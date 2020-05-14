Public Class EDC856O1

    Private tblEDT856O1 As DataTable = Nothing
    Private tblEDT856O2 As DataTable = Nothing
    Private tblEDT856O3 As DataTable = Nothing
    Private tblEDT856O4 As DataTable = Nothing
    Private tblEDT856O5 As DataTable = Nothing
    Private tblEDT856O6 As DataTable = Nothing
    Private tblEDTSYSIH As DataTable = Nothing
    Private tblEDTXREF3 As DataTable = Nothing

    Private rowARTPARM1 As DataRow = Nothing

    Private COMPANY_CODE As String = String.Empty
    Private Const EDI_PROCESS_IND As String = "1"
    
    Private EDI_OUTBOUND_DOC_NO As String = String.Empty

    Private tblSOTSVIA1 As DataTable = Nothing
    Private tblTATTERM1 As DataTable = Nothing
    Private tblEDTTRPM1 As DataTable = Nothing
    Private tblWHTPKGM1 As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing
    Private rowEDT850T1 As DataRow = Nothing

    Private ediAsnPerPoExt As String = String.Empty

    ' Customers having 850 cartons explode to Eaches
    Private lstCartonPackRollUpCustomer As New List(Of String)(New String() {"LOBLAW"})

    Public Sub New(ByRef datasetIn As DataSet)

        InitializeData(datasetIn.Tables("EDTSYSIH"), _
                       datasetIn.Tables("EDT856O1"), _
                       datasetIn.Tables("EDT856O2"), _
                       datasetIn.Tables("EDT856O3"), _
                       datasetIn.Tables("EDT856O4"), _
                       datasetIn.Tables("EDT856O5"), _
                       datasetIn.Tables("EDT856O6"))

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tblEDTSYSIHin">Reference to table EDTSYSIH</param>
    ''' <param name="tblEDT856O1in">Reference to table EDT856O1</param>
    ''' <param name="tblEDT856O2in">Reference to table EDT856O2</param>
    ''' <param name="tblEDT856O3in">Reference to table EDT856O3</param>
    ''' <param name="tblEDT856O4in">Reference to table EDT856O4</param>
    ''' <param name="tblEDT856O5in">Reference to table EDT856O5</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tblEDTSYSIHin As DataTable, _
                   ByRef tblEDT856O1in As DataTable, _
                   ByRef tblEDT856O2in As DataTable, _
                   ByRef tblEDT856O3in As DataTable, _
                   ByRef tblEDT856O4in As DataTable, _
                   ByRef tblEDT856O5in As DataTable, _
                   ByRef tblEDT856O6in As DataTable)

        InitializeData(tblEDTSYSIHin, _
                       tblEDT856O1in, _
                       tblEDT856O2in, _
                       tblEDT856O3in, _
                       tblEDT856O4in, _
                       tblEDT856O5in, _
                       tblEDT856O6in)

    End Sub

    Private Sub InitializeData(ByRef tblEDTSYSIHin As DataTable, _
                   ByRef tblEDT856O1in As DataTable, _
                   ByRef tblEDT856O2in As DataTable, _
                   ByRef tblEDT856O3in As DataTable, _
                   ByRef tblEDT856O4in As DataTable, _
                   ByRef tblEDT856O5in As DataTable, _
                   ByRef tblEDT856O6in As DataTable)

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
        tblEDT856O1 = tblEDT856O1in
        tblEDT856O2 = tblEDT856O2in
        tblEDT856O3 = tblEDT856O3in
        tblEDT856O4 = tblEDT856O4in
        tblEDT856O5 = tblEDT856O5in
        tblEDT856O6 = tblEDT856O6in

        EDI_OUTBOUND_DOC_NO = String.Empty
        tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT SOTSVIA1.*, SOTCARR1.CARRIER_TYPE FROM SOTSVIA1, SOTCARR1 WHERE SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE", "SOTSVIA1", String.Empty, Nothing)
        tblSOTSVIA1.PrimaryKey = New DataColumn() {tblSOTSVIA1.Columns("SHIP_VIA_CODE")}

        tblTATTERM1 = ASCDATA1.GetDataTable("SELECT * FROM TATTERM1", "TATTERM1", String.Empty, Nothing)
        tblEDTTRPM1 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM1 where EDI_DOC_NO = '856'", "EDTTRPM1", String.Empty, Nothing)
        tblWHTPKGM1 = ASCDATA1.GetDataTable("SELECT * FROM WHTPKGM1", "WHTPKGM1")
        rowARTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM ARTPARM1 WHERE AR_PARM_KEY = 'Z'")
        tblEDTSLSP1 = ASCDATA1.GetDataTable("SELECT * FROM EDTSLSP1")
        tblEDTSLSP1.PrimaryKey = New DataColumn() {tblEDTSLSP1.Columns("CUST_CODE")}

        tblEDTXREF3 = ASCDATA1.GetDataTable("SELECT * FROM EDTXREF3")

        'added 6/12/2018
        If Not tblEDT856O4.Columns.Contains("EDI_DTL_SEQ_850") Then
            tblEDT856O4.Columns.Add("EDI_DTL_SEQ_850", GetType(System.Int32))
        End If

        ediAsnPerPoExt = "0"
    End Sub

    ''' <summary>
    ''' Creates the EDT865 table entries
    ''' </summary>
    ''' <param name="SHIP_BOL_NO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateEDI856(ByVal SHIP_BOL_NO As String, ByRef ErrorMessage As String) As String

        ' What are differences between BOL and Master BOL??

        Dim rowARTCUST1 As DataRow = Nothing
        Dim rowARTCUST2 As DataRow = Nothing
        Dim rowEDT856O1 As DataRow = Nothing
        Dim rowEDT856O2 As DataRow = Nothing
        Dim rowEDT856O3 As DataRow = Nothing
        Dim rowEDT856O4 As DataRow = Nothing
        Dim rowEDT856O5 As DataRow = Nothing

        Dim rowEDT850T1 As DataRow = Nothing
        Dim rowEDT850T2 As DataRow = Nothing
        Dim rowEDT850T6 As DataRow = Nothing
        Dim rowEDTTRPM1 As DataRow = Nothing
        Dim rowICTWHSE1 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing
        Dim rowSOTPICK1 As DataRow = Nothing
        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim rowSOTSHIPB As DataRow = Nothing
        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowWHTPKGM1 As DataRow = Nothing
        Dim rowEDTSLSP1 As DataRow = Nothing

        Dim tblSOTCART1 As DataTable = Nothing
        Dim tblSOTCART2 As DataTable = Nothing
        Dim tblSOTORDR1 As DataTable = Nothing
        Dim tblSOTORDR5 As DataTable = Nothing
        Dim tblSOTPICK1 As DataTable = Nothing
        Dim tblSOTSHIP1 As DataTable = Nothing
        Dim tblARTCUST2 As DataTable = Nothing
        Dim tblEDT945T2 As DataTable = Nothing
        Dim tblEDTPPKS1 As DataTable = Nothing

        Dim BILL_OF_LADING_NO As String = String.Empty
        Dim CART_NO As String = String.Empty
        Dim CUST_CODE As String = String.Empty
        Dim EDI_CUSTOMER As String = String.Empty
        Dim EDI_DTL_SEQ As Int32 = 0
        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim PICK_NO As String = String.Empty
        Dim SHIP_856_BATCH_NO As String = String.Empty
        Dim rowSOTSHIP1X As DataRow = Nothing
        Dim EDI_PROMOTION As Boolean = False
        Dim EDI_DOC_SEQ_NO_PROMO As String = String.Empty
        Dim caseFactor As Int16 = 1
        Dim useEdiAsnPerPO As Boolean = False
        Dim SERVICE_LEVEL_3PL As String = String.Empty

        Dim sql As String = String.Empty
        Dim methodType As String = String.Empty

        rowSOTSHIP1 = ASCDATA1.GetDataRow("Select * from SOTSHIP1 where SHIP_BOL_NO = :PARM1 and NVL(SHIP_856_IND, '0') = '1' and (SHIP_856_BATCH_NO IS NULL OR SHIP_856_BATCH_NO = '')", "V", New Object() {SHIP_BOL_NO})
        If rowSOTSHIP1 Is Nothing Then
            ErrorMessage = "Cannot locate shipment " & SHIP_BOL_NO
            Return String.Empty
        End If

        rowICTWHSE1 = ASCDATA1.GetDataRow("SELECT * FROM ICTWHSE1 WHERE WHSE_CODE = :PARM1", "V", New Object() {rowSOTSHIP1.Item("WHSE_CODE") & String.Empty})
        If rowICTWHSE1 Is Nothing Then
            ErrorMessage = "Cannot locate Warehouse Master for shipment " & SHIP_BOL_NO
            Return String.Empty
        End If

        BILL_OF_LADING_NO = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty

        ' 945's have Bill of Ladings, but no SOTSHIPB record.
        Dim wktable As String = String.Empty
        Dim dataFound As Boolean = False
        Dim useBOL As Boolean = False

        ' Collection of Pick_Nos and Order Nos for this shipment
        If BILL_OF_LADING_NO.Length > 0 Then
            sql = " SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO "
            sql &= " FROM SOTSHIPB, SOTSHIP1, SOTPICK1"
            sql &= " WHERE SOTSHIPB.BOL_NO = SOTSHIP1.BILL_OF_LADING_NO"
            sql &= " AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
            sql &= " AND SOTSHIPB.BOL_NO = '" & BILL_OF_LADING_NO & "'"
            sql &= " AND SOTSHIP1.SHIP_856_IND = '1'"
            sql &= " AND SOTSHIP1.EDI_856_CREATED IS NULL"
            wktable = ASCMAIN1.Temp_Table(sql)

            If Val(ASCDATA1.GetDataValue("SELECT COUNT(*) FROM " & wktable) & String.Empty) > 0 Then
                dataFound = True
                useBOL = True
            End If
        End If

        ' Bill of lading number without a SOTORDRB - do a good match down to customer
        If Not dataFound AndAlso BILL_OF_LADING_NO.Length > 0 Then
            sql = " SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO "
            sql &= " FROM SOTSHIP1, SOTPICK1, SOTORDR1, SOTORDR0"
            sql &= " WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
            sql &= " AND SOTSHIP1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'"
            sql &= " AND SOTSHIP1.SHIP_856_IND = '1'"
            sql &= " AND SOTSHIP1.EDI_856_CREATED IS NULL"
            sql &= " AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"
            sql &= " AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO"
            sql &= " AND SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO"
            sql &= " AND SOTORDR0.CUST_CODE = SOTORDR1.CUST_CODE"
            wktable = ASCMAIN1.Temp_Table(sql)

            If Val(ASCDATA1.GetDataValue("SELECT COUNT(*) FROM " & wktable) & String.Empty) > 0 Then
                dataFound = True
                useBOL = True
            End If
        End If

        If Not dataFound Then
            sql = " SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO "
            sql &= " FROM SOTSHIP1, SOTPICK1"
            sql &= " WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
            sql &= " AND SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            sql &= " AND SOTSHIP1.SHIP_856_IND = '1'"
            sql &= " AND SOTSHIP1.EDI_856_CREATED IS NULL"
            wktable = ASCMAIN1.Temp_Table(sql)
        End If

        tblSOTORDR1 = ASCDATA1.GetDataTable("SELECT DISTINCT(CUST_CODE) FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & wktable & ")")
        If tblSOTORDR1.Rows.Count = 0 Then
            ' Nothing Found
            ErrorMessage = "Cannot locate Orders for shipment " & SHIP_BOL_NO
            Return String.Empty
        ElseIf tblSOTORDR1.Rows.Count > 1 Then
            ' multilpe customers 
            ErrorMessage = "There are multiple customers associated with shipment " & SHIP_BOL_NO
            Return String.Empty
        Else
            CUST_CODE = tblSOTORDR1.Rows(0).Item("CUST_CODE") & String.Empty
        End If

        sql = " SELECT SOTORDR2.*, EDT850T2.*"
        sql &= " FROM SOTORDR2, EDT850T2 "
        sql &= " WHERE SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO (+)"
        sql &= " AND SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ (+)"
        sql &= " AND ORDR_NO IN (SELECT ORDR_NO FROM " & wktable & ")"
        Dim tblSOTORDR2 As DataTable = ASCDATA1.GetDataTable(sql, "SOTORDR2")

        rowEDTSLSP1 = tblEDTSLSP1.Rows.Find(CUST_CODE)
        Dim CartonPackRollUpCustomer As Boolean = lstCartonPackRollUpCustomer.Contains(CUST_CODE)


        ' See if the customer wants Separate ASNs for consolidated shipments
        ' Multiple entries in SOTSHIP1 will have the same value in BILL_OF_LADING_NO; however, the customer wants each
        ' SOTSHIP1 record as its own ASN.
        If BILL_OF_LADING_NO.Length > 0 AndAlso tblEDTSLSP1.Columns.Contains("EDI_ASN_PER_PO") AndAlso rowEDTSLSP1.Item("EDI_ASN_PER_PO") & String.Empty = "1" Then
            If ASCDATA1.GetDataTable("SELECT DISTINCT(SHIP_BOL_NO) FROM SOTPICK1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")").Rows.Count > 1 Then
                sql = " SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO "
                sql &= " FROM SOTSHIP1, SOTPICK1"
                sql &= " WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
                sql &= " AND SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                sql &= " AND SOTSHIP1.SHIP_856_IND = '1'"
                sql &= " AND SOTSHIP1.EDI_856_CREATED IS NULL"
                wktable = ASCMAIN1.Temp_Table(sql)
                useBOL = False
                useEdiAsnPerPO = True
                ediAsnPerPoExt = (Val(ediAsnPerPoExt) + 1).ToString.Trim
            End If
        End If

        tblSOTORDR1 = ASCDATA1.GetDataTable("SELECT DISTINCT(CUST_CODE) FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & wktable & ")")
        If tblSOTORDR1.Rows.Count = 0 Then
            ' Nothing Found
            ErrorMessage = "Cannot locate Orders for shipment " & SHIP_BOL_NO
            Return String.Empty
        ElseIf tblSOTORDR1.Rows.Count > 1 Then
            ' multilpe customers 
            ErrorMessage = "There are multiple customers associated with shipment " & SHIP_BOL_NO
            Return String.Empty
        Else
            CUST_CODE = tblSOTORDR1.Rows(0).Item("CUST_CODE") & String.Empty
        End If

        If BILL_OF_LADING_NO.Length > 0 AndAlso useBOL Then
            rowSOTSHIPB = ASCDATA1.GetDataRow("SELECT * FROM SOTSHIPB WHERE BOL_NO = '" & BILL_OF_LADING_NO & "'")
            rowSOTSHIP1X = ASCDATA1.GetDataRow("SELECT * FROM SOTSHIP1 WHERE BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "' AND EDI_856_CREATED IS NULL")
            useBOL = rowSOTSHIPB IsNot Nothing ' Not all records with BILL_OF_LADING_NO are in SOTSHIPB
        Else
            rowSOTSHIP1X = ASCDATA1.GetDataRow("SELECT * FROM SOTSHIP1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "' AND EDI_856_CREATED IS NULL")
        End If

        rowSOTSVIA1 = tblSOTSVIA1.Rows.Find(rowSOTSHIP1X.Item("SHIP_VIA_CODE") & String.Empty)
        If rowSOTSVIA1 Is Nothing Then
            ErrorMessage = "Cannot locate Ship Via or the Ship Vias carrier."
            Return String.Empty
        End If

        Dim EDI_ID_856 As String = rowEDTSLSP1.Item("EDI_ID_856") & String.Empty
        Dim EDI_QUAL_856 As String = rowEDTSLSP1.Item("EDI_QUAL_856") & String.Empty
        Dim EDI_CONSUMER As Boolean = rowEDTSLSP1.Item("EDI_CONSUMER") & String.Empty = "1"
        Dim rowEDTXREF3 As DataRow = Nothing

        sql = "EDI_TP_QUAL = '" & EDI_QUAL_856 & "' and EDI_TP_ID = '" & EDI_ID_856 & "' and EDI_DOC_NO = '856'"

        If tblEDTTRPM1.Select(sql).Length = 0 Then
            ErrorMessage = "Customer (" & CUST_CODE & ") associated with shipment " & SHIP_BOL_NO & " is not setup to receive 856 data."
            Return String.Empty
        Else
            rowEDTTRPM1 = tblEDTTRPM1.Select(sql)(0)
        End If

        ' Set from EDTTRPM1
        Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
        Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty

        Dim SHIP_ADDR_TYPE As String = rowSOTSHIP1X.Item("SHIP_ADDR_TYPE") & String.Empty
        Dim SHIP_ADDR_CODE As String = rowSOTSHIP1X.Item("SHIP_ADDR_CODE") & String.Empty
        Dim numChars As Int16 = 0

        If rowEDTSLSP1 IsNot Nothing Then
            Select Case SHIP_ADDR_TYPE
                Case "MK"
                    numChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
                Case "DC"
                    numChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & String.Empty)
            End Select
        End If

        If numChars > 0 And IsNumeric(SHIP_ADDR_CODE) Then
            SHIP_ADDR_CODE = SHIP_ADDR_CODE.PadLeft(numChars, "0")
            SHIP_ADDR_CODE = StrReverse(StrReverse(SHIP_ADDR_CODE).Substring(0, numChars))
        End If

        ' Reset this value for the MK in EDT865O5
        numChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)

        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
        tblARTCUST2 = ASCDATA1.GetDataTable("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1", "ARTCUST2", "V", New Object() {CUST_CODE})
        tblARTCUST2.PrimaryKey = New DataColumn() {tblARTCUST2.Columns("CUST_CODE"), tblARTCUST2.Columns("CUST_ADDR_TYPE"), tblARTCUST2.Columns("CUST_ADDR_CODE")}

        ' Load Cartons.
        sql = " SELECT * FROM SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTCART1 = ASCDATA1.GetDataTable(sql, "SOTCART1")
        If tblSOTCART1.Rows.Count = 0 Then
            ErrorMessage = "Cannot find cartons associated with shipment " & SHIP_BOL_NO
            Return String.Empty
        End If

        sql = "SELECT * FROM EDT945T2 WHERE EDI_CART_NO IN (SELECT CART_NO FROM SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & wktable & "))"
        tblEDT945T2 = ASCDATA1.GetDataTable(sql, "EDT945T2")

        ' Get data where Cases are exploded to Eaches. Need to put back as cases in 856
        sql = " Select EDTPPKS1.*, ICTSTYL1.CARTON_PACK_QTY CARTON_PACK_QTY_STYLE"
        sql &= "  from EDTPPKS1,ICTSTYL1"
        sql &= "  where ICTSTYL1.STYLE_CODE (+) = EDTPPKS1.STYLE_CODE"
        sql &= "  and EDTPPKS1.CUST_CODE = :PARM1"
        tblEDTPPKS1 = ASCDATA1.GetDataTable(sql, "EDTPPKS1", "V", New Object() {CUST_CODE})

        sql = "SELECT SOTCART2.*,  ICTSTYL1.STYLE_STATUS, ICTSTYL1.STYLE_DESC, ICTSTYL1.REQUIRES_EXP_DATE, SOTORDR2.STYLE_RETAIL"
        sql &= ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.EDI_DOC_SEQ_NO, SOTORDR2.EDI_DTL_SEQ, SOTORDR2.STYLE_CODE_SUB"
        ' Regency Wayfair Set Quantity
        If ASCMAIN1.CLIENT = "RGI" Then
            sql &= " ,SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_SKU, ICTSTYC1.UPC_CODE, NVL(SOTORDR2.SET_QTY, 1) SET_QTY"
        Else
            sql &= " ,SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_SKU, ICTSTYC1.UPC_CODE, 1 SET_QTY"
        End If

        sql &= " FROM SOTCART1, SOTCART2, ICTSTYL1, ICTSTYC1, SOTORDR2"
        sql &= " WHERE SOTCART1.CART_NO = SOTCART2.CART_NO"
        sql &= " AND NVL(SOTORDR2.STYLE_CODE_SUB, SOTCART2.STYLE_CODE) = ICTSTYL1.STYLE_CODE"
        sql &= " AND ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE"
        sql &= " AND SOTORDR2.COLOR_CODE = ICTSTYC1.COLOR_CODE"
        sql &= " AND SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO AND SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO"
        sql &= " AND SOTCART1.PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTCART2 = ASCDATA1.GetDataTable(sql, "SOTCART2")

        If rowARTCUST1.Item("EDI_REQUIRES_EXP_DATE") & String.Empty = "1" Then
            For Each row As DataRow In tblSOTCART2.Select("REQUIRES_EXP_DATE = '1'")
                If row.Item("ITEM_EXP_DATE") & String.Empty = String.Empty Then
                    ErrorMessage = "Shipment " & SHIP_BOL_NO & " contains items that require an Expiration Date and the date is missing."
                    Return String.Empty
                End If
            Next
        End If

        ' Load Pick Tickets
        sql = "SELECT SOTPICK1.*, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DEPT, SOTORDR1.ORDR_DATE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.EDI_DOC_SEQ_NO"
        sql &= " FROM SOTORDR1, SOTPICK1"
        sql &= " WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
        sql &= " AND SOTPICK1.PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTPICK1 = ASCDATA1.GetDataTable(sql, "SOTPICK1")
        tblSOTPICK1.PrimaryKey = New DataColumn() {tblSOTPICK1.Columns("PICK_NO")}

        EDI_DOC_SEQ_NO = tblSOTPICK1.Compute("MAX(EDI_DOC_SEQ_NO)", "") & String.Empty

        ' Load Shipping Addresses
        tblSOTORDR5 = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR5 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & wktable & ")", "SOTORDR5")

        ' Load Shipment header
        sql = "SELECT DISTINCT SOTSHIP1.* FROM SOTSHIP1, SOTPICK1 WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO AND SOTPICK1.PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTSHIP1 = ASCDATA1.GetDataTable(sql, "SOTSHIP1")
        tblSOTSHIP1.PrimaryKey = New DataColumn() {tblSOTSHIP1.Columns("SHIP_BOL_NO")}

        ' UPS requirement – I think we agreed that we would hard code it for RGI specific customer
        ' Ship Confirmation/ ASN Create for Nordstrom (RGI/customer 302941, SOTSVIA1.CARRIER_MODE = 'U')
        ' UPS Separate ASNs for each carton - Bill of Lading = UPS tracking number
        ' Requirement: 1 ASN per carton - must conform to routing instructions weights

        Dim breakoutByCarton As Boolean = False
        If ASCMAIN1.CLIENT = "RGI" Then
            If CUST_CODE = "302941" Then
                If rowSOTSVIA1.Item("CARRIER_MODE") & String.Empty = "U" Then
                    breakoutByCarton = True
                End If
            End If
        End If

        Dim tblCartonsLoop As DataTable = Nothing

        If Not breakoutByCarton Then
            ' forces the loop to execute once
            tblCartonsLoop = New DataTable
            tblCartonsLoop.Columns.Add("CART_NO", GetType(System.String))
            tblCartonsLoop.Rows.Add(New Object() {"."})
        Else
            tblCartonsLoop = tblSOTCART1.Copy
        End If

        For Each rowLoop As DataRow In tblCartonsLoop.Select("", "CART_NO")
            Dim cartonNoLoop As String = rowLoop.Item("CART_NO")
            Dim sqlCartonNoLoop As String = String.Empty

            EDI_OUTBOUND_DOC_NO = Me.CreateEDTSYSIH(EDI_OUR_ID, EDI_TP_ID, "SH", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)
            SHIP_856_BATCH_NO = EDI_OUTBOUND_DOC_NO
            Dim EDI_ADR_SEQ As Int16 = 0

            ' Header data From SOTSHIPB, ICTWHSE1
            rowEDT856O1 = tblEDT856O1.NewRow
            rowEDT856O1.Item("COMPANY_CODE") = COMPANY_CODE
            rowEDT856O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            rowEDT856O1.Item("WHSE_CITY") = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
            rowEDT856O1.Item("WHSE_STATE") = rowICTWHSE1.Item("WHSE_STATE") & String.Empty

            If Not breakoutByCarton Then
                rowEDT856O1.Item("EDI_SHIP_CNT_CARTONS") = tblSOTCART1.Rows.Count
                ' Debbie does not want to enter the weights on the cartons so use weight from pick ticket - 1/30/2013
                rowEDT856O1.Item("EDI_SHIP_TOTAL_WGT") = Val(tblSOTPICK1.Compute("SUM(PICK_TOTAL_WGT)", "") & String.Empty) 'Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "") & String.Empty)
            Else
                rowEDT856O1.Item("EDI_SHIP_CNT_CARTONS") = 1
                rowEDT856O1.Item("EDI_SHIP_TOTAL_WGT") = Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "' AND CART_NO = '" & cartonNoLoop & "'") & String.Empty)
            End If

            rowEDT856O1.Item("EDI_REMIT_NAME") = rowARTPARM1.Item("AR_PARM_REMIT_NAME") & String.Empty
            rowEDT856O1.Item("EDI_TP_ID") = EDI_TP_ID
            rowEDT856O1.Item("EDI_SUPPLIER_NO") = rowEDTTRPM1.Item("EDI_ACCT_REF_NO") & String.Empty
            rowEDT856O1.Item("SHIP_856_BATCH_NO") = SHIP_856_BATCH_NO
            rowEDT856O1.Item("WHSE_ZIP_CODE") = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
            rowEDT856O1.Item("EDI_CUSTOMER") = EDI_CUSTOMER
            rowEDT856O1.Item("CARRIER_MODE") = rowSOTSVIA1.Item("CARRIER_TYPE") & String.Empty
            rowEDT856O1.Item("SHIP_ADDR_CODE") = SHIP_ADDR_CODE
            'rowEDT856O1.Item("EDI_TRACKING_NUMBER") = String.Empty

            rowEDT850T1 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
            If rowEDT850T1 IsNot Nothing Then
                rowEDT856O1.Item("EDI_SUPPLIER_NO") = rowEDT850T1.Item("EDI_SUPPLIER_NO")
            End If

            SERVICE_LEVEL_3PL = String.Empty
            If EDI_CONSUMER Then
                Dim SHIP_VIA_CODE As String = rowSOTSVIA1.Item("SHIP_VIA_CODE") & String.Empty
                Dim sqlXref As String = "SENDER_ID_QUAL = '" & EDI_QUAL_856 & "' AND SENDER_ID = '" & EDI_ID_856 & "' AND SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'"
                If tblEDTXREF3.Select(sqlXref).Length = 1 Then
                    SERVICE_LEVEL_3PL = tblEDTXREF3.Select(sqlXref)(0).Item("SERVICE_LEVEL_3PL") & String.Empty
                End If

                If SERVICE_LEVEL_3PL.Length = 0 Then
                    If rowEDT850T1 IsNot Nothing Then
                        SERVICE_LEVEL_3PL = rowEDT850T1.Item("EDI_SHIPPER") & String.Empty
                    End If
                End If

                rowEDT856O1.Item("SHIP_VIA_DESC") = SERVICE_LEVEL_3PL
            End If

            If BILL_OF_LADING_NO.Length > 0 AndAlso useBOL Then
                rowEDT856O1.Item("BILL_OF_LADING_NO") = rowSOTSHIPB.Item("BOL_NO") & String.Empty ' IIf(rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty <> String.Empty, rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty, rowSOTSHIPB.Item("BOL_NO") & String.Empty)
                rowEDT856O1.Item("SHIP_BOL_NO") = rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty
                rowEDT856O1.Item("EDI_DATE_SHIPPED") = rowSOTSHIPB.Item("SHIPPED_ACTUAL")
                rowEDT856O1.Item("EDI_SCHED_DELIV_DATE") = rowSOTSHIPB.Item("SCHED_DELIV_DATE")
                rowEDT856O1.Item("FRT_TERMS") = IIf(rowSOTSHIPB.Item("FRT_TERMS") & String.Empty = "COL", "CC", "PP")
                rowEDT856O1.Item("EDI_SCAC_CODE") = rowSOTSHIPB.Item("SHIP_VIA_SCAC") & String.Empty
                rowEDT856O1.Item("EDI_PRO_NO") = rowSOTSHIPB.Item("SHIP_REF") & String.Empty
                If Trim(SERVICE_LEVEL_3PL).Length = 0 Then
                    rowEDT856O1.Item("SHIP_VIA_DESC") = rowSOTSHIPB.Item("SHIP_VIA_DESC") & String.Empty
                End If
                rowEDT856O1.Item("MASTER_BILL_OF_LADING_NO") = IIf(rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty <> String.Empty, rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty, rowSOTSHIPB.Item("BOL_NO") & String.Empty)
                'rowEDT856O1.Item("SHIP_ADDR_CODE") = tblSOTPICK1.Rows(0).Item("ORDR_ADDR_TYPE_ST")
                'rowEDT856O1.Item("SHIP_MANIFEST_NO") = rowSOTSHIPB.Item("SHIP_MANIFEST_NO") & String.Empty
                rowEDT856O1.Item("EDI_LOAD_ID") = rowSOTSHIPB.Item("EDI_LOAD_ID") & String.Empty
            Else
                If BILL_OF_LADING_NO.Length > 0 Then
                    rowEDT856O1.Item("BILL_OF_LADING_NO") = BILL_OF_LADING_NO & IIf(useEdiAsnPerPO, "-" & ediAsnPerPoExt, "")
                    rowEDT856O1.Item("MASTER_BILL_OF_LADING_NO") = BILL_OF_LADING_NO
                Else
                    rowEDT856O1.Item("BILL_OF_LADING_NO") = "9" & (rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty).ToString.Substring(1)
                    rowEDT856O1.Item("MASTER_BILL_OF_LADING_NO") = rowEDT856O1.Item("BILL_OF_LADING_NO") 'rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty
                End If
                rowEDT856O1.Item("SHIP_BOL_NO") = rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty
                rowEDT856O1.Item("EDI_DATE_SHIPPED") = rowSOTSHIP1X.Item("SHIP_DATE_SHIPPED")
                ' For now just add 3 days when no Bill of Lading
                rowEDT856O1.Item("EDI_SCHED_DELIV_DATE") = DateAdd(DateInterval.Day, 3, rowSOTSHIP1X.Item("SHIP_DATE_SHIPPED"))
                rowEDT856O1.Item("FRT_TERMS") = IIf(rowSOTSHIP1X.Item("FRT_TERMS") & String.Empty = "COL", "CC", "PP")
                If rowSOTSVIA1 IsNot Nothing Then
                    rowEDT856O1.Item("EDI_SCAC_CODE") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
                    If SERVICE_LEVEL_3PL.Length = 0 Then
                        rowEDT856O1.Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    End If
                End If
                rowEDT856O1.Item("EDI_PRO_NO") = rowSOTSHIP1X.Item("SHIP_REF") & String.Empty
                rowEDT856O1.Item("EDI_LOAD_ID") = rowSOTSHIP1X.Item("EDI_LOAD_ID") & String.Empty
                'rowEDT856O1.Item("SHIP_MANIFEST_NO") = rowSOTSHIPB.Item("SHIP_MANIFEST_NO") & String.Empty
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                If rowEDT856O1.Item("CARRIER_MODE") & String.Empty = "U" Then
                    rowEDT856O1.Item("EDI_TRACKING_NUMBER") = rowSOTSHIP1.Item("SHIP_REF") & String.Empty
                End If
            End If

            tblEDT856O1.Rows.Add(rowEDT856O1)

            Dim EDI_HL2_SEQ As Int16 = 0
            Dim EDI_HL3_SEQ As Int16 = 0
            Dim EDI_HL4_SEQ As Int16 = 0

            ' Pick Tickets
            sqlCartonNoLoop = "SELECT DISTINCT PICK_NO FROM " & wktable
            If breakoutByCarton Then
                sqlCartonNoLoop = "SELECT XX.PICK_NO FROM " & wktable & " XX, SOTCART1 WHERE XX.PICK_NO = SOTCART1.PICK_NO AND SOTCART1.CART_NO = '" & cartonNoLoop & "'"
            End If

            For Each rowPICKNO As DataRow In ASCDATA1.GetDataTable(sqlCartonNoLoop, wktable).Select("", "PICK_NO")
                PICK_NO = rowPICKNO.Item("PICK_NO")
                rowSOTPICK1 = tblSOTPICK1.Rows.Find(PICK_NO)
                SHIP_BOL_NO = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                rowSOTSHIP1 = tblSOTSHIP1.Rows.Find(SHIP_BOL_NO)

                Dim ORDR_GROUP_NO As String = rowSOTSHIP1.Item("ORDR_GROUP_NO") & String.Empty
                methodType = "A"
                Dim rowSOTORDR0 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO = :PARM1", "V", New Object() {ORDR_GROUP_NO})
                If rowSOTORDR0 IsNot Nothing AndAlso rowSOTORDR0.Item("EDI_DOC_SEQ_NO") & String.Empty <> String.Empty Then
                    EDI_DOC_SEQ_NO = rowSOTORDR0.Item("EDI_DOC_SEQ_NO") & String.Empty
                    If EDI_DOC_SEQ_NO.Length > 0 Then
                        If ASCDATA1.GetDataTable("SELECT * FROM EDT850T6 WHERE EDI_DOC_SEQ_NO = :PARM1", "", "V", New Object() {EDI_DOC_SEQ_NO}).Rows.Count > 0 Then
                            methodType = "B"
                        End If
                    End If
                End If

                'New rule on 10/27/2014 - If EDTSLSP1.EDI_IGNORE SLN = 1 and cust_code <> WALMART
                If rowEDTSLSP1 IsNot Nothing AndAlso rowEDTSLSP1.Item("EDI_IGNORE_SLN") & String.Empty = "1" Then
                    If COMPANY_CODE = "NYA" Then
                        If CUST_CODE <> "WALMART" Then
                            methodType = "A"
                        End If
                    Else
                        methodType = "A"
                    End If
                End If

                ' Pick Ticket / Invoice header data
                rowEDT856O2 = tblEDT856O2.NewRow
                rowEDT856O2.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT856O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                EDI_HL2_SEQ += 1
                EDI_ADR_SEQ = 0
                rowEDT856O2.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                rowEDT856O2.Item("ORDR_CUST_PO") = rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                rowEDT856O2.Item("ORDR_DEPT") = rowSOTPICK1.Item("ORDR_DEPT") & String.Empty

                If Not breakoutByCarton Then
                    rowEDT856O2.Item("EDI_ORD_CNT_CARTONS") = tblSOTCART1.Select("PICK_NO = '" & PICK_NO & "'").Length
                    ' Debbie does not want to enter the weights on the cartons so use weight from pick ticket - 1/30/2013
                    rowEDT856O2.Item("EDI_ORD_TOTAL_WGT") = Val(rowSOTPICK1.Item("PICK_TOTAL_WGT") & String.Empty)
                Else
                    rowEDT856O2.Item("EDI_ORD_CNT_CARTONS") = 1
                    rowEDT856O2.Item("EDI_ORD_TOTAL_WGT") = Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "' AND CART_NO = '" & cartonNoLoop & "'") & String.Empty)
                End If
                rowEDT856O2.Item("ORDR_NO") = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                rowEDT856O2.Item("PICK_NO") = rowSOTPICK1.Item("PICK_NO") & String.Empty

                Dim SHIP_REF As String = rowSOTSHIP1.Item("SHIP_REF") & String.Empty
                If SHIP_REF.Length > rowEDT856O2.Table.Columns("PRO_NO").MaxLength Then
                    SHIP_REF = SHIP_REF.Substring(0, rowEDT856O2.Table.Columns("PRO_NO").MaxLength)
                End If
                rowEDT856O2.Item("PRO_NO") = SHIP_REF
                rowEDT856O2.Item("INV_NO") = rowSOTPICK1.Item("INV_NO") & String.Empty
                rowEDT856O2.Item("ORDR_DATE") = rowSOTPICK1.Item("ORDR_DATE") & String.Empty
                rowEDT856O2.Item("EDI_ORDER_STATUS") = rowSOTPICK1.Item("ORDR_STATUS") & String.Empty
                rowEDT856O2.Item("CUST_STORE_NO") = rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                rowEDT856O2.Item("EDI_CUSTOMER") = EDI_CUSTOMER

                ' EDT856O2.EDI_PO_RELEASE_NO = EDT850T1.EDI_PO_RELEASE_NO
                ' Ricks request 06/12/2018
                If rowEDT850T1 IsNot Nothing Then
                    If rowEDT856O2.Table.Columns.Contains("EDI_PO_RELEASE_NO") AndAlso rowEDT850T1.Table.Columns.Contains("EDI_PO_RELEASE_NO") Then
                        rowEDT856O2.Item("EDI_PO_RELEASE_NO") = rowEDT850T1.Item("EDI_PO_RELEASE_NO")
                    End If
                End If

                tblEDT856O2.Rows.Add(rowEDT856O2)

                ' Maria wants a MK for each rowEDT856O2 record
                rowEDT856O5 = tblEDT856O5.NewRow
                rowEDT856O5.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT856O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT856O5.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                EDI_ADR_SEQ += 1

                Dim CUST_ADDR_CODE As String = rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                rowARTCUST2 = tblARTCUST2.Rows.Find(New Object() {CUST_CODE, "MK", CUST_ADDR_CODE})

                rowEDT856O5.Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                rowEDT856O5.Item("EDI_ADDR_TYPE") = "MK"
                rowEDT856O5.Item("EDI_CUST_NAME_ADR") = rowARTCUST2.Item("CUST_NAME") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS1") = rowARTCUST2.Item("CUST_ADDR1") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS2") = rowARTCUST2.Item("CUST_ADDR2") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS3") = rowARTCUST2.Item("CUST_ADDR3") & String.Empty
                rowEDT856O5.Item("EDI_CITY") = rowARTCUST2.Item("CUST_CITY") & String.Empty
                rowEDT856O5.Item("EDI_STATE") = rowARTCUST2.Item("CUST_STATE") & String.Empty

                rowEDT856O5.Item("EDI_ZIPCODE") = (rowARTCUST2.Item("CUST_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                If (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                    rowEDT856O5.Item("EDI_ZIPCODE") = (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                End If

                rowEDT856O5.Item("EDI_COUNTRY") = rowARTCUST2.Item("CUST_COUNTRY") & String.Empty

                'In NYA ASN , when you are building EDT856O5 for MK.
                'Please set EDI_ADDR_CODE  from ARTCUST2.GLOBAL_LOCATION_NUMBER  if GLN is not null
                'This is required for AAFES two recent shipments for location 000339
                If rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty <> String.Empty Then
                    CUST_ADDR_CODE = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty
                ElseIf numChars > 0 AndAlso IsNumeric(CUST_ADDR_CODE) Then
                    CUST_ADDR_CODE = CUST_ADDR_CODE.PadLeft(numChars, "0")
                    CUST_ADDR_CODE = StrReverse(StrReverse(CUST_ADDR_CODE).Substring(0, numChars))
                End If

                rowEDT856O5.Item("EDI_ADDR_CODE") = CUST_ADDR_CODE
                'rowEDT856O5.Item("EDI_ADDR_CODE_QUAL") = rowICTWHSE1.Item("WHSE_EDI_QUAL") & String.Empty
                tblEDT856O5.Rows.Add(rowEDT856O5)
                EDI_PROMOTION = False

                ' Cartons for the above Pick Ticket
                EDI_HL3_SEQ = 0

                ' Need to get to EDT856O3
                sqlCartonNoLoop = "PICK_NO = '" & PICK_NO & "'"
                If breakoutByCarton Then
                    sqlCartonNoLoop = "CART_NO = '" & cartonNoLoop & "'"
                End If

                For Each rowSOTCART1 As DataRow In tblSOTCART1.Select(sqlCartonNoLoop, "CART_NO")
                    CART_NO = rowSOTCART1.Item("CART_NO")

                    rowEDT856O3 = tblEDT856O3.NewRow
                    rowEDT856O3.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT856O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT856O3.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                    EDI_HL3_SEQ += 1
                    rowEDT856O3.Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                    rowEDT856O3.Item("CART_NO") = CART_NO
                    rowEDT856O3.Item("CART_TOTAL_WGT_ACTUAL") = Val(rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty)

                    If Not breakoutByCarton Then
                        rowEDT856O3.Item("CART_SEQ") = Val(rowSOTCART1.Item("CART_SEQ") & String.Empty)
                    Else
                        rowEDT856O3.Item("CART_SEQ") = 1
                    End If

                    ' You will need to change the ASN process to populate in both ABSolution and the service
                    ' SOTCART1.CART_TRACKING_NO -> EDT856O3.EDI_CTN_TRACKING_NUMBER
                    ' 04/18/2016
                    rowEDT856O3.Item("EDI_CTN_TRACKING_NUMBER") = rowSOTCART1.Item("CART_TRACKING_NO") & String.Empty
                    If breakoutByCarton Then
                        rowEDT856O1.Item("BILL_OF_LADING_NO") = rowSOTCART1.Item("CART_TRACKING_NO") & String.Empty
                    End If

                    If tblEDT945T2.Select("EDI_CART_NO = '" & CART_NO & "'").Length > 0 Then
                        Dim rowEDT945T2 As DataRow = tblEDT945T2.Select("EDI_CART_NO = '" & CART_NO & "'")(0)
                        rowEDT856O3.Item("CARTON_LENGTH") = Val(rowEDT945T2.Item("EDI_CART_LENGTH") & String.Empty)
                        rowEDT856O3.Item("CARTON_WIDTH") = Val(rowEDT945T2.Item("EDI_CART_WIDTH") & String.Empty)
                        rowEDT856O3.Item("CARTON_HEIGHT") = Val(rowEDT945T2.Item("EDI_CART_HEIGHT") & String.Empty)
                    ElseIf rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty = "31" Then
                        'When package is other, we can't get the data from WHTPKGM1
                        If rowSOTCART1.Item("PKG_CODE") & "" = "OTHER" Then
                            rowEDT856O3.Item("CARTON_LENGTH") = rowSOTCART1.Item("PKG_L")
                            rowEDT856O3.Item("CARTON_WIDTH") = rowSOTCART1.Item("PKG_W")
                            rowEDT856O3.Item("CARTON_HEIGHT") = rowSOTCART1.Item("PKG_H")
                        Else
                            rowWHTPKGM1 = tblWHTPKGM1.Rows.Find(rowSOTCART1.Item("PKG_CODE"))
                            If rowWHTPKGM1 IsNot Nothing Then
                                rowEDT856O3.Item("CARTON_LENGTH") = rowWHTPKGM1.Item("PKG_L")
                                rowEDT856O3.Item("CARTON_WIDTH") = rowWHTPKGM1.Item("PKG_W")
                                rowEDT856O3.Item("CARTON_HEIGHT") = rowWHTPKGM1.Item("PKG_H")
                            End If
                        End If
                    Else
                        'rowEDT856O3.Item("CARTON_LENGTH") = String.Empty
                        'rowEDT856O3.Item("CARTON_WIDTH") = String.Empty
                        'rowEDT856O3.Item("CARTON_HEIGHT") = String.Empty
                    End If

                    'rowEDT856O3.Item("CARTON_WGT_PER") = String.Empty
                    tblEDT856O3.Rows.Add(rowEDT856O3)

                    ' Carton Contents
                    EDI_HL4_SEQ = 0

                    ' Verify all EDI data was imported and used
                    Dim tblcartons As DataTable = Nothing

                    If methodType = "B" Then

                        'sql = " SELECT SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ"
                        'sql &= " , SOTCART2.CART_NO, SOTORDR2.STYLE_CODE"
                        'sql &= " , NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM) EDI_PRICE_UOM"
                        'sql &= " , ROUND(SUM(SOTCART2.QTY_PACKED) / SUM(EDT850T6.EDI_SLN_QTY), 2) QTY"
                        'sql &= " , ROUND(SUM(SOTORDR2.ORDR_QTY_ORIG) / SUM(EDT850T6.EDI_SLN_QTY), 2) QTYO"
                        'sql &= " FROM SOTORDR2, EDT850T6, EDT850T2, SOTCART2, SOTORDR1"
                        'sql &= " WHERE SOTCART2.CART_NO = :PARM1"
                        'sql &= " AND EDT850T6.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO AND EDT850T6.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ"
                        'sql &= " AND EDT850T6.EDI_SLN_SEQ  = SOTORDR2.EDI_SLN_SEQ"
                        'sql &= " AND EDT850T2.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO AND EDT850T2.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ"
                        'sql &= " AND SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO AND SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO"
                        'sql &= " AND SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO AND SOTORDR1.ORDR_STATUS <> 'C'"
                        'sql &= " GROUP BY SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ, SOTCART2.CART_NO, SOTORDR2.STYLE_CODE, NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM)"

                        '' Trying this since one SOTORDR2 Record can have Many EDT850T6 lines and need to get the total Pre-Pack EDI_SLN_QTY
                        sql = " SELECT ORDR_NO, EDI_DOC_SEQ_NO, EDI_DTL_SEQ, CART_NO, STYLE_CODE"
                        sql &= " , EDI_PRICE_UOM" ' , QTY_PACKED, ORDR_QTY_ORIG"
                        'sql &= " , ROUND(SUM(QTY_PACKED) / SUM(EDI_SLN_QTY), 2) QTY"
                        sql &= " , ROUND(SUM(QTY_PACKED / EDI_SLN_QTY), 2) QTY"
                        sql &= " , ROUND(SUM(ORDR_QTY_ORIG) / SUM(EDI_SLN_QTY), 2) QTYO "
                        sql &= " FROM"
                        sql &= "  ("
                        sql &= " SELECT SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ , SOTCART2.CART_NO, SOTORDR2.STYLE_CODE "
                        sql &= "  , NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM) EDI_PRICE_UOM , SOTCART2.QTY_PACKED"
                        sql &= "  , SOTORDR2.ORDR_QTY_ORIG , SUM(EDT850T6.EDI_SLN_QTY) EDI_SLN_QTY"
                        sql &= "  FROM SOTORDR2, EDT850T6, EDT850T2, SOTCART2, SOTORDR1 "
                        sql &= "  WHERE SOTCART2.CART_NO = :PARM1 "
                        sql &= "  AND EDT850T6.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO"
                        sql &= "  AND EDT850T6.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ"
                        sql &= "  AND EDT850T2.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO "
                        sql &= "  AND EDT850T2.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ "
                        sql &= "  AND SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO "
                        sql &= "  AND SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO "
                        sql &= "  AND SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO "
                        sql &= "  AND SOTORDR1.ORDR_STATUS <> 'C' "
                        ' Added SOTORDR2.ORDR_LNO, since some orders have the same items more than once.
                        sql &= "  GROUP BY SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ , SOTCART2.CART_NO, SOTORDR2.STYLE_CODE "
                        sql &= "  , NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM) , SOTCART2.QTY_PACKED"
                        sql &= "  , SOTORDR2.ORDR_QTY_ORIG"
                        sql &= "  )"
                        sql &= "  GROUP BY ORDR_NO, EDI_DOC_SEQ_NO, EDI_DTL_SEQ, CART_NO, STYLE_CODE, EDI_PRICE_UOM" ' , QTY_PACKED, ORDR_QTY_ORIG"

                        tblcartons = ASCDATA1.GetDataTable(sql, "", "V", New Object() {CART_NO})

                        If tblcartons.Rows.Count = 0 Then
                            methodType = "A"
                        End If
                    End If

                    If methodType = "B" Then

                        Dim CART_LNO As Int16 = 0

                        For Each rowCarton As DataRow In tblcartons.Select("")
                            EDI_DOC_SEQ_NO = rowCarton.Item("EDI_DOC_SEQ_NO") & String.Empty
                            EDI_DTL_SEQ = Val(rowCarton.Item("EDI_DTL_SEQ") & String.Empty)

                            If EDI_PROMOTION = False Then
                                EDI_PROMOTION = True
                                ' prevents duplicate lookups
                                If EDI_DOC_SEQ_NO_PROMO <> EDI_DOC_SEQ_NO Then
                                    rowEDT850T1 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = :PARM1", "V", New Object() {EDI_DOC_SEQ_NO})
                                    EDI_DOC_SEQ_NO_PROMO = EDI_DOC_SEQ_NO
                                End If
                                If rowEDT850T1 IsNot Nothing Then
                                    rowEDT856O2.Item("EDI_PROMOTION") = rowEDT850T1.Item("EDI_PROMOTION") & String.Empty
                                    rowEDT856O2.Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE") & String.Empty
                                End If
                            End If

                            caseFactor = 1
                            Dim STYLE_CODE As String = rowCarton.Item("STYLE_CODE") & String.Empty
                            Dim rowEDTPPKS1 As DataRow = Nothing
                            If tblEDTPPKS1.Select("CUST_CODE = '" & CUST_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "'").Length > 0 Then
                                rowEDTPPKS1 = tblEDTPPKS1.Select("CUST_CODE = '" & CUST_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "'")(0)
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

                            ' Can we add items to an EDI Order? If so, do we send the new items over
                            rowEDT850T2 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T2 WHERE EDI_DOC_SEQ_NO = :PARM1 AND EDI_DTL_SEQ = :PARM1", "VN", New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ})
                            rowEDT850T6 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T6 WHERE EDI_DOC_SEQ_NO = :PARM1 AND EDI_DTL_SEQ = :PARM1", "VN", New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ})

                            'So if EDI_PO4_UOM is null, then use EDI_PO4_UOM
                            rowEDT856O4 = tblEDT856O4.NewRow
                            rowEDT856O4.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT856O4.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            rowEDT856O4.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                            rowEDT856O4.Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                            EDI_HL4_SEQ += 1
                            rowEDT856O4.Item("EDI_HL4_SEQ") = EDI_HL4_SEQ
                            CART_LNO += 1
                            rowEDT856O4.Item("CART_LNO") = CART_LNO
                            rowEDT856O4.Item("STYLE_UPC_CODE") = (rowEDT850T2.Item("EDI_UPC") & String.Empty).ToString.Trim
                            rowEDT856O4.Item("PICK_QTY_CONF") = Convert.ToInt16(Val(rowCarton.Item("QTY") & String.Empty) / caseFactor)

                            ' This happens when TSI places less than 1/2 prepak in a carton.
                            If rowEDT856O4.Item("PICK_QTY_CONF") = 0 Then
                                rowEDT856O4.Item("PICK_QTY_CONF") = 1
                            End If

                            rowEDT856O4.Item("ORDR_QTY_ORIG") = Val(rowEDT850T6.Item("EDI_SLN_QTY") & String.Empty)
                            rowEDT856O4.Item("ORIG_PRICE") = Val(rowEDT850T2.Item("EDI_PRICE") & String.Empty)
                            rowEDT856O4.Item("EDI_SKU") = (rowEDT850T2.Item("EDI_SKU") & String.Empty).ToString.Trim
                            rowEDT856O4.Item("EDI_EAN") = (rowEDT850T2.Item("EDI_EAN") & String.Empty).ToString.Trim

                            ' If EDI_PRICE_UOM is null, then use EDI_PO4_UOM, else EA
                            If (rowCarton.Item("EDI_PRICE_UOM") & String.Empty).ToString.Trim.Length > 0 Then
                                rowEDT856O4.Item("EDI_PO4_UOM") = (rowCarton.Item("EDI_PRICE_UOM") & String.Empty).ToString.Trim
                            ElseIf (rowEDT850T2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim.Length > 0 Then
                                rowEDT856O4.Item("EDI_PO4_UOM") = (rowEDT850T2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim
                            Else
                                rowEDT856O4.Item("EDI_PO4_UOM") = "EA"
                            End If

                            ' Added 06/27/2018
                            If ASCMAIN1.CLIENT = "NYA" AndAlso CartonPackRollUpCustomer Then
                                If tblSOTORDR2.Select("EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "' AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")).Length > 0 Then
                                    Dim rowSOTORDR2 As DataRow = tblSOTORDR2.Select("EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "' AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ"))(0)
                                    Dim EDI_PRICE_UOM As String = rowEDT850T2.Item("EDI_PRICE_UOM") & String.Empty
                                    Dim STYLE_UOM As String = rowSOTORDR2.Item("STYLE_UOM") & String.Empty
                                    Dim CARTON_PACK_QTY As Int16 = Val(rowSOTORDR2.Item("CARTON_PACK_QTY") & String.Empty)
                                    'If (EDI_PRICE_UOM = "CA" OrElse EDI_PRICE_UOM = "CS") AndAlso STYLE_UOM = "EA" AndAlso CARTON_PACK_QTY > 0 Then
                                    ' Changed on 6/21/2018 - As per Walter and Maria since Loblaws may send other EDI_PRICE_UOM in the EDT850T2
                                    If EDI_PRICE_UOM <> "EA" AndAlso CARTON_PACK_QTY > 0 Then
                                        rowEDT856O4.Item("PICK_QTY_CONF") = CInt(Val(rowEDT856O4.Item("PICK_QTY_CONF") & String.Empty) / CARTON_PACK_QTY)
                                    End If
                                End If
                            End If

                            rowEDT856O4.Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
                            rowEDT856O4.Item("STYLE_GTIN_CODE") = (rowEDT850T2.Item("EDI_GTIN") & String.Empty).ToString.Trim
                            rowEDT856O4.Item("COLOR_CODE") = (rowEDT850T2.Item("EDI_COLOR_CODE") & String.Empty).ToString.Trim
                            rowEDT856O4.Item("EDI_SIZE_DESC") = (rowEDT850T2.Item("EDI_SIZE_DESC") & String.Empty).ToString.Trim
                            'EDT850T2.EDI_PO_LNO  >> EDT856O4.EDI_PO_LNO
                            rowEDT856O4.Item("EDI_PO_LNO") = rowEDT850T2.Item("EDI_PO_LNO")

                            'In the 856 processing, please take EDT850T2. EDI_PO4_INNER and populate EDT856O4.EDI_PO4_INNER
                            rowEDT856O4.Item("EDI_PO4_INNER") = rowEDT850T2.Item("EDI_PO4_INNER")
                            rowEDT856O4.Item("EDI_CARTON_GRP") = rowEDT850T2.Item("EDI_CARTON_GRP")

                            'added 6/12/2018 Code in InitializeData adds field if missing from table
                            rowEDT856O4.Item("EDI_DTL_SEQ_850") = rowEDT850T2.Item("EDI_DTL_SEQ")

                            tblEDT856O4.Rows.Add(rowEDT856O4)

                            ' Create EDT856O6 entries using the EDT850T6 entries
                            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM EDT850T6 WHERE EDI_DOC_SEQ_NO = :PARM1 AND EDI_DTL_SEQ = :PARM1", "EDT850T6", "VN", New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ}).Rows

                                Dim rowEDT856O6 As DataRow = tblEDT856O6.NewRow
                                rowEDT856O6.Item("COMPANY_CODE") = COMPANY_CODE
                                rowEDT856O6.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                                rowEDT856O6.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                                rowEDT856O6.Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                                rowEDT856O6.Item("EDI_HL4_SEQ") = EDI_HL4_SEQ
                                rowEDT856O6.Item("CART_LNO") = CART_LNO
                                rowEDT856O6.Item("SLN_SEQ") = row.Item("EDI_SLN_SEQ")
                                rowEDT856O6.Item("EDI_PARENT_UPC") = row.Item("EDI_PARENT_UPC")
                                rowEDT856O6.Item("EDI_PARENT_SKU") = row.Item("EDI_PARENT_SKU")
                                rowEDT856O6.Item("EDI_SLN_QTY") = row.Item("EDI_SLN_QTY")
                                rowEDT856O6.Item("EDI_SLN_UOM") = row.Item("EDI_SLN_UOM")
                                rowEDT856O6.Item("EDI_SLN_PRICE") = row.Item("EDI_SLN_PRICE")
                                rowEDT856O6.Item("EDI_SLN_SKU") = row.Item("EDI_SLN_SKU")
                                rowEDT856O6.Item("EDI_SLN_UPC") = row.Item("EDI_SLN_UPC")
                                rowEDT856O6.Item("EDI_SLN_ITEM") = row.Item("EDI_SLN_ITEM")
                                rowEDT856O6.Item("EDI_SLN_SIZE_DESC") = row.Item("EDI_SLN_SIZE_DESC")
                                rowEDT856O6.Item("EDI_SLN_PO4_UOM") = row.Item("EDI_SLN_PO4_UOM")
                                rowEDT856O6.Item("EDI_PO4_QTY") = row.Item("EDI_PO4_QTY")
                                rowEDT856O6.Item("EDI_PO4_INNER") = row.Item("EDI_PO4_INNER")
                                rowEDT856O6.Item("EDI_SLN_COLOR") = row.Item("EDI_SLN_COLOR")
                                rowEDT856O6.Item("EDI_SLN_LBL_CODE") = row.Item("EDI_SLN_LBL_CODE")
                                rowEDT856O6.Item("EDI_SLN_STYLE") = row.Item("EDI_SLN_STYLE")
                                rowEDT856O6.Item("EDI_SLN_ITEM_DESC") = row.Item("EDI_SLN_ITEM_DESC")
                                rowEDT856O6.Item("EDI_SLN_RETAIL_PRICE") = row.Item("EDI_SLN_RETAIL_PRICE")
                                rowEDT856O6.Item("EDI_SLN_PO_LNO") = row.Item("EDI_SLN_PO_LNO")
                                rowEDT856O6.Item("EDI_SLN_DEPT") = row.Item("EDI_SLN_DEPT")
                                rowEDT856O6.Item("EDI_SLN_LINE_MODE") = row.Item("EDI_SLN_LINE_MODE")
                                ' In the ASN when taking from EDT850T6 to build EDT856T6
                                ' Please add this new field EDI_SLN_ID - 5/28/2013
                                rowEDT856O6.Item("EDI_SLN_ID") = row.Item("EDI_SLN_ID")
                                '5/30 add EDI_SLN_COLOR_CODE, EDI_SLN_SIZE_CODE, EDI_SLN_BUYER_ITEM
                                rowEDT856O6.Item("EDI_SLN_COLOR_CODE") = row.Item("EDI_SLN_COLOR_CODE")
                                rowEDT856O6.Item("EDI_SLN_SIZE_CODE") = row.Item("EDI_SLN_SIZE_CODE")
                                rowEDT856O6.Item("EDI_SLN_BUYER_ITEM") = row.Item("EDI_SLN_BUYER_ITEM")
                                tblEDT856O6.Rows.Add(rowEDT856O6)
                            Next
                        Next
                    Else
                        For Each rowSOTCART2 As DataRow In tblSOTCART2.Select("CART_NO = '" & CART_NO & "'", "CART_LNO")
                            EDI_DOC_SEQ_NO = rowSOTCART2.Item("EDI_DOC_SEQ_NO") & String.Empty
                            EDI_DTL_SEQ = Val(rowSOTCART2.Item("EDI_DTL_SEQ") & String.Empty)

                            If EDI_PROMOTION = False Then
                                EDI_PROMOTION = True
                                ' prevents duplicate lookups
                                If EDI_DOC_SEQ_NO_PROMO <> EDI_DOC_SEQ_NO Then
                                    rowEDT850T1 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = :PARM1", "V", New Object() {EDI_DOC_SEQ_NO})
                                    EDI_DOC_SEQ_NO_PROMO = EDI_DOC_SEQ_NO
                                End If
                                If rowEDT850T1 IsNot Nothing Then
                                    rowEDT856O2.Item("EDI_PROMOTION") = rowEDT850T1.Item("EDI_PROMOTION") & String.Empty
                                    rowEDT856O2.Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE") & String.Empty
                                End If
                            End If

                            ' Can we add items to an EDI Order? If so, do we send the new items over
                            rowEDT850T2 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T2 WHERE EDI_DOC_SEQ_NO = :PARM1 AND EDI_DTL_SEQ = :PARM1", "VN", New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ})

                            caseFactor = 1
                            Dim STYLE_CODE As String = rowSOTCART2.Item("STYLE_CODE") & String.Empty
                            Dim rowEDTPPKS1 As DataRow = Nothing
                            If tblEDTPPKS1.Select("CUST_CODE = '" & CUST_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "'").Length > 0 Then
                                rowEDTPPKS1 = tblEDTPPKS1.Select("CUST_CODE = '" & CUST_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "'")(0)
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

                            rowEDT856O4 = tblEDT856O4.NewRow
                            rowEDT856O4.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT856O4.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            rowEDT856O4.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                            rowEDT856O4.Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                            EDI_HL4_SEQ += 1
                            rowEDT856O4.Item("EDI_HL4_SEQ") = EDI_HL4_SEQ
                            rowEDT856O4.Item("CART_LNO") = rowSOTCART2.Item("CART_LNO")
                            rowEDT856O4.Item("STYLE_UPC_CODE") = (rowEDT850T2.Item("EDI_UPC") & String.Empty).trim
                            rowEDT856O4.Item("STYLE_CODE") = (rowSOTCART2.Item("CUST_STYLE_CODE") & String.Empty).trim
                            rowEDT856O4.Item("PICK_QTY_CONF") = Convert.ToInt16(Val(rowSOTCART2.Item("QTY_PACKED") & String.Empty) / caseFactor)
                            rowEDT856O4.Item("ORDR_QTY_ORIG") = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & String.Empty)
                            rowEDT856O4.Item("STYLE_DESC") = (rowSOTCART2.Item("STYLE_DESC") & String.Empty).ToString.PadRight(35, " ").Substring(0, 35).Trim
                            rowEDT856O4.Item("STYLE_RETAIL") = Val(rowSOTCART2.Item("STYLE_RETAIL") & String.Empty)
                            rowEDT856O4.Item("ORIG_PRICE") = Val(rowEDT850T2.Item("EDI_PRICE") & String.Empty)
                            rowEDT856O4.Item("EDI_SKU") = (rowEDT850T2.Item("EDI_SKU") & String.Empty).trim
                            rowEDT856O4.Item("EDI_EAN") = (rowEDT850T2.Item("EDI_EAN") & String.Empty).ToString.Trim
                            rowEDT856O4.Item("ITEM_EXP_DATE") = rowSOTCART2.Item("ITEM_EXP_DATE")

                            ' Regency Wayfair Set Quantity
                            If ASCMAIN1.CLIENT = "RGI" Then
                                Dim SET_QTY As Int16 = Val(rowSOTCART2.Item("SET_QTY") & String.Empty)
                                If SET_QTY > 1 Then
                                    rowEDT856O4.Item("PICK_QTY_CONF") = Val(rowEDT856O4.Item("PICK_QTY_CONF") & String.Empty) / SET_QTY
                                    'rowEDT856O4.Item("ORDR_QTY_ORIG") = Val(rowEDT856O4.Item("ORDR_QTY_ORIG") & String.Empty) / SET_QTY
                                    'rowEDT856O4.Item("ORIG_PRICE") = Val(rowEDT856O4.Item("ORIG_PRICE") & String.Empty) * SET_QTY
                                End If
                            End If

                            ' If EDI_PRICE_UOM is null, then use EDI_PO4_UOM, else EA
                            If (rowEDT850T2.Item("EDI_PRICE_UOM") & String.Empty).ToString.Trim.Length > 0 Then
                                rowEDT856O4.Item("EDI_PO4_UOM") = (rowEDT850T2.Item("EDI_PRICE_UOM") & String.Empty).ToString.Trim
                            ElseIf (rowEDT850T2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim.Length > 0 Then
                                rowEDT856O4.Item("EDI_PO4_UOM") = (rowEDT850T2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim
                            Else
                                rowEDT856O4.Item("EDI_PO4_UOM") = "EA"
                            End If

                            ' Added 06/27/2018
                            If ASCMAIN1.CLIENT = "NYA" AndAlso CartonPackRollUpCustomer Then
                                If tblSOTORDR2.Select("EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "' AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")).Length > 0 Then
                                    Dim rowSOTORDR2 As DataRow = tblSOTORDR2.Select("EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "' AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ"))(0)
                                    Dim EDI_PRICE_UOM As String = rowEDT850T2.Item("EDI_PRICE_UOM") & String.Empty
                                    Dim STYLE_UOM As String = rowSOTORDR2.Item("STYLE_UOM") & String.Empty
                                    Dim CARTON_PACK_QTY As Int16 = Val(rowSOTORDR2.Item("CARTON_PACK_QTY") & String.Empty)
                                    'If (EDI_PRICE_UOM = "CA" OrElse EDI_PRICE_UOM = "CS") AndAlso STYLE_UOM = "EA" AndAlso CARTON_PACK_QTY > 0 Then
                                    ' Changed on 6/21/2018 - As per Walter and Maria since Loblaws may send other EDI_PRICE_UOM in the EDT850T2
                                    If EDI_PRICE_UOM <> "EA" AndAlso CARTON_PACK_QTY > 0 Then
                                        rowEDT856O4.Item("PICK_QTY_CONF") = CInt(Val(rowEDT856O4.Item("PICK_QTY_CONF") & String.Empty) / CARTON_PACK_QTY)
                                    End If
                                End If
                            End If

                            rowEDT856O4.Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
                            rowEDT856O4.Item("STYLE_GTIN_CODE") = (rowEDT850T2.Item("EDI_GTIN") & String.Empty).ToString.Trim
                            rowEDT856O4.Item("COLOR_CODE") = (rowEDT850T2.Item("EDI_COLOR_CODE") & String.Empty).ToString.Trim
                            rowEDT856O4.Item("EDI_SIZE_DESC") = (rowEDT850T2.Item("EDI_SIZE_DESC") & String.Empty).ToString.Trim
                            'EDT850T2.EDI_PO_LNO  >> EDT856O4.EDI_PO_LNO
                            rowEDT856O4.Item("EDI_PO_LNO") = rowEDT850T2.Item("EDI_PO_LNO")

                            'In the 856 processing, please take EDT850T2. EDI_PO4_INNER and populate EDT856O4.EDI_PO4_INNER
                            rowEDT856O4.Item("EDI_PO4_INNER") = rowEDT850T2.Item("EDI_PO4_INNER")
                            rowEDT856O4.Item("EDI_ITEM") = rowEDT850T2.Item("EDI_ITEM") & String.Empty
                            rowEDT856O4.Item("EDI_CARTON_GRP") = rowEDT850T2.Item("EDI_CARTON_GRP")

                            'added 6/12/2018 Code in InitializeData adds field if missing from table
                            rowEDT856O4.Item("EDI_DTL_SEQ_850") = rowEDT850T2.Item("EDI_DTL_SEQ")

                            tblEDT856O4.Rows.Add(rowEDT856O4)

                            ' Create EDT856O6 entries using the EDT850T6 entries
                            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM EDT850T6 WHERE EDI_DOC_SEQ_NO = :PARM1 AND EDI_DTL_SEQ = :PARM1", "EDT850T6", "VN", New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ}).Rows

                                Dim rowEDT856O6 As DataRow = tblEDT856O6.NewRow
                                rowEDT856O6.Item("COMPANY_CODE") = COMPANY_CODE
                                rowEDT856O6.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                                rowEDT856O6.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                                rowEDT856O6.Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                                rowEDT856O6.Item("EDI_HL4_SEQ") = EDI_HL4_SEQ
                                rowEDT856O6.Item("CART_LNO") = rowSOTCART2.Item("CART_LNO")
                                rowEDT856O6.Item("SLN_SEQ") = row.Item("EDI_SLN_SEQ")
                                rowEDT856O6.Item("EDI_PARENT_UPC") = row.Item("EDI_PARENT_UPC")
                                rowEDT856O6.Item("EDI_PARENT_SKU") = row.Item("EDI_PARENT_SKU")
                                rowEDT856O6.Item("EDI_SLN_QTY") = row.Item("EDI_SLN_QTY")
                                rowEDT856O6.Item("EDI_SLN_UOM") = row.Item("EDI_SLN_UOM")
                                rowEDT856O6.Item("EDI_SLN_PRICE") = row.Item("EDI_SLN_PRICE")
                                rowEDT856O6.Item("EDI_SLN_SKU") = row.Item("EDI_SLN_SKU")
                                rowEDT856O6.Item("EDI_SLN_UPC") = row.Item("EDI_SLN_UPC")
                                rowEDT856O6.Item("EDI_SLN_ITEM") = row.Item("EDI_SLN_ITEM")
                                rowEDT856O6.Item("EDI_SLN_SIZE_DESC") = row.Item("EDI_SLN_SIZE_DESC")
                                rowEDT856O6.Item("EDI_SLN_PO4_UOM") = row.Item("EDI_SLN_PO4_UOM")
                                rowEDT856O6.Item("EDI_PO4_QTY") = row.Item("EDI_PO4_QTY")
                                rowEDT856O6.Item("EDI_PO4_INNER") = row.Item("EDI_PO4_INNER")
                                rowEDT856O6.Item("EDI_SLN_COLOR") = row.Item("EDI_SLN_COLOR")
                                rowEDT856O6.Item("EDI_SLN_LBL_CODE") = row.Item("EDI_SLN_LBL_CODE")
                                rowEDT856O6.Item("EDI_SLN_STYLE") = row.Item("EDI_SLN_STYLE")
                                rowEDT856O6.Item("EDI_SLN_ITEM_DESC") = row.Item("EDI_SLN_ITEM_DESC")
                                rowEDT856O6.Item("EDI_SLN_RETAIL_PRICE") = row.Item("EDI_SLN_RETAIL_PRICE")
                                rowEDT856O6.Item("EDI_SLN_PO_LNO") = row.Item("EDI_SLN_PO_LNO")
                                rowEDT856O6.Item("EDI_SLN_DEPT") = row.Item("EDI_SLN_DEPT")
                                rowEDT856O6.Item("EDI_SLN_LINE_MODE") = row.Item("EDI_SLN_LINE_MODE")
                                ' In the ASN when taking from EDT850T6 to build EDT856T6
                                ' Please add this new field EDI_SLN_ID - 5/28/2013
                                rowEDT856O6.Item("EDI_SLN_ID") = row.Item("EDI_SLN_ID")
                                '5/30 add EDI_SLN_COLOR_CODE, EDI_SLN_SIZE_CODE, EDI_SLN_BUYER_ITEM
                                rowEDT856O6.Item("EDI_SLN_COLOR_CODE") = row.Item("EDI_SLN_COLOR_CODE")
                                rowEDT856O6.Item("EDI_SLN_SIZE_CODE") = row.Item("EDI_SLN_SIZE_CODE")
                                rowEDT856O6.Item("EDI_SLN_BUYER_ITEM") = row.Item("EDI_SLN_BUYER_ITEM")
                                tblEDT856O6.Rows.Add(rowEDT856O6)
                            Next
                        Next ' End SOTCART2
                    End If
                Next ' End SOTCART1

                ' Only one entry for Ship To and Ship From
                ' Ed - When you get to creating EDT810O5 and EDT856O5 rows, I am thinking maybe we should take the rows from EDT850T5 original PO
                If EDI_HL2_SEQ = 1 Then
                    ' Ship From
                    rowEDT856O5 = tblEDT856O5.NewRow
                    rowEDT856O5.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT856O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT856O5.Item("EDI_HL2_SEQ") = 0 'EDI_HL2_SEQ
                    EDI_ADR_SEQ += 1
                    rowEDT856O5.Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                    rowEDT856O5.Item("EDI_ADDR_TYPE") = "SF"
                    rowEDT856O5.Item("EDI_CUST_NAME_ADR") = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                    rowEDT856O5.Item("EDI_ADDRESS1") = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
                    rowEDT856O5.Item("EDI_ADDRESS2") = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
                    rowEDT856O5.Item("EDI_ADDRESS3") = rowICTWHSE1.Item("WHSE_ADDR3") & String.Empty
                    rowEDT856O5.Item("EDI_CITY") = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
                    rowEDT856O5.Item("EDI_STATE") = rowICTWHSE1.Item("WHSE_STATE") & String.Empty

                    rowEDT856O5.Item("EDI_ZIPCODE") = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                    If (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                        rowEDT856O5.Item("EDI_ZIPCODE") = (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Substring(0, rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                    End If

                    rowEDT856O5.Item("EDI_COUNTRY") = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
                    rowEDT856O5.Item("EDI_ADDR_CODE") = rowICTWHSE1.Item("WHSE_EDI_ID") & String.Empty
                    'rowEDT856O5.Item("EDI_ADDR_CODE_QUAL") = rowICTWHSE1.Item("WHSE_EDI_QUAL") & String.Empty
                    tblEDT856O5.Rows.Add(rowEDT856O5)

                    ' Factored Invoices may not have EDT850T5 records
                    Dim tblEDT850T5 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM EDT850T5 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    If tblEDT850T5.Rows.Count = 0 OrElse tblEDT850T5.Select("EDI_ADDR_TYPE = 'ST'").Length = 0 Then
                        ' Ship to
                        rowSOTORDR5 = tblSOTORDR5.Rows.Find(New Object() {rowSOTPICK1.Item("ORDR_NO"), "ST"})
                        If rowSOTORDR5 Is Nothing Then
                            rowSOTORDR5 = ASCDATA1.GetDataRow("select * from ARTCUST1 Where CUST_CODE = :PARM1", "V", CUST_CODE)
                        End If
                        rowEDT856O5 = tblEDT856O5.NewRow
                        rowEDT856O5.Item("COMPANY_CODE") = COMPANY_CODE
                        rowEDT856O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        rowEDT856O5.Item("EDI_HL2_SEQ") = 0 'EDI_HL2_SEQ
                        EDI_ADR_SEQ += 1
                        rowEDT856O5.Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                        rowEDT856O5.Item("EDI_ADDR_TYPE") = "ST"
                        rowEDT856O5.Item("EDI_CUST_NAME_ADR") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                        rowEDT856O5.Item("EDI_ADDRESS1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                        rowEDT856O5.Item("EDI_ADDRESS2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                        rowEDT856O5.Item("EDI_ADDRESS3") = rowSOTORDR5.Item("CUST_ADDR3") & String.Empty
                        rowEDT856O5.Item("EDI_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                        rowEDT856O5.Item("EDI_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty

                        rowEDT856O5.Item("EDI_ZIPCODE") = (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                        If (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                            rowEDT856O5.Item("EDI_ZIPCODE") = (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Replace(" ", "").Replace("-", "").Substring(0, rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                        End If

                        rowEDT856O5.Item("EDI_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
                        rowEDT856O5.Item("EDI_ADDR_CODE") = SHIP_ADDR_CODE
                        'rowEDT856O5.Item("EDI_ADDR_CODE_QUAL") = rowICTWHSE1.Item("WHSE_EDI_QUAL") & String.Empty
                        tblEDT856O5.Rows.Add(rowEDT856O5)
                    End If

                    ' Get the data from EDT850T5 
                    Dim EDI_ADDR_TYPE As List(Of String) = New List(Of String)
                    sqlCartonNoLoop = String.Empty

                    For Each rowEDT850T5 In tblEDT850T5.Select("", "EDI_ADDR_TYPE, EDI_ADR_SEQ")

                        If EDI_ADDR_TYPE.Contains(rowEDT850T5.Item("EDI_ADDR_TYPE") & String.Empty) Then
                            Continue For
                        Else
                            EDI_ADDR_TYPE.Add(rowEDT850T5.Item("EDI_ADDR_TYPE"))
                        End If

                        rowEDT856O5 = tblEDT856O5.NewRow
                        rowEDT856O5.Item("COMPANY_CODE") = COMPANY_CODE
                        rowEDT856O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        rowEDT856O5.Item("EDI_HL2_SEQ") = 0 'EDI_HL2_SEQ
                        EDI_ADR_SEQ += 1
                        rowEDT856O5.Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                        rowEDT856O5.Item("EDI_ADDR_TYPE") = rowEDT850T5.Item("EDI_ADDR_TYPE") & String.Empty
                        rowEDT856O5.Item("EDI_CUST_NAME_ADR") = rowEDT850T5.Item("EDI_CUST_NAME_ADR") & String.Empty
                        rowEDT856O5.Item("EDI_ADDRESS1") = rowEDT850T5.Item("EDI_ADDRESS1") & String.Empty
                        rowEDT856O5.Item("EDI_ADDRESS2") = rowEDT850T5.Item("EDI_ADDRESS2") & String.Empty
                        rowEDT856O5.Item("EDI_ADDRESS3") = String.Empty
                        rowEDT856O5.Item("EDI_CITY") = rowEDT850T5.Item("EDI_CITY") & String.Empty
                        rowEDT856O5.Item("EDI_STATE") = rowEDT850T5.Item("EDI_STATE") & String.Empty

                        rowEDT856O5.Item("EDI_ZIPCODE") = rowEDT850T5.Item("EDI_ZIPCODE") & String.Empty.ToString.Replace("-", "").Replace(" ", "")
                        If (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Length > rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength Then
                            rowEDT856O5.Item("EDI_ZIPCODE") = (rowEDT856O5.Item("EDI_ZIPCODE") & String.Empty).ToString.Replace(" ", "").Replace("-", "").Substring(0, rowEDT856O5.Table.Columns("EDI_ZIPCODE").MaxLength)
                        End If

                        rowEDT856O5.Item("EDI_COUNTRY") = rowEDT850T5.Item("EDI_COUNTRY") & String.Empty
                        rowEDT856O5.Item("EDI_ADDR_CODE") = rowEDT850T5.Item("EDI_ADDR_CODE") & String.Empty
                        rowEDT856O5.Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & String.Empty
                        tblEDT856O5.Rows.Add(rowEDT856O5)
                    Next
                    EDI_ADR_SEQ = 0
                End If

            Next ' End PICK_NO 
        Next

        Return EDI_OUTBOUND_DOC_NO

    End Function

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
        'rowEDTSYSIH.Item("LAST_DATE") = DateTime.Now
        'rowEDTSYSIH.Item("LAST_OPER") = ASCMAIN1.USER_ID
        tblEDTSYSIH.Rows.Add(rowEDTSYSIH)

        CreateEDTSYSIH = ediOutboundDocNo

    End Function

End Class
