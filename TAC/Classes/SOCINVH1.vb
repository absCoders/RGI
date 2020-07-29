Public Class SOCINVH1

    Private tblSOTINVH1 As DataTable = Nothing
    Private tblSOTINVH2 As DataTable = Nothing

    Private tblSOTRTRN1 As DataTable = Nothing
    Private tblSOTRTRN2 As DataTable = Nothing

    Private tblSOTPICK1 As DataTable = Nothing
    Private tblSOTPICK2 As DataTable = Nothing

    Private tblARTOPEN1 As DataTable = Nothing
    Private tblSOTSHIP1 As DataTable = Nothing

    Private tblSOTORDR5 As DataTable = Nothing
    Private tblSOTSVIA1 As DataTable = Nothing

    Private tblSOTINVH9 As DataTable = Nothing
    Private tblSOTINVHM As DataTable = Nothing
    Private tblSOTRNGA1 As DataTable = Nothing

    ' These were made available when the Invoice is getting created without an Order in memory
    ' These can be set to tables in a forms dataset
    Public tblSOTORDR1 As DataTable = Nothing
    Public tblSOTORDR2 As DataTable = Nothing

    Public tblSOTCART1 As DataTable = Nothing

    Private rowGLTPARM1 As DataRow
    Private tblTATTERM1 As DataTable = Nothing

    Private tblARTSTAX1 As DataTable = Nothing

    ''' <summary>
    ''' takes dataset and maps the necessary tables
    ''' </summary>
    ''' <param name="dst"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dst As DataSet)
        tblSOTINVH1 = dst.Tables("SOTINVH1")
        tblSOTINVH2 = dst.Tables("SOTINVH2")
        tblSOTPICK1 = dst.Tables("SOTPICK1")
        tblSOTPICK2 = dst.Tables("SOTPICK2")
        tblARTOPEN1 = dst.Tables("ARTOPEN1")
        tblSOTSHIP1 = dst.Tables("SOTSHIP1")
        tblSOTORDR5 = dst.Tables("SOTORDR5")
        tblSOTINVH9 = dst.Tables("SOTINVH9")
        tblSOTINVHM = dst.Tables("SOTINVHM")
        tblSOTRNGA1 = dst.Tables("SOTRNGA1")

        InitializeClassVariables()

        If dst.Tables.Contains("SOTORDR1") Then
            tblSOTORDR1 = dst.Tables("SOTORDR1")
        End If

        If dst.Tables.Contains("SOTORDR2") Then
            tblSOTORDR2 = dst.Tables("SOTORDR2")
        End If

        If dst.Tables.Contains("SOTCART1") Then
            tblSOTCART1 = dst.Tables("SOTCART1")
        End If

    End Sub

    ''' <summary>
    ''' used to create invoices
    ''' </summary>
    ''' <param name="SOTINVH1"></param>
    ''' <param name="SOTINVH2"></param>
    ''' <param name="SOTPICK1"></param>
    ''' <param name="SOTPICK2"></param>
    ''' <param name="ARTOPEN1"></param>
    ''' <param name="SOTSHIP1"></param>
    ''' <param name="SOTORDR5"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef SOTINVH1 As DataTable, _
                   ByRef SOTINVH2 As DataTable, _
                   ByRef SOTPICK1 As DataTable, _
                   ByRef SOTPICK2 As DataTable, _
                   ByRef ARTOPEN1 As DataTable, _
                   ByRef SOTSHIP1 As DataTable, _
                   ByRef SOTORDR5 As DataTable, _
                   ByRef SOTINVH9 As DataTable, _
                   ByRef SOTINVHM As DataTable, _
                   ByRef SOTRNGA1 As DataTable)

        tblSOTINVH1 = SOTINVH1
        tblSOTINVH2 = SOTINVH2
        tblSOTPICK1 = SOTPICK1
        tblSOTPICK2 = SOTPICK2
        tblARTOPEN1 = ARTOPEN1
        tblSOTSHIP1 = SOTSHIP1
        tblSOTORDR5 = SOTORDR5
        tblSOTINVH9 = SOTINVH9
        tblSOTINVHM = SOTINVHM
        tblSOTRNGA1 = SOTRNGA1

        InitializeClassVariables()
    End Sub

    ''' <summary>
    ''' Used to Create Credits from returns
    ''' </summary>
    ''' <param name="SOTINVH1"></param>
    ''' <param name="SOTINVH2"></param>
    ''' <param name="SOTRTRN1"></param>
    ''' <param name="SOTRTRN2"></param>
    ''' <param name="ARTOPEN1"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef SOTINVH1 As DataTable, _
                   ByRef SOTINVH2 As DataTable, _
                   ByRef SOTRTRN1 As DataTable, _
                   ByRef SOTRTRN2 As DataTable, _
                   ByRef ARTOPEN1 As DataTable)


        tblSOTINVH1 = SOTINVH1
        tblSOTINVH2 = SOTINVH2
        tblSOTRTRN1 = SOTRTRN1
        tblSOTRTRN2 = SOTRTRN2
        tblARTOPEN1 = ARTOPEN1

        InitializeClassVariables()

    End Sub

    ''' <summary>
    ''' This would be used if you are going to use CreateConsolidatedInvoice Only
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    Private Sub InitializeClassVariables()
        rowGLTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM1 WHERE GL_PARM_KEY = 'Z'")
        tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT * FROM SOTSVIA1", "SOTSVIA1")
        tblTATTERM1 = ASCDATA1.GetDataTable("SELECT * FROM TATTERM1", "TATTERM1")
        tblARTSTAX1 = ASCDATA1.GetDataTable("SELECT * FROM ARTSTAX1", "ARTSTAX1")

        tblSOTORDR1 = Nothing
        tblSOTORDR2 = Nothing
        tblSOTCART1 = Nothing
    End Sub

    Public Function CreateInvoices(ByVal SHIP_BOL_NO As String) As Int16
        Dim RFIXMSG As Boolean = False
        Return CreateInvoices(SHIP_BOL_NO, False, RFIXMSG)
    End Function

    Public Function CreateInvoices(ByVal SHIP_BOL_NO As String, ByRef RFIXMSG As Boolean) As Int16
        Return CreateInvoices(SHIP_BOL_NO, RFIXMSG, String.Empty)
    End Function

    Public Function CreateInvoices(ByVal SHIP_BOL_NO As String, ByRef RFIXMSG As Boolean, ByVal CUST_CODE As String) As Int16

        Dim rowSOTINVH1 As DataRow = Nothing
        Dim rowSOTINVH2 As DataRow = Nothing
        Dim rowARTCUST1 As DataRow = Nothing
        Dim rowSOTSHIP1 As DataRow = Nothing

        Dim WHSE_CODE As String = String.Empty
        Dim numInvoices As Int16 = 0
        Dim ORDR_GROUP_NO As String = String.Empty
        Dim CURR_CODE As String = String.Empty
        Dim CURR_EXCH_RATE As Decimal = 1
        Dim INV_NO As String = String.Empty
        Dim PPA_FREIGHT As Decimal = 0

        Dim edi810_customer As Boolean = False
        Dim MarkInvoiceAsPrinted As Boolean = False
        Dim isCustConsInv As Boolean = False
        Dim INV_NO_CONS As String = String.Empty
        Dim INV_TYPE As String = "I"
        Dim INV_MISC_CHG As Decimal = 0
        Dim foreignExchange As Boolean = False


        rowSOTSHIP1 = tblSOTSHIP1.Rows.Find(SHIP_BOL_NO)
        If rowSOTSHIP1 Is Nothing Then
            Return numInvoices
        End If

        WHSE_CODE = rowSOTSHIP1.Item("WHSE_CODE")
        ORDR_GROUP_NO = rowSOTSHIP1.Item("ORDR_GROUP_NO")

        Dim rowICTWHSE1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTWHSE1 WHERE WHSE_CODE = :PARM1", "V", New Object() {WHSE_CODE})
        Dim WHSE_PHYS_STATUS As String = rowICTWHSE1.Item("WHSE_PHYS_STATUS") & ""

        If CUST_CODE.Length = 0 Then
            Dim rowSOTORODR0 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO = :PARM1", "V", New Object() {ORDR_GROUP_NO})
            If rowSOTORODR0 Is Nothing Then
                Throw New Exception("Cannot locate Order Group No: " & ORDR_GROUP_NO)
            End If
            CUST_CODE = rowSOTORODR0.Item("CUST_CODE")
        End If

        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
        ' This means the customer has consolidated invoices by a shipment SOTSHIP1.SHIP_BOL_NO
        If rowARTCUST1.Item("CUST_CONS_INV") & String.Empty = "1" Then
            isCustConsInv = True
        End If

        Dim sql As String = String.Empty
        sql = " SELECT EDTTRPM1.*"
        sql &= " FROM EDTSLSP1, EDTTRPM1"
        sql &= " WHERE EDTSLSP1.EDI_QUAL_810 = EDTTRPM1.EDI_TP_QUAL"
        sql &= " AND EDTSLSP1.EDI_ID_810 = EDTTRPM1.EDI_TP_ID "
        sql &= " AND EDTSLSP1.CUST_CODE = '" & CUST_CODE & "'"
        sql &= " AND EDTTRPM1.EDI_DOC_NO = '810'"
        sql &= " AND EDTTRPM1.EDI_STATUS = 'P'"

        If ASCMAIN1.CLIENT = "VAN" Then
            sql = "Select * from EDTTRPM1 where CUST_CODE = '" & CUST_CODE & "'"
        Else

        End If
        edi810_customer = ASCDATA1.GetDataTable(sql).Rows.Count > 0
        MarkInvoiceAsPrinted = rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty = "N"

        Dim GL_PARM_CURR_CODE As String = rowGLTPARM1.Item("GL_PARM_CURR_CODE") & ""
        CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
        If CURR_CODE = "" OrElse CURR_CODE = GL_PARM_CURR_CODE Then
            CURR_CODE = GL_PARM_CURR_CODE
        Else
            foreignExchange = True
        End If

        ' Standard routine to get the Currency Exchange Rate
        If foreignExchange Then
            CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(rowGLTPARM1, CURR_CODE, CDate(DateTime.Now.ToString("MM/dd/yyyy")), True)
        End If

        ' Should parametize if invoices can cross periods
        If Not IsDate(rowSOTSHIP1.Item("INV_DATE") & String.Empty) Then
            rowSOTSHIP1.Item("INV_DATE") = CDate(DateTime.Now.ToString("MM/dd/yyyy"))
        End If

        Dim ORDR_YYYYPP_UPDATED As String = ConvertDateToPeriod(rowSOTSHIP1.Item("INV_DATE"))

        For Each rowSOTPICK1 As DataRow In tblSOTPICK1.Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "PICK_NO")

            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
            Dim PICK_QTY_CONF As Int32 = Val(tblSOTPICK2.Compute("SUM(PICK_QTY_CONF)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
            If PICK_QTY_CONF = 0 Then Continue For

            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
            Dim rowSOTORDR1 As DataRow = Nothing
            If tblSOTORDR1 IsNot Nothing Then
                rowSOTORDR1 = tblSOTORDR1.Rows.Find(ORDR_NO)
            End If

            If rowSOTORDR1 Is Nothing Then
                rowSOTORDR1 = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO = :PARM1", "V", New Object() {ORDR_NO})
            End If

            Dim tblICTSTYL1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ICTSTYL1 WHERE STYLE_CODE IN (SELECT STYLE_CODE FROM SOTORDR2 WHERE ORDR_NO = :PARM1)", "ICTSTYL1", "V", New Object() {ORDR_NO})
            Dim tblICTSTYC1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ICTSTYC1 WHERE (STYLE_CODE, COLOR_CODE) IN (SELECT STYLE_CODE, COLOR_CODE FROM SOTORDR2 WHERE ORDR_NO = :PARM1)", "ICTSTYC1", "V", New Object() {ORDR_NO})

            Dim SALES_DIVISION_CODE As String = rowSOTPICK1.Item("SALES_DIVISION_CODE") & String.Empty
            INV_NO = rowSOTPICK1.Item("INV_NO") & String.Empty

            If INV_NO.Length = 0 Then
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    INV_NO = ASCMAIN1.Next_Control_No("INV_NO_01")
                Else
                    INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                End If
                rowSOTPICK1.Item("INV_NO") = INV_NO
            End If

            ' Use the first Invoice as the consolidated invoice
            If isCustConsInv AndAlso INV_NO_CONS.Length = 0 Then
                INV_NO_CONS = INV_NO
            End If

            Dim INV_COGS As Decimal = 0
            Dim INV_SALES As Decimal = 0
            Dim INV_SALES_CURR As Decimal = 0

            For Each rowSOTPICK2 As DataRow In tblSOTPICK2.Select("PICK_NO = '" & PICK_NO & "'")
                Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
                Dim STYLE_CODE As String = rowSOTPICK2.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = tblICTSTYL1.Rows.Find(STYLE_CODE)
                Dim COLOR_CODE As String = rowSOTPICK2.Item("COLOR_CODE")
                Dim rowICTSTYC1_FIFO As DataRow = tblICTSTYC1.Rows.Find(New Object() {STYLE_CODE, COLOR_CODE})
                Dim ORDR_UNIT_COST As Decimal = 0

                If rowICTSTYC1_FIFO IsNot Nothing Then
                    ORDR_UNIT_COST = Val(rowICTSTYC1_FIFO.Item("STYLE_COST_FIFO") & "")
                End If

                If ORDR_UNIT_COST = 0 Then
                    ORDR_UNIT_COST = Val(rowICTSTYL1.Item("STYLE_COST") & "")
                End If

                rowSOTINVH2 = tblSOTINVH2.NewRow
                Dim rowSOTORDR2 As DataRow = tblSOTORDR2.Rows.Find(New Object() {rowSOTPICK2.Item("ORDR_NO"), rowSOTPICK2.Item("ORDR_LNO")})

                With rowSOTINVH2
                    .Item("INV_TYPE") = "I"
                    .Item("INV_NO") = INV_NO
                    .Item("INV_LNO") = rowSOTPICK2.Item("PICK_LNO")
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("ORDR_UNIT_COST") = ORDR_UNIT_COST
                    .Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED
                    .Item("STYLE_CUST_CODE") = rowICTSTYL1.Item("CUST_CODE") & String.Empty
                    .Item("RANGE_STYLE_LNO") = Val(rowSOTPICK2.Item("RANGE_STYLE_LNO") & String.Empty)
                    '.Item("LIC_CODE") = ""
                    '.Item("LIC_TYPE") = ""
                    '.Item("LIC_LAST") = ""
                    '.Item("LIC_OPER") = ""

                    ' If Foreign Exchange then use Currency Prices 
                    ' As per Walter on 6/15/2016 - this is how to calculate the Order Unit Price
                    ' 04/11/2018 changed to use SOTORDR2.ORDR_UNIT_PRICE, previously used SOTPICK2.PICK_UNIT_PRICE
                    .Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR")
                    If foreignExchange Then
                        .Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") * CURR_EXCH_RATE
                    Else
                        .Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                    End If

                    If (ASCMAIN1.CLIENT = "RGI" OrElse ASCMAIN1.CLIENT = "NYA") Then
                        .Item("ORDR_PRICE_SOURCE") = rowSOTPICK2.Item("ORDR_PRICE_SOURCE")
                        .Item("COMM_RATE") = rowSOTPICK2.Item("COMM_RATE")
                    End If

                    INV_COGS += (ORDR_QTY_SHIP * ORDR_UNIT_COST)
                    INV_SALES += (ORDR_QTY_SHIP * Val(.Item("ORDR_UNIT_PRICE") & ""))
                    INV_SALES_CURR += (ORDR_QTY_SHIP * Val(.Item("ORDR_UNIT_PRICE_CURR") & ""))
                End With

                tblSOTINVH2.Rows.Add(rowSOTINVH2)
            Next

            INV_SALES = Math.Round(INV_SALES, 2)

            Dim tblSOTORDR9 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR9 WHERE ORDR_NO = :PARM1", "", "V", New Object() {ORDR_NO})
            For Each rowSOTORDR9 As DataRow In tblSOTORDR9.Select()
                Dim RANGE_STYLE_LNO As Int32 = Val(rowSOTORDR9.Item("RANGE_STYLE_LNO"))
                Dim sqlw As String = "INV_NO = '" & INV_NO & "' and INV_TYPE = 'I' and RANGE_STYLE_LNO = " & CStr(RANGE_STYLE_LNO)
                Dim ORDR_QTY_SHIP As Int64 = Val(tblSOTINVH2.Compute("SUM(ORDR_QTY_SHIP)", sqlw) & "")
                Dim rowSOTINVH9 As DataRow = tblSOTINVH9.NewRow
                With rowSOTINVH9
                    .Item("INV_TYPE") = "I"
                    .Item("INV_NO") = INV_NO
                    .Item("RANGE_STYLE_LNO") = rowSOTORDR9.Item("RANGE_STYLE_LNO")
                    .Item("RANGE_STYLE_CODE") = rowSOTORDR9.Item("RANGE_STYLE_CODE")
                    .Item("RANGE_STYLE_QTY_SHIP") = ORDR_QTY_SHIP
                    .Item("RANGE_STYLE_PRICE") = rowSOTORDR9.Item("RANGE_STYLE_PRICE")
                    .Item("RANGE_STYLE_PP_PRICE") = rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE")
                    .Item("RANGE_STYLE_QTY_PER_PP") = rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP")

                    ' Added 04/19/2018 sInce code was commented out in SO.SOFSHIPB
                    .Item("RANGE_STYLE_PRICE_CURR") = rowSOTORDR9.Item("RANGE_STYLE_PRICE") / CURR_EXCH_RATE
                    .Item("RANGE_STYLE_PP_PRICE_CURR") = rowSOTORDR9.Item("RANGE_STYLE_PP_PRICE") / CURR_EXCH_RATE

                    Dim RANGE_STYLE_QTY_PER_PP As Int64 = Val(rowSOTORDR9.Item("RANGE_STYLE_QTY_PER_PP") & "")
                    If RANGE_STYLE_QTY_PER_PP = 0 Then 'Chances are the range is bad.  Add Audit trail Here.
                        Dim rowSOTRNGA1 As DataRow = tblSOTRNGA1.NewRow
                        rowSOTRNGA1.Item("INV_TYPE") = "I"
                        rowSOTRNGA1.Item("INV_NO") = INV_NO
                        rowSOTRNGA1.Item("RANGE_STYLE_LNO") = .Item("RANGE_STYLE_LNO")
                        rowSOTRNGA1.Item("RANGE_STYLE_CODE") = .Item("RANGE_STYLE_CODE")
                        rowSOTRNGA1.Item("RANGE_STYLE_PP_QTY_SHIP") = .Item("RANGE_STYLE_QTY_SHIP")
                        rowSOTRNGA1.Item("RANGE_STYLE_QTY_PER_PP") = .Item("RANGE_STYLE_QTY_PER_PP")
                        rowSOTRNGA1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowSOTRNGA1.Item("LAST_DATE") = DateTime.Now
                        tblSOTRNGA1.Rows.Add(rowSOTRNGA1)

                        .Item("RANGE_STYLE_QTY_PER_PP") = 1
                        .Item("RANGE_STYLE_PP_QTY_SHIP") = .Item("RANGE_STYLE_QTY_SHIP") / 1
                        RFIXMSG = True
                    Else
                        .Item("RANGE_STYLE_PP_QTY_SHIP") = .Item("RANGE_STYLE_QTY_SHIP") / RANGE_STYLE_QTY_PER_PP
                    End If
                End With
                tblSOTINVH9.Rows.Add(rowSOTINVH9)
            Next

            rowSOTINVH1 = tblSOTINVH1.NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = "I"
                .Item("INV_NO") = INV_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = rowSOTPICK1.Item("CUST_STORE_NO")
                .Item("ORDR_CUST_PO") = rowSOTPICK1.Item("ORDR_CUST_PO")
                .Item("ORDR_NO") = rowSOTPICK1.Item("ORDR_NO")
                .Item("WHSE_CODE") = rowSOTPICK1.Item("WHSE_CODE")
                .Item("INV_SALES") = INV_SALES
                .Item("INV_COGS") = INV_COGS

                ' Freight
                PPA_FREIGHT = 0
                If rowSOTPICK1.Table.Columns.Contains("PPA_FREIGHT") Then
                    PPA_FREIGHT = Val(rowSOTPICK1.Item("PPA_FREIGHT") & String.Empty)
                End If

                .Item("INV_FREIGHT") = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Val(rowSOTPICK1.Item("ORDR_FOB") & String.Empty) + PPA_FREIGHT

                If foreignExchange Then
                    .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
                    .Item("INV_FREIGHT") = Math.Round(.Item("INV_FREIGHT") / CURR_EXCH_RATE, 2)
                End If

                ' Miscellaneous Charges
                .Item("INV_MISC_CHG") = 0
                If rowSOTPICK1.Table.Columns.Contains("INV_MISC_CHG") Then
                    .Item("INV_MISC_CHG") = Val(rowSOTPICK1.Item("INV_MISC_CHG") & String.Empty)
                End If

                .Item("INV_TOTAL_AMOUNT") = INV_SALES + Val(.Item("INV_FREIGHT") & "") + Val(.Item("INV_MISC_CHG") & "")
                .Item("REASON_CODE") = "SHP"

                If Not IsDate(rowSOTSHIP1.Item("INV_DATE") & String.Empty) Then
                    rowSOTSHIP1.Item("INV_DATE") = CDate(DateTime.Now.ToString("MM/dd/yyyy"))
                End If
                .Item("INV_DATE") = CDate(rowSOTSHIP1.Item("INV_DATE")).ToShortDateString

                .Item("ORDR_DATE_UPDATED") = .Item("INV_DATE")
                .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED
                .Item("ORDR_BILL_TO_CUST") = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
                .Item("POST_CODE") = rowSOTPICK1.Item("POST_CODE")

                .Item("SHIP_BOL_NO") = rowSOTPICK1.Item("SHIP_BOL_NO")
                .Item("SALES_DIVISION_CODE") = rowSOTPICK1.Item("SALES_DIVISION_CODE")

                If edi810_customer OrElse MarkInvoiceAsPrinted Then
                    .Item("INV_PRINTED") = DateTime.Now
                End If

                '.Item("INV_810_BATCH_NO") = ""
                .Item("INV_NO_CONS") = INV_NO_CONS
                .Item("TERM_CODE") = rowSOTSHIP1.Item("TERM_CODE")
                .Item("INIT_DATE") = DateTime.Now
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("PICK_NO") = PICK_NO
                '.Item("INV_NO_REV") = ""
                .Item("CUST_FACTOR_IND") = rowSOTSHIP1.Item("CUST_FACTOR_TRANS_IND") ' IIf(Factored, "1", "0")
                .Item("CUST_SURCHARGE_IND") = rowARTCUST1.Item("CUST_SURCHARGE_IND")
                .Item("SREP_CODE") = rowSOTSHIP1.Item("SREP_CODE") ' rowSOTSHIP1.Item("SREP_CODE")
                .Item("INV_COMMENT") = rowSOTPICK1.Item("ORDR_INV_COMMENT")
                '.Item("REGISTER_XNO") = ""
                '.Item("INV_NO_REV_BY") = ""
                .Item("SREP2_CODE") = rowSOTSHIP1.Item("SREP2_CODE")
                '.Item("REVISED_CUST_STORE_NO") = ""
                '.Item("LAST_REVISED_DATE") = ""
                '.Item("LAST_REVISED_OPER") = ""
                .Item("EDI_RETRANSMIT_IND") = "0"
                '.Item("ORIG_CUST_STORE_NO") = ""
                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE


                ' Walmart gets a hard-coded 2.5% discount, on all invoices.
                ' Do not do this yet, will do only in factored 810 to Rosenthal

                ' March 2016 - Do this or Invoices
                If (CUST_CODE = "WALMART") AndAlso (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then

                    ASCMAIN1.sql = "SELECT MAX(EDI_DOC_SEQ_NO) EDI_DOC_SEQ_NO FROM SOTORDR2 WHERE ORDR_NO = '" & .Item("ORDR_NO") & "'"
                    Dim EDI_DOC_SEQ_NO As String = ASCDATA1.GetDataValue(ASCMAIN1.sql) & String.Empty
                    If EDI_DOC_SEQ_NO.Length > 0 Then
                        Dim tblEDT850T7 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM EDT850T7 WHERE EDI_DOC_SEQ_NO = :PARM1", "", "V", New Object() {EDI_DOC_SEQ_NO})
                        Dim miscDiscountMultiplier As Double = Math.Round(Val(tblEDT850T7.Compute("SUM(SAH_PERCENT)", "") & String.Empty) / 100, 4)

                        If miscDiscountMultiplier > 0 Then

                            Dim INV_MNO As Integer = 0
                            For Each row As DataRow In tblEDT850T7.Select("SAH_PERCENT <> 0")
                                Dim rowSOTINVHM As DataRow = tblSOTINVHM.NewRow
                                rowSOTINVHM.Item("INV_TYPE") = .Item("INV_TYPE")
                                rowSOTINVHM.Item("INV_NO") = .Item("INV_NO")

                                INV_MNO += 1
                                rowSOTINVHM.Item("INV_MNO") = INV_MNO

                                rowSOTINVHM.Item("MISC_CHG_CODE") = "DI"

                                ' Should we use The Walmart SAH_ALLOW_CODE as the MISC_CHG_CODE??
                                'Dim SAH_ALLOW_CODE As String = row.Item("SAH_ALLOW_CODE") & String.Empty
                                'If SAH_ALLOW_CODE = String.Empty Then
                                '    rowSOTINVHM.Item("MISC_CHG_CODE") = "WMT"
                                'Else
                                '    rowSOTINVHM.Item("MISC_CHG_CODE") = row.Item("SAH_ALLOW_CODE") & String.Empty
                                'End If
                                rowSOTINVHM.Item("MISC_CHG_DESC") = "Walmart Discount"
                                rowSOTINVHM.Item("MISC_CHG_NOTE") = row.Item("SAH_ALLOW_CODE") & String.Empty

                                miscDiscountMultiplier = Math.Round(.Item("INV_TOTAL_AMOUNT") * Val(row.Item("SAH_PERCENT") & String.Empty) / 100, 2, MidpointRounding.AwayFromZero) * -1
                                rowSOTINVHM.Item("INV_MISC_CHG") = miscDiscountMultiplier

                                ' Update SOTINVH1 record
                                .Item("INV_MISC_CHG") = Val(.Item("INV_MISC_CHG") & String.Empty) + rowSOTINVHM.Item("INV_MISC_CHG")

                                rowSOTINVHM.Item("CTL_NO") = DBNull.Value
                                rowSOTINVHM.Item("PO_ORDER_NO") = DBNull.Value
                                tblSOTINVHM.Rows.Add(rowSOTINVHM)
                            Next

                            .Item("INV_TOTAL_AMOUNT") += .Item("INV_MISC_CHG")

                        End If
                    End If
                End If

                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                .Item("INV_SALES_CURR") = INV_SALES_CURR

                If Val(.Item("INV_FREIGHT_CURR") & String.Empty) = 0 Then
                    .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT") / CURR_EXCH_RATE
                End If
                .Item("INV_MISC_CHG_CURR") = .Item("INV_MISC_CHG") / CURR_EXCH_RATE

                ' These two fields are the same
                .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_SALES_CURR") + .Item("INV_FREIGHT_CURR") + .Item("INV_MISC_CHG_CURR")
                .Item("INV_TOTAL_AMT_CURR") = .Item("INV_TOTAL_AMOUNT_CURR")

                '.Item("GST_TAX") = ""
                '.Item("GST_TAX_CURR") = ""
                '.Item("GEN_IND") = ""
                '.Item("GEN_XNO") = ""
                '.Item("GEN_DATE") = ""
                '.Item("DOCUMENTKEY") = ""
                If WHSE_PHYS_STATUS = "1" Then
                    .Item("SHIP_DURING_PHY") = "1"
                End If
                .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                .Item("ORDR_TYPE_CODE") = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                .Item("CUST_BILL_TO_CUST") = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
                .Item("WHSE_CODE_TO") = rowSOTORDR1.Item("WHSE_CODE_TO")

                .Item("CC_TRANS_ID") = rowSOTORDR1.Item("CC_TRANS_ID") & String.Empty
                .Item("CCPA_NO") = rowSOTPICK1.Item("CCPA_NO")

            End With
            tblSOTINVH1.Rows.Add(rowSOTINVH1)

            ' 04/11/2018 - Allow for GST Tax
            If ASCMAIN1.CLIENT = "NYA" Then
                Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO")
                Dim ORDR_ADDR_TYPE_ST As String = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
                Dim rowSOTORDR5 As DataRow = tblSOTORDR5.Rows.Find(New Object() {ORDR_NO, "ST"})
                If rowSOTORDR5 Is Nothing Then
                    rowSOTORDR5 = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR5 WHERE ORDR_NO = :PARM1 AND CUST_ADDR_TYPE = :PARM2", "VV", New Object() {ORDR_NO, "ST"})
                End If
                rowSOTINVH1.Item("CUST_SHIP_TO_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                Dim CUST_ADDR_CODE As String = rowSOTORDR5.Item("CUST_ADDR_CODE") & String.Empty
                Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_ADDR_TYPE = :PARM2 AND CUST_ADDR_CODE = :PARM3", "VVV", New Object() {CUST_CODE, ORDR_ADDR_TYPE_ST, CUST_ADDR_CODE})
                Dim STAX_CODE As String = String.Empty
                If rowARTCUST2 IsNot Nothing Then
                    STAX_CODE = rowARTCUST2.Item("STAX_CODE") & String.Empty
                End If

                If STAX_CODE <> "" Then
                    Dim rowARTSTAX1 As DataRow = tblARTSTAX1.Rows.Find(STAX_CODE)
                    Dim STAX_RATE As Decimal = Val(rowARTSTAX1.Item("STAX_RATE") & "")
                    Dim COUNTRY_CODE As String = rowARTSTAX1.Item("COUNTRY_CODE") & ""

                    INV_SALES_CURR = Val(rowSOTINVH1.Item("INV_SALES_CURR") & "")
                    INV_SALES = Val(rowSOTINVH1.Item("INV_SALES") & "")

                    With rowSOTINVH1
                        Dim INV_STAX_CURR As Decimal = System.Math.Round(STAX_RATE * INV_SALES_CURR / 100, 2)
                        Dim INV_STAX As Decimal = System.Math.Round(STAX_RATE * INV_SALES / 100, 2)

                        .Item("STAX_CODE") = STAX_CODE
                        .Item("STAX_RATE") = STAX_RATE
                        .Item("INV_STAX") = INV_STAX
                        .Item("INV_TOTAL_AMOUNT") += INV_STAX

                        If INV_SALES_CURR <> 0 Then
                            .Item("INV_STAX_CURR") = INV_STAX_CURR

                            If COUNTRY_CODE = "CAN" Then
                                .Item("GST_TAX") = INV_STAX
                                .Item("GST_TAX_CURR") = INV_STAX_CURR
                            End If

                            .Item("INV_TOTAL_AMT_CURR") += INV_STAX_CURR
                            .Item("INV_TOTAL_AMOUNT_CURR") += INV_STAX_CURR
                        End If
                    End With
                End If
            End If

            numInvoices += 1
            If ASCMAIN1.CLIENT = "VAN" Then
                If rowSOTINVH1.Item("ORDR_TYPE_CODE") & "" = "" Then
                    rowSOTINVH1.Item("ORDR_TYPE_CODE") = "REG"
                End If
            End If
            ' Transfer Orders do not create ARTOPEN1 records
            If rowSOTINVH1.Item("ORDR_TYPE_CODE") <> "XFR" Then
                CreateOpenAR(INV_TYPE, INV_NO, CURR_EXCH_RATE)
            End If

            rowSOTPICK1.Item("INV_NO") = INV_NO
            rowSOTPICK1.Item("PICK_SHIPPED") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
            rowSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSOTPICK1.Item("LAST_DATE") = DateTime.Now
        Next

        Return numInvoices

    End Function

    ''' <summary>
    ''' Creates the Open AR Record for the Invoice
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateOpenAR(ByVal INV_TYPE As String, ByVal INV_NO As String, ByVal CURR_EXCH_RATE As Decimal)

        Dim rowARTOPEN1 As DataRow = tblARTOPEN1.NewRow
        Dim rowSOTINVH1 As DataRow = tblSOTINVH1.Rows.Find(New Object() {INV_TYPE, INV_NO})

        Dim INV_SALES_CURR As Decimal = 0
        For Each row As DataRow In tblSOTINVH2.Select("INV_TYPE = '" & INV_TYPE & "' and INV_NO = '" & INV_NO & "'")
            INV_SALES_CURR += (Val(row.Item("ORDR_QTY_SHIP") & String.Empty) * Val(row.Item("ORDR_UNIT_PRICE_CURR") & String.Empty))
        Next

        Dim foreignExchange As Boolean = False
        Dim GL_PARM_CURR_CODE As String = rowGLTPARM1.Item("GL_PARM_CURR_CODE") & ""
        Dim CURR_CODE As String = rowSOTINVH1.Item("CURR_CODE") & ""
        If CURR_CODE = "" OrElse CURR_CODE = GL_PARM_CURR_CODE Then
            CURR_CODE = GL_PARM_CURR_CODE
        Else
            foreignExchange = True
        End If

        For Each fieldName As String In New String() _
                {"CUST_CODE", "INV_TYPE", "INV_DATE", "CUST_STORE_NO", "POST_CODE", _
                 "TERM_CODE", "SREP_CODE", "SREP2_CODE", "ORDR_TYPE_CODE", _
                 "ORDR_NO", "INV_SALES", "INV_FREIGHT", "INV_TOTAL_AMOUNT", _
                 "REASON_CODE", "INIT_OPER", "INIT_DATE", "INV_MISC_CHG", "ORDR_TYPE_CODE", "SALES_DIVISION_CODE", "INV_NO_CONS", "INV_STAX", "GST_TAX", "GST_TAX_CURR"}
            rowARTOPEN1.Item(fieldName) = rowSOTINVH1.Item(fieldName)
        Next

        rowARTOPEN1.Item("INV_TYPE") = INV_TYPE
        rowARTOPEN1.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")

        Dim rowTATTERM1 As DataRow = tblTATTERM1.Rows.Find(rowARTOPEN1.Item("TERM_CODE") & String.Empty)

        Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Nothing, rowSOTINVH1.Item("TERM_CODE") & String.Empty, rowTATTERM1, rowSOTINVH1.Item("INV_DATE"))
        'Dim INV_DUE_DATE As Date = SOCMAIN1.Calculate_INV_DUE_DATE(Nothing, rowSOTINVH1.Item("TERM_CODE") & String.Empty, rowSOTINVH1.Item("INV_DATE"))
        rowARTOPEN1.Item("INV_DUE_DATE") = INV_DUE_DATE.ToShortDateString
        rowARTOPEN1.Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")

        ' If it is a Factored Invoice set Balance = 0 if NYA
        If rowSOTINVH1.Item("CUST_FACTOR_IND") & String.Empty = "1" AndAlso ASCMAIN1.CLIENT = "NYA" Then
            rowARTOPEN1.Item("INV_BALANCE") = 0
        Else
            rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        End If

        rowARTOPEN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTOPEN1.Item("LAST_DATE") = DateTime.Now + ASCMAIN1.NowTSD

        rowARTOPEN1.Item("CUST_CODE_SO") = rowSOTINVH1.Item("CUST_CODE")
        rowARTOPEN1.Item("SEG2_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG2")
        rowARTOPEN1.Item("SEG3_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG3")
        rowARTOPEN1.Item("SEG4_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG4")
        rowARTOPEN1.Item("CURR_CODE") = CURR_CODE ' rowGLTPARM1.Item("GL_PARM_CURR_CODE")
        rowARTOPEN1.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE

        rowARTOPEN1.Item("INV_SALES_CURR") = rowSOTINVH1.Item("INV_SALES_CURR")

        rowARTOPEN1.Item("INV_DISC") = 0
        rowARTOPEN1.Item("INV_PMT") = 0
        rowARTOPEN1.Item("INV_DISC_TAKEN") = 0
        rowARTOPEN1.Item("INV_WRITE_OFF") = 0

        rowARTOPEN1.Item("INV_PMT_CURR") = 0
        rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") = 0
        rowARTOPEN1.Item("INV_WRITE_OFF_CURR") = 0
        rowARTOPEN1.Item("OPS_YYYYPP") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")

        If foreignExchange Then
            If Val(rowARTOPEN1.Item("INV_DISC") & String.Empty) <> 0 Then
                Throw New Exception("Invoice Discount for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            End If

            'If Val(rowARTOPEN1.Item("INV_FREIGHT") & String.Empty) <> 0 Then
            '    Throw New Exception("Invoice Freight for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            'End If

            'If Val(rowARTOPEN1.Item("INV_STAX") & String.Empty) <> 0 Then
            '    Throw New Exception("Invoice Sales Tax for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            'End If

            If Val(rowARTOPEN1.Item("INV_MISC_CHG") & String.Empty) <> 0 Then
                Throw New Exception("Invoice Misc Charge for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            End If

            rowARTOPEN1.Item("INV_DISC_CURR") = 0
            rowARTOPEN1.Item("INV_FREIGHT_CURR") = rowSOTINVH1.Item("INV_FREIGHT_CURR")
            rowARTOPEN1.Item("INV_STAX_CURR") = rowSOTINVH1.Item("INV_STAX_CURR")
            rowARTOPEN1.Item("INV_MISC_CHG_CURR") = 0
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR")

            If rowARTOPEN1.Item("INV_BALANCE") = 0 Then
                rowARTOPEN1.Item("INV_BALANCE_CURR") = 0
            Else
                rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_TOTAL_AMOUNT")
            End If
        Else
            rowARTOPEN1.Item("INV_DISC_CURR") = rowARTOPEN1.Item("INV_DISC")
            rowARTOPEN1.Item("INV_FREIGHT_CURR") = rowARTOPEN1.Item("INV_FREIGHT")
            rowARTOPEN1.Item("INV_STAX_CURR") = rowARTOPEN1.Item("INV_STAX")
            rowARTOPEN1.Item("INV_MISC_CHG_CURR") = rowARTOPEN1.Item("INV_MISC_CHG")
            rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE")
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowARTOPEN1.Item("INV_TOTAL_AMOUNT")
        End If

        rowTATTERM1 = tblTATTERM1.Rows.Find(rowARTOPEN1.Item("TERM_CODE") & String.Empty)
        If rowTATTERM1 IsNot Nothing _
            AndAlso Val(rowTATTERM1.Item("TERM_DISC_PERC") & String.Empty) > 0 _
            AndAlso Val(rowTATTERM1.Item("TERM_DAYS_DISC") & String.Empty) > 0 Then

            rowARTOPEN1.Item("INV_DISC_DATE") = DateAdd(DateInterval.Day, Val(rowTATTERM1.Item("TERM_DAYS_DISC") & String.Empty), rowARTOPEN1.Item("INV_DATE")).ToShortDateString
        End If

        tblARTOPEN1.Rows.Add(rowARTOPEN1)

        ' Other Columns in ARTOPEN1 that are not filled by this procedure
        'rowARTOPEN1.Item("APPLY_TO_INV_NUM") = String.Empty
        'rowARTOPEN1.Item("APPLY_TO_INV_TYPE") = String.Empty
        'rowARTOPEN1.Item("INV_LAST_PMT") = String.Empty
        'rowARTOPEN1.Item("INV_PMT") = String.Empty
        'rowARTOPEN1.Item("INV_DISC_TAKEN") = String.Empty
        'rowARTOPEN1.Item("INV_WRITE_OFF") = String.Empty
        'rowARTOPEN1.Item("INV_LAST_PMT_REF") = String.Empty
        'rowARTOPEN1.Item("INV_LAST_PMT_REF_DT") = String.Empty
        'rowARTOPEN1.Item("OPS_YYYYPP_F") = String.Empty
        'rowARTOPEN1.Item("GST_TAX") = String.Empty
        'rowARTOPEN1.Item("GST_TAX_CURR") = String.Empty
        'rowARTOPEN1.Item("ORDR_CREDIT_APPR_BY") = String.Empty
        'rowARTOPEN1.Item("ORDR_CREDIT_APPR_DATE") = String.Empty
        'rowARTOPEN1.Item("INV_NOTES") = String.Empty
        'rowARTOPEN1.Item("OPS_YYYYPP_PAID") = String.Empty

    End Sub

    ''' <summary>
    ''' Get all invoices where the Invoice Number or Consolidated Invoice Nuumber match the invoiceNumber provided
    ''' Creates a SOTINVH2 header, and SOTINVH2 detaail using the Consolidated Invoice Number
    ''' </summary>
    ''' <param name="invoiceNumber">Invoice number used to get all invoices in teh consolidation</param>
    ''' <param name="rowSOTINVH1">reference to SOTINVH1 datarow</param>
    ''' <param name="tblSOTINVH2">reference to SOTINVH2 datatable</param>
    ''' <returns>returns True if no errors; otherwise, returns false</returns>
    ''' <remarks></remarks>
    Public Function CreateConsolidatedInvoice(ByVal invoiceNumber As String, _
                                              ByRef rowSOTINVH1 As DataRow, _
                                              ByRef tblSOTINVH2 As DataTable, _
                                              Optional useSLN As Boolean = False) As Boolean

        Try

            ' Row and table must have the primay key fields
            If Not rowSOTINVH1.Table.Columns.Contains("INV_NO") Then
                Return False
            End If

            If Not rowSOTINVH1.Table.Columns.Contains("INV_TYPE") Then
                Return False
            End If

            If Not tblSOTINVH2.Columns.Contains("INV_NO") Then
                Return False
            End If

            If Not tblSOTINVH2.Columns.Contains("INV_TYPE") Then
                Return False
            End If

            Dim sqlInvoices As String = "Select * from Sotinvh1 where Inv_no = :PARM1 or INV_NO_CONS = :PARM2"

            Dim tblHeader As DataTable = ASCDATA1.GetDataTable(sqlInvoices, "SOTINVH1", _
                                                               "VV", New Object() {invoiceNumber, invoiceNumber})

            tblHeader.PrimaryKey = New System.Data.DataColumn() {tblHeader.Columns("INV_NO")}

            If tblHeader.Rows.Count <= 1 Then
                Return True
            End If

            Dim tblDetails As DataTable = Nothing
            Dim tblcartons As DataTable = Nothing
            Dim sql As String = String.Empty

            If Not useSLN Then
                sql = "Select * from Sotinvh2 where Inv_no in (" & sqlInvoices.Replace("*", "INV_NO") & ")"
                tblDetails = ASCDATA1.GetDataTable(sql, "SOTINVH2", "VV", New Object() {invoiceNumber, invoiceNumber})
            Else
                sql = "Select SOTINVH2.*, SOTINVH1.ORDR_NO, SOTORDR2.EDI_DOC_SEQ_NO, SOTORDR2.EDI_DTL_SEQ, SOTORDR2.STYLE_UOM"
                sql &= " from SOTINVH1, SOTINVH2, SOTORDR2"
                sql &= " where SOTINVH1.inv_type = SOTINVH2.inv_type"
                sql &= " and SOTINVH1.inv_no = SOTINVH2.inv_no"
                sql &= " and sotordr2.ordr_lno = sotinvh2.inv_lno"
                sql &= " and sotordr2.ordr_no = sotinvh1.ordr_no"
                sql &= " and SOTINVH2.Inv_no in (" & sqlInvoices.Replace("*", "INV_NO") & ")"
                tblDetails = ASCDATA1.GetDataTable(sql, "SOTINVH2", "VV", New Object() {invoiceNumber, invoiceNumber})
                tblDetails.Columns.Add("EXTENDED", GetType(System.Double), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY_SHIP,0)")

                Dim ediData As String = String.Empty
                For Each row As DataRow In ASCDATA1.SelectDistinct(tblDetails, New String() {"ORDR_NO", "EDI_DOC_SEQ_NO", "EDI_DTL_SEQ"}).Rows
                    Dim ORDR_NO As String = (row.Item("ORDR_NO") & String.Empty).ToString.Trim
                    Dim EDI_DOC_SEQ_NO As String = (row.Item("EDI_DOC_SEQ_NO") & String.Empty).ToString.Trim
                    Dim EDI_DTL_SEQ As Int16 = Val(row.Item("EDI_DTL_SEQ") & String.Empty)

                    If ORDR_NO.Length > 0 AndAlso EDI_DOC_SEQ_NO.Length > 0 AndAlso EDI_DTL_SEQ > 0 Then
                        ediData &= ", ('" & ORDR_NO & "', '" & EDI_DOC_SEQ_NO & "', " & EDI_DTL_SEQ & ")"
                    End If
                Next

                If ediData.Length > 0 Then
                    ediData = ediData.Substring(1).Trim
                End If

                If tblcartons IsNot Nothing Then
                    tblcartons.Rows.Clear()
                End If

                'sql = " SELECT SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ"
                'sql &= " , NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM) EDI_PRICE_UOM"
                'sql &= " , SUM(SOTCART2.QTY_PACKED) / SUM(EDT850T6.EDI_SLN_QTY) QTY"
                'sql &= " FROM SOTORDR2, EDT850T6, EDT850T2, SOTCART2"
                'sql &= " where EDT850T6.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO AND EDT850T6.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ"
                'sql &= " AND EDT850T6.EDI_SLN_SEQ  = SOTORDR2.EDI_SLN_SEQ"
                'sql &= " AND EDT850T2.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO AND EDT850T2.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ"
                'sql &= " AND SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO AND SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO"
                'sql &= " AND (SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ) IN "
                'sql &= " (" & ediData & ")"
                'sql &= " GROUP BY SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ, NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM)"

                ' De-Released pick Tickets with Cartons cause the numbers to be incorrect.
                sql = " SELECT SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ, EDT850T6.EDI_SLN_QTY,"
                sql &= " NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM) EDI_PRICE_UOM , SUM(SOTCART2.QTY_PACKED) QTY_PACKED"
                sql &= " FROM SOTORDR2, EDT850T6, EDT850T2, SOTCART2, SOTCART1, SOTPICK1 "
                sql &= " where EDT850T6.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO "
                sql &= " and EDT850T6.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ "
                sql &= " and EDT850T6.EDI_SLN_SEQ  = SOTORDR2.EDI_SLN_SEQ "
                sql &= " and EDT850T2.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO "
                sql &= " and EDT850T2.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ "
                sql &= " and SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO "
                sql &= " and SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO "
                sql &= " and SOTCART1.CART_NO = SOTCART2.CART_NO"
                sql &= " and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
                sql &= " and SOTPICK1.PICK_STATUS = 'F'"
                sql &= " and (SOTORDR2.ORDR_NO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ)  IN "
                sql &= " (" & ediData & ")"
                sql &= " GROUP BY SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ, EDT850T6.EDI_SLN_QTY, NVL(EDT850T2.EDI_PRICE_UOM,EDT850T2.EDI_PO4_UOM)"
                Dim wkSLN As String = ASCMAIN1.Temp_Table(sql)

                sql = " SELECT ORDR_NO, EDI_DOC_SEQ_NO, EDI_DTL_SEQ, EDI_PRICE_UOM, SUM(QTY_PACKED) / SUM(EDI_SLN_QTY) QTY"
                sql &= " FROM  " & wkSLN
                sql &= " GROUP BY ORDR_NO, EDI_DOC_SEQ_NO, EDI_DTL_SEQ, EDI_PRICE_UOM"
 
                tblcartons = ASCDATA1.GetDataTable(sql)

                For Each row As DataRow In tblcartons.Select("", "ORDR_NO, EDI_DOC_SEQ_NO, EDI_DTL_SEQ")
                    Dim ORDR_NO As String = (row.Item("ORDR_NO") & String.Empty).ToString.Trim
                    Dim EDI_DOC_SEQ_NO As String = (row.Item("EDI_DOC_SEQ_NO") & String.Empty).ToString.Trim
                    Dim EDI_DTL_SEQ As Int16 = Val(row.Item("EDI_DTL_SEQ") & String.Empty)

                    sql = "ORDR_NO = '" & ORDR_NO & "' and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_DTL_SEQ = " & EDI_DTL_SEQ
                    Dim qty As Int32 = Val(tblcartons.Compute("SUM(QTY)", sql) & String.Empty)
                    Dim ordrUnitPrice As Double = Val(tblDetails.Compute("SUM(EXTENDED)", sql) & String.Empty) / qty ' Val(row.Item("QTY") & String.Empty)

                    ' Set Order Qty Shipped to 0, left over lines will be deleted later
                    For Each rowDetails As DataRow In tblDetails.Select(sql, "EDI_DTL_SEQ")
                        rowDetails.Item("ORDR_QTY_SHIP") = 0
                    Next

                    ' Apply to only one Line item
                    For Each rowDetails As DataRow In tblDetails.Select(sql, "EDI_DTL_SEQ")
                        rowDetails.Item("ORDR_QTY_SHIP") = qty 'row.Item("QTY")
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
            End If

            tblDetails.PrimaryKey = New System.Data.DataColumn() {tblDetails.Columns("INV_TYPE"), tblDetails.Columns("INV_NO"), tblDetails.Columns("INV_LNO")}

            Dim tblSOTINVH2wk As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTINVH2 WHERE ROWNUM < 1", "SOTINVH2")

            ' use the Consolidated Invoice as the Header Row. Will modify $ and totals later
            Dim rowHeader As DataRow = tblHeader.Rows.Find(invoiceNumber)

            Dim INV_NO As String = rowHeader.Item("INV_NO")
            Dim INV_LNO As Int32 = 0

            Dim STYLE_CODE As String = String.Empty
            Dim COLOR_CODE As String = String.Empty
            Dim firstRow As DataRow = Nothing
            For Each rowItem As DataRow In ASCDATA1.SelectDistinct(tblDetails, New String() {"STYLE_CODE", "COLOR_CODE"}).Rows

                STYLE_CODE = rowItem.Item("STYLE_CODE") & String.Empty
                COLOR_CODE = rowItem.Item("COLOR_CODE") & String.Empty

                firstRow = tblDetails.Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")(0)

                Dim ORDR_UNIT_PRICE As Decimal = 0
                Dim ORDR_QTY_SHIP As Int32 = 0

                For Each rowShipped As DataRow In tblDetails.Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")
                    If Val(rowShipped.Item("ORDR_QTY_SHIP") & String.Empty) <> 0 Then
                        ORDR_QTY_SHIP += Val(rowShipped.Item("ORDR_QTY_SHIP") & String.Empty)
                        ORDR_UNIT_PRICE += Val(rowShipped.Item("ORDR_QTY_SHIP") & String.Empty) * Val(rowShipped.Item("ORDR_UNIT_PRICE") & String.Empty)
                    End If
                Next

                If ORDR_QTY_SHIP = 0 Then Continue For

                Dim rowSOTINVH2 As DataRow = tblSOTINVH2wk.NewRow
                INV_LNO += 1
                rowSOTINVH2.Item("INV_TYPE") = rowHeader.Item("INV_TYPE") & String.Empty
                rowSOTINVH2.Item("INV_NO") = INV_NO
                rowSOTINVH2.Item("INV_LNO") = INV_LNO
                rowSOTINVH2.Item("STYLE_CODE") = STYLE_CODE
                rowSOTINVH2.Item("COLOR_CODE") = COLOR_CODE
                rowSOTINVH2.Item("ORDR_UNIT_COST") = firstRow.Item("ORDR_UNIT_COST")
                rowSOTINVH2.Item("ORDR_UNIT_PRICE") = Math.Round(ORDR_UNIT_PRICE / ORDR_QTY_SHIP, 2)
                rowSOTINVH2.Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
                rowSOTINVH2.Item("CUST_CODE") = firstRow.Item("CUST_CODE")
                rowSOTINVH2.Item("ORDR_YYYYPP_UPDATED") = firstRow.Item("ORDR_YYYYPP_UPDATED")
                rowSOTINVH2.Item("STYLE_CUST_CODE") = firstRow.Item("STYLE_CUST_CODE")
                rowSOTINVH2.Item("RANGE_STYLE_LNO") = firstRow.Item("RANGE_STYLE_LNO")
                rowSOTINVH2.Item("LIC_CODE") = firstRow.Item("LIC_CODE")
                rowSOTINVH2.Item("LIC_TYPE") = firstRow.Item("LIC_TYPE")
                rowSOTINVH2.Item("LIC_LAST") = firstRow.Item("LIC_LAST")
                rowSOTINVH2.Item("LIC_OPER") = firstRow.Item("LIC_OPER")
                rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") = firstRow.Item("ORDR_UNIT_PRICE_CURR")
                tblSOTINVH2wk.Rows.Add(rowSOTINVH2)
            Next

            If useSLN Then
                Dim INV_SALES As Double = Val(tblHeader.Compute("SUM(INV_SALES)", "") & String.Empty)

                tblSOTINVH2wk.Columns.Add("INV_SALES_SLN", GetType(System.Double), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY_SHIP,0)")
                Dim INV_SALES_SLN As Double = Val(tblSOTINVH2wk.Compute("SUM(INV_SALES_SLN)", "") & String.Empty)

                If Math.Abs(Math.Round(INV_SALES, 0, MidpointRounding.AwayFromZero) - Math.Round(INV_SALES_SLN, 0, MidpointRounding.AwayFromZero)) > 1 Then
                    Throw New Exception("Consolidated Invoice for SLN not matching.")
                End If

            End If

            rowHeader.Item("INV_SALES") = Val(tblHeader.Compute("SUM(INV_SALES)", "") & String.Empty)
            rowHeader.Item("INV_COGS") = Val(tblHeader.Compute("SUM(INV_COGS)", "") & String.Empty)
            rowHeader.Item("INV_FREIGHT") = Val(tblHeader.Compute("SUM(INV_FREIGHT)", "") & String.Empty)
            rowHeader.Item("INV_MISC_CHG") = Val(tblHeader.Compute("SUM(INV_MISC_CHG)", "") & String.Empty)
            rowHeader.Item("INV_TOTAL_AMOUNT") = Val(tblHeader.Compute("SUM(INV_TOTAL_AMOUNT)", "") & String.Empty)

            ' Update header row
            For Each col As DataColumn In rowHeader.Table.Columns
                If rowSOTINVH1.Table.Columns.Contains(col.ColumnName) Then
                    rowSOTINVH1.Item(col.ColumnName) = rowHeader.Item(col.ColumnName)
                End If
            Next

            ' Get a list of column names in common
            Dim colList As New List(Of String)
            For Each col As DataColumn In tblDetails.Columns
                If tblSOTINVH2.Columns.Contains(col.ColumnName) Then
                    colList.Add(col.ColumnName)
                End If
            Next

            ' Create the invoice details
            For Each row As DataRow In tblSOTINVH2.Select("INV_NO = '" & invoiceNumber & "'")
                row.Delete()
            Next
            tblSOTINVH2.AcceptChanges()

            For Each row As DataRow In tblSOTINVH2wk.Select("")
                Dim rowSOTINVH2 As DataRow = tblSOTINVH2.NewRow
                For Each field As String In colList
                    rowSOTINVH2.Item(field) = row.Item(field)
                Next
                tblSOTINVH2.Rows.Add(rowSOTINVH2)
            Next

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    ''' <summary>
    '''  Creates a Credit based on a return
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateReturnsCredit(ByVal ORDR_TYPE_CODE As String) As Int16

        Dim numReturns As Int16 = 0
        Dim INV_TYPE As String = "C"

        Dim rowSOTINVH1 As DataRow = Nothing
        Dim rowSOTINVH2 As DataRow = Nothing
        Dim CURR_CODE As String = String.Empty
        Dim CURR_EXCH_RATE As Int16 = 1
        Dim foreignExchange As Boolean = False

        ORDR_TYPE_CODE = ORDR_TYPE_CODE.Trim
        If ORDR_TYPE_CODE.Length = 0 Then
            ORDR_TYPE_CODE = "REG"
        End If

        For Each rowSOTRTRN1 As DataRow In tblSOTRTRN1.Select("", "RTRN_NO")

            Dim RTRN_NO As String = rowSOTRTRN1.Item("RTRN_NO")
            Dim INV_COGS As Decimal = 0
            Dim RTRN_SALES As Decimal = 0
            Dim INV_NO As String = rowSOTRTRN1.Item("INV_NO") & String.Empty

            Dim ORDR_YYYYPP_UPDATED As String = ConvertDateToPeriod(DateTime.Now)
            Dim CUST_CODE As String = rowSOTRTRN1.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
            Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_STORE_NO = :PARM2", "VV", New Object() {CUST_CODE, rowSOTRTRN1.Item("CUST_STORE_NO") & String.Empty})

            Dim GL_PARM_CURR_CODE As String = rowGLTPARM1.Item("GL_PARM_CURR_CODE") & ""
            CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
            If CURR_CODE = "" OrElse CURR_CODE = GL_PARM_CURR_CODE Then
                CURR_CODE = GL_PARM_CURR_CODE
                CURR_EXCH_RATE = 1
            Else
                foreignExchange = True
            End If

            ' Standard routine to get the Currency Exchange Rate
            If foreignExchange Then
                CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(rowGLTPARM1, CURR_CODE, CDate(DateTime.Now.ToString("MM/dd/yyyy")), True)
            End If

            For Each rowSOTRTRN2 As DataRow In tblSOTRTRN2.Select("RTRN_NO = '" & RTRN_NO & "'")
                Dim ORDR_QTY_RTRN As Int32 = Val(rowSOTRTRN2.Item("RTRN_QTY_1") & "") + Val(rowSOTRTRN2.Item("RTRN_QTY_2") & "") + Val(rowSOTRTRN2.Item("RTRN_QTY_3") & "")
                If ORDR_QTY_RTRN = 0 Then Continue For

                Dim ITEM_CODE As String = rowSOTRTRN2.Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_CODE = :PARM1", "V", New Object() {ITEM_CODE})

                Dim RTN_UNIT_COST As Decimal = Val(rowSOTRTRN2.Item("ITEM_COST_STD") & "")
                Dim RTRN_PRICE As Decimal = Val(rowSOTRTRN2.Item("RTRN_PRICE") & "")
                RTRN_SALES += ORDR_QTY_RTRN * RTRN_PRICE

                rowSOTINVH2 = tblSOTINVH2.NewRow
                With rowSOTINVH2
                    .Item("INV_TYPE") = INV_TYPE
                    .Item("INV_NO") = rowSOTRTRN1.Item("INV_NO")
                    .Item("INV_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("ORDR_UNIT_PRICE") = RTRN_PRICE
                    .Item("ORDR_QTY_SHIP") = ORDR_QTY_RTRN * -1

                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO")
                    If .Item("CUST_STORE_NO") & String.Empty = String.Empty Then
                        .Item("CUST_STORE_NO") = "000000"
                    End If
                    .Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
                    .Item("SREP_CODE") = rowSOTRTRN1.Item("SREP_CODE")
                    .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                    .Item("ORDR_UNIT_PRICE_CURR") = RTRN_PRICE
                    .Item("ITEM_UNIT_COST") = RTN_UNIT_COST
                    .Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                    .Item("ITEM_RETAIL_PRICE_CURR") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                    .Item("OPS_YYYYWW") = ASCMAIN1.CYW

                    INV_COGS += (ORDR_QTY_RTRN * RTN_UNIT_COST)
                End With
                tblSOTINVH2.Rows.Add(rowSOTINVH2)
            Next

            If tblSOTINVH2.Select("INV_NO = '" & INV_NO & "'").Length = 0 Then
                Continue For
            End If

            rowSOTINVH1 = tblSOTINVH1.NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = INV_TYPE
                .Item("INV_NO") = INV_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO")
                If .Item("CUST_STORE_NO") & String.Empty = String.Empty Then
                    .Item("CUST_STORE_NO") = "000000"
                End If
                .Item("ORDR_CUST_PO") = rowSOTRTRN1.Item("CUST_CLAIM_NO")
                '.Item("ORDR_NO") = rowSOTRTRN1.Item("ORDR_NO")
                .Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
                .Item("POST_CODE") = rowSOTRTRN1.Item("POST_CODE")
                .Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                .Item("SREP_CODE") = rowSOTRTRN1.Item("SREP_CODE")
                .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE")
                .Item("REASON_CODE") = rowSOTRTRN1.Item("REASON_CODE")

                If rowARTCUST1.Item("CUST_BILL_TO_CUST") & String.Empty = String.Empty Then
                    .Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_CODE")
                Else
                    .Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_BILL_TO_CUST")
                End If

                .Item("BRAND_CODE") = String.Empty
                .Item("EVENT_CODE") = String.Empty

                .Item("INV_SALES") = Val(rowSOTRTRN1.Item("RTRN_SALES") & String.Empty) * -1
                .Item("INV_COGS") = INV_COGS * -1

                .Item("REGISTER_IND") = "0"
                .Item("SHIP_FRT_AMT_ACTUAL") = 0
                .Item("STAX_RATE") = 0
                .Item("STAX_CODE") = rowSOTRTRN1.Item("STAX_CODE")
                .Item("SHIP_FRT_AMT_ACCRUED") = 0

                .Item("INV_FREIGHT") = Val(rowSOTRTRN1.Item("RTRN_FREIGHT") & String.Empty) * -1
                .Item("INV_MISC_CHG") = Val(rowSOTRTRN1.Item("RTRN_HANDLING") & String.Empty) * -1

                .Item("INV_STAX") = Val(rowSOTRTRN1.Item("RTRN_STAX") & String.Empty) * -1
                .Item("INV_TOTAL_AMOUNT") = Val(rowSOTRTRN1.Item("RTRN_AMOUNT") & String.Empty) * -1
                .Item("INV_DATE") = CDate(DateTime.Now.ToString("MM/dd/yyyy"))
                .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                .Item("INIT_DATE") = DateTime.Now
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("REGISTER_XNO") = String.Empty
                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                '.Item("INV_DATE_SHIPPED") = .Item("INV_DATE")

                If .Item("INV_TOTAL_AMOUNT") = 0 AndAlso Not (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI") Then
                    .Item("INV_PRINTED") = DateTime.Now
                End If

                .Item("INV_CARTONS") = 0
                .Item("INV_WEIGHT") = 0
                '.Item("INV_BOL_NO") = String.Empty
                '.Item("INV_PRO_NO") = String.Empty
                '.Item("SHIP_VIA_DESC") = String.Empty
                '.Item("INV_NO_CONS") = String.Empty
                '.Item("SHIP_BOL_NO") = String.Empty

                If rowARTCUST2 IsNot Nothing Then
                    .Item("CUST_SHIP_TO_STATE") = rowARTCUST2.Item("CUST_STORE_STATE") & String.Empty
                End If

                .Item("INV_COMMENT") = rowSOTRTRN1.Item("RTRN_NOTE")
                '.Item("INV_FREIGHT_TAX") = 0
                '.Item("SHIP_VIA_CODE") = String.Empty
                .Item("ORDR_TYPE_CODE") = ORDR_TYPE_CODE
                .Item("OPS_YYYYWW") = ASCMAIN1.CYW
                '.Item("ORDR_DEPT") = String.Empty
                '.Item("CUST_FACTOR_IND") = String.Empty
                '.Item("ORDR_NO_WEB") = String.Empty
                '.Item("SALES_DIVISION_CODE") = rowSOTRTRN1.Item("SALES_DIVISION_CODE")
            End With
            tblSOTINVH1.Rows.Add(rowSOTINVH1)
            numReturns += 1
            CreateOpenAR(INV_TYPE, INV_NO, CURR_EXCH_RATE)
        Next

        Return numReturns

    End Function

    ' Converts a date to a period
    Public Function ConvertDateToPeriod(inDate As Date) As String

        Dim period As String = String.Empty
        Try
            period = ASCDATA1.GetDataValue("Select MIN(OPS_YYYYPP) from gltparm2 where prd_end_date >= '" & inDate.ToString("dd-MMM-yyyy") & "'") & String.Empty
        Catch ex As Exception
            period = ASCMAIN1.CYP
        End Try

        Return period

    End Function

    Public Sub ProcessPickTicketsAndUpdateSalesDetails(ByVal INV_DATE As Date)

        Dim rowSOTORDR1 As DataRow = Nothing

        Dim ORDR_YYYYPP_UPDATED As String = ASCDATA1.GetDataValue("Select MIN(OPS_YYYYPP) from gltparm2 where prd_end_date >= '" & INV_DATE.ToString("dd-MMM-yyyy") & "'") & String.Empty
        If ORDR_YYYYPP_UPDATED.Length = 0 Then
            ORDR_YYYYPP_UPDATED = ASCMAIN1.CYP
        End If

        Dim ORDR_QTY_CANC As Int32 = 0
        Dim ORDR_QTY_CANC_ORIG As Int32 = 0

        Dim ORDR_QTY_BACK As Int32 = 0
        Dim ORDR_QTY_BACK_ORIG As Int32 = 0

        For Each rowSOTSHIP1 As DataRow In tblSOTSHIP1.Rows
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            Dim sql As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            If tblSOTPICK1.Columns.Contains("SELECTED") Then
                sql &= " AND SELECTED = '1'"
            End If

            ' Update Sotordr1 and Sortordr2 and Possibly SOTPICK1,2 
            For Each rowSOTPICK1 As DataRow In tblSOTPICK1.Select(sql, "", DataViewRowState.CurrentRows)
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                rowSOTORDR1 = tblSOTORDR1.Rows.Find(ORDR_NO)

                ' Update Carton Weight based on the weight(s) provided by the user.
                ' Added 1/21/2013 / modified on 6/17/2013 - User may key in a higher weight; therefore, honor it.
                Dim PICK_TOTAL_WGT As Decimal = Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                If PICK_TOTAL_WGT > Val(rowSOTPICK1.Item("PICK_TOTAL_WGT") & String.Empty) Then
                    rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT
                End If

                For Each rowSOTPICK2 As DataRow In tblSOTPICK2.Select("PICK_NO = '" & PICK_NO & "'")

                    Dim ORDR_LNO As Int16 = rowSOTPICK2.Item("ORDR_LNO")
                    Dim rowSOTORDR2 As DataRow = tblSOTORDR2.Rows.Find(New Object() {ORDR_NO, ORDR_LNO})

                    ' Clear out the pick Qty from this Pick Ticket
                    rowSOTORDR2.Item("ORDR_QTY_PICK") = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty) - Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty)
                    If rowSOTORDR2.Item("ORDR_QTY_PICK") < 0 Then
                        rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                    End If

                    ' See if the user incrememnted the ORDR_QTY_CANC
                    If rowSOTPICK2.RowState = DataRowState.Added Then
                        ORDR_QTY_CANC_ORIG = Val(rowSOTPICK2.Item("PICK_QTY_CANC", DataRowVersion.Current) & String.Empty)
                    Else
                        ORDR_QTY_CANC_ORIG = Val(rowSOTPICK2.Item("PICK_QTY_CANC", DataRowVersion.Original) & String.Empty)
                    End If

                    ORDR_QTY_CANC = Val(rowSOTPICK2.Item("PICK_QTY_CANC", DataRowVersion.Current) & String.Empty)

                    If rowSOTPICK2.RowState = DataRowState.Added Then
                        ORDR_QTY_BACK_ORIG = Val(rowSOTPICK2.Item("PICK_QTY_BACK", DataRowVersion.Current) & String.Empty)
                    Else
                        ORDR_QTY_BACK_ORIG = Val(rowSOTPICK2.Item("PICK_QTY_BACK", DataRowVersion.Original) & String.Empty)
                    End If

                    ORDR_QTY_BACK = Val(rowSOTPICK2.Item("PICK_QTY_BACK", DataRowVersion.Current) & String.Empty)

                    If ORDR_QTY_CANC > ORDR_QTY_CANC_ORIG Then
                        ORDR_QTY_CANC -= ORDR_QTY_CANC_ORIG
                    Else
                        'ORDR_QTY_CANC = 0
                    End If

                    If ORDR_QTY_BACK > ORDR_QTY_BACK_ORIG Then
                        ORDR_QTY_BACK -= ORDR_QTY_BACK_ORIG
                    Else
                        'ORDR_QTY_BACK = 0
                    End If

                    ' Increment / Decrement other fields *** What about data in SOTPICK2.PICK_QTY_CANC_REL, PICK_QTY_BACK_REL
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + ORDR_QTY_CANC
                    rowSOTORDR2.Item("ORDR_QTY_SHIP") = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                    rowSOTORDR2.Item("ORDR_QTY_OPEN") = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty) - Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty) _
                        - Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) - Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty)

                    If Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty) < 0 Then
                        rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    End If

                    ' Adjust sales order detail status
                    If Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & String.Empty) > 0 Then
                        rowSOTORDR2.Item("ORDR_STATUS") = "O"
                    ElseIf Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty) > 0 Then
                        rowSOTORDR2.Item("ORDR_STATUS") = "P"
                    ElseIf Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) > 0 Then
                        rowSOTORDR2.Item("ORDR_STATUS") = "F"
                    Else
                        rowSOTORDR2.Item("ORDR_STATUS") = "C"
                    End If
                Next

                ' Adjust sales order header's status
                If tblSOTORDR2.Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_STATUS = 'O'").Length > 0 Then
                    rowSOTORDR1.Item("ORDR_STATUS") = "O"
                ElseIf tblSOTORDR2.Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_STATUS = 'P'").Length > 0 Then
                    rowSOTORDR1.Item("ORDR_STATUS") = "P"
                ElseIf tblSOTORDR2.Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_STATUS = 'F'").Length > 0 Then
                    rowSOTORDR1.Item("ORDR_STATUS") = "F"
                Else
                    rowSOTORDR1.Item("ORDR_STATUS") = "C"
                End If

                rowSOTORDR1.Item("ORDR_DATE_CLOSED") = DateTime.Now.Date
                rowSOTORDR1.Item("ORDR_YYYYPP_CLOSED") = ASCMAIN1.CYP

                rowSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTORDR1.Item("LAST_DATE") = DateTime.Now
            Next
        Next
    End Sub

End Class
