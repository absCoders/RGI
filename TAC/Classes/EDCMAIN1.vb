Public Class EDCMAIN1

    Public Shared Sub Generate_810(ByVal clsASCBASE1 As ASCBASE1, _
        ByVal ORDR_INV_NO As String)
        Dim dt As Date = Now + ASCMAIN1.NowTSD
        Dim EDI_OUTBOUND_DOC_NO As String = ""
        Dim EDI_DOC_SEQ_NO As String = ""
        Dim EDI_TP_QUAL As String = ""
        Dim EDI_TP_ID As String = ""
        Dim E_810 As String = ""
        Dim Pack_Code As String = ""
        Dim MOP_Desc As String = ""
        Dim SAH_Allow_Code As String = ""
        Dim SAH_Handling_Code As String = ""
        Dim EDI_PURPOSE_CODE_ORIG As String = ""
        Dim EDI_PURPOSE_CODE_REV As String = ""
        Dim QTY_CASES As Long = 0
        Dim QTY_UNITS As Double = 0
        Dim PACK_FACTOR As Double = 0
        Dim CASE_PRICE As Double = 0
        Dim Edi_Adr_Seq As Double = 0
        Dim Trans_Code As String = 0


        ASCMAIN1.sql = "Select * from SOTINVH1 Where ORDR_INV_NO = :PARM1"
        Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "V", New String() {ORDR_INV_NO})
        If rowSOTINVH1 IsNot Nothing Then
            If rowSOTINVH1.Item("EDI_DOC_SEQ_NO") & "" = "" Then
                E_810 = "Not and EDI Order"
                Exit Sub
            End If
            EDI_DOC_SEQ_NO = rowSOTINVH1.Item("EDI_DOC_SEQ_NO")

            ASCMAIN1.sql = "Select * from EDT850T1 where EDI_DOC_SEQ_NO = :PARM1"
            Dim rowEDT850T1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "V", New String() {EDI_DOC_SEQ_NO})
            If rowEDT850T1 IsNot Nothing Then
                EDI_OUTBOUND_DOC_NO = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                EDI_TP_QUAL = rowEDT850T1.Item("EDI_TP_QUAL")
                EDI_TP_ID = rowEDT850T1.Item("EDI_TP_ID")
            Else
                ASCMAIN1.sql = "Select * from EDT945T1 where EDI_DOC_SEQ_NO = :PARM1"
                rowEDT850T1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "V", New String() {EDI_DOC_SEQ_NO})

                If rowEDT850T1 IsNot Nothing Then
                    EDI_OUTBOUND_DOC_NO = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                    EDI_TP_QUAL = "16"
                    EDI_TP_ID = "1065144410BJS"
                Else
                    'this is not an edi order
                    Exit Sub
                End If
            End If

            ASCMAIN1.sql = "Select * from EDTTRPMZ MZ, EDTTRPM1 M1" _
            & " Where MZ.EDI_TP_QUAL = M1.EDI_TP_QUAL " _
            & " And MZ.EDI_TP_ID = M1.EDI_TP_ID " _
            & " And MZ.EDI_DOC_NO = M1.EDI_DOC_TYPE " _
            & " And EDI_STATUS = 'A' " _
            & " And EDI_DOC_NO = '810' " _
            & " And MZ.EDI_TP_QUAL = :PARM1" _
            & " And MZ.EDI_TP_ID = :PARM2"
            Dim rowEDTTRPMZ As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "VV", New Object() {EDI_TP_QUAL, EDI_TP_ID})
            If rowEDTTRPMZ IsNot Nothing Then
                If EDI_TP_ID <> "1065144410BJS" Then
                    If rowEDTTRPMZ.Item("ALT_EDI_TP_QUAL") & "" <> "" Then
                        EDI_TP_QUAL = rowEDTTRPMZ.Item("ALT_EDI_TP_QUAL")
                        EDI_TP_ID = rowEDTTRPMZ.Item("ALT_EDI_TP_ID")
                    End If
                End If


                EDI_PURPOSE_CODE_ORIG = rowEDTTRPMZ.Item("EDI_PURPOSE_CODE_ORIG") & ""
                EDI_PURPOSE_CODE_REV = rowEDTTRPMZ.Item("EDI_PURPOSE_CODE_REV_810") & ""
                SAH_Allow_Code = rowEDTTRPMZ.Item("SAH_ALLOW_CODE") & ""
                SAH_Handling_Code = rowEDTTRPMZ.Item("SAH_HANDLING_CODE") & ""

                If rowEDTTRPMZ.Item("REPORT_TRANS_CODE") = "N" Then
                    Trans_Code = ""
                Else
                    ASCMAIN1.sql = "Select * from EDTSVIA1 E1, SOTSVIA1 S1" _
                    & " Where EDI_TP_QUAL = '" & EDI_TP_QUAL & "'" _
                    & " And EDI_TP_ID = '" & EDI_TP_ID & "'" _
                    & " And S1.SHIP_CODE = '" & rowSOTINVH1.Item("SHIP_CODE") & "'" _
                    & " And E1.SHIP_METHOD_CODE = S1.SHIP_METHOD_CODE"
                    Dim rowEDTSVIA1 As DataRow = ASCDATA1.GetDataRow

                    If rowEDTSVIA1 IsNot Nothing Then
                        Trans_Code = rowEDTSVIA1.Item("EDI_TRANS_CODE") & ""
                    Else
                        Exit Sub
                    End If
                End If

                Dim rowEDT810O1 As DataRow = clsASCBASE1.dst.Tables("EDT810O1").NewRow  ' TBLs("EDT810O1").NewRow
                With rowEDT810O1
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("CUST_ORDER_NO") = rowSOTINVH1.Item("CUST_ORDER_NO")
                    .Item("ORDR_DATE") = rowSOTINVH1.Item("ORDR_DATE")
                    .Item("CUST_PU_DATE") = rowSOTINVH1.Item("CUST_PU_DATE")
                    .Item("SHIP_VIA") = rowSOTINVH1.Item("SHIP_VIA")
                    .Item("ORDR_AMT") = rowSOTINVH1.Item("ORDR_AMT")
                    .Item("FRT_TERMS") = rowSOTINVH1.Item("FRT_TERMS")
                    .Item("FRT_RATE") = rowSOTINVH1.Item("FRT_RATE")
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = Format(dt, "dd-MMM-yy")
                    .Item("LAST_OPER") = ""
                    .Item("LAST_DATE") = DBNull.Value
                    .Item("ORDR_INV_NO") = rowSOTINVH1.Item("ORDR_INV_NO")
                    .Item("ORDR_INV_DATE") = rowSOTINVH1.Item("ORDR_INV_DATE")
                    .Item("TERM_CODE") = rowSOTINVH1.Item("TERM_CODE")
                    .Item("CUST_ROUTING_INST") = rowSOTINVH1.Item("CUST_ROUTING_INST")
                    '.item("SCAC_CODE") = rowsotinvh1.item("")
                    .Item("ORDR_WD_CHG") = rowSOTINVH1.Item("ORDR_WD_CHG")
                    .Item("ORDR_TOTAL_AMT") = rowSOTINVH1.Item("ORDR_TOTAL_AMT")
                    .Item("CUST_BILL_CASES") = rowSOTINVH1.Item("CUST_BILL_CASES")
                    .Item("CURR_CODE") = rowSOTINVH1.Item("CURR_CODE")
                    .Item("CURR_EXCH_RATE") = rowSOTINVH1.Item("CURR_EXCH_RATE")
                    .Item("EDI_DOC_SEQ_NO") = rowSOTINVH1.Item("EDI_DOC_SEQ_NO")
                    .Item("EDI_ORDR_DOCUMENTKEY") = DBNull.Value
                    .Item("SO_ORDER_NO") = rowSOTINVH1.Item("SO_ORDER_NO")
                    If EDI_TP_ID = "1065144410BJS" Then
                        .Item("EDI_INT_ORDR_NO") = ""
                    Else
                        .Item("EDI_INT_ORDR_NO") = rowEDT850T1.Item("EDI_INT_ORDR_NO")
                    End If
                    .Item("EDI_TRANS_CODE") = Trans_Code

                    ASCMAIN1.sql = "Select * from EDT810O1 Where SO_ORDER_NO = '" & rowSOTINVH1.Item("SO_ORDER_NO") & "'"
                    Dim rowEDT810OC As DataRow = ASCDATA1.GetDataRow
                    If rowEDT810OC IsNot Nothing Then
                        .Item("EDI_PURPOSE_CODE") = EDI_PURPOSE_CODE_REV
                    Else
                        .Item("EDI_PURPOSE_CODE") = EDI_PURPOSE_CODE_ORIG
                    End If

                    ASCMAIN1.sql = "Select Count(*) as LINE_COUNT, sum(QTY_CASES) as CASE_COUNT from SOTINVH2" _
                    & " Where SO_ORDER_NO = '" & rowSOTINVH1.Item("SO_ORDER_NO") & "'"
                    Dim rowSOTCOUNT As DataRow = ASCDATA1.GetDataRow
                    If rowSOTCOUNT IsNot Nothing Then
                        .Item("NO_LINE_ITEMS") = rowSOTCOUNT.Item("LINE_COUNT")
                        .Item("TOT_CASES") = rowSOTCOUNT.Item("CASE_COUNT")
                    End If

                    ASCMAIN1.sql = "Select * from EDTTERM1 " _
                    & " Where TERM_CODE = :PARM1"
                    Dim rowEDTTERM1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "V", New Object() {rowSOTINVH1.Item("TERM_CODE")})

                    If rowEDTTERM1 IsNot Nothing Then
                        .Item("EDI_TERMS") = rowEDTTERM1.Item("EDI_TERMS")
                        .Item("EDI_TERM_TYPE") = rowEDTTERM1.Item("EDI_TERM_TYPE")
                        .Item("EDI_TERM_BASIS") = rowEDTTERM1.Item("EDI_TERM_BASIS")
                        .Item("EDI_TERM_RATE") = rowEDTTERM1.Item("EDI_TERM_RATE")
                        .Item("EDI_TERM_DSCDAYS") = rowEDTTERM1.Item("EDI_TERM_DSCDAYS")
                        .Item("EDI_TERM_NETDAYS") = rowEDTTERM1.Item("EDI_TERM_NETDAYS")
                        .Item("EDI_TERM_DESC") = rowEDTTERM1.Item("EDI_TERM_DESC")

                        If Val(rowEDTTERM1.Item("EDI_TERM_NETDAYS") & "") > 0 Then
                            .Item("EDI_TERM_NET_DUE_DATE") = DateAdd("d", Val(rowEDTTERM1.Item("EDI_TERM_NETDAYS") & ""), rowSOTINVH1.Item("ORDR_INV_DATE"))
                        End If
                    End If
                End With
                clsASCBASE1.dst.Tables("EDT810O1").Rows.Add(rowEDT810O1)

                ASCMAIN1.sql = "Select * from SOTINVH2" _
                & " Where SO_ORDER_NO = '" & rowSOTINVH1.Item("SO_ORDER_NO") & "'"
                For Each rowSOTINVH23 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim Do_Not_Include As Boolean = False

                    If EDI_TP_ID <> "1065144410BJS" Then
                        ASCMAIN1.sql = "Select ICTITEM0.* from ICTITEM0, EDT850T2 " & vbCrLf _
                        & " Where ITEM_CODE = EDI_ITEM" & vbCrLf _
                        & " And EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                        & " And EDI_DTL_SEQ = '" & rowSOTINVH23.Item("EDI_DETL_SEQ") & "'"
                        Dim rowDONOTINCLUDE As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "V", New Object() {rowSOTINVH1.Item("TERM_CODE")})

                        If rowDONOTINCLUDE IsNot Nothing Then
                            If rowDONOTINCLUDE.Item("DO_NOT_REPORT") & "" = "1" Then
                                Do_Not_Include = True
                            End If
                        End If
                    End If

                    If Do_Not_Include = False Then
                        Dim rowEDT810O2 As DataRow = clsASCBASE1.dst.Tables("EDT810O2").NewRow
                        With rowEDT810O2
                            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            .Item("EDI_DETL_SEQ") = rowSOTINVH23.Item("SO_ORDER_LNO")
                            .Item("PROD_CODE") = rowSOTINVH23.Item("PROD_CODE")
                            .Item("SIZE_CODE") = rowSOTINVH23.Item("SIZE_CODE")
                            .Item("ORIG_CODE") = rowSOTINVH23.Item("ORIG_CODE")
                            .Item("WHSE_CODE") = rowSOTINVH23.Item("WHSE_CODE")
                            .Item("PACK_CODE") = rowSOTINVH23.Item("PACK_CODE")
                            .Item("GRADE_CODE") = rowSOTINVH23.Item("GRADE_CODE")
                            .Item("BRAND_CODE") = rowSOTINVH23.Item("BRAND_CODE")
                            .Item("QTY_CASES") = rowSOTINVH23.Item("QTY_CASES")
                            .Item("QTY_UNITS") = rowSOTINVH23.Item("QTY_UNITS")
                            .Item("GTIN") = rowSOTINVH23.Item("GTIN")

                            If EDI_TP_ID <> "1065144410BJS" Then
                                ASCMAIN1.sql = "Select * from EDT850T2" _
                                & " Where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
                                & " And EDI_DTL_SEQ = " & rowSOTINVH23.Item("EDI_DETL_SEQ")
                                Dim rowEDT850T2 As DataRow = ASCDATA1.GetDataRow

                                If rowEDT850T2 IsNot Nothing Then
                                    ASCMAIN1.sql = " Select * from ICTITEM0, SOTCITM1" _
                                    & " Where ICTITEM0.ITEM_CODE = SOTCITM1.ITEM_CODE" _
                                    & " And CUST_SKU = '" & rowEDT850T2.Item("EDI_SKU") & "'"
                                    Dim rowICTITEM0 As DataRow = ASCDATA1.GetDataRow

                                    If rowICTITEM0 IsNot Nothing Then
                                        .Item("EMP_MFG_ID") = rowICTITEM0.Item("ITEM_CODE") & ""
                                    Else
                                        .Item("EMP_MFG_ID") = ""
                                    End If
                                    If rowEDT850T2.Item("ENTRY_TYPE") & "" <> "M" Then
                                        .Item("EDI_EMP_ITEM_CODE") = rowEDT850T2.Item("EDI_ITEM") & ""
                                    Else
                                        .Item("EDI_EMP_ITEM_CODE") = ""
                                    End If
                                    .Item("EDI_CUST_ITEM_CODE") = rowEDT850T2.Item("EDI_SKU")
                                    If rowEDT850T2.Item("EDI_ITEM_DESC") & "" <> "" Then
                                        .Item("LINE_ITEM_DESCR") = rowEDT850T2.Item("EDI_ITEM_DESC")
                                    Else
                                        ASCMAIN1.sql = " Select PROD_DESC, SIZE_DESC " _
                                        & " from EDT850T2 T2, SOTCITM1 M1, ICTITEM0 M0," _
                                        & " ICTPROD1 D1, ICTSIZE1 E1" _
                                        & " where edi_doc_seq_no =  '" & EDI_DOC_SEQ_NO & "'" _
                                        & " And EDI_DTL_SEQ = " & rowSOTINVH23.Item("EDI_DETL_SEQ") _
                                        & " And T2.EDI_SKU = M1.CUST_SKU" _
                                        & " And M1.ITEM_CODE = M0.ITEM_CODE" _
                                        & " And D1.PROD_CODE = M0.PROD_CODE" _
                                        & " And E1.SIZE_CODE = M0.SIZE_CODE"
                                        Dim rowITEMDESCR As DataRow = ASCDATA1.GetDataRow

                                        If rowITEMDESCR IsNot Nothing Then
                                            .Item("LINE_ITEM_DESCR") = rowITEMDESCR.Item("PROD_DESC") & " " & rowITEMDESCR.Item("SIZE_DESC")
                                        End If
                                    End If

                                    .Item("EDI_PO1_UOM") = rowEDT850T2.Item("EDI_PO1_UOM")
                                    If .Item("EDI_PO1_UOM") = "EA" Then .Item("EDI_PO1_UOM") = "CA" 'added this line b/c I was getting rejected 810's if it was in EA
                                    .Item("EDI_PO4_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
                                    If rowEDT850T2.Item("EDI_PO4_UOM") & "" = "PP" Then
                                        .Item("ORDR_PRICE_GRS_CURR") = rowSOTINVH23.Item("ORDR_PRICE_GRS_CURR")
                                    Else
                                        QTY_CASES = Val(rowSOTINVH23.Item("QTY_CASES") & "")
                                        QTY_UNITS = Val(rowSOTINVH23.Item("QTY_UNITS") & "")
                                        PACK_FACTOR = Val(rowSOTINVH23.Item("QTY_UNITS") & "") / Val(rowSOTINVH23.Item("QTY_CASES") & "")

                                        If rowSOTINVH23.Item("PACK_CODE_INV") & "" <> "" Then
                                            Dim rowICTPACK1 As DataRow = clsASCBASE1.LookUp("ICTPACK1", rowSOTINVH23.Item("PACK_CODE_INV"))

                                            PACK_FACTOR = rowICTPACK1.Item("PACK_FACTOR")
                                            QTY_CASES = QTY_UNITS / PACK_FACTOR
                                        End If
                                        If rowSOTINVH23.Item("EDI_DETL_SEQ") = 12 And ASCMAIN1.Running_in_VS Then Stop
                                        CASE_PRICE = Val(rowSOTINVH23.Item("ORDR_PRICE_GRS") & "") * PACK_FACTOR
                                        'CASE_PRICE = Val(6.9) * PACK_FACTOR
                                        .Item("QTY_CASES") = QTY_CASES

                                        .Item("ORDR_PRICE_GRS_CURR") = CASE_PRICE
                                    End If
                                Else

                                    .Item("LINE_ITEM_DESCR") = rowSOTINVH23.Item("LINE_ITEM_DESCR") & ""
                                    .Item("ORDR_PRICE_GRS_CURR") = (Val(rowSOTINVH23.Item("QTY_UNITS") & "") / Val(rowSOTINVH23.Item("QTY_CASES") & "")) * rowSOTINVH23.Item("ORDR_PRICE_GRS_CURR")
                                End If
                            Else
                                'bj Code here

                                ASCMAIN1.sql = " Select ICTITEM0.*, EDT945T2.CUST_SKU " & vbCrLf _
                                & " from ICTITEM0, EDT945T2, SOTCITM1 " & vbCrLf _
                                & " Where ICTITEM0.ITEM_CODE = SOTCITM1.ITEM_CODE" & vbCrLf _
                                & " And SOTCITM1.CUST_SKU = SUBSTR(EDT945T2.CUST_SKU,8,5) " & vbCrLf _
                                & " And EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                                & " And EDI_DTL_SEQ = '" & rowSOTINVH23.Item("EDI_DETL_SEQ") & "'"
                                Dim rowICTITEM0 As DataRow = ASCDATA1.GetDataRow

                                If rowICTITEM0 IsNot Nothing Then
                                    .Item("EMP_MFG_ID") = rowICTITEM0.Item("ITEM_CODE")
                                Else
                                    .Item("EMP_MFG_ID") = ""
                                End If

                                .Item("EDI_CUST_ITEM_CODE") = rowICTITEM0.Item("CUST_SKU")
                                .Item("EDI_PO1_UOM") = "CA"
                                .Item("EDI_PO4_UOM") = "CA"

                                Dim rowICTSIZE1 As DataRow = clsASCBASE1.LookUp("ICTSIZE1", rowICTITEM0.Item("SIZE_CODE"))
                                Dim rowICTPROD1 As DataRow = clsASCBASE1.LookUp("ICTPROD1", rowICTITEM0.Item("PROD_CODE"))
                                .Item("LINE_ITEM_DESCR") = rowICTPROD1.Item("PROD_DESC") & " " & rowICTSIZE1.Item("SIZE_DESC")

                                QTY_CASES = Val(rowSOTINVH23.Item("QTY_CASES") & "")
                                QTY_UNITS = Val(rowSOTINVH23.Item("QTY_UNITS") & "")
                                PACK_FACTOR = Val(rowSOTINVH23.Item("QTY_UNITS") & "") / Val(rowSOTINVH23.Item("QTY_CASES") & "")

                                If rowSOTINVH23.Item("PACK_CODE_INV") & "" <> "" Then
                                    Dim rowICTPACK1 As DataRow = clsASCBASE1.LookUp("ICTPACK1", rowSOTINVH23.Item("PACK_CODE_INV"))

                                    PACK_FACTOR = rowICTPACK1.Item("PACK_FACTOR")
                                    QTY_CASES = QTY_UNITS / PACK_FACTOR
                                End If
                                CASE_PRICE = Val(rowSOTINVH23.Item("ORDR_PRICE_GRS") & "") * PACK_FACTOR

                                .Item("QTY_CASES") = QTY_CASES
                                .Item("ORDR_PRICE_GRS_CURR") = CASE_PRICE
                            End If

                        End With

                        clsASCBASE1.dst.Tables("EDT810O2").Rows.Add(rowEDT810O2)

                        '----Populate EDT810O3
                        ASCMAIN1.sql = "Select * from SOTINVH3" _
                        & " Where SO_ORDER_NO = '" & rowSOTINVH23.Item("SO_ORDER_NO") & "'" _
                        & " And SO_ORDER_LNO = '" & rowSOTINVH23.Item("SO_ORDER_LNO") & "'"
                        For Each rowSOTINVH3 As DataRow In ASCDATA1.GetDataTable.Rows
                            Dim rowEDT810O3 As DataRow = clsASCBASE1.dst.Tables("EDT810O3").NewRow
                            With rowEDT810O3
                                rowEDT810O3.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                                rowEDT810O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                                rowEDT810O3.Item("EDI_DETL_SEQ") = rowSOTINVH3.Item("SO_ORDER_LNO")
                                rowEDT810O3.Item("EDI_LOT_LNO") = rowSOTINVH3.Item("SO_LOT_LNO")
                                rowEDT810O3.Item("WHSE_CODE") = rowSOTINVH3.Item("WHSE_CODE")
                                rowEDT810O3.Item("LOT_NO") = rowSOTINVH3.Item("LOT_NO")
                                rowEDT810O3.Item("LOT_SEQ_NO") = rowSOTINVH3.Item("LOT_SEQ_NO")
                                rowEDT810O3.Item("LOT_ORDER_QTY") = rowSOTINVH3.Item("LOT_ORDER_QTY")
                                rowEDT810O3.Item("PACK_CODE") = rowSOTINVH3.Item("PACK_CODE")
                                rowEDT810O3.Item("PACK_FACTOR") = rowSOTINVH3.Item("PACK_FACTOR")
                                rowEDT810O3.Item("SO_LOT_CASES") = rowSOTINVH3.Item("SO_LOT_CASES")
                                rowEDT810O3.Item("SO_LOT_UNITS") = rowSOTINVH3.Item("SO_LOT_UNITS")
                                rowEDT810O3.Item("COOL_COMPLIANT") = rowSOTINVH3.Item("COOL_COMPLIANT")
                                rowEDT810O3.Item("MOP") = rowSOTINVH3.Item("MOP")
                                Select Case rowSOTINVH3.Item("MOP") & ""
                                    Case "F"
                                        MOP_Desc = "MOP:Farm Raised;"
                                    Case "W"
                                        MOP_Desc = "MOP:Wild Caught;"
                                    Case "N"
                                        MOP_Desc = "MOP:Not Available;"
                                    Case Else
                                        MOP_Desc = "MOP:Not Available;"
                                End Select

                                ASCMAIN1.sql = "Select ORIG_DESC from ICTLOTD1 D1, ICTORIG1 G1" _
                                & " Where LOT_NO = '" & rowSOTINVH3.Item("LOT_NO") & "'" _
                                & " And LOT_SEQ_NO = '" & rowSOTINVH3.Item("LOT_SEQ_NO") & "'" _
                                & " And WHSE_CODE = '" & rowSOTINVH3.Item("WHSE_CODE") & "'" _
                                & " And D1.ORIG_CODE = G1.ORIG_CODE"
                                Dim rowICTLOTD1 As DataRow = ASCDATA1.GetDataRow
                                If rowICTLOTD1 IsNot Nothing Then
                                    MOP_Desc = MOP_Desc & " Country of Origin:" & rowICTLOTD1.Item("ORIG_DESC") & ""
                                Else
                                    MOP_Desc = MOP_Desc & " Country of Origin:"
                                End If
                                rowEDT810O3.Item("LOT_NOTES") = MOP_Desc
                            End With
                            clsASCBASE1.dst.Tables("EDT810O3").Rows.Add(rowEDT810O3)
                        Next
                    End If
                Next

                '-----Fill EDT810O5
                If EDI_TP_ID <> "1065144410BJS" Then
                    ASCMAIN1.sql = " Insert into EDT810O5" _
                    & " Select '" & ASCMAIN1.DBS_COMPANY & "','" & EDI_OUTBOUND_DOC_NO & "'," _
                    & " EDI_ADR_SEQ, EDI_ADDR_TYPE, EDI_CUST_NAME_ADR, EDI_ADDRESS1," _
                    & " EDI_ADDRESS2, EDI_CITY, EDI_STATE, EDI_ZIPCODE, EDI_COUNTRY, EDI_ADDR_CODE," _
                    & " EDI_ADDR_CODE_QUAL, '' as EDI_HL2_SEQ, EDI_ADDRESS3, EDI_ADDRESS4" _
                    & " from EDT850T5 " _
                    & " where EDI_DOC_SEQ_NO = '" & rowEDT850T1.Item("EDI_DOC_SEQ_NO") & "'" _
                    & " And EDI_ADDR_TYPE <> 'SF' "
                    ASCDATA1.ExecuteSQL()
                Else

                    ASCMAIN1.sql = "  Insert into EDT810O5 Select '" & ASCMAIN1.DBS_COMPANY & "', '" & EDI_OUTBOUND_DOC_NO & "'," & vbCrLf _
                    & " '1' as EDI_ADR_SEQ, 'ST' as EDI_ADDR_TYPE," & vbCrLf _
                    & "  H1.CUST_SHIP_TO_NAME, H1.CUST_SHIP_TO_ADDR1, H1.CUST_SHIP_TO_ADDR2," & vbCrLf _
                    & "  H1.CUST_SHIP_TO_CITY, H1.CUST_SHIP_TO_STATE, H1.CUST_SHIP_TO_ZIP_CODE, " & vbCrLf _
                    & "  H1.CUST_SHIP_TO_COUNTRY, CUST_DUNS_NO as EDI_ADDR_CODE, " & vbCrLf _
                    & "  '9' as  EDI_ADDR_CODE_QUAL, '' as EDI_HL2_SEQ, " & vbCrLf _
                    & "  '' as EDI_ADDRESS3, '' as EDI_ADDRESS4 from SOTINVH1 H1, EDT945T1 T1" & vbCrLf _
                    & "  Where H1.EDI_DOC_SEQ_NO = T1.EDI_DOC_SEQ_NO" & vbCrLf _
                    & "  And H1.EDI_DOC_SEQ_NO = '" & rowEDT850T1.Item("EDI_DOC_SEQ_NO") & "'"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "  Insert into EDT810O5 Select '" & ASCMAIN1.DBS_COMPANY & "',  '" & EDI_OUTBOUND_DOC_NO & "'," & vbCrLf _
                    & " '2' as EDI_ADR_SEQ, 'BT' as EDI_ADDR_TYPE," & vbCrLf _
                    & "  CUST_NAME, '' as CUST_SHIP_TO_ADDR1, '' as CUST_SHIP_TO_ADDR2," & vbCrLf _
                    & "  '' as CUST_SHIP_TO_CITY, '' as CUST_SHIP_TO_STATE, '' as CUST_SHIP_TO_ZIP_CODE, " & vbCrLf _
                    & "  '' as CUST_SHIP_TO_COUNTRY, '' as EDI_ADDR_CODE," & vbCrLf _
                    & "  '' as  EDI_ADDR_CODE_QUAL, '' as EDI_HL2_SEQ, " & vbCrLf _
                    & "  '' as EDI_ADDRESS3, '' as EDI_ADDRESS4 from ARTCUST1 T1" & vbCrLf _
                    & "  Where CUST_CODE = 'E00655'"
                    ASCDATA1.ExecuteSQL()
                End If



                ASCMAIN1.sql = "Select Max(EDI_ADR_SEQ) as MAX_SEQ From EDT810O5" _
                & " Where COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" _
                & " And EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'"
                Dim rowEDT810OA As DataRow = ASCDATA1.GetDataRow

                If rowEDT810OA IsNot Nothing Then
                    Edi_Adr_Seq = rowEDT810OA.Item("MAX_SEQ") + 1
                Else
                    Edi_Adr_Seq = 1
                End If

                Dim rowEDT810O5 As DataRow = clsASCBASE1.dst.Tables("EDT810O5").NewRow
                With rowEDT810O5
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("EDI_ADR_SEQ") = Edi_Adr_Seq
                    .Item("EDI_ADDR_TYPE") = "RE"
                    .Item("EDI_CUST_NAME_ADR") = rowEDTTRPMZ.Item("EDI_CUST_NAME_ADR") & ""
                    .Item("EDI_ADDRESS1") = rowEDTTRPMZ.Item("EDI_ADDRESS1") & ""
                    .Item("EDI_ADDRESS2") = rowEDTTRPMZ.Item("EDI_ADDRESS2") & ""
                    .Item("EDI_CITY") = rowEDTTRPMZ.Item("EDI_CITY") & ""
                    .Item("EDI_STATE") = rowEDTTRPMZ.Item("EDI_STATE") & ""
                    .Item("EDI_ZIPCODE") = rowEDTTRPMZ.Item("EDI_ZIPCODE") & ""
                    .Item("EDI_COUNTRY") = ""
                    .Item("EDI_ADDR_CODE") = "4426297228"
                    .Item("EDI_ADDR_CODE_QUAL") = "91"
                    .Item("EDI_HL2_SEQ") = DBNull.Value
                    .Item("EDI_ADDRESS3") = ""
                    .Item("EDI_ADDRESS4") = ""
                End With
                clsASCBASE1.dst.Tables("EDT810O5").Rows.Add(rowEDT810O5)


                If rowSOTINVH1.Item("ORDR_WD_CHG") <> 0 Then
                    Dim SAH_SEQ_NO As Double = 0

                    ASCMAIN1.sql = "Select Max(SAH_SEQ_NO) as MAX_SEQ From EDT810O7" _
                    & " Where COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" _
                    & " And EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'"
                    Dim rowMAXSAH As DataRow = ASCDATA1.GetDataRow

                    If rowMAXSAH IsNot Nothing Then
                        SAH_SEQ_NO = Val(rowMAXSAH.Item("MAX_SEQ") & "") + 1
                    Else
                        SAH_SEQ_NO = 1
                    End If
                    Dim rowEDT810O7 As DataRow = clsASCBASE1.dst.Tables("EDT810O7").NewRow
                    With rowEDT810O7
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        .Item("SAH_SEQ_NO") = SAH_SEQ_NO
                        .Item("INV_NO") = rowSOTINVH1.Item("ORDR_INV_NO")
                        .Item("SAH_ALLOW_IND") = "C"
                        .Item("SAH_ALLOW_CODE") = SAH_Allow_Code
                        .Item("SAH_AMOUNT") = rowSOTINVH1.Item("ORDR_WD_CHG")
                        .Item("SAH_PERCENT_QUAL") = ""
                        .Item("SAH_PERCENT") = DBNull.Value
                        .Item("SAH_RATE") = DBNull.Value
                        .Item("SAH_UOM_CODE") = ""
                        .Item("SAH_QTY") = DBNull.Value
                        .Item("SAH_HANDLING_CODE") = SAH_Handling_Code
                        .Item("SAH_DESC") = "Withdrawal Charge"
                    End With
                    clsASCBASE1.dst.Tables("EDT810O7").Rows.Add(rowEDT810O7)
                End If

                Call clsASCBASE1.Update_Record_TDA("EDT810O1")
                Call clsASCBASE1.Update_Record_TDA("EDT810O2")
                Call clsASCBASE1.Update_Record_TDA("EDT810O3")
                Call clsASCBASE1.Update_Record_TDA("EDT810O5")
                Call clsASCBASE1.Update_Record_TDA("EDT810O7")
                Call clsASCBASE1.Update_Record_TDA("EDT810O8")

                ASCMAIN1.sql = " Insert into EDTSYSIH values (" _
                & "'" & ASCMAIN1.DBS_COMPANY & "'," _
                & "'" & EDI_OUTBOUND_DOC_NO _
                & "','OIN','" & TAC.TACMAIN1.EDI_PROCESS_IND & "','" _
                & EDI_TP_ID & "')"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTINVH1 Set ORDR_810_IND = '1'" _
                & " Where SO_ORDER_NO = '" & rowSOTINVH1.Item("SO_ORDER_NO") & "'"
                ASCDATA1.ExecuteSQL()
            End If

        End If
    End Sub

    Public Shared Sub Generate_855(ByVal clsASCBASE1 As ASCBASE1, _
                                   ByVal SO_ORDER_NO As String, _
                                   ByVal TABLE_NAME As String)

        Dim EDI_OUR_ID As String = ""
        Dim EDI_DOC_SEQ_NO As String = ""
        Dim EDI_PURP_CODE As String = ""
        Dim Err_Msg As String = ""
        Dim Report_Manual_Lines As String = ""
        Dim dt As Date = Now + ASCMAIN1.NowTSD


        ASCMAIN1.sql = "Select * from " & TABLE_NAME & "1 R1, EDT850T1 T1" _
        & " Where SO_ORDER_NO = '" & SO_ORDER_NO & "'" _
        & " And R1.EDI_DOC_SEQ_NO is Not Null" _
        & " And R1.EDI_DOC_SEQ_NO = T1.EDI_DOC_SEQ_NO"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow
        If rowSOTORDR1 IsNot Nothing Then
            EDI_DOC_SEQ_NO = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
            Dim EDI_Outbound_Doc_No As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
            EDI_OUR_ID = Replace(rowSOTORDR1.Item("EDI_OUR_ID"), " ", "")

            ASCMAIN1.sql = "Select * from EDTTRPMZ MZ, EDTTRPM1 M1" & vbCrLf _
            & " Where MZ.EDI_STATUS = 'A'" & vbCrLf _
            & " And MZ.EDI_DOC_NO = '855'" & vbCrLf _
            & " And MZ.EDI_TP_QUAL = '" & rowSOTORDR1.Item("EDI_TP_QUAL") & "'" & vbCrLf _
            & " And MZ.EDI_TP_ID = rtrim('" & rowSOTORDR1.Item("EDI_TP_ID") & "')" & vbCrLf _
            & " And MZ.EDI_TP_QUAL = M1.EDI_TP_QUAL" & vbCrLf _
            & " And MZ.EDI_TP_ID = M1.EDI_TP_ID" & vbCrLf _
            & " And MZ.EDI_DOC_NO = M1.EDI_DOC_TYPE"
            Dim rowEDTTRPMZ As DataRow = ASCDATA1.GetDataRow

            If rowEDTTRPMZ IsNot Nothing Then
                Dim EDI_TP_ID As String = rowEDTTRPMZ.Item("EDI_TP_ID")

                ASCMAIN1.sql = "Select IH.* from EDT855O1 O1, EDTSYSIH IH" & vbCrLf _
                & " Where SO_ORDER_NO = '" & SO_ORDER_NO & "'" & vbCrLf _
                & " And O1.COMPANY_CODE = IH.COMPANY_CODE" & vbCrLf _
                & " And O1.EDI_OUTBOUND_DOC_NO = IH.EDI_OUTBOUND_DOC_NO"
                Dim rowEDT855OC As DataRow = ASCDATA1.GetDataRow

                If rowEDT855OC IsNot Nothing Then
                    EDI_PURP_CODE = rowEDTTRPMZ.Item("EDI_PURPOSE_CODE_REV") & ""
                Else
                    EDI_PURP_CODE = rowEDTTRPMZ.Item("EDI_PURPOSE_CODE_ORIG") & ""
                End If
                Report_Manual_Lines = rowEDTTRPMZ.Item("REPORT_MANUAL_LINES") & ""

                Dim rowEDT855O1 As DataRow = clsASCBASE1.dst.Tables("EDT855O1").NewRow  ' TBLs("EDT810O1").NewRow
                With rowEDT855O1
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                    .Item("EDI_PURPOSE_CODE") = EDI_PURP_CODE
                    .Item("EDI_PO_NO") = rowSOTORDR1.Item("EDI_PO_NO")
                    .Item("EDI_PO_DATE") = rowSOTORDR1.Item("EDI_PO_DATE")
                    .Item("EDI_CONFIRM_DATE") = dt
                    .Item("EDI_SHIP_DATE") = rowSOTORDR1.Item("CUST_PU_DATE")
                    .Item("EDI_ARRIVAL_DATE") = rowSOTORDR1.Item("ORDR_ETA_DATE")


                    ASCMAIN1.sql = "Select EDI_TRANS_CODE from EDTSVIA1, SOTSVIA1" & vbCrLf _
                    & " where EDI_TP_QUAL = '" & rowSOTORDR1.Item("EDI_TP_QUAL") & "'" & vbCrLf _
                    & " AND EDI_TP_ID = '" & EDI_TP_ID & "'" & vbCrLf _
                    & " AND SHIP_CODE = '" & rowSOTORDR1.Item("SHIP_CODE") & "'" & vbCrLf _
                    & " And SOTSVIA1.SHIP_METHOD_CODE = EDTSVIA1.SHIP_METHOD_CODE"
                    Dim rowEDTSVIA1 As DataRow = ASCDATA1.GetDataRow

                    If rowEDTSVIA1 IsNot Nothing Then
                        .Item("TRANSPORTATION_METHOD") = rowEDTSVIA1.Item("EDI_TRANS_CODE")
                    Else
                        ASCMAIN1.sql = "Select * from EDTXREF5" & vbCrLf _
                        & " where SENDER_ID_QUAL = '" & rowSOTORDR1.Item("EDI_TP_QUAL") & "'" & vbCrLf _
                        & " AND SENDER_ID = '" & EDI_TP_ID & "'" & vbCrLf _
                        & " AND FRT_TERMS = '" & rowSOTORDR1.Item("FRT_TERMS") & "'"
                        Dim rowEDTXREF5 As DataRow = ASCDATA1.GetDataRow

                        If rowEDTXREF5 IsNot Nothing Then
                            .Item("TRANSPORTATION_METHOD") = rowEDTXREF5.Item("TRANSPORTATION_METHOD") & ""
                        Else
                            .Item("TRANSPORTATION_METHOD") = "SR"
                        End If
                    End If

                    ' this should be codified and translated from SR-> PPD and H-> CPU instead of SUPP TRUCK and CUST PU
                    .Item("EDI_SHIP_ADDR_TYPE") = rowSOTORDR1.Item("EDI_SHIP_ADDR_TYPE")
                    If rowSOTORDR1.Item("EDI_SHIP_ADDR_TYPE") & "" = "MK" Then
                        .Item("EDI_STORE") = rowSOTORDR1.Item("CUST_SHIP_TO_CODE")
                        .Item("EDI_SHIP_DC") = rowSOTORDR1.Item("EDI_SHIP_DC")
                    Else
                        .Item("EDI_STORE") = rowSOTORDR1.Item("EDI_STORE")
                        .Item("EDI_SHIP_DC") = rowSOTORDR1.Item("CUST_SHIP_TO_CODE")
                    End If
                    .Item("STORE_GLOBAL_LOCATION_NUMBER") = rowSOTORDR1.Item("STORE_GLOBAL_LOCATION_NUMBER")
                    .Item("DC_GLOBAL_LOCATION_NUMBER") = rowSOTORDR1.Item("DC_GLOBAL_LOCATION_NUMBER")
                    .Item("EDI_DEPARTMENT") = rowSOTORDR1.Item("EDI_DEPARTMENT")
                    .Item("EDI_SUPPLIER_NO") = rowSOTORDR1.Item("EDI_SUPPLIER_NO")
                    .Item("EDI_PROMOTION") = rowSOTORDR1.Item("EDI_PROMOTION")
                    .Item("EDI_MERCH_TYPE") = rowSOTORDR1.Item("EDI_MERCH_TYPE")
                    If rowEDTTRPMZ.Item("EDI_ACK_TYPE") <> "" Then
                        .Item("EDI_ACK_TYPE") = rowEDTTRPMZ.Item("EDI_ACK_TYPE")
                    Else
                        .Item("EDI_ACK_TYPE") = "AC"
                    End If
                    .Item("EDI_DOC_SEQ_NO") = rowSOTORDR1.Item("EDI_DOC_SEQ_NO")
                    .Item("SO_ORDER_NO") = SO_ORDER_NO
                    .Item("EDI_INT_ORDR_NO") = rowSOTORDR1.Item("EDI_INT_ORDR_NO")
                End With
                clsASCBASE1.dst.Tables("EDT855O1").Rows.Add(rowEDT855O1)


                ASCMAIN1.sql = "Select * from EDT850T2" & vbCrLf _
                & " Where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                For Each rowEDT850T2 As DataRow In ASCDATA1.GetDataTable.Rows
                    If rowEDT850T2.Item("ENTRY_TYPE") & "" = "M" Then
                        If Report_Manual_Lines <> "N" Then
                            Dim rowICTITEM0 As DataRow = clsASCBASE1.LookUp("ICTITEM0", rowEDT850T2.Item("EDI_ITEM") & "")
                            If rowICTITEM0 IsNot Nothing Then
                                If rowICTITEM0.Item("DO_NOT_REPORT") & "" <> "1" Then
                                    Add_EDT855O2(clsASCBASE1, rowEDT850T2, SO_ORDER_NO, EDI_Outbound_Doc_No, TABLE_NAME, EDI_TP_ID)
                                End If
                            Else
                                Add_EDT855O2(clsASCBASE1, rowEDT850T2, SO_ORDER_NO, EDI_Outbound_Doc_No, TABLE_NAME, EDI_TP_ID)
                            End If
                        End If

                    Else
                        Dim rowICTITEM0 As DataRow = clsASCBASE1.LookUp("ICTITEM0", rowEDT850T2.Item("EDI_ITEM") & "")
                        If rowICTITEM0 IsNot Nothing Then
                            If rowICTITEM0.Item("DO_NOT_REPORT") & "" <> "1" Then
                                Add_EDT855O2(clsASCBASE1, rowEDT850T2, SO_ORDER_NO, EDI_Outbound_Doc_No, TABLE_NAME, EDI_TP_ID)
                            End If
                        Else
                            Add_EDT855O2(clsASCBASE1, rowEDT850T2, SO_ORDER_NO, EDI_Outbound_Doc_No, TABLE_NAME, EDI_TP_ID)
                        End If
                    End If

                Next

                ASCMAIN1.sql = "Select * from EDT850T5 " & vbCrLf _
                & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                & " And EDI_ADDR_TYPE <> 'SF' "
                For Each rowEDT850T5 As DataRow In ASCDATA1.GetDataTable.Rows

                    Dim rowEDT855O5 As DataRow = clsASCBASE1.dst.Tables("EDT855O5").NewRow  ' TBLs("EDT810O1").NewRow
                    With rowEDT855O5
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                        .Item("EDI_ADR_SEQ") = rowEDT850T5.Item("EDI_ADR_SEQ") & ""
                        .Item("EDI_DOC_NO") = rowEDT850T5.Item("EDI_DOC_NO") & ""
                        .Item("EDI_ADDR_TYPE") = rowEDT850T5.Item("EDI_ADDR_TYPE") & ""
                        .Item("EDI_CUST_NAME_ADR") = rowEDT850T5.Item("EDI_CUST_NAME_ADR") & ""
                        .Item("EDI_ADDRESS1") = rowEDT850T5.Item("EDI_ADDRESS1") & ""
                        .Item("EDI_ADDRESS2") = rowEDT850T5.Item("EDI_ADDRESS2") & ""
                        .Item("EDI_CITY") = rowEDT850T5.Item("EDI_CITY") & ""
                        .Item("EDI_STATE") = rowEDT850T5.Item("EDI_STATE") & ""
                        .Item("EDI_ZIPCODE") = rowEDT850T5.Item("EDI_ZIPCODE") & ""
                        .Item("EDI_COUNTRY") = rowEDT850T5.Item("EDI_COUNTRY") & ""
                        .Item("EDI_ADDR_CODE") = rowEDT850T5.Item("EDI_ADDR_CODE") & ""
                        .Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & ""
                        .Item("EDI_ADDRESS3") = rowEDT850T5.Item("EDI_ADDRESS3") & ""
                        .Item("EDI_ADDRESS4") = rowEDT850T5.Item("EDI_ADDRESS4") & ""
                    End With
                    clsASCBASE1.dst.Tables("EDT855O5").Rows.Add(rowEDT855O5)
                Next


                ASCMAIN1.sql = "Select * from EDT850T7 " & vbCrLf _
                & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                For Each rowEDT850T7 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim rowEDT855O7 As DataRow = clsASCBASE1.dst.Tables("EDT855O7").NewRow  ' TBLs("EDT810O1").NewRow
                    With rowEDT855O7
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                        .Item("SAH_SEQ_NO") = rowEDT850T7.Item("SAH_SEQ_NO") & ""
                        .Item("SAH_ALLOW_IND") = rowEDT850T7.Item("SAH_ALLOW_IND") & ""
                        .Item("SAH_ALLOW_CODE") = rowEDT850T7.Item("SAH_ALLOW_CODE") & ""
                        .Item("SAH_AMOUNT") = rowEDT850T7.Item("SAH_AMOUNT") & ""
                        .Item("SAH_PERCENT_QUAL") = rowEDT850T7.Item("SAH_PERCENT_QUAL") & ""
                        .Item("SAH_PERCENT") = rowEDT850T7.Item("SAH_PERCENT")
                        .Item("SAH_RATE") = rowEDT850T7.Item("SAH_RATE")
                        .Item("SAH_UOM_CODE") = rowEDT850T7.Item("SAH_UOM_CODE") & ""
                        .Item("SAH_QTY") = rowEDT850T7.Item("SAH_QTY")
                        .Item("SAH_HANDLING_CODE") = rowEDT850T7.Item("SAH_HANDLING_CODE") & ""
                        .Item("SAH_DESC") = rowEDT850T7.Item("SAH_DESC") & ""
                    End With
                    clsASCBASE1.dst.Tables("EDT855O7").Rows.Add(rowEDT855O7)
                Next

                Call clsASCBASE1.Update_Record_TDA("EDT855O1")
                Call clsASCBASE1.Update_Record_TDA("EDT855O2")
                Call clsASCBASE1.Update_Record_TDA("EDT855O5")
                Call clsASCBASE1.Update_Record_TDA("EDT855O7")


                ASCMAIN1.sql = " Insert into EDTSYSIH values (" _
                & "'" & ASCMAIN1.DBS_COMPANY & "'," _
                & "'" & EDI_Outbound_Doc_No _
                & "','OPR','" & TAC.TACMAIN1.EDI_PROCESS_IND & "','" _
                & EDI_TP_ID & "')"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update " & TABLE_NAME & "1 " _
                & " Set EDI_OUTBOUND_DOC_NO = '" & EDI_Outbound_Doc_No & "'" _
                & " WHERE SO_ORDER_NO = '" & SO_ORDER_NO & "'"
                ASCDATA1.ExecuteSQL()
            End If


        Else
            Err_Msg = Err_Msg & "Could Not Link Sales Order to 850"
        End If
    End Sub

    Public Shared Sub Add_EDT855O2(ByVal clsASCBASE1 As ASCBASE1, _
                     ByVal row850T2 As DataRow, _
                     ByVal SO_ORDER_NO As String, _
                     ByVal EDI_Outbound_Doc_No As String, _
                     ByVal Table_Name As String, _
                     ByVal EDI_TP_ID As String)

        Dim QTY_UNITS As Double = 0
        Dim QTY_CASES As Long = 0
        Dim CASE_PRICE As Double = 0
        Dim PACK_FACTOR As Double = 0

        Dim rowEDT855O2 As DataRow = clsASCBASE1.dst.Tables("EDT855O2").NewRow  ' TBLs("EDT810O1").NewRow
        With rowEDT855O2
            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
            .Item("EDI_DTL_SEQ") = row850T2.Item("EDI_DTL_SEQ")

            ASCMAIN1.sql = "Select * from " & Table_Name & "2" & vbCrLf _
            & " Where SO_ORDER_NO = '" & SO_ORDER_NO & "'" & vbCrLf _
            & " And EDI_DETL_SEQ = " & row850T2.Item("EDI_DTL_SEQ")
            Dim rowSOTORDR2 As DataRow = ASCDATA1.GetDataRow

            If rowSOTORDR2 IsNot Nothing Then
                QTY_CASES = Val(rowSOTORDR2.Item("QTY_CASES") & "")
                QTY_UNITS = Val(rowSOTORDR2.Item("QTY_UNITS") & "")
                PACK_FACTOR = Val(rowSOTORDR2.Item("QTY_UNITS") & "") / Val(rowSOTORDR2.Item("QTY_CASES") & "")
                If rowSOTORDR2.Item("PACK_CODE_INV") & "" <> "" Then
                    Dim rowICTPACK1 As DataRow = clsASCBASE1.LookUp("ICTPACK1", rowSOTORDR2.Item("PACK_CODE_INV"))

                    PACK_FACTOR = rowICTPACK1.Item("PACK_FACTOR")
                    QTY_CASES = QTY_UNITS / PACK_FACTOR
                End If
                CASE_PRICE = Val(rowSOTORDR2.Item("ORDR_PRICE_GRS") & "") * PACK_FACTOR

                .Item("EDI_PO1_QTY") = QTY_CASES
                .Item("EDI_PRICE") = CASE_PRICE
            Else
                .Item("EDI_PO1_QTY") = 0
                .Item("EDI_PRICE") = 0
            End If
            .Item("ORDR_PRICE_GRS") = 0
            .Item("EDI_PO1_UOM") = row850T2.Item("EDI_PO1_UOM")
            .Item("EDI_PO4_PACK_QTY") = row850T2.Item("EDI_PO4_PACK_QTY")
            .Item("EDI_PO4_INNER_PACK_QTY") = row850T2.Item("EDI_PO4_INNER_PACK_QTY")
            .Item("EDI_PO4_UOM") = row850T2.Item("EDI_PO4_UOM")
            .Item("EDI_ITEM") = row850T2.Item("EDI_ITEM")
            If row850T2.Item("ENTRY_TYPE") & "" = "M" And EDI_TP_ID = "621418185" Then
                .Item("EDI_ITEM") = ""
                .Item("EDI_UPC") = row850T2.Item("EDI_ITEM")
            Else
                .Item("EDI_ITEM") = row850T2.Item("EDI_ITEM")
                .Item("EDI_UPC") = row850T2.Item("EDI_UPC")
            End If

            .Item("EDI_SKU") = row850T2.Item("EDI_SKU")
            .Item("EDI_UPC_2") = row850T2.Item("EDI_UPC_2")
            .Item("EDI_GTIN") = row850T2.Item("EDI_GTIN")
            .Item("EDI_PROD_DESC") = row850T2.Item("EDI_ITEM_DESC")
            .Item("EDI_PROD_DESC2") = row850T2.Item("EDI_PROD_DESC2")
            .Item("EDI_DTL_ACK_TYPE") = ""
        End With
        clsASCBASE1.dst.Tables("EDT855O2").Rows.Add(rowEDT855O2)
    End Sub

    Public Shared Sub Fix_Bad_Styles()
        ASCDATA1.ExecuteSQL("UPDATE EDT850T2 SET EDI_STYLE = 'MTX40592', EDI_COLOR_CODE = 'RDGW' WHERE (EDI_STYLE IS NULL OR EDI_COLOR_CODE IS NULL)  AND TRIM(EDI_SKU) = 'MTX40592'")
        ASCDATA1.ExecuteSQL("UPDATE EDT850T2 SET EDI_STYLE = 'MTX41887', EDI_COLOR_CODE = 'MULT' WHERE (EDI_STYLE IS NULL OR EDI_COLOR_CODE IS NULL)  AND TRIM(EDI_SKU) = 'MTX41887'")
        ASCDATA1.ExecuteSQL("UPDATE EDT850T2 SET EDI_STYLE = 'MTX43352', EDI_COLOR_CODE = 'MULT' WHERE (EDI_STYLE IS NULL OR EDI_COLOR_CODE IS NULL)  AND TRIM(EDI_SKU) = 'MTX43352'")
        ASCDATA1.ExecuteSQL("UPDATE EDT850T2 SET EDI_STYLE = 'MTX46919', EDI_COLOR_CODE = 'ANSI' WHERE (EDI_STYLE IS NULL OR EDI_COLOR_CODE IS NULL)  AND TRIM(EDI_SKU) = 'MTX46919'")
        ASCDATA1.ExecuteSQL("UPDATE EDT850T2 SET EDI_STYLE = 'MTX51794', EDI_COLOR_CODE = 'RDWH' WHERE (EDI_STYLE IS NULL OR EDI_COLOR_CODE IS NULL)  AND TRIM(EDI_SKU) = 'MTX51794'")
    End Sub

End Class