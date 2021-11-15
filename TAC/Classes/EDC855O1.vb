Public Class EDC855O1
    Private rowARTPARM1 As DataRow = Nothing
    Private Const EDI_PROCESS_IND As String = "1"

    Private EDI_OUTBOUND_DOC_NO As String = String.Empty

    Private tblSOTSVIA1 As DataTable = Nothing
    Private tblTATTERM1 As DataTable = Nothing
    Private tblEDTTRPM1 As DataTable = Nothing
    Private tblWHTPKGM1 As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing
    Private Shared tblICTSTYC1 As DataTable = Nothing

    Public Shared Function CreateEDTSYSIH(ByRef dst As DataSet, _
                                    ByVal ediOutboundDocNo As String, _
                                    ByVal EDI_OUR_ID As String, _
                                    ByVal EDI_TP_ID As String, _
                                    ByVal ediApplicationID As String, _
                                    ByVal EDI_STATUS As String) As String

        CreateEDTSYSIH = String.Empty

        Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
        rowEDTSYSIH.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
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
        dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

        CreateEDTSYSIH = ediOutboundDocNo

    End Function

    Public Shared Function Generate_855(ByVal clsASCBASE1 As ASCBASE1, ByVal ORDR_GROUP_NO As String) As Boolean

        Dim EDI_OUR_ID As String = ""
        Dim EDI_DOC_SEQ_NO As String = ""
        Dim EDI_PURP_CODE As String = ""
        Dim Err_Msg As String = ""
        Dim EDI_PO_RELEASE_NO = ""

        Dim FactorCode As String = String.Empty
        Dim rowEDTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM EDTPARM1 WHERE ED_PARM_KEY = 'Z'")

        If rowEDTPARM1 IsNot Nothing AndAlso rowEDTPARM1.Table.Columns.Contains("ED_PARM_FACTOR") Then
            FactorCode = rowEDTPARM1.Item("ED_PARM_FACTOR") & String.Empty
        End If

        ' Added 01/17/2019
        Dim rowARTCUST1 As DataRow = Nothing
        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = (SELECT CUST_CODE FROM SOTORDR0 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')")
        Dim rowGLTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM1 WHERE GL_PARM_KEY = 'Z'")

        If clsASCBASE1.dst.Tables.Contains("EDT855O1") Then
            clsASCBASE1.dst.Tables("EDT855O1").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O2").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O3").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O5").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O7").Rows.Clear()
            clsASCBASE1.dst.Tables("EDTSYSIH").Rows.Clear()
        Else
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O1", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O2", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O3", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O5", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O7", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDTSYSIH", "*")
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            EDI_PO_RELEASE_NO = ", t1.EDI_PO_RELEASE_NO,  r1.ORDR_STATUS"
        End If

        ASCMAIN1.sql = "Select r1.ORDR_CUST_PO, r1.ORDR_GROUP_NO, r1.TERM_CODE, r1.ORDR_SHIP_DATE, r1.ORDR_CANCEL_DATE, r1.EDI_DOC_SEQ_NO, " _
        & " t1.EDI_OUR_ID, t1.EDI_TP_QUAL, t1.EDI_TP_ID, t1.EDI_PO_DATE, t1.EDI_ARRIVAL_DATE, t1.EDI_SUPPLIER_NO, R1.CURR_CODE," _
        & " r0.ORDR_AMT_PICK" & EDI_PO_RELEASE_NO _
        & " from SOTORDR1 R1, EDT850T1 T1, SOTORDR0 R0" _
        & " Where R1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
        & " And R1.ORDR_GROUP_NO = R0.ORDR_GROUP_NO " _
        & " And R1.EDI_DOC_SEQ_NO is Not Null" _
        & " And R1.EDI_DOC_SEQ_NO = T1.EDI_DOC_SEQ_NO"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        If rowSOTORDR1 IsNot Nothing Then

            ASCMAIN1.sql = "Select DISTINCT ICTSTYC1.* " _
                & " FROM ICTSTYC1, SOTORDR1, SOTORDR2" _
                & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
                & " AND SOTORDR2.STYLE_CODE = ICTSTYC1.STYLE_CODE" _
                & " AND SOTORDR2.COLOR_CODE = ICTSTYC1.COLOR_CODE" _
                & " AND SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            tblICTSTYC1 = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            EDI_DOC_SEQ_NO = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
            Dim EDI_Outbound_Doc_No As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
            EDI_OUR_ID = Replace(rowSOTORDR1.Item("EDI_OUR_ID"), " ", "")

            ASCMAIN1.sql = "Select * from EDTTRPM1 M1" & vbCrLf _
            & " Where EDI_DOC_NO = '855'" & vbCrLf _
            & " And EDI_TP_QUAL = '" & rowSOTORDR1.Item("EDI_TP_QUAL") & "'" & vbCrLf _
            & " And EDI_TP_ID = rtrim('" & rowSOTORDR1.Item("EDI_TP_ID") & "')" & vbCrLf
            Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow

            If rowEDTTRPM1 IsNot Nothing Then
                Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID")
                Dim Ack_Type As String = "AC"

                ASCMAIN1.sql = "Select IH.* from EDT855O1 O1, EDTSYSIH IH" & vbCrLf _
                & " Where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                & " and O1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" & vbCrLf _
                & " And O1.COMPANY_CODE = IH.COMPANY_CODE" & vbCrLf _
                & " And O1.EDI_OUTBOUND_DOC_NO = IH.EDI_OUTBOUND_DOC_NO" & vbCrLf _
                & " AND EDI_TP_ID NOT IN (SELECT EDI_TP_ID FROM EDTTRPM1 WHERE CUST_CODE = '" & FactorCode & "')"
                Dim rowEDT855OC As DataRow = ASCDATA1.GetDataRow

                If rowEDT855OC IsNot Nothing Then
                    EDI_PURP_CODE = "05"
                Else
                    EDI_PURP_CODE = "00"
                End If

                If ASCMAIN1.CLIENT = "RGI" And rowEDTTRPM1.Item("EDI_TP_ID") & "" = "AMAZONDS" Then
                    If rowSOTORDR1.Item("ORDR_STATUS") & "" = "C" Then
                        Ack_Type = "RD"
                    Else
                        Ack_Type = "AT"
                    End If
                End If

                Dim rowEDT855O1 As DataRow = clsASCBASE1.dst.Tables("EDT855O1").NewRow  ' TBLs("EDT810O1").NewRow
                With rowEDT855O1
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                    .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                    .Item("ORDR_PO_DATE") = rowSOTORDR1.Item("EDI_PO_DATE")
                    .Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO")
                    .Item("REQUEST_DATE") = Format(Now, "dd-MMM-yy")
                    .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
                    .Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                    .Item("ORDR_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
                    .Item("AS_OF_DATE") = Format(Now, "dd-MMM-yy")
                    .Item("INIT_DATE") = Format(Now, "dd-MMM-yy")
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("EDI_PURPOSE_CODE") = EDI_PURP_CODE
                    .Item("EDI_ARRIVAL_DATE") = rowSOTORDR1.Item("EDI_ARRIVAL_DATE")
                    .Item("EDI_ACK_TYPE") = Ack_Type
                    .Item("EDI_SUPPLIER_NO") = rowSOTORDR1.Item("EDI_SUPPLIER_NO")
                    .Item("CURR_CODE") = rowSOTORDR1.Item("CURR_CODE")

                    If ASCMAIN1.CLIENT = "RGI" Then
                        .Item("ORDR_AMT") = rowSOTORDR1.Item("ORDR_AMT_PICK")
                        .Item("EDI_PO_RELEASE_NO") = rowSOTORDR1.Item("EDI_PO_RELEASE_NO")
                    End If

                    ' Added 01/17/2019
                    If rowEDT855O1.Table.Columns.Contains("SEG4_CODE") Then
                        If rowARTCUST1 IsNot Nothing AndAlso rowARTCUST1.Item("SEG4_CODE") & String.Empty <> String.Empty Then
                            .Item("SEG4_CODE") = rowARTCUST1.Item("SEG4_CODE")
                        Else
                            .Item("SEG4_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG4")
                        End If
                    End If
                End With

                clsASCBASE1.dst.Tables("EDT855O1").Rows.Add(rowEDT855O1)

                ASCMAIN1.sql = "Select * from EDT850T2 Where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                For Each rowEDT850T2 As DataRow In ASCDATA1.GetDataTable.Rows
                    Add_EDT855O2(clsASCBASE1, rowEDT850T2, ORDR_GROUP_NO, EDI_Outbound_Doc_No, EDI_TP_ID, Ack_Type)
                    Add_EDT855O3(clsASCBASE1, rowEDT850T2, ORDR_GROUP_NO, EDI_Outbound_Doc_No)
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
                        .Item("EDI_ADDR_TYPE") = rowEDT850T5.Item("EDI_ADDR_TYPE") & ""
                        .Item("EDI_CUST_NAME_ADR") = rowEDT850T5.Item("EDI_CUST_NAME_ADR") & ""
                        .Item("EDI_ADDRESS1") = rowEDT850T5.Item("EDI_ADDRESS1") & ""
                        .Item("EDI_ADDRESS2") = rowEDT850T5.Item("EDI_ADDRESS2") & ""
                        .Item("EDI_ADDRESS3") = ""
                        .Item("EDI_CITY") = rowEDT850T5.Item("EDI_CITY") & ""
                        .Item("EDI_STATE") = rowEDT850T5.Item("EDI_STATE") & ""
                        .Item("EDI_ZIPCODE") = rowEDT850T5.Item("EDI_ZIPCODE") & ""
                        .Item("EDI_COUNTRY") = rowEDT850T5.Item("EDI_COUNTRY") & ""
                        .Item("EDI_ADDR_CODE") = rowEDT850T5.Item("EDI_ADDR_CODE") & ""
                        .Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & ""
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
                        .Item("SAH_AMOUNT") = Val(rowEDT850T7.Item("SAH_AMOUNT") & "")
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

                clsASCBASE1.Update_Record_TDA("EDT855O1")
                clsASCBASE1.Update_Record_TDA("EDT855O2")
                clsASCBASE1.Update_Record_TDA("EDT855O3")
                clsASCBASE1.Update_Record_TDA("EDT855O5")
                clsASCBASE1.Update_Record_TDA("EDT855O7")

                EDI_Outbound_Doc_No = CreateEDTSYSIH(clsASCBASE1.dst, EDI_Outbound_Doc_No, EDI_OUR_ID, EDI_TP_ID, "PR", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)
                clsASCBASE1.Update_Record_TDA("EDTSYSIH")

            End If
        Else
            'Err_Msg = Err_Msg & "Could Not Link Sales Order to 850"
        End If
    End Function

    Public Shared Sub Add_EDT855O2(ByVal clsASCBASE1 As ASCBASE1, _
                  ByVal rowEDT850T2 As DataRow, _
                  ByVal Ordr_Group_No As String, _
                  ByVal EDI_Outbound_Doc_No As String, _
                  ByVal EDI_TP_ID As String, _
                  ByVal Ack_Type As String)

        Dim EDI_PRICE_ACTUAL As Double = 0
        Dim EDI_QTY_OPEN As Double = 0
        Dim EDI_QTY_PICK As Long = 0
        Dim EDI_QTY_CANCEL As Double = 0

        Dim rowEDT855O2 As DataRow = clsASCBASE1.dst.Tables("EDT855O2").NewRow
        With rowEDT855O2
            If ASCMAIN1.CLIENT = "RGI" Then
                ASCMAIN1.sql = "Select MAX(SOTORDR2.ORDR_UNIT_PRICE * nvl(SOTORDR2.SET_QTY,1)) ORDR_UNIT_PRICE, SUM(nvl(SOTORDR2.ORDR_QTY_OPEN,0) / nvl(SOTORDR2.SET_QTY,1)) ORDR_QTY_OPEN, SUM(nvl(SOTORDR2.ORDR_QTY_PICK,0) / nvl(SOTORDR2.SET_QTY,1)) ORDR_QTY_PICK, SUM(nvl(SOTORDR2.ORDR_QTY_CANC,0) / nvl(SOTORDR2.SET_QTY,1)) ORDR_QTY_CANC " _
               & " ,max(nvl(SOTORDRS.ORDR_QTY_ALLO,0)) ORDR_QTY_ALLO" & vbCrLf _
               & " from SOTORDR1" & vbCrLf _
               & " join SOTORDR2 on (SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO) " & vbCrLf _
               & " left outer join SOTORDRS on ( SOTORDRS.STYLE_CODE  = SOTORDR2.STYLE_CODE and SOTORDRS.COLOR_CODE   = SOTORDR2.COLOR_CODE And SOTORDRS.ORDR_GROUP_NO  = SOTORDR1.ORDR_GROUP_NO) " & vbCrLf _
               & " Where SOTORDR1.ORDR_GROUP_NO = '" & Ordr_Group_No & "'" & vbCrLf _
               & " And SOTORDR2.EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")
            Else
                ASCMAIN1.sql = "Select MAX(SOTORDR2.ORDR_UNIT_PRICE) ORDR_UNIT_PRICE, SUM(SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM(SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK, SUM(SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC " _
               & " from SOTORDR2, SOTORDR1" & vbCrLf _
               & " Where SOTORDR1.ORDR_GROUP_NO = '" & Ordr_Group_No & "'" & vbCrLf _
               & " And SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
               & " And EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")
            End If
           
            Dim rowSOTORDR2 As DataRow = ASCDATA1.GetDataRow

            If rowSOTORDR2 IsNot Nothing Then
                EDI_PRICE_ACTUAL = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                EDI_QTY_OPEN = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                EDI_QTY_PICK = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                EDI_QTY_CANCEL = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")
            Else
                EDI_PRICE_ACTUAL = 0
                EDI_QTY_OPEN = 0
                EDI_QTY_PICK = 0
                EDI_QTY_CANCEL = 0
            End If

            'RD is only set for Amazon
            If ASCMAIN1.CLIENT = "RGI" And Ack_Type = "RD" Then
                If Val(rowSOTORDR2.Item("ORDR_QTY_ALLO") & "") > 0 Then
                    EDI_QTY_OPEN = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")
                    EDI_QTY_CANCEL = 0
                End If
            End If

            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
            .Item("EDI_DTL_SEQ") = rowEDT850T2.Item("EDI_DTL_SEQ")
            .Item("EDI_TOTAL_QTY") = rowEDT850T2.Item("EDI_TOTAL_QTY")
            .Item("EDI_UOM") = rowEDT850T2.Item("EDI_PRICE_UOM") & ""
            .Item("EDI_PRICE") = rowEDT850T2.Item("EDI_PRICE")
            .Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
            .Item("EDI_PO4_INNER") = rowEDT850T2.Item("EDI_PO4_INNER")
            .Item("EDI_PO4_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
            .Item("EDI_ITEM") = rowEDT850T2.Item("EDI_ITEM")
            .Item("EDI_UPC") = rowEDT850T2.Item("EDI_UPC")
            .Item("EDI_SKU") = rowEDT850T2.Item("EDI_SKU")
            .Item("EDI_GTIN") = rowEDT850T2.Item("EDI_GTIN")
            .Item("EDI_ITEM_DESC") = rowEDT850T2.Item("EDI_ITEM_DESC")
            .Item("EDI_PO_LNO") = rowEDT850T2.Item("EDI_PO_LNO")
            .Item("EDI_PRICE_ACTUAL") = EDI_PRICE_ACTUAL
            .Item("EDI_QTY_OPEN") = EDI_QTY_OPEN
            .Item("EDI_QTY_PICK") = EDI_QTY_PICK
            .Item("EDI_QTY_CANC") = EDI_QTY_CANCEL
            .Item("EDI_DIMENSION") = rowEDT850T2.Item("EDI_DIMENSION") & ""
        End With
        clsASCBASE1.dst.Tables("EDT855O2").Rows.Add(rowEDT855O2)
    End Sub

    Public Shared Sub Add_EDT855O3(ByVal clsASCBASE1 As ASCBASE1, _
           ByVal rowEDT850T2 As DataRow, _
           ByVal ORDR_GROUP_NO As String, _
           ByVal EDI_OUTBOUND_DOC_NO As String)

        Dim rowEDT855O3 As DataRow = Nothing
        Dim EDI_SDQ_SEQ As Int32 = 0
        Dim fieldNum As Int16 = 0
        Dim UPC_CODE As String = rowEDT850T2.Item("EDI_UPC") & String.Empty
        If UPC_CODE.Length = 0 Then
            UPC_CODE = rowEDT850T2.Item("EDI_EAN") & String.Empty
        End If

        Dim STYLE_CODE As String = String.Empty
        Dim COLOR_CODE As String = String.Empty

        If UPC_CODE.Length > 0 Then
            If tblICTSTYC1.Select("UPC_CODE = '" & UPC_CODE & "'").Length > 0 Then
                STYLE_CODE = tblICTSTYC1.Select("UPC_CODE = '" & UPC_CODE & "'")(0).Item("STYLE_CODE") & String.Empty
                COLOR_CODE = tblICTSTYC1.Select("UPC_CODE = '" & UPC_CODE & "'")(0).Item("COLOR_CODE") & String.Empty
            End If
        End If

        If clsASCBASE1.dst.Tables("SOTORDR2").Select("CUST_UPC = '" & UPC_CODE & "'").Length > 0 Then
            STYLE_CODE = clsASCBASE1.dst.Tables("SOTORDR2").Select("CUST_UPC = '" & UPC_CODE & "'")(0).Item("STYLE_CODE") & String.Empty
            COLOR_CODE = clsASCBASE1.dst.Tables("SOTORDR2").Select("CUST_UPC = '" & UPC_CODE & "'")(0).Item("COLOR_CODE") & String.Empty
        End If

        For Each rowSOTORDR1 As DataRow In clsASCBASE1.dst.Tables("SOTORDR1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")

            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")
            ASCMAIN1.sql = "SELECT * FROM EDTSLSP1 WHERE CUST_CODE = '" & CUST_CODE & "'"
            Dim rowEDTSLSP1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
            Dim storeMaxLen As Int32 = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")

            Dim sql As String = "ORDR_NO = '" & ORDR_NO & "'" _
                                & " AND ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                                & " AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'" _
                                & " AND EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "'" _
                                & " AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")
            '  & " AND ORDR_QTY_PICK > 0" _


            If Not clsASCBASE1.dst.Tables("SOTORDR2").Columns.Contains("ORDR_GROUP_NO") Then
                sql = "ORDR_NO = '" & ORDR_NO & "'" _
                    & " AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'" _
                    & " AND EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "'" _
                    & " AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")
                ' & " AND ORDR_QTY_PICK >= 0"

            End If

            For Each rowSOTORDR2 As DataRow In clsASCBASE1.dst.Tables("SOTORDR2").Select(sql, "")

                Dim EDI_STORE As String = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty

                If storeMaxLen > 0 And EDI_STORE.Length > storeMaxLen Then
                    EDI_STORE = StrReverse(StrReverse(EDI_STORE).Substring(0, storeMaxLen))
                End If

                If IsNumeric(EDI_STORE) AndAlso storeMaxLen > 0 Then
                    EDI_STORE = EDI_STORE.PadLeft(storeMaxLen, "0")
                End If

                fieldNum += 1

                Select Case fieldNum
                    Case 1
                        rowEDT855O3 = clsASCBASE1.dst.Tables("EDT855O3").NewRow
                        rowEDT855O3.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        rowEDT855O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        rowEDT855O3.Item("EDI_DTL_SEQ") = rowEDT850T2.Item("EDI_DTL_SEQ")
                        EDI_SDQ_SEQ += 1
                        rowEDT855O3.Item("EDI_SDQ_SEQ") = EDI_SDQ_SEQ
                        rowEDT855O3.Item("EDI_SDQ_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
                        'rowEDT855O3.Item("EDI_SDQ_QUAL") = String.Empty
                        clsASCBASE1.dst.Tables("EDT855O3").Rows.Add(rowEDT855O3)
                End Select

                rowEDT855O3.Item("EDI_STORE_" & fieldNum.ToString("00")) = EDI_STORE

                ' Done in case something goes awry
                Dim ORDR_QTY_PICK As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                If ORDR_QTY_PICK < 0 Then
                    ORDR_QTY_PICK = 0
                End If
                rowEDT855O3.Item("EDI_QTY_" & fieldNum.ToString("00")) = ORDR_QTY_PICK

                If fieldNum = 10 Then
                    fieldNum = 0
                End If
            Next
        Next

    End Sub

End Class
