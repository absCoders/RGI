Public Class EDC84601

    Private tblEDT846O1 As DataTable = Nothing
    Private tblEDT846O2 As DataTable = Nothing
    Private tblEDTSYSIH As DataTable = Nothing
    Private tblEDTTRPM1 As DataTable = Nothing

    Private rowASTPARM1 As DataRow = Nothing
    Private COMPANY_CODE As String = String.Empty
    Private Const EDI_PROCESS_IND As String = "1"

    Private EDI_OUTBOUND_DOC_NO As String = String.Empty
    Private sQuery As New System.Text.StringBuilder With {.Length = 0}

    Public Sub New(ByRef datasetIn As DataSet)
        InitializeData(datasetIn.Tables("EDT846O1"),
                       datasetIn.Tables("EDT846O2"),
                       datasetIn.Tables("EDTSYSIH"))
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tblEDT846O1in">Reference to table EDTSYSIH</param>
    ''' <param name="tblEDT846O2in">Reference to table EDT856O1</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tblEDT846O1in As DataTable,
                   ByRef tblEDT846O2in As DataTable,
                   ByRef tblEDTSYSIHin As DataTable)

        InitializeData(tblEDT846O1in,
                       tblEDT846O2in,
                       tblEDTSYSIHin)
    End Sub

    Private Sub InitializeData(ByRef tblEDT846O1in As DataTable,
                               ByRef tblEDT846O2in As DataTable,
                               ByRef tblEDTSYSIHin As DataTable)

        COMPANY_CODE = ASCMAIN1.CLIENT

        tblEDT846O1 = tblEDT846O1in
        tblEDT846O2 = tblEDT846O2in
        tblEDTSYSIH = tblEDTSYSIHin
        tblEDTTRPM1 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM1 where EDI_DOC_NO = '846'", "EDTTRPM1", String.Empty, Nothing)

        rowASTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM ASTPARM1 WHERE AS_PARM_KEY = 'Z'")


        EDI_OUTBOUND_DOC_NO = String.Empty

    End Sub

    ''' <summary>
    ''' Creates the EDT846 table entries
    ''' </summary>
    ''' <param name="ECOM_CODE"></param>
    ''' <param name="ErrorMessage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateEDI846(ByVal ECOM_CODE As String, ByRef ErrorMessage As String) As String

        Dim tblICTSTATX As DataTable = Nothing
        Dim rowECTECOM1 As DataRow = Nothing
        Dim rowEDTTRPM1 As DataRow = Nothing
        'Dim rowEDTXREF4 As DataRow = Nothing
        Dim dataFound As Boolean = False
        Dim tblECTESTY3 As DataTable = Nothing
        Dim tblSOTCSTY1 As DataTable = Nothing

        sQuery.Length = 0
        sQuery.AppendLine("SELECT *")
        sQuery.AppendLine("FROM ECTECOM1")
        sQuery.AppendLine("WHERE ECOM_CODE = :PARM1")
        rowECTECOM1 = ASCDATA1.GetDataRow(sQuery.ToString, "V", New Object() {ECOM_CODE})
        If rowECTECOM1 Is Nothing Then
            ErrorMessage = "Cannot locate ECTECOM1 Record For " & ECOM_CODE
            Return String.Empty
        End If

        Dim EDI_TP_ID As String = rowECTECOM1.Item("EDI_TP_ID") & String.Empty
        Dim EDI_TP_QUAL As String = rowECTECOM1.Item("EDI_TP_QUAL") & String.Empty

        Dim Sql As String = "EDI_TP_QUAL = '" & EDI_TP_QUAL & "' and EDI_TP_ID = '" & EDI_TP_ID & "' and EDI_DOC_NO = '846'"

        If tblEDTTRPM1.Select(Sql).Length = 0 Then
            ErrorMessage = "Customer with Ecommerce Code (" & ECOM_CODE & ") is not setup to receive 846 data."
            Return String.Empty
        Else
            rowEDTTRPM1 = tblEDTTRPM1.Select(Sql)(0)
        End If
        Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty

        Dim ECOM_MIN_QTY_DEFAULT As Integer = Val(rowECTECOM1.Item("ECOM_MIN_QTY_DEFAULT").ToString & "")
        If ECOM_MIN_QTY_DEFAULT = 0 Then
            ErrorMessage = "ECOM_MIN_QTY_DEFAULT Set to 0" & ECOM_CODE
            Return String.Empty
        End If

        Dim FORCE_QTY_ZERO = rowECTECOM1.Item("FORCE_QTY_ZERO").ToString & ""
        If FORCE_QTY_ZERO = "" Then FORCE_QTY_ZERO = "0"

        Dim ECOM_ALLOC_PCT_DEFAULT As Double = Val(rowECTECOM1.Item("ECOM_ALLOC_PCT_DEFAULT").ToString & "")
        If ECOM_ALLOC_PCT_DEFAULT = 0 Then
            ErrorMessage = "ECOM_ALLOC_PCT_DEFAULT Set to 0" & ECOM_CODE
            Return String.Empty
        End If

        Sql = " SELECT * FROM SOTCSTY1 WHERE CUST_CODE IN"
        Sql &= " ("
        Sql &= " SELECT EDTTRPM1.CUST_CODE"
        Sql &= " FROM ECTECOM1, EDTTRPM1"
        Sql &= " where ECTECOM1.EDI_TP_QUAL = EDTTRPM1.EDI_TP_QUAL"
        Sql &= " and ECTECOM1.EDI_TP_ID = EDTTRPM1.EDI_TP_ID"
        Sql &= " and EDTTRPM1.EDI_DOC_NO = '846' "
        Sql &= " AND ECTECOM1.EDI_846_INDICATOR = '1' "
        Sql &= " AND NVL(ECTECOM1.EDI_846_INTERVAL, 0) > 0"
        Sql &= " AND ECTECOM1.ECOM_CODE = '" & ECOM_CODE & "'"
        Sql &= " )"
        tblSOTCSTY1 = ASCDATA1.GetDataTable(Sql)

        'sQuery.Length = 0
        'sQuery.AppendLine("SELECT X4.*")
        'sQuery.AppendLine("FROM EDTXREF4 X4, EDTTRPM1 PM")
        'sQuery.AppendLine("WHERE X4.SENDER_ID_QUAL = PM.EDI_TP_QUAL")
        'sQuery.AppendLine("AND X4.SENDER_ID = PM.EDI_TP_ID")
        'sQuery.AppendLine("AND PM.EDI_DOC_NO = '846'")
        'sQuery.AppendLine("AND X4.WHSE_CODE = 'MS'")
        'sQuery.AppendLine("AND X4.SENDER_ID = :PARM1")
        'rowEDTXREF4 = ASCDATA1.GetDataRow(sQuery.ToString, "V", New Object() {EDI_TP_ID})
        'If rowEDTXREF4 Is Nothing Then
        '    ErrorMessage = "Cannot locate EDTXREF4 Record For " & EDI_TP_ID
        '    Return String.Empty
        'End If

        Dim ICTPEND_TEMP As String = MAKE_ICTPEND_TEMP()

        sQuery.Length = 0
        sQuery.AppendLine("SELECT")
        sQuery.AppendLine("S1.STYLE_CODE,")
        sQuery.AppendLine("S2.COLOR_CODE,")
        sQuery.AppendLine("C1.COLOR_DESC,")
        sQuery.AppendLine("(S1.STYLE_CODE || '-' || S2.COLOR_CODE) AS EDI_SKU,")
        sQuery.AppendLine("S1.STYLE_DESC AS EDI_ITEM_DESC,")
        sQuery.AppendLine("SC.UPC_CODE,")
        sQuery.AppendLine("S1.STYLE_UOM,")
        sQuery.AppendLine("E1.ECOM_UNIT_PRICE,")
        sQuery.AppendLine("CASE WHEN '" & FORCE_QTY_ZERO & "' = '1' THEN 0")
        sQuery.AppendLine(String.Format("WHEN SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0) - NVL(P1.ORDR_QTY_PEND,0)) / GREATEST(DECODE(NVL(E1.SET_QTY,1),0,1, NVL(E1.SET_QTY,1)))) < case when max(nvl(E1.ECOM_MIN_QTY_OVERRIDE,0)) > 0 then max(E1.ECOM_MIN_QTY_OVERRIDE) else {0} end THEN", ECOM_MIN_QTY_DEFAULT))
        sQuery.AppendLine("  0")
        sQuery.AppendLine("ELSE")
        sQuery.AppendLine("  TRUNC(SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0) - NVL(P1.ORDR_QTY_PEND,0)) / GREATEST(DECODE(NVL(E1.SET_QTY,1),0,1, NVL(E1.SET_QTY,1)))))")
        sQuery.AppendLine("END AS ON_HAND,")
        sQuery.AppendLine(String.Format("CASE WHEN SC.STYLE_COLOR_STATUS = 'D' AND SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0) - NVL(P1.ORDR_QTY_PEND,0))  / GREATEST(DECODE(NVL(E1.SET_QTY,1),0,1, NVL(E1.SET_QTY,1)))) < case when max(nvl(E1.ECOM_MIN_QTY_OVERRIDE,0)) > 0 then max(E1.ECOM_MIN_QTY_OVERRIDE) else {0} end THEN", ECOM_MIN_QTY_DEFAULT))
        sQuery.AppendLine("  '002'")
        sQuery.AppendLine("ELSE")
        sQuery.AppendLine("  '001'")
        sQuery.AppendLine("END AS EDI_MAINT_TYPE_CODE")
        sQuery.AppendLine(String.Format("FROM ICTSTYL1 S1, ICTSTAT2 S2, ECTESTY1 E1, ECTESTY2 E2, ICTCOLR1 C1, ICTSTYC1 SC, {0} P1", ICTPEND_TEMP))
        sQuery.AppendLine("WHERE E1.STYLE_CODE = E2.STYLE_CODE")
        sQuery.AppendLine("AND NVL(E1.SHIP_DROP,'0') = '1'")
        sQuery.AppendLine("AND S1.STYLE_CODE = S2.STYLE_CODE (+)")
        sQuery.AppendLine("AND E2.STYLE_CODE = P1.STYLE_CODE (+)")
        sQuery.AppendLine("AND E2.COLOR_CODE = P1.COLOR_CODE (+)")
        sQuery.AppendLine("AND S2.STYLE_CODE = E2.STYLE_CODE (+)")
        sQuery.AppendLine("AND S2.COLOR_CODE = E2.COLOR_CODE (+)")
        sQuery.AppendLine("AND E2.ECOM_CODE = E1.ECOM_CODE")
        sQuery.AppendLine("AND S2.COLOR_CODE = C1.COLOR_CODE")
        sQuery.AppendLine("AND S2.STYLE_CODE = SC.STYLE_CODE")
        sQuery.AppendLine("AND S2.COLOR_CODE = SC.COLOR_CODE")
        sQuery.AppendLine("AND E2.ECOM_STYLE_COLOR_STATUS = 'A'")
        sQuery.AppendLine("AND S2.WHSE_CODE = 'MS'")
        sQuery.AppendLine(String.Format("AND E2.ECOM_CODE = '{0}'", ECOM_CODE))
        sQuery.AppendLine("GROUP BY S1.STYLE_CODE,")
        sQuery.AppendLine("S2.COLOR_CODE,")
        sQuery.AppendLine("C1.COLOR_DESC,")
        sQuery.AppendLine("(S1.STYLE_CODE || '-' || S2.COLOR_CODE),")
        sQuery.AppendLine("S1.STYLE_DESC,")
        sQuery.AppendLine("SC.UPC_CODE,")
        sQuery.AppendLine("SC.STYLE_COLOR_STATUS,")
        sQuery.AppendLine("S1.STYLE_UOM,")
        sQuery.AppendLine("E1.ECOM_UNIT_PRICE")

        Dim wkTable As String = ASCMAIN1.Temp_Table(sQuery.ToString)
        tblICTSTATX = ASCDATA1.GetDataTable($"Select * from {wkTable}")

        If tblICTSTATX.Rows.Count = 0 Then
            ErrorMessage = "No Records Found To Transfer For " & ECOM_CODE
            Return String.Empty
        End If

        Dim tblPOTORDR2 As DataTable = ASCDATA1.GetDataTable($"Select PO_ORDER_NO,STYLE_CODE, COLOR_CODE, PO_QTY_OPN, PO_DATE_ETA
                    from POTORDR2 
                    WHERE (STYLE_CODE, COLOR_CODE) IN (SELECT STYLE_CODE, COLOR_CODE FROM {wkTable})
                    AND PO_QTY_OPN > 0 ")

        Dim EDI_OUTBOUND_DOC_NO As String = Me.CreateEDTSYSIH(EDI_OUR_ID, EDI_TP_ID, "IB", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)

        Dim rowEDT846O1 As DataRow = tblEDT846O1.NewRow
        rowEDT846O1.Item("COMPANY_CODE") = COMPANY_CODE
        rowEDT846O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
        rowEDT846O1.Item("EDI_COMPANY_NAME") = rowASTPARM1.Item("AS_PARM_INST_NAME").ToString & ""
        rowEDT846O1.Item("EDI_ADDR1") = rowASTPARM1.Item("AS_PARM_INST_ADDR1").ToString & ""
        rowEDT846O1.Item("EDI_ADDR2") = rowASTPARM1.Item("AS_PARM_INST_ADDR2").ToString & ""
        rowEDT846O1.Item("EDI_ADDR3") = rowASTPARM1.Item("AS_PARM_INST_ADDR3").ToString & ""
        rowEDT846O1.Item("EDI_CITY") = rowASTPARM1.Item("AS_PARM_INST_CITY").ToString & ""
        rowEDT846O1.Item("EDI_STATE") = rowASTPARM1.Item("AS_PARM_INST_STATE").ToString & ""
        rowEDT846O1.Item("EDI_ZIP_CODE") = rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE").ToString & ""
        rowEDT846O1.Item("EDI_COUNTRY") = rowASTPARM1.Item("AS_PARM_INST_COUNTRY").ToString & ""
        rowEDT846O1.Item("INIT_DATE") = DateTime.Now
        rowEDT846O1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowEDT846O1.Item("EDI_REPORT_DATE") = DateTime.Now
        rowEDT846O1.Item("EDI_SUPPLIER_NO") = rowEDTTRPM1.Item("EDI_ACCT_REF_NO").ToString & ""
        tblEDT846O1.Rows.Add(rowEDT846O1)

        sQuery.Length = 0
        sQuery.AppendLine("SELECT *")
        sQuery.AppendLine("FROM ECTESTY3")
        sQuery.AppendLine(String.Format("WHERE ECOM_CODE = '{0}'", ECOM_CODE))
        sQuery.AppendLine("and SYSDATE between PROMO_START_DATE and PROMO_END_DATE + (1-1/24/60/60)")
        tblECTESTY3 = ASCDATA1.GetDataTable(sQuery.ToString)


        Dim EDI_DOC_LNO As Integer = 0
        For Each rowICTSTATX As DataRow In tblICTSTATX.Select()

            Dim EDI_SKU As String = rowICTSTATX.Item("EDI_SKU") & String.Empty
            If tblSOTCSTY1.Select("STYLE_CODE = '" & rowICTSTATX.Item("STYLE_CODE") & "' and COLOR_CODE = '" & rowICTSTATX.Item("COLOR_CODE") & "'").Length > 0 Then
                EDI_SKU = tblSOTCSTY1.Select("STYLE_CODE = '" & rowICTSTATX.Item("STYLE_CODE") & "' and COLOR_CODE = '" & rowICTSTATX.Item("COLOR_CODE") & "'")(0).Item("CUST_STYLE_CODE") & String.Empty

                ' just is case
                EDI_SKU = EDI_SKU.Trim
                If EDI_SKU.Length = 0 Then
                    EDI_SKU = rowICTSTATX.Item("EDI_SKU") & String.Empty
                End If
            End If

            If ECOM_CODE = "KIRKLANDS" Then
                If EDI_SKU = rowICTSTATX.Item("EDI_SKU") & String.Empty Then
                    'RecordLogEntry("Style/Color " & EDI_SKU & " does not have a Kirklands Sku. Item was skipped in the Kirklands EDI 846.", True)
                    Continue For
                End If
            End If

            Dim RsrvQtyOpen As Int16 = 0
            Dim RsrvSQL As String = "Select SOTRSRV2.* from SOTRSRV2,SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.CUST_CODE = :PARM1 " & vbCrLf _
                & "   and SOTRSRV2.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTRSRV2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_QTY_OPEN > 0" & vbCrLf
            For Each rowSOTRSRV2 As DataRow In ASCDATA1.GetDataTable(RsrvSQL, "", -1, True, -1, "VVV", New Object() {rowECTECOM1.Item("CUST_CODE"), rowICTSTATX.Item("STYLE_CODE"), rowICTSTATX.Item("COLOR_CODE")}).Select("")
                RsrvQtyOpen += Val(rowSOTRSRV2.Item("RSRV_QTY_OPEN").ToString & "")
            Next

            Dim rowEDT846O2 As DataRow = tblEDT846O2.NewRow
            rowEDT846O2.Item("COMPANY_CODE") = COMPANY_CODE
            rowEDT846O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO

            EDI_DOC_LNO += 1
            rowEDT846O2.Item("EDI_DOC_LNO") = EDI_DOC_LNO

            rowEDT846O2.Item("EDI_ITEM") = rowICTSTATX.Item("EDI_SKU").ToString & ""
            rowEDT846O2.Item("EDI_ITEM_DESC") = rowICTSTATX.Item("EDI_ITEM_DESC").ToString & ""
            rowEDT846O2.Item("EDI_SKU") = EDI_SKU ' rowICTSTATX.Item("EDI_SKU").ToString & ""
            rowEDT846O2.Item("EDI_STYLE") = rowICTSTATX.Item("STYLE_CODE").ToString & ""
            rowEDT846O2.Item("EDI_STYLE_NAME") = rowICTSTATX.Item("EDI_ITEM_DESC").ToString & ""
            rowEDT846O2.Item("EDI_COLOR_CODE") = rowICTSTATX.Item("COLOR_CODE").ToString & ""
            rowEDT846O2.Item("EDI_COLOR_NAME") = rowICTSTATX.Item("COLOR_DESC").ToString & ""
            rowEDT846O2.Item("EDI_UPC") = rowICTSTATX.Item("UPC_CODE").ToString & ""
            rowEDT846O2.Item("EDI_AVAIL_QTY") = Val(rowICTSTATX.Item("ON_HAND").ToString & "") + RsrvQtyOpen
            rowEDT846O2.Item("EDI_ITEM_UOM") = rowICTSTATX.Item("STYLE_UOM").ToString & ""
            rowEDT846O2.Item("EDI_MAINT_TYPE_CODE") = rowICTSTATX.Item("EDI_MAINT_TYPE_CODE").ToString & ""
            rowEDT846O2.Item("ECOM_UNIT_PRICE") = Val(rowICTSTATX.Item("ECOM_UNIT_PRICE").ToString & "")

            If rowICTSTATX.Item("EDI_MAINT_TYPE_CODE").ToString & "" = "002" Then
                Dim STYLE_CODE As String = rowICTSTATX.Item("STYLE_CODE").ToString & ""
                Dim COLOR_CODE As String = rowICTSTATX.Item("COLOR_CODE").ToString & ""
                Dim DISC_DATE As String = ASCDATA1.GetDataValue("SELECT TO_CHAR(DISC_DATE,'MM/DD/YYYY') from ECTESTY2 Where ECOM_CODE = :PARM1 and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3", "VVV", New Object() {ECOM_CODE, STYLE_CODE, COLOR_CODE})
                If String.IsNullOrEmpty(DISC_DATE) Then
                    ASCMAIN1.sql = "UPDATE ECTESTY2 SET DISC_DATE = TRUNC(SYSDATE) Where ECOM_CODE = :PARM1 and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {ECOM_CODE, STYLE_CODE, COLOR_CODE})
                    DISC_DATE = Date.Today
                ElseIf DateAdd(DateInterval.Day, CDbl(rowECTECOM1.Item("DISC_DAYS")), CDate(DISC_DATE)) < DateTime.Now Then
                    Continue For
                End If
                rowEDT846O2.Item("EDI_DISCONTINUE_DATE") = CDate(DISC_DATE)
            End If

            ' 09/21/2020
            ' POTORDR2.PO_QTY_OPN >> EDT846O2.EDI_FUTURE_QTY
            ' POTORDR2.PO_DATE_ETA >> EDT846O2.EDI_FUTURE_DATE
            ' Use only the first one closest to today's date.
            Sql = $"STYLE_CODE = '{rowICTSTATX.Item("STYLE_CODE")}' and COLOR_CODE = '{rowICTSTATX.Item("COLOR_CODE")}'"
            For Each rowPOTORDR2 As DataRow In tblPOTORDR2.Select(Sql, "PO_DATE_ETA")
                rowEDT846O2.Item("EDI_FUTURE_QTY") = Val(rowPOTORDR2.Item("PO_QTY_OPN").ToString & "")
                rowEDT846O2.Item("EDI_FUTURE_DATE") = rowPOTORDR2.Item("PO_DATE_ETA")
                Exit For
            Next


            'rowEDT846O2.Item("EDI_SIZE_CODE") = NULL
            'rowEDT846O2.Item("EDI_SIZE_DESC") = NULL
            'rowEDT846O2.Item("EDI_EAN") = NULL
            'rowEDT846O2.Item("EDI_GTIN") = NULL
            'rowEDT846O2.Item("EDI_FUTURE_QTY") = NULL
            'rowEDT846O2.Item("EDI_FUTURE_DATE") = NULL
            'rowEDT846O2.Item("EDI_DISCONTINUE_DATE") = NULL
            For Each rowECTESTY3 As DataRow In tblECTESTY3.Select(String.Format("STYLE_CODE = '{0}' and COLOR_CODE = '{1}'", rowICTSTATX.Item("STYLE_CODE").ToString & "", rowICTSTATX.Item("COLOR_CODE").ToString & ""))
                rowEDT846O2.Item("PROMO_UNIT_PRICE") = Val(rowECTESTY3.Item("PROMO_UNIT_PRICE").ToString & "")
            Next
            tblEDT846O2.Rows.Add(rowEDT846O2)
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

    Private Function MAKE_ICTPEND_TEMP() As String
        Dim RETVAL As String = ""
        Dim s As New System.Text.StringBuilder With {.Length = 0}
        s.AppendLine("SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_QTY_PEND) AS ORDR_QTY_PEND")
        s.AppendLine("FROM")
        s.AppendLine("(")
        s.AppendLine("  SELECT 'L' SOURCE, L2.STYLE_CODE,")
        s.AppendLine("  L2.COLOR_CODE,")
        s.AppendLine("  SUM(L2.ORDR_QTY_OPEN) ORDR_QTY_PEND")
        s.AppendLine("  FROM SOTORDR1_L L1,")
        s.AppendLine("  SOTORDR2_L L2")
        s.AppendLine("  WHERE L1.ORDR_NO   = L2.ORDR_NO")
        s.AppendLine("  AND L1.ORDR_STATUS = 'O'")
        s.AppendLine("  GROUP BY L2.STYLE_CODE,")
        s.AppendLine("  L2.COLOR_CODE")
        s.AppendLine("  UNION")
        s.AppendLine("  SELECT 'E' SOURCE, SC.STYLE_CODE,")
        s.AppendLine("  SC.COLOR_CODE,")
        s.AppendLine("  SUM(EDI_TOTAL_QTY) AS ORDR_QTY_PEND")
        s.AppendLine("  FROM EDTTRPM1 P1, EDTXREF4 X4,")
        s.AppendLine("  EDT850T1 T1,")
        s.AppendLine("  EDT850T2 T2,")
        s.AppendLine("  ICTSTYC1 SC")
        s.AppendLine("  WHERE T1.EDI_DOC_SEQ_NO = T2.EDI_DOC_SEQ_NO")
        s.AppendLine("  AND P1.EDI_TP_ID = TRIM(T1.EDI_TP_ID)")
        s.AppendLine("  AND T1.EDI_TP_QUAL = X4.SENDER_ID_QUAL")
        s.AppendLine("  AND TRIM(T1.EDI_TP_ID) = X4.SENDER_ID")
        s.AppendLine("  and T1.EDI_SUPPLIER_NO = X4.EDI_SUPPLIER_NO")
        s.AppendLine("  AND P1.EDI_DOC_NO = '850'")
        s.AppendLine("  AND P1.EDI_STATUS = 'P'")
        s.AppendLine("  AND X4.WHSE_CODE = 'MS'")
        s.AppendLine("  AND NVL(T1.EDI_PROCESS_IND, '0') = '0'")
        s.AppendLine("  AND SC.STYLE_CODE || '-' || SC.COLOR_CODE in (t2.EDI_UPC, t2.EDI_SKU, t2.EDI_STYLE || '-' || t2.EDI_COLOR_CODE)")
        s.AppendLine("  GROUP BY SC.STYLE_CODE,")
        s.AppendLine("  SC.COLOR_CODE")
        s.AppendLine(") GROUP BY STYLE_CODE, COLOR_CODE")
        ASCMAIN1.sql = s.ToString
        RETVAL = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RETVAL & " Add Primary Key (STYLE_CODE, COLOR_CODE)")
        Return RETVAL
    End Function

End Class
