Public Class EDC753O1

    Private Const EDI_PROCESS_IND As String = "1"
    Private clsASCBASE1 As ASCBASE1

    Private pick2Query As String = String.Empty

    Public Sub New(ByVal clsASCBASE1in As ASCBASE1)
        clsASCBASE1 = clsASCBASE1in
    End Sub

    Private Function CreateEDTSYSIH(ByRef dst As DataSet, _
                                ByVal EDI_OUTBOUND_DOC_NO As String, _
                                ByVal EDI_OUR_ID As String, _
                                ByVal EDI_TP_ID As String, _
                                ByVal ediApplicationID As String, _
                                ByVal EDI_STATUS As String) As String

        CreateEDTSYSIH = String.Empty

        Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
        rowEDTSYSIH.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
        rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
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

        CreateEDTSYSIH = EDI_OUTBOUND_DOC_NO

    End Function

    Public Sub Generate_753(ByVal SHIP_BOL_NO As String)

        Try

            If ASCMAIN1.CLIENT <> "NYA" Then
                Exit Sub
            End If

            pick2Query = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY, SOTPICK1.SHIP_BOL_NO," & vbCrLf _
                & "  ICTSTYC1.STYLE_BIN, ICTSTYC1.STYLE_BIN as LOCATION_CODE, ICTSTYL1.CASE_CUBE, ICTSTYC1.UPC_CODE, ICTSTYL1.CASE_WEIGHT_GRS" & vbCrLf _
                & " from SOTPICK2, SOTPICK1, SOTORDR2, ICTSTYL1, ICTSTYC1" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTPICK2.PICK_QTY <> 0"


            If clsASCBASE1.dst.Tables.Contains("EDT753O1") Then
                clsASCBASE1.dst.Tables("EDT753O1").Rows.Clear()
                clsASCBASE1.dst.Tables("EDT753O2").Rows.Clear()
                clsASCBASE1.dst.Tables("EDT753O3").Rows.Clear()
                clsASCBASE1.dst.Tables("EDTSYSIH").Rows.Clear()

                clsASCBASE1.dst.Tables("SOTSHIP1_WK").Rows.Clear()
                clsASCBASE1.dst.Tables("SOTPICK1_WK").Rows.Clear()
                clsASCBASE1.dst.Tables("SOTPICK2_WK").Rows.Clear()

                clsASCBASE1.dst.Tables("SOTCART1_WK").Rows.Clear()
                clsASCBASE1.dst.Tables("SOTCART2_WK").Rows.Clear()

                clsASCBASE1.dst.Tables("SOTORDR1_WK").Rows.Clear()
                clsASCBASE1.dst.Tables("SOTORDR2_WK").Rows.Clear()
            Else
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT753O1", "*")
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT753O2", "*")
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT753O3", "*")
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDTSYSIH", "*")

                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTSHIP1_WK"), "SOTSHIP1", "*")
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTPICK1_WK"), "SOTPICK1", "*")
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTPICK2_WK"), "SOTPICK2", pick2Query)

                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTCART1_WK"), "SOTCART1", "*")
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTCART2_WK"), "SOTCART2", "*")

                clsASCBASE1.dst.Tables("SOTCART1_WK").Columns.Add("STYLE_CODE", GetType(System.String))
                clsASCBASE1.dst.Tables("SOTCART1_WK").Columns.Add("STYLE_DESC", GetType(System.String))
                clsASCBASE1.dst.Tables("SOTCART1_WK").Columns.Add("NUM_CARTONS", GetType(System.Int32))
                clsASCBASE1.dst.Tables("SOTCART1_WK").Columns.Add("TOTAL_WEIGHT", GetType(System.Decimal))
                clsASCBASE1.dst.Tables("SOTCART1_WK").Columns.Add("CASE_CUBE_FT", GetType(System.Decimal))

                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTORDR1_WK"), "SOTORDR1", "*")
                clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTORDR2_WK"), "SOTORDR2", "*")
            End If

            Load_SOTPICK1(SHIP_BOL_NO)

            Dim rowSOTSHIP1 As DataRow = clsASCBASE1.dst.Tables("SOTSHIP1_WK").Rows.Find(SHIP_BOL_NO)
            If rowSOTSHIP1 Is Nothing Then
                Exit Sub
            End If

            Dim rowICTWHSE1 As DataRow = clsASCBASE1.LookUp("ICTWHSE1", rowSOTSHIP1.Item("WHSE_CODE") & String.Empty)
            Dim rowARTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTPARM1 WHERE AR_PARM_KEY = 'Z'")
            Dim ORDR_GROUP_NO As String = rowSOTSHIP1.Item("ORDR_GROUP_NO") & String.Empty
            Dim rowSOTORDR0 As DataRow = clsASCBASE1.LookUp("SOTORDR0", ORDR_GROUP_NO)
            Dim CUST_CODE As String = rowSOTORDR0.Item("CUST_CODE")

            Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM EDTTRPM1 WHERE CUST_CODE = '" & CUST_CODE & "' AND EDI_DOC_NO = '753'")

            If rowEDTTRPM1 Is Nothing Then
                Exit Sub
            End If

            Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty
            Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty

            For Each rowSOTPICK1 As DataRow In clsASCBASE1.dst.Tables("SOTPICK1_WK").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "PICK_NO")
                Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                Dim rowSOTORDR1 As DataRow = clsASCBASE1.dst.Tables("SOTORDR1_WK").Rows.Find(rowSOTPICK1.Item("ORDR_NO"))

                Dim rowEDT753O1 As DataRow = clsASCBASE1.dst.Tables("EDT753O1").NewRow
                rowEDT753O1("COMPANY_CODE") = ASCMAIN1.CLIENT
                rowEDT753O1("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT753O1("EDI_TRANS_DATE_TIME") = System.DateTime.Now
                rowEDT753O1("WHSE_PHONE") = rowICTWHSE1.Item("WHSE_PHONE") & String.Empty
                rowEDT753O1("WHSE_EXT") = rowICTWHSE1.Item("WHSE_EXT") & String.Empty
                rowEDT753O1("WHSE_FAX") = rowICTWHSE1.Item("WHSE_FAX") & String.Empty
                rowEDT753O1("EDI_SUPPLIER_NO") = rowEDTTRPM1.Item("EDI_ACCT_REF_NO") & String.Empty
                rowEDT753O1("WHSE_ADDR1") = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
                rowEDT753O1("WHSE_ADDR2") = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
                rowEDT753O1("WHSE_CITY") = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
                rowEDT753O1("WHSE_STATE") = rowICTWHSE1.Item("WHSE_STATE") & String.Empty
                rowEDT753O1("WHSE_ZIP_CODE") = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
                rowEDT753O1("WHSE_COUNTRY") = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
                rowEDT753O1("AR_PARM_REMIT_NAME") = rowARTPARM1.Item("AR_PARM_REMIT_NAME") & String.Empty
                rowEDT753O1("WHSE_DESC") = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                ' Hard Coded for Walmart. Currenty only Walmart/NYA uses this.
                rowEDT753O1("SHIP_POINT") = "08885606"
                clsASCBASE1.dst.Tables("EDT753O1").Rows.Add(rowEDT753O1)

                Dim rowEDT753O2 As DataRow = clsASCBASE1.dst.Tables("EDT753O2").NewRow
                rowEDT753O2("COMPANY_CODE") = ASCMAIN1.CLIENT
                rowEDT753O2("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT753O2("EDI_753_SEQ_NO") = 1
                rowEDT753O2("SHIP_ADDR_CODE") = rowSOTORDR1.Item("CUST_DC_NO") & String.Empty
                rowEDT753O2("EDI_RRC_NO") = ASCMAIN1.Next_Control_No("EDT753O2.EDI_RRC_NO")
                rowEDT753O2("EDI_RTS_DATE") = CDate(rowEDT753O1("EDI_TRANS_DATE_TIME")).ToShortDateString
                rowEDT753O2("SHIP_BOL_NO") = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                rowEDT753O2("EDI_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE") & String.Empty
                rowEDT753O2("SHIP_TO_NAME") = rowSOTORDR1.Item("CUST_STORE_NAME") & String.Empty
                clsASCBASE1.dst.Tables("EDT753O2").Rows.Add(rowEDT753O2)

                Dim EDI_753_ORD_SEQ_NO As Int32 = 1
                Dim rowEDT753O3 As DataRow = clsASCBASE1.dst.Tables("EDT753O3").NewRow
                rowEDT753O3("COMPANY_CODE") = ASCMAIN1.CLIENT
                rowEDT753O3("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT753O3("EDI_753_SEQ_NO") = rowEDT753O2("EDI_753_SEQ_NO")
                rowEDT753O3("EDI_753_ORD_SEQ_NO") = EDI_753_ORD_SEQ_NO
                EDI_753_ORD_SEQ_NO += 1
                rowEDT753O3("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                rowEDT753O3("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO") & String.Empty
                rowEDT753O3("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT") & String.Empty

                rowEDT753O3("EDI_753_CNT_CARTONS") = Val(clsASCBASE1.dst.Tables("SOTCART1_WK").Compute("SUM(NUM_CARTONS)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)
                rowEDT753O3("EDI_753_TOTAL_WGT") = Val(clsASCBASE1.dst.Tables("SOTCART1_WK").Compute("SUM(TOTAL_WEIGHT)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)

                ' As per Debbie Stark at NYA, ad an additional 15% to the Cubic feet.
                rowEDT753O3("EDI_753_CUBIC_FEET") = Math.Round(Val(clsASCBASE1.dst.Tables("SOTCART1_WK").Compute("SUM(CASE_CUBE_FT)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty) * 1.15, 0)
                ' Cubic feet should be at least 1
                If rowEDT753O3("EDI_753_CUBIC_FEET") < 1 Then
                    rowEDT753O3("EDI_753_CUBIC_FEET") = 1
                End If
                rowEDT753O3("EDI_753_CNT_PCS") = Val(clsASCBASE1.dst.Tables("SOTCART1_WK").Compute("SUM(CART_TOTAL_UNITS)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)

                clsASCBASE1.dst.Tables("EDT753O3").Rows.Add(rowEDT753O3)

                clsASCBASE1.Update_Record_TDA("EDT753O1")
                clsASCBASE1.Update_Record_TDA("EDT753O2")
                clsASCBASE1.Update_Record_TDA("EDT753O3")

                EDI_OUTBOUND_DOC_NO = CreateEDTSYSIH(clsASCBASE1.dst, EDI_OUTBOUND_DOC_NO, EDI_OUR_ID, EDI_TP_ID, "RF", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)
                clsASCBASE1.Update_Record_TDA("EDTSYSIH")

                For Each tableName As String In New String() {"EDT753O1", "EDT753O2", "EDT753O3", "EDTSYSIH"}
                    clsASCBASE1.dst.Tables(tableName).Rows.Clear()
                Next

                ASCDATA1.ExecuteSQL("UPDATE SOTSHIP1 SET SHIP_XMIT_FLAG = 'H' WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")

            Next

        Catch ex As Exception
            MessageBox.Show("EDI 753 failed due to the following error: " & ex.Message, "Generate EDI 753", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Load_SOTPICK1(ByVal SHIP_BOL_NO As String)

        clsASCBASE1.EnforceConstraints(False)
        Dim PICK_NO_CONS As String = String.Empty

        Dim sqlSOTCART1 As String = "Select SOTCART1.*,SOTPICK1.SHIP_BOL_NO,SOTPICK1.ORDR_NO" & vbCrLf _
            & ", SUBSTR(SOTCART1.CART_NO,11,9) CART_NO_9" & vbCrLf _
            & ", SUBSTR(SOTCART1.CART_NO,20,1) CART_NO_DIGIT" & vbCrLf _
            & ", SUBSTR(SOTCART1.CART_NO,5,6) CART_NO_PFX" & vbCrLf _
            & ", '(00) 0 0 ' || SUBSTR(SOTCART1.CART_NO,5,6) || ' ' || SUBSTR(SOTCART1.CART_NO,11,9) || SUBSTR(SOTCART1.CART_NO,20,1) CART_NO_FMT" & vbCrLf

        ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
            & ", SOTORDR0.CUST_CODE" & vbCrLf _
            & ", DECODE (SOTSHIP1.SHIP_ADDR_TYPE,'DC',SOTSHIP1.SHIP_BOL_NO,'MK') SHIP_BOL_NO_X" & vbCrLf _
            & " from SOTSHIP1,SOTORDR0" & vbCrLf _
            & " where SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
            & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
        clsASCBASE1.Fill_Records("SOTSHIP1_WK", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select T1.PICK_NO, T1.ORDR_NO, T1.PICK_FREIGHT,T1.PICK_PICKER,T1.ORDR_PICK_SEQ,T1.PICK_STATUS,  " & vbCrLf _
            & "T1.PICK_RELEASED,T1.PICK_PRINTED,T1.PICK_PACKED,T1.PICK_SHIPPED,T1.PICK_BATCH_NO,T1.SHIP_BOL_NO,T1.INV_NO, " & vbCrLf _
            & "T1.PICK_CNT_CARTONS,T1.PICK_TOTAL_WGT, t1.INIT_OPER, t1.LAST_OPER, t1.INIT_DATE, t1.LAST_DATE, t1.PICK_PRINTED_OPER, " & vbCrLf _
            & "T1.PICK_NO_REV, t1.CCPA_NO, t1.SHIP_CNTL_NO, t1.CCPA_NO_STATUS, t1.CCPA_NO_AUTH, t1.CONFIG_NO, " & vbCrLf _
            & "'" & PICK_NO_CONS & "'AS PICK_NO_CONS " & vbCrLf _
            & ",SOTORDR1.CUST_STORE_NO from SOTPICK1 T1, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = T1.ORDR_NO" & vbCrLf _
            & "   And T1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        clsASCBASE1.Fill_Records("SOTPICK1_WK", "", True, ASCMAIN1.sql)

        'ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
        '    & " from SOTPICK1,SOTPICK2,SOTORDR2" & vbCrLf _
        '    & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
        '    & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
        '    & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
        '    & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCMAIN1.sql = pick2Query & " and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        clsASCBASE1.Fill_Records("SOTPICK2_WK", "", True, ASCMAIN1.sql)

        ' loop here to update locations
        clsASCBASE1.dst.Tables("SOTCART1_WK").Rows.Clear()
        clsASCBASE1.dst.Tables("SOTCART2_WK").Rows.Clear()

        ASCMAIN1.sql = "Select SOTORDR1.*, 'MK' AS MARK_FOR, 'ST' AS SHIP_TO" & vbCrLf _
           & " from SOTORDR1,SOTPICK1" & vbCrLf _
           & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
           & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        clsASCBASE1.Fill_Records("SOTORDR1_WK", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
           & " from SOTORDR2,SOTORDR1,SOTPICK1" & vbCrLf _
           & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
           & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
           & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        clsASCBASE1.Fill_Records("SOTORDR2_WK", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = sqlSOTCART1 & " from SOTCART1,SOTPICK1 where SOTPICK1.PICK_NO = SOTCART1.PICK_NO  and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        clsASCBASE1.Fill_Records("SOTCART1_WK", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
            & " from SOTCART2,SOTCART1,SOTPICK1" & vbCrLf _
            & " where SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTCART1.PICK_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"

        clsASCBASE1.Fill_Records("SOTCART2_WK", "", True, ASCMAIN1.sql)
        SetCartItemCode()

        clsASCBASE1.EnforceConstraints(True)

    End Sub

    Private Sub SetCartItemCode()

        Dim CART_NO As String = String.Empty
        Dim PICK_NO As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim ORDR_LNO As Int16 = 0
        Dim rowSOTCART2 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing

        For Each rowSOTCART1 As DataRow In clsASCBASE1.dst.Tables("SOTCART1_WK").Select("")
            CART_NO = rowSOTCART1.Item("CART_NO")
            PICK_NO = rowSOTCART1.Item("PICK_NO") & String.Empty

            Dim numItems As Int16 = ASCDATA1.SelectDistinct(clsASCBASE1.dst.Tables("SOTCART2_WK").Select("CART_NO = '" & CART_NO & "'"), "STYLE_CODE").Rows.Count

            rowSOTCART2 = clsASCBASE1.dst.Tables("SOTCART2_WK").Select("CART_NO = '" & CART_NO & "'")(0)
            ORDR_NO = rowSOTCART2.Item("ORDR_NO") & String.Empty
            ORDR_LNO = Val(rowSOTCART2.Item("ORDR_LNO") & String.Empty)

            rowSOTORDR2 = clsASCBASE1.dst.Tables("SOTORDR2_WK").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})

            If numItems = 1 Then
                If rowSOTORDR2 IsNot Nothing Then
                    rowSOTCART1.Item("STYLE_CODE") = rowSOTORDR2.Item("STYLE_CODE")
                    rowSOTCART1.Item("STYLE_DESC") = rowSOTORDR2.Item("STYLE_DESC")
                Else
                    rowSOTCART1.Item("STYLE_CODE") = "Could not resolve Style"
                End If
            Else
                rowSOTCART1.Item("STYLE_CODE") = "Mixed"
            End If

            ' DELETE FROM HERE --------------------------------

            If PICK_NO.Length > 0 Then
                If rowSOTORDR2 IsNot Nothing Then
                    If clsASCBASE1.dst.Tables("SOTPICK2_WK").Select("PICK_NO = '" & PICK_NO & "' and STYLE_CODE = '" & rowSOTORDR2.Item("STYLE_CODE") & "'").Length > 0 Then
                        Dim rowSOTPICK2 As DataRow = clsASCBASE1.dst.Tables("SOTPICK2_WK").Select("PICK_NO = '" & PICK_NO & "' and STYLE_CODE = '" & rowSOTORDR2.Item("STYLE_CODE") & "'")(0)
                        If Val(rowSOTPICK2.Item("CARTON_PACK_QTY") & String.Empty) > 0 Then
                            rowSOTCART1.Item("NUM_CARTONS") = Val(rowSOTCART1.Item("CART_TOTAL_UNITS") & String.Empty) / Val(rowSOTPICK2.Item("CARTON_PACK_QTY") & String.Empty)
                            rowSOTCART1.Item("TOTAL_WEIGHT") = Val(rowSOTCART1.Item("NUM_CARTONS") & String.Empty) * Val(rowSOTPICK2.Item("CASE_WEIGHT_GRS") & String.Empty)
                            rowSOTCART1.Item("CASE_CUBE_FT") = Val(rowSOTCART1.Item("NUM_CARTONS") & String.Empty) * (Val(rowSOTPICK2.Item("CASE_CUBE") & String.Empty) / 1728)
                        End If
                    End If

                End If
            End If

            ' DELETE TO HERE --------------------------------

            ' ED - THIS SECTION IS HOW I THINK THE ABOVE SECTION SHOULD HAVE BEEN WRITTEN


            Dim NUM_CARTONS As Decimal = 0
            Dim TOTAL_WEIGHT As Decimal = 0
            Dim CASE_CUBE_FT As Decimal = 0

            For Each rowSOTCART2 In clsASCBASE1.dst.Tables("SOTCART2_WK").Select("CART_NO = '" & CART_NO & "'")

                ORDR_NO = rowSOTCART2.Item("ORDR_NO")
                ORDR_LNO = Val(rowSOTCART2.Item("ORDR_LNO"))
                Dim QTY_PACKED As Int64 = Val(rowSOTCART2.Item("QTY_PACKED"))

                Dim rowSOTPICK2 As DataRow = clsASCBASE1.dst.Tables("SOTPICK2_WK").Select _
                    ("PICK_NO = '" & PICK_NO & "' and ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO))(0)

                'rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                'Dim CARTON_PACK_QTY As Int32 = Val(rowSOTORDR2.Item("CARTON_PACK_QTY") & "")

                Dim CARTON_PACK_QTY As Int32 = Val(rowSOTPICK2.Item("CARTON_PACK_QTY") & "")
                Dim CASE_WEIGHT_GRS As Decimal = Val(rowSOTPICK2.Item("CASE_WEIGHT_GRS") & "")
                Dim CASE_CUBE As Decimal = Val(rowSOTPICK2.Item("CASE_CUBE") & "")
                If CARTON_PACK_QTY = 0 Then CARTON_PACK_QTY = 1
                Dim CARTONS As Decimal = QTY_PACKED / CARTON_PACK_QTY
                NUM_CARTONS += CARTONS
                TOTAL_WEIGHT += CARTONS * System.Math.Round(CASE_WEIGHT_GRS + 0.51, 0)
                CASE_CUBE_FT += CARTONS * (CASE_CUBE * 1.1)
            Next

            rowSOTCART1.Item("NUM_CARTONS") = NUM_CARTONS
            rowSOTCART1.Item("TOTAL_WEIGHT") = TOTAL_WEIGHT
            rowSOTCART1.Item("CASE_CUBE_FT") = CASE_CUBE_FT / 1728
        Next

    End Sub

End Class
