Public Class EDROUTB2

    ' A SOTSHIP1 RECORD THAT HAS NO CANDIATE INVOICES WILL FOREVER COME UP AS A CANDIDATE IN SOTSHIPX - NEED TO MOTHBALL THESE SOMEEOW
    ' PERHAPS IN THE UPDATE WE CAN SET REGISTERXNO OR BATCH NO TO SOME NON NULL VALUE FOR SELECT SHIP_BOL_NO FROM SOTSHIPX MINUS SELECT DISTINCT SHIP_BOL_NO FROM SOTINVHX

    '   SHIPSTATUS = 'F' FROM SOTSHIP1 FOR 856'S
    '   ADR recs in 856's - Burdines, Meijers, bloomies, Montgom, goodys, boscovs, nordstrom, richs, WALMART, BON MARCHE, MACYSWEST

    ' before turning this on for all custoemrs, need to scan all code looking for CUST_CODE to make sure we are using the variable vs rowSOTORDR1.item("CUST_CODE") properly

    Dim rowSOTINVHC As DataRow
    Dim rowARTCUST1 As DataRow
    Dim rowEDTTRPM1_810 As DataRow
    Dim rowEDTTRPM1_856 As DataRow
    Dim rowEDT850T1 As DataRow
    Dim rowSOTINVH1 As DataRow
    Dim rowSOTORDR1 As DataRow
    Dim rowSOTSHIP1 As DataRow
    Dim rowSOTPICK1 As DataRow
    Dim rowTATTERM1 As DataRow
    Dim rowSOTSVIA1 As DataRow
    Dim rowICTWHSE1 As DataRow

    Dim SHIP_8XX_BATCH_NO As String
    Dim EDI_CUSTOMER As String
    Dim CUST_CODE As String
    Dim ORDR_SOURCE As String
    Dim SHIP_BOL_NO As String
    Dim SHIP_BOL_NO_CONS As String
    Dim BILL_OF_LADING_NO As String

    Dim EDI_DOC_SEQ_NO As String
    Dim EDI_MERCH_TYPE As String

    Dim EDI_OUR_ID As String
    Dim EDI_OUR_QUAL As String
    Dim EDI_STYLE As String

    Dim SHIP_810_BATCH_NO As String
    Dim SHIP_856_BATCH_NO As String
    Dim BILLING_INV_NO As String

    Dim first_invoice As Boolean

    Dim GTIN As String
    Dim UPC As String
    Dim SKU As String

    Dim PROCESS_810 As String
    Dim PROCESS_856 As String

    Dim HDR_TOT_CHK As Double
    Dim DTL_TOT_CHK As Double
    Dim EDI_ADR_SEQ As Long
    Dim EDI_OUTBOUND_DOC_NO As String

    Dim INV_ERR As String
    Dim INV_NO_CONS As String
    Dim CONS_INV As Boolean

    Dim CUST_NAME As String
    Dim CUST_ADDR1 As String
    Dim CUST_ADDR2 As String
    Dim CUST_CITY As String
    Dim CUST_STATE As String
    Dim CUST_ZIP_CODE As String
    Dim CUST_COUNTRY As String
    Dim CUST_GLN As String
    Dim CUST_DC_NO As String
    Dim CUST_ADDR_CODE_DC As String
    Dim CUST_NAME_DC As String
    Dim CUST_ADDR1_DC As String
    Dim CUST_ADDR2_DC As String
    Dim CUST_CITY_DC As String
    Dim CUST_STATE_DC As String
    Dim CUST_ZIP_CODE_DC As String
    Dim CUST_COUNTRY_DC As String
    Dim CUST_GLN_DC As String
    Dim CUST_STORE_NO As String
    Dim CUST_ADDR_CODE_MK As String
    Dim CUST_NAME_MK As String
    Dim CUST_ADDR1_MK As String
    Dim CUST_ADDR2_MK As String
    Dim CUST_CITY_MK As String
    Dim CUST_STATE_MK As String
    Dim CUST_ZIP_CODE_MK As String
    Dim CUST_COUNTRY_MK As String
    Dim CUST_GLN_MK As String

    Dim EDI_HL2_SEQ As Integer

    Dim BYPASS_ZERO_QTY_IN_ASN As String = ""
    Dim SEND_PARENT_STYLE_IN_ASN As String = ""

    Dim HL4_STYLES As New List(Of String)

    Dim EDI_TABLES() As String = {"EDTSYSIH", "EDTREJD1", _
             "EDT810O1", "EDT810O2", "EDT810O5", "EDT810O6", "EDT810O7", "EDT810O8", _
             "EDT856O1", "EDT856O2", "EDT856O3", "EDT856O4", "EDT856O5"}

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("EDTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Public Overrides Sub Print_Report()

        SUBT = "Sending: "
        If Absx1.chkFor("CHK856").Checked Then
            SUBT &= " 856s"
        End If
        If Absx1.chkFor("CHK810").Checked Then
            SUBT &= " 810s"
        End If
        If Absx1.chkFor("CHKRECREATE").Checked Then
            SUBT &= " (Retransmission)"
        Else
            SUBT &= " (Original Transmission)"
        End If

        Generate_Report(RPT, , SUBT)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.chkFor("CHK810").Checked And Absx1.chkFor("CHK856").Checked Then
                EMsg &= vbCr & "You Cannot run 810s and 856s at the Same Time" ' although I do not know why
            Else
                If Not Absx1.chkFor("CHK810").Checked And Not Absx1.chkFor("CHK856").Checked Then
                    EMsg &= vbCr & "You Must Select Process 810s or Process 856s"
                End If
            End If

            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("INV_NO")
            If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                If Not Absx1.chkFor("CHKRECREATE").Checked Then
                    EMsg &= vbCr & "You may not use Selected Invoices unless you are Re-Creating a Batch"
                End If
            End If

            If tblASTDSQLA.Select("EXCLUDE = '1'").Length <> 0 Then
                EMsg &= vbCr & "You may not use Exclusion on any Filter for Outbound EDI"
            End If



            ' UNTIL WE OPEN THIS UP TO ALL CUSTOMERS - ALSO SEE BUILD WORKFILE

            If Absx1.chkFor("CHKRECREATE").Checked Then
                EMsg &= vbCr & "You may not use Re-Create Batch (see Rick or Walter)"
            End If

            'rowASTDSQLA = tblASTDSQLA.Rows.Find("INV_NO")
            'If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
            '    EMsg &= vbCr & "You may not use Selected Invoices (Walmart Multi PO Shipments Only)"
            'End If

            'rowASTDSQLA = tblASTDSQLA.Rows.Find("SHIP_BOL_NO")
            'If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
            '    EMsg &= vbCr & "You may not use Selected Shipments (Walmart Multi PO Shipments Only)"
            'End If

            'rowASTDSQLA = tblASTDSQLA.Rows.Find("CUST_CODE")
            'If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
            '    EMsg &= vbCr & "You may not use Selected Customers (Walmart Multi PO Shipments Only)"
            'End If

        End If

    End Sub

    Overrides Sub Update_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Create Temp Table containing all SOTSHIP1 records to be queued for Outbound

        Dim sqlw As String = ""

        If Absx1.chkFor("CHKRECREATE").Checked Then
            sqlw = ""
            If Absx1.chkFor("CHK810").Checked Then
                sqlw = "SHIP_810_BATCH_NO = '" & Absx1.txtFor("TXTRECREATE").Text & "' "
                If Absx1.chkFor("CHK856").Checked Then
                    sqlw = " and (" & sqlw
                Else
                    sqlw = " and " & sqlw
                End If
            End If
            If Absx1.chkFor("CHK856").Checked Then
                If Absx1.chkFor("CHK810").Checked Then
                    sqlw = sqlw & " or "
                Else
                    sqlw = sqlw & " and "
                End If
                sqlw = sqlw & " SHIP_856_BATCH_NO = '" & Absx1.txtFor("TXTRECREATE").Text & "' "
                If Absx1.chkFor("CHK810").Checked Then
                    sqlw = sqlw & ")"
                End If
            End If
        Else
            sqlw = ""
            If Absx1.chkFor("CHK810").Checked Then
                sqlw = " and (SOTSHIP1.SHIP_810_BATCH_NO IS NULL"
            End If
            If Absx1.chkFor("CHK856").Checked Then
                If Absx1.chkFor("CHK810").Checked Then
                    sqlw = sqlw & " or "
                Else
                    sqlw = " and ("
                End If
                sqlw = sqlw & " SOTSHIP1.SHIP_856_BATCH_NO IS NULL"
            End If
            sqlw = sqlw & ") and SOTSHIP1.REGISTER_XNO is not Null"
        End If

        'If Absx1.chkFor("CHKSKB").Checked Then
        sqlw = sqlw & " and SOTORDR0.ORDR_SOURCE = 'E'"
        'End If

        sqlw &= SQL_in("CUST_CODE", "SOTORDR0.CUST_CODE")
        sqlw &= SQL_in("SHIP_BOL_NO", "SOTSHIP1.SHIP_BOL_NO")

        sqlw = sqlw & " and SOTSHIP1.SHIP_VIA_CODE <> 'UNKO'"

        ' UNTIL WE OPEN THIS UP TO ALL CUSTOMERS - ALSO SEE VERIFY_SPECIAL
        'sqlw &= " and SOTORDR0.CUST_CODE = 'WALMART' and SOTSHIP1.SHIP_BOL_NO_CONS is Not Null"

        ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & ", SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_BOL_NO_CONS, SOTSHIP1.MASTER_SHIP_BOL_NO, SOTSHIP1.EDI_LOAD_ID, SOTSHIP1.SHIP_REF" & vbCrLf _
            & ", SOTORDR0.CUST_CODE, SOTSHIP1.SHIP_ADDR_TYPE, SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
            & " from SOTSHIP1,SOTORDR0 " & vbCrLf _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & "   and SHIP_BOL_NO_REV is Null" & vbCrLf _
            & sqlw

        ASCMAIN1.sql = "Select X.*, ARTCUST2.CUST_ADDR_CODE from (" & vbCrLf _
            & ASCMAIN1.sql & vbCrLf _
            & ") X, ARTCUST2" & vbCrLf _
            & " where ARTCUST2.CUST_CODE (+) = X.CUST_CODE" _
            & "   and ARTCUST2.CUST_ADDR_TYPE (+) = X.SHIP_ADDR_TYPE" _
            & "   and ARTCUST2.CUST_ADDR_CODE (+) = X.SHIP_ADDR_CODE"

        Dim SOTSHIPX As String = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")

        ' Create dataset tables

        With dst
            For Each TABLE_NAME As String In EDI_TABLES
                Create_TDA(.Tables.Add, TABLE_NAME, "*")
            Next
        End With

        ASCMAIN1.sql = "Select * from " & SOTSHIPX
        Create_TDA(dst.Tables.Add, "SOTSHIPX", "**", 0, False, , 1)
        Fill_Records("SOTSHIPX")


        If dst.Tables("SOTSHIPX").Select("CUST_ADDR_CODE is Null").Length <> 0 Then
            Dim row As DataRow = dst.Tables("SOTSHIPX").Select("CUST_ADDR_CODE is Null")(0)
            xErrMsg = "Ship BOL " & row.Item("SHIP_BOL_NO") & " has an Issue joining Shipping Address record to ARTCUST2"
            MsgBox(xErrMsg, MsgBoxStyle.OkOnly, "Cannot Proceed with Outbound EDI")
            RWU &= "0"
            Exit Sub
        End If

        ' MAKE SURE THAT BILL_OF_LADING_NO, MIN () MAX () CUSTS, SHIP_ADDR_CODES,  SHIP_REF, LOAD_ID, master_ship_bol_no  MUST ALL  MATCH
        ASCMAIN1.sql = "Select SHIP_BOL_NO_CONS, Count (*) SHIP_BOL_NOS" & vbCrLf _
            & ", Min (CUST_CODE) CUST_CODE_MIN, Max (CUST_CODE) CUST_CODE_MAX" & vbCrLf _
            & ", Min (SHIP_ADDR_CODE) SHIP_ADDR_CODE_MIN, Max (SHIP_ADDR_CODE) SHIP_ADDR_CODE_MAX" & vbCrLf _
            & ", Min (SHIP_REF) SHIP_REF_MIN, Max (SHIP_REF) SHIP_REF_MAX" & vbCrLf _
            & ", Min (EDI_LOAD_ID) EDI_LOAD_ID_MIN, Max (EDI_LOAD_ID) EDI_LOAD_ID_MAX" & vbCrLf _
            & ", Min (MASTER_SHIP_BOL_NO) MASTER_SHIP_BOL_NO_MIN, Max (MASTER_SHIP_BOL_NO) MASTER_SHIP_BOL_NO_MAX" & vbCrLf _
            & " from " & SOTSHIPX & " where SHIP_BOL_NO_CONS is Not Null" & vbCrLf _
            & " group by SHIP_BOL_NO_CONS"
        ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & ASCMAIN1.sql & ")" & vbCrLf _
            & " where CUST_CODE_MIN <> CUST_CODE_MAX" & vbCrLf _
            & "    or SHIP_ADDR_CODE_MIN <> SHIP_ADDR_CODE_MAX" & vbCrLf _
            & "    or SHIP_REF_MIN <> SHIP_REF_MAX" & vbCrLf _
            & "    or EDI_LOAD_ID_MIN <> EDI_LOAD_ID_MAX" & vbCrLf _
            & "    or MASTER_SHIP_BOL_NO_MIN <> MASTER_SHIP_BOL_NO_MAX"

        Dim rowsNotMatching() As DataRow = ASCDATA1.GetDataTable.Select("")
        If rowsNotMatching.Length <> 0 Then
            Dim row As DataRow = rowsNotMatching(0)
            xErrMsg = "Ship BOL " & row.Item("SHIP_BOL_NO_CONS") & " has an Issue with Non-Matching Data"
            MsgBox(xErrMsg, MsgBoxStyle.OkOnly, "Cannot Proceed with Outbound EDI")
            RWU &= "0"
            Exit Sub
        End If

         



        ' For Consolidated SHIP_BOL_NOs, we must select all in a consolidation

        Dim SHIP_BOL_NO_CONSs As New Dictionary(Of String, List(Of String))
        For Each row As DataRow In dst.Tables("SOTSHIPX").Select("SHIP_BOL_NO_CONS is Not Null")
            Dim SHIP_BOL_NO_CONS As String = row.Item("SHIP_BOL_NO_CONS")
            If Not SHIP_BOL_NO_CONSs.ContainsKey(SHIP_BOL_NO_CONS) Then
                Dim SHIP_BOL_NOs As New List(Of String)

                ASCMAIN1.sql = "Select SHIP_BOL_NO from SOTSHIP1 where SHIP_BOL_NO_CONS = '" & SHIP_BOL_NO_CONS & "' and SHIP_STATUS <> 'D'"
                For Each row2 As DataRow In ASCDATA1.GetDataTable().Select("")
                    Dim SHIP_BOL_NO As String = row2.Item("SHIP_BOL_NO")
                    If dst.Tables("SOTSHIPX").Rows.Find(SHIP_BOL_NO) Is Nothing Then
                        xErrMsg = "Ship BOL " & SHIP_BOL_NO & " not included in Outbound Batch and must be included with " & SHIP_BOL_NO_CONS
                        MsgBox(xErrMsg, MsgBoxStyle.OkOnly, "Cannot Proceed with Outbound EDI")
                        '  Throw New Exception(EMsg)
                        RWU &= "0"
                        Exit Sub
                    End If

                    ' For Consolidated PICK_NOs, only 1 carton per PT

                    ASCMAIN1.sql = "Select SOTCART1.PICK_NO, Count (*) CTNS" & vbCrLf _
                        & " from SOTPICK1,SOTCART1" & vbCrLf _
                        & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                        & "   and SOTPICK1.PICK_STATUS <> 'D' and SOTPICK1.PICK_STATUS <> 'P'" & vbCrLf _
                        & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" & vbCrLf _
                        & " group by SOTCART1.PICK_NO having Count (*) > 1"
                    Dim rowSOTPICK1s() As DataRow = ASCDATA1.GetDataTable().Select()
                    If rowSOTPICK1s.Length <> 0 Then
                        Dim rowPs As DataRow = rowSOTPICK1s(0)
                        xErrMsg = "Ship BOL " & SHIP_BOL_NO & " has a PT (" & rowPs.Item("PICK_NO") & ") with Multiple Cartons"
                        MsgBox(xErrMsg, MsgBoxStyle.OkOnly, "Cannot Proceed with Outbound EDI")
                        RWU &= "0"
                        Exit Sub
                    End If

                    SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                Next

                SHIP_BOL_NO_CONSs.Add(SHIP_BOL_NO_CONS, SHIP_BOL_NOs)

            End If
        Next



        ' WTF ON EDI_RETRANSMIT_IND - WANT TO GET RID OF THIS , BECAUSE IT IS NOT GOING TO WORK FOR CONS INVOICES UNLESS YOU ENSURE ALL CONS INVOICES ARE FLAGGED, AND ALSO THAT THE SHIPMENT BATCH FIELDS ARE NULLED OUT
        ' SHOULDNT THIS BE INVPOCES ONLY AND NOT ASN?
        ' NOTE THAT SOTSHIPX DOES NOT MAKE SURE TO GRAB THE SHIPMENTS FOR INVOICES FLAGGED FOR RETRANSMIT

        ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
            & " from SOTINVH1, " & SOTSHIPX & " W" & vbCrLf _
            & " where (SOTINVH1.SHIP_BOL_NO = W.SHIP_BOL_NO or SOTINVH1.EDI_RETRANSMIT_IND = '1')" & vbCrLf _
            & " and SOTINVH1.INV_NO_REV is Null and SOTINVH1.INV_NO_REV_BY is Null "
        Dim SOTINVHX As String = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHX & " Add Primary Key (INV_TYPE, INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVHX & "_1 on " & SOTINVHX & " (INV_NO_CONS)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVHX & "_2 on " & SOTINVHX & " (ORDR_NO)")

        Dim sqlINVHX As String = " where SOTINVH1.SHIP_BOL_NO = :PARM1" 

        ASCMAIN1.sql = "Select SOTINVH1.* from " & SOTINVHX & " SOTINVH1" & sqlINVHX
        Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, False, "V", 2)

        ASCMAIN1.sql = "Select SOTINVH2.* from " & SOTINVHX & " SOTINVH1,SOTINVH2" & vbCrLf _
            & sqlINVHX & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO"
        Create_TDA(dst.Tables.Add, "SOTINVH2", "**", 0, False, "V", 3)

        ASCMAIN1.sql = "Select SOTINVH9.* from " & SOTINVHX & " SOTINVH1,SOTINVH9" & vbCrLf _
            & sqlINVHX & vbCrLf _
            & "   and SOTINVH9.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH9.INV_NO = SOTINVH1.INV_NO"
        Create_TDA(dst.Tables.Add, "SOTINVH9", "**", 0, False, "V", 3)


        'ASCMAIN1.sql = "Select NVL(SOTINVH1.INV_NO_CONS, SOTINVH1.INV_NO) BILLING_INV_NO, SOTINVH1.* " & vbCrLf _
        '    & " from SOTINVH1, " & SOTINVH1_TEMP & " T" & vbCrLf _
        '    & " where SOTINVH1.INV_TYPE = 'I' " & vbCrLf _
        '    & "   and SOTINVH1.INV_NO  = T.INV_NO "
        'Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, False, "", 3)
        'Fill_Records("SOTINVH1")

        ASCMAIN1.sql = "Select SOTSHIP1.* from SOTSHIP1 where SHIP_BOL_NO IN (" & vbCrLf _
            & " Select Distinct SHIP_BOL_NO from " & SOTINVHX & ")"
        Create_TDA(dst.Tables.Add, "SOTSHIP1", "**", 0, False, "", 1)
        Fill_Records("SOTSHIP1")

        ASCMAIN1.sql = "Select SOTSHIP1.BILL_OF_LADING_NO" & vbCrLf _
            & ", Sum(SOTPICK1.PICK_TOTAL_WGT) TOT_WGT, Sum(SOTPICK1.PICK_CNT_CARTONS) TOT_CTNS " & vbCrLf _
            & ", Sum(CASE WHEN SOTPICK1.PICK_NO = SOTPICK1.PICK_NO_CONS THEN SOTPICK1.PICK_CNT_CARTONS ELSE 0 END) TOT_CTNS_CONS" & vbCrLf _
            & " from SOTSHIP1, SOTORDR0, SOTPICK1, " & SOTSHIPX & " SOTSHIPX" & vbCrLf _
            & " where SOTSHIP1.BILL_OF_LADING_NO = :PARM1" & vbCrLf _
            & "   and SOTORDR0.CUST_CODE = :PARM2" & vbCrLf _
            & "   and SOTSHIP1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO " & vbCrLf _
            & "   and SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & " group by SOTSHIP1.BILL_OF_LADING_NO"
        Create_TDA(dst.Tables.Add, "SOTSHIPT", "**", 0, False, "VV", 1)

        ASCMAIN1.sql = "Select INV_NO_CONS, SUM(PICK_CNT_CARTONS) PICK_CNT_CARTONS, SUM(PICK_TOTAL_WGT) PICK_TOTAL_WGT, SUM(INV_SALES) INV_SALES," & vbCrLf _
            & " SUM(INV_FREIGHT) INV_FREIGHT, SUM(INV_TOTAL_AMOUNT) INV_TOTAL_AMOUNT, SUM(INV_MISC_CHG) INV_MISC_CHG, SUM(INV_TOTAL_AMT_CURR) INV_TOTAL_AMT_CURR, " & vbCrLf _
            & " SUM(INV_SALES_CURR) INV_SALES_CURR, SUM(GST_TAX) GST_TAX, SUM(GST_TAX_CURR) GST_TAX_CURR" & vbCrLf _
            & " from " & SOTINVHX & " SOTINVH1, SOTPICK1" & vbCrLf _
            & " where INV_NO_CONS = :PARM1" & vbCrLf _
            & "   and SOTPICK1.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & " group by INV_NO_CONS"
        Create_TDA(dst.Tables.Add, "SOTINVHC", "**", 0, False, "V", 1)

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  MIN(RANGE_STYLE_PP_PRICE_CURR) RANGE_STYLE_PP_PRICE_CURR" & vbCrLf _
            & ", MIN(RANGE_STYLE_PP_PRICE) RANGE_STYLE_PP_PRICE" & vbCrLf _
            & ", SUM(RANGE_STYLE_PP_QTY_SHIP) RANGE_STYLE_PP_QTY_SHIP" & vbCrLf _
            & " from SOTINVH9, " & SOTINVHX & " SOTINVHX" & vbCrLf _
            & " where SOTINVH9.INV_TYPE = 'I' AND SOTINVH9.INV_NO = SOTINVHX.INV_NO " & vbCrLf _
            & "   and SOTINVHX.INV_NO_CONS = :PARM1 and SOTINVH9.RANGE_STYLE_CODE = :PARM2"
        Create_TDA(dst.Tables.Add, "SOTINVH9R", "**", 0, False, "VV", 0)

        Dim sqlSOTPICK1 As String = "" _
            & " where SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
            & "   and SOTPICK1.PICK_NO_REV is Null" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS <> 'D' and SOTPICK1.PICK_STATUS <> 'P'"
        ' MAYBE WE SHOULD GET PICK_STATUS = F ONLY - STATUS ARE P C D F

        ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1" & vbCrLf _
            & sqlSOTPICK1
        Create_TDA(dst.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

        ASCMAIN1.sql = "Select SOTPICK2.* FROM SOTPICK2, SOTPICK1" & vbCrLf _
            & sqlSOTPICK1 & vbCrLf _
            & "  and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO"
        Create_TDA(dst.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

        ASCMAIN1.sql = "Select SOTCART1.* from SOTCART1, SOTPICK1" & vbCrLf _
            & sqlSOTPICK1 & vbCrLf _
            & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
        Create_TDA(dst.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

        ASCMAIN1.sql = "Select SOTPICK1.PICK_NO_CONS, SOTCART1.CART_NO" & vbCrLf _
            & " from SOTCART1, SOTPICK1, " & SOTSHIPX & " SOTSHIPX" & vbCrLf _
            & " where SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO_CONS is Not Null" & vbCrLf _
            & "   and SOTPICK1.PICK_NO_CONS = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTPICK1.PICK_STATUS <> 'D' and SOTPICK1.PICK_STATUS <> 'P'" & vbCrLf _
            & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
        Create_TDA(dst.Tables.Add, "SOTCARTP", "**", 0, False, , 1)
        Fill_Records("SOTCARTP")

        ASCMAIN1.sql = "Select SOTCART2.* from SOTCART2, SOTCART1, SOTPICK1" & vbCrLf _
            & sqlSOTPICK1 & vbCrLf _
            & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
            & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
        Create_TDA(dst.Tables.Add, "SOTCART2", "**", 0, False, "V", 2)

        ASCMAIN1.sql = "Select Distinct SOTINVH1.ORDR_NO, SOTORDR1.EDI_DOC_SEQ_NO, SOTINVH1.SHIP_BOL_NO" & vbCrLf _
            & " from SOTORDR1, " & SOTINVHX & " SOTINVH1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO"
        Dim SOTORDRX As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRX & " Add Primary Key (ORDR_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTORDRX & "_1 on " & SOTORDRX & " (EDI_DOC_SEQ_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTORDRX & "_2 on " & SOTORDRX & " (SHIP_BOL_NO)")

         For Each TABLE As String In New String() {"SOTORDR1", "SOTORDR2", "SOTORDR9"}
            ASCMAIN1.sql = "Select T.* from " & TABLE & " T, " & SOTORDRX & " X" & vbCrLf _
                & " where T.ORDR_NO =  X.ORDR_NO" & vbCrLf _
                & "   and X.SHIP_BOL_NO = :PARM1"
            Dim kc As Integer = 2
            If TABLE = "SOTORDR1" Then kc = 1
            Create_TDA(dst.Tables.Add, TABLE, "**", 0, False, "V", kc)
        Next

        ASCMAIN1.sql = "Select Distinct EDI_DOC_SEQ_NO, SHIP_BOL_NO from " & SOTORDRX
        Dim EDT850TX As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & EDT850TX & " Add Primary Key (EDI_DOC_SEQ_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & EDT850TX & "_2 on " & EDT850TX & " (SHIP_BOL_NO)")

        For Each TABLE As String In New String() {"EDT850T1", "EDT850T2", "EDT850T6", "EDT850T7", "EDT850T8"}
            ASCMAIN1.sql = "Select T.* from " & TABLE & " T, " & EDT850TX & " X" & vbCrLf _
                & " where T.EDI_DOC_SEQ_NO =  X.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and X.SHIP_BOL_NO = :PARM1"
            Dim kc As Integer = 2
            If TABLE = "EDT850T1" Then kc = 1
            If TABLE = "EDT850T6" Or TABLE = "EDT850T8" Then kc = 3
            Create_TDA(dst.Tables.Add, TABLE, "**", 0, False, "V", kc)
        Next

        Create_TDA(dst.Tables.Add, "SOTCSTY1", "*", 1, False)

        With dst.Tables.Add("SOTCSTYX")
            .Columns.Add("CUST_CODE")
            .Columns.Add("STYLE_CODE")
            .Columns.Add("COLOR_CODE")
            .Columns.Add("UPC")
            .Columns.Add("SKU")
            .PrimaryKey = New DataColumn() {.Columns("CUST_CODE"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
        End With

        Create_TDA(dst.Tables.Add, "ICTSTYC2", "*", 0, False)
        Fill_Records("ICTSTYC2")
        Create_TDA(dst.Tables.Add, "ICTSTYC4", "*", 0, False)
        Fill_Records("ICTSTYC4")

        Create_TDA(dst.Tables.Add, "ICTSTYL1", "*", , False)

        ASCMAIN1.sql = "Select M1.*, CUST_CODE, NUMBER_CHRS_SHIPTO, NUMBER_CHRS_DC" & vbCrLf _
            & ", BYPASS_ZERO_QTY_IN_ASN, NEW_SERVER_FLAG, DEFAULT_ADDR_QUALIFIER, EDI_OUR_ID" & vbCrLf _
            & ", EDI_OUR_QUAL, EDI_DUNS_PLUS4_PREFIX, P1.SEND_PARENT_STYLE_IN_ASN" & vbCrLf _
            & " from EDTTRPMZ M1, EDTSLSP1 P1 where EDI_DOC_NO in ('810','856')" & vbCrLf _
            & " and M1.EDI_TP_QUAL = P1.EDI_TP_QUAL and M1.EDI_TP_ID = P1.EDI_TP_ID "
        Create_TDA(dst.Tables.Add, "EDTTRPM1", "**", 0, False, "", 3)
        Fill_Records("EDTTRPM1")

        For Each TABLE_NAME As String In New String() {"SOTSVIA1", "TATTERM1", "ICTWHSE1"}
            Create_TDA(dst.Tables.Add, TABLE_NAME, "*", 0, False)
            Fill_Records(TABLE_NAME)
        Next

        ASCMAIN1.sql = "Select * from ICTRSTY2"
        Create_TDA(dst.Tables.Add, "ICTRSTY2", "**", 0, False, "", 2)

        ASCMAIN1.sql = "Select * from ARTCUST1"
        Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 1, False, "", 1)

        ASCMAIN1.sql = "Select * from ARTCUST2" '  where CUST_CODE = :CUST_CODE and CUST_ADDR_TYPE = :CODE1 AND CUST_ADDR_CODE = :CODE2"
        Create_TDA(dst.Tables.Add, "ARTCUST2", "**", 1, False, "", 3)

        ASCMAIN1.sql = "Select SHIP_810_BATCH_NO SHIP_8XX_BATCH_NO, SOTINVH1.CUST_CODE, CUST_DC_NO, BILL_OF_LADING_NO, CUST_NAME, SOTSHIP1.SHIP_BOL_NO AS FIRST_SBOL_NO," & vbCrLf _
            & " SOTSHIP1.SHIP_BOL_NO AS LAST_SBOL_NO, INV_TYPE || INV_NO FIRST_INV_NO, INV_TYPE || INV_NO LAST_INV_NO," & vbCrLf _
            & " 0 NUM_OF_810, 0 NUM_OF_856, 0.01 DOLLAR_AMT, SHIP_CNT_CARTONS, SHIP_TOTAL_WGT,  SOTSHIP1.SHIP_VIA_CODE," & vbCrLf _
            & " LPAD(' ',80) ERRORS From SOTSHIP1, SOTINVH1, SOTORDR1 WHERE ROWNUM < 1"
        Create_TDA(dst.Tables.Add, "EDTOUTB1", "**", 0, False, "", 4)

        ASCMAIN1.sql = "Select * from EDTCSCD1"
        Create_TDA(dst.Tables.Add, "EDTCSCD1", "**", 1, False, "", 2)

        If Not Absx1.chkFor("CHKRECREATE").Checked Then
            SHIP_8XX_BATCH_NO = ASCMAIN1.Next_Control_No("EDFOUTB1")
        Else
            SHIP_8XX_BATCH_NO = Absx1.txtFor("TXTRECREATE").Text
        End If

        If Absx1.chkFor("CHK810").Checked Then
            SHIP_810_BATCH_NO = SHIP_8XX_BATCH_NO
        Else
            SHIP_810_BATCH_NO = ""
        End If
        If Absx1.chkFor("CHK856").Checked Then
            SHIP_856_BATCH_NO = SHIP_8XX_BATCH_NO
        Else
            SHIP_856_BATCH_NO = ""
        End If

        Create_TDA(dst.Tables.Add, "ICTGTINT", "*", 0, False, "", 2)

        ASCMAIN1.Progress("Processing EDI", "")

        BeginTrans()

        Dim sqlSOTINVH1_RECREATE As String = SQL_in("INV_NO")

        For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("", "CUST_CODE, BILL_OF_LADING_NO, SHIP_BOL_NO")

            SHIP_BOL_NO = rowSOTSHIPX.Item("SHIP_BOL_NO")
            SHIP_BOL_NO_CONS = rowSOTSHIPX.Item("SHIP_BOL_NO_CONS")

            Setup_SHIP_BOL_NO()

            PROCESS_810 = IIf(Absx1.chkFor("CHK810").Checked, "1", "0")
            PROCESS_856 = IIf(Absx1.chkFor("CHK856").Checked, "1", "0")

            Dim sqlSOTINVH1 As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            If Absx1.chkFor("CHKRECREATE").Checked And sqlSOTINVH1_RECREATE <> "" Then
                sqlSOTINVH1 &= sqlSOTINVH1_RECREATE
            End If
            ' PROBLEM WITH RECREATING A BATCH - FOR 810 - IF THERE ARE CONSOLIDATED INVOICES AND YOU DON'T PICK ALL OF THEM, YOU ARE IN DEEP DOODOO

            INV_ERR = ""
            first_invoice = True
            For Each rowSOTINVH1 In dst.Tables("SOTINVH1").Select(sqlSOTINVH1)
                Process_SOTINVH1s()
            Next

            Dim sqlSOTSHIP1 As String = "Update SOTSHIP1 set "
            If Absx1.chkFor("CHK810").Checked Then
                sqlSOTSHIP1 &= " SHIP_810_BATCH_NO = '" & SHIP_8XX_BATCH_NO & "'"
                ASCDATA1.ExecuteSQL("Update SOTINVH1 set GEN_IND = '0', GEN_XNO = NULL, GEN_DATE = NULL, DOCUMENTKEY = NULL where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
            End If
            If Absx1.chkFor("CHK810").Checked And Absx1.chkFor("CHK856").Checked Then
                sqlSOTSHIP1 &= ","
            End If
            If Absx1.chkFor("CHK856").Checked Then
                sqlSOTSHIP1 &= " SHIP_856_BATCH_NO = '" & SHIP_8XX_BATCH_NO & "'"
                sqlSOTSHIP1 &= ", GEN_IND = '0', GEN_XNO = NULL, GEN_DATE = NULL, DOCUMENTKEY = NULL"
            End If
            sqlSOTSHIP1 &= " Where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            ASCDATA1.ExecuteSQL(sqlSOTSHIP1)
        Next

        'ASCMAIN1.Progress("Invoice Retransmission", "")
        ''INV_NO_CONS = ""
        ''CONS_INV = False

        'If Absx1.chkFor("CHK810").Checked Then
        '    PROCESS_810 = IIf(Absx1.chkFor("CHK810").Checked, "1", "0")

        '    ' RETRANSMITTED_INVOICES = ""
        '    EDI_ADR_SEQ = 0

        '    Dim sqlSOTINVH1 As String = "EDI_RETRANSMIT_IND = '1'"
        '    For Each rowSOTINVH1 In dst.Tables("SOTINVH1").Select(sqlSOTINVH1, "CUST_CODE, SHIP_BOL_NO, INV_NO_CONS, INV_TYPE, INV_NO")

        '        SHIP_BOL_NO = rowSOTINVH1.Item("SHIP_BOL_NO")
        '        Setup_SHIP_BOL_NO()

        '        Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE")
        '        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
        '        ASCMAIN1.Progress("-", INV_NO & " - 0")

        '        Process_SOTINVH1s()
        '        '   RETRANSMITTED_INVOICES = RETRANSMITTED_INVOICES & "," & INV_NO

        '        sql = "Update SOTINVH1 set EDI_RETRANSMIT_IND = '0' WHERE INV_NO = '" & INV_NO & "' AND INV_TYPE = '" & INV_TYPE & "'"
        '        ASCDATA1.ExecuteSQL(sql)
        '    Next
        'End If

        ' Update Oracle

        For Each TABLE_NAME As String In EDI_TABLES
            Update_Record_TDA(TABLE_NAME)
        Next

        CommitTrans()
    End Sub

    Sub Setup_SHIP_BOL_NO()
        ASCMAIN1.Progress("-", SHIP_BOL_NO)
        Fill_Records("SOTCART1", SHIP_BOL_NO)
        Fill_Records("SOTCART2", SHIP_BOL_NO)
        Fill_Records("SOTPICK1", SHIP_BOL_NO)
        Fill_Records("SOTPICK2", SHIP_BOL_NO)
        Fill_Records("SOTINVH1", SHIP_BOL_NO)
        Fill_Records("SOTINVH2", SHIP_BOL_NO)
        Fill_Records("SOTINVH9", SHIP_BOL_NO)

        Fill_Records("SOTORDR1", SHIP_BOL_NO)
        Fill_Records("SOTORDR2", SHIP_BOL_NO)
        Fill_Records("SOTORDR9", SHIP_BOL_NO)

        Fill_Records("EDT850T1", SHIP_BOL_NO)
        Fill_Records("EDT850T2", SHIP_BOL_NO)
        Fill_Records("EDT850T6", SHIP_BOL_NO)
        Fill_Records("EDT850T7", SHIP_BOL_NO)
        Fill_Records("EDT850T8", SHIP_BOL_NO)

        rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
        rowSOTSVIA1 = dst.Tables("SOTSVIA1").Rows.Find(rowSOTSHIP1.Item("SHIP_VIA_CODE"))

        Dim WHSE_CODE As String = rowSOTSHIP1.Item("WHSE_CODE")
        rowICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
        BILL_OF_LADING_NO = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & ""

    End Sub

    Sub Process_SOTINVH1s()
        If rowSOTINVH1.Item("INV_NO_CONS") & "" <> "" Then
            If rowSOTINVH1.Item("INV_NO_CONS") <> INV_NO_CONS Then
                INV_NO_CONS = rowSOTINVH1.Item("INV_NO_CONS")
                rowSOTINVHC = Fill_Record("SOTINVHC", INV_NO_CONS)
                CONS_INV = False
            Else
                CONS_INV = True
            End If
        Else
            INV_NO_CONS = ""
            CONS_INV = False
        End If

        rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(rowSOTINVH1.Item("ORDR_NO"))
        EDI_DOC_SEQ_NO = rowSOTORDR1.Item("EDI_DOC_SEQ_NO")

        rowEDT850T1 = dst.Tables("EDT850T1").Rows.Find(rowSOTORDR1.Item("EDI_DOC_SEQ_NO"))
        EDI_MERCH_TYPE = rowEDT850T1.Item("EDI_MERCH_TYPE") & ""
        EDI_OUR_ID = rowEDT850T1.Item("EDI_OUR_ID") & ""
        EDI_OUR_QUAL = rowEDT850T1.Item("EDI_OUR_QUAL") & ""

        rowSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(rowSOTINVH1.Item("PICK_NO"))
        rowTATTERM1 = dst.Tables("TATTERM1").Rows.Find(rowSOTINVH1.Item("TERM_CODE"))

        ORDR_SOURCE = rowSOTORDR1.Item("ORDR_SOURCE") & ""

        BILLING_INV_NO = IIf(CONS_INV, INV_NO_CONS, rowSOTINVH1.Item("INV_NO")) '  rowSOTINVH1.Item("BILLING_INV_NO")

        EDI_CUSTOMER = rowEDT850T1.Item("EDI_CUSTOMER") & ""

        Dim CUST_CODE_SO = rowSOTINVH1.Item("CUST_CODE")
        If CUST_CODE_SO = "WALMARTCOM" Or CUST_CODE_SO = "SAMSCLUB" Then
            CUST_CODE_SO = "WALMART"
        End If
        If CUST_CODE_SO = "SEARS" Then
            CUST_CODE_SO = "KMART"
        End If
        If CUST_CODE = "SEARSCANOP" Then
            CUST_CODE = "SEARSCAN"
        End If

        If CUST_CODE_SO <> CUST_CODE Then
            CUST_CODE = CUST_CODE_SO
            rowARTCUST1 = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            Load_Customer_Related_Rows()


            ' WHAT ABOUT WHEN WE HAVE KMART/SEARS OR WALMART/SAMESCLUB

            Dim rowEDTTRPM1s() As DataRow
            rowEDTTRPM1s = dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "' and EDI_DOC_NO = '810'")
            If rowEDTTRPM1s.Length = 1 Then
                rowEDTTRPM1_810 = rowEDTTRPM1s(0)
            Else
                rowEDTTRPM1_810 = Nothing
            End If

            BYPASS_ZERO_QTY_IN_ASN = "0"
            SEND_PARENT_STYLE_IN_ASN = "0"

            rowEDTTRPM1s = dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "' and EDI_DOC_NO = '856'")
            If rowEDTTRPM1s.Length = 1 Then
                rowEDTTRPM1_856 = rowEDTTRPM1s(0)
                BYPASS_ZERO_QTY_IN_ASN = rowEDTTRPM1_856.Item("BYPASS_ZERO_QTY_IN_ASN") & ""
                SEND_PARENT_STYLE_IN_ASN = rowEDTTRPM1_856.Item("SEND_PARENT_STYLE_IN_ASN") & ""
            Else
                rowEDTTRPM1_856 = Nothing
            End If
        End If

        If PROCESS_810 = "1" And CONS_INV = False Then
            Build_810()
        End If

        Dim INVCTN As Integer = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_UNITS)", "PICK_NO = '" & rowSOTINVH1.Item("PICK_NO") & "'") & "")

        If PROCESS_856 = "1" Then
            If Val(rowSOTPICK1.Item("PICK_CNT_CARTONS") & "") = 0 Or INVCTN = 0 Then
                ' DO NOTHING - MAYBE BASED ON A FLAG IN THE CUSTOMER TABLE WE NEED TO REPORT EMPTY CARTONS
            Else
                If rowSOTINVH1.Item("EDI_RETRANSMIT_IND") = "1" Then
                    ' DO NOTHING
                Else
                    Build_856()
                End If
            End If
        End If
    End Sub

    Sub Remove_Doc_from_Access()
        Dim sqlw As String = "EDI_OUR_ID = '" & EDI_OUR_ID & "' AND EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'"
        For Each TABLE_NAME As String In EDI_TABLES
            ASCDATA1.DeleteRows(TABLE_NAME, sqlw)
        Next
    End Sub

    Sub Build_810()
        Dim Last_Range_Number As String

        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
        Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
        Dim EDI_DOC_SEQ_NO As String = rowSOTORDR1.Item("EDI_DOC_SEQ_NO")
 
        'OraD.Parameters("CODE1") = rowSOTINVH1.item("INV_NO")
        'OraD.Parameters("CODE2") = rowSOTINVH1.item("ORDR_NO")
        'OraD.Parameters("CODE3") = tblSOWORDR1.Fields("EDI_DOC_SEQ_NO")


        Dim SORT As String = ""

        If INV_NO_CONS = "" Then

            '"Select SOTINVH2.*, SOTINVH1.ORDR_NO, SOTORDR1.EDI_DOC_SEQ_NO " _
            '& "from SOTINVH2, SOTORDR1, SOTINVH1" _
            '& "where SOTINVH2.INV_NO = '0005000425' " _
            '& "  AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
            '& "  AND SOTINVH1.INV_NO   = SOTINVH2.INV_NO" _
            '& "  AND SOTORDR1.ORDR_NO  = SOTINVH1.ORDR_NO" _
            '& "ORDER BY SOTINVH2.RANGE_STYLE_LNO"

            sql = "Select SOWINVH2.*, ORDR_NO, EDI_DOC_SEQ_NO from SOWINVH2, SOWORDR2 where INV_NO = '" & INV_NO & "'"
            sql = sql & " AND ORDR_NO = '" & ORDR_NO & "' AND INV_LNO = ORDR_LNO "
            sql = sql & " ORDER BY SOWINVH2.RANGE_STYLE_LNO"

            SORT = "RANGE_STYLE_LNO"

        Else

            'Select MIN(SOTORDR2.RANGE_STYLE_LNO) as RANGE_STYLE_LNO, SOTORDR2.RANGE_STYLE_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, 
            'MAX(SOTINVH2.ORDR_UNIT_PRICE) as ORDR_UNIT_PRICE, MAX(SOTINVH2.ORDR_UNIT_PRICE_CURR) as ORDR_UNIT_PRICE_CURR, MIN(SOTINVH2.INV_LNO) as INV_LNO, 
            'SUM(SOTINVH2.ORDR_QTY_SHIP) as ORDR_QTY_SHIP, MIN(SOTORDR2.ORDR_NO || SOTORDR2.ORDR_LNO) as ORDR_NO, 
            'MIN(SOTORDR2.EDI_DOC_SEQ_NO || SOTORDR2.EDI_DTL_SEQ) as EDI_DOC_SEQ_NO
            'from SOTINVH2, SOTORDR2, SOTINVH1 where SOTINVH1.INV_NO_CONS = '0005000425'
            'and SOTINVH2.INV_NO = SOTINVH1.INV_NO
            'and SOTORDR2.ORDR_NO = SOTINVH1.ORDR_NO
            'and SOTINVH2.INV_LNO = SOTORDR2.ORDR_LNO
            'group by SOTORDR2.RANGE_STYLE_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTORDR2.CUST_UPC

            'OUT OF BALANCE - MEIJER - REPLACE LINE_NUMBER WITH UPC, BURLINGTON REMOVE LINE_NO
            sql = "Select MIN(SOWORDR2.RANGE_STYLE_LNO) as RANGE_STYLE_LNO, SOWORDR2.RANGE_STYLE_CODE, SOWINVH2.STYLE_CODE, SOWINVH2.COLOR_CODE, "
            sql = sql & " MAX(SOWINVH2.ORDR_UNIT_PRICE) as ORDR_UNIT_PRICE, MAX(SOWINVH2.ORDR_UNIT_PRICE_CURR) as ORDR_UNIT_PRICE_CURR, MIN(SOWINVH2.INV_LNO) as INV_LNO, "
            sql = sql & " SUM(SOWINVH2.ORDR_QTY_SHIP) as ORDR_QTY_SHIP, MIN(SOWORDR2.ORDR_NO & SOWORDR2.ORDR_LNO) as ORDR_NO, "
            sql = sql & " MIN(SOWORDR2.EDI_DOC_SEQ_NO & SOWORDR2.EDI_DTL_SEQ) as EDI_DOC_SEQ_NO"
            sql = sql & " from SOWINVH2, SOWORDR2, SOWINVH1 where INV_NO_CONS = '" & INV_NO_CONS & "'"
            sql = sql & " and SOWINVH2.INV_NO = SOWINVH1.INV_NO"
            sql = sql & " and SOWORDR2.ORDR_NO = SOWINVH1.ORDR_NO"
            sql = sql & " and SOWINVH2.INV_LNO = SOWORDR2.ORDR_LNO"
            sql = sql & " group by SOWORDR2.RANGE_STYLE_CODE, SOWINVH2.STYLE_CODE, SOWINVH2.COLOR_CODE, SOWORDR2.CUST_UPC"

            SORT = "RANGE_STYLE_CODE, STYLE_CODE, COLOR_CODE, CUST_UPC"

        End If


        Dim Hdr As Boolean = True
        Dim CURRENT_RANGE_STYLE_LNO As Integer = 0

        Last_Range_Number = ""
        DTL_TOT_CHK = 0
        HDR_TOT_CHK = 0
        Dim CONS_INV_LNO As Int32 = 0

        For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select(sql, SORT)
            Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & "")

            'Next

            'Do While Not dynSOWINVH2.EOF
            If (ORDR_QTY_SHIP = 0 _
                And BYPASS_ZERO_QTY_IN_ASN = "1") Then
                GoTo get_next
            End If
            GTIN = ""
            UPC = ""
            SKU = ""
            EDI_STYLE = ""

            Dim rowSOTORDR2 As DataRow = Nothing
            If INV_NO_CONS = "" Then
                rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {rowSOTINVH2.Item("ORDR_NO"), rowSOTINVH2.Item("INV_LNO")})
            Else
                rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {Mid(rowSOTINVH2.Item("ORDR_NO"), 1, 10), Val(Mid(rowSOTINVH2.Item("ORDR_NO"), 11))})
            End If

            Dim rowEDT850T2 As DataRow = Nothing
            If ORDR_SOURCE = "E" Then
                If INV_NO_CONS <> "" Then
                    rowEDT850T2 = dst.Tables("EDT850T2").Rows.Find(New Object() {Mid(rowSOTINVH2.Item("EDI_DOC_SEQ_NO"), 1, 10), Val(Mid(rowSOTINVH2.Item("EDI_DOC_SEQ_NO"), 11))})
                ElseIf Val(rowSOTORDR2.Item("EDI_DTL_SEQ")) <> 0 Then
                    rowEDT850T2 = dst.Tables("EDT850T2").Rows.Find(New Object() {rowSOTORDR1.Item("EDI_DOC_SEQ_NO"), Val(rowSOTORDR2.Item("EDI_DTL_SEQ"))})
                End If
                If rowEDT850T2 IsNot Nothing Then
                    EDI_STYLE = rowEDT850T2.Item("EDI_STYLE") & ""
                End If
            End If
            If Hdr Then
                Write_Header_810()
                If PROCESS_810 <> "1" Then
                    Exit For ' Do
                End If
                Ship_Addresses("EDT810O5", rowEDTTRPM1_810)
                Inv_Addresses("EDT810O5", rowEDTTRPM1_810, False)
                Hdr = False
            End If
            If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" = "" Then
                CURRENT_RANGE_STYLE_LNO = 0
                Last_Range_Number = ""
            End If
            If INV_NO_CONS <> "" Then
                CONS_INV_LNO = CONS_INV_LNO + 1
            End If

            Dim rowEDT810O2 As DataRow = Nothing
            Dim rowEDT850T6 As DataRow = Nothing

            Dim PO4_UOM As String

            'The comparison below needs to be modified for burlington, it shouldn't compare line numbers
            If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" And (rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> Last_Range_Number Or rowSOTORDR2.Item("RANGE_STYLE_LNO") <> CURRENT_RANGE_STYLE_LNO) And CUST_CODE <> "*SEARS*" Then
                Dim rowSOTINVH9 As DataRow = dst.Tables("SOTINVH9").Rows.Find(New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO"), rowSOTORDR2.Item("RANGE_STYLE_LNO")})
                Dim rowSOTORDR9 As DataRow = dst.Tables("SOTORDR9").Rows.Find(New Object() {rowSOTORDR2.Item("ORDR_NO"), rowSOTORDR2.Item("RANGE_STYLE_LNO")})
                rowEDT850T6 = dst.Tables("EDT850T6").Rows.Find(New Object() {rowSOTORDR1.Item("EDI_DOC_SEQ_NO"), rowSOTORDR2.Item("EDI_DTL_SEQ"), rowSOTORDR2.Item("RANGE_STYLE_LNO")})

                Last_Range_Number = rowSOTORDR2.Item("RANGE_STYLE_CODE")

                rowEDT810O2 = dst.Tables("EDT810O2").NewRow
                With rowEDT810O2

                    .Item("EDI_OUR_ID") = EDI_OUR_ID
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("INV_NO") = BILLING_INV_NO
                    .Item("EDI_UPC_2") = ""
                    If INV_NO_CONS = "" Then
                        .Item("INV_LNO") = rowSOTINVH9.Item("RANGE_STYLE_LNO")
                    Else
                        .Item("INV_LNO") = CONS_INV_LNO
                    End If
                    .Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
                    .Item("EDI_SIZE_DESC") = rowEDT850T2.Item("EDI_SIZE_DESC") & ""
                    .Item("EDI_LINE_STATUS") = "I"

                    If (CUST_CODE = "BURLING" Or CUST_CODE = "WALMART" Or CUST_CODE = "NORDSTR" Or CUST_CODE = "MEIJER" Or CUST_CODE = "KMART" Or CUST_CODE = "BELKS") _
                        And INV_NO_CONS <> "" Then
                        'sometimes SEARS/Kmart EMP orders

                        Dim rowWK As DataRow = Fill_Record("SOTINVH9R", New String() {INV_NO_CONS, rowSOTINVH2.Item("RANGE_STYLE_CODE")})

                        .Item("ORDR_UNIT_PRICE_CURR") = rowWK.Item(0)
                        .Item("ORDR_UNIT_PRICE") = rowWK.Item(1)
                        .Item("ORDR_QTY_SHIP") = rowWK.Item(2)

                    Else

                        .Item("ORDR_UNIT_PRICE_CURR") = rowSOTINVH9.Item("RANGE_STYLE_PP_PRICE_CURR") & ""
                        .Item("ORDR_UNIT_PRICE") = rowSOTINVH9.Item("RANGE_STYLE_PP_PRICE") & ""
                        .Item("ORDR_QTY_SHIP") = rowSOTINVH9.Item("RANGE_STYLE_PP_QTY_SHIP") & ""
                        If rowSOTINVH9.Item("RANGE_STYLE_QTY_PER_PP") & "" = 1 Then
                            .Item("ORDR_QTY_SHIP") = rowSOTINVH9.Item("RANGE_STYLE_PRICE") * rowSOTINVH9.Item("RANGE_STYLE_PP_QTY_SHIP") / rowSOTINVH9.Item("RANGE_STYLE_PP_PRICE")
                        End If
                    End If

                    .Item("STYLE_CODE") = rowSOTORDR2.Item("RANGE_STYLE_CODE")
                    If ORDR_SOURCE = "E" And rowSOTORDR2.Item("EDI_DTL_SEQ") <> 0 Then
                        If EDI_STYLE <> "" Then
                            .Item("STYLE_CODE") = EDI_STYLE
                        End If
                        GTIN = rowEDT850T2.Item("EDI_GTIN") & ""
                        UPC = rowEDT850T2.Item("EDI_UPC") & ""
                        SKU = rowEDT850T2.Item("EDI_SKU") & ""
                        PO4_UOM = Trim(rowEDT850T2.Item("EDI_PO4_UOM")) & ""
                        If CUST_CODE = "BURLING" Or CUST_CODE = "CHARLOT" Then
                            PO4_UOM = UCase(Trim(rowEDT850T2.Item("EDI_PO4_UOM")) & "")
                        End If
                        If CUST_CODE = "SEARS" Then
                            'I856.PO4_UOM = "EA"
                        End If
                        .Item("ITEM_DESC") = rowSOTORDR9.Item("RANGE_STYLE_DESC") & ""
                    Else
                        PO4_UOM = Trim(rowSOTORDR9.Item("RANGE_STYLE_UOM")) & ""
                        If CUST_CODE = "BURLING" Or CUST_CODE = "CHARLOT" Then
                            PO4_UOM = UCase(Trim(rowSOTORDR9.Item("RANGE_STYLE_UOM")) & "")
                        End If
                        .Item("ITEM_DESC") = rowSOTORDR9.Item("RANGE_STYLE_DESC") & ""
                    End If

                    If SKU = "" Or UPC = "" Then
                        Dim UPCXREF As New Dictionary(Of String, String)
                        UPCXREF.Add("CUST_CODE", rowSOTINVH1.Item("CUST_CODE"))

                        If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" Then
                            UPCXREF.Add("STYLE_CODE", rowSOTORDR2.Item("RANGE_STYLE_CODE") & "")
                            UPCXREF.Add("COLOR_CODE", "AST")
                        Else
                            UPCXREF.Add("STYLE_CODE", rowSOTORDR2.Item("STYLE_CODE"))
                            UPCXREF.Add("COLOR_CODE", rowSOTORDR2.Item("COLOR_CODE"))
                            UPCXREF.Add("CODE", rowSOTORDR2.Item("CUST_COLOR_CODE"))
                        End If

                        Get_UPC(UPCXREF)
                    End If

                    .Item("EDI_UPC") = UPC
                    .Item("EDI_SKU") = SKU
                    .Item("STYLE_GTIN_CODE") = GTIN

                    If CUST_CODE = "KMART" And EDI_MERCH_TYPE <> "J1" And GTIN = "" Then ' change this in the near future to use new field in 850t2
                        Dim rowICTGTINTs() As DataRow = dst.Tables("ICTGTINT").Select("GTIN_UPC_CODE = '" & UPC & "'")
                        If rowICTGTINTs.Length > 0 Then
                            .Item("STYLE_GTIN_CODE") = rowICTGTINTs(0).Item("GTIN_CODE")
                        Else
                            MsgBox("Gtin Missing for Kmart, verify Invoice", vbOKOnly, "Invoice Error")
                        End If
                    End If

                    .Item("EDI_QTY_UOM") = Trim(rowSOTORDR2.Item("STYLE_UOM")) & ""
                    .Item("EDI_PRICE_UOM") = Trim(rowSOTORDR2.Item("STYLE_UOM")) & ""
                    If CUST_CODE = "SEARS" Or CUST_CODE = "BURLING" Or CUST_CODE = "CHARLOT" Then
                        .Item("EDI_QTY_UOM") = UCase(Trim(rowEDT850T2.Item("EDI_PO4_UOM")) & "")
                        .Item("EDI_PRICE_UOM") = UCase(Trim(rowEDT850T2.Item("EDI_PO4_UOM")) & "")
                    End If
                    If CUST_CODE = "WALMARTCA" And 1 = 2 Then ' Don't do special code for WmrtCa
                        DTL_TOT_CHK = DTL_TOT_CHK + (.Item("ORDR_QTY_SHIP") * .Item("ORDR_UNIT_PRICE_CURR"))
                    Else
                        DTL_TOT_CHK = DTL_TOT_CHK + (.Item("ORDR_QTY_SHIP") * .Item("ORDR_UNIT_PRICE"))
                    End If


                    dst.Tables("EDT810O2").Rows.Add(rowEDT810O2)
                End With

                CURRENT_RANGE_STYLE_LNO = rowSOTORDR2.Item("RANGE_STYLE_LNO")
            End If

            rowEDT810O2 = dst.Tables("EDT810O2").NewRow
            With rowEDT810O2
                .Item("EDI_OUR_ID") = EDI_OUR_ID
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("INV_NO") = BILLING_INV_NO

                If INV_NO_CONS = "" Then
                    .Item("INV_LNO") = rowSOTINVH2.Item("INV_LNO")
                Else
                    .Item("INV_LNO") = CONS_INV_LNO
                End If
                .Item("EDI_PO4_QTY") = 0 ' rowEDT850T2.Item("EDI_PO4_QTY")
                .Item("ORDR_UNIT_PRICE_CURR") = rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") & ""
                .Item("ORDR_UNIT_PRICE") = rowSOTINVH2.Item("ORDR_UNIT_PRICE") & ""
                .Item("ORDR_QTY_SHIP") = rowSOTINVH2.Item("ORDR_QTY_SHIP") & ""
                .Item("EDI_LINE_STATUS") = "I"


                Dim REC_ID As String = ""

                If CURRENT_RANGE_STYLE_LNO = 0 Then
                    REC_ID = "DTL"
                    If rowSOTORDR2.Item("STYLE_CODE_SUB") & "" <> "" Then
                        .Item("STYLE_CODE") = rowSOTORDR2.Item("STYLE_CODE_SUB")
                    Else
                        .Item("STYLE_CODE") = rowSOTINVH2.Item("STYLE_CODE")
                    End If

                Else
                    If rowSOTORDR1.Item("ORDR_SOURCE") & "" <> "K" Then
                        REC_ID = "SUB"
                        If rowEDT850T6 Is Nothing Then
                            .Item("STYLE_CODE") = rowSOTINVH2.Item("STYLE_CODE")
                        ElseIf rowEDT850T6.Item("EDI_SLN_STYLE") & "" <> "" Then
                            .Item("STYLE_CODE") = rowEDT850T6.Item("EDI_SLN_STYLE")
                        Else
                            .Item("STYLE_CODE") = rowSOTINVH2.Item("STYLE_CODE")
                        End If
                    End If
                End If
                .Item("COLOR_CODE") = rowSOTINVH2.Item("COLOR_CODE")
                If ORDR_SOURCE = "E" And rowSOTORDR2.Item("EDI_DTL_SEQ") <> 0 Then
                    If CURRENT_RANGE_STYLE_LNO = 0 Or rowEDT850T6 Is Nothing Then
                        If EDI_STYLE <> "" Then
                            .Item("STYLE_CODE") = EDI_STYLE
                        End If
                        .Item("EDI_SIZE_DESC") = rowEDT850T2.Item("EDI_SIZE_DESC") & ""
                        SKU = rowEDT850T2.Item("EDI_SKU") & ""
                        PO4_UOM = Trim(rowEDT850T2.Item("EDI_PO4_UOM")) & ""
                        If CUST_CODE = "SEARS" Then
                            'I856.PO4_UOM = "EA"
                        End If
                        If CUST_CODE = "BURLING" Then
                            PO4_UOM = UCase(Trim(rowEDT850T2.Item("EDI_PO4_UOM")) & "")
                        End If
                        GTIN = rowEDT850T2.Item("EDI_GTIN") & ""
                        If rowEDT850T2.Item("EDI_UPC") & "" <> "" Then
                            UPC = rowEDT850T2.Item("EDI_UPC") & ""
                        ElseIf rowSOTORDR2.Item("CUST_UPC") & "" <> "" Then
                            UPC = rowSOTORDR2.Item("CUST_UPC") & ""
                        End If
                        If rowEDT850T2.Item("EDI_STYLE_NAME") & "" <> "" Then
                            .Item("ITEM_DESC") = rowEDT850T2.Item("EDI_STYLE_NAME") & ""
                        Else
                            .Item("ITEM_DESC") = rowSOTORDR2.Item("STYLE_DESC") & ""
                        End If
                    Else
                        If rowEDT850T6.Item("EDI_SLN_STYLE") & "" <> "" Then
                            .Item("STYLE_CODE") = rowEDT850T6.Item("EDI_SLN_STYLE")
                        End If
                        .Item("EDI_SIZE_DESC") = rowEDT850T6.Item("EDI_SLN_SIZE_DESC") & ""
                        .Item("EDI_SKU") = rowEDT850T6.Item("EDI_SLN_SKU") & ""
                        PO4_UOM = Trim(rowEDT850T6.Item("EDI_SLN_PO4_UOM")) & ""
                        If CUST_CODE = "BURLING" Or CUST_CODE = "CHARLOT" Then
                            PO4_UOM = UCase(Trim(rowEDT850T6.Item("EDI_SLN_PO4_UOM")) & "")
                        End If
                        If rowEDT850T6.Item("EDI_SLN_UPC") & "" <> "" Then
                            UPC = rowEDT850T6.Item("EDI_SLN_UPC") & ""
                        End If

                        .Item("ITEM_DESC") = rowSOTORDR2.Item("STYLE_DESC") & ""

                    End If
                Else
                    PO4_UOM = Trim(rowSOTORDR2.Item("STYLE_UOM")) & ""
                    If CUST_CODE = "BURLING" Then
                        PO4_UOM = UCase(Trim(rowSOTORDR2.Item("STYLE_UOM")) & "")
                    End If
                    If CUST_CODE = "SEARS" Then
                        'I856.PO4_UOM = "EA"
                    End If
                    UPC = rowSOTORDR2.Item("CUST_UPC") & ""
                    If rowSOTORDR2.Item("CUST_UPC") & "" <> "" Then
                        UPC = rowSOTORDR2.Item("CUST_UPC") & ""
                        SKU = rowSOTORDR2.Item("CUST_SKU") & ""
                    End If
                    .Item("ITEM_DESC") = rowSOTORDR2.Item("STYLE_DESC") & ""
                End If

                If SKU = "" Or UPC = "" Then
                    Dim UPCXREF As New Dictionary(Of String, String)
                    UPCXREF.Add("CUST_CODE", rowSOTINVH1.Item("CUST_CODE"))
                    UPCXREF.Add("STYLE_CODE", rowSOTORDR2.Item("STYLE_CODE"))
                    UPCXREF.Add("COLOR_CODE", rowSOTORDR2.Item("COLOR_CODE"))
                    UPCXREF.Add("CODE", rowSOTORDR2.Item("CUST_COLOR_CODE"))

                    Get_UPC(UPCXREF)
                End If

                Stop ' NEED TO SET ORDR_SOURCE = rowSOTORDR1.item("ORDR_SOURCE") USING V2 SYNTAX SOMEWHERE IN TOP OF LOOP
                If CUST_CODE = "KMART" And ORDR_SOURCE <> "K" And GTIN = "" Then
                    If EDI_MERCH_TYPE <> "J1" Then ' change this in the near future to use new field in 850t2
                        Dim rowICTGTINTs() As DataRow = dst.Tables("ICTGTINT").Select("GTIN_UPC_CODE = '" & UPC & "'")
                        If rowICTGTINTs.Length > 0 Then
                            GTIN = rowICTGTINTs(0).Item("GTIN_CODE")
                        Else
                            MsgBox("Gtin Missing for Kmart, verify Invoice", vbOKOnly, "Invoice Error")
                        End If
                    End If
                End If

                .Item("EDI_UPC") = UPC
                .Item("EDI_SKU") = SKU
                .Item("STYLE_GTIN_CODE") = GTIN

                If CUST_CODE = "NORDSTR" Then
                    .Item("ITEM_DESC") = Mid$(.Item("ITEM_DESC"), 1, 25)
                End If
                ' TEST AND SEE IF SUBLINES PREVENT MULTIPLE DETAIL LINES IN TRANSLATION
                If REC_ID <> "SUB" Then
                    .Item("EDI_QTY_UOM") = Trim(rowSOTORDR2.Item("STYLE_UOM")) & ""
                    .Item("EDI_PRICE_UOM") = Trim(rowSOTORDR2.Item("STYLE_UOM")) & ""
                    If CUST_CODE = "HANES" Then
                        .Item("EDI_PRICE_UOM") = "WE"
                    End If
                    DTL_TOT_CHK = DTL_TOT_CHK + (.Item("ORDR_QTY_SHIP") * .Item("ORDR_UNIT_PRICE"))

                    dst.Tables("EDT810O2").Rows.Add(rowEDT810O2)
                End If
            End With
get_next:

        Next ' Loop

        If System.Math.Round(HDR_TOT_CHK, 2) <> System.Math.Round(DTL_TOT_CHK, 2) Then

            ASCMAIN1.Progress("-", "Invoice " & BILLING_INV_NO & " NOT Transmitted")

            Dim rowEDTREJD1 As DataRow = dst.Tables("EDTREJD1").NewRow
            With rowEDTREJD1
                .Item("EDI_OUR_ID") = EDI_OUR_ID
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("EDI_DOC_TYPE") = "810"
                .Item("ABS_DOC_NO") = BILLING_INV_NO
            End With
            dst.Tables("EDTREJD1").Rows.Add(rowEDTREJD1)

            INV_ERR = INV_ERR & "," & BILLING_INV_NO
            WRITE_TO_MDB_810()
            Remove_Doc_from_Access()
            MsgBox("Invoice " & BILLING_INV_NO & " is out of balance", MsgBoxStyle.OkOnly, "Please review report")
        End If
    End Sub

    Sub Write_Header_810()

        If rowEDTTRPM1_810 Is Nothing Then
            PROCESS_810 = ""
            HDR_TOT_CHK = 0
            Return
        End If

        Stop

        Write_to_EDTSYSIH("OIN")

        Dim EDI_TP_QUAL As String = rowEDTTRPM1_810.Item("EDI_TP_QUAL")
        Dim EDI_TP_ID As String = rowEDTTRPM1_810.Item("EDI_TP_ID")
        Dim NUMBER_CHRS_SHIPTO As Integer = Val(rowEDTTRPM1_810.Item("NUMBER_CHRS_SHIPTO") & "")

        Dim rowEDT810O1 As DataRow = dst.Tables("EDT810O1").NewRow
        With rowEDT810O1
            .Item("EDI_OUR_ID") = EDI_OUR_ID
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            .Item("INV_NO") = BILLING_INV_NO
            .Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO")
            .Item("EDI_TP_QUAL") = EDI_TP_QUAL ' GET THIS FROM EDTTRPM1
            .Item("EDI_TP_ID") = EDI_TP_ID ' GET THIS FROM EDTTRPM1

            Dim ORDR_DEPT As String = rowSOTINVH1.Item("ORDR_DEPT") & ""
            .Item("ORDR_DEPT") = ORDR_DEPT
            If CUST_CODE = "SEARS" And Len(ORDR_DEPT) < 3 Then
                .Item("ORDR_DEPT") = ORDR_DEPT & "8"
            End If

            .Item("ORDR_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")

            If ORDR_SOURCE = "E" Then
                If EDI_MERCH_TYPE = "J1" And rowSOTSHIP1.Item("SHIP_ADDR_CODE") & "" <> "" Then
                    .Item("ORDR_CUST_PO") = Format$(rowSOTSHIP1.Item("SHIP_ADDR_CODE"), "00000") & Mid$(rowSOTINVH1.Item("ORDR_CUST_PO"), 6)
                End If
            End If

            .Item("ORDR_DATE") = rowSOTORDR1.Item("ORDR_DATE")
            .Item("INV_DATE") = rowSOTINVH1.Item("INV_DATE")
            .Item("INV_DATE_SHIPPED") = rowSOTINVH1.Item("INV_DATE")
            .Item("SHIP_MANIFEST_NO") = rowSOTSHIP1.Item("SHIP_MANIFEST_NO") & ""
            .Item("INV_BOL_NO") = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & ""
            .Item("INV_PRO_NO") = rowSOTSHIP1.Item("SHIP_REF") & ""
            .Item("SHIP_ADDR_CODE") = rowSOTSHIP1.Item("SHIP_ADDR_CODE")

            .Item("SHIP_VIA_SCAC") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & ""
            .Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & ""
            .Item("WHSE_CODE") = rowSOTSHIP1.Item("WHSE_CODE")
            .Item("CURR_CODE") = "USD" '
            .Item("INV_NO_CONS") = rowSOTINVH1.Item("INV_NO_CONS")
            .Item("EDI_BUS_UNIT_NAME") = "VANDALE INDUSTRIES, INC. "
            .Item("EDI_WAREHOUSE_LOCATION") = "EDISON, NJ"
            .Item("BUS_UNIT_DUNS_NO") = "101140010"
            If CUST_CODE = "SEARS" Or CUST_CODE = "KMART" Then
                .Item("BUS_UNIT_DUNS_NO") = "00101140010"
            End If
            .Item("SHIP_810_BATCH_NO") = SHIP_8XX_BATCH_NO
            If EDI_CUSTOMER <> "" Then
                .Item("EDI_CUSTOMER") = EDI_CUSTOMER
            Else
                .Item("EDI_CUSTOMER") = rowSOTINVH1.Item("CUST_CODE")
            End If
            If INV_NO_CONS = "" Then
                If rowSOTINVH1.Item("REVISED_CUST_STORE_NO") & "" = "" Then
                    If IsNumeric(rowSOTINVH1.Item("CUST_STORE_NO") & "") Then
                        .Item("CUST_STORE_NO") = Format$(rowSOTINVH1.Item("CUST_STORE_NO") & "", Mid$("000000", 1, NUMBER_CHRS_SHIPTO))
                    Else
                        .Item("CUST_STORE_NO") = rowSOTINVH1.Item("CUST_STORE_NO") & ""
                    End If
                Else
                    If IsNumeric(rowSOTINVH1.Item("REVISED_CUST_STORE_NO") & "") Then
                        .Item("CUST_STORE_NO") = Format$(rowSOTINVH1.Item("REVISED_CUST_STORE_NO") & "", Mid$("000000", 1, NUMBER_CHRS_SHIPTO))
                    Else
                        .Item("CUST_STORE_NO") = rowSOTINVH1.Item("REVISED_CUST_STORE_NO") & ""
                    End If
                End If
                .Item("INV_CARTONS") = rowSOTPICK1.Item("PICK_CNT_CARTONS") & ""
                .Item("INV_WEIGHT") = rowSOTPICK1.Item("PICK_TOTAL_WGT") & ""
                .Item("INV_FREIGHT") = rowSOTINVH1.Item("INV_FREIGHT") + 0
                .Item("INV_MISC_CHG") = rowSOTINVH1.Item("INV_MISC_CHG") & ""
                .Item("INV_SALES") = rowSOTINVH1.Item("INV_SALES") & ""
                .Item("INV_TOTAL_AMOUNT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT") ' - Allowances
                .Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMT_CURR")
                .Item("GST_TAX") = rowSOTINVH1.Item("GST_TAX")
                .Item("GST_TAX_CURR") = rowSOTINVH1.Item("GST_TAX_CURR")
                .Item("INV_SALES_CURR") = rowSOTINVH1.Item("INV_SALES_CURR")
            Else
                .Item("CUST_STORE_NO") = ""
                .Item("INV_CARTONS") = Val(rowSOTINVHC.Item("PICK_CNT_CARTONS") & "")
                .Item("INV_WEIGHT") = Val(rowSOTINVHC.Item("PICK_TOTAL_WGT") & "")
                .Item("INV_FREIGHT") = Val(rowSOTINVHC.Item("INV_FREIGHT") & "") + 0
                .Item("INV_MISC_CHG") = Val(rowSOTINVHC.Item("INV_MISC_CHG") & "")
                .Item("INV_TOTAL_AMOUNT") = Val(rowSOTINVHC.Item("INV_TOTAL_AMOUNT") & "") '- Allowances
                .Item("INV_SALES") = Val(rowSOTINVHC.Item("INV_SALES") & "")
                .Item("INV_TOTAL_AMOUNT_CURR") = Val(rowSOTINVHC.Item("INV_TOTAL_AMT_CURR") & "")
                .Item("GST_TAX") = Val(rowSOTINVHC.Item("GST_TAX") & "")
                .Item("GST_TAX_CURR") = Val(rowSOTINVHC.Item("GST_TAX_CURR") & "")
                .Item("INV_SALES_CURR") = Val(rowSOTINVHC.Item("INV_SALES_CURR") & "")
            End If

            If rowSOTORDR1.Item("FRT_TERMS") & "" = "PPD" Then
                .Item("EDI_FOB") = "PP"
            ElseIf rowSOTORDR1.Item("FRT_TERMS") & "" = "PPA" Then
                .Item("EDI_FOB") = "PC"
            Else
                .Item("EDI_FOB") = "CC"
            End If

            Dim temp_date As Date = Now + ASCMAIN1.NowTSD
            If rowTATTERM1.Item("TERM_DUE_TYPE") = "E" Then
                temp_date = DateAdd("D", -1, DateAdd("M", 1, DateValue(Format$(temp_date, "MM/01/YYYY"))))
            End If
            temp_date = DateAdd("M", rowTATTERM1.Item("TERM_ADDL_MOS"), DateAdd("D", rowTATTERM1.Item("TERM_DAYS_DUE"), temp_date))
            If Val(rowTATTERM1.Item("TERM_DAYS_DISC") & "") <> 0 Then
                .Item("EDI_DSC_DUEDATE") = DateAdd("D", rowTATTERM1.Item("TERM_DAYS_DISC"), Now + ASCMAIN1.NowTSD)
                .Item("EDI_TERM_DSCDAYS") = rowTATTERM1.Item("TERM_DAYS_DISC")
            End If
            .Item("EDI_NET_DUEDATE") = Format$(temp_date, "MM/DD/YYYY")
            .Item("EDI_TERM_NETDAYS") = rowTATTERM1.Item("TERM_DAYS_DUE")
            If ORDR_SOURCE = "E" Then
                .Item("EDI_SUPPLIER_NO") = rowEDT850T1.Item("EDI_SUPPLIER_NO") & ""
                If (CUST_CODE = "SEARS" Or CUST_CODE = "KMART") And Len(rowEDT850T1.Item("EDI_SUPPLIER_NO") & "") = 9 Then
                    .Item("EDI_SUPPLIER_NO") = "00" & rowEDT850T1.Item("EDI_SUPPLIER_NO") & ""
                End If
                .Item("EDI_PROMOTION") = rowEDT850T1.Item("EDI_PROMOTION") & ""
                .Item("EDI_MERCH_TYPE") = EDI_MERCH_TYPE
                .Item("EDI_TERM_TYPE") = rowEDT850T1.Item("EDI_TERM_TYPE") & ""
                If rowEDT850T1.Item("EDI_TERM_BASIS") & "" <> "" Then
                    .Item("EDI_TERM_BASIS") = rowEDT850T1.Item("EDI_TERM_BASIS") & ""
                Else
                    .Item("EDI_TERM_BASIS") = "3"
                End If
                .Item("EDI_TERM_DESC") = rowEDT850T1.Item("EDI_TERM_DESC") & ""
                .Item("EDI_TERM_DOM") = rowEDT850T1.Item("EDI_TERM_DOM") & ""
            Else
                .Item("EDI_TERM_TYPE") = "05"
                .Item("EDI_TERM_BASIS") = "3"
                .Item("EDI_TERM_DESC") = rowTATTERM1.Item("TERM_DESC") & ""
            End If

            If .Item("EDI_SUPPLIER_NO") & "" = "" Then
                If rowEDTTRPM1_810.Item("EDI_SUPPLIER_NO") & "" <> "" Then
                    .Item("EDI_SUPPLIER_NO") = rowEDTTRPM1_810.Item("EDI_SUPPLIER_NO") & ""
                Else
                    .Item("EDI_SUPPLIER_NO") = rowARTCUST1.Item("CUST_VEND_REF") & ""
                End If
            End If

            If Val(.Item("EDI_TERM_DISC_AMT")) = 0 Then
                .Item("EDI_DSC_DUEDATE") = Null
                .Item("EDI_TERM_DISC_AMT") = Null
            End If

            'Get header allowances and write them out on the invoice
            If ORDR_SOURCE = "E" Then
                Dim Allowances As Decimal = 0
                For Each rowEDT850T7 As DataRow In dst.Tables("EDT850T7").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    Dim rowEDT810O7 As DataRow = dst.Tables("EDT810O7").NewRow
                    With rowEDT810O7
                        .Item("EDI_OUR_ID") = EDI_OUR_ID
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        .Item("SAH_SEQ_NO") = rowEDT850T7.Item("SAH_SEQ_NO") & ""
                        .Item("INV_NO") = BILLING_INV_NO
                        .Item("SAH_ALLOW_IND") = rowEDT850T7.Item("SAH_ALLOW_IND") & ""
                        .Item("SAH_ALLOW_CODE") = rowEDT850T7.Item("SAH_ALLOW_CODE") & ""
                        If rowEDT850T7.Item("SAH_PERCENT") & "" <> "" Then
                            .Item("SAH_AMOUNT") = System.Math.Round((Val(.Item("INV_SALES") + 0) * Val(rowEDT810O7.Item("SAH_PERCENT") + 0)) / 100 + 0.0001, 2)
                        End If
                        .Item("SAH_PERCENT_QUAL") = rowEDT850T7.Item("SAH_PERCENT_QUAL") & ""
                        .Item("SAH_PERCENT") = rowEDT850T7.Item("SAH_PERCENT") + 0
                        .Item("SAH_RATE") = rowEDT850T7.Item("SAH_RATE") + 0
                        .Item("SAH_UOM_CODE") = rowEDT850T7.Item("SAH_UOM_CODE") & ""
                        .Item("SAH_QTY") = rowEDT850T7.Item("SAH_QTY") + 0
                        .Item("SAH_HANDLING_CODE") = rowEDT850T7.Item("SAH_HANDLING_CODE") & ""
                    End With

                    dst.Tables("EDT810O7").Rows.Add(rowEDT810O7)
                Next

                .Item("INV_TOTAL_AMOUNT") = .Item("INV_TOTAL_AMOUNT") - Allowances
            End If

            HDR_TOT_CHK = Val(.Item("INV_SALES"))
        End With
    End Sub

    Sub WRITE_TO_MDB_810()

        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
        Dim INV_TOTAL_AMOUNT As Decimal = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & "")

        Dim SHIP_ADDR_CODE As String = rowSOTSHIP1.Item("SHIP_ADDR_CODE")
        Dim BILL_OF_LADING_NO As String = rowSOTSHIP1.Item("BILL_OF_LADING_NO")

        Stop
        If INV_NO_CONS = "" Then
            Stop
        End If
        Dim rowEDTOUTB1 As DataRow = dst.Tables("EDTOUTB1").Rows.Find _
                                     (New Object() {SHIP_8XX_BATCH_NO, SHIP_ADDR_CODE, BILL_OF_LADING_NO})

        If rowEDTOUTB1 Is Nothing Then
            rowEDTOUTB1 = dst.Tables("EDTOUTB1").NewRow

            With rowEDTOUTB1
                .Item("SHIP_8XX_BATCH_NO") = SHIP_8XX_BATCH_NO
                .Item("CUST_DC_NO") = SHIP_ADDR_CODE
                .Item("BILL_OF_LADING_NO") = BILL_OF_LADING_NO
                .Item("CUST_CODE") = CUST_CODE

                If INV_NO_CONS = "" Then
                    .Item("FIRST_INV_NO") = INV_NO
                    .Item("LAST_INV_NO") = INV_NO
                    .Item("DOLLAR_AMT") = INV_TOTAL_AMOUNT
                Else

                    .Item("FIRST_INV_NO") = INV_NO_CONS
                    .Item("LAST_INV_NO") = INV_NO_CONS
                    .Item("DOLLAR_AMT") = Val(rowSOTINVHC.Item("INV_TOTAL_AMOUNT") & "")
                End If
                .Item("NUM_OF_810") = 1
            End With

        Else

            With rowEDTOUTB1
                If INV_NO_CONS = "" Then
                    If INV_NO > .Item("LAST_INV_NO") Then
                        .Item("LAST_INV_NO") = INV_NO
                    End If
                    If INV_NO < .Item("FIRST_INV_NO") Then
                        .Item("FIRST_INV_NO") = INV_NO
                    End If
                    .Item("DOLLAR_AMT") = Val(.Item("DOLLAR_AMT") & "") + INV_TOTAL_AMOUNT
                Else
                    If INV_NO_CONS > .Item("LAST_INV_NO") Then
                        .Item("LAST_INV_NO") = INV_NO_CONS
                    End If
                    If INV_NO_CONS < .Item("FIRST_INV_NO") Then
                        .Item("FIRST_INV_NO") = INV_NO_CONS
                    End If
                    .Item("DOLLAR_AMT") = Val(.Item("DOLLAR_AMT") & "") + Val(rowSOTINVHC.Item("INV_TOTAL_AMOUNT") & "")
                End If
                .Item("NUM_OF_810") = .Item("NUM_OF_810") + 1
            End With
        End If

        rowEDTOUTB1.Item("CUST_NAME") = CUST_NAME
        rowEDTOUTB1.Item("ERRORS") = Mid(INV_ERR, 2)
    End Sub

    Sub Build_856()

        If first_invoice Then ' Write EDT856O1 for 1st invoice for each Shipment within a BILL_OF_LADING_NO
            first_invoice = False
            Write_Header_856()
        End If

        EDI_HL2_SEQ = EDI_HL2_SEQ + 1

        If EDI_HL2_SEQ = 1 Then ' assumes only 1 ST address per shipment - this was proven in data EDT856O5
            ' THIS IS THE SPOT THAT WILL MATTER IF WE ever NEED TO WRITE MULTIPLE STS PER SHIPMENT
            Ship_Addresses("EDT856O5", rowEDTTRPM1_856) ' WRITE ST/SF TO EDT856O5
        End If

        Dim rowEDT856O2 As DataRow = dst.Tables("EDT856O2").NewRow
        With rowEDT856O2
            .Item("EDI_OUR_ID") = EDI_OUR_ID
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            .Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
            .Item("ORDR_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
            .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT") & ""
            .Item("EDI_ORD_CNT_CARTONS") = rowSOTPICK1.Item("PICK_CNT_CARTONS")
            .Item("EDI_ORD_TOTAL_WGT") = rowSOTPICK1.Item("PICK_TOTAL_WGT")
            .Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO") & ""
            .Item("INV_NO") = BILLING_INV_NO
            .Item("PICK_NO") = rowSOTINVH1.Item("PICK_NO") & ""
            .Item("PRO_NO") = rowSOTSHIP1.Item("SHIP_REF") & ""
            .Item("ORDR_DATE") = rowSOTORDR1.Item("ORDR_DATE") & ""

            If IsNumeric(rowSOTINVH1.Item("CUST_STORE_NO")) Then
                .Item("CUST_STORE_NO") = Format$(Val(rowSOTINVH1.Item("CUST_STORE_NO")), Mid$("000000", 1, rowEDTTRPM1_856.Item("NUMBER_CHRS_DC")))
            Else
                .Item("CUST_STORE_NO") = rowSOTINVH1.Item("CUST_STORE_NO")
            End If

            .Item("EDI_PROMOTION") = ""
            .Item("EDI_MERCH_TYPE") = rowSOTORDR1.Item("EDI_MERCH_TYPE") & ""
            .Item("EDI_CUSTOMER") = rowEDT850T1.Item("EDI_CUSTOMER")

            If CUST_CODE = "MARSHAL" Then  ' TJX requires TJX/Marshalls Identifier.
                .Item("EDI_MERCH_TYPE") = EDI_MERCH_TYPE
            End If
            If rowSOTORDR1.Item("ORDR_STATUS") = "F" Then
                .Item("EDI_ORDER_STATUS") = "CC"
            Else
                .Item("EDI_ORDER_STATUS") = "PR"
            End If
        End With
        dst.Tables("EDT856O2").Rows.Add(rowEDT856O2)

        Dim EDI_HL3_SEQ As Integer = 0
        Dim EDI_HL4_SEQ As Integer = 0

        Inv_Addresses("EDT856O5", rowEDTTRPM1_856, True) ' WRITE MK TO EDT856O5

        Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
        Dim PICK_NO_CONS As String = rowSOTPICK1.Item("PICK_NO_CONS") & ""

        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_NO")
            If Val(rowSOTCART1.Item("CART_TOTAL_UNITS") & "") <> 0 Or BYPASS_ZERO_QTY_IN_ASN <> "1" Then
                Dim CART_NO As String = ""

                For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("CART_NO = '" & rowSOTCART1.Item("CART_NO") & "'", "CART_LNO")
                    Dim STYLE_CODE As String = rowSOTCART2.Item("STYLE_CODE") & ""
                    Dim ORDR_NO As String = rowSOTCART2.Item("ORDR_NO") & ""
                    Dim ORDR_LNO As Int32 = Val(rowSOTCART2.Item("ORDR_LNO") & "")

                    If CART_NO <> rowSOTCART1.Item("CART_NO") Then

                        Dim sqlT As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
                        Dim T As Int64 = Val(dst.Tables("SOTPICK2").Compute("Sum (PICK_QTY_CONF)", sqlT))

                        If T <> 0 Or BYPASS_ZERO_QTY_IN_ASN <> "1" Then
                            EDI_HL3_SEQ += 1
                            CART_NO = rowSOTCART1.Item("CART_NO")
                            Dim rowEDT856O3 As DataRow = dst.Tables("EDT856O3").NewRow
                            With rowEDT856O3
                                .Item("EDI_OUR_ID") = EDI_OUR_ID
                                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                                .Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                                .Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                                If SHIP_BOL_NO_CONS <> "" Or PICK_NO_CONS <> "" Then
                                    Dim rowSOTCARTP As DataRow = dst.Tables("SOTCARTP").Rows.Find(PICK_NO_CONS)
                                    Dim CART_NO_CONS As String = rowSOTCARTP.Item("CART_NO")
                                    .Item("CART_NO") = CART_NO_CONS
                                Else
                                    .Item("CART_NO") = CART_NO
                                End If

                                .Item("CART_TOTAL_WGT_ACTUAL") = rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL")
                                If Val(rowSOTCART1.Item("CART_SEQ") & "") <> 0 Then
                                    .Item("CART_SEQ") = (rowSOTCART1.Item("CART_SEQ") Mod 1000)
                                End If

                                ' use PKG_CODE to get dimensions

                                Dim rowEDTCSCD1 As DataRow = dst.Tables("EDTCSCD1").Rows.Find(New String() {CUST_CODE, STYLE_CODE})
                                If rowEDTCSCD1 IsNot Nothing Then
                                    .Item("CARTON_LENGTH") = rowEDTCSCD1.Item("CARTON_LENGTH")
                                    .Item("CARTON_WIDTH") = rowEDTCSCD1.Item("CARTON_WIDTH")
                                    .Item("CARTON_HEIGHT") = rowEDTCSCD1.Item("CARTON_HEIGHT")
                                    .Item("CARTON_WGT_PER") = rowEDTCSCD1.Item("CARTON_WGT_PER")
                                End If
                            End With
                            dst.Tables("EDT856O3").Rows.Add(rowEDT856O3)

                            EDI_HL4_SEQ = 0

                            HL4_STYLES.Clear()
                        End If
                    End If

                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                    Dim rowEDT850T2 As DataRow = dst.Tables("EDT850T2").Rows.Find(New Object() {rowSOTORDR2.Item("EDI_DOC_SEQ_NO"), rowSOTORDR2.Item("EDI_DTL_SEQ")})
                    Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").Rows.Find(New Object() {PICK_NO, ORDR_LNO})

                    If rowEDT850T2 IsNot Nothing Then
                        EDI_STYLE = rowEDT850T2.Item("EDI_STYLE") & ""
                    Else
                        EDI_STYLE = ""
                    End If

                    Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
                    If rowICTSTYL1 Is Nothing Then
                        rowICTSTYL1 = Fill_Record("ICTSTYL1", STYLE_CODE, False, False)
                    End If

                    Dim rowSOTINVH9 As DataRow = dst.Tables("SOTINVH9").Rows.Find(New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO"), rowSOTORDR2.Item("RANGE_STYLE_LNO")})
                    Dim RANGE_STYLE_CODE As String = rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" '
                    Dim PICK_QTY_CONF As Int64 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")

                    If (PICK_QTY_CONF <> 0 Or BYPASS_ZERO_QTY_IN_ASN <> "1") And Not HL4_STYLES.Contains(RANGE_STYLE_CODE) Then

                        EDI_HL4_SEQ += 1

                        ' New code to deal with Range and Assortment records that had duplicate H4 recs
                        If RANGE_STYLE_CODE <> "" Then
                            HL4_STYLES.Add(RANGE_STYLE_CODE)
                        End If

                        Dim rowEDT856O4 As DataRow = dst.Tables("EDT856O4").NewRow
                        With rowEDT856O4
                            .Item("EDI_OUR_ID") = EDI_OUR_ID
                            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            .Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                            .Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                            .Item("EDI_HL4_SEQ") = EDI_HL4_SEQ
                            .Item("CART_LNO") = rowSOTCART2.Item("CART_LNO")
                            .Item("STYLE_CODE") = rowSOTORDR2.Item("STYLE_CODE")

                            'If ORDR_SOURCE = "E" And rowSOTORDR2.Item("EDI_DTL_SEQ") <> 0 Then
                            .Item("EDI_SIZE_DESC") = rowEDT850T2.Item("EDI_SIZE_DESC") & ""
                            GTIN = rowEDT850T2.Item("EDI_GTIN") & ""
                            UPC = rowEDT850T2.Item("EDI_UPC") & ""
                            .Item("COLOR_CODE") = rowEDT850T2.Item("EDI_COLOR_CODE") & ""

                            SKU = rowEDT850T2.Item("EDI_SKU") & ""
                            If rowEDT850T2.Item("EDI_PO4_QTY") & "" <> "" Then
                                .Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY") & ""
                            ElseIf rowSOTORDR2.Item("INNER_PACK_QTY") <> 0 Then
                                .Item("EDI_PO4_QTY") = rowSOTORDR2.Item("INNER_PACK_QTY")
                            Else
                                .Item("EDI_PO4_QTY") = 1
                            End If
                            'Else
                            '    GTIN = ""
                            '    UPC = rowEDT850T2.Item("EDI_UPC") & ""
                            '    If UPC = "" Then
                            '        UPC = rowSOTORDR2.Item("CUST_UPC") & ""
                            '    End If
                            '    .Item("COLOR_CODE") = rowSOTORDR2.Item("COLOR_CODE")
                            '    SKU = rowSOTORDR2.Item("CUST_SKU") & ""
                            '    If rowSOTORDR2.Item("INNER_PACK_QTY") <> 0 Then
                            '        .Item("EDI_PO4_QTY") = rowSOTORDR2.Item("INNER_PACK_QTY")
                            '    Else
                            '        .Item("EDI_PO4_QTY") = 1
                            '    End If
                            'End If

                            If rowEDT850T2.Item("EDI_STYLE_NAME") & "" <> "" Then
                                .Item("STYLE_DESC") = rowEDT850T2.Item("EDI_STYLE_NAME") & ""
                            Else
                                .Item("STYLE_DESC") = rowSOTORDR2.Item("STYLE_DESC") & ""
                            End If

                            If rowSOTPICK2.Item("PICK_QTY_CANC") <> 0 Or rowSOTPICK2.Item("PICK_QTY_CANC_REL") <> 0 Then
                                .Item("STYLE_STATUS") = "CP"
                            ElseIf rowSOTPICK2.Item("PICK_QTY_BACK") <> 0 Then
                                .Item("STYLE_STATUS") = "BO"
                            Else
                                .Item("STYLE_STATUS") = "I"
                            End If

                            If RANGE_STYLE_CODE <> "" And CUST_CODE = "WALMART" Then
                                If rowSOTINVH9 Is Nothing Then
                                    MsgBox("Missing Range Error, call Rick", vbCritical, "ASN Error")
                                    Stop
                                End If
                                .Item("PICK_QTY_CONF") = rowSOTINVH9.Item("RANGE_STYLE_PP_QTY_SHIP")
                                .Item("EDI_PO4_UOM") = "EA"
                            Else
                                .Item("PICK_QTY_CONF") = rowSOTCART2.Item("QTY_PACKED")
                                .Item("EDI_PO4_UOM") = "EA"
                            End If

                            If rowEDT850T2.Item("EDI_PO4_UOM") = "CA" And CUST_CODE = "MARSHAL" Then
                                .Item("PICK_QTY_CONF") = 1
                                .Item("EDI_PO4_UOM") = "CA"
                            End If

                            If (SKU = "" Or UPC = "") Or SEND_PARENT_STYLE_IN_ASN = "1" Then

                                Dim UPCXREF As New Dictionary(Of String, String)

                                UPCXREF.Add("CUST_CODE", rowSOTINVH1.Item("CUST_CODE"))

                                If RANGE_STYLE_CODE <> "" And SEND_PARENT_STYLE_IN_ASN = "1" Then ' MAYBE HOLE IN LOGIC IF SEND_PARENT_STYLE_IN_ASN <> "1"
                                    UPCXREF.Add("STYLE_CODE", RANGE_STYLE_CODE)
                                    UPCXREF.Add("COLOR_CODE", "AST")

                                    .Item("STYLE_CODE") = RANGE_STYLE_CODE

                                    If rowSOTINVH9 IsNot Nothing Then
                                        .Item("PICK_QTY_CONF") = rowSOTINVH9.Item("RANGE_STYLE_PP_QTY_SHIP")
                                        'If .Item("PICK_QTY_CONF") > 1 And rowSOTSHIP1.Item("SHIP_BOL_NO") = "0000629982" Then
                                        '    .Item("PICK_QTY_CONF") = rowSOTCART2.Item("QTY_PACKED") / rowSOTINVH9.Item("RANGE_STYLE_QTY_PER_PP")
                                        'End If
                                        If rowSOTINVH9.Item("RANGE_STYLE_PP_QTY_SHIP") > 1 And rowSOTINVH9.Item("RANGE_STYLE_QTY_PER_PP") = 1 Then
                                            If CUST_CODE <> "JCPL" Then
                                                MsgBox("Range Qty problem, call Rick", vbCritical, "ASN Error")
                                                Stop
                                            End If
                                            .Item("PICK_QTY_CONF") = Int(rowSOTCART1.Item("CART_TOTAL_UNITS"))  '/ tblSOWINVH9.Fields("RANGE_STYLE_PP_QTY_SHIP"))
                                        End If

                                        If rowEDT850T2.Item("EDI_PO4_UOM") = "CA" Or (rowEDT850T2.Item("EDI_PO4_UOM") = "AS" And CUST_CODE = "BURLING") Then
                                            .Item("EDI_PO4_UOM") = UCase(rowEDT850T2.Item("EDI_PO4_UOM"))
                                            If CUST_CODE = "KMART" Then
                                                .Item("PICK_QTY_CONF") = 1
                                            End If
                                        End If

                                        If CUST_CODE = "MEIJER" Or CUST_CODE = "CHARLOT" Then
                                            .Item("EDI_PO4_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
                                            If rowSOTORDR2.Item("RANGE_STYLE_CODE") & "" <> "" Then
                                                .Item("PICK_QTY_CONF") = 1
                                            End If
                                            If .Item("EDI_PO4_UOM") = "AS" And CUST_CODE = "CHARLOT" Then
                                                If Val(rowSOTINVH9.Item("RANGE_STYLE_QTY_PER_PP") + 0) <= 1 Then
                                                    MsgBox("Range Qty problem, call Rick", vbCritical, "ASN Error")
                                                End If
                                                .Item("PICK_QTY_CONF") = Int(rowSOTCART1.Item("CART_TOTAL_UNITS")) / rowSOTINVH9.Item("RANGE_STYLE_QTY_PER_PP")
                                            End If
                                        End If
                                    Else

                                        MsgBox("Missing Range Error, call Rick", vbCritical, "ASN Error")
                                        Stop

                                    End If

                                    'this was put here to convert casepack qty to eaches... is there a way to determine if it is a casepack qty in QTY_PACKED
                                    'dynEDW856O4.Fields("PICK_QTY_CONF") = range_style_qty_pp_ship
                                    'dynEDW856O4.Fields("EDI_PO4_UOM") = "AS"

                                Else
                                    If RANGE_STYLE_CODE <> "" Then
                                        SKU = ""
                                    End If
                                    UPCXREF.Add("STYLE_CODE", rowSOTORDR2.Item("STYLE_CODE"))
                                    UPCXREF.Add("COLOR_CODE", rowSOTORDR2.Item("COLOR_CODE"))
                                    UPCXREF.Add("CODE", rowSOTORDR2.Item("CUST_COLOR_CODE"))

                                End If

                                Get_UPC(UPCXREF)

                            End If

                            If CUST_CODE = "KMART" And EDI_MERCH_TYPE <> "J1" And GTIN = "" Then ' change this in the near future to use new field in 850t2
                                Dim rowICTGTINTs() As DataRow = dst.Tables("ICTGTINT").Select("GTIN_UPC_CODE = '" & UPC & "'")
                                If rowICTGTINTs.Length > 0 Then
                                    GTIN = rowICTGTINTs(0).Item("GTIN_CODE")
                                Else
                                    ' MsgBox "Gtin Missing for Kmart, verify ASN", vbOKOnly, "ASN Error"
                                End If
                            End If

                            .Item("STYLE_UPC_CODE") = UPC
                            .Item("EDI_SKU") = SKU
                            .Item("STYLE_RETAIL") = rowICTSTYL1.Item("STYLE_RETAIL")
                            .Item("STYLE_GTIN_CODE") = GTIN
                        End With
                        dst.Tables("EDT856O4").Rows.Add(rowEDT856O4)
                    End If
                Next
            End If
        Next

    End Sub

    Sub Write_Header_856()

        Dim SHIP_ADDR_CODE As String = rowSOTSHIP1.Item("SHIP_ADDR_CODE")
        Dim rowEDTOUTB1 As DataRow = dst.Tables("EDTOUTB1").Rows.Find(New String() {SHIP_8XX_BATCH_NO, CUST_CODE, SHIP_ADDR_CODE, BILL_OF_LADING_NO})

        If rowEDTOUTB1 Is Nothing Then

            Dim rowSOTSHIPT As DataRow = Fill_Record("SOTSHIPT", New String() {rowSOTSHIP1.Item("BILL_OF_LADING_NO"), rowSOTORDR1.Item("CUST_CODE")})
            Dim TOT_CTNS As Int32 = Val(rowSOTSHIPT.Item("TOT_CTNS") & "")
            Dim TOT_WGT As Decimal = Val(rowSOTSHIPT.Item("TOT_WGT") & "")
            Dim TOT_CTNS_CONS As Int32 = Val(rowSOTSHIPT.Item("TOT_CTNS_CONS") & "")


            rowEDTOUTB1 = dst.Tables("EDTOUTB1").NewRow
            With rowEDTOUTB1
                .Item("SHIP_8XX_BATCH_NO") = SHIP_8XX_BATCH_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_DC_NO") = SHIP_ADDR_CODE
                .Item("BILL_OF_LADING_NO") = BILL_OF_LADING_NO
                .Item("FIRST_SBOL_NO") = SHIP_BOL_NO
                .Item("LAST_SBOL_NO") = SHIP_BOL_NO
                .Item("NUM_OF_856") = 1

                If SHIP_BOL_NO_CONS <> "" Then
                    .Item("SHIP_CNT_CARTONS") = TOT_CTNS_CONS
                Else
                    .Item("SHIP_CNT_CARTONS") = TOT_CTNS
                End If

                .Item("SHIP_TOTAL_WGT") = TOT_WGT
                .Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE")
            End With
            dst.Tables("EDTOUTB1").Rows.Add(rowEDTOUTB1)

            Write_EDT856O1(TOT_CTNS, TOT_WGT, TOT_CTNS_CONS)

            rowEDTOUTB1.Item("CUST_NAME") = CUST_NAME

        Else
            With rowEDTOUTB1
                If SHIP_BOL_NO > .Item("LAST_SBOL_NO") & "" Then
                    .Item("LAST_SBOL_NO") = SHIP_BOL_NO
                End If
                If SHIP_BOL_NO < .Item("FIRST_SBOL_NO") & "" _
                    Or .Item("FIRST_SBOL_NO") & "" = "" Then
                    .Item("FIRST_SBOL_NO") = SHIP_BOL_NO
                End If

                '.Item("SHIP_CNT_CARTONS") = .Item("SHIP_CNT_CARTONS") + TOT_CTNS
                '.Item("SHIP_TOTAL_WGT") = .Item("SHIP_TOTAL_WGT") + TOT_WGT
            End With
        End If

        ' rowEDTOUTB1.Item("ERRORS") = Mid$(INV_ERR, 2)
    End Sub

    Sub Write_EDT856O1(TOT_CTNS As Int32, TOT_WGT As Decimal, TOT_CTNS_CONS As Int32)

        Write_to_EDTSYSIH("OSH")

        Dim rowEDT856O1 As DataRow = dst.Tables("EDT856O1").NewRow
        With rowEDT856O1
            .Item("EDI_OUR_ID") = EDI_OUR_ID
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            .Item("SHIP_BOL_NO") = SHIP_BOL_NO
            If rowSOTORDR1.Item("FRT_TERMS") & "" = "PPD" Then
                .Item("FRT_TERMS") = "PP"
            ElseIf rowSOTORDR1.Item("FRT_TERMS") & "" = "PPA" Then
                .Item("FRT_TERMS") = "PC"
            Else
                .Item("FRT_TERMS") = "CC"
            End If
            .Item("SHIP_MANIFEST_NO") = rowSOTSHIP1.Item("SHIP_MANIFEST_NO")
            .Item("EDI_REMIT_NAME") = "" ' ????
            .Item("EDI_TP_ID") = rowEDTTRPM1_856.Item("EDI_TP_ID")  ' potential issue, good for JCP, maybe not for others with cust switches
            .Item("SHIP_856_BATCH_NO") = SHIP_856_BATCH_NO
            .Item("EDI_PRO_NO") = rowSOTSHIP1.Item("SHIP_REF")

            .Item("WHSE_ZIP_CODE") = ""
            If EDI_CUSTOMER <> "" Then
                .Item("EDI_CUSTOMER") = EDI_CUSTOMER
            Else
                .Item("EDI_CUSTOMER") = CUST_CODE
            End If
            If CUST_CODE = "SEARSCAN" Then
                .Item("EDI_TP_ID") = rowEDT850T1.Item("EDI_TP_ID")
            End If
            .Item("EDI_DATE_SHIPPED") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
            .Item("BILL_OF_LADING_NO") = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & ""
            If rowSOTSHIP1.Item("SHIP_ADDR_CODE") & "" <> "" Then
                If IsNumeric(rowSOTSHIP1.Item("SHIP_ADDR_CODE")) Then
                    .Item("SHIP_ADDR_CODE") = Format$(Val(rowSOTSHIP1.Item("SHIP_ADDR_CODE")), Mid$("000000", 1, rowEDTTRPM1_856.Item("NUMBER_CHRS_DC")))
                Else
                    .Item("SHIP_ADDR_CODE") = rowSOTSHIP1.Item("SHIP_ADDR_CODE")
                End If
            End If

            .Item("WHSE_CITY") = rowICTWHSE1.Item("WHSE_CITY")
            .Item("WHSE_STATE") = rowICTWHSE1.Item("WHSE_STATE")

            If SHIP_BOL_NO_CONS <> "" Then
                .Item("EDI_SHIP_CNT_CARTONS") = TOT_CTNS_CONS
            Else
                .Item("EDI_SHIP_CNT_CARTONS") = TOT_CTNS
            End If

            .Item("EDI_SHIP_TOTAL_WGT") = TOT_WGT

            .Item("EDI_SCAC_CODE") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & ""
            .Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & ""

            Dim EDI_SUPPLIER_NO As String = rowEDT850T1.Item("EDI_SUPPLIER_NO") & ""
            If EDI_SUPPLIER_NO <> "" Then
                .Item("EDI_SUPPLIER_NO") = EDI_SUPPLIER_NO
                If (CUST_CODE = "SEARS" Or CUST_CODE = "KMART") And Len(EDI_SUPPLIER_NO) = 9 Then
                    .Item("EDI_SUPPLIER_NO") = "00" & EDI_SUPPLIER_NO
                End If
            Else
                .Item("EDI_SUPPLIER_NO") = rowARTCUST1.Item("CUST_VEND_REF") & ""
            End If
        End With
        dst.Tables("EDT856O1").Rows.Add(rowEDT856O1)

        EDI_HL2_SEQ = 0

    End Sub

    Sub Inv_Addresses(TABLE_NAME As String, rowEDTTRPM1 As DataRow, ASN As Boolean)

        Dim rowARTCUST2 As DataRow = Nothing

        If CUST_CODE <> rowARTCUST1.Item("CUST_CODE") & "" Or CUST_NAME = "" Then
            CUST_DC_NO = ""

            rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {rowSOTINVH1.Item("CUST_CODE"), "BT", rowSOTINVH1.Item("CUST_STORE_NO")})

            If rowARTCUST2 Is Nothing Then
                CUST_NAME = rowARTCUST1.Item("CUST_NAME")
                CUST_ADDR1 = rowARTCUST1.Item("CUST_ADDR1") & ""
                CUST_ADDR2 = rowARTCUST1.Item("CUST_ADDR2") & ""
                CUST_CITY = rowARTCUST1.Item("CUST_CITY") & ""
                CUST_STATE = rowARTCUST1.Item("CUST_STATE") & ""
                CUST_ZIP_CODE = rowARTCUST1.Item("CUST_ZIP_CODE") & ""
                CUST_COUNTRY = rowARTCUST1.Item("CUST_COUNTRY") & ""
                CUST_GLN = ""
            Else
                CUST_NAME = rowARTCUST2.Item("CUST_NAME") & ""
                CUST_ADDR1 = rowARTCUST2.Item("CUST_ADDR1") & ""
                CUST_ADDR2 = rowARTCUST2.Item("CUST_ADDR2") & ""
                CUST_CITY = rowARTCUST2.Item("CUST_CITY") & ""
                CUST_STATE = rowARTCUST2.Item("CUST_STATE") & ""
                CUST_ZIP_CODE = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
                CUST_COUNTRY = rowARTCUST2.Item("CUST_COUNTRY") & ""
                CUST_GLN = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & ""
            End If
        End If

        If Len(CUST_ZIP_CODE) > 9 Then
            CUST_ZIP_CODE = Mid$(CUST_ZIP_CODE, 1, 5) & Mid$(CUST_ZIP_CODE, 7, 4)
        End If

        If ASN = False Then
            Dim row As DataRow = dst.Tables(TABLE_NAME).NewRow
            With row
                EDI_ADR_SEQ = EDI_ADR_SEQ + 1
                .Item("EDI_OUR_ID") = EDI_OUR_ID
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                .Item("EDI_ADDR_TYPE") = "BT"
                .Item("EDI_CUST_NAME_ADR") = CUST_NAME
                .Item("EDI_ADDRESS1") = CUST_ADDR1
                .Item("EDI_ADDRESS2") = CUST_ADDR2
                .Item("EDI_CITY") = CUST_CITY
                .Item("EDI_STATE") = CUST_STATE
                .Item("EDI_ZIPCODE") = CUST_ZIP_CODE
                .Item("EDI_COUNTRY") = CUST_COUNTRY
                .Item("EDI_ADDR_CODE") = CUST_GLN
                .Item("EDI_ADDR_CODE_QUAL") = rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & ""
                If EDI_HL2_SEQ <> 0 Then
                    .Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                End If
                If rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "9" Then
                    .Item("EDI_ADDR_CODE") = rowEDTTRPM1.Item("EDI_DUNS_PLUS4_PREFIX") & "0000"
                ElseIf rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "1" Then
                    .Item("EDI_ADDR_CODE") = rowEDTTRPM1.Item("EDI_DUNS_PLUS4_PREFIX") & ""
                ElseIf rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "UL" Then
                    .Item("EDI_ADDR_CODE") = CUST_GLN_DC
                Else
                    If IsNumeric(CUST_ADDR_CODE_DC) Then
                        .Item("EDI_ADDR_CODE") = Format$(Val(CUST_ADDR_CODE_DC), Mid$("000000", 1, rowEDTTRPM1.Item("NUMBER_CHRS_DC")))
                    Else
                        .Item("EDI_ADDR_CODE") = CUST_ADDR_CODE_DC
                    End If
                End If

            End With
            dst.Tables(TABLE_NAME).Rows.Add(row)
        End If

        If CUST_ADDR_CODE_MK <> rowSOTINVH1.Item("CUST_STORE_NO") And (INV_NO_CONS = "" Or ASN = True) Then
            rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {rowSOTINVH1.Item("CUST_CODE"), "MK", rowSOTINVH1.Item("CUST_STORE_NO")})
            CUST_ADDR_CODE_MK = rowARTCUST2.Item("CUST_ADDR_CODE")
            CUST_NAME_MK = rowARTCUST2.Item("CUST_NAME") & ""
            CUST_ADDR1_MK = rowARTCUST2.Item("CUST_ADDR1") & ""
            CUST_ADDR2_MK = rowARTCUST2.Item("CUST_ADDR2") & ""
            CUST_CITY_MK = rowARTCUST2.Item("CUST_CITY") & ""
            CUST_STATE_MK = rowARTCUST2.Item("CUST_STATE") & ""
            CUST_ZIP_CODE_MK = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
            CUST_COUNTRY_MK = rowARTCUST2.Item("CUST_COUNTRY") & ""
            CUST_GLN_MK = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & ""
        End If

        If ASN = True Or (rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") = "DC" And ASN = False And INV_NO_CONS = "") Then
            Dim row As DataRow = dst.Tables(TABLE_NAME).NewRow
            With row
                EDI_ADR_SEQ = EDI_ADR_SEQ + 1
                .Item("EDI_OUR_ID") = EDI_OUR_ID
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                .Item("EDI_ADDR_TYPE") = "MK"
                .Item("EDI_CUST_NAME_ADR") = Mid$(CUST_NAME_MK, 1, 34)
                .Item("EDI_ADDRESS1") = CUST_ADDR1_MK
                .Item("EDI_ADDRESS2") = CUST_ADDR2_MK
                .Item("EDI_CITY") = CUST_CITY_MK
                .Item("EDI_STATE") = CUST_STATE_MK
                If Len(CUST_ZIP_CODE_MK) = 10 Then
                    .Item("EDI_ZIPCODE") = Mid$(CUST_ZIP_CODE_MK, 1, 5) & Mid$(CUST_ZIP_CODE_MK, 7)
                Else
                    .Item("EDI_ZIPCODE") = CUST_ZIP_CODE_MK
                End If
                .Item("EDI_COUNTRY") = CUST_COUNTRY_MK
                .Item("EDI_ADDR_CODE_QUAL") = rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & ""
                If EDI_HL2_SEQ <> 0 Then
                    .Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                End If
                If rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "9" Then
                    .Item("EDI_ADDR_CODE") = rowEDTTRPM1.Item("EDI_DUNS_PLUS4_PREFIX") & Format$(Val(rowSOTINVH1.Item("CUST_STORE_NO")), "0000")
                ElseIf rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "1" Then
                    .Item("EDI_ADDR_CODE") = rowEDTTRPM1.Item("EDI_DUNS_PLUS4_PREFIX") & ""
                ElseIf rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "UL" Then
                    .Item("EDI_ADDR_CODE") = CUST_GLN_MK
                Else
                    If IsNumeric(CUST_ADDR_CODE_MK) Then
                        .Item("EDI_ADDR_CODE") = Format$(Val(CUST_ADDR_CODE_MK), Mid$("000000", 1, rowEDTTRPM1.Item("NUMBER_CHRS_DC")))
                    Else
                        .Item("EDI_ADDR_CODE") = CUST_ADDR_CODE_MK
                    End If
                End If

                'If CUST_CODE = "BURLING" And INV_NO_CONS <> "" Then
                If CUST_CODE = "BURLING" And rowSOTORDR1.Item("CUST_DC_NO") & "" <> "" Then
                    .Item("EDI_ADDR_TYPE") = "Z7"
                End If

            End With
            dst.Tables(TABLE_NAME).Rows.Add(row)
        End If

        If ASN = False Then
            Dim row As DataRow = dst.Tables(TABLE_NAME).NewRow
            With row
                EDI_ADR_SEQ = EDI_ADR_SEQ + 1
                .Item("EDI_OUR_ID") = EDI_OUR_ID
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                .Item("EDI_ADDR_TYPE") = "RE"
                .Item("EDI_CUST_NAME_ADR") = "VANDALE INDUSTRIES"
                .Item("EDI_ADDRESS1") = "16 EAST 34th STREET, 8TH FLOOR"
                .Item("EDI_CITY") = "NEW YORK"
                .Item("EDI_STATE") = "NY"
                .Item("EDI_ZIPCODE") = "10016"
                .Item("EDI_COUNTRY") = "USA"
                .Item("EDI_ADDR_CODE_QUAL") = "1"
                .Item("EDI_ADDR_CODE") = "101140010"
                If CUST_CODE = "SEARS" Or CUST_CODE = "KMART" Then
                    .Item("EDI_ADDR_CODE") = "00101140010"
                    .Item("EDI_ADDR_CODE_QUAL") = "92"
                End If
                If CUST_CODE = "MARSHAL" Then
                    .Item("EDI_CUST_NAME_ADR") = "VANDALE INDUSTRIES INC"
                    .Item("EDI_COUNTRY") = "US"
                End If
            End With
            dst.Tables(TABLE_NAME).Rows.Add(row)
        End If

    End Sub

    Sub Ship_Addresses(TABLE_NAME As String, rowEDTTRPM1 As DataRow)

        Dim rowARTCUST2 As DataRow = Nothing
        If rowSOTORDR1.Item("CUST_DC_NO") = rowSOTINVH1.Item("CUST_STORE_NO") Or rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & "" = "MK" Then
            If CUST_STORE_NO <> rowSOTINVH1.Item("CUST_STORE_NO") Then
                rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {rowSOTINVH1.Item("CUST_CODE"), "MK", rowSOTINVH1.Item("CUST_STORE_NO")})
                CUST_ADDR_CODE_MK = rowARTCUST2.Item("CUST_ADDR_CODE")
                CUST_NAME_MK = rowARTCUST2.Item("CUST_NAME")
                CUST_ADDR1_MK = rowARTCUST2.Item("CUST_ADDR1") & ""
                CUST_ADDR2_MK = rowARTCUST2.Item("CUST_ADDR2") & ""
                CUST_CITY_MK = rowARTCUST2.Item("CUST_CITY") & ""
                CUST_STATE_MK = rowARTCUST2.Item("CUST_STATE") & ""
                CUST_ZIP_CODE_MK = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
                CUST_COUNTRY_MK = rowARTCUST2.Item("CUST_COUNTRY") & ""
                CUST_GLN_MK = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & ""
            End If
            If IsNumeric(CUST_STORE_NO) Then
                CUST_DC_NO = Format$(Val(CUST_STORE_NO), Mid$("000000", 1, rowEDTTRPM1.Item("NUMBER_CHRS_DC")))
            Else
                CUST_DC_NO = CUST_STORE_NO
            End If
            CUST_ADDR_CODE_DC = CUST_ADDR_CODE_MK
            CUST_NAME_DC = CUST_NAME_MK
            CUST_ADDR1_DC = CUST_ADDR1_MK
            CUST_ADDR2_DC = CUST_ADDR2_MK
            CUST_CITY_DC = CUST_CITY_MK
            CUST_STATE_DC = CUST_STATE_MK
            CUST_ZIP_CODE_DC = CUST_ZIP_CODE_MK
            CUST_COUNTRY_DC = CUST_COUNTRY_MK
            CUST_GLN_DC = CUST_GLN_MK
        Else
            If CUST_ADDR_CODE_DC <> rowSOTORDR1.Item("CUST_DC_NO") & "" Then
                rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {rowSOTORDR1.Item("CUST_CODE"), "DC", rowSOTORDR1.Item("CUST_DC_NO")})
                If rowARTCUST2 IsNot Nothing Then
                    CUST_ADDR_CODE_DC = rowARTCUST2.Item("CUST_ADDR_CODE") & ""
                    CUST_NAME_DC = rowARTCUST2.Item("CUST_NAME") & ""
                    CUST_ADDR1_DC = rowARTCUST2.Item("CUST_ADDR1") & ""
                    CUST_ADDR2_DC = rowARTCUST2.Item("CUST_ADDR2") & ""
                    CUST_CITY_DC = rowARTCUST2.Item("CUST_CITY") & ""
                    CUST_STATE_DC = rowARTCUST2.Item("CUST_STATE") & ""
                    CUST_ZIP_CODE_DC = rowARTCUST2.Item("CUST_ZIP_CODE") & ""
                    CUST_COUNTRY_DC = rowARTCUST2.Item("CUST_COUNTRY") & ""
                    CUST_GLN_DC = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & ""
                End If
            End If
        End If

        Dim rowADDRESSES As DataRow = Nothing

        rowADDRESSES = dst.Tables(TABLE_NAME).NewRow()
        With rowADDRESSES
            EDI_ADR_SEQ = EDI_ADR_SEQ + 1
            .Item("EDI_OUR_ID") = EDI_OUR_ID
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
            .Item("EDI_ADDR_TYPE") = "ST"
            .Item("EDI_CUST_NAME_ADR") = Mid$(CUST_NAME_DC, 1, 34)
            .Item("EDI_ADDRESS1") = CUST_ADDR1_DC
            .Item("EDI_ADDRESS2") = CUST_ADDR2_DC
            .Item("EDI_CITY") = CUST_CITY_DC
            .Item("EDI_STATE") = CUST_STATE_DC
            If Len(CUST_ZIP_CODE_DC) = 10 Then
                .Item("EDI_ZIPCODE") = Mid$(CUST_ZIP_CODE_DC, 1, 5) & Mid$(CUST_ZIP_CODE_DC, 7, 4)
            Else
                .Item("EDI_ZIPCODE") = CUST_ZIP_CODE_DC
            End If
            .Item("EDI_COUNTRY") = CUST_COUNTRY_DC
            .Item("EDI_ADDR_CODE_QUAL") = rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & ""
            If EDI_HL2_SEQ <> 0 Then
                .Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
            End If
            If rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "9" Then
                .Item("EDI_ADDR_CODE") = rowEDTTRPM1.Item("EDI_DUNS_PLUS4_PREFIX") & Format$(Val(rowSOTSHIP1.Item("SHIP_ADDR_CODE")), "0000")
            ElseIf rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "1" Then
                .Item("EDI_ADDR_CODE") = rowEDTTRPM1.Item("EDI_DUNS_PLUS4_PREFIX") & ""
            ElseIf rowEDTTRPM1.Item("DEFAULT_ADDR_QUALIFIER") & "" = "UL" Then
                .Item("EDI_ADDR_CODE") = CUST_GLN_DC
            Else
                If IsNumeric(CUST_ADDR_CODE_DC) Then
                    .Item("EDI_ADDR_CODE") = Format$(Val(CUST_ADDR_CODE_DC), Mid$("000000", 1, rowEDTTRPM1.Item("NUMBER_CHRS_DC")))
                Else
                    .Item("EDI_ADDR_CODE") = CUST_ADDR_CODE_DC
                End If
            End If
        End With
        dst.Tables(TABLE_NAME).Rows.Add(rowADDRESSES)

        rowADDRESSES = dst.Tables(TABLE_NAME).NewRow()
        With rowADDRESSES
            EDI_ADR_SEQ = EDI_ADR_SEQ + 1
            .Item("EDI_OUR_ID") = EDI_OUR_ID
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            If EDI_HL2_SEQ <> 0 Then
                .Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
            End If
            .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
            .Item("EDI_ADDR_TYPE") = "SF"
            .Item("EDI_CUST_NAME_ADR") = rowICTWHSE1.Item("WHSE_DESC") ' "VANDALE WAREHOUSE"
            .Item("EDI_ADDRESS1") = rowICTWHSE1.Item("WHSE_ADDR1") ' "40 EXECUTIVE AVENUE"
            .Item("EDI_CITY") = rowICTWHSE1.Item("WHSE_CITY") ' "EDISON"
            .Item("EDI_STATE") = rowICTWHSE1.Item("WHSE_STATE") ' "NJ"
            .Item("EDI_ZIPCODE") = rowICTWHSE1.Item("WHSE_ZIP_CODE") '"08817"
            .Item("EDI_COUNTRY") = rowICTWHSE1.Item("WHSE_COUNTRY") ' "USA"
            .Item("EDI_ADDR_CODE_QUAL") = "1"
            .Item("EDI_ADDR_CODE") = "101140010"
            If CUST_CODE = "SEARS" Or CUST_CODE = "KMART" Then
                .Item("EDI_ADDR_CODE") = "00101140010"
                .Item("EDI_ADDR_CODE_QUAL") = "92"
            End If

        End With
        dst.Tables(TABLE_NAME).Rows.Add(rowADDRESSES)

    End Sub

    Sub Write_to_EDTSYSIH(EDI_APPLICATION_ID As String)

        EDI_OUTBOUND_DOC_NO = ASCMAIN1.Next_Control_No("EDI_OUTBOUND_DOC_NO", 10)

        Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
        With rowEDTSYSIH
            .Item("EDI_OUR_ID") = EDI_OUR_ID
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            .Item("EDI_APPLICATION_ID") = EDI_APPLICATION_ID
            .Item("EDI_PROCESS_IND") = "1"
            .Item("EDI_TP_ID") = rowEDT850T1.Item("EDI_TP_ID")
            If CUST_CODE = "KMART10" Then
                .Item("EDI_TP_ID") = "006985290AC"
            End If
        End With
        dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

        EDI_ADR_SEQ = 0
    End Sub

    Sub Load_Customer_Related_Rows()
        Fill_Records("SOTCSTY1", CUST_CODE)
        dst.Tables("SOTCSTYX").Rows.Clear()


        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        CUST_NAME = rowARTCUST1.Item("CUST_NAME")

        Fill_Records("ARTCUST2", CUST_CODE)

        If CUST_CODE = "KMART" Then
            Fill_Records("ICTGTINT")
        End If
    End Sub

    Function Get_UPC(UPCXREF As Dictionary(Of String, String)) As DataRow

        Dim STYLE_CODE As String = UPCXREF("STYLE_CODE")
        Dim COLOR_CODE As String = UPCXREF("COLOR_CODE")

        ' Look into SOTCSTY1 first
        ' if you can find UPC/SKU there, then that is your match (however, this search is done using STYLE_CODE and COLOR_CODE
        ' but there there are duplicates, so it is a mystery how this is working
        ' then, if UPC is still = "", look at ICTSTYC4 using STYLE_CODE and COLOR_CODE - and take the last UPC you find for that STYLE/COLOR
        ' but there there are duplicates based on COLOR_CODE_UPC and SIZE_INDEX, so it is a mystery how this is working
        ' then, if UPC is still = "", look at ICTSTYC2 using STYLE_CODE and COLOR_CODE - and take the last UPC you find for that STYLE/COLOR
        ' (original v1 code only used STYLE_CODE and not COLOR_CODE when searching UPCs - WTF?) 
        ' but there there are duplicates based on COLOR_CODE_UPC, so it is a mystery how this is working

        Dim sqlSC As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

        Dim rowSOTCSTYX As DataRow = dst.Tables("SOTCSTYX").Rows.Find(New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
        If rowSOTCSTYX Is Nothing Then
            ' SOTCSTY1 is loaded for all records keyed to the customer, at the customer break in the main transactional loop
            For Each row As DataRow In dst.Tables("SOTCSTY1").Select(sqlSC, "STYLE_CODE,COLOR_CODE")
                Dim CUST_STYLE_UPC_FLAG As String = row.Item("CUST_STYLE_UPC_FLAG") & ""
                If UPC = "" And CUST_STYLE_UPC_FLAG = "1" Then
                    UPC = row.Item("CUST_STYLE_CODE")
                ElseIf SKU = "" And CUST_STYLE_UPC_FLAG <> "1" Then
                    SKU = row.Item("CUST_STYLE_CODE")
                End If

                If UPC <> "" And SKU <> "" Then Exit For
            Next
            rowSOTCSTYX = dst.Tables("SOTCSTYX").Rows.Add(New String() {CUST_CODE, STYLE_CODE, COLOR_CODE, UPC, SKU})
        End If

        UPC = rowSOTCSTYX.Item("UPC") & ""
        SKU = rowSOTCSTYX.Item("SKU") & ""

        If UPC = "" Then
            For Each row As DataRow In dst.Tables("ICTSTYC4").Select(sqlSC & " and ISNULL(UPC_CODE,'?') <> '?'", "STYLE_CODE,COLOR_CODE")
                UPC = row.Item("UPC_CODE")
            Next
        End If

        If UPC = "" Then
            For Each row As DataRow In dst.Tables("ICTSTYC2").Select(sqlSC & " and ISNULL(UPC_CODE,'?') <> '?'", "STYLE_CODE,COLOR_CODE")
                UPC = row.Item("UPC_CODE")
            Next
        End If

        ' Should probably stuff UPC back into SOTCSTYX but what about COLOR_CODE_UPC and SIZE_INDEX?
        Return rowSOTCSTYX
    End Function
End Class