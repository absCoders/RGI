Public Class SORNEWC1

    Dim SO_ORDER_NO As String
    Dim ARTCUST1_COLS As String() = {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE",
        "CUST_ZIP_CODE", "CUST_COUNTRY", "CUST_CONTACT", "CUST_DBA_NAME", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL",
        "CUST_INCL_INV_SHIP", "CUST_XMIT_INV_VIA", "CUST_STMT_EMAIL", "CUST_STMT_CC", "CUST_INV_EMAIL", "CUST_INV_CC",
        "CUST_INV_COMMENT", "CUST_ROUTING_INST", "CUST_SPECIAL_INST", "FRT_TERMS", "SHIP_VIA_CODE", "WHSE_CODE",
        "CUST_SHIP_TO_CODE_DEF", "CUST_EMAIL_SHIP_ACK", "CUST_PRICE_TIER", "CUST_DISC_PCT", "CUST_PRICE_TIER_PVC", "CUST_DISC_PCT_EXTRA"}
    Dim ARTCUST2_COLS As String() = {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE",
        "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_ADDR_NAME", "CUST_ADDR_STATUS",
        "CUST_EMAIL", "GLOBAL_LOCATION_NUMBER", "FDX_ACCT_NO", "CUST_DC_NO", "CUST_ADDR3", "UPS_ACCT_NO", "CUST_ADDR_GROUP", "CUST_EMAIL"}
    Dim ARTCUSTD_COLS As String() = {"CONTACT_NAME", "CONTACT_TITLE", "CONTACT_EMAIL", "CONTACT_PHONE", "CONTACT_EXT",
        "CONTACT_FAX", "CONTACT_TYPE", "CONTACT_PRIMARY", "CONTACT_NOTE", "CONTACT_CELL"}
    Dim ARTCUSTQ_COLS As String() = {"LAST_DATE", "LAST_OPER", "LAST_ORDR_NO", "RESIDENTIAL_ORDR", "INSIDE_REQ", "GATE_LIFT_REQ",
        "LIMITED_ACCESS", "LIMITED_ACCESS_NOTE", "IRREGULAR_HOURS", "IRREGULAR_HOURS_NOTE", "APPOINTMENT_REQUIRED",
        "APPOINTMENT_REQUIRED_NOTE", "BROKER", "BROKER_NOTE"}
    Dim S As New System.Text.StringBuilder With {.Length = 0}


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R" ' Y MEANS REPORT WITH UPDATE AN N IS REPORT ONLY A 'U' IS UPDATE ONLY 
        '  FROM NEEDS RGO. TO PREFIX SOTORDR1 & SOTORDR2, SOTORDR5 - POSSIBLY ALL TABLES 
        ' **** REMEMBER RGO. TO UPDATE ROUTINE 
        ' ALSO REPLACE WHERE CLAUSE TO GET ORDR_STATUS = 'L' INSTEAD OF ORDR_NO & REMOVE ALL OTHER STUFF 
        ' & "WHERE  SOTORDR1.ORDR_NO = '0000362015'  " & vbCrLf _
        ' REMOVE STU

        'WR - Putting a finger in the dike until I can find the source of the hole.
        S.Length = 0
        S.AppendLine("BEGIN DECLARE CURSOR C1 IS")
        S.AppendLine("  SELECT ORDR_NO, LAST_OPER")
        S.AppendLine("  FROM SOTORDR1_L")
        S.AppendLine("  WHERE ORDR_NO IN")
        S.AppendLine("  (")
        S.AppendLine("    SELECT LAST_ORDR_NO FROM ARTCUSTQ_L")
        S.AppendLine("    WHERE NVL(LAST_OPER,'NULL') = 'NULL'")
        S.AppendLine("  );")
        S.AppendLine("BEGIN FOR R1 IN C1 LOOP")
        S.AppendLine("  UPDATE ARTCUSTQ_L")
        S.AppendLine("  SET LAST_OPER = R1.LAST_OPER")
        S.AppendLine("  WHERE LAST_ORDR_NO = R1.ORDR_NO;")
        S.AppendLine("END LOOP; END; END;")
        ASCMAIN1.sql = S.ToString
        ASCDATA1.ExecuteSQL()
        S.Length = 0
        S.AppendLine("UPDATE ARTCUSTQ_L")
        S.AppendLine("SET INSIDE_REQ = '0'")
        S.AppendLine("WHERE NVL(INSIDE_REQ,'X') = 'X'")
        ASCMAIN1.sql = S.ToString
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ARTCUST1_L " _
        & " Set LAST_OPER = '" & XNO & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ARTCUST2_L " _
        & " Set LAST_OPER = '" & XNO & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ARTCUSTD_L " _
        & " Set LAST_OPER = '" & XNO & "'"
        ASCDATA1.ExecuteSQL()

        Create_TDA(dst.Tables.Add(), "ARTCUST1_A", "*")
        ASCMAIN1.sql = "Select ARTCUST1_L.* FROM ARTCUST1_L WHERE LAST_OPER = '" & XNO & "'"
        Fill_Records("ARTCUST1_A", "", True, ASCMAIN1.sql)
        dst.Tables("ARTCUST1_A").Columns.Add("CHANGE_TYPE")

        Create_TDA(dst.Tables.Add(), "ARTCUST2_A", "*")
        ASCMAIN1.sql = "Select ARTCUST2_L.* FROM ARTCUST2_L WHERE LAST_OPER = '" & XNO & "'"
        Fill_Records("ARTCUST2_A", "", True, ASCMAIN1.sql)

        Create_TDA(dst.Tables.Add(), "ARTCUSTD_A", "*")
        ASCMAIN1.sql = "Select ARTCUSTD_L.* FROM ARTCUSTD_L WHERE LAST_OPER = '" & XNO & "'"
        Fill_Records("ARTCUSTD_A", "", True, ASCMAIN1.sql)

        Create_TDA(dst.Tables.Add(), "ARTCUST1", "*")
        ASCMAIN1.sql = "Select ARTCUST1.* FROM ARTCUST1 WHERE CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1_L WHERE LAST_OPER = '" & XNO & "')"
        Fill_Records("ARTCUST1", "", True, ASCMAIN1.sql)

        Create_TDA(dst.Tables.Add(), "ARTCUST2", "*")
        ASCMAIN1.sql = "Select ARTCUST2.* FROM ARTCUST2 WHERE (CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE) IN (SELECT ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_TYPE, ARTCUST2.CUST_ADDR_CODE from artcust2, artcust2_l where(artcust2.cust_code = ARTCUST2_L.CUST_CODE) and  artcust2.cust_addr_type = artcust2_l.cust_addr_type and artcust2.cust_addr_code = artcust2_l.cust_addr_code and ARTCUST2_L.LAST_OPER = '" & XNO & "')"
        Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)

        Create_TDA(dst.Tables.Add(), "ARTCUSTD", "*")
        ASCMAIN1.sql = "Select ARTCUSTD.* FROM ARTCUSTD WHERE (CUST_CODE, CONTACT_NO) IN (SELECT ARTCUSTD.CUST_CODE, ARTCUSTD.CONTACT_NO from artcustD, artcustD_l where(artcustD.cust_code = ARTCUSTD_L.CUST_CODE) AND  ARTCUSTD.CONTACT_NO = ARTCUSTD_l.CONTACT_NO AND ARTCUSTD_L.LAST_OPER = '" & XNO & "')"
        Fill_Records("ARTCUSTD", "", True, ASCMAIN1.sql)

        Create_TDA(dst.Tables.Add(), "ASTAUDT1", "*")

        Create_TDA(dst.Tables.Add(), "ARTCUSTQ_L", "*")
        S.Length = 0
        S.AppendLine("SELECT *")
        S.AppendLine("FROM ARTCUSTQ_L")
        S.AppendLine("WHERE")
        S.AppendLine("(CUST_CODE, CUST_ADDR_CODE, LAST_DATE)")
        S.AppendLine("NOT IN")
        S.AppendLine("(")
        S.AppendLine("SELECT CUST_CODE, CUST_ADDR_CODE, LAST_DATE FROM ARTCUSTQ")
        S.AppendLine(")")
        ASCMAIN1.sql = S.ToString
        Fill_Records("ARTCUSTQ_L", "", False, ASCMAIN1.sql)

        Create_TDA(dst.Tables.Add(), "ARTCUSTQ", "*")
        S.Length = 0
        S.AppendLine("SELECT *")
        S.AppendLine("FROM ARTCUSTQ")
        S.AppendLine("WHERE (CUST_CODE, CUST_ADDR_CODE) IN")
        S.AppendLine("(")
        S.AppendLine(" SELECT CUST_CODE, CUST_ADDR_CODE")
        S.AppendLine(" FROM ARTCUSTQ_L")
        S.AppendLine(" WHERE")
        S.AppendLine(" (CUST_CODE, CUST_ADDR_CODE, LAST_DATE)")
        S.AppendLine(" NOT IN")
        S.AppendLine(" (")
        S.AppendLine("  SELECT CUST_CODE, CUST_ADDR_CODE, LAST_DATE FROM ARTCUSTQ")
        S.AppendLine(" )")
        S.AppendLine(")")
        ASCMAIN1.sql = S.ToString
        dst.Tables("ARTCUSTQ").Columns.Add("CHANGE_TYPE")
        dst.Tables("ARTCUSTQ").Columns.Add("CUST_NAME")
        dst.Tables("ARTCUSTQ").Columns.Add("CUST_NAME_ST")
        Fill_Records("ARTCUSTQ", "", True, ASCMAIN1.sql)

        AddAuditRecords()

        For Each rowARTCUST1_A As DataRow In dst.Tables("ARTCUST1_A").Select()
            'Per Danny we are now stamping new customers INIT_DATE with the date they are imported and The person Who Imported It. - WR 2/22/18
            Dim CUST_CODE As String = rowARTCUST1_A.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find({CUST_CODE})
            If rowARTCUST1 Is Nothing Then
                Dim row As DataRow = dst.Tables("ARTCUST1").NewRow
                rowARTCUST1_A.Item("CHANGE_TYPE") = "New Customers"
                row.Item("CUST_CODE") = rowARTCUST1_A.Item("CUST_CODE")
                For Each COL_NAME As String In ARTCUST1_COLS
                    row.Item(COL_NAME) = rowARTCUST1_A.Item(COL_NAME)
                Next

                'These get set to defaults no matter what.
                row.Item("FRT_TERMS") = "COL"
                row.Item("SHIP_VIA_CODE") = "BST"
                row.Item("CUST_ALLOW_BACKORDER") = "1"
                row.Item("POST_CODE") = "REG"
                row.Item("CUST_STATUS") = "A"
                row.Item("CURR_CODE") = "USD"
                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                row.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                row.Item("LAST_OPER") = rowARTCUST1_A.Item("LAST_OPER")
                row.Item("LAST_DATE") = rowARTCUST1_A.Item("LAST_DATE")

                'These were only being filled in new records.  I have no idea why and don't have the time to figure it out.
                row.Item("CUST_DUNS") = rowARTCUST1_A.Item("CUST_DUNS")
                row.Item("CUST_VEND_REF") = rowARTCUST1_A.Item("CUST_VEND_REF")
                row.Item("CUST_ALT_SORT") = rowARTCUST1_A.Item("CUST_ALT_SORT")
                row.Item("CUST_STMT_IND") = rowARTCUST1_A.Item("CUST_STMT_IND")
                row.Item("CUST_SALES_HOLD") = rowARTCUST1_A.Item("CUST_SALES_HOLD")
                row.Item("TERM_CODE") = rowARTCUST1_A.Item("TERM_CODE")
                row.Item("SREP_CODE") = rowARTCUST1_A.Item("SREP_CODE")
                row.Item("SREP2_CODE") = rowARTCUST1_A.Item("SREP2_CODE")
                row.Item("CUST_CLASS_CODE") = rowARTCUST1_A.Item("CUST_CLASS_CODE")
                row.Item("CUST_PO_REQD") = rowARTCUST1_A.Item("CUST_PO_REQD")
                row.Item("TRADE_CLASS_CODE") = rowARTCUST1_A.Item("TRADE_CLASS_CODE")
                row.Item("CUST_STATUS_DATE") = DATETIME_STAMP.Date
                row.Item("CUST_SHIP_COMPLETE") = rowARTCUST1_A.Item("CUST_SHIP_COMPLETE")
                row.Item("CUST_BILL_SHIP_TO") = rowARTCUST1_A.Item("CUST_BILL_SHIP_TO")
                row.Item("FDX_ACCT_NO") = rowARTCUST1_A.Item("FDX_ACCT_NO")
                row.Item("UPS_ACCT_NO") = rowARTCUST1_A.Item("UPS_ACCT_NO")
                row.Item("CUST_SHIP_BY_CASE") = rowARTCUST1_A.Item("CUST_SHIP_BY_CASE")
                row.Item("CUST_STATUS_COMMENT") = rowARTCUST1_A.Item("CUST_STATUS_COMMENT")
                row.Item("CUST_URL") = rowARTCUST1_A.Item("CUST_URL")
                row.Item("CUST_ADDR3") = rowARTCUST1_A.Item("CUST_ADDR3")
                row.Item("CUST_SHIP_TO_REQD") = rowARTCUST1_A.Item("CUST_SHIP_TO_REQD")
                row.Item("VEND_CODE") = rowARTCUST1_A.Item("VEND_CODE")
                row.Item("CUST_BOL_INST") = rowARTCUST1_A.Item("CUST_BOL_INST")
                dst.Tables("ARTCUST1").Rows.Add(row)
            Else
                rowARTCUST1_A.Item("CHANGE_TYPE") = "Existing Customers"
                For Each COL_NAME As String In ARTCUST1_COLS
                    rowARTCUST1.Item(COL_NAME) = rowARTCUST1_A.Item(COL_NAME)
                Next
            End If
        Next

        For Each rowARTCUST2_A As DataRow In dst.Tables("ARTCUST2_A").Select()
            'Per Danny we are now stamping new customers INIT_DATE with the date they are imported and The person Who Imported It. - WR 2/22/18
            Dim CUST_CODE As String = rowARTCUST2_A.Item("CUST_CODE")
            Dim CUST_ADDR_TYPE As String = rowARTCUST2_A.Item("CUST_ADDR_TYPE")
            Dim CUST_ADDR_CODE As String = rowARTCUST2_A.Item("CUST_ADDR_CODE")
            Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})
            If rowARTCUST2 Is Nothing Then
                Dim row As DataRow = dst.Tables("ARTCUST2").NewRow
                row.Item("CUST_CODE") = rowARTCUST2_A.Item("CUST_CODE")
                row.Item("CUST_ADDR_TYPE") = rowARTCUST2_A.Item("CUST_ADDR_TYPE")
                row.Item("CUST_ADDR_CODE") = rowARTCUST2_A.Item("CUST_ADDR_CODE")
                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                row.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                row.Item("LAST_OPER") = rowARTCUST2_A.Item("LAST_OPER")
                row.Item("LAST_DATE") = rowARTCUST2_A.Item("LAST_DATE")
                For Each COL_NAME As String In ARTCUST2_COLS
                    row.Item(COL_NAME) = rowARTCUST2_A.Item(COL_NAME)
                Next
                dst.Tables("ARTCUST2").Rows.Add(row)
            Else
                For Each COL_NAME As String In ARTCUST2_COLS
                    rowARTCUST2.Item(COL_NAME) = rowARTCUST2_A.Item(COL_NAME)
                Next
            End If
        Next

        For Each rowARTCUSTD_A As DataRow In dst.Tables("ARTCUSTD_A").Select()
            'Per Danny we are now stamping new customers INIT_DATE with the date they are imported and The person Who Imported It. - WR 2/22/18
            Dim CUST_CODE As String = rowARTCUSTD_A.Item("CUST_CODE")
            Dim CONTACT_NO As String = rowARTCUSTD_A.Item("CONTACT_NO")

            Dim rowARTCUSTD As DataRow = dst.Tables("ARTCUSTD").Rows.Find(New Object() {CUST_CODE, CONTACT_NO})
            If rowARTCUSTD Is Nothing Then
                Dim row As DataRow = dst.Tables("ARTCUSTD").NewRow
                row.Item("CUST_CODE") = rowARTCUSTD_A.Item("CUST_CODE")
                row.Item("CONTACT_NO") = rowARTCUSTD_A.Item("CONTACT_NO")
                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                row.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                row.Item("LAST_DATE") = rowARTCUSTD_A.Item("LAST_DATE")
                row.Item("LAST_OPER") = rowARTCUSTD_A.Item("LAST_OPER")
                For Each COL_NAME As String In ARTCUSTD_COLS
                    row.Item(COL_NAME) = rowARTCUSTD_A.Item(COL_NAME)
                Next
                dst.Tables("ARTCUSTD").Rows.Add(row)
            Else
                rowARTCUSTD("INIT_OPER") = rowARTCUSTD_A.Item("INIT_OPER")
                rowARTCUSTD("LAST_DATE") = rowARTCUSTD_A.Item("LAST_DATE")
                rowARTCUSTD("LAST_OPER") = rowARTCUSTD_A.Item("LAST_OPER")
                rowARTCUSTD("INIT_DATE") = rowARTCUSTD_A.Item("INIT_DATE")
                For Each COL_NAME As String In ARTCUSTD_COLS
                    rowARTCUSTD.Item(COL_NAME) = rowARTCUSTD_A.Item(COL_NAME)
                Next
            End If
        Next

        For Each rowARTCUSTQ_L As DataRow In dst.Tables("ARTCUSTQ_L").Select()
            Dim CUST_CODE As String = rowARTCUSTQ_L.Item("CUST_CODE")
            Dim CUST_ADDR_CODE As String = rowARTCUSTQ_L.Item("CUST_ADDR_CODE")
            Dim rowARTCUSTQ As DataRow = dst.Tables("ARTCUSTQ").Rows.Find(New Object() {CUST_CODE, CUST_ADDR_CODE})
            If rowARTCUSTQ Is Nothing Then
                Dim row As DataRow = dst.Tables("ARTCUSTQ").NewRow
                row.Item("CHANGE_TYPE") = "New Records"
                row.Item("CUST_CODE") = rowARTCUSTQ_L.Item("CUST_CODE")
                row.Item("CUST_ADDR_CODE") = rowARTCUSTQ_L.Item("CUST_ADDR_CODE")
                For Each COL_NAME As String In ARTCUSTQ_COLS
                    row.Item(COL_NAME) = rowARTCUSTQ_L.Item(COL_NAME)
                Next
                dst.Tables("ARTCUSTQ").Rows.Add(row)
            Else
                rowARTCUSTQ.Item("CHANGE_TYPE") = "Updated Records"
                For Each COL_NAME As String In ARTCUSTQ_COLS
                    rowARTCUSTQ.Item(COL_NAME) = rowARTCUSTQ_L.Item(COL_NAME)
                Next
            End If
        Next
        For Each rowARTCUSTQ As DataRow In dst.Tables("ARTCUSTQ").Select()
            Dim CUST_CODE As String = rowARTCUSTQ.Item("CUST_CODE").ToString & String.Empty
            Dim CUST_ADDR_CODE As String = rowARTCUSTQ.Item("CUST_ADDR_CODE").ToString & String.Empty
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_ADDR_CODE})
            If Not IsNothing(rowARTCUST1) Then
                rowARTCUSTQ.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME").ToString & String.Empty
            End If
            If Not IsNothing(rowARTCUST2) Then
                rowARTCUSTQ.Item("CUST_NAME_ST") = rowARTCUST2.Item("CUST_NAME").ToString & String.Empty
            End If

        Next

    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report("SORNEWC1", "Customers from Laptops")
        Generate_Report("SORNEWC2", "Customer Ship Tos from Laptops")
        Generate_Report("SORNEWCD", "Customer Contacts from Laptops")
        Generate_Report("SORNEWC3", "Customer Change Audit Trail")
        Generate_Report("SORNEWC4", "Customer Extra Ship To Changes")
        Generate_Report("SORNEWC5", "Additional Ship-to Information")
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()
        dst.Tables("ARTCUST1_A").RejectChanges()
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUSTD"}
            For Each ROW As DataRow In dst.Tables(TABLE_NAME & "_A").Select("")
                ROW.SetAdded()
            Next
            Update_Record_TDA(TABLE_NAME)
            Update_Record_TDA(TABLE_NAME & "_A")
            ASCMAIN1.sql = "DELETE " & TABLE_NAME & "_L" _
            & " WHERE LAST_OPER = '" & XNO & "'"
            ASCDATA1.ExecuteSQL()
        Next
        Update_Record_TDA("ARTCUSTQ")
        Update_Record_TDA("ASTAUDT1")
    End Sub

    Private Sub AddAuditRecords()
        For Each TABLE_NAME_FROM As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUSTD", "ARTCUSTQ"}
            Dim TABLE_NAME_TO As String = String.Format("{0}_A", TABLE_NAME_FROM)
            If TABLE_NAME_FROM = "ARTCUSTQ" Then
                TABLE_NAME_TO = String.Format("{0}_L", TABLE_NAME_FROM)
            End If
            For Each rowTABLE_TO As DataRow In dst.Tables(TABLE_NAME_TO).Select()
                Dim CUST_CODE As String = rowTABLE_TO.Item("CUST_CODE").ToString & String.Empty
                Dim FILTER As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
                Select Case TABLE_NAME_FROM
                    Case "ARTCUST2"
                        Dim CUST_ADDR_TYPE As String = rowTABLE_TO.Item("CUST_ADDR_TYPE").ToString & String.Empty
                        Dim CUST_ADDR_CODE As String = rowTABLE_TO.Item("CUST_ADDR_CODE").ToString & String.Empty
                        FILTER = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_TYPE = '{1}' AND CUST_ADDR_CODE = '{2}'", CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE)
                    Case "ARTCUSTD"
                        Dim CONTACT_NO As String = rowTABLE_TO.Item("CONTACT_NO").ToString & String.Empty
                        FILTER = String.Format("CUST_CODE = '{0}' AND CONTACT_NO = '{1}'", CUST_CODE, CONTACT_NO)
                    Case "ARTCUSTQ"
                        Dim CUST_ADDR_CODE As String = rowTABLE_TO.Item("CUST_ADDR_CODE").ToString & String.Empty
                        FILTER = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CUST_ADDR_CODE)
                End Select

                Dim rowTABLE_FROM As DataRow = dst.Tables(TABLE_NAME_FROM).Select(FILTER).FirstOrDefault
                If Not IsNothing(rowTABLE_FROM) Then
                    Dim FILTERC As String = String.Format("CUST_CODE = '{0}'", CUST_CODE)
                    Dim BASE_TABLE As String = "ARTCUST1_A"
                    If TABLE_NAME_FROM = "ARTCUSTQ" Then
                        BASE_TABLE = "ARTCUSTQ"
                    End If
                    Dim CHANGE_TYPE As String = dst.Tables.Item(BASE_TABLE).Select(FILTERC).FirstOrDefault.Item("CHANGE_TYPE").ToString & String.Empty
                    If CHANGE_TYPE = "" Then
                        For Each DC As DataColumn In dst.Tables.Item(TABLE_NAME_TO).Columns
                            Dim COL_NAME As String = DC.ColumnName
                            Dim VALID_COL As Boolean = False
                            If TABLE_NAME_FROM = "ARTCUST1" And ARTCUST1_COLS.Contains(COL_NAME) Then
                                VALID_COL = True
                            End If
                            If TABLE_NAME_FROM = "ARTCUST2" And ARTCUST2_COLS.Contains(COL_NAME) Then
                                VALID_COL = True
                            End If
                            If TABLE_NAME_FROM = "ARTCUSTD" And ARTCUSTD_COLS.Contains(COL_NAME) Then
                                VALID_COL = True
                            End If
                            If TABLE_NAME_FROM = "ARTCUSTQ" And ARTCUSTQ_COLS.Contains(COL_NAME) Then
                                VALID_COL = True
                            End If
                            If VALID_COL Then
                                If dst.Tables(TABLE_NAME_FROM).Columns.Contains(COL_NAME) Then
                                    If rowTABLE_TO.Item(COL_NAME).ToString & String.Empty <> rowTABLE_FROM.Item(COL_NAME).ToString & String.Empty Then
                                        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                                        With rowASTAUDT1
                                            If TABLE_NAME_FROM = "ARTCUSTQ" Then
                                                .Item("TABLE_NAME") = "ARTCUSTQ"
                                            Else
                                                .Item("TABLE_NAME") = "ARTCUST1"
                                            End If
                                            .Item("KEY_VALUE") = CUST_CODE
                                            Select Case TABLE_NAME_FROM
                                                Case "ARTCUST2"
                                                    .Item("KEY_VALUE2") = rowTABLE_FROM.Item("CUST_ADDR_TYPE").ToString & String.Empty & "-" & rowTABLE_FROM.Item("CUST_ADDR_CODE").ToString & String.Empty
                                                Case "ARTCUSTD"
                                                    .Item("KEY_VALUE2") = rowTABLE_FROM.Item("CONTACT_NO").ToString & String.Empty
                                                Case "ARTCUSTQ"
                                                    .Item("KEY_VALUE2") = rowTABLE_FROM.Item("CUST_ADDR_CODE").ToString & String.Empty & "-" & rowTABLE_FROM.Item("CUST_ADDR_CODE").ToString & String.Empty
                                            End Select
                                            .Item("COLUMN_NAME") = COL_NAME
                                            .Item("USER_ID") = ASCMAIN1.USER_ID
                                            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                                            .Item("OLD_VALUE") = rowTABLE_FROM.Item(COL_NAME).ToString & String.Empty
                                            .Item("NEW_VALUE") = rowTABLE_TO.Item(COL_NAME).ToString & String.Empty
                                            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                                            .Item("SELECTION_NO") = Me.SELECTION_NO
                                            .Item("XNO") = Me.XNO
                                        End With
                                        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        Next
    End Sub
End Class