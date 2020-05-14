Imports ABSolution
Imports Infragistics.Win

Public Class EDR754I1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("EDTPARM1")

        ASCMAIN1.sql = "Select count(*) from gen.EDT754I1 where EDI_PROCESS_IND is Null"
        Dim rows As String = ASCDATA1.GetDataValue(ASCMAIN1.sql)

        If rows = "0" Then
            txtDescription.Text = "No 754 Records Available"
        Else
            txtDescription.Text = "754 Records Available"
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()
        ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options


        ' Extracts from Data Sources
        
        BeginTrans()

        ASCMAIN1.sql = " Insert into " & ASCMAIN1.CLIENT & ".EDT754I1 Select * from gen.EDT754I1 where EDI_PROCESS_IND is Null"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = " Insert into " & ASCMAIN1.CLIENT & ".EDT754I2 "
        ASCMAIN1.sql = ASCMAIN1.sql & " Select * from gen.EDT754I2 where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from gen.EDT754I1 where EDI_PROCESS_IND is Null)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = " Insert into " & ASCMAIN1.CLIENT & ".EDT754I3 "
        ASCMAIN1.sql = ASCMAIN1.sql & " Select * from gen.EDT754I3 where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from gen.EDT754I1 where EDI_PROCESS_IND is Null)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = " Insert into " & ASCMAIN1.CLIENT & ".EDT754I9 "
        ASCMAIN1.sql = ASCMAIN1.sql & " Select * from gen.EDT754I9 where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from gen.EDT754I1 where EDI_PROCESS_IND is Null)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Update gen.EDT754I1 Set EDI_PROCESS_IND = '1'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        CommitTrans()

        ASCMAIN1.sql = "Select * from EDT754I1 where EDI_PROCESS_IND is Null"
        Create_TDA(dst.Tables.Add, "EDT754I1", "**", 0, True, "", 1, "EDI_PROCESS_IND")
        Fill_Records("EDT754I1")

        ASCMAIN1.sql = "Select * from EDT754I2 where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from EDT754I1 where EDI_PROCESS_IND is Null)"
        Create_TDA(dst.Tables.Add, "EDT754I2", "**", 0, True, "", 2)
        Fill_Records("EDT754I2")

        ASCMAIN1.sql = "Select * from EDT754I3 where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from EDT754I1 where EDI_PROCESS_IND is Null)"
        Create_TDA(dst.Tables.Add, "EDT754I3", "**", 0, True, "", 3)
        Fill_Records("EDT754I3")

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_CUST_PO in (Select EDI_PO_NO from EDT754I1, EDT754I3 where EDT754I1.EDI_DOC_SEQ_NO = EDT754I3.EDI_DOC_SEQ_NO and EDI_PROCESS_IND is Null)"
        Create_TDA(dst.Tables.Add, "SOTORDR1", "**", 0, True, "", 1)
        Fill_Records("SOTORDR1")


        'Cant use EDTTRPM1 for JCP to get customer because VAN uses mutliple customer codes for JCP but the same trading partner id
        Create_TDA(dst.Tables.Add, "ARTCUST2", "*", 0, False, "")

        Dim CUST_CODEs As String = ""
        For Each row As DataRow In dst.Tables("SOTORDR1").Select("")
            If Not CUST_CODEs.Contains(row.Item("CUST_CODE")) Then
                CUST_CODEs = CUST_CODEs & ",'" & row.Item("CUST_CODE") & "'"
            End If
        Next
        If CUST_CODEs <> "" Then
            CUST_CODEs = CUST_CODEs.Substring(1)
            Fill_Records("ARTCUST2", , , "SELECT * FROM ARTCUST2 where CUST_CODE in (" & CUST_CODEs & ") and CUST_ADDR_TYPE = 'DC'")
        End If

        sql = "Update EDT754I1 set EDI_PROCESS_IND = '1' Where EDI_PROCESS_IND is Null"
        ASCDATA1.ExecuteSQL(sql)


    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        
        Generate_Report(RPT, , SUBT)
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        'If eItemKey = "Proceed" Then
        '    If Not Absx1.chkFor("CHKPO").Checked And Not Absx1.chkFor("CHKPP").Checked Then
        '        EMsg &= vbCr & "You must select at least 1: Purchase Orders and/or Purchase Shipments"
        '    End If
        'End If
    End Sub

    Public Overrides Sub Update_Record()
        MyBase.Update_Record()

        Dim Invalid_Stores As Boolean

        BeginTrans()
        sql = "Drop Table SOTORDR1_754"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Drop Table SOTORDR5_754"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Drop Table SOTORDR0_754"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Drop Table SOTSHIP1_754"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Create table SOTORDR1_754 as "
        sql = sql & " select * from sotordr1 where cust_code = 'JCPL' and"
        sql = sql & " ORDR_CUST_PO in ("
        sql = sql & " select edi_po_no from edt754i3 Where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO FROM EDT754I1 Where EDI_PROCESS_IND is null))"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Create table SOTORDR5_754 as "
        sql = sql & " select * from sotordr5 where ordr_no in ("
        sql = sql & " select ordr_no from sotordr1 where cust_code = 'JCPL' and"
        sql = sql & " ORDR_CUST_PO in ("
        sql = sql & " select edi_po_no from edt754i3 Where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO FROM EDT754I1 Where EDI_PROCESS_IND is null))) and cust_addr_type = 'ST'"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Create table SOTORDR0_754 as "
        sql = sql & " Select * from sotordr0 where ordr_group_no in ("
        sql = sql & " select ordr_group_no from sotordr1 where cust_code = 'JCPL' and"
        sql = sql & " ORDR_CUST_PO in ("
        sql = sql & " select edi_po_no from edt754i3 Where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO FROM EDT754I1 Where EDI_PROCESS_IND is null)))"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Create table SOTSHIP1_754 as "
        sql = sql & " select * from sotship1 where ordr_group_no in ("
        sql = sql & " select ordr_group_no from sotordr1 where cust_code = 'JCPL' and"
        sql = sql & " ORDR_CUST_PO in ("
        sql = sql & " select edi_po_no from edt754i3 Where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO FROM EDT754I1 Where EDI_PROCESS_IND is null)))"
        ASCDATA1.ExecuteSQL(sql)
        'need to add check address codes that dont exist
        'sql check for invalid Stores
        sql = " Select Distinct I1.EDI_DOC_SEQ_NO, EDI_DOC_NO, '0'||I2.EDI_SHIP_ADDR_CODE ADDR_CODE, "
        sql = sql & "                 I3.EDI_PO_NO, CUST_CODE "
        sql = sql & " From EDTTRPM1 M1, EDT754I1 I1, EDT754I2 I2, EDT754I3 I3 "
        sql = sql & " Where M1.EDI_TP_QUAL = I1.EDI_TP_QUAL "
        sql = sql & " and M1.EDI_TP_ID = Rtrim(I1.EDI_TP_ID) "
        sql = sql & " and M1.EDI_DOC_NO = '753' "
        sql = sql & " and I1.EDI_DOC_SEQ_NO = I2.EDI_DOC_SEQ_NO "
        sql = sql & " and I2.EDI_DOC_SEQ_NO = I3.EDI_DOC_SEQ_NO "
        sql = sql & " and I2.EDI_754_SEQ_NO = I3.EDI_754_SEQ_NO "
        sql = sql & " and I1.EDI_PROCESS_IND is Null"
        sql = sql & " and ('0'||I2.EDI_SHIP_ADDR_CODE, I3.EDI_PO_NO) not in "
        sql = sql & " ( Select Distinct  C2.CUST_ADDR_CODE, I3.EDI_PO_NO "
        sql = sql & "    From EDTTRPM1 M1, EDT754I1 I1, EDT754I2 I2, "
        sql = sql & "         EDT754I3 I3, ARTCUST2 C2, SOTORDR1 O1"
        sql = sql & "    Where M1.EDI_TP_QUAL = I1.EDI_TP_QUAL"
        sql = sql & "     and M1.EDI_TP_ID = Rtrim(I1.EDI_TP_ID)"
        sql = sql & "     and M1.EDI_DOC_NO = '753'"
        sql = sql & "     and M1.CUST_CODE = C2.CUST_CODE"
        sql = sql & "     and C2.CUST_ADDR_TYPE = 'DC'"
        sql = sql & "     and C2.CUST_ADDR_CODE = '0'||I2.EDI_SHIP_ADDR_CODE"
        sql = sql & "     and I1.EDI_DOC_SEQ_NO = I2.EDI_DOC_SEQ_NO"
        sql = sql & "     and I2.EDI_DOC_SEQ_NO = I3.EDI_DOC_SEQ_NO"
        sql = sql & "     and I2.EDI_754_SEQ_NO = I3.EDI_754_SEQ_NO"
        sql = sql & "     and I3.EDI_PO_NO = O1.ORDR_CUST_PO"
        sql = sql & "     and O1.CUST_CODE = M1.CUST_CODE"
        sql = sql & "     and I1.EDI_PROCESS_IND is Null And ORDR_STATUS = 'P')"

        Dim Rows As Integer = ASCDATA1.GetDataTable(sql).Rows.Count
        
        If Rows > 0 Then
            Invalid_Stores = True
        Else
            Invalid_Stores = False
            Stop
            sql = "  BEGIN"
            sql = sql & " DECLARE"
            sql = sql & "    CURSOR C2 IS"
            sql = sql & "      Select Distinct I1.EDI_DOC_SEQ_NO, O1.ORDR_NO, I3.EDI_PO_NO, C2.CUST_ADDR_CODE, "
            sql = sql & "        C2.CUST_NAME, C2.CUST_ADDR1, C2.CUST_ADDR2, C2.CUST_CITY,"
            sql = sql & "        C2.CUST_STATE, C2.CUST_ZIP_CODE, C2.CUST_COUNTRY, C2.CUST_CONTACT, "
            sql = sql & "        C2.CUST_PHONE, C2.CUST_EXT, C2.CUST_FAX , C2.CUST_EMAIL"
            sql = sql & "      From EDTTRPM1 M1, EDT754I1 I1, EDT754I2 I2, "
            sql = sql & "           EDT754I3 I3, ARTCUST2 C2, SOTORDR1 O1"
            sql = sql & "      Where M1.EDI_TP_QUAL = I1.EDI_TP_QUAL"
            sql = sql & "       and M1.EDI_TP_ID = Rtrim(I1.EDI_TP_ID)"
            sql = sql & "       and M1.EDI_DOC_NO = '753'" & vbCr
            sql = sql & "       and M1.CUST_CODE = C2.CUST_CODE"
            sql = sql & "       and C2.CUST_ADDR_TYPE = 'DC'"
            sql = sql & "       and C2.CUST_ADDR_CODE = '0'||I2.EDI_SHIP_ADDR_CODE"
            sql = sql & "       and I1.EDI_DOC_SEQ_NO = I2.EDI_DOC_SEQ_NO"
            sql = sql & "       and I2.EDI_DOC_SEQ_NO = I3.EDI_DOC_SEQ_NO"
            sql = sql & "       and I2.EDI_754_SEQ_NO = I3.EDI_754_SEQ_NO"
            sql = sql & "       and I3.EDI_PO_NO = O1.ORDR_CUST_PO"
            sql = sql & "       and O1.CUST_CODE = M1.CUST_CODE"
            sql = sql & "       and I1.EDI_PROCESS_IND is Null And ORDR_STATUS = 'P'"
            sql = sql & "     Order by I1.EDI_DOC_SEQ_NO, O1.ORDR_NO, C2.CUST_ADDR_CODE;"
            sql = sql & "   BEGIN"
            sql = sql & "    FOR R2 IN C2 LOOP"
            sql = sql & "     UPDATE SOTORDR5 set CUST_ADDR_CODE = R2.CUST_ADDR_CODE, CUST_NAME = R2.CUST_NAME,"
            sql = sql & "        CUST_ADDR1 = R2.CUST_ADDR1, CUST_ADDR2 = R2.CUST_ADDR2, "
            sql = sql & "        CUST_CITY = R2.CUST_CITY, CUST_STATE = R2.CUST_STATE, "
            sql = sql & "        CUST_ZIP_CODE = R2.CUST_ZIP_CODE, CUST_COUNTRY = R2.CUST_COUNTRY,"
            sql = sql & "        CUST_CONTACT = R2.CUST_CONTACT, CUST_PHONE = R2.CUST_PHONE, "
            sql = sql & "        CUST_EXT = R2.CUST_EXT, CUST_FAX = R2.CUST_FAX, CUST_EMAIL = R2.CUST_EMAIL"
            sql = sql & "     where ORDR_NO = R2.ORDR_NO AND CUST_ADDR_TYPE = 'ST' ;"
            sql = sql & "   END LOOP;"
            sql = sql & " END; END;"
            ASCDATA1.ExecuteSQL(sql)

            sql = "  BEGIN"
            sql = sql & " DECLARE"
            sql = sql & "    CURSOR C2 IS"
            sql = sql & "     Select Distinct I1.EDI_DOC_SEQ_NO, C2.CUST_ADDR_CODE, I3.EDI_PO_NO, C2.CUST_CODE"
            sql = sql & "      From EDTTRPM1 M1, EDT754I1 I1, EDT754I2 I2, "
            sql = sql & "           EDT754I3 I3, ARTCUST2 C2, SOTORDR1 O1"
            sql = sql & "      Where M1.EDI_TP_QUAL = I1.EDI_TP_QUAL"
            sql = sql & "        and M1.EDI_TP_ID = Rtrim(I1.EDI_TP_ID)"
            sql = sql & "        and M1.EDI_DOC_NO = '753'"
            sql = sql & "        and M1.CUST_CODE = C2.CUST_CODE"
            sql = sql & "        and C2.CUST_ADDR_TYPE = 'DC'"
            sql = sql & "        and C2.CUST_ADDR_CODE = '0'||I2.EDI_SHIP_ADDR_CODE"
            sql = sql & "        and I1.EDI_DOC_SEQ_NO = I2.EDI_DOC_SEQ_NO"
            sql = sql & "        and I2.EDI_DOC_SEQ_NO = I3.EDI_DOC_SEQ_NO"
            sql = sql & "        and I2.EDI_754_SEQ_NO = I3.EDI_754_SEQ_NO"
            sql = sql & "        and I3.EDI_PO_NO = O1.ORDR_CUST_PO"
            sql = sql & "        and O1.CUST_CODE = M1.CUST_CODE"
            sql = sql & "        and I1.EDI_PROCESS_IND is Null And ORDR_STATUS = 'P'"
            sql = sql & "      Order by I1.EDI_DOC_SEQ_NO, C2.CUST_ADDR_CODE, I3.EDI_PO_NO;"
            sql = sql & "   BEGIN"
            sql = sql & "    FOR R2 IN C2 LOOP"
            sql = sql & "    UPDATE SOTORDR1 set CUST_DC_NO = r2.CUST_ADDR_CODE"
            sql = sql & "                    where ORDR_CUST_PO = r2.EDI_PO_NO"
            sql = sql & "                      and CUST_CODE = r2.CUST_CODE;"
            sql = sql & "   END LOOP;"
            sql = sql & "  END; END;"
            ASCDATA1.ExecuteSQL(sql)

            sql = "  BEGIN"
            sql = sql & "   DECLARE"
            sql = sql & "    CURSOR C2 IS"
            sql = sql & "     Select Distinct I1.EDI_DOC_SEQ_NO, O1.ORDR_GROUP_NO, C2.CUST_ADDR_CODE"
            sql = sql & "      From EDTTRPM1 M1, EDT754I1 I1, EDT754I2 I2, "
            sql = sql & "           EDT754I3 I3, ARTCUST2 C2, SOTORDR1 O1"
            sql = sql & "      Where M1.EDI_TP_QUAL = I1.EDI_TP_QUAL"
            sql = sql & "        and M1.EDI_TP_ID = Rtrim(I1.EDI_TP_ID)"
            sql = sql & "        and M1.EDI_DOC_NO = '753'"
            sql = sql & "        and M1.CUST_CODE = C2.CUST_CODE"
            sql = sql & "        and C2.CUST_ADDR_TYPE = 'DC'"
            sql = sql & "        and C2.CUST_ADDR_CODE = '0'||I2.EDI_SHIP_ADDR_CODE"
            sql = sql & "        and I1.EDI_DOC_SEQ_NO = I2.EDI_DOC_SEQ_NO"
            sql = sql & "        and I2.EDI_DOC_SEQ_NO = I3.EDI_DOC_SEQ_NO"
            sql = sql & "        and I2.EDI_754_SEQ_NO = I3.EDI_754_SEQ_NO"
            sql = sql & "        and I3.EDI_PO_NO = O1.ORDR_CUST_PO"
            sql = sql & "        and O1.CUST_CODE = M1.CUST_CODE"
            sql = sql & "        and I1.EDI_PROCESS_IND is Null And ORDR_STATUS = 'P'"
            sql = sql & "      Order by I1.EDI_DOC_SEQ_NO, O1.ORDR_GROUP_NO, C2.CUST_ADDR_CODE;"
            sql = sql & "   BEGIN"
            sql = sql & "    FOR R2 IN C2 LOOP"
            sql = sql & "    UPDATE SOTORDR0 set CUST_DC_NO = r2.CUST_ADDR_CODE"
            sql = sql & "                    where ORDR_GROUP_NO = r2.ORDR_GROUP_NO;"
            sql = sql & "   END LOOP;"
            sql = sql & "  END; END;"
            ASCDATA1.ExecuteSQL(sql)

            sql = "  BEGIN"
            sql = sql & "   DECLARE"
            sql = sql & "    CURSOR C2 IS"
            sql = sql & "     Select Distinct I1.EDI_DOC_SEQ_NO, O1.ORDR_GROUP_NO, C2.CUST_ADDR_CODE"
            sql = sql & "      From EDTTRPM1 M1, EDT754I1 I1, EDT754I2 I2, "
            sql = sql & "           EDT754I3 I3, ARTCUST2 C2, SOTORDR1 O1"
            sql = sql & "      Where M1.EDI_TP_QUAL = I1.EDI_TP_QUAL"
            sql = sql & "        and M1.EDI_TP_ID = Rtrim(I1.EDI_TP_ID)"
            sql = sql & "        and M1.EDI_DOC_NO = '753'"
            sql = sql & "        and M1.CUST_CODE = C2.CUST_CODE"
            sql = sql & "        and C2.CUST_ADDR_TYPE = 'DC'"
            sql = sql & "        and C2.CUST_ADDR_CODE = '0'||I2.EDI_SHIP_ADDR_CODE"
            sql = sql & "        and I1.EDI_DOC_SEQ_NO = I2.EDI_DOC_SEQ_NO"
            sql = sql & "        and I2.EDI_DOC_SEQ_NO = I3.EDI_DOC_SEQ_NO"
            sql = sql & "        and I2.EDI_754_SEQ_NO = I3.EDI_754_SEQ_NO"
            sql = sql & "        and I3.EDI_PO_NO = O1.ORDR_CUST_PO"
            sql = sql & "        and O1.CUST_CODE = M1.CUST_CODE"
            sql = sql & "        and I1.EDI_PROCESS_IND is Null And ORDR_STATUS = 'P'"
            sql = sql & "      Order by I1.EDI_DOC_SEQ_NO, O1.ORDR_GROUP_NO, C2.CUST_ADDR_CODE;"
            sql = sql & "   BEGIN"
            sql = sql & "    FOR R2 IN C2 LOOP"
            sql = sql & "    UPDATE SOTSHIP1 set SHIP_ADDR_CODE = r2.CUST_ADDR_CODE"
            sql = sql & "                    where ORDR_GROUP_NO = r2.ORDR_GROUP_NO;"
            sql = sql & "   END LOOP;"
            sql = sql & "  END; END;"
            ASCDATA1.ExecuteSQL(sql)




        End If
        CommitTrans()
    End Sub

End Class
