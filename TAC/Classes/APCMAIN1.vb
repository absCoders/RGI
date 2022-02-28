Imports System.Math

Public Class APCMAIN1

    Public Shared Function Prepare_Check_Register( _
    ByRef f As ASFSRPTM, _
    ByRef dst As DataSet, _
    ByVal check_register As Boolean, _
    Optional ByVal RYP1 As String = "", Optional ByVal RYP2 As String = "", _
    Optional ByVal DATE1 As Date = Nothing, Optional ByVal DATE2 As Date = Nothing) As String

        Dim sql As String
        Dim APTCHKR1 As String

        With dst
            sql = "Select 'I' RECORD_TYPE, APTCHCK1.* from APTCHCK1 "
            If check_register Then
                sql = sql & " where APTCHCK1.REGISTER_IND = '0'"
            Else
                If RYP1 = "" Then
                    sql = sql & " where APTCHCK1.CHECK_DATE >= '" & Format$(DATE1, "dd-MMM-yyyy") & "'"
                    sql = sql & "   and APTCHCK1.CHECK_DATE <= '" & Format$(DATE2, "dd-MMM-yyyy") & "'"
                Else
                    If RYP1 = RYP2 Or RYP2 = "" Then
                        sql = sql & " where APTCHCK1.OPS_YYYYPP = '" & RYP1 & "'"
                    Else
                        sql = sql & " where APTCHCK1.OPS_YYYYPP >= '" & RYP1 & "'"
                        sql = sql & "   and APTCHCK1.OPS_YYYYPP <= '" & RYP2 & "'"
                    End If
                End If
            End If
            sql = sql & f.SQL_in("BANK_CODE", "APTCHCK1.BANK_CODE")
            sql = sql & f.SQL_in("VEND_CODE_AP", "APTCHCK1.VEND_CODE_AP")
            sql = sql & f.SQL_in("PYMT_METHOD", "APTCHCK1.PYMT_METHOD")

            APTCHKR1 = ASCMAIN1.Temp_Table(sql)

            sql = "Select 'V' RECORD_TYPE, APTCHCK1.* from APTCHCK1 "
            If check_register Then
                sql = sql & "where APTCHCK1.REGISTER_IND_F = '0'"
            Else

                If RYP1 = "" Then
                    sql = sql & " where APTCHCK1.CHECK_DATE >= '" & Format$(DATE1, "dd-MMM-yyyy") & "'"
                    sql = sql & "   and APTCHCK1.CHECK_DATE <= '" & Format$(DATE2, "dd-MMM-yyyy") & "'"
                Else
                    If RYP1 = RYP2 Or RYP2 = "" Then
                        sql = sql & " where APTCHCK1.OPS_YYYYPP_F = '" & RYP1 & "'"
                    Else
                        sql = sql & " where APTCHCK1.OPS_YYYYPP_F >= '" & RYP1 & "'"
                        sql = sql & "   and APTCHCK1.OPS_YYYYPP_F <= '" & RYP2 & "'"
                    End If
                End If
                sql = sql & " and (APTCHCK1.CHECK_STATUS = 'V' OR APTCHCK1.CHECK_STATUS = 'R')"
            End If
            sql = sql & f.SQL_in("BANK_CODE", "APTCHCK1.BANK_CODE")
            sql = sql & f.SQL_in("VEND_CODE_AP", "APTCHCK1.VEND_CODE_AP")
            sql = sql & f.SQL_in("PYMT_METHOD", "APTCHCK1.PYMT_METHOD")

            ASCDATA1.ExecuteSQL("Insert into " & APTCHKR1 & " " & sql)
            ASCDATA1.ExecuteSQL("Alter table " & APTCHKR1 & " Add Primary Key (RECORD_TYPE, BANK_CODE, CHECK_NUM)")

            sql = "Select * from " & APTCHKR1
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTCHKR1", 3))

            sql = "Select APTCHKR1.RECORD_TYPE, APTCHCK2.*, APTINVH1.POST_CODE " _
            & ", APTCHKR1.OPS_YYYYPP, APTCHKR1.OPS_YYYYPP_F, APTINVH1.LC_FEE " _
            & " from APTCHCK2,APTINVH1," & APTCHKR1 & " APTCHKR1" _
            & " where APTCHKR1.BANK_CODE = APTCHCK2.BANK_CODE " _
            & "   and APTCHKR1.CHECK_NUM = APTCHCK2.CHECK_NUM " _
            & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTCHKR2", 4))


            sql = "Select APTINVH1.* from APTINVH1 " _
            & " where VOUCHER_NO in (Select APTINVH1.VOUCHER_NO" _
            & " from APTCHCK2,APTINVH1," & APTCHKR1 & " APTCHKR1" _
            & " where APTCHKR1.BANK_CODE = APTCHCK2.BANK_CODE " _
            & "   and APTCHKR1.CHECK_NUM = APTCHCK2.CHECK_NUM " _
            & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO)"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTINVH1", 1))

            .Tables.Add(ASCDATA1.GetDataTable("*", "GLTBANK1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "APTPOST1"))

            f.Create_TDA(.Tables.Add, "GLTINTF1", "*")


            If ASCMAIN1.CLIENT = "VAN" Then
                sql = "Select X.BANK_CODE, X.CHECK_NUM, POTSHIP2.*,POTSHIP1.PO_SHIP_VESSEL,POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                    & " from POTSHIP2,POTSHIP1,(" & vbCrLf _
                    & "Select DISTINCT APTCHKR1.BANK_CODE, APTCHKR1.CHECK_NUM, APTINVH5.PO_SHIPMENT_NO, APTINVH5.PO_SHIPMENT_LNO " & vbCrLf _
                    & " from APTINVH5,APTINVH1,APTCHCK2," & APTCHKR1 & " APTCHKR1" & vbCrLf _
                    & " where APTCHCK2.BANK_CODE = APTCHKR1.BANK_CODE" & vbCrLf _
                    & "   and APTCHCK2.CHECK_NUM = APTCHKR1.CHECK_NUM" & vbCrLf _
                    & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
                    & "   and APTINVH1.INV_PYMT_METHOD = 'LC'" & vbCrLf _
                    & "   and APTINVH5.VOUCHER_NO = APTCHCK2.VOUCHER_NO) X" & vbCrLf _
                    & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                    & "   and POTSHIP2.PO_SHIPMENT_NO = X.PO_SHIPMENT_NO" & vbCrLf _
                    & "   and POTSHIP2.PO_SHIPMENT_LNO = X.PO_SHIPMENT_LNO"

                dst.Tables.Add(ASCDATA1.GetDataTable(sql, "POTSHIPX", 4))
            End If

            If check_register Then
                sql = "Select APTINVH2.ACCT_CODE, APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE," & vbCrLf _
                & " APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT," & vbCrLf _
                & " SUM(APTINVH2.INV_LINE_AMT) GL_AMT" & vbCrLf _
                & " From APTCHCK1, APTCHCK2, APTINVH1, APTINVH2," & APTCHKR1 & " APTCHKR1" & vbCrLf _
                & " where APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
                & "   and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM" & vbCrLf _
                & "   and APTCHCK2.BANK_CODE = APTCHKR1.BANK_CODE" & vbCrLf _
                & "   and APTCHCK2.CHECK_NUM = APTCHKR1.CHECK_NUM" & vbCrLf _
                & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
                & "   and APTINVH2.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
                & "   and (APTCHCK1.CHECK_STATUS = 'I' OR NVL(APTCHCK1.OPS_YYYYPP,'??????') <> NVL(APTCHCK1.OPS_YYYYPP_F,'??????'))" & vbCrLf _
                & "   and APTCHKR1.RECORD_TYPE = APTCHCK1.CHECK_STATUS" & vbCrLf _
                & " group by APTINVH2.ACCT_CODE, APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE," & vbCrLf _
                & " APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT"

                sql &= vbCrLf & " union " & vbCrLf _
                & "Select '" & f.ROWs("APTPARM1").Item("AP_PARM_LC_FEE") & "' ACCT_CODE, APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE," & vbCrLf _
                & " APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT," & vbCrLf _
                & " SUM(APTINVH1.LC_FEE) GL_AMT" & vbCrLf _
                & " From APTCHCK1, APTCHCK2, APTINVH1," & APTCHKR1 & " APTCHKR1" & vbCrLf _
                & " where APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
                & "   and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM" & vbCrLf _
                & "   and APTCHCK2.BANK_CODE = APTCHKR1.BANK_CODE" & vbCrLf _
                & "   and APTCHCK2.CHECK_NUM = APTCHKR1.CHECK_NUM" & vbCrLf _
                & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
                & "   and (APTCHCK1.CHECK_STATUS = 'I' OR NVL(APTCHCK1.OPS_YYYYPP,'??????') <> NVL(APTCHCK1.OPS_YYYYPP_F,'??????'))" & vbCrLf _
                & "   and APTCHKR1.RECORD_TYPE = APTCHCK1.CHECK_STATUS" & vbCrLf _
                & "   and APTINVH1.LC_FEE <> 0" & vbCrLf _
                & " group by APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE, APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT"

                If RYP1 <> "" Then
                    sql &= vbCrLf & " union " & Replace(Replace(sql, "(APTCHCK1.CHECK_STATUS = 'I' OR NVL(APTCHCK1.OPS_YYYYPP,'??????') <> NVL(APTCHCK1.OPS_YYYYPP_F,'??????'))", "APTCHCK1.CHECK_STATUS = 'V' and APTCHCK1.OPS_YYYYPP <> APTCHCK1.OPS_YYYYPP_F"), " SUM(", " -1*SUM(")
                End If

            Else

                sql = "Select APTINVH2.ACCT_CODE, APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE," & vbCrLf _
                    & " APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT," & vbCrLf _
                    & " SUM(APTINVH2.INV_LINE_AMT) GL_AMT" & vbCrLf _
                    & " From APTCHCK1, APTCHCK2, APTINVH1, APTINVH2," & APTCHKR1 & " APTCHKR1" & vbCrLf _
                    & " where APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
                    & "   and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM" & vbCrLf _
                    & "   and APTCHCK2.BANK_CODE = APTCHKR1.BANK_CODE" & vbCrLf _
                    & "   and APTCHCK2.CHECK_NUM = APTCHKR1.CHECK_NUM" & vbCrLf _
                    & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
                    & "   and APTINVH2.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
                    & "   and (APTCHCK1.CHECK_STATUS = 'I' OR NVL(APTCHCK1.OPS_YYYYPP,'??????') <> NVL(APTCHCK1.OPS_YYYYPP_F,'??????')) and APTCHCK1.OPS_YYYYPP between '" & RYP1 & "' and '" & RYP2 & "'" & vbCrLf _
                    & " group by APTINVH2.ACCT_CODE, APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE," & vbCrLf _
                    & " APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT"

                sql &= vbCrLf & " union " & vbCrLf _
                    & "Select '" & f.ROWs("APTPARM1").Item("AP_PARM_LC_FEE") & "' ACCT_CODE, APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE," & vbCrLf _
                    & " APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT," & vbCrLf _
                    & " SUM(APTINVH1.LC_FEE) GL_AMT" & vbCrLf _
                    & " From APTCHCK1, APTCHCK2, APTINVH1," & APTCHKR1 & " APTCHKR1" & vbCrLf _
                    & " where APTCHCK2.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
                    & "   and APTCHCK2.CHECK_NUM = APTCHCK1.CHECK_NUM" & vbCrLf _
                    & "   and APTCHCK2.BANK_CODE = APTCHKR1.BANK_CODE" & vbCrLf _
                    & "   and APTCHCK2.CHECK_NUM = APTCHKR1.CHECK_NUM" & vbCrLf _
                    & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" & vbCrLf _
                    & "   and (APTCHCK1.CHECK_STATUS = 'I' OR NVL(APTCHCK1.OPS_YYYYPP,'??????') <> NVL(APTCHCK1.OPS_YYYYPP_F,'??????')) and APTCHCK1.OPS_YYYYPP between '" & RYP1 & "' and '" & RYP2 & "'" & vbCrLf _
                    & "   and APTINVH1.LC_FEE <> 0" & vbCrLf _
                    & " group by APTCHCK1.CHECK_NUM, APTCHCK1.VEND_CODE, APTCHCK1.VEND_NAME, APTCHCK1.CHECK_DATE, APTCHCK1.CHECK_AMT"


                If RYP1 <> "" Then
                    sql &= vbCrLf & " union " & Replace(Replace(sql, "(APTCHCK1.CHECK_STATUS = 'I' OR NVL(APTCHCK1.OPS_YYYYPP,'??????') <> NVL(APTCHCK1.OPS_YYYYPP_F,'??????')) and APTCHCK1.OPS_YYYYPP between", "APTCHCK1.CHECK_STATUS = 'V' and APTCHCK1.OPS_YYYYPP <> APTCHCK1.OPS_YYYYPP_F and APTCHCK1.OPS_YYYYPP_F between"), " SUM(", " -1*SUM(")
                End If
            End If




            dst.Tables.Add(ASCDATA1.GetDataTable(sql, "APTDISB1", 0))


            sql = "Select * from GLTACCT1"
            dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1", 1))
        End With

        For Each rowAPTCHKR1 As DataRow In dst.Tables("APTCHKR1").Select("RECORD_TYPE = 'V'", "")
            rowAPTCHKR1.Item("CHECK_AMT") = -1 * Val(rowAPTCHKR1.Item("CHECK_AMT") & "")
        Next
        For Each rowAPTCHKR2 As DataRow In dst.Tables("APTCHKR2").Select("RECORD_TYPE = 'V'", "")
            rowAPTCHKR2.Item("INV_AMT_APPLIED") = -1 * Val(rowAPTCHKR2.Item("INV_AMT_APPLIED") & "")
            rowAPTCHKR2.Item("INV_DISC_TAKEN") = -1 * Val(rowAPTCHKR2.Item("INV_DISC_TAKEN") & "")
            rowAPTCHKR2.Item("INV_WHSE_ALLOW_AMT") = -1 * Val(rowAPTCHKR2.Item("INV_WHSE_ALLOW_AMT") & "")
            rowAPTCHKR2.Item("LC_FEE") = -1 * Val(rowAPTCHKR2.Item("LC_FEE") & "")
        Next

        If check_register Then

            ' Prepare GL Interface File

            Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
            Dim JOURNAL_TYPE As String = "APCD"
            Dim JOURNAL_LNO As Integer = 0

            Dim DETL_POSTING_AMT As Double
            Dim BANK_CODE As String
            Dim POST_CODE As String
            Dim OPS_YYYYPP As String

            Call f.Summary_Table("APTCHKRB", "APTCHKR1", _
            "BANK_CODE,RECORD_TYPE,OPS_YYYYPP,OPS_YYYYPP_F", _
            "CHECK_AMT")

            For Each rowAPTCHKRB As DataRow In dst.Tables("APTCHKRB").Rows
                BANK_CODE = rowAPTCHKRB("BANK_CODE")
                Dim rowGLTBANK1 As DataRow = dst.Tables("GLTBANK1").Rows.Find(BANK_CODE)
                If rowAPTCHKRB("RECORD_TYPE") = "I" Then
                    OPS_YYYYPP = rowAPTCHKRB("OPS_YYYYPP")
                Else
                    OPS_YYYYPP = rowAPTCHKRB("OPS_YYYYPP_F")
                End If
                DETL_POSTING_AMT = -1 * Val(rowAPTCHKRB("CHECK_AMT") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = rowGLTBANK1("ACCT_CODE")
                    rowGLTINTF1("SEG2_CODE") = IIf(rowGLTBANK1("SEG2_CODE") & "" = "", f.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2"), rowGLTBANK1("SEG2_CODE"))
                    rowGLTINTF1("SEG3_CODE") = IIf(rowGLTBANK1("SEG3_CODE") & "" = "", f.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3"), rowGLTBANK1("SEG3_CODE"))
                    rowGLTINTF1("SEG4_CODE") = IIf(rowGLTBANK1("SEG4_CODE") & "" = "", f.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4"), rowGLTBANK1("SEG4_CODE"))
                    rowGLTINTF1("DETL_CTL_DATE") = Format(f.DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = f.XNO
                    rowGLTINTF1("DETL_CVX_NO") = BANK_CODE
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If
            Next

            Call f.Summary_Table("APTCHKRP", "APTCHKR2",
            "POST_CODE,RECORD_TYPE,OPS_YYYYPP,OPS_YYYYPP_F",
            "INV_AMT_APPLIED,INV_DISC_TAKEN,LC_FEE")

            For Each rowAPTCHKRP As DataRow In dst.Tables("APTCHKRP").Rows
                POST_CODE = rowAPTCHKRP("POST_CODE")
                Dim rowAPTPOST1 As DataRow = dst.Tables("APTPOST1").Rows.Find(POST_CODE)
                If rowAPTCHKRP("RECORD_TYPE") = "I" Then
                    OPS_YYYYPP = rowAPTCHKRP("OPS_YYYYPP")
                Else
                    OPS_YYYYPP = rowAPTCHKRP("OPS_YYYYPP_F")
                End If
                DETL_POSTING_AMT = Val(rowAPTCHKRP("INV_AMT_APPLIED") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = rowAPTPOST1("ACCT_CODE")
                    rowGLTINTF1("SEG2_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG2")
                    rowGLTINTF1("SEG3_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG3")
                    rowGLTINTF1("SEG4_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG4")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(f.DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = f.XNO
                    rowGLTINTF1("DETL_CVX_NO") = POST_CODE
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If

                DETL_POSTING_AMT = -1 * Val(rowAPTCHKRP("INV_DISC_TAKEN") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = f.ROWs("APTPARM1")("AP_PARM_ACCT_CODE_DISC")
                    rowGLTINTF1("SEG2_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG2")
                    rowGLTINTF1("SEG3_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG3")
                    rowGLTINTF1("SEG4_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG4")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(f.DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = f.XNO
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If

                DETL_POSTING_AMT = Val(rowAPTCHKRP("LC_FEE") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = f.ROWs("APTPARM1")("AP_PARM_LC_FEE")
                    rowGLTINTF1("SEG2_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG2")
                    rowGLTINTF1("SEG3_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG3")
                    rowGLTINTF1("SEG4_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG4")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(f.DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = f.XNO
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If

            Next
        End If

        Return APTCHKR1
    End Function
End Class
