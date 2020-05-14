Public Class ARCMAIN2

    Public Shared AGE_DAYS(4) As String
    Public Shared DUE_DAYS(4) As String
    Public Shared AGE_DATE(4) As String
    Public Shared DUE_DATE(4) As String
    Public Shared AGE_DATE_ORA(4) As String
    Public Shared DUE_DATE_ORA(4) As String
    Public Shared DAYS_AND_BUCKETS As String
    Public Shared AGED_TOTALS As String

    Public Shared Sub Record_Customer_Event(ByVal CUST_CODE As String, ByVal EVENT_DESC As String, ByVal EVENT_TYPE As String)
        TAC.TACMAIN1.Record_Event("ARTCUST1", CUST_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, EVENT_TYPE, EVENT_DESC)
    End Sub

    Public Shared Sub Pay_Open_AR_Item( _
    ByVal rowARTOPEN1 As DataRow, _
    ByVal rowARTPYMT2 As DataRow, _
    ByVal CURR_CODE As String, _
    ByVal PYMT_BATCH_DATE As Date, _
    ByVal INV_PMT As Double, _
    ByVal INV_DISC_TAKEN As Double, _
    ByVal INV_WRITE_OFF As Double, _
    ByRef PYMT_BATCH_ILNO As Integer, _
    ByVal F As ASFBASE1)

        If Not F.ROWs.ContainsKey("GLTPARM1") Then
            F.Get_PARM("GLTPARM1")
        End If

        With rowARTOPEN1
            Dim INV_BALANCE As Double = Val(.Item("INV_BALANCE") & "")

            .Item("INV_LAST_PMT") = PYMT_BATCH_DATE
            .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_PMT
            .Item("INV_DISC_TAKEN") = Val(.Item("INV_DISC_TAKEN") & "") + INV_DISC_TAKEN
            .Item("INV_WRITE_OFF") = Val(.Item("INV_WRITE_OFF") & "") + INV_WRITE_OFF
            .Item("INV_BALANCE") = Val(.Item("INV_BALANCE") & "") - (INV_PMT + INV_DISC_TAKEN + INV_WRITE_OFF)
            .Item("INV_LAST_PMT_REF") = rowARTPYMT2.Item("CUST_PYMT_REF_NO")
            .Item("INV_LAST_PMT_REF_DT") = rowARTPYMT2.Item("CUST_PYMT_REF_DATE")
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = F.DATETIME_STAMP
            If CURR_CODE = F.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                .Item("INV_PMT_CURR") = .Item("INV_PMT")
                .Item("INV_DISC_TAKEN_CURR") = .Item("INV_DISC_TAKEN")
                .Item("INV_WRITE_OFF_CURR") = .Item("INV_WRITE_OFF")
                .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
            Else
                MsgBox("ABS Stop Statement - Currency not set up") ' Stop
            End If

            Dim rowARTPYMT3 As DataRow = F.dst.Tables("ARTPYMT3").NewRow

            rowARTPYMT3.Item("PYMT_BATCH_NO") = rowARTPYMT2.Item("PYMT_BATCH_NO")
            rowARTPYMT3.Item("PYMT_BATCH_LNO") = rowARTPYMT2.Item("PYMT_BATCH_LNO")
            PYMT_BATCH_ILNO += 1
            rowARTPYMT3.Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
            rowARTPYMT3.Item("INV_TYPE") = .Item("INV_TYPE")
            rowARTPYMT3.Item("INV_NUM") = .Item("INV_NUM")
            rowARTPYMT3.Item("REASON_CODE") = .Item("REASON_CODE")
            rowARTPYMT3.Item("INV_DATE") = .Item("INV_DATE")
            rowARTPYMT3.Item("INV_DUE_DATE") = .Item("INV_DUE_DATE")
            rowARTPYMT3.Item("CUST_CODE_SO") = .Item("CUST_CODE_SO")
            rowARTPYMT3.Item("CUST_STORE_NO") = .Item("CUST_STORE_NO")
            rowARTPYMT3.Item("INV_CUST_PO") = .Item("INV_CUST_PO")
            rowARTPYMT3.Item("INV_BALANCE") = INV_BALANCE
            rowARTPYMT3.Item("INV_PMT") = INV_PMT
            rowARTPYMT3.Item("INV_DISC_TAKEN") = INV_DISC_TAKEN
            rowARTPYMT3.Item("INV_WRITE_OFF") = INV_WRITE_OFF
            rowARTPYMT3.Item("INV_BALANCE_NEW") = .Item("INV_BALANCE")
            rowARTPYMT3.Item("POST_CODE") = .Item("POST_CODE")
            rowARTPYMT3.Item("SEG2_CODE") = .Item("SEG2_CODE")
            rowARTPYMT3.Item("SEG3_CODE") = .Item("SEG3_CODE")
            rowARTPYMT3.Item("SEG4_CODE") = .Item("SEG4_CODE")
            If CURR_CODE = F.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                rowARTPYMT3.Item("INV_BALANCE_CURR") = rowARTPYMT3.Item("INV_BALANCE")
                rowARTPYMT3.Item("INV_PMT_CURR") = rowARTPYMT3.Item("INV_PMT")
                rowARTPYMT3.Item("INV_DISC_TAKEN_CURR") = rowARTPYMT3.Item("INV_DISC_TAKEN")
                rowARTPYMT3.Item("INV_WRITE_OFF_CURR") = rowARTPYMT3.Item("INV_WRITE_OFF")
                rowARTPYMT3.Item("INV_BALANCE_NEW_CURR") = rowARTPYMT3.Item("INV_BALANCE_NEW")
            Else
                MsgBox("ABS Stop Statement - Currency not set up") ' Stop
            End If
            F.dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
        End With
    End Sub

    Public Shared Function Last_3612_SAvg(ByVal CUST_CREDIT_GROUP_CUST As String) As String

        Dim LAST_3612 As String = ""
        Dim YP3 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)
        Dim YP6 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6)
        Dim YP12 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)

        ASCMAIN1.sql = "SELECT MIN (OPS_YYYYPP) YP" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP3 & "' THEN INV_CLSD_DYS ELSE 0 END) D03" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP3 & "' THEN INV_CLSD_CNT ELSE 0 END) T03" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP6 & "' THEN INV_CLSD_DYS ELSE 0 END) D06" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP6 & "' THEN INV_CLSD_CNT ELSE 0 END) T06" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP12 & "' THEN INV_CLSD_DYS ELSE 0 END) D12" _
        & ", SUM (CASE WHEN OPS_YYYYPP >='" & YP12 & "' THEN INV_CLSD_CNT ELSE 0 END) T12" _
        & " from ARTSTMT1 where OPS_YYYYPP >= '" & YP12 & "'" _
        & "  and ARTSTMT1.CUST_CODE = '" & CUST_CREDIT_GROUP_CUST & "'"
        '& "  and ARTSTMT1.CUST_CODE in (Select CUST_CODE from " & ARTSTMTX & ")"

        Dim row3612 As DataRow = ASCDATA1.GetDataRow
        Dim YP As String = row3612.Item("YP") & ""
        Dim V As Decimal = 0
        Dim DXX As Decimal = 0
        Dim TXX As Decimal = 0
        If YP <= YP3 Then
            DXX = Val(row3612.Item("D03") & "")
            TXX = Val(row3612.Item("T03") & "")
            If TXX = 0 Then
                V = 0
            Else
                V = DXX / TXX
            End If
            LAST_3612 &= "/" & Format(V, "##0")
            If YP <= YP6 Then
                DXX = Val(row3612.Item("D06") & "")
                TXX = Val(row3612.Item("T06") & "")
                If TXX = 0 Then
                    V = 0
                Else
                    V = DXX / TXX
                End If
                LAST_3612 &= "/" & Format(V, "##0")
                If YP <= YP12 Then
                    DXX = Val(row3612.Item("D12") & "")
                    TXX = Val(row3612.Item("T12") & "")
                    If TXX = 0 Then
                        V = 0
                    Else
                        V = DXX / TXX
                    End If
                    LAST_3612 &= "/" & Format(V, "##0")
                End If
            End If
        End If
        LAST_3612 = Mid(LAST_3612, 2)

        Return LAST_3612
    End Function
    ' NEED TO ADD THEBELOW ROUTINE TO ARCMAIN1 - AND MAKE IT REGENCY SPECIFIC

    Public Shared Sub Get_Aging_Data_RGI( _
    ByVal rowARTPARM1 As DataRow, _
    ByVal BASE_DATE As Date, _
    Optional ByVal use_parms_for_days As Boolean = True)

        If use_parms_for_days Then
            For i As Integer = 2 To 4
                AGE_DAYS(i) = Val(rowARTPARM1.Item("AR_PARM_AGE_CATG_" & CStr(i)) & "")
                DUE_DAYS(i) = Val(rowARTPARM1.Item("AR_PARM_DUE_CATG_" & CStr(i)) & "")
            Next
        End If
        For i As Integer = 1 To 4
            Dim PRD_END_DATE As Date = BASE_DATE.AddDays(-1 * AGE_DAYS(i))
            AGE_DATE(i) = "'" & Format(PRD_END_DATE, "MM/dd/yyyy") & "'"
            AGE_DATE_ORA(i) = "'" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
        Next
        For i As Integer = 1 To 4
            Dim PRD_END_DATE As Date = BASE_DATE.AddDays(-1 * DUE_DAYS(i))
            DUE_DATE(i) = "'" & Format(PRD_END_DATE, "MM/dd/yyyy") & "'"
            DUE_DATE_ORA(i) = "'" & Format(PRD_END_DATE, "dd-MMM-yyyy") & "'"
        Next

        'Dim sqlINV As String = "ARTOPEN1.INV_DATE"
        'Dim sqlDUE As String = "ARTOPEN1.INV_DUE_DATE"

        '''DAYS_AND_BUCKETS = "" _
        '''& ", DECODE(ARTOPEN1.OPS_YYYYPP_PAID,NULL,ARTOPEN1.DATE_PAID,TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "')) - ARTOPEN1.INV_DATE DAYS" & vbCrLf _
        '''& ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 or (TATTERM1.TERM_DUE_TYPE = 'S' AND  ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(1) & ") THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(2) & "                                                  THEN '1' ELSE" & vbCrLf _
        '''& "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        '''& "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        '''& "  '4' END END END END) AGE_BUCKET" & vbCrLf _
        '''& ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DATE) AGE" & vbCrLf _
        '''& ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 or (TATTERM1.TERM_DUE_TYPE = 'S'AND  ARTOPEN1.INV_DUE_DATE > " & DUE_DATE_ORA(1) & ") THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(2) & "                                          THEN '1' ELSE" & vbCrLf _
        '''& "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        '''& "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        '''& "  '4' END END END END) DUE_BUCKET" & vbCrLf _
        '''& ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DUE_DATE) DUE"




        DAYS_AND_BUCKETS = "" _
        & ", DECODE(ARTOPEN1.OPS_YYYYPP_PAID,NULL,ARTOPEN1.DATE_PAID,TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "')) - ARTOPEN1.INV_DATE DAYS" & vbCrLf _
        & ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 or (TATTERM1.TERM_DUE_TYPE = 'S' AND  ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(1) & ") THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(2) & "                                                  THEN '1' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        & "  '4' END END END END) AGE_BUCKET" & vbCrLf _
        & ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DATE) AGE" & vbCrLf _
        & ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 or ARTOPEN1.INV_DUE_DATE > " & DUE_DATE_ORA(1) & "  THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(2) & "                                          THEN '1' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        & "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        & "  '4' END END END END) DUE_BUCKET" & vbCrLf _
        & ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DUE_DATE) DUE"












        'DAYS_AND_BUCKETS = "" _
        '& ", DECODE(ARTOPEN1.OPS_YYYYPP_PAID,NULL,ARTOPEN1.DATE_PAID,TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "')) - ARTOPEN1.INV_DATE DAYS" & vbCrLf _
        '& ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(2) & "                                                  THEN '1' ELSE" & vbCrLf _
        '& "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        '& "   CASE WHEN ARTOPEN1.INV_DATE > " & AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE <= " & AGE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        '& "  '4' END END END END) AGE_BUCKET" & vbCrLf _
        '& ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DATE) AGE" & vbCrLf _
        '& ", (CASE WHEN ARTOPEN1.INV_BALANCE = 0 THEN '0' ELSE CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(2) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(1) & " THEN '1' ELSE" & vbCrLf _
        '& "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(2) & " THEN '2' ELSE" & vbCrLf _
        '& "   CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & DUE_DATE_ORA(3) & " THEN '3' ELSE" & vbCrLf _
        '& "  '4' END END END END) DUE_BUCKET" & vbCrLf _
        '& ", TO_NUMBER(TO_DATE('" & Format(BASE_DATE, "dd-MMM-yyyy") & "') - ARTOPEN1.INV_DUE_DATE) DUE"






        Dim sqle As String = " THEN ARTOPEN1.INV_BALANCE ELSE 0 END "

        AGED_TOTALS = "" _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN2.AGE_DATE_ORA(2) & " AND ARTOPEN1.INV_DATE <= " & TAC.ARCMAIN2.AGE_DATE_ORA(1) & sqle & ") AGE_1" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN2.AGE_DATE_ORA(3) & " AND ARTOPEN1.INV_DATE <= " & TAC.ARCMAIN2.AGE_DATE_ORA(2) & sqle & ") AGE_2" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE > " & TAC.ARCMAIN2.AGE_DATE_ORA(4) & " AND ARTOPEN1.INV_DATE <= " & TAC.ARCMAIN2.AGE_DATE_ORA(3) & sqle & ") AGE_3" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DATE                                                              <= " & TAC.ARCMAIN2.AGE_DATE_ORA(4) & sqle & ") AGE_4" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & TAC.ARCMAIN2.DUE_DATE_ORA(2) & " AND ARTOPEN1.INV_DUE_DATE < " & TAC.ARCMAIN2.AGE_DATE_ORA(1) & sqle & ") DUE_1" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & TAC.ARCMAIN2.DUE_DATE_ORA(3) & " AND ARTOPEN1.INV_DUE_DATE < " & TAC.ARCMAIN2.DUE_DATE_ORA(2) & sqle & ") DUE_2" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE >= " & TAC.ARCMAIN2.DUE_DATE_ORA(4) & " AND ARTOPEN1.INV_DUE_DATE < " & TAC.ARCMAIN2.DUE_DATE_ORA(3) & sqle & ") DUE_3" & vbCrLf _
        & ", SUM (CASE WHEN ARTOPEN1.INV_DUE_DATE                                                                   < " & TAC.ARCMAIN2.DUE_DATE_ORA(4) & sqle & ") DUE_4" & vbCrLf


    End Sub
End Class
