Public Class ARRFINC1

    Dim ARTFINC1 As String
    Dim PRD_END_DATE As Date

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        If ASCMAIN1.EOM <> "1" Then
            RWU = "N"
        End If

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
        PRD_END_DATE = rowGLTPARM2.Item("PRD_END_DATE")
        Dim PD As String = Format(PRD_END_DATE.AddDays(0), "dd-MMM-yyyy")
        Dim PDFC As String = Format(PRD_END_DATE.AddDays(-90), "dd-MMM-yyyy")

        ASCMAIN1.sql = "Select CUST_CODE, SUM (INV_BALANCE) INV_BALANCE" & vbCrLf _
            & ", SUM (CASE WHEN INV_BALANCE < 0 THEN INV_BALANCE ELSE 0 END) INV_BALANCE_CR" & vbCrLf _
            & ", SUM (CASE WHEN INV_DUE_DATE < '" & PD & "' AND INV_BALANCE > 0 THEN INV_BALANCE ELSE 0 END) INV_BALANCE_PD" & vbCrLf _
            & ", SUM (CASE WHEN INV_DUE_DATE < '" & PDFC & "' AND INV_BALANCE > 0  THEN INV_BALANCE ELSE 0 END) INV_BALANCE_PDFC" & vbCrLf _
            & " from ARTOPEN1 where INV_BALANCE <> 0" & vbCrLf _
            & " group by CUST_CODE"
        ASCMAIN1.sql = "Select X.* from (" & ASCMAIN1.sql & ") X, ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_FIN_CHG_IND = '1'" & vbCrLf _
            & "   and X.INV_BALANCE_PDFC > 0"
        ARTFINC1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTFINC1 & " Add Primary Key (CUST_CODE)")

        ASCMAIN1.sql = "Select * from " & ARTFINC1
        Create_TDA(dst.Tables.Add, "ARTFINC1", "**", 0, False, "", 1)
        With dst.Tables("ARTFINC1").Columns
            Dim AR_PARM_FIN_CHG_RATE As Decimal = Val(ROWs("ARTPARM1").Item("AR_PARM_FIN_CHG_RATE") & "")
            .Add("FIN_CHG", GetType(System.Decimal), "INV_BALANCE_PDFC * " & CStr(AR_PARM_FIN_CHG_RATE) & " / 100")
            .Add("PD_PCT", GetType(System.Decimal), "IIF(INV_BALANCE = 0, 0, 100 * INV_BALANCE_PD / INV_BALANCE)")
            .Add("PDFC_PCT", GetType(System.Decimal), "IIF(INV_BALANCE = 0, 0, 100 * INV_BALANCE_PDFC / INV_BALANCE)")
        End With
        Fill_Records("ARTFINC1")

        ASCMAIN1.sql = "Select * from ARTCUST1 where CUST_CODE in (Select CUST_CODE from " & ARTFINC1 & ")"
        Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)
        Fill_Records("ARTCUST1")

        Create_TDA(dst.Tables.Add, "SOTINVH1", "*")
        Create_TDA(dst.Tables.Add, "SOTINVHM", "*")
        Create_TDA(dst.Tables.Add, "ARTOPEN1", "*")
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("SUBT", "")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        For Each rowARTFINC1 As DataRow In dst.Tables("ARTFINC1").Select("")
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
            Dim INV_NO As String = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            Dim CUST_CODE As String = rowARTFINC1.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)

            Dim FIN_CHG As Decimal = System.Math.Round(Val(rowARTFINC1.Item("FIN_CHG") & ""), 2)
            Dim INV_BALANCE_PDFC As Decimal = System.Math.Round(Val(rowARTFINC1.Item("INV_BALANCE_PDFC") & ""), 2)

            Dim rowSOTMISC1 As DataRow = LookUp("SOTMISC1", ROWs("ARTPARM1").Item("AR_PARM_FIN_CHG_CODE") & "")

            With rowSOTINVH1
                .Item("INV_TYPE") = "I"
                .Item("INV_NO") = INV_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("INV_MISC_CHG") = FIN_CHG
                .Item("INV_TOTAL_AMOUNT") = FIN_CHG
                .Item("REASON_CODE") = rowSOTMISC1.Item("REASON_CODE")
                .Item("INV_DATE") = PRD_END_DATE
                .Item("ORDR_DATE_UPDATED") = DATETIME_STAMP
                .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP

                .Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE")
                .Item("SALES_DIVISION_CODE") = "RIB"
                .Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0")
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
                .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE")
                .Item("INV_TOTAL_AMOUNT_CURR") = FIN_CHG
                .Item("CURR_CODE") = "USD"
                .Item("CURR_EXCH_RATE") = 1
                .Item("INV_MISC_CHG_CURR") = FIN_CHG
                .Item("INV_TOTAL_AMT_CURR") = FIN_CHG
                .Item("ORDR_TYPE_CODE") = "FIN"
                .Item("CUST_BILL_TO_CUST") = CUST_CODE
            End With
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

            Dim rowSOTINVHM As DataRow = dst.Tables("SOTINVHM").NewRow
            With rowSOTINVHM
                .Item("INV_TYPE") = "I"
                .Item("INV_NO") = INV_NO
                .Item("INV_MNO") = 1
                .Item("MISC_CHG_CODE") = ROWs("ARTPARM1").Item("AR_PARM_FIN_CHG_CODE")
                .Item("MISC_CHG_DESC") = "Finance Charge"
                ' .Item("MISC_CHG_NOTE") = "1% Past Due > 90 Days: " & Format(INV_BALANCE_PDFC, "$#.00")
                .Item("INV_MISC_CHG") = FIN_CHG
            End With
            dst.Tables("SOTINVHM").Rows.Add(rowSOTINVHM)

            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
            With rowARTOPEN1
                For Each COLUMN_NAME As String In New String() _
                    {"CUST_CODE", "INV_TOTAL_AMOUNT", "INV_TYPE", "INV_DATE", "POST_CODE", "TERM_CODE", "SREP_CODE", _
                     "INV_TOTAL_AMOUNT", "INIT_OPER", "INIT_DATE", "INV_MISC_CHG", "CURR_CODE", "CURR_EXCH_RATE", "INV_MISC_CHG_CURR"}
                    .Item(COLUMN_NAME) = rowSOTINVH1.Item(COLUMN_NAME)
                Next
                .Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")
                .Item("INV_DUE_DATE") = .Item("INV_DATE")
                .Item("INV_BALANCE") = .Item("INV_TOTAL_AMOUNT")
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
  
            End With
           
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            rowARTOPEN1.Item("INV_BALANCE_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
            rowARTOPEN1.Item("INV_NOTES") = rowSOTINVH1.Item("INV_COMMENT")
            rowARTOPEN1.Item("ORDR_TYPE_CODE") = rowSOTINVH1.Item("ORDR_TYPE_CODE")
            'rowARTOPEN1.Item("INV_REF") = rowSOTINVH1.Item("INV_REF")
            rowARTOPEN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowARTOPEN1.Item("SALES_DIVISION_CODE") = rowSOTINVH1.Item("SALES_DIVISION_CODE")
            dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Update ARTCUST6 Set CUST_LAST_FCHG_DATE = :PARM1" & vbCrLf _
                & ", CUST_FIN_CHG_MTD = NVL(CUST_FIN_CHG_MTD,0) + :PARM2" & vbCrLf _
                & ", CUST_FIN_CHG_YTD = NVL(CUST_FIN_CHG_YTD,0) + :PARM2" & vbCrLf _
                & ", CUST_NUM_FIN_MTD = NVL(CUST_NUM_FIN_MTD,0) + 1" & vbCrLf _
                & ", CUST_NUM_FIN_YTD = NVL(CUST_NUM_FIN_YTD,0) + 1" & vbCrLf _
                & " where CUST_CODE = :PARM3;" & vbCrLf _
                & " If SQL%NOTFOUND Then " & vbCrLf _
                & "  Insert into ARTCUST6 (CUST_CODE,CUST_FIN_CHG_MTD,CUST_FIN_CHG_YTD,CUST_NUM_FIN_MTD,CUST_NUM_FIN_YTD)" & vbCrLf _
                & "   Values" & vbCrLf _
                & "   (:PARM3,:PARM2,:PARM2,1,1);" & vbCrLf _
                & " End If;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DNV", New Object() {PRD_END_DATE, FIN_CHG, CUST_CODE})
        Next

        Update_Record_TDA("SOTINVH1")
        Update_Record_TDA("SOTINVHM")
        Update_Record_TDA("ARTOPEN1")
    End Sub
End Class