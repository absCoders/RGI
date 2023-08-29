Public Class ICRROYL1
    Dim LEGEND As String = ""
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -60, 0, 0)

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        Dim SQL As New Text.StringBuilder With {.Length = 0}
        SQL.AppendLine("SELECT")
        SQL.AppendLine("I1.ORDR_YYYYPP_UPDATED,")
        SQL.AppendLine("I1.INV_DATE,")
        SQL.AppendLine("I1.INV_NO,")
        SQL.AppendLine("I1.CUST_CODE,")
        SQL.AppendLine("C1.CUST_NAME,")
        SQL.AppendLine("I1.SREP_CODE,")
        SQL.AppendLine("SR.SREP_NAME,")
        SQL.AppendLine("I2.STYLE_CODE,")
        SQL.AppendLine("I2.COLOR_CODE,")
        SQL.AppendLine("O2.STYLE_PRICE,")
        SQL.AppendLine("NVL(I2.ORDR_QTY_SHIP,0) AS ORDR_QTY_SHIP,")
        SQL.AppendLine("NVL(I2.ORDR_UNIT_PRICE,0) AS ORDR_UNIT_PRICE,")
        SQL.AppendLine("(NVL(I2.ORDR_QTY_SHIP,0) * NVL(I2.ORDR_UNIT_PRICE,0)) AS ORDR_TOTAL,")
        SQL.AppendLine("R1.ROYALTY_CODE,")
        SQL.AppendLine("R1.ROYALTY_DESC,")
        SQL.AppendLine("(R2.ROYALTY_PCT / 100) AS ROYALTY_PCT,")
        SQL.AppendLine("NVL(I2.ORDR_UNIT_PRICE,0) * (R2.ROYALTY_PCT / 100) AS ORDR_UNIT_ROYALTY,")
        SQL.AppendLine("(NVL(I2.ORDR_QTY_SHIP,0) * NVL(I2.ORDR_UNIT_PRICE,0)) * (R2.ROYALTY_PCT / 100) AS ORDR_TOTAL_ROYALTY")
        SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST1 C1, SOTORDR1 O1, SOTORDR2 O2, ICTSTYL1 S1, ICTROYL2 R2, ICTROYL1 R1, SOTSREP1 SR")
        SQL.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
        SQL.AppendLine("AND I1.CUST_CODE = C1.CUST_CODE")
        SQL.AppendLine("AND O1.ORDR_NO = O2.ORDR_NO")
        SQL.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO")
        SQL.AppendLine("AND I2.INV_LNO = O2.ORDR_LNO")
        SQL.AppendLine("AND I2.STYLE_CODE = S1.STYLE_CODE")
        SQL.AppendLine("AND R2.ROYALTY_CODE = R1.ROYALTY_CODE")
        SQL.AppendLine("AND I1.SREP_CODE = SR.SREP_CODE")
        SQL.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED = '{RYP}'")
        SQL.AppendLine("AND NVL(S1.ROYALTY_CODE,'NULL') <> 'NULL'")
        SQL.AppendLine("AND NVL(S1.ROYALTY_CODE,'NULL') = R2.ROYALTY_CODE")
        SQL.AppendLine("AND (R2.ROYALTY_BEGIN <= O1.ORDR_DATE AND R2.ROYALTY_END >= O1.ORDR_DATE)")
        SQL.AppendLine("AND I1.CUST_CODE NOT IN ('180000','301758')")
        SQL.AppendLine("AND NVL(I2.ORDR_QTY_SHIP,0) > 0")
        ASCMAIN1.sql = SQL.ToString
        '& " from SOTINVH2 where ORDR_YYYYPP_UPDATED >= '" & Mid(RYP, 1, 4) & "01' and ORDR_YYYYPP_UPDATED <= '" & RYP & "'" _

        Create_TDA(dst.Tables.Add, "ICTROYLX", "**", 0, False, "", 0)
        Fill_Records("ICTROYLX")

        'SQL.Length = 0
        'SQL.AppendLine("SELECT LEGEND")
        'SQL.AppendLine("FROM GLTPARM2")
        'SQL.AppendLine($"WHERE OPS_YYYYPP = '{RYP}'")
        'ASCMAIN1.sql = SQL.ToString()
        'LEGEND = ASCDATA1.GetDataValue

    End Sub

    Public Overrides Sub Print_Report()
        'SUBT = "{RYPLEGEND} - Designers"

        RPT = "ICRROYL1"
        RPT_TITLE = "Designer Royalty Commissions"
        CR_params.Add("SUBT", $"{RYPLEGEND} - Designers")
        CR_params.Add("RPT_FOR", "D")
        Generate_Report(RPT, , SUBT)

        RPT = "ICRROYL1"
        RPT_TITLE = "Designer Royalty Commissions"
        CR_params.Add("SUBT", $"{RYPLEGEND} - Sales Reps")
        CR_params.Add("RPT_FOR", "S")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class