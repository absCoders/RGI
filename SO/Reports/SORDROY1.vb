Public Class SORDROY1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -60, 0, 0)

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "R"

        'If ASCMAIN1.EOM <> "1" Then
        RWU = "N"
        'End If

        Dim SQL As New System.Text.StringBuilder With {.Length = 0}

        SQL.AppendLine("SELECT")
        SQL.AppendLine("R1.VEND_CODE,")
        SQL.AppendLine("V1.VEND_NAME,")
        SQL.AppendLine("R1.ROYALTY_CODE,")
        SQL.AppendLine("R1.ROYALTY_DESC,")
        SQL.AppendLine("R2.ROYALTY_PCT,")
        SQL.AppendLine("I2.STYLE_CODE,")
        SQL.AppendLine("S1.STYLE_DESC,")
        SQL.AppendLine("I1.CUST_CODE,")
        SQL.AppendLine("C1.CUST_NAME,")
        SQL.AppendLine("I1.ORDR_NO,")
        SQL.AppendLine("I1.INV_NO,")
        SQL.AppendLine("I1.INV_DATE,")
        SQL.AppendLine("I1.INV_TYPE,")
        SQL.AppendLine("DECODE(I1.INV_TYPE,'I','Invoice','Credit') AS TYPE,")
        SQL.AppendLine("I2.ORDR_UNIT_PRICE,")
        SQL.AppendLine("SUM(I2.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP,")
        SQL.AppendLine("SUM((I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE)) AS SHIPPED_PRICE,")
        SQL.AppendLine("SUM(DECODE(I1.INV_TYPE,'I',((I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) * (R2.ROYALTY_PCT / 100)),(((I2.ORDR_QTY_SHIP * I2.ORDR_UNIT_PRICE) * -1) * (R2.ROYALTY_PCT / 100)))) AS PAY_AMT")
        SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, ARTCUST1 C1, ICTSTYL1 S1, ICTROYL1 R1, ICTROYL2 R2, APTVEND1 V1")
        SQL.AppendLine("WHERE I1.INV_NO = I2.INV_NO")
        SQL.AppendLine("AND I2.STYLE_CODE = S1.STYLE_CODE")
        SQL.AppendLine("AND I1.CUST_CODE = C1.CUST_CODE")
        SQL.AppendLine("AND S1.ROYALTY_CODE = R1.ROYALTY_CODE")
        SQL.AppendLine("AND R1.ROYALTY_CODE = R2.ROYALTY_CODE")
        SQL.AppendLine("AND R1.VEND_CODE = V1.VEND_CODE")
        SQL.AppendLine("AND R2.ROYALTY_BEGIN <= '01-JUL-2022'")
        SQL.AppendLine("AND R2.ROYALTY_END >= '31-JUL-2022'")
        'SQL.AppendLine("AND I1.ORDR_YYYYPP_UPDATED = '202207'")
        SQL.AppendLine("AND I1.INV_DATE >= '01-JUL-2022'")
        SQL.AppendLine("AND I1.INV_DATE <= '31-JUL-2022'")
        SQL.AppendLine("AND I2.ORDR_QTY_SHIP > 0")
        SQL.AppendLine("GROUP BY")
        SQL.AppendLine("R1.VEND_CODE,")
        SQL.AppendLine("V1.VEND_NAME,")
        SQL.AppendLine("R1.ROYALTY_CODE,")
        SQL.AppendLine("R1.ROYALTY_DESC,")
        SQL.AppendLine("R2.ROYALTY_PCT,")
        SQL.AppendLine("I2.STYLE_CODE,")
        SQL.AppendLine("S1.STYLE_DESC,")
        SQL.AppendLine("I1.CUST_CODE,")
        SQL.AppendLine("C1.CUST_NAME,")
        SQL.AppendLine("I1.ORDR_NO,")
        SQL.AppendLine("I1.INV_NO,")
        SQL.AppendLine("I1.INV_DATE,")
        SQL.AppendLine("I1.INV_TYPE,")
        SQL.AppendLine("DECODE(I1.INV_TYPE,'I','Invoice','Credit'),")
        SQL.AppendLine("I2.ORDR_UNIT_PRICE")
        'ASCMAIN1.sql = SQL.ToString
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SORDROY1"))
        'Fill_Records("SORDROY1")
        Create_TDA(dst.Tables.Add, "SORDROY1", SQL.ToString, 0, False)
        Fill_Records("SORDROY1")

        'ASCMAIN1.Progress("")
    End Sub

    Public Overrides Sub Print_Report()

        ' CR_params.Add("SORT_CUST", IIf(tblASTDSQLA.Rows.Find("CUST_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        ' CR_params.Add("SORT_SREP", IIf(tblASTDSQLA.Rows.Find("SREP_CODE").Item("SEQUENCE") & "" = "", "0", "1"))
        ' CR_params.Add("SORT_WHSE", IIf(tblASTDSQLA.Rows.Find("WHSE_CODE").Item("SEQUENCE") & "" = "", "0", "1"))

        CR_params.Add("SUBT", "")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class