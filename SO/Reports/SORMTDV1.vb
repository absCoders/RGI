Public Class SORMTDV1

#Region "Declarations"
    Dim DTE0 As Date
    Dim DTE1 As Date
    Dim SOTINVH1 As String
    Dim SOTINVH2 As String

    Dim S As New System.Text.StringBuilder() With {.Length = 0}
#End Region

#Region "ABS Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("SOTPARM1")
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -84, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
    End Sub

    Overrides Sub Clear_Record()
    End Sub

    Protected Overrides Sub Build_Workfile()
        RYPLEGEND0 = Absx1.cmbFor("RYP0").Value
        RYP0 = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
        RYPLEGEND1 = Absx1.cmbFor("RYP1").Value
        RYP1 = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)
        If RYP0 = RYP1 Then
            SUBT = "Invoices Posted in " & RYPLEGEND0
        Else
            SUBT = String.Format("Invoices Posted between {0} and {1}", RYPLEGEND0, RYPLEGEND1)
        End If

        Dim sqlw As String = ""
        sqlw &= String.Format("   and SOTINVH1.ORDR_YYYYPP_UPDATED between '{0}' and '{1}'{2}", RYP0, RYP1, vbCrLf)
        sqlw &= SQL_in("SALES_DIVISION_CODE", "ICTSTYL1.SALES_DIVISION_CODE")
        sqlw &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE")
        sqlw &= "    and SOTINVH1.CUST_CODE <> 'TRANSFERS'"
        sqlw &= "    and SOTINVH1.CUST_CODE <> 'SAMPLES'"
        Prepare_Sales_Invoices(sqlw, SOTINVH1, SOTINVH2)

        Check_if_Empty("SOTINVHZ")
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = SUBT & " Excluding Samples & Transfers"
        Generate_Report("SORMTDV1", "Invoice Register", SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
        End Select
    End Sub

#End Region

    Private Function Prepare_Sales_Invoices(
        sqlw As String,
        ByRef SOTINVH1 As String,
        ByRef SOTINVH2 As String) As String

        ASCMAIN1.Progress("Building Work File")

        Dim rowGLTPARM2 As DataRow = Lookup("GLTPARM2", ASCMAIN1.CYP)
        Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")
        Dim NYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)

        ASCMAIN1.Progress("Building Customer Masterfile")
        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", Decode(E.CUST_CODE,NULL,'N','Y') EDI" & vbCrLf _
            & ", Decode(M.CUST_CODE,NULL,'N','Y') MULTI_STORE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & ", (Select Distinct CUST_CODE from EDTTRPM1 where EDI_STATUS = 'P' and EDI_DOC_NO = '810') E" & vbCrLf _
            & ", (Select CUST_CODE from ARTCUST2 where CUST_ADDR_TYPE = 'MK' group by CUST_CODE having COUNT (*) > 1) M" & vbCrLf _
            & " where E.CUST_CODE (+) = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and M.CUST_CODE (+) = ARTCUST1.CUST_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTCUST1", 1))

        ASCMAIN1.Progress("Building Consolidated Invoices")
        S.Length = 0
        S.AppendLine("SELECT DISTINCT SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        S.AppendLine("Sum (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) as TOTAL_SALES,")
        S.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_STAX, SOTINVH1.INV_TOTAL_AMOUNT,")
        S.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        S.AppendLine("SOTINVH1.INV_NO_CONS , SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO , SOTINVH1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE,")
        S.AppendLine("SOTINVH1.GST_TAX, SOTINVH1.ORDR_YYYYPP_UPDATED")
        S.AppendLine("FROM SOTINVH1, SOTINVH2, ICTSTYL1")
        S.AppendLine("WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
        S.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
        S.AppendLine("AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        S.AppendLine("AND SOTINVH1.INV_NO_CONS IS NOT NULL")
        S.AppendLine("AND SOTINVH1.INV_TYPE = 'I'")
        S.AppendLine(sqlw)
        S.AppendLine("GROUP BY")
        S.AppendLine("SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        S.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG,  SOTINVH1.INV_STAX, SOTINVH1.INV_TOTAL_AMOUNT,")
        S.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        S.AppendLine("SOTINVH1.INV_NO_CONS,  SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO, SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.GST_TAX, SOTINVH1.ORDR_YYYYPP_UPDATED")
        ASCMAIN1.sql = S.ToString
        Dim SOTINVHT As String = ASCMAIN1.Temp_Table

        S.Length = 0
        S.AppendLine("UPDATE " & SOTINVHT)
        S.AppendLine("SET H_SALES_DIVISION_CODE = SALES_DIVISION_CODE")
        S.AppendLine("WHERE H_SALES_DIVISION_CODE <> SALES_DIVISION_CODE")
        ASCMAIN1.sql = S.ToString
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Building Invoices")
        Dim SN As New System.Text.StringBuilder() With {.Length = 0}
        SN.AppendLine("SELECT DISTINCT  SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        SN.AppendLine("Sum (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) as TOTAL_SALES,")
        SN.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_STAX,  SOTINVH1.INV_TOTAL_AMOUNT,")
        SN.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        SN.AppendLine("SOTINVH1.INV_NO_CONS , SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO , SOTINVH1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE,")
        SN.AppendLine("SOTINVH1.GST_TAX")
        SN.AppendLine("FROM SOTINVH1, SOTINVH2, ICTSTYL1")
        SN.AppendLine("WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
        SN.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
        SN.AppendLine("AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        SN.AppendLine("AND SOTINVH1.INV_NO_CONS IS NULL")
        SN.AppendLine("AND SOTINVH1.INV_TYPE = 'I'")
        SN.AppendLine(sqlw)
        SN.AppendLine("GROUP BY")
        SN.AppendLine("SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        SN.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_STAX, SOTINVH1.INV_TOTAL_AMOUNT,")
        SN.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        SN.AppendLine("SOTINVH1.INV_NO_CONS,  SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO, SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.GST_TAX")

        Dim SC As New System.Text.StringBuilder() With {.Length = 0}
        SC.AppendLine("")
        SC.AppendLine("SELECT INV_TYPE,")
        SC.AppendLine("INV_NO_CONS AS INV_NO,")
        SC.AppendLine("CUST_CODE,")
        SC.AppendLine("'*CONS*' AS CUST_STORE_NO,")
        SC.AppendLine("MIN(ORDR_CUST_PO) AS ORDR_CUST_PO,")
        SC.AppendLine("MIN(ORDR_NO) AS ORDR_NO,")
        SC.AppendLine("SUM(TOTAL_SALES) AS TOTAL_SALES,")
        SC.AppendLine("SUM(NVL(INV_SALES,0)) AS INV_SALES,")
        SC.AppendLine("SUM(NVL(INV_FREIGHT,0)) AS INV_FREIGHT,")
        SC.AppendLine("SUM(NVL(INV_MISC_CHG,0)) AS INV_MISC_CHG,")
        SC.AppendLine("SUM(NVL(INV_STAX,0)) AS INV_STAX,")
        SC.AppendLine("SUM(NVL(INV_TOTAL_AMOUNT,0)) AS INV_TOTAL_AMOUNT,")
        SC.AppendLine("MIN(INV_DATE) AS INV_DATE,")
        SC.AppendLine("MIN(POST_CODE) AS POST_CODE,")
        SC.AppendLine("MIN(SHIP_BOL_NO) AS SHIP_BOL_NO,")
        SC.AppendLine("SALES_DIVISION_CODE,")
        SC.AppendLine("MIN(INV_NO_CONS) AS INV_NO_CONS,")
        SC.AppendLine("MIN(INIT_DATE) AS INIT_DATE,")
        SC.AppendLine("'*CONS*' AS PICK_NO,")
        SC.AppendLine("H_SALES_DIVISION_CODE,")
        SC.AppendLine("SUM(NVL(GST_TAX,0)) AS GST_TAX")
        SC.AppendLine("FROM " & SOTINVHT)
        SC.AppendLine("GROUP BY")
        SC.AppendLine("INV_TYPE, INV_NO_CONS, CUST_CODE, H_SALES_DIVISION_CODE, SALES_DIVISION_CODE")

        S.Length = 0
        S.AppendLine("SELECT INV_TYPE, INV_NO, CUST_CODE, CUST_STORE_NO, ORDR_CUST_PO, ORDR_NO,")
        S.AppendLine("TOTAL_SALES, INV_SALES, INV_FREIGHT, INV_MISC_CHG, INV_STAX, INV_TOTAL_AMOUNT,")
        S.AppendLine("INV_DATE, POST_CODE, SHIP_BOL_NO, SALES_DIVISION_CODE,")
        S.AppendLine("INV_NO_CONS, INIT_DATE, PICK_NO, H_SALES_DIVISION_CODE, GST_TAX")
        S.AppendLine("FROM (")
        S.AppendLine(SN.ToString)
        S.AppendLine("UNION")
        S.AppendLine(SC.ToString)
        S.AppendLine(")")
        ASCMAIN1.sql = S.ToString
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHZ", 0))
        ASCMAIN1.Progress("Building Consolidated Invoices", "Wrap-up")

        ASCMAIN1.Progress("Building Misc Data")
        ASCMAIN1.sql = "Select SOTSDIV1.* from SOTSDIV1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSDIV1", 1))

        Return ""
    End Function
End Class