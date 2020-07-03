Public Class SORMTDV2

#Region "Declarations"
    Dim DTE0 As Date
    Dim DTE1 As Date
    Dim SOTINVH1 As String
    Dim SOTINVH2 As String
    Dim S As New System.Text.StringBuilder With {.Length = 0}

#End Region

#Region "ABS Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("SOTPARM1")

        grpDATE_RANGE.Top = grpPERIOD_RANGE.Top
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -84, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        Dim AllowCosts As Boolean = InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") > 0
        With Absx1.chkFor("CHKCOSTS")
            .Checked = AllowCosts
            .Enabled = AllowCosts
            .Visible = AllowCosts
        End With
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        If Absx1.optFor("RANGE").Value = "D" Then
            DTE0 = Absx1.dteFor("DTE0").Value
            DTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(DTE0, DTE1) = 0 Then
                SUBT = "Invoices Dated " & Format(DTE0, "MM/dd/yyyy")
            Else
                SUBT = String.Format("Invoices Dated between {0} and {1}", Format(DTE0, "MM/dd/yyyy"), Format(DTE1, "MM/dd/yyyy"))
            End If
        Else
            RYPLEGEND0 = Absx1.cmbFor("RYP0").Value
            RYP0 = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
            RYPLEGEND1 = Absx1.cmbFor("RYP1").Value
            RYP1 = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)
            If RYP0 = RYP1 Then
                SUBT = "Invoices Posted in " & RYPLEGEND0
            Else
                SUBT = String.Format("Invoices Posted between {0} and {1}", RYPLEGEND0, RYPLEGEND1)
            End If
        End If

        Dim sqlw As String = ""
        If optRANGE.Value = "D" Then
            sqlw &= String.Format("   and SOTINVH1.INV_DATE >= '{0}'{1}", Format(DTE0, "dd-MMM-yyyy"), vbCrLf)
            sqlw &= String.Format("   and SOTINVH1.INV_DATE <= '{0}'{1}", Format(DTE1, "dd-MMM-yyyy"), vbCrLf)
        ElseIf optRANGE.Value = "P" Then
            sqlw &= String.Format("   and SOTINVH1.ORDR_YYYYPP_UPDATED between '{0}' and '{1}'{2}", RYP0, RYP1, vbCrLf)
        End If

        sqlw &= SQL_in("SALES_DIVISION_CODE", "ICTSTYL1.SALES_DIVISION_CODE")
        sqlw &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE")
        If optASN.Value = "S" Then
            sqlw &= "   and ICTSTYL1.CUST_CODE is Null"
        End If
        If optASN.Value = "N" Then
            sqlw &= "   and ICTSTYL1.CUST_CODE is Not Null"
        End If
        sqlw &= "    and SOTINVH1.CUST_CODE <> 'TRANSFERS'"
        sqlw &= "    and SOTINVH1.CUST_CODE <> 'SAMPLES'"

        Prepare_Sales_Invoices(sqlw, SOTINVH1, SOTINVH2)
        Check_if_Empty("SOTINVHC")
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = SUBT & " Excluding Samples & Transfers"

        If Absx1.chkFor("CHKSSBC").Checked Then
            RPT_TITLE = "Sales Summary By Customer"
            CR_params.Add("CHKCOSTS", IIf(Absx1.chkFor("CHKCOSTS").Checked, "1", "0"))
            Generate_Report("SORMTDV2", RPT_TITLE, SUBT)
        End If

        If Absx1.chkFor("CHKSBC").Checked Then
            RPT_TITLE = "Sales By Customer"
            CR_params.Add("CHKCOSTS", IIf(Absx1.chkFor("CHKCOSTS").Checked, "1", "0"))
            Generate_Report("SORMTDVA", RPT_TITLE, SUBT)
        End If

        If Absx1.chkFor("CHKDSBCR").Checked Then
            RPT_TITLE = "Daily Sales By Customer Recap"
            'CR_params.Add("CHKCOSTS", "1")
            CR_params.Add("CHKCOSTS", IIf(Absx1.chkFor("CHKCOSTS").Checked, "1", "0"))
            Generate_Report("SORMTDVB", RPT_TITLE, SUBT)
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                If Absx1.optFor("RANGE").Value = "P" Then
                    If Absx1.cmbFor("RYP0").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify a Starting Period"
                    End If
                    If Absx1.cmbFor("RYP1").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify an Ending Period"
                    End If
                Else
                    If Absx1.dteFor("DTE0").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify a Starting Date"
                    End If
                    If Absx1.dteFor("DTE1").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify an Ending Date"
                    End If
                End If
                If Absx1.chkFor("CHKSBC").Checked = False _
                    And Absx1.chkFor("CHKSSBC").Checked = False _
                    And Absx1.chkFor("CHKDSBCR").Checked = False Then
                    EMsg &= vbCr & "You must Pick At Least One Report"
                End If
        End Select
    End Sub

#End Region

    Private Function Prepare_Sales_Invoices(
        sqlw As String,
        ByRef SOTINVH1 As String,
        ByRef SOTINVH2 As String) As String

        ASCMAIN1.Progress("Building Work Files")

        Dim rowGLTPARM2 As DataRow = Lookup("GLTPARM2", ASCMAIN1.CYP)
        Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")
        Dim NYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1)

        ASCMAIN1.Progress("Building Work File - SOTINVH2")
        ASCMAIN1.sql = "Select SOTINVH2.*" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE as ORDR_AMT_SHIP" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST as ORDR_CGS_SHIP" & vbCrLf _
            & " from SOTINVH2, SOTINVH1, ICTSTYL1" & vbCrLf _
            & " where SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and  SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
            & sqlw

        SOTINVH2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add Primary Key (INV_TYPE,INV_NO,INV_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_QTY_CANC NUMBER(6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_AMT_CANC NUMBER(13,2)")
        ASCMAIN1.AnalyzeTable(SOTINVH2)

        ASCMAIN1.Progress("Building Work File - ICTSTYL1")
        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC" _
            & ", ICTSTYL1.STYLE_COST, ICTSTYL1.SALES_DIVISION_CODE from ICTSTYL1" _
            & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTINVH2 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTYL1", 1))

        ASCMAIN1.Progress("Building Work File - ARTCUST1")
        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", Decode(E.CUST_CODE,NULL,'N','Y') EDI" & vbCrLf _
            & ", Decode(M.CUST_CODE,NULL,'N','Y') MULTI_STORE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & ", (Select Distinct CUST_CODE from EDTTRPM1 where EDI_STATUS = 'P' and EDI_DOC_NO = '810') E" & vbCrLf _
            & ", (Select CUST_CODE from ARTCUST2 where CUST_ADDR_TYPE = 'MK' group by CUST_CODE having COUNT (*) > 1) M" & vbCrLf _
            & " where E.CUST_CODE (+) = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and M.CUST_CODE (+) = ARTCUST1.CUST_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTCUST1", 1))

        ASCMAIN1.Progress("Building Work File - SOTINVHC")
        S.Length = 0
        S.AppendLine("Select SALES_DIVISION_CODE, CUST_CODE")
        S.AppendLine(", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS")
        S.AppendLine(", Sum (ORDR_AMT_SHIP) as TOTAL_SALES")
        S.AppendLine(", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS")
        S.AppendLine(", Sum (TARIFF_UNIT_COST) AS TARIFF_UNIT_COST")
        S.AppendLine(" from " & SOTINVH2)
        S.AppendLine(" GROUP BY SALES_DIVISION_CODE, CUST_CODE")
        ASCMAIN1.sql = S.ToString
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHC", 0))

        ASCMAIN1.Progress("Building Work File - SOTINVHD")
        S.Length = 0
        S.AppendLine("Select INV_DATE, SALES_DIVISION_CODE, CUST_CODE")
        S.AppendLine(", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS")
        S.AppendLine(", Sum (ORDR_AMT_SHIP) as TOTAL_SALES")
        S.AppendLine(", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS")
        S.AppendLine(", Sum (TARIFF_UNIT_COST) AS TARIFF_UNIT_COST")
        S.AppendLine(", Max (TARIFF_FLAG) AS TARIFF_FLAG")
        S.AppendLine(", '   ' AS TARIFF_IND")
        S.AppendLine(" from " & SOTINVH2)
        S.AppendLine(" GROUP BY INV_DATE, SALES_DIVISION_CODE, CUST_CODE")
        ASCMAIN1.sql = S.ToString
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHD", 0))
        dst.Tables("SOTINVHD").Columns.Item("TARIFF_IND").ReadOnly = False
        For Each rowSOTINVHD As DataRow In dst.Tables("SOTINVHD").Select()
            Dim TARIFF_IND As String = ""
            If (rowSOTINVHD.Item("TARIFF_FLAG").ToString & String.Empty).Length = 10 Then
                Dim TARIFF_PRE = (rowSOTINVHD.Item("TARIFF_FLAG").ToString & String.Empty).Substring(8, 2)
                If TARIFF_PRE <> "00" Then
                    TARIFF_IND = String.Format("T{0}", TARIFF_PRE)
                End If
            End If
            rowSOTINVHD.Item("TARIFF_IND") = TARIFF_IND
        Next

        ' Master Files
        ASCMAIN1.Progress("Building Work File - Master Files")
        ASCMAIN1.sql = "Select ARTREAS1.* from ARTREAS1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTREAS1", 1))

        ASCMAIN1.sql = "Select SOTSDIV1.* from SOTSDIV1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSDIV1", 1))

        Return ""
    End Function

    Private Sub optRANGE_ValueChanged(sender As Object, e As EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        Else
            'Absx1.cmbFor("RYP0").Value = ""
            'Absx1.cmbFor("RYP1").Value = ""
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub
End Class