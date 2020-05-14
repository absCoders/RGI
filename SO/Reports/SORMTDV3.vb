Public Class SORMTDV3

#Region "Declarations"
    Dim DTE0 As Date
    Dim DTE1 As Date
    Dim SOTINVH1 As String
    Dim SOTINVH2 As String
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
        If optASN.Value = "S" Then
            sqlw &= "   and ICTSTYL1.CUST_CODE is Null"
        End If
        If optASN.Value = "N" Then
            sqlw &= "   and ICTSTYL1.CUST_CODE is Not Null"
        End If
        If Absx1.chkFor("CHKXTRANSF").Checked Then
            sqlw &= "    and SOTINVH1.CUST_CODE <> 'TRANSFERS'"
        End If
        If Absx1.chkFor("CHKXSAMPLES").Checked Then
            sqlw &= "    and SOTINVH1.CUST_CODE <> 'SAMPLES'"
        End If
        'Performance.Clear()
        'PerformanceStart = Now
        Prepare_Sales_Invoices(sqlw, SOTINVH1, SOTINVH2)
        Check_if_Empty("SOTINVH1")
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = SUBT & " Excluding Samples & Transfers"
        Generate_Report("SORMTDV3", "Credit Register", SUBT)

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

        ASCMAIN1.Progress("Building Work File - SOTINVH1")
        ASCMAIN1.sql = "Select DISTINCT SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" & vbCrLf _
            & " from SOTINVH2, SOTINVH1, ICTSTYL1" & vbCrLf _
            & " where SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & sqlw

        SOTINVH2 = ASCMAIN1.Temp_Table
        ASCMAIN1.AnalyzeTable(SOTINVH2)

        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.CUST_DC_NO" & vbCrLf _
            & ", SOTORDR1.EDI_APPOINTMENT" & vbCrLf _
            & " from SOTINVH1, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & " AND SOTINVH1.INV_TYPE = 'C'" & vbCrLf _
            & "   and (INV_TYPE, INV_NO) in (Select Distinct INV_TYPE, INV_NO from " & SOTINVH2 & ")"
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE,INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_1 on " & SOTINVH1 & " (INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_2 on " & SOTINVH1 & " (PICK_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_3 on " & SOTINVH1 & " (ORDR_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_4 on " & SOTINVH1 & " (SHIP_BOL_NO)")
        ASCMAIN1.AnalyzeTable(SOTINVH1)

        ASCMAIN1.sql = "Select SOTINVH1.* from " & SOTINVH1 & " SOTINVH1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVH1", 2))

        ' Credits
        ASCMAIN1.Progress("Building Work File - SOTINVHR")
        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", NULL ORDR_GROUP_NO, NULL ORDR_ADDR_TYPE_ST, NULL CUST_DC_NO" & vbCrLf _
            & ", NULL EDI_APPOINTMENT" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = 'C'"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHR", 2))

        For Each rowSOTINVHR As DataRow In dst.Tables("SOTINVHR").Select()
            Dim INV_NO As String = rowSOTINVHR.Item("INV_NO").ToString & ""
            Dim SQLR As New System.Text.StringBuilder() With {.Length = 0}
            SQLR.AppendLine("SELECT MIN(SALES_DIVISION_CODE) AS SALES_DIVISION_CODE")
            SQLR.AppendLine("FROM SOTINVH2, ICTSTYL1")
            SQLR.AppendLine("WHERE SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            SQLR.AppendLine(String.Format("AND INV_NO = '{0}'", INV_NO))
            ASCMAIN1.sql = SQLR.ToString()
            rowSOTINVHR.Item("SALES_DIVISION_CODE") = ASCDATA1.GetDataValue
        Next

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

        ASCMAIN1.sql = "Select SOTSDIV1.* from SOTSDIV1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSDIV1", 1))

        ASCMAIN1.sql = "Select ARTREAS1.* from ARTREAS1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTREAS1", 1))

        Return ""
    End Function
End Class