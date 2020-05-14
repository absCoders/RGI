Imports System.Text

Public Class SORCOMM1

#Region "General Declarations"
    Dim xRYP0 As String
    Dim xRYP0_legend As String
    Dim SUBTITLE As String
    Dim CHKSUMMARY As String
    Dim CHKSREP2 As String
    Dim SOTINVH1_temp As String
    Dim SQ As New StringBuilder With {.Length = 0}

#End Region

#Region "ABSolution Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Get_SQL("*")
        'RWU = "R"
        Dim sqlw As String = ""

        xRYP0_legend = Absx1.cmbFor("RYP0").Value
        xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)

        SUBTITLE = "Commissions for period " & xRYP0_legend
        If Absx1.chkFor("CHKSREP2").Checked Then
            SUBTITLE = SUBTITLE & " (Second Sales Rep)"
        End If

        'Build Temp Table That Drive The Report
        SQ.Length = 0
        SQ.AppendLine("Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE,")
        If Absx1.chkFor("CHKSREP2").Checked Then
            SQ.AppendLine("SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.INV_NO_CONS, SOTINVH1.SREP2_CODE SREP_CODE, SOTINVH1.INV_SALES,")
        Else
            SQ.AppendLine("SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.INV_NO_CONS, SOTINVH1.SREP_CODE, SOTINVH1.INV_SALES,")
        End If
        SQ.AppendLine("(SOTINVH1.INV_SALES * (TATTERM1.TERM_DISC_PERC/100)) DISCOUNT,")
        SQ.AppendLine("(INV_SALES - (INV_SALES * (TATTERM1.TERM_DISC_PERC/100))) INV_NET,")
        SQ.AppendLine("DECODE(SOTSREP2.SREP_COMM_RATE, NULL,")
        SQ.AppendLine("DECODE(SOTSREP3.SREP_COMM_RATE, NULL, SOTSREP1.SREP_COMM_RATE,")
        SQ.AppendLine("SOTSREP3.SREP_COMM_RATE), SOTSREP2.SREP_COMM_RATE)")
        SQ.AppendLine("INV_COMM_PCT, ((INV_SALES - (INV_SALES * (TATTERM1.TERM_DISC_PERC/100))) *")
        SQ.AppendLine("DECODE(SOTSREP2.SREP_COMM_RATE, NULL,")
        SQ.AppendLine("DECODE(SOTSREP3.SREP_COMM_RATE, NULL, SOTSREP1.SREP_COMM_RATE,")
        SQ.AppendLine("SOTSREP3.SREP_COMM_RATE), SOTSREP2.SREP_COMM_RATE)/100) INV_COMM_AMT")
        SQ.AppendLine("From SOTINVH1, TATTERM1, SOTSREP1, SOTSREP2, SOTSREP3")
        SQ.AppendLine("Where SOTINVH1.TERM_CODE = TATTERM1.TERM_CODE")
        If Absx1.chkFor("CHKSREP2").Checked Then
            SQ.AppendLine("AND SOTINVH1.SREP2_CODE = SOTSREP1.SREP_CODE")
        Else
            SQ.AppendLine("AND SOTINVH1.SREP_CODE = SOTSREP1.SREP_CODE")
        End If
        SQ.AppendLine("AND SOTINVH1.CUST_CODE = SOTSREP2.CUST_CODE (+)")
        If Absx1.chkFor("CHKSREP2").Checked Then
            SQ.AppendLine("AND SOTINVH1.SREP2_CODE = SOTSREP2.SREP_CODE (+)")
        Else
            SQ.AppendLine("AND SOTINVH1.SREP_CODE = SOTSREP2.SREP_CODE (+)")
        End If
        SQ.AppendLine("AND SOTINVH1.SALES_DIVISION_CODE = SOTSREP2.SALES_DIVISION_CODE (+)")
        If Absx1.chkFor("CHKSREP2").Checked Then
            SQ.AppendLine("AND SOTINVH1.SREP2_CODE = SOTSREP3.SREP_CODE (+)")
        Else
            SQ.AppendLine("AND SOTINVH1.SREP_CODE = SOTSREP3.SREP_CODE (+)")
        End If
        SQ.AppendLine("AND SOTINVH1.SALES_DIVISION_CODE = SOTSREP3.SALES_DIVISION_CODE (+)")
        SQ.AppendLine("AND SOTINVH1.ORDR_YYYYPP_UPDATED = '" & xRYP0 & "'")
        SQ.AppendLine("AND SOTINVH1.INV_SALES <> 0")
        If Absx1.chkFor("CHKSREP2").Checked Then
            Dim sql_WHERE2 As String = sql_WHERE.Replace("SOTINVH1.SREP_CODE", "SOTINVH1.SREP2_CODE")
            SQ.AppendLine(sql_WHERE2)
        Else
            SQ.AppendLine(sql_WHERE)
        End If

        ASCMAIN1.sql = SQ.ToString
        SOTINVH1_temp = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Create Unique Index I_" & SOTINVH1_temp & "_1 on " & SOTINVH1_temp & " (INV_NO)")

        SQ.Length = 0
        SQ.AppendLine("Select SOTINVH1.* from " & SOTINVH1_temp & " SOTINVH1 WHERE SOTINVH1.INV_NO_CONS IS NULL")
        Create_TDA(dst.Tables.Add, "SOTINVH1", SQ.ToString, 0, False)
        Fill_Records("SOTINVH1")

        SQ.Length = 0
        SQ.AppendLine("Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO_CONS INV_NO,")
        SQ.AppendLine("MAX(SOTINVH1.CUST_CODE) CUST_CODE, MAX(SOTINVH1.SALES_DIVISION_CODE) SALES_DIVISION_CODE,")
        SQ.AppendLine("SOTINVH1.INV_NO_CONS INV_NO , MAX(SOTINVH1.SREP_CODE) SREP_CODE,")
        SQ.AppendLine("Sum (SOTINVH1.INV_SALES) INV_SALES, Sum(DISCOUNT) DISCOUNT, Sum(INV_NET) INV_NET, AVG(INV_COMM_PCT) INV_COMM_PCT, Sum(INV_COMM_AMT) INV_COMM_AMT")
        SQ.AppendLine("FROM " & SOTINVH1_temp & " SOTINVH1")
        SQ.AppendLine("Where SOTINVH1.INV_NO_CONS Is Not Null")
        SQ.AppendLine("And SOTINVH1.INV_SALES <> 0")
        SQ.AppendLine("GROUP BY SOTINVH1.INV_TYPE, SOTINVH1.INV_NO_CONS")
        Fill_Records("SOTINVH1", "", False, SQ.ToString)

        UpdateINVCOMMPCT()

        Call ASCMAIN1.Progress("", "Invoice Details")
        SQ.Length = 0
        SQ.AppendLine("SELECT SOTINVH2.INV_TYPE, SOTINVH1.INV_NO, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE,")
        SQ.AppendLine("SUM(NVL(SOTINVH2.ORDR_UNIT_COST,0) * NVL(SOTINVH2.ORDR_QTY_SHIP,0)) COSTS,")
        SQ.AppendLine("SUM(NVL(SOTINVH2.ORDR_UNIT_PRICE,0) * NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SALES,")
        SQ.AppendLine("SUM(NVL(SOTINVH2.ORDR_QTY_SHIP,0)) UNITS")
        SQ.AppendLine("FROM " & SOTINVH1_temp & " SOTINVH1, SOTINVH2")
        SQ.AppendLine("Where SOTINVH2.INV_NO = SOTINVH1.INV_NO")
        SQ.AppendLine("and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE")
        SQ.AppendLine("AND SOTINVH1.INV_NO_CONS IS NULL")
        If Absx1.chkFor("CHKSUMMARY").Checked Then
            SQ.AppendLine("and ROWNUM < 1")
        End If
        SQ.AppendLine("GROUP BY SOTINVH2.INV_TYPE, SOTINVH1.INV_NO, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
        Create_TDA(dst.Tables.Add, "SOTINVH2", SQ.ToString, 0, False)
        Fill_Records("SOTINVH2")

        If Not Absx1.chkFor("CHKSUMMARY").Checked Then
            SQ.Length = 0
            SQ.AppendLine("SELECT SOTINVH2.INV_TYPE, SOTINVH1.INV_NO_CONS INV_NO, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE,")
            SQ.AppendLine("SUM(NVL(SOTINVH2.ORDR_UNIT_COST,0) * NVL(SOTINVH2.ORDR_QTY_SHIP,0)) COSTS,")
            SQ.AppendLine("SUM(NVL(SOTINVH2.ORDR_UNIT_PRICE,0) * NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SALES,")
            SQ.AppendLine("SUM(NVL(SOTINVH2.ORDR_QTY_SHIP,0)) UNITS")
            SQ.AppendLine("FROM " & SOTINVH1_temp & " SOTINVH1, SOTINVH2")
            SQ.AppendLine("Where SOTINVH2.INV_NO = SOTINVH1.INV_NO")
            SQ.AppendLine("AND SOTINVH1.INV_NO_CONS IS NOT NULL")
            SQ.AppendLine("and ROWNUM < 1")
            SQ.AppendLine("GROUP BY SOTINVH2.INV_TYPE, SOTINVH1.INV_NO_CONS, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE")
            Fill_Records("SOTINVH2", "", False, SQ.ToString)
        End If
        'If you are bit in the ass during the conversion it could be because you didn't account for this:
        'Call Create_Index("SOWINVH2", "I_SOWINVH2_1", "INV_TYPE, INV_NO, STYLE_CODE, COLOR_CODE")

        Call ASCMAIN1.Progress("", "Cash Write-offs")
        SQ.Length = 0
        If Absx1.chkFor("CHKSREP2").Checked Then
            SQ.AppendLine("SELECT ARTCUST1.SREP2_CODE SREP_CODE, ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE,")
        Else
            SQ.AppendLine("SELECT ARTCUST1.SREP_CODE, ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE,")
        End If
        SQ.AppendLine("SUM(ARTPYMT5.GL_DIST_AMT) GL_DIST_AMT")
        SQ.AppendLine("From ARTPYMT1, ARTPYMT2, ARTPYMT5, ARTCUST1")
        SQ.AppendLine("Where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO")
        SQ.AppendLine("AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO")
        SQ.AppendLine("AND ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT5.PYMT_BATCH_LNO")
        SQ.AppendLine("AND ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE")
        SQ.AppendLine("AND OPS_YYYYPP = '" & xRYP0 & "'")
        'SQ.AppendLine("AND (ARTPYMT5.CHARGEBACK_IND = NULL OR ARTPYMT5.CHARGEBACK_IND <> -1)")
        SQ.AppendLine("AND NVL(ARTPYMT5.CHARGEBACK_IND,'0') <> '1'")
        SQ.AppendLine("AND ARTPYMT5.REASON_CODE IN ( SELECT REASON_CODE FROM ARTREAS1 WHERE COMMISSION_IND = '1')")
        'If Absx1.chkFor("CHKSREP2").Checked Then
        '    SQ.AppendLine("AND ARTCUST1.SREP2_CODE IS NOT NULL")
        'End If
        Dim sql_WHERE3 As String = sql_WHERE
        If Absx1.chkFor("CHKSREP2").Checked Then
            sql_WHERE3 = sql_WHERE3.Replace("SOTINVH1.SREP_CODE", "ARTCUST1.SREP2_CODE")
        Else
            sql_WHERE3 = sql_WHERE3.Replace("SOTINVH1.SREP_CODE", "ARTCUST1.SREP_CODE")
        End If
        'SQ.AppendLine(sql_WHERE3)
        If Absx1.chkFor("CHKSREP2").Checked Then
            SQ.AppendLine("GROUP BY ARTCUST1.SREP2_CODE, ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE")
        Else
            SQ.AppendLine("GROUP BY ARTCUST1.SREP_CODE, ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE")
        End If
        Create_TDA(dst.Tables.Add, "ARTCASHX", SQ.ToString, 0, False)
        Fill_Records("ARTCASHX")

        SQ.Length = 0
        SQ.AppendLine("SELECT SREP_CODE,")
        SQ.AppendLine("99999.99 AS INV_SALES,")
        SQ.AppendLine("99999.99 AS DISCOUNT,")
        SQ.AppendLine("99999.99 AS INV_COMM_AMT,")
        SQ.AppendLine("99999.99 As WRITE_OFF")
        SQ.AppendLine("FROM SOTINVH1")
        Create_TDA(dst.Tables.Add, "SOTCOMMS", SQ.ToString, 0, False, "", 1)
        Dim SREPS As New List(Of String)
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select()
            If Not SREPS.Contains(rowSOTINVH1.Item("SREP_CODE").ToString) Then
                SREPS.Add(rowSOTINVH1.Item("SREP_CODE").ToString())
            End If
        Next
        For Each srep As String In SREPS
            Dim filter As String = "SREP_CODE = '" & srep & "'"
            Dim newSOTCOMMS As DataRow = dst.Tables("SOTCOMMS").NewRow
            newSOTCOMMS.Item("SREP_CODE") = srep
            newSOTCOMMS.Item("INV_SALES") = dst.Tables("SOTINVH1").Compute("Sum(INV_SALES)", filter)
            newSOTCOMMS.Item("DISCOUNT") = dst.Tables("SOTINVH1").Compute("Sum(DISCOUNT)", filter)
            newSOTCOMMS.Item("INV_COMM_AMT") = dst.Tables("SOTINVH1").Compute("Sum(INV_COMM_AMT)", filter)
            newSOTCOMMS.Item("WRITE_OFF") = dst.Tables("ARTCASHX").Compute("Sum(GL_DIST_AMT)", filter)
            dst.Tables("SOTCOMMS").Rows.Add(newSOTCOMMS)
        Next

        SQ.Length = 0
        If Absx1.chkFor("CHKSREP2").Checked Then
            SQ.AppendLine("Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.SREP2_CODE SREP_CODE from ARTCUST1")
        Else
            SQ.AppendLine("Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE from ARTCUST1")
        End If
        Create_TDA(dst.Tables.Add, "ARTCUST1", SQ.ToString, 0, False, "", 1)
        Fill_Records("ARTCUST1")

        'For Each srep As String In SREPS
        '    Dim filter As String = "SREP_CODE = '" & srep & "'"
        '    Dim newSOTCOMMS As DataRow = dst.Tables("SOTCOMMS").NewRow
        '    newSOTCOMMS.Item("SREP_CODE") = srep
        '    newSOTCOMMS.Item("WRITE_OFF") = dst.Tables("ARWCASHX").Compute("Sum(GL_DIST_AMT)", filter)
        '    dst.Tables("SOTCOMMS").Rows.Add(newSOTCOMMS)
        'Next

        SQ.Length = 0
        SQ.AppendLine("Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_COST from ICTSTYL1")
        Create_TDA(dst.Tables.Add, "ICTSTYL1", SQ.ToString, 0, False, "", 1)
        Fill_Records("ICTSTYL1")

        SQ.Length = 0
        SQ.AppendLine("Select * from SOTSDIV1")
        Create_TDA(dst.Tables.Add, "SOTSDIV1", SQ.ToString, 0, False, "", 1)
        Fill_Records("SOTSDIV1")

        SQ.Length = 0
        SQ.AppendLine("Select * From SOTSREP1")
        Create_TDA(dst.Tables.Add, "SOTSREP1", SQ.ToString, 0, False, "", 1)
        Fill_Records("SOTSREP1")

        Check_if_Empty("SOTCOMMS")
    End Sub

    Public Overrides Sub Print_Report()
        RPT_TITLE = "Sales Rep Commission Report"
        'CR_params.Add("SUBT", SUBTITLE)
        'CR_params.Add("SUPPRESS_DETAIL", "0")
        Dim CHKSUMMARY As String = "0"
        If Absx1.chkFor("CHKSUMMARY").Checked Then
            CHKSUMMARY = "1"
        End If
        CR_params.Add("SUMMARY", CHKSUMMARY)

        'Set CR_SubRpt = CR_Rpt.OpenSubreport("SORCOMM2")
        'For i = 1 To CR_SubRpt.Database.Tables.Count
        '        CR_SubRpt.Database.Tables(i).LOCATION = WorkDBName
        '    Next i
        'Set CR_SubRpt = CR_Rpt.OpenSubreport("SORCOMM3")
        'For i = 1 To CR_SubRpt.Database.Tables.Count
        '        CR_SubRpt.Database.Tables(i).LOCATION = WorkDBName
        '    Next i
        'Set CR_SubRpt = Nothing
        Generate_Report("SORCOMM1", RPT_TITLE, SUBTITLE)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If Absx1.cmbFor("RYP0").Value & "" = "" Then
            EMsg &= vbCr & "You must Specify a Period"
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        'If Not Me.Visible Then Clear_dst()
        Return clsASCBASE1
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim workTableName As String = String.Empty

        ' Period to load
        If parms.Length > 0 Then
            If parms(0).ToString.Contains("-") Then
                xRYP0 = parms(0).ToString.Split("-")(0).Trim
                xRYP0_legend = ASCDATA1.GetDataValue("Select LEGEND FROM GLTPARM2 WHERE OPS_YYYYPP = '" & xRYP0 & "'") & String.Empty
            Else
                xRYP0 = parms(0)
                xRYP0_legend = ASCDATA1.GetDataValue("SELECT LEGEND FROM GLTPARM2 WHERE OPS_YYYYPP = '" & xRYP0 & "'") & String.Empty
            End If
        End If

        EnforceConstraints(False)

        SQ.Length = 0
        SQ.AppendLine("SELECT SREP_CODE, SREP_NAME FROM SOTSREP1")
        ASCMAIN1.sql = sql.ToString
        'Fill_Records("SOTSREP1", String.Empty, True, ASCMAIN1.sql)

        'Fill_Records("SOTCOMMT")

        ' Show Only summary for sales reps in the main datatable
        'For Each rowSOTCOMMT As DataRow In dst.Tables("SOTCOMMT").Select("")
        '    If dst.Tables("SOTCOMMS").Select("SREP_CODE = '" & rowSOTCOMMT.Item("SREP_CODE") & "'").Length = 0 Then
        '        rowSOTCOMMT.Delete()
        '    End If
        'Next
        'dst.Tables("SOTCOMMT").AcceptChanges()

        EnforceConstraints(True)

    End Sub
#End Region

#Region "Form Controls"

#End Region

#Region "Custom Methods"
    Private Sub UpdateINVCOMMPCT()
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select()
            rowSOTINVH1.Item("INV_COMM_PCT") = 0
            If Val(rowSOTINVH1.Item("INV_SALES") & "") <> 0 Then
                rowSOTINVH1.Item("INV_COMM_PCT") = Val(rowSOTINVH1.Item("INV_COMM_AMT") & "") / Val(rowSOTINVH1.Item("INV_NET") & "")
            End If
        Next
    End Sub
#End Region
End Class