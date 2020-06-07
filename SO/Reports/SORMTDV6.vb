Public Class SORMTDV6

#Region "Declarations"
    Dim DTE0 As Date
    Dim DTE1 As Date
    Dim SOTINVH1 As String
    Dim SOTINVH2 As String
    Dim Performance As New Dictionary(Of String, Long)
    Dim PerformanceStart As Date
#End Region

#Region "ABS Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("SOTPARM1")

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

        DTE0 = Absx1.dteFor("DTE0").Value
        DTE1 = Absx1.dteFor("DTE1").Value
        If System.DateTime.Compare(DTE0, DTE1) = 0 Then
            SUBT = "Invoices Dated " & Format(DTE0, "MM/dd/yyyy")
        Else
            SUBT = String.Format("Invoices Dated between {0} and {1}", Format(DTE0, "MM/dd/yyyy"), Format(DTE1, "MM/dd/yyyy"))
        End If

        'RYPLEGEND0 = Absx1.cmbFor("RYP0").Value
        'RYP0 = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)
        'RYPLEGEND1 = Absx1.cmbFor("RYP1").Value
        'RYP1 = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)
        'If RYP0 = RYP1 Then
        '    SUBT = "Invoices Posted in " & RYPLEGEND0
        'Else
        '    SUBT = String.Format("Invoices Posted between {0} and {1}", RYPLEGEND0, RYPLEGEND1)
        'End If

        Dim sqlw As String = ""
        'sqlw &= String.Format("   and SOTINVH1.ORDR_YYYYPP_UPDATED between '{0}' and '{1}'{2}", RYP0, RYP1, vbCrLf)
        sqlw &= String.Format("   and SOTINVH1.INV_DATE >= '{0}'{1}", Format(DTE0, "dd-MMM-yyyy"), vbCrLf)
        sqlw &= String.Format("   and SOTINVH1.INV_DATE <= '{0}'{1}", Format(DTE1, "dd-MMM-yyyy"), vbCrLf)


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
        Performance.Clear()
        PerformanceStart = Now
        Prepare_Sales_Invoices(sqlw, SOTINVH1, SOTINVH2)
        Check_if_Empty("SOTINVHD")
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = SUBT & " Excluding Samples & Transfers"
        'CR_params.Add("CHKCOSTS", "1")
        CR_params.Add("CHKCOSTS", IIf(Absx1.chkFor("CHKCOSTS").Checked, "1", "0"))
        Generate_Report("SORMTDV6", "Shipment Summary By Date", SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        DTE0 = Absx1.dteFor("DTE0").Value
        DTE1 = Absx1.dteFor("DTE1").Value
        Select Case eItemKey
            Case "Proceed"
                'If Absx1.cmbFor("RYP0").Value & "" = "" Then
                '    EMsg &= vbCr & "You must Specify a Starting Period"
                'End If
                'If Absx1.cmbFor("RYP1").Value & "" = "" Then
                '    EMsg &= vbCr & "You must Specify an Ending Period"
                'End If
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
                If DTE1 < DTE0 Then
                    EMsg &= vbCr & "Invalid Date Range Selected."
                End If
                If EMsg.Length = 0 Then
                    If DateDiff(DateInterval.Day, DTE0, DTE1) > 7 Then
                        EMsg &= vbCr & "No More Than 7 Day Range Allowed."
                    End If
                End If
                Dim MTHERR As String = SOCMAINL.SalesReportCanRun(CDate(Absx1.dteFor("DTE0").Value),
                                                                  CDate(Absx1.dteFor("DTE1").Value),
                                                                  True, True)
                If MTHERR.Length > 0 Then
                    EMsg &= vbCr & MTHERR
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
        ASCMAIN1.sql = "Select SOTINVH2.*" & vbCrLf _
            & ", ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE as ORDR_AMT_SHIP" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST as ORDR_CGS_SHIP" & vbCrLf _
            & " from SOTINVH2, SOTINVH1, ICTSTYL1" & vbCrLf _
            & " where SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & sqlw

        SOTINVH2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add Primary Key (INV_TYPE,INV_NO,INV_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_QTY_CANC NUMBER(6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add ORDR_AMT_CANC NUMBER(13,2)")
        ASCMAIN1.AnalyzeTable(SOTINVH2)

        ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
            & ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.CUST_DC_NO" & vbCrLf _
            & ", SOTORDR1.EDI_APPOINTMENT" & vbCrLf _
            & " from SOTINVH1, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and (INV_TYPE, INV_NO) in (Select Distinct INV_TYPE, INV_NO from " & SOTINVH2 & ")"
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE,INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_1 on " & SOTINVH1 & " (INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_2 on " & SOTINVH1 & " (PICK_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_3 on " & SOTINVH1 & " (ORDR_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_4 on " & SOTINVH1 & " (SHIP_BOL_NO)")
        ASCMAIN1.AnalyzeTable(SOTINVH1)

        'ASCMAIN1.sql = "Select SOTINVH1.* from " & SOTINVH1 & " SOTINVH1"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVH1", 2))

        '' Credits
        'ASCMAIN1.Progress("Building Work File - SOTINVHR")
        'ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
        '    & ", NULL ORDR_GROUP_NO, NULL ORDR_ADDR_TYPE_ST, NULL CUST_DC_NO" & vbCrLf _
        '    & ", NULL EDI_APPOINTMENT" & vbCrLf _
        '    & " from " & SOTINVH1 & " SOTINVH1" & vbCrLf _
        '    & " where SOTINVH1.INV_TYPE = 'C'"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHR", 2))

        'For Each rowSOTINVHR As DataRow In dst.Tables("SOTINVHR").Select()
        '    Dim INV_NO As String = rowSOTINVHR.Item("INV_NO").ToString & ""
        '    Dim SQLR As New System.Text.StringBuilder() With {.Length = 0}
        '    SQLR.AppendLine("SELECT MIN(SALES_DIVISION_CODE) AS SALES_DIVISION_CODE")
        '    SQLR.AppendLine("FROM SOTINVH2, ICTSTYL1")
        '    SQLR.AppendLine("WHERE SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        '    SQLR.AppendLine(String.Format("AND INV_NO = '{0}'", INV_NO))
        '    ASCMAIN1.sql = SQLR.ToString()
        '    rowSOTINVHR.Item("SALES_DIVISION_CODE") = ASCDATA1.GetDataValue
        'Next

        Dim SQLC As New System.Text.StringBuilder() With {.Length = 0}
        SQLC.AppendLine(String.Format("DELETE FROM {0}", SOTINVH2))
        SQLC.AppendLine("WHERE INV_NO IN (")
        SQLC.AppendLine("SELECT INV_NO")
        SQLC.AppendLine(String.Format("FROM {0}", SOTINVH1))
        SQLC.AppendLine("WHERE INV_TYPE = 'C'")
        SQLC.AppendLine(")")
        ASCMAIN1.sql = SQLC.ToString
        ASCDATA1.ExecuteSQL()

        SQLC.Length = 0
        SQLC.AppendLine(String.Format("DELETE FROM {0}", SOTINVH1))
        SQLC.AppendLine("WHERE INV_TYPE = 'C'")
        ASCMAIN1.sql = SQLC.ToString
        ASCDATA1.ExecuteSQL()

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

        'ASCMAIN1.Progress("Building Work File - SOTFPCT1")
        'ASCMAIN1.sql = "" _
        '    & "Select SOTFPCT1.OPS_YYYYPP, SOTFPCT1.CUST_FACTOR_PERCENT, SOTFPCT1.CUST_SURCHARGE_PERCENT" & vbCrLf _
        '    & " from SOTFPCT1" & vbCrLf _
        '    & " union " & vbCrLf _
        '    & "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT" & vbCrLf _
        '    & " from SOTPARM1 where SO_PARM_KEY = 'Z'" & vbCrLf _
        '    & " union " & vbCrLf _
        '    & "Select '" & NYP & "' OPS_YYYYPP, SO_PARM_FACTOR_PCT CUST_FACTOR_PERCENT, SO_PARM_SURCHARGE_PCT CUST_SURCHARGE_PERCENT" & vbCrLf _
        '    & " from SOTPARM1 where SO_PARM_KEY = 'Z'"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTFPCT1", 1))

        ASCMAIN1.Progress("Building Work File - Report Summaries")
        'ASCMAIN1.Progress("-", "Report Summaries")
        ASCMAIN1.sql = "Select SOTINVH2.SALES_DIVISION_CODE, SOTINVH1.CUST_CODE, SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH1.SREP_CODE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','G:' || SOTORDR1.ORDR_GROUP_NO, 'S:' || SOTINVH1.SHIP_BOL_NO) AS SHIP_BOL_NO_X" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','MK','DC') AS SHIP_ADDR_TYPE" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','000000',SOTORDR1.CUST_DC_NO) AS SHIP_ADDR_CODE" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
            & ", Sum (SOTINVH2.ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
            & ", Sum (SOTINVH2.ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
            & ", Sum (SOTINVH2.ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
            & ", Sum (SOTINVH2.TARIFF_UNIT_COST) as TARIFF_UNIT_COST" & vbCrLf _
            & ", MAX (SOTINVH2.TARIFF_FLAG) as TARIFF_FLAG" & vbCrLf _
            & "  from SOTINVH1," & SOTINVH2 & " SOTINVH2, SOTORDR1" & vbCrLf _
            & "  where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "    and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "    and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
            & " group by SOTINVH2.SALES_DIVISION_CODE, SOTINVH1.CUST_CODE,  SOTINVH1.INV_DATE" & vbCrLf _
            & ", SOTINVH1.SREP_CODE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','G:' || SOTORDR1.ORDR_GROUP_NO, 'S:' || SOTINVH1.SHIP_BOL_NO)" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','MK','DC')" & vbCrLf _
            & ", DECODE (SOTORDR1.ORDR_ADDR_TYPE_ST, 'MK','000000',SOTORDR1.CUST_DC_NO)" & vbCrLf _
            & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf
        Dim SOTINVHD As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add ORDR_QTY_CANC NUMBER (6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add ORDR_AMT_CANC NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add LAST_RCD VARCHAR2(8)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " Add TARIFF_IND VARCHAR2(3)")
        ASCMAIN1.sql = "Update " & SOTINVHD & " Set SHIP_BOL_NO_X = 'G:' || ORDR_GROUP_NO"
        ASCDATA1.ExecuteSQL()

        Dim sq As New Text.StringBuilder With {.Length = 0}
        sq.AppendLine("SELECT")
        sq.AppendLine("SOTINVH2.SALES_DIVISION_CODE,")
        sq.AppendLine("SOTORDR1.ORDR_GROUP_NO,")
        sq.AppendLine("SOTINVH2.STYLE_CODE,")
        sq.AppendLine("SOTINVH2.COLOR_CODE")
        sq.AppendLine("from SOTINVH1, " & SOTINVH2 & " SOTINVH2, SOTORDR1")
        sq.AppendLine("where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
        sq.AppendLine("and SOTINVH1.INV_NO = SOTINVH2.INV_NO")
        sq.AppendLine("and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO")
        sq.AppendLine("GROUP BY")
        sq.AppendLine("SOTINVH2.SALES_DIVISION_CODE,")
        sq.AppendLine("SOTORDR1.ORDR_GROUP_NO,")
        sq.AppendLine("SOTINVH2.STYLE_CODE,")
        sq.AppendLine("SOTINVH2.COLOR_CODE")
        ASCMAIN1.sql = sq.ToString
        Dim CANCELS As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("ALTER TABLE " & CANCELS & " MODIFY ORDR_GROUP_NO VARCHAR2(10) NOT NULL")
        ASCDATA1.ExecuteSQL("ALTER TABLE " & CANCELS & " MODIFY SALES_DIVISION_CODE VARCHAR2(6) NOT NULL")
        ASCDATA1.ExecuteSQL("ALTER TABLE " & CANCELS & " MODIFY STYLE_CODE VARCHAR2(12) NOT NULL")
        ASCDATA1.ExecuteSQL("ALTER TABLE " & CANCELS & " MODIFY COLOR_CODE VARCHAR2(6) NOT NULL")
        ASCDATA1.ExecuteSQL("ALTER TABLE " & CANCELS & " ADD PRIMARY KEY (SALES_DIVISION_CODE,ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE)")

        'ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVHD & "_1 on " & SOTINVHD & " (ORDR_GROUP_NO, SALES_DIVISION_CODE, STYLE_CODE, COLOR_CODE)")
        'ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVHD & "_1 on " & SOTINVHD & " (ORDR_GROUP_NO, SALES_DIVISION_CODE, STYLE_CODE, COLOR_CODE)")
        'ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVHD & "_2 on " & SOTINVHD & " (ORDR_QTY_CANC)")
        'ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVHD & "_3 on " & SOTINVHD & " (STYLE_CODE)")
        'ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVHD & "_4 on " & SOTINVHD & " (COLOR_CODE)")
        'ASCMAIN1.AnalyzeTable(SOTINVHD)

        ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is Select SOTINVHD.ORDR_GROUP_NO, SOTINVHD.SALES_DIVISION_CODE, SOTINVHD.STYLE_CODE, SOTINVHD.COLOR_CODE" _
                & "  , SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC" _
                & "  , SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_CANC" _
                & "  from SOTORDR1,SOTORDR2," & CANCELS & " SOTINVHD" _
                & "  where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                & "    and SOTORDR2.ORDR_QTY_CANC <> 0" _
                & "    and SOTORDR1.ORDR_GROUP_NO = SOTINVHD.ORDR_GROUP_NO" _
                & "    AND SOTORDR2.STYLE_CODE = SOTINVHD.STYLE_CODE" _
                & "    AND SOTORDR2.COLOR_CODE = SOTINVHD.COLOR_CODE" _
                & "  group by SOTINVHD.ORDR_GROUP_NO, SOTINVHD.SALES_DIVISION_CODE, SOTINVHD.STYLE_CODE, SOTINVHD.COLOR_CODE;" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Begin" _
                & "    Update " & SOTINVHD & " Set " _
                & "      ORDR_QTY_CANC = R1.ORDR_QTY_CANC" _
                & "     ,ORDR_AMT_CANC = R1.ORDR_AMT_CANC" _
                & "     where SALES_DIVISION_CODE = R1.SALES_DIVISION_CODE" _
                & "     AND ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
                & "     AND STYLE_CODE = R1.STYLE_CODE" _
                & "     AND COLOR_CODE = R1.COLOR_CODE;" _
                & "   End;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Building Work File - SOTINVHD")
        ASCMAIN1.sql = "Select * from " & SOTINVHD
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHD", 0))

        ASCMAIN1.Progress("Building Work File - SOTINVHN")
        Dim S As New System.Text.StringBuilder() With {.Length = 0}
        S.AppendLine("SELECT 'G:' || SOTORDR1.ORDR_GROUP_NO AS SHIP_BOL_NO_X,")
        S.AppendLine("ICTSTYL1.SALES_DIVISION_CODE,")
        S.AppendLine("SOTORDR1.CUST_CODE,")
        S.AppendLine("MAX(X.INV_DATE) INV_DATE,")
        S.AppendLine("SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0)) AS QTY_CANC,")
        S.AppendLine("SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) AS AMT_CANC")
        S.AppendLine("FROM SOTORDR1, SOTORDR2, ICTSTYL1, (SELECT DISTINCT O1.ORDR_GROUP_NO, I1.INV_DATE FROM SOTINVH1 I1, SOTORDR1 O1 WHERE I1.ORDR_NO = O1.ORDR_NO) X")
        S.AppendLine("WHERE SOTORDR1.ORDR_GROUP_NO = X.ORDR_GROUP_NO")
        S.AppendLine("AND SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO")
        S.AppendLine("AND SOTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        'S.AppendLine(String.Format("AND ORDR_GROUP_NO IN (SELECT DISTINCT ORDR_GROUP_NO FROM {0})", SOTINVH1))
        S.AppendLine(String.Format(" AND X.INV_DATE >= '{0}'", Format(DTE0, "dd-MMM-yyyy")))
        S.AppendLine(String.Format(" AND X.INV_DATE <= '{0}'", Format(DTE1, "dd-MMM-yyyy")))
        S.AppendLine("AND SOTORDR1.ORDR_GROUP_NO IN (SELECT DISTINCT ORDR_GROUP_NO FROM SOTORDR1 WHERE ORDR_NO IN (")
        S.AppendLine("SELECT ORDR_NO FROM " & SOTINVH1)
        S.AppendLine("))")
        S.AppendLine("GROUP BY")
        S.AppendLine("'G:' || SOTORDR1.ORDR_GROUP_NO,")
        S.AppendLine("ICTSTYL1.SALES_DIVISION_CODE,")
        S.AppendLine("SOTORDR1.CUST_CODE")
        'S.AppendLine("X.INV_DATE")
        ASCMAIN1.sql = S.ToString
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHN", 0))

        'ASCMAIN1.Progress("Building Work File - SOTINVHC")
        'ASCMAIN1.sql = "Select INV_DATE, SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
        '    & ", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
        '    & ", Sum (ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
        '    & ", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
        '    & ", Sum (TARIFF_UNIT_COST) AS TARIFF_UNIT_COST" & vbCrLf _
        '    & " from " & SOTINVH2 & vbCrLf _
        '    & " group by INV_DATE, SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHC", 0))

        'ASCMAIN1.Progress("Building Work File - SOTINVHY")
        'ASCMAIN1.sql = "Select INV_DATE, SALES_DIVISION_CODE, CUST_CODE" & vbCrLf _
        '    & ", STYLE_CODE, COLOR_CODE" & vbCrLf _
        '    & ", Sum (ORDR_QTY_SHIP) as TOTAL_UNITS" & vbCrLf _
        '    & ", Sum (ORDR_AMT_SHIP) as TOTAL_SALES" & vbCrLf _
        '    & ", Sum (ORDR_CGS_SHIP) as TOTAL_COSTS" & vbCrLf _
        '    & " from " & SOTINVH2 & vbCrLf _
        '    & " group by INV_DATE, SALES_DIVISION_CODE, CUST_CODE, STYLE_CODE, COLOR_CODE"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHY", 0))

        'ASCMAIN1.Progress("Building Work File - SOTORDR0")
        'ASCMAIN1.sql = "Select * from SOTORDR0 where ORDR_GROUP_NO in" & vbCrLf _
        '    & " (Select DISTINCT ORDR_GROUP_NO from " & SOTINVH1 & ")"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTORDR0", 0))

        'Performance.Add("000003", DateDiff(DateInterval.Second, PerformanceStart, Now()))
        'ASCMAIN1.Progress("-", "Consolidated Invoices")
        'S.Length = 0
        'S.AppendLine("SELECT DISTINCT  SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        'S.AppendLine("Sum (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) as TOTAL_SALES,")
        'S.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG,  SOTINVH1.INV_TOTAL_AMOUNT,")
        'S.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        'S.AppendLine("SOTINVH1.INV_NO_CONS , SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO , SOTINVH1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE,")
        'S.AppendLine("SOTINVH1.GST_TAX")
        'S.AppendLine("FROM SOTINVH1, SOTINVH2, ICTSTYL1")
        'S.AppendLine("WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
        'S.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
        'S.AppendLine("AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        'S.AppendLine("AND SOTINVH1.INV_NO_CONS IS NULL")
        'S.AppendLine("AND SOTINVH1.INV_TYPE = 'I'")
        'S.AppendLine(sqlw)
        'S.AppendLine("GROUP BY")
        'S.AppendLine("SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        'S.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG,  SOTINVH1.INV_TOTAL_AMOUNT,")
        'S.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        'S.AppendLine("SOTINVH1.INV_NO_CONS,  SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO, SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.GST_TAX")
        'ASCMAIN1.sql = S.ToString
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHZ", 0))

        'Performance.Add("000004", DateDiff(DateInterval.Second, PerformanceStart, Now()))

        'S.Length = 0
        'S.AppendLine("SELECT DISTINCT SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        'S.AppendLine("Sum (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) as TOTAL_SALES,")
        'S.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG,  SOTINVH1.INV_TOTAL_AMOUNT,")
        'S.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        'S.AppendLine("SOTINVH1.INV_NO_CONS , SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO , SOTINVH1.SALES_DIVISION_CODE as H_SALES_DIVISION_CODE,")
        'S.AppendLine("SOTINVH1.GST_TAX")
        'S.AppendLine("FROM SOTINVH1, SOTINVH2, ICTSTYL1")
        'S.AppendLine("WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
        'S.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
        'S.AppendLine("AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        'S.AppendLine("AND SOTINVH1.INV_NO_CONS IS NOT NULL")
        'S.AppendLine("AND SOTINVH1.INV_TYPE = 'I'")
        'S.AppendLine(sqlw)
        'S.AppendLine("GROUP BY")
        'S.AppendLine("SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.ORDR_NO,")
        'S.AppendLine("SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT,  SOTINVH1.INV_MISC_CHG,  SOTINVH1.INV_TOTAL_AMOUNT,")
        'S.AppendLine("SOTINVH1.INV_DATE, SOTINVH1.POST_CODE, SOTINVH1.SHIP_BOL_NO, ICTSTYL1.SALES_DIVISION_CODE,")
        'S.AppendLine("SOTINVH1.INV_NO_CONS,  SOTINVH1.INIT_DATE, SOTINVH1.PICK_NO, SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.GST_TAX")
        'ASCMAIN1.sql = S.ToString
        'Dim SOTINVHT As String = ASCMAIN1.Temp_Table

        'S.Length = 0
        'S.AppendLine("UPDATE " & SOTINVHT)
        'S.AppendLine("SET H_SALES_DIVISION_CODE = SALES_DIVISION_CODE")
        'S.AppendLine("WHERE H_SALES_DIVISION_CODE <> SALES_DIVISION_CODE")
        'ASCMAIN1.sql = S.ToString
        'ASCDATA1.ExecuteSQL()

        'S.Length = 0
        'S.AppendLine("SELECT INV_TYPE, INV_NO, SUM(TOTAL_SALES) AS INV_SALES_NEW")
        'S.AppendLine("FROM " & SOTINVHT)
        'S.AppendLine("GROUP BY  INV_TYPE, INV_NO")
        'Dim tbl As DataTable = ASCDATA1.GetDataTable(S.ToString(), String.Empty, "V")
        'For Each rowSOTINVHT As DataRow In tbl.Rows
        '    Dim su As New System.Text.StringBuilder() With {.Length = 0}
        '    su.AppendLine(String.Format("Update {0} SET INV_SALES = {1}", SOTINVHT, Val(rowSOTINVHT.Item("INV_SALES_NEW") & "")))
        '    su.AppendLine(String.Format("WHERE INV_TYPE = '{0}'", rowSOTINVHT.Item("INV_TYPE")))
        '    su.AppendLine(String.Format("AND INV_NO = '{0}'", rowSOTINVHT.Item("INV_NO")))
        '    ASCMAIN1.sql = su.ToString
        '    ASCDATA1.ExecuteSQL()
        'Next

        'Performance.Add("000005", DateDiff(DateInterval.Second, PerformanceStart, Now()))

        'ASCMAIN1.Progress("Building Work File - SOTINVHX")
        'S.Length = 0
        'S.AppendLine("SELECT INV_NO_CONS,  CUST_CODE, SALES_DIVISION_CODE, H_SALES_DIVISION_CODE,")
        'S.AppendLine("MAX(INV_DATE) AS INV_DT,")
        'S.AppendLine("MAX(ORDR_CUST_PO) AS ORDR_PO,")
        'S.AppendLine("SUM(TOTAL_SALES) as TOT_SALES,")
        'S.AppendLine("SUM(INV_SALES) AS INV_SALE,")
        'S.AppendLine("SUM(INV_FREIGHT) AS INV_FR,")
        'S.AppendLine("SUM (INV_MISC_CHG) AS INV_MISC,")
        'S.AppendLine("SUM(GST_TAX) AS GST_TAX,")
        'S.AppendLine("SUM (INV_TOTAL_AMOUNT) AS INV_TOT_AMOUNT")
        'S.AppendLine("FROM " & SOTINVHT)
        'S.AppendLine("GROUP BY INV_NO_CONS, CUST_CODE, SALES_DIVISION_CODE , H_SALES_DIVISION_CODE")
        'Dim tblSOTINVHZ As DataTable = ASCDATA1.GetDataTable(S.ToString(), String.Empty, "V")
        'For Each rowSOTINVHZ As DataRow In tblSOTINVHZ.Rows
        '    Dim newSOTINVHZ As DataRow = dst.Tables("SOTINVHZ").NewRow
        '    newSOTINVHZ.Item("INV_TYPE") = "I"
        '    newSOTINVHZ.Item("INV_NO") = rowSOTINVHZ.Item("INV_NO_CONS").ToString
        '    newSOTINVHZ.Item("INV_DATE") = rowSOTINVHZ.Item("INV_DT").ToString
        '    newSOTINVHZ.Item("CUST_CODE") = rowSOTINVHZ.Item("CUST_CODE").ToString
        '    newSOTINVHZ.Item("SALES_DIVISION_CODE") = rowSOTINVHZ.Item("SALES_DIVISION_CODE").ToString
        '    newSOTINVHZ.Item("TOTAL_SALES") = rowSOTINVHZ.Item("TOT_SALES").ToString
        '    newSOTINVHZ.Item("H_SALES_DIVISION_CODE") = rowSOTINVHZ.Item("H_SALES_DIVISION_CODE").ToString
        '    newSOTINVHZ.Item("ORDR_CUST_PO") = rowSOTINVHZ.Item("ORDR_PO").ToString
        '    newSOTINVHZ.Item("INV_SALES") = Val(rowSOTINVHZ.Item("INV_SALE").ToString & "")
        '    newSOTINVHZ.Item("INV_FREIGHT") = Val(rowSOTINVHZ.Item("INV_FR").ToString & "")
        '    newSOTINVHZ.Item("INV_MISC_CHG") = Val(rowSOTINVHZ.Item("INV_MISC").ToString & "")
        '    newSOTINVHZ.Item("GST_TAX") = Val(rowSOTINVHZ.Item("GST_TAX").ToString & "")
        '    newSOTINVHZ.Item("INV_TOTAL_AMOUNT") = Val(rowSOTINVHZ.Item("INV_TOT_AMOUNT").ToString & "")
        '    dst.Tables.Item("SOTINVHZ").Rows.Add(newSOTINVHZ)
        'Next

        'ASCMAIN1.sql = "Select SOTINVH1.* " & vbCrLf _
        '        & ", 0 AS TOTAL_UNITS" & vbCrLf _
        '        & ", 0 AS TOTAL_UNITS_CANC" & vbCrLf _
        '        & ", 0 AS TOTAL_UNITS_BACK" & vbCrLf _
        '        & " from " & SOTINVH1 & " SOTINVH1, SOTORDR1" & vbCrLf _
        '        & " where SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
        '        & "   and SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
        '        & "   and SOTINVH1.INV_NO_CONS is Null"
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHX", 2))

        'ASCMAIN1.sql = "Select SOTINVH1.INV_NO_CONS, SOTINVH1.ORDR_BILL_TO_CUST as CUST_CODE" & vbCrLf _
        '        & ", Max(SOTINVH1.INV_DATE) AS INV_DATE" & vbCrLf _
        '        & ", Max(SOTINVH1.REASON_CODE) as REASON_CODE" & vbCrLf _
        '        & ", Max(SOTINVH1.SALES_DIVISION_CODE) as SALES_DIVISION_CODE" & vbCrLf _
        '        & ", Max(SOTINVH1.ORDR_CUST_PO) AS ORDR_CUST_PO" & vbCrLf _
        '        & ", Max(SOTINVH1.CUST_FACTOR_IND) AS CUST_FACTOR_IND" & vbCrLf _
        '        & ", Max(SOTINVH1.CUST_SURCHARGE_IND) AS CUST_SURCHARGE_IND" & vbCrLf _
        '        & ", Sum(SOTINVH1.INV_SALES) AS INV_SALES" & vbCrLf _
        '        & ", Sum(SOTINVH1.INV_FREIGHT) AS INV_FREIGHT" & vbCrLf _
        '        & ", Sum(SOTINVH1.INV_MISC_CHG) AS INV_MISC_CHG" & vbCrLf _
        '        & ", Sum(SOTINVH1.GST_TAX) AS GST_TAX" & vbCrLf _
        '        & ", Sum(SOTINVH1.INV_TOTAL_AMOUNT) AS INV_TOTAL_AMOUNT" & vbCrLf _
        '        & ", 0 AS TOTAL_UNITS" & vbCrLf _
        '        & ", 0 AS TOTAL_UNITS_CANC" & vbCrLf _
        '        & ", 0 AS TOTAL_UNITS_BACK" & vbCrLf _
        '        & " from " & SOTINVH1 & " SOTINVH1" & vbCrLf _
        '        & " where SOTINVH1.INV_NO_CONS is Not Null" & vbCrLf _
        '        & " group by SOTINVH1.INV_NO_CONS, SOTINVH1.ORDR_BILL_TO_CUST" & vbCrLf
        'For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
        '    Dim rowSOTINVHX As DataRow = dst.Tables("SOTINVHX").NewRow
        '    With rowSOTINVHX
        '        For Each DCOL As DataColumn In row.Table.Columns
        '            If DCOL.ColumnName = "INV_NO_CONS" Then
        '                .Item("INV_TYPE") = "I"
        '                .Item("INV_NO") = row.Item("INV_NO_CONS")
        '            Else
        '                .Item(DCOL.ColumnName) = row.Item(DCOL.ColumnName)
        '            End If
        '        Next
        '        .Item("CURR_CODE") = "USD"
        '        .Item("CURR_EXCH_RATE") = 1
        '    End With
        '    dst.Tables("SOTINVHX").Rows.Add(rowSOTINVHX)
        'Next

        ASCMAIN1.Progress("Building Work File - SOTINVHG")
        For Each TABLE_NAME As String In New String() {"SOTINVHG1", "SOTINVHG", "SOTINVHG2"}
            With dst.Tables.Add(TABLE_NAME)
                .Columns.Add("SD")
                .Columns.Add("CC")
                If TABLE_NAME <> "SOWINVHG2" Then .Columns.Add("ID", GetType(System.DateTime))
                .Columns.Add("QC", GetType(System.Int64))
                .Columns.Add("AC", GetType(System.Decimal))
                If TABLE_NAME = "SOTINVHG" Then .PrimaryKey = New DataColumn() { .Columns("SD"), .Columns("CC"), .Columns("ID")}
                If TABLE_NAME = "SOTINVHG2" Then .PrimaryKey = New DataColumn() { .Columns("SD"), .Columns("CC")}
            End With
        Next

        For Each row As DataRow In dst.Tables("SOTINVHN").Select("")
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim INV_DATE As Date = row.Item("INV_DATE")
            dst.Tables("SOTINVHG1").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE, row.Item("QTY_CANC"), row.Item("AMT_CANC")})
            If dst.Tables("SOTINVHG").Rows.Find(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE}) Is Nothing Then
                dst.Tables("SOTINVHG").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE, INV_DATE})
            End If
            If dst.Tables("SOTINVHG2").Rows.Find(New Object() {SALES_DIVISION_CODE, CUST_CODE}) Is Nothing Then
                dst.Tables("SOTINVHG2").Rows.Add(New Object() {SALES_DIVISION_CODE, CUST_CODE})
            End If
        Next

        Create_Relation("SOTINVHG", "SOTINVHG1", "SD,CC,ID")
        dst.Tables("SOTINVHG").Columns("QC").Expression = "SUM(CHILD.QC)"
        dst.Tables("SOTINVHG").Columns("AC").Expression = "SUM(CHILD.AC)"

        Create_Relation("SOTINVHG2", "SOTINVHG", "SD,CC")
        dst.Tables("SOTINVHG2").Columns("QC").Expression = "SUM(CHILD.QC)"
        dst.Tables("SOTINVHG2").Columns("AC").Expression = "SUM(CHILD.AC)"

        ' Master Files
        ASCMAIN1.Progress("Building Work File - Master Files")
        ASCMAIN1.sql = "Select ARTREAS1.* from ARTREAS1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTREAS1", 1))

        ASCMAIN1.sql = "Select SOTSDIV1.* from SOTSDIV1"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSDIV1", 1))

        'Performance.Add("000006", DateDiff(DateInterval.Second, PerformanceStart, Now()))

        ASCMAIN1.Progress("Calculating Last Recd Dates")
        For Each rowSOTINVHD As DataRow In dst.Tables("SOTINVHD").Select()
            Dim TARIFF_IND As String = ""
            If (rowSOTINVHD.Item("TARIFF_FLAG").ToString & String.Empty).Length = 10 Then
                Dim TARIFF_PRE = (rowSOTINVHD.Item("TARIFF_FLAG").ToString & String.Empty).Substring(8, 2)
                If TARIFF_PRE <> "00" Then
                    TARIFF_IND = String.Format("T{0}", TARIFF_PRE)
                End If
            End If
            rowSOTINVHD.Item("TARIFF_IND") = TARIFF_IND

            Dim STYLE_CODE As String = rowSOTINVHD.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowSOTINVHD.Item("COLOR_CODE").ToString & String.Empty
            S.Length = 0
            S.AppendLine("SELECT NVL(TO_CHAR(MAX(POTSHIP2.PO_DATE_RECEIVED),'MM/DD/YY'),'') PO_DATE_RECEIVED")
            S.AppendLine("FROM POTORDR2, POTSHIP3, POTSHIP2")
            S.AppendLine("WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
            S.AppendLine("AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO")
            S.AppendLine("AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO")
            S.AppendLine("AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO")
            S.AppendLine(String.Format("AND POTORDR2.STYLE_CODE = '{0}'", STYLE_CODE))
            S.AppendLine(String.Format("AND POTORDR2.COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = S.ToString()
            Dim LAST_RCD_DATE As String = ASCDATA1.GetDataValue
            'If STYLE_CODE = "VCO72239A" Then Stop
            If IsDate(LAST_RCD_DATE) Then
                LAST_RCD_DATE = Format(CDate(LAST_RCD_DATE), "MM/dd/yy")
            Else
                S.Length = 0
                S.AppendLine("SELECT SUM(NVL(WHSE_QTY_TRAN,0)) AS IN_TRAN")
                S.AppendLine("FROM ICTSTAT2")
                S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                S.AppendLine(String.Format("AND WHSE_CODE = '{0}'", "NJE"))
                ASCMAIN1.sql = S.ToString()
                Dim IN_TRAN As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                If IN_TRAN > 0 Then
                    LAST_RCD_DATE = "In-Tran"
                Else
                    S.Length = 0
                    S.AppendLine("SELECT SUM(NVL(WHSE_QTY_ON_ORDER,0)) AS IN_WIP")
                    S.AppendLine("FROM ICTSTAT2")
                    S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                    S.AppendLine(String.Format("AND WHSE_CODE = '{0}'", "NJE"))
                    ASCMAIN1.sql = S.ToString()
                    Dim IN_WIP As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                    If IN_WIP > 0 Then
                        LAST_RCD_DATE = "In-WIP"
                    Else
                        LAST_RCD_DATE = ""
                    End If
                End If
            End If
            'row.Item("LAST_RCD_DATE") = LAST_RCD_DATE
            rowSOTINVHD.Item("LAST_RCD") = LAST_RCD_DATE
        Next


        Return ""
    End Function

    Private Sub UltraDateTimeEditor4_ValueChanged(sender As Object, e As EventArgs) Handles UltraDateTimeEditor4.ValueChanged
        Absx1.dteFor("DTE1").Value = Absx1.dteFor("DTE0").Value
    End Sub

    Private Sub UltraDateTimeEditor3_ValueChanged(sender As Object, e As EventArgs) Handles UltraDateTimeEditor3.ValueChanged
        Absx1.dteFor("DTE0").Value = Absx1.dteFor("DTE1").Value
    End Sub
End Class