Imports System.Text

Public Class SORSHIP6
    Dim BPeriods As New List(Of String)
    Dim EPeriods As New List(Of String)
    Dim S As StringBuilder = New StringBuilder With {.Length = 0}
    Enum DayType
        First
        Last
    End Enum

#Region "Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        RWU = "N"

        'Range_Events(grpSO_DATE_INV1)

        Set_cmbYP("BYP0", ASCMAIN1.CYP, -60, 0, -12)
        Set_cmbYP("EYP0", ASCMAIN1.CYP, -60, 0, -12)
        Set_cmbYP_Child("BYP1", 60, "BYP0", 1)
        Set_cmbYP_Child("EYP1", 60, "EYP0", 1)
        Set_cmbYP_Child("BYP2", 60, "BYP1", 1)
        Set_cmbYP_Child("EYP2", 60, "EYP1", 1)
        Set_cmbYP_Child("BYP3", 60, "BYP2", 1)
        Set_cmbYP_Child("EYP3", 60, "EYP2", 1)
        Set_cmbYP_Child("BYP4", 60, "BYP3", 1)
        Set_cmbYP_Child("EYP4", 60, "EYP3", 1)
        Set_cmbYP_Child("BYP5", 60, "BYP4", 1)
        Set_cmbYP_Child("EYP5", 60, "EYP4", 1)
        Set_cmbYP_Child("BYP6", 60, "BYP5", 1)
        Set_cmbYP_Child("EYP6", 60, "EYP5", 1)
        Set_cmbYP_Child("BYP7", 60, "BYP6", 1)
        Set_cmbYP_Child("EYP7", 60, "EYP6", 1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim maxPeriod As Integer = getMaxPeriods()

        ' Prepare filters from Run-Time Options

        SUBT = ""
        If Absx1.optFor("OPTASN").Value = "S" Then
            SUBT &= "Stock Styles Only"
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            SUBT &= "Non-Stock Styles Only"
        End If

        Dim sql_filter As String = ""

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")
        Dim sql_TABLE_NAMEs_orig As String = sql_TABLE_NAMEs
        Dim sql_JOIN_orig As String = sql_JOIN

        Dim sql_filter2 As String = ""

        If Absx1.optFor("OPTASN").Value = "S" Then
            sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Null"
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Not Null"
        End If

        S.Length = 0
        S.AppendLine(String.Format("SELECT {0}", sql_SELECT_cols))
        S.AppendLine(", SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
        S.AppendLine("SUM(NVL(S2.WHSE_QTY_ON_ORDER,0)) AS WHSE_QTY_ON_ORDER,")
        S.AppendLine("SUM(NVL(S2.WHSE_QTY_TRAN,0)) AS WHSE_QTY_TRAN,")
        S.AppendLine("SUM(NVL(S2.WHSE_QTY_OPEN,0)) AS WHSE_QTY_OPEN,")
        S.AppendLine("SUM(NVL(S2.WHSE_QTY_PICK,0)) AS WHSE_QTY_PICK")
        S.AppendLine(String.Format("FROM ICTSTYL1, ICTSTYC1 C1, ICTSTAT2 S2 {0}", ""))
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = C1.STYLE_CODE")
        S.AppendLine("AND C1.STYLE_CODE = S2.STYLE_CODE (+)")
        S.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
        S.AppendLine(sql_WHERE)
        S.AppendLine(sql_filter2)
        S.AppendLine(String.Format("GROUP BY {0}", sql_GROUP_BY_cols))
        ASCMAIN1.sql = S.ToString

        S.Length = 0
        S.AppendLine("Insert into " & ASTSRPT1)
        S.AppendLine(" ")
        S.AppendLine(ASCMAIN1.sql)
        S.AppendLine(" ")
        ASCDATA1.ExecuteSQL(S.ToString)

        S.Length = 0
        S.AppendLine(String.Format("SELECT {0}", sql_SELECT_cols))
        S.AppendLine(", SUM(SOTINVH2.ORDR_QTY_SHIP) AS QTY_SHP")
        S.AppendLine(", SUM(SOTINVH2.ORDR_UNIT_PRICE * SOTINVH2.ORDR_QTY_SHIP) AS DOL_SHIP")
        S.AppendLine("FROM ICTSTYL1, SOTINVH2")
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE")
        S.AppendLine("AND SOTINVH2.ORDR_YYYYPP_UPDATED >= :PARM1")
        S.AppendLine("AND SOTINVH2.ORDR_YYYYPP_UPDATED <= :PARM2")
        S.AppendLine(sql_WHERE)
        S.AppendLine(sql_filter2)
        S.AppendLine(String.Format("GROUP BY {0}", sql_GROUP_BY_cols))
        'S.AppendLine("SELECT")
        'S.AppendLine("I2.STYLE_CODE,")
        'S.AppendLine("I2.COLOR_CODE,")
        'S.AppendLine("SUM(I2.ORDR_QTY_SHIP) AS QTY_SHP,")
        'S.AppendLine("SUM(I2.ORDR_UNIT_PRICE * I2.ORDR_QTY_SHIP) AS DOL_SHIP")
        'S.AppendLine("FROM SOTINVH2 I2")
        'S.AppendLine("WHERE I2.ORDR_YYYYPP_UPDATED >= :PARM1")
        'S.AppendLine("AND I2.ORDR_YYYYPP_UPDATED <= :PARM2")
        'S.AppendLine("GROUP BY")
        'S.AppendLine("I2.STYLE_CODE,")
        'S.AppendLine("I2.COLOR_CODE")
        ASCMAIN1.sql = S.ToString
        Create_TDA(dst.Tables.Add("SOTINVHX"), "SOTINVH2", "**",, False, "VV")

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        For i As Integer = 1 To 8
            dst.Tables.Item("ASTSRPT1").Columns.Add(String.Format("SP_TY_0{0}", i), GetType(System.Double))
            dst.Tables.Item("ASTSRPT1").Columns.Add(String.Format("SP_LY_0{0}", i), GetType(System.Double))
        Next

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
                rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC " & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        Dim YearPass As Integer = 1
        If Absx1.chkFor("CHKPRIORYEAR").Checked Then
            YearPass = 2
        End If
        For Pass As Integer = 1 To YearPass
            For p As Integer = 1 To 8
                Dim BP As String = ""
                Dim EP As String = ""
                BP = getPeriodFromDescription(Absx1.cmbFor(String.Format("BYP{0}", p - 1), True).Value)
                EP = getPeriodFromDescription(Absx1.cmbFor(String.Format("EYP{0}", p - 1), True).Value)
                If Pass = 2 Then
                    BP = String.Format("{0}{1}{2}", BP.Substring(0, 3), Val(BP.Substring(3, 1)) - 1, BP.Substring(4, 2))
                    EP = String.Format("{0}{1}{2}", EP.Substring(0, 3), Val(EP.Substring(3, 1)) - 1, EP.Substring(4, 2))
                End If

                ASCMAIN1.Progress(String.Format("Fetching Data for {0} - {1}", bp, ep), "")
                Fill_Records("SOTINVHX", New String() {bp, ep})
                For Each rowSOTINVHX As DataRow In dst.Tables("SOTINVHX").Select()
                    Dim filter As String = CalculateFilter(rowSOTINVHX)
                    Dim SP As Double = Val(rowSOTINVHX.Item("QTY_SHP"))
                    ASCMAIN1.Progress(String.Format("Processing {0} - {1}", bp, ep))
                    Dim rowASTSRPT1 As DataRow = dst.Tables.Item("ASTSRPT1").Select(filter).FirstOrDefault
                    If Not IsNothing(rowASTSRPT1) Then
                        If Pass = 1 Then
                            rowASTSRPT1.Item(String.Format("SP_TY_0{0}", p)) = SP
                        Else
                            rowASTSRPT1.Item(String.Format("SP_LY_0{0}", p)) = SP
                        End If

                    End If
                Next
            Next
        Next

        'S.Length = 0
        '    S.AppendLine("Select SOTINVH2.*")
        'For i As Integer = 1 To 8
        '    S.AppendLine(", 0 as ORDR_QTY_SHIP" & i)
        '    S.AppendLine(", 0 as ORDR_QTY_SHIP_LY" & i)
        'Next
        'S.AppendLine("FROM SOTINVH2, ICTSTYL1 ")
        'S.AppendLine("WHERE SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        'S.AppendLine("AND (SOTINVH2.INV_NO, SOTINVH2.INV_LNO) in (Select Distinct INV_NO,INV_LNO FROM " & ASTSRPT1 & ")")
        'ASCMAIN1.sql = S.ToString
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTINVHX", 3))

    End Sub

    Private Function CalculateFilter(ByRef rowSOTINVHX As DataRow) As String
        Dim RetVal As String = ""
        For i As Integer = 0 To 7
            Dim ColName As String = dst.Tables.Item("SOTINVHX").Columns(i).ColumnName
            Dim GCol As String = String.Format("G{0}", i + 1)
            Dim ColVal As String = rowSOTINVHX.Item(i) & ""
            If ColName.Length > 2 Then
                Dim AddAnd As String = ""
                Dim DataAppend As String = ""
                If RetVal.Length > 0 Then
                    AddAnd = " AND "
                End If
                Select Case ColName
                    Case "CUST_CODE"
                        DataAppend = "Customer:"
                    Case "SALES_DIVISION_CODE"
                        DataAppend = "Division:"
                End Select
                RetVal = RetVal & String.Format("{0}{1} = '{2}{3}'", AddAnd, GCol, DataAppend, ColVal)
            Else
                Exit For
            End If
        Next
        Return RetVal
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        If parms.Length > 0 Then
        End If

        EnforceConstraints(False)
        'Fill_Records("ASTSRPT1")
        EnforceConstraints(True)
    End Sub

    Overrides Function Prepare_dst(ByVal perform_fill As Boolean, ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1
    End Function

    Public Overrides Sub Print_Report()
        Dim maxPeriods As Integer = getMaxPeriods()
        Dim dateTitles As List(Of String) = buildDateTitles(maxPeriods)
        CleanZeros()

        'CalculatePeriods()

        'If Absx1.chkFor("CHKPRIORYEAR").Checked Then
        '    RPT = "SORSHIP7"
        'Else
        '    RPT = "SORSHIP6"
        'End If
        RPT = "SORSHIP6"

        For i As Integer = 1 To 8
            CR_params.Add("GroupTitle" & i, dateTitles(i - 1))
            If i > 1 Then
                If Absx1.chkFor("INVGROUP" & i).Checked Then
                    CR_params.Add("GroupShow" & i, "1")
                Else
                    CR_params.Add("GroupShow" & i, "0")
                End If
            End If
        Next

        CR_params.Add("SUBT", txtDescription.Text & SUBT)

        Generate_Report(RPT, , SUBT)
    End Sub

    Private Sub CleanZeros()

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select()
            Dim FoundData As Boolean = False
            Dim Cols As String() = {"PMTPRP11", "PMTPRP12", "PMTPRP13"}
            For Each Col As String In Cols

            Next
            rowASTSRPT1.Item("ASTSRPT1") = "XXXXX"
        Next
    End Sub

    Private Sub UpdateASTSRPT1(ByVal PO_ORDER_NO As String, PO_ORDER_LNO As Integer, Group As Integer)
        Dim filter As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' AND PO_ORDER_LNO = " & PO_ORDER_LNO
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select(filter)
            rowASTSRPT1.Item("SHIP_OPN" & Group) = Val(rowASTSRPT1.Item("SHIP_OPN" & Group) & "") + Val(rowASTSRPT1.Item("SHIP_OPN") & "")
            rowASTSRPT1.Item("SHIP_OPN_LY" & Group) = Val(rowASTSRPT1.Item("SHIP_OPN_LY" & Group) & "") + Val(rowASTSRPT1.Item("SHIP_OPN_LY") & "")
        Next
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length > 4 Then
                EMsg &= vbCr & "Maximum number of Sort Fields for this report is 4"
            End If
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub CalculatePeriods()
        Dim maxPeriod As Integer = getMaxPeriods()
        'buildDates(maxPeriod)
        For Each rowSOTINVHX As DataRow In dst.Tables("SOTINVHX").Select()

            Dim ORDR_YYYYPP_UPDATED As String = rowSOTINVHX.Item("ORDR_YYYYPP_UPDATED")
            For i As Integer = 1 To maxPeriod
                If (ORDR_YYYYPP_UPDATED >= BPeriods(i - 1) And ORDR_YYYYPP_UPDATED <= EPeriods(i - 1)) Then
                    rowSOTINVHX.Table.Columns.Item("SHIP_OPN" & i).ReadOnly = False
                    rowSOTINVHX.Item("SHIP_OPN" & i) = Val(rowSOTINVHX.Item("SHIP_OPN" & i) & "") + Val(rowSOTINVHX.Item("ORDR_QTY_SHIP") & "")
                End If
            Next
            If Absx1.chkFor("CHKPRIORYEAR").Checked Then
                For i As Integer = 1 To maxPeriod
                    Dim BPeriodL As String = ASCMAIN1.Period_Calc(BPeriods(i - 1), -12)
                    Dim EPeriodL As String = ASCMAIN1.Period_Calc(EPeriods(i - 1), -12)
                    If (ORDR_YYYYPP_UPDATED >= BPeriodL And ORDR_YYYYPP_UPDATED <= EPeriodL) Then
                        rowSOTINVHX.Table.Columns.Item("SHIP_OPN_LY" & i).ReadOnly = False
                        rowSOTINVHX.Item("SHIP_OPN_LY" & i) = Val(rowSOTINVHX.Item("SHIP_OPN_LY" & i) & "") + Val(rowSOTINVHX.Item("ORDR_QTY_SHIP") & "")
                    End If
                Next
            End If
        Next

    End Sub

    'Private Sub buildDates(maxPeriod As Integer, Optional ByVal useVerbose As Boolean = False)
    '    Periods.Clear()
    '    If Absx1.chkFor("CHKORDR_YYYYPP_UPDATED1").Checked Then
    '        Periods.Add("200001")
    '    Else
    '        BDate.Add(CDate(Format(Absx1.dteFor("ORDR_YYYYPP_UPDATED1").Value, "dd-MMM-yyyy")))
    '    End If
    '    EDate.Add(CDate(Format(Absx1.dteFor("PO_DATE_ETA_L1").Value, "dd-MMM-yyyy")))

    '    For i As Integer = 2 To maxPeriod
    '        BDate.Add(CDate(Format(Absx1.dteFor("ORDR_YYYYPP_UPDATED" & i).Value, "dd-MMM-yyyy")))
    '        If Absx1.chkFor("CHKPO_DATE_ETA_L" & i).Checked Then
    '            EDate.Add(DateSerial(2100, 12, 31))
    '        Else
    '            EDate.Add(CDate(Format(Absx1.dteFor("PO_DATE_ETA_L" & i).Value, "dd-MMM-yyyy")))
    '        End If
    '    Next
    'End Sub

    Private Function buildDateTitles(maxPeriod As Integer) As List(Of String)
        Dim retVal As New List(Of String)
        Dim date1 As String = ""
        Dim date2 As String = ""
        BPeriods.Clear()
        EPeriods.Clear()
        If Absx1.chkFor("CHKINV_DATE_F1").Checked Then
            date1 = "First"
            BPeriods.Add("200001")
        Else
            date1 = Format(CDate(getDateFromPeriod(getPeriodFromDescription(Absx1.cmbFor("BYP0", True).Value), DayType.First)), "MM/dd/yy")
            BPeriods.Add(getPeriodFromDescription(Absx1.cmbFor("BYP0", True).Value))
        End If
        date2 = Format(CDate(getDateFromPeriod(getPeriodFromDescription(Absx1.cmbFor("EYP0", True).Value), DayType.Last)), "MM/dd/yy")
        retVal.Add(date1 & " to " & date2)
        EPeriods.Add(getPeriodFromDescription(Absx1.cmbFor("EYP0", True).Value))
        For i As Integer = 1 To maxPeriod - 1
            date1 = Format(CDate(getDateFromPeriod(getPeriodFromDescription(Absx1.cmbFor("BYP" & i, True).Value), DayType.First)), "MM/dd/yy")
            BPeriods.Add(getPeriodFromDescription(Absx1.cmbFor("BYP" & i, True).Value))
            If Absx1.chkFor("CHKINV_DATE_L" & i + 1).Checked Then
                date2 = "Last"
                EPeriods.Add("202501")
            Else
                date2 = Format(CDate(getDateFromPeriod(getPeriodFromDescription(Absx1.cmbFor("EYP" & i, True).Value), DayType.Last)), "MM/dd/yy")
                EPeriods.Add(getPeriodFromDescription(Absx1.cmbFor("EYP" & i, True).Value))
            End If
            retVal.Add(date1 & " to " & date2)
        Next
        For i As Integer = maxPeriod + 1 To 8
            retVal.Add("N/A")
        Next
        Return retVal
    End Function

    Function Get_Dates(ByVal Period As Integer, Optional ByVal PriorYear As Boolean = False) As String
        Dim sql As String = ""

        Dim COL_NAMEs() As String
        Dim CONTROL_NAMEs() As String
        COL_NAMEs = New String() {"ORDR_YYYYPP_UPDATED"}
        CONTROL_NAMEs = New String() {"CHKINV_DATE"}

        Dim ctlIndex As Integer = 0
        For Each COLUMN_NAME As String In COL_NAMEs
            Dim CONTROL_NAME As String = CONTROL_NAMEs(ctlIndex)
            Dim BPeriod As String = ""
            If Absx1.chkFor("CHKINV_DATE_F1").Checked And Period = 1 Then
                BPeriod = "200001"
            Else
                BPeriod = getPeriodFromDescription(Absx1.cmbFor("BYP" & Period - 1, True).Value)
                If PriorYear Then
                    BPeriod = ASCMAIN1.Period_Calc(getPeriodFromDescription(Absx1.cmbFor("BYP" & Period - 1, True).Value), -12)
                End If
            End If
            Dim LPeriod As String = ""
            If Period > 1 Then
                If Absx1.chkFor("CHKINV_DATE_L" & Period).Checked And Period > 1 Then
                    LPeriod = "202501"
                Else
                    LPeriod = getPeriodFromDescription(Absx1.cmbFor("EYP" & Period - 1, True).Value)
                    If PriorYear Then
                        LPeriod = ASCMAIN1.Period_Calc(getPeriodFromDescription(Absx1.cmbFor("EYP" & Period - 1, True).Value), -12)
                    End If
                End If
            Else
                LPeriod = getPeriodFromDescription(Absx1.cmbFor("EYP" & Period - 1, True).Value)
                If PriorYear Then
                    LPeriod = ASCMAIN1.Period_Calc(getPeriodFromDescription(Absx1.cmbFor("EYP" & Period - 1, True).Value), -12)
                End If
            End If
            sql = sql & " and A." & COLUMN_NAME & " >= '" & BPeriod & "'"
            sql = sql & " and A." & COLUMN_NAME & " <= '" & LPeriod & "'"
            ctlIndex += 1
        Next
        sql = Replace(sql, "A.ORDR_YYYYPP_UPDATED", "SOTINVH1.ORDR_YYYYPP_UPDATED")
        sql = Replace(sql, "A.", "SOTINVH1.")

        Return sql
    End Function

    Private Function getPeriodFromDescription(ByVal PDESC As String) As String
        Dim RetVal As String = ""
        If PDESC.Length >= 7 Then
            RetVal = PDESC.Substring(0, 4) & PDESC.Substring(5, 2)
        End If
        Return RetVal
    End Function

    Private Function getDateFromPeriod(ByVal YYYYMM As String, ByVal dayType As DayType) As Date
        Dim RetVal As Date = DateSerial(YYYYMM.Substring(0, 4), YYYYMM.Substring(4, 2), 1)
        If dayType = DayType.Last Then
            RetVal = RetVal.AddMonths(1).AddDays(-1)
        End If
        Return RetVal
    End Function

    Private Function getMaxPeriods() As Integer
        Dim retval As Integer = 0
        For i As Integer = 2 To 8
            If Absx1.chkFor("INVGROUP" & i).Visible Then
                If Absx1.chkFor("INVGROUP" & i).Checked Then
                    retval = i
                Else
                    retval = i - 1
                End If
            End If
        Next
        Return retval
    End Function

    Private Sub setINVValues(ByVal Group As Integer)
        If Group < 8 Then
            For gp As Integer = Group + 1 To 8
                If gp >= Group + 1 Then
                    If gp = Group + 1 Then
                        If Absx1.chkFor("INVGROUP" & Group).Visible Then
                            Absx1.chkFor("INVGROUP" & gp).Visible = Absx1.chkFor("INVGROUP" & Group).Checked
                            Absx1.chkFor("INVGROUP" & gp).Checked = False
                        Else
                            Absx1.chkFor("INVGROUP" & gp).Visible = False
                        End If
                    Else
                        Absx1.chkFor("INVGROUP" & gp).Visible = False
                    End If

                    Absx1.CtlFor("GRP_DATE_INV" & gp).Visible = False
                End If
            Next
        End If
        Absx1.CtlFor("GRP_DATE_INV" & Group).Visible = Absx1.chkFor("INVGROUP" & Group).Checked
    End Sub

#End Region

#Region "Form Controls"
#Region "Check Boxes"
    Private Sub chkINVGROUP2_CheckedChanged(sender As Object, e As EventArgs) Handles chkINVGROUP2.CheckedChanged
        setINVValues(2)
    End Sub

    Private Sub chkINVGROUP3_CheckedChanged(sender As Object, e As EventArgs) Handles chkINVGROUP3.CheckedChanged
        setINVValues(3)
    End Sub

    Private Sub chkINVGROUP4_CheckedChanged(sender As Object, e As EventArgs) Handles chkINVGROUP4.CheckedChanged
        setINVValues(4)
    End Sub

    Private Sub chkINVGROUP5_CheckedChanged(sender As Object, e As EventArgs) Handles chkINVGROUP5.CheckedChanged
        setINVValues(5)
    End Sub

    Private Sub chkINVGROUP6_CheckedChanged(sender As Object, e As EventArgs) Handles chkINVGROUP6.CheckedChanged
        setINVValues(6)
    End Sub

    Private Sub chkINVGROUP7_CheckedChanged(sender As Object, e As EventArgs) Handles chkINVGROUP7.CheckedChanged
        setINVValues(7)
    End Sub

    Private Sub chkINVGROUP8_CheckedChanged(sender As Object, e As EventArgs) Handles chkINVGROUP8.CheckedChanged
        setINVValues(8)
    End Sub
#End Region
#End Region
End Class