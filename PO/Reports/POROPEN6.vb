Imports System.Text

Public Class POROPEN6
    Dim BDate As New List(Of Date)
    Dim EDate As New List(Of Date)
    'NOTICE: This report was basterized from POROPEN4.
    'Take care where the original report focused on PO_QTY_OPN, this report is concerned with PO_QTY_REC.
    'You will find instances where the report field PO_QTY_OPN is acually being filled with PO_QTY_REC data.
#Region "Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        RWU = "N"

        Range_Events(grpPO_DATE_ETA1)

        Absx1.dteFor("PO_DATE_ETA_F1").DateTime = DateSerial(Now.Year, 1, 1)
        Absx1.dteFor("PO_DATE_ETA_L1").DateTime = DateSerial(Now.Year, 1, 31)

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim maxPeriod As Integer = getMaxPeriods()

        ' Prepare filters from Run-Time Options

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Dim STOCK_SUB As String = ""
            Select Case Absx1.optFor("OPTASN").Value
                Case "S"
                    STOCK_SUB = "Stock Only"
                Case "N"
                    STOCK_SUB = "Non-Stock Only"
                Case Else
                    STOCK_SUB = "All Stock"
            End Select
            If SUBT.Length = 0 Then
                SUBT = STOCK_SUB
            Else
                SUBT = SUBT & ", " & STOCK_SUB
            End If
        Else
            SUBT = ""
            If Absx1.optFor("OPTASN").Value = "S" Then
                SUBT &= "Stock Styles Only"
            ElseIf Absx1.optFor("OPTASN").Value = "N" Then
                SUBT &= "Non-Stock Styles Only"
            End If
        End If

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        'dst.Tables.Item("ASTSRPT1").Columns.Add("WHSE_QTY_ON_HAND", GetType(System.Double))

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")
        Dim sql_TABLE_NAMEs_orig As String = sql_TABLE_NAMEs
        Dim sql_JOIN_orig As String = sql_JOIN
        Dim YearPass As Integer = 1
        If Absx1.chkFor("CHKPRIORYEAR").Checked Then
            YearPass = 2
        End If
        For Pass As Integer = 1 To YearPass
            For Period As Integer = 1 To maxPeriod
                Dim sql_filter2 As String = ""
                If Pass = 1 Then
                    sql_filter2 = Get_Dates(Period)
                Else
                    sql_filter2 = Get_Dates(Period, True)
                End If

                If Absx1.optFor("OPTASN").Value = "S" Then
                    sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Null"
                ElseIf Absx1.optFor("OPTASN").Value = "N" Then
                    sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Not Null"
                End If
                sql_filter2 &= "" & " and POTSHIP2.PO_SHIP_STATUS = 'C'" & vbCrLf & " and NVL(POTSHIP3.PO_QTY_REC,0) > 0" & vbCrLf

                sql_TABLE_NAMEs = sql_TABLE_NAMEs_orig & ",POTSHIP2,POTSHIP3"
                sql_JOIN = sql_JOIN_orig & " " & " and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf & " and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf & " and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf & " and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf & " and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf
                If Not sql_TABLE_NAMEs.Contains("POTSHIP1") Then
                    sql_TABLE_NAMEs = sql_TABLE_NAMEs & ",POTSHIP1"
                End If

                ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf & ", POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf & ", 'OPENPO' PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf & ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_SHP, POTORDR2.PO_QTY_REC, POTORDR2.PO_QTY_OPN, 0 SHIP_QTY, 0 SHIP_OPN, 0 SHIP_REC" & ", 0 WHSE_QTY_ON_HAND from POTORDR2" & sql_TABLE_NAMEs & vbCrLf & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf

                ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & G1thru9 & ",PO_ORDER_NO,PO_ORDER_LNO,PO_SHIPMENT_NO,PO_SHIPMENT_LNO" & ",PO_QTY_ORD,PO_QTY_SHP,PO_QTY_REC,PO_QTY_OPN,SHIP_QTY,SHIP_OPN,SHIP_REC, WHSE_QTY_ON_HAND" & ") " & " (" & ASCMAIN1.sql & ")")
            Next
        Next

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        For i As Integer = 1 To 12
            dst.Tables.Item("ASTSRPT1").Columns.Add("SHIP_OPN" & Format(i, "0"), GetType(System.Double))
            'dst.Tables.Item("ASTSRPT1").Columns.Add("SHIP_OPN_LY" & Format(i, "0"), GetType(System.Double))
        Next

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC " & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        Dim SQL As StringBuilder = New StringBuilder With {.Length = 0}
        SQL.AppendLine("Select POTORDR2.*,ICTSTYL1.CASE_CUBE, ICTCOLR1.COLOR_DESC")
        SQL.AppendLine(", POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTORDR1.PO_REFERENCE, POTORDR1.WHSE_CODE")
        SQL.AppendLine(", POTORDR1.PO_SPEC_ORDR_NO, POTORDR1.PO_DATE_ORDERED, POTORDR1.PO_DATE_CANCEL,POTORDR1.PORT_CODE_ORIG,POTORDR1.PO_SHIP_VIA")
        For i As Integer = 1 To 12
            SQL.AppendLine(", 0 as PO_QTY_OPN" & i)
            SQL.AppendLine(", 0 as PO_QTY_OPN_LY" & i)
        Next
        'SQL.AppendLine(", 0 as PO_QTY_OPN1, 0 as PO_QTY_OPN2, 0 as PO_QTY_OPN3, 0 as PO_QTY_OPN4, 0 as PO_QTY_OPN5")
        'SQL.AppendLine(", " & IIf(optSORT.Value & "" = "D", "TO_CHAR(POTORDR2.PO_DATE_SHIP_BY,'YYYYMMDD')", IIf(optSORT.Value & "" = "S", "POTORDR2.STYLE_CODE", "POTORDR2.PO_ORDER_NO")) & " SORT1")
        'SQL.AppendLine(", " & IIf(optSORT.Value & "" = "D", "POTORDR2.PO_ORDER_NO", IIf(optSORT.Value & "" = "S", "POTORDR2.COLOR_CODE", "TO_CHAR(POTORDR2.PO_ORDER_LNO,'000000')")) & " SORT2")
        SQL.AppendLine(" from POTORDR2,POTORDR1,ICTCOLR1, ICTSTYL1 ")
        SQL.AppendLine(" where (POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO) in (Select Distinct PO_ORDER_NO,PO_ORDER_LNO from " & ASTSRPT1 & ")")
        SQL.AppendLine("   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
        SQL.AppendLine("   and POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        SQL.AppendLine("   and ICTCOLR1.COLOR_CODE = POTORDR2.COLOR_CODE")

        'If Absx1.chkFor("CHKQTYOH").Checked Then
        '    SQL.AppendLine("UNION")
        '    SQL.AppendLine("SELECT")
        '    SQL.AppendLine("'000000' AS PO_ORDER_NO,")
        '    SQL.AppendLine("1 AS PO_ORDER_LNO,")
        '    SQL.AppendLine("NULL STYLE_CODE,")
        '    SQL.AppendLine("NULL COLOR_CODE,")
        '    SQL.AppendLine("0 PO_QTY_OR,")
        '    SQL.AppendLine("0 PO_QTY_SHP,")
        '    SQL.AppendLine("0 PO_QTY_REC,")
        '    SQL.AppendLine("0 PO_QTY_OPN,")
        '    SQL.AppendLine("0 PO_COST,")
        '    SQL.AppendLine("SYSDATE PO_DATE_SHIP_BY,")
        '    SQL.AppendLine("SYSDATE PO_DATE_ETA,")
        '    SQL.AppendLine("SYSDATE PO_ORIG_DATE_SHIP_BY,")
        '    SQL.AppendLine("SYSDATE PO_ORIG_DATE_ETA,")
        '    SQL.AppendLine("'O' PO_STATUS,")
        '    SQL.AppendLine("0 PO_QTY_UOM,")
        '    SQL.AppendLine("0 PO_COST_VCOST,")
        '    SQL.AppendLine("0 PO_COST_MATLS,")
        '    SQL.AppendLine("NULL CMT_NO,")
        '    SQL.AppendLine("NULL STYLE_NOTES,")
        '    SQL.AppendLine("0 YIELD_QTY,")
        '    SQL.AppendLine("0 SUB_UNIT_PACK_QTY,")
        '    SQL.AppendLine("0 PO_COST_VCOST_DZ,")
        '    SQL.AppendLine("0 PO_COST_MATLS_DZ,")
        '    SQL.AppendLine("0 PO_COST_OTHER,")
        '    SQL.AppendLine("0 PO_COST_COMM,")
        '    SQL.AppendLine("NULL LINE_FINISHED,")
        '    SQL.AppendLine("0 YARDS_CONSUMED,")
        '    SQL.AppendLine("0 FABRIC_COST,")
        '    SQL.AppendLine("NULL INIT_OPER,")
        '    SQL.AppendLine("NULL LAST_OPER,")
        '    SQL.AppendLine("NULL INIT_DATE,")
        '    SQL.AppendLine("SYSDATE LAST_DATE,")
        '    SQL.AppendLine("0 PO_COST_QUOTA,")
        '    SQL.AppendLine("NULL LAST_OPER_SHIP_BY,")
        '    SQL.AppendLine("SYSDATE LAST_DATE_SHIP_BY,")
        '    SQL.AppendLine("NULL SHIP_COST_CHANGE_USER,")
        '    SQL.AppendLine("SYSDATE SHIP_COST_CHANGE_DATE,")
        '    SQL.AppendLine("NULL DFQUOTA,")
        '    SQL.AppendLine("NULL PO_COST_ACCEPTED,")
        '    SQL.AppendLine("NULL PO_COST_ACCEPTED_USERID,")
        '    SQL.AppendLine("SYSDATE PO_COST_ACCEPTED_DATE,")
        '    SQL.AppendLine("0 PO_COST_BUFFER,")
        '    SQL.AppendLine("0 CARTON_PACK_QTY,")
        '    SQL.AppendLine("NULL ORDR_NO,")
        '    SQL.AppendLine("0 ORDR_LNO,")
        '    SQL.AppendLine("NULL PO_CONF_NO,")
        '    SQL.AppendLine("SYSDATE PO_CONF_DATE,")
        '    SQL.AppendLine("0 INNER_PACK_QTY,")
        '    SQL.AppendLine("NULL PO_LINE_NOTE_INT,")
        '    SQL.AppendLine("NULL CASE_CUBE,")
        '    SQL.AppendLine("NULL COLOR_DESC,")
        '    SQL.AppendLine("NULL VEND_CODE,")
        '    SQL.AppendLine("NULL FACTORY_CODE,")
        '    SQL.AppendLine("NULL PO_REFERENCE,")
        '    SQL.AppendLine("NULL WHSE_CODE,")
        '    SQL.AppendLine("NULL PO_SPEC_ORDR_NO,")
        '    SQL.AppendLine("SYSDATE PO_DATE_ORDERED,")
        '    SQL.AppendLine("SYSDATE PO_DATE_CANCEL,")
        '    SQL.AppendLine("NULL PORT_CODE_ORIG,")
        '    SQL.AppendLine("NULL PO_SHIP_VIA,")
        '    SQL.AppendLine("0 as PO_QTY_OPN1,")
        '    SQL.AppendLine("0 as PO_QTY_OPN2,")
        '    SQL.AppendLine("0 as PO_QTY_OPN3,")
        '    SQL.AppendLine("0 as PO_QTY_OPN4,")
        '    SQL.AppendLine("0 as PO_QTY_OPN5,")
        '    SQL.AppendLine("0 as PO_QTY_OPN6,")
        '    SQL.AppendLine("0 as PO_QTY_OPN7,")
        '    SQL.AppendLine("0 as PO_QTY_OPN8,")
        '    SQL.AppendLine("0 as PO_QTY_OPN9,")
        '    SQL.AppendLine("0 AS PO_QTY_OPN10,")
        '    SQL.AppendLine("0 AS PO_QTY_OPN11,")
        '    SQL.AppendLine("0 AS PO_QTY_OPN12")
        '    SQL.AppendLine("FROM DUAL")
        'End If

        ASCMAIN1.sql = SQL.ToString
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDRX", 2))

        ASCMAIN1.sql = "Select POTSHIP3.*" & vbCrLf & ", POTSHIP2.CONTAINER_NO, POTSHIP2.PO_SHIP_STATUS SHIP_STATUS" & vbCrLf & ", POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_REF_NO" & vbCrLf & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf & " from POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf & " where (POTSHIP3.PO_SHIPMENT_NO,POTSHIP3.PO_SHIPMENT_LNO) in (Select Distinct PO_SHIPMENT_NO,PO_SHIPMENT_LNO from " & ASTSRPT1 & ")" & vbCrLf & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTSHIPX", 4))
    End Sub

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
        'Dim isLongReport As Boolean = False
        CalculatePeriods()

        'If maxPeriods <= 8 Then
        '    RPT = "POROPEN4"
        '    isLongReport = False
        'Else
        If Absx1.chkFor("CHKPRIORYEAR").Checked Then
            RPT = "POROPEN7"
            'If Absx1.chkFor("CHKSHOWLINES").Checked Then
            '    CR_params.Add("SHOWLINES", "1")
            'Else
            CR_params.Add("SHOWLINES", "0")
            'End If
        Else
            RPT = "POROPEN6"
        End If

        'isLongReport = True
        'End If

        For i As Integer = 1 To 8
            CR_params.Add("GroupTitle" & i, dateTitles(i - 1))
            If i > 1 Then
                If Absx1.chkFor("ETAGROUP" & i).Checked Then
                    CR_params.Add("GroupShow" & i, "1")
                Else
                    CR_params.Add("GroupShow" & i, "0")
                End If
            End If
        Next
        'If isLongReport Then
        '    For i As Integer = 9 To 12
        '        CR_params.Add("GroupTitle" & i, dateTitles(i - 1))
        '        If Absx1.chkFor("ETAGROUP" & i).Checked Then
        '            CR_params.Add("GroupShow" & i, "1")
        '        Else
        '            CR_params.Add("GroupShow" & i, "0")
        '        End If
        '    Next
        'End If

        CR_params.Add("COSTINIT", "1")

        'If Absx1.chkFor("CHKQTYOH").Checked Then
        '    CR_params.Add("SHOWQTYOH", "1")
        'Else
        CR_params.Add("SHOWQTYOH", "0")
        'End If

        'If Absx1.chkFor("CHKZEROPO").Checked Then
        '    CR_params.Add("ZEROPO", "1")
        'Else
        CR_params.Add("ZEROPO", "0")
        'End If

        CR_params.Add("SUBT", txtDescription.Text & SUBT)

        Generate_Report(RPT, , SUBT)
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
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
            '    EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            'End If
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length > 4 Then
                EMsg &= vbCr & "Maximum number of Sort Fields for this report is 4"
            End If
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub CalculatePeriods()
        Dim maxPeriod As Integer = getMaxPeriods()
        buildDates(maxPeriod)
        For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select()
            'Dim etaDate As Date = CDate(Format(rowPOTORDRX.Item("PO_DATE_ETA"), "dd-MMM-yyyy"))
            Dim etaDate As Date = GetBestETA(rowPOTORDRX.Item("PO_ORDER_NO"), rowPOTORDRX.Item("PO_ORDER_LNO"))
            For i As Integer = 1 To maxPeriod
                If (etaDate >= BDate(i - 1) And etaDate <= EDate(i - 1)) Then
                    rowPOTORDRX.Table.Columns.Item("PO_QTY_OPN" & i).ReadOnly = False
                    rowPOTORDRX.Item("PO_QTY_OPN" & i) = Val(rowPOTORDRX.Item("PO_QTY_OPN" & i) & "") + Val(rowPOTORDRX.Item("PO_QTY_REC") & "")
                End If
            Next
        Next
        If Absx1.chkFor("CHKPRIORYEAR").Checked Then
            For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select()
                Dim etaDate As Date = GetBestETA(rowPOTORDRX.Item("PO_ORDER_NO"), rowPOTORDRX.Item("PO_ORDER_LNO"))
                For y As Integer = 1 To maxPeriod
                    If (etaDate >= BDate(y - 1).AddYears(-1) And etaDate <= EDate(y - 1).AddYears(-1)) Then
                        rowPOTORDRX.Table.Columns.Item("PO_QTY_OPN_LY" & y).ReadOnly = False
                        rowPOTORDRX.Item("PO_QTY_OPN_LY" & y) = Val(rowPOTORDRX.Item("PO_QTY_OPN_LY" & y) & "") + Val(rowPOTORDRX.Item("PO_QTY_REC") & "")
                    End If
                Next
            Next
        End If

        'For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select()
        '    Dim etaDate As Date = GetBestETA(rowASTSRPT1.Item("PO_ORDER_NO"), rowASTSRPT1.Item("PO_ORDER_LNO"))
        '    For i As Integer = 1 To maxPeriod
        '        If (etaDate >= BDate(i - 1) And etaDate <= EDate(i - 1)) Then
        '            rowASTSRPT1.Item("SHIP_OPN" & i) = Val(rowASTSRPT1.Item("SHIP_OPN" & i) & "") + Val(rowASTSRPT1.Item("SHIP_OPN") & "")
        '        End If
        '    Next
        'Next

    End Sub

    Private Sub buildDates(maxPeriod As Integer, Optional ByVal useVerbose As Boolean = False)
        BDate.Clear()
        EDate.Clear()
        If Absx1.chkFor("CHKPO_DATE_ETA_F1").Checked Then
            BDate.Add(DateSerial(2000, 1, 1))
        Else
            BDate.Add(CDate(Format(Absx1.dteFor("PO_DATE_ETA_F1").Value, "dd-MMM-yyyy")))
        End If
        EDate.Add(CDate(Format(Absx1.dteFor("PO_DATE_ETA_L1").Value, "dd-MMM-yyyy")))

        For i As Integer = 2 To maxPeriod
            BDate.Add(CDate(Format(Absx1.dteFor("PO_DATE_ETA_F" & i).Value, "dd-MMM-yyyy")))
            If Absx1.chkFor("CHKPO_DATE_ETA_L" & i).Checked Then
                EDate.Add(DateSerial(2100, 12, 31))
            Else
                EDate.Add(CDate(Format(Absx1.dteFor("PO_DATE_ETA_L" & i).Value, "dd-MMM-yyyy")))
            End If
        Next
    End Sub

    Private Function buildDateTitles(maxPeriod As Integer) As List(Of String)
        Dim retVal As New List(Of String)
        Dim date1 As String = ""
        Dim date2 As String = ""
        If Absx1.chkFor("CHKPO_DATE_ETA_F1").Checked Then
            date1 = "First"
        Else
            date1 = Format(CDate(Absx1.dteFor("PO_DATE_ETA_F1").Value), "MM/dd/yy")
        End If
        date2 = Format(CDate(Absx1.dteFor("PO_DATE_ETA_L1").Value), "MM/dd/yy")
        retVal.Add(date1 & " to " & date2)

        For i As Integer = 2 To maxPeriod
            date1 = Format(CDate(Absx1.dteFor("PO_DATE_ETA_F" & i).Value), "MM/dd/yy")
            If Absx1.chkFor("CHKPO_DATE_ETA_L" & i).Checked Then
                date2 = "Last"
            Else
                date2 = Format(CDate(Absx1.dteFor("PO_DATE_ETA_L" & i).Value), "MM/dd/yy")
            End If
            retVal.Add(date1 & " to " & date2)
        Next
        For i As Integer = maxPeriod + 1 To 12
            retVal.Add("N/A")
        Next
        Return retVal
    End Function

    Function Get_Dates(ByVal Period As Integer, Optional ByVal PriorYear As Boolean = False) As String
        Dim sql As String = ""

        Dim COL_NAMEs() As String
        Dim CONTROL_NAMEs() As String
        'If Type = "O" Then
        COL_NAMEs = New String() {"PO_DATE_ETA"}
        CONTROL_NAMEs = New String() {"PO_DATE_ETA"}
        'Else
        '    COL_NAMEs = New String() {"PO_SHIP_ETA"}
        '    CONTROL_NAMEs = New String() {"PO_DATE_ETA"}
        'End If
        Dim ctlIndex As Integer = 0
        For Each COLUMN_NAME As String In COL_NAMEs
            Dim CONTROL_NAME As String = CONTROL_NAMEs(ctlIndex)
            Dim BPeriod As String = ""
            If Absx1.chkFor("CHKPO_DATE_ETA_F1").Checked And Period = 1 Then
                BPeriod = Format(DateSerial(2000, 1, 1), "dd-MMM-yyyy")
            Else
                BPeriod = Format(Absx1.dteFor(CONTROL_NAME & "_F" & Period).Value, "dd-MMM-yyyy")
                If PriorYear Then
                    BPeriod = Format(CDate(BPeriod).AddYears(-1), "dd-MMM-yyyy")
                End If
            End If
            Dim LPeriod As String = ""
            If Period > 1 Then
                If Absx1.chkFor("CHKPO_DATE_ETA_L" & Period).Checked And Period > 1 Then
                    LPeriod = Format(DateSerial(2100, 12, 1), "dd-MMM-yyyy")
                Else
                    LPeriod = Format(Absx1.dteFor(CONTROL_NAME & "_L" & Period).Value, "dd-MMM-yyyy")
                    If PriorYear Then
                        LPeriod = Format(CDate(LPeriod).AddYears(-1), "dd-MMM-yyyy")
                    End If
                End If
            Else
                LPeriod = Format(Absx1.dteFor(CONTROL_NAME & "_L" & Period).Value, "dd-MMM-yyyy")
                If PriorYear Then
                    LPeriod = Format(CDate(LPeriod).AddYears(-1), "dd-MMM-yyyy")
                End If
            End If
            sql = sql & " and A." & COLUMN_NAME & " >= '" & BPeriod & "'"
            sql = sql & " and A." & COLUMN_NAME & " <= '" & LPeriod & "'"
            ctlIndex += 1
        Next
        'If Type = "S" Then
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            sql = Replace(sql, "A.PO_DATE_ETA", "TRUNC(POTSHIP2.PO_DATE_RECEIVED)")
            sql = Replace(sql, "A.", "POTSHIP1.")
            'Else
            '    sql = Replace(sql, "A.", "POTORDR2.")
            'End If
        Else
            sql = Replace(sql, "A.PO_DATE_ETA", "POTSHIP2.PO_DATE_RECEIVED")
            sql = Replace(sql, "A.", "POTSHIP1.")
            'Else
            '    sql = Replace(sql, "A.", "POTORDR2.")
            'End If
        End If

        Return sql
    End Function

    Private Function getMaxPeriods() As Integer
        Dim retval As Integer = 0
        For i As Integer = 2 To 12
            If Absx1.chkFor("ETAGROUP" & i).Visible Then
                If Absx1.chkFor("ETAGROUP" & i).Checked Then
                    retval = i
                Else
                    retval = i - 1
                End If
            End If
        Next
        Return retval
    End Function

    Private Function GetNextMontFirstDay(ByVal thisDate As Date) As Date
        Dim retVal As Date
        Dim MO As Integer = thisDate.Month
        Dim YR As Integer = thisDate.Year
        Dim DY As Integer = 1
        retVal = DateSerial(YR, MO, DY).AddMonths(1)
        Return retVal
    End Function

    Private Function GetNextMontLastDay(ByVal thisDate As Date) As Date
        Dim retVal As Date
        Dim MO As Integer = thisDate.Month
        Dim YR As Integer = thisDate.Year
        Dim DY As Integer = 1
        retVal = DateSerial(YR, MO, DY).AddMonths(1).AddDays(-1)
        Return retVal
    End Function

    Private Function IsFirstDayOfMonth(ByVal D As Date) As Boolean
        Dim Retval As Boolean = False
        Dim FirstDay As Date = DateSerial(D.Year, D.Month, 1)
        If CDate(Format(D, "MM/dd/yyyy")) = CDate(Format(FirstDay, "MM/dd/yyyy")) Then
            Retval = True
        End If
        Return Retval
    End Function

    Private Function IsLastDayOfMonth(ByVal D As Date) As Boolean
        Dim Retval As Boolean = False
        Dim LastDay As Date = DateSerial(D.Year, D.Month + 1, 1).AddDays(-1)
        If CDate(Format(D, "MM/dd/yyyy")) = CDate(Format(LastDay, "MM/dd/yyyy")) Then
            Retval = True
        End If
        Return Retval
    End Function

    Private Sub setETAValues(ByVal Group As Integer)

        Dim priorFirst As Date
        If IsDate(Absx1.dteFor("PO_DATE_ETA_F" & Group - 1).Value()) Then
            priorFirst = CDate(Absx1.dteFor("PO_DATE_ETA_F" & Group - 1).Value())
        Else
            Dim MO As Integer = CDate(Absx1.dteFor("PO_DATE_ETA_L" & Group - 1).Value()).Month
            Dim YR As Integer = CDate(Absx1.dteFor("PO_DATE_ETA_L" & Group - 1).Value()).Year
            priorFirst = DateSerial(YR, MO, 1)
        End If
        If Not IsFirstDayOfMonth(priorFirst) Then
            priorFirst = DateSerial(priorFirst.Year, priorFirst.Month, 1)
        End If

        Dim priorLast As Date
        If IsDate(Absx1.dteFor("PO_DATE_ETA_L" & Group - 1).Value()) Then
            priorLast = CDate(Absx1.dteFor("PO_DATE_ETA_L" & Group - 1).Value())
            If Not IsLastDayOfMonth(priorLast) Then
                Dim MO As Integer = priorLast.Month
                Dim YR As Integer = priorLast.Year
                Dim DY As Integer = 1
                priorLast = DateSerial(YR, MO, DY).AddMonths(1).AddDays(-1)
            End If
        Else
            priorLast = CDate(Absx1.dteFor("PO_DATE_ETA_F" & Group - 1).Value())
            Dim MO As Integer = priorLast.Month
            Dim YR As Integer = priorLast.Year
            Dim DY As Integer = 1
            priorLast = DateSerial(YR, MO, DY).AddMonths(1).AddDays(-1)
        End If

        Absx1.dteFor("PO_DATE_ETA_F" & Group).DateTime = priorLast.AddDays(1)
        'Absx1.dteFor("PO_DATE_ETA_L" & Group).DateTime = priorFirst.AddMonths(2).AddDays(-1)
        'This might hurt a little
        Absx1.dteFor("PO_DATE_ETA_L" & Group).DateTime = priorLast.AddDays(1).AddMonths(DateDiff(DateInterval.Month, priorFirst, priorLast) + 1).AddDays(-1)

        Absx1.CtlFor("GRP_DATE_ETA" & Group).Visible = Absx1.chkFor("ETAGROUP" & Group).Checked

        If Group < 8 Then


            For gp As Integer = Group + 1 To 12
                'Absx1.CtlFor("CHKPO_DATE_ETA_L" & Group + 1).Visible = False
                If gp >= Group + 1 Then
                    If gp = Group + 1 Then
                        If Absx1.chkFor("ETAGROUP" & Group).Visible Then
                            Absx1.chkFor("ETAGROUP" & gp).Visible = Absx1.chkFor("ETAGROUP" & Group).Checked
                            Absx1.chkFor("ETAGROUP" & gp).Checked = False
                        Else
                            Absx1.chkFor("ETAGROUP" & gp).Visible = False
                        End If
                    Else
                        Absx1.chkFor("ETAGROUP" & gp).Visible = False
                    End If

                    Absx1.CtlFor("GRP_DATE_ETA" & gp).Visible = False

                    'Absx1.chkFor("CHKPO_DATE_ETA_L" & gp).Checked = False

                    'Absx1.dteFor("PO_DATE_ETA_F" & gp).Text = ""
                    'Absx1.dteFor("PO_DATE_ETA_L" & gp).Text = ""
                End If
            Next
        End If
    End Sub

    Private Sub setETALastValue(ByVal group As Integer)
        If Absx1.chkFor("CHKPO_DATE_ETA_L" & group).Checked Then
            Absx1.dteFor("PO_DATE_ETA_L" & group).ReadOnly = True
            Absx1.dteFor("PO_DATE_ETA_L" & group).Text = ""
            If group <> 8 Then
                Absx1.chkFor("ETAGROUP" & group + 1).Checked = False
                Absx1.chkFor("ETAGROUP" & group + 1).Visible = False
            End If
        Else
            Absx1.dteFor("PO_DATE_ETA_L" & group).ReadOnly = False
            Absx1.dteFor("PO_DATE_ETA_L" & group).Text = Absx1.dteFor("PO_DATE_ETA_F2").Text
            If group <> 8 Then
                Absx1.chkFor("ETAGROUP" & group + 1).Checked = False
                Absx1.chkFor("ETAGROUP" & group + 1).Visible = True
            End If
        End If
    End Sub
#End Region

#Region "Form Controls"
#Region "Check Boxes"
    Private Sub chkETAGROUP2_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP2.CheckedChanged
        setETAValues(2)
    End Sub

    Private Sub chkETAGROUP3_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP3.CheckedChanged
        setETAValues(3)
    End Sub

    Private Sub chkETAGROUP4_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP4.CheckedChanged
        setETAValues(4)
    End Sub

    Private Sub chkETAGROUP5_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP5.CheckedChanged
        setETAValues(5)
    End Sub

    Private Sub chkETAGROUP6_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP6.CheckedChanged
        setETAValues(6)
    End Sub

    Private Sub chkETAGROUP7_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP7.CheckedChanged
        setETAValues(7)
    End Sub

    Private Sub chkETAGROUP8_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP8.CheckedChanged
        setETAValues(8)
    End Sub

    Private Sub chkETAGROUP9_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP9.CheckedChanged
        setETAValues(9)
    End Sub

    Private Sub chkETAGROUP10_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP10.CheckedChanged
        setETAValues(10)
    End Sub

    Private Sub chkETAGROUP11_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP11.CheckedChanged
        setETAValues(11)
    End Sub

    Private Sub chkETAGROUP12_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP12.CheckedChanged
        setETAValues(12)
    End Sub

    Private Sub chkPO_DATE_ETA_L2_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L2.CheckedChanged
        setETALastValue(2)
    End Sub

    Private Sub chkPO_DATE_ETA_L3_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L3.CheckedChanged
        setETALastValue(3)
    End Sub

    Private Sub chkPO_DATE_ETA_L4_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L4.CheckedChanged
        setETALastValue(4)
    End Sub

    Private Sub chkPO_DATE_ETA_L5_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L5.CheckedChanged
        setETALastValue(5)
    End Sub

    Private Sub chkPO_DATE_ETA_L6_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L6.CheckedChanged
        setETALastValue(6)
    End Sub

    Private Sub chkPO_DATE_ETA_L7_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L7.CheckedChanged
        setETALastValue(7)
    End Sub

    Private Sub chkPO_DATE_ETA_L8_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L8.CheckedChanged
        setETALastValue(8)
    End Sub

    Private Sub chkPO_DATE_ETA_L9_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L9.CheckedChanged
        setETALastValue(9)
    End Sub

    Private Sub chkPO_DATE_ETA_L10_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L10.CheckedChanged
        setETALastValue(10)
    End Sub

    Private Sub chkPO_DATE_ETA_L11_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L11.CheckedChanged
        setETALastValue(11)
    End Sub

    Private Sub chkPO_DATE_ETA_L12_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L12.CheckedChanged
        setETALastValue(12)
    End Sub

    Private Sub CHKPRIORYEAR_CheckedChanged(sender As Object, e As EventArgs) Handles CHKPRIORYEAR.CheckedChanged
        'CHKSHOWLINES.Visible = CHKPRIORYEAR.Checked
        'CHKSHOWLINES.Checked = CHKPRIORYEAR.Checked
    End Sub
#End Region
#End Region

    Private Function GetBestETA(ByVal PO_ORDER_NO As String, ByVal PO_ORDER_LNO As Integer) As Date
        Dim RetVal As Date = CDate("01/01/1800")
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT TRUNC(POTSHIP2.PO_DATE_RECEIVED) AS PO_DATE_RECEIVED")
        SQLS.AppendLine("FROM POTSHIP2,POTSHIP3")
        SQLS.AppendLine("WHERE POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO")
        SQLS.AppendLine("AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO")
        SQLS.AppendLine(String.Format("AND POTSHIP3.PO_ORDER_NO = '{0}'", PO_ORDER_NO))
        SQLS.AppendLine(String.Format("AND POTSHIP3.PO_ORDER_LNO = {0}", PO_ORDER_LNO))
        ASCMAIN1.sql = SQLS.ToString()
        Dim DATE_STRING As String = ASCDATA1.GetDataValue
        If IsDate(DATE_STRING) Then
            RetVal = CDate(DATE_STRING)
        Else
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then Stop
        End If
        'Dim filter As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' AND PO_ORDER_LNO = " & PO_ORDER_LNO
        ''Dim rowPOTORDRX As DataRow = dst.Tables.Item("POTORDRX").Select(filter).FirstOrDefault()
        ''If Not IsNothing(rowPOTORDRX) Then
        ''    If IsDate(rowPOTORDRX.Item("PO_DATE_ETA").ToString()) Then
        ''        RetVal = CDate(rowPOTORDRX.Item("PO_DATE_ETA").ToString())
        ''    End If
        ''End If

        'Dim rowPOTSHIPX As DataRow = dst.Tables.Item("POTSHIPX").Select(filter).FirstOrDefault()
        'If Not IsNothing(rowPOTSHIPX) Then
        '    If IsDate(rowPOTSHIPX.Item("PO_DATE_RECEIVED").ToString()) Then
        '        RetVal = CDate(rowPOTSHIPX.Item("PO_DATE_RECEIVED").ToString())
        '    End If
        'End If
        Return RetVal
    End Function

    'Private Sub AbsCheckBox2_CheckedChanged(sender As Object, e As EventArgs)
    '    Absx1.chkFor("CHKZEROPO").Visible = Absx1.chkFor("CHKQTYOH").Checked
    'End Sub
End Class