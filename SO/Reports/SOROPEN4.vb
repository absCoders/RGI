Public Class SOROPEN4

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"

        Range_Events(grpPO_DATE_ETA1)

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = ""
        If Absx1.optFor("OPTASN").Value = "S" Then
            SUBT &= "Stock Styles Only"
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            SUBT &= "Non-Stock Styles Only"
        End If

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")
        Dim sql_TABLE_NAMEs_orig As String = sql_TABLE_NAMEs
        Dim sql_JOIN_orig As String = sql_JOIN

        For Period As Integer = 1 To 4
            Dim ProcessPeriod As Boolean = True
            If Period = 1 Then
                ProcessPeriod = True
            Else
                If Absx1.chkFor("ETAGROUP" & Period).Checked Then
                    ProcessPeriod = True
                Else
                    ProcessPeriod = False
                End If
            End If

            If ProcessPeriod Then
                For Each OS As String In New String() {"O", "S"}
                    If Absx1.chkFor("CHKINCL_" & OS).Checked Then
                        Dim sql_filter2 As String = Get_Dates(OS, Period)

                        If Absx1.optFor("OPTASN").Value = "S" Then
                            sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Null"
                        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
                            sql_filter2 &= "   and ICTSTYL1.CUST_CODE is Not Null"
                        End If

                        If Absx1.optFor("OPTSHOW").Value = "O" Then
                            If OS = "O" Then
                                sql_filter2 &= "" _
                                    & " AND SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
                                    & " AND NVL(SOTORDR2.ORDR_QTY_OPEN,0) <> 0" & vbCrLf
                            Else
                                sql_filter2 &= "" _
                                    & " AND SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
                                    & " AND NVL(SOTORDR2.ORDR_QTY_OPEN,0) = 0" & vbCrLf
                            End If
                        End If

                        If OS = "S" Then
                            sql_TABLE_NAMEs = sql_TABLE_NAMEs_orig & ",POTSHIP2,POTSHIP3"
                            'sql_JOIN = sql_JOIN_orig & " " _
                            '    & " and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                            '    & " and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                            '    & " and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                            '    & " and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                            '    & " and POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf
                            If Not sql_TABLE_NAMEs.Contains("POTSHIP1") Then
                                sql_TABLE_NAMEs = sql_TABLE_NAMEs & ",POTSHIP1"
                            End If
                        Else
                            sql_TABLE_NAMEs = sql_TABLE_NAMEs_orig
                            sql_JOIN = sql_JOIN_orig
                        End If

                        ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf _
                            & ", POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                            & IIf(OS = "O", _
                                  ", 'OPENPO' PO_SHIPMENT_NO, 0 PO_SHIPMENT_LNO" & vbCrLf, _
                                  ", POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf) _
                            & IIf(OS = "O", _
                                  ", POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_SHP, POTORDR2.PO_QTY_REC, POTORDR2.PO_QTY_OPN, 0 SHIP_QTY, 0 SHIP_OPN, 0 SHIP_REC" & vbCrLf, _
                                  ", 0 PO_QTY_ORD, 0 PO_QTY_SHP, 0 PO_QTY_REC, 0 PO_QTY_OPN, POTSHIP3.PO_QTY_SHP SHIP_QTY, DECODE (POTSHIP2.PO_SHIP_STATUS,'O',POTSHIP3.PO_QTY_SHP,0) SHIP_OPN, POTSHIP3.PO_QTY_REC SHIP_REC" & vbCrLf) _
                            & " from POTORDR2" & sql_TABLE_NAMEs & vbCrLf _
                            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf

                        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                                            & " (" & G1thru9 _
                                            & ",PO_ORDER_NO,PO_ORDER_LNO,PO_SHIPMENT_NO,PO_SHIPMENT_LNO" _
                                            & ",PO_QTY_ORD,PO_QTY_SHP,PO_QTY_REC,PO_QTY_OPN,SHIP_QTY,SHIP_OPN,SHIP_REC" _
                                            & ") " _
                                            & " (" & ASCMAIN1.sql & ")")

                    End If
                Next
            End If
        Next

        Prepare_dst(True, sql_filter)
    End Sub

    Private Sub CalculatePeriods()
        Dim g1BDate As Date
        If Absx1.chkFor("CHKPO_DATE_ETA_F1").Checked Then
            g1BDate = DateSerial(2000, 1, 1)
        Else
            g1BDate = CDate(Format(Absx1.dteFor("PO_DATE_ETA_F1").Value, "dd-MMM-yyyy"))
        End If
        Dim g1EDate As Date = CDate(Format(Absx1.dteFor("PO_DATE_ETA_L1").Value, "dd-MMM-yyyy"))

        Dim g2BDate As Date = CDate(Format(Absx1.dteFor("PO_DATE_ETA_F2").Value, "dd-MMM-yyyy"))
        Dim g2EDate As Date
        If Absx1.chkFor("CHKPO_DATE_ETA_L2").Checked Then
            g2EDate = DateSerial(2100, 12, 31)
        Else
            g2EDate = CDate(Format(Absx1.dteFor("PO_DATE_ETA_L2").Value, "dd-MMM-yyyy"))
        End If
        Dim g2Checked As Boolean = Absx1.chkFor("ETAGROUP2").Checked

        Dim g3BDate As Date = CDate(Format(Absx1.dteFor("PO_DATE_ETA_F3").Value, "dd-MMM-yyyy"))
        Dim g3EDate As Date
        If Absx1.chkFor("CHKPO_DATE_ETA_L3").Checked Then
            g3EDate = DateSerial(2100, 12, 31)
        Else
            g3EDate = CDate(Format(Absx1.dteFor("PO_DATE_ETA_L3").Value, "dd-MMM-yyyy"))
        End If
        Dim g3Checked As Boolean = Absx1.chkFor("ETAGROUP3").Checked

        Dim g4BDate As Date = CDate(Format(Absx1.dteFor("PO_DATE_ETA_F4").Value, "dd-MMM-yyyy"))
        Dim g4EDate As Date
        If Absx1.chkFor("CHKPO_DATE_ETA_L4").Checked Then
            g4EDate = DateSerial(2100, 12, 31)
        Else
            g4EDate = CDate(Format(Absx1.dteFor("PO_DATE_ETA_L4").Value, "dd-MMM-yyyy"))
        End If
        Dim g4Checked As Boolean = Absx1.chkFor("ETAGROUP4").Checked

        For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select()
            rowPOTORDRX.Table.Columns.Item("PO_QTY_OPN1").ReadOnly = False
            rowPOTORDRX.Table.Columns.Item("PO_QTY_OPN2").ReadOnly = False
            rowPOTORDRX.Table.Columns.Item("PO_QTY_OPN3").ReadOnly = False
            rowPOTORDRX.Table.Columns.Item("PO_QTY_OPN4").ReadOnly = False
            Dim etaDate As Date = CDate(Format(rowPOTORDRX.Item("PO_DATE_ETA"), "dd-MMM-yyyy"))
            If (etaDate >= g1BDate And etaDate <= g1EDate) Then
                rowPOTORDRX.Item("PO_QTY_OPN1") = Val(rowPOTORDRX.Item("PO_QTY_OPN1") & "") + rowPOTORDRX.Item("PO_QTY_OPN")
                UpdateASTSRPT1(rowPOTORDRX.Item("PO_ORDER_NO"), rowPOTORDRX.Item("PO_ORDER_LNO"), 1)
            End If
            If g2Checked And (etaDate >= g2BDate And etaDate <= g2EDate) Then
                rowPOTORDRX.Item("PO_QTY_OPN2") = Val(rowPOTORDRX.Item("PO_QTY_OPN2") & "") + rowPOTORDRX.Item("PO_QTY_OPN")
                UpdateASTSRPT1(rowPOTORDRX.Item("PO_ORDER_NO"), rowPOTORDRX.Item("PO_ORDER_LNO"), 2)
            End If
            If g3Checked And (etaDate >= g3BDate And etaDate <= g3EDate) Then
                rowPOTORDRX.Item("PO_QTY_OPN3") = Val(rowPOTORDRX.Item("PO_QTY_OPN3") & "") + rowPOTORDRX.Item("PO_QTY_OPN")
                UpdateASTSRPT1(rowPOTORDRX.Item("PO_ORDER_NO"), rowPOTORDRX.Item("PO_ORDER_LNO"), 3)
            End If
            If g4Checked And (etaDate >= g4BDate And etaDate <= g4EDate) Then
                rowPOTORDRX.Item("PO_QTY_OPN4") = Val(rowPOTORDRX.Item("PO_QTY_OPN4") & "") + rowPOTORDRX.Item("PO_QTY_OPN")
                UpdateASTSRPT1(rowPOTORDRX.Item("PO_ORDER_NO"), rowPOTORDRX.Item("PO_ORDER_LNO"), 4)
            End If
        Next
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        For i As Integer = 1 To 4
            dst.Tables.Item("ASTSRPT1").Columns.Add("SHIP_OPN" & Format(i, "0"), GetType(System.Double))
        Next

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC " _
            & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        Dim SQL As New Text.StringBuilder With {.Length = 0}
        SQL.AppendLine("Select POTORDR2.*,ICTSTYL1.CASE_CUBE, ICTCOLR1.COLOR_DESC")
        SQL.AppendLine(", POTORDR1.VEND_CODE, POTORDR1.FACTORY_CODE, POTORDR1.PO_REFERENCE, POTORDR1.WHSE_CODE")
        SQL.AppendLine(", POTORDR1.PO_SPEC_ORDR_NO, POTORDR1.PO_DATE_ORDERED, POTORDR1.PO_DATE_CANCEL,POTORDR1.PORT_CODE_ORIG,POTORDR1.PO_SHIP_VIA")
        SQL.AppendLine(", 0 as PO_QTY_OPN1, 0 as PO_QTY_OPN2, 0 as PO_QTY_OPN3, 0 as PO_QTY_OPN4")
        'SQL.AppendLine(", " & IIf(optSORT.Value & "" = "D", "TO_CHAR(POTORDR2.PO_DATE_SHIP_BY,'YYYYMMDD')", IIf(optSORT.Value & "" = "S", "POTORDR2.STYLE_CODE", "POTORDR2.PO_ORDER_NO")) & " SORT1")
        'SQL.AppendLine(", " & IIf(optSORT.Value & "" = "D", "POTORDR2.PO_ORDER_NO", IIf(optSORT.Value & "" = "S", "POTORDR2.COLOR_CODE", "TO_CHAR(POTORDR2.PO_ORDER_LNO,'000000')")) & " SORT2")
        SQL.AppendLine(" from POTORDR2,POTORDR1,ICTCOLR1, ICTSTYL1 ")
        SQL.AppendLine(" where (POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO) in (Select Distinct PO_ORDER_NO,PO_ORDER_LNO from " & ASTSRPT1 & ")")
        SQL.AppendLine("   and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
        SQL.AppendLine("   and POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        SQL.AppendLine("   and ICTCOLR1.COLOR_CODE = POTORDR2.COLOR_CODE")
        ASCMAIN1.sql = SQL.ToString
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDRX", 2))

        ASCMAIN1.sql = "Select POTSHIP3.*" & vbCrLf _
            & ", POTSHIP2.CONTAINER_NO, POTSHIP2.PO_SHIP_STATUS SHIP_STATUS" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_REF_NO" & vbCrLf _
            & ", POTSHIP1.PO_DATE_SHIPPED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
            & " from POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
            & " where (POTSHIP3.PO_SHIPMENT_NO,POTSHIP3.PO_SHIPMENT_LNO) in (Select Distinct PO_SHIPMENT_NO,PO_SHIPMENT_LNO from " & ASTSRPT1 & ")" & vbCrLf _
            & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTSHIPX", 4))
    End Sub

    Public Overrides Sub Print_Report()
        CalculatePeriods()

        'CR_params.Add("SUBT", txtDescription.Text & SUBT)
        RPT = "POROPEN4"
        'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
        '    RPT = "POROPEN3"
        'End If
        Dim F1 As String = Format(Absx1.dteFor("PO_DATE_ETA_F1").Value, "MM/dd/yy")
        If F1 = "" Then
            F1 = "First"
        End If
        Dim L2 As String = Format(Absx1.dteFor("PO_DATE_ETA_L2").Value, "MM/dd/yy")
        If L2 = "" Then
            L2 = "Last"
        End If
        Dim L3 As String = Format(Absx1.dteFor("PO_DATE_ETA_L3").Value, "MM/dd/yy")
        If L3 = "" Then
            L3 = "Last"
        End If
        Dim L4 As String = Format(Absx1.dteFor("PO_DATE_ETA_L4").Value, "MM/dd/yy")
        If L4 = "" Then
            L4 = "Last"
        End If

        CR_params.Add("GroupTitle1", F1 & "-" & Format(Absx1.dteFor("PO_DATE_ETA_L1").Value, "MM/dd/yy"))
        CR_params.Add("GroupTitle2", Format(Absx1.dteFor("PO_DATE_ETA_F2").Value, "MM/dd/yy") & "-" & L2)
        CR_params.Add("GroupTitle3", Format(Absx1.dteFor("PO_DATE_ETA_F3").Value, "MM/dd/yy") & "-" & L3)
        CR_params.Add("GroupTitle4", Format(Absx1.dteFor("PO_DATE_ETA_F4").Value, "MM/dd/yy") & "-" & L4)
        If Absx1.chkFor("ETAGROUP2").Checked Then
            CR_params.Add("GroupShow2", "1")
        Else
            CR_params.Add("GroupShow2", "0")
        End If
        If Absx1.chkFor("ETAGROUP3").Checked Then
            CR_params.Add("GroupShow3", "1")
        Else
            CR_params.Add("GroupShow3", "0")
        End If
        If Absx1.chkFor("ETAGROUP4").Checked Then
            CR_params.Add("GroupShow4", "1")
        Else
            CR_params.Add("GroupShow4", "0")
        End If
        If Absx1.chkFor("CHKCOSTINIT").Checked Then
            CR_params.Add("COSTINIT", "1")
            If SUBT.Length > 0 Then
                SUBT = SUBT & " | Using Initial Cost"
            Else
                SUBT = SUBT & "Using Initial Cost"
            End If

        Else
            CR_params.Add("COSTINIT", "0")
        End If

        CR_params.Add("SUBT", txtDescription.Text & SUBT)

        Generate_Report(RPT, , SUBT)
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

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        If parms.Length > 0 Then
        End If

        EnforceConstraints(False)
        'Fill_Records("ASTSRPT1")
        EnforceConstraints(True)
    End Sub

    Function Get_Dates(TYPE As String, ByVal Period As Integer) As String
        Dim sql As String = ""

        Dim COLUMN_NAMEs() As String
        Dim CONTROL_NAMEs() As String
        If TYPE = "O" Then
            COLUMN_NAMEs = New String() {"PO_DATE_ETA"}
            CONTROL_NAMEs = New String() {"PO_DATE_ETA"}
        Else
            COLUMN_NAMEs = New String() {"PO_SHIP_ETA"}
            CONTROL_NAMEs = New String() {"PO_DATE_ETA"}
        End If
        Dim ctlIndex As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Dim CONTROL_NAME As String = CONTROL_NAMEs(ctlIndex)
            Dim BPeriod As String = ""
            If Absx1.chkFor("CHKPO_DATE_ETA_F1").Checked And Period = 1 Then
                BPeriod = Format(DateSerial(2000, 1, 1), "dd-MMM-yyyy")
            Else
                BPeriod = Format(Absx1.dteFor(CONTROL_NAME & "_F" & Period).Value, "dd-MMM-yyyy")
            End If
            Dim LPeriod As String = ""
            If Period > 1 Then
                If Absx1.chkFor("CHKPO_DATE_ETA_L" & Period).Checked And Period > 1 Then
                    LPeriod = Format(DateSerial(2100, 12, 1), "dd-MMM-yyyy")
                Else
                    LPeriod = Format(Absx1.dteFor(CONTROL_NAME & "_L" & Period).Value, "dd-MMM-yyyy")
                End If
            Else
                LPeriod = Format(Absx1.dteFor(CONTROL_NAME & "_L" & Period).Value, "dd-MMM-yyyy")
            End If
            sql = sql & " and A." & COLUMN_NAME & " >= '" & BPeriod & "'"
            sql = sql & " and A." & COLUMN_NAME & " <= '" & LPeriod & "'"
            ctlIndex += 1
        Next
        If TYPE = "S" Then
            sql = Replace(sql, "A.PO_DATE_RECEIVED", "POTSHIP2.PO_DATE_RECEIVED")
            sql = Replace(sql, "A.", "POTSHIP1.")
        Else
            sql = Replace(sql, "A.", "POTORDR2.")
        End If
        Return sql
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

    Private Sub chkETAGROUP2_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP2.CheckedChanged
        Dim F1 As Date = CDate(Absx1.dteFor("PO_DATE_ETA_F1").Value())
        Dim L1 As Date = CDate(Absx1.dteFor("PO_DATE_ETA_L1").Value())
        Dim dayDiff As Integer = DateDiff("d", F1, L1)

        grpPO_DATE_ETA2.Visible = chkETAGROUP2.Checked
        If Not chkETAGROUP2.Checked Then
            chkETAGROUP3.Checked = False
            chkETAGROUP3.Visible = False
            chkETAGROUP4.Checked = False
            chkETAGROUP4.Visible = False
        Else
            chkETAGROUP3.Visible = True
        End If
        Dim F2 As Date = L1.AddDays(1).ToString()
        Dim L2 As Date = L1.AddDays(1).ToString()
        Absx1.dteFor("PO_DATE_ETA_F2").Text = F2
        Absx1.chkFor("CHKPO_DATE_ETA_L4").Visible = False
        Absx1.chkFor("CHKPO_DATE_ETA_L4").Checked = False
        Absx1.chkFor("CHKPO_DATE_ETA_L3").Visible = False
        Absx1.chkFor("CHKPO_DATE_ETA_L3").Checked = False
        Absx1.chkFor("CHKPO_DATE_ETA_L2").Visible = True
        Absx1.chkFor("CHKPO_DATE_ETA_L2").Checked = False

        If IsFirstDayOfMonth(F1) And IsLastDayOfMonth(L1) Then
            Absx1.dteFor("PO_DATE_ETA_L2").Text = DateSerial(F2.AddDays(1).Year, F2.AddDays(1).Month + 1, 1).AddDays(-1).ToString()
        Else
            If Absx1.chkFor("CHKPO_DATE_ETA_F1").Checked Then
                Absx1.dteFor("PO_DATE_ETA_L2").Text = L2
            Else
                Absx1.dteFor("PO_DATE_ETA_L2").Text = F2.AddDays(dayDiff)
            End If
        End If
    End Sub

    Private Sub chkETAGROUP3_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP3.CheckedChanged
        Dim F2 As Date = CDate(Absx1.dteFor("PO_DATE_ETA_F2").Value())
        Dim L2 As Date = CDate(Absx1.dteFor("PO_DATE_ETA_L2").Value())
        Dim dayDiff As Integer = DateDiff("d", F2, L2)

        grpPO_DATE_ETA3.Visible = chkETAGROUP3.Checked
        If Not chkETAGROUP3.Checked Then
            chkETAGROUP4.Checked = False
            chkETAGROUP4.Visible = False
        Else
            chkETAGROUP4.Visible = True
        End If
        Dim F3 As Date = L2.AddDays(1).ToString()
        Dim L3 As Date = L2.AddDays(1).ToString()
        Absx1.dteFor("PO_DATE_ETA_F3").Text = F3
        Absx1.chkFor("CHKPO_DATE_ETA_L4").Visible = False
        Absx1.chkFor("CHKPO_DATE_ETA_L4").Checked = False
        Absx1.chkFor("CHKPO_DATE_ETA_L3").Visible = True
        Absx1.chkFor("CHKPO_DATE_ETA_L3").Checked = False
        Absx1.chkFor("CHKPO_DATE_ETA_L2").Visible = False
        Absx1.chkFor("CHKPO_DATE_ETA_L2").Checked = False

        If IsFirstDayOfMonth(F2) And IsLastDayOfMonth(L2) Then
            Absx1.dteFor("PO_DATE_ETA_L3").Text = DateSerial(F3.AddDays(1).Year, F3.AddDays(1).Month + 1, 1).AddDays(-1).ToString()
        Else
            Absx1.dteFor("PO_DATE_ETA_L3").Text = F3.AddDays(dayDiff)
        End If
    End Sub

    Private Sub chkETAGROUP4_CheckedChanged(sender As Object, e As EventArgs) Handles chkETAGROUP4.CheckedChanged
        Dim F3 As Date = CDate(Absx1.dteFor("PO_DATE_ETA_F3").Value())
        Dim L3 As Date = CDate(Absx1.dteFor("PO_DATE_ETA_L3").Value())
        Dim dayDiff As Integer = DateDiff("d", F3, L3)

        grpPO_DATE_ETA4.Visible = chkETAGROUP4.Checked
        Dim F4 As Date = L3.AddDays(1).ToString()
        Dim L4 As Date = L3.AddDays(1).ToString()
        Absx1.dteFor("PO_DATE_ETA_F4").Text = F4
        Absx1.chkFor("CHKPO_DATE_ETA_L4").Visible = True
        Absx1.chkFor("CHKPO_DATE_ETA_L4").Checked = False
        Absx1.chkFor("CHKPO_DATE_ETA_L3").Visible = False
        Absx1.chkFor("CHKPO_DATE_ETA_L3").Checked = False
        Absx1.chkFor("CHKPO_DATE_ETA_L2").Visible = False
        Absx1.chkFor("CHKPO_DATE_ETA_L2").Checked = False

        If IsFirstDayOfMonth(F3) And IsLastDayOfMonth(L3) Then
            Absx1.dteFor("PO_DATE_ETA_L4").Text = DateSerial(F4.AddDays(1).Year, F4.AddDays(1).Month + 1, 1).AddDays(-1).ToString()
        Else
            Absx1.dteFor("PO_DATE_ETA_L4").Text = F4.AddDays(dayDiff)
        End If
    End Sub

    Private Sub UpdateASTSRPT1(ByVal PO_ORDER_NO As String, PO_ORDER_LNO As Integer, Group As Integer)
        Dim filter As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "' AND PO_ORDER_LNO = " & PO_ORDER_LNO
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select(filter)
            rowASTSRPT1.Item("SHIP_OPN" & Group) = Val(rowASTSRPT1.Item("SHIP_OPN" & Group) & "") + rowASTSRPT1.Item("SHIP_OPN")
        Next
    End Sub

    Private Sub chkPO_DATE_ETA_L2_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L2.CheckedChanged
        If chkPO_DATE_ETA_L2.Checked Then
            Absx1.dteFor("PO_DATE_ETA_L2").ReadOnly = True
            Absx1.dteFor("PO_DATE_ETA_L2").Text = ""
        Else
            Absx1.dteFor("PO_DATE_ETA_L2").ReadOnly = False
            Absx1.dteFor("PO_DATE_ETA_L2").Text = Absx1.dteFor("PO_DATE_ETA_F2").Text
        End If
    End Sub

    Private Sub chkPO_DATE_ETA_L3_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L3.CheckedChanged
        If chkPO_DATE_ETA_L3.Checked Then
            Absx1.dteFor("PO_DATE_ETA_L3").ReadOnly = True
            Absx1.dteFor("PO_DATE_ETA_L3").Text = ""
        Else
            Absx1.dteFor("PO_DATE_ETA_L3").ReadOnly = False
            Absx1.dteFor("PO_DATE_ETA_L3").Text = Absx1.dteFor("PO_DATE_ETA_F3").Text
        End If
    End Sub

    Private Sub chkPO_DATE_ETA_L4_CheckedChanged(sender As Object, e As EventArgs) Handles chkPO_DATE_ETA_L4.CheckedChanged
        If chkPO_DATE_ETA_L4.Checked Then
            Absx1.dteFor("PO_DATE_ETA_L4").ReadOnly = True
            Absx1.dteFor("PO_DATE_ETA_L4").Text = ""
        Else
            Absx1.dteFor("PO_DATE_ETA_L4").ReadOnly = False
            Absx1.dteFor("PO_DATE_ETA_L4").Text = Absx1.dteFor("PO_DATE_ETA_F4").Text
        End If
    End Sub

End Class