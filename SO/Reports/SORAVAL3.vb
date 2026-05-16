Imports System.Text

Public Class SORAVAL3
    Private ALLOC_BDATE As New List(Of Date)
    Private ALLOC_EDATE As New List(Of Date)
    Private ALLOC_SHOW As New List(Of Boolean)
    Dim SOTDEMD1 As String
    Dim SOTSUPP1 As String
    Dim edi850cust As List(Of String)
    Private xRYP0_legend As String = String.Empty
    Private xRYP0 As String = String.Empty
    Private xRYP1_legend As String = String.Empty
    Private xRYP1 As String = String.Empty
    Private xRYP2_legend As String = String.Empty
    Private xRYP2 As String = String.Empty
    Private xRYP3_legend As String = String.Empty
    Private xRYP3 As String = String.Empty
    Private FormReady As Boolean = False

#Region "ABS Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"
        IsLoading = True
        Get_PARM("SOTPARM1")

        'Range_Events(grpPO_DATE_ETA1)

        For i As Integer = 1 To 8
            ALLOC_BDATE.Add(DateSerial(2099, 1, 1))
            ALLOC_EDATE.Add(DateSerial(2099, 1, 1))
            ALLOC_SHOW.Add(False)
        Next
        Absx1.dteFor("ALLOC_BDATE1").DateTime = DateSerial(Now.Year, Now.Month, 1)
        Absx1.dteFor("ALLOC_EDATE1").DateTime = DateAdd(DateInterval.Month, 1, DateSerial(Now.Year, Now.Month, 1)).AddDays(-1)
        ALLOC_BDATE(0) = Absx1.dteFor("ALLOC_BDATE1").DateTime
        ALLOC_EDATE(0) = Absx1.dteFor("ALLOC_EDATE1").DateTime

        ALLOC_SHOW(0) = True

        Set_cmbYP("RYP0", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -48, 0, 0)
        Set_cmbYP("RYP1", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -48, 0, 0)

        Set_cmbYP("RYP2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -48, 0, 0)
        Set_cmbYP("RYP3", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -48, 0, 0)

        setShowDates(2)
        setALLOC_DATE1()
        FormReady = True
    End Sub

    Protected Overrides Sub Build_Workfile()
        xRYP0_legend = Absx1.cmbFor("RYP0").Value
        xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)

        xRYP1_legend = Absx1.cmbFor("RYP1").Value
        xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)

        xRYP2_legend = Absx1.cmbFor("RYP2").Value
        xRYP2 = Mid(xRYP2_legend, 1, 4) & Mid(xRYP2_legend, 6, 2)

        xRYP3_legend = Absx1.cmbFor("RYP3").Value
        xRYP3 = Mid(xRYP3_legend, 1, 4) & Mid(xRYP3_legend, 6, 2)

        If xRYP0 > xRYP1 Then
            MessageBox.Show("Start Period is greater than Ending Period.", "Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            RWU &= "0"
            xErrMsg = "No Eligible Records"
            Exit Sub
        End If

        If xRYP2 > xRYP3 Then
            MessageBox.Show("Start Period is greater than Ending Period.", "Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            RWU &= "0"
            xErrMsg = "No Eligible Records"
            Exit Sub
        End If

        Dim S As New StringBuilder With {.Length = 0}

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = ""
        'If Absx1.optFor("OPTASN").Value = "S" Then
        '    SUBT &= "Stock Styles Only"
        'ElseIf Absx1.optFor("OPTASN").Value = "N" Then
        '    SUBT &= "Non-Stock Styles Only"
        'End If

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

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

        '-- Shit you may need here --
        'sql_SELECT_cols, sql_TABLE_NAMEs, sql_WHERE, sql_JOIN, sql_filter, sql_filter2
        S.Length = 0
        S.AppendLine("Select " & sql_SELECT_cols)
        S.AppendLine(",ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTYC1.COLOR_CODE,")
        S.AppendLine("ICTSTYL1.STYLE_DESC,")
        S.AppendLine("0.00 AS ALLOC_1,")
        S.AppendLine("0.00 AS ALLOC_2,")
        S.AppendLine("0.00 AS ALLOC_3,")
        S.AppendLine("0.00 AS ALLOC_4,")
        S.AppendLine("0.00 AS ALLOC_5,")
        S.AppendLine("0.00 as ALLOC_6,")
        S.AppendLine("0.00 as ALLOC_7,")
        S.AppendLine("0.00 as ALLOC_8,")
        S.AppendLine("0.00 as ONP,")
        S.AppendLine("0.00 as SHIPPED,")
        S.AppendLine("0.00 as SHIPPED2")
        S.AppendLine("FROM ICTSTYL1, ICTSTYC1")
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE")
        S.AppendLine(sql_WHERE)
        S.AppendLine(sql_filter2)

        ASCMAIN1.sql = S.ToString()

        S.Length = 0
        S.AppendLine("Insert into " & ASTSRPT1)
        S.AppendLine(" (" & G1thru9)
        S.AppendLine(",STYLE_CODE,")
        S.AppendLine("COLOR_CODE,")
        S.AppendLine("STYLE_DESC,")
        S.AppendLine("ALLOC_1,")
        S.AppendLine("ALLOC_2,")
        S.AppendLine("ALLOC_3,")
        S.AppendLine("ALLOC_4,")
        S.AppendLine("ALLOC_5,")
        S.AppendLine("ALLOC_6,")
        S.AppendLine("ALLOC_7,")
        S.AppendLine("ALLOC_8,")
        S.AppendLine("ONP,")
        S.AppendLine("SHIPPED,")
        S.AppendLine("SHIPPED2")
        S.AppendLine(") ")
        S.AppendLine(" (" & ASCMAIN1.sql & ")")
        ASCDATA1.ExecuteSQL(S.ToString())

        With dst
            SOTSUPP1 = ASCMAIN1.Temp_Table("Select * from SOTSUPP1")
            ASCMAIN1.sql = "Select * from " & SOTSUPP1
            Create_TDA(.Tables.Add, "SOTSUPP1", "**", 0, False)

            SOTDEMD1 = ASCMAIN1.Temp_Table("Select * from SOTDEMD1")
            ASCMAIN1.sql = "Select * from " & SOTDEMD1
            Create_TDA(.Tables.Add, "SOTDEMD1", "**", 0, False)

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("ICTSTYL1.STYLE_CODE,")
            S.AppendLine("ICTSTAT2.COLOR_CODE,")
            S.AppendLine("SUM(NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) AS NETPOS")
            S.AppendLine("FROM ICTSTAT2, ICTSTYL1")
            S.AppendLine("WHERE ICTSTAT2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            S.AppendLine("GROUP BY ICTSTYL1.STYLE_CODE,")
            S.AppendLine("ICTSTAT2.COLOR_CODE,")
            S.AppendLine("ICTSTYL1.SALES_DIVISION_CODE,")
            S.AppendLine("ICTSTYL1.SUB_BODY_CODE")
            S.AppendLine("HAVING SUM(NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) <> 0")
            S.AppendLine("ORDER BY ICTSTYL1.STYLE_CODE,")
            S.AppendLine("ICTSTAT2.COLOR_CODE,")
            S.AppendLine("ICTSTYL1.SALES_DIVISION_CODE,")
            S.AppendLine("ICTSTYL1.SUB_BODY_CODE")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "SOTNETPS", "**", 0, False)
            Fill_Records("SOTNETPS")
        End With

        edi850cust = TAC.SOCMAIN1.Get_EDI_Custs("850")

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        'For i As Integer = 1 To 4
        '    dst.Tables.Item("ASTSRPT1").Columns.Add("SHIP_OPN" & Format(i, "0"), GetType(System.Double))
        'Next

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

        TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me,
          "",
          False,
          True,
          False,
          "", Now.Date, "")

        Dim newStyle As Boolean = True
        Dim lastStyle As String = ""
        Dim Zeros As Double()
        ReDim Zeros(8)
        For i As Integer = 0 To 8
            Zeros(i) = 0
        Next
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowASTSRPT1.Item("STYLE_CODE").ToString()
            Dim COLOR_CODE As String = rowASTSRPT1.Item("COLOR_CODE").ToString()
            If STYLE_CODE = lastStyle Then
                newStyle = False
            Else
                newStyle = True
                ASCMAIN1.Progress("Now Allocating Style ", STYLE_CODE)
            End If
            lastStyle = STYLE_CODE

            'If STYLE_CODE = "VCO31022" And COLOR_CODE = "001" Then Stop
            Dim totAlloc As Int64 = 0
            If StyleShouldAllocate(STYLE_CODE) Then
                Dim Allocations As Double() = GetAllocations(STYLE_CODE, COLOR_CODE, TABLE_NAMEs, newStyle)
                For i As Integer = 1 To 8
                    rowASTSRPT1.Item("ALLOC_" & i) = Allocations(i - 1)
                    totAlloc += Allocations(i - 1)
                Next
            Else
                For i As Integer = 1 To 8
                    rowASTSRPT1.Item("ALLOC_" & i) = Zeros(i - 1)
                    totAlloc += Zeros(i - 1)
                Next
            End If

            'This was removed per jimmie.  Data falling outside of the range selected will now be excluded.
            'And I am glad.  I never liked this stupid option anyway. : WR-5/31/17
            'If totAlloc = 0 Then
            '    If chkALLOC_DATE1.Checked Then
            '        Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
            '        If dst.Tables.Item("SOTNETPS").Select(filter).Count() = 1 Then
            '            Dim netPOS As Int64 = Val(dst.Tables.Item("SOTNETPS").Select(filter).FirstOrDefault().Item("NETPOS").ToString() & "")
            '            If netPOS <> 0 Then
            '                rowASTSRPT1.Item("ALLOC_1") = netPOS
            '            End If
            '        End If
            '    End If
            'End If
            rowASTSRPT1.Item("ONP") = getONP(STYLE_CODE, COLOR_CODE)
            If chkShowShipped.Checked Then
                rowASTSRPT1.Item("SHIPPED") = getSHIPPED(STYLE_CODE, COLOR_CODE, 1)
            End If
            If chkShowShipped2.Checked Then
                rowASTSRPT1.Item("SHIPPED2") = getSHIPPED(STYLE_CODE, COLOR_CODE, 2)
            End If

            If chkEXCLUDENEGATIVES.Checked Then
                For i As Integer = 1 To 8
                    If Val(rowASTSRPT1.Item(String.Format("ALLOC_{0}", i)).ToString & String.Empty) < 0 Then
                        rowASTSRPT1.Item(String.Format("ALLOC_{0}", i)) = 0
                    End If
                Next
            End If

        Next

    End Sub

    Public Overrides Sub Print_Report()
        Dim Reports As String() = {"SORAVAL3", "SORAVALP"}
        For Each RPT As String In Reports
            'RPT = "SORAVAL3"
            'CR_params.Add("GroupTitle1", F1 & "-" & Format(Absx1.dteFor("PO_DATE_ETA_L1").Value, "MM/dd/yy"))
            For i As Integer = 1 To 8
                If i > 1 Then
                    Dim LAST_ON_FILE As Boolean = Absx1.chkFor("CHKALLOC_DATE" & i).Checked
                    If LAST_ON_FILE Then
                        CR_params.Add("ALLOC_DATE" & i, Format(ALLOC_BDATE(i - 1), "MM/dd/yy") & " - Last")
                    Else
                        CR_params.Add("ALLOC_DATE" & i, Format(ALLOC_BDATE(i - 1), "MM/dd/yy") & " - " & Format(ALLOC_EDATE(i - 1), "MM/dd/yy"))
                    End If
                Else
                    If Absx1.chkFor("CHKALLOC_DATE" & i).Checked Then
                        CR_params.Add("ALLOC_DATE" & i, "First - " & Format(ALLOC_EDATE(i - 1), "MM/dd/yy"))
                    Else
                        CR_params.Add("ALLOC_DATE" & i, Format(ALLOC_BDATE(i - 1), "MM/dd/yy") & " - " & Format(ALLOC_EDATE(i - 1), "MM/dd/yy"))
                    End If
                End If

                Dim SHOW_ALLOC As String = "0"
                If ALLOC_SHOW(i - 1) Then
                    SHOW_ALLOC = "1"
                End If
                CR_params.Add("SHOW_ALLOC" & i, SHOW_ALLOC)
            Next
            If CHKALLOP.Checked Then
                CR_params.Add("SHOWALLORDPICK", "1")
            Else
                CR_params.Add("SHOWALLORDPICK", "")
            End If


            If chkShowShipped.Checked Then
                CR_params.Add("SHOWSHIPPED", "1")
                CR_params.Add("SHIP1DATES", xRYP0.Substring(4, 2) & "/" & xRYP0.Substring(2, 2) & " - " & xRYP1.Substring(4, 2) & "/" & xRYP1.Substring(2, 2))
            Else
                CR_params.Add("SHOWSHIPPED", "0")
                CR_params.Add("SHIP1DATES", "")
            End If

            If chkShowShipped2.Checked Then
                CR_params.Add("SHOWSHIPPED2", "1")
                CR_params.Add("SHIP2DATES", xRYP2.Substring(4, 2) & "/" & xRYP2.Substring(2, 2) & " - " & xRYP3.Substring(4, 2) & "/" & xRYP3.Substring(2, 2))
            Else
                CR_params.Add("SHOWSHIPPED2", "0")
                CR_params.Add("SHIP2DATES", "")
            End If


            If chkShowONP.Checked Then
                CR_params.Add("SHOWONP", "1")
            Else
                CR_params.Add("SHOWONP", "0")
            End If

            If chkShowTotal.Checked Then
                CR_params.Add("SHOWTOTAL", "1")
            Else
                CR_params.Add("SHOWTOTAL", "0")
            End If

            If ckhSHOWSTYLES.Checked Then
                CR_params.Add("SHOWSTYLES", "1")
            Else
                CR_params.Add("SHOWSTYLES", "0")
            End If

            If chkShowShipped.Checked Then
                SUBT = SUBT + " Shipped From " & xRYP0_legend & " To " & xRYP1_legend
            End If
            If chkShowShipped.Checked Then
                SUBT = SUBT + " And " & xRYP2_legend & " To " & xRYP3_legend
            End If

            'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
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
            'End If

            CR_params.Add("SUBT", txtDescription.Text & SUBT)

            If RPT = "SORAVALP" Then
                If chkProvideProof.Checked Then
                    Generate_Report(RPT, "Available To Sell By Period - Detailed Proof", SUBT)
                End If
            Else
                Generate_Report(RPT, , SUBT)
            End If
        Next
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            setFinalDates()
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
            '    EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            'End If
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length > 4 Then
                EMsg &= vbCr & "Maximum number of Sort Fields for this report is 4"
            End If
        End If
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
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
#End Region

#Region "Form Controls"
#Region "chkALLGROUP"
    Private Sub chkALLGROUPX_CheckedChanged(sender As Object, e As EventArgs) Handles chkALLGROUP2.CheckedChanged,
        chkALLGROUP3.CheckedChanged,
        chkALLGROUP4.CheckedChanged,
        chkALLGROUP5.CheckedChanged,
        chkALLGROUP6.CheckedChanged,
        chkALLGROUP7.CheckedChanged,
        chkALLGROUP8.CheckedChanged
        Dim cb As ABSCS.ABSCheckBox = TryCast(sender, ABSCS.ABSCheckBox)
        setShowDates(cb.Name.Substring(11, 1))
    End Sub
#End Region

#Region "chkALLOC_DATE"
    Private Sub chkALLOC_DATE1_CheckedChanged(sender As Object, e As EventArgs) Handles chkALLOC_DATE1.CheckedChanged
        If FormReady Then
            setALLOC_DATE1()
        End If
    End Sub

    Private Sub chkALLOC_DATEX_CheckedChanged(sender As Object, e As EventArgs) Handles chkALLOC_DATE2.CheckedChanged,
        chkALLOC_DATE3.CheckedChanged,
        chkALLOC_DATE4.CheckedChanged,
        chkALLOC_DATE5.CheckedChanged,
        chkALLOC_DATE6.CheckedChanged,
        chkALLOC_DATE7.CheckedChanged,
        chkALLOC_DATE8.CheckedChanged
        Dim cb As ABSCS.ABSCheckBox = TryCast(sender, ABSCS.ABSCheckBox)
        setShowDates(cb.Name.Substring(13, 1))
    End Sub

#End Region
    Private Sub chkShowShipped_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowShipped.CheckedChanged
        grpPERIOD_RANGE.Visible = chkShowShipped.Checked
        chkShowShipped2.Visible = chkShowShipped.Checked
        chkShowShipped2.Checked = False
    End Sub

    Private Sub chkShowShipped2_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowShipped2.CheckedChanged
        grpPERIOD_RANG2.Visible = chkShowShipped2.Checked
    End Sub
#End Region

#Region "Custom Methods"
    Private Function GetAllocations(ByVal STYLE_CODE As String,
                                    ByVal COLOR_CODE As String,
                                    ByVal TABLE_NAMEs As Dictionary(Of String, String),
                                    ByVal NEWSTYLE As Boolean) As Double()
        Dim RetVal As Double()

        If NEWSTYLE Then
            Dim SOTORDR0 As String = TABLE_NAMEs("SOTORDR0")
            Dim SOTORDR1 As String = TABLE_NAMEs("SOTORDR1")
            Dim SOTORDR2 As String = TABLE_NAMEs("SOTORDR2")
            Dim SOTRSRV1 As String = TABLE_NAMEs("SOTRSRV1")
            Dim SOTRSRV2 As String = TABLE_NAMEs("SOTRSRV2")
            Dim ARTCUST1 As String = TABLE_NAMEs("ARTCUST1")

            For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR0", "ARTCUST1", "ICTSTDQ1", "SOTORDR2", "SOTRSRV1", "SOTRSRV2"}
                ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAMEs(TABLE_NAME))
            Next

            For Each sql As String In TABLE_NAMEs.Keys
                If sql.StartsWith("sql") Then
                    Dim sqlstmt As String = Replace(TABLE_NAMEs(sql), "'STYLE_CODE'", "'" & STYLE_CODE & "'")
                    ASCDATA1.ExecuteSQL(sqlstmt)
                End If
            Next

            dst.Tables("SOTSUPP0").Rows.Clear()
            dst.Tables("SOTSUPPI").Rows.Clear()
            dst.Tables("SOTORDR7").Rows.Clear()
            dst.Tables("ICTSTDQ1").Rows.Clear()
            dst.Tables("ICTSTDQ2").Rows.Clear()

            TAC.SOCMAIN1.Allocation(Me,
                False,
                True,
                 "",
                 "", edi850cust,
                SOTSUPP1, SOTDEMD1, TABLE_NAMEs, True, True, STYLE_CODE, , , , False)
        End If

        ReDim RetVal(7)
        'If STYLE_CODE = "VCO51035" Then Stop
        If dst.Tables.Item("ICTSTDQ1").Rows.Count > 0 Then
            For i As Integer = 0 To 7
                Dim thisBDate As Date = ALLOC_BDATE(i)
                Dim thisEDate As Date = ALLOC_EDATE(i)
                If Not ALLOC_SHOW(i) Then
                    thisEDate = CDate("01/01/2099")
                End If
                Dim filter As String = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' AND STATUS_DATE >= '" & thisBDate & "' AND STATUS_DATE <= '" & thisEDate & "'"
                If i = 0 And chkALLOC_DATE1.Checked Then
                    filter = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' AND STATUS_DATE <= '" & thisEDate & "'"
                End If
                For Each rowICTSTDQ1 As DataRow In dst.Tables("ICTSTDQ1").Select(filter, "STYLE_CODE, COLOR_CODE, STATUS_DATE")
                    If i = 0 Then
                        If Val(rowICTSTDQ1.Item("QTY_ATS_CUM") & "") <> 0 Then
                            RetVal(i) = Val(rowICTSTDQ1.Item("QTY_ATS_CUM") & "")
                        End If
                    Else
                        RetVal(i) += Val(rowICTSTDQ1.Item("QTY_ATS") & "")
                    End If
                    'RetVal(i) += Val(rowICTSTDQ1.Item("QTY_ATS_CUM") & "") 'Check with Walter.  Only get the last CUM?
                    'RetVal(i) = Val(rowICTSTDQ1.Item("QTY_ATS_CUM") & "")
                    'RetVal(i) += Val(rowICTSTDQ1.Item("QTY_ATS") & "")
                    'If Val(rowICTSTDQ1.Item("STATUS_QTY") & "") > 0 Then Stop
                    'RetVal(i) += Val(rowICTSTDQ1.Item("STATUS_QTY") & "")
                Next
            Next
        Else
            For i As Integer = 0 To 7
                RetVal(i) = 0
            Next
        End If
        Return RetVal
    End Function

    Private Function getONP(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As Int64
        Dim retVal As Int64 = 0
        Dim SQLS As New StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT SUM((NVL(WHSE_QTY_OPEN,0) + NVL(WHSE_QTY_PICK,0))) AS ONP")
        SQLS.AppendLine("FROM ICTSTAT2")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        retVal = Val(ASCDATA1.GetDataValue)
        Return retVal
        'If STYLE_CODE = "500752IZ" Then Stop
    End Function

    Private Function getSHIPPED(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal Period As Integer) As Int64
        Dim retVal As Int64 = 0
        Dim SQLS As New StringBuilder With {.Length = 0}
        Dim P1 As String = ""
        Dim P2 As String = ""
        Select Case Period
            Case 1
                P1 = xRYP0
                P2 = xRYP1
            Case 2
                P1 = xRYP2
                P2 = xRYP3
            Case Else
                Stop
        End Select

        SQLS.AppendLine("SELECT NVL(SUM(S2.ORDR_QTY_SHIP),0) as ORDR_QTY_SHIP")
        SQLS.AppendLine("FROM SOTINVH1 S1, SOTINVH2 S2")
        SQLS.AppendLine("WHERE S1.INV_NO = S2.INV_NO")
        SQLS.AppendLine("AND S1.INV_TYPE = S2.INV_TYPE")
        SQLS.AppendLine(String.Format("AND S2.ORDR_YYYYPP_UPDATED >= '{0}'", P1))
        SQLS.AppendLine(String.Format("AND S2.ORDR_YYYYPP_UPDATED <= '{0}'", P2))
        SQLS.AppendLine(String.Format("AND S2.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND S2.COLOR_CODE = '{0}'", COLOR_CODE))
        If Absx1.chkFor("CHKTRANSAMPLES").Checked Then
            SQLS.AppendLine("AND S1.CUST_CODE NOT IN ('TRANSFERS','SAMPLES')")
        End If
        ASCMAIN1.sql = SQLS.ToString()
        retVal = Val(ASCDATA1.GetDataValue)
        'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
        '    If STYLE_CODE = "1458JS" And COLOR_CODE = "331" Then Stop
        'End If
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

    Private Sub setALLOC_DATE1()
        If chkALLOC_DATE1.Checked Then
            Absx1.dteFor("ALLOC_BDATE1").ReadOnly = True
            Absx1.dteFor("ALLOC_BDATE1").Text = ""
            'ALLOC_BDATE(0) = DateSerial(1900, 1, 1)
        Else
            ALLOC_BDATE(0) = ALLOC_BDATE(0)
            Absx1.dteFor("ALLOC_BDATE1").ReadOnly = False
            Absx1.dteFor("ALLOC_BDATE1").Text = ALLOC_BDATE(0)
            Absx1.dteFor("ALLOC_EDATE1").Text = ALLOC_EDATE(0)
        End If
    End Sub

    Private Sub setFinalDates()
        If chkALLOC_DATE1.Checked Then
            ALLOC_BDATE(0) = DateSerial(1900, 1, 1)
        Else
            ALLOC_BDATE(0) = Absx1.dteFor("ALLOC_BDATE1").DateTime
        End If
        ALLOC_EDATE(0) = Absx1.dteFor("ALLOC_EDATE1").DateTime
        For i As Integer = 2 To 8
            If Absx1.chkFor("CHKALLGROUP" & i).Checked Then
                ALLOC_SHOW(i - 1) = Absx1.chkFor("CHKALLGROUP" & i).Checked
                ALLOC_BDATE(i - 1) = Absx1.dteFor("ALLOC_BDATE" & i).DateTime
                If Absx1.chkFor("CHKALLOC_DATE" & i).Checked Then
                    ALLOC_EDATE(i - 1) = DateSerial(2099, 1, 1)
                Else
                    ALLOC_EDATE(i - 1) = Absx1.dteFor("ALLOC_EDATE" & i).DateTime
                End If
            Else
                ALLOC_SHOW(i - 1) = Absx1.chkFor("CHKALLGROUP" & i).Checked
                ALLOC_BDATE(i - 1) = DateSerial(1900, 1, 1)
                ALLOC_EDATE(i - 1) = DateSerial(2099, 1, 1)
            End If
        Next
    End Sub

    Private Sub setShowDates(ByVal DatePeriod As Integer)
        If Absx1.chkFor("CHKALLGROUP" & DatePeriod).Checked Then
            Absx1.dteFor("ALLOC_BDATE" & DatePeriod).DateTime = Absx1.dteFor("ALLOC_EDATE" & DatePeriod - 1).DateTime.AddDays(1)
            Dim MM As Integer = Absx1.dteFor("ALLOC_BDATE" & DatePeriod).DateTime.Month
            Dim YY As Integer = Absx1.dteFor("ALLOC_BDATE" & DatePeriod).DateTime.Year
            If Absx1.chkFor("CHKALLOC_DATE" & DatePeriod).Checked Then
                Absx1.dteFor("ALLOC_EDATE" & DatePeriod).ReadOnly = True
                Absx1.dteFor("ALLOC_EDATE" & DatePeriod).Text = ""
            Else
                Absx1.dteFor("ALLOC_EDATE" & DatePeriod).ReadOnly = False
                Absx1.dteFor("ALLOC_EDATE" & DatePeriod).DateTime = DateSerial(YY, MM, 1).AddMonths(1).AddDays(-1)
            End If
            Absx1.CtlFor("GRPALLGROUP" & DatePeriod).Visible = True
            Absx1.chkFor("CHKALLGROUP" & DatePeriod).Visible = True
            For I As Integer = DatePeriod + 1 To 8
                Absx1.CtlFor("GRPALLGROUP" & I).Visible = False
                Absx1.chkFor("CHKALLGROUP" & I).Visible = False
                Absx1.chkFor("CHKALLGROUP" & I).Checked = False
                Absx1.chkFor("CHKALLOC_DATE" & I).Checked = False
            Next
            If DatePeriod < 8 Then
                'Absx1.chkFor("CHKALLGROUP" & DatePeriod + 1).Checked = True
                Absx1.chkFor("CHKALLGROUP" & DatePeriod + 1).Visible = True
            End If
        Else
            For I As Integer = DatePeriod To 8
                Absx1.CtlFor("GRPALLGROUP" & I).Visible = False
                Absx1.chkFor("CHKALLGROUP" & I).Visible = False
                If I + 1 < 8 Then
                    Absx1.chkFor("CHKALLGROUP" & I + 1).Checked = False
                End If
            Next
            Absx1.chkFor("CHKALLGROUP" & DatePeriod).Visible = True
        End If
    End Sub

    Private Function StyleShouldAllocate(ByVal STYLE_CODE As String) As Boolean
        Dim retVal As Boolean = True
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT SUM(NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)+NVL(WHSE_QTY_OPEN,0)+NVL(WHSE_QTY_PICK,0)) TOT")
        SQLS.AppendLine("FROM ICTSTAT2")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim TOT As Int64 = Val(ASCDATA1.GetDataValue)
        If TOT = 0 Then
            retVal = False
        End If
        Return retVal
    End Function
#End Region
End Class