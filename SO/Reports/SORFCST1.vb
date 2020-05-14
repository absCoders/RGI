Imports System.Text

Public Class SORFCST1
    Dim SOTDEMD1 As String
    Dim SOTSUPP1 As String
    Dim edi850cust As List(Of String)


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"
        'Range_Events(grpPO_DATE_ETA1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim S As New StringBuilder With {.Length = 0}

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = ""
        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")
        Dim sql_TABLE_NAMEs_orig As String = sql_TABLE_NAMEs
        Dim sql_JOIN_orig As String = sql_JOIN

        Dim sql_filter2 As String = ""

        '-- Shit you may need here --
        'sql_SELECT_cols, sql_TABLE_NAMEs, sql_WHERE, sql_JOIN, sql_filter, sql_filter2
        S.Length = 0
        S.AppendLine("Select " & sql_SELECT_cols)
        S.AppendLine(",ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTSTYL1.STYLE_DESC,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) AS WHSE_QTY_ON_ORDER,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_OPEN,0)) AS WHSE_QTY_OPEN,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_PICK,0)) AS WHSE_QTY_PICK,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0)) AS WHSE_QTY_TRAN,")
        S.AppendLine("SUM(NVL(SOTRSRV2.RSRV_QTY_OPEN,0)) AS RSRV_QTY_OPEN")
        S.AppendLine("FROM ICTSTYL1, ICTSTAT2, SOTRSRV2")
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
        S.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTRSRV2.STYLE_CODE (+)")
        S.AppendLine(sql_WHERE)
        S.AppendLine(sql_filter2)
        S.AppendLine("GROUP BY")
        If sql_GROUP_BY_cols.Length > 0 Then
            S.AppendLine(sql_GROUP_BY_cols & ",")
        End If
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTSTYL1.STYLE_DESC")
        ASCMAIN1.sql = S.ToString()

        S.Length = 0
        S.AppendLine("Insert into " & ASTSRPT1)
        S.AppendLine(" (" & G1thru9)
        S.AppendLine(",STYLE_CODE,")
        S.AppendLine("COLOR_CODE,")
        S.AppendLine("STYLE_DESC,")
        S.AppendLine("WHSE_QTY_ON_HAND,")
        S.AppendLine("WHSE_QTY_ON_ORDER,")
        S.AppendLine("WHSE_QTY_OPEN,")
        S.AppendLine("WHSE_QTY_PICK,")
        S.AppendLine("WHSE_QTY_TRAN,")
        S.AppendLine("RSRV_QTY_OPEN")
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

            'S.Length = 0
            'S.AppendLine("SELECT")
            'S.AppendLine("ICTSTYL1.STYLE_CODE,")
            'S.AppendLine("ICTSTAT2.COLOR_CODE,")
            'S.AppendLine("SUM(NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) AS NETPOS")
            'S.AppendLine("FROM ICTSTAT2, ICTSTYL1")
            'S.AppendLine("WHERE ICTSTAT2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            'S.AppendLine("GROUP BY ICTSTYL1.STYLE_CODE,")
            'S.AppendLine("ICTSTAT2.COLOR_CODE,")
            'S.AppendLine("ICTSTYL1.SALES_DIVISION_CODE,")
            'S.AppendLine("ICTSTYL1.SUB_BODY_CODE")
            'S.AppendLine("HAVING SUM(NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) <> 0")
            'S.AppendLine("ORDER BY ICTSTYL1.STYLE_CODE,")
            'S.AppendLine("ICTSTAT2.COLOR_CODE,")
            'S.AppendLine("ICTSTYL1.SALES_DIVISION_CODE,")
            'S.AppendLine("ICTSTYL1.SUB_BODY_CODE")
            'ASCMAIN1.sql = S.ToString()
            'Create_TDA(.Tables.Add, "SOTNETPS", "**", 0, False)
            'Fill_Records("SOTNETPS")
        End With

        edi850cust = TAC.SOCMAIN1.Get_EDI_Custs("850")

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        For i As Integer = 1 To 3
            Dim iFormat As String = Format(i, "00")
            dst.Tables.Item("ASTSRPT1").Columns.Add("ALLOC_DATE_" & iFormat, GetType(System.DateTime))
            dst.Tables.Item("ASTSRPT1").Columns.Add("ALLOC_QTY_" & iFormat, GetType(System.Double))
        Next

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

        TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me, _
          "", _
          False, _
          True,
          False, _
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

            'If STYLE_CODE = "66114WM" Then Stop
            If newStyle Then
                Dim totAlloc As Int64 = 0
                If StyleShouldAllocate(STYLE_CODE) Then
                    Dim Allocations As Boolean = MakeAllocationTable(STYLE_CODE, TABLE_NAMEs, newStyle)
                    Dim DQFilter As String = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
                    Dim iCnt As Integer = 0
                    For Each rowICTSTDQ1 As DataRow In dst.Tables("ICTSTDQ1").Select(DQFilter, "STATUS_DATE")
                        iCnt += 1
                        If iCnt <= 3 Then
                            Dim iFormat As String = Format(iCnt, "00")
                            If IsDate(rowICTSTDQ1.Item("STATUS_DATE").ToString) Then
                                rowASTSRPT1.Item("ALLOC_DATE_" & iFormat) = rowICTSTDQ1.Item("STATUS_DATE").ToString
                            End If
                            rowASTSRPT1.Item("ALLOC_QTY_" & iFormat) = Val(rowICTSTDQ1.Item("QTY_ATS_CUM").ToString)
                        End If
                    Next
                End If
            End If
        Next

    End Sub

    Public Overrides Sub Print_Report()
        'CalculatePeriods()

        'CR_params.Add("SUBT", txtDescription.Text & SUBT)
        RPT = "SORFCST1"

        CR_params.Add("SUBT", txtDescription.Text & SUBT)

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length > 4 Then
            '    EMsg &= vbCr & "Maximum number of Sort Fields for this report is 4"
            'End If
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

    Private Function MakeAllocationTable(ByVal STYLE_CODE As String,
                                    ByVal TABLE_NAMEs As Dictionary(Of String, String),
                                    ByVal NEWSTYLE As Boolean) As Boolean
        Dim RetVal As Boolean = True

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
        Return RetVal
    End Function

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
End Class