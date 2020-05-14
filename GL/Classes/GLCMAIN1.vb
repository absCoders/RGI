Imports System.Math

Public Class GLCMAIN1

    Public Shared Function Prepare_Work_File( _
    ByRef F As ASFBASE1, _
    ByRef TTA As String, _
    ByRef TTB As String, _
    ByRef GLTFINRD As String, _
    ByRef RY As String, _
    ByRef P As Integer, _
    ByRef SQLP() As String, _
    ByRef SQLF() As String, _
    ByRef sqlA_select As String, _
    ByRef sqlA_group_by As String, _
    ByRef sqlA_where As String, _
    ByRef BY_SEG2 As Boolean, _
    ByRef BY_SEG3 As Boolean, _
    ByRef BY_SEG4 As Boolean) As String

        Dim ACCT_TYPEs As New Dictionary(Of String, String)

        ACCT_TYPEs.Add("B", "('A','L','E')")
        ACCT_TYPEs.Add("I", "('I','X')")



        TTA = GL_Prep(F, Format(Val(RY) - 5, "0000"), Format(Val(RY) + 1, "0000"))
        TTB = GL_Prep(F, Format(Val(RY) - 5, "0000"), Format(Val(RY) + 1, "0000"), True)

        ' Clean out data which represents actuals beyond the Report Period Selected

        ASCDATA1.ExecuteSQL("Delete from " & TTA & " where ACCT_YEAR > '" & RY & "'")
        If P <> 12 Then
            Dim sql_clear As String = ""
            For i As Integer = P + 1 To 12
                sql_clear &= ", ACCT_ACT_P" & Format(i, "00") & " = 0"
            Next
            ASCDATA1.ExecuteSQL("Update " & TTA & " Set " & Mid(sql_clear, 2) & " where ACCT_YEAR = '" & RY & "'")
        End If

        Call Setup_SQL(F, SQLP, SQLF, TTA, TTB, sqlA_select, sqlA_group_by, sqlA_where, BY_SEG2, BY_SEG3, BY_SEG4)

        Dim sql As String = ""

        Dim A234 As String = "ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE"

        Dim A234T As String = A234 & ", ACCT_TYPE"
        Dim sqlC As String = _
            "Select Distinct " & A234T & " from " & TTA & " union " & _
            "Select Distinct " & A234T & " from " & TTB

        sql = "Select Distinct " & A234T & " from (" & sqlC & ")"
        Dim TTC As String = ASCMAIN1.Temp_Table(sql)

        ASCMAIN1.sql = "Select * from GLTFINR3 where ROWNUM < 1"
        GLTFINRD = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        Dim STMT_LINE_NO_ALL_ELSE As Integer = 0
        Dim STMT_LINE_NO As Integer = 0
        For Each row As DataRow In F.dst.Tables("GLTFINR2").Select("STMT_LINE_TYPE = 'D'", "STMT_LINE_NO", DataViewRowState.CurrentRows)
            STMT_LINE_NO = Val(row.Item("STMT_LINE_NO"))
            sql = ""
            Dim sqlx As String = "Select Distinct '" & F.HFs("STMT_CODE") & "' STMT_CODE, " & CStr(STMT_LINE_NO) & " STMT_LINE_NO, " & A234
            Dim sqlfrom As String = " from " & TTC

            Select Case row.Item("STMT_LINE_ACCTS") & ""
                Case "S"
                    Dim dvwGLTFINR3 As New DataView(F.dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
                    Dim z As String = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        z &= ",'" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                    Next
                    If z <> "" Then
                        sql = sqlx & sqlfrom & " where ACCT_CODE in (" & Mid(z, 2) & ")"
                        For s As Integer = 2 To 4
                            If row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "S" _
                            Or row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "X" Then
                                Dim dvwGLTFINR4 As New DataView(F.dst.Tables("GLTFINR4"))
                                dvwGLTFINR4.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO) & " and ACCT_SEG_ID = '" & CStr(s) & "'"
                                z = ""
                                For i As Integer = 0 To dvwGLTFINR4.Count - 1
                                    z &= ",'" & dvwGLTFINR4(i).Item("ACCT_SEG_CODE") & "'"
                                Next
                                If z <> "" Then
                                    sql &= " and SEG" & CStr(s) & "_CODE" _
                                        & IIf(row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "X", " NOT", "") _
                                        & " in (" & Mid(z, 2) & ")"
                                End If
                            End If
                        Next
                    End If

                Case "R"
                    sql = sqlx & sqlfrom & " where ACCT_CODE >= '" & row.Item("STMT_LINE_ACCT_RANGE1") & "' and ACCT_CODE <= '" & row.Item("STMT_LINE_ACCT_RANGE2") & "'"

                Case "I"
                    If F.HFs("STMT_TYPE") = "B" Then
                        sql = sqlx & sqlfrom & " where ACCT_TYPE in " & ACCT_TYPEs("I")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "B"
                    If F.HFs("STMT_TYPE") = "I" Then
                        sql = sqlx & sqlfrom & " where ACCT_TYPE in " & ACCT_TYPEs("B")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "X"
                    Dim dvwGLTFINR3 As New DataView(F.dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
                    Dim z As String = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        z &= ",('" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG2_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG3_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG4_CODE") & "')"
                    Next
                    If z <> "" Then
                        sql = sqlx & sqlfrom & " where (" & A234 & ") in (" & Mid(z, 2) & ")"
                    End If
            End Select

            If sql <> "" Then
                sql = "Insert into " & GLTFINRD & " " & sql
                ASCDATA1.ExecuteSQL(sql)
            End If
        Next

        If STMT_LINE_NO_ALL_ELSE <> 0 Then
            sql = "Select '" & F.HFs("STMT_CODE") & "' STMT_CODE, " & CStr(STMT_LINE_NO_ALL_ELSE) & " STMT_LINE_NO, " & A234 & " " _
                    & " from (" _
                    & "Select DISTINCT TT." & Replace(A234, ", ", ", TT.") & " from GLTACCT1," & TTC & " TT where TT.ACCT_CODE = GLTACCT1.ACCT_CODE and GLTACCT1.ACCT_TYPE in " & ACCT_TYPEs(F.HFs("STMT_TYPE")) _
                    & " MINUS " _
                    & "Select DISTINCT " & A234 & " from " & GLTFINRD _
                    & ")"
            sql = "Insert into " & GLTFINRD & " " & sql
            ASCDATA1.ExecuteSQL(sql)
        End If

        Call ASCMAIN1.AnalyzeTable(GLTFINRD)

        Return TTC

    End Function

    Public Shared Sub Setup_SQL( _
    ByRef F As ASFBASE1, _
    ByRef SQLP() As String, _
    ByRef SQLF() As String, _
    ByRef TTA As String, _
    ByRef TTB As String, _
    ByRef sqlA_select As String, _
    ByRef sqlA_group_by As String, _
    ByRef sqlA_where As String, _
    ByRef BY_SEG2 As Boolean, _
    ByRef BY_SEG3 As Boolean, _
    ByRef BY_SEG4 As Boolean)

        ReDim SQLP(4)
        ReDim SQLF(4)
        SQLP(0) = ", Sum (ACCT_BEG_BAL)"
        SQLP(1) = ", Sum (ACCT_BEG_BAL)"
        SQLP(4) = ", Sum (ACCT_BEG_BAL)"
        For i As Integer = 1 To 13
            SQLP(0) = SQLP(0) & ", Sum (ACCT_ACT_P" & Format$(i, "00") & ")"
            SQLP(1) = SQLP(1) & ", Sum (ACCT_BUD_P" & Format$(i, "00") & ")"
            SQLP(4) = SQLP(4) & ", Sum (ACCT_BUD_P" & Format$(i, "00") & ")"
        Next i
        SQLF(0) = " from " & TTA & " X"
        SQLF(1) = " from " & TTB & " X"
        SQLF(4) = " from " & TTB & " X"
        ' NEED AN ORIGINAL BUDGET TEMP TABLE

        sqlA_select = ""
        Dim z As String
        For i As Integer = 2 To 4
            Dim SEGX_CODE As String = "SEG" & CStr(i) & "_CODE"
            If Not New Boolean() {BY_SEG2, BY_SEG3, BY_SEG4}(i - 2) Then
                z = "'" & F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "'"
            Else
                z = "X." & SEGX_CODE
                sqlA_group_by = sqlA_group_by & ", " & z
            End If
            sqlA_select = sqlA_select & ", " & z & " " & SEGX_CODE
        Next

        Dim sql As String = ""
        z = SQLA("ACCT_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.ACCT_CODE " & IIf(SQLA("ACCT_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("SEG2_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.SEG2_CODE " & IIf(SQLA("SEG2_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("SEG3_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.SEG3_CODE " & IIf(SQLA("SEG3_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("SEG4_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.SEG4_CODE " & IIf(SQLA("SEG4_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("ACCT_TYPE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.ACCT_TYPE " & IIf(SQLA("ACCT_TYPE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        sqlA_where = sql
    End Sub

    Public Shared Function GL_Prep( _
    ByRef F As ASFBASE1, _
    ByRef YYYY_beg As String, _
    ByRef YYYY_end As String, _
    Optional ByRef budget As Boolean = False, _
    Optional ByRef OFFSET As Integer = 0, _
    Optional ByRef OFFSET_Y As Integer = 0, _
    Optional ByRef budTblsfx24 As String = "", _
    Optional ByRef TABLE_NAME As String = "") As String

        ' determine YYYY_gyp as lesser of YYYY_beg and GYP
        ' get all years into work table from YYYY_gyp thru endyear
        ' change nulls to zeroes
        ' get balance forward set for years > YYYY_gyp thru YYYY_end
        ' close net profit into RTE for all years from GYP thru YYYY_end

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim z As String
        Dim sqlbs As String
        Dim sqlis As String
        Dim sql As String

        Dim GYP As String
        Dim RTE As String
        Dim RTEsql As String
        Dim RTEsql_group_by As String

        GYP = F.ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
        RTE = F.ROWs("GLTPARM1").Item("GL_PARM_RET_EARN_ACCT")

        RTEsql = ""
        RTEsql_group_by = ""
        Dim seg(4) As String
        seg(2) = F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        seg(3) = F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        seg(4) = F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

        Dim YYYY_gyp As String
        If YYYY_beg < Mid$(GYP, 1, 4) Then
            YYYY_gyp = YYYY_beg
        Else
            YYYY_gyp = Mid$(GYP, 1, 4)
        End If

        Dim YRS As String
        YRS = ""
        For i = Val(YYYY_gyp) To Val(YYYY_end)
            YRS = YRS & ",'" & Format$(i, "0000") & "'"
        Next i
        YRS = Mid$(YRS, 2)

        '' Make sure that all Segment Codes are accounted for in Segment Master File

        'For i = 2 To 4
        '    z = Format$(i, "0")
        '    sql = "INSERT INTO GLTSEGM1 (ACCT_SEG_ID, ACCT_SEG_CODE, ACCT_SEG_DESC)"
        '    sql = sql & " SELECT '" & z & "', SEG" & z & "_CODE, 'Code ' || SEG" & z & "_CODE "
        '    sql = sql & " FROM"
        '    sql = sql & " (SELECT DISTINCT SEG" & z & "_CODE FROM GLTACCT3 "
        '    sql = sql & " MINUS "
        '    sql = sql & "  SELECT ACCT_SEG_CODE FROM GLTSEGM1 WHERE ACCT_SEG_ID = '" & z & "')"
        '    OraD.ExecuteSQL(sql)
        'Next i

        sql = "Select GLTACCT3.*, GLTACCT1.ACCT_TYPE "
        If budget Then
            If budTblsfx24 = "" Then
                budTblsfx24 = "2"
            End If
            sql = sql & " from EMP.GLTACCT" & budTblsfx24 & " GLTACCT3,EMP.GLTACCT1"
        Else
            sql = sql & " from EMP.GLTACCT3,EMP.GLTACCT1"
        End If
        sql = sql & " where GLTACCT1.ACCT_CODE (+) = GLTACCT3.ACCT_CODE"
        sql = sql & "   and GLTACCT3.ACCT_YEAR in (" & YRS & ")"
        Dim TT As String = ""
        If TABLE_NAME <> "" Then
            TT = TABLE_NAME
            ASCDATA1.ExecuteSQL("Delete from " & TT)
            ASCDATA1.ExecuteSQL("Insert into " & TT & " " & sql)
        Else
            TT = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & TT & " add Primary Key (ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,ACCT_YEAR)")
        End If

        For i = 2 To 4
            z = "SEG" & Format$(i, "0")
            If F.ROWs("GLTPARM1").Item("GL_PARM_" & z & "_RTE") & "" = "1" Then
                RTEsql = RTEsql & z & "_CODE,"
                RTEsql_group_by = RTEsql_group_by & z & "_CODE,"
            Else
                RTEsql = RTEsql & "'" & seg(i) & "' " & z & "_CODE,"
                RTEsql_group_by = RTEsql_group_by & "'" & seg(i) & "',"
            End If
        Next i

        ASCDATA1.ExecuteSQL("Update " & TT & " set ACCT_BEG_BAL = 0 where ACCT_BEG_BAL is Null")

        sqlbs = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE, NVL(ACCT_BEG_BAL,0) "
        sqlis = "Select " & RTEsql & " Sum (NVL(ACCT_BEG_BAL,0) "
        For i = 1 To 13
            If budget Then
                z = "ACCT_BUD_P" & Format$(i, "00")
            Else
                z = "ACCT_ACT_P" & Format$(i, "00")
            End If
            ASCDATA1.ExecuteSQL("Update " & TT & " Set " & z & " = 0 where " & z & " is Null")
            sqlbs = sqlbs & " + NVL (" & z & ",0)"
            sqlis = sqlis & " + NVL (" & z & ",0)"
        Next i
        sqlis = sqlis & ")"
        sqlbs = sqlbs & " ACCT_BEG_BAL"

        If TABLE_NAME = "" Then
            Call F.Create_TDA(F.dst.Tables.Add, TT, "*")
            For j = 0 To 13
                F.dst.Tables(TT).Columns(5 + j).DefaultValue = 0
            Next
        Else
            '            Fill_Records(TT)
        End If

        Dim RTE_imax As Integer

        If Val(Mid$(GYP, 1, 4)) <= YYYY_end - 1 Then
            Dim yz As String
            For i = Val(Mid$(GYP, 1, 4)) To YYYY_end - 1
                yz = Format$(i + 1, "0000")
                sql = sqlbs & " from " & TT & " where ACCT_TYPE in ('A','L','E') and ACCT_YEAR = '" & Format$(i, "0000") & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable(sql, "GLTACCTX").Rows
                    Dim ACCT_BEG_BAL As Double = Val(row.Item("ACCT_BEG_BAL") & "")
                    If ACCT_BEG_BAL <> 0 Then
                        Dim rowTT As DataRow = F.Fill_Record(TT, New String() {row.Item("ACCT_CODE"), _
                        row.Item("SEG2_CODE"), row.Item("SEG3_CODE"), row.Item("SEG4_CODE"), yz}, True)
                        rowTT.Item("ACCT_TYPE") = row.Item("ACCT_TYPE")
                        rowTT.Item("ACCT_BEG_BAL") = Val(rowTT.Item("ACCT_BEG_BAL") & "") + ACCT_BEG_BAL
                        F.Update_Record_TDA(TT) ' F.Update_Record_TDA_Rows(TT)
                    End If
                Next
                RTE_imax = i
                Call RTE_Calc(F, i, YYYY_gyp, TT, RTEsql_group_by, sqlis, RTE_imax, RTE)
            Next i
        End If

        If OFFSET <> 0 Then
            Stop ' WHEN WE HAVE A FRESH MIND
            '    Dim jmax As Integer
            '    j = 0
            '    If budget Then
            '        z = "BUD"
            '    Else
            '        z = "ACT"
            '    End If
            '    sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE" & vbCr
            '    sql = sql & ", SUM (DECODE(ACCT_YEAR,'" & Format$(YYYY_gyp, "0000") & "',NVL(ACCT_BEG_BAL,0))) P000" & vbCr
            '    For i = Val(YYYY_gyp) To Val(YYYY_end)
            '        For k = 1 To 12
            '            j = j + 1
            '            sql = sql & ", SUM (DECODE(ACCT_YEAR,'" & Format$(i, "0000") & "', NVL(ACCT_" & z & "_P" & Format$(k, "00") & ",0))) P" & Format$(j, "000") & vbCr
            '        Next k
            '    Next i
            '    sql = sql & " from " & TT & " group by "
            '    sql = sql & "ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE" & vbCr
            '    jmax = j
            '    Dim dyn As OraDynaset
            '    dyn = OraD.CreateDynaset(sql, 8&)
            '    ASCDATA1.ExecuteSQL("Delete from " & TT)
            '    sql = "Select * from " & TT & " where ROWNUM < 1"
            '    Dim dyntt As OraDynaset
            '    dyntt = OraD.CreateDynaset(sql, 0&)
            '    Dim a As Double
            '    Dim AMT() As Double
            '    Do While Not dyn.EOF
            '        ReDim AMT(12)
            '        k = 0
            '        i = OFFSET_Y
            '        For j = 0 To jmax
            '            a = Val(dyn.Fields("P" & Format$(j, "000")).Value & "")
            '            If j <= OFFSET Then
            '                AMT(0) = AMT(0) + a
            '            Else
            '                k = k + 1
            '                AMT(k) = a
            '                If k = 12 Or j = jmax Then
            '                    If InStr("ALE", dyn.Fields("ACCT_TYPE").Value & "") = 0 Then
            '                        If i = OFFSET_Y Then
            '                            dyntt.AddNew()
            '                            dyntt.Fields("ACCT_CODE").Value = dyn.Fields("ACCT_CODE").Value
            '                            dyntt.Fields("SEG2_CODE").Value = dyn.Fields("SEG2_CODE").Value
            '                            dyntt.Fields("SEG3_CODE").Value = dyn.Fields("SEG3_CODE").Value
            '                            dyntt.Fields("SEG4_CODE").Value = dyn.Fields("SEG4_CODE").Value
            '                            dyntt.Fields("ACCT_YEAR").Value = "0000" ' Val(y0) ' + i - 1
            '                            dyntt.Fields("ACCT_TYPE").Value = dyn.Fields("ACCT_TYPE").Value & ""
            '                            dyntt.Fields("ACCT_BEG_BAL").Value = AMT(0)
            '                            dyntt.Update()
            '                        End If
            '                        AMT(0) = 0
            '                    End If
            '                    If Val(YYYY_gyp) + i <= Val(YYYY_end) Then ' And (amt(0) <> 0 Or amt(1) <> 0 Or amt(2) <> 0 Or amt(3) <> 0 Or amt(4) <> 0 Or amt(5) <> 0 Or amt(6) <> 0 Or amt(7) <> 0 Or amt(8) <> 0 Or amt(9) <> 0 Or amt(10) <> 0 Or amt(11) <> 0 Or amt(12) <> 0) Then
            '                        dyntt.AddNew()
            '                        dyntt.Fields("ACCT_CODE").Value = dyn.Fields("ACCT_CODE").Value
            '                        dyntt.Fields("SEG2_CODE").Value = dyn.Fields("SEG2_CODE").Value
            '                        dyntt.Fields("SEG3_CODE").Value = dyn.Fields("SEG3_CODE").Value
            '                        dyntt.Fields("SEG4_CODE").Value = dyn.Fields("SEG4_CODE").Value
            '                        dyntt.Fields("ACCT_YEAR").Value = Val(y0) + i
            '                        dyntt.Fields("ACCT_TYPE").Value = dyn.Fields("ACCT_TYPE").Value & ""
            '                        dyntt.Fields("ACCT_BEG_BAL").Value = AMT(0)
            '                        For k = 1 To 12
            '                            dyntt.Fields("ACCT_" & z & "_P" & Format$(k, "00")).Value = AMT(k)
            '                            If InStr("ALE", dyn.Fields("ACCT_TYPE").Value & "") <> 0 Then
            '                                AMT(0) = AMT(0) + AMT(k)
            '                            End If
            '                            AMT(k) = 0
            '                        Next k
            '                        dyntt.Update()
            '                    End If
            '                    i = i + 1
            '                    k = 0
            '                End If
            '            End If
            '        Next j
            '        dyn.MoveNext()
            '    Loop

            '    i = 0
            '    RTE_imax = Val(YYYY_end)
            'GoSub Calc_RTE

            '    For i = Val(YYYY_gyp) To Val(YYYY_end)
            '        RTE_imax = Val(YYYY_end)
            '    GoSub Calc_RTE
            '    Next i

            '    sql = "Update " & TT & " SET ACCT_BEG_BAL = 0 "
            '    sql = sql & " where ACCT_TYPE in ('I','X') "
            '    sql = sql & " and ACCT_YEAR = '0000'"
            '    ASCDATA1.ExecuteSQL(sql) ' Clear out Accum R/E from periods prior to start of re-calendarized year which was stuffed into Op Accts

        End If

        ASCDATA1.ExecuteSQL("Delete from " & TT & " where ACCT_YEAR < '" & Format$(Val(YYYY_beg) + OFFSET_Y * Sign(Abs(OFFSET)), "0000") & "'")
        ASCDATA1.ExecuteSQL("Delete from " & TT & " where ACCT_YEAR > '" & Format$(Val(YYYY_end), "0000") & "'")

        If budget Then
            z = "BUD"
        Else
            z = "ACT"
        End If
        sql = "Delete from " & TT
        sql = sql & " where NVL(ACCT_BEG_BAL,0) = 0" & vbCr
        For k = 1 To 12
            sql = sql & " and NVL(ACCT_" & z & "_P" & Format$(k, "00") & ",0) = 0" & vbCr
        Next k
        'OraD.ExecuteSQL sql ' this throws off the TBAL where an account may have had activity which nets to 0

        If TABLE_NAME = "" Then
            ASCDATA1.ExecuteSQL("Create Index I_" & TT & "_1 on " & TT & " (ACCT_YEAR,ACCT_TYPE)")
        End If

        Call ASCMAIN1.AnalyzeTable(TT)

        Return TT

    End Function

    Public Shared Function SQLA( _
    ByRef PB_COLUMN_NAME As String, _
    Optional ByRef COLUMN_NAME As String = "CODE_VALUES", _
    Optional ByRef SQL_List As Boolean = False) As String
        Dim rowASTDSQLA As DataRow ' = tblASTDSQLA.Rows.Find(PB_COLUMN_NAME)
        rowASTDSQLA = Nothing ' THIS IS THE ONLY THING THAT WE NEED TO MOVE THESE ROUTINES FROM GLRSTMT1/ASFBASE1 TO THIS MODULE
        If rowASTDSQLA Is Nothing Then
            SQLA = ""
        Else
            SQLA = rowASTDSQLA.Item(COLUMN_NAME) & ""
            If SQL_List And SQLA <> "" Then
                SQLA = "'" & Replace(SQLA, ",", "','") & "'"
            End If
        End If
        Return SQLA
    End Function


    Public Shared Sub RTE_Calc( _
    ByRef F As ASFBASE1, _
    ByRef YYYY As Integer, _
    ByRef YYYY_gyp As String, _
    ByRef TT As String, _
    ByRef RTEsql_group_by As String, _
    ByRef sqlis As String, _
    ByRef RTE_imax As Integer, _
    ByRef RTE As String)

        Dim RTE_imin As Integer
        If YYYY = 0 Then
            RTE_imin = YYYY_gyp
        Else
            RTE_imin = YYYY
        End If
        Dim sql As String
        sql = sqlis & " from " & TT & " where ACCT_TYPE in ('I','X') "
        sql = sql & " and ACCT_YEAR = '" & Format$(YYYY, "0000") & "'"
        sql = sql & " group by " & Mid$(RTEsql_group_by, 1, Len(RTEsql_group_by) - 1)
        For Each rowRTE As DataRow In ASCDATA1.GetDataTable(sql, "").Rows
            Dim ACCT_BEG_BAL As Double = Val(rowRTE.Item(3) & "")
            If ACCT_BEG_BAL <> 0 Then
                For RTE_i As Integer = RTE_imin To RTE_imax
                    Dim row As DataRow = F.Fill_Record(TT, New String() {RTE, _
                    rowRTE.Item("SEG2_CODE"), rowRTE.Item("SEG3_CODE"), rowRTE.Item("SEG4_CODE"), _
                    Format$(RTE_i + 1, "0000")}, True)
                    row.Item("ACCT_TYPE") = "E"
                    row.Item("ACCT_BEG_BAL") = Val(row.Item("ACCT_BEG_BAL") & "") + ACCT_BEG_BAL
                    F.Update_Record_TDA(TT) ' F.Update_Record_TDA_Rows(TT)
                Next RTE_i
            End If
        Next
    End Sub
End Class
