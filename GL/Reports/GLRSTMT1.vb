Imports System.Math

Public Class GLRSTMT1

    Dim MAX_COLUMNS As Integer = 100 ' 13 ' Maximum number of columns supported in a Financial Statement
    ' A NOTE ABOUT MAX_COLUMNS
    ' THE RD() ARRAY STUFFS THE DENOMINATOR FOR VARPCT COLUMNS INTO MAX_COLUMNS + 1, 2, 3, ETC
    ' SO THERE IS AN IMPLICIT MAX OF MAX_COLUMNS REAL COLUMNS AND MAX_COLUMNS MORE FOR RD'S IN CASE THEY ARE NEEDED

    Dim FH(,) As String
    Dim HDGs() As String

    Dim xy(,) As Integer
    Dim rd() As Single  ' report column divisors (for var pct columns)
    Dim ADJ_P As Integer = 0
    Dim ACCT_TYPEs As New Dictionary(Of String, String)
    Dim SHOW_000S As Integer
    Dim SQLP() As String
    Dim SQLF() As String
    'Dim SQLdtl() As String
    Dim rr(,) As Double
    Dim lineref(,) As Double
    Dim A(,,) As Double
    Dim FA(,,,) As Double
    Dim STMT_LINE_REF_SETs As String
    Dim GLTFINRD As String = ""
    Dim RY As String        ' YYYY portion of RYP
    Dim P As Integer        ' PP portion of RYP
    Dim BY_SEG2 As Boolean
    Dim BY_SEG3 As Boolean
    Dim BY_SEG4 As Boolean
    Dim BY_SEG2_CLASS As Boolean
    Dim BY_SEG3_CLASS As Boolean
    Dim BY_SEG4_CLASS As Boolean
    Dim sqlA_where As String
    Dim sqlA_select As String
    Dim sqlA_group_by As String

    Dim TTA As String
    Dim TTB As String
    Dim TTC As String

    Dim STMT_LINE_REFs As New Dictionary(Of String, Integer)
    Dim STFs As New Dictionary(Of Integer, String)

    Dim rowGLTCLAY1 As DataRow
    Dim STMT_CALC_NO_REF_PCT As String

    Dim STMT_LINE_NO_BEG As Int32 = 7
    Dim STMT_LINE_NO_END As Int32 = 9

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ACCT_TYPEs.Add("B", "('A','L','E')")
        ACCT_TYPEs.Add("I", "('I','X')")

        Get_PARM("GLTPARM1")
        Breakout_By()
        Breakout_By_Class()
        Set_cmbYP("RYP", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -60, 24, 0)

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Now Loading Data")

        Get_PARM("GLTPARM1")
        HFs.Add("STMT_CODE", Absx1.cmbFor("STMT_CODE").Text)
        HFs.Add("STMT_LAYOUT_CODE", Absx1.cmbFor("STMT_LAYOUT_CODE").Text)

        ASCMAIN1.sql = "Select * from GLTCLAY2 where STMT_LAYOUT_CODE = '" & HFs("STMT_LAYOUT_CODE") & "'"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "GLTCLAY2"))
        'MAX_COLUMNS = Val(dst.Tables("GLTCLAY2").Compute("MAX(STMT_COL_POS)", "") & "")

        RY = Mid(RYP, 1, 4)
        P = Val(Mid(RYP, 5, 2))
        ADJ_P = P

        Load_Layout()

        If Absx1.chkFor("SHOW_000S").Checked Then
            SHOW_000S = 1000
        Else
            SHOW_000S = 1
        End If

        For i As Integer = 1 To 4
            Dim TABLE_NAME As String = "GLTFINR" & CStr(i)
            Dim sql As String = "Select * from " & TABLE_NAME _
                & " where STMT_CODE = '" & HFs("STMT_CODE") & "'"
            dst.Tables.Add(ASCDATA1.GetDataTable(sql, TABLE_NAME))
        Next

        HFs("STMT_TYPE") = dst.Tables("GLTFINR1").Rows(0).Item("STMT_TYPE")
        BY_SEG2 = Absx1.chkFor("SEG2_CODE").Checked Or SQLA("SEG2_CODE", "SEQUENCE") <> ""
        BY_SEG3 = Absx1.chkFor("SEG3_CODE").Checked Or SQLA("SEG3_CODE", "SEQUENCE") <> ""
        BY_SEG4 = Absx1.chkFor("SEG4_CODE").Checked Or SQLA("SEG4_CODE", "SEQUENCE") <> ""
        BY_SEG2_CLASS = Absx1.chkFor("SEG2_CLASS_CODE").Checked Or SQLA("SEG2_CLASS_CODE", "SEQUENCE") <> ""
        BY_SEG3_CLASS = Absx1.chkFor("SEG3_CLASS_CODE").Checked Or SQLA("SEG3_CLASS_CODE", "SEQUENCE") <> ""
        BY_SEG4_CLASS = Absx1.chkFor("SEG4_CLASS_CODE").Checked Or SQLA("SEG4_CLASS_CODE", "SEQUENCE") <> ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = " and GLTFINR1.STMT_CODE = '" & HFs("STMT_CODE") & "'"

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")


        ' DON'T UNDERSTAND WHY WE NEED THE CODE BELOW
        ' AND WHY IT IS CODED ONLY FOR SEG4 IF WE DO NEED IT

        'Dim SQL_SEG As String = ""

        'SQL_SEG = "GLTACCT3.SEG4_CLASS_CODE"
        'If sql_SELECT_cols.Contains(SQL_SEG) Or sql_WHERE.Contains(SQL_SEG) Or sql_GROUP_BY_cols.Contains(SQL_SEG) Then
        '    If Not sql_TABLE_NAMEs.Contains(",GLTSEGM1 GLTSEGM4") Then
        '        sql_TABLE_NAMEs &= ",GLTSEGM1 GLTSEGM4"
        '        sql_JOIN = " and GLTSEGM4.ACCT_SEG_ID = '4' and GLTSEGM4.ACCT_SEG_CODE = GLTACCT3.SEG4_CODE"
        '        sql_SELECT_cols = Replace(sql_SELECT_cols, SQL_SEG, "GLTSEGM4.ACCT_SEG_CLASS")
        '        sql_WHERE = Replace(sql_WHERE, SQL_SEG, "GLTSEGM4.ACCT_SEG_CLASS")
        '        sql_GROUP_BY_cols = Replace(sql_GROUP_BY_cols, SQL_SEG, "GLTSEGM4.ACCT_SEG_CLASS")
        '    End If
        'End If

        'SQL_SEG = "GLTACCT3.SEG4_GROUP_CODE"
        'If sql_SELECT_cols.Contains(SQL_SEG) Or sql_WHERE.Contains(SQL_SEG) Or sql_GROUP_BY_cols.Contains(SQL_SEG) Then
        '    If Not sql_TABLE_NAMEs.Contains(",GLTSEGG2 GLTSEGG4") Then
        '        sql_TABLE_NAMEs &= ",GLTSEGG2 GLTSEGG4"
        '        sql_JOIN = " and GLTSEGG4.ACCT_SEG_ID = '4' and GLTSEGG4.ACCT_SEG_CODE = GLTACCT3.SEG4_CODE"
        '        sql_SELECT_cols = Replace(sql_SELECT_cols, SQL_SEG, "GLTSEGG4.ACCT_SEG_GROUP_CODE")
        '        sql_WHERE = Replace(sql_WHERE, SQL_SEG, "GLTSEGG4.ACCT_SEG_GROUP_CODE")
        '        sql_GROUP_BY_cols = Replace(sql_GROUP_BY_cols, SQL_SEG, "GLTSEGG4.ACCT_SEG_GROUP_CODE")
        '    End If
        'End If



        Dim TT_Work As String = Prepare_Work_File()

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", GLTFINR1.STMT_CODE" & vbCrLf _
            & ", 1 AS REPORT_NO" & vbCrLf
        sql = sql & " from GLTFINR1," & TT_Work & " GLTACCT3" & sql_TABLE_NAMEs & vbCrLf
        sql = sql & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf
        sql = sql & " group by " & sql_GROUP_BY_cols & vbCrLf
        sql = sql & IIf(sql_GROUP_BY_cols = "", "", ",") & " GLTFINR1.STMT_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        ' this report is a PB report only up to this point
        Build_Report_File()
        PB_Report = False

        sql = "Select TATWORK1.W_LNG REPORT_NO "
        For i As Integer = 2 To 4
            Dim z As String = "SEG" & CStr(i)
            sql = sql & ", GLTSEGM1.ACCT_SEG_CODE " & z & "_CLASS_CODE"
            sql = sql & ", GLTSEGM1.ACCT_SEG_DESC " & z & "_CLASS_DESC"
            sql = sql & ", GLTSEGG1.ACCT_SEG_GROUP_CODE " & z & "_GROUP_CODE"
            sql = sql & ", GLTSEGG1.ACCT_SEG_GROUP_DESC " & z & "_GROUP_DESC"
            sql = sql & ", GLTSEGM1.ACCT_SEG_CODE " & z & "_CODE"
            sql = sql & ", GLTSEGM1.ACCT_SEG_DESC " & z & "_DESC"
        Next
        sql = sql & ", '1' SUPPRESS_PAGE"
        sql = sql & " from TATWORK1,GLTSEGM1,GLTSEGG1 where ROWNUM < 1"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTSTMTX", 1))
        dst.Tables("GLTSTMTX").Columns("SUPPRESS_PAGE").ReadOnly = False

        sql = "Select GLTFINR2.STMT_CODE, TATWORK1.W_LNG REPORT_NO" & vbCrLf _
            & ", GLTFINR2.STMT_LINE_NO, GLTFINR2.STMT_LINE_NO STMT_LINE_NO2" & vbCrLf _
            & ", GLTFINR2.STMT_LINE_REF_PCT" & vbCrLf _
            & ", GLTFINR3.ACCT_CODE" & vbCrLf _
            & ", GLTFINR3.SEG2_CODE, GLTFINR3.SEG3_CODE, GLTFINR3.SEG4_CODE" & vbCrLf _
            & ", GLTFINR3.SEG2_CODE SEG2_CLASS_CODE, GLTFINR3.SEG3_CODE SEG3_CLASS_CODE, GLTFINR3.SEG4_CODE SEG4_CLASS_CODE" & vbCrLf _
            & ", 'X' SUPPRESS_PRINT" & vbCrLf
        For i As Integer = 1 To MAX_COLUMNS
            sql = sql & ", TATWORK1.W_AMT AMT" & Format$(i, "00") & vbCrLf
        Next i
        For i As Integer = 1 To MAX_COLUMNS
            sql = sql & ", TATWORK1.W_AMT PCT" & Format$(i, "00") & vbCrLf
        Next i
        sql = sql & " from TATWORK1,GLTFINR2,GLTFINR3 where ROWNUM < 1"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTFINRX", 4))
        With dst.Tables("GLTFINRX")
            For Each C As String In New String() {"ACCT_CODE",
                                                  "SEG2_CODE", "SEG3_CODE", "SEG4_CODE",
                                                  "SEG2_CLASS_CODE", "SEG3_CLASS_CODE", "SEG4_CLASS_CODE"}
                .Columns(C).AllowDBNull = True
            Next
            .Columns("SUPPRESS_PRINT").ReadOnly = False
        End With

        With dst.Tables.Add("GLTSTMTE")
            For Each COLUMN_NAME As String In New String() _
                {"STMT_CODE", "REPORT_NO", "STMT_LINE_NO", "STMT_LINE_NO2",
                 "STMT_LINE_DESC",
                 "ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "SEG2_CLASS_CODE", "SEG3_CLASS_CODE", "SEG4_CLASS_CODE"}
                If New String() {"REPORT_NO", "STMT_LINE_NO", "STMT_LINE_NO2"}.Contains(COLUMN_NAME) Then
                    .Columns.Add(COLUMN_NAME, GetType(System.Int32))
                Else
                    .Columns.Add(COLUMN_NAME)
                End If
            Next
            For i As Integer = 1 To MAX_COLUMNS
                .Columns.Add("AMT" & Format$(i, "00"), GetType(System.Decimal))
                .Columns.Add("PCT" & Format$(i, "00"), GetType(System.Decimal))
            Next i
            .PrimaryKey = New DataColumn() { .Columns("STMT_CODE"), .Columns("REPORT_NO"), .Columns("STMT_LINE_NO"), .Columns("STMT_LINE_NO2")}
        End With




        Prepare_FS()

        sql = "Select GLTACCT1.* from GLTACCT1,(SELECT DISTINCT ACCT_CODE FROM " & TT_Work & ") TTT where GLTACCT1.ACCT_CODE = TTT.ACCT_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1", 1))

    End Sub

    Sub Load_Layout()

        Dim sql As String = "Select * from GLTCLAY1 where STMT_LAYOUT_CODE = '" & HFs("STMT_LAYOUT_CODE") & "'"
        rowGLTCLAY1 = ASCDATA1.GetDataRow(sql)

        Dim j As Integer
        Dim z As String

        ' Create_Lookup("GLTCALC1")

        ReDim xy(MAX_COLUMNS * 2, 5)
        ReDim rd(MAX_COLUMNS)
        ReDim FH(MAX_COLUMNS, 1)

        STMT_CALC_NO_REF_PCT = ""
        For i As Integer = 1 To MAX_COLUMNS
            Dim rowGLTCLAY2 As DataRow = dst.Tables("GLTCLAY2").Rows.Find(New Object() {cmbSTMT_LAYOUT_CODE.Value, i})
            Dim STMT_CALC_CODE As String = ""
            If rowGLTCLAY2 IsNot Nothing Then
                STMT_CALC_CODE = rowGLTCLAY2.Item("STMT_CALC_CODE") & ""
            End If
            ' Dim STMT_CALC_CODE As String = rowGLTCLAY1.Item("STMT_CALC_CODE_" & Format(i, "00")) & ""
            If STMT_CALC_CODE <> "" Then

                Dim rowGLTCALC1 As DataRow = LookUp("GLTCALC1", STMT_CALC_CODE)
                If rowGLTCALC1.Item("STMT_CALC_NO_REF_PCT") & "" = "1" Then
                    STMT_CALC_NO_REF_PCT &= "1"
                Else
                    STMT_CALC_NO_REF_PCT &= "0"
                End If

                xy(i, 1) = Val(rowGLTCALC1.Item("STMT_CALC_YEAR") & "")
                xy(i, 2) = Val(rowGLTCALC1.Item("STMT_CALC_NO") & "")
                xy(i, 3) = Val(rowGLTCALC1.Item("STMT_CALC_TYPE") & "")
                xy(i, 4) = Val(rowGLTCALC1.Item("STMT_CALC_DATA_TYPE") & "")
                xy(i, 5) = Val(rowGLTCALC1.Item("STMT_CALC_PERIOD") & "")
                If xy(i, 3) = 1 Then
                    rd(0) = rd(0) + 1
                    ' rd(i) = 10 + rd(0)
                    rd(i) = MAX_COLUMNS + rd(0)

                    'MAX_COLUMNS = MAX_COLUMNS + 1
                    'rd(i) = MAX_COLUMNS
                    Select Case Val(rowGLTCALC1.Item("STMT_CALC_DATA_TYPE") & "")
                        Case 2 ' (TY - RB)
                            xy(rd(i), 1) = Val(rowGLTCALC1.Item("STMT_CALC_YEAR") & "")
                            xy(rd(i), 2) = Val(rowGLTCALC1.Item("STMT_CALC_NO") & "")
                            xy(rd(i), 3) = 0 ' Value, not %
                            xy(rd(i), 4) = 1 ' Denominator s/b RB
                            xy(rd(i), 5) = Val(rowGLTCALC1.Item("STMT_CALC_PERIOD") & "")

                        Case 3 ' (TY - LY)
                            xy(rd(i), 1) = Val(rowGLTCALC1.Item("STMT_CALC_YEAR") & "") + 1
                            xy(rd(i), 2) = Val(rowGLTCALC1.Item("STMT_CALC_NO") & "")
                            xy(rd(i), 3) = 0 ' Value, not %
                            xy(rd(i), 4) = 0 ' Denominator s/b LY - so use TY but STMT_CALC_YEAR + 1
                            xy(rd(i), 5) = Val(rowGLTCALC1.Item("STMT_CALC_PERIOD") & "")

                        Case 5 ' (TY - OB) - Denominator s/b OB
                            xy(rd(i), 1) = Val(rowGLTCALC1.Item("STMT_CALC_YEAR") & "")
                            xy(rd(i), 2) = Val(rowGLTCALC1.Item("STMT_CALC_NO") & "")
                            xy(rd(i), 3) = 0 ' Value, not %
                            xy(rd(i), 4) = 4 ' Denominator s/b OB
                            xy(rd(i), 5) = Val(rowGLTCALC1.Item("STMT_CALC_PERIOD") & "")

                        Case 6 ' (RB - OB) - Denominator s/b OB
                            xy(rd(i), 1) = Val(rowGLTCALC1.Item("STMT_CALC_YEAR") & "")
                            xy(rd(i), 2) = Val(rowGLTCALC1.Item("STMT_CALC_NO") & "")
                            xy(rd(i), 3) = 0 ' Value, not %
                            xy(rd(i), 4) = 4 ' Denominator s/b OB
                            xy(rd(i), 5) = Val(rowGLTCALC1.Item("STMT_CALC_PERIOD") & "")

                    End Select
                    'xy(rd(i), 1) = Val(rowGLTCALC1.Item("STMT_CALC_YEAR") & "") + 1
                    'xy(rd(i), 2) = Val(rowGLTCALC1.Item("STMT_CALC_NO") & "")
                    'xy(rd(i), 3) = 0
                    'xy(rd(i), 4) = 0
                    'xy(rd(i), 5) = 0
                End If


                FH(i, 1) = STMT_CALC_CODE ' rowGLTCLAY1.Item("STMT_CALC_CODE_" & Format$(i, "00"))

                If rowGLTCALC1.Item("STMT_CALC_DESC_TOKEN") & "" = "" Then
                    FH(i, 0) = rowGLTCALC1.Item("STMT_CALC_DESC") & ""
                Else
                    FH(i, 0) = rowGLTCALC1.Item("STMT_CALC_DESC_TOKEN") & ""

                    Dim YMS As String
                    Dim YME As String
                    Dim ADJ_S As Integer
                    Dim ADJ_E As Integer
                    Select Case xy(i, 2)
                        Case 1
                            ADJ_S = 0
                            ADJ_E = 0
                        Case 2  ' YTD
                            ADJ_S = 1 - ADJ_P
                            ADJ_E = 0
                            'ADJ_S = 12 - ADJ_P + 1 ' WAS 1 - ADJ_P
                            'ADJ_E = 12 - 0 ' WAS 0
                        Case 4 ' QTD
                            ADJ_S = 1 + 3 * (Int((ADJ_P - 1) / 3)) - ADJ_P
                            ADJ_E = 0
                            'ADJ_S = 12 + 1 + 3 * (Int((ADJ_P - 1) / 3)) - ADJ_P
                            'ADJ_E = 12 + 0
                        Case 6 ' HTD
                            ADJ_S = 1 + 6 * (Int((ADJ_P - 1) / 6)) - ADJ_P
                            ADJ_E = 0
                            'ADJ_S = 12 + 1 + 6 * (Int((ADJ_P - 1) / 6)) - ADJ_P
                            'ADJ_E = 12 + 0
                        Case 3 ' Total Year
                            ADJ_S = 1 - ADJ_P
                            ADJ_E = 12 - ADJ_P
                        Case 5
                            ADJ_S = 3 * Int((ADJ_P - 1) / 3) - ADJ_P - 3 + 1
                            ADJ_E = 3 * Int((ADJ_P - 1) / 3) - ADJ_P
                        Case 7
                            ADJ_S = 6 * Int((ADJ_P - 1) / 6) - ADJ_P - 6 + 1
                            ADJ_E = 6 * Int((ADJ_P - 1) / 6) - ADJ_P
                        Case 8
                            ADJ_S = 1 * xy(i, 5) - ADJ_P - 1 + 1
                            ADJ_E = 1 * xy(i, 5) - ADJ_P
                            'ADJ_S = 12 + 1 * xy(i, 5) - ADJ_P - 1 + 1
                            'ADJ_E = 12 + 1 * xy(i, 5) - ADJ_P
                        Case 9
                            ADJ_S = 3 * xy(i, 5) - ADJ_P - 3 + 1
                            ADJ_E = 3 * xy(i, 5) - ADJ_P
                            'ADJ_S = 12 + 3 * xy(i, 5) - ADJ_P - 3 + 1
                            'ADJ_E = 12 + 3 * xy(i, 5) - ADJ_P
                        Case 10
                            ADJ_S = 6 * xy(i, 5) - ADJ_P - 6 + 1
                            ADJ_E = 6 * xy(i, 5) - ADJ_P
                            'ADJ_S = 12 + 6 * xy(i, 5) - ADJ_P - 6 + 1
                            'ADJ_E = 12 + 6 * xy(i, 5) - ADJ_P
                    End Select

                    'z = ASCMAIN1.Get_YYYYMM(ASCMAIN1.Period_Calc(RYP, ADJ_S - 12 * (xy(i, 1) + 1)), 0)
                    ' CHANGED CALC FOR I01 INCOME STATEMENT PTD WHICH WAS SHOWING MAR'06 WHEN IT SHOULD HAVE BEEN SHOWING MAR'07 FOR EEE'YY
                    z = ASCMAIN1.Get_YYYYMM(ASCMAIN1.Period_Calc(RYP, ADJ_S - 12 * (xy(i, 1))), 0)
                    YMS = Replace(Format(DateValue(Mid$(z, 5, 2) & "/01/" & Mid$(z, 1, 4)), "MMM/yy"), "/", "'")
                    'z = ASCMAIN1.Get_YYYYMM(ASCMAIN1.Period_Calc(RYP, ADJ_E - 12 * (xy(i, 1) + 1)), 0)
                    ' CHANGED CALC FOR I01 INCOME STATEMENT PTD WHICH WAS SHOWING MAR'06 WHEN IT SHOULD HAVE BEEN SHOWING MAR'07 FOR EEE'YY
                    z = ASCMAIN1.Get_YYYYMM(ASCMAIN1.Period_Calc(RYP, ADJ_E - 12 * (xy(i, 1))), 0)
                    YME = Replace(Format(DateValue(Mid$(z, 5, 2) & "/01/" & Mid$(z, 1, 4)), "MMM/yy"), "/", "'")

                    Dim YYYY As String = Format(DateValue(Mid$(z, 5, 2) & "/01/" & Mid$(z, 1, 4)), "yyyy")

                    j = InStr(FH(i, 0), "MMM'YY")
                    If j <> 0 Then
                        FH(i, 0) = Mid$(FH(i, 0), 1, j - 1) & YME & Mid$(FH(i, 0), j + 6)
                    End If
                    j = InStr(FH(i, 0), "SSS'YY")
                    If j <> 0 Then
                        FH(i, 0) = Mid$(FH(i, 0), 1, j - 1) & YMS & Mid$(FH(i, 0), j + 6)
                    End If
                    j = InStr(FH(i, 0), "EEE'YY")
                    If j <> 0 Then
                        FH(i, 0) = Mid$(FH(i, 0), 1, j - 1) & YME & Mid$(FH(i, 0), j + 6)
                    End If
                    j = InStr(FH(i, 0), "YYYY")
                    If j <> 0 Then
                        FH(i, 0) = Mid$(FH(i, 0), 1, j - 1) & YYYY & Mid$(FH(i, 0), j + 4)
                    End If
                End If
            End If
        Next i

        txtREPORT_TITLE.Text = "Financial Report " & Mid(ASCMAIN1.Get_Legend(RYP), 1, 16)

    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("GYPLEGEND", ASCMAIN1.Get_Legend(RYP))
        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("SHOW_000S").Checked, "1", "0"))

        Dim z As String = ""
        If SQLA("SEG2_CODE", "CODE_VALUES") <> "" _
        Or SQLA("SEG3_CODE", "CODE_VALUES") <> "" _
        Or SQLA("SEG4_CODE", "CODE_VALUES") <> "" Then
            For i As Integer = 2 To 4
                Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                Dim CODE_VALUES As String = SQLA(COLUMN_NAME, "CODE_VALUES")
                If CODE_VALUES <> "" Then
                    If SQLA(COLUMN_NAME, "EXCLUDE") = "1" Then
                        z = z & "x"
                    End If
                    z = z & SQLA(COLUMN_NAME, "COLUMN_CAPTION") & ":" & CODE_VALUES & ", "
                End If
            Next i
            If z <> "" Then
                z = Mid$(z, 1, Len(z) - 2)
            End If
        End If

        CR_params.Add("SEL", z)

        CR_params.Add("DTL", IIf(Absx1.chkFor("SHOW_ACCTS").Checked, "1", "0"))
        CR_params.Add("SUPPRESSZERO", "0")
        CR_params.Add("REPORT", Absx1.optFor("REPORT").Value)
        CR_params.Add("NO_REF_PCT", STMT_CALC_NO_REF_PCT)

        ReDim HDGs(MAX_COLUMNS)
        For i As Integer = 1 To MAX_COLUMNS
            Dim rowGLTCLAY2 As DataRow = dst.Tables("GLTCLAY2").Rows.Find(New Object() {cmbSTMT_LAYOUT_CODE.Value, i})
            Dim MIN_PRD As Integer = 0
            If rowGLTCLAY2 IsNot Nothing Then
                MIN_PRD = Val(rowGLTCLAY2("STMT_MIN_PRD") & "")
            End If
            'Dim MIN_PRD As Integer = Val(rowGLTCLAY1("STMT_MIN_PRD_" & Format(i, "00")) & "")
            Dim HDG As String = ""
            If MIN_PRD <> 0 And MIN_PRD > P Then
                CR_params.Add("HD" & Format(i, "00"), "")
            Else
                HDG = Replace(FH(i, 0), "^", vbCrLf) & ""
                CR_params.Add("HD" & Format(i, "00"), HDG)
            End If
            HDGs(i) = HDG
        Next i

        If Absx1.optFor("REPORT").Value = "1" Then
            RPT = "GLRSTMT1"
        ElseIf Absx1.optFor("REPORT").Value = "2" Then
            RPT = "GLRSTMT1" ' "GLRSTMT2"
        ElseIf Absx1.optFor("REPORT").Value = "3" Then
            RPT = "GLRSTMT3"
        End If

        If rowGLTCLAY1.Item("RPT") & "" <> "" Then
            RPT = rowGLTCLAY1.Item("RPT")
        End If

        Dim REPORT_TITLE As String = Absx1.txtFor("REPORT_TITLE").Text
        CR_params.Add("REPORT_TITLE", REPORT_TITLE)

        If Absx1.optFor("REPORT").Value = "0" Then

            dst.Tables("GLTSTMTE").Rows.Clear()
            For Each rowGLTFINRX As DataRow In dst.Tables("GLTFINRX").Select("")
                Dim STMT_CODE As String = rowGLTFINRX.Item("STMT_CODE")
                Dim REPORT_NO As Integer = Val(rowGLTFINRX.Item("REPORT_NO") & "")
                Dim STMT_LINE_NO As Integer = Val(rowGLTFINRX.Item("STMT_LINE_NO") & "")
                Dim STMT_LINE_NO2 As Integer = Val(rowGLTFINRX.Item("STMT_LINE_NO2") & "")
                Dim rowGLTSTMTE As DataRow = dst.Tables("GLTSTMTE").NewRow
                For Each DC As DataColumn In dst.Tables("GLTSTMTE").Columns
                    If DC.ColumnName = "STMT_LINE_DESC" Then
                        Dim rowGLTFINR2 As DataRow = dst.Tables("GLTFINR2").Rows.Find(New Object() {STMT_CODE, STMT_LINE_NO})
                        If STMT_LINE_NO2 = 0 Then
                            rowGLTSTMTE.Item("STMT_LINE_DESC") = rowGLTFINR2.Item("STMT_LINE_DESC")
                        Else
                            Dim ACCT_CODE As String = rowGLTFINRX.Item("ACCT_CODE")
                            Dim rowGLTACCT1 As DataRow = dst.Tables("GLTACCT1").Rows.Find(ACCT_CODE)
                            Dim ACCT_DESC As String = rowGLTACCT1.Item("ACCT_DESC")
                            rowGLTSTMTE.Item("STMT_LINE_DESC") = ACCT_DESC
                        End If
                    Else
                        'If DC.ColumnName = "AMT01" Then
                        '    If Val(rowGLTFINRX.Item("AMT01") & "") <> 0 And Val(rowGLTFINRX.Item("AMT02") & "") Then
                        '        rowGLTSTMTE.Item("PCT01") = Val(rowGLTFINRX.Item("AMT01") & "") / Val(rowGLTFINRX.Item("AMT02") & "")
                        '    End If
                        'End If
                        'If DC.ColumnName <> "PCT01" Then
                        rowGLTSTMTE.Item(DC.ColumnName) = rowGLTFINRX.Item(DC.ColumnName)
                        'End If
                    End If
                Next
                dst.Tables("GLTSTMTE").Rows.Add(rowGLTSTMTE)
            Next

            Dim workbook As SpreadsheetGear.IWorkbook
            Dim XLS_FILENAME As String = ASCMAIN1.Folders("Work") & XNO & ".xlsX"
            workbook = SpreadsheetGear.Factory.GetWorkbook()

            Dim REPORT_NO_max As Integer = Val(dst.Tables("GLTSTMTX").Compute("MAX(REPORT_NO)", "") & "")

            Dim SFX As String = ""
            If REPORT_NO_max > 1 Then
                Dim MASK As String = "0"
                If REPORT_NO_max >= 10 Then MASK = "00"
                If REPORT_NO_max >= 100 Then MASK = "000"
                If REPORT_NO_max >= 1000 Then MASK = "000"
                If REPORT_NO_max >= 10000 Then MASK = "0000"
                SFX = "_" & Format(REPORT_NO, MASK)
            End If

            Dim ws As Integer = 0
            For Each row As DataRow In dst.Tables("GLTSTMTX").Select("", "REPORT_NO")
                Dim STMT_CODE As String = HFs("STMT_CODE")
                Dim REPORT_NO As Integer = Val(row.Item("REPORT_NO"))
                Dim SUPPRESS_PAGE As String = row.Item("SUPPRESS_PAGE") & ""
                If SUPPRESS_PAGE <> "1" Then

                    Dim description As String = ""
                    Dim sheet_name As String = ""

                    If tblASTDSQLA.Select("SEQUENCE IS NOT NULL").Length >= 1 Then
                        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE IS NOT NULL", "SEQUENCE")
                            Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")
                            Dim CODE_VALUE As String = row.Item(COLUMN_NAME)
                            If CODE_VALUE = "" Then CODE_VALUE = "000"
                            description &= "," & rowASTDSQLA.Item("COLUMN_CAPTION") & " " & CODE_VALUE
                        Next
                        sheet_name = Mid(description, 2)
                    Else
                        For Each T As String In New String() {"", "_GROUP", "_CLASS"}
                            For I As Integer = 2 To 4
                                Dim COLUMN_NAME As String = "SEG" & CStr(I) & T & "_CODE"
                                If row.Item(COLUMN_NAME) & "" <> "" Then
                                    Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(COLUMN_NAME)
                                    description &= "," & rowASTDSQLA.Item("COLUMN_CAPTION") & " " & row.Item(COLUMN_NAME) ' & " " & row.Item(Replace(COLUMN_NAME, "_CODE", "_DESC"))
                                    If sheet_name = "" Then
                                        sheet_name = rowASTDSQLA.Item("COLUMN_CAPTION") & " " & row.Item(COLUMN_NAME) & " " & row.Item(Replace(COLUMN_NAME, "_CODE", "_DESC"))
                                    Else
                                        sheet_name = Mid(description, 2) ' CStr(REPORT_NO)
                                    End If
                                End If
                            Next
                        Next
                    End If

                    ws += 1
                    If ws > 1 Then
                        workbook.Worksheets.Add()
                    End If
                    Excel_Extract(STMT_CODE, REPORT_NO, workbook, Mid(description, 2), sheet_name)
                End If
            Next

            workbook.Worksheets(0).Select()
            workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
            Show_Document(XLS_FILENAME)
        Else
            Generate_Report(RPT, , Absx1.cmbFor("STMT_LAYOUT_CODE").ActiveRow.Cells("STMT_LAYOUT_DESC").Text)
        End If

        PB_Report = True

    End Sub

    Function Prepare_Work_File() As String


        TTA = GL_Prep(Format(Val(RY) - 5, "0000"), Format(Val(RY) + 1, "0000"))
        TTB = GL_Prep(Format(Val(RY) - 5, "0000"), Format(Val(RY) + 1, "0000"), True)
        TTC = GL_Prep(Format(Val(RY) - 5, "0000"), Format(Val(RY) + 1, "0000"), True, , , "4")

        ' Clean out data which represents actuals beyond the Report Period Selected

        ASCDATA1.ExecuteSQL("Delete from " & TTA & " where ACCT_YEAR > '" & RY & "'")
        If P <> 12 Then
            Dim sql_clear As String = ""
            For i As Integer = P + 1 To 12
                sql_clear &= ", ACCT_ACT_P" & Format(i, "00") & " = 0"
            Next
            ASCDATA1.ExecuteSQL("Update " & TTA & " Set " & Mid(sql_clear, 2) & " where ACCT_YEAR = '" & RY & "'")
        End If



        Setup_SQL()

        Dim sql As String = ""

        Dim A234 As String = "ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE"

        Dim A234T As String = A234 & ", ACCT_TYPE"
        Dim sqlC As String =
            "Select Distinct " & A234T & " from " & TTA & " union " &
            "Select Distinct " & A234T & " from " & TTB & " union " &
            "Select Distinct " & A234T & " from " & TTC

        sql = "Select Distinct " & A234T & " from (" & sqlC & ")"
        Dim TTCA234 As String = ASCMAIN1.Temp_Table(sql)

        ASCMAIN1.sql = "Select GLTFINR3.*" & vbCrLf _
            & ", GLTFINR3.SEG2_CODE SEG2_CLASS_CODE" & vbCrLf _
            & ", GLTFINR3.SEG3_CODE SEG3_CLASS_CODE" & vbCrLf _
            & ", GLTFINR3.SEG4_CODE SEG4_CLASS_CODE" & vbCrLf _
            & " from GLTFINR3 where ROWNUM < 1"
        GLTFINRD = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        Dim STMT_LINE_NO_ALL_ELSE As Integer = 0
        Dim STMT_LINE_NO As Integer = 0
        For Each row As DataRow In dst.Tables("GLTFINR2").Select("STMT_LINE_TYPE = 'D'", "STMT_LINE_NO", DataViewRowState.CurrentRows)
            STMT_LINE_NO = Val(row.Item("STMT_LINE_NO"))
            sql = ""
            Dim sqlx As String = "Select Distinct '" & HFs("STMT_CODE") & "' STMT_CODE, " & CStr(STMT_LINE_NO) & " STMT_LINE_NO, " & A234
            Dim sqlf As String = " from " & TTCA234

            Select Case row.Item("STMT_LINE_ACCTS") & ""
                Case "S"
                    Dim dvwGLTFINR3 As New DataView(dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
                    Dim z As String = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        z &= ",'" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                    Next
                    If z <> "" Then
                        sql = sqlx & sqlf & " where ACCT_CODE in (" & Mid(z, 2) & ")"
                        For s As Integer = 2 To 4
                            If row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "S" _
                            Or row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "X" Then
                                Dim dvwGLTFINR4 As New DataView(dst.Tables("GLTFINR4"))
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
                    sql = sqlx & sqlf & " where ACCT_CODE >= '" & row.Item("STMT_LINE_ACCT_RANGE1") & "' and ACCT_CODE <= '" & row.Item("STMT_LINE_ACCT_RANGE2") & "'"

                Case "I"
                    If HFs("STMT_TYPE") = "B" Then
                        sql = sqlx & sqlf & " where ACCT_TYPE in " & ACCT_TYPEs("I")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "B"
                    If HFs("STMT_TYPE") = "I" Then
                        sql = sqlx & sqlf & " where ACCT_TYPE in " & ACCT_TYPEs("B")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "X"
                    Dim dvwGLTFINR3 As New DataView(dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO) & " AND (SEG2_CODE <> '*' AND SEG3_CODE <> '*' AND SEG4_CODE <> '*')"
                    Dim z As String = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        z &= ",('" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG2_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG3_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG4_CODE") & "')"
                    Next
                    If z <> "" Then
                        sql = sqlx & sqlf & " where (" & A234 & ") in (" & Mid(z, 2) & ")"
                    End If

                    Dim dvwGLTFINR3wc As New DataView(dst.Tables("GLTFINR3"))
                    dvwGLTFINR3wc.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO) & " AND (SEG2_CODE = '*' or SEG3_CODE = '*' or SEG4_CODE = '*')"
                    For i As Integer = 0 To dvwGLTFINR3wc.Count - 1
                        Dim sqlWildCard As String = ""
                        Dim sqlWildCard_where As String = " where (ACCT_CODE"
                        Dim sqlWildCard_in As String = "(('" & dvwGLTFINR3wc(i).Item("ACCT_CODE") & "'"
                        If dvwGLTFINR3wc(i).Item("SEG2_CODE") <> "*" Then
                            sqlWildCard_where &= ",SEG2_CODE"
                            sqlWildCard_in &= ",'" & dvwGLTFINR3wc(i).Item("SEG2_CODE") & "'"
                        End If
                        If dvwGLTFINR3wc(i).Item("SEG3_CODE") <> "*" Then
                            sqlWildCard_where &= ",SEG3_CODE"
                            sqlWildCard_in &= ",'" & dvwGLTFINR3wc(i).Item("SEG3_CODE") & "'"
                        End If
                        If dvwGLTFINR3wc(i).Item("SEG4_CODE") <> "*" Then
                            sqlWildCard_where &= ",SEG4_CODE"
                            sqlWildCard_in &= ",'" & dvwGLTFINR3wc(i).Item("SEG4_CODE") & "'"
                        End If
                        sqlWildCard_where &= ")"
                        sqlWildCard_in &= "))"
                        sqlWildCard = "Insert into " & GLTFINRD & " " & sqlx & sqlf & sqlWildCard_where & " in " & sqlWildCard_in

                        ASCDATA1.ExecuteSQL(sqlWildCard)
                    Next
            End Select

            If sql <> "" Then
                sql = "Insert into " & GLTFINRD & " " & vbCrLf _
                    & "Select X.*" & vbCrLf _
                    & ", GLTSEGM2.ACCT_SEG_CLASS SEG2_CLASS_CODE" & vbCrLf _
                    & ", GLTSEGM3.ACCT_SEG_CLASS SEG3_CLASS_CODE" & vbCrLf _
                    & ", GLTSEGM4.ACCT_SEG_CLASS SEG4_CLASS_CODE" & vbCrLf _
                    & " from (" & sql & ") X" & vbCrLf _
                    & ", GLTSEGM1 GLTSEGM2, GLTSEGM1 GLTSEGM3, GLTSEGM1 GLTSEGM4" & vbCrLf _
                    & " where GLTSEGM2.ACCT_SEG_ID (+) = '2' and GLTSEGM2.ACCT_SEG_CODE (+) = X.SEG2_CODE" & vbCrLf _
                    & "   and GLTSEGM3.ACCT_SEG_ID (+) = '3' and GLTSEGM3.ACCT_SEG_CODE (+) = X.SEG3_CODE" & vbCrLf _
                    & "   and GLTSEGM4.ACCT_SEG_ID (+) = '4' and GLTSEGM4.ACCT_SEG_CODE (+) = X.SEG4_CODE"
                ASCDATA1.ExecuteSQL(sql)
            End If
        Next

        If STMT_LINE_NO_ALL_ELSE <> 0 Then
            sql = "Select '" & HFs("STMT_CODE") & "' STMT_CODE, " & CStr(STMT_LINE_NO_ALL_ELSE) & " STMT_LINE_NO, " & A234 & " " _
                    & " from (" _
                    & "Select DISTINCT TT." & Replace(A234, ", ", ", TT.") & " from GLTACCT1," & TTCA234 & " TT where TT.ACCT_CODE = GLTACCT1.ACCT_CODE and GLTACCT1.ACCT_TYPE in " & ACCT_TYPEs(HFs("STMT_TYPE")) _
                    & " MINUS " _
                    & "Select DISTINCT " & A234 & " from " & GLTFINRD _
                    & ")"
            sql = "Insert into " & GLTFINRD & " " & vbCrLf _
                & "Select X.*" & vbCrLf _
                & ", GLTSEGM2.ACCT_SEG_CLASS SEG2_CLASS_CODE" & vbCrLf _
                & ", GLTSEGM3.ACCT_SEG_CLASS SEG3_CLASS_CODE" & vbCrLf _
                & ", GLTSEGM4.ACCT_SEG_CLASS SEG4_CLASS_CODE" & vbCrLf _
                & " from (" & sql & ") X" & vbCrLf _
                & ", GLTSEGM1 GLTSEGM2, GLTSEGM1 GLTSEGM3, GLTSEGM1 GLTSEGM4" & vbCrLf _
                & " where GLTSEGM2.ACCT_SEG_ID (+) = '2' and GLTSEGM2.ACCT_SEG_CODE (+) = X.SEG2_CODE" & vbCrLf _
                & "   and GLTSEGM3.ACCT_SEG_ID (+) = '3' and GLTSEGM3.ACCT_SEG_CODE (+) = X.SEG3_CODE" & vbCrLf _
                & "   and GLTSEGM4.ACCT_SEG_ID (+) = '4' and GLTSEGM4.ACCT_SEG_CODE (+) = X.SEG4_CODE"
            ASCDATA1.ExecuteSQL(sql)
        End If

        ASCMAIN1.AnalyzeTable(GLTFINRD)

        Return TTCA234

    End Function

    Private Sub UltraCheckEditor5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraCheckEditor5.CheckedChanged
        grpBreakoutBy.Visible = Absx1.chkFor("SHOW_ACCTS").Checked
        Absx1.chkFor("SHOW_TRANS").Visible = Absx1.chkFor("SHOW_ACCTS").Checked
        If Not Absx1.chkFor("SHOW_ACCTS").Checked Then
            Absx1.chkFor("SHOW_TRANS").Checked = False
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("STMT_CODE").Text = "" Then
                EMsg &= vbCr & "You must pick a Statement"
            End If
            If Absx1.cmbFor("STMT_LAYOUT_CODE").Text = "" Then
                EMsg &= vbCr & "You must pick a Column Layout"
            End If
        End If

    End Sub

    Sub Prepare_FS()

        Dim REPORT_NO As Integer = 0
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "G1,G2,G3,G4,G5,G6,G7,G8,G9")
            REPORT_NO = REPORT_NO + 1
            rowASTSRPT1.Item("REPORT_NO") = REPORT_NO
            Process_FS_Report(REPORT_NO, rowASTSRPT1)
        Next

        For Each rowGLTFINRX As DataRow In dst.Tables("GLTFINRX").Rows
            For i As Integer = 1 To MAX_COLUMNS
                rowGLTFINRX.Item("AMT" & Format(i, "00")) = Round(rowGLTFINRX.Item("AMT" & Format(i, "00")), 2)
            Next
        Next

        For Each rowGLTSTMTX As DataRow In dst.Tables("GLTSTMTX").Rows
            Dim T As Double = 0
            sql = "STMT_CODE = '" & HFs("STMT_CODE") & "' AND REPORT_NO = " & rowGLTSTMTX.Item("REPORT_NO") ' & " AND ACCT_CODE IS NOT NULL"

            For I As Integer = 1 To MAX_COLUMNS
                T = Val(dst.Tables("GLTFINRX").Compute("SUM(AMT" & Format(I, "00") & ")", sql) & "")
                If Abs(T) >= 1 Then
                    Exit For
                Else
                    T = Val(dst.Tables("GLTFINRX").Compute("MIN(AMT" & Format(I, "00") & ")", sql) & "")
                    If Abs(T) >= 1 Then
                        Exit For
                    Else
                        T = Val(dst.Tables("GLTFINRX").Compute("MAX(AMT" & Format(I, "00") & ")", sql) & "")
                        If Abs(T) >= 1 Then
                            Exit For
                        End If
                    End If
                End If
            Next
            If Abs(T) >= 1 Then
                sql = "STMT_CODE = '" & HFs("STMT_CODE") & "' and (STMT_LINE_TYPE = 'D' or STMT_LINE_TYPE = 'S') "
                If Not Absx1.chkFor("SUPPRESS_ZERO").Checked Then
                    sql = sql & "and STMT_LINE_PRINT = '0'"
                End If
                For Each rowGLTFINR2 As DataRow In dst.Tables("GLTFINR2").Select(sql, "STMT_LINE_NO")

                    Dim z As String = ""
                    For i As Integer = 1 To MAX_COLUMNS
                        'z = z & " and ABS(AMT" & Format(i, "00") & ") < 1"
                        z = z & " and AMT" & Format(i, "00") & " = 0"
                    Next i
                    sql = "STMT_CODE = '" & HFs("STMT_CODE") & "' AND REPORT_NO = " & rowGLTSTMTX.Item("REPORT_NO") & " and STMT_LINE_NO = " & rowGLTFINR2.Item("STMT_LINE_NO") & z

                    'Dim dv As New DataView(dst.Tables("GLTFINRX"), sql, "STMT_LINE_NO2", DataViewRowState.CurrentRows)
                    'For Each rowGLTFINRX As DataRow In dv.ToTable.Rows
                    '    rowGLTFINRX.Item("SUPPRESS_PRINT") = "1"
                    'Next
                    For Each rowGLTFINRX As DataRow In
                    dst.Tables("GLTFINRX").Select(sql, "STMT_LINE_NO2")

                        rowGLTFINRX.Item("SUPPRESS_PRINT") = "1"
                    Next
                Next
            Else
                rowGLTSTMTX.Item("SUPPRESS_PAGE") = "1"
            End If
        Next

        'If Absx1.chkFor("SHOW_ACCTS").Checked Then
        '    For Each rowGLTFINRX As DataRow In dst.Tables("GLTFINRX").Select("STMT_LINE_NO2 = 0")
        '        rowGLTFINRX.Item("STMT_LINE_NO2") = 9999
        '    Next
        'End If

    End Sub

    Sub Process_FS_Report(ByVal REPORT_NO As Integer, ByVal rowASTSRPT1 As DataRow)

        Dim rowGLTSTMTX As DataRow = dst.Tables("GLTSTMTX").NewRow
        rowGLTSTMTX.Item("REPORT_NO") = REPORT_NO

        rowGLTSTMTX.Item("SEG2_CODE") = ""
        rowGLTSTMTX.Item("SEG3_CODE") = ""
        rowGLTSTMTX.Item("SEG4_CODE") = ""
        rowGLTSTMTX.Item("SEG2_CLASS_CODE") = ""
        rowGLTSTMTX.Item("SEG3_CLASS_CODE") = ""
        rowGLTSTMTX.Item("SEG4_CLASS_CODE") = ""
        rowGLTSTMTX.Item("SEG2_GROUP_CODE") = ""
        rowGLTSTMTX.Item("SEG3_GROUP_CODE") = ""
        rowGLTSTMTX.Item("SEG4_GROUP_CODE") = ""


        dst.Tables("GLTSTMTX").Rows.Add(rowGLTSTMTX)


        ReDim rr(6, MAX_COLUMNS + MAX_COLUMNS)
        ReDim lineref(3, MAX_COLUMNS)

        Dim sql As String = "STMT_CODE = '" & HFs("STMT_CODE") & "'"
        STMT_LINE_REF_SETs = ""
        Dim STMT_LINE_NO_max As Integer = 0

        Dim sql_tables As String = ""

        Dim sql_where As String = ""
        For i As Integer = 1 To 9
            If rowASTSRPT1.Item("G" & CStr(i)) & "" <> "x" Then
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Select("SEQUENCE = " & CStr(i), "", DataViewRowState.CurrentRows)(0)
                Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")
                Dim ACCT_SEG_ID As String = Mid(COLUMN_NAME, 4, 1)
                Dim CODE_VALUE As String = Split(rowASTSRPT1.Item("G" & CStr(i)), ":")(1)
                Dim DESC_VALUE As String = ""
                Dim rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(rowASTSRPT1.Item("G" & CStr(i)) & "")
                If rowASTGROUP IsNot Nothing Then
                    DESC_VALUE = rowASTGROUP.ITEM("GROUP_DESC") & ""
                End If

                Select Case COLUMN_NAME
                    Case "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"
                        rowGLTSTMTX.Item("SEG" & ACCT_SEG_ID & "_CODE") = CODE_VALUE
                        rowGLTSTMTX.Item("SEG" & ACCT_SEG_ID & "_DESC") = DESC_VALUE

                        sql_where &= " and X." & COLUMN_NAME

                    Case "SEG2_GROUP_CODE", "SEG3_GROUP_CODE", "SEG4_GROUP_CODE"
                        rowGLTSTMTX.Item("SEG" & ACCT_SEG_ID & "_GROUP_CODE") = CODE_VALUE
                        rowGLTSTMTX.Item("SEG" & ACCT_SEG_ID & "_GROUP_DESC") = DESC_VALUE

                        sql_tables &= ", GLTSEGG2 " & "GLTSEGG" & ACCT_SEG_ID
                        sql_where &= " and GLTSEGG" & ACCT_SEG_ID & ".ACCT_SEG_ID = '" & ACCT_SEG_ID & "'"
                        'sql_where = " and GLTSEGG" & ACCT_SEG_ID & ".ACCT_SEG_CODE = X." & COLUMN_NAME
                        sql_where &= " and GLTSEGG" & ACCT_SEG_ID & ".ACCT_SEG_CODE = X." & Replace(COLUMN_NAME, "_GROUP_", "_")
                        sql_where &= " and GLTSEGG" & ACCT_SEG_ID & ".ACCT_SEG_GROUP_CODE"

                    Case "SEG2_CLASS_CODE", "SEG3_CLASS_CODE", "SEG4_CLASS_CODE"
                        rowGLTSTMTX.Item("SEG" & ACCT_SEG_ID & "_CLASS_CODE") = CODE_VALUE
                        rowGLTSTMTX.Item("SEG" & ACCT_SEG_ID & "_CLASS_DESC") = DESC_VALUE

                        sql_tables &= ", GLTSEGM1 " & "GLTSEGM" & ACCT_SEG_ID
                        sql_where &= " and GLTSEGM" & ACCT_SEG_ID & ".ACCT_SEG_ID = '" & ACCT_SEG_ID & "'"
                        'sql_where = sql_where & " and GLTSEGM" & ACCT_SEG_ID & ".ACCT_SEG_CODE = X." & COLUMN_NAME
                        sql_where &= " and GLTSEGM" & ACCT_SEG_ID & ".ACCT_SEG_CODE = X." & Replace(COLUMN_NAME, "_CLASS_", "_")
                        sql_where &= " and GLTSEGM" & ACCT_SEG_ID & ".ACCT_SEG_CLASS"
                End Select
                If CODE_VALUE = "" Then
                    sql_where &= " IS NULL"
                Else
                    sql_where &= " = '" & CODE_VALUE & "'"
                End If

            End If
        Next

        STMT_LINE_REFs.Clear()
        STFs.Clear()
        Dim STBs(6) As String

        For Each rowGLTFINR2 As DataRow In dst.Tables("GLTFINR2").Select(sql, "STMT_LINE_NO")

            Dim STMT_LINE_NO As Integer = Val(rowGLTFINR2.Item("STMT_LINE_NO") & "")
            Dim STMT_LINE_NO2 As Integer = 0
            Dim STMT_LINE_DC As String = rowGLTFINR2.Item("STMT_LINE_DC") & ""
            Dim STMT_LINE_TYPE As String = rowGLTFINR2.Item("STMT_LINE_TYPE") & ""

            If STMT_LINE_TYPE = "D" Then

                For B As Integer = 1 To 6
                    Dim STMT_SUBT_ADD As String = rowGLTFINR2.Item("STMT_SUBT_ADD" & CStr(B)) & ""
                    If STMT_SUBT_ADD = "1" Then
                        STBs(B) &= IIf(STMT_LINE_DC = "C", "-", "+") & "L" & Format(STMT_LINE_NO, "000")
                    End If
                Next

                Dim sqlS As String = ""
                sqlS = sqlS & " where GLTFINRD.STMT_CODE = '" & HFs("STMT_CODE") & "'"
                sqlS = sqlS & "   and GLTFINRD.STMT_LINE_NO = " & CStr(STMT_LINE_NO)

                Get_Details(sqlS & sqlA_where & sql_where, sql_tables, STMT_LINE_NO, False)
                Write_Details(rowGLTFINR2, REPORT_NO, STMT_LINE_NO2)

                If Absx1.chkFor("SHOW_ACCTS").Checked Then
                    sql = "Select X.ACCT_CODE" & sqlA_select _
                        & " from " & GLTFINRD & " X " & sql_tables _
                        & Replace(sqlS, "GLTFINRD.", "X.") & Replace(sqlA_where, "GLTFINRD.", "X.") & sql_where _
                        & " group by X.ACCT_CODE" & sqlA_group_by
                    For Each rowGLTFINRD As DataRow In ASCDATA1.GetDataTable(sql).Select("", "ACCT_CODE,SEG2_CLASS_CODE,SEG3_CLASS_CODE,SEG4_CLASS_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE")
                        Get_Details(sqlS & sqlA_where & sql_where, sql_tables, STMT_LINE_NO, True,
                            rowGLTFINRD.Item("ACCT_CODE"),
                            rowGLTFINRD.Item("SEG2_CODE"),
                            rowGLTFINRD.Item("SEG3_CODE"),
                            rowGLTFINRD.Item("SEG4_CODE"),
                            rowGLTFINRD.Item("SEG2_CLASS_CODE") & "",
                            rowGLTFINRD.Item("SEG3_CLASS_CODE") & "",
                            rowGLTFINRD.Item("SEG4_CLASS_CODE") & "")
                        STMT_LINE_NO2 = STMT_LINE_NO2 + 1
                        Write_Details(rowGLTFINR2, REPORT_NO, STMT_LINE_NO2,
                            rowGLTFINRD.Item("ACCT_CODE"),
                            rowGLTFINRD.Item("SEG2_CODE"),
                            rowGLTFINRD.Item("SEG3_CODE"),
                            rowGLTFINRD.Item("SEG4_CODE"),
                            rowGLTFINRD.Item("SEG2_CLASS_CODE") & "",
                            rowGLTFINRD.Item("SEG3_CLASS_CODE") & "",
                            rowGLTFINRD.Item("SEG4_CLASS_CODE") & "")
                    Next
                End If

            Else

                If STMT_LINE_TYPE = "S" Then
                    Dim STMT_SUBT_SHOW As String = Val(rowGLTFINR2.Item("STMT_SUBT_SHOW") & "")
                    If STMT_SUBT_SHOW <> 0 Then
                        STFs.Add(STMT_LINE_NO, STBs(STMT_SUBT_SHOW))
                        STBs(STMT_SUBT_SHOW) = ""
                    End If

                    For B As Integer = 1 To 6
                        Dim STMT_SUBT_ADD As String = rowGLTFINR2.Item("STMT_SUBT_ADD" & CStr(B)) & ""
                        If STMT_SUBT_ADD = "1" Then
                            STBs(B) &= IIf(STMT_LINE_DC = "C", "-", "+") & "L" & Format(STMT_LINE_NO, "000")
                        End If
                    Next

                End If

                Write_Details(rowGLTFINR2, REPORT_NO, 0)
            End If
        Next

        If STMT_LINE_REF_SETs <> "" Then
            For j As Integer = 1 To Len(STMT_LINE_REF_SETs)
                Dim STMT_LINE_REF_PCT As String = Mid(STMT_LINE_REF_SETs, j, 1)

                sql = "STMT_CODE = '" & HFs("STMT_CODE") & "'" _
                    & " and REPORT_NO = " & CStr(REPORT_NO) _
                    & " and STMT_LINE_REF_PCT = '" & STMT_LINE_REF_PCT & "'"
                For Each rowGLTFINRX As DataRow In dst.Tables("GLTFINRX") _
                    .Select(sql, "STMT_LINE_NO,STMT_LINE_NO2")
                    For i As Integer = 1 To MAX_COLUMNS
                        If xy(i, 3) = 2 Then
                            rowGLTFINRX.Item("PCT" & Format(i, "00")) = Null
                        ElseIf lineref(STMT_LINE_REF_PCT, i) = 0 Then
                            rowGLTFINRX.Item("PCT" & Format(i, "00")) = 0
                        Else
                            rowGLTFINRX.Item("PCT" & Format(i, "00")) = 100 * CStr(SHOW_000S) * rowGLTFINRX.Item("AMT" & Format(i, "00")) / lineref(STMT_LINE_REF_PCT, i)
                        End If
                    Next
                Next
            Next
        End If

        Exit Sub


    End Sub

    Sub Write_Details(
    ByRef rowGLTFINR2 As DataRow,
    ByVal REPORT_NO As Integer,
    ByRef STMT_LINE_NO2 As Integer,
    Optional ByVal ACCT_CODE As String = "",
    Optional ByVal SEG2_CODE As String = "",
    Optional ByVal SEG3_CODE As String = "",
    Optional ByVal SEG4_CODE As String = "",
    Optional ByVal SEG2_CLASS_CODE As String = "",
    Optional ByVal SEG3_CLASS_CODE As String = "",
    Optional ByVal SEG4_CLASS_CODE As String = "")

        Dim S As Integer = 0
        If rowGLTFINR2.Item("STMT_LINE_DC") & "" = "D" Then
            S = 1
        Else
            S = -1
        End If

        Dim STMT_LINE_TYPE As String = rowGLTFINR2.Item("STMT_LINE_TYPE") & ""
        If STMT_LINE_TYPE = "S" Then
            Dim STMT_SUBT_ADD As Integer = Val(rowGLTFINR2.Item("STMT_SUBT_SHOW") & "")
            If STMT_SUBT_ADD <> 0 Then
                For i As Integer = 1 To MAX_COLUMNS * 2
                    rr(0, i) = rr(STMT_SUBT_ADD, i)
                    rr(STMT_SUBT_ADD, i) = 0
                Next
            End If
        End If
        If STMT_LINE_TYPE = "H" Then
            For i As Integer = 1 To MAX_COLUMNS * 2
                rr(0, i) = 0
            Next
        End If
        If (STMT_LINE_TYPE = "D" Or STMT_LINE_TYPE = "S") _
            And STMT_LINE_NO2 = 0 Then
            For STMT_SUBT_ADD As Integer = 1 To 6
                If rowGLTFINR2.Item("STMT_SUBT_ADD" & Format$(STMT_SUBT_ADD, "0")) & "" = "1" Then
                    For j As Integer = 1 To MAX_COLUMNS * 2
                        rr(STMT_SUBT_ADD, j) = rr(STMT_SUBT_ADD, j) + rr(0, j)
                    Next
                End If
            Next
        End If

        Dim rowGLTFINRX As DataRow = dst.Tables("GLTFINRX").NewRow
        rowGLTFINRX.Item("STMT_CODE") = HFs("STMT_CODE")
        rowGLTFINRX.Item("REPORT_NO") = REPORT_NO
        rowGLTFINRX.Item("STMT_LINE_NO") = rowGLTFINR2.Item("STMT_LINE_NO")
        rowGLTFINRX.Item("STMT_LINE_NO2") = STMT_LINE_NO2
        If STMT_LINE_NO2 <> 0 Then
            rowGLTFINRX.Item("ACCT_CODE") = ACCT_CODE
            rowGLTFINRX.Item("SEG2_CODE") = SEG2_CODE
            rowGLTFINRX.Item("SEG3_CODE") = SEG3_CODE
            rowGLTFINRX.Item("SEG4_CODE") = SEG4_CODE
            rowGLTFINRX.Item("SEG2_CLASS_CODE") = SEG2_CLASS_CODE
            rowGLTFINRX.Item("SEG3_CLASS_CODE") = SEG3_CLASS_CODE
            rowGLTFINRX.Item("SEG4_CLASS_CODE") = SEG4_CLASS_CODE
        End If

        'STMT_LINE_NO2 = STMT_LINE_NO2 + 1

        Dim amt() As Double
        ReDim amt(MAX_COLUMNS)
        For i As Integer = 1 To MAX_COLUMNS
            If xy(i, 3) = 1 Then ' 1 = PCT 0 = AMT

                '  If rd(i) > MAX_COLUMNS Then rd(i) = MAX_COLUMNS

                If rr(0, rd(i)) = 0 Then
                    amt(i) = 0
                Else
                    amt(i) = 100 * rr(0, i) / rr(0, rd(i))
                End If
            Else
                amt(i) = S * rr(0, i) / SHOW_000S
            End If
            rowGLTFINRX.Item("AMT" & Format$(i, "00")) = amt(i)
            If Val(rowGLTFINR2.Item("STMT_LINE_REF_PCT") & "") <> 0 Then
                rowGLTFINRX.Item("STMT_LINE_REF_PCT") = Val(rowGLTFINR2.Item("STMT_LINE_REF_PCT") & "")
            End If
        Next

        dst.Tables("GLTFINRX").Rows.Add(rowGLTFINRX)

        If STMT_LINE_NO2 = 0 Then
            Dim STMT_LINE_REF_SET As Integer = Val(rowGLTFINR2.Item("STMT_LINE_REF_SET") & "")
            If STMT_LINE_REF_SET <> 0 Then
                If InStr(STMT_LINE_REF_SETs, CStr(STMT_LINE_REF_SET)) = 0 Then
                    STMT_LINE_REF_SETs = STMT_LINE_REF_SETs & CStr(STMT_LINE_REF_SET)
                    STMT_LINE_REFs.Add(STMT_LINE_REF_SET, Val(rowGLTFINR2.Item("STMT_LINE_NO") & ""))
                End If
                For i As Integer = 1 To MAX_COLUMNS
                    lineref(STMT_LINE_REF_SET, i) = rr(0, i) * S
                Next
            End If
        End If

    End Sub

    Sub Get_Details(
    ByVal sql_AB As String,
    ByVal sql_tables As String,
    STMT_LINE_NO As Int32,
    ByVal acct_details As Boolean,
    Optional ByVal ACCT_CODE As String = "",
    Optional ByVal SEG2_CODE As String = "",
    Optional ByVal SEG3_CODE As String = "",
    Optional ByVal SEG4_CODE As String = "",
    Optional ByVal SEG2_CLASS_CODE As String = "",
    Optional ByVal SEG3_CLASS_CODE As String = "",
    Optional ByVal SEG4_CLASS_CODE As String = "")

        ReDim A(5, 14, 6)
        ReDim FA(5, MAX_COLUMNS * 2, 2, 6)

        ' need to go to -1 for next year

        ' f(years, calculations, types, AB); type 0:ledger, 1:var$, 2:var%

        ' Draw Actuals, Revised Budgets, and Original Budgets in from db

        For AB As Integer = 0 To 4

            ' 0 (ACT) = GLTACCT3
            ' 1 (RBD) = GLTACCT2
            ' 4 (OBD) = GLTACCT4

            If AB = 2 Then
                AB = 4
            End If
            sql = "Select ACCT_YEAR" & vbCrLf
            sql = sql & SQLP(AB) & SQLF(AB) & ", " & GLTFINRD & " GLTFINRD" & vbCrLf _
                & sql_tables & sql_AB
            sql = sql & "   and X.ACCT_CODE = GLTFINRD.ACCT_CODE" & vbCrLf
            sql = sql & "   and X.SEG2_CODE = GLTFINRD.SEG2_CODE" & vbCrLf
            sql = sql & "   and X.SEG3_CODE = GLTFINRD.SEG3_CODE" & vbCrLf
            sql = sql & "   and X.SEG4_CODE = GLTFINRD.SEG4_CODE" & vbCrLf

            If acct_details Then
                sql = sql & " and X.ACCT_CODE = '" & ACCT_CODE & "'"
                Dim z As String
                For i As Integer = 2 To 4
                    z = "SEG" & CStr(i) & "_CODE"
                    If New Boolean() {BY_SEG2, BY_SEG3, BY_SEG4}(i - 2) Then
                        sql = sql & "   and X." & z & " = '" & New String() {SEG2_CODE, SEG3_CODE, SEG4_CODE}(i - 2) & "'"
                    End If
                    z = "SEG" & CStr(i) & "_CLASS_CODE"
                    If New Boolean() {BY_SEG2_CLASS, BY_SEG3_CLASS, BY_SEG4_CLASS}(i - 2) Then
                        Dim V As String = New String() {SEG2_CLASS_CODE, SEG3_CLASS_CODE, SEG4_CLASS_CODE}(i - 2)
                        If V = "" Then
                            sql = sql & "   and GLTFINRD." & z & " IS NULL"
                        Else
                            sql = sql & "   and GLTFINRD." & z & " = '" & V & "'"
                        End If
                    End If
                Next
            End If
            sql = sql & " group by X.ACCT_YEAR"

            For Each row As DataRow In ASCDATA1.GetDataTable(sql).Rows
                Dim y As Integer = RY - Val(row.Item("ACCT_YEAR") & "")
                If y >= 0 Then ' FOR NOW - NO NEXT YEAR
                    For j As Integer = 0 To 13 ' this 13 might be the number of periods
                        A(y, j, AB) = A(y, j, AB) + Val(row.Item(1 + j) & "")
                    Next j
                End If
            Next

            Dim QTR_IN_BEG_BAL As Integer = 1 + 3 * (Int((ADJ_P - 1) / 3))



            Dim INVTY(5, 6, 3) As Decimal
            If STMT_LINE_NO = STMT_LINE_NO_BEG Or STMT_LINE_NO = STMT_LINE_NO_END Then
                sql = Replace(sql, $"GLTFINRD.STMT_LINE_NO = {STMT_LINE_NO} ", $"GLTFINRD.STMT_LINE_NO = {STMT_LINE_NO_END} ")
                sql = Replace(sql, $"and X.ACCT_CODE = '1200'", $"")
                For Each row As DataRow In ASCDATA1.GetDataTable(sql).Rows
                    Dim y As Integer = RY - Val(row.Item("ACCT_YEAR") & "")
                    If y >= 0 Then ' FOR NOW - NO NEXT YEAR
                        For j As Integer = 0 To ADJ_P
                            INVTY(y, AB, 0) += Val(row.Item(1 + j) & "") ' YTD
                            If j = ADJ_P Then
                                INVTY(y, AB, 1) = Val(row.Item(1 + j) & "") ' MTD
                            End If
                            If j <= QTR_IN_BEG_BAL - 1 Then
                                INVTY(y, AB, 2) += Val(row.Item(1 + j) & "") ' QTR BEG
                            End If
                            If j >= QTR_IN_BEG_BAL Then
                                INVTY(y, AB, 3) += Val(row.Item(1 + j) & "") ' QTR
                            End If

                        Next j
                    End If
                Next
            End If


            Dim k As Integer = 0

            For i As Integer = -1 To 5          ' for relative years -1 thru 5
                If i >= 0 Then ' FOR NOW - NO NEXT YEAR


                    Dim INV_ADJ As Decimal = -1 * INVTY(i, AB, 0) + INVTY(i, AB, 1)
                    FA(i, 1, 0, AB) = A(i, ADJ_P, AB)       ' 1 mtd
                    If STMT_LINE_NO = STMT_LINE_NO_BEG Then ' Beg Invty
                        FA(i, 1, 0, AB) += INV_ADJ
                    ElseIf STMT_LINE_NO = STMT_LINE_NO_END Then ' End Invty
                        FA(i, 1, 0, AB) -= INV_ADJ
                    End If


                    For j As Integer = 0 To 13 ' this 13 might be the number of periods
                        If j <= ADJ_P Then
                            FA(i, 2, 0, AB) += A(i, j, AB)   ' 2 ytd
                        End If
                        FA(i, 3, 0, AB) += A(i, j, AB)       ' 3 total year
                    Next


                    k = Int((ADJ_P - 1) / 3) * 3
                    For j As Integer = k + 1 To k + 3
                        If j <= ADJ_P Then
                            INV_ADJ = INVTY(i, AB, 2) + INVTY(i, AB, 3)
                            If STMT_LINE_NO = STMT_LINE_NO_BEG Then ' Beg Invty

                                If k = 0 Then
                                    FA(0, 4, 0, 0) = A(0, 1, 0)
                                    '   FA(i, 4, 0, AB) = A(i, j, AB)
                                Else
                                    FA(i, 4, 0, AB) = (-1 * INVTY(i, AB, 2))
                                End If

                                '    Stop
                                ' FA(i, 4, 0, AB) += A(i, j, AB)
                            ElseIf STMT_LINE_NO = STMT_LINE_NO_END Then ' End Invty
                                FA(i, 4, 0, AB) = INV_ADJ
                            Else
                                FA(i, 4, 0, AB) += A(i, j, AB)   ' 4 qtd
                            End If
                        End If
                        FA(i, 5, 0, AB) += A(i, j, AB)       ' 5 total quarter
                    Next

                    k = Int((ADJ_P - 1) / 6) * 6
                    For j As Integer = k + 1 To k + 6
                        If j <= ADJ_P Then
                            FA(i, 6, 0, AB) += A(i, j, AB)   ' 6 htd
                        End If
                        FA(i, 7, 0, AB) += A(i, j, AB)       ' 7 total half
                    Next
                End If
            Next
        Next

        ' Calculate Variances
        ' f(years, calculations, types, AB); type 0:ledger, 1:var$, 2:var%

        For i As Integer = -1 To 5
            If i >= 0 Then ' FOR NOW 

                For j As Integer = 1 To MAX_COLUMNS * 2 '20 ' THIS WAS FOR WHEN MAX_COLUMNS WAS 10, +10 FOR DENOMINATORS, REALLY NEED TO SAY MAX_COLUMNS * 2 HERE
                    FA(i, j, 0, 2) = FA(i, j, 0, 0) - FA(i, j, 0, 1)            ' AB=2: TY-RB
                    If i < 5 Then
                        FA(i, j, 0, 3) = FA(i, j, 0, 0) - FA(i + 1, j, 0, 0)    ' AB=3: TY-LY
                    End If
                    FA(i, j, 0, 5) = FA(i, j, 0, 0) - FA(i, j, 0, 4)            ' AB=5: TY-OB
                    FA(i, j, 0, 6) = FA(i, j, 0, 1) - FA(i, j, 0, 4)            ' AB=6: RB-OB

                    ' If j = 14 And ASCMAIN1.Running_in_VS Then Stop

                    '' DO NOT DO %'AGES IN GET DETAILS - SUBTOTALS NEED %'AGES TOO - LEAVE NUMERATOR IN AND DO DENOMINATOR IN WRITE_DETAILS
                    '' 02/13/16 - CHANGED REMMED OUT CODE TO BUCKET 2 BUT DON'T THINK IT MATTERS
                    ''If FA(i, j, 0, 1) = 1 Then
                    'FA(i, j, 2, 2) = 100 * FA(i, j, 0, 2) / FA(i, j, 0, 1)      ' AB=2: TY-RB / RB
                    'If i < 5 Then
                    '    FA(i, j, 2, 3) = 100 * FA(i, j, 0, 3) / FA(i + 1, j, 0, 0)  ' AB=3: TY-LY / LY
                    'End If
                    'FA(i, j, 2, 5) = 100 * FA(i, j, 0, 5) / FA(i, j, 0, 4)      ' AB=5: TY-OB / OB
                    'FA(i, j, 2, 6) = 100 * FA(i, j, 0, 6) / FA(i, j, 0, 4)      ' AB=6: RB-OB / OB
                    ''End If

                    FA(i, j, 1, 2) = FA(i, j, 0, 2)
                    FA(i, j, 1, 3) = FA(i, j, 0, 3)
                    FA(i, j, 1, 5) = FA(i, j, 0, 5)
                    FA(i, j, 1, 6) = FA(i, j, 0, 6)
                Next j
            End If

        Next i

        Dim k1 As Integer
        Dim k2 As Integer

        For STMT_COL_NO As Integer = 1 To MAX_COLUMNS + MAX_COLUMNS
            If xy(STMT_COL_NO, 2) = 8 Then      ' Specified Month
                k1 = xy(STMT_COL_NO, 5)
                k2 = xy(STMT_COL_NO, 5)
                Calc_TOTAL_AMT(STMT_COL_NO, k1, k2, True)
            ElseIf xy(STMT_COL_NO, 2) = 9 Then  ' Specified Quarter
                k1 = (xy(STMT_COL_NO, 5) - 1) * 3 + 1
                k2 = (xy(STMT_COL_NO, 5) - 1) * 3 + 3
                Calc_TOTAL_AMT(STMT_COL_NO, k1, k2, True)
            ElseIf xy(STMT_COL_NO, 2) = 10 Then ' Specified Half
                k1 = (xy(STMT_COL_NO, 5) - 1) * 6 + 1
                k2 = (xy(STMT_COL_NO, 5) - 1) * 6 + 6
                Calc_TOTAL_AMT(STMT_COL_NO, k1, k2, True)
            Else
                rr(0, STMT_COL_NO) = FA(xy(STMT_COL_NO, 1), xy(STMT_COL_NO, 2), xy(STMT_COL_NO, 3), xy(STMT_COL_NO, 4))
            End If
        Next STMT_COL_NO

    End Sub

    Sub Calc_TOTAL_AMT(
    ByRef STMT_COL_NO As Integer,
    ByVal k1 As Integer,
    ByVal k2 As Integer,
    Optional ByVal is_for_specific_MQH As Boolean = False)

        ' note that we are not calculating properly for var% 
        ' - need another pot of coffee for that one

        Dim BALANCE_AMT As Double = 0
        Dim BALANCE_AMTZ As Double = 0

        For k As Integer = 0 To k1 - 1
            If xy(STMT_COL_NO, 4) = 2 Then
                BALANCE_AMTZ = A(xy(STMT_COL_NO, 1), k, 0) - A(xy(STMT_COL_NO, 1), k, 1)
            ElseIf xy(STMT_COL_NO, 4) = 3 Then
                BALANCE_AMTZ = A(xy(STMT_COL_NO, 1), k, 0) - A(xy(STMT_COL_NO, 1) - 1, k, 0)
            ElseIf xy(STMT_COL_NO, 4) = 5 Then
                BALANCE_AMTZ = A(xy(STMT_COL_NO, 1), k, 0) - A(xy(STMT_COL_NO, 1), k, 4)
            ElseIf xy(STMT_COL_NO, 4) = 6 Then
                BALANCE_AMTZ = A(xy(STMT_COL_NO, 1), k, 1) - A(xy(STMT_COL_NO, 1), k, 4)
            Else
                BALANCE_AMTZ = A(xy(STMT_COL_NO, 1), k, xy(STMT_COL_NO, 4))
            End If
            BALANCE_AMT = BALANCE_AMT + BALANCE_AMTZ
        Next k

        Dim TOTAL_AMT As Double = 0
        Dim TOTAL_AMTZ As Double = 0

        For k As Integer = k1 To k2
            If xy(STMT_COL_NO, 4) = 2 Then
                TOTAL_AMTZ = A(xy(STMT_COL_NO, 1), k, 0) - A(xy(STMT_COL_NO, 1), k, 1)
            ElseIf xy(STMT_COL_NO, 4) = 3 Then
                TOTAL_AMTZ = A(xy(STMT_COL_NO, 1), k, 0) - A(xy(STMT_COL_NO, 1) - 1, k, 0)
            ElseIf xy(STMT_COL_NO, 4) = 5 Then
                TOTAL_AMTZ = A(xy(STMT_COL_NO, 1), k, 0) - A(xy(STMT_COL_NO, 1), k, 4)
            ElseIf xy(STMT_COL_NO, 4) = 6 Then
                TOTAL_AMTZ = A(xy(STMT_COL_NO, 1), k, 1) - A(xy(STMT_COL_NO, 1), k, 4)
            Else
                TOTAL_AMTZ = A(xy(STMT_COL_NO, 1), k, xy(STMT_COL_NO, 4))
            End If
            TOTAL_AMT = TOTAL_AMT + TOTAL_AMTZ
        Next k

        If is_for_specific_MQH And HFs("STMT_TYPE") = "B" Then
            ' note that this adjustment to the month or quarter or half means that we have no way of seeing activity for a month, quarter or half on the balance sheet
            TOTAL_AMT = TOTAL_AMT + BALANCE_AMT
        End If

        rr(0, STMT_COL_NO) = TOTAL_AMT

    End Sub

    Sub Setup_SQL()

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
        SQLF(4) = " from " & TTC & " X"

        sqlA_select = ""
        sqlA_group_by = ""
        Dim z As String
        For i As Integer = 2 To 4
            Dim SEGX_CODE As String = "SEG" & CStr(i) & "_CODE"
            If Not New Boolean() {BY_SEG2, BY_SEG3, BY_SEG4}(i - 2) Then
                z = "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "'"
            Else
                z = "X." & SEGX_CODE
                sqlA_group_by = sqlA_group_by & ", " & z
            End If
            sqlA_select = sqlA_select & ", " & z & " " & SEGX_CODE
        Next
        For i As Integer = 2 To 4
            Dim SEGX_CODE As String = "SEG" & CStr(i) & "_CLASS_CODE"
            If Not New Boolean() {BY_SEG2_CLASS, BY_SEG3_CLASS, BY_SEG4_CLASS}(i - 2) Then
                z = "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "'"
            Else
                z = "X." & SEGX_CODE
                sqlA_group_by = sqlA_group_by & ", " & z
            End If
            sqlA_select = sqlA_select & ", " & z & " " & SEGX_CODE
        Next

        Dim sql As String = ""

        For Each C As String In New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE",
                                              "SEG2_CLASS_CODE", "SEG3_CLASS_CODE", "SEG4_CLASS_CODE", "ACCT_TYPE"}
            z = SQLA(C, "CODE_VALUES", True)
            If z <> "" Then
                sql &= " and GLTFINRD." & C & " " & IIf(SQLA(C, "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
            End If
        Next

        sqlA_where = sql
    End Sub

    Private Sub cmbSTMT_LAYOUT_CODE_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSTMT_LAYOUT_CODE.ValueChanged
        'Dim rowGLTCLAY1 As DataRow = LookUp("GLTCLAY1", cmbSTMT_LAYOUT_CODE.Text, True)
        'If rowGLTCLAY1.Item("RPT") & "" = "" Then
        '    grpReportVersion.Visible = True
        'Else
        '    grpReportVersion.Visible = False
        'End If
    End Sub

    Sub Excel_Extract(STMT_CODE As String, REPORT_NO As Integer,
                      workbook As SpreadsheetGear.IWorkbook,
                      description As String,
                      sheet_name As String)

        Dim Start_Row As Integer = 5
        Dim Start_Col As Integer = 2
        Dim Total_Cols As Integer = 10

        Dim Transactions As String = "Transactions" & IIf(REPORT_NO = 1, "", " " & CStr(REPORT_NO))

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        worksheet = workbook.Worksheets(workbook.Worksheets.Count - 1)

        Dim STMT_LAYOUT_DESC As String = Absx1.cmbFor("STMT_LAYOUT_CODE").ActiveRow.Cells("STMT_LAYOUT_DESC").Text
        If sheet_name = "" Then sheet_name = STMT_LAYOUT_DESC
        worksheet.Name = ASCMAIN1.Excel_Sheet_Name(sheet_name) ' REPORT_NO ' ASCMAIN1.Excel_Sheet_Name(STMT_LAYOUT_DESC)
        worksheet.Outline.SummaryRow = SpreadsheetGear.SummaryRow.Above

        Dim Rx As Integer = Start_Row

        ' Get DataTable Prepared

        Dim SQL As String = "STMT_CODE = '" & STMT_CODE & "' and REPORT_NO = " & CStr(REPORT_NO)
        Dim DVW As New DataView(dst.Tables("GLTSTMTE"), SQL, "STMT_LINE_NO,STMT_LINE_NO2", DataViewRowState.CurrentRows)
        Dim tbl As DataTable = DVW.ToTable


        ' Format 1st Columns as Textual

        Dim Col0 As Integer = Start_Col + tbl.Columns.IndexOf("AMT01") - 1

        Dim Col0_STMT_LINE_DESC As Integer = Start_Col + tbl.Columns.IndexOf("STMT_LINE_DESC")
        range = worksheet.Cells(Excel_Cell0(0, Col0_STMT_LINE_DESC)).EntireColumn
        range.ColumnWidth = range.ColumnWidth * 5

        Format_XLS_based_on_tbl(tbl, worksheet, Start_Col)

        'range = worksheet.Cells(Excel_Cell0(-1, 0) & ":" & Excel_Cell0(-1, Col0))
        'range.EntireColumn.NumberFormat = "@"


        ' Load Data in to Excel

        worksheet.Range(Excel_Cell0(Rx, Start_Col + 1 - 1) & ":" & Excel_Cell0(Rx, Start_Col + tbl.Columns.Count - 1)).CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.None)
        Rx += tbl.Rows.Count + 1


        ' Format

        With worksheet.Cells(1, Col0_STMT_LINE_DESC)
            .Value = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME")
            If ASCMAIN1.CLIENT = "AHA" Then
                .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            Else
                .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End If
        End With

        Dim REPORT_TITLE As String = Absx1.txtFor("REPORT_TITLE").Text
        If REPORT_TITLE = "" Then REPORT_TITLE = RPT_TITLE
        With worksheet.Cells(2, Col0_STMT_LINE_DESC)
            .Value = REPORT_TITLE
            .Font.Bold = True
            If ASCMAIN1.CLIENT = "AHA" Then
                .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            Else
                .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End If
            .Font.Color = SpreadsheetGear.Colors.Blue
        End With

        With worksheet.Cells(3, Col0_STMT_LINE_DESC)
            .Value = STMT_LAYOUT_DESC
            ' .Font.Size = 14
            .Font.Bold = True
            If ASCMAIN1.CLIENT = "AHA" Then
                .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            Else
                .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End If
            .Font.Color = SpreadsheetGear.Colors.Green
        End With

        worksheet.Cells(1, Col0_STMT_LINE_DESC + 1 + 4).Value = "'" & Format(Now, "MM/dd/yy HH:mm")
        worksheet.Cells(2, Col0_STMT_LINE_DESC + 1 + 4).Value = ASCMAIN1.USER_ID
        worksheet.Cells(3, Col0_STMT_LINE_DESC + 1 + 4).Value = "As Of " & Mid(ASCMAIN1.Get_Legend(RYP), 10, 6)
        worksheet.Cells(1, Col0_STMT_LINE_DESC + 1 + 4, 3, Col0_STMT_LINE_DESC + 1 + 4).HorizontalAlignment = SpreadsheetGear.HAlign.Right

        worksheet.Cells(0, Col0_STMT_LINE_DESC + 1 + 4).EntireRow.RowHeight = 2

        ' worksheet.Cells(4, Col0_STMT_LINE_DESC + 1 + 4).EntireRow.RowHeight = 2

        With worksheet.Cells(4, Col0_STMT_LINE_DESC)
            .AddComment(Join(Page0.ToArray, vbCrLf))
            ' .Font.Color = SpreadsheetGear.Colors.Red
            '.Value = description
            .Value = "'" & Format(Now, "MM/dd/yy HH:mm") & "   " & ASCMAIN1.USER_ID
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With

        With worksheet.Cells(1, Col0_STMT_LINE_DESC, 3, Col0_STMT_LINE_DESC + 1 + 4)
            .Interior.Color = SpreadsheetGear.Colors.AliceBlue
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        With worksheet.Cells(Start_Row, Col0_STMT_LINE_DESC, Start_Row, Col0_STMT_LINE_DESC + 4 + 3 + MAX_COLUMNS * 2)
            .Interior.Color = SpreadsheetGear.Colors.AliceBlue
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        Dim i As Integer = 0

        worksheet.Cells(Start_Row, Col0_STMT_LINE_DESC).Value = "Description"
        range = worksheet.Cells(Start_Row + 1, Col0_STMT_LINE_DESC + 1)
        range.Select()
        worksheet.WindowInfo.FreezePanes = True

        worksheet.Cells(Start_Row, Col0_STMT_LINE_DESC).EntireRow.RowHeight = 45

        Dim xlFs As New Dictionary(Of Integer, String)
        For Each j As Integer In STFs.Keys
            xlFs.Add(j, STFs(j))
        Next

        Dim STMT_LINE_REF_Rxs As New Dictionary(Of String, Integer) ' relating a Line Ref to a Worksheet Row
        Dim L2R As New Dictionary(Of Integer, Integer)

        Dim STMT_LINE_NO2_last = 0
        Dim STMT_LINE_NO2_last_Rx = -1

        Rx = Start_Row

        For Each row As DataRow In tbl.Select("", "STMT_LINE_NO,STMT_LINE_NO2")
            Dim STMT_LINE_NO As Integer = Val(row.Item("STMT_LINE_NO") & "")
            Dim STMT_LINE_NO2 As Integer = Val(row.Item("STMT_LINE_NO2") & "")
            If STMT_LINE_NO2 = 0 Then
                If STMT_LINE_NO2_last <> 0 Then

                    worksheet.Cells(CStr(STMT_LINE_NO2_last_Rx + 1) & ":" & Rx + 1).Rows.OutlineLevel = 2

                    'For C As Integer = 15 To 15
                    '    worksheet.Cells(Excel_Cell(STMT_LINE_NO2_last_Rx, C)).Formula = "=SUM(" & Excel_Cell(STMT_LINE_NO2_last_Rx + 1, C) & ":" & Excel_Cell(Rx + 1, C) & ")"
                    'Next

                    For i = 1 To MAX_COLUMNS
                        Dim Cx As Integer = Col0 + (i - 1) * 2
                        If HDGs(i) = "" Then
                        Else
                            Dim Fx As String = "=SUM(" & Excel_Cell(STMT_LINE_NO2_last_Rx + 1, Cx + 2) & ":" & Excel_Cell(Rx + 1, Cx + 2) & ")"
                            worksheet.Cells(STMT_LINE_NO2_last_Rx - 1, Cx + 1).Formula = Fx
                        End If
                    Next

                    STMT_LINE_NO2_last_Rx = -1
                    STMT_LINE_NO2_last = 0
                End If
            End If

            Dim rowGLTFINR2 As DataRow = dst.Tables("GLTFINR2").Rows.Find(New Object() {STMT_CODE, STMT_LINE_NO})
            Dim STMT_LINE_LEVEL As Integer = Val(rowGLTFINR2.Item("STMT_LINE_LEVEL") & "")
            Dim STMT_LINE_REF_SET As String = rowGLTFINR2.Item("STMT_LINE_REF_SET") & ""

            Rx += 1

            Dim STMT_LINE_TYPE As String = rowGLTFINR2.Item("STMT_LINE_TYPE") & ""
            Dim STMT_LINE_PRINT As String = rowGLTFINR2.Item("STMT_LINE_PRINT") & ""
            Dim STMT_BOLD_LINE As String = rowGLTFINR2.Item("STMT_BOLD_LINE") & ""
            Dim STMT_SKIP_LINE As String = rowGLTFINR2.Item("STMT_SKIP_LINE") & ""
            Dim STMT_DRAW_LINE As String = rowGLTFINR2.Item("STMT_DRAW_LINE") & ""
            Dim rowGLTFINRX As DataRow = dst.Tables("GLTFINRX").Rows.Find(New Object() {STMT_CODE, REPORT_NO, STMT_LINE_NO, STMT_LINE_NO2})
            Dim SUPPRESS_PRINT As String = rowGLTFINRX.Item("SUPPRESS_PRINT") & ""
            Dim STMT_LINE_DC As String = rowGLTFINR2.Item("STMT_LINE_DC") & ""
            Dim STMT_LINE_DESC_YELLOW As String = rowGLTFINR2.Item("STMT_LINE_DESC") & ""

            If STMT_LINE_TYPE = "H" Then
                worksheet.Cells(Rx, Col0 + (1 - 1) * 2, Rx, Col0 + (MAX_COLUMNS - 1) * 2).Clear()
                worksheet.Cells(Rx, Col0_STMT_LINE_DESC).Font.Bold = True
            Else
                If STMT_SKIP_LINE = "1" And STMT_LINE_NO2 = 0 Then
                    worksheet.Cells(Rx, 0).EntireRow.Insert(SpreadsheetGear.InsertShiftDirection.Down)
                    Rx += 1
                End If

                If STMT_LINE_NO2 = 0 Then
                    L2R.Add(STMT_LINE_NO, Rx)
                End If

                If STMT_LINE_REF_SET <> "" And STMT_LINE_REF_SET <> "0" Then
                    If Not STMT_LINE_REF_Rxs.ContainsKey(STMT_LINE_REF_SET) Then
                        STMT_LINE_REF_Rxs.Add(STMT_LINE_REF_SET, Rx)
                    End If
                End If

                If STMT_BOLD_LINE = "1" And STMT_LINE_NO2 = 0 Then
                    worksheet.Cells(Rx, Col0_STMT_LINE_DESC).EntireRow.Font.Bold = True
                End If
                If SUPPRESS_PRINT = "1" Then
                    worksheet.Cells(Rx, Col0_STMT_LINE_DESC).EntireRow.Hidden = True
                End If

                If ASCMAIN1.CLIENT = "VAN" Then
                    If (STMT_LINE_DESC_YELLOW = "Warehouse Expenses" Or STMT_LINE_DESC_YELLOW = "Sales and Design Expenses" Or STMT_LINE_DESC_YELLOW = "General/Admin Expenses" Or STMT_LINE_DESC_YELLOW = "Other Income & Expense" Or STMT_LINE_DESC_YELLOW = "Sales and Design Expenses") And STMT_LINE_TYPE = "D" And STMT_LINE_NO2 = "0" Then
                        With worksheet.Cells(Rx, Col0_STMT_LINE_DESC, Rx, Col0_STMT_LINE_DESC + 4 + 3 + MAX_COLUMNS * 2)
                            .Interior.Color = SpreadsheetGear.Colors.LightGoldenrodYellow
                            '.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                            '.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                            '.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        End With

                    End If

                End If
                If STMT_LINE_TYPE = "S" Then

                    Dim F As String = ""
                    If xlFs.ContainsKey(STMT_LINE_NO) Then
                        F = xlFs(STMT_LINE_NO)

                        If Not F.StartsWith("=") Then
                            ' Mid(F, 1, 1) = "="
                            F = "=" & F
                            Do While F.Contains("L")
                                Dim z As Integer = InStr(F, "L")
                                Dim l As Integer = Val(Mid(F, z + 1, 3))
                                Mid(F, z, 4) = Space(4) ' required, because if length of replacement string is less than 4, the remaining spaces will retain previous values
                                Mid(F, z, 4) = "A" & CStr(L2R(l) + 1)
                            Loop
                            F = Replace(F, " ", "")
                            If STMT_LINE_DC = "C" Then
                                F = Replace(F, "-", "_")
                                F = Replace(F, "+", "-")
                                F = Replace(F, "_", "+")
                            End If
                            xlFs(STMT_LINE_NO) = F
                        End If

                        For i = 1 To MAX_COLUMNS
                            Dim Cx As Integer = Col0 + (i - 1) * 2
                            If HDGs(i) = "" Then
                            Else
                                Dim Fx As String = Replace(F, "A", Excel_Cell0(-1, Cx + 1))
                                worksheet.Cells(Rx, Cx + 1).Formula = Fx
                            End If
                        Next

                    End If

                    With worksheet.Cells(Rx, Col0_STMT_LINE_DESC, Rx, Col0_STMT_LINE_DESC + 4 + 3 + MAX_COLUMNS * 2)
                        .Interior.Color = SpreadsheetGear.Colors.AliceBlue
                        '.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        '.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        '.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    End With

                End If
            End If

            worksheet.Cells(Rx, Col0_STMT_LINE_DESC).IndentLevel = STMT_LINE_LEVEL - 1 + IIf(STMT_LINE_NO2 = 0, 0, 4)

            If Absx1.chkFor("SHOW_ACCTS").Checked Then
                If STMT_LINE_NO2 <> 0 Then
                    STMT_LINE_NO2_last = STMT_LINE_NO2
                    If STMT_LINE_NO2_last_Rx = -1 Then STMT_LINE_NO2_last_Rx = Rx

                    Dim ACCT_CODE As String = row.Item("ACCT_CODE")

                    If Absx1.chkFor("SHOW_TRANS").Checked Then
                        Dim worksheetFS As SpreadsheetGear.IWorksheet = worksheet ' workbook.Worksheets(0)
                        worksheetFS.Hyperlinks.Add(worksheetFS.Cells(Rx, Col0_STMT_LINE_DESC + 1),
                                                    "",
                                                    "'" & Transactions & "'!AC_" & ACCT_CODE,
                                                    "Click Here to Navigate to Transactions",
                                                    "")
                    End If
                End If
            End If
        Next

        Dim RxMax As Integer = Rx

        ' Set % to Net Sales

        Rx = Start_Row

        For Each row As DataRow In tbl.Select("", "STMT_LINE_NO,STMT_LINE_NO2")

            Dim STMT_LINE_NO As Integer = Val(row.Item("STMT_LINE_NO") & "")
            Dim STMT_LINE_NO2 As Integer = Val(row.Item("STMT_LINE_NO2") & "")
            Dim rowGLTFINR2 As DataRow = dst.Tables("GLTFINR2").Rows.Find(New Object() {STMT_CODE, STMT_LINE_NO})

            Rx += 1

            Dim STMT_LINE_TYPE As String = rowGLTFINR2.Item("STMT_LINE_TYPE") & ""
            Dim STMT_SKIP_LINE As String = rowGLTFINR2.Item("STMT_SKIP_LINE") & ""

            If STMT_LINE_TYPE = "H" Then
            Else
                If STMT_SKIP_LINE = "1" And STMT_LINE_NO2 = 0 Then
                    Rx += 1
                End If

                Dim STMT_LINE_REF_PCT As String = rowGLTFINR2.Item("STMT_LINE_REF_PCT") & ""
                If STMT_LINE_REF_PCT <> "0" And STMT_LINE_REF_PCT <> "" Then
                    For i = 1 To MAX_COLUMNS
                        Dim Cx As Integer = Col0 + (i - 1) * 2
                        If HDGs(i) = "" Then
                        Else
                            Dim Rx_REF As Integer = STMT_LINE_REF_Rxs(STMT_LINE_REF_PCT)
                            worksheet.Cells(Rx, Cx + 2).Formula = "=IFERROR(" & Excel_Cell0(Rx, Cx + 1) & "/" & Excel_Cell0(Rx_REF, Cx + 1) & ",0)"
                        End If
                    Next
                End If
            End If
        Next



        range = worksheet.Cells(Excel_Cell0(-1, 1) & ":" & Excel_Cell0(-1, Col0_STMT_LINE_DESC - 1)).EntireColumn
        range.Hidden = True
        worksheet.Cells(Excel_Cell0(0, 0)).EntireColumn.ColumnWidth = 1

        i = 0
        For Each C As String In New String() {"SHOW_ACCTS", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "SEG2_CLASS_CODE", "SEG3_CLASS_CODE", "SEG4_CLASS_CODE"}
            i += 1
            range = worksheet.Cells(Excel_Cell0(Start_Row, Col0_STMT_LINE_DESC + i))
            If Not Absx1.chkFor(C).Checked Or Not Absx1.chkFor("SHOW_ACCTS").Checked Then
            Else
                If i = 1 Then
                    range.Value = "Account"
                Else
                    If i <= 4 Then
                        range.Value = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC")
                    Else
                        range.Value = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i - 3) & "_CLASS_DESC")
                    End If
                End If
            End If

            range.EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Left
            range.EntireColumn.NumberFormat = "@"
            range.EntireColumn.Hidden = Not Absx1.chkFor(C).Checked Or Not Absx1.chkFor("SHOW_ACCTS").Checked
        Next

        ' determine when to hide

        For i = 1 To MAX_COLUMNS
            Dim Cx As Integer = Col0 + (i - 1) * 2
            Dim rowGLTCLAY2 As DataRow = dst.Tables("GLTCLAY2").Rows.Find(New Object() {cmbSTMT_LAYOUT_CODE.Value, i})
            Dim STMT_CALC_CODE As String = ""
            If rowGLTCLAY2 IsNot Nothing Then
                STMT_CALC_CODE = rowGLTCLAY2.Item("STMT_CALC_CODE") & ""
            End If
            'Dim STMT_CALC_CODE As String = rowGLTCLAY1.Item("STMT_CALC_CODE_" & Format(i, "00")) & ""
            Dim rowGLTCALC1 As DataRow = LookUp("GLTCALC1", STMT_CALC_CODE)

            If HDGs(i) = "" Then
                range = worksheet.Cells(Excel_Cell0(-1, Cx + 1) & ":" & Excel_Cell0(-1, Cx + 2))
                range.Clear()
                range.EntireColumn.Hidden = True
            Else
                range = worksheet.Cells(Excel_Cell0(Start_Row, Cx + 1))
                range.Value = HDGs(i)
                If ASCMAIN1.CLIENT = "AHA" Then
                    range.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    range.EntireColumn.AutoFit()
                Else
                    range.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End If
                range.EntireColumn.ColumnWidth = 13
                range.EntireColumn.NumberFormat = "#,##0_);[Red](#,##0)" ' "#,##0"

                If Absx1.chkFor("SUPPRESS_ZERO_COLS").Checked Then
                    If dst.Tables("GLTFINRX").Select("AMT" & Format(i, "00") & " <> 0").Length = 0 Then
                        range.EntireColumn.Hidden = True
                    End If
                End If
                If ASCMAIN1.CLIENT = "AHA" Then
                    range.EntireColumn.AutoFit()

                    'If i = 1 Then
                    '    worksheet.Cells(Excel_Cell0(Start_Row + 1, Cx + 1) & ":" & Excel_Cell0(Rx, Cx + 1)).Interior.Color = SpreadsheetGear.Colors.Pink
                    'End If

                End If


                range = worksheet.Cells(Excel_Cell0(Start_Row, Cx + 2))
                range.Value = "%"
                If ASCMAIN1.CLIENT = "AHA" Then
                    range.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                Else
                    range.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End If
                range.EntireColumn.ColumnWidth = 8
                'range.EntireColumn.NumberFormatType = SpreadsheetGear.NumberFormatType.Percent
                range.EntireColumn.NumberFormat = "#0.0%"

                If Absx1.chkFor("SUPPRESS_ZERO_COLS").Checked Then
                    If dst.Tables("GLTFINRX").Select("AMT" & Format(i, "00") & " <> 0").Length = 0 Then
                        range.EntireColumn.Hidden = True
                    End If
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    range.EntireColumn.AutoFit()
                End If

                With worksheet.Cells(Start_Row, Cx + 1, RxMax, Cx + 2)
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With

                If rowGLTCALC1.Item("STMT_CALC_NO_REF_PCT") & "" = "1" Then
                    range = worksheet.Cells(Excel_Cell0(-1, Cx + 2) & ":" & Excel_Cell0(-1, Cx + 2))
                    range.Clear()
                    range.EntireColumn.Hidden = True
                End If

            End If
        Next

        If Absx1.chkFor("SHOW_ACCTS").Checked And Absx1.chkFor("SHOW_TRANS").Checked Then
            worksheet = workbook.Worksheets.Add
            worksheet.Name = Transactions
            ASCMAIN1.sql = "Select " _
                & "  GLTDETL1.OPS_YYYYPP" & vbCrLf _
                & ", GLTDETL1.JOURNAL_NO" & vbCrLf _
                & ", GLTJRNL1.JOURNAL_DESC" & vbCrLf _
                & ", GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
                & ", GLTDETL1.JOURNAL_LNO" & vbCrLf _
                & ", GLTDETL1.ACCT_CODE" & vbCrLf _
                & ", GLTDETL1.SEG2_CODE" & vbCrLf _
                & ", GLTDETL1.SEG3_CODE" & vbCrLf _
                & ", GLTDETL1.SEG4_CODE" & vbCrLf _
                & ", GLTDETL1.DETL_CTL_DATE" & vbCrLf _
                & ", GLTDETL1.DETL_CTL_NO" & vbCrLf _
                & ", GLTDETL1.DETL_CTL_LNO" & vbCrLf _
                & ", GLTDETL1.DETL_EXE_NO" & vbCrLf _
                & ", GLTDETL1.DETL_POSTING_AMT" & vbCrLf _
                & ", GLTDETL1.DETL_DESC" & vbCrLf _
                & ", GLTDETL1.DETL_EXP_CTL_NO" & vbCrLf _
                & ", GLTDETL1.DETL_CVX_NO" & vbCrLf _
                & ", GLTDETL1.DETL_CVX_REF_DATE" & vbCrLf _
                & ", GLTDETL1.DETL_CVX_REF_NO" & vbCrLf _
                & ", GLTDETL1.DETL_CVX_REF_LNO" & vbCrLf _
                & ", GLTDETL1.DETL_CTL_TYPE" & vbCrLf _
                & ", GLTDETL1.DETL_CVX_TYPE" & vbCrLf _
                & ", GLTJRNL1.JOURNAL_REVERSED" & vbCrLf _
                & ", GLTJRNL1.JOURNAL_REVERSED_IND" & vbCrLf _
                & ", GLTJRNL1.INIT_OPER" & vbCrLf _
                & ", GLTJRNL1.INIT_DATE" & vbCrLf _
                & ", GLTJRNL1.LAST_OPER" & vbCrLf _
                & ", GLTJRNL1.LAST_DATE" & vbCrLf _
                & ", GLTJRNL1.JOURNAL_COMMENT" & vbCrLf _
                & ", GLTSEGM2.ACCT_SEG_CLASS SEG2_CLASS_CODE" & vbCrLf _
                & ", GLTSEGM3.ACCT_SEG_CLASS SEG3_CLASS_CODE" & vbCrLf _
                & ", GLTSEGM4.ACCT_SEG_CLASS SEG4_CLASS_CODE" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1" & vbCrLf _
                & ", GLTSEGM1 GLTSEGM2, GLTSEGM1 GLTSEGM3, GLTSEGM1 GLTSEGM4" & vbCrLf _
                & " where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
                & "   and GLTDETL1.OPS_YYYYPP >= '" & Format(Val(Mid(RYP, 1, 4)) - 1, "0000") & "01'" & vbCrLf _
                & "   and GLTDETL1.OPS_YYYYPP <= '" & Mid(RYP, 1, 4) & "12'" & vbCrLf _
                & "   and GLTSEGM2.ACCT_SEG_ID (+) = '2' and GLTSEGM2.ACCT_SEG_CODE (+) = GLTDETL1.SEG2_CODE" & vbCrLf _
                & "   and GLTSEGM3.ACCT_SEG_ID (+) = '3' and GLTSEGM3.ACCT_SEG_CODE (+) = GLTDETL1.SEG3_CODE" & vbCrLf _
                & "   and GLTSEGM4.ACCT_SEG_ID (+) = '4' and GLTSEGM4.ACCT_SEG_CODE (+) = GLTDETL1.SEG4_CODE"

            Dim SQLW As String = ""
            Dim rowGLTSTMTX As DataRow = dst.Tables("GLTSTMTX").Rows.Find(New Object() {REPORT_NO})
            For S As Integer = 2 To 4
                Dim COLUMN_NAME As String = "SEG" & CStr(S) & "_CODE"
                Dim SEGX_CODE As String = rowGLTSTMTX.Item(COLUMN_NAME) & ""
                If SEGX_CODE <> "" Then
                    SQLW &= " and X." & COLUMN_NAME & " = '" & SEGX_CODE & "'"
                End If
                COLUMN_NAME = "SEG" & CStr(S) & "_CLASS_CODE"
                SEGX_CODE = rowGLTSTMTX.Item(COLUMN_NAME) & ""
                If SEGX_CODE <> "" Then
                    SQLW &= " and X." & COLUMN_NAME & " = '" & SEGX_CODE & "'"
                End If
            Next

            ASCMAIN1.sql = "Select X.* from (" & ASCMAIN1.sql & ") X " & ASCMAIN1.SQL_Add_WHERE(SQLW & Replace(sqlA_where, "GLTFINRD", "X"))

            tbl = ASCDATA1.GetDataTable
            tbl = New DataView(tbl, "", "ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,OPS_YYYYPP,JOURNAL_NO,JOURNAL_LNO", DataViewRowState.CurrentRows).ToTable

            Format_XLS_based_on_tbl(tbl, worksheet)

            ' Dump DataTable to Spreadsheet

            Dim Rx0 As Integer = 1
            worksheet.Range(Rx0, 0).CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.None)
            worksheet.Range(Rx0, 0, Rx0, tbl.Columns.Count - 1).EntireColumn.AutoFit()

            range = worksheet.Cells(Rx0 + 1, 0)
            range.Select()
            worksheet.WindowInfo.FreezePanes = True

            For C As Integer = 0 To tbl.Columns.Count - 1
                worksheet.Range(0, C).EntireColumn.AutoFilter()
            Next

            Dim ACCT_CODE_last As String = ""
            For Each row As DataRow In tbl.Select("", "ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,OPS_YYYYPP,JOURNAL_NO,JOURNAL_LNO")
                Dim ACCT_CODE As String = row.Item("ACCT_CODE")
                Rx0 += 1
                If ACCT_CODE <> ACCT_CODE_last Then
                    workbook.Names.Add("AC_" & ACCT_CODE, "='" & Transactions & "'!F" & CStr(Rx0 + 1 + 2))
                    ACCT_CODE_last = ACCT_CODE
                End If
            Next
        End If

        'Formulas - need to observe DRCR, and formula if accts shown

    End Sub

    Sub Format_XLS_based_on_tbl(tbl As DataTable, worksheet As SpreadsheetGear.IWorksheet, Optional C_offset As Integer = 0)

        For C As Integer = 0 To tbl.Columns.Count - 1
            Dim DC As DataColumn = tbl.Columns(C)
            If DC.DataType.Name = "String" Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "@"
            ElseIf DC.DataType.Name = "DateTime" Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "MM/DD/YY"
                worksheet.Range(0, C + C_offset).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
            ElseIf DC.DataType.Name = "Decimal" Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "#,##0.00"
                worksheet.Range(0, C + C_offset).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            ElseIf DC.DataType.Name.StartsWith("Int") Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "#,##0"
                worksheet.Range(0, C + C_offset).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End If
        Next
    End Sub

End Class