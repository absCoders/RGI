Public Class SORSLSD1

    Dim MD() As String
    Dim mos As Integer

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Absx1.chkFor("CHKSORTBYSEL").Checked = False
            Absx1.chkFor("CHKSORTBYSEL").Visible = True
            Absx1.chkFor("CHKSHOWBOTH").Checked = False
            Absx1.chkFor("CHKSHOWBOTH").Visible = True
            chkShowGPFilter.Visible = True
            chkShowGPFilter.Checked = False
            grpShowGPFilter.Visible = False
        Else
            Absx1.chkFor("CHKSORTBYSEL").Checked = False
            Absx1.chkFor("CHKSORTBYSEL").Visible = False
            Absx1.chkFor("CHKSHOWBOTH").Checked = False
            Absx1.chkFor("CHKSHOWBOTH").Visible = False
            chkShowGPFilter.Visible = False
            chkShowGPFilter.Checked = False
            grpShowGPFilter.Visible = False
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables
 
        mos = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1

        ReDim md(12)
        For i As Integer = 1 To 12
            Dim Z As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP0, (i - 1)))
            MD(i) = Mid$(z, 10, 6)
        Next i

        ' Prepare filters from Run-Time Options
  
        MyBase.Get_SQL("*")

        Dim sql_filter As String = ""

        If Absx1.optFor("OPTASN").Value = "S" Then
            sql_filter &= "   and ICTSTYL1.CUST_CODE is Null"
        End If
        If Absx1.optFor("OPTASN").Value = "N" Then
            sql_filter &= "   and ICTSTYL1.CUST_CODE is Not Null"
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ' TODO: when SOTINVH1 FIELD IS POPULATED PROPERLY FOR XFR, THEN USE SOTINVH1.ORDR_TYPE_CODE
            If Absx1.chkFor("CHKXSAMPLES").Checked Then
                sql_filter &= "   and SOTINVH2.CUST_CODE <> 'SAMPLES'"
            End If
            If Absx1.chkFor("CHKXTRANSF").Checked Then
                sql_filter &= "   and SOTINVH2.CUST_CODE <> 'TRANSFERS'"
            End If
        Else
            If Absx1.chkFor("CHKXSAMPLES").Checked Then
                sql_filter &= "   and SOTINVH1.ORDR_TYPE_CODE <> 'SAM'"
            End If
            If Absx1.chkFor("CHKXTRANSF").Checked Then
                sql_filter &= "   and SOTINVH1.ORDR_TYPE_CODE <> 'XFR'"
            End If
        End If

        Dim EXP As String = ""
        If Absx1.chkFor("CHKUNITS").Checked Then
            EXP = "NVL(SOTINVH2.ORDR_QTY_SHIP,0)"
        Else
            EXP = "NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)"
        End If

        For Each XY As String In New String() {"TY", "LY", "PY"}

            Dim YP As String = ""
            If XY = "TY" Then
                YP = RYP0
            ElseIf XY = "LY" Then
                YP = ASCMAIN1.Period_Calc(RYP0, -12)
            ElseIf XY = "PY" Then
                If chkShowGPFilter.Checked Then
                    YP = RYP0
                Else
                    YP = ASCMAIN1.Period_Calc(RYP0, -24)
                End If

            End If

            Dim sql_filter2 As String = " and SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN '" & YP & "'" & vbCrLf _
                                        & " and '" & ASCMAIN1.Period_Calc(YP, mos - 1) & "'" & vbCrLf

            Dim COLS As String = ""
            For I As Integer = 1 To mos
                COLS &= "," & XY & "_" & Format(I, "00")
            Next
            COLS &= "," & XY & "_TOT_COST"

            ' Prepare Work File with Data from Server

            Dim sql_Data As String = ""
            For i As Integer = 1 To mos
                Dim YPZ As String = ASCMAIN1.Period_Calc(YP, i - 1)
                sql_Data &= ", SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YPZ & "', " & EXP & ",0)) " & XY & "_" & Format$(i, "00")
            Next i
            sql_Data = sql_Data & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST" & ") " & XY & "_" & "TOT_COST"

            'sql = "Select " & sql_SELECT_cols & vbCrLf & sql_Data _
            '    & " from SOTINVH2" & sql_TABLE_NAMEs & vbCrLf _
            '    & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf _
            '    & " group by " & sql_GROUP_BY_cols
            Dim s As New Text.StringBuilder With {.Length = 0}
            s.AppendLine("Select " & sql_SELECT_cols)
            s.AppendLine(sql_Data)
            s.AppendLine("from SOTINVH2" & sql_TABLE_NAMEs)
            s.AppendLine(ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2))
            If chkShowGPFilter.Checked Then
                If XY <> "PY" Then
                    Dim GPFilterBeg As Double = numGPFilterBeg.Value / 100
                    Dim GPFilterEnd As Double = numGPFilterEnd.Value / 100
                    s.AppendLine(" AND (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) <> 0 ")
                    If chkGPFilterBeg.Checked Then
                        Dim txtStrBeg As String = ">"
                        If chkGPFilterBegEq.Checked Then
                            txtStrBeg += "="
                        End If
                        txtStrBeg += " "
                        s.AppendLine(String.Format(" AND ((SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) - (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST)) / (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) {0} {1}", txtStrBeg, GPFilterBeg))
                    End If
                    If chkGPFilterEnd.Checked Then
                        Dim txtStrEnd As String = "<"
                        If chkGPFilterEndEq.Checked Then
                            txtStrEnd += "="
                        End If
                        txtStrEnd += " "
                        s.AppendLine(String.Format(" AND ((SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) - (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST)) / (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) {0} {1}", txtStrEnd, GPFilterEnd))
                    End If
                End If
            End If
            s.AppendLine(" group by " & sql_GROUP_BY_cols)
            sql = s.ToString

            ASCMAIN1.sql = "Insert into " & ASTSRPT1 _
                & "(G1,G2,G3,G4,G5,G6,G7,G8,G9" _
                & COLUMN_NAMEs_appended _
                & COLS & ")" & vbCrLf _
                & " (" & sql & ")"

            ASCDATA1.ExecuteSQL()

        Next XY

        If Absx1.chkFor("CHKDEDUCT").Checked Then
            Dim ACCRUALS As Boolean = False
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                If chkUseAccrDed.Checked Then
                    ACCRUALS = True
                End If
            End If
            If ACCRUALS Then
                Dim CUST_CODE_exp As String = "Decode(ARTPYMT5.CUST_CODE_SO, NULL, X.CUST_CODE,ARTPYMT5.CUST_CODE_SO)"
                Dim SQLC As String = SQLA_filter("CUST_CODE", "", CUST_CODE_exp)
                Dim SQLC2 As String = SQLA_filter("CUST_CODE", "", "C1.CUST_CODE")

                Dim SQLD As String = ", Sum (Case When ARTPYMT1.OPS_YYYYPP between 'RYP0' and 'RYP1' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End)"
                Dim sd As New Text.StringBuilder With {.Length = 0}
                sd.AppendLine("SELECT CUST_CODE, G2, G3, G4, G5, G6, G7, G8, G9,")
                sd.AppendLine("SUM(TY_TOT_DED) AS TY_TOT_DED,")
                sd.AppendLine("SUM(LY_TOT_DED) AS LY_TOT_DED,")
                sd.AppendLine("SUM(PY_TOT_DED) AS PY_TOT_DED")
                sd.AppendLine("FROM")
                sd.AppendLine("(")
                sd.AppendLine("Select " & CUST_CODE_exp & " CUST_CODE, 'x' G2, 'x' G3, 'x' G4, 'x' G5, 'x' G6, 'x' G7, 'x' G8, 'x' G9")
                sd.AppendLine(Replace(Replace(SQLD, "RYP0", RYP0), "RYP1", RYP1) & " TY_TOT_DED")
                sd.AppendLine(Replace(Replace(SQLD, "RYP0", ASCMAIN1.Period_Calc(RYP0, -12)), "RYP1", ASCMAIN1.Period_Calc(RYP1, -12)) & " LY_TOT_DED")
                If chkShowGPFilter.Checked Then
                    sd.AppendLine(Replace(Replace(SQLD, "RYP0", RYP0), "RYP1", RYP1) & " PY_TOT_DED")
                Else
                    sd.AppendLine(Replace(Replace(SQLD, "RYP0", ASCMAIN1.Period_Calc(RYP0, -24)), "RYP1", ASCMAIN1.Period_Calc(RYP1, -24)) & " PY_TOT_DED")
                End If
                sd.AppendLine(" from ARTPYMT5, ARTPYMT2 X, ARTPYMT1")
                sd.AppendLine(" where NVL(ARTPYMT5.CHARGEBACK_IND,'0') <> '1' ")
                sd.AppendLine("   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO")
                sd.AppendLine("   and ARTPYMT5.PYMT_BATCH_NO = X.PYMT_BATCH_NO")
                sd.AppendLine("   and ARTPYMT5.PYMT_BATCH_LNO = X.PYMT_BATCH_LNO")
                sd.AppendLine("   and ARTPYMT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP0, -24) & "'")
                sd.AppendLine("   and ARTPYMT1.OPS_YYYYPP <= '" & RYP1 & "'")
                sd.AppendLine(SQLC)
                sd.AppendLine("AND (Decode(ARTPYMT5.CUST_CODE_SO, NULL, X.CUST_CODE,ARTPYMT5.CUST_CODE_SO), ARTPYMT5.REASON_CODE)")
                sd.AppendLine("NOT IN")
                sd.AppendLine("(")
                sd.AppendLine("SELECT C2.CUST_CODE, C2.REASON_CODE")
                sd.AppendLine("FROM ARTCRES1 C1, ARTCRES2 C2")
                sd.AppendLine("WHERE C1.CUST_CODE = C2.CUST_CODE")
                sd.AppendLine("AND C1.USE_DED_EST = '1'")
                sd.AppendLine(")")
                sd.AppendLine(" group by " & CUST_CODE_exp)
                sd.AppendLine("UNION")
                sd.AppendLine("Select C1.CUST_CODE, 'x' G2, 'x' G3, 'x' G4, 'x' G5, 'x' G6, 'x' G7, 'x' G8, 'x' G9")
                sd.AppendLine(", Sum (Case When C3.OPS_YYYYPP between '202108' and '202112' Then NVL(C3.TOT_DED,0) Else 0 End) TY_TOT_DED")
                sd.AppendLine(", Sum (Case When C3.OPS_YYYYPP between '202008' and '202012' Then NVL(C3.TOT_DED,0) Else 0 End) LY_TOT_DED")
                sd.AppendLine(", Sum (Case When C3.OPS_YYYYPP between '201908' and '201912' Then NVL(C3.TOT_DED,0) Else 0 End) PY_TOT_DED")
                sd.AppendLine("from ARTCRES1 C1, ARTCRES2 C2, ARTCRES3 C3")
                sd.AppendLine("WHERE C1.CUST_CODE = C2.CUST_CODE")
                sd.AppendLine("AND C2.CUST_CODE = C3.CUST_CODE")
                sd.AppendLine("AND C2.REASON_CODE = C3.REASON_CODE")
                sd.AppendLine("AND C1.USE_DED_EST = '1'")
                sd.AppendLine(SQLC2)
                sd.AppendLine("group by C1.CUST_CODE")
                sd.AppendLine(")")
                sd.AppendLine("GROUP BY CUST_CODE, G2, G3, G4, G5, G6, G7, G8, G9")
                sql = sd.ToString

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 _
                    & " (G1,G2,G3,G4,G5,G6,G7,G8,G9" _
                    & COLUMN_NAMEs_appended _
                    & ",TY_DED,LY_DED,PY_DED)" & vbCrLf _
                    & " (" & sql & ")"

                ASCDATA1.ExecuteSQL()
            Else
                Dim CUST_CODE_exp As String = "Decode(ARTPYMT5.CUST_CODE_SO, NULL, X.CUST_CODE,ARTPYMT5.CUST_CODE_SO)"
                Dim SQLC As String = SQLA_filter("CUST_CODE", "", CUST_CODE_exp)

                Dim SQLD As String = ", Sum (Case When ARTPYMT1.OPS_YYYYPP between 'RYP0' and 'RYP1' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End)"
                Dim sd As New Text.StringBuilder With {.Length = 0}
                sd.AppendLine("Select " & CUST_CODE_exp & " CUST_CODE, 'x' G2, 'x' G3, 'x' G4, 'x' G5, 'x' G6, 'x' G7, 'x' G8, 'x' G9")
                sd.AppendLine(Replace(Replace(SQLD, "RYP0", RYP0), "RYP1", RYP1) & " TY_TOT_DED")
                sd.AppendLine(Replace(Replace(SQLD, "RYP0", ASCMAIN1.Period_Calc(RYP0, -12)), "RYP1", ASCMAIN1.Period_Calc(RYP1, -12)) & " LY_TOT_DED")
                If chkShowGPFilter.Checked Then
                    sd.AppendLine(Replace(Replace(SQLD, "RYP0", RYP0), "RYP1", RYP1) & " PY_TOT_DED")
                Else
                    sd.AppendLine(Replace(Replace(SQLD, "RYP0", ASCMAIN1.Period_Calc(RYP0, -24)), "RYP1", ASCMAIN1.Period_Calc(RYP1, -24)) & " PY_TOT_DED")
                End If
                sd.AppendLine(" from ARTPYMT5, ARTPYMT2 X, ARTPYMT1")
                sd.AppendLine(" where NVL(ARTPYMT5.CHARGEBACK_IND,'0') <> '1' ")
                sd.AppendLine("   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO")
                sd.AppendLine("   and ARTPYMT5.PYMT_BATCH_NO = X.PYMT_BATCH_NO")
                sd.AppendLine("   and ARTPYMT5.PYMT_BATCH_LNO = X.PYMT_BATCH_LNO")
                sd.AppendLine("   and ARTPYMT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP0, -24) & "'")
                sd.AppendLine("   and ARTPYMT1.OPS_YYYYPP <= '" & RYP1 & "'")
                sd.AppendLine(SQLC)
                sd.AppendLine(" group by " & CUST_CODE_exp)
                sql = sd.ToString

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 _
                    & " (G1,G2,G3,G4,G5,G6,G7,G8,G9" _
                    & COLUMN_NAMEs_appended _
                    & ",TY_DED,LY_DED,PY_DED)" & vbCrLf _
                    & " (" & sql & ")"

                ASCDATA1.ExecuteSQL()
            End If

        End If

    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        ' Set up 12 month Total fields using ADO.Net expression

        Dim sqlT As String = ""
        For i As Integer = 1 To 12
            sqlT &= "+ISNULL(XY_" & Format(i, "00") & ",0)"
        Next i
        For Each XY As String In New String() {"TY", "LY", "PY"}
            dst.Tables("ASTSRPT1").Columns(XY & "_TOT").Expression = Replace(Mid(sqlT, 2), "XY", XY)
        Next


        ' Ranking Last Element Sorted within each Pre-Fix combination

        If COLUMN_NAMEs.Count > 1 Then

            Dim Gs As String = ""
            Dim Gst(0) As String
            If COLUMN_NAMEs.Count > 1 Then
                For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                    Gs &= "," & "G" & CStr(i)
                    ReDim Preserve Gst(i - 1)
                    Gst(i - 1) = "G" & CStr(i)
                Next
                Gs = Mid(Gs, 2)

            End If

            'For Each row As DataRow In ASCMAIN1.Distinct_Values("", dst.Tables("ASTSRPT1"), Gs).Select("", Gs)
            For Each row As DataRow In ASCMAIN1.Distinct_Values("", dst.Tables("ASTSRPT1"), Gst).Select("", Gs)

                Dim Gw As String = ""
                If COLUMN_NAMEs.Count > 1 Then
                    For i As Integer = 1 To COLUMN_NAMEs.Count - 1
                        If row.Item(i - 1) & "" = "" Then
                            Gw &= " and " & "G" & CStr(i) & " is Null'"
                        Else
                            Gw &= " and " & "G" & CStr(i) & " = '" & row.Item(i - 1) & "'"
                        End If
                    Next
                    Gw = Mid(Gw, 6)
                End If

                Rank(row, Gw)
            Next
        Else
            Rank()
        End If

        If Absx1.chkFor("CHKDEDUCT").Checked Then
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                If chkUseAccrDed.Checked Then
                    Dim sql As New Text.StringBuilder With {.Length = 0}
                    sql.AppendLine("SELECT C1.CUST_CODE")
                    sql.AppendLine("FROM ARTCRES1 C1")
                    sql.AppendLine("WHERE C1.USE_DED_EST = '1'")
                    Dim tblARTCRES1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())
                    If tblARTCRES1.Rows.Count > 0 Then
                        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Rows
                            Dim CUST_CODE As String = rowASTGROUP.Item("GROUP_CODE").ToString & String.Empty
                            Dim FLT As String = $"CUST_CODE = '{CUST_CODE}'"
                            Dim rowARTCRES1 As DataRow = tblARTCRES1.Select(FLT).FirstOrDefault
                            If Not IsNothing(rowARTCRES1) Then
                                rowASTGROUP.Item("GROUP_DESC") = "* " + rowASTGROUP.Item("GROUP_DESC").ToString & String.Empty
                            End If
                        Next
                    End If
                End If
            End If
        End If


        ' Delete if > Cut-Off

        If Val(Absx1.numFor("FPCUTOFF").Value & "") > 0 Then
            ASCDATA1.DeleteRows("ASTSRPT1", "RANK > " & Absx1.numFor("FPCUTOFF").Value)
        End If

    End Sub

    Sub Rank(Optional row As DataRow = Nothing, Optional Gw As String = "")

        Dim GX As String = "G" & CStr(COLUMN_NAMEs.Count)

        Dim RANK As Integer = 0
        For Each rowASTSPRT1 As DataRow In dst.Tables("ASTSRPT1").Select(Gw, "TY_TOT DESC")
            RANK += 1
            rowASTSPRT1.Item("RANK") = RANK

            Dim GX_VALUE As String = rowASTSPRT1.Item(GX)
            Dim rowASTGROUP As DataRow = dst.Tables("ASTGROUP").Rows.Find(New String() {GX_VALUE})

            If Not Absx1.chkFor("CHKSORTBYSEL").Checked Then
                GX_VALUE = Format(RANK, "000000") & GX_VALUE
            End If
            rowASTSPRT1.Item(GX) = GX_VALUE
            If dst.Tables("ASTGROUP").Rows.Find(GX_VALUE) Is Nothing Then
                dst.Tables("ASTGROUP").Rows.Add(New String() {GX_VALUE, rowASTGROUP.Item(1), rowASTGROUP.Item(2)})
            End If
        Next
    End Sub

    Public Overrides Sub Print_Report()

        If Absx1.chkFor("CHKSORTBYSEL").Checked Then
            RPT = "SORSLSDV"
        End If

        SUBT = txtDescription.Text
        SUBT &= " Ranked by YTD Sales, "

        If Absx1.chkFor("CHKDEDUCT").Checked Then
            SUBT &= " Showing Deductions, "
            CR_params.Add("DEDUCT", "1")
        Else
            CR_params.Add("DEDUCT", "")
        End If

        If Absx1.chkFor("CHKUNITS").Checked Then
            SUBT &= " Reporting In Units, "
            CR_params.Add("INUNITS", "1")
        Else
            CR_params.Add("INUNITS", "")
        End If


        Dim SMP As String = ""
        If Absx1.chkFor("CHKXSAMPLES").Checked Then
            SMP = " Excluding Samples"
        Else
            SMP = ""
        End If
        If Absx1.chkFor("CHKXTRANSF").Checked Then
            If SMP = "" Then
                SMP = " Excluding Transfers"
            Else
                SMP = " Excluding Samples & Transfers"
            End If
        End If

        Dim SGP As String = ""
        If chkShowGPFilter.Checked Then
            SGP = "Filtering GP "
            Dim txtStrBeg As String = ""
            If chkGPFilterBeg.Checked Then
                txtStrBeg += ">"
                If chkGPFilterBegEq.Checked Then
                    txtStrBeg += "="
                End If
                txtStrBeg += " " & numGPFilterBeg.Value
                SGP += txtStrBeg & " "
            End If
            Dim txtStrEnd As String = "<"
            If chkGPFilterEnd.Checked Then
                txtStrEnd = "<"
                If chkGPFilterEndEq.Checked Then
                    txtStrEnd += "="
                End If
                txtStrEnd += " " & numGPFilterEnd.Value
                SGP += txtStrEnd
            End If
            'SGP = String.Format("Filtering GP From {0} To {1}", numGPFilterBeg.Value, numGPFilterEnd.Value)
            If SUBT.Length > 0 Or SMP.Length > 0 Then
                SGP = ", " & SGP
            End If
            CR_params.Add("GPF", "TY TOT")
        Else
            CR_params.Add("GPF", "PY")
        End If

        SUBT &= SMP & SGP

        Dim ACCRUE As String = "0"
        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Dim STOCK_SUB As String = ""
            Select Case Absx1.optFor("OPTASN").Value
                Case "S"
                    STOCK_SUB = "Stock Only"
                Case "N"
                    STOCK_SUB = "Non-Stock"
                Case Else
                    STOCK_SUB = "All Stock"
            End Select
            If SUBT.Length = 0 Then
                SUBT = STOCK_SUB
            Else
                SUBT = SUBT & ", " & STOCK_SUB
            End If


            If Absx1.chkFor("CHKDEDUCT").Checked Then
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    If chkUseAccrDed.Checked Then
                        ACCRUE = "1"
                    End If
                End If
            End If
        End If
        CR_params.Add("ACCRUE", ACCRUE)

        For i As Integer = 1 To 12
            CR_params.Add("MD" & Format$(i, "00"), md(i))
        Next i
        CR_params.Add("MD1", md(1))
        CR_params.Add("MD2", md(mos))
        CR_params.Add("MOS", CStr(mos))

        Dim FPCUTOFF As Integer = Val(Absx1.numFor("FPCUTOFF").Value & "")

        If CStr(FPCUTOFF) <> 0 Then
            CR_params.Add("cut", CStr(FPCUTOFF))
        Else
            FPCUTOFF = "1000"
            CR_params.Add("cut", CStr(FPCUTOFF))
        End If

        CR_params.Add("THOUSANDS", "N")

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.chkFor("CHKDEDUCT").Checked Then
                Dim rows() As DataRow = tblASTDSQLA.Select("COLUMN_NAME <> 'CUST_CODE' and (ISNULL(SEQUENCE,0) <> 0 or ISNULL(CODE_VALUES,'')<>'')")
                If rows.Length <> 0 Then
                    EMsg = EMsg & vbCrLf & "If you want to see Deductions, you cannot sort or filter by " & rows(0).Item("COLUMN_CAPTION")
                End If
            End If
            If Absx1.chkFor("CHKSHOWBOTH").Checked Then
                EMsg = EMsg & vbCrLf & "The Option To Show Both  Units & Dollars Is Still Under Construction."
                Dim R0 As String = Absx1.cmbFor("RYP0").SelectedRow.Cells("OPS_YYYYPP").Value
                Dim R1 As String = Absx1.cmbFor("RYP1").SelectedRow.Cells("OPS_YYYYPP").Value
                mos = ASCMAIN1.Period_Diff(R0, R1) + 1
                If mos > 6 Then
                    EMsg = EMsg & vbCrLf & "You May Not Select More Than 6 Months When Showing Units & Dollars."
                End If
            End If
            If chkShowGPFilter.Checked Then
                If Absx1.chkFor("CHKSORTBYSEL").Checked Then
                    EMsg = EMsg & vbCrLf & "You May Not Filter GP and Sort By Sort & Filter On the Same Report"
                End If
                If chkGPFilterBeg.Checked = True And chkGPFilterEnd.Checked = True Then
                    If numGPFilterBeg.Value > numGPFilterEnd.Value Then
                        EMsg = EMsg & vbCrLf & "When Filtering GP FROM Must Be Less Than TO"
                    End If
                End If
                If chkGPFilterBeg.Checked = False And chkGPFilterEnd.Checked = False Then
                    EMsg = EMsg & vbCrLf & "You Must Pick Either Greater Than Or Less Than"
                End If
            End If
        End If
    End Sub

    Private Sub chkShowGPFilter_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowGPFilter.CheckedChanged
        If chkShowGPFilter.Checked Then
            grpShowGPFilter.Visible = True
            numGPFilterBeg.Value = 0
            numGPFilterEnd.Value = 100
        Else
            grpShowGPFilter.Visible = False
        End If
    End Sub

    Private Sub chkGPFilterBegEq_CheckedChanged(sender As Object, e As EventArgs) Handles chkGPFilterBegEq.CheckedChanged
        If chkGPFilterBegEq.Checked Then
            chkGPFilterBeg.Checked = True
        End If
    End Sub

    Private Sub chkGPFilterEndEq_CheckedChanged(sender As Object, e As EventArgs) Handles chkGPFilterEndEq.CheckedChanged
        If chkGPFilterEndEq.Checked Then
            chkGPFilterEnd.Checked = True
        End If
    End Sub

    Private Sub chkGPFilterBeg_CheckedChanged(sender As Object, e As EventArgs) Handles chkGPFilterBeg.CheckedChanged
        If chkGPFilterBeg.Checked = False Then
            chkGPFilterBegEq.Checked = False
        End If
    End Sub

    Private Sub chkGPFilterEnd_CheckedChanged(sender As Object, e As EventArgs) Handles chkGPFilterEnd.CheckedChanged
        If chkGPFilterEnd.Checked = False Then
            chkGPFilterEndEq.Checked = False
        End If
    End Sub

    Private Sub chkCHKDEDUCT_CheckedChanged(sender As Object, e As EventArgs) Handles chkCHKDEDUCT.CheckedChanged
        If chkCHKDEDUCT.Checked = True Then
            chkUseAccrDed.Visible = True
        Else
            chkUseAccrDed.Visible = False
            chkUseAccrDed.Checked = False
        End If
    End Sub
End Class