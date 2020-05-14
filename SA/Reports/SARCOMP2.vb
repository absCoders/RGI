Public Class SARCOMP2

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)

        Set_cmbYP("RYP0", Mid(ASCMAIN1.CYP, 1, 4) & "01", -60, 60, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0")

        ASCMAIN1.sql = "Select ASTDSQLA.COLUMN_NAME, NVL(ASTDSQLK.COLUMN_CAPTION,NVL(ASTDSQLA.COLUMN_CAPTION,ASTDSQLA.COLUMN_NAME)) COLUMN_CAPTION" _
        & " from ASTDSQLA,ASTDSQLK where ASTDSQLA.FORM_NAME = '" & Me.Name & "'" _
        & " and ASTDSQLK.COLUMN_NAME (+) = ASTDSQLA.COLUMN_NAME"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        cmbRANKBY.DataSource = DT
        cmbMIN.DataSource = DT

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        RYP = RYP1
        RYPLEGEND = RYPLEGEND1

        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -12)
        Dim LYP_NEXT As String = ASCMAIN1.Period_Calc(RYP, -12 + 1)
        Dim RYP_NEXT As String = ASCMAIN1.Period_Calc(RYP, +1)

        ' RYPLEGEND = ""

        Dim RYP_01 As String = RYP0 ' Mid(RYP, 1, 4) & "01"
        Dim LYP_01 As String = ASCMAIN1.Period_Calc(RYP0, -12) ' Mid(LYP, 1, 4) & "01"
        Dim LYP_12 As String = ASCMAIN1.Period_Calc(LYP_01, +11)  ' Mid(LYP, 1, 4) & "12"

        '    Dim SATCOMP2 As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        Dim rowASTDSQLC As DataRow = tblASTDSQLC.Rows.Find(New String() {"SARCOMP2", "*", "SREP_CODE"})
        tblASTDSQLC.Columns("TABLE_NAME").ReadOnly = False
        If optSREP.Value = "S" Then
            rowASTDSQLC.Item("TABLE_NAME") = "SOTINVH1"
        Else
            rowASTDSQLC.Item("TABLE_NAME") = "ARTCUST1"
        End If

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        Dim FX As String = ""
        Dim FACTOR As Decimal = 1
        If Absx1.chkFor("THOUSANDS").Checked Then
            FACTOR = 1000
        End If
        If FACTOR <> 1 Then
            FX = "/" & CStr(FACTOR)
        End If

        Dim sql_Data As String = ""
        Dim EXP_AMT As String = ""
        Dim EXP_GPD As String = ""


        ASCMAIN1.Progress("Shipments")
        MyBase.Get_SQL("*")

        Dim SOURCE_TABLE_NAME As String = "SOTINVH2"
        Dim YP As String = "SATCOMP2.ORDR_YYYYPP_UPDATED"


        EXP_AMT = "NVL(SATCOMP2.ORDR_QTY_SHIP,0) * NVL(SATCOMP2.ORDR_UNIT_PRICE,0)" & FX
        EXP_GPD = EXP_AMT & " - " & "NVL(SATCOMP2.ORDR_QTY_SHIP,0) * NVL(SATCOMP2.ORDR_UNIT_COST,0)" & FX

        Dim sql_MTD As String = "" _
        & ", Sum (CASE WHEN " & YP & " = 'XXXXXX' AND SATCOMP2.INV_TYPE = 'I' THEN " & EXP_AMT & " ELSE 0 END) XYMTDGRS" & vbCrLf _
        & ", Sum (CASE WHEN " & YP & " = 'XXXXXX' AND SATCOMP2.INV_TYPE = 'C' THEN " & EXP_AMT & " ELSE 0 END) XYMTDRTN" & vbCrLf _
        & ", Sum (CASE WHEN " & YP & " = 'XXXXXX' THEN " & EXP_AMT & " ELSE 0 END) XYMTDNET" & vbCrLf _
        & ", Sum (CASE WHEN " & YP & " = 'XXXXXX' THEN " & EXP_GPD & " ELSE 0 END) XYMTDGPD" & vbCrLf

        sql_Data &= "" _
            & Replace(Replace(sql_MTD, "XYMTD", "TYMTD"), "XXXXXX", RYP) _
            & Replace(Replace(sql_MTD, "XYMTD", "LYMON"), "XXXXXX", LYP)

        Dim sql_YTD As String = "" _
        & ", Sum (CASE WHEN " & YP & " Between 'XXXXX1' and 'XXXXX2' AND SATCOMP2.INV_TYPE = 'I' THEN " & EXP_AMT & " ELSE 0 END) XYYTDGRS" & vbCrLf _
        & ", Sum (CASE WHEN " & YP & " Between 'XXXXX1' and 'XXXXX2' AND SATCOMP2.INV_TYPE = 'C' THEN " & EXP_AMT & " ELSE 0 END) XYYTDRTN" & vbCrLf _
        & ", Sum (CASE WHEN " & YP & " Between 'XXXXX1' and 'XXXXX2' THEN " & EXP_AMT & " ELSE 0 END) XYYTDNET" & vbCrLf _
        & ", Sum (CASE WHEN " & YP & " Between 'XXXXX1' and 'XXXXX2' THEN " & EXP_GPD & " ELSE 0 END) XYYTDGPD" & vbCrLf

        sql_Data &= "" _
            & Replace(Replace(Replace(sql_YTD, "XYYTD", "TYYTD"), "XXXXX1", RYP_01), "XXXXX2", RYP) _
            & Replace(Replace(Replace(sql_YTD, "XYYTD", "LYYTD"), "XXXXX1", LYP_01), "XXXXX2", LYP) _
            & Replace(Replace(Replace(sql_YTD, "XYYTD", "LYTOT"), "XXXXX1", LYP_01), "XXXXX2", LYP_12)


        sql = "Select " & sql_SELECT_cols & vbCrLf & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " SATCOMP1 " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

       
        Dim sql_Col_X As String = ",TYMTDGRS,TYMTDRTN,TYMTDNET,TYMTDGPD"

        Dim sql_Cols As String = "" _
            & Replace(sql_Col_X, "TYMTD", "TYMTD") _
            & Replace(sql_Col_X, "TYMTD", "LYMON") _
            & Replace(sql_Col_X, "TYMTD", "TYYTD") _
            & Replace(sql_Col_X, "TYMTD", "LYYTD") _
            & Replace(sql_Col_X, "TYMTD", "LYTOT")

        sql_filter = "" _
            & " and " & YP & " BETWEEN '" & LYP_01 & "' AND '" & RYP & "'" & vbCrLf _
            & " and SATCOMP2.ORDR_QTY_SHIP <> 0"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTINVH2 SATCOMP2" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ' Eliminate 0s

        Dim sqlz As String = ""
        For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
            sqlz &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        Next
        ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlz))
    End Sub
     
    Public Overrides Sub Print_Report()
        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        CR_params.Add("YP_LEGEND", RYPLEGEND)
        Dim pp = Val(Mid(ASCMAIN1.CYP, 5, 2))
        SUBT = CStr(ASCMAIN1.Period_Diff(RYP0, RYP1) + 1) & " Months Ending " & RYPLEGEND
        If optRANKBY.Value <> "C" Then
            SUBT &= ", Ranked by " & optRANKBY.Text
        End If
        If chkYOY.Checked Then
            RPT = "SARCOMP4"
            CR_params.Add("GN", optRANKBYSALES.Text)
        Else
            RPT = "SARCOMP2"
        End If
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If Trim(ASCMAIN1.USER_CODES) = "FS" Then
            '    Dim SREP_CODE_filter As String = SQLA("SREP_CODE")
            '    If SREP_CODE_filter <> TAC.TACMAIN1.SREP_CODE Then
            '        EMsg &= vbCr & "You must leave Filter set to Sales Rep " & TAC.TACMAIN1.SREP_CODE & " only"
            '    Else
            '        Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("SREP_CODE")
            '        If rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
            '            EMsg &= vbCr & "You must leave Filter set INCLUDE to Sales Rep " & TAC.TACMAIN1.SREP_CODE & " only"
            '            EMsg &= vbCr & "(Nice Try)"
            '        End If
            '    End If
            'End If
            If optRANKBY.Value & "" <> "C" Then
                Dim COLUMN_NAME As String = cmbRANKBY.Value & ""
                If COLUMN_NAME = "" Then
                    EMsg &= vbCr & "You Must choose a Code to Rank by " & optRANKBY.Text
                Else
                    Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(COLUMN_NAME)
                    If rowASTDSQLA.Item("SEQUENCE") & "" = "" Then
                        EMsg &= vbCr & "You must choose a Code to Rank that is part of the Orginal Sort Seq"
                    End If
                End If
            End If

            If chkMIN.Checked Then
                Dim COLUMN_NAME As String = cmbMIN.Value & ""
                If COLUMN_NAME = "" Then
                    EMsg &= vbCr & "You Must choose a Code to Use for Minimums"
                Else
                    Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(COLUMN_NAME)
                    If rowASTDSQLA.Item("SEQUENCE") & "" = "" Then
                        EMsg &= vbCr & "You must choose a Code for Minimums that is part of the Orginal Sort Seq"
                    End If
                End If
            End If

            'RYPLEGEND0 = Absx1.cmbFor("RYP0", True).Value
            'RYP0 = Mid(RYPLEGEND0, 1, 4) & Mid(RYPLEGEND0, 6, 2)

            'RYPLEGEND1 = Absx1.cmbFor("RYP1", True).Value
            'RYP1 = Mid(RYPLEGEND1, 1, 4) & Mid(RYPLEGEND1, 6, 2)

            'If ASCMAIN1.Period_Diff(RYP0, RYP1) > 12 Then
            '    EMsg &= vbCr & "Period Range may not be more than 12 months"
            'End If
            'Stop

        End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        Dim GN As String = IIf(optRANKBYSALES.Value = "G", "GRS", "NET")
        Dim Y As String = IIf(optRANKBYTYLY.Value = "TY", "TY", "LY")

        If chkMIN.Checked Then
            Dim COLUMN_NAME_min As String = cmbMIN.Value
            Dim GMIN As String = ""
            For I As Integer = 1 To 7
                If COLUMN_NAMEs(I - 1) = COLUMN_NAME_min Then
                    GMIN = "G" & CStr(I)
                    Exit For
                End If
            Next

            Dim FX As String = ""
            Dim FACTOR As Decimal = 1
            If Absx1.chkFor("THOUSANDS").Checked Then
                FACTOR = 1000
            End If
            If FACTOR <> 1 Then
                FX = "/" & CStr(FACTOR)
            End If

            ASCMAIN1.sql = "Select " & GMIN _
                & " from " & ASTSRPT1 _
                & " group by " & GMIN _
                & " having SUM (NVL(" & Y & "YTD" & GN & ",0)) < " & CStr(numMIN.Value / FACTOR)
            Dim DT As DataTable = ASCDATA1.GetDataTable

            For Each row As DataRow In DT.Rows
                Dim GV As String = row.Item(0)
                ASCDATA1.DeleteRows(tblASTSRPT1, GMIN & "= '" & GV & "'")
            Next
        End If

        If optRANKBY.Value <> "C" Then
            Dim COLUMN_NAME_to_rank As String = cmbRANKBY.Value
            Dim COLUMN_NAME_to_rank_XY As String = IIf(optRANKBY.Value = "S", _
                                                       IIf(optRANKBYTYLY.Value = "TY", "TYSLS", "LYSLS"), _
                                                       IIf(optRANKBY.Value = "$", "YOYSLS", "YOYPCT"))

            Dim CP As Integer = 0
            For Each CN As String In COLUMN_NAMEs
                CP += 1
                If CN = COLUMN_NAME_to_rank Then
                    Exit For
                End If
            Next

            Dim C As String = ""
            If COLUMN_NAMEs.Count >= 1 Then
                C = Mid(G1thru9, 1, CP * 3 - 1)
            End If

            ASCMAIN1.sql = "Select " & C & vbCrLf _
                & ", TYSLS TYSLS" & vbCrLf _
                & ", LYSLS LYSLS" & vbCrLf _
                & ", TYSLS-LYSLS YOYSLS" & vbCrLf _
                & ", TRUNC(100 * CASE WHEN LYSLS = 0 THEN 0 ELSE 100 * (TYSLS-LYSLS) / LYSLS END) / 100 YOYPCT" & vbCrLf _
                & " from " & vbCrLf _
                & " (Select " & C & vbCrLf _
                & ", SUM (NVL(TYYTD" & GN & ",0)) TYSLS" & vbCrLf _
                & ", SUM (NVL(LYYTD" & GN & ",0)) LYSLS" & vbCrLf _
                & " from " & ASTSRPT1 & " group by " & C & ")" & vbCrLf _
                & " ORDER BY " & COLUMN_NAME_to_rank_XY & " DESC"
            Dim DT As DataTable = ASCDATA1.GetDataTable
            Dim GX As String = "G" & Format(CP, "0")
            Dim RANK As Integer = 0
            For Each row As DataRow In DT.Select("", COLUMN_NAME_to_rank_XY & " DESC")
                Dim sqlw As String = ""
                For I As Integer = 1 To CP
                    sqlw &= " and G" & CStr(I) & " = '" & row.Item(I - 1) & "'"
                Next

                RANK += 1
                Dim KEY_rank As String = Format(RANK, "000000" & "") & " " & row.Item(GX)

                For Each row2 As DataRow In tblASTSRPT1.Select(Mid(sqlw, 5))
                    row2.Item(GX) = KEY_rank
                Next

                Dim rowASTGROUP2 As DataRow = tblASTGROUP.Rows.Find(KEY_rank)
                If rowASTGROUP2 Is Nothing Then
                    Dim rowASTGROUP As DataRow = tblASTGROUP.Rows.Find(row.Item(GX))
                    rowASTGROUP2 = tblASTGROUP.NewRow
                    rowASTGROUP2.Item("GROUP_KEY") = KEY_rank
                    rowASTGROUP2.Item("GROUP_CODE") = rowASTGROUP.Item("GROUP_CODE")
                    rowASTGROUP2.Item("GROUP_DESC") = rowASTGROUP.Item("GROUP_DESC")
                    tblASTGROUP.Rows.Add(rowASTGROUP2)
                End If

            Next

        End If

        tblASTSRPT1.AcceptChanges()
        tblASTGROUP.AcceptChanges()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then
            'If Trim(ASCMAIN1.USER_CODES) = "FS" Then
            '    Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("SREP_CODE")
            '    rowASTDSQLA.Item("CODE_VALUES") = TAC.TACMAIN1.SREP_CODE
            'End If
        End If
    End Sub

    Private Sub optRANKBY_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optRANKBY.ValueChanged
        cmbRANKBY.Visible = (optRANKBY.Value <> "C")
    End Sub

    Private Sub chkMIN_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkMIN.CheckedChanged
        numMIN.Visible = chkMIN.Checked
        cmbMIN.Visible = chkMIN.Checked
    End Sub

    Private Sub optRANKBYSALES_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optRANKBYSALES.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Captions()
    End Sub

    Sub Set_Captions()
        Dim GN As String = IIf(optRANKBYSALES.Value = "G", "Gross", "Net")
        Dim Y As String = IIf(optRANKBYTYLY.Value = "TY", "TY", "LY")
        chkMIN.Text = "Minimum " & GN & " Sales " & Y

        Dim VL As ValueListItem = Nothing
        VL = optRANKBY.ValueList.ValueListItems(1)
        VL.DisplayText = "Y/Y " & GN & " Sales $"
        VL = optRANKBY.ValueList.ValueListItems(2)
        VL.DisplayText = "Y/Y " & GN & " Sales %"
        VL = optRANKBY.ValueList.ValueListItems(3)
        VL.DisplayText = GN & " Sales " & Y
    End Sub

    Private Sub optRANKBYTYLY_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optRANKBYTYLY.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Captions()
    End Sub
End Class