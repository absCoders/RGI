Public Class ICRWSUP1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        SUBT = "Showing All Open and Afloated Purchase Orders"
        If optShow.Value = "C" Then
            SUBT = "Cases"
        ElseIf optShow.Value = "U" Then
            SUBT = "Units"
        ElseIf optShow.Value = "A" Then
            SUBT = "Extended Purchase Cost"
        End If

        If chkNet.Checked Then
            SUBT &= ", Less Sales Commitments"
        End If

        SUBT &= ", based on Model Stock of " & CStr(numMODEL_STOCK.Value) & " Weeks of Supply"



        ' Extracts from Data Sources

        Dim sqlX As String = ""
        Dim sql_filter2 As String = ""

        Dim sqlONH As String = ""
        Dim sqlONH_FDA As String = ""
        Dim sqlONH_COM As String = ""

        ASCMAIN1.Progress("On Hand")
        Get_SQL("*")

        sqlONH_COM = "NVL(ICTLOTD2.QTY_COMMITTED,0)"
        sqlONH = "NVL(ICTLOTD2.QTY_ON_HAND,0)"
        'If chkNet.Checked Then
        '    sqlONH = "(" & sqlONH & "-" & sqlONH_COM & ")"
        'End If

        sqlX = ""
        If optShow.Value = "C" Then
        Else
            sqlX = " * DECODE(ICTLOTD2.PACK_CODE,'" & TAC.TACMAIN1.CATCH_PACK & "',NVL(ICTLOTD2.CATCH_WEIGHT,0),NVL(ICTPACK1.PACK_FACTOR,0))"
            If optShow.Value = "A" Then
                sqlX &= " * NVL(ICTLOTD2.PURCHASE_COST,0)"
            End If
        End If

        sql_filter2 = "" _
        & " and ICTPACK1.PACK_CODE = ICTLOTD2.PACK_CODE" & vbCrLf _
        & " and NVL(ICTLOTD2.ON_HOLD_FLAG,'M') <> 'R'" & vbCrLf

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", Sum (" & sqlONH & sqlX & ") ONH" & vbCrLf _
        & ", Sum (DECODE(NVL(ICTLOTD2.ON_HOLD_FLAG,'M'),'M',0," & sqlONH & sqlX & ")) ONH_FDA" & vbCrLf _
        & ", Sum (" & sqlONH_COM & sqlX & ") ONH_COM" & vbCrLf _
        & " from ICTLOTD2,ICTPACK1" & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                            & " (" & G1thru9 & ",ONH,ONH_FDA,ONH_COM" & ") " _
                            & " (" & sql & ")")


        Dim sqlQ As String = ""
        Dim sqlC As String = ""

        ASCMAIN1.Progress("On Order && Afloat")
        Get_SQL("P")

        sqlC = "NVL(POTORDR2.PO_CASES_PRESOLD,0)"
        sqlQ = "NVL(POTORDR2.PO_CASES,0)"
        'If chkNet.Checked Then
        '    sqlQ = "(" & sqlQ & "-" & sqlC & ")"
        'End If

        sqlX = ""
        If optShow.Value = "C" Then
        Else
            sqlC = Replace(sqlC, "_CASES", "_UNITS")
            sqlQ = Replace(sqlQ, "_CASES", "_UNITS")
            If optShow.Value = "A" Then
                sqlX &= " * NVL(POTORDR2.PURCHASE_COST,0)"
            End If
        End If

        sql_filter2 = "" _
        & " and POTORDR1.PO_STATUS_CODE = 'O'" & vbCrLf

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", Sum (DECODE(POTORDR1.IMPORT_NO,NULL,0," & sqlQ & sqlX & ")) AFL" & vbCrLf _
        & ", Sum (DECODE(POTORDR1.IMPORT_NO,NULL," & sqlQ & sqlX & ",0)) OPO" & vbCrLf _
        & ", Sum (DECODE(POTORDR1.IMPORT_NO,NULL,0," & sqlC & sqlX & ")) AFL_COM" & vbCrLf _
        & ", Sum (DECODE(POTORDR1.IMPORT_NO,NULL," & sqlC & sqlX & ",0)) OPO_COM" & vbCrLf _
        & " from POTORDR2 " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                            & " (" & G1thru9 & ",AFL,OPO,AFL_COM,OPO_COM" & ") " _
                            & " (" & sql & ")")



        ASCMAIN1.Progress("In Transit")
        Get_SQL("T")

        sqlC = "NVL(ICTTRANX.XFR_CASES_COM,0)"
        sqlQ = "NVL(ICTTRANX.TRANSFER_CASES,0)"
        'If chkNet.Checked Then
        '    sqlQ = "(" & sqlQ & "-" & sqlC & ")"
        'End If

        sqlX = ""
        If optShow.Value = "C" Then
        Else
            sqlC = Replace(sqlC, "_CASES", "_UNITS")
            sqlQ = Replace(sqlQ, "_CASES", "_UNITS")
            If optShow.Value = "A" Then
                sqlX &= " * NVL(ICTTRANX.PURCHASE_COST,0)"
            End If
        End If

        sql_filter2 = "" _
        & " and ICTLOTD1.WHSE_CODE = ICTTRANX.WHSE_CODE" & vbCrLf _
        & " and ICTLOTD1.LOT_NO = ICTTRANX.LOT_NO" & vbCrLf _
        & " and ICTLOTD1.LOT_SEQ_NO = ICTTRANX.LOT_SEQ_NO" & vbCrLf

        Dim sql_subquery As String = "Select ICTLOTD1.*,ICTTRAN1.WHSE_CODE_TO" & vbCrLf _
        & ",ICTTRAN2.TRANSFER_CASES, ICTTRAN2.XFR_CASES_COM" & vbCrLf _
        & ",ICTTRAN2.TRANSFER_UNITS, ICTTRAN2.XFR_UNITS_COM" & vbCrLf _
        & " from ICTTRAN1, ICTTRAN2, ICTLOTD1, ICTPACK1" _
        & " where ICTTRAN1.TRANSFER_NO = ICTTRAN2.TRANSFER_NO" & vbCrLf _
        & "   and ICTLOTD1.WHSE_CODE = ICTTRAN2.WHSE_CODE" & vbCrLf _
        & "   and ICTLOTD1.LOT_NO = ICTTRAN2.LOT_NO" & vbCrLf _
        & "   and ICTLOTD1.LOT_SEQ_NO = ICTTRAN2.LOT_SEQ_NO" & vbCrLf _
        & "   and ICTPACK1.PACK_CODE = ICTLOTD1.PACK_CODE" & vbCrLf _
        & "   and ICTTRAN1.TRANSFER_STATUS = 'O' AND ICTTRAN2.TRANSFER_LNO_STATUS = 'O'" & vbCrLf

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", Sum (" & sqlQ & sqlX & ") XIT" & vbCrLf _
        & ", Sum (" & sqlC & sqlX & ") XIT_COM" & vbCrLf _
        & " from (" & sql_subquery & ") ICTTRANX, ICTLOTD1" & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                            & " (" & G1thru9 & ",XIT,XIT_COM" & ") " _
                            & " (" & sql & ")")


        ASCMAIN1.Progress("Historical On Hand")
        Get_SQL("X")

        sqlONH = "NVL(ICTLOTD4.BAL_CASES,0)"

        sqlX = ""
        If optShow.Value = "C" Then
        Else
            sqlONH = Replace(sqlONH, "CASES", "UNITS")
            If optShow.Value = "A" Then
                sqlX &= " * NVL(ICTLOTD1.PURCHASE_COST,0)"
            End If
        End If

        sql_filter2 = "" _
        & " and ICTLOTD1.WHSE_CODE = ICTLOTD4.WHSE_CODE" & vbCrLf _
        & " and ICTLOTD1.LOT_NO = ICTLOTD4.LOT_NO" & vbCrLf _
        & " and ICTLOTD1.LOT_SEQ_NO = ICTLOTD4.LOT_SEQ_NO" & vbCrLf _
        & " and NVL(ICTLOTD1.ON_HOLD_FLAG,'M') <> 'R'" & vbCrLf _
        & " and ICTLOTD4.TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + 1), "dd-MMM-yyyy") & "'" & vbCrLf

        sql = "Select " & sql_SELECT_cols & ", ICTLOTD4.TRAN_DATE, SUM (" & sqlONH & sqlX & ") ONH" & vbCrLf _
        & " from ICTLOTD4,ICTLOTD1 " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols & ", ICTLOTD4.TRAN_DATE"

        Dim sql_COLUMN_NAMEs As String = ""
        For Each COLUMN_NAME In COLUMN_NAMEs
            sql_COLUMN_NAMEs &= "," & COLUMN_NAME
        Next
        sql_COLUMN_NAMEs = Mid(sql_COLUMN_NAMEs, 2)
        Dim sql_COLUMN_NAMEs_gby As String = sql_COLUMN_NAMEs

        Dim gx() As String = Split(sql_SELECT_cols, ",")

        For i As Integer = COLUMN_NAMEs.Count + 1 To 9
            sql_COLUMN_NAMEs &= "," & gx(i - 1)
        Next
        sql = "Select " & sql_COLUMN_NAMEs & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (4 * 7) + 1), "dd-MMM-yyyy") & "' THEN 1 ELSE 0 END) DAYS_04" & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (4 * 7) + 1), "dd-MMM-yyyy") & "' THEN ONH ELSE 0 END) / (4 * 7) INVTY_04" & vbCrLf _
        & ", MIN (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (4 * 7) + 1), "dd-MMM-yyyy") & "' THEN ONH ELSE NULL END) LOW_04" & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (12 * 7) + 1), "dd-MMM-yyyy") & "' THEN 1 ELSE 0 END) DAYS_12" & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (12 * 7) + 1), "dd-MMM-yyyy") & "' THEN ONH ELSE 0 END) / (12 * 7) INVTY_12" & vbCrLf _
        & ", MIN (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (12 * 7) + 1), "dd-MMM-yyyy") & "' THEN ONH ELSE NULL END) LOW_12" & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + 1), "dd-MMM-yyyy") & "' THEN 1 ELSE 0 END) DAYS_52" & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + 1), "dd-MMM-yyyy") & "' THEN ONH ELSE 0 END) / (52 * 7) INVTY_52" & vbCrLf _
        & ", MIN (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + 1), "dd-MMM-yyyy") & "' THEN ONH ELSE NULL END) LOW_52" & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + 1), "dd-MMM-yyyy") & "' and TRAN_DATE <= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + (12 * 7)), "dd-MMM-yyyy") & "' THEN 1 ELSE 0 END) DAYS_N12" & vbCrLf _
        & ", SUM (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + 1), "dd-MMM-yyyy") & "' and TRAN_DATE <= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + (12 * 7)), "dd-MMM-yyyy") & "' THEN ONH ELSE 0 END) / (12 * 7) INVTY_N12" & vbCrLf _
        & ", MIN (CASE WHEN TRAN_DATE >= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + 1), "dd-MMM-yyyy") & "' and TRAN_DATE <= '" & Format(DATETIME_STAMP.AddDays(-1 * (52 * 7) + (12 * 7)), "dd-MMM-yyyy") & "' THEN ONH ELSE NULL END) LOW_N12" & vbCrLf _
        & " from (" & sql & ") X" & vbCrLf _
        & " group by " & sql_COLUMN_NAMEs_gby

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                            & " (" & G1thru9 & ",DAYS_04,INVTY_04,LOW_04,DAYS_12,INVTY_12,LOW_12,DAYS_52,INVTY_52,LOW_52,DAYS_N12,INVTY_N12,LOW_N12" & ") " _
                            & " (" & sql & ")")





        Dim sqlSALES As String

        ASCMAIN1.Progress("Sales History")
        Get_SQL("S")

        sqlSALES = "NVL(SOTINVH0.QTY_CASES,0)"

        sqlX = ""
        If optShow.Value = "C" Then
        Else
            sqlSALES = Replace(sqlSALES, "_CASES", "_UNITS")
            If optShow.Value = "A" Then
                sqlX &= " * NVL(ICTLOTD1.PURCHASE_COST,0)"
            End If
        End If

        sql_filter2 = "" _
        & " and ICTLOTD1.WHSE_CODE = SOTINVH0.WHSE_CODE" & vbCrLf _
        & " and ICTLOTD1.LOT_NO = SOTINVH0.LOT_NO" & vbCrLf _
        & " and ICTLOTD1.LOT_SEQ_NO = SOTINVH0.LOT_SEQ_NO" & vbCrLf _
        & " and SOTINVH0.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'" & vbCrLf

        Dim DT04 As String = Format(Now.AddDays(-7 * 4 + 1), "dd-MMM-yyyy")
        Dim DT12 As String = Format(Now.AddDays(-7 * 12 + 1), "dd-MMM-yyyy")
        Dim DT52 As String = Format(Now.AddDays(-7 * 52 + 1), "dd-MMM-yyyy")

        Dim DTN12_START As String = Format(Now.AddDays(-7 * 52 + 1), "dd-MMM-yyyy")
        Dim DTN12_END As String = Format(Now.AddDays(-7 * 52 + 7 * 12), "dd-MMM-yyyy")

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", SUM (CASE WHEN SOTINVH0.ORDR_INV_DATE >= '" & DT04 & "' THEN " & sqlSALES & " ELSE 0 END) / 4 SALES_04" & vbCrLf _
        & ", SUM (CASE WHEN SOTINVH0.ORDR_INV_DATE >= '" & DT12 & "' THEN " & sqlSALES & " ELSE 0 END) / 12 SALES_12" & vbCrLf _
        & ", SUM (CASE WHEN SOTINVH0.ORDR_INV_DATE >= '" & DT52 & "' THEN " & sqlSALES & " ELSE 0 END) / 52 SALES_52" & vbCrLf _
        & ", SUM (CASE WHEN SOTINVH0.ORDR_INV_DATE >= '" & DTN12_START & "' AND SOTINVH0.ORDR_INV_DATE <= '" & DTN12_END & "' THEN " & sqlSALES & " ELSE 0 END) / 12 SALES_N12" & vbCrLf _
        & " from SOTINVH0, ICTLOTD1" & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & sql_filter2) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                            & " (" & G1thru9 & ",SALES_04,SALES_12,SALES_52,SALES_N12" & ") " _
                            & " (" & sql & ")")

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("TOTAL_DESC", optHAP.Text & IIf(chkNet.Checked, ", less Sales Committments", ""))
        CR_params.Add("SET_DESC", txtDescription.Text) '  SET_DESC & "")
        CR_params.Add("MODEL_STOCK", CStr(numMODEL_STOCK.Value))
        CR_params.Add("HAP", optHAP.Value)
        CR_params.Add("LESS_COM", IIf(chkNet.Checked, "1", "0"))

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
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

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShow.ValueChanged

        If SELECTION_NO = 0 Then Exit Sub

    End Sub
End Class