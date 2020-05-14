Public Class SORGOPN1
    Dim ICTCOSTX As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        RWU = "N"
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim SQLMain As New System.Text.StringBuilder() With {.Length = 0}

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = "For Period " & RYPLegend

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")

        Dim FYP As String = Mid(RYP, 1, 4) & "01"
        FYP = RYP

        Dim sql_filter2 As String = ""
        'ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf _
        '    & ", ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
        '    & ", DECODE (ICTSTAT1.OPS_YYYYPP,'" & FYP & "',ICTSTAT1.WHSE_QTY_BEG,0) WHSE_QTY_BEG" & vbCrLf _
        '    & ", ICTSTAT1.WHSE_QTY_SHP, ICTSTAT1.WHSE_QTY_RTN " & vbCrLf _
        '    & ", ICTSTAT1.WHSE_QTY_REC, ICTSTAT1.WHSE_QTY_ADJ, ICTSTAT1.WHSE_QTY_XFR " & vbCrLf _
        '    & ", ICTSTAT1.WHSE_QTY_PHY " & vbCrLf

        'ASCMAIN1.sql &= "" _
        '    & ", ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_ON_ORDER " & vbCrLf _
        '    & ", ICTSTAT2.WHSE_QTY_TRAN, ICTSTAT2.WHSE_QTY_OPEN" & vbCrLf _
        '    & ", ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO" & vbCrLf
        'If RYP = ASCMAIN1.CYP Then
        '    ASCMAIN1.sql &= " from ICTSTAT2 ICTSTAT2, ICTSTAT1 "
        'Else
        '    ASCMAIN1.sql &= " from ICTSTAT5 ICTSTAT2, ICTSTAT1"
        'End If
        'ASCMAIN1.sql &= sql_TABLE_NAMEs & vbCrLf
        'ASCMAIN1.sql &= " where ICTSTAT1.OPS_YYYYPP (+) = '" & RYP & "'" & vbCrLf
        'If RYP <> ASCMAIN1.CYP Then
        '    ASCMAIN1.sql &= "" _
        '        & "   and ICTSTAT2.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
        '        & "   and ICTSTAT2.OPS_YYYYPP = ICTSTAT1.OPS_YYYYPP (+) " & vbCrLf
        'End If
        'ASCMAIN1.sql &= sql_WHERE & sql_JOIN & sql_filter & sql_filter2 & vbCrLf
        'ASCMAIN1.sql &= "" _
        '    & "   and ICTSTAT2.STYLE_CODE = ICTSTAT1.STYLE_CODE (+) " & vbCrLf _
        '    & "   and ICTSTAT2.COLOR_CODE = ICTSTAT1.COLOR_CODE (+) " & vbCrLf _
        '    & "   and ICTSTAT2.WHSE_CODE = ICTSTAT1.WHSE_CODE (+) " & vbCrLf _
        '    & "   and ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE"
        'If Absx1.chkFor("CHKNEG").Checked = "1" Then
        '    'ASCMAIN1.sql &= "   and ICTSTAT2.WHSE_QTY_ON_HAND < 0"
        '    'ICTSTAT2.WHSE_ON_HAND-ICTSTAT2.WHSE_OPEN-ICTSTAT2.WHSE_PICK
        '    ASCMAIN1.sql &= "   and (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0)) > 0"
        'End If
        'If Absx1.optFor("OPTASN").Value = "S" Then
        '    ASCMAIN1.sql &= "   and ICTSTYL1.CUST_CODE is Null"
        'ElseIf Absx1.optFor("OPTASN").Value = "N" Then
        '    ASCMAIN1.sql &= "   and ICTSTYL1.CUST_CODE is Not Null"
        'End If

        SQLMain.AppendLine("Select " & sql_SELECT_cols)
        SQLMain.AppendLine(", ICTSTAT2.STYLE_CODE")
        SQLMain.AppendLine(", ICTSTAT2.COLOR_CODE")
        SQLMain.AppendLine(String.Format(", SUM(DECODE (ICTSTAT1.OPS_YYYYPP,'{0}',ICTSTAT1.WHSE_QTY_BEG,0)) WHSE_QTY_BEG", FYP))
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT1.WHSE_QTY_SHP,0)) WHSE_QTY_SHP")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT1.WHSE_QTY_RTN,0)) WHSE_QTY_RTN")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT1.WHSE_QTY_REC,0)) WHSE_QTY_REC")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT1.WHSE_QTY_ADJ,0)) WHSE_QTY_ADJ")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT1.WHSE_QTY_XFR,0)) WHSE_QTY_XFR")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT1.WHSE_QTY_PHY,0)) WHSE_QTY_PHY")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) WHSE_QTY_ON_HAND")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) WHSE_QTY_ON_ORDER")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0)) WHSE_QTY_TRAN")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT2.WHSE_QTY_OPEN,0)) WHSE_QTY_OPEN")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT2.WHSE_QTY_PICK,0)) WHSE_QTY_PICK")
        SQLMain.AppendLine(", SUM(NVL(ICTSTAT2.WHSE_QTY_ALLO,0)) WHSE_QTY_ALLO")
        If RYP = ASCMAIN1.CYP Then
            SQLMain.AppendLine(" from ICTSTAT2 ICTSTAT2, ICTSTAT1")
        Else
            SQLMain.AppendLine(" from ICTSTAT5 ICTSTAT2, ICTSTAT1")
        End If
        SQLMain.AppendLine(sql_TABLE_NAMEs)
        SQLMain.AppendLine(String.Format("where ICTSTAT1.OPS_YYYYPP (+) = '{0}'", RYP))
        If RYP <> ASCMAIN1.CYP Then
            SQLMain.AppendLine(String.Format("and ICTSTAT2.OPS_YYYYPP = '{0}'", RYP))
            SQLMain.AppendLine("and ICTSTAT2.OPS_YYYYPP = ICTSTAT1.OPS_YYYYPP (+) ")
        End If
        SQLMain.AppendLine(sql_WHERE & sql_JOIN & sql_filter & sql_filter2)
        SQLMain.AppendLine("and ICTSTAT2.STYLE_CODE = ICTSTAT1.STYLE_CODE (+)")
        SQLMain.AppendLine("and ICTSTAT2.COLOR_CODE = ICTSTAT1.COLOR_CODE (+) ")
        SQLMain.AppendLine("and ICTSTAT2.WHSE_CODE = ICTSTAT1.WHSE_CODE (+) ")
        SQLMain.AppendLine("and ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
        If Absx1.optFor("OPTASN").Value = "S" Then
            SQLMain.AppendLine("and ICTSTYL1.CUST_CODE is Null")
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            SQLMain.AppendLine("and ICTSTYL1.CUST_CODE is Not Null")
        End If
        If Absx1.chkFor("CHKNEG").Checked = "1" Then
            SQLMain.AppendLine("Having Sum((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0))) > 0")
        End If
        If sql_GROUP_BY_cols.Length = 0 Then
            SQLMain.AppendLine("Group By ICTSTAT2.STYLE_CODE")
        Else
            SQLMain.AppendLine("Group By " & sql_GROUP_BY_cols)
            SQLMain.AppendLine(", ICTSTAT2.STYLE_CODE")
        End If
        SQLMain.AppendLine(", ICTSTAT2.COLOR_CODE")

        Dim SQLi As New System.Text.StringBuilder() With {.Length = 0}
        SQLi.AppendLine("Insert into " & ASTSRPT1)
        SQLi.AppendLine(" (" & G1thru9)
        SQLi.AppendLine(",STYLE_CODE,COLOR_CODE")
        SQLi.AppendLine(",WHSE_BEG,WHSE_SHP,WHSE_RTN,WHSE_REC,WHSE_ADJ,WHSE_XFR,WHSE_PHY")
        SQLi.AppendLine(",WHSE_ON_HAND,WHSE_ON_ORDER,WHSE_TRAN,WHSE_OPEN,WHSE_PICK,WHSE_ALLO")
        SQLi.AppendLine(") ")
        SQLi.AppendLine(String.Format(" ({0})", SQLMain))

        ASCDATA1.ExecuteSQL(SQLi.ToString)

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)

        'Suppress Style/Colors w/All 0's
        ASCMAIN1.sql = ""
        For Each COLUMN_NAME As String In New String() _
            {"WHSE_BEG", "WHSE_SHP", "WHSE_RTN", "WHSE_REC", "WHSE_ADJ", "WHSE_XFR", "WHSE_PHY", _
             "WHSE_ON_HAND", "WHSE_ON_ORDER", "WHSE_TRAN", "WHSE_OPEN", "WHSE_PICK", "WHSE_ALLO"}
            ASCMAIN1.sql &= " and NVL(" & COLUMN_NAME & ",0) = 0"
        Next
        ASCMAIN1.sql = "Delete from " & TT & ASCMAIN1.SQL_Add_WHERE(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL()

        If Absx1.chkFor("CHKOTSMIN").Checked Then
            ASCMAIN1.sql = "Delete from " & TT & " where NVL(WHSE_ON_HAND,0)" _
                            & " - NVL(WHSE_PICK,0)" _
                            & " + NVL(WHSE_ON_ORDER,0)" _
                            & " + NVL(WHSE_TRAN,0)" _
                            & " - NVL(WHSE_OPEN,0) < " & Absx1.numFor("NUMOTSMIN").Value
            ASCDATA1.ExecuteSQL()
        End If
        If Absx1.chkFor("CHKOTSMAX").Checked Then
            ASCMAIN1.sql = "Delete from " & TT & " where NVL(WHSE_ON_HAND,0)" _
                            & " - NVL(WHSE_PICK,0)" _
                            & " + NVL(WHSE_ON_ORDER,0)" _
                            & " + NVL(WHSE_TRAN,0)" _
                            & " - NVL(WHSE_OPEN,0) > " & Absx1.numFor("NUMOTSMAX").Value
            ASCDATA1.ExecuteSQL()
        End If


        ' COULD USE TT INSTEAD OF ASTRPT1
        ASCMAIN1.sql = "Select Distinct ASTSRPT1.STYLE_CODE, ASTSRPT1.COLOR_CODE, ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
            & " from " & TT & " ASTSRPT1,ICTSTYC1" & vbCrLf _
            & " where ICTSTYC1.STYLE_CODE (+) = ASTSRPT1.STYLE_CODE" & vbCrLf _
            & "   and  ICTSTYC1.COLOR_CODE (+) = ASTSRPT1.COLOR_CODE"
        ICTCOSTX = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add Primary Key (STYLE_CODE, COLOR_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_LAST_SHP DATE")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_LAST_REC DATE")

        'ASCMAIN1.sql = "" _
        '    & "Begin" & vbCrLf _
        '    & " Declare Cursor C1 is " & vbCrLf _
        '    & "  Select ICTCOSTA.* from ICTCOSTA," & ICTCOSTX & " ICTCOSTX" & vbCrLf _
        '    & "   where ICTCOSTA.STYLE_CODE = ICTCOSTX.STYLE_CODE" & vbCrLf _
        '    & "     and ICTCOSTA.COLOR_CODE = ICTCOSTX.COLOR_CODE" & vbCrLf _
        '    & "     and ICTCOSTA.OPS_YYYYPP = '" & RYP & "';" & vbCrLf _
        '    & " Begin" & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   Update " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
        '    & "    Set STYLE_COST = R1.STYLE_COST, DATE_LAST_SHP = R1.DATE_LAST_SHP, DATE_LAST_REC = R1.DATE_LAST_REC" & vbCrLf _
        '    & "   where ICTCOSTX.STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
        '    & "     and ICTCOSTX.COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
        '    & "  End Loop;" & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End;"
        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select ICTCOSTA.* from ICTCOSTA," & ICTCOSTX & " ICTCOSTX" & vbCrLf _
            & "   where ICTCOSTA.STYLE_CODE = ICTCOSTX.STYLE_CODE" & vbCrLf _
            & "     and ICTCOSTA.COLOR_CODE = ICTCOSTX.COLOR_CODE" & vbCrLf _
            & "     and ICTCOSTA.STYLE_COST > 0;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
            & "    Set STYLE_COST = R1.STYLE_COST, DATE_LAST_SHP = R1.DATE_LAST_SHP, DATE_LAST_REC = R1.DATE_LAST_REC" & vbCrLf _
            & "   where ICTCOSTX.STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
            & "     and ICTCOSTX.COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        ASCMAIN1.Progress("Now Loading Style Activity")

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_COST, ICTSTYL1.CARTON_PACK_QTY " _
            & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        ASCMAIN1.sql = "Select * from " & ICTCOSTX
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTCOSTX", 2))
        For i As Integer = 1 To 4
            dst.Tables("ICTCOSTX").Columns.Add("DATE_" & Format(i, "0"), GetType(System.DateTime))
            dst.Tables("ICTCOSTX").Columns.Add("QTY_" & Format(i, "0"), GetType(System.Int64))
        Next

        ASCMAIN1.sql = "Select * from ICTSTDQ1 where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
        Create_TDA(dst.Tables.Add, "ICTSTDQ1", "**", 0, False, "VV", 0)

        Dim need_to_prepare_ICTCOST1 As Boolean = True
        For Each rowICTCOSTX As DataRow In dst.Tables("ICTCOSTX").Select("")
            Dim STYLE_CODE As String = rowICTCOSTX.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTCOSTX.Item("COLOR_CODE")
            ASCMAIN1.Progress("-", STYLE_CODE & "-" & COLOR_CODE)

            Dim STYLE_COST As Decimal = Val(Split(TAC.ICCMAIN1.Calc_Cost_OH(Me, RYP, STYLE_CODE, COLOR_CODE, need_to_prepare_ICTCOST1), "|")(0))
            need_to_prepare_ICTCOST1 = False

            rowICTCOSTX.Item("STYLE_COST") = STYLE_COST
        Next

    End Sub

    Public Overrides Sub Print_Report()

        'For Each COLUMN_NAME As String In New String() {"DATE_LAST_REC"}
        '    Dim Z As String = Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text & ":"
        '    Dim real_date_selected As Boolean = False
        '    If Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
        '        Z &= " from First"
        '    Else
        '        Z &= " from " & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "MM/dd/yyyy")
        '        real_date_selected = True
        '    End If
        '    If Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
        '        Z &= " to Last"
        '    Else
        '        Z &= " to " & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "MM/dd/yyyy")
        '        real_date_selected = True
        '    End If
        '    If real_date_selected Then
        '        SUBT &= ", " & Z
        '    End If
        '    Page0.Add(Z)
        'Next

        CR_params.Add("SUBT", txtDescription.Text & SUBT)

        'CR_params.Add("NEG", IIf(Absx1.chkFor("CHKNEG").Checked, "1", "0"))

        CR_params.Add("SUB1", IIf(Absx1.chkFor("CHK1").Checked, "1", "0"))
        CR_params.Add("SUB2", IIf(Absx1.chkFor("CHK2").Checked, "1", "0"))
        CR_params.Add("SUB3", IIf(Absx1.chkFor("CHK3").Checked, "1", "0"))
        CR_params.Add("SUB4", IIf(Absx1.chkFor("CHK4").Checked, "1", "0"))

        'CR_params.Add("UD", "U")
        'CR_params.Add("TD", "M")
        'CR_params.Add("AD", "A")

        'CR_params.Add("COST", "1")

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

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

    Private Sub UltraTabControl1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If e.Tab.Key = "Other Run-Time Options" Then
            For S As Integer = 1 To 4
                Dim row() As DataRow = tblASTDSQLA.Select("SEQUENCE = " & CStr(S))
                Absx1.chkFor("CHK" & CStr(S)).Visible = (row.Length = 1)
                If row.Length = 1 Then
                    Absx1.chkFor("CHK" & CStr(S)).Text = row(0).Item("COLUMN_CAPTION")
                End If
            Next
        End If
    End Sub

End Class