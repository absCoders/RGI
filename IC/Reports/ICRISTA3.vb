Imports System.Text

Public Class ICRISTA3
    Dim ICTCOSTX As String
    Dim CUST_CODE_OD As String = ""
    Dim S As New StringBuilder With {.Length = 0}
    Dim RunCount As Integer = 0
    Dim GROUP_CODE As String = ""
    Dim GROUP_CODE_MULT As New List(Of String)
    Dim FormLoading As Boolean = True
    Dim tblICTISTA4 As DataTable = Nothing
    Dim SQLDELRECS As New List(Of String)
    Dim OPT_SUB As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        'RWU = "N"
        RWU = "R"

        Range_Events(grpDATE_LAST_REC)

        'dteLimitOP_C.DateTime = CDate(Now().ToShortDateString)
        dteLimitOP_C.Value = Null

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("X5") Then
                Absx1.chkFor("CHKCOST").Checked = False
                Absx1.chkFor("CHKCOST").Visible = False
                Dim VL As ValueList = Absx1.optFor("OPTAD").ValueList
                VL.ValueListItems.Remove(2)
            End If
        End If

        Set_Labels()
        Setup_Options()

        FormLoading = False
        fillICTISTA4(True)
        grdICTISTA4.DataSource = dst.Tables.Item("ICTISTA4")
        ASCMAIN1.Add_Value_List(grdICTISTA4, "OPTASN", , New String() {":", "A:All", "S:Stock", "N:Non-Stock"})

        With grdICTISTA4.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True
            For i As Integer = 0 To .Bands(0).Columns.Count - 1
                .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
            Next i
            For Each COLNAME As String In New String() {"SELECTED"}
                .Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
        End With

    End Sub

    Private Sub fillICTISTA4(ByVal CreateTable As Boolean)
        If Not FormLoading Then
            Dim sFill As New StringBuilder With {.Length = 0}
            Dim OPS_YYYYPP As String = UltraCombo1.Text
            If OPS_YYYYPP.Length <> 0 Then
                OPS_YYYYPP = OPS_YYYYPP.Substring(0, 4) & OPS_YYYYPP.Substring(5, 2)
            Else
                OPS_YYYYPP = ASCMAIN1.CYP
            End If
            Dim GROUP_CODE As String = ""
            If Not IsNothing(tblASTDSQLA) Then
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Select("SEQUENCE = 1").FirstOrDefault
                If Not IsNothing(rowASTDSQLA) Then
                    GROUP_CODE = rowASTDSQLA.Item("COLUMN_NAME").ToString & String.Empty
                End If
            End If
            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("'0' SELECTED,")
            S.AppendLine("OPS_YYYYPP,")
            S.AppendLine("MAX(OPTASN) AS OPTASN,")
            S.AppendLine("MAX(GROUP_CODE) As GROUP_CODE,")
            S.AppendLine("MAX(GROUP_CAPTION) As GROUP_CAPTION,")
            S.AppendLine("MAX(OPTFILTERS) As OPTFILTERS,")
            S.AppendLine("MAX(LAST_OPER) As LAST_OPER,")
            S.AppendLine("MAX(LAST_DATE) As LAST_DATE")
            S.AppendLine("FROM ICTISTA3")
            sFill.Append(S.ToString)
            If Not dst.Tables.Contains("ICTISTA4") Then
                S.AppendLine("GROUP BY")
                S.AppendLine("OPS_YYYYPP,")
                S.AppendLine("OPTASN")
                ASCMAIN1.sql = S.ToString
                Create_TDA(dst.Tables.Add, "ICTISTA4", "**", 0, False)
            End If
            'sFill.AppendLine(String.Format(" WHERE OPS_YYYYPP <= '{0}'", OPS_YYYYPP))
            If GROUP_CODE.Length > 0 Then
                sFill.AppendLine(String.Format("WHERE GROUP_CODE = '{0}'", GROUP_CODE))
            End If
            sFill.AppendLine("GROUP BY")
            sFill.AppendLine("OPS_YYYYPP,")
            sFill.AppendLine("OPTASN")
            Fill_Records("ICTISTA4",, True, sFill.ToString)
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        RunCount += 1

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = "For Period " & RYPLEGEND
        If Absx1.chkFor("CHKCOST").Checked Then
            SUBT = SUBT & " - Showing Costs"
        End If
        If Absx1.optFor("OPTAD").Value = "D" Then
            SUBT = SUBT & " By Date"
        ElseIf Absx1.optFor("OPTAD").Value = "C" Then
            SUBT = SUBT & " - FIFO Valuation"
        End If

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")

        Dim TDX As Integer = 1
        If Absx1.optFor("OPTTD").Value = "Y" And Mid(RYP, 5, 2) <> "01" Then
            TDX = 2
        End If
        Dim FYP As String = Mid(RYP, 1, 4) & "01"
        If Absx1.optFor("OPTTD").Value = "M" Then
            FYP = RYP
        End If

        For TD As Integer = 1 To TDX

            Dim sql_filter2 As String = ""

            ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf _
                & ", ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE" & vbCrLf _
                & ", DECODE (ICTSTAT1.OPS_YYYYPP,'" & FYP & "',ICTSTAT1.WHSE_QTY_BEG,0) WHSE_QTY_BEG" & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_SHP, ICTSTAT1.WHSE_QTY_RTN " & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_REC, ICTSTAT1.WHSE_QTY_ADJ, ICTSTAT1.WHSE_QTY_XFR " & vbCrLf _
                & ", ICTSTAT1.WHSE_QTY_PHY " & vbCrLf
            If TD = 1 Then
                ASCMAIN1.sql &= "" _
                    & ", ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_ON_ORDER " & vbCrLf _
                    & ", ICTSTAT2.WHSE_QTY_TRAN, ICTSTAT2.WHSE_QTY_OPEN" & vbCrLf _
                    & ", ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO" & vbCrLf
            Else
                ASCMAIN1.sql &= "" _
                    & ", 0 WHSE_QTY_ON_HAND, 0 WHSE_QTY_ON_ORDER " & vbCrLf _
                    & ", 0 WHSE_QTY_TRAN, 0 WHSE_QTY_OPEN" & vbCrLf _
                    & ", 0 WHSE_QTY_PICK, 0 WHSE_QTY_ALLO" & vbCrLf
            End If

            If RYP = ASCMAIN1.CYP And TD = 1 Then
                ASCMAIN1.sql &= " from ICTSTAT2 ICTSTAT2, ICTSTAT1 "
            Else
                ASCMAIN1.sql &= " from ICTSTAT5 ICTSTAT2, ICTSTAT1"
            End If
            ASCMAIN1.sql &= sql_TABLE_NAMEs & vbCrLf

            If TD = 1 Then
                ASCMAIN1.sql &= " where ICTSTAT1.OPS_YYYYPP (+) = '" & RYP & "'" & vbCrLf
                If RYP <> ASCMAIN1.CYP Then
                    ASCMAIN1.sql &= "" _
                        & "   and ICTSTAT2.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
                        & "   and ICTSTAT2.OPS_YYYYPP = ICTSTAT1.OPS_YYYYPP (+) " & vbCrLf
                End If
            Else
                ASCMAIN1.sql &= "" _
                    & " where ICTSTAT1.OPS_YYYYPP >= '" & FYP & "'" & vbCrLf _
                    & "   and ICTSTAT1.OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(RYP, -1) & "'" & vbCrLf _
                    & "   and ICTSTAT2.OPS_YYYYPP = ICTSTAT1.OPS_YYYYPP (+) " & vbCrLf
            End If

            ASCMAIN1.sql &= sql_WHERE & sql_JOIN & sql_filter & sql_filter2 & vbCrLf

            ASCMAIN1.sql &= "" _
                & "   and ICTSTAT2.STYLE_CODE = ICTSTAT1.STYLE_CODE (+) " & vbCrLf _
                & "   and ICTSTAT2.COLOR_CODE = ICTSTAT1.COLOR_CODE (+) " & vbCrLf _
                & "   and ICTSTAT2.WHSE_CODE = ICTSTAT1.WHSE_CODE (+) " & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE"

            If Absx1.chkFor("CHKNEG").Checked = "1" Then
                If btnWIP.Text = "Neg OH" Then
                    ASCMAIN1.sql &= "   and (ICTSTAT2.WHSE_QTY_ON_HAND - ICTSTAT2.WHSE_QTY_PICK + ICTSTAT2.WHSE_QTY_ON_ORDER + ICTSTAT2.WHSE_QTY_TRAN - ICTSTAT2.WHSE_QTY_OPEN) < 0"
                Else
                    ASCMAIN1.sql &= "   and ICTSTAT2.WHSE_QTY_ON_HAND < 0"
                End If
            End If
            If Absx1.optFor("OPTASN").Value = "S" Then
                ASCMAIN1.sql &= "   and ICTSTYL1.CUST_CODE is Null"
            ElseIf Absx1.optFor("OPTASN").Value = "N" Then
                ASCMAIN1.sql &= "   and ICTSTYL1.CUST_CODE is Not Null"
            End If

            'ASCMAIN1.sql &= " group by " & sql_GROUP_BY_cols _
            '    & ", ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE"

            ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                                & " (" & G1thru9 _
                                & ",STYLE_CODE,COLOR_CODE" _
                                & ",WHSE_BEG,WHSE_SHP,WHSE_RTN,WHSE_REC,WHSE_ADJ,WHSE_XFR,WHSE_PHY" _
                                & ",WHSE_ON_HAND,WHSE_ON_ORDER,WHSE_TRAN,WHSE_OPEN,WHSE_PICK,WHSE_ALLO" _
                                & ") " _
                                & " (" & ASCMAIN1.sql & ")")
        Next TD

        ' Replacing next line SO THAT i HAVE FIFO COSTS FOR ALL VERSIONS OF THIS REPORT
        ' If Absx1.optFor("OPTAD").Value = "C" Then
        If Absx1.optFor("OPTAD").Value = "C" Or Absx1.chkFor("CHKCOST").Checked Then

            ' ANNA WANTS TRUE FIFO CALCULATIONS

            'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            '    TAC.ICCMAIN1.Calculate_FIFO_Cost_OH(Me, RYP)

            '    With dst.Tables("ICTCOSTA").Columns
            '        .Add("AGE_1", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)<=" & numDAYS1.Value & ",LOT_AMT_ONHD,0)")
            '        .Add("AGE_2", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)>" & numDAYS1.Value & " AND ISNULL(LOT_DAYS,0)<=" & numDAYS2.Value & ",LOT_AMT_ONHD,0)")
            '        .Add("AGE_3", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)>" & numDAYS2.Value & ",LOT_AMT_ONHD,0)")
            '    End With

            'Else
            Dim ICTCOSTA As String = ""
            TAC.ICCMAIN1.Calculate_FIFO(Me, RYP, False, ICTCOSTA)

            Create_Relation("ICTCOSTA", "ICTCOSTL", "STYLE_CODE,COLOR_CODE")

            With dst.Tables("ICTCOSTL").Columns
                .Add("AGE_1", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)<=" & numDAYS1.Value & ",LOT_AMT_ONHD,0)")
                .Add("AGE_2", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)>" & numDAYS1.Value & " AND ISNULL(LOT_DAYS,0)<=" & numDAYS2.Value & ",LOT_AMT_ONHD,0)")
                .Add("AGE_3", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)>" & numDAYS2.Value & ",LOT_AMT_ONHD,0)")
            End With
            With dst.Tables("ICTCOSTA").Columns
                .Add("AGE_1", GetType(System.Decimal), "SUM(CHILD(ICTCOSTA_ICTCOSTL).AGE_1)")
                .Add("AGE_2", GetType(System.Decimal), "SUM(CHILD(ICTCOSTA_ICTCOSTL).AGE_2)")
                .Add("AGE_3", GetType(System.Decimal), "SUM(CHILD(ICTCOSTA_ICTCOSTL).AGE_3)")
            End With
            'End If
        End If

        fillICTISTA4(True)

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)

        If Absx1.chkFor("CHKZERO").Checked Then
            If Absx1.optFor("OPTAD").Value = "C" Then
                ASCMAIN1.sql = "Delete from " & TT & " where NVL(WHSE_ON_HAND,0) = 0"
                ASCDATA1.ExecuteSQL()
            Else
                ASCMAIN1.sql = ""
                For Each COLUMN_NAME As String In New String() _
                    {"WHSE_BEG", "WHSE_SHP", "WHSE_RTN", "WHSE_REC", "WHSE_ADJ", "WHSE_XFR", "WHSE_PHY",
                     "WHSE_ON_HAND", "WHSE_ON_ORDER", "WHSE_TRAN", "WHSE_OPEN", "WHSE_PICK", "WHSE_ALLO"}
                    ASCMAIN1.sql &= " and NVL(" & COLUMN_NAME & ",0) = 0"
                Next
                ASCMAIN1.sql = "Delete from " & TT & ASCMAIN1.SQL_Add_WHERE(ASCMAIN1.sql)
                ASCDATA1.ExecuteSQL()
            End If
        End If

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

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select ICTCOSTA.* from ICTCOSTA," & ICTCOSTX & " ICTCOSTX" & vbCrLf _
            & "   where ICTCOSTA.STYLE_CODE = ICTCOSTX.STYLE_CODE" & vbCrLf _
            & "     and ICTCOSTA.COLOR_CODE = ICTCOSTX.COLOR_CODE" & vbCrLf _
            & "     and ICTCOSTA.OPS_YYYYPP = '" & RYP & "';" & vbCrLf _
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

        Dim SQLD As String = Get_Dates()
        If SQLD <> "" Then
            Dim sqlC As String = "Select Distinct STYLE_CODE, COLOR_CODE from " & TT & vbCrLf _
                & " minus " & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE from " & ICTCOSTX & " ICTCOSTX" & ASCMAIN1.SQL_Add_WHERE(SQLD)
            ASCMAIN1.sql = "Delete from " & TT & " where (STYLE_CODE,COLOR_CODE) in (" & vbCrLf & sqlC & vbCrLf & ")"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Delete from " & ICTCOSTX & " where (STYLE_CODE,COLOR_CODE) in (" & vbCrLf & sqlC & vbCrLf & ")"
            ASCDATA1.ExecuteSQL()
        End If

        If optFORMAT.Value = "C" Then

        End If

    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        ASCMAIN1.Progress("Now Loading Style Activity")

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        'If RunCount = 1 Then
        '    S.AppendLine("SELECT OPS_YYYYPP,")
        '    S.AppendLine(String.Format("'{0}' AS G1", "".PadRight(50, " ")))
        '    S.AppendLine("FROM GLTINTF1")
        '    S.AppendLine("WHERE ROWNUM < 0")
        '    ASCMAIN1.sql = S.ToString
        '    dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTISTA3", 2))
        '    With dst.Tables("ICTISTA3").Columns
        '        .Add("VAL_FULL", GetType(System.Double))
        '        .Add("OP_01", GetType(System.Double))
        '        .Add("OP_02", GetType(System.Double))
        '        .Add("OP_03", GetType(System.Double))
        '        .Add("AGE_01", GetType(System.Double))
        '        .Add("AGE_02", GetType(System.Double))
        '        .Add("AGE_03", GetType(System.Double))
        '    End With
        'End If

        'S.Length = 0
        'S.AppendLine("Select *")
        'S.AppendLine("FROM ICTISTA3")
        'S.AppendLine(String.Format("WHERE OPS_YYYYPP <= '{0}'", RYP))
        'S.AppendLine(String.Format("AND GROUP_CODE = '{0}'", GROUP_CODE))
        'S.AppendLine(String.Format("AND OPTASN = '{0}'", optASN.Value))
        'ASCMAIN1.sql = S.ToString
        'Create_TDA(dst.Tables.Add, "ICTISTA3", "**", 0, True)
        'Fill_Records("ICTISTA3")

        S.Length = 0
        S.AppendLine("SELECT SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, SUM(NVL(SOTRSRV2.RSRV_QTY_OPEN,0)) AS RSRV_QTY_OPEN")
        S.AppendLine("FROM SOTRSRV1, SOTRSRV2")
        S.AppendLine("WHERE SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO")
        S.AppendLine("AND ORDR_CANCEL_DATE < :PARM1")
        S.AppendLine("HAVING SUM(NVL(SOTRSRV2.RSRV_QTY_OPEN,0)) > 0")
        S.AppendLine("GROUP BY SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE")
        ASCMAIN1.sql = S.ToString
        Create_TDA(dst.Tables.Add, "SOTRSRV2", "**", 0, False, "D")

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("O2.STYLE_CODE, COLOR_CODE, SUM(NVL(ORDR_QTY_OPEN,0)) AS ORDR_QTY_OPEN")
        S.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
        S.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
        S.AppendLine("AND O1.ORDR_STATUS = 'O'")
        S.AppendLine("AND O1.ORDR_CANCEL_DATE < :PARM1")
        S.AppendLine("GROUP BY O2.STYLE_CODE, COLOR_CODE")
        S.AppendLine("HAVING SUM(NVL(ORDR_QTY_OPEN,0)) > 0")
        ASCMAIN1.sql = S.ToString
        Create_TDA(dst.Tables.Add("SOTCANCL"), "SOTORDR2", "**", 0, False, "D")

        S.Length = 0
        S.AppendLine("Select *")
        S.AppendLine("FROM ICTISTA3")
        ASCMAIN1.sql = S.ToString
        Create_TDA(dst.Tables.Add, "ICTISTA3", "**", 0, True)
        dst.Tables("ICTISTA3").Columns.Add("GROUP_CODE2")
        dst.Tables("ICTISTA3").Columns.Add("GROUP_DESC2")
        dst.Tables("ICTISTA3").Columns.Add("GROUP_CAPTION2")
        dst.Tables("ICTISTA3").Columns.Add("QTY_01", GetType(System.Int64))
        dst.Tables("ICTISTA3").Columns.Add("QTY_02", GetType(System.Int64))
        dst.Tables("ICTISTA3").Columns.Add("QTY_03", GetType(System.Int64))
        dst.Tables("ICTISTA3").Columns.Add("QTYOP_01", GetType(System.Int64))
        dst.Tables("ICTISTA3").Columns.Add("QTYOP_02", GetType(System.Int64))
        dst.Tables("ICTISTA3").Columns.Add("QTYOP_03", GetType(System.Int64))



        SQLDELRECS.Clear()
        For Each rowICTISTA4 As DataRow In tblICTISTA4.Select("SELECTED = '1'")
            S.Length = 0
            S.AppendLine("Select *")
            S.AppendLine("FROM ICTISTA3")
            S.AppendLine(String.Format("WHERE OPS_YYYYPP = '{0}'", rowICTISTA4.Item("OPS_YYYYPP").ToString & String.Empty))
            S.AppendLine(String.Format("AND GROUP_CODE = '{0}'", rowICTISTA4.Item("GROUP_CODE").ToString & String.Empty))
            S.AppendLine(String.Format("AND OPTASN = '{0}'", rowICTISTA4.Item("OPTASN").ToString & String.Empty))
            Fill_Records("ICTISTA3",, False, S.ToString)
            S.Length = 0
            S.AppendLine("DELETE")
            S.AppendLine("FROM ICTISTA3")
            S.AppendLine(String.Format("WHERE OPS_YYYYPP = '{0}'", rowICTISTA4.Item("OPS_YYYYPP").ToString & String.Empty))
            S.AppendLine(String.Format("AND GROUP_CODE = '{0}'", rowICTISTA4.Item("GROUP_CODE").ToString & String.Empty))
            S.AppendLine(String.Format("AND OPTASN = '{0}'", rowICTISTA4.Item("OPTASN").ToString & String.Empty))
            SQLDELRECS.Add(S.ToString)
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_COST, ICTSTYL1.CARTON_PACK_QTY " _
            & " from ICTSTYL1 where STYLE_CODE In (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        'ASCMAIN1.sql = "Select Distinct ASTSRPT1.STYLE_CODE, ASTSRPT1.COLOR_CODE, ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
        '    & " from " & ASTSRPT1 & " ASTSRPT1,ICTSTYC1" & vbCrLf _
        '    & " where ICTSTYC1.STYLE_CODE (+) = ASTSRPT1.STYLE_CODE" & vbCrLf _
        '    & "   And  ICTSTYC1.COLOR_CODE (+) = ASTSRPT1.COLOR_CODE"
        'ICTCOSTX = ASCMAIN1.Temp_Table
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add Primary Key (STYLE_CODE, COLOR_CODE)")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_LAST_SHP Date")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_LAST_REC Date")

        'ASCMAIN1.sql = "" _
        '    & "Begin" & vbCrLf _
        '    & " Declare Cursor C1 Is " & vbCrLf _
        '    & "  Select ICTCOSTA.* from ICTCOSTA," & ICTCOSTX & " ICTCOSTX" & vbCrLf _
        '    & "   where ICTCOSTA.STYLE_CODE = ICTCOSTX.STYLE_CODE" & vbCrLf _
        '    & "     And ICTCOSTA.COLOR_CODE = ICTCOSTX.COLOR_CODE" & vbCrLf _
        '    & "     And ICTCOSTA.OPS_YYYYPP = '" & RYP & "';" & vbCrLf _
        '    & " Begin" & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   Update " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
        '    & "    Set STYLE_COST = R1.STYLE_COST, DATE_LAST_SHP = R1.DATE_LAST_SHP, DATE_LAST_REC = R1.DATE_LAST_REC" & vbCrLf _
        '    & "   where ICTCOSTX.STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
        '    & "     and ICTCOSTX.COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
        '    & "  End Loop;" & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End;"
        'ASCDATA1.ExecuteSQL()


        'Dim sqlLAST_SHIPPED As String = "" _
        '    & "Select MAX(INV_DATE) AS INV_DATE_LAST_SHIPPED" & vbCrLf _
        '    & " from SOTINVH1 S1, SOTINVH2 S2" & vbCrLf _
        '    & " WHERE S1.INV_TYPE = S2.INV_TYPE" & vbCrLf _
        '    & " AND S1.INV_NO = S2.INV_NO" & vbCrLf _
        '    & " AND S2.STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
        '    & " AND COLOR_CODE = R1.COLOR_CODE"

        'Dim sqlLAST_RECD As String = "" _
        '    & "Select MAX(PO_DATE_RECEIVED) AS LAST_RECD" & vbCrLf _
        '    & " from POTSHIP2 S2, POTSHIP3 S3, POTORDR2 O2" & vbCrLf _
        '    & " where S2.PO_SHIPMENT_NO = S3.PO_SHIPMENT_NO" & vbCrLf _
        '    & "   and S2.PO_SHIPMENT_LNO = S3.PO_SHIPMENT_LNO" & vbCrLf _
        '    & "   and S3.PO_ORDER_NO = O2.PO_ORDER_NO" & vbCrLf _
        '    & "   and S3.PO_ORDER_LNO = O2.PO_ORDER_LNO" & vbCrLf _
        '    & "   and O2.STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
        '    & "   and O2.COLOR_CODE = R1.COLOR_CODE"

        'ASCMAIN1.sql = "" _
        '    & "Begin" & vbCrLf _
        '    & " Declare Cursor C1 is Select * from " & ICTCOSTX & " for Update; " & vbCrLf _
        '    & " Begin" & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   Update " & ICTCOSTX & " Set LAST_SHIPPED = (" & sqlLAST_SHIPPED & ")" & vbCrLf _
        '    & "    where Current of C1;" & vbCrLf _
        '    & "   Update " & ICTCOSTX & " Set LAST_RECD = (" & sqlLAST_RECD & ")" & vbCrLf _
        '    & "    where Current of C1;" & vbCrLf _
        '    & "  End Loop;" & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End;"
        'ASCDATA1.ExecuteSQL()


        'ASCMAIN1.Progress("Now calculating FIFO Costs")

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

            If Absx1.optFor("OPTAD").Value = "D" Then
                Fill_Records("ICTSTDQ1", New String() {STYLE_CODE, COLOR_CODE})
                Dim R As Integer = 0
                For Each rowICTSTDQ1 As DataRow In dst.Tables("ICTSTDQ1").Select("", "STATUS_DATE")
                    R += 1
                    rowICTCOSTX.Item("DATE_" & CStr(R)) = rowICTSTDQ1.Item("STATUS_DATE")
                    rowICTCOSTX.Item("QTY_" & CStr(R)) = rowICTSTDQ1.Item("STATUS_QTY")
                    If R = 4 Then Exit For
                Next
            End If

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                If Absx1.optFor("OPTAD").Value = "C" Or Absx1.chkFor("CHKCOST").Checked Then
                    'If STYLE_CODE = "C1007" Then Stop

                    ' SOMEHOW, THIS REPORT IS ZEROING OUT VALUE OF ITEMS WITH NEG ON HAND

                    Dim rowICTCOSTA As DataRow = dst.Tables("ICTCOSTA").Rows.Find(New String() {RYP, STYLE_CODE, COLOR_CODE})
                    If rowICTCOSTA Is Nothing Then
                        rowICTCOSTX.Item("STYLE_COST") = 0
                    Else
                        rowICTCOSTX.Item("STYLE_COST") = rowICTCOSTA.Item("STYLE_COST")
                    End If

                    'Dim STYLE_COST As Decimal = Val(Split(TAC.ICCMAIN1.Calc_Cost_OH(Me, RYP, STYLE_CODE, COLOR_CODE, need_to_prepare_ICTCOST1), "|")(0))
                    'need_to_prepare_ICTCOST1 = False

                    'rowICTCOSTX.Item("STYLE_COST") = STYLE_COST
                    'If Absx1.optFor("OPTUD").Value = "D" Then
                    '    rowICTCOSTX.Item("STYLE_COST") = Val(rowICTCOSTX.Item("STYLE_COST") & "") / 12 ' NOT CORRECT - NEED TO MULTIPLY BY SUB_UNIT_PACK
                    'End If

                    ''For Each rowICTCOSTA As DataRow In dst.Tables("ICTCOSTA").Select("")
                    ''    Dim DAYS As Integer = Val(rowICTCOSTA.Item("DAYS") & "")
                    ''    If DAYS > 360 Then
                    ''        rowICTCOSTA.Item("AGE_3") = rowICTCOSTA.Item("LOT_AMT_ONHD")
                    ''    ElseIf DAYS > 180 Then
                    ''        rowICTCOSTA.Item("AGE_2") = rowICTCOSTA.Item("LOT_AMT_ONHD")
                    ''    Else
                    ''        rowICTCOSTA.Item("AGE_1") = rowICTCOSTA.Item("LOT_AMT_ONHD")
                    ''    End If
                    ''Next
                End If
            End If
        Next

        'If Absx1.optFor("OPTAD").Value = "C" Then

        '    If Absx1.optFor("OPTONLY").Value & "" <> "1" Then
        '        Dim sqlw As String = ""
        '        If Absx1.optFor("OPTONLY").Value & "" = "2" Then
        '            sqlw = "AGE_2 = 0 and AGE_3 = 0"
        '        Else
        '            sqlw = "AGE_3 = 0"
        '        End If

        '        For Each row As DataRow In dst.Tables("ICTCOSTA").Select(sqlw)
        '            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
        '            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
        '            ASCDATA1.DeleteRows("ASTSRPT1", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
        '        Next
        '        ASCDATA1.DeleteRows("ICTCOSTA", sqlw)

        '    End If
        'End If




        'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        '    If ASCMAIN1.Running_in_VS Then
        '        ASCDATA1.ExecuteSQL("Truncate table WZTISTA1")

        '        Create_TDA(dst.Tables.Add, "WZTISTA1", "*")

        '        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("")
        '            Dim rowWZ As DataRow = dst.Tables("WZTISTA1").NewRow
        '            rowWZ.Item("OPS_YYYYPP") = RYP
        '            rowWZ.Item("STYLE_CODE") = row.Item("STYLE_CODE")
        '            rowWZ.Item("COLOR_CODE") = row.Item("COLOR_CODE")
        '            Dim rowICTCOSTX As DataRow = dst.Tables("ICTCOSTX").Rows.Find(New String() {row.Item("STYLE_CODE"), row.Item("COLOR_CODE")})
        '            Dim COST As Decimal = Val(rowICTCOSTX.Item("STYLE_COST") & "")
        '            Dim OH As Int64 = Val(row.Item("WHSE_ON_HAND") & "")
        '            If COST <> 0 Or OH <> 0 Then
        '                rowWZ.Item("COST") = COST
        '                rowWZ.Item("OH") = OH
        '                dst.Tables("WZTISTA1").Rows.Add(rowWZ)
        '            End If
        '        Next
        '        Update_Record_TDA("WZTISTA1")


        '        'ASCMAIN1.sql = "INSERT INTO WZTISTA1 SELECT '" & RYP & "' OPS_YYYYPP" & vbCrLf _
        '        '    & ", ASTSRPT1.STYLE_CODE, ASTSRPT1.COLOR_CODE" & vbCrLf _
        '        '    & ", ICTCOSTX.STYLE_COST COST, ASTSRPT1.WHSE_ON_HAND QTY" & vbCrLf _
        '        '    & "FROM " & ASTSRPT1 & " ASTSRPT1, " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
        '        '    & " where ICTCOSTX.STYLE_CODE = ASTSRPT1.STYLE_CODE" & vbCrLf _
        '        '    & "   and ICTCOSTX.COLOR_CODE = ASTSRPT1.COLOR_CODE" & vbCrLf _
        '        '    & "   and (NVL(ICTCOSTX.STYLE_COST,0) <> 0 or NVL(ASTSRPT1.WHSE_ON_HAND,0) <> 0)"
        '        'ASCDATA1.ExecuteSQL()
        '    End If
        'End If
    End Sub

    Public Overrides Sub Print_Report()
        'RPT = "ICRISTAV"

        dst.Tables("ASTSRPT1").Columns.Add("SHOWFL1", GetType(System.String))
        dst.Tables("ASTSRPT1").Columns.Add("SHOWFL2", GetType(System.String))
        dst.Tables("ASTSRPT1").Columns.Add("OP_COST", GetType(System.Double))
        dst.Tables("ASTSRPT1").Columns.Add("AGE_01", GetType(System.Double))
        dst.Tables("ASTSRPT1").Columns.Add("AGE_02", GetType(System.Double))
        dst.Tables("ASTSRPT1").Columns.Add("AGE_03", GetType(System.Double))
        dst.Tables("ASTSRPT1").Columns.Add("EXT_COST", GetType(System.Double))
        dst.Tables("ASTSRPT1").Columns.Add("QTY_01", GetType(System.Int64))
        dst.Tables("ASTSRPT1").Columns.Add("QTY_02", GetType(System.Int64))
        dst.Tables("ASTSRPT1").Columns.Add("QTY_03", GetType(System.Int64))
        dst.Tables("ASTSRPT1").Columns.Add("QTYOP_01", GetType(System.Int64))
        dst.Tables("ASTSRPT1").Columns.Add("QTYOP_02", GetType(System.Int64))
        dst.Tables("ASTSRPT1").Columns.Add("QTYOP_03", GetType(System.Int64))


        CalcExtCost()
        If GROUP_CODE_MULT.Count = 0 Then
            FillASTSRPT1()
        Else
            FillASTSRPT1_2()
        End If


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

        Dim ReportSubtitle As String = SUBT
        Dim Reporttitle As String = ""

        For i As Integer = 1 To 3
            RPT = "ICRISTA3"
            For Each COLUMN_NAME As String In New String() {"DATE_LAST_REC"}
                Dim Z As String = Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text & ":"
                Dim real_date_selected As Boolean = False
                If Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                    Z &= " from First"
                Else
                    Z &= " from " & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "MM/dd/yyyy")
                    real_date_selected = True
                End If
                If Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                    Z &= " to Last"
                Else
                    Z &= " to " & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "MM/dd/yyyy")
                    real_date_selected = True
                End If
                If real_date_selected Then
                    ReportSubtitle &= ", " & Z
                End If
                Page0.Add(Z)
            Next

            CR_params.Add("SUBT", txtDescription.Text & SUBT)

            CR_params.Add("NEG", IIf(Absx1.chkFor("CHKNEG").Checked, "1", "0"))

            CR_params.Add("SUB1", IIf(Absx1.chkFor("CHK1").Checked, "1", "0"))
            CR_params.Add("SUB2", IIf(Absx1.chkFor("CHK2").Checked, "1", "0"))
            CR_params.Add("SUB3", IIf(Absx1.chkFor("CHK3").Checked, "1", "0"))
            CR_params.Add("SUB4", IIf(Absx1.chkFor("CHK4").Checked, "1", "0"))

            CR_params.Add("UD", Absx1.optFor("OPTUD").Value & "")
            CR_params.Add("TD", Absx1.optFor("OPTTD").Value & "")
            CR_params.Add("AD", Absx1.optFor("OPTAD").Value & "")

            Dim rsf As String = ""

            CR_params.Add("COST", "1")

            CR_params.Add("AGE_1", Replace(lblNUMDAYS1.Text, " Days", ""))
            CR_params.Add("AGE_2", Replace(lblNUMDAYS2.Text, " Days", ""))
            CR_params.Add("AGE_3", Replace(lblNUMDAYS3.Text, " Days", ""))

            CR_params.Add("DAYS_1", numDAYS1.Value)
            CR_params.Add("DAYS_2", numDAYS2.Value)

            CR_params.Add("OPEN_PICK", IIf(Absx1.chkFor("CHKOPEN_PICK").Checked, "1", "0"))
            CR_params.Add("OPEN_PICK_UNITS", IIf(Absx1.chkFor("CHKOPEN_PICK_UNITS").Checked, "1", "0"))

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                CR_params.Add("USE_LOTS_FOR_AGE", "0")
            Else
                CR_params.Add("USE_LOTS_FOR_AGE", "1")
            End If

            'If Absx1.optFor("OPTONLY").Value & "" <> "1" Then
            '    If Absx1.optFor("OPTONLY").Value & "" = "2" Then
            '        rsf = "{@AGE_2} <> 0 or {@AGE_3} <> 0"
            '    Else
            '        rsf = "{@AGE_3} <> 0"
            '    End If
            'End If

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                Select Case i
                    Case 1
                        ReportSubtitle = SUBT & " Up to 180"
                        Reporttitle = "Original FIFO Report Up to 180"
                        rsf = "{@AGE_1} <> 0"
                    Case 2
                        rsf = "{@AGE_2} <> 0"
                        Reporttitle = "Original FIFO Report 181-360"
                    Case 3
                        rsf = "{@AGE_3} <> 0"
                        Reporttitle = "Original FIFO Report Over 360"
                End Select

            Else
                Select Case i
                    Case 1
                        ReportSubtitle = SUBT & " Up to 180"
                        Reporttitle = "Original FIFO Report Up to 180"
                    Case 2
                        rsf = "{@AGE_2} <> 0 or {@AGE_3} <> 0"
                        Reporttitle = "Original FIFO Report 181-360"
                    Case 3
                        rsf = "{@AGE_3} <> 0"
                        Reporttitle = "Original FIFO Report Over 360"
                End Select
            End If

            If chkHIDEFIFOCOST.Checked Then
                CR_params.Add("HIDEFIFOCOST", "1")
            Else
                CR_params.Add("HIDEFIFOCOST", "0")
            End If
            If Absx1.chkFor("CHKTOTALSONLY").Checked Then
                CR_params.Add("TOTALSONLY", "1")
            Else
                CR_params.Add("TOTALSONLY", "0")
            End If
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                If ASCMAIN1.CYP = RYP Then
                    CR_params.Add("BEGIN_AGE_DATE", DateSerial(Now().Year, Now().Month, Now().Day))
                Else
                    CR_params.Add("BEGIN_AGE_DATE", DateSerial(RYP.Substring(0, 4), RYP.Substring(4, 2), 1).AddMonths(1).AddDays(-1))
                End If
            Else
                CR_params.Add("BEGIN_AGE_DATE", DateSerial(Now().Year, Now().Month, Now().Day))
            End If
            Generate_Report(RPT, Reporttitle, ReportSubtitle, rsf)
        Next

        OPT_SUB = ""
        If chkLimitOP_O.Checked = True Then
            OPT_SUB = OPT_SUB & " Open"
        End If

        If chkLimitOP_P.Checked = True Then
            OPT_SUB = OPT_SUB & " Pick"
        End If

        If chkLimitOP_R.Checked = True Then
            OPT_SUB = OPT_SUB & " Res"
        End If

        If chkLimitOP_C.Checked Then
            OPT_SUB = OPT_SUB & " Cancel < " & CDate(dteLimitOP_C.DateTime.ToShortDateString)
        End If
        If GROUP_CODE_MULT.Count = 0 Then
            If chkUNITS.Checked Then
                CR_params.Add("UNITS", "1")
            Else
                CR_params.Add("UNITS", "0")
            End If
            If CHKUNITSONLY.Checked Then
                OPT_SUB = OPT_SUB & " - UNITS ONLY"
                Generate_Report("ICRISTA6", "", OPT_SUB)
            Else
                Generate_Report("ICRISTA4", "", OPT_SUB)
            End If

            RWU = "R"
        Else
            Generate_Report("ICRISTA5", "", OPT_SUB)
            RWU = "N"
        End If

    End Sub

    Overrides Sub Update_Record()
        SQLDELRECS.Clear() 'Someday we will stop loading this up top.
        Dim YYYYPP As String = UltraCombo1.Text
        If YYYYPP.Length <> 0 Then
            YYYYPP = YYYYPP.Substring(0, 4) & YYYYPP.Substring(5, 2)
        Else
            YYYYPP = ASCMAIN1.CYP
        End If
        Dim SQ As New StringBuilder With {.Length = 0}
        SQ.AppendLine("DELETE")
        SQ.AppendLine("FROM ICTISTA3")
        SQ.AppendLine(String.Format("WHERE OPS_YYYYPP = '{0}'", YYYYPP))
        SQ.AppendLine(String.Format("AND GROUP_CODE = '{0}'", GROUP_CODE))
        SQ.AppendLine(String.Format("AND OPTASN = '{0}'", optASN.Value))
        SQLDELRECS.Add(SQ.ToString)
        For Each SQLDEL As String In SQLDELRECS
            ASCMAIN1.sql = SQLDEL
            ASCDATA1.ExecuteSQL()
        Next
        Update_Record_TDA("ICTISTA3")
    End Sub

    Private Sub FillASTSRPT1()
        Dim GROUP_DESCS As New List(Of String)
        Dim OPTFILTERS As String = setOPTFILTERS
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select()
            If Not GROUP_DESCS.Contains(rowASTSRPT1.Item("G1").ToString & String.Empty) Then
                GROUP_DESCS.Add(rowASTSRPT1.Item("G1").ToString & String.Empty)
            End If
        Next

        For Each GROUP_DESC As String In GROUP_DESCS
            'Dim OPS_YYYYPP As String = ASCMAIN1.CYP
            Dim OPS_YYYYPP As String = RYP
            Dim VAL_FULL As Double = 0
            Dim OP_01 As Double = 0
            Dim OP_02 As Double = 0
            Dim OP_03 As Double = 0
            Dim AGE_01 As Double = 0
            Dim AGE_02 As Double = 0
            Dim AGE_03 As Double = 0
            Dim QTY_01 As Double = 0
            Dim QTY_02 As Double = 0
            Dim QTY_03 As Double = 0
            Dim QTYOP_01 As Double = 0
            Dim QTYOP_02 As Double = 0
            Dim QTYOP_03 As Double = 0

            Dim filter As String = String.Format("G1 = '{0}'", GROUP_DESC)
            For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select(filter)
                VAL_FULL = VAL_FULL + Val(rowASTSRPT1.Item("EXT_COST").ToString & String.Empty)
                OP_01 = OP_01 + Val(rowASTSRPT1.Item("OP_COST").ToString & String.Empty)
                AGE_01 = AGE_01 + Val(rowASTSRPT1.Item("AGE_01").ToString & String.Empty)
                AGE_02 = AGE_02 + Val(rowASTSRPT1.Item("AGE_02").ToString & String.Empty)
                AGE_03 = AGE_03 + Val(rowASTSRPT1.Item("AGE_03").ToString & String.Empty)
                QTY_01 = QTY_01 + Val(rowASTSRPT1.Item("QTY_01").ToString & String.Empty)
                QTY_02 = QTY_02 + Val(rowASTSRPT1.Item("QTY_02").ToString & String.Empty)
                QTY_03 = QTY_03 + Val(rowASTSRPT1.Item("QTY_03").ToString & String.Empty)
                QTYOP_01 = QTYOP_01 + Val(rowASTSRPT1.Item("QTYOP_01").ToString & String.Empty)
                QTYOP_02 = QTYOP_02 + Val(rowASTSRPT1.Item("QTYOP_02").ToString & String.Empty)
                QTYOP_03 = QTYOP_03 + Val(rowASTSRPT1.Item("QTYOP_03").ToString & String.Empty)


                If Val(rowASTSRPT1.Item("AGE_02").ToString & String.Empty) <> 0 Or Val(rowASTSRPT1.Item("AGE_03").ToString & String.Empty) <> 0 Then
                    OP_02 = OP_02 + Val(rowASTSRPT1.Item("OP_COST").ToString & String.Empty)
                End If
                If Val(rowASTSRPT1.Item("AGE_03").ToString & String.Empty) <> 0 Then
                    OP_03 = OP_03 + Val(rowASTSRPT1.Item("OP_COST").ToString & String.Empty)
                End If
            Next

            Dim Fltr As String = String.Format("OPS_YYYYPP = '{0}' AND GROUP_CODE = '{1}' AND GROUP_DESC = '{2}' AND OPTASN = '{3}'", OPS_YYYYPP, GROUP_CODE, GROUP_DESC, optASN.Value)
            If dst.Tables.Item("ICTISTA3").Select(Fltr).Count = 1 Then
                Dim rowICTISTA3 As DataRow = dst.Tables("ICTISTA3").Select(Fltr).FirstOrDefault
                rowICTISTA3.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowICTISTA3.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                rowICTISTA3.Item("VAL_FULL") = VAL_FULL
                rowICTISTA3.Item("OP_01") = OP_01
                rowICTISTA3.Item("OP_02") = OP_02
                rowICTISTA3.Item("OP_03") = OP_03
                rowICTISTA3.Item("AGE_01") = AGE_01
                rowICTISTA3.Item("AGE_02") = AGE_02
                rowICTISTA3.Item("AGE_03") = AGE_03
                rowICTISTA3.Item("QTY_01") = QTY_01
                rowICTISTA3.Item("QTY_02") = QTY_02
                rowICTISTA3.Item("QTY_03") = QTY_03
                rowICTISTA3.Item("QTYOP_01") = QTYOP_01
                rowICTISTA3.Item("QTYOP_02") = QTYOP_02
                rowICTISTA3.Item("QTYOP_03") = QTYOP_03

            Else
                Dim GROUP_CAPTION As String = ""
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Select("SEQUENCE = 1").FirstOrDefault
                If Not IsNothing(rowASTDSQLA) Then
                    GROUP_CAPTION = rowASTDSQLA.Item("COLUMN_CAPTION").ToString & String.Empty
                End If
                Dim newICTISTA3 As DataRow = dst.Tables("ICTISTA3").NewRow
                newICTISTA3.Item("OPS_YYYYPP") = OPS_YYYYPP
                newICTISTA3.Item("GROUP_CODE") = GROUP_CODE
                newICTISTA3.Item("GROUP_DESC") = GROUP_DESC
                newICTISTA3.Item("GROUP_CAPTION") = GROUP_CAPTION
                newICTISTA3.Item("OPTFILTERS") = OPTFILTERS
                newICTISTA3.Item("OPTASN") = optASN.Value
                newICTISTA3.Item("LAST_OPER") = ASCMAIN1.USER_ID
                newICTISTA3.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                newICTISTA3.Item("VAL_FULL") = VAL_FULL
                newICTISTA3.Item("OP_01") = OP_01
                newICTISTA3.Item("OP_02") = OP_02
                newICTISTA3.Item("OP_03") = OP_03
                newICTISTA3.Item("AGE_01") = AGE_01
                newICTISTA3.Item("AGE_02") = AGE_02
                newICTISTA3.Item("AGE_03") = AGE_03
                newICTISTA3.Item("QTY_01") = QTY_01
                newICTISTA3.Item("QTY_02") = QTY_02
                newICTISTA3.Item("QTY_03") = QTY_03
                newICTISTA3.Item("QTYOP_01") = QTYOP_01
                newICTISTA3.Item("QTYOP_02") = QTYOP_02
                newICTISTA3.Item("QTYOP_03") = QTYOP_03

                dst.Tables("ICTISTA3").Rows.Add(newICTISTA3)
            End If

        Next
    End Sub

    Private Sub FillASTSRPT1_2()
        Dim GROUP_DESCS As New List(Of KeyValuePair(Of String, String))
        Dim OPTFILTERS As String = setOPTFILTERS()
        Dim PK As DataColumn() = dst.Tables("ICTISTA3").PrimaryKey
        ReDim Preserve PK(5)
        PK(4) = dst.Tables("ICTISTA3").Columns.Item("GROUP_CODE2")
        PK(5) = dst.Tables("ICTISTA3").Columns.Item("GROUP_DESC2")
        dst.Tables("ICTISTA3").PrimaryKey = PK
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select()
            Dim KVP As New KeyValuePair(Of String, String)(rowASTSRPT1.Item("G1").ToString & String.Empty, rowASTSRPT1.Item("G2").ToString & String.Empty)
            If Not GROUP_DESCS.Contains(KVP) Then
                GROUP_DESCS.Add(KVP)
            End If
        Next

        'For Each GC As String In GROUP_CODE_MULT
        For Each GROUP_DESC As KeyValuePair(Of String, String) In GROUP_DESCS
            'If GROUP_DESC.Value = GC Then
            Dim OPS_YYYYPP As String = RYP
            Dim VAL_FULL As Double = 0
            Dim OP_01 As Double = 0
            Dim OP_02 As Double = 0
            Dim OP_03 As Double = 0
            Dim AGE_01 As Double = 0
            Dim AGE_02 As Double = 0
            Dim AGE_03 As Double = 0
            Dim QTY_01 As Double = 0
            Dim QTY_02 As Double = 0
            Dim QTY_03 As Double = 0
            Dim QTYOP_01 As Double = 0
            Dim QTYOP_02 As Double = 0
            Dim QTYOP_03 As Double = 0


            Dim filter As String = $"G1 = '{GROUP_DESC.Key}' AND G2 = '{GROUP_DESC.Value}'"
            For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select(filter)
                VAL_FULL = VAL_FULL + Val(rowASTSRPT1.Item("EXT_COST").ToString & String.Empty)
                OP_01 = OP_01 + Val(rowASTSRPT1.Item("OP_COST").ToString & String.Empty)
                AGE_01 = AGE_01 + Val(rowASTSRPT1.Item("AGE_01").ToString & String.Empty)
                AGE_02 = AGE_02 + Val(rowASTSRPT1.Item("AGE_02").ToString & String.Empty)
                AGE_03 = AGE_03 + Val(rowASTSRPT1.Item("AGE_03").ToString & String.Empty)
                QTY_01 = QTY_01 + Val(rowASTSRPT1.Item("QTY_01").ToString & String.Empty)
                QTY_02 = QTY_02 + Val(rowASTSRPT1.Item("QTY_02").ToString & String.Empty)
                QTY_03 = QTY_03 + Val(rowASTSRPT1.Item("QTY_03").ToString & String.Empty)
                QTYOP_01 = QTYOP_01 + Val(rowASTSRPT1.Item("QTYOP_01").ToString & String.Empty)
                QTYOP_02 = QTYOP_02 + Val(rowASTSRPT1.Item("QTYOP_02").ToString & String.Empty)
                QTYOP_03 = QTYOP_03 + Val(rowASTSRPT1.Item("QTYOP_03").ToString & String.Empty)


                If Val(rowASTSRPT1.Item("AGE_02").ToString & String.Empty) <> 0 Or Val(rowASTSRPT1.Item("AGE_03").ToString & String.Empty) <> 0 Then
                    OP_02 = OP_02 + Val(rowASTSRPT1.Item("OP_COST").ToString & String.Empty)
                End If
                If Val(rowASTSRPT1.Item("AGE_03").ToString & String.Empty) <> 0 Then
                    OP_03 = OP_03 + Val(rowASTSRPT1.Item("OP_COST").ToString & String.Empty)
                End If
            Next

            Dim Fltr As String = $"OPS_YYYYPP = '{OPS_YYYYPP}' AND GROUP_CODE = '{GROUP_DESC.Key}' AND GROUP_CODE2 = '{GROUP_DESC.Value}' AND OPTASN = '{optASN.Value}'"
            If dst.Tables.Item("ICTISTA3").Select(Fltr).Count = 1 Then
                Dim rowICTISTA3 As DataRow = dst.Tables("ICTISTA3").Select(Fltr).FirstOrDefault
                rowICTISTA3.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowICTISTA3.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                rowICTISTA3.Item("VAL_FULL") = VAL_FULL
                rowICTISTA3.Item("OP_01") = OP_01
                rowICTISTA3.Item("OP_02") = OP_02
                rowICTISTA3.Item("OP_03") = OP_03
                rowICTISTA3.Item("AGE_01") = AGE_01
                rowICTISTA3.Item("AGE_02") = AGE_02
                rowICTISTA3.Item("AGE_03") = AGE_03
                rowICTISTA3.Item("QTY_01") = QTY_01
                rowICTISTA3.Item("QTY_02") = QTY_02
                rowICTISTA3.Item("QTY_03") = QTY_03
                rowICTISTA3.Item("QTYOP_01") = QTYOP_01
                rowICTISTA3.Item("QTYOP_02") = QTYOP_02
                rowICTISTA3.Item("QTYOP_03") = QTYOP_03
            Else
                Dim GROUP_CAPTION As String = ""
                Dim GROUP_CAPTION2 As String = ""
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Select("SEQUENCE = 1").FirstOrDefault
                If Not IsNothing(rowASTDSQLA) Then
                    GROUP_CAPTION = rowASTDSQLA.Item("COLUMN_CAPTION").ToString & String.Empty
                End If
                Dim rowASTDSQLA2 As DataRow = tblASTDSQLA.Select("SEQUENCE = 2").FirstOrDefault
                If Not IsNothing(rowASTDSQLA2) Then
                    GROUP_CAPTION2 = rowASTDSQLA2.Item("COLUMN_CAPTION").ToString & String.Empty
                End If
                Dim newICTISTA3 As DataRow = dst.Tables("ICTISTA3").NewRow
                newICTISTA3.Item("OPS_YYYYPP") = OPS_YYYYPP
                newICTISTA3.Item("GROUP_CODE") = GROUP_CODE_MULT(0)
                newICTISTA3.Item("GROUP_DESC") = GROUP_DESC.Key
                newICTISTA3.Item("GROUP_CAPTION") = GROUP_CAPTION
                newICTISTA3.Item("GROUP_CODE2") = GROUP_CODE_MULT(1)
                newICTISTA3.Item("GROUP_DESC2") = GROUP_DESC.Value
                newICTISTA3.Item("GROUP_CAPTION2") = GROUP_CAPTION2
                newICTISTA3.Item("OPTFILTERS") = OPTFILTERS
                newICTISTA3.Item("OPTASN") = optASN.Value
                newICTISTA3.Item("LAST_OPER") = ASCMAIN1.USER_ID
                newICTISTA3.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                newICTISTA3.Item("VAL_FULL") = VAL_FULL
                newICTISTA3.Item("OP_01") = OP_01
                newICTISTA3.Item("OP_02") = OP_02
                newICTISTA3.Item("OP_03") = OP_03
                newICTISTA3.Item("AGE_01") = AGE_01
                newICTISTA3.Item("AGE_02") = AGE_02
                newICTISTA3.Item("AGE_03") = AGE_03
                newICTISTA3.Item("QTY_01") = QTY_01
                newICTISTA3.Item("QTY_02") = QTY_02
                newICTISTA3.Item("QTY_03") = QTY_03
                newICTISTA3.Item("QTYOP_01") = QTYOP_01
                newICTISTA3.Item("QTYOP_02") = QTYOP_02
                newICTISTA3.Item("QTYOP_03") = QTYOP_03


                dst.Tables("ICTISTA3").Rows.Add(newICTISTA3)
            End If
            'End If
        Next
        'Next

    End Sub

    Private Function setOPTFILTERS() As String
        Dim retVal As String = ""
        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select()
            If rowASTDSQLA.Item("CODE_VALUES").ToString & String.Empty <> "" Then
                retVal = retVal & "| " & rowASTDSQLA.Item("COLUMN_CAPTION").ToString & String.Empty & ": " & rowASTDSQLA.Item("CODE_VALUES").ToString
            End If
        Next
        If retVal.Length > 2 Then
            retVal = retVal.Substring(2)
        End If
        Return retVal
    End Function

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE Is Not Null", "SEQUENCE").Length = 1 Then
                GROUP_CODE = tblASTDSQLA.Select("SEQUENCE Is Not Null", "SEQUENCE").FirstOrDefault.Item("COLUMN_NAME").ToString & String.Empty
            Else
                If tblASTDSQLA.Select("SEQUENCE Is Not Null", "SEQUENCE").Length = 2 Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Testing"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("Running This Report With")
                    iMSG.AppendLine("2 Sorts/Groups Is In Testing")
                    iMSG.AppendLine("Mode.  Please Review The")
                    iMSG.AppendLine("Results Carefully.")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Proceed?")

                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg &= vbCr & "2 Group Option Cancelled."
                    Else
                        For Each DR As DataRow In tblASTDSQLA.Select("SEQUENCE Is Not Null", "SEQUENCE")
                            GROUP_CODE_MULT.Add(DR.Item("COLUMN_NAME").ToString & String.Empty)
                        Next
                    End If
                Else
                    EMsg &= vbCr & "You Must Select One Or Two Fields To Group By At A Time."
                End If
            End If
            If (optFORMAT.Value = "C") Then
                If numDAYS1.Value >= numDAYS2.Value Then
                    EMsg &= vbCr & "1St Aging days must be a smaller number than 2nd Aging days"
                End If

                'If Absx1.cmbFor("RYP").Value = "" Then

                'End If
            End If
            If chkSHOWOD.Checked Then
                CUST_CODE_OD = ""
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("CUST_CODE")
                If Not IsNothing(rowASTDSQLA) Then
                    If rowASTDSQLA.Item("CODE_VALUES") & "" = "" Then
                        EMsg &= vbCr & "Yo Must Select A Customer When Order Dates Option Is Selected"
                    Else
                        If rowASTDSQLA.Item("CODE_VALUES").ToString.Contains(",") Then
                            EMsg &= vbCr & "You Can Only Select One Customer When Order Dates Option Is Selected"
                        End If
                        CUST_CODE_OD = rowASTDSQLA.Item("CODE_VALUES").ToString
                    End If
                End If
            End If
            If EMsg.Length = 0 Then
                tblICTISTA4 = dst.Tables.Item("ICTISTA4").Copy
            End If

        End If
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " And ROWNUM < 1"

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
        fillICTISTA4(False)
    End Sub

    Private Sub numDAYS1_ValueChanged(sender As System.Object, e As System.EventArgs)
        Set_Labels()
    End Sub

    Private Sub numDAYS2_ValueChanged(sender As System.Object, e As System.EventArgs)
        Set_Labels()
    End Sub

    Sub Set_Labels()
        If SELECTION_NO = 0 Then Exit Sub
        lblNUMDAYS1.Text = "Up To " & numDAYS1.Value & " Days"
        lblNUMDAYS2.Text = CStr(Val(numDAYS1.Value & "") + 1) & " - " & numDAYS2.Value & " Days"
        lblNUMDAYS3.Text = "Over " & numDAYS2.Value & " Days"

        With Absx1.optFor("OPTONLY").ValueList
            .ValueListItems(0).DisplayText = lblNUMDAYS1.Text
            .ValueListItems(1).DisplayText = lblNUMDAYS2.Text
            .ValueListItems(2).DisplayText = lblNUMDAYS3.Text
        End With
    End Sub

    Private Sub optFORMAT_ValueChanged(sender As System.Object, e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Options()
    End Sub

    Sub Setup_Options()
        'Stop
        'grpDATE_LAST_REC.Visible = (optFORMAT.Value = "C")
        'grpNUMDAYS.Visible = (optFORMAT.Value = "C")
        'grpShowOnly.Visible = (optFORMAT.Value = "C")
        'Absx1.chkFor("CHKCOST").Visible = Not (optFORMAT.Value = "C")
        'If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") _
        'Or (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
        '    chkOPEN_PICK.Visible = (optFORMAT.Value = "C")
        '    setPickDollarsOptions()
        'End If
        'If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") Then
        '    chkOPEN_PICK.Visible = (optFORMAT.Value = "C")
        '    setPickDollarsOptions()
        '    chkHIDEFIFOCOST.Visible = (optFORMAT.Value = "C")
        '    chkCOMBINEOPEN_PICK.Visible = (optFORMAT.Value = "A")
        '    chkCOMBINEOPEN_PICK.Checked = False
        '    chkSHOWFL.Checked = False
        '    chkSHOWFL.Visible = True
        '    chkSHOWRS.Visible = True
        '    chkSHOWOD.Visible = True
        '    lblSHOWOD.Visible = True
        'Else
        '    chkCOMBINEOPEN_PICK.Visible = False
        '    chkCOMBINEOPEN_PICK.Checked = False
        '    chkSHOWFL.Checked = False
        '    chkSHOWFL.Visible = False
        '    chkSHOWRS.Visible = False
        '    chkSHOWOD.Visible = False
        '    lblSHOWOD.Visible = False
        'End If

    End Sub

    Function Get_Dates() As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"DATE_LAST_REC"}
            Dim TABLE_NAME As String = "ICTCOSTX"
            'If COLUMN_NAME = "PICK_RELEASED" Then TABLE_NAME = "SOTPICK1"
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                sql = sql & " And " & TABLE_NAME & "." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
            End If
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                sql = sql & " and " & TABLE_NAME & "." & COLUMN_NAME & " <= '" & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "dd-MMM-yyyy") & "'"
            End If
        Next
        Return sql
    End Function

    Private Sub CalcFL()
        ASCMAIN1.Progress("Now calculating First / Last Dates")
        'dst.Tables("ASTSRPT1").Columns.Add("SHOWFL", GetType(System.String))
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select()
            If chkSHOWFL.Checked Then
                Dim SQL As New StringBuilder With {.Length = 0}
                SQL.AppendLine("SELECT (NVL(TO_CHAR(MIN(INV_DATE),'MM/DD/YY'),'')) AS FL")
                SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
                SQL.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
                SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
                SQL.AppendLine(String.Format("AND I2.STYLE_CODE = '{0}'", rowASTSRPT1.Item("STYLE_CODE").ToString()))
                SQL.AppendLine(String.Format("AND I2.COLOR_CODE = '{0}'", rowASTSRPT1.Item("COLOR_CODE").ToString()))
                ASCMAIN1.sql = SQL.ToString()
                Dim SHOWFL1 As String = ASCDATA1.GetDataValue
                rowASTSRPT1.Item("SHOWFL1") = SHOWFL1
                SQL.Length = 0
                SQL.AppendLine("SELECT (NVL(TO_CHAR(MAX(INV_DATE),'MM/DD/YY'),'')) AS FL")
                SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
                SQL.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
                SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
                SQL.AppendLine(String.Format("AND I2.STYLE_CODE = '{0}'", rowASTSRPT1.Item("STYLE_CODE").ToString()))
                SQL.AppendLine(String.Format("AND I2.COLOR_CODE = '{0}'", rowASTSRPT1.Item("COLOR_CODE").ToString()))
                ASCMAIN1.sql = SQL.ToString()
                Dim SHOWFL2 As String = ASCDATA1.GetDataValue
                rowASTSRPT1.Item("SHOWFL2") = SHOWFL2
            Else
                If chkSHOWRS.Checked Then
                    Dim SQL As New StringBuilder With {.Length = 0}
                    SQL.AppendLine("SELECT NVL(TO_CHAR(MAX(INV_DATE),'MM/DD/YY'),'') AS LS")
                    SQL.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2")
                    SQL.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
                    SQL.AppendLine("AND I1.INV_NO = I2.INV_NO")
                    SQL.AppendLine(String.Format("AND I2.STYLE_CODE = '{0}'", rowASTSRPT1.Item("STYLE_CODE").ToString()))
                    SQL.AppendLine(String.Format("AND I2.COLOR_CODE = '{0}'", rowASTSRPT1.Item("COLOR_CODE").ToString()))
                    ASCMAIN1.sql = SQL.ToString()
                    Dim LS As String = ASCDATA1.GetDataValue
                    rowASTSRPT1.Item("SHOWFL2") = LS
                    SQL.Length = 0
                    SQL.AppendLine("SELECT NVL(TO_CHAR(MAX(POTSHIP2.PO_DATE_RECEIVED),'MM/DD/YY'),'') PO_DATE_RECEIVED")
                    SQL.AppendLine("FROM POTORDR2, POTSHIP3, POTSHIP2")
                    SQL.AppendLine("WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
                    SQL.AppendLine("AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO")
                    SQL.AppendLine("AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO")
                    SQL.AppendLine("AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO")
                    SQL.AppendLine(String.Format("AND POTORDR2.STYLE_CODE = '{0}'", rowASTSRPT1.Item("STYLE_CODE").ToString()))
                    SQL.AppendLine(String.Format("AND POTORDR2.COLOR_CODE = '{0}'", rowASTSRPT1.Item("COLOR_CODE").ToString()))
                    ASCMAIN1.sql = SQL.ToString()
                    Dim LR As String = ASCDATA1.GetDataValue
                    rowASTSRPT1.Item("SHOWFL1") = LR
                Else
                    If chkSHOWOD.Checked Then
                        Dim SQL As New StringBuilder With {.Length = 0}
                        SQL.AppendLine("SELECT NVL(TO_CHAR(MIN(ORDR_DATE),'MM/DD/YY'),'') AS OD1")
                        SQL.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
                        SQL.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
                        SQL.AppendLine(String.Format("AND O1.CUST_CODE = '{0}'", CUST_CODE_OD))
                        SQL.AppendLine(String.Format("AND O2.STYLE_CODE = '{0}'", rowASTSRPT1.Item("STYLE_CODE").ToString()))
                        SQL.AppendLine(String.Format("AND O2.COLOR_CODE = '{0}'", rowASTSRPT1.Item("COLOR_CODE").ToString()))
                        ASCMAIN1.sql = SQL.ToString()
                        Dim OD1 As String = ASCDATA1.GetDataValue
                        rowASTSRPT1.Item("SHOWFL1") = OD1
                        SQL.Length = 0
                        SQL.AppendLine("SELECT NVL(TO_CHAR(MAX(ORDR_DATE),'MM/DD/YY'),'') AS OD2")
                        SQL.AppendLine("FROM SOTORDR1 O1, SOTORDR2 O2")
                        SQL.AppendLine("WHERE O1.ORDR_NO = O2.ORDR_NO")
                        SQL.AppendLine(String.Format("AND O1.CUST_CODE = '{0}'", CUST_CODE_OD))
                        SQL.AppendLine(String.Format("AND O2.STYLE_CODE = '{0}'", rowASTSRPT1.Item("STYLE_CODE").ToString()))
                        SQL.AppendLine(String.Format("AND O2.COLOR_CODE = '{0}'", rowASTSRPT1.Item("COLOR_CODE").ToString()))
                        ASCMAIN1.sql = SQL.ToString()
                        Dim OD2 As String = ASCDATA1.GetDataValue
                        rowASTSRPT1.Item("SHOWFL2") = OD2
                    End If
                End If
            End If

        Next
    End Sub

    Private Sub CalcExtCost()
        ASCMAIN1.Progress("Now calculating Extended Cost")
        Dim CANCEL_DATE As Date = DateSerial(2099, 1, 1)
        If chkLimitOP_C.Checked Then
            CANCEL_DATE = CDate(dteLimitOP_C.DateTime)
        End If
        Fill_Records("SOTRSRV2", CANCEL_DATE)
        Fill_Records("SOTCANCL", CANCEL_DATE)

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select()
            'If rowASTSRPT1.Item("STYLE_CODE").ToString() = "501286XIZ" And rowASTSRPT1.Item("COLOR_CODE").ToString() = "401" Then Stop

            Dim Filter As String = "STYLE_CODE = '" & rowASTSRPT1.Item("STYLE_CODE").ToString() & "' AND COLOR_CODE = '" & rowASTSRPT1.Item("COLOR_CODE").ToString() & "'"
            Dim rowICTCOSTX As DataRow = dst.Tables.Item("ICTCOSTX").Select(Filter).FirstOrDefault()
            Dim EXT_COST As Double = 0
            Dim OP_COST As Double = 0
            Dim QTYOH As Double = 0
            Dim QTYOP As Double = 0
            Dim AGE_01 As Double = 0
            Dim AGE_02 As Double = 0
            Dim AGE_03 As Double = 0
            Dim QTY_01 As Double = 0
            Dim QTY_02 As Double = 0
            Dim QTY_03 As Double = 0
            Dim QTYOP_01 As Double = 0
            Dim QTYOP_02 As Double = 0
            Dim QTYOP_03 As Double = 0
            If Not IsNothing(rowICTCOSTX) Then
                EXT_COST = Val(rowASTSRPT1.Item("WHSE_ON_HAND").ToString() & "") * Val(rowICTCOSTX.Item("STYLE_COST").ToString & "")
                QTYOH = Val(rowASTSRPT1.Item("WHSE_ON_HAND").ToString() & "")
                Dim WHSE_OPEN As Int64 = 0
                Dim WHSE_PICK As Int64 = 0
                Dim WHSE_RESV As Int64 = 0
                Dim WHSE_CANC As Int64 = 0
                Dim rowSOTRSRV2 As DataRow = dst.Tables("SOTRSRV2").Select(Filter).FirstOrDefault
                If Not IsNothing(rowSOTRSRV2) Then
                    WHSE_RESV = Val(rowSOTRSRV2.Item("RSRV_QTY_OPEN").ToString & String.Empty)
                End If

                Dim rowSOTCANCL As DataRow = dst.Tables("SOTCANCL").Select(Filter).FirstOrDefault
                If Not IsNothing(rowSOTCANCL) Then
                    WHSE_CANC = Val(rowSOTCANCL.Item("ORDR_QTY_OPEN").ToString & String.Empty)
                End If

                If chkLimitOP_P.Checked = True Then
                    WHSE_PICK = Val(rowASTSRPT1.Item("WHSE_PICK").ToString() & "")
                Else
                    WHSE_PICK = 0
                End If

                If chkLimitOP_O.Checked = True Then
                    If chkLimitOP_C.Checked Then
                        WHSE_OPEN = WHSE_CANC
                    Else
                        WHSE_OPEN = Val(rowASTSRPT1.Item("WHSE_OPEN").ToString() & "") - WHSE_RESV
                    End If
                    'WHSE_OPEN = Val(rowASTSRPT1.Item("WHSE_OPEN").ToString() & "") - WHSE_RESV
                Else
                    WHSE_OPEN = 0
                End If

                If chkLimitOP_R.Checked = False Then
                    WHSE_RESV = 0
                End If

                QTYOP = WHSE_OPEN + WHSE_RESV + WHSE_PICK
                OP_COST = (WHSE_OPEN + WHSE_RESV + WHSE_PICK) * Val(rowICTCOSTX.Item("STYLE_COST").ToString & "")
                'OP_COST = (Val(rowASTSRPT1.Item("WHSE_OPEN").ToString() & "") + Val(rowASTSRPT1.Item("WHSE_PICK").ToString() & "")) * Val(rowICTCOSTX.Item("STYLE_COST").ToString & "")

                'Dim FilterA As String = "STYLE_CODE = '" & rowASTSRPT1.Item("STYLE_CODE").ToString() & "'"
                Dim rowICTCOSTA As DataRow = dst.Tables.Item("ICTCOSTA").Select(Filter).FirstOrDefault()
                If Not IsNothing(rowICTCOSTA) Then
                    '        .Add("AGE_1", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)<=" & numDAYS1.Value & ",LOT_AMT_ONHD,0)")
                    '        .Add("AGE_2", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)>" & numDAYS1.Value & " AND ISNULL(LOT_DAYS,0)<=" & numDAYS2.Value & ",LOT_AMT_ONHD,0)")
                    '        .Add("AGE_3", GetType(System.Decimal), "IIF(ISNULL(LOT_DAYS,0)>" & numDAYS2.Value & ",LOT_AMT_ONHD,0)")
                    'DateDiff ("d",{ICTCOSTX.DATE_LAST_REC},{?BEGIN_AGE_DATE})

                    Dim BEGIN_AGE_DATE As Date = DateSerial(Now().Year, Now().Month, Now().Day)
                    If ASCMAIN1.CYP = RYP Then
                        BEGIN_AGE_DATE = DateSerial(Now().Year, Now().Month, Now().Day)
                    Else
                        BEGIN_AGE_DATE = DateSerial(RYP.Substring(0, 4), RYP.Substring(4, 2), 1).AddMonths(1).AddDays(-1)
                    End If
                    Dim DaysCnt As Integer = 0
                    If IsDate(rowICTCOSTX.Item("DATE_LAST_REC").ToString() & "") Then
                        DaysCnt = DateDiff("d", CDate(CDate(rowICTCOSTX.Item("DATE_LAST_REC").ToString()).ToShortDateString) & "", BEGIN_AGE_DATE)
                    End If

                    If DaysCnt <= numDAYS1.Value Then
                        AGE_01 = EXT_COST
                        QTY_01 = QTYOH
                        QTYOP_01 = QTYOP
                    End If
                    If DaysCnt > numDAYS1.Value And DaysCnt <= numDAYS2.Value Then
                        AGE_02 = EXT_COST
                        QTY_02 = QTYOH
                        QTYOP_02 = QTYOP
                    End If
                    If DaysCnt > numDAYS2.Value Then
                        AGE_03 = EXT_COST
                        QTY_03 = QTYOH
                        QTYOP_03 = QTYOP
                    End If
                End If
            End If
            rowASTSRPT1.Item("EXT_COST") = EXT_COST
            rowASTSRPT1.Item("OP_COST") = OP_COST
            rowASTSRPT1.Item("AGE_01") = AGE_01
            rowASTSRPT1.Item("AGE_02") = AGE_02
            rowASTSRPT1.Item("AGE_03") = AGE_03
            rowASTSRPT1.Item("QTY_01") = QTY_01
            rowASTSRPT1.Item("QTY_02") = QTY_02
            rowASTSRPT1.Item("QTY_03") = QTY_03
            rowASTSRPT1.Item("QTYOP_01") = QTYOP_01
            rowASTSRPT1.Item("QTYOP_02") = QTYOP_02
            rowASTSRPT1.Item("QTYOP_03") = QTYOP_03
        Next
    End Sub

    Private Sub chkSHOWFL_CheckedChanged(sender As Object, e As EventArgs)
        If chkSHOWFL.Checked Then
            chkSHOWRS.Checked = False
            chkSHOWOD.Checked = False
        End If
    End Sub

    Private Sub chkSHOWRS_CheckedChanged(sender As Object, e As EventArgs)
        If chkSHOWRS.Checked Then
            chkSHOWFL.Checked = False
            chkSHOWOD.Checked = False
        End If
    End Sub

    Private Sub chkSHOWOD_CheckedChanged(sender As Object, e As EventArgs)
        chkSHOWRS.Checked = False
        chkSHOWFL.Checked = False
    End Sub

    Private Sub chkOPEN_PICK_CheckedChanged(sender As Object, e As EventArgs)
        setPickDollarsOptions()
    End Sub

    Private Sub setPickDollarsOptions()
        'If chkOPEN_PICK.Checked Then
        '    chkOPEN_PICK_DOLLARS.Checked = True
        '    chkOPEN_PICK_UNITS.Checked = False
        '    chkOPEN_PICK_DOLLARS.Visible = True
        '    chkOPEN_PICK_UNITS.Visible = True
        'Else
        '    chkOPEN_PICK_DOLLARS.Checked = False
        '    chkOPEN_PICK_UNITS.Checked = False
        '    chkOPEN_PICK_DOLLARS.Visible = False
        '    chkOPEN_PICK_UNITS.Visible = False
        'End If
    End Sub

    Private Sub chkOPEN_PICK_DOLLARS_CheckedChanged(sender As Object, e As EventArgs)
        chkOPEN_PICK_UNITS.Checked = Not chkOPEN_PICK_DOLLARS.Checked
    End Sub

    Private Sub chkOPEN_PICK_UNITS_CheckedChanged(sender As Object, e As EventArgs)
        chkOPEN_PICK_DOLLARS.Checked = Not chkOPEN_PICK_UNITS.Checked
    End Sub

    Private Sub chkNEG_CheckedChanged(sender As Object, e As EventArgs)
        If chkNEG.Checked Then
            btnWIP.Visible = True
        Else
            btnWIP.Visible = False
            btnWIP.Text = "Neg OH"
            chkNEG.Text = "Show Style/Colors w/ Neg WIP Only"
        End If
    End Sub

    Private Sub btnWIP_Click(sender As Object, e As EventArgs)
        If btnWIP.Text = "Neg OH" Then
            btnWIP.Text = "Neg WIP"
            chkNEG.Text = "Show Style/Colors w/ Neg OH Only"
        Else
            btnWIP.Text = "Neg OH"
            chkNEG.Text = "Show Style/Colors w/ Neg WIP Only"
        End If
    End Sub

    Private Sub UltraCombo1_AfterCloseUp(sender As Object, e As EventArgs)
        fillICTISTA4(False)
    End Sub

    Private Sub ASFBASE1_Fill_Panel_Paint(sender As Object, e As PaintEventArgs) Handles ASFBASE1_Fill_Panel.Paint

    End Sub

    Private Sub chkLimitOP_C_CheckedChanged(sender As Object, e As EventArgs) Handles chkLimitOP_C.CheckedChanged
        If chkLimitOP_C.Checked Then
            dteLimitOP_C.Value = CDate(Now().ToShortDateString)
        Else
            dteLimitOP_C.Value = Null
        End If
    End Sub

    Private Sub AbsCheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CHKUNITSONLY.CheckedChanged

    End Sub
End Class