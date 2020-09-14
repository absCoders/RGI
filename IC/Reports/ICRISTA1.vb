Imports System.Text

Public Class ICRISTA1
    Dim ICTCOSTX As String
    Dim CUST_CODE_OD As String = ""
    Dim ICTCOSTA As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        RWU = "N"

        Range_Events(grpDATE_LAST_REC)

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
    End Sub

    Protected Overrides Sub Build_Workfile()

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

            ICTCOSTA = ""
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
                    {"WHSE_BEG", "WHSE_SHP", "WHSE_RTN", "WHSE_REC", "WHSE_ADJ", "WHSE_XFR", "WHSE_PHY", _
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
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add UPC_CODE VARCHAR2(12)")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add UPCS NUMBER (6,0)")


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

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select ICTUPCH1.STYLE_CODE, ICTUPCH1.COLOR_CODE, Min (ICTUPCH1.UPC_CODE) UPC_CODE, Count (*) UPCS" & vbCrLf _
            & "    from ICTUPCH1" & vbCrLf _
            & "   where ICTUPCH1.STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTCOSTX & ")" & vbCrLf _
            & "   group by ICTUPCH1.STYLE_CODE, ICTUPCH1.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
            & "    Set UPC_CODE = R1.UPC_CODE, UPCS = R1.UPCS" & vbCrLf _
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

        If Absx1.chkFor("CHKCOST").Checked And chkGenerateMarkdownSS.Checked Then
            Generate_Markdown_Spreadsheet()
        End If

    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        ASCMAIN1.Progress("Now Loading Style Activity")

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_COST, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.SUB_UNIT_PACK_QTY " _
            & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        'ASCMAIN1.sql = "Select Distinct ASTSRPT1.STYLE_CODE, ASTSRPT1.COLOR_CODE, ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
        '    & " from " & ASTSRPT1 & " ASTSRPT1,ICTSTYC1" & vbCrLf _
        '    & " where ICTSTYC1.STYLE_CODE (+) = ASTSRPT1.STYLE_CODE" & vbCrLf _
        '    & "   and  ICTSTYC1.COLOR_CODE (+) = ASTSRPT1.COLOR_CODE"
        'ICTCOSTX = ASCMAIN1.Temp_Table
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add Primary Key (STYLE_CODE, COLOR_CODE)")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_LAST_SHP DATE")
        'ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_LAST_REC DATE")

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

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            RPT = "ICRISTAV"
            dst.Tables("ASTSRPT1").Columns.Add("SHOWFL1", GetType(System.String))
            dst.Tables("ASTSRPT1").Columns.Add("SHOWFL2", GetType(System.String))
        End If
        dst.Tables("ASTSRPT1").Columns.Add("EXT_COST", GetType(System.Double))
        CalcExtCost()

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
                SUBT &= ", " & Z
            End If
            Page0.Add(Z)
        Next

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

        CR_params.Add("NEG", IIf(Absx1.chkFor("CHKNEG").Checked, "1", "0"))

        CR_params.Add("SUB1", IIf(Absx1.chkFor("CHK1").Checked, "1", "0"))
        CR_params.Add("SUB2", IIf(Absx1.chkFor("CHK2").Checked, "1", "0"))
        CR_params.Add("SUB3", IIf(Absx1.chkFor("CHK3").Checked, "1", "0"))
        CR_params.Add("SUB4", IIf(Absx1.chkFor("CHK4").Checked, "1", "0"))

        CR_params.Add("UD", Absx1.optFor("OPTUD").Value & "")
        CR_params.Add("TD", Absx1.optFor("OPTTD").Value & "")
        CR_params.Add("AD", Absx1.optFor("OPTAD").Value & "")

        Dim rsf As String = ""

        If Absx1.optFor("OPTAD").Value = "C" Then
            RPT = "ICRISTA3"
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


            If Absx1.optFor("OPTONLY").Value & "" <> "1" Then
                If Absx1.optFor("OPTONLY").Value & "" = "2" Then
                    rsf = "{@AGE_2} <> 0 or {@AGE_3} <> 0"
                Else
                    rsf = "{@AGE_3} <> 0"
                End If
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
        Else
            CR_params.Add("COST", IIf(Absx1.chkFor("CHKCOST").Checked, "1", "0"))
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                If chkCOMBINEOPEN_PICK.Checked Then
                    CR_params.Add("COMBINEOP", "1")
                Else
                    CR_params.Add("COMBINEOP", "0")
                End If
                If chkSHOWOD.Checked Then
                    CR_params.Add("SHOWOD", "1")
                    CalcFL()
                Else
                    CR_params.Add("SHOWOD", "0")
                End If

                If Absx1.chkFor("CHKTOTALSONLY").Checked Then
                    CR_params.Add("TOTALSONLY", "1")
                    CR_params.Add("SHOWFL", "0")
                    CR_params.Add("SHOWRS", "0")
                Else
                    CR_params.Add("TOTALSONLY", "0")

                    If chkSHOWFL.Checked Then
                        CR_params.Add("SHOWFL", "1")
                        CalcFL()
                    Else
                        CR_params.Add("SHOWFL", "0")
                    End If

                    If chkSHOWRS.Checked Then
                        CR_params.Add("SHOWRS", "1")
                        CalcFL()
                    Else
                        CR_params.Add("SHOWRS", "0")
                    End If

                End If
            Else
                If Absx1.chkFor("CHKTOTALSONLY").Checked Then
                    CR_params.Add("TOTALSONLY", "1")
                Else
                    CR_params.Add("TOTALSONLY", "0")
                End If
            End If
        End If

        Generate_Report(RPT, , SUBT, rsf)

        If ASCMAIN1.CLIENT = "VAN" Then
            Prepare_Data_Extracts()
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
            '    EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            'End If
            If (optFORMAT.Value = "C") Then
                If numDAYS1.Value >= numDAYS2.Value Then
                    EMsg &= vbCr & "1st Aging days must be a smaller number than 2nd Aging days"
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
                            EMsg &= vbCr & "Yo Can Only Select One Customer When Order Dates Option Is Selected"
                        End If
                        CUST_CODE_OD = rowASTDSQLA.Item("CODE_VALUES").ToString
                    End If
                End If
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

    Private Sub numDAYS1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles numDAYS1.ValueChanged
        Set_Labels()
    End Sub

    Private Sub numDAYS2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles numDAYS2.ValueChanged
        Set_Labels()
    End Sub

    Sub Set_Labels()
        If SELECTION_NO = 0 Then Exit Sub
        lblNUMDAYS1.Text = "Up to " & numDAYS1.Value & " Days"
        lblNUMDAYS2.Text = CStr(Val(numDAYS1.Value & "") + 1) & " - " & numDAYS2.Value & " Days"
        lblNUMDAYS3.Text = "Over " & numDAYS2.Value & " Days"

        With Absx1.optFor("OPTONLY").ValueList
            .ValueListItems(0).DisplayText = lblNUMDAYS1.Text
            .ValueListItems(1).DisplayText = lblNUMDAYS2.Text
            .ValueListItems(2).DisplayText = lblNUMDAYS3.Text
        End With
    End Sub

    Private Sub optFORMAT_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optFORMAT.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Options()
    End Sub

    Sub Setup_Options()
        grpDATE_LAST_REC.Visible = (optFORMAT.Value = "C")
        grpNUMDAYS.Visible = (optFORMAT.Value = "C")
        grpShowOnly.Visible = (optFORMAT.Value = "C")
        Absx1.chkFor("CHKCOST").Visible = Not (optFORMAT.Value = "C")
        If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") _
        Or (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
            chkOPEN_PICK.Visible = (optFORMAT.Value = "C")
            setPickDollarsOptions()
        End If
        If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") Then
            chkOPEN_PICK.Visible = (optFORMAT.Value = "C")
            setPickDollarsOptions()
            chkHIDEFIFOCOST.Visible = (optFORMAT.Value = "C")
            chkCOMBINEOPEN_PICK.Visible = (optFORMAT.Value = "A")
            chkCOMBINEOPEN_PICK.Checked = False
            chkSHOWFL.Checked = False
            chkSHOWFL.Visible = True
            chkSHOWRS.Visible = True
            chkSHOWOD.Visible = True
            lblSHOWOD.Visible = True
            grpLimitLastShipped.Visible = False
            chkLimitLastShipped.Checked = False
            chkLimitLastShippedB.Checked = False
            chkLimitLastShippedE.Checked = False
            dteLimitLastShippedB.DateTime = Now()
            dteLimitLastShippedE.DateTime = Now()
        Else
            chkCOMBINEOPEN_PICK.Visible = False
            chkCOMBINEOPEN_PICK.Checked = False
            chkSHOWFL.Checked = False
            chkSHOWFL.Visible = False
            chkSHOWRS.Visible = False
            chkSHOWOD.Visible = False
            lblSHOWOD.Visible = False
            grpLimitLastShipped.Visible = False
            chkLimitLastShipped.Checked = False
        End If

    End Sub

    Function Get_Dates() As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"DATE_LAST_REC"}
            Dim TABLE_NAME As String = "ICTCOSTX"
            'If COLUMN_NAME = "PICK_RELEASED" Then TABLE_NAME = "SOTPICK1"
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                sql = sql & " and " & TABLE_NAME & "." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
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
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim SC As String = rowASTSRPT1.Item("STYLE_CODE").ToString() & "-" & rowASTSRPT1.Item("STYLE_CODE").ToString()
            ASCMAIN1.Progress("-", SC)
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
                If chkLimitLastShipped.Checked Then
                    If IsDate(SHOWFL2) Then
                        Dim SHOWFL2_DATE = CDate(SHOWFL2)
                        Dim ExcludeStyle As Boolean = False
                        If chkLimitLastShippedB.Checked = False Then
                            If SHOWFL2_DATE < CDate(dteLimitLastShippedB.DateTime.ToShortDateString) Then
                                ExcludeStyle = True
                            End If
                        End If
                        If chkLimitLastShippedE.Checked = False Then
                            If SHOWFL2_DATE > CDate(dteLimitLastShippedE.DateTime.ToShortDateString) Then
                                ExcludeStyle = True
                            End If
                        End If
                        If ExcludeStyle Then
                            rowASTSRPT1.Delete()
                        End If
                    End If
                End If
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
                    If chkLimitLastShipped.Checked Then
                        If IsDate(LS) Then
                            Dim SHOWFL2_DATE = CDate(LS)
                            Dim ExcludeStyle As Boolean = False
                            If chkLimitLastShippedB.Checked = False Then
                                If SHOWFL2_DATE < CDate(dteLimitLastShippedB.DateTime.ToShortDateString) Then
                                    ExcludeStyle = True
                                End If
                            End If
                            If chkLimitLastShippedE.Checked = False Then
                                If SHOWFL2_DATE > CDate(dteLimitLastShippedE.DateTime.ToShortDateString) Then
                                    ExcludeStyle = True
                                End If
                            End If
                            If ExcludeStyle Then
                                rowASTSRPT1.Delete()
                            End If
                        End If
                    End If
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
    Private Sub Generate_Markdown_Spreadsheet()
        ASCMAIN1.Progress("Now Creating Workbook")

        '  Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
        Dim DataTable As DataTable
        Dim r As Integer = 0
        Dim c As Integer = 0

        Dim ssgx As String = ASCMAIN1.Folders("Work") & "Markdowns_" & XNO & ".xlsX"

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook() '(FILENAME)
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePasteTo As SpreadsheetGear.IRange = Nothing

        'build datatable 
        ASCMAIN1.sql = "Select W.OPS_YYYYPP, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.SUB_BODY_CODE, ICTSTYL1.FABRIC_CODE" & vbCrLf _
            & ", W.STYLE_CODE, W.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & ", W.STYLE_COST COST_NOW, W.WHSE_QTY_ON_HAND ON_HAND, W.LOT_AMT_ONHD VALUE_NOW" & vbCrLf _
            & ", MKDN.TRAN_COST COST_NEW, NULL VALUE_NEW, NULL MARKDOWN" & vbCrLf _
            & ", ICTCOSTX.DATE_LAST_SHP, ICTCOSTX.DATE_LAST_REC" & vbCrLf _
            & " from " & ICTCOSTA & " W, ICTSTYL1, " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
            & ", (SELECT STYLE_CODE, COLOR_CODE, MIN(TRAN_COST) TRAN_COST FROM ICTCOST1 WHERE OPS_YYYYPP = '" & RYP & "' AND TRAN_TYPE = 'M' GROUP BY STYLE_CODE, COLOR_CODE) MKDN" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE = W.STYLE_CODE " & vbCrLf _
            & "   and ICTCOSTX.STYLE_CODE = W.STYLE_CODE and ICTCOSTX.COLOR_CODE = W.COLOR_CODE" & vbCrLf _
            & "   and MKDN.STYLE_CODE (+) = W.STYLE_CODE and MKDN.COLOR_CODE (+) = W.COLOR_CODE" & vbCrLf _
            & " order by OPS_YYYYPP,SALES_DIVISION_CODE,SUB_BODY_CODE, FABRIC_CODE, STYLE_CODE, COLOR_CODE"

        'ASCMAIN1.sql = "Select W.OPS_YYYYPP, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.SUB_BODY_CODE, ICTSTYL1.FABRIC_CODE" & vbCrLf _
        '    & ", W.STYLE_CODE, W.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
        '    & ", W.STYLE_COST COST_NOW, W.WHSE_QTY_ON_HAND ON_HAND, W.LOT_AMT_ONHD VALUE_NOW" & vbCrLf _
        '    & ", NULL COST_NEW, NULL VALUE_NEW, NULL MARKDOWN" & vbCrLf _
        '    & ", DTS.DATE_LAST_SHP, DTS.DATE_LAST_REC" & vbCrLf _
        '    & " from " & ICTCOSTA & " W, ICTSTYL1, " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
        '    & " (Select STYLE_CODE, COLOR_CODE, Max (SHOWFL1) DATE_LAST_SHP, Max (SHOWFL2) DATE_LAST_REC from " & ASTSRPT1 & vbCrLf _
        '    & " group by STYLE_CODE, COLOR_CODE) DTS" & vbCrLf _
        '    & " where ICTSTYL1.STYLE_CODE = W.STYLE_CODE " & vbCrLf _
        '    & "   and ICTCOSTX.STYLE_CODE = W.STYLE_CODE and ICTCOSTX.COLOR_CODE = W.COLOR_CODE" & vbCrLf _
        '    & " order by OPS_YYYYPP,SALES_DIVISION_CODE,SUB_BODY_CODE, FABRIC_CODE, STYLE_CODE, COLOR_CODE"

        DataTable = ASCDATA1.GetDataTable
        Dim cdr As Integer = 0

        Dim R0 As Integer = 1 ' 0 based starting row for headings just prior to data

        Dim COLS As Integer = DataTable.Columns.Count
        Dim ROWS As Integer = DataTable.Rows.Count

        worksheet.Cells(0, 0, 0, 7).EntireColumn.NumberFormat = "@"

        range = worksheet.Range(R0, 0)
        range.CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.None)

        worksheet.UsedRange.Columns.AutoFit()
        worksheet.Range(R0, 0, R0, COLS - 1).Interior.Color = SpreadsheetGear.Colors.AliceBlue


        worksheet.Cells(R0 + 1, 0, R0 + ROWS, COLS - 1).Interior.Color = SpreadsheetGear.Colors.Linen

        Dim col As Integer

        col = 7
        Dim col_COST_NOW As Integer = col
        worksheet.Cells(0, 0, 0, col).EntireColumn.NumberFormat = "@"
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "#,##0.0000"

        col = 8
        Dim col_ON_HAND As Integer = col
        worksheet.Cells(0, col).Formula = "=SUBTOTAL(9," & Excel_Cell0(R0 + 1, col) & ":" & Excel_Cell0(R0 + ROWS, col) & ")"
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "#,##0"

        col = 9
        Dim col_VALUE_NOW As Integer = col
        worksheet.Cells(0, col).Formula = "=SUBTOTAL(9," & Excel_Cell0(R0 + 1, col) & ":" & Excel_Cell0(R0 + ROWS, col) & ")"
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "#,##0.00"

        worksheet.Cells(R0 + 1, col).Formula = "=" & Excel_Cell0(R0 + 1, col_ON_HAND) & "*" & Excel_Cell0(R0 + 1, col_COST_NOW)
        rangeCopyFrom = worksheet.Cells(R0 + 1, col)
        rangePasteTo = worksheet.Cells(R0 + 1, col, R0 + ROWS, col)
        rangeCopyFrom.Copy(rangePasteTo, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

        col = 10
        Dim col_COST_NEW As Integer = col
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "#,##0.0000"
        worksheet.Cells(0, col).EntireColumn.Locked = False

        col = 11
        Dim col_VALUE_NEW As Integer = col
        worksheet.Cells(0, col).Formula = "=SUBTOTAL(9," & Excel_Cell0(R0 + 1, col) & ":" & Excel_Cell0(R0 + ROWS, col) & ")"
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "#,##0.00"

        worksheet.Cells(R0 + 1, col).Formula = "=IF(" & Excel_Cell0(R0 + 1, col_COST_NEW) & "=0," & Excel_Cell0(R0 + 1, col_VALUE_NOW) & "," & Excel_Cell0(R0 + 1, col_COST_NEW) & "*" & Excel_Cell0(R0 + 1, col_ON_HAND) & ")"
        rangeCopyFrom = worksheet.Cells(R0 + 1, col)
        rangePasteTo = worksheet.Cells(R0 + 1, col, R0 + ROWS, col)
        rangeCopyFrom.Copy(rangePasteTo, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

        worksheet.Cells(R0 + 1, col_COST_NEW, R0 + ROWS, col_COST_NEW).Interior.Color = SpreadsheetGear.Colors.AliceBlue

        col = 12
        Dim col_MARKDOWN As Integer = 12
        worksheet.Cells(0, col).Formula = "=SUBTOTAL(9," & Excel_Cell0(R0 + 1, col) & ":" & Excel_Cell0(R0 + ROWS, col) & ")"
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "#,##0.00"

        worksheet.Cells(R0 + 1, col).Formula = "=" & Excel_Cell0(R0 + 1, col_VALUE_NEW) & "-" & Excel_Cell0(R0 + 1, col_VALUE_NOW)
        rangeCopyFrom = worksheet.Cells(R0 + 1, col)
        rangePasteTo = worksheet.Cells(R0 + 1, col, R0 + ROWS, col)
        rangeCopyFrom.Copy(rangePasteTo, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)


        col = 13
        Dim col_DATE_LAST_SHP As Integer = 13
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "mm/dd/yyyy"
        worksheet.Cells(0, col).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center

        col = 14
        Dim col_DATE_LAST_REC As Integer = 14
        worksheet.Cells(0, col).EntireColumn.NumberFormat = "mm/dd/yyyy"
        worksheet.Cells(0, col).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center

        ' Headings
        worksheet.Range(R0, 0, R0, COLS - 1).AutoFilter()
        For CX As Integer = 0 To COLS - 1
            worksheet.Cells(R0, CX).EntireColumn.ColumnWidth *= 1.25
        Next

        worksheet.Cells(R0 + 1, 0).Activate()
        worksheet.WindowInfo.FreezePanes = True

        worksheet.Protect("")

        workbook.SaveAs(ssgx, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        range = Nothing
        worksheet = Nothing
        workbook = Nothing

        Show_Document(ssgx)

    End Sub

    Private Sub CalcExtCost()
        ASCMAIN1.Progress("Now calculating Extended Cost")
        'dst.Tables("ASTSRPT1").Columns.Add("SHOWFL", GetType(System.String))

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select()
            'If rowASTSRPT1.Item("STYLE_CODE").ToString() = "0100IZ" And rowASTSRPT1.Item("COLOR_CODE").ToString() = "270" Then

            'End If
            Dim Filter As String = "STYLE_CODE = '" & rowASTSRPT1.Item("STYLE_CODE").ToString() & "' AND COLOR_CODE = '" & rowASTSRPT1.Item("COLOR_CODE").ToString() & "'"
            Dim rowICTCOSTX As DataRow = dst.Tables.Item("ICTCOSTX").Select(Filter).FirstOrDefault()
            Dim EXT_COST As Double = 0
            If Not IsNothing(rowICTCOSTX) Then
                EXT_COST = Val(rowASTSRPT1.Item("WHSE_ON_HAND").ToString() & "") * Val(rowICTCOSTX.Item("STYLE_COST").ToString & "")
            End If
            rowASTSRPT1.Item("EXT_COST") = EXT_COST
        Next
    End Sub

    Private Sub chkSHOWFL_CheckedChanged(sender As Object, e As EventArgs) Handles chkSHOWFL.CheckedChanged
        chkLimitLastShipped.Checked = False
        chkLimitLastShippedB.Checked = False
        chkLimitLastShippedE.Checked = False
        If chkSHOWFL.Checked Then
            chkSHOWRS.Checked = False
            chkSHOWOD.Checked = False
            grpLimitLastShipped.Visible = True
            grpLimitLastShipped.Text = "Limit To Styles Last Shipped"
        Else
            grpLimitLastShipped.Visible = False
        End If
    End Sub

    Private Sub chkSHOWRS_CheckedChanged(sender As Object, e As EventArgs) Handles chkSHOWRS.CheckedChanged
        If chkSHOWRS.Checked Then
            chkSHOWFL.Checked = False
            chkSHOWOD.Checked = False
            grpLimitLastShipped.Visible = True
            grpLimitLastShipped.Text = "Limit To Styles Last Shipped"
        Else
            grpLimitLastShipped.Visible = False
        End If
    End Sub

    Private Sub chkSHOWOD_CheckedChanged(sender As Object, e As EventArgs) Handles chkSHOWOD.CheckedChanged
        chkSHOWRS.Checked = False
        chkSHOWFL.Checked = False
    End Sub

    Private Sub chkOPEN_PICK_CheckedChanged(sender As Object, e As EventArgs) Handles chkOPEN_PICK.CheckedChanged
        setPickDollarsOptions()
    End Sub

    Private Sub setPickDollarsOptions()
        If chkOPEN_PICK.Checked Then
            chkOPEN_PICK_DOLLARS.Checked = True
            chkOPEN_PICK_UNITS.Checked = False
            chkOPEN_PICK_DOLLARS.Visible = True
            chkOPEN_PICK_UNITS.Visible = True
        Else
            chkOPEN_PICK_DOLLARS.Checked = False
            chkOPEN_PICK_UNITS.Checked = False
            chkOPEN_PICK_DOLLARS.Visible = False
            chkOPEN_PICK_UNITS.Visible = False
        End If
    End Sub

    Private Sub chkOPEN_PICK_DOLLARS_CheckedChanged(sender As Object, e As EventArgs) Handles chkOPEN_PICK_DOLLARS.CheckedChanged
        chkOPEN_PICK_UNITS.Checked = Not chkOPEN_PICK_DOLLARS.Checked
    End Sub

    Private Sub chkOPEN_PICK_UNITS_CheckedChanged(sender As Object, e As EventArgs) Handles chkOPEN_PICK_UNITS.CheckedChanged
        chkOPEN_PICK_DOLLARS.Checked = Not chkOPEN_PICK_UNITS.Checked
    End Sub

    Private Sub chkNEG_CheckedChanged(sender As Object, e As EventArgs) Handles chkNEG.CheckedChanged
        If chkNEG.Checked Then
            btnWIP.Visible = True
        Else
            btnWIP.Visible = False
            btnWIP.Text = "Neg OH"
            chkNEG.Text = "Show Style/Colors w/ Neg WIP Only"
        End If
    End Sub

    Private Sub btnWIP_Click(sender As Object, e As EventArgs) Handles btnWIP.Click
        If btnWIP.Text = "Neg OH" Then
            btnWIP.Text = "Neg WIP"
            chkNEG.Text = "Show Style/Colors w/ Neg OH Only"
        Else
            btnWIP.Text = "Neg OH"
            chkNEG.Text = "Show Style/Colors w/ Neg WIP Only"
        End If
    End Sub

    Private Sub chkShowCosts_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowCosts.CheckedChanged
        Absx1.chkFor("CHKGENERATE_MKDN_SS").Visible = ((ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") And chkShowCosts.Visible And chkShowCosts.Checked)
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        With dst.Tables("ASTSRPT1").Columns
            '  .Add("STYLE_CODE")
            .Add("STYLE_DESC")
            .Add("CARTON_PACK_QTY", GetType(System.Int64))
            .Add("SUB_UNIT_PACK_QTY", GetType(System.Int64))

            .Add("DATE_LAST_SHP", GetType(System.DateTime))
            .Add("DATE_LAST_REC", GetType(System.DateTime))
            .Add("STYLE_COST", GetType(System.Decimal))

            .Add("DATE_FRST_SHP", GetType(System.DateTime))
            .Add("DATE_FRST_REC", GetType(System.DateTime))
            .Add("UPC_CODE")
            .Add("UPCS", GetType(System.Int32))

            '.Add("WHSE_REC", GetType(System.Int64))
            '.Add("WHSE_ON_HAND", GetType(System.Int64))
            '.Add("WHSE_PICK", GetType(System.Int64))

            .Add("NET_SHP", GetType(System.Int64), "ISNULL(WHSE_SHP,0)-ISNULL(WHSE_RTN,0)")
            .Add("NET_ADJ", GetType(System.Int64), "ISNULL(WHSE_ADJ,0)+ISNULL(WHSE_PHY,0)")

            .Add("NET_WIP", GetType(System.Int64), "ISNULL(WHSE_TRAN,0)+ISNULL(WHSE_ON_ORDER,0)")

            .Add("OTS_ONH", GetType(System.Int64), "ISNULL(WHSE_ON_HAND,0)-ISNULL(WHSE_PICK,0)")
            .Add("NET_POS", GetType(System.Int64), "ISNULL(WHSE_ON_HAND,0)-ISNULL(WHSE_PICK,0)+ISNULL(WHSE_ON_ORDER,0)+ISNULL(WHSE_TRAN,0)-ISNULL(WHSE_OPEN,0)")
        End With

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("")
            For I As Integer = 1 To COLUMN_NAMEs.Count
                Dim CODE_VALUE As String = row.Item("G" & CStr(I))
                row.Item("G" & CStr(I)) = Split(CODE_VALUE, ":")(1)
            Next
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
            row.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            row.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
            row.Item("SUB_UNIT_PACK_QTY") = rowICTSTYL1.Item("SUB_UNIT_PACK_QTY")

            Dim COLOR_CODE As String = row.Item("COLOR_CODE")

            Dim rowICTCOSTX As DataRow = dst.Tables("ICTCOSTX").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If rowICTCOSTX IsNot Nothing Then
                For Each C As String In New String() {"DATE_LAST_SHP", "DATE_LAST_REC", "UPC_CODE", "UPCS"}
                    ' {"DATE_LAST_SHP", "DATE_LAST_REC", "STYLE_COST", "DATE_FRST_SHP", "DATE_FRST_REC", "UPC_CODE", "UPCS"}
                    row.Item(C) = rowICTCOSTX.Item(C)
                Next
            End If
        Next

        grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")

        grdASTEXPT1.Text = MENU_ITEM_DESC

        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , Color.Gold)
            grdASTEXPT1.DisplayLayout.Bands(0).Columns("G" & CStr(G)).Header.Fixed = True
        Next

        Set_DX_Column(grdASTEXPT1, "STYLE_CODE", "Style Code", 120)
        Set_DX_Column(grdASTEXPT1, "STYLE_DESC", "Description", 200)
        Set_DX_Column(grdASTEXPT1, "COLOR_CODE", "Color", 60)

        Set_DX_Column(grdASTEXPT1, "CARTON_PACK_QTY", "#/Ctn", 50, "##0", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "SUB_UNIT_PACK_QTY", "Pcs", 40, "##0", , Color.Pink)

        Set_DX_Column(grdASTEXPT1, "UPC_CODE", "UPC Code", 120, , , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "UPCS", "#UPCs", 50, "##0", , Color.Pink)


        Set_DX_Column(grdASTEXPT1, "STYLE_COST", "Unit Cost", 90, "#.0000", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "WHSE_REC", "MTD Rec", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "NET_SHP", "MTD Sls", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "NET_ADJ", "MTD Adj", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_ON_HAND", "On Hand", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_PICK", "In Pick", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "OTS_ONH", "OTS", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "WHSE_TRAN", "In Transit", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "WHSE_ON_ORDER", "On PO", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_OPEN", "Open Orders", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "NET_WIP", "In-Xit+PO", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "NET_POS", "OTS WIP", 90, "#,##0", , Color.LightBlue)

        'For I As Integer = 1 To 4
        '    Set_DX_Column(grdASTEXPT1, "ADD_" & Format(I, "0"), "Fut Qty " & CStr(I), 90, "#,##0", , Color.LightGreen)
        '    Set_DX_Column(grdASTEXPT1, "DATE_" & Format(I, "0"), "Date " & CStr(I), 90, "MM/dd/yy", , Color.LightGreen)
        'Next

        'Set_DX_Column(grdASTEXPT1, "DATE_FRST_SHP", "1st Shp (2yr)", 90, "MM/dd/yy", , Color.Orange)
        'Set_DX_Column(grdASTEXPT1, "DATE_LAST_REC", "Last Rec", 90, "MM/dd/yy", , Color.Orange)
        'Set_DX_Column(grdASTEXPT1, "DATE_LAST_SHP", "Last Shp", 90, "MM/dd/yy", , Color.Orange)

        If chkSHOWFL.Checked Then
            Set_DX_Column(grdASTEXPT1, "SHOWFL1", "First Shp", 90, "MM/dd/yy", , Color.Orange)
            Set_DX_Column(grdASTEXPT1, "SHOWFL2", "Last Shp", 90, "MM/dd/yy", , Color.Orange)
        Else
            If chkSHOWRS.Checked Then
                Set_DX_Column(grdASTEXPT1, "SHOWFL1", "Last Rec", 90, "MM/dd/yy", , Color.Orange)
                Set_DX_Column(grdASTEXPT1, "SHOWFL2", "Last Shp", 90, "MM/dd/yy", , Color.Orange)
            Else
                If chkSHOWOD.Checked Then
                    Set_DX_Column(grdASTEXPT1, "SHOWFL1", "First Ord", 90, "MM/dd/yy", , Color.Orange)
                    Set_DX_Column(grdASTEXPT1, "SHOWFL2", "Last Ord", 90, "MM/dd/yy", , Color.Orange)
                End If
            End If
        End If

        grdASTEXPT1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "STYLE_CODE")

    End Sub

    Overrides Sub grdASTEXPT1_InitializeRow_Custom(sender As Object, e As UltraWinGrid.InitializeRowEventArgs)
        If e.Row.IsDataRow Then
            Dim NET_POS As Int64 = Val(e.Row.Cells("NET_POS").Value & "")
            If NET_POS < 0 Then
                e.Row.CellAppearance.ForeColor = Color.Red
            Else
                e.Row.CellAppearance.ForeColor = Color.Empty
            End If
        End If
    End Sub

    Private Sub chkLimitLastShippedB_CheckedChanged(sender As Object, e As EventArgs) Handles chkLimitLastShippedB.CheckedChanged
        If chkLimitLastShippedB.Checked Then
            dteLimitLastShippedB.Value = Null
        Else
            dteLimitLastShippedB.DateTime = Now()
        End If
    End Sub

    Private Sub chkLimitLastShippedE_CheckedChanged(sender As Object, e As EventArgs) Handles chkLimitLastShippedE.CheckedChanged
        If chkLimitLastShippedE.Checked Then
            dteLimitLastShippedE.Value = Null
        Else
            dteLimitLastShippedE.DateTime = Now()
        End If
    End Sub
End Class