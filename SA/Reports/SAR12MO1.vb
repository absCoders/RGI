Public Class SAR12MO1
    Dim PTD1 As Int16 = 1
    Dim PTD2 As Int16 = 12

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 12, -1)

        grpShow.Visible = (ASCMAIN1.DBS_COMPANY = "CPU") ' OPTION PROVIDED FOR CPU
        SplitContainer5.Panel2Collapsed = False ' probably should have not put this splitter and grid in stds - probably should have implemented the whole thing here
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        PTD1 = utb1.Value
        PTD2 = utb2.Value

        Dim YEAR_END As Integer = -99
        Dim YEAR_BEGIN As Integer = 99

        'For Each row As DataRow In dst.Tables("ASTRECAP").Rows
        '    Dim YEAR As Integer = 0
        '    If row.Item("YEAR") & "" = "N" Then
        '        YEAR = -1
        '    Else
        '        YEAR = Val(row.Item("YEAR") & "")
        '    End If
        '    If YEAR_END < YEAR Then
        '        YEAR_END = YEAR
        '    End If
        '    If YEAR_BEGIN > YEAR Then
        '        YEAR_BEGIN = YEAR
        '    End If
        'Next

        For M As Integer = 1 To 12
            For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("COLUMN_NAME = 'M" & Format(M, "00") & "'")
                Dim YP As String = ASCMAIN1.Period_Calc(RYP, -12 + M)
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YP)
                rowASTDSQLS.Item("COLUMN_CAPTION") = Mid(LEGEND, 10, 6)
            Next
        Next


        Dim ICTITEM1 As String = "ICTITEM1"
        If chkHISTCAT.Checked Then
            ICTITEM1 = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP, True)
        End If

        Dim SAT12MO1 As String = ""

        Dim SAT12MO1_ACN_R As String = ""
        Dim SAT12MO1_ACN_W As String = ""
        If optACN.Value <> "A" Then
            Dim YP_MIN As String = ""
            Dim YP_MAX As String = ""
            For Each rowASTRECAP As DataRow In tblASTRECAP.Rows
                Dim GRN As String = rowASTRECAP.Item("GRN")
                If GRN = "T" Then
                    SAT12MO1_ACN_R = "*"
                Else
                    SAT12MO1_ACN_W = "*"
                End If
                Dim YEAR As String = rowASTRECAP.Item("YEAR")
                Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))
                Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
                Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)
                If YP_01 < YP_MIN Or YP_MIN = "" Then YP_MIN = YP_01
                If YP_12 > YP_MAX Or YP_MAX = "" Then YP_MAX = YP_12
            Next
            If YP_MAX > ASCMAIN1.CYP Then
                YP_MAX = ASCMAIN1.CYP
            End If

            Dim MOS As Int32 = ASCMAIN1.Period_Diff(YP_MIN, YP_MAX) + 1

            If SAT12MO1_ACN_R = "*" Then
                ASCMAIN1.sql = "" _
                & "Select CUST_CODE, CUST_STORE_NO" _
                & ", COUNT (DISTINCT OPS_YYYYPP) MOS" _
                & " from RSTRETL1 where AMT_SOLD > 0" _
                & " and OPS_YYYYPP between '" & YP_MIN & "' and '" & YP_MAX & "'" _
                & IIf(SQLA("CUST_CODE") <> "", " and CUST_CODE in (" & SQLA("CUST_CODE", , True) & ")", "") _
                & " group by CUST_CODE, CUST_STORE_NO"
                SAT12MO1_ACN_R = ASCMAIN1.Temp_Table

                ASCDATA1.ExecuteSQL("Delete from " & SAT12MO1_ACN_R & " where MOS <> " & CStr(MOS))
                ASCDATA1.ExecuteSQL("Alter Table " & SAT12MO1_ACN_R & " Add Primary Key (CUST_CODE, CUST_STORE_NO)")
            End If

            If SAT12MO1_ACN_W = "*" Then
                ASCMAIN1.sql = "" _
                & "Select CUST_CODE, CUST_STORE_NO" _
                & ", COUNT (DISTINCT ORDR_YYYYPP_UPDATED) MOS" _
                & " from SOTINVH2 where ORDR_QTY_SHIP > 0" _
                & " and ORDR_YYYYPP_UPDATED between '" & YP_MIN & "' and '" & YP_MAX & "'" _
                & IIf(SQLA("CUST_CODE") <> "", " and CUST_CODE in (" & SQLA("CUST_CODE", , True) & ")", "") _
                & " group by CUST_CODE, CUST_STORE_NO"
                SAT12MO1_ACN_W = ASCMAIN1.Temp_Table

                ASCDATA1.ExecuteSQL("Delete from " & SAT12MO1_ACN_W & " where MOS <> " & CStr(MOS))
                ASCDATA1.ExecuteSQL("Alter Table " & SAT12MO1_ACN_W & " Add Primary Key (CUST_CODE, CUST_STORE_NO)")
            End If

        End If

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        'Stop

        Dim SOURCE_TABLE_NAME As String = ""
        Dim by_Item As Boolean = False
        If COLUMN_NAMEs.Contains("ITEM_CODE") Then ' THIS NEEDS TO BE EXPANDED UPON
            by_Item = True
        End If
        Dim by_Store As Boolean = False
        If COLUMN_NAMEs.Contains("CUST_STORE_NO") Then ' THIS NEEDS TO BE EXPANDED UPON
            by_Store = True
        End If

        by_Item = True
        by_Store = True


        ' Wholesale Shipments

        If by_Item And by_Store Then
            SOURCE_TABLE_NAME = "SOTINVH2"
        Else
            If Not by_Item And Not by_Store Then
                SOURCE_TABLE_NAME = "SATSSUM0"
            Else
                If by_Item Then
                    SOURCE_TABLE_NAME = "SATSSUMI"
                Else
                    SOURCE_TABLE_NAME = "SATSSUMS"
                End If
            End If
        End If

        If chkHISTCAT.Checked Then
            sql_TABLE_NAMEs = Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1 & " ICTITEM1")
        End If

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            Dim YP As String = ""
            Dim COLUMN_NAME As String = ""
            Dim sql_GRN As String = ""

            If (GRN = "G" Or GRN = "R" Or GRN = "N" Or GRN = "T") And (USC = "U" Or USC = "S" Or USC = "C" Or USC = "G") Then

                If GRN = "T" Then
                    sql_GRN = ""
                    YP = ""
                    COLUMN_NAME = ""
                    YP = "SAT12MO1.OPS_YYYYPP"
                    If USC = "U" Then
                        COLUMN_NAME = "NVL(SAT12MO1.QTY_SOLD,0)"
                    ElseIf USC = "S" Then
                        COLUMN_NAME = "NVL(SAT12MO1.AMT_SOLD,0)"
                    ElseIf USC = "C" Then
                        COLUMN_NAME = "0"
                    ElseIf USC = "G" Then
                        COLUMN_NAME = "0"
                    End If
                Else

                    sql_GRN = ""
                    If GRN = "G" Then
                        sql_GRN = " AND SAT12MO1.INV_TYPE = 'I'"
                    ElseIf GRN = "R" Then
                        sql_GRN = " AND SAT12MO1.INV_TYPE = 'C'"
                    End If

                    YP = ""
                    COLUMN_NAME = ""
                    Select Case SOURCE_TABLE_NAME
                        Case "SOTINVH2"
                            YP = "SAT12MO1.ORDR_YYYYPP_UPDATED"
                            If USC = "U" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0)"
                            ElseIf USC = "S" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0) * NVL(SAT12MO1.ORDR_UNIT_PRICE,0)"
                            ElseIf USC = "C" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0) * NVL(SAT12MO1.ITEM_UNIT_COST,0)"
                            ElseIf USC = "G" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0) * (NVL(SAT12MO1.ORDR_UNIT_PRICE,0) - NVL(SAT12MO1.ITEM_UNIT_COST,0))"
                            End If
                        Case Else
                            YP = "SAT12MO1.OPS_YYYYPP"
                            If USC = "U" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0)"
                            ElseIf USC = "S" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_AMT_SHIP,0)"
                            ElseIf USC = "C" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_CGS_SHIP,0)"
                            ElseIf USC = "G" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_AMT_SHIP,0) - NVL(SAT12MO1.ORDR_CGS_SHIP,0)"
                            End If
                    End Select
                End If

                sql_filter = " and " & YP & " BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'" & vbCrLf _
                & " and " & COLUMN_NAME & " <> 0"

                If optACN.Value <> "A" And IIf(GRN = "T", SAT12MO1_ACN_R, SAT12MO1_ACN_W) <> "" Then
                    sql_filter &= vbCrLf & " and SAT12MO1_ACN.CUST_CODE (+) = SAT12MO1.CUST_CODE"
                    sql_filter &= vbCrLf & " and SAT12MO1_ACN.CUST_STORE_NO (+) = SAT12MO1.CUST_STORE_NO"
                    sql_filter &= vbCrLf & " and NVL(SAT12MO1_ACN.MOS,0) " & IIf(optACN.Value = "C", "<>0", "=0")
                End If


                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "'" & sql_GRN & " THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from " & IIf(GRN = "T", "RSTRETL1", SOURCE_TABLE_NAME) & " SAT12MO1 " & sql_TABLE_NAMEs & vbCrLf _
                & IIf(optACN.Value <> "A" And IIf(GRN = "T", SAT12MO1_ACN_R, SAT12MO1_ACN_W) <> "", "," & IIf(GRN = "T", SAT12MO1_ACN_R, SAT12MO1_ACN_W) & " SAT12MO1_ACN", "") & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If

        Next


        ASCMAIN1.sql = "Select * from ICTITEM1 where ROWNUM < 1"
        Dim ICTITEM1_BUDGETS As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1_BUDGETS & " Add Primary Key (ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTITEM1_BUDGETS & "_1 on " & ICTITEM1_BUDGETS & " (COLLECTION_CODE, ITEM_CATGY_CODE)")

        '' Wholesale (Gross Shipment) Budgets

        'ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1)

        'ASCMAIN1.sql = "Insert into " & ICTITEM1 _
        '& " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE) " _
        '& " Select ROWNUM, 'Item for ' || COLLECTION_CODE || '-' || ITEM_CATGY_CODE" _
        '& ", COLLECTION_CODE, ITEM_CATGY_CODE " _
        '& " from " _
        '& " (Select Distinct COLLECTION_CODE, ITEM_CATGY_CODE from SATBUDG1)"
        'ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" _
        '& ", OPS_YYYYPP, AMT_SOLD BUDGET " _
        '& " from RSTRETL1 where ROWNUM < 1"
        'Dim SATBUDGI As String = ASCMAIN1.Temp_Table
        'For P As Integer = 1 To 12
        '    ASCMAIN1.sql = "Insert into " & SATBUDGI _
        '    & " Select SATBUDG1.CUST_CODE, '000000' CUST_STORE_NO, ICTITEM1.ITEM_CODE" _
        '    & ", OPS_YYYY || '" & Format(P, "00") & "'" _
        '    & ", SATBUDG1.BUDGET_P" & Format(P, "00") & " BUDGET" _
        '    & " from SATBUDG1," & ICTITEM1 & " ICTITEM1" _
        '    & " where SATBUDG1.BUDGET_P" & Format(P, "00") & " <> 0" _
        '    & "   and ICTITEM1.COLLECTION_CODE = SATBUDG1.COLLECTION_CODE" _
        '    & "   and ICTITEM1.ITEM_CATGY_CODE = SATBUDG1.ITEM_CATGY_CODE"
        '    ASCDATA1.ExecuteSQL()
        'Next

        'SOURCE_TABLE_NAME = SATBUDGI

        'For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

        '    Dim USC As String = rowASTRECAP.Item("USC")
        '    Dim GRN As String = rowASTRECAP.Item("GRN")
        '    Dim YEAR As String = rowASTRECAP.Item("YEAR")
        '    Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

        '    Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
        '    Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

        '    Dim YP As String = ""
        '    Dim COLUMN_NAME As String = ""
        '    Dim sql_GRN As String = ""

        '    If GRN = "G" And USC = "B" Then

        '        If GRN = "G" Then
        '            sql_GRN = ""
        '            YP = "SAT12MO1.OPS_YYYYPP"
        '            COLUMN_NAME = "NVL(SAT12MO1.BUDGET,0)"
        '        End If

        '        sql_filter = " and " & YP & " BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'"

        '        Dim sql_Data As String = ""
        '        For M As Integer = 1 To 12
        '            Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
        '            sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "'" & sql_GRN & " THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
        '        Next

        '        sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
        '        & " from " & SOURCE_TABLE_NAME & " SAT12MO1 " & sql_TABLE_NAMEs & vbCrLf _
        '        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        '        & " group by " & sql_GROUP_BY_cols

        '        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        '        & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
        '        & COLUMN_NAMEs_appended _
        '        & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
        '        & "(" & sql & ")"
        '        ASCDATA1.ExecuteSQL()

        '    End If
        'Next

        ' Wholesale Shipment Budgets

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1_BUDGETS)

        ASCMAIN1.sql = "Insert into " & ICTITEM1_BUDGETS _
        & " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE) " _
        & " Select ROWNUM, 'Item for ' || COLLECTION_CODE || '-' || ITEM_CATGY_CODE" _
        & ", COLLECTION_CODE, ITEM_CATGY_CODE " _
        & " from " _
        & " (Select Distinct COLLECTION_CODE, ITEM_CATGY_CODE from SATBUDD1)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" _
        & ", OPS_YYYYPP, AMT_SOLD BUDGET " _
        & " from RSTRETL1 where ROWNUM < 1"
        Dim SATBUDDI As String = ASCMAIN1.Temp_Table
        For P As Integer = 1 To 12
            ASCMAIN1.sql = "Insert into " & SATBUDDI _
            & " Select SATBUDD1.CUST_CODE, SATBUDD1.CUST_STORE_NO, ICTITEM1.ITEM_CODE" _
            & ", OPS_YYYY || '" & Format(P, "00") & "'" _
            & ", SATBUDD1.BUDGET_P" & Format(P, "00") & " BUDGET" _
            & " from SATBUDD1," & ICTITEM1_BUDGETS & " ICTITEM1" _
            & " where SATBUDD1.BUDGET_P" & Format(P, "00") & " <> 0" _
            & "   and ICTITEM1.COLLECTION_CODE = SATBUDD1.COLLECTION_CODE" _
            & "   and ICTITEM1.ITEM_CATGY_CODE = SATBUDD1.ITEM_CATGY_CODE"
            ASCDATA1.ExecuteSQL()
        Next

        Dim PA As Int32 = 1 - ASCMAIN1.PCO ' -1 * ((12 + ASCMAIN1.PCO) Mod 12) + 1

        ASCMAIN1.sql = "Update " & SATBUDDI & " Set OPS_YYYYPP = PERIOD_CALC(OPS_YYYYPP," & CStr(PA) & ")"
        ASCDATA1.ExecuteSQL()

        SOURCE_TABLE_NAME = SATBUDDI

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            Dim YP As String = ""
            Dim COLUMN_NAME As String = ""
            Dim sql_GRN As String = ""

            If (GRN = "G" Or GRN = "N") And USC = "B" Then

                If GRN = "G" Or GRN = "N" Then
                    sql_GRN = ""
                    YP = "SAT12MO1.OPS_YYYYPP"
                    COLUMN_NAME = "NVL(SAT12MO1.BUDGET,0)"
                End If

                sql_filter = " and " & YP & " BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'"

                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "'" & sql_GRN & " THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from " & SOURCE_TABLE_NAME & " SAT12MO1 " & Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1_BUDGETS & " ICTITEM1") & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If
        Next

        ' Retail Budgets

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1_BUDGETS)

        ASCMAIN1.sql = "Insert into " & ICTITEM1_BUDGETS _
        & " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE) " _
        & " Select ROWNUM, 'Item for ' || COLLECTION_CODE || '-' || ITEM_CATGY_CODE" _
        & ", COLLECTION_CODE, ITEM_CATGY_CODE " _
        & " from " _
        & " (Select Distinct COLLECTION_CODE, ITEM_CATGY_CODE from RSTBUDR1)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" _
        & ", OPS_YYYYPP, AMT_SOLD BUDGET " _
        & " from RSTRETL1 where ROWNUM < 1"
        Dim RSTBUDRI As String = ASCMAIN1.Temp_Table
        For P As Integer = 1 To 12
            ASCMAIN1.sql = "Insert into " & RSTBUDRI _
            & " Select RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, ICTITEM1.ITEM_CODE" _
            & ", OPS_YYYY || '" & Format(P, "00") & "'" _
            & ", RSTBUDR1.BUDGET_P" & Format(P, "00") & " BUDGET" _
            & " from RSTBUDR1," & ICTITEM1_BUDGETS & " ICTITEM1" _
            & " where RSTBUDR1.BUDGET_P" & Format(P, "00") & " <> 0" _
            & "   and ICTITEM1.COLLECTION_CODE = RSTBUDR1.COLLECTION_CODE" _
            & "   and ICTITEM1.ITEM_CATGY_CODE = RSTBUDR1.ITEM_CATGY_CODE"
            ASCDATA1.ExecuteSQL()
        Next

        'Dim PA As Int32 = 1 - ASCMAIN1.PCO ' -1 * ((12 + ASCMAIN1.PCO) Mod 12) + 1

        ASCMAIN1.sql = "Update " & RSTBUDRI & " Set OPS_YYYYPP = PERIOD_CALC(OPS_YYYYPP," & CStr(PA) & ")"
        ASCDATA1.ExecuteSQL()

        SOURCE_TABLE_NAME = RSTBUDRI

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            Dim YP As String = ""
            Dim COLUMN_NAME As String = ""
            Dim sql_GRN As String = ""

            If GRN = "T" And USC = "B" Then

                If GRN = "T" Then
                    sql_GRN = ""
                    YP = "SAT12MO1.OPS_YYYYPP"
                    COLUMN_NAME = "NVL(SAT12MO1.BUDGET,0)"
                End If

                sql_filter = " and " & YP & " BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'"

                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "'" & sql_GRN & " THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from " & SOURCE_TABLE_NAME & " SAT12MO1 " & Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1_BUDGETS & " ICTITEM1") & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If
        Next


        Dim YTDcols As String = ""
        Dim TOTALcols As String = ""
        For M As Integer = 1 To 12
            If M >= PTD1 And M <= PTD2 Then
                YTDcols &= "+M" & Format(M, "00")
            End If
            TOTALcols &= "+M" & Format(M, "00")
        Next
        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set YTD = " & YTDcols & ", TOTAL = " & TOTALcols
        ASCDATA1.ExecuteSQL()

    End Sub

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)

        If Absx1.chkFor("THOUSANDS").Checked Then
            Dim sql As String = ""
            For Each COLUMN_NAME As String In New String() _
            {"M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12", "YTD", "TOTAL"}
                sql &= ", " & COLUMN_NAME & " = " & COLUMN_NAME & " / 1000"
            Next

            ASCMAIN1.sql = "Update " & TT & " Set " & Mid(sql, 2)
            ASCDATA1.ExecuteSQL()
        End If

        For Each rowASTRECAP As DataRow In tblASTRECAP.Select("USC = 'P'")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim rowS() As DataRow = tblASTRECAP.Select("USC = 'S' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
            Dim rowG() As DataRow = tblASTRECAP.Select("USC = 'G' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
            'Dim rowC() As DataRow = tblASTRECAP.Select("USC = 'C' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
            ASCMAIN1.sql = "Delete from " & TT & " where ASTSRPT1_RECAP_ROW_NO = " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO")
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO " & TT & " Select S.G1,S.G2,S.G3,S.G4,S.G5,S.G6,S.G7,S.G8,S.G9," & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf
            For Each COLUMN_NAME As String In New String() _
            {"M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12", "YTD", "TOTAL"}
                ASCMAIN1.sql &= Replace(", TRUNC(10000*DECODE(NVL(S.M01,0),0,0,NVL(G.M01,0)/NVL(S.M01,0)))/100 M01" & vbCrLf, "M01", COLUMN_NAME)
            Next
            ASCMAIN1.sql &= ",null, null"
            ASCMAIN1.sql &= " FROM " & vbCrLf _
            & "(SELECT * FROM " & TT & " WHERE ASTSRPT1_RECAP_ROW_NO = " & rowS(0).Item("ASTSRPT1_RECAP_ROW_NO") & ") S," & vbCrLf _
            & "(SELECT * FROM " & TT & " WHERE ASTSRPT1_RECAP_ROW_NO = " & rowG(0).Item("ASTSRPT1_RECAP_ROW_NO") & ") G" & vbCrLf _
            & " WHERE S.G1 = G.G1 AND S.G2 = G.G2 AND S.G3 = G.G3 AND S.G4 = G.G4 AND S.G5 = G.G5" & vbCrLf _
            & "   AND S.G6 = G.G6 AND S.G7 = G.G7 AND S.G8 = G.G8 AND S.G9 = G.G9"

            ASCDATA1.ExecuteSQL()
        Next
    End Sub

    Overrides Sub Build_Report_File_Post_Process()


        Dim rowASTGROUP As DataRow = dst.Tables("ASTGROUP").Rows.Find(aRC)
        rowASTGROUP.Item("GROUP_CODE") = "Recap"
        rowASTGROUP.Item("GROUP_DESC") = "All Above"

        Dim P As New List(Of Int16)
        For Each ROWASTRECAP As DataRow In dst.Tables("ASTRECAP").Rows
            If ROWASTRECAP.Item("TYPE") & "" = "P" Then
                P.Add(ROWASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO"))
            End If
        Next

        Dim J As Int16 = 0
        Dim rowT() As DataRow = Nothing

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "G1,G2,G3,G4,G5,G6,G7,G8,G9")
            If P.Contains(Val(rowASTSRPT1.Item("ASTSRPT1_RECAP_ROW_NO"))) Then
                rowASTSRPT1.Item("PPTT") = DBNull.Value
                rowASTSRPT1.Item("TPTT") = DBNull.Value
            Else
                For I As Int16 = 1 To 9
                    If rowASTSRPT1.Item("G" & CStr(I)) & "" = "x" Or rowASTSRPT1.Item("G" & CStr(I)) & "" = aRC Then
                        J = I - 1
                        Exit For
                    End If
                Next
                Dim SQL As String = ""
                If J > 1 Then
                    For I As Int16 = 1 To J - 1
                        SQL &= " and G" & CStr(I) & " = '" & rowASTSRPT1.Item("G" & CStr(I)) & "'"
                    Next
                End If
                SQL = "G" & CStr(J) & " = '" & aRC & "' and ASTSRPT1_RECAP_ROW_NO = " & rowASTSRPT1.Item("ASTSRPT1_RECAP_ROW_NO") & SQL

                If J = 0 Then
                    rowASTSRPT1.Item("PPTT") = DBNull.Value
                    rowASTSRPT1.Item("TPTT") = DBNull.Value
                Else
                    rowT = dst.Tables("ASTSRPT1").Select(SQL)

                    If rowT IsNot Nothing AndAlso rowT.Length = 1 Then
                        Dim YTD As Decimal = Val(rowASTSRPT1.Item("YTD") & "")
                        Dim YTD_T As Decimal = Val(rowT(0).Item("YTD") & "")
                        Dim PPTT As Decimal = 0
                        If YTD_T <> 0 Then PPTT = 100 * YTD / YTD_T
                        rowASTSRPT1.Item("PPTT") = PPTT
                        Dim TOTAL As Decimal = Val(rowASTSRPT1.Item("TOTAL") & "")
                        Dim TOTAL_T As Decimal = Val(rowT(0).Item("TOTAL") & "")
                        Dim TPTT As Decimal = 0
                        If TOTAL_T <> 0 Then TPTT = 100 * TOTAL / TOTAL_T
                        rowASTSRPT1.Item("TPTT") = TPTT
                    End If
                End If

            End If
        Next
    End Sub

    Public Overrides Sub Print_Report()

        'Dim SUBT As String = ""
        If optACN.Value = "C" Then
            SUBT &= " -Incl Comp Stores Only"
        ElseIf optACN.Value = "N" Then
            SUBT &= " -Incl Non-Comp Stores Only"
        End If

        CR_params.Add("RYPLEGEND", RYPLEGEND)

        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        CR_params.Add("HIDE", "0")
        For M As Int32 = 1 To 12
            Dim MONTH_DESC As String = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -1 * (12 - M)), False, True), 1, 6)
            CR_params.Add("MD" & Format(M, "00"), MONTH_DESC)
            If M = PTD1 Then
                CR_params.Add("T1", Format(M, "00"))
                If Absx1.chkFor("FISCALYTD").Checked Then
                    CR_params.Add("MD1", "Fiscal YTD")
                Else
                    CR_params.Add("MD1", MONTH_DESC)
                End If
            End If
            If M = PTD2 Then
                CR_params.Add("T2", Format(M, "00"))
                If Absx1.chkFor("FISCALYTD").Checked Then
                    CR_params.Add("MD2", Mid(RYP, 1, 4))
                Else
                    CR_params.Add("MD2", MONTH_DESC)
                End If
            End If
        Next
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Report Period"
            End If

            For Each rowASTRECAP As DataRow In tblASTRECAP.Select("USC = 'P'")
                Dim GRN As String = rowASTRECAP.Item("GRN")
                Dim YEAR As String = rowASTRECAP.Item("YEAR")
                Dim rowS() As DataRow = tblASTRECAP.Select("USC = 'S' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
                Dim rowG() As DataRow = tblASTRECAP.Select("USC = 'G' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
                'Dim rowC() As DataRow = tblASTRECAP.Select("USC = 'C' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
                If rowS.Length <> 1 Or rowG.Length <> 1 Then ' Or rowC.Length <> 1 Then
                    EMsg &= "You Must Select 1 line for Sales and 1 line for GP in order to have GP% for Sales Type:" & GRN & ", Year: " & YEAR
                End If
            Next
        End If
    End Sub

    Private Sub utb1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles utb1.ValueChanged
        lblutb1.Text = Format(utb1.Value, "00")
    End Sub

    Private Sub utb2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles utb2.ValueChanged
        lblutb2.Text = Format(utb2.Value, "00")
    End Sub

    Private Sub UltraCheckEditor2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraCheckEditor2.CheckedChanged

        utb1.Enabled = Not Absx1.chkFor("FISCALYTD").Checked
        utb2.Enabled = Not Absx1.chkFor("FISCALYTD").Checked
        If Absx1.chkFor("FISCALYTD").Checked Then
            Set_utb()
        End If

    End Sub

    Sub Set_utb()
        Dim YP As String = Absx1.cmbFor("RYP").Value
        Dim P As Integer = Val(Mid(YP, 6, 2))
        utb1.Value = 12 - P + 1
        utb2.Value = 12
    End Sub

    Private Sub UltraCombo1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles UltraCombo1.InitializeLayout

    End Sub

    Private Sub UltraCombo1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraCombo1.ValueChanged
        If Absx1.chkFor("FISCALYTD").Checked Then
            Set_utb()
        End If
    End Sub

    Public Overrides Sub Post_Process_Special()
        MyBase.Post_Process_Special()
        Try
            'Prepare_XLS()

            Dim opt As String = optSL.Value & "" '  "L" ' S = Stacked, L = in Line

            Dim colors() As System.Drawing.Color = _
                {System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige, _
                 System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige, _
                 System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige}

            Dim colorsF() As System.Drawing.Color = _
                {System.Drawing.Color.ForestGreen, System.Drawing.Color.Purple, System.Drawing.Color.Orange, _
                 System.Drawing.Color.Blue, System.Drawing.Color.Brown, System.Drawing.Color.Red, _
                 System.Drawing.Color.LimeGreen, System.Drawing.Color.Turquoise, System.Drawing.Color.Salmon}

            Dim XC As Integer = 0
            Dim XR As Integer = 0

            GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

            Dim FILENAME As String = FORM_NAME & "_" & XNO & ".XLSX"
            Dim myWorkbook As New GemBox.Spreadsheet.ExcelFile
            Dim ws As GemBox.Spreadsheet.ExcelWorksheet = myWorkbook.Worksheets.Add(MENU_ITEM_OBJECT)


            Dim M_DESC(12) As String
            For M As Integer = 1 To 12
                Dim MONTH_DESC As String = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -1 * (12 - M)), False, True), 1, 6)
                M_DESC(M) = MONTH_DESC
            Next

            Dim R As Integer = -1
            ws.Cells(1, 0).Style.Font.Color = System.Drawing.Color.Blue
            ws.Cells(1, 0).Style.Font.Size = 300
            ws.Cells(1, 0).Style.Font.Name = "Times New Roman"
            ws.Cells(1, 0).Value = MENU_ITEM_DESC
            ws.Cells(0, 1).Value = MENU_ITEM_OBJECT
            ws.Cells(2, 0).Value = SUBT

            With ws.Cells(0, 0)
                .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Left
                .Style.NumberFormat = "mm/dd/yy;@"
                .Value = Now
            End With

            R = 2

            Dim LVLS As Integer = COLUMN_CAPTION_by_Lvl.Count - 1

            R += 2
            For C As Integer = 1 To 9
                ws.Columns(C - 1).Style.Font.Name = "Verdana"

                ws.Cells(R, C - 1).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
                ws.Cells(R, C - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

                If C <= LVLS Then
                    ws.Cells(R, C - 1).Value = COLUMN_CAPTION_by_Lvl(C)
                    ws.Cells(R, C - 1).Style.Font.Color = colorsF(C - 1)

                    ws.Columns(C - 1).Width = 4000
                Else
                    ws.Columns(C - 1).Hidden = True
                    ws.Columns(C - 1).Width = 0
                End If
            Next

            ws.Cells(R, 9).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
            ws.Cells(R, 9).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

            ws.Cells(R, 9).Value = "Description"
            ws.Columns(9).Width = 10000

            Dim MBAY As Integer = 15
            Dim GBAY As Integer = 11

            Dim Ms As New Dictionary(Of Integer, Integer)
            Dim DATA_TYPEs(dst.Tables("ASTRECAP").Rows.Count) As String

            Dim M_MAX As Integer = 0
            For Each row As DataRow In dst.Tables("ASTRECAP").Select("", "ASTSRPT1_RECAP_ROW_NO")
                M_MAX += 1

                Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(row.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                Ms.Add(ASTSRPT1_RECAP_ROW_NO, M_MAX)

                Dim M As Integer = M_MAX

                DATA_TYPEs(M) = row.Item("ASTSRPT1_RECAP_ROW_CAPTION") & ""

                If opt <> "S" Or M = 1 Then
                    If opt = "S" Then
                        ws.Cells(R, GBAY - 1).Value = "Data Type"
                        ws.Cells(R, GBAY - 1).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
                        ws.Cells(R, GBAY - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

                        ws.Columns(GBAY - 1).Width = 4000
                    Else
                        ws.Cells(R - 1, GBAY + MBAY * (M - 1)).Value = row.Item("ASTSRPT1_RECAP_ROW_CAPTION")
                        ws.Columns(GBAY + MBAY * (M - 1) + 0 - 1).Width = 100
                    End If
                    For C As Integer = 1 To 12
                        XC = GBAY + MBAY * (M - 1) + C - 1
                        With ws.Columns(XC)
                            .Style.Font.Name = "Verdana"
                            .Style.NumberFormat = "#,##0;@"
                            .Style.Font.Name = "Verdana"
                            .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                            .Width = 3000
                        End With
                        With ws.Cells(R, XC)
                            .Value = M_DESC(C)
                            .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                            .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                        End With
                    Next

                    XC = GBAY + MBAY * (M - 1) + 13 - 1
                    With ws.Columns(XC)
                        .Style.Font.Name = "Verdana"
                        .Style.NumberFormat = "#,##0;@"
                        .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                        .Width = 3000
                    End With
                    With ws.Cells(R, XC)
                        .Value = "YTD"
                        .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                        .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                    End With

                    XC = GBAY + MBAY * (M - 1) + 14 - 1
                    With ws.Columns(XC)
                        .Style.Font.Name = "Verdana"
                        .Style.NumberFormat = "#,##0;@"
                        .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                        .Width = 3000
                    End With
                    With ws.Cells(R, XC)
                        .Value = "Total"
                        .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                        .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                    End With

                End If
            Next

            Dim LAST_LVL As Integer = 0
            Dim LAST_KEY = ""
            Dim START_ROW As Integer = -1
            Dim LEVEL_LAST_ROW As Integer = 0
            Dim C_LVL As Integer = 0

            Dim START_ROWS(LVLS) As List(Of Integer)
            For L As Integer = 0 To LVLS
                START_ROWS(L) = New List(Of Integer)
            Next

            Dim MROW As Integer = 0
            Dim RSTART As Integer = R
            Dim CODEs(LVLS) As String
            Dim First_Row_at_Level As Boolean = False

            For Each row As DataRow In dst.Tables("ASTSRPT1").Select("", "G1,G2,G3,G4,G5,G6,G7,G8,G9,ASTSRPT1_RECAP_ROW_NO")
                Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(row.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                Dim M As Integer = Ms(ASTSRPT1_RECAP_ROW_NO)
                Dim THIS_KEY = ""
                For I As Integer = 1 To 9
                    THIS_KEY &= vbTab & row.Item("G" & CStr(I))
                Next

                If LAST_KEY <> THIS_KEY Then
                    LAST_KEY = THIS_KEY
                    LEVEL_LAST_ROW = C_LVL
                    If opt = "S" Then
                        MROW += 1
                    Else
                        R += 1
                    End If
                    ReDim CODEs(LVLS)
                    First_Row_at_Level = True
                Else
                    First_Row_at_Level = False
                End If

                If opt = "S" Then
                    R = RSTART + (MROW - 1) * M_MAX + M
                End If

                Dim CODE_VALUE As String = ""
                C_LVL = 0
                For C As Integer = 1 To LVLS
                    Dim Z As String = row.Item("G" & CStr(C)) & ""
                    If InStr(Z, ":") = 0 Then
                        Exit For
                    End If
                    CODE_VALUE = Split(Z & ":", ":")(1)
                    C_LVL = C
                    If First_Row_at_Level Then
                        CODEs(C) = CODE_VALUE
                        ws.Cells(R, C - 1).Value = CODE_VALUE
                        ' this could be done at end for entire range of cells
                        ws.Cells(R, C - 1).Style.Font.Color = colorsF(C - 1) ' see end of loop
                    End If
                Next

                If LEVEL_LAST_ROW <> C_LVL Then
                    If C_LVL = LVLS Then
                        If First_Row_at_Level Then
                            'If LEVEL_LAST_ROW = C_LVL Then
                            '    Stop
                            'End If
                            START_ROW = R
                            START_ROWS(C_LVL).Clear()
                        End If
                    End If
                End If

                If opt = "S" Then
                    If First_Row_at_Level Then
                        For MM As Integer = 1 To M_MAX
                            ws.Cells(RSTART + (MROW - 1) * M_MAX + MM, GBAY - 1).Value = DATA_TYPEs(MM)
                            If C_LVL <> 0 Then ws.Cells(RSTART + (MROW - 1) * M_MAX + MM, GBAY - 1).Style.Font.Color = colorsF(C_LVL - 1)
                        Next
                        If C_LVL = LVLS Then START_ROWS(C_LVL).Add(R)
                    End If
                End If

                Dim rowASTGROUP As DataRow
                Dim DESC_VALUE As String
                If First_Row_at_Level Then
                    rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(COLUMN_CAPTION_by_Lvl(C_LVL) & ":" & CODE_VALUE)
                    DESC_VALUE = ""
                    If rowASTGROUP IsNot Nothing Then
                        DESC_VALUE = rowASTGROUP.Item("GROUP_DESC")
                        ws.Cells(R, 9).Value = DESC_VALUE
                        ws.Cells(R, 9).Style.Font.Color = colorsF(C_LVL - 1)
                        ws.Cells(R, 9).Style.Indent = (C_LVL - 1) * 1
                    End If
                End If

                Dim MC As Integer = IIf(opt = "S", 1, M)

                If C_LVL = LVLS Then
                    For C As Integer = 1 To 12
                        XC = GBAY + MBAY * (MC - 1) + C - 1
                        ws.Cells(R, XC).Value = row.Item("M" & Format(C, "00"))
                    Next

                    XC += 1 : ws.Cells(R, XC).Value = row.Item("YTD")
                    XC += 1 : ws.Cells(R, XC).Value = row.Item("TOTAL")

                Else
                    If C_LVL = LVLS - 1 And opt <> "S" Then
                        For C As Integer = 1 To 12
                            XC = GBAY + MBAY * (MC - 1) + C - 1
                            ws.Cells(R, XC).Formula = "=SUM(" & ASCMAIN1.Excel_Cell(START_ROW + 1, XC + 1) & ":" & ASCMAIN1.Excel_Cell(R - 1 + 1, XC + 1) & ")"
                        Next

                        XC += 1 : ws.Cells(R, XC).Formula = "=SUM(" & ASCMAIN1.Excel_Cell(START_ROW + 1, XC + 1) & ":" & ASCMAIN1.Excel_Cell(R - 1 + 1, XC + 1) & ")"
                        XC += 1 : ws.Cells(R, XC).Formula = "=SUM(" & ASCMAIN1.Excel_Cell(START_ROW + 1, XC + 1) & ":" & ASCMAIN1.Excel_Cell(R - 1 + 1, XC + 1) & ")"

                        If Not START_ROWS(C_LVL).Contains(R) Then
                            For rr As Integer = START_ROW To R - 1
                                ws.Rows(rr).OutlineLevel = LVLS
                            Next

                            START_ROWS(C_LVL).Add(R)
                        End If


                    Else
                        Dim TT As String = ""
                        For Each RR As Integer In START_ROWS(C_LVL + 1)
                            TT &= "," & ASCMAIN1.Excel_Cell(RR + IIf(opt = "S", M, 1), 1)
                            If First_Row_at_Level Then
                                For RRM As Integer = RR To RR + (M_MAX - 1)
                                    ws.Rows(RRM).OutlineLevel = C_LVL + 1
                                Next
                            End If
                        Next


                        Dim TTC As String = ""
                        For C As Integer = 1 To 12
                            XC = GBAY + MBAY * (MC - 1) + C - 1
                            TTC = ASCMAIN1.Excel_Cell(0, XC + 1)
                            ws.Cells(R, XC).Formula = "=SUM(" & Mid(Replace(TT, ",A", "," & TTC), 2) & ")"
                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)
                        Next

                        XC += 1 : TTC = ASCMAIN1.Excel_Cell(0, XC + 1) : ws.Cells(R, XC).Formula = "=SUM(" & Mid(Replace(TT, ",A", "," & TTC), 2) & ")" : If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)
                        XC += 1 : TTC = ASCMAIN1.Excel_Cell(0, XC + 1) : ws.Cells(R, XC).Formula = "=SUM(" & Mid(Replace(TT, ",A", "," & TTC), 2) & ")" : If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)

                        If First_Row_at_Level Then
                            'If Not START_ROWS(C_LVL).Contains(R) Then
                            START_ROWS(C_LVL).Add(R)
                            If C_LVL = 0 Then ws.Cells(R, 0).Value = "Totals"
                            'End If
                        End If

                    End If
                End If
            Next

            Dim RFINAL As Integer = R
            If opt = "S" Then
                RFINAL = RSTART + MROW * M_MAX - 1
            End If

            'For C As Integer = 1 To LVLS
            '    Dim cr As GemBox.Spreadsheet.CellRange = ws.Cells.GetSubrange(ASCMAIN1.Excel_Cell(RSTART, C), ASCMAIN1.Excel_Cell(RFINAL, C))
            '    cr.Style.Font.Color = colorsF(C - 1)
            'Next

            Gembox_Export_to_Excel_Show(myWorkbook, FILENAME)


        Catch ex As Exception
            If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
        End Try
    End Sub


    Public Overrides Function Prepare_XLS_Summary_Columns(ByVal COLUMN_NAME_sum As Dictionary(Of String, String)) As String

        SUBT = ASCMAIN1.Get_Legend_Wk(RYW)

        'For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Rows
        '    Dim ITEM_CODE As String = rowASTSRPT1.Item("ITEM_CODE") & ""
        '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", "")
        '    If rowICTITEM1 IsNot Nothing Then
        '        rowASTSRPT1.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        '        rowASTSRPT1.Item("LAUNCH_DATE") = rowICTITEM1.Item("LAUNCH_DATE")
        '    End If

        'Next
        '        Return "QTY_EOW,TY_WTD_S,LY_WTD_S,TY_MTD_S,LY_MTD_S,LY_MTL_S,TY_STD_S,LY_STD_S,LY_STL_S,TY_YTD_S,LY_YTD_S,LY_YTL_S"

        'If Not dst.Tables("ASTSRPT1").Columns.Contains("LAUNCH_DATE") Then
        '    With dst.Tables("ASTSRPT1")
        '        .Columns.Add("LAUNCH_DATE")
        '        .Columns.Add("WST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_WTD_S)=0,0,100*TY_WTD_S/(QTY_EOW+TY_WTD_S))")
        '        .Columns.Add("WWOH", GetType(System.Decimal), "IIF(TY_WTD_S=0,0,QTY_EOW/TY_WTD_S)")
        '        .Columns.Add("MTD_PCT", GetType(System.Decimal), "IIF(LY_MTD_S=0,0,100*(TY_MTD_S-LY_MTD_S)/LY_MTD_S)")
        '        .Columns.Add("MST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_MTD_S)=0,0,100*TY_MTD_S/(QTY_EOW+TY_MTD_S))")
        '        .Columns.Add("STD_PCT", GetType(System.Decimal), "IIF(LY_STD_S=0,0,100*(TY_STD_S-LY_STD_S)/LY_STD_S)")
        '        .Columns.Add("STL_PCT", GetType(System.Decimal), "IIF(LY_STL_S=0,0,100*TY_STD_S/LY_STL_S)")
        '        .Columns.Add("SST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_STD_S)=0,0,100*TY_STD_S/(QTY_EOW+TY_STD_S))")
        '        .Columns.Add("YTD_PCT", GetType(System.Decimal), "IIF(LY_YTD_S=0,0,100*(TY_YTD_S-LY_YTD_S)/LY_YTD_S)")
        '        .Columns.Add("YTL_PCT", GetType(System.Decimal), "IIF(LY_YTL_S=0,0,100*TY_YTD_S/LY_YTL_S)")
        '        .Columns.Add("YST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_YTD_S)=0,0,100*TY_YTD_S/(QTY_EOW+TY_YTD_S))")
        '    End With
        'End If
        Return "M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,YTD,TOTAL"

    End Function

    Overrides Sub Prepare_XLS_Prepare_row(ByVal row As DataRow)
        'Dim GMAX As Integer = COLUMN_NAMEs.Count

        'If COLUMN_NAMEs(GMAX - 1) <> "ITEM_CODE" Then
        '    Exit Sub
        'End If

        'Dim ITEM_CODE As String = row.Item("ITEM_CODE")
        'Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE) ' LookUp("ICTITEM1", ITEM_CODE)
        'row.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        'row.Item("LAUNCH_DATE") = rowICTITEM1.Item("LAUNCH_DATE")
    End Sub


End Class