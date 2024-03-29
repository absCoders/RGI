Public Class ICRPHYV1

    Dim LYP As String = ""

    Dim ICTPHYV1 As String = ""
    Dim WHTPHYV1 As String = ""

    Dim WHTCYCLX As String = ""
    Dim WHTCYCLC As String = ""

    Dim WHSE_CODEs As String = ""
    Dim current_period_physical As Boolean = False
    Dim RGI_FILEDS As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("ICTPARM1")

        If ASCMAIN1.CLIENT = "RGI" Then
            current_period_physical = True
            RGI_FILEDS = " and NVL(ICTPHYC2.STATUS,'A') = 'A' "
        End If

        grpStock.Visible = (ASCMAIN1.CLIENT = "VAN")
        chkL.Visible = (ASCMAIN1.CLIENT = "VAN")
        grpUnitVarianceRange.Visible = (ASCMAIN1.CLIENT = "VAN")

        chkShowImpactedItemsOnly.Visible = (ASCMAIN1.CLIENT = "VAN")

        chkUpdateVariances.Visible = Not (ASCMAIN1.CLIENT = "VAN")
        lblFutureDays.Visible = Not (ASCMAIN1.CLIENT = "VAN")

        LYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables
        WHSE_CODEs = SQLA("WHSE_CODE")

        If chkUpdateVariances.Checked Then
            RWU = "R"
        Else
            RWU = "N"
        End If
        'RWU = "R" - TOO DANGEROUS - DO THIS IN A SCREEN

        Prepare_Work_File()

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""
        'sql_filter &= SQL_in("SALES_DIVISION_CODE", "ICTSTYL1.SALES_DIVISION_CODE")

        If numVARU.Value <> 0 Then
            RWU = "N"
            ASCDATA1.ExecuteSQL("Delete from " & ICTPHYV1 & " where ABS(NVL(BOOK,0) - NVL(PHYS,0)) < " & CStr(numVARU.Value))
        End If
        If numVARC.Value <> 0 Then
            RWU = "N"
            ASCDATA1.ExecuteSQL("Delete from " & ICTPHYV1 & " where ABS(NVL(STYLE_COST,0) * (NVL(BOOK,0) - NVL(PHYS,0))) < " & CStr(numVARC.Value))
        End If

        If chkUnitVarianceRange.Checked Then
            RWU = "N"
            ASCDATA1.ExecuteSQL($"Delete from {ICTPHYV1} where (ABS(NVL(BOOK,0) - NVL(PHYS,0)) < {CStr(numVARMIN.Value)} or ABS(NVL(BOOK,0) - NVL(PHYS,0)) > {CStr(numVARMAX.Value)})")
        End If

        If optSORT.Value = "I" Then
            ASCDATA1.ExecuteSQL("Update " & ICTPHYV1 & " Set SORT_VALUE = STYLE_CODE || '-' || COLOR_CODE")
        ElseIf optSORT.Value = "U" Then
            ASCDATA1.ExecuteSQL("Update " & ICTPHYV1 & " Set SORT_VALUE = TRIM(TO_CHAR(9999999999 - ABS(NVL(PHYS,0) - NVL(BOOK,0)),'0000000000'))")
        ElseIf optSORT.Value = "C" Then
            ASCDATA1.ExecuteSQL("Update " & ICTPHYV1 & " Set SORT_VALUE = TRIM(TO_CHAR(9999999999 - NVL(STYLE_COST,0) * ABS(NVL(PHYS,0) - NVL(BOOK,0)),'0000000000'))")
        End If

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & ICTPHYV1, "ICTPHYV1", 3))
        With dst.Tables("ICTPHYV1")
            .Columns.Add("VAR", GetType(System.Int64), "ISNULL(PHYS,0) - ISNULL(BOOK,0)")
            .Columns.Add("PHYS_AMT", GetType(System.Int64), "ISNULL(PHYS,0) * ISNULL(STYLE_COST,0)")
            .Columns.Add("BOOK_AMT", GetType(System.Int64), "ISNULL(BOOK,0) * ISNULL(STYLE_COST,0)")
            .Columns.Add("VAR_AMT", GetType(System.Int64), "ISNULL(VAR,0) * ISNULL(STYLE_COST,0)")
        End With

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & WHTPHYV1, "WHTPHYV1", 4))

        dst.Tables.Add(ASCDATA1.GetDataTable("Select ICTSTYL1.* from ICTSTYL1 where ICTSTYL1.STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTPHYV1 & ")", "ICTSTYL1", 1))

        If ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.sql = "Select ICTSTYC1.*, ICTSTYC1.STYLE_COST_FIFO STYLE_COST from ICTSTYC1" & vbCrLf _
            & " where (ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE) in (Select Distinct STYLE_CODE, COLOR_CODE from " & ICTPHYV1 & ")"
        Else
            ASCMAIN1.sql = "Select ICTSTYC1.*, ICTCOSTA.STYLE_COST, ICTCOSTA.DATE_LAST_SHP, ICTCOSTA.DATE_LAST_REC from ICTSTYC1,ICTCOSTA" & vbCrLf _
            & " where (ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE) in (Select Distinct STYLE_CODE, COLOR_CODE from " & ICTPHYV1 & ")" & vbCrLf _
            & "   and ICTCOSTA.STYLE_CODE (+) = ICTSTYC1.STYLE_CODE" & vbCrLf _
            & "   and ICTCOSTA.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP (+) = '" & LYP & "' "
        End If
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYC1", 2))

        If ASCMAIN1.CLIENT = "RGI" Then
            'need to update style_cost on dataset for ICTSTYL1
            'the where clause avoids selecting errored entries made in lower case
            ASCMAIN1.sql = "SELECT * FROM ICTSTYV1 where STYLE_CODE < 'a'"
            For Each row As DataRow In ASCDATA1.GetDataTable().Select("PO_COST > 0")
                For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("STYLE_CODE = '" & row.Item("STYLE_CODE") & "' and VEND_CODE = '" & row.Item("VEND_CODE") & "'")
                    If Val(rowICTSTYL1.Item("STYLE_COST") & "") = 0 Then
                        If Not IsDBNull(row.Item("NEW_PO_COST_DATE")) AndAlso row.Item("NEW_PO_COST_DATE") < DATETIME_STAMP Then
                            rowICTSTYL1.Item("STYLE_COST") = row.Item("NEW_PO_COST")
                        Else
                            rowICTSTYL1.Item("STYLE_COST") = row.Item("PO_COST")
                        End If
                    End If
                Next

                For Each rowICTPHYV1 As DataRow In dst.Tables("ICTPHYV1").Select("STYLE_CODE = '" & row.Item("STYLE_CODE") & "'")
                    If Val(rowICTPHYV1.Item("STYLE_COST") & "") = 0 Then
                        If Not IsDBNull(row.Item("NEW_PO_COST_DATE")) AndAlso row.Item("NEW_PO_COST_DATE") < DATETIME_STAMP Then
                            rowICTPHYV1.Item("STYLE_COST") = row.Item("NEW_PO_COST")
                        Else
                            rowICTPHYV1.Item("STYLE_COST") = row.Item("PO_COST")
                        End If
                    End If
                Next

                For Each rowWHTPHYV1 As DataRow In dst.Tables("WHTPHYV1").Select("STYLE_CODE = '" & row.Item("STYLE_CODE") & "'")
                    If Val(rowWHTPHYV1.Item("STYLE_COST") & "") = 0 Then
                        If Not IsDBNull(row.Item("NEW_PO_COST_DATE")) AndAlso row.Item("NEW_PO_COST_DATE") < DATETIME_STAMP Then
                            rowWHTPHYV1.Item("STYLE_COST") = row.Item("NEW_PO_COST")
                        Else
                            rowWHTPHYV1.Item("STYLE_COST") = row.Item("PO_COST")
                        End If
                    End If
                Next
            Next
        End If


        If ASCMAIN1.CLIENT = "VAN" Then
            dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & WHTCYCLX, "WHTCYCLX", 3))
            dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & WHTCYCLC, "WHTCYCLC", 3))
        End If



        ' Extracts from Data Sources

        MyBase.Get_SQL("*", ICTPHYV1)

        Dim SOURCE_TABLE_NAME As String = "ICTPHYVX"
        ' Dim x As String = ASTSRPT1_sum_columns
        ' Dim y As String = ASTSRPT1_sql_sum
        Dim sql_Data As String = ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", " & SOURCE_TABLE_NAME & ".SORT_VALUE, " & SOURCE_TABLE_NAME & ".STYLE_CODE, " & SOURCE_TABLE_NAME & ".COLOR_CODE" & vbCrLf _
        & ASTSRPT1_sum_columns _
        & " from " & ICTPHYV1 & " " & SOURCE_TABLE_NAME & " " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols _
        & ", " & SOURCE_TABLE_NAME & ".SORT_VALUE, " & SOURCE_TABLE_NAME & ".STYLE_CODE, " & SOURCE_TABLE_NAME & ".COLOR_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        If ASCMAIN1.CLIENT = "VAN" Then
            RPT = "ICRPHYV2"
        End If

        If numVARU.Value <> 0 Then
            Page0.Add("Unit Variance Threshold: " & CStr(numVARU.Value))
        End If
        If numVARC.Value <> 0 Then
            Page0.Add("Cost Variance Threshold: " & CStr(numVARC.Value))
        End If
        If chkUnitVarianceRange.Checked Then
            Page0.Add($"Unit Variances between {CStr(numVARMIN.Value)} and {CStr(numVARMAX.Value)}")
        End If
        Select Case optSORT.Value
            Case "I"
                Page0.Add("Sorted by Item")
            Case "U"
                Page0.Add("Ranked by Unit Variance")
            Case "C"
                Page0.Add("Ranked by Cost Variance")
        End Select

        If ASCMAIN1.CLIENT = "VAN" Then
            If optASN.Value = "S" Then
                Page0.Add("Stock Only")
                SUBT = "Stock Only"
            ElseIf optASN.Value = "N" Then
                Page0.Add("Non-Stock Only")
                SUBT = "Non-Stock Only"
            End If

            If chkUnitVarianceRange.Checked Then
                If SUBT <> "" Then SUBT &= ", "
                SUBT &= $"Unit Variances between {CStr(numVARMIN.Value)} and {CStr(numVARMAX.Value)}"
            End If


            CR_params.Add("CHKL", IIf(chkL.Checked, "1", "0"))

        End If

        CR_params.Add("OPTR", optSORT.Value)
        Generate_Report(RPT, , SUBT)
    End Sub

    Sub Prepare_Work_File()

        Dim sqlASN As String = ""

        If ASCMAIN1.CLIENT = "VAN" Then

            If optASN.Value = "A" Then
            ElseIf optASN.Value = "S" Then
                sqlASN = " and ICTSTYL1.CUST_CODE is Null"
            ElseIf optASN.Value = "N" Then
                sqlASN = " and ICTSTYL1.CUST_CODE is NOT Null"
            End If

            ASCMAIN1.sql = "TRUNCATE TABLE ICTPHYC1"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO ICTPHYC1 
            SELECT WHSE_CODE, TICKET_NO, NULL, LOCATION_CODE, INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE, 'A', NULL,NULL
            FROM WHTPHYC1 WHERE TICKET_STATUS = 'A'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "TRUNCATE TABLE ICTPHYC2"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO ICTPHYC2
            SELECT WHSE_CODE, TICKET_NO, ROWNUM, STYLE_CODE, COLOR_CODE, 0, 0, PHYS_UNITS, BAR_CODE, 'A', NULL
            FROM WHTPHYC3 WHERE (WHSE_CODE, TICKET_NO) IN (SELECT WHSE_CODE, TICKET_NO FROM WHTPHYC1 WHERE TICKET_STATUS = 'A')"
            If chkShowImpactedItemsOnly.Checked Then
                ASCMAIN1.sql &= " and (STYLE_CODE, COLOR_CODE) in 
                    (Select Distinct STYLE_CODE, COLOR_CODE from WJZ_DUPLPN_429)"
            End If
            ASCDATA1.ExecuteSQL()

            If chkIncludeCycleCountAdjustments.Checked Then
                Dim cycle_counts As String = "SELECT Y.*, WHTCYCL2.BAR_CODE, WHTCYCL2.CYCLE_SCAN, WHTCYCL2.CYCLE_NEW, WHTLOCB2.WHSE_TRAN_QTY, WHTLOCB2.STYLE_CODE, WHTLOCB2.COLOR_CODE
                 FROM WHTCYCL2, WHTLOCB2, (
                SELECT X.*, WHTCYCL1.WHSE_CODE, WHTCYCL1.LOCATION_CODE, WHTCYCL1.CYCLE_STATUS, WHTCYCL1.CYCLE_RESOLUTION, WHTCYCL1.CYCLE_TYPE,'M' WHSE_TRAN_TYPE, WHTCYCL1.WHSE_TRAN_NO,
                 WHTCYCL1.CASES_BOOK, WHTCYCL1.CASES_PHYS FROM WHTCYCL1,(
                SELECT CYCLE_NO
                , SUM (CASE WHEN CYCLE_SCAN IS NULL THEN 1 ELSE 0 END) NOSCANS
                , SUM (CASE WHEN CYCLE_SCAN = '1' THEN 1 ELSE 0 END) SCANS
                , SUM (CASE WHEN CYCLE_NEW = '1' THEN 1 ELSE 0 END) NEW
                 FROM WHTCYCL2
                GROUP BY CYCLE_NO
                ) X WHERE WHTCYCL1.CYCLE_NO = X.CYCLE_NO
                AND WHTCYCL1.CYCLE_STATUS = 'D' AND WHTCYCL1.CYCLE_RESOLUTION = 'U' AND WHTCYCL1.CYCLE_TYPE = 'V'
                AND WHTCYCL1.INIT_DATE > '06-FEB-2023'
                ) Y  WHERE Y.CYCLE_NO = WHTCYCL2.CYCLE_NO
                AND (WHTCYCL2.CYCLE_SCAN IS NULL OR WHTCYCL2.CYCLE_NEW = '1')
                AND WHTLOCB2.WHSE_CODE = Y.WHSE_CODE AND WHTLOCB2.WHSE_TRAN_TYPE = Y.WHSE_TRAN_TYPE AND WHTLOCB2.WHSE_TRAN_NO = Y.WHSE_TRAN_NO
                AND WHTLOCB2.LOCATION_CODE= '00-LNF-A' AND (WHTLOCB2.BAR_CODE = WHTCYCL2.BAR_CODE OR WHTLOCB2.BAR_CODE_OTHER = WHTCYCL2.BAR_CODE)
                ORDER BY Y.CYCLE_NO"

                ASCMAIN1.sql = $"INSERT INTO ICTPHYC1 
                SELECT 'NJC' WHSE_CODE, 'C' || SUBSTR(CYCLE_NO,6,5), 'cycle', LOCATION_CODE, INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE, 'A', NULL, NULL
                FROM WHTCYCL1 WHERE CYCLE_NO IN (SELECT DISTINCT CYCLE_NO FROM ({cycle_counts}))"
                ASCDATA1.ExecuteSQL()

                Dim sqlCYCLE_ADJ As String = $"SELECT Z.WHSE_CODE, 'C' || SUBSTR(Z.CYCLE_NO,6,5), ROWNUM TICKET_LNO
                , Z.STYLE_CODE, Z.COLOR_CODE, 0 COUNT_CTNS, 0 CARTON_PACK_QTY, -1 * Z.WHSE_TRAN_QTY COUNT_LOOSE
                , Z.BAR_CODE, 'A' STATUS, NULL NEW_TICKET
                FROM ({cycle_counts}) Z"

                ASCMAIN1.sql = $"INSERT INTO ICTPHYC2 {sqlCYCLE_ADJ}"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = $"Select WHSE_CODE, STYLE_CODE, COLOR_CODE, SUM (COUNT_LOOSE) CYCLE_ADJ from ({sqlCYCLE_ADJ}) group by WHSE_CODE, STYLE_CODE, COLOR_CODE"
                WHTCYCLX = ASCMAIN1.Temp_Table
            Else
                ASCMAIN1.sql = $"Select WHSE_CODE, STYLE_CODE, COLOR_CODE, 0 CYCLE_ADJ from ICTSTAT2 where ROWNUM < 1"
                WHTCYCLX = ASCMAIN1.Temp_Table
            End If

            ASCMAIN1.sql = $"Select * from WHTCYCLC"
            WHTCYCLC = ASCMAIN1.Temp_Table


        End If

        Dim SQLW As String = ""
        If WHSE_CODEs <> "" Then
            SQLW = " and X.WHSE_CODE in ('" & Replace(WHSE_CODEs, ",", "','") & "')"
        End If
        If ASCMAIN1.CLIENT = "VAN" Then
            'SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_PHYS_STATUS = 'C')"
            SQLW &= " and X.WHSE_CODE = 'NJC'"
        Else
            SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_PHYS_STATUS = 'C')"
        End If

        'Code below replaces standard select for RGI for booked inventory from ictstat2 
        'With booked inventory In WHTLOCB0 because RGI continues to ship e-comm during inventory.
        'inventory falls around easter a big e-comm event.
        'Prior to inventory verify that ICTSTAT2 matches WHTLOCB1 inventory.
        If ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.sql = "Select X.*, ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
            & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & ", ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.ROYALTY_CODE" & vbCrLf _
            & " from ICTSTYL1, ICTSTYC1, (" & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE, WHSE_CODE, Sum (PHYS) PHYS, Sum (BOOK) BOOK from (" & vbCrLf _
            & "Select ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE" & vbCrLf _
            & ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS, 0 BOOK" _
            & " from ICTPHYC1,ICTPHYC2 where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" _
            & RGI_FILEDS & vbCrLf _
            & Replace(SQLW, "X.WHSE_CODE", "ICTPHYC2.WHSE_CODE") _
            & " group by ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & IIf(current_period_physical,
                "Select WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE" & vbCrLf _
                & ", 0 PHYS,  Sum (NVL(WHTLOCB0.LOCATION_QTY,0)) - Sum (nvl(WHTLOCB0.BOOK_INVTY_ADJ, 0)) BOOK" _
                & " from WHTLOCB0,ICTSTYL1 where ICTSTYL1.STYLE_CODE = WHTLOCB0.STYLE_CODE" _
                & Replace(SQLW, "X.WHSE_CODE", "WHTLOCB0.WHSE_CODE") _
                & " group by WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE" _
                & " having Sum (NVL(WHTLOCB0.LOCATION_QTY,0)) - Sum (nvl(WHTLOCB0.BOOK_INVTY_ADJ, 0)) <> 0",
                "Select ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE, ICTSTAT1.WHSE_CODE" & vbCrLf _
                & ", 0 PHYS, Sum (NVL(ICTSTAT1.WHSE_QTY_BEG,0)) BOOK" _
                & " from ICTSTAT1 where NVL(ICTSTAT1.WHSE_QTY_BEG,0) <> 0" _
                & Replace(SQLW, "X.WHSE_CODE", "ICTSTAT1.WHSE_CODE") _
                & " and ICTSTAT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                & " group by ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE, ICTSTAT1.WHSE_CODE") & vbCrLf _
            & ") group by STYLE_CODE, COLOR_CODE, WHSE_CODE) X" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
            & " and ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
            & " and ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE"
        Else
            ASCMAIN1.sql = "Select X.*, ICTCOSTA.STYLE_COST" & vbCrLf _
           & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
           & ", ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.ROYALTY_CODE, ICTSTYL1.CUST_CODE" & vbCrLf _
           & " from ICTSTYL1, ICTCOSTA, (" & vbCrLf _
           & "Select STYLE_CODE, COLOR_CODE, WHSE_CODE, Sum (PHYS) PHYS, Sum (BOOK) BOOK from (" & vbCrLf _
           & "Select ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE" & vbCrLf _
           & ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS, 0 BOOK" _
           & " from ICTPHYC1,ICTPHYC2 where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" _
           & Replace(SQLW, "X.WHSE_CODE", "ICTPHYC2.WHSE_CODE") _
           & " group by ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE" & vbCrLf _
           & " union " & vbCrLf _
           & IIf(current_period_physical,
               "Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_CODE" & vbCrLf _
               & ", 0 PHYS, Sum (NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)) BOOK" _
               & " from ICTSTAT2 where NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) <> 0" _
               & Replace(SQLW, "X.WHSE_CODE", "ICTSTAT2.WHSE_CODE") _
               & " group by ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_CODE",
                "Select WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE" & vbCrLf _
                & ", 0 PHYS,  Sum (NVL(WHTLOCB0.LOCATION_QTY,0)) - Sum (nvl(WHTLOCB0.BOOK_INVTY_ADJ, 0)) BOOK" _
                & " from WHTLOCB0,ICTSTYL1 where ICTSTYL1.STYLE_CODE = WHTLOCB0.STYLE_CODE" _
                & Replace(SQLW, "X.WHSE_CODE", "WHTLOCB0.WHSE_CODE") _
                & " group by WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE" _
                & " having Sum (NVL(WHTLOCB0.LOCATION_QTY,0)) - Sum (nvl(WHTLOCB0.BOOK_INVTY_ADJ, 0)) <> 0" & vbCrLf
               ) & vbCrLf _
           & ") group by STYLE_CODE, COLOR_CODE, WHSE_CODE) X" & vbCrLf _
           & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
           & sqlASN _
           & " and ICTCOSTA.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
           & " and ICTCOSTA.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
           & " and ICTCOSTA.OPS_YYYYPP (+) = '" & LYP & "'"
        End If

        If chkShowImpactedItemsOnly.Checked Then
            ASCMAIN1.sql &= " and (X.STYLE_CODE, X.COLOR_CODE) in 
                    (Select Distinct STYLE_CODE, COLOR_CODE from WJZ_DUPLPN_429)"
        End If

        '"Select ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE, ICTSTAT1.WHSE_CODE" & vbCrLf _
        '       & ", 0 PHYS, Sum (NVL(ICTSTAT1.WHSE_QTY_BEG,0)) BOOK" _
        '       & " from ICTSTAT1 where NVL(ICTSTAT1.WHSE_QTY_BEG,0) <> 0" _
        '       & Replace(SQLW, "X.WHSE_CODE", "ICTSTAT1.WHSE_CODE") _
        '       & " and ICTSTAT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
        '       & " group by ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE, ICTSTAT1.WHSE_CODE"

        ICTPHYV1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Alter Table " & ICTPHYV1 & " Add Primary Key (STYLE_CODE,COLOR_CODE,WHSE_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTPHYV1 & " Add SORT_VALUE VARCHAR2(30)")

        SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_LOCATOR = '1')"
        If ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.sql = "Select X.*, ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.ROYALTY_CODE" & vbCrLf _
                & " from ICTSTYL1, ICTSTYC1, (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, WHSE_CODE, LOCATION_CODE, Sum (PHYS) PHYS, Sum (BOOK) BOOK from (" & vbCrLf _
                & "Select ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS, 0 BOOK" _
                & " from ICTPHYC1,ICTPHYC2 where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" _
                & RGI_FILEDS & vbCrLf _
                & Replace(SQLW, "X.WHSE_CODE", "ICTPHYC2.WHSE_CODE") _
                & " group by ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", 0 PHYS, Sum (NVL(WHTLOCB0.LOCATION_QTY,0)) - Sum (nvl(WHTLOCB0.BOOK_INVTY_ADJ, 0)) BOOK" _
                & " from WHTLOCB0,ICTSTYL1 where ICTSTYL1.STYLE_CODE = WHTLOCB0.STYLE_CODE" _
                & Replace(SQLW, "X.WHSE_CODE", "WHTLOCB0.WHSE_CODE") _
                & " group by WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE, WHSE_CODE, LOCATION_CODE) X" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & " and ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & " and ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE"
        Else
            ASCMAIN1.sql = "Select X.*, ICTCOSTA.STYLE_COST" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_GROUP_CODE, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.ROYALTY_CODE, ICTSTYL1.CUST_CODE" & vbCrLf _
                & " from ICTSTYL1, ICTCOSTA, (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, WHSE_CODE, LOCATION_CODE, Sum (PHYS) PHYS, Sum (BOOK) BOOK from (" & vbCrLf _
                & "Select ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS, 0 BOOK" _
                & " from ICTPHYC1,ICTPHYC2 where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" _
                & Replace(SQLW, "X.WHSE_CODE", "ICTPHYC2.WHSE_CODE") _
                & " group by ICTPHYC2.STYLE_CODE, ICTPHYC2.COLOR_CODE, ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", 0 PHYS, Sum (NVL(WHTLOCB0.LOCATION_QTY,0)) BOOK" _
                & " from WHTLOCB0,ICTSTYL1 where ICTSTYL1.STYLE_CODE = WHTLOCB0.STYLE_CODE" _
                & Replace(SQLW, "X.WHSE_CODE", "WHTLOCB0.WHSE_CODE") _
                & " group by WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE, WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE, WHSE_CODE, LOCATION_CODE) X" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & sqlASN _
                & " and ICTCOSTA.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & " and ICTCOSTA.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                & " and ICTCOSTA.OPS_YYYYPP (+) = '" & LYP & "'"
        End If
        WHTPHYV1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Alter Table " & WHTPHYV1 & " Add Primary Key (STYLE_CODE,COLOR_CODE,WHSE_CODE,LOCATION_CODE)")
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            End If

            If chkUpdateVariances.Checked Then
                For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("")
                    Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")

                    If COLUMN_NAME = "WHSE_CODE" Then
                        If Val(rowASTDSQLA.Item("SEQUENCE") & "") <> 1 Then
                            EMsg &= vbCr & "Warehouse MUST be the 1st sort sequence when updating"
                        End If
                    End If
                    If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                        If COLUMN_NAME = "WHSE_CODE" Then
                            ' this is ok
                        Else
                            EMsg &= vbCr & "You may NOT specify filter criteria for any field (other than Warehouse) when updating"
                        End If
                    End If
                Next
                If optSORT.Value <> "I" Then
                    EMsg &= vbCr & "Invalid Sort Option when enabling Update - must be by Item Code"
                End If
            End If

            If chkUnitVarianceRange.Checked Then
                If Val(numVARMIN.Value) > Val(numVARMAX.Value) Or Val(numVARMAX.Value) <= 0 Or Val(numVARMIN.Value) < 0 Then
                    EMsg &= vbCr & "Invalid Unit Variance Range"
                End If
                If Val(numVARU.Value) <> 0 Then
                    EMsg &= vbCr & "Cannot specify a Unit Variance threshhold when specifying a Unit Variance Range"
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

        If ASCMAIN1.CLIENT = "VAN" Then
            Throw New Exception("Cannot Use this Variance Report to record Variances")
        End If

        If current_period_physical Then
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select * from " & ICTPHYV1 & ";" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTSTAT1 Set WHSE_QTY_PHY = NVL(WHSE_QTY_PHY,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_PHY) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & ASCMAIN1.CYP & "',NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
                & "   Update ICTSTAT2 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT2 (STYLE_CODE,COLOR_CODE,WHSE_CODE,WHSE_QTY_ON_HAND) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
                & "  End Loop; " & vbCrLf _
                & " End; " & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select * from " & ICTPHYV1 & ";" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTSTAT1 Set WHSE_QTY_BEG = R1.PHYS where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_BEG) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & ASCMAIN1.CYP & "',R1.PHYS); End If;" & vbCrLf _
                & "   Update ICTSTAT1 Set WHSE_QTY_PHY = NVL(WHSE_QTY_PHY,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & LYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_PHY) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & LYP & "',NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
                & "   Update ICTSTAT2 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT2 (STYLE_CODE,COLOR_CODE,WHSE_CODE,WHSE_QTY_ON_HAND) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
                & "   Update ICTSTAT5 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & LYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT5 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_ON_HAND) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & LYP & "',NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
                & "  End Loop; " & vbCrLf _
                & " End; " & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from " & WHTPHYV1 & " where NVL(PHYS,0) - NVL(BOOK,0) <> 0;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update WHTLOCB1 Set LOCATION_QTY = NVL(LOCATION_QTY,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and LOCATION_CODE = R1.LOCATION_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then Insert into WHTLOCB1 (WHSE_CODE,LOCATION_CODE,BAR_CODE,STYLE_CODE,COLOR_CODE,LOCATION_QTY) Values (R1.WHSE_CODE,R1.LOCATION_CODE,'0000000000',R1.STYLE_CODE,R1.COLOR_CODE,NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
            & "   Insert into WHTLOCB2 (WHSE_CODE,LOCATION_CODE,BAR_CODE,STYLE_CODE,COLOR_CODE,WHSE_TRAN_QTY,WHSE_TRAN_TYPE,WHSE_TRAN_NO,WHSE_TRAN_LNO,INIT_DATE,INIT_OPER,LOCATION_CODE_OTHER,SESSION_NO) " & vbCrLf _
            & "    Values (R1.WHSE_CODE,R1.LOCATION_CODE,'0000000000',R1.STYLE_CODE,R1.COLOR_CODE,NVL(R1.PHYS,0) - NVL(R1.BOOK,0),'P','0000000000',0,SYSDATE,'" & ASCMAIN1.USER_ID & "',NULL,'" & ASCMAIN1.SESSION_NO & "');" & vbCrLf _
            & "  End Loop; " & vbCrLf _
            & " End; " & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Dim SQLW As String = ""
        If WHSE_CODEs <> "" Then
            SQLW = " and X.WHSE_CODE in ('" & Replace(WHSE_CODEs, ",", "','") & "')"
        End If
        SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_PHYS_STATUS = 'C')"

        ASCMAIN1.sql = "" _
          & "Update ICTWHSE1 X Set WHSE_YYYYPP_LAST_PHY = '" & LYP & "', WHSE_PHYS_STATUS = NULL" & ASCMAIN1.SQL_Add_WHERE(SQLW)
        ASCDATA1.ExecuteSQL()

    End Sub

End Class