Public Class SOROPEN1
    Dim tblSOTGROUP As DataTable
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Range_Events(grpORDR_DATE_BOOKED)
        Range_Events(grpORDR_SHIP_DATE)
        Range_Events(grpORDR_CANCEL_DATE)

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            chkShowAllo.Checked = True
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            chkShowPrice.Visible = True
            Dim SHIPPED_DATE_F As DateTime = DateSerial(Now().Year, Now().Month, 1)
            Dim SHIPPED_DATE_L As DateTime = SHIPPED_DATE_F.AddMonths(1).AddDays(-1)
            Absx1.dteFor("SHIPPED_DATE_F").Value = Format(SHIPPED_DATE_F, "MM/dd/yyyy")
            Absx1.dteFor("SHIPPED_DATE_L").Value = Format(SHIPPED_DATE_L, "MM/dd/yyyy")
        Else
            chkShowPrice.Visible = False
            For Each vli As ValueListItem In optSTATUS.Items
                If vli.DataValue = "S" Then
                    optSTATUS.Items.Remove(vli)
                End If
            Next
        End If
        chkShowPrice.Checked = False

        sql = "Select ORDR_GROUP_NO FROM SOTORDR1 WHERE ROWNUM < 0"
        tblSOTGROUP = ASCDATA1.GetDataTable(sql, "SOTGROUP")

        grdSOTGROUP.DataSource = tblSOTGROUP

        Get_PARM("SOTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")

        For Each ORDR_TYPE As String In New String() {"O", "R"}
            If Absx1.optFor("OPTORDERS").Value = "A" Or Absx1.optFor("OPTORDERS").Value = ORDR_TYPE Then
                If ORDR_TYPE = "R" And (Absx1.optFor("OPTEDI").Value = "E" Or Absx1.optFor("OPTSTATUS").Value = "P") Then
                Else
                    If ORDR_TYPE = "O" Then
                        If Absx1.optFor("OPTEDI").Value <> "A" Then sql_filter = " and SOTORDR1.ORDR_SOURCE = '" & Absx1.optFor("OPTEDI").Value & "'"
                        If Absx1.optFor("OPTSTATUS").Value = "A" Then
                            sql_filter = " and SOTORDR1.ORDR_STATUS <> 'D'"
                        Else
                            If Absx1.optFor("OPTSTATUS").Value = "X" Then
                                sql_filter = " and SOTORDR1.ORDR_STATUS in ('O','P')"
                            Else
                                If Absx1.optFor("OPTSTATUS").Value = "S" Then
                                    Dim TEMPGROUP As String = ""
                                    If tblSOTGROUP.Rows.Count = 0 Then
                                        Dim OS As String = Format(Absx1.dteFor("SHIPPED_DATE_F").Value, "dd-MMM-yyyy")
                                        Dim OE As String = Format(Absx1.dteFor("SHIPPED_DATE_L").Value, "dd-MMM-yyyy")
                                        ASCMAIN1.sql = String.Format("SELECT DISTINCT ORDR_NO FROM SOTINVH1 WHERE INV_DATE >= '{0}' AND INV_DATE <= '{1}'", OS, OE)
                                        TEMPGROUP = ASCMAIN1.Temp_Table
                                    Else
                                        ASCMAIN1.sql = "SELECT DISTINCT ORDR_NO FROM SOTINVH1 WHERE ROWNUM < 0"
                                        TEMPGROUP = ASCMAIN1.Temp_Table
                                        For Each rowSOTGROUP As DataRow In tblSOTGROUP.Select()
                                            Dim ORDR_GROUP_NO As String = rowSOTGROUP.Item("ORDR_GROUP_NO") & String.Empty
                                            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                                            SQLS.AppendLine(String.Format("INSERT INTO {0} ", TEMPGROUP))
                                            SQLS.AppendLine("SELECT ORDR_NO FROM SOTORDR1")
                                            SQLS.AppendLine(String.Format("WHERE ORDR_GROUP_NO = '{0}'", rowSOTGROUP.Item("ORDR_GROUP_NO").ToString))
                                            ASCMAIN1.sql = SQLS.ToString
                                            ASCDATA1.ExecuteSQL()
                                        Next
                                    End If
                                    sql_filter = String.Format(" and SOTORDR1.ORDR_NO IN (SELECT DISTINCT ORDR_NO FROM {0})", TEMPGROUP)
                                Else
                                    sql_filter = " and SOTORDR1.ORDR_STATUS = '" & Absx1.optFor("OPTSTATUS").Value & "'"
                                    If Absx1.optFor("OPTSTATUS").Value = "O" Then
                                        sql_filter &= " and SOTORDR2.ORDR_QTY_OPEN <> 0"
                                    Else
                                        sql_filter &= " and SOTORDR2.ORDR_QTY_PICK <> 0"
                                    End If
                                End If
                            End If
                        End If
                    Else
                        sql_filter = Replace(sql_WHERE, "SOTORDR1.ORDR_GROUP_NO", "'0000000000'")
                        sql_filter = " and SOTORDR1.RSRV_STATUS = 'O'"
                        sql_filter = " and SOTORDR2.RSRV_QTY_OPEN <> 0"
                    End If

                    If Absx1.optFor("OPTASN").Value = "S" Then sql_filter &= " and ICTSTYL1.CUST_CODE is Null"
                    If Absx1.optFor("OPTASN").Value = "N" Then sql_filter &= " and ICTSTYL1.CUST_CODE is Not Null"

                    Dim OPT_G As String = "SOTORDR1.ORDR_GROUP_NO"
                    Dim OPT_S As String = "SOTORDR2.STYLE_CODE"
                    Dim OPT_C As String = "SOTORDR2.COLOR_CODE"

                    If Absx1.optFor("OPTDTL").Value = "3" Then
                        OPT_G = "'0000000000'"
                    End If

                    If Absx1.optFor("OPTDTL").Value <> "2" Then
                        OPT_S = "'X'"
                        OPT_C = "'X'"
                    End If

                    sql = "Select " & sql_SELECT_cols & vbCrLf _
                        & ", '" & ORDR_TYPE & "' ORDR_TYPE" & vbCrLf _
                        & ", " & OPT_G & " ORDR_GROUP_NO" & vbCrLf _
                        & ", " & OPT_S & " STYLE_CODE, " & OPT_C & " COLOR_CODE" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY,0)) ORDR_QTY" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_ALLO,0)) ORDR_QTY_ALLO" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_PICK,0)) ORDR_QTY_PICK" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0)) ORDR_QTY_CANC" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_ALLO,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_ALLO" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_OPEN,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_OPEN" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_PICK,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_PICK" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_SHIP,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_SHIP" & vbCrLf _
                        & ", SUM(NVL(SOTORDR2.ORDR_QTY_CANC,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_CANC" & vbCrLf _
                        & " from " & "SOTORDR1" & sql_TABLE_NAMEs & vbCrLf _
                        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & Get_Dates(ORDR_TYPE)) & vbCrLf _
                        & " group by " & IIf(sql_GROUP_BY_cols = "", "'" & ORDR_TYPE & "'", sql_GROUP_BY_cols) & vbCrLf _
                        & ", " & OPT_G & "" & vbCrLf _
                        & ", " & OPT_S & ", " & OPT_C & "" & vbCrLf

                    If ORDR_TYPE = "R" Then
                        sql = Replace(sql, ".ORDR_GROUP_NO", ".RSRV_NO")
                        sql = Replace(sql, "NVL(SOTORDR2.ORDR_QTY_PICK,0)", "0")
                        sql = Replace(sql, "NVL(SOTORDR2.ORDR_QTY_SHIP,0)", "0")
                        sql = Replace(sql, ".ORDR_QTY", ".RSRV_QTY")
                        sql = Replace(sql, ".ORDR_NO", ".RSRV_NO")
                        sql = Replace(sql, "ORDR_STATUS", "RSRV_STATUS")
                        sql = Replace(sql, "SOTORDR1", "SOTRSRV1")
                        sql = Replace(sql, "SOTORDR2", "SOTRSRV2")
                    End If
                    ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")
                End If
            End If
        Next

        sql = "Select ORDR_TYPE, ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE" & ASTSRPT1_sum_columns _
            & " from " & ASTSRPT1 & " group by ORDR_TYPE, ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDRY", 4))

        sql = "Select SOTAUTH1.*" _
            & " from SOTAUTH1 where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTAUTH1", 1))

        sql = "Select SOTORDRS.*" _
            & " from SOTORDRS where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDRS", 3))

        sql = "Select SOTORDRG.*" _
            & " from SOTORDRG where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDRG", 1))

        sql = "Select X.STYLE_CODE, ICTSTYL1.STYLE_DESC" _
            & " from (Select Distinct STYLE_CODE from " & ASTSRPT1 & ") X, ICTSTYL1" _
            & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTSTYLX", 1))

        sql = "Select X.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
            & " from (Select Distinct COLOR_CODE from " & ASTSRPT1 & ") X, ICTCOLR1" _
            & " where ICTCOLR1.COLOR_CODE = X.COLOR_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTCOLRX", 1))


        If Absx1.optFor("OPTDTL").Value = "2" Then
            ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " Set STYLE_CODE = NULL, COLOR_CODE = NULL")
        End If

        sql = "" _
            & "Select 'O' ORDR_TYPE, ORDR_GROUP_NO, CUST_CODE, ORDR_DATE" & vbCrLf _
            & ", ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_CUST_PO, ORDR_DEPT" & vbCrLf _
            & ", ORDR_CNT, ORDR_CNT_OPEN, ORDR_CNT_PICK, CUST_DC_NO" & vbCrLf _
            & ", SALES_DIVISION_CODE, WHSE_CODE from SOTORDR0 " & vbCrLf _
            & " where ORDR_GROUP_NO in " _
            & " (Select Distinct ORDR_GROUP_NO from " & ASTSRPT1 & " where ORDR_TYPE = 'O')" & vbCrLf _
            & " union " & vbCrLf _
            & "Select 'R' ORDR_TYPE, RSRV_NO ORDR_GROUP_NO, CUST_CODE, TRUNC(INIT_DATE) ORDR_DATE" & vbCrLf _
            & ", ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_CUST_PO, ORDR_DEPT" & vbCrLf _
            & ", 1 ORDR_CNT, 1 ORDR_CNT_OPEN, 0 ORDR_CNT_PICK, NULL CUST_DC_NO" & vbCrLf _
            & ", SALES_DIVISION_CODE, WHSE_CODE from SOTRSRV1 " & vbCrLf _
            & " where RSRV_NO in " _
            & " (Select Distinct ORDR_GROUP_NO from " & ASTSRPT1 & " where ORDR_TYPE = 'R')" & vbCrLf
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDRX", 2))



    End Sub

    Function Get_Dates(TYPE As String) As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"ORDR_DATE_BOOKED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                sql = sql & " and SOTORDR1." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
            End If
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                sql = sql & " and SOTORDR1." & COLUMN_NAME & " <= '" & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "dd-MMM-yyyy") & "'"
            End If
        Next
        sql = Replace(sql, "SOTORDR1.ORDR_DATE_BOOKED", "TRUNC(SOTORDR1.INIT_DATE)")
        Return sql
    End Function

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        'ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = 0")
        'ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = TRUNC(100 * ORDR_AMT / ORDR_QTY) / 100 where ORDR_QTY <> 0")
    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = ""

        Page0.Add("Styles: " & Absx1.optFor("OPTASN").Text)
        If Absx1.optFor("OPTASN").Value <> "A" Then SUBT &= Absx1.optFor("OPTASN").Text & ", "

        Page0.Add("Report Detail: " & Absx1.optFor("OPTDTL").Text)

        Page0.Add("Orders & Reservations: " & Absx1.optFor("OPTORDERS").Text)
        If Absx1.optFor("OPTORDERS").Value <> "A" Then SUBT &= Absx1.optFor("OPTORDERS").Text & ", "

        Page0.Add("Status: " & Absx1.optFor("OPTSTATUS").Text)
        If Absx1.optFor("OPTSTATUS").Value <> "A" Then SUBT &= Absx1.optFor("OPTSTATUS").Text & ", "

        Page0.Add("Orders: " & Absx1.optFor("OPTEDI").Text)
        If Absx1.optFor("OPTEDI").Value <> "A" Then SUBT &= Absx1.optFor("OPTEDI").Text & ", "

        For Each COLUMN_NAME As String In New String() {"ORDR_DATE_BOOKED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            Dim Z As String = Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text & ":"
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                Z &= " from First"
            Else
                Z &= " from " & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "MM/dd/yyyy")
            End If
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                Z &= " to Last"
            Else
                Z &= " to " & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "MM/dd/yyyy")
            End If
            Page0.Add(Z)
        Next

        '   ASCMAIN1.CR_RPT.RecordSelectionFormula = "RSF"
        ' Stop ' CR_Rpt.ParameterFields("LVLS").SetCurrentValue(CStr(cfmax + 1))
        CR_params.Add("CHKDTL", IIf(Absx1.optFor("OPTDTL").Value = "2", "1", "0"))
        CR_params.Add("OPTDTL", Absx1.optFor("OPTDTL").Value)

        If Absx1.optFor("OPTSTATUS").Value = "X" Then
            RPT_TITLE = "Sales Order Report"
        End If

        If optDTL.Value = "2" And chkShowAllo.Checked Then RPT = "SOROPENN"

        If Absx1.chkFor("CHKSHOW_PRICE").Checked Then RPT = "SOROPENP"
        'CHKSHOW_PRICE

        Generate_Report(RPT, RPT_TITLE, SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("STYLE_CODE")

                If Absx1.optFor("OPTDTL").Value = "2" And Val(rowASTDSQLA("SEQUENCE") & "") <> 0 Then
                    EMsg &= "You Must NOT Sort by Style when showing Details"
                End If

                If Absx1.optFor("OPTSTATUS").Value = "S" Then
                    If Not (IsDate(Absx1.dteFor("SHIPPED_DATE_F").Value) And IsDate(Absx1.dteFor("SHIPPED_DATE_L").Value)) Then
                        EMsg &= "You Must Select Dates When Filtering By Shipped Orders"
                    End If
                    If Absx1.optFor("OPTORDERS").Value <> "O" Then
                        EMsg &= "You May Only Select OrdersWhen Filtering By Shipped Orders"
                    End If
                End If

                If ASCMAIN1.CLIENT = "NYA" AndAlso ASCMAIN1.USER_CODES = "CA" Then

                    rowASTDSQLA = tblASTDSQLA.Rows.Find("WHSE_CODE")

                    ' Force Canada settings
                    rowASTDSQLA("EXCLUDE") = "0"
                    rowASTDSQLA("CODE_VALUES") = TAC.TACMAIN1.NyaCanadaWhseCommaSeparatedString  ' "18"

                    'If rowASTDSQLA("EXCLUDE") & "" = "0" And rowASTDSQLA("CODE_VALUES") & "" = "18" Then
                    '    ' OK
                    'Else
                    '    EMsg &= "Whse 18 Only"
                    'End If
                End If

        End Select
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Sub Pivot_Prepare_PreProcess(dt As DataTable)
        dt.Columns.Remove("ORDR_GROUP_NO")
        dt.Columns.Remove("STYLE_CODE")
        dt.Columns.Remove("COLOR_CODE")
    End Sub

    Private Sub optDTL_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optDTL.ValueChanged
        chkShowAllo.Visible = (optDTL.Value = "2")
    End Sub

    Private Sub optSTATUS_ValueChanged(sender As Object, e As EventArgs) Handles optSTATUS.ValueChanged
        If optSTATUS.Value = "S" Then
            grpOrdersShipped.Visible = True
            grpSelectGroups.Visible = True
            optORDERS.Value = "O"
        Else
            grpOrdersShipped.Visible = False
            grpSelectGroups.Visible = False
            optORDERS.Value = "A"
        End If
    End Sub

    Private Sub btnSelectGroups_Click(sender As Object, e As EventArgs) Handles btnSelectGroups.Click
        Dim OS As String = Format(Absx1.dteFor("SHIPPED_DATE_F").Value, "dd-MMM-yyyy")
        Dim OE As String = Format(Absx1.dteFor("SHIPPED_DATE_L").Value, "dd-MMM-yyyy")

        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("SELECT")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("I1.INV_DATE,")
        S.AppendLine("O1.ORDR_GROUP_NO,")
        S.AppendLine("I1.ORDR_CUST_PO,")
        S.AppendLine("SUM(I1.INV_TOTAL_AMOUNT) AS INV_TOTAL_AMOUNT")
        S.AppendLine("FROM SOTINVH1 I1, SOTORDR1 O1")
        S.AppendLine("WHERE I1.ORDR_NO = O1.ORDR_NO")
        S.AppendLine("AND I1.INV_TYPE = 'I'")
        S.AppendLine(String.Format("AND (I1.INV_DATE >= '{0}' AND I1.INV_DATE <= '{1}')", OS, OE))
        S.AppendLine("GROUP BY")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("I1.INV_DATE,")
        S.AppendLine("O1.ORDR_GROUP_NO,")
        S.AppendLine("I1.ORDR_CUST_PO")
        S.AppendLine("ORDER BY")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("I1.INV_DATE,")
        S.AppendLine("O1.ORDR_GROUP_NO,")
        S.AppendLine("I1.ORDR_CUST_PO")
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Please Select PO(s)"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For Each dr As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                Dim ORDR_GROUP_NO As String = dr.Item("ORDR_GROUP_NO") & String.Empty
                Dim rowSOTGROUP As DataRow = tblSOTGROUP.NewRow
                rowSOTGROUP.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                tblSOTGROUP.Rows.Add(rowSOTGROUP)
            Next
            grdSOTGROUP.Refresh()
        End If
    End Sub
End Class