Public Class ICRPHYVV
    Dim LYY As String = ""
    Dim LYP As String = ""
    Dim ICTPHYVX As String = ""
    Dim Record_No As Integer = 0

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'setup report definition with these keys(sql = "Select 0 RECORD_NO, 0 COUNTER, WHSE_CODE, STYLE_CODE, COLOR_CODE, ")
        numVARC.Value = 0
        numVARU.Value = 0
        optSORT.Value = "S"
        optASN.Value = "A"
        optBA.Value = "A"

        'Per Gabe and Maruice.  No Updates for anyone except Gabe. Gary Updated by accident 1/6/03.
        IIf(UCase(ASCMAIN1.USER_ID) <> UCase("wayne"), RWU = "N", RWU = "Y")

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options


        Dim sql_filter As String = ""
        sql_filter &= SQL_in("SALES_DIVISION_CODE", "ICTSTYL1.SALES_DIVISION_CODE")
        sql_filter &= SQL_in("FABRIC_CODE", "ICTSTYL1.FABRIC_CODE")
        sql_filter &= SQL_in("SEASON_CODE", "ICTSTYL1.SEASON_CODE")
        sql_filter &= SQL_in("SUB_BODY_CODE", "ICTSTYL1.SUB_BODY_CODE")
        sql_filter &= SQL_in("CUST_CODE", "ICTSTYL1.CUST_CODE")
        sql_filter &= SQL_in("FASHION_PROMO", "ICTSTYL1.FASHION_PROMO")
        sql_filter &= SQL_in("CMT_NO", "ICTSTYL1.CMT_NO")
        sql_filter &= SQL_in("WHSE_CODE", "X.WHSE_CODE")
        sql_filter &= SQL_in("STYLE_CODE", "X.STYLE_CODE")
        Record_No = 0


        If optBA.Value <> "F" Or sql_filter <> "" Or numVARU.Value <> 0 Or numVARC.Value <> 0 Or optASN.Value <> "A" Then
            RWU = "N"
        End If

        LYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        LYY = Val(Mid(LYP, 4)) - 1
        ASCMAIN1.Progress("Now Preparing Dataset")

        'Create Table ICTPHYVX as 
        ASCMAIN1.sql = "Select 0 RECORD_NO, ICTSTAT1.WHSE_CODE, " & vbCrLf _
            & " ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE, " & vbCrLf _
            & " ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_COST, " & vbCrLf _
            & " ICTSTYL1.SALES_DIVISION_CODE, ICTSTAT1.WHSE_QTY_BEG," & vbCrLf _
            & " ICTSTYL1.FABRIC_CODE, ICTSTYL1.SEASON_CODE, ICTSTYL1.SUB_BODY_CODE, " & vbCrLf _
            & " ICTSTYL1.CUST_CODE, ICTSTYL1.FASHION_PROMO, ICTSTYL1.CMT_NO, " & vbCrLf _
            & " 0 PHYS_COUNT, 0 BOOK, 0 PHYS, 0 MF_COST, 0 LYV, 0 PRE_COUNT" & vbCrLf _
            & " from ICTSTAT1, ICTSTYL1, ICTWHSE1 Where Rownum < 1" & vbCrLf
        ICTPHYVX = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCMAIN1.sql = "Alter Table " & ICTPHYVX & " Add Primary Key (RECORD_NO)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Call MyBase.Get_SQL("*", ICTPHYVX)

        Prepare_dst(True, sql_filter)

        sql = "Select  " & sql_SELECT_cols _
        & " , RECORD_NO " & vbCrLf _
        & " from " & ICTPHYVX & " ICTPHYVX " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN) & vbCrLf _
        & " group by  " & sql_GROUP_BY_cols & vbCrLf _
        & " , RECORD_NO "
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")


    End Sub



    Public Overrides Sub Print_Report()

        If numVARU.Value <> 0 Then
            Write_Page0("Unit Variance Threshold: " & CStr(numVARU.Value))
        End If
        If numVARC.Value <> 0 Then
            Write_Page0("Cost Variance Threshold: " & CStr(numVARC.Value))
        End If

        Select Case optASN.Value
            Case "A"
                Write_Page0("All Styles")
            Case "S"
                Write_Page0("Stock Styles Only")
            Case "N"
                Write_Page0("Non-Stock Styles Only")
        End Select
        If chkDZ.Checked Then
            Write_Page0("In Dozens")
        End If

        RPT = IIf(optSORT.Value = "S", "ICRPHYVV2", "ICRPHYVV")
        CR_params.Add("SUBT", txtDescription.Text & IIf(chkFYCOST.Checked, " Using Prior Fiscal Year Costs", ""))
        CR_params.Add("OPTR", optSORT.Value)
        CR_params.Add("CHKDZ", chkDZ.CheckedValue)
        Generate_Report(RPT, , SUBT)

    End Sub

    Sub Write_Page0(Desc As String)
        Dim row As DataRow = dst.Tables("ASTPAGE0").NewRow
        With row
            .Item("LINE_NO") = Val(dst.Tables("ASTPAGE0").Compute("MAX(LINE_NO)", "") & "") + 1
            .Item("LINE_DATA") = Desc
        End With
        dst.Tables("ASTPAGE0").Rows.Add(row)

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
        ' If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst
            Create_TDA(.Tables.Add, "ICTPHYV9", "*", 0, True, "", 4)

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from " & ICTPHYVX
            Create_TDA(.Tables.Add, ICTPHYVX, "**", 0, True, "", 1)

            ASCMAIN1.sql = "Select * from " & ICTPHYVX
            Create_TDA(.Tables.Add, "ICTPHYVV", "**", 0, False, "", 1)
            .Tables("ICTPHYVV").Columns.Add("SORTED_BY")
            .Tables("ICTPHYVV").Columns.Add("ABS_VARIANCE", GetType(System.Decimal))
            .Tables("ICTPHYVV").Columns.Add("ABS_STYLE_COST", GetType(System.Decimal))

            ASCMAIN1.sql = "SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, " & vbCrLf _
            & " ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
            & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
            & " FROM ICTCOST1" & vbCrLf
            Create_TDA(.Tables.Add, "ICTCOST1", "**", 0, False, "", 0)
            .Tables("ICTCOST1").Columns.Add("TRAN_REF")
            .Tables("ICTCOST1").Columns.Add("QTY_USED", GetType(System.Decimal))
            .Tables("ICTCOST1").Columns.Add("COST_TOTAL", GetType(System.Decimal), "ISNULL(TRAN_COST,0) * ISNULL(QTY_USED,0)")
        End With

        ASCMAIN1.sql = "Select '" & ASCMAIN1.CYP & "' OPS_YYYYPP, " & vbCrLf _
        & " STYLE_CODE, COLOR_CODE, WHSE_CODE, " & vbCrLf _
        & " WHSE_QTY_ON_HAND WHSE_QTY_BEG" & vbCrLf _
        & " FROM BATSTAT2"
        Dim TT As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCMAIN1.sql = "Alter Table " & TT & " Add Primary Key (OPS_YYYYPP, STYLE_CODE, COLOR_CODE, WHSE_CODE)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.Progress("Calculating Book", "")
        Get_Book(TT, sqlw)


        'this was a fix for 2012 only

        ''Create Table BATSTAT2_SHIPMENT as
        ''Select '201301' OPS_YYYYPP, R2.STYLE_CODE, R2.COLOR_CODE, R1.WHSE_CODE,
        ''Sum(P3.PO_QTY_REC) as WHSE_QTY_BEG 
        ''from POTSHIP2 P2, potship3 P3, POTORDR2 R2, POTORDR1 R1
        ''Where P3.PO_SHIPMENT_NO in ('015276','015290')
        ''And P2.PO_SHIPMENT_NO = P3.PO_SHIPMENT_NO
        ''And P2.PO_SHIPMENT_LNO = P3.PO_SHIPMENT_LNO
        ''And R1.PO_ORDER_NO = R2.PO_ORDER_NO
        ''And P3.PO_ORDER_NO = R2.PO_ORDER_NO
        ''And P3.PO_ORDER_LNO = R2.PO_ORDER_LNO
        ''Group By  R2.STYLE_CODE, R2.COLOR_CODE, R1.WHSE_CODE

        ''Insert into BATSTAT2_SHIPMENT
        ''Select '201301' OPS_YYYYPP, R2.STYLE_CODE, R2.COLOR_CODE, R1.WHSE_CODE,
        ''Sum(P3.PO_QTY_REC) as WHSE_QTY_BEG 
        ''from POTSHIP2 P2, potship3 P3, POTORDR2 R2, POTORDR1 R1
        ''Where P3.PO_SHIPMENT_NO in ('015277')
        ''And P2.PO_SHIPMENT_NO = P3.PO_SHIPMENT_NO
        ''And P2.PO_SHIPMENT_LNO = P3.PO_SHIPMENT_LNO
        ''And R1.PO_ORDER_NO = R2.PO_ORDER_NO
        ''And P3.PO_ORDER_NO = R2.PO_ORDER_NO
        ''And P3.PO_ORDER_LNO = R2.PO_ORDER_LNO
        ''And CONTAINER_NO in ('DRYU4038096','GESU5066134','CLHU4592717')
        ''Group By R2.STYLE_CODE, R2.COLOR_CODE, R1.WHSE_CODE

        ASCMAIN1.Progress("Adding Shipments to Book", "")
        Get_Book("BATSTAT2_SHIPMENT", sqlw) 'this was to include closed shipments ADS had closed

        'Create table BATSTAT2_PPK as 
        'Select '201301' as OPS_YYYYPP, STYLE_CODE, COLOR_CODE, 'NJ' WHSE_CODE , QTY_REC as WHSE_QTY_BEG
        'FROM ICTIREC2 WHERE RECEIPT_NO in ('079972','079973')
        ASCMAIN1.Progress("Loading PPK Fix", "")
        Get_Book("BATSTAT2_PPK", sqlw) ' this was used to correct PPK error for 2 shipments

        ASCMAIN1.Progress("Loading Post Fix", "")
        Get_Count_SQL("BATLOCP1_FIX", sqlw) ' this was used to bring remove shipments that ADS had open but Vandale had billed

        ASCMAIN1.Progress("Loading CAHEL Counts", "")
        Get_Count_SQL("BATLOCP1_ADJUSTMENTS", sqlw) ' this was used to act as if CAHEL was counted

        ASCMAIN1.Progress("Loading CAHEL Counts", "")
        Get_Count_SQL("BATLOCP1_CAHEL", sqlw) ' this was used to act as if CAHEL was counted
        'End of Fix for 2012 Variance

        ASCMAIN1.Progress("Loading Pre Counts", "")
        Get_Count_SQL("BATLOCP1_PRE", sqlw)
        ASCMAIN1.Progress("Loading Post Counts", "")
        Get_Count_SQL("BATLOCP1_POST", sqlw)

        'Calculate Costs
        If chkCOST.Checked Then
            ASCMAIN1.Progress("Loading Cost", "")
            Dim SC As Double = 0
            Dim CostSource As Integer = 0
            Dim W As String = ""

            Dim COSTYP As String = IIf(chkFYCOST.Checked, Str(Val(Mid(ASCMAIN1.CYP, 1, 4))) - 1 & "12", ASCMAIN1.CYP)
            For Each rowICTPHYVX As DataRow In dst.Tables(ICTPHYVX).Select("", "STYLE_CODE, COLOR_CODE")
                ASCMAIN1.Progress("Costing", rowICTPHYVX.Item("STYLE_CODE") & "" & " - " & rowICTPHYVX.Item("COLOR_CODE") & "")
                SC = 0
                CostSource = 0
                W = TAC.ICCMAIN1.Calc_Cost_OH(Me, COSTYP, rowICTPHYVX.Item("STYLE_CODE"), rowICTPHYVX.Item("COLOR_CODE"), False)

                Dim a() As String = Split(W, "|")
                SC = Val(CDbl(a(0)))

                If SC = 0 Then 'Use the Masterfile cost.
                    Dim rowICTSTYL1 As DataRow = clsASCBASE1.LookUp("ICTSTYL1", rowICTPHYVX.Item("STYLE_CODE"))
                    SC = Val(rowICTSTYL1.Item("STYLE_COST") & "")
                    CostSource = 2
                End If
                rowICTPHYVX.Item("STYLE_COST") = SC
                rowICTPHYVX.Item("MF_COST") = CostSource
            Next
        End If


        If numVARU.Value <> 0 Then
            ASCDATA1.DeleteRows(ICTPHYVX, " (WHSE_QTY_BEG - PHYS_COUNT) < " & numVARU.Value & "  And (WHSE_QTY_BEG - PHYS_COUNT) >= 0")
            ASCDATA1.DeleteRows(ICTPHYVX, " (WHSE_QTY_BEG - PHYS_COUNT) > " & (numVARU.Value * -1) & "  And (WHSE_QTY_BEG - PHYS_COUNT) <= 0")
        End If
        If numVARC.Value <> 0 Then
            ASCDATA1.DeleteRows(ICTPHYVX, " (STYLE_COST * (WHSE_QTY_BEG - PHYS_COUNT)) < " & numVARC.Value & "  And (WHSE_QTY_BEG - PHYS_COUNT) >= 0")
            ASCDATA1.DeleteRows(ICTPHYVX, " (STYLE_COST * (WHSE_QTY_BEG - PHYS_COUNT)) > " & (numVARC.Value * -1) & "  And (WHSE_QTY_BEG - PHYS_COUNT) <= 0")
        End If

        ASCMAIN1.Progress("Loading Temp Table", "")
        Update_Record_TDA(ICTPHYVX)
        Fill_Records("ICTPHYVV")

        ASCMAIN1.sql = "Select * from ICTSTYL1 Where STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTPHYVX & ")"
        Fill_Records("ICTSTYL1", , , ASCMAIN1.sql)

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Sub Get_Book(Table_Name As String, sqlw As String)
        ASCMAIN1.sql = "Select X.WHSE_CODE, " & vbCrLf _
       & " X.STYLE_CODE, X.COLOR_CODE, " & vbCrLf _
       & " ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_COST, " & vbCrLf _
       & " ICTSTYL1.SALES_DIVISION_CODE, X.WHSE_QTY_BEG, " & vbCrLf _
       & " ICTSTYL1.FABRIC_CODE, ICTSTYL1.SEASON_CODE, ICTSTYL1.SUB_BODY_CODE, " & vbCrLf _
       & " ICTSTYL1.CUST_CODE, ICTSTYL1.FASHION_PROMO, ICTSTYL1.CMT_NO," & vbCrLf _
       & " 0 PHYS_COUNT, 1 BOOK, 0 PHYS, 0 MF_COST, " & vbCrLf _
       & " NVL(ICTPHYV9.VARIANCE,0) AS LYV, 0 PRE_COUNT" & vbCrLf _
       & " from " & Table_Name & " X," & vbCrLf _
       & " ICTSTYL1, ICTWHSE1, ICTPHYV9" & sql_JOIN & vbCrLf _
       & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
       & "   and ICTWHSE1.WHSE_CODE = X.WHSE_CODE" & vbCrLf _
       & "   AND X.WHSE_CODE = ICTPHYV9.WHSE_CODE (+)" & vbCrLf _
       & "   AND X.STYLE_CODE = ICTPHYV9.STYLE_CODE (+)" & vbCrLf _
       & "   AND X.COLOR_CODE = ICTPHYV9.COLOR_CODE (+)" & vbCrLf _
       & "   and X.OPS_YYYYPP like '" & Format(Now, "yyyy") & "%'" & vbCrLf _
       & "   AND ICTPHYV9.OPS_YYYY (+) = '" & LYY & "'" _
       & "   and X.WHSE_QTY_BEG <> 0" & vbCrLf _
       & IIf(optASN.Value = "S", "   and ICTSTYL1.CUST_CODE is Null", "") & vbCrLf _
       & IIf(optASN.Value = "N", "   and ICTSTYL1.CUST_CODE is Not Null", "") & vbCrLf _
       & sql_JOIN & sqlw & vbCrLf
        '& "   and X.OPS_YYYYPP = '" & Format(Now, "yyyy") & "01'" & vbCrLf _ 'Removed this line.. shouldnt be concernced about period since it is a snapshot
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Write_Records(row)
        Next
    End Sub

    Sub Get_Count_SQL(Table_Name As String, sqlw As String)
        ASCMAIN1.sql = "Select X.WHSE_CODE, " & vbCrLf _
        & " X.STYLE_CODE, X.COLOR_CODE, " & vbCrLf _
        & " ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_COST, " & vbCrLf _
        & " ICTSTYL1.SALES_DIVISION_CODE, 0 WHSE_QTY_BEG, " & vbCrLf _
        & " ICTSTYL1.FABRIC_CODE, ICTSTYL1.SEASON_CODE, ICTSTYL1.SUB_BODY_CODE, " & vbCrLf _
        & " ICTSTYL1.CUST_CODE, ICTSTYL1.FASHION_PROMO, ICTSTYL1.CMT_NO, " & vbCrLf _
        & IIf(Table_Name = "BATLOCP1_PRE", " 0 ", " X.") & "PHYS_COUNT, " & vbCrLf _
        & " 0 BOOK, 1 PHYS, 0 MF_COST, 0 LYV," & vbCrLf _
        & IIf(Table_Name = "BATLOCP1_PRE", " X.PHYS_COUNT", " 0 ") & " as PRE_COUNT " & vbCrLf _
        & " from (" & vbCrLf _
        & " SELECT WHSE_CODE, LOCATION_CODE TICKET_NO," & vbCrLf _
        & " ROWNUM TICKET_LNO, STYLE_CODE, COLOR_CODE," & vbCrLf _
        & " SUM(LOCATION_QTY) PHYS_COUNT FROM " & Table_Name & vbCrLf _
        & " WHERE LOCATION_CODE NOT IN ('00005A','00008A')" & vbCrLf _
        & " GROUP BY WHSE_CODE, LOCATION_CODE," & vbCrLf _
        & " ROWNUM, STYLE_CODE, COLOR_CODE" & vbCrLf _
        & " HAVING SUM(LOCATION_QTY) <> 0" & vbCrLf _
        & " ) X, ICTSTYL1, ICTWHSE1" & sql_JOIN & vbCrLf _
        & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
        & "   and ICTWHSE1.WHSE_CODE = X.WHSE_CODE" & vbCrLf _
        & IIf(optASN.Value = "S", "   and ICTSTYL1.CUST_CODE is Null", "") & vbCrLf _
        & IIf(optASN.Value = "N", "   and ICTSTYL1.CUST_CODE is Not Null", "") & vbCrLf _
        & sql_JOIN & sqlw & vbCrLf

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Write_Records(row)
        Next

    End Sub
    Sub Write_Records(row As DataRow)

        Dim Whse_Code As String = IIf(chkConsolidate.Checked, "XX", row.Item("WHSE_CODE") & "")
        Dim rows() As DataRow = dst.Tables(ICTPHYVX).Select("WHSE_CODE = '" & Whse_Code & "'" _
                                       & " And STYLE_CODE = '" & row.Item("STYLE_CODE") & "'" _
                                       & " And COLOR_CODE = '" & row.Item("COLOR_CODE") & "'")
        If rows.Length <> 0 Then
            'this was not working
            'Dim rowSEARCH As DataRow = dst.Tables("ICTPHYVV").Rows.Find(New Object() {row.Item("WHSE_CODE"), row.Item("STYLE_CODE"), row.Item("COLOR_CODE")})
            'If row IsNot Nothing Then
            rows(0).Item("WHSE_QTY_BEG") += row.Item("WHSE_QTY_BEG")
            rows(0).Item("PHYS_COUNT") += row.Item("PHYS_COUNT")
            rows(0).Item("PRE_COUNT") += row.Item("PRE_COUNT")
            rows(0).Item("PHYS") += 1
        Else
            Dim rowICTPHYVV As DataRow = dst.Tables(ICTPHYVX).NewRow
            For Each DC As DataColumn In dst.Tables(ICTPHYVX).Columns
                Dim COLUMN_NAME As String = DC.ColumnName
                With rowICTPHYVV
                    If COLUMN_NAME = "RECORD_NO" Then
                        Record_No += 1
                        .Item(COLUMN_NAME) = Record_No
                    ElseIf COLUMN_NAME = "WHSE_CODE" Then
                        .Item(COLUMN_NAME) = Whse_Code
                    Else
                        Select Case dst.Tables(ICTPHYVX).Columns(COLUMN_NAME).DataType.Name
                            Case "Int16", "Int64", "Int32", "Double", "Decimal"
                                .Item(COLUMN_NAME) = Val(row.Item(COLUMN_NAME) & "")
                            Case "DateTime"
                                If .Item(COLUMN_NAME) & "" <> "" Then
                                    .Item(COLUMN_NAME) = row.Item(COLUMN_NAME) & ""
                                End If
                            Case Else
                                .Item(COLUMN_NAME) = row.Item(COLUMN_NAME) & ""
                        End Select
                    End If
                End With
            Next
            dst.Tables(ICTPHYVX).Rows.Add(rowICTPHYVV)
        End If

    End Sub


    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        If parms.Length > 0 Then
        End If

        EnforceConstraints(False)
        'Fill_Records("ASTSRPT1")
        EnforceConstraints(True)
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If e.Tab.Key = "Other Run-Time Options" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        MsgBox("No Update Allowed Without ABS", vbOKOnly, "Cannot Update")
        Exit Sub

        'For Each rowICTPHYVV As DataRow In dst.Tables("ICTPHYVV").Select("WHSE_QTY_BEG <> PHYS_COUNT", "WHSE_CODE, STYLE_CODE, COLOR_CODE")
        '    Dim Whse_Code As String = rowICTPHYVV.Item("WHSE_CODE") & ""
        '    Dim Style_Code As String = rowICTPHYVV.Item("STYLE_CODE") & ""
        '    Dim Color_Code As String = rowICTPHYVV.Item("COLOR_CODE") & ""
        '    Dim QTY As Long = Val(rowICTPHYVV.Item("PHYS_COUNT") & "") - Val(rowICTPHYVV.Item("WHSE_QTY_BEG") & "")

        '    Dim rowICTSTAT1 As DataRow = dst.Tables("ICTSTAT1").Rows.Find(New Object() {ASCMAIN1.CYP, Style_Code, Color_Code, Whse_Code})
        '    If rowICTSTAT1 IsNot Nothing Then
        '        ASCMAIN1.sql = "Update ICTSTAT1 Set WHSE_QTY_PHY = " & Val(rowICTSTAT1.Item("WHSE_QTY_PHY") & "") + QTY & vbCrLf _
        '            & " Where OPS_YYYYPP  ='" & ASCMAIN1.CYP & "'" & vbCrLf _
        '            & " And STYLE_CODE  ='" & Style_Code & "'" & vbCrLf _
        '            & " And COLOR_CODE  ='" & Color_Code & "'" & vbCrLf _
        '            & " And WHSE_CODE  ='" & Whse_Code & "'" & vbCrLf
        '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        '    Else
        '        ASCMAIN1.sql = "Insert into ICTSTAT1 ('OPS_YYYYPP','STYLE_CODE','COLOR_CODE','WHSE_CODE','WHSE_QTY_PHY') Values " _
        '            & " ('" & ASCMAIN1.CYP & "','" & Style_Code & "','" & Color_Code & "','" & Whse_Code & "','" & QTY & "')"
        '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        '    End If

        '    Dim rowICTSTAT2 As DataRow = dst.Tables("ICTSTAT2").Rows.Find(New Object() {Style_Code, Color_Code, Whse_Code})
        '    If rowICTSTAT2 IsNot Nothing Then
        '        ASCMAIN1.sql = "Update ICTSTAT1 Set WHSE_QTY_ON_HAND = " & Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "") + QTY & vbCrLf _
        '            & " Where STYLE_CODE  ='" & Style_Code & "'" & vbCrLf _
        '            & " And COLOR_CODE  ='" & Color_Code & "'" & vbCrLf _
        '            & " And WHSE_CODE  ='" & Whse_Code & "'" & vbCrLf
        '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        '    Else
        '        ASCMAIN1.sql = "Insert into ICTSTAT2 ('STYLE_CODE','COLOR_CODE','WHSE_CODE','WHSE_QTY_ON_HAND') Values " _
        '         & " ('" & Style_Code & "','" & Color_Code & "','" & Whse_Code & "','" & QTY & "')"
        '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        '    End If

        '    Dim rowICTPHYV9 As DataRow = dst.Tables("ICTPHYV9").NewRow
        '    With rowICTPHYV9
        '        .Item("OPS_YYYY") = Val(Mid(ASCMAIN1.CYP, 1, 4)) - 1
        '        .Item("WHSE_CODE") = Whse_Code
        '        .Item("STYLE_CODE") = Style_Code
        '        .Item("COLOR_CODE") = Color_Code
        '        .Item("VARIANCE") = QTY
        '    End With
        '    dst.Tables("ICTPHYV9").Rows.Add(rowICTPHYV9)

        'Next
        'Update_Record_TDA("ICTPHYV9")

        'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ICTPHYVV").Select(sql), "WHSE_CODE").Rows
        '    Dim Current_Whse As String = row.Item("WHSE_CODE") & ""
        '    ASCMAIN1.sql = "Update ICTSTAT1 Set WHSE_YYYYPP_LAST_PHY = '" & ASCMAIN1.CYP & "', WHSE_PHYS_STATUS = ''" & vbCrLf _
        '        & " Where STYLE_CODE  ='" & Current_Whse & "'" & vbCrLf
        '    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        'Next




    End Sub
End Class


