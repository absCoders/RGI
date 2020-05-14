Public Class ICRISTA2

    Dim ICTCOSTX As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        If Absx1.chkFor("CHKALLOCATION").Checked Then
            Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

            TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me, _
              "", _
              False, _
              True,
              True, _
              "", Now.Date, "")


            TAC.SOCMAIN1.Allocation(Me, _
                False, _
                True, _
                 "", _
                 "", New List(Of String), _
                "", "", TABLE_NAMEs, , (ROWs("SOTPARM1").Item("SO_PARM_ALLO_SEQ") & "" = "1"))

        End If

        ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = "Showing Status Qtys, and Qtys Available to Sell by Date"

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")

        Dim sql_filter2 As String = ""

        ASCMAIN1.sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_CODE" & vbCrLf _
            & ", ICTSTDQ2.DATE_1, ICTSTDQ2.QTY_1" & vbCrLf _
            & ", ICTSTDQ2.DATE_2, ICTSTDQ2.QTY_2" & vbCrLf _
            & ", ICTSTDQ2.DATE_3, ICTSTDQ2.QTY_3" & vbCrLf _
            & ", ICTSTDQ2.DATE_4, ICTSTDQ2.QTY_4" & vbCrLf _
            & ", ICTSTDQ2.DATE_5, ICTSTDQ2.QTY_5" & vbCrLf _
            & ", ICTSTDQ2.DATE_6, ICTSTDQ2.QTY_6" & vbCrLf _
            & ", ICTSTDQ2.DATE_7, ICTSTDQ2.QTY_7" & vbCrLf _
            & ", ICTSTDQ2.DATE_8, ICTSTDQ2.QTY_8" & vbCrLf _
            & ", ICTSTDQ2.DATE_9, ICTSTDQ2.QTY_9" & vbCrLf _
            & ", ICTSTDQ2.ADD_1, ICTSTDQ2.ADD_2, ICTSTDQ2.ADD_3, ICTSTDQ2.ADD_4, ICTSTDQ2.ADD_5, ICTSTDQ2.ADD_6, ICTSTDQ2.ADD_7, ICTSTDQ2.ADD_8, ICTSTDQ2.ADD_9" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_ON_ORDER " & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_TRAN, ICTSTAT2.WHSE_QTY_OPEN" & vbCrLf _
            & ", ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO" & vbCrLf _
            & " from ICTSTAT2,ICTSTDQ2" & vbCrLf _
            & sql_TABLE_NAMEs & vbCrLf _
            & " where ICTSTDQ2.STYLE_CODE (+) = ICTSTAT2.STYLE_CODE" & vbCrLf _
            & "   and ICTSTDQ2.COLOR_CODE (+) = ICTSTAT2.COLOR_CODE" & vbCrLf _
            & "   and ICTSTDQ2.WHSE_CODE (+) = ICTSTAT2.WHSE_CODE" & vbCrLf _
            & sql_WHERE & sql_JOIN & sql_filter & sql_filter2 & vbCrLf

        ' ICTSTYL1 is Always Joined from the Report Definition

        If Absx1.chkFor("CHKNEG").Checked = "1" Then
            ASCMAIN1.sql &= "   and ICTSTAT2.WHSE_QTY_ON_HAND < 0"
        End If
        If Absx1.optFor("OPTASN").Value = "S" Then
            ASCMAIN1.sql &= "   and ICTSTYL1.CUST_CODE is Null"
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            ASCMAIN1.sql &= "   and ICTSTYL1.CUST_CODE is Not Null"
        End If

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                            & " (" & G1thru9 _
                            & ", STYLE_CODE, COLOR_CODE, WHSE_CODE" _
                            & ", DATE_1, QTY_1, DATE_2, QTY_2, DATE_3, QTY_3, DATE_4, QTY_4, DATE_5, QTY_5, DATE_6, QTY_6, DATE_7, QTY_7, DATE_8, QTY_8, DATE_9, QTY_9" _
                            & ", ADD_1, ADD_2, ADD_3, ADD_4, ADD_5, ADD_6, ADD_7, ADD_8, ADD_9" _
                            & ", WHSE_ON_HAND, WHSE_ON_ORDER, WHSE_TRAN, WHSE_OPEN, WHSE_PICK, WHSE_ALLO" _
                            & ") " _
                            & " (" & ASCMAIN1.sql & ")")

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)

        If Absx1.chkFor("CHKZERO").Checked Then
            ASCMAIN1.sql = ""
            For Each COLUMN_NAME As String In New String() _
                {"WHSE_ON_HAND", "WHSE_ON_ORDER", "WHSE_TRAN", "WHSE_OPEN", "WHSE_PICK", "WHSE_ALLO"}
                ASCMAIN1.sql &= " and NVL(" & COLUMN_NAME & ",0) = 0"
            Next
            ASCMAIN1.sql = "Delete from " & TT & ASCMAIN1.SQL_Add_WHERE(ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL()
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
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_FRST_SHP DATE")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add DATE_FRST_REC DATE")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add UPC_CODE VARCHAR2(12)")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTX & " Add UPCS NUMBER (6,0)")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select ICTCOSTA.* from ICTCOSTA," & ICTCOSTX & " ICTCOSTX" & vbCrLf _
            & "   where ICTCOSTA.STYLE_CODE = ICTCOSTX.STYLE_CODE" & vbCrLf _
            & "     and ICTCOSTA.COLOR_CODE = ICTCOSTX.COLOR_CODE" & vbCrLf _
            & "     and ICTCOSTA.OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
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
            & "  Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, Min (SOTINVH1.INV_DATE) DATE_FRST_SHP" & vbCrLf _
            & "    from SOTINVH1,SOTINVH2" & vbCrLf _
            & "   where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "     and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "     and SOTINVH2.STYLE_CODE in (Select Distinct STYLE_CODE from " & ICTCOSTX & ")" & vbCrLf _
            & "     and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "'" & vbCrLf _
            & "   group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ICTCOSTX & " ICTCOSTX" & vbCrLf _
            & "    Set DATE_FRST_SHP = R1.DATE_FRST_SHP" & vbCrLf _
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

    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        ASCMAIN1.Progress("Now Loading Style Activity")

        For Each rowASTGROUP As DataRow In dst.Tables("ASTGROUP").Select("GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')")
            rowASTGROUP.Item("GROUP_CODE") = "STOCK"
            rowASTGROUP.Item("GROUP_DESC") = "Stock Item"
        Next

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_COST, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.SUB_UNIT_PACK_QTY, ICTSTYL1.INNER_PACK_QTY " _
            & " from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        ASCMAIN1.sql = "Select * from " & ICTCOSTX
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTCOSTX", 2))
        'For i As Integer = 1 To 4
        '    dst.Tables("ICTCOSTX").Columns.Add("DATE_" & Format(i, "0"), GetType(System.DateTime))
        '    dst.Tables("ICTCOSTX").Columns.Add("QTY_" & Format(i, "0"), GetType(System.Int64))
        'Next

        ASCMAIN1.sql = "Select SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, SOTRSRV1.WHSE_CODE" _
            & ", SUM (SOTRSRV2.RSRV_QTY_OPEN) RSRV_QTY_OPEN" _
            & " from SOTRSRV2,SOTRSRV1 where SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO" _
            & " group by SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE, SOTRSRV1.WHSE_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTRSRVX", 3))

    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("SUBT", txtDescription.Text & SUBT)

        CR_params.Add("NEG", IIf(Absx1.chkFor("CHKNEG").Checked, "1", "0"))
        CR_params.Add("COST", IIf(Absx1.chkFor("CHKCOST").Checked, "1", "0"))

        Generate_Report(RPT, , SUBT)

        If ASCMAIN1.CLIENT = "VAN" Then
            Prepare_Data_Extracts()
        End If


    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
            '    EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            'End If
        End If
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
                For Each C As String In New String() {"DATE_LAST_SHP", "DATE_LAST_REC", "STYLE_COST", "DATE_FRST_SHP", "DATE_FRST_REC", "UPC_CODE", "UPCS"}
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

        Set_DX_Column(grdASTEXPT1, "UPC_CODE", "UPC Code", 100, , , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "UPCS", "#UPCs", 50, "##0", , Color.Pink)

        Set_DX_Column(grdASTEXPT1, "DATE_FRST_SHP", "1st Shp (2yr)", 90, "MM/dd/yy", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "DATE_LAST_SHP", "Last Shp", 90, "MM/dd/yy", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "DATE_LAST_REC", "Last Rec", 90, "MM/dd/yy", , Color.Orange)

        Set_DX_Column(grdASTEXPT1, "STYLE_COST", "Unit Cost", 90, "#.0000", , Color.Orange)

        Set_DX_Column(grdASTEXPT1, "WHSE_ON_HAND", "On Hand", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_PICK", "In Pick", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "OTS_ONH", "OTS", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_TRAN", "In Transit", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_ON_ORDER", "On PO", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "WHSE_OPEN", "Open Orders", 90, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "NET_POS", "Avail to Sell", 90, "#,##0", , Color.LightBlue)

        For I As Integer = 1 To 4
            Set_DX_Column(grdASTEXPT1, "ADD_" & Format(I, "0"), "Fut Qty " & CStr(I), 90, "#,##0", , Color.LightGreen)
            Set_DX_Column(grdASTEXPT1, "DATE_" & Format(I, "0"), "Date " & CStr(I), 90, "MM/dd/yy", , Color.LightGreen)
        Next

        grdASTEXPT1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "STYLE_CODE")

    End Sub

End Class