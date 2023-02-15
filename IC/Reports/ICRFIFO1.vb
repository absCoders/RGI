Public Class ICRFIFO1

    Dim ICTCOSTA As String
    Dim ICTCOSTL As String
    Dim ICTCOSTU As String
    Dim ICTCOSTG As String
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -60, 0, 0)
        Get_PARM("ICTPARM1")
        Get_PARM("GLTPARM1")

        ASCMAIN1.sql = "Select * from ICTCOSTP where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
        grdICTCOSTP.DataSource = ASCDATA1.GetDataTable
        Sort_grdColumns(grdICTCOSTP, "OPS_YYYYPP".ToLower)

        If ASCMAIN1.CLIENT = "NYA" Then
            chk001.Visible = True
        End If
    End Sub

    Overrides Sub Clear_Record()
        lblWarning.Visible = False
    End Sub

    Protected Overrides Sub Build_Workfile()



        ' ALTER TABLE SOTINVH2 ADD TARIFF_UNIT_COST NUMBER (12,6);
        ' ALTER TABLE ICTCOSTL ADD TARIFF_UNIT_COST NUMBER (12,6)'

        RWU = "N"
        If chkRebuild_FIFO.Checked Then
            RWU = "R"
            ASCMAIN1.Progress("Now Calculating FIFO Costs")
            ICTCOSTA = ""
            ICTCOSTL = ""
            ICTCOSTU = ""
            TAC.ICCMAIN1.Calculate_FIFO(Me, RYP, True, ICTCOSTA, ICTCOSTL, ICTCOSTU, ICTCOSTG)
        Else

            Dim conditionalSQL As String = ""
            If ASCMAIN1.CLIENT = "NYA" Then
                If chk001.Checked Then
                    conditionalSQL = " AND ICTSTYL1.SALES_DIVISION_CODE IN (SELECT SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001')"

                    ASCMAIN1.sql = "Select * from ICTCOSTA where OPS_YYYYPP = '" & RYP & "'"
                    ICTCOSTA = ASCMAIN1.Temp_Table

                    ASCMAIN1.sql = "Select * from ICTCOSTL where OPS_YYYYPP = '" & RYP & "'"
                    ICTCOSTL = ASCMAIN1.Temp_Table

                End If
            End If

            ASCMAIN1.sql = "Select ICTCOSTL.* from " & ICTCOSTL & " ICTCOSTL,ICTSTYL1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = ICTCOSTL.STYLE_CODE" & vbCrLf _
                & conditionalSQL _
                & "   and ICTCOSTL.OPS_YYYYPP_FIFO = :PARM1"
            Create_TDA(dst.Tables.Add, "ICTCOSTL", "**", 0, False, "V")

            Fill_Records("ICTCOSTL", RYP)

            If ASCMAIN1.CLIENT = "NYA" Then
                If chk001.Checked Then
                    ASCMAIN1.sql = "Update " & ICTCOSTA & " Set LOT_AMT_ONHD = 0, LOT_AMT_USED = 0"
                    ASCDATA1.ExecuteSQL()

                    Dim DR As String = ""
                    Dim DRD As Date
                    Dim CURR_EXCH_RATE As Decimal = 0
                    For Each row As DataRow In dst.Tables("ICTCOSTL").Select("", "TRAN_DATE")
                        If Format(row.Item("TRAN_DATE"), "yyyyMMdd") <> DR Then
                            DR = Format(row.Item("TRAN_DATE"), "yyyyMMdd")
                            DRD = row.Item("TRAN_DATE")
                            CURR_EXCH_RATE = TAC.TACMAIN1.Get_Historical_Exchange_Rate("CAD", DRD)
                        End If
                        For Each C As String In New String() {"TRAN_COST", "LOT_AMT_ONHD", "LOT_AMT_USED", "LOT_AMT_SHP", "LOT_AMT_RTN", "LOT_AMT_ADJ"}
                            row.Item(C) = System.Math.Round(Val(row.Item(C) & "") / CURR_EXCH_RATE, 6)
                        Next

                        ' THIS NEXT PIECE PROBABLY NEEDS TO BE RE-ARCHITECTED IF WE GET MANY STYLES
                        ASCMAIN1.sql = "Update " & ICTCOSTA & " Set LOT_AMT_ONHD = LOT_AMT_ONHD + " & CStr(Val(row.Item("LOT_AMT_ONHD") & "")) & ", LOT_AMT_USED = LOT_AMT_USED + " & CStr(Val(row.Item("LOT_AMT_SHP") & "") + Val(row.Item("LOT_AMT_RTN") & "") + Val(row.Item("LOT_AMT_ADJ") & "")) & " where STYLE_CODE = '" & row.Item("STYLE_CODE") & "' and COLOR_CODE = '" & row.Item("COLOR_CODE") & "'"
                        ASCDATA1.ExecuteSQL("")
                    Next
                End If
            End If

            ASCMAIN1.sql = "Select ICTCOSTA.* from " & ICTCOSTA & " ICTCOSTA,ICTSTYL1" & vbCrLf _
                & "  where ICTSTYL1.STYLE_CODE = ICTCOSTA.STYLE_CODE" & vbCrLf _
                & conditionalSQL _
                & "   and ICTCOSTA.OPS_YYYYPP = :PARM1"
            Create_TDA(dst.Tables.Add, "ICTCOSTA", "**", 0, False, "V")

            Fill_Records("ICTCOSTA", RYP)
        End If


        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ' ANNA WANTS STYLES WITH NEG ON HAND TO HAVE ZERO COST
            ' the j/e is not based on these numbers

            ' BASED ON 02/13/23 CONVERSATION W/ANNA, WE NOW WILL PERMIT NEGATIVE ON HANDS TO FLOW INTO FIFO REPORT
            If False Then
                For Each rowICTCOSTA As DataRow In dst.Tables("ICTCOSTA").Select("WHSE_QTY_ON_HAND < 0")
                    Dim STYLE_CODE As String = rowICTCOSTA.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowICTCOSTA.Item("COLOR_CODE")
                    For Each rowICTCOSTL As DataRow In dst.Tables("ICTCOSTL").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                        rowICTCOSTL.Item("LOT_AMT_ONHD") = 0
                    Next
                Next
            End If


            'If ASCMAIN1.Running_in_VS Then
            '    ASCDATA1.ExecuteSQL("Truncate table WZTFIFO1")
            '    ASCMAIN1.sql = "INSERT INTO WZTFIFO1 SELECT OPS_YYYYPP, STYLE_CODE, COLOR_CODE" & vbCrLf _
            '        & ", AMT/QTY, QTY FROM (" & vbCrLf _
            '        & "SELECT OPS_YYYYPP, STYLE_CODE, COLOR_CODE" & vbCrLf _
            '        & ", SUM (LOT_AMT_ONHD) AMT, SUM(LOT_QTY_ONHD) QTY" & vbCrLf _
            '        & "FROM " & ICTCOSTL & " GROUP BY OPS_YYYYPP, STYLE_CODE, COLOR_CODE" & vbCrLf _
            '        & "HAVING SUM (LOT_QTY_ONHD) <> 0)"
            '    ASCDATA1.ExecuteSQL()
            'End If
        End If

        ASCMAIN1.sql = "Select * from ICTSTYL1"
        Create_TDA(dst.Tables.Add, "ICTSTYL1", "**", 0, False)
        Fill_Records("ICTSTYL1")

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
            & " from " & ICTCOSTA & " ICTCOSTA, ICTCLAS1, SOTSDIV1, ICTSTYL1" & vbCrLf _
            & " where ICTSTYL1.STYLE_CODE (+) = ICTCOSTA.STYLE_CODE" & vbCrLf _
            & "   and ICTCLAS1.STYLE_CLASS_CODE (+) = ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
            & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and (ICTCLAS1.STYLE_CLASS_DESC IS NULL OR SOTSDIV1.SALES_DIVISION_NAME IS NULL)" & vbCrLf _
            & "   and ICTCOSTA.LOT_AMT_ONHD <> 0"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        If tbl.Rows.Count <> 0 Then
            Using frm As New ASFMSGBF
                frm.Show_grd(tbl, Me, "Styles with Invalid or Missing Values for Division or Class")
            End Using
        End If

        If chkRebuild_FIFO.Checked Then
            ASCMAIN1.sql = "Select ICTCOSTL.*, ICTCOSTA.WHSE_QTY_ON_HAND" & vbCrLf _
                & " from " & ICTCOSTL & " ICTCOSTL, " & ICTCOSTA & " ICTCOSTA" & vbCrLf _
                & " where ICTCOSTL.TRAN_TYPE = 'Z'" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP (+) = ICTCOSTL.OPS_YYYYPP" & vbCrLf _
                & "   and ICTCOSTA.STYLE_CODE (+) = ICTCOSTL.STYLE_CODE" & vbCrLf _
                & "   and ICTCOSTA.COLOR_CODE (+) = ICTCOSTL.COLOR_CODE" & vbCrLf _
                & "   and (NVL(ICTCOSTL.LOT_QTY_USED,0) <> 0 or NVL(ICTCOSTA.WHSE_QTY_ON_HAND,0) <> 0)"
            Dim ICTFIFOZ As String = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTFIFOZ & " Add Primary Key (OPS_YYYYPP_FIFO, STYLE_CODE, COLOR_CODE, RECORD_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTFIFOZ & " Add RECEIPT_NO VARCHAR2(6)")
            ASCMAIN1.sql = "Update " & ICTFIFOZ & " ICTFIFOZ Set RECEIPT_NO = (Select MIN(RECEIPT_NO) from ICTIREC2 where STYLE_CODE = ICTFIFOZ.STYLE_CODE and COLOR_CODE = ICTFIFOZ.COLOR_CODE and OPS_YYYYPP > '" & RYP & "')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select * from " & ICTFIFOZ
            Create_TDA(dst.Tables.Add, "ICTFIFOZ", "**", 0, False)
            Fill_Records("ICTFIFOZ")

            ASCMAIN1.sql = "Select * from ICTIREC1 where RECEIPT_NO in (Select Distinct RECEIPT_NO from " & ICTFIFOZ & ")"
            Create_TDA(dst.Tables.Add, "ICTIREC1", "**", 0, False)
            Fill_Records("ICTIREC1")

        End If

        Prepare_GL_Interface("ICVA")

        ASCMAIN1.Progress("")
    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT_sfx As String = ""
        If ASCMAIN1.CLIENT = "NYA" Then
            If chk001.Checked Then
                SUBT_sfx = " - NYAG Canada (in CAD)"
            End If
        End If

        SUBT = "FIFO Calculations for " & RYPLEGEND & SUBT_sfx
        Generate_Report(RPT, , SUBT)

        If chkRebuild_FIFO.Checked Then
            SUBT = "Zero Lots Generated for Styles with Qty On Hand or Activity in " & RYPLEGEND
            Generate_Report("ICRFIFOZ", , SUBT)
        End If

        Dim RWU_pre As String = RWU
        Print_GL()
        ''If ASCMAIN1.CLIENT = "VAN" Then
        ''    If chkGL.Checked Then
        ''        Print_GL()
        ''    End If
        ''Else
        ''    Print_GL()
        ''End If

        If Not chkGL.Checked Then RWU = RWU_pre

        If ASCMAIN1.CLIENT = "VAN" Then
            'EnforceConstraints(False)
            'Dim IRECORD As Int64 = 0 ' to find orphans
            'For Each row As DataRow In dst.Tables("ICTCOSTA").Select("")
            '    Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            '    IRECORD += 1
            '    Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
            '    If rowICTSTYL1 Is Nothing Then Stop
            'Next
            Create_Relation("ICTSTYL1", "ICTCOSTA", "STYLE_CODE")
            'EnforceConstraints(True)
            dst.Tables("ICTCOSTA").Columns.Add("STYLE_DESC", GetType(System.String), "PARENT.STYLE_DESC")
            Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook() '(FILENAME)
            Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)

            Dim range As SpreadsheetGear.IRange = Nothing
            Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
            Dim rangePasteTo As SpreadsheetGear.IRange = Nothing

            worksheet.Cells(0, 0).CopyFromDataTable(dst.Tables("ICTCOSTA"), SpreadsheetGear.Data.SetDataFlags.None)

            For i As Integer = 0 To dst.Tables("ICTCOSTA").Columns.Count - 1
                Dim dcol As DataColumn = dst.Tables("ICTCOSTA").Columns(i)
                Dim T As String = dcol.DataType.ToString
                Dim F As String = ""
                If T = "System.String" Then

                ElseIf T = "System.DateTime" Then
                    F = "MM/dd/yy"
                ElseIf T = "System.Int32" Or T = "System.Int64" Or T = "System.Integer" Then
                    F = "#,##0"
                ElseIf T = "System.Decimal" Then
                    F = "#,##0.00"
                End If
                If F <> "" Then
                    worksheet.Cells(0, i).EntireColumn.NumberFormat = F
                End If

            Next

            Dim XLS_FILENAME As String = Me.Name & "_" & XNO & ".XLSX"

            workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
            workbook.Close()

            Show_Document(XLS_FILENAME)

        End If


    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            Dim RYP As String = Mid(Absx1.cmbFor("RYP").Value, 1, 4) & Mid(Absx1.cmbFor("RYP").Value, 6, 2)

            lblWarning.Visible = False
            If chkRebuild_FIFO.Checked Then
                ASCMAIN1.sql = "Select Count (*), MIN (POTSHIP1.PO_SHIPMENT_NO) PO_SHIPMENT_NO from POTSHIP2,POTSHIP1" _
                    & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" _
                    & "   and NVL(POTSHIP1.COST_COMPLETE,'0') <> '1' and POTSHIP2.OPS_YYYYPP <= '" & RYP & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If Val(row.Item(0) & "") <> 0 Then
                    lblWarning.Visible = True
                    If chkGL.Checked Then
                        MsgBox(lblWarning.Text, MsgBoxStyle.OkOnly, "You May Not Update GL with Incomplete Landed Costs Data")
                        lblWarning.Visible = False
                        EMsg &= vbCr & "Please review costs in PO Shipments Cost Entry(see Shipment " & row.Item(1) & ")"
                    Else
                        If MsgBox(lblWarning.Text & vbCrLf & vbCrLf & "Proceed Anyway?", MsgBoxStyle.YesNo, "Please Acknowledge") = MsgBoxResult.No Then
                            EMsg &= vbCr & "Please review costs in PO Shipments Cost Entry (see Shipment " & row.Item(1) & ")"
                        End If
                    End If
                End If
                If ASCMAIN1.CLIENT = "VAN" Then
                    ASCMAIN1.sql = "Select Count (*) from ICTCOSTP" _
                    & " where NVL(ICTCOSTP.UPDATED,'0') = '1' and ICTCOSTP.OPS_YYYYPP = '" & RYP & "'"
                    Dim rowICTCOSTP As DataRow = ASCDATA1.GetDataRow
                    If Val(rowICTCOSTP.Item(0) & "") <> 0 Then
                        'MsgBox(lblWarning.Text, MsgBoxStyle.OkOnly, "You May Not Update GL for this period")
                        'lblWarning.Visible = False
                        EMsg &= vbCr & "FIFO G/L Journal Has already been Updated for this period, Cannot Rebuild FIFO Lot Costs"
                    End If
                    ' CHECK TO SEE UF MONTH END HAS BEEN UPDATED FOR PERIOD
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

        ASCDATA1.ExecuteSQL("Delete from ICTCOSTL where OPS_YYYYPP_FIFO = :PARM1", "V", New Object() {RYP})
        ASCDATA1.ExecuteSQL("Insert into ICTCOSTL Select * from " & ICTCOSTL)

        ASCDATA1.ExecuteSQL("Delete from ICTCOSTA where OPS_YYYYPP = '" & RYP & "'")
        ASCDATA1.ExecuteSQL("Insert into ICTCOSTA Select * from " & ICTCOSTA)

        ASCDATA1.ExecuteSQL("Update SOTINVH2 Set ORDR_UNIT_COST = 0 where ORDR_YYYYPP_UPDATED = '" & RYP & "'")
        ASCDATA1.ExecuteSQL("Update ICTIADJ2 Set STYLE_COST = 0 where OPS_YYYYPP = '" & RYP & "'")
        'TARIFF_UNIT_COST
        ASCDATA1.ExecuteSQL("Update SOTINVH2 Set ORDR_UNIT_COST = (Select TRAN_COST from " & ICTCOSTU & " where TRAN_TYPE = SOTINVH2.INV_TYPE and TRAN_NO = SOTINVH2.INV_NO and TRAN_LNO = SOTINVH2.INV_LNO) where ORDR_YYYYPP_UPDATED = '" & RYP & "'")
        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSQL("Update SOTINVH2 Set TARIFF_UNIT_COST = (Select TARIFF_UNIT_COST from " & ICTCOSTU & " where TRAN_TYPE = SOTINVH2.INV_TYPE and TRAN_NO = SOTINVH2.INV_NO and TRAN_LNO = SOTINVH2.INV_LNO) where ORDR_YYYYPP_UPDATED = '" & RYP & "'")
            ASCDATA1.ExecuteSQL("Update SOTINVH2 Set TARIFF_FLAG = (Select TARIFF_FLAG from " & ICTCOSTU & " where TRAN_TYPE = SOTINVH2.INV_TYPE and TRAN_NO = SOTINVH2.INV_NO and TRAN_LNO = SOTINVH2.INV_LNO) where ORDR_YYYYPP_UPDATED = '" & RYP & "'")
        End If
        ASCDATA1.ExecuteSQL("Update SOTINVH1 Set INV_COGS = (Select Sum (ORDR_QTY_SHIP * ORDR_UNIT_COST) from SOTINVH2 where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO) where ORDR_YYYYPP_UPDATED = '" & RYP & "'")

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSQL("Delete from ICTSTAT4 where OPS_YYYYPP = '" & RYP & "'")
            Dim sql_I As String = "DECODE(SOTINVH2.INV_TYPE,'I',SOTINVH2.ORDR_QTY_SHIP,0)"
            Dim sql_C As String = "DECODE(SOTINVH2.INV_TYPE,'C',SOTINVH2.ORDR_QTY_SHIP,0)"
            ASCDATA1.ExecuteSQL("Insert into ICTSTAT4 Select SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.WHSE_CODE" & vbCrLf _
                                & ", SUM (" & sql_I & ") WHSE_QTY_SHP, SUM (" & sql_I & " * SOTINVH2.ORDR_UNIT_PRICE) WHSE_SLS_SHP, SUM (" & sql_I & " * SOTINVH2.ORDR_UNIT_COST) WHSE_CST_SHP" & vbCrLf _
                                & ", SUM (" & sql_C & ") WHSE_QTY_RTN, SUM (" & sql_C & " * SOTINVH2.ORDR_UNIT_PRICE) WHSE_SLS_RTN, SUM (" & sql_C & " * SOTINVH2.ORDR_UNIT_COST) WHSE_CST_RTN" & vbCrLf _
                                & " from SOTINVH2,SOTINVH1" & vbCrLf _
                                & " where SOTINVH2.ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCrLf _
                                & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                                & " group by SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.WHSE_CODE")

        End If

        ASCDATA1.ExecuteSQL("Update ICTIADJ2 Set STYLE_COST = (Select TRAN_COST from " & ICTCOSTU & " where TRAN_TYPE = 'A' and TRAN_NO = ICTIADJ2.ADJ_NO and TRAN_LNO = ICTIADJ2.ADJ_LNO) where OPS_YYYYPP = '" & RYP & "'")
        ASCDATA1.ExecuteSQL("Update ICTIADJ1 Set TOTAL_COSTS = (Select Sum (ADJ_QTY * STYLE_COST) from ICTIADJ2 where ICTIADJ2.ADJ_NO = ICTIADJ1.ADJ_NO) where OPS_YYYYPP = '" & RYP & "'")

        ASCMAIN1.sql = "Update ICTIADJ3 Set DIST_AMT = " & vbCrLf _
            & "(Select ADJ_QTY * STYLE_COST from ICTIADJ2 where ADJ_NO = ICTIADJ3.ADJ_NO and ADJ_LNO = ICTIADJ3.ADJ_LNO) " & vbCrLf _
            & " where DIST_TYPE = 'INVTY' " & vbCrLf _
            & "   and ADJ_NO IN (SELECT ADJ_NO from ICTIADJ1 where OPS_YYYYPP = '" & RYP & "')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ICTIADJ3 Set DIST_AMT = " & vbCrLf _
            & "(Select -1 * ADJ_QTY * STYLE_COST from ICTIADJ2 where ADJ_NO = ICTIADJ3.ADJ_NO and ADJ_LNO = ICTIADJ3.ADJ_LNO) " & vbCrLf _
            & " where DIST_TYPE = 'INVADJ' " & vbCrLf _
            & "   and ADJ_NO IN (SELECT ADJ_NO from ICTIADJ1 where OPS_YYYYPP = '" & RYP & "')"
        ASCDATA1.ExecuteSQL()

        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & "  Select STYLE_CODE, COLOR_CODE, TRUNC(1000000 * LOT_AMT_ONHD / LOT_QTY_ONHD)/1000000 STYLE_COST" & vbCrLf _
                & "   from ICTCOSTA where OPS_YYYYPP = '" & RYP & "' and LOT_QTY_ONHD <> 0;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTSTYC1 Set STYLE_COST_FIFO = R1.STYLE_COST" & vbCrLf _
                & "    where STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
                & "      and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & "  Select STYLE_CODE, TRUNC(1000000 * LOT_AMT_ONHD / LOT_QTY_ONHD)/1000000 STYLE_COST" & vbCrLf _
                & "   from (Select STYLE_CODE, Sum (LOT_AMT_ONHD) LOT_AMT_ONHD, Sum (LOT_QTY_ONHD) LOT_QTY_ONHD" & vbCrLf _
                & "   from ICTCOSTA where OPS_YYYYPP = '" & RYP & "' and LOT_QTY_ONHD <> 0 group by STYLE_CODE having Sum (LOT_QTY_ONHD) <> 0);" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTSTYL1 Set STYLE_COST = R1.STYLE_COST" & vbCrLf _
                & "    where STYLE_CODE = R1.STYLE_CODE;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        If chkRebuild_FIFO.Checked Then

            ASCDATA1.ExecuteSQL("Delete from ICTCOSTP where OPS_YYYYPP = '" & RYP & "'")

            ASCMAIN1.sql = "Insert into ICTCOSTP" _
                & " Select OPS_YYYYPP, COUNT (*), SUM (LOT_QTY_ONHD), SUM (LOT_AMT_ONHD), NULL" _
                & ", SYSDATE, '" & ASCMAIN1.USER_ID & "', NULL, NULL" _
                & " from ICTCOSTA where OPS_YYYYPP = '" & RYP & "' group by OPS_YYYYPP"
            ASCDATA1.ExecuteSQL()

        End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is" & vbCrLf _
                & "Select * from (Select X.*" & vbCrLf _
                & ", CASE WHEN X.LOT_QTY_ONHD <> 0 THEN TRUNC(1000000 * X.LOT_AMT_ONHD / X.LOT_QTY_ONHD) / 1000000 ELSE" & vbCrLf _
                & "  CASE WHEN X.LOT_QTY_USED <> 0 THEN TRUNC(1000000 * X.LOT_AMT_USED / X.LOT_QTY_USED) / 1000000 ELSE" & vbCrLf _
                & " X.STYLE_COST END END STYLE_COST_CALC" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select STYLE_CODE, OPS_YYYYPP, MAX(STYLE_COST) STYLE_COST" & vbCrLf _
                & ", SUM (LOT_QTY_ONHD) LOT_QTY_ONHD, SUM (LOT_AMT_ONHD) LOT_AMT_ONHD" & vbCrLf _
                & ", SUM (LOT_QTY_USED) LOT_QTY_USED, SUM (LOT_AMT_USED) LOT_AMT_USED" & vbCrLf _
                & " from ICTCOSTA where (STYLE_CODE, OPS_YYYYPP) IN (" & vbCrLf _
                & "Select STYLE_CODE, MAX(OPS_YYYYPP) OPS_YYYYPP FROM ICTCOSTA" & vbCrLf _
                & " where (NVL(WHSE_QTY_ON_HAND,0) <> 0 OR NVL(LOT_QTY_USED,0) <> 0)" & vbCrLf _
                & " group by STYLE_CODE)" & vbCrLf _
                & " group by STYLE_CODE, OPS_YYYYPP) X) where NVL(STYLE_COST_CALC,0) <> 0;" & vbCrLf _
                & "Begin For R1 in C1 Loop" & vbCrLf _
                & " Update ICTSTYL1 Set STYLE_COST = R1.STYLE_COST_CALC" & vbCrLf _
                & "  where STYLE_CODE = R1.STYLE_CODE;" & vbCrLf _
                & "End Loop; End; End;"
            ASCDATA1.ExecuteSQL()

        End If

        ' UN REM VAN CHECK BELOW WHEN GOING LIVE
        If chkGL.Checked Then
            If (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") Then
            Else
                GL_Update()
            End If

            ASCMAIN1.sql = "Update ICTCOSTP" _
                & " Set UPDATED = '1', LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                & " where OPS_YYYYPP = '" & RYP & "'"
            ASCDATA1.ExecuteSQL()
        End If

    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        If Not dst.Tables.Contains("GLTINTF1") Then
            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End If

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))

        Dim NYP As String = ASCMAIN1.Period_Calc(RYP, 1)

        ASCMAIN1.sql = "Select ICTCLAS1.ACCT_CODE_ONH, ICTSTYL1.SALES_DIVISION_CODE, SUM (ICTCOSTA.LOT_AMT_ONHD) LOT_AMT_ONHD " & vbCrLf _
         & " from " & ICTCOSTA & " ICTCOSTA, ICTCLAS1, ICTSTYL1" & vbCrLf _
         & " where ICTSTYL1.STYLE_CODE (+) = ICTCOSTA.STYLE_CODE" _
         & "   and ICTCLAS1.STYLE_CLASS_CODE (+) = ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
         & "   and ICTCOSTA.LOT_AMT_ONHD <> 0" & vbCrLf _
         & "   and ICTCOSTA.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
         & IIf(chk001.Checked, " AND ICTSTYL1.SALES_DIVISION_CODE IN (SELECT SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001')", "") _
         & " group by ICTCLAS1.ACCT_CODE_ONH, ICTSTYL1.SALES_DIVISION_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            Dim DETL_POSTING_AMT As Decimal = Val(row.Item("LOT_AMT_ONHD") & "")

            Dim DETL_CVX_NO As String = ""
            Dim DETL_CVX_REF_NO As String = row.Item("SALES_DIVISION_CODE") & ""
            Dim DETL_CVX_TYPE As String = ""

            Dim SEG2_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            Dim SEG3_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            Dim SEG4_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            For I As Integer = 1 To 2
                Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                rowGLTINTF1("OPS_YYYYPP") = RYP
                rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                JOURNAL_LNO += 1
                rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO

                Dim ACCT_CODE As String = ""
                If I = 1 Then
                    ' ACCT_CODE = row.Item("ACCT_CODE_ONH") & "" - use this if BS inventory is specific to class
                    ACCT_CODE = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_INVTY") & ""
                    SEG3_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") ' row.Item("SALES_DIVISION_CODE") & ""
                Else
                    DETL_POSTING_AMT = -1 * DETL_POSTING_AMT
                    ACCT_CODE = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_INVTY_END") & ""
                    SEG3_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") 'row.Item("SALES_DIVISION_CODE") & ""
                End If

                rowGLTINTF1("ACCT_CODE") = ACCT_CODE
                rowGLTINTF1("SEG2_CODE") = SEG2_CODE
                rowGLTINTF1("SEG3_CODE") = SEG3_CODE
                rowGLTINTF1("SEG4_CODE") = SEG4_CODE
                rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
                rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
                rowGLTINTF1("DETL_EXE_NO") = XNO
                rowGLTINTF1("DETL_CTL_NO") = DBNull.Value
                rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
                rowGLTINTF1("DETL_CVX_NO") = DETL_CVX_NO
                rowGLTINTF1("DETL_CVX_REF_DATE") = DETL_CTL_DATE
                rowGLTINTF1("DETL_CVX_REF_NO") = DETL_CVX_REF_NO
                rowGLTINTF1("DETL_DESC") = DBNull.Value
                rowGLTINTF1("DETL_CVX_TYPE") = DETL_CVX_TYPE
                rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)

                Dim rowGLTINTF1_reversal As DataRow = dst.Tables("GLTINTF1").NewRow
                rowGLTINTF1_reversal.ItemArray = rowGLTINTF1.ItemArray
                rowGLTINTF1_reversal.Item("OPS_YYYYPP") = NYP
                rowGLTINTF1_reversal.Item("DETL_POSTING_AMT") = -1 * DETL_POSTING_AMT
                If I = 2 Then
                    If NYP.EndsWith("01") Then
                        ACCT_CODE = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_INVTY_BEG") & ""
                    Else
                        ACCT_CODE = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_INVTY_END") & ""
                    End If
                    rowGLTINTF1_reversal.Item("ACCT_CODE") = ACCT_CODE
                End If
                dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1_reversal)
            Next
        Next

        Return JOURNAL_NO
    End Function

    Private Sub chkRebuild_FIFO_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkRebuild_FIFO.CheckedChanged
        If chkRebuild_FIFO.Checked Then
            If ASCMAIN1.CLIENT = "VAN" Then
                chkGL.Visible = (ASCMAIN1.USER_ID = "dgj" Or ASCMAIN1.USER_ID = "anna" Or ASCMAIN1.USER_ID = "wjz")
            Else
                chkGL.Visible = True
            End If
        Else
            chkGL.Visible = False
            chkGL.Checked = False
        End If

        chk001.Visible = Not chkRebuild_FIFO.Checked And ASCMAIN1.CLIENT = "NYA"
        If Not chk001.Visible Then
            chk001.Checked = False
        End If
    End Sub


#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTCOSTP, "B", "Re-Open Last Period Closed")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.Name = "grd" Then
            Exit Sub
        End If

        Select Case e.SourceControl.Name
         
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                'Case "Show Style/Colors w/Zero Status"
                '    tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Re-Open Last Period Closed"
                Dim RYP_last As String = ""
                Dim tbl As DataTable = DirectCast(grdICTCOSTP.DataSource, DataTable)
                RYP_last = tbl.Compute("MAX(OPS_YYYYPP)", "")

                If RYP_last = "" Then
                    MsgBox("No Periods Closed - nothing to re-open", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If
                If MsgBox("OK to Re-Open Period " & RYP_last & " for Costing?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                Else
                    Dim row As DataRow = tbl.Rows.Find(RYP_last)
                    row.Delete()
                    BeginTrans()
                    Write_Audit_Trail(row, "D")
                    ASCDATA1.ExecuteSQL("Delete from ICTCOSTP where OPS_YYYYPP = '" & RYP_last & "'")
                    CommitTrans("Period " & RYP_last & " has been Successfully Re-Opened")
                End If

            Case "Show Style/Colors w/Zero Status"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then Exit Sub
        End Select

        If grd Is Nothing OrElse grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub
#End Region
End Class