Public Class SAFSLSI1

    Dim RYP As String
    Dim SATSLSI1 As String
    Dim ICTSTKL2 As String
    Dim SATSLSI0 As String
    Dim TBL_PLAN As DataTable

    Dim MAXPOSQTY As Integer = 4
    Dim MINPOSQTY As Integer = 2
    Dim MAXPOSCST As Integer = 4
    Dim MINPOSCST As Integer = 2

    'anna inventory position report uses ICTCOSTL so on hand and PO positions should be built right after fifo update

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            Create_SATSLSI1("")
            ASCMAIN1.sql = "Select * from " & SATSLSI1
            Create_TDA(.Tables.Add, "SATSLSI1", "**", 0, False, "", 0)

            Create_TDA(.Tables.Add, "SOTSDIV1", "*", 0, False)
            Create_TDA(.Tables.Add, "ICTPCAT1", "*", 0, False)
        End With

        Fill_Records("SOTSDIV1")
        Fill_Records("ICTPCAT1")

        grdSATSLSI1.DataSource = dst.Tables("SATSLSI1")

        Create_Summary(grdSATSLSI1, "STYLE_CODE", "Count")

 
        With grdSATSLSI1.DisplayLayout.Bands("SATSLSI1")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
 
        End With

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP
 
        Show_Filter(grdSATSLSI1, True)
        grdSATSLSI1.DisplayLayout.GroupByBox.Hidden = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("OPS_YYYYPP")
                Dim rowICTCOSTP As DataRow = LookUp("ICTCOSTP", ASCMAIN1.CYP)
                If rowICTCOSTP Is Nothing Then
                    EMsg &= vbCr & "You must run and update the FIFO report before running this screen"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdPLAN.Visible = tf
        btnPlusDuty.Visible = tf

        grdSATSLSI1.Visible = tf
        Absx1.txtFor("OPS_YYYYPP").ReadOnly = True

        If ScreenMode Then
            grdSATSLSI1.Text = "Inventory Position Analysis for " & Absx1.txtFor("LEGEND").Text
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("SATSLSI1").Rows.Clear()
        dst.EnforceConstraints = True
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)




        ' Prepare Data for Inventory Position Analysis

        Dim z As String = Absx1.txtFor("OPS_YYYYPP").Text
        RYP = z
        Create_SATSLSI1(RYP)

        ASCMAIN1.Progress("Now Loading Data from Database")

        dst.EnforceConstraints = False

        Fill_Records("SATSLSI1")

        grdSATSLSI1.Rows.ExpandAll(True)

        EnforceConstraints(True)

        '       Load_XLS()
        Create_Pivot()

        If RYP = ASCMAIN1.CYP Then
            ICTSTKL2 = TAC.POCMAIN1.Aged_PO(Me, ASCMAIN1.CYP)
        Else
            ICTSTKL2 = "ICTSTKL2"
        End If





        ' Prepare Data for Planning Report

        Dim sql1 As String = ""
        Dim sql1_zero As String = ""
        For I As Integer = 0 To 24
            Dim YP As String = ASCMAIN1.Period_Calc(RYP, -1 * I)
            Dim sql1x As String = ", SUM (DECODE(ICTSTAT1.OPS_YYYYPP,'" & YP & "',ICTSTAT1."
            Dim sql4x As String = ", SUM (DECODE(ICTSTAT1.OPS_YYYYPP,'" & YP & "',ICTSTAT4."

            Dim Ixx As String = "H" & Format(I, "00")

            sql1 &= sql1x & "WHSE_QTY_BEG,0)) BOM_QTY_" & Ixx
            sql1 &= sql1x & "WHSE_QTY_BEG,0) * ICTCOSTA.STYLE_COST) BOM_CST_" & Ixx
            sql1 &= sql4x & "WHSE_QTY_SHP,0)) SHP_QTY_" & Ixx
            sql1 &= sql4x & "WHSE_CST_SHP,0)) SHP_CST_" & Ixx
            sql1 &= sql4x & "WHSE_SLS_SHP,0)) SHP_SLS_" & Ixx
            sql1 &= sql4x & "WHSE_QTY_RTN,0)) RTN_QTY_" & Ixx
            sql1 &= sql4x & "WHSE_CST_RTN,0)) RTN_CST_" & Ixx
            sql1 &= sql4x & "WHSE_SLS_RTN,0)) RTN_SLS_" & Ixx
            sql1 &= sql1x & "WHSE_QTY_REC,0)) REC_QTY_" & Ixx
            sql1 &= sql1x & "WHSE_QTY_REC,0) * ICTCOSTA.STYLE_COST) REC_CST_" & Ixx
            sql1 &= sql1x & "WHSE_QTY_ADJ,0)) ADJ_QTY_" & Ixx
            sql1 &= sql1x & "WHSE_QTY_ADJ,0) * ICTCOSTA.STYLE_COST) ADJ_CST_" & Ixx
            sql1_zero &= ", 0 BOM_QTY_" & Ixx
            sql1_zero &= ", 0 BOM_CST_" & Ixx
            sql1_zero &= ", 0 SHP_QTY_" & Ixx
            sql1_zero &= ", 0 SHP_CST_" & Ixx
            sql1_zero &= ", 0 SHP_SLS_" & Ixx
            sql1_zero &= ", 0 RTN_QTY_" & Ixx
            sql1_zero &= ", 0 RTN_CST_" & Ixx
            sql1_zero &= ", 0 RTN_SLS_" & Ixx
            sql1_zero &= ", 0 REC_QTY_" & Ixx
            sql1_zero &= ", 0 REC_CST_" & Ixx
            sql1_zero &= ", 0 ADJ_QTY_" & Ixx
            sql1_zero &= ", 0 ADJ_CST_" & Ixx
        Next



 

        Dim sql2 As String = ""
        Dim sql2_zero As String = ""

        sql2 = "" _
            & ", SUM (ICTSTKL2.PO_QTY_XIT_CMO) PO_QTY_XIT_F00" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_XIT_NMO) PO_QTY_XIT_F01" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_XIT_2NMO) PO_QTY_XIT_F02" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_OPN_CMO) PO_QTY_OPN_F00" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_OPN_NMO) PO_QTY_OPN_F01" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_OPN_2NMO) PO_QTY_OPN_F02" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_OPN_3NMO) PO_QTY_OPN_F03" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_OPN_4NMO) PO_QTY_OPN_F04" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_OPN_5NMO) PO_QTY_OPN_F05" & vbCrLf _
            & ", SUM (ICTSTKL2.PO_QTY_OPN_6NMO) PO_QTY_OPN_F06" & vbCrLf
        sql2 &= Replace(sql2, "PO_QTY", "PO_CST")

        sql2_zero &= ", 0 PO_QTY_XIT_F00"
        sql2_zero &= ", 0 PO_QTY_XIT_F01"
        sql2_zero &= ", 0 PO_QTY_XIT_F02"
        sql2_zero &= ", 0 PO_QTY_OPN_F00"
        sql2_zero &= ", 0 PO_QTY_OPN_F01"
        sql2_zero &= ", 0 PO_QTY_OPN_F02"
        sql2_zero &= ", 0 PO_QTY_OPN_F03"
        sql2_zero &= ", 0 PO_QTY_OPN_F04"
        sql2_zero &= ", 0 PO_QTY_OPN_F05"
        sql2_zero &= ", 0 PO_QTY_OPN_F06"
        sql2_zero &= Replace(sql2_zero, "PO_QTY", "PO_CST")


        Dim sql3 As String = ""
        Dim sql3_zero As String = ""

        For I As Integer = -3 To 12
            Dim YP_FC As String = ASCMAIN1.Period_Calc(RYP, I)
            Dim C As String = "F" & Format(I, "00")
            Dim YP As String = RYP
            If I < 0 Then
                YP = YP_FC
                C = "H" & Format(-1 * I, "00")
            End If

            Dim sql3x As String = ", SUM (CASE WHEN ICTSTKL0.OPS_YYYYPP = '" & YP & "' AND ICTSTKL0.OPS_YYYYPP_FC = '" & YP_FC & "' THEN ICTSTKL0.QTY_PROJ ELSE 0 END"
            sql3 &= sql3x & ") FC_QTY_" & C & vbCrLf
            sql3 &= sql3x & " * ICTSTKL0.AVG_PRICE) FC_SLS_" & C & vbCrLf
            sql3 &= sql3x & " * ICTSTKL0.AVG_COST) FC_CST_" & C & vbCrLf

            sql3_zero &= ", 0 FC_QTY_" & C
            sql3_zero &= ", 0 FC_SLS_" & C
            sql3_zero &= ", 0 FC_CST_" & C
        Next


        Dim sqlStyle As String = ", ICTSTYL1.SALES_DIVISION_CODE, ICTBODY2.CATGY_CODE, DECODE(ICTSTYL1.CUST_CODE,NULL,'S','N') STK, ICTSTYL1.CUST_CODE"

        ASCMAIN1.sql = "Select ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE" & vbCrLf _
            & sqlStyle & vbCrLf _
            & sql1 & vbCrLf _
            & sql2_zero & vbCrLf _
            & sql3_zero & vbCrLf _
            & " from ICTSTAT1,ICTSTAT4,ICTCOSTA,ICTSTYL1,ICTBODY2" & vbCrLf _
            & " where ICTSTAT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -24) & "' and ICTSTAT1.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP (+) = ICTSTAT1.OPS_YYYYPP" & vbCrLf _
            & "   and ICTCOSTA.STYLE_CODE (+) = ICTSTAT1.STYLE_CODE" & vbCrLf _
            & "   and ICTCOSTA.COLOR_CODE (+) = ICTSTAT1.COLOR_CODE" & vbCrLf _
            & "   and ICTSTAT4.OPS_YYYYPP (+) = ICTSTAT1.OPS_YYYYPP" & vbCrLf _
            & "   and ICTSTAT4.STYLE_CODE (+) = ICTSTAT1.STYLE_CODE" & vbCrLf _
            & "   and ICTSTAT4.COLOR_CODE (+) = ICTSTAT1.COLOR_CODE" & vbCrLf _
            & "   and ICTSTAT4.WHSE_CODE (+) = ICTSTAT1.WHSE_CODE" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = ICTSTAT1.STYLE_CODE" & vbCrLf _
            & "   and ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf _
            & " group by ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE" & vbCrLf _
            & Replace(sqlStyle, ") STK,", "),")

        ASCMAIN1.sql &= vbCrLf & " union " & vbCrLf _
            & "Select ICTSTKL2.STYLE_CODE, ICTSTKL2.COLOR_CODE" & vbCrLf _
                        & sqlStyle & vbCrLf _
            & sql1_zero & vbCrLf _
            & sql2 & vbCrLf _
            & sql3_zero & vbCrLf _
            & " from " & ICTSTKL2 & " ICTSTKL2,ICTSTYL1,ICTBODY2" & vbCrLf _
            & " where ICTSTKL2.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = ICTSTKL2.STYLE_CODE" & vbCrLf _
            & "   and ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf _
            & " group by ICTSTKL2.STYLE_CODE, ICTSTKL2.COLOR_CODE" & vbCrLf _
            & Replace(sqlStyle, ") STK,", "),")

        ASCMAIN1.sql &= vbCrLf & " union " & vbCrLf _
            & "Select NULL STYLE_CODE, NULL COLOR_CODE" & vbCrLf _
            & ", ICTSTKL0.SALES_DIVISION_CODE, ICTSTKL0.CATGY_CODE, ICTSTKL0.STK, ICTSTKL0.CUST_CODE" & vbCrLf _
            & sql1_zero & vbCrLf _
            & sql2_zero & vbCrLf _
            & sql3 & vbCrLf _
            & " from ICTSTKL0" & vbCrLf _
            & " where ICTSTKL0.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -3) & "' and ICTSTKL0.OPS_YYYYPP <= '" & RYP & "'" & vbCrLf _
            & " group by ICTSTKL0.SALES_DIVISION_CODE, ICTSTKL0.CATGY_CODE, ICTSTKL0.STK, ICTSTKL0.CUST_CODE"

        SATSLSI1 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Select * from " & SATSLSI1 & " where ROWNUM < 1"
        Dim TBL As DataTable = ASCDATA1.GetDataTable


        Dim sqlSums As String = ""
        For Each DC As DataColumn In TBL.Columns
            Dim C As String = DC.ColumnName
            If DC.DataType.ToString <> "System.String" Then
                sqlSums &= ", Sum (" & C & ")" & C & vbCrLf
            End If
        Next

        ASCMAIN1.sql = "Select NVL(SALES_DIVISION_CODE,'BB') SALES_DIVISION_CODE, NVL(CATGY_CODE,'BLANK') CATGY_CODE, STK" & vbCrLf _
            & sqlSums _
            & " from " & SATSLSI1 & vbCrLf _
            & " group by NVL(SALES_DIVISION_CODE,'BB'), NVL(CATGY_CODE,'BLANK'), STK"

        SATSLSI0 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Select * from " & SATSLSI0
        TBL_PLAN = ASCDATA1.GetDataTable





        With TBL_PLAN.Columns
            For Each DT As String In New String() {"Q", "C"}

                For I As Integer = 1 To 3

                    Dim F_EOM As String = ""
                    For J As Integer = 1 To 12
                        F_EOM &= "+ISNULL(BOM_QTY_H" & Format(I - 1 + J - 1, "00") & ",0)"
                    Next
                    Dim C_EOM As String = "T" & DT & "EOM_" & Format(I, "00")
                    .Add(C_EOM, GetType(System.Decimal), "(" & Mid(F_EOM, 2) & ")/12")

                    Dim F_SHP As String = ""
                    For J As Integer = 1 To 12
                        F_SHP &= "+ISNULL(SHP_QTY_H" & Format(I - 1 + J - 0, "00") & ",0)"
                    Next
                    Dim C_SHP As String = "T" & DT & "SHP_" & Format(I, "00")
                    .Add(C_SHP, GetType(System.Decimal), Mid(F_SHP, 2))

                    Dim F As String = "IIF(" & C_EOM & " = 0,0," & C_SHP & "/" & C_EOM & ")"
                    Dim C As String = "T" & DT & Format(I, "00")
                    .Add(C, GetType(System.Decimal), F)

                Next
                For I As Integer = 0 To 6
                    Dim C As String = "P" & DT & Format(I, "00")
                    .Add(C, GetType(System.Decimal))
                Next
                For I As Integer = 0 To 6
                    Dim C As String = "OS" & DT & Format(I, "00")
                    .Add(C, GetType(System.Decimal))
                Next
            Next
        End With

        'Dim MAXPOSQTY As Integer = 4
        'Dim MINPOSQTY As Integer = 2
        'Dim MAXPOSCST As Integer = 4
        'Dim MINPOSCST As Integer = 2

        For Each row As DataRow In TBL_PLAN.Select("")

            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE") & ""
            Dim CATGY_CODE As String = row.Item("CATGY_CODE") & ""
            Dim STK As String = row.Item("STK") & ""

            Dim FCQ(12) As Int64
            Dim FCC(12) As Decimal
            Dim BOMQ(6) As Int64
            Dim BOMC(6) As Decimal
            For I As Integer = 0 To 12
                Dim Ixx As String = Format(I, "00")
                FCQ(I) = Val(row.Item("FC_QTY_F" & Ixx) & "")
                FCC(I) = Val(row.Item("FC_CST_F" & Ixx) & "")
                If I <= 6 Then
                    BOMQ(I) = Val(row.Item("FC_QTY_F" & Ixx) & "")
                    BOMC(I) = Val(row.Item("FC_CST_F" & Ixx) & "")
                End If
            Next

            Dim EOMQ(6) As Int64
            Dim FQ0 As Int64 = Val(row.Item("FC_QTY_F00") & "")
            If Val(row.Item("SHP_QTY_H00") & "") > FQ0 Then
                FQ0 = Val(row.Item("SHP_QTY_H00") & "")
            End If
            EOMQ(0) = Val(row.Item("BOM_QTY_H00") & "") - FQ0 + Val(row.Item("PO_QTY_OPN_F00") & "") + Val(row.Item("PO_QTY_XIT_F00") & "") + Val(row.Item("REC_QTY_H00") & "") + Val(row.Item("ADJ_QTY_H00") & "")
            EOMQ(1) = EOMQ(0) - Val(row.Item("FC_QTY_F01") & "") + Val(row.Item("PO_QTY_OPN_F01") & "") + Val(row.Item("PO_QTY_XIT_F01") & "")
            EOMQ(2) = EOMQ(1) - Val(row.Item("FC_QTY_F02") & "") + Val(row.Item("PO_QTY_OPN_F02") & "") + Val(row.Item("PO_QTY_XIT_F02") & "")
            EOMQ(3) = EOMQ(2) - Val(row.Item("FC_QTY_F03") & "") + Val(row.Item("PO_QTY_OPN_F03") & "")
            EOMQ(4) = EOMQ(3) - Val(row.Item("FC_QTY_F04") & "") + Val(row.Item("PO_QTY_OPN_F04") & "")
            EOMQ(5) = EOMQ(4) - Val(row.Item("FC_QTY_F05") & "") + Val(row.Item("PO_QTY_OPN_F05") & "")
            EOMQ(6) = EOMQ(5) - Val(row.Item("FC_QTY_F06") & "") + Val(row.Item("PO_QTY_OPN_F06") & "")
            Dim EOMC(6) As Decimal
            Dim FC0 As Int64 = Val(row.Item("FC_CST_F00") & "")
            If Val(row.Item("SHP_CST_H00") & "") > FQ0 Then
                FC0 = Val(row.Item("SHP_CST_H00") & "")
            End If
            EOMC(0) = Val(row.Item("BOM_CST_H00") & "") - FC0 + Val(row.Item("PO_CST_OPN_F00") & "") + Val(row.Item("PO_CST_XIT_F00") & "") + Val(row.Item("REC_CST_H00") & "") + Val(row.Item("ADJ_CST_H00") & "")
            EOMC(1) = EOMC(0) - Val(row.Item("FC_CST_F01") & "") + Val(row.Item("PO_CST_OPN_F01") & "") + Val(row.Item("PO_CST_XIT_F01") & "")
            EOMC(2) = EOMC(1) - Val(row.Item("FC_CST_F02") & "") + Val(row.Item("PO_CST_OPN_F02") & "") + Val(row.Item("PO_CST_XIT_F02") & "")
            EOMC(3) = EOMC(2) - Val(row.Item("FC_CST_F03") & "") + Val(row.Item("PO_CST_OPN_F03") & "")
            EOMC(4) = EOMC(3) - Val(row.Item("FC_CST_F04") & "") + Val(row.Item("PO_CST_OPN_F04") & "")
            EOMC(5) = EOMC(4) - Val(row.Item("FC_CST_F05") & "") + Val(row.Item("PO_CST_OPN_F05") & "")
            EOMC(6) = EOMC(5) - Val(row.Item("FC_CST_F06") & "") + Val(row.Item("PO_CST_OPN_F06") & "")

            Dim C As String = ""
            For I As Integer = 0 To 6

                Dim DQ As Int64 = 0
                Dim PQ As Decimal = 0
                Dim OSQ As Int64 = 0
                For J As Integer = I + 1 To 12
                    If FCQ(J) > 0 Then
                        If EOMQ(I) > DQ + FCQ(J) Then ' IF I = 0 THEN WE NEED TO USE MAX OF SLS AND FC
                            PQ += 1
                            DQ += FCQ(J)
                        Else
                            PQ += (EOMQ(I) - DQ) / FCQ(J)
                            Exit For
                        End If
                    End If
                Next

                If PQ < 0 Then PQ = 0

                Dim MAXQTY As Int64 = 0
                Dim MINQTY As Int64 = 0
                For J As Integer = 1 To MAXPOSQTY
                    If I + J > 12 Then Exit For
                    If J <= MINPOSQTY Then
                        MINQTY += Val(row.Item("FC_QTY_F" & Format(I + J, "00")) & "")
                    End If
                    MAXQTY += Val(row.Item("FC_QTY_F" & Format(I + J, "00")) & "")
                Next
                If EOMQ(I) > MAXQTY Then
                    OSQ = MAXQTY - EOMQ(I)
                ElseIf EOMQ(I) < MINQTY Then
                    OSQ = MINQTY - EOMQ(I)
                End If
                row.Item("PQ" & Format(I, "00")) = PQ
                row.Item("OSQ" & Format(I, "00")) = OSQ

                ' If SALES_DIVISION_CODE = "15" And CATGY_CODE = "BRAS" And STK = "N" Then Stop


                Dim DC As Decimal = 0
                Dim PC As Decimal = 0
                Dim OSC As Decimal = 0
                For J As Integer = I + 1 To 12
                    If FCC(J) > 0 Then
                        If EOMC(I) > DC + FCC(J) Then ' IF I = 0 THEN WE NEED TO USE MAX OF SLS AND FC
                            PC += 1
                            DC += FCC(J)
                        Else
                            PC += (EOMC(I) - DC) / FCC(J)
                            Exit For
                        End If
                    End If
                Next

                If PC < 0 Then PC = 0

                Dim MAXCST As Int64 = 0
                Dim MINCST As Int64 = 0
                For J As Integer = 1 To MAXPOSCST
                    If I + J > 12 Then Exit For
                    If J <= MINPOSCST Then
                        MINCST += Val(row.Item("FC_CST_F" & Format(I + J, "00")) & "")
                    End If
                    MAXCST += Val(row.Item("FC_CST_F" & Format(I + J, "00")) & "")
                Next
                If EOMC(I) > MAXCST Then
                    OSC = MAXCST - EOMC(I)
                ElseIf EOMC(I) < MINCST Then
                    OSC = MINCST - EOMC(I)
                End If
                row.Item("PC" & Format(I, "00")) = PC
                row.Item("OSC" & Format(I, "00")) = OSC

                '  If SALES_DIVISION_CODE = "01" And CATGY_CODE = "BRAS" And STK = "N" Then Stop
            Next
        Next



        grdPLAN.DataSource = TBL_PLAN
        Show_Filter(grdPLAN, True)
        Sort_grdColumns(grdPLAN, "SALES_DIVISION_CODE,CATGY_CODE,STK")
        '   Generate_Plan_XLS()







        ' Prepare Gross Profit & Aged Inventory Summary 







        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSATSLSI1, "SSBSB", "Show Filter", "Show GroupBox", "Item Status Inquiry", "Show Inactive", "Restore Original Sort")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = Nothing
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
         
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Sub Create_SATSLSI1(ByVal RYP As String)

        Dim DTES(13) As String
        For I As Integer = 0 To 13
            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, I))
            DTES(I) = Format(rowGLTPARM2.Item("PRD_END_DATE"), "dd-MMM-yyyy")
        Next

        '& ", OPNCURMO * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPOCURMO" & vbCrLf _
        '& ", OPNNXTMO * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPONXTMO" & vbCrLf _
        '& ", OPN2NXTMO * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO2NXTMO" & vbCrLf _
        '& ", OPN3NXTMO * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO3NXTMO" & vbCrLf _
        '& ", OPN4NXTMO * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO4NXTMO" & vbCrLf _

        ASCMAIN1.sql = "" _
            & "Select Q.*" & vbCrLf _
            & ", (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) COST_USED" & vbCrLf _
            & ", PICK * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_PICK" & vbCrLf _
            & ", OPEN * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_OPEN" & vbCrLf _
            & ", ONPO * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO" & vbCrLf _
            & ", TRAN * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_TRAN" & vbCrLf _
            & ", OPN_M00 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M00" & vbCrLf _
            & ", OPN_M01 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M01" & vbCrLf _
            & ", OPN_M02 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M02" & vbCrLf _
            & ", OPN_M03 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M03" & vbCrLf _
            & ", OPN_M04 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M04" & vbCrLf _
            & ", OPN_M05 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M05" & vbCrLf _
            & ", OPN_M06 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M06" & vbCrLf _
            & ", OPN_M07 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M07" & vbCrLf _
            & ", OPN_M08 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M08" & vbCrLf _
            & ", OPN_M09 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M09" & vbCrLf _
            & ", OPN_M10 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M10" & vbCrLf _
            & ", OPN_M11 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M11" & vbCrLf _
            & ", OPN_M12 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M12" & vbCrLf _
            & ", OPN_M13 * (CASE WHEN NVL(COST,0) = 0 THEN NVL(POCOST1,0) ELSE NVL(COST,0) END) AMT_ONPO_M13" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select Z.*, CASE WHEN NVL(Z.QTY,0) = 0 THEN 0 ELSE ROUND(Z.AMT/Z.QTY,6) END COST" & vbCrLf _
            & ", P.OPN_M00, P.OPN_M01, P.OPN_M02, P.OPN_M03, P.OPN_M04, P.OPN_M05, P.OPN_M06, P.OPN_M07, P.OPN_M08, P.OPN_M09, P.OPN_M10, P.OPN_M11, P.OPN_M12, P.OPN_M13" & vbCrLf _
            & ", ICTSTYL1.CUST_CODE, DECODE(ICTSTYL1.CUST_CODE,NULL,'STOCK','NONSTOCK') STK, ICTSTYL1.SUB_BODY_CODE, ICTBODY2.MASTER_BODY_CODE, ICTSTYL1.SALES_DIVISION_CODE" & vbCrLf _
            & " from ICTBODY2,ICTSTYL1," & vbCrLf _
            & "(Select STYLE_CODE, COLOR_CODE, SUM (PO_QTY_OPN) QTY_OPN" & vbCrLf _
            & ", SUM (CASE WHEN                                     PO_DATE_ETA <= '" & DTES(0) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M00" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(0) & "' and PO_DATE_ETA <= '" & DTES(1) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M01" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(1) & "' and PO_DATE_ETA <= '" & DTES(2) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M02" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(2) & "' and PO_DATE_ETA <= '" & DTES(3) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M03" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(3) & "' and PO_DATE_ETA <= '" & DTES(4) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M04" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(4) & "' and PO_DATE_ETA <= '" & DTES(5) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M05" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(5) & "' and PO_DATE_ETA <= '" & DTES(6) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M06" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(6) & "' and PO_DATE_ETA <= '" & DTES(7) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M07" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(7) & "' and PO_DATE_ETA <= '" & DTES(8) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M08" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(8) & "' and PO_DATE_ETA <= '" & DTES(9) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M09" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(9) & "' and PO_DATE_ETA <= '" & DTES(10) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M10" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(10) & "' and PO_DATE_ETA <= '" & DTES(11) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M11" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(11) & "' and PO_DATE_ETA <= '" & DTES(12) & "' THEN PO_QTY_OPN ELSE 0 END) OPN_M12" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(12) & "'                                      THEN PO_QTY_OPN ELSE 0 END) OPN_M13" & vbCrLf _
            & " from POTORDR2 WHERE PO_STATUS = 'O' AND PO_QTY_OPN <> 0 GROUP BY STYLE_CODE, COLOR_CODE) P," & vbCrLf _
            & "(" & vbCrLf _
            & "Select X.*, Y.POCOST1, Y.OPN, Y.SHP, ICTSTYC1.STYLE_COST_FIFO" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", SUM (QTY) QTY, SUM (AMT) AMT" & vbCrLf _
            & ", SUM (QTYCUR) QTYCUR, SUM (AMTCUR) AMTCUR" & vbCrLf _
            & ", SUM (QTY180) QTY180, SUM (AMT180) AMT180" & vbCrLf _
            & ", SUM (QTY360) QTY360, SUM (AMT360) AMT360" & vbCrLf _
            & ", SUM (PICK) PICK, SUM (OPEN) OPEN" & vbCrLf _
            & ", SUM (ONPO) ONPO, SUM (TRAN) TRAN" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", SUM (LOT_QTY_ONHD) QTY" & vbCrLf _
            & ", SUM (LOT_AMT_ONHD) AMT" & vbCrLf _
            & ", SUM (CASE WHEN SYSDATE - TRAN_DATE < 180 THEN LOT_QTY_ONHD ELSE 0 END) QTYCUR" & vbCrLf _
            & ", SUM (CASE WHEN SYSDATE - TRAN_DATE < 180 THEN LOT_AMT_ONHD ELSE 0 END) AMTCUR" & vbCrLf _
            & ", SUM (CASE WHEN SYSDATE - TRAN_DATE >= 180 AND SYSDATE - TRAN_DATE < 360 THEN LOT_QTY_ONHD ELSE 0 END) QTY180" & vbCrLf _
            & ", SUM (CASE WHEN SYSDATE - TRAN_DATE >= 180 AND SYSDATE - TRAN_DATE < 360 THEN LOT_AMT_ONHD ELSE 0 END) AMT180" & vbCrLf _
            & ", SUM (CASE WHEN SYSDATE - TRAN_DATE >= 360 THEN LOT_QTY_ONHD ELSE 0 END) QTY360" & vbCrLf _
            & ", SUM (CASE WHEN SYSDATE - TRAN_DATE >= 360 THEN LOT_AMT_ONHD ELSE 0 END) AMT360" & vbCrLf _
            & ", 0 PICK, 0 OPEN, 0 ONPO, 0 TRAN" & vbCrLf _
            & " from ICTCOSTL" & vbCrLf _
            & "WHERE OPS_YYYYPP_FIFO = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
            & "UNION" & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE, 0 QTY, 0 AMT, 0 QTYCUR, 0 AMTCUR, 0 QTY180, 0 AMT180, 0 QTY360, 0 AMT360" & vbCrLf _
            & ", SUM (WHSE_QTY_PICK) PICK" & vbCrLf _
            & ", SUM (WHSE_QTY_OPEN) OPEN" & vbCrLf _
            & ", SUM (WHSE_QTY_ON_ORDER) ONPO" & vbCrLf _
            & ", SUM (WHSE_QTY_TRAN) TRAN" & vbCrLf _
            & " from ICTSTAT2 GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ") GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ") X," & vbCrLf _
            & "(" & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", COUNT (DISTINCT PO_ORDER_NO) POS" & vbCrLf _
            & ", MIN (PO_ORDER_NO) MINPO, MAX (PO_ORDER_NO) MAXPO" & vbCrLf _
            & ", MIN (PO_COST) POCOST1, MAX (PO_COST) POCOST2" & vbCrLf _
            & ", SUM (PO_QTY_OPN) OPN" & vbCrLf _
            & ", SUM (PO_QTY_SHP) SHP" & vbCrLf _
            & " from POTORDR2 WHERE PO_STATUS = 'O' AND PO_QTY_OPN <> 0" & vbCrLf _
            & "GROUP BY STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ") Y," & vbCrLf _
            & "ICTSTYC1" & vbCrLf _
            & "WHERE Y.STYLE_CODE (+) = X.STYLE_CODE AND Y.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
            & "AND ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE AND ICTSTYC1.COLOR_CODE(+) = X.COLOR_CODE" & vbCrLf _
            & ") Z WHERE ICTSTYL1.STYLE_CODE = Z.STYLE_CODE AND ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE" & vbCrLf _
            & "AND P.STYLE_CODE (+) = Z.STYLE_CODE AND P.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
            & ") Q"


        If SATSLSI1 = "" Then
            SATSLSI1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSI1 & " Add Primary Key (STYLE_CODE, COLOR_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSI1)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSI1 & " " & ASCMAIN1.sql)
        End If

    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Print_Report_End()
    End Sub

    Sub Load_XLS()

        Dim DT As DataTable = dst.Tables("SATSLSI1")
        Dim xls_filename As String = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & "_XLS") & ".XLS"
        Dim xls_template As String = ASCMAIN1.Folders("Archive") & "\Templates\Inventory Positions.xls"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A2")
 
        Dim RX As Int32 = 0
        Dim CX As Int32 = 0

        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If
        Next

        'With oSheet.Cells(Excel_Cell(RX + 0 + 1, CX + 0 + 1) & ":" & Excel_Cell(RX + 0 + 1, CX + DT.Columns.Count - 1 + 1))
        '    .Font.Color = SpreadsheetGear.Colors.White
        '    .Font.Bold = True
        '    .Interior.Color = SpreadsheetGear.Colors.Blue
        'End With

        'For c As Integer = 0 To DT.Columns.Count - 1

        '    Dim COLUMN_NAME As String = DT.Columns(c).ColumnName
        '    If COLUMN_NAME = "STORES" Or COLUMN_NAME = "ORD29" Or COLUMN_NAME = "ORD30" Or COLUMN_NAME.StartsWith("UNITS") Then

        '        With oSheet.Cells(Excel_Cell(RX + 0 + 1, CX + c + 1)).EntireColumn
        '            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        '            .NumberFormat = "#,##0"
        '            If COLUMN_NAME.StartsWith("UNITS") Then
        '                .ColumnWidth = .ColumnWidth * 2
        '            End If

        '        End With

        '        oSheet.Cells(Excel_Cell(RX + 1, CX + c + 1)).Formula = "=SUBTOTAL(9," & Excel_Cell(RX + 1 + 1 + 1, CX + c + 1) & ":" & Excel_Cell(RX + 1 + DT.Rows.Count, CX + c + 1) & ")"
        '    Else
        '        With oSheet.Cells(Excel_Cell(RX + 0 + 1, CX + c + 1)).EntireColumn
        '            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        '            .NumberFormat = "@"
        '            If COLUMN_NAME = "PO29" Or COLUMN_NAME = "PO30" Then
        '                .ColumnWidth = .ColumnWidth * 2
        '            End If

        '        End With
        '    End If
        'Next

        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        'With oSheet.Cells(Excel_Cell(RX + 1, CX + 1)).EntireRow
        '    .Insert(SpreadsheetGear.InsertShiftDirection.Down)
        'End With

        'For c As Integer = 0 To DT.Columns.Count - 1
        '    Dim COLUMN_NAME As String = DT.Columns(c).ColumnName
        '    If COLUMN_NAME = "STORES" Or COLUMN_NAME = "ORD29" Or COLUMN_NAME = "ORD30" Or COLUMN_NAME.StartsWith("UNITS") Then
        '        oSheet.Cells(Excel_Cell(RX + 1, CX + c + 1)).Formula = "=SUBTOTAL(9," & Excel_Cell(RX + 1 + 1 + 1, CX + c + 1) & ":" & Excel_Cell(RX + 1 + 1 + DT.Rows.Count, CX + c + 1) & ")"
        '    End If
        'Next

        oWB.SaveAs(ASCMAIN1.Folders("Temp") & xls_filename, SpreadsheetGear.FileFormat.Excel8)
        oWB.Close()
        range = Nothing
        oSheet = Nothing
        oWB = Nothing
        '  Dim p As Process = Process.Start(ASCMAIN1.Folders("Temp") & xls_filename)


    End Sub


    Sub Create_Pivot()

        Dim xls_filename As String = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & "_XLS") & ".XLSM"
        Dim xls_template As String = ASCMAIN1.Folders("Archive") & "\Templates\Inventory Positions.xls"

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(xls_template)

        Dim SheetName As String = "Data"
        ws = wb.Worksheets(SheetName)

        Dim DataTable As DataTable = dst.Tables("SATSLSI1")

        Dim iRx As Integer = 2

        Dim r As Integer = iRx
        For Each row As DataRow In DataTable.Select("")
            r += 1
            ws.Range(Excel_Cell(r, 1) & ":" & Excel_Cell(r, DataTable.Columns.Count)).Value2 = row.ItemArray
        Next

        ASCMAIN1.Progress("-", "Pivot")
        wb.Names.Add("DATA", "=" & SheetName & "!" & Excel_Cell(iRx + 0, 1, 3) & ":" & Excel_Cell(iRx + DataTable.Rows.Count, DataTable.Columns.Count, 3))
        'excel.Run("ResetData")

        For c As Integer = 1 To DataTable.Columns.Count
            If ws.Range(Excel_Cell(1, c)).Formula & "" <> "" Then
                ws.Range(Excel_Cell(1, c)).Formula = "=SUBTOTAL(9," & Excel_Cell(iRx + 1, c) & ":" & Excel_Cell(iRx + DataTable.Rows.Count, c) & ")"
            End If
        Next



        'refresh the pivotcache
        'ws.PivotTables("PivotTable1").PivotCache.Refresh()

        ' the line below was disabled by wjz on 02/08 as part of disabling all refreshes -but not sure if this one was a problem - this one might have worked
        '   wb.Sheets("YTD").PivotTables(1).PivotCache.Refresh()

        'Dim pt As Microsoft.Office.Interop.Excel.PivotTable
        'pt = wb.Sheets("YTD").PivotTables("PivotTable1")
        'pt.PivotCache.Refresh()

        'Marshal.ReleaseComObject(pt)

        'ws.PivotTables("PivotTable1").RefreshTable()
        'ws.PivotTables("PivotTable1").Update()

        ASCMAIN1.Progress("Now Saving Workbook")
 

        Dim objOpt As Object = Nothing
        wb.SaveAs(ASCMAIN1.Folders("Temp") & xls_filename _
                           , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
        wb.Close(False, objOpt, objOpt)

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        ReleaseCOMObject(xlDestRange)
        ReleaseCOMObject(xlSourceRange)
        ReleaseCOMObject(ws)
        ReleaseCOMObject(wb)
        ReleaseCOMObject(excel)
        Dim p As Process = Process.Start(ASCMAIN1.Folders("Temp") & xls_filename)

        ASCMAIN1.Progress("")
    End Sub


    Sub Generate_Plan_XLS()



        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP)
        Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")

        Dim DT As DataTable = TBL_PLAN

        Dim TP As New DataTable
        With TP.Columns
            .Add("COLUMN_NAME")
            .Add("TYPE")
            .Add("DATA_VALUE")
            .Add("R", GetType(System.Int32))
            .Add("C", GetType(System.Int32))
        End With

        TP.PrimaryKey = New DataColumn() {TP.Columns("COLUMN_NAME")}
        TP.Columns("R").DefaultValue = -1

        TP.Rows.Add(New Object() {"SALES_DIVISION_CODE", "HEADER"})
        TP.Rows.Add(New Object() {"CATGY_CODE", "HEADER"})
        TP.Rows.Add(New Object() {"LEGEND", "HEADER"})
        TP.Rows.Add(New Object() {"SALES_DIVISION_NAME", "HEADER"})
        TP.Rows.Add(New Object() {"CATGY_DESC", "HEADER"})
        TP.Rows.Add(New Object() {"STOCK", "HEADER"})
        TP.Rows.Add(New Object() {"CUST_CODE", "HEADER"})

        For i As Integer = 1 To 3
            Dim YP As String = ASCMAIN1.Period_Calc(RYP, -1 * i)
            rowGLTPARM2 = LookUp("GLTPARM2", YP)
            Dim L As String = rowGLTPARM2.Item("LEGEND")
            TP.Rows.Add(New Object() {"P" & Format(i, "00"), "HEADER", Mid(L, 10, 6)})
        Next
        For i As Integer = 0 To 6
            Dim YP As String = ASCMAIN1.Period_Calc(RYP, +1 * i)
            rowGLTPARM2 = LookUp("GLTPARM2", YP)
            Dim L As String = rowGLTPARM2.Item("LEGEND")
            TP.Rows.Add(New Object() {"F" & Format(i, "00"), "HEADER", Mid(L, 10, 6)})
        Next

        'Dim MAXPOSQTY As Integer = 4
        'Dim MINPOSQTY As Integer = 2
        TP.Rows.Add(New Object() {"MAXPOSQTY", "HEADER", MAXPOSQTY})
        TP.Rows.Add(New Object() {"MINPOSQTY", "HEADER", MINPOSQTY})

        'Dim MAXPOSCST As Integer = 4
        'Dim MINPOSCST As Integer = 2
        TP.Rows.Add(New Object() {"MAXPOSCST", "HEADER", MAXPOSCST})
        TP.Rows.Add(New Object() {"MINPOSCST", "HEADER", MINPOSCST})

        For Each DC As DataColumn In TBL_PLAN.Columns
            If DC.DataType.ToString = "System.String" Then
            Else
                TP.Rows.Add(New Object() {DC.ColumnName, "DATA"})
            End If
        Next


        Dim xls_filename As String = "Plan_" & ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & "_XLS") & ".XLS"
        Dim xls_template As String = ASCMAIN1.Folders("Archive") & "\Templates\Plan.xlsx"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(xls_template)

        Dim oSheet_Summary As SpreadsheetGear.IWorksheet = oWB.Worksheets("Summary")
        Dim oSheet_Plan As SpreadsheetGear.IWorksheet = oWB.Worksheets("Plan")

        For RP As Integer = 0 To oSheet_Plan.UsedRange.RowCount - 1
            For CP As Integer = 0 To oSheet_Plan.UsedRange.ColumnCount - 1
                With oSheet_Plan.Cells(RP, CP)
                    If .Value & "" <> "" Then
                        Dim COLUMN_NAME As String = .Value
                        Dim row As DataRow = TP.Rows.Find(COLUMN_NAME)
                        If row IsNot Nothing Then
                            row.Item("R") = RP
                            row.Item("C") = CP
                            .ClearContents()
                        End If
                    End If
                End With
            Next
        Next

        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing '  oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = Nothing ' oSheet.Cells("A2")

        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing '  oSheet_Plan.Cells(0, 0, oSheet_Plan.Cells.RowCount, oSheet_Plan.Cells.ColumnCount)
        Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

        Dim RX As Int32 = 3
        Dim CX As Int32 = 3

        Dim HEADER_DATA As New Dictionary(Of String, String)

        For Each row As DataRow In DT.Select("", "SALES_DIVISION_CODE, CATGY_CODE, STK")
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE") & ""
            Dim CATGY_CODE As String = row.Item("CATGY_CODE") & ""
            Dim STK As String = row.Item("STK") & ""
            Dim STOCK As String = IIf(STK = "S", "Stock", "Non-Stock")

            Dim CUST_CODE As String = "" ' = row.Item("CUST_CODE") & ""

            Dim rowSOTSDIV1 As DataRow = dst.Tables("SOTSDIV1").Rows.Find(SALES_DIVISION_CODE)
            Dim SALES_DIVISION_NAME As String = "?"
            If rowSOTSDIV1 IsNot Nothing Then SALES_DIVISION_NAME = rowSOTSDIV1.Item("SALES_DIVISION_NAME") & ""
            Dim rowICTPCAT1 As DataRow = dst.Tables("ICTPCAT1").Rows.Find(CATGY_CODE)
            Dim CATGY_DESC As String = "?"
            If rowICTPCAT1 IsNot Nothing Then CATGY_DESC = rowICTPCAT1.Item("CATGY_DESC") & ""

            HEADER_DATA.Clear()

            HEADER_DATA.Add("SALES_DIVISION_CODE", SALES_DIVISION_CODE)
            HEADER_DATA.Add("CATGY_CODE", CATGY_CODE)
            HEADER_DATA.Add("LEGEND", LEGEND)
            HEADER_DATA.Add("SALES_DIVISION_NAME", SALES_DIVISION_NAME)
            HEADER_DATA.Add("CATGY_DESC", CATGY_DESC)
            HEADER_DATA.Add("STOCK", STOCK)
            HEADER_DATA.Add("CUST_CODE", CUST_CODE)
 


            oSheet = oWB.Worksheets("Plan").CopyAfter(oWB.Worksheets(oWB.Worksheets.Count - 1))

            oSheet.Name = SALES_DIVISION_CODE & "-" & CATGY_CODE & "-" & STK

            'oSheet.Cells("A1").Value = SALES_DIVISION_CODE
            'oSheet.Cells("A2").Value = CATGY_CODE

            RX += 1
            oSheet_Summary.Cells(RX, CX + 0).Value = SALES_DIVISION_CODE
            oSheet_Summary.Cells(RX, CX + 1).Value = SALES_DIVISION_NAME
            oSheet_Summary.Cells(RX, CX + 2).Value = CATGY_CODE
            oSheet_Summary.Cells(RX, CX + 3).Value = CATGY_DESC
            oSheet_Summary.Cells(RX, CX + 4).Value = STOCK
            oSheet_Summary.Hyperlinks.Add(oSheet_Summary.Cells(RX, CX - 1), "", "'" & oSheet.Name & "'!A1", "Click Here to Navigate to " & oSheet.Name, "See Plan")
            oSheet.Hyperlinks.Add(oSheet.Cells("A4"), "", "'" & oSheet_Summary.Name & "'!A1", "Click Here to Navigate back to Summary Sheet", "Summary")

            Dim CXO As Integer

            CXO = 5 ' Revenue
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G5"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G6"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G8"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G11"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G12"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G14"

            CXO = 12 ' Costs
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G20"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G21"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G22"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G23"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G28"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G29"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G31"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G32"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G33"

            CXO = 24 ' Units
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G35"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G36"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G37"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G38"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G43"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G44"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G46"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G47"
            CXO += 1 : oSheet_Summary.Cells(RX, CX + CXO).Formula = "='" & oSheet.Name & "'!G48"

            For Each rowTP As DataRow In TP.Select("TYPE = 'DATA'")
                Dim COLUMN_NAME As String = rowTP.Item("COLUMN_NAME") & ""
                Dim R As Integer = Val(rowTP.Item("R") & "")
                Dim C As Integer = Val(rowTP.Item("C") & "")
                If R <> -1 Then
                    oSheet.Cells(R, C).Value = row.Item(COLUMN_NAME)
                End If
            Next

            For Each rowTP As DataRow In TP.Select("TYPE = 'HEADER'")
                Dim COLUMN_NAME As String = rowTP.Item("COLUMN_NAME") & ""
                Dim DATA_VALUE As String = rowTP.Item("DATA_VALUE") & ""
                Dim R As Integer = Val(rowTP.Item("R") & "")
                Dim C As Integer = Val(rowTP.Item("C") & "")
                If R <> -1 Then
                    If DATA_VALUE <> "" Then
                        oSheet.Cells(R, C).Value = DATA_VALUE
                    Else
                        oSheet.Cells(R, C).Value = HEADER_DATA(COLUMN_NAME)
                    End If
                End If
            Next
        Next

        Dim RMAX As Integer = TBL_PLAN.Rows.Count

        Dim ROWSUB As Integer = 1
        For COLSUB As Integer = CX To oSheet_Summary.UsedRange.ColumnCount - 1
            With oSheet_Summary.Cells(ROWSUB, COLSUB)
                Dim FORMULA As String = .Formula
                If FORMULA.StartsWith("=SUBTOTAL(9") Then
                    .Formula = "=SUBTOTAL(9," & Excel_Cell0(ROWSUB + 2 + 1, COLSUB) & ":" & Excel_Cell0(ROWSUB + 2 + RMAX, COLSUB) & ")"
                End If
            End With
        Next

            range = oSheet_Summary.Cells(3, 3, 3 + 1 + RMAX, 7)
        range.AutoFilter()


        oSheet_Summary.Cells(0, CX - 1).Value = Format(Now, "MM/dd/yyyy HH:mm")
        oSheet_Summary.Cells(0, CX - 0).Value = ASCMAIN1.USER_ID
        Dim rowGLTPARM2_RYP As DataRow = LookUp("GLTPARM2", RYP)
        oSheet_Summary.Cells(1, CX - 1).Value = rowGLTPARM2_RYP.Item("LEGEND")


        oWB.Worksheets("Plan").Visible = False

        oWB.Worksheets("Summary").Select()


        oWB.SaveAs(ASCMAIN1.Folders("Temp") & xls_filename, SpreadsheetGear.FileFormat.Excel8)
        oWB.Close()
        range = Nothing
        oSheet = Nothing
        oWB = Nothing
        Dim p As Process = Process.Start(ASCMAIN1.Folders("Temp") & xls_filename)


    End Sub

    Private Sub btnPlusDuty_Click(sender As Object, e As EventArgs) Handles btnPlusDuty.Click
        Generate_Plan_XLS()
    End Sub
End Class