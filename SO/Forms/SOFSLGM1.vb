Public Class SOFSLGM1

#Region "Declarations"
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -48, 0, -11)

        Set_cmbYP("RYP1", ASCMAIN1.CYP, -48, 0, 0)

        With dst
            ASCMAIN1.sql = "SELECT X.*, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_CUST_PO, ARTCUST1.CUST_NAME FROM SOTORDR1, ARTCUST1, (" & vbCrLf _
                & "SELECT CUST_CODE, COUNT(*) ORDS, MAX(ORDR_NO) ORDR_NO FROM(" & vbCrLf _
                & "Select ORDR_NO, CUST_CODE, COUNT(*) POS, MAX(PO_ORDER_NO) PO_ORDER_NO from POTORDR1 WHERE ORDR_NO IS NOT NULL" & vbCrLf _
                & "GROUP BY ORDR_NO, CUST_CODE)" & vbCrLf _
                & "GROUP BY CUST_CODE) X WHERE ARTCUST1.CUST_CODE = X.CUST_CODE AND SOTORDR1.ORDR_NO = X.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTSLGMX", "**", 0, False)

            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO" _
                & " from SOTORDR1" _
                & " where SOTORDR1.ORDR_NO IN (SELECT DISTINCT (ORDR_NO) FROM SOTINVH1 WHERE SOTINVH1.CUST_CODE = :PARM1 AND SOTINVH1.ORDR_YYYYPP_UPDATED >= :PARM2 AND SOTINVH1.ORDR_YYYYPP_UPDATED <= :PARM3 AND SOTINVH1.ORDR_TYPE_CODE = 'BTB')"

            Create_TDA(.Tables.Add, "SOTSLGM1", "**", 0, False, "VVV")
            With .Tables("SOTSLGM1").Columns
                .Add("PO_ORDER_NOS")
                .Add("PO_SHIPMENT_NOS")
                .Add("SALES_REVENUES", GetType(System.Decimal))
                .Add("CB_TO_FACTORIES", GetType(System.Decimal))
                .Add("INSPECTION_SHORTAGE", GetType(System.Decimal))
                .Add("DEDUCTIONS", GetType(System.Decimal))
                .Add("CGS")
                .Add("PURCHASES", GetType(System.Decimal))
                .Add("CDS", GetType(System.Decimal))
                .Add("CBFEES", GetType(System.Decimal))
                .Add("OCEAN", GetType(System.Decimal))
                .Add("COMM", GetType(System.Decimal))
                .Add("DEMUR", GetType(System.Decimal))
                .Add("DET", GetType(System.Decimal))
                .Add("AIR", GetType(System.Decimal))
                .Add("CUSTOM", GetType(System.Decimal))
                .Add("HANDLE20", GetType(System.Decimal))
                .Add("HAUL", GetType(System.Decimal))
                .Add("HNDL", GetType(System.Decimal))
                .Add("INSPEC", GetType(System.Decimal))
                .Add("LABFEE", GetType(System.Decimal))
                .Add("LBL", GetType(System.Decimal))
                .Add("DUTY", GetType(System.Decimal))
                .Add("FUMIG", GetType(System.Decimal))
                .Add("LBLSHP", GetType(System.Decimal))
                .Add("LC", GetType(System.Decimal))
                .Add("LCLCFS", GetType(System.Decimal))
                .Add("MARINE", GetType(System.Decimal))
                .Add("PIER", GetType(System.Decimal))
                .Add("RWRK", GetType(System.Decimal))
                .Add("SAM", GetType(System.Decimal))
                .Add("SAMSHP", GetType(System.Decimal))
                .Add("SCOM", GetType(System.Decimal))
                .Add("STG", GetType(System.Decimal))
                .Add("TELEX", GetType(System.Decimal))
                .Add("THC_ORC", GetType(System.Decimal))
                .Add("TRUCK", GetType(System.Decimal))
                .Add("FWL", GetType(System.Decimal))
                .Add("TRUCKING", GetType(System.Decimal))
                .Add("LABELS", GetType(System.Decimal))
                .Add("GROSS_MARGIN", GetType(System.Decimal), "ISNULL(SALES_REVENUES,0) - ISNULL( PURCHASES,0) - ISNULL( AIR,0) - ISNULL(CDS,0) - ISNULL(CBFEES,0) - ISNULL(COMM,0) - ISNULL(CUSTOM,0) - ISNULL(DEMUR,0) - ISNULL(DET,0) - ISNULL( DUTY,0) - ISNULL(FUMIG,0) - ISNULL( FWL,0) - ISNULL( HANDLE20,0) -  ISNULL( HAUL ,0) - ISNULL(HNDL ,0) - ISNULL(INSPEC ,0) - ISNULL(LABFEE,0) - ISNULL(LBL,0) - ISNULL(LBLSHP ,0) - ISNULL(LC ,0) - ISNULL(LCLCFS ,0) - ISNULL(MARINE  ,0) - ISNULL(OCEAN ,0) - ISNULL(PIER  ,0) - ISNULL(RWRK ,0) - ISNULL(SAM,0) - ISNULL(SAMSHP  ,0) - ISNULL(SCOM  ,0) - ISNULL(STG ,0) - ISNULL(TELEX ,0) - ISNULL(THC_ORC ,0) - ISNULL(TRUCK, 0) ")
                .Add("GROSS_MARGIN_PCT", GetType(System.Decimal), "IIF(ISNULL(SALES_REVENUES,0)=0, 0, 100 * ISNULL(GROSS_MARGIN,0) /  ISNULL(SALES_REVENUES,0) )")
            End With


        End With

        grdSOTSLGMX.DataSource = dst.Tables("SOTSLGMX")
        grdSOTSLGM1.DataSource = dst.Tables("SOTSLGM1")


        grdSOTSLGM1.DisplayLayout.UseFixedHeaders = True
        With grdSOTSLGM1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).Header.VisiblePosition = 1
            Next
        End With

        grdSOTSLGMX.DisplayLayout.UseFixedHeaders = True
        With grdSOTSLGMX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTSLGM1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                '    If gcol.Key = "STYLE_DESC" Or gcol.Key = "COLOR_DESC" Or gcol.Key = "STATE" Then
                '        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                '        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                '    ElseIf New String() {"INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE"}.Contains(gcol.Key) Then
                '        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                '        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                '    Else
                '        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                '    End If
                gcol.Header.Appearance.TextHAlign = HAlign.Left

                For Each COLUMN_NAME As String In New String() {"ORDR_NO"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 1
                    .Columns(COLUMN_NAME).Header.Caption = "Order No"
                Next

                For Each COLUMN_NAME As String In New String() {"ORDR_CUST_PO"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 2
                    .Columns(COLUMN_NAME).Header.Caption = "Customer PO"
                Next

                For Each COLUMN_NAME As String In New String() {"PO_ORDER_NOS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 3
                    .Columns(COLUMN_NAME).Header.Caption = "PO Order Nos"
                Next

                For Each COLUMN_NAME As String In New String() {"PO_SHIPMENT_NOS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 4
                    .Columns(COLUMN_NAME).Header.Caption = "PO Shipments Nos"
                Next

                For Each COLUMN_NAME As String In New String() {"SALES_REVENUES"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 5
                Next

                For Each COLUMN_NAME As String In New String() {"PURCHASES"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 6
                    .Columns(COLUMN_NAME).Header.Caption = "Purchases"
                Next

                For Each COLUMN_NAME As String In New String() {"CB_TO_FACTORIES"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 6
                Next

                For Each COLUMN_NAME As String In New String() {"CB_TO_FACTORIES"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 7
                Next

                For Each COLUMN_NAME As String In New String() {"INSPECTION_SHORTAGE"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 8
                    .Columns(COLUMN_NAME).Header.Caption = "Inspection Shortage"
                Next

                For Each COLUMN_NAME As String In New String() {"DEDUCTIONS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 9
                    .Columns(COLUMN_NAME).Header.Caption = "Deductons"
                Next

                For Each COLUMN_NAME As String In New String() {"CGS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 10
                Next

                For Each COLUMN_NAME As String In New String() {"CDS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 11
                Next

                For Each COLUMN_NAME As String In New String() {"CBFEES"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 12
                Next


                For Each COLUMN_NAME As String In New String() {"DUTY"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 13
                    .Columns(COLUMN_NAME).Header.Caption = "Duty"
                Next

                For Each COLUMN_NAME As String In New String() {"OCEAN"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 14
                    .Columns(COLUMN_NAME).Header.Caption = "Ocean"
                Next

                For Each COLUMN_NAME As String In New String() {"CUSTOM"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 15
                    .Columns(COLUMN_NAME).Header.Caption = "Customs"
                Next

                For Each COLUMN_NAME As String In New String() {"COMM"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 16
                Next

                For Each COLUMN_NAME As String In New String() {"TRUCKING"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 17
                    .Columns(COLUMN_NAME).Header.Caption = "Trucking"
                Next

                For Each COLUMN_NAME As String In New String() {"DEMUR"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 18
                Next

                For Each COLUMN_NAME As String In New String() {"LABELS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 19
                Next

                For Each COLUMN_NAME As String In New String() {"DET"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 20
                Next

                For Each COLUMN_NAME As String In New String() {"AIR"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 21
                Next

                For Each COLUMN_NAME As String In New String() {"HANDLE20"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 22
                Next

                For Each COLUMN_NAME As String In New String() {"HAUL"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 23
                Next



                For Each COLUMN_NAME As String In New String() {"HNDL"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 24
                Next

                For Each COLUMN_NAME As String In New String() {"INSPEC"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 25
                Next

                For Each COLUMN_NAME As String In New String() {"LABFEE"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 26
                Next

                For Each COLUMN_NAME As String In New String() {"LBL"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 27
                Next

                For Each COLUMN_NAME As String In New String() {"FUMIG"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 28
                    .Columns(COLUMN_NAME).Header.Caption = "Fumigation"
                Next

                For Each COLUMN_NAME As String In New String() {"LBLSHP"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 29
                Next

                '888

                For Each COLUMN_NAME As String In New String() {"LC"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 30
                Next

                For Each COLUMN_NAME As String In New String() {"LCLCFS"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 31
                Next

                For Each COLUMN_NAME As String In New String() {"MARINE"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 32
                    .Columns(COLUMN_NAME).Header.Caption = "Marine"
                Next

                For Each COLUMN_NAME As String In New String() {"PIER"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 33
                    .Columns(COLUMN_NAME).Header.Caption = "Pier"
                Next

                For Each COLUMN_NAME As String In New String() {"RWRK"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 34
                    .Columns(COLUMN_NAME).Header.Caption = "Rework"
                Next

                For Each COLUMN_NAME As String In New String() {"SAM"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 35
                Next

                For Each COLUMN_NAME As String In New String() {"SAMSHP"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 36
                Next

                For Each COLUMN_NAME As String In New String() {"SCOM"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 37
                Next

                For Each COLUMN_NAME As String In New String() {"STG"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 38
                Next

                For Each COLUMN_NAME As String In New String() {"TELEX"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 39
                    .Columns(COLUMN_NAME).Header.Caption = "Telex"
                Next

                For Each COLUMN_NAME As String In New String() {"THC_ORC"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 40
                Next

                For Each COLUMN_NAME As String In New String() {"TRUCK"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 41
                    .Columns(COLUMN_NAME).Header.Caption = "Truck"
                Next


                For Each COLUMN_NAME As String In New String() {"FWL"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 42
                Next

                For Each COLUMN_NAME As String In New String() {"GROSS_MARGIN"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 43
                    .Columns(COLUMN_NAME).Header.Caption = "Gross Margin $ "
                Next

                For Each COLUMN_NAME As String In New String() {"GROSS_MARGIN_PCT"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                    .Columns(COLUMN_NAME).Header.VisiblePosition = 44
                    .Columns(COLUMN_NAME).Header.Caption = "Gross Margin %"
                Next

            Next

        End With
        Show_Filter(grdSOTSLGMX, True)
        Create_Summary(grdSOTSLGMX, "CUST_CODE", "Count")
        Create_Summary(grdSOTSLGM1, "ORDR_NO", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Visible = (EntryMode = "V" And ScreenMode)
        End With

        grdSOTSLGMX.Visible = Not ScreenMode
        grdSOTSLGM1.Visible = ScreenMode
        ' grpPERIOD_RANGE.Enabled = Not ScreenMode
        Set_Read_Only(grpPERIOD_RANGE, ScreenMode)
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CUST_CODE").Text = ""
        CUST_CODE = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSLGM1", "SOTSLGMX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdSOTSLGM1.Rows.ColumnFilters.ClearAllFilters()


        Load_SOTSLGMX("")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value
        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value
        Fill_Records("SOTSLGM1", New String() {CUST_CODE, RYP0, RYP1})

        For Each rowSOTSLGM1 As DataRow In dst.Tables("SOTSLGM1").Select("")
            Dim ORDR_NO As String = rowSOTSLGM1.Item("ORDR_NO")

            ASCMAIN1.sql = "SELECT SUM(INV_SALES) FROM SOTINVH1 WHERE ORDR_NO = '" & ORDR_NO & "'"
            Dim SALES_REVENUES As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("SALES_REVENUES") = SALES_REVENUES

            ASCMAIN1.sql = "SELECT SUM(PO_QTY_REC * PO_COST) FROM POTORDR2, POTORDR1 WHERE POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
            Dim PURCHASES As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("PURCHASES") = PURCHASES




            ASCMAIN1.sql = "SELECT PO_ORDER_NO FROM POTORDR1 WHERE ORDR_NO = '" & ORDR_NO & "'"
            Dim PO_ORDER_NOS As String = ""
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                PO_ORDER_NOS &= "," & row.Item("PO_ORDER_NO")
            Next
            rowSOTSLGM1.Item("PO_ORDER_NOS") = Mid(PO_ORDER_NOS, 2)

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                        & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                        & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTLCST1.COST_CATGY_CODE = 'CBFEES'"
            Dim CBFEES As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("CBFEES") = CBFEES


            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                       & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                       & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                       & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                       & " AND POTLCST1.COST_CATGY_CODE = 'DUTY'"
            Dim DUTY As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("DUTY") = DUTY


            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                     & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                     & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                     & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                     & " AND POTLCST1.COST_CATGY_CODE = 'LBL'"
            Dim LBL As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("LBL") = LBL


            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                     & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                     & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                     & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                     & " AND POTLCST1.COST_CATGY_CODE = 'CDS'"
            Dim CDS As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("CDS") = CDS


            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                     & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                     & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                     & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                     & " AND POTLCST1.COST_CATGY_CODE = 'COMM'"
            Dim COMM As Decimal = Val(ASCDATA1.GetDataValue)

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                      & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                      & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                      & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                      & " AND POTLCST1.COST_CATGY_CODE = 'AIR'"
            Dim AIR As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("AIR") = AIR

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                       & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                       & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                       & " AND POTLCST1.COST_CATGY_CODE = 'CUSTOM'"
            Dim CUSTOM As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("CUSTOM") = CUSTOM

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                     & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                     & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                     & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                     & " AND POTLCST1.COST_CATGY_CODE = 'DEMUR'"
            Dim DEMUR As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("DEMUR") = DEMUR

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                    & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                    & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                    & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                    & " AND POTLCST1.COST_CATGY_CODE = 'DET'"
            Dim DET As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("DET") = DET

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                      & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                      & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                      & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                      & " AND POTLCST1.COST_CATGY_CODE = 'FUMIG'"
            Dim FUMIG As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("FUMIG") = FUMIG

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                   & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                   & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                   & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                   & " AND POTLCST1.COST_CATGY_CODE = 'FWL'"
            Dim FWL As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("FWL") = FWL

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                    & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                    & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                    & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                    & " AND POTLCST1.COST_CATGY_CODE = 'OCEAN'"
            Dim OCEAN As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("OCEAN") = OCEAN

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                        & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                        & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTLCST1.COST_CATGY_CODE = 'HANDLE20'"
            Dim HANDLE20 As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("HANDLE20") = HANDLE20

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                    & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                    & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                    & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                    & " AND POTLCST1.COST_CATGY_CODE = 'HAUL'"
            Dim HAUL As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("HAUL") = HAUL

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                        & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                        & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTLCST1.COST_CATGY_CODE = 'HNDL'"
            Dim HNDL As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("HNDL") = HNDL

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                        & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                        & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTLCST1.COST_CATGY_CODE = 'INSPEC'"
            Dim INSPEC As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("INSPEC") = INSPEC

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                        & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                        & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTLCST1.COST_CATGY_CODE = 'LABFEE'"
            Dim LABFEE As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("LABFEE") = LABFEE

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'LBLSHP'"
            Dim LBLSHP As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("LBLSHP") = LBLSHP

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'LC'"
            Dim LC As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("LC") = LC

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'LCLCFS'"
            Dim LCLCFS As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("LCLCFS") = LCLCFS


            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                          & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                          & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                          & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                          & " AND POTLCST1.COST_CATGY_CODE = 'MARINE'"
            Dim MARINE As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("MARINE") = MARINE

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                        & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                        & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTLCST1.COST_CATGY_CODE = 'PIER'"
            Dim PIER As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("PIER") = PIER

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'RWRK'"
            Dim RWRK As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("RWRK") = RWRK

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'SAM'"
            Dim SAM As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("SAM") = SAM

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'SAMSHP'"
            Dim SAMSHP As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("SAMSHP") = SAMSHP

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'SCOM'"
            Dim SCOM As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("SCOM") = SCOM

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                          & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                          & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                          & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                          & " AND POTLCST1.COST_CATGY_CODE = 'STG'"
            Dim STG As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("STG") = STG

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'TELEX'"
            Dim TELEX As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("TELEX") = TELEX

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                        & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                        & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                        & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                        & " AND POTLCST1.COST_CATGY_CODE = 'THC_ORC'"
            Dim THC_ORC As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("THC_ORC") = THC_ORC

            ASCMAIN1.sql = "SELECT SUM(POTLCST2.COST_ACT_PO) FROM POTLCST2, POTLCST1, POTORDR1  " _
                         & " WHERE POTLCST2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                         & " AND POTLCST2.CTL_NO = POTLCST1.CTL_NO" _
                         & " AND POTORDR1.ORDR_NO = '" & ORDR_NO & "'" _
                         & " AND POTLCST1.COST_CATGY_CODE = 'TRUCK'"
            Dim TRUCK As Decimal = Val(ASCDATA1.GetDataValue)
            rowSOTSLGM1.Item("TRUCK") = TRUCK


        Next
        dst.Tables("SOTSLGM1").AcceptChanges()






        Sort_grdColumns(grdSOTSLGM1, "ORDR_NO")

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        Stop
        'If EntryMode = "N" Then Exit Sub
        'For Each TABLE_NAME As String In New String() _
        '    {"SOTRSRV1", "SOTRSRV2"}
        '    Delete_Records_1(TABLE_NAME)
        'Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        'ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where CUST_CODE = '" & CUST_CODE & "'"
        'ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()


        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "EDI_DOC_SEQ_NO"
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View"

                Dim CUST_CODE As String = Split(key, ":")(0)
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Click_Command("View")
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSLGMX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTSLGM1, "SSSBS", "Show Filter", "Show GroupBox", "Show Pins", _
                        "Sales Order Inquiry", "Card View")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        If tlb_pop.Tools.Exists("Card View") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Card View"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = grd.DisplayLayout.Bands(0).CardView
        End If




        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '    e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTORDR0"


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Card View"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "CUST_CODE"
            '    If Not ScreenMode Then
            '        Load_SOTORDRX()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("Select")

        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTSLGMX(Optional PARM1 As String = "", Optional CUST_CODE As String = "")
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Customers", "")

        Fill_Records("SOTSLGMX")
        Sort_grdColumns(grdSOTSLGMX, "CUST_CODE")
        grdSOTSLGMX.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdSOTCSTYX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSLGMX.DoubleClickRow
        If Not ScreenMode Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
            Click_Command("View")
        End If
    End Sub
End Class