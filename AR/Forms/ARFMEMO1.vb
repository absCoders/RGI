Public Class ARFMEMO1

    ' PICK FROM A LIST OF INVOICES (ESP USEFUL FOR WEB CREDITS)
    ' IF WEB CREDIT, SEND CR MEMO 
    Dim rowSOTINVH1 As DataRow
    Dim rowSOTINVHC As DataRow
    Dim rowSOTMISC1 As DataRow
    Dim sqlPOTLCST2 As String
    Dim PO_SHIPMENT_NO As String

    Dim auto_CR As Boolean
    Dim CURR_CODE As String = ""
    Dim CURR_EXCH_RATE As Decimal = 0

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVHM", "*", 2)

            ASCMAIN1.sql = "Select INV_TYPE, INV_NUM, INV_DATE, INV_BALANCE, '0' SEL" _
            & " from ARTOPEN1 where CUST_CODE = :PARM1 and INV_BALANCE <> 0"
            Call Create_TDA(.Tables.Add, "ARTOPENA", "**", 0, False, "V", 0)
            .Tables("ARTOPENA").Columns.Add("INV_BALANCE_APPLIED", GetType(System.Decimal))

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")


            ASCMAIN1.sql = "Select INV_TYPE, INV_NO, INV_DATE, CUST_CODE, INIT_OPER, INIT_DATE" _
                & ", ORDR_CUST_PO, SALES_DIVISION_CODE, INV_COMMENT" _
                & ", ORDR_TYPE_CODE" _
                & ", INV_SALES, INV_FREIGHT, INV_MISC_CHG, INV_TOTAL_AMOUNT" _
                & " from SOTINVH1 where ORDR_TYPE_CODE = 'TOP' and ORDR_YYYYPP_UPDATED = :PARM1"
            Call Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "V", 2)
            '  .Tables("SOTINVHX").Columns.Add("INV_BALANCE_APPLIED", GetType(System.Decimal))



            ASCMAIN1.sql = "Select INV_TYPE, INV_NO, INV_DATE, CUST_CODE, INIT_OPER, INIT_DATE" _
                & ", ORDR_CUST_PO, SALES_DIVISION_CODE, INV_COMMENT" _
                & ", ORDR_TYPE_CODE, CCPA_NO, CC_TRANS_ID" _
                & ", INV_SALES, INV_FREIGHT, INV_MISC_CHG, INV_TOTAL_AMOUNT" _
                & " from SOTINVH1 where INV_TYPE = 'I' and CUST_CODE = :PARM1 and ORDR_YYYYPP_UPDATED = :PARM2"
            Call Create_TDA(.Tables.Add, "SOTINVHC", "**", 0, False, "VV", 2)

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")
            Create_TDA(.Tables.Add, "ARTPYMT3", "*")


            ASCMAIN1.sql = "Select POTLCST2.*" & vbCrLf _
                & ", POTLCST1.VEND_CODE, POTLCST1.COST_CATGY_CODE, POTLCST1.COST_ACT, POTLCST1.VOUCHER_NO" & vbCrLf _
                & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE" & vbCrLf _
                & ", SOTORDR1.ORDR_NO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from POTLCST2,POTLCST1,APTINVH1,POTORDR1,SOTORDR1,POTSHIP1" & vbCrLf _
                & " where POTLCST1.CTL_NO = POTLCST2.CTL_NO" & vbCrLf _
                & "   and APTINVH1.VOUCHER_NO (+) = POTLCST1.VOUCHER_NO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTLCST2.PO_ORDER_NO" & vbCrLf _
                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = POTORDR1.ORDR_NO"
            sqlPOTLCST2 = ASCMAIN1.sql

            Create_TDA(.Tables.Add, "POTLCST2", "**", 0, True, "", , "INV_NO,CHARGEBACK_STATUS")

        End With

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        grdARTOPENA.DataSource = dst.Tables("ARTOPENA")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdSOTINVHC.DataSource = dst.Tables("SOTINVHC")
        grdSOTINVHM.DataSource = dst.Tables("SOTINVHM")
        grdPOTLCST2.DataSource = dst.Tables("POTLCST2")

        Create_Summary(grdARTOPENA, "INV_TYPE", "Count")
        Create_Summary(grdARTOPENA, New String() {"INV_BALANCE", "INV_BALANCE_APPLIED"})

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, New String() {"INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT"})

        Create_Summary(grdSOTINVHM, "INV_MNO", "Count")
        Create_Summary(grdSOTINVHM, New String() {"INV_MISC_CHG"})

        Create_Summary(grdPOTLCST2, "CTL_NO", "Count")
        Create_Summary(grdPOTLCST2, New String() {"COST_ACT_PO"})

        Set_Read_Only_for_ctl(Absx1.numFor("INV_TOTAL_AMOUNT"), True)

        Show_Filter(grdSOTINVHC, True)

        tabMemo.Tabs("Apply to Open Items").Visible = False

        Bind_Controls(grpSOTINVHC, "SOTINVHC")

        ASCMAIN1.Add_Value_List(grdPOTLCST2, "CHARGEBACK_STATUS", Nothing, New String() {":", "0:Absorb", "1:Pending", "2:Re-Billed"})

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            tab0.Tabs("Invoices to Credit").Visible = False
        End If

        dteEnteredSince.Value = Now.Date.AddDays(-90)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("CUST_CODE")
                If cdr IsNot Nothing Then
                    If cdr.Item("CURR_CODE") & "" <> "USD" Then
                        MsgBox("Please Note - all amounts should be in " & cdr.Item("CURR_CODE") & "", MsgBoxStyle.OkOnly, "Verification")
                    End If
                End If

                If Absx1.dteFor("INV_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Date Specified for Entry"
                End If


                'If Absx1.optFor("INV_TYPE").Value & "" = "" Then
                '    EMsg &= vbCr & "Memo Type Required"
                'End If

                If PO_SHIPMENT_NO <> "" Then
                    If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then Exit Sub
                End If

                ' MULTITASKING

            Case "View"
                'Validate_Code("INV_NO")
                If Absx1.txtFor("INV_NO").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Document No"
                Else
                    ASCMAIN1.sql = "Select * from SOTINVH1 where INV_NO = '" & Absx1.txtFor("INV_NO").Text & "'"
                    Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow
                    If rowSOTINVH1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Document No"
                    Else
                        If rowSOTINVH1.Item("ORDR_TYPE_CODE") <> "TOP" Then
                            EMsg &= vbCr & "Invalid Type of Document to view with this screen"
                        Else
                            Absx1.optFor("INV_TYPE").Value = rowSOTINVH1.Item("INV_TYPE")
                        End If
                    End If
                End If

            Case "Update"

                If Val(Absx1.numFor("INV_TOTAL_AMOUNT").Value & "") = 0 Then
                    EMsg &= vbCr & "Total Amount is Zero"
                Else
                    If Val(Absx1.numFor("INV_TOTAL_AMOUNT").Value & "") > 0 And Absx1.optFor("INV_TYPE").Value = "C" Then
                        EMsg &= vbCr & "Credit Amount may not be Positive"
                    End If
                    If Val(Absx1.numFor("INV_TOTAL_AMOUNT").Value & "") < 0 And Absx1.optFor("INV_TYPE").Value = "D" Then
                        EMsg &= vbCr & "Debit Amount may not be Negative"
                    End If
                End If

                Dim SALES_DIVISION_CODE As String = Absx1.txtFor("SALES_DIVISION_CODE").Text
                If SALES_DIVISION_CODE = "" Then
                    EMsg &= vbCr & "Sales Division is required"
                Else
                    LookUp("SOTSDIV1", SALES_DIVISION_CODE)
                    If cdr Is Nothing Then
                        EMsg &= vbCr & "Invalid Value specified for Sales Division Code"
                    Else
                        If ASCMAIN1.CLIENT = "NYA" Then
                            Dim SEG4_CODE_DIV As String = cdr.Item("SEG4_CODE") & ""
                            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                            Dim SEG4_CODE_CUST As String = rowARTCUST1.Item("SEG4_CODE") & ""
                            If (SEG4_CODE_CUST = "001" Or SEG4_CODE_DIV = "001") And SEG4_CODE_CUST <> SEG4_CODE_DIV Then
                                If MsgBox("You are mixing NYAG Canada with NYAG US" & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If

                'Validate_Code("SALES_DIVISION_CODE")
                'Validate_Code("REASON_CODE")

                Dim DT As Date = Absx1.dteFor("INV_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                End If

                Dim ORDR_NO As String = ""
                For Each rowPOTLCST2 As DataRow In dst.Tables("POTLCST2").Select("INV_NO = '" & Absx1.txtFor("INV_NO").Text & "'")
                    If rowPOTLCST2.Item("ORDR_NO") & "" <> "" Then
                        If ORDR_NO = "" Then
                            ORDR_NO = rowPOTLCST2.Item("ORDR_NO")
                        Else
                            If rowPOTLCST2.Item("ORDR_NO") & "" <> ORDR_NO Then
                                EMsg &= vbCr & "Cannot Mix Orders on a Single Misc Chg Invoice" ' Anastasia asked for this restriction so that all invoices related to a single order may be shown in Sales Order Inquiry
                            End If
                        End If

                        If Not CHECK_FOR_COST_COMPLETE(rowPOTLCST2) Then
                            Exit Sub
                        End If

                    End If
                Next

                If EMsg = "" Then
                    If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
                        CURR_EXCH_RATE = 1
                    Else
                        Dim INV_DATE As Date = Absx1.dteFor("INV_DATE").Value ' rowSOTINVH1.Item("INV_DATE")
                        'CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, INV_DATE)
                        CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me.ROWs("GLTPARM1"), CURR_CODE, INV_DATE)

                        If CURR_EXCH_RATE <= 0 Then
                            EMsg &= vbCr & "Problem with Currency Exchange Rate"
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If chkIssueCCCredit.Visible And Not chkIssueCCCredit.Checked Then
                        If MsgBox("The Customer's Terms are Credit Card" _
                                  & vbCrLf & "Yet this Credit is not indicated to be Credited via Credit Card" _
                                  & vbCrLf & vbCrLf & "Continue with Update?", _
                                  MsgBoxStyle.YesNo, "Option to Credit Customer's Credit Card") <> MsgBoxResult.Yes Then
                            Exit Sub
                        End If
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Print Credit Memo"
                Print_Credit_Memo()

            Case "email Credit Memo"
                Print_Credit_Memo(True)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode

                If ScreenMode And EntryMode <> "N" Then
                    .Groups("Screen Control").Items("Update").Settings.Enabled = not_iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
                Else
                    .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                End If

                If ScreenMode And EntryMode <> "V" Then
                    .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
                Else
                    .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                End If

                .Groups("Show If Entered in").Visible = Not ScreenMode
                .Groups("Chargebacks").Visible = Not ScreenMode

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    .Groups("Screen Control").Items("Print Credit Memo").Visible = (InquiryMode OrElse EntryMode = "V" OrElse ScreenMode) _
                            AndAlso (rowSOTINVH1 IsNot Nothing)
                    .Groups("Screen Control").Items("email Credit Memo").Visible = (InquiryMode OrElse EntryMode = "V" OrElse ScreenMode) _
                             AndAlso (rowSOTINVH1 IsNot Nothing)
                Else
                    .Groups("Screen Control").Items("Print Credit Memo").Visible = (InquiryMode OrElse EntryMode = "V") _
                            AndAlso (rowSOTINVH1 IsNot Nothing)
                    .Groups("Screen Control").Items("email Credit Memo").Visible = (InquiryMode OrElse EntryMode = "V") _
                            AndAlso (rowSOTINVH1 IsNot Nothing)
                End If

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpMemoDetails.Visible = tf
        ' grdSOTINVHX.Visible = Not tf

        grdARTOPENA.Visible = False
        tab0.Visible = Not tf

        lblNote.Visible = ScreenMode And (Absx1.optFor("INV_TYPE").Value = "C")
        lblCURR_CODE.Visible = ScreenMode

        Set_Read_Only(grpMemoDetails, (EntryMode = "V"))
        'Set_Read_Only_for_ctl(Absx1.numFor("INV_TOTAL_AMOUNT"), True)
        If ScreenMode Then
            If EntryMode = "N" Then Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), auto_CR)
            grpSOTINVHC.Visible = auto_CR
        End If
        Set_Read_Only(grpSOTINVHC, True)

        lblDRCR.Text = Absx1.optFor("INV_TYPE").Text
        If Absx1.optFor("INV_TYPE").Value = "C" Then
            lblDRCR.Appearance.ForeColor = Drawing.Color.Red
        Else
            lblDRCR.Appearance.ForeColor = Drawing.Color.Green
        End If

        If Not ScreenMode Then
            grdPOTLCST2.Parent = tab0.Tabs("Re-Billable Charges").TabPage
            optShowChargebacks.Value = "A"
            grdPOTLCST2.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = False
            grdPOTLCST2.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = False
        Else
            grdPOTLCST2.Parent = tabMemo.Tabs("Re-Billable Charges").TabPage
            optShowChargebacks.Value = "C"
            tabMemo.Tabs("Re-Billable Charges").Visible = (dst.Tables("POTLCST2").Rows.Count > 0) And (EntryMode = "N")
            grdPOTLCST2.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
            grdPOTLCST2.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
        End If

        With grdSOTINVHM.DisplayLayout.Override
            If EntryMode = "N" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        If ScreenMode Then
            UltraExplorerBar1.Groups("Re-Billable Filter").Visible = False

            If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                lblCURR_CODE.Text = CURR_CODE
                lblCURR_CODE.Visible = True
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTINVH1", "ARTOPENA", "ARTPYMT1", "ARTPYMT2", "ARTPYMT3", "SOTINVHC", "SOTINVHM"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
        Setup_tab0()

        PO_SHIPMENT_NO = ""
        Absx1.dteFor("INV_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.optFor("INV_TYPE").Value = "C"
        If HFs.ContainsKey("CUST_CODE") Then Absx1.txtFor("CUST_CODE").Text = HFs("CUST_CODE")
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
      
        If EntryMode = "N" Then
            rowSOTINVH1 = dst.Tables("SOTINVH1").NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = HFs("INV_TYPE")
                .Item("INV_NO") = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                .Item("CUST_CODE") = HFs("CUST_CODE")
                .Item("CUST_BILL_TO_CUST") = HFs("CUST_CODE")
                .Item("CUST_STORE_NO") = "000000"
                .Item("INV_DATE") = Now.Date ' HFs("INV_DATE")
                .Item("ORDR_TYPE_CODE") = "TOP"

                .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                .Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE")
                .Item("INV_SALES") = 0
                .Item("INV_COGS") = 0
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP

                '  .Item("INV_PRINTED") = "1"
                '  .Item("REGISTER_IND") = "0"

                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    .Item("SALES_DIVISION_CODE") = "RIB"
                End If

                If auto_CR Then
                    .Item("ORDR_CUST_PO") = rowSOTINVHC.Item("ORDR_CUST_PO")
                    .Item("SALES_DIVISION_CODE") = rowSOTINVHC.Item("SALES_DIVISION_CODE")
                End If

                ' Place the original Invoice Number and CC Trans ID from the Invoice on the credit
                If .Item("INV_TYPE") = "C" Then
                    If rowSOTINVHC IsNot Nothing AndAlso rowSOTINVHC.Item("CC_TRANS_ID") & String.Empty <> String.Empty Then
                        .Item("CC_TRANS_ID") = rowSOTINVHC.Item("CC_TRANS_ID") & String.Empty
                        .Item("INV_NO_CR") = rowSOTINVHC.Item("INV_NO") & String.Empty
                    ElseIf MyBase.Absx1.txtFor("CCPA_NO").Text.Trim.Length > 0 Then
                        Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = '" & MyBase.Absx1.txtFor("CCPA_NO").Text & "'")
                        If rowARTCCPA1 IsNot Nothing Then
                            .Item("CC_TRANS_ID") = rowARTCCPA1.Item("TRANS_ID") & String.Empty
                            ' .Item("INV_NO_CR") = txtInvNo.Text
                        End If
                    End If
                End If


                'remaining SOTINVH1 fields
                'CUST_SHIP_TO_NO
                'ORDR_NO
                'WHSE_CODE
                'REGISTER_XNO
                'SHIPMENT_NO
                'PICK_NO
                'CUST_SHIP_TO_STATE
                'REGISTER_IND
                'APPLY_TO_INV_TYPE
                'APPLY_TO_INV_NO
                'INV_REVERSED
                'INV_REVERSED_INV_NO
                'INV_NO_RESHIP
                'INV_RESHIP

                'defaults from customer
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", New String() {HFs("CUST_CODE")}, True)
                .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")

                CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
                If CURR_CODE = "" Then CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = 1

                If .Item("INV_TYPE") = "I" Then
                    .Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                Else
                    .Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0") ' rowARTCUST1.Item("TERM_CODE")
                End If
            End With

            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1.ItemArray)

            Fill_Records("ARTOPENA", HFs("CUST_CODE"))

            chkIssueCCCredit.Visible = False
            chkIssueCCCredit.Checked = False

            Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", rowSOTINVH1.Item("TERM_CODE") & "")
            If rowTATTERM1 IsNot Nothing Then
                If rowTATTERM1.Item("TERM_TYPE") & "" = "D" Then
                    chkIssueCCCredit.Visible = True
                    ' chkIssueCCCredit.Checked = True
                End If
            End If
        Else
            ASCMAIN1.sql = "Select * from SOTINVH1 where INV_NO = '" & Absx1.txtFor("INV_NO").Text & "'"
            rowSOTINVH1 = Fill_Record("SOTINVH1", New String() {Absx1.optFor("INV_TYPE").Value, Absx1.txtFor("INV_NO").Text})
            dst.AcceptChanges()

            Fill_Records("SOTINVHM", New String() {Absx1.optFor("INV_TYPE").Value, Absx1.txtFor("INV_NO").Text})

            chkIssueCCCredit.Visible = False

            With UltraExplorerBar1.Groups("Screen Control").Items("Print Credit Memo")
                If Absx1.optFor("INV_TYPE").Value = "I" Then
                    .Text = "Print Invoice"
                Else
                    .Text = "Print Credit Memo"
                End If
            End With
            With UltraExplorerBar1.Groups("Screen Control").Items("email Credit Memo")
                If Absx1.optFor("INV_TYPE").Value = "I" Then
                    .Text = "email Invoice"
                Else
                    .Text = "email Credit Memo"
                End If
            End With

        End If


        Sort_grdColumns(grdSOTINVHM, "INV_MNO")


        If ASCMAIN1.CLIENT = "NYA" Then
            CURR_CODE = rowSOTINVH1.Item("CURR_CODE") & ""
            If CURR_CODE = "" Then
                CURR_CODE = "USD"
                CURR_EXCH_RATE = 1
            Else
                If CURR_CODE = "USD" Then
                    CURR_EXCH_RATE = 1
                Else
                    'CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, Now.Date)
                    CURR_EXCH_RATE = Val(rowSOTINVH1.Item("CURR_EXCH_RATE") & "")
                End If
            End If
        End If


        ' NEED TO FILL ARTOPENA DIFFERENTLY IF THIS FORM IS TO BE USED IN INQUIRY MODE

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        'ALTER TABLE SOTINVHM ADD INV_MISC_CHG_CURR  NUMBER(13,2);
        For Each rowSOTINVHM As DataRow In dst.Tables("SOTINVHM").Select("")
            Dim INV_MISC_CHG_CURR As Decimal = Val(rowSOTINVHM.Item("INV_MISC_CHG") & "")
            rowSOTINVHM.Item("INV_MISC_CHG_CURR") = INV_MISC_CHG_CURR
            If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
                Dim INV_MISC_CHG As Decimal = System.Math.Round(INV_MISC_CHG_CURR * CURR_EXCH_RATE, 2)
                rowSOTINVHM.Item("INV_MISC_CHG") = INV_MISC_CHG
            End If
        Next


        rowSOTINVH1 = dst.Tables("SOTINVH1").Rows(0)
        rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowSOTINVH1.Item("INV_SALES_CURR") = rowSOTINVH1.Item("INV_SALES")
        rowSOTINVH1.Item("INV_FREIGHT_CURR") = rowSOTINVH1.Item("INV_FREIGHT")
        rowSOTINVH1.Item("INV_MISC_CHG_CURR") = rowSOTINVH1.Item("INV_MISC_CHG")
        rowSOTINVH1.Item("INV_TOTAL_AMT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowSOTINVH1.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE

        If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
            For Each C As String In New String() {"INV_TOTAL_AMOUNT", "INV_SALES", "INV_FREIGHT", "INV_MISC_CHG"}
                Dim AMT As Decimal = Val(rowSOTINVH1.Item(C & "_CURR") & "")
                AMT = System.Math.Round(AMT * CURR_EXCH_RATE, 2)
                rowSOTINVH1.Item(C) = AMT
            Next
        End If

        For Each rowPOTLCST2 As DataRow In dst.Tables("POTLCST2").Select("INV_NO <> ''")
            rowPOTLCST2.Item("CHARGEBACK_STATUS") = "2"
            ' rowSOTINVH1.Item("ORDR_TYPE_CODE") = "BTB" ' EVEN THO IT IS NOT REALLY A BTB, IT IS AN INVOICE ASSOCIATED WITH A BTB ORDER
            ' if you don't leave it as TOP, then the grid showing all entries made won't show this invoice
            Dim ORDR_NO As String = rowPOTLCST2.Item("ORDR_NO")
            If ORDR_NO <> "" Then
                rowSOTINVH1.Item("ORDR_NO") = ORDR_NO
            End If

            CHECK_FOR_COST_COMPLETE(rowPOTLCST2, True)

        Next

        Update_Record_TDA("POTLCST2")

        Update_Record_TDA("SOTINVH1")

        ' rowARTOPEN1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0")

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        rowARTOPEN1.Item("CUST_CODE") = HFs("CUST_CODE")
        rowARTOPEN1.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")

        Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, rowSOTINVH1.Item("TERM_CODE"), Nothing, rowSOTINVH1.Item("INV_DATE"))

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            If Absx1.optFor("INV_TYPE").Value <> "I" Then
                INV_DUE_DATE = rowSOTINVH1.Item("INV_DATE") & ""
            End If
        Else
            INV_DUE_DATE = rowSOTINVH1.Item("INV_DATE") & ""
        End If



        rowARTOPEN1.Item("INV_DUE_DATE") = INV_DUE_DATE
        '  rowARTOPEN1.Item("INV_DUE_DATE") = rowSOTINVH1.Item("INV_DATE")
        rowARTOPEN1.Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
        rowARTOPEN1.Item("INV_DISC") = 0
        rowARTOPEN1.Item("INV_DISC_CURR") = 0
        rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("INV_BALANCE_CURR") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT_CURR")
        rowARTOPEN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTOPEN1.Item("INIT_DATE") = DATETIME_STAMP
        rowARTOPEN1.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowARTOPEN1.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowARTOPEN1.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowARTOPEN1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        rowARTOPEN1.Item("INV_NOTES") = rowSOTINVH1.Item("INV_COMMENT")
        rowARTOPEN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP

        With rowARTOPEN1
            For Each C As String In New String() {"INV_TYPE", "POST_CODE", "TERM_CODE", "SREP_CODE", "STAX_CODE", "INV_DATE", "ORDR_NO", _
                                                  "CURR_CODE", "CURR_EXCH_RATE", "ORDR_TYPE_CODE", "SALES_DIVISION_CODE", _
                                                  "INV_SALES", "INV_FREIGHT", "INV_TOTAL_AMOUNT", "INV_MISC_CHG", _
                                                    "INV_SALES_CURR", "INV_FREIGHT_CURR", "INV_TOTAL_AMOUNT_CURR", "INV_MISC_CHG_CURR"}
                rowARTOPEN1.Item(C) = rowSOTINVH1.Item(C)
            Next
        End With
       
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        ' APPLY TO OPEN AR ITEMS

        'If "THERE WAS SOMETHING APPLIED" Then

        ' ESTABLISH A SINGLE ROW IN ARTPYMT1
        ' ESTABLISH A SINGLE ROW IN ARTPYMT2

        '    Dim INV_PMT_TOTAL As Decimal = 0
        '    For Each rowARTOPENA As DataRow In dst.Tables("ARTOPENA").Select("SEL = '1'")
        '        rowARTOPEN1 = Fill_Record("ARTOPEN1", New String() _
        '        {HFs("CUST_CODE"), _
        '        rowARTOPENA.Item("INV_TYPE"), _
        '        rowARTOPENA.Item("INV_NUM")}, False, False)

        '        Dim INV_BALANCE As Double = Val(rowARTOPEN1.Item("INV_BALANCE") & "")
        '        Dim INV_PMT As Double = INV_BALANCE
        '        INV_PMT_TOTAL += INV_PMT
        '        ' WE WILL HAVE TO HANDLE THESE WHEN WE DO A BUYING GROUP WITH ANTIC DISC
        '        Dim INV_DISC_TAKEN As Double = 0
        '        Dim INV_WRITE_OFF As Double = 0

        '        Pay_Open_AR_Item(rowARTOPEN1, rowARTPYMT2_BOX, _
        '        ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"), _
        '         PYMT_BATCH_DATE, INV_PMT, INV_DISC_TAKEN, INV_WRITE_OFF, PYMT_BATCH_ILNO)

        '    Next
        'End If

        Update_Record_TDA("ARTOPEN1")

        ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV", _
                           New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")}, _
                           New String() {"INV_TYPE_IN", "INV_NO_IN"})

        Update_Record_TDA("SOTINVHM")

        CommitTrans("Update Complete")

        ' Create Web Invoices
        Try
            ASCMAIN1.Progress("Creating Web Invoice", "")
            For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                'WR waiting for testing to put this live
                'Dim CM_EXISTS As Boolean = CHK_CM_EXISTS(row.Item("INV_TYPE"), row.Item("INV_NO"))
                TAC.SOCMAIN1.CreateWebInvoice(Me, row.Item("INV_TYPE"), row.Item("INV_NO"))
                'If Not CM_EXISTS Then
                '    EMAIL_SREP(row.Item("INV_TYPE"), row.Item("INV_NO"), row.Item("CUST_CODE"))
                'End If
            Next
        Catch ex As Exception

        Finally
            ASCMAIN1.Progress("", "")

        End Try

    End Sub

    Private Sub EMAIL_SREP(ByVal CUST_CODE As String, ByVal INV_TYPE As String, ByVal INV_NO As String)
        Dim Attachment As String = ""
        If (ASCMAIN1.DBS_SERVER = "RGI" AndAlso ASCMAIN1.DBS_COMPANY = "RGI") Then
            'Only permit this to work in Production for regency.
            If (ASCMAIN1.DBS_SERVER = ASCMAIN1.DBS_COMPANY) Then
                Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPARM1 WHERE SO_PARM_KEY = 'Z'")
                If rowSOTPARM1.Table.Columns.Contains("SO_PARM_WEB_INVOICES") Then
                    Dim SO_PARM_WEB_INVOICES As String = (rowSOTPARM1.Item("SO_PARM_WEB_INVOICES") & String.Empty).ToString.Trim
                    If SO_PARM_WEB_INVOICES.Length > 0 Then
                        If My.Computer.FileSystem.DirectoryExists(SO_PARM_WEB_INVOICES) Then
                            If Not SO_PARM_WEB_INVOICES.EndsWith("\") Then
                                SO_PARM_WEB_INVOICES &= "\"
                            End If
                            SO_PARM_WEB_INVOICES &= CUST_CODE
                            SO_PARM_WEB_INVOICES &= "\"
                            Attachment = SO_PARM_WEB_INVOICES & INV_NO & ".pdf"
                        End If
                    End If
                End If

            End If
        End If
        If Attachment.Length > 0 Then

        End If
    End Sub

    Private Function CHK_CM_EXISTS(ByVal INV_TYPE As String, ByVal INV_NO As String) As Boolean
        'Only do this if RGI
        Dim RetVal As Boolean = False
        If (ASCMAIN1.DBS_SERVER = "RGI" AndAlso ASCMAIN1.DBS_COMPANY = "RGI") Then
            'Only permit this to work in Production for regency.
            If (ASCMAIN1.DBS_SERVER = ASCMAIN1.DBS_COMPANY) Then
                Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPARM1 WHERE SO_PARM_KEY = 'Z'")
                If rowSOTPARM1.Table.Columns.Contains("SO_PARM_WEB_INVOICES") Then
                    Dim SO_PARM_WEB_INVOICES As String = (rowSOTPARM1.Item("SO_PARM_WEB_INVOICES") & String.Empty).ToString.Trim
                    If SO_PARM_WEB_INVOICES.Length > 0 Then
                        If My.Computer.FileSystem.DirectoryExists(SO_PARM_WEB_INVOICES) Then
                            RetVal = True
                        End If
                    End If
                End If

            End If
        Else
            Return RetVal
        End If
        Return RetVal
    End Function

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Call Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    ' Click_Command("New", e)
                    Load_SOTINVHC()
                End If
            Case "INV_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                '   Click_Command("New")
                Load_SOTINVHC()
            Case "INV_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        Call LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then
                            Load_SOTINVHC()
                        End If
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT", "INV_MISC_CHG"
                Absx1.numFor("INV_TOTAL_AMOUNT").Value = _
                    Val(Absx1.numFor("INV_FREIGHT").Value & "") + _
                    Val(Absx1.numFor("INV_MISC_CHG").Value & "")
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTINVHX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdPOTLCST2, "SBB", "Show Filter", "Close", "Re-Open")
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

        Select Case grd.Name
            Case "grdPOTLCST2"
                tlb_btn = DirectCast(tlb_pop.Tools("Close"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (optShowChargebacks.Value <> "X")
                tlb_btn = DirectCast(tlb_pop.Tools("Re-Open"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (optShowChargebacks.Value = "X")
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Close", "Re-Open"

                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("First Select Rows by Clicking on the Row Selector to the left of each Row")
                    Exit Sub
                End If
                Dim CTL_NOs As String = ""
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim CTL_NO As String = grow.Cells("CTL_NO").Value
                    CTL_NOs &= ",'" & CTL_NO & "'"
                Next
                ' WHAT IF SOMEONE IS IN AP, OR IS IN AR MEMO ENTRY?  NO BIGGIE, SINCE WE ARE JUST MOVING THE STATUS, AND NOTHING FINANCIAL
                If CTL_NOs <> "" Then
                    Dim CHARGEBACK_STATUS_OLD As String = IIf(e.Tool.Key = "Close", "1", "0")
                    Dim CHARGEBACK_STATUS_NEW As String = IIf(e.Tool.Key = "Close", "0", "1")
                    ASCMAIN1.sql = "Update POTLCST2 Set CHARGEBACK_STATUS = '" & CHARGEBACK_STATUS_NEW & "'" _
                        & " where CTL_NO in (" & Mid(CTL_NOs, 2) & ") and CHARGEBACK_STATUS = '" & CHARGEBACK_STATUS_OLD & "'"
                    ASCDATA1.ExecuteSQL()
                    Refresh_POTLCST2()
                End If

        End Select
    End Sub

#End Region

    Private Sub grdARTOPENA_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTOPENA.AfterCellUpdate
        If e.Cell.Column.Key = "SEL" Then
            If e.Cell.Value = "1" Then
                e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = e.Cell.Row.Cells("INV_BALANCE").Value
            Else
                e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = 0
            End If
        End If
    End Sub

    Private Sub grdARTOPENA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTOPENA.BeforeRowUpdate
        'If e.Row.Cells("SEL").Value = "1" Then
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = e.Row.Cells("INV_BALANCE").Value
        'Else
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = 0
        'End If
    End Sub

    Sub Pay_Open_AR_Item( _
    ByVal rowARTOPEN1 As DataRow, _
    ByVal rowARTPYMT2 As DataRow, _
    ByVal CURR_CODE As String, _
    ByVal PYMT_BATCH_DATE As Date, _
    ByVal INV_PMT As Double, _
    ByVal INV_DISC_TAKEN As Double, _
    ByVal INV_WRITE_OFF As Double, _
    ByRef PYMT_BATCH_ILNO As Integer)

        With rowARTOPEN1
            Dim INV_BALANCE As Double = Val(.Item("INV_BALANCE") & "")

            .Item("INV_LAST_PMT") = PYMT_BATCH_DATE
            .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_PMT
            .Item("INV_DISC_TAKEN") = Val(.Item("INV_DISC_TAKEN") & "") + INV_DISC_TAKEN
            .Item("INV_WRITE_OFF") = Val(.Item("INV_WRITE_OFF") & "") + INV_WRITE_OFF
            .Item("INV_BALANCE") = Val(.Item("INV_BALANCE") & "") - (INV_PMT + INV_DISC_TAKEN + INV_WRITE_OFF)
            .Item("INV_LAST_PMT_REF") = rowARTPYMT2.Item("CUST_PYMT_REF_NO")
            .Item("INV_LAST_PMT_REF_DT") = rowARTPYMT2.Item("CUST_PYMT_REF_DATE")
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                .Item("INV_PMT_CURR") = .Item("INV_PMT")
                .Item("INV_DISC_TAKEN_CURR") = .Item("INV_DISC_TAKEN")
                .Item("INV_WRITE_OFF_CURR") = .Item("INV_WRITE_OFF")
                .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
            Else
                Stop
            End If

            Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow

            rowARTPYMT3.Item("PYMT_BATCH_NO") = rowARTPYMT2.Item("PYMT_BATCH_NO")
            rowARTPYMT3.Item("PYMT_BATCH_LNO") = rowARTPYMT2.Item("PYMT_BATCH_LNO")
            PYMT_BATCH_ILNO += 1
            rowARTPYMT3.Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
            rowARTPYMT3.Item("INV_TYPE") = .Item("INV_TYPE")
            rowARTPYMT3.Item("INV_NUM") = .Item("INV_NUM")
            rowARTPYMT3.Item("REASON_CODE") = .Item("REASON_CODE")
            rowARTPYMT3.Item("INV_DATE") = .Item("INV_DATE")
            rowARTPYMT3.Item("INV_DUE_DATE") = .Item("INV_DUE_DATE")
            rowARTPYMT3.Item("CUST_CODE_SO") = .Item("CUST_CODE_SO")
            rowARTPYMT3.Item("CUST_SHIP_TO_NO") = .Item("CUST_SHIP_TO_NO")
            rowARTPYMT3.Item("INV_CUST_PO") = .Item("INV_CUST_PO")
            rowARTPYMT3.Item("INV_BALANCE") = INV_BALANCE
            rowARTPYMT3.Item("INV_PMT") = INV_PMT
            rowARTPYMT3.Item("INV_DISC_TAKEN") = INV_DISC_TAKEN
            rowARTPYMT3.Item("INV_WRITE_OFF") = INV_WRITE_OFF
            rowARTPYMT3.Item("INV_BALANCE_NEW") = .Item("INV_BALANCE")
            rowARTPYMT3.Item("POST_CODE") = .Item("POST_CODE")
            rowARTPYMT3.Item("SEG2_CODE") = .Item("SEG2_CODE")
            rowARTPYMT3.Item("SEG3_CODE") = .Item("SEG3_CODE")
            rowARTPYMT3.Item("SEG4_CODE") = .Item("SEG4_CODE")
            If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                rowARTPYMT3.Item("INV_BALANCE_CURR") = rowARTPYMT3.Item("INV_BALANCE")
                rowARTPYMT3.Item("INV_PMT_CURR") = rowARTPYMT3.Item("INV_PMT")
                rowARTPYMT3.Item("INV_DISC_TAKEN_CURR") = rowARTPYMT3.Item("INV_DISC_TAKEN")
                rowARTPYMT3.Item("INV_WRITE_OFF_CURR") = rowARTPYMT3.Item("INV_WRITE_OFF")
                rowARTPYMT3.Item("INV_BALANCE_NEW_CURR") = rowARTPYMT3.Item("INV_BALANCE_NEW")
            Else
                Stop
            End If
            dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
        End With
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Fill_Records("SOTINVHX", YP)
        grdSOTINVHX.Text = "Entered in " & cbeYP.Text

        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            Load_SOTINVHC()
        End If

        Refresh_POTLCST2()

        Me.Cursor = Cursors.Default
    End Sub

    Sub Refresh_POTLCST2()

        ASCMAIN1.sql = sqlPOTLCST2 & vbCrLf

        If optShowChargebacks.Value = "A" Then
            grdPOTLCST2.Text = "Pending Chargebacks"
            ASCMAIN1.sql &= " and POTLCST2.CHARGEBACK_STATUS = '1'"
        ElseIf optShowChargebacks.Value = "C" Then
            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
            grdPOTLCST2.Text = "Pending Chargebacks for Customer " & CUST_CODE
            ASCMAIN1.sql &= " and POTLCST2.CHARGEBACK_STATUS = '1' and SOTORDR1.CUST_CODE = '" & CUST_CODE & "'"
        ElseIf optShowChargebacks.Value = "X" Then
            Dim OPS_YYYYPP As String = cbeYP.Value
            grdPOTLCST2.Text = "Chargebacks created in " & cbeYP.Text & " which have been Written-Off (Not to be re-billed)"
            ASCMAIN1.sql &= " and POTLCST2.CHARGEBACK_STATUS = '0' and POTLCST2.OPS_YYYYPP = '" & OPS_YYYYPP & "'"
        End If

        If chkFEOnly.Checked Then
            ASCMAIN1.sql &= " and POTSHIP1.WHSE_CODE = 'FE'"
        End If
        If chkEnteredSince.Checked Then
            ASCMAIN1.sql &= " and POTSHIP1.PO_DATE_SHIPPED >= '" & Format(dteEnteredSince.Value, "dd-MMM-yyyy") & "'"
        End If

        Fill_Records("POTLCST2", "", True, ASCMAIN1.sql)

    End Sub
    Private Sub cbeYP_ValueChanged(sender As System.Object, e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Private Sub grdSOTINVHX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("INV_NO").Text = e.Row.Cells("INV_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        UltraExplorerBar1.Groups("Show If Entered in").Visible = (tab0.SelectedTab.Key = "Re-Billable Charges" And optShowChargebacks.Value = "X") Or (tab0.SelectedTab.Key = "Memos Entered") Or (tab0.SelectedTab.Key = "Invoices to Credit")
        UltraExplorerBar1.Groups("Chargebacks").Visible = (tab0.SelectedTab.Key = "Re-Billable Charges")
        UltraExplorerBar1.Groups("Re-Billable Filter").Visible = (tab0.SelectedTab.Key = "Re-Billable Charges")
    End Sub

    Sub Load_SOTINVHC()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Fill_Records("SOTINVHC", New String() {CUST_CODE, YP})
        grdSOTINVHC.Text = "Invoices to Credit for " & CUST_CODE & " Posted in " & cbeYP.Text
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdSOTINVHC_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHC.DoubleClickRow
        If e.Row.IsDataRow Then

            If Absx1.optFor("INV_TYPE").Value = "C" Then
                Dim ORDR_CUST_PO As String = e.Row.Cells("ORDR_CUST_PO").Value
                ASCMAIN1.sql = "Select * from SOTINVH1 where INV_TYPE = 'C' and ORDR_CUST_PO = :PARM1"
                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {ORDR_CUST_PO})
                If row IsNot Nothing Then
                    If MsgBox("A CR Memo has already been posted for Customer PO " & ORDR_CUST_PO _
                              & vbCrLf & vbCrLf & "Continue with Credit Memo?",
                              MsgBoxStyle.YesNo, _
                              "CR Memo " & row.Item("INV_NO") & " has already been posted with reference to Customer PO " & ORDR_CUST_PO) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If


            auto_CR = True

            Dim INV_NO As String = e.Row.Cells("INV_NO").Value
            Dim INV_TYPE As String = e.Row.Cells("INV_TYPE").Value

            rowSOTINVHC = dst.Tables("SOTINVHC").Rows.Find(New String() {INV_TYPE, INV_NO})
            Click_Command("New")
            auto_CR = False
        End If
    End Sub

    Private Sub grdSOTINVHC_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTINVHC.InitializeLayout

    End Sub

    Private Sub Print_Credit_Memo(Optional email As Boolean = False)
        Try
            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Now Preparing Credit Memo for Printing")

            'Dim REPORT_NAME As String = "SORINVP1"
            'Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
            'If RPT = "" Then RPT = REPORT_NAME

            'If Not REPORTS.ContainsKey(REPORT_NAME) Then
            '    REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            '    REPORTS(REPORT_NAME).Prepare_dst(False, "")
            'End If

            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim ATTACHMENT As String = ""
            Dim FILENAME As String = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO)

            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowSOTINVH1.Item("CUST_CODE"))

            If email Then
                Dim SUBJECT As String = "" = "Invoice " & INV_NO
                Dim SEND_NO As String = TAC.SOCMAIN1.email_Invoice(Me, _
                  rowSOTINVH1.Item("CUST_CODE"), _
                 rowARTCUST1.Item("CUST_NAME"), _
                  rowARTCUST1.Item("CUST_EMAIL") & "", _
                  rowARTCUST1.Item("CUST_CONTACT") & "", _
                  FILENAME, IIf(ATTACHMENT = "", FILENAME, ATTACHMENT), SUBJECT, INV_NO)

                'If SEND_NO <> "" Then
                '    ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTORDR1' and TABLE_KEY = '" & ORDR_NO & "' and EVENT_KEY = '" & SEND_NO & "'"
                '    Fill_Records("TATEVNT1", "", False, ASCMAIN1.sql)

                '    Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
                'End If

            Else
                Show_Document(FILENAME)
            End If


            'Dim sql As String = " and SOTINVH1.INV_TYPE = '" & rowSOTINVH1.Item("INV_TYPE") & "' and SOTINVH1.INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "'"
            'Dim tempFileName As String = "Memo" & DateTime.Now.ToString("yyyyMMddHHmmss")

            'REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
            'Dim FILENAME As String = ""
            'With REPORTS(REPORT_NAME).clsASCBASE1
            '    .Print_Report_Begin()
            '    .CR_params.Add("SUBT", "")
            '    .CR_params.Add("CONS_INV", "")
            '    Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
            '    FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
            '    .Print_Report_End(, True)
            'End With

            'Show_Document(FILENAME)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print Invoice / Credit Memo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

#Region "grdSOTINVHM"

    Private Sub grdSOTINVHM_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTINVHM.AfterCellUpdate
        With grdSOTINVHM.ActiveRow
            Select Case e.Cell.Column.Key
                Case "MISC_CHG_CODE"
                    Dim MISC_CHG_CODE As String = Validate_MISC_CHG_CODE(.Cells("MISC_CHG_CODE").Value & "")
                    If MISC_CHG_CODE <> "" Then
                        .Cells("MISC_CHG_DESC").Value = rowSOTMISC1.Item("MISC_CHG_DESC")
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSOTINVHM_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTINVHM.AfterRowActivate

        If Trim(grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE").Value & "") = "" And _
            (grdSOTINVHM.ActiveCell Is Nothing OrElse _
             (grdSOTINVHM.ActiveCell.Column.Key <> "MISC_CHG_CODE")) _
        Then
            grdSOTINVHM.ActiveCell = grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE")
            Exit Sub
        End If

        If grdSOTINVHM.ActiveRow.IsAddRow Then
            If grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE").Value & "" = "" Then
                grdSOTINVHM.ActiveCell = grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE")
            End If
        Else
            With grdSOTINVHM.DisplayLayout.Bands(0)
                Validate_MISC_CHG_CODE(grdSOTINVHM.ActiveRow.Cells("MISC_CHG_CODE").Value & "")
            End With
        End If
    End Sub

    Private Sub grdSOTINVHM_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTINVHM.AfterRowsDeleted
        Dim DELKEYS As List(Of String) = DirectCast(grdSOTINVHM.Tag, List(Of String))
        For Each DELKEY As String In DELKEYS
            Dim CTL_NO As String = Split(DELKEY, vbTab)(0)
            Dim PO_ORDER_NO As String = Split(DELKEY, vbTab)(1)
            Dim rowPOTLCST2 As DataRow = dst.Tables("POTLCST2").Rows.Find(New String() {CTL_NO, PO_ORDER_NO})
            rowPOTLCST2.Item("INV_NO") = ""
        Next
        Display_Totals()
    End Sub

    Private Sub grdSOTINVHM_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTINVHM.AfterRowUpdate
        If e.Row.Cells("CTL_NO").Value & "" <> "" Then
            Dim CTL_NO As String = e.Row.Cells("CTL_NO").Value & ""
            Dim PO_ORDER_NO As String = e.Row.Cells("PO_ORDER_NO").Value & ""
            Dim rowPOTLCST2 As DataRow = dst.Tables("POTLCST2").Rows.Find(New String() {CTL_NO, PO_ORDER_NO})
            rowPOTLCST2.Item("INV_NO") = e.Row.Cells("INV_NO").Value
        End If
        Display_Totals()
    End Sub

    Private Sub grdSOTINVHM_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTINVHM.BeforeExitEditMode
        If grdSOTINVHM.ActiveCell IsNot Nothing Then
            With grdSOTINVHM.ActiveCell
                Select Case .Column.Key
                    Case "MISC_CHG_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTINVHM_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTINVHM.BeforeRowsDeleted
        Dim DELKEYS As New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim CTL_NO As String = grow.Cells("CTL_NO").Value & ""
            If CTL_NO <> "" Then
                Dim PO_ORDER_NO As String = grow.Cells("PO_ORDER_NO").Value
                DELKEYS.Add(CTL_NO & vbTab & PO_ORDER_NO)
            End If
        Next
        grdSOTINVHM.Tag = DELKEYS
    End Sub

    Private Sub grdSOTINVHM_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTINVHM.BeforeRowUpdate

        If Validate_MISC_CHG_CODE(e.Row.Cells("MISC_CHG_CODE").Value & "") = "" Then
            e.Cancel = True
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("INV_TYPE").Value = Absx1.optFor("INV_TYPE").Value
            e.Row.Cells("INV_NO").Value = Absx1.txtFor("INV_NO").Text
            Dim INV_MNO As Int64 = Val(dst.Tables("SOTINVHM").Compute("MAX(INV_MNO)", "") & "") + 1
            e.Row.Cells("INV_MNO").Value = INV_MNO
        End If
    End Sub

    Private Sub grdSOTINVHM_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTINVHM.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "MISC_CHG_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTINVHM, sql_where)
            End Select
        End With
    End Sub

    Function Validate_MISC_CHG_CODE(MISC_CHG_CODE As String) As String
        rowSOTMISC1 = LookUp("SOTMISC1", MISC_CHG_CODE)
        If rowSOTMISC1 Is Nothing Then
            Return ""
        Else
            Return rowSOTMISC1.Item("MISC_CHG_CODE")
        End If
    End Function
#End Region

    Sub Display_Totals()
        Synch_TABLE_NAME("SOTINVH1")
        Dim INV_MISC_CHG As Decimal = Val(dst.Tables("SOTINVHM").Compute("SUM(INV_MISC_CHG)", "") & "")
        Absx1.numFor("INV_MISC_CHG").Value = INV_MISC_CHG
    End Sub

    Private Sub optShowChargebacks_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShowChargebacks.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Refresh_POTLCST2()
        Setup_tab0()
    End Sub

    Private Sub grdPOTLCST2_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTLCST2.DoubleClickRow
        If e.Row.IsDataRow Then '
            Dim COST_CATGY_CODE As String = e.Row.Cells("COST_CATGY_CODE").Value
            Dim CTL_NO As String = e.Row.Cells("CTL_NO").Value
            Dim PO_ORDER_NO As String = e.Row.Cells("PO_ORDER_NO").Value
            Dim INV_NO As String = e.Row.Cells("INV_NO").Value & ""
            Dim COST_ACT_PO As Decimal = Val(e.Row.Cells("COST_ACT_PO").Value & "")

            If INV_NO <> "" Then
                MsgBox("Re-Billable Charge " & CTL_NO & " for " & Format(COST_ACT_PO, "$#,##0.00") & " has already been selected", MsgBoxStyle.OkOnly, "Cannot Add Re-Billable Charge as Requested")
                Exit Sub
            End If

            If Not ScreenMode Then
                Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Absx1.optFor("INV_TYPE").Value = "I"

                Dim rowPOTLCST1 As DataRow = LookUp("POTLCST1", CTL_NO)
                PO_SHIPMENT_NO = rowPOTLCST1.Item("PO_SHIPMENT_NO")

                Click_Command("New")
            End If



            If ScreenMode Then
                Add_grdSOTINVHM(COST_CATGY_CODE, CTL_NO, PO_ORDER_NO, COST_ACT_PO)
                tabMemo.SelectedTab = tabMemo.Tabs("Misc Charges")
            End If
        End If
     
    End Sub

    Sub Add_grdSOTINVHM(COST_CATGY_CODE As String, CTL_NO As String, PO_ORDER_NO As String, COST_ACT_PO As Decimal)
        With grdSOTINVHM
            If .ActiveRow IsNot Nothing AndAlso .ActiveRow.DataChanged Then
                .ActiveRow.CancelUpdate()
            End If

            Dim rowPOTCATG1 As DataRow = LookUp("POTCATG1", COST_CATGY_CODE)

            .DisplayLayout.Bands(0).AddNew()
            With .ActiveRow
                .Cells("MISC_CHG_CODE").Value = rowPOTCATG1.Item("MISC_CHG_CODE")
                .Cells("CTL_NO").Value = CTL_NO
                .Cells("PO_ORDER_NO").Value = PO_ORDER_NO
                .Cells("INV_MISC_CHG").Value = COST_ACT_PO
                .Update()
            End With
        End With

    End Sub

    Private Sub grdPOTLCST2_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTLCST2.InitializeLayout

    End Sub

    Private Sub grdPOTLCST2_InitializeRowsCollection(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowsCollectionEventArgs) Handles grdPOTLCST2.InitializeRowsCollection

    End Sub

    Private Sub chkFEOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkFEOnly.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_POTLCST2()
    End Sub


    Private Sub chkEnteredSince_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkEnteredSince.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_POTLCST2()
    End Sub

    Private Sub dteEnteredSince_ValueChanged(sender As System.Object, e As System.EventArgs) Handles dteEnteredSince.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_POTLCST2()
    End Sub

    Function Check_for_COST_COMPLETE(rowPOTLCST2 As DataRow, Optional update As Boolean = False) As Boolean
        Dim CTL_NO As String = rowPOTLCST2.Item("CTL_NO") & ""
        If CTL_NO <> "" Then
            Dim rowPOTLCST1 As DataRow = LookUp("POTLCST1", CTL_NO)
            If rowPOTLCST1 IsNot Nothing Then
                Dim PO_SHIMENT_NO As String = rowPOTLCST1.Item("PO_SHIPMENT_NO")
                Dim rowPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIMENT_NO)
                If rowPOTSHIP1.Item("COST_COMPLETE") = "1" Then
                    If update Then
                        ASCMAIN1.sql = "Update POTSHIP1 Set COST_COMPLETE = '0' where PO_SHIPMENT_NO = '" & PO_SHIMENT_NO & "'"
                        ASCDATA1.ExecuteSQL()
                    Else
                        If MsgBox("This entry impacts a Shipment which whose Costing has been Completed" _
                                  & vbCrLf & "OK to Proceed with this Entry?", MsgBoxStyle.YesNo, "Shipments which have been Completed Costed will be Re-Opened") = MsgBoxResult.No Then
                            Return False
                        Else
                            Return True
                        End If
                    End If
                Else
                    Return True
                End If
            End If
        End If
    End Function

    Private Sub txtINV_FREIGHT_GotFocus(sender As Object, e As EventArgs) Handles txtINV_FREIGHT.GotFocus
        txtINV_FREIGHT.SelectAll()
    End Sub
End Class