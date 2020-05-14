Public Class EDF865O1
    'Load_Events

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name

    Dim ORDR_GROUP_NO As String
    Dim EDI_OUTBOUND_DOC_NO As String
    Dim ORDR_CUST_PO As String      ' Customer's PO No
    Dim EDI_PURPOSE_CODE As String
    Dim ORDR_PICKED As Boolean

    Dim rowEDT865O1 As DataRow
    Dim rowARTCUST1 As DataRow      ' ARTCUST1 for the Sold-To
    Dim rowICTSTYL1 As DataRow
    Dim rowSOTORDR0 As DataRow
    Dim rowEDT850T1 As DataRow
    Dim rowEDTSYSIH As DataRow
    Dim sqlSOTORDR0 As String
    Dim SOTORDR0 As String
    Dim sqlSOTORDRS As String
    Dim EDI_DOC_SEQ_NO As String

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            sqlSOTORDR0 = "Select 'O' ORDR_TYPE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
          & ", SOTORDR0.CUST_DC_NO, SOTORDR0.ORDR_DEPT, EDT850T1.EDI_MERCH_TYPE, SOTORDR0.SALES_DIVISION_CODE, SOTORDR0.ORDR_DATE" & vbCrLf _
          & ", SOTORDR0.ORDR_SHIP_DATE,SOTORDR0. ORDR_CANCEL_DATE, SOTORDR0.ORDR_ORIG_SHIP_DATE, SOTORDR0.ORDR_ORIG_CANCEL_DATE" & vbCrLf _
          & ", SOTORDR0.WHSE_CODE, SOTORDR0.SREP_CODE" & vbCrLf _
          & ", SOTORDR0.ORDR_AMT, SOTORDR0.ORDR_AMT_OPEN, SOTORDR0.ORDR_AMT_PICK, SOTORDR0.ORDR_AMT_SHIP, SOTORDR0.ORDR_AMT_CANC" & vbCrLf _
          & ", SOTORDR0.ORDR_QTY, SOTORDR0.ORDR_QTY_OPEN, SOTORDR0.ORDR_QTY_PICK, SOTORDR0.ORDR_QTY_SHIP, SOTORDR0.ORDR_QTY_CANC" & vbCrLf _
          & ", SOTORDR0.ORDR_CNT, SOTORDR0.ORDR_CNT_OPEN, SOTORDR0.ORDR_CNT_PICK, EDT865O1.ORDR_CUST_PO EDI_PO" & vbCrLf _
          & " from SOTORDR0 " & vbCrLf _
          & " join EDT850T1 on EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO " & vbCrLf _
          & " left outer join (select distinct ORDR_CUST_PO from EDT865O1) EDT865O1 on  EDT865O1.ORDR_CUST_PO = sotordr0.ORDR_CUST_PO "
            ASCMAIN1.sql = sqlSOTORDR0 & " where EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO And SOTORDR0.CUST_CODE = ''"
            ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") X, ARTCUST1" _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
            SOTORDR0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add WAVE_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_TYPE, ORDR_GROUP_NO)")
            ASCMAIN1.sql = "Select * from " & SOTORDR0
            'Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "V", 2)
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "", 2)




            sqlSOTORDRS = "Select SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC, SOTORDR2.RANGE_STYLE_CODE" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY" & vbCrLf _
                & ", (SOTORDR2.ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_ALLO" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTORDR2.ORDR_RELEASE_AVAIL" & vbCrLf _
                & ", SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                & " from SOTORDR2,ICTCOLR1,SOTORDR1 " & vbCrLf _
                & " where ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_QTY_CANC > 0"
            ASCMAIN1.sql = sqlSOTORDRS & " and ROWNUM < 1 "
            Create_TDA(.Tables.Add, "SOTORDRS", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "select sotordr2.EDI_DTL_SEQ, sum(ORDR_QTY_ORIG) ORDR_QTY_ORIG, SUM(ORDR_QTY_OPEN) ORDR_QTY_OPEN, sum(ORDR_QTY_PICK) ORDR_QTY_PICK, sum(ORDR_QTY_CANC) ORDR_QTY_CANC from sotordr2, sotordr1" & vbCrLf _
                & " where sotordr2.ORDR_NO = sotordr1.ORDR_NO and ordr_group_no = :PARM1" & vbCrLf _
                & " group by EDI_DTL_SEQ "
            Create_TDA(.Tables.Add, "SOTORDRD", "**", 0, False, "V", 1)


            ASCMAIN1.sql = "Select EDT865O1.* from EDT865O1" & vbCrLf _
                & " where COMPANY_CODE = :PARM1 and EDI_OUTBOUND_DOC_NO = :PARM2"
            Create_TDA(.Tables.Add, "EDT865OX", "**", 0, False, "VV", 2)

            Create_TDA(.Tables.Add, "EDT865O1", "*", 2)

            'ASCMAIN1.sql = "Select EDT865O2.*" _
            '& " from EDT865O2"
            Create_TDA(.Tables.Add, "EDT865O2", "*", 3)
            .Tables("EDT865O2").Columns.Add("EDI_ITEM_SOLDOUT_FLAG")
            .Tables("EDT865O2").Columns.Add("EDI_QTY_CANC_CALC", GetType(System.Int64), "ISNULL(ORDR_QTY_OPEN,0) - ISNULL(EDI_QTY_OPEN,0)")

            Create_TDA(.Tables.Add, "EDT865O5", "*", 3)
            Create_TDA(.Tables.Add, "EDTSYSIH", "*", 2)

            'With .Tables.Add("SOTRSRVT")
            '    .Columns.Add("KEY", GetType(System.Int32))
            '    .Columns.Add("STATUS")
            '    .Columns.Add("QTY", GetType(System.Int32))
            '    .Columns.Add("AMT", GetType(System.Decimal))
            '    .PrimaryKey = New DataColumn() {.Columns("KEY")}
            'End With

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)
        End With

        grdEDT865OX.DataSource = dst.Tables("EDT865OX")
        grdEDT865O2.DataSource = dst.Tables("EDT865O2")
        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTORDRS.DataSource = dst.Tables("SOTORDRS")

        grdEDT865OX.DisplayLayout.UseFixedHeaders = True
        With grdEDT865OX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"EDI_OUTBOUND_DOC_NO", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        'grdEDT865O2.DisplayLayout.UseFixedHeaders = True
        'With grdEDT865O2.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"RSRV_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
        '        .Columns(COLUMN_NAME).Header.Fixed = True
        '    Next
        'End With

        With grdEDT865O2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"EDI_QTY_OPEN", "EDI_ITEM_SOLDOUT_FLAG"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"ORDR_QTY_OPEN", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    'ElseIf New String() {"rsrv_amt", "rsrv_amt_open", "rsrv_amt_allo", "rsrv_amt_used", "rsrv_amt_canc"}.Contains(gcol.Key) Then
                    '    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    '    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    'ElseIf New String() {"style_code", "style_desc", "color_code", "color_desc"}.Contains(gcol.Key) Then
                    '    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    '    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With

        'With grdSOTORDRS.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}
        '        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
        '            If COLUMN_NAME = "STYLE_CODE" Then .Columns(COLUMN_NAME).Header.Fixed = True
        '        Else
        '            .Columns(COLUMN_NAME).Header.Fixed = True
        '        End If
        '    Next
        'End With

        Create_Summary(grdEDT865OX, "EDI_OUTBOUND_DOC_NO", "Count")



        Create_Summary(grdSOTORDRS, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRS, New String() {"ORDR_QTY", "ORDR_AMT", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})


        'Create_Summary(grdEDT865O2, "RSRV_LNO", "Count")
        'Create_Summary(grdEDT865O2, New String() {"RSRV_QTY", "RSRV_QTY_OPEN", "RSRV_QTY_ALLO", "RSRV_QTY_USED", "RSRV_QTY_CANC", "RSRV_AMT"})

        'With dst.Tables("SOTRSRVT").Rows
        '    .Add(New Object() {1, "Rsrv", 0, 0})
        '    .Add(New Object() {2, "Open", 0, 0})
        '    .Add(New Object() {3, "Allo", 0, 0})
        '    .Add(New Object() {4, "Used", 0, 0})
        '    .Add(New Object() {5, "Canc", 0, 0})
        'End With
        'Sort_grdColumns(grdSOTRSRVT, "KEY", True)

        Show_Filter(grdSOTORDR0, True)

        Show_Filter(grdEDT865OX, True)
        grdEDT865OX.DisplayLayout.GroupByBox.Hidden = False

        'SplitContainer1.Panel2Collapsed = True
        tabInfo.Tabs("Usage").Visible = False

        '  Check_InquiryMode()

    End Sub

    'Sub Check_InquiryMode()
    '    InquiryMode = (MENU_ITEM_OBJECT = "SOFRSRVI")
    'End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        'ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                        'If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                        '    EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                        'End If
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If Absx1.txtFor("ORDR_GROUP_NO").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify an Order Group No"
                Else
                    rowSOTORDR0 = LookUp("SOTORDR0", Absx1.txtFor("ORDR_GROUP_NO").Text)
                    If rowSOTORDR0 IsNot Nothing Then
                        ORDR_GROUP_NO = Absx1.txtFor("ORDR_GROUP_NO").Text
                        ORDR_CUST_PO = rowSOTORDR0.Item("ORDR_CUST_PO") & ""
                    Else
                        EMsg &= vbCr & "No Record of Order Group " & Absx1.txtFor("ORDR_GROUP_NO").Text
                    End If
                End If

                If Absx1.txtFor("EDI_PURPOSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a transaction type"
                Else
                    EDI_PURPOSE_CODE = Absx1.txtFor("EDI_PURPOSE_CODE").Text
                End If

                If EMsg = "" Then

                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                End If

            Case "View"

                CUST_CODE = ""
                EDI_OUTBOUND_DOC_NO = ""
                ORDR_GROUP_NO = ""

                If Absx1.txtFor("EDI_OUTBOUND_DOC_NO").Text = "" Then
                    EMsg &= vbCr & "No Reservation No Specified"
                Else
                    EDI_OUTBOUND_DOC_NO = Absx1.txtFor("EDI_OUTBOUND_DOC_NO").Text
                    rowEDT865O1 = LookUp("EDT865O1", EDI_OUTBOUND_DOC_NO)
                    If rowEDT865O1 Is Nothing Then
                        EMsg &= vbCr & "No Record of 865 Doc No " & EDI_OUTBOUND_DOC_NO
                    Else
                        CUST_CODE = rowEDT865O1.Item("CUST_CODE")
                        ORDR_GROUP_NO = rowEDT865O1.Item("ORDR_GROUP_NO")
                    End If
                End If

            Case "Update"
                If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                    Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                Else
                    If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                     > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                    End If
                End If

                'If Absx1.txtFor("TERM_CODE").Text = "" Then
                '    EMsg &= vbCr & "Terms are required"
                'Else
                '    If LookUp("TATTERM1", Absx1.txtFor("TERM_CODE").Text) Is Nothing Then
                '        EMsg &= vbCr & "Invalid Terms"
                '    End If
                'End If

                If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                    EMsg &= vbCr & "Customer PO is required"
                End If

                'If grdEDT865O2.Rows.Count = 0 Then
                '    EMsg &= vbCr & "No Items on Reservation"
                'Else
                '    If Val(dst.Tables("SOTRSRV2").Compute("COUNT(RSRV_LNO)", "RSRV_QTY > 0") & "") = 0 Then
                '        EMsg &= vbCr & "No Items on Reservation with Qty >0"
                '    End If
                'End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
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

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

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
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
            End With

            .Groups("Totals").Visible = ScreenMode
        End With

        'lblStatus.Visible = ScreenMode

        tab865.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), InquiryMode Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")))
        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))

        If ScreenMode Then
            If EntryMode = "V" Then
                grdEDT865O2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdEDT865O2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdEDT865O2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdEDT865O2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No ' FixedAddRowOnTop
                grdEDT865O2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdEDT865O2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

                If EntryMode <> "E" Then
                    grdEDT865O2.DisplayLayout.Bands(0).Columns("X").Hidden = True
                Else
                    grdEDT865O2.DisplayLayout.Bands(0).Columns("X").Hidden = False
                End If
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("EDI_OUTBOUND_DOC_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("ORDR_GROUP_NO").Text = ""

        CUST_CODE = ""
        EDI_OUTBOUND_DOC_NO = ""
        ORDR_GROUP_NO = ""
        ORDR_PICKED = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"EDT865O1", "EDT865O2", "EDT865O5", "EDTSYSIH"} ', "SOTORDRS"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Load_EDT865OX()
        Load_SOTORDR0()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            EDI_OUTBOUND_DOC_NO = ASCMAIN1.Next_Control_No("EDI_OUTBOUND_DOC_NO")
            rowSOTORDR0 = LookUp("SOTORDR0", ORDR_GROUP_NO)
            EDI_DOC_SEQ_NO = rowSOTORDR0.Item("EDI_DOC_SEQ_NO") & ""
            ORDR_PICKED = False
            If rowSOTORDR0.Item("ORDR_QTY_PICK") > 0 Then
                ORDR_PICKED = True
            End If


            rowEDT850T1 = LookUp("EDT850T1", EDI_DOC_SEQ_NO)
            rowEDT865O1 = dst.Tables("EDT865O1").NewRow
            With rowEDT865O1
                .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("EDI_TP_QUAL") = rowEDT850T1.Item("EDI_TP_QUAL")
                .Item("EDI_TP_ID") = rowEDT850T1.Item("EDI_TP_ID")
                .Item("CUST_CODE") = CUST_CODE
                .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                .Item("EDI_PURPOSE_CODE") = EDI_PURPOSE_CODE
                .Item("ORDR_PO_DATE") = rowEDT850T1.Item("EDI_PO_DATE")
                .Item("REQUEST_DATE") = DATETIME_STAMP
                .Item("TERM_CODE") = "" 'rowSOTORDR0.Item("TERM_CODE")
                .Item("AS_OF_DATE") = DATETIME_STAMP
                '.Item("ORDR_AMT") = rowEDT850T1.Item("")
                .Item("EDI_ARRIVAL_DATE") = rowEDT850T1.Item("EDI_RECEIVED_DATE")
                '.Item("EDI_ACK_TYPE") = rowEDT850T1.Item("")
                .Item("EDI_SUPPLIER_NO") = rowEDT850T1.Item("EDI_SUPPLIER_NO")
                .Item("EDI_DEPT_NO") = rowEDT850T1.Item("EDI_DEPARTMENT")
                .Item("EDI_REF_CODE") = rowEDT850T1.Item("EDI_MERCH_TYPE")
                .Item("EDI_REF_CODE2") = rowEDT850T1.Item("EDI_PROMOTION")
                .Item("ORDR_SHIP_DATE") = rowSOTORDR0.Item("ORDR_SHIP_DATE")
                .Item("ORDR_CANCEL_DATE") = rowSOTORDR0.Item("ORDR_CANCEL_DATE")

                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                'Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                'If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
                'If WHSE_CODE = "" Then WHSE_CODE = ""
                '.Item("WHSE_CODE") = WHSE_CODE
                '.Item("RSRV_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
            End With
            dst.Tables("EDT865O1").Rows.Add(rowEDT865O1)

            rowEDTSYSIH = dst.Tables("EDTSYSIH").NewRow
            With rowEDTSYSIH
                .Item("EDI_OUR_ID") = rowEDT850T1.Item("EDI_OUR_ID")
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                .Item("EDI_APPLICATION_ID") = "OPC"
                .Item("EDI_TP_ID") = rowEDT850T1.Item("EDI_TP_ID")
                .Item("EDI_PROCESS_IND") = "1"
            End With
            dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

            Fill_Records("SOTORDRD", ORDR_GROUP_NO)

            ASCMAIN1.sql = "Select * from edt850t2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim EDI_DTL_SEQ As Int32 = Val(row.Item("EDI_DTL_SEQ") & "")
                Dim rowEDT865O2 As DataRow = dst.Tables("EDT865O2").NewRow
                With rowEDT865O2
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                    .Item("EDI_TOTAL_QTY") = row.Item("EDI_TOTAL_QTY")
                    '.Item("EDI_UOM") = "EDI_UOM"
                    .Item("EDI_PRICE") = row.Item("EDI_PRICE")
                    '.Item("EDI_PO4_QTY") = "EDI_PO4_QTY"
                    '.Item("EDI_PO4_INNER") = "EDI_PO4_INNER"
                    .Item("EDI_PO4_UOM") = row.Item("EDI_PO4_UOM")
                    .Item("EDI_ITEM") = row.Item("EDI_STYLE")
                    .Item("EDI_UPC") = row.Item("EDI_UPC")
                    .Item("EDI_SKU") = row.Item("EDI_SKU")
                    '.Item("EDI_GTIN") = "EDI_GTIN"
                    .Item("EDI_ITEM_DESC") = row.Item("EDI_STYLE_NAME")
                    '.Item("EDI_PO_LNO") = "EDI_PO_LNO"
                    '.Item("EDI_PRICE_ACTUAL") = "EDI_PRICE_ACTUAL"
                    Dim rowSOTORDRD As DataRow = dst.Tables("SOTORDRD").Rows.Find(EDI_DTL_SEQ)
                    .Item("ORDR_QTY_OPEN") = IIf(ORDR_PICKED, Val(rowSOTORDRD.Item("ORDR_QTY_ORIG") & ""), Val(rowSOTORDRD.Item("ORDR_QTY_OPEN") & ""))
                    .Item("ORDR_QTY_CANC") = Val(rowSOTORDRD.Item("ORDR_QTY_CANC") & "")
                    .Item("EDI_QTY_OPEN") = IIf(ORDR_PICKED, Val(rowSOTORDRD.Item("ORDR_QTY_PICK") & ""), Val(rowSOTORDRD.Item("ORDR_QTY_OPEN") & ""))
                    '.Item("EDI_QTY_PICK") = "EDI_QTY_PICK"
                    .Item("EDI_QTY_CANC") = Val(rowSOTORDRD.Item("ORDR_QTY_CANC") & "")
                    .Item("EDI_DIMENSION") = "EDI_DIMENSION"
                    .Item("EDI_ITEM_SOLDOUT_FLAG") = "" ' "EDI_ITEM_SOLDOUT"
                    .Item("EDI_ITEM_CHANGE_TYPE") = "" ' "EDI_ITEM_CHANGE_TYPE"
                End With
                dst.Tables("EDT865O2").Rows.Add(rowEDT865O2)
            Next

            ASCMAIN1.sql = "Select * from edt850t5 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim EDI_ADR_SEQ As Int32 = Val(row.Item("EDI_ADR_SEQ") & "")
                Dim rowEDT865O5 As DataRow = dst.Tables("EDT865O5").NewRow
                With rowEDT865O5
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                    .Item("EDI_ADDR_TYPE") = row.Item("EDI_ADDR_TYPE")
                    .Item("EDI_CUST_NAME_ADR") = row.Item("EDI_CUST_NAME_ADR")
                    .Item("EDI_ADDRESS1") = row.Item("EDI_ADDRESS1")
                    .Item("EDI_ADDRESS2") = row.Item("EDI_ADDRESS2")
                    .Item("EDI_ADDRESS3") = ""
                    .Item("EDI_CITY") = row.Item("EDI_CITY")
                    .Item("EDI_STATE") = row.Item("EDI_STATE")
                    .Item("EDI_ZIPCODE") = row.Item("EDI_ZIPCODE")
                    .Item("EDI_COUNTRY") = row.Item("EDI_COUNTRY")
                    .Item("EDI_ADDR_CODE") = row.Item("EDI_ADDR_CODE")
                    .Item("EDI_ADDR_CODE_QUAL") = row.Item("EDI_ADDR_CODE_QUAL")
                End With
                dst.Tables("EDT865O5").Rows.Add(rowEDT865O5)
            Next

        Else
            rowEDT865O1 = Fill_Record("EDT865O1", New String() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO})
            CUST_CODE = rowEDT865O1.Item("CUST_CODE")
            ORDR_CUST_PO = rowEDT865O1.Item("ORDR_CUST_PO")
            ORDR_GROUP_NO = rowEDT865O1.Item("ORDR_GROUP_NO")
            rowSOTORDR0 = LookUp("SOTORDR0", ORDR_GROUP_NO)
            EDI_DOC_SEQ_NO = rowSOTORDR0.Item("EDI_DOC_SEQ_NO")

            Fill_Records("EDT865O2", New String() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO})
            '   Sort_grdColumns(grdEDT865O2, "RSRV_LNO")

        End If


        CUST_CODE = rowEDT865O1.Item("CUST_CODE")
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)


        lblINIT_DATE.Text = "Entered on " & Format(rowEDT865O1.Item("INIT_DATE"), "MM/dd/yyyy")

        With grdEDT865O2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"EDI_QTY_OPEN", "EDI_ITEM_SOLDOUT_FLAG"}.Contains(gcol.Key) And EDI_PURPOSE_CODE = "19" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next
        End With


        'If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
        '    With grdEDT865O2.DisplayLayout.Override
        '        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        '        .AllowUpdate = DefaultableBoolean.True
        '        .AllowDelete = DefaultableBoolean.True
        '    End With
        '    grdEDT865O2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Set_Read_Only(splHeader, False)
        'Else
        '    With grdEDT865O2.DisplayLayout.Override
        '        .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '        .AllowUpdate = DefaultableBoolean.False
        '        .AllowDelete = DefaultableBoolean.False
        '    End With
        '    grdEDT865O2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        '    Set_Read_Only(splHeader, True)
        'End If

        'Load_SOTORDRS()

        Display_Totals()
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
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, EDI_OUTBOUND_DOC_NO)
        For Each TABLE_NAME As String In New String() _
            {"SOTRSRV1", "SOTRSRV2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        If EntryMode <> "N" Then Delete_Records()

        If EntryMode = "N" Then
            rowEDT865O1.Item("ORDR_ORIG_SHIP_DATE") = rowEDT865O1.Item("ORDR_SHIP_DATE")
            rowEDT865O1.Item("ORDR_ORIG_CANCEL_DATE") = rowEDT865O1.Item("ORDR_CANCEL_DATE")
        End If

        Dim sqlwhere As String

        If EDI_PURPOSE_CODE = "19" Then
            sqlwhere = "EDI_QTY_CANC_CALC = 0 AND ISNULL(EDI_ITEM_SOLDOUT_FLAG,'0') <> '1'"
        Else
            sqlwhere = "EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'"
        End If


        ASCDATA1.DeleteRows("EDT865O2", sqlwhere)

        For Each row As DataRow In dst.Tables("EDT865O2").Select()
            If Val(row.Item("EDI_ITEM_SOLDOUT_FLAG") & "") = "1" Then
                row.Item("EDI_ITEM_SOLDOUT") = "PSO"
                row.Item("EDI_ITEM_CHANGE_TYPE") = "DI"
            Else
                If Val(row.Item("ORDR_QTY_OPEN") & "") = Val(row.Item("EDI_QTY_CANC_CALC") & "") And (rowEDT865O1.Item("EDI_REF_CODE") & "" <> "SAMS FXD") Then
                    row.Item("EDI_ITEM_CHANGE_TYPE") = "DI"
                Else
                    row.Item("EDI_ITEM_CHANGE_TYPE") = "QD"
                    row.Item("EDI_QTY_CANC") = Val(row.Item("EDI_QTY_CANC_CALC") & "")
                End If
            End If
        Next

        'INIT_LAST("SOTRSRV1", False, , True)
        Dim sqldelete As String = "EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'"
        Update_Record_TDA("EDT865O1", sqldelete)
        Update_Record_TDA("EDT865O2", sqldelete)
        Update_Record_TDA("EDT865O5", sqldelete)
        Update_Record_TDA("EDTSYSIH", sqldelete)

        'Dependent_Updates(1, EDI_OUTBOUND_DOC_NO)

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "EDI_OUTBOUND_DOC_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTRSRV1.RSRV_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTRSRV1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTRSRV1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If
        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("EDI_OUTBOUND_DOC_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTRSRV1"
            E.COLUMN_NAME = "EDI_OUTBOUND_DOC_NO"
            E.CODE_VALUE = Absx1.txtFor("EDI_OUTBOUND_DOC_NO").Text
            E.DESC_VALUE = "Reservation"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRSRV1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT865OX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdEDT865O2, "BB", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTORDRS, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")

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
            Case "grdSOTRSRV2"
                tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

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

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_EDT865OX()
                    Load_SOTORDR0()
                End If

            Case "ORDR_CUST_PO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode _
                       And Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "EDI_OUTBOUND_DOC_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_EDT865OX()
                    Load_SOTORDR0()

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_EDT865OX()
                Load_SOTORDR0()
            Case "EDI_OUTBOUND_DOC_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_EDT865OX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE = "" Then
            ASCMAIN1.sql = "Select * from EDT865O1"
            Fill_Records("EDT865OX", "", , ASCMAIN1.sql)
            grdEDT865OX.Text = "865 History"
            Sort_grdColumns(grdEDT865OX, "EDI_OUTBOUND_DOC_NO".ToLower)
        Else
            ASCMAIN1.sql = "Select * from EDT865O1 where CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("EDT865OX", "", , ASCMAIN1.sql)
            grdEDT865OX.Text = "865 History with " & CUST_CODE
            Sort_grdColumns(grdEDT865OX, "EDI_OUTBOUND_DOC_NO".ToLower)
        End If
        grdEDT865OX.Visible = True
    End Sub

    Sub Load_SOTORDR0(Optional PARM1 As String = "", Optional CUST_CODE As String = "")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Order Summary", "")

        CUST_CODE = Absx1.txtFor("CUST_CODE").Text

        If CUST_CODE <> "" Then ' ScreenMode Then
            ASCMAIN1.sql = sqlSOTORDR0
            Dim sqlw As String = " where EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf

            grdSOTORDR0.Text = "Order Groups for " & CUST_CODE ' & "; Status: " & optOrders.Text

            ASCMAIN1.sql &= sqlw

            'If (optOrders.Value = "A" Or optOrders.Value = "O" Or optOrders.Value = "OP" Or optOrders.Value = "C") And chkReservations.Checked Then
            '    ASCMAIN1.sql &= Replace(sqlReservations, " group by ", "   and SOTRSRV1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf & " group by ")
            'End If

        Else
            ASCMAIN1.sql = sqlSOTORDR0
            PARM1 = Replace(Replace(PARM1, ";", ""), "'", "")

            Dim sqlORDR_STATUS As String = ""

            grdSOTORDR0.Text = "Orders which are either Open or In Pick"
            ASCMAIN1.sql &= " where EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO and (SOTORDR0.ORDR_CNT_OPEN <> 0 or SOTORDR0.ORDR_CNT_PICK <> 0)"

        End If

        ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY, NULL WAVE_NO from (" & ASCMAIN1.sql & ") X,ARTCUST1" _
             & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
        'Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Delete from " & SOTORDR0)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0 & " " & ASCMAIN1.sql)

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select ORDR_GROUP_NO, MIN (WAVE_NO) WAVE_NO" & vbCrLf _
                & "   from SOTSHIP1 where ORDR_GROUP_NO in " & vbCrLf _
                & "    (Select ORDR_GROUP_NO from " & SOTORDR0 & " where ORDR_TYPE = 'O')" & vbCrLf _
                & "   group by ORDR_GROUP_NO;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set WAVE_NO = R1.WAVE_NO where ORDR_TYPE = 'O' and ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        Fill_Records("SOTORDR0")

        Setup_SOTORDR0()
        Sort_grdColumns(grdSOTORDR0, "ORDR_GROUP_NO".ToLower)

        grdSOTORDR0.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub


    Sub Setup_SOTORDR0()

        'If grdSOTORDR0.ActiveRow Is Nothing OrElse Not grdSOTORDR0.ActiveRow.IsDataRow Then
        '    tabDetails.Visible = False
        'Else
        '    tabDetails.Visible = True
        '    ORDR_GROUP_NO = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value
        '    Dim ORDR_CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
        '    Dim ORDR_TYPE As String = grdSOTORDR0.ActiveRow.Cells("ORDR_TYPE").Value
        '    'ASCMAIN1.Progress("Now Setting up Details")
        '    If ORDR_TYPE = "R" Then
        '        chkShowSelectedOrder.Checked = False
        '        chkShowSelectedOrder.Enabled = False
        '    Else
        '        chkShowSelectedOrder.Enabled = True
        '    End If
        '    EnforceConstraints(False)
        '    dst.Tables("SOTPICK2").Rows.Clear()
        '    dst.Tables("SOTSHIP2").Rows.Clear()
        '    If ORDR_TYPE = "O" Then
        '        Fill_Records("SOTORDR1", ORDR_GROUP_NO)
        '        If tabDetails.SelectedTab.Key = "All Orders" Then
        '        Else
        '            Sort_grdColumns(grdSOTORDR1, "ORDR_NO")
        '            grdSOTORDR1.Text = "Sales Orders for Order Group " & ORDR_GROUP_NO
        '        End If

        '        Fill_Records("SOTPICK1", ORDR_GROUP_NO)
        '        Sort_grdColumns(grdSOTPICK1, "PICK_NO")
        '        grdSOTPICK1.Text = "Pick Tickets for Order Group " & ORDR_GROUP_NO

        '        Fill_Records("SOTSHIP1", ORDR_GROUP_NO)
        '        Sort_grdColumns(grdSOTSHIP1, "SHIP_BOL_NO")
        '        grdSOTSHIP1.Text = "Shipments for Order Group " & ORDR_GROUP_NO
        '    Else
        '        dst.Tables("SOTORDR1").Rows.Clear()
        '        dst.Tables("SOTPICK1").Rows.Clear()
        '        dst.Tables("SOTSHIP1").Rows.Clear()
        '    End If
        '    EnforceConstraints(True)
        '    Load_SOTORDRS()
        '    ' ASCMAIN1.Progress("")
        'End If

    End Sub


    Sub Load_SOTORDRS()

        Setup_SOTORDRS()

        'With grdSOTORDRS.DisplayLayout.Bands(0)
        '    If ORDR_TYPE = "O" Then
        '        .Columns("ORDR_QTY_SHIP").Header.Caption = "#Ship"
        '        ' .Columns("ORDR_AMT_SHIP").Header.Caption = "$Ship"
        '    Else
        '        .Columns("ORDR_QTY_SHIP").Header.Caption = "#Used"
        '        ' .Columns("ORDR_AMT_SHIP").Header.Caption = "$Used"
        '    End If
        'End With

        '  Setup_Summary()
    End Sub

    Sub Setup_SOTORDRS()

        If grdSOTORDR0.ActiveRow Is Nothing Then Exit Sub

        ASCMAIN1.Progress("Now Getting Style Details")

        Dim GROUP_NO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
        Dim CUST_PO As String = grdSOTORDR0.ActiveRow.Cells("ORDR_CUST_PO").Value & ""
        Dim sql As String = ""

        sql = sqlSOTORDRS & " and SOTORDR1.ORDR_GROUP_NO = '" & GROUP_NO & "'"
        grdSOTORDRS.Text = "Cancel Style Summary for Order Group " & GROUP_NO & ", Customer PO " & CUST_PO

        'If Not chkShowSelectedOrder.Checked Then
        '    sql = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' group by ")
        '    grdSOTORDRS.Text = "Style Summary for Order Group " & ORDR_GROUP_NO & ", Customer PO " & ORDR_CUST_PO
        'Else
        '    Dim ORDR_NO As String = grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value
        '    Dim CUST_STORE_NO As String = grdSOTORDRX.ActiveRow.Cells("CUST_STORE_NO").Value
        '    sql = Replace(sqlSOTORDRS, " group by ", " and SOTORDR1.ORDR_NO = '" & ORDR_NO & "' group by ")
        '    grdSOTORDRS.Text = "Style Details for Order No " & ORDR_NO & ", Customer PO " & ORDR_CUST_PO & ", Store No " & CUST_STORE_NO
        'End If

        For Each COLUMN_NAME As String In New String() _
                {"STYLE_CODE", "COLOR_CODE", "CUST_UPC", "RANGE_STYLE_CODE",
                 "CUST_STYLE_CODE", "CUST_COLOR_CODE", "CUST_SIZE_CODE", "CUST_SKU"}

            grdSOTORDRS.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False ' Not Absx1.chkFor("SHOW_" & COLUMN_NAME).Checked

            If Not True Then '  Not Absx1.chkFor("SHOW_" & COLUMN_NAME).Checked Then
                sql = Replace(sql, "SOTORDR2." & COLUMN_NAME, "NULL " & COLUMN_NAME, , 1)
                sql = Replace(sql, "SOTORDR2." & COLUMN_NAME, "NULL")
                If COLUMN_NAME = "STYLE_CODE" Then
                    sql = Replace(sql, "SOTORDR2.STYLE_DESC", "NULL " & "STYLE_DESC", , 1)
                    sql = Replace(sql, "SOTORDR2.STYLE_DESC", "NULL")
                End If
                If COLUMN_NAME = "COLOR_CODE" Then
                    sql = Replace(sql, "ICTCOLR1.COLOR_DESC", "NULL " & "COLOR_DESC", , 1)
                    sql = Replace(sql, "ICTCOLR1.COLOR_DESC", "NULL")
                End If
            End If
        Next

        sql = Replace(sql, "ICTCOLR1.COLOR_CODE (+) = NULL", "ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE")
        'If ORDR_TYPE = "R" Then
        '    sql = Replace(sql, "SOTORDR2.QTY_PICK", "0")
        '    sql = Replace(sql, "SOTORDR2.QTY_SHIP", "SOTORDR2.QTY_USED")
        '    sql = Replace(Replace(sql, "SOTORDR1", "SOTRSRV1"), "SOTORDR2", "SOTRSRV2")
        'End If

        Fill_Records("SOTORDRS", "", True, sql)
        Sort_grdColumns(grdSOTORDRS, "ORDR_CUST_PO, CUST_STORE_NO, STYLE_CODE, COLOR_CODE, RANGE_STYLE_CODE, CUST_STYLE_CODE, CUST_COLOR_CODE, CUST_SKU")

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Dim extra_decimals As Boolean = False
            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("", "")
                Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTORDRS.Item("ORDR_UNIT_PRICE") & "")
                If ORDR_UNIT_PRICE <> Val(Format(ORDR_UNIT_PRICE, "#.00")) Then
                    extra_decimals = True
                    Exit For
                End If
            Next
            If extra_decimals Then
                grdSOTORDRS.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = "#,##0.0000"
            Else
                grdSOTORDRS.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Format = "#,##0.00"
            End If
        End If

        ASCMAIN1.Progress("")
    End Sub





    Private Sub grdEDT865OX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT865OX.DoubleClickRow
        Absx1.txtFor("EDI_OUTBOUND_DOC_NO").Text = e.Row.Cells("EDI_OUTBOUND_DOC_NO").Value
        Click_Command("View")
    End Sub

    Sub Dependent_Updates(S As Integer, EDI_OUTBOUND_DOC_NO As String)

        ' update header import table

        'ASCDATA1.ExecuteSP("ICPSTAT2", "VVVNNNNNN", _
        '           New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, _
        '                         0, 0, 0, _
        '                         QTY, 0, 0}, _
        '           New String() {"STYLE_CODE_IN", "COLOR_CODE_IN", "WHSE_CODE_IN", _
        '                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in", _
        '                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})

    End Sub

    Sub Display_Totals()

        'Dim KEY As Int32 = 0
        'For Each SFX As String In New String() {"", "OPEN", "ALLO", "USED", "CANC"}
        '    If SFX <> "" Then SFX = "_" & SFX
        '    KEY += 1
        '    Dim rowSOTRSRVT As DataRow = dst.Tables("SOTRSRVT").Rows.Find(KEY)
        '    rowSOTRSRVT.Item("QTY") = Val(dst.Tables("SOTRSRV2").Compute("SUM(RSRV_QTY" & SFX & ")", "") & "")
        '    rowSOTRSRVT.Item("AMT") = Val(dst.Tables("SOTRSRV2").Compute("SUM(RSRV_AMT" & SFX & ")", "") & "")
        'Next
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdEDT865O2.ActiveRow
            Select Case COLUMN_NAME
                'Case "STYLE_CODE"
                '    Dim STYLE_CODE As String = ""
                '    If Trim(.Cells("STYLE_CODE").Value & "") <> "" Then
                '        STYLE_CODE = Validate_Style(.Cells("STYLE_CODE").Value & "")
                '    End If
                '    Cancel = (STYLE_CODE = "")

                'Case "RSRV_QTY"
                '    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                '        Cancel = True
                '        Exit Sub
                '    End If
                '    If Trim(.Cells("RSRV_QTY").Value & "") = "" Then
                '        MsgBox("Qty Not Specified", vbOKOnly, "Cannot Update Record")
                '        Cancel = True
                '        grdEDT865O2.ActiveCell = grdEDT865O2.ActiveRow.Cells("RSRV_QTY")
                '        Exit Sub
                '    End If
                '    If Val(.Cells("RSRV_QTY").Value & "") < 0 Then
                '        MsgBox("Qty May Not be Negative", vbOKOnly, "Invalid Quantity")
                '        Cancel = True
                '    End If
            End Select
        End With
    End Sub


#Region "grdSOTRSRV2"

    Private Sub grdSOTRSRV2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDT865O2.AfterCellUpdate
        With grdEDT865O2.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    'Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value)
                    'If STYLE_CODE <> "" Then
                    '    .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")

                    'End If

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & ""
                    If COLOR_CODE <> "" Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        If rowICTCOLR1 IsNot Nothing Then
                            .Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

                Case "RSRV_QTY"
                    .Cells("RSRV_QTY_OPEN").Value = .Cells("RSRV_QTY").Value

            End Select
        End With
    End Sub

    Private Sub grdSOTRSRV2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdEDT865O2.AfterRowActivate

    End Sub

    Private Sub grdSOTRSRV2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdEDT865O2.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTRSRV2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdEDT865O2.BeforeCellUpdate

    End Sub

    Private Sub grdSOTRSRV2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdEDT865O2.BeforeExitEditMode
        If grdEDT865O2.ActiveCell IsNot Nothing Then
            With grdEDT865O2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub


    Private Sub grdSOTRSRV2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdEDT865O2.BeforeRowUpdate

        Validate_Columns("STYLE_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("COLOR_CODE", e.Cancel)
        End If
        If Not e.Cancel Then
            Validate_Columns("RSRV_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("EDI_OUTBOUND_DOC_NO").Value = EDI_OUTBOUND_DOC_NO
            Dim RSRV_LNO As Int64 = Val(dst.Tables("SOTRSRV2").Compute("MAX(RSRV_LNO)", "") & "") + 1
            e.Row.Cells("RSRV_LNO").Value = RSRV_LNO
        End If
    End Sub

    Private Sub grdSOTRSRV2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDT865O2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE IN (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE ='" & grdEDT865O2.ActiveRow.Cells("STYLE_CODE").Value & "')"
                    grdClickCellButton(grdEDT865O2, sql_where)
            End Select
        End With

    End Sub
#End Region

    Sub Load_Events()

    End Sub

    Private Sub grdSOTORDR0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR0.AfterRowActivate
        Load_SOTORDRS()
    End Sub

    Private Sub grdSOTORDR0_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDR0.DoubleClickRow
        If e.Row.IsDataRow Then
            If Not ScreenMode Then
                Dim ORDR_GROUP_NO As String = e.Row.Cells("ORDR_GROUP_NO").Value & ""
                Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
                Absx1.txtFor("ORDR_GROUP_NO").Text = ORDR_GROUP_NO
                Click_Command("New")
                'For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                '    If grow.Cells("ORDR_GROUP_NO").Value = ORDR_GROUP_NO Then
                '        grdSOTORDR0.ActiveRow = grow
                '        grdSOTORDR0.DisplayLayout.RowScrollRegions(0).FirstRow = grow
                '    End If
                'Next
            End If
        End If
    End Sub

    Private Sub grdSOTORDR0_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR0.InitializeRow
        With e.Row.Cells("EDI_PO")
            If IsDBNull(.Value) Then
                e.Row.Appearance.BackColor = System.Drawing.Color.Empty
            Else
                e.Row.Appearance.BackColor = System.Drawing.Color.LightGreen
            End If
        End With
    End Sub

End Class