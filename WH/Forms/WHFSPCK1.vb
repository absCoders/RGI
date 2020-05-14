Public Class WHFSPCK1
    Dim WHSE_CODE As String
    Dim LP_CODE As String
    Dim rowICTWHSE1 As DataRow
    Dim rowWHTTPLP1 As DataRow

    Dim WHTSTYLX As String

    'NEED TO MT AGAINST REL & DE-REL FOR THE WHSE SELECTED
    'NEED TO MT FOR SENDITEMS
    'DO NOT ALLOW PT PRINT FOR A LP WHSE
    'CHECK MT ON DESIGN RECALL OF SHIPMENTS/PICK TICKETS
    ' CHECK EVENT PROCEDURE FIRING WHEN CLICKING CANCEL - WHEN CYCLING MODES TO FALSE, DONT WANT TO LOAD_SOTSHIPX
    Dim SOTSHIPX As String = ""
    Dim Shipments As Integer = 0
    Dim LP_XNO As String
    Dim ASW As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                WHTSTYLX = TAC.WHCMAIN1.Prepare_WHTSTYLX("", "", True)
            End If

            ASCMAIN1.sql = "Select SHIP_BOL_NO, '1' SEL, '1' EDI856, '1' SHIP_CART_REQD from SOTSHIP1 where ROWNUM < 1"
            SOTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_ADDR_TYPE_ST VARCHAR2(2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add CUST_CODE VARCHAR2(10)")

            ASCMAIN1.sql = "Select SOTSHIP1.*, SOTSHIPX.SEL, SOTSHIPX.EDI856" & vbCrLf _
                & ", SOTSHIPX.ORDR_NO, SOTSHIP1.SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST" & vbCrLf _
                & ", WHTSHIP1.LP_XNO LP_XNO_XMIT" & vbCrLf _
                & ", SOTORDR0.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
                & ", SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & " from " & SOTSHIPX & " SOTSHIPX,SOTSHIP1,SOTORDR0,ARTCUST1,WHTSHIP1" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and WHTSHIP1.LP_XNO (+) = SOTSHIP1.LP_XNO" & vbCrLf _
                & "   and WHTSHIP1.SHIP_BOL_NO (+) = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'"
            Create_TDA(.Tables.Add("SOTSHIPX"), SOTSHIPX, "**", 0, True, "", 1, "SEL")
            '.Tables("SOTSHIPX").Columns.Add("SEL")


            ASCMAIN1.sql = "Select ICTWHSE1.WHSE_CODE, ICTWHSE1.WHSE_DESC, ICTWHSE1.LP_CODE" & vbCrLf _
                & ", X.SHIPS" & vbCrLf _
                & "  from ICTWHSE1" & vbCrLf _
                & ", (Select SOTSHIP1.WHSE_CODE, Count (*) SHIPS" & vbCrLf _
                & "  from SOTSHIP1" & vbCrLf _
                & " where SOTSHIP1.SHIP_STATUS = 'P' and SOTSHIP1.LP_STATUS = '0'" & vbCrLf _
                & " group by SOTSHIP1.WHSE_CODE) X" & vbCrLf _
                & " where ICTWHSE1.LP_CODE is Not Null" & vbCrLf _
                & "   and X.WHSE_CODE (+) = ICTWHSE1.WHSE_CODE"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False)
            .Tables("ICTWHSEX").Columns("SHIPS").DataType = GetType(System.Int32)


            ASCMAIN1.sql = "Select * from SOTPICK1 where SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
                & ", SUM (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC) PICK_QTY_CANC" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK) PICK_QTY_BACK" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK_REL) PICK_QTY_BACK_REL" _
                & ", MAX (SOTORDR2.STYLE_DESC) STYLE_DESC" _
                & ", SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.INNER_PACK_QTY, SOTORDR2.CUST_STYLE_CODE" _
                & ", SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" _
                & ", MAX(ICTCOLR1.COLOR_DESC) COLOR_DESC" _
                & " from SOTPICK2,SOTORDR2,SOTPICK1,ICTCOLR1 " _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
                & " and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO and SOTPICK1.SHIP_BOL_NO = :PARM1" _
                & " and ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                & ", SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.INNER_PACK_QTY, SOTORDR2.CUST_STYLE_CODE" _
                & ", SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_SIZE_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU"
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select * from WHTLPXN1 where LP_XNO_SOURCE = '" & MENU_ITEM_OBJECT & "'" _
                & " and INIT_DATE >= :PARM1 and INIT_DATE -1  < :PARM2"
            Create_TDA(.Tables.Add, "WHTLPXN1", "**", 0, False, "DD", 1)

            Create_TDA(dst.Tables.Add, "TATCNTRY", "*", 0, False)

        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        grdWHTLPXN1.DataSource = dst.Tables("WHTLPXN1")

        Fill_Records("ICTWHSEX")
        Fill_Records("TATCNTRY")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEX, "SHIPS")

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTSHIPX, "SEL")
        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")

        Create_Summary(grdSOTPICKX, "STYLE_CODE", "Count")
        Create_Summary(grdSOTPICKX, New String() _
                       {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK" _
                       , "PICK_QTY_CANC_REL", "PICK_QTY_BACK_REL"})

        Create_Summary(grdWHTLPXN1, "LP_XNO", "Count")


        grdSOTSHIPX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For Each C As UltraWinGrid.UltraGridColumn In grdSOTSHIPX.DisplayLayout.Bands(0).Columns
            If C.Key = "SEL" Then
                C.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                C.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        With grdSOTSHIPX.DisplayLayout.Bands("SOTSHIPX")
            .Columns("SHIP_BOL_NO").Header.Fixed = True
            .Columns("SEL").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
        End With
        With grdSOTPICKX.DisplayLayout.Bands("SOTPICKX")
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        calFrom.Value = Now.Date.AddDays(-10)
        calTo.Value = Now.Date

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each TABLE_NAME As String In New String() _
               {"SOTSHIP1_3PL", "SOTPICK1_3PL", "SOTPICK2_3PL", "SOTCART1_3PL", "SOTCART2_3PL"}
                ASW.Add(TABLE_NAME, ASCMAIN1.Temp_Table("Select * from " & TABLE_NAME & " where ROWNUM <1"))
            Next
        End If

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
        Else
            optPending.ValueList.ValueListItems(1).DisplayText = "Transmitted"
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    Else
                        If rowICTWHSE1.Item("LP_CODE") & "" = "" Then
                            EMsg &= vbCr & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " is not set up as a 3PL"
                        Else
                            rowWHTTPLP1 = LookUp("WHTTPLP1", rowICTWHSE1.Item("LP_CODE"))
                            If rowWHTTPLP1 Is Nothing Then
                                EMsg &= vbCrLf & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " Does NOT have a valid value specified for its 3PL"
                            End If
                        End If

                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                    LP_CODE = rowICTWHSE1.Item("LP_CODE")

                    If Not ASCMAIN1.Logical_Lock("WHTSPCK1", WHSE_CODE) Then Exit Sub

                End If

            Case "Send"
                Shipments = dst.Tables("SOTSHIPX").Select("SEL = '1'").Length
                If Shipments = 0 Then
                    EMsg &= vbCr & "No Shipments Selected"
                    Exit Select
                End If

                Shipments = dst.Tables("SOTSHIPX").Select("SEL = '1' AND SHIP_XMIT_FLAG = 'H'").Length
                If Shipments > 0 Then
                    Dim zMsg As String = "There are " & Shipments & " selected shipment(s) on Hold. Do you want to Send these anyway?"
                    If MessageBox.Show(zMsg, "Send", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

                If EMsg = "" Then
                    If optPending.Value = "0" Then
                        If MsgBox("You are about to send Pick Tickets Electronically over to the 3PL" _
                              & vbCrLf _
                              & vbCrLf & "No Changes or De-Releases are Permitted" _
                              & vbCrLf & " to these Orders once they are sent to the 3PL" _
                              & vbCrLf & " without getting the 3PL to Void the corresponding Record in their System" _
                              & vbCrLf _
                              & vbCrLf & "OK To Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    Else
                        If MsgBox("You are about to mark the Selected Shipments as De-Transmitted." _
                              & vbCrLf _
                              & vbCrLf & "Once De-Transmitted, these shipments may be either" _
                              & vbCrLf & " 1) Re-Cartonized, or" _
                              & vbCrLf & " 2) De-Released, and then cancelled or changed and then Re-Released" _
                              & vbCrLf _
                              & vbCrLf & "If you Re-Cartonize or De-Release/Re-Release, " _
                              & vbCrLf & " you will need to Re-Transmit these Pick Tickets to the 3PL." _
                              & vbCrLf _
                              & vbCrLf & "OK To Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Request 940 Cancel"
                Shipments = dst.Tables("SOTSHIPX").Select("SEL='1'").Length
                If Shipments = 0 Then
                    EMsg &= vbCr & "No Shipments Selected"
                End If

                If EMsg = "" Then
                    If MsgBox("You are about to send a Request to Cancel Pick Tickets to the 3PL" _
                             & vbCrLf _
                             & vbCrLf & "You should have communicated with your CSR before doing this" _
                             & vbCrLf & " to make sure that these Pick Tickets are able to be Cancelled." _
                             & vbCrLf _
                             & vbCrLf & "Once you get a confirmation email, you should then De-Transmit these Shipments." _
                             & vbCrLf _
                             & vbCrLf & "OK To Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
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
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Send"
                Update_Record()
                Mode_Settings(False)

            Case "Request 940 Cancel"
                Request_940_Cancel()
                Load_SOTSHIPX()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Send").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Transmission Controls").Visible = ScreenMode

                If ASCMAIN1.DBS_SERVER = "NYA" OrElse ASCMAIN1.DBS_COMPANY = "NYA" Then
                    .Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
                Else
                    .Groups("Screen Control").Items("Request 940 Cancel").Visible = False
                End If

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        splShipments.Visible = ScreenMode


        If ScreenMode Then
            grdSOTSHIPX.Dock = DockStyle.None
            grdSOTSHIPX.Parent = splShipments.Panel1
            grdSOTSHIPX.Dock = DockStyle.Fill
            grdSOTSHIPX.Text = "Shipments Pending Transmission to 3PL"
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("SEL").Hidden = False
            grdSOTSHIPX.Visible = True
        Else
            Clear_Record()
            grdSOTSHIPX.Dock = DockStyle.None
            grdSOTSHIPX.Parent = splTransmissions.Panel2
            grdSOTSHIPX.Dock = DockStyle.Fill
            grdSOTSHIPX.Text = "Shipments Transmitted"
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIPX", "SOTPICK1", "SOTPICKX", "WHTLPXN1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optPending.Value = "0"
        Load_WHTLPXN1()
        Fill_Records("ICTWHSEX")
        Setup_tab0()
        UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Load_SOTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ASCMAIN1.sql = "Truncate Table " & WHTSTYLX
            ASCDATA1.ExecuteSQL()

            For Each TABLE_NAME As String In New String() _
               {"SOTSHIP1_3PL", "SOTPICK1_3PL", "SOTPICK2_3PL", "SOTCART1_3PL", "SOTCART2_3PL"}
                ASCDATA1.ExecuteSQL("Truncate Table " & ASW(TABLE_NAME))
            Next
        End If

        '  ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
        ASCDATA1.DeleteRows(dst.Tables("SOTSHIPX"), "ISNULL(SEL,'0') <> '1'")
        Update_Record_TDA("SOTSHIPX", "1=1") ' this should work, but stds don't know that this is a temp table

        BeginTrans()

        LP_XNO = TAC.WHCMAIN1.Get_LP_XNO(MENU_ITEM_OBJECT, Shipments)

        ' THIS TABLE IS PROBABLY UNNEC SINCE WE ARE MARKING SOTSHIP1
        ASCMAIN1.sql = "Insert into WHTSHIP1 (LP_XNO, SHIP_BOL_NO) " _
            & " Select '" & LP_XNO & "' LP_XNO, SHIP_BOL_NO from " & SOTSHIPX
        ASCDATA1.ExecuteSQL()

        If optPending.Value = "0" Then
            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '1', LP_XNO = '" & LP_XNO & "', LP_XMIT_DATE = SYSDATE, SHIP_XMIT_FLAG = NULL"
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Else
                ASCMAIN1.sql &= ", SHIP_PICK_PRINTED = SYSDATE"
            End If
            ASCMAIN1.sql &= " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Send_3PL()

        Else
            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL"
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Else
                ASCMAIN1.sql &= ", SHIP_PICK_PRINTED = NULL"
            End If
            ASCMAIN1.sql &= " where SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            De_Transmit_3PL()
        End If

        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWHSEX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTSHIPX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins" _
                        , "Select All", "De-Select All", "Select All in Group", "Recall Shipment")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTSHIPX"
                    tlb_pop.Tools("Select All").SharedProps.Visible = ScreenMode
                    tlb_pop.Tools("De-Select All").SharedProps.Visible = ScreenMode
                    tlb_pop.Tools("Recall Shipment").SharedProps.Visible = Not ScreenMode

                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    Else
                        tlb_pop.Tools("Recall Shipment").SharedProps.Visible = False
                    End If


                    tlb_pop.Tools("Select All in Group").SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow)

                    If Not ScreenMode Then
                        If grd.ActiveRow.Cells("LP_XNO").Value & "" <> grd.ActiveRow.Cells("LP_XNO_XMIT").Value & "" _
                        Or grd.ActiveRow.Cells("LP_STATUS").Value & "" <> "1" Then
                            tlb_pop.Tools("Recall Shipment").SharedProps.Visible = False
                        End If
                    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All"

                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTSHIPX").Rows
                    rowSOTWSHIPX.Item("SEL") = "1"
                Next

                MsgBox("You have selected " & dst.Tables("SOTSHIPX").Select("SEL = '1'").Length & " Records by Selecting All", MsgBoxStyle.OkOnly, "Verification")

            Case "De-Select All"

                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTSHIPX").Rows
                    rowSOTWSHIPX.Item("SEL") = "0"
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Recall Shipment"
                Dim SHIP_BOL_NO As String = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim LP_XNO As String = grdSOTSHIPX.ActiveRow.Cells("LP_XNO").Value
                If MsgBox("Are you sure you want to Recall all Pick Tickets and Shipments for Shipment " & SHIP_BOL_NO, _
                          MsgBoxStyle.YesNo, _
                          "Verification to Recall Pick Tickets and Shipments from a 3PL") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                Recall_Shipment(SHIP_BOL_NO, LP_XNO)


            Case "Select All in Group"
                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("ORDR_GROUP_NO = '" & grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & "'")
                    rowSOTWSHIPX.Item("SEL") = "1"
                Next
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("Load")
        End Select
    End Sub
#End Region

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("Load")
    End Sub

    Public Function Build_List_of_Objects(Of C As {New})(sql As String) As List(Of C)

        Dim objList As New List(Of C)
        Dim ALL_COLUMNS As Dictionary(Of String, System.Reflection.FieldInfo) _
            = Get_Columns_from_Class(GetType(C))

        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)
        Dim row_count_total As Int32 = tbl.Rows.Count
        Dim row_counter As Int32 = 0

        For Each row As DataRow In tbl.Rows
            row_counter += 1

            Dim objItem As New C

            If 1 <> 1 Then
                ALL_COLUMNS = Get_Columns_from_Class(GetType(C))
            End If

            For Each COLUMN_NAME In ALL_COLUMNS.Keys
                If row.Item(COLUMN_NAME) & "" = "" Then
                Else
                   

                    If row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.DateTime" Then
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, row.Item(COLUMN_NAME))
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.String" Then
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, row.Item(COLUMN_NAME))
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Double" Then
                        Dim V As Decimal = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int32" Then
                        Dim V As Int32 = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int16" Then
                        Dim V As Int16 = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    Else
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, Val(row.Item(COLUMN_NAME)))
                    End If

                End If
            Next
            objList.Add(objItem)
        Next

        Return objList
    End Function

    Public Shared Function Get_Columns_from_Class(T As Type) _
        As Dictionary(Of String, System.Reflection.FieldInfo)

        Dim COLUMN_NAMEs As New Dictionary(Of String, System.Reflection.FieldInfo)
        ' Dim COLUMN_NAMEs As New Dictionary(Of String, System.Reflection.PropertyInfo)

        'Dim t As Type = XX.GetType
        Dim fieldName As String
        ' Dim propertyValue As Object

        ' Use each property of the business object passed in 
        'For Each pi As System.Reflection.PropertyInfo In _
        '        T.GetProperties(System.Reflection.BindingFlags.Instance Or _
        '                        System.Reflection.BindingFlags.Public Or _
        '                        System.Reflection.BindingFlags.NonPublic)
        '    ' Get the name and value of the property 
        '    If pi.Name <> "ExtensionData" Then
        '        fieldName = pi.Name
        '        COLUMN_NAMEs.Add(fieldName, pi)
        '    End If

        '    ' Get the value of the property 
        '    ' propertyValue = pi.GetValue(XX, Nothing)
        '    'Console.WriteLine(fieldName & ": " &
        '    'If(propertyValue Is Nothing, "Nothing", propertyValue.ToString))
        'Next

        For Each pi As System.Reflection.FieldInfo In _
               T.GetFields(System.Reflection.BindingFlags.Instance Or _
                               System.Reflection.BindingFlags.Public Or _
                               System.Reflection.BindingFlags.NonPublic)
            If pi.MemberType = Reflection.MemberTypes.Field Then
                fieldName = pi.Name
                If fieldName <> "SQL" Then
                    ' Debug.Write(pi.Name & ":" & pi.MemberType.ToString)
                    COLUMN_NAMEs.Add(fieldName, pi)
                End If
            End If
        Next
        Return COLUMN_NAMEs
    End Function

    Private Sub grdWHTLPXN1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTLPXN1.AfterRowActivate
        Setup_WHTLPXN1()
    End Sub

    Sub Setup_WHTLPXN1()
        If grdWHTLPXN1.ActiveRow Is Nothing Then
            grdSOTSHIPX.Visible = False
        Else
            grdSOTSHIPX.Visible = True
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
            Dim LP_XNO As String = grdWHTLPXN1.ActiveRow.Cells("LP_XNO").Value
            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " (SHIP_BOL_NO) Select SHIP_BOL_NO from WHTSHIP1 where LP_XNO = '" & LP_XNO & "'")
            Fill_Records("SOTSHIPX")
            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO")
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        UltraExplorerBar1.Groups("Transmission History").Visible = Not ScreenMode And tab0.SelectedTab.Key = "Transmissions"
    End Sub

    Private Sub btnLoadHistory_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadHistory.Click
        Load_WHTLPXN1()
    End Sub

    Sub Load_WHTLPXN1()
        Fill_Records("WHTLPXN1", New Object() {calFrom.Value, calTo.Value})
        Sort_grdColumns(grdWHTLPXN1, "LP_XNO".ToLower)
        Setup_WHTLPXN1()
    End Sub

    Private Sub grdSOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSHIPX.AfterRowActivate
        Setup_SOTSHIPX()
    End Sub

    Sub Setup_SOTSHIPX()
        If grdSOTSHIPX.ActiveRow Is Nothing OrElse Not grdSOTSHIPX.ActiveRow.IsDataRow Then
            tabShipment.Visible = False
        Else
            tabShipment.Visible = True
            Dim SHIP_BOL_NO As String = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
            grdSOTPICK1.Text = "Pick Tickets for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICK1", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICKX.Text = "Style/Color Summary for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICKX", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTPICKX, "STYLE_CODE,COLOR_CODE")
        End If
    End Sub

    Sub Recall_Shipment(SHIP_BOL_NO As String, LP_XNO As String)

        If Not ASCMAIN1.Logical_Lock("WHTSPCK1", WHSE_CODE) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub

        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
        If rowSOTSHIP1.Item("LP_XNO") & "" <> LP_XNO _
        Or rowSOTSHIP1.Item("LP_STATUS") <> "1" Then
            ' DO NOTHING, SOMETHING HAS CHANGED
        Else
            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL where SHIP_BOL_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {SHIP_BOL_NO})
        End If

        Mode_Settings(False)

    End Sub

    Private Sub grdSOTSHIPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIPX.InitializeRow
        If Not ScreenMode And EntryMode = "" Then
            If grdWHTLPXN1.ActiveRow IsNot Nothing Then
                If e.Row.Cells("LP_XNO_XMIT").Value & "" <> grdWHTLPXN1.ActiveRow.Cells("LP_XNO").Value & "" Then
                    e.Row.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    e.Row.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            End If
        Else
            e.Row.CellAppearance.BackColor = Drawing.Color.Empty
        End If

        If e.Row.Cells("SHIP_XMIT_FLAG").Value & String.Empty = "H" Then
            e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Shipment is On Hold pending Revised Ship Date"
        End If

    End Sub

    Sub De_Transmit_3PL()

        Dim sqlSHIP_BOL_NO As String = ""
        For Each row As DataRow In dst.Tables("SOTSHIPX").Select("")
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            sqlSHIP_BOL_NO &= ",'" & SHIP_BOL_NO & "'"
        Next
        sqlSHIP_BOL_NO = " where SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NO, 2) & ")"
        Dim sqlPICK_NO As String = " where PICK_NO in (Select PICK_NO from ADS.SOTPICK1_3PL@ADSIIS" & sqlSHIP_BOL_NO & ")"

    End Sub

    Sub Send_3PL()


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Sending to 3PL")

        ' Dim ORDR_VAS As New Dictionary(Of String, String)

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then

            MsgBox("RED PRARIE IS SHUT OFF")

        ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then

            If Not dst.Tables.Contains("EDT940O1") Then
                Create_TDA(dst.Tables.Add, "EDT940O1", "*")
                Create_TDA(dst.Tables.Add, "EDT940O2", "*")
                Create_TDA(dst.Tables.Add, "EDT940O4", "*")
                Create_TDA(dst.Tables.Add, "EDT940O5", "*")
                Create_TDA(dst.Tables.Add, "EDT940O6", "*")
            End If

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
            Dim WHSE_EMAIL_VAS As String = rowICTWHSE1.Item("WHSE_EMAIL_VAS") & ""

            '  Dim EMAIL_ADDRESS_list As List(Of String) = Split(WHSE_EMAIL_VAS, ";").ToList
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            If WHSE_EMAIL_VAS <> "" Then
                For Each EMAIL_ADDRESS As String In Split(WHSE_EMAIL_VAS, ";")
                    EMAIL_ADDRESSs.Add(EMAIL_ADDRESS, "VAS Coordinator")
                Next
            End If

            Dim rowEDTTRPM1 As DataRow = LookUp("EDTTRPM1", _
                                                New String() {rowICTWHSE1.Item("WHSE_EDI_QUAL"), rowICTWHSE1.Item("WHSE_EDI_ID"), "943"})

            ' BeginTrans()

            For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
                Dim CUST_CODE As String = rowSOTSHIPX.Item("CUST_CODE")
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                Dim rowEDTSLSP1 As DataRow = LookUp("EDTSLSP1", CUST_CODE)
                Dim EDI_SLN_TOT_IND As String = ""
                If rowEDTSLSP1 IsNot Nothing Then
                    EDI_SLN_TOT_IND = rowEDTSLSP1.Item("EDI_SLN_TOT_IND") & ""
                End If

                Dim MAX_CHARS_DC As Integer = 0
                If rowEDTSLSP1 IsNot Nothing Then
                    Dim NUMBER_CHARS_STORE As Integer = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & "")
                    Dim NUMBER_CHARS_DC As Integer = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & "")
                    If NUMBER_CHARS_DC <> 0 Then
                        MAX_CHARS_DC = NUMBER_CHARS_DC
                    ElseIf NUMBER_CHARS_STORE <> 0 Then
                        MAX_CHARS_DC = NUMBER_CHARS_STORE
                    End If
                End If

                Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("SHIP_BOL_NO")

                Dim SHIP_VIA_CODE As String = rowSOTSHIPX.Item("SHIP_VIA_CODE") & ""
                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE, True)
                Dim rowSOTCARR1 As DataRow = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty, True)

                Dim FDX_ACCT_NO As String = rowARTCUST1.Item("FDX_ACCT_NO") & ""
                Dim UPS_ACCT_NO As String = rowARTCUST1.Item("UPS_ACCT_NO") & ""
                Dim SHIP_ACCT_NO As String = rowSOTCARR1.Item("SHIP_ACCT_NO") & String.Empty

                ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P'"
                For Each rowSOTPICK1 As DataRow In ASCDATA1.GetDataTable.Select("", "PICK_NO")

                    Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                    Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)

                    ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                        & " Values ('SOTORDR1', '" & ORDR_NO & "', SYSDATE, '" & ASCMAIN1.USER_ID & "'" _
                        & ", 'RELORD','Order Released to Whse', '" & EDI_OUTBOUND_DOC_NO & "')"
                    ASCDATA1.ExecuteSQL()

                    If ASCMAIN1.CLIENT = "NYA" Then
                        If rowSOTORDR1.Item("ORDR_INCL_VAS") & "" = "1" And WHSE_EMAIL_VAS <> "" Then
                            Dim CUST_NAME As String = rowSOTORDR1.Item("CUST_NAME") & ""
                            Dim ORDR_CUST_PO As String = rowSOTORDR1.Item("ORDR_CUST_PO") & ""
                            Dim EMAIL_BODY As String = "VAS Coordinator - please read - IMPORTANT" & vbCrLf & vbCrLf _
                                                       & "VAS Instructions:" & vbCrLf & vbCrLf _
                                                       & rowSOTORDR1.Item("ORDR_SHIP_INSTR") & ""
                            Dim SEND_NO As String = ""

                            Dim EMAIL_SUBJECT = "VAS for " & CUST_CODE & ", PO " & ORDR_CUST_PO & "; Sales Order " & ORDR_NO & "; Pick Ticket " & PICK_NO
                            EMAIL_SUBJECT = "VAS for " & CUST_CODE & " PO " & ORDR_CUST_PO & " Sales Order " & ORDR_NO & " Pick Ticket " & PICK_NO
                            EMAIL_SUBJECT = Replace(EMAIL_SUBJECT, ".", " ")

                            SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing, _
                                     EMAIL_SUBJECT, "VAS", True,
                                     False, CUST_CODE, CUST_NAME, "Customer", EMAIL_BODY)

                            ASCMAIN1.Record_Event("SOTORDR1", ORDR_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, _
                                                  "VASEML", "VAS email", SEND_NO)
                        End If
                    End If

                    Dim EDI_DOC_SEQ_NO As String = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
                    Dim rowEDT850T1 As DataRow = LookUp("EDT850T1", EDI_DOC_SEQ_NO)

                    Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & ""
                    Dim CUST_DC_NO As String = rowSOTORDR1.Item("CUST_DC_NO") & ""

                    Dim rowARTCUST2_MK As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
                    Dim rowARTCUST2_DC As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "DC", CUST_DC_NO})

                    Dim rowEDT940O1 As DataRow = dst.Tables("EDT940O1").NewRow
                    With rowEDT940O1
                        .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_STORE_NO") = CUST_STORE_NO
                        .Item("PICK_NO") = PICK_NO
                        .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                        .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                        .Item("EDI_SUPPLIER_NO") = rowARTCUST1.Item("CUST_VEND_REF")
                        .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                        .Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                        .Item("ORDR_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
                        .Item("ORDR_PO_DATE") = rowSOTORDR1.Item("ORDR_DATE")
                        .Item("UPS_REF_1") = CreateUpsRef1(CUST_CODE, EDI_DOC_SEQ_NO)

                        Dim FRT_TERMS As String = rowSOTSHIPX.Item("FRT_TERMS") & ""
                        Dim FRT_TERMS_EDI As String = ""
                        Select Case FRT_TERMS
                            Case "PPD", "PPA"
                                FRT_TERMS_EDI = "PP"
                            Case "COL"
                                FRT_TERMS_EDI = "CC"
                            Case "3PY"
                                FRT_TERMS_EDI = "TP"
                        End Select
                        .Item("FRT_TERMS") = FRT_TERMS_EDI
                        '.Item("EDI_TRANS_METH_CODE") = "?"
                        .Item("EDI_SCAC_CODE") = "ROUT" ' rowSOTSVIA1.Item("SHIP_VIA_SCAC")

                        If SHIP_VIA_CODE <> "ROUT" And SHIP_VIA_CODE <> "" Then
                            ASCMAIN1.sql = "Select * from EDTXREF3" _
                                & " where SENDER_ID_QUAL = :PARM1 and SENDER_ID = :PARM2 and SHIP_VIA_CODE = :PARM3"

                            Dim rowEDTXREF3 As DataRow = ASCDATA1.GetDataRow _
                                                         (ASCMAIN1.sql, "VVV", New String() _
                                                          {rowICTWHSE1.Item("WHSE_EDI_QUAL"), _
                                                           rowICTWHSE1.Item("WHSE_EDI_ID"), _
                                                           SHIP_VIA_CODE})

                            If rowSOTSVIA1.Item("CARRIER_CODE") & "" = "FEDEX" Then
                                .Item("EDI_SCAC_CODE") = "FEDX"
                            End If
                            If rowSOTSVIA1.Item("CARRIER_CODE") & "" = "UPS" Then
                                .Item("EDI_SCAC_CODE") = "UPSN"
                            End If

                            If rowEDTXREF3 IsNot Nothing Then
                                .Item("EDI_SERVICE_LEVEL") = rowEDTXREF3.Item("SERVICE_LEVEL_3PL")

                                If FRT_TERMS = "COL" Or FRT_TERMS = "3PY" Then
                                    Dim FDX_ACCT_NO_MK As String = rowARTCUST2_MK.Item("FDX_ACCT_NO") & String.Empty
                                    Dim UPS_ACCT_NO_MK As String = rowARTCUST2_MK.Item("UPS_ACCT_NO") & String.Empty
                                    If rowARTCUST2_MK IsNot Nothing Then
                                        If rowARTCUST2_MK.Item("FDX_ACCT_NO") & String.Empty <> String.Empty Then FDX_ACCT_NO_MK = rowARTCUST2_MK.Item("FDX_ACCT_NO") & String.Empty
                                        If rowARTCUST2_MK.Item("UPS_ACCT_NO") & String.Empty <> String.Empty Then UPS_ACCT_NO_MK = rowARTCUST2_MK.Item("UPS_ACCT_NO") & String.Empty
                                    End If

                                    Dim EDI_TP_BILLING_ACCT As String = String.Empty
                                    Select Case rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
                                        Case "FEDEX"
                                            If FDX_ACCT_NO_MK.Length > 0 Then
                                                EDI_TP_BILLING_ACCT = FDX_ACCT_NO_MK
                                            Else
                                                EDI_TP_BILLING_ACCT = FDX_ACCT_NO
                                            End If

                                        Case "UPS"
                                            If UPS_ACCT_NO_MK.Length > 0 Then
                                                EDI_TP_BILLING_ACCT = UPS_ACCT_NO_MK
                                            Else
                                                EDI_TP_BILLING_ACCT = UPS_ACCT_NO
                                            End If
                                    End Select

                                    .Item("EDI_TP_BILLING_ACCT") = EDI_TP_BILLING_ACCT
                                    ' IF FRT_TERMS WAS COL OR 3RD PARTY WE WOULD SEND THE 3RD PARTY ACCT NUMBER"
                                End If
                            End If
                        End If

                        If .Item("EDI_TP_BILLING_ACCT") & String.Empty = String.Empty _
                            AndAlso (FRT_TERMS = "PPA" OrElse FRT_TERMS = "PPD") _
                            AndAlso SHIP_ACCT_NO.Length > 0 Then
                            .Item("EDI_TP_PPD_BILLING_ACCT") = SHIP_ACCT_NO
                        End If

                        .Item("EDI_DIVISION_CODE") = rowICTWHSE1.Item("LP_WHSE_ID")

                        ' .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("CUST_VEND_REF")
                        .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("LABEL_TEMPLATE_CODE")
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("ORDR_TYPE_CODE") = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                        'If rowEDT850T1 IsNot Nothing Then
                        '    .Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")
                        'End If
                        .Item("EDI_MERCH_TYPE") = rowSOTORDR1.Item("EDI_MERCH_TYPE")
                        .Item("ORDR_STATUS_CODE") = "N"

                        If rowEDT850T1 IsNot Nothing Then
                            .Item("EDI_CHAIN") = rowEDT850T1.Item("EDI_CHAIN")
                            .Item("EDI_FACILITY") = rowEDT850T1.Item("EDI_FACILITY")
                        End If

                        If rowARTCUST2_MK IsNot Nothing AndAlso rowARTCUST2_MK.Item("CUST_ADDR_GROUP") & "" <> "" Then
                            .Item("CUST_ADDR_GROUP") = rowARTCUST2_MK.Item("CUST_ADDR_GROUP")
                        End If

                    End With
                    dst.Tables("EDT940O1").Rows.Add(rowEDT940O1)

                    Dim sqlMLK As String = "SOTORDR2.EDI_DTL_SEQ"
                    If EDI_DOC_SEQ_NO = "" Then
                        sqlMLK = "SOTORDR2.STYLE_CODE"
                    End If

                    ASCMAIN1.sql = "Select SOTPICK2.*," & sqlMLK & " MLK" & vbCrLf _
                        & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_UOM, SOTORDR2.STYLE_DESC" & vbCrLf _
                        & ", SOTORDR2.EDI_DTL_SEQ, SOTORDR2.EDI_SLN_SEQ" & vbCrLf _
                        & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE" & vbCrLf _
                        & ", SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU, SOTORDR2.CUST_SIZE_CODE" & vbCrLf _
                        & " from SOTPICK2,SOTORDR2" & vbCrLf _
                        & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                        & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                        & "   and SOTPICK2.PICK_NO = '" & PICK_NO & "'"

                    Dim tbl As DataTable = ASCDATA1.GetDataTable

                    ' Make a list of combinable details, counting how many lines would be combined
                    '  for EDI Orders we combine based on EDI_DTL_SEQ_NO, 
                    '  and for Non-EDI we combine based on STYLE_CODE

                    sqlMLK = Replace(sqlMLK, "SOTORDR2.", "")
                    ASCMAIN1.sql = "Select " & sqlMLK & " MLK, Count (*) LINES, Min (PICK_LNO) PICK_LNO from (" _
                    & ASCMAIN1.sql & ") group by " & sqlMLK & " having Count (*) > 1"
                    Dim tblML As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ML", 1)
                    tblML.Columns.Add("SLN_LNO", GetType(System.Int32))
                    tblML.Columns("SLN_LNO").DefaultValue = 0

                    ' no multi-lines if the order is an EDI order and the flag is not set

                    If EDI_DOC_SEQ_NO <> "" Then ' if this is an EDI Order
                        If EDI_SLN_TOT_IND = "1" Then ' if this flag is set
                            ' ok for multi-line
                        Else
                            tblML.Rows.Clear()
                        End If
                    End If

                    ' NO MULTI-LINE NOW UNTIL WE ARE DONE TESTING
                    '    tblML.Rows.Clear()

                    Dim MULTI_LINE_NO As Integer = 0

                    Dim lead_item_processed As Boolean = False
                    For Each rowSOTPICK2 As DataRow In tbl.Select("", "PICK_LNO")

                        Dim PICK_LNO As Int32 = Val(rowSOTPICK2.Item("PICK_LNO") & "")
                        Dim PICK_QTY As Int64 = Val(rowSOTPICK2.Item("PICK_QTY") & "")

                        Dim EDI_DTL_SEQ As Int32 = Val(rowSOTPICK2.Item("EDI_DTL_SEQ") & "")
                        Dim rowEDT850T2 As DataRow = LookUp("EDT850T2", New String() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ})
                        Dim EDI_SLN_SEQ As Int32 = Val(rowSOTPICK2.Item("EDI_SLN_SEQ") & "")
                        Dim rowEDT850T6 As DataRow = LookUp("EDT850T6", New String() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ, EDI_SLN_SEQ})

                        Dim STYLE_CODE As String = rowSOTPICK2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTPICK2.Item("COLOR_CODE")

                        Dim skip_row As Boolean = False

                        Dim MLK As String = rowSOTPICK2.Item("MLK")
                        Dim rowML As DataRow = tblML.Rows.Find(MLK)

                        ' Dim MultiLine As Boolean = (rowML IsNot Nothing)

                        If rowML IsNot Nothing Then
                            MULTI_LINE_NO = rowML.Item("PICK_LNO")
                            'If EDI_DOC_SEQ_NO = "" Then
                            '    If Not STYLE_CODEs.ContainsKey(STYLE_CODE) Then
                            '        STYLE_CODEs.Add(STYLE_CODE, STYLE_CODEs.Count + 1)
                            '    End If
                            '    MULTI_LINE_NO = STYLE_CODEs(STYLE_CODE)
                            'Else
                            '    MULTI_LINE_NO = EDI_DOC_SEQ_NO
                            'End If
                            Dim SLN_LNO As Integer = Val(rowML.Item("SLN_LNO") & "") + 1
                            rowML.Item("SLN_LNO") = SLN_LNO

                            If MULTI_LINE_NO = 0 Then
                                Throw New Exception("Cannot Have an Order Detail without a link to an EDI Detail - See Line " & CStr(PICK_LNO))
                            End If

                            Dim rowEDT940O2_cum As DataRow = dst.Tables("EDT940O2").Rows.Find _
                                 (New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, MULTI_LINE_NO})

                            If rowEDT940O2_cum IsNot Nothing Then
                                rowEDT940O2_cum.Item("PICK_QTY") = Val(rowEDT940O2_cum.Item("PICK_QTY") & "") + PICK_QTY
                                skip_row = True
                            End If

                            Dim rowEDT940O6 As DataRow = dst.Tables("EDT940O6").NewRow
                            With rowEDT940O6
                                .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                                .Item("PICK_LNO") = MULTI_LINE_NO
                                .Item("SLN_LNO") = SLN_LNO
                                .Item("PICK_QTY") = PICK_QTY
                                .Item("ORDR_LNO") = rowSOTPICK2.Item("ORDR_LNO")

                                If EDI_DOC_SEQ_NO = "" Then

                                    'EDI_PARENT_UPC,
                                    'EDI_PARENT_SKU,

                                    'EDI_SLN_PRICE,
                                    'EDI_SLN_ITEM,
                                    'EDI_SLN_PO4_UOM,
                                    'EDI_PO4_QTY,
                                    'EDI_PO4_INNER,
                                    'EDI_SLN_LBL_CODE,
                                    'EDI_SLN_RETAIL_PRICE,
                                    'EDI_SLN_PO_LNO,
                                    'EDI_SLN_DEPT,
                                    'EDI_SLN_LINE_MODE,
                                    'EDI_SLN_ID,
                                    'EDI_SLN_BUYER_ITEM,
                                    'EDI_SHIP_DC,

                                    .Item("EDI_SLN_UOM") = "EA"
                                    .Item("EDI_SLN_ITEM_DESC") = rowSOTPICK2.Item("STYLE_DESC")
                                    .Item("EDI_SLN_UPC") = rowSOTPICK2.Item("CUST_UPC")
                                    .Item("EDI_SLN_QTY") = rowSOTPICK2.Item("PICK_QTY")
                                    .Item("EDI_SLN_UPC") = rowSOTPICK2.Item("CUST_UPC")
                                    .Item("EDI_SLN_STYLE") = rowSOTPICK2.Item("CUST_STYLE_CODE")
                                    .Item("EDI_SLN_COLOR") = rowSOTPICK2.Item("CUST_COLOR_CODE")
                                    '.Item("EDI_SLN_COLOR_CODE") = rowSOTPICK2.Item("CUST_COLOR_CODE")
                                    .Item("EDI_SLN_SKU") = rowSOTPICK2.Item("CUST_SKU")
                                    .Item("EDI_SLN_SIZE_CODE") = rowSOTPICK2.Item("CUST_SIZE_CODE")
                                    .Item("EDI_SLN_SIZE_DESC") = rowSOTPICK2.Item("CUST_SIZE_CODE")
                                Else
                                    For Each DC As DataColumn In dst.Tables("EDT940O6").Columns
                                        If New String() {"COMPANY_CODE", "EDI_OUTBOUND_DOC_NO", "PICK_LNO", "SLN_LNO", "PICK_QTY", "ORDR_LNO"}.Contains(DC.ColumnName) Then
                                        Else
                                            .Item(DC.ColumnName) = rowEDT850T6.Item(DC.ColumnName)
                                        End If
                                    Next
                                End If

                            End With
                            dst.Tables("EDT940O6").Rows.Add(rowEDT940O6)
                        End If

                        If skip_row Then
                        Else

                            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", New String() {STYLE_CODE})

                            If Not lead_item_processed Then
                                Dim EDI_DIVISION_CODE As String = rowEDT940O1.Item("EDI_DIVISION_CODE")
                                If rowICTSTYL1.Item("CUST_CODE") & "" = "DOLGEN" Then
                                    EDI_DIVISION_CODE = "NYDG"
                                ElseIf rowICTSTYL1.Item("CUST_CODE") & "" = "WALMART" Then
                                    If rowICTSTYL1.Item("STYLE_GROUP_CODE") & "" = "07" Then
                                        EDI_DIVISION_CODE = "NYWB"
                                    Else
                                        EDI_DIVISION_CODE = "NYWM"
                                    End If
                                End If
                                rowEDT940O1.Item("EDI_DIVISION_CODE") = EDI_DIVISION_CODE
                                lead_item_processed = True
                            End If

                            Dim rowEDT940O2 As DataRow = dst.Tables("EDT940O2").NewRow
                            With rowEDT940O2
                                .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                                .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                                .Item("PICK_LNO") = PICK_LNO
                                .Item("PICK_QTY") = PICK_QTY
                                .Item("STYLE_UOM") = rowSOTPICK2.Item("STYLE_UOM")

                                Dim STYLE_CODE_EDI As String = STYLE_CODE & COLOR_CODE
                                If rowICTSTYC1.Item("HIDE_COLOR_3PL") & "" = "1" Then
                                    STYLE_CODE_EDI = STYLE_CODE
                                End If

                                .Item("STYLE_CODE") = STYLE_CODE_EDI
                                .Item("STYLE_DESC") = rowSOTPICK2.Item("STYLE_DESC")
                                .Item("UNIT_PRICE") = rowSOTPICK2.Item("PICK_UNIT_PRICE")

                                .Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE") & ""

                                'If MultiLine Then
                                'Else
                                If rowSOTPICK2.Item("CUST_UPC") & "" <> "" Then
                                    .Item("UPC_CODE") = rowSOTPICK2.Item("CUST_UPC")
                                End If
                                If rowSOTPICK2.Item("CUST_STYLE_CODE") & "" <> "" Then
                                    .Item("EDI_STYLE") = rowSOTPICK2.Item("CUST_STYLE_CODE")
                                Else
                                    If rowEDT850T2 IsNot Nothing Then
                                        .Item("EDI_STYLE") = rowEDT850T2.Item("EDI_GTIN") & ""
                                    End If
                                End If

                                .Item("EDI_COLOR") = rowSOTPICK2.Item("CUST_COLOR_CODE")
                                .Item("EDI_SKU") = rowSOTPICK2.Item("CUST_SKU")
                                .Item("EDI_SIZE") = rowSOTPICK2.Item("CUST_SIZE_CODE")
                                'End If

                                ' MAYBE THESE CAN BE ELIMINATED FROM EDT940O2?
                                '.Item("EDI_PROD_CLASS") = "?"
                                '.Item("NMFC") = "?"
                                '.Item("EDI_STYLE_DESC") = "?"
                                '.Item("EDI_RETAIL_PRICE") = 0
                                '.Item("EDI_SIZE") = "?"

                                If rowEDT850T6 IsNot Nothing Then
                                    .Item("EDI_PO_LNO") = rowEDT850T6.Item("EDI_SLN_PO_LNO")
                                    .Item("EDI_DEPT") = rowEDT850T6.Item("EDI_SLN_DEPT")
                                    .Item("EDI_LINE_MODE") = rowEDT850T6.Item("EDI_SLN_LINE_MODE")
                                    .Item("PACK_SIZE") = rowEDT850T6.Item("EDI_SLN_QTY") '  rowEDT850T6.Item("EDI_PO4_QTY")
                                    .Item("PACK_UOM") = rowEDT850T6.Item("EDI_SLN_PO4_UOM")
                                    .Item("EDI_PARENT_SKU") = rowEDT850T6.Item("EDI_PARENT_SKU")
                                End If

                                If rowICTSTYL1.Item("COUNTRY_CODE") & "" <> "" Then
                                    Dim rowTATCNTRY As DataRow = dst.Tables("TATCNTRY").Rows.Find(rowICTSTYL1.Item("COUNTRY_CODE"))
                                    If rowTATCNTRY IsNot Nothing Then
                                        .Item("COUNTRY_NAME") = rowTATCNTRY.Item("COUNTRY_NAME")
                                    Else
                                        .Item("COUNTRY_NAME") = rowICTSTYL1.Item("COUNTRY_CODE")
                                    End If
                                End If
                            End With
                            dst.Tables("EDT940O2").Rows.Add(rowEDT940O2)
                        End If
                    Next

                    If tblML.Rows.Count <> 0 Then
                        For Each rowML As DataRow In tblML.Select("")
                            Dim PICK_LNO As Int32 = Val(rowML.Item("PICK_LNO") & "")
                            Dim rowEDT940O2 As DataRow = dst.Tables("EDT940O2").Rows.Find _
                                 (New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, PICK_LNO})
                            Dim STYLE_CODE As String = rowEDT940O2.Item("STYLE_CODE")
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            rowEDT940O2.Item("PACK_SIZE") = rowICTSTYL1.Item("INNER_PACK_QTY")
                            Dim PACK_SIZE As Int64 = Val(rowEDT940O2.Item("PACK_SIZE"))
                            Dim PICK_QTY As Int64 = Val(rowEDT940O2.Item("PICK_QTY"))
                            If PICK_QTY = 0 Or PACK_SIZE = 0 Then
                                Throw New Exception("Pick or Pack Qty is 0 for Order " & ORDR_NO)
                            End If
                            If (PICK_QTY Mod PACK_SIZE) <> 0 Then
                                Throw New Exception("Pick Qty not Divisible by Pack Qty for Order " & ORDR_NO)
                            End If
                            Dim PACK_QTY As Int64 = PICK_QTY / PACK_SIZE
                            rowEDT940O2.Item("PACK_QTY") = PACK_QTY
                            rowEDT940O2.Item("EDI_STYLE") = DBNull.Value
                            rowEDT940O2.Item("EDI_COLOR") = DBNull.Value
                            rowEDT940O2.Item("EDI_SKU") = DBNull.Value
                            rowEDT940O2.Item("EDI_SIZE") = DBNull.Value

                            Dim sqlw As String = "" _
                                & "COMPANY_CODE = '" & rowEDT940O2.Item("COMPANY_CODE") & "'" _
                                & " and EDI_OUTBOUND_DOC_NO = '" & rowEDT940O2.Item("EDI_OUTBOUND_DOC_NO") & "'" _
                                & " and PICK_LNO = " & rowEDT940O2.Item("PICK_LNO")

                            If EDI_DOC_SEQ_NO = "" Then
                                For Each rowEDT940O6 As DataRow In dst.Tables("EDT940O6").Select(sqlw)
                                    With rowEDT940O6
                                        .Item("EDI_SLN_QTY") = Val(.Item("EDI_SLN_QTY") & "") / PACK_QTY
                                    End With
                                Next
                            End If
                        Next
                    End If

                    'Dim CUST_ROUTING_INST As String = rowARTCUST1.Item("CUST_ROUTING_INST") & ""
                    'If CUST_ROUTING_INST <> "" Then '(TRA) Transportation/Routing : Special shipping instructions
                    '    Write_Notes(EDI_OUTBOUND_DOC_NO, "TRA", CUST_ROUTING_INST)
                    'End If

                    Dim ORDR_MESSAGE As String = "" ' THERE IS NO rowSOTORDR1.Item("ORDR_MESSAGE") & ""
                    ASCMAIN1.sql = "Select SOTORDR4.* from SOTORDR4 where ORDR_NO = '" & ORDR_NO & "'"
                    For Each rowSOTORDR4 As DataRow In ASCDATA1.GetDataTable.Select("", "ORDR_CLNO")
                        If ORDR_MESSAGE <> "" Then ORDR_MESSAGE &= vbCrLf
                        ORDR_MESSAGE &= rowSOTORDR4.Item("ORDR_COMMENT")
                    Next

                    Dim ORDR_SHIP_INSTR As String = rowSOTORDR1.Item("ORDR_SHIP_INSTR") & ""
                    If ORDR_SHIP_INSTR <> "" Then '(WHI) Warehouse Instructions : Instructions viewed by pickers in warehouse.
                        Write_Notes(EDI_OUTBOUND_DOC_NO, "WHI", ORDR_SHIP_INSTR)
                    End If

                    If ORDR_MESSAGE <> "" Then '(GEN) - General Order Instructions : Standard order instructions.
                        Write_Notes(EDI_OUTBOUND_DOC_NO, "GEN", ORDR_MESSAGE)
                    End If

                    Dim BT_GEN As Boolean = False

                    ASCMAIN1.sql = "Select SOTORDR5.* from SOTORDR5 where ORDR_NO = :PARM1"
                    Dim EDI_ADR_SEQ As Int32 = 0
                    For Each rowSOTORDR5 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {ORDR_NO}).Select("", "")

                        Dim CUST_ADDR_TYPE As String = rowSOTORDR5.Item("CUST_ADDR_TYPE") & ""
                        Dim CUST_ADDR_CODE As String = rowSOTORDR5.Item("CUST_ADDR_CODE")

                        Dim rowEDT940O5 As DataRow = dst.Tables("EDT940O5").NewRow
                        With rowEDT940O5
                            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            EDI_ADR_SEQ += 1
                            .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                            If CUST_ADDR_TYPE = "BT" Then
                                BT_GEN = True
                            End If
                            .Item("EDI_ADDR_TYPE") = CUST_ADDR_TYPE
                            .Item("EDI_CUST_NAME_ADR") = Mid(rowSOTORDR5.Item("CUST_NAME") & "", 1, 35)
                            .Item("EDI_ADDRESS1") = Mid(rowSOTORDR5.Item("CUST_ADDR1") & "", 1, 35)
                            .Item("EDI_ADDRESS2") = Mid(rowSOTORDR5.Item("CUST_ADDR2") & "", 1, 35)
                            .Item("EDI_ADDRESS3") = Mid(rowSOTORDR5.Item("CUST_ADDR3") & "", 1, 35)
                            .Item("EDI_CITY") = rowSOTORDR5.Item("CUST_CITY")
                            .Item("EDI_STATE") = rowSOTORDR5.Item("CUST_STATE")
                            .Item("EDI_ZIPCODE") = rowSOTORDR5.Item("CUST_ZIP_CODE")
                            .Item("EDI_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY")

                            If .Item("EDI_STATE") & "" = "" Then .Item("EDI_STATE") = "."
                            If .Item("EDI_ZIPCODE") & "" = "" Then .Item("EDI_ZIPCODE") = "00000"

                            If CUST_ADDR_TYPE = "ST" Then
                                Dim CUST_ADDR_CODE_orig As String = CUST_ADDR_CODE
                                If MAX_CHARS_DC <> 0 And Len(CUST_ADDR_CODE) > MAX_CHARS_DC Then
                                    CUST_ADDR_CODE = Mid(CUST_ADDR_CODE, Len(CUST_ADDR_CODE) - MAX_CHARS_DC + 1, MAX_CHARS_DC)
                                End If

                                If CUST_CODE = "AAFES" Then 'DOES THIS NEED TO BE PARAMETERIZED IN EDTSLSP1? - SEND GLN AS "STORE"
                                    Dim ORDR_ADDR_TYPE_ST As String = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
                                    Dim rowARTCUST2_ST As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, ORDR_ADDR_TYPE_ST, CUST_ADDR_CODE_orig})
                                    'AAFES BETTER HAVE THIS NUMBER
                                    .Item("EDI_ADDR_CODE") = rowARTCUST2_ST.Item("GLOBAL_LOCATION_NUMBER")

                                ElseIf CUST_CODE = "WALMART" Then 'DOES THIS NEED TO BE PARAMETERIZED IN EDTSLSP1? - SEND GLN AS "STORE"
                                    .Item("EDI_ADDR_CODE") = "0" & Mid(CUST_ADDR_CODE, 1, 4)

                                Else
                                    .Item("EDI_ADDR_CODE") = CUST_ADDR_CODE
                                End If
                            End If

                            .Item("EDI_ADDR_CODE_QUAL") = CUST_ADDR_TYPE
                        End With
                        dst.Tables("EDT940O5").Rows.Add(rowEDT940O5)
                    Next


                    If CUST_CODE = "JCPLIQ" Then 'DOES THIS NEED TO BE PARAMETERIZED IN EDTSLSP1? - SEND MK AS WELL AS ST
                        Dim rowARTCUST2_JCP As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
                        If rowARTCUST2_JCP Is Nothing Then
                            rowARTCUST2_JCP = LookUp("ARTCUST2", New String() {CUST_CODE, "DC", CUST_STORE_NO})
                        End If
                        Dim rowEDT940O5 As DataRow = dst.Tables("EDT940O5").NewRow
                        With rowEDT940O5
                            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            EDI_ADR_SEQ += 1
                            .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                            .Item("EDI_ADDR_TYPE") = "MK"
                            .Item("EDI_CUST_NAME_ADR") = Mid(rowARTCUST2_JCP.Item("CUST_NAME") & "", 1, 35)
                            .Item("EDI_ADDRESS1") = Mid(rowARTCUST2_JCP.Item("CUST_ADDR1") & "", 1, 35)
                            .Item("EDI_ADDRESS2") = Mid(rowARTCUST2_JCP.Item("CUST_ADDR2") & "", 1, 35)
                            .Item("EDI_ADDRESS3") = Mid(rowARTCUST2_JCP.Item("CUST_ADDR3") & "", 1, 35)
                            .Item("EDI_CITY") = rowARTCUST2_JCP.Item("CUST_CITY")
                            .Item("EDI_STATE") = rowARTCUST2_JCP.Item("CUST_STATE")
                            .Item("EDI_ZIPCODE") = rowARTCUST2_JCP.Item("CUST_ZIP_CODE")
                            .Item("EDI_COUNTRY") = rowARTCUST2_JCP.Item("CUST_COUNTRY")

                            Dim CUST_ADDR_CODE As String = CUST_STORE_NO
                            Dim CUST_ADDR_CODE_orig As String = CUST_ADDR_CODE
                            If MAX_CHARS_DC <> 0 And Len(CUST_ADDR_CODE) > MAX_CHARS_DC Then
                                CUST_ADDR_CODE = Mid(CUST_ADDR_CODE, Len(CUST_ADDR_CODE) - MAX_CHARS_DC + 1, MAX_CHARS_DC)
                            End If
                            .Item("EDI_ADDR_CODE") = CUST_ADDR_CODE
                            .Item("EDI_ADDR_CODE_QUAL") = "MK"
                        End With
                        dst.Tables("EDT940O5").Rows.Add(rowEDT940O5)
                    End If

                    If Not BT_GEN Then
                        Dim rowEDT940O5 As DataRow = dst.Tables("EDT940O5").NewRow
                        With rowEDT940O5
                            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            EDI_ADR_SEQ += 1
                            .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ

                            .Item("EDI_ADDR_TYPE") = "BT"
                            .Item("EDI_CUST_NAME_ADR") = Mid(rowARTCUST1.Item("CUST_NAME") & "", 1, 35)
                            .Item("EDI_ADDRESS1") = Mid(rowARTCUST1.Item("CUST_ADDR1") & "", 1, 35)
                            .Item("EDI_ADDRESS2") = Mid(rowARTCUST1.Item("CUST_ADDR2") & "", 1, 35)
                            .Item("EDI_ADDRESS3") = Mid(rowARTCUST1.Item("CUST_ADDR3") & "", 1, 35)
                            .Item("EDI_CITY") = rowARTCUST1.Item("CUST_CITY")
                            .Item("EDI_STATE") = rowARTCUST1.Item("CUST_STATE")
                            .Item("EDI_ZIPCODE") = rowARTCUST1.Item("CUST_ZIP_CODE")
                            .Item("EDI_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY")
                            .Item("EDI_ADDR_CODE") = rowARTCUST1.Item("CUST_CODE")
                            .Item("EDI_ADDR_CODE_QUAL") = "BT"
                        End With
                        dst.Tables("EDT940O5").Rows.Add(rowEDT940O5)
                    End If

                    If CUST_STORE_NO <> CUST_DC_NO And CUST_DC_NO <> "" Then
                        Dim rowEDT940O5 As DataRow = dst.Tables("EDT940O5").NewRow
                        With rowEDT940O5
                            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                            .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                            EDI_ADR_SEQ += 1
                            .Item("EDI_ADR_SEQ") = EDI_ADR_SEQ

                            .Item("EDI_ADDR_TYPE") = "MK"
                            .Item("EDI_CUST_NAME_ADR") = Mid(rowARTCUST2_MK.Item("CUST_NAME") & "", 1, 35)
                            .Item("EDI_ADDRESS1") = Mid(rowARTCUST2_MK.Item("CUST_ADDR1") & "", 1, 35)
                            .Item("EDI_ADDRESS2") = Mid(rowARTCUST2_MK.Item("CUST_ADDR2") & "", 1, 35)
                            .Item("EDI_ADDRESS3") = Mid(rowARTCUST2_MK.Item("CUST_ADDR3") & "", 1, 35)
                            .Item("EDI_CITY") = rowARTCUST2_MK.Item("CUST_CITY")
                            .Item("EDI_STATE") = rowARTCUST2_MK.Item("CUST_STATE")
                            .Item("EDI_ZIPCODE") = rowARTCUST2_MK.Item("CUST_ZIP_CODE")
                            .Item("EDI_COUNTRY") = rowARTCUST2_MK.Item("CUST_COUNTRY")

                            Dim CUST_ADDR_CODE As String = CUST_STORE_NO
                            If MAX_CHARS_DC <> 0 And Len(CUST_ADDR_CODE) > MAX_CHARS_DC Then
                                CUST_ADDR_CODE = Mid(CUST_ADDR_CODE, Len(CUST_ADDR_CODE) - MAX_CHARS_DC + 1, MAX_CHARS_DC)
                            End If
                            .Item("EDI_ADDR_CODE") = CUST_ADDR_CODE

                            .Item("EDI_ADDR_CODE_QUAL") = "MK"
                        End With
                        dst.Tables("EDT940O5").Rows.Add(rowEDT940O5)
                    End If


                    '4012453780      TAYLORED

                    ASCMAIN1.sql = "Insert into EDTSYSIH (COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_APPLICATION_ID,EDI_PROCESS_IND," _
                        & "EDI_OUR_ID,EDI_TP_ID,INIT_DATE,INIT_OPER)" _
                        & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,'" & ASCMAIN1.USER_ID & "')"
                    Dim EDI_APPLICATION_ID As String = "OW"
                    Dim EDI_PROCESS_IND As String = "1"
                    ' EDI_PROCESS_IND = "T"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVVV", _
                            New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, EDI_APPLICATION_ID, EDI_PROCESS_IND, _
                                          rowEDTTRPM1.Item("EDI_OUR_ID"), rowICTWHSE1.Item("WHSE_EDI_ID")})
                Next

            Next

            For Each TABLE_NAME As String In New String() _
                {"EDT940O1", "EDT940O2", "EDT940O4", "EDT940O5", "EDT940O6"}
                For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                    For Each DCOL As DataColumn In dst.Tables(TABLE_NAME).Columns
                        If DCOL.DataType.ToString = "System.String" Then
                            If row.Item(DCOL.ColumnName) & "" <> "" Then
                                Dim DV As String = row.Item(DCOL.ColumnName)
                                ' IF DV CONTAINS ANY UNPRINTABLE CHARACTERS 
                                '   THEN REPLACE THEM WITH ""
                            End If
                        End If
                    Next
                Next
            Next

            Update_Record_TDA("EDT940O1")
            Update_Record_TDA("EDT940O2")
            Update_Record_TDA("EDT940O4")
            Update_Record_TDA("EDT940O5")
            Update_Record_TDA("EDT940O6")

        End If

        '  CommitTrans()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Write_Notes(EDI_OUTBOUND_DOC_NO As String, EDI_NTE_TYPE As String, NOTES As String)

        Dim EDI_NTE_SEQ_NO As Int32 = 0
        For Each NOTE As String In Split(NOTES, vbCrLf)
            NOTE = Trim(NOTE)
            Do While NOTE <> ""
                EDI_NTE_SEQ_NO += 1

                Dim EDI_NTE As String
                If NOTE.Length > 40 Then
                    EDI_NTE = Mid(NOTE, 1, 40)
                    NOTE = Mid(NOTE, 41)
                Else
                    EDI_NTE = NOTE
                    NOTE = ""
                End If

                EDI_NTE = Replace(EDI_NTE, "*", "@")

                Dim rowEDT940O4 As DataRow = dst.Tables("EDT940O4").NewRow
                With rowEDT940O4
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("EDI_NTE_TYPE") = EDI_NTE_TYPE
                    .Item("EDI_NTE_SEQ_NO") = EDI_NTE_SEQ_NO
                    .Item("EDI_NTE") = EDI_NTE
                End With
                dst.Tables("EDT940O4").Rows.Add(rowEDT940O4)
            Loop
        Next

    End Sub

    Private Sub optPending_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPending.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        ' dont want to be here if closing down
        Load_SOTSHIPX()
    End Sub

    Sub Load_SOTSHIPX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Shipments Queue")

        If (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
            With UltraExplorerBar1.Groups("Screen Control")
                If optPending.Value = "D" Then
                    .Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.True
                    .Items("Send").Settings.Enabled = DefaultableBoolean.False
                Else
                    .Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
                    .Items("Send").Settings.Enabled = DefaultableBoolean.True
                End If
            End With
        End If

        If optPending.Value = "0" Then
            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & ", '1' SEL, '0' EDI856, SOTSHIP1.SHIP_CART_REQD, NULL ORDR_NO" & vbCrLf _
                & ", SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST, NULL CUST_CODE from SOTSHIP1" _
                & " where SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"

            UltraExplorerBar1.Groups("Screen Control").Items("Send").Text = "Transmit"
            ASCMAIN1.sql &= " and SOTSHIP1.LP_STATUS = '0'"
            grdSOTSHIPX.Text = "Shipments Pending Transmission to 3PL (" & WHSE_CODE & ")"
            UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False

        Else
            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                 & ", '0' SEL, '0' EDI856, SOTSHIP1.SHIP_CART_REQD, NULL ORDR_NO" & vbCrLf _
                 & ", SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST, NULL CUST_CODE from SOTSHIP1" _
                 & " where SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"



            If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") Then
                UltraExplorerBar1.Groups("Screen Control").Items("Send").Text = "De-Transmit"

            End If

            ASCMAIN1.sql &= " and SOTSHIP1.LP_STATUS = '1' and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf
            grdSOTSHIPX.Text = "Shipments Sent to 3PL (" & WHSE_CODE & ")"

        End If
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
        ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " " & ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
                            & "Set ORDR_NO = (Select Min (ORDR_NO) ORDR_NO from SOTPICK1 " _
                            & " where SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO)")
        ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
                    & "Set CUST_CODE = (Select CUST_CODE from SOTORDR1 where ORDR_NO = SOTSHIPX.ORDR_NO)")

        ' NOT TRUSING THE SOTSHIP1.SHIP_856_IND
        '  - IF WE WERE, WE WOULD BE USING IT INSTEAD OF SETTING THIS FIELD TO '0' ABOVE, AND THEN FIXING IT HERE
        '  - WISH I HAD COHONES LIKE 1999
        ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
                    & "Set EDI856 = '1' where CUST_CODE in (Select Distinct CUST_CODE from EDTTRPM1 " _
                    & " where EDI_DOC_NO = '856' and EDI_STATUS = 'P')")

        Fill_Records("SOTSHIPX")
        Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)

        Setup_SOTSHIPX()

        ' Default to UnSelected. Force the user to select them
        For Each row As DataRow In dst.Tables("SOTSHIPX").Select("SHIP_XMIT_FLAG = 'H'")
            row.Item("SEL") = 0
        Next

        dst.Tables("SOTSHIPX").AcceptChanges()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Request_940_Cancel()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")

        BeginTrans()

        If Not dst.Tables.Contains("EDT940O1") Then
            Create_TDA(dst.Tables.Add, "EDT940O1", "*")
            Create_TDA(dst.Tables.Add, "EDT940O2", "*")
            Create_TDA(dst.Tables.Add, "EDT940O4", "*")
            Create_TDA(dst.Tables.Add, "EDT940O5", "*")
            Create_TDA(dst.Tables.Add, "EDT940O6", "*")
        End If

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
        Dim rowEDTTRPM1 As DataRow = LookUp("EDTTRPM1", _
                                            New String() {rowICTWHSE1.Item("WHSE_EDI_QUAL"), rowICTWHSE1.Item("WHSE_EDI_ID"), "943"})


        For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
            Dim CUST_CODE As String = rowSOTSHIPX.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim rowEDTSLSP1 As DataRow = LookUp("EDTSLSP1", CUST_CODE)
            Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("SHIP_BOL_NO")
            Dim SHIP_VIA_CODE As String = rowSOTSHIPX.Item("SHIP_VIA_CODE") & ""
            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE, True)

            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL, SHIP_PICK_PRINTED = NULL" _
               & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P'"
            For Each rowSOTPICK1 As DataRow In ASCDATA1.GetDataTable.Select("")

                Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                Dim EDI_DOC_SEQ_NO As String = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
                'Dim rowEDT850T1 As DataRow = LookUp("EDT850T1", EDI_DOC_SEQ_NO)

                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & ""
                Dim CUST_DC_NO As String = rowSOTORDR1.Item("CUST_DC_NO") & ""

                Dim rowEDT940O1 As DataRow = dst.Tables("EDT940O1").NewRow
                With rowEDT940O1
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("PICK_NO") = PICK_NO
                    .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                    .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                    .Item("EDI_SUPPLIER_NO") = rowARTCUST1.Item("CUST_VEND_REF")
                    .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                    .Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                    .Item("ORDR_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
                    .Item("ORDR_PO_DATE") = rowSOTORDR1.Item("ORDR_DATE")
                    .Item("UPS_REF_1") = CreateUpsRef1(CUST_CODE, EDI_DOC_SEQ_NO)

                    Dim FRT_TERMS As String = rowSOTSHIPX.Item("FRT_TERMS") & ""
                    Dim FRT_TERMS_EDI As String = ""
                    Select Case FRT_TERMS
                        Case "PPD", "PPA"
                            FRT_TERMS_EDI = "PP"
                        Case "COL"
                            FRT_TERMS_EDI = "CC"
                    End Select
                    .Item("FRT_TERMS") = FRT_TERMS_EDI

                    '.Item("EDI_TRANS_METH_CODE") = "?"
                    '.Item("EDI_SERVICE_LEVEL") = "?"
                    '.Item("EDI_TP_BILLING_ACCT") = "? ' IF FRT_TERMS WAS 3RD PARTY WE WOULD SEND THE 3RD PARTY ACCT NUMBER
                    .Item("EDI_SCAC_CODE") = "ROUT" ' rowSOTSVIA1.Item("SHIP_VIA_SCAC")

                    Dim EDI_DIVISION_CODE As String = rowICTWHSE1.Item("LP_WHSE_ID") & ""

                    ASCMAIN1.sql = "Select * from EDT940O1 where PICK_NO = '" & PICK_NO & "'"
                    Dim rowEDT940O1_prior As DataRow = ASCDATA1.GetDataRow
                    If rowEDT940O1_prior IsNot Nothing Then
                        EDI_DIVISION_CODE = rowEDT940O1_prior.Item("EDI_DIVISION_CODE") & ""
                    End If
                    .Item("EDI_DIVISION_CODE") = EDI_DIVISION_CODE

                    ' .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("CUST_VEND_REF")
                    .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("LABEL_TEMPLATE_CODE")
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("ORDR_TYPE_CODE") = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                    'If rowEDT850T1 IsNot Nothing Then
                    '    .Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")
                    'End If
                    .Item("EDI_MERCH_TYPE") = rowSOTORDR1.Item("EDI_MERCH_TYPE")
                    .Item("ORDR_STATUS_CODE") = "V"
                End With
                dst.Tables("EDT940O1").Rows.Add(rowEDT940O1)

                ASCMAIN1.sql = "Insert into EDTSYSIH (COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_APPLICATION_ID,EDI_PROCESS_IND," _
                    & "EDI_OUR_ID,EDI_TP_ID,INIT_DATE,INIT_OPER)" _
                    & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,'" & ASCMAIN1.USER_ID & "')"
                Dim EDI_APPLICATION_ID As String = "OW"
                Dim EDI_PROCESS_IND As String = "1"
                ' EDI_PROCESS_IND = "T"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVVV", _
                        New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, EDI_APPLICATION_ID, EDI_PROCESS_IND, _
                                      rowEDTTRPM1.Item("EDI_OUR_ID"), rowICTWHSE1.Item("WHSE_EDI_ID")})
            Next
        Next

        Update_Record_TDA("EDT940O1")
        Update_Record_TDA("EDT940O2")
        Update_Record_TDA("EDT940O4")
        Update_Record_TDA("EDT940O5")

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Function CreateUpsRef1(ByVal CUST_CODE As String, ByVal EDI_DOC_SEQ_NO As String) As String

        Dim UPS_REF_1 As String = String.Empty

        If CUST_CODE <> "JCPE" OrElse ASCMAIN1.CLIENT <> "NYA" Then
            Return UPS_REF_1
        End If

        ASCMAIN1.sql = "SELECT T1.EDI_PO_NO, T2.EDI_PO_LNO" _
                & " FROM EDT850T1 T1, EDT850T2 T2 " _
                & " WHERE T1.EDI_DOC_SEQ_NO = T2.EDI_DOC_SEQ_NO" _
                & " AND T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"

        Dim tblx As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        If tblx.Rows.Count = 0 Then
            Return UPS_REF_1
        End If

        ' Get the PO Number - do not include underscore and characters after the underscore
        UPS_REF_1 = tblx.Rows(0).Item("EDI_PO_NO") & String.Empty
        UPS_REF_1 = UPS_REF_1.Split("_")(0)

        ' Each Line number needs to be three characters with leading zeros
        For Each row As DataRow In tblx.Select("", "EDI_PO_LNO")
            UPS_REF_1 &= (row.Item("EDI_PO_LNO") & String.Empty).ToString.PadLeft(3, "0")
            ' Max field size is 30
            If UPS_REF_1.Length > 27 Then
                Exit For
            End If
        Next

        Return UPS_REF_1

    End Function
End Class