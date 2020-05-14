Public Class SOFSHIP1

    Dim sqlSOTSHIPX As String
    Dim SOTSHIPX As String

    Dim SOTSHIPX_1 As String
    Dim SOTSHIPX_2 As String
    Dim SOTSHIPX_3 As String

    Dim CUST_CODE As String
    Dim ORDR_CUST_PO As String
    Dim PICK_NO As String
    Dim SHIP_BOL_NO As String

    Dim SHIP_BOL_NOs As New List(Of String)
    Dim ORDR_GROUP_NOs As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_Temp_Table("")

        With dst

            ASCMAIN1.sql = "Select * from " & SOTSHIPX
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "", 1)
            .Tables("SOTSHIPX").Columns.Add("SELECTED")
            .Tables("SOTSHIPX").Columns("SELECTED").DefaultValue = "0"

            ASCMAIN1.sql = "Select SOTSHIP3.*" _
                & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, ARTCUST1.CUST_NAME" _
                & " from SOTSHIP3, SOTSHIP1, SOTORDR0, ARTCUST1" _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTSHIP3.SHIP_BOL_NO" _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTSHIP3", "**", 0)
            '.Tables("SOTSHIP3").Columns.Add("ACK")
            '.Tables("SOTSHIP3").Columns("ACK").DefaultValue = "0"

            ASCMAIN1.sql = "Select INIT_DATE, USER_ID INIT_OPER, COLUMN_NAME, OLD_VALUE, NEW_VALUE" _
                & " from ASTAUDT1 where TABLE_NAME = 'SOTSHIP1' and KEY_VALUE = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIPA", "**", 0, False, "V", 0)

            With .Tables.Add("SOTORDRM")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("QTY")
            End With

            ASCMAIN1.sql = "Select SOTPICK1.PICK_NO, SOTORDR1.CUST_STORE_NO, SOTPICK1.ORDR_NO" & vbCrLf _
                & ", SOTPICK1.PICK_STATUS, SOTPICK1.PICK_RELEASED, SOTPICK1.PICK_FREIGHT" & vbCrLf _
                & ", SOTPICK1.PICK_PICKER, SOTPICK1.PICK_NO_REV" & vbCrLf _
                & ", SOTPICK1.PICK_PRINTED, SOTPICK1.PICK_PACKED, SOTPICK1.PICK_SHIPPED" & vbCrLf _
                & ", SOTPICK1.PICK_BATCH_NO, SOTPICK1.SHIP_BOL_NO, SOTPICK1.INV_NO" & vbCrLf _
                & ", SOTPICK1.PICK_CNT_CARTONS, SOTPICK1.PICK_TOTAL_WGT" & vbCrLf _
                & ", SOTPICK1.INIT_OPER, SOTPICK1.LAST_OPER, SOTPICK1.INIT_DATE, SOTPICK1.LAST_DATE" & vbCrLf _
                & ", SOTPICK0.PICK_FORCED" & vbCrLf _
                & " from SOTPICK1,SOTORDR1,SOTPICK0 " & vbCrLf _
                & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTPICK0.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTPICK1.SHIP_BOL_NO, EDT850T2.EDI_COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf _
                & " from SOTPICK1,SOTPICK2,SOTORDR2,ICTCOLR1,EDT850T2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")


            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTSHIP1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", , , "V")

        End With

        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTORDRM.DataSource = dst.Tables("SOTORDRM")
        grdSOTSHIPA.DataSource = dst.Tables("SOTSHIPA")
        grdSOTSHIP3.DataSource = dst.Tables("SOTSHIP3")
        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTSHIPX, New String() {"ORDR_QTY_SHIP", "ORDR_AMT_SHIP", "SHIP_TOTAL_WGT", "SHIP_CNT_CARTONS", "PICK_NO_COUNT", "PICK_FREIGHT", "PICK_QTY", "PICK_QTY_PICK"})

        grdSOTSHIPX.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        ' grdSOTSHIPX.DisplayLayout.UseFixedHeaders = True
        With grdSOTSHIPX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
            Next


            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_ADDR_TYPE", "SHIP_ADDR_CODE", "SHIP_CART_REQD", "ORDR_PICK_TYPE", "ORDR_GROUP_NO", "PICK_BATCH_NO", "FRT_TERMS", "WHSE_CODE", "INIT_DATE", "INIT_OPER", "SREP2_CODE", "SREP_CODE", "ORDR_DEPT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_NOTES", "SHIP_DATE_RECEIVED", "SHIPPED_ACTUAL", "SHIP_DATE_PACKED", "ORDR_QTY_SHIP", "ORDR_AMT_SHIP"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_DATE_SHIPPED", "BOL_PRINTED", "SHIP_PICK_PRINTED", "LP_XMIT_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_TRAILER_NO", "SHIP_LOAD_NO", "SHIP_APPT_NO", "SHIP_DATE_PLANNED", "SHIP_DATE_ROUTED", "SHIP_NOTES_3PL", "SHIP_VIA_CODE", "BILL_OF_LADING_NO"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next
        End With

        'grdSOTSHIP3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        With grdSOTSHIP3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                'If gcol.Key = "ACK" Then
                '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                '    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                'End If
            Next
            .Columns("ACK").Hidden = True
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_SHIP_DATE_OLD", "ORDR_CANCEL_DATE_OLD"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_SHIP_DATE_NEW", "ORDR_CANCEL_DATE_NEW"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
            Next
            For Each COLUMN_NAME As String In New String() {"INIT_DATE", "INIT_OPER", "SHIP_CHGREQ_REASON", "SHIP_CHGREQ_CONTACT", "SHIP_CHGREQ_EMAIL"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"LAST_DATE", "LAST_OPER", "ACK"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
        End With

        Set_Enable_Update()

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
        End If

        calFrom.Value = Now.Date.AddDays(-7)
        calTo.Value = Now.Date.AddDays(7)
        calDate.Value = Now.Date

        Bind_Controls(grpSOTSHIP1, "SOTSHIPX")

        Show_Filter(grdSOTSHIPX, True)
        grdSOTSHIPX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTSHIPX, "SHIP_STATUS", Nothing, New String() {":", "P:Pick", "F:Shipped", "C:Cancelled", "D:Deleted"})
        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS", Nothing, New String() {":", "P:Pick", "F:Shipped", "C:Cancelled", "D:Deleted"})
        ASCMAIN1.Add_Value_List(grdSOTSHIPX, "ORDR_PICK_TYPE", Nothing, New String() {":", "P:Pick&Pack", "C:Full Case"})
        splSOTSHIPX.Panel2Collapsed = Not (chkShowDetails.Checked)

        ASCMAIN1.Add_Value_List(cmbType, "SHIP_EVENT")

        Dim COLUMN_NAMEs As New List(Of String)
        COLUMN_NAMEs.Add(":")
        For Each COLUMN_NAME As String In New String() {"SHIP_DATE_ROUTED", "SHIP_DATE_PLANNED", "SHIP_DATE_PACKED", _
                                                        "SHIP_APPT_NO", "SHIP_NOTES", "SHIP_NOTES_3PL"}
            COLUMN_NAMEs.Add(COLUMN_NAME & ":" & grdSOTSHIPX.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption)
        Next
        ASCMAIN1.Add_Value_List(grdSOTSHIPA, "COLUMN_NAME", Nothing, COLUMN_NAMEs.ToArray)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"
                CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                PICK_NO = Absx1.txtFor("PICK_NO").Text
                SHIP_BOL_NO = Absx1.txtFor("SHIP_BOL_NO").Text

                If PICK_NO <> "" Then
                    Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", PICK_NO)
                    If rowSOTPICK1 Is Nothing Then
                        EMsg &= vbCr & "Cannot Find Pick Ticket No " & PICK_NO
                    Else
                        SHIP_BOL_NO = rowSOTPICK1.Item("SHIP_BOL_NO")
                    End If
                    CUST_CODE = ""
                    ORDR_CUST_PO = ""

                ElseIf CUST_CODE <> "" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Cannot Find Customer " & CUST_CODE
                    End If
                End If

            Case "Edit Status"
                SHIP_BOL_NO = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value
                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub

                Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                If rowSOTSHIP1.Item("SHIP_STATUS") & "" <> "P" Then
                    EMsg &= "Shipment " & SHIP_BOL_NO & " is No Longer In Pick"
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Update"

            Case "Print"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                SHIP_BOL_NO = ""
                Load_SOTSHIPX()

            Case "Edit Status"
                Load_SOTSHIPX()
                Mode_Settings(True)

            Case "Update Status"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel Edits"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                End With

                .Groups("Shipment Status").Visible = Not ScreenMode
                .Groups("Shipping Window").Visible = False ' Not ScreenMode

                With .Groups("Shipment Status Updates")
                    .Items("Edit Status").Settings.Enabled = not_iScreenMode
                    .Items("Update Status").Settings.Enabled = iScreenMode
                    .Items("Cancel Edits").Settings.Enabled = iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpSOTSHIP1, Not ScreenMode)
        Setup_tabSOTSHIPX()

        If ScreenMode Then
            Dim dvw As DataView = DirectCast(grdSOTSHIPX.DataSource, DataTable).DefaultView
            dvw.RowFilter = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            Show_Filter(grdSOTSHIPX, False)
        Else
            Clear_Record()
            grdSOTSHIPX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            Show_Filter(grdSOTSHIPX, True)
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTSHIPX"}
            ' dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("ORDR_CUST_PO").Text = ""
        Absx1.txtFor("PICK_NO").Text = ""
        CUST_CODE = ""
        ORDR_CUST_PO = ""
        PICK_NO = ""
        'SHIP_BOL_NO = ""
        Set_Date_Type()

        Load_SOTSHIPX()
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()

        Synch_TABLE_NAME("SOTSHIPX")

        BeginTrans()

        dst.Tables("ASTAUDT1").Rows.Clear()
        Dim row As DataRow = dst.Tables("SOTSHIPX").Rows(0)

        For Each COLUMN_NAME As String In New String() _
            {"SHIP_DATE_RECEIVED", "SHIPPED_ACTUAL", "SHIP_DATE_PLANNED", "SHIP_NOTES", "SHIP_NOTES_3PL", _
             "SHIP_VIA_CODE", "SHIP_APPT_NO", "SHIP_LOAD_NO", "SHIP_DATE_PACKED", "SHIP_DATE_ROUTED"}
            If row.Item(COLUMN_NAME) & "" <> row.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                Record_Audit(SHIP_BOL_NO, COLUMN_NAME, row.Item(COLUMN_NAME, DataRowVersion.Original) & "", row.Item(COLUMN_NAME) & "")
            End If
        Next
        Update_Record_TDA("ASTAUDT1")

        ASCMAIN1.sql = "Update SOTSHIP1 Set " & vbCrLf _
            & "  SHIP_DATE_RECEIVED = :PARM1" & vbCrLf _
            & ", SHIPPED_ACTUAL = :PARM2" & vbCrLf _
            & ", SHIP_DATE_PLANNED = :PARM3" & vbCrLf _
            & ", SHIP_NOTES = :PARM4" & vbCrLf _
            & ", SHIP_NOTES_3PL = :PARM5" & vbCrLf _
            & ", SHIP_VIA_CODE = :PARM6" & vbCrLf _
            & ", SHIP_APPT_NO = :PARM7" & vbCrLf _
            & ", SHIP_LOAD_NO = :PARM8" & vbCrLf _
            & ", SHIP_DATE_PACKED = :PARM9" & vbCrLf _
            & ", SHIP_DATE_ROUTED = :PARM10" & vbCrLf _
            & " where SHIP_BOL_NO = :PARM11"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DDDVVVVVDDV", New Object() { _
                            row.Item("SHIP_DATE_RECEIVED"), _
                            row.Item("SHIPPED_ACTUAL"), _
                            row.Item("SHIP_DATE_PLANNED"), _
                            row.Item("SHIP_NOTES"), _
                            row.Item("SHIP_NOTES_3PL"), _
                            row.Item("SHIP_VIA_CODE"), _
                            row.Item("SHIP_APPT_NO"), _
                            row.Item("SHIP_LOAD_NO"), _
                            row.Item("SHIP_DATE_PACKED"), _
                            row.Item("SHIP_DATE_ROUTED"), _
                            row.Item("SHIP_BOL_NO")})
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Generate_Report("PORWREC2")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "USER_ID"
                sql_where = "USER_ID in (Select Distinct WO_ASSIGNED_TO from SOTWORK1 union Select Distinct WO_ASSIGNED_TO from SOTWORK2)"
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPX, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Order Inquiry", "Edit Shipping Window", "Wave Inquiry")
        Load_Popup_Menu(grdSOTSHIP3, "SSSS", "Show Filter", "Show GroupBox", "Show Pins", "Enable Ack")
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

        Select Case e.SourceControl.Name
            Case "grdSOTSHIPX"
                If ScreenMode Then e.Cancel = True
                If grpSHIP_WINDOW.Tag & "" = "X" Then e.Cancel = True
                '  Exit Sub

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdSOTSHIPX"
                Dim WAVE_NO As String = ""
                If grd.ActiveRow IsNot Nothing And grd.ActiveRow.IsDataRow Then
                    WAVE_NO = grd.ActiveRow.Cells("WAVE_NO").Value & ""
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow

                tlb_btn = DirectCast(tlb_pop.Tools("Edit Shipping Window"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode ' not while editing shipments

                ' security

                tlb_btn = DirectCast(tlb_pop.Tools("Wave Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (WAVE_NO <> "") ' And ScreenMode

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Edit Shipping Window"
                SHIP_BOL_NOs.Clear()
                ORDR_GROUP_NOs.Clear()

                Dim EMsg As String = ""

                If grdSOTSHIPX.Selected.Rows.Count = 0 Then
                    If grdSOTSHIPX.ActiveRow IsNot Nothing AndAlso grdSOTSHIPX.ActiveRow.IsDataRow Then
                        grdSOTSHIPX.ActiveRow.Selected = True
                    End If
                End If

                If grdSOTSHIPX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must First Select Shipments to Edit"
                Else
                    Dim CUST_CODE_selected As String = ""
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTSHIPX.Selected.Rows
                        If CUST_CODE_selected = "" Then
                            CUST_CODE_selected = grow.Cells("CUST_CODE").Value
                        Else
                            If CUST_CODE_selected <> grow.Cells("CUST_CODE").Value Then
                                EMsg &= vbCr & "You May NOT Select Shipments from More than 1 Customer to Edit the Shipping Window"
                                Exit For
                            End If
                        End If
                        Dim SHIP_BOL_NO As String = grow.Cells("SHIP_BOL_NO").Value
                        Dim ORDR_GROUP_NO As String = grow.Cells("ORDR_GROUP_NO").Value
                        SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                        dteORDR_SHIP_DATE.Value = grow.Cells("ORDR_SHIP_DATE").Value
                        dteORDR_CANCEL_DATE.Value = grow.Cells("ORDR_CANCEL_DATE").Value
                        If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                        If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then
                            ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        End If

                        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                        If rowSOTSHIP1.Item("SHIP_STATUS") & "" <> "P" Then
                            EMsg &= "Shipment " & SHIP_BOL_NO & " is No Longer In Pick"
                        End If
                    Next
                End If

                If EMsg <> "" Then
                    MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Edit Shipping Window")
                    ASCMAIN1.MultiTask_Release()
                Else
                    For Each row As DataRow In dst.Tables("SOTSHIPX").Select("")
                        row.Item("SELECTED") = "0"
                        If SHIP_BOL_NOs.Contains(row.Item("SHIP_BOL_NO")) Then
                            row.Item("SELECTED") = "1"
                        End If
                    Next
                    grdSOTSHIPX.Selected.Rows.Clear()
                    chkShowDetails.Checked = False
                    Toggle_Shipping_Window(True)
                End If

            Case "Enable Ack"
                tlb_sbt = DirectCast(tlb.Tools("Enable Ack"), UltraWinToolbars.StateButtonTool)
                grdSOTSHIP3.DisplayLayout.Bands(0).Columns("ACK").Hidden = Not tlb_sbt.Checked

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")

            Case "Wave Inquiry"
                Dim WAVE_NO As String = grd.ActiveRow.Cells("WAVE_NO").Value
                If WAVE_NO <> "" Then
                    Context_Launch("View", WAVE_NO, e.Tool.Key, "WHFWAVEI")
                End If
        End Select
    End Sub

#End Region

    Sub Toggle_Shipping_Window(tf As Boolean)
        UltraExplorerBar1.Groups("Shipping Window").Visible = tf
        UltraExplorerBar1.Groups("Scope (Blank for All)").Visible = Not tf
        UltraExplorerBar1.Groups("Shipment Status").Visible = Not tf
        UltraExplorerBar1.Groups("Screen Control").Visible = Not tf
        tabMain.Tabs("Changes to Shipping Window").Visible = Not tf
        Dim dvw As DataView = DirectCast(grdSOTSHIPX.DataSource, DataTable).DefaultView
        If Not tf Then
            grpSHIP_WINDOW.Tag = ""
            ASCMAIN1.MultiTask_Release()
            dvw.RowFilter = ""
            grdSOTSHIPX.Selected.Rows.Clear()
        Else
            grpSHIP_WINDOW.Tag = "X"
            dvw.RowFilter = "SELECTED = '1'"
            txtSHIP_WINDOW_REASON.Text = ""
            txtSHIP_WINDOW_CONTACT.Text = ""
            txtSHIP_WINDOW_EMAIL.Text = ""
        End If
        chkEnableUpdate.Visible = Not tf
        Show_Filter(grdSOTSHIPX, Not tf)
        grdSOTSHIPX.DisplayLayout.GroupByBox.Hidden = tf
    End Sub

    Private Sub cmdSHIP_WINDOW_Update_Click(sender As System.Object, e As System.EventArgs) Handles cmdSHIP_WINDOW_Update.Click
        Update_Shipping_Window()
        Toggle_Shipping_Window(False)
    End Sub

    Private Sub cmdSHIP_WINDOW_Cancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdSHIP_WINDOW_Cancel.Click
        Toggle_Shipping_Window(False)
    End Sub

    Sub Update_Shipping_Window()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        dst.Tables("SOTSHIP3").Rows.Clear()
        dst.Tables("ASTAUDT1").Rows.Clear()

        BeginTrans()

        For Each SHIP_BOL_NO In SHIP_BOL_NOs
            Dim rowSOTSHIPX As DataRow = dst.Tables("SOTSHIPX").Rows.Find(SHIP_BOL_NO)
            Dim ORDR_GROUP_NO As String = rowSOTSHIPX.Item("ORDR_GROUP_NO")
            Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)

            Dim date_changed As Boolean = False
            Dim ORDR_SHIP_DATE As Date = dteORDR_SHIP_DATE.Value
            Dim ORDR_CANCEL_DATE As Date = dteORDR_CANCEL_DATE.Value

            If Format(rowSOTORDR0.Item("ORDR_SHIP_DATE"), "MM/dd/yy") <> Format(ORDR_SHIP_DATE, "MM/dd/yy") Then
                Record_Event("SHIP_DATE", "Ship Date Changed from " & Format(rowSOTORDR0.Item("ORDR_SHIP_DATE"), "MM/dd/yy") & " -> " & Format(ORDR_SHIP_DATE, "MM/dd/yy"))
                date_changed = True
                rowSOTSHIPX.Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE
            End If
            If Format(rowSOTORDR0.Item("ORDR_CANCEL_DATE"), "MM/dd/yy") <> Format(ORDR_CANCEL_DATE, "MM/dd/yy") Then
                Record_Event("CANCEL_DATE", "Cancel Date Changed from " & Format(rowSOTORDR0.Item("ORDR_CANCEL_DATE"), "MM/dd/yy") & " -> " & Format(ORDR_CANCEL_DATE, "MM/dd/yy"))
                date_changed = True
                rowSOTSHIPX.Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
            End If

            If date_changed Then
                Dim SHIP_CHGREQ_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP3.SHIP_CHGREQ_NO")
                Dim rowSOTSHIP3 As DataRow = dst.Tables("SOTSHIP3").NewRow
                With rowSOTSHIP3
                    .Item("SHIP_CHGREQ_NO") = SHIP_CHGREQ_NO
                    .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                    .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("ORDR_SHIP_DATE_OLD") = rowSOTORDR0.Item("ORDR_SHIP_DATE")
                    .Item("ORDR_CANCEL_DATE_OLD") = rowSOTORDR0.Item("ORDR_CANCEL_DATE")
                    .Item("ORDR_SHIP_DATE_NEW") = ORDR_SHIP_DATE
                    .Item("ORDR_CANCEL_DATE_NEW") = ORDR_CANCEL_DATE
                    .Item("SHIP_CHGREQ_REASON") = txtSHIP_WINDOW_REASON.Text
                    .Item("SHIP_CHGREQ_CONTACT") = txtSHIP_WINDOW_CONTACT.Text
                    .Item("SHIP_CHGREQ_EMAIL") = txtSHIP_WINDOW_EMAIL.Text
                End With
                dst.Tables("SOTSHIP3").Rows.Add(rowSOTSHIP3)

                ASCMAIN1.sql = "Update SOTORDR1 Set ORDR_SHIP_DATE = :PARM1, ORDR_CANCEL_DATE = :PARM2" & vbCrLf _
                    & " where ORDR_NO in " & vbCrLf _
                    & " (Select ORDR_NO from SOTPICK1 where SHIP_BOL_NO = :PARM3)"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DDV", New Object() {ORDR_SHIP_DATE, ORDR_CANCEL_DATE, SHIP_BOL_NO})

                ASCMAIN1.sql = "Update " & SOTSHIPX & " Set ORDR_SHIP_DATE = :PARM1, ORDR_CANCEL_DATE = :PARM2 where SHIP_BOL_NO = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DDV", New Object() {ORDR_SHIP_DATE, ORDR_CANCEL_DATE, SHIP_BOL_NO})
            End If
        Next

        Update_Record_TDA("SOTSHIP3")
        Update_Record_TDA("ASTAUDT1")

        ' Group Record
        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
            ASCDATA1.ExecuteSQL()
        Next

        Fill_Records("SOTSHIP3")

        CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_VIA_CODE"
                If ScreenMode Then
                    Absx1.dteFor("SHIP_DATE_ROUTED").Value = DATETIME_STAMP.Date
                End If
        End Select
    End Sub
#End Region

    Sub Load_SOTSHIPX()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Data")

        If ASCMAIN1.DBS_COMPANY = "VAN" And ASCMAIN1.DBS_SERVER = "VAN" Then
            TAC.WHCMAIN1.Update_ADS_SOTSHIP1()
        End If

        EnforceConstraints(False)

        If SHIP_BOL_NO <> "" Then
            ASCMAIN1.sql = " and SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            grdSOTSHIPX.Text = "Shipment " & SHIP_BOL_NO & IIf(PICK_NO = "", "", " (includes Pick Ticket " & PICK_NO & ")")
        Else
            ASCMAIN1.sql = "" _
                           & IIf(CUST_CODE = "", "", " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'") _
                           & IIf(ORDR_CUST_PO = "", "", " and SOTORDR0.ORDR_CUST_PO = '" & ORDR_CUST_PO & "'")

            Select Case optStatus.Value
                Case "RNT"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.LP_XMIT_DATE is Null"
                    grdSOTSHIPX.Text = "Shipments Released not Transmitted"
                Case "RNP"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is Null"
                    grdSOTSHIPX.Text = "Shipments Released not Printed"
                Case "PNS"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is Not Null"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIPPED_ACTUAL is Null"
                    grdSOTSHIPX.Text = "Shipments Printed not Shipped"
                Case "PNC"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is Not Null"
                    grdSOTSHIPX.Text = "Shipments Printed not Confirmed (Billed)"
                Case "C"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'F'"
                    If optDATE_TYPE.Value <> "*" Then
                        If Not chkFirst.Checked Then ASCMAIN1.sql &= " and SOTSHIP1.SHIP_DATE_SHIPPED >= '" & Format(calFrom.Value, "dd-MMM-yyyy") & "'"
                        If Not chkLast.Checked Then ASCMAIN1.sql &= " and SOTSHIP1.SHIP_DATE_SHIPPED <= '" & Format(calTo.Value, "dd-MMM-yyyy") & "'"
                    End If
                    grdSOTSHIPX.Text = "Shipments Confirmed (Billed)" & IIf(optDATE_TYPE.Value = "*", "", " with " & optDATE_TYPE.Text & " between " & IIf(chkFirst.Checked, "First", calFrom.Value) & " and " & IIf(chkLast.Checked, "Last", calTo.Value))
                Case "P"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    Dim DATE_TYPE As String = optDATE_TYPE.Value
                    If optDATE_TYPE.Value <> "*" Then
                        If Not chkFirst.Checked Then ASCMAIN1.sql &= " and " & DATE_TYPE & " >= '" & Format(calFrom.Value, "dd-MMM-yyyy") & "'"
                        If Not chkLast.Checked Then ASCMAIN1.sql &= " and " & DATE_TYPE & " <= '" & Format(calTo.Value, "dd-MMM-yyyy") & "'"
                    End If
                    grdSOTSHIPX.Text = "Shipments In Pick" & IIf(optDATE_TYPE.Value = "*", "", " with " & optDATE_TYPE.Text & " between " & IIf(chkFirst.Checked, "First", calFrom.Value) & " and " & IIf(chkLast.Checked, "Last", calTo.Value))
            End Select
        End If

        If CUST_CODE <> "" Then grdSOTSHIPX.Text &= " associated with " & CUST_CODE
        ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS in ('F','P')"
        Create_Temp_Table(ASCMAIN1.sql)

        Fill_Records("SOTSHIPX")

        Dim sqlx As String = " where SHIP_BOL_NO in (Select SHIP_BOL_NO from SOTSHIP1 where SHIP_STATUS = 'P')"

        ASCMAIN1.sql = "Select SOTSHIP3.*" & vbCrLf _
            & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, ARTCUST1.CUST_NAME" & vbCrLf _
            & " from SOTSHIP3, SOTSHIP1, SOTORDR0, ARTCUST1" & vbCrLf _
            & " where SOTSHIP1.SHIP_BOL_NO = SOTSHIP3.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
            & "   and SOTSHIP1.SHIP_STATUS = 'P'"
        Fill_Records("SOTSHIP3", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdSOTSHIP3, "SHIP_CHGREQ_NO".ToLower)

        Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
        grdSOTSHIPX.Visible = True

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If optStatus.Value = "P" Then
            optDATE_TYPE.Value = "*"
            chkFirst.Checked = True
        Else
            If optDATE_TYPE.Value = "*" Then
                optDATE_TYPE.Value = "SOTORDR0.ORDR_CANCEL_DATE"
            End If
        End If
        Set_Date_Type()
        Click_Command("Refresh")
    End Sub

    Sub Set_Date_Type()
        grpConfirmed.Visible = (optStatus.Value = "C" Or optStatus.Value = "P")
        optDATE_TYPE.Visible = (optStatus.Value = "P")

        calFrom.Visible = Not chkFirst.Checked
        lblXFrom.Visible = Not chkFirst.Checked

        calTo.Visible = Not chkLast.Checked
        lblXTo.Visible = Not chkLast.Checked

        If optStatus.Value = "C" Then
            lblDATE_TYPE.Text = "Confirmed as Shipped"
        Else
            If optStatus.Value = "P" Then
                If optDATE_TYPE.Value = "*" Then
                    lblDATE_TYPE.Text = "All"
                Else
                    lblDATE_TYPE.Text = optDATE_TYPE.Text & " Range"
                End If
            End If
        End If
    End Sub

    Private Sub grdSOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSHIPX.AfterRowActivate
        If grpSHIP_WINDOW.Tag & "" = "X" Then
            ' DO NOTHING
        Else
            Setup_grdSOTSHIPX()
        End If
    End Sub

    Private Sub grdSOTSHIPX_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTSHIPX.AfterRowUpdate
        Dim row As DataRow = LookUp("SOTSHIP1", e.Row.Cells("SHIP_BOL_NO").Value)
        ASCMAIN1.sql = "Update SOTSHIP1 Set SHIP_DATE_RECEIVED = :PARM1, SHIP_NOTES = :PARM2, SHIP_DATE_PACKED = :PARM3, SHIPPED_ACTUAL = :PARM4" _
            & " where SHIP_BOL_NO = :PARM5"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVDDV", New Object() { _
                            e.Row.Cells("SHIP_DATE_RECEIVED").Value, _
                            e.Row.Cells("SHIP_NOTES").Value, _
                            e.Row.Cells("SHIP_DATE_PACKED").Value, _
                            e.Row.Cells("SHIPPED_ACTUAL").Value, _
                            e.Row.Cells("SHIP_BOL_NO").Value})

        dst.Tables("ASTAUDT1").Rows.Clear()
        Dim INIT_DATE As Date = Now + ASCMAIN1.NowTSD
        For Each COLUMN_NAME As String In New String() {"SHIP_DATE_RECEIVED", "SHIP_NOTES", "SHIP_DATE_PACKED", "SHIPPED_ACTUAL"}
            If row.Item(COLUMN_NAME) & "" <> e.Row.Cells(COLUMN_NAME).Value & "" Then
                Record_Audit(e.Row.Cells("SHIP_BOL_NO").Value, _
                             COLUMN_NAME, _
                             row.Item(COLUMN_NAME) & "", _
                             e.Row.Cells(COLUMN_NAME).Value & "")
            End If
        Next
        Update_Record_TDA("ASTAUDT1")
        Fill_Records("SOTSHIPA", e.Row.Cells("SHIP_BOL_NO").Value)
    End Sub

    Private Sub grdSOTSHIPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIPX.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("SHIP_STATUS").Value = "D" Then
                e.Row.Cells("SHIP_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Deleted"
            ElseIf e.Row.Cells("SHIP_STATUS").Value = "F" Then
                e.Row.Cells("SHIP_STATUS").Appearance.BackColor = Drawing.Color.LightGreen
                e.Row.ToolTipText = "Shipped"
            Else
                If e.Row.Cells("SHIP_WAVE_STATUS").Value & "" = "1" Then
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Blue
                    e.Row.Cells("SHIP_BOL_NO").ToolTipText = "Waved"
                Else
                    e.Row.Cells("SHIP_BOL_NO").Appearance.ForeColor = Drawing.Color.Empty
                    e.Row.Cells("SHIP_BOL_NO").ToolTipText = ""
                End If
            End If

        End If
    End Sub

    Sub Setup_grdSOTSHIPX()
        If Not chkShowDetails.Checked OrElse grdSOTSHIPX.ActiveRow Is Nothing OrElse Not grdSOTSHIPX.ActiveRow.IsDataRow Then
            tabSOTSHIPX.Visible = False
        Else
            tabSOTSHIPX.Visible = True
            EnforceConstraints(False)
            SHIP_BOL_NO = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value

            Fill_Records("SOTSHIPA", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTSHIPA, "INIT_DATE")
            grdSOTSHIPA.Text = "Audit Trail for Shipment " & SHIP_BOL_NO

            Fill_Records("SOTPICK1", SHIP_BOL_NO)
            Fill_Records("SOTPICK2", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICK1.Text = "Pick Tickets for Shipment " & SHIP_BOL_NO

            If tabSOTSHIPX.SelectedTab.Key = "Styles" Then Setup_Summary_SOTORDRM(SHIP_BOL_NO)

            lblShipment.Text = "Details for Shipment " & SHIP_BOL_NO

            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", Absx1.txtFor("SHIP_VIA_CODE").Text)
            If rowSOTSVIA1 IsNot Nothing Then
                Absx1.txtFor("SHIP_VIA_DESC").Text = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
            Else
                Absx1.txtFor("SHIP_VIA_DESC").Clear()
            End If

            ' Set up events for this shipment
            Dim CUST_CODE As String = grdSOTSHIPX.ActiveRow.Cells("CUST_CODE").Value
            Dim ORDR_CUST_PO As String = grdSOTSHIPX.ActiveRow.Cells("ORDR_CUST_PO").Value
            Fill_Records("TATEVNT1", New String() {SHIP_BOL_NO})
            grdTATEVNT1.Text = "Events for Shipment " & SHIP_BOL_NO & "; " & CUST_CODE & " PO " & ORDR_CUST_PO
            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
            cmbType.Value = DBNull.Value
            txtNote.Text = ""

            EnforceConstraints(True)
        End If
        Setup_tabSOTSHIPX()
    End Sub

    Sub Create_Temp_Table(SQLW As String)
        If SOTSHIPX = "" Then
            sqlSOTSHIPX = "Select SOTSHIP1.*" & vbCrLf _
                 & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                 & ",ARTCUST1.CUST_NAME" & vbCrLf _
                 & " from SOTSHIP1,SOTORDR0,ARTCUST1" & vbCrLf _
                 & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                 & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            SOTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_NO_MIN VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_NO_MAX VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_NO_MIN VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_NO_MAX VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_NO_COUNT NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_QTY NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_QTY_PICK NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add PICK_FREIGHT NUMBER (13,2)")

            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_QTY_SHIP NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_AMT_SHIP NUMBER (13,2)")

            ASCMAIN1.AnalyzeTable(SOTSHIPX)

            ASCMAIN1.sql = "Select SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "  , SUM (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
                & "  , SUM (DECODE(SOTPICK1.PICK_STATUS,'P',SOTPICK2.PICK_QTY,0)) PICK_QTY_PICK" & vbCrLf _
                & "   from SOTPICK2,SOTPICK1," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
                & "   where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "     and SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   group by SOTSHIPX.SHIP_BOL_NO" & vbCrLf
            SOTSHIPX_1 = ASCMAIN1.Temp_Table

        Else
            ASCMAIN1.sql = Replace(sqlSOTSHIPX, " from ", ", NULL PICK_NO_MIN, NULL PICK_NO_MAX, NULL ORDR_NO_MIN, NULL ORDR_NO_MAX, 0 PICK_NO_COUNT, 0 PICK_QTY, 0 PICK_QTY_PICK, 0 PICK_FREIGHT, 0 ORDR_QTY_SHIP, 0 ORDR_AMT_SHIP from ") & SQLW
            ASCDATA1.ExecuteSQL("Delete from " & SOTSHIPX)
            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " " & ASCMAIN1.sql)

            ASCMAIN1.AnalyzeTable(SOTSHIPX)

            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX_1)

            ASCMAIN1.sql = "Select SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "  , SUM (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
                & "  , SUM (DECODE(SOTPICK1.PICK_STATUS,'P',SOTPICK2.PICK_QTY,0)) PICK_QTY_PICK" & vbCrLf _
                & "   from SOTPICK2,SOTPICK1," & SOTSHIPX & " SOTSHIPX" & vbCrLf _
                & "   where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "     and SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   group by SOTSHIPX.SHIP_BOL_NO" & vbCrLf
            SOTSHIPX_1 = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select * from " & SOTSHIPX_1 & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTSHIPX & " Set " & vbCrLf _
                & "     PICK_QTY = R1.PICK_QTY" & vbCrLf _
                & "   , PICK_QTY_PICK = R1.PICK_QTY_PICK" & vbCrLf _
                & "   where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select SOTINVH1.SHIP_BOL_NO" & vbCrLf _
                & "  , SUM (SOTINVH2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & "  , SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
                & "   from SOTINVH2,SOTINVH1" & vbCrLf _
                & "   where SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
                & "     and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
                & "     and SOTINVH1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")" & vbCrLf _
                & "   group by SOTINVH1.SHIP_BOL_NO;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTSHIPX & " Set " & vbCrLf _
                & "     ORDR_QTY_SHIP = R1.ORDR_QTY_SHIP" & vbCrLf _
                & "   , ORDR_AMT_SHIP = R1.ORDR_AMT_SHIP" & vbCrLf _
                & "   where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "  , MIN (SOTPICK1.PICK_NO) PICK_NO_MIN, MAX (SOTPICK1.PICK_NO) PICK_NO_MAX" & vbCrLf _
                & "  , MIN (SOTPICK1.ORDR_NO) ORDR_NO_MIN, MAX (SOTPICK1.ORDR_NO) ORDR_NO_MAX" & vbCrLf _
                & "  , COUNT (SOTPICK1.PICK_NO) PICK_NO_COUNT" & vbCrLf _
                & "  , SUM (SOTPICK1.PICK_FREIGHT) PICK_FREIGHT" & vbCrLf _
                & "   from SOTPICK1" & vbCrLf _
                & "   where SOTPICK1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIPX & ")" & vbCrLf _
                & "   group by SOTPICK1.SHIP_BOL_NO;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTSHIPX & " Set " & vbCrLf _
                & "     PICK_NO_MIN = R1.PICK_NO_MIN" & vbCrLf _
                & "   , PICK_NO_MAX = R1.PICK_NO_MAX" & vbCrLf _
                & "   , ORDR_NO_MIN = R1.ORDR_NO_MIN" & vbCrLf _
                & "   , ORDR_NO_MAX = R1.ORDR_NO_MAX" & vbCrLf _
                & "   , PICK_NO_COUNT = R1.PICK_NO_COUNT" & vbCrLf _
                & "   , PICK_FREIGHT = R1.PICK_FREIGHT" & vbCrLf _
                & "   where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

        End If
    End Sub

    Private Sub optDATE_TYPE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optDATE_TYPE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Date_Type()
        Click_Command("Refresh")
    End Sub

    Private Sub chkEnableUpdate_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkEnableUpdate.CheckedChanged
        Set_Enable_Update()
    End Sub

    Sub Set_Enable_Update()

        With grdSOTSHIPX.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            If chkEnableUpdate.Checked Then
                .AllowUpdate = DefaultableBoolean.True
            Else
                .AllowUpdate = DefaultableBoolean.False
            End If
            .AllowDelete = DefaultableBoolean.False
        End With

        With grdSOTSHIPX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If chkEnableUpdate.Checked And _
                    New String() {"SHIP_NOTES", "SHIP_DATE_RECEIVED", "SHIPPED_ACTUAL", "SHIP_DATE_PACKED"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.LightCyan
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
        End With

        Setup_tabSOTSHIPX()
    End Sub

    Private Sub calFrom_ValueChanged(sender As Object, e As System.EventArgs) Handles calFrom.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Dim D As Integer = CDate(calFrom.Value).Subtract(Now.Date).Days
        If D >= 0 Then
            lblXFrom.Text = "+" & CStr(D)
        Else
            lblXFrom.Text = D
        End If
        If Not IsLoading Then Click_Command("Refresh")
    End Sub

    Private Sub calTo_ValueChanged(sender As Object, e As System.EventArgs) Handles calTo.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Dim D As Integer = CDate(calTo.Value).Subtract(Now.Date).Days
        If D >= 0 Then
            lblXTo.Text = "+" & CStr(D)
        Else
            lblXTo.Text = D
        End If
        If Not IsLoading Then Click_Command("Refresh")
    End Sub

    Private Sub optFieldName_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optFieldName.ValueChanged
        lblFieldName.Text = optFieldName.Text
        calDate.Visible = Not (optFieldName.Value = "SHIP_NOTES")
        txtSHIP_NOTE.Visible = (optFieldName.Value = "SHIP_NOTES")
    End Sub

    Private Sub btnApply_Click(sender As System.Object, e As System.EventArgs) Handles btnApply.Click
        If grdSOTSHIPX.Selected.Rows.Count = 0 Then
            MsgBox("No Rows Selected", MsgBoxStyle.OkOnly, "Cannot Perform Reqested Action")
        Else
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTSHIPX.Selected.Rows
                If optFieldName.Value = "SHIP_NOTES" Then
                    grow.Cells(optFieldName.Value).Value = txtSHIP_NOTE.Text
                Else
                    grow.Cells(optFieldName.Value).Value = calDate.Value
                End If
                grow.Update()
            Next
        End If
    End Sub

    Private Sub chkFirst_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkFirst.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Date_Type()
        Click_Command("Refresh")
    End Sub

    Private Sub chkLast_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkLast.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Date_Type()
        Click_Command("Refresh")
    End Sub

    Sub Setup_Summary_SOTORDRM(SHIP_BOL_NO As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Summary by Store")

        Dim COLUMN_NAME As String = "PICK_QTY"
        Dim SHIP_STATUS As String = "P"
        If grdSOTSHIPX.ActiveRow IsNot Nothing Then
            SHIP_STATUS = grdSOTSHIPX.ActiveRow.Cells("SHIP_STATUS").Value & ""
            If SHIP_STATUS = "F" Then
                COLUMN_NAME = "PICK_QTY_CONF"
            End If
        End If

        ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
        ASCMAIN1.sql &= ",Sum (" & COLUMN_NAME & ") QTY"
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTPICK1"), New String() {"CUST_STORE_NO"}).Select("", "CUST_STORE_NO")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            ASCMAIN1.sql &= ", Sum (Decode(SOTORDR1.CUST_STORE_NO,'" & CUST_STORE_NO & "',SOTPICK2." & COLUMN_NAME & ",0)) QTY_" & CUST_STORE_NO
        Next
        ASCMAIN1.sql &= " from SOTORDR1,SOTORDR2,SOTPICK2,SOTPICK1" _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & "   and SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
            & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
        grdSOTORDRM.DataSource = Nothing
        grdSOTORDRM.DisplayLayout.Bands(0).Summaries.Clear()
        grdSOTORDRM.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        dst.Tables.Remove("SOTORDRM")
        Dim t As DataTable = ASCDATA1.GetDataTable
        t.TableName = "SOTORDRM"
        dst.Tables.Add(t)
        grdSOTORDRM.DataSource = t
        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTORDRM.DisplayLayout.Bands(0).Columns
            If gcol.Key = "STYLE_CODE" Then
                gcol.Width = 90
                gcol.Header.Caption = "Style"
                Create_Summary(grdSOTORDRM, "STYLE_CODE", "Count")
            ElseIf gcol.Key = "COLOR_CODE" Then
                gcol.Width = 40
                gcol.Header.Caption = "Color"
            ElseIf gcol.Key = "QTY" Then
                gcol.Width = 70
                gcol.Header.Caption = "Total"
                gcol.Format = "#,##0"
                Create_Summary(grdSOTORDRM, "QTY")
            Else
                gcol.Width = 70
                gcol.Header.Caption = Mid(gcol.Key, 5)
                gcol.Format = "#,##0"
                Create_Summary(grdSOTORDRM, gcol.Key)
            End If
        Next

        grdSOTORDRM.Text = "Shipment " & SHIP_BOL_NO & ", Style Summary by Store" & IIf(SHIP_STATUS = "P", " - Qtys in Pick", " - Qtys Shipped")
        Sort_grdColumns(grdSOTORDRM, "STYLE_CODE,COLOR_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub chkShowDetails_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        splSOTSHIPX.Panel2Collapsed = Not (chkShowDetails.Checked)
        If chkShowDetails.Checked Then
            Setup_grdSOTSHIPX()
        End If
    End Sub

    Private Sub tabSOTSHIPX_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSOTSHIPX.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabSOTSHIPX()
    End Sub

    Sub Setup_tabSOTSHIPX()
        UltraExplorerBar1.Groups("Shipment Status").Visible = tabMain.SelectedTab.Key = "Shipments"
        UltraExplorerBar1.Groups("Shipment Status Updates").Visible = tabMain.SelectedTab.Key = "Shipments" And chkShowDetails.Checked And (tabSOTSHIPX.SelectedTab.Key = "Shipment Status Updates")
        UltraExplorerBar1.Groups("Multiple Record Changes").Visible = tabMain.SelectedTab.Key = "Shipments" And chkEnableUpdate.Checked
        UltraExplorerBar1.Groups("Scope (Blank for All)").Visible = tabMain.SelectedTab.Key = "Shipments" And Not ScreenMode And Not chkEnableUpdate.Checked And Not UltraExplorerBar1.Groups("Shipment Status Updates").Visible

        If tabSOTSHIPX.SelectedTab.Key = "Styles" Then
            Setup_Summary_SOTORDRM(SHIP_BOL_NO)
        End If
    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabSOTSHIPX()
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If cmbType.Value & "" = "" Then
            MsgBox("No Event Type Selected", MsgBoxStyle.OkOnly, "Cannot Add Event")
            Exit Sub
        End If
        If txtNote.Text = "" Then
            MsgBox("No Event Description Provided", MsgBoxStyle.OkOnly, "Cannot Add Event")
            Exit Sub
        End If

        Record_Event(cmbType.Value, txtNote.Text)

        cmbType.Value = DBNull.Value
        txtNote.Text = ""
    End Sub

    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String)
        Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
        With rowTATEVNT1
            .Item("TABLE_NAME") = "SOTSHIP1"
            .Item("TABLE_KEY") = SHIP_BOL_NO
            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("EVENT_TYPE") = EVENT_TYPE
            .Item("EVENT_DESC") = EVENT_DESC
            .Item("EVENT_KEY") = ""
            .Item("FORM_NAME") = Me.Name
        End With
        dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
        Update_Record_TDA("TATEVNT1")
    End Sub

    Sub Record_Audit(SHIP_BOL_NO As String, COLUMN_NAME As String, OLD_VALUE As String, NEW_VALUE As String)

        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
        With rowASTAUDT1
            .Item("TABLE_NAME") = "SOTSHIP1"
            .Item("KEY_VALUE") = SHIP_BOL_NO
            .Item("COLUMN_NAME") = COLUMN_NAME
            .Item("USER_ID") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            .Item("OLD_VALUE") = OLD_VALUE
            .Item("NEW_VALUE") = NEW_VALUE
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("SELECTION_NO") = Me.SELECTION_NO
            .Item("XNO") = Me.XNO
        End With
        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
    End Sub

    Private Sub grdSOTSHIP3_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSHIP3.ClickCellButton
        If e.Cell.Row.Cells("LAST_OPER").Value & "" = "" Then
            e.Cell.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
            e.Cell.Row.Cells("LAST_DATE").Value = Now + ASCMAIN1.NowTSD
            e.Cell.Row.Update()
            Update_Record_TDA("SOTSHIP3")
            'dst.Tables ("SOTSHIP3").AcceptChanges 
        End If
    End Sub
End Class