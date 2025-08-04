Imports Infragistics.Win.UltraWinGrid

Public Class WHFAREC1
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "WHFARECI" Then
            InquiryMode = True
        End If

        Get_PARM("SOTPARM1")
        With dst


            ASCMAIN1.sql = "Select ICTWHSE1.WHSE_CODE, ICTWHSE1.WHSE_DESC, shipx.containers, shipx.records from ICTWHSE1 ,
                            (Select POTSHIP1.WHSE_CODE, count(distinct CONTAINER_NO)CONTAINERS, count(1) records from POTSHIP1,POTSHIP2
                            where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO
                            and POTSHIP2.PO_SHIP_STATUS = 'O'
                            group by POTSHIP1.WHSE_CODE) SHIPX
                            where ICTWHSE1.WHSE_CODE  = SHIPX.WHSE_CODE"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False, "", 1)


            ASCMAIN1.sql = "Select POTSHIP2.*, POTSHIP1.PO_SHIP_VESSEL, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
                & " from POTSHIP1,POTSHIP2" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIP_STATUS = 'O' and POTSHIP1.WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "POTSHIP2", "**", 0, False, "V", 2)
            With .Tables("POTSHIP2")
                .Columns.Add("SELECT")
            End With

            ASCMAIN1.sql = "Select POTSHIP3.* " & vbCrLf _
                & ", POTORDR1.VEND_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE " & vbCrLf _
                & ", POTORDR2.PO_QTY_OPN, POTORDR2.PO_QTY_UOM, POTORDR2.PO_COST ORDR2_COST" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_BODY_CODE, POTORDR2.SUB_UNIT_PACK_QTY, POTORDR2.CARTON_PACK_QTY" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
                & ", POTSHIP3.PO_QTY_REC PO_QTY_REC_OLD" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_OPEN, ICTSTAT2.WHSE_QTY_PICK" & vbCrLf _
                & " from POTSHIP3,POTORDR2,ICTSTYL1,POTORDR1,ICTSTAT2,POTSHIP1" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTSHIP1.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTAT2.WHSE_CODE = POTSHIP1.WHSE_CODE" & vbCrLf _
                & "   and ICTSTAT2.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTAT2.COLOR_CODE = POTORDR2.COLOR_CODE" & vbCrLf _
                & "   and (POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO) in (Select Distinct POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO " & vbCrLf _
                & "  from POTSHIP1,POTSHIP2" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIP_STATUS = 'O' and POTSHIP1.WHSE_CODE = :PARM1)"
            Create_TDA(.Tables.Add, "POTSHIP3", "**", 0, True, "V", 4)
            'With .Tables("POTSHIP3")
            '    .Columns.Add("ORDR_QTY_OPEN", GetType(System.Int64))
            '    .Columns.Add("ORDR_QTY_PICK", GetType(System.Int64))
            '    .Columns.Add("WHSE_QTY_ON_HAND", GetType(System.Int64))
            'End With

            Create_Relation("POTSHIP2", "POTSHIP3", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")

            With .Tables("POTSHIP2")
                .Columns.Add("PO_QTY_SHP", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).PO_QTY_SHP)")
                .Columns.Add("WHSE_QTY_OPEN", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).WHSE_QTY_OPEN)")
                .Columns.Add("WHSE_QTY_PICK", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).WHSE_QTY_PICK)")
                .Columns.Add("WHSE_QTY_ON_HAND", GetType(System.Int64), "SUM(CHILD(POTSHIP2_POTSHIP3).WHSE_QTY_ON_HAND)")
            End With




            ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                & " where WHSE_CODE = :PARM1 and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3" & vbCrLf _
                & " and LOCATION_QTY <> 0"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "VVV")



            ASCMAIN1.sql = "SELECT 'O' ORDR_TYPE, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.CUST_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY ORDR" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN OPEN" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_ALLO ALLO" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP SHIP" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_CANC CANC" & vbCrLf _
                & ", 0 ORDERS" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_OPEN" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_CANC" & vbCrLf _
                & ", SOTORDR1.CUST_NAME" & vbCrLf _
                & ", SOTORDR1.ORDR_DATE_RECD" & vbCrLf _
                & " From SOTORDR2, SOTORDR1" & vbCrLf _
                & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "", 2)



            '----------------------------------
            'PICK TICKET PRINT TEST
            'ALSO SEE CMDPRINT

            Dim PICK_NO As String = "0000608761"
            ASCMAIN1.sql = "Select * from SOTPICK1 where PICK_NO = '" & PICK_NO & "'"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "", 1)
            ASCMAIN1.sql = "Select * from SOTPICK2 where PICK_NO = '" & PICK_NO & "'"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "", 2)
            Dim ORDR_NO As String = "0000440040"
            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)
            ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)
            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE from SOTORDR1,SOTORDR2 where ROWNUM < 1"
            Create_TDA(.Tables.Add, "WHTLOCBE", "**", 0, False, "", 2)
            .Tables("WHTLOCBE").Columns.Add("LOCATIONS")

        End With

        Fill_Records("SOTPICK1")
        Fill_Records("SOTPICK2")
        Fill_Records("SOTORDR1")
        Fill_Records("SOTORDR2")

        dst.Tables("WHTLOCBE").Rows.Clear()

        Dim WHSE_CODE As String = "MS" ' THIS WOULD COME FROM SOTORDR1
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR2"), New String() {"STYLE_CODE", "COLOR_CODE"}).Select("")
            'For Each row As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")

            Dim rowWHTLOCBE As DataRow = dst.Tables("WHTLOCBE").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If rowWHTLOCBE Is Nothing Then

                Dim LOCATIONS As String = ""

                ASCMAIN1.sql = "Select * from WHTLOCB1 where WHSE_CODE = '" & WHSE_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and LOCATION_QTY > 0"
                For Each row2 As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim LOCATION_CODE As String = row2.Item("LOCATION_CODE")
                    LOCATIONS &= "," & LOCATION_CODE
                Next

                rowWHTLOCBE = dst.Tables("WHTLOCBE").NewRow
                rowWHTLOCBE.Item("STYLE_CODE") = STYLE_CODE
                rowWHTLOCBE.Item("COLOR_CODE") = COLOR_CODE
                rowWHTLOCBE.Item("LOCATIONS") = Mid(LOCATIONS, 2)
                dst.Tables("WHTLOCBE").Rows.Add(rowWHTLOCBE)
            End If

        Next



        '----------------------------------

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdPOTSHIP2.DataSource = dst.Tables("POTSHIP2")
        grdPOTSHIP3.DataSource = dst.Tables("POTSHIP3")
        grdWHTLOCB1.DataSource = dst.Tables("WHTLOCB1")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")

        With grdPOTSHIP2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "SELECT" Then
                    gcol.CellActivation = Activation.AllowEdit
                Else
                    gcol.CellActivation = Activation.NoEdit
                    'gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                End If
                If New String() {"WHSE_QTY_ON_HAND", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf gcol.Key = "PO_QTY_SHP" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdPOTSHIP3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"WHSE_QTY_ON_HAND", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf gcol.Key = "PO_QTY_SHP" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        Fill_Records("ICTWHSEX")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")

        Create_Summary(grdPOTSHIP2, "PO_SHIPMENT_LNO", "Count")
        Create_Summary(grdPOTSHIP2, New String() {"PO_QTY_SHP", "WHSE_QTY_ON_HAND", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"})

        Create_Summary(grdPOTSHIP3, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTSHIP3, New String() {"PO_QTY_SHP", "WHSE_QTY_ON_HAND", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"})

        Create_Summary(grdSOTORDRX, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDRX, New String() {"ORDERS", "ORDR", "OPEN", "SHIP", "PICK", "CANC", "ALLO", "ORDR_AMT"}) ', "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP"


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                    '  If Not ASCMAIN1.Logical_Open("WHTPACK1", WHSE_CODE) Then Exit Sub
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

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode


                    .Items("Done").Settings.Enabled = iScreenMode



                End With

                .Groups("Packing Filters").Visible = False ' JUST KEEPING THIS GROUP AROUND IN CASE WE NEED SOMETHING LIKE THIS ' ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        splShipments.Visible = ScreenMode


        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        chkAutoRefresh.Checked = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ICTWHSEX", "POTSHIP2", "POTSHIP3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
        '     Setup_tab0()


    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        'EnforceConstraints(False)
        Fill_Records("POTSHIP2", WHSE_CODE)
        Fill_Records("POTSHIP3", WHSE_CODE)
        'EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
     
    Sub Update_Record()

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Preparing to Update")

        'Update_Record_TDA("SOTCART2")

        'CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        'Call BeginTrans()
        'Stop
        ''Call Delete_Records("table")
        'Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWHSEX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdPOTSHIP2, "SBB", "Show Filter", "PO Shipment Inquiry", "Whse Carrier")
        Load_Popup_Menu(grdPOTSHIP3, "SBB", "Show Filter", "Style Status Inquiry", "PO Inquiry")
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
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTPACKX"
                Case "grdSOTORDRX"
                    Dim ORDR_TYPE As String = ""
                    If grdSOTORDRX.ActiveRow IsNot Nothing Then
                        ORDR_TYPE = grdSOTORDRX.ActiveRow.Cells("ORDR_TYPE").Value
                    End If
                    tlb_btn = DirectCast(tlb_pop.Tools("Sales Order Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O")

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
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

                'Case "Style Master"
                '    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                '    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                '    If rowICTSTYL1 IsNot Nothing Then
                '        Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                '    End If


            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "PO Shipment Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = ""
                Dim ORDR_GROUP_NO As String = ""

                ORDR_GROUP_NO = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                ASCMAIN1.sql = "Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                ORDR_NO = ASCDATA1.GetDataValue

                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")

            Case "Whse Carrier"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Dim PO_SHIPMENT_LNO As String = grd.ActiveRow.Cells("PO_SHIPMENT_LNO").Text
                Dim WH_CARRIER As String = ""

                For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = '" & PO_SHIPMENT_LNO & "'", "")
                    WH_CARRIER = rowPOTSHIP2.Item("WH_CARRIER") & ""
                Next

                Dim frmASFMSGBF As New ASFMSGBF
                Dim label As New System.Text.StringBuilder With {.Length = 0}
                label.AppendLine("Enter Carrier Id (10 chars max) ")
                Dim Caption As String = "Warehouse Carrier Id"
                WH_CARRIER = frmASFMSGBF.Get_txtblock_from_User(label.ToString, Caption, WH_CARRIER, False, 10)

                ASCMAIN1.sql = $"Update POTSHIP2 set WH_CARRIER = '{WH_CARRIER}'
                                    Where PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'
                                    and   PO_SHIPMENT_LNO = '{PO_SHIPMENT_LNO}'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                For Each rowPOTSHIP2 As DataRow In dst.Tables("POTSHIP2").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = '" & PO_SHIPMENT_LNO & "'", "")
                    rowPOTSHIP2.Item("WH_CARRIER") = WH_CARRIER
                Next
                grdPOTSHIP2.ActiveRow.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)


        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("View")
        End Select
    End Sub
#End Region

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("View")
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()

    End Sub

    Private Sub grdPOTSHIP2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTSHIP2.AfterRowActivate
        Setup_POTSHIP3()
    End Sub
    Sub Setup_POTSHIP3()
        If grdPOTSHIP2.ActiveRow Is Nothing OrElse grdPOTSHIP2.ActiveRow.IsFilterRow OrElse Not grdPOTSHIP2.ActiveRow.IsDataRow Then
            grdPOTSHIP3.Visible = False
        Else
            grdPOTSHIP3.Visible = True
            Dim PO_SHIPMENT_NO As String = grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_NO").Value
            Dim PO_SHIPMENT_LNO As Int32 = Val(grdPOTSHIP2.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")

            Dim DVW As DataView = DirectCast(grdPOTSHIP3.DataSource, DataTable).DefaultView

            DVW.RowFilter = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)

            grdPOTSHIP3.Text = "Commercial Invoice " & grdPOTSHIP2.ActiveRow.Cells("COMM_INV_NO").Value _
                & " / Bill of Lading " & grdPOTSHIP2.ActiveRow.Cells("BOL_NO").Value _
                & " / Container " & grdPOTSHIP2.ActiveRow.Cells("CONTAINER_NO").Value _
                & " - Details"
        End If

    End Sub 

    Private Sub grdPOTSHIP3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTSHIP3.AfterRowActivate
        Setup_WHTLOCB1()
        Setup_SOTORDRX()
    End Sub

    Sub Setup_WHTLOCB1()
        If grdPOTSHIP3.ActiveRow Is Nothing OrElse grdPOTSHIP3.ActiveRow.IsFilterRow OrElse Not grdPOTSHIP3.ActiveRow.IsDataRow Then
            grdWHTLOCB1.Visible = False
        Else
            grdWHTLOCB1.Visible = True
            Dim STYLE_CODE As String = grdPOTSHIP3.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdPOTSHIP3.ActiveRow.Cells("COLOR_CODE").Value
 
            grdWHTLOCB1.Text = "Locations for Style-Color " & STYLE_CODE & "-" & COLOR_CODE
            Fill_Records("WHTLOCB1", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
        End If
    End Sub
    Sub Setup_SOTORDRX()

        If grdPOTSHIP3.ActiveRow Is Nothing OrElse grdPOTSHIP3.ActiveRow.IsFilterRow OrElse Not grdPOTSHIP3.ActiveRow.IsDataRow Then
            grdSOTORDRX.Visible = False
        Else
            grdSOTORDRX.Visible = True
            Dim STYLE_CODE As String = grdPOTSHIP3.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdPOTSHIP3.ActiveRow.Cells("COLOR_CODE").Value

       
            grdSOTORDRX.Text = "Sales Order Details for " & STYLE_CODE & "-" & COLOR_CODE


            ASCMAIN1.sql = "SELECT 'O' ORDR_TYPE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
               & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
               & ", MIN(SOTORDR1.SREP_CODE) SREP_CODE, MIN(SOTORDR1.WHSE_CODE) WHSE_CODE, SOTORDR0.ORDR_TYPE_CODE" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY) ORDR, SUM (SOTORDR2.ORDR_QTY_OPEN) OPEN" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY_PICK) PICK, SUM (SOTORDR2.ORDR_QTY_ALLO) ALLO" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY_SHIP) SHIP, SUM (SOTORDR2.ORDR_QTY_CANC) CANC" & vbCrLf _
               & ", COUNT (DISTINCT SOTORDR1.ORDR_NO) ORDERS" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY      * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
               & ", SUM (SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
               & ", ARTCUST1.CUST_NAME" & vbCrLf _
               & ", SOTORDR1.ORDR_DATE_RECD" & vbCrLf _
               & " From SOTORDR2, SOTORDR1, SOTORDR0, ARTCUST1" & vbCrLf

            ASCMAIN1.sql &= "" _
                & " where SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf


            ASCMAIN1.sql &= "   and (SOTORDR2.ORDR_STATUS = 'O' OR SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf

            ASCMAIN1.sql &= "" _
                & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE"

            ASCMAIN1.sql &= "" _
                & " group by SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, ARTCUST1.CUST_NAME, SOTORDR0.ORDR_TYPE_CODE, SOTORDR1.ORDR_DATE_RECD" & vbCrLf


            Fill_Records("SOTORDRX", "", True, ASCMAIN1.sql)
        End If
    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click
        print_record()
    End Sub


    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Pick Ticket")

        Dim REPORT_NAME As String = "SORPICKE"

        Print_Report_Begin()
        Generate_Report(REPORT_NAME, "eCommerce Pick Ticket", "")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdPOTSHIP2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTSHIP2.InitializeRow
        If e.Row.Cells("WH_CARRIER").Value & "" <> "" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightBlue
            'e.Row.Cells("ORDR_GROUP_NO").ToolTipText = "Some or All Orders are In Pick"
            e.Row.ToolTipText = "Carrier: " & e.Row.Cells("WH_CARRIER").Value
        Else
            e.Row.Appearance.BackColor = Nothing
            e.Row.ResetToolTipText()
        End If

    End Sub
End Class