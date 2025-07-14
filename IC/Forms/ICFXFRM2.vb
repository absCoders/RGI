Public Class ICFXFRM2
    Dim ICTSTYLX As String
    Dim sqlICTSTYLX As String

    Dim SOTORDRX As String
    Dim sqlSOTORDRX As String

    Dim SOTINVHX As String
    Dim sqlSOTINVHX As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_WorkTables(True)

        With dst

            ASCMAIN1.sql = $"Select * from {SOTORDRX} where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add("SOTORDRX"), SOTORDRX, "**", 0, False, "VV")

            ASCMAIN1.sql = "Select * from " & ICTSTYLX
            Create_TDA(.Tables.Add("ICTSTYLX"), ICTSTYLX, "**", 0, False)
            .Tables("ICTSTYLX").Columns.Add("QTY2XFR", GetType(System.Int32))
            .Tables("ICTSTYLX").Columns.Add("AVA_MS", GetType(System.Int32), "ISNULL(ONHD_MS,0)-ISNULL(PICK_MS,0)-ISNULL(OPEN_MS,0)+ISNULL(ONPO_MS,0)+ISNULL(TRAN_MS,0)")
            .Tables("ICTSTYLX").Columns.Add("AVA_US", GetType(System.Int32), "ISNULL(ONHD_US,0)-ISNULL(PICK_US,0)-ISNULL(OPEN_US,0)")
            .Tables("ICTSTYLX").Columns.Add("QTY_SHORT", GetType(System.Int32), "IIF(ORDR_QTY_OPEN > 0 AND AVA_MS - ORDR_QTY_OPEN <= 0, AVA_MS - ORDR_QTY_OPEN, NULL)")

            ASCMAIN1.sql = $"Select * from {SOTINVHX} where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add("SOTINVHX"), SOTINVHX, "**", 0, False, "VV")

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDRG", "*")

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SOTORDR1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V")

        End With


        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        Create_Summary(grdICTSTYLX, "STYLE_CODE", "Count")
        Create_Summary(grdICTSTYLX, "QTY2XFR")
        Create_Summary(grdICTSTYLX, "ORDR_QTY_OPEN")
        Create_Summary(grdICTSTYLX, "QTY_SHORT")

        Create_Summary(grdSOTORDRX, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRX, "ORDR_QTY_OPEN")

        Create_Summary(grdSOTINVHX, "ECOM_CODE", "Count")
        Create_Summary(grdSOTINVHX, "QTY_SHIPPED_TOT")
        Create_Summary(grdSOTINVHX, "AVG_WEEK_SALES")

        For i As Integer = 1 To 18
            Dim colName As String = $"MTH_{i.ToString("D2")}_QTY"
            Create_Summary(grdSOTINVHX, colName)
        Next


        With grdICTSTYLX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.True

            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_STATUS").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("COLOR_DESC").Header.Fixed = True
            .Columns("STYLE_COLOR_STATUS").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "QTY2XFR" Then
                    gcol.CellAppearance.BackColor = Color.PaleGreen
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                If New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "COLOR_DESC", "STYLE_STATUS", "STYLE_COLOR_STATUS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf gcol.Key.EndsWith("_MS") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf gcol.Key.EndsWith("_US") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"CARTON_PACK_QTY", "INNER_PACK_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY_OPEN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleTurquoise
                ElseIf New String() {"QTY2XFR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                ElseIf New String() {"QTY_SHORT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                End If
            Next
        End With

        'grdSOTORDR0.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDRX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"ORDR_DATE_RECD", "ORDR_PRIORITY",
                                 "ORDR_RELEASE_AVAIL_MIN", "ORDR_RELEASE_AVAIL_MAX", "ORDR_REL_SHORT", "ORDR_REL_SHORT_OPER",
                                 "ORDR_REL_ACTION_DATE", "ORDR_REL_ACTION_OPER", "TERM_CODE", "LAST_DATE", "LAST_OPER", "ORDR_SHIP_INSTR", "ORDR_MESSAGE", "EDI_PO_TYPE"}.Contains(gcol.Key) Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If

                If New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf gcol.Key = "ORDR_QTY_ALLO_3PL" Or gcol.Key = "ORDR_AMT_ALLO_3PL" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LimeGreen
                ElseIf gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleTurquoise
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT", "WHSE_CODE", "EDI_MERCH_TYPE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "SALES_DIVISION_CODE", "SREP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"CUST_CITY", "CUST_STATE", "CUST_COUNTRY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                Else
                    gcol.Header.Appearance.BackColor = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next

        End With

        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("ECOM_NAME").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdICTSTYLX, "STYLE_STATUS")
        ASCMAIN1.Add_Value_List(grdICTSTYLX, "STYLE_COLOR_STATUS")

        'spl.Panel1Collapsed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            'need to change ava_ms back to ava_us after testing
            Case "Update"
                If dst.Tables("ICTSTYLX").Select("QTY2XFR IS NOT NULL AND QTY2XFR > 0").Length = 0 Then
                    EMsg &= "No transfer quantities were entered. Nothing to update."
                End If


                For Each row As DataRow In dst.Tables("ICTSTYLX").Select("QTY2XFR IS NOT NULL AND QTY2XFR > 0")
                    Dim QTY2XFR As Integer = row.Field(Of Integer)("QTY2XFR")
                    Dim AVA_US As Integer = row.Field(Of Integer)("AVA_US")

                    If QTY2XFR > AVA_US Then
                        EMsg &= $"Transfer qty for {row("STYLE_CODE")}-{row("COLOR_CODE")} exceeds available US qty ({QTY2XFR} > {AVA_US}).{vbCrLf}"
                        Continue For
                    End If

                    Dim valid As Boolean = True
                    Dim INNER_PACK_QTY = row("INNER_PACK_QTY")
                    Dim CARTON_PACK_QTY = row("CARTON_PACK_QTY")

                    If Not IsDBNull(INNER_PACK_QTY) AndAlso CInt(INNER_PACK_QTY) > 0 Then
                        If QTY2XFR Mod CInt(INNER_PACK_QTY) <> 0 Then
                            valid = False
                            EMsg &= $"Transfer qty for {row("STYLE_CODE")}-{row("COLOR_CODE")} must be a multiple of INNER_PACK_QTY ({INNER_PACK_QTY}).{vbCrLf}"
                        End If
                    ElseIf Not IsDBNull(CARTON_PACK_QTY) AndAlso CInt(CARTON_PACK_QTY) > 0 Then
                        If QTY2XFR Mod CInt(CARTON_PACK_QTY) <> 0 Then
                            valid = False
                            EMsg &= $"Transfer qty for {row("STYLE_CODE")}-{row("COLOR_CODE")} must be a multiple of CARTON_PACK_QTY ({CARTON_PACK_QTY}).{vbCrLf}"
                        End If
                    End If
                Next

                Stop ' check that all non-0 qtys are positive, a multiple of case or inner, and also less than qty avail
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Auto Populate #ToXfr"
                Auto_Populate()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Auto Populate #ToXfr").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Edit").Visible = (ScreenMode And EntryMode = "L")
                    .Items("Done").Visible = (ScreenMode And EntryMode = "L")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = True

        If ScreenMode Then

            With grdICTSTYLX.DisplayLayout.Bands(0).Columns("QTY2XFR")
                If EntryMode = "E" Then
                    .CellAppearance.BackColor = Color.LightGreen
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .CellAppearance.BackColor = Color.Empty
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End With

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTSTYLX", "SOTORDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        SplitContainer1.Visible = False
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        Create_WorkTables()

        Fill_Records("SOTORDRX")
        Fill_Records("ICTSTYLX")

        Sort_grdColumns(grdICTSTYLX, "STYLE_CODE,COLOR_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
        Dim ORDR_GROUP_NO As String = ORDR_NO

        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow
        With rowSOTORDR1
            .Item("ORDR_NO") = ORDR_NO
            .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            .Item("ORDR_TYPE_CODE") = "XFR"
            .Item("ORDR_STATUS") = "O"
            .Item("ORDR_SOURCE") = "K"
            .Item("ORDR_DATE") = Now.Date
            .Item("ORDR_DATE_BOOKED") = Now.Date
            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
            .Item("ORDR_SHIP_DATE") = Now.Date
            .Item("ORDR_CANCEL_DATE") = Now.Date.AddDays(10)
            .Item("INIT_DATE") = Now
            .Item("LAST_DATE") = Now
            .Item("INIT_OPER") = ASCMAIN1.USER_ID


            .Item("WHSE_CODE") = "NC"
            .Item("WHSE_CODE_TO") = "MS"

            'will need to change this
            .Item("CUST_CODE") = "030758"
            .Item("CUST_NAME") = "COUNTRYSIDE GARDEN CENTER"
            .Item("CUST_STORE_NO") = "000000"
            .Item("CUST_STORE_NAME") = "COUNTRYSIDE GARDEN CENTER"
            .Item("ORDR_FOB") = "Altamahaw,NC"

            .Item("TERM_CODE") = "N30"
            .Item("FRT_TERMS") = "COL"

            .Item("CURR_CODE") = "USD"
            .Item("CURR_EXCH_RATE") = 1D

            .Item("SHIP_VIA_CODE") = "BST"
            .Item("POST_CODE") = "REG"
            .Item("SREP_CODE") = "TN"
            .Item("SALES_DIVISION_CODE") = "RIB"
            .Item("ORDR_ADDR_TYPE_ST") = "MK"
            .Item("ORDR_REL_HOLD_CODES") = "0"
            .Item("CUST_BILL_TO_CUST") = "030758"
        End With
        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

        Dim nextLno As Integer = 0
        For Each r As DataRow In dst.Tables("ICTSTYLX") _
                               .Select("QTY2XFR IS NOT NULL AND QTY2XFR > 0")
            nextLno += 1
            Dim d As DataRow = dst.Tables("SOTORDR2").NewRow
            With d
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = nextLno
                .Item("STYLE_CODE") = r("STYLE_CODE")
                .Item("STYLE_DESC") = r("STYLE_DESC")
                .Item("STYLE_UOM") = r("STYLE_UOM")
                .Item("COLOR_CODE") = r("COLOR_CODE")
                .Item("ORDR_QTY") = r("QTY2XFR")
                .Item("ORDR_QTY_OPEN") = r("QTY2XFR")
                .Item("ORDR_QTY_ORIG") = r("QTY2XFR")
                .Item("ORDR_STATUS") = "O"
                .Item("INNER_PACK_QTY") = r("INNER_PACK_QTY")
                .Item("CARTON_PACK_QTY") = r("CARTON_PACK_QTY")
                .Item("ORDR_UNIT_PRICE") = 0

                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", .Item("STYLE_CODE"))
                If rowICTSTYL1 IsNot Nothing Then
                    .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                    .Item("STYLE_PRICE") = rowICTSTYL1.Item("STYLE_PRICE")
                    .Item("ORDR_UNIT_PRICE_STD") = rowICTSTYL1.Item("STYLE_PRICE")
                    '.Item("ORDR_UNIT_PRICE_CALC") = 
                    .Item("ORDR_UNIT_PRICE_MANUAL") = 0D
                    '.Item("ORDR_PRICE_SOURCE") = 
                    '.Item("COMM_RATE") = 
                End If
            End With
            dst.Tables("SOTORDR2").Rows.Add(d)
        Next

        Dim SOTORDR5 As DataTable = dst.Tables("SOTORDR5")

        Dim rowBT As DataRow = SOTORDR5.NewRow
        Dim billToCust As DataRow = LookUp("ARTCUST1", "030758")
        With rowBT
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_ADDR_TYPE") = "BT"
            .Item("CUST_ADDR_CODE") = "030758"
            .Item("CUST_NAME") = billToCust.Item("CUST_NAME")
            .Item("CUST_ADDR1") = billToCust.Item("CUST_ADDR1")
            .Item("CUST_ADDR2") = billToCust.Item("CUST_ADDR2")
            .Item("CUST_CITY") = billToCust.Item("CUST_CITY")
            .Item("CUST_STATE") = billToCust.Item("CUST_STATE")
            .Item("CUST_ZIP_CODE") = billToCust.Item("CUST_ZIP_CODE")
            .Item("CUST_COUNTRY") = billToCust.Item("CUST_COUNTRY")
            .Item("CUST_CONTACT") = billToCust.Item("CUST_CONTACT")
            .Item("CUST_PHONE") = billToCust.Item("CUST_PHONE")
            .Item("CUST_FAX") = billToCust.Item("CUST_FAX")
            .Item("CUST_EMAIL") = billToCust.Item("CUST_EMAIL")
        End With
        SOTORDR5.Rows.Add(rowBT)

        Dim rowST As DataRow = SOTORDR5.NewRow
        With rowST
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_ADDR_TYPE") = "ST"
            .Item("CUST_ADDR_CODE") = "MS"
            .Item("CUST_NAME") = "MOUNTAINSIDE WAREHOUSE"
            .Item("CUST_ADDR1") = "1112 Bristol Road"
            .Item("CUST_CITY") = "Mountainside"
            .Item("CUST_STATE") = "NJ"
            .Item("CUST_ZIP_CODE") = "07092"
            .Item("CUST_COUNTRY") = "USA"
            .Item("CUST_PHONE") = "9086541515"
            .Item("CUST_FAX") = "9086543582"
        End With
        SOTORDR5.Rows.Add(rowST)

        Record_Event("UPDT", "Transfer Order Updated", ORDR_NO)
        Dim SQLD As String = $"ORDR_NO = '{ORDR_NO}'"
        INIT_LAST("SOTORDR1", False,, True)
        Update_Record_TDA("SOTORDR1", SQLD)
        Update_Record_TDA("SOTORDR2", SQLD)
        Update_Record_TDA("SOTORDR5", SQLD)
        Update_Record_TDA("TATEVNT1")

        Dependent_Updates(1, ORDR_NO)
        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

        'If ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI" Then
        '    ASCDATA1.ExecuteSP("SOPORDR1_COMM", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        'End If

        Dim rowSOTORDRG As DataRow = Fill_Record("SOTORDRG", ORDR_GROUP_NO)

        If rowSOTORDRG Is Nothing Then
            rowSOTORDRG = dst.Tables("SOTORDRG").NewRow
            rowSOTORDRG.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            dst.Tables("SOTORDRG").Rows.Add(rowSOTORDRG)
        End If

        If rowSOTORDRG.Item("ORDR_REL_SHORT").ToString() <> "1" Then
            rowSOTORDRG.Item("ORDR_REL_SHORT") = "1"
            rowSOTORDRG.Item("ORDR_REL_SHORT_OPER") = ASCMAIN1.USER_ID
            rowSOTORDRG.Item("ORDR_REL_SHORT_DATE") = DATETIME_STAMP
        End If

        Update_Record_TDA("SOTORDRG")

        Stop
        ' create Sales Order of type XFR - SOTORDR1/2/5 SOTORDR0
        ' release the order - SOTPICK1/2 SOTPICK0, SOTSHIP1, SOTCART1/2

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYLX, "SSB", "Show Filter", "Show GroupBox", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTORDRX, "SSBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Customer Order Inquiry")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        Select Case e.SourceControl.Name
            Case "grdICTSTYLX"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdICTSTYLX"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = Lookup("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Dim rowARTCUST1 As DataRow = Lookup("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

#End Region

    Sub Create_WorkTables(Optional initialize As Boolean = False)

        If initialize Then

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME" & vbCrLf _
                & ", SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.SREP_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_SOURCE, SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
                & ", SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_DATE_RECD" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR1.ORDR_PRIORITY" & vbCrLf _
                & " from SOTORDR1,SOTORDR2,ARTCUST1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_QTY_OPEN > 0" & vbCrLf _
                & "   and SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
                & "   and SOTORDR1.WHSE_CODE = 'MS'" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "" ' "   and (ARTCUST1.CUST_COUNTRY <> 'USA' or ARTCUST1.CUST_COUNTRY <> 'US')"

            sqlSOTORDRX = ASCMAIN1.sql
            ASCMAIN1.sql = $"{sqlSOTORDRX} and ROWNUM < 1"
            SOTORDRX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {SOTORDRX} Add Primary Key (ORDR_NO, ORDR_LNO)")

            ASCMAIN1.sql = "Select C1.STYLE_CODE, C1.COLOR_CODE, X.ORDR_QTY_OPEN" & vbCrLf _
            & ", L1.STYLE_DESC, L1.CARTON_PACK_QTY, L1.INNER_PACK_QTY" & vbCrLf _
            & ", L1.STYLE_STATUS, C1.STYLE_COLOR_STATUS" & vbCrLf _
            & ", R1.COLOR_DESC" & vbCrLf _
            & ", L1.STYLE_UOM, L1.VEND_CODE" & vbCrLf _
            & ", XMS.ONHD_MS, XMS.PICK_MS, XMS.OPEN_MS, XMS.ONPO_MS, XMS.TRAN_MS" & vbCrLf _
            & ", XUS.ONHD_US, XUS.PICK_US, XUS.OPEN_US, XUS.ONPO_US, XUS.TRAN_US" & vbCrLf _
            & ", NVL(SALES.AVG_WEEK_SALES, 0) AVG_WEEK_SALES" & vbCrLf _
            & ", CASE WHEN NVL(SALES.AVG_WEEK_SALES,0) = 0 THEN NULL ELSE ROUND((NVL(XMS.ONHD_MS,0) - NVL(XMS.PICK_MS,0) - NVL(XMS.OPEN_MS,0) + NVL(XMS.ONPO_MS,0) + NVL(XMS.TRAN_MS,0)) / SALES.AVG_WEEK_SALES, 1) END WOS_MS" & vbCrLf _
            & ", CASE WHEN NVL(SALES.AVG_WEEK_SALES,0) = 0 THEN NULL ELSE ROUND((NVL(XUS.ONHD_US,0) - NVL(XUS.PICK_US,0) - NVL(XUS.OPEN_US,0) + NVL(XUS.ONPO_US,0) + NVL(XUS.TRAN_US,0)) / SALES.AVG_WEEK_SALES, 1) END WOS_US" & vbCrLf _
            & " from ICTSTYL1 L1, ICTCOLR1 R1, ICTSTYC1 C1" & vbCrLf _
            & ", (Select STYLE_CODE, COLOR_CODE, SUM(CASE WHEN ORDR_TYPE_CODE = 'B2C' THEN ORDR_QTY_OPEN ELSE 0 END) AS ORDR_QTY_OPEN" & vbCrLf _
            & $"   from {SOTORDRX} " & vbCrLf _
            & "   group by STYLE_CODE, COLOR_CODE) X" & vbCrLf _
            & ", (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & "   , Sum (WHSE_QTY_ON_HAND) ONHD_MS" & vbCrLf _
            & "   , Sum (WHSE_QTY_PICK) PICK_MS" & vbCrLf _
            & "   , Sum (WHSE_QTY_OPEN) OPEN_MS" & vbCrLf _
            & "   , Sum (WHSE_QTY_ON_ORDER) ONPO_MS" & vbCrLf _
            & "   , Sum (WHSE_QTY_TRAN) TRAN_MS" & vbCrLf _
            & "   from ICTSTAT2 where WHSE_CODE = 'MS'" & vbCrLf _
            & "   group by STYLE_CODE, COLOR_CODE) XMS" & vbCrLf _
            & ", (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & "   , Sum (WHSE_QTY_ON_HAND) ONHD_US" & vbCrLf _
            & "   , Sum (WHSE_QTY_PICK) PICK_US" & vbCrLf _
            & "   , Sum (WHSE_QTY_OPEN) OPEN_US" & vbCrLf _
            & "   , Sum (WHSE_QTY_ON_ORDER) ONPO_US" & vbCrLf _
            & "   , Sum (WHSE_QTY_TRAN) TRAN_US" & vbCrLf _
            & "   from ICTSTAT2 where WHSE_CODE = 'US'" & vbCrLf _
            & "   group by STYLE_CODE, COLOR_CODE) XUS" & vbCrLf _
            & ", (Select h2.STYLE_CODE, h2.COLOR_CODE" & vbCrLf _
            & "   , ROUND(SUM(h2.ORDR_QTY_SHIP) / 78, 2) AS AVG_WEEK_SALES" & vbCrLf _
            & "   from SOTINVH1 h1, SOTINVH2 h2" & vbCrLf _
            & "   where h1.INV_TYPE = h2.INV_TYPE" & vbCrLf _
            & "     and h1.INV_NO   = h2.INV_NO" & vbCrLf _
            & " AND h1.INV_DATE >= ADD_MONTHS(TRUNC(SYSDATE, 'MM'), -17)" & vbCrLf _
            & "     and h2.ORDR_QTY_SHIP > 0" & vbCrLf _
            & "   group by h2.STYLE_CODE, h2.COLOR_CODE) SALES" & vbCrLf _
            & " where L1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
            & "   and R1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
            & "   and C1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
            & "   and C1.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
            & "   and XMS.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
            & "   and XMS.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
            & "   and XUS.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
            & "   and XUS.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
            & "   and SALES.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
            & "   and SALES.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
            & "   and X.STYLE_CODE in (Select Distinct STYLE_CODE from ECTESTY1)"


            sqlICTSTYLX = ASCMAIN1.sql
            ASCMAIN1.sql = $"{sqlICTSTYLX} and ROWNUM < 1"
            ICTSTYLX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {ICTSTYLX} Add Primary Key (STYLE_CODE, COLOR_CODE)")

            'ASCMAIN1.sql =
            '"SELECT e.ECOM_CODE, e.ECOM_NAME" & vbCrLf &
            '", h2.STYLE_CODE, h2.COLOR_CODE" & vbCrLf &
            '", SUM(h2.ORDR_QTY_SHIP)             QTY_SHIPPED_TOT" & vbCrLf &
            '", SUM(h2.ORDR_QTY_SHIP/4)  AVG_WEEK_SALES" & vbCrLf &
            '" FROM SOTINVH1 h1" & vbCrLf &
            '", SOTINVH2 h2" & vbCrLf &
            '", ECTECOM1  e" & vbCrLf &
            '"WHERE h1.INV_TYPE = h2.INV_TYPE" & vbCrLf &
            '" AND h1.INV_NO   = h2.INV_NO" & vbCrLf &
            '" AND h1.CUST_CODE = e.CUST_CODE" & vbCrLf &
            '" AND h1.INV_DATE  >= TRUNC(SYSDATE)-28" & vbCrLf &
            '" AND h2.ORDR_QTY_SHIP > 0" & vbCrLf &
            '"GROUP BY e.ECOM_CODE, e.ECOM_NAME, h2.STYLE_CODE, h2.COLOR_CODE"

            'sqlSOTINVHX = ASCMAIN1.sql
            'ASCMAIN1.sql = "SELECT * FROM (" & sqlSOTINVHX & ") WHERE ROWNUM < 1"
            'SOTINVHX = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL($"ALTER TABLE {SOTINVHX} ADD PRIMARY KEY (ECOM_CODE, STYLE_CODE, COLOR_CODE)")

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("SELECT e.ECOM_CODE, e.ECOM_NAME")
            sb.AppendLine(", h2.STYLE_CODE, h2.COLOR_CODE")
            sb.AppendLine(", SUM(h2.ORDR_QTY_SHIP) QTY_SHIPPED_TOT")
            sb.AppendLine(", ROUND(SUM(h2.ORDR_QTY_SHIP) / 78, 2) AS AVG_WEEK_SALES")

            For i As Integer = 0 To 17
                Dim colAlias As String = $"MTH_{(i + 1).ToString("D2")}_QTY"
                Dim monthExpr As String = $"TO_CHAR(ADD_MONTHS(TRUNC(SYSDATE, 'MM'), -{i}), 'YYYYMM')"
                sb.AppendLine($", ROUND(SUM(CASE WHEN TO_CHAR(h1.INV_DATE, 'YYYYMM') = {monthExpr} THEN h2.ORDR_QTY_SHIP ELSE 0 END), 0) AS {colAlias}")
            Next

            sb.AppendLine("FROM SOTINVH1 h1, SOTINVH2 h2, ECTECOM1 e")
            sb.AppendLine("WHERE h1.INV_TYPE = h2.INV_TYPE")
            sb.AppendLine("AND h1.INV_NO   = h2.INV_NO")
            sb.AppendLine("AND h1.CUST_CODE = e.CUST_CODE")
            sb.AppendLine("AND h1.INV_DATE >= ADD_MONTHS(TRUNC(SYSDATE, 'MM'), -17)")
            sb.AppendLine("AND h2.ORDR_QTY_SHIP > 0")
            sb.AppendLine("GROUP BY e.ECOM_CODE, e.ECOM_NAME, h2.STYLE_CODE, h2.COLOR_CODE")

            ASCMAIN1.sql = sb.ToString()

            sqlSOTINVHX = ASCMAIN1.sql
            ASCMAIN1.sql = "SELECT * FROM (" & sqlSOTINVHX & ") WHERE ROWNUM < 1"
            SOTINVHX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"ALTER TABLE {SOTINVHX} ADD PRIMARY KEY (ECOM_CODE, STYLE_CODE, COLOR_CODE)")

        Else

            ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDRX}")
            ASCDATA1.ExecuteSQL($"Insert into {SOTORDRX} {sqlSOTORDRX}")

            ASCDATA1.ExecuteSQL($"Truncate Table {ICTSTYLX}")
            ASCDATA1.ExecuteSQL($"Insert into {ICTSTYLX} {sqlICTSTYLX}")

            ASCDATA1.ExecuteSQL($"TRUNCATE TABLE {SOTINVHX}")
            ASCDATA1.ExecuteSQL($"INSERT INTO {SOTINVHX} {sqlSOTINVHX}")

        End If
    End Sub

    Private Sub grdICTSTYLX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTSTYLX.AfterRowActivate
        If grdICTSTYLX.ActiveRow Is Nothing OrElse Not grdICTSTYLX.ActiveRow.IsDataRow Then
            SplitContainer1.Panel2Collapsed = True
        Else
            Setup_SOTORDRX()
            Setup_SOTINVHX()
            SplitContainer1.Panel2Collapsed = False
        End If
    End Sub
    Sub Setup_SOTORDRX()
        If grdICTSTYLX.ActiveRow Is Nothing OrElse Not grdICTSTYLX.ActiveRow.IsDataRow Then
            grdSOTORDRX.Visible = False
        Else
            Dim STYLE_CODE As String = grdICTSTYLX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTYLX.ActiveRow.Cells("COLOR_CODE").Value
            Fill_Records("SOTORDRX", New String() {STYLE_CODE, COLOR_CODE})

            grdSOTORDRX.Visible = True
            grdSOTORDRX.Text = $"Open Orders for SC {STYLE_CODE}-{COLOR_CODE}"

        End If
    End Sub
    Sub Setup_SOTINVHX()
        If grdICTSTYLX.ActiveRow Is Nothing Then Exit Sub

        Dim STYLE_CODE As String = grdICTSTYLX.ActiveRow.Cells("STYLE_CODE").Text
        Dim COLOR_CODE As String = grdICTSTYLX.ActiveRow.Cells("COLOR_CODE").Text
        Fill_Records("SOTINVHX", New String() {STYLE_CODE, COLOR_CODE})

        grdSOTINVHX.Text = $"eComm Partner Sales (Last 18 Months) – {STYLE_CODE}-{COLOR_CODE}"

        If grdSOTINVHX.DisplayLayout.Bands.Count > 0 Then
            Dim band = grdSOTINVHX.DisplayLayout.Bands(0)
            Dim baseDate As Date = DateSerial(Year(Today), Month(Today), 0)

            For i As Integer = 0 To 17
                Dim colKey As String = $"MTH_{(i + 1).ToString("D2")}_QTY"
                If band.Columns.Exists(colKey) Then
                    Dim labelDate As Date = DateAdd(DateInterval.Month, -i, baseDate)
                    band.Columns(colKey).Header.Caption = labelDate.ToString("MMM yy")
                    band.Columns(colKey).Format = "###0"
                    band.Columns(colKey).Width = 60
                End If
            Next
        End If
    End Sub
    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String, ORDR_NO As String)
        Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
        rowTATEVNT1.Item("TABLE_NAME") = "SOTORDR1"
        rowTATEVNT1.Item("TABLE_KEY") = ORDR_NO
        rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
        rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowTATEVNT1.Item("EVENT_TYPE") = EVENT_TYPE
        rowTATEVNT1.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
    End Sub
    Sub Dependent_Updates(S As Integer, ORDR_NO As String)

        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE")
                TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", S * QTY_TO_COMMIT)
            End If
        Next

    End Sub
    Sub Auto_Populate()
        Dim ANY_POPULATED As Boolean = False
        For Each row As DataRow In dst.Tables("ICTSTYLX").Select("QTY_SHORT < 0")
            Dim QTY_SHORT As Integer = Math.Abs(Val(row("QTY_SHORT") & ""))
            Dim AVA_US As Integer = Val(row("AVA_US") & "")
            Dim INNER_PACK_QTY As Integer = Val(row("INNER_PACK_QTY") & "")
            Dim CARTON_PACK_QTY As Integer = Val(row("CARTON_PACK_QTY") & "")

            If QTY_SHORT > AVA_US Then
                row("QTY2XFR") = DBNull.Value
                Continue For
            End If

            Dim QTY2XFR As Integer = QTY_SHORT

            If INNER_PACK_QTY > 0 Then
                QTY2XFR = Math.Ceiling(QTY_SHORT / INNER_PACK_QTY) * INNER_PACK_QTY
            ElseIf CARTON_PACK_QTY > 0 Then
                QTY2XFR = Math.Ceiling(QTY_SHORT / CARTON_PACK_QTY) * CARTON_PACK_QTY
            End If

            If QTY2XFR <= AVA_US Then
                row("QTY2XFR") = QTY2XFR
                ANY_POPULATED = True
            Else
                row("QTY2XFR") = DBNull.Value
            End If
        Next

        If ANY_POPULATED Then
            MsgBox("Transfer quantities auto-populated based on shortage and available US inventory.", MsgBoxStyle.Information)
        Else
            MsgBox("No rows qualified for auto-population.", MsgBoxStyle.Exclamation)
        End If

    End Sub
    'Function Order_Release() As Boolean

    '    ASCMAIN1.Progress("Now Releasing Orders for Shipment", "")

    '    ' Customer Master : Bill-To

    '    ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
    '        & ", ARTCUST1.CUST_PD_GRACE_DAYS, ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.CUST_CREDIT_HOLD" & vbCrLf _
    '        & ", ARTCUST1.CUST_CRED_LIMIT_REV, ARTCUST1.CUST_CREDIT_RELEASE" & vbCrLf _
    '        & " from ARTCUST1" & vbCrLf _
    '        & " where ARTCUST1.CUST_CODE in (Select Distinct CUST_BILL_TO_CUST from " & ARTCUST1 & ")"
    '    Dim ARTCUST1_BT As String = ASCMAIN1.Temp_Table
    '    ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add Primary Key (CUST_CODE)")
    '    ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_AMT_PICK NUMBER (13,2)")
    '    ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_AMT_OPEN NUMBER (13,2)")
    '    ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_AMT_PICK_NOW NUMBER (13,2)")
    '    ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_BALANCE NUMBER (13,2)")
    '    ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_BALANCE_PD NUMBER (13,2)")
    '    ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_HOLDS_CREDIT VARCHAR2(20)")

    '    ASCMAIN1.sql = "" _
    '        & "Begin" & vbCrLf _
    '        & " Declare Cursor C1 is" & vbCrLf _
    '        & "  Select ARTOPEN1.CUST_CODE" & vbCrLf _
    '        & ", Sum (ARTOPEN1.INV_BALANCE) CUST_BALANCE" & vbCrLf _
    '        & ", Sum (CASE WHEN INV_BALANCE > 0 and INV_TYPE = 'I' and ARTOPEN1.INV_DUE_DATE + NVL(ARTCUST1.CUST_PD_GRACE_DAYS,0) +1 < SYSDATE THEN ARTOPEN1.INV_BALANCE ELSE 0 END) CUST_BALANCE_PD" & vbCrLf _
    '        & "   from ARTOPEN1," & ARTCUST1_BT & " ARTCUST1" & vbCrLf _
    '        & "   where ARTCUST1.CUST_CODE = ARTOPEN1.CUST_CODE" & vbCrLf _
    '        & "   group by ARTOPEN1.CUST_CODE;" & vbCrLf _
    '        & " Begin" & vbCrLf _
    '        & "  For R1 in C1 Loop" & vbCrLf _
    '        & "   Update " & ARTCUST1_BT & " ARTCUST1 Set" & vbCrLf _
    '        & "    CUST_BALANCE = R1.CUST_BALANCE" & vbCrLf _
    '        & "   ,CUST_BALANCE_PD = R1.CUST_BALANCE_PD" & vbCrLf _
    '        & "    where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
    '        & "  End Loop;" & vbCrLf _
    '        & " End;" & vbCrLf _
    '        & "End;"
    '    ASCDATA1.ExecuteSQL()

    '    ASCMAIN1.sql = "" _
    '        & "Begin" & vbCrLf _
    '        & " Declare Cursor C1 is" & vbCrLf _
    '        & "  Select NVL(ARTCUST1.CUST_BILL_TO_CUST,SOTORDR0.CUST_CODE) CUST_CODE" & vbCrLf _
    '        & ", Sum (SOTORDR0.ORDR_AMT_OPEN) CUST_AMT_OPEN" & vbCrLf _
    '        & ", Sum (SOTORDR0.ORDR_AMT_PICK) CUST_AMT_PICK" & vbCrLf _
    '        & "   from SOTORDR0, ARTCUST1" & vbCrLf _
    '        & "   where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
    '        & "   group by NVL(ARTCUST1.CUST_BILL_TO_CUST,SOTORDR0.CUST_CODE);" & vbCrLf _
    '        & " Begin" & vbCrLf _
    '        & "  For R1 in C1 Loop" & vbCrLf _
    '        & "   Update " & ARTCUST1_BT & " ARTCUST1 Set" & vbCrLf _
    '        & "    CUST_AMT_PICK = R1.CUST_AMT_PICK" & vbCrLf _
    '        & "   ,CUST_AMT_OPEN = R1.CUST_AMT_OPEN" & vbCrLf _
    '        & "    where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
    '        & "  End Loop;" & vbCrLf _
    '        & " End;" & vbCrLf _
    '        & "End;"
    '    ASCDATA1.ExecuteSQL()

    '    ASCDATA1.ExecuteSQL("Update " & ARTCUST1_BT & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'P' where NVL(CUST_BALANCE_PD,0) > 0 and NVL(CUST_CREDIT_RELEASE,'?') NOT IN ('I','N')")
    '    ASCDATA1.ExecuteSQL("Update " & ARTCUST1_BT & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'C' where NVL(CUST_CREDIT_HOLD,'0') = '1'")
    '    ASCDATA1.ExecuteSQL("Update " & ARTCUST1_BT & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'Z' where (NVL(CUST_CREDIT_LIMIT,0) <=0 or CUST_CRED_LIMIT_REV is Null or CUST_CRED_LIMIT_REV < SYSDATE) and NVL(CUST_CREDIT_RELEASE,'?') <> 'N'")


    '    ' A: Hold Orders where 
    '    ' Customer is on Credit Hold, 
    '    ' or is Past Due, 
    '    ' or has no Credit Limit, 
    '    ' or Credit Limit has expired, 
    '    ' or Customer Aging is Past Due beyond Grace Period

    '    ASCMAIN1.sql = "" _
    '            & "Select CUST_CODE from " & ARTCUST1 & " where CUST_CREDIT_HOLD = '1'" & vbCrLf _
    '            & " union " & vbCrLf _
    '            & "Select CUST_CODE from " & ARTCUST1_BT & " where CUST_HOLDS_CREDIT is not Null"


    '    ASCMAIN1.sql = "Begin Declare Cursor C1 is Select Distinct CUST_CODE from (" & ASCMAIN1.sql & ");" _
    '        & " Begin For R1 in C1 Loop" _
    '        & "  Update " & SOTORDR1 & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'A'" _
    '        & "   where CUST_CODE = R1.CUST_CODE or CUST_BILL_TO_CUST = R1.CUST_CODE or CUST_CREDIT_GROUP_CUST = R1.CUST_CODE;" _
    '        & " End Loop; End; End;"
    '    ASCDATA1.ExecuteSQL()

    '    If (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
    '        If ORDR_GROUP_NO_sql.Length = 0 Then
    '            ASCMAIN1.sql = "Update " & SOTORDR1 & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'B'" _
    '                   & " WHERE ORDR_NO IN " _
    '                   & "(SELECT ORDR_NO" _
    '                   & " FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & SOTORDR1 & ")" _
    '                   & " AND CUST_ORDR_CALL_B4_SHIPPING = '1')"
    '            ASCDATA1.ExecuteSQL()
    '        End If
    '    End If

    '    ' Down Below, Hold Orders if Customer is Over Credit Limit, 

    '    ASCMAIN1.sql = "Select * from " & ARTCUST1_BT
    '    Create_TDA(dst.Tables.Add("ARTCUST1_BT"), ARTCUST1_BT, "**", 0)
    '    Fill_Records("ARTCUST1_BT")

    '    Create_Relation("ARTCUST1_BT", "ARTCUST1", "CUST_CODE", "CUST_BILL_TO_CUST")
    '    ' NEED TO GET THIS SET UP FOR RGI - ARTCUST1 DOES NOT HAVE THESE CHILD FIELDS RIGHT NOW
    '    'With dst.Tables("ARTCUST1_BT")
    '    '    .Columns("CUST_AMT_PICK").Expression = "SUM(CHILD(ARTCUST1_BT_ARTCUST1).CUST_AMT_PICK)"
    '    '    .Columns("CUST_AMT_OPEN").Expression = "SUM(CHILD(ARTCUST1_BT_ARTCUST1).CUST_AMT_OPEN)"
    '    '    .Columns("CUST_AMT_PICK_NOW").Expression = "SUM(CHILD(ARTCUST1_BT_ARTCUST1).CUST_AMT_PICK_NOW)"
    '    'End With

    '    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
    '    Else
    '        ASCMAIN1.sql = " Select ORDR_NO" _
    '             & " from TATTERM1," & SOTORDR1 & " SOTORDR1, ARTCUST1, TATTERM1 TATTERMC " _
    '             & " where SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE" _
    '             & " and TATTERM1.TERM_CODE (+) = SOTORDR1.TERM_CODE" _
    '             & " and TATTERMC.TERM_CODE (+) = ARTCUST1.TERM_CODE" _
    '             & " and (NVL(TATTERM1.TERM_TYPE,'?') = 'R' OR NVL(TATTERMC.TERM_TYPE,'?') = 'R')"

    '        ' Regency will Queue the Credit Card Auths Up and Allow Customer Servce to process in a batch
    '        If Not (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
    '            ASCMAIN1.sql &= " UNION " _
    '                & " Select ORDR_NO" & vbCrLf _
    '                & " from TATTERM1," & SOTORDR1 & " SOTORDR1" _
    '                & " where TATTERM1.TERM_CODE = SOTORDR1.TERM_CODE" & vbCrLf _
    '                & "   and TATTERM1.TERM_TYPE = 'D'" & vbCrLf _
    '                & "   and (SOTORDR1.CCPA_NO is Null or SOTORDR1.CC_TRANS_ID is Null)"
    '        End If

    '        'ASCMAIN1.sql = "Select ORDR_NO" & vbCrLf _
    '        '    & " from TATTERM1," & SOTORDR1 & " SOTORDR1" _
    '        '    & " where TATTERM1.TERM_CODE = SOTORDR1.TERM_CODE" & vbCrLf _
    '        '    & "   and TATTERM1.TERM_TYPE = 'D'" & vbCrLf _
    '        '    & "   and (SOTORDR1.CCPA_NO is Null or SOTORDR1.CC_TRANS_ID is Null)" _
    '        '    & " Union " _
    '        '    & " Select ORDR_NO" _
    '        '    & " from TATTERM1," & SOTORDR1 & " SOTORDR1, ARTCUST1, TATTERM1 TATTERMC " _
    '        '    & " where SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE" _
    '        '    & " and TATTERM1.TERM_CODE (+) = SOTORDR1.TERM_CODE" _
    '        '    & " and TATTERMC.TERM_CODE (+) = ARTCUST1.TERM_CODE" _
    '        '    & " and (NVL(TATTERM1.TERM_TYPE,'?') = 'R' OR NVL(TATTERMC.TERM_TYPE,'?') = 'R')"

    '        ASCMAIN1.sql = "Update " & SOTORDR1 _
    '            & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'R'" _
    '            & "   where ORDR_NO in (" & ASCMAIN1.sql & ")"
    '        ASCDATA1.ExecuteSQL()
    '    End If

    '    ' P: Hold Orders where Pre-Pack Qty is not evenly divisible by Inner Pack, or Styles are not the same for all PPK Lines
    '    ' S: Hold Orders where Customer is on Sales Hold
    '    ' S: Hold all Orders for large customers such as SEARS or KMART unless they were explicitly Requested, or specific groups were selected

    '    If ORDR_GROUP_NO_sql = "" Then
    '        ASCMAIN1.sql = "Update " & ARTCUST1 _
    '            & " Set CUST_SALES_HOLD = '1' " _
    '            & " where CUST_REL_EXPLICITLY = '1'" _
    '            & "   and CUST_CODE NOT in (" & CUST_CODE_sql & ")"
    '    End If

    '    ' O: Hold Orders where Specific Order was indicated to be held

    '    ASCMAIN1.sql = "Update " & SOTORDR1 _
    '        & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'O'" _
    '        & " where ORDR_HOLD = '1'"
    '    ASCDATA1.ExecuteSQL()


    '    ' Set ORDR_RELEASE flag to Ship Style/Color Short if CUST/STYLE Parameter Table indicates ORDR_RELEASE = '1'

    '    '    sql = "Update SOWORDR2,SOWORDR1,ICWSTYC1"
    '    '    & " Set SOWORDR2.ORDR_RELEASE = '1'"
    '    '    & " where SOWORDR2.ORDR_QTY_OPEN <> SOWORDR2.ORDR_QTY_ALLO"
    '    '    & "   and SOWORDR2.ORDR_RELEASE is Null"
    '    '    & "   and SOWORDR2.RANGE_STYLE_CODE is Null"
    '    '    & "   and SOWORDR1.ORDR_NO = SOWORDR2.ORDR_NO"
    '    '    & "   and ICWSTYC1.STYLE_CODE = SOWORDR2.STYLE_CODE"
    '    '    & "   and ICWSTYC1.COLOR_CODE = SOWORDR2.COLOR_CODE"
    '    '    & "   and ICWSTYC1.ORDR_RELEASE = '1'"

    '    ' Set ORDR_RELEASE flag to Ship Style/Color Short if CUST/STYLE/COLOR Parameter Table indicates ORDR_RELEASE = '1'

    '    '    sql = "Update SOWORDR2,SOWORDR1,SOWCSTP1"
    '    '    & " Set SOWORDR2.ORDR_RELEASE = '1'"
    '    '    & " where SOWORDR2.ORDR_QTY_OPEN <> SOWORDR2.ORDR_QTY_ALLO"
    '    '    & "   and SOWORDR2.ORDR_RELEASE is Null"
    '    '    & "   and SOWORDR2.RANGE_STYLE_CODE is Null"
    '    '    & "   and SOWORDR1.ORDR_NO = SOWORDR2.ORDR_NO"
    '    '    & "   and SOWCSTP1.CUST_CODE = SOWORDR2.CUST_CODE"
    '    '    & "   and SOWCSTP1.STYLE_CODE = SOWORDR2.STYLE_CODE"
    '    '    & "   and SOWCSTP1.COLOR_CODE = SOWORDR2.COLOR_CODE"
    '    '    & "   and SOWCSTP1.ORDR_RELEASE = '1'"

    '    ' I, E, F
    '    Inventory_Shortages()

    '    ' N: Hold all orders whose lines total 0 qty allocated

    '    '& "   and (SOWORDR2.ORDR_RELEASE IS NULL OR (SOWORDR2.ORDR_RELEASE <> 'S' AND SOWORDR2.ORDR_RELEASE <> 'C')) "

    '    ' If ARTCUST1.CUST_SHIP_COMPLETE_DETAIL AND SOTORDR2.ORDR_QTY <> SOTORDR2.ORDR_QTY_ALLO_CUR then the sales order cannot be released
    '    ASCMAIN1.sql = "Select Distinct SOTORDR2.ORDR_NO" & vbCrLf _
    '            & " from " & SOTORDR2 & " SOTORDR2" & vbCrLf _
    '            & " where NVL(SOTORDR2.ORDR_QTY,0) <> NVL(SOTORDR2.ORDR_QTY_ALLO_CUR,0) AND ORDR_STATUS IN ('O', 'P')"
    '    ' & "   and SOTORDR2.ORDR_RELEASE is Null"

    '    ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
    '        & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'M'" & vbCrLf _
    '        & " where ORDR_NO in (" & ASCMAIN1.sql & ") AND CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE CUST_SHIP_COMPLETE_DETAIL = '1')"
    '    ASCDATA1.ExecuteSQL()

    '    ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO" _
    '        & " from " & SOTORDR1 & " SOTORDR1," & SOTORDR2 & " SOTORDR2" _
    '        & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
    '        & "   and SOTORDR1.ORDR_REL_HOLD_CODES is Null" _
    '        & " group by SOTORDR1.ORDR_NO" _
    '        & " having SUM (SOTORDR2.ORDR_QTY_ALLO_CUR) = 0"
    '    Dim TBL As DataTable = ASCDATA1.GetDataTable()
    '    ASCMAIN1.sql = "Update " & SOTORDR1 _
    '        & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'N'" _
    '        & " where ORDR_NO in (" & ASCMAIN1.sql & ")"
    '    ASCDATA1.ExecuteSQL()

    '    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
    '    Else

    '        ' Check if Over Credit Limit

    '        ASCMAIN1.sql = "" _
    '            & "Begin Declare Cursor C0 is Select * from " & ARTCUST1_BT & " for Update;" & vbCrLf _
    '            & " Begin" & vbCrLf _
    '            & "  For R0 in C0 Loop" & vbCrLf _
    '            & "   Begin" & vbCrLf _
    '            & "    Declare" & vbCrLf _
    '            & "     CUST_AMT_ALLO_NOW NUMBER (13,2);" & vbCrLf _
    '            & "     Cursor C1 is" & vbCrLf _
    '            & "      Select SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
    '            & "      , SUM (NVL(SOTORDR2.ORDR_QTY_ALLO,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_ALLO" & vbCrLf _
    '            & "       from " & SOTORDR1 & " SOTORDR1," & SOTORDR2 & " SOTORDR2" & vbCrLf _
    '            & "      where SOTORDR1.CUST_BILL_TO_CUST = R0.CUST_CODE" & vbCrLf _
    '            & "        and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
    '            & "      group by SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
    '            & "      order by SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_GROUP_NO;" & vbCrLf _
    '            & "    Begin " & vbCrLf _
    '            & "     CUST_AMT_ALLO_NOW := 0;" & vbCrLf _
    '            & "     For R1 in C1 Loop" & vbCrLf _
    '            & "      CUST_AMT_ALLO_NOW := CUST_AMT_ALLO_NOW + NVL(R1.ORDR_AMT_ALLO,0);" & vbCrLf _
    '            & "      Update " & ARTCUST1_BT & " Set CUST_AMT_PICK_NOW = CUST_AMT_ALLO_NOW where CURRENT of C0;" & vbCrLf _
    '            & "      If NVL(R0.CUST_CREDIT_LIMIT,0) < NVL(R0.CUST_AMT_PICK,0) + CUST_AMT_ALLO_NOW Then" & vbCrLf _
    '            & "       Update " & SOTORDR1 & vbCrLf _
    '            & "        Set ORDR_REL_HOLD_CODES = ORDR_REL_HOLD_CODES || 'L'" & vbCrLf _
    '            & "        where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
    '            & "      End If;" & vbCrLf _
    '            & "     End Loop;" & vbCrLf _
    '            & "    End;" & vbCrLf _
    '            & "   End;" & vbCrLf _
    '            & "  End Loop;" & vbCrLf _
    '            & " End;" & vbCrLf _
    '            & "End;"
    '        ASCDATA1.ExecuteSQL()

    '    End If


    '    ' C: Integrity Checks

    '    Dim Records_Updated As Int64 = 0
    '    Dim ok_to_release As Boolean = True

    '    ASCMAIN1.sql = "Select ORDR_NO from " & SOTORDR1 & " where ORDR_ADDR_TYPE_ST = 'DC' and CUST_DC_NO is Null"
    '    ASCMAIN1.sql = "Update " & SOTORDR1 _
    '        & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'C'" _
    '        & " where ORDR_NO in (" & ASCMAIN1.sql & ")"
    '    'If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
    '    '    If Format(Now, "yyyyMMdd") < "20130320" Then
    '    '        ASCMAIN1.sql &= " and 1<>1"
    '    '    End If
    '    'End If
    '    Records_Updated = ASCDATA1.ExecuteSQL()
    '    If Records_Updated <> 0 Then ok_to_release = False

    '    If Records_Updated <> 0 Then
    '        Dim f As New ASFMSGBF
    '        Dim tblNoDC As DataTable = ASCDATA1.GetDataTable("Select ORDR_NO, CUST_CODE, CUST_NAME, ORDR_CUST_PO, ORDR_GROUP_NO, CUST_STORE_NO, ORDR_ADDR_TYPE_ST, CUST_DC_NO from " & SOTORDR1 & " where ORDR_ADDR_TYPE_ST = 'DC' and CUST_DC_NO is Null")
    '        f.Show_grd(tblNoDC, Me, "Orders shipping to DC with No DC")
    '    End If

    '    ' X: Hold all orders within a group if both of the following are true:
    '    '      1) any single order in the group is being held from release
    '    '      2) the customer is flagged with CUST_SHIP_COMPLETE - DISABLED 2/27/01 BY WJZ - 3 STORES OUT OF 28 NATIO10 RELEASED BECAUSE CUSTOMER WAS NOT MARKED AS SHIP_COMPLETE

    '    ASCMAIN1.sql = "Select Distinct SOTORDR1.ORDR_GROUP_NO" _
    '        & " from " & SOTORDR1 & " SOTORDR1," & ARTCUST1 & " ARTCUST1" _
    '        & " where ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" _
    '        & " and SOTORDR1.ORDR_REL_HOLD_CODES is Not Null"
    '    ' ****************************** VAN SPECIFIC SEE #2 ABOVE ****************************
    '    '& " and ARTCUST1.CUST_SHIP_COMPLETE = '1' "
    '    ' ****************************** VAN SPECIFIC SEE #2 ABOVE ****************************
    '    Dim TBL2 As DataTable = ASCDATA1.GetDataTable
    '    ASCMAIN1.sql = "Update " & SOTORDR1 _
    '        & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'X'" _
    '        & " where ORDR_REL_HOLD_CODES is Null and ORDR_GROUP_NO in (" & ASCMAIN1.sql & ")"
    '    ASCDATA1.ExecuteSQL()

    '    ' Clear out any holds against any orders which belong to force-picked groups
    '    ' Identify all Orders which should be part of this Release by setting ORDR_REL_BATCH_NO to XNO


    '    ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1 set ORDR_REL_BATCH_NO = '" & XNO & "'" _
    '    '& "   and ORDR_REL_HOLD_CODES is Null " ' THIS WAY WE MARK ALL ORDERS THAT WERE ATTEMPTED TO BE RELEASED, WHETHER SUCCESSFUL OR NOT
    '    RELEASE_SQL = " where ORDR_SHIP_DATE <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'"


    '    If SALES_DIVISION_CODE_sql <> "" Then
    '        RELEASE_SQL &= " and SALES_DIVISION_CODE in (" & SALES_DIVISION_CODE_sql & ")"
    '    End If

    '    If CUST_CODE_sql <> "" Then
    '        RELEASE_SQL &= " and CUST_CODE in (" & CUST_CODE_sql & ")"
    '    End If

    '    If ORDR_GROUP_NO_sql <> "" Then
    '        If blnORDR_GROUP_NO_sql_NOT Then
    '            RELEASE_SQL &= " and ORDR_GROUP_NO not in (" & ORDR_GROUP_NO_sql & ")"
    '        Else
    '            RELEASE_SQL &= " and ORDR_GROUP_NO in (" & ORDR_GROUP_NO_sql & ")"
    '        End If
    '    End If

    '    If TERM_CODE_sql <> "" Then
    '        RELEASE_SQL &= " and TERM_CODE in (" & TERM_CODE_sql & ")"
    '    End If

    '    If blnMANUAL_ONLY Then

    '        If ASCMAIN1.CLIENT = "RGI" Then
    '            RELEASE_SQL &= " and ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDRG_manual & ")"
    '        End If

    '        'Dim sqlU As String = ""
    '        'For Each row As DataRow In DirectCast(grdSOTORDRU.DataSource, DataTable).Select("SEL='1'")
    '        '    sqlU &= ",'" & row.Item("USER_ID") & "'"
    '        'Next
    '        'If sqlU <> "" Then
    '        '    sqlU = " and SOTORDRG.ORDR_REL_SHORT_OPER in (" & Mid(sqlU, 2) & ")"
    '        'End If

    '        'RELEASE_SQL &= " and ORDR_GROUP_NO in " _
    '        '    & " (Select SOTORDRG.ORDR_GROUP_NO from SOTORDRG," & SOTORDR1 & " SOTORDR1" _
    '        '    & "   where SOTORDRG.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO and SOTORDRG.ORDR_REL_SHORT = '1'" & sqlU & ")"
    '    End If


    '    ASCMAIN1.sql &= RELEASE_SQL
    '    ASCDATA1.ExecuteSQL()

    '    Release_Exceptions()
    '    ASCMAIN1.Progress(String.Empty, String.Empty)

    '    ' 03/21/2019 - Try Credit Card auth here not after release. Mark with Hold Code R if it fails
    '    If ASCMAIN1.CLIENT = "RGI" Then

    '        ASCMAIN1.sql = "Select * from " & SOTORDR1 & " where ORDR_REL_BATCH_NO = '" & XNO & "' and ORDR_REL_HOLD_CODES is Null and TERM_CODE IN (SELECT TERM_CODE FROM TATTERM1 WHERE TERM_TYPE = 'D')"
    '        Dim tblCC As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

    '        ASCMAIN1.sql = "Select * from " & SOTORDR2 & " WHERE ORDR_NO IN (" & ASCMAIN1.sql.Replace("*", "ORDR_NO") & ")"
    '        Dim tblSOTORDR2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR2")

    '        For Each rowSOTORDR1_rel As DataRow In tblCC.Select("", "CUST_CODE, ORDR_NO")
    '            Dim ORDR_NO As String = rowSOTORDR1_rel.Item("ORDR_NO") & String.Empty

    '            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_QTY_OPEN <> 0"
    '            Dim totalSales As Decimal = 0

    '            For Each rowSOTORDR2_rel As DataRow In tblSOTORDR2.Select(sqlw, "ORDR_LNO")

    '                Dim qA As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_ALLO_CUR") & "")
    '                Dim qO As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") & "")


    '                totalSales += (qA * Val(rowSOTORDR2_rel.Item("ORDR_UNIT_PRICE") & String.Empty))
    '            Next
    '        Next
    '    End If

    '    ASCMAIN1.sql = "Select Distinct SOTORDRG.ORDR_GROUP_NO" _
    '        & " from " & SOTORDR1 & " SOTORDR1, SOTORDRG" _
    '        & " where SOTORDRG.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" _
    '        & "   and SOTORDRG.ORDR_REL_SHORT = '1'" _
    '        & "   and SOTORDR1.ORDR_REL_BATCH_NO = '" & XNO & "'" _
    '        & "   and SOTORDR1.ORDR_REL_HOLD_CODES is Null"
    '    ASCMAIN1.sql = "Update SOTORDRG" _
    '        & " Set ORDR_REL_SHORT = '0', ORDR_REL_BATCH_NO = '" & XNO & "'" _
    '        & " where ORDR_GROUP_NO in (" & ASCMAIN1.sql & ")"
    '    ASCDATA1.ExecuteSQL()

    '    ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1 set ORDR_REL_HOLD_CODES = null" _
    '        & " where ORDR_REL_BATCH_NO is Null"
    '    ASCDATA1.ExecuteSQL()

    '    Return ok_to_release

    'End Function




End Class