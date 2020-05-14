Public Class SOFCARTP

    Public CUST_CODE As String
    Public ORDR_GROUP_NOs As List(Of String)
    Public PACK_NOs As New List(Of String)
    Public frm As ASFBASE0
    Public PACK_NO As String
    Dim SOTPCKC4 As String
    Public tblSOTPCKP2_existing As DataTable
    Dim PackColumnCount As Integer = 0
    Dim PACK_CART_NO_max As Integer = 1

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst


            Create_TDA(.Tables.Add, "SOTPCKP1", "*")
            Create_TDA(.Tables.Add, "SOTPCKP2", "*")

            Create_TDA(.Tables.Add, "SOTPCKC1", "*")
            With .Tables("SOTPCKC1")
                .Columns.Add("ORDR_QTY_TOTAL", GetType(System.Int32), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDERS,0)")
            End With

            Create_TDA(.Tables.Add, "SOTPCKC2", "*")
            With .Tables("SOTPCKC2")
                Dim TOTAL_PACKS As String = ""
                For p As Integer = 1 To 9
                    Dim COLUMN_NAME As String = "ORDR_QTY_PACK_" & Format(p, "0")
                    TOTAL_PACKS &= "+ISNULL(" & COLUMN_NAME & ",0)"
                    .Columns.Add(COLUMN_NAME, GetType(System.Int32))
                Next
                .Columns.Add("ORDR_QTY_PACK", GetType(System.Int32), Mid(TOTAL_PACKS, 2))
                .Columns.Add("ORDR_QTY_LEFT", GetType(System.Int32), "ISNULL(ORDR_QTY_OPEN,0) - ISNULL(ORDR_QTY_PACK,0)")

            End With


            Create_TDA(.Tables.Add, "SOTPCKC3", "*", 1)
            Create_TDA(.Tables.Add, "SOTPCKC4", "*")


            PACK_NO = ASCMAIN1.Next_Control_No("SOTPCKP1.PACK_NO")

            Dim tbl As DataTable = frm.dst.Tables("SOTORDR0").Clone
            .Tables.Add(tbl)
            '   tbl.Columns.Add("PACK_NO")
            tblSOTPCKP2_existing.PrimaryKey = New DataColumn() {tblSOTPCKP2_existing.Columns("ORDR_GROUP_NO")}

            PACK_NOs.Clear()

            For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
                Dim rowSOTORDR0 As DataRow = frm.dst.Tables("SOTORDR0").Rows.Find(New String() {"O", ORDR_GROUP_NO})
                Dim row As DataRow = .Tables("SOTORDR0").Rows.Add(rowSOTORDR0.ItemArray)

                Dim rowSOTORDR2_existing = tblSOTPCKP2_existing.Rows.Find(ORDR_GROUP_NO)
                If rowSOTORDR2_existing IsNot Nothing Then
                    Dim PACK_NO_existing As String = rowSOTORDR2_existing.Item("PACK_NO")
                    If Not PACK_NOs.Contains(PACK_NO_existing) Then
                        PACK_NOs.Add(PACK_NO_existing)
                        Fill_Records("SOTPCKC3", PACK_NO_existing)
                    End If
                    '  row.Item("PACK_NO") = PACK_NO_existing
                End If

                .Tables("SOTPCKP2").Rows.Add(New String() {PACK_NO, ORDR_GROUP_NO, "A"})
            Next

            Dim rowSOTPCKP1 As DataRow = dst.Tables("SOTPCKP1").NewRow
            With rowSOTPCKP1
                .Item("PACK_NO") = PACK_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
            End With
            dst.Tables("SOTPCKP1").Rows.Add(rowSOTPCKP1)

            Create_SOTPCKC4_Worktable()

            Dim PACK_CONFIG_NOs As New List(Of String)

            For Each rowSOTPCKC3 As DataRow In dst.Tables("SOTPCKC3").Select("PACK_CART_NO > 1")
                Dim PACK_CONFIG_NO As String = rowSOTPCKC3.Item("PACK_CONFIG_NO")
                Dim PACK_CART_NO As Integer = Val(rowSOTPCKC3.Item("PACK_CART_NO") & "")
                Dim ORDR_LNO As Integer = Val(rowSOTPCKC3.Item("ORDR_LNO") & "")
                Dim ORDR_QTY_PACK As Integer = Val(rowSOTPCKC3.Item("ORDR_QTY_PACK") & "")

                If PACK_CART_NO > PACK_CART_NO_max Then
                    PACK_CART_NO_max = PACK_CART_NO
                End If

                If Not PACK_CONFIG_NOs.Contains(PACK_CONFIG_NO) Then
                    PACK_CONFIG_NOs.Add(PACK_CONFIG_NO)
                End If

                Dim rowSOTPCKC2 As DataRow = dst.Tables("SOTPCKC2").Rows.Find(New Object() {PACK_NO, PACK_CONFIG_NO, ORDR_LNO})
                If rowSOTPCKC2 IsNot Nothing Then
                    rowSOTPCKC2.Item("ORDR_QTY_PACK_" & Format(PACK_CART_NO, "0")) = ORDR_QTY_PACK
                    rowSOTPCKC2.Item("ORDR_QTY_PACK_1") = Val(rowSOTPCKC2.Item("ORDR_QTY_PACK_1")) - ORDR_QTY_PACK
                End If

            Next

            For Each PACK_CONFIG_NO As String In PACK_CONFIG_NOs
                Dim rowSOTPCKC1 As DataRow = dst.Tables("SOTPCKC1").Rows.Find(New String() {PACK_NO, PACK_CONFIG_NO})
                rowSOTPCKC1.Item("PACKS") = PackCount(PACK_CONFIG_NO)
            Next

        End With


        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTPCKC1.DataSource = dst.Tables("SOTPCKC1")
        grdSOTPCKC2.DataSource = dst.Tables("SOTPCKC2")
        grdSOTPCKC4.DataSource = dst.Tables("SOTPCKC4")

        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("PACK_NO").Text = PACK_NO

        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If New String() {"CUST_CODE", "CUST_NAME", "SREP_CODE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE", "WAVE_NO", "EDI_LOAD_ID", _
                                 "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_LAST4"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                End If

                If gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If
                ', "ORDR_ARRIVAL_DATE", "ORDR_LAST_ARRIVAL_DATE"
                If New String() {"ORDR_DATE_RECD", "ORDR_PRIORITY",
                                 "ORDR_RELEASE_AVAIL_MIN", "ORDR_RELEASE_AVAIL_MAX", "ORDR_REL_SHORT", "ORDR_REL_SHORT_OPER",
                                 "ORDR_REL_ACTION_DATE", "ORDR_REL_ACTION_OPER", "TERM_CODE", "LAST_DATE", "LAST_OPER", "ORDR_SHIP_INSTR", "ORDR_MESSAGE", "EDI_PO_TYPE"}.Contains(gcol.Key) Then
                    gcol.Hidden = Not (ASCMAIN1.CLIENT = "RGI")
                End If

                If New String() {"ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_CNT_PICK", "EDI_MERCH_TYPE", "EDI_CONS_NO"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                End If

                If New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT", "WHSE_CODE", "EDI_MERCH_TYPE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "SALES_DIVISION_CODE", "SREP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"CUST_CITY", "CUST_STATE", "CUST_COUNTRY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Hidden = True
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next
        End With


        With grdSOTPCKC2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next

            For Each COLUMN_NAME As String In New String() {"ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "ORDR_QTY_OPEN", "ORDR_QTY_PACK", "ORDR_QTY_LEFT"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
            Next

            .Override.AllowMultiCellOperations = UltraWinGrid.AllowMultiCellOperation.All

            For P As Integer = 1 To 9
                Dim COLUMN_NAME As String = "ORDR_QTY_PACK_" & Format(P, "0")
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    .Header.Caption = "Pack " & Format(P, "0")
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    ' If P > 3 Then
                    .Hidden = True
                    ' End If

                    If P <= 3 Or P <= PACK_CART_NO_max Then
                        AddPackColumn()
                    End If
                End With
            Next
        End With


        With grdSOTPCKC4.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        With grdSOTPCKC1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}, , , "#,##0")

        Create_Summary(grdSOTPCKC1, "PACK_CONFIG_NO", "Count")
        Create_Summary(grdSOTPCKC1, New String() {"ORDERS", "ORDR_QTY_TOTAL"})

        Create_Summary(grdSOTPCKC4, "ORDR_NO", "Count")
        Create_Summary(grdSOTPCKC2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTPCKC2, New String() {"ORDR_QTY_OPEN", "ORDR_QTY_PACK", "ORDR_QTY_LEFT", _
                                                  "ORDR_QTY_PACK_1", "ORDR_QTY_PACK_2", "ORDR_QTY_PACK_3", _
                                                  "ORDR_QTY_PACK_4", "ORDR_QTY_PACK_5", "ORDR_QTY_PACK_6", _
                                                  "ORDR_QTY_PACK_7", "ORDR_QTY_PACK_8", "ORDR_QTY_PACK_9"})

        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), True)
        Set_Read_Only_for_ctl(Absx1.txtFor("PACK_NO"), True)

        Me.Width = 0.9 * frm.Width
        Me.Height = frm.Height

        Me.Top = 0.05 * frm.Height
        Me.Left = 0.05 * frm.Width

        cmdDelete.Visible = (PACK_NOs.Count > 0)
    End Sub


#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPCKC2, "BB", "Split to new Pack", "Add Pack Column")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If
        If Not GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
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

        Select Case e.SourceControl.Name

            Case "grdSOTPCKC2"
                'tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                'tlb_btn = DirectCast(tlb_pop.Tools("Create POs"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = Not ScreenMode And (grdSOTORDRX.ActiveRow IsNot Nothing Or grdSOTORDRX.Selected.Rows.Count <> 0) And (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
 
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
            ' NEED TO GET PAST HERE FOR STYLE MULTI-COLOR WHEN THERE ARE NO ROWS IN THE GRID
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTPCKC2"
                    'tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    'tlb_sbt = DirectCast(tlb_pop.Tools("Show UPC/SKU"), UltraWinToolbars.StateButtonTool)
                    'tlb_sbt.SharedProps.Visible = True ' (Absx1.optFor("ORDR_SOURCE").Value = "K")
                    'tlb_sbt.Tag = "X"
                    'tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Split to new Pack"
                Dim PACK_CONFIG_NO As String = grdSOTPCKC1.ActiveRow.Cells("PACK_CONFIG_NO").Value
                Dim splitP As Integer = 0
                For p As Integer = 2 To 9
                    Dim SQLW As String = "PACK_CONFIG_NO = '" & PACK_CONFIG_NO & "' AND ISNULL(ORDR_QTY_PACK_" & CStr(p) & ",0)<>0"
                    If dst.Tables("SOTPCKC2").Select(SQLW).Length = 0 Then
                        splitP = p
                        Exit For
                    End If
                Next
                If splitP = 0 Then
                    MsgBox("Cannot Split Pack", MsgBoxStyle.OkOnly, "No more packs available, Max = 9")
                    Exit Sub
                End If

                If splitP > PackColumnCount Then
                    AddPackColumn()

                    If splitP > PackColumnCount Then
                        MsgBox("Cannot Split Pack", MsgBoxStyle.OkOnly, "No more packs available, Max = 9")
                        Exit Sub
                    End If
                End If

                If grdSOTPCKC2.Selected.Rows.Count = 0 Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTPCKC2.Rows
                        grow.Selected = True
                    Next
                End If

                Dim COLUMN_NAME_split_from As String = "ORDR_QTY_PACK_1"
                Dim COLUMN_NAME_split_to As String = "ORDR_QTY_PACK_" & CStr(splitP)
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTPCKC2.Selected.Rows
                    Dim QTY As Integer = Val(grow.Cells(COLUMN_NAME_split_from).Value)
                    QTY = QTY / 2
                    grow.Cells(COLUMN_NAME_split_from).Value = Val(grow.Cells(COLUMN_NAME_split_from).Value) - QTY
                    grow.Cells(COLUMN_NAME_split_to).Value = QTY
                    ' grow.Update()
                Next
                grdSOTPCKC2.UpdateData()
                grdSOTPCKC2.Selected.Rows.Clear()


            Case "Add Pack Column"
                AddPackColumn()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Show Document"
 
                 
        End Select
    End Sub

#End Region


#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"
                'If e.KeyCode = Windows.Forms.Keys.Enter Then
                '    Prepare_ICTCOLRM()
                'End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"
                'Prepare_ICTCOLRM()
        End Select
    End Sub
#End Region

    Sub Create_SOTPCKC4_Worktable()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Orders for Pack Configuration")

        ASCMAIN1.sql = "Select '" & PACK_NO & "' PACK_NO, SOTORDR1.ORDR_NO, '000000' PACK_CONFIG_NO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_DC_NO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, COUNT (*) STYLES" & vbCrLf _
            & " from SOTORDR1,SOTORDR2" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_GROUP_NO in ('" & Join(ORDR_GROUP_NOs.ToArray, "','") & "')" & vbCrLf _
            & " group by SOTORDR1.ORDR_NO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_DC_NO"

        SOTPCKC4 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Select * from " & SOTPCKC4
        Fill_Records("SOTPCKC4", , , ASCMAIN1.sql)

        Dim PACK_CONFIG_NO_ctr As Int32 = 0

        ASCMAIN1.sql = "Select ORDR_QTY_OPEN, STYLES from " & SOTPCKC4 & " group by ORDR_QTY_OPEN, STYLES"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ORDR_QTY_OPEN As Int64 = Val(row.Item("ORDR_QTY_OPEN") & "")
            Dim STYLES As Int64 = Val(row.Item("STYLES") & "")

            Dim ORDR_NO_CONFIG As String = ""
            Dim PACK_CONFIG_NO As String = ""

            Dim sqlw As String = "ORDR_QTY_OPEN = " & CStr(ORDR_QTY_OPEN) & " and STYLES = " & CStr(STYLES)

            For Each rowSOTPCKC4 In dst.Tables("SOTPCKC4").Select(sqlw)
                Dim ORDR_NO As String = rowSOTPCKC4.Item("ORDR_NO")
                PACK_CONFIG_NO = ""

                For Each rowSOTPCKC1 As DataRow In dst.Tables("SOTPCKC1").Select("ORDR_QTY_OPEN = " & CStr(ORDR_QTY_OPEN) & " and STYLES = " & CStr(STYLES))

                    ASCMAIN1.sql = "Select Count (*) from (" _
                        & "(" & vbCrLf _
                        & " Select '1' DIFF, ORDR_LNO, STYLE_CODE, COLOR_CODE, ORDR_QTY_OPEN from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                        & "  Minus " & vbCrLf _
                        & " Select '1' DIFF, ORDR_LNO, STYLE_CODE, COLOR_CODE, ORDR_QTY_OPEN from SOTORDR2 where ORDR_NO = '" & ORDR_NO_CONFIG & "'" & vbCrLf _
                        & ")" & vbCrLf _
                        & "  Union " & vbCrLf _
                        & "(" & vbCrLf _
                        & " Select '2' DIFF, ORDR_LNO, STYLE_CODE, COLOR_CODE, ORDR_QTY_OPEN from SOTORDR2 where ORDR_NO = '" & ORDR_NO_CONFIG & "'" & vbCrLf _
                        & "  Minus " & vbCrLf _
                        & " Select '2' DIFF, ORDR_LNO, STYLE_CODE, COLOR_CODE, ORDR_QTY_OPEN from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'" & vbCrLf _
                        & ")" & vbCrLf _
                        & ")"

                    Dim C As Int32 = ASCDATA1.GetDataValue
                    If C = 0 Then
                        PACK_CONFIG_NO = rowSOTPCKC1.Item("PACK_CONFIG_NO")
                        rowSOTPCKC1.Item("ORDERS") = Val(rowSOTPCKC1.Item("ORDERS") & "") + 1
                        rowSOTPCKC1.Item("PACKS") = 1
                        Exit For
                    End If
                Next

                If PACK_CONFIG_NO = "" Then
                    PACK_CONFIG_NO_ctr += 1
                    PACK_CONFIG_NO = Format(PACK_CONFIG_NO_ctr, "000000")

                    If ORDR_NO_CONFIG = "" Then ORDR_NO_CONFIG = ORDR_NO

                    dst.Tables("SOTPCKC1").Rows.Add(New Object() {PACK_NO, PACK_CONFIG_NO, ORDR_NO, ORDR_QTY_OPEN, STYLES, 1})

                    ASCMAIN1.sql = "Select '" & PACK_NO & "' PACK_NO, '" & PACK_CONFIG_NO & "' PACK_CONFIG_NO, ORDR_LNO, STYLE_CODE, COLOR_CODE, ORDR_QTY_OPEN from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
                    Fill_Records("SOTPCKC2", , False, ASCMAIN1.sql)

                End If

                rowSOTPCKC4.Item("PACK_CONFIG_NO") = PACK_CONFIG_NO
            Next
        Next

        For Each rowSOTPCKC2 As DataRow In dst.Tables("SOTPCKC2").Select("")
            rowSOTPCKC2.Item("ORDR_QTY_PACK_1") = rowSOTPCKC2.Item("ORDR_QTY_OPEN")
        Next


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Determine_Store_Configurations()

    End Sub

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click

        EMsg = ""
        Dim msg As String = ""

        dst.Tables("SOTPCKC3").Rows.Clear()

        For Each rowSOTPCKC1 As DataRow In dst.Tables("SOTPCKC1").Select("")
            Dim PACK_CONFIG_NO As String = rowSOTPCKC1.Item("PACK_CONFIG_NO")

            Dim PACK_CART_NOs As New List(Of Integer)

            For Each rowSOTPCKC2 As DataRow In dst.Tables("SOTPCKC2").Select("PACK_CONFIG_NO = '" & PACK_CONFIG_NO & "'")

                Dim STYLE_CODE As String = rowSOTPCKC2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTPCKC2.Item("COLOR_CODE")

                If Val(rowSOTPCKC2.Item("ORDR_QTY_LEFT") & "") <> 0 Then
                    msg = "Unpacked Qtys in Pack Configuration " & PACK_CONFIG_NO
                    If Not EMsg.Contains(msg) Then
                        EMsg &= vbCr & msg & " - see " & STYLE_CODE & "-" & COLOR_CODE
                    End If
                End If
                For P As Integer = 1 To 9
                    Dim ORDR_QTY_PACK_P As Int32 = Val(rowSOTPCKC2.Item("ORDR_QTY_PACK_" & Format(P, "0")) & "")

                    If ORDR_QTY_PACK_P <> 0 Then
                        If ORDR_QTY_PACK_P < 0 Then
                            msg = "Negative Qtys in Pack Configuration " & PACK_CONFIG_NO
                            If Not EMsg.Contains(msg) Then
                                EMsg &= vbCr & msg & " - see " & STYLE_CODE & "-" & COLOR_CODE
                            End If
                        Else
                            If Not PACK_CART_NOs.Contains(P) Then
                                PACK_CART_NOs.Add(P)
                            End If
                            dst.Tables("SOTPCKC3").Rows.Add(New Object() {PACK_NO, PACK_CONFIG_NO, P, _
                                rowSOTPCKC2.Item("ORDR_LNO"), Val(rowSOTPCKC2.Item("ORDR_QTY_PACK_" & Format(P, "0")) & "")})
                        End If
                    End If

                Next
            Next

            rowSOTPCKC1.Item("PACKS") = PACK_CART_NOs.Count
        Next

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If




        BeginTrans()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Pack Configuration")

        Delete_Previously_Packed()

        Dim sqlDelete As String = "PACK_NO = '" & PACK_NO & "'"
        Update_Record_TDA("SOTPCKP1", sqlDelete)
        Update_Record_TDA("SOTPCKP2", sqlDelete)

        Update_Record_TDA("SOTPCKC1", sqlDelete)
        Update_Record_TDA("SOTPCKC2", sqlDelete)
        Update_Record_TDA("SOTPCKC3", sqlDelete)
        Update_Record_TDA("SOTPCKC4", sqlDelete)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        CommitTrans("Pack " & PACK_NO & " has been Updated")


        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            Dim rowSOTORDR0 As DataRow = frm.dst.Tables("SOTORDR0").Rows.Find(New String() {"O", ORDR_GROUP_NO})
            rowSOTORDR0.Item("PACK_NO") = PACK_NO
        Next



        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
    End Sub

    Private Sub grdSOTPCKC1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPCKC1.AfterRowActivate

        Dim PACK_CONFIG_NO As String = grdSOTPCKC1.ActiveRow.Cells("PACK_CONFIG_NO").Value

        Dim dvw As DataView = DirectCast(grdSOTPCKC4.DataSource, DataTable).DefaultView
        dvw.RowFilter = "PACK_CONFIG_NO = '" & PACK_CONFIG_NO & "'"
        Sort_grdColumns(grdSOTPCKC4, "ORDR_NO")
        grdSOTPCKC4.Text = "Orders with Configuation " & PACK_CONFIG_NO

        dvw = DirectCast(grdSOTPCKC2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "PACK_CONFIG_NO = '" & PACK_CONFIG_NO & "'"
        Sort_grdColumns(grdSOTPCKC2, "ORDR_LNO")
        grdSOTPCKC2.Text = "Style/Color/Qty Details for Configuration " & PACK_CONFIG_NO
    End Sub
 
    Private Sub cmdDelete_Click(sender As Object, e As EventArgs) Handles cmdDelete.Click
        If MsgBox("Packs to be Delete from Order Groups Selected: " & Join(PACK_NOs.ToArray, ",") _
            & vbCrLf & vbCrLf & "Are you sure that you want to Delete all previously entered Pack Configurations for all Order Groups shown?", MsgBoxStyle.YesNo, "Warning - this will Delete Packs Configured") = MsgBoxResult.No Then
            Exit Sub
        End If


        BeginTrans()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Deleting Previously Entered Pack Configurations")

        Delete_Previously_Packed()

        'For Each PACK_NO As String In PACK_NOs
        '    For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
        '        ASCMAIN1.sql = "Update SOTPCKP2 Set PACK_GROUP_STATUS = :PARM1, LAST_OPER = :PARM2, LAST_DATE = SYSDATE where ORDR_GROUP_NO = :PARM3"
        '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New String() {"D", ASCMAIN1.USER_ID, ORDR_GROUP_NO})
        '    Next
        'Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        CommitTrans("Packs have been Deleted")

        Me.Close()

    End Sub

    Sub Delete_Previously_Packed()

        For Each ROW As DataRow In tblSOTPCKP2_existing.Select("")
            Dim PACK_NO As String = ROW.Item("PACK_NO")
            Dim ORDR_GROUP_NO As String = ROW.Item("ORDR_GROUP_NO")
            ASCMAIN1.sql = "Update SOTPCKP2 Set PACK_GROUP_STATUS = :PARM1, LAST_OPER = :PARM2, LAST_DATE = SYSDATE where ORDR_GROUP_NO = :PARM3 and PACK_NO = :PARM4 and PACK_GROUP_STATUS = 'A'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VvVV", New String() {"D", ASCMAIN1.USER_ID, ORDR_GROUP_NO, PACK_NO})

            Dim rowSOTORDR0 As DataRow = frm.dst.Tables("SOTORDR0").Rows.Find(New String() {"O", ORDR_GROUP_NO})
            rowSOTORDR0.Item("PACK_NO") = DBNull.Value
        Next
    End Sub


    Function SplitColumn() As Boolean

    End Function

    Function AddPackColumn() As Integer

        If PackColumnCount < 9 Then
            PackColumnCount += 1
            Dim COLUMN_NAME As String = "ORDR_QTY_PACK_" & Format(PackColumnCount, "0")
            grdSOTPCKC2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
        End If

        Return PackColumnCount
    End Function

    Function PackCount(PACK_CONFIG_NO As String) As Integer
        Dim PACK_COUNT As Integer = 0
        For p As Integer = 1 To 9
            If dst.Tables("SOTPCKC2").Select("ISNULL(ORDR_QTY_PACK_" & CStr(p) & ",0)<>0").Length > 0 Then
                PACK_COUNT += 1
            End If
        Next
        Return PACK_COUNT
    End Function

    Private Sub grdSOTPCKC2_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSOTPCKC2.AfterRowUpdate
        Dim PACK_CONFIG_NO As String = e.Row.Cells("PACK_CONFIG_NO").Value
        Dim rowSOTPCKC1 As DataRow = dst.Tables("SOTPCKC1").Rows.Find(New String() {PACK_NO, PACK_CONFIG_NO})
        rowSOTPCKC1.Item("PACKS") = PackCount(PACK_CONFIG_NO)
    End Sub
End Class