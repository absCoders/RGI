Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFP2LC1

#Region "Declarations"

    Dim WHTWAVEX As String = ""

    Dim WAVE_NO As String = ""
    Dim rowWHTWAVE1 As DataRow

    Dim SHIP_BOL_NOs As String = ""
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

    Dim AppearanceRed As New Infragistics.Win.Appearance
    Dim AppearanceEmpty As New Infragistics.Win.Appearance

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        AppearanceEmpty.ForeColor = Color.Empty
        AppearanceRed.ForeColor = Color.Red

        Create_WorkTables()

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select WHTWAVEX.*, WHTWAVE1.CUST_CODE, WHTWAVE1.WAVE_DATE, WHTWAVE1.WHSE_CODE, WHTWAVE1.P2L_LINE_ID" & vbCrLf _
                & $" from {WHTWAVEX} WHTWAVEX, WHTWAVE1 where WHTWAVE1.WAVE_NO = WHTWAVEX.WAVE_NO"
            Create_TDA(.Tables.Add, "WHTWAVEX", "**", 0, False)

            ASCMAIN1.sql = "Select WHTWAVE3.*, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.ORDR_CUST_PO, SOTSHIP1.SHIP_ADDR_CODE" & vbCrLf _
                & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR0.ORDR_CNT_PICK, SOTORDR0.ORDR_QTY_PICK" & vbCrLf _
                & " from WHTWAVE3, SOTORDR0, SOTSHIP1" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "WHTWAVE3", "**", 0, True, "V", 2)
            .Tables("WHTWAVE3").Columns.Add("SELECTED")
            .Tables("WHTWAVE3").Columns("SELECTED").DefaultValue = "0"


            ASCMAIN1.sql = "Select WHTWAVE3.SHIP_BOL_NO, SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE" & vbCrLf _
                & ", Sum (QTY_PACKED) QTY_PACKED" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTCART1, SOTCART2" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & " group by WHTWAVE3.SHIP_BOL_NO,SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTWAVEZ", "**", 0, False, "V", 3)

            Create_Relation("WHTWAVE3", "WHTWAVEZ", "SHIP_BOL_NO")
            With .Tables("WHTWAVEZ").Columns
                .Add("SELECTED", GetType(System.String), "PARENT(WHTWAVE3_WHTWAVEZ).SELECTED")
                .Add("QTY_2BI", GetType(System.Int32), "IIF(SELECTED='1',QTY_PACKED,0)")
            End With

            ASCMAIN1.sql = "Select SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE" & vbCrLf _
                & ", Sum (QTY_PACKED) QTY_PACKED" & vbCrLf _
                & ", Sum (DECODE(WHTWAVE3.P2L_SHIP_STATUS,'P', QTY_PACKED,0)) QTY_P2L_P" & vbCrLf _
                & ", Sum (DECODE(WHTWAVE3.P2L_SHIP_STATUS,'O', QTY_PACKED,0)) QTY_P2L_O" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTCART1, SOTCART2" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & " group by SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTWAVES", "**", 0, False, "V", 2)

            Create_Relation("WHTWAVES", "WHTWAVEZ", "STYLE_CODE,COLOR_CODE")
            With .Tables("WHTWAVES").Columns
                .Add("QTY_2BI", GetType(System.Int32), "SUM(CHILD(WHTWAVES_WHTWAVEZ).QTY_2BI)")
                .Add("QTY_ON_HAND", GetType(System.Int32))
                .Add("QTY_COMM", GetType(System.Int32))
                .Add("QTY_AVA", GetType(System.Int32), "ISNULL(QTY_ON_HAND,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_2BI,0)")
                .Add("QTY_WO", GetType(System.Int32))
                .Add("QTY_NET", GetType(System.Int32), "ISNULL(QTY_AVA,0)+ISNULL(QTY_WO,0)")
            End With

            ASCMAIN1.sql = "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
                & ", Sum (LOCATION_QTY) LOCATION_QTY" & vbCrLf _
                & ", Sum (LOCATION_QTY_WAVE) LOCATION_QTY_WAVE" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTCART1, SOTCART2, WHTLOCB1, WHTWAVE1" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and WHTWAVE1.WAVE_NO = WHTWAVE3.WAVE_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & "   and WHTLOCB1.WHSE_CODE = WHTWAVE1.WHSE_CODE" & vbCrLf _
                & "   and WHTLOCB1.LOCATION_CODE = WHTWAVE1.LOCATION_CODE_DEPOSIT" & vbCrLf _
                & "   and WHTLOCB1.BAR_CODE = '0000000000'" & vbCrLf _
                & "   and WHTLOCB1.STYLE_CODE = SOTCART2.STYLE_CODE" & vbCrLf _
                & "   and WHTLOCB1.COLOR_CODE = SOTCART2.COLOR_CODE" & vbCrLf _
                & " group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTWAVEQ", "**", 0, False, "V", 3)

            ASCMAIN1.sql = "Select SOTCART1.CART_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_DC_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_NO" & vbCrLf _
                & " from SOTCART1, SOTPICK1, SOTORDR1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCART2.CART_NO, SOTCART2.QTY_PACKED, SOTCART2.STYLE_CODE, SOTCART2.COLOR_CODE, WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & " from SOTCART2, SOTCART1, SOTPICK1, WHTSCSEQ, WHTLOCM1, SOTORDR1" & vbCrLf _
                & " where SOTPICK1.SHIP_BOL_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and WHTSCSEQ.STYLE_CODE = SOTCART2.STYLE_CODE" & vbCrLf _
                & "   And WHTSCSEQ.COLOR_CODE = SOTCART2.COLOR_CODE" & vbCrLf _
                & "   And WHTSCSEQ.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "   And WHTLOCM1.LOCATION_ROUTE_SEQ = WHTSCSEQ.STYLE_SEQ" & vbCrLf _
                & "   And WHTLOCM1.LOCATION_CODE Like :PARM2" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = SOTORDR1.WHSE_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "VV", 0)

        End With

        grdWHTWAVEX.DataSource = dst.Tables("WHTWAVEX")
        grdWHTWAVE3.DataSource = dst.Tables("WHTWAVE3")
        grdWHTWAVES.DataSource = dst.Tables("WHTWAVES")

        With grdWHTWAVEX.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
            'For Each COLUMN_NAME As String In New String() {"SELECTED", "SHIP_BOL_NO", "ORDR_GROUP_NO", "PICK_BATCH_NO", "CUST_CODE", "CUST_CODE", "ORDR_CUST_PO"}
            '    .Columns(COLUMN_NAME).Header.Fixed = True
            'Next

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If GCOL.Key = "SHIP_STYLES" Then
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                ElseIf New String() {"WAVE_NO", "ORDR_GROUP_NO", "WAVE_DATE", "CUST_CODE", "WHSE_CODE", "ORDR_TYPE_CODE", "P2L_LINE_ID"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                Else
                    If GCOL.Key.EndsWith("_2BI") Then
                        GCOL.Header.Appearance.BackColor2 = Color.Violet
                    Else
                        GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                    End If
                End If
            Next
        End With


        With grdWHTWAVE3.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit

                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = Activation.AllowEdit
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                ElseIf New String() {"SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO", "SHIP_ADDR_CODE", "ORDR_GROUP_NO", "ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                Else
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                End If
            Next
        End With


        With grdWHTWAVES.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit

                If GCOL.Key = "QTY_2BI" Then
                    GCOL.Header.Appearance.BackColor2 = Color.Orange
                ElseIf New String() {"STYLE_CODE", "COLOR_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf New String() {"QTY_PACKED", "QTY_P2L_P", "QTY_P2L_O"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Violet
                Else
                    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                End If
            Next
        End With

        Create_Summary(grdWHTWAVEX, "WAVE_NO", "Count")
        Create_Summary(grdWHTWAVEX, New String() {"SHIP_CNT", "SHIP_CNT_2BI", "SHIP_CTNS", "SHIP_CTNS_2BI", "SHIP_UNITS", "SHIP_UNITS_2BI"})

        Create_Summary(grdWHTWAVE3, "SHIP_BOL_NO", "Count")
        Create_Summary(grdWHTWAVE3, New String() {"SELECTED", "ORDR_CNT_PICK", "ORDR_QTY_PICK"})

        Create_Summary(grdWHTWAVES, "STYLE_CODE", "Count")
        Create_Summary(grdWHTWAVES, New String() {"QTY_PACKED", "QTY_P2L_P", "QTY_P2L_O", "QTY_2BI", "QTY_ON_HAND", "QTY_COMM", "QTY_AVA", "QTY_WO", "QTY_NET"})

        Show_Filter(grdWHTWAVEX, True)

        'ASCMAIN1.Add_Value_List(grdWHTWAVE3, "ORDR_SOURCE", Nothing, New String() {":", "K:Keyboard", "W:Web", "E:EDI"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("WAVE_NO").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Wave"
                Else
                    WAVE_NO = Absx1.txtFor("WAVE_NO").Text
                    rowWHTWAVE1 = LookUp("WHTWAVE1", WAVE_NO)
                    If rowWHTWAVE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Wave"
                    Else
                        If rowWHTWAVE1.item("P2L_WAVE_STATUS") <> "P" Then
                            EMsg &= vbCrLf & "Wave is not Pending P2L Induction"
                        End If
                    End If
                End If

                If Not ASCMAIN1.Logical_Lock("WHTWAVE1", WAVE_NO) Then Exit Sub

                If EMsg = "" Then

                End If

            Case "Refresh"

            Case "Update"

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

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Refresh"
                Refresh_WHTWAVEX()

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Load").Settings.Enabled = not_iScreenMode
                .Items("Refresh").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Done").Visible = False
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End With

        grdWHTWAVEX.Visible = Not ScreenMode
        splMain.Visible = ScreenMode


        lblCUST_CODE.Visible = ScreenMode
        txtCUST_CODE.Visible = ScreenMode
        lblP2L_LINE_ID.Visible = ScreenMode
        txtP2L_LINE_ID.Visible = ScreenMode
        lblWAVE_DATE.Visible = ScreenMode
        dteWAVE_DATE.Visible = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        'For Each TABLE_NAME As String In New String() _
        '    {"SOTORDR1", "SOTORDR2"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next
        EnforceConstraints(True)

        WHSE_CODE = ""
        Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""
        Refresh_WHTWAVEX()

    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        'WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        'Refresh_SOTSHIPX()

        EnforceConstraints(False)

        rowWHTWAVE1 = LookUp("WHTWAVE1", WAVE_NO)
        txtCUST_CODE.Text = rowWHTWAVE1.Item("CUST_CODE")
        txtP2L_LINE_ID.Text = rowWHTWAVE1.Item("P2L_LINE_ID")
        dteWAVE_DATE.Value = rowWHTWAVE1.Item("WAVE_DATE")

        Fill_Records("WHTWAVE3", WAVE_NO)
        Sort_grdColumns(grdWHTWAVE3, "SHIP_BOL_NO")

        Fill_Records("WHTWAVES", WAVE_NO)
        Sort_grdColumns(grdWHTWAVES, "STYLE_CODE, COLOR_CODE")

        Fill_Records("WHTWAVEZ", WAVE_NO)

        Fill_Records("WHTWAVEQ", WAVE_NO)
        For Each rowWHTWAVEQ As DataRow In dst.Tables("WHTWAVEQ").Select("")
            Dim STYLE_CODE As String = rowWHTWAVEQ.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTWAVEQ.Item("COLOR_CODE")
            Dim LOCATION_QTY As Int32 = Val(rowWHTWAVEQ.Item("LOCATION_QTY") & "")
            Dim LOCATION_QTY_WAVE As Int32 = Val(rowWHTWAVEQ.Item("LOCATION_QTY_WAVE") & "")
            Dim rowWHTWAVES As DataRow = dst.Tables("WHTWAVES").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            rowWHTWAVES.Item("QTY_ON_HAND") = LOCATION_QTY
            rowWHTWAVES.Item("QTY_COMM") = LOCATION_QTY_WAVE
        Next


        EnforceConstraints(True)

        Setup_tabWHTWAVEX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        dst.Tables("WHTWAVE3").AcceptChanges()

        For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("SELECTED = '1' and P2L_SHIP_STATUS = 'O'")
            Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
            rowWHTWAVE3.Item("P2L_SHIP_STATUS") = "P"
            Create_P2L_xml(SHIP_BOL_NO)
        Next

        For Each rowWHTWAVE3 As DataRow In dst.Tables("WHTWAVE3").Select("SELECTED = '0' and P2L_SHIP_STATUS = 'P'")
            Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
            rowWHTWAVE3.Item("P2L_SHIP_STATUS") = "O"
            Create_P2L_Delete_xml(SHIP_BOL_NO)
        Next

        Update_Record_TDA("WHTWAVE3")

        CommitTrans("")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTWAVE3, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Select All X")
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
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdSOTPICKX"
                    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All X"), UltraWinToolbars.ButtonTool)
                    If grdWHTWAVE3.ActiveCell Is Nothing OrElse
                            (grdWHTWAVE3.ActiveCell.Value & "" = "" _
                             Or Not New String() {"ORDR_GROUP_NO", "CUST_CODE", "PICK_BATCH_NO"}.Contains(grdWHTWAVE3.ActiveCell.Column.Key)) Then
                        tlb_btn.SharedProps.Visible = False
                        tlb_btn.Tag = ""
                    Else
                        tlb_btn.Tag = grdWHTWAVE3.ActiveCell.Column.Key & " = '" & grdWHTWAVE3.ActiveCell.Value & "'"
                        tlb_btn.SharedProps.Caption = "Select All " & grdWHTWAVE3.ActiveCell.Column.Header.Caption & " = " & grdWHTWAVE3.ActiveCell.Value
                        tlb_btn.SharedProps.Visible = True
                    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All", "Select All X"

                If grd.Name = "grdSOTPICK1" Or grd.Name = "grdSOTCART1" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        grow.Selected = (e.Tool.Key = "Select All")
                    Next
                Else
                    If e.Tool.Key = "Select All X" Then
                        Dim sqlw As String = IIf(e.Tool.Key = "Select All X", e.Tool.Tag, "")
                        For Each rowSOTPICKX As DataRow In dst.Tables("SOTPICKX").Select(sqlw)
                            rowSOTPICKX.Item("SELECTED") = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                        Next
                    Else
                        For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                            grow.Cells("SELECTED").Value = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                            grow.Update()
                        Next
                    End If
                    '    Display_Totals()
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")

            Case "Shipment Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("LAST_SHIP_KEY").Value
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")

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

    Sub Refresh_WHTWAVEX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Waves")

        Create_WorkTables()
        Fill_Records("WHTWAVEX")
        Sort_grdColumns(grdWHTWAVEX, "WAVE_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Create_WorkTables()

        Dim sqlWHTWAVEX As String = "Select WHTWAVE1.WAVE_NO from WHTWAVE1 where P2L_WAVE_STATUS = 'P'"

        If WHTWAVEX = "" Then
            WHTWAVEX = ASCMAIN1.Temp_Table(sqlWHTWAVEX)
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CNT NUMBER (3,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CTNS NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_UNITS NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CNT_2BI NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CTNS_2BI NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_UNITS_2BI NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_STYLES NUMBER (7,0)")
        Else
            ASCMAIN1.sql = $"Truncate Table {WHTWAVEX}"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Insert into {WHTWAVEX} (WAVE_NO) " & sqlWHTWAVEX
            ASCDATA1.ExecuteSQL()

            Dim sqlC As String = " where Current of C1"
            ASCMAIN1.sql = "" _
                & $"Begin" & vbCrLf _
                & $" Declare Cursor C1 is Select * from {WHTWAVEX} for Update;" & vbCrLf _
                & $" Begin" & vbCrLf _
                & $"  For R1 in C1 Loop" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CNT       = (Select Count(*) from WHTWAVE3 where WAVE_NO = R1.WAVE_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CNT_2BI   = (Select Count(*) from WHTWAVE3 where WAVE_NO = R1.WAVE_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS      = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS_2BI  = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS     = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS_2BI = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_STYLES    = (Select Count (Distinct SOTCART2.STYLE_CODE) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO) {sqlC};" & vbCrLf _
                & $"  End Loop;" & vbCrLf _
                & $" End;" & vbCrLf _
                & $"End;"
            ASCDATA1.ExecuteSQL()
        End If

    End Sub

    Private Sub grdWHTWAVEX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdWHTWAVEX.DoubleClickRow
        If e.Row.IsFilterRow Or Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Dim WAVE_NO As String = e.Row.Cells("WAVE_NO").Value
        Absx1.txtFor("WAVE_NO").Text = WAVE_NO
        Click_Command("Load")
    End Sub

    Private Sub tabWHTWAVEX_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabWHTWAVEX.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabWHTWAVEX()
    End Sub

    Sub Setup_tabWHTWAVEX()

        Dim dvw As DataView = DirectCast(grdWHTWAVE3.DataSource, DataTable).DefaultView

        If tabWHTWAVEX.SelectedTab.Key = "To Be Inducted" Then
            grdWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            grdWHTWAVE3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            ' grdWHTWAVE3.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False
            dvw.RowFilter = "P2L_SHIP_STATUS = 'O'"
            grdWHTWAVE3.Text = "Shipments to be Inducted"

        ElseIf tabWHTWAVEX.SelectedTab.Key = "Already Inducted" Then
            grdWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            grdWHTWAVE3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            ' grdWHTWAVE3.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True
            dvw.RowFilter = "P2L_SHIP_STATUS = 'P'"
            grdWHTWAVE3.Text = "Shipments already Inducted"
        End If

    End Sub

    Private Sub grdWHTWAVES_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTWAVES.InitializeRow

        Dim QTY_AVA As Int32 = Val(e.Row.Cells("QTY_AVA").Value & "")
        If QTY_AVA < 0 Then
            e.Row.Cells("QTY_AVA").Appearance = AppearanceRed
        Else
            e.Row.Cells("QTY_AVA").Appearance = AppearanceEmpty
        End If
    End Sub

    Private Sub Create_P2L_xml(SHIP_BOL_NO As String)

        Dim xmlString As New System.Text.StringBuilder

        Dim P2L_LINE_ID As String = rowWHTWAVE1.Item("P2L_LINE_ID")

        Fill_Records("SOTCART1", SHIP_BOL_NO)
        Fill_Records("SOTCART2", New String() {SHIP_BOL_NO, P2L_LINE_ID & "%"})

        xmlString.AppendLine("<LPXML>")
        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("", "CART_NO")
            Dim CART_NO As String = rowSOTCART1("CART_NO")
            xmlString.AppendLine($"<PickOrder PickOrderNumber='{CART_NO}'>")

            Dim ORDR_CUST_PO As String = rowSOTCART1("ORDR_CUST_PO")
            Dim CUST_DC_NO As String = rowSOTCART1("CUST_DC_NO")
            Dim CUST_STORE_NO As String = rowSOTCART1("CUST_STORE_NO")
            Dim ORDR_NO As String = rowSOTCART1("ORDR_NO")
            Dim PICK_NO As String = rowSOTCART1("PICK_NO")
            xmlString.AppendLine($"<PickOrderXtra ORDR_CUST_PO='{ORDR_CUST_PO}' CUST_DC_NO='{CUST_DC_NO}' CUST_STORE_NO='{CUST_STORE_NO}' ORDR_NO='{ORDR_NO}' PICK_NO='{PICK_NO}'/>")

            For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select($"CART_NO = '{CART_NO}' and QTY_PACKED <> 0", "LOCATION_CODE")
                Dim LOCATION_CODE As String = rowSOTCART2("LOCATION_CODE")
                Dim STYLE_CODE As String = rowSOTCART2("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTCART2("COLOR_CODE")
                Dim QTY_PACKED As Int32 = Val(rowSOTCART2("QTY_PACKED") & "")
                xmlString.AppendLine($"<PickLine LocationName='{LOCATION_CODE}' PickOrderQty='{CStr(QTY_PACKED)}'>")
                xmlString.AppendLine($"<PickLineXtra STYLE_CODE='{STYLE_CODE}' COLOR_CODE='{COLOR_CODE}'/>")
                xmlString.AppendLine("</PickLine>")
            Next

            xmlString.AppendLine("</PickOrder>")
        Next
        xmlString.AppendLine("</LPXML>")

        Dim doc As New System.Xml.XmlDocument()
        doc.LoadXml(xmlString.ToString)
        doc.Save($"{ASCMAIN1.Folders("Work")}{SHIP_BOL_NO}.xml")

        'INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES(xmlString.ToString)

    End Sub

    Private Sub Create_P2L_Delete_xml(SHIP_BOL_NO As String)

        Dim xmlString As New System.Text.StringBuilder
        Stop

    End Sub

End Class