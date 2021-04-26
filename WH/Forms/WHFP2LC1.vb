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

    Dim CUST_CODE As String
    Dim P2L_LINE_ID As String

    Dim sqlCS As String = ""

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
                & ", SOTORDR0.ORDR_CNT_PICK PTS, SOTORDR0.ORDR_QTY_PICK UNITS" & vbCrLf _
                & " from WHTWAVE3, SOTORDR0, SOTSHIP1" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "WHTWAVE3", "**", 0, True, "V", 2)

            With .Tables("WHTWAVE3")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
                .Columns.Add("CTNS", GetType(System.Int32))
                .Columns.Add("CTNS_WIP", GetType(System.Int32))
                .Columns.Add("UNITS_WIP", GetType(System.Int32))
            End With


            ASCMAIN1.sql = "Select WHTWAVE3.SHIP_BOL_NO, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                & ", SOTCART1.CART_NO, SOTCART1.CART_PACKER, SOTCART1.CART_PACKED, SOTCART1.PICK_NO, SOTCART1.CART_TOTAL_UNITS" & vbCrLf _
                & " from WHTWAVE3, SOTPICK1, SOTORDR1, SOTCART1" & vbCrLf _
                & " where WHTWAVE3.WAVE_NO = :PARM1" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTCART1.PICK_NO = SOTPICK1.PICK_NO"
            Create_TDA(.Tables.Add, "WHTWAVEC", "**", 0, False, "V", 3)




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
                .Add("P2L_SHIP_STATUS", GetType(System.String), "PARENT(WHTWAVE3_WHTWAVEZ).P2L_SHIP_STATUS")
                .Add("QTY_2BI", GetType(System.Int32), "IIF(SELECTED='1' and P2L_SHIP_STATUS = 'O',QTY_PACKED,0)")
                .Add("QTY_2BD", GetType(System.Int32), "IIF(SELECTED='0' and P2L_SHIP_STATUS = 'P',QTY_PACKED,0)")
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
                .Add("QTY_2BD", GetType(System.Int32), "SUM(CHILD(WHTWAVES_WHTWAVEZ).QTY_2BD)")
                .Add("QTY_ON_HAND", GetType(System.Int32))
                .Add("QTY_WO_PICK", GetType(System.Int32))
                .Add("QTY_COMM", GetType(System.Int32))
                .Add("QTY_AVA", GetType(System.Int32), "ISNULL(QTY_ON_HAND,0)+ISNULL(QTY_WO_PICK,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_2BI,0)+ISNULL(QTY_2BD,0)")
                .Add("QTY_WO_OPEN", GetType(System.Int32))
                .Add("QTY_NET", GetType(System.Int32), "ISNULL(QTY_AVA,0)+ISNULL(QTY_WO_OPEN,0)")
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

            ' maybe this sql needs to avoid looking at waves that have been deposited, 
            '  since the pick would already in the on hand of the Deposit location
            ASCMAIN1.sql = "Select WHTWAVE2.STYLE_CODE, WHTWAVE2.COLOR_CODE" & vbCrLf _
                & ", SUM (DECODE(WHTINST1.WAVE_INST_STATUS,'0',WHTINST2.LOCATION_QTY_WAVE,0)) OPEN" & vbCrLf _
                & ", SUM (DECODE(WHTINST1.WAVE_INST_STATUS,'1',WHTINST2.LOCATION_QTY_PICK,0)) PICK" & vbCrLf _
                & " from WHTINST2,WHTINST1,WHTWAVE2,WHTWAVE1" & vbCrLf _
                & " where WHTINST2.WAVE_INST_NO = WHTINST1.WAVE_INST_NO" & vbCrLf _
                & "   and WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                & "   and WHTWAVE2.WAVE_NO = WHTINST1.WAVE_NO AND WHTWAVE2.WAVE_LNO = WHTINST1.WAVE_LNO" & vbCrLf _
                & "   and WHTWAVE1.P2L_LINE_ID = :PARM1 and WHTWAVE1.P2L_WAVE_STATUS = 'P'" & vbCrLf _
                & " group by WHTWAVE2.STYLE_CODE, WHTWAVE2.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTINSTX", "**", 0, False, "V", 0)

            Create_TDA(.Tables.Add, "WHTP2LP1", "*")
            Create_TDA(.Tables.Add, "WHTP2LP2", "*")

        End With

        grdWHTWAVEX.DataSource = dst.Tables("WHTWAVEX")
        grdWHTWAVE3.DataSource = dst.Tables("WHTWAVE3")
        grdWHTWAVEC.DataSource = dst.Tables("WHTWAVEC")
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
                    ElseIf GCOL.Key.EndsWith("_2BP") Then
                        GCOL.Header.Appearance.BackColor2 = Color.Gold
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
                ElseIf New String() {"CTNS_WIP", "UNITS_WIP"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Color.Gold
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

                If GCOL.Key = "QTY_2BI" Or GCOL.Key = "QTY_2BD" Then
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

        With grdWHTWAVEC.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = Color.White
                GCOL.Header.Appearance.BackColor2 = Color.Gray
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = Activation.NoEdit

                'If GCOL.Key = "QTY_2BI" Or GCOL.Key = "QTY_2BD" Then
                '    GCOL.Header.Appearance.BackColor2 = Color.Orange
                'ElseIf New String() {"STYLE_CODE", "COLOR_CODE"}.Contains(GCOL.Key) Then
                '    GCOL.Header.Appearance.BackColor2 = Color.LightBlue
                'ElseIf New String() {"QTY_PACKED", "QTY_P2L_P", "QTY_P2L_O"}.Contains(GCOL.Key) Then
                '    GCOL.Header.Appearance.BackColor2 = Color.Violet
                'Else
                '    GCOL.Header.Appearance.BackColor2 = Color.LightGreen
                'End If
            Next
        End With

        Create_Summary(grdWHTWAVEX, "WAVE_NO", "Count")
        Create_Summary(grdWHTWAVEX, New String() {"SHIP_CNT", "SHIP_CNT_2BI", "SHIP_CTNS", "SHIP_CTNS_2BI", "SHIP_UNITS", "SHIP_UNITS_2BI", "SHIP_CNT_2BP", "SHIP_CTNS_2BP", "SHIP_UNITS_2BP"})

        Create_Summary(grdWHTWAVE3, "SHIP_BOL_NO", "Count")
        Create_Summary(grdWHTWAVE3, New String() {"SELECTED", "PTS", "UNITS", "CTNS", "CTNS_WIP", "UNITS_WIP"})

        Create_Summary(grdWHTWAVEC, "CART_NO", "Count")
        Create_Summary(grdWHTWAVEC, New String() {"CART_TOTAL_UNITS"})

        Create_Summary(grdWHTWAVES, "STYLE_CODE", "Count")
        Create_Summary(grdWHTWAVES, New String() {"QTY_PACKED", "QTY_P2L_P", "QTY_P2L_O", "QTY_2BI", "QTY_2BD", "QTY_ON_HAND", "QTY_WO_PICK", "QTY_COMM", "QTY_AVA", "QTY_WO_OPEN", "QTY_NET"})

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

            Case "Import P2L Picks"
                'Import_Picks()
                Poll_P2L()

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

                .Items("Import P2L Picks").Visible = Not ScreenMode

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

        CUST_CODE = rowWHTWAVE1.Item("CUST_CODE")
        P2L_LINE_ID = rowWHTWAVE1.Item("P2L_LINE_ID")

        rowWHTWAVE1 = LookUp("WHTWAVE1", WAVE_NO)
        txtCUST_CODE.Text = CUST_CODE
        txtP2L_LINE_ID.Text = P2L_LINE_ID
        dteWAVE_DATE.Value = rowWHTWAVE1.Item("WAVE_DATE")

        Fill_Records("WHTWAVEC", WAVE_NO)

        Fill_Records("WHTWAVE3", WAVE_NO)
        Sort_grdColumns(grdWHTWAVE3, "SHIP_BOL_NO")
        For Each row As DataRow In dst.Tables("WHTWAVE3").Select("P2L_SHIP_STATUS = 'P'")
            row.Item("SELECTED") = "1"
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            Dim CTNS As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}'"))
            row.Item("CTNS") = CTNS
            Dim CTNS_WIP As Int32 = Val(dst.Tables("WHTWAVEC").Compute("COUNT(CART_NO)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NULL"))
            row.Item("CTNS_WIP") = CTNS_WIP
            Dim UNITS_WIP As Int32 = Val(dst.Tables("WHTWAVEC").Compute("SUM (CART_TOTAL_UNITS)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}' and CART_PACKER IS NULL"))
            row.Item("UNITS_WIP") = UNITS_WIP
        Next

        dst.Tables("WHTWAVE3").AcceptChanges()

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


        Fill_Records("WHTINSTX", P2L_LINE_ID)
        For Each rowWHTINSTX As DataRow In dst.Tables("WHTINSTX").Select("")
            Dim STYLE_CODE As String = rowWHTINSTX.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTINSTX.Item("COLOR_CODE")
            Dim OPEN As Int32 = Val(rowWHTINSTX.Item("OPEN") & "")
            Dim PICK As Int32 = Val(rowWHTINSTX.Item("PICK") & "")
            Dim rowWHTWAVES As DataRow = dst.Tables("WHTWAVES").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            'rowWHTWAVES.Item("QTY_ON_HAND") = Val(rowWHTWAVES.Item("QTY_ON_HAND") & "") + PICK
            rowWHTWAVES.Item("QTY_WO_PICK") = PICK
            rowWHTWAVES.Item("QTY_WO_OPEN") = OPEN
        Next

        EnforceConstraints(True)

        tabWHTWAVEX.SelectedTab = tabWHTWAVEX.Tabs("To Be Inducted")
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
            Create_P2L_xml(rowWHTWAVE3)
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
        Load_Popup_Menu(grdWHTWAVE3, "SSBB", "Show Filter", "Show GroupBox", "Select All", "De-Select All")
        Load_Popup_Menu(grdWHTWAVES, "SSB", "Show Filter", "Show GroupBox", "Style Status Inquiry")

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

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                If grd.Name = "grdWHTWAVE3" Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Executing: " & e.Tool.Key)
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        Dim SHIP_BOL_NO As String = grow.Cells("SHIP_BOL_NO").Value
                        ASCMAIN1.Progress("-", SHIP_BOL_NO)
                        grow.Cells("SELECTED").Value = IIf(e.Tool.Key.StartsWith("Select"), "1", "0")
                        grow.Update()
                    Next
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")

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
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CNT_2BP NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_CTNS_2BP NUMBER (7,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {WHTWAVEX} Add SHIP_UNITS_2BP NUMBER (7,0)")
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
                & $"   Update {WHTWAVEX} Set SHIP_CNT_2BP   = (Select Count(*) from WHTWAVE3 where WAVE_NO = R1.WAVE_NO and WHTWAVE3.P2L_SHIP_STATUS = 'P') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS      = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS_2BI  = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_CTNS_2BP  = (Select Count(*) from WHTWAVE3,SOTPICK1,SOTCART1 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and WHTWAVE3.P2L_SHIP_STATUS = 'P' and SOTCART1.CART_PACKER IS NULL) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS     = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO) {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS_2BI = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO and WHTWAVE3.P2L_SHIP_STATUS = 'O') {sqlC};" & vbCrLf _
                & $"   Update {WHTWAVEX} Set SHIP_UNITS_2BP = (Select Sum (SOTCART2.QTY_PACKED) from WHTWAVE3,SOTPICK1,SOTCART1,SOTCART2 where WAVE_NO = R1.WAVE_NO and SOTPICK1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO and SOTCART1.PICK_NO = SOTPICK1.PICK_NO and SOTCART2.CART_NO = SOTCART1.CART_NO and WHTWAVE3.P2L_SHIP_STATUS = 'P' and SOTCART1.CART_PACKER IS NULL) {sqlC};" & vbCrLf _
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
            'grdWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            splWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            grdWHTWAVE3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            ' grdWHTWAVE3.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False
            grdWHTWAVE3.DisplayLayout.Bands(0).Columns("SELECTED").Header.Caption = "Sel"
            dvw.RowFilter = "P2L_SHIP_STATUS = 'O'"
            grdWHTWAVE3.Text = "Shipments to be Inducted"

        ElseIf tabWHTWAVEX.SelectedTab.Key = "Already Inducted" Then
            'grdWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            splWHTWAVE3.Parent = tabWHTWAVEX.SelectedTab.TabPage
            grdWHTWAVE3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            ' grdWHTWAVE3.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True
            grdWHTWAVE3.DisplayLayout.Bands(0).Columns("SELECTED").Header.Caption = "Del"
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

    Private Sub Create_P2L_xml(rowWHTWAVE3 As DataRow)

        Dim SHIP_BOL_NO As String = rowWHTWAVE3.Item("SHIP_BOL_NO")
        Dim QTY_PACKED_WHTWAVEZ As Int32 = Val(dst.Tables("WHTWAVEZ").Compute("SUM(QTY_PACKED)", $"SHIP_BOL_NO = '{SHIP_BOL_NO}'") & "")

        Dim xmlString As New System.Text.StringBuilder

        Dim P2L_LINE_ID As String = rowWHTWAVE1.Item("P2L_LINE_ID")

        Fill_Records("SOTCART1", SHIP_BOL_NO)
        Fill_Records("SOTCART2", New String() {SHIP_BOL_NO, P2L_LINE_ID & "%"})

        Dim QTY_PACKED_SOTCART2 As Int32 = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "") & "")
        If QTY_PACKED_SOTCART2 <> QTY_PACKED_WHTWAVEZ Then
            Throw New Exception("Qty about to send to P2L does not agree with Shipment Qty Released")
        End If

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
        'INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES('<LPXML>…</LPXML>')

        sqlCS = "Data Source= ABSSVR2019; Initial Catalog=LPPick; User Id= abs; Password= v4n$4L3"
        Dim sqlConn As New System.Data.SqlClient.SqlConnection(sqlCS)
        sqlConn.Open()

        'Dim sqlP As New System.Data.SqlClient.SqlParameter("@parm", SqlDbType.Xml)
        'sqlP.Value = ""
        ' Dim sql As String = "Insert into xxx values (@parm1)"
        Dim sql As String = $"INSERT INTO [LPPick].[dbo].[XmlInput] ([XmlInputData]) VALUES('{doc.InnerXml}')"
        Using sqlCmd As New System.Data.SqlClient.SqlCommand(sql, sqlConn)
            sqlCmd.ExecuteNonQuery()
        End Using

    End Sub

    Private Sub Create_P2L_Delete_xml(SHIP_BOL_NO As String)

        Dim xmlString As New System.Text.StringBuilder
        Stop

    End Sub

    Sub Poll_P2L()

        Dim XML As String = "<PickMade EventDateTime='2011-03-02 08:45:19' EventVersion='3' OpenLineCount='0' Source='PTL'><Area AreaName='Area 1'/><Bay BayName='Bay 1'/><Box BoxId='2' BoxBarCode='00000001'/><BoxLine BoxLineId='2' Qty='1' PickTime='2011-03-02 08:45:19' IsCasePick='0' CartNumber=''/><Picker PickerId='1' PickerName='Luke' PickerBarCode='EMP01'/><PickLine PickLineId='2' LocationName='01-01-A' LocationBarCode='' ProductName='' ProductBarCode='' ProductDescription='' ProductInnerPackQty='1' PickOrderQty='1' PickedQty='1' PickLineSeqNo='0' PickLineStatus='Picked' DisplayAttribute=''/><PickOrder PickOrderId='2' BatchNumber='' PickOrderNumber='00000001' PickOrderBarCode='00000001' PickTicketNumber='' PickTicketBarCode='' PickOrderStatus='Normal' OrderType='otPtl'/><WorkPlan WorkPlanName='1 Picker'/><Zone ZoneName='Zone 1'/></PickMade>"

        Load_PickMade(XML)
        Exit Sub

        Dim sqlConn As New System.Data.SqlClient.SqlConnection(sqlCS)
        sqlConn.Open()
        Dim sql As String = "Select [XmlOutputId], [XmlOutputData] FROM [XmlOutput]" & vbCrLf _
            & " where [XmlOutputProcessed] = 0 ORDER BY [XmlOutputId] ASC"
        Dim sqlCmd As New System.Data.SqlClient.SqlCommand(sql, sqlConn)
        Dim tbl As New DataTable

        Using dr As System.Data.SqlClient.SqlDataReader = sqlCmd.ExecuteReader()

            Do While dr.Read
                Dim XmlOutputId As String = dr("XmlOutputId")
                Dim XmlOutputTime As Date = dr("XmlOutputTime")
                Dim XmlOutputData As String = dr("XmlOutputData")

                BeginTrans()

                Try

                    ASCMAIN1.sql = $"Insert into WHTP2LX1 Values ({XmlOutputId},'{XmlOutputTime}',0,NULL,'{XmlOutputData}')"
                    ASCDATA1.ExecuteSQL()

                    If XmlOutputData.StartsWith("<PickMade ") Then

                    End If

                    CommitTrans()

                    sql = "Update [XmlOutput] SET [XmlOutputProcessed] = 1, [XmlOutputProcessedTime] = GETDATE()" & vbCrLf _
                        & $" where [XmlOutputId] = {XmlOutputId}"
                    Dim sqlCmd2 As New System.Data.SqlClient.SqlCommand(sql, sqlConn)
                    sqlCmd2.ExecuteNonQuery()

                Catch ex As Exception

                    MsgBox(ex.InnerException.Message, MsgBoxStyle.OkOnly, "Error Occurred")
                    Rollback()

                End Try

            Loop

            '.Close()
            '.Dispose()
        End Using


        'The Xml data from the Lighting Pick database table named XmlOutput can be read by the customer.
        'Select [XmlOutputId], [XmlOutputData] FROM [LPPick].[dbo].[XmlOutput]
        'WHERE [XmlOutputProcessed] = 0 ORDER BY [XmlOutputId] ASC

        'When finished processing the record, run the following statement
        'Update [LPPick].[dbo].[XmlOutput] 
        'SET [XmlOutputProcessed] = 1, [XmlOutputProcessedTime] = GETDATE()
        'WHERE [XmlOutputId] = (XmlOutputId from above)


    End Sub

    Sub Load_PickMade(XML As String)
        Dim doc As New System.Xml.XmlDocument()
        doc.LoadXml(XML.ToString)

        Dim EVENTDATETIME As Date = Now
        Dim BOXBARCODE As String = ""
        Dim PICKTIME As String = ""
        Dim PICKERID As String = ""
        Dim PICKORDERBARCODE As String = ""

        Dim PICKMADE As String = ASCMAIN1.Next_Control_No("WHTP2LP1.PICKMADE")

        Dim elem As System.Xml.XmlElement = Nothing

        Dim elems As New Dictionary(Of String, System.Xml.XmlNodeList)
        For Each c As String In New String() {"Box", "BoxLine", "Picker", "PickLine", "PickOrder"}
            elems.Add(c, doc.DocumentElement.GetElementsByTagName(c))
        Next


        Dim rowWHTP2LP1 As DataRow = dst.Tables("WHTP2LP1").NewRow
        With rowWHTP2LP1
            .Item("PICKMADE") = PICKMADE
            '.Item("EVENTDATETIME") = CDate(doc.DocumentElement.GetAttribute("EventDateTime"))
            .Item("EVENTDATETIME") = CDate(doc.DocumentElement.Attributes("EventDateTime").Value)
            elem = doc.DocumentElement



            elem = doc.DocumentElement.GetElementsByTagName("Box")(0)
            .Item("BOXBARCODE") = elem.GetAttribute("BoxBarCode")
            .Item("BOXBARCODE") = elems("Box")(0).Attributes("BoxBarCode").Value

            .Item("PICKTIME") = CDate(elems("BoxLine")(0).Attributes("PickTime").Value)
            .Item("PICKERID") = elems("Picker")(0).Attributes("PickerId").Value
            .Item("PICKORDERBARCODE") = elems("PickOrder")(0).Attributes("PickOrderBarCode").Value
        End With
        dst.Tables("WHTP2LP1").Rows.Add(rowWHTP2LP1)

        Dim PICKMADE_LNO As Int32 = 0



        For Each elem In elems("PickLine")

            Dim LOCATIONBARCODE As String = elem.Attributes("LocationBarCode").Value
            Dim PRODUCTBARCODE As String = elem.Attributes("ProductBarCode").Value
            Dim PICKORDERQTY As String = elem.Attributes("PickOrderQty").Value

            Dim rowWHTP2LP2 As DataRow = dst.Tables("WHTP2LP2").NewRow
            With rowWHTP2LP2
                .Item("PICKMADE") = PICKMADE
                PICKMADE_LNO += 1
                .Item("PICKMADE_LNO") = PICKMADE_LNO
                .Item("LOCATIONBARCODE") = LOCATIONBARCODE
                .Item("PRODUCTBARCODE") = PRODUCTBARCODE
                .Item("PICKORDERQTY") = PICKORDERQTY
            End With
            dst.Tables("WHTP2LP2").Rows.Add(rowWHTP2LP2)
        Next


    End Sub

    Sub Import_Picks()

        sqlCS = "Data Source= ABSSVR2019; Initial Catalog=test; User Id= sa; Password= 0ff1c3ABS"

        sqlCS = "Data Source= ABSSVR2019; Initial Catalog=test; User Id= test; Password= test"
        sqlCS = "Data Source= SVR-VDI-NJ-PK1; Initial Catalog=LPPick; User Id= abs; Password= v4n$4L3"
        sqlCS = "Data Source= ABSSVR2019; Initial Catalog=LPPick; User Id= abs; Password= v4n$4L3"

        Dim sqlConn As New System.Data.SqlClient.SqlConnection(sqlCS)
        sqlConn.Open()

    End Sub

    Private Sub grdWHTWAVE3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTWAVE3.AfterRowActivate
        If grdWHTWAVE3.ActiveRow Is Nothing OrElse Not grdWHTWAVE3.ActiveRow.IsDataRow Then
            grdWHTWAVEC.Visible = False
        Else
            grdWHTWAVEC.Visible = True
            Dim SHIP_BOL_NO As String = grdWHTWAVE3.ActiveRow.Cells("SHIP_BOL_NO").Value
            Dim SHIP_ADDR_CODE As String = grdWHTWAVE3.ActiveRow.Cells("SHIP_ADDR_CODE").Value
            grdWHTWAVEC.Text = $"Cartons in Shipment {SHIP_BOL_NO} - DC {SHIP_ADDR_CODE}"

            Dim dvw As DataView = DirectCast(grdWHTWAVEC.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"SHIP_BOL_NO = '{SHIP_BOL_NO}'"
            Sort_grdColumns(grdWHTWAVEC, "CART_NO")
        End If
    End Sub
End Class