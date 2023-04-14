Imports System.Drawing
Imports System.Math
Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinTabControl

Public Class WHFCYCLA
    Dim rowICTIADJ1 As DataRow
    Dim location_support As Boolean = False
    Dim rowICTWHSE1 As DataRow
    Dim tblADJ_REF As String = String.Empty
    Dim CYCLE_STATUS As String = ""
    Dim CYCLE_TYPE As String = ""
    Dim WHSE_TRAN_NO As String = ""
    Dim WHSE_CODE As String = ""


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFIADJI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        With dst
            'ASCMAIN1.sql = "Select ICTIADJ1.*" _
            '& " from ICTIADJ1 where ICTIADJ1.OPS_YYYYPP = :PARM1"
            'Create_TDA(.Tables.Add, "ICTIADJX", "**", 0, False, "V")

            'ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC" _
            '& ", ICTIADJ1.ADJ_DATE, ICTIADJ1.WHSE_CODE, ICTIADJ1.REASON_CODE" _
            '& ", ICTIADJ1.ADJ_NOTE, ICTIADJ1.INIT_OPER, ICTIADJ1.INIT_DATE" _
            '& ", ICTIADJ1.ADJ_SOURCE, ICTIADJ1.OPS_YYYYPP, ICTIADJ1.RTRN_NO" _
            '& " from ICTIADJ1,ICTIADJ3,GLTACCT1 where ICTIADJ1.OPS_YYYYPP = :PARM1" _
            '& " and GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE" _
            '& " and ICTIADJ3.ADJ_NO = ICTIADJ1.ADJ_NO"
            'Create_TDA(.Tables.Add, "ICTIADJG", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "ICTIADJ1", "*")

            ASCMAIN1.sql = "Select ICTIADJ2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from ICTIADJ2,ICTSTYL1,ICTCOLR1 where ICTSTYL1.STYLE_CODE = ICTIADJ2.STYLE_CODE" _
            & " and ICTCOLR1.COLOR_CODE = ICTIADJ2.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTIADJ2", "**", 1)
            .Tables("ICTIADJ2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(ADJ_QTY,0) * ISNULL(STYLE_COST,0)")

            ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC" _
            & " from ICTIADJ3,GLTACCT1 where GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE"
            Create_TDA(.Tables.Add, "ICTIADJ3", "**", 1)

            ASCMAIN1.sql = "SELECT DISTINCT WHTCYCL1.* FROM WHTCYCL2,WHTCYCL1" _
            & " WHERE WHTCYCL1.CYCLE_NO = WHTCYCL2.CYCLE_NO AND NVL(WHTCYCL1.UPDATED_INV_ADJ,0) <> 'X'" _
            & " AND WHTCYCL1.CYCLE_TYPE ='V' AND WHTCYCL1.CYCLE_STATUS = 'D'" _
            & " And WHTCYCL1.INIT_DATE >= :PARM1"
            Create_TDA(.Tables.Add, "WHTCYCL1", "**", 0, True, "D")
            .Tables("WHTCYCL1").Columns.Add("SEL")
            .Tables("WHTCYCL1").Columns("SEL").DefaultValue = "0"

            ' new Closed Carton SQL

            ASCMAIN1.sql = "Select distinct WHTPHYC4.* From WHTPHYC4, WHTPHYC5" _
            & " WHERE WHTPHYC4.WHSE_CODE = WHTPHYC5.WHSE_CODE" _
            & " And WHTPHYC4.seq_no = WHTPHYC5.seq_no" _
            & " And whtphyc5.location_qty <> whtphyc5.qty_count" _
            & " And WHTPHYC4.INIT_DATE >= :PARM1"
            Create_TDA(.Tables.Add, "WHTPHYC4", "**", 0, True, "D")
            .Tables("WHTPHYC4").Columns.Add("SEL")
            .Tables("WHTPHYC4").Columns("SEL").DefaultValue = "0"


            ASCMAIN1.sql = "SELECT * FROM WHTLOCB2 WHERE WHSE_TRAN_TYPE= 'M'" _
                & " AND WHSE_TRAN_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTLOCB2", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "WHTCYCLC", "*", 0)

        End With

        Dim CUTOFF As Date = Nothing

        dteDATE_CUTOFF.Value = "06-FEB-2023"
        CUTOFF = dteDATE_CUTOFF.Value
        OptResolution.Value = "U"
        Dim CYCLE_RESOLUTION As String = "U"
        Dim CYCLE_TYPE As String = "V"


        Fill_Records("WHTCYCL1", New Object() {CUTOFF})

        Fill_Records("WHTPHYC4", New Object() {CUTOFF})

        grdWHTCYCL1.DataSource = dst.Tables("WHTCYCL1")
        grdWHTPHYC4.DataSource = dst.Tables("WHTPHYC4")
        grdWHTLOCB2.DataSource = dst.Tables("WHTLOCB2")
        grdWHTCYCLC.DataSource = dst.Tables("WHTCYCLC")

        Create_Summary(grdWHTCYCL1, "CYCLE_NO", "Count")
        Create_Summary(grdWHTPHYC4, "SEQ_NO", "Count")
        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_NO", "Count")
        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_QTY")

        With grdWHTCYCL1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        With grdWHTPHYC4.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        With grdWHTCYCL1.DisplayLayout.Bands(0)
            .Columns("WHSE_TRAN_NO").CellAppearance.BackColor = Color.Beige
        End With
        With grdWHTPHYC4.DisplayLayout.Bands(0)
            .Columns("WHSE_TRAN_NO").CellAppearance.BackColor = Color.Beige
        End With


        With grdWHTLOCB2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdWHTLOCB2.DisplayLayout.Bands(0)
            .Columns("WHSE_TRAN_NO").CellAppearance.BackColor = Color.Beige
        End With

        With grdWHTCYCL1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdWHTPHYC4.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

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

            Case "Update"
                'Update_Record()
                'Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    'If ScreenMode And EntryMode <> "N" Then
                    '    .Items("Update").Settings.Enabled = not_iScreenMode
                    '    .Items("Cancel").Settings.Enabled = not_iScreenMode
                    'Else
                    '    .Items("Update").Settings.Enabled = iScreenMode
                    '    .Items("Cancel").Settings.Enabled = iScreenMode
                    'End If


                End With

                .Groups("Screen Control").Visible = False


                If tab0.Tabs(0).Selected = True Then
                    '   .Groups("Screen Control").Visible = True
                    If (ASCMAIN1.USER_ID = "dgj" Or ASCMAIN1.USER_ID = "jimmie" Or ASCMAIN1.USER_ID = "melvin") Then
                        .Groups("Cycle Count Adjustment").Visible = True
                    Else
                        .Groups("Cycle Count Adjustment").Visible = False
                    End If
                Else
                    '   .Groups("Screen Control").Visible = False
                    .Groups("Cycle Count Adjustment").Visible = False
                End If


            End With
        End If
        If (ASCMAIN1.USER_ID = "dgj" Or ASCMAIN1.USER_ID = "jimmie" Or ASCMAIN1.USER_ID = "melvin") Then
            tab0.Tabs("Cycle Count Adjustments").Visible = True
        Else
            tab0.Tabs("Cycle Count Adjustments").Visible = False
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        'For Each TABLE_NAME As String In New String() {"ICTIADJ0", "ICTIADJ1", "ICTIADJ2", "ICTIADJ3"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next
        EnforceConstraints(True)

        CYCLE_TYPE = "V"
        CYCLE_STATUS = "D"
        OptResolution.Value = "A"
        OptResolution.Value = "U"
        optTYPE.Value = "U"
        grdWHTPHYC4.Visible = False


        Refresh_Cycle_Counts()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        'BeginTrans()

        'ICCMAIN1.Update_Adjustment(Me)

        'If location_support Then

        '    ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
        '             New Object() {"A", rowICTIADJ1.Item("ADJ_NO"), ASCMAIN1.SESSION_NO},
        '             New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
        'End If
        'CommitTrans("Update Complete")

    End Sub

    Sub Update_WHTLOCBX()


        Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").Rows(0)
        For Each row As DataRow In dst.Tables("ICTIADJ2").Select("")
            Dim TRAN_NO As String = row.Item("ADJ_NO")
            Dim TRAN_LNO As Integer = row.Item("ADJ_LNO")
            Dim WHSE_CODE As String = rowICTIADJ1.Item("WHSE_CODE")
            Dim BAR_CODE As String = "0000000000" ' row.Item("BAR_CODE")
            Dim LOCATION_CODE As String = row.Item("LOCATION_CODE")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim ADJ_QTY As Int64 = Val(row.Item("ADJ_QTY") & "")

            Dim rowWHTLOCB1 As DataRow = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE})
            If rowWHTLOCB1 Is Nothing Then
                Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE}, False)
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE})
            End If

            If rowWHTLOCB1 Is Nothing Then
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").NewRow
                With rowWHTLOCB1
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOCATION_CODE") = LOCATION_CODE
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("LOCATION_QTY") = ADJ_QTY
                End With
                dst.Tables("WHTLOCB1").Rows.Add(rowWHTLOCB1)
            Else
                rowWHTLOCB1.Item("LOCATION_QTY") = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "") + ADJ_QTY
            End If

            Dim rowWHTLOCB2 As DataRow = dst.Tables("WHTLOCB2").NewRow
            With rowWHTLOCB2
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("LOCATION_CODE") = LOCATION_CODE
                .Item("BAR_CODE") = BAR_CODE
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("WHSE_TRAN_QTY") = ADJ_QTY
                .Item("WHSE_TRAN_TYPE") = "A"
                .Item("WHSE_TRAN_NO") = TRAN_NO
                .Item("WHSE_TRAN_LNO") = TRAN_LNO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LOCATION_CODE_OTHER") = ""
                .Item("SESSION_ID") = ""
            End With
            dst.Tables("WHTLOCB2").Rows.Add(rowWHTLOCB2)
        Next

        Update_Record_TDA("WHTLOCB1")
        Update_Record_TDA("WHTLOCB2")

        dst.Tables("WHTLOCB1").Rows.Clear()
        dst.Tables("WHTLOCB2").Rows.Clear()
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub


    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("ADJ_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTCYCL1, "B", "Inventory Adjustment Inquiry")
        Load_Popup_Menu(grdWHTPHYC4, "B", "Inventory Adjustment Inquiry")
        Load_Popup_Menu(grdWHTLOCB2, "SS", "Show Filter", "Show GroupBox")
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

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Inventory Adjustment Inquiry"
                Dim ADJ_NO As String = grd.ActiveRow.Cells("ADJ_NO").Text
                Dim rowICTIADJ1 As DataRow = LookUp("ICTIADJ1", ADJ_NO)
                If rowICTIADJ1 IsNot Nothing Then
                    Context_Launch("Select", ADJ_NO, e.Tool.Key, "ICFIADJI")
                End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "WHSE_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            If Not InquiryMode Then
        '                Click_Command("New", e)
        '            End If
        '        End If
        '    Case "ADJ_NO"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Click_Command("View", e)
        '        End If
        'End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "WHSE_CODE"
        '        If Not InquiryMode Then
        '            Click_Command("New")
        '        End If
        '    Case "ADJ_NO"
        '        Click_Command("View")
        'End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region




#Region "grdWHTCYCLC"

    Private Sub grdWHTCYCLC_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTCYCLC.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

                grdCodeDesc(grdWHTCYCLC, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE As String = e.Cell.Value
                    'e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                Else
                    grdWHTCYCLC.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "COLOR_CODE"
                grdCodeDesc(grdWHTCYCLC, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE COLOR_DESC
                If cdr IsNot Nothing Then
                    'e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")
                End If

        End Select
    End Sub


    Private Sub grdWHTCYCLC_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTCYCLC.AfterRowActivate
        With grdWHTCYCLC.DisplayLayout.Bands(0)
            If grdWHTCYCLC.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdWHTCYCLC.ActiveCell = grdWHTCYCLC.ActiveRow.Cells("STYLE_CODE")
                grdWHTCYCLC.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

    End Sub


    Private Sub grdWHTCYCLC_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWHTCYCLC.BeforeExitEditMode
        If grdWHTCYCLC.ActiveCell Is Nothing Then Exit Sub
        With grdWHTCYCLC.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTSTYL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If

                Case "COLOR_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTCOLR1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Color Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                        If Not e.Cancel Then
                            cdr = LookUp("ICTSTYC1", New String() { .Row.Cells("STYLE_CODE").Value, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Color Code (" & .Text & ") not set up with Style (" & .Row.Cells("STYLE_CODE").Value & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdWHTCYCLC_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTCYCLC.BeforeRowUpdate
        With grdWHTCYCLC
            If e.Row.Cells("STYLE_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Row.Cells("COLOR_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTCOLR1", e.Row.Cells("COLOR_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Color Code (" & e.Row.Cells("COLOR_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
                If Not e.Cancel Then
                    LookUp("ICTSTYC1", New String() {e.Row.Cells("STYLE_CODE").Text, e.Row.Cells("COLOR_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Color Code (" & e.Row.Cells("COLOR_CODE").Text & ") not set up for Style (" & e.Row.Cells("STYLE_CODE").Text & ")",
                               MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If


            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("WHSE_CODE").Text = "" Then
                    .ActiveRow.Cells("WHSE_CODE").Value = "NJC"
                    .ActiveRow.Cells("INIT_DATE").Value = Now
                    .ActiveRow.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                Else
                    .ActiveRow.Cells("LAST_DATE").Value = Now
                    .ActiveRow.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                End If
            End If
        End With
    End Sub

    Private Sub grdWHTCYCLC_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTCYCLC.ClickCellButton

        If grdWHTCYCLC.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

            Case "COLOR_CODE"
                sql_where = "COLOR_CODE in (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & e.Cell.Row.Cells("STYLE_CODE").Value & "')"

        End Select
        grdClickCellButton(grdWHTCYCLC, sql_where, False)

    End Sub

    Private Sub grdWHTCYCLC_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdWHTCYCLC.Error
        grdWHTCYCLC.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show("Are you sure you want to Refresh based on Date, You will lose any Cycles you have Selected?", "Confirm Refresh",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Refresh_Cycle_Counts()

    End Sub

    Sub Refresh_Cycle_Counts()
        Dim CUTOFF As Date = Nothing
        CUTOFF = dteDATE_CUTOFF.Value
        Fill_Records("WHTCYCL1", CUTOFF)
        OptResolution.Value = "p"
        OptResolution.Value = "U"

        CYCLE_TYPE = "V"
        CYCLE_STATUS = "D"
        OptResolution.Value = "U"
        chkUpdated.Checked = False

        Fill_Records("WHTCYCLC")

        Sort_grdColumns(grdWHTCYCLC, "STYLE_CODE, COLOR_CODE")

    End Sub
    Private Sub cmdUpdateCycles_Click(sender As Object, e As EventArgs) Handles cmdUpdateCycles.Click

        If optTYPE.Value = "U" Then
            If DirectCast(grdWHTCYCL1.DataSource, DataTable).Select("SEL='1'").Length = 0 Then
                EMsg = "You Must Select a Cycle to Update"
                MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
                Exit Sub
            End If
        Else
            If DirectCast(grdWHTPHYC4.DataSource, DataTable).Select("SEL='1'").Length = 0 Then
                EMsg = "You Must Select a Closed Carton to Update"
                MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
                Exit Sub
            End If


        End If
        If optTYPE.Value = "U" Then
            If MessageBox.Show("Are you sure you want to Update Cycles Selected?", "Confirm Update",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            Call UPDATE_CYCLE_ADJUSTMENTS()
            Me.Cursor = Cursors.Default

        Else
            If MessageBox.Show("Are you sure you want to Update Closed Cartons Selected?", "Confirm Update",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            Call UPDATE_CLOSED_CARTON_ADJUSTMENTS()
            Me.Cursor = Cursors.Default

        End If


    End Sub

    Sub Load_WHTCYCL1()

        If optTYPE.Value = "C" Then
            Dim SQLW As String = "SEQ_NO = SEQ_NO"
            grdWHTCYCL1.Visible = False

            If chkUpdated.Checked = False Then
                cmdUpdateCycles.Enabled = True
                grdWHTPHYC4.DisplayLayout.Bands(0).Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit

                ' grdWHTCYCL1.DisplayLayout.Override.AllowUpdate = False
            Else
                cmdUpdateCycles.Enabled = False
                '   grdWHTCYCL1.DisplayLayout.Override.AllowUpdate = False
                grdWHTPHYC4.DisplayLayout.Bands(0).Columns("SEL").CellActivation = UltraWinGrid.Activation.NoEdit
            End If


            If chkUpdated.Checked = True Then
                grdWHTCYCL1.Text = "Closed Cartons That have been Updated"
            Else
                grdWHTCYCL1.Text = "Closed Cartons That have been Not been Updated"
            End If

            If chkUpdated.Checked = True Then
                SQLW = SQLW & " and ADJ_NO IS NOT NULL"
            Else
                SQLW = SQLW & " and ADJ_NO IS NULL"
            End If

            If chkSel.Checked = True Then
                SQLW = SQLW & " and SEL = '1'"
            End If

            Dim dvw As DataView
            dvw = DirectCast(grdWHTPHYC4.DataSource, DataTable).DefaultView
            dvw.RowFilter = SQLW
        Else
            grdWHTCYCL1.Visible = True
            Dim SQLW As String = "CYCLE_NO = CYCLE_NO"

            If CYCLE_TYPE = "V" And OptResolution.Value = "U" And CYCLE_STATUS = "D" And chkUpdated.Checked = False Then
                cmdUpdateCycles.Enabled = True
            Else
                cmdUpdateCycles.Enabled = False
            End If

            If OptResolution.Value = "A" Then
            Else
                SQLW = SQLW & " And CYCLE_RESOLUTION = '" & OptResolution.Value & "'"
            End If

            SQLW = SQLW & " and CYCLE_TYPE = '" & CYCLE_TYPE & "'"

            SQLW = SQLW & " and CYCLE_STATUS = '" & CYCLE_STATUS & "'"

            If chkUpdated.Checked = False Then
                SQLW = SQLW & " and UPDATED_INV_ADJ IS NULL"
                grdWHTCYCL1.DisplayLayout.Bands(0).Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                SQLW = SQLW & " and UPDATED_INV_ADJ = '1'"
                grdWHTCYCL1.DisplayLayout.Bands(0).Columns("SEL").CellActivation = UltraWinGrid.Activation.NoEdit

            End If

            If chkSel.Checked = True Then
                SQLW = SQLW & " and SEL = '1'"
            End If

            Dim dvw As DataView
            dvw = DirectCast(grdWHTCYCL1.DataSource, DataTable).DefaultView
            dvw.RowFilter = SQLW

            If chkUpdated.Checked = True Then
                grdWHTCYCL1.Text = "Cycles That have been Updated"
            Else
                grdWHTCYCL1.Text = "Cycles That have been not been Updated"
            End If


        End If


    End Sub
    Sub LOAD_WHTLOCB2()
        Dim SQLW As String = "LOCATION_CODE = LOCATION_CODE"

        If chkLF.Checked = True Then
            SQLW = SQLW & " and LOCATION_CODE = '00-LNF-A'"
        End If

        Dim dvw As DataView
        dvw = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView
        dvw.RowFilter = SQLW



    End Sub

    Private Sub OptResolution_ValueChanged(sender As Object, e As EventArgs) Handles OptResolution.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()

    End Sub

    Private Sub chkUpdated_CheckedValueChanged(sender As Object, e As EventArgs) Handles chkUpdated.CheckedValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()
    End Sub
    Sub UPDATE_CYCLE_ADJUSTMENTS()
        WHSE_CODE = ""


        dst.Tables("ICTIADJ1").Rows.Clear()
        dst.Tables("ICTIADJ2").Rows.Clear()
        DATETIME_STAMP = Now + ASCMAIN1.NowTSD


        ' PRE UDPATE LOCK CYCLES, CAN'T LOCK, GET OUT
        For Each ROW As DataRow In DirectCast(grdWHTCYCL1.DataSource, DataTable).Select("SEL='1'")
            If ROW Is Nothing Then
                EMsg &= vbCr & "You Must Select at least 1 Cycle to Update, Cannot Proceed"
                Exit Sub

            End If
            If WHSE_CODE & "" = "" Then
                WHSE_CODE = ROW.Item("WHSE_CODE") & ""
            End If
            If WHSE_CODE <> ROW.Item("WHSE_CODE") & "" Then
                EMsg &= vbCr & "Multiple Wareshouses Selected, Cannot Proceed"
                Exit Sub
            End If
            Dim CYCLE_NO = ROW.Item("CYCLE_NO") & ""
            If Not ASCMAIN1.Logical_Lock("WHTCYCL1", CYCLE_NO) Then
                ASCMAIN1.MultiTask_Release(, , 1)
                Exit Sub
            End If
            ' CHECK WHTCYCL1 UPDATED FLAG HERE FOR CYCL1
            ASCMAIN1.sql = "Select * from WHTCYCL1 where CYCLE_NO = '" & CYCLE_NO & "' AND UPDATED_INV_ADJ = '1'"
            Dim tblWHTCYCL1 As DataTable = ASCDATA1.GetDataTable()
            If tblWHTCYCL1.Rows.Count > 0 Then
                EMsg &= vbCr & "Cycle " & CYCLE_NO & " Had Already Been Updated, Cannot Proceed"
                Exit Sub
            End If

        Next

        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")

        Dim WHSE_TRAN_NO As String = ""
        Dim LOCATION_CODE As String = ""
        For Each ROW As DataRow In DirectCast(grdWHTCYCL1.DataSource, DataTable).Select("SEL='1'")
            WHSE_TRAN_NO = ROW.Item("WHSE_TRAN_NO") & ""
            Dim ADJ_NO As String = ASCMAIN1.Next_Control_No("TRAN_NO_A")
            Update_Adjustment(WHSE_TRAN_NO, ADJ_NO)
            ROW.Item("UPDATED_INV_ADJ") = "1"
            ROW.Item("ADJ_NO") = ADJ_NO
        Next

        ' NEW Update
        BeginTrans()

        Update_Record_TDA("WHTCYCL1")

        ICCMAIN1.Update_Adjustment(Me)

        If location_support Then
            For Each rowICTIADJ1 As DataRow In dst.Tables("ICTIADJ1").Select("")
                ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                     New Object() {"A", rowICTIADJ1.Item("ADJ_NO"), ASCMAIN1.SESSION_NO},
                     New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

            Next
        End If
        CommitTrans("Update Complete")

        ASCMAIN1.MultiTask_Release(, , 1)

        Mode_Settings(False)

    End Sub
    Sub UPDATE_CLOSED_CARTON_ADJUSTMENTS()
        WHSE_CODE = ""


        dst.Tables("ICTIADJ1").Rows.Clear()
        dst.Tables("ICTIADJ2").Rows.Clear()
        DATETIME_STAMP = Now + ASCMAIN1.NowTSD


        ' PRE UDPATE LOCK SESSIONS WHTPHYC4, CAN'T LOCK, GET OUT
        For Each ROW As DataRow In DirectCast(grdWHTPHYC4.DataSource, DataTable).Select("SEL='1'")
            If ROW Is Nothing Then
                EMsg &= vbCr & "You Must Select at least 1 Session to Update, Cannot Proceed"
                Exit Sub

            End If
            If WHSE_CODE & "" = "" Then
                WHSE_CODE = ROW.Item("WHSE_CODE") & ""
            End If
            If WHSE_CODE <> ROW.Item("WHSE_CODE") & "" Then
                EMsg &= vbCr & "Multiple Wareshouses Selected, Cannot Proceed"
                Exit Sub
            End If
            Dim SEQ_NO = ROW.Item("SEQ_NO") & ""
            If Not ASCMAIN1.Logical_Lock("WHTPHYC4", SEQ_NO) Then
                ASCMAIN1.MultiTask_Release(, , 1)
                Exit Sub
            End If
            '' CHECK WHTPHYC4 UPDATED ADJ NO HERE FOR WHTPHYC4
            'ASCMAIN1.sql = "Select * from WHTPHYC4 where WHSE_CODE = '" & WHSE_CODE & "' and SEQ_NO = '" & SEQ_NO & "' AND NVL(WHTPHYC4.ADJ_NO,0) <> 0"
            'Dim tblWHTPHYC4 As DataTable = ASCDATA1.GetDataTable()
            'If tblWHTPHYC4.Rows.Count > 0 Then
            '    EMsg &= vbCr & "Closed Carton " & SEQ_NO & " Has Already Been Updated, Cannot Proceed"
            '    Exit Sub
            'End If


        Next
        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")

        Dim WHSE_TRAN_NO As String = ""
        Dim LOCATION_CODE As String = ""
        For Each ROW As DataRow In DirectCast(grdWHTPHYC4.DataSource, DataTable).Select("SEL='1'")

            WHSE_TRAN_NO = ROW.Item("WHSE_TRAN_NO") & ""
            Dim ADJ_NO As String = ASCMAIN1.Next_Control_No("TRAN_NO_A")
            Update_Adjustment(WHSE_TRAN_NO, ADJ_NO)
            ROW.Item("ADJ_NO") = ADJ_NO
        Next

        ' NEW Update
        BeginTrans()

        Update_Record_TDA("WHTPHYC4")

        ICCMAIN1.Update_Adjustment(Me)

        If location_support Then
            For Each rowICTIADJ1 As DataRow In dst.Tables("ICTIADJ1").Select("")
                ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                     New Object() {"A", rowICTIADJ1.Item("ADJ_NO"), ASCMAIN1.SESSION_NO},
                     New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

            Next
        End If
        CommitTrans("Update Complete")

        ASCMAIN1.MultiTask_Release(, , 1)

        Mode_Settings(False)
    End Sub
    Private Sub chkSel_CheckedValueChanged(sender As Object, e As EventArgs) Handles chkSel.CheckedValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()

    End Sub

    Private Sub grdWHTCYCL1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTCYCL1.AfterRowActivate
        Dim WHSE_TRAN_NO As String = grdWHTCYCL1.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""

        Fill_Records("WHTLOCB2", WHSE_TRAN_NO)

        grdWHTLOCB2.Text = "LOCB2 Records for Whse Transaction No" & " " & grdWHTCYCL1.ActiveRow.Cells("WHSE_TRAN_NO").Value

        Dim SQLW As String = "WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'"
        If chkLF.Checked = True Then
            SQLW = SQLW & " and LOCATION_CODE = '00-LNF-A'"
        End If

        Dim dvw As DataView = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView
        dvw.RowFilter = SQLW

        SET_LOCB2(WHSE_TRAN_NO)

    End Sub

    Private Sub grdWHTCYCLC_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdWHTCYCLC.AfterRowUpdate
        Update_Record_TDA("WHTCYCLC")
    End Sub

    Private Sub tab0_Click(sender As Object, e As EventArgs) Handles tab0.Click
        With UltraExplorerBar1

            If tab0.Tabs(0).Selected = True Then
                .Groups("Screen Control").Visible = True
                .Groups("Cycle Count Adjustment").Visible = True
            Else
                .Groups("Screen Control").Visible = False
                .Groups("Cycle Count Adjustment").Visible = False

            End If

        End With


    End Sub

    Private Sub optTYPE_ValueChanged(sender As Object, e As EventArgs) Handles optTYPE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        '    Load_WHTCYCL1()
        Dim dvw As DataView = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "WHSE_CODE = '999'"


        If optTYPE.Value = "C" Then
            UltraLabel7.Visible = False
            OptResolution.Visible = False
            dteDATE_CUTOFF.Visible = False
            UltraLabel3.Visible = False
            Button2.Visible = False
            chkUpdated.Text = "Updated Closed Cartons"
            grdWHTCYCL1.Visible = False
            grdWHTPHYC4.Visible = True


        Else
            UltraLabel7.Visible = True
            OptResolution.Visible = True
            dteDATE_CUTOFF.Visible = True
            UltraLabel3.Visible = True
            Button2.Visible = True
            chkUpdated.Text = "Updated Cycle Counts"
            grdWHTPHYC4.Visible = False
            grdWHTCYCL1.Visible = True
            OptResolution.Value = "A"
            OptResolution.Value = "U"


        End If

        Load_WHTCYCL1()
        OptResolution.Value = "A"
        OptResolution.Value = "U"


    End Sub

    Private Sub chkUpdated_CheckedChanged(sender As Object, e As EventArgs) Handles chkUpdated.CheckedChanged

    End Sub

    Private Sub grdWHTPHYC4_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTPHYC4.AfterRowActivate
        Dim WHSE_TRAN_NO As String = grdWHTPHYC4.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""

        Fill_Records("WHTLOCB2", WHSE_TRAN_NO)

        grdWHTLOCB2.Text = "LOCB2 Records for Whse Transaction No" & " " & grdWHTPHYC4.ActiveRow.Cells("WHSE_TRAN_NO").Value

        SET_LOCB2(WHSE_TRAN_NO)
    End Sub
    Sub SET_LOCB2(WHSE_TRAN_NO As String)

        Dim SQLW As String = "WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'"
        If chkLF.Checked = True Then
            SQLW = SQLW & " and LOCATION_CODE = '00-LNF-A'"
        End If

        Dim dvw As DataView = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView
        dvw.RowFilter = SQLW
    End Sub
    Private Sub chkLF_CheckedValueChanged(sender As Object, e As EventArgs) Handles chkLF.CheckedValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        LOAD_WHTLOCB2()

    End Sub
    Sub Update_Adjustment(WHSE_TRAN_NO As String, ADJ_NO As String)
        Dim LOCATION_CODE As String = "00-LNF-A"
        Fill_Records("WHTLOCB2", WHSE_TRAN_NO)

        rowICTIADJ1 = dst.Tables("ICTIADJ1").NewRow
        rowICTIADJ1.Item("ADJ_NO") = ADJ_NO
        rowICTIADJ1.Item("WHSE_CODE") = WHSE_CODE
        rowICTIADJ1.Item("ADJ_DATE") = DATETIME_STAMP.Date
        rowICTIADJ1.Item("ADJ_SOURCE") = "E"
        rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("INIT_DATE") = DATETIME_STAMP
        rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("LAST_DATE") = DATETIME_STAMP
        rowICTIADJ1.Item("REGISTER_IND") = "0"
        rowICTIADJ1.Item("JOURNAL_IND") = "0"
        rowICTIADJ1.Item("REASON_CODE") = "WHLOC"
        rowICTIADJ1.Item("TOTAL_COSTS") = 0

        dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)


        Dim rowICTIADJ2 As DataRow
        For Each rowWHTLOCB2 As DataRow In dst.Tables("WHTLOCB2").Select("WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "' and LOCATION_CODE = '" & LOCATION_CODE & "'")
            cdr = LookUp("ICTSTYL1", rowWHTLOCB2("STYLE_CODE"))
            Dim STYLE_DESC As String = cdr.Item("STYLE_DESC") & ""
            Dim STYLE_CLASS_CODE As String = cdr.Item("STYLE_CLASS_CODE") & ""
            Dim SALES_DIVISION_CODE As String = cdr.Item("SALES_DIVISION_CODE") & ""
            Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")

            cdr = LookUp("ICTCOLR1", rowWHTLOCB2("COLOR_CODE"))
            Dim COLOR_DESC As String = cdr.Item("COLOR_DESC") & ""


            rowICTIADJ2 = dst.Tables("ICTIADJ2").NewRow
            With rowICTIADJ2
                .Item("ADJ_NO") = ADJ_NO
                .Item("ADJ_LNO") = Val(dst.Tables("ICTIADJ2").Compute("Max(ADJ_LNO)", "") & "") + 1
                .Item("STYLE_CODE") = rowWHTLOCB2("STYLE_CODE")
                .Item("STYLE_DESC") = STYLE_DESC
                .Item("COLOR_CODE") = rowWHTLOCB2("COLOR_CODE")
                .Item("COLOR_DESC") = COLOR_DESC
                .Item("ADJ_QTY") = Val(rowWHTLOCB2("WHSE_TRAN_QTY") & "") * -1
                .Item("STYLE_COST") = STYLE_COST
                ''    '.Item("STYLE_COST") = Val(row("STYLE_COST") & "") ' TEMP
                .Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("LOCATION_CODE") = rowWHTLOCB2("LOCATION_CODE")
                .Item("BAR_CODE") = rowWHTLOCB2("BAR_CODE")
                .Item("ADJ_REF") = ""
                .Item("ADJ_REF") = rowWHTLOCB2("WHSE_TRAN_NO")
            End With
            dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
        Next

    End Sub

    Private Sub grdWHTPHYC4_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdWHTPHYC4.InitializeLayout

    End Sub
End Class