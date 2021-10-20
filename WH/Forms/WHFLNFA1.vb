Imports System.Drawing
Imports System.Drawing.Printing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFLNFA1
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

    Dim WHTLOCBX As String
    Dim WHTLOCBC As String
    Dim WHTLOCLX As String


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Create_Temp_Table(True)

        If MENU_ITEM_OBJECT = "WHFPACKI" Then
            InquiryMode = True
        End If

        Get_PARM("SOTPARM1")
        With dst

            ASCMAIN1.sql = "Select * from " & WHTLOCBX
            Create_TDA(.Tables.Add, "WHTLOCBX", "**", 0, False, "", 2)


            ASCMAIN1.sql = "Select X.*, Y.DATE_LAST_COUNTED from (" & vbCrLf _
                & "Select WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & ", SUM (WHTLOCB1.LOCATION_QTY) LOCATION_QTY, SUM (WHTLOCB1.LOCATION_QTY_WAVE) LOCATION_QTY_WAVE" & vbCrLf _
                & ", MIN (WHTLOCB1.INIT_DATE) INIT_DATE, MAX (WHTLOCB1.LAST_DATE) LAST_DATE, COUNT (DISTINCT WHTLOCB1.BAR_CODE) CARTONS" & vbCrLf _
                & " from WHTLOCB1 " & vbCrLf _
                & " where WHTLOCB1.WHSE_CODE = :PARM1 And WHTLOCB1.STYLE_CODE = :PARM2 and WHTLOCB1.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and (WHTLOCB1.LOCATION_QTY <> 0 OR WHTLOCB1.LOCATION_QTY_WAVE <> 0)" & vbCrLf _
                & " group by WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_CODE) X, " & vbCrLf _
                & " (SELECT LOCATION_CODE, MAX(INIT_DATE) DATE_LAST_COUNTED FROM WHTCYCL4 WHERE WHSE_CODE = :PARM1 AND STYLE_CODE = :PARM2 AND COLOR_CODE = :PARM3 GROUP BY LOCATION_CODE) Y" & vbCrLf _
                & " where Y.LOCATION_CODE (+) = X.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCBY", "**", 0, False, "VVV", 4)

            ASCMAIN1.sql = "Select WHTLOCB2.*" & vbCrLf _
                & " from WHTLOCB2" & vbCrLf _
                & " where WHTLOCB2.WHSE_CODE = :PARM1 And WHTLOCB2.STYLE_CODE = :PARM2 and WHTLOCB2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and WHTLOCB2.LOCATION_CODE = :PARM4"
            Create_TDA(.Tables.Add, "WHTLOCBZ", "**", 0, False, "VVVV", 0)

            ASCMAIN1.sql = "Select * from " & WHTLOCLX
            Create_TDA(.Tables.Add, "WHTLOCLX", "**", 0, False, "", 1)


            ASCMAIN1.sql = "Select WHTCYCL4.* from WHTCYCL4" & vbCrLf _
                & "  where WHTCYCL4.WHSE_CODE = :PARM1 and WHTCYCL4.LOCATION_CODE = :PARM2"
            Create_TDA(.Tables.Add, "WHTLOCLY", "**", 0, False, "VV", 0)


            ASCMAIN1.sql = "Select ICTIADJ2.*, ICTIADJ1.ADJ_DATE, ICTIADJ1.REASON_CODE, ICTIADJ1.ADJ_NOTE" & vbCrLf _
                & " from ICTIADJ1, ICTIADJ2" & vbCrLf _
                & " where ICTIADJ1.WHSE_CODE = :PARM1 And ICTIADJ2.STYLE_CODE = :PARM2 and ICTIADJ2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and ICTIADJ2.LOCATION_CODE = :PARM4" & vbCrLf _
                & "   and ICTIADJ1.ADJ_NO = ICTIADJ2.ADJ_NO"
            Create_TDA(.Tables.Add, "ICTIADJ2", "**", 0, False, "VVVV", 0)

            dteAdjFrom.DateTime = DateAdd(DateInterval.Day, -10, DateTime.Now)
            dteAdjTo.DateTime = DateTime.Now


            ASCMAIN1.sql = "Select ICTIADJ2.*, ICTIADJ1.ADJ_DATE, ICTIADJ1.REASON_CODE, ICTIADJ1.ADJ_NOTE" & vbCrLf _
                & " from ICTIADJ1, ICTIADJ2" & vbCrLf _
                & " where ICTIADJ1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and ICTIADJ1.ADJ_NO = ICTIADJ2.ADJ_NO"
            Create_TDA(.Tables.Add, "ICTIADJX", "**", 0, False, "V", 0)
            '    & "   and ICTIADJ1.INIT_DATE >= '" & Format(dteAdjFrom, "dd-MMM-yyyy") & "'" _
            '    & "   and ICTIADJ1.INIT_DATE <=  '" & Format(dteAdjTo, "dd-MMM-yyyy") & "'"





            ASCMAIN1.sql = "Select ICTWHSE1.* from ICTWHSE1" & vbCrLf _
                & "  where ICTWHSE1.WHSE_STATUS = 'A' and ICTWHSE1.WHSE_LOCATOR = '1'"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "ICTWHSE1", "*",, False)

            ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                & " where WHSE_CODE = :PARM1 And LOCATION_CODE = : PARM2 And STYLE_CODE = : PARM3 And COLOR_CODE = : PARM4" & vbCrLf _
                & "   And LOCATION_QTY <> 0"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, True, "VVVV")

            ASCMAIN1.sql = "Select * from WHTLOCB2 where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2 and STYLE_CODE = :PARM3 and COLOR_CODE = :PARM4"
            Create_TDA(.Tables.Add, "WHTLOCB2", "**", 0, True, "VVVV")
        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdWHTLOCBX.DataSource = dst.Tables("WHTLOCBX")
        grdWHTLOCBY.DataSource = dst.Tables("WHTLOCBY")
        grdWHTLOCBZ.DataSource = dst.Tables("WHTLOCBZ")
        grdICTIADJ2.DataSource = dst.Tables("ICTIADJ2")
        grdICTIADJX.DataSource = dst.Tables("ICTIADJX")

        grdWHTLOCLX.DataSource = dst.Tables("WHTLOCLX")
        grdWHTLOCLY.DataSource = dst.Tables("WHTLOCLY")

        Fill_Records("ICTWHSEX")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")

        Create_Summary(grdWHTLOCLX, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCLY, "INIT_DATE", "Count")

        Create_Summary(grdWHTLOCBX, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCBX, New String() {"ONH", "WAV", "RECA", "RECB", "GUN", "LNF", "SHP", "RTN", "FIN", "ADJ", "LOC"})

        Create_Summary(grdWHTLOCBY, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCBY, New String() {"LOCATION_QTY", "LOCATION_QTY_WAVE"})

        Create_Summary(grdWHTLOCBZ, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCBZ, New String() {"WHSE_TRAN_QTY"})

        Create_Summary(grdICTIADJ2, "ADJ_NO", "Count")
        Create_Summary(grdICTIADJ2, New String() {"ADJ_QTY"})

        Create_Summary(grdICTIADJX, "ADJ_NO", "Count")
        Create_Summary(grdICTIADJX, New String() {"ADJ_QTY"})

        With grdWHTLOCBX.DisplayLayout.Bands("WHTLOCBX")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightGray
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    If gcol.Key = "LNF" Then
                        .BackColor2 = Drawing.Color.LightBlue
                    End If
                End With

            Next
        End With

        Show_Filter(grdWHTLOCBX)

        dteAdjFrom.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dteAdjTo.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)


        numDays.Value = 90

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load", "View"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    Else
                        If rowICTWHSE1.Item("WHSE_STATUS") & "" <> "A" Then
                            EMsg &= vbCrLf & "Invalid Value specified for Warehouse - Inactive"
                        End If
                        If rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
                            EMsg &= vbCrLf & "Invalid Value specified for Warehouse - No Locator"
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If eItemKey = "Load" Then
                        WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                        If Not ASCMAIN1.Logical_Open("WHTLNFA1", WHSE_CODE) Then Exit Sub
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

            Case "Load", "View"
                EntryMode = "L"
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
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Load").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Load").Settings.Enabled = not_iScreenMode
                    End If
                    ' .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Update").Visible = False

                    .Items("Load").Visible = Not InquiryMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        splWHTLOCBX.Visible = ScreenMode

        UltraExplorerBar1.Groups("Options").Visible = ScreenMode And (ASCMAIN1.CLIENT = "VAN")
        UltraExplorerBar1.Groups("Adjustment Date Range").Visible = ScreenMode And (ASCMAIN1.CLIENT = "VAN")



        chkEnableAdjustment.Visible = ScreenMode And (optLNF.Value = "LNF")
        Set_Read_Only_for_ctl(chkEnableAdjustment, False)

        chkPastDueCountsOnly.Visible = ScreenMode
        Set_Read_Only_for_ctl(chkPastDueCountsOnly, False)

        optLNF.Visible = ScreenMode
        Set_Read_Only_for_ctl(optLNF, False)

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTLOCBX", "WHTLOCBY", "WHTLOCBZ"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        chkEnableAdjustment.Checked = False

        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
        Setup_tab0()
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        rowICTWHSE1 = Fill_Record("ICTWHSE1", WHSE_CODE)

        Create_Temp_Table(False)

        Fill_Records("WHTLOCBX")
        Sort_grdColumns(grdWHTLOCBX, "STYLE_CODE, COLOR_CODE")
        Set_WHTLOCBX()

        Fill_Records("WHTLOCLX")
        Sort_grdColumns(grdWHTLOCLX, "LOCATION_CODE")

        Fill_Records("ICTIADJX", New String() {WHSE_CODE})
        Sort_grdColumns(grdICTIADJX, "ADJ_NO")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")
        BeginTrans()
        Update_Record_TDA("SOTCART2")
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
        Load_Popup_Menu(grdICTWHSEX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdWHTLOCBX, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Adjust")
        Load_Popup_Menu(grdWHTLOCBY, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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
                Case "grdWHTLOCBX"
                    Dim enableAdjustment As Boolean = chkEnableAdjustment.Checked
                    If grdWHTLOCBX.ActiveRow Is Nothing Or grdWHTLOCBX.ActiveCell Is Nothing Then
                        enableAdjustment = False
                    Else
                        If grdWHTLOCBX.ActiveCell.Column.Key = "LNF" Then

                        Else
                            enableAdjustment = False
                        End If
                    End If

                    tlb_pop.Tools("Adjust").SharedProps.Visible = enableAdjustment
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
            Case "Adjust"
                Dim LNF As Int64 = Val(grd.ActiveRow.Cells("LNF").Value & "")
                If LNF <> 0 Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value & ""
                    Dim COLOR_CODE As String = grd.ActiveRow.Cells("COLOR_CODE").Value & ""
                    Dim LOCATION_CODE As String = "00-LNF-A" ' grd.ActiveRow.Cells("LOCATION_CODE").Value & ""
                    Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE})
                    Adjustment(STYLE_CODE, COLOR_CODE, LOCATION_CODE)
                End If

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

    Private Sub grdWHTLOCBX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTLOCBX.AfterRowActivate
        Setup_WHTLOCBY()
    End Sub

    Sub Setup_WHTLOCBY()
        If grdWHTLOCBX.ActiveRow IsNot Nothing AndAlso grdWHTLOCBX.ActiveRow.IsDataRow AndAlso Not grdWHTLOCBX.ActiveRow.IsFilterRow Then
            Dim STYLE_CODE As String = grdWHTLOCBX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdWHTLOCBX.ActiveRow.Cells("COLOR_CODE").Value

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Locations", "")

            Fill_Records("WHTLOCBY", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            Sort_grdColumns(grdWHTLOCBY, "LOCATION_CODE")

            grdWHTLOCBY.Visible = True
            grdWHTLOCBY.Text = $"Location Details for {STYLE_CODE}-{COLOR_CODE}"

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")

        Else
            grdWHTLOCBY.Visible = False
        End If
    End Sub

    Sub SETUP_ICTIADJ2()
        If grdWHTLOCBY.ActiveRow IsNot Nothing AndAlso grdWHTLOCBY.ActiveRow.IsDataRow AndAlso Not grdWHTLOCBY.ActiveRow.IsFilterRow Then
            Dim STYLE_CODE As String = grdWHTLOCBY.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdWHTLOCBY.ActiveRow.Cells("COLOR_CODE").Value
            Dim LOCATION_CODE As String = grdWHTLOCBY.ActiveRow.Cells("LOCATION_CODE").Value

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Adjustments", "")

            Fill_Records("ICTIADJ2", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE, LOCATION_CODE})
            Sort_grdColumns(grdICTIADJ2, "ADJ_NO")

            grdICTIADJ2.Visible = True
            grdICTIADJ2.Text = $"Adjustments Details for {STYLE_CODE}-{COLOR_CODE} in Location {LOCATION_CODE}"

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")

        Else
            grdICTIADJ2.Visible = False
        End If
    End Sub

    Private Sub optLNF_ValueChanged(sender As Object, e As EventArgs) Handles optLNF.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        chkEnableAdjustment.Visible = ScreenMode And (optLNF.Value = "LNF")
        Set_WHTLOCBX()
    End Sub

    Sub Set_WHTLOCLX()

        Dim sqlwhere As String = ""
        Dim dvw As DataView = DirectCast(grdWHTLOCLX.DataSource, DataTable).DefaultView

        If chkPastDueCountsOnly.Checked Then
            Dim DATE_LAST_CYCLE_COUNT_cutoff As Date = Now.Date.AddDays(-1 * Val(numDays.Value & ""))
            sqlwhere &= $" and LAST_CYCLE_COUNT is Null or LAST_CYCLE_COUNT < '{Format(DATE_LAST_CYCLE_COUNT_cutoff, "MM/dd/yyyy")}'"
        End If

        dvw.RowFilter = Mid(sqlwhere, 5)

    End Sub

    Sub Set_WHTLOCBX()

        Dim sqlwhere As String = ""
        Dim dvw As DataView = DirectCast(grdWHTLOCBX.DataSource, DataTable).DefaultView
        If optLNF.Value = "ALL" Then
        Else
            sqlwhere &= " and LNF <> 0"
        End If

        If chkPastDueCountsOnly.Checked Then
            Dim DATE_LAST_CYCLE_COUNT_cutoff As Date = Now.Date.AddDays(-1 * Val(numDays.Value & ""))
            sqlwhere &= $" and DATE_LAST_CYCLE_COUNT is Null or DATE_LAST_CYCLE_COUNT < '{Format(DATE_LAST_CYCLE_COUNT_cutoff, "MM/dd/yyyy")}'"
        End If

        dvw.RowFilter = Mid(sqlwhere, 5)

    End Sub


    Private Sub grdWHTLOCBY_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTLOCBY.AfterRowActivate
        Setup_WHTLOCBZ()
        SETUP_ICTIADJ2()
    End Sub

    Sub Setup_WHTLOCBZ()
        If grdWHTLOCBY.ActiveRow IsNot Nothing AndAlso grdWHTLOCBY.ActiveRow.IsDataRow AndAlso Not grdWHTLOCBY.ActiveRow.IsFilterRow Then

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Adjustments", "")

            Dim STYLE_CODE As String = grdWHTLOCBY.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdWHTLOCBY.ActiveRow.Cells("COLOR_CODE").Value
            Dim LOCATION_CODE As String = grdWHTLOCBY.ActiveRow.Cells("LOCATION_CODE").Value

            Fill_Records("WHTLOCBZ", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE, LOCATION_CODE})
            Sort_grdColumns(grdWHTLOCBZ, "INIT_DATE")

            grdWHTLOCBZ.Visible = True
            grdWHTLOCBZ.Text = $"Transaction Details for {STYLE_CODE}-{COLOR_CODE} in Location {LOCATION_CODE}"

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
        Else
            dst.Tables("WHTLOCBZ").Rows.Clear()
        End If
    End Sub

    Sub Adjustment(STYLE_CODE As String, COLOR_CODE As String, LOCATION_CODE As String)

        Dim confirm_only As Boolean = True ' (movement_type = "LNF") Or (movement_type = "CMB")
        Dim dupBarcode As String = ""
        Dim LOCATION_CODE_TO As String = ""

        Using ff As New TAC.TAFLOCM1()

            Dim BAR_CODE_CMB As String = ""


            ff.confirm_only = confirm_only
            ff.movement_type = "ADJ"
            ff.rowICTWHSE1 = rowICTWHSE1
            ff.WHSE_CODE = WHSE_CODE

            ff.disableUpdate = True


            Dim BCs As New List(Of String)
            Dim LOCs As New List(Of String)

            Dim sqlw As String = $"WHSE_CODE = '{WHSE_CODE}' and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' and LOCATION_CODE = '{LOCATION_CODE}' and LOCATION_QTY <> 0 and BAR_CODE <> '{rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")}'"

            For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("")

                Dim BAR_CODE As String = rowWHTLOCB1.Item("BAR_CODE") & ""
                Dim LOAD_NO As String = "" ' row.Item("LOAD_NO") & ""

                If Val(rowWHTLOCB1.Item("LOCATION_QTY_WAVE") & "") <> 0 Then
                    MsgBox("Cannot Change or Move a Case which has been committed to a Wave", MsgBoxStyle.OkOnly, "Cannot Move")
                    Exit Sub
                End If

                Dim LOCATION_QTY As Int64 = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "")

                ff.AddItemToMove(WHSE_CODE,
                        LOCATION_CODE,
                            STYLE_CODE,
                            COLOR_CODE,
                            BAR_CODE,
                            LOAD_NO,
                            LOCATION_QTY,
                            LOCATION_CODE_TO, BAR_CODE_CMB)
            Next


            ff.ShowDialog()

            Dim WHSE_TRAN_NO As String = ff.WHSE_TRAN_NO
            If WHSE_TRAN_NO <> "" Then

                Dim sqlWHTLOCB2where As String = $"WHSE_CODE = '{WHSE_CODE}' and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' and LOCATION_CODE = '{LOCATION_CODE}' and BAR_CODE <> '{rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")}'"

                ASCMAIN1.sql = "Select * from WHTLOCB2" _
                    & sqlWHTLOCB2where _
                    & " and WHSE_TRAN_TYPE = 'M' and WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'"

                Fill_Records("WHTLOCB2", "", False, ASCMAIN1.sql)

                ASCMAIN1.Progress("Refreshing Data", "")
                'Load_WHTLOCB1()
            End If

            'Setup_grdWHTLOCB2()
            'Setup_grdWHTMOVEX()
            ASCMAIN1.Progress("", "")
        End Using
    End Sub

    Sub Create_Temp_Table(Initialize As Boolean)

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

        Dim sqlWHTLOCBC As String = "Select WHTCYCL4.STYLE_CODE, WHTCYCL4.COLOR_CODE" & vbCrLf _
            & ", Max(WHTCYCL4.INIT_DATE) DATE_LAST_CYCLE_COUNT" & vbCrLf _
            & " From WHTCYCL4" & vbCrLf _
            & $" Where WHTCYCL4.WHSE_CODE = '{WHSE_CODE}'" & vbCrLf _
            & " group by WHTCYCL4.STYLE_CODE, WHTCYCL4.COLOR_CODE" & vbCrLf

        If Initialize Then
            ASCMAIN1.sql = sqlWHTLOCBC
            WHTLOCBC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & WHTLOCBC & " Add Primary Key (STYLE_CODE,COLOR_CODE)")
        Else
            ASCDATA1.ExecuteSQL("DELETE FROM " & WHTLOCBC)
            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCBC & " " & sqlWHTLOCBC)
        End If

        Dim locations_to_avoid As String = "'00-REC-A','00-REC-B','00-LNF-A','00-SHP-A','00-RTN-A','00-FIN-A','00-ADJ-A'"

        Dim sqlWHTLOCBX As String = $"Select X.*,WHTLOCBC.DATE_LAST_CYCLE_COUNT from {WHTLOCBC} WHTLOCBC, (" & vbCrLf _
             & "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
             & ", SUM (WHTLOCB1.LOCATION_QTY) ONH, SUM (WHTLOCB1.LOCATION_QTY_WAVE) WAV" & vbCrLf _
             & ", SUM (DECODE(WHTLOCB1.LOCATION_CODE,'00-REC-A',WHTLOCB1.LOCATION_QTY,0)) RECA" & vbCrLf _
             & ", SUM (DECODE(WHTLOCB1.LOCATION_CODE,'00-REC-B',WHTLOCB1.LOCATION_QTY,0)) RECB" & vbCrLf _
             & ", SUM (DECODE(NVL(WHTLOCM1.LOCATION_USE,'?'),'G',WHTLOCB1.LOCATION_QTY,0)) GUN" & vbCrLf _
             & ", SUM (DECODE(WHTLOCB1.LOCATION_CODE,'00-LNF-A',WHTLOCB1.LOCATION_QTY,0)) LNF" & vbCrLf _
             & ", SUM (DECODE(WHTLOCB1.LOCATION_CODE,'00-SHP-A',WHTLOCB1.LOCATION_QTY,0)) SHP" & vbCrLf _
             & ", SUM (DECODE(WHTLOCB1.LOCATION_CODE,'00-RTN-A',WHTLOCB1.LOCATION_QTY,0)) RTN" & vbCrLf _
             & ", SUM (DECODE(WHTLOCB1.LOCATION_CODE,'00-FIN-A',WHTLOCB1.LOCATION_QTY,0)) FIN" & vbCrLf _
             & ", SUM (DECODE(WHTLOCB1.LOCATION_CODE,'00-ADJ-A',WHTLOCB1.LOCATION_QTY,0)) ADJ" & vbCrLf _
             & $", SUM (CASE WHEN NVL(WHTLOCM1.LOCATION_USE,'?')<>'G' AND WHTLOCB1.LOCATION_CODE NOT IN ({locations_to_avoid}) THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) LOC" & vbCrLf _
             & " from WHTLOCB1,WHTLOCM1" & vbCrLf _
             & $" where WHTLOCB1.WHSE_CODE = '{WHSE_CODE}'" & vbCrLf _
             & "   and (WHTLOCB1.LOCATION_QTY <> 0 OR WHTLOCB1.LOCATION_QTY_WAVE <> 0)" & vbCrLf _
             & "   and WHTLOCM1.WHSE_CODE = WHTLOCB1.WHSE_CODE and WHTLOCM1.LOCATION_CODE = WHTLOCB1.LOCATION_CODE" & vbCrLf _
             & " group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
             & ") X WHERE WHTLOCBC.STYLE_CODE(+) = X.STYLE_CODE AND WHTLOCBC.COLOR_CODE(+) = X.COLOR_CODE"

        If Initialize Then
            ASCMAIN1.sql = sqlWHTLOCBX
            WHTLOCBX = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("DELETE FROM " & WHTLOCBX)
            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCBX & " " & sqlWHTLOCBX)
        End If


        Dim sqlWHTLOCLX As String = "Select WHTCYCL4.LOCATION_CODE, MAX(WHTCYCL4.INIT_DATE) LAST_CYCLE_COUNT" & vbCrLf _
            & $" from WHTCYCL4 where WHTCYCL4.WHSE_CODE = '{WHSE_CODE}'" & vbCrLf _
            & " group by LOCATION_CODE"

        If Initialize Then
            ASCMAIN1.sql = sqlWHTLOCLX
            WHTLOCLX = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("DELETE FROM " & WHTLOCLX)
            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCLX & " " & sqlWHTLOCLX)

            Dim sqlOtherLocations As String = "Select LOCATION_CODE from WHTLOCM1" & vbCrLf _
                & $" where WHSE_CODE = '{WHSE_CODE}' and LOCATION_CODE not in ({locations_to_avoid})" & vbCrLf _
                & $" and LOCATION_CODE in (Select Distinct LOCATION_CODE from WHTLOCB1 where WHSE_CODE = '{WHSE_CODE}' and LOCATION_QTY <> 0)" & vbCrLf _
                & " minus " & vbCrLf _
                & $" Select LOCATION_CODE from {WHTLOCLX}"
            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCLX & " (LOCATION_CODE) " & sqlOtherLocations)
        End If

    End Sub

    Private Sub grdWHTLOCBX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTLOCBX.InitializeRow

        If e.Row.Cells("DATE_LAST_CYCLE_COUNT").Value & "" <> "" Then
            Dim DAYS_WINDOW As Integer = Val(numDays.Value & "")
            Dim DATE_LAST_CYCLE_COUNT As Date = e.Row.Cells("DATE_LAST_CYCLE_COUNT").Value & ""
            If Format(DATE_LAST_CYCLE_COUNT.AddDays(DAYS_WINDOW), "yyyyMMdd") < Format(Now.Date, "yyyyMMdd") Then
                e.Row.Cells("DATE_LAST_CYCLE_COUNT").Appearance.ForeColor = System.Drawing.Color.Red
            Else
                e.Row.Cells("DATE_LAST_CYCLE_COUNT").Appearance.ForeColor = System.Drawing.Color.Empty
            End If
        Else
            e.Row.Cells("DATE_LAST_CYCLE_COUNT").Appearance.ForeColor = System.Drawing.Color.Empty
        End If

    End Sub

    Private Sub grdWHTLOCLX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTLOCLX.AfterRowActivate
        Setup_WHTLOCLY()
    End Sub

    Sub Setup_WHTLOCLY()
        If grdWHTLOCLX.ActiveRow IsNot Nothing AndAlso grdWHTLOCLX.ActiveRow.IsDataRow AndAlso Not grdWHTLOCLX.ActiveRow.IsFilterRow Then
            Dim LOCATION_CODE As String = grdWHTLOCLX.ActiveRow.Cells("LOCATION_CODE").Value

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Location Cycle Count History", "")

            Fill_Records("WHTLOCLY", New String() {WHSE_CODE, LOCATION_CODE})
            Sort_grdColumns(grdWHTLOCLY, "INIT_DATE".ToLower)

            grdWHTLOCLY.Visible = True
            grdWHTLOCLY.Text = $"Cycle Count History for Location {LOCATION_CODE}"

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")

        Else
            grdWHTLOClY.Visible = False
        End If
    End Sub

    Private Sub chkPastDueCountsOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkPastDueCountsOnly.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        Set_WHTLOCBX()
        Set_WHTLOClX()
    End Sub

    Private Sub grdWHTLOCLX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTLOCLX.InitializeRow

        If e.Row.Cells("LAST_CYCLE_COUNT").Value & "" <> "" Then
            Dim DAYS_WINDOW As Integer = Val(numDays.Value & "")
            Dim LAST_CYCLE_COUNT As Date = e.Row.Cells("LAST_CYCLE_COUNT").Value & ""
            If Format(LAST_CYCLE_COUNT.AddDays(DAYS_WINDOW), "yyyyMMdd") < Format(Now.Date, "yyyyMMdd") Then
                e.Row.Cells("LAST_CYCLE_COUNT").Appearance.ForeColor = System.Drawing.Color.Red
            Else
                e.Row.Cells("LAST_CYCLE_COUNT").Appearance.ForeColor = System.Drawing.Color.Empty
            End If
        Else
            e.Row.Cells("LAST_CYCLE_COUNT").Appearance.ForeColor = System.Drawing.Color.Empty
        End If
    End Sub

    Private Sub cmdRefreshAdj_Click(sender As Object, e As EventArgs) Handles cmdRefreshAdj.Click
        grdICTIADJX.Text = "Adjustments Date Range From" & Format(dteAdjFrom.Value, "MM/dd/yyyy") & " thru " & Format(dteAdjTo.Value, "MM/dd/yyyy")
        ' need to refresh grdICTIADJX
        Fill_Records("ICTIADJX", New String() {WHSE_CODE})
        Sort_grdColumns(grdICTIADJX, "ADJ_NO")


    End Sub

    Private Sub WHFLNFA1_MouseHover(sender As Object, e As EventArgs) Handles Me.MouseHover

    End Sub
End Class