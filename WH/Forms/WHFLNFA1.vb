Imports System.Drawing
Imports System.Drawing.Printing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFLNFA1
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

    Dim WHTLOCBX As String
    Dim WHTLOCBC As String
    Dim WHTLOCLX As String
    Dim WHTLOCBL As String
    Dim LOCBY_LAST_SEL As String


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

            '                & "   and (WHTLOCB1.LOCATION_QTY <> 0 OR WHTLOCB1.LOCATION_QTY_WAVE <> 0)" & vbCrLf _

            ASCMAIN1.sql = "Select X.*, Y.DATE_LAST_COUNTED from (" & vbCrLf _
                & "Select WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & ", SUM (WHTLOCB1.LOCATION_QTY) LOCATION_QTY, SUM (WHTLOCB1.LOCATION_QTY_WAVE) LOCATION_QTY_WAVE" & vbCrLf _
                & ", MIN (WHTLOCB1.INIT_DATE) INIT_DATE, MAX (WHTLOCB1.LAST_DATE) LAST_DATE, COUNT (DISTINCT WHTLOCB1.BAR_CODE) CARTONS" & vbCrLf _
                & " from WHTLOCB1 " & vbCrLf _
                & " where WHTLOCB1.WHSE_CODE = :PARM1 And WHTLOCB1.STYLE_CODE = :PARM2 and WHTLOCB1.COLOR_CODE = :PARM3" & vbCrLf _
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

            '                & "   and ICTIADJ2.LOCATION_CODE = :PARM4" & vbCrLf _
            ASCMAIN1.sql = "Select ICTIADJ2.*, ICTIADJ1.ADJ_DATE, ICTIADJ1.REASON_CODE, ICTIADJ1.ADJ_NOTE" & vbCrLf _
                & " from ICTIADJ1, ICTIADJ2" & vbCrLf _
                & " where ICTIADJ1.WHSE_CODE = :PARM1 And ICTIADJ2.STYLE_CODE = :PARM2 and ICTIADJ2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and ICTIADJ1.ADJ_NO = ICTIADJ2.ADJ_NO"
            Create_TDA(.Tables.Add, "ICTIADJ2", "**", 0, False, "VVV", 0)

            dteAdjFrom.DateTime = DateAdd(DateInterval.Day, -10, DateTime.Now)
            dteAdjTo.DateTime = DateTime.Now


            ASCMAIN1.sql = "Select ICTIADJ2.*, ICTIADJ1.ADJ_DATE, ICTIADJ1.REASON_CODE, ICTIADJ1.ADJ_NOTE" & vbCrLf _
                & " from ICTIADJ1, ICTIADJ2" & vbCrLf _
                & " where ICTIADJ1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and ICTIADJ1.INIT_DATE >= :PARM2" & vbCrLf _
                & "   and ICTIADJ1.INIT_DATE <  :PARM3" & vbCrLf _
                & "   and ICTIADJ1.ADJ_NO = ICTIADJ2.ADJ_NO"
            Create_TDA(.Tables.Add, "ICTIADJX", "**", 0, False, "VDD", 0)


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
        Create_Summary(grdWHTLOCLX, New String() {"LOCATION_QTY"})

        Create_Summary(grdWHTLOCBX, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCBX, New String() {"ONH", "PICK", "WAV", "RECA", "RECB", "GUN", "LNF", "SHP", "RTN", "FIN", "ADJ", "LOC"})

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

        ASCMAIN1.Add_Value_List(grdWHTLOCBZ, "WHSE_TRAN_TYPE")

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

            Case "Load", "View", "Refresh"
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
                    .Items("Refresh").Settings.Enabled = iScreenMode
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

        optLOC.Visible = ScreenMode
        Set_Read_Only_for_ctl(optLOC, False)



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
        Set_WHTLOCBX()
        Sort_grdColumns(grdWHTLOCBX, "STYLE_CODE, COLOR_CODE")

        Fill_Records("WHTLOCLX")
        Set_WHTLOCLX()
        Sort_grdColumns(grdWHTLOCLX, "LOCATION_CODE")

        Set_ICTIADJX()


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
        Load_Popup_Menu(grdWHTLOCBX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Adjust", "Style Status Inquiry")
        Load_Popup_Menu(grdWHTLOCBY, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Move", "Consolidate")
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
                Case "grdWHTLOCBY"
                    tlb_btn = DirectCast(tlb_pop.Tools("Move"), UltraWinToolbars.ButtonTool)
                    Dim enableMove As Boolean = False
                    Dim FROM_LOC As String = ""
                    If grd.Selected.Rows.Count = 2 Then
                        enableMove = True
                        For Each grdR As UltraGridRow In grd.Selected.Rows
                            If grdR.Cells("LOCATION_CODE").Value <> LOCBY_LAST_SEL Then
                                FROM_LOC = grdR.Cells("LOCATION_CODE").Value
                            End If
                        Next
                        tlb_btn.SharedProps.Caption = $"Move {FROM_LOC} to {LOCBY_LAST_SEL}"
                    End If
                    tlb_btn.SharedProps.Visible = enableMove

                    tlb_btn = DirectCast(tlb_pop.Tools("Consolidate"), UltraWinToolbars.ButtonTool)
                    Dim enableCons = False
                    FROM_LOC = ""
                    If grd.Selected.Rows.Count = 1 Then
                        Dim LOCATION_USE As String = ASCDATA1.GetDataValue($"SELECT LOCATION_USE FROM WHTLOCM1 where WHSE_CODE = '{WHSE_CODE}' and LOCATION_CODE = '{LOCBY_LAST_SEL}'")
                        If LOCATION_USE = "S" Or LOCATION_USE = "L" Then
                            enableCons = True
                            tlb_btn.SharedProps.Caption = $"Consolidate to {LOCBY_LAST_SEL}"
                        End If
                    End If
                    tlb_btn.SharedProps.Visible = enableCons
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
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress($"Building records for {e.Tool.Key}")
        Select Case e.Tool.Key
            Case "Adjust"
                Dim LNF As Int64 = Val(grd.ActiveRow.Cells("LNF").Value & "")
                If LNF <> 0 Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value & ""
                    Dim COLOR_CODE As String = grd.ActiveRow.Cells("COLOR_CODE").Value & ""
                    Dim LOCATION_CODE As String = "00-LNF-A" ' grd.ActiveRow.Cells("LOCATION_CODE").Value & ""
                    Dim LOCATION_CODE_TO As String = ""
                    Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE})
                    Adjustment(STYLE_CODE, COLOR_CODE, LOCATION_CODE, LOCATION_CODE_TO, "ADJ")
                End If
            Case "Move"
                Dim STYLE_CODE As String = grdWHTLOCBX.ActiveRow.Cells("STYLE_CODE").Value & ""
                Dim COLOR_CODE As String = grdWHTLOCBX.ActiveRow.Cells("COLOR_CODE").Value & ""
                Dim LOCATION_CODE As String = ""
                Dim LOCATION_CODE_TO As String = LOCBY_LAST_SEL
                If grd.Selected.Rows.Count = 2 Then
                    For Each grdR As UltraGridRow In grd.Selected.Rows
                        If grdR.Cells("LOCATION_CODE").Value <> LOCBY_LAST_SEL Then
                            LOCATION_CODE = grdR.Cells("LOCATION_CODE").Value
                        End If
                    Next
                    Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE})
                    Adjustment(STYLE_CODE, COLOR_CODE, LOCATION_CODE, LOCATION_CODE_TO, "") ' Type Move is default
                End If
            Case "Consolidate"
                Dim STYLE_CODE As String = grdWHTLOCBX.ActiveRow.Cells("STYLE_CODE").Value & ""
                Dim COLOR_CODE As String = grdWHTLOCBX.ActiveRow.Cells("COLOR_CODE").Value & ""
                Dim LOCATION_CODE As String = ""
                Dim LOCATION_CODE_TO As String = LOCBY_LAST_SEL
                If grd.Selected.Rows.Count = 1 Then
                    dst.Tables("WHTLOCB1").Rows.Clear()
                    For Each grdR As UltraGridRow In grd.Rows
                        If grdR.Cells("LOCATION_CODE").Value <> LOCBY_LAST_SEL And ("00-,99-".Contains(grdR.Cells("LOCATION_CODE").Value.ToString.Substring(0, 2))) Then
                            LOCATION_CODE = grdR.Cells("LOCATION_CODE").Value
                            Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE}, False)
                        End If
                    Next
                    Adjustment(STYLE_CODE, COLOR_CODE, LOCATION_CODE, LOCATION_CODE_TO, "CONS") ' Type Move is default
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
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

    Private Sub grdWHTLOCBY_AfterSelectChange(sender As Object, e As AfterSelectChangeEventArgs) Handles grdWHTLOCBY.AfterSelectChange
        'Debug.Print(e.Type.ToString)
        If e.Type.Name = "UltraGridRow" Then
            If grdWHTLOCBY.Selected.Rows.Count = 0 Then
                LOCBY_LAST_SEL = ""
            Else
                LOCBY_LAST_SEL = grdWHTLOCBY.ActiveRow.Cells("LOCATION_CODE").Value
            End If
        End If
    End Sub
    Sub Setup_WHTLOCBY()
        If grdWHTLOCBX.ActiveRow IsNot Nothing AndAlso grdWHTLOCBX.ActiveRow.IsDataRow AndAlso Not grdWHTLOCBX.ActiveRow.IsFilterRow Then
            Dim STYLE_CODE As String = grdWHTLOCBX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdWHTLOCBX.ActiveRow.Cells("COLOR_CODE").Value

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Locations", "")

            Fill_Records("WHTLOCBY", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            Set_WHTLOCBY()
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

            '            Fill_Records("ICTIADJ2", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE, LOCATION_CODE})
            Fill_Records("ICTIADJ2", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            Sort_grdColumns(grdICTIADJ2, "ADJ_NO")

            grdICTIADJ2.Visible = True
            grdICTIADJ2.Text = $"Adjustments Details for {STYLE_CODE}-{COLOR_CODE}"  '  in Location {LOCATION_CODE}"

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

    Private Sub optLOC_ValueChanged(sender As Object, e As EventArgs) Handles optLOC.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_WHTLOCLX()
        Set_WHTLOCBY()
    End Sub

    Sub Set_WHTLOCLX()

        Dim sqlwhere As String = ""
        Dim dvw As DataView = DirectCast(grdWHTLOCLX.DataSource, DataTable).DefaultView

        If optLOC.Value = "ALL" Then
        Else
            sqlwhere &= " and LOCATION_QTY <> 0"
        End If

        If chkPastDueCountsOnly.Checked Then
            Dim DATE_LAST_CYCLE_COUNT_cutoff As Date = Now.Date.AddDays(-1 * Val(numDays.Value & ""))
            sqlwhere &= $" and LAST_CYCLE_COUNT is Null or LAST_CYCLE_COUNT < '{Format(DATE_LAST_CYCLE_COUNT_cutoff, "MM/dd/yyyy")}'"
        End If

        dvw.RowFilter = Mid(sqlwhere, 5)

    End Sub

    Sub Set_WHTLOCBY()

        Dim sqlwhere As String = ""
        Dim dvw As DataView = DirectCast(grdWHTLOCBY.DataSource, DataTable).DefaultView

        If optLOC.Value = "ALL" Then
        Else
            sqlwhere &= " and (LOCATION_QTY <> 0 or DATE_LAST_COUNTED is Not Null)"
        End If

        'If chkPastDueCountsOnly.Checked Then
        '    Dim DATE_LAST_CYCLE_COUNT_cutoff As Date = Now.Date.AddDays(-1 * Val(numDays.Value & ""))
        '    sqlwhere &= $" and LAST_CYCLE_COUNT is Null or LAST_CYCLE_COUNT < '{Format(DATE_LAST_CYCLE_COUNT_cutoff, "MM/dd/yyyy")}'"
        'End If

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

    Sub Adjustment(STYLE_CODE As String, COLOR_CODE As String, LOCATION_CODE As String, LOCATION_CODE_TO As String, movement_type As String)

        Dim confirm_only As Boolean = True ' (movement_type = "LNF") Or (movement_type = "CMB")
        Dim dupBarcode As String = ""

        If LOCATION_CODE = "00-FIN-A" Or LOCATION_CODE_TO = "00-FIN-A" Then
            MsgBox("Cannot Change or Move Cases in Financial Location", MsgBoxStyle.OkOnly, "Cannot Move")
            Exit Sub
        End If

        Using ff As New TAC.TAFLOCM1()

            Dim BAR_CODE_CMB As String = ""


            ff.confirm_only = confirm_only
            ff.movement_type = movement_type
            ff.rowICTWHSE1 = rowICTWHSE1
            ff.WHSE_CODE = WHSE_CODE

            If movement_type = "ADJ" Then
                ff.REASON_CODE = "WHLOC"
            End If
            'ff.disableUpdate = True
            If movement_type = "CONS" Then
                ' Move is default
                ff.movement_type = ""
            End If


            Dim BCs As New List(Of String)
            Dim LOCs As New List(Of String)
            Dim prePacks As Boolean = False
            Dim SkippedCnt As Int32 = 0
            Dim OkToMoveDefaultBarcode As Boolean = False

            Dim sqlw As String = $"WHSE_CODE = '{WHSE_CODE}' and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' and LOCATION_CODE = '{LOCATION_CODE}' and LOCATION_QTY <> 0 and BAR_CODE <> '{rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")}'"
            If movement_type = "CONS" Then
                sqlw = $"WHSE_CODE = '{WHSE_CODE}' and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' and LOCATION_CODE <> '{LOCATION_CODE_TO}' and LOCATION_QTY <> 0"
            End If
            If movement_type = "" Then
                Dim LOCATION_USE As String = ASCDATA1.GetDataValue($"SELECT LOCATION_USE FROM WHTLOCM1 where WHSE_CODE = '{WHSE_CODE}' and LOCATION_CODE = '{LOCATION_CODE}'")
                If LOCATION_USE = "S" Or LOCATION_USE = "L" Then
                    Dim LOCATION_USE_TO As String = ASCDATA1.GetDataValue($"SELECT LOCATION_USE FROM WHTLOCM1 where WHSE_CODE = '{WHSE_CODE}' and LOCATION_CODE = '{LOCATION_CODE_TO}'")
                    If LOCATION_USE_TO = "S" Or LOCATION_USE_TO = "L" Then
                        OkToMoveDefaultBarcode = True
                    End If
                End If
                End If


                For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("")

                Dim BAR_CODE As String = rowWHTLOCB1.Item("BAR_CODE") & ""
                Dim LOAD_NO As String = "" ' row.Item("LOAD_NO") & ""
                If movement_type = "CONS" Then
                    LOCATION_CODE = rowWHTLOCB1.Item("LOCATION_CODE") & ""
                End If

                If LOCATION_CODE = "00-FIN-A" Then
                    MsgBox("Cannot Change or Move Cases in Financial Location", MsgBoxStyle.OkOnly, "Cannot Move")
                    Exit Sub
                End If
                If BAR_CODE = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE") Then
                    If movement_type <> "ADJ" And movement_type <> "CONS" And OkToMoveDefaultBarcode = False Then
                        MsgBox("Cannot Change or Move a Case with no LPN", MsgBoxStyle.OkOnly, "Cannot Move")
                        Exit Sub
                    End If
                Else
                    ASCMAIN1.sql = $"SELECT * FROM WHTLOCB1 Where WHSE_CODE = '{WHSE_CODE}'  and LOCATION_CODE = '{LOCATION_CODE}' and LOCATION_QTY <> 0 and BAR_CODE = '{BAR_CODE}'"
                    For Each rowStyClr As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select($"STYLE_CODE <> '{STYLE_CODE}' and COLOR_CODE <> '{COLOR_CODE}'")
                        If Val(rowStyClr.Item("LOCATION_QTY_WAVE") & "") <> 0 Then
                            MsgBox("Cannot Change or Move a Case which has been committed to a Wave", MsgBoxStyle.OkOnly, "Cannot Move")
                            Exit Sub
                        End If
                        prePacks = True
                        If movement_type = "CONS" Then
                            'MsgBox("Cannot Consolidate pre-packs", MsgBoxStyle.OkOnly, "Skipping Carton " & BAR_CODE)
                            'skip the carton for other styles
                            SkippedCnt += 1
                            Exit For
                        End If
                        Dim STYLE_CODE1 As String = rowStyClr("STYLE_CODE")
                        Dim COLOR_CODE1 As String = rowStyClr("COLOR_CODE")
                        Dim LOCATION_QTY1 As Int64 = Val(rowStyClr("LOCATION_QTY") & "")
                        'Additional Carton Styles
                        ff.AddItemToMove(WHSE_CODE,
                        LOCATION_CODE,
                            STYLE_CODE1,
                            COLOR_CODE1,
                            BAR_CODE,
                            LOAD_NO,
                            LOCATION_QTY1,
                            LOCATION_CODE_TO, BAR_CODE_CMB)
                    Next

                End If

                If Val(rowWHTLOCB1.Item("LOCATION_QTY_WAVE") & "") <> 0 Then
                    If movement_type = "CONS" Then
                        Continue For
                    End If
                    MsgBox("Cannot Change or Move a Case which has been committed to a Wave", MsgBoxStyle.OkOnly, "Cannot Move")
                    Exit Sub
                End If

                Dim LOCATION_QTY As Int64 = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "")

                If movement_type = "CONS" And prePacks = True Then
                    'skip the carton - for current style
                    prePacks = False
                Else
                    ff.AddItemToMove(WHSE_CODE,
                        LOCATION_CODE,
                            STYLE_CODE,
                            COLOR_CODE,
                            BAR_CODE,
                            LOAD_NO,
                            LOCATION_QTY,
                            LOCATION_CODE_TO, BAR_CODE_CMB)
                End If
            Next

            If prePacks = True Then
                MsgBox("Cases whith pre-Packs have been selected, this will affect other styles", MsgBoxStyle.OkOnly, "Verify Styles")
            End If

            If SkippedCnt > 0 Then
                MsgBox("Cases whith pre-Packs have been skipped to avoid other styles", MsgBoxStyle.OkOnly, $"{SkippedCnt} cases skipped")
            End If

            ff.ShowDialog()

            Dim WHSE_TRAN_NO As String = IIf(movement_type = "ADJ", ff.ADJ_NO, ff.WHSE_TRAN_NO)
            If WHSE_TRAN_NO <> "" Then

                Dim sqlWHTLOCB2where As String = $" Where WHSE_CODE = '{WHSE_CODE}' and STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}' and LOCATION_CODE = '{LOCATION_CODE}' and BAR_CODE <> '{rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")}'"

                ASCMAIN1.sql = "Select * from WHTLOCB2" _
                    & sqlWHTLOCB2where _
                    & $" and WHSE_TRAN_TYPE = '{IIf(movement_type = "ADJ", "A", "M")}' and WHSE_TRAN_NO = '{WHSE_TRAN_NO}'"

                Fill_Records("WHTLOCB2", "", False, ASCMAIN1.sql)

                ASCMAIN1.Progress("Refreshing Data", "")
                Dim special_locs As String = "'00-REC-A','00-REC-B','00-LNF-A','00-SHP-A','00-RTN-A','00-FIN-A','00-ADJ-A'"
                Dim grdRow As UltraGridRow = grdWHTLOCBX.ActiveRow
                Dim AdjQty As Int64 = Val(dst.Tables("WHTLOCB2").Compute("SUM(WHSE_TRAN_QTY)", $"LOCATION_CODE = '{LOCATION_CODE}'") & "")
                If movement_type = "ADJ" Then
                    grdRow.Cells("LNF").Value = Val(grdRow.Cells("LNF").Value) + AdjQty
                    grdRow.Cells("ONH").Value = Val(grdRow.Cells("ONH").Value) + AdjQty
                Else
                    Dim s = 1
                    For Each loc As String In New Object() {LOCATION_CODE, LOCATION_CODE_TO}
                        Dim LOCATION_USE As String = ASCDATA1.GetDataValue($"Select nvl(LOCATION_USE,'A') From WHTLOCM1 where WHSE_CODE = '{WHSE_CODE}' and LOCATION_CODE = '{loc}'") & ""
                        grdRow.Cells("RECA").Value = Val(grdRow.Cells("RECA").Value) + IIf(loc = "00-REC-A", AdjQty * s, 0)
                        grdRow.Cells("RECB").Value = Val(grdRow.Cells("RECB").Value) + IIf(loc = "00-REC-B", AdjQty * s, 0)
                        grdRow.Cells("GUN").Value = Val(grdRow.Cells("GUN").Value) + IIf(LOCATION_USE = "G", AdjQty * s, 0)
                        grdRow.Cells("LNF").Value = Val(grdRow.Cells("LNF").Value) + IIf(loc = "00-LNF-A", AdjQty * s, 0)
                        grdRow.Cells("SHP").Value = Val(grdRow.Cells("SHP").Value) + IIf(loc = "00-SHP-A", AdjQty * s, 0)
                        grdRow.Cells("RTN").Value = Val(grdRow.Cells("RTN").Value) + IIf(loc = "00-RTN-A", AdjQty * s, 0)
                        grdRow.Cells("FIN").Value = Val(grdRow.Cells("FIN").Value) + IIf(loc = "00-FIN-A", AdjQty * s, 0)
                        grdRow.Cells("ADJ").Value = Val(grdRow.Cells("ADJ").Value) + IIf(loc = "00-ADJ-A", AdjQty * s, 0)
                        grdRow.Cells("LOC").Value = Val(grdRow.Cells("LOC").Value) + IIf(LOCATION_USE <> "G" And Not special_locs.Contains(loc), AdjQty * s, 0)
                        s = s * -1
                    Next
                End If
                'Load_WHTLOCB1()
                Setup_WHTLOCBY()
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

        Dim sqlWHTLOCBX As String = $"Select X.*, 0 PICK, ' ' STYLE_DESC, WHTLOCBC.DATE_LAST_CYCLE_COUNT from {WHTLOCBC} WHTLOCBC, (" & vbCrLf _
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
            ASCDATA1.ExecuteSQL("Alter Table " & WHTLOCBX & " MODIFY PICK NUMBER(8)")
            ASCDATA1.ExecuteSQL("Alter Table " & WHTLOCBX & " MODIFY STYLE_DESC VARCHAR2(256)")
        Else
            ASCDATA1.ExecuteSQL("DELETE FROM " & WHTLOCBX)
            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCBX & " " & sqlWHTLOCBX)
        End If
        ASCMAIN1.sql = $"Update {WHTLOCBX} x set STYLE_DESC = (Select STYLE_DESC from ICTSTYL1 where STYLE_CODE = x.STYLE_CODE)" & vbCrLf _
            & " where exists (Select STYLE_DESC from ICTSTYL1 where STYLE_CODE = x.STYLE_CODE)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = $"Update {WHTLOCBX} x set PICK = (" & vbCrLf _
            & $" Select WHSE_QTY_PICK from ICTSTAT2 where STYLE_CODE = x.STYLE_CODE and COLOR_CODE = x.COLOR_CODE and WHSE_CODE = '{WHSE_CODE}')" & vbCrLf _
            & $" where exists (Select STYLE_DESC from ICTSTAT2 where STYLE_CODE = x.STYLE_CODE and COLOR_CODE = x.COLOR_CODE and WHSE_CODE = '{WHSE_CODE}')"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Dim sqlWHTLOCBL As String = "Select WHTLOCB1.LOCATION_CODE" & vbCrLf _
            & ", Count ( Distinct (Case when LOCATION_QTY <> 0 THEN WHTLOCB1.STYLE_CODE || WHTLOCB1.COLOR_CODE else null end)) SCS" & vbCrLf _
            & ", Sum (WHTLOCB1.LOCATION_QTY) LOCATION_QTY, MAX(WHTLOCB1.LAST_DATE) LAST_DATE" & vbCrLf _
            & $" from WHTLOCB1 where WHTLOCB1.WHSE_CODE = '{WHSE_CODE}'" & vbCrLf _
            & " group by WHTLOCB1.LOCATION_CODE"

        If Initialize Then
            ASCMAIN1.sql = sqlWHTLOCBL
            WHTLOCBL = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("DELETE FROM " & WHTLOCBL)
            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCBL & " " & sqlWHTLOCBL)
        End If


        Dim sqlWHTLOCLX As String = "Select WHTCYCL4.LOCATION_CODE, MAX(WHTCYCL4.INIT_DATE) LAST_CYCLE_COUNT" & vbCrLf _
            & $" from WHTCYCL4 where WHTCYCL4.WHSE_CODE = '{WHSE_CODE}'" & vbCrLf _
            & " group by LOCATION_CODE"

        sqlWHTLOCLX = $"Select X.*, WHTLOCBL.SCS, WHTLOCBL.LOCATION_QTY, WHTLOCBL.LAST_DATE from ({sqlWHTLOCLX}) X, {WHTLOCBL} WHTLOCBL where WHTLOCBL.LOCATION_CODE (+) = x.LOCATION_CODE"

        If Initialize Then
            ASCMAIN1.sql = sqlWHTLOCLX
            WHTLOCLX = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("DELETE FROM " & WHTLOCLX)
            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCLX & " " & sqlWHTLOCLX)

            '                & $" where WHSE_CODE = '{WHSE_CODE}' and LOCATION_CODE not in ({locations_to_avoid})" & vbCrLf _
            Dim sqlOtherLocations As String = "Select LOCATION_CODE from WHTLOCM1" & vbCrLf _
                & $" where WHSE_CODE = '{WHSE_CODE}'" & vbCrLf _
                & $" and LOCATION_CODE in (Select Distinct LOCATION_CODE from WHTLOCB1 where WHSE_CODE = '{WHSE_CODE}')" & vbCrLf _
                & " minus " & vbCrLf _
                & $" Select LOCATION_CODE from {WHTLOCLX}"

            sqlOtherLocations = $"Select X.*, WHTLOCBL.SCS, WHTLOCBL.LOCATION_QTY, WHTLOCBL.LAST_DATE from ({sqlOtherLocations}) X, {WHTLOCBL} WHTLOCBL where WHTLOCBL.LOCATION_CODE (+) = x.LOCATION_CODE"



            ASCDATA1.ExecuteSQL("Insert into " & WHTLOCLX & " (LOCATION_CODE, SCS, LOCATION_QTY, LAST_DATE) " & sqlOtherLocations)
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
            grdWHTLOCLY.Visible = False
        End If
    End Sub

    Private Sub chkPastDueCountsOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkPastDueCountsOnly.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        Set_WHTLOCBX()
        Set_WHTLOCLX()
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
        Set_ICTIADJX()
    End Sub

    Sub Set_ICTIADJX()


        Dim ODT_FROM As Date = CDate(dteAdjFrom.Value)
        Dim ODT_TO As Date = CDate(dteAdjTo.Value).AddDays(1)

        Dim DT_FROM As String = Format(CDate(dteAdjFrom.Value), "dd-MMM-yyyy")
        Dim DT_TO As String = Format(CDate(dteAdjTo.Value), "dd-MMM-yyyy")

        Fill_Records("ICTIADJX", New Object() {WHSE_CODE, ODT_FROM, ODT_TO})
        Sort_grdColumns(grdICTIADJX, "ADJ_NO")
        grdICTIADJX.Text = $"{WHSE_CODE} Locator Adjustment History from {DT_FROM} TO {DT_TO}"
    End Sub
End Class