Imports System.Drawing
Imports System.Drawing.Printing
Imports Infragistics.Win.UltraWinGrid

Public Class WHFLNFA1
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "WHFPACKI" Then
            InquiryMode = True
        End If

        Get_PARM("SOTPARM1")
        With dst

            ASCMAIN1.sql = "Select * from (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", SUM (DECODE(LOCATION_CODE,'00-REC-A',LOCATION_QTY,0)) RECA" & vbCrLf _
                & ", SUM (DECODE(LOCATION_CODE,'00-REC-B',LOCATION_QTY,0)) RECB" & vbCrLf _
                & ", SUM (DECODE(LOCATION_CODE,'00-LNF-A',LOCATION_QTY,0)) LNF" & vbCrLf _
                & ", SUM (DECODE(LOCATION_CODE,'00-SHP-A',LOCATION_QTY,0)) SHP" & vbCrLf _
                & ", SUM (DECODE(LOCATION_CODE,'00-FIN-A',LOCATION_QTY,0)) FIN" & vbCrLf _
                & ", SUM (DECODE(LOCATION_CODE,'00-ADJ-A',LOCATION_QTY,0)) ADJ" & vbCrLf _
                & ", SUM (CASE WHEN LOCATION_CODE NOT IN ('00-REC-A','00-REC-B','00-LNF-A','00-SHP-A','00-FIN-A','00-ADJ-A') THEN LOCATION_QTY ELSE 0 END) LOC" & vbCrLf _
                & " from WHTLOCB1" & vbCrLf _
                & " where WHSE_CODE = :PARM1" & vbCrLf _
                & "   and (LOCATION_QTY <> 0 OR LOCATION_QTY_WAVE <> 0)" & vbCrLf _
                & " group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ") where LNF <> 0"
            Create_TDA(.Tables.Add, "WHTLOCBX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_CODE" & vbCrLf _
                & ", SUM (WHTLOCB1.LOCATION_QTY) LOCATION_QTY, SUM (WHTLOCB1.LOCATION_QTY_WAVE) LOCATION_QTY_WAVE" & vbCrLf _
                & ", MIN (WHTLOCB1.INIT_DATE) INIT_DATE, MAX (WHTLOCB1.LAST_DATE) LAST_DATE, COUNT (DISTINCT WHTLOCB1.BAR_CODE) CARTONS" & vbCrLf _
                & " from WHTLOCB1" & vbCrLf _
                & " where WHTLOCB1.WHSE_CODE = :PARM1 And WHTLOCB1.STYLE_CODE = :PARM2 and WHTLOCB1.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and (WHTLOCB1.LOCATION_QTY <> 0 OR WHTLOCB1.LOCATION_QTY_WAVE <> 0)" & vbCrLf _
                & " group by WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTLOCBY", "**", 0, False, "VVV", 4)

            ASCMAIN1.sql = "Select ICTWHSE1.* from ICTWHSE1" & vbCrLf _
                & "  where ICTWHSE1.WHSE_STATUS = 'A' and ICTWHSE1.WHSE_LOCATOR = '1'"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "ICTWHSE1", "*",, False)
        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdWHTLOCBX.DataSource = dst.Tables("WHTLOCBX")
        grdWHTLOCBY.DataSource = dst.Tables("WHTLOCBY")

        Fill_Records("ICTWHSEX")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")


        Create_Summary(grdWHTLOCBX, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCBX, New String() {"RECA", "RECB", "LNF", "SHP", "FIN", "ADJ", "LOC"})

        With grdWHTLOCBX.DisplayLayout.Bands("WHTLOCBX")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
        End With

        Show_Filter(grdWHTLOCBX)


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

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTLOCBX", "WHTLOCBY"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
        Setup_tab0()
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        rowICTWHSE1 = Fill_Record("ICTWHSE1", WHSE_CODE)

        Fill_Records("WHTLOCBX", WHSE_CODE)

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
        Load_Popup_Menu(grdICTWHSEX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdWHTLOCBX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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
        SETUP_WHTLOCBY
    End Sub

    Sub Setup_WHTLOCBY()
        If grdWHTLOCBX.ActiveRow IsNot Nothing AndAlso grdWHTLOCBX.ActiveRow.IsDataRow AndAlso Not grdWHTLOCBX.ActiveRow.IsFilterRow Then
            Dim STYLE_CODE As String = grdWHTLOCBX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdWHTLOCBX.ActiveRow.Cells("COLOR_CODE").Value

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Locations", "")

            Fill_Records("WHTLOCBY", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
            Sort_grdColumns(grdWHTLOCBY, "LOCATION_CODE")
            grdWHTLOCBY.Visible = True
            grdWHTLOCBY.Text = $"Location Details for {STYLE_CODE}-{COLOR_CODE}"
        Else
            grdWHTLOCBY.Visible = False
        End If
    End Sub
End Class