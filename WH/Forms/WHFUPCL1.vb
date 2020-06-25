Public Class WHFUPCL1

    Dim rowWHTUPCL1 As DataRow
    Dim rowWHTUPCLP As DataRow
    '   Public DymoAddIn As Dymo.DymoAddIn
    '   Public DymoLabels As Dymo.DymoLabels


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("WHTPARM1")

        '        DymoAddIn = New Dymo.DymoAddIn
        '        DymoLabels = New Dymo.DymoLabels
        '        Dim LabelName As String = ASCMAIN1.Folders("Images") & "UPC.label"

        'C:\VS\VDI\WH\RGI.label
        '        DymoAddIn.Open(LabelName)

        ' obtain the currently selected printer
        '        SetupLabelWriterSelection(True)

        With dst
            Create_TDA(.Tables.Add, "WHTUPCL1", "*")

            ASCMAIN1.sql = "Select WHTUPCL1.* from WHTUPCL1 where PROCESS_IND = '1'"
            Create_TDA(.Tables.Add, "WHTUPCLX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select WHTUPCL1.* from WHTUPCL1 where PROCESS_IND = '0'"
            Create_TDA(.Tables.Add, "WHTUPCLP", "**", 0, False, "", 1)
            .Tables("WHTUPCLP").Columns.Add("SEL")
            .Tables("WHTUPCLP").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select WHTUPCL1.* from WHTUPCL1"
            Create_TDA(.Tables.Add, "WHTUPCLR", "**", 0, False, "", 1)
            .Tables("WHTUPCLR").Columns.Add("SEL")
            .Tables("WHTUPCLR").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select ASTUSER2.* from ASTUSER2 where SECURITY_CODE = 'WS'"
            Create_TDA(.Tables.Add, "ASTUSER2", "**", 0, False, "", 1)
            .Tables("ASTUSER2").Columns.Add("SEL")
            .Tables("ASTUSER2").Columns("SEL").DefaultValue = "0"

        End With

        grdWHTUPCLX.DataSource = dst.Tables("WHTUPCLX")
        grdWHTUPCLP.DataSource = dst.Tables("WHTUPCLP")
        grdWHTUPCLR.DataSource = dst.Tables("WHTUPCLR")
        grdASTUSER2.DataSource = dst.Tables("ASTUSER2")


        grdWHTUPCLP.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For Each gcol As UltraWinGrid.UltraGridColumn In grdWHTUPCLP.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SEL" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
            End If
        Next

        grdWHTUPCLR.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For Each gcol As UltraWinGrid.UltraGridColumn In grdWHTUPCLR.DisplayLayout.Bands(0).Columns
            If gcol.Key = "LBL_QTY" Or gcol.Key = "SEL" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
            End If
        Next

        grdASTUSER2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For Each gcol As UltraWinGrid.UltraGridColumn In grdASTUSER2.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SEL" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
            End If
        Next

        Dim rows() As DataRow = ASCDATA1.GetDataTable("SELECT *  FROM WHTLPRT1").Select("")
        For Each row As DataRow In rows
            cbxLabelPrinter.Items.Add(row.Item("LABEL_PRINTER_ID"))
        Next
        cbxLabelPrinter.SelectedIndex = 0

        Create_Summary(grdWHTUPCLX, "LBL_REQ_NO", "Count")

        Create_Summary(grdWHTUPCLP, "LBL_REQ_NO", "Count")
        Create_Summary(grdWHTUPCLP, "SEL")

        Create_Summary(grdWHTUPCLR, "LBL_REQ_NO", "Count")

        Show_Filter(grdWHTUPCLP, True)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Validate_Code("UPC_CODE") Then
                    Absx1.txtFor("STYLE_CODE").Text = cdr.Item("STYLE_CODE")
                    Absx1.txtFor("COLOR_CODE").Text = cdr.Item("COLOR_CODE")
                Else
                    EMsg &= vbCr & "You Must Specify a valid style and color for upc"
                End If

            Case "View"
                If Absx1.txtFor("LBL_REQ_NO").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Label Request No"
                Else
                    ASCMAIN1.sql = "Select * from WHTUPCL1 where LBL_REQ_NO = '" & Absx1.txtFor("LBL_REQ_NO").Text & "'"
                    Dim rowWHTUPCL1 As DataRow = ASCDATA1.GetDataRow
                    If rowWHTUPCL1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Label Request No"
                    Else

                    End If
                End If

            Case "Print"

                If Val(Absx1.numFor("LBL_QTY").Value & "") <= 0 Then
                    EMsg &= vbCr & "Invalid Number of Labels Requested"
                End If

                If EMsg = "" Then

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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Print"
                rowWHTUPCL1 = dst.Tables("WHTUPCL1").Rows(0)
                With rowWHTUPCL1
                    .Item("PROCESS_IND") = "1"
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                End With
                printreq(rowWHTUPCL1)
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode

                If ScreenMode And EntryMode <> "N" Then
                    .Groups("Screen Control").Items("Print").Settings.Enabled = not_iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = not_iScreenMode
                Else
                    .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                End If

                If ScreenMode And EntryMode <> "V" Then
                    .Groups("Screen Control").Items("Done").Settings.Enabled = not_iScreenMode
                Else
                    .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                End If

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        optLabel.CheckedIndex = 0

        tab0.Visible = Not tf

        Set_Read_Only(UltraGroupBox2, True)

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTUPCL1", "WHTUPCLP", "WHTUPCLX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
        Setup_tab0()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowWHTUPCL1 = dst.Tables("WHTUPCL1").NewRow
            With rowWHTUPCL1
                .Item("LBL_REQ_NO") = ASCMAIN1.Next_Control_No("WHTUPCL1.LBL_REQ_NO")
                .Item("UPC_CODE") = HFs("UPC_CODE")
                .Item("STYLE_CODE") = HFs("STYLE_CODE")
                .Item("COLOR_CODE") = HFs("COLOR_CODE")
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LBL_QTY") = 1
            End With

            dst.Tables("WHTUPCL1").Rows.Add(rowWHTUPCL1.ItemArray)
        Else
            rowWHTUPCL1 = Fill_Record("WHTUPCL1", New String() {Absx1.txtFor("LBL_REQ_NO").Text})
            dst.AcceptChanges()
        End If

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        Update_Record_TDA("WHTUPCL1")

        CommitTrans("Print Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Call Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "UPC_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    ' Click_Command("New", e)
                    Load_WHTUPCLP()
                End If
            Case "LBL_REQ_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "UPC_CODE"
                '   Click_Command("New")
                Load_WHTUPCLP()
            Case "LBL_REQ_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "UPC_CODE"
                If EntryMode = "" Then
                    'If Absx1.txtFor("UPC_CODE").Text <> "" Then
                    '    LookUp("ARTCUST1", Absx1.txtFor("UPC_CODE").Text)
                    '    If cdr IsNot Nothing Then
                    '        Load_WHTUPCLP()
                    '    End If
                    'End If
                End If

                'Case "STYLE_CODE"
                '    If .Text <> "" And STYLE_CODE <> .Text Then
                '        Dim Valid As Boolean = Validate_Style(.Text, True)
                '        If STYLE_CODE <> "" Then
                '            If COLOR_CODEs.Count = 1 Then
                '                Absx1.txtFor("COLOR_CODE").Text = COLOR_CODEs(0)
                '                Absx1.txtFor("UPC_CODE").Text = UPC_CODE
                '            End If
                '        End If
                '    End If

            Case "COLOR_CODE"
                If ctl.Text <> "" Then
                    LookUp("ICTSTYC1", New String() {Absx1.txtFor("STYLE_CODE").Text, Absx1.txtFor("COLOR_CODE").Text})
                    If cdr IsNot Nothing Then
                        Absx1.txtFor("UPC_CODE").Text = cdr.Item("UPC_CODE")
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "COLOR_CODE"
                sql_where = " COLOR_CODE IN (SELECT COLOR_CODE FROM ICTSTYC1 WHERE STYLE_CODE = '" & Absx1.txtFor("STYLE_CODE").Text & "')"

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LBL_QTY"

        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTUPCLX, "B", "Re-Print")
        Load_Popup_Menu(grdWHTUPCLP, "BBB", "Select All", "Clear All", "Print Selected Labels")
        Load_Popup_Menu(grdWHTUPCLR, "BB", "Select All", "Clear All")
        Load_Popup_Menu(grdASTUSER2, "BBB", "Select All", "Clear All", "Print Name Tags")
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

        Select Case grd.Name
            'Case "grdPOTLCST2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Close"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (optShowChargebacks.Value <> "X")
            '    tlb_btn = DirectCast(tlb_pop.Tools("Re-Open"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (optShowChargebacks.Value = "X")
        End Select


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

            Case "Select All"
                If tab0.ActiveTab.Text = "Bulk Request" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdWHTUPCLR.Rows
                        If Not grow.IsFilteredOut Then
                            grow.Cells("SEL").Value = "1"
                            grow.Update()
                        End If
                    Next
                ElseIf tab0.ActiveTab.Text = "Name Tags" Then
                    For Each grow As UltraWinGrid.UltraGridRow In grdASTUSER2.Rows
                        If Not grow.IsFilteredOut Then
                            grow.Cells("SEL").Value = "1"
                            grow.Update()
                        End If
                    Next
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grdWHTUPCLP.Rows
                        If Not grow.IsFilteredOut Then
                            grow.Cells("SEL").Value = "1"
                            grow.Update()
                        End If
                    Next
                End If

            Case "Clear All"
                If tab0.ActiveTab.Text = "Bulk Request" Then
                    For Each row As DataRow In dst.Tables("WHTUPCLR").Select("SEL='1'")
                        row.Item("SEL") = 0
                    Next
                ElseIf tab0.ActiveTab.Text = "Name Tags" Then
                    For Each row As DataRow In dst.Tables("ASTUSER2").Select("SEL='1'")
                        row.Item("SEL") = 0
                    Next
                Else
                    For Each row As DataRow In dst.Tables("WHTUPCLP").Select("SEL='1'")
                        row.Item("SEL") = 0
                    Next
                End If

            Case "Print Selected Labels"
                If dst.Tables("WHTUPCLP").Select("SEL='1'").Length = 0 Then
                    MsgBox("No Labels Selected to Print", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                If MsgBox(String.Format("OK to Print {0} Selected Labels?", dst.Tables("WHTUPCLP").Select("SEL='1'").Length), MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If
                PrintSelected()
            Case "Print Name Tags"
                If dst.Tables("ASTUSER2").Select("SEL='1'").Length = 0 Then
                    MsgBox("No Names Selected to Print", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                If MsgBox(String.Format("OK to Print {0} Selected Tags?", dst.Tables("ASTUSER2").Select("SEL='1'").Length), MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If
                PrintNameTags()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Re-Print"

                If grd.Selected.Rows.Count = 0 Then
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("First Select Rows by Clicking on the Row Selector to the left of each Row")
                    Exit Sub
                End If
                ReprintSelected()
        End Select
    End Sub

#End Region

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor

        Fill_Records("WHTUPCLX")
        Sort_grdColumns(grdWHTUPCLX, "LBL_REQ_NO,LOCATION_CODE,STYLE_CODE,COLOR_CODE")
        grdWHTUPCLX.Text = "Previously Printed Labels"

        Load_WHTUPCLP()
        Load_ASTUSER2()

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdWHTUPCLX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTUPCLX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("LBL_REQ_NO").Text = e.Row.Cells("LBL_REQ_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        ' UltraExplorerBar1.Groups("Re-Billable Filter").Visible = (tab0.SelectedTab.Key = "Re-Billable Charges")
    End Sub

    Sub Load_WHTUPCLP()
        Me.Cursor = Cursors.WaitCursor
        Fill_Records("WHTUPCLP")
        Sort_grdColumns(grdWHTUPCLP, "LBL_REQ_NO,LOCATION_CODE,STYLE_CODE,COLOR_CODE")
        grdWHTUPCLP.Text = "Label Print Requests Pending"
        Me.Cursor = Cursors.Default
    End Sub

    Sub Load_ASTUSER2()
        Me.Cursor = Cursors.WaitCursor
        Fill_Records("ASTUSER2")
        Sort_grdColumns(grdASTUSER2, "USER_ID")
        grdASTUSER2.Text = "Name Tags for Users"
        Me.Cursor = Cursors.Default
    End Sub

    Sub Display_Totals()

    End Sub
    Sub PrintSelected()
        dst.Tables("WHTUPCL1").Rows.Clear()

        For Each row As DataRow In dst.Tables("WHTUPCLP").Select("SEL='1'")
            Dim LBL_REQ_NO As String = row.Item("LBL_REQ_NO")
            Dim rowWHTUPCL1 As DataRow = Fill_Record("WHTUPCL1", LBL_REQ_NO, False, False)
            With rowWHTUPCL1
                .Item("PROCESS_IND") = "1"
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            End With
            printreq(row)
        Next
        Update_Record()
        Fill_Records("WHTUPCLP")
        Sort_grdColumns(grdWHTUPCLP, "LBL_REQ_NO,LOCATION_CODE,STYLE_CODE,COLOR_CODE")
    End Sub

    Sub PrintNameTags()
        For Each row As DataRow In dst.Tables("ASTUSER2").Select("SEL='1'")
            printreq(row)
        Next

    End Sub

    Sub ReprintSelected()
        For Each grow As UltraWinGrid.UltraGridRow In grdWHTUPCLX.Selected.Rows
            Dim row As DataRow = dst.Tables("WHTUPCLX").Rows.Find(grow.Cells("LBL_REQ_NO").Value)
            printreq(row)
        Next
        grdWHTUPCLX.Selected.Rows.Clear()

    End Sub
    Sub printreq(row As DataRow)
        ' Code for Dymo printer
        Dim CARTONS_PER_UNIT As Int16 = 0
        Dim cnt As Int16 = 0
        Dim rowICTSTYL1 As DataRow
        If row.Table.TableName <> "ASTUSER2" Then
            'ASCMAIN1.sql = "Select STYLE_DESC from ICTSTYL1 where STYLE_CODE = '" & row.Item("STYLE_CODE") & "'"
            rowICTSTYL1 = ASCDATA1.GetDataRow("Select STYLE_DESC, nvl(CARTONS_PER_UNIT,1) CARTONS_PER_UNIT from ICTSTYL1 where STYLE_CODE = :PARM1", "V", row.Item("STYLE_CODE"))
            CARTONS_PER_UNIT = rowICTSTYL1.Item("CARTONS_PER_UNIT")
        End If

        Using ipp As New nsoftware.IPWorks.Ipport
            ipp.RuntimeLicense = "31504E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004B4857525953375A4A5A375A0000"
            If ASCMAIN1.Running_in_VS Then
                ipp.Connect("192.168.1.19", "4444") 'ipp.Connect("192.168.120.67", "4444") '"192.168.4.117", "4444")
            Else
                ipp.Connect("192.168.110.223", "4444")
            End If
            Dim data As String '= "upc123" ' & vbCrLf a new line is needed to send the data across
            Try
                Do
                    cnt += 1
                    data = cbxLabelPrinter.SelectedItem 'Printer
                    If row.Table.TableName = "ASTUSER2" Then
                        data &= "|" & row.Item("USER_ID")
                    Else
                        If optLabel.Value = "L" Then
                            data &= "|" & row.Item("STYLE_CODE") &
                            "|" & row.Item("COLOR_CODE") &
                            "|" & rowICTSTYL1.Item("STYLE_DESC") &
                            "|" & IIf(CARTONS_PER_UNIT > 1, String.Format("{0}-{1}", row.Item("UPC_CODE").ToString.Substring(0, 12), cnt), row.Item("UPC_CODE")) &
                            "|" & row.Item("LBL_QTY")
                            If tab.ActiveTab.Text = "0" And tab0.ActiveTab.Text = "Bulk Request" Then
                                data &= "|" & row.Item("LOCATION_CODE")
                            End If
                        Else
                            data = "NEWER|smallupc.lbx|" & cbxLabelPrinter.SelectedItem
                            data &= "|" & row.Item("STYLE_CODE") &
                            "  " & row.Item("COLOR_CODE") &
                            "|" & rowICTSTYL1.Item("STYLE_DESC") &
                            "|" & row.Item("UPC_CODE") &
                            "|" & row.Item("LBL_QTY")
                            ' small labels aren't for cartons
                            cnt += CARTONS_PER_UNIT
                        End If
                    End If

                    ipp.SendLine(data)
                Loop While CARTONS_PER_UNIT > cnt

            Catch ex As Exception

            End Try

            ipp.Disconnect()
        End Using

    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Me.Cursor = Cursors.WaitCursor
        If btnLoad.Text = "Load" Then
            btnLoad.Text = "Clear"
            ASCMAIN1.sql = "select whtlocb1.LOCATION_CODE, ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYC1.UPC_CODE, round(whtlocb1.LOCATION_QTY/NVL(ICTSTYL1.INNER_PACK_QTY, 0)) LBL_QTY" & vbCrLf _
                            & " from whtlocb1, ICTSTYC1, ICTSTYL1" & vbCrLf _
                            & " where ictstyc1.style_code = whtlocb1.style_code" & vbCrLf _
                            & " and ictstyc1.color_code = whtlocb1.color_code" & vbCrLf _
                            & " and ICTSTYL1.STYLE_CODE = whtlocb1.style_code" & vbCrLf _
                            & " and whtlocb1.WHSE_CODE = 'MS'" & vbCrLf _
                            & " And whtlocb1.LOCATION_QTY > 0" & vbCrLf _
                            & " And ICTSTYL1.INNER_PACK_QTY > 1" & vbCrLf _
                            & " and whtlocb1.LOCATION_CODE like '" & txtLAisle.Value.ToString.ToUpper & "%'"

            Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("", "LOCATION_CODE")
            If rows.Length > 0 Then
                For Each ROW As DataRow In rows
                    rowWHTUPCL1 = dst.Tables("WHTUPCLR").NewRow
                    With rowWHTUPCL1
                        .Item("LBL_REQ_NO") = ASCMAIN1.Next_Control_No("WHTUPCL1.LBL_REQ_NO")
                        .Item("UPC_CODE") = ROW.Item("UPC_CODE")
                        .Item("STYLE_CODE") = ROW.Item("STYLE_CODE")
                        .Item("COLOR_CODE") = ROW.Item("COLOR_CODE")
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("LBL_QTY") = ROW.Item("LBL_QTY") + 1
                        .Item("LOCATION_CODE") = ROW.Item("LOCATION_CODE")
                    End With
                    dst.Tables("WHTUPCLR").Rows.Add(rowWHTUPCL1.ItemArray)
                Next
            End If
        Else
            btnLoad.Text = "Load"
            txtLAisle.Value = ""
            grdWHTUPCLX.Selected.Rows.Clear()
            dst.Tables("WHTUPCLR").Rows.Clear()
        End If
        dst.AcceptChanges()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click
        For Each row As DataRow In dst.Tables("WHTUPCLR").Select("SEL='1'")
            With row
                rowWHTUPCL1 = dst.Tables("WHTUPCL1").NewRow
                rowWHTUPCL1.Item("LBL_REQ_NO") = .Item("LBL_REQ_NO")
                rowWHTUPCL1.Item("UPC_CODE") = .Item("UPC_CODE")
                rowWHTUPCL1.Item("STYLE_CODE") = .Item("STYLE_CODE")
                rowWHTUPCL1.Item("COLOR_CODE") = .Item("COLOR_CODE")
                rowWHTUPCL1.Item("INIT_OPER") = .Item("INIT_OPER")
                rowWHTUPCL1.Item("INIT_DATE") = .Item("INIT_DATE")
                rowWHTUPCL1.Item("LBL_QTY") = .Item("LBL_QTY")
                rowWHTUPCL1.Item("LOCATION_CODE") = .Item("LOCATION_CODE")
                rowWHTUPCL1.Item("PROCESS_IND") = "1"
                rowWHTUPCL1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowWHTUPCL1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            End With
            If row.Item("LBL_QTY") > 0 Then
                printreq(row)
            End If
            dst.Tables("WHTUPCL1").Rows.Add(rowWHTUPCL1.ItemArray)
        Next
        grdWHTUPCLX.Selected.Rows.Clear()
        dst.Tables("WHTUPCLR").Rows.Clear()
        Update_Record()
        btnLoad.Text = "Load"
    End Sub

    'Public Sub SetupLabelWriterSelection(ByVal InitCmb As Boolean)
    '    Dim PrtNames As String
    '    Dim i As Integer

    '    ' get the objects on the label
    '    If InitCmb Then
    '        ' clear all items first
    '        LabelWriterCmb.Items.Clear()

    '        PrtNames = DymoAddIn.GetDymoPrinters()

    '        If Not (PrtNames Is Nothing) Then
    '            ' parse the result
    '            i = PrtNames.IndexOf("|")
    '            While i >= 0
    '                LabelWriterCmb.Items.Add(PrtNames.Substring(0, i))
    '                PrtNames = PrtNames.Remove(0, i + 1)
    '                i = PrtNames.IndexOf("|")
    '            End While
    '            If PrtNames.Length > 0 Then
    '                LabelWriterCmb.Items.Add(PrtNames)
    '            End If

    '            PrtNames = DymoAddIn.GetCurrentPrinterName()
    '            If Not (PrtNames Is Nothing) Then
    '                LabelWriterCmb.SelectedIndex = LabelWriterCmb.Items.IndexOf(PrtNames)
    '            Else
    '                LabelWriterCmb.SelectedIndex = 0
    '            End If
    '        End If
    '    End If

    ' check if selected/current printer is a twin turbo printer
    'TrayCmb.Enabled = DymoAddIn.IsTwinTurboPrinter(LabelWriterCmb.Text)
    'If TrayCmb.Enabled Then
    '    ' show the current tray selection if the printer
    '    ' is a twin turbo
    '    i = DymoAddIn.GetCurrentPaperTray()
    '    If i = 0 Then
    '        TrayCmb.SelectedIndex = 0 ' left tray
    '    ElseIf i = 1 Then
    '        TrayCmb.SelectedIndex = 1 ' right tray
    '    ElseIf i = 2 Then
    '        TrayCmb.SelectedIndex = 2 ' auto switch
    '    Else
    '        TrayCmb.SelectedIndex = 2 ' tray selection not set, so default to auto switch
    '    End If
    'End If
    'End Sub

    'Private Sub LabelWriterCmb_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles LabelWriterCmb.SelectionChangeCommitted
    '    DymoAddIn.SelectPrinter(LabelWriterCmb.Text)
    '    SetupLabelWriterSelection(False)
    'End Sub

End Class