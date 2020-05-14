Public Class WHFPCKT1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("WHTPARM1")

        With dst
            Create_TDA(.Tables.Add, "WHTUPCL1", "*")

            ASCMAIN1.sql = "Select WHTUPCL1.* from WHTUPCL1 where PROCESS_IND = '1'"
            Create_TDA(.Tables.Add, "WHTPKTX", "**", 0, False, "", 1)


        End With

        grdWHTPKTX.DataSource = dst.Tables("WHTPKTX")



        Dim rows() As DataRow = ASCDATA1.GetDataTable("SELECT *  FROM WHTLPRT1").Select("")
        For Each row As DataRow In rows
            cbxLabelPrinter.Items.Add(row.Item("LABEL_PRINTER_ID"))
        Next
        cbxLabelPrinter.SelectedIndex = 0

        Create_Summary(grdWHTPKTX, "LBL_REQ_NO", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("UPC_CODE")
                Absx1.txtFor("STYLE_CODE").Text = cdr.Item("STYLE_CODE")
                Absx1.txtFor("COLOR_CODE").Text = cdr.Item("COLOR_CODE")
                ' MULTITASKING

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

        '      Set_Read_Only(UltraGroupBox2, True)

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTUPCL1", "WHTUPCLP", "WHTPKTX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()


    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            'rowWHTUPCL1 = dst.Tables("WHTUPCL1").NewRow
            'With rowWHTUPCL1
            '    .Item("LBL_REQ_NO") = ASCMAIN1.Next_Control_No("WHTUPCL1.LBL_REQ_NO")
            '    .Item("UPC_CODE") = HFs("UPC_CODE")
            '    .Item("STYLE_CODE") = HFs("STYLE_CODE")
            '    .Item("COLOR_CODE") = HFs("COLOR_CODE")
            '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
            '    .Item("INIT_DATE") = DATETIME_STAMP
            '    .Item("LBL_QTY") = 1
            'End With

            ' dst.Tables("WHTUPCL1").Rows.Add(rowWHTUPCL1.ItemArray)
        Else
            'rowWHTUPCL1 = Fill_Record("WHTUPCL1", New String() {Absx1.txtFor("LBL_REQ_NO").Text})
            'dst.AcceptChanges()
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
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    ' Click_Command("New", e)
                    ' SHOW PICK TICKETS FOR CUSTOMER SELECTED
                End If
            Case "PICK_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                '   Click_Command("New")

            Case "PICK_NO"
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
        Load_Popup_Menu(grdWHTPKTX, "B", "Re-Print")
        
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

            
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

           
        End Select
    End Sub

#End Region

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor

        Fill_Records("WHTPKTX")
        Sort_grdColumns(grdWHTPKTX, "PICK_NO")
        grdWHTPKTX.Text = "Pick Tickets"

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdWHTPKTX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTPKTX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("LBL_REQ_NO").Text = e.Row.Cells("LBL_REQ_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Display_Totals()

    End Sub
End Class