Public Class SOTTRCK1

    Dim zplPrint As New TAC.TACZPLT1

    Private lstFramesDC As New List(Of String)
    Private lstStockLensDC As New List(Of String)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from SOTTOTE1 where TRUCK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTTOTE1", "**", 0, True, "V", 1)
            .Tables("SOTTOTE1").Columns.Add("SEL")
            .Tables("SOTTOTE1").Columns("SEL").DefaultValue = "0"
            .Tables("SOTTOTE1").Columns.Add("PRT")

            'lstFramesDC = TAC.TACMAIN1.GetDCCodes(TAC.TACMAIN1.DCTypes.Frames)
            'lstStockLensDC = TAC.TACMAIN1.GetDCCodes(TAC.TACMAIN1.DCTypes.StockLenses)

        End With

        grdSOTTOTE1.DataSource = dst.Tables("SOTTOTE1")
        With grdSOTTOTE1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With
        With grdSOTTOTE1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                If GCOL.Key = "SEL" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        'Show_Filter(grdSOTTOTE1, True)

        Set_Read_Only_for_ctl(optTRUCK_TYPE, True)
        Set_Read_Only_for_ctl(txtMini, True)
        txtMini.Text = ASCMAIN1.MiniLabelPrinterIPAddress

        Create_Summary(grdSOTTOTE1, "TOTE_NO", "Count")
        Create_Summary(grdSOTTOTE1, "TOTE_LABEL_PRINT_IND", "Sum")
        Create_Summary(grdSOTTOTE1, "SEL", "Sum")
    End Sub

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTTOTE1, "BB", "Select All", "De-Select All")
    End Sub

    Private Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles tlb.BeforeToolDropdown

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool
            If tlb_pop.Tools.Exists("Show Filter") Then
                tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            End If

            If tlb_pop.Tools.Exists("Show GroupBox") Then
                tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
            End If

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Private Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles tlb.ToolClick
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each rowSOTTOTE1 As DataRow In dst.Tables("SOTTOTE1").Select("")
                    rowSOTTOTE1.Item("SEL") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next

        End Select
    End Sub

#End Region


#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

                Absx1.txtFor("TRUCK_NO").Text = Absx1.txtFor("TRUCK_NO").Text.Trim.ToUpper
                Dim TRUCK_NO As String = Absx1.txtFor("TRUCK_NO").Text
                If TRUCK_NO.Length <> 4 OrElse Not (TRUCK_NO.StartsWith("T") Or TRUCK_NO.StartsWith("N") Or TRUCK_NO.StartsWith("K")) OrElse Not IsNumeric(TRUCK_NO.Substring(1)) Then
                    EMsg &= vbCr & "Truck ID must be 4 characters starting with T, N or K followed by 3 numeric values"
                End If

            Case "Edit"

            Case "Update"

                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                Dim TRUCK_NO As String = Absx1.txtFor("TRUCK_NO").Text

                If WHSE_CODE = "" Then
                    EMsg &= vbCr & "DC Code is Mandatory"
                Else
                    If LookUp("ICTWHSE1", WHSE_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid value specificed for DC Code"
                    Else
                        'If lstStockLensDC.Contains(WHSE_CODE) AndAlso TRUCK_NO.StartsWith("N") Then
                        '    ' This is Okay
                        'ElseIf lstFramesDC.Contains(WHSE_CODE) AndAlso (TRUCK_NO.StartsWith("T") OrElse TRUCK_NO.StartsWith("K")) Then
                        '    ' This is Okay
                        'Else
                        '    EMsg &= vbCr & "Invalid Truck No for DC specified"
                        'End If

                    End If
                End If

                If EntryMode = "New" And EMsg = "" Then
                    If optTRUCK_TYPE.Value & "" = "" Then
                        optTRUCK_TYPE.Value = "R"
                        dst.Tables("SOTTRCK1").Rows(0).Item("TRUCK_TYPE") = "R"
                    End If

                    If optTRUCK_TYPE.Value = "P" Then
                        If dst.Tables("SOTTOTE1").Rows.Count = 0 Then
                            EMsg &= vbCr & "No Totes defined for Pre-Configured Truck"
                        Else
                            For Each rowSOTTOTE1 As DataRow In dst.Tables("SOTTOTE1").Select("")
                                rowSOTTOTE1.Item("WHSE_CODE") = WHSE_CODE
                            Next
                        End If
                    End If

                    If optTRUCK_TYPE.Value = "X" Then
                        EMsg &= vbCr & "You cannot set Custom Truck here - this is done when Building a Truck"
                    End If
                End If

                ' CHECK THAT IF THERE ARE TOTES THAT THE PRECONFIG FLAG IS SET, AND VICE VERSA
                If EMsg.Length = 0 Then
                    If optTRUCK_TYPE.CheckedIndex = -1 OrElse dst.Tables("SOTTRCK1").Rows(0).Item("TRUCK_TYPE") & String.Empty = String.Empty Then
                        EMsg& = vbCr & "Truck Type is required."
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Show_Record_Special()
        Dim TRUCK_NO As String = Absx1.txtFor("TRUCK_NO").Text
        Fill_Records("SOTTOTE1", TRUCK_NO)
        Sort_grdColumns(grdSOTTOTE1, "SLOT_NO")
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTTOTE1").Rows.Clear()
            EnforceConstraints(False)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        btnPrintTruckPlacard.Visible = tf
        If EntryMode = "New" Then
            Set_Read_Only_for_ctl(optTRUCK_TYPE, False)
            Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), False)
        Else
            Set_Read_Only_for_ctl(optTRUCK_TYPE, True)
            Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), True)
        End If

        Toggle_PreConfig()
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        If EntryMode = "New" Then
            Update_Record_TDA("SOTTOTE1")
        End If
        ' ALLOW UNCHECK PRECONFIGURED TO WIPE OUT TOTES FIELDS FOR TRUCK AND SLOT
    End Sub

#End Region

    Sub Toggle_PreConfig()
        Dim isPreConfig As Boolean = (optTRUCK_TYPE.Value & "" = "P")
        grdSOTTOTE1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        numTotes.Visible = isPreConfig AndAlso EntryMode = "New"
        lblTotes.Visible = isPreConfig AndAlso EntryMode = "New"

        ' tried to get this to work with View but grid did not permit ckecking Sel
        splTote1.Visible = isPreConfig
        'grdSOTTOTE1.Visible = isPreConfig
        'grpPrintToteLabels.Visible = isPreConfig AndAlso (EntryMode = "Edit" OrElse EntryMode = "View")

        If isPreConfig Then
            'Set_Read_Only_for_ctl(grdSOTTOTE1, False)
            grdSOTTOTE1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

        Set_Read_Only_for_ctl(txtMini, True)

    End Sub

    Private Sub btnPrintToteLabels_Click(sender As Object, e As EventArgs) Handles btnPrintToteLabels.Click
        Dim sql As String = "ISNULL(TOTE_LABEL_PRINT_IND,'0') = '0'"
        Dim C As Integer = Queue_Label_Print(sql)
        If C > 0 Then
            Print_Tote_Labels()
            'Print_Report()
            If MsgBox("OK to mark these labels as Printed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                For Each rowSOTTOTE1 As DataRow In dst.Tables("SOTTOTE1").Select("PRT = '1'")
                    rowSOTTOTE1.Item("TOTE_LABEL_PRINT_IND") = "1"
                Next
                Update_Record_TDA("SOTTOTE1")
            End If
        End If
    End Sub

    Private Sub btnPrintSelected_Click(sender As Object, e As EventArgs) Handles btnPrintSelected.Click
        Dim sql As String = "ISNULL(SEL,'0') = '1'"
        Dim C As Integer = Queue_Label_Print(sql)
        If C > 0 Then
            Print_Tote_Labels()
            'Print_Report()
        End If
    End Sub

    Sub Print_Tote_Labels()

        If ASCMAIN1.MiniLabelPrinterIPAddress.Length = 0 Then
            MessageBox.Show("You are not assigned a Mini Label IP Address", "Print Tote Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Tote Labels")

        zplPrint.Print_Tote_Labels(dst.Tables("SOTTOTE1").Select("PRT = '1'", "TOTE_NO"), chkOneTotePerLabel.Checked)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Queue_Label_Print(sql As String) As Integer
        Dim C As Integer = 0
        For Each rowSOTTOTE1 As DataRow In dst.Tables("SOTTOTE1").Select("")
            rowSOTTOTE1.Item("PRT") = "0"
        Next
        For Each rowSOTTOTE1 As DataRow In dst.Tables("SOTTOTE1").Select(sql)
            rowSOTTOTE1.Item("PRT") = "1"
            C += 1
        Next
        If C = 0 Then
            MsgBox("Nothing to Print", MsgBoxStyle.OkOnly, "Cannot Print Labels")
        End If
        Return C
    End Function


    'Sub Print_Report()
    '    Print_Report_Begin()
    '    Generate_Report("SORTRCKL", "Placard for Truck")
    '    Print_Report_End()
    'End Sub

    Private Sub numTotes_ValueChanged(sender As Object, e As EventArgs) Handles numTotes.ValueChanged

    End Sub

    Private Sub numTotes_KeyDown(sender As Object, e As KeyEventArgs) Handles numTotes.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "New" Then
            Dim n As Integer = Val(numTotes.Value & "")

            If n > 0 And n < 45 Then
                dst.Tables("SOTTOTE1").Rows.Clear()
                Dim TRUCK_NO As String = Absx1.txtFor("TRUCK_NO").Text
                For T As Integer = 1 To n
                    Dim TOTE_NO As String = TRUCK_NO & Format(T, "00")
                    Dim rowSOTTOTE1 As DataRow = dst.Tables("SOTTOTE1").NewRow
                    With rowSOTTOTE1
                        .Item("TOTE_NO") = TOTE_NO
                        .Item("TOTE_CLASS_CODE") = "A"
                        .Item("WHSE_CODE") = Absx1.txtFor("WHSE_CODE").Text
                        .Item("TRUCK_NO") = TRUCK_NO
                        .Item("SLOT_NO") = T
                        .Item("TOTE_TYPE") = "P"
                    End With
                    dst.Tables("SOTTOTE1").Rows.Add(rowSOTTOTE1)
                Next
            End If
        End If
    End Sub

    Private Sub numTotes_KeyPress(sender As Object, e As KeyPressEventArgs) Handles numTotes.KeyPress

    End Sub

    Private Sub optTRUCK_TYPE_ValueChanged(sender As Object, e As EventArgs) Handles optTRUCK_TYPE.ValueChanged
        Toggle_PreConfig()
        If Not (optTRUCK_TYPE.Value & "" = "P") Then
            dst.Tables("SOTTOTE1").Rows.Clear()
        End If
    End Sub

    Private Sub btnPrintTruckPlacard_Click(sender As Object, e As EventArgs) Handles btnPrintTruckPlacard.Click
        Dim TRUCK_NO As String = Absx1.txtFor("TRUCK_NO").Text
        zplPrint.Print_Truck_ID(TRUCK_NO)
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        If ASCMAIN1.MiniLabelPrinterIPAddress.Length = 0 Then
            MessageBox.Show("You are not assigned a Mini Label IP Address", "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        zplPrint.PrintItemLabel("Test", 1)
    End Sub

    Private Sub UltraTextEditor10_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor10.ValueChanged

    End Sub
End Class