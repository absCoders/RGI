Public Class SOTTOTE0

    Dim zplPrint As New TAC.TACZPLT1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from SOTTOTE1 where WHSE_CODE = :PARM1 and TOTE_CLASS_CODE = :PARM2 AND NVL(TOTE_TYPE,'?') <> 'P' AND NVL(TOTE_TYPE,'?') <> 'C'"
            Create_TDA(.Tables.Add, "SOTTOTE1", "**", 0,, "VV", 1, "TOTE_LABEL_PRINT_IND")
            .Tables("SOTTOTE1").Columns.Add("SEL")
            .Tables("SOTTOTE1").Columns("SEL").DefaultValue = "0"
            .Tables("SOTTOTE1").Columns.Add("PRT")
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

        Show_Filter(grdSOTTOTE1, True)
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

            Case "Edit"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Show_Record_Special()

        'EnforceConstraints(False)
        'Fill_Records("SOTTOTE1", New String() {Absx1.txtFor("MATL_CODE").Text})
        'EnforceConstraints(True)

        If EntryMode = "Edit" And Absx1.txtFor("TOTE_CLASS_CODE").Text <> "X" Then
            grpToteManagement.Visible = True
            grdSOTTOTE1.Visible = False
            Set_Read_Only(grpToteManagement, False)
            Set_Read_Only_for_ctl(txtWHSE_CODE, False)
            txtWHSE_CODE.Enabled = True
        Else
            grpToteManagement.Visible = False
            grdSOTTOTE1.Visible = False
        End If

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTTOTE1").Rows.Clear()
            EnforceConstraints(False)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then
            grdSOTTOTE1.Visible = False
            grpCreateNewTotes.Visible = False
            grpPrintToteLabels.Visible = False
            grpToteManagement.Visible = False

            grdSOTTOTE1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        End If
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Update_Record_TDA("SOTTOTE1")
    End Sub

    Private Sub txtWHSE_CODE_Leave(sender As Object, e As EventArgs) Handles txtWHSE_CODE.Leave

    End Sub

    Private Sub txtWHSE_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtWHSE_CODE.ValueChanged
        Dim WHSE_CODE As String = txtWHSE_CODE.Text
        Dim TOTE_CLASS_CODE As String = Absx1.txtFor("TOTE_CLASS_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 IsNot Nothing Then
            Fill_Records("SOTTOTE1", New String() {WHSE_CODE, TOTE_CLASS_CODE})
            Sort_grdColumns(grdSOTTOTE1, "TOTE_NO")
            grdSOTTOTE1.Visible = True
            grpCreateNewTotes.Visible = True
            grpPrintToteLabels.Visible = True

            grdSOTTOTE1.Text = $"Class {Absx1.txtFor("TOTE_CLASS_CODE").Text} Totes in DC {WHSE_CODE}"

        Else
            grdSOTTOTE1.Visible = False
            grpCreateNewTotes.Visible = False
            grpPrintToteLabels.Visible = False
        End If
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
    '    Generate_Report("SORTOTEL", "Bar-Coded Labels for Totes")
    '    Print_Report_End()
    'End Sub


    Private Sub btnCreateTotes_Click(sender As Object, e As EventArgs) Handles btnCreateTotes.Click

        Dim N As Integer = Val(numCreateQty.Value)

        If N >= 1 Then

            If MsgBox($"OK to Create {CStr(N)} New Totes?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

            Dim TOTE_NO_ctr As Integer = Val(ASCMAIN1.Next_Control_No("SOTTOTE1.TOTE_NO", N)) - 1

            Dim TOTE_CLASS_CODE As String = Absx1.txtFor("TOTE_CLASS_CODE").Text
            Dim WHSE_CODE As String = txtWHSE_CODE.Value
            For i As Integer = 1 To N
                TOTE_NO_ctr += 1
                Dim TOTE_NO As String = Format(TOTE_NO_ctr, "000000")
                Dim rowSOTTOTE1 As DataRow = dst.Tables("SOTTOTE1").NewRow
                rowSOTTOTE1.Item("TOTE_NO") = TOTE_NO
                rowSOTTOTE1.Item("TOTE_CLASS_CODE") = TOTE_CLASS_CODE
                rowSOTTOTE1.Item("WHSE_CODE") = WHSE_CODE
                rowSOTTOTE1.Item("TOTE_TYPE") = "R"
                rowSOTTOTE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTTOTE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTTOTE1.Item("INIT_DATE") = Me.DATETIME_STAMP
                rowSOTTOTE1.Item("LAST_DATE") = Me.DATETIME_STAMP
                dst.Tables("SOTTOTE1").Rows.Add(rowSOTTOTE1)
            Next

            Update_Record_TDA("SOTTOTE1")

            MsgBox($"{CStr(N)} New Totes have been Created - Remember to Print Labels", vbOKOnly, "Success")
        End If
    End Sub

#End Region

End Class