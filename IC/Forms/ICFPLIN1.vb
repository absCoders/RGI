Public Class ICFPLIN1

    Dim STYLE_CODE_PLM As String
    Dim SEASON_CODEs As New List(Of String)
    Dim rowICTPLIN2 As DataRow
    Dim auto_generated_style As Boolean = False
    Dim imgba() As Byte = Nothing

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        InquiryMode = (MENU_ITEM_OBJECT = "ICFPLINI")

        With dst

            Create_TDA(.Tables.Add, "ICTPLIN1", "*")

            ASCMAIN1.sql = "Select * from ICTPLIN2"
            Create_TDA(.Tables.Add, "ICTPLINX", "**", 0, False)

            ASCMAIN1.sql = "Select STYLE_CODE_PLM from ICTPLIN2 where SALES_DIVISION_CODE = :PARM1 and STYLE_CLASS_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTPLINN", "**", 0, False, "VV")

            Create_TDA(.Tables.Add, "ICTPLIN2", "*", 1)
            .Tables("ICTPLIN2").Columns.Add("IMAGE", GetType(System.Byte()))
       
            Create_TDA(.Tables.Add, "ICTPLIN3", "*", 1)
            Create_Relation("ICTPLIN2", "ICTPLIN3", "STYLE_CODE_PLM")

            With .Tables("ICTPLIN3").Columns
                 .Add("DUTY_COST_CALC", GetType(System.Decimal), "(ISNULL(PO_COST,0) + ISNULL(OTHER_COST,0)) * ISNULL(DUTY_RATE,0) / 100")
                '.Add("LANDED_COST_CALC", GetType(System.Decimal), "ISNULL(PO_COST,0) + ISNULL(OTHER_COST,0) + ISNULL(DUTY_COST,0) + ISNULL(FREIGHT_COST,0) + ISNULL(BRKR_COST,0) + ISNULL(MISC_COST,0) + ISNULL(INLAND_COST,0) + ISNULL(LABOR_COST,0)")
                .Add("LANDED_COST_CALC", GetType(System.Decimal), "ISNULL(PO_COST,0) + ISNULL(OTHER_COST,0) + ISNULL(DUTY_COST_CALC,0) + ISNULL(FREIGHT_COST,0) + ISNULL(BRKR_COST,0) + ISNULL(MISC_COST,0) + ISNULL(INLAND_COST,0) + ISNULL(LABOR_COST,0)")
                .Add("STYLE_PRICE_CALC", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "ICTPLIN4", "*", 1)

            For Each TABLE_NAME As String In New String() {"ICTCLAS1", "ICTSEAS1", "SOTSDIV1", "APTVEND1", "TATCNTRY"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                If TABLE_NAME = "APTVEND1" Then ASCMAIN1.sql &= " where VEND_TYPE = 'S'"
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False)
                Fill_Records(TABLE_NAME)
            Next

        End With

        grdICTPLINX.DataSource = dst.Tables("ICTPLINX")
        grdICTPLIN3.DataSource = dst.Tables("ICTPLIN3")
        grdICTPLIN4.DataSource = dst.Tables("ICTPLIN4")
        grdICTPLINN.DataSource = dst.Tables("ICTPLINN")

        Create_Summary(grdICTPLINX, "STYLE_CODE_PLM", "Count")

        With grdICTPLIN4.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"STYLE_DTL_SEQ", "STYLE_DTL_QTY", "STYLE_DTL_COLOR", "STYLE_DTL_PANTONE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightGray
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "STYLE_DTL_SEQ" Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next

            Dim colors() As System.Drawing.Color = {Color.Orange, _
                                                    Color.LightBlue, _
                                                    Color.Yellow, _
                                                    Color.LightGreen, _
                                                    Color.Gold, _
                                                    Color.Pink, _
                                                    Color.Azure, _
                                                    Color.Gold}

            For i As Integer = 1 To 8

                .Columns("COLOR_" & CStr(i)).Width = 70
                .Columns("PANTONE_" & CStr(i)).Width = 70
                .Columns("COLOR_" & CStr(i)).Header.Caption = "Color" & CStr(i)
                .Columns("PANTONE_" & CStr(i)).Header.Caption = "Pant#" & CStr(i)
                .Columns("COLOR_" & CStr(i)).Header.Appearance.BackColor2 = colors(i - 1)
                .Columns("PANTONE_" & CStr(i)).Header.Appearance.BackColor2 = colors(i - 1)
            Next
        End With

        With grdICTPLIN3.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"STYLE_TYPE_SEQ", "DUTY_CATGY_CODE", "STYLE_CONTENT", "STYLE_WEIGHT", "STYLE_SPEC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "STYLE_TYPE_SEQ" Or gcol.Key = "STYLE_PRICE_CALC" Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightGray
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("LANDED_COST_CALC").CellAppearance.BackColor = Color.LightGray
            .Columns("OTHER_COST").Hidden = True

            .Columns("BRKR_COST").Hidden = True
            .Columns("MISC_COST").Hidden = True
            .Columns("INLAND_COST").Hidden = True
            .Columns("LABOR_COST").Hidden = True
        End With

        With grdICTPLINX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        Season(Now.AddMonths(18))
        chkSNNext.Text = Season(Now.AddMonths(12))
        chkSNThis.Text = Season(Now.AddMonths(6))
        chkSNLast.Text = Season(Now.AddMonths(0))
        chkSNPrev.Text = Season(Now.AddMonths(-6)) & " and Prior"

        Show_Filter(grdICTPLINX, True)

    End Sub

    Function Season(dt As Date) As String
        Dim YYYY As String = Format(dt, "yyyy")
        Dim MM As String = Format(dt, "MM")
        Dim SN As String = IIf(MM >= "02" And MM <= "07", "S", "F")

        SEASON_CODEs.Add(YYYY & SN)

        Return YYYY & SN

    End Function

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                STYLE_CODE_PLM = Absx1.txtFor("STYLE_CODE_PLM").Text
                If STYLE_CODE_PLM = "" Then

                    EMsg &= vbCr & "You Must first Enter a Style Code"
                End If
                Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                If rowICTPLIN2 IsNot Nothing Then
                    EMsg &= vbCr & "Record for Style " & STYLE_CODE_PLM & " already Exists"
                End If

                ASCMAIN1.sql = "Select STYLE_CODE from ICTSTYL1 where STYLE_CODE like :PARM1 || '%'"
                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", STYLE_CODE_PLM)
                If row IsNot Nothing Then
                    If ASCMAIN1.USER_ID = "rcohen" Or ASCMAIN1.USER_ID = "wjz" Then
                        Dim msg As String = "A Style Master Record for Style " & STYLE_CODE_PLM & " already Exists (see " & row.Item(0) & ")"
                        msg &= vbCrLf & vbCrLf & "You are allowed to continue because you are an Administrator" & vbCrLf & "You should be creating a Parent PLM for Orphaned ERP Styles" & vbCrLf & " which begin with " & STYLE_CODE_PLM
                        MsgBox(msg, MsgBoxStyle.OkOnly, "Warning - Please Proceed with Caution")
                    Else
                        EMsg &= vbCr & "A Style Master Record for Style " & STYLE_CODE_PLM & " already Exists (see " & row.Item(0) & ")"
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("STYLE_CODE_PLM").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Edit", "View"

                STYLE_CODE_PLM = Absx1.txtFor("STYLE_CODE_PLM").Text
                Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                If rowICTPLIN2 Is Nothing Then
                    EMsg &= vbCr & "No record of Style " & STYLE_CODE_PLM
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("STYLE_CODE_PLM").Text) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Update"

                If Absx1.txtFor("STYLE_DESC").Text = "" Then
                    EMsg &= vbCr & "Style Description is Mandatory"
                End If

                cdr = LookUp("SOTSDIV1", Absx1.txtFor("SALES_DIVISION_CODE").Text)
                If cdr Is Nothing Then
                    EMsg &= vbCr & "Invalid Sales Division Specified"
                End If
                cdr = LookUp("ICTSEAS1", Absx1.txtFor("SEASON_CODE").Text)
                If cdr Is Nothing Then
                    EMsg &= vbCr & "Invalid Season Specified"
                End If
                cdr = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                If cdr Is Nothing Then
                    ' EMsg &= vbCr & "Invalid Supplier Specified"
                End If

            Case "Cancel"
                If MsgBox("OK to Lose any Changes Made to Style " & STYLE_CODE_PLM, _
                        MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

            Case "Delete"
                If MsgBox("OK to Delete all information relating to Style " & STYLE_CODE_PLM, _
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

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

                'If ASCMAIN1.Running_in_VS Then
                '    Stop
                '    ASCMAIN1.sql = "Select * from ICTCLAS1"
                '    For Each row1 As DataRow In ASCDATA1.GetDataTable.Select("")
                '        Dim STYLE_CLASS_CODE As String = row1.Item("STYLE_CLASS_CODE")
                '        ASCMAIN1.Progress(STYLE_CLASS_CODE, "")
                '        ASCMAIN1.sql = "Select * from ICTPLIN2 where STYLE_CODE LIKE '" & STYLE_CLASS_CODE & "-%'"
                '        Dim SEQ As Integer = 0
                '        For Each ROW2 As DataRow In ASCDATA1.GetDataTable.Select("")
                '            Dim STYLE_CODE As String = ROW2.Item("STYLE_CODE_PLM")
                '            Dim SSEQ As String = Split(STYLE_CODE, "-")(1)
                '            If IsNumeric(SSEQ) Then
                '                If Val(SSEQ) > SEQ Then SEQ = Val(SSEQ)
                '            End If
                '        Next
                '        ASCMAIN1.sql = "Update ICTCLAS1 SET STYLE_CLASS_STYLE_SEQ = " & CStr(SEQ) & " WHERE STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'"
                '        ASCDATA1.ExecuteSQL()
                '    Next
                '    ASCMAIN1.Progress("", "")
                '    MsgBox("Done")
                '    Exit Sub
                'End If

                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Done"
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Print"
                Print_Record()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If

                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode

                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                .Items("Print").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode

                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode And Not InquiryMode
                .Items("Delete").Visible = (EntryMode = "E") And Not InquiryMode And Not InquiryMode
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode And Not InquiryMode
            End With

            .Groups("Scope").Visible = Not ScreenMode
            .Groups("Navigator").Visible = (EntryMode = "V")
            .Groups("New Style").Visible = Not ScreenMode And Not InquiryMode
            .Groups("Copy Style to ...").Visible = ScreenMode And (EntryMode = "V")
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdICTPLINX.Visible = Not ScreenMode
        splStyle.Visible = ScreenMode

        If ScreenMode Then

            If EntryMode = "V" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPLIN3, grdICTPLIN4}
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                Next

            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPLIN3, grdICTPLIN4}
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                Next
            End If

            Set_Read_Only(splNotes.Panel1, (EntryMode = "V"))
            Set_Read_Only(splNotes.Panel2, (EntryMode = "V"))
            Set_Read_Only(grpStyle, (EntryMode = "V"))
            Set_Read_Only(grpStyleImage, (EntryMode = "V"))
            Set_Read_Only(Absx1.txtFor("STYLE_DESC"), (EntryMode = "V"))

            ' Set_Read_Only_for_ctl(Absx1.txtFor("STYLE_CLASS_CODE"), EntryMode <> "N" Or auto_generated_style)
            ' Set_Read_Only_for_ctl(Absx1.txtFor("SEASON_CODE"), EntryMode <> "N" Or auto_generated_style)
            ' Set_Read_Only_for_ctl(Absx1.txtFor("SALES_DIVISION_CODE"), EntryMode <> "N" Or auto_generated_style)
            ' Set_Read_Only_for_ctl(Absx1.txtFor("ROYALTY_CODE"), EntryMode <> "N" Or auto_generated_style)

            With grdICTPLIN3.DisplayLayout.Bands(0)
                .Columns("INIT_DATE").Hidden = Not (EntryMode = "V")
                .Columns("INIT_OPER").Hidden = Not (EntryMode = "V")
                .Columns("LAST_DATE").Hidden = Not (EntryMode = "V")
                .Columns("LAST_OPER").Hidden = Not (EntryMode = "V")
            End With
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTPLIN1", "ICTPLIN2", "ICTPLIN3", "ICTPLIN4"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        'If HFs.Count > 0 Then
        '    Absx1.txtFor("SEASON_CODE").Text = HFs("SEASON_CODE")
        '    Absx1.txtFor("SALES_DIVISION_CODE").Text = HFs("SALES_DIVISION_CODE")
        'End If
        Load_ICTPLINX()

        auto_generated_style = False
        txtStyleCode_CopyTo.Text = ""

        Setup_Tab()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        rowICTPLIN2 = Fill_Record("ICTPLIN2", STYLE_CODE_PLM, EntryMode = "N")
        If EntryMode = "N" Then
            rowICTPLIN2.Item("STYLE_DESC") = HFs("STYLE_DESC")
        Else
            Save_Header_Fields(UltraGroupBox1)
        End If
        Fill_Records("ICTPLIN3", STYLE_CODE_PLM)

        Dim ROYALTY_PCT As Decimal = Val(Absx1.numFor("ROYALTY_PCT").Value & "")
        For Each rowICTPLIN3 As DataRow In dst.Tables("ICTPLIN3").Select("")
            rowICTPLIN3.Item("STYLE_PRICE_CALC") = TAC.ICCMAIN1.Calculate_Suggested_SP(Val(rowICTPLIN3.Item("LANDED_COST") & ""), ROYALTY_PCT)
        Next
        Fill_Records("ICTPLIN4", STYLE_CODE_PLM)

        Sort_grdColumns(grdICTPLIN3, "STYLE_TYPE_SEQ")
        Sort_grdColumns(grdICTPLIN4, "STYLE_DTL_SEQ")

        EnforceConstraints(True)

        If EntryMode = "V" Then
            Fill_Records("ICTPLINN", New String() {Absx1.txtFor("SALES_DIVISION_CODE").Text, _
                                                   Absx1.txtFor("STYLE_CLASS_CODE").Text})
            Sort_grdColumns(grdICTPLINN, "STYLE_CODE_PLM")
            For Each grow As UltraWinGrid.UltraGridRow In grdICTPLINN.Rows
                If grow.Cells("STYLE_CODE_PLM").Value & "" = STYLE_CODE_PLM Then
                    grdICTPLINN.ActiveRow = grow
                    Exit For
                End If
            Next
        End If

        imgba = Nothing

        Dim IMAGE_NAME As String = STYLE_CODE_PLM & ".jpg"
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & "\" & Absx1.txtFor("SALES_DIVISION_CODE").Text
        If My.Computer.FileSystem.FileExists(FOLDER_NAME & "\" & STYLE_CODE_PLM & ".png") Then
            IMAGE_NAME = STYLE_CODE_PLM & ".png"
        End If
        'If ASCMAIN1.Running_in_VS Then FOLDER_NAME = "C:\Users\wjz\Desktop\Data\Database\Images\" & Absx1.txtFor("SALES_DIVISION_CODE").Text
        imgSTYLE.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)

        'If IMAGE_NAME <> "" Then
        '    Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & "\" & Absx1.txtFor("SALES_DIVISION_CODE").Text
        '    'If ASCMAIN1.Running_in_VS Then FOLDER_NAME = "C:\Users\wjz\Desktop\Data\Database\Images\" & Absx1.txtFor("SALES_DIVISION_CODE").Text
        '    imgSTYLE.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
        'Else
        '    imgSTYLE.Image = Nothing
        'End If


        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Setup_Tab()
    End Sub

    Sub Delete_Record()
        BeginTrans()

        For Each TABLE_NAME As String In New String() {"ICTPLIN2", "ICTPLIN3", "ICTPLIN4"}
            ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where STYLE_CODE_PLM = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", STYLE_CODE_PLM)
        Next

        CommitTrans("Delete Complete")
    End Sub

    Sub Update_Record()
        INIT_LAST("ICTPLIN2", True, "", True)

        BeginTrans()
        EnforceConstraints(False)
        Update_Record_TDA("ICTPLIN2")

        ' these lines are repeated in Print_Record
        For Each rowICTPLIN3 As DataRow In dst.Tables("ICTPLIN3").Select("")
            rowICTPLIN3.Item("LANDED_COST") = rowICTPLIN3.Item("LANDED_COST_CALC")
            rowICTPLIN3.Item("DUTY_COST") = rowICTPLIN3.Item("DUTY_COST_CALC")
            rowICTPLIN3.Item("STYLE_PRICE") = rowICTPLIN3.Item("STYLE_PRICE_CALC")
        Next

        Update_Record_TDA("ICTPLIN3")
        Update_Record_TDA("ICTPLIN4")
        EnforceConstraints(True)
        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "SEASON_CODE", "ICTSTYL1.SEASON_CODE"
                sql_where = "SEASON_CODE IN ('" & Join(SEASON_CODEs.ToArray, "','") & "')"
            Case "VEND_CODE"
                sql_where = "VEND_TYPE = 'S'"
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("STYLE_CODE_PLM").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ICTPLIN2"
            E.COLUMN_NAME = "STYLE_CODE_PLM"
            E.CODE_VALUE = Absx1.txtFor("STYLE_CODE_PLM").Text
            E.DESC_VALUE = "PLM Style Code"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTPLINX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins")
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

        Select Case e.SourceControl.Name

            'Case "grdTATEVNT1"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Show email"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("EVENT_TYPE").Value = "EML" Or grd.ActiveRow.Cells("EVENT_TYPE").Value = "EML"))

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
            ' NEED TO GET PAST HERE FOR STYLE MULTI-COLOR WHEN THERE ARE NO ROWS IN THE GRID
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                'Case "grdSOTORDRS"

                '    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                '    tlb_btn.SharedProps.Visible = show_qty_copy_option
                '    tlb_btn.SharedProps.Caption = "Update Qty to " & CStr(ORDR_QTY) & " for All Stores"


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

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Style Master File"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE_PLM").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If


        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE_PLM"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not ScreenMode Then Click_Command("View")
                End If
                'Case "SEASON_CODE"
                '    If e.KeyCode = Windows.Forms.Keys.Enter Then
                '        Load_ICTPLINX()
                '    End If
                'Case "SALES_DIVISION_CODE"
                '    If e.KeyCode = Windows.Forms.Keys.Enter Then
                '        Load_ICTPLINX()
                '    End If

            Case "STYLE_CODE_NAV"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim row As DataRow = LookUp("ICTPLIN2", txtSTYLE_CODE_NAV.Text)
                    If row Is Nothing Then
                        MsgBox("No record of Style " & txtSTYLE_CODE_NAV.Text)
                        Exit Sub
                    Else
                        Click_Done_and_View_Style(txtSTYLE_CODE_NAV.Text)
                        txtSTYLE_CODE_NAV.Text = ""
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            'Case "SEASON_CODE"
            '    If Not ScreenMode Then
            '        Load_ICTPLINX()
            '    End If
            'Case "SALES_DIVISION_CODE"
            '    If Not ScreenMode Then
            '        Load_ICTPLINX()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE_PLM"
                Click_Command("View")
                'Case "CUST_CODE"
                '    Load_ICTPLINX()
                'Case "SALES_DIVISION_CODE"
                '    Load_ICTPLINX()
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ROYALTY_PCT"
                If ScreenMode And (EntryMode = "E" Or EntryMode = "N") Then Apply_Royalty()
        End Select

    End Sub

#End Region

    Sub Setup_Tab()

    End Sub

    Sub Load_ICTPLINX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        dst.Tables("ICTPLINX").Rows.Clear()

        Dim t As String = ""

        ASCMAIN1.Progress("Now Building List of Styles", "")
        'Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
        'Dim SALES_DIVISION_CODE As String = Absx1.txtFor("SALES_DIVISION_CODE").Text

        Dim sqlw As String = ""
        If chkSNLast.Checked And chkSNNext.Checked And chkSNPrev.Checked And chkSNThis.Checked Then
        Else
            If chkSNNext.Checked Then sqlw &= " or SEASON_CODE = '" & chkSNNext.Text & "'"
            If chkSNThis.Checked Then sqlw &= " or SEASON_CODE = '" & chkSNThis.Text & "'"
            If chkSNLast.Checked Then sqlw &= " or SEASON_CODE = '" & chkSNLast.Text & "'"
            If chkSNPrev.Checked Then
                sqlw &= " or SEASON_CODE = '" & Mid(chkSNPrev.Text, 1, 5) & "'"
                If Mid(chkSNPrev.Text, 1, 4) = "F" Then
                    sqlw &= " or SUBSTR(SEASON_CODE,1,4) <= '" & Mid(chkSNPrev.Text, 1, 4) & "'"
                Else
                    sqlw &= " or SUBSTR(SEASON_CODE,1,4) < '" & Mid(chkSNPrev.Text, 1, 4) & "'"
                End If
            End If

            If sqlw = "" Then
                sqlw = " and 1<>1"
            Else
                sqlw = " and (" & Mid(sqlw, 5) & ")"
            End If
        End If

        'If SEASON_CODE <> "" Then
        '    sqlw &= " and SEASON_CODE = '" & SEASON_CODE & "'"
        '    t &= ", Season " & SEASON_CODE
        'End If
        'If SALES_DIVISION_CODE <> "" Then
        '    sqlw &= " and SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
        '    t &= ", Division " & SALES_DIVISION_CODE
        'End If

        ASCMAIN1.sql = "Select * from ICTPLIN2 " & ASCMAIN1.SQL_Add_WHERE(sqlw)

        Fill_Records("ICTPLINX", "", , ASCMAIN1.sql)
        Sort_grdColumns(grdICTPLINX, "STYLE_CODE_PLM")
        grdICTPLINX.Text = "All Styles in " & Mid(t, 3)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdICTPLINX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPLINX.DoubleClickRow
        Absx1.txtFor("STYLE_CODE_PLM").Text = e.Row.Cells("STYLE_CODE_PLM").Value & ""
        Click_Command("View")
    End Sub

#Region "grdICTPLIN3"

    Private Sub grdICTPLIN3_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPLIN3.AfterCellUpdate
        With grdICTPLIN3.ActiveRow
            Select Case e.Cell.Column.Key
                Case "DUTY_RATE_CODE"
                    Dim DUTY_RATE_CODE As String = e.Cell.Value & ""
                    If DUTY_RATE_CODE <> "" Then
                        Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", DUTY_RATE_CODE)
                        If rowICTDUTY1 IsNot Nothing Then
                            .Cells("DUTY_RATE").Value = rowICTDUTY1.Item("DUTY_RATE")
                        End If
                    End If
                    '     .Cells("DUTY_COST").Value = .Cells("DUTY_COST_CALC").Value

                Case "LANDED_COST_CALC"
                    Dim ROYALTY_PCT As Decimal = Val(Absx1.numFor("ROYALTY_PCT").Value & "")
                    .Cells("STYLE_PRICE_CALC").Value = TAC.ICCMAIN1.Calculate_Suggested_SP(Val(.Cells("STYLE_PRICE_CALC").Value & ""), ROYALTY_PCT)
            End Select
        End With
    End Sub

    Private Sub grdICTPLIN3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTPLIN3.AfterRowUpdate
        Dim ROYALTY_PCT As Decimal = Val(Absx1.numFor("ROYALTY_PCT").Value & "")
        Dim STYLE_PRICE_CALC As Decimal = TAC.ICCMAIN1.Calculate_Suggested_SP(Val(e.Row.Cells("LANDED_COST_CALC").Value & ""), ROYALTY_PCT)
        If Val(e.Row.Cells("STYLE_PRICE_CALC").Value & "") <> STYLE_PRICE_CALC Then
            e.Row.Cells("STYLE_PRICE_CALC").Value = STYLE_PRICE_CALC
            e.Row.Update()
        End If
    End Sub


    Private Sub grdICTPLIN3_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTPLIN3.BeforeRowUpdate

        Dim DUTY_CATGY_CODE As String = e.Row.Cells("DUTY_CATGY_CODE").Value & ""
        If DUTY_CATGY_CODE = "" Then
            e.Cancel = True
        Else
            If LookUp("ICTDUTY2", DUTY_CATGY_CODE) Is Nothing Then
                e.Cancel = True
            End If
        End If

        If e.Cancel Then MsgBox("Invalid Duty Category Code")

        ' e.Row.Cells("STYLE_PRICE_CALC_BEF_ROY").Value = TAC.ICCMAIN1.Calculate_Suggested_SP(Val(e.Row.Cells("LANDED_COST_CALC").Value & ""))
        ' LANDED_COST_CALC DOES NOT HAVE THE CORRECT FINAL VALUE IN IT
        If e.Row.IsAddRow Then
            e.Row.Cells("STYLE_TYPE_SEQ").Value = Val(dst.Tables("ICTPLIN3").Compute("MAX(STYLE_TYPE_SEQ)", "") & "") + 1
            e.Row.Cells("STYLE_CODE_PLM").Value = STYLE_CODE_PLM
            e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
        Else
            e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        End If
    End Sub

    Private Sub grdICTPLIN3_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPLIN3.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

                Case "DUTY_CATGY_CODE"
                    If grdICTPLIN3.ActiveRow IsNot Nothing Then
                        'If grdICTPLIN3.ActiveRow IsNot Nothing AndAlso grdICTPLIN3.ActiveRow.IsAddRow Then
                        Dim sql_where As String = ""
                        grdClickCellButton(grdICTPLIN3, sql_where)
                    End If

                Case "DUTY_RATE_CODE"
                    If grdICTPLIN3.ActiveRow IsNot Nothing Then
                        Dim DUTY_CATGY_CODE As String = grdICTPLIN3.ActiveRow.Cells("DUTY_CATGY_CODE").Value & ""
                        Dim sql_where As String = ""
                        If DUTY_CATGY_CODE <> "" Then sql_where = "DUTY_CATGY_CODE = '" & DUTY_CATGY_CODE & "'"
                        grdClickCellButton(grdICTPLIN3, sql_where)
                    End If
            End Select
        End With

    End Sub

#End Region

#Region "grdICTPLIN4"

    Private Sub grdICTPLIN4_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTPLIN4.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("STYLE_DTL_SEQ").Value = Val(dst.Tables("ICTPLIN4").Compute("MAX(STYLE_DTL_SEQ)", "") & "") + 1
            e.Row.Cells("STYLE_CODE_PLM").Value = STYLE_CODE_PLM
            'e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            'e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
        Else
            'e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            'e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        End If
    End Sub
#End Region

    Private Sub chkSNNext_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSNNext.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ICTPLINX()
    End Sub

    Private Sub chkSNThis_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSNThis.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ICTPLINX()
    End Sub

    Private Sub chkSNLast_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSNLast.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ICTPLINX()
    End Sub

    Private Sub chkSNPrev_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSNPrev.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ICTPLINX()
    End Sub

    Private Sub cmdNewStyle_Click(sender As System.Object, e As System.EventArgs) Handles cmdNewStyle.Click

        Dim STYLE_CLASS_CODE As String = txtSTYLE_CLASS_CODE.Text
        Dim SEASON_CODE As String = txtSEASON_CODE.Text
        Dim SALES_DIVISION_CODE As String = txtSALES_DIVISION_CODE.Text
        Dim ROYALTY_CODE As String = txtROYALTY_CODE.Text

        If STYLE_CLASS_CODE = "" _
        Or SEASON_CODE = "" _
        Or SALES_DIVISION_CODE = "" Then
            MsgBox("You must first specify a Class, Season and Division Code before generating a New Style", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim rowICTCLAS1 As DataRow = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
        If rowICTCLAS1 Is Nothing Then
            MsgBox("Invalid Value Specifed for Style Class Code", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim rowICTSEAS1 As DataRow = LookUp("ICTSEAS1", SEASON_CODE)
        If rowICTSEAS1 Is Nothing Then
            MsgBox("Invalid Value Specifed for Season Code", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim rowSOTSDIV1 As DataRow = LookUp("SOTSDIV1", SALES_DIVISION_CODE)
        If rowSOTSDIV1 Is Nothing Then
            MsgBox("Invalid Value Specifed for Sales Division Code", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim STYLE_PREFIX As String = ""
        If ROYALTY_CODE <> "" Then
            Dim rowICTROYL1 As DataRow = LookUp("ICTROYL1", ROYALTY_CODE)
            If rowICTROYL1 Is Nothing Then
                MsgBox("Invalid Value Specifed for Royalty Code", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Exit Sub
            Else
                STYLE_PREFIX = rowICTROYL1.Item("STYLE_PREFIX") & ""
            End If
        End If

        If Not ASCMAIN1.Logical_Lock("ICTCLAS1", STYLE_CLASS_CODE, False, True, True, 2) Then Exit Sub

        Dim EMsg As String = Generate_Next_Style(STYLE_CLASS_CODE, STYLE_PREFIX)


        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Auto-Generate Style")
        Else
            If ScreenMode Then
                Absx1.txtFor("STYLE_CLASS_CODE").Text = STYLE_CLASS_CODE
                Absx1.txtFor("SEASON_CODE").Text = SEASON_CODE
                Absx1.txtFor("SALES_DIVISION_CODE").Text = SALES_DIVISION_CODE
                Absx1.txtFor("ROYALTY_CODE").Text = ROYALTY_CODE
            End If
        End If

        'ASCMAIN1.MultiTask_Release(, , 2)

    End Sub

    Function Generate_Next_Style(STYLE_CLASS_CODE As String, STYLE_PREFIX As String) As String

        Dim EMsg As String = ""

        Dim rowICTCLAS1 As DataRow = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
        If rowICTCLAS1 Is Nothing Then
            EMsg = "Could not find Class Record for " & STYLE_CLASS_CODE
        Else
            Dim STYLE_CLASS_STYLE_MASK As String = rowICTCLAS1.Item("STYLE_CLASS_STYLE_MASK") & ""
            If STYLE_CLASS_STYLE_MASK = "" Then
                STYLE_CLASS_STYLE_MASK = STYLE_CLASS_CODE & "-" & "#"
                'STYLE_CLASS_STYLE_MASK = "##########"
            ElseIf Not STYLE_CLASS_STYLE_MASK.Contains("#") Then
                STYLE_CLASS_STYLE_MASK &= "##########"
            End If
            If STYLE_CLASS_STYLE_MASK.Length > Absx1.txtFor("STYLE_CODE_PLM").MaxLength Then
                EMsg = "Total Mask Size Exceeds Maximum Length of Style Code"
            Else
                Dim STYLE_SEQ_START As Integer = InStr(STYLE_CLASS_STYLE_MASK, "#")
                Dim S As String = STYLE_CLASS_STYLE_MASK.Substring(STYLE_SEQ_START - 1)
                Dim STYLE_SEQ_LENGTH As Integer = 1
                Do While S.Length > STYLE_SEQ_LENGTH And Mid(S, STYLE_SEQ_LENGTH, 1) = "#"
                    STYLE_SEQ_LENGTH += 1
                Loop
                If STYLE_CLASS_STYLE_MASK <> rowICTCLAS1.Item("STYLE_CLASS_STYLE_MASK") & "" Then
                    ASCDATA1.ExecuteSQL("Update ICTCLAS1 Set STYLE_CLASS_STYLE_MASK = '" & STYLE_CLASS_STYLE_MASK & "' where STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'")
                End If
                ASCDATA1.ExecuteSQL("Update ICTCLAS1 Set STYLE_CLASS_STYLE_SEQ = NVL(STYLE_CLASS_STYLE_SEQ,0) + 1 where STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'")
                Dim STYLE_CLASS_STYLE_SEQ As Int64 = Val(rowICTCLAS1.Item("STYLE_CLASS_STYLE_SEQ") & "")
                rowICTCLAS1 = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
                If STYLE_CLASS_STYLE_SEQ + 1 <> Val(rowICTCLAS1.Item("STYLE_CLASS_STYLE_SEQ") & "") Then
                    EMsg = "Problem with Sequence Control"
                Else
                    STYLE_CLASS_STYLE_SEQ = Val(rowICTCLAS1.Item("STYLE_CLASS_STYLE_SEQ") & "")
                    Dim SF As String = "".PadLeft(STYLE_SEQ_LENGTH, "#")
                    Dim SS As String = Format(STYLE_CLASS_STYLE_SEQ, Replace(SF, "#", "0"))
                    Dim STYLE_CODE_PLM As String = STYLE_PREFIX & Replace(STYLE_CLASS_STYLE_MASK, SF, SS, , 1)
                    Dim row As DataRow = LookUp("ICTSTYL1", STYLE_CODE_PLM)

                    If row Is Nothing Then
                        ASCMAIN1.sql = "Select STYLE_CODE from ICTSTYL1 where STYLE_CODE like :PARM1 || '%'"
                        row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", STYLE_CODE_PLM)
                    End If

                    If row IsNot Nothing Then
                        EMsg = "Auto-Generated Next Style Code (" & STYLE_CODE_PLM & ") already exists in ERP Style Table." _
                            & vbCrLf & "See ERP Style Code " & row.Item(0) & "." _
                            & vbCrLf & "(You may want to try again to generate the next Style Number Suffix)"

                    Else

                        row = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                        If row IsNot Nothing Then
                            EMsg = "Auto-Generated Next Style Code (" & STYLE_CODE_PLM & ") already exists in PLM Style Table" _
                               & vbCrLf & "(You may want to try again to generate the next Style Number Suffix)"
                        Else
                            If Not ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE_PLM, False, True, True, 2) Then Return ""
                            If Not ASCMAIN1.Logical_Lock("ICTPLIN2", STYLE_CODE_PLM, False, True, True, 2) Then Return ""
                            Absx1.txtFor("STYLE_CODE_PLM").Text = STYLE_CODE_PLM
                            auto_generated_style = True
                            Click_Command("New")
                            auto_generated_style = ScreenMode
                        End If
                    End If
                End If
            End If
        End If

        ASCMAIN1.MultiTask_Release(, , 2)

        Return EMsg

    End Function

    Sub Print_Record()
        Synch_TABLE_NAME("ICTPLIN2")

        ' these lines are repeated in Update_Record
        For Each rowICTPLIN3 As DataRow In dst.Tables("ICTPLIN3").Select("")
            rowICTPLIN3.Item("LANDED_COST") = rowICTPLIN3.Item("LANDED_COST_CALC")
            rowICTPLIN3.Item("DUTY_COST") = rowICTPLIN3.Item("DUTY_COST_CALC")
            rowICTPLIN3.Item("STYLE_PRICE") = rowICTPLIN3.Item("STYLE_PRICE_CALC")
        Next

        rowICTPLIN2.Item("IMAGE") = imgba

        Print_Report_Begin()
        Generate_Report("ICRPLIN2", Me.Text, "Style " & STYLE_CODE_PLM & " Information Sheet")
        Print_Report_End()
    End Sub

    Private Sub btnFirst_Click(sender As System.Object, e As System.EventArgs) Handles btnFirst.Click
        Style_Navigate("<<")
    End Sub

    Private Sub btnPrev_Click(sender As System.Object, e As System.EventArgs) Handles btnPrev.Click
        Style_Navigate("<")
    End Sub

    Private Sub btnNext_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click
        Style_Navigate(">")
    End Sub

    Private Sub btnLast_Click(sender As System.Object, e As System.EventArgs) Handles btnLast.Click
        Style_Navigate(">>")
       
    End Sub

    Sub Style_Navigate(action As String)

        Dim STYLE_CODE_PLM As String = ""

        If Not Absx1.txtFor("STYLE_CODE_PLM").Text.Contains("-") Then
            Dim rows() As DataRow = dst.Tables("ICTPLINN").Select("", "STYLE_CODE_PLM")

            Select Case action
                Case "<<"
                    STYLE_CODE_PLM = rows(0).Item("STYLE_CODE_PLM")

                Case "<"
                    rows = dst.Tables("ICTPLINN").Select("STYLE_CODE_PLM < '" & Absx1.txtFor("STYLE_CODE_PLM").Text & "'", "STYLE_CODE_PLM DESC")
                    If rows.Length <> 0 Then
                        STYLE_CODE_PLM = rows(0).Item("STYLE_CODE_PLM")
                    End If

                Case ">"
                    rows = dst.Tables("ICTPLINN").Select("STYLE_CODE_PLM > '" & Absx1.txtFor("STYLE_CODE_PLM").Text & "'", "STYLE_CODE_PLM")
                    If rows.Length <> 0 Then
                        STYLE_CODE_PLM = rows(0).Item("STYLE_CODE_PLM")
                    End If

                Case ">>"
                    STYLE_CODE_PLM = rows(rows.Length - 1).Item("STYLE_CODE_PLM")
            End Select
        Else
            Dim S1 As String = Split(Absx1.txtFor("STYLE_CODE_PLM").Text, "-")(0)
            Dim S2 As String = Split(Absx1.txtFor("STYLE_CODE_PLM").Text, "-")(1)

            Dim SMAX As Integer = 2000 ' SHOULD CALCULATE THIS AS SMAX WHEN WE LOAD ICTPLINN
            Dim S As String = ""

            Select Case action
                Case "<<"
                    For I As Integer = 0 To SMAX
                        Dim SFX As String = CStr(I)
                        S = S1 & "-" & SFX
                        If dst.Tables("ICTPLINN").Rows.Find(S) IsNot Nothing Then
                            STYLE_CODE_PLM = S
                            Exit For
                        End If

                    Next

                Case "<"
                    For I As Integer = Val(S2) - 1 To 0 Step -1
                        Dim SFX As String = CStr(I)
                        S = S1 & "-" & SFX
                        If dst.Tables("ICTPLINN").Rows.Find(S) IsNot Nothing Then
                            STYLE_CODE_PLM = S
                            Exit For
                        End If
                    Next
                Case ">"
                    For I As Integer = Val(S2) + 1 To SMAX
                        Dim SFX As String = CStr(I)
                        S = S1 & "-" & SFX
                        If dst.Tables("ICTPLINN").Rows.Find(S) IsNot Nothing Then
                            STYLE_CODE_PLM = S
                            Exit For
                        End If
                    Next
                Case ">>"
                    For I As Integer = SMAX To 0 Step -1
                        Dim SFX As String = CStr(I)
                        S = S1 & "-" & SFX
                        If dst.Tables("ICTPLINN").Rows.Find(S) IsNot Nothing Then
                            STYLE_CODE_PLM = S
                            Exit For
                        End If

                    Next
            End Select

        End If

        If STYLE_CODE_PLM <> "" Then
            Click_Done_and_View_Style(STYLE_CODE_PLM)
        End If
    End Sub

    Private Sub txtSTYLE_CODE_NAV_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtSTYLE_CODE_NAV.ValueChanged
       
    End Sub

    Sub Click_Done_and_View_Style(STYLE_CODE_PLM As String)

        UltraExplorerBar1.Groups("Navigator").Visible = False
        Click_Command("Done")
        Absx1.txtFor("STYLE_CODE_PLM").Text = STYLE_CODE_PLM
        Click_Command("View")
    End Sub

    Private Sub grdICTPLINN_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPLINN.DoubleClickRow
        Click_Done_and_View_Style(e.Row.Cells("STYLE_CODE_PLM").Value)
    End Sub

    Private Sub UltraGrid1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPLINN.InitializeLayout

    End Sub

    Private Sub cmdMultiStyle_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub chkNextStyleNo_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkNextStyleNo.CheckedChanged
        If chkNextStyleNo.Checked Then
            txtStyleCode_CopyTo.Text = ""
            txtStyleCode_CopyTo.Enabled = False
        Else
            txtStyleCode_CopyTo.Enabled = True
        End If
    End Sub

    Private Sub cmdCopyStyle_Click(sender As System.Object, e As System.EventArgs) Handles cmdCopyStyle.Click

        Dim EMsg As String = ""

        Dim STYLE_CODE_PLM_copy_from As String = STYLE_CODE_PLM

        Dim STYLE_CLASS_CODE = Absx1.txtFor("STYLE_CLASS_CODE").Text
        Dim rowICTCLAS1 As DataRow = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
        If rowICTCLAS1 Is Nothing Then
            EMsg &= vbCr & "Invalid Style Class Code"
        End If

        Dim SALES_DIVISION_CODE = Absx1.txtFor("SALES_DIVISION_CODE").Text
        Dim rowSOTSDIV1 As DataRow = LookUp("SOTSDIV1", SALES_DIVISION_CODE)
        If rowSOTSDIV1 Is Nothing Then
            EMsg &= vbCr & "Invalid Sales Division Code"
        End If

        Dim SEASON_CODE = Absx1.txtFor("SEASON_CODE").Text
        Dim rowICTSEAS1 As DataRow = LookUp("ICTSEAS1", SEASON_CODE)
        If rowICTSEAS1 Is Nothing Then
            EMsg &= vbCr & "Invalid Season Code"
        End If

        Dim STYLE_PREFIX As String = ""
        Dim ROYALTY_CODE = Absx1.txtFor("ROYALTY_CODE").Text
        Dim rowICTROYL1 As DataRow = LookUp("ICTROYL1", ROYALTY_CODE)
        If rowICTROYL1 Is Nothing Then
            '   EMsg &= vbCr & "Invalid Royalty Code"
        Else
            STYLE_PREFIX = rowICTROYL1.Item("STYLE_PREFIX") & ""
        End If

        Dim row_original As DataRow = dst.Tables("ICTPLIN2").NewRow
        row_original.ItemArray = rowICTPLIN2.ItemArray

        If EMsg = "" Then

            If chkNextStyleNo.Checked Then
                Generate_Next_Style(STYLE_CLASS_CODE, STYLE_PREFIX)
            Else
                Dim STYLE_CODE_PLM As String = txtStyleCode_CopyTo.Text
                If STYLE_CODE_PLM = "" Then
                    EMsg &= vbCr & "Invalid Style Code to Copy To - No Style Code Specified"
                Else
                    Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
                    If rowICTPLIN2 IsNot Nothing Then
                        EMsg &= vbCr & "Invalid Style Code to Copy To - Style " & STYLE_CODE_PLM & " exists"
                    Else
                        Click_Command("Done")
                        Absx1.txtFor("STYLE_CODE_PLM").Text = STYLE_CODE_PLM
                        Click_Command("New")
                    End If
                End If
            End If
        End If
 
        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Copy Style")
        Else
            If EntryMode = "N" Then

                EnforceConstraints(False)

                For Each dcol As DataColumn In dst.Tables("ICTPLIN2").Columns
                    If dcol.ColumnName <> "STYLE_CODE_PLM" Then
                        rowICTPLIN2.Item(dcol.ColumnName) = row_original.Item(dcol.ColumnName)
                    End If
                Next

                For Each T As String In New String() {"ICTPLIN3", "ICTPLIN4"}
                    Fill_Records(T, STYLE_CODE_PLM_copy_from)
                    For Each row As DataRow In dst.Tables(T).Select()
                        row.Item("STYLE_CODE_PLM") = Absx1.txtFor("STYLE_CODE_PLM").Text
                        row.AcceptChanges() '
                        row.SetAdded()
                    Next
                Next

                EnforceConstraints(True)

                'Sort_grdColumns(grdICTPLIN3, "STYLE_TYPE_SEQ")
                'Sort_grdColumns(grdICTPLIN4, "STYLE_DTL_SEQ")
            End If
        End If
    End Sub

    Sub Apply_Royalty()
        Dim ROYALTY_PCT As Decimal = Val(Absx1.numFor("ROYALTY_PCT").Value & "")
        For Each rowICTPLIN3 As DataRow In dst.Tables("ICTPLIN3").Select("")
            rowICTPLIN3.Item("STYLE_PRICE_CALC") = TAC.ICCMAIN1.Calculate_Suggested_SP(Val(rowICTPLIN3.Item("LANDED_COST") & ""), ROYALTY_PCT)
        Next
    End Sub

    Private Sub grdICTPLIN3_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPLIN3.InitializeLayout

    End Sub

    Private Sub grdICTPLIN3_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTPLIN3.InitializeRow
        Dim STYLE_PRICE As Decimal = Val(e.Row.Cells("STYLE_PRICE").Value & "")
        Dim STYLE_PRICE_OVERRIDE As Decimal = Val(e.Row.Cells("STYLE_PRICE_OVERRIDE").Value & "")
        If STYLE_PRICE_OVERRIDE <> 0 And STYLE_PRICE_OVERRIDE < STYLE_PRICE Then
            e.Row.Cells("STYLE_PRICE_OVERRIDE").Appearance.ForeColor = Color.Red
        Else
            e.Row.Cells("STYLE_PRICE_OVERRIDE").Appearance.ForeColor = Color.Empty
        End If
    End Sub
End Class