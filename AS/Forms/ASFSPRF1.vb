Public Class ASFSPRF1
    Dim tblASTSPRF1 As New DataTable
    Dim REPORT_NO As String
    Dim MENU_ITEM_OBJECTs As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Create_Lookup("ASTUSER1")
        dst.Tables.Add(tblASTSPRF1)

        For Each MENU_ITEM_OBJECT As String In ASCMAIN1.MENU_ITEM_OBJECTs
            MENU_ITEM_OBJECTs &= ",'" & MENU_ITEM_OBJECT & "'"
        Next
        Show_Filter(grdASTSPRF1, True)
    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Reports"
                If Absx1.txtFor("USER_ID").Text <> "" Then
                    Validate_Code("USER_ID")
                End If

                If dteREPORT_DATE.Value Is Nothing Then
                    EMsg &= vbCr & "No Date Specified"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Reports"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                'If ASCMAIN1.Running_in_VS Then Move_Reports() : Exit Sub
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load Reports").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpLoadOptions, ScreenMode)

        grdASTSPRF1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Report from Archive")
        Me.Cursor = Cursors.WaitCursor

        Dim sql As String = ""
        If Absx1.txtFor("USER_ID").Text <> "" Then
            If chkIncludeOtherUsers.Checked Then
                'grdASTSPRF1.DisplayLayout.Bands(0).Columns("USER_ID").Hidden = False
                sql = " and MENU_ITEM_OBJECT in (" & Mid(MENU_ITEM_OBJECTs, 2) & ")"
            Else
                'grdASTSPRF1.DisplayLayout.Bands(0).Columns("USER_ID").Hidden = True
                sql = " and USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
            End If
        End If
        If Absx1.txtFor("FORM_NAME").Text <> "" Then
            sql = " and FORM_NAME = '" & Absx1.txtFor("FORM_NAME").Text & "'"
        End If
        If optLoadOption.Value = "0" Then
            sql = sql & " and REPORT_DATE >= '" & Format(dteREPORT_DATE.Value, "dd-MMM-yyyy") & "'"
            sql = sql & " and REPORT_DATE < '" & Format(DateValue(dteREPORT_DATE.Value.ToString).AddDays(1), "dd-MMM-yyyy") & "'"
        Else
            sql = sql & " and REPORT_DATE >= '" & Format(DateValue(dteREPORT_DATE.Value.ToString).AddDays(-100), "dd-MMM-yyyy") & "'"
            sql = sql & " and REPORT_DATE < '" & Format(DateValue(dteREPORT_DATE.Value.ToString).AddDays(1), "dd-MMM-yyyy") & "'"
        End If
        ASCMAIN1.sql = "Select * from ASTSPRF1 " & ASCMAIN1.SQL_Add_WHERE(sql)

        If optLoadOption.Value = "1" Then
            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Select * from", "Select TOP 100 * from")
            Else
                ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & " order by REPORT_DATE DESC) where ROWNUM < 101"
            End If
        End If

        tblASTSPRF1 = ASCDATA1.GetDataTable
        grdASTSPRF1.DataSource = tblASTSPRF1
        Sort_grdColumns(grdASTSPRF1, "REPORT_DATE")

        If Absx1.txtFor("USER_ID").Text <> "" Then
            If chkIncludeOtherUsers.Checked Then
                grdASTSPRF1.DisplayLayout.Bands(0).Columns("USER_ID").Hidden = False
            Else
                grdASTSPRF1.DisplayLayout.Bands(0).Columns("USER_ID").Hidden = True
            End If
        Else
            grdASTSPRF1.DisplayLayout.Bands(0).Columns("USER_ID").Hidden = False
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()
        If Absx1.txtFor("USER_ID").Text = "" Then
            Absx1.txtFor("USER_ID").Text = ASCMAIN1.USER_ID
        End If

        If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("SY") Then
            Set_Read_Only(Absx1.txtFor("USER_ID"), True)
        End If
        optLoadOption.CheckedIndex = 0

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdASTSPRF1, "SSS", "Show Filter", "Show GroupBox", "Select Entire Report Group")
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        'If tlb_pop.Tools.Exists("Select Entire Report Group") Then
        '    tlb_sbt = DirectCast(tlb_pop.Tools("Select Entire Report Group"), UltraWinToolbars.StateButtonTool)
        '    tlb_sbt.Checked = grd.DisplayLayout.Bands(0).CardView
        'End If

        If grd.Name = "grdASTSPRF1" _
        Or grd.Name = "grdASTSPRF1" Then
            Exit Sub
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""
                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select Entire Report Group"
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        If COLUMN_NAME = "USER_ID" Then
            If ctl.Text <> "" Then
                'Call Click_Command("Load Reports")
            End If
        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "USER_ID" Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                Call Click_Command("Load Reports", e)
            End If
        End If
    End Sub
#End Region

#Region "grdASTSPRF1"

    Private Sub grdASTSPRF1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdASTSPRF1.DoubleClickRow
        If grdASTSPRF1.ActiveRow Is Nothing OrElse Not grdASTSPRF1.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Retrieving Report")

        REPORT_NO = grdASTSPRF1.ActiveRow.Cells("REPORT_NO").Text

        Dim FORM_NAME As String = grdASTSPRF1.ActiveRow.Cells("FORM_NAME").Text
        Dim XNO As String = grdASTSPRF1.ActiveRow.Cells("XNO").Text

        Dim dvw As New DataView(tblASTSPRF1)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Select Entire Report Group"), UltraWinToolbars.StateButtonTool)
        If tlb_sbt.Checked Then
            dvw.RowFilter = "FORM_NAME = '" & FORM_NAME & "' and XNO = '" & XNO & "'"
        Else
            dvw.RowFilter = "REPORT_NO = '" & REPORT_NO & "'"
        End If
        Dim f As New ASFSRPTV
        f.Set_Table(dvw.ToTable)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
#End Region


    Sub Move_Reports()
        'see Proceed Done

        If MsgBox("Move Reports into Subfolders?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Moving Files")

        Dim i As Int64 = 0
        For Each filename As String In My.Computer.FileSystem.GetFiles(ASCMAIN1.Folders("Archive") & "Reports\")
            'Stop
            Dim FILENAME_FIXED As String = Replace(filename, " ", "")
            If FILENAME_FIXED.Length <> 39 Then Stop
            'Stop
            Dim FN() As String = Split(FILENAME_FIXED, "\")
            If FN.Length <> 5 Then Stop
            'Stop
            Dim F As String = ASCMAIN1.Folders("Archive") & "Reports\" & FN(1) & "\" & Mid(FN(4), 5, 5) & "\"
            My.Computer.FileSystem.CreateDirectory(F)
            My.Computer.FileSystem.MoveFile(filename, F & FN(4))
            i += 1
            If i Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(i) & ":" & FN(4))
                Application.DoEvents()
            End If
            'Stop
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        MsgBox("Move Complete")
    End Sub
End Class