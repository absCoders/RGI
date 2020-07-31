Imports System.ComponentModel
Imports System.Text

Public Class SOFWBIMG

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Private FULL_FILE_NAME As String = ""
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFWBIMI" Then
            InquiryMode = True
        End If

        Check_Form_Options()

        With dst
            Dim SQLB As New System.Text.StringBuilder
            SQLB.Length = 0
            SQLB.AppendLine("SELECT *")
            SQLB.AppendLine("FROM SOTIMGM1")
            SQLB.AppendLine("WHERE IMG_NO = :PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTIMGM1", "**", 0, True, "V")

            SQLB.Length = 0
            SQLB.Length = 0
            SQLB.AppendLine("SELECT *")
            SQLB.AppendLine("FROM SOTIMGM2")
            SQLB.AppendLine("WHERE IMG_NO = :PARM1")
            ASCMAIN1.sql = SQLB.ToString
            Create_TDA(.Tables.Add, "SOTIMGM2", "**", 0, True, "V")

        End With

        grdSOTIMGM2.DataSource = dst.Tables("SOTIMGM2")

        'ASCMAIN1.Add_Value_List(grdSOFCSTMX, "REPORT_TYPE", , New String() {":", "I:Initial", "A:Amended", "S:Subsequent", "R:Revised"})

        'Create_Summary(grdSOFPRIC1, "ORDRED_TY", "Sum")
        'Sort_grdColumns(grdSOFPRIC1, "STYLE_CODE", True)

        'With grdSOFPRIC1.DisplayLayout.Bands(0)
        '    For Each COL_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC"}
        '        .Columns(COL_NAME).Header.Fixed = True
        '    Next
        'End With

        TABLE_NAME = "SOTIMGM1"

        EntryMode = "E"
        Call Load_Record()
        Call Mode_Settings(True)
    End Sub

    Sub Check_Inquiry_Mode()
        If InquiryMode Then
        Else
        End If
    End Sub

    Sub Check_Form_Options()
        'With UltraExplorerBar1.Groups("Screen Control")
        '    .Items("New").Visible = (Me.Name = "PMFVIST1")
        'End With
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

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                Dim OD As New OpenFileDialog
                OD.InitialDirectory = "C:\"
                OD.Filter = "jpg files (*.jpg)|*.jpg|All files (*.*)|*.*"
                OD.Title = "Please Select An Image To Begin"
                If OD.ShowDialog() = DialogResult.OK Then
                    Dim iResult As String = InputBox("Please Provide The Name Of This Image", "Image Name")
                    If iResult <> "" Then
                        Absx1.txtFor("IMG_NAME").Text = iResult
                        FULL_FILE_NAME = OD.FileName
                        PictureBox1.ImageLocation = FULL_FILE_NAME
                        Absx1.txtFor("FILE_NAME").Text = OD.SafeFileName
                        Absx1.txtFor("IMG_NO").Text = ASCMAIN1.Next_Control_No("SOTIMGM1.IMG_NO")
                        Call Mode_Settings(True)
                    Else
                        MsgBox("You Must Provide An Image Name", vbOKOnly, "Please Insert Another Quarter And Try Again.")
                        Call Mode_Settings(False)
                    End If
                End If
            Case "Load"
                Dim S As New Text.StringBuilder With {.Length = 0}
                S.AppendLine("SELECT IMG_NO, IMG_NAME, FILE_NAME FROM SOTIMGM1")
                With ASCMAIN1.CodeSelector
                    .SQL = S.ToString
                    .MultipleSelections = False
                    .PreviouslySelectedCodes0 = ""
                    .Caption = "Select Image To Load"
                    .TABLE_NAME = ""
                    .VIEW_NAME = ""
                    .VIEW_DESC = ""
                    .COLUMN_NAME = ""
                    .COLUMN_PREKEYs = New Dictionary(Of String, String)
                    .Custom_sql_where = ""
                    .tblASTVIEW1 = New DataTable
                End With
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    ASCMAIN1.Progress("please Wait... Loading Image")
                    Application.DoEvents()
                    Dim IMG_NO As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("IMG_NO") & ""
                    Fill_Records("SOTIMGM1", IMG_NO)
                    Fill_Records("SOTIMGM2", IMG_NO)
                    Dim rowSOTIMGM1 As DataRow = dst.Tables("SOTIMGM1").Rows(0)
                    Absx1.txtFor("IMG_NO").Text = rowSOTIMGM1.Item("IMG_NO").ToString & String.Empty
                    Absx1.txtFor("IMG_NAME").Text = rowSOTIMGM1.Item("IMG_NAME").ToString & String.Empty
                    Absx1.txtFor("FILE_NAME").Text = rowSOTIMGM1.Item("FILE_NAME").ToString & String.Empty
                    FULL_FILE_NAME = rowSOTIMGM1.Item("FULL_FILE_NAME").ToString & String.Empty
                    PictureBox1.ImageLocation = FULL_FILE_NAME
                    Me.Cursor = Cursors.WaitCursor
                    Application.DoEvents()
                    Timer1.Interval = 5000
                    Timer1.Start()
                    Call Mode_Settings(True)
                End If
            Case "Save Script"

            Case "Done"
                Update_Record()
                Call Mode_Settings(False)
                Me.Close()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Save Script").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'With grdSOFPRIC1.DisplayLayout.Bands(0)
        '    For Each thisCOL As String In New String() {"ORDRED_TY", "SHIPPED_TY", "CANCELLED_TY"}
        '        .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Khaki
        '    Next
        '    For Each thisCOL As String In New String() {"ORDRED_LY", "SHIPPED_LY", "CANCELLED_LY"}
        '        .Columns.Item(thisCOL).Header.Appearance.BackColor = Drawing.Color.Bisque
        '    Next
        'End With


        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("SOTIMGM1").Rows.Clear()
        dst.Tables("SOTIMGM2").Rows.Clear()

        Absx1.txtFor("IMG_NO").Text = ""
        Absx1.txtFor("IMG_NAME").Text = ""
        Absx1.txtFor("FILE_NAME").Text = ""
        FULL_FILE_NAME = ""

        'Dim dvw As DataView = DirectCast(grdPMTVIST1.DataSource, DataTable).DefaultView
        'dvw.RowStateFilter = DataViewRowState.CurrentRows

        'Fill_Records("PMTVIST1")
        'Process_SVRs()

        'Sort_grdColumns(grdPMTVIST1, "DATE_VISITED".ToLower)
        'Sort_grdColumns(grdPMTVISTH, "DATE_VISITED".ToLower)
        dst.EnforceConstraints = True
        'Setup_Summary()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        'Setup_Summary()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()
        BeginTrans()
        'INIT_LAST("PMTVIST1", True, "", True)
        Dim rowSOTIMGM1 As DataRow = Nothing
        If dst.Tables("SOTIMGM1").Rows.Count = 1 Then
            rowSOTIMGM1 = dst.Tables("SOTIMGM1").Rows(0)
            rowSOTIMGM1.Item("IMG_NO") = Absx1.txtFor("IMG_NO").Text & String.Empty
            rowSOTIMGM1.Item("IMG_NAME") = Absx1.txtFor("IMG_NAME").Text & String.Empty
            rowSOTIMGM1.Item("FILE_NAME") = Absx1.txtFor("FILE_NAME").Text & String.Empty
            rowSOTIMGM1.Item("FULL_FILE_NAME") = FULL_FILE_NAME
        Else
            rowSOTIMGM1 = dst.Tables("SOTIMGM1").NewRow
            rowSOTIMGM1.Item("IMG_NO") = Absx1.txtFor("IMG_NO").Text & String.Empty
            rowSOTIMGM1.Item("IMG_NAME") = Absx1.txtFor("IMG_NAME").Text & String.Empty
            rowSOTIMGM1.Item("FILE_NAME") = Absx1.txtFor("FILE_NAME").Text & String.Empty
            rowSOTIMGM1.Item("FULL_FILE_NAME") = FULL_FILE_NAME
            dst.Tables("SOTIMGM1").Rows.Add(rowSOTIMGM1)
        End If
        Update_Record_TDA("SOTIMGM1")
        Update_Record_TDA("SOTIMGM2")
        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        'Call Load_Popup_Menu(grdSOFPRIC1, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
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
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'Select Case e.Tool.Key
        '    Case "Project Center"
        '        Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
        '        Context_Launch("Edit", Column_Values("JOB_NO", JOB_NO), e.Tool.Key, "PMFJOBM1")
        '    Case "Show Report"
        '        Dim FILENAME As String = "C:\Documents and Settings\wjz\Desktop\randfromdrc\RandInvoices\310 West 52nd Street - 30760.pdf"
        '        Show_Document(FILENAME)

        'End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        'Case "EMPLOYEE_CODE"
        '    If e.KeyCode = Windows.Forms.Keys.Enter Then
        '        Setup_Summary()
        '    End If
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        'Select Case COLUMN_NAME
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "EMPLOYEE_CODE"
        '        Setup_Summary()
        'End Select
    End Sub

#End Region

    Sub Setup_Summary()
        Dim sqlwhere As String = ""
        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor
        dst.Tables("SOTPRIC1").Rows.Clear()

        Fill_Records("SOTPRIC1")
        'Fill_Pricing()

        dst.EnforceConstraints = False

        ASCMAIN1.Progress("")
        'grdSOFPRIC1.Update()
        'grdSOFPRIC1.Refresh()
        Me.Cursor = Cursors.Default
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs)
        'Setup_Summary()
    End Sub

    'Private Sub btnLoad_Click(sender As Object, e As EventArgs)
    '    Dim OD As New OpenFileDialog
    '    OD.InitialDirectory = "Z:\Wayne On My Mac\Dropbox\Zoho Documents\Regency International\Laptop Image Mapper"
    '    OD.Filter = "png files (*.jpg)|*.jpg|All files (*.*)|*.*"
    '    If OD.ShowDialog() = DialogResult.OK Then
    '        PictureBox1.ImageLocation = OD.FileName
    '    End If

    'End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Dim arg As System.Windows.Forms.MouseEventArgs = e
        Dim frmASFMSGBF As New ASFMSGBF
        'Dim STYLE_CODE As String = frmASFMSGBF.Get_txtblock_from_User("Please Type Or Scan The Style", "Style Code", "", False, 25)
        Dim STYLE_CODE As String = InputBox("Please Type Or Scan The Style", "Add Style")
        If STYLE_CODE <> "" Then
            Dim x_pos As Int64 = arg.X
            Dim y_pos As Int64 = arg.Y
            Dim p As New System.Drawing.Pen(Drawing.Brushes.Orange, 4)
            Dim g As System.Drawing.Graphics
            g = PictureBox1.CreateGraphics
            g.DrawEllipse(p, x_pos, y_pos, 10, 10)
            Dim rowSOTIMGM2 As DataRow = dst.Tables("SOTIMGM2").NewRow
            rowSOTIMGM2.Item("IMG_NO") = Absx1.txtFor("IMG_NO").Text
            rowSOTIMGM2.Item("IMG_LINE") = Val(dst.Tables("SOTIMGM2").Compute("Max(IMG_LINE)", "") & "") + 1
            rowSOTIMGM2.Item("STYLE_CODE") = STYLE_CODE
            rowSOTIMGM2.Item("X_POS") = x_pos
            rowSOTIMGM2.Item("Y_POS") = y_pos
            dst.Tables("SOTIMGM2").Rows.Add(rowSOTIMGM2)
        End If

    End Sub

    Private Sub PictureBox1_MouseHover(sender As Object, e As EventArgs) Handles PictureBox1.MouseMove
        Dim arg As System.Windows.Forms.MouseEventArgs = e
        Dim x_pos As Int64 = arg.X
        Dim y_pos As Int64 = arg.Y
        Dim x_pos_last As Int64
        Dim y_pos_last As Int64
        grdSOTIMGM2.Selected.Rows.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTIMGM2.Rows
            x_pos_last = grow.Cells.Item("X_POS").Text
            y_pos_last = grow.Cells.Item("Y_POS").Text
            If x_pos > x_pos_last And x_pos < x_pos_last + 10 Then
                If y_pos > y_pos_last And y_pos < y_pos_last + 10 Then
                    grow.Selected = True
                Else
                    grow.Selected = False
                End If
            End If
        Next

    End Sub

    Private Sub grdSOTIMGM2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTIMGM2.AfterRowActivate
        'Dim x_pos As Int64 = Val(grdSOTIMGM2.ActiveRow.Cells.Item("X_POS").Text.ToString() & String.Empty)
        'Dim y_pos As Int64 = Val(grdSOTIMGM2.ActiveRow.Cells.Item("Y_POS").Text.ToString() & String.Empty)
        'Dim p As New System.Drawing.Pen(Drawing.Brushes.Orange, 4)
        'Dim g As System.Drawing.Graphics
        'g = PictureBox1.CreateGraphics
        'g.DrawEllipse(p, x_pos, y_pos, 10, 10)
    End Sub

    'Private Sub PictureBox1_LoadCompleted(sender As Object, e As AsyncCompletedEventArgs) Handles PictureBox1.LoadCompleted
    '    For Each rowSOTIMGM2 As DataRow In dst.Tables("SOTIMGM2").Select()
    '        Dim x_pos As Int64 = Val(rowSOTIMGM2.Item("X_POS").ToString() & String.Empty)
    '        Dim y_pos As Int64 = Val(rowSOTIMGM2.Item("Y_POS").ToString() & String.Empty)
    '        Dim p As New System.Drawing.Pen(Drawing.Brushes.Orange, 4)
    '        Dim g As System.Drawing.Graphics
    '        g = PictureBox1.CreateGraphics
    '        g.DrawEllipse(p, x_pos, y_pos, 10, 10)
    '    Next
    '    Me.Cursor = Cursors.Default
    'End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        For Each rowSOTIMGM2 As DataRow In dst.Tables("SOTIMGM2").Select()
            Dim x_pos As Int64 = Val(rowSOTIMGM2.Item("X_POS").ToString() & String.Empty)
            Dim y_pos As Int64 = Val(rowSOTIMGM2.Item("Y_POS").ToString() & String.Empty)
            Dim p As New System.Drawing.Pen(Drawing.Brushes.Orange, 4)
            Dim g As System.Drawing.Graphics
            g = PictureBox1.CreateGraphics
            g.DrawEllipse(p, x_pos, y_pos, 10, 10)
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    'Private Sub btnPHP_Click(sender As Object, e As EventArgs)
    '    Dim imsg As New StringBuilder With {.Length = 0}
    '    imsg.AppendLine("$products = array(")
    '    For Each rowSOTWBIMG As DataRow In dst.Tables("SOTWBIMG").Select()
    '        Dim STYLE_CODE As String = rowSOTWBIMG.Item("STYLE_CODE").ToString()
    '        Dim X_POS As Int64 = Val(rowSOTWBIMG.Item("X_POS").ToString())
    '        Dim Y_POS As Int64 = Val(rowSOTWBIMG.Item("Y_POS").ToString())
    '        imsg.AppendLine(String.Format("{0}{1}.html{0} => array({2}, {3}),", Chr(34), STYLE_CODE, X_POS, Y_POS))
    '    Next
    '    imsg.AppendLine(");")
    '    MsgBox(imsg.ToString(), vbOKOnly, "Paste This To Shopsite")

    'End Sub
End Class