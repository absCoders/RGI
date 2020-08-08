Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Threading

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
        'picImagMapper.Height = picImagMapper.Parent.Height - 10
        'picImagMapper.Width = picImagMapper.Parent.Width - 10

        Dim Scale As New List(Of String)
        Scale.Add("1.0")
        Scale.Add("1.5")
        Scale.Add("2.0")
        Scale.Add("2.5")
        Scale.Add("3.0")
        Scale.Add("3.5")
        Scale.Add("4.0")
        cboScale.DataSource = Scale
        cboScale.SelectedIndex = 0

        Dim imgSize As New List(Of String)
        imgSize.Add("Zoom")
        imgSize.Add("AutoSize")
        imgSize.Add("Normal")
        imgSize.Add("Stretched")
        imgSize.Add("Center")

        cboImgSize.DataSource = imgSize
        SetImgSizeMode(True)

        grdSOTIMGM2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

        TABLE_NAME = "SOTIMGM1"

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
        'Me.Cursor = Cursors.Default
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

    Sub Update_Record(ByVal Optional iMsg As String = "Update Complete")
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
        CommitTrans(iMsg)
    End Sub

    Sub Print_Report()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'CR_params.Add("GROUP_BY", optGroupBy.Value)
        'Generate_Report("PMRVIST1", "Open Site Visit Report")
        'Print_Report_End()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Update"
            Case "Save Script"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Save Script"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Prompt You For A Folder To")
                iMSG.AppendLine("Save Your Image And Shopsite Script.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Are You Finished And Prepared For")
                iMSG.AppendLine("Saving Your Work?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg = "Save Script Aborted"
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
                        Dim BM As Bitmap = Bitmap.FromFile(FULL_FILE_NAME)
                        picImagMapper.Height = BM.Height / Val(cboScale.Text)
                        picImagMapper.Width = BM.Width / Val(cboScale.Text)
                        picImagMapper.Image = BM
                        Absx1.txtFor("FILE_NAME").Text = OD.SafeFileName
                        Absx1.txtFor("IMG_NO").Text = ASCMAIN1.Next_Control_No("SOTIMGM1.IMG_NO")
                        Call Mode_Settings(True)
                    Else
                        MsgBox("You Must Provide An Image Name", vbOKOnly, "Please Insert Another Quarter And Try Again.")
                        Call Mode_Settings(False)
                    End If
                End If
            Case "Load"
                Dim S As New System.Text.StringBuilder With {.Length = 0}
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
                    FULL_FILE_NAME = rowSOTIMGM1.Item("FULL_FILE_NAME").ToString & String.Empty
                    If Not IO.File.Exists(FULL_FILE_NAME) Then
                        MsgBox(FULL_FILE_NAME, vbOKOnly, "File Specified Does Not Exist")
                        Clear_Record()
                        Call Mode_Settings(False)
                    Else
                        Absx1.txtFor("IMG_NO").Text = rowSOTIMGM1.Item("IMG_NO").ToString & String.Empty
                        Absx1.txtFor("IMG_NAME").Text = rowSOTIMGM1.Item("IMG_NAME").ToString & String.Empty
                        Absx1.txtFor("FILE_NAME").Text = rowSOTIMGM1.Item("FILE_NAME").ToString & String.Empty
                        Dim BM As Bitmap = Bitmap.FromFile(FULL_FILE_NAME)
                        'Dim propItems As PropertyItem() = BM.PropertyItems
                        picImagMapper.Image = BM
                        picImagMapper.Height = BM.Height / Val(cboScale.Text)
                        picImagMapper.Width = BM.Width / Val(cboScale.Text)
                        picImagMapper.BorderStyle = BorderStyle.FixedSingle
                        RefreshDots()
                        Call Mode_Settings(True)
                    End If
                End If
            Case "Save Script"
                SaveScript()
            Case "Done"
                Update_Record()
                Call Mode_Settings(False)
                Me.Close()
            Case "Refresh Dots"
                RefreshDots()
        End Select

    End Sub

    Private Sub RefreshDots()
        ASCMAIN1.Progress("please Wait... Refreshing Images")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Timer1.Interval = 3000
        Timer1.Start()
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Refresh Dots").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Save Script").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            EntryMode = "E"
        Else
            Clear_Record()
            EntryMode = ""
        End If

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

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "JOB_NO"
        '        sql_where = "JOB_STATUS = 'O' and SITE_VISITS > 0"
        'End Select
    End Sub
#End Region

#Region "Form Controls"
    Private Sub chkImgPos_CheckedChanged(sender As Object, e As EventArgs) Handles chkImgPos.CheckedChanged
        If chkImgPos.Checked Then
            lblImgPos.Visible = True
        Else
            lblImgPos.Visible = False
        End If
    End Sub

    Private Sub cboImgSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboImgSize.SelectedIndexChanged
        SetImgSizeMode()
    End Sub

    Private Sub grdSOTIMGM2_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSOTIMGM2.AfterRowsDeleted
        RefreshDots()
    End Sub

    Private Sub picImagMapper_Click(sender As Object, e As EventArgs) Handles picImagMapper.Click
        If EntryMode = "E" Then
            Dim arg As System.Windows.Forms.MouseEventArgs = e
            Dim frmASFMSGBF As New ASFMSGBF
            'Dim STYLE_CODE As String = frmASFMSGBF.Get_txtblock_from_User("Please Type Or Scan The Style", "Style Code", "", False, 25)
            Dim STYLE_CODE As String = InputBox("Please Type Or Scan The Style", "Add Style")
            If STYLE_CODE <> "" Then
                STYLE_CODE = STYLE_CODE.Replace(" ", "")
                STYLE_CODE = STYLE_CODE.ToUpper
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If IsNothing(rowICTSTYL1) Then
                    MsgBox(String.Format("{0}: Style Not Found In Masterfile", STYLE_CODE), vbOKOnly, "Bad Style")
                Else
                    Dim x_pos As Int64 = arg.X
                    Dim y_pos As Int64 = arg.Y
                    Dim p As New System.Drawing.Pen(Drawing.Brushes.Red, 4)
                    Dim g As System.Drawing.Graphics
                    g = picImagMapper.CreateGraphics
                    g.DrawEllipse(p, x_pos, y_pos, 10, 10)
                    Dim rowSOTIMGM2 As DataRow = dst.Tables("SOTIMGM2").NewRow
                    rowSOTIMGM2.Item("IMG_NO") = Absx1.txtFor("IMG_NO").Text
                    rowSOTIMGM2.Item("IMG_LINE") = Val(dst.Tables("SOTIMGM2").Compute("Max(IMG_LINE)", "") & "") + 1
                    rowSOTIMGM2.Item("STYLE_CODE") = STYLE_CODE
                    rowSOTIMGM2.Item("X_POS") = x_pos * Val(cboScale.Text)
                    rowSOTIMGM2.Item("Y_POS") = y_pos * Val(cboScale.Text)
                    dst.Tables("SOTIMGM2").Rows.Add(rowSOTIMGM2)
                End If
            End If
        End If

    End Sub

    Private Sub picImagMapper_MouseHover(sender As Object, e As EventArgs) Handles picImagMapper.MouseMove
        Dim arg As System.Windows.Forms.MouseEventArgs = e
        If chkImgPos.Checked = True Then
            lblImgPos.Visible = True
            'Dim realPos As Point = TranslatePoints(arg.Location)
            'lblImgPos.Text = String.Format("X: {0} | Y: {1}", realPos.X, realPos.Y)
            lblImgPos.Text = String.Format("X: {0} | Y: {1}", arg.X * Val(cboScale.Text), arg.Y * Val(cboScale.Text))
        Else
            lblImgPos.Visible = False
        End If

        Dim x_pos As Int64 = arg.X
        Dim y_pos As Int64 = arg.Y
        Dim x_pos_last As Int64
        Dim y_pos_last As Int64
        grdSOTIMGM2.Selected.Rows.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTIMGM2.Rows
            x_pos_last = grow.Cells.Item("X_POS").Text / Val(cboScale.Text)
            y_pos_last = grow.Cells.Item("Y_POS").Text / Val(cboScale.Text)
            If x_pos > x_pos_last And x_pos < x_pos_last + 10 Then
                If y_pos > y_pos_last And y_pos < y_pos_last + 10 Then
                    grow.Selected = True
                Else
                    grow.Selected = False
                End If
            End If
        Next

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        For Each rowSOTIMGM2 As DataRow In dst.Tables("SOTIMGM2").Select()
            Dim x_pos As Int64 = Val(rowSOTIMGM2.Item("X_POS").ToString() & String.Empty) / Val(cboScale.Text)
            Dim y_pos As Int64 = Val(rowSOTIMGM2.Item("Y_POS").ToString() & String.Empty) / Val(cboScale.Text)
            Dim p As New System.Drawing.Pen(Drawing.Brushes.Red, 4)
            Dim g As System.Drawing.Graphics
            g = picImagMapper.CreateGraphics
            g.DrawEllipse(p, x_pos, y_pos, 10, 10)
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        lblRefreshRequired.Visible = False
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub SaveScript()

        Update_Record("")
        Dim fld As New FolderBrowserDialog

        If fld.ShowDialog() = DialogResult.OK Then

            Try
                Dim path As String = fld.SelectedPath
                Dim FN_IMG As String = dst.Tables.Item("SOTIMGM1").Rows(0).Item("FILE_NAME").ToString
                Dim FN_TXT As String = FN_IMG.Replace("jpg", "txt")
                FN_TXT = FN_TXT.Replace("JPG", "txt")
                FN_TXT = FN_TXT.Replace("Jpg", "txt")

                Dim PROD_NAME As String = FN_IMG.Replace(".jpg", "")
                If IO.File.Exists(path & "\" & FN_IMG) Then
                    IO.File.Delete(path & "\" & FN_IMG)
                End If
                If IO.File.Exists(path & "\" & FN_TXT) Then
                    IO.File.Delete(path & "\" & FN_TXT)
                End If
                picImagMapper.Image.Save(path & "\" & FN_IMG)
                Dim imsg As New StringBuilder With {.Length = 0}
                imsg.AppendLine(String.Format("${0} = array(", PROD_NAME))
                Dim STYLE_CODE_LAST As String = ""
                Dim STYLE_NUM_LAST As Int64 = 0
                Dim STYLE_SUFFIX As String = ""
                For Each rowSOTIMGM2 As DataRow In dst.Tables("SOTIMGM2").Select("", "STYLE_CODE, IMG_LINE")
                    Dim STYLE_CODE As String = rowSOTIMGM2.Item("STYLE_CODE").ToString()
                    If STYLE_CODE <> STYLE_CODE_LAST Then
                        STYLE_SUFFIX = ""
                        STYLE_CODE_LAST = STYLE_CODE
                        STYLE_NUM_LAST = 0
                    Else
                        STYLE_NUM_LAST += 1
                        STYLE_SUFFIX = "#" + STYLE_NUM_LAST.ToString
                    End If
                    Dim OrigPoint As New Point With {.X = CDbl(rowSOTIMGM2.Item("X_POS").ToString()), .Y = CDbl(rowSOTIMGM2.Item("Y_POS").ToString())}
                    'Dim newPos As Point = TranslatePoints(OrigPoint)
                    imsg.AppendLine(String.Format("     {0}{1}.html{2}{0} => array({3}, {4}),", Chr(34), STYLE_CODE, STYLE_SUFFIX, OrigPoint.X, OrigPoint.Y))
                Next
                imsg.AppendLine(");")
                IO.File.WriteAllText(path & "\" & FN_TXT, imsg.ToString)
                MsgBox("Your Files Are Saved", vbOKOnly, "Done")
            Catch ex As Exception
                MsgBox(ex.InnerException.ToString, vbOKOnly, "Error Saving Files")
            End Try

        End If
    End Sub

    Private Sub SetImgSizeMode(Optional ByVal setDefault As Boolean = False)
        If setDefault = True Then
            picImagMapper.SizeMode = PictureBoxSizeMode.Zoom
        Else
            Select Case cboImgSize.Text
                Case "Normal"
                    picImagMapper.SizeMode = PictureBoxSizeMode.Normal
                Case "Stretched"
                    picImagMapper.SizeMode = PictureBoxSizeMode.StretchImage
                Case "AutoSize"
                    picImagMapper.SizeMode = PictureBoxSizeMode.AutoSize
                Case "Center"
                    picImagMapper.SizeMode = PictureBoxSizeMode.CenterImage
                Case "Zoom"
                    picImagMapper.SizeMode = PictureBoxSizeMode.Zoom
            End Select
        End If
        RefreshDots()
    End Sub

    Private Function TranslatePoints(ByVal truP As Point) As Point
        'Dim p As Point = picImagMapper.PointToClient(Cursor.Position)
        Dim RetVal As Point = New Point()
        If Not IsNothing(picImagMapper.Image) Then
            Dim w_i As Integer = picImagMapper.Image.Width
            Dim h_i As Integer = picImagMapper.Image.Height
            Dim w_c As Integer = picImagMapper.Width
            Dim h_c As Integer = picImagMapper.Height

            Dim imageRatio As Single = w_i / CSng(h_i)
            Dim containerRatio As Single = w_c / CSng(h_c)

            If imageRatio >= containerRatio Then
                Dim scaleFactor As Single = w_c / CSng(w_i)
                Dim scaledHeight As Single = h_i * scaleFactor
                Dim filler As Single = Math.Abs(h_c - scaledHeight) / 2
                'RetVal.X = CInt((p.X / scaleFactor))
                'RetVal.Y = CInt(((p.Y - filler) / scaleFactor))
                RetVal.X = CInt(((truP.X - filler) / scaleFactor))
                RetVal.Y = CInt((truP.Y / scaleFactor))
            Else
                Dim scaleFactor As Single = h_c / CSng(h_i)
                Dim scaledWidth As Single = w_i * scaleFactor
                Dim filler As Single = Math.Abs(w_c - scaledWidth) / 2
                'RetVal.X = CInt(((p.X - filler) / scaleFactor))
                'RetVal.Y = CInt((p.Y / scaleFactor))
                RetVal.X = CInt(((truP.X - filler) / scaleFactor))
                RetVal.Y = CInt((truP.Y / scaleFactor))
            End If
        End If

        Return RetVal
    End Function

    Private Sub cboScale_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboScale.SelectedIndexChanged
        If Not IsNothing(picImagMapper.Image) Then
            Dim BM As Bitmap = picImagMapper.Image
            picImagMapper.Height = BM.Height / Val(cboScale.Text)
            picImagMapper.Width = BM.Width / Val(cboScale.Text)
            'picImagMapper.Image = BM
            RefreshDots()
        End If
    End Sub

    Private Sub panImage_Paint(sender As Object, e As PaintEventArgs) Handles panImage.Paint
        lblRefreshRequired.Visible = True
    End Sub

    Private Sub panImage_Scroll(sender As Object, e As ScrollEventArgs) Handles panImage.Scroll
        lblRefreshRequired.Visible = True
    End Sub

#End Region
End Class