
Imports System.Text
Imports System.IO

Public Class WBFIMGWB
    Dim S As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
    Dim TTM As New UltraWinToolTip.UltraToolTipManager
    Dim FTPImages As Boolean = False
    Dim ImageListDownload As New List(Of String)
    Dim FileList As New Dictionary(Of String, String)
    Dim ImageGetProgress As Boolean = True
    Dim ImageListFTP As New Dictionary(Of String, Int64)
    Dim ImageListLocal As New Dictionary(Of String, Int64)

    Private Event OnDirList As nsoftware.IPWorks.Ftp.OnDirListHandler
    'Private Event OnDirListS As nsoftware.IPWorksSSH.Sftp.OnDirListHandler
    ' Dim sqlSOTWORK1 As String
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp
    'Dim WithEvents FtpS As New nsoftware.IPWorksSSH.Sftp


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

        SetFileLocations()

        With dst

            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("RSLT.*,")
            S.AppendLine("(STYLE_CODE || '-' || COLOR_CODE || '.JPG') AS SC")
            S.AppendLine("FROM")
            S.AppendLine("(")
            S.AppendLine("SELECT")
            S.AppendLine("S1.STYLE_CODE,")
            S.AppendLine("C1.COLOR_CODE,")
            S.AppendLine("S1.STYLE_STATUS,")
            S.AppendLine("C1.STYLE_COLOR_STATUS,")
            S.AppendLine("S1.STYLE_DESC,")
            S.AppendLine("SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))) AS AVAIL")
            S.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1, ICTSTAT2 S2")
            S.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE")
            S.AppendLine("AND C1.STYLE_CODE = S2.STYLE_CODE (+)")
            S.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
            S.AppendLine("AND S2.WHSE_CODE = 'MS'")
            S.AppendLine("GROUP BY")
            S.AppendLine("S1.STYLE_CODE,")
            S.AppendLine("C1.COLOR_CODE,")
            S.AppendLine("S1.STYLE_STATUS,")
            S.AppendLine("C1.STYLE_COLOR_STATUS,")
            S.AppendLine("S1.STYLE_DESC")
            S.AppendLine(") RSLT")
            S.AppendLine("WHERE (STYLE_COLOR_STATUS = 'A' OR AVAIL > 0)")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "WBTIMGWB", "**", 0, False)
            With .Tables("WBTIMGWB").Columns
                .Add("IS_WEB", GetType(System.String))
                .Add("IS_LOCAL", GetType(System.String))
                .Add("IS_FTP", GetType(System.String))
                .Add("LOCAL_SIZE", GetType(System.Int64))
                .Add("FTP_SIZE", GetType(System.Int64))
                .Add("IS_MATCHED", GetType(System.String), "LOCAL_SIZE = FTP_SIZE")
            End With

            S.Length = 0
            S.AppendLine("SELECT STYLE_CODE, COLOR_CODE")
            S.AppendLine("FROM WBTSTYLD")
            S.AppendLine("WHERE WEB_IND = 'W'")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "WBTSTYLD", "**", 0, False)
            Fill_Records("WBTSTYLD")

        End With

        grdWBTIMGWB.DataSource = dst.Tables("WBTIMGWB")

        Create_Summary(grdWBTIMGWB, "STYLE_CODE", "Count", "", "###,##0")

        ASCMAIN1.Add_Value_List(grdWBTIMGWB, "STYLE_STATUS", , New String() {":", "A:Active", "N:No Re-Order", "D:Discontinued"})
        ASCMAIN1.Add_Value_List(grdWBTIMGWB, "STYLE_COLOR_STATUS", , New String() {":", "A:Active", "N:No Re-Order", "D:Discontinued"})

        Sort_grdColumns(grdWBTIMGWB, "STYLE_CODE, COLOR_CODE", False)

        grdWBTIMGWB.DisplayLayout.Bands(0).Columns("AVAIL").Format = "###,##0"

        With grdWBTIMGWB.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        For i As Integer = 0 To grdWBTIMGWB.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBTIMGWB.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

        Load_Record()

        tab.Visible = False
        isFormLoading = False

    End Sub

    Private Sub SetFileLocations()
        txtREMOTE_FOLDER.Text = "/www/media/product/"
        txtLOCAL_FOLDER.Text = "S:\Images\"
        txtFILE_EXT.Text = "*.jpg"

        If Not (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
            If Not IsNothing(rowSOTPARM3) Then
                If rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString.EndsWith("\") Then
                    txtLOCAL_FOLDER.Text = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString & String.Empty
                Else
                    txtLOCAL_FOLDER.Text = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString & "\"
                End If
            End If
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Refresh"

            Case "Exit"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                'Load_Record()
                RefreshData()

            Case "Exit"
                Call Mode_Settings(False)
                Me.Close()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Visible = True
                .Groups("Screen Control").Items("Exit").Visible = True
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        'dst.Tables("SOTQRDR1").Rows.Clear()
    End Sub

    Sub Load_Record()
        'Call Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing Data", "")

        EnforceConstraints(False)

        'Fill_Records("SOTQRDR1")

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'Update_Record_TDA("XXXXXXX")
        Call CommitTrans("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        '    Print_Report_Begin()
        '    'frm.CR_params.Add("SUBT", "")
        '    'Fill SOTORDRP records
        '    Fill_Records("SOTQRDR5", ORDR_NO, True)
        '    For Each rowSOTQRDR1 As DataRow In dst.Tables("SOTQRDR1").Select()
        '        If rowSOTQRDR1.Item("ORDR_NO") = ORDR_NO Then
        '            rowSOTQRDR1.Item("ERRORS") = "NEW"
        '        Else
        '            rowSOTQRDR1.Item("ERRORS") = ""
        '        End If
        '    Next
        '    'Generate_Report("SORQRDRO")
        '    Generate_Report("WBRWEBQT", "Quotes Imported From Web", "Re-printed From Quote Maint.")
        '    '    Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWBTIMGWB, "SSB", "Show Filter", "Show GroupBox", "View Local Image", "View FTP Image")
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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTORDR1"
            '    If Not InquiryOnly Then
            '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
            '    End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Selected.Rows.Count = 0 Then
        '    MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
        '    Exit Sub
        'End If

        Dim STYLE_CODE As String = grd.ActiveRow.Cells.Item("STYLE_CODE").Value & String.Empty
        Dim COLOR_CODE As String = grd.ActiveRow.Cells.Item("COLOR_CODE").Value & String.Empty
        If Not (STYLE_CODE.Length = 0 Or COLOR_CODE.Length = 0) Then
            Select Case e.Tool.Key
                Case "View Local Image"
                    Dim frmIMAGE As New TAC.TAFIMGV1(Me, STYLE_CODE, COLOR_CODE, "M", False)
                    With frmIMAGE
                        .ShowDialog(Me)
                    End With
                Case "View FTP Image"
                    Dim frmIMAGE As New TAC.TAFIMGV1(Me, STYLE_CODE, COLOR_CODE, "M", True)
                    With frmIMAGE
                        .ShowDialog(Me)
                    End With
            End Select
        End If
        Update_Record()
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub RefreshData()
        ASCMAIN1.Progress("Refreshing Styles", "")
        btnAddNew.Visible = False
        chkAddBothLocations.Visible = False
        grdWBTIMGWB.Enabled = False
        Fill_Records("WBTIMGWB")

        FillImageStats()

        grdWBTIMGWB.Enabled = True
        btnAddNew.Visible = True
        chkAddBothLocations.Visible = True
        ASCMAIN1.Progress("", "")
    End Sub

    'Private Sub getFTPImages()
    '    Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
    '    Dim IMAGES_FOLDER As String = "C:\"

    '    If Not IsNothing(rowSOTPARM3) Then
    '        If rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString.EndsWith("\") Then
    '            IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
    '        Else
    '            IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString & "\"
    '        End If
    '    End If
    '    Ftp1.User = "regency-rib"
    '    Ftp1.Password = "joydHUJ3"
    '    Ftp1.RemoteHost = "regency-rib.com"
    '    Ftp1.Logon()
    '    Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
    '    Ftp1.RemoteFile = RemoteFolder & "*"
    '    Ftp1.LocalFile = IMAGES_FOLDER & "*"
    '    Ftp1.Overwrite = True
    '    For Each DLFile As String In ImageListDownload
    '        Application.DoEvents()
    '        If ImageGetProgress Then
    '            Ftp1.LocalFile = IMAGES_FOLDER & DLFile
    '            Ftp1.RemoteFile = RemoteFolder & DLFile
    '            ASCMAIN1.Progress("Fetching " & DLFile)
    '            Ftp1.Download()
    '        Else
    '            Exit For
    '        End If
    '    Next
    '    Ftp1.Logoff()
    '    'For Each DELFile As String In ImageListDelete
    '    '    System.IO.File.Delete(IMAGES_FOLDER & DELFile)
    '    '    ASCMAIN1.Progress("Deleting " & DELFile)
    '    'Next
    '    'FillImageStats()
    '    MsgBox("File Sync Complete!", vbOKOnly, "Sync")
    '    'btnGetImages.Visible = False
    'End Sub

    Private Sub GetFileInfo(sender As Object, e As nsoftware.IPWorks.FtpDirListEventArgs) Handles Ftp1.OnDirList
        If FTPImages Then
            If Not e.IsDir Then
                Dim localfile As String = Ftp1.LocalFile
                Dim RemoteFile As String = Ftp1.RemoteFile
                If RemoteFile.Substring(RemoteFile.Length - 1, 1) = "*" Then
                    RemoteFile = RemoteFile.Substring(0, RemoteFile.Length - 1) + e.FileName
                Else
                    Exit Sub
                End If
                Dim FileSize As String = e.FileSize
                If IsNumeric(FileSize) Then
                    If Not ImageListFTP.ContainsKey(e.FileName.Replace(txtREMOTE_FOLDER.Text, "").ToUpper) Then
                        ImageListFTP.Add(e.FileName.Replace(txtREMOTE_FOLDER.Text, "").ToUpper, FileSize)
                    End If
                Else
                    ImageListFTP.Add(e.FileName, 0)
                End If


            End If
        Else
            If Not e.IsDir Then
                Dim localfile As String = Ftp1.LocalFile
                Dim RemoteFile As String = Ftp1.RemoteFile
                If RemoteFile.Substring(RemoteFile.Length - 1, 1) = "*" Then
                    RemoteFile = RemoteFile.Substring(0, RemoteFile.Length - 1) + e.FileName
                Else
                    Exit Sub
                End If
                If localfile.Length > 0 Then
                    If localfile.Substring(localfile.Length - 1, 1) = "*" Then
                        localfile = localfile.Substring(0, localfile.Length - 1) + e.FileName
                        If IO.File.Exists(localfile) Then
                            Dim localDT As Date = IO.File.GetLastWriteTime(localfile)
                            If e.FileTime > localDT Then
                                FileList.Add(localfile, RemoteFile)
                            End If
                        Else
                            FileList.Add(localfile, RemoteFile)
                        End If
                    End If
                Else
                    Exit Sub
                End If
            End If
        End If
    End Sub

    Private Sub FillImageStats()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing File Information", "")
        Application.DoEvents()

        FTPImages = True
        ImageListDownload.Clear()
        ImageListLocal.Clear()
        ImageListFTP.Clear()

        If System.IO.Directory.Exists(txtLOCAL_FOLDER.Text) Then

            Dim LocalFiles As String() = System.IO.Directory.GetFiles(txtLOCAL_FOLDER.Text, txtFILE_EXT.Text)
            For Each LocalFile As String In LocalFiles
                If System.IO.File.Exists(LocalFile) Then
                    Dim fileInfo As New FileInfo(LocalFile)
                    Dim LocalAttrib As FileAttributes = System.IO.File.GetAttributes(LocalFile)
                    ImageListLocal.Add(LocalFile.Replace(txtLOCAL_FOLDER.Text, "").ToUpper(), fileInfo.Length)
                End If
            Next

            Ftp1.User = "regency-rib"
            Ftp1.Password = "joydHUJ3"
            Ftp1.RemoteHost = "regency-rib.com"
            Ftp1.Logon()
            Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
            Ftp1.RemoteFile = txtREMOTE_FOLDER.Text & "*"
            Ftp1.LocalFile = txtLOCAL_FOLDER.Text & "*"
            Ftp1.Overwrite = True
            FileList.Clear()
            Ftp1.ListDirectoryLong()
            Ftp1.Logoff()

            'Dim ImageListFTPtmp As New Dictionary(Of String, Date)

            Dim iCNT As Int64 = dst.Tables("WBTIMGWB").Rows.Count
            Dim cCNT As Int64 = 0
            For Each rowWBTIMGWB As DataRow In dst.Tables("WBTIMGWB").Select("", "STYLE_CODE, COLOR_CODE")
                cCNT += 1
                Dim STYLE_CODE As String = (rowWBTIMGWB.Item("STYLE_CODE").ToString & String.Empty).ToUpper()
                Dim COLOR_CODE As String = (rowWBTIMGWB.Item("COLOR_CODE").ToString & String.Empty).ToUpper()
                Dim FLTR As String = $"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'"
                ASCMAIN1.Progress($"{STYLE_CODE} - {COLOR_CODE}", $"{cCNT} of {iCNT}")
                Dim IMG_NAME As String = $"{STYLE_CODE}-{COLOR_CODE}.JPG"
                If ImageListLocal.ContainsKey(IMG_NAME) Then
                    rowWBTIMGWB.Item("IS_LOCAL") = "1"
                    rowWBTIMGWB.Item("LOCAL_SIZE") = Val(ImageListLocal.Item(IMG_NAME).ToString)
                End If
                If ImageListFTP.ContainsKey(IMG_NAME) Then
                    rowWBTIMGWB.Item("IS_FTP") = "1"
                    rowWBTIMGWB.Item("FTP_SIZE") = Val(ImageListFTP.Item(IMG_NAME).ToString)
                End If
                If Not IsNothing(dst.Tables.Item("WBTSTYLD").Select(FLTR).FirstOrDefault) Then
                    rowWBTIMGWB.Item("IS_WEB") = "1"
                End If
            Next
        Else
            MsgBox("Error with Image Parameters", MsgBoxStyle.Critical, "Parameters")
        End If
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Application.DoEvents()


        FTPImages = False
    End Sub

    Private Sub btnAddNew_Click(sender As Object, e As EventArgs) Handles btnAddNew.Click
        Dim eMsg As New StringBuilder With {.Length = 0}
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Select Image(s) To Add"
        openFileDialog.Filter = $"Files|{txtFILE_EXT.Text}"
        openFileDialog.Multiselect = True ' Allow multiple file selection
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim selectedFiles As String() = openFileDialog.FileNames
            For Each file As String In selectedFiles
                Dim slashAt As Int64 = file.LastIndexOf("\")
                If slashAt = -1 Then
                    eMsg.AppendLine($"Bad File Name: {file}")
                Else
                    Dim fileName As String = file.Substring(slashAt + 1, file.Length - slashAt - 1)
                    Dim dashAt As Int64 = fileName.LastIndexOf("-")
                    Dim dotAt As Int64 = fileName.LastIndexOf(".")
                    If dashAt = -1 Or dotAt = -1 Then
                        eMsg.AppendLine($"Bad File Name: {file}")
                    Else
                        Dim STYLE_CODE As String = fileName.Substring(0, dashAt)
                        Dim COLOR_CODE As String = fileName.Substring(dashAt + 1, fileName.Length - dotAt)
                        Dim EXT As String = fileName.Substring(dotAt + 1, fileName.Length - dotAt - 1)
                        If EXT <> "jpg" Then
                            eMsg.AppendLine($"Bad Extension: {file}")
                        Else
                            Dim fltr As String = $"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'"
                            If dst.Tables("WBTIMGWB").Select(fltr).Count <> 0 Then
                                eMsg.AppendLine($"Not Valid Style / Color: {file}")
                            End If
                        End If
                    End If
                End If
            Next
            If eMsg.Length > 0 Then
                MsgBox(eMsg.ToString, vbCritical, "Problem With Selected File(s)")
            Else
                If chkAddBothLocations.Checked Then
                    Ftp1.User = "regency-rib"
                    Ftp1.Password = "joydHUJ3"
                    Ftp1.RemoteHost = "regency-rib.com"
                    Ftp1.Logon()
                    Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                    Ftp1.Overwrite = True
                End If

                For Each file As String In selectedFiles
                    Dim slashAt As Int64 = file.LastIndexOf("\")
                    Dim fileName As String = file.Substring(slashAt + 1, file.Length - slashAt - 1)
                    System.IO.File.Copy(file, txtLOCAL_FOLDER.Text & fileName)
                    Application.DoEvents()
                    If chkAddBothLocations.Checked Then
                        Ftp1.LocalFile = file
                        Ftp1.RemoteFile = txtREMOTE_FOLDER.Text & fileName
                        Ftp1.Upload()
                    End If
                Next
                If chkAddBothLocations.Checked Then
                    Ftp1.Logoff()
                End If
            End If
        Else
            MsgBox("No files selected.", vbCritical, "What?")
        End If
    End Sub
#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class