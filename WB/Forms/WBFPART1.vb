Imports System.Text
Imports System.IO

Public Class WBFPART1
    Dim InquiryOnly As Boolean = False
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp
    Dim rowWBTPARM1 As DataRow = LookUp("WBTPARM1", "Z")

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst
            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTPART1 WHERE PART_NAME = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTPART1", "**", 0, True, "V", 1)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM WBTPART2 WHERE PART_NAME = :PARM1")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "WBTPART2", "**", 0, True, "V", 2)
        End With

        tab.Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "New"
                Dim PART_NAME As String = Absx1.txtFor("PART_NAME").Text
                Dim SQLS As New System.Text.StringBuilder
                SQLS.Length = 0
                SQLS.AppendLine(String.Format("SELECT COUNT(*) FROM WBTPART1 WHERE PART_NAME = '{0}'", PART_NAME))
                ASCMAIN1.sql = SQLS.ToString()
                Dim PCount As Int16 = Val(ASCDATA1.GetDataValue)
                If PCount = 1 Then
                    EMsg &= vbCr & "This Page Already Exists."
                End If
                If Absx1.txtFor("PART_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "You Must Provide A Part Name To Create."
                End If
            Case "Edit"
                If Absx1.txtFor("PART_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "You Must Provide A Part Name To Edit."
                End If
            Case "Cancel"

            Case "Update"
                If Absx1.txtFor("PART_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "Part Name Can Not Be Blank."
                End If
                If Absx1.txtFor("PART_PAGE_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "Page Name Can Not Be Blank."
                End If
                If Not Absx1.txtFor("PART_PAGE_NAME").Text.EndsWith(".HTML") Then
                    EMsg &= vbCr & "Page Name Must End With .HTML"
                End If
                If Absx1.txtFor("PART_DESC").Text.Length = 0 Then
                    EMsg &= vbCr & "Page Description Can Not Be Blank."
                End If
            Case "Done"
                Mode_Settings(False)
            Case "Upload"
                If Absx1.txtFor("PART_PAGE_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "Page Name Can Not Be Blank."
                End If
                If Not Absx1.txtFor("PART_PAGE_NAME").Text.EndsWith(".HTML") Then
                    EMsg &= vbCr & "Page Name Must End With .HTML"
                End If
                If rowWBTPARM1.Item("WB_PARM_SITE_IP").ToString & "" = "" Then
                    EMsg &= vbCr & "Parameter Missing For Site Address"
                End If
                If rowWBTPARM1.Item("WB_PARM_SITE_USER").ToString & "" = "" Then
                    EMsg &= vbCr & "Parameter Missing For Site User"
                End If
                If rowWBTPARM1.Item("WB_PARM_SITE_PWD").ToString & "" = "" Then
                    EMsg &= vbCr & "Parameter Missing For Site Password"
                End If
                If EMsg = "" Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Upload File"
                    Dim iMSG As New System.Text.StringBuilder
                    iMSG.AppendLine("This Wil Upload Your Changes And Over-Write")
                    iMSG.AppendLine("The Existing Files On The Server.  Are You")
                    iMSG.AppendLine("Sure You Want To Proceed?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.No Then
                        EMsg &= vbCr & "Upload Aborted"
                    End If
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
                EntryMode = "N"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Update"
                Call CreateWBTPART2()
                Call Update_Record()
                Call Mode_Settings(False)
            Case "Cancel", "Done"
                Call Mode_Settings(False)
            Case "Upload"
                Call ftp_File()
                MsgBox("Upload Complete.  Save Your Work!", MsgBoxStyle.Exclamation, "Upload")
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Edit").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode

                .Groups("Screen Control").Items("Update").Visible = ScreenMode
                .Groups("Screen Control").Items("Cancel").Visible = ScreenMode
                .Groups("Screen Control").Items("Upload").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If Not tf Then
            Call Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        dst.Tables("WBTPART1").Rows.Clear()
        dst.Tables("WBTPART2").Rows.Clear()
        txtPART_HTLM.Text = ""
        txtPART_NAME.Text = ""
        txtPART_PAGE_NAME.Text = ""
        txtPART_DESC.Text = ""
        txtPART_NAME.ReadOnly = False
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            dst.Tables("WBTPART1").Clear()
            dst.Tables("WBTPART2").Clear()
            Dim rowWBTPART1 As DataRow
            rowWBTPART1 = dst.Tables("WBTPART1").NewRow()
            rowWBTPART1.Item("PART_NAME") = txtPART_NAME.Text
            dst.Tables("WBTPART1").Rows.Add(rowWBTPART1)
        Else
            Call Fill_Records("WBTPART1", Absx1.txtFor("PART_NAME").Text, True)
            Call Fill_Records("WBTPART2", Absx1.txtFor("PART_NAME").Text, True)
        End If

        EnforceConstraints(True)

        For Each rowWBTPART2 As DataRow In dst.Tables("WBTPART2").Select()
            txtPART_HTLM.Text = txtPART_HTLM.Text + rowWBTPART2.Item("PART_HTLM")
        Next

        If dst.Tables("WBTPART1").Rows.Count = 1 Then
            txtPART_PAGE_NAME.Text = dst.Tables("WBTPART1").Rows(0).Item("PART_PAGE_NAME") & ""
            txtPART_DESC.Text = dst.Tables("WBTPART1").Rows(0).Item("PART_DESC") & ""
        End If

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
        txtPART_NAME.ReadOnly = True
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()
        'For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        'Next
        'Call CommitTrans("Order / Quote Deleted")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Update_Record_TDA("WBTPART1")
        Update_Record_TDA("WBTPART2", String.Format("PART_NAME = '{0}'", Absx1.txtFor("PART_NAME").Text))
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        'Print_Report_Begin()
        'Generate_Report("SORORDRO")
        'Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSOTORDRX, "SSB", "Show Filter", "Show GroupBox")
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

        Select Case e.Tool.Key
            'Case "Edit Ship To"
            '    If Not InquiryOnly Then
            '        MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
            '    End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PART_NAME"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "PART_NAME"
                txtPART_NAME.Text = txtPART_NAME.Text.ToUpper
            Case "PART_PAGE_NAME"
                txtPART_PAGE_NAME.Text = txtPART_PAGE_NAME.Text.ToUpper
            Case "PART_DESC"
                txtPART_DESC.Text = txtPART_DESC.Text.ToUpper
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
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
    Public Sub CreateWBTPART2()
        If dst.Tables("WBTPART1").Rows.Count = 1 Then
            dst.Tables("WBTPART1").Rows(0).Item("PART_PAGE_NAME") = txtPART_PAGE_NAME.Text
            dst.Tables("WBTPART1").Rows(0).Item("PART_DESC") = txtPART_DESC.Text
        End If
        dst.Tables("WBTPART2").Clear()
        If txtPART_HTLM.Text.Length > 0 Then
            Dim TotalLines As Double = txtPART_HTLM.Text.Length / 2000
            Dim FullLines As Double = Math.Truncate(TotalLines)
            If TotalLines > FullLines Then
                TotalLines = FullLines + 1
            Else
                TotalLines = FullLines
            End If
            Dim PART_HTLM As String = ""
            Dim LastStart As Integer = 0
            For i As Integer = 1 To TotalLines
                If i = TotalLines Then
                    PART_HTLM = txtPART_HTLM.Text.Substring(LastStart)
                Else
                    PART_HTLM = txtPART_HTLM.Text.Substring(LastStart, 2000)
                    LastStart = LastStart + 2000
                End If
                Dim rowWBTPART2 As DataRow = dst.Tables("WBTPART2").NewRow
                rowWBTPART2.Item("PART_NAME") = Absx1.txtFor("PART_NAME").Text
                rowWBTPART2.Item("PART_LNO") = i
                rowWBTPART2.Item("PART_HTLM") = PART_HTLM
                dst.Tables("WBTPART2").Rows.Add(rowWBTPART2)
            Next
        End If

    End Sub

    Private Sub BuildFTPFile()

    End Sub

    Sub ftp_File()
        'Stop 'Need to get the correct Specs from Parameters and test.
        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        Dim TransferFile As String = String.Format("{0}{1}", ASCMAIN1.Folders("Temp"), txtPART_PAGE_NAME.Text)
        Dim fs As FileStream = File.Create(TransferFile)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(txtPART_HTLM.Text)
        fs.Write(info, 0, info.Length)
        fs.Close()
        Ftp1.User = rowWBTPARM1.Item("WB_PARM_SITE_USER").ToString
        Ftp1.Password = rowWBTPARM1.Item("WB_PARM_SITE_PWD").ToString
        Ftp1.RemoteHost = "69.39.227.201"
        'Ftp1.RemotePath = "www/partials"
        Ftp1.ChangeRemotePath("www/partials")
        Ftp1.Logon()
        'Ftp1.TransferMode = nsoftware.IPWorks.FTPTransferModes.tmBinary
        Ftp1.ChangeTransferMode(nsoftware.IPWorks.FTPTransferModes.tmBinary)
        Ftp1.LocalFile = TransferFile
        Ftp1.RemoteFile = (txtPART_PAGE_NAME.Text).ToLower()
        Ftp1.Overwrite = True
        Ftp1.Upload()
        Ftp1.Logoff()
    End Sub
#End Region

End Class