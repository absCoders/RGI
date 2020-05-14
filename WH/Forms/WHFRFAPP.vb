Public Class WHFRFAPP
    Dim TXTs As New Dictionary(Of Integer, UltraWinEditors.UltraTextEditor)
    Dim LBLs As New Dictionary(Of Integer, Misc.UltraLabel)
    Dim Cs As New Dictionary(Of Integer, WHC.WHCRF000)
    Dim Bs As New Dictionary(Of Integer, Misc.UltraButton())

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify a Whse"
                Else
                    If Absx1.txtFor("APP_ID").Text = "" Then
                        EMsg &= vbCr & "You must specify an App ID"
                    Else
                        Dim row As DataRow = LookUp("WHTGUNA1", Absx1.txtFor("APP_ID").Text)
                        If row Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for App ID " & Absx1.txtFor("APP_ID").Text
                        End If
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

            Case "Load"

                Start_App()

                'EntryMode = "L"
                'Load_Record()
                'Mode_Settings(True) 

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                '  .Groups("Screen Control").Items("Send").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        'For Each TABLE_NAME As String In New String() _
        '    {"SOTSHIPX", "SOTPICK1", "SOTPICKX", "WHTLPXN1"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        '  Load_SOTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")
        BeginTrans()
        CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "APP_ID"
                sql_where = "USE_CLASS = '1'"
        End Select
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        ' Load_Popup_Menu(grdICTWHSEX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
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

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "APP_ID"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "APP_ID"
                Click_Command("Load")
        End Select
    End Sub
#End Region

    Sub RF_APP(THREAD_NO As Integer, APP_ID As String, APP_DESC As String, TXT As UltraWinEditors.UltraTextEditor)

        'Dim C As New WHCRF001(THREAD_NO, APP_ID, APP_DESC)
        'TXT.Text = C.Hello
        'AddHandler C.RespondToScan, AddressOf Display_Text

        'Dim RESPONSE As String = ""
        'Cs.Add(THREAD_NO, C)
    End Sub

    Sub Start_App()
        Dim APP_ID As String = Absx1.txtFor("APP_ID").Text
        Dim APP_DESC As String = Absx1.txtFor("APP_DESC").Text
        Dim rowWHTGUNA1 As DataRow = LookUp("WHTGUNA1", APP_ID)
        Dim PROCEDURE_NAME As String = rowWHTGUNA1.Item("PROCEDURE_NAME") & ""
        Dim GUN_PARAM As String = rowWHTGUNA1.Item("PICK_TYPE") & ""
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

        Dim THREAD_NO As Integer = tabMain.Tabs.Count
        tabMain.Tabs.Add(CStr(THREAD_NO))

        Dim txt As New UltraWinEditors.UltraTextEditor
        Dim spl As New SplitContainer
        Dim lbl As New Misc.UltraLabel

        With tabMain.Tabs(THREAD_NO)

            txt.Multiline = True
            txt.Scrollbars = ScrollBars.Vertical

            .TabPage.Controls.Add(spl)
            spl.Dock = DockStyle.Fill

            spl.Panel1.Controls.Add(lbl)
            LBLs(THREAD_NO) = lbl
            lbl.Height = 150
            lbl.Width = 300

            spl.Panel2.Controls.Add(txt)
            txt.Visible = True
            txt.Dock = DockStyle.Fill

            spl.SplitterDistance = spl.Height * 0.8
            spl.Tag = THREAD_NO

            Dim txt2 As New UltraWinEditors.UltraTextEditor
            spl.Panel1.Controls.Add(txt2)
            txt2.Tag = THREAD_NO
            txt2.Top = lbl.Height
            AddHandler txt2.KeyDown, AddressOf txt2_KeyDown

            Dim btns(5) As Misc.UltraButton
            For i As Integer = 1 To 5
                Dim btn As New Misc.UltraButton
                spl.Panel1.Controls.Add(btn)
                btn.Height = 25
                btn.Width = 100
                btn.Top = txt2.Top + txt2.Height + 10 + (i - 1) * (btn.Height + 1)
                btn.Visible = False
                btns(i) = btn
                AddHandler btn.Click, AddressOf UltraButton1_Click
            Next
            Bs.Add(THREAD_NO, btns)

            .Text = APP_ID & ":" & APP_DESC

            txt2.Focus()

        End With

        ' Dim APP = New System.Threading.Thread(Sub() Me.RF_APP(T, APP_ID, APP_DESC, txt))
        'APP.Start()
        TXTs.Add(THREAD_NO, txt)

        Dim GUN_LOC As String = "GUN" & Format(THREAD_NO, "000")
        GUN_LOC = "99-G00-A"

        Dim C As WHC.WHCRF000 = WHC.WHCFACT1.CreateWhcClass(PROCEDURE_NAME, New WHC.GunEnvironment With
            {.DBS_COMPANY = ASCMAIN1.DBS_COMPANY, .DBS_SERVER = ASCMAIN1.DBS_SERVER, .DBS_PASSWORD = ASCMAIN1.DBS_PASSWORD,
             .THREAD_NO = THREAD_NO, .APP_ID = APP_ID, .APP_DESC = APP_DESC,
             .USER_ID = ASCMAIN1.USER_ID, .GUN_LOC = GUN_LOC, .PICK_TYPE = GUN_PARAM, .WHSE_CODE = WHSE_CODE})

        'Dim C As New WHC.WHCRF002(New WHC.GunEnvironment With _
        '    {.DBS_COMPANY = ASCMAIN1.DBS_COMPANY, .DBS_SERVER = ASCMAIN1.DBS_SERVER, .DBS_PASSWORD = ASCMAIN1.DBS_PASSWORD, _
        '     .THREAD_NO = THREAD_NO, .APP_ID = APP_ID, .APP_DESC = APP_DESC, _
        '     .USER_ID = ASCMAIN1.USER_ID, .GUN_LOC = "GUN" & Format("000"), .PICK_TYPE = "C", .WHSE_CODE = "NJE"})

        Cs.Add(THREAD_NO, C)
        AddHandler C.RespondToScan, AddressOf Display_Text

        Display_Text(THREAD_NO, C.Hello)

        Setup_Tab()

    End Sub

    Sub Display_Text(THREAD_NO As Integer, TXT As String)

        Dim BTNs() As String = Split(TXT, "|")
        If BTNs.Length > 1 Then
            TXT = BTNs(0)
        End If

        For i As Integer = 1 To 5
            With Bs(THREAD_NO)(i)
                If BTNs.Length - 1 < i OrElse BTNs(i) = "" Then
                    .Visible = False
                Else
                    .Visible = True
                    .Text = BTNs(i)
                End If
            End With
        Next

        LBLs(THREAD_NO).TEXT = TXT

        TXTs(THREAD_NO).Text &= vbCrLf & "Thread " & THREAD_NO & ":" & TXT
        TXTs(THREAD_NO).SelectionStart = TXTs(THREAD_NO).Text.Length - 1
        TXTs(THREAD_NO).ScrollToCaret()

        If Split(TXT, vbCrLf)(0) = "EXIT" Then
            grd.DataSource = Nothing
            Cs(THREAD_NO) = Nothing
            BTNs(THREAD_NO) = Nothing
            tabMain.Tabs.Remove(tabMain.Tabs(THREAD_NO))
        End If
    End Sub

    Sub txt2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Try
                Dim txt2 As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
                Dim THREAD_NO As Integer = txt2.Tag
                Cs(THREAD_NO).GetResponseToScan(txt2.Text)
                txt2.Text = ""
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub UltraButton1_Click(sender As System.Object, e As System.EventArgs)
        Dim btn As Misc.UltraButton = DirectCast(sender, Misc.UltraButton)
        Dim THREAD_NO As Integer = Val(btn.Parent.Tag)
        Cs(THREAD_NO).GetResponseToScan(btn.Text)
    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If tabMain.SelectedTab Is Nothing Then Exit Sub
        setup_Tab()
    End Sub

    Sub Setup_Tab()

        Dim THREAD_NO As Integer = Val(tabMain.SelectedTab.Key)
        If Not Cs.ContainsKey(THREAD_NO) Then Exit Sub
        grd.DataSource = Cs(THREAD_NO).tbl
        grd.Text = tabMain.SelectedTab.Text
    End Sub

    Private Sub UltraTextEditor1_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor1.ValueChanged

    End Sub
End Class