Imports ABSolution

Public Class SOFPICKP

    Private pickQuery As String = String.Empty
    Private WithEvents timerFocus As New System.Windows.Forms.Timer
    Private gridFormatted As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        InquiryMode = MENU_ITEM_OBJECT = "SOFPICKI"

        With dst

            pickQuery = "Select SOTPICK1.*, SOTORDR1.ORDR_DATE, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME," _
                    & " SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_STORE_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" _
                    & " FROM SOTORDR1, SOTPICK1" _
                    & " WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
                    & " AND PICK_STATUS = 'P'"

            Create_TDA(.Tables.Add, "SOTPICK1", "*", , , , , "PICK_PICKER,PICK_PACKED")

            .Tables("SOTPICK1").Columns.Add("ORDR_DATE", GetType(System.DateTime))
            .Tables("SOTPICK1").Columns.Add("CUST_CODE", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("CUST_NAME", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("CUST_STORE_NO", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("CUST_STORE_NAME", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("ORDR_CUST_PO", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("ORDR_SHIP_DATE", GetType(System.DateTime))
            .Tables("SOTPICK1").Columns.Add("ORDR_CANCEL_DATE", GetType(System.DateTime))

        End With

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        Create_Summary(grdSOTPICK1, "CUST_CODE", "Count")

        splSOTPICK1.Panel1Collapsed = InquiryMode

        If InquiryMode Then
            grdSOTPICK1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        Else
            grdSOTPICK1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
        End If

        timerFocus.Interval = 500
        timerFocus.Stop()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Not InquiryMode Then
                    Validate_Code("USER_ID")
                End If

            Case "Cancel"

            Case "Update"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Refresh"
                EntryMode = "E"
                Load_Record()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode

                .Groups("Screen Control").Items("Update").Visible = Not InquiryMode
                .Groups("Screen Control").Items("Refresh").Visible = InquiryMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        Set_Read_Only(splUser.Panel1, (ScreenMode AndAlso Not InquiryMode))
        Set_Read_Only(splUser.Panel2, (Not ScreenMode AndAlso Not InquiryMode))

        grdSOTPICK1.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SOTPICK1").Rows.Clear()
        EnforceConstraints(True)

        Absx1.txtFor("USER_ID").Clear()
        Absx1.txtFor("PICK_NO").Clear()

    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        If InquiryMode Then
            Fill_Records("SOTPICK1", String.Empty, True, pickQuery)
        End If

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            For Each row As DataRow In dst.Tables("SOTPICK1").Select("")
                row.Item("PICK_PACKED") = DateTime.Now
            Next

            Update_Record_TDA("SOTPICK1")
            CommitTrans()
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region


#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICK1, "SS", "Show Filter", "Show GroupBox")
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

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)

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

    Public Overrides Sub txt_KeyDown(sender As Object, e As KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Select Case Absx1.GetABSColumnName(sender)

            Case "PICK_NO"

                Try
                    If e.KeyCode = Windows.Forms.Keys.Enter Then
                        EMsg = String.Empty
                        Absx1.txtFor("PICK_NO").Text = Absx1.txtFor("PICK_NO").Text.Trim
                        Validate_Code("PICK_NO")

                        If EMsg.Length > 0 Then
                            MessageBox.Show(EMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Absx1.txtFor("PICK_NO").Clear()
                            Exit Sub
                        End If

                        If dst.Tables("SOTPICK1").Select("PICK_NO = '" & Absx1.txtFor("PICK_NO").Text & "'").Length > 0 Then
                            MessageBox.Show("The scanned Pick Ticket is already listed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Absx1.txtFor("PICK_NO").Clear()
                            Exit Sub
                        End If

                        If cdr.Item("PICK_STATUS") & String.Empty <> "P" Then
                            MessageBox.Show("The scanned Pick Ticket does not have a status of In Pick.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Absx1.txtFor("PICK_NO").Clear()
                            Exit Sub
                        End If

                        If cdr.Item("PICK_PICKER") & String.Empty <> String.Empty Then
                            Dim zMsg As String = "The Pick Ticket is already assigned to user (" & cdr.Item("PICK_PICKER") & ")." _
                                                 & Environment.NewLine & Environment.NewLine _
                                                 & "Do you want to continue?"
                            If MessageBox.Show(zMsg, "Pick Ticket", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                                Absx1.txtFor("PICK_NO").Clear()
                                Exit Sub
                            End If
                        End If

                        ASCMAIN1.sql = pickQuery.Replace("PICK_STATUS = 'P'", "PICK_NO = '" & Absx1.txtFor("PICK_NO").Text & "'")
                        Fill_Records("SOTPICK1", String.Empty, False, ASCMAIN1.sql)

                        dst.Tables("SOTPICK1").Select("PICK_NO = '" & Absx1.txtFor("PICK_NO").Text & "'")(0).Item("PICK_PICKER") = Absx1.txtFor("USER_ID").Text

                        If Not gridFormatted AndAlso dst.Tables("SOTPICK1").Rows.Count > 0 Then
                            System.Threading.Thread.Sleep(500)
                            ASCMAIN1.Progress("Formatting Grid Data")
                            grdSOTPICK1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                            ASCMAIN1.Progress("")
                            gridFormatted = True
                        End If

                        Absx1.txtFor("PICK_NO").Clear()
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message)

                Finally
                    timerFocus.Start()
                End Try

        End Select
    End Sub

    Private Sub timerFocus_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles timerFocus.Tick
        timerFocus.Stop()
        Absx1.txtFor("PICK_NO").Focus()
    End Sub

    Private Sub grdSOTPICK1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTPICK1.InitializeLayout
        If InquiryMode Then
            e.Layout.UseFixedHeaders = True
            e.Layout.Bands(0).Columns("CUST_CODE").Header.Fixed = True
            e.Layout.Bands(0).Columns("CUST_NAME").Header.Fixed = True
            e.Layout.Bands(0).Columns("CUST_STORE_NO").Header.Fixed = True
            e.Layout.Bands(0).Columns("CUST_STORE_NAME").Header.Fixed = True
        End If
    End Sub

    Private Sub grdSOTPICK1_VisibleChanged(sender As Object, e As EventArgs) Handles grdSOTPICK1.VisibleChanged

        If Not gridFormatted AndAlso dst.Tables.Contains("SOTPICK1") AndAlso dst.Tables("SOTPICK1").Rows.Count > 0 Then
            ASCMAIN1.Progress("Formatting Grid Data")
            grdSOTPICK1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
            gridFormatted = True
            ASCMAIN1.Progress("")
        End If
    End Sub
End Class