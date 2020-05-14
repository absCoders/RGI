Public Class ARFCBDA1
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim REASON_CODE As String
    Dim rowARTREAS1 As DataRow
    Dim RYP0 As String
    Dim RYP1 As String
    Dim ARTCBDAA As String
    Dim ARTCBDA1 As String
    Dim sqlARTCBDA1 As String
    Dim RYP0_Legend As String = ""
    Dim RYP1_Legend As String = ""
    Dim MOS As Integer = 0

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Create_ARTCBDAA("", "")
            TAC.ARCMAIN1.Create_ARTCBDA1(ARTCBDA1, "", "")

            ASCMAIN1.sql = "Select ARTCBDAA.REASON_CODE, ARTREAS1.REASON_DESC" _
                & " from ARTREAS1," & ARTCBDAA & " ARTCBDAA where ARTREAS1.REASON_CODE = ARTCBDAA.REASON_CODE"
            Create_TDA(.Tables.Add, "ARTCBDA2", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCBDAA.*" & vbCrLf _
                & ",ARTCUST1.CUST_NAME,ARTREAS1.REASON_DESC" & vbCrLf _
                & " from ARTCUST1,ARTREAS1," & ARTCBDAA & " ARTCBDAA" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = ARTCBDAA.CUST_CODE" & vbCrLf _
                & "   and ARTREAS1.REASON_CODE (+) = ARTCBDAA.REASON_CODE" & vbCrLf _
                & "   and ARTCBDAA.ACTIVITY = :PARM1"
            Create_TDA(.Tables.Add, "ARTCBDAA", "**", 0, False, "V", 2)
            .Tables("ARTCBDAA").Columns.Add("P00", GetType(System.Decimal), "ISNULL(P01,0)+ISNULL(P02,0)+ISNULL(P03,0)+ISNULL(P04,0)+ISNULL(P05,0)+ISNULL(P06,0)+ISNULL(P07,0)+ISNULL(P08,0)+ISNULL(P09,0)+ISNULL(P10,0)+ISNULL(P11,0)+ISNULL(P12,0)")

            ASCMAIN1.sql = "Select * from ARTREAS1"
            Create_TDA(.Tables.Add, "ARTREAS1", "**", 0, False)

            sqlARTCBDA1 = "Select ARTCBDA1.CUST_CODE CODE_VALUE, ARTCUST1.CUST_NAME DESC_VALUE" _
                & ", Sum (BEG_B) BEG_B, Sum (NEW_B) NEW_B, Sum (APP_B) APP_B, Sum (END_B) END_B" & vbCrLf _
                & ", Sum (BEG_C) BEG_C, Sum (NEW_C) NEW_C, Sum (APP_C) APP_C, Sum (END_C) END_C, Sum (NEW_X) NEW_X" & vbCrLf _
                & " from ARTCUST1," & ARTCBDA1 & " ARTCBDA1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = ARTCBDA1.CUST_CODE" & vbCrLf _
                & " group by ARTCBDA1.CUST_CODE, ARTCUST1.CUST_NAME"
            ASCMAIN1.sql = sqlARTCBDA1
            Create_TDA(.Tables.Add, "ARTCBDA1", "**", 0, False, "", 1)
            With .Tables("ARTCBDA1")
                .Columns.Add("END_ALL", GetType(System.Decimal), "ISNULL(END_B,0) + ISNULL(END_C,0)")
                .Columns.Add("TOTAL", GetType(System.Decimal), "ISNULL(NEW_C,0) + ISNULL(NEW_X,0)")
            End With
        End With

        Fill_Records("ARTREAS1")

        grdARTCBDAA.DataSource = dst.Tables("ARTCBDAA")
        grdARTCBDA1.DataSource = dst.Tables("ARTCBDA1")
        grdARTCBDA2.DataSource = dst.Tables("ARTCBDA2")

        Create_Summary(grdARTCBDAA, "CUST_CODE", "Count")
        Create_Summary(grdARTCBDAA, "REASON_CODE", "Count")

        For I As Integer = 0 To 12
            Dim C As String = "P" & Format(I, "00")
            Create_Summary(grdARTCBDAA, C)
        Next I

        With grdARTCBDAA.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "REASON_CODE", "REASON_DESC", "P00"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
        End With

        With grdARTCBDA1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
            Dim g As UltraWinGrid.UltraGridGroup

            g = .Groups.Add("CODES")
            g.Header.Fixed = True
            g.Header.Caption = "Codes"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.LightGreen
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"CODE_VALUE", "DESC_VALUE"}
                .Columns(COLUMN_NAME).Group = g
            Next
            Create_Summary(grdARTCBDA1, "CODE_VALUE", "Count")

            g = .Groups.Add("CB")
            g.Header.Caption = "AR Chargebacks and CRs On/Account"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.LightBlue
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"BEG_B", "NEW_B", "APP_B", "END_B"}
                .Columns(COLUMN_NAME).Group = g
                .Columns(COLUMN_NAME).Format = "#,##0.00"
                .Columns(COLUMN_NAME).Width = 100
                Create_Summary(grdARTCBDA1, COLUMN_NAME)
            Next

            g = .Groups.Add("CR")
            g.Header.Caption = "Misc Charges and Credits"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.Orange
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"BEG_C", "NEW_C", "APP_C", "END_C"}
                .Columns(COLUMN_NAME).Group = g
                .Columns(COLUMN_NAME).Format = "#,##0.00"
                .Columns(COLUMN_NAME).Width = 100
                Create_Summary(grdARTCBDA1, COLUMN_NAME)
            Next

            g = .Groups.Add("DED")
            g.Header.Caption = "Totals"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.Violet
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"END_ALL", "NEW_X", "TOTAL"}
                .Columns(COLUMN_NAME).Group = g
                .Columns(COLUMN_NAME).Format = "#,##0.00"
                .Columns(COLUMN_NAME).Width = 100
                Create_Summary(grdARTCBDA1, COLUMN_NAME)
            Next
        End With

        ASCMAIN1.sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2" _
            & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        cbeYP0.DataSource = New DataView(DT, "", "OPS_YYYYPP", DataViewRowState.CurrentRows)
        cbeYP1.DataSource = New DataView(DT, "", "OPS_YYYYPP", DataViewRowState.CurrentRows)

        tab1.Tabs("Detail").Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    Validate_Code("CUST_CODE")
                    If EMsg = "" Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Customer Code " & CUST_CODE
                        Else
                            REASON_CODE = ""
                            Absx1.txtFor("REASON_CODE").Text = ""
                        End If
                    End If
                ElseIf Absx1.txtFor("REASON_CODE").Text <> "" Then
                    Validate_Code("REASON_CODE")
                    If EMsg = "" Then
                        REASON_CODE = Absx1.txtFor("REASON_CODE").Text
                        rowARTREAS1 = LookUp("ARTREAS1", REASON_CODE)
                        If rowARTREAS1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Reason Code " & REASON_CODE
                        Else
                            CUST_CODE = ""
                            Absx1.txtFor("CUST_CODE").Text = ""
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If Not Load_Periods() Then
                        Exit Sub
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                If optShowRollForward.Value = "C" Then
                    optShowRollForward.Value = "R"
                Else
                    optShowRollForward.Value = "C"
                End If
                Mode_Settings(False)

            Case "Print"
                Print_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Print").Visible = False ' (InquiryMode Or EntryMode = "V")
                End With
                .Groups("Period Range").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(cbeYP0, ScreenMode)
        Set_Read_Only_for_ctl(cbeYP1, ScreenMode)

        splARTCBDA1.Visible = ScreenMode
        tab0.Visible = Not ScreenMode

        If ScreenMode Then
            grdARTCBDA1.Parent = splARTCBDA1.Panel1
            grdARTCBDAA.Parent = tab1.Tabs("Activity").TabPage
        Else
            Clear_Record()
            grdARTCBDA1.Parent = tab0.Tabs("Roll Forward").TabPage
            grdARTCBDAA.Parent = tab0.Tabs("Activity").TabPage
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ARTCBDA1", "ARTCBDA2", "ARTCBDAA"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("REASON_CODE").Text = ""

        If cbeYP0.Value & "" = "" Then
            cbeYP0.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -11)
            cbeYP1.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0)
            Load_ARTCBDAA()
            Load_ARTCBDA1()
        Else
            Setup_ARTCBDAA()
            Setup_ARTCBDA1()
        End If

        Setup_tab0()

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        'Load_ARTCBDAA()
        Setup_ARTCBDAA()

        If optShowRollForward.Value = IIf(REASON_CODE <> "", "C", "R") Then
            Setup_ARTCBDA1()
        Else
            optShowRollForward.Value = IIf(REASON_CODE <> "", "C", "R")
        End If

        EnforceConstraints(True)

        Setup_tab0()

        SETUP_grdARTCBDA2()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCBDAA, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdARTCBDA1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Inquiry")
        Load_Popup_Menu(grdARTCBDA2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Inquiry")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdARTCBDA1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String
                If grd.Name = "grdARTCBDA1" Then
                    CUST_CODE = grd.ActiveRow.Cells("CODE_VALUE").Value
                Else
                    CUST_CODE = grd.ActiveRow.Cells("CUST_CODE").Value
                End If
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    Context_Launch("Select", CUST_CODE, e.Tool.Key, "ARFCINQ1")
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        Absx1.txtFor("REASON_CODE").Text = ""
                        Click_Command("View")
                    End If
                End If
            Case "REASON_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("REASON_CODE").Text <> "" Then
                        Absx1.txtFor("CUST_CODE").Text = ""
                        Click_Command("View")
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Absx1.txtFor("REASON_CODE").Text = ""
                Click_Command("View")
            Case "REASON_CODE"
                Absx1.txtFor("CUST_CODE").Text = ""
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr Is Nothing Then
                            Absx1.txtFor("CUST_CODE").Text = ""
                        End If
                    End If
                End If
            Case "REASON_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("REASON_CODE").Text <> "" Then
                        LookUp("ARTREAS1", Absx1.txtFor("REASON_CODE").Text)
                        If cdr Is Nothing Then
                            Absx1.txtFor("REASON_CODE").Text = ""
                        End If
                    End If
                End If
        End Select
    End Sub
#End Region

#Region "grdARTCBDA1"

    Private Sub grdARTCBDA1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdARTCBDA1.AfterRowActivate
        If ScreenMode Then SETUP_grdARTCBDA2()
    End Sub

#End Region

    Sub SETUP_grdARTCBDA2()
        If grdARTCBDA1.ActiveRow Is Nothing OrElse Not grdARTCBDA1.ActiveRow.IsDataRow Then
            splARTCBDA1.Panel2Collapsed = True
        Else
            grdARTCBDA2.Text = "Roll Forward Details for " & ""
            splARTCBDA1.Panel2Collapsed = False
        End If
    End Sub

    Private Sub grdARTCBDAA_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCBDAA.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
            Absx1.txtFor("REASON_CODE").Text = ""
            Click_Command("View")
        End If
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Report")

        'Print_Report_Begin()
        'CR_params.Add("NOTES", "1")
        'Generate_Report("BMRLIST1", "Bill of Materials", "")
        'Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Load_Periods() As Boolean
        RYP0 = cbeYP0.Value
        RYP1 = cbeYP1.Value

        Dim P As Integer = 11
        If RYP0 <> "" Then P = ASCMAIN1.Period_Diff(RYP0, RYP1)

        If P < 0 Or P > 11 Then
            MsgBox("Periods range may span from 1 to 12 months")
            Return False
        Else

            RYP0_Legend = ""
            RYP1_Legend = ""
            MOS = 0

            With grdARTCBDAA.DisplayLayout.Bands(0)
                For I As Integer = 1 To 12
                    Dim RYP As String = ASCMAIN1.Period_Calc(RYP0, (I - 1))
                    Dim C As String = "P" & Format(I, "00")
                    If RYP <= RYP1 Then
                        .Columns(C).Hidden = False
                        Dim LEGEND = ASCMAIN1.Get_Legend(RYP)
                        If I = 1 Then RYP0_Legend = LEGEND
                        If RYP = RYP1 Then RYP1_Legend = LEGEND
                        .Columns(C).Header.Caption = Mid(LEGEND, 10, 6)
                        MOS += 1
                    Else
                        .Columns(C).Hidden = True
                    End If
                Next
            End With

            Return True
        End If
    End Function

    Sub Load_ARTCBDAA()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Activity")

        Load_Periods()
        Create_ARTCBDAA(RYP0, RYP1)
        Setup_ARTCBDAA()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Sub Load_ARTCBDA1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Roll Forward")

        TAC.ARCMAIN1.Create_ARTCBDA1(ARTCBDA1, RYP0, RYP1)
        Setup_ARTCBDA1()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Sub Create_ARTCBDAA(RYP0 As String, RYP1 As String)

        Dim sqlSum As String = ""
        For I As Integer = 1 To 12
            Dim RYP As String = ASCMAIN1.Period_Calc(RYP0, (I - 1))
            Dim C As String = "P" & Format(I, "00")
            If RYP <= RYP1 Or ARTCBDAA = "" Then
                sqlSum &= ", SUM (DECODE(ARTPYMT1.OPS_YYYYPP,'" & RYP & "',ARTPYMT5.GL_DIST_AMT,0)) " & C & vbCrLf
            Else
                sqlSum &= ", 0 " & C & vbCrLf
            End If
        Next

        ASCMAIN1.sql = "Select ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE" & vbCrLf _
            & ", DECODE(ARTPYMT5.CHARGEBACK_IND,'1','B','D') ACTIVITY" & vbCrLf _
            & sqlSum _
            & "  from ARTPYMT1,ARTPYMT2,ARTPYMT5" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & RYP1 & "'"
        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and ARTPYMT2.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        End If
        If Absx1.txtFor("REASON_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and ARTPYMT5.REASON_CODE = '" & Absx1.txtFor("REASON_CODE").Text & "'"
        End If
        If ARTCBDAA = "" Then
            ASCMAIN1.sql &= "   and ROWNUM < 1"
        End If
        ASCMAIN1.sql &= " group by CUST_CODE, ARTPYMT5.REASON_CODE" & vbCrLf _
            & ", DECODE(ARTPYMT5.CHARGEBACK_IND,'1','B','D')"

        If ARTCBDAA = "" Then
            ARTCBDAA = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCBDAA)

            ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select SOTINVH1.CUST_CODE, SOTINVH1.REASON_CODE" & vbCrLf _
           & ", 'C' ACTIVITY" & vbCrLf _
           & Replace(Replace(Replace(Replace(sqlSum, "ARTPYMT1", "SOTINVH1"), "OPS_YYYYPP", "ORDR_YYYYPP_UPDATED"), "ARTPYMT5", "SOTINVH1"), "GL_DIST_AMT", "INV_SALES") _
           & "  from SOTINVH1" & vbCrLf _
           & " where SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" & vbCrLf _
           & "   and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'" & vbCrLf _
           & "   and SOTINVH1.ORDR_TYPE_CODE in ('TOP','DIF')"
        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        End If
        If Absx1.txtFor("REASON_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and SOTINVH1.REASON_CODE = '" & Absx1.txtFor("REASON_CODE").Text & "'"
        End If
        ASCMAIN1.sql &= " group by SOTINVH1.CUST_CODE, SOTINVH1.REASON_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & ASCMAIN1.sql)

    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        If Load_Periods() Then
            Load_ARTCBDAA()
            Load_ARTCBDA1()
        End If
    End Sub

    Sub Setup_ARTCBDA1()

        Dim SQLW As String = ""

        If optShowRollForward.Value = "C" Then
            ASCMAIN1.sql = sqlARTCBDA1
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption = "Customer"
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = "Name"
            If (EntryMode = "V") Then SQLW = " ARTCBDA1.REASON_CODE = '" & REASON_CODE & "'"
        Else
            ASCMAIN1.sql = _
                Replace( _
                Replace( _
                Replace(sqlARTCBDA1, _
                    "CUST_CODE", "REASON_CODE"), _
                    "CUST_NAME", "REASON_DESC"), _
                    "ARTCUST1", "ARTREAS1")
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption = "Reason"
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = "Description"
            If (EntryMode = "V") Then SQLW = " ARTCBDA1.CUST_CODE = '" & CUST_CODE & "'"
        End If

        If SQLW <> "" Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, " where ", " where " & SQLW & " and ")

        Fill_Records("ARTCBDA1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdARTCBDA1, "CODE_VALUE")

        grdARTCBDA1.Text = "Roll Forward" & _
            IIf(EntryMode = "V", " for " & IIf(optShowRollForward.Value = "C", REASON_CODE, CUST_CODE), "") _
            & ": " & RYP0_Legend & " thru " & RYP1_Legend
    End Sub

    Sub Setup_ARTCBDAA()
        Fill_Records("ARTCBDAA", optShowActivity.Value)
        Sort_grdColumns(grdARTCBDAA, "CUST_CODE,REASON_CODE")

        With grdARTCBDAA.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Hidden = (EntryMode = "V") And (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("CUST_NAME").Hidden = (EntryMode = "V") And (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("REASON_CODE").Hidden = (EntryMode = "V") And (Absx1.txtFor("REASON_CODE").Text <> "")
            .Columns("REASON_DESC").Hidden = (EntryMode = "V") And (Absx1.txtFor("REASON_CODE").Text <> "")
        End With

        Dim DVW As DataView = DirectCast(grdARTCBDAA.DataSource, DataTable).DefaultView
        If EntryMode = "V" Then
            If CUST_CODE <> "" Then
                DVW.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"
            Else
                DVW.RowFilter = "REASON_CODE = '" & REASON_CODE & "'"
            End If
        Else
            DVW.RowFilter = ""
        End If

        grdARTCBDAA.Visible = True
        grdARTCBDAA.Text = optShowActivity.Text & " Activity for the " & CStr(MOS) & " Months " & RYP0_Legend & " to " & RYP1_Legend
    End Sub

    Private Sub optShowActivity_ValueChanged(sender As Object, e As EventArgs) Handles optShowActivity.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_ARTCBDAA()
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        optShowActivity.Visible = tab0.SelectedTab.Key = "Activity"
        optShowRollForward.Visible = tab0.SelectedTab.Key = "Roll Forward"
    End Sub

    Private Sub optShowRollForward_ValueChanged(sender As Object, e As EventArgs) Handles optShowRollForward.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_ARTCBDA1()
    End Sub

    Private Sub grdARTCBDA1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCBDA1.DoubleClickRow

        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            Dim CODE_VALUE As String = e.Row.Cells("CODE_VALUE").Value & ""
            If optShowRollForward.Value = "C" Then
                Absx1.txtFor("CUST_CODE").Text = CODE_VALUE
                Absx1.txtFor("REASON_CODE").Text = ""
            Else
                Absx1.txtFor("CUST_CODE").Text = CODE_VALUE
                Absx1.txtFor("REASON_CODE").Text = ""
            End If
            Click_Command("View")
        End If
      
    End Sub

    Private Sub grdARTCBDA1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTCBDA1.InitializeLayout

    End Sub
End Class