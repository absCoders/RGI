Public Class ICFSHOW1
    Private SQL As New System.Text.StringBuilder With {.Length = 0}
    Private IMAGE_PATH As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        If ASCMAIN1.Running_in_VS Then
            Stop
            IMAGE_PATH = "\\192.168.180.35\g\VAN\images\"
        Else
            IMAGE_PATH = (ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & "").Replace("G:", "R:")
            If Not IMAGE_PATH.EndsWith("\") Then
                IMAGE_PATH = IMAGE_PATH & "\"
            End If
        End If

        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT * FROM ICTSTYL1 WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "V", 1)

            SQL.Length = 0
            SQL.AppendLine("Select X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC")
            SQL.AppendLine(", SUM(X.BEG) BEG, SUM(X.SHP) SHP, SUM(X.RTN) RTN, SUM(X.REC) REC")
            SQL.AppendLine(", SUM(X.ADJ) ADJ, SUM(X.XFR) XFR, SUM(X.PHY) PHY, SUM(X.ON_HAND) ON_HAND")
            SQL.AppendLine(", SUM(X.ON_ORDER) ON_ORDER, SUM(X.TRAN) TRAN, SUM(X.OPEN) OPEN")
            SQL.AppendLine(", SUM(X.PICK) PICK, SUM(X.ALLO) ALLO, SUM(X.COMM) COMM, SUM(X.PROD) PROD")
            SQL.AppendLine(", MAX(UPC_CODE) UPC_CODE, MAX(STYLE_COLOR_STATUS) STYLE_COLOR_STATUS from ICTCOLR1, ICTSTYL1, (")
            SQL.AppendLine("(Select ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE")
            SQL.AppendLine(", SUM(ICTSTAT1.WHSE_QTY_BEG) BEG")
            SQL.AppendLine(", SUM(ICTSTAT1.WHSE_QTY_SHP) SHP, SUM(ICTSTAT1.WHSE_QTY_RTN) RTN")
            SQL.AppendLine(", SUM(ICTSTAT1.WHSE_QTY_REC) REC, SUM(ICTSTAT1.WHSE_QTY_ADJ) ADJ")
            SQL.AppendLine(", SUM(ICTSTAT1.WHSE_QTY_XFR) XFR, SUM(ICTSTAT1.WHSE_QTY_PHY) PHY")
            SQL.AppendLine(", SUM(0) ON_HAND, SUM (0) ON_ORDER, SUM (0) TRAN, SUM (0) OPEN, SUM (0) PICK, SUM (0) ALLO, SUM (0) COMM, SUM (0) PROD")
            SQL.AppendLine(", NULL UPC_CODE, NULL STYLE_COLOR_STATUS from ICTSTAT1")
            SQL.AppendLine(" where ICTSTAT1.STYLE_CODE = :PARM1")
            SQL.AppendLine("   and ICTSTAT1.OPS_YYYYPP = :PARM2")
            SQL.AppendLine(" group by ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE)")
            SQL.AppendLine(" union")
            SQL.AppendLine("(Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE")
            SQL.AppendLine(", SUM(0) BEG, SUM (0) SHP, SUM (0) RTN, SUM (0) REC, SUM (0) ADJ, SUM (0) XFR, SUM (0) PHY")
            SQL.AppendLine(", SUM(ICTSTAT2.WHSE_QTY_ON_HAND) ON_HAND")
            SQL.AppendLine(", SUM(ICTSTAT2.WHSE_QTY_ON_ORDER) ON_ORDER, SUM(ICTSTAT2.WHSE_QTY_TRAN) TRAN")
            SQL.AppendLine(", SUM(ICTSTAT2.WHSE_QTY_OPEN) OPEN, SUM(ICTSTAT2.WHSE_QTY_PICK) PICK")
            SQL.AppendLine(", SUM(ICTSTAT2.WHSE_QTY_ALLO) ALLO")
            SQL.AppendLine(", SUM(ICTSTAT2.WHSE_QTY_COMM) COMM, SUM(ICTSTAT2.WHSE_QTY_PROD) PROD")
            SQL.AppendLine(", NULL UPC_CODE, NULL STYLE_COLOR_STATUS from ICTSTAT2")
            SQL.AppendLine(" where ICTSTAT2.STYLE_CODE = :PARM1")
            SQL.AppendLine(" group by ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE)")
            SQL.AppendLine(" union")
            SQL.AppendLine("(Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE")
            SQL.AppendLine(", 0 BEG, 0 SHP, 0 RTN, 0 REC, 0 ADJ, 0 XFR, 0 PHY")
            SQL.AppendLine(", 0 ON_HAND, 0 ON_ORDER, 0 TRAN, 0 OPEN, 0 PICK, 0 ALLO, 0 COMM, 0 PROD, ICTSTYC1.UPC_CODE, ICTSTYC1.STYLE_COLOR_STATUS")
            SQL.AppendLine(" from ICTSTYC1")
            SQL.AppendLine(" where ICTSTYC1.STYLE_CODE = :PARM1)")
            SQL.AppendLine(") X")
            SQL.AppendLine(" where ICTCOLR1.COLOR_CODE (+) = X.COLOR_CODE")
            SQL.AppendLine("   and ICTSTYL1.STYLE_CODE (+) = X.STYLE_CODE")
            SQL.AppendLine(" group by X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ICTSTATA", "**", 0, False, "VV", 2)
            With .Tables("ICTSTATA").Columns
                .Add("OTS_INV", GetType(System.Int64), "ISNULL(ON_HAND,0) - ISNULL(PICK,0)")
                .Add("OTS_WIP", GetType(System.Int64), "ISNULL(OTS_INV,0) + ISNULL(TRAN,0) + ISNULL(ON_ORDER,0)")
                .Add("NET_POS", GetType(System.Int64), "ISNULL(OTS_WIP,0) - ISNULL(OPEN,0) - ISNULL(COMM,0) - ISNULL(PROD,0)")
            End With

            'For Each TABLE_NAME As String In New String() {"ICTSTATA"}
            'TABLE_NAME = "ICTSTATA"
            'With .Tables.Add(TABLE_NAME)
            '    For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE", "WHSE_CODE", "STYLE_DESC", "COLOR_DESC", "WHSE_DESC"}
            '        If TABLE_NAME = "ICTSTATA" And (COLUMN_NAME = "WHSE_CODE" Or COLUMN_NAME = "WHSE_DESC") Then
            '        ElseIf TABLE_NAME = "ICTSTATW" And (COLUMN_NAME = "STYLE_DESC" Or COLUMN_NAME = "COLOR_DESC") Then
            '        Else
            '            .Columns.Add(COLUMN_NAME)
            '        End If
            '    Next
            '    For Each COLUMN_NAME As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "PHY",
            '                                                    "ON_HAND", "ON_ORDER", "TRAN", "OPEN", "PICK", "ALLO", "COMM", "PROD"}
            '        .Columns.Add(COLUMN_NAME, GetType(System.Int64))
            '    Next
            '    If TABLE_NAME = "ICTSTATA" Then
            '        .Columns.Add("UPC_CODE")
            '        .Columns.Add("STYLE_COLOR_STATUS")
            '        .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            '        .Columns.Add("THEME_DESC")
            '    Else
            '        .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE"), .Columns("WHSE_CODE")}
            '    End If
            '    .Columns.Add("OTS_INV", GetType(System.Int64), "ISNULL(ON_HAND,0) - ISNULL(PICK,0)")
            '    .Columns.Add("OTS_WIP", GetType(System.Int64), "ISNULL(OTS_INV,0) + ISNULL(TRAN,0) + ISNULL(ON_ORDER,0)")
            '    .Columns.Add("NET_POS", GetType(System.Int64), "ISNULL(OTS_WIP,0) - ISNULL(OPEN,0) - ISNULL(COMM,0) - ISNULL(PROD,0)")
            'End With
        End With

        grdICTSTATA.DataSource = dst.Tables("ICTSTATA")

        'Create_Summary(grdICTSTYC1, "UPC_CODE", "Count")

        'spl.Panel1Collapsed = True

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            optStockNon.Visible = True
        Else
            optStockNon.Visible = False
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                Fill_Records("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text)
                If dst.Tables.Item("ICTSTYL1").Rows.Count = 1 Then
                    LoadStyleInformation()
                End If
            Case "Done"

                'Dim rowEDT846T1 As DataRow = LookUp("EDT846T1", EDI_DOC_SEQ_NO)
                'If rowEDT846T1 Is Nothing Then
                '    Exit Sub
                'End If


                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("SEASON_CODE").Text) Then
                '        Exit Sub
                '    End If
                'End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Private Sub LoadStyleInformation()
        Dim rowICTSTYL1 As DataRow = dst.Tables.Item("ICTSTYL1").Rows(0)
        Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME").ToString & String.Empty
        If IMAGE_NAME.Length > 0 Then
            picStyle.ImageLocation = IMAGE_PATH & IMAGE_NAME
        End If
        lblStyleDesc.Text = rowICTSTYL1.Item("STYLE_DESC").ToString & String.Empty

        If chkShowColorAvail.Checked Then
            Fill_Records("ICTSTATA", {rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty, ASCMAIN1.CYP})
        End If

    End Sub

    Sub Proceed(ByVal eItemKey As String)
        Select Case eItemKey
            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)
            Case "Done"
                Clear_Record()
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Visible = False
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        If tf = True Then
            btnOptions.Visible = False
            grpOptions.Visible = False
        Else
            btnOptions.Visible = True
        End If
        grdICTSTATA.Visible = tf
        grpStyleInfo.Visible = tf
        If chkShowColorAvail.Checked Then
            grdICTSTATA.Visible = tf
            grdICTSTATA.Width = grdICTSTATA.Parent.Width - 10
        Else
            grdICTSTATA.Visible = False
        End If
    End Sub

    Sub Clear_Record()
        For Each TABLE_NAME As String In New String() {"ICTSTYL1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        picStyle.ImageLocation = ""
        Absx1.txtFor("STYLE_CODE").Text = ""
        lblStyleDesc.Text = ""
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            'rowICTPLIN1.Item("STMT_TYPE") = HFs("STMT_TYPE")
            'rowICTPLIN1.Item("STMT_DESC") = HFs("STMT_DESC")
        Else

        End If

        If EntryMode = "N" Then
        Else

        End If

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
    ByVal ctl As Control,
    ByVal COLUMN_NAME As String,
    Optional ByRef sql_where As String = "",
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                Select Case optStockNon.Value
                    Case "S"
                        sql_where = "NVL(CUST_CODE,'NULL') = 'NULL'"
                    Case "N"
                        sql_where = "NVL(CUST_CODE,'NULL') <> 'NULL'"
                    Case Else
                        sql_where = ""
                End Select
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdICTSTYC1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"
            '    Using F As New TAC.ICFSTYCX
            '        F.STYLE_CODE = ""
            '        F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
            '        F.ShowDialog()
            '        If F.STYLE_CODE <> "" Then
            '            Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
            '        End If
            '    End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Select Case Absx1.GetABSColumnName(sender)
                Case "STYLE_CODE"
                    If Absx1.txtFor("STYLE_CODE").Text.Length > 0 Then
                        Click_Command("View")
                    End If
            End Select
        End If
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "LP_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Me.UltraGroupBox1.Select() ' to force txt_Leave event to fire, for formatting
        '            Load_ICTSTYC1()
        '        End If
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "LP_CODE"
        '        Load_ICTSTYC1()
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        'With Absx1.txtFor(COLUMN_NAME)
        '    Select Case COLUMN_NAME

        '        Case "LP_CODE"
        '            Load_ICTSTYC1()

        '    End Select

        'End With
    End Sub

    Private Sub btnOptions_Click(sender As Object, e As EventArgs) Handles btnOptions.Click
        If grpOptions.Visible = True Then
            grpOptions.Visible = False
        Else
            grpOptions.Visible = True
        End If
    End Sub

#End Region
End Class