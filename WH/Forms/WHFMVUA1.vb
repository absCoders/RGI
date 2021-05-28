Public Class WHFMVUA1
    Dim tblWHTMOVEA As New DataTable
    Dim SECURITY_CODEsW As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SECURITY_CODEsW.Add("WH")
        SECURITY_CODEsW.Add("WS")

        Create_Lookup("ASTUSER1")
        dst.Tables.Add(tblWHTMOVEA)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Statistics"
                If Absx1.txtFor("USER_ID").Text <> "" Then
                    Validate_Code("USER_ID")
                End If

                If dteINIT_DATE.Value Is Nothing Then
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

            Case "Load Statistics"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load Statistics").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Errors from Log")
        Me.Cursor = Cursors.WaitCursor

        Dim sql As String = ""

        ASCMAIN1.sql = "Select WHTMOVE1.INIT_OPER, WHTMOVE1.INIT_DATE, to_char(WHTMOVE1.INIT_DATE, 'HH24') H24" & vbCrLf _
                & " ,case when WHTMOVE1.WHSE_TRAN_TYPE = 'W' then 'Receiving'" & vbCrLf _
                & "     when LOCATION_CODE_FROM = '00-RCV' then 'PutAway'" & vbCrLf _
                & "     when LOCATION_CODE_TO = '00-SHP' then 'Picking' " & vbCrLf _
                & "     else 'Move' end TRANSACTION," & vbCrLf _
                & " WHTMOVE2.STYLE_CODE, WHTMOVE2.COLOR_CODE, WHTMOVE2.WHSE_TRAN_QTY" & vbCrLf _
                & " ,case when nvl(ICTSTYL1.CARTON_PACK_QTY,0) > WHTMOVE2.WHSE_TRAN_QTY then 0" & vbCrLf _
                & " else trunc(WHTMOVE2.WHSE_TRAN_QTY / ICTSTYL1.CARTON_PACK_QTY) end CARTONS" & vbCrLf _
                & " ,ICTSTYL1.CARTON_PACK_QTY, WHTMOVE2.LOCATION_CODE_FROM, WHTMOVE2.LOCATION_CODE_TO" & vbCrLf _
                & " from WHTMOVE1, WHTMOVE2, ICTSTYL1" & vbCrLf _
                & "  where WHTMOVE1.WHSE_TRAN_NO =  WHTMOVE2.WHSE_TRAN_NO" & vbCrLf _
                & "  and ICTSTYL1.STYLE_CODE = WHTMOVE2.STYLE_CODE" & vbCrLf _
                & "  and WHTMOVE1.INIT_OPER = :PARM1" & vbCrLf _
                & "  and TRUNC(WHTMOVE1.INIT_DATE) = :PARM2"

        tblWHTMOVEA = ASCDATA1.GetDataTable(ASCMAIN1.sql, "WHTMOVEA", "VV", New Object() {Absx1.txtFor("USER_ID").Text, Format(DateValue(dteINIT_DATE.Value.ToString), "dd-MMM-yyyy")})
        grdWHTMOVEA.DataSource = tblWHTMOVEA

        'Sort_grdColumns(grdWHTMOVEA, "INIT_DATE".ToLower)

        grdWHTMOVEA.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdWHTMOVEA.DisplayLayout.Bands(0).SortedColumns.Add("H24", False, True)
        grdWHTMOVEA.Text = $"Transaction Activity for  {Absx1.txtFor("USER_ID").Text} on {Format(DateValue(dteINIT_DATE.Value.ToString), "dd-MMM-yyyy")}"
        Show_Filter(grdWHTMOVEA)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()


    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTMOVEA, "SS", "Show Filter", "Show GroupBox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        If COLUMN_NAME = "INIT_DATE" Then
            If ctl.Text <> "" Then
                Call Click_Command("Load Statistics")
            End If
        End If
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If Absx1.GetABSColumnName(sender) = "USER_ID" Then
            If e.KeyCode = Windows.Forms.Keys.Enter Then
                Call Click_Command("Load Statistics", e)
            End If
        End If
    End Sub
    Overrides Sub Prepare_for_View_Lookup_Special(
     ByVal ctl As Control,
     ByVal COLUMN_NAME As String,
     Optional ByRef sql_where As String = "",
     Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "USER_ID"
                sql_where = "USER_ID in (" _
                    & "(Select Distinct USER_ID from ASTUSER2 where SECURITY_CODE in ('" & Join(SECURITY_CODEsW.ToArray, "','") & "'))" _
                    & " minus " _
                    & "(Select Distinct USER_ID from ASTUSER2 where SECURITY_CODE Not in ('" & Join(SECURITY_CODEsW.ToArray, "','") & "'))" _
                    & ")"
        End Select
    End Sub
#End Region

#Region "grdWHTMOVEA"

#End Region

    Private Sub grdASTERROR_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTMOVEA.AfterRowActivate
        'txtSTACKTRACE.Text = grdWHTMOVEA.ActiveRow.Cells("STACKTRACE").Text
        'grpERR_TEXT.Text = grdWHTMOVEA.ActiveRow.Cells("ERR_TEXT").Text
    End Sub
End Class