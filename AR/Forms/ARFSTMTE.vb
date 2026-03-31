Imports Infragistics.Win.UltraWinGrid

Public Class ARFSTMTE

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "ARTSTMTE", "*", 1)
        End With

        grdARTSTMTE.DataSource = dst.Tables("ARTSTMTE")
        ASCMAIN1.Add_Value_List(grdARTSTMTE, "CUST_STMT_IND",, {":", "B:Both", "E:Email"})
        Create_Summary(grdARTSTMTE, "CUST_CODE", "Count")

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)

        grdARTSTMTE.DisplayLayout.UseFixedHeaders = True
        grdARTSTMTE.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.Fixed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("OPS_YYYYPP")

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
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        grdARTSTMTE.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("ARTSTMTE").Rows.Clear()

        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)
        ASCMAIN1.Progress("Now Loading Customers")

        EnforceConstraints(False)
        Fill_Records("ARTSTMTE", txtOPS_YYYYPP.Text)
        EnforceConstraints(True)

        grdARTSTMTE.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
        Sort_grdColumns(grdARTSTMTE, "CUST_CODE")

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()
            MyBase.CommitTrans("Update Complete")
        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub


#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTSTMTE, "SSB", "Show Filter", "Show GroupBox", "Customer Master File")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name


            End Select
        End If

    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Master File"
                Dim CUST_CODE As String = grd.ActiveRow.Cells.Item("CUST_CODE").Text
                If CUST_CODE.Length > 0 Then
                    Context_Launch("View", Column_Values("CUST_CODE", CUST_CODE), e.Tool.Key, "ARTCUST1")
                End If
        End Select
    End Sub

#End Region

#Region "Controls and Procedures"

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As Control, COLUMN_NAME As String, ByRef Optional sql_where As String = "", ByRef Optional Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                sql_where = $"OPS_YYYYPP BETWEEN {ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)} AND {ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)}"
        End Select
    End Sub

    Private Sub grdARTSTMTE_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdARTSTMTE.InitializeRow

        If e.Row.Cells("INVALID_TO_EMAIL").Value & String.Empty = "1" Then
            e.Row.Cells("CUST_STMT_EMAIL").Appearance.BackColor = Drawing.Color.Pink
        End If

        If e.Row.Cells("INVALID_CC_EMAIL").Value & String.Empty = "1" Then
            e.Row.Cells("CUST_STMT_CC").Appearance.BackColor = Drawing.Color.Pink
        End If
    End Sub

#End Region

End Class