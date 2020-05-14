Imports System.Drawing

Public Class SOFREASP

    Private sqlData As String = String.Empty
    Private startDate As Date
    Private endDate As Date
    Private CUST_CODE As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "select sotinvh1.inv_no, sotinvh1.inv_date, SOTORDR1.ORDR_NO, sotordr1.cust_code, sotordr1.CUST_NAME, sotordr1.CUST_STORE_NO, sotordr1.CUST_STORE_NAME," _
            & " SOTPICK2.SHORT_REASON_CODE, SOTPICK2.SHORT_REASON_COMMENT, sotordr2.ORDR_QTY, sotordr2.style_code, sotordr2.color_code, sotordr2.style_desc, sotpick1.ship_bol_no," _
            & " sotpick2.pick_no, sotpick2.PICK_QTY, sotpick2.PICK_QTY_CONF, sotpick2.PICK_QTY_CANC, sotpick2.PICK_QTY_BACK, (sotpick2.PICK_QTY - sotpick2.PICK_QTY_CONF) SHORTAGE" _
            & " from sotordr1, sotordr2, sotpick1, sotpick2, SOTINVH1" _
            & " where sotordr1.ordr_no = sotordr2.ordr_no" _
            & " and sotordr2.ordr_no = sotpick2.ordr_no" _
            & " and sotordr2.ordr_lno = sotpick2.ordr_lno" _
            & " and sotpick1.ordr_no = sotordr1.ordr_no" _
            & " and sotpick1.pick_no = sotpick2.pick_no " _
            & " and sotpick2.SHORT_REASON_CODE is not null" _
            & " and sotinvh1.ordr_no = sotordr1.ordr_no"
            sqlData = ASCMAIN1.sql

            Create_TDA(.Tables.Add, "SOTPICKX", ASCMAIN1.sql)
        End With

        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        ASCMAIN1.Add_Value_List(grdSOTPICKX, "SHORT_REASON_CODE")

        dteStartDate.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dteStartDate.MinDate = DateAdd(DateInterval.Year, -3, DateTime.Now)

        dteEndDate.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dteEndDate.MinDate = DateAdd(DateInterval.Year, -3, DateTime.Now)

        dteStartDate.DateTime = DateTime.Now
        dteEndDate.DateTime = DateTime.Now

        grdSOTPICKX.DisplayLayout.Bands(0).GroupHeadersVisible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"
                startDate = dteStartDate.DateTime.ToShortDateString
                endDate = dteEndDate.DateTime.ToShortDateString

                If DateDiff(DateInterval.Day, endDate, startDate) > 0 Then
                    EMsg &= vbCr & "End Date must be greater equal Start Date."
                End If

                txtCUST_CODE.Text = txtCUST_CODE.Text.Trim
                If txtCUST_CODE.TextLength > 0 Then
                    Validate_Code("CUST_CODE")
                End If

            Case "Done"

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
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

        Else
            Clear_Record()
        End If

        grdSOTPICKX.Visible = ScreenMode

    End Sub

    Sub Clear_Record()

        dst.Tables("SOTPICKX").Rows.Clear()

        CUST_CODE = String.Empty
        Absx1.txtFor("CUST_CODE").Clear()

        Clear_All_Filters(grdSOTPICKX)

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")

        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        ASCMAIN1.sql = sqlData

        Dim grdText As String = String.Empty

        If txtCUST_CODE.TextLength > 0 Then
            ASCMAIN1.sql &= " and SOTINVH1.CUST_CODE = '" & txtCUST_CODE.Text & "'"
            grdText = "Customer: " & txtCUST_CODE.Text & ", "
        End If

        ASCMAIN1.sql &= " and SOTINVH1.INV_DATE BETWEEN '" & startDate.ToString("dd-MMM-yyyy") & "' and '" & endDate.ToString("dd-MMM-yyyy") & "'"
        Fill_Records("SOTPICKX", String.Empty, True, ASCMAIN1.sql)

        grdText &= "Invoiced between " & startDate & " and " & endDate

        Sort_grdColumns(grdSOTPICKX, "CUST_CODE, CUST_STORE_NO, ORDR_NO", False)
        grdSOTPICKX.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        grdSOTPICKX.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Width = grdSOTPICKX.DisplayLayout.Bands(0).Columns("CUST_CODE").Width
        grdSOTPICKX.DisplayLayout.Bands(0).Columns("PICK_QTY").Width = grdSOTPICKX.DisplayLayout.Bands(0).Columns("ORDR_QTY").Width
        grdSOTPICKX.DisplayLayout.Bands(0).Columns("PICK_QTY_CONF").Width = grdSOTPICKX.DisplayLayout.Bands(0).Columns("PICK_QTY_CANC").Width

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()

        Try

        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTPICKX, "SSB", "Show Filter", "Show GroupBox", "Shipment Confirmation Inquiry")
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

        Select Case e.Tool.Key

            Case "Shipment Confirmation Inquiry"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                If SHIP_BOL_NO <> "" Then
                    Context_Launch("Select", SHIP_BOL_NO, e.Tool.Key, "SOFSHIPI", "F", "SO")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"
    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

#End Region

#Region "Form Controls"

#End Region


End Class