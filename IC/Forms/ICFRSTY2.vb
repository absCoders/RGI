Imports System.Drawing

Public Class ICFRSTY2
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim rowICTRSTY1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
            & " from SOTINVH1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " and ORDR_INV_TYPE = 'I' and ORDR_INV_REG IS NULL"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTINVHW", "*")
        End With

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdICTRSTYX.DataSource = dst.Tables("SOTINVHX")

        Create_Summary(grdICTRSTYX, "RANGE_STYLE_CODE", "Count")
        'Set_Read_Only(grpInvoice, True)

        Show_Filter(grdICTRSTYX)

        'With grdICTRSTYX.DisplayLayout.Bands("SOTINVHX")
        '    .Columns("SO_ORDER_NO").Header.Fixed = True
        '    .Columns("ORDR_TOTAL_AMT").Header.Fixed = True
        '    .Columns("ORDR_INV_NO").Header.Fixed = True
        '    .Columns("ORDR_INV_DATE").Header.Fixed = True
        '    .Columns("ORDR_DIV_CODE").Header.Fixed = True
        'End With

        ASCMAIN1.Add_Value_List(grdICTRSTYX, "TRANSFER_TYPE")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Select Order"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                Load_SOTINVHX()

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1


                .Groups("Screen Control").Items("Approve").Visible = (MENU_ITEM_OBJECT = "SOFINVHA")
                .Groups("Screen Control").Items("Refresh").Visible = (MENU_ITEM_OBJECT <> "SOFINVHA")
                .Groups("Screen Control").Items("Update").Visible = (MENU_ITEM_OBJECT <> "SOFINVHA")
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Select Order").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Approve").Settings.Enabled = iScreenMode
                .Groups("Totals").Visible = ScreenMode
                .Groups("Show").Visible = Not ScreenMode And MENU_ITEM_OBJECT <> "SOFINVHA"
                .Groups("Batch Defaults").Visible = Not ScreenMode And MENU_ITEM_OBJECT <> "SOFINVHA"
            End With
        End If

        grpHeader.Visible = ScreenMode

        grdICTRSTYX.Visible = Not ScreenMode

        Set_Read_Only(grpHeader, True)
        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
            Show_Filter(grdICTRSTYX, True)
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SOTINVHX", "SOTORDRT" _
            , "SOTINVH1", "SOTINVH2", "SOTINVH3", "SOTINVH5", "SOTINVHB"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

        Load_SOTINVHX()
        Absx1.txtFor("SO_ORDER_NO").Focus()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Selecting Order for Invoicing")

        EnforceConstraints(False)


        EnforceConstraints(True)


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Overrides Function Remote_Control( _
 ByVal command As String, _
 Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Select Order"
                Absx1.txtFor("SO_ORDER_NO").Text = key
                Click_Command("Select Order")
        End Select

        Return return_key
    End Function

    Sub Update_Record()

        ' ANY CHANGES TO THIS METHOD SHOULD BE CHECKED INTO SOFINVH2.REVERSE_INVOICE

        BeginTrans()

        ASCMAIN1.Progress("", "")
        CommitTrans("Update Complete")

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SO_ORDER_NO"

            Case Else
        End Select

    End Sub


    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTORDR1"
            E.COLUMN_NAME = "SO_ORDER_NO"
            E.CODE_VALUE = Absx1.txtFor("SO_ORDER_NO").Text
            E.DESC_VALUE = "Sales Order"
            E.ATTACHMENT_NOTES = ""
        End If

        Return E
    End Function
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        If MENU_ITEM_OBJECT = "SOFINVHA" Then
            Load_Popup_Menu(grdICTRSTYX, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Inquiry", "Approve", "Approve Batch")
        Else
            Load_Popup_Menu(grdICTRSTYX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Inquiry")
        End If
        Load_Popup_Menu(grdSOTORDR3, "S", "Show Filter")
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
            Case "grdSOTINVHX"
                e.Tool.ToolbarsManager.Tools("Sales Order Inquiry").SharedProps.Visible = True
                If MENU_ITEM_OBJECT = "SOFINVHA" Then

                    If grd.ActiveRow.Cells("BATCH_NO").Text <> "" Then
                        e.Tool.ToolbarsManager.Tools("Approve Batch").SharedProps.Caption = "Approve Batch: " & grd.ActiveRow.Cells("BATCH_NO").Text
                    End If
                    e.Tool.ToolbarsManager.Tools("Approve Batch").SharedProps.Visible = (grd.ActiveRow.Cells("BATCH_NO").Text <> "")
                    e.Tool.ToolbarsManager.Tools("Approve").SharedProps.Caption = "Approve Order: " & grd.ActiveRow.Cells("SO_ORDER_NO").Text
                End If
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
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)

        Select Case COLUMN_NAME
            Case "SO_ORDER_NO"
                If Absx1.txtFor("SO_ORDER_NO").Text <> "" Then
                    Click_Command("Select Order")
                End If
        End Select
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)

        Select Case COLUMN_NAME

            Case "SO_ORDER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Select Order", e)
                End If

        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_CODE"
                If Not Me.IsLoading Then
                    ' IF THE FRT_TERMS TEXT BOX.TEXT= "" THEN .TEXT = "SOMETHING"
                End If

                If Absx1.txtFor("SHIP_CODE").Text <> "" Then
                    LookUp("SOTSVIA1", Absx1.txtFor("SHIP_CODE").Text)

                    If cdr IsNot Nothing Then
                        Absx1.txtFor("SHIP_VIA").Text = cdr.Item("SHIP_DESC")
                    End If
                End If

            Case "FRT_TERMS"
                If Absx1.txtFor("FRT_TERMS").Text = "TIS" Or Absx1.txtFor("FRT_TERMS").Text = "CPU" Then
                    Absx1.txtFor("SHIP_CODE").Text = Absx1.txtFor(COLUMN_NAME).Text
                End If

                If Absx1.txtFor("FRT_TERMS").Text <> "PPD" Then
                    Absx1.numFor("FRT_RATE").Value = 0
                    Set_Read_Only_for_ctl(Absx1.numFor("FRT_RATE"), True)
                Else
                    Set_Read_Only_for_ctl(Absx1.numFor("FRT_RATE"), False)
                End If

                If Not IsLoading Then
                    If use_CUST_PU_DATE And Absx1.txtFor("FRT_TERMS").Text <> "PPD" _
                    And Format(Absx1.dteFor("ORDR_DATE_SHIPPED").Value, "yyyyMMdd") _
                                         >= Format(Absx1.dteFor("CUST_PU_DATE").Value, "yyyyMMdd") Then
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("CUST_PU_DATE").Value
                    Else
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("ORDR_DATE_SHIPPED").Value
                    End If

                End If

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "ITEM_CODE"
            '    If Absx1.txtFor("ITEM_CODE").Tag = "Y" Then
            '        Absx1.txtFor("ITEM_CODE").Tag = ""
            '        'Dim X As CurrencyManager = Me.BindingContext(dst.Tables("SOTORDR2"))
            '        'X.EndCurrentEdit()

            '        Click_Command("Find Lots")
            '    End If
        End Select
    End Sub

    Public Overrides Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_Leave(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "SHIP_CODE"
                If Absx1.txtFor(COLUMN_NAME).Text = "TIS" Or Absx1.txtFor(COLUMN_NAME).Text = "CPU" Then
                    Absx1.txtFor("FRT_TERMS").Text = Absx1.txtFor(COLUMN_NAME).Text
                End If
            Case "FRT_TERMS"
                If Absx1.txtFor(COLUMN_NAME).Text = "TIS" Or Absx1.txtFor(COLUMN_NAME).Text = "CPU" Then
                    Absx1.txtFor("SHIP_CODE").Text = Absx1.txtFor(COLUMN_NAME).Text
                End If

        End Select

    End Sub

    Public Overrides Sub opt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "TRANSFER_TYPE"
                ' Toggle_Transfer_Fields()

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "FRT_RATE"
                If ScreenMode Then
                    Synch_TABLE_NAME("SOTORDR1")
                    Display_Totals()
                End If
            Case "SVC_CHG_RATE"
                If ScreenMode Then
                    Synch_TABLE_NAME("SOTORDR1")
                    Display_Totals()
                End If
        End Select

    End Sub

    Public Overrides Sub dte_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ORDR_DATE_SHIPPED"
                If Not IsLoading Then
                    If use_CUST_PU_DATE And Absx1.txtFor("FRT_TERMS").Text <> "PPD" _
                    And Format(Absx1.dteFor("ORDR_DATE_SHIPPED").Value, "yyyyMMdd") _
                                         >= Format(Absx1.dteFor("CUST_PU_DATE").Value, "yyyyMMdd") Then
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("CUST_PU_DATE").Value
                    Else
                        Absx1.dteFor("ORDR_INV_DATE").Value = Absx1.dteFor("ORDR_DATE_SHIPPED").Value
                    End If

                End If

                If Absx1.dteFor("ORDR_DATE_SHIPPED").Value = Absx1.dteFor("ORDR_INV_DATE").Value Then
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Empty
                Else
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Red
                End If

            Case "ORDR_INV_DATE"

                If Absx1.dteFor("ORDR_DATE_SHIPPED").Value = Absx1.dteFor("ORDR_INV_DATE").Value Then
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Empty
                Else
                    Absx1.dteFor("ORDR_INV_DATE").Appearance.ForeColor = Drawing.Color.Red
                End If

        End Select

    End Sub
#End Region

    Private Sub Load_SOTINVHX()
        If SELECTION_NO = 0 Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ASCMAIN1.sql = "Select * from "
        Fill_Records("ICTRSTYX", "", , ASCMAIN1.sql)

        Setup_grdSOTINVHX()
        Sort_grdColumns(grdICTRSTYX, "CUST_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_grdSOTINVHX()
       
    End Sub

    Private Sub grdSOTINVHX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTRSTYX.DoubleClickRow
        If optShow.Value = "P" Or optShow.Value = "R" Or optShow.Value = "A" Then
            Absx1.txtFor("SO_ORDER_NO").Text = e.Row.Cells("SO_ORDER_NO").Value
            Click_Command("Select Order")
        End If
    End Sub

    Sub Display_Totals()
        If SELECTION_NO = 0 Then Exit Sub


    End Sub

#Region "grdSOTORDR3"

    Private Sub grdSOTORDR3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR3.AfterCellUpdate
        If e.Cell.Column.Key = "SO_LOT_CASES" Then
            If e.Cell.Row.Cells("PACK_CODE").Text <> TAC.TACMAIN1.CATCH_PACK Then
                e.Cell.Row.Cells("SO_LOT_UNITS").Value = _
                Val(e.Cell.Row.Cells("SO_LOT_CASES").Value & "") * _
                Val(e.Cell.Row.Cells("PACK_FACTOR").Value & "")
            End If
        End If
    End Sub

    Private Sub grdSOTORDR3_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDR3.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTORDR3_BeforeCellActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdSOTORDR3.BeforeCellActivate
        If BATCH_NO <> "" Then
            e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit
        End If
        If e.Cell.Column.Key = "SO_LOT_UNITS" Then
            If e.Cell.Row.Cells("PACK_CODE").Text <> TAC.TACMAIN1.CATCH_PACK Then
                e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit
            Else
                e.Cell.Column.CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End If
    End Sub

    Private Sub grdSOTORDR3_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDR3.BeforeCellUpdate

        If e.Cell.Column.Key = "SO_LOT_UNITS" Then
            If e.Cell.Row.Cells("PACK_CODE").Text <> TAC.TACMAIN1.CATCH_PACK Then
                If grdSOTORDR3.ActiveCell IsNot Nothing AndAlso grdSOTORDR3.ActiveCell.Column.Key = "SO_LOT_CASES" Then
                Else
                    e.Cancel = True
                End If
            End If
        End If

    End Sub

    Private Sub grdSOTORDR3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR3.InitializeRow

        If e.Row.Cells("PACK_CODE").Text = TAC.TACMAIN1.CATCH_PACK Then
            e.Row.Cells("PACK_DESC").Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("PACK_DESC").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("SO_LOT_UNITS").Appearance.BackColor = Drawing.Color.Empty
        End If
        If Val(e.Row.Cells("SO_LOT_CASES").Text) <> Val(e.Row.Cells("SO_LOT_CASES_ORIG").Text) Then
            If e.Row.Cells("ACK").Text = "Checked" Then
                e.Row.Cells("ACK").Appearance.BackColor = Drawing.Color.Green
            Else
                e.Row.Cells("ACK").Appearance.BackColor = Drawing.Color.Red
            End If
        End If

    End Sub

    Private Sub grdSOTORDR3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdSOTORDR3.KeyPress
        If grdSOTORDR3.ActiveCell Is Nothing Then
            Exit Sub
        End If
    End Sub

#End Region

End Class