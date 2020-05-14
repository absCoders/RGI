Imports System.Drawing

Public Class SOFSHPB1
    Dim CUST_CODE As String
    Dim CUST_BILL_TO_CUST As String
    Dim SO_ORDER_NO As String
    Dim SO_ORDER_NO_init As String
    Dim BATCH_NO As String
    Dim SHIP_CODE_ORIG As String
    Dim CURR_CODE As String
    Dim CURR_EXCH_RATE As Decimal
    Dim ORDR_TYPE_CODE As String
    Dim use_CUST_PU_DATE As Boolean

    Dim rowARTCUST1_SOLDTO As DataRow
    Dim rowSOTORDR1 As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")

        With dst
            ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
            & " from SOTINVH1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " and ORDR_INV_TYPE = 'I' and ORDR_INV_REG IS NULL"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTINVHW", "*")
        End With

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
      
        Create_Summary(grdSOTINVHX, "SO_ORDER_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_TOTAL_AMT")

        Set_Read_Only(grpInvoice, True)

        Show_Filter(grdSOTINVHX)

        With grdSOTINVHX.DisplayLayout.Bands("SOTINVHX")
            .Columns("SO_ORDER_NO").Header.Fixed = True
            .Columns("ORDR_TOTAL_AMT").Header.Fixed = True
            .Columns("ORDR_INV_NO").Header.Fixed = True
            .Columns("ORDR_INV_DATE").Header.Fixed = True
            .Columns("ORDR_DIV_CODE").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdSOTINVHX, "TRANSFER_TYPE")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select Order"

                If dteDATE_SHIPPED.Value & "" = "" Then
                    EMsg &= vbCr & "No value specified for Invoice Date"
                End If


            Case "Update"

               
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
        Absx1.optFor("TRANSFER_TYPE").Visible = ScreenMode AndAlso (ORDR_TYPE_CODE = "T")

        lblORDR_TYPE_CODE.Visible = ScreenMode And ORDR_TYPE_CODE <> "T"
        Absx1.cbeFor("ORDR_TYPE_CODE").Visible = ScreenMode And ORDR_TYPE_CODE <> "T"

   
        grdSOTINVHX.Visible = Not ScreenMode

        Set_Read_Only(grpHeader, True)
        Set_Read_Only(UltraGroupBox1, ScreenMode)
     
        If ScreenMode Then
        Else
            Clear_Record()
            Show_Filter(grdSOTINVHX, True)
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
            Load_Popup_Menu(grdSOTINVHX, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Inquiry", "Approve", "Approve Batch")
        Else
            Load_Popup_Menu(grdSOTINVHX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry", "Customer Inquiry")
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
            Case "Sales Order Inquiry"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Context_Launch("Load", SO_ORDER_NO, e.Tool.Key, "SOFORDRI", "F", "SO")

            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")
            Case "Approve"
                Update_Approval_Fields(grd.ActiveRow.Cells("SO_ORDER_NO").Text)
            Case "Approve Batch"
                Update_Approval_Batch(grd.ActiveRow.Cells("BATCH_NO").Text)
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


        Select Case optShow.Value
            Case "N"
                ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
                & " from SOTINVH1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
                & " and ORDR_INV_TYPE = 'I' and ORDR_INV_REG IS NULL"

            Case "P"
                ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1 where ORDR_REL = '1' "

            Case "A"
                ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1 where ORDR_REL = '1' " _
                & " and BILLING_APPR_DATE is not Null"


            Case "D"
                ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1" & vbCrLf _
                & " where ORDR_INV_DATE = '" & Format(dteShow.Value, "dd-MMM-yyyy") & "'"

            Case "J"
                ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1" & vbCrLf _
                & " where ORDR_INV_REG_DATE = '" & Format(dteShow.Value, "dd-MMM-yyyy") & "'"
            Case "R"
                ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1 where ORDR_REL = '1' " _
                & " and BILLING_APPR_DATE is Null"
        End Select

        If chkMyInvoicesOnly.Checked Then
            ASCMAIN1.sql &= " and LAST_OPER = '" & ASCMAIN1.USER_ID & "'"

        End If

        Fill_Records("SOTINVHX", "", , ASCMAIN1.sql)

        Setup_grdSOTINVHX()
        Sort_grdColumns(grdSOTINVHX, "SO_ORDER_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_grdSOTINVHX()
        Dim dvw As DataView = DirectCast(grdSOTINVHX.DataSource, DataTable).DefaultView
        dvw.RowFilter = ""
        grdSOTINVHX.Text = Replace(optShow.Text, "...", Format(dteShow.Value, "MM/dd/yyyy"))
    End Sub

    Private Sub grdSOTINVHX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHX.DoubleClickRow
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


#Region "VB6"

    Dim CUST_CODE As String         ' Sold-To Customer Code
    Dim ORDR_GROUP_NO As String     ' ORDR_GROUP_NO for Order currently in process
    Dim ORDR_CUST_PO As String      ' Customer's PO No

    Dim tp As Boolean               'Using a Third Party
    Dim CUST_NAME_TP As String      'Third Party Cust Name
    Dim CUST_RECIP_TP As String     'Third Party Contact
    Dim CUST_ADDR1_TP As String     'Third Party Addr1
    Dim CUST_ADDR2_TP As String     'Third Party Addr2
    Dim CUST_CITY_TP As String      'Third Party City
    Dim CUST_STATE_TP As String     'Third Party State
    Dim CUST_ZIP_TP As String       'Third Party Zip
    Dim CUST_ACCT_NO_TP As String   'Third Party Account No

    Dim STYLE_CODE As String        ' current STYLE_CODE
    Dim COLOR_CODE As String        ' current COLOR_CODE
    Dim SHIPPER_CODE_CAPTION As String ' Used in grid caption
    Dim SHIP_ORDR_TYPE_CAPTION As String ' Used in grid caption

    Dim gNo As String               'Group Number if using Carton Pack Configuration
    Dim SHIPPER_CODE As String      'Shipper Code
    Dim SHIP_ORDR_TYPE As String    'Ship Order Type; B for Batched, K for Keyed Import
    Public weightLbs As Double      'Package Weight
    Dim sqlSOWSHPW1 As String       'Work File Sql

    Dim SQL_PRINT_SEQ As String     'Label Printing Sequesce
    Dim sqlPrint() As String        'Array used to hold descriptions and field names for printing order

    Dim dynARTCUST1 As OraDynaset   'Customer Master Dynaset
    Dim dynSOTORDR1 As OraDynaset   'Order Header Dynaset
    Dim dynSOTPICK2 As OraDynaset
    Dim Path As String

    Function cmdExecute_Check(Index As Integer) As String
        Dim i As Integer
        Dim z As String
        Dim EMsg As String

        cmdExecute_Check = "0"

        Select Case Index

            Case 2
                'Load Group

                OraD.Parameters("CUST_CODE").Value = txtCode(0).Text
                dynARTCUST1.Refresh()
                If dynARTCUST1.EOF Then
                    EMsg = EMsg & vbCr & txtCode(0).Text & " Is NOT a Valid Customer."
                Else
                    If dynARTCUST1.Fields("CUST_SHIPPER_ID").Value & "" = "" And SHIPPER_CODE <> "DHL" Then
                        EMsg = EMsg & vbCr & txtCode(0).Text & " Is a Valid Customer, but no Shipper ID has been specified."
                        EMsg = EMsg & vbCr & " Please use the Customer Master File to set a Shipper ID for " & txtCode(0) & "."
                    End If
                End If

                If txtCode(2).Text <> "" Then
                    OraD.Parameters("CODE").Value = Trim(txtCode(2).Text & "")
                    dynSOTORDR1.Refresh()
                    If dynSOTORDR1.EOF Then
                        EMsg = EMsg & vbCr & txtCode(2).Text & " Is not a Valid Customer PO."
                    End If
                End If

            Case 3
                'Update
                Dim paramCaption As String
                For i = 0 To 2
                    If Text1(i).Text = "" Then
                        paramCaption = Mid$(lblShipmentParam(i).Caption, 1, Len(lblShipmentParam(i).Caption) - 1)
                        If paramCaption <> "Package Type" Then
                            EMsg = EMsg & vbCr & "Please specifiy a " & paramCaption & "."
                        End If
                    End If
                Next i
                If SHIP_ORDR_TYPE = "B" And SQL_PRINT_SEQ = "" Then
                    EMsg = EMsg & vbCr & "Please Select a Print Sequence for Batched Labels"
                End If
                'CHECK FOR RECORDS WITH PRINT_STATUS = 'A'
            Case 4
                'Cancel
            Case 6
                'Import Log
            Case Else
                Stop
        End Select

        If EMsg <> "" Then
            MsgBox(Mid$(EMsg, 2), 16 + 0, "Cannot Proceed")
            Exit Function
        End If
        cmdExecute_Check = "1"

    End Function

    Sub cmdExecute(Index As Integer)
        Select Case Index
            Case 2
                EMode = "E"
                Call Setup_Record()
            Case 3
                Call Update_Record()
            Case 4
                Call Clean_Up()
                Call Modes(False)
                Call Set_Control()
            Case 6
                'Call Select_File
        End Select
    End Sub
    Sub Clean_Up()

        Call Delete_WK("SOWCONF4")
        Call Delete_WK("SOWSHPW1")

    End Sub


    Sub Init_Form_Load()

        'ORDER PREP WORK TABLE (USED TO CREATE BATCH AND KEYED IMPORT RECORDS FOR BOTH SHIPPERS)
        SQL = "  SELECT SOTORDR1.ORDR_GROUP_NO, SOTPICK1.PICK_NO, SOTCART1.CART_NO"
        SQL = SQL & "  , NULL TRACK_NO, ARTCUST2.CUST_NAME CUST_NAME"
        SQL = SQL & "  , 'RECEIVING' CUST_CONTACT, ARTCUST2.CUST_ADDR1 CUST_ADDR1, ARTCUST2.CUST_ADDR2 CUST_ADDR2"
        SQL = SQL & "  , ARTCUST2.CUST_CITY CUST_CITY, ARTCUST2.CUST_STATE CUST_STATE, TRIM(SUBSTR(ARTCUST2.CUST_ZIP_CODE,1,5)) CUST_ZIP_CODE"
        SQL = SQL & "  , ARTCUST2.CUST_PHONE, 'US' CUST_COUNTRY, NULL CUST_NAME_TP, NULL CUST_CONTACT_TP, NULL CUST_ADDR1_TP, NULL CUST_ADDR2_TP"
        SQL = SQL & "  , NULL CUST_CITY_TP, NULL CUST_STATE_TP, NULL CUST_ZIP_CODE_TP, NULL CUST_PHONE_TP, NULL CUST_COUNTRY_TP"
        SQL = SQL & "  , NULL SHIPPER_ACCT_CODE_TP, SUM(SOTCART2.QTY_PACKED * ICTSTYL1.STYLE_WEIGHT) CALC_CARTON_WEIGHT, SUM(SOTCART2.QTY_PACKED * ICTSTYL1.STYLE_WEIGHT) PACKAGE_WEIGHT, 'XXXXXXXXXXX' PACKAGE_TYPE"
        SQL = SQL & "  , 'XXXXXXXXXXXXXXXXXXX' PAYMENT_TYPE, 'XXXXXXXXXXXXXXXXXX' SERVICE_TYPE, SOTORDR1.ORDR_NO, NULL BATCH_GROUP_NO, NULL SHIPPER_ACCT_CODE"
        SQL = SQL & "  , ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_TYPE, ARTCUST2.CUST_ADDR_CODE, NULL SHIP_DATE, 'XXX' SHIPPER_CODE, 'X' SHIP_ORDR_TYPE"
        SQL = SQL & "  , 0 SORT_BY, 'X' PROCESS_IND, 0 LBL_GROUP_NO, ARTCUST2.CUST_EMAIL, DECODE(NVL(ARTCUST2.CONSIGNEE_BILLED,'N'),'N','N','Y') CONSIGNEE_BILLED , 'X' PRINT_STATUS, 'XXXX' PRINT_MACHINE"
        SQL = SQL & "  , TRANSIT_BUS_DAYS, 0 as PICK_TOT"
        SQL = SQL & "   FROM ARTCUST1, ARTCUST2, SOTORDR1, SOTPICK1, SOTCART1, SOTCART2, ICTSTYL1, SOTSVIA2,"
        SQL = SQL & "  (SELECT ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_TYPE, ARTCUST2.CUST_ADDR_CODE"
        SQL = SQL & "  , NVL(ARTCUST2.CUST_ADDR_CODE_DOMESTIC, ARTCUST2.CUST_ADDR_CODE) CUST_ADDR_CODE_OVER"
        SQL = SQL & "  FROM ARTCUST1, ARTCUST2"
        SQL = SQL & "  Where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE"
        SQL = SQL & "  AND ARTCUST1.CUST_SHIPPER_ID IS NOT NULL) ARTCUSTX"
        SQL = SQL & "  WHERE ROWNUM < 1"
        SQL = SQL & "  Group By"
        SQL = SQL & "  SOTORDR1.ORDR_GROUP_NO, SOTPICK1.PICK_NO, SOTCART1.CART_NO"
        SQL = SQL & "  , ARTCUST2.CUST_NAME, ARTCUST2.CUST_ADDR1, ARTCUST2.CUST_ADDR2"
        SQL = SQL & "  , ARTCUST2.CUST_CITY, ARTCUST2.CUST_STATE, TRIM(SUBSTR(ARTCUST2.CUST_ZIP_CODE,1,5))"
        SQL = SQL & "  , ARTCUST2.CUST_PHONE, SOTORDR1.ORDR_NO, ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_TYPE, ARTCUST2.CUST_ADDR_CODE"
        SQL = SQL & "  , ARTCUST2.CUST_EMAIL, DECODE(NVL(ARTCUST2.CONSIGNEE_BILLED,'N'),'N','N','Y'), TRANSIT_BUS_DAYS"
        Call Ora_to_Acc(Nothing, "SOWSHPW1", 0, "", SQL)

        SQL = "Select * from ARTCUST1 where CUST_CODE = :CUST_CODE"
        SQL = SQL & " and CUST_STATUS = 'A'"
        dynARTCUST1 = OraD.CreateDynaset(SQL, 8&)

        SQL = "SELECT * FROM SOTORDR1 WHERE ORDR_CUST_PO = :CODE"
        SQL = SQL & " AND CUST_CODE = :CUST_CODE"
        SQL = SQL & " AND ORDR_STATUS IN ('P','F')"
        dynSOTORDR1 = OraD.CreateDynaset(SQL, 8&)

        SQL = "Select Sum(PICK_QTY) as PICK_TOT from SOTPICK2"
        SQL = SQL & " Where PICK_NO = :PICK_NO"
        dynSOTPICK2 = OraD.CreateDynaset(SQL, 8&)

        SQL = "SELECT CONFIG_NO, COUNT(*) STORE_COUNT, 0.01 WEIGHT, '0' PRINT_GROUP "
        SQL = SQL & ", BATCHED ALREADY_PRINTED"
        SQL = SQL & " From "
        SQL = SQL & " SOTCONF2 , SOTORDR1 "
        SQL = SQL & " Where ROWNUM < 1 "
        SQL = SQL & " GROUP BY CONFIG_NO, BATCHED "
        Call Ora_to_Acc(Nothing, "SOWCONF4", 1, "", SQL)

        SQL = "SELECT * "
        SQL = SQL & " From "
        SQL = SQL & " SOTLBLG1 "
        SQL = SQL & " Where ROWNUM < 1 "
        Call Ora_to_Acc(Nothing, "SOWLBLG1", 1, "", SQL)

        SQL = "SELECT SOTLBLG2.*, 0 LBL_COUNT, 'XXXXXX' STORE_BEG, 'XXXXXX' STORE_END "
        SQL = SQL & " From "
        SQL = SQL & " SOTLBLG2 "
        SQL = SQL & " Where ROWNUM < 1 "
        Call Ora_to_Acc(Nothing, "SOWLBLG2", 2, "", SQL)

        SQL = " SELECT SHIP_PARM_CODE SERVICE_CODE_FDX, SHIP_PARM_CODE SERVICE_CODE_UPS" & vbCr
        SQL = SQL & ", SHIP_PARM_DESC SERVICE_DESC_FDX, SHIP_PARM_DESC SERVICE_DESC_UPS" & vbCr
        SQL = SQL & ", SHIP_PARM_CODE BILLING_CODE_FDX, SHIP_PARM_CODE BILLING_CODE_UPS" & vbCr
        SQL = SQL & ", SHIP_PARM_DESC BILLING_DESC_FDX, SHIP_PARM_DESC BILLING_DESC_UPS" & vbCr
        SQL = SQL & ", SHIP_PARM_CODE PACKAGE_CODE_FDX, SHIP_PARM_CODE PACKAGE_CODE_UPS" & vbCr
        SQL = SQL & ", SHIP_PARM_DESC PACKAGE_DESC_FDX, SHIP_PARM_DESC PACKAGE_DESC_UPS" & vbCr
        SQL = SQL & " FROM SOTSHPP1 WHERE ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "SOWSHPP1", 0, "", SQL)

        SQL = " SELECT * FROM SOTTPBL1 WHERE ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "SOWTPBL1", 2, "", SQL)

        Dim k As Integer
        Dim i As Integer

        Call Load_Data_Controls(Me)
    End Sub

    Sub Modes(t As Integer)
        Dim i As Integer
        mc(0, 0, mcif) = t
        If t = False Then
            EMode = ""
            txtCode(0).Text = ""
            txtCode(2).Text = ""
            cmbRecord(1).Text = ""
            cmdImport.Visible = False
            Path = ""
            For i = 0 To 2
                Text1(i).Text = ""
                Text2(i).Text = ""
            Next i
            grdSort.removeAll()

        End If

        Call Check_Shipments_Log()

        Me.KeyPreview = Not t
        cmdCodeLookup(0).Enabled = Not t
        frmHeader.Enabled = Not t
        frmD.Visible = t

        If t = False Then
            Call Set_mc(mcif, "SX")
        Else
            Call Set_mc(mcif, "UC")
        End If

    End Sub
 
    Sub Setup_WK(t As String)

        Dim i As Integer
        Dim CP As Integer
        Dim groupNo As Integer
        Dim dynWK As Recordset

        Dim OPENPICK As String
        Dim batTot As String
        Dim z As String
        Dim Ranges As String
        Dim ORDR_GROUP_NO As String
        Dim CART_WEIGHT As Double
        Dim sb() As String

        ORDR_GROUP_NO = cmbRecord(1).Text
        CART_WEIGHT = fpCarton.Value
        Text1(2).Visible = True
        Text2(2).Visible = True
        lblShipmentParam(2).Visible = True

        If optSHPR(0) Then
            SHIPPER_CODE = "FDX"
            SHIPPER_CODE_CAPTION = "FedEx"
        ElseIf optSHPR(1) Then
            SHIPPER_CODE = "UPS"
            SHIPPER_CODE_CAPTION = "UPS"
        Else
            SHIPPER_CODE = "DHL"
            SHIPPER_CODE_CAPTION = "DHL"
            Text1(2).Visible = False
            Text2(2).Visible = False
            lblShipmentParam(2).Visible = False
        End If
        If optType(0) Then
            SHIP_ORDR_TYPE = "B"
            SHIP_ORDR_TYPE_CAPTION = "Batch Import"
        Else
            SHIP_ORDR_TYPE = "K"
            SHIP_ORDR_TYPE_CAPTION = "Keyed Import"
        End If

        Call Prompt("Building Order Details", "This may take a moment")

        'SQL = sqlSOWSHPW1


        Sql = "  SELECT SOTORDR1.ORDR_GROUP_NO, SOTPICK1.PICK_NO, SOTCART1.CART_NO"
        Sql = Sql & "  , NULL TRACK_NO, ARTCUST2.CUST_NAME CUST_NAME"
        Sql = Sql & "  , 'RECEIVING' CUST_CONTACT, ARTCUST2.CUST_ADDR1 CUST_ADDR1, ARTCUST2.CUST_ADDR2 CUST_ADDR2"
        Sql = Sql & "  , ARTCUST2.CUST_CITY CUST_CITY, ARTCUST2.CUST_STATE CUST_STATE, TRIM(SUBSTR(ARTCUST2.CUST_ZIP_CODE,1,5)) CUST_ZIP_CODE"
        Sql = Sql & "  , ARTCUST2.CUST_PHONE, 'US' CUST_COUNTRY, NULL CUST_NAME_TP, NULL CUST_CONTACT_TP, NULL CUST_ADDR1_TP, NULL CUST_ADDR2_TP"
        Sql = Sql & "  , NULL CUST_CITY_TP, NULL CUST_STATE_TP, NULL CUST_ZIP_CODE_TP, NULL CUST_PHONE_TP, NULL CUST_COUNTRY_TP"
        Sql = Sql & "  , NULL SHIPPER_ACCT_CODE_TP, SUM(SOTCART2.QTY_PACKED * ICTSTYL1.STYLE_WEIGHT) CALC_CARTON_WEIGHT, SUM(SOTCART2.QTY_PACKED * ICTSTYL1.STYLE_WEIGHT) PACKAGE_WEIGHT, 'XXXXXXXXXXX' PACKAGE_TYPE"
        Sql = Sql & "  , 'XXXXXXXXXXXXXXXXXXX' PAYMENT_TYPE, 'XXXXXXXXXXXXXXXXXX' SERVICE_TYPE, SOTORDR1.ORDR_NO, NULL BATCH_GROUP_NO, NULL SHIPPER_ACCT_CODE"
        Sql = Sql & "  , ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_TYPE, ARTCUST2.CUST_ADDR_CODE, NULL SHIP_DATE, 'XXX' SHIPPER_CODE, 'X' SHIP_ORDR_TYPE"
        Sql = Sql & "  , 0 SORT_BY, 'X' PROCESS_IND, 0 LBL_GROUP_NO, ARTCUST2.CUST_EMAIL, DECODE(NVL(ARTCUST2.CONSIGNEE_BILLED,'N'),'N','N','Y') CONSIGNEE_BILLED , 'X' PRINT_STATUS, 'XXXX' PRINT_MACHINE "
        Sql = Sql & "  , TRANSIT_BUS_DAYS, 0 as PICK_TOT"
        Sql = Sql & "   FROM ARTCUST1, ARTCUST2, SOTORDR1, SOTPICK1, SOTCART1, SOTCART2, ICTSTYL1, SOTSVIA2,"
        Sql = Sql & "  (SELECT ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_TYPE, ARTCUST2.CUST_ADDR_CODE"
        Sql = Sql & "  , NVL(ARTCUST2.CUST_ADDR_CODE_DOMESTIC, ARTCUST2.CUST_ADDR_CODE) CUST_ADDR_CODE_OVER"
        Sql = Sql & "  FROM ARTCUST1, ARTCUST2"
        Sql = Sql & "  Where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE"
        Sql = Sql & "  AND ARTCUST1.CUST_SHIPPER_ID IS NOT NULL) ARTCUSTX"
        Sql = Sql & "   WHERE  " & vbCrLf
        Sql = Sql & "    ARTCUST1.CUST_CODE  = ARTCUST2.CUST_CODE  " & vbCrLf
        Sql = Sql & "    AND ARTCUST2.CUST_CODE  = SOTORDR1.CUST_CODE  " & vbCrLf
        Sql = Sql & "    AND ARTCUST2.CUST_ADDR_TYPE = SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf
        Sql = Sql & "    AND ARTCUST2.CUST_ADDR_CODE = DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC', SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO)" & vbCrLf
        Sql = Sql & "    AND ARTCUST2.CUST_CODE = ARTCUSTX.CUST_CODE(+)"
        Sql = Sql & "    AND ARTCUST2.CUST_ADDR_TYPE = ARTCUSTX.CUST_ADDR_TYPE(+)"
        Sql = Sql & "    AND ARTCUST2.CUST_ADDR_CODE = ARTCUSTX.CUST_ADDR_CODE_OVER(+)"
        Sql = Sql & "    AND SOTPICK1.ORDR_NO  = SOTORDR1.ORDR_NO  " & vbCrLf
        Sql = Sql & "    AND SOTPICK1.PICK_NO = SOTCART1.PICK_NO  " & vbCrLf
        Sql = Sql & "    AND SOTCART1.CART_NO = SOTCART2.CART_NO" & vbCrLf
        Sql = Sql & "    AND SOTCART2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf
        Sql = Sql & "    AND ARTCUST2.CUST_STATE  = SOTSVIA2.STATE_CODE(+)"
        Sql = Sql & "    AND PICK_STATUS IN ('P','F')  " & vbCrLf
        If optBATCH_TYPE(0).Value = True Then
            Sql = Sql & "    AND SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'  " & vbCrLf
        Else
            Sql = Sql & "    AND SOTPICK1.PICK_BATCH_NO = '" & ORDR_GROUP_NO & "'  " & vbCrLf
        End If
        Sql = Sql & "   GROUP BY SOTORDR1.ORDR_GROUP_NO, TRANSIT_BUS_DAYS, SOTPICK1.PICK_NO,  SOTCART1.CART_NO, ARTCUST2.CUST_NAME  " & vbCrLf
        Sql = Sql & "    , ARTCUST2.CUST_CONTACT,  ARTCUST2.CUST_ADDR1, ARTCUST2.CUST_ADDR2,  ARTCUST2.CUST_CITY, ARTCUST2.CUST_STATE  " & vbCrLf
        Sql = Sql & "    , ARTCUST2.CUST_ZIP_CODE, ARTCUST2.CUST_PHONE,  ARTCUST2.CUST_COUNTRY, SOTCART1.CART_TOTAL_WGT_CALC, SOTORDR1.ORDR_NO, SOTPICK1.PICK_TOTAL_WGT  " & vbCrLf
        Sql = Sql & "    , ARTCUST2.CUST_CODE, ARTCUST2.CUST_ADDR_TYPE, ARTCUST2.CUST_ADDR_CODE"
        Sql = Sql & " , ARTCUST2.CUST_EMAIL, DECODE(NVL(ARTCUST2.CONSIGNEE_BILLED,'N'),'N','N','Y')"
        Call Ora_to_Acc(Nothing, "SOWSHPW1", 0, "", Sql)

        Sql = "Update SOWSHPW1 set PACKAGE_TYPE = '', SERVICE_TYPE = ''"
        AccD.Execute(Sql)

        If chkByLine.Value = "1" Then
            Sql = "Update SOWSHPW1 set PAYMENT_TYPE = '' WHERE PAYMENT_TYPE = 'XXXXXXXXXXXXXXXXXXX'"
            AccD.Execute(Sql)
        Else
            Sql = "Update SOWSHPW1 set PACKAGE_TYPE = ''"
            AccD.Execute(Sql)
        End If

        Call Prompt("Building Pick Totals", "")



        Dim dynSOWSHPW1 As Recordset
        Sql = "Select * from SOWSHPW1"
        dynSOWSHPW1 = AccD.OpenRecordset(Sql, dbOpenDynaset)

        Do While Not dynSOWSHPW1.EOF
            ' SQL = "Select Sum(PICK_QTY) as PICK_TOT from SOTPICK2"
            ' SQL = SQL & " Where PICK_NO = '" & dynSOWSHPW1.Fields("PICK_NO").Value & "'"
            ' Set dynSOTPICK2 = OraD.CreateDynaset(SQL, 8&)

            OraD.Parameters("PICK_NO").Value = dynSOWSHPW1.Fields("PICK_NO").Value
            dynSOTPICK2.Refresh()


            If Not dynSOTPICK2.EOF Then
                dynSOWSHPW1.Edit()
                dynSOWSHPW1.Fields("PICK_TOT").Value = dynSOTPICK2.Fields("PICK_TOT").Value
                dynSOWSHPW1.Update()
            End If
            dynSOTPICK2.Close()
            dynSOWSHPW1.MoveNext()
        Loop
        dynSOWSHPW1.Close()


        Call Prompt("Building Order Details", "Finished!")


        If chkSSOR.Value = 1 Then
            'Store Specific Order, we do know the weights, there will only be one group
            Sql = "Update SOWSHPW1 set BATCH_GROUP_NO = '000001'"
            AccD.Execute(Sql)
        Else
            'Non Store Specific Order, we do not know the weights, there may be multiple groups of configurations
            'configuration here should match that of the store order configuration report
            If SHIP_ORDR_TYPE = "B" Then
                Call Prompt("Retrieving Store Order Configuration", "")
                Sql = "SELECT CONFIG_NO, COUNT(*) STORE_COUNT, 0.01 WEIGHT, 0 PRINT_GROUP" & vbCrLf
                Sql = Sql & ", BATCHED ALREADY_PRINTED "
                Sql = Sql & " From" & vbCrLf
                Sql = Sql & " SOTCONF2 , SOTORDR1" & vbCrLf
                Sql = Sql & " Where" & vbCrLf
                Sql = Sql & " SOTCONF2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf
                Sql = Sql & " AND SOTCONF2.CUST_STORE_NO = DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC', SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO)" & vbCrLf
                Sql = Sql & " AND SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf
                Sql = Sql & " GROUP BY CONFIG_NO, BATCHED"
                Call Ora_to_Acc(Nothing, "SOWCONF4", 1, "", Sql)

                Sql = "UPDATE SOWCONF4 SET WEIGHT = 0"
                AccD.Execute(Sql)
            Else
                Sql = "Update SOWSHPW1 set BATCH_GROUP_NO = '000001'"
                AccD.Execute(Sql)
            End If
        End If

        If SHIPPER_CODE = "UPS" And chkByLine.Value = "1" Then
            Sql = "UPDATE SOWSHPW1 SET PAYMENT_TYPE = 'Consignee Billed' WHERE CONSIGNEE_BILLED = 'Y'"
            AccD.Execute(Sql)
        End If

        Sql = " SELECT * FROM SOTTPBL1 WHERE CUST_CODE = '" & CUST_CODE & "'"
        Call Ora_to_Acc(Nothing, "SOWTPBL1", 2, "", Sql)

        'Add default Carton Weight
        Sql = "Update SOWSHPW1 set SHIPPER_CODE = '" & SHIPPER_CODE & "', SHIP_ORDR_TYPE = '" & SHIP_ORDR_TYPE & "'"
        Sql = Sql & ", PROCESS_IND = 'O', PACKAGE_WEIGHT = PACKAGE_WEIGHT + " & CART_WEIGHT
        AccD.Execute(Sql)

        Call Prompt("Building Label Groups", "Default Sort")
        If SHIP_ORDR_TYPE = "B" Then
            Call Build_LBL_Groups()
        End If

        Call Prompt("Setting Screen Components", "")
        Call Set_Screen_Components()
        Call Prompt("Setting Screen Components", "Finished!")

    End Sub
    Sub Build_LBL_Groups()

        Dim i As Integer
        Dim j As Integer 'Group Ceiling
        Dim n As Integer 'Label groups
        Dim c As Integer 'default print machine, 1 or 2
        Dim records As Integer
        Dim INIT_DATE As Date
        Dim INIT_OPER As String

        Dim dynWK As Recordset
        Dim dynVARS As Recordset
        Dim dynSOWLBLG1 As Recordset
        Dim dynSOWLBLG2 As Recordset

        Dim storeBeg As String
        Dim storeEnd As String
        c = 1
        INIT_DATE = Now + NowTSD
        INIT_OPER = UserID


        ORDR_GROUP_NO = cmbRecord(1).Text
        SQL = "Delete from SOWLBLG1"
        AccD.Execute(SQL)
        dynSOWLBLG1 = AccD.OpenRecordset("SOWLBLG1", dbOpenDynaset)

        'Dim dynGroup As Recordset
        'SQL = "Select Distinct ORDR_GROUP_NO from SOWSHPW1"
        'Set dynGroup = AccD.OpenRecordset(SQL, dbOpenForwardOnly)

        'Do While Not dynGroup.EOF
        dynSOWLBLG1.AddNew()
        dynSOWLBLG1.Fields("ORDR_GROUP_NO").Value = ORDR_GROUP_NO
        dynSOWLBLG1.Fields("INIT_DATE").Value = INIT_DATE
        dynSOWLBLG1.Fields("INIT_OPER").Value = INIT_OPER
        dynSOWLBLG1.Fields("LAST_DATE").Value = INIT_DATE
        dynSOWLBLG1.Fields("LAST_OPER").Value = INIT_OPER
        dynSOWLBLG1.Update()
        '   dynGroup.MoveNext
        'Loop

        SQL = "Delete from SOWLBLG2"
        AccD.Execute(SQL)
        dynSOWLBLG2 = AccD.OpenRecordset("SOWLBLG2", dbOpenDynaset)

        If SQL_PRINT_SEQ = "" Then
            SQL_PRINT_SEQ = "CUST_ADDR_CODE"
        End If

        i = 0
        n = 1
        j = 900
        SQL = "Select * from SOWSHPW1 order by " & SQL_PRINT_SEQ
        dynWK = AccD.OpenRecordset(SQL, dbOpenDynaset)
        Do While Not dynWK.EOF
            i = i + 1
            If i > j Then
                n = n + 1
                j = 900 * n
            End If
            dynWK.Edit()
            dynWK.Fields("SORT_BY").Value = i
            dynWK.Update()
            dynWK.MoveNext()
        Loop
        dynWK.Close()

        If n = 1 Then
            SQL = "Update SOWSHPW1 set LBL_GROUP_NO = 1, PRINT_STATUS = 'A', PRINT_MACHINE = '" & SHIPPER_CODE & c & "'"
            AccD.Execute(SQL)

            'SQL = "Select Distinct ORDR_GROUP_NO from SOWSHPW1"
            'Set dynGroup = AccD.OpenRecordset(SQL, dbOpenForwardOnly)

            'Do While Not dynGroup.EOF

            SQL = "Select CUST_ADDR_CODE from SOWSHPW1 WHERE SORT_BY = 1"
            '            SQL = SQL & " And ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            dynVARS = AccD.OpenRecordset(SQL, dbOpenDynaset)
            storeBeg = dynVARS.Fields(0).Value & ""

            SQL = "Select CUST_ADDR_CODE from SOWSHPW1 WHERE SORT_BY = " & i
            ' SQL = SQL & " And ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            dynVARS = AccD.OpenRecordset(SQL, dbOpenDynaset)
            storeEnd = dynVARS.Fields(0).Value & ""


            dynSOWLBLG2.AddNew()
            dynSOWLBLG2.Fields("ORDR_GROUP_NO").Value = ORDR_GROUP_NO
            dynSOWLBLG2.Fields("LBL_GROUP_NO").Value = 1
            dynSOWLBLG2.Fields("PRINT_STATUS").Value = "A"
            dynSOWLBLG2.Fields("PRINT_MACHINE").Value = SHIPPER_CODE & c
            dynSOWLBLG2.Fields("LBL_COUNT").Value = i
            dynSOWLBLG2.Fields("STORE_BEG").Value = storeBeg
            dynSOWLBLG2.Fields("STORE_END").Value = storeEnd
            dynSOWLBLG2.Update()
            'dynGroup.MoveNext
            'Loop

        Else
            Dim minSortBy As Integer
            minSortBy = j - (j / n)
            Do While n >= 1
                SQL = "Update SOWSHPW1 set LBL_GROUP_NO = " & n & " where SORT_BY <= " & i & " and SORT_BY > " & minSortBy
                AccD.Execute(SQL)

                'SQL = "Select Distinct ORDR_GROUP_NO from SOWSHPW1"
                'Set dynGroup = AccD.OpenRecordset(SQL, dbOpenForwardOnly)

                'Do While Not dynGroup.EOF

                SQL = "Select CUST_ADDR_CODE from SOWSHPW1 WHERE SORT_BY = " & minSortBy + 1
                '  SQL = SQL & " And ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                dynVARS = AccD.OpenRecordset(SQL, dbOpenDynaset)
                storeBeg = dynVARS.Fields(0).Value & ""

                SQL = "Select CUST_ADDR_CODE from SOWSHPW1 WHERE SORT_BY = " & i
                '  SQL = SQL & " And ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                dynVARS = AccD.OpenRecordset(SQL, dbOpenDynaset)
                storeEnd = dynVARS.Fields(0).Value & ""



                dynSOWLBLG2.AddNew()
                dynSOWLBLG2.Fields("ORDR_GROUP_NO").Value = ORDR_GROUP_NO
                dynSOWLBLG2.Fields("LBL_GROUP_NO").Value = n
                If n <= 2 Then
                    dynSOWLBLG2.Fields("PRINT_STATUS").Value = "A"
                Else
                    dynSOWLBLG2.Fields("PRINT_STATUS").Value = "Q"
                End If
                If ((n) Mod 2 = 0) Then
                    dynSOWLBLG2.Fields("PRINT_MACHINE").Value = SHIPPER_CODE & c + 1
                Else
                    dynSOWLBLG2.Fields("PRINT_MACHINE").Value = SHIPPER_CODE & c
                End If
                dynSOWLBLG2.Fields("LBL_COUNT").Value = i - minSortBy
                dynSOWLBLG2.Fields("STORE_BEG").Value = storeBeg
                dynSOWLBLG2.Fields("STORE_END").Value = storeEnd
                dynSOWLBLG2.Update()

                i = minSortBy
                minSortBy = minSortBy - (j / n)
                j = j - (j / n)
                n = n - 1
                '   dynGroup.MoveNext
                'Loop
            Loop

        End If

        dynVARS.Close()
        dynSOWLBLG2.Close()

    End Sub
    Sub Set_Screen_Components()
        Dim sb() As String
        Dim i As Integer
        Dim dyn As OraDynaset
        Dim z As String

        Text1(0).DATAFIELD = "SERVICE_CODE_" & SHIPPER_CODE
        Text1(1).DATAFIELD = "BILLING_CODE_" & SHIPPER_CODE
        Text1(2).DATAFIELD = "PACKAGE_CODE_" & SHIPPER_CODE
        Text2(0).DATAFIELD = "SERVICE_DESC_" & SHIPPER_CODE
        Text2(1).DATAFIELD = "BILLING_DESC_" & SHIPPER_CODE
        Text2(2).DATAFIELD = "PACKAGE_DESC_" & SHIPPER_CODE
        Text2(0).tag = "SERVICE_CODE_" & SHIPPER_CODE
        Text2(1).tag = "BILLING_CODE_" & SHIPPER_CODE
        Text2(2).tag = "PACKAGE_CODE_" & SHIPPER_CODE

        Call FM_Init_Text2()

        If chkSSOR.Value = 0 And SHIP_ORDR_TYPE = "B" Then
            ReDim sb(4, 1)
            sb(0, 0) = "BATCH_GROUP_NO"
            sb(0, 1) = "Group No"
            sb(1, 0) = "CUST_ADDR_CODE"
            sb(1, 1) = "Store"
            sb(2, 0) = "PICK_NO"
            sb(2, 1) = "Pick Ticket"
            sb(3, 0) = "TRANSIT_BUS_DAYS"
            sb(3, 1) = "Trans Days"
            sb(4, 0) = "PICK_TOT"
            sb(4, 1) = "Pick Tot"


            SSTab1.TabEnabled(3) = True
            datSOWCONF4.Refresh()

            lblTOT.Visible = False
            lblPrinted.Visible = False
            lblGroupTOT.Visible = True
            lblGroupMsg.Visible = True
        Else
            ReDim sb(3, 1)
            sb(0, 0) = "CUST_ADDR_CODE"
            sb(0, 1) = "Store"
            sb(1, 0) = "PICK_NO"
            sb(1, 1) = "Pick Ticket"
            sb(2, 0) = "TRANSIT_BUS_DAYS"
            sb(2, 1) = "Trans Days"
            sb(3, 0) = "PICK_TOT"
            sb(3, 1) = "Pick Tot"

            SSTab1.TabEnabled(3) = False

            lblTOT.Visible = True
            lblPrinted.Visible = True
            lblGroupTOT.Visible = False
            lblGroupMsg.Visible = False

        End If

        If SHIP_ORDR_TYPE = "B" Then
            grdSort.Visible = True
            For i = 0 To UBound(sb, 1)
                If sb(i, 0) = "CUST_ADDR_CODE" Then
                    grdSort.AddItem(sb(i, 0) & Chr$(9) & sb(i, 1) & Chr$(9) & "1")
                Else
                    grdSort.AddItem(sb(i, 0) & Chr$(9) & sb(i, 1))
                End If
            Next i
        Else
            grdSort.Visible = False
        End If

        ssdStatus.removeAll()
        SQL = "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SOTSHPW1' and COLUMN_NAME = 'PRINT_STATUS' "
        SQL = SQL & " and T_CODE in ('A','Q','P') order by T_CODE"
        dyn = OraD.CreateDynaset(SQL, 8&)
        Do While Not dyn.EOF
            z = dyn.Fields(0).Value & vbTab & dyn.Fields(1).Value
            ssdStatus.AddItem(z)
            dyn.MoveNext()
        Loop

        ssdTarget.removeAll()
        SQL = "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SOTSHPW1' and COLUMN_NAME = 'PRINT_MACHINE' "
        If SHIPPER_CODE = "FDX" Then
            SQL = SQL & " and T_CODE like '" & SHIPPER_CODE & "%' order by T_CODE"
        Else
            SQL = SQL & " and T_CODE like '" & SHIPPER_CODE & "%' order by T_CODE"
        End If
        dyn = OraD.CreateDynaset(SQL, 8&)
        Do While Not dyn.EOF
            z = dyn.Fields(0).Value & vbTab & dyn.Fields(1).Value
            ssdTarget.AddItem(z)
            dyn.MoveNext()
        Loop
        dyn.Close()

        datSOWSHPW1.Refresh()
        grdSOWSHPW1.Caption = "Details for " & SHIPPER_CODE_CAPTION & " " & SHIP_ORDR_TYPE_CAPTION
        datSOWLBLG2.Refresh()

        Call Calc_Group_Stats("000001")
        Call Load_Style_Weight()
        SSTab1.TabEnabled(4) = False
        SSTab1.Tab = 0

    End Sub
    Sub Calc_Group_Stats(group_no As String)

        Dim dynTOT As Recordset
        Dim batTot As String
        Dim grp As Integer

        If chkSSOR.Value = 0 Then
            If SHIP_ORDR_TYPE = "B" Then
                SQL = "select sum(STORE_COUNT) from SOWCONF4 where PRINT_GROUP = '1'"
                grp = Val(grdSOWCONF4.Columns("CONFIG_NO").Text & "")
            Else
                SQL = "select count (*) from SOWSHPW1"
                grp = 1
            End If
        Else
            SQL = "select count (*) from SOWSHPW1"
            grp = 1
        End If
        dynTOT = AccD.OpenRecordset(SQL)
        batTot = dynTOT.Fields(0).Value & ""
        lblGroupTOT.Caption = batTot
        lblTOT.Caption = batTot
        dynTOT.Close()

        SQL = "SELECT * FROM SOWSHPW1 WHERE BATCH_GROUP_NO = '" & Format$(grp, "000000") & "'"
        datSOWSHPW1.RecordSource = SQL
        datSOWSHPW1.Refresh()

    End Sub
 
    Sub Update_Record()

        Screen.MousePointer = 11

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim dynWK As Recordset

        Dim Range As String
        Dim pick As String
        Dim z As String
        Dim msgInfo As String
        Dim grp As Integer
        Dim SVC_TYP As String
        Dim PAY_TYP As String
        Dim PAC_TYP As String
        Dim SHP_DATE As String
        Dim viewToUse As String

        OraS.BeginTrans()

        'Get shipment parameters, FedEx uses codes, UPS uses descriptions
        If SHIPPER_CODE = "FDX" Then
            Call Prompt("Now Sending Records to FedEx ...", "")
            SVC_TYP = Text1(0).Text
            PAY_TYP = Text1(1).Text
            PAC_TYP = Text1(2).Text

        ElseIf SHIPPER_CODE = "UPS" Then
            Call Prompt("Now Sending Records to UPS ...", "")
            SVC_TYP = Text2(0).Text
            PAY_TYP = Text2(1).Text
            PAC_TYP = Text2(2).Text
        Else
            Call Create_DHL()
            OraS.CommitTrans()
        GoSub Leave_Routine
        End If

        If SHIPPER_CODE = "FDX" Then
            SHP_DATE = Format$(SSMonth1.Date, "MM/DD/YYYY")
        Else
            SHP_DATE = Format$(SSMonth1.Date, "MM/DD/YYYY")
        End If

        'Update Shipment Parameters
        SQL = "UPDATE SOWSHPW1 SET "
        SQL = SQL & " SHIP_DATE = '" & SHP_DATE & "'"
        AccD.Execute(SQL)

        'Delete Records Not Being Printed, only applicable to batched orders
        If (chkSSOR.Value <> "1" And SHIP_ORDR_TYPE = "K") Then
            SQL = "SELECT * FROM SOWCONF4"
            dynWK = AccD.OpenRecordset(SQL, dbOpenDynaset)
            Do While Not dynWK.EOF
                If dynWK.Fields("PRINT_GROUP").Value & "" <> "1" Then
                    grp = Val(dynWK.Fields("CONFIG_NO").Value & "")
                    SQL = "Delete from SOWSHPW1 where"
                    SQL = SQL & " BATCH_GROUP_NO = '" & Format(grp, "000000") & "'"
                    AccD.Execute(SQL)
                End If
                dynWK.MoveNext()
            Loop
            dynWK.Close()
        End If

        'Update Label Printing Parameters
        SQL = "Update SOWSHPW1,SOWLBLG2"
        SQL = SQL & " Set SOWSHPW1.PRINT_STATUS = SOWLBLG2.PRINT_STATUS"
        SQL = SQL & ", SOWSHPW1.PRINT_MACHINE = SOWLBLG2.PRINT_MACHINE"
        SQL = SQL & " where SOWSHPW1.LBL_GROUP_NO = SOWLBLG2.LBL_GROUP_NO"
        AccD.Execute(SQL)

        'Third Party Info if Necessary
        If tp = True Then
            For i = 3 To 11
                SQL = "Update SOWSHPW1 set " & Text2(i).tag & "_TP = '" & Text2(i).Text & "'"
                AccD.Execute(SQL)
            Next i
        End If

        'Clean Up Un-Matchable Multi-Carton Pick Tickets, FedEx Only
        If SHIP_ORDR_TYPE = "K" And SHIPPER_CODE = "FDX" Then
            SQL = "Select PICK_NO, COUNT(*) from SOWSHPW1 GROUP BY PICK_NO HAVING COUNT(*) > 1"
            dynWK = AccD.OpenRecordset(SQL, dbOpenDynaset)
            Do While Not dynWK.EOF
                SQL = "UPDATE SOWSHPW1 SET CART_NO = null WHERE PICK_NO = '" & dynWK.Fields("PICK_NO").Value & "'"
                AccD.Execute(SQL)
                dynWK.MoveNext()
            Loop
            dynWK.Close()
        End If

        Dim dynDelete As Recordset
        SQL = "Select Distinct ORDR_GROUP_NO from SOWSHPW1"
        dynDelete = AccD.OpenRecordset(SQL, dbOpenForwardOnly)

        Do While Not dynDelete.EOF
            'Prepare Oracle for Update
            SQL = "DELETE FROM SOTSHPW1"
            SQL = SQL & " WHERE ORDR_GROUP_NO = '" & dynDelete.Fields("ORDR_GROUP_NO").Value & "'"
            OraD.ExecuteSQL(SQL)


            dynDelete.MoveNext()
        Loop

        SQL = "DELETE FROM SOTLBLG1"
        SQL = SQL & " WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
        OraD.ExecuteSQL(SQL)

        SQL = "DELETE FROM SOTLBLG2"
        SQL = SQL & " WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
        OraD.ExecuteSQL(SQL)


        SQL = "UPDATE SOTSHPW1"
        SQL = SQL & " SET PRINT_STATUS = 'Q' WHERE PRINT_STATUS = 'A'"
        OraD.ExecuteSQL(SQL)


        'Send Records to Oracle
        Call Acc_to_Ora("SOWSHPW1", "", "")
        Call Acc_to_Ora("SOWLBLG1", "", "")
        Call Acc_to_Ora("SOWLBLG2", "", "")

        'Assign correct View to use for PL/SQL Driver
        viewToUse = "SOV" & SHIPPER_CODE & SHIP_ORDR_TYPE & "1"

        'update acct no from customer table
        SQL = " BEGIN"
        SQL = SQL & "  DECLARE"
        SQL = SQL & "   CURSOR C2 IS"
        SQL = SQL & " Select Distinct W1.CUST_CODE, W1.CUST_ADDR_CODE, FDX_ACCT_NO from SOTSHPW1 W1, ARTCUST2 T2"
        SQL = SQL & " Where W1.PRINT_STATUS = 'A'"
        SQL = SQL & " And W1.CUST_CODE = T2.CUST_CODE"
        SQL = SQL & " And W1.CUST_ADDR_CODE = T2.CUST_ADDR_CODE"
        SQL = SQL & " And FDX_ACCT_NO is not null;"
        SQL = SQL & "  BEGIN"
        SQL = SQL & "   FOR R2 IN C2 LOOP"
        SQL = SQL & " UPDATE SOTSHPW1 SET SHIPPER_ACCT_CODE = R2.FDX_ACCT_NO WHERE CUST_CODE = R2.CUST_CODE"
        SQL = SQL & "  and CUST_ADDR_CODE = R2.CUST_ADDR_CODE and PRINT_STATUS = 'A';"
        SQL = SQL & "  END LOOP;"
        SQL = SQL & " END; END;"
        OraD.ExecuteSQL(SQL)

        'Update Pickt Ticket Records
        Call Prompt("Now Updating Pick Table with Package Weights ", "")
        SQL = " BEGIN DECLARE CURSOR C1 IS"
        SQL = SQL & " SELECT PICK_NO, SUM(PACKAGE_WEIGHT) PACKAGE_WEIGHT FROM " & viewToUse & " GROUP BY PICK_NO;"
        SQL = SQL & " BEGIN FOR R1 IN C1 LOOP"
        SQL = SQL & "  UPDATE SOTPICK1 SET PICK_TOTAL_WGT = R1.PACKAGE_WEIGHT WHERE SOTPICK1.PICK_NO = R1.PICK_NO;"
        SQL = SQL & " END LOOP;END;END;"
        OraD.ExecuteSQL(SQL)

        'Update Cart records if possible
        If SHIP_ORDR_TYPE = "B" Or SHIPPER_CODE = "UPS" Then
            Call Prompt("Now Updating Carton Table with Package Weights ", "")
            SQL = " BEGIN DECLARE CURSOR C1 IS"
            SQL = SQL & " SELECT CART_NO, PACKAGE_WEIGHT FROM " & viewToUse & ";"
            SQL = SQL & " BEGIN FOR R1 IN C1 LOOP"
            SQL = SQL & "  UPDATE SOTCART1 SET CART_TOTAL_WGT_CALC = R1.PACKAGE_WEIGHT WHERE SOTCART1.CART_NO = R1.CART_NO;"
            SQL = SQL & " END LOOP;END;END;"
            OraD.ExecuteSQL(SQL)
        End If

        'Update Configuration Table if Necessary
        If chkSSOR.Value <> "1" And SHIP_ORDR_TYPE = "B" Then
            Call Prompt("Now Updating Configuration Table.", "")
            SQL = " BEGIN DECLARE CURSOR C1 IS"
            SQL = SQL & " SELECT ORDR_NO, CUST_ADDR_CODE, SHIP_DATE FROM " & viewToUse & ";"
            SQL = SQL & " BEGIN FOR R1 IN C1 LOOP"
            SQL = SQL & "  UPDATE SOTCONF2 SET BATCHED = 'Y' "
            SQL = SQL & ", SHIP_DATE = TO_DATE(R1.SHIP_DATE,'yyyymmdd')"
            SQL = SQL & "  WHERE SOTCONF2.ORDR_NO = R1.ORDR_NO"
            SQL = SQL & "  AND SOTCONF2.CUST_STORE_NO = R1.CUST_ADDR_CODE;"
            SQL = SQL & " END LOOP;END;END;"
            OraD.ExecuteSQL(SQL)
        End If

        'Finish up, now thats a nicely documented update routine
        Call Prompt("Done", "")
        OraS.CommitTrans()

Leave_Routine:
        Screen.MousePointer = 0
        MsgBox("Records Successfully Batched", 48 + 0, "Success")
        Call cmdExecute(4)

    End Sub
    Sub Load_Style_Weight()

        SQL = "SELECT SOTCART2.STYLE_CODE, SUM(QTY_PACKED) QTY, ICTSTYL1.STYLE_WEIGHT" & vbCrLf
        SQL = SQL & ", SOTCART2.CART_NO,  SOTCART2.SIZE_DESC, SOTCART2.ORDR_NO, SOTCART2.ORDR_LNO " & vbCrLf
        SQL = SQL & "  FROM SOTCART2, ICTSTYL1 WHERE" & vbCrLf
        SQL = SQL & "  SOTCART2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf
        SQL = SQL & "  AND SOTCART2.CART_NO = '" & grdSOWSHPW1.Columns("CART_NO").Text & "'" & vbCrLf
        SQL = SQL & " GROUP BY SOTCART2.STYLE_CODE, ICTSTYL1.STYLE_WEIGHT"
        SQL = SQL & ", SOTCART2.CART_NO,  SOTCART2.SIZE_DESC, SOTCART2.ORDR_NO, SOTCART2.ORDR_LNO"
        SQL = SQL & " ORDER BY SOTCART2.ORDR_NO, SOTCART2.ORDR_LNO"

        datSOTCARTX.RecordSource = SQL
        datSOTCARTX.Refresh()

    End Sub

    Private Sub cmbRecord_CloseUp(Index As Integer)
        If cmbRecord(Index).Text = "" Then
            Exit Sub
        Else
            If optBATCH_TYPE(0).Value = True Then
                Call Validate_Record(Index)
            Else
                Call Validate_Batch()
            End If
            ASFMAIN1.cmdMain_Click(2)
        End If
    End Sub


    Private Sub cmbRecord_DropDown(Index As Integer)

        If txtCode(0).Text = "" Then
            MsgBox("You Must Specify a Customer Code")
            Exit Sub
        End If

        SQL = "Select * from SOTORDR0 where SOTORDR0.CUST_CODE = '" & txtCode(0).Text & "'"
        If txtCode(2).Text <> "" Then
            SQL = SQL & " and SOTORDR0.ORDR_CUST_PO = '" & Trim(txtCode(2).Text & "") & "'"
        End If

        datSOTORDR0.RecordSource = SQL
        datSOTORDR0.Refresh()

    End Sub

    Private Sub cmbRecord_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            If cmbRecord(Index).Text = "" Then
                Exit Sub
            Else
                If optBATCH_TYPE(0).Value = True Then
                    Call Validate_Record(Index)
                Else
                    Call Validate_Batch()
                End If
                ASFMAIN1.cmdMain_Click(2)
            End If
        End If
    End Sub

    Sub Validate_Record(Index As Integer)

        Dim dyn As OraDynaset

        If cmbRecord(Index).Text <> "" Then
            SQL = "Select ORDR_GROUP_NO from SOTORDR0 where ORDR_GROUP_NO"
            SQL = SQL & " = '" & Format$(Val(cmbRecord(Index).Text), "0000000000") & "'"
            dyn = OraD.CreateDynaset(SQL, 8&)

            If Not dyn.EOF Then
                cmbRecord(Index).Text = dyn.Fields(0).Value
                SQL = "Select CUST_CODE, ORDR_CUST_PO from SOTORDR1 where ORDR_GROUP_NO"
                SQL = SQL & " = '" & cmbRecord(Index).Text & "'"
                dyn = OraD.CreateDynaset(SQL, 8&)
                txtCode(2).Text = dyn.Fields(1).Value & ""
                CUST_CODE = dyn.Fields(0).Value & ""
                ORDR_GROUP_NO = Format$(Val(cmbRecord(Index).Text), "0000000000")
            Else
                MsgBox("No record of " & lblRecord(Index).Caption & " " & cmbRecord(Index).Text)
                cmbRecord(Index).Text = ""
                Exit Sub
            End If
            dyn.Close()
        End If
    End Sub
    Sub Import_Log(fn As String)
        Call Prompt("Connecting to warehouse......", "")
        Call Import_Shipment_Log(, fn)
        Call Prompt("Information Import Complete", "")
    End Sub

    Sub Select_File()
        Dim z As String

        CommonDialog1.InitDir = "G:\VAN\WMS\Shipping\Export"
        CommonDialog1.FileName = "*.log"
        CommonDialog1.ShowOpen()
        z = CommonDialog1.FileName

        If z <> "" And z <> "*.log" Then
            Call Import_Log(z)
        End If

    End Sub

    Private Sub cmdClear_Click(Index As Integer)
        Dim i As Integer
        For i = 3 To 10
            Text2(i).Text = ""
        Next i
    End Sub

    Private Sub cmdImport_Click()
        Dim z As String
        CommonDialog1.InitDir = Path
        CommonDialog1.FileName = "*.log"
        CommonDialog1.ShowOpen()
        z = CommonDialog1.FileName

        If z <> "" And z <> "*.log" Then
            Call Import_Log(z)
        End If
        Call Check_Shipments_Log()
    End Sub

    Sub Check_Shipments_Log()
        If UserID = "gcv" Then Exit Sub
        cmdImport.Visible = False
        File1.Path = "G:\VAN\WMS\Shipping\Export\FedEx01\"
        File1.Pattern = "Shipments.log"
        If File1.ListCount > 0 Then
            cmdImport.Visible = True
            Path = "G:\VAN\WMS\Shipping\Export\FedEx01\"
        End If
        File1.Path = "G:\VAN\WMS\Shipping\Export\FedEx02\"
        File1.Pattern = "Shipments.log"
        If File1.ListCount > 0 Then
            cmdImport.Visible = True
            Path = "G:\VAN\WMS\Shipping\Export\FedEx02\"
        End If

        Dim dynSOTSHPTT As OraDynaset
        SQL = "Select * from SOTSHPTT"
        SQL = SQL & " WHERE SHIPPER_CODE = 'UPS' "
        SQL = SQL & " And IMPORT_STATUS is Null"
        dynSOTSHPTT = OraD.CreateDynaset(SQL, 8&)

        If dynSOTSHPTT.EOF Then
            cmdUPS.Visible = False
        Else
            cmdUPS.Visible = True
        End If
        dynSOTSHPTT.Close()
    End Sub

    Private Sub cmdTP_Click(Index As Integer)
        Select Case Index
            Case 0
                'load saved Third Party Billing
                Call Prompt("Retreiving Third Party Info......", "")
                Call Load_Saved_TP_Info()
                Call Prompt("", "")
            Case 1
                'Save Third Party Billing
                Call Prompt("Saving Third Party Info......", "")
                Call Save_TP_Info()
                MsgBox("Third Party Info Saved", vbOKOnly, "Success!")
                Call Prompt("", "")
        End Select
    End Sub

    Sub Load_Saved_TP_Info()
        Dim i As Integer
        Dim j As Integer

        Dim dynSOTTPBL1 As OraDynaset

        aMessage = CUST_CODE
        SOFTPBL1.Show(1)

        If aMessage <> "" Then
            i = CInt(aMessage)

            SQL = "Select CUST_NAME, CUST_ADDR1, CUST_ADDR2, CUST_CITY, CUST_STATE, CUST_ZIP_CODE, CUST_CONTACT, THIRD_PARTY_ACCT, CUST_COUNTRY"
            SQL = SQL & " from SOTTPBL1 where CUST_CODE = '" & CUST_CODE & "' AND TP_LNO = " & i
            dynSOTTPBL1 = OraD.CreateDynaset(SQL, 8&)

            For j = 0 To 8
                Text2(j + 3).Text = dynSOTTPBL1.Fields(j).Value & ""
            Next j

            dynSOTTPBL1.Close()
        End If

    End Sub
    Sub Save_TP_Info()
        OraS.BeginTrans()
        Dim dyn As OraDynaset
        Dim dynSOTTPBL1 As OraDynaset
        Dim tpLno As Integer

        SQL = "Select MAX(TP_LNO) from SOTTPBL1 where CUST_CODE = '" & CUST_CODE & "'"
        dyn = OraD.CreateDynaset(SQL, 8&)
        tpLno = Val(dyn.Fields(0).Value & "") + 1
        dyn.Close()

        SQL = "Select * from SOTTPBL1 where rownum < 1"
        dynSOTTPBL1 = OraD.CreateDynaset(SQL, 8&)

        dynSOTTPBL1.AddNew()
        dynSOTTPBL1.Fields("CUST_CODE").Value = CUST_CODE
        dynSOTTPBL1.Fields("TP_LNO").Value = tpLno
        dynSOTTPBL1.Fields("CUST_NAME").Value = Text2(3).Text
        dynSOTTPBL1.Fields("CUST_ADDR1").Value = Text2(4).Text
        dynSOTTPBL1.Fields("CUST_ADDR2").Value = Text2(5).Text
        dynSOTTPBL1.Fields("CUST_CITY").Value = Text2(6).Text
        dynSOTTPBL1.Fields("CUST_STATE").Value = Text2(7).Text
        dynSOTTPBL1.Fields("CUST_ZIP_CODE").Value = Text2(8).Text
        dynSOTTPBL1.Fields("CUST_COUNTRY").Value = Text2(11).Text
        dynSOTTPBL1.Fields("CUST_CONTACT").Value = Text2(9).Text
        dynSOTTPBL1.Fields("THIRD_PARTY_ACCT").Value = Text2(10).Text
        dynSOTTPBL1.Update()

        dynSOTTPBL1.Close()
        OraS.CommitTrans()

    End Sub

    Private Sub cmdUPS_Click()
        Dim dynSOTSHPB1 As OraDynaset
        Dim dynSOTSHPM1 As OraDynaset

        Dim dynSOTSHPTT As OraDynaset
        SQL = " Select ORDR_GROUP_NO, K1.ORDR_NO, TT.PICK_NO," & vbCr
        SQL = SQL & " TT.CART_NO, TRACK_NO, '' REFERENCE_NO, PACKAGE_WEIGHT," & vbCr
        SQL = SQL & " TRACK_NO MASTER_TRACK_NO,  ORDR_SHIP_DATE AS SHIP_DATE," & vbCr
        SQL = SQL & " SHIP_BOL_NO AS SHIPMENT_CTL_NO" & vbCr
        SQL = SQL & " from SOTSHPTT TT, SOTPICK1 K1, SOTORDR1 R1 " & vbCr
        SQL = SQL & " Where TT.PICK_NO = K1.PICK_NO " & vbCr
        SQL = SQL & " And K1.ORDR_NO = R1.ORDR_NO" & vbCr
        SQL = SQL & " and SHIPPER_CODE = 'UPS' " & vbCr
        SQL = SQL & " and IMPORT_STATUS is Null " & vbCr
        'SQL = SQL & " and pick_status = 'P'"
        'SQL = SQL & " And ORDR_GROUP_NO = '0000047777'" & vbCr
        dynSOTSHPTT = OraD.CreateDynaset(SQL, 8&)

        OraS.BeginTrans()

        Do While Not dynSOTSHPTT.EOF
            SQL = "Select * from SOTSHPB1 Where PICK_NO = '" & dynSOTSHPTT.Fields("PICK_NO").Value & "'"
            SQL = SQL & " And Cart_no = '" & dynSOTSHPTT.Fields("CART_NO").Value & "'"
            SQL = SQL & " And ORDR_NO = '" & dynSOTSHPTT.Fields("ORDR_NO").Value & "'"
            SQL = SQL & " And ORDR_GROUP_NO = '" & dynSOTSHPTT.Fields("ORDR_GROUP_NO").Value & "'"
            dynSOTSHPB1 = OraD.CreateDynaset(SQL, 8&)
            If dynSOTSHPB1.EOF Then
                dynSOTSHPB1.AddNew()
            Else
                dynSOTSHPB1.Edit()
            End If
            dynSOTSHPB1.Fields("ORDR_GROUP_NO").Value = dynSOTSHPTT.Fields("ORDR_GROUP_NO").Value
            dynSOTSHPB1.Fields("ORDR_NO").Value = dynSOTSHPTT.Fields("ORDR_NO").Value
            dynSOTSHPB1.Fields("PICK_NO").Value = dynSOTSHPTT.Fields("PICK_NO").Value
            dynSOTSHPB1.Fields("TRACK_NO").Value = dynSOTSHPTT.Fields("TRACK_NO").Value
            dynSOTSHPB1.Fields("CART_NO").Value = dynSOTSHPTT.Fields("CART_NO").Value
            dynSOTSHPB1.Fields("REFERENCE_NO").Value = ""
            dynSOTSHPB1.Fields("PACKAGE_WEIGHT").Value = dynSOTSHPTT.Fields("PACKAGE_WEIGHT").Value
            dynSOTSHPB1.Fields("MASTER_TRACK_NO").Value = dynSOTSHPTT.Fields("TRACK_NO").Value
            dynSOTSHPB1.Fields("SHIP_DATE").Value = dynSOTSHPTT.Fields("SHIP_DATE").Value
            dynSOTSHPB1.Fields("INIT_DATE").Value = Format$(Now + NowTSD, "dd-mmm-yy")
            dynSOTSHPB1.Fields("INIT_USER").Value = UserID
            dynSOTSHPB1.Fields("SHIPMENT_CTL_NO").Value = dynSOTSHPTT.Fields("SHIPMENT_CTL_NO").Value
            dynSOTSHPB1.Update()

            SQL = "Select * from SOTSHPB1 Where PICK_NO = '" & dynSOTSHPTT.Fields("PICK_NO").Value & "'"
            SQL = SQL & " And ORDR_NO = '" & dynSOTSHPTT.Fields("ORDR_NO").Value & "'"
            SQL = SQL & " And ORDR_GROUP_NO = '" & dynSOTSHPTT.Fields("ORDR_GROUP_NO").Value & "'"
            SQL = SQL & " And TRACK_NO = '" & dynSOTSHPTT.Fields("TRACK_NO").Value & "'"
            dynSOTSHPM1 = OraD.CreateDynaset(SQL, 8&)
            If dynSOTSHPM1.EOF Then
                dynSOTSHPM1.AddNew()
            Else
                dynSOTSHPM1.Edit()
            End If
            dynSOTSHPM1.Fields("ORDR_GROUP_NO").Value = dynSOTSHPTT.Fields("ordr_group_no").Value
            dynSOTSHPM1.Fields("ORDR_NO").Value = dynSOTSHPTT.Fields("ORDR_NO").Value
            dynSOTSHPM1.Fields("PICK_NO").Value = dynSOTSHPTT.Fields("PICK_NO").Value
            dynSOTSHPM1.Fields("TRACK_NO").Value = dynSOTSHPTT.Fields("TRACK_NO").Value
            dynSOTSHPM1.Fields("REFERENCE_NO").Value = ""
            dynSOTSHPM1.Fields("PACKAGE_WEIGHT").Value = dynSOTSHPTT.Fields("PACKAGE_WEIGHT").Value
            dynSOTSHPM1.Fields("MASTER_TRACK_NO").Value = dynSOTSHPTT.Fields("TRACK_NO").Value
            dynSOTSHPM1.Fields("SHIP_DATE").Value = dynSOTSHPTT.Fields("SHIP_DATE").Value
            dynSOTSHPM1.Fields("INIT_DATE").Value = Format$(Now + NowTSD, "dd-mmm-yy")
            dynSOTSHPM1.Fields("INIT_USER").Value = UserID
            dynSOTSHPM1.Fields("SHIPMENT_CTL_NO").Value = dynSOTSHPTT.Fields("SHIPMENT_CTL_NO").Value
            dynSOTSHPM1.Update()

            dynSOTSHPTT.MoveNext()
        Loop
        dynSOTSHPTT.Close()
        OraS.CommitTrans()

        OraS.BeginTrans()
        SQL = " BEGIN DECLARE CURSOR C1 IS"
        SQL = SQL & "  SELECT SOTPICK1.SHIP_BOL_NO, SOTSHPB1.PICK_NO, Max(SOTSHPB1.MASTER_TRACK_NO) MASTER_TRACK_NO"
        SQL = SQL & " , SUM(SOTSHPB1.PACKAGE_WEIGHT) PACKAGE_WEIGHT FROM SOTSHPB1, SOTPICK1, SOTSHPTT "
        SQL = SQL & "  WHERE SOTSHPTT.PICK_NO = SOTPICK1.PICK_NO"
        SQL = SQL & "  And SOTSHPB1.PICK_NO = SOTPICK1.PICK_NO"
        SQL = SQL & "  And SHIPPER_CODE = 'UPS' "
        SQL = SQL & "  And IMPORT_STATUS is Null "
        SQL = SQL & "  GROUP BY SOTPICK1.SHIP_BOL_NO, SOTSHPB1.PICK_NO;"
        SQL = SQL & " BEGIN FOR R1 IN C1 LOOP"
        SQL = SQL & "  UPDATE SOTPICK1 SET PICK_TOTAL_WGT = R1.PACKAGE_WEIGHT WHERE SOTPICK1.PICK_NO = R1.PICK_NO;"
        SQL = SQL & "  UPDATE SOTSHIP1 SET BILL_OF_LADING_NO = R1.MASTER_TRACK_NO "
        SQL = SQL & "  WHERE SOTSHIP1.SHIP_BOL_NO = R1.SHIP_BOL_NO;"
        SQL = SQL & " END LOOP;END;END;"
        OraD.ExecuteSQL(SQL)

        'Call Prompt("Updating Cartons with Package Weights and Tracking No", "Shipment Ctl No: " & SHIPMENT_CTL_NO)
        SQL = " BEGIN DECLARE CURSOR C1 IS"
        SQL = SQL & "  SELECT CART_NO, TRACK_NO, PACKAGE_WEIGHT  FROM  SOTSHPTT"
        SQL = SQL & "  WHERE SHIPPER_CODE = 'UPS' "
        SQL = SQL & "  And IMPORT_STATUS is Null; "
        SQL = SQL & " BEGIN FOR R1 IN C1 LOOP"
        SQL = SQL & "  UPDATE SOTCART1 SET CART_TOTAL_WGT_CALC = R1.PACKAGE_WEIGHT"
        SQL = SQL & " , CART_TRACKING_NO = R1.TRACK_NO WHERE SOTCART1.CART_NO = R1.CART_NO;"
        SQL = SQL & " END LOOP;END;END;"
        OraD.ExecuteSQL(SQL)

        SQL = "Update SOTSHPTT Set IMPORT_STATUS = '1'"
        SQL = SQL & "  WHERE SHIPPER_CODE = 'UPS' "
        SQL = SQL & "  And IMPORT_STATUS is Null"
        OraD.ExecuteSQL(SQL)

        cmdUPS.Visible = False
        OraS.CommitTrans()
    End Sub


    Private Sub fpCarton_Change()

        Sql = "Update SOWSHPW1 set PACKAGE_WEIGHT = CALC_CARTON_WEIGHT + " & fpCarton.Value
        AccD.Execute(Sql)

        datSOWSHPW1.Refresh()

    End Sub

    Private Sub grdSort_BtnClick()
        grdSort.Redraw = False
        If grdSort.Columns("Seq").Text <> "" Then
            grdSort.Columns("Seq").Text = ""
        Else
            grdSort.Columns("Seq").Text = "8"
        End If
        grdSort.Update()
        Dim r As String
        Dim rr As String
        Dim V As Object
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim rFirst As Integer
        Dim rlast As Integer
        Dim fFirst As String
        Dim fLast As String

        Dim fr As Object
        fr = grdSort.FirstRow
        r = Space$(grdSort.Rows)
        rr = Space$(grdSort.Rows)
        rFirst = 0
        rlast = 0
        For i = 0 To grdSort.Rows - 1
            V = grdSort.AddItemBookmark(i)
            Mid$(r, i + 1, 1) = grdSort.Columns("Seq").CellText(V)
            If grdSort.Columns("FIELD").CellText(V) = fFirst Then
                rFirst = i + 1
            End If
            If grdSort.Columns("FIELD").CellText(V) = fLast Then
                rlast = i + 1
            End If
        Next i

        If rFirst <> 0 Then
            Mid$(r, rFirst, 1) = "0"
        End If
        If rlast <> 0 Then
            Mid$(r, rlast, 1) = "9"
        End If
        j = 1
        For i = 0 To 9
            k = InStr(r, Format$(i, "0"))
            If k <> 0 Then
                Mid$(r, k, 1) = " "
                Mid$(rr, k, 1) = Format$(j, "0")
                j = j + 1
                i = i - 1
            End If
        Next i

        ReDim sqlPrint(grdSort.Rows)
        For i = 0 To grdSort.Rows - 1
            V = grdSort.AddItemBookmark(i)
            grdSort.Bookmark = V
            If Mid$(rr, i + 1, 1) <> " " Then
                grdSort.Columns("Seq").Text = Mid$(rr, i + 1, 1)
                If Mid$(rr, i + 1, 1) <> "" Then
                    sqlPrint(CInt(Mid$(rr, i + 1, 1))) = grdSort.Columns("FIELD").Value
                End If
            End If
        Next i

        SQL_PRINT_SEQ = ""
        For i = 1 To UBound(sqlPrint)
            If sqlPrint(i) <> "" Then
                If sqlPrint(i) = "TRANSIT_BUS_DAYS" Or sqlPrint(i) = "PICK_TOT" Then
                    SQL_PRINT_SEQ = SQL_PRINT_SEQ & ", " & sqlPrint(i) & " Desc"
                Else
                    SQL_PRINT_SEQ = SQL_PRINT_SEQ & ", " & sqlPrint(i)
                End If
            End If
        Next i

        If SQL_PRINT_SEQ <> "" Then
            SQL_PRINT_SEQ = Mid$(SQL_PRINT_SEQ, 2)
        End If

        grdSort.Update()
        grdSort.FirstRow = fr
        grdSort.Redraw = True

        Call Build_LBL_Groups()
        datSOWLBLG2.Refresh()

    End Sub

    Private Sub grdSOWCONF4_AfterUpdate(RtnDispErrMsg As Integer)
        Dim lbs As Double
        Dim grp As Integer
        lbs = Val(grdSOWCONF4.Columns("WEIGHT").Text & "")
        grp = Val(grdSOWCONF4.Columns("CONFIG_NO").Text & "")
        SQL = "Update SOWSHPW1 set PACKAGE_WEIGHT = " & lbs & " where "
        SQL = SQL & " BATCH_GROUP_NO = '" & Format$(grp, "000000") & "'"
        AccD.Execute(SQL)
        Call Calc_Group_Stats(CStr(Format$(grp, "000000")))
    End Sub

    Private Sub grdSOWCONF4_BeforeColUpdate(ByVal ColIndex As Integer, ByVal OldValue As Object, Cancel As Integer)
        If grdSOWCONF4.Columns(ColIndex).NAME = "PRINT_GROUP" Then
            If grdSOWCONF4.Columns(ColIndex).Text = "-1" Then
                grdSOWCONF4.Columns(ColIndex).Text = "1"
            End If
        End If
    End Sub

    Private Sub grdSOWCONF4_BeforeUpdate(Cancel As Integer)
        If grdSOWCONF4.Columns("PRINT_GROUP").Text = "" Then
            grdSOWCONF4.Columns("PRINT_GROUP").Text = "0"
        End If
        If Val(grdSOWCONF4.Columns("WEIGHT").Text & "") > 0 Then
            'grdSOWCONF4.Columns("PRINT_GROUP").Text = "1"
        Else
            'grdSOWCONF4.Columns("PRINT_GROUP").Text = "0"
        End If

    End Sub

    Private Sub grdSOWCONF4_KeyPress(KeyAscii As Integer)
        If KeyAscii <> 46 Then
            Call ASBMAIN1.grd_KeyPress(KeyAscii, grdSOWCONF4)
        End If
    End Sub

    Private Sub grdSOWCONF4_LostFocus()
        grdSOWCONF4.Update()
    End Sub

    Private Sub grdSOWCONF4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        grdSOWCONF4.ActiveCell.SelStart = 0
        grdSOWCONF4.ActiveCell.SelLength = Len(grdSOWCONF4.ActiveCell.Text)
    End Sub

    Private Sub grdSOWCONF4_RowColChange(ByVal LastRow As Object, ByVal LastCol As Integer)
        gNo = grdSOWCONF4.Columns("CONFIG_NO").Text & ""
        grdSOWCONF4.Update()
        Call Calc_Group_Stats(gNo)
    End Sub


    Private Sub grdSOWCONF4_RowLoaded(ByVal Bookmark As Object)
        If grdSOWCONF4.Columns("ALREADY_PRINTED").Text = "Y" Then
            grdSOWCONF4.Columns("PRINT_GROUP").CellStyleSet("Red")
        End If
    End Sub

    Private Sub grdSOWLBLG2_InitColumnProps()
        grdSOWLBLG2.Columns("PRINT_STATUS").DropDownHwnd = ssdStatus.hWnd
        grdSOWLBLG2.Columns("PRINT_MACHINE").DropDownHwnd = ssdTarget.hWnd
    End Sub

    Private Sub grdSOWSHPW1_HeadClick(ByVal ColIndex As Integer)
        Call Sort_By_Headclick(grdSOWSHPW1, ColIndex)
    End Sub

    Private Sub grdSOWSHPW1_RowColChange(ByVal LastRow As Object, ByVal LastCol As Integer)
        Call Load_Style_Weight()
    End Sub

    Private Sub optBATCH_TYPE_Click(Index As Integer)
        If Index = 0 Then
            lblRecord(1).Caption = "Group No"
        Else
            lblRecord(1).Caption = "Batch No"
        End If
    End Sub

    Private Sub optSHPR_Click(Index As Integer)
        If optSHPR(1) Then
            chkSSOR.Value = "1"
            chkSSOR.Visible = False
        Else
            chkSSOR.Value = "1"
            chkSSOR.Visible = True
        End If
    End Sub

    Private Sub optType_Click(Index As Integer)
        If optType(0) Then
            chkSSOR.Value = "1"
            chkSSOR.Visible = True
        Else
            chkSSOR.Value = "0"
            chkSSOR.Visible = False
        End If
    End Sub

    Private Sub ssdStatus_InitColumnProps()
        grdSOWLBLG2.Update()
    End Sub

    Private Sub ssdTarget_CloseUp()
        grdSOWLBLG2.Update()
    End Sub

    Private Sub Text1_Change(Index As Integer)
        Call FM_Text1_Change(Text1(Index))
        Select Case Mid$(Text1(Index).DATAFIELD, 1, Len(Text1(Index).DATAFIELD) - 4)
            Case "BILLING_CODE"
                If Text1(Index).Text = "3" Or Text1(Index).Text = "T" Then
                    SSTab1.TabEnabled(4) = True
                    SSTab1.Tab = 4
                    tp = True
                Else
                    SSTab1.TabEnabled(4) = False
                    tp = False
                End If
                If SHIPPER_CODE = "FDX" Then
                    SQL = "Update SOWSHPW1 set PAYMENT_TYPE = '" & Text1(1).Text & "'"
                Else
                    If chkByLine = "1" Then
                        SQL = "Update SOWSHPW1 set PAYMENT_TYPE = '" & Text2(1).Text & "' WHERE PAYMENT_TYPE <> 'Consignee Billed'"
                    Else
                        SQL = "Update SOWSHPW1 set PAYMENT_TYPE = '" & Text2(1).Text & "'"
                    End If
                End If
                AccD.Execute(SQL)
            Case "SERVICE_CODE"
                If SHIPPER_CODE = "FDX" Then
                    If Text1(Index).Text = "92" Then
                        Text1(2).Text = "1"
                        Call FM_Text1_Change(Text1(2))
                    End If
                End If
                If SHIPPER_CODE = "FDX" Then
                    SQL = "Update SOWSHPW1 set SERVICE_TYPE = '" & Text1(0).Text & "'"
                Else
                    SQL = "Update SOWSHPW1 set SERVICE_TYPE = '" & Text2(0).Text & "'"
                End If
                AccD.Execute(SQL)
            Case "PACKAGE_CODE"
                If SHIPPER_CODE = "FDX" Then
                    SQL = "Update SOWSHPW1 set PACKAGE_TYPE = '" & Text1(2).Text & "'"
                Else
                    SQL = "Update SOWSHPW1 set PACKAGE_TYPE = '" & Text2(2).Text & "'"
                End If
                AccD.Execute(SQL)
        End Select
        datSOWSHPW1.Refresh()

    End Sub

    Sub Create_DHL()
        Dim fn As String
        Dim Batch_File As Integer
        Dim Start_Pos As Integer
        Dim End_Pos As Integer
        Dim X As Integer
        Dim ctr As Double
        Dim CUST_NAME As String
        Dim ADDR1 As String
        Dim ADDR2 As String
    Dim BATCH_Line As String * 300

        fn = "G:\van\DHL\DHL_BATCH.txt"
        On Error Resume Next
        X = Shell("DELETE " & fn, 1)
        X = Shell("DELETE G:\van\DHL\DHL_IMPORT.csv", 1)
        Batch_File = FreeFile
        On Error GoTo 0
    Open fn For Output Access Write As #Batch_File ' Len = rlen

        Dim dynUATHEALS As OraDynaset
        Dim Elig_date As String
        Dim Grp_No As String

        Dim dynSOWSHPW1 As Recordset
        SQL = "Select * from SOWSHPW1"
        SQL = SQL & " Order By PICK_NO"
        dynSOWSHPW1 = AccD.OpenRecordset(SQL, dbOpenDynaset)

        Do While Not dynSOWSHPW1.EOF
            Call Prompt("Adding Member " & dynSOWSHPW1.Fields("CUST_ADDR_CODE").Value & "", "")
            BATCH_Line = Space$(300)
            'End_Pos = Len(cmbRecord(1).Text & ",")
            End_Pos = Len(dynSOWSHPW1.Fields("CART_NO").Value & "" & ",")
            Start_Pos = 1
            'Mid$(BATCH_Line, Start_Pos, End_Pos) = cmbRecord(1).Text & ","
            Mid$(BATCH_Line, Start_Pos, End_Pos) = dynSOWSHPW1.Fields("CART_NO").Value & "" & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(txtCode(0).Text & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = txtCode(0).Text & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(dynSOWSHPW1.Fields("CUST_ADDR_CODE").Value & "" & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = dynSOWSHPW1.Fields("CUST_ADDR_CODE").Value & "" & ","

            CUST_NAME = Replace(dynSOWSHPW1.Fields("CUST_NAME").Value, ",", "")
            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(CUST_NAME & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = CUST_NAME & ","

            ADDR1 = Replace(dynSOWSHPW1.Fields("CUST_ADDR1").Value, ",", "")
            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(ADDR1 & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = ADDR1 & ","

            ADDR2 = Replace(dynSOWSHPW1.Fields("CUST_ADDR2").Value & "", ",", "")
            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(ADDR2 & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = ADDR2 & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(dynSOWSHPW1.Fields("CUST_CITY").Value & "" & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = dynSOWSHPW1.Fields("CUST_CITY").Value & "" & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(dynSOWSHPW1.Fields("CUST_STATE").Value & "" & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = dynSOWSHPW1.Fields("CUST_STATE").Value & "" & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(dynSOWSHPW1.Fields("CUST_ZIP_CODE").Value & "" & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = dynSOWSHPW1.Fields("CUST_ZIP_CODE").Value & "" & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len("US,")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = "US," 'dynSOWSHPW1.Fields("COUNTRY").Value & ""

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(dynSOWSHPW1.Fields("PACKAGE_WEIGHT").Value & "" & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = dynSOWSHPW1.Fields("PACKAGE_WEIGHT").Value & "" & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(Text1(0).Text & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = Text1(0).Text & "," 'service type

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(Text1(1).Text & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = Text1(1).Text & ","  'Billing code

            If Text2(10).Text <> "" Then
                Start_Pos = End_Pos + 1
                End_Pos = End_Pos + Len(Text2(10).Text & ",")
                Mid$(BATCH_Line, Start_Pos, End_Pos) = Text2(10).Text & ","
            Else
                Start_Pos = End_Pos + 1
                End_Pos = End_Pos + Len(",")
                Mid$(BATCH_Line, Start_Pos, End_Pos) = "," 'Acct #
            End If

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len("1,")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = "1,"

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len(Format$(SSMonth1.Date, "MM/DD/YYYY") & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = Format$(SSMonth1.Date, "MM/DD/YYYY") & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len("Store No: " & dynSOWSHPW1.Fields("CUST_ADDR_CODE").Value & "" & ",")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = "Store No: " & dynSOWSHPW1.Fields("CUST_ADDR_CODE").Value & "" & ","

            Start_Pos = End_Pos + 1
            End_Pos = End_Pos + Len("Pick No: " & dynSOWSHPW1.Fields("PICK_NO").Value & "")
            Mid$(BATCH_Line, Start_Pos, End_Pos) = "Pick No: " & dynSOWSHPW1.Fields("PICK_NO").Value & ""

        Print #Batch_File, BATCH_Line
            dynSOWSHPW1.MoveNext()
        Loop
        dynSOWSHPW1.Close()

    Close #Batch_File
        FileCopy("G:\van\DHL\DHL_BATCH.txt", "G:\van\DHL\DHL_IMPORT.csv")
        Call Prompt("Now Printing Report", "")
    End Sub
#End Region
End Class