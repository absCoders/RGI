Imports System.Drawing
Imports System.Math

Public Class WHFWRTN1
    Private rowWHTWRTN1 As DataRow
    Private WH_RTN_NO As String
    Private rowICTWHSE1 As DataRow
    Private rowARTCUST1 As DataRow
    Dim Wh_Rtn_Status As String
    Dim BAR_CODE_FIRST_deleted As String
    Dim BAR_CODE_LAST_deleted As String
    Dim CUST_CODE_new As String
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")
        With dst
            ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_LOCATOR = '1' And WHSE_CODE = '" & ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & "'"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "WHTWRTN1", "*", , True)
            Create_TDA(.Tables.Add, "WHTWRTN2", "*", , True)
            .Tables("WHTWRTN2").Columns.Add("LPN_FIRST")
            .Tables("WHTWRTN2").Columns.Add("LPN_LAST")
            .Tables("WHTWRTN2").Columns.Add("STYLE_DESC")
            .Tables("WHTWRTN2").Columns.Add("COLOR_DESC")

            Create_TDA(.Tables.Add, "WHTWRTN3", "*", , True)
            '.Tables("WHTWRTN3").Columns.Add("UNITS", GetType(System.Int32))

            Create_Relation("WHTWRTN2", "WHTWRTN3", "WH_RTN_NO,WH_RTN_LNO")
            .Tables("WHTWRTN2").Columns.Add("CASES", GetType(System.Int32), "COUNT(CHILD(WHTWRTN2_WHTWRTN3).BAR_CODE)")
            .Tables("WHTWRTN2").Columns.Add("UNITS", GetType(System.Int32), "SUM(CHILD(WHTWRTN2_WHTWRTN3).QTY_RTN)")

            ASCMAIN1.sql = "Select * from WHTWRTN2"
            Create_TDA(.Tables.Add, "WHTWRTNS", "**", 0, False, "", 2)
            .Tables("WHTWRTNS").Columns.Add("STYLE_DESC")
            .Tables("WHTWRTNS").Columns.Add("COLOR_DESC")
            .Tables("WHTWRTNS").Columns.Add("CASES", GetType(System.Int32))
            .Tables("WHTWRTNS").Columns.Add("UNITS", GetType(System.Int32))

            ASCMAIN1.sql = "Select * from WHTWRTN1 Where WH_RTN_STATUS in ('S','C')"
            Create_TDA(.Tables.Add, "WHTWRTNX", "**", 0, False, "", 1)
            .Tables("WHTWRTNX").Columns.Add("CUSTOMER_NAME")

            ASCMAIN1.sql = "Select * from WHTBARC1 where TRAN_TYPE = 'L' and TRAN_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTBARC1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select * from WHTBARC0 where LOAD_NO in (Select Distinct LOAD_NO from WHTBARC1 where TRAN_TYPE = 'L' and TRAN_NO = :PARM1)"
            Create_TDA(.Tables.Add, "WHTBARC0", "**", 0, True, "V", 1)


        End With
        Fill_Records("ICTWHSE1")
        grdWHTRTRNX.DataSource = dst.Tables("WHTWRTNX")
        grdWHTWRTN2.DataSource = dst.Tables("WHTWRTN2")
        grdWHTWRTN3.DataSource = dst.Tables("WHTWRTN3")
        grdWHTWRTNS.DataSource = dst.Tables("WHTWRTNS")

        With grdWHTWRTN2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"CASES", "UNITS", "WH_RTN_NO", "WH_RTN_LNO", "STYLE_DESC", "COLOR_DESC"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"STYLE_CODE", "COLOR_CODE", "CTN_PACK_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                ElseIf New String() {"CASES", "UNITS"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With
        Sort_grdColumns(grdWHTWRTN2, "wh_rtn_lno")

        Create_Summary(grdWHTWRTNS, New String() {"CASES", "UNITS"})
        ASCMAIN1.Add_Value_List(grdWHTRTRNX, "WH_RTN_STATUS", , New String() {":", "S:SAVED", "C:COMPLETED", "F:FINALIZED"})
        ASCMAIN1.Add_Value_List(grdWHTRTRNX, "CUST_CODE", "Select CUST_CODE, CUST_NAME from ARTCUST1")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "New"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "Enter a Valid Customer Code"
                End If
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "Enter a Valid Whse Code"
                End If

            Case "Edit"
                WH_RTN_NO = Absx1.txtFor("WHSE_CODE").Text
                If Not ASCMAIN1.Logical_Lock("WHTWRTN1", WH_RTN_NO) Then
                    Exit Sub
                End If

            Case "Save", "Complete"
                If Absx1.dteFor("WH_RTN_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Return Date is Mandatory"
                End If

                If dst.Tables("WHTWRTN2").Rows.Count = 0 Then
                    EMsg &= vbCr & "No Returns Entered"
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
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
            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Save"
                Update_Record("S")
                Mode_Settings(False)

            Case "Complete"
                Update_Record("C")
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Save").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Complete").Settings.Enabled = iScreenMode

                    .Items("Cancel").Visible = ScreenMode And Not (EntryMode = "V")
                    .Items("Done").Visible = ScreenMode And (EntryMode = "V")
                    .Items("Save").Visible = IIf(Not ScreenMode, True, IIf(Wh_Rtn_Status = "C", False, True))
                    .Items("Complete").Visible = IIf(Not ScreenMode, True, IIf(Wh_Rtn_Status = "C", False, True))
                End With
            End With
        End If

        UltraExplorerBar1.Groups("Change Customer").Visible = Not ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        tab0.Visible = Not tf

        If ScreenMode Then
            Set_Read_Only(SplitContainer1, IIf(rowWHTWRTN1.Item("WH_RTN_STATUS") = "S", False, True))


            With grdWHTWRTN2.DisplayLayout.Override
                If EntryMode = "V" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                End If
            End With

        Else
            Clear_Record()
        End If


    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTWRTN1", "WHTWRTN2", "WHTWRTN3", "WHTBARC1", "WHTBARC0"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        optLPN_VIEW.Value = "A"
        Wh_Rtn_Status = ""

        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            If dst.Tables("ICTWHSE1").Rows.Count = 1 Then
                Absx1.txtFor("WHSE_CODE").Text = dst.Tables("ICTWHSE1").Rows(0).Item("WHSE_CODE")
            End If
        End If

        Fill_Records("WHTWRTNX")
        Sort_grdColumns(grdWHTRTRNX, "WH_RTN_NO")
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then

            WH_RTN_NO = ASCMAIN1.Next_Control_No("WHTWRTN1.WH_RTN_NO")
            Absx1.txtFor("WH_RTN_NO").Text = WH_RTN_NO

            rowWHTWRTN1 = dst.Tables("WHTWRTN1").NewRow
            rowWHTWRTN1.Item("WH_RTN_NO") = WH_RTN_NO
            rowWHTWRTN1.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
            rowWHTWRTN1.Item("WHSE_CODE") = Absx1.txtFor("WHSE_CODE").Text
            rowWHTWRTN1.Item("WH_RMA_NO") = ""
            rowWHTWRTN1.Item("WH_RTN_DATE") = DATETIME_STAMP.Date
            rowWHTWRTN1.Item("WH_RTN_COMMENT") = ""
            rowWHTWRTN1.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTWRTN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTWRTN1.Item("WH_RTN_STATUS") = "S"
            dst.Tables("WHTWRTN1").Rows.Add(rowWHTWRTN1)


            New_WHTBARC0()

        Else
            WH_RTN_NO = Absx1.txtFor("WH_RTN_NO").Text
            rowWHTWRTN1 = Fill_Record("WHTWRTN1", WH_RTN_NO)

            ASCMAIN1.sql = "Select * from  WHTWRTN2 Where WH_RTN_NO = '" & WH_RTN_NO & "'"
            Fill_Records("WHTWRTN2", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from  WHTWRTN3 Where WH_RTN_NO = '" & WH_RTN_NO & "'"
            Fill_Records("WHTWRTN3", "", True, ASCMAIN1.sql)

            Fill_Records("WHTBARC0", New Object() {WH_RTN_NO}, False)

            If dst.Tables("WHTBARC0").Select.Count = 0 Then
                New_WHTBARC0()
            End If
        End If

        Fill_Records("WHTBARC1", New Object() {WH_RTN_NO}, False)
        Wh_Rtn_Status = rowWHTWRTN1.Item("WH_RTN_STATUS")

        If dst.Tables("WHTBARC1").Rows.Count = 1 Then
            txtRMA_LPN.Text = dst.Tables("WHTBARC1").Rows(0)("BAR_CODE")
        End If

        rowICTWHSE1 = LookUp("ICTWHSE1", rowWHTWRTN1.Item("WHSE_CODE"))
        rowARTCUST1 = LookUp("ARTCUST1", rowWHTWRTN1.Item("CUST_CODE"))

        For Each rowWHTWRTN2 As DataRow In dst.Tables("WHTWRTN2").Select
            rowWHTWRTN2.Item("LPN_FIRST") = CStr((dst.Tables("WHTWRTN3").Compute("MIN(BAR_CODE)", "WH_RTN_LNO = " & rowWHTWRTN2.Item("WH_RTN_LNO")) & "")).PadLeft(8, "0")
            rowWHTWRTN2.Item("LPN_LAST") = CStr((dst.Tables("WHTWRTN3").Compute("MAX(BAR_CODE)", "WH_RTN_LNO = " & rowWHTWRTN2.Item("WH_RTN_LNO")) & "")).PadLeft(8, "0")

            LookUp("ICTSTYL1", rowWHTWRTN2.Item("STYLE_CODE") & "")
            If cdr Is Nothing Then
                rowWHTWRTN2.Item("STYLE_DESC") = ""
            Else
                rowWHTWRTN2.Item("STYLE_DESC") = cdr.Item("STYLE_DESC") & ""
            End If
            LookUp("ICTCOLR1", rowWHTWRTN2.Item("COLOR_CODE") & "")
            If cdr Is Nothing Then
                rowWHTWRTN2.Item("COLOR_DESC") = ""
            Else
                rowWHTWRTN2.Item("COLOR_DESC") = cdr.Item("COLOR_DESC") & ""
            End If
        Next
        Calc_LPN_Total()

        ASCMAIN1.Progress("")
    End Sub

    Sub New_WHTBARC0()
        Dim LOAD_NO As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")

        Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
        rowWHTBARC0.Item("LOAD_NO") = LOAD_NO
        rowWHTBARC0.Item("WHSE_CODE") = Absx1.txtFor("WHSE_CODE").Text
        rowWHTBARC0.Item("INIT_DATE") = DATETIME_STAMP
        rowWHTBARC0.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTBARC0.Item("LOAD_STATUS") = "R"
        rowWHTBARC0.Item("LOCATION_CODE") = Absx1.txtFor("WHSE_CODE").Text
        rowWHTBARC0.Item("TRAN_TYPE") = "L"
        rowWHTBARC0.Item("TRAN_NO") = WH_RTN_NO
        dst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
    End Sub

    Sub Update_Record(Update_Type As String)

        BeginTrans()
        Dim LOAD_NO As String = dst.Tables("WHTBARC0")(0).Item("LOAD_NO") & ""

        For Each rowWHTBARC1 As DataRow In dst.Tables("WHTBARC1").Select("")
            rowWHTBARC1.Item("TRAN_TYPE") = "L"
            rowWHTBARC1.Item("TRAN_NO") = WH_RTN_NO
            rowWHTBARC1.Item("PO_DATE_RECEIVED") = dteRETURN_DATE.Value
            rowWHTBARC1.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTBARC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTBARC1.Item("STATUS_CODE") = "R"
            rowWHTBARC1.Item("LOAD_NO") = LOAD_NO
        Next

        rowWHTWRTN1.Item("WH_RTN_STATUS") = Update_Type

        Update_Record_TDA("WHTWRTN1", "WH_RTN_NO = '" & WH_RTN_NO & "'")
        Update_Record_TDA("WHTWRTN2", "WH_RTN_NO = '" & WH_RTN_NO & "'")
        Update_Record_TDA("WHTWRTN3", "WH_RTN_NO = '" & WH_RTN_NO & "'")
        Update_Record_TDA("WHTBARC1", "TRAN_TYPE = 'L' and TRAN_NO = '" & WH_RTN_NO & "'")
        Update_Record_TDA("WHTBARC0", "TRAN_TYPE = 'L' and TRAN_NO = '" & WH_RTN_NO & "'")

        Dim Completed As String = False
        If Update_Type = "C" Then
            ASCDATA1.ExecuteSP("WHPLOCB2", "VVV", _
                   New Object() {"L", WH_RTN_NO, ASCMAIN1.SESSION_NO}, _
                   New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
        End If
        CommitTrans()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where &= " AND WHSE_CTN_CTL = 'C'"
        End Select
    End Sub

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        'If ScreenMode Then
        '    E.TABLE_NAME = "SOTRTRN1"
        '    E.COLUMN_NAME = "RTRN_NO"
        '    E.CODE_VALUE = Absx1.txtFor("RTRN_NO").Text
        '    E.DESC_VALUE = "Return"
        '    E.ATTACHMENT_NOTES = ""
        'End If

        Return E
    End Function

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTWRTN2, "B", "Location Inquiry")
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
            Case "grdSOTRTRN2"


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
            Case ""

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Location Inquiry"
                Dim Style_Code As String = "S:" & grd.ActiveRow.Cells("STYLE_CODE").Text
                Context_Launch("Select", Style_Code, e.Tool.Key, "WHFLOCS1", "F", "WHREC")

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("WHSE_CODE").Text <> "" Then
                        Click_Command("New", e)
                    End If
                End If

            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("CUST_CODE").Text <> "" Then
                        Click_Command("New", e)
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If Not InquiryMode And Absx1.txtFor("WHSE_CODE").Text <> "" Then
                    ' Click_Command("New")
                End If
            Case "WH_RTN_NO"
                'Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "REASON_CODE"


        End Select
    End Sub


#End Region

#Region "grdSOTRTRN2"

    Private Sub grdWHTWRTN2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTWRTN2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

                grdCodeDesc(grdWHTWRTN2, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")

                    ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
                    Dim rowICTSTYC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rowICTSTYC1s.Length = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = rowICTSTYC1s(0).Item("COLOR_CODE")
                    End If
                    If txtRMA_LPN.Text <> "" Then
                        e.Cell.Row.Cells("LPN_FIRST").Value = txtRMA_LPN.Text
                        e.Cell.Row.Cells("LPN_LAST").Value = txtRMA_LPN.Text
                    End If
                Else
                    grdWHTWRTN2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If
            Case "COLOR_CODE"
                grdCodeDesc(grdWHTWRTN2, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
                If cdr IsNot Nothing Then
                    e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")
                End If

        End Select
    End Sub

    Private Sub grdWHTWRTN2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTWRTN2.AfterRowActivate
        Dim Filter As String = ""

        If grdWHTWRTN2.ActiveRow.IsAddRow Then
            Filter = IIf(optLPN_VIEW.Value = "A", "", "WH_RTN_LNO = 0")

            grdWHTWRTN2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdWHTWRTN2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdWHTWRTN2.DisplayLayout.Bands(0).Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdWHTWRTN2.DisplayLayout.Bands(0).Columns("CTN_PACK_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            Filter = IIf(optLPN_VIEW.Value = "A", "", "WH_RTN_LNO = " & grdWHTWRTN2.ActiveRow.Cells("WH_RTN_LNO").Value)

            grdWHTWRTN2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdWHTWRTN2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdWHTWRTN2.DisplayLayout.Bands(0).Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdWHTWRTN2.DisplayLayout.Bands(0).Columns("CTN_PACK_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
        End If

        Dim dvw As DataView = DirectCast(grdWHTWRTN3.DataSource, DataTable).DefaultView
        dvw.RowFilter = (Filter)
    End Sub

    Private Sub optLPN_VIEW_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optLPN_VIEW.ValueChanged
        If ScreenMode = False Then Exit Sub
        If grdWHTWRTN2.ActiveRow Is Nothing Then Exit Sub
        Dim dvw As DataView = DirectCast(grdWHTWRTN3.DataSource, DataTable).DefaultView
        dvw.RowFilter = IIf(optLPN_VIEW.Value = "A", "", "WH_RTN_LNO = " & grdWHTWRTN2.ActiveRow.Cells("WH_RTN_LNO").Value)
    End Sub

    Private Sub grdWHTWRTN2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWHTWRTN2.AfterRowsDeleted
        Dim LOAD_NOs As New List(Of String)

        Dim BAR_CODE As String = BAR_CODE_FIRST_deleted
        Dim BAR_CODE2 As String = BAR_CODE_LAST_deleted
        Dim QTY As Int64 '= Val(BAR_CODE2) - Val(BAR_CODE) + 1
        If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
            QTY = Val(BAR_CODE2.Substring(1)) - Val(BAR_CODE.Substring(1)) + 1
        Else
            QTY = Val(BAR_CODE2) - Val(BAR_CODE) + 1
        End If
        'Dim QTY = Val(BAR_CODE_LAST_deleted) - Val(BAR_CODE_FIRST_deleted) + 1

        Dim BAR_CODE_first As Int64 '= Val(BAR_CODE)

        If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
            BAR_CODE_first = Val(BAR_CODE.Substring(1))
        Else
            BAR_CODE_first = Val(BAR_CODE)
        End If

        For i As Integer = 1 To QTY
            Dim BAR_CODE_T As String = Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
            If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
                BAR_CODE_T = BAR_CODE.ToUpper.Substring(0, 1) & Format(BAR_CODE_first + i - 1, "".PadLeft(7, "0"))
            Else
                BAR_CODE_T = Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
            End If
            If Not BAR_CODE_T = txtRMA_LPN.Text Then
                Dim rowWHTBARC1 As DataRow = dst.Tables("WHTBARC1").Rows.Find(BAR_CODE)
                If rowWHTBARC1 IsNot Nothing Then
                    Dim LOAD_NO As String = rowWHTBARC1.Item("LOAD_NO") & ""
                    If Not LOAD_NOs.Contains(LOAD_NO) Then
                        LOAD_NOs.Add(LOAD_NO)
                    End If
                    rowWHTBARC1.Delete()
                End If
            End If

        Next

        If dst.Tables("WHTWRTN2").Select.Count = 0 Then
            For Each LOAD_NO As String In LOAD_NOs
                Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").Rows.Find(LOAD_NO)
                'rowWHTBARC0.Delete()
            Next
        End If
        Calc_LPN_Total()
        BAR_CODE_FIRST_deleted = ""
        BAR_CODE_LAST_deleted = ""
    End Sub

    Private Sub grdWHTWRTN2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTWRTN2.AfterRowUpdate
        Write_LPNs()
        Calc_LPN_Total()
    End Sub

    Private Sub grdWHTWRTN2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdWHTWRTN2.BeforeRowsDeleted
        BAR_CODE_FIRST_deleted = ""
        BAR_CODE_LAST_deleted = ""
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            BAR_CODE_FIRST_deleted = grow.Cells("LPN_FIRST").Value
            BAR_CODE_LAST_deleted = grow.Cells("LPN_LAST").Value
        Next
    End Sub

    Private Sub grdWHTWRTN2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTWRTN2.BeforeRowUpdate
        With grdWHTWRTN2
            If Val(e.Row.Cells("CTN_PACK_QTY").Text & "") = 0 Then
                MsgBox("Missing Ctn Pack Qty", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If
            If e.Row.Cells("STYLE_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Row.Cells("COLOR_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTCOLR1", e.Row.Cells("COLOR_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Color Code (" & e.Row.Cells("COLOR_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
                If Not e.Cancel Then
                    LookUp("ICTSTYC1", New String() {e.Row.Cells("STYLE_CODE").Text, e.Row.Cells("COLOR_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Color Code (" & e.Row.Cells("COLOR_CODE").Text & ") not set up for Style (" & e.Row.Cells("STYLE_CODE").Text & ")",
                               MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If
            If txtRMA_LPN.Text <> "" Then
                If e.Row.Cells("LPN_FIRST").Value & "" <> txtRMA_LPN.Text Or e.Row.Cells("LPN_LAST").Value & "" <> txtRMA_LPN.Text Then
                    MsgBox("Cannot Enter Individual LPN's for RMA_LPN option",
                                 MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If
            If e.Row.Cells("LPN_FIRST").Value & "" = "" Or e.Row.Cells("LPN_LAST").Value & "" = "" Then
                MsgBox("Please Enter Starting and Ending LPN", _
                             MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                e.Row.Cells("LPN_FIRST").Value = Check_BAR_CODE(e.Row.Cells("LPN_FIRST").Value)
                Dim Bar_Code As String = e.Row.Cells("LPN_FIRST").Value
                If Bar_Code <> txtRMA_LPN.Text Then
                    LookUp("WHTBARC1", Bar_Code)
                    If cdr IsNot Nothing Then
                        MsgBox("Starting LPN: " & e.Row.Cells("LPN_FIRST").Value & " already exists in Database" & vbCrLf _
                        & "Received on PO " & cdr.Item("PO_ORDER_NO") _
                        & " on " & cdr.Item("PO_DATE_RECEIVED"),
                        MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                        Bar_Code = ""
                        e.Cancel = True
                    Else
                        If dst.Tables("WHTWRTN3").Select("BAR_CODE = '" & Bar_Code & "'").Length <> 0 Then
                            MsgBox("Starting LPN: " & e.Row.Cells("LPN_FIRST").Value & " already exists in Current Receipt" & vbCrLf & vbCrLf,
                            MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                            Bar_Code = ""
                            e.Cancel = True
                        End If
                    End If
                End If
                e.Row.Cells("LPN_LAST").Value = Check_BAR_CODE(e.Row.Cells("LPN_LAST").Value)
                Dim Bar_Code2 As String = e.Row.Cells("LPN_LAST").Value
                If Bar_Code2 <> txtRMA_LPN.Text Then
                    LookUp("WHTBARC1", Bar_Code2)
                    If cdr IsNot Nothing Then
                        MsgBox("Ending LPN: " & e.Row.Cells("LPN_LAST").Value & " already exists in Database" & vbCrLf _
                        & "Received on PO " & cdr.Item("PO_ORDER_NO") _
                        & " on " & cdr.Item("PO_DATE_RECEIVED"),
                        MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                        Bar_Code = ""
                        e.Cancel = True
                    Else
                        If dst.Tables("WHTWRTN3").Select("BAR_CODE = '" & Bar_Code2 & "'").Length <> 0 Then
                            MsgBox("Ending LPN: " & e.Row.Cells("LPN_LAST").Value & " already exists in Current Receipt" & vbCrLf & vbCrLf,
                            MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                            Bar_Code = ""
                            e.Cancel = True
                        End If
                    End If
                End If
            End If

        End With
        If e.Cancel <> True Then
            If Val(e.Row.Cells("WH_RTN_LNO").Value & "") = 0 Then
                e.Row.Cells("WH_RTN_LNO").Value = Val(dst.Tables("WHTWRTN2").Compute("MAX(WH_RTN_LNO)", "") & "") + 1
            End If
            e.Row.Cells("WH_RTN_NO").Value = WH_RTN_NO
        End If

    End Sub
    Sub Calc_LPN_Total()
        Dim Lno As Integer = 0
        dst.Tables("WHTWRTNS").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("WHTWRTN2").Select(""), "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC").Rows
            Lno += 1
            Dim rowWHTWRTNS As DataRow = dst.Tables("WHTWRTNS").NewRow
            rowWHTWRTNS.Item("WH_RTN_NO") = WH_RTN_NO
            rowWHTWRTNS.Item("WH_RTN_LNO") = Lno
            rowWHTWRTNS.Item("STYLE_CODE") = row.Item("STYLE_CODE")
            rowWHTWRTNS.Item("STYLE_DESC") = row.Item("STYLE_DESC")
            rowWHTWRTNS.Item("COLOR_CODE") = row.Item("COLOR_CODE")
            rowWHTWRTNS.Item("COLOR_DESC") = row.Item("COLOR_DESC")
            rowWHTWRTNS.Item("CASES") = Val(dst.Tables("WHTWRTN2").Compute("SUM(CASES)", "STYLE_CODE  ='" & row.Item("STYLE_CODE") & "' And COLOR_CODE = '" & row.Item("COLOR_CODE") & "'") & "")
            rowWHTWRTNS.Item("UNITS") = Val(dst.Tables("WHTWRTN2").Compute("SUM(UNITS)", "STYLE_CODE  ='" & row.Item("STYLE_CODE") & "' And COLOR_CODE = '" & row.Item("COLOR_CODE") & "'") & "")

            dst.Tables("WHTWRTNS").Rows.Add(rowWHTWRTNS)
        Next

    End Sub

    Private Sub grdWHTWRTN2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTWRTN2.ClickCellButton

        If grdWHTWRTN2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
            Case "COLOR_CODE"
                sql_where = "COLOR_CODE in (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & e.Cell.Row.Cells("STYLE_CODE").Value & "')"
        End Select
        grdClickCellButton(grdWHTWRTN2, sql_where, False)

    End Sub

    Function Check_BAR_CODE(BAR_CODE As String) As String

        Dim prefix As String = ""
        If BAR_CODE = "" Then Return BAR_CODE

        If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
            prefix = BAR_CODE.ToUpper.Substring(0, 1)
            BAR_CODE = BAR_CODE.Substring(1)
        End If

        If BAR_CODE.PadLeft(8, "0") <> Format(Val(BAR_CODE), "".PadLeft(8, "0")) Then
            BAR_CODE = ""
        Else
            If prefix = "" Then
                BAR_CODE = BAR_CODE.PadLeft(8, "0")
            Else
                BAR_CODE = prefix & BAR_CODE.PadLeft(7, "0")
            End If
        End If
        Return BAR_CODE

    End Function

    Sub Validate_BAR_CODE(Barc_Code As String)
        If Not ScreenMode Then Exit Sub
        Dim BAR_CODE As String = Check_BAR_CODE(Barc_Code)

        If BAR_CODE <> "" Then
            LookUp("WHTBARC1", BAR_CODE)
            If cdr IsNot Nothing Then
                MsgBox("LPN already exists in Database" & vbCrLf _
                    & "Received on PO " & cdr.Item("PO_ORDER_NO") _
                    & " on " & cdr.Item("PO_DATE_RECEIVED"), _
                    MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                BAR_CODE = ""
            Else
                Dim rowWHTBARC1 As DataRow = dst.Tables("WHTBARC1").Rows.Find(BAR_CODE)
                If rowWHTBARC1 IsNot Nothing Then
                    Dim rowPOTSHIPX As DataRow = dst.Tables("POTSHIPX").Rows.Find(New Object() {rowWHTBARC1.Item("PO_SHIPMENT_NO"), rowWHTBARC1.Item("PO_SHIPMENT_LNO")})
                    MsgBox("LPN already exists in Current Receipt" & vbCrLf & vbCrLf _
                        & "Received on Container " & rowPOTSHIPX.Item("CONTAINER_NO") _
                        & " Shipment Line " & rowWHTBARC1.Item("PO_SHIPMENT_LNO") _
                        & " as part of Carton Type " & rowWHTBARC1.Item("CARTON_NO"), _
                        MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                    BAR_CODE = ""
                End If
            End If

            'txtBAR_CODE.Text = BAR_CODE
            'If BAR_CODE = "" Then
            '    txtBAR_CODE.Focus()
            'Else
            '    txtBAR_CODE2.Focus()
            'End If
        End If
    End Sub
#End Region

    Private Sub grdWHTRTRNX_DoubleClick(sender As Object, e As System.EventArgs) Handles grdWHTRTRNX.DoubleClick
        WH_RTN_NO = grdWHTRTRNX.ActiveRow.Cells("WH_RTN_NO").Value
        Absx1.txtFor("WH_RTN_NO").Text = WH_RTN_NO

        Wh_Rtn_Status = grdWHTRTRNX.ActiveRow.Cells("WH_RTN_STATUS").Value
        If Wh_Rtn_Status = "S" Then
            Click_Command("Edit")
        Else
            Click_Command("View")
        End If

    End Sub

    Sub Write_LPNs()


        dst.Tables("WHTWRTN3").Rows.Clear()
        dst.Tables("WHTBARC1").Rows.Clear()
        For Each rowWHTWRTN2 As DataRow In dst.Tables("WHTWRTN2").Select
            Dim BAR_CODE As String = rowWHTWRTN2.Item("LPN_FIRST")
            Dim BAR_CODE2 As String = rowWHTWRTN2.Item("LPN_LAST")
            Dim QTY As Int64 '= Val(BAR_CODE2) - Val(BAR_CODE) + 1
            If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
                QTY = Val(BAR_CODE2.Substring(1)) - Val(BAR_CODE.Substring(1)) + 1
            Else
                QTY = Val(BAR_CODE2) - Val(BAR_CODE) + 1
            End If

            Dim BAR_CODE_first As Int64 '= Val(BAR_CODE)

            If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
                BAR_CODE_first = Val(BAR_CODE.Substring(1))
            Else
                BAR_CODE_first = Val(BAR_CODE)
            End If

            For i As Integer = 1 To QTY
                Dim BAR_CODE_T As String
                If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
                    BAR_CODE_T = BAR_CODE.ToUpper.Substring(0, 1) & Format(BAR_CODE_first + i - 1, "".PadLeft(7, "0"))
                Else
                    BAR_CODE_T = Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
                End If

                Dim rowWHTWRTN3 As DataRow = dst.Tables("WHTWRTN3").NewRow
                rowWHTWRTN3.Item("WH_RTN_NO") = WH_RTN_NO
                rowWHTWRTN3.Item("WH_RTN_LNO") = rowWHTWRTN2.Item("WH_RTN_LNO")
                rowWHTWRTN3.Item("BAR_CODE") = BAR_CODE_T 'Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
                rowWHTWRTN3.Item("QTY_RTN") = rowWHTWRTN2.Item("CTN_PACK_QTY")
                dst.Tables("WHTWRTN3").Rows.Add(rowWHTWRTN3)

                If txtRMA_LPN.Text = "" Or dst.Tables("WHTBARC1").Rows.Count = 0 Then
                    Dim rowWHTBARC1 As DataRow = dst.Tables("WHTBARC1").NewRow

                    rowWHTBARC1.Item("BAR_CODE") = BAR_CODE_T 'Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
                    rowWHTBARC1.Item("LOAD_NO") = dst.Tables("WHTBARC0")(0).Item("LOAD_NO") & ""
                    dst.Tables("WHTBARC1").Rows.Add(rowWHTBARC1)
                End If
            Next
        Next


        Sort_grdColumns(grdWHTWRTN3, "BAR_CODE".ToLower)
    End Sub

    Private Sub cmdChangeCustomer_Click(sender As Object, e As EventArgs) Handles cmdChangeCustomer.Click
        If grdWHTRTRNX.Selected.Rows.Count = 0 Then
            MsgBox("No Returns Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        Else

            For Each grow As UltraWinGrid.UltraGridRow In grdWHTRTRNX.Selected.Rows
                Dim WH_RTN_NO As String = grow.Cells("WH_RTN_NO").Value
                Dim WH_RTN_STATUS As String = grow.Cells("WH_RTN_STATUS").Value

                If Not ASCMAIN1.Logical_Lock("WHTWRTN1", WH_RTN_NO) Then
                    Exit Sub
                End If

                If WH_RTN_STATUS <> "C" And WH_RTN_STATUS <> "S" Then
                    MsgBox("You may correct customers only if the Status of the Returns record is 'Completed' or 'Saved' - see " & WH_RTN_NO, _
                           MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If
 
            Next


            Dim CUST_CODE_NEW As String = Absx1.txtFor("CUST_CODE_NEW").Text
            If CUST_CODE_NEW = "" Then
                MsgBox("No New Customer Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Else
                Dim row As DataRow = LookUp("ARTCUST1", CUST_CODE_NEW)
                If row Is Nothing Then
                    MsgBox("Customer Selected is invalid", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Else
                    If MsgBox("Change the Customer for the " & grdWHTRTRNX.Selected.Rows.Count & " returns selected to " & CUST_CODE_NEW & "?", vbYesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    Else
                        For Each grow As UltraWinGrid.UltraGridRow In grdWHTRTRNX.Selected.Rows
                            Dim WH_RTN_NO As String = grow.Cells("WH_RTN_NO").Value
                            grow.Cells("CUST_CODE").Value = CUST_CODE_NEW
                            grow.Cells("CUSTOMER_NAME").Value = row.Item("CUST_NAME") & ""
                            grow.Update()
                            ASCMAIN1.sql = "Update WHTWRTN1 set CUST_CODE = :PARM1 where WH_RTN_NO = :PARM2"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {CUST_CODE_NEW, WH_RTN_NO})
                            Dim rowWHTWRTN1 As DataRow = Fill_Record("WHTWRTN1", WH_RTN_NO)
                            rowWHTWRTN1.Item("CUST_CODE") = CUST_CODE_NEW
                            Write_Audit_Trail(rowWHTWRTN1, "E")
                        Next
                        dst.Tables("WHTWRTN1").Rows.Clear()
                        MsgBox("Customer(s) have been corrected for the selected Returns", MsgBoxStyle.OkOnly, "Success")
                        grdWHTRTRNX.Selected.Rows.Clear()
                        Absx1.txtFor("CUST_CODE_NEW").Text = ""
                    End If
                End If
            End If
        End If

        ASCMAIN1.MultiTask_Release()

    End Sub
End Class