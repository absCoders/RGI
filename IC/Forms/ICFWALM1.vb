Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

Public Class ICFWALM1

    Dim PROGRAM_CODE As String
    Dim rowICTWALM1 As DataRow

    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow


    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing

    Dim rangeCopyFrom As SpreadsheetGear.IRange
    Dim rangePaste_To As SpreadsheetGear.IRange

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Get_PARM("ICTPARM1")

        With dst

            ASCMAIN1.sql = "Select ICTWALM1.*" & vbCrLf _
                & " from ICTWALM1" & vbCrLf _
                & " where CUST_CODE = :PARM1 and ORDR_DATE = :PARM2 and WHSE_CODE = :PARM3"
            Create_TDA(.Tables.Add, "ICTWALM1", "**", 0, False, "VDV", 1)

            Create_TDA(.Tables.Add, "ICTWALM2", "*", 1)
            .Tables("ICTWALM2").Columns.Add("ITM_SEQ", GetType(System.Int32))

            Create_TDA(.Tables.Add, "ICTWALM3", "*", 1)
            .Tables("ICTWALM3").Columns.Add("STR_SEQ", GetType(System.Int32))

            With .Tables.Add("ICTWALMQ")
                .Columns.Add("CUST_SKU")
                .Columns.Add("CUST_STORE_NO")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("ITM_SEQ", GetType(System.Int32))
                .Columns.Add("STR_SEQ", GetType(System.Int32))
                .Columns.Add("PACK")
                .PrimaryKey = New DataColumn() { .Columns("CUST_SKU"), .Columns("CUST_STORE_NO")}
            End With

            With .Tables.Add("ICTWALMS")
                .Columns.Add("CUST_SKU")
                .Columns.Add("CUST_STORE_NO")
                .Columns.Add("QTY", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("CUST_SKU"), .Columns("CUST_STORE_NO")}
            End With


            Create_TDA(.Tables.Add, "ICTVOLG1", "*", 0, False)
            .Tables("ICTVOLG1").Columns.Add("VOL_GRADE") ', GetType(System.String), "VOL_GRADE_RANK")
            .Tables("ICTVOLG1").Columns.Add("SEL")
            .Tables("ICTVOLG1").Columns("SEL").DefaultValue = "0"
        End With

        Fill_Records("ICTVOLG1")
        For Each row As DataRow In dst.Tables("ICTVOLG1").Select
            row.Item("SEL") = "1"
            row.Item("VOL_GRADE") = Format(Val(row.Item("VOL_GRADE_RANK") & ""), "000")
        Next

        grdICTVOLG1.DataSource = dst.Tables("ICTVOLG1")
        Sort_grdColumns(grdICTVOLG1, "VOL_GRADE")

        grdICTWALM1.DataSource = dst.Tables("ICTWALM1")
        grdICTWALM2.DataSource = dst.Tables("ICTWALM2")
        grdICTWALM3.DataSource = dst.Tables("ICTWALM3")

        Create_Summary(grdICTWALM1, "PROGRAM_CODE", "Count")

        'grdICTWALM1.DisplayLayout.Bands(0).Columns("SELECTED").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        'With grdICTWALM1.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"SELECTED", "ORDR_GROUP_NO", "ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT"}
        '        .Columns(COLUMN_NAME).Header.Fixed = True
        '    Next
        '    For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
        '        If GCOL.Key = "SELECTED" Then
        '            GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
        '        Else
        '            GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
        '            GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
        '        End If
        '        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
        '        GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        If GCOL.Key.StartsWith("ORDR_AMT") Then
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
        '            GCOL.Width = 80
        '            GCOL.Format = "#,##0"
        '            Create_Summary(grdICTWALM1, GCOL.Key)
        '        ElseIf GCOL.Key.StartsWith("ORDR_QTY") Then
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
        '            GCOL.Width = 70
        '            GCOL.Format = "#,##0"
        '            Create_Summary(grdICTWALM1, GCOL.Key)
        '        ElseIf GCOL.Key.StartsWith("ORDR_CNT") Then
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
        '            GCOL.Width = 50
        '            GCOL.Format = "#,##0"
        '            Create_Summary(grdICTWALM1, GCOL.Key)
        '        ElseIf New String() {"ORDR_GROUP_NO", "ORDR_CUST_PO"}.Contains(GCOL.Key) Then
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightCoral
        '            GCOL.Width = 110
        '        ElseIf New String() {"CUST_DC_NO", "ORDR_DEPT"}.Contains(GCOL.Key) Then
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightCoral
        '            GCOL.Width = 70
        '        ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
        '        ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(GCOL.Key) Then
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
        '            GCOL.Width = 90
        '        Else
        '            GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
        '        End If
        '    Next
        'End With

        'ASCMAIN1.Add_Value_List(grdICTSTYLX, "ITEM_STATUS")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                If Absx1.txtFor("PROGRAM_NAME").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Program Name"
                End If

                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        'ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                        'If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                        '    EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                        'End If
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If

                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 IsNot Nothing Then
                        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Warehouse " & Absx1.txtFor("WHSE_CODE").Text
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTWALM1", CUST_CODE) Then Exit Sub
                End If

            Case "View", "Edit"

                CUST_CODE = Absx1.txtFor("CUST_CODE").Text

                If CUST_CODE = "" Then
                    EMsg &= vbCr & "No Customer Defined"
                Else
                    If grdICTWALM1.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Orders in Selection Grid"
                    Else
                        Dim rows() As DataRow = dst.Tables("SOTORDR0").Select("SELECTED='1'")
                        If rows.Length = 0 Then
                            EMsg &= vbCr & "No Orders Selected"
                        Else
                            If rows(0).Item("CUST_CODE") <> CUST_CODE Then
                                EMsg &= vbCr & "Orders in Selection grid do not appear to belong to Customer Defined"
                            End If
                        End If
                    End If
                End If

                If EMsg <> "" Then
                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Code"
                    End If
                    rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Warehouse Code"
                    End If
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        For Each row As DataRow In dst.Tables("SOTORDR0").Select("SELECTED = '1'")
                            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        Next
                        If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("Are you sure you want to Cancel your changes?",
                            MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If dst.Tables("SOTCANCY").Select("ORDR_QTY_OPEN <> ORIG_QTY_OPEN").Length = 0 Then
                    EMsg &= vbCr & "No records have been updated"
                End If
        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View", "Edit"
                MyBase.EntryMode = Mid(eItemKey, 1, 1)
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Cancel"
                Me.Mode_Settings(False)

            Case "Update"
                Me.Update_Record()
                Me.Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    If Not ScreenMode Or (EntryMode = "V") Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode

                    .Items("Cancel").Visible = (EntryMode = "N" Or EntryMode = "E")
                    .Items("Update").Visible = (EntryMode = "N" Or EntryMode = "E")

                    '.Items("Print").Settings.Enabled = iScreenMode
                End With

                .Groups("Display Options").Visible = False
                .Groups("Commands").Visible = ScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        xlsPlan.Visible = ScreenMode
        grdICTWALM1.Visible = Not ScreenMode

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid _
                In New UltraWinGrid.UltraGrid() _
                {grdICTWALM1}
                With grd.DisplayLayout.Override
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        If grd.Name = "grdSOTCANCX" Or grd.Name = "grdICTSTYLX" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        End If
                        '  .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.True
                        .AllowDelete = DefaultableBoolean.False
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    End If
                End With
            Next
        End If

        'With grdSOTORDR0.DisplayLayout.Override
        '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .AllowUpdate = DefaultableBoolean.True
        '    .AllowDelete = DefaultableBoolean.False
        'End With

        'With grdICTSTYLX.DisplayLayout.Override
        '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .AllowUpdate = DefaultableBoolean.True
        '    .AllowDelete = DefaultableBoolean.False
        'End With

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ICTWALM1", "ICTWALM2", "ICTWALM3", "ICTWALMQ"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        MyBase.EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE

        If ASCMAIN1.Running_in_VS Then
            Absx1.txtFor("CUST_CODE").Text = "WALMART"
            Absx1.txtFor("WHSE_CODE").Text = "NJC"
            Absx1.dteFor("ORDR_DATE").Text = Now.Date
        End If

        Clear_All_Filters(grdICTWALM1)
    End Sub

    Private Sub Load_Record()


        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            PROGRAM_CODE = ASCMAIN1.Next_Control_No("ICTWALM1.PROGRAM_CODE")

            rowICTWALM1 = dst.Tables("ICTWALM1").NewRow
            With rowICTWALM1
                .Item("PROGRAM_CODE") = PROGRAM_CODE
                .Item("CUST_CODE") = CUST_CODE
                '  .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                '   .Item("RSRV_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                .Item("WHSE_CODE") = WHSE_CODE
                '  .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
            End With
            dst.Tables("ICTWALM1").Rows.Add(rowICTWALM1)

        Else
            rowICTWALM1 = Fill_Record("ICTWALM1", PROGRAM_CODE)
        End If

        CUST_CODE = rowICTWALM1.Item("CUST_CODE")
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)


        '  lblINIT_DATE.Text = "Entered on " & Format(rowSOTRSRV1.Item("INIT_DATE"), "MM/dd/yyyy")

        'If EntryMode = "N" Then
        '    lblStatus.Text = "New Order"
        'Else
        '    Select Case rowSOTRSRV1.Item("RSRV_STATUS")
        '        Case "O"
        '            lblStatus.Text = "Open"
        '        Case "C"
        '            lblStatus.Text = "Cancelled"
        '        Case "D"
        '            lblStatus.Text = "Deleted"
        '    End Select
        'End If

        xlsPlan.GetLock()
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & "A.XLS"
        xlsPlan.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        xlsPlan.ReleaseLock()


        xlsSSO.GetLock()
        Dim FILENAME_SSO As String = ASCMAIN1.Folders("Work") & "B.XLSm"
        xlsSSO.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME_SSO)
        xlsSSO.ReleaseLock()

        '    Display_Totals()
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()
            Update_Record_TDA("ICTWALM1")
            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWALM1, "S", "Show Filter")
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

            Case "grdICTSTYLX"

            Case "grdSOTORDR0"


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else

            Select Case e.SourceControl.Name
                Case "grdSOTALLOX", "grdICTITEM1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.Key <> "Show All Levels" Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

        End Select

        If grd Is Nothing OrElse (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow) Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode Then
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        Set_SOTORDR0()
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)

            Case "CUST_CODE"
                Set_SOTORDR0()

        End Select
    End Sub

#End Region

    Sub Set_SOTORDR0()

        If ScreenMode Then Exit Sub
        Fill_Records("SOTORDR0", New Object() {Absx1.txtFor("CUST_CODE").Text, dteORDR_DATE.Value, Absx1.txtFor("WHSE_CODE").Text})
        grdICTWALM1.Text = "Order Groups for " & Absx1.txtFor("CUST_CODE").Text & " with Order Date of " & dteORDR_DATE.Value
    End Sub

    Sub Update_Totals()

    End Sub

    Private Sub cmdLoadStyles_Click(sender As Object, e As EventArgs) Handles cmdLoadStyles.Click
        xlsPlan.GetLock()

        workbook = xlsPlan.ActiveWorkbook
        worksheet = workbook.Worksheets("IS")

        Dim RX As Integer = 8
        Dim C0 As Integer = 12
        For S As Integer = 1 To 48
            Dim rowICTWALM2 As DataRow = dst.Tables("ICTWALM2").NewRow
            With rowICTWALM2
                .Item("PROGRAM_CODE") = PROGRAM_CODE
                .Item("STYLE_CODE") = worksheet.Cells(RX + 0, C0 + S - 1).Value
                .Item("COLOR_CODE") = "AST" '  worksheet.Cells(RX, C0 + S - 1).Value
                .Item("STYLE_DESC") = worksheet.Cells(RX + 4, C0 + S - 1).Value
                .Item("CUST_STYLE_CODE") = worksheet.Cells(RX + 3, C0 + S - 1).Value
                .Item("CUST_COLOR_CODE") = worksheet.Cells(RX + 5, C0 + S - 1).Value
                .Item("CUST_SIZE_CODE") = worksheet.Cells(RX + 1, C0 + S - 1).Value
                ' .Item("CUST_UPC") = worksheet.Cells(RX, C0 + S - 1).Value
                .Item("CUST_SKU") = worksheet.Cells(RX + 2, C0 + S - 1).Value
                .Item("STYLE_RETAIL") = worksheet.Cells(RX + 7, C0 + S - 1).Value
                .Item("STYLE_PRICE") = worksheet.Cells(RX + 6, C0 + S - 1).Value

                .Item("PACK_A") = worksheet.Cells(RX - 5, C0 + S - 1).Value
                .Item("PACK_B") = worksheet.Cells(RX - 4, C0 + S - 1).Value
                .Item("PACK_C") = worksheet.Cells(RX - 3, C0 + S - 1).Value

                .Item("ITM_SEQ") = S
            End With
            dst.Tables("ICTWALM2").Rows.Add(rowICTWALM2)
        Next

        xlsPlan.ReleaseLock()

        Sort_grdColumns(grdICTWALM2, "STYLE_CODE,COLOR_CODE")

    End Sub

    Private Sub cmdLoadStores_Click(sender As Object, e As EventArgs) Handles cmdLoadStores.Click
        xlsPlan.GetLock()

        workbook = xlsPlan.ActiveWorkbook
        worksheet = workbook.Worksheets("IS")

        Dim RX As Integer = 20
        Dim C0 As Integer = 12
        Dim r As Integer = 1

        Do While worksheet.Cells(RX + r, 0).Value & "" <> ""
            Dim rowICTWALM3 As DataRow = dst.Tables("ICTWALM3").NewRow
            With rowICTWALM3
                .Item("PROGRAM_CODE") = PROGRAM_CODE
                .Item("CUST_STORE_NO") = worksheet.Cells(RX + r, 0).Value
                .Item("CUST_STORE_NAME") = worksheet.Cells(RX + r, 1).Value
                .Item("CUST_STORE_VOL_GRADE") = worksheet.Cells(RX + r, 2).Value
                .Item("CUST_STORE_VOL_RANK") = worksheet.Cells(RX + r, 3).Value
                .Item("STR_SEQ") = r
            End With
            dst.Tables("ICTWALM3").Rows.Add(rowICTWALM3)
            r += 1
        Loop

        xlsPlan.ReleaseLock()

        Sort_grdColumns(grdICTWALM3, "CUST_STORE_VOL_RANK".ToLower)
    End Sub

    Private Sub cmdLoadSSO_Click(sender As Object, e As EventArgs) Handles cmdLoadSSO.Click

        dst.Tables("ICTWALMQ").Rows.Clear()

        Dim P As Integer = 0
        For Each PACK As String In New String() {"A", "B", "C"}
            P += 1

            xlsPlan.GetLock()

            workbook = xlsPlan.ActiveWorkbook
            worksheet = workbook.Worksheets("IS")

            Dim RX As Integer = 20
            Dim C0 As Integer = 12
            Dim r As Integer = 1

            Do While worksheet.Cells(RX + r, 0).Value & "" <> ""

                Dim Q As Integer = Val(worksheet.Cells(RX + r, 3 + P).Value & "")
                If Q <> 0 Then

                    For Each row As DataRow In dst.Tables("ICTWALM2").Select("PACK_" & PACK & " <> 0")
                        Dim CUST_SKU As String = row.Item("CUST_SKU")
                        Dim PQ As Integer = Val(row.Item("PACK_" & PACK) & "")
                        Dim rowICTWALMQ As DataRow = dst.Tables("ICTWALMQ").NewRow
                        With rowICTWALMQ
                            .Item("CUST_SKU") = CUST_SKU
                            .Item("CUST_STORE_NO") = worksheet.Cells(RX + r, 0).Value
                            .Item("QTY") = Q * PQ
                            .Item("ITM_SEQ") = row.Item("ITM_SEQ")
                            .Item("STR_SEQ") = r
                            .Item("PACK") = PACK
                        End With
                        dst.Tables("ICTWALMQ").Rows.Add(rowICTWALMQ)
                    Next
                End If
                r += 1
            Loop

            xlsPlan.ReleaseLock()
        Next

        xlsSSO.GetLock()

        workbook = xlsSSO.ActiveWorkbook
        worksheet = workbook.Worksheets("SSO Form")

        range = worksheet.Cells("E2")
        Dim v As DataView = dst.Tables("ICTWALMQ").DefaultView
        v.Sort = "PACK,ITM_SEQ,STR_SEQ"
        Dim TBL As DataTable = v.ToTable

        range.CopyFromDataTable(TBL, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)


        xlsSSO.ReleaseLock()

    End Sub
End Class