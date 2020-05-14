Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

Public Class SOFWALMP

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Get_PARM("ICTPARM1")

        With dst

            ASCMAIN1.sql = "Select SOTORDR0.*" & vbCrLf _
                & " from SOTORDR0" & vbCrLf _
                & " where CUST_CODE = 'WALMART' and ORDR_CUST_PO LIKE '%' || :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "SOTORDR0", "*", 0, False, "", 1)
            .Tables("SOTORDR0").Columns.Add("DC")
            .Tables("SOTORDR0").Columns.Add("NO", GetType(System.Int32))


            With .Tables.Add("SOTORDRP")
                .Columns.Add("NO", GetType(System.Int32))
                .Columns.Add("PO")
                .Columns.Add("DC")
                .Columns.Add("ORDR_GROUP_NO")
                .Columns.Add("GROUPS")
                .Columns.Add("STYLE_CODE_1")
                .Columns.Add("STYLE_CODE_2")
                .Columns.Add("ORDR_QTY", GetType(System.Int64))
                .Columns.Add("ORDR_QTY_OPEN", GetType(System.Int64))
                .Columns.Add("ORDR_QTY_PICK", GetType(System.Int64))
                .Columns.Add("ORDR_QTY_SHIP", GetType(System.Int64))
                .Columns.Add("ORDR_QTY_CANC", GetType(System.Int64))
                .PrimaryKey = New DataColumn() { .Columns("NO")}
            End With

            '  Create_Relation("SOTORDRP", "SOTORDR0", "NO")
        End With

        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTORDRP.DataSource = dst.Tables("SOTORDRP")

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        'Create_Summary(grdSOTORDR0, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})

        Create_Summary(grdSOTORDRP, "PO", "Count")
        'Create_Summary(grdSOTORDRP, New String() {"GROUPS", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTORDR0, grdSOTORDRP}
            With grd.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"SELECTED", "PO", "ORDR_GROUP_NO", "ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT"}
                    If .Columns.Exists(COLUMN_NAME) Then
                        .Columns(COLUMN_NAME).Header.Fixed = True
                    End If

                Next
                For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                    If GCOL.Key = "SELECTED" Then
                        GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                        GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                    End If
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                    GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If GCOL.Key.StartsWith("ORDR_AMT") Then
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                        GCOL.Width = 80
                        GCOL.Format = "#,##0"
                        Create_Summary(grd, GCOL.Key)
                    ElseIf GCOL.Key.StartsWith("ORDR_QTY") Then
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                        GCOL.Width = 70
                        GCOL.Format = "#,##0"
                        Create_Summary(grd, GCOL.Key)
                    ElseIf GCOL.Key.StartsWith("ORDR_CNT") Then
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                        GCOL.Width = 50
                        GCOL.Format = "#,##0"
                        Create_Summary(grdSOTORDR0, GCOL.Key)
                    ElseIf New String() {"ORDR_GROUP_NO", "ORDR_CUST_PO"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightCoral
                        GCOL.Width = 110
                    ElseIf New String() {"CUST_DC_NO", "ORDR_DEPT"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightCoral
                        GCOL.Width = 70
                    ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                    ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(GCOL.Key) Then
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                        GCOL.Width = 90
                    Else
                        GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    End If
                Next
            End With
        Next




        'ASCMAIN1.Add_Value_List(grdICTSTYLX, "ITEM_STATUS")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Verify POs"

                If txtDCPO.Text = "" Then
                    EMsg &= vbCr & "You Must First Provide a Block of Walmart DC/POs"
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

            Case "Verify POs"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Me.Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Verify POs").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                End With

            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpDCPO.Visible = Not ScreenMode
        grdSOTORDR0.Visible = ScreenMode

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid _
                In New UltraWinGrid.UltraGrid() _
                {grdSOTORDR0}
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

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR0", "SOTORDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        MyBase.EnforceConstraints(True)

        Clear_All_Filters(grdSOTORDR0)
    End Sub

    Private Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        dst.Tables("SOTORDR0").Rows.Clear()
        dst.Tables("SOTORDRP").Rows.Clear()

        Dim BAD_ROWS As String = ""

        Dim DCPOs As String = txtDCPO.Text
        For Each DCPOX As String In Split(DCPOs, vbCrLf)
            Dim DCPO2() As String = Split(DCPOX, vbTab)
            If DCPOX <> "" Then
                If DCPO2.Length <> 8 Then
                    BAD_ROWS &= vbCrLf & DCPOX
                End If
                If DCPO2.Length >= 2 Then Write_PO(DCPO2(0), DCPO2(1))
                If DCPO2.Length >= 4 Then Write_PO(DCPO2(2), DCPO2(3))
                If DCPO2.Length >= 6 Then Write_PO(DCPO2(4), DCPO2(5))
                If DCPO2.Length >= 8 Then Write_PO(DCPO2(6), DCPO2(7))
            End If
        Next

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Write_PO(DC As String, PO As String)

        Dim rowSOTORDRP As DataRow = dst.Tables("SOTORDRP").NewRow
        With rowSOTORDRP
            .Item("NO") = dst.Tables("SOTORDRP").Rows.Count + 1
            .Item("PO") = PO
            .Item("DC") = DC
        End With
        dst.Tables("SOTORDRP").Rows.Add(rowSOTORDRP)

        Fill_Records("SOTORDRX", New String() {PO})
        For Each ROW As DataRow In dst.Tables("SOTORDRX").Select("")

            Dim rowSOTORDR0 As DataRow = dst.Tables("SOTORDR0").NewRow
            For Each DCOL As DataColumn In ROW.Table.Columns
                rowSOTORDR0.Item(DCOL.ColumnName) = ROW.Item(DCOL.ColumnName)
            Next
            dst.Tables("SOTORDR0").Rows.Add(rowSOTORDR0)
            rowSOTORDR0.Item("DC") = DC
            rowSOTORDR0.Item("NO") = rowSOTORDRP.Item("NO")

            Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO")

            ASCMAIN1.sql = "Select ORDR_GROUP_NO" & vbCrLf _
                & ", MIN (STYLE_CODE) STYLE_CODE_1, MAX (STYLE_CODE) STYLE_CODE_2" & vbCrLf _
                & " from SOTORDR1,SOTORDR2" & vbCrLf _
                & " where SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & " group by SOTORDR1.ORDR_GROUP_NO"
            Dim rowSTYLE_CODE As DataRow = ASCDATA1.GetDataRow

            With rowSOTORDRP
                .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                .Item("GROUPS") = Val(.Item("GROUPS") & "") + 1
                .Item("STYLE_CODE_1") = rowSTYLE_CODE.Item("STYLE_CODE_1")
                .Item("STYLE_CODE_2") = rowSTYLE_CODE.Item("STYLE_CODE_2")
                .Item("ORDR_QTY") = Val(.Item("ORDR_QTY") & "") + Val(ROW.Item("ORDR_QTY") & "")
                .Item("ORDR_QTY_OPEN") = Val(.Item("ORDR_QTY_OPEN") & "") + Val(ROW.Item("ORDR_QTY_OPEN") & "")
                .Item("ORDR_QTY_PICK") = Val(.Item("ORDR_QTY_PICK") & "") + Val(ROW.Item("ORDR_QTY_PICK") & "")
                .Item("ORDR_QTY_SHIP") = Val(.Item("ORDR_QTY_SHIP") & "") + Val(ROW.Item("ORDR_QTY_SHIP") & "")
                .Item("ORDR_QTY_CANC") = Val(.Item("ORDR_QTY_CANC") & "") + Val(ROW.Item("ORDR_QTY_CANC") & "")
            End With
        Next

    End Sub

    Private Sub Update_Record()

        'Try
        '    MyBase.BeginTrans()
        '    Update_Record_TDA("ICTWALM1")
        '    MyBase.CommitTrans("Update Complete")

        'Catch ex As Exception
        '    MyBase.Rollback(ex.Message)
        'End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        '   Load_Popup_Menu(grdICTWALM1, "S", "Show Filter")
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
            'Case "CUST_CODE"
            '    If Not ScreenMode Then
            '        If e.KeyCode = System.Windows.Forms.Keys.Enter Then
            '            Set_SOTORDR0()
            '        End If
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)

            'Case "CUST_CODE"
            '    Set_SOTORDR0()

        End Select
    End Sub
#End Region

End Class