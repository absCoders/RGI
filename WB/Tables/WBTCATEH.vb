Imports Infragistics.Win
Imports Infragistics.Win.UltraWinGrid

Public Class WBTCATEH
    Private PAGE_CODE As String = String.Empty
    Private SB As New Text.StringBuilder With {.Length = 0}
    Private SQL_LOAD As New Text.StringBuilder With {.Length = 0}

#Region "ABS Standards"
    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey
            Case "New"
                Absx1.txtFor("PAGE_CODE").Text = ASCMAIN1.Next_Control_No("WBTCATEH.PAGE_CODE")
            Case "Edit"
                setIS_PVC()
        End Select
    End Sub

    Public Overrides Sub Proceed_Update_Special_Pre()
        MyBase.Proceed_Update_Special_Pre()
        'Stop

        Dim sql As String = String.Empty
        Dim rowWBTCATEH As DataRow = dst.Tables("WBTCATEH").Select("PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")(0)

        Update_Record_TDA("WBTCATED", "DELETE FROM WBTCATED WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")

        saveWBTCATU1()
        Update_Record_TDA("WBTCATU1", "DELETE FROM WBTCATU1 WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")
        'Update_Record_TDA("WBTCATU1")
        Update_Record_TDA("WBTCATU2", "DELETE FROM WBTCATU2 WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")
        setIS_PVC()
    End Sub

    Overrides Sub Show_Record_Special()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Loading Data")

        PAGE_CODE = MyBase.Absx1.txtFor("PAGE_CODE").Text.Trim

        Dim sql As String = String.Empty

        MyBase.EnforceConstraints(False)

        Dim S As New Text.StringBuilder With {.Length = 0}
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("WBTCATED.PAGE_CODE,")
        S.AppendLine("WBTCATED.STYLE_CODE,")
        S.AppendLine("WBTCATED.COLOR_CODE,")
        S.AppendLine("WBTCATED.COLOR_CODES,")
        S.AppendLine("WBTCATED.STYLE_ORDR,")
        S.AppendLine("ICTSTYL1.STYLE_DESC,")
        S.AppendLine("ICTSTYL1.STYLE_STATUS,")
        S.AppendLine("ICTPVC01.HEIGHT,")
        S.AppendLine("ICTPVC01.DIAMETER,")
        S.AppendLine("ICTPVC01.PVC_LENGTH,")
        S.AppendLine("ICTPVCLT.LIGHT_TYPE_DESC,")
        S.AppendLine("ICTPVC01.LIGHT_COUNT,")
        S.AppendLine("ICTPVCLC.LIGHT_COLOR_DESC,")
        S.AppendLine("ICTPVC01.TIP_COUNT,")
        S.AppendLine("ICTPVC01.G40_COUNT,")
        S.AppendLine("('*' || WBTCATED.STYLE_CODE || '*') AS BAR_CODE,")
        S.AppendLine("9999.99 AS FULL_CASE,")
        S.AppendLine("9999.99 AS FIVE_CASE,")
        S.AppendLine("ICTSTYL1.CARTON_PACK_QTY,")
        S.AppendLine("ICTSTYL1.CASE_CUBE")
        S.AppendLine("FROM WBTCATED, ICTSTYL1, ICTPVC01, ICTPVCLT, ICTPVCLC")
        S.AppendLine("WHERE WBTCATED.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        S.AppendLine("AND ICTSTYL1.STYLE_CODE = ICTPVC01.STYLE_CODE (+)")
        S.AppendLine("AND ICTPVC01.LIGHT_TYPE_CODE = ICTPVCLT.LIGHT_TYPE_CODE (+)")
        S.AppendLine("AND ICTPVC01.LIGHT_COLOR_CODE = ICTPVCLC.LIGHT_COLOR_CODE(+)")
        S.AppendLine($"AND WBTCATED.PAGE_CODE = '{PAGE_CODE}'")
        Call Fill_Records("WBTCATED", String.Empty, True, S.ToString)
        setIS_PVC()


        SB.Length = 0
        SB.AppendLine("SELECT *")
        SB.AppendLine("FROM WBTCATU1")
        SB.AppendLine("WHERE PAGE_CODE = '" & PAGE_CODE & "'")
        Call Fill_Records("WBTCATU1", String.Empty, True, SB.ToString)
        setWBTCATU1()

        SB.Length = 0
        SB.AppendLine("SELECT *")
        SB.AppendLine("FROM WBTCATU2")
        SB.AppendLine("WHERE PAGE_CODE = '" & PAGE_CODE & "'")
        Call Fill_Records("WBTCATU2", String.Empty, True, SB.ToString)

        UpdateExtraData()

        grdWBTCATED.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdWBTCATED.DisplayLayout.Bands(0).SortedColumns.Add("STYLE_ORDR", False)
        grdWBTCATED.DisplayLayout.Bands(0).SortedColumns.Add("STYLE_CODE", False)

        grdWBTCATU2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdWBTCATU2.DisplayLayout.Bands(0).SortedColumns.Add("ATTR_CODE", False)

        setColsVisable()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub UpdateExtraData()
        'Fill_Records("ICTSTYC1")
        Dim rowARTCUST1 As DataRow = Nothing
        For Each rowWBTCATED As DataRow In dst.Tables("WBTCATED").Select("", "STYLE_ORDR, STYLE_CODE")
            Dim STYLE_CODE As String = rowWBTCATED.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowWBTCATED.Item("COLOR_CODE").ToString & String.Empty

            'Current On-Hand
            Dim fltrS As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            Dim rowICTSTYCX As DataRow = dst.Tables("ICTSTYCX").Select(fltrS).FirstOrDefault
            If Not IsNothing(rowICTSTYCX) Then
                rowWBTCATED.Item("CURR_ON_HAND") = Val(rowICTSTYCX.Item("MSOH").ToString & String.Empty) + Val(rowICTSTYCX.Item("MSFT").ToString & String.Empty)
            End If

            'Pricing
            Dim Discounts As List(Of DISCOUNTS)
            Discounts = SOCMAIN2.Price_Discounts(Me, "", rowARTCUST1, STYLE_CODE, False)
            If Discounts(2).DISCOUNT_QTY = 0 Then
                rowWBTCATED.Item("FULL_CASE") = Null
                rowWBTCATED.Item("FULL_CASE") = Null
            Else
                rowWBTCATED.Item("FULL_CASE") = Discounts(2).DISCOUNT_QTY
                rowWBTCATED.Item("FULL_CASE") = Format(Discounts(2).DISCOUNT_PRICE, "###,##0.00")
            End If

            If Discounts(1).DISCOUNT_QTY = 0 Then
                rowWBTCATED.Item("FIVE_CASE") = Null
                rowWBTCATED.Item("FIVE_CASE") = Null
            Else
                rowWBTCATED.Item("FIVE_CASE") = Discounts(1).DISCOUNT_QTY
                rowWBTCATED.Item("FIVE_CASE") = Format(Discounts(1).DISCOUNT_PRICE, "###,##0.00")
            End If

            'Colors
            Dim COLORS As String = ""
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(fltrS, "COLOR_CODE")
                Dim CC As String = rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty
                COLORS = COLORS & $"{CC},"
            Next
            If COLORS.Length > 0 Then
                COLORS = COLORS.Substring(0, COLORS.Length - 1)
            End If
            rowWBTCATED.Item("COLOR_CODES") = COLORS

            'Theme
            Dim THEME_DESC As String = ""
            Dim sql As New Text.StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT DISTINCT TH1.THEME_DESC")
            sql.AppendLine("FROM ICTSTYL1 ST1, ICTSTYC1 CL1, ICTTHEME TH1")
            sql.AppendLine("WHERE ST1.STYLE_CODE = CL1.STYLE_CODE")
            sql.AppendLine("AND CL1.THEME_CODE = TH1.THEME_CODE")
            sql.AppendLine($"AND ST1.STYLE_CODE = '{STYLE_CODE}'")
            sql.AppendLine("GROUP BY TH1.THEME_DESC")
            sql.AppendLine("ORDER BY TH1.THEME_DESC DESC")
            Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString())
            For Each rowICTTHEME As DataRow In tbl.Select("", "theme_desc")
                THEME_DESC = $"{THEME_DESC} | {rowICTTHEME.Item("THEME_DESC").ToString & String.Empty}"
            Next
            If THEME_DESC.Length > 3 Then
                THEME_DESC = THEME_DESC.Substring(3, THEME_DESC.Length - 3)
            End If
            rowWBTCATED.Item("THEME_DESC") = THEME_DESC

            'STYLE_ORDR if empty
            Dim STYLE_ORDR As String = rowWBTCATED.Item("STYLE_ORDR").ToString & String.Empty
            If STYLE_ORDR.Length = 0 Then
                rowWBTCATED.Item("STYLE_ORDR") = getNextSort()
            End If
        Next
        grdWBTCATED.UpdateData()
        grdWBTCATED.Refresh()
    End Sub

    Overrides Sub Clear_Record_Special()

        If ScreenMode Then
            MyBase.EnforceConstraints(False)
            dst.Tables("WBTCATED").Rows.Clear()
            dst.Tables("WBTCATU1").Rows.Clear()
            dst.Tables("WBTCATU2").Rows.Clear()
            txtSTYLE_CLASS_CODE.Text = ""
            txtTHEME_CODE.Text = ""
            txtSTYLE_CLASS_CODE2.Text = ""
            setIS_PVC()
            MyBase.EnforceConstraints(True)

            PAGE_CODE = String.Empty
        End If

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        With grdWBTCATU2.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
        End With

        'With grdWBTCATED.DisplayLayout
        '    .Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
        '    .Override.AllowDelete = DefaultableBoolean.True
        '    .Override.AllowUpdate = DefaultableBoolean.True
        '    For i As Integer = 0 To .Bands(0).Columns.Count - 1
        '        .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        '    Next i
        '    .Bands(0).Columns("STYLE_ORDR").CellActivation = UltraWinGrid.Activation.AllowEdit
        '    .Bands(0).Columns("STYLE_ORDR").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        'End With

        setIS_PVC()
        setColsVisable()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWBTCATU2, "SSB", "Show Filter", "Show GroupBox", "Add Attributes")
        Load_Popup_Menu(grdWBTCATED, "SSB", "Show Filter", "Show GroupBox", "Clear All")
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

            Case "grdWBTCATU2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Attributes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdWBTCATED"
                tlb_btn = DirectCast(tlb_pop.Tools("Clear All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            Case "Add Attributes"
                If grd.Name = "grdWBTCATU2" Then
                    Add_Attributes(grdWBTCATU2, "ICTATTR1", "ATTR_CODE", "Attributes")
                End If
            Case "Clear All"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Clear All"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Clear All Items")
                iMSG.AppendLine("Associated For This Page!!")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    dst.Tables.Item("WBTCATED").Clear()
                    setIS_PVC()
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

    Private Sub Add_Attributes(grd As UltraWinGrid.UltraGrid,
        TABLE_NAME As String,
        COLUMN_NAME As String,
        Codes_Caption As String)

        'Dim TABLE_NAME_grid As String = DirectCast(grd.DataSource, DataTable).TableName
        'ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME, , sql_where)
        Dim S As New Text.StringBuilder With {.Length = 0}
        Dim sql_where As String = Get_List_of_Codes(TABLE_NAME & "." & COLUMN_NAME & " not in", "WBTCATU2", COLUMN_NAME)
        If sql_where.Length = 0 Then
            'ASCMAIN1.CodeSelector.SQL = String.Format("SELECT ATTR_CODE, ATTR_DESC, '{0}' AS PAGE_CODE FROM ICTATTR1", Absx1.txtFor("PAGE_CODE").Text)
            'ASCMAIN1.CodeSelector.SQL = "SELECT ATTR_CODE, ATTR_DESC FROM ICTATTR1 ORDER BY ATTR_CODE"
            S.AppendLine("SELECT ATTR_CODE, ATTR_DESC FROM ICTATTR1 ORDER BY ATTR_CODE")
        Else
            'ASCMAIN1.CodeSelector.SQL = String.Format("SELECT ATTR_CODE, ATTR_DESC FROM ICTATTR1 WHERE {0} ORDER BY ATTR_CODE", sql_where)
            S.AppendLine(String.Format("SELECT ATTR_CODE, ATTR_DESC FROM ICTATTR1 WHERE {0} ORDER BY ATTR_CODE", sql_where))
        End If
        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Page Attributes"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For i As Int64 = 0 To ASCMAIN1.CodeSelector.Selections - 1
                Dim ATTR_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(i).Item("ATTR_CODE") & ""
                Dim rowWBTCATU2 As DataRow = dst.Tables.Item("WBTCATU2").NewRow
                rowWBTCATU2.Item("PAGE_CODE") = Absx1.txtFor("PAGE_CODE").Text
                rowWBTCATU2.Item("ATTR_CODE") = ATTR_CODE
                dst.Tables.Item("WBTCATU2").Rows.Add(rowWBTCATU2)
            Next
        End If


        'If ASCMAIN1.CodeSelector.SQL <> "" Then
        '    ASCMAIN1.CodeSelector.MultipleSelections = True
        '    ASCMAIN1.CodeSelector.tblASTVIEW1 = New DataTable
        '    Dim F As New ASFCODE1

        '    F.ShowDialog()
        '    F.Dispose()
        '    If ASCMAIN1.CodeSelector.Selections <> 0 Then
        '        Me.Cursor = Cursors.WaitCursor
        '        ASCMAIN1.Progress("Now Loading " & Codes_Caption)

        '        grd.Visible = False
        '        If grd.ActiveRow IsNot Nothing Then grd.ActiveRow.CancelUpdate()
        '        For Each CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
        '            grd.ActiveRow = grd.DisplayLayout.Bands(0).AddNew
        '            grd.ActiveRow.Cells(COLUMN_NAME).Value = CODE
        '            grd.ActiveRow.Update()
        '        Next
        '        grd.Visible = True
        '        Me.Cursor = Cursors.Default
        '        ASCMAIN1.Progress("")
        '    End If
        'End If

    End Sub
#End Region

#Region "Form Controls"

    Private Sub btnCLASS_ADD_Click(sender As Object, e As EventArgs) Handles btnCLASS_ADD.Click
        If (EntryMode = "Edit" Or EntryMode = "New") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Add By Class Code"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}

            If txtSTYLE_CLASS_CODE.Text.Length = 0 Then
                iMSG.AppendLine("Missing Class Code")
                MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            Else
                iMSG.AppendLine("This Will Add All Styles")
                iMSG.AppendLine("To This Page With Class Code")
                iMSG.AppendLine(String.Format("Of {0}.", txtSTYLE_CLASS_CODE.Text))
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    ADD_TO_WBTCATED(1)
                End If
            End If
        End If
    End Sub

    Private Sub btnTHEME_ADD_Click(sender As Object, e As EventArgs) Handles btnTHEME_ADD.Click
        If (EntryMode = "Edit" Or EntryMode = "New") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Add By Theme Code"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}

            If txtTHEME_CODE.Text.Length = 0 Then
                iMSG.AppendLine("Missing Theme Code")
                MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            Else
                iMSG.AppendLine("This Will Add All Styles")
                iMSG.AppendLine("To This Page With Theme Code")
                iMSG.AppendLine(String.Format("Of {0}.", txtTHEME_CODE.Text))
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    ADD_TO_WBTCATED(2)
                End If
            End If
        End If
    End Sub

    Private Sub btnATTR_ADD_Click(sender As Object, e As EventArgs) Handles btnATTR_ADD.Click
        If (EntryMode = "Edit" Or EntryMode = "New") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Add By Attribute Code"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}

            If dst.Tables.Item("WBTCATU2").Rows.Count = 0 Then
                iMSG.AppendLine("Missing Attribute Codes")
                MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            Else
                iMSG.AppendLine("This Will Add All Styles")
                iMSG.AppendLine("To This Page With Selected")
                iMSG.AppendLine("Class Code (if Any) and Attributes.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    ADD_TO_WBTCATED(3)
                End If
            End If
        End If
    End Sub

    Private Sub btnBuildOnRules_Click(sender As Object, e As EventArgs) Handles btnBuildOnRules.Click
        If (EntryMode = "Edit" Or EntryMode = "New") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Add By All Rules"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("This Will Add All Styles")
            iMSG.AppendLine("To This Page With All Rules")
            iMSG.AppendLine("Defined.")
            iMSG.AppendLine("")
            iMSG.AppendLine("Is That What You Want?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                If txtSTYLE_CLASS_CODE.Text.Length > 0 Then
                    ADD_TO_WBTCATED(1)
                End If
                If txtTHEME_CODE.Text.Length > 0 Then
                    ADD_TO_WBTCATED(2)
                End If
                If dst.Tables.Item("WBTCATU2").Rows.Count > 0 Then
                    ADD_TO_WBTCATED(3)
                End If
            End If
        End If
    End Sub

    Private Sub grdWBTCATED_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWBTCATED.BeforeRowUpdate
        e.Row.Cells("PAGE_CODE").Value = PAGE_CODE

        Dim STYLE_CODE As String = (e.Row.Cells("STYLE_CODE").Value & String.Empty).ToString.ToUpper.Trim
        Dim styleInList As Boolean = dst.Tables("WBTCATED").Select("STYLE_CODE = '" & STYLE_CODE & "'").Length > 0
        e.Row.Cells("STYLE_CODE").Value = STYLE_CODE

        If e.Row.IsAddRow AndAlso styleInList Then
            e.Cancel = True
            MessageBox.Show("Style already belongs to the Page.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf Not e.Row.IsAddRow AndAlso styleInList AndAlso
                e.Row.Cells("STYLE_CODE").Value & String.Empty <> e.Row.Cells("STYLE_CODE").OriginalValue & String.Empty Then
            e.Cancel = True
            MessageBox.Show("Style already belongs to the Page.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            e.Row.Cells("STYLE_ORDR").Value = getNextSort()
            e.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
            e.Row.Cells("STYLE_STATUS").Value = rowICTSTYL1.Item("STYLE_STATUS") & String.Empty
        End If
    End Sub

    Private Sub grdWBTCATED_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWBTCATED.ClickCellButton
        If grdWBTCATED.ActiveRow Is Nothing Then Exit Sub
        If Not e.Cell.Row.IsAddRow Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim sql_where As String = ""
        Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
        Dim VIEW_NAME As String = String.Empty

        Select Case grd.ActiveCell.Column.Key
            Case "STYLE_CODE"
                Dim S As New Text.StringBuilder With {.Length = 0}
                S.AppendLine("SELECT")
                S.AppendLine("S1.STYLE_CODE,")
                S.AppendLine("C1.COLOR_CODE,")
                S.AppendLine("S1.STYLE_DESC,")
                S.AppendLine("C1.THEME_CODE")
                S.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1")
                S.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE")
                S.AppendLine($"AND (S1.STYLE_CODE, C1.COLOR_CODE) NOT IN (SELECT STYLE_CODE, COLOR_CODE FROM WBTCATED WHERE PAGE_CODE = '{PAGE_CODE}')")
                S.AppendLine("ORDER BY")
                S.AppendLine("S1.STYLE_CODE,")
                S.AppendLine("C1.COLOR_CODE")
                With ASCMAIN1.CodeSelector
                    .SQL = S.ToString
                    .MultipleSelections = False
                    .PreviouslySelectedCodes0 = ""
                    .Caption = "Select Styles"
                    .TABLE_NAME = ""
                    .VIEW_NAME = ""
                    .VIEW_DESC = ""
                    .COLUMN_NAME = ""
                    .COLUMN_PREKEYs = New Dictionary(Of String, String)
                    '.Custom_sql_where = ""
                    .tblASTVIEW1 = New DataTable
                End With
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                        Dim STYLE_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STYLE_CODE") & ""
                        Dim COLOR_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("COLOR_CODE") & ""
                        Dim STYLE_DESC As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STYLE_DESC") & ""
                        Dim THEME_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("THEME_CODE") & ""
                        e.Cell.Row.Cells.Item("STYLE_CODE").Value = STYLE_CODE
                        e.Cell.Row.Cells.Item("COLOR_CODE").Value = COLOR_CODE
                        e.Cell.Row.Cells.Item("STYLE_DESC").Value = STYLE_DESC
                        'e.Cell.Row.Cells.Item("THEME_CODE").Value = THEME_CODE
                    Next
                    UpdateExtraData()
                End If

        End Select
        'Call grdClickCellButton(grdWBTCATED, sql_where, True)
    End Sub

    Private Sub grdWBTCATED_BeforeCellActivate(sender As Object, e As CancelableCellEventArgs) Handles grdWBTCATED.BeforeCellActivate
        If Not e.Cell.Row.IsAddRow Then
            e.Cancel = True
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub ADD_TO_WBTCATED(ByVal AddType As Integer, Optional ShowErr As Boolean = True)
        '1 = Class
        '2 = Theme
        '3 = Class / Attribute
        '4 = Discontinued

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Loading Data")

        Dim err As New Text.StringBuilder With {.Length = 0}
        Dim sql As New Text.StringBuilder With {.Length = 0}
        Dim sqlin As New Text.StringBuilder With {.Length = 0}
        Dim sqlgb As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine(String.Format("'{0}' AS PAGE_CODE,", Absx1.txtFor("PAGE_CODE").Text))
        sql.AppendLine("S1.STYLE_CODE,")
        sql.AppendLine("C1.COLOR_CODE,")
        sql.AppendLine("S1.STYLE_DESC,")
        sql.AppendLine("S1.STYLE_STATUS")
        sql.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1")
        sql.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE ")

        sqlgb.AppendLine(" GROUP BY S1.STYLE_CODE, C1.COLOR_CODE,  S1.STYLE_DESC, S1.STYLE_STATUS")

        Select Case AddType
            Case 1
                If txtSTYLE_CLASS_CODE.Text.Length = 0 Then
                    err.AppendLine("Missing Class Code")
                Else
                    sqlin.AppendLine("AND S1.STYLE_CODE IN")
                    sqlin.AppendLine("(")
                    sqlin.AppendLine("  SELECT DISTINCT STYLE_CODE")
                    sqlin.AppendLine("  FROM ICTSTYL1")
                    sqlin.AppendLine(String.Format("  WHERE STYLE_CLASS_CODE = '{0}'", txtSTYLE_CLASS_CODE.Text))
                    sqlin.AppendLine(")")
                End If
            Case 2
                If txtTHEME_CODE.Text.Length = 0 Then
                    err.AppendLine("Missing Theme Code")
                Else
                    sqlin.AppendLine("AND S1.STYLE_CODE IN")
                    sqlin.AppendLine("(")
                    sqlin.AppendLine("  SELECT DISTINCT STYLE_CODE")
                    sqlin.AppendLine("  FROM ICTSTYC1")
                    sqlin.AppendLine(String.Format("  WHERE THEME_CODE = '{0}'", txtTHEME_CODE.Text))
                    sqlin.AppendLine(")")
                End If
            Case 3
                If txtSTYLE_CLASS_CODE2.Text.Length = 0 Then
                    If dst.Tables.Item("WBTCATU2").Rows.Count = 0 Then
                        err.AppendLine("Missing Class and Attribute Codes")
                    End If
                End If
                If err.Length = 0 Then
                    If txtSTYLE_CLASS_CODE2.Text.Length > 0 Then
                        sqlin.AppendLine("AND S1.STYLE_CODE IN")
                        sqlin.AppendLine("(")
                        sqlin.AppendLine("  SELECT DISTINCT STYLE_CODE")
                        sqlin.AppendLine("  FROM ICTSTYL1")
                        sqlin.AppendLine(String.Format("  WHERE STYLE_CLASS_CODE = '{0}'", txtSTYLE_CLASS_CODE2.Text))
                        sqlin.AppendLine(")")
                    End If
                    Dim ATIN As String = ""
                    For Each rowWBTCATU2 As DataRow In dst.Tables("WBTCATU2").Select()
                        ATIN = ATIN & "'" & rowWBTCATU2.Item("ATTR_CODE") & "',"
                    Next
                    ATIN = ATIN.Substring(0, ATIN.Length - 1)
                    'Stop
                    sqlin.AppendLine("AND S1.STYLE_CODE IN")
                    sqlin.AppendLine("(")
                    sqlin.AppendLine("  SELECT DISTINCT STYLE_CODE")
                    sqlin.AppendLine("  FROM ICTSTYL3")
                    sqlin.AppendLine(String.Format("  WHERE ATTR_CODE IN ({0})", ATIN))
                    sqlin.AppendLine(")")
                End If
            Case 4
                sqlin.AppendLine("AND S1.STYLE_CODE IN")
                sqlin.AppendLine("(")
                sqlin.AppendLine("  SELECT DISTINCT STYLE_CODE")
                sqlin.AppendLine("  FROM ICTSTYC1")
                sqlin.AppendLine("  WHERE STYLE_COLOR_STATUS = 'D'")
                sqlin.AppendLine(")")
        End Select
        If err.Length > 0 And ShowErr Then
            MsgBox(err, vbOKOnly, "Add Errors")
        Else
            Dim SQLF As String = sql.ToString & sqlin.ToString & sqlgb.ToString
            'Stop
            Dim tblICTSTYL1 As DataTable = ASCDATA1.GetDataTable(SQLF, String.Empty)
            For Each rowICTSTYL1 As DataRow In tblICTSTYL1.Rows
                Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
                Dim COLOR_CODE As String = rowICTSTYL1.Item("COLOR_CODE").ToString & String.Empty
                Dim filter As String = String.Format($"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'")
                If dst.Tables.Item("WBTCATED").Select(filter).Count = 0 Then
                    Dim newWBTCATU2 As DataRow = dst.Tables.Item("WBTCATED").NewRow
                    For Each COL As String In New String() {"PAGE_CODE", "STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "STYLE_STATUS"}
                        newWBTCATU2.Item(COL) = rowICTSTYL1.Item(COL).ToString & String.Empty
                    Next
                    newWBTCATU2.Item("STYLE_ORDR") = getNextSort()
                    dst.Tables.Item("WBTCATED").Rows.Add(newWBTCATU2)
                End If
            Next
            UpdateExtraData()
        End If
        setIS_PVC()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Function getNextSort() As Int64
        Dim RETVAL As Int64
        RETVAL = Val(dst.Tables("WBTCATED").Compute("MAX(STYLE_ORDR)", "").ToString & String.Empty) + 1
        Return RETVAL
    End Function
    Private Sub saveWBTCATU1()
        If dst.Tables.Item("WBTCATU1").Rows.Count = 1 Then
            Dim rowWBTCATU1 As DataRow = dst.Tables.Item("WBTCATU1").Rows(0)
            rowWBTCATU1.Item("STYLE_CLASS_CODE") = txtSTYLE_CLASS_CODE.Text
            rowWBTCATU1.Item("THEME_CODE") = txtTHEME_CODE.Text
            rowWBTCATU1.Item("STYLE_CLASS_CODE2") = txtSTYLE_CLASS_CODE2.Text
        End If
    End Sub

    Private Sub setWBTCATU1()
        If dst.Tables.Item("WBTCATU1").Rows.Count = 0 Then
            Dim rowWBTCATU1 As DataRow = dst.Tables.Item("WBTCATU1").NewRow
            rowWBTCATU1.Item("PAGE_CODE") = PAGE_CODE
            dst.Tables.Item("WBTCATU1").Rows.Add(rowWBTCATU1)
        End If
        txtSTYLE_CLASS_CODE.Text = dst.Tables.Item("WBTCATU1").Rows(0).Item("STYLE_CLASS_CODE").ToString & String.Empty
        txtTHEME_CODE.Text = dst.Tables.Item("WBTCATU1").Rows(0).Item("THEME_CODE").ToString & String.Empty
        txtSTYLE_CLASS_CODE2.Text = dst.Tables.Item("WBTCATU1").Rows(0).Item("STYLE_CLASS_CODE2").ToString & String.Empty
    End Sub

    Private Sub WBTCATEH_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "WBTCATU1", "*", 1, True, String.Empty, 1)
            Create_TDA(.Tables.Add, "WBTCATU2", "*", 2, True, String.Empty, 2)

            Dim S As New Text.StringBuilder With {.Length = 0}
            S.Length = 0
            S.AppendLine("SELECT")
            S.AppendLine("WBTCATED.PAGE_CODE,")
            S.AppendLine("WBTCATED.STYLE_CODE,")
            S.AppendLine("WBTCATED.COLOR_CODE,")
            S.AppendLine("WBTCATED.COLOR_CODES,")
            S.AppendLine("WBTCATED.STYLE_ORDR,")
            S.AppendLine("ICTSTYL1.STYLE_DESC,")
            S.AppendLine("ICTSTYL1.STYLE_STATUS,")
            S.AppendLine("ICTPVC01.HEIGHT,")
            S.AppendLine("ICTPVC01.DIAMETER,")
            S.AppendLine("ICTPVC01.PVC_LENGTH,")
            S.AppendLine("ICTPVCLT.LIGHT_TYPE_DESC,")
            S.AppendLine("ICTPVC01.LIGHT_COUNT,")
            S.AppendLine("ICTPVCLC.LIGHT_COLOR_DESC,")
            S.AppendLine("ICTPVC01.TIP_COUNT,")
            S.AppendLine("ICTPVC01.G40_COUNT,")
            S.AppendLine("('*' || WBTCATED.STYLE_CODE || '*') AS BAR_CODE,")
            S.AppendLine("9999.99 AS FULL_CASE,")
            S.AppendLine("9999.99 AS FIVE_CASE,")
            S.AppendLine("ICTSTYL1.CARTON_PACK_QTY,")
            S.AppendLine("ICTSTYL1.CASE_CUBE")
            S.AppendLine("FROM WBTCATED, ICTSTYL1, ICTPVC01, ICTPVCLT, ICTPVCLC")
            S.AppendLine("WHERE WBTCATED.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            S.AppendLine("AND ICTSTYL1.STYLE_CODE = ICTPVC01.STYLE_CODE (+)")
            S.AppendLine("AND ICTPVC01.LIGHT_TYPE_CODE = ICTPVCLT.LIGHT_TYPE_CODE (+)")
            S.AppendLine("AND ICTPVC01.LIGHT_COLOR_CODE = ICTPVCLC.LIGHT_COLOR_CODE (+)")
            ASCMAIN1.sql = S.ToString
            Create_TDA(.Tables.Add, "WBTCATED", "**", 2, True, String.Empty, 2)
            'Create_TDA(.Tables.Add, "WBTCATED", S.ToString)
            With dst.Tables("WBTCATED")
                .Columns.Add("CURR_ON_HAND", GetType(System.Int64))
                .Columns.Add("THEME_DESC", GetType(System.String))
            End With

            Dim SB As New Text.StringBuilder With {.Length = 0}
            SB.Length = 0
            SB.AppendLine("SELECT * FROM")
            SB.AppendLine("  (")
            SB.AppendLine("   SELECT C1.STYLE_CODE,")
            SB.AppendLine("   CASE WHEN")
            SB.AppendLine("   SUM(")
            SB.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            SB.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SB.AppendLine("     ELSE 0")
            SB.AppendLine("     END) < 0")
            SB.AppendLine("   THEN")
            SB.AppendLine("     0")
            SB.AppendLine("   ELSE")
            SB.AppendLine("   SUM(")
            SB.AppendLine("     CASE S2.WHSE_CODE")
            SB.AppendLine("     WHEN 'MS'")
            SB.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SB.AppendLine("     ELSE 0")
            SB.AppendLine("     END)")
            SB.AppendLine("   END AS MSOH,")
            SB.AppendLine("   CASE WHEN")
            SB.AppendLine("   SUM(")
            SB.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            SB.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SB.AppendLine("     ELSE 0")
            SB.AppendLine("     END) <= 0")
            SB.AppendLine("   THEN")
            SB.AppendLine("     0")
            SB.AppendLine("   ELSE")
            SB.AppendLine("     CASE WHEN")
            SB.AppendLine("       SUM(")
            SB.AppendLine("       CASE S2.WHSE_CODE")
            SB.AppendLine("       WHEN 'MS'")
            SB.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SB.AppendLine("       ELSE 0")
            SB.AppendLine("       END) < 0")
            SB.AppendLine("     THEN")
            SB.AppendLine("       0")
            SB.AppendLine("     ELSE")
            SB.AppendLine("     SUM(")
            SB.AppendLine("       CASE S2.WHSE_CODE")
            SB.AppendLine("       WHEN 'MS'")
            SB.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SB.AppendLine("       ELSE 0")
            SB.AppendLine("       END) END")
            SB.AppendLine("   END AS MSFT")
            SB.AppendLine("   FROM ICTSTYC1 C1")
            SB.AppendLine("   LEFT JOIN ICTSTAT2 S2")
            SB.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
            SB.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
            SB.AppendLine("   INNER JOIN ICTCOLR1 C2")
            SB.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
            SB.AppendLine("   GROUP BY C1.STYLE_CODE")
            SB.AppendLine("  )")
            ASCMAIN1.sql = SB.ToString
            Create_TDA(dst.Tables.Add, "ICTSTYCX", "**", 0, False, "", 2)
            Fill_Records("ICTSTYCX")

            S.Length = 0
            S.AppendLine("SELECT STYLE_CODE, COLOR_CODE from ICTSTYC1")
            ASCMAIN1.sql = S.ToString
            Create_TDA(dst.Tables.Add, "ICTSTYC1", "**", 0, False, "", 2)
            Fill_Records("ICTSTYC1")

        End With

        grdWBTCATED.DataSource = dst.Tables("WBTCATED")
        grdWBTCATU2.DataSource = dst.Tables("WBTCATU2")
        Create_Summary(grdWBTCATED, "STYLE_CODE", "Count")

        Create_Summary(grdWBTCATU2, "ATTR_CODE", "Count")

        With grdWBTCATED.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Appearance.ForeColor = System.Drawing.Color.White
            .Columns("STYLE_CODE").Header.Appearance.BackColor2 = System.Drawing.Color.Blue
            .Columns("STYLE_CODE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("COLOR_CODE").Header.Appearance.ForeColor = System.Drawing.Color.White
            .Columns("COLOR_CODE").Header.Appearance.BackColor2 = System.Drawing.Color.Blue
            .Columns("COLOR_CODE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        End With

        With grdWBTCATU2.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
        End With


    End Sub

    Private Sub btnDiscontinued_Click(sender As Object, e As EventArgs) Handles btnDiscontinued.Click
        ADD_TO_WBTCATED(4)
    End Sub

    Private Sub btnUploadXLS_Click(sender As Object, e As EventArgs) Handles btnUploadXLS.Click
        If EntryMode = "Edit" Or EntryMode = "New" Then
            Dim str As New Text.StringBuilder With {.Length = 0}
            Dim SQLS As New Text.StringBuilder With {.Length = 0}
            Dim tableData As New DataTable

            'Dim ECOM_CODE As String = cboPartnerUpsert.Text

            str.AppendLine("This Will Allow You To Upload")
            str.AppendLine("A File To Add New Styles")
            str.AppendLine("")
            str.AppendLine("It Should Be In The Same Format As If")
            str.AppendLine("You Has Exported The Grid Below.")
            str.AppendLine("")
            str.AppendLine("It Needs To Have At Least Two Column")
            str.AppendLine("Titled: Style & Color")
            str.AppendLine("")
            str.AppendLine("Are You Ready?")
            Dim iResult As MsgBoxResult = MsgBox(str.ToString, vbYesNo, "Upload Styles?")
            Dim fileToImport As String = String.Empty
            If iResult = MsgBoxResult.Yes Then
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Open File To Upsert"
                    openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx"
                    openFileDialog1.FilterIndex = 1
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        fileToImport = openFileDialog1.FileName
                    End If

                    openFileDialog1.Dispose()
                End Using
                If fileToImport.Length = 0 Then
                    Exit Sub
                End If
                '---------------
                If fileToImport <> "" Then

                    Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(fileToImport)
                    Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                    Dim range As SpreadsheetGear.IRange = Nothing

                    Dim BAD_STYLES As New List(Of String)
                    Dim DUPE_STYLES As New List(Of String)

                    ASCMAIN1.Progress("Now Loading from XLS")

                    Dim r As Integer = 0
                    Dim Blanks As Int64 = 0
                    Dim HeaderFound As Boolean = False
                    Do While Blanks <> 25
                        Dim STYLE_CODE As String = oSheet.Cells(r, 0).Text & ""
                        Dim COLOR_CODE As String = ""
                        If Not chkIS_PVC.Checked Then
                            COLOR_CODE = oSheet.Cells(r, 1).Text & ""
                        End If
                        If STYLE_CODE = "" Or IsNumeric(STYLE_CODE) Or STYLE_CODE = "Totals" Then
                            Blanks += 1
                        End If
                        If STYLE_CODE = "Style" Then
                            HeaderFound = True
                        End If
                        If HeaderFound = True And STYLE_CODE.Length > 0 And STYLE_CODE <> "Style" And Not IsNumeric(STYLE_CODE) And STYLE_CODE <> "Totals" Then
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            If IsNothing(rowICTSTYL1) Then
                                If Not BAD_STYLES.Contains(STYLE_CODE) Then
                                    BAD_STYLES.Add(STYLE_CODE)
                                End If
                            Else
                                Dim flt As String = ""
                                If chkIS_PVC.Checked Then
                                    flt = $"STYLE_CODE = '{STYLE_CODE}'"
                                Else
                                    flt = $"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'"
                                End If

                                If dst.Tables.Item("WBTCATED").Select(flt).Count > 0 Then
                                    Dim STYLE_COLOR As String = ""
                                    If chkIS_PVC.Checked Then
                                        STYLE_COLOR = $"{STYLE_CODE}"
                                    Else
                                        STYLE_COLOR = $"{STYLE_CODE}-{COLOR_CODE}"
                                    End If
                                    If Not DUPE_STYLES.Contains(STYLE_COLOR) Then
                                        DUPE_STYLES.Add(STYLE_COLOR)
                                    End If
                                Else
                                    Dim newWBTCATED As DataRow = dst.Tables.Item("WBTCATED").NewRow
                                    newWBTCATED.Item("PAGE_CODE") = MyBase.Absx1.txtFor("PAGE_CODE").Text
                                    newWBTCATED.Item("STYLE_CODE") = STYLE_CODE
                                    If Not chkIS_PVC.Checked Then
                                        newWBTCATED.Item("COLOR_CODE") = COLOR_CODE
                                    Else
                                        newWBTCATED.Item("COLOR_CODE") = "PVC"
                                    End If
                                    newWBTCATED.Item("STYLE_ORDR") = getNextSort()
                                    newWBTCATED.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC").ToString & String.Empty
                                    newWBTCATED.Item("STYLE_STATUS") = rowICTSTYL1.Item("STYLE_STATUS") & String.Empty
                                    dst.Tables.Item("WBTCATED").Rows.Add(newWBTCATED)
                                End If
                            End If
                        End If
                        r += 1
                        ASCMAIN1.Progress("-", CStr(r))
                    Loop

                    If BAD_STYLES.Count <> 0 Then
                        MsgBox("The following invalid Styles have been encountered: " & Join(BAD_STYLES.ToArray, ","), MsgBoxStyle.OkOnly, "Warning")
                    End If
                    If DUPE_STYLES.Count <> 0 Then
                        MsgBox("The following Duplicate Styles have been encountered: " & Join(DUPE_STYLES.ToArray, ","), MsgBoxStyle.OkOnly, "Warning")
                    End If
                End If
                UpdateExtraData()
                setIS_PVC()
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
                Dim mg As New Text.StringBuilder With {.Length = 0}
                mg.AppendLine("Import Complete.")
                mg.AppendLine("Please Save and Re-load")
                mg.AppendLine("To See PVC Data.")

                MsgBox(mg.ToString, vbOKOnly, "Done")
            End If
        End If
    End Sub

    Private Sub chkIS_PVC_CheckedChanged(sender As Object, e As EventArgs) Handles chkIS_PVC.CheckedChanged
        setIS_PVC()

        If chkIS_PVC.Checked Then

        End If
        setColsVisable()
    End Sub
    Private Sub setIS_PVC()
        Dim DETL_RECS As Int64 = dst.Tables.Item("WBTCATED").Rows.Count
        If DETL_RECS > 0 Then
            chkIS_PVC.Enabled = False
        Else
            chkIS_PVC.Enabled = True
        End If
    End Sub
    Private Sub setColsVisable()
        With grdWBTCATED.DisplayLayout.Bands(0)
            .Columns.Item("STYLE_CODE").Hidden = False
            .Columns.Item("STYLE_ORDR").Hidden = False
            .Columns.Item("HEIGHT").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("DIAMETER").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("PVC_LENGTH").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("COLOR_CODE").Hidden = chkIS_PVC.Checked
            .Columns.Item("COLOR_CODES").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("LIGHT_TYPE_DESC").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("LIGHT_COUNT").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("TIP_COUNT").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("G40_COUNT").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("BAR_CODE").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("FULL_CASE").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("FIVE_CASE").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("CARTON_PACK_QTY").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("CASE_CUBE").Hidden = Not chkIS_PVC.Checked
            .Columns.Item("STYLE_DESC").Hidden = False
            .Columns.Item("STYLE_STATUS").Hidden = chkIS_PVC.Checked
            .Columns.Item("CURR_ON_HAND").Hidden = chkIS_PVC.Checked
            .Columns.Item("THEME_DESC").Hidden = chkIS_PVC.Checked
        End With
    End Sub
#End Region
End Class