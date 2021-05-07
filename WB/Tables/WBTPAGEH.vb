Imports Infragistics.Win
Imports Infragistics.Win.UltraWinGrid

Public Class WBTPAGEH

    Private PAGE_CODE As String = String.Empty

#Region "ABS Standards"
    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey
            Case "New"
                Absx1.txtFor("PAGE_CODE").Text = ASCMAIN1.Next_Control_No("WBTPAGEH.PAGE_CODE")
        End Select
    End Sub

    Public Overrides Sub Proceed_Update_Special_Pre()
        MyBase.Proceed_Update_Special_Pre()
        'Stop

        Dim sql As String = String.Empty
        Dim rowWBTPAGEH As DataRow = dst.Tables("WBTPAGEH").Select("PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")(0)

        Update_Record_TDA("WBTPAGED", "DELETE FROM WBTPAGED WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")

        saveWBTPAGU1()
        Update_Record_TDA("WBTPAGU1", "DELETE FROM WBTPAGU1 WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")
        'Update_Record_TDA("WBTPAGU1")
        Update_Record_TDA("WBTPAGU2", "DELETE FROM WBTPAGU2 WHERE PAGE_CODE = '" & MyBase.Absx1.txtFor("PAGE_CODE").Text & "'")

        Update_Record_TDA("WBTSTYLD")

    End Sub

    Overrides Sub Show_Record_Special()

        PAGE_CODE = MyBase.Absx1.txtFor("PAGE_CODE").Text.Trim

        Dim sql As String = String.Empty

        MyBase.EnforceConstraints(False)

        Dim SB As New Text.StringBuilder
        SB.Length = 0
        SB.AppendLine("SELECT")
        SB.AppendLine("WBTPAGED.PAGE_CODE,")
        SB.AppendLine("WBTPAGED.STYLE_CODE,")
        SB.AppendLine("ICTSTYL1.STYLE_DESC,")
        SB.AppendLine("MIN(WBTSTYLD.STYLE_STATUS) AS STYLE_STATUS,")
        SB.AppendLine("0 AS CURR_ON_HAND")
        SB.AppendLine("FROM WBTPAGED, WBTSTYLD, ICTSTYL1")
        SB.AppendLine("WHERE WBTPAGED.STYLE_CODE = WBTSTYLD.STYLE_CODE")
        SB.AppendLine("AND WBTSTYLD.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        SB.AppendLine("AND WBTPAGED.PAGE_CODE = '" & PAGE_CODE & "'")
        SB.AppendLine("GROUP BY")
        SB.AppendLine("WBTPAGED.PAGE_CODE,")
        SB.AppendLine("WBTPAGED.STYLE_CODE,")
        SB.AppendLine("ICTSTYL1.STYLE_DESC")
        Call Fill_Records("WBTPAGED", String.Empty, True, SB.ToString)

        SB.Length = 0
        SB.AppendLine("SELECT *")
        SB.AppendLine("FROM WBTPAGU1")
        SB.AppendLine("WHERE PAGE_CODE = '" & PAGE_CODE & "'")
        Call Fill_Records("WBTPAGU1", String.Empty, True, SB.ToString)
        setWBTPAGU1()

        SB.Length = 0
        SB.AppendLine("SELECT *")
        SB.AppendLine("FROM WBTPAGU2")
        SB.AppendLine("WHERE PAGE_CODE = '" & PAGE_CODE & "'")
        Call Fill_Records("WBTPAGU2", String.Empty, True, SB.ToString)

        UpdateInventory()

        grdWBTPAGED.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdWBTPAGED.DisplayLayout.Bands(0).SortedColumns.Add("STYLE_CODE", False)

        grdWBTPAGU2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdWBTPAGU2.DisplayLayout.Bands(0).SortedColumns.Add("ATTR_CODE", False)



    End Sub

    Private Sub UpdateInventory()
        For Each rowWBTPAGED As DataRow In dst.Tables("WBTPAGED").Select()
            Dim STYLE_CODE As String = rowWBTPAGED.Item("STYLE_CODE").ToString & String.Empty
            Dim filter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Select(filter).FirstOrDefault
            If Not IsNothing(rowICTSTYC1) Then
                rowWBTPAGED.Item("CURR_ON_HAND") = Val(rowICTSTYC1.Item("MSOH").ToString & String.Empty) + Val(rowICTSTYC1.Item("MSFT").ToString & String.Empty)
            End If
        Next
        grdWBTPAGED.UpdateData()
        grdWBTPAGED.Refresh()
    End Sub

    Overrides Sub Clear_Record_Special()

        If ScreenMode Then
            MyBase.EnforceConstraints(False)
            dst.Tables("WBTPAGED").Rows.Clear()
            dst.Tables("WBTPAGU1").Rows.Clear()
            dst.Tables("WBTPAGU2").Rows.Clear()
            txtSTYLE_CLASS_CODE.Text = ""
            txtTHEME_CODE.Text = ""
            txtSTYLE_CLASS_CODE2.Text = ""
            MyBase.EnforceConstraints(True)

            PAGE_CODE = String.Empty
        End If

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        With grdWBTPAGU2.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
        End With
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWBTPAGU2, "SSB", "Show Filter", "Show GroupBox", "Add Attributes")
        Load_Popup_Menu(grdWBTPAGED, "SSB", "Show Filter", "Show GroupBox", "Clear All", "Add All To Full Upload", "Remove All From Full Upload")
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

            Case "grdWBTPAGU2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Attributes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdWBTPAGED"
                tlb_btn = DirectCast(tlb_pop.Tools("Clear All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

                tlb_btn = DirectCast(tlb_pop.Tools("Add All To Full Upload"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

                tlb_btn = DirectCast(tlb_pop.Tools("Remove All From Full Upload"), UltraWinToolbars.ButtonTool)
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
                If grd.Name = "grdWBTPAGU2" Then
                    Add_Attributes(grdWBTPAGU2, "ICTATTR1", "ATTR_CODE", "Attributes")
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
                    dst.Tables.Item("WBTPAGED").Clear()
                End If
            Case "Add All To Full Upload"
                Me.Cursor = Cursors.WaitCursor
                dst.Tables.Item("WBTSTYLD").Clear()
                For Each rowWBTPAGED As DataRow In dst.Tables("WBTPAGED").Select()
                    Dim STYLE_CODE As String = rowWBTPAGED.Item("STYLE_CODE").ToString & String.Empty
                    Dim TEMP_SEL As String = String.Format("SELECT * FROM WBTSTYLD WHERE STYLE_CODE = '{0}'", STYLE_CODE)
                    Fill_Records("WBTSTYLD", {STYLE_CODE}, False, TEMP_SEL)
                Next
                For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
                    rowWBTSTYLD.Item("FULL_UPLOAD") = "1"
                Next
                Me.Cursor = Cursors.Default
                MsgBox("Done", vbOKOnly, "Marked As Full")
            Case "Remove All From Full Upload"
                dst.Tables.Item("WBTSTYLD").Clear()
                For Each rowWBTPAGED As DataRow In dst.Tables("WBTPAGED").Select()
                    Dim STYLE_CODE As String = rowWBTPAGED.Item("STYLE_CODE").ToString & String.Empty
                    Dim TEMP_SEL As String = String.Format("SELECT * FROM WBTSTYLD WHERE STYLE_CODE = '{0}'", STYLE_CODE)
                    Fill_Records("WBTSTYLD", {STYLE_CODE}, False, TEMP_SEL)
                Next
                For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
                    rowWBTSTYLD.Item("FULL_UPLOAD") = "0"
                Next
                Me.Cursor = Cursors.Default
                MsgBox("Done", vbOKOnly, "Un-Marked As Full")
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
        Dim sql_where As String = Get_List_of_Codes(TABLE_NAME & "." & COLUMN_NAME & " not in", "WBTPAGU2", COLUMN_NAME)
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
                Dim rowWBTPAGU2 As DataRow = dst.Tables.Item("WBTPAGU2").NewRow
                rowWBTPAGU2.Item("PAGE_CODE") = Absx1.txtFor("PAGE_CODE").Text
                rowWBTPAGU2.Item("ATTR_CODE") = ATTR_CODE
                dst.Tables.Item("WBTPAGU2").Rows.Add(rowWBTPAGU2)
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
                iMSG.AppendLine("This Will Add All Web Styles")
                iMSG.AppendLine("To This Page With Class Code")
                iMSG.AppendLine(String.Format("Of {0}.", txtSTYLE_CLASS_CODE.Text))
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    ADD_TO_WBTPAGED(1)
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
                iMSG.AppendLine("This Will Add All Web Styles")
                iMSG.AppendLine("To This Page With Theme Code")
                iMSG.AppendLine(String.Format("Of {0}.", txtTHEME_CODE.Text))
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    ADD_TO_WBTPAGED(2)
                End If
            End If
        End If
    End Sub

    Private Sub btnATTR_ADD_Click(sender As Object, e As EventArgs) Handles btnATTR_ADD.Click
        If (EntryMode = "Edit" Or EntryMode = "New") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Add By Attribute Code"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}

            If dst.Tables.Item("WBTPAGU2").Rows.Count = 0 Then
                iMSG.AppendLine("Missing Attribute Codes")
                MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            Else
                iMSG.AppendLine("This Will Add All Web Styles")
                iMSG.AppendLine("To This Page With Selected")
                iMSG.AppendLine("Class Code (if Any) and Attributes.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    ADD_TO_WBTPAGED(3)
                End If
            End If
        End If
    End Sub

    Private Sub btnBuildOnRules_Click(sender As Object, e As EventArgs) Handles btnBuildOnRules.Click
        If (EntryMode = "Edit" Or EntryMode = "New") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Add By All Rules"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("This Will Add All Web Styles")
            iMSG.AppendLine("To This Page With All Rules")
            iMSG.AppendLine("Defined.")
            iMSG.AppendLine("")
            iMSG.AppendLine("Is That What You Want?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                If txtSTYLE_CLASS_CODE.Text.Length > 0 Then
                    ADD_TO_WBTPAGED(1)
                End If
                If txtTHEME_CODE.Text.Length > 0 Then
                    ADD_TO_WBTPAGED(2)
                End If
                If dst.Tables.Item("WBTPAGU2").Rows.Count > 0 Then
                    ADD_TO_WBTPAGED(3)
                End If
            End If
        End If
    End Sub

    Private Sub grdWBTPAGED_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWBTPAGED.BeforeRowUpdate
        e.Row.Cells("PAGE_CODE").Value = PAGE_CODE

        Dim STYLE_CODE As String = (e.Row.Cells("STYLE_CODE").Value & String.Empty).ToString.ToUpper.Trim
        Dim styleInList As Boolean = dst.Tables("WBTPAGED").Select("STYLE_CODE = '" & STYLE_CODE & "'").Length > 0
        e.Row.Cells("STYLE_CODE").Value = STYLE_CODE

        Dim rowWBTSTYLD As DataRow = ASCDATA1.GetDataRow("SELECT * FROM WBTSTYLD WHERE STYLE_CODE = :PARM1", "V", New String() {STYLE_CODE})

        If rowWBTSTYLD Is Nothing Then
            e.Cancel = True
            MessageBox.Show("Invalid or missing Style Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If e.Row.IsAddRow AndAlso styleInList Then
                e.Cancel = True
                MessageBox.Show("Style already belongs to the Page.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf Not e.Row.IsAddRow AndAlso styleInList AndAlso
                e.Row.Cells("STYLE_CODE").Value & String.Empty <> e.Row.Cells("STYLE_CODE").OriginalValue & String.Empty Then
                e.Cancel = True
                MessageBox.Show("Style already belongs to the Page.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowWBTSTYLD.Item("STYLE_CODE") & String.Empty)
                e.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
                e.Row.Cells("STYLE_STATUS").Value = rowICTSTYL1.Item("STYLE_STATUS") & String.Empty
            End If
        End If
    End Sub

    Private Sub grdWBTPAGED_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWBTPAGED.ClickCellButton
        If grdWBTPAGED.ActiveRow Is Nothing Then Exit Sub
        If Not e.Cell.Row.IsAddRow Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Dim sql_where As String = ""
        Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
        Dim VIEW_NAME As String = String.Empty

        Select Case grd.ActiveCell.Column.Key
            Case "STYLE_CODE"
                sql_where = String.Format(" WBTSTYLH.STYLE_CODE NOT IN (SELECT STYLE_CODE FROM WBTPAGED WHERE PAGE_CODE = '{0}')", PAGE_CODE)
                VIEW_NAME = "STYLE_WEB.WBTSTYLH"

        End Select

        Call grdClickCellButton(grdWBTPAGED, sql_where, True)
    End Sub

    Private Sub grdWBTPAGED_BeforeCellActivate(sender As Object, e As CancelableCellEventArgs) Handles grdWBTPAGED.BeforeCellActivate
        If Not e.Cell.Row.IsAddRow Then
            e.Cancel = True
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub ADD_TO_WBTPAGED(ByVal AddType As Integer, Optional ShowErr As Boolean = True)
        '1 = Class
        '2 = Theme
        '3 = Class / Attribute
        '4 = Discontinued
        Dim err As New Text.StringBuilder With {.Length = 0}
        Dim sql As New Text.StringBuilder With {.Length = 0}
        Dim sqlin As New Text.StringBuilder With {.Length = 0}
        Dim sqlgb As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine(String.Format("'{0}' AS PAGE_CODE,", Absx1.txtFor("PAGE_CODE").Text))
        sql.AppendLine("WD.STYLE_CODE,")
        sql.AppendLine("S1.STYLE_DESC,")
        sql.AppendLine("WD.STYLE_STATUS")
        sql.AppendLine("FROM WBTSTYLD WD, ICTSTYL1 S1")
        sql.AppendLine("WHERE WD.STYLE_CODE = S1.STYLE_CODE ")

        sqlgb.AppendLine(" GROUP BY WD.STYLE_CODE, S1.STYLE_DESC, WD.STYLE_STATUS")

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
                    If dst.Tables.Item("WBTPAGU2").Rows.Count = 0 Then
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
                    For Each rowWBTPAGU2 As DataRow In dst.Tables("WBTPAGU2").Select()
                        ATIN = ATIN & "'" & rowWBTPAGU2.Item("ATTR_CODE") & "',"
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
            Dim tblWBTSTYLD As DataTable = ASCDATA1.GetDataTable(SQLF, String.Empty)
            For Each rowWBTSTYLD As DataRow In tblWBTSTYLD.Rows
                Dim filter As String = String.Format("STYLE_CODE = '{0}'", rowWBTSTYLD.Item("STYLE_CODE").ToString & String.Empty)
                If dst.Tables.Item("WBTPAGED").Select(filter).Count = 0 Then
                    Dim newWBTPAGU2 As DataRow = dst.Tables.Item("WBTPAGED").NewRow
                    For Each COL As String In New String() {"PAGE_CODE", "STYLE_CODE", "STYLE_DESC", "STYLE_STATUS"}
                        newWBTPAGU2.Item(COL) = rowWBTSTYLD.Item(COL).ToString & String.Empty
                    Next
                    dst.Tables.Item("WBTPAGED").Rows.Add(newWBTPAGU2)
                End If
            Next
        End If
    End Sub

    Private Sub saveWBTPAGU1()
        If dst.Tables.Item("WBTPAGU1").Rows.Count = 1 Then
            Dim rowWBTPAGU1 As DataRow = dst.Tables.Item("WBTPAGU1").Rows(0)
            rowWBTPAGU1.Item("STYLE_CLASS_CODE") = txtSTYLE_CLASS_CODE.Text
            rowWBTPAGU1.Item("THEME_CODE") = txtTHEME_CODE.Text
            rowWBTPAGU1.Item("STYLE_CLASS_CODE2") = txtSTYLE_CLASS_CODE2.Text
        End If
    End Sub

    Private Sub setWBTPAGU1()
        If dst.Tables.Item("WBTPAGU1").Rows.Count = 0 Then
            Dim rowWBTPAGU1 As DataRow = dst.Tables.Item("WBTPAGU1").NewRow
            rowWBTPAGU1.Item("PAGE_CODE") = PAGE_CODE
            dst.Tables.Item("WBTPAGU1").Rows.Add(rowWBTPAGU1)
        End If
        txtSTYLE_CLASS_CODE.Text = dst.Tables.Item("WBTPAGU1").Rows(0).Item("STYLE_CLASS_CODE").ToString & String.Empty
        txtTHEME_CODE.Text = dst.Tables.Item("WBTPAGU1").Rows(0).Item("THEME_CODE").ToString & String.Empty
        txtSTYLE_CLASS_CODE2.Text = dst.Tables.Item("WBTPAGU1").Rows(0).Item("STYLE_CLASS_CODE2").ToString & String.Empty
    End Sub

    Private Sub WBTPAGEH_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "WBTPAGU1", "*", 1, True, String.Empty, 1)
            Create_TDA(.Tables.Add, "WBTPAGU2", "*", 2, True, String.Empty, 2)

            Create_TDA(.Tables.Add, "WBTPAGED", "*", 2, True, String.Empty, 2)
            With dst.Tables("WBTPAGED")
                .Columns.Add("STYLE_DESC", GetType(System.String))
                .Columns.Add("STYLE_STATUS", GetType(System.String))
                .Columns.Add("CURR_ON_HAND", GetType(System.Int64))
            End With

            'Create_TDA(.Tables.Add, "WBTSTYLD", "*", 2, True, String.Empty, 2)

            Dim SB As New Text.StringBuilder With {.Length = 0}
            SB.AppendLine("SELECT * FROM WBTSTYLD WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SB.ToString
            Call Create_TDA(dst.Tables.Add, "WBTSTYLD", "**", 1, True, "V",, "FULL_UPLOAD")

            SB.Length = 0
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
            Create_TDA(dst.Tables.Add, "ICTSTYC1", "**", 0, False, "", 2)
            Fill_Records("ICTSTYC1")
        End With

        grdWBTPAGED.DataSource = dst.Tables("WBTPAGED")
        grdWBTPAGU2.DataSource = dst.Tables("WBTPAGU2")
        Create_Summary(grdWBTPAGED, "STYLE_CODE", "Count")

        Create_Summary(grdWBTPAGU2, "ATTR_CODE", "Count")

        With grdWBTPAGED.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Appearance.ForeColor = System.Drawing.Color.White
            .Columns("STYLE_CODE").Header.Appearance.BackColor2 = System.Drawing.Color.Blue
            .Columns("STYLE_CODE").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        End With

        With grdWBTPAGU2.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
        End With

        'With grdWBTPAGED.DisplayLayout
        '    '.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
        '    '.Override.AllowDelete = DefaultableBoolean.True
        '    '.Override.AllowUpdate = DefaultableBoolean.True
        '    .Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        '    'For i As Integer = 0 To .Bands(0).Columns.Count - 1
        '    '    .Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        '    'Next i
        'End With
    End Sub

    Private Sub btnDiscontinued_Click(sender As Object, e As EventArgs) Handles btnDiscontinued.Click
        ADD_TO_WBTPAGED(4)
    End Sub
#End Region
End Class