Public Class SOTPICK1

    Dim sqlSOTPICK2 As String = ""
    Dim REM_CUBE As Decimal = 0
    Dim PICK_NO As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "SELECT SOTPICK2.PICK_NO, SOTPICK2.PICK_LNO, SOTPICK2.PICK_QTY" & vbCrLf _
                & ", SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO, SOTPICK2.PICK_QTY_CONF" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.SUB_BODY_CODE, ICTBODY2.STANDARD_CUBE_PER_UNIT" & vbCrLf _
                & " FROM SOTPICK2, SOTORDR2, ICTSTYL1, ICTBODY2" & vbCrLf _
                & "WHERE SOTPICK2.PICK_NO = :PARM1" & vbCrLf _
                & "AND SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "AND ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & "AND ICTBODY2.SUB_BODY_CODE = ICTSTYL1.SUB_BODY_CODE"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, True, "V", 2)
            .Tables("SOTPICK2").Columns.Add("CUBE_REQD", GetType(System.Decimal), "PICK_QTY * STANDARD_CUBE_PER_UNIT")
            .Tables("SOTPICK2").Columns.Add("CART_NO", GetType(System.String))


            Create_TDA(.Tables.Add, "SOTCART1", "*", 0, False)
            .Tables("SOTCART1").Columns.Add("PKG_CUBE", GetType(System.Decimal))
            .Tables("SOTCART1").Columns.Add("PKG_CUBE_PACK", GetType(System.Decimal))

            Create_TDA(.Tables.Add, "SOTCART2", "*", 0, False)

            ASCMAIN1.sql = "Select PKG_CODE, INNER_CUBE from WHTPKGM1_N where USE_FOR_P2L = '1' ORDER BY INNER_CUBE DESC"
            Create_TDA(.Tables.Add, "WHTPKGM1", "**", 0, False)


        End With

        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdWHTPKGM1.DataSource = dst.Tables("WHTPKGM1")

        Fill_Records("WHTPKGM1")
        Sort_grdColumns(grdWHTPKGM1, "INNER_CUBE".ToLower)

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTPICK2.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = Drawing.Color.White
            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            If gcol.Key = "CUBE_REQD" Then
                '  gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                '      gcol.Format = "#.00000"

            End If

        Next


        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICK2, New String() {"PICK_QTY", "CUBE_REQD"})

        'numBuffer.Value = 10
        'numPKGBuffer.Value = 2

        'With grdSOTPICK2.DisplayLayout.Bands(0)
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        If gcol.Key = "CUBE_REQD" Then
        '            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        '            gcol.Format = "#.00000"
        '        End If

        '        gcol.Header.Appearance.BackColor = Drawing.Color.Tomato
        '        gcol.Header.Appearance.BackColor = Drawing.Color.White
        '        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '    Next
        'End With


    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICK2, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

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
        'if not new or edit - hide add codes

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

            Case "grdSOTPICK2"
                'tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
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
            'Case "Add Codes"
            '    If grd.Name = "grdSOTPICK2" Then
            '        Add_Codes(grdSOTPICK2, "ICTSTYL1", "STYLE_CODE", "Items")
            '    End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Show_Record_Special()

        PICK_NO = Absx1.txtFor("PICK_NO").Text

        EnforceConstraints(False)
        Fill_Records("SOTPICK2", New String() {PICK_NO})
        Sort_grdColumns(grdSOTPICK2, "PICK_LNO")
        grdSOTPICK2.Text = "Pick Ticket Details for " & PICK_NO

        Sort_grdColumns(grdSOTCART1, "CART_NO")
        grdSOTCART1.Text = "Cartons for Pick Ticket " & PICK_NO

        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTPICK2").Rows.Clear()
            dst.Tables("SOTCART1").Rows.Clear()
            dst.Tables("SOTCART2").Rows.Clear()
            EnforceConstraints(True)
        End If

        grdSOTPICK2.Text = "Pick Ticket Details"
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        cmdCartonize.Visible = tf
        'Set_Read_Only_for_ctl(numBuffer, False)
        'Set_Read_Only_for_ctl(numPKGBuffer, False)
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTPICK2, grdSOTCART1}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub


    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTCART1.AfterRowActivate
        SETUP_SOTCART2
    End Sub

    Sub Setup_SOTCART2()
        If grdSOTCART1.ActiveRow Is Nothing Then
            grdSOTCART2.Visible = False
        Else

            Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value & ""
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"CART_NO = {CART_NO}"
            grdSOTCART2.Visible = True
        End If
    End Sub

    Private Sub cmdCartonize_Click(sender As Object, e As EventArgs) Handles cmdCartonize.Click
        dst.Tables("SOTCART1").Rows.Clear()
        dst.Tables("SOTCART2").Rows.Clear()


        TAC.SOCMAIN1.Create_Cartons_For_PICK_NO(Me, PICK_NO)


        'dst.Tables("SOTCART1").Rows.Clear()
        'dst.Tables("SOTCART2").Rows.Clear()

        ''  Dim REM_CUBE_PICK As Decimal = Val(dst.Tables("SOTPICK2").Compute("SUM(CUBE_REQD)", "") & "")
        'REM_CUBE = Val(dst.Tables("SOTPICK2").Compute("SUM(CUBE_REQD)", "") & "")

        '' REM_CUBE = REM_CUBE + (REM_CUBE * Val((numBuffer.Value / 100)))


        'Dim row() As DataRow = dst.Tables("WHTPKGM1").Select("", "INNER_CUBE DESC")
        'Dim MAX_PKG As String = row(0).Item("PKG_CODE")
        'Dim MAX_PKG_CUBE As Decimal = Val(row(0).Item("INNER_CUBE"))

        'Dim J As Integer = 0

        'Do
        '    If REM_CUBE <= 0 Then
        '        Exit Do
        '    Else
        '        If REM_CUBE > MAX_PKG_CUBE Then
        '            J = J + 1
        '            Dim CUBE As Decimal = TAC.SOCMAIN1.Create_Carton(Me, MAX_PKG, Val(MAX_PKG_CUBE) - (Val(MAX_PKG_CUBE) * Val((numPKGBuffer.Value / 100))), J)
        '            REM_CUBE = REM_CUBE - CUBE
        '            ' ISSUE CARTON WITH MAX PKG
        '        Else
        '            For Each rowWHTPKGM1 As DataRow In dst.Tables("WHTPKGM1").Select("", "INNER_CUBE")
        '                If Val(rowWHTPKGM1.Item("INNER_CUBE")) - (Val(rowWHTPKGM1.Item("INNER_CUBE")) * Val((numPKGBuffer.Value / 100))) > Val(REM_CUBE) Then
        '                    J = J + 1
        '                    Dim CUBE As Decimal = TAC.SOCMAIN1.Create_Carton(Me, rowWHTPKGM1.Item("PKG_CODE"), Val(rowWHTPKGM1.Item("INNER_CUBE")) - (Val(rowWHTPKGM1.Item("INNER_CUBE")) * Val((numPKGBuffer.Value / 100))), J)
        '                    REM_CUBE = REM_CUBE - CUBE
        '                    Exit For
        '                End If
        '            Next
        '        End If
        '    End If
        'Loop

    End Sub

    Private Sub UltraTextEditor1_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor1.ValueChanged

    End Sub
#End Region
    Function CREATE_CARTON(PKG_CODE As String, PKG_CUBE As Decimal, CART As Integer)

        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        With rowSOTCART1
            .Item("CART_nO") = Format(CART, "00000000000000000000")
            .Item("PKG_CODE") = PKG_CODE
            .Item("PKG_CUBE") = PKG_CUBE
        End With
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

        Dim CARTON_CUBE As Decimal = 0
        Dim CART_LNO As Integer = 0
        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("CART_NO Is NULL", "CUBE_REQD DESC")
            If CARTON_CUBE + Val(rowSOTPICK2.Item("CUBE_REQD")) > PKG_CUBE Then
                Exit For
            Else
                CART_LNO += 1
                Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                With rowSOTCART2
                    .Item("CART_NO") = rowSOTCART1.Item("CART_NO")
                    .Item("CART_LNO") = Format(CART_LNO, "000")
                    .Item("STYLE_CODE") = rowSOTPICK2.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = rowSOTPICK2.Item("COLOR_CODE")
                    .Item("QTY_REL") = rowSOTPICK2.Item("PICK_QTY")
                End With
                dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                CARTON_CUBE = CARTON_CUBE + Val(rowSOTPICK2.Item("CUBE_REQD"))
                rowSOTPICK2.Item("CART_NO") = rowSOTCART1.Item("CART_NO")
                '    REM_CUBE = REM_CUBE - Val(rowSOTPICK2.Item("CUBE_REQD"))
            End If '
        Next
        rowSOTCART1.Item("PKG_CUBE_PACK") = CARTON_CUBE
        Return CARTON_CUBE

    End Function


End Class