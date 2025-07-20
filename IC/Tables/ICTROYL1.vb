Imports Infragistics.Win.UltraWinGrid

Public Class ICTROYL1
    Dim isRGI As Boolean = False
    Dim SQL As New Text.StringBuilder With {.Length = 0}
    Dim LABEL_LOCATION As String = "S:\Archive\royalty\labels\"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        isRGI = ASCMAIN1.CLIENT = "RGI"
        If ASCMAIN1.useUNCPath Then
            LABEL_LOCATION = $"{ASCMAIN1.Folders("SharedRoot")}\Archive\royalty\labels\"
        End If
        Get_PARM("ICTPARM1")

        If isRGI Then
            With dst
                Create_TDA(.Tables.Add, "ICTROYL2", "*", 1, True)
                SQL.Length = 0
                SQL.AppendLine("SELECT S1.STYLE_CODE, S1.STYLE_DESC, S1.STYLE_PRICE, S1.LABEL_TYPE_CODE")
                SQL.AppendLine("FROM ICTSTYL1 S1")
                SQL.AppendLine("WHERE S1.ROYALTY_CODE = :PARM1")
                ASCMAIN1.sql = SQL.ToString
                Create_TDA(.Tables.Add("ICTSTROY"), "ICTSTYL1", "**", 0, False, "V", 1)
                With .Tables("ICTSTROY")
                    .Columns.Add("STYLE_PRICE_CALC", GetType(System.Decimal))
                    .Columns.Add("STYLE_PRICE_ROY", GetType(System.Decimal))
                    .Columns.Add("STYLE_PRICE_NEW", GetType(System.Decimal))
                    .Columns.Add("STYLE_PRICE_VAR", GetType(System.Decimal))
                End With

                SQL.Length = 0
                SQL.AppendLine("SELECT")
                SQL.AppendLine("I1.STYLE_CODE,")
                SQL.AppendLine("I1.STYLE_DESC,")
                SQL.AppendLine("I1.STYLE_STATUS,")
                SQL.AppendLine("I1.VEND_CODE AS SUPPLIER_VEND_CODE,")
                SQL.AppendLine("A1.VEND_NAME AS SUPPLIER_VEND_NAME,")
                SQL.AppendLine("R1.ROYALTY_CODE AS DESIGNER_CODE,")
                SQL.AppendLine("R1.ROYALTY_DESC AS DESIGNER_DESC,")
                SQL.AppendLine("R1.ROYALTY_NAME AS DESIGNER_NAME,")
                SQL.AppendLine("R1.ROYALTY_ADDR1 AS DESIGNER_ADDR1,")
                SQL.AppendLine("R1.ROYALTY_ADDR2 AS DESIGNER_ADDR2,")
                SQL.AppendLine("R1.ROYALTY_CITY AS DESIGNER_CITY,")
                SQL.AppendLine("R1.ROYALTY_STATE AS DESIGNER_STATE,")
                SQL.AppendLine("R1.ROYALTY_ZIP_CODE AS DESIGNER_ZIP_CODE,")
                SQL.AppendLine("AV.VEND_NAME AS AGENT_VEND_NAME,")
                SQL.AppendLine("AV.VEND_ADDR1 AS AGENT_ADDR1,")
                SQL.AppendLine("AV.VEND_ADDR2 AS AGENT_ADDR2,")
                SQL.AppendLine("AV.VEND_CITY AS AGENT_CITY,")
                SQL.AppendLine("AV.VEND_STATE AS AGENT_STATE,")
                SQL.AppendLine("AV.VEND_ZIP_CODE AS AGENT_ZIP_CODE,")
                SQL.AppendLine("I1.STYLE_PRICE")
                SQL.AppendLine("FROM ICTSTYL1 I1, APTVEND1 A1, ICTROYL1 R1, APTVEND1 AV")
                SQL.AppendLine("WHERE NVL(I1.ROYALTY_CODE,'NAN') <> 'NAN'")
                SQL.AppendLine("AND I1.VEND_CODE = A1.VEND_CODE")
                SQL.AppendLine("AND I1.ROYALTY_CODE = R1.ROYALTY_CODE")
                SQL.AppendLine("AND R1.VEND_CODE = AV.VEND_CODE")
                ASCMAIN1.sql = SQL.ToString
                Create_TDA(.Tables.Add, "ICTSTROX", "**", 0, False, "", 1)
                With .Tables("ICTSTROX")
                    .Columns.Add("STYLE_PRICE_CALC", GetType(System.Decimal))
                    .Columns.Add("STYLE_PRICE_ROY", GetType(System.Decimal))
                    .Columns.Add("STYLE_PRICE_NEW", GetType(System.Decimal))
                    .Columns.Add("STYLE_PRICE_VAR", GetType(System.Decimal))
                End With
            End With
            grdICTSTROY.DataSource = dst.Tables("ICTSTROY")
            grdICTROYL2.DataSource = dst.Tables("ICTROYL2")
            grdICTSTROX.DataSource = dst.Tables("ICTSTROX")
        End If

        Create_Summary(grdICTSTROX, "STYLE_CODE", "Count")

        SetVisability()

        Fill_Records("ICTSTROX")
        setListPriceCalc("ICTSTROX")

    End Sub

    Overrides Sub Show_Record_Special()
        'Dim txtctl As UltraWinEditors.UltraTextEditor
        'txtctl = Absx1.txtFor("VEND_CODE")
        If EntryMode <> "Edit" Then
            Clear_Record_Special()
            Load_Report_Form()
        End If

    End Sub

    Private Sub SetVisability()
        lblSTYLE_PREFIX.Visible = Not isRGI
        txtSTYLE_PREFIX.Visible = Not isRGI
        lblROYALTY_PCT.Visible = Not isRGI
        txtROYALTY_PCT.Visible = Not isRGI
        txtROYALTY_NAME.Visible = isRGI
        lblROYALTY_NAME.Visible = isRGI
        txtROYALTY_COMMENTS.Visible = isRGI
        lblROYALTY_COMMENTS.Visible = isRGI
        txtVEND_CODE.Visible = isRGI
        lblVEND_CODE.Visible = isRGI
        txtVEND_NAME.Visible = isRGI
        grdICTROYL2.Visible = isRGI
        grdICTSTROY.Visible = isRGI
        tabICTSTROX.Tabs.Item("All Royalty Styles").Visible = True
        tabICTSTROX.Tabs.Item("Royalty Style Details").Visible = False
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            If ASCMAIN1.CLIENT = "RGI" Then
                dst.Tables("ICTSTROY").Rows.Clear()
                dst.Tables("ICTROYL2").Rows.Clear()
            End If
            EnforceConstraints(True)
            picStryle.Image = Nothing
            picCopyright.Image = Nothing
            tabICTSTROX.Tabs.Item("All Royalty Styles").Visible = True
            tabICTSTROX.Tabs.Item("Royalty Style Details").Visible = False
        Else
            tabICTSTROX.Tabs.Item("All Royalty Styles").Visible = False
            tabICTSTROX.Tabs.Item("Royalty Style Details").Visible = True
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If ASCMAIN1.CLIENT = "RGI" Then
            With grdICTROYL2.DisplayLayout.Bands(0)
                .Columns.Item("ROYALTY_BEGIN").Format = "MM/dd/yy"
                .Columns.Item("ROYALTY_END").Format = "MM/dd/yy"
                .Columns.Item("ROYALTY_PCT").Format = "###,##0.0"

            End With
            With grdICTSTROY.DisplayLayout.Override
                .AllowUpdate = DefaultableBoolean.False
                .AllowAddNew = False
                .AllowDelete = False
            End With
            With grdICTSTROX.DisplayLayout.Override
                .AllowUpdate = DefaultableBoolean.False
                .AllowAddNew = False
                .AllowDelete = False
            End With
        End If
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTSTROY} ', grdICTROYL2, grdICTSTROX
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                If EntryMode = "Edit" Then
                    .AllowUpdate = DefaultableBoolean.True
                Else
                    .AllowUpdate = DefaultableBoolean.False
                End If
                'If EntryMode = "New" Or EntryMode = "Edit" Then
                '    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                '    .AllowDelete = DefaultableBoolean.True
                '    .AllowUpdate = DefaultableBoolean.True
                'Else
                '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                '    .AllowUpdate = DefaultableBoolean.False
                '    .AllowDelete = DefaultableBoolean.False
                'End If

            End With
        Next
    End Sub

    Sub Load_Report_Form()
        Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Text

        Sort_grdColumns(grdICTSTROY, "STYLE_CODE")
        Sort_grdColumns(grdICTROYL2, "ROYALTY_BEGIN")

        EnforceConstraints(False)

        If ASCMAIN1.CLIENT = "RGI" Then
            Fill_Records("ICTSTROY", ROYALTY_CODE)
            Fill_Records("ICTROYL2", ROYALTY_CODE)
            setListPriceCalc("ICTSTROY")
            setCopyRightImage()
            setLabelImage("A")
            setLabelImage("B")
        End If

        EnforceConstraints(True)

    End Sub

    Private Sub setListPriceCalc(ByVal TABLE_NAME As String)
        For Each rowDATA As DataRow In dst.Tables(TABLE_NAME).Select()
            Dim STYLE_CODE As String = rowDATA.Item("STYLE_CODE").ToString & String.Empty
            Dim STYLE_PRICE As Decimal = Val(rowDATA.Item("STYLE_PRICE").ToString & String.Empty)
            Dim STYLE_PRICE_CALC As Decimal = 0
            Dim STYLE_PRICE_ROY As Decimal = 0
            Dim STYLE_PRICE_NEW As Decimal = 0
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            If Not IsNothing(rowICTSTYL1) Then
                STYLE_PRICE_CALC = TAC.ICCMAIN1.Calculate_Style_Price(Me, True, STYLE_CODE, rowICTSTYL1)
                STYLE_PRICE_ROY = TAC.ICCMAIN1.Calculate_Style_Royalty_Markup(Me, STYLE_CODE)
            Else
                STYLE_PRICE_CALC = STYLE_PRICE
            End If
            STYLE_PRICE_NEW = Math.Round(STYLE_PRICE_CALC + STYLE_PRICE_ROY, 1)
            Dim STYLE_PRICE_VAR As Decimal = STYLE_PRICE_NEW - STYLE_PRICE
            rowDATA.Item("STYLE_PRICE_CALC") = STYLE_PRICE_CALC
            rowDATA.Item("STYLE_PRICE_ROY") = STYLE_PRICE_ROY
            rowDATA.Item("STYLE_PRICE_NEW") = STYLE_PRICE_NEW
            rowDATA.Item("STYLE_PRICE_VAR") = STYLE_PRICE_VAR
        Next
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sql As String = ""

        If ASCMAIN1.CLIENT = "RGI" Then
            Update_Record_TDA("ICTROYL2")
        End If
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                If ASCMAIN1.CLIENT = "RGI" Then
                    Dim ROY_ERR As Boolean = False
                    Dim ROYALTY_END_LAST As String = ""
                    Dim ROYALTY_END_MISSING As Int64 = 0
                    For Each rowICTROYL2 As DataRow In dst.Tables("ICTROYL2").Select("", "ROYALTY_BEGIN")
                        ROYALTY_END_LAST = rowICTROYL2.Item("ROYALTY_BEGIN").ToString & String.Empty
                        Dim ROYALTY_BEGIN As Date = CDate(rowICTROYL2.Item("ROYALTY_BEGIN").ToString & String.Empty)
                        If IsDate(ROYALTY_END_LAST) Then
                            If ROYALTY_BEGIN > CDate(ROYALTY_END_LAST) Then
                                ROY_ERR = True
                            End If
                        Else
                            ROYALTY_END_MISSING += 1
                        End If
                    Next
                    If ROYALTY_END_MISSING > 1 Then
                        ROY_ERR = True
                    End If
                    If ROY_ERR Then
                        EMsg &= EMsg & "Please Check Your Royalty Dates."
                    End If
                End If

        End Select

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        MyBase.Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTROX, "SSBB", "Show Filter", "Show GroupBox", "Refresh", "Style Masterfile")
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
            Case "grdICTSTROX"
                tlb_btn = DirectCast(tlb_pop.Tools("Style Masterfile"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = True

                tlb_btn = DirectCast(tlb_pop.Tools("Refresh"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = True
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
        'MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Style Masterfile"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim keys As New Dictionary(Of String, Object)
                keys.Add("STYLE_CODE", STYLE_CODE)
                Context_Launch("Edit", keys, e.Tool.Key, "ICTSTYL1")
            Case "Refresh"
                Fill_Records("ICTSTROX")
                setListPriceCalc("ICTSTROX")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "grdAPTVENR2"

    Private Sub grdICTROYL2_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdICTROYL2.BeforeRowUpdate
        If e.Row.IsAddRow Then
            Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Value
            If ROYALTY_CODE.Length > 0 Then
                e.Row.Cells("ROYALTY_CODE").Value = ROYALTY_CODE
                e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("INIT_DATE").Value = Now()
                e.Row.Cells("LAST_DATE").Value = Now()
            Else
                MsgBox("Error with Vendor Code", vbOKOnly, "Royalty Problem")
                e.Cancel = True
            End If

        Else

        End If
        Dim iMsg As New Text.StringBuilder With {.Length = 0}
        Dim ROYALTY_BEGIN As String = e.Row.Cells("ROYALTY_BEGIN").Text.ToString
        Dim ROYALTY_PCT As String = e.Row.Cells("ROYALTY_PCT").Text.ToString
        If Not IsDate(ROYALTY_BEGIN) Then
            iMsg.AppendLine("Invalid Begin Date.")
        End If
        If IsNumeric(ROYALTY_PCT) Then
            If Val(ROYALTY_PCT) <= 0 Or Val(ROYALTY_PCT) >= 100 Then
                iMsg.AppendLine("Invalid Percentage.")
            End If
        Else
            iMsg.AppendLine("Invalid Percentage.")
        End If
        If iMsg.Length > 0 Then
            MsgBox(iMsg.ToString, vbOKOnly, "Please Fix The Following")
            e.Cancel = True
        End If
    End Sub

    Private Sub btnCopyrightImage_Click(sender As Object, e As EventArgs) Handles btnCopyrightImage.Click

        If ScreenMode Then
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Select Copyright Image"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.JPG)|*.JPG"
            'fd.FilterIndex = 2
            'fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                fd.Dispose()

                If strFileName.Length > 0 Then
                    Dim imgba() As Byte = Nothing
                    picCopyright.Image = ASCMAIN1.Get_Image("C:\", "", True, , , imgba)
                    picCopyright.Image = Nothing
                    Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Text.ToString & String.Empty
                    Dim COPYRIGHT_IMAGE As String = $"ROYALTY_CODE_COPYRIGHT_{ROYALTY_CODE}"
                    Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
                    If Not FOLDER_NAME.EndsWith("\") Then
                        FOLDER_NAME = FOLDER_NAME + "\"
                    End If
                    If System.IO.File.Exists($"{FOLDER_NAME}{COPYRIGHT_IMAGE}.JPG") Then
                        System.IO.File.Delete($"{FOLDER_NAME}{COPYRIGHT_IMAGE}.JPG")
                    End If
                    System.IO.File.Copy(strFileName, $"{FOLDER_NAME}{COPYRIGHT_IMAGE}.JPG")
                    setCopyRightImage()
                End If
            End If
        End If

    End Sub

    Private Sub grdICTSTROY_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTSTROY.AfterRowActivate
        picStryle.Image = Nothing
        picStryle.SizeMode = PictureBoxSizeMode.StretchImage
        If Not IsNothing(grdICTSTROY.ActiveRow) Then
            Dim STYLE_CODE As String = grdICTSTROY.ActiveRow.Cells("STYLE_CODE").Value
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine($"SELECT MIN(COLOR_CODE) FROM ICTSTYC1 WHERE STYLE_CODE = '{STYLE_CODE}'")
            ASCMAIN1.sql = SQLS.ToString()
            Dim COLOR_CODE As String = ASCDATA1.GetDataValue
            Dim IMAGE_NAME As String = STYLE_CODE & "-" & COLOR_CODE

            Dim imgba() As Byte = Nothing
            If IMAGE_NAME <> "" Then
                Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
                picStryle.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
            End If
        End If
    End Sub

    Private Sub setLabelImage(ByVal LABEL_TYPE As String)
        Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Text.ToString & String.Empty
        Select Case LABEL_TYPE
            Case "A"
                picLabelA.Image = Nothing
                picLabelA.SizeMode = PictureBoxSizeMode.StretchImage
                Dim LABEL_IMAGE As String = $"{ROYALTY_CODE}_A"
                Dim imgba() As Byte = Nothing
                picLabelA.Image = ASCMAIN1.Get_Image(LABEL_LOCATION, LABEL_IMAGE, True, , , imgba)
            Case "B"
                picLabelB.Image = Nothing
                picLabelB.SizeMode = PictureBoxSizeMode.StretchImage
                Dim LABEL_IMAGE As String = $"{ROYALTY_CODE}_B"
                Dim imgba() As Byte = Nothing
                picLabelB.Image = ASCMAIN1.Get_Image(LABEL_LOCATION, LABEL_IMAGE, True, , , imgba)
            Case Else
                picLabelA.Image = Nothing
                picLabelB.Image = Nothing
        End Select

    End Sub

    Private Sub btnLabelAImage_Click(sender As Object, e As EventArgs) Handles btnLabelA.Click
        If ScreenMode Then
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Select Label Image"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.JPG)|*.JPG"

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                fd.Dispose()

                If strFileName.Length > 0 Then
                    Dim imgba() As Byte = Nothing
                    picLabelA.Image = ASCMAIN1.Get_Image("C:\", "", True, , , imgba)
                    picLabelB.Image = Nothing
                    Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Text.ToString & String.Empty
                    Dim LABEL_IMAGE As String = $"{ROYALTY_CODE}_A"
                    If System.IO.File.Exists($"{LABEL_LOCATION}{LABEL_IMAGE}.JPG") Then
                        System.IO.File.Delete($"{LABEL_LOCATION}{LABEL_IMAGE}.JPG")
                    End If
                    System.IO.File.Copy(strFileName, $"{LABEL_LOCATION}{LABEL_IMAGE}.JPG")
                    setLabelImage("A")
                End If
            End If
        End If
    End Sub

    Private Sub btnLabelBImage_Click(sender As Object, e As EventArgs) Handles btnLabelB.Click
        If ScreenMode Then
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Select Label Image"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.JPG)|*.JPG"

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                fd.Dispose()

                If strFileName.Length > 0 Then
                    Dim imgba() As Byte = Nothing
                    picLabelB.Image = ASCMAIN1.Get_Image("C:\", "", True, , , imgba)
                    picLabelB.Image = Nothing
                    Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Text.ToString & String.Empty
                    Dim LABEL_IMAGE As String = $"{ROYALTY_CODE}_B"
                    If System.IO.File.Exists($"{LABEL_LOCATION}{LABEL_IMAGE}.JPG") Then
                        System.IO.File.Delete($"{LABEL_LOCATION}{LABEL_IMAGE}.JPG")
                    End If
                    System.IO.File.Copy(strFileName, $"{LABEL_LOCATION}{LABEL_IMAGE}.JPG")
                    setLabelImage("B")
                End If
            End If
        End If
    End Sub
    Private Sub setCopyRightImage()
        picCopyright.Image = Nothing
        picCopyright.SizeMode = PictureBoxSizeMode.StretchImage
        Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Text.ToString & String.Empty
        Dim COPYRIGHT_IMAGE As String = $"ROYALTY_CODE_COPYRIGHT_{ROYALTY_CODE}"
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        'If Not FOLDER_NAME.Equals("\") Then
        '    FOLDER_NAME = FOLDER_NAME + "\"
        'End If
        Dim imgba() As Byte = Nothing
        picCopyright.Image = ASCMAIN1.Get_Image(FOLDER_NAME, COPYRIGHT_IMAGE, True, , , imgba)
    End Sub

    Private Sub grdICTSTROY_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdICTSTROY.ClickCellButton
        Stop
        If e.Cell.Text = "Label" Then
            Dim S As New Text.StringBuilder With {.Length = 0}
            S.AppendLine("SELECT COL1, COL2 FROM TABLE")
            With ASCMAIN1.CodeSelector
                .SQL = S.ToString
                .MultipleSelections = False
                .PreviouslySelectedCodes0 = ""
                .Caption = "Title For Pop-Up"
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
                Dim COL1 As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("COL1") & ""
                'Do Something with the returned Value
            End If
        End If
    End Sub

    Private Sub btnLabelA_Click(sender As Object, e As EventArgs) Handles btnLabelA.Click

    End Sub
#End Region
End Class