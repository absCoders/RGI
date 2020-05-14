Imports System.Drawing
Imports System.Text

Public Class WBTSTYL1

    Private WB_PARM_IMAGES_DIR As String
    Private WB_PARM_IMAGES_UPLOADED_DIR As String
    Private STYLE_CODE As String = String.Empty
    Private WB_PARM_MAX_REC_STYLE As Int16 = 0
    Private CurrentImage As String = String.Empty

    Private Sub WBTSTYL1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim SB As New StringBuilder
        With dst
            Get_PARM("WBTPARM1")
            WB_PARM_IMAGES_DIR = (ROWs("WBTPARM1").Item("WB_PARM_IMAGES_DIR") & String.Empty).ToString.Trim
            WB_PARM_IMAGES_UPLOADED_DIR = String.Empty

            If WB_PARM_IMAGES_DIR.Length > 0 Then
                If Not WB_PARM_IMAGES_DIR.EndsWith("\") Then WB_PARM_IMAGES_DIR &= "\"
                WB_PARM_IMAGES_UPLOADED_DIR = WB_PARM_IMAGES_DIR & "Uploaded\"
                If Not My.Computer.FileSystem.DirectoryExists(WB_PARM_IMAGES_DIR) Then
                    WB_PARM_IMAGES_DIR = String.Empty
                End If
            End If

            'ASCMAIN1.sql = "Select '0' SELECTED, ICTFEAT1.* FROM ICTFEAT1"
            'Create_TDA(.Tables.Add, "ICTFEAT1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select '0' SELECTED, WBTPAGE1.* FROM WBTPAGE1 WHERE NVL(PAGE_STATUS, 'A') = 'A'"
            Create_TDA(.Tables.Add, "WBTPAGE1", "**", 0, False, "", 2)

            'Create_TDA(.Tables.Add, "WBTITEM1", "*")
            'Create_TDA(.Tables.Add, "ICTSTYL2", "*")
            Create_TDA(.Tables.Add, "WBTSTYL3", "*")

            'With dst.Tables("ICTITEM1")
            '    .Columns.Add("SIZE_SEQ_NO", GetType(System.Int16))
            '    .Columns.Add("WHSE_QTY_ON_HAND", GetType(System.Int16))
            'End With

            SB.Length = 0
            SB.AppendLine("Select ICTSTYLR.*,  ICTSTYL1.STYLE_DESC, ICTSTYLR.STYLE_CODE_REC STYLE_CODE_REC_ORIG, ICTSTYLR.SEQ_NO SEQ_NO_ORIG, WBTSTYL1.STYLE_STATUS")
            SB.AppendLine("FROM ICTSTYLR, WBTSTYL1, ICTSTYL1")
            SB.AppendLine("WHERE ICTSTYLR.STYLE_CODE_REC = WBTSTYL1.STYLE_CODE")
            SB.AppendLine("AND WBTSTYL1.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            ASCMAIN1.sql = SB.ToString
            Create_TDA(.Tables.Add, "ICTSTYLR", "**", 1)

            'Dim sql As String
            'sql = "SELECT * FROM ASTAUDT1"
            'sql &= " WHERE TABLE_NAME IN ('ICTITEM1', 'ICTIREC2')"
            'sql &= " AND KEY_VALUE LIKE :PARM1 || ':%'"
            'sql &= " AND COLUMN_NAME = 'BIN_LOC'"
            'Create_TDA(.Tables.Add, "ASTAUDTX", sql, 0, False, "V", 0)
        End With

        grdWBTSTYL3.DataSource = dst.Tables("WBTPAGE1")

        grdICTSTYLR.DataSource = dst.Tables("ICTSTYLR")

        Create_Lookup("WBTPAGE1")
        Create_Lookup("WBTSTYL1")
        Create_Lookup("ICTCOLR1")
        Create_Lookup("ICTSIZE1")

        Get_PARM("WBTPARM1")
        If ROWs("WBTPARM1") IsNot Nothing Then
            WB_PARM_MAX_REC_STYLE = Val(ROWs("WBTPARM1").Item("WB_PARM_MAX_REC_STYLE") & String.Empty)
        End If

        For Each txtctl As Infragistics.Win.UltraWinEditors.UltraTextEditor In _
            New Infragistics.Win.UltraWinEditors.UltraTextEditor() {txtSTYLE_IMAGE, txtSTYLE_IMAGE_OTHER1, _
                        txtSTYLE_IMAGE_OTHER2}

            If txtctl.Name <> txtSTYLE_IMAGE.Name Then
                Dim btn As New UltraWinEditors.EditorButton
                btn.Key = "Open"
                txtctl.ButtonsRight.Add(btn)
                btn.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "ARROW_UP_BLUE")
            End If

            Dim btn1 As New UltraWinEditors.EditorButton
            btn1.Key = "View"
            txtctl.ButtonsRight.Add(btn1)
            btn1.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "VIEW")

            AddHandler txtctl.EditorButtonClick, AddressOf txt_EditorButtonClick
        Next

        'If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("WM") Then
        '    Dim tabKeys As String = ",Sizes,Bin Changes,Images,".ToUpper
        '    For Each tabp As Infragistics.Win.UltraWinTabControl.UltraTab In tabStyle.Tabs
        '        If Not tabKeys.Contains("," & tabp.Key.ToUpper & ",") Then
        '            tabp.Visible = False
        '        End If
        '    Next
        'End If
    End Sub

    Private Sub txt_DisplayImage(ByVal txtctl As UltraWinEditors.UltraTextEditor)

        Dim ImageFile As String = txtctl.Text.Trim
        If ImageFile.Length = 0 Then
            Exit Sub
        End If

        CurrentImage = ImageFile
        AutosizeImage()
    End Sub

#Region "Overrides"

    Public Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        MyBase.Proceed_PreReq_Special(eItemKey)

        Select Case eItemKey

            Case "Update"

                MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text = MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text.Replace(vbCrLf, Space(1)).Trim
                MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text = MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text.Replace(Space(2), Space(1)).Trim

                'MyBase.Absx1.txtFor("STYLE_ADDL_KEYWORDS").Text = MyBase.Absx1.txtFor("STYLE_ADDL_KEYWORDS").Text.Replace(vbCrLf, Space(1)).Trim
                'MyBase.Absx1.txtFor("STYLE_ADDL_KEYWORDS").Text = MyBase.Absx1.txtFor("STYLE_ADDL_KEYWORDS").Text.Replace(Space(2), Space(1)).Trim

                dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_FULL_DESC") = MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text
                'dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_ADDL_KEYWORDS") = MyBase.Absx1.txtFor("STYLE_ADDL_KEYWORDS").Text

                Dim wMsg As String = String.Empty
                'If dst.Tables("ICTITEM1").Select("ITEM_STATUS = 'I' AND ISNULL(BIN_LOC, 0) <> 9999", "").Length > 0 Then
                '    wMsg &= vbCr & "There are Inactive items not in Bin Location 9999."
                'End If

                'If dst.Tables("ICTITEM1").Select("ITEM_STATUS = 'A' AND ISNULL(BIN_LOC, 0) = 9999", "").Length > 0 Then
                '    wMsg &= vbCr & "There are Active items in Bin Location 9999."
                'End If

                'If dst.Tables("ICTITEM1").Select("ITEM_STATUS = 'A'", "").Length = 0 _
                '        AndAlso Not "ID".Contains(MyBase.Absx1.optFor("STYLE_STATUS").Value) Then
                '    wMsg &= vbCr & "All Items are Inactive; however, the style is not Inactive or Discontinued."
                'End If

                'If dst.Tables("ICTITEM1").Select("ITEM_STATUS = 'A'", "").Length > 0 _
                '        AndAlso "ID".Contains(MyBase.Absx1.optFor("STYLE_STATUS").Value) Then
                '    wMsg &= vbCr & "Some Items are Active; however, the style is Inactive or Discontinued."
                'End If

                'If optStyleStatus.Value = "A" _
                '        AndAlso Not Absx1.chkFor("STYLE_EXCL_FROM_DATAFEED").Checked _
                '        AndAlso Absx1.txtFor("BRAND_CODE").TextLength > 0 Then
                '    If dst.Tables("ICTITEM1").Select("ISNULL(ITEM_GTIN, '') = ''").Length > 0 Then
                '        wMsg &= vbCr & "Some Items are missing Gtins."
                '    End If

                'End If

                If wMsg.Length > 0 Then
                    If MessageBox.Show(wMsg & vbCr & vbCr & "Update anyway?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.No Then
                        EMsg = "Update cancelled by user."
                    End If
                    'Else
                    '    If MyBase.Absx1.optFor("STYLE_STATUS").Value = "A" Then
                    '        If MyBase.Absx1.optFor("STYLE_PRICE_TYPE").Value = "R" AndAlso Val(MyBase.Absx1.numFor("STYLE_PRICE_OUR").Value & "") <= 0 Then
                    '            EMsg &= Environment.NewLine & "The company name is WebUndies not FreeUndies; therefore, Our Price must be greater $0.00"
                    '        ElseIf MyBase.Absx1.optFor("STYLE_PRICE_TYPE").Value = "S" AndAlso Val(MyBase.Absx1.numFor("STYLE_PRICE_SALE").Value & "") <= 0 Then
                    '            EMsg &= Environment.NewLine & "The company name is WebUndies not FreeUndies; therefore, Sale Price must be greater $0.00"
                    '        End If
                    '    End If
                End If

        End Select
    End Sub

    Public Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        MyBase.Proceed_PreReq(eItemKey)

        MyBase.Absx1.txtFor("STYLE_CODE").Text = MyBase.Absx1.txtFor("STYLE_CODE").Text.ToUpper.Trim

    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim sql As String = "Delete from ICTSTYL2 where STYLE_CODE = '" & STYLE_CODE & "'"
        'ASCDATA1.ExecuteSQL(sql)

        'dst.Tables("ICTSTYL2").Clear()
        'dst.Tables("ICTSTYL2").AcceptChanges()

        'Dim rowICTSTYL2 As DataRow = Nothing
        'For Each rowICTFEAT1 As DataRow In dst.Tables("ICTFEAT1").Select("SELECTED = '1'")
        '    rowICTSTYL2 = dst.Tables("ICTSTYL2").NewRow
        '    rowICTSTYL2.Item("STYLE_CODE") = Absx1.txtFor("STYLE_CODE").Text
        '    rowICTSTYL2.Item("FEATURE_CODE") = rowICTFEAT1.Item("FEATURE_CODE")
        '    dst.Tables("ICTSTYL2").Rows.Add(rowICTSTYL2)
        'Next


        dst.Tables("WBTSTYL3").Clear()
        dst.Tables("WBTSTYL3").AcceptChanges()

        Dim rowWBTSTYL3 As DataRow = Nothing
        For Each rowWBTPAGE1 As DataRow In dst.Tables("WBTPAGE1").Select("SELECTED = '1'")
            rowWBTSTYL3 = dst.Tables("WBTSTYL3").NewRow
            rowWBTSTYL3.Item("STYLE_CODE") = Absx1.txtFor("STYLE_CODE").Text
            rowWBTSTYL3.Item("PAGE_CODE") = rowWBTPAGE1.Item("PAGE_CODE")
            dst.Tables("WBTSTYL3").Rows.Add(rowWBTSTYL3)
        Next


        'For Each tableName As String In New String() {"ICTITEM1", "ICTSTYL2"}
        '    For Each row As DataRow In dst.Tables(tableName).Select("")
        '        If row.RowState = DataRowState.Added Then
        '            Call Write_Audit_Trail(row, Nothing, "N")
        '        ElseIf row.RowState = DataRowState.Modified Then
        '            Call Write_Audit_Trail(row, Nothing, "E")
        '        End If
        '    Next
        'Next

        Dim seqNo As Integer = 1
        For Each rowICTSTYLR As DataRow In dst.Tables("ICTSTYLR").Select("", "SEQ_NO", DataViewRowState.CurrentRows)
            rowICTSTYLR.Item("SEQ_NO") = seqNo
            seqNo += 1
        Next

        'INIT_LAST("ICTITEM1", True, , True)
        'Update_Record_TDA("ICTITEM1")
        'Update_Record_TDA("ICTSTYL2")
        Update_Record_TDA("ICTSTYLR", "DELETE FROM ICTSTYLR WHERE STYLE_CODE = '" & STYLE_CODE & "'")
        Update_Record_TDA("WBTSTYL3", "DELETE FROM WBTSTYL3 WHERE STYLE_CODE = '" & STYLE_CODE & "'")
    End Sub

    Public Overrides Sub Proceed_Update_Special_Post()
        MyBase.Proceed_Update_Special_Post()
        ASCDATA1.ExecuteSQL("UPDATE WBTSTYL1 SET WEB_IND = '1' WHERE STYLE_CODE = :PARM1", "V", STYLE_CODE)
    End Sub

    Public Overrides Sub txt_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs)
        MyBase.txt_EditorButtonClick(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor
        txtctl = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        Dim imageFields As String = ",STYLE_IMAGE,STYLE_IMAGE_OTHER1,STYLE_IMAGE_OTHER2, "

        If imageFields.Contains("," & MyBase.Absx1.GetABSColumnName(txtctl) & ",") Then

            Select Case e.Button.Key

                Case "Open"
                    Using fdlg As OpenFileDialog = New OpenFileDialog()
                        fdlg.Title = "Images Open File Dialog"
                        fdlg.InitialDirectory = WB_PARM_IMAGES_DIR
                        fdlg.Filter = "Image Files|*.jpg;*.gif;*.bmp;*.png;*.jpeg|All Files|*.*"
                        fdlg.FilterIndex = 1
                        fdlg.RestoreDirectory = True
                        If fdlg.ShowDialog() = DialogResult.OK Then
                            If My.Computer.FileSystem.FileExists(fdlg.FileName) Then
                                txtctl.Text = My.Computer.FileSystem.GetName(fdlg.FileName)
                            End If
                        End If
                    End Using

                Case "View"
                    txt_DisplayImage(txtctl)

            End Select
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        STYLE_CODE = MyBase.Absx1.txtFor("STYLE_CODE").Text.Trim

        Dim sql As String = String.Empty

        MyBase.EnforceConstraints(False)

        txtSTYLE_IMAGE.Text = STYLE_CODE.ToLower & ".jpg"
        CurrentImage = txtSTYLE_IMAGE.Text
        AutosizeImage()

        Call Fill_Records("WBTSTYL3", String.Empty, True, "SELECT * FROM WBTSTYL3 WHERE STYLE_CODE = '" & STYLE_CODE & "'")
        Call Fill_Records("ICTSTYLR", STYLE_CODE, )

        'Call Fill_Records("ICTITEM1", String.Empty, True, "SELECT * FROM ICTITEM1 WHERE STYLE_CODE = '" & STYLE_CODE & "'")
        'sql = "SELECT ICTITEM1.*, ICTSIZE1.SIZE_SEQ_NO, NVL(ICTSTAT2.WHSE_QTY_ON_HAND, 0) WHSE_QTY_ON_HAND"
        'sql &= " FROM ICTITEM1, ICTSIZE1, ICTSTAT2"
        'sql &= " WHERE ICTITEM1.STYLE_CODE = '" & STYLE_CODE & "'"
        'sql &= " AND ICTITEM1.SIZE_CODE = ICTSIZE1.SIZE_CODE (+)"
        'sql &= " AND ICTITEM1.STYLE_CODE = ICTSTAT2.STYLE_CODE (+)"
        'sql &= " AND ICTITEM1.COLOR_CODE = ICTSTAT2.COLOR_CODE (+)"
        'sql &= " AND ICTITEM1.SIZE_CODE = ICTSTAT2.SIZE_CODE (+)"
        'Call Fill_Records("ICTITEM1", String.Empty, True, sql)

        'Call Fill_Records("ICTFEAT1")
        Call Fill_Records("WBTPAGE1")
        'Call Fill_Records("ASTAUDTX", New Object() {STYLE_CODE})

        'For Each rowICTSTYL2 As DataRow In dst.Tables("ICTSTYL2").Rows
        '    If dst.Tables("ICTFEAT1").Select("FEATURE_CODE = '" & rowICTSTYL2.Item("FEATURE_CODE") & "'").Length > 0 Then
        '        dst.Tables("ICTFEAT1").Select("FEATURE_CODE = '" & rowICTSTYL2.Item("FEATURE_CODE") & "'")(0).Item("SELECTED") = 1
        '    End If
        'Next
        'dst.Tables("ICTSTYL2").AcceptChanges()

        For Each rowWBTSTYL3 As DataRow In dst.Tables("WBTSTYL3").Rows
            If dst.Tables("WBTPAGE1").Select("PAGE_CODE = '" & rowWBTSTYL3.Item("PAGE_CODE") & "'").Length > 0 Then
                dst.Tables("WBTPAGE1").Select("PAGE_CODE = '" & rowWBTSTYL3.Item("PAGE_CODE") & "'")(0).Item("SELECTED") = 1
            End If
        Next
        dst.Tables("WBTSTYL3").AcceptChanges()

        SetupICTSTYLR()

        MyBase.EnforceConstraints(True)

        With grdWBTSTYL3.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("SELECTED", True)

        End With

        If EntryMode = "New" Then
            'dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_COST") = 0
            dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_IMAGE_OTHER1") = String.Empty
            dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_IMAGE_OTHER2") = String.Empty
        End If

    End Sub

    Overrides Sub Clear_Record_Special()

        If ScreenMode Then
            MyBase.EnforceConstraints(False)
            'dst.Tables("ICTSTYL2").Rows.Clear()
            dst.Tables("WBTSTYL3").Rows.Clear()
            dst.Tables("ICTSTYLR").Rows.Clear()
            'dst.Tables("ICTFEAT1").Rows.Clear()
            dst.Tables("WBTPAGE1").Rows.Clear()
            'dst.Tables("ICTITEM1").Rows.Clear()
            'dst.Tables("ASTAUDTX").Rows.Clear()
            MyBase.EnforceConstraints(True)

            STYLE_CODE = String.Empty
            tabStyle.SelectedTab = tabStyle.Tabs(0)

            CurrentImage = String.Empty
        End If

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        tabStyle.Visible = tf

    End Sub

    Private Sub SetupICTSTYLR()
        If dst.Tables("ICTSTYLR").Select("", "", DataViewRowState.CurrentRows).Length >= WB_PARM_MAX_REC_STYLE Then
            grdICTSTYLR.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        Else
            If grdICTSTYLR.DisplayLayout.Override.AllowAddNew <> UltraWinGrid.AllowAddNew.FixedAddRowOnTop Then
                grdICTSTYLR.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            End If
        End If
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdICTSTYLR_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTSTYLR.AfterRowsDeleted
        SetupICTSTYLR()
    End Sub

    Private Sub grdICTSTYLR_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTSTYLR.AfterRowUpdate
        SetupICTSTYLR()
    End Sub

    Private Sub grdICTSTYLR_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYLR.BeforeRowUpdate

        e.Row.Cells("STYLE_CODE").Value = STYLE_CODE

        EMsg = String.Empty

        If e.Row.Cells("SEQ_NO").Value & String.Empty = String.Empty Then
            e.Row.Cells("SEQ_NO").Value = Val(dst.Tables("ICTSTYLR").Compute("MAX(SEQ_NO)", "") & String.Empty) + 1
        End If

        Dim SEQ_NO As Integer = Val(e.Row.Cells("SEQ_NO").Value & String.Empty)
        If dst.Tables("ICTSTYLR").Select("SEQ_NO = " & SEQ_NO).Length > 0 AndAlso _
            (e.Row.IsAddRow OrElse e.Row.Cells("SEQ_NO").Value <> e.Row.Cells("SEQ_NO_ORIG").Value & String.Empty) Then
            EMsg &= Environment.NewLine & "Sequence Number is already used"
        End If

        e.Row.Cells("STYLE_CODE_REC").Value = (e.Row.Cells("STYLE_CODE_REC").Value & String.Empty).ToString.Trim.ToUpper
        Dim STYLE_CODE_REC As String = e.Row.Cells("STYLE_CODE_REC").Value & String.Empty

        If STYLE_CODE_REC.Length = 0 Then
            EMsg &= Environment.NewLine & "Recommended Style is required"
        ElseIf STYLE_CODE_REC = STYLE_CODE Then
            EMsg &= Environment.NewLine & "Recommended Style must be different than the Style Code"
        Else
            If dst.Tables("ICTSTYLR").Select("STYLE_CODE_REC = '" & STYLE_CODE_REC & "'", "", DataViewRowState.CurrentRows).Length > 0 AndAlso _
                (e.Row.IsAddRow OrElse e.Row.Cells("STYLE_CODE_REC").Value <> e.Row.Cells("STYLE_CODE_REC_ORIG").Value & String.Empty) Then
                EMsg &= Environment.NewLine & "The recommended style code is already in use"
            End If
        End If

        If EMsg.Length = 0 Then
            Dim rowWBTSTYL1 As DataRow = MyBase.LookUp("WBTSTYL1", e.Row.Cells("STYLE_CODE_REC").Value)
            If rowWBTSTYL1 Is Nothing Then
                EMsg &= Environment.NewLine & "Invalid recommended style code"
            Else
                e.Row.Cells("STYLE_DESC").Value = rowWBTSTYL1.Item("STYLE_DESC") & ""
                e.Row.Cells("STYLE_STATUS").Value = rowWBTSTYL1.Item("STYLE_STATUS") & ""
            End If
        End If

        If EMsg.Length > 0 Then
            e.Cancel = True
            MessageBox.Show(EMsg, "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        e.Row.Cells("STYLE_CODE_REC_ORIG").Value = e.Row.Cells("STYLE_CODE_REC").Value
        e.Row.Cells("SEQ_NO_ORIG").Value = e.Row.Cells("SEQ_NO").Value

    End Sub

    Private Sub grdICTSTYLR_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYLR.ClickCellButton

        If grdICTSTYLR.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Dim COLUMN_NAME As String = grdICTSTYLR.ActiveCell.Column.Key
        Dim VIEW_NAME As String = String.Empty

        If COLUMN_NAME = "STYLE_CODE_REC" Then COLUMN_NAME = "STYLE_CODE"

        Call grdClickCellButton(grdICTSTYLR, sql_where, False, COLUMN_NAME, VIEW_NAME)

        If ASCMAIN1.CodeSelector.SelectedRows.Count > 0 Then
            grdICTSTYLR.ActiveRow.Cells("STYLE_DESC").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STYLE_DESC") & ""
            grdICTSTYLR.ActiveRow.Cells("STYLE_STATUS").Value = ASCMAIN1.CodeSelector.SelectedRows(0).Item("STYLE_STATUS") & ""
        End If
    End Sub

    Private Sub grdICTSTYLR_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTSTYLR.InitializeLayout

        Dim SQL As String = "SELECT 'A', 'Active' FROM DUAL"
        SQL &= " UNION SELECT 'I', 'Inactive' FROM DUAL"
        SQL &= " UNION SELECT 'P', 'Pending' FROM DUAL"
        SQL &= " UNION SELECT 'D', 'Discontinued' FROM DUAL"
        ASCMAIN1.Add_Value_List(grdICTSTYLR, "STYLE_STATUS", SQL)
    End Sub

    Private Sub grdWBTSTYL3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWBTSTYL3.InitializeRow

        Dim pageCode As String = e.Row.Cells("PAGE_CODE").Value

        Try
            e.Row.Cells("PAGE_NAME").Value = MyBase.LookUp("WBTPAGE1", pageCode, True).Item("PAGE_NAME") & String.Empty
        Catch ex As Exception

        End Try

    End Sub

    Public Sub AutosizeImage(Optional ByVal pSizeMode As PictureBoxSizeMode = PictureBoxSizeMode.CenterImage)
        Try

            If CurrentImage.Length = 0 Then
                Exit Sub
            End If

            picBox.Image = Nothing

            Select Case optImageSize.Value
                Case "A"
                    picBox.SizeMode = PictureBoxSizeMode.AutoSize
                Case "C"
                    picBox.SizeMode = PictureBoxSizeMode.CenterImage
                Case "N"
                    picBox.SizeMode = PictureBoxSizeMode.Normal
                Case "S"
                    picBox.SizeMode = PictureBoxSizeMode.StretchImage
                Case "Z"
                    picBox.SizeMode = PictureBoxSizeMode.Zoom
                Case Else
                    picBox.SizeMode = PictureBoxSizeMode.AutoSize
            End Select

            Dim displayFile As String = WB_PARM_IMAGES_UPLOADED_DIR & CurrentImage
            If Not My.Computer.FileSystem.FileExists(displayFile) Then
                displayFile = WB_PARM_IMAGES_DIR & CurrentImage
                If Not My.Computer.FileSystem.FileExists(displayFile) Then
                    Exit Sub
                End If
            End If

            If System.IO.File.Exists(displayFile) Then
                Dim imgOrg As Bitmap
                Dim imgShow As Bitmap
                Dim g As Graphics
                Dim divideBy, divideByH, divideByW As Double
                imgOrg = DirectCast(Bitmap.FromFile(displayFile), Bitmap)

                divideByW = imgOrg.Width / picBox.Width
                divideByH = imgOrg.Height / picBox.Height
                If divideByW > 1 Or divideByH > 1 Then
                    If divideByW > divideByH Then
                        divideBy = divideByW
                    Else
                        divideBy = divideByH
                    End If

                    imgShow = New Bitmap(CInt(CDbl(imgOrg.Width) / divideBy), CInt(CDbl(imgOrg.Height) / divideBy))
                    imgShow.SetResolution(imgOrg.HorizontalResolution, imgOrg.VerticalResolution)
                    g = Graphics.FromImage(imgShow)
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(imgOrg, New Rectangle(0, 0, CInt(CDbl(imgOrg.Width) / divideBy), CInt(CDbl(imgOrg.Height) / divideBy)), 0, 0, imgOrg.Width, imgOrg.Height, GraphicsUnit.Pixel)
                    g.Dispose()
                Else
                    imgShow = New Bitmap(imgOrg.Width, imgOrg.Height)
                    imgShow.SetResolution(imgOrg.HorizontalResolution, imgOrg.VerticalResolution)
                    g = Graphics.FromImage(imgShow)
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(imgOrg, New Rectangle(0, 0, imgOrg.Width, imgOrg.Height), 0, 0, imgOrg.Width, imgOrg.Height, GraphicsUnit.Pixel)
                    g.Dispose()
                End If
                imgOrg.Dispose()

                picBox.Image = imgShow
            Else
                picBox.Image = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub UltraOptionSet2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optImageSize.ValueChanged
        AutosizeImage()
    End Sub

#End Region

    Private Sub UltraTextEditor1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraTextEditor1.KeyUp
        If Not ScreenMode Then
            Dim code As String = Absx1.txtFor("STYLE_CODE").Text.ToUpper
            Absx1.txtFor("STYLE_DESC").Text = LookUp("WBTSTYL1", code, True).Item("STYLE_DESC") & String.Empty
        End If
    End Sub

    Function ValidateUPC_EAN(ByVal UPC_EANnumber As String) As Boolean

        Dim originalCheckDigit As String = String.Empty
        Dim checkDigitSubtotal As Int16 = 0

        Try
            UPC_EANnumber = UPC_EANnumber.Trim.Replace(" ", "")
            If Not IsNumeric(UPC_EANnumber) Then
                Return False
            End If

            If UPC_EANnumber.Length = 14 Then
                If UPC_EANnumber.StartsWith("00") Then
                    UPC_EANnumber = UPC_EANnumber.Substring(2)
                ElseIf UPC_EANnumber.StartsWith("0") Then
                    UPC_EANnumber = UPC_EANnumber.Substring(1)
                End If
            End If

            ' Possible valid input includes 12 char UPC, 13 char EAN
            Select Case Len(UPC_EANnumber)
                Case 12
                    ' Let's strip your check digit and calculate ours from scratch, then compare the two
                    originalCheckDigit = UPC_EANnumber.Substring(11, 1)
                    UPC_EANnumber = UPC_EANnumber.Substring(0, 11)

                    ' Now we need to do the UPC A check digit calculation.
                    ' Add up the numbers in the odd positions left to right. Multiple the result by 3.
                    ' Add up the numbers in the even positions. Now add the first subtotal to the second.
                    ' The UPC barcode check digit is the single digit number makes the total a multiple of 10.
                    checkDigitSubtotal = (Val(Microsoft.VisualBasic.Left(UPC_EANnumber, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 3, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 5, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 7, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 9, 1))) _
                                            + (Val(Microsoft.VisualBasic.Right(UPC_EANnumber, 1)))

                    checkDigitSubtotal = (3 * checkDigitSubtotal) _
                                            + (Val(Mid(UPC_EANnumber, 2, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 4, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 6, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 8, 1))) _
                                            + (Val(Mid(UPC_EANnumber, 10, 1)))

                    Dim CalcUPCcheckDigit = Microsoft.VisualBasic.Right(Str(300 - checkDigitSubtotal), 1)
                    Return CalcUPCcheckDigit = originalCheckDigit

                Case 13
                    ' Ean
                    originalCheckDigit = UPC_EANnumber.Substring(12, 1)
                    UPC_EANnumber = UPC_EANnumber.Substring(0, 12)

                    ' Calculate the EAN-13 check digit.
                    checkDigitSubtotal = 3 * (Val(Mid(UPC_EANnumber, 2, 1)) _
                                              + Val(Mid(UPC_EANnumber, 4, 1)) _
                                              + Val(Mid(UPC_EANnumber, 6, 1)) _
                                              + Val(Mid(UPC_EANnumber, 8, 1)) _
                                              + Val(Mid(UPC_EANnumber, 10, 1)) _
                                              + Val(Microsoft.VisualBasic.Right(UPC_EANnumber, 1)))

                    checkDigitSubtotal = checkDigitSubtotal _
                                            + Val(Microsoft.VisualBasic.Left(UPC_EANnumber, 1)) _
                                            + Val(Mid(UPC_EANnumber, 3, 1)) _
                                            + Val(Mid(UPC_EANnumber, 5, 1)) _
                                            + Val(Mid(UPC_EANnumber, 7, 1)) _
                                            + Val(Mid(UPC_EANnumber, 9, 1)) _
                                            + Val(Mid(UPC_EANnumber, 11, 1))

                    Dim CalcEANcheckDigit As String = Microsoft.VisualBasic.Right(Str(300 - checkDigitSubtotal), 1)
                    Return CalcEANcheckDigit = originalCheckDigit

                Case Else
                    Return False
            End Select


        Catch ex As Exception
            Return False

        End Try

    End Function

End Class