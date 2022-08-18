Public Class ICTROYL1
    Dim isRGI As Boolean = False
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        isRGI = ASCMAIN1.CLIENT = "RGI"

        Get_PARM("ICTPARM1")

        If isRGI Then
            With dst
                Create_TDA(.Tables.Add, "ICTROYL2", "*", 1, True)

                ASCMAIN1.sql = "SELECT STYLE_CODE, STYLE_DESC FROM ICTSTYL1 WHERE ROYALTY_CODE = :PARM1"
                Create_TDA(.Tables.Add("ICTSTROY"), "ICTSTYL1", "**", 0, False, "V", 1)
            End With

            grdICTSTROY.DataSource = dst.Tables("ICTSTROY")
            grdICTROYL2.DataSource = dst.Tables("ICTROYL2")

        End If

        SetVisability()

    End Sub

    Overrides Sub Show_Record_Special()
        'Dim txtctl As UltraWinEditors.UltraTextEditor
        'txtctl = Absx1.txtFor("VEND_CODE")
        Clear_Record_Special()
        Load_Report_Form()
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
        End If
    End Sub

    Sub Load_Report_Form()
        Dim ROYALTY_CODE As String = Absx1.txtFor("ROYALTY_CODE").Text

        Sort_grdColumns(grdICTSTROY, "STYLE_CODE")
        Sort_grdColumns(grdICTROYL2, "ROYALTY_BEGIN")

        EnforceConstraints(False)

        If ASCMAIN1.CLIENT = "RGI" Then
            Fill_Records("ICTSTROY", ROYALTY_CODE)
            Fill_Records("ICTROYL2", ROYALTY_CODE)
            setCopyRightImage()
        End If

        EnforceConstraints(True)

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

#Region "grdAPTVENR2"

    Private Sub grdICTROYL2_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs)
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
#End Region
End Class