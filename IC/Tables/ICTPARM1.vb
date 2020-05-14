Public Class ICTPARM1

    Private Sub ICTPARM1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If ASCMAIN1.Running_in_VS And ASCMAIN1.CLIENT = "NY" And ASCMAIN1.USER_ID = "wjz" Then
            btnJoeFresh.Visible = True

            Create_TDA(dst.Tables.Add, "JF", "*")
            Create_TDA(dst.Tables.Add, "JFC", "*", 0, False, , 1)

            Fill_Records("JFC")
        End If
    End Sub

    Private Sub btnJoeFresh_Click(sender As Object, e As EventArgs) Handles btnJoeFresh.Click

        'Exit Sub

        Dim pictures_only As Boolean = True
        Dim no_pictures As Boolean = False

        Try
            Dim FILENAME As String = ""
            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                openFileDialog1.RestoreDirectory = True
                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    FILENAME = openFileDialog1.FileName
                End If
            End Using

            If FILENAME <> "" Then

                Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
                Dim range As SpreadsheetGear.IRange = Nothing

                Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                Dim folder As String = fi.DirectoryName
                Dim fname As String = fi.Name

                For Each oSheet In oWB.Worksheets

                    If oSheet.Cells(1, 0).Value & "" <> "IMAGE" Then
                        MsgBox("IMAGE")
                    End If

                    range = oSheet.UsedRange

                    Dim rmax As Integer = range.RowCount

                    Dim TYPE As String = ""
                    Dim started As Boolean = False

                    Dim INVALID_STYLES As String = ""

                    'For i As Integer = 0 To oSheet.Shapes.Count - 1
                    '    If oSheet.Shapes(i).Type = SpreadsheetGear.Shapes.ShapeType.Picture Then
                    '        Dim IMAGE As System.Drawing.Image = SpreadsheetGear.Drawing.Image.GetImage(oSheet.Shapes(i).PictureFormat)
                    '    End If
                    'Next

                    If no_pictures Then
                    Else

                        ' Loop through all shapes on the worksheet
                        For Each shape As SpreadsheetGear.Shapes.IShape In oSheet.Shapes
                            ' Ensure we only look at "picture" shapes
                            If shape.Type = SpreadsheetGear.Shapes.ShapeType.Picture Then
                                ' Get cell "under" top-left edge of the shape
                                Dim imageCell As SpreadsheetGear.IRange = shape.TopLeftCell
                                ' Offset one column to the right
                                Dim NYAG_STYLECell As SpreadsheetGear.IRange = imageCell.Offset(0, 1)
                                ' Get sku text
                                Dim NYAG_STYLE As String = NYAG_STYLECell.Value.ToString()
                                NYAG_STYLE = Replace(Replace(Replace(Trim(Split(NYAG_STYLE, " ")(0)), vbCrLf, ""), vbCr, ""), vbLf, "")

                                Dim JF_STYLECell As SpreadsheetGear.IRange = imageCell.Offset(0, 2)
                                ' Get sku text
                                Dim JF_STYLE As String = JF_STYLECell.Value.ToString()
                                JF_STYLE = Replace(Replace(Replace(Trim(Split(JF_STYLE, " ")(0)), vbCrLf, ""), vbCr, ""), vbLf, "")


                                ' Offset one column to the right
                                Dim clrCell As SpreadsheetGear.IRange = imageCell.Offset(0, 4)
                                ' Get sku text
                                Dim clr As String = clrCell.Value.ToString()
                                'Dim rowJFC As DataRow = dst.Tables("JFC").Rows.Find(clr)
                                'If rowJFC IsNot Nothing Then
                                '    clr = rowJFC.Item("SFX")
                                'End If

                                Dim picname As String = NYAG_STYLE & clr
                                ASCMAIN1.sql = "Select STYLE_CODE_PLM from JF" _
                                    & " where WORKBOOK = '" & fname & "'" _
                                    & "   and WORKSHEET = '" & oSheet.Name & "'" _
                                    & "   and NYAG_STYLE = '" & NYAG_STYLE & "'" _
                                    & "   and JF_STYLE= '" & JF_STYLE & "'" _
                                    & "   and COLOR = '" & clr & "'"
                                Dim rowJF As DataRow = ASCDATA1.GetDataRow


                                If rowJF Is Nothing Then
                                    ASCMAIN1.sql = "Select STYLE_CODE_PLM from JF" _
                                   & " where WORKBOOK = '" & fname & "'" _
                                   & "   and WORKSHEET = '" & oSheet.Name & "'" _
                                   & "   and NYAG_STYLE = '" & NYAG_STYLE & "'" _
                                   & "   and JF_STYLE LIKE '" & JF_STYLE & "%'" _
                                   & "   and COLOR = '" & clr & "'"
                                    rowJF = ASCDATA1.GetDataRow
                                End If

                                If rowJF Is Nothing Then
                                    ASCMAIN1.sql = "Select STYLE_CODE_PLM from JF" _
                                   & " where WORKBOOK = '" & fname & "'" _
                                   & "   and WORKSHEET = '" & oSheet.Name & "'" _
                                   & "   and NYAG_STYLE LIKE '" & NYAG_STYLE & "%'" _
                                   & "   and JF_STYLE LIKE '" & JF_STYLE & "%'" _
                                   & "   and COLOR LIKE '" & Trim(clr) & "%'"
                                    rowJF = ASCDATA1.GetDataRow
                                End If

                                If rowJF IsNot Nothing Then
                                    picname = rowJF.Item(0)
                                Else
                                    Stop
                                End If

                                ' OPTION 1:
                                'Get actual image file used for shape (this image's size may not match that displayed on the worksheet).
                                Using image As System.Drawing.Image = SpreadsheetGear.Drawing.Image.GetImage(shape.PictureFormat)
                                    ' Save image with sku filename
                                    ' image.Save(String.Format(folder & "\Pictures\Images\{0}{1}.png", sku, clr), System.Drawing.Imaging.ImageFormat.Png)
                                    image.Save(folder & "\Pictures\Images\" & picname & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
                                End Using

                                ' OPTION 2:
                                ' Pass in the IShape you want to render (this also accepts IRange and IChart objects)
                                Dim imageCreator As SpreadsheetGear.Drawing.Image = New SpreadsheetGear.Drawing.Image(shape)

                                ' Render a Bitmap object
                                'Using bitmap As System.Drawing.Bitmap = imageCreator.GetBitmap()
                                '    ' Save the file
                                '    bitmap.Save(String.Format(folder & "\Pictures\Bitmaps\{0}-{1}.png", sku, clr), System.Drawing.Imaging.ImageFormat.Png)
                                'End Using
                            End If
                        Next
                    End If

            If pictures_only Then
            Else

                Dim r As Integer = 2
                Do While r < rmax



                    Dim NYAG_STYLE As String = Trim(oSheet.Cells(r, 1).Value & "")
                    Dim JF_STYLE As String = Trim(oSheet.Cells(r, 2).Value & "")
                    Dim STYLE_DESC As String = Trim(oSheet.Cells(r, 3).Value & "")
                    '  If NYAG_STYLE = "KO310" Then Stop
                    '  If JF_STYLE.StartsWith("MF8C490015") Then Stop
                    Dim issue As Boolean = False
                    If NYAG_STYLE = "" And JF_STYLE = "" And STYLE_DESC = "" Then
                    Else
                        If NYAG_STYLE = "" Or JF_STYLE = "" Or STYLE_DESC = "" Then
                            MsgBox(NYAG_STYLE & ":" & JF_STYLE & ":" & STYLE_DESC, MsgBoxStyle.OkOnly, "Issue")
                            issue = True
                        End If
                    End If

                    If NYAG_STYLE <> "" Then
                        Dim row As DataRow = dst.Tables("JF").NewRow
                                For C As Integer = 1 To 26
                                    row.Item(C) = Trim(oSheet.Cells(r, C).Value & "")
                                Next
                        row.Item("WORKBOOK") = fname
                        row.Item("WORKSHEET") = oSheet.Name

                        dst.Tables("JF").Rows.Add(row)
                    End If

                    r += 1
                Loop
            End If
                Next

            Update_Record_TDA("JF")

            MsgBox("Workbook " & FILENAME & " has been Loaded", MsgBoxStyle.OkOnly, "Success")
            End If
        Catch ex As Exception
            MsgBox("Error " & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load this Workbook")
        End Try

    End Sub
End Class