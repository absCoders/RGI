Imports System.Drawing.Printing

Public Class TAFZPLT1

    Public Rotate As RotateFlipType = RotateFlipType.RotateNoneFlipNone
    Public ResizeTo4by6 As Boolean = False

    Public WriteOnly Property zplPNGFilename As String
        Set(value As String)
            Try
                Dim FS As New IO.FileStream(value, IO.FileMode.Open)
                picPNG.Image = Image.FromStream(FS)
                FS.Close()
                FS.Dispose()
            Catch ex As Exception
            End Try
        End Set
    End Property

    Private Sub TAFZPLT1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        If ResizeTo4by6 Then
            picPNG.Image = New Bitmap(picPNG.Image, New Size(64 * 8, 64 * 12))
        End If
        picPNG.Image.RotateFlip(Rotate)
    End Sub

    Private Sub picPNG_DoubleClick(sender As Object, e As EventArgs) Handles picPNG.DoubleClick
        If MessageBox.Show("Do you want to save the label image?", "Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            Exit Sub
        End If

        Using sfd As New SaveFileDialog()
            sfd.Title = "Save Picture"
            sfd.Filter = "PNG Image|*.png|JPEG Image|*.jpg|Bitmap Image|*.bmp"
            sfd.DefaultExt = "png"
            sfd.AddExtension = True

            If sfd.ShowDialog() = DialogResult.OK Then
                Using bmp As New Bitmap(picPNG.Width, picPNG.Height)
                    picPNG.DrawToBitmap(bmp, New Rectangle(0, 0, bmp.Width, bmp.Height))

                    ' Pick format based on extension
                    Select Case IO.Path.GetExtension(sfd.FileName).ToLower()
                        Case ".jpg"
                            bmp.Save(sfd.FileName, Imaging.ImageFormat.Jpeg)
                        Case ".bmp"
                            bmp.Save(sfd.FileName, Imaging.ImageFormat.Bmp)
                        Case Else
                            bmp.Save(sfd.FileName, Imaging.ImageFormat.Png)
                    End Select
                End Using
            End If
        End Using
    End Sub

End Class