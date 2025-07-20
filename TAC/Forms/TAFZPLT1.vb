Public Class TAFZPLT1

    Public Rotate As RotateFlipType = RotateFlipType.RotateNoneFlipNone
    Public ResizsTo4by6 As Boolean = False

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
        If ResizsTo4by6 Then
            picPNG.Image = New Bitmap(picPNG.Image, New Size(64 * 8, 64 * 12))
        End If
        picPNG.Image.RotateFlip(Rotate)
    End Sub

End Class