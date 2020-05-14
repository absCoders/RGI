Public Class ICFSIMG1

    Public Sub New(ByVal image_name As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        PictureBox1.ImageLocation = image_name
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

    End Sub

    Private Sub btnDone_Click(sender As Object, e As EventArgs) Handles btnDone.Click
        Me.Close()
    End Sub
End Class