Public Class ICTPANTC

    Private Sub txtRGB_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtRGB.ValueChanged
        If txtRGB.Text = "" Then
            txtSwatch.Appearance.BackColor = System.Drawing.Color.Empty
        Else
            Dim RGB() As String = Split(txtRGB.Text & ",,,", ",")
            Dim R As Integer = Val(RGB(0)) : If R < 0 Or R > 255 Then R = 0
            Dim G As Integer = Val(RGB(1)) : If G < 0 Or G > 255 Then G = 0
            Dim B As Integer = Val(RGB(2)) : If B < 0 Or B > 255 Then B = 0
            txtSwatch.Appearance.BackColor = System.Drawing.Color.FromArgb(255, R, G, B)
        End If
    End Sub
End Class