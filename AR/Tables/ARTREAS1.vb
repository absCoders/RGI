Public Class ARTREAS1

    Private Sub ARTREAS1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
        Else
            chkShippingViolation.Visible = False
        End If

        grpSegments.Visible = False
    End Sub
End Class