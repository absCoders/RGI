Public Class ICTSGRP1
    '
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If ASCMAIN1.CLIENT = "RGI" Then
            lblGROUP_CODE.Text = "Family"
            txtSTYLE_GROUP_CODE.ReadOnly = True
            btnNextCode.Visible = True
        Else
            txtSTYLE_GROUP_CODE.ReadOnly = False
            btnNextCode.Visible = False
        End If
    End Sub
    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        btnNextCode.Enabled = Not tf
    End Sub

    Private Sub btnNextCode_Click(sender As Object, e As EventArgs) Handles btnNextCode.Click
        txtSTYLE_GROUP_CODE.Text = ASCMAIN1.Next_Control_No("ICTSGRP1.STYLE_GROUP_CODE")
    End Sub
End Class