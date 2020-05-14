Public Class SOTCOMMD

#Region "Overrides"

    Overrides Sub Show_Record_Special()

        Dim COMM_TYPE As String = Absx1.txtFor("COMM_TYPE").Text

        For I As Integer = 0 To 9
            Absx1.numFor("COMM_DISC_" & Format(I, "0")).Visible = (COMM_TYPE = "REG" Or COMM_TYPE = "PVC")
        Next
        lblDisc.Visible = (COMM_TYPE = "REG" Or COMM_TYPE = "PVC")

        Select Case COMM_TYPE
            Case "REG", "PVC"
                lblDesc0.Text = ""
                lblDesc1.Text = ""
                lblDesc2.Text = ""
                lblDesc3.Text = ""
                lblDesc4.Text = ""

            Case "BTB"

                lblDesc0.Text = ""
                lblDesc1.Text = "Over 7% Discount off Std Price"
                lblDesc2.Text = "Up to 7% Discount off Std Price"
                lblDesc3.Text = ""
                lblDesc4.Text = ""

            Case "DIS"

                lblDesc0.Text = "Over 90% Discount"
                lblDesc1.Text = "Up to 90% Discount"
                lblDesc2.Text = "Up to 80% Discount"
                lblDesc3.Text = "Up to 75% Discount"
                lblDesc4.Text = "Up to 70% Discount"

            Case "PRO"

                lblDesc0.Text = ""
                lblDesc1.Text = "Net Price"
                lblDesc2.Text = ""
                lblDesc3.Text = ""
                lblDesc4.Text = ""

        End Select

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        If Not tf Then
            lblDesc0.Text = ""
            lblDesc1.Text = ""
            lblDesc2.Text = ""
            lblDesc3.Text = ""
            lblDesc4.Text = ""
        End If
    End Sub
#End Region

    Private Sub UltraLabel14_Click(sender As System.Object, e As System.EventArgs) Handles lblTier1.Click

    End Sub
End Class