Public Class ECTECOMA
    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey
            Case "New"
                Absx1.txtFor("ATTRIB_CODE").Text = ASCMAIN1.Next_Control_No("ECTECOMA.ATTRIB_CODE")
        End Select
    End Sub

    Private Sub txtSTYLE_CODE_PLM_SOURCE_ValueChanged(sender As Object, e As EventArgs)

    End Sub
End Class