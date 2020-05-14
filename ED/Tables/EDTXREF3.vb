Public Class EDTXREF3

    Public Overrides Sub txt_EditorButtonClick(sender As Object, e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs)
        MyBase.txt_EditorButtonClick(sender, e)
    End Sub

    Public Overrides Sub txt_EditorButtonClick_Special(txtctl As Infragistics.Win.UltraWinEditors.UltraTextEditor)
        MyBase.txt_EditorButtonClick_Special(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)

            Case "SENDER_ID_QUAL"
                If ASCMAIN1.CodeSelector.Selections = 1 Then
                    MyBase.Absx1.txtFor("SENDER_ID").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("SENDER_ID") & String.Empty
                    MyBase.Absx1.txtFor("SERVICE_LEVEL_3PL").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("SERVICE_LEVEL_3PL") & String.Empty
                End If

        End Select

    End Sub

End Class