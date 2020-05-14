Public Class ICTCOLRN

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                If Absx1.txtFor("COLOR_CODE").Text = "PPK" Then
                    EMsg &= vbCr & "You may not use PPK as a Color Code"
                End If

            Case "Edit"
            Case "Update"
                If Absx1.txtFor("COLOR_CODE").Text = "PPK" Then
                    EMsg &= vbCr & "You may not use PPK as a Color Code"
                End If
        End Select
    End Sub

#End Region

End Class