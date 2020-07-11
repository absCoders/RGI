Public Class ICTIMAGT

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                'If txtSTYLE_CODE_PLM_SOURCE.Tag & "" = "" Then
                '    EMsg &= vbCr & "You Must Use the Create Style from PLM function to add a New Style"
                'End If
                Absx1.txtFor("IMAGE_CODE").Text = Absx1.txtFor("IMAGE_CODE").Text.ToUpper
                If Absx1.txtFor("IMAGE_CODE").Text.Length < 4 Then
                    EMsg &= vbCr & "Image Code Must Be Between 4-6 Characters"
                End If
            Case "Edit"

            Case "View"

            Case "Update"

        End Select
    End Sub
End Class