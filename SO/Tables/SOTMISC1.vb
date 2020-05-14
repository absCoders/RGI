Public Class SOTMISC1
    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "Edit"
            Case "Update", "Save"
                If Absx1.txtFor("ACCT_CODE").Value & "" = "" Then
                    EMsg &= vbCr & "You Must Select an Account Code"
                End If
        End Select
    End Sub
End Class