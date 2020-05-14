Public Class ARTPARM1
#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

                If Absx1.chkFor("AR_PARM_USE_DISC").Checked Then
                    If Absx1.txtFor("AR_PARM_HDG_DISC").Text = "" Then
                        EMsg &= vbCr & "Please provide a Heading for Discounts"
                    End If
                    If LookUp("ARTREAS1", Absx1.txtFor("AR_PARM_REASON_CODE_DISC").Text) Is Nothing Then
                        EMsg &= vbCr & "Please provide a Valid Reason Code for Discounts"
                    End If
                End If
                If Absx1.chkFor("AR_PARM_USE_WOFF").Checked Then
                    If Absx1.txtFor("AR_PARM_HDG_WOFF").Text = "" Then
                        EMsg &= vbCr & "Please provide a Heading for Write-Offs"
                    End If
                    If LookUp("ARTREAS1", Absx1.txtFor("AR_PARM_REASON_CODE_WOFF").Text) Is Nothing Then
                        EMsg &= vbCr & "Please provide a Valid Reason Code for Write-Offs"
                    End If
                End If

        End Select
    End Sub

#End Region

End Class