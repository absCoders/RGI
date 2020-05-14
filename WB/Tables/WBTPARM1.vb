Public Class WBTPARM1


    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Update"
                'Dim WB_PARM_SITE_IP As String = MyBase.Absx1.txtFor("WB_PARM_SITE_IP").Text.Trim

                'If WB_PARM_SITE_IP.Length > 0 Then
                '    If Not Net.IPAddress.TryParse(WB_PARM_SITE_IP, Nothing) Then
                '        EMsg = "Could not connect to the provided Site IP Address"
                '    End If
                'End If
        End Select
    End Sub
End Class