Public Class ECTECOM1
    Overrides Sub Show_Record_Special()
        If EntryMode = "New" Then
            SetNewDefaults()
        End If
    End Sub

    Private Sub SetNewDefaults()
        Absx1.numFor("ECOM_MIN_QTY_DEFAULT").Value = 4
        Absx1.numFor("ECOM_ALLOC_PCT_DEFAULT").Value = 100
        Absx1.numFor("ECOM_SHIP_WINDOW").Value = 7
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Update"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Dim CUST_CODE_CNT As Int16 = 0
                If CUST_CODE.Length > 0 Then
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine(String.Format("SELECT COUNT(*) FROM ARTCUST1 WHERE CUST_CODE = '{0}'", CUST_CODE))
                    ASCMAIN1.sql = SQLS.ToString()
                    CUST_CODE_CNT = Val(ASCDATA1.GetDataValue)
                    If CUST_CODE_CNT <> 1 Then
                        EMsg &= "Invalid Value Specified for Cust Code"
                    End If
                Else
                    EMsg &= "Invalid Value Specified for Cust Code"
                End If
        End Select
    End Sub

    Private Sub chkEDI_846_INDICATOR_CheckedChanged(sender As Object, e As EventArgs) Handles chkEDI_846_INDICATOR.CheckedChanged
        If chkEDI_846_INDICATOR.Checked Then
            'lblEDI_846_INTERVAL.Visible = True
            'numEDI_846_INTERVAL.Visible = True
        Else
            'lblEDI_846_INTERVAL.Visible = False
            'numEDI_846_INTERVAL.Visible = False
        End If
    End Sub

    Private Sub UltraLabel11_Click(sender As Object, e As EventArgs) Handles UltraLabel11.Click

    End Sub

    Private Sub UltraNumericEditor3_ValueChanged(sender As Object, e As EventArgs) Handles UltraNumericEditor3.ValueChanged

    End Sub
End Class