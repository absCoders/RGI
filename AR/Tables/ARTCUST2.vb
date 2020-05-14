Public Class ARTCUST2

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
            rowASFBASE1.Item("CUST_ADDR_STATUS") = "A"
            Dim CUST_ADDR_CODE As String = MyBase.Absx1.txtFor("CUST_ADDR_CODE").Text.Trim
            CUST_ADDR_CODE = ASCMAIN1.Format_Field(CUST_ADDR_CODE, "CUST_ADDR_CODE")
            MyBase.Absx1.txtFor("CUST_ADDR_CODE").Text = CUST_ADDR_CODE
        End If

    End Sub

End Class