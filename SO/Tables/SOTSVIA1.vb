Public Class SOTSVIA1

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As System.Windows.Forms.Control, COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "CARRIER_PROD_CODE"
                sql_where = "CARRIER_CODE = '" & MyBase.Absx1.txtFor("CARRIER_CODE").Text & "'"
        End Select

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"

                Dim SHIP_VIA_DESC As String = " " & MyBase.Absx1.txtFor("SHIP_VIA_DESC").Text.Trim & " "
                SHIP_VIA_DESC = StrConv(SHIP_VIA_DESC, VbStrConv.ProperCase)
                SHIP_VIA_DESC = SHIP_VIA_DESC.Replace(" Ups ", "UPS")
                SHIP_VIA_DESC = SHIP_VIA_DESC.Trim
                MyBase.Absx1.txtFor("SHIP_VIA_DESC").Text = SHIP_VIA_DESC
                MyBase.Absx1.txtFor("CARRIER_CODE").Text = MyBase.Absx1.txtFor("CARRIER_CODE").Text.Trim
                MyBase.Absx1.txtFor("CARRIER_MODE").Text = MyBase.Absx1.txtFor("CARRIER_MODE").Text.Trim
                MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text = MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text.Trim
                Validate_Code("CARRIER_CODE", False, False)
                Validate_Code("CARRIER_MODE", False, False)

                'SOTCARR2, CARRIER_CODE, CARRIER_PROD_CODE
                If EMsg.Length > 0 AndAlso MyBase.Absx1.txtFor("CARRIER_CODE").TextLength > 0 _
                    AndAlso MyBase.Absx1.txtFor("CARRIER_PROD_CODE").TextLength > 0 Then
                    If LookUp("SOTCARR2", New String() {MyBase.Absx1.txtFor("CARRIER_CODE").Text, MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text}) Is Nothing Then
                        EMsg &= "Invalid Carrier Product for the supplied Product."
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub Show_Record_Special()
        MyBase.Show_Record_Special()

        If EntryMode = "New" Then
            MyBase.Absx1.txtFor("CARRIER_CODE").Text = "TRUCK"
            MyBase.Absx1.txtFor("CARRIER_MODE").Text = "M"
            MyBase.Absx1.optFor("SHIP_VIA_STATUS").Text = "A"
        End If

    End Sub

    'Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
    '    Select Case COLUMN_NAME
    '        Case "CARRIER_PROD_CODE"
    '            Dim rowSOTCARR2 As DataRow = LookUp("SOTCARR2", New String() {MyBase.Absx1.txtFor("CARRIER_CODE").Text, MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text})
    '            If rowSOTCARR2 IsNot Nothing Then

    '            End If
    '    End Select
    'End Sub

End Class