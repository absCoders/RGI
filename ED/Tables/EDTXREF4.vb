Public Class EDTXREF4

    Private rowEDTTRPM1 As DataRow = Nothing

    Public Overrides Sub txt_EditorButtonClick(sender As Object, e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs)
        MyBase.txt_EditorButtonClick(sender, e)
    End Sub

    Public Overrides Sub txt_EditorButtonClick_Special(txtctl As Infragistics.Win.UltraWinEditors.UltraTextEditor)
        MyBase.txt_EditorButtonClick_Special(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)

            Case "SENDER_ID_QUAL"
                If ASCMAIN1.CodeSelector.Selections = 1 Then
                    MyBase.Absx1.txtFor("SENDER_ID").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("SENDER_ID") & String.Empty
                    MyBase.Absx1.txtFor("EDI_SUPPLIER_NO").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("EDI_SUPPLIER_NO") & String.Empty
                End If

        End Select

    End Sub

    Public Overrides Sub Proceed_PreReq_Special(eItemKey As String)
        MyBase.Proceed_PreReq_Special(eItemKey)

        Select Case eItemKey

            Case "New"
                Dim EDI_TP_QUAL As String = Absx1.txtFor("SENDER_ID_QUAL").Text.Trim
                Dim EDI_TP_ID As String = Absx1.txtFor("SENDER_ID").Text.Trim

                rowEDTTRPM1 = ASCDATA1.GetDataRow("SELECT * FROM EDTTRPM1 WHERE EDI_TP_QUAL = :PARM1 and EDI_TP_ID = :PARM2", "VV", New Object() {EDI_TP_QUAL, EDI_TP_ID})
                If rowEDTTRPM1 Is Nothing Then
                    EMsg &= "Qualifer and ID combination cannot be found in EDTTRPM1."
                End If

            Case "Update"
                Validate_Code("WHSE_CODE")

                If rowEDTTRPM1 Is Nothing Then
                    Dim EDI_TP_QUAL As String = Absx1.txtFor("SENDER_ID_QUAL").Text.Trim
                    Dim EDI_TP_ID As String = Absx1.txtFor("SENDER_ID").Text.Trim
                    rowEDTTRPM1 = ASCDATA1.GetDataRow("SELECT * FROM EDTTRPM1 WHERE EDI_TP_QUAL = :PARM1 and EDI_TP_ID = :PARM2", "VV", New Object() {EDI_TP_QUAL, EDI_TP_ID})
                End If

                If rowEDTTRPM1 Is Nothing Then
                    EMsg &= vbCr & "Qualifer and ID combination cannot be found in EDTTRPM1."
                End If

        End Select
    End Sub

    Public Overrides Sub Show_Record_Special()
        Dim EDI_TP_QUAL As String = Absx1.txtFor("SENDER_ID_QUAL").Text.Trim
        Dim EDI_TP_ID As String = Absx1.txtFor("SENDER_ID").Text.Trim

        rowEDTTRPM1 = ASCDATA1.GetDataRow("SELECT * FROM EDTTRPM1 WHERE EDI_TP_QUAL = :PARM1 and EDI_TP_ID = :PARM2", "VV", New Object() {EDI_TP_QUAL, EDI_TP_ID})
        If rowEDTTRPM1 Is Nothing Then
            Absx1.txtFor("CUST_CODE").Clear()
        Else
            Absx1.txtFor("CUST_CODE").Text = rowEDTTRPM1.Item("CUST_CODE") & String.Empty
        End If

    End Sub

End Class