Public Class ICTCOLR1


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        txtCOLOR_CODE_LONG.Visible = (ASCMAIN1.CLIENT = "RGI")
        lblCOLOR_CODE_LONG.Visible = (ASCMAIN1.CLIENT = "RGI")
        lblCOLOR_CODE_LONG2.Visible = (ASCMAIN1.CLIENT = "RGI")

    End Sub

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
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim S As New System.Text.StringBuilder With {.Length = 0}
                    S.AppendLine("SELECT COUNT(*) FROM ICTCOLR1")
                    S.AppendLine("WHERE NVL(COLOR_CODE_LONG,'NULL') = '" & txtCOLOR_CODE_LONG.Text & "'")
                    S.AppendLine(String.Format("AND COLOR_CODE <> '{0}'", Absx1.txtFor("COLOR_CODE").Text))
                    S.AppendLine("AND NVL(COLOR_STATUS,'A') = 'A'")
                    ASCMAIN1.sql = S.ToString()
                    Dim RecCntL As Int16 = Val(ASCDATA1.GetDataValue)
                    If RecCntL > 0 Then
                        EMsg &= vbCr & "Long Code " & txtCOLOR_CODE_LONG.Text & " Is Already Used On Another Active Color"
                    End If

                    S.Length = 0
                    S.AppendLine("SELECT COUNT(*) FROM ICTCOLR1")
                    S.AppendLine("WHERE NVL(COLOR_DESC,'NULL') = '" & txtCOLOR_DESC.Text & "'")
                    S.AppendLine(String.Format("AND COLOR_CODE <> '{0}'", Absx1.txtFor("COLOR_CODE").Text))
                    S.AppendLine("AND NVL(COLOR_STATUS,'A') = 'A'")
                    ASCMAIN1.sql = S.ToString()
                    Dim RecCntD As Int16 = Val(ASCDATA1.GetDataValue)
                    If RecCntD > 0 Then
                        EMsg &= vbCr & "Color Description " & txtCOLOR_DESC.Text & " Is Already Used On Another Active Color"
                    End If

                    If txtCOLOR_CODE_LONG.Text & String.Empty = "" Then
                        EMsg &= vbCr & "Long Code Can Not Be Blank "
                    End If

                    If txtCOLOR_DESC.Text & String.Empty = "" Then
                        EMsg &= vbCr & "Description Can Not Be Blank "
                    End If

                    S.Length = 0
                    Dim COLOR_STATUS As String = Absx1.optFor("COLOR_STATUS").Value
                    If COLOR_STATUS <> "A" And COLOR_STATUS <> "I" Then
                        EMsg &= vbCr & "Invalid Color Status"
                    End If
                End If
        End Select
    End Sub

#End Region

End Class