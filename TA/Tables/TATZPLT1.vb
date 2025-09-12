Public Class TATZPLT1

    Private clsTACZPLT1 As New TAC.TACZPLT1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst

        End With

        optLabelSize.CheckedIndex = 0
    End Sub

#Region "Overrides"
    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "Edit"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Show_Record_Special()

        If Not ScreenMode Then

        Else

        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then

        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then

        End If

        lblIP.Visible = ScreenMode And (EntryMode = "View")
        txtIP.Visible = ScreenMode And (EntryMode = "View")
        btnPrint.Visible = ScreenMode And (EntryMode = "View")

    End Sub

#End Region

    Private Sub btnPrint_Click(sender As Object, e As EventArgs) Handles btnPrint.Click

        Dim zpl As String = Absx1.txtFor("ZPL_BODY").Text
        clsTACZPLT1.SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, zpl)
    End Sub

    Private Sub btnViewLabel_Click(sender As Object, e As EventArgs) Handles btnViewLabel.Click

        Try
            Absx1.txtFor("ZPL_BODY").Text = Absx1.txtFor("ZPL_BODY").Text.Trim
            If Absx1.txtFor("ZPL_BODY").TextLength = 0 Then
                MessageBox.Show("There is no label defined.", "Display Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Select Case optLabelSize.Value
                Case 4
                    clsTACZPLT1.ShowLabelDialog(TACZPLT1.LabelSizes.label4x6, Absx1.txtFor("ZPL_BODY").Text)
                Case 2
                    clsTACZPLT1.ShowLabelDialog(TACZPLT1.LabelSizes.label225x125, Absx1.txtFor("ZPL_BODY").Text)
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message, "View Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnPrintLabel_Click(sender As Object, e As EventArgs) Handles btnPrintLabel.Click

        Try
            Absx1.txtFor("ZPL_BODY").Text = Absx1.txtFor("ZPL_BODY").Text.Trim
            If Absx1.txtFor("ZPL_BODY").TextLength = 0 Then
                MessageBox.Show("There is no label defined.", "Display Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Select Case optLabelSize.CheckedIndex
                Case 0
                    If ASCMAIN1.LabelPrinterIPAddress.Length = 0 Then
                        MessageBox.Show("You are not assigned a 4x6 label printer", "Print Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    clsTACZPLT1.SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, Absx1.txtFor("ZPL_BODY").Text)
                Case 1
                    If ASCMAIN1.MiniLabelPrinterIPAddress.Length = 0 Then
                        MessageBox.Show("You are not assigned a mini label printer", "Print Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    clsTACZPLT1.SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, Absx1.txtFor("ZPL_BODY").Text)

                Case Else
                    MessageBox.Show("Select a Label Size", "Print Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try



    End Sub
End Class