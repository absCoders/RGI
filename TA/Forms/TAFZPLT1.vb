Imports ABSolution

Public Class TAFZPLT1

    Private labelImage As String = String.Empty
    Private clsTACZPLT1 As New TAC.TACZPLT1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            labelImage = "^XA
                                ^CF0,60
                                ^FO40,50^FDTEST LABEL^FS
                                ^FO50,150^GB700,3,3^FS
                                ^CFA,30
                                ^FO40,215^FDShip From^FS
                                ^FO60,255^FD{Sender.Company}^FS
                                ^FO60,295^FD{Sender.Address1}^FS
                                ^FO60,335^FD{Sender.Address2}^FS
                                ^FO60,375^FD{Sender.Address3}^FS
                                ^FO60,415^FD{Sender.City, State, ZipCode}^FS
                                ^FO40,500^FDShip To^FS
                                ^FO60,540^FD{Recipient.Company}^FS
                                ^FO60,580^FD{Recipient.Address1}^FS
                                ^FO60,620^FD{Recipient.Address2}^FS
                                ^FO60,660^FD{Recipient.Address3}^FS
                                ^FO60,700^FD{Recipient.City, State, ZipCode}^FS
                                ^BY3,2,150
                                ^FO100,825^BC^FD1234567890^FS
                                ^CF0,60
                                ^FO60,1100^FDTEST - Do Not Ship^FS
                                ^XZ"

            lblDirections.Text = "Instructions"
            lblDirections.Text &= Environment.NewLine & Environment.NewLine
            lblDirections.Text &= "Turn OFF the Label printer and wait 10 seconds."
            lblDirections.Text &= Environment.NewLine & Environment.NewLine
            lblDirections.Text &= "Turn ON the label printer."
            lblDirections.Text &= Environment.NewLine & Environment.NewLine
            lblDirections.Text &= "When the Label printer's light is green, click the 'Print ZPL Label' button."
            lblDirections.Text &= Environment.NewLine & Environment.NewLine
            lblDirections.Text &= "Let ABS know if the label does not print or if you get an Error Message."
            lblDirections.Text &= Environment.NewLine & Environment.NewLine
            lblDirections.Text &= "Turn OFF the Label printer, wait 10 seconds then Turn ON the label printer."

            txtLabelPrinter.Appearance.BackColor = Drawing.Color.LightGreen
            If ASCMAIN1.LabelPrinterIPAddress.Length > 0 Then
                txtLabelPrinter.Text = ASCMAIN1.LabelPrinterIPAddress
            Else
                txtLabelPrinter.Text = "No Port"
                txtLabelPrinter.Appearance.BackColor = Drawing.Color.Red
                txtLabelPrinter.Appearance.ForeColor = Drawing.Color.White
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Label Printer", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = String.Empty

        Dim sql As String = String.Empty
        Dim zMsg As String = String.Empty

        Select Case eItemKey

            Case "Load"

            Case "Generate"

            Case "Cancel"

            Case "Update"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"

            Case "Generate"

            Case "Cancel"

            Case "Update"

        End Select


    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

    End Sub

    Private Sub Clear_Record()

    End Sub

    Private Sub Load_Record()

    End Sub

    Private Sub Update_Record()

    End Sub

    Private Sub btnLabelPrinter_Click(sender As Object, e As EventArgs) Handles btnLabelPrinter.Click

        Try
            clsTACZPLT1.SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Test ZPL", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#End Region

End Class