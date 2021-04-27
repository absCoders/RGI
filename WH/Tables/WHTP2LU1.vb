Public Class WHTP2LU1

    ' ALTER TABLE TATSTATE ADD STATE_REL_PCT NUMBER (3)


#Region "Overrides"

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
    End Sub
    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

    End Sub


    Overrides Sub Proceed_Update_Special_Pre()
    End Sub

    Overrides Sub Show_Record_Special()
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"WHTP2LU1"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

#End Region

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub UltraTextEditor1_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor1.ValueChanged

    End Sub

    Private Sub btnP2L_Click(sender As Object, e As EventArgs) Handles btnP2L.Click
        Dim Label As String = ""
        Dim LOCATION_CODE As String = ""
        Dim DrawArrow As String = ""

        Label = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", "WHTP2LU1")) & ""

        If String.IsNullOrEmpty(Absx1.txtFor("P2L_USER_ID").Text) Then
            MsgBox("Enter Pick to Light User ID", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If String.IsNullOrEmpty(Absx1.txtFor("USER_NAME").Text) Then
            MsgBox("Enter an User Name", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        ASCMAIN1.sql = "Select * from WHTP2LU1" _
         & " Where WH_USER_ID = '" & Absx1.txtFor("WH_USER_ID").Text & "'"

        For Each rowWK As DataRow In ASCDATA1.GetDataTable.Rows
            '   LOCATION_CODE = rowWK.Item("LOCATION_CODE")
            ShippingLabel.SendToLabelPrinter(String.Format(Label, Absx1.txtFor("USER_NAME").Text, Absx1.txtFor("P2L_USER_ID").Text, Absx1.txtFor("P2L_USER_ID").Text), cbxLabelPrinter.Text)
        Next

    End Sub

    Private Sub SetUpPortsAndPrinters()
        Dim tooltip As New System.Windows.Forms.ToolTip()

        ' Label Printer Port
        Try
            If ASCMAIN1.CLIENT = "RGI" Or ASCMAIN1.CLIENT = "VAN" Then
                txtLabelPrinter.Visible = False

                If ASCMAIN1.CLIENT = "VAN" Then
                    Dim ZebraPrinters As New List(Of String)
                    For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                        If printerName.ToUpper.StartsWith("ZDESIGNER") Or printerName.ToUpper.StartsWith("MONARCH") Or printerName.ToUpper.StartsWith("AVERY") Or printerName.ToUpper.StartsWith("ZEBRA") Then
                            ZebraPrinters.Add(printerName)
                        End If
                    Next printerName
                    If ZebraPrinters.Count >= 1 Then
                        cbxLabelPrinter.DataSource = ZebraPrinters
                    End If
                Else
                    btnP2L.Visible = False
                    Dim rows() As DataRow = ASCDATA1.GetDataTable("SELECT *  FROM WHTLPRT1").Select("")
                    For Each row As DataRow In rows
                        cbxLabelPrinter.Items.Add(row.Item("LABEL_PRINTER_ID"))
                    Next
                    cbxLabelPrinter.SelectedIndex = 0
                End If

            Else
                cbxLabelPrinter.Visible = False
                btnP2L.Visible = False

                txtLabelPrinter.BackColor = Drawing.Color.Red

                If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                    txtLabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                    tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
                    txtLabelPrinter.BackColor = Drawing.Color.Yellow
                    If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                        ASCMAIN1.LabelPrinterSerialPort.Open()
                    End If

                    If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                        txtLabelPrinter.BackColor = Drawing.Color.Green
                    End If
                ElseIf ASCMAIN1.LabelPrinterName.Length > 0 Then
                    txtLabelPrinter.Text = ASCMAIN1.LabelPrinterName
                    tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
                    txtLabelPrinter.BackColor = Drawing.Color.Green
                Else
                    Me.txtLabelPrinter.Text = "No Port / Printer"
                    tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
                End If
            End If
        Catch ex As Exception
            If ASCMAIN1.CLIENT = "RGI" Then
                cbxLabelPrinter.BackColor = Drawing.Color.Red
            Else
                txtLabelPrinter.BackColor = Drawing.Color.Red
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            End If
        End Try

    End Sub

    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    Private Sub txtLabelPrinter_ValueChanged(sender As Object, e As EventArgs) Handles txtLabelPrinter.ValueChanged

    End Sub
End Class