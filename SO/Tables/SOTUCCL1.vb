Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing.Printing

Public Class SOTUCCL1
    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        '  SetUpPortsAndPrinters()
    End Sub

    Private Sub btnTest_Click(sender As System.Object, e As System.EventArgs) Handles btnTest.Click

        Dim UCC128 As String = Absx1.txtFor("UCC128_COMMANDS").Text

        If UCC128 = "" Then
            MsgBox("Nothing to Test", MsgBoxStyle.OkOnly, "No Template Defined")
            Exit Sub
        End If

        Dim PrinterName As String = ""

        If ASCMAIN1.CLIENT = "VAN" Then
            Dim ZebraPrinter As String = cboZebraPrinter.SelectedValue
            Dim PRINTER_PORT As String = ZebraPrinter.Split("|")(2)
            PrinterName = PRINTER_PORT
        ElseIf ASCMAIN1.CLIENT = "RGI" Then
            Dim ZebraPrinter As String = cboZebraPrinter.SelectedValue
            'Dim PRINTER_PORT As String = ZebraPrinter.Split("|")(2)
            PrinterName = ZebraPrinter
        End If

        If txtCartonNo.Text <> "" Then
            Dim cartonLabel As New TestLabel(Absx1.txtFor("LABEL_TEMPLATE_CODE").Text, txtCartonNo.Text)
            Try
                cartonLabel.PrintLabel(1, PrinterName)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            ShippingLabel.SendToLabelPrinter(UCC128, PrinterName)
        End If


    End Sub

    Private Sub SetUpPortsAndPrinters()
        Dim tooltip As New System.Windows.Forms.ToolTip()

        ' Label Printer Port
        Try
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

        Catch ex As Exception
            txtLabelPrinter.BackColor = Drawing.Color.Red
            tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
        End Try

    End Sub

    Private Sub SOTUCCL1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim ZebraPrinters As New List(Of String)
        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Select * from ASTPRNT1"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim PRINTER_CODE As String = row.Item("PRINTER_CODE")
                Dim PRINTER_NAME As String = row.Item("PRINTER_NAME")
                Dim PRINTER_PORT As String = row.Item("PRINTER_PORT")

                Dim ZebraPrinter As String = PRINTER_CODE & "|" & PRINTER_NAME & "|" & PRINTER_PORT
                ZebraPrinters.Add(ZebraPrinter)
            Next
            cboZebraPrinter.DataSource = ZebraPrinters
        ElseIf ASCMAIN1.CLIENT = "RGI" Then
            Try
                ASCMAIN1.sql = "Select * From ICTWHSEL"
                For Each rowICTWHSEL As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "LABEL_IP_ADDRESS")
                    Dim LABEL_IP_ADDRESS As String = rowICTWHSEL.Item("LABEL_IP_ADDRESS") & String.Empty
                    If LABEL_IP_ADDRESS.Length = 0 Then Continue For
                    If ZebraPrinters.Contains(LABEL_IP_ADDRESS) Then Continue For

                    ZebraPrinters.Add(rowICTWHSEL.Item("LABEL_IP_ADDRESS") & String.Empty)
                Next
            Catch ex As Exception

            End Try

            Try
                If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                    ZebraPrinters.Add(ASCMAIN1.LabelPrinterSerialPort.PortName)
                End If
            Catch ex As Exception

            End Try

            cboZebraPrinter.DataSource = ZebraPrinters
        Else
            For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                If printerName.ToUpper.StartsWith("ZDESIGNER") Or printerName.ToUpper.StartsWith("MONARCH") Or printerName.ToUpper.StartsWith("AVERY") Or printerName.ToUpper.StartsWith("ZEBRA") Then
                    ZebraPrinters.Add(printerName)
                End If
            Next printerName
            If ZebraPrinters.Count >= 1 Then
                cboZebraPrinter.DataSource = ZebraPrinters
            End If
        End If
    End Sub


#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

    End Sub

    Overrides Sub Clear_Record_Special()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        grpTest.Visible = ScreenMode
    End Sub

#End Region
     
End Class