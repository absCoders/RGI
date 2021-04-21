Public Class WHTLOCM1
    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    Private Sub WHTLOCM1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        grpLocationFormat.Visible = (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN")
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                Dim LOCATION_CODE As String = Absx1.txtFor("LOCATION_CODE").Text
                LOCATION_CODE = LOCATION_CODE.ToUpper

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    If Len(LOCATION_CODE) = 6 And InStr(LOCATION_CODE, "-") = 0 Then
                        LOCATION_CODE = Mid(LOCATION_CODE, 1, 2) & "-" & Mid(LOCATION_CODE, 3, 3) & "-" & Mid(LOCATION_CODE, 6, 1)
                    End If

                    Absx1.txtFor("LOCATION_CODE").Text = LOCATION_CODE

                    If Mid(LOCATION_CODE, 1, 2) <> "00" And (Len(LOCATION_CODE) <> 8 Or Mid(LOCATION_CODE, 3, 1) <> "-" Or Mid(LOCATION_CODE, 7, 1) <> "-") Then
                        EMsg &= vbCr & "Invalid Format for Location Code"
                    Else
                        Dim LOC1 As String = Mid(LOCATION_CODE, 1, 2)
                        Dim LOC2 As String = Mid(LOCATION_CODE, 4, 3)
                        Dim LOC3 As String = Mid(LOCATION_CODE, 8, 1)

                        If Format(Val(LOC1), "00") <> LOC1 Or Val(LOC1) < 0 Or Val(LOC1) > 99 Then
                            EMsg &= vbCr & "Invalid Character or Format for 1st segment of Location Code"
                        End If

                        For i As Integer = 1 To LOC2.Length
                            Dim X As String = Mid(LOC2, i, 1)
                            If (X >= "A" And X <= "Z") Or (X >= "0" And X <= "9") Then
                            Else
                                EMsg &= vbCr & "Invalid Character or Format for 2nd segment of Location Code"
                            End If
                        Next


                        If InStr("ABCDEF", LOC3) = 0 And LOC1 <> "00" Then
                            EMsg &= vbCr & "Invalid Character or Format for 3rd segment of Location Code"
                        End If

                    End If
                End If
                

            Case "Edit"
            Case "Update"
                'If Absx1.txtFor("LP_CODE").Text <> "" And Absx1.chkFor("WHSE_LOCATOR").Checked Then
                '    EMsg &= vbCr & "A 3PL Warehouse Cannot be Set Up with Locator Support"
                'End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'If rowASFBASE1.Item("WHSE_LOCATOR") & "" = "1" Then
        '    Dim WHSE_CODE As String = rowASFBASE1.Item("WHSE_CODE")
        '    ASCMAIN1.sql = "Insert into WHTLOCM1 (WHSE_CODE, LOCATION_CODE, LOCATION_DESC)" & vbCrLf _
        '        & "Select '" & WHSE_CODE & "' WHSE_CODE, LOCATION_CODE, LOCATION_DESC from WHTLOCM0" & vbCrLf _
        '        & " where LOCATION_CODE in " & vbCrLf _
        '        & "(Select LOCATION_CODE from WHTLOCM0 minus " & vbCrLf _
        '        & " Select LOCATION_CODE from WHTLOCM1 where WHSE_CODE = '" & WHSE_CODE & "')"
        '    ASCDATA1.ExecuteSQL()
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        'Dim WHSE_LOCATOR As String = rowASFBASE1.Item("WHSE_LOCATOR") & ""
        ' grpLOCATIONs.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked) Or (EntryMode = "New")   

        Set_Read_Only(grpSpecial, (EntryMode <> "New") Or Not ASCMAIN1.USER_SECURITY_CODEs.Contains("SY"))
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        '  Set_Read_Only_for_ctl(Absx1.chkFor("WHSE_LOCATOR"), Not (EntryMode = "New"))
        grpPrintLabels.Visible = Not ScreenMode
        Set_Read_Only_for_ctl(txtLOCATION_FROM, ScreenMode)
        Set_Read_Only_for_ctl(txtLOCATION_TO, ScreenMode)
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click

        Dim LocationLabel As String = "^XA^FO200,75^BY8^BCR,500,N,N,N^FD{0}^FS^CF0,190^FWR^FO10,75^FD{1}^FS^XZ"

        If Not ScreenMode Then
            If String.IsNullOrEmpty(Absx1.txtFor("WHSE_CODE").Text) Then
                MsgBox("Enter a Whse code", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtLOCATION_FROM.Text) Or String.IsNullOrEmpty(txtLOCATION_TO.Text) Then
                MsgBox("Enter both From and To Locations to print", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            ASCMAIN1.sql = "Select * from WHTLOCM1" _
         & " Where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text _
         & "' and LOCATION_CODE between '" & txtLOCATION_FROM.Text & "' and '" & txtLOCATION_TO.Text & "'"
            For Each rowWK As DataRow In ASCDATA1.GetDataTable.Rows
                If ASCMAIN1.CLIENT = "VAN" Then
                    ShippingLabel.SendToLabelPrinter(String.Format(LocationLabel, rowWK.Item("LOCATION_CODE"), rowWK.Item("LOCATION_CODE")), cbxLabelPrinter.Text)
                ElseIf ASCMAIN1.CLIENT = "RGI" Then
                    PrintService_Label(rowWK.Item("LOCATION_CODE"))
                Else
                    ShippingLabel.SendToLabelPrinter(String.Format(LocationLabel, rowWK.Item("LOCATION_CODE"), rowWK.Item("LOCATION_CODE")))
                End If
            Next
        Else
            If ASCMAIN1.CLIENT = "VAN" Then
                ShippingLabel.SendToLabelPrinter(String.Format(LocationLabel, Absx1.txtFor("LOCATION_CODE").Text, Absx1.txtFor("LOCATION_CODE").Text), cbxLabelPrinter.Text)
            ElseIf ASCMAIN1.CLIENT = "RGI" Then
                PrintService_Label(Absx1.txtFor("LOCATION_CODE").Text)
            Else
                ShippingLabel.SendToLabelPrinter(String.Format(LocationLabel, Absx1.txtFor("LOCATION_CODE").Text, Absx1.txtFor("LOCATION_CODE").Text))
            End If

        End If

    End Sub

    Private Sub PrintService_Label(Loc As String)
        Dim Printer = cbxLabelPrinter.Text

        Dim Label = "NEWER|BARCODE_1TXT.lbx|" & Printer & "|" & Loc & "|"

        Using ipp As New nsoftware.IPWorks.Ipport
            ipp.RuntimeLicense = "31504E3941413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600004B4857525953375A4A5A375A0000"
            If ASCMAIN1.Running_in_VS Then
                ipp.Connect("192.168.120.52", "4444") 'ipp.Connect("192.168.120.67", "4444") '"192.168.4.117", "4444")
            Else
                ipp.Connect("192.168.110.223", "4444")
            End If

            ipp.SendLine(Label)
            ipp.Disconnect()
        End Using


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

    Private Sub btnP2L_Click(sender As Object, e As EventArgs) Handles btnP2L.Click
        Dim LocationLabel As String = ""
        Dim LOCATION_CODE As String = ""
        LocationLabel = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", "P2L_BAYS")) & ""

        If String.IsNullOrEmpty(Absx1.txtFor("WHSE_CODE").Text) Then
                MsgBox("Enter a Whse code", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            If String.IsNullOrEmpty(txtLOCATION_FROM.Text) Or String.IsNullOrEmpty(txtLOCATION_TO.Text) Then
                MsgBox("Enter both From and To Locations to print", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
        ASCMAIN1.sql = "Select * from WHTLOCM1" _
         & " Where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text _
         & "' and LOCATION_CODE like '" & txtLOCATION_FROM.Text & "-__-A-1'"
        For Each rowWK As DataRow In ASCDATA1.GetDataTable.Rows
            LOCATION_CODE = rowWK.Item("LOCATION_CODE")
            ShippingLabel.SendToLabelPrinter(String.Format(LocationLabel, LOCATION_CODE.Substring(0, 5), rowWK.Item("LOCATION_CODE")), cbxLabelPrinter.Text)
        Next

    End Sub

#End Region


End Class