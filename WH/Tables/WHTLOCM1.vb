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
                If Absx1.txtFor("LOCATION_USE").Text = "L" Then
                    EMsg = "P2L Location, Edit not allowed"
                End If

            Case "Update"
                If Absx1.txtFor("LOCATION_USE").Text = "L" And EntryMode = "E" Then
                    EMsg = "P2L Location, Changes not allowed"
                End If
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
        Set_Read_Only_for_ctl(optArrow, ScreenMode)
        'not drawing an arrow anymore
        optArrow.Visible = False

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
        Dim DrawArrow As String = ""
        Dim LeftArrow As String = "^FO320,690^GFA,4608,4608,24,,::::::::::::gH01FF8,g03LFC" _
                                & ",Y03NFE,X07PFE,W03RFC,V01TF8,V07TFE,U03VFC,U0XF,T03XFC,T07YF" _
                                & ",S01gF8,S07gFE,R01gHF8,R03gHFC,R0gJF,Q01gJF8,Q03gJFC,Q07gKF" _
                                & ",P01gLF8,P03gLFC,P07gLFE,P0gNF,O01gNF8,O03gNFC,O07gNFE,O0gPF" _
                                & ",N01TFE7TF8,N03TF00TFC,N07SFE007SFE,N07SFC003TF,N0TF8001TF" _
                                & ",M01TFJ0TF8,M03SFEJ07SFC,M07SFCJ03SFE,M07SF8J01TF,M0TFL0TF,L01SFEL07SF8" _
                                & ",L01SFCL03SFC,L03SF8L01SFC,L03SFN0SFE,L07RFEN0SFE,L0SFCN0TF,L0SF8N07SF8" _
                                & ",K01SFO07SF8,K01RFEO0TFC,K03RFCO0TFC,K03RF8O0TFE,K07RFO01TFE,K07QFEO03UF" _
                                & ",K0RFCO07UF,K0RF8O0VF,K0RFO01VF8,J01QFEO03VF8,J01QFCO07VFC,J03QF8O0WFC" _
                                & ",J03QFO01WFC,J03PFEO03WFE,J07PFCO07WFE,J07PF8O0XFE,J07PFO01XFE,J07OFEO03YF" _
                                & ",J0PFCO07YF,J0PF8O0gF,J0PFO01gF,J0OFEO03gF8,I01OFCO07gF8,I01OF8O0gGF8,I01OFO01gGF8" _
                                & ",I01NFEgI01MF8,I01NFCgJ0MFC,I03NF8gJ07LFC,I03NFgK03LFC,I03MFEgK03LFC,I03MFCgK01LFC" _
                                & ",I03MF8gK01LFC,I03MFgL01LFC,I03LFEgL01LFC,I03LFCgL01LFC,:I03LF8gL01LFE,I03LF8gL01LFC" _
                                & ",:::I03LF8gL01LFE,I03LFCgL01LFC,I03LFEgL01LFC,:I03MF8gK01LFC,:I03MFCgK01LFC,I03NFgK03LFC" _
                                & ",:I01NF8gJ07LFC,I01NFEgJ0MFC,I01NFEO01RFENF8,I01OFO01gGF8,I01OFCO0gGF8,I01OFCO07gF8" _
                                & ",J0OFEO03gF8,J0PF8N01gF,J0PF8O0gF,J0PFCO07YF,J07PFO03YF,J07PFO01XFE,J07PF8O0XFE,J07PFEO07WFE" _
                                & ",J03PFEO03WFC,J03QFO01WFC,J01QFCO0WFC,J01QFCO07VF8,J01QFEO03VF8,K0RF8N01VF,K0RF8O0VF,K0RFCO07UF" _
                                & ",K07RFO03TFE,K07RFO01TFE,K03RF8O0TFC,K03RFEO0TFC,K01RFEO0TF8,K01SFO07SF8,L0SFCN0TF,:L07RFEN0SFE" _
                                & ",L03SF8L01SFC,:L01SFCL03SF8,M0TFL07SF8,M0TFL0TF,M07SF8J01SFE,M03SFEJ03SFE,M03SFEJ07SFC,M01TFJ0TF8" _
                                & ",N0TFC001TF,N07SFC003SFE,N03TF00TFE,N01TF81TFC,O0gPF8,O0gPF,O07gNFE,O03gNFC,O01gNF8,P07gMF,P03gLFE,P01gLF8" _
                                & ",Q0gLF,Q07gJFE,Q03gJFC,R0gJF,R07gHFE,R01gHFC,S0gHF,S03gFC,T0gF8,T07XFE,U0XF8,U07VFE,U01VF8,V03TFC,W07RFE" _
                                & ",X0RF8,X01PF8,g0NF8,gG03JFC,,:::::::::::::^FS"
        Dim RightArrow As String = "^FO320,690^GFA,4608,4608,24,,:::::::::::::gG03JFC,Y01NF,X01PF8,W01RF,W07RFE,V03TFC,U01VF8,U07VFE,T01XF" _
                                & ",T07XFE,S01gF,S03gFC,S0gHF,R03gHF8,R07gHFE,R0gJF,Q03gJFC,Q07gJFE,Q0gLF,P01gLF8,P07gLFC,P0gMFE,O01gNF8,O03gNFC" _
                                & ",O07gNFE,O0gPF,N01gPF,N03TF81TF8,N07TF00TFC,N07SFC003SFE,N0TF8003TF,M01TFJ0TF8,M03SFEJ07SFC,M07SFCJ07SFC,M07SF8J01SFE" _
                                & ",M0TFL0TF,L01SFEL0TF,L01SFCL03SF8,L03SF8L01SFC,:L07SFN07RFE,L0TFN03SF,:K01SFEO0SF8,K01TFO07RF8,K03TFO07RFC,K03TFO01RFC" _
                                & ",K07TF8O0RFE,K07TFCO0RFE,K0UFEO03RF,K0VFO01RF,K0VF8N01RF,J01VFCO07QF8,J01VFEO03QF8,J03WFO03QF8,J03WF8O0QFC,J03WFCO07PFC" _
                                & ",J07WFEO07PFE,J07XFO01PFE,J07XF8O0PFE,J0YFCO0PFE,J0YFEO03PF,J0gFO01PF,J0gF8N01PF,I01gFCO07OF,I01gFEO03OF8,I01gGFO03OF8" _
                                & ",I01gGF8O0OF8,I01NF7RF8O07NF8,I03MFgJ07NF8,I03LFEgJ01NF8,I03LFCgK0NFC,:I03LF8gK03MFC,I03LF8gK01MFC,:I03LF8gL07LFC" _
                                & ",:I03LF8gL03LFC,I07LF8gL01LFC,I03LF8gL01LFC,:::I07LF8gL01LFC,I03LF8gL03LFC,:I03LF8gL07LFC,I03LF8gL0MFC,I03LF8gK01MFC" _
                                & ",I03LF8gK03MFC,I03LFCgK07MFC,I03LFCgK0NFC,I03LFEgJ01NFC,I03MFgJ03NF8,I01MF8gI07NF8,I01gGF8O0OF8,I01gGFO01OF8,I01gFEO03OF8" _
                                & ",I01gFCO07OF,J0gF8O0PF,J0gFO01PF,J0YFEO03PF,J0YFCO07OFE,J07XF8O0PFE,J07XFO01PFE,J07WFEO03PFE,J07WFCO07PFC,J03WF8O0QFC" _
                                & ",J03WFO01QFC,J03VFEO03QF8,J01VFCO07QF8,J01VF8O0RF,K0VFO01RF,K0UFEO03RF,K0UFCO07QFE,K07TF8O0RFE,K07TFO01RFC,K03TFO03RFC" _
                                & ",K03TFO07RF8,K01SFEO0SF8,K01SFEN01SF,L0TFN03SF,L07SFN07RFE,L07SFN0SFC,L03SF8L01SFC,L03SFCL03SF8,L01SFEL07SF8,M0TFL0TF" _
                                & ",M0TF8J01SFE,M07SFCJ03SFE,M03SFEJ07SFC,M01TFJ0TF8,N0TF8001TF,N0TFC003SFE,N07SFE007SFE,N03TF00TFC,N01TFE7TF8,O0gPF,O07gNFE" _
                                & ",O03gNFC,O01gNF8,P0gNF,P07gLFE,P03gLFC,P01gLF8,Q0gKFE,Q03gJFC,Q01gJF8,R0gJF,R03gHFC,R01gHF8,S07gFE,S01gF8,T0YFE,T03XFC" _
                                & ",U0XF,U03VFC,V07TFE,V01TF8,W03RFC,X07PFE,Y07NFC,g03LFC,gH01FF8,,::::::::::::^FS
"

        LocationLabel = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", "P2L_BAYS")) & ""

        If String.IsNullOrEmpty(Absx1.txtFor("WHSE_CODE").Text) Then
            MsgBox("Enter a Whse code", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If String.IsNullOrEmpty(txtLOCATION_FROM.Text) Or String.IsNullOrEmpty(txtLOCATION_TO.Text) Then
            MsgBox("Enter both From and To Locations to print", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        'If optArrow.CheckedItem Is Nothing Then
        '    MsgBox("Select arrow direction for P2L labels")
        '    Exit Sub
        'End If
        'If optArrow.CheckedItem.DisplayText = "Left Arrow" Then
        '    DrawArrow = LeftArrow
        'Else
        '    DrawArrow = RightArrow
        'End If

        ASCMAIN1.sql = "Select * from WHTLOCM1" _
         & " Where WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text _
         & "' and LOCATION_CODE like '" & txtLOCATION_FROM.Text & "-__-A-1'"
        For Each rowWK As DataRow In ASCDATA1.GetDataTable.Rows
            LOCATION_CODE = rowWK.Item("LOCATION_CODE")
            ShippingLabel.SendToLabelPrinter(String.Format(LocationLabel, rowWK.Item("LOCATION_ZONE"), rowWK.Item("LOCATION_CODE"), LOCATION_CODE.Substring(0, 5)), cbxLabelPrinter.Text)
        Next

    End Sub

#End Region


End Class