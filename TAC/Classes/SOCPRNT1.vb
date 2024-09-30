Imports System.Text.RegularExpressions
Imports System
Imports System.IO

Public MustInherit Class ShippingLabel
    Private Property lastPrintedData As String

    Public Sub PrintLabel(Optional ByVal printQty As Integer = 1, Optional ByVal PrinterName As String = "", Optional ByVal labelTemplateOverride As String = "")
        Dim labelData As Dictionary(Of String, DataRow) = GetLabelData()
        Dim labelTemplate As String = ""
        If IsNothing(labelData) Then
            Exit Sub
        End If
        If labelTemplateOverride <> "" Then
            labelTemplate = labelTemplateOverride
            If ASCMAIN1.CLIENT = "VAN" Then
                labelTemplate = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", labelTemplateOverride)) & ""
            End If
        Else
            labelTemplate = GetLabelTemplate()
        End If

        If (ASCMAIN1.CLIENT = "VAN" And InStr(labelTemplate, "DSDC Multi-PO") > 0) Then
            Dim row As DataRow = labelData("SOTPICK1")
            'if SOTPICK1 is empty skip printing
            If row("PICK_NO") & "" = "" Then
                Exit Sub
            End If
        End If

        For i As Integer = 1 To printQty
            ChangeLabelData(labelData, i, printQty)
            Dim labeltoPrint As String = FillLabelTemplateWithData(labelTemplate, labelData)
            SendToLabelPrinter(labeltoPrint, PrinterName)
        Next
    End Sub

    Public Sub SaveLabelToFile(ByVal fileName As String, ByVal append As Boolean, Optional ByVal fromLastPrinted As Boolean = True)
        If fromLastPrinted AndAlso Not String.IsNullOrEmpty(lastPrintedData) Then
            If append Then
                File.AppendAllText(fileName, lastPrintedData)
            Else
                File.WriteAllText(fileName, lastPrintedData)
            End If
        Else
            Dim labelData As Dictionary(Of String, DataRow) = GetLabelData()
            Dim labelTemplate As String = GetLabelTemplate()

            ChangeLabelData(labelData, 1, 1)
            Dim labeltoPrint As String = FillLabelTemplateWithData(labelTemplate, labelData)
            If append Then
                File.AppendAllText(fileName, labeltoPrint)
            Else
                File.WriteAllText(fileName, labeltoPrint)
            End If
        End If
    End Sub

    Protected MustOverride Function GetLabelData() As Dictionary(Of String, DataRow)
    Protected MustOverride Function GetLabelTemplate() As String

    Protected Overridable Sub ChangeLabelData(ByVal labelData As Dictionary(Of String, DataRow), ByVal currentIndex As Integer, ByVal lastIndex As Integer)

    End Sub

    Protected Function FillLabelTemplateWithData(labelTemplate As String, labelData As Dictionary(Of String, DataRow)) As String
        'Matches <<TABLE.COLUMN>...>, and if the value of TABLE.COLUMN is null, it omits this line from the ZPL
        'Used for hiding a section of label if the data is unavailable
        labelTemplate = Regex.Replace(labelTemplate, "\<\<(?<table>[\w_]+)\.(?<column>[\w_]+)\>(?<command>.*)\>",
                        Function(m) If(labelData(m.Groups("table").Value).Item(m.Groups("column").Value) & "" = "",
                                       "", m.Groups("command").Value))


        'Regex matches {TABLE.COLUMN} and replaces with values from rowUCC128
        labelTemplate = Regex.Replace(labelTemplate, "\{(?<table>[\w_]+)\.(?<column>[\w_]+)\}",
                        Function(m) labelData(m.Groups("table").Value).Item(m.Groups("column").Value) & "")


        If labelTemplate.Contains("{CARTONDETAILS}") Then
            labelTemplate = WriteCartonDetails(labelTemplate, labelData("SOTCART1").Item("CART_NO") & String.Empty)
        End If

        ' Special Processing for Regency Customer 230514 WINNERS MERCHANTS INC. UCC 128 LABEL
        If ASCMAIN1.CLIENT = "RGI" Then
            If labelTemplate.Contains("{SOTCART2:") Then
                If labelData.ContainsKey("SOTCART2_1") Then
                    WriteCartonDetailsRGI(labelTemplate, labelData)
                End If
            End If
        End If

        Return labelTemplate
    End Function

    Private Sub WriteCartonDetailsRGI(ByRef labelTemplate As String, labelData As Dictionary(Of String, DataRow))
        If Not labelTemplate.Contains("{SOTCART2:") Then
            Exit Sub
        End If

        If Not labelData.ContainsKey("SOTCART2_1") Then
            Exit Sub
        End If

        ' Example entry
        '^FO{X},{Y}^A1N,22,20^FD {SOTCART2:340:10:STYLE_CODE,020:COLOR_CODE,0320:QTY_PACKED,0520}^FS
        ' 340 replaces {Y}
        ' 10 is how much to Increment {Y}
        ' STYLE_CODE,020 - place the style code at 020, 340 

        Dim splitData As String() = labelTemplate.Split(Environment.NewLine)
        Dim lineToReplace As String = String.Empty
        For Each entry As String In splitData
            If entry.Contains("{SOTCART2:") Then
                lineToReplace = entry
                Exit For
            End If
        Next

        If lineToReplace.Length = 0 Then
            Exit Sub
        End If

        Dim extractedRules As String = String.Empty
        Dim startLoc As Int32 = lineToReplace.IndexOf("{SOTCART2")
        Dim endLoc As Int32 = lineToReplace.IndexOf("}", startLoc + 1)

        If startLoc < 0 Then
            Exit Sub
        End If

        If endLoc < 0 Then
            Exit Sub
        End If

        If endLoc <= startLoc Then
            Exit Sub
        End If

        '^FO{X},{Y}^A1N,22,20^FD {SOTCART2:340:10:STYLE_CODE,020:COLOR_CODE,0320:QTY_PACKED,0520}^FS
        Dim dataToParce As String = lineToReplace.Substring(startLoc + 1, endLoc - (startLoc + 1))
        ' dataToParce should be: SOTCART2:340:10STYLE_CODE,020:COLOR_CODE,0320:QTY_PACKED,0520
        Dim contentToReplace As String = lineToReplace.Substring(startLoc, endLoc - (startLoc - 1))
        ' contentToReplace should be {SOTCART2:340:10STYLE_CODE,020:COLOR_CODE,0320:QTY_PACKED,0520}

        splitData = dataToParce.Split(":")
        ' sample data in splitdata
        ' SOTCART2
        ' 340
        ' 10
        ' STYLE_CODE,020
        ' COLOR_CODE, 0320
        ' QTY_PCKED, 0520

        ' ^FO{Y},{X}^A1N,22,20^FD {SOTCART2:340:10:STYLE_CODE,020:COLOR_CODE,0320:QTY_PACKED,0520}^FS
        Dim rowFound As Boolean = True
        Dim rowIndex As Int32 = 1
        Dim replacementCode As String = String.Empty
        Dim workString As String = lineToReplace.Replace(contentToReplace, "{@@@}")
        Dim increment As Int64 = Val(splitData(2))
        Dim yLoc As Int64 = Val(splitData(1))

        While rowFound
            If Not labelData.ContainsKey($"SOTCART2_{rowIndex}") Then
                rowFound = False
                Continue While
            End If

            For iLoop As Int16 = 3 To splitData.Length - 1
                If splitData(iLoop).Contains(",") Then
                    Dim fieldName As String = splitData(iLoop).Split(",")(0)
                    Dim xCoord As String = splitData(iLoop).Split(",")(1)

                    If labelData($"SOTCART2_{rowIndex}").Table.Columns.Contains(fieldName) Then
                        replacementCode &= workString.Replace("{Y}", yLoc).Replace("{X}", xCoord).Replace("{@@@}", labelData($"SOTCART2_{rowIndex}").Item(fieldName) & String.Empty)
                    End If
                End If

            Next

            yLoc += increment
            rowIndex += 1
        End While

        If replacementCode.Length > 0 Then
            labelTemplate = labelTemplate.Replace(lineToReplace, replacementCode)
        End If

    End Sub

    Private Function WriteCartonDetails(ByVal labelTemplate As String, ByVal CART_NO As String) As String
        Dim wklabelTemplate As String = String.Empty
        Dim pages As Int16 = 1
        Dim totalQuantity As Int16 = 0
        Dim labelIndex As Int16 = 1
        Dim labelArray(labelIndex) As String
        Dim labelDetails As String = String.Empty

        Dim startYCoord As Int16 = 280
        Dim detailLineSpacing As Int16 = 30
        Dim maxYCoord As Int16 = 1100
        Dim ycoord As Int16 = startYCoord

        Dim lblCART_LNO As String = "^FO05,{Y}^ABN,22,12^FD {CART_LNO}^FS"
        Dim lblSTYLE_CODE As String = "^FO070,{Y}^ABN,22,12^FD {STYLE_CODE}^FS"
        Dim lblCOLOR_CODE As String = "^FO0325,{Y}^ABN,22,12^FD {COLOR_CODE}^FS"
        Dim lblSIZE_DESC As String = "^FO0500,{Y}^ABN,22,12^FD {SIZE_DESC}^FS"
        Dim lblQTY_PACKED As String = "^FO0700,{Y}^ABN,22,12^FD {QTY_PACKED}^FS"

        Dim tblSOTCART2 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTCART2 WHERE CART_NO = :PARM1 AND QTY_PACKED > 0", "SOTCART2", "V", New Object() {CART_NO})

        If tblSOTCART2.Rows.Count = 0 Then
            labelTemplate = labelTemplate.Replace("{CARTONDETAILS}", "")
            labelTemplate = labelTemplate.Replace("{PAGE}", 1)
            labelTemplate = labelTemplate.Replace("{PAGEN}", 1)
            labelTemplate = labelTemplate.Replace("{TOTALQUANTITY}", 0)
            Return labelTemplate
        End If

        For Each rowSOTART2 As DataRow In tblSOTCART2.Select("", "STYLE_CODE,COLOR_CODE")

            If ycoord > maxYCoord Then
                labelArray(labelIndex) = labelTemplate.Replace("{CARTONDETAILS}", labelDetails)
                labelArray(labelIndex) = labelArray(labelIndex).Replace("{TOTALQUANTITY}", totalQuantity)
                labelDetails = String.Empty
                labelIndex += 1
                ReDim Preserve labelArray(labelIndex)
                pages += 1
                totalQuantity = 0
                ycoord = startYCoord
            End If

            labelDetails &= lblCART_LNO.Replace("{CART_LNO}", rowSOTART2.Item("CART_LNO")) & vbCrLf _
                            & lblSTYLE_CODE.Replace("{STYLE_CODE}", rowSOTART2.Item("STYLE_CODE") & "") & vbCrLf _
                            & lblCOLOR_CODE.Replace("{COLOR_CODE}", rowSOTART2.Item("COLOR_CODE") & "") & vbCrLf _
                            & lblSIZE_DESC.Replace("{SIZE_DESC}", rowSOTART2.Item("SIZE_DESC") & "") & vbCrLf _
                            & lblQTY_PACKED.Replace("{QTY_PACKED}", rowSOTART2.Item("QTY_PACKED") & "") & vbCrLf
            labelDetails = labelDetails.Replace("{Y}", ycoord)
            totalQuantity += Val(rowSOTART2.Item("QTY_PACKED") & "")
            ycoord += detailLineSpacing
        Next

        ' Update the last label 
        labelArray(labelIndex) = labelTemplate.Replace("{CARTONDETAILS}", labelDetails)
        labelArray(labelIndex) = labelArray(labelIndex).Replace("{TOTALQUANTITY}", totalQuantity)

        For pageNumber As Int16 = 1 To labelIndex
            labelArray(pageNumber) = labelArray(pageNumber).Replace("{PAGE}", pageNumber)
            labelArray(pageNumber) = labelArray(pageNumber).Replace("{PAGEN}", pages)
            labelArray(pageNumber) = labelArray(pageNumber).Replace(vbCrLf & vbCrLf, vbCrLf)

            ' hard coded to set modecraft to Burlington
            labelArray(pageNumber) = labelArray(pageNumber).Replace("MODECRAFT FASHIONS", "Burlington Coat Factory")

            wklabelTemplate &= labelArray(pageNumber)
        Next

        Return wklabelTemplate
    End Function

    Public Shared Function SendToLabelPrinter(ByVal labelData As String, Optional ByVal PrinterName As String = "") As Boolean
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then

            If ASCMAIN1.Running_in_VS Then Stop ' TO PRINT TO ZEBRA VIA IP PRINTING DO NOT RETURN ON NEXT LINE

            If PrinterName.Contains(":") And PrinterName.StartsWith("192.168.") Then
            Else
                Return PrintShippingLabelForVandale(labelData, PrinterName)
            End If
            ' the above line was restored so that Doug can print KOHLS labels on the Avery printer
            ' probably need an if stmt for KOHLS and WALMART

            ' SHOULD BE SENDING IN PRINTER_CODE, OR MAYBE rowASTPRNT1, AND NOT JUST PrinterName
            Dim PRINTER_CODE As String = PrinterName
            Dim PRINTER_NAME As String = PrinterName
            Dim PRINTER_PORT As String = PRINTER_NAME

            '    ASCMAIN1.sql = "Select * from ASTPRNT1 where PRINTER_CODE = :PARM1"
            ASCMAIN1.sql = "Select * from ASTPRNT1 where PRINTER_NAME = :PARM1"
            Dim rowASTPRNT1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {PRINTER_NAME})
            If rowASTPRNT1 IsNot Nothing Then
                PRINTER_PORT = rowASTPRNT1.Item("PRINTER_PORT") & ""
            End If

            'PRINTER_NAME = PRINTER_PORT ' FOR VANDALE, USE IP:PORT FOR PRINTER_NAME

            Using ipp As New nsoftware.IPWorks.Ipport
                ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")
                ipp.Connect(Split(PRINTER_PORT, ":")(0), Val(Split(PRINTER_PORT, ":")(1)))

                Dim array() As Byte = System.Text.Encoding.ASCII.GetBytes(labelData)
                ipp.Send(array)
                ipp.Disconnect()
            End Using

        Else
            Try

                If ASCMAIN1.Running_in_VS AndAlso 1 = 2 Then
                    PrintShippingLabelFromDevMachine(labelData)
                ElseIf PrinterName.Length > 0 AndAlso PrinterName.Contains(":") AndAlso PrinterName.Split(":").Length = 2 Then
                    Using ipp As New nsoftware.IPWorks.Ipport
                        ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")
                        ipp.Connect(Split(PrinterName, ":")(0), Val(Split(PrinterName, ":")(1)))

                        Dim array() As Byte = System.Text.Encoding.ASCII.GetBytes(labelData)
                        ipp.Send(array)
                        ipp.Disconnect()
                    End Using
                Else
                    ASCMAIN1.LabelPrinterSerialPort.WriteLine(labelData)
                End If

                Return True
            Catch ex As Exception
                MessageBox.Show(String.Format("Error Printing Label: {0}", ex.Message))
                Return False
            End Try
        End If
    End Function

    Private Shared Function PrintShippingLabelForVandale(ByVal labelData As String, ByVal PrinterName As String) As Boolean
        Dim zebraPrinter As String = PrinterName

        If zebraPrinter.Length = 0 Then
            zebraPrinter = FindZebraPrinter()
        End If

        Dim vLabelPrinter As New ASCPRINT
        Return vLabelPrinter.SendStringToPrinter(zebraPrinter, labelData)
    End Function

    Private Shared Function PrintShippingLabelFromDevMachine(ByVal labelData As String) As Boolean
        If ASCMAIN1.USER_ID <> "edz" AndAlso ASCMAIN1.USER_ID <> "wjz" AndAlso ASCMAIN1.USER_ID <> "wayne" Then
            If MessageBox.Show(labelData, "Continue with label print?", MessageBoxButtons.YesNo) = DialogResult.No Then
                Return True
            End If
        End If

        Dim zebraPrinter As String = String.Empty
        If ASCMAIN1.LabelPrinterName.Length > 0 Then
            zebraPrinter = ASCMAIN1.LabelPrinterName
        Else
            zebraPrinter = FindZebraPrinter()
        End If

        Dim vLabelPrinter As New ASCPRINT
        Return vLabelPrinter.SendStringToPrinter(zebraPrinter, labelData)
    End Function

    Private Shared Function FindZebraPrinter() As String

        If ASCMAIN1.DBS_COMPANY = "VAN" Then
            If ASCMAIN1.DBS_SERVER = "" Then
                'Return "ZDesigner ZM400 200 dpi (ZPL) (Copy 1)"
                'Return "Monarch 9855 203dpi (USB003)"
                'Return "Monarch 9855 203dpi (USB002)"
                Return "ZD420" '  "Monarch 9855 203dpi NYC"
            Else
                If ASCMAIN1.USER_ID = "wayne" Then
                    'Stop
                    'Return "ZDesigner ZM400 200 dpi (ZPL) (Vandale)"
                    Return "Monarch 9855 203dpi NYC"
                Else
                    Return "Monarch 9855 300dpi"
                End If


                'Return "Monarch 9855 300dpi"
                'For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                '    If printerName.ToUpper.StartsWith("ZDESIGNER") Then
                '        Return printerName
                '    End If
                'Next printerName
            End If
        Else
            For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                If printerName.ToUpper.StartsWith("ZEBRA") Then
                    Return printerName
                End If
            Next printerName

            For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                If printerName.ToUpper.StartsWith("ZP") Then
                    Return printerName
                End If
            Next printerName
        End If

        Return ""
    End Function

End Class

Public Class CartonLabel
    Inherits ShippingLabel

    Public CartonError As String
    Private Property CartonNo
    Private LabelTemplateName As String = String.Empty

    Public Sub New(ByVal cartonNo As String, ByVal LabeltemplateName As String)
        Me.CartonNo = cartonNo
        Me.LabelTemplateName = LabeltemplateName
    End Sub

    Public Sub New(ByVal cartonNo As String)
        Me.CartonNo = cartonNo
        LabelTemplateName = String.Empty
    End Sub

    Protected Overrides Function GetLabelData() As Dictionary(Of String, DataRow)
        ' Dim CUST_CODE As String = ""

        ASCMAIN1.sql = "SELECT X.*,WH1.*,O5.*,SUBSTR(O5.CUST_ZIP_CODE,1,5) CUST_ZIP_CODE_5,AC1.CUST_VEND_REF," & vbCrLf _
                    & " NVL(ET1.EDI_SUPPLIER_NO,AC1.CUST_VEND_REF) VENDOR_ID, SUBSTR(NVL(ET1.EDI_SUPPLIER_NO,AC1.CUST_VEND_REF),1,6) VENDOR_ID6" & vbCrLf _
                    & ", SUBSTR(X.CART_NO,11,9) CART_NO_9," & vbCrLf _
                    & " SUBSTR(X.CART_NO,20,1) CART_NO_DIGIT," & vbCrLf _
                    & " SUBSTR(X.CART_NO,5,6) CART_NO_PFX," & vbCrLf _
                    & " NVL(ET1.EDI_PO_TYPE,' ') EDI_PO_TYPE, NVL(ET1.EDI_DEPT_DESC,' ') EDI_DEPT_DESC," & vbCrLf _
                    & " TRUNC(SYSDATE) CURRENT_DATE, " & vbCrLf _
                    & " AC2.CUST_NAME CUST_STORE_NAME, AC2.CUST_ADDR_NAME CUST_STORE_ADDR_NAME, AC2.CUST_ADDR1 CUST_STORE_ADDR1, AC2.CUST_ADDR2 CUST_STORE_ADDR2, " & vbCrLf _
                    & " AC2.CUST_CITY CUST_STORE_CITY, AC2.CUST_STATE CUST_STORE_STATE, AC2.CUST_ZIP_CODE CUST_STORE_ZIP_CODE,AC2.CUST_ADDR_GROUP," & vbCrLf _
                    & " X.CART_SERIAL_NO || ' of ' || X.CART_SEQ_MAX CART_1_OF_9,ET1.EDI_PO_RELEASE_NO FROM" & vbCrLf _
                    & " (SELECT ROW_NUMBER() OVER (ORDER BY C1.CART_NO) CART_SERIAL_NO,C1.CART_NO,C1.PICK_NO,O1.EDI_DOC_SEQ_NO,C1.CART_TOTAL_UNITS, " & vbCrLf _
                    & " COUNT(*) OVER () CART_SEQ_MAX,SUM(C2.QTY_PACKED) CART_QTY_PACKED, RPAD(MAX(C1.PKG_CODE),35,' ') PKG_CODE, " & vbCrLf _
                    & " MAX(IS1.STYLE_CODE) STYLE_CODE," & vbCrLf _
                    & " NVL(MAX(O2.STYLE_CODE_SUB),MAX(IS1.STYLE_CODE)) STYLE_CODE_SUB," & vbCrLf _
                    & " MAX(IS1.STYLE_CODE || IC1.COLOR_CODE) STYLE_COLOR_CODE, " & vbCrLf _
                    & " MAX(IC1.COLOR_CODE) COLOR_CODE, MAX(IC1.COLOR_DESC) COLOR_DESC, MAX(IS1.STYLE_DESC) STYLE_DESC," & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN 'Mixed' ELSE MAX(NVL(O2.CUST_UPC,ISC.UPC_CODE)) END UPC_CODE, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN '' ELSE MAX(NVL(O2.CUST_UPC,ISC.UPC_CODE)) END UPC_CODE_ONLY, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN 'Pick And Pack' ELSE MAX(IS1.STYLE_DESC || ' ' || IC1.COLOR_DESC) END STYLE_COLOR_DESC, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN '' ELSE TO_CHAR(SUM(C2.QTY_PACKED)) END CART_QTY_PACKED_ONLY," & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN '' ELSE TO_CHAR(MAX(ET2.EDI_PO4_QTY)) END EDI_PO4_QTY, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN '' ELSE MAX(ET2.EDI_SKU) END EDI_SKU, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN '' ELSE MAX(ET2.EDI_STYLE) END EDI_STYLE, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT O2.STYLE_CODE) > 1 THEN '' ELSE MAX(O2.STYLE_CODE) END ORDR_STYLE, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT O2.CUST_SKU) > 1 THEN '' ELSE MAX(O2.CUST_SKU) END CUST_SKU," & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT O2.CUST_SKU) > 1 THEN 'Mixed' ELSE MAX(O2.CUST_SKU) END CUST_SKU_MIX, " & vbCrLf _
                    & " CASE WHEN COUNT(DISTINCT IS1.STYLE_CODE || IC1.COLOR_CODE) > 1 THEN 'Mixed' ELSE TO_CHAR(C1.CART_TOTAL_UNITS) END CART_TOTAL_UNITS_MIX, " & vbCrLf _
                    & " MAX(O2.CUST_STYLE_CODE) CUST_STYLE_CODE, " & vbCrLf _
                    & " MAX(O2.CUST_COLOR_CODE) CUST_COLOR_CODE, MAX(C1.CART_TOTAL_WGT_CALC) CART_TOTAL_WGT_CALC, " & vbCrLf _
                    & " MAX(O1.ORDR_CUST_PO) ORDR_CUST_PO, " & vbCrLf _
                    & " MAX(O1.ORDR_NO) ORDR_NO, " & vbCrLf _
                    & " MAX(O1.CUST_CODE) CUST_CODE, " & vbCrLf _
                    & " MAX(O1.WHSE_CODE) WHSE_CODE, " & vbCrLf _
                    & " MAX(O1.CUST_STORE_NO) CUST_STORE_NO, " & vbCrLf _
                    & " MAX(O1.ORDR_SHIP_DATE) ORDR_SHIP_DATE, " & vbCrLf _
                    & " MAX(O1.EDI_MERCH_TYPE) EDI_MERCH_TYPE, " & vbCrLf _
                    & " MAX(CNTRY.COUNTRY_NAME) COUNTRY_NAME, MAX(O1.ORDR_DEPT) ORDR_DEPT from" & vbCrLf _
                    & " SOTCART1 C1 JOIN SOTCART2 C2 ON (C1.CART_NO=C2.CART_NO) JOIN " & vbCrLf _
                    & " SOTORDR1 O1 ON (C2.ORDR_NO=O1.ORDR_NO) JOIN " & vbCrLf _
                    & " SOTORDR2 O2 ON (C2.ORDR_NO=O2.ORDR_NO AND C2.ORDR_LNO=O2.ORDR_LNO) JOIN " & vbCrLf _
                    & " ICTSTYL1 IS1 ON (C2.STYLE_CODE=IS1.STYLE_CODE) LEFT JOIN " & vbCrLf _
                    & " TATCNTRY CNTRY ON (IS1.COUNTRY_CODE=CNTRY.COUNTRY_CODE) JOIN" & vbCrLf _
                    & " ICTCOLR1 IC1 ON (C2.COLOR_CODE=IC1.COLOR_CODE) JOIN " & vbCrLf _
                    & " ICTSTYC1 ISC ON (C2.STYLE_CODE=ISC.STYLE_CODE AND C2.COLOR_CODE=ISC.COLOR_CODE) LEFT JOIN" & vbCrLf _
                    & " EDT850T2 ET2 ON (O2.EDI_DOC_SEQ_NO=ET2.EDI_DOC_SEQ_NO AND O2.EDI_DTL_SEQ=ET2.EDI_DTL_SEQ)" & vbCrLf _
                    & " WHERE C1.PICK_NO=(SELECT PICK_NO FROM SOTCART1 WHERE CART_NO=:PARM1) " & vbCrLf _
                    & " GROUP BY C1.CART_NO,C1.PICK_NO,O1.EDI_DOC_SEQ_NO,C1.CART_TOTAL_UNITS) X" & vbCrLf _
                    & " JOIN ICTWHSE1 WH1 ON (X.WHSE_CODE=WH1.WHSE_CODE) " & vbCrLf _
                    & " JOIN SOTORDR5 O5 ON (X.ORDR_NO=O5.ORDR_NO AND O5.CUST_ADDR_TYPE='ST') " & vbCrLf _
                    & " JOIN ARTCUST1 AC1 ON (X.CUST_CODE=AC1.CUST_CODE)" & vbCrLf _
                    & " LEFT JOIN EDT850T1 ET1 ON (X.EDI_DOC_SEQ_NO=ET1.EDI_DOC_SEQ_NO)" & vbCrLf _
                    & " LEFT JOIN ARTCUST2 AC2 ON (X.CUST_CODE=AC2.CUST_CODE AND X.CUST_STORE_NO=AC2.CUST_ADDR_CODE AND AC2.CUST_ADDR_TYPE='MK')" & vbCrLf _
                    & " WHERE CART_NO=:PARM1"

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "ET1.EDI_PO_RELEASE_NO", " NULL EDI_PO_RELEASE_NO")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "AC2.CUST_ADDR_GROUP", " NULL CUST_ADDR_GROUP")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "MAX(O2.CUST_COLOR_CODE) CUST_COLOR_CODE,", " MAX(O2.CUST_COLOR_CODE) CUST_COLOR_CODE, MAX(O2.CUST_SIZE_CODE) CUST_SIZE_CODE,")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "MAX(O1.EDI_MERCH_TYPE) EDI_MERCH_TYPE", "MAX(O1.EDI_MERCH_TYPE) EDI_MERCH_TYPE, MAX(NVL(IS1.CASE_CUBE,0)) CASE_CUBE")
        End If

        Dim rowSOTCART1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, True, "V", New Object() {CartonNo})
        rowSOTCART1.Table.Columns.Add("CUST_STYLE_CODE_SOTCSTY1", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("INNER_PACK_QTY", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("CARTON_PACK_QTY", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("TOTAL_INNER_PACKS", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("PARTIAL_CASE", GetType(System.String))
        rowSOTCART1.Table.Columns("EDI_PO_TYPE").MaxLength = 50
        rowSOTCART1.Table.Columns.Add("CUST_ADDR_NAME", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("GTIN_CODE", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("SHIP_VIA_DESC", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("BILL_OF_LADING_NO", GetType(System.String))
        rowSOTCART1.Table.Columns.Add("SHIP_REF", GetType(System.String))

        Dim CUST_CODE As String = rowSOTCART1.Item("CUST_CODE") & String.Empty
        Dim STYLE_CODE As String = rowSOTCART1.Item("STYLE_CODE") & String.Empty
        Dim CART_NO As String = rowSOTCART1.Item("CART_NO") & String.Empty
        Dim GTIN_CODE As String = ""

        If rowSOTCART1("UPC_CODE_ONLY") & "" <> "" And CUST_CODE = "WALMARTCOM" Then
            Select Case Val(rowSOTCART1("CART_TOTAL_UNITS") + 0)
                Case 2
                    GTIN_CODE = "10"
                Case 4
                    GTIN_CODE = "20"
                Case 6
                    GTIN_CODE = "30"
                Case 8
                    GTIN_CODE = "40"
                Case 12, 18
                    GTIN_CODE = "50"
                Case 24
                    GTIN_CODE = "60"
                Case 36
                    GTIN_CODE = "70"
                Case 48
                    GTIN_CODE = "80"
                Case Else
                    MsgBox("Walmartcom GTIN Error in label, label not printed!", MsgBoxStyle.Critical, "Label Error")
                    Throw New NotImplementedException("GTIN Error")
            End Select
            GTIN_CODE &= rowSOTCART1("UPC_CODE_ONLY") & ""
            If GTIN_CODE.Length = 14 Then
                GTIN_CODE = GTIN_CODE.Substring(0, 13)
            End If
        End If
        rowSOTCART1("GTIN_CODE") = GTIN_CODE

        ASCMAIN1.sql = "select CUST_ADDR_NAME from ARTCUST2 where CUST_CODE = :PARM1 and CUST_ADDR_CODE = :PARM2"
        rowSOTCART1.Item("CUST_ADDR_NAME") = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New Object() {CUST_CODE, rowSOTCART1.Item("CUST_ADDR_CODE") & String.Empty}) & String.Empty

        ASCMAIN1.sql = "SELECT MAX(SOTCSTY1.CUST_STYLE_CODE) 
                FROM SOTCSTY1 
                WHERE CUST_CODE = :PARM1
                AND (STYLE_CODE, COLOR_CODE) IN (SELECT STYLE_CODE, COLOR_CODE FROM SOTCART2 WHERE CART_NO = :PARM2 AND STYLE_CODE = :PARM3)"
        rowSOTCART1.Item("CUST_STYLE_CODE_SOTCSTY1") = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VVV", New Object() {CUST_CODE, CART_NO, STYLE_CODE}) & String.Empty

        Dim rowICTSTYL1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTSTYL1 WHERE STYLE_CODE = :PARM1", True, "V", New Object() {STYLE_CODE})
        rowSOTCART1.Item("INNER_PACK_QTY") = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & String.Empty)
        rowSOTCART1.Item("CARTON_PACK_QTY") = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & String.Empty)
        rowSOTCART1.Item("TOTAL_INNER_PACKS") = 0

        If Val(rowSOTCART1.Item("INNER_PACK_QTY") & String.Empty) > 0 Then
            rowSOTCART1.Item("TOTAL_INNER_PACKS") = Math.Round(Val(rowSOTCART1.Item("CARTON_PACK_QTY") & String.Empty) / Val(rowSOTCART1.Item("INNER_PACK_QTY") & String.Empty), 0)
        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI" Then
            If CUST_CODE = "230058" Then
                If Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & String.Empty) > Val(rowSOTCART1.Item("CART_TOTAL_UNITS_MIX") & String.Empty) Then
                    rowSOTCART1.Table.Columns("PARTIAL_CASE").ReadOnly = False
                    rowSOTCART1.Item("PARTIAL_CASE") = "Partial Case, EA"
                End If
            End If
        End If
        ' NEEDTO GET RID OF THIS TOO
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then  'If I see one more "If VAN Then" I am going to throw up...
            For Each dc In rowSOTCART1.Table.Columns
                dc.ReadOnly = False
            Next
            'would like to Remove the code below, when template name is not the same as the customer we lose the customer code, why even try to get the customer again?
            'adding a condition to search only if cust_code is empty
            If String.IsNullOrEmpty(CUST_CODE) Then
                Dim S As New Text.StringBuilder With {.Length = 0}
                S.AppendLine("SELECT AC.CUST_CODE FROM")
                S.AppendLine("SOTCART1 C1 JOIN")
                S.AppendLine("SOTPICK1 P1 ON (C1.PICK_NO=P1.PICK_NO) JOIN")
                S.AppendLine("SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) JOIN")
                S.AppendLine("ARTCUST1 AC ON (O1.CUST_CODE=AC.CUST_CODE) JOIN")
                S.AppendLine("SOTUCCL1 U1 ON (AC.CUST_CODE=U1.LABEL_TEMPLATE_CODE)")
                S.AppendLine("WHERE C1.CART_NO=:PARM1")
                CUST_CODE = ASCDATA1.GetDataValue(S.ToString, "V", New Object() {CartonNo}) & ""
            End If
            If CUST_CODE = "WALMART" And rowSOTCART1.Item("PKG_CODE") & "" = "" Then
                ASCMAIN1.sql = "select max(SOTORDR9.RANGE_STYLE_DESC) from SOTCART2, SOTORDR2, SOTORDR9" & vbCrLf _
                    & " where CART_NO = :PARM1" & vbCrLf _
                    & " and SOTCART2.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                    & " and SOTCART2.ORDR_LNO = SOTORDR2.ORDR_LNO" & vbCrLf _
                    & " and SOTCART2.ORDR_NO = SOTORDR9.ORDR_NO" & vbCrLf _
                    & " and SOTORDR2.RANGE_STYLE_LNO = SOTORDR9.RANGE_STYLE_LNO"
                rowSOTCART1.Item("PKG_CODE") = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {CartonNo}) & ""
            End If
        End If

        Dim PICK_NO As String = rowSOTCART1.Item("PICK_NO") & ""
        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Select nvl(PICK_NO_CONS,PICK_NO) PICK_NO from SOTPICK1 where PICK_NO = :PARM1"
            PICK_NO = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {PICK_NO}) & ""

            If CUST_CODE = "COSTCOUS" Then
                ASCMAIN1.sql = $"select SHIP_REF, BILL_OF_LADING_NO, SHIP_VIA_DESC from SOTCART1, SOTPICK1, SOTSHIP1, SOTSVIA1
                                where SOTSVIA1.SHIP_VIA_CODE = SOTSHIP1.SHIP_VIA_CODE 
                                AND SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                                AND SOTPICK1.PICK_NO = SOTCART1.PICK_NO
                                AND SOTCART1.CART_NO = '{CART_NO}'"
                Dim rowCarrier As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                rowSOTCART1("SHIP_REF") = rowCarrier.Item("SHIP_REF") & String.Empty
                rowSOTCART1("BILL_OF_LADING_NO") = rowCarrier.Item("BILL_OF_LADING_NO") & String.Empty
                rowSOTCART1("SHIP_VIA_DESC") = rowCarrier.Item("SHIP_VIA_DESC") & String.Empty
            End If

        End If

        Dim sqlMultiPO As String = " (SELECT PICK_NO_CONS" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 1 THEN ORDR_CUST_PO ELSE NULL END) ORDR_CUST_PO_1" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 2 THEN ORDR_CUST_PO ELSE NULL END) ORDR_CUST_PO_2" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 3 THEN ORDR_CUST_PO ELSE NULL END) ORDR_CUST_PO_3" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 4 THEN ORDR_CUST_PO ELSE NULL END) ORDR_CUST_PO_4" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 5 THEN ORDR_CUST_PO ELSE NULL END) ORDR_CUST_PO_5" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 1 THEN ORDR_DEPT ELSE NULL END) ORDR_DEPT_1" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 2 THEN ORDR_DEPT ELSE NULL END) ORDR_DEPT_2" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 3 THEN ORDR_DEPT ELSE NULL END) ORDR_DEPT_3" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 4 THEN ORDR_DEPT ELSE NULL END) ORDR_DEPT_4" & vbCrLf _
            & ", MAX(CASE WHEN ROW_NO = 5 THEN ORDR_DEPT ELSE NULL END) ORDR_DEPT_5" & vbCrLf _
            & "FROM (" & vbCrLf _
            & "SELECT X.*, ROWNUM ROW_NO FROM (" & vbCrLf _
            & "SELECT SOTPICK1.PICK_NO_CONS, SOTORDR1.EDI_MERCH_TYPE, SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO" & vbCrLf _
            & ", SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DEPT, SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
            & " from SOTPICK1,SOTORDR1" & vbCrLf _
            & " where SOTPICK1.PICK_NO_CONS = '" & PICK_NO & "'" & vbCrLf _  ' I think this may cause a problem maybe PICK_NO_CONS
            & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
            & " order by CASE WHEN SOTPICK1.PICK_NO_CONS = SOTPICK1.PICK_NO THEN 0 ELSE 1 END" & vbCrLf _
            & ") X) Y GROUP BY PICK_NO_CONS) M" & vbCrLf

        Dim sqlMultiPOjoin As String = " and M.PICK_NO (+) = NVL(SOTPICK1.PICK_NO_CONS,'?') " & vbCrLf

        sqlMultiPOjoin = " RIGHT JOIN " & sqlMultiPO & " ON (M.PICK_NO_CONS = NVL(SOTPICK1.PICK_NO_CONS,'?'))"
        sqlMultiPO = ""

        Dim sqlMultiPOcolumns = "SOTCART1.CART_NO, SOTCART1.PKG_CODE, M.ORDR_DEPT_1,M.ORDR_DEPT_2,M.ORDR_DEPT_3,M.ORDR_DEPT_4,M.ORDR_DEPT_5," _
                                & " M.ORDR_CUST_PO_1,M.ORDR_CUST_PO_2,M.ORDR_CUST_PO_3,M.ORDR_CUST_PO_4,M.ORDR_CUST_PO_5,"

        ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.CUST_STORE_NO, SUBSTR(SOTSHIP1.BILL_OF_LADING_NO,-10) SHIP_BOL_NO_LAST_10, " & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN", sqlMultiPOcolumns, "") _
                    & " NVL(SOTORDR0.ORDR_DEPT,SOTSHIP1.ORDR_DEPT) ORDR_DEPT,SOTSHIP1.*,SOTORDR0.CUST_CODE,SOTORDR0.ORDR_CUST_PO, " & vbCrLf _
                    & " COALESCE(SOTSVIA1.SHIP_VIA_DESC,SOTSHIP1.SHIP_VIA_CODE,'') SHIP_VIA_DESC, " & vbCrLf _
                    & " SOTSVIA1.SHIP_VIA_SCAC, SOTSHIP1.SHIP_REF," & vbCrLf _
                    & " SUBSTR(SOTORDR1.CUST_STORE_NO,-4) CUST_STORE_NO_4," & vbCrLf _
                    & " SUBSTR(SOTORDR1.CUST_STORE_NO,-5) CUST_STORE_NO_5," & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN",
                          " SOTORDR1.CUST_STORE_NO as CUST_STORE_NO_X,",
                          " SUBSTR(SOTORDR1.CUST_STORE_NO,-1 * EDTSLSP1.NUMBER_CHARS_STORE) CUST_STORE_NO_X,") & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN",
                          " SUBSTR(SOTORDR1.CUST_DC_NO,-1 * NVL(EDTSLSP1.NUMBER_CHRS_DC,0)) CUST_DC_NO_X,",
                          " SUBSTR(SOTORDR1.CUST_DC_NO,-1 * NVL(EDTSLSP1.NUMBER_CHARS_DC,0)) CUST_DC_NO_X,") & vbCrLf _
                    & " SUBSTR(SOTORDR1.CUST_DC_NO,-4) CUST_DC_NO_4," & vbCrLf _
                    & " SOTORDR1.EDI_MERCH_TYPE," & vbCrLf _
                    & " SOTORDR1.CUST_DC_NO," & vbCrLf _
                    & " SOTORDR1.CUST_STORE_NO," & vbCrLf _
                    & " COALESCE(SOTSHIP1.SHIP_DATE_PLANNED,TRUNC(SYSDATE)) SHIP_DATE" & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN", " from SOTCARM1 SOTCART1 ", " from SOTCART1 ") & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN", sqlMultiPO, "") _
                    & " JOIN SOTPICK1 ON (SOTCART1.PICK_NO=SOTPICK1.PICK_NO)" & vbCrLf _
                    & " JOIN SOTORDR1 ON (SOTPICK1.ORDR_NO=SOTORDR1.ORDR_NO) " & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN",
                          " LEFT JOIN EDT850T1 ON (EDT850T1.EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO)" & vbCrLf,
                          "") _
                    & " JOIN SOTSHIP1 ON (SOTPICK1.SHIP_BOL_NO=SOTSHIP1.SHIP_BOL_NO) " & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN",
                          " LEFT JOIN EDTSLSP1 ON (SOTORDR1.CUST_CODE=EDTSLSP1.CUST_CODE and EDTSLSP1.EDI_TP_QUAL = EDT850T1.EDI_TP_QUAL and EDTSLSP1.EDI_TP_ID = EDT850T1.EDI_TP_ID)",
                          " LEFT JOIN EDTSLSP1 ON (SOTORDR1.CUST_CODE=EDTSLSP1.CUST_CODE)") & vbCrLf _
                    & " LEFT JOIN SOTORDR0 ON (SOTSHIP1.ORDR_GROUP_NO=SOTORDR0.ORDR_GROUP_NO) " & vbCrLf _
                    & " LEFT JOIN SOTSVIA1 ON (SOTSHIP1.SHIP_VIA_CODE=SOTSVIA1.SHIP_VIA_CODE) " & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "VAN", sqlMultiPOjoin, "") _
                    & " where " & vbCrLf _
                    & " SOTCART1.CART_NO = :PARM1"

        'If ASCMAIN1.CLIENT = "VAN" Then
        '    ASCMAIN1.sql = ASCMAIN1.sql.Replace("", "")
        'End If

        Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, True, "V", New Object() {CartonNo})

        If ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI" Then
            If CUST_CODE = "230058" Then
                If rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty = String.Empty Then
                    rowSOTPICK1.Table.Columns("ORDR_CUST_PO").ReadOnly = False
                    rowSOTPICK1.Item("ORDR_CUST_PO") = "PO # Unavailable"
                End If
            End If
        End If

        Dim labelData As New Dictionary(Of String, DataRow)
        labelData.Add("SOTCART1", rowSOTCART1)
        labelData.Add("SOTPICK1", rowSOTPICK1)
        labelData.Add("SOTORDR1", rowSOTCART1)
        labelData.Add("SOTSHIPX", rowSOTPICK1)
        labelData.Add("ICTWHSE1", rowSOTCART1)
        labelData.Add("SOTORDR5", rowSOTCART1)
        labelData.Add("ARTCUST1", rowSOTCART1)

        If ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.sql = "SELECT CART_NO, STYLE_CODE, COLOR_CODE, SUM(QTY_PACKED) QTY_PACKED FROM SOTCART2 WHERE CART_NO = :PARM1 GROUP BY CART_NO, STYLE_CODE, COLOR_CODE"
            Dim tblSOTCART2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTCART2", "V", New Object() {CartonNo})
            Dim ictr As Int32 = 1
            For Each rowSOTCART2 As DataRow In tblSOTCART2.Select("")
                labelData.Add($"SOTCART2_{ictr}", rowSOTCART2)
                ictr += 1
            Next
        End If

        ' SEE ABOVE - NEED TO ELIMINATE THAT CODE AND PASS IN CUST_CODE
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then  'If I see one more "If VAN Then" I am going to throw up...
            'If CUST_CODE = "KOHLS" Or CUST_CODE = "WALMART" Then
            VENDORFORMAT(CUST_CODE, rowSOTCART1, labelData)
            'End If
        End If

        Return labelData
    End Function

    Private Sub GetVendorOverideTemplate(ByRef labelTemplate As String)
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("SELECT AC.CUST_CODE FROM")
        S.AppendLine("SOTCART1 C1 JOIN")
        S.AppendLine("SOTPICK1 P1 ON (C1.PICK_NO=P1.PICK_NO) JOIN")
        S.AppendLine("SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) JOIN")
        S.AppendLine("ARTCUST1 AC ON (O1.CUST_CODE=AC.CUST_CODE) JOIN")
        S.AppendLine("SOTUCCL1 U1 ON (AC.LABEL_TEMPLATE_CODE=U1.LABEL_TEMPLATE_CODE)")
        S.AppendLine("WHERE C1.CART_NO='" & CartonNo & "'")
        Dim LABEL_TEMPLATE_CODE As String = ASCDATA1.GetDataValue(S.ToString)

        S.Length = 0
        S.AppendLine("SELECT MIN(ORDR_NO) AS ORDR_NO")
        S.AppendLine("FROM SOTCART2")
        S.AppendLine(String.Format("WHERE CART_NO = '{0}'", CartonNo))
        ASCMAIN1.sql = S.ToString()
        Dim ORDR_NO As String = ASCDATA1.GetDataValue

        S.Length = 0
        S.AppendLine("SELECT COUNT(*) AS RNG_CNT")
        S.AppendLine("FROM SOTORDR9")
        S.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
        ASCMAIN1.sql = S.ToString()
        Dim RNG_CNT As Integer = Val(ASCDATA1.GetDataValue & "")

        Select Case LABEL_TEMPLATE_CODE
            Case Is = "KOHLS"
                S.Length = 0
                S.AppendLine("SELECT NVL(EDI_DEPT_DESC,'NULL')")
                S.AppendLine("FROM EDT850T1")
                S.AppendLine("WHERE EDI_JRNL_NO IN (")
                S.AppendLine("  SELECT O1.EDI_JRNL_NO")
                S.AppendLine("  FROM")
                S.AppendLine("    SOTCART1 C1 JOIN")
                S.AppendLine("    SOTPICK1 P1 ON (C1.PICK_NO = P1.PICK_NO)")
                S.AppendLine("    JOIN")
                S.AppendLine("    SOTORDR1 O1 ON (P1.ORDR_NO = O1.ORDR_NO)")
                S.AppendLine("    JOIN")
                S.AppendLine("    ARTCUST1 AC ON (O1.CUST_CODE = AC.CUST_CODE)")
                S.AppendLine("    JOIN")
                S.AppendLine("    SOTUCCL1 U1 ON (AC.LABEL_TEMPLATE_CODE = U1.LABEL_TEMPLATE_CODE)")
                S.AppendLine(String.Format("  WHERE C1.CART_NO = '{0}'", CartonNo))
                S.AppendLine(")")
                ASCMAIN1.sql = S.ToString()
                Dim EDI_DEPT_DESC As String = ASCDATA1.GetDataValue
                'MsgBox("Switching Of Kohls Label types Need To be Tested!", vbCritical, "Kohls Labels")
                'Stop
                If EDI_DEPT_DESC = "BULK" Then
                    LABEL_TEMPLATE_CODE = "KOHLS2"
                End If
                labelTemplate = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", LABEL_TEMPLATE_CODE)) & ""

            Case Is = "WALMART"
                ASCMAIN1.sql = "Select EDT850T1.EDI_CONS_NO from EDT850T1,SOTORDR1,SOTPICK1,SOTCART1 where SOTCART1.CART_NO = :PARM1 and SOTPICK1.PICK_NO = SOTCART1.PICK_NO and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO and EDT850T1.EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO"
                Dim EDI_CONS_NO As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {CartonNo})
                If EDI_CONS_NO <> "" Then
                    LABEL_TEMPLATE_CODE = "WALMARTCON"
                    labelTemplate = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", LABEL_TEMPLATE_CODE)) & ""
                End If

            Case Is = "MEIJER"
                'If RNG_CNT > 0 Then
                '    S.Length = 0
                '    S.AppendLine("select ordr_cnt from sotordr0")
                '    S.AppendLine(" where ordr_cust_po in (")
                '    S.AppendLine("   select ordr_cust_po from sotordr1")
                '    S.AppendLine(String.Format("   where ordr_no = '{0}')", ORDR_NO))
                '    ASCMAIN1.sql = S.ToString()
                '    Dim ORDR_CNT As Integer = Val(ASCDATA1.GetDataValue & "")
                '    If ORDR_CNT = 1 Then
                '        LABEL_TEMPLATE_CODE = "MEIJERR"
                '        labelTemplate = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", LABEL_TEMPLATE_CODE)) & ""
                '    End If
                'End If
                ' 8/20/2024 - count Range Styles not styles in range
                'S.Length = 0
                'S.AppendLine("select count(1) from SOTCART2")
                'S.AppendLine(String.Format("   where CART_NO = :PARM1", ORDR_NO))
                If RNG_CNT > 0 Then
                    ASCMAIN1.sql = "select count(distinct RANGE_STYLE_CODE) from SOTCART2, SOTORDR2
                                where SOTCART2.CART_NO = :PARM1
                                and SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO
                                and SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO"
                Else
                    ASCMAIN1.sql = "select count(1) from SOTCART2
                                where SOTCART2.CART_NO = :PARM1"
                End If
                Dim UPC_CNT As Integer = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {CartonNo}) & "")
                If UPC_CNT = 1 Then
                    LABEL_TEMPLATE_CODE = "MEIJERR"
                    labelTemplate = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", LABEL_TEMPLATE_CODE)) & ""
                End If


            Case Is = "BURLING", "BURLINMEN"
                ASCMAIN1.sql = "Select SOTORDR1.ORDR_CUST_PO from SOTORDR1,SOTPICK1,SOTCART1 where SOTCART1.CART_NO = :PARM1 and SOTPICK1.PICK_NO = SOTCART1.PICK_NO and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
                Dim CUST_PO As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {CartonNo})
                If CUST_PO.Length = 9 Then
                    LABEL_TEMPLATE_CODE = "BURLING99"
                    labelTemplate = ASCDATA1.GetDataValue(String.Format("SELECT UCC128_COMMANDS FROM  SOTUCCL1 U1  WHERE U1.LABEL_TEMPLATE_CODE='{0}'", LABEL_TEMPLATE_CODE)) & ""
                End If
        End Select
    End Sub
    Protected Overrides Function GetLabelTemplate() As String

        Dim labelTemplate As String = String.Empty

        If LabelTemplateName.Length > 0 Then
            labelTemplate = ASCDATA1.GetDataValue(
                "SELECT UCC128_COMMANDS FROM " &
                " SOTUCCL1 U1 " &
                " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {LabelTemplateName}) & ""
        Else
            labelTemplate = ASCDATA1.GetDataValue(
                "SELECT UCC128_COMMANDS FROM " &
                " SOTCART1 C1 JOIN " &
                " SOTPICK1 P1 ON (C1.PICK_NO=P1.PICK_NO) JOIN " &
                " SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) JOIN " &
                " ARTCUST1 AC ON (O1.CUST_CODE=AC.CUST_CODE) JOIN " &
                " SOTUCCL1 U1 ON (AC.LABEL_TEMPLATE_CODE=U1.LABEL_TEMPLATE_CODE) " &
                " WHERE C1.CART_NO=:PARM1", "V", New Object() {CartonNo}) & ""
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                GetVendorOverideTemplate(labelTemplate)
            End If
        End If

        If labelTemplate = String.Empty AndAlso ASCMAIN1.USER_ID = "edz" AndAlso ASCMAIN1.Running_in_VS Then
            labelTemplate = ASCDATA1.GetDataValue(
                "SELECT UCC128_COMMANDS FROM " &
                " SOTUCCL1 U1 " &
                " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {"WINNERSTJX"}) & ""
        End If

        If labelTemplate = "" Then Throw New Exception("No UCC128 label template assigned for this customer")

        Return labelTemplate
    End Function

    Private Shared Function FillSOTCART2(ByVal CART_NO As String, ByVal CUST_CODE As String, ByVal MaxRows As Integer, ByRef Row As DataRow, ByRef labelData As Dictionary(Of String, DataRow), Optional WarnIfMax As Boolean = False, Optional MaxDescLen As Integer = 0) As String
        Dim RetVal As String = ""
        Dim z As Integer = 0
        Dim S As New Text.StringBuilder With {.Length = 0}

        S.Length = 0
        S.AppendLine(String.Format("Select Count(*) from SOTCART2 where CART_NO = '{0}'", CART_NO))
        ASCMAIN1.sql = S.ToString()
        Dim CART_ROWS As Int16 = Val(ASCDATA1.GetDataValue)
        If CART_ROWS > MaxRows And WarnIfMax Then
            RetVal = "Carton Contents Exceeds Label Maximum Rows Of " & MaxRows
        Else
            Dim SR As New Text.StringBuilder With {.Length = 0}
            SR.AppendLine("SELECT NVL(RANGE_STYLE_CODE,'') AS RANGE_STYLE_CODE")
            SR.AppendLine("FROM SOTORDR2")
            SR.AppendLine("WHERE (ORDR_NO, ORDR_LNO) IN")
            SR.AppendLine("(")
            SR.AppendLine("  SELECT MIN(ORDR_NO), MIN(ORDR_LNO)")
            SR.AppendLine("  FROM SOTCART2")
            SR.AppendLine(String.Format("  WHERE CART_NO = '{0}'", CART_NO))
            SR.AppendLine(")")
            ASCMAIN1.sql = SR.ToString()
            Dim RANGE_STYLE_CODE As String = ASCDATA1.GetDataValue & ""
            If RANGE_STYLE_CODE.Length = 0 Then
                S.Length = 0
                S.AppendLine(String.Format("SELECT * FROM SOTCART1 WHERE CART_NO = '{0}'", CART_NO))
                ASCMAIN1.sql = S.ToString
                Dim rowSOTCART2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                For i As Integer = 1 To MaxRows
                    rowSOTCART2.Table.Columns.Add("STYLE_CODE_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("STYLE_DESC_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("COLOR_CODE_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("SIZE_DESC_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("UPC_CODE_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("QTY_PACKED_" & Format(i, "0#"))
                Next
                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("  C2.STYLE_CODE,")
                S.AppendLine("  S1.STYLE_DESC,")
                S.AppendLine("  NVL(C2.COLOR_CODE, 'AST') AS COLOR_CODE,")
                S.AppendLine("  NVL(C2.SIZE_DESC, 'AST')  AS SIZE_DESC,")
                S.AppendLine("  NVL(C2.UPC_CODE, 'AST')   AS UPC_CODE,")
                S.AppendLine("  C2.QTY_PACKED")
                S.AppendLine("FROM SOTCART2 C2, ICTSTYL1 S1")
                S.AppendLine("WHERE C2.STYLE_CODE = S1.STYLE_CODE")
                S.AppendLine(String.Format("      AND C2.CART_NO = '{0}'", CART_NO))
                Dim tbl As DataTable = ASCDATA1.GetDataTable(S.ToString())
                For Each rowSOTCARTX As DataRow In tbl.Rows
                    z += 1
                    If z > MaxRows Then
                        Exit For
                    End If
                    rowSOTCART2.Item("STYLE_CODE_" & Format(z, "0#")) = rowSOTCARTX.Item("STYLE_CODE").ToString & ""
                    Dim STYLE_DESC As String = rowSOTCARTX.Item("STYLE_DESC").ToString & ""
                    If MaxDescLen > 0 And STYLE_DESC.Length > MaxDescLen Then
                        STYLE_DESC = STYLE_DESC.Substring(0, MaxDescLen - 1)
                    End If
                    rowSOTCART2.Item("STYLE_DESC_" & Format(z, "0#")) = STYLE_DESC
                    rowSOTCART2.Item("COLOR_CODE_" & Format(z, "0#")) = rowSOTCARTX.Item("COLOR_CODE").ToString & ""
                    rowSOTCART2.Item("SIZE_DESC_" & Format(z, "0#")) = rowSOTCARTX.Item("SIZE_DESC").ToString & ""
                    rowSOTCART2.Item("UPC_CODE_" & Format(z, "0#")) = rowSOTCARTX.Item("UPC_CODE").ToString & ""
                    rowSOTCART2.Item("QTY_PACKED_" & Format(z, "0#")) = rowSOTCARTX.Item("QTY_PACKED").ToString & ""
                Next
                If CUST_CODE = "COSTCOUS" And z > 1 Then
                    Row.Item("EDI_SKU") = "Mixed"
                    rowSOTCART2.Item("STYLE_CODE_" & Format(1, "0#")) = "Mixed"
                    rowSOTCART2.Item("STYLE_DESC_" & Format(1, "0#")) = "Mixed"
                    rowSOTCART2.Item("COLOR_CODE_" & Format(1, "0#")) = "AST"
                    rowSOTCART2.Item("SIZE_DESC_" & Format(1, "0#")) = "AST"
                    rowSOTCART2.Item("UPC_CODE_" & Format(1, "0#")) = "AST"
                    rowSOTCART2.Item("QTY_PACKED_" & Format(1, "0#")) = Val(rowSOTCART2.Item("CART_TOTAL_UNITS").ToString & "")
                End If
                labelData.Add("SOTCART2", rowSOTCART2)
            Else
                S.Length = 0
                S.AppendLine(String.Format("SELECT * FROM SOTCART1 WHERE CART_NO = '{0}'", CART_NO))
                ASCMAIN1.sql = S.ToString
                Dim rowSOTCART2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                For i As Integer = 1 To MaxRows
                    rowSOTCART2.Table.Columns.Add("STYLE_CODE_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("STYLE_DESC_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("COLOR_CODE_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("SIZE_DESC_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("UPC_CODE_" & Format(i, "0#"))
                    rowSOTCART2.Table.Columns.Add("QTY_PACKED_" & Format(i, "0#"))
                Next

                If CUST_CODE = "COSTCOUS" Then
                    ASCMAIN1.sql = $"select EDI_SKU, EDI_STYLE_NAME from EDT850T2
                                    where (edi_doc_seq_no, edi_dtl_seq) in (
                                    select edi_doc_seq_no, edi_dtl_seq from SOTORDR2, SOTCART2
                                    where SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO
                                    and SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO
                                    and SOTCART2.CART_NO = '{CART_NO}')"
                    Dim rowEDISKU As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                    If rowEDISKU IsNot Nothing Then
                        Row.Item("EDI_SKU") = rowEDISKU("EDI_SKU") & ""
                        rowSOTCART2.Item("STYLE_CODE_" & Format(1, "0#")) = rowEDISKU("EDI_SKU") & ""
                        rowSOTCART2.Item("STYLE_DESC_" & Format(1, "0#")) = rowEDISKU("EDI_STYLE_NAME") & ""
                        rowSOTCART2.Item("COLOR_CODE_" & Format(1, "0#")) = "AST"
                        rowSOTCART2.Item("SIZE_DESC_" & Format(1, "0#")) = "AST"
                        rowSOTCART2.Item("UPC_CODE_" & Format(1, "0#")) = "AST"
                        rowSOTCART2.Item("QTY_PACKED_" & Format(1, "0#")) = Val(rowSOTCART2.Item("CART_TOTAL_UNITS").ToString & "")
                        labelData.Add("SOTCART2", rowSOTCART2)
                    Else
                        Throw New Exception("Costco Label EDI SKU not found")
                    End If
                Else

                    rowSOTCART2.Item("STYLE_CODE_" & Format(1, "0#")) = RANGE_STYLE_CODE
                    rowSOTCART2.Item("STYLE_DESC_" & Format(1, "0#")) = ""
                    rowSOTCART2.Item("COLOR_CODE_" & Format(1, "0#")) = "AST"
                    rowSOTCART2.Item("SIZE_DESC_" & Format(1, "0#")) = "AST"
                    rowSOTCART2.Item("UPC_CODE_" & Format(1, "0#")) = "AST"
                    rowSOTCART2.Item("QTY_PACKED_" & Format(1, "0#")) = Val(rowSOTCART2.Item("CART_TOTAL_UNITS").ToString & "")
                    labelData.Add("SOTCART2", rowSOTCART2)
                End If
            End If
        End If
        Return RetVal
    End Function
    Private Sub VENDORFORMAT(ByVal CUST_CODE As String, ByRef Row As DataRow, ByRef labelData As Dictionary(Of String, DataRow))
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.Length = 0
        S.AppendLine("SELECT MIN(ORDR_NO) AS ORDR_NO")
        S.AppendLine("FROM SOTCART2")
        S.AppendLine(String.Format("WHERE CART_NO = '{0}'", Row.Item("CART_NO")))
        ASCMAIN1.sql = S.ToString()
        Dim ORDR_NO As String = ASCDATA1.GetDataValue

        S.Length = 0
        S.AppendLine("SELECT COUNT(*) AS RNG_CNT")
        S.AppendLine("FROM SOTORDR9")
        S.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
        ASCMAIN1.sql = S.ToString()
        Dim RNG_CNT As Integer = Val(ASCDATA1.GetDataValue & "")

        Select Case CUST_CODE
            Case Is = "BEDBATH"
                If Row.Item("CUST_ZIP_CODE").ToString.Length > 5 Then
                    Row.Item("CUST_ZIP_CODE") = Row.Item("CUST_ZIP_CODE").ToString.Substring(0, 5)
                End If
            Case Is = "BELKS"
                Dim EDI_DOC_SEQ_NO As String = Row.Item("EDI_DOC_SEQ_NO").ToString
                Dim SQLS As New Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT EDI_DEPT_DESC")
                SQLS.AppendLine("FROM EDT850T1")
                SQLS.AppendLine(String.Format("WHERE EDI_DOC_SEQ_NO = '{0}'", EDI_DOC_SEQ_NO))
                ASCMAIN1.sql = SQLS.ToString()
                Row.Item("WHSE_EDI_ID") = ASCDATA1.GetDataValue
            Case Is = "BURLING", "BURLINMEN"
                Row.Item("CUST_STORE_NO") = Row.Item("CUST_STORE_NO").ToString.Substring(3, 3)
            Case Is = "CHARLOT"
                Dim CART_NO As String = Row.Item("CART_NO").ToString

                Dim SQLS As New Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("Select CART_TOTAL_WGT_CALC")
                SQLS.AppendLine("FROM SOTCART1")
                SQLS.AppendLine(String.Format("WHERE CART_NO = '{0}'", CART_NO))
                ASCMAIN1.sql = SQLS.ToString()
                Row.Item("VENDOR_ID") = ASCDATA1.GetDataValue

                SQLS.Length = 0
                SQLS.AppendLine("SELECT MIN(RANGE_STYLE_CODE) AS RANGE_STYLE_CODE")
                SQLS.AppendLine("FROM SOTORDR2, SOTCART2")
                SQLS.AppendLine("WHERE SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO")
                SQLS.AppendLine("      AND SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO")
                SQLS.AppendLine(String.Format("AND CART_NO = '{0}'", CART_NO))
                ASCMAIN1.sql = SQLS.ToString()
                Dim RANGE_STYLE_CODE As String = ASCDATA1.GetDataValue

                If RANGE_STYLE_CODE.Length > 0 Then
                    Row.Item("STYLE_CODE_SUB") = ASCDATA1.GetDataValue
                    Row.Item("COLOR_DESC") = ""

                    SQLS.Length = 0
                    SQLS.AppendLine("SELECT RANGE_STYLE_QTY_PER_PP")
                    SQLS.AppendLine("FROM SOTORDR9")
                    SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
                    SQLS.AppendLine(String.Format("      AND RANGE_STYLE_CODE = '{0}'", RANGE_STYLE_CODE))
                    ASCMAIN1.sql = SQLS.ToString()
                    Dim RANGE_STYLE_QTY_PER_PP As Integer = Val(ASCDATA1.GetDataValue)
                    If RANGE_STYLE_QTY_PER_PP > 0 And Val(Row.Item("CART_TOTAL_UNITS") > 0) Then
                        Row.Item("CART_TOTAL_UNITS") = Val(Row.Item("CART_TOTAL_UNITS")) / RANGE_STYLE_QTY_PER_PP
                    End If
                Else
                    CartonError = "Non-Range Style CHARLOT Labels Have Not Been Tested Yet!"
                End If
            Case Is = "KMART"
                Dim STYLE_CODE As String = Row.Item("STYLE_CODE")
                Dim SQLS As New Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT MIN(UPC_CODE) AS UPC_CODE")
                SQLS.AppendLine("FROM ICVLUPC1")
                SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                ASCMAIN1.sql = SQLS.ToString()
                Dim UPC_CODE As String = ASCDATA1.GetDataValue
                SQLS.Length = 0
                SQLS.AppendLine("SELECT MIN(GTIN_CODE) AS GTIN_CODE")
                SQLS.AppendLine("FROM ICTGTINT")
                SQLS.AppendLine(String.Format("WHERE GTIN_UPC_CODE = '{0}'", UPC_CODE))
                ASCMAIN1.sql = SQLS.ToString()
                Dim GTIN As String = ASCDATA1.GetDataValue
                SQLS.Length = 0
                SQLS.AppendLine("SELECT GTIN_CODE_NEW")
                SQLS.AppendLine("FROM ICTGTINK")
                SQLS.AppendLine("WHERE GTIN_CODE_OLD = '" & GTIN & "'")
                ASCMAIN1.sql = SQLS.ToString()
                Dim GTIN_NEW As String = ASCDATA1.GetDataValue
                If GTIN_NEW.Length > 0 Then
                    GTIN = GTIN_NEW
                End If
                If GTIN.Length = 0 Then
                    CartonError = "GTIN Missing for Style " & STYLE_CODE
                Else
                    Row.Item("UPC_CODE_ONLY") = GTIN
                End If

                SQLS.Length = 0
                SQLS.AppendLine("SELECT MIN(NVL(E1.EDI_MERCH_TYPE,'NULL')) AS EDI_MERCH_TYPE")
                SQLS.AppendLine("FROM SOTORDR1 S1, EDT850T1 E1")
                SQLS.AppendLine("WHERE S1.EDI_DOC_SEQ_NO = E1.EDI_DOC_SEQ_NO (+)")
                SQLS.AppendLine("AND S1.EDI_JRNL_NO = E1.EDI_JRNL_NO (+)")
                SQLS.AppendLine(String.Format("AND S1.ORDR_NO = '{0}'", ORDR_NO))
                ASCMAIN1.sql = SQLS.ToString()
                Dim EDI_MERCH_TYPE As String = ASCDATA1.GetDataValue
                Select Case EDI_MERCH_TYPE
                    Case Is = "R"
                        'Do Nothing.  This is the default.
                    Case Is = "A"
                        Row.Item("STYLE_CODE") = Row.Item("EDI_SKU")
                    Case Else
                        CartonError = "Invalid EDI Merch Type.  Run Labels In ABSolution Version 1"
                End Select
            Case Is = "KOHLS"
                Row.Item("CUST_STORE_NO") = Row.Item("CUST_STORE_NO").ToString.Substring(1, 5)
                If Row.Item("CUST_ZIP_CODE").ToString.Length > 5 Then
                    Row.Item("CUST_ZIP_CODE") = Row.Item("CUST_ZIP_CODE").ToString.Substring(0, 5)
                End If
            Case Is = "MARSHAL", Is = "COSTCOUS"
                Dim CART_NO As String = Row.Item("CART_NO").ToString
                Dim SQLS As String
                SQLS = $"Select SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_REF, SOTSVIA1.SHIP_VIA_DESC 
                         from SOTSHIP1, SOTPICK1, SOTCART1, SOTSVIA1
                         where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                         and SOTPICK1.PICK_NO = SOTCART1.PICK_NO
                         and SOTSHIP1.SHIP_VIA_CODE=SOTSVIA1.SHIP_VIA_CODE
                         and SOTCART1.CART_NO = '{CART_NO}'"
                Dim TmpRow As DataRow = ASCDATA1.GetDataRow(SQLS)
                Row("BILL_OF_LADING_NO") = TmpRow("BILL_OF_LADING_NO")
                Row("SHIP_REF") = TmpRow("SHIP_REF")
                Row("SHIP_VIA_DESC") = TmpRow("SHIP_VIA_DESC")
                If CUST_CODE = "COSTCOUS" Then
                    Dim CASE_CUBE As Integer = 0
                    SQLS = $"SELECT trunc(NVL(PKG_L,0) * NVL(PKG_W,0) * NVL(PKG_H,0) / 1728) CUBE FROM SOTCART1 WHERE SOTCART1.CART_NO = '{CART_NO}'"
                    'get cubic inches transform to cubic feet / 1728
                    CASE_CUBE = ASCDATA1.GetDataValue(SQLS)
                    Row.Item("CASE_CUBE") = CASE_CUBE
                    Row.Item("cart_total_wgt_calc") = ASCDATA1.GetDataValue($"SELECT CART_TOTAL_WGT_ACTUAL FROM SOTCART1 WHERE SOTCART1.CART_NO = '{CART_NO}'")
                End If

                If Row.Item("CUST_STORE_NO").ToString.Length > 4 Then
                    Row.Item("CUST_STORE_NO") = Row.Item("CUST_STORE_NO").ToString.Substring(Row.Item("CUST_STORE_NO").ToString.Length - 4, 4)
                End If
                Dim Msg As String = FillSOTCART2(CART_NO, CUST_CODE, 3, Row, labelData, , 25)
                If Msg.Length > 0 Then
                    CartonError = Msg
                End If
            Case Is = "MEIJER"
                Dim CUST_STORE_NO As String = Row.Item("CUST_STORE_NO").ToString
                If CUST_STORE_NO.Length >= 3 Then
                    CUST_STORE_NO = CUST_STORE_NO.Substring(CUST_STORE_NO.Length - 3, 3)
                Else
                    CartonError = String.Format("Customer Store Number {0} Not 3 Digits Long!", CUST_STORE_NO)
                End If
                Row.Item("CUST_STORE_NO") = CUST_STORE_NO

                If RNG_CNT > 0 Then
                    Dim CART_NO As String = Row.Item("CART_NO").ToString

                    ASCMAIN1.sql = $"select ICTRSTY1.RANGE_STYLE_CODE, ICTRSTY1.RANGE_UPC_CODE from SOTORDR2, ICTRSTY1
                                    where  ICTRSTY1.CUST_CODE = 'MEIJER'
                                    and ICTRSTY1.RANGE_STYLE_CODE = SOTORDR2.RANGE_STYLE_CODE
                                    and (ORDR_NO, ORDR_LNO) in (SELECT MIN(ORDR_NO), MIN(ORDR_LNO)
                                    from SOTCART2
                                    where CART_NO = '{CART_NO}')"

                    Dim rowRANGE As DataRow = ASCDATA1.GetDataRow
                    Row.Item("STYLE_CODE") = rowRANGE("RANGE_STYLE_CODE") & ""
                    Row.Item("UPC_CODE") = rowRANGE("RANGE_UPC_CODE") & ""
                    Row.Item("UPC_CODE_ONLY") = rowRANGE("RANGE_UPC_CODE") & ""
                End If
            Case Is = "SAMSCLUB"
                Dim MaxRows = 8
                Dim z = 0
                Dim CART_NO As String = Row.Item("CART_NO").ToString
                ASCMAIN1.sql = $"SELECT * FROM SOTCART1
                                WHERE CART_NO = '{CART_NO}'"
                Dim rowSOTCART2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                For i As Integer = 1 To MaxRows
                    rowSOTCART2.Table.Columns.Add("EDI_SKU_" & Format(i, "0#"))
                Next
                ASCMAIN1.sql = $"SELECT SOTORDR2.* FROM SOTCART2, SOTORDR2
                                WHERE CART_NO = '{CART_NO}'
                                AND SOTCART2.ORDR_NO = SOTORDR2.ORDR_NO
                                AND SOTCART2.ORDR_LNO = SOTORDR2.ORDR_LNO"
                Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                For Each rowSOTCARTX As DataRow In tbl.Rows
                    z += 1
                    If z > MaxRows Then
                        Exit For
                    End If
                    rowSOTCART2.Item("EDI_SKU_" & Format(z, "0#")) = rowSOTCARTX.Item("CUST_SKU").ToString & ""
                Next
                labelData.Add("SOTCART2", rowSOTCART2)

            Case Is = "SEARS"
                Dim SQLS As New Text.StringBuilder With {.Length = 0}
                SQLS.Length = 0
                SQLS.AppendLine("SELECT MIN(NVL(E1.EDI_MERCH_TYPE,'NULL')) AS EDI_MERCH_TYPE")
                SQLS.AppendLine("FROM SOTORDR1 S1, EDT850T1 E1")
                SQLS.AppendLine("WHERE S1.EDI_DOC_SEQ_NO = E1.EDI_DOC_SEQ_NO (+)")
                SQLS.AppendLine("AND S1.EDI_JRNL_NO = E1.EDI_JRNL_NO (+)")
                SQLS.AppendLine(String.Format("AND S1.ORDR_NO = '{0}'", ORDR_NO))
                ASCMAIN1.sql = SQLS.ToString()
                Dim EDI_MERCH_TYPE As String = ASCDATA1.GetDataValue

                SQLS.Length = 0
                SQLS.AppendLine("SELECT MIN(NVL(E1.EDI_TP_ID,'NULL')) AS EDI_TP_ID")
                SQLS.AppendLine("FROM SOTORDR1 S1, EDT850T1 E1")
                SQLS.AppendLine("WHERE S1.EDI_DOC_SEQ_NO = E1.EDI_DOC_SEQ_NO (+)")
                SQLS.AppendLine("AND S1.EDI_JRNL_NO = E1.EDI_JRNL_NO (+)")
                SQLS.AppendLine(String.Format("AND S1.ORDR_NO = '{0}'", ORDR_NO))
                ASCMAIN1.sql = SQLS.ToString()
                Dim EDI_TP_ID As String = ASCDATA1.GetDataValue

                If EDI_MERCH_TYPE <> "J1" Or EDI_TP_ID = "6111250011" Then
                    CartonError = String.Format("Not An EDI J1 Order!  Use ABSolution V1!")
                End If
            Case Is = "STEINM"
                Dim CUST_STORE_NO As String = Row.Item("CUST_STORE_NO").ToString
                If CUST_STORE_NO.Length = 6 Then
                    CUST_STORE_NO = CUST_STORE_NO.Substring(2, 4)
                Else
                    CartonError = String.Format("Customer Store Number{0} Not 6 Digits Long!", CUST_STORE_NO)
                End If
                Row.Item("CUST_STORE_NO") = CUST_STORE_NO
            Case Is = "TARGET"
                If Row.Item("EDI_SKU").ToString.Length = 0 Then
                    Dim CART_NO As String = Row.Item("CART_NO").ToString
                    Dim EDI_SKU As String = ""
                    Dim SKU_CNT As Integer = 0
                    Dim SQLPRE As String = "SELECT COUNT(EDT850T2.EDI_SKU) "
                    Dim SQLS As New Text.StringBuilder With {.Length = 0}
                    SQLS.AppendLine("FROM SOTCART2, SOTORDR2, EDT850T2")
                    SQLS.AppendLine("WHERE SOTCART2.ORDR_NO = SOTORDR2.ORDR_NO")
                    SQLS.AppendLine("      AND SOTCART2.ORDR_LNO = SOTORDR2.ORDR_LNO")
                    SQLS.AppendLine("      AND EDT850T2.EDI_DOC_SEQ_NO = SOTORDR2.EDI_DOC_SEQ_NO")
                    SQLS.AppendLine("      AND EDT850T2.EDI_DTL_SEQ = SOTORDR2.EDI_DTL_SEQ")
                    SQLS.AppendLine(String.Format("      AND CART_NO = '{0}'", CART_NO))
                    ASCMAIN1.sql = SQLPRE.ToString & SQLS.ToString()
                    SKU_CNT = Val(ASCDATA1.GetDataValue)
                    If SKU_CNT = 1 Then
                        SQLPRE = "SELECT EDT850T2.EDI_SKU "
                        ASCMAIN1.sql = SQLPRE.ToString & SQLS.ToString()
                        EDI_SKU = ASCDATA1.GetDataValue
                        Row.Item("EDI_SKU") = EDI_SKU
                    Else
                        CartonError = String.Format("More than 1 EDI SKU Found In Carton: {0}", CART_NO)
                    End If
                End If
            Case Is = "WALMART"
                Dim SQLS As New Text.StringBuilder With {.Length = 0}
                Dim CART_NO As String = Row.Item("CART_NO").ToString
                Dim CUST_STORE_NO As String = Row.Item("CUST_STORE_NO").ToString
                If CUST_STORE_NO.Length = 6 Then
                    CUST_STORE_NO = CUST_STORE_NO.Substring(1, 5)
                Else
                    CartonError = String.Format("Customer Store Number{0} Not 6 Digits Long!", CUST_STORE_NO)
                End If
                Row.Item("CUST_STORE_NO") = CUST_STORE_NO

                SQLS.Length = 0
                SQLS.AppendLine("SELECT EDI_MERCH_TYPE")
                SQLS.AppendLine("FROM SOTORDR1")
                SQLS.AppendLine("WHERE ORDR_NO IN (")
                SQLS.AppendLine("  SELECT ORDR_NO")
                SQLS.AppendLine("  FROM SOTPICK1")
                SQLS.AppendLine("  WHERE PICK_NO IN (")
                SQLS.AppendLine("    SELECT PICK_NO")
                SQLS.AppendLine("    FROM SOTCART1")
                SQLS.AppendLine(String.Format("    WHERE CART_NO = '{0}'", CART_NO))
                SQLS.AppendLine("  )")
                SQLS.AppendLine(")")
                ASCMAIN1.sql = SQLS.ToString()
                Dim EDI_MERCH_TYPE As String = ASCDATA1.GetDataValue
                Row.Item("UPC_CODE_ONLY") = EDI_MERCH_TYPE

                ASCMAIN1.sql = "Select EDT850T1.EDI_CONS_NO from EDT850T1,SOTORDR1,SOTPICK1,SOTCART1 where SOTCART1.CART_NO = :PARM1 and SOTPICK1.PICK_NO = SOTCART1.PICK_NO and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO and EDT850T1.EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO"
                Dim EDI_CONS_NO As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {CartonNo})
                If EDI_CONS_NO = "" Then
                    ASCMAIN1.sql = " Select ICTRSTY1.RANGE_STYLE_DESC, EDT850T2.EDI_UPC
                                     from EDT850T2,SOTORDR9,ICTRSTY1, SOTPICK1,SOTCART1 
                                     where SOTCART1.CART_NO = :PARM1 
                                     and SOTPICK1.PICK_NO = SOTCART1.PICK_NO 
                                     and SOTORDR9.ORDR_NO = SOTPICK1.ORDR_NO 
                                     and EDT850T2.EDI_DOC_SEQ_NO = SOTORDR9.EDI_DOC_SEQ_NO
                                     and SOTORDR9.EDI_DTL_SEQ =  EDT850T2.EDI_DTL_SEQ
                                     and ICTRSTY1.CUST_CODE = 'WALMART'
                                     and ICTRSTY1.RANGE_STYLE_CODE = EDT850T2.EDI_STYLE"
                    'testing the following lookup until we are comfortable that it works reliably
                    ASCMAIN1.sql = "Select distinct ICTRSTY1.RANGE_STYLE_DESC, ICTRSTY1.RANGE_UPC_CODE, ICTRSTY1.RANGE_STYLE_CODE
                                    from ICTRSTY1, SOTCART2, SOTORDR2 
                                    where SOTCART2.CART_NO = :PARM1 
                                    and SOTCART2.ORDR_NO = SOTORDR2.ORDR_NO 
                                    and SOTCART2.ORDR_LNO = SOTORDR2.ORDR_LNO
                                    and ICTRSTY1.CUST_CODE = 'WALMART'
                                    and ICTRSTY1.RANGE_SKU = SOTORDR2.RANGE_STYLE_CODE"
                    Dim rowWALMART As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {CartonNo})
                    If Not rowWALMART Is Nothing Then
                        Row.Item("PKG_CODE") = rowWALMART("RANGE_STYLE_DESC")
                        Row.Item("UPC_CODE") = rowWALMART("RANGE_UPC_CODE") 'rowWALMART("EDI_UPC")
                    End If
                End If

                Dim CUST_ADDR_CODE As String = Row.Item("CUST_ADDR_CODE").ToString
                If CUST_ADDR_CODE.Length = 5 Then
                    CUST_ADDR_CODE = CUST_ADDR_CODE.Substring(0, 4)
                Else
                    CartonError = String.Format("Customer Address Code {0} Not 5 Digits Long!", CUST_ADDR_CODE)
                End If
                Row.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
            Case Is = "BLOOMOUT", Is = "MACYSECOM", Is = "MACYS"
                If Row.Item("EDI_PO_TYPE") = "RE" Then
                    Row.Item("EDI_PO_TYPE") = "REPLENISHMENT"
                Else
                    Row.Item("EDI_PO_TYPE") = " "
                End If
                Row.Item("CUST_STORE_NO") = Row.Item("CUST_STORE_NO").ToString.Substring(2)

        End Select
    End Sub

End Class

Public Class AddressLabel
    Inherits ShippingLabel

    Private Property PickNo As String
    Private Property LabelComment As String
    Private Property LabelTemplateCode As String
    Private Property RowSOTORDR5 As DataRow
    Private Property StartPos As Integer = 0
    Private Property TotalLabels As Integer = 0

    Public Sub New(ByVal pickNo As String, ByVal labelComment As String, ByVal labelTemplateCode As String, ByVal rowSOTORDR5 As DataRow)
        Me.PickNo = pickNo
        Me.LabelComment = labelComment
        Me.RowSOTORDR5 = rowSOTORDR5
        Me.LabelTemplateCode = labelTemplateCode
    End Sub

    Protected Overrides Function GetLabelData() As Dictionary(Of String, DataRow)
        'retrieve data for label based off pick #

        ASCMAIN1.sql = "SELECT P1.PICK_NO,P1.SHIP_BOL_NO,O1.CUST_STORE_NO FROM " _
            & " SOTPICK1 P1 JOIN SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) " _
            & " WHERE P1.PICK_NO=:PARM1"
        Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {PickNo})

        ASCMAIN1.sql = "Select SOTSHIP1.*,SOTORDR0.CUST_CODE,SOTORDR0.ORDR_CUST_PO,SOTORDR0.ORDR_DEPT" _
               & ", COALESCE(SOTSVIA1.SHIP_VIA_DESC,SOTSHIP1.SHIP_VIA_CODE,'') SHIP_VIA_DESC" _
               & " from SOTSHIP1,SOTORDR0,SOTSVIA1 " _
               & " where SOTORDR0.ORDR_GROUP_NO (+) = SOTSHIP1.ORDR_GROUP_NO" _
               & "   and SOTSVIA1.SHIP_VIA_CODE (+) = SOTSHIP1.SHIP_VIA_CODE" _
               & "   and SOTSHIP1.SHIP_BOL_NO = :PARM1"
        Dim rowSOTSHIPX As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {rowSOTPICK1.Item("SHIP_BOL_NO")})
        rowSOTSHIPX.Table.Columns.Add("ORDR_COMMENTS")
        rowSOTSHIPX.Table.Columns.Add("ADDR_1_OF_9")
        rowSOTSHIPX.Item("ORDR_COMMENTS") = LabelComment

        ASCMAIN1.sql = "SELECT W1.* FROM" _
                    & " SOTPICK1 P1 JOIN" _
                    & " SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) JOIN" _
                    & " ICTWHSE1 W1 ON (O1.WHSE_CODE=W1.WHSE_CODE) WHERE P1.PICK_NO=:PARM1"
        Dim rowICTWHSE1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {PickNo})

        Dim labelData As New Dictionary(Of String, DataRow)
        labelData.Add("SOTCART1", rowSOTSHIPX)
        labelData.Add("SOTPICK1", rowSOTPICK1)
        labelData.Add("SOTSHIPX", rowSOTSHIPX)
        labelData.Add("ICTWHSE1", rowICTWHSE1)
        labelData.Add("SOTORDR5", RowSOTORDR5)
        Return labelData
    End Function

    Public Sub Set1of9(ByVal startPos As Integer, ByVal totalLabels As Integer)
        Me.StartPos = startPos
        Me.TotalLabels = totalLabels
    End Sub

    Protected Overrides Sub ChangeLabelData(labelData As Dictionary(Of String, DataRow), currentIndex As Integer, lastIndex As Integer)
        If labelData("SOTCART1").Table.Columns.Contains("ADDR_1_OF_9") Then
            labelData("SOTCART1").Item("ADDR_1_OF_9") = CStr(currentIndex + If(StartPos > 0, StartPos, 1) - 1) & " of " & CStr(If(TotalLabels > 0, TotalLabels, lastIndex))
        End If
    End Sub

    Protected Overrides Function GetLabelTemplate() As String
        Dim labelTemplate As String = ASCDATA1.GetDataValue(
            "SELECT UCC128_COMMANDS FROM " &
            " SOTUCCL1 U1 " &
            " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {LabelTemplateCode}) & ""
        If labelTemplate = "" Then Throw New Exception("Label Template '" & LabelTemplateCode & "' not found")
        Return labelTemplate
    End Function
End Class

Public Class TestLabel
    Inherits CartonLabel

    Private Property labelTemplateCode As String

    Public Sub New(ByVal labelTemplateCode As String, ByVal cartonNo As String)
        MyBase.New(cartonNo)
        Me.labelTemplateCode = labelTemplateCode
    End Sub

    ''' <summary>
    ''' Use this call to print your own data to the label defined and the printer selected
    ''' </summary>
    Public Sub PrintTestLabel(PrinterNAme As String, labelData As Dictionary(Of String, DataRow))
        Dim labelTemplate = GetLabelTemplate()
        Dim labeltoPrint As String = FillLabelTemplateWithData(labelTemplate, labelData)
        ShippingLabel.SendToLabelPrinter(labeltoPrint, PrinterNAme)

    End Sub
    Public Sub PrintRawZPL(PrinterNAme As String, labelRawData As String)

        ShippingLabel.SendToLabelPrinter(labelRawData, PrinterNAme)

    End Sub
    Protected Overrides Function GetLabelData() As Dictionary(Of String, DataRow)
        Return MyBase.GetLabelData()
    End Function

    Protected Overrides Function GetLabelTemplate() As String
        Dim labelTemplate As String = ASCDATA1.GetDataValue(
            "SELECT UCC128_COMMANDS FROM " &
            " SOTUCCL1 U1 " &
            " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {labelTemplateCode}) & ""
        If labelTemplate = "" Then Throw New Exception("Label Template '" & labelTemplateCode & "' not found")
        Return labelTemplate
    End Function
End Class

Public Class CustomLabel
    Inherits CartonLabel

    Private Property labelTemplateCode As String
    Public tblLabelData As DataTable

    Public Sub New(ByVal labelTemplateCode As String, ByVal cartonNo As String)
        MyBase.New(cartonNo)
        Me.labelTemplateCode = labelTemplateCode

        tblLabelData = New DataTable
        tblLabelData.TableName = "LABELDATA"
        With tblLabelData
            For iLoop As Int16 = 1 To 100
                .Columns.Add("FIELD" & iLoop.ToString.Trim, GetType(System.String))
            Next
        End With
    End Sub

    Protected Overrides Function GetLabelData() As Dictionary(Of String, DataRow)
        Dim labelData As New Dictionary(Of String, DataRow)

        Dim rowLabelData As DataRow = Nothing
        If tblLabelData.Rows.Count > 0 Then
            rowLabelData = tblLabelData.Rows(0)
        Else
            rowLabelData = tblLabelData.NewRow
        End If

        labelData.Add("LABELDATA", rowLabelData)
        Return labelData
    End Function

    Protected Overrides Function GetLabelTemplate() As String
        Dim labelTemplate As String = ASCDATA1.GetDataValue(
            "SELECT UCC128_COMMANDS FROM " &
            " SOTUCCL1 U1 " &
            " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {labelTemplateCode}) & ""
        If labelTemplate = "" Then Throw New Exception("Label Template '" & labelTemplateCode & "' not found")
        Return labelTemplate
    End Function
End Class

Public Class CharmingLabel
    Inherits CartonLabel

    Private Property labelTemplateCode As String
    Public tblLabelData As DataTable
    Public tblShippingData As DataTable
    Public cartonString As String
    Public poNumber As String
    Public division As String

    Public Sub New(ByVal labelTemplateCode As String)
        MyBase.New("")
        Me.labelTemplateCode = labelTemplateCode
    End Sub

    Protected Overrides Function GetLabelData() As Dictionary(Of String, DataRow)
        Dim labelData As New Dictionary(Of String, DataRow)

        Dim dtCharmingData As DataTable = New DataTable()


        dtCharmingData.Columns.Add("CARTONX")
        dtCharmingData.Columns.Add("DIVISION")
        dtCharmingData.Columns.Add("DEPT")
        dtCharmingData.Columns.Add("PO_NO")
        dtCharmingData.Columns.Add("PO_LINES")
        dtCharmingData.Columns.Add("MODE")
        dtCharmingData.Columns.Add("T6DATA")
        dtCharmingData.Columns.Add("CUST_NAME")
        dtCharmingData.Columns.Add("CUST_ADDR1")
        dtCharmingData.Columns.Add("CUST_ADDR2")
        dtCharmingData.Columns.Add("CUST_CITY")
        dtCharmingData.Columns.Add("CUST_STATE")
        dtCharmingData.Columns.Add("CUST_ZIP_CODE")
        dtCharmingData.Columns.Add("CUST_COLOR_CODE")
        dtCharmingData.Columns.Add("CUST_SIZE_CODE")
        dtCharmingData.Columns.Add("INNER_PACK_QTY")
        dtCharmingData.Columns.Add("CARTON_PACK_QTY")
        dtCharmingData.Columns.Add("CASE_WEIGHT_GRS")
        dtCharmingData.Columns.Add("COUNTRY_CODE")
        dtCharmingData.Columns.Add("COUNTRY_NAME")
        dtCharmingData.Columns.Add("STYLE_CODE")
        dtCharmingData.Columns.Add("NET_WEIGHT")
        dtCharmingData.Columns.Add("ORDR_CUST_PO")
        dtCharmingData.Columns.Add("ORDR_NO")
        dtCharmingData.Columns.Add("ORDR_LNO")
        dtCharmingData.Columns.Add("CUST_SKU")
        dtCharmingData.Columns.Add("ORDR_DEPT")
        Dim rowLabelData As DataRow = dtCharmingData.NewRow()

        For Each colName As String In {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE"}
            rowLabelData.Item(colName) = tblShippingData.Rows(0).Item(colName)
        Next

        If Me.division = "CATHERINE" Then
            rowLabelData.Item("DIVISION") = "Catherine's"
        ElseIf Me.division = "LANE BRYANT" Then
            rowLabelData.Item("DIVISION") = "Lane Bryant"
        ElseIf Me.division = "DRESSBARN" Then
            rowLabelData.Item("DIVISION") = "Dress Barn"
        End If

        rowLabelData.Item("PO_NO") = Me.poNumber
        rowLabelData.Item("CARTONX") = cartonString


        If Me.division = "DRESSBARN" Then
            Dim drLoadedData As DataRow = tblLabelData.Rows(0)

            For Each col In {"CUST_SKU", "ORDR_CUST_PO", "CASE_WEIGHT_GRS", "INNER_PACK_QTY", "COUNTRY_NAME",
                            "CARTON_PACK_QTY", "CUST_SIZE_CODE", "CUST_COLOR_CODE", "ORDR_NO", "ORDR_LNO", "STYLE_CODE", "ORDR_DEPT", "NET_WEIGHT"}
                rowLabelData.Item(col) = drLoadedData.Item(col)
            Next
        Else
            Dim T6ROW = "\&" & vbCrLf &
                   "Color: {0}\&" & vbCrLf &
                   "Size: {1} Total: {2}\&"
            Dim T6DATA As String = ""
            For Each row As DataRow In tblLabelData.Rows
                T6DATA &= String.Format(T6ROW, row.Item("EDI_SLN_COLOR"), row.Item("SIZE_STRING"), row.Item("TOTAL_QTY"))
                rowLabelData.Item("MODE") = row.Item("EDI_SLN_LINE_MODE")
                rowLabelData.Item("DEPT") = row.Item("EDI_SLN_DEPT")
            Next

            rowLabelData.Item("T6DATA") = T6DATA

            Dim poLineNos As String = String.Join(",", tblLabelData.AsEnumerable().Select(Function(x) x.Item("EDI_SLN_PO_LNO")).Distinct().ToArray())
            rowLabelData.Item("PO_LINES") = poLineNos
        End If

        labelData.Add("LABELDATA", rowLabelData)
        Return labelData
    End Function

    Protected Overrides Function GetLabelTemplate() As String
        Dim labelTemplate As String = ASCDATA1.GetDataValue(
            "SELECT UCC128_COMMANDS FROM " &
            " SOTUCCL1 U1 " &
            " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {labelTemplateCode}) & ""
        If labelTemplate = "" Then Throw New System.Exception("Label Template '" & labelTemplateCode & "' not found")
        Return labelTemplate
    End Function

End Class