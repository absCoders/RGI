Imports System.IO
Imports System.Net
Imports System.Text

Public Class TACZPLT1

    ' ******************************************
    ' ASCMAIN1.MiniLabelPrinterIPAddress is a 2.25 by 1.25 label Label Printer
    ' Use this value to force a label to this printer
    ' SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)
    '
    ' To print to the 4 by 6 printer use the following
    ' SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)
    ' ******************************************

    Private tblTATZPLT1 As DataTable = Nothing
    Private rowTATZPLT1 As DataRow = Nothing

    Public Enum LabelSizes
        label4x6
        label225x125
    End Enum

    Public ErrorMessge As String = String.Empty

#Region "Instantiate Class"

    Public Sub New(ByVal DisplayLabel As Boolean)
        InitializeClass(DisplayLabel)
    End Sub

    Public Sub New()
        InitializeClass(False)
    End Sub

    Private Sub InitializeClass(ByVal DisplayLabel As Boolean)
        Me.DisplayLabel = DisplayLabel
        tblTATZPLT1 = ASCDATA1.GetDataTable("SELECT * FROM TATZPLT1", "TATZPLT1")
    End Sub

    Public Property DisplayLabel() As Boolean = False

#End Region

#Region "Public Procedures"

    Public Sub PrintNPILabel(ByVal CUST_ORDER_NO As String,
                             ByVal CUST_SHIP_TO_NO As String,
                             ByVal DELIVERY_DATE As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("NPI_LABEL")
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{CUST_ORDER_NO}", CUST_ORDER_NO)
        labelImage = labelImage.Replace("{CUST_SHIP_TO_NO}", CUST_SHIP_TO_NO)
        labelImage = labelImage.Replace("{DELIVERY_DATE}", DELIVERY_DATE)

        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintStockLensTransferLabel(ByVal INV_NO As String,
                                           ByVal ORDR_NO As String,
                                           ByVal ORDR_CUST_PO As String,
                                           ByVal CUST_CODE As String,
                                           ByVal CUST_SHIP_TO_NO As String,
                                           ByVal SHIP_TO_NAME As String,
                                           ByVal ITEM1 As String,
                                           ByVal ITEM1_DESC As String,
                                           ByVal ITEM2 As String,
                                           ByVal ITEM2_DESC As String,
                                           ByVal BIN_NO As String,
                                           ByVal INV_DATE As String,
                                           ByVal TRUCK_NO As String,
                                           ByVal TOTE_NO As String,
                                           ByVal SLOT_NO As String,
                                           ByVal BIN_NO_R2 As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("OPTRANSFER")
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{INV_NO}", INV_NO)
        labelImage = labelImage.Replace("{ORDR_NO}", ORDR_NO)
        labelImage = labelImage.Replace("{CUSTCODE}", CUST_CODE)
        labelImage = labelImage.Replace("{SHIPTO}", CUST_SHIP_TO_NO)
        labelImage = labelImage.Replace("{SHIPTONAME}", SHIP_TO_NAME)
        labelImage = labelImage.Replace("{ORDR_CUST_PO}", ORDR_CUST_PO)
        labelImage = labelImage.Replace("{ITEM1}", ITEM1)
        labelImage = labelImage.Replace("{ITEM1_DESC}", ITEM1_DESC)
        labelImage = labelImage.Replace("{ITEM2}", ITEM2)
        labelImage = labelImage.Replace("{ITEM2_DESC}", ITEM2_DESC)
        labelImage = labelImage.Replace("{INV_DATE}", INV_DATE)
        labelImage = labelImage.Replace("{TRUCK_NO}", TRUCK_NO)
        labelImage = labelImage.Replace("{TOTE_NO}", TOTE_NO)
        labelImage = labelImage.Replace("{SLOT_NO}", SLOT_NO)

        ' INC0167996 PIT Insource Job Flow Change
        If BIN_NO_R2.Length > 0 Then
            labelImage = labelImage.Replace("{BIN_NO_R2}", "Frm Bin: " & BIN_NO_R2)
        Else
            labelImage = labelImage.Replace("{BIN_NO_R2}", String.Empty)
        End If

        If BIN_NO.Length > 0 Then
            labelImage = labelImage.Replace("{BIN_NO}", "Bin No:")
            labelImage = labelImage.Replace("{BIN_NO_BC}", BIN_NO)
        Else
            'labelImage = labelImage.Replace("{BIN_NO}", "")
            'labelImage = labelImage.Replace("{BIN_NO_BC}", "")

            ' Need to delete the commnads to avoid a useless barcode on the label
            Dim lstLabelImage As New List(Of String)
            Dim newLabelImage As String = String.Empty

            lstLabelImage = labelImage.Split(vbLf).ToList
            For Each command As String In lstLabelImage
                If command.Contains("{BIN_NO}") Then
                    ' Do not use the command
                ElseIf command.Contains("{BIN_NO_BC}") Then
                    ' Do not use the command
                Else
                    newLabelImage &= command
                End If
            Next
            labelImage = newLabelImage
        End If

        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)
    End Sub

    Public Sub PrintVendorRALabel(ByVal RA_NO As String, ByVal VEND_RA_NO As String, ByVal PARTNER_NAME As String)
        rowTATZPLT1 = tblTATZPLT1.Rows.Find("VENDOR_RA_LABEL")
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{RA_NO}", RA_NO)
        labelImage = labelImage.Replace("{VEND_RA_NO}", VEND_RA_NO)
        labelImage = labelImage.Replace("{PARTNER_NAME}", PARTNER_NAME)

        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintCaseReceiptLabel(ByVal ITEM_CODE As String,
                                     ByVal ITEM_DESC As String,
                                     ByVal QTY_REC As Int32,
                                     ByVal LOC_CODE As String)


        rowTATZPLT1 = tblTATZPLT1.Rows.Find("CASE_RECEIPT")
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{ITEM_CODE}", ITEM_CODE)
        labelImage = labelImage.Replace("{ITEM_DESC}", ITEM_DESC)
        labelImage = labelImage.Replace("{QTY_REC}", QTY_REC)
        labelImage = labelImage.Replace("{LOC_CODE}", LOC_CODE)

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)
    End Sub

    Public Sub PrintInvoiceDetailsLabel(ByVal INV_NO As String,
                                        ByVal ITEM_DESC As String,
                                        ByVal FRAME_DESC As String,
                                        ByVal CUST_LINE_REF As String,
                                        ByVal ORDR_NO As String,
                                        ByVal TOTE_NO As String,
                                        ByVal LAB_CODE As String,
                                        ByVal LABEL_TYPE As String,
                                        ByVal BOX_NO As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("INV_DETAILS")
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{INV_NO}", INV_NO)
        labelImage = labelImage.Replace("{ITEM_DESC}", ITEM_DESC)
        labelImage = labelImage.Replace("{FRAME_DESC}", FRAME_DESC)
        labelImage = labelImage.Replace("{CUST_LINE_REF}", CUST_LINE_REF)
        labelImage = labelImage.Replace("{ORDR_NO}", ORDR_NO)
        labelImage = labelImage.Replace("{TOTE_NO}", TOTE_NO)
        labelImage = labelImage.Replace("{LAB_CODE}", LAB_CODE)
        labelImage = labelImage.Replace("{LABEL_TYPE}", LABEL_TYPE)
        labelImage = labelImage.Replace("{BOX_NO}", BOX_NO)

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintEmployeeLabel(ByVal WH_OPER_ID As String,
                                  ByVal WH_OPER_NAME As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("EMP_LABEL")
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{WH_OPER_ID}", WH_OPER_ID)
        labelImage = labelImage.Replace("{WH_OPER_NAME}", WH_OPER_NAME)

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintPlacard(ByVal PALLET_NO As String,
                            ByVal INIT_OPER As String,
                            ByVal INIT_DATE As Date,
                            ByVal PALLET_TYPE As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("PALLET")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{PALLET_NO}", PALLET_NO)
        labelImage = labelImage.Replace("{INIT_OPER}", INIT_OPER)
        labelImage = labelImage.Replace("{INIT_DATE}", INIT_DATE)
        labelImage = labelImage.Replace("{PALLET_TYPE}", PALLET_TYPE)

        ' Prints to a 4 by 6 label
        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintDCCartonLabel(ByVal CARTON_NO As String,
                            ByVal DC_CODE_FROM As String,
                            ByVal DC_CODE_TO As String,
                            ByVal DC_TRANS_NO As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("DC_CARTON")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{CARTON_NO}", CARTON_NO)
        labelImage = labelImage.Replace("{DC_CODE_FROM}", DC_CODE_FROM)
        labelImage = labelImage.Replace("{DC_CODE_TO}", DC_CODE_TO)
        labelImage = labelImage.Replace("{DC_TRANS_NO}", DC_TRANS_NO)

        ' Prints to a 4 by 6 label
        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)

    End Sub


    Public Sub PrintTruckCourierShippingLabel(ByVal TRACKING_NO As String,
                            ByVal DC_CODE_FROM As String,
                            ByVal DC_CODE_TO As String,
                            ByVal DC_TRANS_NO As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("DC_COURIER_LABEL")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{TRACKING_NO}", TRACKING_NO)
        labelImage = labelImage.Replace("{DC_CODE_FROM}", DC_CODE_FROM)
        labelImage = labelImage.Replace("{DC_CODE_TO}", DC_CODE_TO)
        labelImage = labelImage.Replace("{DC_TRANS_NO}", DC_TRANS_NO)

        ' Prints to a 4 by 6 label
        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub Print_Truck_ID(ByVal TRUCK_NO As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("SOTTRCK1_TRUCK_ID")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{TRUCK_NO}", TRUCK_NO)

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintShippingBoxLabel(ByVal BOX_NO As String,
                                     ByVal CUST_CODE As String,
                                     ByVal CUST_SHIP_TO_NO As String,
                                     ByVal CUST_SHIP_TO_NAME As String,
                                     ByVal SHIP_VIA_CODE As String,
                                     ByVal SHIP_VIA_DESC As String,
                                     ByVal PARTNER_CODE As String,
                                     ByVal PARTNER_NAME As String,
                                     ByVal LAB_CODE As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("BOX_LABEL")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{BOX_NO}", BOX_NO)
        labelImage = labelImage.Replace("{CUST_CODE}", CUST_CODE)
        labelImage = labelImage.Replace("{CUST_SHIP_TO_NO}", CUST_SHIP_TO_NO)
        labelImage = labelImage.Replace("{CUST_SHIP_TO_NAME}", CUST_SHIP_TO_NAME)
        labelImage = labelImage.Replace("{SHIP_VIA_CODE}", SHIP_VIA_CODE)
        labelImage = labelImage.Replace("{SHIP_VIA_DESC}", SHIP_VIA_DESC)
        labelImage = labelImage.Replace("{PARTNER_CODE}", PARTNER_CODE)
        labelImage = labelImage.Replace("{PARTNER_NAME}", PARTNER_NAME)
        labelImage = labelImage.Replace("{LAB_CODE}", LAB_CODE)

        ' Prints to a 4 by 6 label
        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintFrameLabel(ByVal FRAME_UPC_CODE As String, ByVal FRAME_DESC As String, Optional labelQty As Integer = 1)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("FRAME_LABEL")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{FRAME_UPC_CODE}", FRAME_UPC_CODE)
        labelImage = labelImage.Replace("{FRAME_DESC}", FRAME_DESC)

        For i As Integer = 1 To labelQty
            SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)
        Next

    End Sub

    Public Sub PrintItemLabel(ByVal ITEM_CODE As String, Optional labelQty As Integer = 1)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("ITEM_LABEL")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{ITEM_CODE}", ITEM_CODE)

        For i As Integer = 1 To labelQty
            SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)
        Next

    End Sub

    Public Sub Print_Tote_Labels(rows() As DataRow, ByVal OneTotePerLabel As Boolean)

        If OneTotePerLabel Then
            Print_Tote_Labels_OneTotePerLabel(rows)
            Exit Sub
        End If

        Dim ZPL_KEY As String = "SOTTOTE1"
        If rows(0).Item("TOTE_TYPE") & "" = "P" Then
            ZPL_KEY = "SOTTOTE1_PRE"
        End If
        rowTATZPLT1 = tblTATZPLT1.Rows.Find(ZPL_KEY)
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim zplx_orig As String = rowTATZPLT1.Item("ZPL_BODY") & ""
        Dim zplx As String = rowTATZPLT1.Item("ZPL_BODY") & ""

        Dim zpl As String = ""

        Dim label_count As Integer = 0
        Dim zplx_save As String = ""
        For Each rowSOTTOTE1 As DataRow In rows

            Dim TOTE_TYPE As String = rowSOTTOTE1.Item("TOTE_TYPE") & ""
            Dim TOTE_CLASS_CODE As String = rowSOTTOTE1.Item("TOTE_CLASS_CODE") & ""

            label_count += 1

            Dim TOTE_NO As String = rowSOTTOTE1.Item("TOTE_NO") & ""

            If label_count Mod 2 = 1 Then
                zplx = Replace(zplx, "{TOTE_NO1}", TOTE_NO)
                If TOTE_TYPE = "P" Then
                    zplx = Replace(zplx, "{TOTE_NO_SFX1}", Mid(TOTE_NO, 5, 2))
                Else
                    zplx = Replace(zplx, "{TOTE_CLASS_CODE1}", TOTE_CLASS_CODE)
                End If
                zplx_save = zplx
            Else
                zplx = Replace(zplx, "{TOTE_NO2}", TOTE_NO)
                If TOTE_TYPE = "P" Then
                    zplx = Replace(zplx, "{TOTE_NO_SFX2}", Mid(TOTE_NO, 5, 2))
                Else
                    zplx = Replace(zplx, "{TOTE_CLASS_CODE2}", TOTE_CLASS_CODE)
                End If
                zpl &= zplx
                zplx_save = ""
                zplx = zplx_orig
            End If
            'zplx = Replace(zplx, "{TOTE_NO}", TOTE_NO)

            'zpl &= zplx
        Next

        If zplx_save <> "" Then
            zpl &= zplx_save
        End If

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, zpl)
    End Sub

    Public Sub Print_Tote_Labels_OneTotePerLabel(rows() As DataRow)

        Dim ZPL_KEY As String = "SOTTOTE1"
        If rows(0).Item("TOTE_TYPE") & "" = "P" Then
            ZPL_KEY = "SOTTOTE1_PRE"
        End If
        rowTATZPLT1 = tblTATZPLT1.Rows.Find(ZPL_KEY)
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim zplx_orig As String = rowTATZPLT1.Item("ZPL_BODY") & ""
        Dim zplx As String = rowTATZPLT1.Item("ZPL_BODY") & ""

        Dim zpl As String = ""

        Dim label_count As Integer = 0
        Dim zplx_save As String = ""
        For Each rowSOTTOTE1 As DataRow In rows

            Dim TOTE_TYPE As String = rowSOTTOTE1.Item("TOTE_TYPE") & ""
            Dim TOTE_CLASS_CODE As String = rowSOTTOTE1.Item("TOTE_CLASS_CODE") & ""

            label_count += 1

            Dim TOTE_NO As String = rowSOTTOTE1.Item("TOTE_NO") & ""
            zplx = Replace(zplx, "{TOTE_NO1}", TOTE_NO)
            If TOTE_TYPE = "P" Then
                zplx = Replace(zplx, "{TOTE_NO_SFX1}", Mid(TOTE_NO, 5, 2))
            Else
                zplx = Replace(zplx, "{TOTE_CLASS_CODE1}", TOTE_CLASS_CODE)
            End If
            zplx_save = zplx

            zplx = Replace(zplx, "{TOTE_NO2}", TOTE_NO)
            If TOTE_TYPE = "P" Then
                zplx = Replace(zplx, "{TOTE_NO_SFX2}", Mid(TOTE_NO, 5, 2))
            Else
                zplx = Replace(zplx, "{TOTE_CLASS_CODE2}", TOTE_CLASS_CODE)
            End If

            zpl &= zplx
            zplx_save = ""
            zplx = zplx_orig
        Next

        If zplx_save <> "" Then
            zpl &= zplx_save
        End If

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, zpl)

    End Sub

    Public Sub Print_Custom_Tote_Labels(LABEL_ACTION As String, TRUCK_NO As String, Optional TOTE_NO_only As String = "",
                                        Optional KEY1 As String = "", Optional LNO1 As Int32 = 0,
                                        Optional KEY2 As String = "", Optional LNO2 As Int32 = 0)

        'https://www.zebra.com/us/en/support-downloads/knowledge-articles/how-to-preview-a-zpl-label-with-a-network-printer.html
        'http://labelary.com/viewer.html

        ' ASCMAIN1.sql = $"Select * from SOTTOTE1 where TRUCK_NO = '{TRUCK_NO}'"
        ASCMAIN1.sql = "SELECT SOTTOTE1.*, SOTPICK1.ORDR_NO, SOTPICK0.PARTNER_CODE" & vbCrLf _
            & ", SOTPICK0.LAB_CODE, SOTPICK0.NPIX_NO, POTNPIX1.NPIX_TYPE" & vbCrLf _
            & ", SOTPICK0.FRAME_SOURCE_TYPE, SOTPICK1.CUST_STORE_NO" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_SHIP_TO_NAME, SOTSVIA1.SHIP_VIA_DESC" & vbCrLf _
            & ", SOTORDR1.ORDR_SHIP_COMPLETE" & vbCrLf _
            & " FROM SOTTOTE1,SOTPICK1,SOTPICK0,SOTORDR1,POTNPIX1,SOTSVIA1" & vbCrLf _
            & "WHERE SOTPICK1.PICK_NO = SOTTOTE1.PICK_NO" & vbCrLf _
            & "And SOTPICK0.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO" & vbCrLf _
            & "And POTNPIX1.NPIX_NO (+) = SOTPICK0.NPIX_NO" & vbCrLf _
            & "And SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "And SOTSVIA1.SHIP_VIA_CODE (+) = SOTPICK1.SHIP_VIA_CODE" & vbCrLf _
            & $"And SOTTOTE1.TRUCK_NO = '{TRUCK_NO}'"

        If TOTE_NO_only <> "" Then
            ASCMAIN1.sql &= $" and SOTTOTE1.TOTE_NO = '{TOTE_NO_only}'"
        End If

        ' NEED TO KNOW IF SHIP COMPLETE

        Dim zpl As String = ""

        Dim ZPL_KEY As String = ""

        For Each rowSOTTOTE1 As DataRow In ASCDATA1.GetDataTable.Select("", "TOTE_NO")
            Dim TOTE_NO As String = rowSOTTOTE1.Item("TOTE_NO") & ""
            Dim SLOT_NO As Int32 = Val(rowSOTTOTE1.Item("SLOT_NO") & "")
            Dim ORDR_NO As String = rowSOTTOTE1.Item("ORDR_NO") & ""
            Dim PICK_NO As String = rowSOTTOTE1.Item("PICK_NO") & ""
            Dim ORDR_CUST_PO As String = rowSOTTOTE1.Item("ORDR_CUST_PO") & ""
            Dim CUST_SHIP_TO_NAME As String = rowSOTTOTE1.Item("CUST_SHIP_TO_NAME") & ""
            Dim SHIP_VIA_DESC As String = rowSOTTOTE1.Item("SHIP_VIA_DESC") & ""
            Dim CUST_STORE_NO As String = rowSOTTOTE1.Item("CUST_STORE_NO") & ""
            Dim PARTNER_CODE As String = rowSOTTOTE1.Item("PARTNER_CODE") & ""
            Dim NPIX_NO As String = rowSOTTOTE1.Item("NPIX_NO") & ""
            Dim LAB_CODE As String = rowSOTTOTE1.Item("LAB_CODE") & ""
            Dim ORDR_SHIP_COMPLETE As String = rowSOTTOTE1.Item("ORDR_SHIP_COMPLETE") & ""

            Dim DC_CODE As String = rowSOTTOTE1.Item("DC_CODE") & ""

            ASCMAIN1.sql = $"Select Sum (PICK_QTY) PICK_QTY from SOTPICK2 where PICK_NO = '{PICK_NO}'"
            Dim PICK_QTY As Integer = Val(ASCDATA1.GetDataValue & "")

            Dim FRAME_SOURCE_TYPE As String = rowSOTTOTE1.Item("FRAME_SOURCE_TYPE") & ""
            Dim FRAME_SOURCE_TYPE_DESC As String = ""
            If FRAME_SOURCE_TYPE = "J" Then FRAME_SOURCE_TYPE_DESC = "Jobs (Rx)"
            If FRAME_SOURCE_TYPE = "R" Then FRAME_SOURCE_TYPE_DESC = "Rx Replen"
            If FRAME_SOURCE_TYPE = "U" Then FRAME_SOURCE_TYPE_DESC = "PO Replen"

            Dim NPIX_TYPE As String = rowSOTTOTE1.Item("NPIX_TYPE") & ""
            Dim NPIX_TYPE_DESC As String = ""
            If NPIX_TYPE = "N" Then NPIX_TYPE_DESC = "NPI Order"
            If NPIX_TYPE = "K" Then NPIX_TYPE_DESC = "Store Kit"

            ZPL_KEY = "SOTTOTE1"
            If LAB_CODE <> "" Then ZPL_KEY = "SOTTOTE1_LAB"
            If NPIX_NO <> "" Then ZPL_KEY = "SOTTOTE1_NPI"
            If DC_CODE = "NY" Then ZPL_KEY = "SOTTOTE1_DC"
            Dim rowTATZPLT1 As DataRow = tblTATZPLT1.Rows.Find(ZPL_KEY)

            Dim zplx As String = rowTATZPLT1.Item("ZPL_BODY") & ""

            'zpl &= "^XA" & vbCrLf
            'zpl &= "^CF0,30" & vbCrLf
            'zpl &= $"^FO50,50^FDTote No^FS^FO200,50^FD{TOTE_NO}^FS" & vbCrLf
            'zpl &= $"^FO50,100^A0N,30,30^BCN,50,N,N,N ^FD{TOTE_NO}^FS" & vbCrLf
            'zpl &= $"^FO50,250^FDOrder No^FS^FO200,250^FD{ORDR_NO}^FS" & vbCrLf
            'zpl &= $"^FO50,300^FDPartner^FS^FO200,300^FD{PARTNER_CODE}^FS" & vbCrLf
            'If LAB_CODE <> "" Then zpl &= $"^FO50,350^FDLab^FS^FO200,350^FD{LAB_CODE}^FS^FO300,350^FD{FRAME_SOURCE_TYPE_DESC}^FS" & vbCrLf
            'If NPIX_NO <> "" Then zpl &= $"^FO50,350^FDNPI#^FS^FO200,350^FD{NPIX_NO}^FS^FO300,350^FD{CUST_STORE_NO}^FS" & vbCrLf
            'zpl &= $"^FO50,400^FDTruck^FS^FO200,400^FD{TRUCK_NO}^FS" & vbCrLf
            'zpl &= $"^FO50,450^FDSlot^FS^FO200,450^FD{Format(SLOT_NO, "00")}^FS" & vbCrLf
            'zpl &= "^XZ" & vbCrLf

            zplx = Replace(zplx, "{TOTE_NO}", TOTE_NO)
            zplx = Replace(zplx, "{ORDR_NO}", ORDR_NO)
            zplx = Replace(zplx, "{PARTNER_CODE}", PARTNER_CODE)
            zplx = Replace(zplx, "{CUST_STORE_NO}", CUST_STORE_NO)

            If LAB_CODE <> "" Then
                zplx = Replace(zplx, "{LAB_CODE}", LAB_CODE)
                zplx = Replace(zplx, "{FRAME_SOURCE_TYPE_DESC}", FRAME_SOURCE_TYPE_DESC)
            ElseIf NPIX_NO <> "" Then
                zplx = Replace(zplx, "{NPIX_NO}", NPIX_NO)
                zplx = Replace(zplx, "{NPIX_TYPE_DESC}", NPIX_TYPE_DESC)
            ElseIf DC_CODE = "NY" Then
                ' PLACE CUSTOM REPLACEMENTS HERE
            Else

            End If

            zplx = Replace(zplx, "{TRUCK_NO}", TRUCK_NO)
            zplx = Replace(zplx, "{SLOT_NO}", Format(SLOT_NO, "00"))

            Dim SHIP_COMPLETE As String = ""
            If ORDR_SHIP_COMPLETE = "1" Then
                SHIP_COMPLETE = "*** SHIP COMPLETE ***"
            End If
            zplx = Replace(zplx, "{SHIP_COMPLETE}", SHIP_COMPLETE)

            zplx = Replace(zplx, "{ORDR_CUST_PO}", ORDR_CUST_PO)
            zplx = Replace(zplx, "{SHIP_VIA_DESC}", SHIP_VIA_DESC)
            zplx = Replace(zplx, "{CUST_SHIP_TO_NAME}", CUST_SHIP_TO_NAME)
            zplx = Replace(zplx, "{PICK_QTY}", CStr(PICK_QTY))
            zplx = Replace(zplx, "{DT}", Format(Now, "MM/dd/yy HH:mm:ss"))

            zpl &= zplx
        Next

        'SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, zpl)
        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, zpl, LABEL_ACTION, ZPL_KEY, KEY1, LNO1, KEY2, LNO2)
    End Sub

    Public Sub Print_Truck_Pick_Tag(LABEL_ACTION As String, ByVal TRUCK_NO As String, PICK_DESCRIPTION As String, ORDER_COUNT As Integer,
                                        Optional KEY1 As String = "", Optional LNO1 As Int32 = 0,
                                        Optional KEY2 As String = "", Optional LNO2 As Int32 = 0)

        Dim ZPL_KEY As String = "SOTTRCK1_PICK_TAG"
        rowTATZPLT1 = tblTATZPLT1.Rows.Find(ZPL_KEY)
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & ""
        labelImage = labelImage.Replace("{TRUCK_NO}", TRUCK_NO)
        labelImage = labelImage.Replace("{PICK_DESCRIPTION}", PICK_DESCRIPTION)
        labelImage = labelImage.Replace("{ORDER_COUNT}", CStr(ORDER_COUNT))
        labelImage = labelImage.Replace("{DATE_TIME}", Format(Now, "MM/dd/yy HH:mm"))

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage, LABEL_ACTION, ZPL_KEY, KEY1, LNO1, KEY2, LNO2)
    End Sub

    Public Sub Print_UPC_Label(
                  STORE As String,
                  Optional UPC As String = "",
                  Optional STYLE As String = "",
                  Optional DESCRIPTION As String = "",
                  Optional FRAME_SIZE As String = "",
                  Optional COLLECTION As String = "",
                  Optional RETAIL_PRICE As Decimal = 0,
                  Optional NPIX_REFERENCE As String = "",
                  Optional PARTNER_CODE As String = "",
                  Optional LABELS As Integer = 0,
                  Optional BIN_NO As String = "")

        Dim LABEL As String = ""

        Dim VERT As String = "250"
        Dim HORZ As Integer = "650"
        Dim HORZPLUS As Integer = 22

        If NPIX_REFERENCE <> "" Then ' Single Labels per Store - Header Label

            LABEL = Chr(27) & "A" ' start

            LABEL &= Chr(27) & "%2" ' rotate 180 degrees

            LABEL &= Chr(27) & "V" & "220"
            LABEL &= Chr(27) & "H" & "650"
            LABEL &= Chr(27) & "FT" & ",50,8,25,0"

            LABEL &= Chr(27) & "%1" ' rotate 90 degrees

            LABEL &= Chr(27) & "V" & "250"
            LABEL &= Chr(27) & "H" & "700"
            LABEL &= Chr(27) & "XM" & PARTNER_CODE

            LABEL &= Chr(27) & "V" & "100"
            LABEL &= Chr(27) & "H" & "700"
            LABEL &= Chr(27) & "S" & CStr(LABELS) & " Labels"

            LABEL &= Chr(27) & "%1" ' rotate 90 degrees

            LABEL &= Chr(27) & "V" & "250"
            LABEL &= Chr(27) & "H" & "800"
            LABEL &= Chr(27) & "P" & "2"
            LABEL &= Chr(27) & "L" & "0101"
            LABEL &= Chr(27) & "XM" & NPIX_REFERENCE

            LABEL &= Chr(27) & "Q" & "1" ' LABEL QTY
            LABEL &= Chr(27) & "Z"

        ElseIf UPC = "" Then ' STORE HEADER LABEL

            LABEL = Chr(27) & "A" ' start

            LABEL &= Chr(27) & "%2" ' rotate 180 degrees

            LABEL &= Chr(27) & "V" & "220"
            LABEL &= Chr(27) & "H" & "650"
            LABEL &= Chr(27) & "FT" & ",50,8,25,0"

            LABEL &= Chr(27) & "%1" ' rotate 90 degrees

            LABEL &= Chr(27) & "V" & "250"
            LABEL &= Chr(27) & "H" & "700"
            LABEL &= Chr(27) & "P" & "2"
            LABEL &= Chr(27) & "L" & "0304"
            LABEL &= Chr(27) & "XM" & STORE

            ' INC0147246 - LAB AEG Replen labels - print bin on lead frame tag
            If STORE.Length > 0 Then
                If BIN_NO.Length > 0 Then
                    LABEL &= Chr(27) & "V" & "200"
                    LABEL &= Chr(27) & "H" & "800"
                    LABEL &= Chr(27) & "P" & "2"
                    LABEL &= Chr(27) & "L" & "0101"
                    LABEL &= Chr(27) & $"XM" & $"Bin: {BIN_NO}"
                End If
            End If

            LABEL &= Chr(27) & "Q" & "1" ' LABEL QTY
            LABEL &= Chr(27) & "Z"

        Else

            LABEL = Chr(27) & "A" ' start

            LABEL &= Chr(27) & "V130" & Chr(27) & "H500"
            LABEL &= Chr(27) & "XM" & STORE

            LABEL &= Chr(27) & "%1" ' rotate 90 degrees

            LABEL &= Chr(27) & "V" & VERT & Chr(27) & "H" & CStr(HORZ + 0)
            LABEL &= Chr(27) & "XS" & STYLE ' U is really tiny

            LABEL &= Chr(27) & "V" & VERT & Chr(27) & "H" & CStr(HORZ + 1 * HORZPLUS)
            LABEL &= Chr(27) & "S" & DESCRIPTION

            LABEL &= Chr(27) & "V" & VERT & Chr(27) & "H" & CStr(HORZ + 2 * HORZPLUS)
            LABEL &= Chr(27) & "S" & FRAME_SIZE

            LABEL &= Chr(27) & "V" & "100" & Chr(27) & "H" & CStr(HORZ + 2 * HORZPLUS)
            LABEL &= Chr(27) & "XS" & Format(RETAIL_PRICE, "$##0.00")

            LABEL &= Chr(27) & "V" & VERT & Chr(27) & "H" & CStr(HORZ + 3 * HORZPLUS)
            LABEL &= Chr(27) & "XS" & COLLECTION

            LABEL &= Chr(27) & "%3"

            If Len(UPC) = 12 Then
                LABEL &= Chr(27) & "V" & "040" & Chr(27) & "H" & CStr(HORZ + 8 * HORZPLUS)
                LABEL &= Chr(27) & "BG" & "02" & "050" & ">I" & UPC
            ElseIf Len(UPC) = 13 Then
                LABEL &= Chr(27) & "V" & "015" & Chr(27) & "H" & CStr(HORZ + 8 * HORZPLUS)
                LABEL &= Chr(27) & "BG" & "02" & "050" & ">H" & Mid(UPC, 1, 1) & ">C" & Mid(UPC, 2)
            End If
            'BG06211>I0123456789

            LABEL &= Chr(27) & "%3"
            LABEL &= Chr(27) & "V" & "050" & Chr(27) & "H" & CStr(HORZ + 5 * HORZPLUS)
            LABEL &= Chr(27) & "P" & "08"
            LABEL &= Chr(27) & "XS" & UPC

            LABEL &= Chr(27) & "Q" & "1" ' LABEL QTY
            LABEL &= Chr(27) & "Z"
        End If

        SendLabelToPrinter(ASCMAIN1.UPCFramePrinterIPAddress, LABEL)

        'Try
        '    Using ipp As New nsoftware.IPWorks.Ipport
        '        ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")
        '        ipp.Connect(ASCMAIN1.UPCFramePrinterIPAddress, 9100)
        '        Dim array() As Byte = System.Text.Encoding.ASCII.GetBytes(LABEL)
        '        ipp.Send(array)
        '        ipp.Disconnect()
        '    End Using
        'Catch ex As Exception
        '    MessageBox.Show($"Error Printing Label: {ex.Message}", "Print Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
        'End Try

    End Sub

    Public Sub Print_VAN_UPC_Label(STYLE_CODE As String, COLOR_CODE As String, SIZE_CODE As String, UPC_CODE As String)
        Dim ZPL_KEY As String = "VAN_UPC"
        rowTATZPLT1 = tblTATZPLT1.Rows.Find(ZPL_KEY)
        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & ""
        labelImage = labelImage.Replace("{STYLE_CODE}", STYLE_CODE)
        labelImage = labelImage.Replace("{COLOR_CODE}", COLOR_CODE)
        labelImage = labelImage.Replace("{SIZE_CODE}", SIZE_CODE)
        labelImage = labelImage.Replace("{UPC_CODE}", UPC_CODE)

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)
    End Sub

    Public Sub Print_Location_Label(LOC_CODE As String, LOC_TYPE As String)

        Dim LOC_CODE_FMT As String = LOC_CODE
        If LOC_TYPE = "S" Or LOC_TYPE = "O" Then
            LOC_CODE_FMT = Mid(LOC_CODE, 1, 2) & "-" & Mid(LOC_CODE, 3, 2) & "-" & Mid(LOC_CODE, 5, 2) & "-" & Mid(LOC_CODE, 7, 2)
            If LOC_TYPE = "S" Then
                LOC_CODE_FMT &= "-" & Mid(LOC_CODE, 9, 1)
            End If
        End If

        ASCMAIN1.Progress("-", LOC_CODE_FMT)

        'Dim rowTATZPLT1 As DataRow = Lookup("TATZPLT1", "ICTILOC1_LBL")
        Dim rowTATZPLT1 As DataRow = tblTATZPLT1.Rows.Find("ICTILOC1_LBL")
        Dim ZPL_BODY As String = rowTATZPLT1.Item("ZPL_BODY") & ""

        Dim zpl As String = ZPL_BODY

        zpl = Replace(zpl, "{LOC_CODE_FMT}", LOC_CODE_FMT)
        zpl = Replace(zpl, "{LOC_CODE}", LOC_CODE)


        'Zebra.SendStringToIPLabelPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, zpl)
        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, zpl)

    End Sub

    Public Sub PrintSampleShippingLabel()

        Dim labelImage As String = "^XA
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

        SendLabelToPrinter(ASCMAIN1.LabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub PrintSampleMiniLabel()

        PrintFrameLabel("123456789012", "Text Label")

    End Sub

    Public Sub SendLabelToPrinter(ByVal labelPrinterIP As String, ByVal labelImage As String,
                                  Optional LABEL_ACTION As String = "", Optional ZPL_KEY As String = "",
                                  Optional KEY1 As String = "", Optional LNO1 As Int32 = 0,
                                  Optional KEY2 As String = "", Optional LNO2 As Int32 = 0)

        ErrorMessge = String.Empty

        If (ASCMAIN1.Running_in_VS AndAlso (ASCMAIN1.USER_ID = "edz" Or ASCMAIN1.USER_ID = "wjz")) Then
            If LABEL_ACTION <> "" And ZPL_KEY <> "" Then
                ASCMAIN1.sql = $"Insert into TATZPLH1 (ZPL_CTL_NO,PRINTER_IP,ZPL_KEY,ZPL_BODY,INIT_DATE,INIT_OPER,MENU_ITEM_OBJECT,SESSION_NO,SELECTION_NO,RE_XNO,KEY1,LNO1,KEY2,LNO2, LABEL_ACTION)" & vbCrLf _
                    & $" Values (TAPCTLN1('TATZPLH1.ZPL_CTL_NO',1), '{labelPrinterIP}','{ZPL_KEY}','{labelImage}',SYSDATE,'{ASCMAIN1.USER_ID}','{ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT}','{ASCMAIN1.SESSION_NO}',{CStr(ASCMAIN1.ActiveForm.SELECTION_NO)},{CStr(ASCMAIN1.ActiveForm.RE_XNO)},'{KEY1}',{CStr(LNO1)},'{KEY2}',{CStr(LNO2)},'{LABEL_ACTION}') "
                ASCDATA1.ExecuteSQL()
            End If

        End If


        If labelPrinterIP <> ASCMAIN1.UPCFramePrinterIPAddress Then
            If ASCMAIN1.Running_in_VS Then
                If MessageBox.Show("Do you want to display the label?", "Display Label", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    ShowLabelDialog(labelPrinterIP, labelImage)
                    If MessageBox.Show("Do you also want to print the label?", "Send Label To Printer", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        Try
            labelPrinterIP = (labelPrinterIP & String.Empty).ToString.Trim

            If Not (Net.IPAddress.TryParse(labelPrinterIP, Nothing) AndAlso labelPrinterIP.Split(".").Length = 4) Then
                ErrorMessge = $"Could not determine the IP Address of the printer. Label Printer IP = {labelPrinterIP}"
                Exit Sub
            End If

            labelImage = (labelImage & String.Empty).ToString.Trim
            If labelImage.Length = 0 Then
                Exit Sub
            End If

            Dim port As Int16 = 9100

            Using ipp As New nsoftware.IPWorks.TCPClient ' .Ipport
                ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")

                ' ipp.Connect(labelPrinterIP, port)

                Dim array() As Byte = System.Text.Encoding.ASCII.GetBytes(labelImage)
                ipp.Send(array)

                ipp.Disconnect()
                ipp.Dispose()
            End Using

        Catch ex As Exception
            ErrorMessge = ex.Message
            If ex.InnerException IsNot Nothing Then
                If ex.InnerException.Message & String.Empty <> String.Empty Then
                    ErrorMessge &= " - Inner Exception: " & ex.InnerException.Message
                End If
            End If
        End Try
    End Sub

    Public Sub ShowLabelDialog(ByVal labelSize As LabelSizes, ByVal LabelImage As String)

        Try
            ' "http://api.labelary.com/v1/printers/8dpmm/labels/4x6/0/"
            Dim requestUriString As String = "http://api.labelary.com/v1/printers/8dpmm/labels/"
            Dim fn As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("NOTEPAD") & ".png"

            Select Case labelSize
                Case LabelSizes.label225x125
                    requestUriString &= "2.25x1.25/0/"

                Case LabelSizes.label4x6
                    requestUriString &= "4x6/0/"
            End Select


            Dim zpl() As Byte = Encoding.UTF8.GetBytes(LabelImage)

            ' adjust print density (8dpmm), label width (4 inches), label height (6 inches), and label index (0) as necessary
            Dim request As HttpWebRequest = WebRequest.Create(requestUriString)
            request.Method = "POST"
            'request.Accept = "application/pdf" ' omit this line to get PNG images back
            request.ContentType = "application/x-www-form-urlencoded"
            request.ContentLength = zpl.Length

            Dim requestStream As Stream = request.GetRequestStream()
            requestStream.Write(zpl, 0, zpl.Length)
            requestStream.Close()

            Try
                Dim response As HttpWebResponse = request.GetResponse()
                Dim responseStream As Stream = response.GetResponseStream()
                Dim fileStream As Stream = File.Create(fn)
                responseStream.CopyTo(fileStream)
                responseStream.Close()
                fileStream.Close()

            Catch e As WebException
                MessageBox.Show(e.Message, "Display ZPL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            'Using frm As New TAFZPLT1
            '    frm.zplPNGFilename = fn
            '    frm.ShowDialog()
            'End Using

            Dim frm As New TAFZPLT1
            frm.zplPNGFilename = fn
            frm.Show()

            My.Computer.FileSystem.DeleteFile(fn)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Display ZPL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Sub PrintRXLabel(ByVal LABRXNUMBER As String,
                            ByVal LABORDERACCOUNT As String,
                            ByVal BAYCOLOR As String,
                            ByVal ORDERNUMBER As String,
                            ByVal CASEBIN As String)

        Dim ZPL_KEY As String = "RX_LABEL"
        rowTATZPLT1 = tblTATZPLT1.Rows.Find(ZPL_KEY)

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{LABRXNUMBER}", LABRXNUMBER)
        labelImage = labelImage.Replace("{LABORDERACCOUNT}", LABORDERACCOUNT)
        labelImage = labelImage.Replace("{BAYCOLOR}", BAYCOLOR)
        labelImage = labelImage.Replace("{ORDERNUMBER}", ORDERNUMBER)
        labelImage = labelImage.Replace("{CASEBIN}", CASEBIN)

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage, "Printig Rx Labels", ZPL_KEY, LABRXNUMBER, 0, ORDERNUMBER, 0)

    End Sub

    Public Sub PrintStoreOrderLabel(ByVal STORE_NO As String, ByVal ORDER_NO As String)

        rowTATZPLT1 = tblTATZPLT1.Rows.Find("STORE_ORDER_LABEL")

        If rowTATZPLT1 Is Nothing Then
            Exit Sub
        End If

        Dim labelImage As String = rowTATZPLT1.Item("ZPL_BODY") & String.Empty
        labelImage = labelImage.Replace("{STORE_NO}", STORE_NO)
        labelImage = labelImage.Replace("{ORDER_NO}", ORDER_NO)

        SendLabelToPrinter(ASCMAIN1.MiniLabelPrinterIPAddress, labelImage)

    End Sub

    Public Sub ShowLabelDialog(ByVal LabelPrinterIP As String, ByVal LabelImage As String)

        Try
            ' "http://api.labelary.com/v1/printers/8dpmm/labels/4x6/0/"
            Dim requestUriString As String = "http://api.labelary.com/v1/printers/8dpmm/labels/"
            Dim fn As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("NOTEPAD") & ".png"
            Dim ResizeTo4by6 As Boolean = False

            Select Case LabelPrinterIP
                Case ASCMAIN1.MiniLabelPrinterIPAddress
                    requestUriString &= "2.25x1.25/0/"

                Case Else
                    requestUriString &= "4x6/0/"
                    ResizeTo4by6 = True
            End Select


            Dim zpl() As Byte = Encoding.UTF8.GetBytes(LabelImage)

            ' adjust print density (8dpmm), label width (4 inches), label height (6 inches), and label index (0) as necessary
            Dim request As HttpWebRequest = WebRequest.Create(requestUriString)
            request.Method = "POST"
            'request.Accept = "application/pdf" ' omit this line to get PNG images back
            request.ContentType = "application/x-www-form-urlencoded"
            request.ContentLength = zpl.Length

            Dim requestStream As Stream = request.GetRequestStream()
            requestStream.Write(zpl, 0, zpl.Length)
            requestStream.Close()

            Try
                Dim response As HttpWebResponse = request.GetResponse()
                Dim responseStream As Stream = response.GetResponseStream()
                Dim fileStream As Stream = File.Create(fn)
                responseStream.CopyTo(fileStream)
                responseStream.Close()
                fileStream.Close()

            Catch e As WebException
                Exit Sub
            End Try

            Using frm As New TAFZPLT1
                frm.zplPNGFilename = fn
                If ResizeTo4by6 Then
                    frm.Rotate = RotateFlipType.Rotate180FlipNone
                    frm.ResizeTo4by6 = True
                End If
                frm.ShowDialog()
            End Using

            My.Computer.FileSystem.DeleteFile(fn)

        Catch ex As Exception

        End Try

    End Sub

#End Region

End Class
