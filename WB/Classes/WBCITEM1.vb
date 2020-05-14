Imports System.Xml
Imports System.Net.Mail

Public Class WBCITEM1

    Private xmlLabelRequest As XmlDocument
    Private productNodeCount As Int16 = 1
    Private nodeShopSite As XmlNode
    Private nodeProducts As XmlNode

    Private tblICTSIZE1 As DataTable = Nothing
    Private tblICTSIZE2 As DataTable = Nothing
    Private tblICTSTYLS As DataTable = Nothing


    Private Const lt = "-&lt"
    Private Const gt = "-&gt"
    Private styleListInactive As List(Of String) = New List(Of String)

    Public Sub New()
        Me.InitiailizeClass()
    End Sub

    ''' <summary>
    ''' Initialze Class variables
    ''' </summary>
    ''' <remarks></remarks>
    ''' 

    Private Sub InitiailizeClass()

        productNodeCount = 1

        tblICTSIZE1 = ASCDATA1.GetDataTable("SELECT * FROM ICTSIZE1")
        tblICTSIZE2 = ASCDATA1.GetDataTable("SELECT * FROM ICTSIZE2")
        tblICTSTYLS = ASCDATA1.GetDataTable("SELECT * FROM ICTSTYLS")

        '"xml", "version=""1.0"" encoding=""UTF-8"""
        xmlLabelRequest = New XmlDocument
        xmlLabelRequest.PreserveWhitespace = True
        xmlLabelRequest.AppendChild(xmlLabelRequest.CreateProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8"" standalone=""no"""))

        nodeShopSite = Nothing
        nodeShopSite = xmlLabelRequest.CreateElement("ShopSiteProducts")

        xmlLabelRequest.AppendChild(nodeShopSite)

        nodeProducts = Nothing
        nodeProducts = xmlLabelRequest.CreateElement("Products")

        nodeShopSite.AppendChild(nodeProducts)

    End Sub

    ''' <summary>
    ''' Clears all product Nodes and sets up for a new XML Document
    ''' </summary>
    ''' <remarks></remarks>
    ''' 

    Public Sub Clear()
        InitiailizeClass()
    End Sub

    ''' <summary>
    ''' Returns the Outer XML for the XML Document Created
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public ReadOnly Property GetXmlOuterXml() As String
        Get
            Return xmlLabelRequest.OuterXml
        End Get
    End Property

    ''' <summary>
    ''' Returns the XMl Document Object created by the class
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public ReadOnly Property GetXMLDocument() As XmlDocument
        Get
            Return xmlLabelRequest
        End Get
    End Property

    ''' <summary>
    ''' Retuns teh number of Product Nodes in the XML Document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public ReadOnly Property NumProductNodes() As Int16
        Get
            Return productNodeCount - 1
        End Get
    End Property

    ''' <summary>
    ''' Adds the Style and all associated Items to the XMLs Products Node
    ''' </summary>
    ''' <param name="StyleCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function AddStyle(ByVal StyleCode As String, ByVal DelList As List(Of String)) As Integer
        styleListInactive = DelList
        Dim nodesProcessed As Integer = 0
        Dim ictr As Integer = 1

        Dim tbl As DataTable = MakeStyleRow(StyleCode)
        For Each rowWBTSTYL1 As DataRow In tbl.Rows
            Dim nodeProduct As XmlNode = Nothing

            If nodesProcessed = 0 Then
                nodeProduct = CreateProductNode(rowWBTSTYL1, True, tbl)
                If nodeProduct IsNot Nothing Then
                    nodeProducts.AppendChild(nodeProduct)
                End If
            End If
            'nodeProduct = MakeProductNode(rowWBTSTYL1, False, Nothing)
            If nodeProduct IsNot Nothing Then
                nodeProducts.AppendChild(nodeProduct)
            End If
            nodesProcessed += 1
        Next
    End Function

    ''' <summary>
    ''' Updates the Inventory level to Shopsite
    ''' </summary>
    ''' <param name="STYLE_CODE"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function AddInventory(ByVal STYLE_CODE As String, ByVal DelList As List(Of String)) As Integer
        styleListInactive = DelList

        Dim tbl As DataTable = MakeStyleRow(STYLE_CODE)
        Dim sql As String = String.Empty
        Dim nodesProcessed As Integer = 0
        Dim ictr As Integer = 1

        For Each rowWBTSTYL1 As DataRow In tbl.Rows
            Dim nodeProduct As XmlNode = Nothing
            nodeProduct = CreateProductInventoryNode(rowWBTSTYL1)
            If nodeProduct IsNot Nothing Then
                nodeProducts.AppendChild(nodeProduct)
            End If

            nodesProcessed += 1
        Next
        Return nodesProcessed
    End Function

    Private Function CreateProductInventoryNode(ByVal rowWBTSTYL1 As DataRow) As XmlNode
        Dim nodeProduct As XmlNode = Nothing
        Dim StyleError As String = ""
        Try
            If rowWBTSTYL1 Is Nothing Then
                Return Nothing
            End If

            nodeProduct = Nothing
            nodeProduct = xmlLabelRequest.CreateElement("Product")

            Dim tblCOLORS As DataTable = RefreshColorTable(rowWBTSTYL1.Item("STYLE_CODE") & String.Empty)
            Dim STYLE_CODE As String = rowWBTSTYL1.Item("STYLE_CODE") & String.Empty
            StyleError = STYLE_CODE

            MakeXMLNode(nodeProduct, "Name", (rowWBTSTYL1.Item("STYLE_DESC") & String.Empty).ToString.Trim.Replace("<", lt).Replace(">", gt))
            MakeXMLNode(nodeProduct, "SKU", STYLE_CODE)
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine("SELECT")
            SQLS.AppendLine("SUM(CASE")
            SQLS.AppendLine("     WHEN (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) < 0 THEN 0")
            SQLS.AppendLine("     ELSE (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0))")
            SQLS.AppendLine("END) AS AVAILABLE")
            SQLS.AppendLine("FROM ICTSTAT2")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine("AND WHSE_CODE = 'MS'")
            ASCMAIN1.sql = SQLS.ToString()
            Dim AVAILABLE As Int16 = Val(ASCDATA1.GetDataValue)
            If AVAILABLE < 0 Then
                AVAILABLE = 0
            End If
            MakeXMLNode(nodeProduct, "QuantityOnHand", AVAILABLE)
            MakeXMLNode(nodeProduct, "MinimumQuantity", GetMinimumQuantity(STYLE_CODE))
            MakeXMLNode(nodeProduct, "SearchKeywords", CreateAttributes(STYLE_CODE))
            Dim nodeOptionMenus As XmlNode = Nothing
            nodeOptionMenus = Nothing
            nodeOptionMenus = xmlLabelRequest.CreateElement("OptionMenus")
            With nodeOptionMenus
                Dim nodeMenu As XmlNode = Nothing
                nodeMenu = xmlLabelRequest.CreateElement("Menu")
                With nodeMenu
                    MakeXMLNode(nodeMenu, "MenuItem", "Color;n")
                    For Each rowCOLORS As DataRow In tblCOLORS.Rows
                        MakeXMLNode(nodeMenu, "MenuItem", rowCOLORS.Item("COLOR_CODE_LONG").ToString())
                    Next
                End With
                .AppendChild(nodeMenu)
            End With
            nodeProduct.AppendChild(nodeOptionMenus)

            MakeXMLNode(nodeProduct, "OptionColumnHeaders", "Color")
            MakeXMLNode(nodeProduct, "OptionAppendSKU", "checked")
            MakeXMLNode(nodeProduct, "OptionUseMultiMenus", "checked")
            MakeXMLNode(nodeProduct, "OptionSelectDefault")

            'Re-Load this if you need to re-populate all the ProductField11 fields again.

            If rowWBTSTYL1.Item("STYLE_SORT") & "" = "" Then
                MakeXMLNode(nodeProduct, "ProductField11", STYLE_CODE)
            Else
                MakeXMLNode(nodeProduct, "ProductField11", rowWBTSTYL1.Item("STYLE_SORT"))
            End If
            MakeXMLNode(nodeProduct, "ProductField12", "wholesale|wholesale1")
            MakeXMLNode(nodeProduct, "ProductField1", GetCOLOR_GROUP_CODE(STYLE_CODE))
            Dim default_image As String = "product/" & rowWBTSTYL1.Item("DEFAULT_IMAGE").ToString & ""
            default_image = default_image.Replace(".JPG", ".jpg")
            MakeXMLNode(nodeProduct, "Graphic", default_image)
            MakeXMLNode(nodeProduct, "ProductImageSize", "2")
            Dim nodeProductOptions As XmlNode = Nothing
            nodeProductOptions = xmlLabelRequest.CreateElement("ProductOptions")
            With nodeProductOptions
                For Each rowCOLOR As DataRow In tblCOLORS.Rows
                    Dim nodeProductOption As XmlNode = Nothing
                    nodeProductOption = xmlLabelRequest.CreateElement("ProductOption")
                    Dim typeAttr As XmlAttribute = xmlLabelRequest.CreateAttribute("Name")
                    typeAttr.Value = rowCOLOR.Item("COLOR_CODE_LONG")
                    nodeProductOption.Attributes.Append(typeAttr)
                    With nodeProductOption
                        MakeXMLNode(nodeProductOption, "Use", "checked")
                        MakeXMLNode(nodeProductOption, "Menu1", rowCOLOR.Item("COLOR_CODE_LONG"))
                        For i As Integer = 2 To 4
                            MakeXMLNode(nodeProductOption, "Menu" & i.ToString)
                        Next
                        MakeXMLNode(nodeProductOption, "AppendText", rowCOLOR.Item("COLOR_CODE"))
                        'This was changed based on testing done with Mario. 5/1/14.
                        'MakeXMLNode(nodeProductOption, "SKU", String.Format("{0}-{1}", STYLE_CODE, rowCOLOR.Item("COLOR_CODE")))
                        MakeXMLNode(nodeProductOption, "SKU", "-" & rowCOLOR.Item("COLOR_CODE"))
                        MakeXMLNode(nodeProductOption, "PriceModifier")
                        MakeXMLNode(nodeProductOption, "WeightModifier")
                        MakeXMLNode(nodeProductOption, "QuantityOnHand", Val(rowCOLOR.Item("MSOH") & "") + Val(rowCOLOR.Item("MSFT") & ""))
                        MakeXMLNode(nodeProductOption, "LowStockThreshold", Val(rowCOLOR.Item("MSOH") & "") + Val(rowCOLOR.Item("MSFT") & ""))
                        MakeXMLNode(nodeProductOption, "OutofStockLimit", "0")
                        If rowCOLOR.Item("IMG_NAME") & "" <> "" Then
                            Dim ImageFile As String = rowCOLOR.Item("IMG_NAME")
                            With ImageFile
                                If .Length > 3 Then
                                    If .Substring(.Length - 3, 3) = "JPG" Then
                                        ImageFile = .Substring(0, .Length - 3) & "jpg"
                                    End If
                                End If
                            End With
                            MakeXMLNode(nodeProductOption, "Image", "product/" & ImageFile)
                        Else
                            If (ASCMAIN1.Running_in_VS) Then
                                'Stop
                            End If
                        End If
                        
                    End With
                    .AppendChild(nodeProductOption)
                Next
            End With
            nodeProduct.AppendChild(nodeProductOptions)

            Dim nodeQuantityPricing As XmlNode = FillQuantityPricing(STYLE_CODE)
            nodeProduct.AppendChild(nodeQuantityPricing)

            'MakeProductOnOagesNode
            Dim nodeProductOnOages As XmlNode = xmlLabelRequest.CreateElement("ProductOnPages")
            With nodeProductOnOages
                Dim nodeName As XmlNode = xmlLabelRequest.CreateElement("Name")
                'I Know this is dumb but they keep changing their minds on classes.
                Dim STYLE_CLASS_DESC As String = ""
                Select Case rowWBTSTYL1.Item("STYLE_CLASS_DESC") & String.Empty
                    Case "Fall"
                        STYLE_CLASS_DESC = "HARVEST"
                    Case Else
                        STYLE_CLASS_DESC = rowWBTSTYL1.Item("STYLE_CLASS_DESC") & String.Empty
                End Select
                If styleListInactive.Contains(rowWBTSTYL1.Item("STYLE_CODE") & "") Then
                    nodeName.InnerText = ""
                    .AppendChild(nodeName)
                Else
                    nodeName.InnerText = STYLE_CLASS_DESC
                    .AppendChild(nodeName)
                    Dim SQLS1 As New System.Text.StringBuilder() With {.Length = 0}
                    SQLS1.AppendLine(String.Format("SELECT NVL(STYLE_STATUS,'A') AS STYLE_STATUS FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    ASCMAIN1.sql = SQLS1.ToString()
                    Dim STYLE_STATUS As String = ASCDATA1.GetDataValue
                    If STYLE_STATUS = "D" Then
                        Dim nodeDisc As XmlNode = xmlLabelRequest.CreateElement("Name")
                        nodeDisc.InnerText = "Discontinued"
                        .AppendChild(nodeDisc)
                    End If
                End If
            End With
            nodeProduct.AppendChild(nodeProductOnOages)

            Dim InjectScripts As New System.Text.StringBuilder
            InjectScripts.Length = 0
            InjectScripts.AppendLine("<script src='js/productpage.js'></script>")
            InjectScripts.AppendLine("<div id='injectinv'></div>")
            InjectScripts.AppendLine("<script src='http://ajax.googleapis.com/ajax/libs/jquery/1.8.0/jquery.min.js' type='text/javascript'></script>")
            InjectScripts.AppendLine("<script>")
            InjectScripts.AppendLine("$.get('invtbl/" & STYLE_CODE & ".html', function( data ) {")
            InjectScripts.AppendLine("     $('#injectinv').html(data);")
            InjectScripts.AppendLine("    });")
            InjectScripts.AppendLine("</script>")
            InjectScripts.AppendLine("<style>")
            InjectScripts.AppendLine("div#menu-left{")
            InjectScripts.AppendLine("background: #E2FFCF;")
            InjectScripts.AppendLine("color: #445a3d;")
            InjectScripts.AppendLine("}")
            InjectScripts.AppendLine("</style>")
            MakeXMLNode(nodeProduct, "MoreInformationText", InjectScripts.ToString)


            'Dim BASE As New ASFBASE0
            'Dim nodeQuantityPricing As XmlNode = xmlLabelRequest.CreateElement("QuantityPricing")
            'With nodeQuantityPricing
            '    MakeXMLNode(nodeQuantityPricing, "Comment", GetComments(BASE, STYLE_CODE))
            'End With
            'nodeProduct.AppendChild(nodeQuantityPricing)

        Catch ex As Exception
            nodeProduct = Nothing
            SendErrorEMail(StyleError & " - " & ex.Message, False)
        End Try

        Return nodeProduct
    End Function
    Private Sub SendErrorEMail(ByVal MsgBody As String, Optional ByVal StopProcess As Boolean = True)

        Const FROM_ADDRESS As String = "new.accounts@regency-rib.com"
        Const FROM_NAME As String = "Regency Auto-Sync Manager"
        Const SERVER_IP As String = "192.168.110.221"
        Const SERVER_PORT As Integer = 25
        Const SERVER_ACCOUNT As String = "new.accounts@regency-rib.com"
        Const SERVER_PASSWORD As String = "0ff1c3"
        Const EMAIL_ADDRESS As String = "mariog@regency-rib.comt"
        Const EMAIL_NAME As String = "Mario Arenas Jr."
        Dim CC_ADDRESS As String = "whr@waynerichmond.net"
        Dim CC_NAME As String = "Wayne Richmond"


        Try
            Dim HTMLBody As String = "A Message Was Reported From The Regency Auto-Refresh Process As Follows:" & vbCrLf & vbCrLf & MsgBody
            Dim mail As New MailMessage() With {.From = New MailAddress(FROM_ADDRESS, FROM_NAME)}
            mail.To.Add(New MailAddress(EMAIL_ADDRESS, EMAIL_NAME))
            mail.Subject = "Message From Regency Auto-Refresh Process"
            mail.IsBodyHtml = True
            mail.Body = HTMLBody
            'If Not WayneOnly Then
            mail.CC.Add(New MailAddress(CC_ADDRESS, CC_NAME))
            'End If

            Dim smtp As New SmtpClient(SERVER_IP, SERVER_PORT)
            If smtp IsNot Nothing Then
                smtp.Credentials = New System.Net.NetworkCredential(SERVER_ACCOUNT, SERVER_PASSWORD)
            Else
                MsgBox("SMTP Client could not be created.", MsgBoxStyle.OkOnly, "Error")
            End If

            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                smtp.Send(mail)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Sending Mail")
        End Try
    End Sub

    Private Shared Function GetMinimumQuantity(ByVal STYLE_CODE As String) As Int16
        Dim MinimumQuantity As Int16 = 0
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine(String.Format("Select NVL(STYLE_SO_QTY_MIN,1) from ICTSTYL1 where STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        MinimumQuantity = Val(ASCDATA1.GetDataValue)
        If MinimumQuantity = 0 Then
            MinimumQuantity = 1
        End If

        SQLS.Length = 0
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("SUM(CASE")
        SQLS.AppendLine("     WHEN (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0)) < 0 THEN 0")
        SQLS.AppendLine("     ELSE (NVL(WHSE_QTY_ON_HAND,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0) - NVL(WHSE_QTY_PICK,0))")
        SQLS.AppendLine("END) AS AVAILABLE")
        SQLS.AppendLine("FROM ICTSTAT2")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine("AND WHSE_CODE = 'MS'")
        ASCMAIN1.sql = SQLS.ToString()
        Dim AVAILABLE As Int16 = Val(ASCDATA1.GetDataValue)
        If AVAILABLE < 0 Then
            AVAILABLE = 0
        End If

        If AVAILABLE < MinimumQuantity Then
            MinimumQuantity = AVAILABLE
        End If

        Return MinimumQuantity
    End Function

    ''' <summary>
    ''' Creates each Product Node for the XML document
    ''' </summary>
    ''' <param name="rowWBTSTYL1"></param>
    ''' <param name="mainStyle"></param>
    ''' <param name="tblSubProducts"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Private Function CreateProductNode(ByVal rowWBTSTYL1 As DataRow, ByVal mainStyle As Boolean, ByRef tblSubProducts As DataTable) As XmlNode

        'Dim innerText As String = String.Empty
        'Dim node As XmlNode = Nothing
        'Dim sql As String = String.Empty
        'Dim rowICTSTAT2 As DataRow = Nothing
        Dim nodeProduct As XmlNode = Nothing

        Try

            If rowWBTSTYL1 Is Nothing Then
                Return Nothing
            End If

            'Dim rowASTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ASTPARM1 WHERE AS_PARM_KEY = 'Z'")

            nodeProduct = Nothing
            nodeProduct = xmlLabelRequest.CreateElement("Product")



            Dim tblCOLORS As DataTable = RefreshColorTable(rowWBTSTYL1.Item("STYLE_CODE") & String.Empty)
            Dim STYLE_CODE As String = rowWBTSTYL1.Item("STYLE_CODE") & String.Empty
            'Dim Default_Image As String = GetImageName(STYLE_CODE, tblCOLORS.Rows(0).Item("COLOR_CODE").ToString & String.Empty)
            'Dim Default_Image As String = GetImageName(STYLE_CODE, tblCOLORS.Rows(0).Item("COLOR_CODE").ToString & String.Empty)
            'If Default_Image.Length = 0 Then
            'Dim Default_Image As String = String.Format("{0}-{1}.jpg", STYLE_CODE, tblCOLORS.Rows(0).Item("COLOR_CODE"))

            Dim Default_Image As String = ""
            Dim First_Image As String = ""
            For Each rowWBTSTYL2 As DataRow In tblCOLORS.Rows
                If rowWBTSTYL2.Item("COLOR_STATUS").ToString = "A" Then
                    Default_Image = String.Format("{0}-{1}.jpg", STYLE_CODE, rowWBTSTYL2.Item("COLOR_CODE"))
                    Exit For
                Else
                    If Val(rowWBTSTYL2.Item("MSOH").ToString) > 0 Or Val(rowWBTSTYL2.Item("MSFT").ToString) > 0 Then
                        Default_Image = String.Format("{0}-{1}.jpg", STYLE_CODE, rowWBTSTYL2.Item("COLOR_CODE"))
                        Exit For
                    Else
                        If First_Image.Length = 0 Then
                            First_Image = String.Format("{0}-{1}.jpg", STYLE_CODE, rowWBTSTYL2.Item("COLOR_CODE"))
                        End If
                    End If
                End If
            Next
            If Default_Image.Length = 0 Then
                If First_Image.Length > 0 Then
                    Default_Image = First_Image
                Else
                    MsgBox("Could Not Assign Default Image for Style " & STYLE_CODE, MsgBoxStyle.Critical, "Gotta Problem Here")
                    Stop
                End If
            End If
            'End If

            MakeXMLNode(nodeProduct, "Name", (rowWBTSTYL1.Item("STYLE_DESC") & String.Empty).ToString.Trim.Replace("<", lt).Replace(">", gt))
            MakeXMLNode(nodeProduct, "SKU", STYLE_CODE)
            MakeXMLNode(nodeProduct, "Price", "0.00")
            MakeXMLNode(nodeProduct, "SaleAmount")
            MakeXMLNode(nodeProduct, "ProductDisabled", "uncheck")
            MakeXMLNode(nodeProduct, "Taxable", "checked")
            MakeXMLNode(nodeProduct, "AvaTaxCode")
            MakeXMLNode(nodeProduct, "VAT", "0")
            MakeXMLNode(nodeProduct, "Weight", "0")
            MakeXMLNode(nodeProduct, "QuantityOnHand")
            MakeXMLNode(nodeProduct, "LowStockThreshhold")
            MakeXMLNode(nodeProduct, "OutOfStockLimit")
            MakeXMLNode(nodeProduct, "GroundShipping", "0.00")
            MakeXMLNode(nodeProduct, "SecondDayShipping", "0.00")
            MakeXMLNode(nodeProduct, "NextDayShipping", "0.00")
            For w As Integer = 3 To 9
                MakeXMLNode(nodeProduct, String.Format("Shipping{0}", w))
            Next
            MakeXMLNode(nodeProduct, "Graphic", "product/" & Default_Image)
            MakeXMLNode(nodeProduct, "ProductImageSize", "2")
            MakeXMLNode(nodeProduct, "SearchKeywords", CreateAttributes(STYLE_CODE))
            MakeXMLNode(nodeProduct, "SearchMakePage")
            MakeXMLNode(nodeProduct, "ProductDescription")
            MakeXMLNode(nodeProduct, "ProductGUID", "")
            MakeXMLNode(nodeProduct, "OptionText")
            'OptionMenus
            Dim nodeOptionMenus As XmlNode = Nothing
            nodeOptionMenus = Nothing
            nodeOptionMenus = xmlLabelRequest.CreateElement("OptionMenus")
            With nodeOptionMenus
                Dim nodeMenu As XmlNode = Nothing
                nodeMenu = xmlLabelRequest.CreateElement("Menu")
                With nodeMenu
                    MakeXMLNode(nodeMenu, "MenuItem", "Color;n")
                    For Each rowCOLORS As DataRow In tblCOLORS.Rows
                        If rowCOLORS.Item("COLOR_STATUS").ToString = "A" Or (Val(rowCOLORS.Item("MSOH").ToString) > 0 Or Val(rowCOLORS.Item("MSFT").ToString) > 0) Then
                            MakeXMLNode(nodeMenu, "MenuItem", rowCOLORS.Item("COLOR_CODE_LONG").ToString())
                        End If
                    Next
                End With
                .AppendChild(nodeMenu)
            End With
            nodeProduct.AppendChild(nodeOptionMenus)

            MakeXMLNode(nodeProduct, "OptionColumnHeaders", "Color")
            MakeXMLNode(nodeProduct, "OptionAppendSKU", "checked")
            MakeXMLNode(nodeProduct, "OptionUseMultiMenus", "checked")
            MakeXMLNode(nodeProduct, "OptionSelectDefault")

            'ProductOptions
            Dim nodeProductOptions As XmlNode = Nothing
            nodeProductOptions = xmlLabelRequest.CreateElement("ProductOptions")
            With nodeProductOptions
                Dim rowFilter As String = "IMG_FOUND = '1'"
                For Each rowCOLOR As DataRow In tblCOLORS.Select(rowFilter)
                    If rowCOLOR.Item("COLOR_STATUS").ToString = "A" Or (Val(rowCOLOR.Item("MSOH").ToString) > 0 Or Val(rowCOLOR.Item("MSFT").ToString) > 0) Then
                        Dim nodeProductOption As XmlNode = Nothing
                        nodeProductOption = xmlLabelRequest.CreateElement("ProductOption")
                        Dim typeAttr As XmlAttribute = xmlLabelRequest.CreateAttribute("Name")
                        typeAttr.Value = rowCOLOR.Item("COLOR_CODE_LONG")
                        nodeProductOption.Attributes.Append(typeAttr)
                        With nodeProductOption
                            MakeXMLNode(nodeProductOption, "Use", "checked")
                            MakeXMLNode(nodeProductOption, "Menu1", rowCOLOR.Item("COLOR_CODE_LONG"))
                            For i As Integer = 2 To 4
                                MakeXMLNode(nodeProductOption, "Menu" & i.ToString)
                            Next
                            MakeXMLNode(nodeProductOption, "AppendText", rowCOLOR.Item("COLOR_CODE"))
                            'Changed based on testing with Mario. 5/1/14.
                            'MakeXMLNode(nodeProductOption, "SKU", String.Format("{0}-{1}", STYLE_CODE, rowCOLOR.Item("COLOR_CODE")))
                            MakeXMLNode(nodeProductOption, "SKU", "-" & rowCOLOR.Item("COLOR_CODE"))
                            MakeXMLNode(nodeProductOption, "PriceModifier")
                            MakeXMLNode(nodeProductOption, "WeightModifier")
                            MakeXMLNode(nodeProductOption, "QuantityOnHand", Val(rowCOLOR.Item("MSOH") & "") + Val(rowCOLOR.Item("MSFT") & ""))
                            MakeXMLNode(nodeProductOption, "LowStockThreshold", Val(rowCOLOR.Item("MSOH") & "") + Val(rowCOLOR.Item("MSFT") & ""))
                            MakeXMLNode(nodeProductOption, "OutofStockLimit", "0")
                            Dim ImageFile As String = rowCOLOR.Item("IMG_NAME")
                            With ImageFile
                                If .Length > 3 Then
                                    If .Substring(.Length - 3, 3) = "JPG" Then
                                        ImageFile = .Substring(0, .Length - 3) & "jpg"
                                    End If
                                End If
                            End With
                            MakeXMLNode(nodeProductOption, "Image", "product/" & ImageFile)
                        End With
                        .AppendChild(nodeProductOption)
                    End If
                Next
            End With
            nodeProduct.AppendChild(nodeProductOptions)

            MakeXMLNode(nodeProduct, "CustomerTextEntryBox", "uncheck")
            MakeXMLNode(nodeProduct, "CustomerTextEntryheader")
            MakeXMLNode(nodeProduct, "CustomerTextEntryColumns", "40")
            MakeXMLNode(nodeProduct, "CustomerTextEntryRows", "4")
            MakeXMLNode(nodeProduct, "CrossSell")

            'MakeProductOnOagesNode
            Dim nodeProductOnOages As XmlNode = Nothing
            nodeProductOnOages = Nothing
            nodeProductOnOages = xmlLabelRequest.CreateElement("ProductOnPages")
            With nodeProductOnOages
                Dim nodeName As XmlNode = Nothing
                nodeName = xmlLabelRequest.CreateElement("Name")
                'I Know this is dumb but they keep changing their minds on classes.
                Dim STYLE_CLASS_DESC As String = ""
                Select Case rowWBTSTYL1.Item("STYLE_CLASS_DESC") & String.Empty
                    Case "Fall"
                        STYLE_CLASS_DESC = "HARVEST"
                    Case Else
                        STYLE_CLASS_DESC = rowWBTSTYL1.Item("STYLE_CLASS_DESC") & String.Empty
                End Select
                If styleListInactive.Contains(rowWBTSTYL1.Item("STYLE_CODE") & "") Then
                    nodeName.InnerText = ""
                    .AppendChild(nodeName)
                Else
                    nodeName.InnerText = STYLE_CLASS_DESC
                    .AppendChild(nodeName)
                    Dim SQLS1 As New System.Text.StringBuilder
                    SQLS1.Length = 0
                    SQLS1.AppendLine(String.Format("SELECT NVL(STYLE_STATUS,'A') AS STYLE_STATUS FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    ASCMAIN1.sql = SQLS1.ToString()
                    Dim STYLE_STATUS As String = ASCDATA1.GetDataValue
                    If STYLE_STATUS = "D" Then
                        Dim nodeDisc As XmlNode = Nothing
                        nodeDisc = xmlLabelRequest.CreateElement("Name")
                        nodeDisc.InnerText = "Discontinued"
                        .AppendChild(nodeDisc)
                    End If
                End If
            End With
            nodeProduct.AppendChild(nodeProductOnOages)

            MakeXMLNode(nodeProduct, "AddToPages")

            MakeXMLNode(nodeProduct, "DisplayMoreInformationPage_", "checked")
            Dim InjectScripts As New System.Text.StringBuilder
            InjectScripts.Length = 0
            InjectScripts.AppendLine("<script src='js/productpage.js'></script>")
            InjectScripts.AppendLine("<div id='injectinv'></div>")
            InjectScripts.AppendLine("<script src='http://ajax.googleapis.com/ajax/libs/jquery/1.8.0/jquery.min.js' type='text/javascript'></script>")
            InjectScripts.AppendLine("<script>")
            InjectScripts.AppendLine("$.get('invtbl/" & STYLE_CODE & ".html', function( data ) {")
            InjectScripts.AppendLine("     $('#injectinv').html(data);")
            InjectScripts.AppendLine("    });")
            InjectScripts.AppendLine("</script>")
            InjectScripts.AppendLine("<style>")
            InjectScripts.AppendLine("div#menu-left{")
            InjectScripts.AppendLine("background: #E2FFCF;")
            InjectScripts.AppendLine("color: #445a3d;")
            InjectScripts.AppendLine("}")
            InjectScripts.AppendLine("</style>")
            MakeXMLNode(nodeProduct, "MoreInformationText", InjectScripts.ToString)
            MakeXMLNode(nodeProduct, "MoreInformationGraphic", "product/" & Default_Image)
            MakeXMLNode(nodeProduct, "MoreInformationSize", "1")
            MakeXMLNode(nodeProduct, "MoreInformationTitle")
            MakeXMLNode(nodeProduct, "MoreInformationMetaKeywords")
            MakeXMLNode(nodeProduct, "MoreInfoMetaDescription")
            MakeXMLNode(nodeProduct, "FileName", STYLE_CODE & ".html")
            MakeXMLNode(nodeProduct, "ProductSitemap", "checked")
            MakeXMLNode(nodeProduct, "ProductSitemapPriority", "Google Default")
            MakeXMLNode(nodeProduct, "ProductCrossSell", "checked")
            MakeXMLNode(nodeProduct, "GlobalCrossSell", "uncheck")
            For i As Integer = 1 To 20
                MakeXMLNode(nodeProduct, "MoreInfoImage" & i.ToString, "none")
            Next
            MakeXMLNode(nodeProduct, "MoreInfoImageExtraSize", "3")
            MakeXMLNode(nodeProduct, "Template", "Dynamic-Product_wholesale.sst")
            MakeXMLNode(nodeProduct, "DisplayName", "checked")
            MakeXMLNode(nodeProduct, "DisplaySKU", "checked")
            MakeXMLNode(nodeProduct, "DisplayPrice", "uncheck")
            MakeXMLNode(nodeProduct, "DisplayGraphic", "checked")
            MakeXMLNode(nodeProduct, "SaleOn", "uncheck")
            MakeXMLNode(nodeProduct, "NameStyle", "Bold")
            MakeXMLNode(nodeProduct, "NameSize", "Normal")
            MakeXMLNode(nodeProduct, "PriceStyle", "Plain")
            MakeXMLNode(nodeProduct, "PriceSize", "Normal")
            MakeXMLNode(nodeProduct, "SKUStyle", "Plain")
            MakeXMLNode(nodeProduct, "SKUSize", "Small")
            MakeXMLNode(nodeProduct, "DescriptionStyle", "Plain")
            MakeXMLNode(nodeProduct, "DescriptionSize", "Normal")
            MakeXMLNode(nodeProduct, "ImageAlignment", "Center")
            MakeXMLNode(nodeProduct, "TextWrap", "On")
            MakeXMLNode(nodeProduct, "AddtoCartButton", "Add To Cart")
            MakeXMLNode(nodeProduct, "ViewCartButton", "View Cart")
            MakeXMLNode(nodeProduct, "ProductType", "Tangible")
            MakeXMLNode(nodeProduct, "DisplayOrderQuantity_", "checked")
            MakeXMLNode(nodeProduct, "DisplayOrderingOptions_", "checked")
            MakeXMLNode(nodeProduct, "UseAddtoCartImage_", "0")
            MakeXMLNode(nodeProduct, "AddtoCartImage", "[shopsite-images]/buttons/defaults/Add_To_Cart.gif")
            MakeXMLNode(nodeProduct, "UseViewCartImage_", "0")
            MakeXMLNode(nodeProduct, "ViewCartImage", "[shopsite-images]/buttons/defaults/View_Cart.gif")
            MakeXMLNode(nodeProduct, "ViewCartButton", "View Cart")
            MakeXMLNode(nodeProduct, "ProductDownloadLocation", "none")
            MakeXMLNode(nodeProduct, "SearchDestType", "selected")
            MakeXMLNode(nodeProduct, "SearchDest", "Store")
            MakeXMLNode(nodeProduct, "DimensionOptions", "1")
            MakeXMLNode(nodeProduct, "DimensionText")
            MakeXMLNode(nodeProduct, "DimensionSelected")
            MakeXMLNode(nodeProduct, "FedExContainer")
            MakeXMLNode(nodeProduct, "USPSContainer")
            MakeXMLNode(nodeProduct, "NoShippingCharges", "uncheck")
            MakeXMLNode(nodeProduct, "ExtraHandlingCharge", "0.00")
            MakeXMLNode(nodeProduct, "ProhibitedShippingMethods")
            MakeXMLNode(nodeProduct, "QBImport")
            MakeXMLNode(nodeProduct, "MinimumQuantity", GetMinimumQuantity(STYLE_CODE))
            Dim nodeQuantityPricing As XmlNode = FillQuantityPricing(STYLE_CODE)
            nodeProduct.AppendChild(nodeQuantityPricing)
            MakeXMLNode(nodeProduct, "QuantityPricingGroup")
            MakeXMLNode(nodeProduct, "DisplayQuantityPricing", "checked")
            MakeXMLNode(nodeProduct, "VariablePrice_", "uncheck")
            MakeXMLNode(nodeProduct, "VariableName_", "uncheck")
            MakeXMLNode(nodeProduct, "VariableSKU_", "uncheck")
            MakeXMLNode(nodeProduct, "GoogleBase", "unchecked")
            MakeXMLNode(nodeProduct, "Brand")
            MakeXMLNode(nodeProduct, "GTIN")
            MakeXMLNode(nodeProduct, "ManufacturerPartNumber")
            MakeXMLNode(nodeProduct, "GoogleProductType")
            MakeXMLNode(nodeProduct, "GoogleProductCategory")
            MakeXMLNode(nodeProduct, "Availability", "in stock")
            MakeXMLNode(nodeProduct, "GoogleCondition", "New")
            MakeXMLNode(nodeProduct, "GoogleAgeGroup", "none")
            MakeXMLNode(nodeProduct, "GoogleGender", "none")
            MakeXMLNode(nodeProduct, "GoogleUseAdvancedOrderingOptions", "unchecked")
            MakeXMLNode(nodeProduct, "GoogleColorColumn")
            MakeXMLNode(nodeProduct, "GoogleSizeColumn")
            MakeXMLNode(nodeProduct, "GooglePatternColumn")
            MakeXMLNode(nodeProduct, "GoogleMaterialColumn")
            MakeXMLNode(nodeProduct, "GoogleListAsFreeShipping", "uncheck")
            MakeXMLNode(nodeProduct, "DobaItemID")
            MakeXMLNode(nodeProduct, "SubscriptionProduct")
            MakeXMLNode(nodeProduct, "SubPaymentIntervalUnit", "monthly")
            MakeXMLNode(nodeProduct, "SubRegularPaymentAmount")
            MakeXMLNode(nodeProduct, "SubEndingPeriod", "never")
            MakeXMLNode(nodeProduct, "SubBillOn", "order_date")
            MakeXMLNode(nodeProduct, "SubTrialCheckBox")
            MakeXMLNode(nodeProduct, "SubTrialPaymentAmount")
            MakeXMLNode(nodeProduct, "SubTrialPeriods")
            MakeXMLNode(nodeProduct, "SubOneTimeCheckBox")
            MakeXMLNode(nodeProduct, "SubOneTimePaymentAmount")
            MakeXMLNode(nodeProduct, "Subproducts")
            For i As Integer = 1 To 25
                If i = 1 Then
                    MakeXMLNode(nodeProduct, "ProductField1", GetMinimumQuantity(STYLE_CODE))
                Else
                    If i > 10 Then
                        Select Case i
                            Case Is = 11
                                If rowWBTSTYL1.Item("STYLE_SORT") & "" = "" Then
                                    MakeXMLNode(nodeProduct, "ProductField" & i.ToString, STYLE_CODE)
                                Else
                                    MakeXMLNode(nodeProduct, "ProductField" & i.ToString, rowWBTSTYL1.Item("STYLE_SORT"))
                                End If
                            Case Is = 12
                                MakeXMLNode(nodeProduct, "ProductField" & i.ToString, "wholesale|wholesale1")
                            Case Else
                                MakeXMLNode(nodeProduct, "ProductField" & i.ToString)
                        End Select
                    Else
                        MakeXMLNode(nodeProduct, "ProductField" & i.ToString, "F" & i.ToString)
                    End If
                End If
            Next
            productNodeCount += 1
        Catch ex As Exception
            nodeProduct = Nothing

        End Try

        Return nodeProduct
    End Function

    Private Function CreateAttributes(ByVal STYLE_CODE As String) As String
        Dim SQLS As New System.Text.StringBuilder
        Dim ATTRS As String = ""
        SQLS.Length = 0
        SQLS.AppendLine("SELECT A1.ATTR_DESC")
        SQLS.AppendLine("FROM ICTSTYL3 I3, ICTATTR1 A1")
        SQLS.AppendLine("WHERE I3.ATTR_CODE = A1.ATTR_CODE")
        SQLS.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine("AND NVL(A1.ATTR_DESC,'NULL') <> 'NULL'")
        SQLS.AppendLine("UNION")
        SQLS.AppendLine("SELECT C1.STYLE_CLASS_DESC AS ATTR_DESC")
        SQLS.AppendLine("FROM ICTSTYL1 S1, ICTCLAS1 C1")
        SQLS.AppendLine("WHERE S1.STYLE_CLASS_CODE = C1.STYLE_CLASS_CODE")
        SQLS.AppendLine(String.Format("AND S1.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine("UNION")
        SQLS.AppendLine("SELECT Z1.SIZE_DESC AS ATTR_DESC")
        SQLS.AppendLine("FROM ICTSTYL1 S1, ICTSIZE1 Z1")
        SQLS.AppendLine("WHERE S1.SIZE_CODE = Z1.SIZE_CODE")
        SQLS.AppendLine(String.Format("AND S1.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine("UNION")
        SQLS.AppendLine("SELECT W1.ATTR_DESC")
        SQLS.AppendLine("FROM ICTSTYL3 I3, ICTATTR1 A1, WBTATTR1 W1")
        SQLS.AppendLine("WHERE I3.ATTR_CODE = A1.ATTR_CODE")
        SQLS.AppendLine("AND A1.ATTR_CODE = W1.ATTR_CODE")
        SQLS.AppendLine(String.Format("AND I3.STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine("AND NVL(A1.ATTR_DESC,'NULL') <> 'NULL'")
        For Each rowATTR_CODE As DataRow In ASCDATA1.GetDataTable(SQLS.ToString(), String.Empty, "V", STYLE_CODE).Rows
            'ATTRS = String.Format("{0},{1}", ATTRS, rowATTR_CODE.Item("ATTR_DESC"))
            Dim ATR As String = rowATTR_CODE.Item("ATTR_DESC").ToString & ""
            If ATR = "DÉCOR" Then
                ATR = "DECOR"
            End If
            ATTRS += ATR.ToLower & ", "
            ATTRS += ATR.ToUpper & ", "
            ATTRS += StrConv(ATR.ToLower, VbStrConv.ProperCase) & ", "
        Next
        Dim SQLS1 As New System.Text.StringBuilder
        SQLS1.Length = 0
        SQLS1.AppendLine(String.Format("SELECT NVL(STYLE_STATUS,'A') AS STYLE_STATUS FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS1.ToString()
        Dim STYLE_STATUS As String = ASCDATA1.GetDataValue
        If STYLE_STATUS = "D" Then
            ATTRS = String.Format("{0}, {1}", ATTRS, "Discontinued")
        End If
        ATTRS = ATTRS.Substring(0, ATTRS.Length - 2)
        Return ATTRS
    End Function

    Private Sub MakeXMLNode(ByRef NodeToAppend As XmlNode, _
                                 ByVal Element As String, Optional ByVal InnerText As String = "", _
                                 Optional ByVal Attr As String = "", _
                                 Optional ByVal AttrValue As String = "")
        Dim node As XmlNode = Nothing
        node = xmlLabelRequest.CreateElement(Element)
        If Attr.Length > 0 Then
            Dim typeAttr As XmlAttribute = xmlLabelRequest.CreateAttribute(Attr)
            typeAttr.Value = AttrValue
            node.Attributes.Append(typeAttr)
        End If
        node.InnerText = InnerText
        NodeToAppend.AppendChild(node)
    End Sub

    Private Function GetImageName(ByVal STYLE_CODE As Object, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE)
        Dim url As New System.Uri("http://api.regency-rib.com:8181/images/product/" & RetVal)
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
        Dim resptest As System.Net.WebResponse
        Dim ErrorsFound As Boolean = False
        req.Timeout = 3000
        Try
            resptest = req.GetResponse()
        Catch ex As Exception
            ErrorsFound = True
            RetVal = ""
            req = Nothing
        End Try
        Return RetVal
    End Function

    Private Function RefreshColorTable(ByVal Style_Code As String) As DataTable
        Dim Padded As String = StrDup(50, " ")
        Dim SQLS As New System.Text.StringBuilder
        'SQLS.Length = 0
        'SQLS.AppendLine("SELECT * FROM")
        'SQLS.AppendLine("  (")
        'SQLS.AppendLine("   SELECT C1.STYLE_CODE, C1.COLOR_CODE,")
        'SQLS.AppendLine("   C2.COLOR_DESC AS COLOR_CODE_LONG,")
        'SQLS.AppendLine("   C1.STYLE_COLOR_STATUS,")
        'SQLS.AppendLine("   CASE WHEN")
        'SQLS.AppendLine("   SUM(")
        'SQLS.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        'SQLS.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        'SQLS.AppendLine("     ELSE 0")
        'SQLS.AppendLine("     END) < 0")
        'SQLS.AppendLine("   THEN")
        'SQLS.AppendLine("     0")
        'SQLS.AppendLine("   ELSE")
        'SQLS.AppendLine("   SUM(")
        'SQLS.AppendLine("     CASE S2.WHSE_CODE")
        'SQLS.AppendLine("     WHEN 'MS'")
        'SQLS.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        'SQLS.AppendLine("     ELSE 0")
        'SQLS.AppendLine("     END)")
        'SQLS.AppendLine("   END AS MSOH,")
        'SQLS.AppendLine("   CASE WHEN")
        'SQLS.AppendLine("   SUM(")
        'SQLS.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        'SQLS.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'SQLS.AppendLine("     ELSE 0")
        'SQLS.AppendLine("     END) <= 0")
        'SQLS.AppendLine("   THEN")
        'SQLS.AppendLine("     0")
        'SQLS.AppendLine("   ELSE")
        'SQLS.AppendLine("     CASE WHEN")
        'SQLS.AppendLine("       SUM(")
        'SQLS.AppendLine("       CASE S2.WHSE_CODE")
        'SQLS.AppendLine("       WHEN 'MS'")
        'SQLS.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'SQLS.AppendLine("       ELSE 0")
        'SQLS.AppendLine("       END) < 0")
        'SQLS.AppendLine("     THEN")
        'SQLS.AppendLine("       0")
        'SQLS.AppendLine("     ELSE")
        'SQLS.AppendLine("     SUM(")
        'SQLS.AppendLine("       CASE S2.WHSE_CODE")
        'SQLS.AppendLine("       WHEN 'MS'")
        'SQLS.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'SQLS.AppendLine("       ELSE 0")
        'SQLS.AppendLine("       END) END")
        'SQLS.AppendLine("   END AS MSFT,")
        'SQLS.AppendLine(String.Format("   '{0}' AS PRODUCT_IMAGE", Padded))
        'SQLS.AppendLine("   FROM ICTSTYC1 C1")
        'SQLS.AppendLine("   LEFT JOIN ICTSTAT2 S2")
        'SQLS.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
        'SQLS.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
        'SQLS.AppendLine("   INNER JOIN ICTCOLR1 C2")
        'SQLS.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
        'SQLS.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
        'SQLS.AppendLine("  )")
        'SQLS.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') OR (MSOH <> 0) OR (MSFT <> 0))")
        'SQLS.AppendLine("  AND STYLE_CODE = :PARM1")
        SQLS.AppendLine("SELECT * FROM WBTSTYL2")
        SQLS.AppendLine("WHERE STYLE_CODE = :PARM1")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(SQLS.ToString(), "COLORCODES", "V", Style_Code)
        'For Each rowCOLOR_CODE As DataRow In tbl.Rows
        '    rowCOLOR_CODE.Item("PRODUCT_IMAGE") = String.Format("{0}-{1}.jpg", Style_Code, rowCOLOR_CODE.Item("COLOR_CODE"))
        'Next
        Return tbl
    End Function

    Private Function MakeStyleRow(StyleCode As String) As DataTable
        Dim sql As New Text.StringBuilder
        Dim Padded As String = StrDup(50, " ")
        sql.Length = 0
        sql.AppendLine(String.Format("SELECT ICTSTYLW.*, ICTSTYL1.STYLE_DESC,ICTCLAS1.STYLE_CLASS_DESC , WBTSTYL1.*, '{0}' AS GRAPHIC_NAME", Padded))
        sql.AppendLine("FROM ICTSTYLW, WBTSTYL1, ICTSTYL1, ICTCLAS1")
        sql.AppendLine("WHERE ICTSTYLW.STYLE_CODE = WBTSTYL1.STYLE_CODE")
        sql.AppendLine("AND WBTSTYL1.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        sql.AppendLine("AND ICTSTYL1.STYLE_CLASS_CODE = ICTCLAS1.STYLE_CLASS_CODE")
        sql.AppendLine("AND ICTSTYLW.STYLE_CODE = :PARM1")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", StyleCode)
        'tbl.Rows(0).Item("IMAGE_NAME") = GetImageName(StyleCode)
        Return tbl
    End Function

    Private Function FillQuantityPricing(ByVal STYLE_CODE As String) As XmlNode
        Dim Discounts As List(Of DISCOUNTS)
        Dim rowARTCUST1 As DataRow = Nothing
        Dim BASE As New ASFBASE0
        Dim MaxBreak As Integer = 4
        Dim nodeQuantityPricing As XmlNode
        Discounts = SOCMAIN2.Price_Discounts(BASE, "", rowARTCUST1, STYLE_CODE, False)
        nodeQuantityPricing = xmlLabelRequest.CreateElement("QuantityPricing")
        For i As Integer = 1 To 4
            If Discounts(i - 1).DISCOUNT_QTY = 0 Then
                MaxBreak = i - 1
                Exit For
            End If
        Next
        With nodeQuantityPricing
            MakeXMLNode(nodeQuantityPricing, "Enabled", "checked")
            MakeXMLNode(nodeQuantityPricing, "NumberPriceBreaks", MaxBreak)
            MakeXMLNode(nodeQuantityPricing, "Comment", BuildComments(BASE, STYLE_CODE))
            Dim nodeBackgroundColors As XmlNode = Nothing
            nodeBackgroundColors = xmlLabelRequest.CreateElement("BackgroundColors")
            With nodeBackgroundColors
                MakeXMLNode(nodeBackgroundColors, "QuantityColor", "#F6F6F6")
                MakeXMLNode(nodeBackgroundColors, "PriceAndCommentColor", "#FFFFFF")
                MakeXMLNode(nodeBackgroundColors, "OnSaleColor", "#F6F6F6")
            End With
            nodeQuantityPricing.AppendChild(nodeBackgroundColors)
            Dim nodePriceBreaks As XmlNode = Nothing
            nodePriceBreaks = xmlLabelRequest.CreateElement("PriceBreaks")
            With nodePriceBreaks
                Dim nodePriceBreak As XmlNode = Nothing
                For ictr As Integer = MaxBreak To 1 Step -1
                    nodePriceBreak = Nothing
                    nodePriceBreak = xmlLabelRequest.CreateElement("PriceBreak")
                    MakeXMLNode(nodePriceBreak, "StartingQuantity", Discounts(ictr - 1).DISCOUNT_QTY)
                    MakeXMLNode(nodePriceBreak, "UnitPrice", Format(Discounts(ictr - 1).DISCOUNT_PRICE, "###,###,##0.00"))
                    nodePriceBreaks.AppendChild(nodePriceBreak)
                Next
            End With
            nodeQuantityPricing.AppendChild(nodePriceBreaks)
        End With
        Return nodeQuantityPricing
    End Function

    Private Function BuildComments(frm As ASFBASE0, ByVal STYLE_CODE As String) As String
        Dim retVal As String = ""
        Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", STYLE_CODE)
        If Not IsNothing(rowICTSTYL1) Then
            retVal = retVal & "<link rel='stylesheet' type='text/css' href='http://www.regency-rib.com/css/regency.css'>"
            retVal = retVal & "<SPAN>"
            retVal = retVal & "<STRONG>UOM:</STRONG>" & rowICTSTYL1.Item("STYLE_UOM").ToString & " | "
            retVal = retVal & "<STRONG>BAG:</STRONG>" & rowICTSTYL1.Item("SUB_UNIT_PACK_QTY").ToString & " | "
            retVal = retVal & "<STRONG>BOX:</STRONG>" & rowICTSTYL1.Item("INNER_PACK_QTY").ToString & " | "
            retVal = retVal & "<STRONG>CART:</STRONG>" & rowICTSTYL1.Item("CARTON_PACK_QTY").ToString & ""
            retVal = retVal & "</SPAN>"
            If rowICTSTYL1.Item("STYLE_STATUS").ToString = "D" Then
                retVal = retVal & "<SPAN>"
                retVal = retVal & "<font color='blue'>* </font>"
                retVal = retVal & "<font color='red'>Item Is Discontinued</font>"
                retVal = retVal & "<font color='blue'> *</font>"
                retVal = retVal & "</SPAN>"
            End If
            'retVal = retVal & "<span> <input class='RGIBUTTON' type='button' onclick='history.back();' value='Continue Shopping'></span>"
            retVal = retVal & "<span> <input class='add' type='submit' onclick='history.back();' value='Continue Shopping'> </span>"
        End If
        Return retVal
    End Function

    Private Function GetCOLOR_GROUP_CODE(STYLE_CODE As String) As String
        Dim RetVal As String = ""
        Dim SQL As New System.Text.StringBuilder() With {.Length = 0}
        SQL.AppendLine("SELECT C1.COLOR_GROUP_CODE")
        SQL.AppendLine("FROM ICTSTYC1 S1, ICTCOLR1 C1")
        SQL.AppendLine("WHERE S1.COLOR_CODE = C1.COLOR_CODE")
        SQL.AppendLine("AND NVL(C1.COLOR_GROUP_CODE,'NULL') <> 'NULL'")
        SQL.AppendLine(String.Format("AND S1.STYLE_CODE = '{0}'", STYLE_CODE))
        SQL.AppendLine("GROUP BY C1.COLOR_GROUP_CODE")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), String.Empty)
        For Each rowICTCOLR1 As DataRow In tbl.Rows
            RetVal += rowICTCOLR1.Item("COLOR_GROUP_CODE").ToString.ToLower & ", "
            RetVal += rowICTCOLR1.Item("COLOR_GROUP_CODE").ToString.ToUpper & ", "
            RetVal += StrConv(rowICTCOLR1.Item("COLOR_GROUP_CODE").ToString.ToLower, VbStrConv.ProperCase) & ", "
        Next
        If RetVal.Length > 2 Then
            RetVal = RetVal.Substring(0, RetVal.Length - 2)
        Else
            RetVal = ""
        End If
        'Stop
        Return RetVal
    End Function

End Class
