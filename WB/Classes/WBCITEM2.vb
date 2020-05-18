Imports System.Xml
Imports System.Net.Mail
Imports System.String

Public Class WBCITEM2
    '-----------------------------------------------------
    '-This is The Newest Class That Produces XML Documents
    '-For ShopSite.  Needs To Be Swapped Out Before Going
    '-Live With New Site Being Designed By Kyle
    '-Scheduled For Fall of 2018.
    '-----------------------------------------------------
    Private xmlLabelRequest As XmlDocument
    Private productNodeCount As Int16 = 1
    Private nodeShopSite As XmlNode
    Private nodeProducts As XmlNode
    Private Const lt = "-&lt"
    Private Const gt = "-&gt"
    Private styleListInactive As List(Of String) = New List(Of String)
    Private dstWBTSTYLD As DataTable

#Region "Class Public Methods"
    Public Sub New(ByRef _dstWBTSTYLD As DataTable)
        Me.InitiailizeClass()
        dstWBTSTYLD = _dstWBTSTYLD
    End Sub

    Public Sub Clear()
        InitiailizeClass()
    End Sub

    Public Function AddStyle(ByVal StyleCode As String,
                             ByVal ColorCode As String,
                             ByVal DelList As List(Of String),
                             ByVal Optional UploadInventoryOnly As Boolean = True,
                             ByVal Optional isParent As Boolean = False,
                             ByVal Optional UpdatePricing As Boolean = False) As Integer
        styleListInactive = DelList
        Dim nodesProcessed As Integer = 0
        Dim ictr As Integer = 1

        Dim tbl As DataTable = GetStyleRow(StyleCode, ColorCode)
        For Each rowWBTSTYLD As DataRow In tbl.Rows
            Dim nodeProduct As XmlNode = Nothing

            If nodesProcessed = 0 Then
                nodeProduct = MakeProductNode(rowWBTSTYLD, True, tbl, UploadInventoryOnly, isParent, UpdatePricing)
                If nodeProduct IsNot Nothing Then
                    nodeProducts.AppendChild(nodeProduct)
                End If
            End If
            If nodeProduct IsNot Nothing Then
                nodeProducts.AppendChild(nodeProduct)
            End If
            nodesProcessed += 1
        Next
        Return nodesProcessed
    End Function

    Public ReadOnly Property GetXMLDocument() As XmlDocument
        Get
            Return xmlLabelRequest
        End Get
    End Property

    Public ReadOnly Property GetXmlOuterXml() As String
        Get
            Return xmlLabelRequest.OuterXml
        End Get
    End Property

    Public ReadOnly Property NumProductNodes() As Int16
        Get
            Return productNodeCount - 1
        End Get
    End Property
#End Region

#Region "Class Private Methods"
    Private Sub InitiailizeClass()
        productNodeCount = 1

        '"xml", "version=""1.0"" encoding=""UTF-8"""
        xmlLabelRequest = New XmlDocument
        xmlLabelRequest.PreserveWhitespace = True
        'xmlLabelRequest.AppendChild(xmlLabelRequest.CreateProcessingInstruction("xml", "version=""1.0""  encoding=""UTF-8"" standalone=""no"""))
        xmlLabelRequest.AppendChild(xmlLabelRequest.CreateProcessingInstruction("xml", "version=""1.0"" encoding=""iso-8859-1"""))

        nodeShopSite = Nothing
        nodeShopSite = xmlLabelRequest.CreateElement("ShopSiteProducts")

        xmlLabelRequest.AppendChild(nodeShopSite)

        nodeProducts = Nothing
        nodeProducts = xmlLabelRequest.CreateElement("Products")

        nodeShopSite.AppendChild(nodeProducts)

    End Sub
#End Region

#Region "XMLNodeCreation"
    Private Function MakeCrossSellNode(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As XmlNode
        Dim SQL As New Text.StringBuilder() With {.Length = 0}
        Dim nodeCrossSellNode As XmlNode = xmlLabelRequest.CreateElement("CrossSell")

        SQL.AppendLine("Select")
        SQL.AppendLine("WD.STYLE_CODE, ")
        SQL.AppendLine("WD.COLOR_CODE, ")
        SQL.AppendLine("WH.STYLE_DESC_SHORT, ")
        SQL.AppendLine("S1.STYLE_DESC")
        SQL.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1")
        SQL.AppendLine("WHERE WH.STYLE_CODE = WD.STYLE_CODE")
        SQL.AppendLine("And WH.STYLE_CODE = S1.STYLE_CODE")
        'SQL.AppendLine("And WEB_IND IN ('U','W')")
        SQL.AppendLine("And CURR_ON_HAND > 0")
        SQL.AppendLine(Format("AND WD.STYLE_CODE = '{0}'", STYLE_CODE))
        SQL.AppendLine(Format("AND WD.COLOR_CODE <> '{0}'", COLOR_CODE))
        SQL.AppendLine("ORDER BY WD.COLOR_CODE")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), String.Empty)
        For Each rowWBTSTYLD As DataRow In tbl.Rows
            Dim COLOR_DESC As String = GetCOLOR_DESC(rowWBTSTYLD.Item("COLOR_CODE").ToString)
            Dim nodeCrossSellItem As XmlNode = xmlLabelRequest.CreateElement("CrossSellItem")
            Dim nodeName As XmlNode = Nothing
            nodeName = xmlLabelRequest.CreateElement("Name")
            nodeName.InnerText = String.Format("{0} {1}", GetSTYLE_DESC_SHORT(rowWBTSTYLD), COLOR_DESC)
            'nodeName.InnerText = rowWBTSTYLD.Item("STYLE_DESC_SHORT") & " (" & rowWBTSTYLD.Item("COLOR_CODE") & ")"
            'nodeName.InnerText = String.Format("{0} {1}", rowWBTSTYLD.Item("STYLE_DESC_SHORT"), COLOR_DESC)
            nodeCrossSellItem.AppendChild(nodeName)

            Dim nodeSKU As XmlNode = Nothing
            nodeSKU = xmlLabelRequest.CreateElement("SKU")
            'nodeSKU.InnerText = rowWBTSTYLD.Item("STYLE_CODE") & "-" & rowWBTSTYLD.Item("COLOR_CODE")
            nodeSKU.InnerText = String.Format("{0}-{1}", rowWBTSTYLD.Item("STYLE_CODE"), rowWBTSTYLD.Item("COLOR_CODE").ToString)
            nodeCrossSellItem.AppendChild(nodeSKU)
            nodeCrossSellNode.AppendChild(nodeCrossSellItem)
        Next
        Return nodeCrossSellNode
    End Function

    Private Function MakeProductNode(ByVal rowWBTSTYLD As DataRow,
                                     ByVal mainStyle As Boolean,
                                     ByRef tblSubProducts As DataTable,
                                     ByVal Optional UploadInventoryOnly As Boolean = True,
                                     ByVal Optional isParent As Boolean = False,
                                     ByVal Optional UpdatePricing As Boolean = False) As XmlNode
        Dim nodeProduct As XmlNode = Nothing
        Try
            If rowWBTSTYLD Is Nothing Then
                Return Nothing
            End If

            nodeProduct = Nothing
            nodeProduct = xmlLabelRequest.CreateElement("Product")

            Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE") & String.Empty
            Dim COLOR_CODE As String = rowWBTSTYLD.Item("COLOR_CODE") & String.Empty
            Dim COLOR_DESC As String = GetCOLOR_DESC(COLOR_CODE)
            Dim ATTRIBUTES As String = GetAttributes(STYLE_CODE)
            Dim tblCOLORS As DataTable = RefreshColorTable(STYLE_CODE)

            Dim FTR_AVAIL As Double = Val(rowWBTSTYLD.Item("FTR_AVAIL").ToString & String.Empty)
            If FTR_AVAIL <= 0 Then
                FTR_AVAIL = 0
            End If

            MakeXMLNode(nodeProduct, "Name", GetWEB_DESC(rowWBTSTYLD, isParent))
            If isParent Then
                MakeXMLNode(nodeProduct, "SKU", STYLE_CODE)
                MakeXMLNode(nodeProduct, "QuantityOnHand", 0)
            Else
                MakeXMLNode(nodeProduct, "SKU", STYLE_CODE & "-" & COLOR_CODE)
                MakeXMLNode(nodeProduct, "QuantityOnHand", FTR_AVAIL)
            End If
            MakeXMLNode(nodeProduct, "ProductDisabled", GetProductDisabled(STYLE_CODE, COLOR_CODE, isParent))
            If Not UploadInventoryOnly Then
                MakeXMLNode(nodeProduct, "ProductDescription", GetSTYLE_DESC_LONG(STYLE_CODE))
                MakeXMLNode(nodeProduct, "MinimumQuantity", GetMinimumQuantity(STYLE_CODE))
                nodeProduct.AppendChild(MakeSubproductsNode(STYLE_CODE, COLOR_CODE, isParent))
                nodeProduct.AppendChild(MakeOptionMenusNode(tblCOLORS))
                nodeProduct.AppendChild(MakeProductOptionsNode(tblCOLORS))
                If isParent Then
                    MakeXMLNode(nodeProduct, "FileName", STYLE_CODE & ".html")
                Else
                    MakeXMLNode(nodeProduct, "FileName", STYLE_CODE & "-" & COLOR_CODE & ".html")
                End If
                nodeProduct.AppendChild(MakeQuantityPricingNode(STYLE_CODE))
                MakeXMLNode(nodeProduct, "SearchKeywords", ATTRIBUTES)
                MakeXMLNode(nodeProduct, "Price", "0.00")
                MakeXMLNode(nodeProduct, "SaleAmount")
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
                MakeXMLNode(nodeProduct, "DobaItemID")
                MakeXMLNode(nodeProduct, "VariablePrice", "checked")
                MakeXMLNode(nodeProduct, "VariableName", "uncheck")
                MakeXMLNode(nodeProduct, "VariableSKU", "uncheck")
                MakeXMLNode(nodeProduct, "VariableWeight", "uncheck")
                MakeXMLNode(nodeProduct, "Taxable", "checked")
                MakeXMLNode(nodeProduct, "AvaTaxCode")
                MakeXMLNode(nodeProduct, "VAT", "0")
                MakeXMLNode(nodeProduct, "Graphic", "product/" & STYLE_CODE & "-" & COLOR_CODE & ".jpg")
                MakeXMLNode(nodeProduct, "ProductImageDesc")
                MakeXMLNode(nodeProduct, "ProductImageSize", "2")
                MakeXMLNode(nodeProduct, "SearchDestType", "selected")
                MakeXMLNode(nodeProduct, "SearchDest", "Store")
                MakeXMLNode(nodeProduct, "SearchMakePage")
                MakeXMLNode(nodeProduct, "MerchantProductInstructions")
                MakeXMLNode(nodeProduct, "Video")
                MakeXMLNode(nodeProduct, "Returns")
                MakeXMLNode(nodeProduct, "Warranty")
                MakeXMLNode(nodeProduct, "Specifications")
                MakeXMLNode(nodeProduct, "MaterialsOrIngredients")
                MakeXMLNode(nodeProduct, "HowToUse")
                MakeXMLNode(nodeProduct, "ShippingDetails")
                MakeXMLNode(nodeProduct, "SizeAndFitGuide", "I")
                MakeXMLNode(nodeProduct, "SizeAndFitImage", "none")
                MakeXMLNode(nodeProduct, "SizeAndFitText")
                MakeXMLNode(nodeProduct, "SizeAndFitImageDesc")
                MakeXMLNode(nodeProduct, "AsSeenInImage", "none")
                MakeXMLNode(nodeProduct, "AsSeenInImageDesc")
                MakeXMLNode(nodeProduct, "QuantityPricingGroup")
                MakeXMLNode(nodeProduct, "Weight", "0")
                MakeXMLNode(nodeProduct, "ItemSize")
                MakeXMLNode(nodeProduct, "DimensionOptions", "0")
                MakeXMLNode(nodeProduct, "DimensionText")
                MakeXMLNode(nodeProduct, "DimensionSelected")
                MakeXMLNode(nodeProduct, "FedExContainer")
                MakeXMLNode(nodeProduct, "USPSContainer")
                MakeXMLNode(nodeProduct, "CanadaPostContainer")
                MakeXMLNode(nodeProduct, "AustraliaPostContainer")
                MakeXMLNode(nodeProduct, "NoShippingCharges", "uncheck")
                MakeXMLNode(nodeProduct, "ExtraHandlingCharge", "0.00")
                MakeXMLNode(nodeProduct, "ProhibitedShippingMethods")
                MakeXMLNode(nodeProduct, "ProductType", "Tangible")
                MakeXMLNode(nodeProduct, "ProductDownloadLocation", "none")
                MakeXMLNode(nodeProduct, "GroundShipping", "0.00")
                MakeXMLNode(nodeProduct, "SecondDayShipping", "0.00")
                MakeXMLNode(nodeProduct, "NextDayShipping", "0.00")
                MakeXMLNode(nodeProduct, "Shipping3", "0.00")
                MakeXMLNode(nodeProduct, "Shipping4", "0.00")
                MakeXMLNode(nodeProduct, "Shipping5", "0.00")
                MakeXMLNode(nodeProduct, "Shipping6", "0.00")
                MakeXMLNode(nodeProduct, "Shipping7", "0.00")
                MakeXMLNode(nodeProduct, "Shipping8", "0.00")
                MakeXMLNode(nodeProduct, "Shipping9", "0.00")
                MakeXMLNode(nodeProduct, "LowStockThreshold", "0")
                MakeXMLNode(nodeProduct, "OutOfStockLimit", "0")
                MakeXMLNode(nodeProduct, "OptionText")
                MakeXMLNode(nodeProduct, "CustomerTextEntryBox", "uncheck")
                MakeXMLNode(nodeProduct, "CustomerTextEntryHeader")
                MakeXMLNode(nodeProduct, "CustomerTextEntryColumns", "40")
                MakeXMLNode(nodeProduct, "CustomerTextEntryRows", "4")
                MakeXMLNode(nodeProduct, "OptionColumnHeaders", "Color")
                MakeXMLNode(nodeProduct, "OptionAppendSKU")
                MakeXMLNode(nodeProduct, "OptionUseMultiMenus")
                MakeXMLNode(nodeProduct, "OptionSelectDefault")
                MakeXMLNode(nodeProduct, "GoogleBase", "uncheck")
                MakeXMLNode(nodeProduct, "Brand")
                MakeXMLNode(nodeProduct, "GTIN")
                MakeXMLNode(nodeProduct, "ManufacturerPartNumber")
                MakeXMLNode(nodeProduct, "GoogleProductType")
                MakeXMLNode(nodeProduct, "GoogleProductCategory")
                MakeXMLNode(nodeProduct, "GoogleCustomProduct", "uncheck")
                MakeXMLNode(nodeProduct, "Availability", "in stock")
                MakeXMLNode(nodeProduct, "GoogleCondition", "New")
                MakeXMLNode(nodeProduct, "GoogleAgeGroup", "none")
                MakeXMLNode(nodeProduct, "GoogleGender", "none")
                MakeXMLNode(nodeProduct, "GoogleColorColumn")
                MakeXMLNode(nodeProduct, "GoogleSizeColumn")
                MakeXMLNode(nodeProduct, "GooglePatternColumn")
                MakeXMLNode(nodeProduct, "GoogleMaterialColumn")
                MakeXMLNode(nodeProduct, "GoogleUseAdvancedOrderingOptions", "uncheck")
                MakeXMLNode(nodeProduct, "GoogleListAsFreeShipping", "uncheck")
                MakeXMLNode(nodeProduct, "CrossSell")
                nodeProduct.AppendChild(MakeProductOnPagesNode(rowWBTSTYLD, isParent))
                MakeXMLNode(nodeProduct, "AddToPages")
                If isParent Then
                    MakeXMLNode(nodeProduct, "DisplayMoreInformationPage", "checked")
                Else
                    MakeXMLNode(nodeProduct, "DisplayMoreInformationPage", "uncheck")
                End If

                If isParent Then
                    MakeXMLNode(nodeProduct, "MoreInfoTitle", STYLE_CODE)
                Else
                    MakeXMLNode(nodeProduct, "MoreInfoTitle", STYLE_CODE & "-" & COLOR_CODE)
                End If
                MakeXMLNode(nodeProduct, "MoreInformationText", MakeMoreInformationTextNode(STYLE_CODE))
                MakeXMLNode(nodeProduct, "MoreInformationGraphic", "product/" & STYLE_CODE & "-" & COLOR_CODE & ".jpg")
                MakeXMLNode(nodeProduct, "MoreInfoImageSize", "1")
                MakeXMLNode(nodeProduct, "MoreInfoMetaKeywords")
                MakeXMLNode(nodeProduct, "MoreInfoMetaDescription")
                MakeXMLNode(nodeProduct, "OneLineAdvertisement")
                MakeXMLNode(nodeProduct, "ProductSitemap", "checked")
                MakeXMLNode(nodeProduct, "ProductSitemapPriority", "Google Default")
                MakeXMLNode(nodeProduct, "ProductCrossSell", "checked")
                MakeXMLNode(nodeProduct, "GlobalCrossSell", "uncheck")
                MakeMoreInfoImages(nodeProduct, STYLE_CODE, isParent)
                MakeXMLNode(nodeProduct, "MoreInfoImageExtraSize", "3")
                MakeXMLNode(nodeProduct, "MoreInformationImageDesc")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc1")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc2")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc3")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc4")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc5")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc6")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc7")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc8")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc9")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc10")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc11")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc12")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc13")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc14")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc15")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc16")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc17")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc18")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc19")
                MakeXMLNode(nodeProduct, "MoreInfoImageDesc20")
                MakeXMLNode(nodeProduct, "Template", "RI-Product.sst")
                MakeXMLNode(nodeProduct, "DisplayName", "checked")
                MakeXMLNode(nodeProduct, "DisplaySKU", "checked")
                MakeXMLNode(nodeProduct, "DisplayPrice", "uncheck")
                MakeXMLNode(nodeProduct, "SaleOn", "uncheck")
                MakeXMLNode(nodeProduct, "DisplayGraphic", "checked")
                MakeXMLNode(nodeProduct, "DisplayOrderQuantity", "checked")
                MakeXMLNode(nodeProduct, "DisplayQuantityPricing", "checked")
                MakeXMLNode(nodeProduct, "DisplayOrderingOptions", "uncheck")
                MakeXMLNode(nodeProduct, "DisplayAddToCart", "All Pages")
                MakeXMLNode(nodeProduct, "NameStyle", "Bold")
                MakeXMLNode(nodeProduct, "NameSize", "Normal")
                MakeXMLNode(nodeProduct, "PriceStyle", "Bold")
                MakeXMLNode(nodeProduct, "PriceSize", "Normal")
                MakeXMLNode(nodeProduct, "SKUStyle", "Plain")
                MakeXMLNode(nodeProduct, "SKUSize", "Small")
                MakeXMLNode(nodeProduct, "DescriptionStyle", "Plain")
                MakeXMLNode(nodeProduct, "DescriptionSize", "Normal")
                MakeXMLNode(nodeProduct, "ImageAlignment", "Left")
                MakeXMLNode(nodeProduct, "TextWrap", "Off")
                MakeXMLNode(nodeProduct, "AddtoCartButton", "Add To Cart")
                MakeXMLNode(nodeProduct, "ViewCartButton", "View Cart")
                MakeXMLNode(nodeProduct, "UseAddtoCartImage", "0")
                MakeXMLNode(nodeProduct, "AddtoCartImage", "[shopsite-images]/buttons/defaults/Add_To_Cart.gif")
                MakeXMLNode(nodeProduct, "AddtoCartImageDesc")
                MakeXMLNode(nodeProduct, "UseViewCartImage", "0")
                MakeXMLNode(nodeProduct, "ViewCartImage", "[shopsite-images]/buttons/defaults/View_Cart.gif")
                MakeXMLNode(nodeProduct, "ViewCartImageDesc")
                MakeXMLNode(nodeProduct, "QBImport")
            End If
            MakeProductFieldNodes(nodeProduct, STYLE_CODE, COLOR_CODE, isParent, UploadInventoryOnly, UpdatePricing)
            '<ProductID>12</ProductID> 'Make Sure We Don't Have to set this.
            'MakeXMLNode(nodeProduct, "BlankEntry")
            productNodeCount += 1
        Catch ex As Exception
            nodeProduct = Nothing
        End Try

        Return nodeProduct
    End Function

    Private Function MakeOptionMenusNode(ByVal tblCOLORS As DataTable) As XmlNode
        Dim nodeOptionMenus As XmlNode = Nothing
        nodeOptionMenus = xmlLabelRequest.CreateElement("OptionMenus")
        Dim nodeMenu As XmlNode = Nothing
        nodeMenu = xmlLabelRequest.CreateElement("Menu")
        MakeXMLNode(nodeMenu, "MenuItem", "Color;n")
        For Each rowTbl As DataRow In tblCOLORS.Rows
            MakeXMLNode(nodeMenu, "MenuItem", rowTbl.Item("COLOR_CODE").ToString & String.Empty)
        Next
        nodeOptionMenus.AppendChild(nodeMenu)
        Return nodeOptionMenus
    End Function

    Private Function MakeProductOptionsNode(tblCOLORS As DataTable) As XmlNode
        Dim nodeProductOptions As XmlNode = Nothing
        nodeProductOptions = xmlLabelRequest.CreateElement("ProductOptions")
        With nodeProductOptions
            For Each rowCOLOR As DataRow In tblCOLORS.Select()
                'If rowCOLOR.Item("STYLE_STATUS").ToString = "A" Or (Val(rowCOLOR.Item("CURR_ON_HAND").ToString) > 0) Then
                If rowCOLOR.Item("STYLE_STATUS").ToString = "A" Then
                    Dim nodeProductOption As XmlNode = Nothing
                    nodeProductOption = xmlLabelRequest.CreateElement("ProductOption")
                    Dim typeAttr As XmlAttribute = xmlLabelRequest.CreateAttribute("Name")
                    typeAttr.Value = rowCOLOR.Item("COLOR_CODE")
                    nodeProductOption.Attributes.Append(typeAttr)
                    With nodeProductOption
                        MakeXMLNode(nodeProductOption, "Use")
                        MakeXMLNode(nodeProductOption, "Menu1", rowCOLOR.Item("COLOR_CODE"))
                        MakeXMLNode(nodeProductOption, "Menu2")
                        MakeXMLNode(nodeProductOption, "Menu3")
                        MakeXMLNode(nodeProductOption, "Menu4")
                        MakeXMLNode(nodeProductOption, "AppendText")
                        MakeXMLNode(nodeProductOption, "SKU")
                        MakeXMLNode(nodeProductOption, "PriceModifier")
                        MakeXMLNode(nodeProductOption, "WeightModifier")
                        MakeXMLNode(nodeProductOption, "QuantityOnHand")
                        MakeXMLNode(nodeProductOption, "LowStockThreshold")
                        MakeXMLNode(nodeProductOption, "OutOfStockLimit")
                        Dim ImageFile As String = rowCOLOR.Item("DEFAULT_IMAGE")
                        With ImageFile
                            If .Length > 3 Then
                                If .Substring(.Length - 3, 3) = "JPG" Then
                                    ImageFile = .Substring(0, .Length - 3) & "jpg"
                                End If
                            End If
                        End With
                        MakeXMLNode(nodeProductOption, "Image", "product/" & ImageFile)
                        MakeXMLNode(nodeProductOption, "GTIN")
                        MakeXMLNode(nodeProductOption, "QBImport")
                        MakeXMLNode(nodeProductOption, "ExtraField1")
                        MakeXMLNode(nodeProductOption, "ExtraField2")
                    End With
                    .AppendChild(nodeProductOption)
                End If
            Next
        End With
        Return nodeProductOptions
    End Function

    Private Sub MakeProductFieldNodes(ByRef nodeProduct As XmlNode,
                                      ByVal STYLE_CODE As String,
                                      ByVal COLOR_CODE As String,
                                      ByVal isParent As Boolean,
                                      ByVal UploadInventoryOnly As Boolean,
                                      Optional ByVal UpdatePricing As Boolean = False)
        Dim BASE As New ASFBASE0
        Dim rowICTSTYL1 As DataRow = BASE.LookUp("ICTSTYL1", STYLE_CODE)
        For i As Integer = 1 To 20
            Select Case i
                Case 1
                    If Not UploadInventoryOnly Then
                        MakeXMLNode(nodeProduct, "ProductField1", GetMinimumQuantity(STYLE_CODE))
                    End If
                Case 2
                    MakeXMLNode(nodeProduct, "ProductField2", GetNextDelDate(STYLE_CODE, COLOR_CODE))
                Case 4
                    'MakeXMLNode(nodeProduct, "ProductField4", STYLE_CODE & "-" & COLOR_CODE)
                Case 5
                    If Not UploadInventoryOnly Then
                        MakeXMLNode(nodeProduct, "ProductField5", GetColorsAvail(STYLE_CODE))
                    End If
                Case 6
                    'If IsNumeric(rowICTSTYL1.Item("INNER_PACK_QTY").ToString & String.Empty) Then
                    '    Dim INNER_PACK_QTY As Integer = Val(rowICTSTYL1.Item("INNER_PACK_QTY").ToString & String.Empty)
                    '    If INNER_PACK_QTY > 0 Then
                    '        MakeXMLNode(nodeProduct, "ProductField6", INNER_PACK_QTY)
                    '    Else
                    '        MakeXMLNode(nodeProduct, "ProductField6")
                    '    End If
                    'Else
                    '    MakeXMLNode(nodeProduct, "ProductField6")
                    'End If
                Case 7
                    If Not UploadInventoryOnly Then
                        If isParent Then
                            MakeXMLNode(nodeProduct, "ProductField7")
                        Else
                            MakeXMLNode(nodeProduct, "ProductField7", GetCOLOR_DESC(COLOR_CODE))
                        End If
                    End If
                Case 10
                    If Not UploadInventoryOnly Then
                        Dim STYLE_UOM As String = rowICTSTYL1.Item("STYLE_UOM").ToString & String.Empty
                        MakeXMLNode(nodeProduct, "ProductField10", STYLE_UOM)
                    End If

                Case 11
                    If Not UploadInventoryOnly Then
                        'Per Danny If Inner is blank or zero use STYLE_SO_QTY_MIN if > 0 - 1/30/19
                        If IsNumeric(rowICTSTYL1.Item("INNER_PACK_QTY").ToString & String.Empty) Then
                            Dim INNER_PACK_QTY As Integer = Val(rowICTSTYL1.Item("INNER_PACK_QTY").ToString & String.Empty)
                            Dim STYLE_SO_QTY_MIN As Integer = Val(rowICTSTYL1.Item("STYLE_SO_QTY_MIN").ToString & String.Empty)
                            If INNER_PACK_QTY > 0 Then
                                MakeXMLNode(nodeProduct, "ProductField11", INNER_PACK_QTY)
                            Else
                                If STYLE_SO_QTY_MIN > 0 Then
                                    MakeXMLNode(nodeProduct, "ProductField11", STYLE_SO_QTY_MIN)
                                Else
                                    MakeXMLNode(nodeProduct, "ProductField11", 1)
                                End If
                            End If
                        Else
                            MakeXMLNode(nodeProduct, "ProductField11", 1)
                        End If
                    End If
                Case 12
                    If Not UploadInventoryOnly Then
                        If IsNumeric(rowICTSTYL1.Item("CARTON_PACK_QTY").ToString & String.Empty) Then
                            Dim CARTON_PACK_QTY As Integer = Val(rowICTSTYL1.Item("CARTON_PACK_QTY").ToString & String.Empty)
                            If CARTON_PACK_QTY > 0 Then
                                MakeXMLNode(nodeProduct, "ProductField12", CARTON_PACK_QTY)
                            Else
                                MakeXMLNode(nodeProduct, "ProductField12", 1)
                            End If
                        Else
                            MakeXMLNode(nodeProduct, "ProductField12", 1)
                        End If
                    End If
                Case 18
                    If Not UploadInventoryOnly Then
                        MakeXMLNode(nodeProduct, "ProductField18", GetColorGroups(STYLE_CODE))
                    End If
                Case 19
                    If Not UploadInventoryOnly Then
                        MakeXMLNode(nodeProduct, "ProductField19", GetSizeGroup(STYLE_CODE))
                    End If
                Case 20
                    If Not UploadInventoryOnly Then
                        MakeXMLNode(nodeProduct, "ProductField20", GetCategoryGroup(STYLE_CODE))
                    End If
                Case Else
                    If Not UploadInventoryOnly Then
                        MakeXMLNode(nodeProduct, "ProductField" & i.ToString)
                    End If
            End Select
        Next
        If UploadInventoryOnly = False Or UpdatePricing = True Then
            Dim rowARTCUST1 As DataRow = BASE.LookUp("ARTCUST1", "180000")

            Dim rowICTCLAS1 As DataRow = BASE.LookUp("ICTCLAS1", rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString & String.Empty)
            Dim IsPVC As Boolean = rowICTCLAS1.Item("DISC_CODE").ToString = "PVC"
            Dim isDiscontunued As Boolean = rowICTSTYL1.Item("STYLE_STATUS").ToString = "D"
            If isDiscontunued Then
                Dim rowNOTHING As DataRow = Nothing
                Dim Discounts As List(Of DISCOUNTS)
                Discounts = SOCMAIN2.Price_Discounts(BASE, "", rowNOTHING, STYLE_CODE, False, False)
                'Stop
                MakeXMLNode(nodeProduct, "ProductField9", Discounts(0).DISCOUNT_PRICE)
            End If
            For i As Integer = 28 To 32
                If IsPVC Or isDiscontunued Then
                    MakeXMLNode(nodeProduct, "ProductField" & i.ToString)
                Else
                    Dim CUST_PRICE_TIER As String = "PC"
                    Dim CUST_DISC_PCT_EXTRA As String = "0"
                    Dim CUST_DISC_PCT As Integer = 0
                    Select Case i
                        Case 28
                            CUST_PRICE_TIER = "PC"
                            CUST_DISC_PCT_EXTRA = "1"
                        Case 29
                            CUST_PRICE_TIER = "PC"
                            CUST_DISC_PCT_EXTRA = "2"
                        Case 30
                            CUST_PRICE_TIER = "HC"
                        Case 31
                            CUST_PRICE_TIER = "FC"
                        Case 32
                            CUST_PRICE_TIER = "SP"
                            CUST_DISC_PCT = 54
                        Case 33
                        Case 34
                            'This was reversed by Danny.  Anyone SP is always 54, Otherwise default pricing.
                            'CUST_PRICE_TIER = "SP"
                            'CUST_DISC_PCT = 54
                    End Select
                    rowARTCUST1.Item("CUST_PRICE_TIER") = CUST_PRICE_TIER
                    rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = CUST_DISC_PCT_EXTRA
                    rowARTCUST1.Item("CUST_DISC_PCT") = CUST_DISC_PCT
                    'MakeXMLNode(nodeProduct, "ProductField" & i.ToString)
                    Dim Discounts As List(Of DISCOUNTS)
                    Dim MaxBreak As Integer = 4
                    'Dim nodeQuantityPricing As XmlNode
                    Discounts = SOCMAIN2.Price_Discounts(BASE, "", rowARTCUST1, STYLE_CODE, True, False)
                    For z As Integer = 1 To 4
                        If Discounts(z - 1).DISCOUNT_QTY = 0 Then
                            MaxBreak = z - 1
                            Exit For
                        End If
                    Next
                    Dim DISC_DESC As New Text.StringBuilder With {.Length = 0}
                    For ictr As Integer = MaxBreak - 1 To 0 Step -1
                        If ictr = 0 Then
                            'DISC_DESC.AppendLine(String.Format("{0}", Math.Round(Discounts(ictr).DISCOUNT_PRICE, 2)))
                            DISC_DESC.Append(String.Format("{0}", Math.Round(Discounts(ictr).DISCOUNT_PRICE, 2)))
                        Else
                            If CUST_PRICE_TIER = "HC" And ictr = 3 Then
                                DISC_DESC.AppendLine(String.Format("{0}|{1}", Discounts(ictr - 1).DISCOUNT_QTY - 1, Math.Round(Discounts(ictr - 1).DISCOUNT_PRICE, 2)))
                            Else
                                If CUST_PRICE_TIER = "FC" And ictr <= 3 Then
                                    DISC_DESC.AppendLine(String.Format("{0}|{1}", Discounts(ictr - 1).DISCOUNT_QTY - 1, Math.Round(Discounts(1).DISCOUNT_PRICE, 2)))
                                Else
                                    DISC_DESC.AppendLine(String.Format("{0}|{1}", Discounts(ictr - 1).DISCOUNT_QTY - 1, Math.Round(Discounts(ictr).DISCOUNT_PRICE, 2)))
                                End If
                            End If
                        End If
                    Next
                    MakeXMLNode(nodeProduct, "ProductField" & i.ToString, DISC_DESC.ToString)
                End If
            Next
        End If
    End Sub

    Private Function GetCategoryGroup(ByVal STYLE_CODE As String) As String
        Dim CAT_DESC As New Text.StringBuilder With {.Length = 0}

        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        Sql.AppendLine("DISTINCT A1.ATTR_DESC")
        Sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1, ICTSTYL3 S3, ICTATTR1 A1")
        sql.AppendLine("WHERE WH.STYLE_CODE (+) = WD.STYLE_CODE")
        sql.AppendLine("AND WD.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND S1.STYLE_CODE = S3.STYLE_CODE")
        Sql.AppendLine("AND S3.ATTR_CODE = A1.ATTR_CODE")
        sql.AppendLine(String.Format("AND WD.STYLE_CODE = '{0}'", STYLE_CODE))
        sql.AppendLine("AND NVL(A1.ATT_RANK,0) = 1")
        'GO_LIVE_CHANGES
        'sql.AppendLine("AND WD.CURR_ON_HAND > 0")
        Dim tblICTATTR1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())

        For Each rowICTATTR1 As DataRow In tblICTATTR1.Rows
            If CAT_DESC.Length = 0 Then
                CAT_DESC.Append(rowICTATTR1.Item("ATTR_DESC").ToString & String.Empty)
            Else
                CAT_DESC.Append("|" & rowICTATTR1.Item("ATTR_DESC").ToString & String.Empty)
            End If
        Next

        Return CAT_DESC.ToString
    End Function

    Private Function GetSizeGroup(ByVal STYLE_CODE As String) As String
        Dim SIZE_DESC As New Text.StringBuilder With {.Length = 0}

        Dim sql As New Text.StringBuilder With {.Length = 0}

        sql.AppendLine("SELECT")
        sql.AppendLine("DISTINCT Z1.SIZE_CODE")
        sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1, ICTSIZE1 Z1")
        sql.AppendLine("WHERE WH.STYLE_CODE = WD.STYLE_CODE")
        sql.AppendLine("AND WH.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND S1.SIZE_CODE = Z1.SIZE_CODE")
        sql.AppendLine(String.Format("AND WD.STYLE_CODE = '{0}'", STYLE_CODE))
        sql.AppendLine("AND S1.STYLE_CLASS_CODE = 'PVC'")
        'GO_LIVE_CHANGES
        'sql.AppendLine("AND WD.CURR_ON_HAND > 0")
        Dim tblICTSIZE1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())

        For Each rowICTCOLR1 As DataRow In tblICTSIZE1.Rows
            If SIZE_DESC.Length = 0 Then
                SIZE_DESC.Append(rowICTCOLR1.Item("SIZE_CODE").ToString & String.Empty)
            Else
                SIZE_DESC.Append("|" & rowICTCOLR1.Item("SIZE_CODE").ToString & String.Empty)
            End If
        Next

        Return SIZE_DESC.ToString
    End Function

    Private Function GetColorGroups(ByVal STYLE_CODE As String) As String
        Dim COLOR_DESC As New Text.StringBuilder With {.Length = 0}

        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("C1.COLOR_DESC")
        sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1, ICTCOLR1 C1")
        sql.AppendLine("WHERE WH.STYLE_CODE (+) = WD.STYLE_CODE")
        sql.AppendLine("AND WD.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND WD.COLOR_CODE = C1.COLOR_CODE")
        sql.AppendLine(String.Format("AND WD.STYLE_CODE = '{0}'", STYLE_CODE))
        'GO_LIVE_CHANGES
        'sql.AppendLine("AND WD.CURR_ON_HAND > 0")
        Dim tblICTCOLR1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())

        For Each rowICTCOLR1 As DataRow In tblICTCOLR1.Rows
            If COLOR_DESC.Length = 0 Then
                COLOR_DESC.Append(rowICTCOLR1.Item("COLOR_DESC").ToString & String.Empty)
            Else
                COLOR_DESC.Append("|" & rowICTCOLR1.Item("COLOR_DESC").ToString & String.Empty)
            End If
        Next

        Return COLOR_DESC.ToString
    End Function

    Private Function GetColorsAvail(ByVal STYLE_CODE As String) As String
        Dim Retval As String = ""

        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT")
        SQLS.AppendLine("COUNT(DISTINCT WD.COLOR_CODE) AS CLR_CNT")
        SQLS.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1")
        SQLS.AppendLine("WHERE WH.STYLE_CODE = WD.STYLE_CODE")
        SQLS.AppendLine("AND WH.STYLE_CODE = S1.STYLE_CODE")
        SQLS.AppendLine(Format("AND WD.STYLE_CODE = '{0}'", STYLE_CODE))
        'GO_LIVE_CHANGES
        'SQLS.AppendLine("AND WD.CURR_ON_HAND > 0")
        ASCMAIN1.sql = SQLS.ToString()
        Dim CLR_CNT As Int16 = Val(ASCDATA1.GetDataValue)
        If CLR_CNT > 1 Then
            Retval = String.Format("{0} Colors Available", CLR_CNT)
        End If

        Return Retval
    End Function

    Private Function GetNextDelDate(STYLE_CODE As String, COLOR_CODE As String) As String
        Dim retVal As String = ""

        'GO_LIVE_CHANGES
        Dim tbl As DataTable = GetStyleRow(STYLE_CODE, COLOR_CODE)
        Dim STYLE_STATUS As String = "A"
        Dim STYLE_COLOR_STATUS As String = "A"
        Dim FTR_AVAIL As Int64 = 0
        If tbl.Rows.Count = 1 Then
            Dim rowWBTSTYLD As DataRow = tbl.Rows(0)
            STYLE_STATUS = rowWBTSTYLD.Item("STYLE_STATUS").ToString & ""
            STYLE_COLOR_STATUS = rowWBTSTYLD.Item("STYLE_COLOR_STATUS").ToString & ""
            FTR_AVAIL = Val(rowWBTSTYLD.Item("FTR_AVAIL").ToString & "")
            If FTR_AVAIL < 0 Then
                FTR_AVAIL = 0
            End If
        End If
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("SELECT MAX(STATUS_DATE) AS NEXT_DATE")
        SQLS.AppendLine("FROM ICTSTDQ1")
        SQLS.AppendLine("WHERE WHSE_CODE = 'MS'")
        SQLS.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim NEXT_DATE As String = ASCDATA1.GetDataValue
        If Not IsNothing(NEXT_DATE) Then
            If IsDate(NEXT_DATE) Then
                If NEXT_DATE > Now.AddDays(1) Then
                    retVal = Format(CDate(NEXT_DATE), "MM/dd/yy").ToString
                End If
            End If
        End If

        'GO_LIVE_CHANGES
        'If retVal = "" And STYLE_COLOR_STATUS = "A" And FTR_AVAIL = 0 Then
        '    retVal = DateSerial(2019, 8, 1)
        'End If
        Return retVal
    End Function

    Private Function MakeMoreInformationTextNode(ByVal STYLE_CODE As String) As String
        'Dim InjectScripts As New System.Text.StringBuilder
        'InjectScripts.Length = 0
        'InjectScripts.AppendLine("<script src='js/productpage.js'></script>")
        'InjectScripts.AppendLine("<div id='injectinv'></div>")
        'InjectScripts.AppendLine("<script src='http://ajax.googleapis.com/ajax/libs/jquery/1.8.0/jquery.min.js' type='text/javascript'></script>")
        'InjectScripts.AppendLine("<script>")
        'InjectScripts.AppendLine("$.get('invtbl/" & STYLE_CODE & ".html', function( data ) {")
        'InjectScripts.AppendLine("     $('#injectinv').html(data);")
        'InjectScripts.AppendLine("    });")
        'InjectScripts.AppendLine("</script>")
        'InjectScripts.AppendLine("<style>")
        'InjectScripts.AppendLine("div#menu-left{")
        'InjectScripts.AppendLine("background: #E2FFCF;")
        'InjectScripts.AppendLine("color: #445a3d;")
        'InjectScripts.AppendLine("}")
        'InjectScripts.AppendLine("</style>")
        'Return InjectScripts.ToString
        Return String.Format("<div id='injectinv' sku='{0}'></div>", STYLE_CODE)
    End Function

    Private Function MakeProductOnPagesNode(ByRef rowWBTSTYLD As DataRow, ByVal isParent As Boolean) As XmlNode
        Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE") & String.Empty
        Dim nodeProductOnOages As XmlNode = Nothing
        nodeProductOnOages = Nothing
        nodeProductOnOages = xmlLabelRequest.CreateElement("ProductOnPages")
        If isParent Then
            With nodeProductOnOages
                Dim sql As New Text.StringBuilder With {.Length = 0}
                sql.AppendLine("SELECT")
                sql.AppendLine("WBTPAGED.STYLE_CODE, WBTPAGEH.PAGE_CODE, WBTPAGEH.PAGE_NAME")
                sql.AppendLine("FROM WBTPAGEH, WBTPAGED")
                sql.AppendLine("WHERE WBTPAGEH.PAGE_CODE = WBTPAGED.PAGE_CODE")
                sql.AppendLine(String.Format("AND WBTPAGED.STYLE_CODE  = '{0}'", STYLE_CODE))
                Dim tblWBTPAGED As DataTable = ASCDATA1.GetDataTable(sql.ToString())
                For Each rowWBTPAGED As DataRow In tblWBTPAGED.Rows
                    Dim nodeName As XmlNode = Nothing
                    nodeName = xmlLabelRequest.CreateElement("Name")
                    Dim PAGE_NAME As String = rowWBTPAGED.Item("PAGE_NAME").ToString & String.Empty
                    nodeName.InnerText = PAGE_NAME
                    .AppendChild(nodeName)
                Next

                If rowWBTSTYLD.Item("STYLE_STATUS") & "" = "D" Then
                    Dim nodeDisc As XmlNode = Nothing
                    nodeDisc = xmlLabelRequest.CreateElement("Name")
                    nodeDisc.InnerText = "Discontinued"
                    .AppendChild(nodeDisc)
                End If
            End With
        End If
        Return nodeProductOnOages
    End Function

    'Private Function MakeProductOnPagesNode(ByRef rowWBTSTYLD As DataRow) As XmlNode
    '    Dim nodeProductOnOages As XmlNode = Nothing
    '    nodeProductOnOages = Nothing
    '    nodeProductOnOages = xmlLabelRequest.CreateElement("ProductOnPages")
    '    With nodeProductOnOages
    '        Dim nodeName As XmlNode = Nothing
    '        nodeName = xmlLabelRequest.CreateElement("Name")
    '        'I Know this is dumb but they keep changing their minds on classes.
    '        Dim STYLE_CLASS_DESC As String = ""
    '        Select Case rowWBTSTYLD.Item("STYLE_CLASS_CODE") & String.Empty
    '            Case "DECOR"
    '                STYLE_CLASS_DESC = "Home Decor"
    '            Case "FLOWER"
    '                STYLE_CLASS_DESC = "Flower"
    '            Case "FOLIAGE"
    '                STYLE_CLASS_DESC = "Foliage"
    '            Case "EASTER"
    '                STYLE_CLASS_DESC = "Easter"
    '            Case "GARDEN"
    '                STYLE_CLASS_DESC = "Garden"
    '            Case "VALENTINE"
    '                STYLE_CLASS_DESC = "Valentine"
    '            Case "GENERAL"
    '                STYLE_CLASS_DESC = "General"
    '            Case "HALLOWEEN"
    '                STYLE_CLASS_DESC = "Halloween"
    '            Case "PVC"
    '                STYLE_CLASS_DESC = "Christmas Greens"
    '            Case "XMAS"
    '                STYLE_CLASS_DESC = "Christmas"
    '            Case "FALL"
    '                STYLE_CLASS_DESC = "Harvest"
    '            Case Else
    '                STYLE_CLASS_DESC = rowWBTSTYLD.Item("STYLE_CLASS_CODE") & String.Empty
    '        End Select
    '        nodeName.InnerText = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(STYLE_CLASS_DESC.ToLower)
    '        .AppendChild(nodeName)
    '        If rowWBTSTYLD.Item("STYLE_COLOR_STATUS") & "" = "D" Then
    '            Dim nodeDisc As XmlNode = Nothing
    '            nodeDisc = xmlLabelRequest.CreateElement("Name")
    '            nodeDisc.InnerText = "Discontinued"
    '            .AppendChild(nodeDisc)
    '        End If

    '        'TODO: Mario Says He Is going to Provide Directions on this
    '        Dim nodeNextopia As XmlNode = Nothing
    '        nodeNextopia = xmlLabelRequest.CreateElement("Name")
    '        nodeNextopia.InnerText = "Nextopia Example - Page Name Here"
    '        .AppendChild(nodeNextopia)
    '    End With
    '    Return nodeProductOnOages
    'End Function

    Private Function MakeQuantityPricingNode(ByVal STYLE_CODE As String) As XmlNode
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
            MakeXMLNode(nodeQuantityPricing, "Comment", GetComments(BASE, STYLE_CODE))
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
                    MakeXMLNode(nodePriceBreak, "StartingQuantity", Discounts(ictr - 1).DISCOUNT_QTY.ToString("F0"))
                    MakeXMLNode(nodePriceBreak, "UnitPrice", Discounts(ictr - 1).DISCOUNT_PRICE.ToString("F2"))
                    nodePriceBreaks.AppendChild(nodePriceBreak)
                Next
            End With
            nodeQuantityPricing.AppendChild(nodePriceBreaks)
        End With
        Return nodeQuantityPricing
    End Function

    Private Sub MakeMoreInfoImages(ByVal nodeProduct As XmlNode, ByVal STYLE_CODE As String, ByVal isParent As Boolean)
        Dim SQLS As New System.Text.StringBuilder
        'GO_LIVE_CHANGES
        'SQLS.AppendLine("SELECT * FROM WBTSTYLD")
        'SQLS.AppendLine("WHERE CURR_ON_HAND <> 0")
        'SQLS.AppendLine("AND STYLE_CODE = :PARM1")
        SQLS.AppendLine("SELECT * FROM WBTSTYLD")
        SQLS.AppendLine("WHERE STYLE_CODE = :PARM1")
        Dim tblWBTSTYLD As DataTable = ASCDATA1.GetDataTable(SQLS.ToString(), "COLORCODES", "V", STYLE_CODE)
        For x As Integer = 1 To 20
            If isParent Then
                If tblWBTSTYLD.Rows.Count >= x Then
                    Dim nodeStr As String = String.Format("product/{0}-{1}.jpg", tblWBTSTYLD.Rows(x - 1).Item("STYLE_CODE").ToString & String.Empty, tblWBTSTYLD.Rows(x - 1).Item("COLOR_CODE").ToString & String.Empty)
                    MakeXMLNode(nodeProduct, String.Format("MoreInfoImage{0}", x.ToString), nodeStr)
                Else
                    MakeXMLNode(nodeProduct, String.Format("MoreInfoImage{0}", x.ToString), "none")
                End If
            Else
                MakeXMLNode(nodeProduct, String.Format("MoreInfoImage{0}", x.ToString), "none")
            End If
        Next
    End Sub

    Private Function MakeSubproductsNode(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal isParent As Boolean) As XmlNode
        Dim SQL As New Text.StringBuilder() With {.Length = 0}
        Dim nodeSubproductsNode As XmlNode = xmlLabelRequest.CreateElement("Subproducts")

        SQL.AppendLine("SELECT")
        SQL.AppendLine("WD.STYLE_CODE,")
        SQL.AppendLine("WD.COLOR_CODE,")
        SQL.AppendLine("WH.STYLE_DESC_SHORT,")
        SQL.AppendLine("S1.STYLE_DESC")
        SQL.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1")
        SQL.AppendLine("WHERE WH.STYLE_CODE = WD.STYLE_CODE")
        SQL.AppendLine("AND WH.STYLE_CODE = S1.STYLE_CODE")
        'SQL.AppendLine("AND WEB_IND IN ('U','W')")
        'SQL.AppendLine("AND CURR_ON_HAND > 0")
        SQL.AppendLine(Format("AND WD.STYLE_CODE = '{0}'", STYLE_CODE))
        'If Not isParent Then
        '    SQL.AppendLine(Format("AND WD.COLOR_CODE <> '{0}'", COLOR_CODE))
        'End If
        'GO_LIVE_CHANGES
        'SQL.AppendLine("AND WD.CURR_ON_HAND > 0")
        SQL.AppendLine("ORDER BY WD.COLOR_CODE")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), String.Empty)
        For Each rowWBTSTYLD As DataRow In tbl.Rows
            Dim COLOR_DESC As String = GetCOLOR_DESC(rowWBTSTYLD.Item("COLOR_CODE").ToString)
            Dim nodeSubproductsItem As XmlNode = xmlLabelRequest.CreateElement("Subproduct")
            Dim nodeName As XmlNode = Nothing
            nodeName = xmlLabelRequest.CreateElement("Name")
            'nodeName.InnerText = String.Format("{0} ({1})", GetSTYLE_DESC_SHORT(rowWBTSTYLD), COLOR_DESC)
            nodeName.InnerText = String.Format("{0}", COLOR_DESC)
            nodeSubproductsItem.AppendChild(nodeName)

            Dim nodeSKU As XmlNode = Nothing
            nodeSKU = xmlLabelRequest.CreateElement("SKU")
            nodeSKU.InnerText = String.Format("{0}-{1}", rowWBTSTYLD.Item("STYLE_CODE"), rowWBTSTYLD.Item("COLOR_CODE").ToString)
            nodeSubproductsItem.AppendChild(nodeSKU)
            nodeSubproductsNode.AppendChild(nodeSubproductsItem)
        Next
        Return nodeSubproductsNode
    End Function

    Private Sub MakeXMLNode(ByRef NodeToAppend As XmlNode,
                                 ByVal Element As String, Optional ByVal InnerText As String = "",
                                 Optional ByVal Attr As String = "",
                                 Optional ByVal AttrValue As String = "")
        Dim node As XmlNode = Nothing
        node = xmlLabelRequest.CreateElement(Element)
        If Attr.Length > 0 Then
            Dim typeAttr As XmlAttribute = xmlLabelRequest.CreateAttribute(Attr)
            typeAttr.Value = AttrValue
            node.Attributes.Append(typeAttr)
        End If
        If InnerText.Length > 0 Then
            node.InnerText = InnerText
        End If
        NodeToAppend.AppendChild(node)
    End Sub
#End Region

#Region "Custom Methods"
    Private Function GetAttributes(ByVal STYLE_CODE As String) As String
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

    Private Function GetCOLOR_DESC(COLOR_CODE As String) As String
        Dim RetVal As String = ""
        ASCMAIN1.sql = String.Format("Select NVL(COLOR_DESC,'') from ICTCOLR1 where COLOR_CODE = '{0}'", COLOR_CODE)
        Dim COLOR_DESC As String = ASCDATA1.GetDataValue
        If COLOR_DESC.Length = 0 Then
            RetVal = COLOR_CODE
        Else
            Dim RepVals As New Dictionary(Of String, String)
            RepVals.Add("/", " ")
            RepVals.Add(",", "")
            RepVals.Add(".", "")
            RepVals.Add("#", "")
            RepVals.Add("?", "")
            RepVals.Add("!", "")
            RepVals.Add("*", "")
            RepVals.Add("%", "")
            RepVals.Add("@", "")
            For Each RepVal As KeyValuePair(Of String, String) In RepVals
                COLOR_DESC = COLOR_DESC.Replace(RepVal.Key, RepVal.Value)
            Next
            RetVal = COLOR_DESC
        End If
        Return RetVal
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

    Private Function GetComments(frm As ASFBASE0, ByVal STYLE_CODE As String) As String
        Dim retVal As String = ""
        Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", STYLE_CODE)
        If Not IsNothing(rowICTSTYL1) Then
            retVal = retVal & "<link rel='stylesheet' type='text/css' href='http://www.regency-rib.com/css/regency.css'>"
            retVal = retVal & "<SPAN>"
            retVal = retVal & "<STRONG>UOM:</STRONG>" & rowICTSTYL1.Item("STYLE_UOM").ToString & " | "
            'Removing Bags per Mario - 4/27/18 wr
            'retVal = retVal & "<STRONG>BAG:</STRONG>" & rowICTSTYL1.Item("SUB_UNIT_PACK_QTY").ToString & " | "
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
            retVal = retVal & "<span> <input class='add' type='submit' onclick='history.back();' value='Continue Shopping'> </span>"
        End If
        Return retVal
    End Function

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

    Private Function GetMinimumQuantity(ByVal STYLE_CODE As String) As Int16
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

    Private Function GetProductDisabled(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal isParent As Boolean) As String
        Dim RetVal As String = "uncheck"
        Dim hasValidColor As Boolean = False
        'Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND WEB_IND = 'W'", STYLE_CODE, COLOR_CODE)
        Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND CURR_ON_HAND > 0", STYLE_CODE, COLOR_CODE)
        If isParent Then
            'filter = String.Format("STYLE_CODE = '{0}' AND WEB_IND = 'W'", STYLE_CODE)
            filter = String.Format("STYLE_CODE = '{0}' AND CURR_ON_HAND > 0", STYLE_CODE)
        End If
        For Each rowWBTSTYLD2 As DataRow In dstWBTSTYLD.Select(Filter)
            If Val(rowWBTSTYLD2.Item("CURR_ON_HAND").ToString & String.Empty) > 0 Then
                hasValidColor = True
            End If
        Next
        If Not hasValidColor Then
            RetVal = "checked"
        End If
        Return RetVal
    End Function

    Private Function GetSTYLE_DESC_SHORT(ByRef rowWBTSTYLD As DataRow) As String
        Dim RetVal As String = ""
        RetVal = rowWBTSTYLD.Item("STYLE_DESC_SHORT") & String.Empty
        If RetVal.Length = 0 Then
            RetVal = rowWBTSTYLD.Item("STYLE_DESC") & String.Empty
        End If
        Return RetVal
    End Function

    Private Function GetSTYLE_DESC_LONG(ByVal STYLE_CODE As String) As String
        Dim RetVal As String = ""
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine(String.Format("SELECT NVL(STYLE_DESC_LONG,'') AS STYLE_DESC_LONG FROM WBTSTYLH WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        RetVal = ASCDATA1.GetDataValue
        Return RetVal
    End Function

    Private Function GetStyleRow(ByVal StyleCode As String, ByVal ColorCode As String) As DataTable
        Dim sql As New Text.StringBuilder
        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("SL.STYLE_CODE,")
        sql.AppendLine("SC.COLOR_CODE,")
        sql.AppendLine("SL.STYLE_DESC,")
        sql.AppendLine("C1.COLOR_DESC,")
        sql.AppendLine("C1.COLOR_CODE_LONG,")
        sql.AppendLine("SL.STYLE_STATUS,")
        sql.AppendLine("WS.WEB_IND,")
        sql.AppendLine("SC.STYLE_COLOR_STATUS,")
        sql.AppendLine("SL.INNER_PACK_QTY,")
        sql.AppendLine("SL.CARTON_PACK_QTY,")
        sql.AppendLine("SL.STYLE_UOM,")
        sql.AppendLine("SL.SUB_UNIT_PACK_QTY,")
        sql.AppendLine("SL.STYLE_CLASS_CODE,")
        sql.AppendLine("CL.STYLE_CLASS_DESC,")
        sql.AppendLine("SL.STYLE_SO_QTY_MIN,")
        sql.AppendLine("SL.STYLE_MATL_DESC,")
        sql.AppendLine("WH.STYLE_DESC_SHORT,")
        sql.AppendLine("NVL(SL.SIZE_CODE,'') AS SIZE_CODE,")
        sql.AppendLine("SC.UPC_CODE,")
        sql.AppendLine("WH.SEARCH_KEYWORDS,")
        sql.AppendLine("WH.MATERIALS,")
        sql.AppendLine("WH.META_DESC,")
        sql.AppendLine("NVL(SC.THEME_CODE,'') AS THEME_CODE,")
        sql.AppendLine("((NVL(ST.WHSE_QTY_ON_HAND,0) - NVL(ST.WHSE_QTY_PICK,0)) + NVL(ST.WHSE_QTY_TRAN,0) + NVL(ST.WHSE_QTY_ON_ORDER,0) - NVL(ST.WHSE_QTY_OPEN,0)) AS FTR_AVAIL")
        sql.AppendLine("FROM WBTSTYLD WS, WBTSTYLH WH, ICTSTYL1 SL, ICTSTYC1 SC, ICTSTAT2 ST, ICTCLAS1 CL, ICTCOLR1 C1")
        sql.AppendLine("WHERE WS.STYLE_CODE = SL.STYLE_CODE")
        sql.AppendLine("AND WS.STYLE_CODE = WH.STYLE_CODE (+)")
        sql.AppendLine("AND WS.COLOR_CODE = SC.COLOR_CODE")
        sql.AppendLine("AND SL.STYLE_CODE = SC.STYLE_CODE")
        sql.AppendLine("AND SC.STYLE_CODE = ST.STYLE_CODE")
        sql.AppendLine("AND SC.COLOR_CODE = ST.COLOR_CODE")
        sql.AppendLine("AND SL.STYLE_CLASS_CODE = CL.STYLE_CLASS_CODE")
        sql.AppendLine("AND SC.COLOR_CODE = C1.COLOR_CODE")
        sql.AppendLine("AND ST.WHSE_CODE = 'MS'")
        'sql.AppendLine("AND SL.STYLE_STATUS = 'A'")
        'GO_LIVE_CHANGES
        'sql.AppendLine("AND SC.STYLE_COLOR_STATUS = 'A'")
        'sql.AppendLine("AND ((NVL(ST.WHSE_QTY_ON_HAND,0) - NVL(ST.WHSE_QTY_PICK,0)) + NVL(ST.WHSE_QTY_TRAN,0) + NVL(ST.WHSE_QTY_ON_ORDER,0) - NVL(ST.WHSE_QTY_OPEN,0)) > 0")
        sql.AppendLine("AND SL.STYLE_CODE = :PARM1")
        sql.AppendLine("AND SC.COLOR_CODE = :PARM2")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "VV", New String() {StyleCode, ColorCode})
        Return tbl
    End Function

    Private Function GetWEB_DESC(ByRef rowWBTSTYLD As DataRow, ByVal isParent As Boolean) As String
        Dim RetVal As String = ""
        Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE").ToString & String.Empty
        Dim COLOR_CODE As String = rowWBTSTYLD.Item("COLOR_CODE").ToString & String.Empty
        Dim COLOR_DESC As String = GetCOLOR_DESC(COLOR_CODE)
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine(String.Format("SELECT NVL(WEB_DESC,'') AS WEB_DESC FROM WBTSTYLH WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim WEB_DESC As String = ASCDATA1.GetDataValue
        If WEB_DESC.Length > 0 Then
            RetVal = WEB_DESC
        Else
            RetVal = rowWBTSTYLD.Item("STYLE_DESC") & String.Empty
        End If
        If isParent Then
            RetVal = String.Format("{0}", RetVal)
        Else
            RetVal = String.Format("{0} ({1})", RetVal, COLOR_DESC)
        End If
        Return RetVal
    End Function

    Private Function RefreshColorTable(ByVal Style_Code As String) As DataTable
        Dim SQLS As New System.Text.StringBuilder
        SQLS.AppendLine("SELECT * FROM WBTSTYLD")
        SQLS.AppendLine("WHERE STYLE_CODE = :PARM1")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(SQLS.ToString(), "COLORCODES", "V", Style_Code)
        Return tbl
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
#End Region
End Class
