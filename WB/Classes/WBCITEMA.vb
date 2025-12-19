Imports System.Xml
Imports System.Net.Mail
Imports System.String
Imports System.Runtime.Remoting.Contexts
Imports Infragistics.Win.UltraWinGrid
Imports System.Text

Public Class WBCITEMA
    '-----------------------------------------------------
    '-This is The Newest Newest Class That Produces XML Documents
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
    Private data As DataSet
    Private BASE As New ASFBASE0
    Private testing As Boolean = False
    Private DISABLED_STYLES As List(Of String)

#Region "Class Public Methods"
    Public Sub New(ByRef _data As DataSet, ByRef _DISABLED_STYLES As List(Of String))
        Me.InitiailizeClass()
        data = _data
        DISABLED_STYLES = _DISABLED_STYLES
    End Sub

    Public Sub Clear()
        InitiailizeClass()
    End Sub

    Public Function AddStyle(ByVal StyleCode As String,
                             ByVal ColorCode As String,
                             ByVal DelList As List(Of String),
                             ByVal Optional UploadInventoryOnly As Boolean = True,
                             ByVal Optional isParent As Boolean = False,
                             ByVal Optional UpdatePricing As Boolean = False,
                             ByVal Optional isTesting As Boolean = False) As Integer
        Dim nodesProcessed As Integer = 0
        Dim ictr As Integer = 1
        testing = isTesting

        Dim CFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", StyleCode, ColorCode)
        For Each rowWBTSTYLD As DataRow In data.Tables("STATUS").Select(CFilter)
            Dim nodeProduct As XmlNode = Nothing

            If nodesProcessed = 0 Then
                nodeProduct = MakeProductNode(rowWBTSTYLD, True, UploadInventoryOnly, isParent, UpdatePricing)
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
        Dim CFilter As String = String.Format("STYLE_CODE = '{0} AND COLOR_CODE = '{1}''", STYLE_CODE, COLOR_CODE)
        For Each rowCROSS As DataRow In data.Tables("CROSS").Select(CFilter, "COLOR_CODE")
            Dim COLOR_DESC As String = GetCOLOR_DESC(rowCROSS.Item("COLOR_CODE").ToString)
            Dim nodeCrossSellItem As XmlNode = xmlLabelRequest.CreateElement("CrossSellItem")
            Dim nodeName As XmlNode = Nothing
            nodeName = xmlLabelRequest.CreateElement("Name")
            nodeName.InnerText = String.Format("{0} {1}", GetSTYLE_DESC_SHORT(rowCROSS), COLOR_DESC)
            nodeCrossSellItem.AppendChild(nodeName)

            Dim nodeSKU As XmlNode = Nothing
            nodeSKU = xmlLabelRequest.CreateElement("SKU")
            nodeSKU.InnerText = String.Format("{0}-{1}", rowCROSS.Item("STYLE_CODE"), rowCROSS.Item("COLOR_CODE").ToString)
            nodeCrossSellItem.AppendChild(nodeSKU)
            nodeCrossSellNode.AppendChild(nodeCrossSellItem)
        Next
        Return nodeCrossSellNode
    End Function

    Private Function MakeProductNode(ByVal rowWBTSTYLD As DataRow,
                                     ByVal mainStyle As Boolean,
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

            MakeXMLNode(nodeProduct, "Name", GetWEB_DESC(rowWBTSTYLD, isParent))
            If isParent Then
                MakeXMLNode(nodeProduct, "SKU", STYLE_CODE)
                'MakeXMLNode(nodeProduct, "QuantityOnHand", 0)
            Else
                MakeXMLNode(nodeProduct, "SKU", STYLE_CODE & "-" & COLOR_CODE)
                'MakeXMLNode(nodeProduct, "QuantityOnHand", Val(rowWBTSTYLD.Item("CURR_QTY_AVAIL") & String.Empty))
            End If
            MakeXMLNode(nodeProduct, "ProductDisabled", GetProductDisabled(STYLE_CODE, COLOR_CODE, isParent, rowWBTSTYLD))
            nodeProduct.AppendChild(MakeProductOnPagesNode(rowWBTSTYLD, isParent))
            MakeXMLNode(nodeProduct, "AddToPages")
            If Not UploadInventoryOnly Then
                MakeXMLNode(nodeProduct, "ProductDescription", GetSTYLE_DESC_LONG(STYLE_CODE))
                MakeXMLNode(nodeProduct, "MinimumQuantity", GetMinimumQuantity(STYLE_CODE))
                nodeProduct.AppendChild(MakeSubproductsNode(STYLE_CODE, COLOR_CODE, isParent))
                nodeProduct.AppendChild(MakeOptionMenusNode(STYLE_CODE))
                nodeProduct.AppendChild(MakeProductOptionsNode(STYLE_CODE))
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
                'MakeXMLNode(nodeProduct, "CustomerTextEntryBox", "uncheck")
                MakeXMLNode(nodeProduct, "CustomerTextEntryBox", "checked")
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
                MakeXMLNode(nodeProduct, "Availability", "In stock")
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
                'nodeProduct.AppendChild(MakeProductOnPagesNode(rowWBTSTYLD, isParent))
                'MakeXMLNode(nodeProduct, "AddToPages")
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
            MakeProductFieldNodes(nodeProduct, STYLE_CODE, COLOR_CODE, isParent, UploadInventoryOnly, rowWBTSTYLD, UpdatePricing)
            '<ProductID>12</ProductID> 'Make Sure We Don't Have to set this.
            'MakeXMLNode(nodeProduct, "BlankEntry")
            productNodeCount += 1
        Catch ex As Exception
            nodeProduct = Nothing
        End Try

        Return nodeProduct
    End Function

    Private Function MakeOptionMenusNode(ByVal STYLE_CODE As String) As XmlNode
        Dim nodeOptionMenus As XmlNode = Nothing
        nodeOptionMenus = xmlLabelRequest.CreateElement("OptionMenus")
        Dim nodeMenu As XmlNode = Nothing
        nodeMenu = xmlLabelRequest.CreateElement("Menu")
        MakeXMLNode(nodeMenu, "MenuItem", "Color;n")
        Dim CFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        For Each rowTbl As DataRow In data.Tables("WBTSTYLD").Select(CFilter, "COLOR_CODE")
            MakeXMLNode(nodeMenu, "MenuItem", rowTbl.Item("COLOR_CODE").ToString & String.Empty)
        Next
        nodeOptionMenus.AppendChild(nodeMenu)
        Return nodeOptionMenus
    End Function

    Private Function MakeProductOptionsNode(ByVal STYLE_CODE As String) As XmlNode
        Dim nodeProductOptions As XmlNode = Nothing
        nodeProductOptions = xmlLabelRequest.CreateElement("ProductOptions")
        With nodeProductOptions
            Dim CFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            For Each rowCOLOR As DataRow In data.Tables("WBTSTYLD").Select(CFilter, "COLOR_CODE")
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
                        'MakeXMLNode(nodeProductOption, "QuantityOnHand")
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
                                      ByVal rowWBTSTYLD As DataRow,
                                      Optional ByVal UpdatePricing As Boolean = False)
        'Dim BASE As New ASFBASE0
        'Dim rowICTSTYL1 As DataRow = BASE.LookUp("ICTSTYL1", STYLE_CODE)
        Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        Dim rowICTSTYL1 As DataRow = data.Tables("ICTSTYL1").Select(SFilter).FirstOrDefault
        For i As Integer = 1 To 20
            Select Case i
                Case 1
                    If Not UploadInventoryOnly Then
                        MakeXMLNode(nodeProduct, "ProductField1", GetMinimumQuantity(STYLE_CODE))
                    End If
                Case 2
                    MakeXMLNode(nodeProduct, "ProductField2", GetNextDelDate(STYLE_CODE, COLOR_CODE, rowWBTSTYLD))
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
                Case 8
                    If Not UploadInventoryOnly Then
                        If isParent Then
                            Dim PFC8 As String = ""
                            If rowWBTSTYLD.Item("FLAG_NEW").ToString & String.Empty = "1" Then
                                PFC8 = "featured"
                            End If
                            MakeXMLNode(nodeProduct, "ProductField8", PFC8)
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
                Case 13
                    If Not UploadInventoryOnly Then
                        Dim BULLETS As String = MakeBullets(rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty)
                        If BULLETS.Length > 0 Then
                            MakeXMLNode(nodeProduct, "ProductField13", BULLETS)
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
            'Dim rowARTCUST1 As DataRow = BASE.LookUp("ARTCUST1", "180000")
            Dim CFilter As String = String.Format("CUST_CODE = '{0}'", "180000")
            Dim rowARTCUST1 As DataRow = data.Tables("ARTCUST1").Select(CFilter).FirstOrDefault

            'Dim rowICTCLAS1 As DataRow = BASE.LookUp("ICTCLAS1", rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString & String.Empty)
            Dim SCFilter As String = String.Format("STYLE_CLASS_CODE = '{0}'", rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString & String.Empty)
            Dim rowICTCLAS1 As DataRow = data.Tables("ICTCLAS1").Select(SCFilter).FirstOrDefault

            Dim IsPVC As Boolean = rowICTCLAS1.Item("DISC_CODE").ToString = "PVC"
            Dim isDiscontunued As Boolean = rowICTSTYL1.Item("STYLE_STATUS").ToString = "D"
            If isDiscontunued Then
                Dim rowNOTHING As DataRow = Nothing
                Dim Discounts As List(Of DISCOUNTS)
                Discounts = SOCMAIN2.Price_Discounts(BASE, "", rowNOTHING, STYLE_CODE, False, True, True)
                'Stop
                MakeXMLNode(nodeProduct, "ProductField9", Discounts(0).DISCOUNT_PRICE)
            End If
            If testing Then
                Stop
                'For i As Integer = 28 To 32
                '    If IsPVC Or isDiscontunued Then
                '        MakeXMLNode(nodeProduct, "ProductField" & i.ToString)
                '    Else
                '        Dim CUST_PRICE_TIER As String = "PC"
                '        Dim CUST_DISC_PCT_EXTRA As String = "0"
                '        Dim CUST_DISC_PCT As Integer = 0
                '        Select Case i
                '            Case 28
                '                CUST_PRICE_TIER = "PC"
                '                CUST_DISC_PCT_EXTRA = "1"
                '            Case 29
                '                CUST_PRICE_TIER = "PC"
                '                CUST_DISC_PCT_EXTRA = "2"
                '            Case 30
                '                CUST_PRICE_TIER = "HC"
                '            Case 31
                '                CUST_PRICE_TIER = "FC"
                '            Case 32
                '                'This was adjusted again per Danny and Rich from 54 to 50.  W.R.-1/15/21
                '                CUST_PRICE_TIER = "SP"
                '                CUST_DISC_PCT = 50
                '            Case 33
                '            Case 34
                '                'This was reversed by Danny.  Anyone SP is always 54, Otherwise default pricing.
                '                'CUST_PRICE_TIER = "SP"
                '                'CUST_DISC_PCT = 54
                '        End Select
                '        rowARTCUST1.Item("CUST_PRICE_TIER") = CUST_PRICE_TIER
                '        rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = CUST_DISC_PCT_EXTRA
                '        rowARTCUST1.Item("CUST_DISC_PCT") = CUST_DISC_PCT
                '        'MakeXMLNode(nodeProduct, "ProductField" & i.ToString)
                '        Dim Discounts As List(Of DISCOUNTS)
                '        Dim MaxBreak As Integer = 4
                '        'Dim nodeQuantityPricing As XmlNode
                '        Discounts = SOCMAIN2.Price_Discounts(BASE, "", rowARTCUST1, STYLE_CODE, True, False, True)
                '        For z As Integer = 1 To 4
                '            If Discounts(z - 1).DISCOUNT_QTY = 0 Then
                '                MaxBreak = z - 1
                '                Exit For
                '            End If
                '        Next
                '        Dim DISC_DESC As New Text.StringBuilder With {.Length = 0}
                '        For ictr As Integer = MaxBreak - 1 To 0 Step -1
                '            If ictr = 0 Then
                '                'DISC_DESC.AppendLine(String.Format("{0}", Math.Round(Discounts(ictr).DISCOUNT_PRICE, 2)))
                '                DISC_DESC.Append(String.Format("{0}", Math.Round(Discounts(ictr).DISCOUNT_PRICE, 2)))
                '            Else
                '                If CUST_PRICE_TIER = "HC" And ictr = 3 Then
                '                    DISC_DESC.AppendLine(String.Format("{0}|{1}", Discounts(ictr - 1).DISCOUNT_QTY - 1, Math.Round(Discounts(ictr - 1).DISCOUNT_PRICE, 2)))
                '                Else
                '                    If CUST_PRICE_TIER = "FC" And ictr <= 3 Then
                '                        DISC_DESC.AppendLine(String.Format("{0}|{1}", Discounts(ictr - 1).DISCOUNT_QTY - 1, Math.Round(Discounts(1).DISCOUNT_PRICE, 2)))
                '                    Else
                '                        DISC_DESC.AppendLine(String.Format("{0}|{1}", Discounts(ictr - 1).DISCOUNT_QTY - 1, Math.Round(Discounts(ictr).DISCOUNT_PRICE, 2)))
                '                    End If
                '                End If
                '            End If
                '        Next
                '        MakeXMLNode(nodeProduct, "ProductField" & i.ToString, DISC_DESC.ToString)
                '    End If
                'Next
            Else
                Dim PGoups As Int64() = {28, 29, 30, 31, 32, 33, 38, 39, 40, 41}
                For i As Integer = 28 To 42
                    If PGoups.Contains(i) Then
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
                                    CUST_DISC_PCT = 52
                                Case 33
                                    CUST_PRICE_TIER = "SP"
                                    CUST_DISC_PCT = 54
                                Case 38
                                    CUST_PRICE_TIER = "SP"
                                    CUST_DISC_PCT = 55
                                Case 39
                                    CUST_PRICE_TIER = "SP"
                                    CUST_DISC_PCT = 56
                                Case 40
                                    CUST_PRICE_TIER = "SP"
                                    CUST_DISC_PCT = 57
                                Case 41
                                    CUST_PRICE_TIER = "SP"
                                    CUST_DISC_PCT = 59
                            End Select
                            rowARTCUST1.Item("CUST_PRICE_TIER") = CUST_PRICE_TIER
                            rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") = CUST_DISC_PCT_EXTRA
                            rowARTCUST1.Item("CUST_DISC_PCT") = CUST_DISC_PCT
                            'MakeXMLNode(nodeProduct, "ProductField" & i.ToString)
                            Dim Discounts As List(Of DISCOUNTS)
                            Dim MaxBreak As Integer = 4
                            'Dim nodeQuantityPricing As XmlNode
                            Discounts = SOCMAIN2.Price_Discounts(BASE, "", rowARTCUST1, STYLE_CODE, True, False, True)
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
                    End If

                Next
                MakeXMLNode(nodeProduct, "ProductField33", "")
            End If


            Dim PFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)

            If data.Tables("PROMOS").Select(PFilter).Count > 0 Then
                Dim rowICTPROM1 As DataRow = data.Tables("PROMOS").Select(PFilter, "PROMO_START_DATE").FirstOrDefault
                Dim PROMO_UNIT_PRICE As Double = Val(rowICTPROM1.Item("PROMO_UNIT_PRICE").ToString & String.Empty)
                Dim PROMO_START_DATE As DateTime = CDate(rowICTPROM1.Item("PROMO_START_DATE").ToString & String.Empty)
                Dim PROMO_END_DATE As DateTime = CDate(rowICTPROM1.Item("PROMO_END_DATE").ToString & String.Empty)
                If PROMO_UNIT_PRICE > 0 Then
                    Dim P1S As String = PROMO_UNIT_PRICE.ToString("#####0.00")
                    Dim PD1 As String = String.Format("{0}/{1}/{2}", PROMO_START_DATE.Month.ToString("00"), PROMO_START_DATE.Day.ToString("00"), PROMO_START_DATE.Year.ToString("0000"))
                    Dim PD2 As String = String.Format("{0}/{1}/{2}", PROMO_END_DATE.Month.ToString("00"), PROMO_END_DATE.Day.ToString("00"), PROMO_END_DATE.Year.ToString("0000"))
                    Dim PromoString As String = String.Format("{0}|{1}+{2}", P1S, PD1, PD2)
                    MakeXMLNode(nodeProduct, "ProductField34", PromoString)
                End If
            Else
                MakeXMLNode(nodeProduct, "ProductField34", "")
            End If
            If isParent Then
                Dim PFC35 As String = ""
                If rowWBTSTYLD.Item("FLAG_NEW").ToString & String.Empty = "1" Then
                    PFC35 = "NEW"
                End If
                MakeXMLNode(nodeProduct, "ProductField35", PFC35)

                Dim PFC36 As String = ""
                Select Case rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString & String.Empty
                    Case "DECOR"
                        PFC36 = "Home Decor"
                    Case "EASTER"
                        PFC36 = "Easter"
                    Case "FALL"
                        PFC36 = "Fall"
                    Case "FLOWER"
                        PFC36 = "Flowers"
                    Case "FOLIAGE"
                        PFC36 = "Foliage"
                    Case "GARDEN"
                        PFC36 = "Garden"
                    Case "PVC"
                        PFC36 = "Christmas Greens (PVC)"
                    Case "XMAS"
                        PFC36 = "Christmas Décor"
                End Select
                MakeXMLNode(nodeProduct, "ProductField36", PFC36)
            Else
                MakeXMLNode(nodeProduct, "ProductField35", "")
                MakeXMLNode(nodeProduct, "ProductField36", "")
            End If
            MakeXMLNode(nodeProduct, "ProductField37", "")
        End If
    End Sub

    Private Function MakeBullets(ByVal STYLE_CODE As String) As String
        Dim RetVal As String = ""
        Dim fltr As String = $"STYLE_CODE = '{STYLE_CODE}'"
        For Each rowICTBULT1 As DataRow In data.Tables("ICTBULT1").Select(fltr, "LINE_NO")
            RetVal = RetVal & rowICTBULT1.Item("BULLET_TEXT") & vbCrLf
        Next
        If RetVal.Length > 2 Then
            RetVal = RetVal.Substring(0, RetVal.Length - 2)
        End If
        Return RetVal
    End Function

    Private Function GetCategoryGroup(ByVal STYLE_CODE As String) As String
        Dim CAT_DESC As New Text.StringBuilder With {.Length = 0}

        Dim AFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        For Each rowICTATTR1 As DataRow In data.Tables("ATTR_DESC").Select(AFilter)
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

        Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)

        For Each rowICTCOLR1 As DataRow In data.Tables("SIZE_CODE").Select(SFilter)
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

        Dim CFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        For Each rowICTCOLR1 As DataRow In data.Tables("COLOR_DESC").Select(CFilter)
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

    Private Function GetNextDelDate(STYLE_CODE As String, COLOR_CODE As String, ByVal rowWBTSTYLD As DataRow) As String
        Dim retVal As String = ""
        If (rowWBTSTYLD.Item("FUT_DATE") & String.Empty) <> "" And Val(rowWBTSTYLD.Item("FUT_QTY_AVAIL") & String.Empty) > 0 Then
            retVal = (rowWBTSTYLD.Item("FUT_DATE") & String.Empty) & "|" & Val(rowWBTSTYLD.Item("FUT_QTY_AVAIL") & String.Empty)
        End If
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
        Dim IsPromotions As Boolean = False
        If isParent Then
            With nodeProductOnOages
                Dim WFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
                For Each rowWBTPAGEX As DataRow In data.Tables("WBTPAGEX").Select(WFilter)
                    Dim nodeName As XmlNode = Nothing
                    nodeName = xmlLabelRequest.CreateElement("Name")
                    Dim PAGE_NAME As String = rowWBTPAGEX.Item("PAGE_NAME").ToString & String.Empty
                    If PAGE_NAME = "Promotions" Then
                        IsPromotions = True
                    End If
                    nodeName.InnerText = PAGE_NAME
                    .AppendChild(nodeName)
                Next

                If rowWBTSTYLD.Item("STYLE_STATUS") & "" = "D" Then
                    Dim nodeDisc As XmlNode = Nothing
                    nodeDisc = xmlLabelRequest.CreateElement("Name")
                    nodeDisc.InnerText = "Discontinued"
                    .AppendChild(nodeDisc)
                End If

                If Not IsPromotions Then

                    Dim S As New StringBuilder With {.Length = 0}
                    S.AppendLine("SELECT")
                    S.AppendLine("COUNT(*)")
                    S.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2")
                    S.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
                    S.AppendLine(String.Format("AND P2.STYLE_CODE = '{0}'", STYLE_CODE))
                    S.AppendLine("AND (SYSDATE >= PROMO_START_DATE AND SYSDATE <= PROMO_END_DATE)")
                    ASCMAIN1.sql = S.ToString()
                    Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
                    If REC_CNT > 0 Then
                        Dim nodeDisc As XmlNode = Nothing
                        nodeDisc = xmlLabelRequest.CreateElement("Name")
                        nodeDisc.InnerText = "Promotions"
                        .AppendChild(nodeDisc)
                    End If
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
        'Dim BASE As New ASFBASE0
        Dim MaxBreak As Integer = 4
        Dim nodeQuantityPricing As XmlNode
        Discounts = SOCMAIN2.Price_Discounts(BASE, "", rowARTCUST1, STYLE_CODE, False, True, True)
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
            MakeXMLNode(nodeQuantityPricing, "Comment", GetComments(STYLE_CODE))
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

        Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        data.Tables("WBTSTYLD").Select(SFilter)
        For x As Integer = 1 To 20
            If isParent Then
                If data.Tables("WBTSTYLD").Select(SFilter).Count >= x Then
                    Dim nodeStr As String = String.Format("product/{0}-{1}.jpg", data.Tables("WBTSTYLD").Select(SFilter).ElementAt(x - 1).Item("STYLE_CODE").ToString & String.Empty, data.Tables("WBTSTYLD").Select(SFilter).ElementAt(x - 1).Item("COLOR_CODE").ToString & String.Empty)
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

        Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        For Each rowWBTSTYLX As DataRow In data.Tables("COLORS").Select(SFilter, "COLOR_CODE")
            Dim COLOR_DESC As String = GetCOLOR_DESC(rowWBTSTYLX.Item("COLOR_CODE").ToString)
            Dim nodeSubproductsItem As XmlNode = xmlLabelRequest.CreateElement("Subproduct")
            Dim nodeName As XmlNode = Nothing
            nodeName = xmlLabelRequest.CreateElement("Name")
            'nodeName.InnerText = String.Format("{0} ({1})", GetSTYLE_DESC_SHORT(rowWBTSTYLD), COLOR_DESC)
            nodeName.InnerText = String.Format("{0}", COLOR_DESC)
            nodeSubproductsItem.AppendChild(nodeName)

            Dim nodeSKU As XmlNode = Nothing
            nodeSKU = xmlLabelRequest.CreateElement("SKU")
            nodeSKU.InnerText = String.Format("{0}-{1}", rowWBTSTYLX.Item("STYLE_CODE"), rowWBTSTYLX.Item("COLOR_CODE").ToString)
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
        Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)

        For Each rowATTR_CODE As DataRow In data.Tables("ATTR_CODE").Select(SFilter)
            Dim ATR As String = rowATTR_CODE.Item("ATTR_DESC").ToString & ""
            If ATR = "DÉCOR" Then
                ATR = "DECOR"
            End If
            ATTRS += ATR.ToLower & ", "
            ATTRS += ATR.ToUpper & ", "
            ATTRS += StrConv(ATR.ToLower, VbStrConv.ProperCase) & ", "
        Next
        Dim STYLE_STATUS As String = data.Tables("ICTSTYL1").Select(SFilter).FirstOrDefault.Item("STYLE_STATUS").ToString & String.Empty
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
        Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        For Each rowICTCOLR1 As DataRow In data.Tables("COLOR_GROUP_CODE").Select(SFilter)
            RetVal += rowICTCOLR1.Item("COLOR_GROUP_CODE").ToString.ToLower & ", "
            RetVal += rowICTCOLR1.Item("COLOR_GROUP_CODE").ToString.ToUpper & ", "
            RetVal += StrConv(rowICTCOLR1.Item("COLOR_GROUP_CODE").ToString.ToLower, VbStrConv.ProperCase) & ", "
        Next
        If RetVal.Length > 2 Then
            RetVal = RetVal.Substring(0, RetVal.Length - 2)
        Else
            RetVal = ""
        End If

        Return RetVal
    End Function

    Private Function GetComments(ByVal STYLE_CODE As String) As String
        Dim retVal As String = ""
        'Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", STYLE_CODE)
        Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        Dim rowICTSTYL1 As DataRow = data.Tables("ICTSTYL1").Select(SFilter).FirstOrDefault
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

    'Private Function GetImageName(ByVal STYLE_CODE As Object, ByVal COLOR_CODE As String) As String
    '    Dim RetVal As String = String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE)
    '    Dim url As New System.Uri("http://api.regency-rib.com:8181/images/product/" & RetVal)
    '    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
    '    Dim resptest As System.Net.WebResponse
    '    Dim ErrorsFound As Boolean = False
    '    req.Timeout = 3000
    '    Try
    '        resptest = req.GetResponse()
    '    Catch ex As Exception
    '        ErrorsFound = True
    '        RetVal = ""
    '        req = Nothing
    '    End Try
    '    Return RetVal
    'End Function

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

        If AVAILABLE < MinimumQuantity Then
            MinimumQuantity = AVAILABLE
        End If

        Return MinimumQuantity
    End Function

    Private Function GetProductDisabled(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal isParent As Boolean, ByRef rowWBTSTYLD As DataRow) As String
        'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
        '    If STYLE_CODE = "MT24880" Then Stop
        'End If

        Dim RetVal As String = "uncheck"
        Dim hasValidColor As Boolean = False
        Dim filter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        For Each rowWBTSTYLD_1 As DataRow In data.Tables("WBTSTYLD").Select(filter)
            If rowWBTSTYLD_1.Item("STYLE_STATUS") = "A" Then
                hasValidColor = True
            Else
                Dim filterC As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                Dim rowSTATUS As DataRow = data.Tables.Item("STATUS").Select(filterC).FirstOrDefault
                If Not IsNothing(rowSTATUS) Then
                    Dim CURR_QTY_AVAIL As Int64 = Val(rowSTATUS.Item("CURR_QTY_AVAIL").ToString & String.Empty)
                    Dim FUT_QTY_AVAIL As Int64 = Val(rowSTATUS.Item("FUT_QTY_AVAIL").ToString & String.Empty)
                    Dim ALT_FUT_QTY As Int64 = Val(rowSTATUS.Item("ALT_FUT_QTY").ToString & String.Empty)
                    If (CURR_QTY_AVAIL + FUT_QTY_AVAIL + ALT_FUT_QTY) > 0 Then
                        hasValidColor = True
                    End If
                End If
            End If
        Next

        'Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND WEB_IND = 'W'", STYLE_CODE, COLOR_CODE)
        'Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND (CURR_ON_HAND + FUT_QTY_AVAIL) > 0", STYLE_CODE, COLOR_CODE)
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT SUM(AVL) AS AVL FROM (")
        SQLS.AppendLine("SELECT SUM(QTY_ATS) AS AVL")
        SQLS.AppendLine("FROM ICTSTDQ1")
        SQLS.AppendLine("WHERE WHSE_CODE = 'MS'")
        SQLS.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
        If Not isParent Then
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        End If
        SQLS.AppendLine("UNION")
        SQLS.AppendLine("SELECT SUM(NVL(ALT_FUT_QTY,0)) AS AVL")
        SQLS.AppendLine("FROM WBTSTYLD")
        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        If Not isParent Then
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        End If
        SQLS.AppendLine(")")
        ASCMAIN1.sql = SQLS.ToString()
        Dim AVL As Int64 = Val(ASCDATA1.GetDataValue)
        If AVL > 0 Then
            hasValidColor = True
        End If
        If hasValidColor Then
            If DISABLED_STYLES.Contains(STYLE_CODE) Then
                DISABLED_STYLES.Remove(STYLE_CODE)
            End If
        Else
            RetVal = "checked"
            If Not DISABLED_STYLES.Contains(STYLE_CODE) Then
                DISABLED_STYLES.Add(STYLE_CODE)
            End If
            'Dim DATE_DISABLED As String = rowWBTSTYLD.Item("DATE_DISABLED").ToString & String.Empty
            'Dim WEB_IND As String = rowWBTSTYLD.Item("WEB_IND").ToString & String.Empty
            'Dim STYLE_GROUP As String = rowWBTSTYLD.Item("STYLE_GROUP").ToString & String.Empty
            'If DATE_DISABLED.Length = 0 And WEB_IND = "W" And STYLE_GROUP <> "999" Then
            '    rowWBTSTYLD.Item("DATE_DISABLED") = Format(Now(), "MM/DD/YYYY")
            '    rowWBTSTYLD.Item("WEB_IND") = "I"
            '    rowWBTSTYLD.Item("STYLE_GROUP") = "999"
            'End If
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
