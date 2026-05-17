Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'*********************************
' Need to define Package Types
'Package Type	Description
'   1           Box
'   2           Polybag
'   3           Tube
'*********************************

' https://developer.dhl.com/api-catalog?f%5B0%5D=api_catalog_dhl_division%3A41&f%5B1%5D=api_catalog_region%3A45

' Vandale Credentials
'   Pickup # 5318184
'   Distribution Center: USORD1
'   Environment: Sandbox
'   API Version: v4
'   Client ID: vGSGRIIehJ96z2mCeVChblb8kqAHNm2T
'   Client Secret: kxsUu77CGYRuv2Ry

'INSERT INTO SOTCARR1
'(CARRIER_CODE, CARRIER_DESC, CARRIER_TYPE, CARRIER_REMOTE_HOST_IP, INIT_OPER, INIT_DATE, CARRIER_ARCHIVE_DIR, CARRIER_SHIP_TYPE, PROVIDER_TYPE, LABEL_FORMAT, CARRIER_PPA_TYPE, CARRIER_SURCHARGE_BASE)
'VALUES
'('DHL', 'DHL', 'U','https://api-sandbox.dhlecs.com', 'edz', SYSDATE, 'R:\SHIPMENTS\DHL', 'D', 'D', 'Z', 'L', 'L');

'INSERT INTO SOTCARR2
'(CARRIER_CODE, CARRIER_PROD_CODE, CARRIER_PROD_DESC, SERVICE_CODE, TRACKING_ID_TYPE, AVG_DELIVERY_DAYS)
'VALUES
'('DHL', '15', 'DHL Parcel Ground', 'D', '1', 5);


'INSERT INTO SOTCARR2
'(CARRIER_CODE, CARRIER_PROD_CODE, CARRIER_PROD_DESC, SERVICE_CODE, TRACKING_ID_TYPE, AVG_DELIVERY_DAYS)
'VALUES
'('DHL', '14', 'DHL Parcel Expedited', 'D', '1', 5);

'INSERT INTO SOTCARR3
'(CARRIER_CODE, DIVISION_CODE, CARRIER_ACCOUNT_NO, SHIPPER_DIVISION_CODE, SHIPPER_ID, CLIENT_ID, CLIENT_SECRET, SERVER_TOKEN_URL, SHIPPER_PASSWORD)
'VALUES
'('DHL', 'VAN', '5318184', 'SKINCOM', 'USORD1', 'vGSGRIIehJ96z2mCeVChblb8kqAHNm2T', 'kxsUu77CGYRuv2Ry', 'https://api-sandbox.dhlecs.com/auth/v4/accesstoken', 'DHL');

'INSERT INTO SOTCARR4
'(CARRIER_CODE, PACKAGE_CODE, PACKAGE_DESC)
'VALUES
'('DHL', '31', 'Our Packaging');

'INSERT INTO SOTSVIA1
'(SHIP_VIA_CODE, SHIP_VIA_DESC, SHIP_VIA_SCAC, CARRIER_CODE, CARRIER_PROD_CODE, SHIP_VIA_STATUS)
'VALUES
'('DHLEXP', 'DHL Parcel Expedited', 'DHLE', 'DHL', '14', 'A');

'INSERT INTO SOTSVIA1
'(SHIP_VIA_CODE, SHIP_VIA_DESC, SHIP_VIA_SCAC, CARRIER_CODE, CARRIER_PROD_CODE, SHIP_VIA_STATUS)
'VALUES
'('DHLGND', 'DHL Parcel Ground', 'DHLG', 'DHL', '15', 'I');

' ALTER TABLE SOTSHIP1 MODIFY SHIP_REF VARCHAR2(50);

' ALTER TABLE SOTCART1 MODIFY CART_TRACKING_NO VARCHAR2(50);

' ALTER TABLE WHTSHPC2 ADD PACKAGE_ID VARCHAR2(30);
' ALTER TABLE WHTSHPC2 ADD MANIFESTED DATE;
' ALTER TABLE WHTSHPC2 ADD REQUEST_ID VARCHAR2(30);

'CREATE TABLE WHTSHPCM( 
'CARRIER_CODE VARCHAR2(6),
'REQUEST_ID VARCHAR2(30),
'INIT_DATE DATE,
'INIT_OPER VARCHAR2(20),
'FILE_NAME VARCHAR2(300),
'PRIMARY KEY(CARRIER_CODE,REQUEST_ID));

Public Class WHCDHLS1

#Region "Variables / Properties"

    Private _AccessToken As String = String.Empty
    Private _TokenExpiration As DateTime = DateTime.Now.AddDays(-1)

    Public Property LastError As String
        Get
            Return clsLastError
        End Get
        Private Set(value As String)
            clsLastError = value
        End Set
    End Property

    Private clsLastError As String
    Public TestMode As Boolean = True

    ' Product Finder v4
    ' https://api-sandbox.dhlecs.com/shipping/v4/products
    Private Const productFinderSandboxUrl As String = "https://api-sandbox.dhlecs.com/shipping/v4/products"
    Private Const productFinderProductionUrl As String = "https://api.dhlecs.com/shipping/v4/products"

    ' AccessToken
    ' https://api-sandbox.dhlecs.com/auth/v4/accesstoken
    Private Const accessTokenSandboxUrl As String = "https://api-sandbox.dhlecs.com/auth/v4/accesstoken"
    Private Const accessTokenProductionUrl As String = "https://api.dhlecs.com/auth/v4/accesstoken"

    ' LabelRequest
    ' https://api-bat.dhlecs.com/shipping/v4/label
    Private Const labelRequestSandboxUrl As String = "https://api-sandbox.dhlecs.com/shipping/v4/label"
    Private Const labelRequestProductionUrl As String = "https://api.dhlecs.com/shipping/v4/label"

    ' Manifest
    Private Const manifestRequestSandboxUrl As String = "https://api-sandbox.dhlecs.com/shipping/v4/manifest"
    Private Const manifestRequestProductionUrl As String = "https://api.dhlecs.com/shipping/v4/manifest"

    Private tblWHTSHPC2 As DataTable = New DataTable

    Public Enum ServiceTypes
        GND = 15
        EXP = 14
    End Enum

#End Region

#Region "Public Properties"

    Public Property DHLClientID As String
    Public Property DHLClientSecret As String
    Public Property DHLDistributionCenter As String
    Public Property DHLPickupNumber As String

    Public Property ConsigneeAddress As New Address
    Public Property ReturnAddress As New Address
    Public Property PickupAddress As New Address
    Public Property ShipperAddress As New Address
    Public Property PackageDetails As List(Of PackageDetail)
    Public Property Rate As New RateRequestOptions
    Public Property EstimatedDeliveryOptions As New EddRequestOptions

    Public Property RequestedProduct As Int16

    Private clsRawRequest As String
    Public ReadOnly Property RawRequest As String
        Get
            Return clsRawRequest
        End Get
    End Property

    Private clsRawResponse As String
    Public ReadOnly Property RawResponse As String
        Get
            Return clsRawResponse
        End Get
    End Property

    Private clsLabelResponse As List(Of LabelResponse)
    Public ReadOnly Property ShippingingLabelResponse As List(Of LabelResponse)
        Get
            Return clsLabelResponse
        End Get
    End Property


#End Region

#Region "Instantiate Class"

    Public Sub New()

    End Sub

#End Region

#Region "Rates Request Classes"

    Private Class ProductFinderRequest
        Public Property pickup As String
        Public Property distributionCenter As String

        Public Property consigneeAddress As New Address
        Public Property returnAddress As New Address

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property pickupAddress As New Address
        Public Property shipperAddress As New Address

        Public Property packageDetail As PackageDetail

        Public Property rate As New RateRequestOptions
        Public Property estimatedDeliveryDate As New EddRequestOptions
    End Class

    Public Class Address
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property name As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property companyName As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property address1 As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property address2 As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property address3 As String

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property city As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property state As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property country As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property postalCode As String

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property email As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property phone As String
    End Class

    Public Class PackageDetail
        Public Property packageId As String

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property packageDescription As String

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property packageReference As String

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property orderSource As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property contentCategory As String

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property billingReference1 As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property billingReference2 As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property customLabelText1 As String
        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property customLabelText2 As String

        Public Property weight As New Weight
        Public Property dimension As New Dimension

        Public Property shippingCost As New ShippingCost

        <JsonIgnore>
        Public Property ShippingRate As Decimal
        <JsonIgnore>
        Public Property TrackingNo As String
        <JsonIgnore>
        Public Property ShippingLabel As String
        <JsonIgnore>
        Public Property PackageDetailRates As New List(Of Product)
        <JsonIgnore>
        Public Property ezshipPackageDetail As New DPayments.DShippingSDK.PackageDetail
        <JsonIgnore>
        Public Property cartKey As String
    End Class

    Public Class Weight
        Public Property value As Decimal
        Public Property unitOfMeasure As String
    End Class

    Public Class Dimension
        Public Property length As Decimal
        Public Property width As Decimal
        Public Property height As Decimal

        Public Property unitOfMeasure As String
        Public Property packageType As String
    End Class

    Public Class ShippingCost
        Public Property currency As String
        Public Property freight As Decimal
        Public Property declaredValue As Decimal
        Public Property insuredValue As Decimal
        Public Property duty As Decimal

        Public Property dutiesPaid As Boolean
        Public Property taxesPaid As Boolean
        Public Property tax As Decimal
    End Class

    Public Class RateRequestOptions
        Public Property calculate As Boolean = True
        Public Property currency As String = "USD"
        Public Property rateDate As String = DateTime.Now.ToString("yyyy-MM-dd")
        Public Property maxPrice As Decimal = 99999
    End Class

    Public Class EddRequestOptions
        Public Property calculate As Boolean = True
        Public Property deliveryBy As String = DateTime.Now.AddDays(14).ToString("yyy-MM-dd")
        Public Property expectedShipDate As String = DateTime.Now.ToString("yyy-MM-dd")
        Public Property expectedTransit As Integer = 14
    End Class

#End Region

#Region "Rates Response Classes"

    Public packageRates As New List(Of PackageDetail)

    Public Class ProductFinderResponse
        Public Property pickup As String
        Public Property distributionCenter As String
        Public Property shipDate As String

        Public Property consigneeAddress As Address
        Public Property returnAddress As Address
        Public Property shipperAddress As Address

        Public Property packageDetail As PackageDetail
        Public Property rate As RateRequestOptions
        Public Property estimatedDeliveryDate As EstimatedDeliveryRequest

        Public Property products As List(Of Product)
    End Class

    Public Class EstimatedDeliveryRequest
        Public Property calculate As Boolean
        Public Property expectedShipDate As String
        Public Property expectedTransit As Integer
        Public Property deliveryBy As String
    End Class

    Public Class Product
        Public Property orderedProductId As String
        Public Property productName As String
        Public Property description As String
        Public Property trackingAvailable As String

        Public Property rate As ProductRate
        Public Property estimatedDeliveryDate As ProductEstimatedDelivery

    End Class

    Private dictRates As New Dictionary(Of String, List(Of Product))

    Public Class ProductRate
        Public Property priceZone As String
        Public Property amount As Decimal
        Public Property currency As String
        Public Property effectiveFrom As String

        Public Property rateComponents As List(Of RateComponent)
    End Class

    Public Class RateComponent
        Public Property rateComponentId As String
        Public Property partId As String
        Public Property description As String
        Public Property amount As Decimal
    End Class

    Public Class ProductEstimatedDelivery
        Public Property isGuaranteed As Boolean
        Public Property deliveryDaysMin As Integer
        Public Property deliveryDaysMax As Integer
        Public Property estimatedDeliveryMin As String
        Public Property estimatedDeliveryMax As String
    End Class

#End Region

#Region "Label Request Classes"

    Public Class LabelRequest

        Public Property pickup As String

        Public Property distributionCenter As String

        Public Property orderedProductId As String

        Public Property consigneeAddress As Address

        Public Property returnAddress As Address

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property shipperAddress As Address

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property pickupAddress As Address

        Public Property packageDetail As PackageDetail

        <JsonProperty(NullValueHandling:=NullValueHandling.Ignore)>
        Public Property customsDetails As List(Of CustomsDetail)

    End Class

    Public Class CustomsDetail

        Public Property itemDescription As String

        Public Property countryOfOrigin As String

        Public Property hsCode As String

        Public Property packagedQuantity As Integer

        Public Property skuNumber As String

        Public Property itemValue As Decimal

        Public Property currency As String

        Public Property productUrl As String

    End Class

#End Region

#Region "Label Response Classes"

    Public Class LabelResponse

        Public Property timestamp As DateTime

        Public Property pickup As String

        Public Property distributionCenter As String

        Public Property orderedProductId As String

        Public Property labels As List(Of LabelItem)

    End Class

    Public Class LabelItem

        Public Property createdOn As DateTime

        Public Property packageId As String

        Public Property dhlPackageId As String

        Public Property trackingId As String

        Public Property labelData As String

        Public Property encodeType As String

        Public Property format As String

        Public Property link As String

        Public Property labelDetail As LabelDetail

    End Class

    Public Class LabelDetail

        Public Property serviceLevel As String

        Public Property outboundSortCode As String

        Public Property sortingSetupVersion As String

        Public Property inboundSortCode As String

        Public Property serviceEndorsement As String

        Public Property intendedReceivingFacility As String

        Public Property mailBanner As String

        Public Property customsDetailsProvided As Boolean

    End Class

#End Region

#Region "Manifest Classes"
    Public Class CreateManifestRequest
        <JsonProperty("pickup")>
        Public Property Pickup As String

        <JsonProperty("products")>
        Public Property Products As List(Of String)

        <JsonProperty("manifests")>
        Public Property Manifests As List(Of ManifestRequest)
    End Class

    Public Class ManifestRequest
        <JsonProperty("packageIds")>
        Public Property PackageIds As List(Of String)

        <JsonProperty("dhlPackageIds")>
        Public Property DhlPackageIds As List(Of String)
    End Class
    Public Class ManifestResponse
        Public Property requestId As String
        Public Property status As String
        Public Property pickup As String
    End Class

    Public Class ManifestSummary
        <JsonProperty("total")>
        Public Property Total As Integer
    End Class

    Public Class ManifestDocumentResponse
        Public Property status As String
        Public Property documents As List(Of ManifestDocument)
    End Class

    Public Class ManifestDownloadResponse
        Public Property status As String
        Public Property documents As List(Of ManifestDocument)
    End Class

    Public Class ManifestDocument
        Public Property format As String
        Public Property content As String
    End Class
#End Region

#Region "Token Classes"

    Public Class TokenResponse
        Public Property access_token As String
        Public Property client_id As String
        Public Property token_type As String
        Public Property expires_in As Integer
    End Class

#End Region

#Region "Private Procedures"

    Private Function GetAccessToken() As String

        If DateTime.Now < _TokenExpiration Then
            If _AccessToken.Length > 0 Then
                Return _AccessToken
            End If
        End If

        Dim url As String = IIf(TestMode, accessTokenSandboxUrl, accessTokenProductionUrl)

        Using client As New HttpClient

            Dim values = New List(Of KeyValuePair(Of String, String)) From {
                New KeyValuePair(Of String, String)(
                    "grant_type",
                    "client_credentials"
                ),
                New KeyValuePair(Of String, String)(
                    "client_id",
                    DHLClientID
                ),
                New KeyValuePair(Of String, String)(
                    "client_secret",
                    DHLClientSecret
                )
            }

            Dim content As New FormUrlEncodedContent(values)
            Dim response = client.PostAsync(url, content).Result
            Dim responseBody As String = response.Content.ReadAsStringAsync().Result

            If response.IsSuccessStatusCode Then
                Dim token = JsonConvert.DeserializeObject(Of TokenResponse)(responseBody)
                _AccessToken = token.access_token
                _TokenExpiration = DateTime.Now.AddSeconds(token.expires_in - 60)
            Else

            End If

        End Using

        Return _AccessToken

    End Function

#End Region

#Region "ProductFinder Procedures"

    Public Function GetRates() As List(Of Product)

        Dim url As String = IIf(TestMode, productFinderSandboxUrl, productFinderProductionUrl)
        Dim lstProduct As New List(Of Product)
        dictRates.Clear()
        packageRates.Clear()

        clsRawRequest = String.Empty
        clsRawResponse = String.Empty

        Using client As New HttpClient

            Dim accessToken As String = GetAccessToken()
            If String.IsNullOrWhiteSpace(accessToken) Then
                LastError = "Unable to obtain access token."
                Return lstProduct
            End If

            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)

            For Each pkg As PackageDetail In PackageDetails
                Dim productFinderRequest As ProductFinderRequest = CreateProductFinderRequest(pkg)
                Dim json = JsonConvert.SerializeObject(productFinderRequest)
                clsRawRequest = json
                Dim content = New StringContent(json, Encoding.UTF8, "application/json")
                Dim response = client.PostAsync(url, content).Result
                Dim responseBody = response.Content.ReadAsStringAsync().Result
                clsRawResponse = responseBody

                If response.IsSuccessStatusCode Then
                    Dim responseObject = JsonConvert.DeserializeObject(Of ProductFinderResponse)(responseBody)
                    Dim rateResponse = JsonConvert.DeserializeObject(Of ProductFinderResponse)(responseBody)
                    If rateResponse IsNot Nothing AndAlso rateResponse.products IsNot Nothing Then
                        ' Need to loop through Products and create a rate object
                        Dim lstProd As New List(Of Product)
                        For Each prod As Product In rateResponse.products
                            lstProd.Add(prod)
                        Next
                        pkg.PackageDetailRates = lstProd
                        dictRates.Add(pkg.packageId, lstProd)
                        packageRates.Add(pkg)
                    End If
                Else
                    Return lstProduct
                End If
            Next
        End Using

        If dictRates.Count = 1 Then
            lstProduct = dictRates.Values.First()
            Return lstProduct
        End If

        ' Find common orderedProductIds
        Dim commonIds As HashSet(Of String) = Nothing

        For Each productList In dictRates.Values
            Dim currentIds As New HashSet(Of String)(
                productList.
                Where(Function(p) p IsNot Nothing).
                Select(Function(p) p.orderedProductId)
            )

            If commonIds Is Nothing Then
                commonIds = currentIds
            Else
                commonIds.IntersectWith(currentIds)
            End If
        Next

        ' Build final product list
        For Each id In commonIds

            ' Grab all matching products
            Dim matchingProducts =
            dictRates.Values.
            SelectMany(Function(list) list).
            Where(Function(p) p.orderedProductId = id).
            ToList()

            ' Use first product as base object
            Dim baseProduct As Product = matchingProducts.First()

            ' Sum all rate amounts
            Dim totalAmount As Decimal =
            matchingProducts.
            Where(Function(p) p.rate IsNot Nothing).
            Sum(Function(p) p.rate.amount)

            ' Create final product
            Dim finalProduct As New Product With {
            .orderedProductId = baseProduct.orderedProductId,
            .productName = baseProduct.productName,
            .description = baseProduct.description,
            .trackingAvailable = baseProduct.trackingAvailable,
            .estimatedDeliveryDate = baseProduct.estimatedDeliveryDate,
            .rate = New ProductRate With {
                .priceZone = baseProduct.rate.priceZone,
                .currency = baseProduct.rate.currency,
                .effectiveFrom = baseProduct.rate.effectiveFrom,
                .rateComponents = baseProduct.rate.rateComponents,
                .amount = totalAmount
                }
            }
            lstProduct.Add(finalProduct)
        Next

        Return lstProduct

    End Function

    Private Function CreateProductFinderRequest(inPackageDetail As PackageDetail) As ProductFinderRequest

        Dim requestObject As New ProductFinderRequest With {
            .pickup = DHLPickupNumber,
            .distributionCenter = DHLDistributionCenter,
            .consigneeAddress = ConsigneeAddress,
            .returnAddress = ReturnAddress,
            .shipperAddress = ShipperAddress,
            .pickupAddress = Nothing,
            .packageDetail = inPackageDetail,
            .rate = Rate,
            .estimatedDeliveryDate = EstimatedDeliveryOptions
            }
        Return requestObject

    End Function

#End Region

#Region "Label Creation Procedures"

    Public Function RequestLabel() As Boolean

        Dim url As String = IIf(TestMode, labelRequestSandboxUrl, labelRequestProductionUrl) & "?format=ZPL"

        clsRawRequest = String.Empty
        clsRawResponse = String.Empty

        clsLabelResponse = New List(Of LabelResponse)
        Dim accessToken As String = GetAccessToken()
        If String.IsNullOrWhiteSpace(accessToken) Then
            LastError = "Unable to obtain access token."
            Return False
        End If

        Dim pkgLabelResponse As New LabelResponse

        Dim ipackageCounter As Int16 = 1
        Using client As New HttpClient
            For Each pkg In PackageDetails
                pkg.customLabelText2 = $"Carton {ipackageCounter} of {PackageDetails.Count}"
                ipackageCounter += 1
                Dim labelRequest As LabelRequest = CreateLabelRequest(pkg)

                client.DefaultRequestHeaders.Clear()
                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)

                client.DefaultRequestHeaders.Accept.Clear()
                client.DefaultRequestHeaders.Add("User-Agent", "ABSolution Label Request")

                Dim settings As New JsonSerializerSettings With {
                        .NullValueHandling = NullValueHandling.Ignore
                        }

                Dim requestJson As String = JsonConvert.SerializeObject(labelRequest, Formatting.Indented, settings)
                clsRawRequest = requestJson
                Dim content As New StringContent(requestJson, Encoding.UTF8, "application/json")

                Dim response As HttpResponseMessage = client.PostAsync(url, content).Result
                Dim responseJson As String = response.Content.ReadAsStringAsync().Result

                Dim parsedJson = JToken.Parse(responseJson)
                Dim formatted As String = parsedJson.ToString(Formatting.Indented)
                responseJson = formatted

                clsRawResponse = responseJson

                If Not response.IsSuccessStatusCode Then
                    LastError = $"HTTP {(CInt(response.StatusCode))}" & Environment.NewLine & responseJson
                    Return False
                End If

                pkgLabelResponse = JsonConvert.DeserializeObject(Of LabelResponse)(responseJson)
                pkg.TrackingNo = pkgLabelResponse.labels(0).trackingId
                pkg.ShippingLabel = pkgLabelResponse.labels(0).labelData
                pkg.ShippingRate = 0
                Dim orderedProductId As String = CType(RequestedProduct, ServiceTypes).ToString()

                ' See if we can match the carton dimension to a rates dimension
                If packageRates IsNot Nothing Then
                    For Each rp As PackageDetail In packageRates
                        If rp.cartKey = pkg.cartKey Then
                            For Each pkgRate As Product In rp.PackageDetailRates
                                If pkgRate.orderedProductId = orderedProductId Then
                                    rp.ShippingRate = pkgRate.rate.amount
                                    rp.packageId = pkg.packageId
                                    rp.ezshipPackageDetail.Id = pkg.packageId
                                    rp.cartKey = pkg.cartKey
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                End If
            Next

            Return True
        End Using

    End Function

    Private Function CreateLabelRequest(inPackageDetail As PackageDetail) As LabelRequest

        Dim orderedProductId As String = CType(RequestedProduct, ServiceTypes).ToString()

        Dim requestObject As New LabelRequest With {
            .pickup = DHLPickupNumber,
            .distributionCenter = DHLDistributionCenter,
            .consigneeAddress = ConsigneeAddress,
            .returnAddress = ReturnAddress,
            .shipperAddress = ShipperAddress,
            .pickupAddress = Nothing,
            .packageDetail = inPackageDetail,
            .orderedProductId = orderedProductId
            }

        Return requestObject

    End Function

#End Region

#Region "Manifest Request Procedures"

    Public Function RequestManifest() As Boolean

        tblWHTSHPC2 = New DataTable
        Dim url As String = IIf(TestMode, manifestRequestSandboxUrl, manifestRequestProductionUrl)

        clsRawRequest = String.Empty
        clsRawResponse = String.Empty

        Dim accessToken As String = GetAccessToken()
        If String.IsNullOrWhiteSpace(accessToken) Then
            LastError = "RequestManifest Error: Unable to obtain access token."
            Return False
        End If

        Dim manifestRequest As CreateManifestRequest = GenerateManifestRequest()
        If manifestRequest Is Nothing Then
            Return True
        End If
        Dim mResponse As New ManifestResponse

        Using client As New HttpClient
            client.DefaultRequestHeaders.Clear()
            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
            client.DefaultRequestHeaders.Add("User-Agent", "ABSolution Manifest Request")

            Dim settings As New JsonSerializerSettings With {
                        .NullValueHandling = NullValueHandling.Ignore
                        }

            Dim requestJson As String = JsonConvert.SerializeObject(manifestRequest, Formatting.Indented, settings)
            clsRawRequest = requestJson
            Dim content As New StringContent(requestJson, Encoding.UTF8, "application/json")

            Dim response = client.PostAsync(url, content).Result
            Dim responseBody = response.Content.ReadAsStringAsync().Result

            If response.IsSuccessStatusCode Then
                ' Do something
                mResponse = JsonConvert.DeserializeObject(Of ManifestResponse)(responseBody)
                If Not RequestManifestPDF(mResponse) Then
                    Return False
                End If
            Else
                LastError = $"HTTP {(CInt(response.StatusCode))}" & Environment.NewLine & responseBody
                Return False
            End If

            Return True

        End Using

    End Function

    Private Function GenerateManifestRequest() As CreateManifestRequest

        Dim sql As String = "SELECT WHTSHPC2.* 
                 FROM WHTSHPC1, WHTSHPC2
                 WHERE WHTSHPC1.SHIP_CNTL_NO = WHTSHPC2.SHIP_CNTL_NO
                 AND WHTSHPC1.CARRIER_CODE = 'DHL'
                 AND WHTSHPC2.PACKAGE_ID IS NOT NULL
                 AND NVL(WHTSHPC2.MANIFESTED, '0') = '0'
                 AND WHTSHPC2.REQUEST_ID IS NULL "

        tblWHTSHPC2 = ASCDATA1.GetDataTable(sql, "WHTSHPC2")

        If tblWHTSHPC2.Rows.Count = 0 Then
            Return Nothing
        End If

        Dim packageIds As New List(Of String)

        For Each dr As DataRow In tblWHTSHPC2.Rows
            Dim packageId = dr("PACKAGE_ID").ToString().Trim()
            packageIds.Add(packageId)
        Next

        Dim request As New CreateManifestRequest With {
            .Pickup = DHLPickupNumber,
            .Manifests = New List(Of ManifestRequest) From {
                New ManifestRequest With {
                    .PackageIds = packageIds
                }
            }
        }

        Return request

    End Function

    Private Function RequestManifestPDF(mResponse As ManifestResponse) As Boolean

        Dim requestId As String = mResponse.requestId
        Dim pickupId As String = mResponse.pickup
        Dim Status As String = mResponse.status
        RequestManifestPDF = True

        Dim ManifestFileName As String = String.Empty
        If Not DownloadManifestPdf(requestId, pickupId, ManifestFileName) Then
            Return False
        End If

        If Not My.Computer.FileSystem.FileExists(ManifestFileName) Then
            LastError = $"Cannot locate DHL Manifest: {ManifestFileName}"
            Return False
        End If

        Try
            Process.Start(New ProcessStartInfo With {
                                    .FileName = ManifestFileName,
                                    .UseShellExecute = True
                                })
        Catch ex As Exception
            LastError = $"Error trying to display Manifest PDF file ({ManifestFileName}): {ex.Message}"
            RequestManifestPDF = False
        End Try

        Dim lstPackageIds As New List(Of String)
        For Each dr In tblWHTSHPC2.Select("")
            lstPackageIds.Add(dr.Item("PACKAGE_ID"))
        Next

        Dim sql As String = $"UPDATE WHTSHPC2 
            SET MANIFESTED = '1', 
            REQUEST_ID = {requestId}
            WHERE PACKAGE_ID IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"

        ASCDATA1.ExecuteSQL(sql, "V", String.Join(",", lstPackageIds.ToArray))

    End Function

    Private Function DownloadManifestPdf(requestId As String, pickupId As String, ByRef ManifestFileName As String) As Boolean

        Try
            Dim url As String = IIf(TestMode, manifestRequestSandboxUrl, manifestRequestProductionUrl) & $"/{pickupId}/{requestId}"

            Dim CARRIER_ARCHIVE_DIR As String = String.Empty
            Dim drSOTCARR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTCARR1 WHERE CARRIER_CODE = 'DHL'")
            If drSOTCARR1 Is Nothing Then
                LastError = "Cannot locate master record for DHL"
                Return False
            End If

            CARRIER_ARCHIVE_DIR = (drSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            If CARRIER_ARCHIVE_DIR.Length = 0 Then '
                LastError = "DHL is not assigned a Carrier Archive Directory"
                Return False
            End If

            If ASCMAIN1.Running_in_VS Then
                Stop
                CARRIER_ARCHIVE_DIR = CARRIER_ARCHIVE_DIR.Replace("R:\", "C:\")
                CARRIER_ARCHIVE_DIR = CARRIER_ARCHIVE_DIR.Replace("S:\", "C:\")
            End If

            If Not My.Computer.FileSystem.DirectoryExists(CARRIER_ARCHIVE_DIR) Then
                My.Computer.FileSystem.CreateDirectory(CARRIER_ARCHIVE_DIR)
            End If


            Dim accessToken As String = GetAccessToken()
            If String.IsNullOrWhiteSpace(accessToken) Then
                LastError = "DownloadManifestPdf Error: Unable to obtain access token."
                Return False
            End If

            Using client As New HttpClient()

                client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", accessToken)
                Dim response = client.GetAsync(url).Result
                Dim responseBody As String = response.Content.ReadAsStringAsync().Result
                If Not response.IsSuccessStatusCode Then
                    LastError = responseBody
                    Return False
                End If

                Dim manifestResult = JsonConvert.DeserializeObject(Of ManifestDocumentResponse)(responseBody)

                If manifestResult Is Nothing Then
                    LastError = "Empty manifest response."
                    Return False
                End If

                If manifestResult.status <> "COMPLETED" Then
                    LastError = $"Manifest status: {manifestResult.status}"
                    Return False
                End If

                Dim base64Pdf As String = manifestResult.documents(0).content
                Dim pdfBytes() As Byte = Convert.FromBase64String(base64Pdf)
                ManifestFileName = Path.Combine(CARRIER_ARCHIVE_DIR, $"DHLManifest_{requestId}.pdf")
                File.WriteAllBytes(ManifestFileName, pdfBytes)

                Return True

            End Using

        Catch ex As Exception
            LastError = $"DownloadManifestPdf Error: {ex.Message}"
            Return False
        End Try

    End Function

#End Region

End Class
