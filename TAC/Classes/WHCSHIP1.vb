Imports nsoftware.InShip
Imports System.Net
Imports System.Text
Imports System.IO
Imports Newtonsoft.Json

' Needed for USPS Pitney Bowes
'alter table SOTCARR1 add USPS_PARTNER VARCHAR2(1);
'alter table SOTCARR3 modify ACCESSLICENSENUMBER VARCHAR2(100);
'alter table SOTCARR3 modify ADDR_VALID_LICENSE VARCHAR2(100);
'alter table SOTCARR3 modify METER_NUMBER VARCHAR2(500);
'alter table SOTCARR3 modify SHIPPER_ID VARCHAR2(100);
'alter table SOTCARR3 add TOKEN_EXPIRES DATE;

Public Class WHCSHIP1

#Region "Variables"

    Private objEzShip As nsoftware.InShip.Ezship
    Private objEzRates As nsoftware.InShip.Ezrates

    Private objFedexShip As nsoftware.InShip.Fedexship
    Private objFedexShipIntl As nsoftware.InShip.Fedexshipintl
    Private objFedexRates As nsoftware.InShip.Fedexrates
    Private objFedexTrack As nsoftware.InShip.Fedextrack

    Private objUpsShip As nsoftware.InShip.Upsship
    Private objUpsShipIntl As nsoftware.InShip.Upsshipintl
    Private objUpsRates As nsoftware.InShip.Upsrates
    Private objUpsTrack As nsoftware.InShip.Upstrack

    Private objUpsFreight As nsoftware.InShip.Upsfreightship
    Private objUpsFreightRates As nsoftware.InShip.Upsfreightrates

    Private objUspsShip As nsoftware.InShip.Uspsship
    Private objUspsRates As nsoftware.InShip.Uspsrates
    Private objUspsTrack As nsoftware.InShip.Uspstrack

    Public Enum ServiceProviders
        FederalExpress = EzshipProviders.pFedEx
        UPS = EzshipProviders.pUPS
        USPS = EzshipProviders.pUSPS
        CanadaPost = EzshipProviders.pCanadaPost
        UPSFreight = 4
        FederalExpressInternational = 5
        UPSInternational = 6
        Unknown = 7
    End Enum

    Public Enum USPSPostageProviders
        None = 0
        Endicia = 1
        StampsCom = 2
        PitneyBowes = 3
    End Enum

    Public USPSPostageProvider As USPSPostageProviders = USPSPostageProviders.None

    Public DropOffType As FedexshipintlDropoffTypes = FedexshipintlDropoffTypes.dtRegularPickup
    Public EzshipLabelImage As EzshipLabelImageTypes = EzshipLabelImageTypes.itZebra
    Public UPSPickupType As UpsratesPickupTypes = UpsratesPickupTypes.ptDailyPickup
    Private cCustomerType As UpsratesCustomerTypes = UpsratesCustomerTypes.ccDaily

    Public ShippingLabelDirectory As String = String.Empty
    Public ShippingLabelPrefix As String = String.Empty
    Public PackageDetailList As New List(Of nsoftware.InShip.PackageDetail)

    Public ShipDate As Date = DateTime.Now
    Public ShipmentSpecialServices As Long = 0
    Public CommodityDetailList As New List(Of nsoftware.InShip.CommodityDetail)
    Public RequestedServicesRates As New List(Of ServiceDetail)
    Public HandlingUnit As String = String.Empty

    Public USPSTestMode As Boolean = True

    Private cServiceProvider As ServiceProviders = ServiceProviders.Unknown
    Private cServer As String = String.Empty
    Private cUserId As String = String.Empty
    Private cPassword As String = String.Empty
    Private cIntegrationID As String = String.Empty

    Private cOAuthKey As String = String.Empty
    Private cOAuthSecret As String = String.Empty
    Private cOAuthUrl As String = String.Empty

    Private cRequestedServiceType As ServiceTypes
    Private cAccountNumber As String = String.Empty
    Private cTotalCustomsValue As Decimal = 0
    Private cSignatureRequired As Boolean = False

    Private cFedexDeveloperKey As String = String.Empty
    Private cFedexMeterNumber As String = String.Empty
    Private cLabelStockType As String = String.Empty

    Private cUPSAccessKey As String = String.Empty

    Private cUSPSEndiciaCustomerId As String = String.Empty
    Private cUSPSEndiciaTransactionId As String = String.Empty

    Private cSenderContact As New Contact
    Private cAccountContact As New Contact
    Private cRecipientContact As New Contact
    Private cReturnAddress As New Contact
    Private cPayorContact As New Contact
    Private cDutiesPayorContact As New Contact
    Private cHoldAtLocation As New Contact
    Private cFedexSmartPost As New SmartPost

    Public Payor As TPayorTypes = TPayorTypes.ptSender
    Public DutiesPayor As TPayorTypes = TPayorTypes.ptRecipient

    Public LastError As String = String.Empty
    Private cMasterTrackingNumber As String = String.Empty
    Private cRawRequest As String = String.Empty
    Private cRawResponse As String = String.Empty
    Private inShipLicense As String = ASCMAIN1.nSoftwareKeys("nSoftwareInship")

    Public FedexClose As New CloseDetail

    Public Const ProviderTypeFedex = "F"
    Public Const ProviderTypeUPS = "U"
    Public Const ProviderTypeUSPS = "P"
    Public Const ProviderTypeCanada = "C"

    Public ShipmentBaseCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentDiscountCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentListCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentNetCharge As New Dictionary(Of Int32, Decimal)
    Public ShipmentSurCharge As New Dictionary(Of Int32, Decimal)

    Private cFedexCustomContent As String = String.Empty
    Private Const SSLEnabledProtocols As Int32 = 4032

    Public cRatesTotalValue As Decimal = 0
    Private cShipmentDescription As String = String.Empty

    Private clsInternationalFormsFile As String = String.Empty

    Public Structure CommercialInvoice
        Dim Purpose As CommercialInvoicePurposes
        Dim Terms As CommercialInvoiceTerms
        Dim Comments As String
        Dim InvoiceDate As String
        Dim CustomersInvoiceNumber As String
        Dim ShipperInsurance As Decimal
        Dim FreightCharge As Decimal
    End Structure

    Public Structure ShippersExportDeclaration
        Dim InBond As TInBondCodes
        Dim LicenseExceptionCode As TExceptionCodes
        Dim LicenseNumber As String
        Dim LicenseDate As String
        Dim ImportEntryNumber As String
        Dim PointOfOrigin As String
        Dim ShippersTaxID As String
        Dim TransPortType As String
        Dim ExportingCarrier As String
        Dim ExportingDate As String
    End Structure

    Public Structure UPSInternationalForms
        Dim CommercialInvoice As Boolean
        Dim ShippersExportDeclaration As Boolean
        Dim CertificateOfOrigin As Boolean
        Dim NAFTACertificateOfOrigin As Boolean

        Dim CertificateOfOriginExportingCarrier As String
        Dim CertificateOfOriginExportDate As String

        Dim NAFTABlanketPeriodStartDate As String
        Dim NAFTABlanketPeriodEndDate As String

        Dim CommercialInvoiceInfo As CommercialInvoice
        Dim ShippersExportDeclarationInfo As ShippersExportDeclaration

    End Structure

    Public Enum NotifictaionTypes
        On_Shipment = 1
        On_Exception = 2
        On_Deleivery = 4
        'On_Tender = 8
        'On_Return_UPS = 10 ' Not Used by FedEx
        'HTML_FedEx = 20 ' Not Used by UPS
        'Text_Fedex = 40  ' Not Used by UPS
        'Wireless_Fedex = 80 ' Not Used by UPS
    End Enum

    Public Class Notifications
        Public email As String = String.Empty
        Public NotificationFlags As NotifictaionTypes
        Public Message As String = String.Empty
    End Class

    Public ShipmentNotifications As New List(Of Notifications)

    Public Enum GroundFreightPayor
        Prepaid = 0
        FreightCollect = 1
        Prepaid_Thirdparty = 2
    End Enum

    Public Class GroundFreightPayorClass
        Public PayorType As GroundFreightPayor = GroundFreightPayor.Prepaid
        Public AccountNumber As String = String.Empty
        Public CountryCode As String = String.Empty
        Public ZipCode As String = String.Empty
    End Class

    Public UPSGroundFreightPayor As New GroundFreightPayorClass

    Public Enum ShipperRates
        List = 0
        Account = 1
    End Enum

    ' Test Regions
    'https://gatewaybeta.fedex.com:443/xml
    'https://wsbeta.fedex.com:443/web-services
    'https://wwwcie.ups.com/ups.app/xml
    'https://swsim.testing.stamps.com/swsim/SwsimV45.asmx
    'https://ct.soa-gw.canadapost.ca" 'development server
    ' endicia Server for PPS

    ' Production Regions 
    ' https://ws.fedex.com:443/xml
    ' https://onlinetools.ups.com/ups.app/xml/

    Public Structure RateList
        Dim ServiceType As String
        Dim ServiceTypeDescription As String
        Dim TransitTime As String
        Dim DeliveryTime As String
        Dim AccountNetCharge As Decimal
        Dim ListNetCharge As Decimal
        Dim ReferenceIndex As Int16
        Dim OfferID As String
        Dim ServiceCode As String
        Dim DeliveryDate As String
    End Structure

    Public UPSFreightCharges As Dictionary(Of String, Decimal)
    Public Const UPSFreightProductCode As String = "4343" ' nsoftware.InShip.ServiceTypes.stUPSGround
    Public UPSFreightBOLID As String = String.Empty
    Public UPSFreightShipmentNumber As String = String.Empty
    Public UPSFreightLabels As New List(Of String)

    Public RequestedUPSInternationalForms As UPSInternationalForms
    Private cInternationalForms As String = String.Empty

    Public Const UPSnternationalFormsExtension As String = "_INTL.PDF"

    Public Class TrackingData
        Public Status As String = String.Empty
        Public [Date] As String = String.Empty
        Public Time As String = String.Empty
        Public Address1 As String = String.Empty
        Public Address2 As String = String.Empty
        Public City As String = String.Empty
        Public State As String = String.Empty
        Public ZipCode As String = String.Empty
        Public CountryCode As String = String.Empty
        Public Location As String = String.Empty
    End Class

    Private clsTrackingData As New TrackingData

    Public ReadOnly Property TrackingInfo As TrackingData
        Get
            Return clsTrackingData
        End Get
    End Property

#End Region

#Region "Instantiate Class"

    Public Sub New()
        InitializeVariables()
    End Sub

    Public Sub New(ByVal ServiceType As ServiceProviders)
        InitializeVariables()
        cServiceProvider = ServiceType
    End Sub

    ''' <summary>
    ''' Set all Objects to the default values
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Initialize class objects
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeVariables()

        objEzShip = Nothing
        objEzRates = Nothing

        objFedexShip = Nothing
        objFedexShipIntl = Nothing
        objFedexRates = Nothing

        objUpsShip = Nothing
        objUpsShipIntl = Nothing
        objUpsRates = Nothing
        objUpsTrack = Nothing

        objUpsFreight = Nothing
        objUpsFreightRates = Nothing

        objUspsShip = Nothing
        objUspsRates = Nothing
        objUspsTrack = Nothing

        FedexClose = New CloseDetail

        DropOffType = FedexshipintlDropoffTypes.dtRegularPickup
        'PackageDetail = New nsoftware.InShip.PackageDetail
        EzshipLabelImage = EzshipLabelImageTypes.itZebra
        ShippingLabelDirectory = String.Empty
        ShippingLabelPrefix = String.Empty
        HandlingUnit = String.Empty
        ShipDate = DateTime.Now
        ShipmentSpecialServices = 0
        CommodityDetailList = New List(Of nsoftware.InShip.CommodityDetail)
        USPSPostageProvider = USPSPostageProviders.None

        ShipmentBaseCharge.Clear()
        ShipmentDiscountCharge.Clear()
        ShipmentListCharge.Clear()
        ShipmentNetCharge.Clear()
        ShipmentSurCharge.Clear()

        cServiceProvider = ServiceProviders.Unknown
        cServer = String.Empty
        cUserId = String.Empty
        cPassword = String.Empty
        cIntegrationID = String.Empty

        cRequestedServiceType = New ServiceTypes
        cAccountNumber = String.Empty
        cLabelStockType = String.Empty

        cFedexDeveloperKey = String.Empty
        cFedexMeterNumber = String.Empty

        cUPSAccessKey = String.Empty

        cUSPSEndiciaCustomerId = String.Empty
        cUSPSEndiciaTransactionId = String.Empty
        cTotalCustomsValue = 0
        cMasterTrackingNumber = String.Empty

        cSenderContact = New Contact
        cRecipientContact = New Contact
        cPayorContact = New Contact
        cDutiesPayorContact = New Contact
        cHoldAtLocation = New Contact
        cReturnAddress = New Contact
        cAccountContact = New Contact

        cFedexSmartPost = New SmartPost

        Payor = TPayorTypes.ptSender
        DutiesPayor = TPayorTypes.ptRecipient

        cRawRequest = String.Empty
        cRawResponse = String.Empty
        cFedexCustomContent = String.Empty

        PackageDetailList = New List(Of nsoftware.InShip.PackageDetail)
        cSignatureRequired = False

        ShipmentNotifications = New List(Of Notifications)
        UPSFreightCharges = New Dictionary(Of String, Decimal)
        UPSGroundFreightPayor = New GroundFreightPayorClass

        cInternationalForms = String.Empty

        With RequestedUPSInternationalForms
            .CertificateOfOrigin = False
            .CommercialInvoice = False
            .NAFTACertificateOfOrigin = False
            .ShippersExportDeclaration = False

            .CertificateOfOriginExportDate = String.Empty
            .CertificateOfOriginExportingCarrier = String.Empty

            .NAFTABlanketPeriodStartDate = String.Empty
            .NAFTABlanketPeriodEndDate = String.Empty

            With .CommercialInvoiceInfo
                .Comments = String.Empty
                .CustomersInvoiceNumber = String.Empty
                .FreightCharge = 0
                .InvoiceDate = String.Empty
                .Purpose = CommercialInvoicePurposes.cipGift
                .ShipperInsurance = 0
                .Terms = CommercialInvoiceTerms.citCpt
            End With

            With .ShippersExportDeclarationInfo
                .ImportEntryNumber = String.Empty
                .InBond = TInBondCodes.ibcNotInBond
                .LicenseDate = String.Empty
                .LicenseExceptionCode = TExceptionCodes.ecNLR
                .LicenseNumber = String.Empty
                .PointOfOrigin = "US"
                .ShippersTaxID = String.Empty
                .TransPortType = String.Empty
            End With
        End With

        Dim pbService As New PitneyBowesServices
        PitneyBowesServiceDictionary = New Dictionary(Of String, PitneyBowesServices)

        pbService.ServiceCode = nsoftware.InShip.ServiceTypes.stUSPSPriorityExpress '"85"
        pbService.ServiceDescription = "Priority Mail Express"
        PitneyBowesServiceDictionary.Add("EM", pbService)

        pbService = New PitneyBowesServices
        pbService.ServiceCode = nsoftware.InShip.ServiceTypes.stUSPSPriority ' "72"
        pbService.ServiceDescription = "Priority Mail"
        PitneyBowesServiceDictionary.Add("PM", pbService)

        pbService = New PitneyBowesServices
        pbService.ServiceCode = nsoftware.InShip.ServiceTypes.stUSPSFirstClass ' "71"
        pbService.ServiceDescription = "First-Class Mail"
        PitneyBowesServiceDictionary.Add("FCM", pbService)

        pbService = New PitneyBowesServices
        pbService.ServiceCode = nsoftware.InShip.ServiceTypes.stUSPSParcelSelect ' "79"
        pbService.ServiceDescription = "Parcel Select"
        PitneyBowesServiceDictionary.Add("PRCLSEL", pbService)

        pbService = New PitneyBowesServices
        pbService.ServiceCode = nsoftware.InShip.ServiceTypes.stUSPSStandardMail  ' "81"
        pbService.ServiceDescription = "Standard Post"
        PitneyBowesServiceDictionary.Add("STDPOST", pbService)

        pbService = New PitneyBowesServices
        pbService.ServiceCode = nsoftware.InShip.ServiceTypes.stUSPSMedia  ' "75"
        pbService.ServiceDescription = "Media Mail"
        PitneyBowesServiceDictionary.Add("MEDIA", pbService)

        pbService = New PitneyBowesServices
        pbService.ServiceCode = nsoftware.InShip.ServiceTypes.stUSPSLibrary  ' "76"
        pbService.ServiceDescription = "Library Mail"
        PitneyBowesServiceDictionary.Add("LIB", pbService)

    End Sub

#End Region

#Region "Class Properties"

    Public ReadOnly Property InternationalFormsFile() As String
        Get
            Return clsInternationalFormsFile
        End Get
    End Property

    Public Property oAuthKey As String
        Get
            Return cOAuthKey
        End Get
        Set(ByVal value As String)
            cOAuthKey = value
        End Set
    End Property

    Public Property OAuthSecret As String
        Get
            Return cOAuthSecret
        End Get
        Set(ByVal value As String)
            cOAuthSecret = value
        End Set
    End Property

    Public Property OAuthUrl As String
        Get
            Return cOAuthUrl
        End Get
        Set(ByVal value As String)
            cOAuthUrl = value
        End Set
    End Property


    Public ReadOnly Property UPSFreightClasses As Dictionary(Of Decimal, String)
        Get
            Dim FreightClass As New Dictionary(Of Decimal, String)
            FreightClass.Add(1, "500")
            FreightClass.Add(2, "400")
            FreightClass.Add(3, "300")
            FreightClass.Add(4, "250")
            FreightClass.Add(5, "200")
            FreightClass.Add(6, "175")
            FreightClass.Add(7, "150")
            FreightClass.Add(8, "125")
            FreightClass.Add(9, "110")
            FreightClass.Add(10.5, "100")
            FreightClass.Add(12, "92.5")
            FreightClass.Add(13.5, "85")
            FreightClass.Add(15, "77.5")
            FreightClass.Add(22.5, "70")
            FreightClass.Add(30, "65")
            FreightClass.Add(35, "60")
            FreightClass.Add(50, "55")
            FreightClass.Add(5000, "50")

            Return FreightClass
        End Get
    End Property

    Public ReadOnly Property InternationalForms As String
        Get
            Return cInternationalForms
        End Get
    End Property

    Public Property ShipmentDescription As String
        Get
            Return cShipmentDescription
        End Get
        Set(ByVal value As String)
            cShipmentDescription = value
        End Set
    End Property


    ''' <summary>
    ''' Set / Get rates total Value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RatesTotalValue() As Decimal
        Get
            Return cRatesTotalValue
        End Get
        Set(ByVal value As Decimal)
            cRatesTotalValue = value
        End Set
    End Property

    Public Property CustomerType() As UpsratesCustomerTypes
        Get
            Return cCustomerType
        End Get
        Set(ByVal value As UpsratesCustomerTypes)
            cCustomerType = value
        End Set
    End Property

    ''' <summary>
    '''  Custom Content to place on a Fedex Label
    '''  When using the CustomContent, the LabelStockType must be either 4 (Stock 4x8) or 5 (Stock 4x9 Leading Doc Tab). 
    '''  Also LabelFormatType must be 0 (Common2D) and LabelImageType must 2 (fitEltron), 3 (fitZebra) or 4 (fitUniMark). 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexCustomContent As String
        Get
            Return cFedexCustomContent
        End Get
        Set(value As String)
            cFedexCustomContent = value
        End Set
    End Property


    ''' <summary>
    ''' Gets set if a signature is required.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SignatureRequired As Boolean
        Get
            Return cSignatureRequired
        End Get
        Set(value As Boolean)
            cSignatureRequired = value
        End Set
    End Property

    ''' <summary>
    ''' Gets the raw request sent to shipper
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RawRequest As String
        Get
            Return cRawRequest
        End Get
    End Property

    ''' <summary>
    ''' Gets the Raw Response returned from the shipper
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RawResponse As String
        Get
            Return cRawResponse
        End Get
    End Property

    ''' <summary>
    ''' Get Shipment Master Tracking Number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property MasterTrackingNumber
        Get
            Return cMasterTrackingNumber
        End Get
    End Property

    ''' <summary>
    ''' Get / Set Service Provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Service As ServiceProviders
        Get
            Return cServiceProvider
        End Get
        Set(value As ServiceProviders)
            cServiceProvider = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set the Url for Service where requests are to be sent
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Server As String
        Get
            Server = cServer
        End Get
        Set(value As String)
            cServer = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set User Id for logging into the server. Not Required for Federal Express
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UserId As String
        Get
            Return cUserId
        End Get
        Set(value As String)
            cUserId = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Password for logging into the server
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Password As String
        Get
            Return (cPassword)
        End Get
        Set(value As String)
            cPassword = value
        End Set
    End Property

    Public Property IntegrationID As String
        Get
            IntegrationID = cIntegrationID
        End Get
        Set(value As String)
            cIntegrationID = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set the domestic service used in the ship request
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RequestedServiceType As ServiceTypes
        Get
            Return cRequestedServiceType
        End Get
        Set(value As ServiceTypes)
            cRequestedServiceType = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set the Shippers Account Number
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AccountNumber As String
        Get
            Return cAccountNumber
        End Get
        Set(value As String)
            cAccountNumber = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Identifting part of the authenication key useed for the sender's identity
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexDeveloperKey As String
        Get
            Return cFedexDeveloperKey
        End Get
        Set(value As String)
            cFedexDeveloperKey = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Meter Number to use for submitting request to the Fedex Server.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexMeterNumber As String
        Get
            Return cFedexMeterNumber
        End Get
        Set(value As String)
            cFedexMeterNumber = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set and identifer required to connect to UPS
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UPSAccessKey As String
        Get
            Return cUPSAccessKey
        End Get
        Set(value As String)
            cUPSAccessKey = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Mandatory Custoder Id for Endicia
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property USPSEndiciaCustomerId As String
        Get
            Return cUSPSEndiciaCustomerId
        End Get
        Set(value As String)
            cUSPSEndiciaCustomerId = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Mandatory Transaction ID for Endicia
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property USPSEndiciaTransactionId As String
        Get
            Return cUSPSEndiciaTransactionId
        End Get
        Set(value As String)
            cUSPSEndiciaTransactionId = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Set Sender Contact Information
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property Sender As Contact
        Get
            Return cSenderContact
        End Get
        Set(value As Contact)
            cSenderContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Sender Contact Information
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property Account As Contact
        Get
            Return cAccountContact
        End Get
        Set(value As Contact)
            cAccountContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Recipient Information
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property Recipient As Contact
        Get
            Return cRecipientContact
        End Get
        Set(value As Contact)
            cRecipientContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set ReturnAddress Information
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public Property ReturnAddress As Contact
        Get
            Return cReturnAddress
        End Get
        Set(value As Contact)
            cReturnAddress = value
        End Set
    End Property

    ''' <summary>
    ''' Set Payor Information
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PayorContact As Contact
        Get
            Return cPayorContact
        End Get
        Set(value As Contact)
            cPayorContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Duties Payor Information
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DutiesPayorContact As Contact
        Get
            Return cDutiesPayorContact
        End Get
        Set(value As Contact)
            cDutiesPayorContact = value
        End Set
    End Property

    ''' <summary>
    ''' Set Hold At Location Information
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property HoldAtLocation As Contact
        Get
            Return cHoldAtLocation
        End Get
        Set(value As Contact)
            cHoldAtLocation = value
        End Set
    End Property

    ''' <summary>
    ''' Set / Get total customs Value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TotalCustomsValue As Decimal
        Get
            Return cTotalCustomsValue
        End Get
        Set(value As Decimal)
            cTotalCustomsValue = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set The Indicia type used for a FedEx SmartPost shipment.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FedexSmartPost As SmartPost
        Get
            Return cFedexSmartPost
        End Get
        Set(value As SmartPost)
            cFedexSmartPost = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set label stock type
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LabelStockType As String
        Get
            Return cLabelStockType
        End Get
        Set(value As String)
            cLabelStockType = value
        End Set
    End Property


#End Region

#Region "Public Class Procedures"

    ''' <summary>
    ''' Request Shipping Label
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RequestLabel() As Boolean

        Try
            cMasterTrackingNumber = String.Empty
            cRawRequest = String.Empty
            cRawResponse = String.Empty
            LastError = String.Empty
            RequestedServicesRates.Clear()
            clsInternationalFormsFile = String.Empty

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            ValidateInputFieldLengths(Recipient.Address1,
                Recipient.Address2,
                Recipient.Address3,
                Recipient.FirstName,
                Recipient.City,
                Recipient.State,
                Recipient.ZipCode,
                Recipient.Company)

            ValidateInputFieldLengths(Sender.Address1,
                Sender.Address2,
                Sender.Address3,
                Sender.FirstName,
                Sender.City,
                Sender.State,
                Sender.ZipCode,
                Sender.Company)

            ValidateInputFieldLengths(ReturnAddress.Address1,
                ReturnAddress.Address2,
                ReturnAddress.Address3,
                ReturnAddress.FirstName,
                ReturnAddress.City,
                ReturnAddress.State,
                ReturnAddress.ZipCode,
                ReturnAddress.Company)

            ValidateInputFieldLengths(Account.Address1,
                Account.Address2,
                Account.Address3,
                Account.FirstName,
                Account.City,
                Account.State,
                Account.ZipCode,
                Account.Company)

            ValidateCommidtyManufacturer()

            Select Case cServiceProvider
                Case ServiceProviders.FederalExpressInternational
                    Return RequestFedexInternaltionalLabel()
                Case ServiceProviders.FederalExpress
                    Return RequestFedexLabel()
                Case ServiceProviders.UPS
                    Return RequestUPSLabel()
                Case ServiceProviders.UPSInternational
                    Return RequestUPSInternaltionalLabel()
                Case ServiceProviders.UPSFreight
                    Return RequestUpsFreightLabel()
                Case ServiceProviders.USPS
                    Return RequestUSPSLabel()
                Case Else
                    Return RequestLabelOther()
            End Select

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

    Private Function RequestUPSInternaltionalLabel() As Boolean

        Try

            LastError = String.Empty
            objUpsShipIntl = New nsoftware.InShip.Upsshipintl
            objUpsShipIntl.RuntimeLicense = inShipLicense
            objUpsShipIntl.Reset()
            objUpsShipIntl.RuntimeLicense = inShipLicense

            If cServiceProvider <> ServiceProviders.UPSInternational Then
                LastError = "Service type not a UPS International."
                Return False
            End If

            Dim useSoap As Boolean = cServer.ToUpper.Contains("WEBSERVICES")
            If useSoap Then
                objUpsShipIntl.Config("UseSOAP=true")
                If Not cServer.EndsWith("/") Then
                    objUpsShipIntl.UPSAccount.Server = cServer & "/Ship"
                Else
                    objUpsShipIntl.UPSAccount.Server = cServer & "Ship"
                End If
            Else
                objUpsShipIntl.Config("UseSOAP=false")
                If Not cServer.EndsWith("/") Then
                    objUpsShipIntl.UPSAccount.Server = cServer & "/ShipConfirm"
                Else
                    objUpsShipIntl.UPSAccount.Server = cServer & "ShipConfirm"
                End If
            End If

            objUpsShipIntl.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            objUpsShipIntl.UPSAccount.AccessKey = UPSAccessKey
            objUpsShipIntl.UPSAccount.Password = cPassword
            objUpsShipIntl.UPSAccount.AccountNumber = cAccountNumber
            objUpsShipIntl.UPSAccount.UserId = cUserId
            objUpsShipIntl.ServiceType = cRequestedServiceType
            objUpsShipIntl.ShipmentDescription = ShipmentDescription

            ' Get Sender Information
            With cSenderContact
                objUpsShipIntl.SenderContact.FirstName = .FirstName
                objUpsShipIntl.SenderContact.LastName = .LastName
                objUpsShipIntl.SenderContact.MiddleInitial = .MiddleInitial

                If .FirstName = String.Empty AndAlso .LastName = String.Empty Then
                    objUpsShipIntl.SenderContact.FirstName = "Warehouse Supervisor"
                End If

                objUpsShipIntl.SenderContact.Phone = .Phone
                objUpsShipIntl.SenderContact.Fax = .Fax
                objUpsShipIntl.SenderContact.Email = .eMail

                objUpsShipIntl.SenderContact.Company = .Company
                objUpsShipIntl.SenderAddress.Address1 = .Address1
                objUpsShipIntl.SenderAddress.Address2 = .Address2
                objUpsShipIntl.Config("SenderAddress3=" & .Address3)

                objUpsShipIntl.SenderAddress.City = .City
                objUpsShipIntl.SenderAddress.State = .State
                objUpsShipIntl.SenderAddress.ZipCode = .ZipCode
                objUpsShipIntl.SenderAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objUpsShipIntl.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShipIntl.SenderAddress.AddressFlags = &H1 'PO Box
                End If
            End With

            With cRecipientContact
                objUpsShipIntl.RecipientContact.FirstName = .FirstName
                objUpsShipIntl.RecipientContact.LastName = .LastName
                objUpsShipIntl.RecipientContact.MiddleInitial = .MiddleInitial
                objUpsShipIntl.RecipientContact.Phone = .Phone
                objUpsShipIntl.RecipientContact.Fax = .Fax
                objUpsShipIntl.RecipientContact.Email = .eMail

                objUpsShipIntl.RecipientContact.Company = .Company
                objUpsShipIntl.RecipientAddress.Address1 = .Address1
                objUpsShipIntl.RecipientAddress.Address2 = .Address2
                objUpsShipIntl.Config("RecipientAddress3=" & .Address3)

                objUpsShipIntl.RecipientAddress.City = .City
                objUpsShipIntl.RecipientAddress.State = .State
                objUpsShipIntl.RecipientAddress.ZipCode = .ZipCode
                objUpsShipIntl.RecipientAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objUpsShipIntl.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShipIntl.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            ' No ground service to Puerto Rico
            If objUpsShipIntl.RecipientAddress.State = "PR" AndAlso objUpsShipIntl.RecipientAddress.CountryCode = "US" Then
                objUpsShipIntl.RecipientAddress.CountryCode = "PR"
                objUpsShipIntl.RecipientAddress.State = ""
                If objUpsShipIntl.ServiceType = ServiceTypes.stUPSStandard Then
                    objUpsShipIntl.ServiceType = ServiceTypes.stUPSGround
                End If
            End If

            Select Case EzshipLabelImage
                Case EzshipLabelImageTypes.itEPL
                    objUpsShipIntl.LabelImageType = UpsshipLabelImageTypes.uitEPL
                Case EzshipLabelImageTypes.itGIF
                    objUpsShipIntl.LabelImageType = UpsshipLabelImageTypes.uitGIF
                Case EzshipLabelImageTypes.itSPL
                    objUpsShipIntl.LabelImageType = UpsshipLabelImageTypes.uitSPL
                Case EzshipLabelImageTypes.itZPL
                    objUpsShipIntl.LabelImageType = UpsshipLabelImageTypes.uitZPL
                Case Else
                    objUpsShipIntl.LabelImageType = UpsshipLabelImageTypes.uitEPL
            End Select

            Dim extension As String = objUpsShipIntl.LabelImageType.ToString
            If extension.StartsWith("uit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            ' Set Shipping Label File
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If Not ShippingLabelDirectory.EndsWith("\") Then
                    ShippingLabelDirectory &= "\"
                End If

                For Each shippingPackageDetail In PackageDetailList
                    Dim id As String = idCtr.ToString
                    idCtr += 1
                    shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                    shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                    shippingPackageDetail.ReturnReceiptFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_RTN" & extension
                Next
            End If

            Dim totalWeight As Double = 0
            Dim totalInsured As Double = 0

            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objUpsShipIntl.Packages.Add(shippingPackageDetail)
            Next

            objUpsShipIntl.ShipDate = ShipDate.ToString("yyyyMMdd")
            objUpsShipIntl.TotalCustomsValue = Format(TotalCustomsValue, "###0.00")

            ' Payor Types
            objUpsShipIntl.Payor.PayorType = Payor
            objUpsShipIntl.Payor.AccountNumber = cPayorContact.AccountNumber
            objUpsShipIntl.Payor.CountryCode = cPayorContact.CountryCode
            objUpsShipIntl.Payor.ZipCode = cPayorContact.ZipCode

            objUpsShipIntl.Payor.Name = cPayorContact.Company
            objUpsShipIntl.Payor.Address1 = cPayorContact.Address1
            objUpsShipIntl.Payor.Address2 = cPayorContact.Address2
            objUpsShipIntl.Payor.City = cPayorContact.City
            objUpsShipIntl.Payor.State = cPayorContact.State

            If objUpsShipIntl.Payor.Address1.Length > 35 Then
                objUpsShipIntl.Payor.Address1 = objUpsShipIntl.Payor.Address1.Substring(0, 35)
                objUpsShipIntl.Payor.Address2 = String.Empty
            Else
                Dim length As Int16 = 35 - objUpsShipIntl.Payor.Address1.Length
                If objUpsShipIntl.Payor.Address2.Length > length Then
                    objUpsShipIntl.Payor.Address2 = objUpsShipIntl.Payor.Address2.Substring(0, length)
                End If
            End If

            If objUpsShipIntl.Payor.AccountNumber <> cDutiesPayorContact.AccountNumber Then
                objUpsShipIntl.DutiesPayor.PayorType = DutiesPayor
                objUpsShipIntl.DutiesPayor.AccountNumber = cDutiesPayorContact.AccountNumber
                objUpsShipIntl.DutiesPayor.CountryCode = cDutiesPayorContact.CountryCode
                objUpsShipIntl.DutiesPayor.ZipCode = cDutiesPayorContact.ZipCode

                objUpsShipIntl.DutiesPayor.Name = cDutiesPayorContact.Company
                objUpsShipIntl.DutiesPayor.Address1 = cDutiesPayorContact.Address1
                objUpsShipIntl.DutiesPayor.Address2 = cDutiesPayorContact.Address2
                objUpsShipIntl.DutiesPayor.City = cDutiesPayorContact.City
                objUpsShipIntl.DutiesPayor.State = cDutiesPayorContact.State

                If objUpsShipIntl.DutiesPayor.Address1.Length > 35 Then
                    objUpsShipIntl.DutiesPayor.Address1 = objUpsShipIntl.DutiesPayor.Address1.Substring(0, 35)
                    objUpsShipIntl.DutiesPayor.Address2 = String.Empty
                Else
                    Dim length As Int16 = 35 - objUpsShipIntl.DutiesPayor.Address1.Length
                    If objUpsShipIntl.DutiesPayor.Address2.Length > length Then
                        objUpsShipIntl.DutiesPayor.Address2 = objUpsShipIntl.DutiesPayor.Address2.Substring(0, length)
                    End If
                End If
            End If


            Dim specialService As Long = ShipmentSpecialServices

            objUpsShipIntl.ShipmentSpecialServices = specialService

            For Each CommDetail As CommodityDetail In CommodityDetailList
                CommDetail.Weight = Format(Val(CommDetail.Weight), "###0.0")
                CommDetail.Description = CommDetail.Description.Replace("&", " ").Replace("<", " ").Replace(">", " ").Replace(Chr(34), " ")
                If CommDetail.Description.Length > 35 Then
                    CommDetail.Description = CommDetail.Description.Substring(0, 35).Trim
                End If

                If CommDetail.Manufacturer.Length = 0 Then
                    CommDetail.Manufacturer = "US"
                End If

                TotalCustomsValue += (Val(CommDetail.UnitPrice & String.Empty) * Val(CommDetail.NumberOfPieces & String.Empty))

                objUpsShipIntl.Commodities.Add(CommDetail)
            Next

            objUpsShipIntl.TotalCustomsValue = Format(TotalCustomsValue, "###0.00")

        Catch ex As Exception
            LastError = ex.Message
            objUpsShipIntl.Dispose()
            objUpsShipIntl = Nothing
            Return False
        End Try

        ' Notifications
        Dim notificationsIndex As Int16 = 0
        If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
            For Each sn As Notifications In ShipmentNotifications
                sn.email = (sn.email & String.Empty).Trim
                If sn.email.Length = 0 Then
                    Continue For
                End If

                Dim notify As New nsoftware.InShip.NotifyDetail
                With notify
                    .Email = sn.email
                    .NotificationFlags = CInt(sn.NotificationFlags)
                    .Message = (sn.Message & String.Empty).ToString.Trim
                End With

                objUpsShipIntl.Notify.Add(notify)

                notificationsIndex += 1
                If notificationsIndex = 3 Then Exit For
            Next
        End If

        'objUpsShipIntl.Config("AdditionalDocumentIndicator=True")
        'objUpsShipIntl.Config("OverridePaperlessIndicator=True")

        Dim ciDetail As New CommercialInvoiceDetail()
        If RequestedUPSInternationalForms.CommercialInvoice = True Then
            With RequestedUPSInternationalForms.CommercialInvoiceInfo
                ciDetail.Purpose = .Purpose
                ciDetail.Terms = .Terms
                ciDetail.Comments = .Comments
                ciDetail.Date = CDate(.InvoiceDate).ToString("yyyyMMdd")
                ciDetail.Number = .CustomersInvoiceNumber
                ciDetail.Insurance = .ShipperInsurance
                ciDetail.FreightCharge = .FreightCharge
            End With
            objUpsShipIntl.CommercialInvoice = ciDetail
            objUpsShipIntl.FormTypes = objUpsShipIntl.FormTypes Or &H1
        Else
            objUpsShipIntl.CommercialInvoice = ciDetail
        End If

        Dim sedDetail As New SEDDetail()
        If RequestedUPSInternationalForms.ShippersExportDeclaration = True AndAlso 1 = 2 Then
            With RequestedUPSInternationalForms.ShippersExportDeclarationInfo
                sedDetail.InBondCode = .InBond
                sedDetail.ExceptionCode = .LicenseExceptionCode
                sedDetail.LicenseNumber = .LicenseNumber
                sedDetail.LicenseDate = .LicenseDate
                sedDetail.EntryNumber = .ImportEntryNumber
                sedDetail.PointOfOrigin = .PointOfOrigin
                sedDetail.ShipperTaxId = .ShippersTaxID
                sedDetail.TransportType = .TransPortType
                objUpsShipIntl.SED = sedDetail
                objUpsShipIntl.Config("ExportingCarrier=" & .ExportingCarrier)
                objUpsShipIntl.Config("ExportDate=" & .ExportingDate)
            End With
            objUpsShipIntl.FormTypes = objUpsShipIntl.FormTypes Or &H2
        Else
            objUpsShipIntl.SED = sedDetail
        End If

        If RequestedUPSInternationalForms.CertificateOfOrigin = True Then
            objUpsShipIntl.Config("ExportingCarrier=" & RequestedUPSInternationalForms.CertificateOfOriginExportingCarrier)
            objUpsShipIntl.Config("ExportDate=" & RequestedUPSInternationalForms.CertificateOfOriginExportDate)
            objUpsShipIntl.FormTypes = objUpsShipIntl.FormTypes Or &H4
        End If

        If RequestedUPSInternationalForms.NAFTACertificateOfOrigin = True Then
            objUpsShipIntl.NAFTABlanketPeriod = RequestedUPSInternationalForms.NAFTABlanketPeriodStartDate & "-" & RequestedUPSInternationalForms.NAFTABlanketPeriodEndDate
            objUpsShipIntl.FormTypes = objUpsShipIntl.FormTypes Or &H8
        End If

        objUpsShipIntl.InternationalFormsFile = ShippingLabelDirectory & ShippingLabelPrefix & UPSnternationalFormsExtension
        clsInternationalFormsFile = objUpsShipIntl.InternationalFormsFile

        Try
            objUpsShipIntl.RuntimeLicense = inShipLicense
            objUpsShipIntl.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            ' For multi UPS package shipments the total cost exists in all packages
            ' so spread the costs

            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objUpsShipIntl.Packages.Count - 1
                PackageDetailList.Add(objUpsShipIntl.Packages(ictr))
                GetPackageCosts(objUpsShipIntl.Packages(ictr), objUpsShipIntl)
                Dim key As Integer = Val(objUpsShipIntl.Packages(ictr).Id)
                ShipmentBaseCharge(key) = Math.Round(ShipmentBaseCharge(key) / objUpsShipIntl.Packages.Count, 2)
                ShipmentDiscountCharge(key) = Math.Round(ShipmentDiscountCharge(key) / objUpsShipIntl.Packages.Count, 2)
                ShipmentSurCharge(key) = Math.Round(ShipmentSurCharge(key) / objUpsShipIntl.Packages.Count, 2)
                ShipmentNetCharge(key) = Math.Round(ShipmentNetCharge(key) / objUpsShipIntl.Packages.Count, 2)
                ShipmentListCharge(key) = Math.Round(ShipmentListCharge(key) / objUpsShipIntl.Packages.Count, 2)
            Next

            If objUpsShipIntl.Packages.Count = 1 Then
                cMasterTrackingNumber = objUpsShipIntl.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objUpsShipIntl.MasterTrackingNumber
            End If

            cInternationalForms = objUpsShipIntl.InternationalForms

            Return True

        Catch ex As nsoftware.InShip.InShipUpsshipintlException
            LastError = ex.Message
            Return False

        Catch exc As Exception
            LastError = exc.Message
            Return False

        Finally
            cRawRequest = objUpsShipIntl.Config("RawRequest")
            cRawResponse = objUpsShipIntl.Config("RawResponse")
            objUpsShipIntl.Dispose()
            objUpsShipIntl = Nothing
        End Try

        Return True

    End Function

    Private Sub ValidateInputFieldLengths(ByRef AddressLine1 As String,
            ByRef AddressLine2 As String,
            ByRef AddressLine3 As String,
            ByRef AttentionName As String,
            ByRef City As String,
            ByRef StateProvinceCode As String,
            ByRef PostalCode As String,
            ByRef CompanyName As String)

        AddressLine1 = AddressLine1 & String.Empty
        If AddressLine1.Length > 35 Then
            AddressLine1 = AddressLine1.Substring(0, 35).Trim
        End If

        AddressLine2 = AddressLine2 & String.Empty
        If AddressLine2.Length > 35 Then
            AddressLine2 = AddressLine2.Substring(0, 35).Trim
        End If

        AddressLine3 = AddressLine3 & String.Empty
        If AddressLine3.Length > 35 Then
            AddressLine3 = AddressLine3.Substring(0, 35).Trim
        End If

        AttentionName = AttentionName & String.Empty
        If AttentionName.Length > 35 Then
            AttentionName = AttentionName.Substring(0, 35).Trim
        End If

        City = City & String.Empty
        If City.Length > 30 Then
            City = City.Substring(0, 30).Trim
        End If

        StateProvinceCode = StateProvinceCode & String.Empty
        If StateProvinceCode.Length > 5 Then
            StateProvinceCode = StateProvinceCode.Substring(0, 5).Trim
        End If

        PostalCode = PostalCode & String.Empty
        If PostalCode.Length > 10 Then
            PostalCode = PostalCode.Substring(0, 10).Trim
        End If

        CompanyName = CompanyName & String.Empty
        If CompanyName.Length > 35 Then
            CompanyName = CompanyName.Substring(0, 35).Trim
        End If

    End Sub

    Private Sub ValidateCommidtyManufacturer()

        If CommodityDetailList Is Nothing Then
            Exit Sub
        End If

        For Each CommDetail As CommodityDetail In CommodityDetailList

            Select Case (CommDetail.Manufacturer & String.Empty).Trim.ToUpper

                Case "BAN" ' BANGLADESH
                    CommDetail.Manufacturer = "BD"

                Case "CHE" 'SWITZERLAND
                    CommDetail.Manufacturer = "CH"

                Case "CHN", "CHI" 'CHINA
                    CommDetail.Manufacturer = "CN"

                Case "HK" ' HONG KONG
                    CommDetail.Manufacturer = "HK"

                Case "HND" 'HONDURAS
                    CommDetail.Manufacturer = "HN"

                Case "INA" ' INDONESIA
                    CommDetail.Manufacturer = "ID"

                Case "IND" 'INDIA
                    CommDetail.Manufacturer = "IN"

                Case "KHM" 'CAMBODIA
                    CommDetail.Manufacturer = "KH"

                Case "MAL" ' MALAYSIA
                    CommDetail.Manufacturer = "MY"

                Case "PHL" 'PHILIPPINES
                    CommDetail.Manufacturer = "PH"

                Case "REU" 'REUNIO
                    CommDetail.Manufacturer = "RE"

                Case "THA" 'THAILAND
                    CommDetail.Manufacturer = "TH"

                Case "TWN" 'TAIWAN, PROVINCE OF CHINA
                    CommDetail.Manufacturer = "TW"

                Case "USA", "", "ZZZ" 'UNITED STATES
                    CommDetail.Manufacturer = "US"

            End Select
        Next
    End Sub

    Public Function CancelShipment(ByVal TrackingNumber As String) As Boolean
        Return CancelShipment(TrackingNumber, False, 0)
    End Function

    Public Function CancelShipment(ByVal TrackingNumber As String, ByVal isMultiPackage As Boolean, ByVal FedexTrackingIDType As Int32) As Boolean

        Select Case cServiceProvider
            Case ServiceProviders.CanadaPost
                Return False
            Case ServiceProviders.FederalExpress, ServiceProviders.FederalExpressInternational
                Return CancelFedexShipment(TrackingNumber, isMultiPackage, FedexTrackingIDType)
            Case ServiceProviders.UPS
                Return CancelUpsShipment(TrackingNumber, isMultiPackage)
            Case ServiceProviders.USPS
                PitneyBowesUniqueTransactionID = FedexTrackingIDType
                Return VoidShipment(TrackingNumber)
            Case Else
                LastError = "Unknown Carrier"
                Return False
        End Select

    End Function

    Public Function RequestShipmentRates() As Decimal

        Try
            RequestedServicesRates.Clear()
            objEzRates = New nsoftware.InShip.Ezrates
            objEzRates.RuntimeLicense = inShipLicense
            objEzRates.Reset()
            objEzRates.RuntimeLicense = inShipLicense
            objEzRates.Provider = Service

            Dim CODTotalAmount As Decimal = 0
            Dim InsuredValue As Decimal = 0
            Dim Totalweight As Decimal = 0

            For Each PackageDetail As nsoftware.InShip.PackageDetail In PackageDetailList
                CODTotalAmount += Val(PackageDetail.CODAmount & String.Empty)
                InsuredValue += Val(PackageDetail.InsuredValue & String.Empty)
                PackageDetail.Weight = Format(Val(PackageDetail.Weight) / 16, "###0.0")
                Totalweight += PackageDetail.Weight
                objEzRates.Packages.Add(PackageDetail)
            Next

            With objEzRates
                .TotalWeight = Format(Totalweight, "###0.0")

                Dim EzAccount As New nsoftware.InShip.EzAccount
                With EzAccount
                    .AccountNumber = cAccountNumber
                    .DeveloperKey = cFedexDeveloperKey
                    .MeterNumber = cFedexMeterNumber
                    .Password = cPassword
                    .Server = cServer
                    .UserId = cUserId
                    .AccessKey = cUPSAccessKey
                End With

                .Account = EzAccount

                With cRecipientContact
                    objEzRates.RecipientAddress.Address1 = .Address1
                    objEzRates.RecipientAddress.Address2 = .Address2
                    'objEzRates.Config("RecipientAddress3=" & .Address3)

                    objEzRates.RecipientAddress.City = .City
                    objEzRates.RecipientAddress.ZipCode = .ZipCode
                    objEzRates.RecipientAddress.State = .State
                    objEzRates.RecipientAddress.CountryCode = .CountryCode

                    If .IsResidental Then
                        objEzRates.RecipientAddress.AddressFlags = &H2 'Residential
                    ElseIf .IsPOBox Then
                        objEzRates.RecipientAddress.AddressFlags = &H1 'PO Box
                    End If
                End With

                With cSenderContact
                    objEzRates.SenderAddress.Address1 = .Address1
                    objEzRates.SenderAddress.Address2 = .Address2
                    'objEzRates.Config("SenderAddress3=" & .Address3)

                    objEzRates.SenderAddress.City = .City
                    objEzRates.SenderAddress.ZipCode = .ZipCode
                    objEzRates.SenderAddress.State = .State
                    objEzRates.SenderAddress.CountryCode = .CountryCode

                    If .IsResidental Then
                        objEzRates.SenderAddress.AddressFlags = &H2 'Residential
                    ElseIf .IsPOBox Then
                        objEzRates.RecipientAddress.AddressFlags = &H1 'PO Box
                    End If
                End With

                Dim specialService As Long = ShipmentSpecialServices
                .GetRates()

                For Each reqSer As ServiceDetail In .Services
                    RequestedServicesRates.Add(reqSer)
                Next

            End With

        Catch ex As Exception
            LastError = ex.Message
            Return -1
        Finally
            objEzRates.Dispose()
            objEzRates = Nothing
        End Try
    End Function

    Public Function GetRates() As Decimal
        Try
            cMasterTrackingNumber = String.Empty
            cRawRequest = String.Empty
            cRawResponse = String.Empty
            LastError = String.Empty
            RequestedServicesRates.Clear()

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            Select Case cServiceProvider
                Case ServiceProviders.FederalExpressInternational
                    Return GetFedexRates()
                Case ServiceProviders.FederalExpress
                    Return GetFedexRates()
                Case ServiceProviders.UPS
                    Return GetUPSRates()
                Case ServiceProviders.UPSInternational
                    Return GetUPSRates()
                Case Else
                    Return 0
            End Select

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "Pitney Bowes"

    ' Pitney Bowes URLs

    'oAuth
    'Sandbox:  https://api-sandbox.pitneybowes.com/oauth/token
    'Production:  https://api.pitneybowes.com/oauth/token 

    ' Rates
    'Sandbox: https://api-sandbox.pitneybowes.com/shippingservices/v1/rates 
    'Production: https://api.pitneybowes.com/shippingservices/v1/rates

    ' Shipping Labels
    'Sandbox: https://api-sandbox.pitneybowes.com/shippingservices/v1/shipments 
    'Production: https://api.pitneybowes.com/shippingservices/v1/shipments

    Public PitneyBowesUniqueTransactionID As String = String.Empty
    Public PitneyBowesInductionPostalCode As String = String.Empty

    Public Class PitneyBowesPackageInformation
        Public TrackingNumber As String = String.Empty
        Public ShipmentID As String = String.Empty
        Public ShippingLabel As String = String.Empty
        Public CODLabel As String = String.Empty
        Public ReturnReceipt As String = String.Empty
    End Class

    Private Class PitneyBowesServices
        Public ServiceCode As String = String.Empty
        Public ServiceDescription As String = String.Empty
    End Class

    Private PitneyBowesServiceDictionary As Dictionary(Of String, PitneyBowesServices)

    Private Class OAuthtoken
        Public Authorization As String = String.Empty
        Public Content_Type As String = String.Empty
        Public client_id As String = String.Empty
        Public scope As String = String.Empty
    End Class

    Private Class OAuthtokenPitneyBowesResponse
        Public access_token As String = String.Empty
        Public tokenType As String = String.Empty
        Public issuedAt As String = String.Empty
        Public expiresIn As String = String.Empty
        Public clientID As String = String.Empty
        Public org As String = String.Empty
    End Class

    Private Class PitneyBowesCancelLabelRequest
        Public carrier As String = String.Empty
        Public cancelInitiator As String = String.Empty
    End Class

    Private Class PitneyBowesCancelLabelResponse
        Public carrier As String = String.Empty
        Public cancelInitiator As String = String.Empty
        Public totalCarrierCharge As Decimal = 0
        Public parcelTrackingNumber As String = String.Empty
        Public status As String = String.Empty
    End Class

    Private Class PitneyBowesInputParameters
        Public name As String = String.Empty
        Public value As String = String.Empty
    End Class

    Private Class PitneyBowesSpecialServices
        Public specialServiceId As String = String.Empty
        Public inputParameters(1) As PitneyBowesInputParameters

        <JsonPropertyAttribute("fee", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public fee As Decimal = 0
    End Class

    Private Class PitneyBowesSurcharges
        Public name As String = String.Empty
        Public fee As Decimal = 0
    End Class

    Private Class PitneyBowesDeliveryCommitment
        Public minEstimatedNumberOfDays As String = String.Empty
        Public maxEstimatedNumberOfDays As String = String.Empty
        Public estimatedDeliveryDateTime As String = String.Empty
        Public guarantee As String = String.Empty
        Public additionalDetails As String = String.Empty
    End Class

    Private Class PitneyBowesRatesRequestObject
        Public fromAddress As New PitneyBowesAddress
        Public toAddress As New PitneyBowesAddress
        Public parcel As New PitneyBowesParcel
        Public rates() As PitneyBowesRatesRequest
    End Class

    Private Class PitneyBowesRatesResponseObject
        Public fromAddress As New PitneyBowesAddress
        Public toAddress As New PitneyBowesAddress
        Public parcel As New PitneyBowesParcel
        Public rates(1) As PitneyBowesRatesResponse
    End Class

    Public Class PitneyBowesAddress
        Public addressLines(2) As String ' Required. Street address or P.O. Box. Include apartment number if applicable. You can specify up to 3 address lines.

        <JsonPropertyAttribute("cityTown", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public cityTown As String = String.Empty ' Conditional: The city or town.

        <JsonPropertyAttribute("stateProvince", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public stateProvince As String = String.Empty ' Conditional: The state or province. For US address, use the 2-letter state code.

        Public postalCode As String = String.Empty ' Required. Postal/ZIP code. For US addresses, either the 5-digit or 9-digit ZIP code.
        Public countryCode As String = String.Empty ' Required. Two-character country code from the ISO country list.

        <JsonPropertyAttribute("company", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public company As String = String.Empty ' Conditional: The name of the company.

        <JsonPropertyAttribute("name", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public name As String = String.Empty ' Conditional: The first and last name.

        <JsonPropertyAttribute("phone", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public phone As String = String.Empty ' Conditional: The phone number.

        <JsonPropertyAttribute("email", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public email As String = String.Empty ' Conditional: The email address.

        Public residential As Boolean = True ' Indicates whether this is a residential address. It is recommended that this parameter be passed in as the address verification process is more accurate with it.

        <JsonPropertyAttribute("deliveryPoint", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public deliveryPoint As String = String.Empty ' The 2-digit delivery point, when available.

        <JsonPropertyAttribute("carrierRoute", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public carrierRoute As String = String.Empty ' The last four characters of the USPS carrier route code. The carrier route is the area served by a particular USPS mail carrier. The full carrier route code is a nine-character string comprising the five-digit postal code appended by these four characters.

        <JsonPropertyAttribute("taxId", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public taxId As String = String.Empty ' Pickup Request Only. Tax identification number. This is optional for pickup requests.

        <JsonPropertyAttribute("status", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public status As String = Nothing ' Response Only. Indicates whether the address is valid and whether the validation check made changes to the address.

        ' Valid Status Values
        '        Note: The response does not return this field if minimal address validation is enabled.
        '       Valid Values are:
        '           VALIDATED_CHANGED: The address is valid. The validation check made changes to the address.
        '           VALIDATED_AND_NOT_CHANGED: The address is valid. No changes were made.
        '           NOT_CHANGED: The address could not be validated. No changes were made.
    End Class

    Private Class PitneyBowesContact
        Public email As String   'Required. Email address.
        Public firstName As String  'Required. First name.
        Public lastName As String   'Required. Last name.
        Public company As String  'Required. Company name.
        Public phone As String   'Required. A valid phone number for the merchant.
        'For USPS, this must be a valid 10-digit number. The string should contain 10 numeric characters and no additional characters. For example: "8442566444"
        Public addressLines(3) As String   'Required. Street address, including the apartment number if applicable.
        'Note: For USPS, the address cannot be a P.O. Box.
        Public cityTown As String   ' Conditional: The city or town. Required if postalCode is absent.
        Public stateProvince As String ' Conditional: The state or province. For US addresses, use the 2-letter state code.
        ' Required if postalCode is absent; otherwise optional. Note that in some cases where cityTown is a unique name within the country, this can be left out, even if postalCode is absent. But the best practice is to include this field whenever postalCode absent.
        Public postalCode As String  ' Conditional: The postal/ZIP code. For US addresses, use either the 5-digit or 9-digit ZIP code.
        'Required if you are creating a shipment. Optional if you are rating a package.
        Public countryCode As String   ' Required. Two-character country code from the ISO country list.
    End Class

    Private Class PitneyBowesWeight
        Public weight As Decimal = 0
        Public unitOfMeasurement As String = "OZ"
    End Class

    Private Class PitneyBowesDimension
        Public length As Decimal = 0
        Public width As Decimal = 0
        Public height As Decimal = 0
        Public unitOfMeasurement As String = "IN"

        <JsonPropertyAttribute("irregularParcelGirth", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public irregularParcelGirth As Decimal = Nothing
    End Class

    Private Class PitneyBowesParcel
        Public weight As New PitneyBowesWeight
        Public dimension As New PitneyBowesDimension
    End Class

    Private Class PitneyBowespages
        Public contents As String = String.Empty
    End Class

    Private Class PitneyBowesDocument
        '"type": "SHIPPING_LABEL",
        '"contentType": "BASE64",
        '"size": "DOC_4X6",
        '"fileFormat": "PNG",
        '"resolution": "DPI_203",
        '"printDialogOption": "EMBED_PRINT_DIALOG"

        Public type As String = String.Empty
        Public contentType As String = String.Empty
        Public size As String = String.Empty
        Public fileFormat As String = String.Empty
        Public resolution As String = String.Empty

        <JsonPropertyAttribute("printDialogOption", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public printDialogOption As String = String.Empty

        <JsonPropertyAttribute("docTab", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public docTab(1) As PitneyBowesDoctab

        <JsonPropertyAttribute("contents", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public contents As String = String.Empty

        <JsonPropertyAttribute("pages", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public pages(1) As PitneyBowespages

    End Class

    Private Class PitneyBowesShipmentOption
        '"name": "SHIPPER_ID",
        '"value": "9024324564"
        Public name As String = String.Empty
        Public value As String = String.Empty
    End Class

    Private Class PitneyBowesDoctab
        Public name As String = String.Empty
        Public displayName As String = String.Empty
        Public value As String = String.Empty
    End Class

    Private Class PitneyBowesRatesRequest
        Public carrier As String = "USPS"

        <JsonPropertyAttribute("serviceId", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public serviceId As String = Nothing

        Public parcelType As String = "PKG"

        <JsonPropertyAttribute("specialServices", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public specialServices() As PitneyBowesSpecialServices = Nothing

        <JsonPropertyAttribute("inductionPostalCode", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public inductionPostalCode As String = String.Empty

        <JsonPropertyAttribute("dimensionalWeight", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public dimensionalWeight As PitneyBowesDimension = Nothing

        <JsonPropertyAttribute("baseCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public baseCharge As Decimal = Nothing ' 

        <JsonPropertyAttribute("totalCarrierCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public totalCarrierCharge As Decimal = Nothing ' 

        <JsonPropertyAttribute("surcharges", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public surcharges() As PitneyBowesSurcharges = Nothing

        <JsonPropertyAttribute("alternateBaseCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public alternateBaseCharge As Decimal = Nothing ' 

        <JsonPropertyAttribute("alternateTotalCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public alternateTotalCharge As Decimal = Nothing ' 

        <JsonPropertyAttribute("deliveryCommitment", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public deliveryCommitment As PitneyBowesDeliveryCommitment = Nothing

        <JsonPropertyAttribute("currencyCode", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public currencyCode As String = Nothing

        <JsonPropertyAttribute("destinationZone", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public destinationZone As Int16 = Nothing ' 

        <JsonPropertyAttribute("rateTypeId", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public rateTypeId As String = Nothing
    End Class

    Private Class PitneyBowesRatesResponse
        Public carrier As String = String.Empty

        <JsonPropertyAttribute("serviceId", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public serviceId As String = String.Empty

        Public parcelType As String = String.Empty

        <JsonPropertyAttribute("specialServices", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public specialServices(1) As PitneyBowesSpecialServices

        <JsonPropertyAttribute("inductionPostalCode", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public inductionPostalCode As String = String.Empty

        <JsonPropertyAttribute("dimensionalWeight", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public dimensionalWeight As New PitneyBowesDimension

        <JsonPropertyAttribute("baseCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public baseCharge As Decimal = 0

        <JsonPropertyAttribute("totalCarrierCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public totalCarrierCharge As Decimal = 0

        <JsonPropertyAttribute("surcharges", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public surcharges(1) As PitneyBowesSurcharges

        <JsonPropertyAttribute("alternateBaseCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public alternateBaseCharge As Decimal = 0

        <JsonPropertyAttribute("alternateTotalCharge", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public alternateTotalCharge As Decimal = 0

        <JsonPropertyAttribute("deliveryCommitment", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public deliveryCommitment As New PitneyBowesDeliveryCommitment

        <JsonPropertyAttribute("currencyCode", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public currencyCode As String = String.Empty

        <JsonPropertyAttribute("destinationZone", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public destinationZone As Int16 = 0

        <JsonPropertyAttribute("rateTypeId", DefaultValueHandling:=NullValueHandling.Ignore)>
        Public rateTypeId As String = Nothing

    End Class

    Private Class PitneyBowesLabelRequestObject
        Public fromAddress As New PitneyBowesAddress
        Public toAddress As New PitneyBowesAddress
        Public parcel As New PitneyBowesParcel
        Public rates() As PitneyBowesRatesRequest
        Public documents() As PitneyBowesDocument
        Public shipmentOptions() As PitneyBowesShipmentOption
    End Class

    Private Class PitneyBowesLabelResponsetObject
        Public fromAddress As New PitneyBowesAddress
        Public toAddress As New PitneyBowesAddress
        Public parcel As New PitneyBowesParcel
        Public rates(1) As PitneyBowesRatesRequest
        Public documents(1) As PitneyBowesDocument
        Public shipmentOptions(1) As PitneyBowesShipmentOption
        Public shipmentId As String = String.Empty
        Public parcelTrackingNumber As String = String.Empty
    End Class

    Private Class PitneyBowesLabelErrors
        Public errorCode As String = String.Empty
    End Class

    Private Class PitneyBowesLabelErrorDescription
        Public additionalInfo As String = String.Empty
        Public parameters(1) As String
    End Class

    Private Class PitneyBowesLabelErrorsObject
        Public errors(1) As PitneyBowesLabelErrors
        Public errorDescription As PitneyBowesLabelErrorDescription
    End Class

    Private Property OAuthAuthenticationKey(ByVal CARRIER_CODE As String) As String
        Get
            Dim row As DataRow = ASCDATA1.GetDataRow("Select * from SOTCARR3 WHERE CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = 'OAUTH'")
            If row IsNot Nothing Then
                Return row.Item("METER_NUMBER") & String.Empty
            Else
                Return String.Empty
            End If
        End Get

        Set(value As String)
            ASCDATA1.ExecuteSQL("Update SOTCARR3 set METER_NUMBER = '" & value & "' where CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = 'OAUTH'")
        End Set
    End Property

    Private Property OAuthAuthenticationKeyExpires(ByVal CARRIER_CODE As String) As String
        Get
            Dim row As DataRow = ASCDATA1.GetDataRow("Select * from SOTCARR3 WHERE CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = 'OAUTH'")
            If row IsNot Nothing Then
                Return row.Item("TOKEN_EXPIRES") & String.Empty
            Else
                Return String.Empty
            End If
        End Get

        Set(value As String)
            ASCDATA1.ExecuteSQL("Update SOTCARR3 set TOKEN_EXPIRES = sysdate + interval '" & value & "' second where CARRIER_CODE = '" & CARRIER_CODE & "' AND SHIPPER_DIVISION_CODE = 'OAUTH'")
        End Set
    End Property

    Private Function GetoAuthToken(ByVal url As String, ByVal CARRIER_CODE As String) As String

        Dim access_token As String = OAuthAuthenticationKey(CARRIER_CODE) & String.Empty
        Dim access_token_expires As String = OAuthAuthenticationKeyExpires(CARRIER_CODE) & String.Empty
        Dim postResult As String = String.Empty

        Dim needNewToken As Boolean = False

        If access_token.Length = 0 Then
            needNewToken = True
        ElseIf Not IsDate(access_token_expires) Then
            needNewToken = True
        Else
            Dim secs As Double = CDate(access_token_expires).Subtract(DateTime.Now).TotalSeconds
            If secs < 300 Then
                needNewToken = True
            End If
        End If

        If Not needNewToken Then
            Return access_token
        End If

        Dim encodedKey As String = String.Empty

        Dim enc As ASCIIEncoding
        enc = New System.Text.ASCIIEncoding()

        ' Convert to base 64
        Dim data As Byte()
        data = enc.GetBytes(cOAuthKey & ":" & cOAuthSecret)
        encodedKey = System.Convert.ToBase64String(data)

        Try
            'Authorization: Basic {base64Value}
            'Content-Type: application/x-www-form-urlencoded 
            'POST https://api.pitneybowes.com/oauth/token ' "https://api-sandbox.pitneybowes.com/oauth/token"
            ' grant_type = client_credentials

            url = cOAuthUrl
            Dim headers As New System.Net.WebHeaderCollection
            headers.Add("Authorization", "Basic " & encodedKey)

            Dim contentType As String = "application/x-www-form-urlencoded"
            Dim method As String = "POST"
            Dim postData As String = "grant_type=client_credentials"

            postResult = PostHttpRequest(url, Nothing, postData, contentType, method, headers)

            '{
            '"access_token": "<oauth_token>",
            '"tokenType": "BearerToken",
            '"issuedAt": "1498168771002",
            '"expiresIn": "35999",
            '"clientID": "a3cDEFghI1jK2LMnOP3qRstU4vWX5Yz",
            '"org": "pitneybowes"
            '}

            Dim tokenRepsonse As OAuthtokenPitneyBowesResponse = JsonConvert.DeserializeObject(Of OAuthtokenPitneyBowesResponse)(postResult)
            access_token = tokenRepsonse.access_token
            Dim token_type As String = tokenRepsonse.tokenType
            Dim expires_in As Int64 = tokenRepsonse.expiresIn

            OAuthAuthenticationKey(CARRIER_CODE) = access_token
            OAuthAuthenticationKeyExpires(CARRIER_CODE) = expires_in.ToString

        Catch ex As Exception
            LastError = ex.Message
            access_token = String.Empty
        End Try

        Return access_token

    End Function

    Private Function PostHttpRequest(ByVal Url As String,
                              ByVal jsonObject As Object,
                              ByVal postData As String,
                              ByVal contentType As String,
                              ByVal method As String,
                              Optional headers As System.Net.WebHeaderCollection = Nothing) As String

        Dim result As System.Net.WebResponse = Nothing
        LastError = String.Empty
        Dim Response As String = String.Empty

        Try
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            Dim httpRequest As HttpWebRequest
            httpRequest = HttpWebRequest.Create(Url)

            Dim jSonData As String = String.Empty
            Dim postdatabytes As Byte() = Nothing

            If jsonObject IsNot Nothing Then
                jSonData = JsonConvert.SerializeObject(jsonObject, Newtonsoft.Json.Formatting.Indented)
                postdatabytes = Encoding.ASCII.GetBytes(jSonData)
            ElseIf postData.Length > 0 Then
                postdatabytes = Encoding.ASCII.GetBytes(postData)
            End If

            If headers IsNot Nothing Then
                'httpRequest.Headers = headers
                For Each key As String In headers.AllKeys
                    If key & String.Empty <> String.Empty AndAlso headers(key) & String.Empty <> String.Empty Then
                        httpRequest.Headers.Add(key, headers(key))
                    End If
                Next
            End If

            httpRequest.Method = method
            httpRequest.ContentType = contentType
            If postdatabytes IsNot Nothing Then
                httpRequest.ContentLength = postdatabytes.Length
            Else
                httpRequest.ContentLength = 0
            End If

            If postdatabytes IsNot Nothing Then
                Using stream = httpRequest.GetRequestStream()
                    stream.Write(postdatabytes, 0, postdatabytes.Length)
                End Using
            End If

            Try
                result = httpRequest.GetResponse()
            Catch exw As System.Net.WebException
                Response = New StreamReader(exw.Response.GetResponseStream()).ReadToEnd
                LastError = exw.Message & " " & Response
            Catch ex As Exception
                LastError = ex.Message
            End Try

            If result IsNot Nothing Then
                Response = New StreamReader(result.GetResponseStream()).ReadToEnd
            End If

            Return Response

        Catch exw As System.Net.WebException
            LastError = exw.Message
            Return String.Empty

        Catch ex As Exception
            LastError = ex.Message
            Return String.Empty
        End Try

    End Function

    Public Function GetPitneyBowesAddresses(ByVal Recipient As Contact) As PitneyBowesAddress

        LastError = String.Empty

        Try

            Dim oAuthToken As String = GetoAuthToken("", "USPS")
            Dim numAddresses As Decimal = 0
            Dim responseAddresses(1) As PitneyBowesAddress

            If oAuthToken Is Nothing OrElse oAuthToken.Length = 0 Then
                Return Nothing
            End If

            Dim shipToAddress As New PitneyBowesAddress
            With shipToAddress
                numAddresses = 0

                If Recipient.Address1.Length > 0 Then
                    numAddresses += 1
                End If
                If Recipient.Address2.Length > 0 Then
                    numAddresses += 1
                End If
                If Recipient.Address3.Length > 0 Then
                    numAddresses += 1
                End If

                ReDim .addressLines(numAddresses - 1)
                numAddresses = 0
                If Recipient.Address1.Length > 0 Then
                    .addressLines(numAddresses) = Recipient.Address1
                    numAddresses += 1
                End If
                If Recipient.Address2.Length > 0 Then
                    .addressLines(numAddresses) = Recipient.Address2
                    numAddresses += 1
                End If
                If Recipient.Address3.Length > 0 Then
                    .addressLines(numAddresses) = Recipient.Address3
                    numAddresses += 1
                End If

                .carrierRoute = Nothing

                .cityTown = Recipient.City
                If .cityTown.Length = 0 Then
                    .cityTown = Nothing
                End If

                .company = Recipient.Company
                If .company.Length = 0 Then
                    .company = Nothing
                End If

                .countryCode = Recipient.CountryCode
                If .countryCode.Length = 0 Then
                    .countryCode = Nothing
                End If

                .deliveryPoint = Nothing

                .email = Recipient.eMail
                If .email.Length = 0 Then
                    .email = Nothing
                End If

                .name = (Recipient.FirstName & " " & Recipient.LastName).ToString.Trim
                If .name.Length = 0 Then
                    .name = Nothing
                End If

                .phone = Recipient.Phone
                If .phone.Length = 0 Then
                    .phone = Nothing
                End If

                .postalCode = Recipient.ZipCode
                If .postalCode.Length <> 5 AndAlso .postalCode.Length <> 9 Then
                    If .postalCode.Length >= 5 Then
                        .postalCode = .postalCode.Substring(0, 5)
                    End If
                End If
                If .postalCode.Length = 0 Then
                    .postalCode = Nothing
                End If

                .residential = Recipient.IsResidental

                .stateProvince = Recipient.State
                If .stateProvince.Length = 0 Then
                    .stateProvince = Nothing
                End If

                .taxId = Nothing
                .status = Nothing

            End With

            'curl -X POST .../v1/addresses/verify?minimalAddressValidation=false \
            '-H "Authorization: Bearer <oauth_token>" \
            '-H "Content-Type: application/json" \
            '-H "X-PB-UnifiedErrorStructure: true" \
            '-d '

            Dim url As String = cServer & "/v1/addresses/verify?minimalAddressValidation=false"
            Dim headers As New System.Net.WebHeaderCollection
            headers.Add("Authorization", "Bearer " & oAuthToken)
            headers.Add("X-PB-UnifiedErrorStructure", "true")

            Dim contentType As String = "application/json"
            Dim method As String = "POST"

            Dim postResult As String = PostHttpRequest(url, shipToAddress, String.Empty, contentType, method, headers)

            If postResult = Nothing Then
                Return Nothing
            End If

            If postResult.Length = 0 Then
                Return Nothing
            End If

            If postResult.StartsWith("{""errors") Then
                LastError = postResult
                Return Nothing
            End If

            Dim addressResponse As PitneyBowesAddress = JsonConvert.DeserializeObject(Of PitneyBowesAddress)(postResult)


            Return addressResponse

        Catch exn As Newtonsoft.Json.JsonException
            LastError = exn.Message
            Return Nothing

        Catch ex As Exception
            LastError = ex.Message
            Return Nothing
        End Try

    End Function

    Private Function GetPitneyBowesRates() As RateList()

        Dim requestedRateList As RateList()
        ReDim requestedRateList(1)
        LastError = String.Empty

        Try
            Dim oAuthToken As String = GetoAuthToken("", "USPS")
            Dim numAddresses As Decimal = 0

            If oAuthToken Is Nothing OrElse oAuthToken.Length = 0 Then
                Return Nothing
            End If

            Dim dicReturnRates As New Dictionary(Of String, RateList)

            Dim ratesObject As New PitneyBowesRatesRequestObject
            With ratesObject
                With .fromAddress
                    numAddresses = 0
                    If Sender.Address1.Length > 0 Then
                        numAddresses += 1
                    End If
                    If Sender.Address2.Length > 0 Then
                        numAddresses += 1
                    End If
                    If Sender.Address3.Length > 0 Then
                        numAddresses += 1
                    End If

                    ReDim .addressLines(numAddresses - 1)
                    numAddresses = 0
                    If Sender.Address1.Length > 0 Then
                        .addressLines(numAddresses) = Sender.Address1
                        numAddresses += 1
                    End If

                    If Sender.Address2.Length > 0 Then
                        .addressLines(numAddresses) = Sender.Address2
                        numAddresses += 1
                    End If

                    If Sender.Address3.Length > 0 Then
                        .addressLines(numAddresses) = Sender.Address3
                        numAddresses += 1
                    End If

                    .carrierRoute = Nothing

                    .cityTown = Sender.City
                    If .cityTown.Length = 0 Then
                        .cityTown = Nothing
                    End If

                    .company = Sender.Company
                    If .company.Length = 0 Then
                        .company = Nothing
                    End If

                    .countryCode = Sender.CountryCode
                    If .countryCode.Length = 0 Then
                        .countryCode = Nothing
                    End If

                    .deliveryPoint = Nothing

                    .email = Sender.eMail
                    If .email.Length = 0 Then
                        .email = Nothing
                    End If

                    .name = (Sender.FirstName & " " & Sender.LastName).ToString.Trim
                    If .name.Length = 0 Then
                        .name = Nothing
                    End If

                    .phone = Sender.Phone
                    If .phone.Length = 0 Then
                        .phone = Nothing
                    End If

                    .postalCode = Sender.ZipCode
                    If .postalCode.Length <> 5 AndAlso .postalCode.Length <> 9 Then
                        If .postalCode.Length >= 5 Then
                            .postalCode = .postalCode.Substring(0, 5)
                        End If
                    End If
                    If .postalCode.Length = 0 Then
                        .postalCode = Nothing
                    End If

                    .residential = False
                    .stateProvince = Sender.State
                    If .stateProvince.Length = 0 Then
                        .stateProvince = Nothing
                    End If

                    .taxId = Nothing
                    .status = Nothing
                End With

                With .toAddress

                    numAddresses = 0
                    If Recipient.Address1.Length > 0 Then
                        numAddresses += 1
                    End If
                    If Recipient.Address2.Length > 0 Then
                        numAddresses += 1
                    End If
                    If Recipient.Address3.Length > 0 Then
                        numAddresses += 1
                    End If

                    ReDim .addressLines(numAddresses - 1)
                    numAddresses = 0

                    If Recipient.Address1.Length > 0 Then
                        .addressLines(numAddresses) = Recipient.Address1
                        numAddresses += 1
                    End If

                    If Recipient.Address2.Length > 0 Then
                        .addressLines(numAddresses) = Recipient.Address2
                        numAddresses += 1
                    End If

                    If Recipient.Address3.Length > 0 Then
                        .addressLines(numAddresses) = Recipient.Address3
                        numAddresses += 1
                    End If

                    .carrierRoute = Nothing

                    .cityTown = Recipient.City
                    If .cityTown.Length = 0 Then
                        .cityTown = Nothing
                    End If

                    .company = Recipient.Company
                    If .company.Length = 0 Then
                        .company = Nothing
                    End If

                    .countryCode = Recipient.CountryCode
                    If .countryCode.Length = 0 Then
                        .countryCode = Nothing
                    End If

                    .deliveryPoint = Nothing

                    .email = Recipient.eMail
                    If .email.Length = 0 Then
                        .email = Nothing
                    End If

                    .name = (Recipient.FirstName & " " & Recipient.LastName).ToString.Trim
                    If .name.Length = 0 Then
                        .name = Nothing
                    End If

                    .phone = Recipient.Phone
                    If .phone.Length = 0 Then
                        .phone = Nothing
                    End If

                    .postalCode = Recipient.ZipCode
                    If .postalCode.Length <> 5 AndAlso .postalCode.Length <> 9 Then
                        If .postalCode.Length >= 5 Then
                            .postalCode = .postalCode.Substring(0, 5)
                        End If
                    End If
                    If .postalCode.Length = 0 Then
                        .postalCode = Nothing
                    End If

                    .residential = Recipient.IsResidental

                    .stateProvince = Recipient.State
                    If .stateProvince.Length = 0 Then
                        .stateProvince = Nothing
                    End If

                    .taxId = Nothing
                    .status = Nothing
                End With

                ReDim .rates(0)
                .rates(0) = New PitneyBowesRatesRequest
                With .rates(0)
                    .carrier = "USPS"
                    .parcelType = "PKG"
                    .inductionPostalCode = Sender.ZipCode
                    If .inductionPostalCode.Length > 5 Then
                        .inductionPostalCode = .inductionPostalCode.Substring(0, 5)
                    End If
                End With

                ' Need to process each carton separately

                Dim firstRequest As Boolean = True

                For Each shippingPackageDetail In PackageDetailList
                    .parcel = New PitneyBowesParcel
                    With .parcel
                        .weight.weight = shippingPackageDetail.Weight
                        .weight.unitOfMeasurement = "OZ"

                        .dimension.height = shippingPackageDetail.Height
                        .dimension.length = shippingPackageDetail.Length
                        .dimension.unitOfMeasurement = "IN"
                        .dimension.width = shippingPackageDetail.Width
                        .dimension.irregularParcelGirth = 2 * (shippingPackageDetail.Height + shippingPackageDetail.Width)
                    End With

                    Dim url As String = cServer & "/v1/rates?includeDeliveryCommitment=true"
                    Dim headers As New System.Net.WebHeaderCollection
                    headers.Add("Authorization", "Bearer " & oAuthToken)
                    headers.Add("X-PB-UnifiedErrorStructure", "true")
                    headers.Add("X-PB-Shipper-Rate-Plan", "VSCM_SRP_NEWBLUE02")

                    Dim contentType As String = "application/json"
                    Dim method As String = "POST"

                    Dim postResult As String = PostHttpRequest(url, ratesObject, String.Empty, contentType, method, headers)

                    If postResult = Nothing Then
                        Return Nothing
                    End If

                    If postResult.Length = 0 Then
                        Return Nothing
                    End If

                    If postResult.StartsWith("{""errors") Then
                        LastError = postResult
                        Return Nothing
                    End If

                    Dim ratesResponse As PitneyBowesRatesResponseObject = JsonConvert.DeserializeObject(Of PitneyBowesRatesResponseObject)(postResult)

                    ' use the first package as the base
                    If firstRequest Then
                        For iloop As Int32 = 0 To ratesResponse.rates.Count - 1
                            Dim rate As PitneyBowesRatesResponse = ratesResponse.rates(iloop)
                            Dim serviceType As String = ""

                            If PitneyBowesServiceDictionary.ContainsKey(rate.serviceId) Then
                                serviceType = PitneyBowesServiceDictionary(rate.serviceId).ServiceCode
                                dicReturnRates.Add(rate.serviceId, New RateList)
                            End If
                        Next
                        firstRequest = False
                    Else
                        ' if we received a shipping method for a previous carton but not on this carton then remove the service from list of services
                        Dim lstDeleteService As New List(Of String)

                        For Each kvp As KeyValuePair(Of String, RateList) In dicReturnRates
                            Dim serviceKey As String = kvp.Key
                            Dim found As Boolean = False

                            For iloop As Int32 = 0 To ratesResponse.rates.Count - 1
                                Dim rate As PitneyBowesRatesResponse = ratesResponse.rates(iloop)
                                If serviceKey = rate.serviceId Then
                                    found = True
                                    Exit For
                                End If
                            Next

                            If Not found Then
                                lstDeleteService.Add(serviceKey)
                            End If
                        Next

                        For Each key As String In lstDeleteService
                            dicReturnRates.Remove(key)
                        Next

                    End If

                    For iloop As Int32 = 0 To ratesResponse.rates.Count - 1
                        Dim rate As PitneyBowesRatesResponse = ratesResponse.rates(iloop)
                        Dim serviceType As String = ""

                        If Not PitneyBowesServiceDictionary.ContainsKey(rate.serviceId) Then
                            Continue For
                        End If

                        serviceType = rate.serviceId

                        If Not dicReturnRates.ContainsKey(serviceType) Then
                            Continue For
                        End If

                        Dim dictRateList As New RateList
                        dictRateList = dicReturnRates(serviceType)

                        With dictRateList

                            If rate.alternateBaseCharge > 0 Then
                                .AccountNetCharge += rate.alternateBaseCharge
                            Else
                                .AccountNetCharge += rate.totalCarrierCharge
                            End If

                            Dim estdate As String = rate.deliveryCommitment.estimatedDeliveryDateTime & String.Empty
                            Dim dt() As String = Split(estdate, "-")
                            If dt.Length = 3 Then
                                estdate = dt(1) & "/" & dt(2) & "/" & dt(0)
                                If Not IsDate(estdate) Then
                                    estdate = rate.deliveryCommitment.estimatedDeliveryDateTime & String.Empty
                                End If
                            End If

                            If Not IsDate(.DeliveryDate) Then
                                .DeliveryDate = estdate
                                .DeliveryTime = rate.deliveryCommitment.additionalDetails & String.Empty
                                .TransitTime = rate.deliveryCommitment.maxEstimatedNumberOfDays & String.Empty
                            ElseIf IsDate(estdate) Then
                                ' use the max date
                                Dim d1 = DateTime.Parse(.DeliveryDate)
                                Dim d2 = DateTime.Parse(estdate)

                                If (d2 - d1).Days > 0 Then
                                    .DeliveryDate = estdate
                                    .DeliveryTime = rate.deliveryCommitment.additionalDetails & String.Empty
                                    .TransitTime = rate.deliveryCommitment.maxEstimatedNumberOfDays & String.Empty
                                End If
                            End If

                            If rate.alternateTotalCharge > 0 Then
                                .ListNetCharge += rate.alternateTotalCharge
                            Else
                                .ListNetCharge += rate.totalCarrierCharge
                            End If

                            .OfferID = ""

                            If Val(.ReferenceIndex) = 0 Then
                                .ReferenceIndex = iloop + 1
                            End If

                            If PitneyBowesServiceDictionary.ContainsKey(rate.serviceId) Then
                                .ServiceType = PitneyBowesServiceDictionary(rate.serviceId).ServiceCode
                                .ServiceTypeDescription = PitneyBowesServiceDictionary(rate.serviceId).ServiceDescription
                            End If

                            .ServiceCode = rate.serviceId

                        End With

                        dicReturnRates(serviceType) = dictRateList
                    Next
                Next

            End With

            If dicReturnRates.Count = 0 Then
                Return Nothing
            End If

            ReDim requestedRateList(dicReturnRates.Count - 1)
            Dim ictr As Int16 = 0
            For Each kvp As KeyValuePair(Of String, RateList) In dicReturnRates
                requestedRateList(ictr) = kvp.Value
                ictr += 1
            Next

            Return requestedRateList

        Catch exn As Newtonsoft.Json.JsonException
            LastError = exn.Message
            Return Nothing

        Catch ex As Exception
            LastError = ex.Message
            Return Nothing
        End Try

    End Function

    Private Function RequestPitneyBowesLabel() As Boolean

        RequestPitneyBowesLabel = False
        cMasterTrackingNumber = String.Empty

        Try
            Dim oAuthToken As String = GetoAuthToken("", "USPS")

            If oAuthToken Is Nothing OrElse oAuthToken.Length = 0 Then
                LastError = "oAuth Token could not be found."
                Return False
            End If


            Dim shipLabelObject As New PitneyBowesLabelRequestObject
            Dim numAddresses As Int16 = 0

            With shipLabelObject

                ReDim .shipmentOptions(1)
                .shipmentOptions(0) = New PitneyBowesShipmentOption()
                With .shipmentOptions(0)
                    .name = "SHIPPER_ID"
                    .value = cUserId
                End With

                .shipmentOptions(1) = New PitneyBowesShipmentOption()
                With .shipmentOptions(1)
                    .name = "HIDE_TOTAL_CARRIER_CHARGE"
                    .value = "true"
                End With

                '.shipmentOptions(2) = New PitneyBowesShipmentOption()
                'With .shipmentOptions(2)
                '    .name = "ADD_TO_MANIFEST"
                '    .value = "true"
                'End With

                '.shipmentOptions(3) = New PitneyBowesShipmentOption()
                'With .shipmentOptions(3)
                '    .name = "MINIMAL_ADDRESS_VALIDATION"
                '    .value = "true"
                'End With

                If Account.Address1.Length > 0 AndAlso Account.City.Length > 0 AndAlso Account.State.Length > 0 Then
                    With .fromAddress
                        numAddresses = 0
                        If Account.Address1.Length > 0 Then
                            numAddresses += 1
                        End If
                        If Account.Address2.Length > 0 Then
                            numAddresses += 1
                        End If
                        If Account.Address3.Length > 0 Then
                            numAddresses += 1
                        End If

                        ReDim .addressLines(numAddresses - 1)
                        numAddresses = 0
                        If Account.Address1.Length > 0 Then
                            .addressLines(numAddresses) = Account.Address1
                            numAddresses += 1
                        End If

                        If Account.Address2.Length > 0 Then
                            .addressLines(numAddresses) = Account.Address2
                            numAddresses += 1
                        End If

                        If Account.Address3.Length > 0 Then
                            .addressLines(numAddresses) = Account.Address3
                            numAddresses += 1
                        End If

                        .carrierRoute = Nothing

                        .cityTown = Account.City
                        If .cityTown.Length = 0 Then
                            .cityTown = Nothing
                        End If

                        .company = Account.Company
                        If .company.Length = 0 Then
                            .company = Nothing
                        End If

                        .countryCode = Account.CountryCode
                        If .countryCode.Length = 0 Then
                            .countryCode = Nothing
                        End If

                        .deliveryPoint = Nothing

                        .email = Account.eMail
                        If .email.Length = 0 Then
                            .email = Nothing
                        End If

                        .name = (Account.FirstName & " " & Account.LastName).ToString.Trim
                        If .name.Length = 0 Then
                            .name = Nothing
                        End If

                        .phone = Account.Phone
                        If .phone.Length = 0 Then
                            .phone = Nothing
                        End If

                        .postalCode = Account.ZipCode
                        If .postalCode.Length <> 5 AndAlso .postalCode.Length <> 9 Then
                            If .postalCode.Length >= 5 Then
                                .postalCode = .postalCode.Substring(0, 5)
                            End If
                        End If
                        If .postalCode.Length = 0 Then
                            .postalCode = Nothing
                        End If

                        .residential = False
                        .stateProvince = Account.State
                        If .stateProvince.Length = 0 Then
                            .stateProvince = Nothing
                        End If

                        .taxId = Nothing
                        .status = Nothing
                    End With
                Else

                    With .fromAddress
                        numAddresses = 0
                        If Sender.Address1.Length > 0 Then
                            numAddresses += 1
                        End If
                        If Sender.Address2.Length > 0 Then
                            numAddresses += 1
                        End If
                        If Sender.Address3.Length > 0 Then
                            numAddresses += 1
                        End If

                        ReDim .addressLines(numAddresses - 1)
                        numAddresses = 0
                        If Sender.Address1.Length > 0 Then
                            .addressLines(numAddresses) = Sender.Address1
                            numAddresses += 1
                        End If

                        If Sender.Address2.Length > 0 Then
                            .addressLines(numAddresses) = Sender.Address2
                            numAddresses += 1
                        End If

                        If Sender.Address3.Length > 0 Then
                            .addressLines(numAddresses) = Sender.Address3
                            numAddresses += 1
                        End If

                        .carrierRoute = Nothing

                        .cityTown = Sender.City
                        If .cityTown.Length = 0 Then
                            .cityTown = Nothing
                        End If

                        .company = Sender.Company
                        If .company.Length = 0 Then
                            .company = Nothing
                        End If

                        .countryCode = Sender.CountryCode
                        If .countryCode.Length = 0 Then
                            .countryCode = Nothing
                        End If

                        .deliveryPoint = Nothing

                        .email = Sender.eMail
                        If .email.Length = 0 Then
                            .email = Nothing
                        End If

                        .name = (Sender.FirstName & " " & Sender.LastName).ToString.Trim
                        If .name.Length = 0 Then
                            .name = Nothing
                        End If

                        .phone = Sender.Phone
                        If .phone.Length = 0 Then
                            .phone = Nothing
                        End If

                        .postalCode = Sender.ZipCode
                        If .postalCode.Length <> 5 AndAlso .postalCode.Length <> 9 Then
                            If .postalCode.Length >= 5 Then
                                .postalCode = .postalCode.Substring(0, 5)
                            End If
                        End If
                        If .postalCode.Length = 0 Then
                            .postalCode = Nothing
                        End If

                        .residential = False
                        .stateProvince = Sender.State
                        If .stateProvince.Length = 0 Then
                            .stateProvince = Nothing
                        End If

                        .taxId = Nothing
                        .status = Nothing
                    End With

                End If

                With .toAddress
                    numAddresses = 0
                    If Recipient.Address1.Length > 0 Then
                        numAddresses += 1
                    End If
                    If Recipient.Address2.Length > 0 Then
                        numAddresses += 1
                    End If
                    If Recipient.Address3.Length > 0 Then
                        numAddresses += 1
                    End If

                    ReDim .addressLines(numAddresses - 1)
                    numAddresses = 0
                    If Recipient.Address1.Length > 0 Then
                        .addressLines(numAddresses) = Recipient.Address1
                        numAddresses += 1
                    End If

                    If Recipient.Address2.Length > 0 Then
                        .addressLines(numAddresses) = Recipient.Address2
                        numAddresses += 1
                    End If

                    If Recipient.Address3.Length > 0 Then
                        .addressLines(numAddresses) = Recipient.Address3
                        numAddresses += 1
                    End If

                    .carrierRoute = Nothing

                    .cityTown = Recipient.City
                    If .cityTown.Length = 0 Then
                        .cityTown = Nothing
                    End If

                    .company = Recipient.Company
                    If .company.Length = 0 Then
                        .company = Nothing
                    End If

                    .countryCode = Recipient.CountryCode
                    If .countryCode.Length = 0 Then
                        .countryCode = Nothing
                    End If

                    .deliveryPoint = Nothing

                    .email = Recipient.eMail
                    If .email.Length = 0 Then
                        .email = Nothing
                    End If

                    .name = (Recipient.FirstName & " " & Recipient.LastName).ToString.Trim
                    If .name.Length = 0 Then
                        .name = Nothing
                    End If

                    .phone = Recipient.Phone
                    If .phone.Length = 0 Then
                        .phone = Nothing
                    End If

                    .postalCode = Recipient.ZipCode
                    If .postalCode.Length <> 5 AndAlso .postalCode.Length <> 9 Then
                        If .postalCode.Length >= 5 Then
                            .postalCode = .postalCode.Substring(0, 5)
                        End If
                    End If
                    If .postalCode.Length = 0 Then
                        .postalCode = Nothing
                    End If

                    .residential = Recipient.IsResidental

                    .stateProvince = Recipient.State
                    If .stateProvince.Length = 0 Then
                        .stateProvince = Nothing
                    End If

                    .taxId = Nothing
                    .status = Nothing
                End With

                ReDim .rates(0)
                .rates(0) = New PitneyBowesRatesRequest
                With .rates(0)
                    .carrier = "USPS"

                    For Each kvp As KeyValuePair(Of String, PitneyBowesServices) In PitneyBowesServiceDictionary
                        Dim serviceKey As String = kvp.Key

                        If Val(PitneyBowesServiceDictionary(serviceKey).ServiceCode) = Val(cRequestedServiceType) Then
                            .serviceId = serviceKey
                            Exit For
                        End If
                    Next

                    .parcelType = "PKG"

                    .inductionPostalCode = Recipient.ZipCode
                    If PitneyBowesInductionPostalCode IsNot Nothing AndAlso PitneyBowesInductionPostalCode.Length > 0 Then
                        .inductionPostalCode = PitneyBowesInductionPostalCode
                    End If

                    If .inductionPostalCode.Length > 5 Then
                        .inductionPostalCode = .inductionPostalCode.Substring(0, 5)
                    End If

                    ' prevent "errorCode":"1022026", errorDescription":"The Service or at least one of the included special services must be trackable.",
                    ReDim .specialServices(0)
                    .specialServices(0) = New PitneyBowesSpecialServices
                    With .specialServices(0)
                        .specialServiceId = "DelCon"

                        ReDim .inputParameters(0)
                        .inputParameters(0) = New PitneyBowesInputParameters
                        With .inputParameters(0)
                            .name = "INPUT_VALUE"
                            .value = "0"
                        End With
                        .fee = Nothing
                    End With

                End With

                ReDim .documents(0)
                .documents(0) = New PitneyBowesDocument
                With .documents(0)
                    .contentType = "BASE64"
                    .fileFormat = "ZPL2"
                    .printDialogOption = "NO_PRINT_DIALOG"
                    .resolution = "DPI_203"
                    .size = "DOC_4X6"
                    .type = "SHIPPING_LABEL"

                    .contents = Nothing
                    .docTab = Nothing
                    .pages = Nothing
                End With

                For Each shippingPackageDetail In PackageDetailList
                    .parcel = New PitneyBowesParcel
                    With .parcel
                        .weight.weight = shippingPackageDetail.Weight
                        .weight.unitOfMeasurement = "OZ"

                        .dimension.height = shippingPackageDetail.Height
                        .dimension.length = shippingPackageDetail.Length
                        .dimension.unitOfMeasurement = "IN"
                        .dimension.width = shippingPackageDetail.Width
                        .dimension.irregularParcelGirth = 2 * (shippingPackageDetail.Height + shippingPackageDetail.Width)
                    End With

                    Dim url As String = cServer & "/v1/shipments?includeDeliveryCommitment=true"
                    Dim headers As New System.Net.WebHeaderCollection
                    headers.Add("Authorization", "Bearer " & oAuthToken)
                    headers.Add("X-PB-UnifiedErrorStructure", "true")
                    headers.Add("X-PB-Shipper-Rate-Plan", "VSCM_SRP_NEWBLUE02")

                    headers.Add("X-PB-TransactionId", PitneyBowesUniqueTransactionID & "_" & Val(shippingPackageDetail.Id)) ' Required. A unique identifier for the transaction, up to 25 characters.
                    Dim contentType As String = "application/json"
                    Dim method As String = "POST"

                    Dim postResult As String = PostHttpRequest(url, shipLabelObject, String.Empty, contentType, method, headers)

                    If postResult = Nothing Then
                        LastError = "No label return by Pitney Bowes"
                        Return False
                    End If

                    If postResult.Length = 0 Then
                        LastError = "No label return by Pitney Bowes"
                        Return False
                    End If

                    If postResult.StartsWith("{""errors") Then
                        LastError = postResult
                        Return False
                    End If

                    Dim labelResponse As PitneyBowesLabelResponsetObject = JsonConvert.DeserializeObject(Of PitneyBowesLabelResponsetObject)(postResult)

                    Dim key As Int32 = Val(shippingPackageDetail.Id)
                    Dim extension As String = ".zpl"
                    Dim charge As Decimal = labelResponse.rates(0).alternateBaseCharge
                    If charge <= 0 Then
                        charge = labelResponse.rates(0).baseCharge
                    End If
                    ShipmentBaseCharge.Add(key, charge)
                    ShipmentDiscountCharge.Add(key, 0)
                    ShipmentSurCharge.Add(key, 0)

                    charge = labelResponse.rates(0).alternateTotalCharge
                    If charge <= 0 Then
                        charge = labelResponse.rates(0).totalCarrierCharge
                    End If

                    ShipmentNetCharge.Add(key, charge)
                    ShipmentListCharge.Add(key, charge)

                    Dim PitneyBowesPackageInformation As New PitneyBowesPackageInformation
                    With PitneyBowesPackageInformation
                        .TrackingNumber = labelResponse.parcelTrackingNumber
                        .ShipmentID = labelResponse.shipmentId
                    End With

                    If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                        If Not ShippingLabelDirectory.EndsWith("\") Then
                            ShippingLabelDirectory &= "\"
                        End If

                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & key & extension
                        Using sw As New StreamWriter(shippingPackageDetail.ShippingLabelFile)
                            For iDoc As Int16 = 0 To labelResponse.documents.Count - 1
                                For ipage As Int16 = 0 To labelResponse.documents(iDoc).pages.Count - 1
                                    Dim content As String = labelResponse.documents(iDoc).pages(ipage).contents
                                    Dim base64Encoded As String = content
                                    Dim base64Decoded As String
                                    Dim data() As Byte
                                    data = System.Convert.FromBase64String(base64Encoded)
                                    base64Decoded = System.Text.ASCIIEncoding.ASCII.GetString(data)

                                    If base64Decoded.Length > 0 Then
                                        sw.WriteLine(base64Decoded)
                                        PitneyBowesPackageInformation.ShippingLabel &= base64Decoded
                                    End If
                                Next
                            Next

                            sw.Close()
                            sw.Dispose()
                        End Using
                    End If

                    Dim jSonData As String = JsonConvert.SerializeObject(PitneyBowesPackageInformation, Newtonsoft.Json.Formatting.Indented)
                    shippingPackageDetail.Reference = jSonData
                    cMasterTrackingNumber = labelResponse.parcelTrackingNumber
                Next
            End With

            RequestPitneyBowesLabel = True

        Catch ex As Exception
            LastError = ex.Message
            RequestPitneyBowesLabel = False
        End Try

    End Function

    Private Function VoidShipment(ByVal ShipmentID As String) As Boolean

        LastError = String.Empty

        Try

            Dim oAuthToken As String = GetoAuthToken("", "USPS")
            Dim numAddresses As Decimal = 0
            Dim cancelRequest As New PitneyBowesCancelLabelRequest

            If oAuthToken Is Nothing OrElse oAuthToken.Length = 0 Then
                LastError = "Could not access oAuth Token"
                Return False
            End If


            'curl -X DELETE .../v1/shipments/USPS2200116649182781 \
            '-H "Authorization: Bearer N7WPfrO0AfgDWXuKRTDlxpOKY3sG" \
            '-H "Content-Type: application/json"  \
            '-H "X-PB-TransactionId: 252d8bac-7520-4209-bff6-97b7" \
            '-H "X-PB-UnifiedErrorStructure: true"  \
            '-d '
            '{
            '    "carrier": "USPS",
            '    "cancelInitiator": "SHIPPER"
            '}'

            cancelRequest.carrier = "USPS"
            cancelRequest.cancelInitiator = "SHIPPER"

            Dim url As String = cServer & "/v1/shipments/" & ShipmentID
            Dim headers As New System.Net.WebHeaderCollection
            headers.Add("Authorization", "Bearer " & oAuthToken)
            headers.Add("X-PB-UnifiedErrorStructure", "true")
            headers.Add("X-PB-TransactionId", PitneyBowesUniqueTransactionID) ' Required. A unique identifier for the transaction, up to 25 characters.

            Dim contentType As String = "application/json"
            Dim method As String = "DELETE"

            Dim postResult As String = PostHttpRequest(url, cancelRequest, String.Empty, contentType, method, headers)

            If postResult = Nothing Then
                Return Nothing
            End If

            If postResult.Length = 0 Then
                Return Nothing
            End If

            If postResult.StartsWith("{""errors") Then
                LastError = postResult
                Return Nothing
            End If

            Dim cancelResponse As PitneyBowesCancelLabelResponse = JsonConvert.DeserializeObject(Of PitneyBowesCancelLabelResponse)(postResult)

            If cancelResponse.status & String.Empty = "INITIATED" Then
                Return True
            Else
                LastError = "Cancel Status returned by Pitney Bowes: " & cancelResponse.cancelInitiator & String.Empty
                Return False
            End If

        Catch exn As Newtonsoft.Json.JsonException
            LastError = exn.Message
            Return False

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function


#End Region

#Region "Federal Express"

    ''' <summary>
    ''' Request Fedex International Shipping label
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestFedexInternaltionalLabel() As Boolean

        Try
            LastError = String.Empty
            objFedexShipIntl = New nsoftware.InShip.Fedexshipintl
            objFedexShipIntl.RuntimeLicense = inShipLicense
            objFedexShipIntl.Reset()
            objFedexShipIntl.RuntimeLicense = inShipLicense

            If cServiceProvider <> ServiceProviders.FederalExpressInternational Then
                LastError = "Service type not a Fedex International."
                Return False
            End If

            ' Set credentials
            objFedexShipIntl.FedExAccount.Server = cServer
            objFedexShipIntl.FedExAccount.DeveloperKey = cFedexDeveloperKey
            objFedexShipIntl.FedExAccount.Password = cPassword
            objFedexShipIntl.FedExAccount.AccountNumber = cAccountNumber
            objFedexShipIntl.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexShipIntl.ServiceType = cRequestedServiceType

            If cServer.ToUpper.Contains("WEB-SERVICES") Then
                objFedexShipIntl.Config("UseSOAP=true")
            Else
                objFedexShipIntl.Config("UseSOAP=false")
            End If

            objFedexShipIntl.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            ' Get Sender Information
            With cSenderContact
                objFedexShipIntl.SenderContact.FirstName = .FirstName
                objFedexShipIntl.SenderContact.LastName = .LastName
                objFedexShipIntl.SenderContact.MiddleInitial = .MiddleInitial
                objFedexShipIntl.SenderContact.Phone = .Phone
                objFedexShipIntl.SenderContact.Fax = .Fax
                objFedexShipIntl.SenderContact.Email = .eMail

                objFedexShipIntl.SenderContact.Company = .Company
                objFedexShipIntl.SenderAddress.Address1 = .Address1
                objFedexShipIntl.SenderAddress.Address2 = .Address2
                'objFedexShipIntl.Config("SenderAddress3=" & .Address3)

                objFedexShipIntl.SenderAddress.City = .City
                objFedexShipIntl.SenderAddress.ZipCode = .ZipCode
                objFedexShipIntl.SenderAddress.State = .State
                objFedexShipIntl.SenderAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objFedexShipIntl.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexShipIntl.SenderAddress.AddressFlags = &H1 'PO Box
                End If
            End With

            With cRecipientContact
                objFedexShipIntl.RecipientContact.FirstName = .FirstName
                objFedexShipIntl.RecipientContact.LastName = .LastName
                objFedexShipIntl.RecipientContact.MiddleInitial = .MiddleInitial
                objFedexShipIntl.RecipientContact.Phone = .Phone
                objFedexShipIntl.RecipientContact.Fax = .Fax
                objFedexShipIntl.RecipientContact.Email = .eMail

                objFedexShipIntl.RecipientContact.Company = .Company
                objFedexShipIntl.RecipientAddress.Address1 = .Address1
                objFedexShipIntl.RecipientAddress.Address2 = .Address2
                'objFedexShipIntl.Config("RecipientAddress3=" & .Address3)

                objFedexShipIntl.RecipientAddress.City = .City
                objFedexShipIntl.RecipientAddress.ZipCode = .ZipCode
                objFedexShipIntl.RecipientAddress.State = .State
                objFedexShipIntl.RecipientAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objFedexShipIntl.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexShipIntl.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            With cReturnAddress
                If .Company & String.Empty <> String.Empty Then
                    objFedexShipIntl.ReturnContact.FirstName = .FirstName
                    objFedexShipIntl.ReturnContact.LastName = .LastName
                    objFedexShipIntl.ReturnContact.MiddleInitial = .MiddleInitial
                    objFedexShipIntl.ReturnContact.Phone = .Phone
                    objFedexShipIntl.ReturnContact.Fax = .Fax
                    objFedexShipIntl.ReturnContact.Email = .eMail

                    objFedexShipIntl.ReturnContact.Company = .Company
                    objFedexShipIntl.ReturnAddress.Address1 = .Address1
                    objFedexShipIntl.ReturnAddress.Address2 = .Address2
                    'objFedexShipIntl.Config("ReturnAddress3=" & .Address3)

                    objFedexShipIntl.ReturnAddress.City = .City
                    objFedexShipIntl.ReturnAddress.ZipCode = .ZipCode
                    objFedexShipIntl.ReturnAddress.State = .State
                    objFedexShipIntl.ReturnAddress.CountryCode = .CountryCode
                End If
            End With

            Select Case EzshipLabelImage
                Case EzshipLabelImageTypes.itEltron
                    objFedexShipIntl.LabelImageType = FedexshipintlLabelImageTypes.fitEltron
                Case EzshipLabelImageTypes.itPDF
                    objFedexShipIntl.LabelImageType = FedexshipintlLabelImageTypes.fitPDF
                Case EzshipLabelImageTypes.itPNG
                    objFedexShipIntl.LabelImageType = FedexshipintlLabelImageTypes.fitPNG
                Case EzshipLabelImageTypes.itUniMark
                    objFedexShipIntl.LabelImageType = FedexshipintlLabelImageTypes.fitUniMark
                Case EzshipLabelImageTypes.itZebra
                    objFedexShipIntl.LabelImageType = FedexshipintlLabelImageTypes.fitZebra
                Case Else ' if not a valid option default to fitEltron
                    objFedexShipIntl.LabelImageType = FedexshipintlLabelImageTypes.fitEltron
            End Select

            Dim extension As String = objFedexShipIntl.LabelImageType.ToString
            If extension.StartsWith("fit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            If LabelStockType.Length > 0 Then
                objFedexShipIntl.Config("LabelStockType=" & LabelStockType)
            End If

            ' Set Shipping Label File
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If Not ShippingLabelDirectory.EndsWith("\") Then
                    ShippingLabelDirectory &= "\"
                End If

                For Each shippingPackageDetail In PackageDetailList
                    Dim id As String = idCtr.ToString
                    idCtr += 1
                    shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                    shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                Next
            End If

            Dim totalWeight As Double = 0
            Dim totalInsured As Double = 0

            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objFedexShipIntl.Packages.Add(shippingPackageDetail)
            Next

            objFedexShipIntl.TotalWeight = Format(totalWeight, "###0.0")
            objFedexShipIntl.InsuredValue = Format(totalInsured, "###0.00")
            objFedexShipIntl.ShipDate = ShipDate.ToString("yyyy-MM-dd")
            objFedexShipIntl.TotalCustomsValue = Format(TotalCustomsValue, "###0.00")

            ' Service Type
            objFedexShipIntl.ServiceType = cRequestedServiceType
            objFedexShipIntl.DropoffType = DropOffType
            objFedexShipIntl.Payor.PayorType = Payor

            objFedexShipIntl.Payor.AccountNumber = cSenderContact.AccountNumber
            objFedexShipIntl.Payor.CountryCode = cSenderContact.CountryCode

            objFedexShipIntl.DutiesPayor.PayorType = DutiesPayor
            objFedexShipIntl.DutiesPayor.AccountNumber = cDutiesPayorContact.AccountNumber
            objFedexShipIntl.DutiesPayor.CountryCode = cDutiesPayorContact.CountryCode

            Dim specialService As Long = ShipmentSpecialServices

            If objFedexShipIntl.ShipDate > DateTime.Now.ToString("yyyy-MM-dd") Then
                specialService = specialService Or &H20000000L
            End If

            objFedexShipIntl.ShipmentSpecialServices = specialService

            objFedexShipIntl.HoldAtLocation.Address1 = cHoldAtLocation.Address1
            objFedexShipIntl.HoldAtLocation.Address2 = cHoldAtLocation.Address2
            objFedexShipIntl.HoldAtLocation.City = cHoldAtLocation.City
            objFedexShipIntl.HoldAtLocation.State = cHoldAtLocation.State
            objFedexShipIntl.HoldAtLocation.ZipCode = cHoldAtLocation.ZipCode
            objFedexShipIntl.HoldAtLocationPhone = cHoldAtLocation.Phone

            With objFedexShipIntl.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            With objFedexShipIntl.DutiesPayor
                .PayorType = DutiesPayor
                .AccountNumber = DutiesPayorContact.AccountNumber
                .CountryCode = DutiesPayorContact.CountryCode
                .ZipCode = DutiesPayorContact.ZipCode
            End With

            For Each CommDetail As CommodityDetail In CommodityDetailList
                CommDetail.Weight = Format(Val(CommDetail.Weight), "###0.0")
                CommDetail.Description = CommDetail.Description.Replace("&", " ").Replace("<", " ").Replace(">", " ")
                objFedexShipIntl.Commodities.Add(CommDetail)
            Next

        Catch ex As Exception
            objFedexShipIntl.Dispose()
            objFedexShipIntl = Nothing
            Return False
        End Try

        ' Notifications - Not Supported by FedEx International
        Dim notificationsIndex As Int16 = 0
        If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
            For Each sn As Notifications In ShipmentNotifications
                sn.email = (sn.email & String.Empty).Trim
                If sn.email.Length = 0 Then
                    Continue For
                End If

                Dim notify As New nsoftware.InShip.NotifyDetail
                With notify
                    .Email = sn.email
                    .NotificationFlags = CInt(sn.NotificationFlags)
                    .Message = (sn.Message & String.Empty).ToString.Trim
                End With

                notificationsIndex += 1
                If notificationsIndex = 3 Then Exit For
            Next
        End If

        Try

            objFedexShipIntl.RuntimeLicense = inShipLicense
            objFedexShipIntl.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objFedexShipIntl.Packages.Count - 1
                PackageDetailList.Add(objFedexShipIntl.Packages(ictr))
                GetPackageCosts(objFedexShipIntl.Packages(ictr), objFedexShipIntl)
            Next

            If objFedexShipIntl.Packages.Count = 1 Then
                cMasterTrackingNumber = objFedexShipIntl.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objFedexShipIntl.MasterTrackingNumber
            End If

            Return True

        Catch ex As nsoftware.InShip.InShipFedexshipintlException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objFedexShipIntl.Config("RawRequest")
            cRawResponse = objFedexShipIntl.Config("RawResponse")
            objFedexShipIntl.Dispose()
            objFedexShipIntl = Nothing
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Request Shipping label. Not used for Fedex Intenational
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestFedexLabel() As Boolean

        Try
            LastError = String.Empty
            objFedexShip = New nsoftware.InShip.Fedexship
            objFedexShip.RuntimeLicense = inShipLicense
            objFedexShip.Reset()
            objFedexShip.RuntimeLicense = inShipLicense

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            objFedexShip.FedExAccount.Server = cServer
            objFedexShip.FedExAccount.AccountNumber = cAccountNumber
            objFedexShip.FedExAccount.Password = cPassword
            objFedexShip.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexShip.FedExAccount.DeveloperKey = cFedexDeveloperKey

            If cServer.ToUpper.Contains("WEB-SERVICES") Then
                objFedexShip.Config("UseSOAP=true")
            Else
                objFedexShip.Config("UseSOAP=false")
            End If

            objFedexShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            objFedexShip.ServiceType = cRequestedServiceType

            If objFedexShip.ServiceType = ServiceTypes.stFedExSmartPost Then
                objFedexShip.Config("SmartPostIndicia=" & FedexSmartPost.Indicia)
                objFedexShip.Config("SmartPostHubId=" & FedexSmartPost.HubId)
                objFedexShip.Config("SmartPostPhysicalPackaging=" & FedexSmartPost.PhysicalPackaging)
                objFedexShip.Config("SmartPostAncillaryEndorsement=" & FedexSmartPost.AncillaryEndorsement)
            End If

            objFedexShip.ShipDate = ShipDate.ToString("yyyy-MM-dd")

            With cSenderContact
                objFedexShip.SenderContact.FirstName = .FirstName
                objFedexShip.SenderContact.LastName = .LastName
                objFedexShip.SenderContact.MiddleInitial = .MiddleInitial
                objFedexShip.SenderContact.Phone = .Phone
                objFedexShip.SenderContact.Fax = .Fax
                objFedexShip.SenderContact.Email = .eMail

                objFedexShip.SenderContact.Company = .Company
                objFedexShip.SenderAddress.Address1 = .Address1
                objFedexShip.SenderAddress.Address2 = .Address2
                'objFedexShip.Config("SenderAddress3=" & .Address3)

                objFedexShip.SenderAddress.City = .City
                objFedexShip.SenderAddress.ZipCode = .ZipCode
                objFedexShip.SenderAddress.State = .State
                objFedexShip.SenderAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    'objFedexShip.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    'objFedexShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            With cAccountContact
                If .Company & String.Empty <> String.Empty Then

                    objFedexShip.OriginContact.FirstName = objFedexShip.SenderContact.FirstName
                    objFedexShip.OriginContact.LastName = objFedexShip.SenderContact.LastName
                    objFedexShip.OriginContact.MiddleInitial = objFedexShip.SenderContact.MiddleInitial
                    objFedexShip.OriginContact.Phone = objFedexShip.SenderContact.Phone
                    objFedexShip.OriginContact.Fax = objFedexShip.SenderContact.Fax
                    objFedexShip.OriginContact.Email = objFedexShip.SenderContact.Email

                    objFedexShip.OriginContact.Company = objFedexShip.SenderContact.Company
                    objFedexShip.OriginAddress.Address1 = objFedexShip.SenderAddress.Address1
                    objFedexShip.OriginAddress.Address2 = objFedexShip.SenderAddress.Address2
                    'objFedexShip.Config("AccountAddress3=" & .Address3)

                    objFedexShip.OriginAddress.City = objFedexShip.SenderAddress.City
                    objFedexShip.OriginAddress.ZipCode = objFedexShip.SenderAddress.ZipCode
                    objFedexShip.OriginAddress.State = objFedexShip.SenderAddress.State
                    objFedexShip.OriginAddress.CountryCode = objFedexShip.SenderAddress.CountryCode

                    objFedexShip.OriginAddress.CountryCode = objFedexShip.OriginAddress.CountryCode.ToUpper
                    If objFedexShip.OriginAddress.CountryCode = "" OrElse objFedexShip.OriginAddress.CountryCode = "USA" Then
                        objFedexShip.OriginAddress.CountryCode = "US"
                    End If

                    If objFedexShip.OriginAddress.CountryCode = "CAN" Then
                        objFedexShip.OriginAddress.CountryCode = "CA"
                    End If

                    If .IsResidental Then
                        'objFedexShip.OriginAddress.AddressFlags = &H2 'Residential
                    ElseIf .IsPOBox Then
                        'objFedexShip.OriginAddress.AddressFlags = &H1 'PO Box
                    End If

                    objFedexShip.SenderContact.FirstName = .FirstName
                    objFedexShip.SenderContact.LastName = .LastName
                    objFedexShip.SenderContact.MiddleInitial = .MiddleInitial
                    objFedexShip.SenderContact.Phone = .Phone
                    objFedexShip.SenderContact.Fax = .Fax
                    objFedexShip.SenderContact.Email = .eMail

                    objFedexShip.SenderContact.Company = .Company
                    objFedexShip.SenderAddress.Address1 = .Address1
                    objFedexShip.SenderAddress.Address2 = .Address2
                    'objFedexShip.Config("SenderAddress3=" & .Address3)

                    objFedexShip.SenderAddress.City = .City
                    objFedexShip.SenderAddress.ZipCode = .ZipCode
                    objFedexShip.SenderAddress.State = .State
                    objFedexShip.SenderAddress.CountryCode = .CountryCode

                    If .IsResidental Then
                        'objFedexShip.SenderAddress.AddressFlags = &H2 'Residential
                    ElseIf .IsPOBox Then
                        'objFedexShip.RecipientAddress.AddressFlags = &H1 'PO Box
                    End If

                End If
            End With

            With cRecipientContact
                objFedexShip.RecipientContact.FirstName = .FirstName
                objFedexShip.RecipientContact.LastName = .LastName
                objFedexShip.RecipientContact.MiddleInitial = .MiddleInitial
                objFedexShip.RecipientContact.Phone = .Phone
                objFedexShip.RecipientContact.Fax = .Fax
                objFedexShip.RecipientContact.Email = .eMail

                objFedexShip.RecipientContact.Company = .Company
                objFedexShip.RecipientAddress.Address1 = .Address1
                objFedexShip.RecipientAddress.Address2 = .Address2
                'objFedexShip.Config("RecipientAddress3=" & .Address3)

                objFedexShip.RecipientAddress.City = .City
                objFedexShip.RecipientAddress.ZipCode = .ZipCode
                objFedexShip.RecipientAddress.State = .State
                objFedexShip.RecipientAddress.CountryCode = .CountryCode

                If .IsResidental Then
                    objFedexShip.RecipientAddress.AddressFlags = &H2 'Residential
                    If objFedexShip.ServiceType = ServiceTypes.stFedExGround Then
                        objFedexShip.ServiceType = ServiceTypes.stFedExGroundHomeDelivery
                    End If
                ElseIf .IsPOBox Then
                    objFedexShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            With cReturnAddress
                If .Company & String.Empty <> String.Empty Then
                    objFedexShip.ReturnContact.FirstName = .FirstName
                    objFedexShip.ReturnContact.LastName = .LastName
                    objFedexShip.ReturnContact.MiddleInitial = .MiddleInitial
                    objFedexShip.ReturnContact.Phone = .Phone
                    objFedexShip.ReturnContact.Fax = .Fax
                    objFedexShip.ReturnContact.Email = .eMail

                    objFedexShip.ReturnContact.Company = .Company
                    objFedexShip.ReturnAddress.Address1 = .Address1
                    objFedexShip.ReturnAddress.Address2 = .Address2
                    'objFedexShip.Config("ReturnAddress3=" & .Address3)

                    objFedexShip.ReturnAddress.City = .City
                    objFedexShip.ReturnAddress.ZipCode = .ZipCode
                    objFedexShip.ReturnAddress.State = .State
                    objFedexShip.ReturnAddress.CountryCode = .CountryCode
                End If
            End With

            'With cAccountContact
            '    If .Company & String.Empty <> String.Empty Then
            '        objFedexShip.ReturnContact.FirstName = .FirstName
            '        objFedexShip.ReturnContact.LastName = .LastName
            '        objFedexShip.ReturnContact.MiddleInitial = .MiddleInitial
            '        objFedexShip.ReturnContact.Phone = .Phone
            '        objFedexShip.ReturnContact.Fax = .Fax
            '        objFedexShip.ReturnContact.Email = .eMail

            '        objFedexShip.ReturnContact.Company = .Company
            '        objFedexShip.ReturnAddress.Address1 = .Address1
            '        objFedexShip.ReturnAddress.Address2 = .Address2
            '        'objFedexShip.Config("ReturnAddress3=" & .Address3)

            '        objFedexShip.ReturnAddress.City = .City
            '        objFedexShip.ReturnAddress.ZipCode = .ZipCode
            '        objFedexShip.ReturnAddress.State = .State
            '        objFedexShip.ReturnAddress.CountryCode = .CountryCode
            '    End If
            'End With

        Catch ex As Exception
            LastError = ex.Message
            objFedexShip.Dispose()
            objFedexShip = Nothing
            Return False
        End Try

        Try
            Select Case EzshipLabelImage
                Case EzshipLabelImageTypes.itEltron
                    objFedexShip.LabelImageType = nsoftware.InShip.FedexshipLabelImageTypes.fitEltron
                Case EzshipLabelImageTypes.itPDF
                    objFedexShip.LabelImageType = nsoftware.InShip.FedexshipLabelImageTypes.fitPDF
                Case EzshipLabelImageTypes.itPNG
                    objFedexShip.LabelImageType = nsoftware.InShip.FedexshipLabelImageTypes.fitPNG
                Case EzshipLabelImageTypes.itUniMark
                    objFedexShip.LabelImageType = nsoftware.InShip.FedexshipLabelImageTypes.fitUniMark
                Case EzshipLabelImageTypes.itZebra
                    objFedexShip.LabelImageType = nsoftware.InShip.FedexshipLabelImageTypes.fitZebra
                Case Else
                    objFedexShip.LabelImageType = nsoftware.InShip.FedexshipLabelImageTypes.fitEltron
            End Select

            Dim extension As String = objFedexShip.LabelImageType.ToString
            If extension.StartsWith("fit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            ' Set shipping directory to store the labels
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If
                    For Each shippingPackageDetail In PackageDetailList
                        Dim id As String = idCtr.ToString
                        idCtr += 1
                        If Val(shippingPackageDetail.Id) > 0 Then
                            id = Val(shippingPackageDetail.Id)
                        End If

                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                        shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                    Next
                End If
            End If

            ' Add packages (package weight is in Ounces - Convert to Pounds)
            Dim TotalWeight As Decimal = 0
            Dim totalInsured As Decimal = 0
            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                TotalWeight += Val(shippingPackageDetail.Weight & String.Empty)

                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = 0
                End If
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)

                ' FedEx requires a direct sign if Insured Value > 500, or if signature is required
                If Val(shippingPackageDetail.InsuredValue & String.Empty) > 500 OrElse cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objFedexShip.Packages.Add(shippingPackageDetail)
            Next

            objFedexShip.TotalWeight = Format(TotalWeight, "###0.0")
            objFedexShip.InsuredValue = Format(totalInsured, "###0.00")

            With objFedexShip.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            '  Custom Content to place on a Fedex Label
            '  When using the CustomContent, the LabelStockType must be either 4 (Stock 4x8) or 5 (Stock 4x9 Leading Doc Tab). 
            '  Also LabelFormatType must be 0 (Common2D) and LabelImageType must 2 (fitEltron), 3 (fitZebra) or 4 (fitUniMark). 

            ' NS-HF048860852E
            If cFedexCustomContent.Length > 0 Then
                objFedexShip.Config("CustomContent=" & cFedexCustomContent)

                ' Need to check the labal type
                ' 4 = 4x8, 5=4x9 - has pull tab
                If LabelStockType <> "4" And LabelStockType <> "5" Then
                    LabelStockType = "4"
                End If

                If objFedexShip.LabelImageType <> FedexshipLabelImageTypes.fitEltron _
                    AndAlso objFedexShip.LabelImageType <> FedexshipLabelImageTypes.fitZebra _
                    AndAlso objFedexShip.LabelImageType <> FedexshipLabelImageTypes.fitUniMark Then
                    objFedexShip.LabelImageType = FedexshipLabelImageTypes.fitEltron
                End If
            End If

            If LabelStockType.Length > 0 Then
                'MessageBox.Show("LabelStockType=" & LabelStockType)
                objFedexShip.Config("LabelStockType=" & LabelStockType)
            End If

            Dim specialService As Long = ShipmentSpecialServices

            'If objFedexShipIntl.ShipDate > DateTime.Now.ToString("yyyy-MM-dd") Then
            '    specialService = specialService Or &H20000000L
            'End If

            objFedexShip.ShipmentSpecialServices = specialService

            ' Notifications
            Dim notificationsIndex As Int16 = 0
            If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
                For Each sn As Notifications In ShipmentNotifications
                    sn.email = (sn.email & String.Empty).Trim
                    If sn.email.Length = 0 Then
                        Continue For
                    End If

                    Dim notify As New nsoftware.InShip.NotifyDetail
                    With notify
                        .Email = sn.email
                        .NotificationFlags = CInt(sn.NotificationFlags)
                        .Message = (sn.Message & String.Empty).ToString.Trim
                    End With

                    objFedexShip.Notify.Add(notify)

                    notificationsIndex += 1
                    If notificationsIndex = 3 Then Exit For
                Next
            End If

            objFedexShip.RuntimeLicense = inShipLicense
            objFedexShip.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objFedexShip.Packages.Count - 1
                PackageDetailList.Add(objFedexShip.Packages(ictr))
                GetPackageCosts(objFedexShip.Packages(ictr), objFedexShip)
            Next

            cRawResponse = objFedexShip.Config("RawResponse")
            Dim dsFedExRates As New DataSet

            Try
                Dim stream As IO.StringReader
                stream = New IO.StringReader(cRawResponse)
                dsFedExRates.ReadXml(stream)
            Catch ex As Exception
                dsFedExRates = Nothing
            End Try

            Dim tablesExist As Boolean = False
            tablesExist = dsFedExRates.Tables.Contains("ProcessShipmentReply") _
                        AndAlso dsFedExRates.Tables.Contains("CompletedShipmentDetail") _
                        AndAlso dsFedExRates.Tables.Contains("ShipmentRating") _
                        AndAlso dsFedExRates.Tables.Contains("ShipmentRateDetails") _
                        AndAlso dsFedExRates.Tables.Contains("TotalNetFedExCharge")

            Dim TotalNetFedExCharge As Decimal = 0
            If tablesExist Then
                For Each rowPSR As DataRow In dsFedExRates.Tables("ProcessShipmentReply").Select("")
                    Dim ProcessShipmentReply_ID As String = rowPSR.Item("ProcessShipmentReply_ID") & String.Empty
                    For Each rowCSD As DataRow In dsFedExRates.Tables("CompletedShipmentDetail").Select("ProcessShipmentReply_ID = '" & ProcessShipmentReply_ID & "'")
                        Dim CompletedShipmentDetail_ID As String = rowCSD.Item("CompletedShipmentDetail_ID") & String.Empty
                        For Each rowSR As DataRow In dsFedExRates.Tables("ShipmentRating").Select("CompletedShipmentDetail_ID = '" & CompletedShipmentDetail_ID & "' and ActualRateType = 'PAYOR_ACCOUNT_SHIPMENT'")
                            Dim ShipmentRating_ID As String = rowSR.Item("ShipmentRating_ID") & String.Empty
                            Dim ActualRateType As String = rowSR.Item("ActualRateType") & String.Empty
                            For Each rowSRD As DataRow In dsFedExRates.Tables("ShipmentRateDetails").Select("RateType = '" & ActualRateType & "' and ShipmentRating_ID = '" & ShipmentRating_ID & "'")
                                Dim ShipmentRateDetails_ID As String = rowSRD.Item("ShipmentRateDetails_ID") & String.Empty
                                For Each rowTNFC As DataRow In dsFedExRates.Tables("TotalNetFedExCharge").Select("ShipmentRateDetails_ID = '" & ShipmentRateDetails_ID & "'")
                                    TotalNetFedExCharge += Val(rowTNFC.Item("AMOUNT") & String.Empty)
                                Next
                            Next
                        Next
                    Next
                Next
            End If

            Dim totalApplied As Decimal = 0
            If TotalNetFedExCharge > 0 AndAlso ShipmentNetCharge.Count > 0 Then
                For iloop As Int16 = 1 To ShipmentNetCharge.Count
                    ShipmentNetCharge(iloop) = Math.Round((TotalNetFedExCharge / (ShipmentNetCharge.Count)), 2, MidpointRounding.AwayFromZero)
                    totalApplied += ShipmentNetCharge(iloop)
                Next
                If totalApplied < TotalNetFedExCharge Then
                    ShipmentNetCharge(1) += (TotalNetFedExCharge - totalApplied)
                End If
            End If


            If objFedexShip.Packages.Count = 1 Then
                cMasterTrackingNumber = objFedexShip.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objFedexShip.MasterTrackingNumber
            End If

            If cMasterTrackingNumber.Length = 0 Then
                LastError = "Shipper did not return a tracking number"
                Return False
            End If

            Dim smartShipTracking As String = objFedexShip.Config("SmartPostTrackingNumbers")
            For Each track As String In smartShipTracking.Split(",")
                If track.Length > 0 Then
                    FedexSmartPost.TrackingNumbers.Add(track)
                End If
            Next

            RequestFedexLabel = True

        Catch ex As nsoftware.InShip.InShipFedexshipException
            LastError = ex.Message
            RequestFedexLabel = False

        Catch exc As Exception
            LastError = exc.Message
            RequestFedexLabel = False

        Finally
            cRawRequest = objFedexShip.Config("RawRequest")
            cRawResponse = objFedexShip.Config("RawResponse")
            objFedexShip.Dispose()
            objFedexShip = Nothing
        End Try

    End Function

    ''' <summary>
    ''' Close Ground Shipment for the day
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FedexCloseGroundShipments() As Boolean
        Try

            LastError = String.Empty
            objFedexShip = New nsoftware.InShip.Fedexship
            objFedexShip.RuntimeLicense = inShipLicense
            objFedexShip.Reset()
            objFedexShip.RuntimeLicense = inShipLicense

            objFedexShip.FedExAccount.Server = cServer
            objFedexShip.FedExAccount.AccountNumber = cAccountNumber
            objFedexShip.FedExAccount.Password = cPassword
            objFedexShip.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexShip.FedExAccount.DeveloperKey = cFedexDeveloperKey

            If cServer.ToUpper.Contains("WEB-SERVICES") Then
                objFedexShip.Config("UseSOAP=true")
            Else
                objFedexShip.Config("UseSOAP=false")
            End If

            objFedexShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            With objFedexShip.CloseRequest
                .Date = FedexClose.Date
                .ReportFile = FedexClose.ReportFile
                .ReportType = FedexClose.ReportType
                .Time = FedexClose.Time
            End With

            objFedexShip.CloseGroundShipments()
            Return True

        Catch ex As Exception
            LastError = ex.Message
            Return False
        Finally
            cRawRequest = objFedexShip.Config("RawRequest")
            cRawResponse = objFedexShip.Config("RawResponse")
            objFedexShip.Dispose()
            objFedexShip = Nothing
        End Try

    End Function

    ''' <summary>
    ''' Cancel / Void Shipment
    ''' </summary>
    ''' <param name="TrackingNumber"></param>
    ''' <param name="isMultiPackage"></param>
    ''' <param name="FedexTrackingIDType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CancelFedexShipment(ByVal TrackingNumber As String, ByVal isMultiPackage As Boolean, ByVal FedexTrackingIDType As Int16) As Boolean

        Try

            LastError = String.Empty
            objFedexShip = New nsoftware.InShip.Fedexship
            objFedexShip.RuntimeLicense = inShipLicense
            objFedexShip.Reset()
            objFedexShip.RuntimeLicense = inShipLicense

            If cServiceProvider <> ServiceProviders.FederalExpress AndAlso cServiceProvider <> ServiceProviders.FederalExpressInternational Then
                LastError = "Invalid Service Type for Fedex shipment cancellation"
                Return False
            End If

            objFedexShip.FedExAccount.Server = cServer
            objFedexShip.FedExAccount.AccountNumber = cAccountNumber
            objFedexShip.FedExAccount.Password = cPassword
            objFedexShip.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexShip.FedExAccount.DeveloperKey = cFedexDeveloperKey

            If cServer.ToUpper.Contains("WEB-SERVICES") Then
                objFedexShip.Config("UseSOAP=true")
            Else
                objFedexShip.Config("UseSOAP=false")
            End If

            objFedexShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            If isMultiPackage Then
                objFedexShip.CancelShipment(TrackingNumber, FedexTrackingIDType)
            Else
                objFedexShip.CancelPackage(TrackingNumber, FedexTrackingIDType)
            End If

            Return True

        Catch ex As Exception
            LastError = ex.Message
            Return False
        Finally
            cRawRequest = objFedexShip.Config("RawRequest")
            cRawResponse = objFedexShip.Config("RawResponse")
            objFedexShip.Dispose()
            objFedexShip = Nothing
        End Try
    End Function

    Public Function GetFedexRates() As Decimal

        Try
            LastError = String.Empty
            objFedexRates = New nsoftware.InShip.Fedexrates
            objFedexRates.RuntimeLicense = inShipLicense
            objFedexRates.Reset()
            objFedexRates.RuntimeLicense = inShipLicense

            ' Set credentials
            objFedexRates.FedExAccount.Server = cServer ' "https://gatewaybeta.fedex.com:443/xml"
            objFedexRates.FedExAccount.DeveloperKey = cFedexDeveloperKey
            objFedexRates.FedExAccount.Password = cPassword
            objFedexRates.FedExAccount.AccountNumber = cAccountNumber
            objFedexRates.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexRates.RequestedService = cRequestedServiceType

            If cServer.ToUpper.Contains("WEB-SERVICES") Then
                objFedexRates.Config("UseSOAP=true")
            Else
                objFedexRates.Config("UseSOAP=false")
            End If

            objFedexRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            ' Get Sender Information
            With cSenderContact
                objFedexRates.SenderAddress.State = .State
                objFedexRates.SenderAddress.ZipCode = .ZipCode
                objFedexRates.SenderAddress.CountryCode = .CountryCode
                If objFedexRates.SenderAddress.CountryCode.Length = 0 Then
                    objFedexRates.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.SenderAddress.AddressFlags = &H2 'Residential
                End If
            End With

            With cRecipientContact
                objFedexRates.RecipientAddress.State = .State
                objFedexRates.RecipientAddress.ZipCode = .ZipCode
                objFedexRates.RecipientAddress.CountryCode = .CountryCode
                If objFedexRates.RecipientAddress.CountryCode.Length = 0 Then
                    objFedexRates.RecipientAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0
            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)

                shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                totalInsured += shippingPackageDetail.InsuredValue

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objFedexRates.Packages.Add(shippingPackageDetail)
            Next

            objFedexRates.TotalWeight = Format(Val(totalWeight), "###0.0")
            objFedexRates.InsuredValue = Format(Val(totalInsured), "###0.00")

            objFedexRates.ShipmentSpecialServices = 0
            objFedexRates.DropoffType = FedexratesDropoffTypes.dtRegularPickup
            objFedexRates.Config("WeightUnit=LB")

            If IsDate(Me.ShipDate) Then
                objFedexRates.ShipDate = CDate(Me.ShipDate).ToString("yyyy-MM-dd")
            Else
                objFedexRates.ShipDate = DateTime.Now.ToString("yyyy-MM-dd")
            End If

            objFedexRates.ShipmentSpecialServices = ShipmentSpecialServices

            objFedexRates.GetRates()

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            GetFedexRates = 0

            If objFedexRates.Config("Warning") = String.Empty OrElse objFedexRates.Services.Count > 0 Then
                For i As Integer = 0 To objFedexRates.Services.Count - 1
                    ShipmentBaseCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountBaseCharge)))
                    ShipmentDiscountCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountTotalDiscount)))
                    ShipmentListCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).ListBaseCharge)))
                    ShipmentNetCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountNetCharge)))
                    ShipmentSurCharge.Add(i, Convert.ToDecimal(Val(objFedexRates.Services(i).AccountTotalSurcharge)))
                    GetFedexRates += Convert.ToDecimal(Val(objFedexRates.Services(i).ListBaseCharge))
                Next
            End If
        Catch ex As Exception
            LastError = ex.Message
            Return 0
        Finally
            cRawRequest = objFedexRates.Config("RawRequest")
            cRawResponse = objFedexRates.Config("RawResponse")
            objFedexRates.Dispose()
            objFedexRates = Nothing
        End Try

    End Function

    Public Function GetFedExRatesList() As RateList()

        Try

            Dim requestedRateList As RateList()
            ReDim requestedRateList(1)

            LastError = String.Empty
            objFedexRates = New nsoftware.InShip.Fedexrates
            objFedexRates.RuntimeLicense = inShipLicense
            objFedexRates.Reset()
            objFedexRates.RuntimeLicense = inShipLicense

            ' Set credentials
            objFedexRates.FedExAccount.Server = cServer ' "https://gatewaybeta.fedex.com:443/xml"
            objFedexRates.FedExAccount.DeveloperKey = cFedexDeveloperKey
            objFedexRates.FedExAccount.Password = cPassword
            objFedexRates.FedExAccount.AccountNumber = cAccountNumber
            objFedexRates.FedExAccount.MeterNumber = cFedexMeterNumber
            objFedexRates.RequestedService = cRequestedServiceType

            If cServer.ToUpper.Contains("WEB-SERVICES") Then
                objFedexRates.Config("UseSOAP=true")
            Else
                objFedexRates.Config("UseSOAP=false")
            End If

            objFedexRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            ' Get Sender Information
            With cSenderContact
                objFedexRates.SenderAddress.State = .State
                objFedexRates.SenderAddress.ZipCode = .ZipCode
                objFedexRates.SenderAddress.CountryCode = .CountryCode
                If objFedexRates.SenderAddress.CountryCode.Length = 0 Then
                    objFedexRates.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.SenderAddress.AddressFlags = &H2 'Residential
                End If
            End With

            With cRecipientContact
                objFedexRates.RecipientAddress.State = .State
                objFedexRates.RecipientAddress.ZipCode = .ZipCode
                objFedexRates.RecipientAddress.CountryCode = .CountryCode
                If objFedexRates.RecipientAddress.CountryCode.Length = 0 Then
                    objFedexRates.RecipientAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objFedexRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objFedexRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0
            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                ' Add packages (package weight is in Ounces - Convert to Pounds)
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)

                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = 0
                End If

                shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                totalInsured += shippingPackageDetail.InsuredValue

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objFedexRates.Packages.Add(shippingPackageDetail)
            Next

            objFedexRates.TotalWeight = Format(Val(totalWeight), "###0.0")
            objFedexRates.InsuredValue = Format(Val(totalInsured), "###0.00")

            objFedexRates.ShipmentSpecialServices = 0
            objFedexRates.DropoffType = FedexratesDropoffTypes.dtRegularPickup
            objFedexRates.Config("WeightUnit=LB")

            If IsDate(Me.ShipDate) Then
                objFedexRates.ShipDate = CDate(Me.ShipDate).ToString("yyyy-MM-dd")
            Else
                objFedexRates.ShipDate = DateTime.Now.ToString("yyyy-MM-dd")
            End If

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            objFedexRates.ShipmentSpecialServices = ShipmentSpecialServices

            objFedexRates.GetRates()

            cRawResponse = objFedexRates.Config("RawResponse")
            Dim dsFedExRates As New DataSet

            Try
                Dim stream As IO.StringReader
                stream = New IO.StringReader(cRawResponse)
                dsFedExRates.ReadXml(stream)
            Catch ex As Exception
                dsFedExRates = Nothing
            End Try

            Dim tablesExist As Boolean = False
            tablesExist = dsFedExRates.Tables.Contains("RateReply") _
                        AndAlso dsFedExRates.Tables.Contains("RateReplyDetails") _
                        AndAlso dsFedExRates.Tables.Contains("RatedShipmentDetails") _
                        AndAlso dsFedExRates.Tables.Contains("ShipmentRateDetail") _
                        AndAlso dsFedExRates.Tables.Contains("EffectiveNetDiscount")


            ReDim requestedRateList(objFedexRates.Services.Count)
            If objFedexRates.Config("Warning") = String.Empty OrElse objFedexRates.Services.Count > 0 Then

                For iLoop As Integer = 0 To objFedexRates.Services.Count - 1
                    With requestedRateList(iLoop)
                        .ServiceType = objFedexRates.Services(iLoop).ServiceType
                        .ServiceTypeDescription = StrConv(objFedexRates.Services(iLoop).ServiceTypeDescription.Replace("_", " "), VbStrConv.ProperCase)
                        .AccountNetCharge = Val(objFedexRates.Services(iLoop).AccountNetCharge & String.Empty)
                        .DeliveryTime = objFedexRates.Services(iLoop).DeliveryTime
                        .ListNetCharge = Val(objFedexRates.Services(iLoop).ListNetCharge & String.Empty)
                        Dim ServiceType As String = objFedexRates.Services(iLoop).ServiceTypeDescription & String.Empty

                        Dim EffectiveNetDiscount As Decimal = 0

                        If tablesExist Then
                            If dsFedExRates.Tables("RateReply").Rows.Count = 1 Then
                                Dim RateReply_Id As String = dsFedExRates.Tables("RateReply").Rows(0).Item("RateReply_Id") & String.Empty

                                For Each rowRRD As DataRow In dsFedExRates.Tables("RateReplyDetails").Select("RateReply_Id = '" & RateReply_Id & "' and ServiceType = '" & ServiceType & "'", "RateReplyDetails_Id")
                                    Dim RateReplyDetails_Id As String = rowRRD.Item("RateReplyDetails_Id") & String.Empty
                                    For Each rowRSD As DataRow In dsFedExRates.Tables("RatedShipmentDetails").Select("RateReplyDetails_Id = '" & RateReplyDetails_Id & "'", "")
                                        Dim RatedShipmentDetails_Id As String = rowRSD.Item("RatedShipmentDetails_Id") & String.Empty
                                        For Each rowSRD As DataRow In dsFedExRates.Tables("ShipmentRateDetail").Select("RateType = 'PAYOR_ACCOUNT_PACKAGE' and RatedShipmentDetails_Id = '" & RatedShipmentDetails_Id & "'", "")
                                            Dim ShipmentRateDetail_Id As String = rowSRD.Item("ShipmentRateDetail_Id") & String.Empty
                                            ' TotalNetFreight
                                            For Each rowTNF As DataRow In dsFedExRates.Tables("TotalNetFedExCharge").Select("ShipmentRateDetail_Id = '" & ShipmentRateDetail_Id & "'", "")
                                                EffectiveNetDiscount += Val(rowTNF.Item("Amount") & String.Empty)
                                            Next
                                        Next
                                    Next
                                Next
                            End If
                        End If

                        If EffectiveNetDiscount > 0 Then
                            .AccountNetCharge = EffectiveNetDiscount
                        End If

                        .TransitTime = "0"

                        If tablesExist _
                            AndAlso dsFedExRates.Tables.Contains("COMMITDETAILS") _
                            AndAlso dsFedExRates.Tables("COMMITDETAILS").Select("ServiceType = '" & ServiceType & "'").Length = 1 Then

                            Dim row As DataRow = dsFedExRates.Tables("COMMITDETAILS").Select("ServiceType = '" & ServiceType & "'")(0)
                            Dim CommitTimeStamp As String = row.Item("CommitTimeStamp") & String.Empty
                            .DeliveryTime = CommitTimeStamp
                            If IsDate(CommitTimeStamp) Then
                                CommitTimeStamp = CDate(CommitTimeStamp).ToShortDateString
                                Dim transittime As Int16 = Math.Abs(DateDiff(DateInterval.Day, CDate(CommitTimeStamp), System.DateTime.Now)) - 1
                                If transittime > 0 Then
                                    .TransitTime = transittime.ToString
                                End If
                            End If
                        End If

                        If Val(.TransitTime <= 0) Then
                            Select Case objFedexRates.Services(iLoop).TransitTime
                                Case "ONE_DAY"
                                    .TransitTime = "1"
                                Case "TWO_DAYS"
                                    .TransitTime = "2"
                                Case "THREE_DAYS"
                                    .TransitTime = "3"
                                Case "FOUR_DAYS"
                                    .TransitTime = "4"
                                Case "FIVE_DAYS"
                                    .TransitTime = "5"
                                Case "SIX_DAYS"
                                    .TransitTime = "6"
                                Case "SEVEN_DAYS"
                                    .TransitTime = "7"
                                Case "EIGHT_DAYS"
                                    .TransitTime = "8"
                                Case "NINE_DAYS"
                                    .TransitTime = "9"
                                Case "TEN_DAYS"
                                    .TransitTime = "10"
                                Case Else
                                    .TransitTime = objFedexRates.Services(iLoop).TransitTime
                            End Select
                        End If

                        If .AccountNetCharge = 0 Then
                            .AccountNetCharge = .ListNetCharge
                        End If
                    End With
                Next
            End If

            Return requestedRateList

        Catch ex As Exception
            LastError = ex.Message
            Return Nothing
        Finally
            cRawRequest = objFedexRates.Config("RawRequest")
            cRawResponse = objFedexRates.Config("RawResponse")
            objFedexRates.Dispose()
            objFedexRates = Nothing
        End Try

    End Function

    Public Function FedExTrack(ByVal TrackingNumber As String) As String

        Dim response As String = String.Empty

        Try
            LastError = String.Empty
            clsTrackingData = New TrackingData

            objFedexTrack = New nsoftware.InShip.Fedextrack
            objFedexTrack.RuntimeLicense = inShipLicense
            objFedexTrack.Reset()
            objFedexTrack.RuntimeLicense = inShipLicense

            ' Set credentials
            objFedexTrack.FedExAccount.Server = cServer ' "https://gatewaybeta.fedex.com:443/xml"
            objFedexTrack.FedExAccount.DeveloperKey = cFedexDeveloperKey
            objFedexTrack.FedExAccount.Password = cPassword
            objFedexTrack.FedExAccount.AccountNumber = cAccountNumber
            objFedexTrack.FedExAccount.MeterNumber = cFedexMeterNumber

            objFedexTrack.TrackShipment(TrackingNumber)

            If objFedexTrack.TrackEvents.Count > 0 Then
                Dim index As Int16 = 0 'objFedexTrack.TrackEvents.Count - 1
                response = String.Empty
                response &= "Status: " & objFedexTrack.TrackEvents(index).Status & Environment.NewLine
                response &= "Date: " & objFedexTrack.TrackEvents(index).Date & Environment.NewLine
                response &= "Time: " & objFedexTrack.TrackEvents(index).Time & Environment.NewLine
                response &= "City: " & objFedexTrack.TrackEvents(index).City & Environment.NewLine
                response &= "State: " & objFedexTrack.TrackEvents(index).State & Environment.NewLine
                response &= "CountryCode: " & objFedexTrack.TrackEvents(index).CountryCode & Environment.NewLine
                response &= "Location: " & objFedexTrack.TrackEvents(index).Location

                clsTrackingData.Status = objFedexTrack.TrackEvents(index).Status & String.Empty
                clsTrackingData.Date = objFedexTrack.TrackEvents(index).Date & String.Empty
                clsTrackingData.Time = objFedexTrack.TrackEvents(index).Time & String.Empty
                clsTrackingData.City = objFedexTrack.TrackEvents(index).City & String.Empty
                clsTrackingData.State = objFedexTrack.TrackEvents(index).State & String.Empty
                clsTrackingData.CountryCode = objFedexTrack.TrackEvents(index).CountryCode & String.Empty
                clsTrackingData.Location = objFedexTrack.TrackEvents(index).Location & String.Empty
                clsTrackingData.Address1 = objFedexTrack.TrackEvents(index).Address1 & String.Empty
                clsTrackingData.Address2 = objFedexTrack.TrackEvents(index).Address2 & String.Empty
                clsTrackingData.ZipCode = objFedexTrack.TrackEvents(index).ZipCode & String.Empty
            End If

        Catch ex As nsoftware.InShip.InShipUpsshipException
            LastError = ex.Message
            response = ex.Message

        Catch exc As Exception
            LastError = exc.Message
            response = exc.Message

        Finally
            cRawRequest = objFedexTrack.Config("RawRequest")
            cRawResponse = objFedexTrack.Config("RawResponse")
            objFedexTrack.Dispose()
            objFedexTrack = Nothing
        End Try

        Return response

    End Function


#End Region

#Region "UPS"

    Public Function UPSTrack(ByVal TrackingNumber As String) As String

        Dim response As String = String.Empty

        Try

            LastError = String.Empty
            clsTrackingData = New TrackingData

            objUpsTrack = New nsoftware.InShip.Upstrack
            objUpsTrack.RuntimeLicense = inShipLicense
            objUpsTrack.Reset()
            objUpsTrack.RuntimeLicense = inShipLicense

            Dim useSoap As Boolean = cServer.ToUpper.Contains("WEBSERVICES")
            If useSoap Then
                objUpsTrack.Config("UseSOAP=true")
                If Not cServer.EndsWith("/") Then
                    objUpsTrack.UPSAccount.Server = cServer & "/Track"
                Else
                    objUpsTrack.UPSAccount.Server = cServer & "Track"
                End If
            Else
                objUpsTrack.Config("UseSOAP=false")
                If Not cServer.EndsWith("/") Then
                    objUpsTrack.UPSAccount.Server = cServer & "/Track"
                Else
                    objUpsTrack.UPSAccount.Server = cServer & "Track"
                End If
            End If

            objUpsTrack.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            objUpsTrack.UPSAccount.AccessKey = cUPSAccessKey
            objUpsTrack.UPSAccount.AccountNumber = cAccountNumber
            objUpsTrack.UPSAccount.Password = cPassword
            objUpsTrack.UPSAccount.UserId = cUserId

            objUpsTrack.TrackShipment(TrackingNumber)

            If objUpsTrack.TrackEvents.Count > 0 Then
                Dim index As Int16 = 0 'objUpsTrack.TrackEvents.Count - 1
                response = String.Empty
                response &= "Status: " & objUpsTrack.TrackEvents(index).Status & Environment.NewLine
                response &= "Date: " & objUpsTrack.TrackEvents(index).Date & Environment.NewLine
                response &= "Time: " & objUpsTrack.TrackEvents(index).Time & Environment.NewLine
                response &= "City: " & objUpsTrack.TrackEvents(index).City & Environment.NewLine
                response &= "State: " & objUpsTrack.TrackEvents(index).State & Environment.NewLine
                response &= "CountryCode: " & objUpsTrack.TrackEvents(index).CountryCode & Environment.NewLine
                response &= "Location: " & objUpsTrack.TrackEvents(index).Location

                clsTrackingData.Status = objUpsTrack.TrackEvents(index).Status & String.Empty
                clsTrackingData.Date = objUpsTrack.TrackEvents(index).Date & String.Empty
                clsTrackingData.Time = objUpsTrack.TrackEvents(index).Time & String.Empty
                clsTrackingData.City = objUpsTrack.TrackEvents(index).City & String.Empty
                clsTrackingData.State = objUpsTrack.TrackEvents(index).State & String.Empty
                clsTrackingData.CountryCode = objUpsTrack.TrackEvents(index).CountryCode & String.Empty
                clsTrackingData.Location = objUpsTrack.TrackEvents(index).Location & String.Empty
                clsTrackingData.Address1 = objUpsTrack.TrackEvents(index).Address1 & String.Empty
                clsTrackingData.Address2 = objUpsTrack.TrackEvents(index).Address2 & String.Empty
                clsTrackingData.ZipCode = objUpsTrack.TrackEvents(index).ZipCode & String.Empty

            End If

        Catch ex As nsoftware.InShip.InShipUpsshipException
            LastError = ex.Message
            response = ex.Message

        Catch exc As Exception
            LastError = exc.Message
            response = exc.Message

        Finally
            cRawRequest = objUpsTrack.Config("RawRequest")
            cRawResponse = objUpsTrack.Config("RawResponse")
            objUpsTrack.Dispose()
            objUpsTrack = Nothing
        End Try

        Return response

    End Function

    Public Function GetUPSRates() As Decimal

        Try

            LastError = String.Empty
            objUpsRates = New nsoftware.InShip.Upsrates
            objUpsRates.RuntimeLicense = inShipLicense
            objUpsRates.Reset()
            objUpsRates.RuntimeLicense = inShipLicense
            GetUPSRates = 0

            ShipmentBaseCharge.Clear()
            ShipmentDiscountCharge.Clear()
            ShipmentListCharge.Clear()
            ShipmentNetCharge.Clear()
            ShipmentSurCharge.Clear()

            If cServer.ToUpper.Contains("WEBSERVICES") Then
                objUpsRates.Config("UseSOAP=true")
            Else
                objUpsRates.Config("UseSOAP=false")
            End If

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            objUpsRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12


            If Not cServer.EndsWith("/") Then
                objUpsRates.UPSAccount.Server = cServer & "/Rate"
            Else
                objUpsRates.UPSAccount.Server = cServer & "Rate"
            End If

            objUpsRates.UPSAccount.AccessKey = cUPSAccessKey
            objUpsRates.UPSAccount.AccountNumber = cAccountNumber
            objUpsRates.UPSAccount.Password = cPassword
            objUpsRates.UPSAccount.UserId = cUserId
            objUpsRates.RequestedService = cRequestedServiceType

            objUpsRates.PickupType = UpsratesPickupTypes.ptDailyPickup
            'objUpsRates.CustomerType = UpsratesCustomerTypes.ccRetail

            ' Insured Value is Positive, Declared Value is Negative
            Dim iPackage As Int16 = 0
            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0

            For Each shippingPackageDetail In PackageDetailList

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objUpsRates.Packages.Add(shippingPackageDetail)

                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                    objUpsRates.Config("PackageDeclaredValueType[" & iPackage & "]=0")
                End If

                ' Format weight 
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured = Val(shippingPackageDetail.InsuredValue)

                iPackage += 1
            Next

            objUpsRates.TotalWeight = Format(Val(totalWeight), "###0.0")

            With cSenderContact
                objUpsRates.SenderAddress.Address1 = .Address1
                objUpsRates.SenderAddress.Address2 = .Address2
                'objUpsRates.Config("SenderAddress3=" & .Address3)

                objUpsRates.SenderAddress.City = .City
                objUpsRates.SenderAddress.ZipCode = .ZipCode
                objUpsRates.SenderAddress.State = .State
                objUpsRates.SenderAddress.CountryCode = .CountryCode
                If objUpsRates.SenderAddress.CountryCode.ToUpper = "USA" Then
                    objUpsRates.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objUpsRates.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If
            End With

            With cRecipientContact
                objUpsRates.RecipientAddress.Address1 = .Address1
                objUpsRates.RecipientAddress.Address2 = .Address2
                'objUpsRates.Config("RecipientAddress3=" & .Address3)

                objUpsRates.RecipientAddress.City = .City
                objUpsRates.RecipientAddress.ZipCode = .ZipCode
                objUpsRates.RecipientAddress.State = .State
                objUpsRates.RecipientAddress.CountryCode = .CountryCode
                If objUpsRates.RecipientAddress.CountryCode.ToUpper = "USA" Then
                    objUpsRates.RecipientAddress.CountryCode = "US"
                End If
            End With

            If IsDate(Me.ShipDate) Then
                objUpsRates.ShipDate = CDate(Me.ShipDate).ToString("yyyyMMdd")
            Else
                objUpsRates.ShipDate = DateTime.Now.ToString("yyyyMMdd")
            End If

            objUpsRates.ShipmentSpecialServices = ShipmentSpecialServices

            objUpsRates.GetRates()

            For i As Integer = 0 To objUpsRates.Services.Count - 1
                ShipmentBaseCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountBaseCharge)))
                ShipmentDiscountCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountTotalDiscount)))
                ShipmentListCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).ListBaseCharge)))
                ShipmentNetCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountNetCharge)))
                ShipmentSurCharge.Add(i, Convert.ToDecimal(Val(objUpsRates.Services(i).AccountTotalSurcharge)))
                GetUPSRates += Convert.ToDecimal(Val(objUpsRates.Services(i).ListBaseCharge))
            Next

        Catch ex As Exception
            LastError = ex.Message
            objUpsRates.Dispose()
            objUpsRates = Nothing
        End Try

    End Function

    ''' <summary>
    ''' Request Shipping label. Not used for UPS Intenational
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestUPSLabel() As Boolean

        Try
            LastError = String.Empty
            objUpsShip = New nsoftware.InShip.Upsship
            objUpsShip.RuntimeLicense = inShipLicense
            objUpsShip.Reset()
            objUpsShip.RuntimeLicense = inShipLicense

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            Dim useSoap As Boolean = cServer.ToUpper.Contains("WEBSERVICES")
            If useSoap Then
                objUpsShip.Config("UseSOAP=true")
                If Not cServer.EndsWith("/") Then
                    objUpsShip.UPSAccount.Server = cServer & "/Ship"
                Else
                    objUpsShip.UPSAccount.Server = cServer & "Ship"
                End If
            Else
                objUpsShip.Config("UseSOAP=false")
                If Not cServer.EndsWith("/") Then
                    objUpsShip.UPSAccount.Server = cServer & "/ShipConfirm"
                Else
                    objUpsShip.UPSAccount.Server = cServer & "ShipConfirm"
                End If
            End If

            objUpsShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            objUpsShip.UPSAccount.AccessKey = cUPSAccessKey
            objUpsShip.UPSAccount.AccountNumber = cAccountNumber
            objUpsShip.UPSAccount.Password = cPassword
            objUpsShip.UPSAccount.UserId = cUserId
            objUpsShip.ServiceType = cRequestedServiceType

            objUpsShip.ShipDate = ShipDate.ToString("yyyyMMdd")

            With cSenderContact
                objUpsShip.SenderContact.FirstName = .FirstName
                objUpsShip.SenderContact.LastName = .LastName
                objUpsShip.SenderContact.MiddleInitial = .MiddleInitial
                objUpsShip.SenderContact.Phone = .Phone
                objUpsShip.SenderContact.Fax = .Fax
                objUpsShip.SenderContact.Email = .eMail

                objUpsShip.SenderContact.Company = .Company
                objUpsShip.SenderAddress.Address1 = .Address1
                objUpsShip.SenderAddress.Address2 = .Address2
                objUpsShip.Config("SenderAddress3=" & .Address3)

                objUpsShip.SenderAddress.City = .City
                objUpsShip.SenderAddress.ZipCode = .ZipCode
                objUpsShip.SenderAddress.State = .State
                objUpsShip.SenderAddress.CountryCode = .CountryCode

                objUpsShip.SenderAddress.CountryCode = objUpsShip.SenderAddress.CountryCode.ToUpper
                If objUpsShip.SenderAddress.CountryCode = "" OrElse objUpsShip.SenderAddress.CountryCode = "USA" Then
                    objUpsShip.SenderAddress.CountryCode = "US"
                End If

                If objUpsShip.SenderAddress.CountryCode = "CAN" Then
                    objUpsShip.SenderAddress.CountryCode = "CA"
                End If

                If .IsResidental Then
                    objUpsShip.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShip.SenderAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            With cAccountContact
                If .Company & String.Empty <> String.Empty Then
                    objUpsShip.AccountContact.FirstName = .FirstName
                    objUpsShip.AccountContact.LastName = .LastName
                    objUpsShip.AccountContact.MiddleInitial = .MiddleInitial
                    objUpsShip.AccountContact.Phone = .Phone
                    objUpsShip.AccountContact.Fax = .Fax
                    objUpsShip.AccountContact.Email = .eMail

                    objUpsShip.AccountContact.Company = .Company
                    objUpsShip.AccountAddress.Address1 = .Address1
                    objUpsShip.AccountAddress.Address2 = .Address2
                    objUpsShip.Config("AccountAddress3=" & .Address3)

                    objUpsShip.AccountAddress.City = .City
                    objUpsShip.AccountAddress.ZipCode = .ZipCode
                    objUpsShip.AccountAddress.State = .State
                    objUpsShip.AccountAddress.CountryCode = .CountryCode

                    objUpsShip.AccountAddress.CountryCode = objUpsShip.AccountAddress.CountryCode.ToUpper
                    If objUpsShip.AccountAddress.CountryCode = "" OrElse objUpsShip.AccountAddress.CountryCode = "USA" Then
                        objUpsShip.AccountAddress.CountryCode = "US"
                    End If

                    If objUpsShip.AccountAddress.CountryCode = "CAN" Then
                        objUpsShip.AccountAddress.CountryCode = "CA"
                    End If

                    If .IsResidental Then
                        objUpsShip.AccountAddress.AddressFlags = &H2 'Residential
                    ElseIf .IsPOBox Then
                        objUpsShip.AccountAddress.AddressFlags = &H1 'PO Box
                    End If
                End If
            End With


            With cRecipientContact
                objUpsShip.RecipientContact.FirstName = .FirstName
                objUpsShip.RecipientContact.LastName = .LastName
                objUpsShip.RecipientContact.MiddleInitial = .MiddleInitial
                objUpsShip.RecipientContact.Phone = .Phone
                objUpsShip.RecipientContact.Fax = .Fax
                objUpsShip.RecipientContact.Email = .eMail

                objUpsShip.RecipientContact.Company = .Company
                objUpsShip.RecipientAddress.Address1 = .Address1
                objUpsShip.RecipientAddress.Address2 = .Address2
                objUpsShip.Config("RecipientAddress3=" & .Address3)

                objUpsShip.RecipientAddress.City = .City
                objUpsShip.RecipientAddress.ZipCode = .ZipCode
                objUpsShip.RecipientAddress.State = .State
                objUpsShip.RecipientAddress.CountryCode = .CountryCode

                objUpsShip.RecipientAddress.CountryCode = objUpsShip.RecipientAddress.CountryCode.ToUpper
                If objUpsShip.RecipientAddress.CountryCode = "" OrElse objUpsShip.RecipientAddress.CountryCode = "USA" Then
                    objUpsShip.RecipientAddress.CountryCode = "US"
                End If

                If objUpsShip.RecipientAddress.CountryCode = "CAN" Then
                    objUpsShip.RecipientAddress.CountryCode = "CA"
                End If

                If .IsResidental Then
                    objUpsShip.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

        Catch ex As Exception
            LastError = ex.Message
            objUpsShip.Dispose()
            objUpsShip = Nothing
            Return False
        End Try

        Try
            Select Case EzshipLabelImage
                Case EzshipLabelImageTypes.itEPL
                    objUpsShip.LabelImageType = UpsshipLabelImageTypes.uitEPL
                Case EzshipLabelImageTypes.itGIF
                    objUpsShip.LabelImageType = UpsshipLabelImageTypes.uitGIF
                Case EzshipLabelImageTypes.itSPL
                    objUpsShip.LabelImageType = UpsshipLabelImageTypes.uitSPL
                Case EzshipLabelImageTypes.itZPL
                    objUpsShip.LabelImageType = UpsshipLabelImageTypes.uitZPL
                Case Else
                    objUpsShip.LabelImageType = UpsshipLabelImageTypes.uitEPL
            End Select

            Dim extension As String = EzshipLabelImage.ToString
            If extension.StartsWith("uit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            ' Set shipping directory to store the labels
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If
                    For Each shippingPackageDetail In PackageDetailList
                        Dim id As String = idCtr.ToString
                        idCtr += 1
                        If Val(shippingPackageDetail.Id) > 0 Then
                            id = Val(shippingPackageDetail.Id)
                        End If

                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                        shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                        shippingPackageDetail.ReturnReceiptFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_RTN" & extension
                    Next
                End If
            End If

            ' Add packages 
            Dim TotalWeight As Decimal = 0
            Dim totalInsured As Decimal = 0
            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                TotalWeight += shippingPackageDetail.Weight
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objUpsShip.Packages.Add(shippingPackageDetail)

                ' Force Sure Post to be greater than 1 pound
                If objUpsShip.ServiceType = ServiceTypes.stUPSSurePost1LBOrGreater AndAlso shippingPackageDetail.Weight < 1 Then
                    shippingPackageDetail.Weight = 1.1
                End If

                If objUpsShip.ServiceType = ServiceTypes.stUPSSurePost1LBOrGreater AndAlso shippingPackageDetail.Weight > 9 Then
                    LastError = "'Sure Post 1 LB or Greater' must be between 1 and 9 pounds"
                    Return False
                End If
            Next

            With objUpsShip.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
                .Name = PayorContact.Company
                .Address1 = PayorContact.Address1
                .Address2 = PayorContact.Address2
                .City = PayorContact.City
                .State = PayorContact.State
            End With

            objUpsShip.RuntimeLicense = inShipLicense

            ' Notifications
            Dim notificationsIndex As Int16 = 0
            If ShipmentNotifications.Count > 0 AndAlso Not ASCMAIN1.Running_in_VS Then
                For Each sn As Notifications In ShipmentNotifications
                    sn.email = (sn.email & String.Empty).Trim
                    If sn.email.Length = 0 Then
                        Continue For
                    End If

                    Dim notify As New nsoftware.InShip.NotifyDetail
                    With notify
                        .Email = sn.email
                        .NotificationFlags = CInt(sn.NotificationFlags)
                        .Message = (sn.Message & String.Empty).ToString.Trim
                    End With

                    objUpsShip.Notify.Add(notify)

                    notificationsIndex += 1
                    If notificationsIndex = 3 Then Exit For
                Next
            End If

            ' UPS Ground Freight
            If CommodityDetailList.Count > 0 Then
                Dim commodityDetail As New nsoftware.InShip.CommodityDetail
                commodityDetail = CommodityDetailList(0)

                objUpsShip.Config("ReturnFreightPrices=True")
                objUpsShip.Config("FRSCommodityCount=" & commodityDetail.NumberOfPieces)
                objUpsShip.Config("FRSPaymentType=" & UPSGroundFreightPayor.PayorType)
                objUpsShip.Config("FRSPaymentDescription=" & commodityDetail.Description)
                objUpsShip.Config("FRSPaymentAccountNumber=" & UPSGroundFreightPayor.AccountNumber)
                objUpsShip.Config("FRSPaymentPostalCode=" & UPSGroundFreightPayor.ZipCode)
                objUpsShip.Config("FRSPaymentCountryCode=" & UPSGroundFreightPayor.CountryCode)

                For iloop As Int16 = 0 To commodityDetail.NumberOfPieces - 1
                    objUpsShip.Config("FRSCommodityFreightClass[" & iloop & "]=" & commodityDetail.FreightClass)
                    objUpsShip.Config("FRSCommodityFreightNMFC[" & iloop & "]=" & commodityDetail.FreightNMFC)
                Next
                objUpsShip.ServiceType = ServiceTypes.stUPSGround

                objUpsShip.UPSAccount.Server = objUpsShip.UPSAccount.Server.Replace("ups.app/xml", "webservices")
                objUpsShip.UPSAccount.Server = objUpsShip.UPSAccount.Server.Replace("/ShipConfirm", "/Ship")
                objUpsShip.Config("UseSOAP=true")
            End If


            'USPSEndorsement:   The USPS endorsement type for Mail Innovations and SurePost shipments.
            'This contains the USPS endorsement type and is required when using a Mail Innovations or SurePost ServiceType. The Valid values are as follows:

            ' Value Meaning
            '0	No Service Selected
            '1	Return Service Selected
            '2	Forwarding Service Requested
            '3	Address Service Requested
            '4	Change Service Requested

            Select Case objUpsShip.ServiceType
                Case ServiceTypes.stUPSEconomyMailInnovations, ServiceTypes.stUPSExpeditedMailInnovations, ServiceTypes.stUPSPriorityMailInnovations,
                    ServiceTypes.stUPSSurePost1LBOrGreater, ServiceTypes.stUPSSurePostBPM, ServiceTypes.stUPSSurePostLessThan1LB, ServiceTypes.stUPSSurePostMedia
                    objUpsShip.Config("USPSEndorsement=0")
            End Select

            objUpsShip.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            ' For multi UPS package shipments the total cost exists in all packages
            ' so spread the costs
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objUpsShip.Packages.Count - 1
                PackageDetailList.Add(objUpsShip.Packages(ictr))
                GetPackageCosts(objUpsShip.Packages(ictr), objUpsShip)
                Dim key As Integer = Val(objUpsShip.Packages(ictr).Id)
                ShipmentBaseCharge(key) = Math.Round(ShipmentBaseCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentDiscountCharge(key) = Math.Round(ShipmentDiscountCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentSurCharge(key) = Math.Round(ShipmentSurCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentNetCharge(key) = Math.Round(ShipmentNetCharge(key) / objUpsShip.Packages.Count, 2)
                ShipmentListCharge(key) = Math.Round(ShipmentListCharge(key) / objUpsShip.Packages.Count, 2)
            Next

            If objUpsShip.Packages.Count = 1 Then
                cMasterTrackingNumber = objUpsShip.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objUpsShip.MasterTrackingNumber
            End If

            If cMasterTrackingNumber.Length = 0 Then
                LastError = "Shipper did not return a tracking number"
                Return False
            End If

            Return True

        Catch ex As nsoftware.InShip.InShipUpsshipException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objUpsShip.Config("RawRequest")
            cRawResponse = objUpsShip.Config("RawResponse")
            objUpsShip.Dispose()
            objUpsShip = Nothing
        End Try

        Return True

    End Function

    ''' <summary>
    ''' Cancel / Void UPS Shipment / Package
    ''' </summary>
    ''' <param name="TrackingNumber"></param>
    ''' <param name="isMultiPackage"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CancelUpsShipment(ByVal TrackingNumber As String, ByVal isMultiPackage As Boolean)

        Try

            LastError = String.Empty
            objUpsShip = New nsoftware.InShip.Upsship
            objUpsShip.RuntimeLicense = inShipLicense
            objUpsShip.Reset()
            objUpsShip.RuntimeLicense = inShipLicense

            If cServiceProvider <> ServiceProviders.UPS Then
                LastError = "Invalid Service Type for UPS shipment cancellation"
                Return False
            End If

            'wwwcie is the test environment
            If cServer.Contains("wwwcie.ups.com") Then
                objUpsShip.UPSAccount.Server = "https://wwwcie.ups.com/ups.app/xml/Void"
            Else
                If Not cServer.EndsWith("/") Then
                    objUpsShip.UPSAccount.Server = cServer & "/Void"
                Else
                    objUpsShip.UPSAccount.Server = cServer & "Void"
                End If
            End If

            If cServer.ToUpper.Contains("WEBSERVICES") Then
                objUpsShip.Config("UseSOAP=true")
            Else
                objUpsShip.Config("UseSOAP=false")
            End If

            objUpsShip.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            objUpsShip.UPSAccount.AccessKey = cUPSAccessKey
            objUpsShip.UPSAccount.AccountNumber = cAccountNumber
            objUpsShip.UPSAccount.Password = cPassword
            objUpsShip.UPSAccount.UserId = cUserId

            If isMultiPackage Then
                objUpsShip.CancelShipment(TrackingNumber)
            Else
                objUpsShip.CancelPackage(TrackingNumber, TrackingNumber)
            End If
            Return True
        Catch ex As Exception
            LastError = ex.Message
            objUpsShip.Dispose()
            objUpsShip = Nothing
            Return False
        End Try

    End Function

    Public Function GetUPSRatesList() As RateList()

        Try

            Dim requestedRateList As RateList()
            ReDim requestedRateList(1)
            LastError = String.Empty

            objUpsRates = New nsoftware.InShip.Upsrates
            objUpsRates.RuntimeLicense = inShipLicense
            objUpsRates.Reset()
            objUpsRates.RuntimeLicense = inShipLicense

            If cServer.ToUpper.Contains("WEBSERVICES") Then
                objUpsRates.Config("UseSOAP=true")
            Else
                objUpsRates.Config("UseSOAP=false")
            End If

            objUpsRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            If Not cServer.EndsWith("/") Then
                objUpsRates.UPSAccount.Server = cServer & "/Rate"
            Else
                objUpsRates.UPSAccount.Server = cServer & "Rate"
            End If

            objUpsRates.UPSAccount.AccessKey = cUPSAccessKey
            objUpsRates.UPSAccount.UserId = cUserId
            objUpsRates.UPSAccount.AccountNumber = cAccountNumber
            objUpsRates.UPSAccount.Password = cPassword

            objUpsRates.RequestedService = cRequestedServiceType
            objUpsRates.PickupType = UPSPickupType
            objUpsRates.CustomerType = cCustomerType

            ' Insured Value is Positive, Declared Value is Negative
            Dim iPackage As Int16 = 0
            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0

            For Each shippingPackageDetail In PackageDetailList

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objUpsRates.Packages.Add(shippingPackageDetail)

                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                    objUpsRates.Config("PackageDeclaredValueType[" & iPackage & "]=0")
                End If

                ' Format weight 
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured = Val(shippingPackageDetail.InsuredValue)

                iPackage += 1
            Next

            objUpsRates.TotalWeight = Format(Val(totalWeight), "###0.0")

            With cSenderContact
                objUpsRates.SenderAddress.ZipCode = .ZipCode
                objUpsRates.SenderAddress.State = .State
                objUpsRates.SenderAddress.CountryCode = .CountryCode
            End With

            With cRecipientContact
                objUpsRates.RecipientAddress.ZipCode = .ZipCode

                If .CountryCode = "US" And .State = "PR" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .State
                ElseIf .CountryCode = "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                ElseIf .CountryCode <> "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                    objUpsRates.RecipientAddress.Address1 = .Address1
                    objUpsRates.RecipientAddress.Address2 = .Address2
                    objUpsRates.RecipientAddress.City = .City
                End If

                If .IsResidental Then
                    objUpsRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            If IsDate(Me.ShipDate) Then
                objUpsRates.ShipDate = CDate(Me.ShipDate).ToString("yyyyMMdd")
            Else
                objUpsRates.ShipDate = DateTime.Now.ToString("yyyyMMdd")
            End If

            objUpsRates.ShipmentSpecialServices = ShipmentSpecialServices

            Dim isGroundFreight As Boolean = False

            If CommodityDetailList.Count > 0 Then
                Dim commodityDetail As New nsoftware.InShip.CommodityDetail
                commodityDetail = CommodityDetailList(0)

                objUpsRates.Config("ReturnFreightPrices=True")
                objUpsRates.Config("FRSCommodityCount=" & commodityDetail.NumberOfPieces)
                objUpsRates.Config("FRSPaymentType=" & UPSGroundFreightPayor.PayorType)
                objUpsRates.Config("FRSPaymentDescription=" & commodityDetail.Description)
                objUpsRates.Config("FRSPaymentAccountNumber=" & UPSGroundFreightPayor.AccountNumber)
                objUpsRates.Config("FRSPaymentPostalCode=" & UPSGroundFreightPayor.ZipCode)
                objUpsRates.Config("FRSPaymentCountryCode=" & UPSGroundFreightPayor.CountryCode)

                For iloop As Int16 = 0 To commodityDetail.NumberOfPieces - 1
                    objUpsRates.Config("FRSCommodityFreightClass[" & iloop & "]=" & commodityDetail.FreightClass)
                    objUpsRates.Config("FRSCommodityFreightNMFC[" & iloop & "]=" & commodityDetail.FreightNMFC)
                Next
                objUpsRates.RequestedService = ServiceTypes.stUPSGround

                isGroundFreight = True
                objUpsRates.UPSAccount.Server = objUpsRates.UPSAccount.Server.Replace("ups.app/xml", "webservices")
                objUpsRates.Config("UseSOAP=true")
            End If

            objUpsRates.GetRates()

            ReDim requestedRateList(objUpsRates.Services.Count)
            For iLoop As Integer = 0 To objUpsRates.Services.Count - 1
                With requestedRateList(iLoop)
                    .ServiceType = objUpsRates.Services(iLoop).ServiceType
                    .ServiceTypeDescription = objUpsRates.Services(iLoop).ServiceTypeDescription
                    If isGroundFreight AndAlso .ServiceType = "43" Then
                        .ServiceTypeDescription = "UPS Ground Freight"
                    End If
                    .AccountNetCharge = Val(objUpsRates.Services(iLoop).AccountNetCharge & String.Empty)
                    .DeliveryTime = objUpsRates.Services(iLoop).DeliveryTime
                    .ListNetCharge = Val(objUpsRates.Services(iLoop).ListNetCharge & String.Empty)
                    .TransitTime = objUpsRates.Services(iLoop).TransitTime
                    .DeliveryDate = objUpsRates.Services(iLoop).DeliveryDate

                    If .AccountNetCharge = 0 Then
                        .AccountNetCharge = .ListNetCharge
                    End If
                End With
            Next

            Try
                objUpsRates.UPSAccount.Server = objUpsRates.UPSAccount.Server.Replace("/Rate", "/TimeInTransit")
                objUpsRates.TotalValue = RatesTotalValue
                objUpsRates.GetShippingTime()

                For outerLoop As Integer = 0 To objUpsRates.Services.Count - 1
                    For innerLoop As Integer = 0 To requestedRateList.Count - 1
                        If (requestedRateList(innerLoop).ServiceTypeDescription & String.Empty).ToUpper = objUpsRates.Services(outerLoop).ServiceTypeDescription.ToUpper _
                            OrElse requestedRateList(innerLoop).ServiceType = objUpsRates.Services(outerLoop).ServiceType Then
                            requestedRateList(innerLoop).DeliveryDate = objUpsRates.Services(outerLoop).DeliveryDate
                            requestedRateList(innerLoop).TransitTime = objUpsRates.Services(outerLoop).TransitTime
                            requestedRateList(innerLoop).DeliveryTime = objUpsRates.Services(outerLoop).DeliveryTime
                            Exit For
                        End If
                    Next
                Next
            Catch ex As Exception

            End Try

            Return requestedRateList

        Catch ex As Exception
            LastError = ex.Message
            Return Nothing

        Finally
            cRawRequest = objUpsRates.Config("RawRequest")
            cRawResponse = objUpsRates.Config("RawResponse")
            objUpsRates.Dispose()
            objUpsRates = Nothing
        End Try

    End Function

    Public Function GetUPSShippingTimeList() As RateList()

        Try

            Dim requestedRateList As RateList()
            ReDim requestedRateList(1)
            LastError = String.Empty

            objUpsRates = New nsoftware.InShip.Upsrates
            objUpsRates.RuntimeLicense = inShipLicense
            objUpsRates.Reset()
            objUpsRates.RuntimeLicense = inShipLicense

            If cServer.ToUpper.Contains("WEBSERVICES") Then
                objUpsRates.Config("UseSOAP=true")
            Else
                objUpsRates.Config("UseSOAP=false")
            End If

            objUpsRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            If Not cServer.EndsWith("/") Then
                objUpsRates.UPSAccount.Server = cServer & "/Rate"
            Else
                objUpsRates.UPSAccount.Server = cServer & "Rate"
            End If

            objUpsRates.UPSAccount.AccessKey = cUPSAccessKey
            objUpsRates.UPSAccount.UserId = cUserId
            objUpsRates.UPSAccount.AccountNumber = cAccountNumber
            objUpsRates.UPSAccount.Password = cPassword

            objUpsRates.RequestedService = cRequestedServiceType
            objUpsRates.PickupType = UPSPickupType
            objUpsRates.CustomerType = cCustomerType

            ' Insured Value is Positive, Declared Value is Negative
            Dim iPackage As Int16 = 0
            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0

            For Each shippingPackageDetail In PackageDetailList

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objUpsRates.Packages.Add(shippingPackageDetail)

                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                If Val(shippingPackageDetail.InsuredValue & String.Empty) < 0 Then
                    shippingPackageDetail.InsuredValue = Math.Abs(Val(shippingPackageDetail.InsuredValue & String.Empty))
                    objUpsRates.Config("PackageDeclaredValueType[" & iPackage & "]=0")
                End If

                ' Format weight 
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured = Val(shippingPackageDetail.InsuredValue)

                iPackage += 1
            Next

            objUpsRates.TotalWeight = Format(Val(totalWeight), "###0.0")

            With cSenderContact
                objUpsRates.SenderAddress.ZipCode = .ZipCode
                objUpsRates.SenderAddress.State = .State
                objUpsRates.SenderAddress.CountryCode = .CountryCode
            End With

            With cRecipientContact
                objUpsRates.RecipientAddress.ZipCode = .ZipCode

                If .CountryCode = "US" And .State = "PR" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .State
                ElseIf .CountryCode = "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                ElseIf .CountryCode <> "US" Then
                    objUpsRates.RecipientAddress.State = .State
                    objUpsRates.RecipientAddress.CountryCode = .CountryCode
                End If

                If .IsResidental Then
                    objUpsRates.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUpsRates.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            If IsDate(Me.ShipDate) Then
                objUpsRates.ShipDate = CDate(Me.ShipDate).ToString("yyyyMMdd")
            Else
                objUpsRates.ShipDate = DateTime.Now.ToString("yyyyMMdd")
            End If

            objUpsRates.ShipmentSpecialServices = ShipmentSpecialServices
            objUpsRates.UPSAccount.Server = objUpsRates.UPSAccount.Server.Replace("/Rate", "/TimeInTransit")
            objUpsRates.GetShippingTime()

            ReDim requestedRateList(objUpsRates.Services.Count)
            For innerLoop As Integer = 0 To objUpsRates.Services.Count - 1
                requestedRateList(innerLoop).ServiceType = objUpsRates.Services(innerLoop).ServiceType
                requestedRateList(innerLoop).DeliveryDate = objUpsRates.Services(innerLoop).DeliveryDate
                requestedRateList(innerLoop).TransitTime = objUpsRates.Services(innerLoop).TransitTime
                requestedRateList(innerLoop).DeliveryTime = objUpsRates.Services(innerLoop).DeliveryTime
                requestedRateList(innerLoop).ServiceTypeDescription = objUpsRates.Services(innerLoop).ServiceTypeDescription
                requestedRateList(innerLoop).AccountNetCharge = objUpsRates.Services(innerLoop).AccountNetCharge
                requestedRateList(innerLoop).ListNetCharge = objUpsRates.Services(innerLoop).ListNetCharge
            Next

            Return requestedRateList
        Catch ex As Exception
            LastError = ex.Message
            Return Nothing
        Finally
            cRawRequest = objUpsRates.Config("RawRequest")
            cRawResponse = objUpsRates.Config("RawResponse")
            objUpsRates.Dispose()
            objUpsRates = Nothing
        End Try

    End Function

#End Region

#Region "UPS Freight"

    Private Function RequestUpsFreightLabel() As Boolean

        Try
            LastError = String.Empty
            UPSFreightBOLID = String.Empty
            UPSFreightShipmentNumber = String.Empty
            UPSFreightLabels.Clear()

            objUpsFreight = New nsoftware.InShip.Upsfreightship
            objUpsFreight.RuntimeLicense = inShipLicense
            objUpsFreight.Reset()
            objUpsFreight.RuntimeLicense = inShipLicense

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            cServer = "https://wwwcie.ups.com/webservices"

            If Not cServer.EndsWith("/") Then
                objUpsFreight.UPSAccount.Server = cServer & "/FreightShip"
            Else
                objUpsFreight.UPSAccount.Server = cServer & "FreightShip"
            End If

            If cServer.ToUpper.Contains("WEBSERVICES") Then
                objUpsFreight.Config("UseSOAP=true")
            Else
                objUpsFreight.Config("UseSOAP=false")
            End If

            objUpsFreight.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            objUpsFreight.UPSAccount.AccessKey = cUPSAccessKey
            objUpsFreight.UPSAccount.AccountNumber = cAccountNumber
            objUpsFreight.UPSAccount.Password = cPassword
            objUpsFreight.UPSAccount.UserId = cUserId
            objUpsFreight.ServiceType = UpsfreightshipServiceTypes.stUPSFreight

            With cSenderContact
                objUpsFreight.SenderContact.Company = .Company
                objUpsFreight.SenderAddress.Address1 = .Address1
                objUpsFreight.SenderAddress.Address2 = .Address2
                objUpsFreight.Config("SenderAddress3=" & .Address3)

                objUpsFreight.SenderAddress.City = .City
                objUpsFreight.SenderAddress.ZipCode = .ZipCode
                objUpsFreight.SenderAddress.State = .State
                objUpsFreight.SenderAddress.CountryCode = .CountryCode
                objUpsFreight.SenderContact.Phone = .Phone
            End With

            With cRecipientContact
                objUpsFreight.RecipientContact.Company = .Company
                objUpsFreight.RecipientAddress.Address1 = .Address1
                objUpsFreight.RecipientAddress.Address2 = .Address2
                objUpsFreight.Config("RecipientAddress3=" & .Address3)

                objUpsFreight.RecipientAddress.City = .City
                objUpsFreight.RecipientAddress.ZipCode = .ZipCode
                objUpsFreight.RecipientAddress.State = .State
                objUpsFreight.RecipientAddress.CountryCode = .CountryCode
                objUpsFreight.RecipientContact.Phone = .Phone
            End With

            With objUpsFreight.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With

            For Each CommDetail As CommodityDetail In CommodityDetailList
                CommDetail.Weight = Format(Val(CommDetail.Weight), "###0.0")
                CommDetail.Description = CommDetail.Description.Replace("&", " ").Replace("<", " ").Replace(">", " ")
                objUpsFreight.Commodities.Add(CommDetail)
            Next

            objUpsFreight.HandlingUnit = HandlingUnit

            Dim docLabelFileName As String = String.Empty
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If
                    docLabelFileName = ShippingLabelDirectory & ShippingLabelPrefix & "_doc" & ".epl"
                End If
            End If

            Dim TotalWeight As Decimal = 0
            For Each shippingPackageDetail In PackageDetailList
                TotalWeight += Val(shippingPackageDetail.Weight) / 16
            Next

            Dim refDetail As New ReferenceDetail
            With refDetail
                .NumberOfCartons = PackageDetailList.Count
                .Number = ShippingLabelPrefix
                .ReferenceType = TFreightReferenceTypes.frtOTHER
                .Weight = TotalWeight.ToString("###0.0")
            End With
            objUpsFreight.References.Add(refDetail)

            Dim docLabel As New DocumentInfo
            UPSFreightLabels.Add(docLabelFileName)
            docLabel.DocumentType = TFreightDocumentTypes.ftcLabel
            docLabel.FileName = docLabelFileName
            docLabel.LabelsPerPage = TFreightLabelsPerPages.flppOne
            docLabel.PrintFormat = TFreightPrintFormats.fpfThermal
            docLabel.PrintSize = TFreightPrintSizes.fpsSize4X6
            objUpsFreight.Documents.Add(docLabel)

            'UPSFreightLabels.Add(docLabelFileName.Replace("_doc", "_BOL"))
            'docLabel.FileName = docLabelFileName.Replace("_doc", "_BOL")
            'docLabel.PrintFormat = TFreightPrintFormats.fpfThermal
            'docLabel.PrintSize = TFreightPrintSizes.fpsSize4X6
            'docLabel.DocumentType = TFreightDocumentTypes.ftcUPSBOL
            'objUpsFreight.Documents.Add(docLabel)

            'UPSFreightLabels.Add(docLabelFileName.Replace("_doc", "_VICS"))
            'docLabel.FileName = docLabelFileName.Replace("_doc", "_VICS")
            'docLabel.PrintFormat = TFreightPrintFormats.fpfThermal
            'docLabel.PrintSize = TFreightPrintSizes.fpsSize4X6
            'docLabel.DocumentType = TFreightDocumentTypes.ftcVICSBOL
            'objUpsFreight.Documents.Add(docLabel)

            'UPSFreightLabels.Add(docLabelFileName.Replace("_doc", "_AWB"))
            'docLabel.FileName = docLabelFileName.Replace("_doc", "_AWB")
            'docLabel.PrintFormat = TFreightPrintFormats.fpfThermal
            'docLabel.PrintSize = TFreightPrintSizes.fpsSize4X6
            'docLabel.DocumentType = TFreightDocumentTypes.ftcAWB
            'objUpsFreight.Documents.Add(docLabel)

            objUpsFreight.RuntimeLicense = inShipLicense
            objUpsFreight.GetShipmentDocuments()

            UPSFreightBOLID = objUpsFreight.BOLID
            UPSFreightShipmentNumber = objUpsFreight.ShipmentNumber

            Dim SHIP_PACKAGE_NO As String = "1"

            ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(objUpsFreight.TotalCharge))
            ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, 0)
            ShipmentSurCharge.Add(SHIP_PACKAGE_NO, 0)
            ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(objUpsFreight.TotalCharge))
            ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(objUpsFreight.TotalCharge))

            Return True

        Catch ex As nsoftware.InShip.InShipUpsfreightshipException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objUpsFreight.Config("RawRequest")
            cRawResponse = objUpsFreight.Config("RawResponse")
            objUpsFreight.Dispose()
            objUpsFreight = Nothing
        End Try

        Return True

    End Function

    Public Function GetUPSFreightRatesList() As RateList()

        Try

            Dim requestedRateList As RateList()
            ReDim requestedRateList(1)
            LastError = String.Empty

            objUpsFreightRates = New nsoftware.InShip.Upsfreightrates
            objUpsFreightRates.RuntimeLicense = inShipLicense
            objUpsFreightRates.Reset()
            objUpsFreightRates.RuntimeLicense = inShipLicense

            'For integration testing, you should direct your Freight Ship test software to:
            'https://wwwcie.ups.com/webservices/FreightShip
            'https://onlinetools.ups.com/webservices/FreightShip
            'https://onlinetools.ups.com/webservices/FreightRate

            If cServer.ToUpper.Contains("wwwcie.ups.com".ToUpper) Then
                cServer = "https://wwwcie.ups.com/webservices"
            Else
                cServer = "https://onlinetools.ups.com/webservices"
            End If

            If cServer.ToUpper.Contains("WEBSERVICES") Then
                objUpsFreightRates.Config("UseSOAP=true")
            Else
                objUpsFreightRates.Config("UseSOAP=false")
            End If

            objUpsFreightRates.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            If Not cServer.EndsWith("/") Then
                objUpsFreightRates.UPSAccount.Server = cServer & "/FreightRate"
            Else
                objUpsFreightRates.UPSAccount.Server = cServer & "FreightRate"
            End If

            objUpsFreightRates.UPSAccount.AccessKey = cUPSAccessKey
            objUpsFreightRates.UPSAccount.UserId = cUserId
            objUpsFreightRates.UPSAccount.AccountNumber = cAccountNumber
            objUpsFreightRates.UPSAccount.Password = cPassword


            objUpsFreightRates.SenderName = cSenderContact.Company
            objUpsFreightRates.SenderAddress.Address1 = cSenderContact.Address1
            objUpsFreightRates.SenderAddress.Address2 = cSenderContact.Address2
            objUpsFreightRates.SenderAddress.City = cSenderContact.City
            objUpsFreightRates.SenderAddress.State = cSenderContact.State
            objUpsFreightRates.SenderAddress.ZipCode = cSenderContact.ZipCode
            objUpsFreightRates.SenderAddress.CountryCode = cSenderContact.CountryCode

            objUpsFreightRates.RecipientName = cRecipientContact.Company
            objUpsFreightRates.RecipientAddress.Address1 = cRecipientContact.Address1
            objUpsFreightRates.RecipientAddress.Address2 = cRecipientContact.Address2
            objUpsFreightRates.RecipientAddress.City = cRecipientContact.City
            objUpsFreightRates.RecipientAddress.State = cRecipientContact.State
            objUpsFreightRates.RecipientAddress.ZipCode = cRecipientContact.ZipCode
            objUpsFreightRates.RecipientAddress.CountryCode = cRecipientContact.CountryCode

            With objUpsFreightRates.Payor
                .PayorType = Payor
                .AccountNumber = PayorContact.AccountNumber
                .CountryCode = PayorContact.CountryCode
                .ZipCode = PayorContact.ZipCode
            End With
            objUpsFreightRates.HandlingUnit = HandlingUnit

            For Each CommDetail As CommodityDetail In CommodityDetailList
                CommDetail.Weight = Format(Val(CommDetail.Weight), "###0")
                CommDetail.Description = CommDetail.Description.Replace("&", " ").Replace("<", " ").Replace(">", " ")
                objUpsFreightRates.Commodities.Add(CommDetail)
            Next

            objUpsFreightRates.GetRates()

            ReDim requestedRateList(1)
            With requestedRateList(0)
                .ServiceType = UPSFreightProductCode
                .ServiceTypeDescription = "UPS Freight"
                .AccountNetCharge = objUpsFreightRates.TotalCharge
                .ListNetCharge = objUpsFreightRates.TotalCharge
                .AccountNetCharge = objUpsFreightRates.TotalCharge
            End With

            Return requestedRateList

        Catch ex As Exception
            LastError = ex.Message
            Return Nothing
        Finally
            cRawRequest = objUpsFreightRates.Config("RawRequest")
            cRawResponse = objUpsFreightRates.Config("RawResponse")
            objUpsFreightRates.Dispose()
            objUpsFreightRates = Nothing
        End Try

    End Function

#End Region

#Region "USPS"

    Public Function USPSAddressValidation() As List(Of Contact)


        Dim addressVer As New nsoftware.InShip.Uspsaddress

        Try
            addressVer.RuntimeLicense = inShipLicense

            Select Case USPSPostageProvider
                Case USPSPostageProviders.Endicia
                    addressVer.Config("PostageProvider=1")

                    If USPSTestMode Then
                        addressVer.Config("EndiciaTestMode=1") 'Test request to Sandbox Server
                    Else
                        'addressVer.Config("EndiciaTestMode=2") 'Test request to Production Server
                    End If
                Case USPSPostageProviders.StampsCom
                    addressVer.Config("PostageProvider=2")

                Case USPSPostageProviders.PitneyBowes
                    Return Nothing

                Case Else
                    Return Nothing
            End Select

            addressVer.USPSAccount.Server = cServer
            addressVer.USPSAccount.UserId = cUserId
            addressVer.USPSAccount.Password = cPassword
            addressVer.USPSAccount.AccountNumber = cIntegrationID

            With cRecipientContact
                addressVer.Config("FullName=" & (.FirstName & " " & .LastName).ToString.Trim)

                addressVer.Address.Address1 = .Address1
                addressVer.Address.Address2 = .Address2
                addressVer.Address.City = .City
                addressVer.Address.State = .State
                addressVer.Address.ZipCode = .ZipCode
                addressVer.Company = .Company
            End With

            addressVer.ValidateAddress()

            Dim returnAddress As New Contact
            Dim returnAddressList As New List(Of Contact)

            If USPSPostageProvider = USPSPostageProviders.StampsCom Then
                If (addressVer.Config("AddressMatch") & String.Empty).Trim.ToUpper = "TRUE" Then
                    With returnAddress
                        .Address1 = cRecipientContact.Address1
                        .Address2 = cRecipientContact.Address2
                        .Address3 = cRecipientContact.Address3
                        .City = cRecipientContact.City
                        .State = cRecipientContact.State
                        .ZipCode = cRecipientContact.ZipCode
                    End With

                    returnAddressList.Add(returnAddress)
                    Return returnAddressList
                End If

                If (addressVer.Config("CityStateZipOK") & String.Empty).Trim.ToUpper = "TRUE" Then
                    'Return ""
                End If
            End If

            If addressVer.Matches.Count = 0 Then
                Return Nothing
            End If

            For iLoop As Integer = 0 To addressVer.Matches.Count - 1
                returnAddress = New Contact
                With returnAddress
                    .Address1 = addressVer.Matches(iLoop).Address1
                    .Address2 = addressVer.Matches(iLoop).Address2
                    .City = addressVer.Matches(iLoop).City
                    .State = addressVer.Matches(iLoop).State
                    .ZipCode = addressVer.Matches(iLoop).ZipCode
                End With
                returnAddressList.Add(returnAddress)
            Next

            Return returnAddressList

        Catch ex As nsoftware.InShip.InShipUspsshipException
            LastError = ex.Message
            Return Nothing

        Catch exc As Exception
            LastError = exc.Message
            Return Nothing

        Finally
            cRawRequest = addressVer.Config("RawRequest")
            cRawResponse = addressVer.Config("RawResponse")
        End Try
    End Function

    Public Function GetUSPSRatesList() As RateList()

        Try

            Dim requestedRateList As RateList()
            ReDim requestedRateList(1)
            LastError = String.Empty

            Select Case USPSPostageProvider
                Case USPSPostageProviders.Endicia
                    'objUspsRates.PostageProvider = UspsratesPostageProviders.ppEndicia
                Case USPSPostageProviders.StampsCom
                    'objUspsRates.PostageProvider = UspsratesPostageProviders.ppStamps
                Case USPSPostageProviders.PitneyBowes
                    Return GetPitneyBowesRates()
                Case Else
                    'Return Nothing
            End Select


            objUspsRates = New nsoftware.InShip.Uspsrates
            objUspsRates.RuntimeLicense = inShipLicense
            objUspsRates.Reset()
            objUspsRates.RuntimeLicense = inShipLicense

            objUspsRates.USPSAccount.Server = cServer
            objUspsRates.USPSAccount.UserId = cUserId
            objUspsRates.USPSAccount.Password = cPassword
            objUspsRates.USPSAccount.AccountNumber = cIntegrationID

            Select Case USPSPostageProvider
                Case USPSPostageProviders.Endicia
                    objUspsRates.PostageProvider = UspsratesPostageProviders.ppEndicia
                Case USPSPostageProviders.StampsCom
                    objUspsRates.PostageProvider = UspsratesPostageProviders.ppStamps
                Case USPSPostageProviders.PitneyBowes
                    Return GetPitneyBowesRates()
                Case Else
                    objUspsRates.Dispose()
                    objUspsRates = Nothing
                    Return Nothing
            End Select

            Dim shipDate As String = DateTime.Now.ToString("yyyy-MM-dd")
            objUspsRates.Config("ShipDate=" + shipDate)

            objUspsRates.RequestedService = ServiceTypes.stUnspecified
            objUspsRates.SenderAddress.ZipCode = cSenderContact.ZipCode
            objUspsRates.RecipientAddress.ZipCode = cRecipientContact.ZipCode

            ' Insured Value is Positive, Declared Value is Negative
            Dim iPackage As Int16 = 0
            Dim totalWeight As Double = 0
            Dim totalInsured As Decimal = 0

            For Each shippingPackageDetail In PackageDetailList

                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If

                objUspsRates.Packages.Add(shippingPackageDetail)

                ' Add packages (package weight is in Ounces - Convert to Pounds/Ounces)
                Dim pounds As Int16 = shippingPackageDetail.Weight \ 16
                Dim ounces As Int16 = shippingPackageDetail.Weight Mod 16
                shippingPackageDetail.Weight = pounds & " lbs"
                If ounces > 0 Then
                    shippingPackageDetail.Weight &= " " & ounces + " oz"
                End If

                'shippingPackageDetail.Size = cmboPackageSize.SelectedIndex
                shippingPackageDetail.Girth = (2 * shippingPackageDetail.Length) + (2 * shippingPackageDetail.Width)
                objUspsRates.Machinable = False

                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")

                ' Format weight 
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                totalWeight += Val(shippingPackageDetail.Weight)
                totalInsured = Val(shippingPackageDetail.InsuredValue)

                iPackage += 1
            Next

            objUspsRates.GetRates()

            ReDim requestedRateList(objUspsRates.Services.Count)
            For iLoop As Integer = 0 To objUspsRates.Services.Count - 1
                With requestedRateList(iLoop)
                    .ServiceType = objUspsRates.Services(iLoop).ServiceType
                    .ServiceTypeDescription = objUspsRates.Services(iLoop).ServiceTypeDescription.Replace("&lt;sup&gt;&amp;reg;&lt;/sup&gt;", "®")
                    .AccountNetCharge = Val(objUspsRates.Services(iLoop).AccountNetCharge & String.Empty)
                    .DeliveryTime = objUspsRates.Services(iLoop).DeliveryTime
                    .ListNetCharge = Val(objUspsRates.Services(iLoop).ListNetCharge & String.Empty)

                    .TransitTime = objUspsRates.Services(iLoop).TransitTime
                    .DeliveryDate = objUspsRates.Services(iLoop).DeliveryDate
                    'Console.WriteLine("It will take approximately " + rates.Services[i].DeliveryDay + " days to ship.")

                    If .AccountNetCharge = 0 Then
                        .AccountNetCharge = .ListNetCharge
                    End If
                End With
            Next

            Return requestedRateList

        Catch ex As nsoftware.InShip.InShipUspsshipException
            LastError = ex.Message
            Return Nothing

        Catch exc As Exception
            LastError = exc.Message
            Return Nothing

        Finally
            cRawRequest = objUspsRates.Config("RawRequest")
            cRawResponse = objUspsRates.Config("RawResponse")
            objUspsRates.Dispose()
            objUspsRates = Nothing
        End Try

    End Function

    Private Function RequestUSPSLabel() As Boolean

        Try

            LastError = String.Empty


            Select Case USPSPostageProvider
                Case USPSPostageProviders.Endicia
                    'objUspsShip.PostageProvider = UspsratesPostageProviders.ppEndicia
                Case USPSPostageProviders.StampsCom
                    'objUspsShip.PostageProvider = UspsratesPostageProviders.ppStamps
                Case USPSPostageProviders.PitneyBowes
                    Return RequestPitneyBowesLabel()
                Case Else
                    LastError = "Unknown / Invalid USPS settings"
                    Return False
            End Select

            objUspsShip = New nsoftware.InShip.Uspsship
            objUspsShip.RuntimeLicense = inShipLicense
            objUspsShip.Reset()
            objUspsShip.RuntimeLicense = inShipLicense

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                Return False
            End If

            objUspsShip.Config("Certify=true")

            objUspsShip.USPSAccount.Server = cServer
            objUspsShip.USPSAccount.UserId = cUserId
            objUspsShip.USPSAccount.Password = cPassword
            objUspsShip.USPSAccount.AccountNumber = cIntegrationID

            objUspsShip.ServiceType = cRequestedServiceType
            objUspsShip.LabelOption = "1"
            objUspsShip.ShipDate = ShipDate.ToString("yyyy-MM-dd")

            Select Case USPSPostageProvider
                Case USPSPostageProviders.Endicia
                    objUspsShip.PostageProvider = UspsratesPostageProviders.ppEndicia
                Case Else
                    objUspsShip.PostageProvider = UspsratesPostageProviders.ppStamps
            End Select

            'If USPSTestMode Then
            '    objUspsShip.Config("Certify=True")
            'End If

            With cSenderContact
                objUspsShip.SenderContact.FirstName = .FirstName
                objUspsShip.SenderContact.LastName = .LastName
                objUspsShip.SenderContact.MiddleInitial = .MiddleInitial
                objUspsShip.SenderContact.Phone = .Phone
                objUspsShip.SenderContact.Fax = .Fax
                objUspsShip.SenderContact.Email = .eMail

                objUspsShip.SenderContact.Company = .Company
                objUspsShip.SenderAddress.Address1 = .Address1
                objUspsShip.SenderAddress.Address2 = .Address2
                objUspsShip.Config("SenderAddress3=" & .Address3)

                objUspsShip.SenderAddress.City = .City
                objUspsShip.SenderAddress.ZipCode = .ZipCode
                objUspsShip.SenderAddress.State = .State
                objUspsShip.SenderAddress.CountryCode = .CountryCode

                objUspsShip.SenderAddress.CountryCode = objUspsShip.SenderAddress.CountryCode.ToUpper
                If objUspsShip.SenderAddress.CountryCode = "" OrElse objUspsShip.SenderAddress.CountryCode = "USA" Then
                    objUspsShip.SenderAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objUspsShip.SenderAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUspsShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

            With cRecipientContact
                objUspsShip.RecipientContact.FirstName = .FirstName
                objUspsShip.RecipientContact.LastName = .LastName
                objUspsShip.RecipientContact.MiddleInitial = .MiddleInitial
                objUspsShip.RecipientContact.Phone = .Phone
                objUspsShip.RecipientContact.Fax = .Fax
                objUspsShip.RecipientContact.Email = .eMail

                objUspsShip.RecipientContact.Company = .Company
                objUspsShip.RecipientAddress.Address1 = .Address1
                objUspsShip.RecipientAddress.Address2 = .Address2
                objUspsShip.Config("RecipientAddress3=" & .Address3)

                objUspsShip.RecipientAddress.City = .City
                objUspsShip.RecipientAddress.ZipCode = .ZipCode
                objUspsShip.RecipientAddress.State = .State
                objUspsShip.RecipientAddress.CountryCode = .CountryCode

                objUspsShip.RecipientAddress.CountryCode = objUspsShip.RecipientAddress.CountryCode.ToUpper
                If objUspsShip.RecipientAddress.CountryCode = "" OrElse objUspsShip.RecipientAddress.CountryCode = "USA" Then
                    objUspsShip.RecipientAddress.CountryCode = "US"
                End If

                If .IsResidental Then
                    objUspsShip.RecipientAddress.AddressFlags = &H2 'Residential
                ElseIf .IsPOBox Then
                    objUspsShip.RecipientAddress.AddressFlags = &H1 'PO Box
                End If

            End With

        Catch ex As Exception
            LastError = ex.Message
            objUspsShip.Dispose()
            objUspsShip = Nothing
            Return False
        End Try


        Try
            Select Case EzshipLabelImage
                Case EzshipLabelImageTypes.itEPL
                    objUspsShip.LabelImageType = UpsshipLabelImageTypes.uitEPL
                Case EzshipLabelImageTypes.itGIF
                    objUspsShip.LabelImageType = UpsshipLabelImageTypes.uitGIF
                Case EzshipLabelImageTypes.itSPL
                    objUspsShip.LabelImageType = UpsshipLabelImageTypes.uitSPL
                Case EzshipLabelImageTypes.itZPL
                    objUspsShip.LabelImageType = UpsshipLabelImageTypes.uitZPL
                Case Else
                    objUspsShip.LabelImageType = UpsshipLabelImageTypes.uitEPL
            End Select

            Dim extension As String = EzshipLabelImage.ToString
            If extension.StartsWith("uit") Then
                extension = "." & extension.Substring(3)
            Else
                extension = String.Empty
            End If

            ' Set shipping directory to store the labels
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If
                    For Each shippingPackageDetail In PackageDetailList
                        Dim id As String = idCtr.ToString
                        idCtr += 1
                        If Val(shippingPackageDetail.Id) > 0 Then
                            id = Val(shippingPackageDetail.Id)
                        End If

                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                        shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                        shippingPackageDetail.ReturnReceiptFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_RTN" & extension
                    Next
                End If
            End If

            ' Add packages 
            Dim TotalWeight As Decimal = 0
            Dim totalInsured As Decimal = 0
            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.InsuredValue = Format(Val(shippingPackageDetail.InsuredValue), "###0.00")
                TotalWeight += shippingPackageDetail.Weight
                totalInsured += Val(shippingPackageDetail.InsuredValue & String.Empty)
                If cSignatureRequired Then
                    shippingPackageDetail.SignatureType = TSignatureTypes.stDirect
                End If
                objUspsShip.Packages.Add(shippingPackageDetail)

                If (objUspsShip.PostageProvider = nsoftware.InShip.UspsshipPostageProviders.ppNone) Then
                    Dim pounds As Int16 = Val(shippingPackageDetail.Weight) \ 16
                    Dim ounces As Int16 = Val(shippingPackageDetail.Weight) Mod 16

                    'Note: If Oz is 0 it should be omitted from the string (just send lbs in that case)
                    shippingPackageDetail.Weight = pounds & " lbs"
                    If ounces > 0 Then
                        shippingPackageDetail.Weight &= " " & ounces & " oz"
                    End If

                    shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")

                Else 'Endicia requires weight in the form of N.N
                    shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight) / 16, "###0.0")
                End If

            Next

            objUspsShip.RuntimeLicense = inShipLicense
            objUspsShip.GetPackageLabel()

            ' Reset the object to have the updated data returned
            ' For multi UPS package shipments the total cost exists in all packages
            ' so spread the costs
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objUspsShip.Packages.Count - 1
                PackageDetailList.Add(objUspsShip.Packages(ictr))
                GetPackageCosts(objUspsShip.Packages(ictr), objUspsShip)
                Dim key As Integer = Val(objUspsShip.Packages(ictr).Id)
                ShipmentBaseCharge(key) = Math.Round(ShipmentBaseCharge(key) / objUspsShip.Packages.Count, 2)
                ShipmentDiscountCharge(key) = Math.Round(ShipmentDiscountCharge(key) / objUspsShip.Packages.Count, 2)
                ShipmentSurCharge(key) = Math.Round(ShipmentSurCharge(key) / objUspsShip.Packages.Count, 2)
                ShipmentNetCharge(key) = Math.Round(ShipmentNetCharge(key) / objUspsShip.Packages.Count, 2)
                ShipmentListCharge(key) = Math.Round(ShipmentListCharge(key) / objUspsShip.Packages.Count, 2)
            Next

            cMasterTrackingNumber = objUspsShip.Packages(0).TrackingNumber

            If cMasterTrackingNumber.Length = 0 Then
                LastError = "Shipper did not return a tracking number"
                Return False
            End If

            Return True

        Catch ex As nsoftware.InShip.InShipUpsshipException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objUspsShip.Config("RawRequest")
            cRawResponse = objUspsShip.Config("RawResponse")
            objUspsShip.Dispose()
            objUspsShip = Nothing
        End Try

        Return True

    End Function

    Public Function USPSTrack(ByVal TrackingNumber As String) As String

        Dim response As String = String.Empty

        Try

            LastError = String.Empty
            clsTrackingData = New TrackingData

            objUspsTrack = New nsoftware.InShip.Uspstrack
            objUspsTrack.RuntimeLicense = inShipLicense
            objUspsTrack.Reset()
            objUspsTrack.RuntimeLicense = inShipLicense

            objUspsTrack.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or SecurityProtocolType.Tls Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls12

            objUspsTrack.USPSAccount.UserId = cUserId
            objUspsTrack.USPSAccount.Password = cPassword
            objUspsTrack.USPSAccount.Server = cServer

            objUspsTrack.TrackShipment(TrackingNumber)

            response = String.Empty
            For index As Integer = 0 To objUspsTrack.TrackEvents.Count - 1
                response &= "Status: " & objUspsTrack.TrackEvents(index).Status & Environment.NewLine
                response &= "Date: " & objUspsTrack.TrackEvents(index).Date & Environment.NewLine
                response &= "Time: " & objUspsTrack.TrackEvents(index).Time & Environment.NewLine
                response &= "City: " & objUspsTrack.TrackEvents(index).City & Environment.NewLine
                response &= "State: " & objUspsTrack.TrackEvents(index).State & Environment.NewLine
                response &= "CountryCode: " & objUspsTrack.TrackEvents(index).CountryCode & Environment.NewLine
                response &= "Location: " & objUspsTrack.TrackEvents(index).Location & Environment.NewLine

                clsTrackingData.Status = objUspsTrack.TrackEvents(index).Status & String.Empty
                clsTrackingData.Date = objUspsTrack.TrackEvents(index).Date & String.Empty
                clsTrackingData.Time = objUspsTrack.TrackEvents(index).Time & String.Empty
                clsTrackingData.City = objUspsTrack.TrackEvents(index).City & String.Empty
                clsTrackingData.State = objUspsTrack.TrackEvents(index).State & String.Empty
                clsTrackingData.CountryCode = objUspsTrack.TrackEvents(index).CountryCode & String.Empty
                clsTrackingData.Location = objUspsTrack.TrackEvents(index).Location & String.Empty
                clsTrackingData.Address1 = objUspsTrack.TrackEvents(index).Address1 & String.Empty
                clsTrackingData.Address2 = objUspsTrack.TrackEvents(index).Address2 & String.Empty
                clsTrackingData.ZipCode = objUspsTrack.TrackEvents(index).ZipCode & String.Empty
            Next

        Catch ex As nsoftware.InShip.InShipUspsshipException
            LastError = ex.Message
            response = ex.Message

        Catch exc As Exception
            LastError = exc.Message
            response = exc.Message

        Finally
            cRawRequest = objUspsTrack.Config("RawRequest")
            cRawResponse = objUspsTrack.Config("RawResponse")
            objUspsTrack.Dispose()
            objUspsTrack = Nothing
        End Try

        Return response

    End Function


#End Region

#Region "Private Class Procedures"

    Private Sub GetPackageCosts(ByVal package As nsoftware.InShip.PackageDetail, ByVal shipObject As Object)

        Dim xdoc As New System.Xml.XmlDocument
        Dim PayorListPackageNetAmount As Decimal = 0
        Dim processingPayorList As Boolean = False
        Dim netFreight As Boolean = False
        Dim SHIP_PACKAGE_NO As String = String.Empty

        Try
            SHIP_PACKAGE_NO = package.Id

            If TypeOf shipObject Is nsoftware.InShip.Fedexshipintl Then
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, 0)
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, 0)
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                Exit Try

            ElseIf TypeOf shipObject Is nsoftware.InShip.Fedexship Then
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(package.BaseCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, Val(package.TotalDiscount))
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, Val(package.TotalSurcharges))
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(package.NetCharge))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(package.NetCharge))
                If package.RatingAggregate.Length > 0 Then
                    Try
                        Dim packageRatingAggregate As String = "<?xml version=""1.0""?>" & vbCrLf & package.RatingAggregate.Replace("v9:", "").Replace("v12:", "")
                        Dim fedexAggDoc As New System.Xml.XmlDocument
                        fedexAggDoc.LoadXml(packageRatingAggregate)
                        Dim root As XmlNode = fedexAggDoc.DocumentElement
                        Dim PAYOR_LIST_PACKAGE As XmlNode = root.SelectSingleNode("descendant::PackageRateDetails[RateType=""PAYOR_LIST_PACKAGE""]")
                        Dim listNetCharge As Double = Val(PAYOR_LIST_PACKAGE.SelectSingleNode("NetCharge/Amount").InnerText & String.Empty)
                        ShipmentListCharge(SHIP_PACKAGE_NO) = listNetCharge
                    Catch ex As Exception

                    End Try
                    Exit Try
                End If

            ElseIf TypeOf shipObject Is nsoftware.InShip.Upsship Then
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalBaseCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, 0)
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalSurcharges))
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.Config("AccountTotalNetCharge") & String.Empty))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                Exit Try
            ElseIf TypeOf shipObject Is nsoftware.InShip.Upsshipintl Then
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalBaseCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, 0)
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalSurcharges))
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.Config("AccountTotalNetCharge") & String.Empty))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalNetCharge))
                Exit Try

            Else
                ShipmentBaseCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.BaseCharge))
                ShipmentDiscountCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalDiscount))
                ShipmentSurCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.TotalSurcharges))
                ShipmentNetCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.NetCharge))
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, Val(shipObject.NetCharge))
                Exit Try
            End If

        Catch ex As Exception
            If Not ShipmentListCharge.ContainsKey(SHIP_PACKAGE_NO) Then
                ShipmentListCharge.Add(SHIP_PACKAGE_NO, PayorListPackageNetAmount)
            End If

        Finally
            If ShipmentListCharge(SHIP_PACKAGE_NO) < ShipmentNetCharge(SHIP_PACKAGE_NO) Then
                ShipmentListCharge(SHIP_PACKAGE_NO) = ShipmentNetCharge(SHIP_PACKAGE_NO)
            End If
        End Try

    End Sub

    Private Function ValidateEmail(ByVal emailAddress As String) As Boolean

        Dim strDomainName As String = String.Empty
        Dim strDomainType As String = String.Empty
        Dim strUserName As String = String.Empty
        Const sInvalidChars As String = "!#$%^&*()=+{}[]|\;:'/?>,< "
        Dim i As Integer

        If Trim(emailAddress) = "" Then
            Return False
        End If

        'Check to see if there is a double quote
        If InStr(1, emailAddress, Chr(34)) > 0 Then Return False

        'Check to see if there are consecutive dots
        If InStr(1, emailAddress, "..") > 0 Then Return False

        ' Check for invalid characters.
        If Len(emailAddress) > Len(sInvalidChars) Then
            For i = 1 To Len(sInvalidChars)
                If InStr(emailAddress, Mid(sInvalidChars, i, 1)) > 0 Then
                    Return False
                End If
            Next
        Else
            For i = 1 To Len(emailAddress)
                If InStr(sInvalidChars, Mid(emailAddress, i, 1)) > 0 Then
                    Return False
                End If
            Next
        End If

        'Check for an @ symbol
        If InStr(1, emailAddress, "@") <= 1 Then
            Return False
        End If

        If emailAddress.EndsWith("@") Then
            Return False
        End If

        strUserName = emailAddress.Substring(0, InStr(1, emailAddress, "@") - 1)
        Dim domain As String = emailAddress.Substring(InStr(1, emailAddress, "@"))

        'Check to see if there are too many @'s
        If InStr(1, domain, "@") > 0 Then
            Return False
        End If

        For Each part As String In domain.Split(".")
            If Trim(part) = "" Then
                Return False
            End If
        Next

        Return True

    End Function

#End Region

#Region "Generic Service Provider"

    ''' <summary>
    ''' Request Shipping label. Not used for Fedex Intenational
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' 
    Private Function RequestLabelOther() As Boolean

        Try
            objEzShip = New nsoftware.InShip.Ezship
            objEzShip.RuntimeLicense = inShipLicense
            objEzShip.Reset()
            objEzShip.RuntimeLicense = inShipLicense

            If cServiceProvider = ServiceProviders.Unknown Then
                LastError = "Unknown Service Type"
                objEzShip.Dispose()
                objEzShip = Nothing
                Return False
            End If

            objEzShip.Provider = cServiceProvider
            objEzShip.Account.Server = cServer
            objEzShip.Account.AccountNumber = cAccountNumber
            objEzShip.Account.UserId = cUserId
            objEzShip.Account.Password = cPassword
            objEzShip.ServiceType = cRequestedServiceType

            If cServiceProvider = ServiceProviders.FederalExpress Then
                objEzShip.Account.MeterNumber = cFedexMeterNumber
                objEzShip.Account.DeveloperKey = cFedexDeveloperKey
            Else
                objEzShip.Account.MeterNumber = String.Empty
                objEzShip.Account.DeveloperKey = String.Empty
            End If

            If cServiceProvider = ServiceProviders.UPS Then
                objEzShip.Account.AccessKey = cUPSAccessKey

                If Not objEzShip.Account.Server.EndsWith("/") Then
                    objEzShip.Account.Server &= "/"
                End If
                objEzShip.Account.Server &= "ShipConfirm"
            Else
                objEzShip.Account.AccessKey = String.Empty
            End If

            If cServiceProvider = ServiceProviders.USPS Then
                objEzShip.Config("PostageProvider=1") 'Use Endicia instead of USPS directly.
                objEzShip.Config("CustomerId=" & cUSPSEndiciaCustomerId) 'Mandatory for Endicia
                objEzShip.Config("TransactionId=" & cUSPSEndiciaTransactionId) 'Mandatory for Endicia
            End If

            objEzShip.ShipDate = ShipDate.ToString("yyyy-MM-dd")

            With cSenderContact
                objEzShip.SenderContact.FirstName = .FirstName
                objEzShip.SenderContact.LastName = .LastName
                objEzShip.SenderContact.MiddleInitial = .MiddleInitial
                objEzShip.SenderContact.Phone = .Phone
                objEzShip.SenderContact.Fax = .Fax
                objEzShip.SenderContact.Email = .eMail

                objEzShip.SenderContact.Company = .Company
                objEzShip.SenderAddress.Address1 = .Address1
                objEzShip.SenderAddress.Address2 = .Address2
                objEzShip.Config("SenderAddress3=" & .Address3)

                objEzShip.SenderAddress.City = .City
                objEzShip.SenderAddress.ZipCode = .ZipCode
                objEzShip.SenderAddress.State = .State
                objEzShip.SenderAddress.CountryCode = .CountryCode
            End With

            With cRecipientContact
                objEzShip.RecipientContact.FirstName = .FirstName
                objEzShip.RecipientContact.LastName = .LastName
                objEzShip.RecipientContact.MiddleInitial = .MiddleInitial
                objEzShip.RecipientContact.Phone = .Phone
                objEzShip.RecipientContact.Fax = .Fax
                objEzShip.RecipientContact.Email = .eMail

                objEzShip.RecipientContact.Company = .Company
                objEzShip.RecipientAddress.Address1 = .Address1
                objEzShip.RecipientAddress.Address2 = .Address2
                objEzShip.Config("RecipientAddress3=" & .Address3)

                objEzShip.RecipientAddress.City = .City
                objEzShip.RecipientAddress.ZipCode = .ZipCode
                objEzShip.RecipientAddress.State = .State
                objEzShip.RecipientAddress.CountryCode = .CountryCode
            End With

        Catch ex As Exception
            LastError = ex.Message
            objEzShip.Dispose()
            objEzShip = Nothing
            Return False
        End Try

        Try

            Dim extension As String = EzshipLabelImage.ToString
            If extension.StartsWith("it") Then
                extension = "." & extension.Substring(2)
            Else
                extension = String.Empty
            End If

            ' Fix label type if an error
            Try
                objEzShip.LabelImageType = EzshipLabelImage
            Catch ex As Exception
                Select Case cServiceProvider
                    Case ServiceProviders.CanadaPost
                        objEzShip.LabelImageType = EzshipLabelImageTypes.itZPL
                    Case ServiceProviders.FederalExpress
                        objEzShip.LabelImageType = EzshipLabelImageTypes.itEltron
                    Case ServiceProviders.FederalExpressInternational
                        objEzShip.LabelImageType = EzshipLabelImageTypes.itEltron
                    Case ServiceProviders.UPS
                        objEzShip.LabelImageType = EzshipLabelImageTypes.itZPL
                    Case ServiceProviders.USPS
                        objEzShip.LabelImageType = EzshipLabelImageTypes.itZPL
                End Select
            End Try


            ' Set shipping directory to store the labels
            Dim idCtr As Int16 = 1
            If ShippingLabelDirectory.Length > 0 AndAlso ShippingLabelPrefix.Length > 0 Then
                If My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then

                    If Not ShippingLabelDirectory.EndsWith("\") Then
                        ShippingLabelDirectory &= "\"
                    End If

                    For Each shippingPackageDetail In PackageDetailList
                        Dim id As String = idCtr.ToString
                        idCtr += 1
                        shippingPackageDetail.ShippingLabelFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & extension
                        shippingPackageDetail.CODFile = ShippingLabelDirectory & ShippingLabelPrefix & "_" & id & "_COD" & extension
                    Next
                End If
            End If

            ' Add packages
            For Each shippingPackageDetail In PackageDetailList
                shippingPackageDetail.Weight = Format(Val(shippingPackageDetail.Weight), "###0.0")
                objEzShip.Packages.Add(shippingPackageDetail)
            Next

            objEzShip.RuntimeLicense = inShipLicense
            objEzShip.GetShipmentLabels()

            ' Reset the object to have the updated data returned
            PackageDetailList.Clear()
            For ictr As Int16 = 0 To objEzShip.Packages.Count - 1
                PackageDetailList.Add(objEzShip.Packages(ictr))
                GetPackageCosts(objEzShip.Packages(ictr), objEzShip)
            Next

            If objEzShip.Packages.Count = 1 Then
                cMasterTrackingNumber = objEzShip.Packages(0).TrackingNumber
            Else
                cMasterTrackingNumber = objEzShip.MasterTrackingNumber
            End If

            Return True

        Catch ex As nsoftware.InShip.InShipException
            LastError = ex.Message
            Return False
        Catch exc As Exception
            LastError = exc.Message
            Return False
        Finally
            cRawRequest = objEzShip.Config("RawRequest")
            cRawResponse = objEzShip.Config("RawResponse")
            objEzShip.Dispose()
            objEzShip = Nothing
        End Try

        Return True

    End Function

#End Region

#Region "Internal Classes"

    Class Contact
        Public FirstName As String = String.Empty
        Public LastName As String = String.Empty
        Public MiddleInitial As String = String.Empty
        Public Phone As String = String.Empty
        Public eMail As String = String.Empty
        Public Fax As String = String.Empty
        Public Company As String = String.Empty

        ' Address Attributes
        Public Address1 As String = String.Empty
        Public Address2 As String = String.Empty
        Public Address3 As String = String.Empty
        Public City As String = String.Empty
        Public State As String = String.Empty
        Public ZipCode As String = String.Empty
        Public CountryCode As String = String.Empty
        Public IsResidental As Boolean = False
        Public IsPOBox As Boolean = False

        Public AccountNumber As String = String.Empty
    End Class

    Class SmartPost
        Public AncillaryEndorsement As String = "0"
        Public CustomerManifestId As String = String.Empty
        Public HubId As String = "5531"
        Public Indicia As String = "1"
        Public PhysicalPackaging As String = "4"
        Public TrackingNumbers As List(Of String) = New List(Of String)
    End Class

#End Region

End Class

