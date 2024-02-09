Imports DPayments.DShippingSDK

Public Class SHCFDXCR
    Private objFedExship As New Fedexship
    Private objFedExRates As New Fedexrates
    Private clsCredentials As New Credentials
    Private runTimeLicense As String = "42584E35414131535542524131535542313246303932303200000000000000000000000000000000444B374D344D483900004A43465542393457373133590000"
    Private clsLastError As String = String.Empty
    Private oneDayTransitServiceType As ServiceTypes = ServiceTypes.stFedExGround
    Private multiTransitServiceType As ServiceTypes = ServiceTypes.stFedExStandardOvernight
    Private tblRates As New DataTable

    Private cRawRequest As String = String.Empty
    Private cRawResponse As String = String.Empty

    Private Const SSLEnabledProtocols As Int32 = 4032


#Region "Class Enums"
    Public Enum LabelDeliveryMethods
        CarrierMailsLabelsToCustomer
        CarrierEmailLabelsToCustomer
        Printlabels
        EmailLabelsToCustomer
        UPSOneAttemptToPickup
        UPSThreeAttemptsToPickup
    End Enum

    Public Enum ServiceType
        GroundService = ServiceTypes.stFedExGround
        StandardOverNight = ServiceTypes.stFedExStandardOvernight
    End Enum

#End Region

#Region "Class Classes"

    Public Class ReturnLabelRequestDetail
        Public ShippingPackages As New List(Of Package)
        Public Shipper As New Address
        Public Recipient As New Address
        Public ServiceType As ServiceType = SHCUPSCR.ServiceType.GroundService
        Public ShipFastestMethod As Boolean = False

        Public labelEmailAddress As String = String.Empty
        Public labelDeliveryMethod As LabelDeliveryMethods = LabelDeliveryMethods.Printlabels

        Public Payor As New Payor
    End Class

    Public Class Credentials
        Public Server As String = String.Empty
        Public AccessKey As String = String.Empty
        Public AccountNumber As String = String.Empty
        Public AddressVaildationAccessLicenseNumber As String = String.Empty
        Public UserId As String = String.Empty
        Public Password As String = String.Empty
        Public LabelImageType As UpsshipLabelImageTypes = UpsshipLabelImageTypes.uitEPL
        Public FedexMeterNumber As String = String.Empty
        Public FedexDeveloperKey As String = String.Empty
    End Class

    Public Class Package
        Public PackagingType As TPackagingTypes = TPackagingTypes.ptYourPackaging
        Public Width As Int32 = 0
        Public Height As Int32 = 0
        Public Length As Int32 = 0
        Public Weight As Int32 = 0
        Public Description As String = String.Empty
        Public InsuredValue As Int32 = 0
        Public AdditionalHandling As Boolean = False
        Public LargePackage As Boolean = False
        Public ShippingLabelFile As String = String.Empty
        Public Reference As String = String.Empty
    End Class

    Public Class Payor
        Public PayorType As TPayorTypes = TPayorTypes.ptSender
        Public AccountNumber As String = String.Empty
        Public AccountZipCode As String = String.Empty
        Public CountryCode As String = String.Empty
    End Class

    Public Class Address
        Public Name As String = String.Empty
        Public Company As String = String.Empty
        Public Phone As String = String.Empty
        Public Address1 As String = String.Empty
        Public Address2 As String = String.Empty
        Public Address3 As String = String.Empty
        Public City As String = String.Empty
        Public State As String = String.Empty
        Public ZipCode As String = String.Empty
        Public isResidental As Boolean = False
        Public Country As String = String.Empty
        Public eMail As String = String.Empty
        Public isPOBox As Boolean = False
    End Class

    Public Class Response
        Public TrackingNumber As String = String.Empty
        Public TotalBaseCharge As Decimal = 0
        Public TotalNetCharge As Decimal = 0
        Public TotalSurcharges As Decimal = 0
        Public ServiceType As ServiceTypes = ServiceTypes.stFedExGround
        Public ShippingLabelFile As String = String.Empty
    End Class

#End Region

#Region "Properties"

    Public ReadOnly Property LastError As String
        Get
            Return clsLastError
        End Get
    End Property

#End Region

#Region "Instantiation"

    Public Sub New(ByVal FedExCredentials As Credentials)
        objFedExship.RuntimeLicense = runTimeLicense
        objFedExship.Reset()
        objFedExship.RuntimeLicense = runTimeLicense

        objFedExRates.RuntimeLicense = runTimeLicense
        objFedExRates.Reset()
        objFedExRates.RuntimeLicense = runTimeLicense

        clsCredentials = FedExCredentials

        objFedExship.FedExAccount.Server = clsCredentials.Server
        objFedExship.FedExAccount.AccountNumber = clsCredentials.AccountNumber
        objFedExship.FedExAccount.Password = clsCredentials.Password
        objFedExship.FedExAccount.MeterNumber = clsCredentials.FedexMeterNumber
        objFedExship.FedExAccount.DeveloperKey = clsCredentials.FedexDeveloperKey

        If clsCredentials.Server.ToUpper.Contains("WEB-SERVICES") Then
            objFedExship.Config("UseSOAP=true")
        Else
            objFedExship.Config("UseSOAP=false")
        End If

        objFedExship.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)

        With tblRates
            .Columns.Add("AccountNetCharge", GetType(System.Decimal))
            .Columns.Add("ServiceTypeDescription", GetType(System.String))
            .Columns.Add("TransitTime", GetType(System.String))
            .Columns.Add("DeliveryTime", GetType(System.String))
            .Columns.Add("ServiceType", GetType(System.Int32))
        End With
    End Sub

#End Region

#Region "Procedures"

    Public Function RequestReturnLabel(ByVal ShipDetails As ReturnLabelRequestDetail) As Response
        clsLastError = String.Empty
        RequestReturnLabel = New Response
        Try
            For Each pkg As Package In ShipDetails.ShippingPackages
                If Not BuildPackages(pkg) Then
                    Return Nothing
                End If
            Next

            ValidateInputFieldLengths(ShipDetails.Recipient.Address1, _
            ShipDetails.Recipient.Address2, _
            ShipDetails.Recipient.Address3, _
            ShipDetails.Recipient.Name, _
            ShipDetails.Recipient.City, _
            ShipDetails.Recipient.State, _
            ShipDetails.Recipient.ZipCode, _
            ShipDetails.Recipient.Company)

            ValidateInputFieldLengths(ShipDetails.Shipper.Address1, _
            ShipDetails.Shipper.Address2, _
            ShipDetails.Shipper.Address3, _
            ShipDetails.Shipper.Name, _
            ShipDetails.Shipper.City, _
            ShipDetails.Shipper.State, _
            ShipDetails.Shipper.ZipCode, _
            ShipDetails.Shipper.Company)

            If Not GetShipper(ShipDetails.Shipper) Then
                Return Nothing
            End If

            If Not GetSender(ShipDetails.Recipient) Then
                Return Nothing
            End If

            If Not GetRecipient(ShipDetails.Shipper) Then
                Return Nothing
            End If

            If Not GetPayor(ShipDetails.Payor) Then
                Return Nothing
            End If

            If Not GetServiceType(ShipDetails) Then
                Return Nothing
            End If

            If Not SetReturnLabelSpecialServices() Then
                Return Nothing
            End If

            objFedExship.SenderContact.Email = ShipDetails.Shipper.eMail
            objFedExship.RecipientContact.Email = ShipDetails.Shipper.eMail
            objFedExship.SenderContact.Email = ShipDetails.Recipient.eMail

            Select Case ShipDetails.labelDeliveryMethod
                Case LabelDeliveryMethods.CarrierMailsLabelsToCustomer
                    objFedExship.Config("ReturnPrintAndMail")
                Case LabelDeliveryMethods.CarrierEmailLabelsToCustomer
                    objFedExship.Config("ElectronicReturnLabel")
                Case LabelDeliveryMethods.EmailLabelsToCustomer
                    objFedExship.LabelImageType = FedexshipLabelImageTypes.fitPNG
                Case LabelDeliveryMethods.Printlabels
                    objFedExship.LabelImageType = FedexshipLabelImageTypes.fitEltron
                Case LabelDeliveryMethods.UPSOneAttemptToPickup
                    objFedExship.Config("ReturnServiceFirstAttempt")
                Case LabelDeliveryMethods.UPSThreeAttemptsToPickup
                    objFedExship.Config("ReturnServiceThirdAttempt")
            End Select

            objFedExship.GetShipmentLabels()

            Dim Resp As New Response
            Resp.TrackingNumber = objFedExship.Packages(0).TrackingNumber
            Resp.ServiceType = objFedExship.ServiceType
            If tblRates.Select("ServiceType = " & Resp.ServiceType).Length > 0 Then
                Dim rowRates As DataRow = tblRates.Select("ServiceType = " & Resp.ServiceType)(0)
                Resp.TotalBaseCharge = Val(rowRates.Item("AccountNetCharge") & String.Empty)
                Resp.TotalNetCharge = Val(rowRates.Item("AccountNetCharge") & String.Empty)
                Resp.TotalSurcharges = 0
            Else
                Resp.TotalBaseCharge = Val(objFedExship.TotalNetCharge & String.Empty)
                Resp.TotalNetCharge = Val(objFedExship.TotalNetCharge & String.Empty)
                Resp.TotalSurcharges = 0
                For Each pkg As PackageDetail In objFedExship.Packages
                    Resp.TotalSurcharges += Val(pkg.TotalSurcharges & String.Empty)
                Next
            End If
            Resp.ShippingLabelFile = objFedExship.Packages(0).ShippingLabelFile
            Return Resp

        Catch ex As Exception
            RequestReturnLabel = Nothing
            clsLastError = "RequestReturnLabel Error: " & ex.Message
        Finally
            cRawRequest = objFedExship.Config("RawRequest")
            cRawResponse = objFedExship.Config("RawResponse")
        End Try
    End Function

    Private Function BuildPackages(ByVal PackageDetails As Package) As Boolean
        clsLastError = String.Empty
        BuildPackages = True
        Try
            objFedExship.Packages.Add(New PackageDetail())
            Dim packageIndex As Int16 = objFedExship.Packages.Count - 1

            objFedExship.Packages(packageIndex).PackagingType = PackageDetails.PackagingType
            objFedExship.Packages(packageIndex).Width = PackageDetails.Width
            objFedExship.Packages(packageIndex).Height = PackageDetails.Height
            objFedExship.Packages(packageIndex).Length = PackageDetails.Length

            ' Convert Ounces to Pounds
            objFedExship.Packages(packageIndex).Weight = Format(PackageDetails.Weight, "##0.0")

            ' For FedEx, this is the description that appears in the email to identify this package. This is optional.
            'If (PackageDetails.Description & String.Empty).Trim.Length > 0 Then
            '    objFedExship.Packages(packageIndex).Description = PackageDetails.Description
            'Else
            '    objFedExship.Packages(packageIndex).Description = "Return Label"
            'End If

            objFedExship.Packages(packageIndex).InsuredValue = PackageDetails.InsuredValue
            objFedExship.Packages(packageIndex).Reference = PackageDetails.Reference & String.Empty

            Dim specialService As Integer = 0
            If PackageDetails.AdditionalHandling Then
                specialService = specialService Or &H1
            End If
            If PackageDetails.LargePackage Then
                specialService = specialService Or &H2
            End If

            objFedExship.Packages(packageIndex).SpecialServices = specialService
            objFedExship.Packages(packageIndex).ShippingLabelFile = PackageDetails.ShippingLabelFile

        Catch ex As Exception
            clsLastError = "BuildPackages Error: " & ex.Message
            BuildPackages = False
        End Try
    End Function

    Private Function GetShipper(ByVal ShipperAddress As Address) As Boolean

        clsLastError = String.Empty
        GetShipper = True

        Try
            objFedExship.SenderContact.Company = ShipperAddress.Company

            objFedExship.SenderContact.FirstName = ShipperAddress.Name
            If objFedExship.SenderContact.FirstName.Length = 0 Then
                objFedExship.SenderContact.FirstName = "Returns Department"
            End If

            objFedExship.SenderContact.Company = ShipperAddress.Company
            If objFedExship.SenderContact.Company.Length = 0 Then
                objFedExship.SenderContact.Company = objFedExship.SenderContact.FirstName
            End If

            objFedExship.SenderContact.Phone = ShipperAddress.Phone
            objFedExship.SenderAddress.Address1 = ShipperAddress.Address1
            objFedExship.SenderAddress.Address2 = ShipperAddress.Address2

            If ShipperAddress.Address3.Length > 0 Then
                objFedExship.Config("SenderAddress3=" & ShipperAddress.Address3)
            End If

            objFedExship.SenderAddress.City = ShipperAddress.City
            objFedExship.SenderAddress.State = ShipperAddress.State
            objFedExship.SenderAddress.ZipCode = ShipperAddress.ZipCode
            objFedExship.SenderContact.Email = ShipperAddress.eMail

        Catch ex As Exception
            GetShipper = False
            clsLastError = "GetShipper Error: " & ex.Message
        End Try

    End Function

    Private Function GetSender(ByVal SenderAddress As Address) As Boolean
        clsLastError = String.Empty
        GetSender = True
        Try
            objFedExship.SenderContact.FirstName = SenderAddress.Name
            If objFedExship.SenderContact.FirstName.Length = 0 Then
                objFedExship.SenderContact.FirstName = "Returns Department"
            End If
            objFedExship.SenderContact.Company = SenderAddress.Company
            If objFedExship.SenderContact.Company.Length = 0 Then
                objFedExship.SenderContact.Company = objFedExship.SenderContact.FirstName
            End If

            objFedExship.SenderContact.Phone = SenderAddress.Phone
            objFedExship.SenderAddress.Address1 = SenderAddress.Address1
            objFedExship.SenderAddress.Address2 = SenderAddress.Address2

            If SenderAddress.Address3.Length > 0 Then
                objFedExship.Config("SenderAddress3=" & SenderAddress.Address3)
            End If

            objFedExship.SenderAddress.City = SenderAddress.City
            objFedExship.SenderAddress.State = SenderAddress.State
            objFedExship.SenderAddress.ZipCode = SenderAddress.ZipCode
            objFedExship.SenderContact.Email = SenderAddress.eMail

        Catch ex As Exception
            GetSender = False
            clsLastError = "GetSender Error: " & ex.Message
        End Try
    End Function

    Private Function GetRecipient(ByVal RecipientAddress As Address) As Boolean
        clsLastError = String.Empty
        GetRecipient = True
        Try
            objFedExship.RecipientContact.FirstName = RecipientAddress.Name
            objFedExship.RecipientContact.Company = RecipientAddress.Company
            If objFedExship.RecipientContact.Company.Length = 0 Then
                objFedExship.RecipientContact.Company = objFedExship.RecipientContact.FirstName
            End If

            objFedExship.RecipientContact.Phone = RecipientAddress.Phone
            objFedExship.RecipientAddress.Address1 = RecipientAddress.Address1
            objFedExship.RecipientAddress.Address2 = RecipientAddress.Address2

            If RecipientAddress.Address3.Length > 0 Then
                objFedExship.Config("RecipientAddress3=" & RecipientAddress.Address3)
            End If

            objFedExship.RecipientAddress.City = RecipientAddress.City
            objFedExship.RecipientAddress.State = RecipientAddress.State
            objFedExship.RecipientAddress.ZipCode = RecipientAddress.ZipCode

            If (RecipientAddress.isResidental) Then
                objFedExship.RecipientAddress.AddressFlags = &H2 'Residential
            End If

            objFedExship.RecipientContact.Email = RecipientAddress.eMail

        Catch ex As Exception
            GetRecipient = False
            clsLastError = "GetRecipient Error: " & ex.Message
        End Try
    End Function

    Private Function GetPayor(ByVal Payor As Payor) As Boolean
        clsLastError = String.Empty
        GetPayor = True
        Try
            objFedExship.Payor.PayorType = Payor.PayorType
            objFedExship.Payor.AccountNumber = Payor.AccountNumber
            objFedExship.Payor.ZipCode = Payor.AccountZipCode
            objFedExship.Payor.CountryCode = Payor.CountryCode

        Catch ex As Exception
            GetPayor = False
            clsLastError = "GetPayor Error: " & ex.Message
        End Try
    End Function

    Private Function GetServiceType(ByVal ShipDetails As ReturnLabelRequestDetail) As Boolean
        clsLastError = String.Empty
        GetServiceType = True
        Try
            objFedExship.ServiceType = ShipDetails.ServiceType
            If ShipDetails.ShipFastestMethod Then
                Dim oneDayTransitServicetypeCost As Decimal = 0
                Dim multiTransitServicetypeCost As Decimal = 0
                ' Need to get transit time.
                Dim tblRates As DataTable = GetServiceTypeBasedOnTimeInTransit(ShipDetails)
                If tblRates.Rows.Count > 0 Then
                    ' Next Day Saver returns an error from UPS = Not a valid Ship Method for Return Label
                    For Each ServiceType As Int32 In New Int32() {oneDayTransitServiceType, multiTransitServiceType}
                        If tblRates.Select("ServiceType = " & ServiceType).Length > 0 Then
                            Select Case ServiceType
                                Case oneDayTransitServiceType
                                    oneDayTransitServicetypeCost = Val(tblRates.Select("ServiceType = " & ServiceType)(0).Item("AccountNetCharge") & String.Empty)
                                Case multiTransitServiceType
                                    multiTransitServicetypeCost = Val(tblRates.Select("ServiceType = " & ServiceType)(0).Item("AccountNetCharge") & String.Empty)
                            End Select
                        End If
                    Next
                End If
            End If

        Catch ex As Exception
            GetServiceType = False
            clsLastError = "GetServiceType Error: " & ex.Message
        End Try
    End Function

    Private Function SetReturnLabelSpecialServices() As Boolean
        clsLastError = String.Empty
        SetReturnLabelSpecialServices = True
        Try
            Dim specialService As Long = 0
            specialService = specialService Or &H8000000L ' Return Label 
            objFedExship.ShipmentSpecialServices = specialService

        Catch ex As Exception
            SetReturnLabelSpecialServices = False
            clsLastError = "SetReturnLabelSpecialServices Error: " & ex.Message
        End Try
    End Function

    Public Function GetServiceTypeBasedOnTimeInTransit(ByVal ShipDetails As ReturnLabelRequestDetail) As DataTable
        clsLastError = String.Empty
        GetServiceTypeBasedOnTimeInTransit = New DataTable
        tblRates.Rows.Clear()

        If 1 = 1 Then
            Return tblRates
        End If

        Try
            objFedExRates.RequestedService = ServiceTypes.stUnspecified
            'objFedExRates.PickupType = UpsratesPickupTypes.ptDailyPickup

            Dim packagesIndex As Int16 = 0
            For Each pkg As Package In ShipDetails.ShippingPackages
                objFedExRates.Packages.Add(New PackageDetail())
                objFedExRates.Packages(packagesIndex).PackagingType = pkg.PackagingType
                'objFedExRates.CustomerType = UpsratesCustomerTypes.ccRetail
                objFedExRates.Packages(packagesIndex).Weight = pkg.Weight
                objFedExRates.Packages(packagesIndex).Length = pkg.Length
                objFedExRates.Packages(packagesIndex).Width = pkg.Width
                objFedExRates.Packages(packagesIndex).Height = pkg.Height
                packagesIndex = 1
            Next

            objFedExRates.SenderAddress.State = ShipDetails.Recipient.State
            objFedExRates.SenderAddress.ZipCode = ShipDetails.Recipient.ZipCode
            objFedExRates.SenderAddress.CountryCode = ShipDetails.Recipient.Country
            objFedExRates.RecipientAddress.State = ShipDetails.Shipper.State
            objFedExRates.RecipientAddress.ZipCode = ShipDetails.Shipper.ZipCode
            objFedExRates.RecipientAddress.CountryCode = ShipDetails.Shipper.Country
            objFedExRates.GetRates()

            For i As Integer = 0 To objFedExRates.Services.Count - 1
                Dim rowRates As DataRow = tblRates.NewRow
                If (objFedExRates.Services(i).AccountNetCharge <> "") Then
                    rowRates.Item("AccountNetCharge") = Convert.ToDecimal(objFedExRates.Services(i).AccountNetCharge)
                Else
                    rowRates.Item("AccountNetCharge") = Convert.ToDecimal(objFedExRates.Services(i).ListNetCharge)
                End If
                rowRates.Item("ServiceTypeDescription") = objFedExRates.Services(i).ServiceTypeDescription & String.Empty
                rowRates.Item("TransitTime") = objFedExRates.Services(i).TransitTime & String.Empty
                rowRates.Item("DeliveryTime") = objFedExRates.Services(i).DeliveryTime & String.Empty
                rowRates.Item("ServiceType") = objFedExRates.Services(i).ServiceType
                tblRates.Rows.Add(rowRates)
            Next
            Return tblRates
        Catch ex As Exception
            clsLastError = "GetServiceTypeBasedOnTimeInTransit Error: " & ex.Message
            Return tblRates
        Finally
        End Try
    End Function

    Private Sub ValidateInputFieldLengths(ByRef AddressLine1 As String, _
    ByRef AddressLine2 As String, _
    ByRef AddressLine3 As String, _
    ByRef AttentionName As String, _
    ByRef City As String, _
    ByRef StateProvinceCode As String, _
    ByRef PostalCode As String, _
    ByRef CompanyName As String)
        ' MT 7103 - validate fiels lengths
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

#End Region

End Class

