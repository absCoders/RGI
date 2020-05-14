Imports nsoftware.InShip
Public Class SHCUPSCR
    Private objUpsship As New nsoftware.InShip.Upsship
    Private objUpsrates As New nsoftware.InShip.Upsrates
    Private clsCredentials As New Credentials
    Private clsLastError As String = String.Empty
    Private oneDayTransitServiceType As nsoftware.InShip.ServiceTypes = nsoftware.InShip.ServiceTypes.stUPSGround
    Private multiTransitServiceType As nsoftware.InShip.ServiceTypes = nsoftware.InShip.ServiceTypes.stUPSNextDayAir
    Private tblRates As New DataTable

    Private cRawRequest As String = String.Empty
    Private cRawResponse As String = String.Empty

    Private Const SSLEnabledProtocols As Int32 = 4032
    Private inShipLicense As String = ASCMAIN1.nSoftwareKeys("nSoftwareInship")


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
        GroundService = nsoftware.InShip.ServiceTypes.stUPSGround
        NextDayService = nsoftware.InShip.ServiceTypes.stUPSNextDayAir
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
        Public LabelImageType As nsoftware.InShip.UpsshipLabelImageTypes = UpsshipLabelImageTypes.uitEPL
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
        Public PayorType As nsoftware.InShip.TPayorTypes = TPayorTypes.ptSender
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
        Public ServiceType As nsoftware.InShip.ServiceTypes = ServiceTypes.stUPSGround
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

    Public Sub New(ByVal UPSCredentials As Credentials)
        objUpsship.RuntimeLicense = inShipLicense
        objUpsship.Reset()
        objUpsship.RuntimeLicense = inShipLicense

        objUpsrates.RuntimeLicense = inShipLicense
        objUpsrates.Reset()
        objUpsrates.RuntimeLicense = inShipLicense

        clsCredentials = UPSCredentials

        With clsCredentials
            If Not .Server.EndsWith("/") Then
                .Server &= "/"
            End If
            If .Server.ToUpper.Contains("WEBSERVICES") Then
                objUpsship.Config("UseSOAP=true")
                objUpsship.UPSAccount.Server = .Server & "Ship"
            Else
                objUpsship.Config("UseSOAP=false")
                objUpsship.UPSAccount.Server = .Server & "ShipConfirm"
            End If
            objUpsship.LabelImageType = .LabelImageType
            objUpsship.UPSAccount.AccessKey = .AccessKey
            objUpsship.UPSAccount.AccountNumber = .AccountNumber
            objUpsship.UPSAccount.UserId = .UserId
            objUpsship.UPSAccount.Password = .Password
            If .Server.ToUpper.Contains("WEBSERVICES") Then
                objUpsrates.Config("UseSOAP=true")
            Else
                objUpsrates.Config("UseSOAP=false")
            End If
            objUpsrates.UPSAccount.Server = .Server & "Rate"
            objUpsrates.UPSAccount.AccessKey = .AccessKey
            objUpsrates.UPSAccount.AccountNumber = .AccountNumber
            objUpsrates.UPSAccount.UserId = .UserId
            objUpsrates.UPSAccount.Password = .Password
        End With

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

            objUpsship.AccountContact.Email = ShipDetails.Shipper.eMail
            objUpsship.RecipientContact.Email = ShipDetails.Shipper.eMail
            objUpsship.SenderContact.Email = ShipDetails.Recipient.eMail

            Select Case ShipDetails.labelDeliveryMethod
                Case LabelDeliveryMethods.CarrierMailsLabelsToCustomer
                    objUpsship.Config("ReturnPrintAndMail")
                Case LabelDeliveryMethods.CarrierEmailLabelsToCustomer
                    objUpsship.Config("ElectronicReturnLabel")
                Case LabelDeliveryMethods.EmailLabelsToCustomer
                    objUpsship.LabelImageType = UpsshipLabelImageTypes.uitGIF
                Case LabelDeliveryMethods.Printlabels
                    objUpsship.LabelImageType = UpsshipLabelImageTypes.uitEPL
                Case LabelDeliveryMethods.UPSOneAttemptToPickup
                    objUpsship.Config("ReturnServiceFirstAttempt")
                Case LabelDeliveryMethods.UPSThreeAttemptsToPickup
                    objUpsship.Config("ReturnServiceThirdAttempt")
            End Select

            objUpsship.Config("SSLEnabledProtocols=" & SSLEnabledProtocols)
            objUpsship.GetShipmentLabels()

            Dim Resp As New Response
            Resp.TrackingNumber = objUpsship.Packages(0).TrackingNumber
            Resp.ServiceType = objUpsship.ServiceType
            If tblRates.Select("ServiceType = " & Resp.ServiceType).Length > 0 Then
                Dim rowRates As DataRow = tblRates.Select("ServiceType = " & Resp.ServiceType)(0)
                Resp.TotalBaseCharge = Val(rowRates.Item("AccountNetCharge") & String.Empty)
                Resp.TotalNetCharge = Val(rowRates.Item("AccountNetCharge") & String.Empty)
                Resp.TotalSurcharges = 0
            Else
                Resp.TotalBaseCharge = objUpsship.TotalBaseCharge
                Resp.TotalNetCharge = objUpsship.TotalNetCharge
                Resp.TotalSurcharges = objUpsship.TotalSurcharges
            End If
            Resp.ShippingLabelFile = objUpsship.Packages(0).ShippingLabelFile
            Return Resp

        Catch ex As Exception
            RequestReturnLabel = Nothing
            clsLastError = "RequestReturnLabel Error: " & ex.Message
        Finally
            cRawRequest = objUpsship.Config("RawRequest")
            cRawResponse = objUpsship.Config("RawResponse")
        End Try
    End Function

    Private Function BuildPackages(ByVal PackageDetails As Package) As Boolean
        clsLastError = String.Empty
        BuildPackages = True
        Try
            objUpsship.Packages.Add(New PackageDetail())
            Dim packageIndex As Int16 = objUpsship.Packages.Count - 1

            objUpsship.Packages(packageIndex).PackagingType = PackageDetails.PackagingType
            objUpsship.Packages(packageIndex).Width = PackageDetails.Width
            objUpsship.Packages(packageIndex).Height = PackageDetails.Height
            objUpsship.Packages(packageIndex).Length = PackageDetails.Length

            ' Convert Ounces to Pounds
            objUpsship.Packages(packageIndex).Weight = Format(PackageDetails.Weight / 16, "##0.0")

            If (PackageDetails.Description & String.Empty).Trim.Length > 0 Then
                objUpsship.Packages(packageIndex).Description = PackageDetails.Description
            Else
                objUpsship.Packages(packageIndex).Description = "Return Label"
            End If

            objUpsship.Packages(packageIndex).InsuredValue = PackageDetails.InsuredValue

            Dim specialService As Integer = 0
            If PackageDetails.AdditionalHandling Then
                specialService = specialService Or &H1
            End If
            If PackageDetails.LargePackage Then
                specialService = specialService Or &H2
            End If

            objUpsship.Packages(packageIndex).SpecialServices = specialService
            objUpsship.Packages(packageIndex).ShippingLabelFile = PackageDetails.ShippingLabelFile

        Catch ex As Exception
            clsLastError = "BuildPackages Error: " & ex.Message
            BuildPackages = False
        End Try
    End Function

    Private Function GetShipper(ByVal ShipperAddress As Address) As Boolean

        clsLastError = String.Empty
        GetShipper = True

        Try
            objUpsship.AccountContact.Company = ShipperAddress.Company

            objUpsship.AccountContact.FirstName = ShipperAddress.Name
            If objUpsship.AccountContact.FirstName.Length = 0 Then
                objUpsship.AccountContact.FirstName = "Returns Department"
            End If

            objUpsship.AccountContact.Company = ShipperAddress.Company
            If objUpsship.AccountContact.Company.Length = 0 Then
                objUpsship.AccountContact.Company = objUpsship.SenderContact.FirstName
            End If

            objUpsship.AccountContact.Phone = ShipperAddress.Phone
            objUpsship.AccountAddress.Address1 = ShipperAddress.Address1
            objUpsship.AccountAddress.Address2 = ShipperAddress.Address2

            If ShipperAddress.Address3.Length > 0 Then
                objUpsship.Config("SenderAddress3=" & ShipperAddress.Address3)
            End If

            objUpsship.AccountAddress.City = ShipperAddress.City
            objUpsship.AccountAddress.State = ShipperAddress.State
            objUpsship.AccountAddress.ZipCode = ShipperAddress.ZipCode
            objUpsship.AccountContact.Email = ShipperAddress.eMail

        Catch ex As Exception
            GetShipper = False
            clsLastError = "GetShipper Error: " & ex.Message
        End Try

    End Function

    Private Function GetSender(ByVal SenderAddress As Address) As Boolean
        clsLastError = String.Empty
        GetSender = True
        Try
            objUpsship.SenderContact.FirstName = SenderAddress.Name
            If objUpsship.SenderContact.FirstName.Length = 0 Then
                objUpsship.SenderContact.FirstName = "Returns Department"
            End If
            objUpsship.SenderContact.Company = SenderAddress.Company
            If objUpsship.SenderContact.Company.Length = 0 Then
                objUpsship.SenderContact.Company = objUpsship.SenderContact.FirstName
            End If

            objUpsship.SenderContact.Phone = SenderAddress.Phone
            objUpsship.SenderAddress.Address1 = SenderAddress.Address1
            objUpsship.SenderAddress.Address2 = SenderAddress.Address2

            If SenderAddress.Address3.Length > 0 Then
                objUpsship.Config("SenderAddress3=" & SenderAddress.Address3)
            End If

            objUpsship.SenderAddress.City = SenderAddress.City
            objUpsship.SenderAddress.State = SenderAddress.State
            objUpsship.SenderAddress.ZipCode = SenderAddress.ZipCode
            objUpsship.SenderContact.Email = SenderAddress.eMail

        Catch ex As Exception
            GetSender = False
            clsLastError = "GetSender Error: " & ex.Message
        End Try
    End Function

    Private Function GetRecipient(ByVal RecipientAddress As Address) As Boolean
        clsLastError = String.Empty
        GetRecipient = True
        Try
            objUpsship.RecipientContact.FirstName = RecipientAddress.Name
            objUpsship.RecipientContact.Company = RecipientAddress.Company
            If objUpsship.RecipientContact.Company.Length = 0 Then
                objUpsship.RecipientContact.Company = objUpsship.RecipientContact.FirstName
            End If

            objUpsship.RecipientContact.Phone = RecipientAddress.Phone
            objUpsship.RecipientAddress.Address1 = RecipientAddress.Address1
            objUpsship.RecipientAddress.Address2 = RecipientAddress.Address2

            If RecipientAddress.Address3.Length > 0 Then
                objUpsship.Config("RecipientAddress3=" & RecipientAddress.Address3)
            End If

            objUpsship.RecipientAddress.City = RecipientAddress.City
            objUpsship.RecipientAddress.State = RecipientAddress.State
            objUpsship.RecipientAddress.ZipCode = RecipientAddress.ZipCode

            If (RecipientAddress.isResidental) Then
                objUpsship.RecipientAddress.AddressFlags = &H2 'Residential
            End If

            objUpsship.RecipientContact.Email = RecipientAddress.eMail

        Catch ex As Exception
            GetRecipient = False
            clsLastError = "GetRecipient Error: " & ex.Message
        End Try
    End Function

    Private Function GetPayor(ByVal Payor As Payor) As Boolean
        clsLastError = String.Empty
        GetPayor = True
        Try
            objUpsship.Payor.PayorType = Payor.PayorType
            objUpsship.Payor.AccountNumber = Payor.AccountNumber
            objUpsship.Payor.ZipCode = Payor.AccountZipCode

        Catch ex As Exception
            GetPayor = False
            clsLastError = "GetPayor Error: " & ex.Message
        End Try
    End Function

    Private Function GetServiceType(ByVal ShipDetails As ReturnLabelRequestDetail) As Boolean
        clsLastError = String.Empty
        GetServiceType = True
        Try
            objUpsship.ServiceType = ShipDetails.ServiceType
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
            objUpsship.ShipmentSpecialServices = specialService

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
            objUpsrates.RequestedService = ServiceTypes.stUnspecified
            objUpsrates.PickupType = UpsratesPickupTypes.ptDailyPickup

            Dim packagesIndex As Int16 = 0
            For Each pkg As Package In ShipDetails.ShippingPackages
                objUpsrates.Packages.Add(New PackageDetail())
                objUpsrates.Packages(packagesIndex).PackagingType = pkg.PackagingType
                objUpsrates.CustomerType = UpsratesCustomerTypes.ccRetail
                objUpsrates.Packages(packagesIndex).Weight = pkg.Weight
                objUpsrates.Packages(packagesIndex).Length = pkg.Length
                objUpsrates.Packages(packagesIndex).Width = pkg.Width
                objUpsrates.Packages(packagesIndex).Height = pkg.Height
                packagesIndex = 1
            Next

            objUpsrates.SenderAddress.State = ShipDetails.Recipient.State
            objUpsrates.SenderAddress.ZipCode = ShipDetails.Recipient.ZipCode
            objUpsrates.SenderAddress.CountryCode = ShipDetails.Recipient.Country
            objUpsrates.RecipientAddress.State = ShipDetails.Shipper.State
            objUpsrates.RecipientAddress.ZipCode = ShipDetails.Shipper.ZipCode
            objUpsrates.RecipientAddress.CountryCode = ShipDetails.Shipper.Country
            objUpsrates.GetRates()

            For i As Integer = 0 To objUpsrates.Services.Count - 1
                Dim rowRates As DataRow = tblRates.NewRow
                If (objUpsrates.Services(i).AccountNetCharge <> "") Then
                    rowRates.Item("AccountNetCharge") = Convert.ToDecimal(objUpsrates.Services(i).AccountNetCharge)
                Else
                    rowRates.Item("AccountNetCharge") = Convert.ToDecimal(objUpsrates.Services(i).ListNetCharge)
                End If
                rowRates.Item("ServiceTypeDescription") = objUpsrates.Services(i).ServiceTypeDescription & String.Empty
                rowRates.Item("TransitTime") = objUpsrates.Services(i).TransitTime & String.Empty
                rowRates.Item("DeliveryTime") = objUpsrates.Services(i).DeliveryTime & String.Empty
                rowRates.Item("ServiceType") = objUpsrates.Services(i).ServiceType
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

