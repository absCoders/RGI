Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports ABSolution
Imports Newtonsoft.Json

Public Class WHFGLOBM

    Public Class OrderRequest
        Public Property OrderIds As List(Of String)
        ' Public Property HubCode As String
    End Class

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "SELECT SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.SHIP_VIA_CODE, SOTPICK1.PICK_SHIPPED, 
                                    SOTORDR1.ORDR_NO_WEB, SOTORDR1.ORDR_WEB_ID, SOTPICK1.PICK_NO
                                FROM SOTORDR1, SOTPICK1, SOTSVIA1
                                WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO
                                AND SOTORDR1.CUST_CODE = 'SKINCOM'
                                AND SOTPICK1.DATE_SHIPPED IS NULL
                                AND SOTPICK1.PICK_STATUS = 'F'
                                AND SOTORDR1.ORDR_STATUS = 'F'
                                AND SOTSVIA1.SHIP_VIA_CODE = SOTORDR1.SHIP_VIA_CODE
                                AND SOTSVIA1.CARRIER_CODE = 'GLOBAL'
                                AND SOTORDR1.ORDR_NO_WEB IS NOT NULL"

            Create_TDA(.Tables.Add, "MANIFEST", ASCMAIN1.sql, 0, False)

            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Fill_Records("SOTCARR1", String.Empty, True, "SELECT * FROM SOTCARR1")

            Create_TDA(.Tables.Add, "SOTCARR3", "*")
            Fill_Records("SOTCARR3", String.Empty, True, "SELECT * FROM SOTCARR3")

        End With

        grdManifest.DataSource = dst.Tables("MANIFEST")
        Create_Summary(grdManifest, "ORDR_NO", "Count")

        Mode_Settings(False)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Cancel"

            Case "Dispatch Orders"
                If dst.Tables("MANIFEST").Rows.Count = 0 Then
                    EMsg &= vbCr & "There are no orders to Dispatch."
                    Exit Select
                End If

                If MessageBox.Show("Do you want to Dispatch the Orders?", "Dispatch Orders", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Dispatch Orders"
                Update_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Dispatch Orders").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

        grdManifest.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("MANIFEST").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading shipments")

        EnforceConstraints(False)

        Fill_Records("MANIFEST", String.Empty, True)

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()
        Try
            BeginTrans()

            Dim lstOrderIds As New List(Of String)

            For Each drManifest As DataRow In dst.Tables("MANIFEST").Select("")
                lstOrderIds.Add("#" & drManifest.Item("ORDR_NO_WEB"))
            Next

            Dim request As New OrderRequest() With {
                        .OrderIds = lstOrderIds
                    }

            ' Production API URL
            '   https://api.global-e.com/
            'GetShippingDocuments API endpoint
            '   https://{globale_api_domain}/Order/GetShippingDocuments
            'Dispatch/ Manifest API endpoint 
            '   https://{globale_api_domain}/Order/DispatchOrders

            Dim drSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Rows.Find("GLOBAL")
            Dim endpoint As String = String.Empty
            ' https://api.global-e.com/
            endpoint = $"{drSOTCARR1.Item("CARRIER_REMOTE_HOST_IP")}Order/DispatchOrders"

            Dim drSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'GLOBAL'")(0)
            Dim MerchantGUID As String = drSOTCARR3.Item("SHIPPER_ID")

            Using client As New HttpClient()
                client.DefaultRequestHeaders.Clear()
                client.DefaultRequestHeaders.Add("MerchantGUID", MerchantGUID)
                client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

                Dim settings As New JsonSerializerSettings With {
                                .NullValueHandling = NullValueHandling.Ignore
                            }

                Dim jsonBody As String = JsonConvert.SerializeObject(request, Formatting.Indented)
                Dim content As New StringContent(jsonBody, Encoding.UTF8, "application/json")

                Dim response = client.PostAsync(endpoint, content).Result
                Dim responseBody = response.Content.ReadAsStringAsync().Result

                If Not responseBody.IsSuccess Then
                    Throw New Exception($"Unable to Dispatch Orders: {responseBody.ErrorText}")
                End If

                Dim docResponse = JsonConvert.DeserializeObject(Of OrderDocumentsResponse)(responseBody)

                Stop
                ' Need to look at the response to see if we need any data

                For Each drManifest As DataRow In dst.Tables("MANIFEST").Select("")
                    Dim PICK_NO As String = drManifest.Item("PICK_NO")
                    ASCMAIN1.sql = "UPDATE SOTPICK1 SET DATE_SHIPPED = SYSDATE WHERE PICK_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PICK_NO})
                Next
            End Using

            CommitTrans("Orders Dispatched")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Global-e Objects"

    Public Class OrderDocumentsResponse
        Public Property IsSuccess As Boolean
        Public Property ErrorText As String
        Public Property Documents As List(Of OrderDocument)
        Public Property ShipperManifests As List(Of OrderDocument)
        Public Property ParcelsTracking As List(Of ParcelTracking)
        Public Property TrackingDetails As TrackingDetails
        Public Property DeliveryAdviceInformation As List(Of DeliveryAdviceInformation)
    End Class

    Public Class OrderDocument
        Public Property CreationDateTime As DateTime?
        Public Property DocumentData As String
        Public Property DocumentExtension As String
        Public Property DocumentReference As String
        Public Property DocumentTypeCode As String
        Public Property DocumentTypeName As String
        Public Property ErrorMessage As String
        Public Property ParcelCode As String
        Public Property ShippingServiceName As String
        Public Property TrackingNumber As String
        Public Property URL As String
    End Class

    Public Class ParcelTracking
        Public Property ParcelTrackingNumber As String
        Public Property ParcelTrackingUrl As String
        Public Property ParcelCode As String
    End Class

    Public Class TrackingDetails
        ' Add fields based on what you expect in the tracking details response
        Public Property TrackingNumber As String
        Public Property ShipperName As String
        Public Property TrackingURL As String

    End Class

    Public Class DeliveryAdviceInformation
        ' Add properties per documentation if you need Delivery Advice support
        Public Property ParcelCode As String
        Public Property CommercialInvoiceNumber As String
        Public Property TotalValue As Decimal
        Public CurrencyCode As String
    End Class

#End Region

End Class