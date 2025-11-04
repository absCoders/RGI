Imports System.IO
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json.Linq

Public Class SOCSHOPF

    Private rowECTECOMD As DataRow = Nothing
    Private Const ECOM_CODE As String = "SHOPIFY"
    Private ECOM_CUST_CODE As String = String.Empty
    Private ECOM_CUST_ADDR_CODE As String = String.Empty
    Private ECOM_SITE_USER As String = String.Empty
    Private ECOM_SITE_PWD As String = String.Empty
    Private ShopifyUrl As String = String.Empty
    Private shopAccessToken As String = String.Empty
    Public LastError As String = String.Empty


    Public Sub New()

        rowECTECOMD = ASCDATA1.GetDataRow("SELECT * FROM ECTECOMD WHERE ECOM_CODE = :PARM1", "V", {ECOM_CODE})
        If rowECTECOMD Is Nothing Then
            Throw New Exception($"ECOM_CODE: {ECOM_CODE} not found in ECTECOMD")
        End If

        ECOM_CUST_CODE = rowECTECOMD.Item("ECOM_CUST_CODE") & String.Empty
        ECOM_CUST_ADDR_CODE = rowECTECOMD.Item("ECOM_CUST_ADDR_CODE") & String.Empty
        ECOM_SITE_USER = rowECTECOMD.Item("ECOM_SITE_USER") & String.Empty
        ECOM_SITE_PWD = rowECTECOMD.Item("ECOM_SITE_PWD") & String.Empty

        ShopifyUrl = rowECTECOMD.Item("ECOM_URL") & String.Empty
        ShopifyUrl = ShopifyUrl.Replace("{USER_ID}", ECOM_SITE_USER).Replace("{PASSWORD}", ECOM_SITE_PWD)
        shopAccessToken = rowECTECOMD.Item("ECOM_SITE_SECRET_KEY") & String.Empty

    End Sub

    Public Class Receipt
        Public id As String
        Public status As String
    End Class

    Public Class CreditCardTransaction
        Public id As String
        Public order_id As String
        Public kind As String
        Public gateway As String
        Public status As String
        Public message As String
        Public created_at As String
        Public test As String
        Public authorization As String
        Public amount As Decimal
        Public currency As String
        Public parent_id As String
        Public user_id As String
        Public location_id As String
        Public device_id As String
        Public error_code As String
        Public source_name As String
        Public payment_id As String
        Public processed_at As String
        Public receipt As New Receipt
    End Class

    Public Function CaptureAuthorizedCreditCard(ByVal ShopifyOrderID As String, ByVal amountToCapture As Decimal) As CreditCardTransaction

        ' You can’t run this until the order is in authorized state (card already approved at checkout).
        ' The authorization hold is usually 7 days (sometimes up to 30 depending on gateway). Capture must happen before it expires.
        ' You'll never touch raw credit card numbers — Shopify handles that part securely.

        Dim creditCardTransaction As New CreditCardTransaction
        LastError = String.Empty

        Try
            If rowECTECOMD Is Nothing Then
                LastError = "Ecommerce Partner {ECOM_CODE} does not have Shopify credentials"
                Return creditCardTransaction
            End If

            Dim ShopifyUrl As String = rowECTECOMD.Item("ECOM_URL") & String.Empty
            If ShopifyUrl.Length = 0 Then
                LastError = "Ecommerce Partner {ECOM_CODE} is not assigned a URL"
                Return creditCardTransaction
            End If
            Dim ECOM_SITE_USER As String = rowECTECOMD.Item("ECOM_SITE_USER") & String.Empty
            Dim ECOM_SITE_PWD As String = rowECTECOMD.Item("ECOM_SITE_PWD") & String.Empty
            ShopifyUrl = ShopifyUrl.Replace("{USER_ID}", ECOM_SITE_USER).Replace("{PASSWORD}", ECOM_SITE_PWD)

            Dim shopAccessToken As String = rowECTECOMD.Item("ECOM_SITE_SECRET_KEY") & String.Empty
            If shopAccessToken.Length = 0 Then
                LastError = $"Ecommerce Partner {ECOM_CODE} is not assigned an Access Token"
                Return creditCardTransaction
            End If

            If IsPaymentCaptured(ShopifyOrderID, shopAccessToken) Then
                creditCardTransaction.status = "SUCCESS"
                Return creditCardTransaction
            End If

            Dim url As String = $"{ShopifyUrl}orders/{ShopifyOrderID}/transactions.json"

            Dim jsonPayload As String = $"{{""transaction"": {{""kind"": ""capture"", ""amount"": ""{amountToCapture}""}}}}"

            Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            request.Method = "POST"
            request.Headers.Add("X-Shopify-Access-Token", shopAccessToken)
            request.ContentType = "application/json"

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(jsonPayload)
            request.ContentLength = byteArray.Length
            Using dataStream As IO.Stream = request.GetRequestStream()
                dataStream.Write(byteArray, 0, byteArray.Length)
            End Using

            Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Dim responseText As String = String.Empty
                Using reader As New IO.StreamReader(response.GetResponseStream())
                    responseText = reader.ReadToEnd()
                End Using

                Dim json As JObject = JObject.Parse(responseText)
                Dim transaction = json("transaction")

                If transaction IsNot Nothing Then
                    With creditCardTransaction
                        .id = transaction("id").ToString()
                        .order_id = transaction("order_id").ToString()
                        .kind = transaction("kind").ToString()
                        .gateway = transaction("gateway").ToString()
                        .status = transaction("status").ToString()
                        .message = transaction("message").ToString()
                        .created_at = transaction("created_at").ToString()
                        .test = transaction("test").ToString()
                        .authorization = transaction("authorization").ToString()
                        .amount = Val(transaction("amount").ToString())
                        .currency = transaction("currency").ToString()
                        .parent_id = transaction("parent_id").ToString()
                        .user_id = transaction("user_id").ToString()
                        .location_id = transaction("location_id").ToString()
                        .device_id = transaction("device_id").ToString()
                        .error_code = transaction("error_code").ToString()
                        .source_name = transaction("order_id").ToString()
                        .payment_id = transaction("payment_id").ToString()
                        .processed_at = transaction("processed_at").ToString()

                        Dim receipt As JObject = CType(transaction("receipt"), JObject)
                        If receipt IsNot Nothing Then
                            .receipt.id = receipt("id").ToString()
                            .receipt.status = receipt("status").ToString()
                        End If
                    End With
                End If

                response.Close()
            End Using

        Catch ex As Exception
            LastError = $"CaptureAuthorizedCreditCard Error: {ex.Message}"
            creditCardTransaction.status = "error"
            creditCardTransaction.error_code = "99999"
            creditCardTransaction.message = ex.Message
        End Try

        Return creditCardTransaction

    End Function

    Function IsPaymentCaptured(ShopifyOrderID As String, accessToken As String) As Boolean
        Try
            Dim url As String = $"{ShopifyUrl}orders/{ShopifyOrderID}/transactions.json"
            Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
            request.Method = "GET"
            request.Headers.Add("X-Shopify-Access-Token", accessToken)

            Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
                Using reader As New StreamReader(response.GetResponseStream())
                    Dim json As JObject = JObject.Parse(reader.ReadToEnd())
                    For Each txn In json("transactions")
                        If txn("kind").ToString() = "capture" AndAlso txn("status").ToString() = "success" Then
                            Return True
                        ElseIf txn("kind").ToString() = "sale" AndAlso txn("status").ToString() = "success" Then
                            Return True
                        End If
                    Next
                End Using
            End Using
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Is Payment Captured", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

End Class
