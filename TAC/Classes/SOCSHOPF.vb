Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
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
        Initialize()
    End Sub

    Private Sub Initialize()
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

            ShopifyUrl = rowECTECOMD.Item("ECOM_URL") & String.Empty
            If ShopifyUrl.Length = 0 Then
                LastError = "Ecommerce Partner {ECOM_CODE} is not assigned a URL"
                Return creditCardTransaction
            End If
            ECOM_SITE_USER = rowECTECOMD.Item("ECOM_SITE_USER") & String.Empty
            ECOM_SITE_PWD = rowECTECOMD.Item("ECOM_SITE_PWD") & String.Empty
            ShopifyUrl = ShopifyUrl.Replace("{USER_ID}", ECOM_SITE_USER).Replace("{PASSWORD}", ECOM_SITE_PWD)

            shopAccessToken = rowECTECOMD.Item("ECOM_SITE_SECRET_KEY") & String.Empty
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
                            '.receipt.id = receipt("id").ToString()
                            '.receipt.status = receipt("status").ToString()
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

            Return False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Is Payment Captured", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Public Function GetShopifyProducts(ByRef NumItemsUpdated As Int16) As Boolean

        Try
            NumItemsUpdated = 0
            Dim sql As String = String.Empty
            Dim lstDiscontinuedItems As New List(Of String)

            Initialize()

            If rowECTECOMD Is Nothing Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} does not have Shopify credentials")
                Return False
            End If

            If ECOM_CUST_CODE.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned to a Customer")
                Return False
            End If

            ShopifyUrl = rowECTECOMD.Item("ECOM_URL") & String.Empty
            If ShopifyUrl.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned a URL")
                Return False
            End If

            ECOM_SITE_USER = rowECTECOMD.Item("ECOM_SITE_USER") & String.Empty
            ECOM_SITE_PWD = rowECTECOMD.Item("ECOM_SITE_PWD") & String.Empty
            ShopifyUrl = ShopifyUrl.Replace("{USER_ID}", ECOM_SITE_USER).Replace("{PASSWORD}", ECOM_SITE_PWD)

            shopAccessToken = rowECTECOMD.Item("ECOM_SITE_SECRET_KEY") & String.Empty
            If shopAccessToken.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned an Access Token")
                Return False
            End If

            Dim client As New HttpClient() With {
                .BaseAddress = New Uri(ShopifyUrl)
            }

            ' Authentication
            client.DefaultRequestHeaders.Accept.Clear()
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            client.DefaultRequestHeaders.Add("X-Shopify-Access-Token", shopAccessToken)

            Dim response As New HttpResponseMessage
            Dim nextPageUrl As String = "products.json?limit=100&published_status=any"
            Dim counter As Int32 = 0

            While Not String.IsNullOrEmpty(nextPageUrl)
                response = client.GetAsync(nextPageUrl).Result
                response.EnsureSuccessStatusCode()
                System.Threading.Thread.Sleep(3000)

                If response.IsSuccessStatusCode Then
                    Dim responseStr As String = response.Content.ReadAsStringAsync().Result
                    Dim jsonObj As JObject = JObject.Parse(responseStr)
                    counter += jsonObj("products").Count

                    For Each product As JObject In jsonObj("products")
                        Dim ECOM_PRODUCT_ID As String = product("id").ToString
                        Dim WEB_DESCRIPTION As String = product("title").ToString
                        Dim BODY_HTML As String = product("body_html").ToString

                        For Each jVariant As JObject In product("variants")
                            Dim UPC_CODE As String = jVariant("sku").ToString
                            Dim ECOM_VARIANT_ID As String = jVariant("id").ToString
                            Dim ECOM_INV_VARIANT_ID As String = jVariant("inventory_item_id").ToString
                            ASCMAIN1.Progress("-", UPC_CODE)

                            If ECOM_PRODUCT_ID.Length = 0 OrElse UPC_CODE.Length = 0 OrElse ECOM_VARIANT_ID.Length = 0 OrElse ECOM_INV_VARIANT_ID.Length = 0 Then
                                'Stop
                            Else
                                sql = "UPDATE ICTSTYCW SET ECOM_PRODUCT_ID = :PARM1, ECOM_VARIANT_ID = :PARM2, ECOM_INV_VARIANT_ID = :PARM3, ECOM_PRODUCT_LAST_UPDATED = SYSDATE,
                                                    WEB_DESCRIPTION = :PARM4, BODY_HTML = :PARM5
                                                    WHERE (STYLE_CODE, COLOR_CODE, SIZE_INDEX) IN
                                                    (SELECT STYLE_CODE, COLOR_CODE, SIZE_INDEX FROM ICTSTYC4 WHERE UPC_CODE = :PARM6)
                                                    AND (ECOM_PRODUCT_ID IS NULL OR ECOM_VARIANT_ID IS NULL OR ECOM_INV_VARIANT_ID IS NULL)"
                                NumItemsUpdated += ASCDATA1.ExecuteSQL(sql, "VVVVVV", {ECOM_PRODUCT_ID, ECOM_VARIANT_ID, ECOM_INV_VARIANT_ID, WEB_DESCRIPTION, BODY_HTML, UPC_CODE})
                            End If
                        Next
                    Next
                End If

                nextPageUrl = Nothing ' reset first
                If response.Headers.Contains("Link") Then
                    Dim linkHeader = response.Headers.GetValues("Link").FirstOrDefault()
                    If linkHeader IsNot Nothing Then
                        Dim parts() As String = linkHeader.Split(","c)

                        For Each part As String In parts
                            If part.Contains("rel=""next""") Then
                                Dim start = part.IndexOf("<") + 1
                                Dim [end] = part.IndexOf(">")
                                Dim fullUrl = part.Substring(start, [end] - start)

                                ' Shopify requires using the exact URL (absolute), not just relative
                                nextPageUrl = fullUrl
                            End If
                        Next
                    End If
                End If

                ' If we found an absolute URL, strip base only if needed
                If Not String.IsNullOrEmpty(nextPageUrl) AndAlso nextPageUrl.StartsWith(client.BaseAddress.ToString()) Then
                    'nextPageUrl = nextPageUrl.Replace(client.BaseAddress.ToString(), "")
                    nextPageUrl = nextPageUrl.Replace(client.BaseAddress.ToString(), "")
                End If
            End While

            Return True

        Catch ex As Exception
            MessageBox.Show($"GetShopifyProducts Error: {ex.Message}")
            Return False
        End Try

    End Function

    Public Function GetShopifyProductsGraphQL(ByRef NumItemsUpdated As Int16) As Boolean

        Try
            NumItemsUpdated = 0
            Dim sql As String = String.Empty
            Dim lstDiscontinuedItems As New List(Of String)

            Initialize()

            If rowECTECOMD Is Nothing Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} does not have Shopify credentials")
                Return False
            End If

            If ECOM_CUST_CODE.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned to a Customer")
                Return False
            End If

            ShopifyUrl = rowECTECOMD.Item("ECOM_URL") & String.Empty
            If ShopifyUrl.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned a URL")
                Return False
            End If

            ECOM_SITE_USER = rowECTECOMD.Item("ECOM_SITE_USER") & String.Empty
            ECOM_SITE_PWD = rowECTECOMD.Item("ECOM_SITE_PWD") & String.Empty
            ShopifyUrl = ShopifyUrl.Replace("{USER_ID}", ECOM_SITE_USER).Replace("{PASSWORD}", ECOM_SITE_PWD)

            shopAccessToken = rowECTECOMD.Item("ECOM_SITE_SECRET_KEY") & String.Empty
            If shopAccessToken.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned an Access Token")
                Return False
            End If

            sql = "SELECT ICTSTYC4.UPC_CODE, ICTSTYCW.*
                        FROM ICTSTYCW, ICTSTYC4
                        WHERE ICTSTYCW.STYLE_CODE = ICTSTYC4.STYLE_CODE
                        AND ICTSTYCW.COLOR_CODE = ICTSTYC4.COLOR_CODE
                        AND ICTSTYCW.SIZE_INDEX = ICTSTYC4.SIZE_INDEX
                        AND (ICTSTYCW.ECOM_PRODUCT_ID IS NULL AND ICTSTYCW.ECOM_VARIANT_ID IS NULL AND ICTSTYCW.ECOM_INV_VARIANT_ID IS NULL)"

            Dim tblICTSTYCW As DataTable = ASCDATA1.GetDataTable(sql)


            Dim hasNext As Boolean = True
            Dim cursor As String = Nothing
            Dim itemCount As Int32 = 0

            While hasNext
                Dim gql As String = $"{{ 
                                        products(first: 250, query: ""status:active OR status:draft OR status:archived"" {If(cursor IsNot Nothing, $", after: ""{cursor}""", "")}) {{
                                            nodes {{
                                                id
                                                title
                                                bodyHtml
                                                variants(first: 250) {{
                                                    nodes {{
                                                        id
                                                        sku
                                                        inventoryItem {{
                                                            id
                                                        }}
                                                    }}
                                                }}
                                            }}
                                            pageInfo {{
                                                hasNextPage
                                                endCursor
                                            }}
                                        }}
                                    }}"

                Dim result = PostGraphQL(gql)
                Dim data = result("data")("products")

                itemCount += data("nodes").Count

                For Each product In data("nodes")
                    Dim ECOM_PRODUCT_ID As String = product("id").ToString
                    Dim WEB_DESCRIPTION As String = product("title").ToString
                    Dim BODY_HTML As String = product("bodyHtml").ToString

                    For Each jVariant In product("variants")("nodes")
                        Dim UPC_CODE As String = jVariant("sku").ToString
                        Dim ECOM_VARIANT_ID As String = jVariant("id").ToString.Replace("gid://shopify/ProductVariant/", "")
                        Dim ECOM_INV_VARIANT_ID As String = jVariant("inventoryItem")("id").ToString.Replace("gid://shopify/InventoryItem/", "")
                        ASCMAIN1.Progress("-", UPC_CODE)

                        If tblICTSTYCW.Select($"UPC_CODE = '{UPC_CODE}'").Length > 0 Then
                            sql = "UPDATE ICTSTYCW SET ECOM_PRODUCT_ID = :PARM1, ECOM_VARIANT_ID = :PARM2, ECOM_INV_VARIANT_ID = :PARM3, ECOM_PRODUCT_LAST_UPDATED = SYSDATE,
                                                    WEB_DESCRIPTION = :PARM4, BODY_HTML = :PARM5
                                                    WHERE (STYLE_CODE, COLOR_CODE, SIZE_INDEX) IN
                                                    (SELECT STYLE_CODE, COLOR_CODE, SIZE_INDEX FROM ICTSTYC4 WHERE UPC_CODE = :PARM6)
                                                    AND (ECOM_PRODUCT_ID IS NULL AND ECOM_VARIANT_ID IS NULL OR ECOM_INV_VARIANT_ID AND NULL)"
                            NumItemsUpdated += ASCDATA1.ExecuteSQL(sql, "VVVVVV", {ECOM_PRODUCT_ID, ECOM_VARIANT_ID, ECOM_INV_VARIANT_ID, WEB_DESCRIPTION, BODY_HTML, UPC_CODE})
                        Else
                            Dim tblICTSTYC4 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ICTSTYC4 WHERE UPC_CODE = :PARM1 AND STYLE_CODE NOT IN (SELECT STYLE_CODE FROM ICTSTYCW)", "", "V", {UPC_CODE})
                            If tblICTSTYC4.Rows.Count > 0 Then
                                Stop
                            End If
                        End If
                    Next
                Next

                hasNext = data("pageInfo")("hasNextPage").ToObject(Of Boolean)()
                cursor = If(hasNext, data("pageInfo")("endCursor").ToString(), Nothing)
            End While

            Return True
        Catch ex As Exception
            MessageBox.Show($"GetShopifyProductsGraphQL Error: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function PostGraphQL(query As String) As JObject

        Using client As New HttpClient
            client.DefaultRequestHeaders.Accept.Clear()
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            client.DefaultRequestHeaders.Add("X-Shopify-Access-Token", shopAccessToken)

            Dim payload = New JObject() From {
                {"query", query}
            }

            Dim content = New StringContent(payload.ToString(), Encoding.UTF8, "application/json")
            Dim response = client.PostAsync(ShopifyUrl & "graphql.json", content).Result
            Dim json = response.Content.ReadAsStringAsync().Result
            Return JObject.Parse(json)
        End Using

    End Function

    Private Function GetShopifyPayouts(ByVal startDate As Date, ByVal endDate As Date) As DataSet

        Try
            Initialize()

            If rowECTECOMD Is Nothing Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} does not have Shopify credentials")
                Return Nothing
            End If

            If ECOM_CUST_CODE.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned to a Customer")
                Return Nothing
            End If

            ShopifyUrl = rowECTECOMD.Item("ECOM_URL") & String.Empty
            If ShopifyUrl.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned a URL")
                Return Nothing
            End If

            ECOM_SITE_USER = rowECTECOMD.Item("ECOM_SITE_USER") & String.Empty
            ECOM_SITE_PWD = rowECTECOMD.Item("ECOM_SITE_PWD") & String.Empty
            ShopifyUrl = ShopifyUrl.Replace("{USER_ID}", ECOM_SITE_USER).Replace("{PASSWORD}", ECOM_SITE_PWD)
            ' https://{USER_ID}:{PASSWORD}@skinlingerie.myshopify.com/admin/api/2025-04/
            ' $"/admin/api/2024-01/shopify_payments/payouts.json?limit=250&date_min={dateMin}&date_max={dateMax}"

            shopAccessToken = rowECTECOMD.Item("ECOM_SITE_SECRET_KEY") & String.Empty
            If shopAccessToken.Length = 0 Then
                MessageBox.Show($"UpdateProducts, Ecommerce Partner {ECOM_CODE} is not assigned an Access Token")
                Return Nothing
            End If

            Dim dtPayment As DataTable = GetPayoutsTable(startDate, endDate)
            Dim dtTransactions As DataTable = GetTransactionsTable(dtPayment)

            Dim dsPayments As New DataSet
            dsPayments.Tables.Add(dtPayment)
            dsPayments.Tables.Add(dtTransactions)

            Return dsPayments

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Get Shopify Payouts", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

    End Function

    Private Function GetPayoutsTable(startDate As Date, endDate As Date) As DataTable

        Try
            Dim dtPayment As New DataTable
            With dtPayment
                .Columns.Add("PAYOUT_ID", GetType(String))
                .Columns.Add("PAYOUT_STATUS", GetType(String))
                .Columns.Add("DATE", GetType(Date))
                .Columns.Add("CURRENCY", GetType(String))
                .Columns.Add("AMOUNT", GetType(Decimal))

                .Columns.Add("ADJUSTMENTS_FEE_AMOUNT", GetType(Decimal))
                .Columns.Add("ADJUSTMENTS_GROSS_AMOUNT", GetType(Decimal))
                .Columns.Add("CHARGES_FEE_AMOUNT", GetType(Decimal))
                .Columns.Add("CHARGES_GROSS_AMOUNT", GetType(Decimal))
                .Columns.Add("REFUNDS_FEE_AMOUNT", GetType(Decimal))
                .Columns.Add("REFUNDS_GROSS_AMOUNT", GetType(Decimal))
                .Columns.Add("RESERVED_FUNDS_FEE_AMOUNT", GetType(Decimal))
                .Columns.Add("RESERVED_FUNDS_GROSS_AMOUNT", GetType(Decimal))
                .Columns.Add("RETRIED_PAYOUTS_FEE_AMOUNT", GetType(Decimal))
                .Columns.Add("RETRIED_PAYOUTS_GROSS_AMOUNT", GetType(Decimal))
            End With
            dtPayment.TableName = "PAYMENTS"

            Using client As New HttpClient() With {
                .BaseAddress = New Uri(ShopifyUrl)
                }

                client.DefaultRequestHeaders.Add("X-Shopify-Access-Token", shopAccessToken)

                Dim currentStart As Date = startDate

                While currentStart < endDate
                    Dim dateMin As String = currentStart.ToString("yyyy-MM-dd")
                    Dim response As New HttpResponseMessage

                    ' GET /admin/api/2025-04/shopify_payments/payouts.json?date_min=YYYY-MM-DD&date_max=YYYY-MM-DD
                    Dim endPoint As String = ShopifyUrl & $"shopify_payments/payouts.json?limit=250&date_min={dateMin}&date_max={dateMin}&status=paid"

                    response = client.GetAsync(endPoint).Result
                    Dim json As String = response.Content.ReadAsStringAsync().Result

                    Dim parsed = JObject.Parse(json)
                    Dim payouts As JArray = parsed("payouts")

                    For Each p In payouts
                        Dim drPayment As DataRow = dtPayment.NewRow
                        With drPayment
                            .Item("PAYOUT_ID") = p("id").ToString()
                            .Item("PAYOUT_STATUS") = p("status").ToString()
                            .Item("DATE") = CDate(p("date").ToString() & String.Empty)
                            .Item("CURRENCY") = p("currency").ToString()
                            .Item("AMOUNT") = Val(p("amount").ToString() & String.Empty)
                        End With
                        dtPayment.Rows.Add(drPayment)

                        Try
                            If p("summary") IsNot Nothing Then
                                Dim summary = p("summary")
                                drPayment.Item("ADJUSTMENTS_FEE_AMOUNT") = Val(summary("adjustments_fee_amount").ToString())
                                drPayment.Item("ADJUSTMENTS_GROSS_AMOUNT") = Val(summary("adjustments_gross_amount").ToString())
                                drPayment.Item("CHARGES_FEE_AMOUNT") = Val(summary("charges_fee_amount").ToString())
                                drPayment.Item("CHARGES_GROSS_AMOUNT") = Val(summary("charges_gross_amount").ToString())
                                drPayment.Item("REFUNDS_FEE_AMOUNT") = Val(summary("refunds_fee_amount").ToString())
                                drPayment.Item("REFUNDS_GROSS_AMOUNT") = Val(summary("refunds_gross_amount").ToString())
                                drPayment.Item("RESERVED_FUNDS_FEE_AMOUNT") = Val(summary("reserved_funds_fee_amount").ToString())
                                drPayment.Item("RESERVED_FUNDS_GROSS_AMOUNT") = Val(summary("reserved_funds_gross_amount").ToString())
                                drPayment.Item("RETRIED_PAYOUTS_FEE_AMOUNT") = Val(summary("retried_payouts_fee_amount").ToString())
                                drPayment.Item("RETRIED_PAYOUTS_GROSS_AMOUNT") = Val(summary("retried_payouts_gross_amount").ToString())
                            End If
                        Catch ex As Exception

                        End Try
                    Next
                    currentStart = currentStart.AddDays(1)
                End While

            End Using

            Return dtPayment

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Function GetTransactionsTable(dtPayment As DataTable) As DataTable
        Try
            Dim dtTransaction As New DataTable
            With dtTransaction
                .Columns.Add("TRANS_ID", GetType(String))
                .Columns.Add("TRANS_TYPE", GetType(String))
                .Columns.Add("PAYOUT_ID", GetType(String))
                .Columns.Add("PAYOUT_STATUS", GetType(String))
                .Columns.Add("CURRENCY", GetType(String))
                .Columns.Add("AMOUNT", GetType(Decimal))
                .Columns.Add("FEE", GetType(Decimal))
                .Columns.Add("NET", GetType(Decimal))
                .Columns.Add("ORDR_WEB_ID", GetType(Decimal))
                .Columns.Add("DATE_PROCESED", GetType(Date))
                .Columns.Add("INV_NO", GetType(String))
            End With
            dtTransaction.TableName = "TRANSACTIONS"

            'Summary Table
            'Type	    Meaning	                                    Affects Invoice?    Notes
            'payment	Customer payment	                        ✅ Yes	            What you apply to invoice
            'refund	    Refund to customer	                        ❌ No	            Reduces payout
            'adjustment	Shopify manual adjustments	                ❌ No	            Usually not order-specific
            'dispute	Chargeback/dispute resolution	            ❌ Maybe	        Can affect payout for the order
            'credit	    Shopify credit (refunds, fees returned)	    ❌ No	            Usually special cases
            'payout	    Payout summary	                            ❌ No	            Total batch amount

            For Each drPayment As DataRow In dtPayment.Select("")
                Dim payoutId As String = drPayment.Item("PAYOUT_ID") & String.Empty

                ' GET /admin/api/2025-04/shopify_payments/payouts/{payout_id}/transactions.json
                Dim endpoint As String = ShopifyUrl & $"shopify_payments/payouts/{payoutId}/transactions.json"

                Using client As New HttpClient()
                    client.BaseAddress = New Uri(ShopifyUrl)
                    client.DefaultRequestHeaders.Add("X-Shopify-Access-Token", shopAccessToken)

                    Dim response = client.GetAsync(endpoint).Result
                    Dim json As String = response.Content.ReadAsStringAsync().Result

                    Dim parsed = JObject.Parse(json)
                    Dim transactions As JArray = parsed("transactions")

                    For Each t In transactions

                        Dim drTransaction As DataRow = dtTransaction.NewRow
                        With drTransaction
                            .Item("TRANS_ID") = t("id").ToString()
                            .Item("TRANS_TYPE") = t("type").ToString()
                            .Item("PAYOUT_ID") = t("payout_id").ToString()
                            .Item("PAYOUT_STATUS") = t("payout_status").ToString()
                            .Item("CURRENCY") = t("currency").ToString()
                            .Item("AMOUNT") = Val(t("amount").ToString() & String.Empty)
                            .Item("FEE") = Val(t("fee").ToString() & String.Empty)
                            .Item("NET") = Val(t("net").ToString() & String.Empty)

                            If t("source_order_id") IsNot Nothing AndAlso t("source_order_id").ToString() <> "" Then
                                .Item("ORDR_WEB_ID") = t("source_order_id").ToString()
                            ElseIf t("order_id") IsNot Nothing AndAlso t("order_id").ToString() <> "" Then
                                '.Item("ORDR_WEB_ID") = t("order_id").ToString()
                            ElseIf t("source_id") IsNot Nothing AndAlso IsNumeric(t("source_id").ToString()) Then
                                '.Item("ORDR_WEB_ID") = t("source_id").ToString()
                            End If

                            If t("processed_at") IsNot Nothing AndAlso t("processed_at").ToString() <> "" Then
                                If IsDate(t("processed_at").ToString()) Then
                                    .Item("DATE_PROCESED") = CDate(t("processed_at").ToString())
                                End If
                            End If

                            Dim ORDR_WEB_ID As String = .Item("ORDR_WEB_ID") & String.Empty
                            Dim INV_TOTAL_AMOUNT As Decimal = VAL(.Item("AMOUNT") & String.EMPTY)

                            If ORDR_WEB_ID.Length > 0 Then
                                Dim sql As String = "SELECT SOTINVH1.INV_NO, SOTINVH1.INV_TOTAL_AMOUNT
                                                            FROM SOTORDR1, SOTINVH1
                                                            WHERE SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO
                                                            AND SOTORDR1.ORDR_WEB_ID = :PARM1"
                                Dim drSOTINVH1 As DataRow = ASCDATA1.GetDataRow(sql, "V", ORDR_WEB_ID)
                                If drSOTINVH1 IsNot Nothing AndAlso drSOTINVH1.Item("INV_TOTAL_AMOUNT") = INV_TOTAL_AMOUNT Then
                                    .Item("INV_NO") = drSOTINVH1.Item("INV_NO") & String.Empty
                                Else
                                    ' See if this is a refund invoice
                                    ' --- SOTRTNL1, ORDR_NO, INV_NO
                                    sql = ""
                                End If
                            End If

                        End With
                        dtTransaction.Rows.Add(drTransaction)
                    Next

                End Using
            Next
            Return dtTransaction

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Class
