Public Class APCCHECK

    Private tblAPTVEND1 As New DataTable

    Public Sub New()
        tblAPTVEND1 = ASCDATA1.GetDataTable("SELECT * FROM APTVEND1 WHERE VEND_BANK_ROUTING_NO IS NOT NULL AND VEND_BANK_ACCT_ID IS NOT NULL", "APTVEND1")
    End Sub

    Public Enum eCheckTypes
        Authorize
        Credit
    End Enum

    Private Class eCheckResponseClass
        Public ApprovalCode As String = String.Empty
        Public ResponseCode As String = String.Empty
        Public ResponseInvoice As String = String.Empty
        Public ResponseText As String = String.Empty
        Public ResponseTransid As String = String.Empty
        Public ErrorMessage As String = String.Empty
    End Class

    Private eCheckResponse As New eCheckResponseClass

    Public ReadOnly Property CheckResponse
        Get
            Return eCheckResponse
        End Get
    End Property

    Public Function Send_eChecks(ByRef rowAPTCHCK1 As DataRow, ByVal method As eCheckTypes) As Boolean

        Send_eChecks = False

        Try
            eCheckResponse = New eCheckResponseClass

            If Val(rowAPTCHCK1.Item("CHECK_AMT") & String.Empty) = 0 Then
                Return True
            End If

            Dim VEND_CODE As String = rowAPTCHCK1.Item("VEND_CODE") & String.Empty
            Dim rowAPTVEND1 As DataRow = tblAPTVEND1.Rows.Find(VEND_CODE)

            Dim epay As New nsoftware.InPay.Echeck
            Dim b As New nsoftware.InPay.EPBank(
                 routingNumber:=rowAPTVEND1.Item("VEND_BANK_ROUTING_NO"),
                 accountNumber:=rowAPTVEND1.Item("VEND_BANK_ACCT_ID"),
                 accountClass:=nsoftware.InPay.AccountClass.acPersonal,
                 accountType:=nsoftware.InPay.AccountTypes.atChecking,
                 name:=rowAPTVEND1.Item("VEND_NAME"),
                 accountHolderName:=rowAPTVEND1.Item("VEND_NAME"))
            epay.Bank = b

            epay.MerchantLogin = "Login" ' txtLogin.Text
            epay.MerchantPassword = "Password" ' txtPassword.Text
            ' epay.GatewayURL = ""
            epay.CheckNumber = rowAPTCHCK1.Item("CHECK_NUM")
            epay.CompanyName = "ABS"
            epay.TransactionAmount = Format(Val(rowAPTCHCK1.Item("CHECK_AMT") & String.Empty), "#.00")
            epay.TransactionDesc = "payment"
            epay.TransactionId = "123"
            epay.PaymentType = nsoftware.InPay.EcheckPaymentTypes.ptBOC
            epay.Gateway = nsoftware.InPay.EcheckGateways.gwACHPayments

            Select Case method
                Case eCheckTypes.Authorize
                    epay.Authorize()
                Case eCheckTypes.Credit
                    epay.Credit("", epay.TransactionAmount)
            End Select

            With eCheckResponse
                .ApprovalCode = epay.Response.ApprovalCode
                .ResponseCode = epay.Response.Code
                .ResponseInvoice = epay.Response.InvoiceNumber
                .ResponseText = epay.Response.Text
                .ResponseTransid = epay.Response.TransactionId
            End With

            Send_eChecks = True

        Catch ex As Exception
            eCheckResponse.ErrorMessage = ex.Message
        End Try

    End Function

End Class
