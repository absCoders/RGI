Public Class APCCHECK

    Private tblAPTVEND1 As New DataTable
    Private eCheckResponse As New eCheckResponseClass
    Private processingTable As Boolean = False
    Private errorList As List(Of String)

    Public Sub New()
        tblAPTVEND1 = ASCDATA1.GetDataTable("SELECT * FROM APTVEND1 WHERE VEND_BANK_ROUTING_NO IS NOT NULL AND VEND_BANK_ACCT_ID IS NOT NULL", "APTVEND1")
        errorList = New List(Of String)
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

            Dim VEND_BANK_ACCT_CLASS As String = rowAPTVEND1.Item("VEND_BANK_ACCT_CLASS") & ""
            Dim VEND_BANK_ACCT_TYPE As String = rowAPTVEND1.Item("VEND_BANK_ACCT_TYPE") & ""

            Dim errorList As New List(Of String)
            errorList = ValidateEntry(rowAPTCHCK1)
            If errorList.Count > 0 Then
                For Each errorMsg As String In errorList
                    errorMsg = errorMsg.Trim
                    If errorMsg.Length = 0 Then Continue For
                    eCheckResponse.ErrorMessage = Environment.NewLine & errorMsg
                Next
            End If

            If eCheckResponse.ErrorMessage.Length > 0 Then
                Return False
            End If

            Dim epay As New nsoftware.InPay.Echeck
            epay.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareInPay")
            Dim bankInfo As New nsoftware.InPay.EPBank(
                 routingNumber:=rowAPTVEND1.Item("VEND_BANK_ROUTING_NO"),
                 accountNumber:=rowAPTVEND1.Item("VEND_BANK_ACCT_ID"),
                 accountClass:=If(VEND_BANK_ACCT_CLASS = "B", nsoftware.InPay.AccountClass.acBusiness, nsoftware.InPay.AccountClass.acPersonal),
                 accountType:=If(VEND_BANK_ACCT_TYPE = "C", nsoftware.InPay.AccountTypes.atChecking, nsoftware.InPay.AccountTypes.atSavings),
                 name:=rowAPTVEND1.Item("VEND_NAME"),
                 accountHolderName:=rowAPTVEND1.Item("VEND_NAME"))
            epay.Bank = bankInfo

            epay.MerchantLogin = "213079" 'EpFI1F ' txtLogin.Text
            epay.MerchantPassword = "0ff1c3ABS$*+" ' txtPassword.Text
            ' epay.GatewayURL = ""
            epay.CheckNumber = rowAPTCHCK1.Item("CHECK_NUM") & String.Empty
            epay.CompanyName = "ABS"
            epay.TransactionAmount = Format(Val(rowAPTCHCK1.Item("CHECK_AMT") & String.Empty), "#.00")
            epay.TransactionDesc = "payment"
            Dim TRAN_NO As String = ASCMAIN1.Next_Control_No("APTCHCK1.TRAN_NO")
            epay.TransactionId = TRAN_NO
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

    Public Function ValidateEntry(ByRef tblAPTCHCK1 As DataTable) As List(Of String)
        processingTable = True
        errorList = New List(Of String)

        Dim tbl As DataTable = ASCDATA1.SelectDistinct(tblAPTCHCK1, New String() {"VEND_CODE"})

        For Each row As DataRow In tbl.Select("", "VEND_CODE")
            Dim rowAPTCHCK1 As DataRow = tblAPTCHCK1.Select($"VEND_CODE = '{row.Item("VEND_CODE")}'")(0)
            ValidateEntry(rowAPTCHCK1)
        Next

        processingTable = False
        Return errorList

    End Function

    Public Function ValidateEntry(ByRef rowAPTCHCK1 As DataRow) As List(Of String)
        If Not processingTable Then
            errorList = New List(Of String)
        End If

        Dim VEND_CODE As String = rowAPTCHCK1.Item("VEND_CODE") & String.Empty
        Dim rowAPTVEND1 As DataRow = tblAPTVEND1.Rows.Find(VEND_CODE)

        If rowAPTVEND1 Is Nothing Then
            errorList.Add($"Vendor {VEND_CODE} is invalid or missing banking information")
            Return errorList
        End If

        If rowAPTVEND1.Item("VEND_BANK_ROUTING_NO") & String.Empty = String.Empty Then
            errorList.Add($"Vendor {VEND_CODE} is missing the Bank Account Routing No")
        End If

        If rowAPTVEND1.Item("VEND_BANK_ACCT_ID") & String.Empty = String.Empty Then
            errorList.Add($"Vendor {VEND_CODE} is missing the Bank Account No")
        End If

        If rowAPTVEND1.Item("VEND_BANK_ACCT_CLASS") & String.Empty = String.Empty Then
            errorList.Add($"Vendor {VEND_CODE} is missing the Bank Account Class")
        End If

        If rowAPTVEND1.Item("VEND_BANK_ACCT_TYPE") & String.Empty = String.Empty Then
            errorList.Add($"Vendor {VEND_CODE} is missing the Bank Account Type")
        End If

        Return errorList
    End Function

End Class
