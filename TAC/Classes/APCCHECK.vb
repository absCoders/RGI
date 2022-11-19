Imports DPayments.InPay

Public Class APCCHECK

    Private tblAPTVEND1 As New DataTable
    Private tblGLTBANK1 As New DataTable
    Private eCheckResponse As New eCheckResponseClass
    Private processingTable As Boolean = False
    Private errorList As List(Of String)

    Public Sub New()
        tblAPTVEND1 = ASCDATA1.GetDataTable("SELECT * FROM APTVEND1 WHERE VEND_BANK_ROUTING_NO IS NOT NULL AND VEND_BANK_ACCT_ID IS NOT NULL", "APTVEND1")
        tblGLTBANK1 = ASCDATA1.GetDataTable("SELECT * FROM GLTBANK1", "GLTBANK1")
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

            Dim BANK_CODE As String = rowAPTCHCK1.Item("BANK_CODE") & String.Empty
            Dim rowGLTBANK1 As DataRow = tblGLTBANK1.Rows.Find(BANK_CODE)

            Dim epay As New Echeck
            epay.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
            Dim bankInfo As New EPBank(
                 routingNumber:=rowAPTVEND1.Item("VEND_BANK_ROUTING_NO"),
                 accountNumber:=rowAPTVEND1.Item("VEND_BANK_ACCT_ID"),
                 accountClass:=If(VEND_BANK_ACCT_CLASS = "B", AccountClass.acBusiness, AccountClass.acPersonal),
                 accountType:=If(VEND_BANK_ACCT_TYPE = "C", AccountTypes.atChecking, AccountTypes.atSavings))
            epay.Bank = bankInfo

            Dim Customer As New EPCustomer
            With Customer
                .FirstName = rowAPTVEND1.Item("VEND_CODE") & String.Empty
                .LastName = rowAPTVEND1.Item("VEND_NAME") & String.Empty
            End With
            epay.Customer = Customer

            'name:=rowAPTVEND1.Item("VEND_NAME"),
            'accountHolderName:=rowAPTVEND1.Item("VEND_NAME"))

            epay.MerchantLogin = rowGLTBANK1.Item("BANK_MERCHANT_ID") & String.Empty
            epay.MerchantPassword = rowGLTBANK1.Item("BANK_MERCHANT_PASSWORD") & String.Empty
            epay.CheckNumber = rowAPTCHCK1.Item("CHECK_NUM") & String.Empty
            epay.CompanyName = "ABS"
            epay.TransactionAmount = Format(Val(rowAPTCHCK1.Item("CHECK_AMT") & String.Empty), "#.00")
            epay.TransactionDesc = "Payment"
            Dim TRAN_NO As String = ASCMAIN1.Next_Control_No("APTCHCK1.TRAN_NO")
            epay.TransactionId = TRAN_NO
            epay.PaymentType = EcheckPaymentTypes.ptBOC
            epay.Gateway = EcheckGateways.ecgwForte

            If epay.CheckNumber & String.Empty <> String.Empty Then
                epay.InvoiceNumber = epay.CheckNumber
            Else
                epay.InvoiceNumber = "T" & TRAN_NO
            End If

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

        Dim BANK_CODE As String = rowAPTCHCK1.Item("BANK_CODE") & String.Empty
        Dim rowGLTBANK1 As DataRow = tblGLTBANK1.Rows.Find(BANK_CODE)

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

        If rowGLTBANK1 Is Nothing Then
            errorList.Add($"Bank {BANK_CODE} is invalid or missing banking information")
            Return errorList
        End If

        If rowGLTBANK1.Item("BANK_MERCHANT_ID") & String.Empty = String.Empty Then
            errorList.Add($"Bank {BANK_CODE} is missing the Bank Merchant ID")
        End If

        If rowGLTBANK1.Item("BANK_MERCHANT_PASSWORD") & String.Empty = String.Empty Then
            errorList.Add($"Bank {BANK_CODE} is missing the Bank Merchant Password")
        End If

        Return errorList
    End Function

End Class
