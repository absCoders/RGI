Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Http.Formatting
Imports Newtonsoft.Json
Imports Newtonsoft
Imports System.Xml

Public Class SOFORDCC
    Private FF As ASFBASE1
    Public CCProcessed As Boolean = False
    Public CUST_CODE As String = ""

    Public Sub New(ByVal frmASFBASE1 As ASFBASE1, ByVal in_CUST_CODE As String)
        FF = frmASFBASE1
        CUST_CODE = in_CUST_CODE
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If CUST_CODE.Length = 0 Then
            Me.Close()
        End If
        grdSOTORDRS.DataSource = FF.dst.Tables("SOTORDR1")
        CalculateTotal()
        SetGroupLabel()
        FillExpDate()
        SplitContainer1.SplitterDistance = SplitContainer1.Parent.Height - btnCustAdd.Height - 12
    End Sub

    Private Sub cmdFinished_Click(sender As System.Object, e As System.EventArgs) Handles cmdFinished.Click
        Dim Msg As String = ""
        Msg = VerifyCC()
        If Msg.Length > 0 Then
            MsgBox(Msg, MsgBoxStyle.Critical, "Credit Card Issues")
            Exit Sub
        Else
            MsgBox("Credit Card Recorded For Order(s)", MsgBoxStyle.Information, "Credit Card Posted")
            CCProcessed = True
            Me.Close()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        CCProcessed = False
        Me.Close()
    End Sub

    Private Sub CalculateTotal()
        Dim OTotal As Double = 0
        For Each rowSOTORDR1 As DataRow In FF.dst.Tables("SOTORDR1").Select()
            OTotal += Val(CDbl(Val(rowSOTORDR1.Item("TORDR") & "")))
        Next
        txtCCPA_AMT.Text = Format(OTotal, "#,###,##0.00")
    End Sub

    Private Sub SetGroupLabel()
        Dim ORDR_GROUP_NO As String = ""
        For Each rowSOTORDR1 As DataRow In FF.dst.Tables("SOTORDR1").Select()
            ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO") & ""
            If ORDR_GROUP_NO = "" Then
                ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_BATCH_NO") & ""
            End If
        Next
        grpSOTORDRS.Text = "Orders On Group " & ORDR_GROUP_NO
    End Sub

    Private Sub txtCCNO_TextChanged(sender As Object, e As System.EventArgs) Handles txtCUST_CREDIT_CARD_NO.TextChanged
        If txtCUST_CREDIT_CARD_NO.Text.Length > 4 Then
            txtLAST4.Text = txtCUST_CREDIT_CARD_NO.Text.Substring((txtCUST_CREDIT_CARD_NO.Text.Length - 4), 4)
        Else
            txtLAST4.Text = ""
        End If
    End Sub

    Private Sub txtShowNumbers_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles txtShowNumbers.CheckedChanged
        If txtShowNumbers.Checked = True Then
            txtCUST_CREDIT_CARD_NO.PasswordChar = ""
        Else
            txtCUST_CREDIT_CARD_NO.PasswordChar = "*"
        End If

    End Sub

    Private Function VerifyCC() As String
        Dim iResult As String = ""
        Dim eMsg As String = ""
        eMsg = VerifyControls()
        If eMsg.Length = 0 Then
            For Each rowSOTORDR1 As DataRow In FF.dst.Tables("SOTORDR1").Select()
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & ""
                Dim iResult_temp As String = UpdateARTCCPA1(ORDR_NO)
                If iResult_temp.Length > 0 Then
                    iResult = String.Format("{0}Order {1}: {2}", vbCrLf, ORDR_NO, iResult_temp)
                End If
            Next
        Else
            iResult = eMsg
        End If
        Return iResult
    End Function

#Region "Remote Service Calls"
    Function UpdateARTCCPA1(ByVal ORDR_NO As String) As String
        Dim iResult As String = ""
        Dim authController As String = ""
        Dim API_BASE As String = ""
        'authController = "AuthorizeCard"
        authController = "AuthorizeCardSSL"
        Dim API_CONTROLLER As String = "api/SalesOrder/" & authController
        Dim url As New System.Uri("https://api.regency-rib.com:8182/")
        'Dim url As New System.Uri("http://localhost:4055/")
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
        Dim resptest As System.Net.WebResponse
        req.Timeout = 20000 '20 Seconds
        Try
            resptest = req.GetResponse()
            resptest.Close()
            req = Nothing
            API_BASE = "https://api.regency-rib.com:8182/"
            'API_BASE = "http://localhost:4055/"
        Catch ex As Exception
            req = Nothing
            'API_BASE = "https://192.168.110.224:8182/"
            API_BASE = "https://api.regency-rib.com:8182/"
        End Try

        lblAuth.Text = "Attempting to Authorize..."
        lblAuth.Visible = True

        Dim _ARTCCPA1 As New ARTCCPA1
        If FF.dst.Tables("SOTORDR1").Rows.Count > 0 Then
            _ARTCCPA1.CUST_CODE = FF.dst.Tables("SOTORDR1").Rows(0).Item("CUST_CODE").ToString
        End If
        _ARTCCPA1.CUST_CREDIT_CARD_NAME = txtCUST_CREDIT_CARD_NAME.Text
        _ARTCCPA1.CUST_CREDIT_CARD_ADDR1 = txtCUST_CREDIT_CARD_ADDR1.Text
        _ARTCCPA1.CUST_CREDIT_CARD_CITY = txtCUST_CREDIT_CARD_CITY.Text
        _ARTCCPA1.CUST_CREDIT_CARD_STATE = txtCUST_CREDIT_CARD_STATE.Text
        _ARTCCPA1.CUST_CREDIT_CARD_ZIP_CODE = txtCUST_CREDIT_CARD_ZIP_CODE.Text
        _ARTCCPA1.CUST_CREDIT_CARD_COUNTRY = "US"
        _ARTCCPA1.ORDR_NO = ORDR_NO
        _ARTCCPA1.INIT_OPER = ASCMAIN1.USER_ID

        _ARTCCPA1.CUST_CREDIT_CARD_NO = txtCUST_CREDIT_CARD_NO.Text
        _ARTCCPA1.CUST_CREDIT_CARD_EXP_DATE = cboCUST_CREDIT_CARD_EXP_DATE.Text.Substring(0, 2) & cboCUST_CREDIT_CARD_EXP_DATE.Text.Substring(cboCUST_CREDIT_CARD_EXP_DATE.Text.Length - 2, 2)
        _ARTCCPA1.CUST_CREDIT_CARD_VER_CODE = txtCUST_CREDIT_CARD_VER_CODE.Text

        '_ARTCCPA1.CCPA_AMT = txtCCPA_AMT.Text
        _ARTCCPA1.CCPA_AMT = 1

        Dim client As New HttpClient()
        client.BaseAddress = New Uri(API_BASE)

        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))

        Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()

        Dim content As HttpContent = New ObjectContent(Of ARTCCPA1)(_ARTCCPA1, frmtr)

        Dim resp As HttpResponseMessage = Nothing
        Dim resp_err As String = ""
        Try
            resp = client.PostAsync(API_CONTROLLER, content).Result
        Catch ex As Exception
            resp_err = ex.InnerException.InnerException.Message
        End Try
        lblAuth.Visible = False

        Const iTitle As String = "Credit Card Charge"
        Dim iMSG As New System.Text.StringBuilder
        If resp_err.Length = 0 Then
            'Dim apiResponseString As String = JsonConvert.SerializeObject(resp)
            If resp.IsSuccessStatusCode Then
                Dim apiResponseString As String = ""
                Dim responseObject As Object = Nothing
                responseObject = resp.Content.ReadAsAsync(Of IEnumerable(Of ARTCCPA1))().Result
                apiResponseString = JsonConvert.SerializeObject(responseObject)
                'txtResponse.Text = apiResponseString
                Dim ARTCCPA1_R As New ARTCCPA1
                ARTCCPA1_R = responseObject(0)
                Dim ccpaNo As String = ARTCCPA1_R.CCPA_NO
                Dim responseText As String = ARTCCPA1_R.RESPONSE_TEXT
                If ARTCCPA1_R.RESPONSE_CODE = "A" Then 'Accepted
                    iMSG.AppendLine("Credit Card Information Recorded")
                    'TODO: Record CCPAS_NO on Order.
                    For Each rowSOTORDR1 As DataRow In FF.dst.Tables("SOTORDR1").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
                        rowSOTORDR1.Item("CCPA_NO") = ARTCCPA1_R.CCPA_NO
                        'Removed Per Ed Z on 5/22/14
                        'rowSOTORDR1.Item("CC_TRANS_ID") = ARTCCPA1_R.TRANS_ID
                    Next
                Else
                    If IsNothing(responseText) Then
                        responseText = "No Response From Credit Card Provider"
                    End If
                    iMSG.AppendLine(responseText)
                    iResult = iMSG.ToString
                End If

            Else
                'iResult = String.Format("{0} ({1})", CInt(resp.StatusCode), resp.ReasonPhrase)
                iResult = "Error Processing Credit Card.  Please Check Information Provided"
                iMSG.AppendLine(iResult)
            End If
        Else
            'iResult = String.Format(resp_err)
            iResult = "Error Processing Credit Card.  Please Check Information Provided"
            iMSG.AppendLine(iResult)
        End If
        MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        Return iResult
    End Function

    Function VerifyControls() As String
        Dim iResult As String = ""
        Dim Controls As String() = {"CUST_CREDIT_CARD_NO",
                                    "CUST_CREDIT_CARD_VER_CODE",
                                    "CUST_CREDIT_CARD_NAME",
                                    "CUST_CREDIT_CARD_ADDR1",
                                    "CUST_CREDIT_CARD_CITY",
                                    "CUST_CREDIT_CARD_STATE",
                                    "CUST_CREDIT_CARD_ZIP_CODE"}
        For Each Control As String In Controls
            If Absx1.txtFor(Control).Text.Length = 0 Then
                iResult += vbCrLf & Absx1.txtFor(Control).Tag.ToString & " Can Not Be Empty."
            End If
        Next
        If optCC_TYPE.CheckedIndex = -1 Then
            iResult += vbCrLf & "Credit Card Type Must Be Defined."
        End If
        If cboCUST_CREDIT_CARD_EXP_DATE.Text.Length = 0 Then
            iResult += vbCrLf & "Expiration Date Can Not Be Empty."
        Else
            If cboCUST_CREDIT_CARD_EXP_DATE.Text.Length <> 7 Then '01/2012
                iResult += vbCrLf & "Expiration Date Must Be In Format MM/YYYY."
            Else
                If cboCUST_CREDIT_CARD_EXP_DATE.Text.Substring(2, 1) <> "/" Then
                    iResult += vbCrLf & "Expiration Date Must Be In Format MM/YYYY."
                End If
            End If
        End If

        Return iResult
    End Function
    'Function UpdateARTCCPA1(ByVal ORDR_NO As String) As String
    '    Dim iresult As String = ""
    '    Dim API_BASE As String = ""
    '    Dim url As New System.Uri("http://50.75.200.254:8181/")
    '    Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
    '    Dim resptest As System.Net.WebResponse
    '    req.Timeout = 20000 '20 Seconds
    '    Try
    '        resptest = req.GetResponse()
    '        resptest.Close()
    '        req = Nothing
    '        API_BASE = "http://50.75.200.254:8181/"
    '    Catch ex As Exception
    '        req = Nothing
    '        API_BASE = "http://192.168.110.224:8181/"
    '    End Try

    '    Dim API_CONTROLLER As String = "api/XXXXXXXX/XXXXXXXXXX"
    '    Dim XXXXXXXX0 As String = ""
    '    Dim XXXXXXXX1 As String = ""
    '    Dim API_QUERY_STRING As String = String.Format("?XXXXXXXX0={0}&XXXXXXXX1={1}&CSQL=", XXXXXXXX0, XXXXXXXX1)

    '    Dim client As New HttpClient()
    '    client.Timeout = TimeSpan.FromSeconds(20)
    '    client.BaseAddress = New Uri(API_BASE)

    '    client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

    '    Dim response As HttpResponseMessage
    '    Try
    '        response = client.GetAsync(API_CONTROLLER & API_QUERY_STRING).Result
    '    Catch ex As Exception
    '        iresult = ex.Message
    '        Return iresult
    '        Exit Function
    '    End Try

    '    If response.IsSuccessStatusCode Then
    '        Try
    '            Dim ARTCCPA1_RESPONSE As New List(Of String)
    '            Dim apiResponseString As String = ""
    '            Dim responseObject As Object = response.Content.ReadAsAsync(Of IEnumerable(Of String))().Result

    '            ARTCCPA1_RESPONSE = responseObject
    '            If ARTCCPA1_RESPONSE.Count > 0 Then
    '                If ARTCCPA1_RESPONSE(0).Length = 0 Then
    '                    iresult = ""
    '                Else
    '                    iresult = ARTCCPA1_RESPONSE(0)
    '                End If
    '            End If
    '        Catch ex As Exception
    '            iresult = ex.Message
    '        End Try

    '    Else
    '        iresult = response.ReasonPhrase
    '    End If
    '    Return iresult
    'End Function
#End Region

    Private Sub FillExpDate()
        Dim CurMonth As Integer = Now.Month()
        Dim CurYear As Integer = Now.Year
        Dim CUST_CREDIT_CARD_EXP_DATE As String()
        Dim CurPointer As Integer = 0
        ReDim CUST_CREDIT_CARD_EXP_DATE(CurPointer)
        Dim FillYears As Integer = CurYear + 7
        Do While CurYear < FillYears
            CUST_CREDIT_CARD_EXP_DATE(CurPointer) = String.Format("{0}/{1}", Format(CurMonth, "00"), Format(CurYear, "0000"))
            If CurMonth = 12 Then
                CurYear += 1
                CurMonth = 1
            Else
                CurMonth += 1
            End If
            CurPointer += 1
            ReDim Preserve CUST_CREDIT_CARD_EXP_DATE(CurPointer)
        Loop
        cboCUST_CREDIT_CARD_EXP_DATE.DataSource = CUST_CREDIT_CARD_EXP_DATE
    End Sub

    Private Sub btnCustAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnCustAdd.Click
        Dim x As Integer = dst.Tables.Count()
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        If Not IsNothing(rowARTCUST1) Then
            txtCUST_CREDIT_CARD_ADDR1.Text = rowARTCUST1.Item("CUST_ADDR1").ToString
            txtCUST_CREDIT_CARD_CITY.Text = rowARTCUST1.Item("CUST_CITY").ToString
            txtCUST_CREDIT_CARD_STATE.Text = rowARTCUST1.Item("CUST_STATE").ToString
            txtCUST_CREDIT_CARD_ZIP_CODE.Text = rowARTCUST1.Item("CUST_ZIP_CODE").ToString
        End If
    End Sub

    Private Sub optCC_TYPE_ValueChanged(sender As Object, e As EventArgs) Handles optCC_TYPE.ValueChanged
        If optCC_TYPE.Value = "A" Then
            lblAMEX.Visible = True
        Else
            lblAMEX.Visible = False
        End If
    End Sub
End Class