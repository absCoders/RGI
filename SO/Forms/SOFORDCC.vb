Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Http.Formatting
Imports Newtonsoft.Json
Imports Newtonsoft
Imports System.Xml
Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinEditors

Public Class SOFORDCC
    Private FF As ASFBASE1
    Private API_BASE As String = "https://api2.regency-rib.com:8086/"
    'Private API_BASE As String = "https://be03-172-254-190-138.ngrok.io/"

    Public CCProcessed As Boolean = False
    Public CUST_CODE As String = ""
    Public CREDIT_CARD_ON_FILE As Boolean = False

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
        setDefaultCountry()
        CalculateTotal()
        SetGroupLabel()
        FillExpDate()
        fetchCConFile()
        SplitContainer2.SplitterDistance = SplitContainer2.Parent.Height - btnCustAdd.Height - 50
    End Sub

    Private Sub setDefaultCountry()
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT T1.COUNTRY_CODE2")
        SQLS.AppendLine("FROM ARTCUST1 A1, TATCNTRY T1")
        SQLS.AppendLine("WHERE A1.CUST_COUNTRY = T1.COUNTRY_CODE3")
        SQLS.AppendLine($"AND A1.CUST_CODE = '{CUST_CODE}'")
        ASCMAIN1.sql = SQLS.ToString()
        Dim COUNTRY_CODE As String = ASCDATA1.GetDataValue
        If COUNTRY_CODE.Length = 2 Then
            txtCUST_CREDIT_CARD_COUNTRY.Text = COUNTRY_CODE
        End If
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
        grdSOTORDRS.Text = "Orders On Group " & ORDR_GROUP_NO
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
                    If Not iResult_temp.StartsWith("Approved") Then
                        iResult = String.Format("{0}Order {1}: {2}", vbCrLf, ORDR_NO, iResult_temp)
                    End If
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
        authController = "AuthorizeCardSSL"
        Dim API_CONTROLLER As String = "api/RGI/SO/" & authController
        Dim url As New System.Uri(API_BASE)
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)

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
        _ARTCCPA1.CUST_CREDIT_CARD_COUNTRY = txtCUST_CREDIT_CARD_COUNTRY.Text
        _ARTCCPA1.ORDR_NO = ORDR_NO
        _ARTCCPA1.INIT_OPER = ASCMAIN1.USER_ID
        _ARTCCPA1.CUST_CREDIT_CARD_EXP_DATE = cboCUST_CREDIT_CARD_EXP_DATE.Text.Substring(0, 2) & cboCUST_CREDIT_CARD_EXP_DATE.Text.Substring(cboCUST_CREDIT_CARD_EXP_DATE.Text.Length - 2, 2)
        If chkCUST_CREDIT_CARD_PREFERRED.Checked Then
            _ARTCCPA1.CUST_CREDIT_CARD_PREFERRED = "1"
        Else
            _ARTCCPA1.CUST_CREDIT_CARD_PREFERRED = "0"
        End If

        If CREDIT_CARD_ON_FILE Then
            _ARTCCPA1.CUST_CREDIT_CARD_KEY = txtCUST_CREDIT_CARD_NO.Text
        Else
            _ARTCCPA1.CUST_CREDIT_CARD_NO = txtCUST_CREDIT_CARD_NO.Text
            _ARTCCPA1.CUST_CREDIT_CARD_VER_CODE = txtCUST_CREDIT_CARD_VER_CODE.Text
        End If

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
            If resp.IsSuccessStatusCode Then
                Dim apiResponseString As String = ""
                Dim responseObject As Object = Nothing
                responseObject = resp.Content.ReadAsAsync(Of IEnumerable(Of ARTCCPA1))().Result
                apiResponseString = JsonConvert.SerializeObject(responseObject)
                Dim ARTCCPA1_R As New ARTCCPA1
                ARTCCPA1_R = responseObject(0)
                Dim ccpaNo As String = ARTCCPA1_R.CCPA_NO
                Dim responseText As String = ARTCCPA1_R.RESPONSE_TEXT
                If ARTCCPA1_R.RESPONSE_CODE = "A" Or ARTCCPA1_R.RESPONSE_TEXT = "Approved" Then 'Accepted
                    iMSG.AppendLine("Credit Card Information Recorded")
                    For Each rowSOTORDR1 As DataRow In FF.dst.Tables("SOTORDR1").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
                        rowSOTORDR1.Item("CCPA_NO") = ARTCCPA1_R.CCPA_NO
                    Next
                Else
                    If IsNothing(responseText) Then
                        responseText = "No Response From Credit Card Provider"
                    Else
                        responseText = ARTCCPA1_R.RESPONSE_TEXT
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
            iResult = "Error Processing Credit Card.  Please Check Information Provided"
            iMSG.AppendLine(iResult)
        End If
        MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
        Return iResult
    End Function

    Function VerifyControls() As String
        Dim iResult As String = ""
        Dim Controls As String() = {"CUST_CREDIT_CARD_NO",
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

        If Not CREDIT_CARD_ON_FILE Then
            If optCC_TYPE.CheckedIndex = -1 Then
                iResult += vbCrLf & "Credit Card Type Must Be Defined."
            End If
            If Absx1.txtFor("CUST_CREDIT_CARD_VER_CODE").Text.Length = 0 Then
                iResult += "CVV2 Code Can Not Be Empty."
            End If
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

    Private Sub btdCardsOnFile_Click(sender As Object, e As EventArgs) Handles btdCardsOnFile.Click
        fetchCConFile()
    End Sub

    Private Sub fetchCConFile()
        btdCardsOnFile.Enabled = True
        grdCCONFILE.Text = "Waitinig For Credit Card info From Server."
        Dim iResult As String = ""
        Dim API_CONTROLLER As String = "api/RGI/SO/GetCustomerCards"
        Dim url As New System.Uri(API_BASE & "api/RGI/SO/ServerStatus")
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
        req.Timeout = 20000 '20 Seconds
        System.Net.ServicePointManager.Expect100Continue = True
        System.Net.ServicePointManager.SecurityProtocol = 3072

        lblAuth.Text = "Attempting to Fetch Cards On File..."
        lblAuth.Visible = True
        Dim CUST_CODE As String = FF.dst.Tables("SOTORDR1").Rows(0).Item("CUST_CODE").ToString
        If CUST_CODE.Length > 0 Then
            Dim client As New HttpClient()
            client.BaseAddress = New Uri(API_BASE)

            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))

            Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()
            Dim CUST_REQ As New CCRequest With {.CUST_CODE = CUST_CODE}
            Dim content As HttpContent = New ObjectContent(Of CCRequest)(CUST_REQ, frmtr)

            Dim resp As HttpResponseMessage = Nothing
            Dim resp_err As String = ""
            Try
                resp = client.PostAsync(API_CONTROLLER, content).Result
            Catch ex As Exception
                resp_err = ex.InnerException.InnerException.Message
            End Try
            lblAuth.Visible = False
            Dim iMSG As New System.Text.StringBuilder
            If resp_err.Length = 0 Then
                If resp.IsSuccessStatusCode Then
                    Dim apiResponseString As String = ""
                    Dim responseObject As Object = Nothing
                    responseObject = resp.Content.ReadAsAsync(Of Object)().Result
                    apiResponseString = JsonConvert.SerializeObject(responseObject("CustomerCards"))
                    Dim dt As New System.Data.DataTable
                    dt = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataTable)(apiResponseString)
                    If dt.Rows.Count > 0 Then
                        grdCCONFILE.DataSource = dt
                        grdCCONFILE.Text = "Credit Card info On File. (Double-Click To Select)"
                        btdCardsOnFile.Enabled = False
                        Sort_grdColumns(grdCCONFILE, "cust_credit_card_preferred,CUST_CREDIT_CARD_NAME", False)
                    Else
                        grdCCONFILE.DataSource = Nothing
                        grdCCONFILE.Text = "No Credit Cards On File For This Customer."
                        btdCardsOnFile.Enabled = True
                    End If

                Else
                    iResult = "Error Requesting Credit Cards On File.  Please Check Information Provided"
                    grdCCONFILE.Text = "No Credit Card info From Server."
                    iMSG.AppendLine(iResult)
                End If
            Else
                iResult = "Error Requesting Credit Cards On File.  Please Check Information Provided"
                grdCCONFILE.Text = "No Credit Card info From Server."
                iMSG.AppendLine(iResult)
                MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, "Error")
            End If
        Else
            grdCCONFILE.Text = "No Credit Card info From Server."
        End If
    End Sub

    Private Sub grdCCONFILE_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdCCONFILE.DoubleClickCell
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Use Card on File"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        iMSG.AppendLine("Are You Sure You Want To Use")
        iMSG.AppendLine("The Selected Card on File?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            If Not IsNothing(grdCCONFILE.ActiveRow) Then
                txtCUST_CREDIT_CARD_NAME.Text = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_NAME").Text & String.Empty
                txtCUST_CREDIT_CARD_NO.Text = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_KEY").Text & String.Empty
                txtCUST_CREDIT_CARD_NO.Enabled = False
                txtCUST_CREDIT_CARD_VER_CODE.Enabled = False
                txtCUST_CREDIT_CARD_ADDR1.Text = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_ADDR1").Text & String.Empty
                txtCUST_CREDIT_CARD_CITY.Text = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_CITY").Text & String.Empty
                txtCUST_CREDIT_CARD_STATE.Text = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_STATE").Text & String.Empty
                txtCUST_CREDIT_CARD_ZIP_CODE.Text = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_ZIP_CODE").Text & String.Empty
                txtLAST4.Text = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_LAST4").Text & String.Empty
                txtCUST_CREDIT_CARD_VER_CODE.Text = ""
                txtCUST_CREDIT_CARD_COUNTRY.Text = ""
                If grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_PREFERRED").Value & String.Empty = "1" Then
                    chkCUST_CREDIT_CARD_PREFERRED.Checked = True
                Else
                    chkCUST_CREDIT_CARD_PREFERRED.Checked = False
                End If
                txtShowNumbers.Enabled = False
                optCC_TYPE.Enabled = False
                txtCUST_CREDIT_CARD_VER_CODE.Enabled = False
                Dim CUST_CREDIT_CARD_EXP_DATE As String = grdCCONFILE.ActiveRow.Cells("CUST_CREDIT_CARD_EXP_DATE").Text & String.Empty
                Dim CC_EXP As String = ""
                If CUST_CREDIT_CARD_EXP_DATE.Length = 4 Then
                    CC_EXP = CUST_CREDIT_CARD_EXP_DATE.Substring(0, 2) & "/20" & CUST_CREDIT_CARD_EXP_DATE.Substring(2, 2)
                    cboCUST_CREDIT_CARD_EXP_DATE.Text = CC_EXP
                End If
                CREDIT_CARD_ON_FILE = True
            End If
        End If
    End Sub

    Private Sub txtCUST_CREDIT_CARD_COUNTRY_EditorButtonClick(sender As Object, e As EditorButtonEventArgs) Handles txtCUST_CREDIT_CARD_COUNTRY.EditorButtonClick
        Dim s As New Text.StringBuilder With {.Length = 0}
        s.Length = 0
        s.AppendLine("SELECT")
        s.AppendLine("COUNTRY_CODE2,")
        s.AppendLine("COUNTRY_NAME,")
        s.AppendLine("TO_CHAR(ROWNUM +4,'fm000') AS COUNTRY_SORT")
        s.AppendLine("FROM TATCNTRY")
        s.AppendLine("ORDER BY COUNTRY_NAME")
        ASCMAIN1.sql = s.ToString
        Dim TATCNTRY_TEMP As String = ASCMAIN1.Temp_Table

        Dim TOPS As New Dictionary(Of String, String)
        TOPS.Add("US", "001")
        TOPS.Add("CA", "002")
        TOPS.Add("MX", "003")
        TOPS.Add("PR", "004")

        For Each TOP As KeyValuePair(Of String, String) In TOPS
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine($"UPDATE {TATCNTRY_TEMP} SET COUNTRY_SORT = '{TOP.Value}' WHERE COUNTRY_CODE2 = '{TOP.Key}'")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
        Next

        s.Length = 0
        s.AppendLine($"SELECT COUNTRY_SORT AS NO, COUNTRY_NAME AS NAME, COUNTRY_CODE2 AS CODE FROM {TATCNTRY_TEMP}")
        'Dim sql As New Text.StringBuilder With {.Length = 0}
        'sql.AppendLine("SELECT *")
        'sql.AppendLine("FROM ASTVIEW1")
        'sql.AppendLine("WHERE ROWNUM < 0")
        'Dim tblTMP As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        'Dim rowTMP As DataRow = tblTMP.NewRow
        'rowTMP.Item("VIEW_NAME") = TATCNTRY_TEMP
        'rowTMP.Item("TABLE_NAME") = TATCNTRY_TEMP
        'rowTMP.Item("ORDER_BY") = "NO, CODE"
        'rowTMP.Item("ORDER_BY_DESC") = "1"
        'tblTMP.Rows.Add(rowTMP)
        With ASCMAIN1.CodeSelector
            .SQL = s.ToString
            .MultipleSelections = False
            .PreviouslySelectedCodes0 = ""
            .Caption = "Select Country"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable 'tblTMP
        End With
        Dim F As New ASFCODE1
        'F.ControlCollection("grd")
        'Dim grd As f.Controls("ASFBASE2_Fill_Panel").Controls("SplitContainer1").Controls(1).Controls("grd").GetType() = F.Controls("ASFBASE2_Fill_Panel").Controls("SplitContainer1").Controls(1).Controls("grd")

        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            Dim CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CODE") & ""
            txtCUST_CREDIT_CARD_COUNTRY.Text = CODE
        End If

    End Sub
End Class

Class RootObject(Of T)
    Public Property Table As T
End Class
Class CCRequest
    Public CUST_CODE As String
End Class

Class CCResults
    Public CUST_CREDIT_CARD_PREFERRED As String
    Public CUST_CREDIT_CARD_KEY As String
    Public CUST_CREDIT_CARD_EXP_DATE As String
    Public CUST_CREDIT_CARD_LAST4 As String
    'Public CUST_CREDIT_CARD_VER_CODE As String
    Public CUST_CREDIT_CARD_NAME As String
    Public CUST_CREDIT_CARD_ADDR1 As String
    Public CUST_CREDIT_CARD_CITY As String
    Public CUST_CREDIT_CARD_STATE As String
    Public CUST_CREDIT_CARD_ZIP_CODE As String
End Class