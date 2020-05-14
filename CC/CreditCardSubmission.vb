Imports Infragistics.Win

Public Class CreditCardSubmission

    Private queryGridControl As String = String.Empty
    Private sCustomerCode As String = String.Empty
    Private gridInitialized As Boolean = False
    Private AbsCon As Object = Nothing 'New ABSConnector

    Private clsTACENCRY As TAC.ASCENCRY
    Private EncryptionType As TAC.ASCENCRY.EncrytpionTypes = TAC.ASCENCRY.EncrytpionTypes.AdvancedEncryptionStandard_AES

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        queryGridControl = "Select ARTCCPA1.* from ARTCCPA1 where ARTCCPA1.CUST_CODE = :PARM1 and CCPA_STATUS <> '0' and CCPA_STATUS <> 'D' and CCPA_STATUS <> 'X'"
        sCustomerCode = String.Empty
        gridInitialized = False

        clsTACENCRY = New TAC.ASCENCRY()
        ValidateEncryption()
    End Sub

    Public Sub New(ByVal EncryptionCode As String)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

       ' Add any initialization after the InitializeComponent() call.
        queryGridControl = "Select ARTCCPA1.* from ARTCCPA1 where ARTCCPA1.CUST_CODE = :PARM1 and CCPA_STATUS <> '0' and CCPA_STATUS <> 'D' and CCPA_STATUS <> 'X'"
        sCustomerCode = String.Empty
        gridInitialized = False

        If AbsCon Is Nothing Then
            AbsCon = New ABSConnector
        End If

        EncryptionCode = EncryptionCode.Trim
        Dim rowTATENCRY As DataRow = AbsCon.GetDataRow("Select * from TATENCRY Where ENCRYPT_CODE = '" & EncryptionCode & "'")

        If EncryptionCode.Length > 0 AndAlso rowTATENCRY IsNot Nothing Then
            EncryptionType = DirectCast(CInt(Val(rowTATENCRY.Item("ENCRYPT_TYPE") & String.Empty)), TAC.ASCENCRY.EncrytpionTypes)
            clsTACENCRY = New TAC.ASCENCRY(EncryptionType)
            clsTACENCRY.Key = rowTATENCRY.Item("ENCRYPT_KEY") & String.Empty
            clsTACENCRY.PaddingMode = rowTATENCRY.Item("ENCRYPT_PADDING") & String.Empty
            clsTACENCRY.CipherMode = rowTATENCRY.Item("ENCRYPT_CIPHER") & String.Empty
        Else
            clsTACENCRY = New TAC.ASCENCRY()
        End If
        ValidateEncryption()
    End Sub

    Public Sub New(ByVal EncryptionType As TAC.ASCENCRY.EncrytpionTypes, ByVal CipherMode As TAC.ASCENCRY.CipherTypes, _
                   ByVal PaddingMode As TAC.ASCENCRY.PaddingTypes, ByVal EncryptionKey As String)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        queryGridControl = "Select ARTCCPA1.* from ARTCCPA1 where ARTCCPA1.CUST_CODE = :PARM1 and CCPA_STATUS <> '0' and CCPA_STATUS <> 'D' and CCPA_STATUS <> 'X'"
        sCustomerCode = String.Empty
        gridInitialized = False

        clsTACENCRY = New TAC.ASCENCRY(EncryptionType)
        clsTACENCRY.Key = EncryptionKey
        clsTACENCRY.PaddingMode = PaddingMode
        clsTACENCRY.CipherMode = CipherMode
        ValidateEncryption()
    End Sub

    Private Sub ValidateEncryption()

        If AbsCon Is Nothing Then
            AbsCon = New ABSConnector
        End If

        Dim rowASTPARMP As DataRow = AbsCon.GetDataRow("Select * from ASTPARMP WHERE AS_PARM_KEY = 'Z'")
        If rowASTPARMP Is Nothing OrElse Not rowASTPARMP.Table.Columns.Contains("AS_PARM_USE_ENCRYPTION") OrElse rowASTPARMP.Item("AS_PARM_USE_ENCRYPTION") & String.Empty <> "1" Then
            clsTACENCRY.UseEncryption = False
        Else
            clsTACENCRY.UseEncryption = True
        End If
    End Sub

#Region "Control Properties"

    ''' <summary>
    ''' Customer Code to get CC data for
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CustomerCode() As String
        Get
            Return sCustomerCode
        End Get
        Set(ByVal value As String)
            sCustomerCode = value
        End Set
    End Property

    ''' <summary>
    ''' Gives access to the User Controls Grid Control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property UserControlGrid() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return grdControl
        End Get

    End Property

#End Region

#Region "Control Public Methods"

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <param name="displayCustomerCode">Customer Code used to fill query</param>
    ''' <remarks></remarks>
    Public Sub DisplayData(ByVal displayCustomerCode As String)

        Try
            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            sCustomerCode = displayCustomerCode

            Dim tblARTCCPAP As DataTable = AbsCon.GetDataTable(queryGridControl, "ARTCCPA1", "V", New Object() {displayCustomerCode})
            tblARTCCPAP.TableName = "ARTCCPA1"
            DecryptARTCCPA1()

            grdControl.DataSource = tblARTCCPAP

            If Not gridInitialized Then
                AbsCon.Create_Summary(grdControl, "CUST_CREDIT_CARD_LAST4", "Count")
                AbsCon.Create_Summary(grdControl, "CCPA_AMT")

                AbsCon.Add_Value_List(grdControl, "CCPA_STATUS")
                AbsCon.Add_Value_List(grdControl, "CCPA_REASON")
                AbsCon.Add_Value_List(grdControl, "RESPONSE_CODE")
                AbsCon.Add_Value_List(grdControl, "CCPA_TYPE")
            End If
            gridInitialized = True

            AbsCon.Sort_grdColumns(grdControl, "LAST_DATE".ToLower)
            grdControl.Text = "Credit Card Submission History for " & displayCustomerCode
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK)
        End Try

    End Sub

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DisplayData()
        DisplayData(sCustomerCode)
    End Sub

#End Region

#Region "User Control Controls"

    Private Sub grdControl_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdControl.InitializeRow

        If e.Row.Cells("RESPONSE_CODE").Text = "A" Then
            e.Row.Cells("RESPONSE_TEXT").Appearance.ForeColor = Drawing.Color.Green
        ElseIf e.Row.Cells("RESPONSE_CODE").Text = "E" Then
            e.Row.Cells("RESPONSE_TEXT").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("RESPONSE_TEXT").Appearance.ForeColor = Drawing.Color.Empty
        End If

        If e.Row.Cells("CCPA_STATUS").Value = "A" Then
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.Green
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "E" Then
            e.Row.Cells("CCPA_STATUS").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.White
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "D" Then
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "S" Then
            e.Row.Cells("CCPA_STATUS").Appearance.BackColor = Drawing.Color.Green
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.White
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "V" Then
            e.Row.Cells("CCPA_STATUS").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.White
        Else
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.Empty
        End If

    End Sub

#End Region

#Region "Private Procedures"

    Private Sub DecryptARTCCPA1()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCCPA1 As DataRow In AbsCon.dst.Tables("ARTCCPA1").rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE"}
                rowARTCCPA1.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1.Item(field & "_E") & String.Empty)
            Next
        Next
    End Sub

    Private Sub EncryptARTCCPA1()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCCPA1 As DataRow In AbsCon.dst.Tables("ARTCCPA1").rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE"}
                rowARTCCPA1.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCCPA1.Item(field) & String.Empty)
                rowARTCCPA1.Item(field) = DBNull.Value
            Next
        Next
    End Sub

    Private Sub DecryptARTCUSTC()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCUSTC As DataRow In AbsCon.dst.Tables("ARTCUSTC").rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE"}
                rowARTCUSTC.Item(field) = clsTACENCRY.DecryptString(rowARTCUSTC.Item(field & "_E") & String.Empty)
            Next
        Next
    End Sub

    Private Sub EncryptARTCUSTC()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCUSTC As DataRow In AbsCon.dst.Tables("ARTCUSTC").rows
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE"}
                rowARTCUSTC.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCUSTC.Item(field) & String.Empty)
                rowARTCUSTC.Item(field) = DBNull.Value
            Next
        Next
    End Sub

    Private Sub DecryptARTCUSPA()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCUSPA As DataRow In AbsCon.dst.Tables("ARTCUSPA").rows
            For Each field As String In New String() {"ACH_ROUTING_NO", "ACH_ACCT_NO"}
                rowARTCUSPA.Item(field) = clsTACENCRY.DecryptString(rowARTCUSPA.Item(field) & String.Empty)
            Next
        Next
    End Sub

    Private Sub EncryptARTCUSPA()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCUSPA As DataRow In AbsCon.dst.Tables("ARTCUSPA").rows
            For Each field As String In New String() {"ACH_ROUTING_NO", "ACH_ACCT_NO"}
                rowARTCUSPA.Item(field) = clsTACENCRY.EncryptString(rowARTCUSPA.Item(field) & String.Empty)
            Next
        Next
    End Sub

#End Region

End Class
