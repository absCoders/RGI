Imports Infragistics.Win.UltraWinGrid

Public Class CreditCardQueue

    Private sCustomerCode As String = String.Empty
    Private initialized As Boolean = False

    Private rowSOTPARM1 As DataRow = Nothing
    Private rowARTCUST1 As DataRow = Nothing
    Private rowARTCCPRC As DataRow

    Private AbsCon As Object = Nothing ' New ABSConnector 
    Private objCCProcessor As Object ' As New TAC.ARCCCARD()
    Private allowACHEntries As Boolean = False
    Private sAllowAutoAuthForm As Boolean = False

    Public RecordCustomerEvent As Boolean = False
    Public isInEditMode As Boolean = False

    ' Backwards Compatibility
    Public encryptionKey As String = String.Empty

    Private clsTACENCRY As TAC.ASCENCRY
    Private EncryptionType As TAC.ASCENCRY.EncrytpionTypes = TAC.ASCENCRY.EncrytpionTypes.AdvancedEncryptionStandard_AES
    Private clsAllowEdit As Boolean = False
    Private CUST_COUNTRY As String = String.Empty

#Region "Instantiate Class"

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        sCustomerCode = String.Empty
        initialized = False
        AllowAutoAuthForm = False
        isInEditMode = False

        clsTACENCRY = New TAC.ASCENCRY()
        ValidateEncryption()

    End Sub

    Public Sub New(ByVal EncryptionCode As String)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        sCustomerCode = String.Empty
        initialized = False
        AllowAutoAuthForm = False
        isInEditMode = False
        clsAllowEdit = False

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

    Public Sub New(ByVal EncryptionType As TAC.ASCENCRY.EncrytpionTypes, ByVal CipherMode As TAC.ASCENCRY.CipherTypes,
                   ByVal PaddingMode As TAC.ASCENCRY.PaddingTypes, ByVal EncryptionKey As String)
        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        sCustomerCode = String.Empty
        initialized = False
        clsAllowEdit = False
        AllowAutoAuthForm = False
        isInEditMode = False

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

#End Region

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
            grpHeader.Text = "Accounts for " & sCustomerCode
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
            Return grdCC
        End Get

    End Property

    ''' <summary>
    ''' Gives access to the User Controls Grid Control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property UserControlGridACH() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return grdACH
        End Get

    End Property

    ''' <summary>
    ''' Get whether in Edit Mode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AllowEdit() As Boolean
        Get
            Return clsAllowEdit
        End Get
        Set(value As Boolean)
            clsAllowEdit = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Allow Auto Authorization Form
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AllowAutoAuthForm() As Boolean
        Get
            Return sAllowAutoAuthForm
        End Get
        Set(ByVal value As Boolean)
            sAllowAutoAuthForm = value
        End Set
    End Property

    ''' <summary>
    ''' Fires if the Update is Successful. Passes in the Customer Master Row with modified data.
    ''' </summary>
    ''' <param name="rowARTCUST1"></param>
    ''' <remarks></remarks>
    Public Event UpdateClickEvent(ByRef rowARTCUST1 As DataRow)

    ''' <summary>
    ''' Fires when the user clicks the Cancel Button
    ''' </summary>
    ''' <remarks></remarks>
    Public Event CancelClickEvent()

#End Region

#Region "Control Public Methods"

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <param name="displayCustomerCode">Customer Code used to fill query</param>
    ''' <remarks></remarks>
    Public Sub DisplayData(ByVal displayCustomerCode As String, Optional ByVal unCheckEdit As Boolean = True)

        Try
            Dim sql As String = String.Empty

            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            sCustomerCode = displayCustomerCode
            If unCheckEdit Then
                clsAllowEdit = False
            End If
            ToggleCCEditable()

            With AbsCon.dst
                If Not AbsCon.dst.Tables.Contains("ARTCUSTC") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCUSTC", "*", 1)
                End If

                If Not AbsCon.dst.Tables.Contains("ARTCCPA1") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCCPA1", "*")
                End If

                If Not AbsCon.dst.Tables.Contains("ARTCUST1") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCUST1", "*")
                End If

                If Not AbsCon.dst.Tables.Contains("ARTCUSPA") Then
                    AbsCon.Create_TDA(.Tables.Add, "ARTCUSPA", "*")
                    .Tables("ARTCUSPA").Columns.Add("ACH_ACCT_NO_LAST4", GetType(System.String), "IIF(LEN(ACH_ACCT_NO) <= 4, ACH_ACCT_NO, SUBSTRING(ACH_ACCT_NO, LEN(ACH_ACCT_NO) - 3, 4))")
                End If

                If Not AbsCon.dst.Tables.Contains("TATCNTRY") Then
                    AbsCon.Create_TDA(.Tables.Add, "TATCNTRY", "*")
                    AbsCon.Fill_Records("TATCNTRY", String.Empty, True, "SELECT * FROM TATCNTRY")
                End If

            End With

            AbsCon.Fill_Records("ARTCUSTC", displayCustomerCode)
            DecryptARTCUSTC()

            AbsCon.Fill_Records("ARTCUST1", displayCustomerCode)

            AbsCon.Fill_Records("ARTCUSPA", String.Empty, True, "SELECT * FROM ARTCUSPA WHERE CUST_CODE = '" & displayCustomerCode & "' AND ACH_ACCT_STATUS = 'A'")
            DecryptARTCUSPA()

            AbsCon.dst.Tables("ARTCCPA1").Rows.Clear()

            ' Fill in Control Values
            CUST_COUNTRY = String.Empty
            If AbsCon.dst.Tables("ARTCUST1").Rows.count > 0 Then
                rowARTCUST1 = AbsCon.dst.Tables("ARTCUST1").Rows(0)
                grpHeader.Text = Space(20) & "Accounts for " & rowARTCUST1.Item("CUST_CODE") & " " & rowARTCUST1.Item("CUST_NAME")
                CUST_COUNTRY = rowARTCUST1.Item("CUST_COUNTRY") & String.Empty
                If AbsCon.Dst.Tables("TATCNTRY").Rows.Find(CUST_COUNTRY) IsNot Nothing Then
                    CUST_COUNTRY = AbsCon.Dst.Tables("TATCNTRY").Rows.Find(CUST_COUNTRY).ITEM("COUNTRY_CODE2") & String.Empty
                Else
                    CUST_COUNTRY = String.Empty
                End If
            Else
                rowARTCUST1 = AbsCon.dst.Tables("ARTCUST1").newrow
                rowARTCUST1.Item("CUST_CODE") = displayCustomerCode
                grpHeader.Text = Space(20) & "Accounts for " & rowARTCUST1.Item("CUST_CODE") & " " & rowARTCUST1.Item("CUST_NAME")
            End If

            grdCC.DataSource = AbsCon.dst.Tables("ARTCUSTC")
            ABSolution.ASCMAIN1.Add_Value_List(grdCC, "CUST_CREDIT_CARD_COUNTRY", "SELECT COUNTRY_CODE2 COUNTRY_CODE, COUNTRY_NAME FROM TATCNTRY")

            grdACH.DataSource = AbsCon.dst.Tables("ARTCUSPA")

            ' Initialization is done here; otherwise, the control crashes if the code in in 'New' when it looks for an Oralce Connection
            If Not initialized Then
                txtCUST_AUTO_CCPA_NOTE.MaxLength = AbsCon.dst.Tables("ARTCUST1").Columns("CUST_AUTO_CCPA_NOTE").MaxLength
                cbeCUST_AUTO_CC_OPER.DataSource = AbsCon.GetDataTable("SELECT USER_ID, USER_NAME FROM ASTUSER1")

                Dim CC_PROC_FOLDER As String = ""

                rowSOTPARM1 = AbsCon.GetDataRow("Select * From SOTPARM1 Where SO_PARM_KEY = 'Z'")
                rowARTCCPRC = AbsCon.Lookup("ARTCCPRC", rowSOTPARM1.Item("SO_PARM_CC_PROC_CODE") & String.Empty)
                If rowARTCCPRC IsNot Nothing Then
                    CC_PROC_FOLDER = rowARTCCPRC.Item("CC_PROC_FOLDER") & ""
                End If
                Dim SO_PARM_CC_PROC_CODE As String = rowSOTPARM1.Item("SO_PARM_CC_PROC_CODE") & String.Empty

                objCCProcessor = New TAC.ARCCCARD(CC_PROC_FOLDER, Val(SO_PARM_CC_PROC_CODE))

                AbsCon.Add_Value_List(grdCC, "CUST_CREDIT_CARD_TYPE")

                With grdCC.DisplayLayout.Bands(0)
                    .Columns("CUST_CREDIT_CARD_NO").Header.Fixed = True
                    .Columns("CUST_CREDIT_CARD_LAST4").Header.Fixed = True
                    .Columns("CUST_CREDIT_CARD_TYPE").Header.Fixed = True
                End With

                Dim vl As Infragistics.Win.ValueList

                If (Not grdACH.DisplayLayout.ValueLists.Exists("ACH_ACCT_TYPE_ID")) Then
                    vl = grdACH.DisplayLayout.ValueLists.Add("ACH_ACCT_TYPE_ID")
                    vl.ValueListItems.Add("0", "Unknown")
                    vl.ValueListItems.Add("1", "Business")
                    vl.ValueListItems.Add("2", "Personal")

                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_TYPE_ID").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_TYPE_ID").ValueList = grdACH.DisplayLayout.ValueLists("ACH_ACCT_TYPE_ID")
                End If


                If (Not grdACH.DisplayLayout.ValueLists.Exists("ACH_ACCT_STATUS")) Then
                    vl = grdACH.DisplayLayout.ValueLists.Add("ACH_ACCT_STATUS")
                    vl.ValueListItems.Add("", "Unknown")
                    vl.ValueListItems.Add("A", "Active")
                    vl.ValueListItems.Add("X", "Deleted")

                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_STATUS").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_STATUS").ValueList = grdACH.DisplayLayout.ValueLists("ACH_ACCT_STATUS")

                End If

                With grdACH.DisplayLayout.Bands(0)
                    .Columns("ACH_ACCT_NAME").Header.Fixed = True
                End With

                initialized = True
            End If

            SetupCustomerMasterData()
            chkCUST_AUTO_CC_AUTH.Enabled = sAllowAutoAuthForm

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

    Public Sub ClearData()
        Try
            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            If clsAllowEdit Then
                CancelUpdate()
            End If

            If AbsCon IsNot Nothing AndAlso AbsCon.dst.Tables.Count > 0 Then
                AbsCon.dst.Tables("ARTCUSTC").Rows.Clear()
                AbsCon.dst.Tables("ARTCUSPA").Rows.Clear()
                AbsCon.dst.Tables("ARTCCPA1").Rows.Clear()
                AbsCon.dst.Tables("ARTCUST1").Rows.Clear()
            End If

            optCUST_AUTO_CCPA.Value = "0"
            txtCUST_AUTO_CCPA_NOTE.Clear()
            dteCUST_AUTO_CC_AUTH_DATE.Value = String.Empty
            cbeCUST_AUTO_CC_OPER.Text = String.Empty
            chkCUST_AUTO_CC_AUTH.Checked = False
            chkCUST_AUTOQ_WEB.Checked = False
            clsAllowEdit = False
            grpHeader.Text = "."

            ToggleCCEditable()

        Catch ex As Exception

        End Try

    End Sub

    Public Function ValidateCardData() As String

        Dim emsg As String = String.Empty

        Try

            Dim CUST_AUTO_CCPA As String = optCUST_AUTO_CCPA.Value & String.Empty
            Dim CUST_AUTO_CCPA_NOTE As String = txtCUST_AUTO_CCPA_NOTE.Text.Trim
            Dim CUST_AUTO_CC_AUTH As String = IIf(chkCUST_AUTO_CC_AUTH.Checked, "1", "0")
            Dim CUST_AUTOQ_WEB As String = IIf(chkCUST_AUTOQ_WEB.Checked, "1", "0")
            Dim CUST_AUTO_CC_AUTH_DATE As String = dteCUST_AUTO_CC_AUTH_DATE.Value & String.Empty
            Dim CUST_AUTO_CC_OPER As String = cbeCUST_AUTO_CC_OPER.Value & String.Empty

            If AbsCon.dst.Tables("ARTCUSTC").Rows.Count = 0 AndAlso ",1,2,".Contains(CUST_AUTO_CCPA) Then
                emsg &= vbCr & "There must be at least one credit card on file for Reminder or Auto Charge."
            End If

            If AbsCon.dst.Tables("ARTCUSPA").Rows.Count = 0 AndAlso ",3,".Contains(CUST_AUTO_CCPA) Then
                emsg &= vbCr & "There must be at least one Bank Account on file for ACH Processing."
            End If

            If ",3,".Contains(CUST_AUTO_CCPA) AndAlso AbsCon.dst.Tables("ARTCUSPA").Rows.Count > 0 AndAlso AbsCon.dst.Tables("ARTCUSPA").SELECT("ACH_AUTO_PAY_IND = '1'").LENGTH = 0 Then
                emsg &= vbCr & "There must be a Bank Account with 'Auto Pay' choosen for ACH Processing."
            End If

            If ",2,3,".Contains(CUST_AUTO_CCPA) AndAlso chkCUST_AUTO_CC_AUTH.Checked = False Then
                emsg &= vbCr & "You cannot select 'Auto Charge Queue' or 'ACH Queue' if the customer has not provided an Authorization Form."
            End If

            For Each rowARTCUSTC As DataRow In AbsCon.dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_STATUS = 'A'", "", DataViewRowState.CurrentRows)
                Dim CUST_CREDIT_CARD_NO As String = rowARTCUSTC.Item("CUST_CREDIT_CARD_NO") & ""
                Dim CUST_CREDIT_CARD_EXP_DATE As String = rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE") & ""
                Dim CUST_CREDIT_CARD_LAST4 As String = rowARTCUSTC.Item("CUST_CREDIT_CARD_LAST4") & ""

                Try
                    objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                    objCCProcessor.CustomerCreditCard.CardExpMonth = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 1, 2) & "")
                    objCCProcessor.CustomerCreditCard.CardExpYear = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 3, 2) & "")
                    objCCProcessor.ValidateCard()
                Catch ex As Exception
                    emsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ":" & ex.Message
                    Continue For
                End Try

                If CUST_CREDIT_CARD_EXP_DATE.Length = 0 OrElse Not IsNumeric(CUST_CREDIT_CARD_EXP_DATE) Then
                    emsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ": Expiration Date is Not Valid"
                End If

                If rowARTCUSTC.Item("CUST_CREDIT_CARD_NAME") & "" = "" Then
                    emsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ": Name Missing"
                End If

                If Not objCCProcessor.DigitCheckPassed Then
                    emsg &= vbCr & CUST_CREDIT_CARD_LAST4 & ": CC Check Digit failed"
                End If
            Next

        Catch ex As Exception
            emsg = $"Error Validating Credit Card Data: {ex.Message}"
        End Try

        Return emsg
    End Function

    Public Sub UpdateData()

        Dim beginTransaction As Boolean = False
        Dim rowSOTPARM1 As DataRow = Nothing
        Dim EMsg As String = String.Empty
        Dim sql As String = String.Empty
        Dim CUST_CODE As String = sCustomerCode

        Try
            If clsAllowEdit Then
                Dim CUST_AUTO_CCPA As String = optCUST_AUTO_CCPA.Value & String.Empty
                Dim CUST_AUTO_CCPA_NOTE As String = txtCUST_AUTO_CCPA_NOTE.Text.Trim
                Dim CUST_AUTO_CC_AUTH As String = IIf(chkCUST_AUTO_CC_AUTH.Checked, "1", "0")
                Dim CUST_AUTOQ_WEB As String = IIf(chkCUST_AUTOQ_WEB.Checked, "1", "0")
                Dim CUST_AUTO_CC_AUTH_DATE As String = dteCUST_AUTO_CC_AUTH_DATE.Value & String.Empty
                Dim CUST_AUTO_CC_OPER As String = cbeCUST_AUTO_CC_OPER.Value & String.Empty

                If CUST_AUTO_CCPA.Length = 0 Then CUST_AUTO_CCPA = "0"

                If chkCUST_AUTO_CC_AUTH.Checked Then
                    If Not IsDate(CUST_AUTO_CC_AUTH_DATE) Then
                        CUST_AUTO_CC_AUTH_DATE = DateTime.Now.ToString("dd-MMM-yyyy")
                    End If

                    If CUST_AUTO_CC_OPER.Length = 0 Then
                        CUST_AUTO_CC_OPER = AbsCon.UserID
                    End If
                End If

                ' Record event if the user changes the CUST_AUTO_CCPA or CUST_AUTO_CCPA_NOTE
                Dim zMsg As String = String.Empty
                If rowARTCUST1.Item("CUST_AUTO_CCPA", DataRowVersion.Original) & String.Empty <> CUST_AUTO_CCPA Then
                    zMsg = "CC Queue changed from "

                    Select Case rowARTCUST1.Item("CUST_AUTO_CCPA", DataRowVersion.Original) & String.Empty
                        Case "1"
                            zMsg &= optCUST_AUTO_CCPA.Items(1).DisplayText
                        Case "2"
                            zMsg &= optCUST_AUTO_CCPA.Items(2).DisplayText
                        Case Else
                            zMsg &= optCUST_AUTO_CCPA.Items(0).DisplayText
                    End Select

                    zMsg &= " to "
                    zMsg &= optCUST_AUTO_CCPA.CheckedItem.DisplayText

                    Record_Customer_Event(CUST_CODE, zMsg, "C")
                End If

                If (rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE", DataRowVersion.Original) & String.Empty).ToString.Trim.ToUpper <> CUST_AUTO_CCPA_NOTE.Trim.ToUpper Then
                    zMsg = "CC Note changed to: " & CUST_AUTO_CCPA_NOTE
                    Record_Customer_Event(CUST_CODE, zMsg, "C")
                End If

                If (rowARTCUST1.Item("CUST_AUTO_CCPA", DataRowVersion.Original) & String.Empty).ToString.Trim.ToUpper <> CUST_AUTO_CCPA.Trim.ToUpper Then
                    zMsg = "CC Queue changed to: " & CUST_AUTO_CCPA
                    Record_Customer_Event(CUST_CODE, zMsg, "C")
                End If

                rowARTCUST1.Item("CUST_AUTO_CCPA") = CUST_AUTO_CCPA
                rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE") = CUST_AUTO_CCPA_NOTE

                Dim rowARTCUST1_orig As DataRow = AbsCon.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
                rowARTCUST1_orig.Table.TableName = "ARTCUST1"

                rowARTCUST1.Item("CUST_AUTO_CCPA") = CUST_AUTO_CCPA
                rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE") = CUST_AUTO_CCPA_NOTE
                rowARTCUST1.Item("CUST_AUTO_CC_AUTH") = CUST_AUTO_CC_AUTH
                rowARTCUST1.Item("CUST_AUTOQ_WEB") = CUST_AUTOQ_WEB

                If CUST_AUTO_CC_AUTH_DATE.Length = 0 Then
                    sql = "Update ARTCUST1 Set"
                    sql &= "  CUST_AUTO_CCPA = :PARM1"
                    sql &= ", CUST_AUTO_CCPA_NOTE = :PARM2"
                    sql &= ", CUST_AUTO_CC_AUTH = :PARM3"
                    sql &= ", CUST_AUTOQ_WEB = :PARM4"
                    sql &= " where CUST_CODE = :PARM5"

                    AbsCon.ExecuteSQL(sql, "VVVVV", New Object() {
                        CUST_AUTO_CCPA,
                        CUST_AUTO_CCPA_NOTE,
                        CUST_AUTO_CC_AUTH, CUST_AUTOQ_WEB,
                        CUST_CODE})
                Else

                    sql = "Update ARTCUST1 Set"
                    sql &= "  CUST_AUTO_CCPA = :PARM1"
                    sql &= ", CUST_AUTO_CCPA_NOTE = :PARM2"
                    sql &= ", CUST_AUTO_CC_AUTH = :PARM3"
                    sql &= ", CUST_AUTO_CC_AUTH_DATE = :PARM4"
                    sql &= ", CUST_AUTO_CC_OPER = :PARM5"
                    sql &= ", CUST_AUTOQ_WEB = :PARM6"
                    sql &= " where CUST_CODE = :PARM7"

                    AbsCon.ExecuteSQL(sql, "VVVDVVV", New Object() {
                        CUST_AUTO_CCPA,
                        CUST_AUTO_CCPA_NOTE,
                        CUST_AUTO_CC_AUTH,
                        IIf(CUST_AUTO_CC_AUTH_DATE.Length > 0, CDate(CUST_AUTO_CC_AUTH_DATE).ToString("dd-MMM-yyyy"), ""),
                        CUST_AUTO_CC_OPER, CUST_AUTOQ_WEB,
                        CUST_CODE})

                    rowARTCUST1.Item("CUST_AUTO_CC_AUTH_DATE") = IIf(CUST_AUTO_CC_AUTH_DATE.Length > 0, CUST_AUTO_CC_AUTH_DATE, "")
                    rowARTCUST1.Item("CUST_AUTO_CC_OPER") = CUST_AUTO_CC_OPER

                End If

                Dim rowARTCUST1_curr As DataRow = AbsCon.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
                rowARTCUST1_curr.Table.TableName = "ARTCUST1"
                rowARTCUST1_curr.Item("CUST_AUTO_CCPA") = CUST_AUTO_CCPA
                rowARTCUST1_curr.Item("CUST_AUTO_CCPA_NOTE") = CUST_AUTO_CCPA_NOTE
                rowARTCUST1_curr.Item("CUST_AUTO_CC_AUTH") = CUST_AUTO_CC_AUTH
                rowARTCUST1_curr.Item("CUST_AUTOQ_WEB") = CUST_AUTOQ_WEB
                If CUST_AUTO_CC_AUTH_DATE.Length > 0 Then
                    rowARTCUST1_curr.Item("CUST_AUTO_CC_AUTH_DATE") = CUST_AUTO_CC_AUTH_DATE
                End If
                rowARTCUST1_curr.Item("CUST_AUTO_CC_OPER") = CUST_AUTO_CC_OPER

                AbsCon.DATETIME_STAMP = DateTime.Now
                AbsCon.Write_Audit_Trail(rowARTCUST1_curr, rowARTCUST1_orig, "E")

                AbsCon.INIT_LAST("ARTCUSTC", True, "CUST_CODE = '" & CUST_CODE & "'")
                EncryptARTCUSTC()
                AbsCon.Update_Record_TDA("ARTCUSTC", "CUST_CODE = '" & CUST_CODE & "'")

                AbsCon.INIT_LAST("ARTCUSPA", True, "CUST_CODE = '" & CUST_CODE & "'")
                EncryptARTCUSPA()
                AbsCon.Update_Record_TDA("ARTCUSPA", "CUST_CODE = '" & CUST_CODE & "'")

                Record_Customer_Event(CUST_CODE, "CC Info Updated", "C")

                clsAllowEdit = False
            End If
            RaiseEvent UpdateClickEvent(rowARTCUST1)

        Catch ex As Exception
            MessageBox.Show("Error updating Credit Card Data: " & ex.Message, "Error", MessageBoxButtons.OK)
        Finally
            For Each tableName As String In New String() {"ARTCUST1", "ARTCUSTC", "ARTCUSPA"}
                AbsCon.dst.Tables(tableName).Rows.clear
            Next
        End Try

    End Sub

    Public Sub CancelUpdate()
        clsAllowEdit = False
        sCustomerCode = String.Empty
        ClearData()
    End Sub

    Public Sub SetUpScreen()

        ToggleCCEditable()

        Try
            If clsAllowEdit Then
                If sCustomerCode.Length = 0 Then
                    clsAllowEdit = False
                    Exit Sub
                End If

                If Not AbsCon.Logical_Lock("ARTCUST1", sCustomerCode) Then
                    clsAllowEdit = False
                    Exit Sub
                End If

                If Not AbsCon.Logical_Lock("ARTCUSTC", sCustomerCode) Then
                    clsAllowEdit = False
                    Exit Sub
                End If

                ' Need to refresh the data incase it is stale.
                DisplayData(sCustomerCode, False)
                ' As of Today only Web Customers can set ACH; therefore, do not allow it to be changed on this screen
                ' it needs to be changed my the customer on the Web.
                If chkCUST_AUTOQ_WEB.Checked OrElse optCUST_AUTO_CCPA.Value = "3" Then
                    optCUST_AUTO_CCPA.Enabled = False
                    chkCUST_AUTOQ_WEB.Enabled = False
                Else
                    optCUST_AUTO_CCPA.Enabled = True
                    chkCUST_AUTOQ_WEB.Enabled = True
                End If

                chkCUST_AUTO_CC_AUTH.Enabled = sAllowAutoAuthForm
            Else
                AbsCon.MultiTask_Release()

                If AbsCon.DST.TABLES.CONTAINS("ARTCUSTC") Then
                    AbsCon.Fill_Records("ARTCUSTC", sCustomerCode)
                    DecryptARTCUSTC()
                End If

                If AbsCon.DST.TABLES.CONTAINS("ARTCUSPA") Then
                    AbsCon.Fill_Records("ARTCUSPA", String.Empty, True, "SELECT * FROM ARTCUSPA WHERE CUST_CODE = '" & sCustomerCode & "' AND ACH_ACCT_STATUS = 'A'")
                    DecryptARTCUSPA()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            clsAllowEdit = False
        End Try

        isInEditMode = (clsAllowEdit = True)

        ToggleCCEditable()
    End Sub


#End Region

#Region "Grid Control"

    Private Sub grdCC_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdCC.AfterCellUpdate
        Select Case e.Cell.Column.Key

            Case "CUST_CREDIT_CARD_PREFERRED"
                If Not grdCC.ActiveRow.IsDataRow Then
                    grdCC.UpdateData()
                End If

            Case "CUST_CREDIT_CARD_NO"
                Dim CUST_CREDIT_CARD_NO As String = e.Cell.Text
                If Len(CUST_CREDIT_CARD_NO) >= 15 Then
                    e.Cell.Row.Cells("CUST_CREDIT_CARD_LAST4").Value = Mid(CUST_CREDIT_CARD_NO, Len(CUST_CREDIT_CARD_NO) - 3, 4)
                Else
                    e.Cell.Row.Cells("CUST_CREDIT_CARD_LAST4").Value = ""
                End If

                objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                Dim CUST_CREDIT_CARD_TYPE As String = String.Empty
                Try
                    Select Case objCCProcessor.GetCreditCardType()
                        Case Nothing
                            ' Nothing 
                        Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                            CUST_CREDIT_CARD_TYPE = "AMEX"
                        Case TAC.ARCCCARD.CreditCardTypes.vctMasterCard
                            CUST_CREDIT_CARD_TYPE = "MSTR"
                        Case TAC.ARCCCARD.CreditCardTypes.vctVisa
                            CUST_CREDIT_CARD_TYPE = "VISA"
                        Case TAC.ARCCCARD.CreditCardTypes.vctDiscover
                            CUST_CREDIT_CARD_TYPE = "DISC"
                    End Select
                Catch ex As Exception

                End Try

                If CUST_CREDIT_CARD_TYPE.Length > 0 Then
                    Try
                        e.Cell.Row.Cells("CUST_CREDIT_CARD_TYPE").Value = (AbsCon.Get_Image(AbsCon.Folders("Images") & "\ABS\CC\", CUST_CREDIT_CARD_TYPE & ".GIF"))
                    Catch ex As Exception

                    End Try
                End If
        End Select
    End Sub

    Private Sub grdCC_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdCC.AfterRowUpdate
        If e.Row.Cells("CUST_CREDIT_CARD_PREFERRED").Value & "" = "1" Then
            Dim CUST_CREDIT_CARD_NO As String = e.Row.Cells("CUST_CREDIT_CARD_NO").Text
            For Each row As DataRow In AbsCon.dst.Tables("ARTCUSTC").Select("CUST_CREDIT_CARD_NO <> '" & CUST_CREDIT_CARD_NO & "'", "")
                row.Item("CUST_CREDIT_CARD_PREFERRED") = ""
            Next
        End If
    End Sub

    Private Sub grdCC_AfterRowInsert(sender As Object, e As RowEventArgs) Handles grdCC.AfterRowInsert
        Try
            e.Row.Cells("CUST_CREDIT_CARD_STATUS").Value = "A"
            e.Row.Cells("CUST_CREDIT_CARD_COUNTRY").Value = CUST_COUNTRY
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdCC_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdCC.BeforeRowUpdate

        grdCC.RowUpdateCancelAction = Infragistics.Win.UltraWinGrid.RowUpdateCancelAction.RetainDataAndActivation

        If CustomerCode.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Cannot Determine the Customer. Please contact ABS.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        With grdCC

            e.Row.Cells("CUST_CODE").Value = CustomerCode

            'CUST_CREDIT_CARD_STATUS - set to active if no value
            If (e.Row.Cells("CUST_CREDIT_CARD_STATUS").Value & String.Empty).ToString.Trim.Length = 0 Then
                e.Row.Cells("CUST_CREDIT_CARD_STATUS").Value = "A"
            ElseIf Not ("AI").Contains(e.Row.Cells("CUST_CREDIT_CARD_STATUS").Value & "") Then
                e.Cancel = True
                MessageBox.Show("Status must be set to Active or Inactive.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim CUST_CREDIT_CARD_KEY As String = e.Row.Cells("CUST_CREDIT_CARD_KEY").Value & String.Empty
            If CUST_CREDIT_CARD_KEY.Length = 0 Then
                CUST_CREDIT_CARD_KEY = ABSolution.ASCMAIN1.Next_Control_No("ARTCUSTC.CUST_CREDIT_CARD_KEY")
                e.Row.Cells("CUST_CREDIT_CARD_KEY").Value = CUST_CREDIT_CARD_KEY
            End If

            Dim CUST_CREDIT_CARD_NO As String = e.Row.Cells("CUST_CREDIT_CARD_NO").Text
            Dim CUST_CREDIT_CARD_EXP_DATE As String = e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Text
            If CUST_CREDIT_CARD_EXP_DATE.Length <> 4 Then
                e.Cancel = True
                MessageBox.Show("Exp MMYY must be in the format MMYY.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim expMonth As Int16 = Val(CUST_CREDIT_CARD_EXP_DATE.Substring(0, 2))
            Dim expYear As Int16 = Val(CUST_CREDIT_CARD_EXP_DATE.Substring(2, 2))
            If expMonth < 1 OrElse expMonth > 12 Then
                e.Cancel = True
                MessageBox.Show("Exp Month must be between 01 and 12.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim currentYear As Int16 = Val(DateTime.Now.ToString("yy"))
            Dim currentMonth As Int16 = Val(DateTime.Now.ToString("MM"))
            If expYear < currentYear Then
                e.Cancel = True
                MessageBox.Show("Exp Year must be greater equal the current year.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If expMonth < currentMonth AndAlso expYear <= currentYear Then
                e.Cancel = True
                MessageBox.Show("Exp Month/Year must be greater equal the current Month/Year.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim CUST_CREDIT_CARD_NAME As String = e.Row.Cells("CUST_CREDIT_CARD_NAME").Text & String.Empty
            CUST_CREDIT_CARD_NAME = CUST_CREDIT_CARD_NAME.Trim
            If CUST_CREDIT_CARD_NAME.Length = 0 Then
                e.Cancel = True
                MessageBox.Show("Name on Card is required.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim CUST_CREDIT_CARD_ZIP_CODE As String = e.Row.Cells("CUST_CREDIT_CARD_ZIP_CODE").Text & String.Empty
            CUST_CREDIT_CARD_ZIP_CODE = CUST_CREDIT_CARD_ZIP_CODE.Trim
            If CUST_CREDIT_CARD_ZIP_CODE.Length = 0 Then
                e.Cancel = True
                MessageBox.Show("Zip Code is required.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim CUST_CREDIT_CARD_COUNTRY As String = (e.Row.Cells("CUST_CREDIT_CARD_COUNTRY").Value & String.Empty).Trim.ToUpper
            e.Row.Cells("CUST_CREDIT_CARD_COUNTRY").Value = CUST_CREDIT_CARD_COUNTRY
            If CUST_CREDIT_CARD_COUNTRY.Length = 0 Then
                e.Cancel = True
                MessageBox.Show("Country is required.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim CUST_CREDIT_CARD_VER_CODE As String = e.Row.Cells("CUST_CREDIT_CARD_VER_CODE").Text & String.Empty
            CUST_CREDIT_CARD_VER_CODE = CUST_CREDIT_CARD_VER_CODE.Trim
            If CUST_CREDIT_CARD_VER_CODE.Length = 0 Then
                e.Cancel = True
                MessageBox.Show("CVV2 is required.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            e.Row.Cells("CUST_CREDIT_CARD_NO").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Appearance.ForeColor = Drawing.Color.Empty
            Try
                objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
                objCCProcessor.CustomerCreditCard.CardExpMonth = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 1, 2) & "")
                objCCProcessor.CustomerCreditCard.CardExpYear = Val(Mid$(CUST_CREDIT_CARD_EXP_DATE, 3, 2) & "")
                objCCProcessor.ValidateCard()
            Catch ex As Exception
                e.Row.Cells("CUST_CREDIT_CARD_EXP_DATE").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("CUST_CREDIT_CARD_NO").Appearance.ForeColor = Drawing.Color.Red
                'e.Cancel = True
                MessageBox.Show("The Credit Card Number appears to be an invlaid Card Number.", "CC Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                'Exit Sub
            End Try

        End With
    End Sub

    Private Sub grdCC_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdCC.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdCC.ActiveCell.Column.Key
            'Case "CUST_CREDIT_CARD_STATE"
            '    sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "' or LENS_DESIGNER_CODE is Null"
        End Select

        AbsCon.grdClickCellButton(grdCC, sql_where, False)
    End Sub

    Private Sub grdCC_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdCC.InitializeRow
        If e.Row.IsAddRow Then
        Else
            Dim row As DataRow = AbsCon.dst.Tables("ARTCUSTC").Rows.Find(New Object() {e.Row.Cells("CUST_CODE").Value & "", e.Row.Cells("CUST_CREDIT_CARD_NO").Value & ""})
            If row IsNot Nothing AndAlso row.RowState = DataRowState.Added Then
                e.Row.Cells("CUST_CREDIT_CARD_NO").Hidden = False
                e.Row.Cells("CUST_CREDIT_CARD_VER_CODE").Hidden = False
            Else
                e.Row.Cells("CUST_CREDIT_CARD_NO").Hidden = True
                e.Row.Cells("CUST_CREDIT_CARD_VER_CODE").Hidden = False
            End If
        End If

        If e.Row Is Nothing OrElse Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Try
            If objCCProcessor Is Nothing Then
                Exit Sub
            End If

            Dim CUST_CREDIT_CARD_NO As String = e.Row.Cells("CUST_CREDIT_CARD_NO").Text

            objCCProcessor.CustomerCreditCard.CardNumber = CUST_CREDIT_CARD_NO
            Dim CUST_CREDIT_CARD_TYPE As String = String.Empty
            Try
                Select Case objCCProcessor.GetCreditCardType()
                    Case Nothing
                            ' Nothing 
                    Case TAC.ARCCCARD.CreditCardTypes.vctAmex
                        CUST_CREDIT_CARD_TYPE = "AMEX"
                    Case TAC.ARCCCARD.CreditCardTypes.vctMasterCard
                        CUST_CREDIT_CARD_TYPE = "MSTR"
                    Case TAC.ARCCCARD.CreditCardTypes.vctVisa
                        CUST_CREDIT_CARD_TYPE = "VISA"
                    Case TAC.ARCCCARD.CreditCardTypes.vctDiscover
                        CUST_CREDIT_CARD_TYPE = "DISC"
                End Select
            Catch ex As Exception

            End Try

            If CUST_CREDIT_CARD_TYPE.Length > 0 Then
                Try
                    e.Row.Cells("CUST_CREDIT_CARD_TYPE").Value = (AbsCon.Get_Image(AbsCon.Folders("Images") & "\ABS\CC\", CUST_CREDIT_CARD_TYPE & ".GIF"))
                Catch ex As Exception

                End Try
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub grdACH_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdACH.AfterRowActivate

        ' As of today 4/18/2012 as per Maria no adding entries
        If Not allowACHEntries Then Exit Sub

        If Not grdACH.ActiveRow.IsAddRow AndAlso grdACH.ActiveRow.Cells("WEB_IND").Value & String.Empty = "1" Then
            grdACH.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Else
            If grdACH.DisplayLayout.Override.AllowUpdate <> Infragistics.Win.DefaultableBoolean.True Then
                grdACH.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            End If
        End If
    End Sub

    Private Sub grdACH_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdACH.AfterRowUpdate

        ' As of today 4/18/2012 as per Maria no adding entries
        If Not allowACHEntries Then Exit Sub

        If e.Row.Cells("ACH_DEFAULT_ACCT_IND").Value & "" = "1" Then
            Dim ACH_ROUTING_NO As String = e.Row.Cells("ACH_ROUTING_NO").Text
            For Each row As DataRow In AbsCon.dst.Tables("ARTCUSPA").Select("ACH_ROUTING_NO <> '" & ACH_ROUTING_NO & "'", "")
                row.Item("ACH_DEFAULT_ACCT_IND") = ""
            Next
        End If
    End Sub

    Private Sub grdACH_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdACH.BeforeRowUpdate

        ' As of today 4/18/2012 as per Maria no adding entries
        If Not allowACHEntries Then
            e.Cancel = True
            Exit Sub
        End If

        With grdACH
            Dim errorMessage As String = String.Empty
            If Not ValidateRoutingNumber(e.Row.Cells("ACH_ROUTING_NO").Text, errorMessage) Then
                MessageBox.Show(errorMessage, "ACH", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            End If

            If e.Row.Cells("ACH_ACCT_STATUS").Text = "" Then
                .ActiveRow.Cells("ACH_ACCT_STATUS").Value = "A"
            End If

            If e.Row.IsAddRow Then

                If e.Row.Cells("CUST_CODE").Text = "" Then
                    .ActiveRow.Cells("CUST_CODE").Value = sCustomerCode
                End If

                .ActiveRow.Cells("ACH_ACCT_ID").Value = ABSolution.ASCMAIN1.Next_Control_No("ARTCUSPA.ACH_ACCT_ID")
            End If
        End With

    End Sub

#End Region

#Region "Private Subs / Functions"

    Private Sub ToggleCCEditable()

        Try
            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            If clsAllowEdit Then

                AbsCon.Set_Read_Only(grpCCPA, False)
                With grdCC.DisplayLayout.Override
                    .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = Infragistics.Win.DefaultableBoolean.True
                    .AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                End With
                grdCC.DisplayLayout.Bands(0).Columns("CUST_CREDIT_CARD_NO").Hidden = False

                If allowACHEntries Then
                    With grdACH.DisplayLayout.Override
                        .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = Infragistics.Win.DefaultableBoolean.True
                        .AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                    End With
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ROUTING_NO").Hidden = False
                    grdACH.DisplayLayout.Bands(0).Columns("ACH_ACCT_NO").Hidden = False
                End If
                chkCUST_AUTO_CC_AUTH.Enabled = AllowAutoAuthForm
                If optCUST_AUTO_CCPA.Value = "3" Then
                    optCUST_AUTO_CCPA.Enabled = False
                Else
                    optCUST_AUTO_CCPA.Enabled = True
                End If
                isInEditMode = True
                chkShowActive.Checked = True
            Else
                AbsCon.Set_Read_Only(grpCCPA, True)
                With grdCC.DisplayLayout.Override
                    .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
                    .AllowDelete = Infragistics.Win.DefaultableBoolean.False
                    .AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                End With
                If grdCC.DisplayLayout.Bands(0).Columns.Contains("CUST_CREDIT_CARD_NO") Then
                    grdCC.DisplayLayout.Bands(0).Columns("CUST_CREDIT_CARD_NO").Hidden = True
                End If

                With grdACH.DisplayLayout.Override
                    .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
                    .AllowDelete = Infragistics.Win.DefaultableBoolean.False
                    .AllowUpdate = Infragistics.Win.DefaultableBoolean.False
                End With
                isInEditMode = False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        chkShowActive.Checked = isInEditMode
        chkShowActive_CheckedChanged(Nothing, Nothing)

    End Sub

    Private Sub Record_Customer_Event(ByVal CUST_CODE As String, ByVal EVENT_DESC As String, ByVal EVENT_TYPE As String)
        If Not RecordCustomerEvent Then Exit Sub
        TAC.ARCMAIN1.Record_Customer_Event(CUST_CODE, EVENT_DESC, EVENT_TYPE)
    End Sub

    Private Sub SetupCustomerMasterData()

        Try
            Select Case rowARTCUST1.Item("CUST_AUTO_CCPA") & String.Empty

                Case "0", "1", "2", "3"
                    optCUST_AUTO_CCPA.Value = rowARTCUST1.Item("CUST_AUTO_CCPA") & String.Empty
                Case Else
                    optCUST_AUTO_CCPA.Value = "0"
            End Select
            txtCUST_AUTO_CCPA_NOTE.Text = rowARTCUST1.Item("CUST_AUTO_CCPA_NOTE") & String.Empty

            If rowARTCUST1.Item("CUST_AUTO_CC_AUTH") & String.Empty = "1" Then
                chkCUST_AUTO_CC_AUTH.Checked = True
            Else
                chkCUST_AUTO_CC_AUTH.Checked = False
            End If

            If rowARTCUST1.Item("CUST_AUTOQ_WEB") & String.Empty = "1" Then
                chkCUST_AUTOQ_WEB.Checked = True
            Else
                chkCUST_AUTOQ_WEB.Checked = False
            End If

            If chkCUST_AUTOQ_WEB.Checked Then
                optCUST_AUTO_CCPA.Enabled = False
                chkCUST_AUTOQ_WEB.Enabled = False
            Else
                optCUST_AUTO_CCPA.Enabled = True
                chkCUST_AUTOQ_WEB.Enabled = True
            End If

            cbeCUST_AUTO_CC_OPER.Value = rowARTCUST1.Item("CUST_AUTO_CC_OPER") & String.Empty

            If IsDate(rowARTCUST1.Item("CUST_AUTO_CC_AUTH_DATE") & String.Empty) Then
                dteCUST_AUTO_CC_AUTH_DATE.Value = CDate(rowARTCUST1.Item("CUST_AUTO_CC_AUTH_DATE") & String.Empty)
            Else
                dteCUST_AUTO_CC_AUTH_DATE.Value = String.Empty
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Validate a Bank Routing Number
    ''' </summary>
    ''' <param name="routingNumber">Routing Number to evaluate</param>
    ''' <param name="resp">If returnsd false, error description is in resp</param>
    ''' <remarks></remarks>
    Public Shared Function ValidateRoutingNumber(ByVal routingNumber As String, ByRef resp As String) As Boolean

        If Not System.Text.RegularExpressions.Regex.IsMatch(routingNumber, "^\d+$") Then
            resp = "Routing number cannot have non-numeric characters"
            Return False

        ElseIf routingNumber.Length <> 9 Then
            resp = "Routing number must be 9 digits long"
            Return False

        Else
            Dim isValid As Boolean
            isValid = (Integer.Parse(routingNumber(0)) * 3 + Integer.Parse(routingNumber(1)) * 7 + Integer.Parse(routingNumber(2)) * 1 +
                      Integer.Parse(routingNumber(3)) * 3 + Integer.Parse(routingNumber(4)) * 7 + Integer.Parse(routingNumber(5)) * 1 +
                      Integer.Parse(routingNumber(6)) * 3 + Integer.Parse(routingNumber(7)) * 7 + Integer.Parse(routingNumber(8)) * 1) Mod 10 = 0
            If Not isValid Then
                resp = "Routing number is invalid"
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub DecryptARTCCPA1()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCCPA1 As DataRow In AbsCon.dst.Tables("ARTCCPA1").Select("")
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                rowARTCCPA1.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1.Item(field & "_E") & String.Empty)
                rowARTCCPA1.Item(field & "_E") = DBNull.Value
            Next
        Next
    End Sub

    Private Sub EncryptARTCCPA1()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If
        For Each rowARTCCPA1 As DataRow In AbsCon.dst.Tables("ARTCCPA1").Select("")
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                rowARTCCPA1.Item(field & "_E") = clsTACENCRY.EncryptString(rowARTCCPA1.Item(field) & String.Empty)
                rowARTCCPA1.Item(field) = DBNull.Value
            Next
        Next
    End Sub

    Private Sub DecryptARTCUSTC()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If

        For Each rowARTCUSTC As DataRow In AbsCon.dst.Tables("ARTCUSTC").Select("")
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                rowARTCUSTC.Item(field) = clsTACENCRY.DecryptString(rowARTCUSTC.Item(field & "_E") & String.Empty)
                rowARTCUSTC.Item(field & "_E") = DBNull.Value
            Next
        Next
    End Sub

    Private Sub EncryptARTCUSTC()
        If clsTACENCRY.UseEncryption = False Then
            Exit Sub
        End If

        For Each rowARTCUSTC As DataRow In AbsCon.dst.Tables("ARTCUSTC").Select("")
            For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} '  "CUST_CREDIT_CARD_EXP_DATE",
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

    Private Sub chkShowActive_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowActive.CheckedChanged

        If AbsCon Is Nothing Then
            Exit Sub
        End If

        If Not AbsCon.dst.Tables.contains("ARTCUSTC") Then
            Exit Sub
        End If

        Dim dvw As DataView = DirectCast(AbsCon.dst.Tables("ARTCUSTC"), DataTable).DefaultView
        Dim sql As String = "CUST_CREDIT_CARD_STATUS = 'A'"
        If Not chkShowActive.Checked Then
            sql = "CUST_CREDIT_CARD_STATUS <> '@@@@'"
        End If
        dvw.RowFilter = sql
        grdCC.DataSource = dvw
    End Sub

#End Region

End Class
