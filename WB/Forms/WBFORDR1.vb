Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.OleDb
Imports Infragistics.Win

Public Class WBFORDR1

    Private rowSOTPART1 As DataRow = Nothing
    Private PARTNER_SITE_ORDERS_POST_URL As String = String.Empty
    Private PARTNER_SITE_IP As String = String.Empty
    Private PARTNER_SITE_USER As String = String.Empty
    Private PARTNER_SITE_PWD As String = String.Empty
    Private PARTNER_SITE_OUTPUT_DIR As String = String.Empty
    Private PARTNER_ORDERS_DIR As String = String.Empty
    Private PARTNER_LAST_SALES_ORDER As String = String.Empty
    Private PARTNER_CODE As String = String.Empty
    Private PARTNER_ORDR_SOURCE_CODE As String = String.Empty

    Private viewSOTORDRV As DataView
    Private rowICTITEM1 As DataRow = Nothing
    Private rowICTSTYL1 As DataRow = Nothing

    Private clsSOCMAIN1 As New TAC.SOCMAIN1
    Private tblSOTORDR2 As DataTable
    Private tblSOTORDR5 As DataTable

    Private ftpFileList As List(Of String)
    Private WithEvents Sftp1 As New nsoftware.IPWorksSSH.Sftp
    Private WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

    Private skippedSalesOrder As List(Of String)

    Private SO_PARM_PP_SIGNATURE As String = String.Empty
    Private SO_PARM_PP_USERNAME As String = String.Empty
    Private SO_PARM_PP_PASSWORD As String = String.Empty
    Private SO_PARM_PP_URL As String = String.Empty

    Private testCreditCardNo As String = String.Empty
    Private testCreditCardExp As String = String.Empty
    Private testCreditCardCCV2 As String = String.Empty

    Private Const shopCCQuery As String = "PYMT_METHOD_CODE = 'CC' AND ORDR_SOURCE_CODE = 'SHP'"

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Define the Tables that will be used in the Form
        Dim sql As String = String.Empty
        Dim SOTERRCD As String

        Get_PARM("SOTPARM1")

        With dst
            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            With dst.Tables("SOTORDR1")
                .Columns.Add("HAS_ERRORS", GetType(System.String))
            End With

            Create_TDA(.Tables.Add, "SOTORDR2", "*")
            Create_TDA(.Tables.Add, "SOTORDR5", "*")
            Create_TDA(.Tables.Add, "SOTORDRG", "*")
            Create_TDA(.Tables.Add, "SOTPAYPL", "*")

            Create_TDA(.Tables.Add, "SOTORDRE", "*")
            Create_TDA(.Tables.Add, "SOTORDRV", "*")
            With dst.Tables("SOTORDRV")
                .Columns.Add("ERROR_DESC", GetType(System.String), "ERROR_CODE")
            End With

            Create_TDA(.Tables.Add, "SOTERRCD", "*")
            Create_TDA(.Tables.Add, "SOTORDRI", "*")
            Create_TDA(.Tables.Add, "TATCNTRY", "*")

            dst.Tables.Add("FILES_PROCESSED")
            With dst.Tables("FILES_PROCESSED")
                .Columns.Add("FileName", GetType(System.String))
                .Columns.Add("FilePath", GetType(System.String))
            End With

            CreateAmazonDataTable()
            CreateShopComDataTable()
            CreateBuyComDataTable()

            SOTERRCD = ASCMAIN1.Temp_Table("SELECT ERROR_CODE, ERROR_DESC FROM SOTERRCD")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & SOTERRCD & " ADD PRIMARY KEY (ERROR_CODE)")

            Create_TDA(.Tables.Add, SOTERRCD, "*")
            Fill_Records(SOTERRCD, String.Empty, True, "SELECT * FROM " & SOTERRCD)
            ' ADD THE EXTRA CODES
            For ERC As Integer = 1 To 100
                dst.Tables(SOTERRCD).Rows.Add(New Object() {"I_" & ERC, "Item Error, Line No " & ERC})
                dst.Tables(SOTERRCD).Rows.Add(New Object() {"P_" & ERC, "Zero Price, Line No " & ERC})
            Next

            MyBase.Update_Record_TDA(SOTERRCD)



        End With

        viewSOTORDRV = New DataView(dst.Tables("SOTORDRV"))
        grdSOTORDRV.DataSource = viewSOTORDRV

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDRI.DataSource = dst.Tables("SOTORDRI")
        ASCMAIN1.Add_Value_List(grdSOTORDRI, "PARTNER_CODE", "SELECT PARTNER_CODE, PARTNER_NAME FROM SOTPART1")

        Create_Lookup("SOTSVIA1")
        Create_Lookup("SOTSVIA2")
        Create_Lookup("ICTITEM1", "*", "ITEM_CODE = :PARM1", "V", False)
        Create_Lookup("ICTSTYL1")

        Fill_Records("SOTERRCD", String.Empty, True, "SELECT * FROM SOTERRCD")
        For ERC As Integer = 1 To 100
            dst.Tables("SOTERRCD").Rows.Add(New Object() {"I_" & ERC, "Item Error, Line No " & ERC})
            dst.Tables("SOTERRCD").Rows.Add(New Object() {"P_" & ERC, "Zero Price, Line No " & ERC})
        Next

        ASCMAIN1.Add_Value_List(grdSOTORDR1, "SHIP_VIA_CODE", "Select Ship_via_code, Ship_Via_Desc from SOTSVIA1")
        ASCMAIN1.Add_Value_List(grdSOTORDR1, "PYMT_METHOD_CODE", "SELECT PYMT_METHOD_CODE,  PYMT_METHOD_DESC FROM SOTPYMT1")

        Dim dicORDR_STATUS As Dictionary(Of String, String) = CodeValues("ORDR_STATUS")

        If dicORDR_STATUS IsNot Nothing Then
            sql = String.Empty
            For Each kvp As KeyValuePair(Of String, String) In dicORDR_STATUS
                Dim v1 As String = kvp.Key
                Dim v2 As String = kvp.Value
                sql &= " Union Select '" & v1 & "', '" & v2 & "' from dual"
            Next

            If sql.Length > 0 Then
                sql = sql.Substring(7).Trim
                ASCMAIN1.Add_Value_List(grdSOTORDR1, "ORDR_STATUS", sql)
            End If
        End If

        MyBase.Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")

        With grdSOTORDR1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
            {"ORDR_SALES_AMT", "ORDR_DISC_AMT", "ORDR_STAX_AMT", "ORDR_FRT_AMT", "ORDR_TAX_AMT", "ORDR_TOT_AMT", "ORDR_GIFTCERT_APPL"}
                If dst.Tables("SOTORDR1").Columns.Contains(COLUMN_NAME) Then
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.DodgerBlue
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Columns(COLUMN_NAME).Width = 75
                    Create_Summary(grdSOTORDR1, COLUMN_NAME, "Sum")
                End If
            Next
        End With

        MyBase.Create_Summary(grdSOTORDRV, "ERROR_CODE", "Count")

        ASCMAIN1.Add_Value_List(grdSOTORDRV, "ERROR_DESC", "SELECT ERROR_CODE, ERROR_DESC FROM " & SOTERRCD)

        dteStart.MaxDate = DateTime.Now
        dteEnd.MaxDate = DateTime.Now

        dteStart.MinDate = DateAdd(DateInterval.Year, -1, DateTime.Now)
        dteEnd.MinDate = DateAdd(DateInterval.Year, -1, DateTime.Now)

        dteEnd.Value = dteEnd.MaxDate
        dteStart.Value = dteStart.MaxDate

        btnGetData_Click(Nothing, Nothing)

        Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty
        Dim zmsg As String = String.Empty

        Select Case eItemKey

            Case "Cancel"

            Case "Get Sales Orders"

                Dim ORDR_SOURCE_CODE As String = MyBase.Absx1.txtFor("ORDR_SOURCE_CODE").Text.Trim
                Validate_Code("ORDR_SOURCE_CODE")

                If EMsg.Length = 0 Then
                    rowSOTPART1 = ASCDATA1.GetDataRow("SELECT * FROM SOTPART1 WHERE NVL(PARTNER_STATUS, 'A') = 'A' AND PARTNER_ORDR_SOURCE_CODE = :PARM1", "V", ORDR_SOURCE_CODE)

                    If rowSOTPART1 Is Nothing Then
                        EMsg = "Invalid or Inactive partner selected."
                    Else
                        PARTNER_SITE_ORDERS_POST_URL = rowSOTPART1.Item("PARTNER_SITE_ORDERS_POST_URL") & String.Empty
                        PARTNER_SITE_IP = rowSOTPART1.Item("PARTNER_SITE_IP") & String.Empty
                        PARTNER_SITE_USER = rowSOTPART1.Item("PARTNER_SITE_USER") & String.Empty
                        PARTNER_SITE_PWD = rowSOTPART1.Item("PARTNER_SITE_PWD") & String.Empty
                        PARTNER_ORDERS_DIR = rowSOTPART1.Item("PARTNER_ORDERS_DIR") & String.Empty
                        PARTNER_LAST_SALES_ORDER = rowSOTPART1.Item("PARTNER_LAST_SALES_ORDER") & String.Empty
                        PARTNER_SITE_OUTPUT_DIR = rowSOTPART1.Item("PARTNER_SITE_OUTPUT_DIR") & String.Empty
                        PARTNER_CODE = rowSOTPART1.Item("PARTNER_CODE") & String.Empty
                        PARTNER_ORDR_SOURCE_CODE = rowSOTPART1.Item("PARTNER_ORDR_SOURCE_CODE") & String.Empty

                        If PARTNER_ORDERS_DIR.Length > 0 AndAlso Not PARTNER_ORDERS_DIR.EndsWith("\") Then
                            PARTNER_ORDERS_DIR &= "\"
                        End If

                        If Not ASCMAIN1.Logical_Lock("IMPSVC01", PARTNER_CODE, False, True, True, 1) Then
                            Exit Sub
                        End If

                    End If
                End If
        End Select

        If EMsg <> String.Empty Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Cancel"
                dst.Tables("SOTORDR1").Rows.Clear()
                dst.Tables("SOTORDRV").Rows.Clear()
                grdSOTORDRV.Text = "Sales Order Errors"
                Call Mode_Settings(False)


            Case "Get Sales Orders"

                dst.Tables("SOTORDR1").Rows.Clear()
                dst.Tables("SOTORDRV").Rows.Clear()
                Fill_Records("TATCNTRY", String.Empty, True, "Select * From TATCNTRY")
                skippedSalesOrder = New List(Of String)

                ' Leave the last imported orders on the screen
                ' Clear them out when we go to get new orders
                Select Case MyBase.Absx1.txtFor("ORDR_SOURCE_CODE").Text
                    Case "WUN"
                        GetShopSiteSalesOrders()

                    Case "SHP"
                        GetShopComSalesOrdersXML()

                    Case "AMZ"
                        GetAmazonSalesOrders()

                    Case "BUY"
                        GetBuyComSalesOrders()

                End Select

                'WBCMAIN1.ValidateOrderData(dst.Tables("SOTORDR1"), dst.Tables("SOTORDR2"), dst.Tables("SOTORDR5"), dst.Tables("SOTORDRV"))

                For Each rowDISTINCT As DataRow In ASCDATA1.SelectDistinct("SOTORDRV", "ORDR_NO").Rows
                    dst.Tables("SOTORDR1").Select("ORDR_NO = '" & rowDISTINCT.Item("ORDR_NO") & "'")(0).Item("HAS_ERRORS") = "1"
                Next

                ASCMAIN1.Progress("Updating", String.Empty)
                Update_Record()
                ASCMAIN1.Progress(String.Empty, String.Empty)

                ' Process ShopSite Credit Cards
                If MyBase.Absx1.txtFor("ORDR_SOURCE_CODE").Text = "SHP" Then
                    'ProcessCreditCards()
                    Try
                        MyBase.BeginTrans()

                        Dim ORDR_NO As String = String.Empty
                        Dim sql As String = String.Empty
                        Dim ORDR_QTY As Int16 = 0

                        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(shopCCQuery & " AND ORDR_STATUS = 'O'")
                            ORDR_NO = rowSOTORDR1.Item("ORDR_NO")

                            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
                                ORDR_QTY = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)

                                sql = "Update ICTSTAT2 "
                                sql &= " SET WHSE_QTY_HOLD = NVL(WHSE_QTY_HOLD, 0) - " & ORDR_QTY
                                sql &= " , WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN, 0) + " & ORDR_QTY
                                sql &= " WHERE STYLE_CODE = '" & rowSOTORDR2.Item("STYLE_CODE") & "'"
                                sql &= " AND COLOR_CODE = '" & rowSOTORDR2.Item("COLOR_CODE") & "'"
                                sql &= " AND SIZE_CODE = '" & rowSOTORDR2.Item("SIZE_CODE") & "'"
                                sql &= " AND WHSE_CODE = '" & rowSOTORDR1.Item("WHSE_CODE") & "'"
                                ASCDATA1.ExecuteSQL(sql)
                            Next
                        Next

                        Update_Record_TDA("SOTORDR1")
                        Update_Record_TDA("SOTORDRV")
                        Update_Record_TDA("SOTORDRE")
                        Update_Record_TDA("SOTPAYPL")

                        MyBase.CommitTrans()
                    Catch ex As Exception
                        MyBase.Rollback("Process Shop Site Credit Cards: " & ex.Message)
                    End Try
                End If

                ' As per Maria/Deb 11/2/10 - do not show Amazon duplicates.
                If skippedSalesOrder.Count > 0 _
                        AndAlso MyBase.Absx1.txtFor("ORDR_SOURCE_CODE").Text <> "AMZ" _
                        AndAlso MyBase.Absx1.txtFor("ORDR_SOURCE_CODE").Text <> "BUY" Then

                    For Each orderID As String In skippedSalesOrder
                        EMsg &= ", " & orderID
                    Next
                    EMsg = EMsg.Substring(1).Trim
                    EMsg = "The following sales order were skipped as duplicate partner order numbers: " & EMsg

                    MessageBox.Show(EMsg, "Import", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    EMsg = String.Empty
                End If

                Call Mode_Settings(True)
                ASCMAIN1.MultiTask_Release()

        End Select

        ASCMAIN1.Progress(String.Empty, String.Empty)
        ASCMAIN1.MultiTask_Release()

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("Get Sales Orders").Settings.Enabled = not_iScreenMode
            .Items("Cancel").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        MyBase.EnforceConstraints(False)

        dst.Tables("SOTORDR2").Rows.Clear()
        dst.Tables("SOTORDR5").Rows.Clear()
        dst.Tables("SOTORDRE").Rows.Clear()
        dst.Tables("SOTORDRG").Rows.Clear()
        dst.Tables("SOTPAYPL").Rows.Clear()

        dst.Tables("FILES_PROCESSED").Rows.Clear()
        dst.Tables("AMAZON").Rows.Clear()
        dst.Tables("SHOPCOM1").Rows.Clear()
        dst.Tables("SHOPCOM2").Rows.Clear()

        MyBase.EnforceConstraints(True)

        MyBase.Absx1.txtFor("ORDR_SOURCE_CODE").Clear()

    End Sub

    Sub Load_Record()

    End Sub

    Sub Update_Record()

        Try
            If dst.Tables("SOTORDR1").Rows.Count > 0 Then
                MyBase.BeginTrans()

                MyBase.INIT_LAST("SOTORDR1", True, "", True)
                MyBase.Update_Record_TDA("SOTORDR1")
                MyBase.Update_Record_TDA("SOTORDR2")
                MyBase.Update_Record_TDA("SOTORDR5")
                MyBase.Update_Record_TDA("SOTORDRV")
                MyBase.Update_Record_TDA("SOTORDRG")
                MyBase.Update_Record_TDA("SOTPAYPL")

                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Rows
                    Record_Event(rowSOTORDR1.Item("ORDR_NO"), "Order Imported")
                    'WBCMAIN1.UpdateInventoryOpenHoldPickStatus(rowSOTORDR1.Item("ORDR_NO"), 1)
                Next

                MyBase.Update_Record_TDA("SOTORDRE")

                ' Update the last sales order imported from shopsite
                ASCDATA1.ExecuteSQL("UPDATE SOTPART1 SET PARTNER_LAST_SALES_ORDER = '" & PARTNER_LAST_SALES_ORDER & "' WHERE PARTNER_CODE = '" & PARTNER_CODE & "'")
                MyBase.CommitTrans(dst.Tables("SOTORDR1").Rows.Count & " Sales Orders Imported.")
            Else
                MessageBox.Show("No sales orders to Import.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            For Each rowFILES_PROCESSED As DataRow In dst.Tables("FILES_PROCESSED").Rows
                Dim FileName As String = rowFILES_PROCESSED.Item("FileName")
                Dim FilePath As String = rowFILES_PROCESSED.Item("FilePath")

                My.Computer.FileSystem.MoveFile(FilePath, PARTNER_ORDERS_DIR & "Processed\" & FileName, True)
            Next

        Catch ex As Exception

            MyBase.Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Amazon Procedures"

    Private Sub CreateAmazonDataTable()

        With dst

            dst.Tables.Add("AMAZON")
            With dst.Tables("AMAZON")

                .Columns.Add("ORDER_ID", GetType(System.String))
                .Columns.Add("ORDER_ITEM_ID", GetType(System.String))
                .Columns.Add("PURCHASE_DATE", GetType(System.String))
                .Columns.Add("PAYMENTS_DATE", GetType(System.String))

                .Columns.Add("BUYER_EMAIL", GetType(System.String))
                .Columns.Add("BUYER_NAME", GetType(System.String))
                .Columns.Add("BUYER_PHONE_NUMBER", GetType(System.String))

                .Columns.Add("SKU", GetType(System.String))
                .Columns.Add("PRODUCT_NAME", GetType(System.String))
                .Columns.Add("QUANTITY_PURCHASED", GetType(System.String))
                .Columns.Add("ITEM_PRICE", GetType(System.String))
                .Columns.Add("ITEM_TAX", GetType(System.String))
                .Columns.Add("SHIPPING_PRICE", GetType(System.String))
                .Columns.Add("SHIPPING_TAX", GetType(System.String))
                .Columns.Add("GIFT_WRAP_PRICE", GetType(System.String))
                .Columns.Add("GIFT_WRAP_TAX", GetType(System.String))

                .Columns.Add("SHIP_SERVICE_LEVEL", GetType(System.String))

                .Columns.Add("RECIPIENT_NAME", GetType(System.String))
                .Columns.Add("SHIP_ADDRESS_1", GetType(System.String))
                .Columns.Add("SHIP_ADDRESS_2", GetType(System.String))
                .Columns.Add("SHIP_ADDRESS_3", GetType(System.String))
                .Columns.Add("SHIP_CITY", GetType(System.String))
                .Columns.Add("SHIP_STATE", GetType(System.String))
                .Columns.Add("SHIP_POSTAL_CODE", GetType(System.String))
                .Columns.Add("SHIP_COUNTRY", GetType(System.String))
                .Columns.Add("SHIP_PHONE_NUMBER", GetType(System.String))

                .Columns.Add("GIFT_WRAP_TYPE", GetType(System.String))
                .Columns.Add("GIFT_MESSAGE_TEXT", GetType(System.String))
                .Columns.Add("SALES_CHANNEL", GetType(System.String))
                .Columns.Add("ITEM_PROMOTION_DISCOUNT", GetType(System.String))
                .Columns.Add("ITEM_PROMOTION_ID", GetType(System.String))
                .Columns.Add("SHIP_PROMOTION_DISCOUNT", GetType(System.String))
                .Columns.Add("SHIP_PROMOTION_ID", GetType(System.String))
                .Columns.Add("FILE_NAME", GetType(System.String))
            End With

        End With
    End Sub

    Private Sub GetAmazonSalesOrders()

        dst.Tables("AMAZON").Rows.Clear()

        If PARTNER_SITE_OUTPUT_DIR.Length = 0 Then Exit Sub
        If PARTNER_ORDERS_DIR.Length = 0 Then Exit Sub

        If Not My.Computer.FileSystem.DirectoryExists(PARTNER_SITE_OUTPUT_DIR) Then
            Exit Sub
        End If

        If Not My.Computer.FileSystem.DirectoryExists(PARTNER_ORDERS_DIR) Then
            Exit Sub
        End If

        If Not PARTNER_ORDERS_DIR.EndsWith("\") Then PARTNER_ORDERS_DIR &= "\"

        ' Move files to orders staging area
        For Each amazonSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_SITE_OUTPUT_DIR, FileIO.SearchOption.SearchTopLevelOnly, "order*.txt")
            My.Computer.FileSystem.MoveFile(amazonSalesFile, PARTNER_ORDERS_DIR & My.Computer.FileSystem.GetName(amazonSalesFile))
        Next

        dst.Tables("FILES_PROCESSED").Rows.Clear()
        Dim rowSalesOrders As DataRow = Nothing
        For Each amazonSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly, "order*.txt")
            If ReadFromAmazonFile(amazonSalesFile) Then
                rowSalesOrders = dst.Tables("FILES_PROCESSED").NewRow
                rowSalesOrders.Item("FileName") = My.Computer.FileSystem.GetName(amazonSalesFile)
                rowSalesOrders.Item("FilePath") = amazonSalesFile
                dst.Tables("FILES_PROCESSED").Rows.Add(rowSalesOrders)
            End If
        Next

        CreateAmazonSalesOrders()

    End Sub

    Public Function ReadFromAmazonFile(ByVal amazonFileName As String) As Boolean

        Dim amazonColumnCount As Int16 = dst.Tables("AMAZON").Columns.Count
        Dim currentrow(amazonColumnCount) As String

        Dim boolFileExists As Boolean = False
        Dim rowAMAZON As DataRow = Nothing

        Dim ORDER_ID As String = String.Empty
        Dim ORDER_ITEM_ID As String = String.Empty
        Dim FILE_NAME As String = String.Empty

        Try

            boolFileExists = My.Computer.FileSystem.FileExists(amazonFileName)

            If boolFileExists = False Then
                Return False
            End If

            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(amazonFileName)

                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(vbTab)

                While Not MyReader.EndOfData
                    currentrow = MyReader.ReadFields()

                    ' Ship any header records
                    If currentrow(0).ToString.Trim.ToUpper = "ORDER-ID" Then
                        Continue While
                    End If

                    ORDER_ID = currentrow(0)
                    ORDER_ITEM_ID = currentrow(1)

                    If dst.Tables("AMAZON").Select("ORDER_ID = '" & ORDER_ID & "' AND ORDER_ITEM_ID = '" & ORDER_ITEM_ID & "'").Length > 0 Then
                        FILE_NAME = dst.Tables("AMAZON").Select("ORDER_ID = '" & ORDER_ID & "' AND ORDER_ITEM_ID = '" & ORDER_ITEM_ID & "'")(0).Item("FILE_NAME") & String.Empty
                        If FILE_NAME.ToUpper <> amazonFileName.ToUpper Then
                            Continue While
                        End If
                    End If

                    '@marketplace.amazon.com
                    ' if the buyer email address does not end with the above then skip the order.
                    ' when on shopsite and paying with Amazon the order comes in the Shopsite download and the Amazon download
                    ' this prevents the Amazon import of the order.
                    Dim BUYER_EMAIL As String = currentrow(4) & String.Empty
                    If Not BUYER_EMAIL.ToUpper.Trim.Contains("@marketplace.amazon.com".ToUpper) Then
                        Continue While
                    End If

                    rowAMAZON = dst.Tables("AMAZON").NewRow
                    With rowAMAZON
                        .Item("ORDER_ID") = currentrow(0)
                        .Item("ORDER_ITEM_ID") = currentrow(1)
                        .Item("PURCHASE_DATE") = currentrow(2)
                        .Item("PAYMENTS_DATE") = currentrow(3)

                        .Item("BUYER_EMAIL") = currentrow(4)
                        .Item("BUYER_NAME") = currentrow(5)
                        .Item("BUYER_PHONE_NUMBER") = currentrow(6)

                        .Item("SKU") = currentrow(7)
                        .Item("PRODUCT_NAME") = currentrow(8)
                        .Item("QUANTITY_PURCHASED") = currentrow(9)
                        .Item("ITEM_PRICE") = currentrow(10)
                        .Item("ITEM_TAX") = currentrow(11)
                        .Item("SHIPPING_PRICE") = currentrow(12)
                        .Item("SHIPPING_TAX") = currentrow(13)
                        .Item("GIFT_WRAP_PRICE") = currentrow(14)
                        .Item("GIFT_WRAP_TAX") = currentrow(15)

                        .Item("SHIP_SERVICE_LEVEL") = currentrow(16)

                        .Item("RECIPIENT_NAME") = currentrow(17)
                        .Item("SHIP_ADDRESS_1") = currentrow(18)
                        .Item("SHIP_ADDRESS_2") = currentrow(19)
                        .Item("SHIP_ADDRESS_3") = currentrow(20)
                        .Item("SHIP_CITY") = currentrow(21)
                        .Item("SHIP_STATE") = currentrow(22)
                        .Item("SHIP_POSTAL_CODE") = currentrow(23)
                        .Item("SHIP_COUNTRY") = currentrow(24)
                        .Item("SHIP_PHONE_NUMBER") = currentrow(25)

                        .Item("GIFT_WRAP_TYPE") = currentrow(26)
                        .Item("GIFT_MESSAGE_TEXT") = currentrow(27)
                        .Item("SALES_CHANNEL") = currentrow(28)
                        .Item("ITEM_PROMOTION_DISCOUNT") = currentrow(29)
                        .Item("ITEM_PROMOTION_ID") = currentrow(30)
                        .Item("SHIP_PROMOTION_DISCOUNT") = currentrow(31)
                        .Item("SHIP_PROMOTION_ID") = currentrow(32)
                        .Item("FILE_NAME") = amazonFileName.ToUpper.Trim
                    End With

                    dst.Tables("AMAZON").Rows.Add(rowAMAZON)
                End While
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Amazon Import Error", MessageBoxButtons.OK)
            Return False
        End Try

        Return True

    End Function

    Private Sub CreateAmazonSalesOrders()

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing

        Dim ORDR_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_BATCH_NO")
        Dim ORDR_NO As String = String.Empty
        Dim ORDR_LNO As Int16 = 0
        Dim ORDER_ID As String = String.Empty
        Dim telePhone As String = String.Empty
        Dim custName As String = String.Empty
        Dim maxLength As Int16 = 0

        Dim ORDR_SALES_AMT As Decimal = 0
        Dim ORDR_DISC_AMT As Decimal = 0
        Dim ORDR_DISC_PCT As Decimal = 0
        Dim ORDR_STAX_AMT As Decimal = 0
        Dim ORDR_STAX_RATE As Decimal = 0
        Dim ORDR_FRT_AMT As Decimal = 0
        Dim ORDR_TOT_AMT As Decimal = 0
        Dim SHIP_COUNTRY As String = String.Empty

        Dim ITEM_CODE As String = String.Empty

        Dim rowAMAZON As DataRow = Nothing
        Dim sql As String = String.Empty

        ASCMAIN1.Progress("Importing Amazon Orders", String.Empty)

        ' Get the distinct Order Numbers
        For Each rowAMAZONx As DataRow In ASCDATA1.SelectDistinct("AMAZON", New String() {"ORDER_ID"}).Rows

            ORDER_ID = rowAMAZONx.Item("ORDER_ID") & String.Empty
            ASCMAIN1.Progress("-", ORDER_ID)

            ' See if we have this Amazon Sales Order
            If ORDER_ID.Trim.Length > 0 Then
                sql = "Select * From SOTORDR1 WHERE ORDR_SOURCE_CODE = '" & PARTNER_ORDR_SOURCE_CODE & "' AND PARTNER_ORDR_NO = '" & ORDER_ID & "'"
                If ASCDATA1.GetDataTable(sql).Rows.Count > 0 Then
                    skippedSalesOrder.Add(ORDER_ID)
                    Continue For
                End If
            End If

            ' Get a single record to create the header
            rowAMAZON = dst.Tables("AMAZON").Select("ORDER_ID = '" & ORDER_ID & "'")(0)

            ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")

            ORDR_SALES_AMT = 0
            ORDR_DISC_AMT = 0
            ORDR_DISC_PCT = 0
            ORDR_STAX_AMT = 0
            ORDR_STAX_RATE = 0
            ORDR_FRT_AMT = 0
            ORDR_TOT_AMT = 0

            rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
            rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
            rowSOTORDR1.Item("WHSE_CODE") = "001"
            rowSOTORDR1.Item("ORDR_BATCH_NO") = ORDR_BATCH_NO
            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

            If IsDate(rowAMAZON.Item("PURCHASE_DATE") & String.Empty) Then
                rowSOTORDR1.Item("ORDR_DATE") = CDate(rowAMAZON.Item("PURCHASE_DATE") & String.Empty).ToString("dd-MMM-yyyy")
            Else
                rowSOTORDR1.Item("ORDR_DATE") = DateTime.Now.ToString("dd-MMM-yyyy")
            End If

            rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
            rowSOTORDR1.Item("ORDR_SOURCE_CODE") = PARTNER_ORDR_SOURCE_CODE
            'rowSOTORDR1.Item("ORDR_NO_ORIG") = ORDER_ID
            rowSOTORDR1.Item("PARTNER_ORDR_NO") = ORDER_ID
            'rowSOTORDR1.Item("AFFILIATE_NO") = String.Empty
            'rowSOTORDR1.Item("IP_ADDRESS") = String.Empty
            'rowSOTORDR1.Item("IP_A") = String.Empty
            'rowSOTORDR1.Item("IP_B") =  String.Empty
            'rowSOTORDR1.Item("IP_C") =  String.Empty
            'rowSOTORDR1.Item("IP_D") =  String.Empty
            'rowSOTORDR1.Item("IP_NUMBER") =  String.Empty
            'rowSOTORDR1.Item("IP_COUNTRY") = TruncateField(rowAMAZON.Item("SHIP_COUNTRY") & String.Empty, "SOTORDR1", "IP_COUNTRY")
            'rowSOTORDR1.Item("PICK_NO") =  String.Empty
            rowSOTORDR1.Item("ORDR_STATUS") = "O"
            rowSOTORDR1.Item("SHIP_VIA_ORIG") = (rowAMAZON.Item("SHIP_SERVICE_LEVEL") & String.Empty).ToString.Trim.ToUpper
            'rowSOTORDR1.Item("INV_DATE") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("SHIP_DATE") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("SHIP_REF_NO") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("USPS_ZONE") = rowAMAZON.Item("") & String.Empty
            rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") = 0
            rowSOTORDR1.Item("ORDR_GIFT_MESSAGE") = TruncateField(rowAMAZON.Item("GIFT_MESSAGE_TEXT") & String.Empty, "SOTORDR1", "ORDR_GIFT_MESSAGE")
            'rowSOTORDR1.Item("ORDR_NOTES") = rowAMAZON.Item("") & String.Empty

            'SALES_CHANNEL
            If (rowAMAZON.Item("SALES_CHANNEL") & String.Empty).ToString.Trim.Length > 0 Then
                rowSOTORDR1.Item("REFERRAL") = TruncateField(rowAMAZON.Item("SALES_CHANNEL") & String.Empty, "SOTORDR1", "REFERRAL")
            Else
                rowSOTORDR1.Item("REFERRAL") = "Amazon"
            End If

            'rowSOTORDR1.Item("BAD_CUST_MATCH") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("ADDRESS_TYPE") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_METHOD") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AUTH_SERVICE") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_TYPE") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_CARD_NO") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_EXP_DATE") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_CARD_CVV") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_REF_CD") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AUTH_CD") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AMT") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_CD") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_STREET") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_ZIP") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_CVV2") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_RECD") = rowAMAZON.Item("") & String.Empty

            If IsDate(rowAMAZON.Item("PAYMENTS_DATE") & String.Empty) Then
                rowSOTORDR1.Item("PYMT_RECD_DATE") = CDate(rowAMAZON.Item("PAYMENTS_DATE") & String.Empty).ToString("dd-MMM-yyyy")
                rowSOTORDR1.Item("PYMT_RECD") = "1"
            End If

            'rowSOTORDR1.Item("PYMT_RECD_DATE") = rowAMAZON.Item("") & String.Empty
            'rowSOTORDR1.Item("INIT_OPER") = String.Empty
            'rowSOTORDR1.Item("INIT_DATE") = String.Empty
            'rowSOTORDR1.Item("LAST_OPER") = String.Empty
            'rowSOTORDR1.Item("LAST_DATE") = String.Empty
            rowSOTORDR1.Item("SHIP_VIA_ORIG") = TruncateField(rowAMAZON.Item("SHIP_SERVICE_LEVEL") & String.Empty, "SOTORDR1", "SHIP_VIA_ORIG")
            'rowSOTORDR1.Item("ORDR_INSTR") = rowAMAZON.Item("") & String.Empty
            rowSOTORDR1.Item("PYMT_METHOD_CODE") = "AMZ"
            rowSOTORDR1.Item("PYMT_TYPE_CODE") = "AMZ"

            ' SOTORDR5
            For Each addrType As String In New String() {"BT", "ST"}

                rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR5.Item("CUST_ADDR_TYPE") = addrType
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                'rowSOTORDR5.Item("CUST_FIRST_NAME") = rowAMAZON.Item("") & String.Empty
                'rowSOTORDR5.Item("CUST_LAST_NAME") = rowAMAZON.Item("") & String.Empty

                Select Case addrType

                    Case "BT"
                        custName = rowAMAZON.Item("BUYER_NAME") & String.Empty
                        telePhone = rowAMAZON.Item("BUYER_PHONE_NUMBER") & String.Empty

                    Case "ST"
                        custName = rowAMAZON.Item("RECIPIENT_NAME") & String.Empty
                        telePhone = rowAMAZON.Item("SHIP_PHONE_NUMBER") & String.Empty
                End Select

                custName = TruncateField(custName, "SOTORDR5", "CUST_FULL_NAME")
                telePhone = TruncateField(telePhone, "SOTORDR5", "CUST_PHONE")

                rowSOTORDR5.Item("CUST_FULL_NAME") = custName
                rowSOTORDR5.Item("CUST_ADDR1") = TruncateField(rowAMAZON.Item("SHIP_ADDRESS_1") & String.Empty, "SOTORDR5", "CUST_ADDR1")
                rowSOTORDR5.Item("CUST_ADDR2") = TruncateField(rowAMAZON.Item("SHIP_ADDRESS_2") & String.Empty, "SOTORDR5", "CUST_ADDR2")
                rowSOTORDR5.Item("CUST_ADDR3") = TruncateField(rowAMAZON.Item("SHIP_ADDRESS_3") & String.Empty, "SOTORDR5", "CUST_ADDR3")
                rowSOTORDR5.Item("CUST_CITY") = TruncateField(rowAMAZON.Item("SHIP_CITY") & String.Empty, "SOTORDR5", "CUST_CITY")
                rowSOTORDR5.Item("CUST_STATE") = TruncateField(ConvertState(rowAMAZON.Item("SHIP_STATE") & String.Empty), "SOTORDR5", "CUST_STATE")
                rowSOTORDR5.Item("CUST_ZIP_CODE") = TruncateField(rowAMAZON.Item("SHIP_POSTAL_CODE") & String.Empty, "SOTORDR5", "CUST_ZIP_CODE")
                rowSOTORDR5.Item("CUST_COUNTRY") = TruncateField(ConvertCountry(rowAMAZON.Item("SHIP_COUNTRY") & String.Empty), "SOTORDR5", "CUST_COUNTRY")
                rowSOTORDR5.Item("CUST_CONTACT") = rowSOTORDR5.Item("CUST_FULL_NAME") & String.Empty
                rowSOTORDR5.Item("CUST_PHONE") = telePhone
                'rowSOTORDR5.Item("CUST_EXT") = rowAMAZON.Item("") & String.Empty
                'rowSOTORDR5.Item("CUST_FAX") = rowAMAZON.Item("") & String.Empty
                rowSOTORDR5.Item("CUST_EMAIL") = TruncateField(rowAMAZON.Item("BUYER_EMAIL"), "SOTORDR5", "CUST_EMAIL")
                'rowSOTORDR5.Item("CUST_ZIP_MATCH") = rowAMAZON.Item("") & String.Empty
                'rowSOTORDR5.Item("CUST_COMPANY_NAME") = rowAMAZON.Item("") & String.Empty
            Next

            ORDR_LNO = 0
            For Each rowDetails As DataRow In dst.Tables("AMAZON").Select("ORDER_ID = '" & ORDER_ID & "'", "ORDER_ITEM_ID")
                ORDR_LNO += 1

                rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR2.Item("ORDR_LNO") = ORDR_LNO
                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                ITEM_CODE = (rowDetails.Item("SKU") & String.Empty).ToString.Trim.ToUpper
                rowSOTORDR2.Item("ITEM_CODE") = ITEM_CODE
                rowSOTORDR2.Item("ITEM_DESC") = TruncateField(rowDetails.Item("PRODUCT_NAME") & String.Empty, "SOTORDR2", "ITEM_DESC")
                UpdateItemInfo(ITEM_CODE, rowSOTORDR2)

                'rowSOTORDR2.Item("ORDR_PRICE_SOURCE") = TruncateField(rowDetails.Item("") & String.Empty, "SOTORDR2", "ORDR_PRICE_SOURCE")

                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Val(rowDetails.Item("ITEM_PRICE") & String.Empty)
                rowSOTORDR2.Item("ORDR_QTY") = Val(rowDetails.Item("QUANTITY_PURCHASED") & String.Empty)

                ' Amazon Item Price is tht Total Line Price
                If rowSOTORDR2.Item("ORDR_QTY") > 1 Then
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE") /= rowSOTORDR2.Item("ORDR_QTY")
                End If

                rowSOTORDR2.Item("ORDR_EXT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE") * rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0 'rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
                rowSOTORDR2.Item("ORDR_QTY_CANC") = 0
                rowSOTORDR2.Item("ORDR_QTY_ORIG") = rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("UNIT_WEIGHT") = 0
                rowSOTORDR2.Item("PARTNER_LN_ID") = TruncateField(rowDetails.Item("ORDER_ITEM_ID") & String.Empty, "SOTORDR2", "PARTNER_LN_ID")

                ORDR_SALES_AMT += Val(rowDetails.Item("GIFT_WRAP_PRICE") & String.Empty) + rowSOTORDR2.Item("ORDR_EXT_PRICE")

                ORDR_DISC_AMT += Val(rowDetails.Item("ITEM_PROMOTION_DISCOUNT") & String.Empty) + Val(rowDetails.Item("SHIP_PROMOTION_DISCOUNT") & String.Empty)
                ORDR_DISC_AMT = Math.Abs(ORDR_DISC_AMT) * -1

                ORDR_DISC_PCT += 0
                ORDR_STAX_AMT += Val(rowDetails.Item("ITEM_TAX") & String.Empty) + Val(rowDetails.Item("SHIPPING_TAX") & String.Empty) + Val(rowDetails.Item("GIFT_WRAP_TAX") & String.Empty)
                ORDR_STAX_RATE += 0
                ORDR_FRT_AMT += Val(rowDetails.Item("SHIPPING_PRICE") & String.Empty)
            Next

            ORDR_TOT_AMT = ORDR_SALES_AMT + ORDR_DISC_AMT + ORDR_STAX_AMT + ORDR_FRT_AMT

            rowSOTORDR1.Item("ORDR_SALES_AMT") = ORDR_SALES_AMT
            rowSOTORDR1.Item("ORDR_COGS_AMT") = 0
            rowSOTORDR1.Item("ORDR_DISC_AMT") = ORDR_DISC_AMT
            rowSOTORDR1.Item("ORDR_DISC_PCT") = ORDR_DISC_PCT
            rowSOTORDR1.Item("ORDR_STAX_AMT") = ORDR_STAX_AMT
            rowSOTORDR1.Item("ORDR_STAX_RATE") = ORDR_STAX_RATE
            rowSOTORDR1.Item("ORDR_FRT_AMT") = ORDR_FRT_AMT
            rowSOTORDR1.Item("ORDR_TOT_AMT") = ORDR_TOT_AMT
            rowSOTORDR1.Item("ORDR_TOT_WT") = 0

        Next
    End Sub

#End Region

#Region "Shop.Com"

    Private Sub CreateShopComDataTable()

        With dst

            dst.Tables.Add("SHOPCOM1")
            With dst.Tables("SHOPCOM1")

                .Columns.Add("ORDER_ID", GetType(System.String))
                .Columns.Add("INVOICE_ID", GetType(System.String))
                .Columns.Add("ORDER_DATE", GetType(System.String))
                '.Columns.Add("EMAIL", GetType(System.String))
                .Columns.Add("SHOPPER_ID", GetType(System.String))
                .Columns.Add("IP_ADDRESS", GetType(System.String))

                .Columns.Add("BT_FIRST_NAME", GetType(System.String))
                .Columns.Add("BT_LAST_NAME", GetType(System.String))
                .Columns.Add("BT_COMPANY_NAME", GetType(System.String))
                .Columns.Add("BT_STREET1", GetType(System.String))
                .Columns.Add("BT_STREET2", GetType(System.String))
                .Columns.Add("BT_CITY", GetType(System.String))
                .Columns.Add("BT_STATE", GetType(System.String))
                .Columns.Add("BT_ZIP", GetType(System.String))
                .Columns.Add("BT_COUNTRY", GetType(System.String))
                .Columns.Add("BT_REGION", GetType(System.String))
                .Columns.Add("BT_TELEPHONE", GetType(System.String))
                .Columns.Add("BT_EMAIL", GetType(System.String))

                .Columns.Add("DELIVERY_METHOD", GetType(System.String))
                .Columns.Add("SUB_TOTAL", GetType(System.String))
                .Columns.Add("FREIGHT", GetType(System.String))
                .Columns.Add("TAX_MULT", GetType(System.String))
                .Columns.Add("TAX", GetType(System.String))
                .Columns.Add("DISCOUNT", GetType(System.String))
                .Columns.Add("TOTAL", GetType(System.String))

                .Columns.Add("CC_TYPE", GetType(System.String))
                .Columns.Add("CC_NUMBER", GetType(System.String))
                .Columns.Add("CC_EXP", GetType(System.String))
                .Columns.Add("CC_CCV", GetType(System.String))
                .Columns.Add("NAME_ON_CC", GetType(System.String))

                .Columns.Add("ST_FIRST_NAME", GetType(System.String))
                .Columns.Add("ST_LAST_NAME", GetType(System.String))
                .Columns.Add("ST_COMPANY_NAME", GetType(System.String))
                .Columns.Add("ST_STREET1", GetType(System.String))
                .Columns.Add("ST_STREET2", GetType(System.String))
                .Columns.Add("ST_CITY", GetType(System.String))
                .Columns.Add("ST_STATE", GetType(System.String))
                .Columns.Add("ST_ZIP", GetType(System.String))
                .Columns.Add("ST_COUNTRY", GetType(System.String))
                .Columns.Add("ST_REGION", GetType(System.String))
                .Columns.Add("ST_TELEPHONE", GetType(System.String))
                .Columns.Add("ST_EMAIL", GetType(System.String))

                .Columns.Add("CATALOG_ID", GetType(System.String))
                .Columns.Add("CATALOG_NAME", GetType(System.String))
                .Columns.Add("MULT_PAYMENT_QTY", GetType(System.String))
                .Columns.Add("CAN_SELL_NAME", GetType(System.String))
                .Columns.Add("CAN_SEND_OFFERS", GetType(System.String))
                .Columns.Add("COMMENTS", GetType(System.String))
            End With

            dst.Tables.Add("SHOPCOM2")
            With dst.Tables("SHOPCOM2")
                .Columns.Add("INVOICE_ID", GetType(System.String))
                .Columns.Add("PURCHASE_ID", GetType(System.String))
                .Columns.Add("VOLUME_ID", GetType(System.String))
                .Columns.Add("VOLUME_NAME", GetType(System.String))
                .Columns.Add("SOURCE_CODE", GetType(System.String))
                .Columns.Add("PRODUCT_SKU", GetType(System.String))
                .Columns.Add("PRODUCT_DESC", GetType(System.String))
                .Columns.Add("QUANTITY", GetType(System.String))
                .Columns.Add("UNIT_PRICE", GetType(System.String))
                .Columns.Add("EXTENDED_PRICE", GetType(System.String))
                .Columns.Add("COUPON_CODE", GetType(System.String))
            End With
        End With
    End Sub

    Private Sub GetShopComSalesOrdersCSV()

        If PARTNER_SITE_IP.Length = 0 Then Exit Sub
        If PARTNER_SITE_USER.Length = 0 Then Exit Sub
        If PARTNER_SITE_PWD.Length = 0 Then Exit Sub
        If PARTNER_SITE_OUTPUT_DIR.Length = 0 Then Exit Sub
        If PARTNER_ORDERS_DIR.Length = 0 Then Exit Sub

        Dim localFilename As String = String.Empty

        If Not My.Computer.FileSystem.DirectoryExists(PARTNER_ORDERS_DIR) Then
            Exit Sub
        End If

        If Not PARTNER_ORDERS_DIR.EndsWith("\") Then PARTNER_ORDERS_DIR &= "\"

        ' FTP file dowm from Shop.Com
        Try
            ASCMAIN1.Progress("Creating FTP Connection to Shop.com", "")

            Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")
            ASCMAIN1.Progress("-", "RemoteHost")

            ASCMAIN1.Progress("-", "User")
            Sftp1.SSHUser = PARTNER_SITE_USER

            ASCMAIN1.Progress("-", "Password")
            Sftp1.SSHPassword = PARTNER_SITE_PWD

            ASCMAIN1.Progress("-", "RemoteFile")
            Sftp1.RemoteFile = String.Empty

            ASCMAIN1.Progress("-", "Timeout")
            Sftp1.Timeout = 300

            ASCMAIN1.Progress("-", "Logon")
            Sftp1.SSHAuthMode = nsoftware.IPWorksSSH.SftpSSHAuthModes.amPassword
            Try
                Sftp1.SSHLogoff()
                Sftp1.SSHLogon(PARTNER_SITE_IP, 22)
            Catch ex As Exception
                Sftp1.SSHLogoff()
                Sftp1.SSHLogon(PARTNER_SITE_IP, 22)
            End Try

            Sftp1.RemotePath = PARTNER_SITE_OUTPUT_DIR
            ftpFileList = New List(Of String)
            Sftp1.ListDirectory()

            For Each fileFtp As String In ftpFileList

                ASCMAIN1.Progress("Downloading: " & fileFtp, String.Empty)
                Sftp1.RemoteFile = fileFtp

                localFilename = fileFtp

                If localFilename.EndsWith(".shop.txt") Then
                    localFilename = localFilename.Replace(".shop.txt", ".csv")
                ElseIf localFilename.EndsWith(".shop") Then
                    localFilename = localFilename.Replace(".shop", ".xml")
                End If

                Sftp1.LocalFile = PARTNER_ORDERS_DIR & localFilename
                Sftp1.Download()
                Sftp1.DeleteFile(fileFtp)

            Next

        Catch ex As Exception
            MessageBox.Show("Error downloading order files: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Sftp1.SSHLogoff()
            Sftp1.Dispose()

        End Try

        dst.Tables("FILES_PROCESSED").Rows.Clear()
        Dim rowSalesOrders As DataRow = Nothing

        ' Need to create a schema.ini file so the data can be put into a temp datatable where the 
        ' data file does not have a header row (column headings)
        Using rs As New StreamWriter(PARTNER_ORDERS_DIR & "schema.ini", False)
            For Each shopcomSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly, "*.csv")
                rs.WriteLine("[" & My.Computer.FileSystem.GetName(shopcomSalesFile) & "]")
                rs.WriteLine("ColNameHeader = False")
                rs.WriteLine("Format = Delimited(,)")
                rs.WriteLine("")
            Next
            rs.Close()
        End Using

        For Each shopcomSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly, "*.csv")
            ASCMAIN1.Progress("Importing: " & shopcomSalesFile, String.Empty)
            If ReadFromShopComFileCSV(shopcomSalesFile) Then
                rowSalesOrders = dst.Tables("FILES_PROCESSED").NewRow
                rowSalesOrders.Item("FileName") = My.Computer.FileSystem.GetName(shopcomSalesFile)
                rowSalesOrders.Item("FilePath") = shopcomSalesFile
                dst.Tables("FILES_PROCESSED").Rows.Add(rowSalesOrders)
            End If
        Next

        ' grab the XML files an archive them as well, not currently used
        For Each shopcomSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly, "*.shop")
            ASCMAIN1.Progress("Importing: " & shopcomSalesFile, String.Empty)
            rowSalesOrders = dst.Tables("FILES_PROCESSED").NewRow
            rowSalesOrders.Item("FileName") = My.Computer.FileSystem.GetName(shopcomSalesFile)
            rowSalesOrders.Item("FilePath") = shopcomSalesFile
            dst.Tables("FILES_PROCESSED").Rows.Add(rowSalesOrders)
        Next

        CreateShopComSalesOrders()

    End Sub

    Private Sub GetShopComSalesOrdersXML()

        If PARTNER_SITE_IP.Length = 0 Then Exit Sub
        If PARTNER_SITE_USER.Length = 0 Then Exit Sub
        If PARTNER_SITE_PWD.Length = 0 Then Exit Sub
        If PARTNER_SITE_OUTPUT_DIR.Length = 0 Then Exit Sub
        If PARTNER_ORDERS_DIR.Length = 0 Then Exit Sub

        Dim localFilename As String = String.Empty

        If Not My.Computer.FileSystem.DirectoryExists(PARTNER_ORDERS_DIR) Then
            Exit Sub
        End If

        If Not PARTNER_ORDERS_DIR.EndsWith("\") Then PARTNER_ORDERS_DIR &= "\"

        ' FTP file dowm from Shop.Com
        Try
            ASCMAIN1.Progress("Creating FTP Connection to Shop.com", "")

            Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")
            ASCMAIN1.Progress("-", "RemoteHost")

            ASCMAIN1.Progress("-", "User")
            Sftp1.SSHUser = PARTNER_SITE_USER

            ASCMAIN1.Progress("-", "Password")
            Sftp1.SSHPassword = PARTNER_SITE_PWD

            ASCMAIN1.Progress("-", "RemoteFile")
            Sftp1.RemoteFile = String.Empty

            ASCMAIN1.Progress("-", "Timeout")
            Sftp1.Timeout = 300

            ASCMAIN1.Progress("-", "Logon")
            Sftp1.SSHAuthMode = nsoftware.IPWorksSSH.SftpSSHAuthModes.amPassword
            Try
                Sftp1.SSHLogoff()
                Sftp1.SSHLogon(PARTNER_SITE_IP, 22)
            Catch ex As Exception
                Sftp1.SSHLogoff()
                Sftp1.SSHLogon(PARTNER_SITE_IP, 22)
            End Try

            Sftp1.RemotePath = PARTNER_SITE_OUTPUT_DIR
            ftpFileList = New List(Of String)
            Sftp1.ListDirectory()

            For Each fileFtp As String In ftpFileList

                ASCMAIN1.Progress("Downloading: " & fileFtp, String.Empty)
                Sftp1.RemoteFile = fileFtp

                localFilename = fileFtp

                If localFilename.EndsWith(".shop.txt") Then
                    localFilename = localFilename.Replace(".shop.txt", ".csv")
                ElseIf localFilename.EndsWith(".shop") Then
                    localFilename = localFilename.Replace(".shop", ".xml")
                End If

                Sftp1.LocalFile = PARTNER_ORDERS_DIR & localFilename
                Sftp1.Download()
                Sftp1.DeleteFile(fileFtp)

            Next

        Catch ex As Exception
            MessageBox.Show("Error downloading order files: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Sftp1.SSHLogoff()
            Sftp1.Dispose()

        End Try

        dst.Tables("FILES_PROCESSED").Rows.Clear()
        Dim rowSalesOrders As DataRow = Nothing

        ' grab the CSV files an archive them as well, not currently used
        For Each shopcomSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly, "*.csv")
            ASCMAIN1.Progress("Archiving: " & shopcomSalesFile, String.Empty)
            'If ReadFromShopComFileCSV(shopcomSalesFile) Then
            rowSalesOrders = dst.Tables("FILES_PROCESSED").NewRow
            rowSalesOrders.Item("FileName") = My.Computer.FileSystem.GetName(shopcomSalesFile)
            rowSalesOrders.Item("FilePath") = shopcomSalesFile
            dst.Tables("FILES_PROCESSED").Rows.Add(rowSalesOrders)
            'End If
        Next

        ' grab the XML files an archive them as well, not currently used
        For Each shopcomSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")
            ASCMAIN1.Progress("Importing: " & shopcomSalesFile, String.Empty)
            If ReadFromShopComFileXML(shopcomSalesFile) Then
                rowSalesOrders = dst.Tables("FILES_PROCESSED").NewRow
                rowSalesOrders.Item("FileName") = My.Computer.FileSystem.GetName(shopcomSalesFile)
                rowSalesOrders.Item("FilePath") = shopcomSalesFile
                dst.Tables("FILES_PROCESSED").Rows.Add(rowSalesOrders)
            End If
        Next

        CreateShopComSalesOrders()

    End Sub


    Private Function ReadFromShopComFileCSV(ByVal shopcomFileName As String) As Boolean

        Dim dt As DataTable = New DataTable
        Dim connectionString As String

        Dim importFolder As String
        Dim strFileName As String

        strFileName = My.Computer.FileSystem.GetName(shopcomFileName)
        Dim filePath As New System.IO.DirectoryInfo(shopcomFileName)
        importFolder = filePath.Parent.FullName

        connectionString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & importFolder & ";"
        Using conn As System.Data.Odbc.OdbcConnection = New Odbc.OdbcConnection(connectionString)
            Using da As System.Data.Odbc.OdbcDataAdapter = New System.Data.Odbc.OdbcDataAdapter("Select * from [" & strFileName & "]", conn)
                da.Fill(dt)
            End Using
        End Using

        For Each row As DataRow In dt.Rows

            Select Case row.Item("F1") & String.Empty

                Case "H"

                    Dim rowSHOPCOM1 As DataRow = dst.Tables("SHOPCOM1").NewRow

                    rowSHOPCOM1.Item("ORDER_ID") = row.Item("F2") & String.Empty
                    rowSHOPCOM1.Item("INVOICE_ID") = row.Item("F3") & String.Empty
                    rowSHOPCOM1.Item("ORDER_DATE") = row.Item("F4") & String.Empty
                    rowSHOPCOM1.Item("BT_EMAIL") = row.Item("F5") & String.Empty
                    rowSHOPCOM1.Item("SHOPPER_ID") = row.Item("F6") & String.Empty

                    rowSHOPCOM1.Item("BT_FIRST_NAME") = row.Item("F7") & String.Empty
                    rowSHOPCOM1.Item("BT_LAST_NAME") = row.Item("F8") & String.Empty
                    rowSHOPCOM1.Item("BT_COMPANY_NAME") = row.Item("F9") & String.Empty
                    rowSHOPCOM1.Item("BT_STREET1") = row.Item("F10") & String.Empty
                    rowSHOPCOM1.Item("BT_STREET2") = row.Item("F11") & String.Empty
                    rowSHOPCOM1.Item("BT_CITY") = row.Item("F12") & String.Empty
                    rowSHOPCOM1.Item("BT_STATE") = row.Item("F13") & String.Empty
                    rowSHOPCOM1.Item("BT_ZIP") = row.Item("F14") & String.Empty
                    rowSHOPCOM1.Item("BT_COUNTRY") = row.Item("F15") & String.Empty
                    rowSHOPCOM1.Item("BT_REGION") = row.Item("F16") & String.Empty
                    rowSHOPCOM1.Item("BT_TELEPHONE") = row.Item("F17") & String.Empty

                    rowSHOPCOM1.Item("DELIVERY_METHOD") = row.Item("F18") & String.Empty
                    rowSHOPCOM1.Item("SUB_TOTAL") = row.Item("F19") & String.Empty
                    rowSHOPCOM1.Item("FREIGHT") = row.Item("F20") & String.Empty
                    rowSHOPCOM1.Item("TAX_MULT") = row.Item("F21") & String.Empty
                    rowSHOPCOM1.Item("TAX") = row.Item("F22") & String.Empty
                    rowSHOPCOM1.Item("DISCOUNT") = row.Item("F23") & String.Empty
                    rowSHOPCOM1.Item("TOTAL") = row.Item("F24") & String.Empty

                    rowSHOPCOM1.Item("CC_TYPE") = row.Item("F25") & String.Empty
                    rowSHOPCOM1.Item("CC_NUMBER") = row.Item("F26") & String.Empty
                    rowSHOPCOM1.Item("CC_EXP") = row.Item("F27") & String.Empty
                    rowSHOPCOM1.Item("CC_CCV") = row.Item("F28") & String.Empty
                    rowSHOPCOM1.Item("NAME_ON_CC") = row.Item("F29") & String.Empty

                    rowSHOPCOM1.Item("ST_FIRST_NAME") = row.Item("F30") & String.Empty
                    rowSHOPCOM1.Item("ST_LAST_NAME") = row.Item("F31") & String.Empty
                    rowSHOPCOM1.Item("ST_COMPANY_NAME") = row.Item("F32") & String.Empty
                    rowSHOPCOM1.Item("ST_STREET1") = row.Item("F33") & String.Empty
                    rowSHOPCOM1.Item("ST_STREET2") = row.Item("F34") & String.Empty
                    rowSHOPCOM1.Item("ST_CITY") = row.Item("F35") & String.Empty
                    rowSHOPCOM1.Item("ST_STATE") = row.Item("F36") & String.Empty
                    rowSHOPCOM1.Item("ST_ZIP") = row.Item("F37") & String.Empty
                    rowSHOPCOM1.Item("ST_COUNTRY") = row.Item("F38") & String.Empty
                    rowSHOPCOM1.Item("ST_REGION") = row.Item("F39") & String.Empty
                    rowSHOPCOM1.Item("ST_TELEPHONE") = row.Item("F40") & String.Empty

                    rowSHOPCOM1.Item("CATALOG_ID") = row.Item("F41") & String.Empty
                    rowSHOPCOM1.Item("CATALOG_NAME") = row.Item("F42") & String.Empty
                    rowSHOPCOM1.Item("MULT_PAYMENT_QTY") = row.Item("F43") & String.Empty
                    rowSHOPCOM1.Item("CAN_SELL_NAME") = row.Item("F44") & String.Empty
                    rowSHOPCOM1.Item("CAN_SEND_OFFERS") = row.Item("F45") & String.Empty
                    rowSHOPCOM1.Item("COMMENTS") = row.Item("F46") & String.Empty

                    dst.Tables("SHOPCOM1").Rows.Add(rowSHOPCOM1)

                Case "D"

                    Dim rowSHOPCOM2 As DataRow = dst.Tables("SHOPCOM2").NewRow

                    rowSHOPCOM2.Item("INVOICE_ID") = row.Item("F2") & String.Empty
                    rowSHOPCOM2.Item("PURCHASE_ID") = row.Item("F3") & String.Empty
                    rowSHOPCOM2.Item("VOLUME_ID") = row.Item("F4") & String.Empty
                    rowSHOPCOM2.Item("VOLUME_NAME") = row.Item("F5") & String.Empty
                    rowSHOPCOM2.Item("SOURCE_CODE") = row.Item("F6") & String.Empty
                    rowSHOPCOM2.Item("PRODUCT_SKU") = row.Item("F7") & String.Empty
                    rowSHOPCOM2.Item("PRODUCT_DESC") = row.Item("F8") & String.Empty
                    rowSHOPCOM2.Item("QUANTITY") = row.Item("F9") & String.Empty
                    rowSHOPCOM2.Item("UNIT_PRICE") = row.Item("F10") & String.Empty
                    rowSHOPCOM2.Item("EXTENDED_PRICE") = row.Item("F11") & String.Empty
                    rowSHOPCOM2.Item("COUPON_CODE") = row.Item("F12") & String.Empty

                    dst.Tables("SHOPCOM2").Rows.Add(rowSHOPCOM2)
                Case "C"
                    ' Nothing at this time
            End Select

        Next

        Return True

    End Function

    Private Function ReadFromShopComFileXML(ByVal shopcomFileName As String) As Boolean

        Dim xReader As XmlTextReader = New XmlTextReader(shopcomFileName)
        Dim rowSHOPCOM1 As DataRow = Nothing
        Dim addressPrefix As String = String.Empty
        Dim rowSHOPCOM2 As DataRow = Nothing
        Dim INVOICE_ID As String = String.Empty
        Dim CATALOG_ID As String = String.Empty

        Do While xReader.Read()

            Select Case xReader.NodeType

                Case XmlNodeType.Element ' node is an element 

                    Select Case xReader.Name
                        Case "CC_TRANSMISSION"
                            If xReader.AttributeCount > 0 Then
                                While xReader.MoveToNextAttribute()
                                    Select Case xReader.Name
                                        Case "CATALOG_ID"
                                            CATALOG_ID = xReader.Value
                                    End Select
                                End While
                            End If

                        Case "BILLING_LABEL"
                            addressPrefix = "BT"
                        Case "SHIPPING_LABEL"
                            addressPrefix = "ST"

                        Case "CC_ORDER"
                            rowSHOPCOM1 = dst.Tables("SHOPCOM1").NewRow
                            rowSHOPCOM1.Item("CATALOG_ID") = CATALOG_ID
                            If xReader.AttributeCount > 0 Then
                                While xReader.MoveToNextAttribute()
                                    Select Case xReader.Name
                                        Case "INVOICE_NO"
                                            INVOICE_ID = xReader.Value
                                            rowSHOPCOM1.Item("INVOICE_ID") = INVOICE_ID
                                        Case "ORDER_NO"
                                            rowSHOPCOM1.Item("ORDER_ID") = xReader.Value
                                        Case "IP_ADDRESS"
                                            rowSHOPCOM1.Item("IP_ADDRESS") = xReader.Value
                                    End Select
                                End While
                            End If
                            dst.Tables("SHOPCOM1").Rows.Add(rowSHOPCOM1)

                            ' Credit Card Fields
                        Case "CC_TYPE"
                            Dim ccType As String = xReader.ReadElementContentAsString
                            Select Case ccType.ToUpper.Substring(0, 3)
                                Case "AME" : rowSHOPCOM1.Item("CC_TYPE") = "1"
                                Case "DIN" : rowSHOPCOM1.Item("CC_TYPE") = "4"
                                Case "DIS" : rowSHOPCOM1.Item("CC_TYPE") = "5"
                                Case "MAS" : rowSHOPCOM1.Item("CC_TYPE") = "6"
                                Case "VIS" : rowSHOPCOM1.Item("CC_TYPE") = "9"
                            End Select
                        Case "CC_NUMBER"
                            rowSHOPCOM1.Item("CC_NUMBER") = xReader.ReadElementContentAsString()
                        Case "CC_EXPIRATION"
                            Dim CC_EXPIRATION As String = String.Empty
                            For Each chNumber As Char In xReader.ReadElementContentAsString()
                                If Char.IsDigit(chNumber) Then
                                    CC_EXPIRATION &= chNumber
                                End If
                            Next
                            rowSHOPCOM1.Item("CC_EXP") = CC_EXPIRATION
                        Case "CC_SECURITY_NUMBER"
                            rowSHOPCOM1.Item("CC_CCV") = xReader.ReadElementContentAsString()
                        Case "CC_NAMEONCARD"
                            rowSHOPCOM1.Item("NAME_ON_CC") = StrConv(xReader.ReadElementContentAsString(), VbStrConv.ProperCase)


                        Case "CUSTOMER"
                            If xReader.AttributeCount > 0 AndAlso addressPrefix = "BT" Then
                                While xReader.MoveToNextAttribute()
                                    Select Case xReader.Name
                                        Case "OK_CONTACT"
                                            rowSHOPCOM1.Item("CAN_SEND_OFFERS") = IIf(xReader.Value.ToUpper = "FALSE", "N", "Y")
                                        Case "OK_RENT"
                                            rowSHOPCOM1.Item("CAN_SELL_NAME") = IIf(xReader.Value.ToUpper = "FALSE", "N", "Y")
                                    End Select
                                End While
                            End If

                            ' Billiing Information
                        Case "CU_FIRST_NAME"
                            rowSHOPCOM1.Item(addressPrefix & "_FIRST_NAME") = StrConv(xReader.ReadElementContentAsString(), VbStrConv.ProperCase)
                        Case "CU_LAST_NAME"
                            rowSHOPCOM1.Item(addressPrefix & "_LAST_NAME") = StrConv(xReader.ReadElementContentAsString(), VbStrConv.ProperCase)
                        Case "CU_COMPANY"
                            rowSHOPCOM1.Item(addressPrefix & "_COMPANY_NAME") = StrConv(xReader.ReadElementContentAsString(), VbStrConv.ProperCase)
                        Case "CU_EMAIL"
                            rowSHOPCOM1.Item(addressPrefix & "_EMAIL") = xReader.ReadElementContentAsString()
                        Case "CU_PHONE"
                            rowSHOPCOM1.Item(addressPrefix & "_TELEPHONE") = xReader.ReadElementContentAsString()
                        Case "CU_SHOPPER_ID"
                            If addressPrefix = "BT" Then rowSHOPCOM1.Item("SHOPPER_ID") = xReader.ReadElementContentAsString()
                        Case "AD_ADDRESS1"
                            rowSHOPCOM1.Item(addressPrefix & "_STREET1") = xReader.ReadElementContentAsString()
                        Case "AD_FLAT"
                            rowSHOPCOM1.Item(addressPrefix & "_STREET2") = xReader.ReadElementContentAsString()
                        Case "AD_ADDRESS2"
                            rowSHOPCOM1.Item(addressPrefix & "_STREET2") = xReader.ReadElementContentAsString()
                        Case "AD_CITY"
                            rowSHOPCOM1.Item(addressPrefix & "_CITY") = StrConv(xReader.ReadElementContentAsString(), VbStrConv.ProperCase)
                        Case "AD_STATE"
                            rowSHOPCOM1.Item(addressPrefix & "_STATE") = xReader.ReadElementContentAsString()
                        Case "AD_ZIP"
                            rowSHOPCOM1.Item(addressPrefix & "_ZIP") = xReader.ReadElementContentAsString()
                        Case "AD_COUNTRY"
                            rowSHOPCOM1.Item(addressPrefix & "_COUNTRY") = xReader.ReadElementContentAsString()
                        Case "AD_COUNTRY_CODE"
                            rowSHOPCOM1.Item(addressPrefix & "_REGION") = xReader.ReadElementContentAsString()
                        Case "AD_PROVINCE"
                            Dim province As String = xReader.ReadElementContentAsString().ToString.Trim
                            If province.Length > 0 AndAlso (rowSHOPCOM1.Item(addressPrefix & "_STATE") & String.Empty).ToString.Trim.Length = 0 Then
                                rowSHOPCOM1.Item(addressPrefix & "_CITY") = xReader.ReadElementContentAsString()
                            End If

                            ' Order total Information
                        Case "TL_ORDER_DATE"
                            rowSHOPCOM1.Item("ORDER_DATE") = CDate(xReader.ReadElementContentAsString()).ToString("MM/dd/yyyy")
                        Case "TL_SUBTOTAL"
                            rowSHOPCOM1.Item("SUB_TOTAL") = xReader.ReadElementContentAsString().Replace("$", "")

                        Case "TL_TAX"
                            rowSHOPCOM1.Item("TAX") = Val(rowSHOPCOM1.Item("TAX") & String.Empty) + Val(xReader.ReadElementContentAsString().Replace("$", ""))
                        Case "TL_EXCISETAX"
                            rowSHOPCOM1.Item("TAX") = Val(rowSHOPCOM1.Item("TAX") & String.Empty) + Val(xReader.ReadElementContentAsString().Replace("$", ""))

                        Case "TL_SHIPPING"
                            rowSHOPCOM1.Item("FREIGHT") = xReader.ReadElementContentAsString().Replace("$", "")
                        Case "TL_TOTAL"
                            rowSHOPCOM1.Item("TOTAL") = xReader.ReadElementContentAsString().Replace("$", "")
                        Case "TL_TAX_RATE"
                            rowSHOPCOM1.Item("TAX_MULT") = xReader.ReadElementContentAsString().Replace("$", "")
                        Case "SL_METHOD"
                            rowSHOPCOM1.Item("DELIVERY_METHOD") = xReader.ReadElementContentAsString()
                        Case "MULTIPLE_PAYMENTS_QTY"
                            rowSHOPCOM1.Item("MULT_PAYMENT_QTY") = xReader.ReadElementContentAsString()
                        Case "SHOPPER_COMMENTS"
                            rowSHOPCOM1.Item("COMMENTS") = xReader.ReadElementContentAsString()


                            ' Order details
                        Case "ITEM"
                            rowSHOPCOM2 = dst.Tables("SHOPCOM2").NewRow
                            rowSHOPCOM2.Item("INVOICE_ID") = INVOICE_ID
                            dst.Tables("SHOPCOM2").Rows.Add(rowSHOPCOM2)
                        Case "IT_PURCHASE_ID"
                            rowSHOPCOM2.Item("PURCHASE_ID") = xReader.ReadElementContentAsString()
                        Case "IT_VID"
                            rowSHOPCOM2.Item("VOLUME_ID") = xReader.ReadElementContentAsString()
                        Case "IT_SKU"
                            rowSHOPCOM2.Item("PRODUCT_SKU") = xReader.ReadElementContentAsString()
                        Case "IT_DESCRIPTION"
                            rowSHOPCOM2.Item("PRODUCT_DESC") = xReader.ReadElementContentAsString()
                        Case "IT_QUANTITY"
                            rowSHOPCOM2.Item("QUANTITY") = xReader.ReadElementContentAsString()
                        Case "IT_UNIT_PRICE"
                            rowSHOPCOM2.Item("UNIT_PRICE") = xReader.ReadElementContentAsString().Replace("$", "")
                        Case "IT_SUB_TOTAL"
                            rowSHOPCOM2.Item("EXTENDED_PRICE") = xReader.ReadElementContentAsString().Replace("$", "")

                    End Select

                Case Else
                    Continue Do
            End Select

        Loop

        xReader.Close()

        'rowSHOPCOM1.Item("DISCOUNT") = "" ' row.Item("F23") & String.Empty
        'rowSHOPCOM1.Item("CATALOG_NAME") = "" ' row.Item("F42") & String.Empty
        'rowSHOPCOM2.Item("VOLUME_NAME") = ""
        'rowSHOPCOM2.Item("SOURCE_CODE") = ""
        'rowSHOPCOM2.Item("COUPON_CODE") = ""


        Return True

    End Function


    Private Sub CreateShopComSalesOrders()

        Dim INVOICE_ID As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim ORDR_LNO As Int16 = 0
        Dim ORDR_DATE As String = String.Empty
        Dim ITEM_CODE As String = String.Empty

        Dim PYMT_METHOD_CODE As String = String.Empty
        Dim PYMT_TYPE_CODE As String = String.Empty

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing
        Dim ORDER_ID As String = String.Empty
        Dim sql As String = String.Empty

        ASCMAIN1.Progress("Importing Shop.com Orders", String.Empty)

        Dim ORDR_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_BATCH_NO")

        For Each rowSHOPCOM1 As DataRow In dst.Tables("SHOPCOM1").Select("", "INVOICE_ID")
            INVOICE_ID = rowSHOPCOM1.Item("INVOICE_ID") & String.Empty
            ORDER_ID = rowSHOPCOM1.Item("ORDER_ID") & String.Empty

            ASCMAIN1.Progress("-", rowSHOPCOM1.Item("ORDER_ID") & String.Empty)

            ' See if we have this Buy.Com Sales Order
            If ORDER_ID.Trim.Length > 0 Then
                sql = "Select * From SOTORDR1 WHERE ORDR_SOURCE_CODE = '" & PARTNER_ORDR_SOURCE_CODE & "' AND PARTNER_ORDR_NO = '" & ORDER_ID & "'"
                If ASCDATA1.GetDataTable(sql).Rows.Count > 0 Then
                    skippedSalesOrder.Add(ORDER_ID)
                    Continue For
                End If
            End If

            ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
            rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
            rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
            rowSOTORDR1.Item("WHSE_CODE") = "001"
            rowSOTORDR1.Item("ORDR_BATCH_NO") = ORDR_BATCH_NO
            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

            ORDR_DATE = rowSHOPCOM1.Item("ORDER_DATE") & String.Empty
            If ORDR_DATE.Length = 8 Then
                ORDR_DATE = ORDR_DATE.Substring(0, 2) & "/" & ORDR_DATE.Substring(2, 2) & "/" & ORDR_DATE.Substring(4, 4)
            End If

            If Not IsDate(ORDR_DATE) Then
                ORDR_DATE = DateTime.Now.ToString("dd-MMM-yyyy")
            Else
                ORDR_DATE = CDate(ORDR_DATE).ToString("dd-MMM-yyyy")
            End If

            rowSOTORDR1.Item("ORDR_DATE") = ORDR_DATE
            rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
            rowSOTORDR1.Item("ORDR_SOURCE_CODE") = PARTNER_ORDR_SOURCE_CODE
            'rowSOTORDR1.Item("ORDR_NO_ORIG") = rowSHOPCOM1.Item("") & String.Empty
            rowSOTORDR1.Item("PARTNER_ORDR_NO") = rowSHOPCOM1.Item("ORDER_ID") & String.Empty
            'rowSOTORDR1.Item("AFFILIATE_NO") = rowSHOPCOM1.Item("") & String.Empty
            rowSOTORDR1.Item("IP_ADDRESS") = rowSHOPCOM1.Item("IP_ADDRESS") & String.Empty

            Dim IP_ADDRESS As String() = (rowSHOPCOM1.Item("IP_ADDRESS") & String.Empty).ToString.Split(".")
            If IP_ADDRESS.Length > 0 Then rowSOTORDR1.Item("IP_A") = IP_ADDRESS(0) & String.Empty
            If IP_ADDRESS.Length > 1 Then rowSOTORDR1.Item("IP_B") = IP_ADDRESS(1) & String.Empty
            If IP_ADDRESS.Length > 2 Then rowSOTORDR1.Item("IP_C") = IP_ADDRESS(2) & String.Empty
            If IP_ADDRESS.Length > 3 Then rowSOTORDR1.Item("IP_D") = IP_ADDRESS(3) & String.Empty
            'rowSOTORDR1.Item("IP_NUMBER") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("IP_COUNTRY") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("PICK_NO") = rowSHOPCOM1.Item("") & String.Empty
            rowSOTORDR1.Item("ORDR_STATUS") = "O"
            'rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("INV_DATE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("SHIP_DATE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("SHIP_REF_NO") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("CARRIER_CODE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("USPS_ZONE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("ORDR_GIFT_MESSAGE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("ORDR_NOTES") = rowSHOPCOM1.Item("") & String.Empty
            rowSOTORDR1.Item("REFERRAL") = "Shop.com"
            rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") = 0
            'rowSOTORDR1.Item("BAD_CUST_MATCH") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("ADDRESS_TYPE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AUTH_SERVICE") = rowSHOPCOM1.Item("") & String.Empty

            PYMT_TYPE_CODE = rowSHOPCOM1.Item("CC_TYPE") & String.Empty
            PYMT_METHOD_CODE = String.Empty

            Select Case PYMT_TYPE_CODE
                Case "1" ' Amex
                    PYMT_METHOD_CODE = "CC"
                    PYMT_TYPE_CODE = "AMEX"

                Case "5" ' DISC
                    PYMT_METHOD_CODE = "CC"
                    PYMT_TYPE_CODE = "DISC"

                Case "6" ' MC
                    PYMT_METHOD_CODE = "CC"
                    PYMT_TYPE_CODE = "MC"

                Case "9" ' Visa
                    PYMT_METHOD_CODE = "CC"
                    PYMT_TYPE_CODE = "VISA"

                Case Else
                    PYMT_METHOD_CODE = "NONE"
            End Select

            'rowSOTORDR1.Item("PYMT_METHOD") = PYMT_METHOD
            'rowSOTORDR1.Item("PYMT_TYPE") = PYMT_TYPE

            rowSOTORDR1.Item("PYMT_CARD_NO") = (rowSHOPCOM1.Item("CC_NUMBER") & String.Empty).ToString.Replace(" ", "").Replace("-", "")
            rowSOTORDR1.Item("PYMT_EXP_DATE") = rowSHOPCOM1.Item("CC_EXP") & String.Empty
            rowSOTORDR1.Item("PYMT_CARD_CVV") = rowSHOPCOM1.Item("CC_CCV") & String.Empty
            rowSOTORDR1.Item("PYMT_CARD_FULL_NAME") = rowSHOPCOM1.Item("NAME_ON_CC") & String.Empty
            'rowSOTORDR1.Item("PYMT_REF_CD") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AUTH_CD") = rowSHOPCOM1.Item("") & String.Empty

            ' Do not fill in, let the credit card processing fill it in
            'rowSOTORDR1.Item("PYMT_AMT") = Val(rowSHOPCOM1.Item("TOTAL") & String.Empty)

            'rowSOTORDR1.Item("PYMT_AVS_CD") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_STREET") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_ZIP") = rowSHOPCOM1.Item("") & String.Empty

            'rowSOTORDR1.Item("PYMT_RECD") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("PYMT_RECD_DATE") = rowSHOPCOM1.Item("") & String.Empty

            rowSOTORDR1.Item("ORDR_SALES_AMT") = Val(rowSHOPCOM1.Item("SUB_TOTAL") & String.Empty)
            rowSOTORDR1.Item("ORDR_COGS_AMT") = 0
            rowSOTORDR1.Item("ORDR_DISC_AMT") = Math.Abs(Val(rowSHOPCOM1.Item("DISCOUNT") & String.Empty)) * -1
            'rowSOTORDR1.Item("ORDR_DISC_PCT") = rowSHOPCOM1.Item("") & String.Empty
            rowSOTORDR1.Item("ORDR_STAX_AMT") = Val(rowSHOPCOM1.Item("TAX") & String.Empty)
            rowSOTORDR1.Item("ORDR_STAX_RATE") = Val(rowSHOPCOM1.Item("TAX_MULT") & String.Empty)
            rowSOTORDR1.Item("ORDR_FRT_AMT") = rowSHOPCOM1.Item("FREIGHT") & String.Empty
            rowSOTORDR1.Item("ORDR_TOT_AMT") = rowSHOPCOM1.Item("TOTAL") & String.Empty
            'rowSOTORDR1.Item("ORDR_TOT_WT") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("INIT_OPER") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("INIT_DATE") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("LAST_OPER") = rowSHOPCOM1.Item("") & String.Empty
            'rowSOTORDR1.Item("LAST_DATE") = rowSHOPCOM1.Item("") & String.Empty
            rowSOTORDR1.Item("SHIP_VIA_ORIG") = rowSHOPCOM1.Item("DELIVERY_METHOD") & String.Empty
            rowSOTORDR1.Item("ORDR_INSTR") = rowSHOPCOM1.Item("COMMENTS") & String.Empty

            rowSOTORDR1.Item("PYMT_METHOD_CODE") = PYMT_METHOD_CODE
            rowSOTORDR1.Item("PYMT_TYPE_CODE") = PYMT_TYPE_CODE

            rowSOTORDR1.Item("PARTNER_INV_NO") = rowSHOPCOM1.Item("INVOICE_ID") & String.Empty
            rowSOTORDR1.Item("PARTNER_SHOPPER_ID") = rowSHOPCOM1.Item("SHOPPER_ID") & String.Empty
            rowSOTORDR1.Item("CAN_SELL_NAME") = IIf(rowSHOPCOM1.Item("CAN_SELL_NAME") & String.Empty = "N", "0", "1")
            rowSOTORDR1.Item("CAN_SEND_OFFERS") = IIf(rowSHOPCOM1.Item("CAN_SEND_OFFERS") & String.Empty = "N", "0", "1")

            For Each custAddrType As String In New String() {"BT", "ST"}

                rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR5.Item("CUST_ADDR_TYPE") = custAddrType
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                rowSOTORDR5.Item("CUST_FIRST_NAME") = rowSHOPCOM1.Item(custAddrType & "_FIRST_NAME") & String.Empty
                rowSOTORDR5.Item("CUST_LAST_NAME") = rowSHOPCOM1.Item(custAddrType & "_LAST_NAME") & String.Empty
                'rowSOTORDR5.Item("CUST_FULL_NAME") = rowSHOPCOM1.Item("") & String.Empty
                rowSOTORDR5.Item("CUST_ADDR1") = rowSHOPCOM1.Item(custAddrType & "_STREET1") & String.Empty
                rowSOTORDR5.Item("CUST_ADDR2") = rowSHOPCOM1.Item(custAddrType & "_STREET2") & String.Empty
                'rowSOTORDR5.Item("CUST_ADDR3") = rowSHOPCOM1.Item("") & String.Empty
                rowSOTORDR5.Item("CUST_CITY") = rowSHOPCOM1.Item(custAddrType & "_CITY") & String.Empty
                rowSOTORDR5.Item("CUST_STATE") = rowSHOPCOM1.Item(custAddrType & "_STATE") & String.Empty
                rowSOTORDR5.Item("CUST_ZIP_CODE") = rowSHOPCOM1.Item(custAddrType & "_ZIP") & String.Empty
                rowSOTORDR5.Item("CUST_COUNTRY") = rowSHOPCOM1.Item(custAddrType & "_COUNTRY") & String.Empty
                'rowSOTORDR5.Item("CUST_CONTACT") = rowSHOPCOM1.Item("") & String.Empty
                rowSOTORDR5.Item("CUST_PHONE") = rowSHOPCOM1.Item(custAddrType & "_TELEPHONE") & String.Empty
                'rowSOTORDR5.Item("CUST_EXT") = rowSHOPCOM1.Item("") & String.Empty
                'rowSOTORDR5.Item("CUST_FAX") = rowSHOPCOM1.Item("") & String.Empty
                rowSOTORDR5.Item("CUST_EMAIL") = TruncateField(rowSHOPCOM1.Item(custAddrType & "_EMAIL") & String.Empty, "SOTORDDR5", "CUST_EMAIL")
                'rowSOTORDR5.Item("CUST_ZIP_MATCH") = rowSHOPCOM1.Item("") & String.Empty
                rowSOTORDR5.Item("CUST_COMPANY_NAME") = rowSHOPCOM1.Item(custAddrType & "_COMPANY_NAME") & String.Empty

                If (rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.ToUpper = "United States".ToUpper Then
                    rowSOTORDR5.Item("CUST_COUNTRY") = "US"
                End If

            Next

            ORDR_LNO = 0
            For Each rowSHOPCOM2 As DataRow In dst.Tables("SHOPCOM2").Select("INVOICE_ID = '" & INVOICE_ID & "'")
                ORDR_LNO += 1

                rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR2.Item("ORDR_LNO") = ORDR_LNO
                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                rowSOTORDR2.Item("ITEM_DESC") = rowSHOPCOM2.Item("PRODUCT_DESC") & String.Empty
                ITEM_CODE = (rowSHOPCOM2.Item("PRODUCT_SKU") & String.Empty).ToString.ToUpper.Trim
                rowSOTORDR2.Item("ITEM_CODE") = ITEM_CODE
                UpdateItemInfo(ITEM_CODE, rowSOTORDR2)

                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Val(rowSHOPCOM2.Item("UNIT_PRICE") & String.Empty)
                rowSOTORDR2.Item("ORDR_QTY") = rowSHOPCOM2.Item("QUANTITY") & String.Empty
                rowSOTORDR2.Item("ORDR_EXT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE") * rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0 'rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
                rowSOTORDR2.Item("ORDR_QTY_CANC") = 0
                rowSOTORDR2.Item("ORDR_QTY_ORIG") = 0
                rowSOTORDR2.Item("UNIT_WEIGHT") = 0
                rowSOTORDR2.Item("PARTNER_LN_ID") = rowSHOPCOM2.Item("PURCHASE_ID") & String.Empty
            Next
        Next
    End Sub

    Private Sub sFtp1_OnDirList(ByVal sender As System.Object, ByVal e As nsoftware.IPWorksSSH.SftpDirListEventArgs) Handles Sftp1.OnDirList
        Dim filename As String = e.FileName

        Select Case PARTNER_CODE

            Case "SHOP"
                If filename.EndsWith(".shop.txt") OrElse filename.EndsWith(".shop") Then
                    ftpFileList.Add(filename)
                End If
        End Select

    End Sub

    Private Sub Ftp1_OnDirList(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.FtpDirListEventArgs) Handles Ftp1.OnDirList
        Dim filename As String = e.FileName

        Select Case PARTNER_CODE
            Case "BUY"
                If filename.EndsWith(".txt") Then
                    ftpFileList.Add(filename)
                End If
        End Select

    End Sub

#End Region

#Region "ShopSite Procedures"

    Private Sub DownLoadShopSiteOrders()

        Dim maxorders As String = "25"

        Dim salesFile As String = PARTNER_ORDERS_DIR & "SO_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
        Dim salesOrders As String = String.Empty

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim dataInStream As Boolean = True

        ASCMAIN1.Progress("Get Webundies.com Sales Orders", String.Empty)

        Dim script As String = PARTNER_SITE_ORDERS_POST_URL
        script &= "startorder=" & Val(PARTNER_LAST_SALES_ORDER) + 1
        'script &= "&maxorder=" & maxorders
        script &= "&pay=yes"
        'script &= "&secure=1"

        'Host: 69.94.109.131:443
        ' Port 443 is the secure port for ShopSite
        Dim sendText As String = String.Empty
        sendText &= "GET " & script & " HTTP/1.0" & Chr(10)
        sendText &= "Host: " & PARTNER_SITE_IP & ":443" & Chr(13) & Chr(10)
        Dim pwd As [Byte]() = Encoding.ASCII.GetBytes(PARTNER_SITE_USER & ":" & PARTNER_SITE_PWD)
        sendText &= "Authorization: Basic " & Convert.ToBase64String(pwd) & Chr(13) & Chr(10) & Chr(10)

        Dim shopSiteResponse As String = String.Empty

        Using tcpClient As New System.Net.Sockets.TcpClient()

            Try
                tcpClient.Connect(PARTNER_SITE_IP, 80)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Shop Site Orders", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try


            Using networkStream As Net.Sockets.NetworkStream = tcpClient.GetStream()

                ' Post the request
                Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(sendText)
                networkStream.Write(sendBytes, 0, sendBytes.Length)

                ' Read the NetworkStream into a byte buffer.
                Dim bytes(tcpClient.ReceiveBufferSize) As Byte

                Dim myReadBuffer(1024) As Byte
                Dim myCompleteMessage As StringBuilder = New StringBuilder()
                Dim numberOfBytesRead As Integer = 0

                If networkStream.CanRead Then

                    ' Incoming message may be larger than the buffer size.
                    ' need to pause to allow buffer to fill up
                    Do
                        numberOfBytesRead = networkStream.Read(myReadBuffer, 0, myReadBuffer.Length)
                        'System.Threading.Thread.Sleep(500)
                        myCompleteMessage.AppendFormat("{0}", Encoding.ASCII.GetString(myReadBuffer, 0, numberOfBytesRead))
                        System.Threading.Thread.Sleep(500)
                    Loop While networkStream.DataAvailable
                End If

                shopSiteResponse = myCompleteMessage.ToString

            End Using

        End Using

        If Not shopSiteResponse.Contains("</ShopSiteOrders>") Then Exit Sub

        salesOrders = shopSiteResponse.Substring(shopSiteResponse.Split("<?xml")(0).Length).Trim
        ' trim the shit characters at the end of the file
        Dim lastChar = InStr(salesOrders, "</ShopSiteOrders>") + "</ShopSiteOrders>".Length

        If lastChar > salesOrders.Length Then
            lastChar = salesOrders.Length
        End If

        salesOrders = salesOrders.Substring(0, lastChar)
        salesOrders = salesOrders.Trim

        salesOrders = salesOrders.Replace(">" & Chr(10) & "<", ">" & Environment.NewLine & "<")

        Using objReader As New StreamWriter(salesFile)
            objReader.Write(salesOrders)
            objReader.Close()
        End Using

        ' Read each xmldocument to create a sales order
        ' Place these in a sorted list so they are processed in numeric order
        dst.Tables("FILES_PROCESSED").Rows.Clear()
        Dim rowSalesOrders As DataRow = Nothing
        For Each salesOrderFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly)
            rowSalesOrders = dst.Tables("FILES_PROCESSED").NewRow
            rowSalesOrders.Item("FileName") = My.Computer.FileSystem.GetName(salesOrderFile)
            rowSalesOrders.Item("FilePath") = salesOrderFile
            dst.Tables("FILES_PROCESSED").Rows.Add(rowSalesOrders)
        Next

    End Sub

    ''' <summary>
    ''' Get sales orders from ShopSite
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub GetShopSiteSalesOrders()

        Me.DownLoadShopSiteOrders()

        If dst.Tables("FILES_PROCESSED").Rows.Count = 0 Then
            Exit Sub
        End If

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim ORDR_NO As String = String.Empty
        Dim sql As String = String.Empty
        Dim duplicateOrder As Boolean = False
        Dim SHIP_COUNTRY As String = String.Empty
        Dim gcLineNo As Int16 = 0

        Dim ORDR_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_BATCH_NO")
        ASCMAIN1.Progress("Importing Webundies.com Orders", String.Empty)

        ' Process the Sales Orders in Order Number Order
        Dim CreatingOrder As Boolean = False
        For Each rowSalesOrders As DataRow In dst.Tables("FILES_PROCESSED").Select("", "FileName")

            Dim FilePath As String = rowSalesOrders.Item("FilePath") & String.Empty

            Dim doc As XmlDocument = New XmlDocument()
            doc.Load(FilePath)

            Dim nodeShopSiteOrder As XmlNode = doc.SelectNodes("ShopSiteOrders")(0)

            If nodeShopSiteOrder.Name <> "ShopSiteOrders" Then
                Continue For
            End If

            'Loop through the records in that node, Should be 1 Response node and a collection of Order nodes
            For Each nodeMain As XmlNode In nodeShopSiteOrder.ChildNodes

                'Get the data we need from the node

                Select Case nodeMain.Name

                    Case "Response"
                        ' One means Successful
                        If Not Response(nodeMain) Then
                            Exit For
                        End If

                    Case "Order"

                        ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                        rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
                        rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
                        rowSOTORDR1.Item("ORDR_SOURCE_CODE") = PARTNER_ORDR_SOURCE_CODE
                        rowSOTORDR1.Item("ORDR_BATCH_NO") = ORDR_BATCH_NO
                        rowSOTORDR1.Item("ORDR_STATUS") = "O"
                        rowSOTORDR1.Item("ORDR_STAX_AMT") = 0
                        rowSOTORDR1.Item("ORDR_COGS_AMT") = 0
                        rowSOTORDR1.Item("ORDR_STAX_RATE") = 0
                        rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") = 0
                        rowSOTORDR1.Item("ORDR_SALES_AMT") = 0
                        rowSOTORDR1.Item("ORDR_DISC_AMT") = 0
                        rowSOTORDR1.Item("ORDR_STAX_AMT") = 0
                        rowSOTORDR1.Item("ORDR_FRT_AMT") = 0
                        rowSOTORDR1.Item("ORDR_TOT_AMT") = 0
                        rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") = 0
                        rowSOTORDR1.Item("WHSE_CODE") = "001"
                        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)
                        ASCMAIN1.Progress("-", ORDR_NO)

                        gcLineNo = 0

                        For Each nodeOrder As XmlNode In nodeMain.ChildNodes

                            Select Case nodeOrder.Name
                                Case "OrderNumber"
                                    rowSOTORDR1.Item("PARTNER_ORDR_NO") = nodeOrder.InnerText & String.Empty
                                    ASCMAIN1.Progress("-", nodeOrder.InnerText & String.Empty)

                                    If Val(rowSOTORDR1.Item("PARTNER_ORDR_NO")) > Val(PARTNER_LAST_SALES_ORDER) Then
                                        PARTNER_LAST_SALES_ORDER = rowSOTORDR1.Item("PARTNER_ORDR_NO")
                                    End If

                                    ' See if this order Number has been processed.
                                    If (nodeOrder.InnerText & String.Empty).Length > 0 Then
                                        sql = "Select * From SOTORDR1 WHERE ORDR_SOURCE_CODE = '" & PARTNER_ORDR_SOURCE_CODE & "' AND PARTNER_ORDR_NO = '" & nodeOrder.InnerText & String.Empty & "'"
                                        If ASCDATA1.GetDataTable(sql).Rows.Count > 0 _
                                            OrElse dst.Tables("SOTORDR1").Select("ORDR_SOURCE_CODE = '" & PARTNER_ORDR_SOURCE_CODE & "' AND PARTNER_ORDR_NO = '" & nodeOrder.InnerText & String.Empty & "'").Length > 1 Then
                                            For Each row As DataRow In dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "'")
                                                row.Delete()
                                            Next

                                            For Each row As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
                                                row.Delete()
                                            Next

                                            For Each row As DataRow In dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")
                                                row.Delete()
                                            Next

                                            ' Grab the next Order since this is a duplicate
                                            rowSOTORDR1 = Nothing
                                            skippedSalesOrder.Add(nodeOrder.InnerText)
                                            Exit For
                                        End If
                                    End If

                                Case "ShopSiteTransactionID"
                                    rowSOTORDR1.Item("PARTNER_INV_NO") = nodeOrder.InnerText & String.Empty

                                Case "OrderDate"
                                    Dim odate As String = nodeOrder.InnerText & String.Empty
                                    If Not IsDate(odate) Then
                                        odate = DateTime.Now.ToString
                                    End If
                                    odate = CDate(odate).ToString("dd-MMM-yyyy")
                                    rowSOTORDR1.Item("ORDR_DATE") = odate
                                    ' CREATE EVENT WITH THIS DATE

                                Case "Billing"
                                    CreateShipping(nodeOrder, ORDR_NO, "BT")

                                Case "Shipping"
                                    CreateShipping(nodeOrder, ORDR_NO, "ST")

                                Case "Payment"
                                    OrderPayment(nodeOrder, rowSOTORDR1)

                                Case "Totals"
                                    OrderTotals(nodeOrder, rowSOTORDR1)

                                Case "Coupon"
                                    Dim Name As String = String.Empty
                                    Dim Status As String = String.Empty
                                    Dim Total As String = String.Empty
                                    Dim ApplyCoupon As String = String.Empty

                                    If nodeOrder.HasChildNodes Then
                                        For Each nodeCoupon As XmlNode In nodeOrder.ChildNodes
                                            Select Case nodeCoupon.Name
                                                Case "Name"
                                                    Name = nodeCoupon.InnerText & String.Empty
                                                Case "Status"
                                                    Status = nodeCoupon.InnerText & String.Empty
                                                Case "Total"
                                                    Total = nodeCoupon.InnerText & String.Empty
                                                Case "ApplyCoupon"
                                                    ApplyCoupon = nodeCoupon.InnerText & String.Empty
                                            End Select
                                        Next
                                    End If

                                    rowSOTORDR1.Item("ORDR_DISC_AMT") = Math.Abs(Val(Total)) * -1

                                    If Status.Length > 0 OrElse Name.Length > 0 Then
                                        rowSOTORDR1.Item("ORDR_DISC_NAME") = TruncateField(Status & "/" & Name, "SOTORDR1", "ORDR_DISC_NAME")
                                    End If

                                Case "Other"
                                    OtherInfo(nodeOrder, rowSOTORDR1)

                                Case "GiftCertificate"

                                    Dim ID As String = String.Empty
                                    Dim AmountUsed As Double = 0
                                    Dim AmountRemaining As Double = 0

                                    If nodeOrder.HasChildNodes Then
                                        For Each nodeGC As XmlNode In nodeOrder.ChildNodes
                                            Select Case nodeGC.Name
                                                Case "ID"
                                                    ID = nodeGC.InnerText & String.Empty
                                                Case "AmountUsed"
                                                    AmountUsed = Val(nodeGC.InnerText.Replace(",", "") & String.Empty)
                                                Case "AmountRemaining"
                                                    AmountRemaining = Val(nodeGC.InnerText.Replace(",", "") & String.Empty)
                                            End Select
                                        Next
                                    End If

                                    If ID.Length > 0 Then
                                        gcLineNo += 1
                                        Dim rowSOTORDRG As DataRow
                                        rowSOTORDRG = dst.Tables("SOTORDRG").NewRow
                                        rowSOTORDRG.Item("ORDR_NO") = ORDR_NO
                                        rowSOTORDRG.Item("ORDR_LNO") = gcLineNo
                                        rowSOTORDRG.Item("GIFTCERT_TRANS_TYPE") = "A"
                                        rowSOTORDRG.Item("GIFTCERT_ID") = ID
                                        rowSOTORDRG.Item("GIFTCERT_APPL") = AmountUsed
                                        rowSOTORDRG.Item("GIFTCERT_BAL") = AmountRemaining
                                        dst.Tables("SOTORDRG").Rows.Add(rowSOTORDRG)

                                        rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") = Val(rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") & String.Empty) - Math.Abs(AmountUsed)
                                    End If

                            End Select
                        Next
                End Select
            Next
        Next

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    ''' <summary>
    ''' Determines if the XML request was successful
    ''' </summary>
    ''' <param name="nodeA"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Response(ByVal nodeA As XmlNode) As Boolean

        For Each nodeC As XmlNode In nodeA.ChildNodes
            Select Case nodeC.Name
                Case "ResponseCode"
                    ' One means Successful
                    Return nodeC.InnerText = "1"
            End Select
        Next

        Return False
    End Function

    ''' <summary>
    ''' Creates The SOTORDR5 entries for the Sales order
    ''' </summary>
    ''' <param name="nodeShip"></param>
    ''' <param name="ORDR_NO"></param>
    ''' <param name="CUST_ADDR_TYPE"></param>
    ''' <remarks></remarks>
    Private Sub CreateShipping(ByVal nodeShip As XmlNode, ByVal ORDR_NO As String, ByVal CUST_ADDR_TYPE As String)

        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
        rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
        rowSOTORDR5.Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

        For Each nodeAddress As XmlNode In nodeShip.ChildNodes

            Select Case nodeAddress.Name
                Case "FullName"
                    rowSOTORDR5.Item("CUST_FULL_NAME") = TruncateField(nodeAddress.InnerText & String.Empty, "SOTORDR5", "CUST_FULL_NAME")

                Case "Company"
                    rowSOTORDR5.Item("CUST_COMPANY_NAME") = TruncateField(nodeAddress.InnerText & String.Empty, "SOTORDR5", "CUST_COMPANY_NAME")

                Case "NameParts"
                    For Each nodeNameParts As XmlNode In nodeAddress.ChildNodes
                        Select Case nodeNameParts.Name
                            Case "FirstName"
                                rowSOTORDR5.Item("CUST_FIRST_NAME") = TruncateField(nodeNameParts.InnerText & String.Empty, "SOTORDR5", "CUST_FIRST_NAME")
                            Case "LastName"
                                rowSOTORDR5.Item("CUST_LAST_NAME") = TruncateField(nodeNameParts.InnerText & String.Empty, "SOTORDR5", "CUST_LAST_NAME")
                        End Select
                    Next

                Case "Email"
                    rowSOTORDR5.Item("CUST_EMAIL") = TruncateField(nodeAddress.InnerText & String.Empty, "SOTORDR5", "CUST_EMAIL")

                Case "Phone"
                    rowSOTORDR5.Item("CUST_PHONE") = TruncateField(nodeAddress.InnerText & String.Empty, "SOTORDR5", "CUST_PHONE")

                Case "Address"
                    For Each nodeAddressx As XmlNode In nodeAddress.ChildNodes
                        Select Case nodeAddressx.Name
                            Case "Street1"
                                rowSOTORDR5.Item("CUST_ADDR1") = TruncateField(nodeAddressx.InnerText & String.Empty, "SOTORDR5", "CUST_ADDR1")
                            Case "Street2"
                                rowSOTORDR5.Item("CUST_ADDR2") = TruncateField(nodeAddressx.InnerText & String.Empty, "SOTORDR5", "CUST_ADDR2")
                            Case "City"
                                rowSOTORDR5.Item("CUST_CITY") = TruncateField(nodeAddressx.InnerText & String.Empty, "SOTORDR5", "CUST_CITY")
                            Case "State"
                                rowSOTORDR5.Item("CUST_STATE") = TruncateField(nodeAddressx.InnerText & String.Empty, "SOTORDR5", "CUST_STATE")
                            Case "Code"
                                rowSOTORDR5.Item("CUST_ZIP_CODE") = TruncateField(nodeAddressx.InnerText & String.Empty, "SOTORDR5", "CUST_ZIP_CODE")
                            Case "Country"
                                rowSOTORDR5.Item("CUST_COUNTRY") = TruncateField(nodeAddressx.InnerText & String.Empty, "SOTORDR5", "CUST_COUNTRY")
                        End Select
                    Next

                    ' Try to convert Country Name to a Country Code
                    Dim countryCode As String = (rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.Trim.ToUpper
                    If countryCode.Length > 0 Then
                        If dst.Tables("TATCNTRY").Select("COUNTRY_NAME = '" & countryCode & "'").Length > 0 Then
                            rowSOTORDR5.Item("CUST_COUNTRY") = dst.Tables("TATCNTRY").Select("COUNTRY_NAME = '" & countryCode & "'")(0).Item("COUNTRY_CODE") & String.Empty
                        End If
                    End If

                    If Not ",US,CA,".Contains("," & rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty & ",") Then
                        'rowSOTORDR5.Item("CUST_STATE") = String.Empty
                    End If

                Case "Products"
                    CreateOrderDetails(nodeAddress, ORDR_NO)

            End Select
        Next

        If (rowSOTORDR5.Item("CUST_FULL_NAME") & String.Empty).ToString.Trim.Length = 0 Then
            rowSOTORDR5.Item("CUST_FULL_NAME") = (rowSOTORDR5.Item("CUST_FIRST_NAME") & " " & rowSOTORDR5.Item("CUST_LAST_NAME")).ToString.Trim
        End If

    End Sub

    ''' <summary>
    ''' Create Sotordr2 entries
    ''' </summary>
    ''' <param name="nodeProducts"></param>
    ''' <param name="ORDR_NO"></param>
    ''' <remarks></remarks>
    Private Sub CreateOrderDetails(ByVal nodeProducts As XmlNode, ByVal ORDR_NO As String)

        Dim ORDR_LNO As Int16 = 0
        Dim ITEM_CODE As String = String.Empty
        Dim rowSOTORDR2 As DataRow = Nothing
        Dim rowSOTORDRG As DataRow = Nothing
        Dim isGiftCert As Boolean = False

        For Each nodeProduct As XmlNode In nodeProducts.ChildNodes
            Select Case nodeProduct.Name
                Case "Product"
                    ORDR_LNO += 1
                    isGiftCert = False
                    rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                    rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDR2.Item("ORDR_LNO") = ORDR_LNO
                    rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                    rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = 0
                    rowSOTORDR2.Item("ORDR_PRICE_SOURCE") = "W"
                    dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                    For Each nodeItem As XmlNode In nodeProduct.ChildNodes
                        Select Case nodeItem.Name
                            Case "SKU"
                                ITEM_CODE = (nodeItem.InnerText & String.Empty).Trim.ToUpper
                                If ITEM_CODE.Length = 0 Then Exit Select
                                rowSOTORDR2.Item("ITEM_CODE") = ITEM_CODE
                                UpdateItemInfo(ITEM_CODE, rowSOTORDR2)

                            Case "ProdType"
                                ITEM_CODE = (nodeItem.InnerText & String.Empty).Trim.ToUpper
                                If ITEM_CODE.Length = 0 Then Exit Select

                                If ITEM_CODE.ToUpper <> "E-mailGiftCertificate".ToUpper Then
                                    Exit Select
                                End If
                                isGiftCert = True
                                ITEM_CODE = "GIFTCERT"
                                rowSOTORDR2.Item("ITEM_CODE") = ITEM_CODE
                                UpdateItemInfo(ITEM_CODE, rowSOTORDR2)

                            Case "Name"
                                rowSOTORDR2.Item("ITEM_DESC") = nodeItem.InnerText & String.Empty

                            Case "Quantity"
                                rowSOTORDR2.Item("ORDR_QTY") = Val(nodeItem.InnerText.Replace(",", "") & String.Empty)
                                rowSOTORDR2.Item("ORDR_QTY_ORIG") = rowSOTORDR2.Item("ORDR_QTY")

                            Case "ItemPrice"
                                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Val(nodeItem.InnerText.Replace(",", "") & String.Empty)

                            Case "Total"
                                rowSOTORDR2.Item("ORDR_EXT_PRICE") = Val(nodeItem.InnerText.Replace(",", "") & String.Empty)

                            Case "Weight"
                                rowSOTORDR2.Item("UNIT_WEIGHT") = Val(nodeItem.InnerText.Replace(",", "") & String.Empty)

                            Case "GiftCertificatePurchase"
                                If Not isGiftCert Then Exit Select

                                Dim GIFTCERT_EMAIL As String = String.Empty
                                Dim GIFTCERT_TO As String = String.Empty
                                Dim GIFTCERT_FROM As String = String.Empty
                                Dim GIFTCERT_MSG As String = String.Empty

                                For Each nodeItemGCP As XmlNode In nodeItem.ChildNodes
                                    Select Case nodeItemGCP.Name
                                        Case "EmailTo"
                                            For Each nodeItemeEmail As XmlNode In nodeItemGCP.ChildNodes
                                                If (nodeItemeEmail.InnerText & String.Empty) <> String.Empty Then
                                                    GIFTCERT_EMAIL &= " " & nodeItemeEmail.InnerText & String.Empty
                                                End If
                                            Next
                                            GIFTCERT_EMAIL = GIFTCERT_EMAIL.Trim
                                        Case "From"
                                            GIFTCERT_FROM = nodeItemGCP.InnerText & String.Empty
                                        Case "To"
                                            GIFTCERT_TO = nodeItemGCP.InnerText & String.Empty
                                        Case "Message"
                                            GIFTCERT_MSG = nodeItemGCP.InnerText & String.Empty
                                    End Select
                                Next

                                TruncateField(GIFTCERT_EMAIL, "SOTORDRG", "GIFTCERT_EMAIL")
                                TruncateField(GIFTCERT_FROM, "SOTORDRG", "GIFTCERT_FROM")
                                TruncateField(GIFTCERT_TO, "SOTORDRG", "GIFTCERT_TO")
                                TruncateField(GIFTCERT_MSG, "SOTORDRG", "GIFTCERT_MSG")

                                rowSOTORDRG = dst.Tables("SOTORDRG").NewRow
                                rowSOTORDRG.Item("ORDR_NO") = ORDR_NO
                                rowSOTORDRG.Item("ORDR_LNO") = ORDR_LNO
                                rowSOTORDRG.Item("GIFTCERT_TRANS_TYPE") = "P"
                                rowSOTORDRG.Item("GIFTCERT_EMAIL") = GIFTCERT_EMAIL
                                rowSOTORDRG.Item("GIFTCERT_FROM") = GIFTCERT_FROM
                                rowSOTORDRG.Item("GIFTCERT_TO") = GIFTCERT_TO
                                rowSOTORDRG.Item("GIFTCERT_MSG") = GIFTCERT_MSG
                                rowSOTORDRG.Item("GIFTCERT_BAL") = Val(rowSOTORDR2.Item("ORDR_EXT_PRICE") & String.Empty)
                                dst.Tables("SOTORDRG").Rows.Add(rowSOTORDRG)

                        End Select
                    Next

                    ' See that we got an item code
                    If rowSOTORDR2.Item("ITEM_CODE") & String.Empty = String.Empty Then
                        ITEM_CODE = "Unknown"
                        rowSOTORDR2.Item("ITEM_CODE") = ITEM_CODE
                        UpdateItemInfo(ITEM_CODE, rowSOTORDR2)
                    End If

            End Select
        Next
    End Sub

    ''' <summary>
    ''' Sale Order Header Payment Info
    ''' </summary>
    ''' <param name="nodePayment"></param>
    ''' <param name="rowSOTORDR1"></param>
    ''' <remarks></remarks>
    Private Sub OrderPayment(ByVal nodePayment As XmlNode, ByRef rowSOTORDR1 As DataRow)

        For Each nodePaymentType As XmlNode In nodePayment.ChildNodes
            Select Case nodePaymentType.Name

                Case "CreditCard"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "CC"
                    For Each nodeCC As XmlNode In nodePaymentType.ChildNodes
                        Select Case nodeCC.Name
                            Case "Issuer"
                                Select Case nodeCC.InnerText
                                    Case "MasterCard"
                                        rowSOTORDR1.Item("PYMT_TYPE_CODE") = "MC"
                                    Case "Visa"
                                        rowSOTORDR1.Item("PYMT_TYPE_CODE") = "VISA"
                                    Case "Discover"
                                        rowSOTORDR1.Item("PYMT_TYPE_CODE") = "DISC"
                                    Case "AMEX", "American Express", "Amex"
                                        rowSOTORDR1.Item("PYMT_TYPE_CODE") = "AMEX"
                                    Case Else
                                        rowSOTORDR1.Item("PYMT_TYPE_CODE") = "NONE"
                                End Select

                            Case "Number"
                                rowSOTORDR1.Item("PYMT_CARD_NO") = (nodeCC.InnerText & String.Empty).ToString.Replace(" ", "").Replace("-", "")

                            Case "VerificationValue"
                                rowSOTORDR1.Item("PYMT_CARD_CVV") = nodeCC.InnerText

                            Case "FullName"
                                rowSOTORDR1.Item("PYMT_CARD_FULL_NAME") = nodeCC.InnerText

                            Case "Company"
                                rowSOTORDR1.Item("PYMT_CARD_COMPANY") = nodeCC.InnerText

                            Case "ExpirationDate"
                                nodeCC.InnerText = nodeCC.InnerText.Trim
                                If (nodeCC.InnerText & String.Empty).Length > 4 Then
                                    rowSOTORDR1.Item("PYMT_EXP_DATE") = nodeCC.InnerText.Substring(0, 2)
                                    rowSOTORDR1.Item("PYMT_EXP_DATE") &= nodeCC.InnerText.Substring(nodeCC.InnerText.Length - 2, 2)
                                Else
                                    rowSOTORDR1.Item("PYMT_EXP_DATE") = nodeCC.InnerText
                                End If

                            Case "PaymentGateway"

                            Case "PaymentStatus"

                            Case "OrderProcessingInfo"
                                ShopSiteCreditCardTransaction(rowSOTORDR1, nodeCC)

                        End Select
                    Next

                Case "Amazon"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "AP"
                    rowSOTORDR1.Item("PYMT_TYPE_CODE") = "AP"
                    If IsDate(rowSOTORDR1.Item("ORDR_DATE") & String.Empty) Then
                        rowSOTORDR1.Item("PYMT_RECD_DATE") = rowSOTORDR1.Item("ORDR_DATE")
                    Else
                        rowSOTORDR1.Item("PYMT_RECD_DATE") = CDate(DateTime.Now.ToShortDateString)
                    End If
                    rowSOTORDR1.Item("PYMT_RECD") = "1"

                Case "COD"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "Check"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"
                    For Each nodeChk As XmlNode In nodePaymentType.ChildNodes
                        Select Case nodeChk.Name
                            Case "RoutingNumber"
                                rowSOTORDR1.Item("PYMT_REF_CD") = nodeChk.InnerText
                            Case "Account Number"
                                rowSOTORDR1.Item("PYMT_CARD_NO") = (nodeChk.InnerText & String.Empty).ToString.Replace(" ", "").Replace("-", "")
                        End Select
                    Next

                Case "eCheck"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "PurchaseOrder"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "Generic1"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "MP"

                Case "Generic2"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "PayPal"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "PayPalExpress"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "PP"
                    rowSOTORDR1.Item("PYMT_TYPE_CODE") = "PP"

                    For Each ppNode As XmlNode In nodePaymentType.ChildNodes
                        If ppNode.Name = "OrderProcessingInfo" Then
                            ShopSiteCreditCardTransaction(rowSOTORDR1, nodePaymentType)
                        End If
                    Next

                Case "WorldPay"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "Google"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "NetBanx"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "Solo"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "Switch"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "Delta"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "UKE"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

                Case "JCB"
                    rowSOTORDR1.Item("PYMT_METHOD_CODE") = "NONE"

            End Select
        Next
    End Sub

    Private Sub ShopSiteCreditCardTransaction(ByRef rowSOTORDR1 As DataRow, ByRef nodeOrderProcessingInfo As XmlNode)

        Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
        Dim payPalExpress As String = String.Empty
        Dim SEQ_NO As Int16 = 0
        Dim ppData() As String
        Dim ppDate As String = String.Empty

        ppData = nodeOrderProcessingInfo.InnerText.Replace("|n|", Chr(13)).Split(Chr(13))
        If ppData.Length <= 1 Then
            Exit Sub
        End If

        SEQ_NO += 1
        Dim rowSOTPAYPL As DataRow = dst.Tables("SOTPAYPL").NewRow

        rowSOTPAYPL.Item("ORDR_NO") = ORDR_NO
        rowSOTPAYPL.Item("SEQ_NO") = SEQ_NO
        dst.Tables("SOTPAYPL").Rows.Add(rowSOTPAYPL)

        For Each ppDataValue As String In ppData

            If ppDataValue.StartsWith("TRANSACTIONID:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("TRANS_ID") = ppDataValue.Substring("TRANSACTIONID:".Length).Trim
                rowSOTORDR1.Item("PYMT_REF_CD") = ppDataValue.Substring("TRANSACTIONID:".Length).Trim
            ElseIf ppDataValue.StartsWith("*PAYPALEXPRESS-", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("PP_TYPE") = ppDataValue.Substring("*PAYPALEXPRESS-".Length).Replace("*", "").Trim
                rowSOTORDR1.Item("PYMT_AUTH_SERVICE") = "PAYPALEXPRESS"
            ElseIf ppDataValue.StartsWith("*PAYPALDIRECT-", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("PP_TYPE") = ppDataValue.Substring("*PAYPALDIRECT-".Length).Replace("*", "").Trim
                rowSOTORDR1.Item("PYMT_AUTH_SERVICE") = "PAYPALDIRECT"
            ElseIf ppDataValue.StartsWith("TRANSACTIONTYPE:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("TRANS_TYPE") = ppDataValue.Substring("TRANSACTIONTYPE:".Length).Trim
            ElseIf ppDataValue.StartsWith("PAYMENTTYPE:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("PYMT_TYPE") = ppDataValue.Substring("PAYMENTTYPE:".Length).Trim
            ElseIf ppDataValue.StartsWith("ORDERTIME:", StringComparison.OrdinalIgnoreCase) OrElse ppDataValue.StartsWith("TIMESTAMP:", StringComparison.OrdinalIgnoreCase) Then

                ppDataValue = ppDataValue.ToUpper.Replace("TIMESTAMP:", "ORDERTIME:")
                ppDate = ppDataValue.Substring("ORDERTIME:".Length).Trim
                ppDate = ppDate.Replace("T", " ").Replace("Z", " ").Trim

                If IsDate(ppDate) Then
                    rowSOTPAYPL.Item("ORDER_TIME") = Convert.ToDateTime(ppDate)
                    rowSOTORDR1.Item("PYMT_RECD_DATE") = Convert.ToDateTime(ppDate).ToString("dd-MMM-yyyy")
                    rowSOTORDR1.Item("PYMT_RECD") = "1"
                End If
            ElseIf ppDataValue.StartsWith("AMT:", StringComparison.OrdinalIgnoreCase) OrElse ppDataValue.StartsWith("AMOUNT:", StringComparison.OrdinalIgnoreCase) Then
                ppDataValue = ppDataValue.ToUpper.Replace("AMOUNT:", "AMT:")
                rowSOTPAYPL.Item("AMOUNT") = Val(ppDataValue.Substring("AMT:".Length).Trim)
                rowSOTORDR1.Item("PYMT_AMT") = Val(ppDataValue.Substring("AMT:".Length).Trim)
            ElseIf ppDataValue.StartsWith("FEEAMT:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("FEE_AMT") = Val(ppDataValue.Substring("FEEAMT:".Length).Trim)
            ElseIf ppDataValue.StartsWith("PAYMENTSTATUS:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("STATUS") = ppDataValue.Substring("PAYMENTSTATUS:".Length).Trim
                If (rowSOTPAYPL.Item("STATUS") & String.Empty).ToString.Trim.ToUpper = "PENDING" Then
                    SalesOrderError(ORDR_NO, "PP")
                End If
            ElseIf ppDataValue.StartsWith("PENDINGREASON:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("PEND_REAS") = ppDataValue.Substring("PENDINGREASON:".Length).Trim
            ElseIf ppDataValue.StartsWith("AVSCODE:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("AVS_CODE") = ppDataValue.Substring("AVSCode:".Length).Trim
                rowSOTORDR1.Item("PYMT_AVS_CD") = ppDataValue.Substring("AVSCode:".Length).Trim
            ElseIf ppDataValue.StartsWith("CVV2CODE:", StringComparison.OrdinalIgnoreCase) Then
                rowSOTPAYPL.Item("CVV2_CODE") = ppDataValue.Substring("CVV2CODE:".Length).Trim
                rowSOTORDR1.Item("PYMT_AVS_CVV2") = ppDataValue.Substring("CVV2CODE:".Length).Trim
            End If
        Next

    End Sub

    ''' <summary>
    ''' Sales Order Header Totals
    ''' </summary>
    ''' <param name="nodeTotals"></param>
    ''' <param name="rowSOTORDR1"></param>
    ''' <remarks></remarks>
    Private Sub OrderTotals(ByVal nodeTotals As XmlNode, ByRef rowSOTORDR1 As DataRow)

        Dim SHIP_SERVICE_LEVEL As String = String.Empty
        Dim rowSOTSVIA2 As DataRow = Nothing
        Dim SHIP_COUNTRY As String = String.Empty

        For Each nodeTotal As XmlNode In nodeTotals.ChildNodes

            Select Case nodeTotal.Name
                Case "ProductTotal"
                    rowSOTORDR1.Item("ORDR_SALES_AMT") = Val(nodeTotal.InnerText.Replace(",", "") & String.Empty)

                Case "Discount"
                    rowSOTORDR1.Item("ORDR_DISC_AMT") = Math.Abs(Val(nodeTotal.InnerText.Replace(",", "") & String.Empty)) * -1
                    If rowSOTORDR1.Item("ORDR_DISC_AMT") = 0 OrElse rowSOTORDR1.Item("ORDR_SALES_AMT") = 0 Then
                        rowSOTORDR1.Item("ORDR_DISC_PCT") = 0
                    Else
                        rowSOTORDR1.Item("ORDR_DISC_PCT") = rowSOTORDR1.Item("ORDR_DISC_AMT") / rowSOTORDR1.Item("ORDR_SALES_AMT")
                    End If

                Case "Tax"
                    For Each nodeTax As XmlNode In nodeTotal.ChildNodes
                        Select Case nodeTax.Name
                            Case "TaxRate"
                                rowSOTORDR1.Item("ORDR_STAX_RATE") = Val(nodeTax.InnerText.Replace(",", "") & String.Empty)
                            Case "TaxAmount"
                                rowSOTORDR1.Item("ORDR_STAX_AMT") = Val(nodeTax.InnerText.Replace(",", "") & String.Empty)
                        End Select
                    Next

                Case "GrandTotal"
                    rowSOTORDR1.Item("ORDR_TOT_AMT") = Val(nodeTotal.InnerText.Replace(",", "") & String.Empty)

                Case "ShippingTotal"
                    For Each nodeShippingTotal As XmlNode In nodeTotal.ChildNodes
                        Select Case nodeShippingTotal.Name
                            Case "Total"
                                rowSOTORDR1.Item("ORDR_FRT_AMT") = Val(nodeShippingTotal.InnerText.Replace(",", "") & String.Empty)
                            Case "Description"
                                rowSOTORDR1.Item("SHIP_VIA_ORIG") = nodeShippingTotal.InnerText & String.Empty

                        End Select
                    Next
            End Select
        Next
    End Sub

    Private Sub Coupon()

    End Sub

    Private Sub OtherInfo(ByVal nodeOther As XmlNode, ByRef rowSOTORDR1 As DataRow)

        For Each nodeOtherData As XmlNode In nodeOther.ChildNodes

            Select Case nodeOtherData.Name

                Case "Comments"
                    rowSOTORDR1.Item("ORDR_NOTES") = nodeOtherData.InnerText & String.Empty

                Case "OrderInstructions"

                Case "OrderUTC"
                    'rowSOTORDR1.Item("IP_NUMBER") = nodeOtherData.InnerText & String.Empty

                Case "IpHostname"
                    Dim IP_ADDRESS As String = nodeOtherData.InnerText & String.Empty

                    While IP_ADDRESS.Length > 0 AndAlso Not Char.IsDigit(IP_ADDRESS.Substring(0, 1))
                        IP_ADDRESS = IP_ADDRESS.Substring(1)
                    End While

                    Dim loc As Int16 = 0
                    If IP_ADDRESS.Length > 0 Then
                        While loc < IP_ADDRESS.Length AndAlso (Char.IsDigit(IP_ADDRESS.Substring(loc, 1)) OrElse IP_ADDRESS.Substring(loc, 1) = ".")
                            loc += 1
                        End While
                        IP_ADDRESS = IP_ADDRESS.Substring(0, loc)
                    End If

                    If IP_ADDRESS.Length > 0 Then
                        rowSOTORDR1.Item("IP_ADDRESS") = IP_ADDRESS
                    End If

                Case "TotalOrderWeight"
                    rowSOTORDR1.Item("ORDR_TOT_WT") = Val(nodeOtherData.InnerText.Replace(",", "") & String.Empty)

                Case "CustomCheckoutField"
                    Dim field As String = String.Empty
                    For Each nodeCCF As XmlNode In nodeOtherData.ChildNodes
                        Select Case nodeCCF.Name

                            Case "FieldName"

                                Select Case nodeCCF.InnerText.ToUpper
                                    Case "Boxergram".ToUpper
                                        field = "ORDR_GIFT_MESSAGE"
                                    Case "Referral".ToUpper
                                        field = "REFERRAL"
                                    Case "BoxergramTo".ToUpper
                                        field = "ORDR_GIFT_MESSAGE_TO"
                                    Case "BoxergramFrom".ToUpper
                                        field = "ORDR_GIFT_MESSAGE_FROM"
                                End Select

                            Case "FieldValue"
                                If field.Length > 0 Then
                                    If Not (field = "REFERRAL" AndAlso (nodeCCF.InnerText & String.Empty).Trim.ToUpper.Contains("NONE SELECTED")) Then
                                        If (nodeCCF.InnerText & String.Empty).ToString.Length > 0 AndAlso (nodeCCF.InnerText & String.Empty).ToString.Substring(0, 1) = "," Then
                                            nodeCCF.InnerText = (nodeCCF.InnerText & String.Empty).ToString.Substring(1)
                                        End If
                                        rowSOTORDR1.Item(field) = nodeCCF.InnerText & String.Empty
                                    End If
                                End If
                                field = String.Empty

                        End Select
                    Next
            End Select
        Next
    End Sub

    'Private Sub ProcessCreditCards()

    '    If dst.Tables("SOTORDR1").Select(shopCCQuery).Length = 0 Then
    '        Exit Sub
    '    End If

    '    MessageBox.Show("System will now process credit cards.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '    ASCMAIN1.Progress("Processing Credit cards", String.Empty)

    '    SO_PARM_PP_SIGNATURE = (ROWs("SOTPARM1").Item("SO_PARM_PP_SIGNATURE") & String.Empty).ToString.Trim
    '    SO_PARM_PP_USERNAME = (ROWs("SOTPARM1").Item("SO_PARM_PP_USERNAME") & String.Empty).ToString.Trim
    '    SO_PARM_PP_PASSWORD = (ROWs("SOTPARM1").Item("SO_PARM_PP_PASSWORD") & String.Empty).ToString.Trim
    '    SO_PARM_PP_URL = (ROWs("SOTPARM1").Item("SO_PARM_PP_URL") & String.Empty).ToString.Trim

    '    Dim rowSOTORDR5 As DataRow = Nothing
    '    Dim ORDR_NO As String = String.Empty
    '    Dim PYMT_EXP_DATE As String = String.Empty

    '    Dim clsTACPAYPL As New TAC.TACPAYPL

    '    CheckTestMode()
    '    With clsTACPAYPL

    '        .Signature = SO_PARM_PP_SIGNATURE
    '        .User = SO_PARM_PP_USERNAME
    '        .Password = SO_PARM_PP_PASSWORD
    '        .URL = SO_PARM_PP_URL

    '        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(shopCCQuery)
    '            ORDR_NO = rowSOTORDR1.Item("ORDR_NO")

    '            ASCMAIN1.Progress("-", ORDR_NO)
    '            Record_Event(ORDR_NO, "Start Credit Card processing")

    '            rowSOTORDR5 = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'BT'")(0)

    '            .ChargeAmount = Val(rowSOTORDR1.Item("ORDR_TOT_AMT") & String.Empty)
    '            .TransactionDescription = "Shop.com Order: " & rowSOTORDR1.Item("PARTNER_ORDR_NO")
    '            .InvoiceNumber = ORDR_NO

    '            If .ChargeAmount <= 0 Then Continue For
    '            PYMT_EXP_DATE = rowSOTORDR1.Item("PYMT_EXP_DATE") & String.Empty

    '            ' format to MMYY
    '            If PYMT_EXP_DATE.Length >= 4 Then
    '                PYMT_EXP_DATE = PYMT_EXP_DATE.Substring(0, 2) & PYMT_EXP_DATE.Substring(PYMT_EXP_DATE.Length - 2, 2)
    '            Else
    '                SalesOrderError(ORDR_NO, "CC")
    '                Continue For
    '            End If

    '            If testCreditCardNo.Length > 0 Then
    '                .CreditCard.CreditCardNumber = testCreditCardNo
    '                .CreditCard.CreditCardExpireMonth = testCreditCardExp.Substring(0, 2)
    '                .CreditCard.CreditCardExpireYear = testCreditCardExp.Substring(2, 2)
    '                .CreditCard.CreditCardCCV = testCreditCardCCV2
    '            Else
    '                .CreditCard.CreditCardNumber = rowSOTORDR1.Item("PYMT_CARD_NO") & String.Empty
    '                .CreditCard.CreditCardExpireMonth = PYMT_EXP_DATE.Substring(0, 2)
    '                .CreditCard.CreditCardExpireYear = PYMT_EXP_DATE.Substring(2, 2)
    '                .CreditCard.CreditCardCCV = rowSOTORDR1.Item("PYMT_CARD_CVV") & String.Empty
    '            End If

    '            .Payee.LastName = rowSOTORDR5.Item("CUST_LAST_NAME") & String.Empty
    '            .Payee.FirstName = rowSOTORDR5.Item("CUST_FIRST_NAME") & String.Empty
    '            .Payee.City = rowSOTORDR5.Item("CUST_CITY") & String.Empty
    '            .Payee.State = rowSOTORDR5.Item("CUST_STATE") & String.Empty
    '            .Payee.Street1 = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
    '            .Payee.Street2 = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
    '            .Payee.CountryCode = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
    '            .Payee.Zip = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
    '            .Payee.EmailAddress = rowSOTORDR5.Item("CUST_EMAIL") & String.Empty

    '            .DirectSale()


    '            If .Errors.Count > 0 Then
    '                For Each Err As TAC.TAICCARD.CreditCardErrors In .Errors
    '                    Record_Event(ORDR_NO, Err.ToString)
    '                Next
    '            End If

    '            If .ListErrors.Count > 0 Then
    '                For Each Err As String In .ListErrors
    '                    Record_Event(ORDR_NO, Err)
    '                Next
    '            End If

    '            Dim rowSOTPAYPL As DataRow = dst.Tables("SOTPAYPL").NewRow
    '            rowSOTPAYPL.Item("ORDR_NO") = ORDR_NO
    '            rowSOTPAYPL.Item("SEQ_NO") = "1"
    '            rowSOTPAYPL.Item("PP_TYPE") = "SALE: " & .Response.Acknowledgement.ToString
    '            rowSOTPAYPL.Item("TRANS_ID") = .Response.TransactionID & String.Empty
    '            rowSOTPAYPL.Item("TRANS_TYPE") = "Shop.com"
    '            rowSOTPAYPL.Item("PYMT_TYPE") = String.Empty
    '            rowSOTPAYPL.Item("ORDER_TIME") = DateTime.Now
    '            rowSOTPAYPL.Item("AMOUNT") = Val(.Response.Amount & String.Empty)
    '            rowSOTPAYPL.Item("FEE_AMT") = 0
    '            rowSOTPAYPL.Item("STATUS") = .Response.Acknowledgement
    '            rowSOTPAYPL.Item("PEND_REAS") = String.Empty
    '            rowSOTPAYPL.Item("AVS_CODE") = .Response.AVS & String.Empty
    '            rowSOTPAYPL.Item("CVV2_CODE") = .Response.CCV & String.Empty
    '            dst.Tables("SOTPAYPL").Rows.Add(rowSOTPAYPL)

    '            Record_Event(ORDR_NO, "Credit Card processing status: " & .Response.Acknowledgement.ToString)

    '            If .Response.Acknowledgement = TAC.TAICCARD.RequestResponseCodes.Success Then

    '                rowSOTORDR1.Item("PYMT_RECD") = "1"
    '                rowSOTORDR1.Item("PYMT_RECD_DATE") = DateTime.Now.ToString("dd-MMM-yyyy")
    '                rowSOTORDR1.Item("PYMT_REF_CD") = .Response.TransactionID
    '                rowSOTORDR1.Item("PYMT_AMT") = Val(.Response.Amount & String.Empty)

    '                ' remove any credit card needs to be processed holds
    '                For Each row As DataRow In dst.Tables("SOTORDRV").Select("ORDR_NO = '" & ORDR_NO & "' AND ERROR_CODE = 'CCP'")
    '                    row.Item("ERROR_STATUS") = "1"
    '                Next

    '                ' If there are no other holds then set the Order to Open
    '                If dst.Tables("SOTORDRV").Select("ORDR_NO = '" & ORDR_NO & "' AND ISNULL(ERROR_STATUS, '0') = '0'").Length = 0 _
    '                    AndAlso rowSOTORDR1.Item("ORDR_STATUS") = "H" Then
    '                    rowSOTORDR1.Item("ORDR_STATUS") = "O"
    '                End If
    '            Else
    '                SalesOrderError(ORDR_NO, "CC")
    '                rowSOTORDR1.Item("ORDR_STATUS") = "H"
    '                rowSOTORDR1.Item("HAS_ERRORS") = "1"
    '            End If
    '        Next
    '    End With

    '    ASCMAIN1.Progress(String.Empty, String.Empty)
    '    MessageBox.Show("Credit card processing completed.", "Import", MessageBoxButtons.OK, MessageBoxIcon.Information)

    'End Sub

    Private Sub CheckTestMode()

        testCreditCardNo = String.Empty
        testCreditCardExp = String.Empty
        testCreditCardCCV2 = String.Empty

        If (ASCMAIN1.DBS_COMPANY <> "WUN" AndAlso ASCMAIN1.DBS_SERVER <> "WUN") Then
            SO_PARM_PP_SIGNATURE = "ArZVhsgMt2Xnh-0gEb7CLza-WEc8Anplrd01chp0nKBXOTWQFcUxhPZX"
            SO_PARM_PP_USERNAME = "xtest1_1199718490_biz_api1.hotmail.com"
            SO_PARM_PP_PASSWORD = "1199718508"
            SO_PARM_PP_URL = "https://api.sandbox.paypal.com/nvp"
            testCreditCardNo = "4832419131427146"
            testCreditCardExp = "0117"
            testCreditCardCCV2 = "123"
        End If
    End Sub

    Private Sub SalesOrderError(ByVal ORDR_NO As String, ByVal ERROR_CODE As String)

        If dst.Tables("SOTORDRV").Select("ORDR_NO = '" & ORDR_NO & "' AND ERROR_CODE = '" & ERROR_CODE & "'").Length > 0 Then
            Exit Sub
        End If

        ERROR_CODE = ERROR_CODE.Trim

        If ERROR_CODE.Length = 0 Then
            Exit Sub
        End If

        Dim rowSOTORDRV As DataRow = dst.Tables("SOTORDRV").NewRow
        rowSOTORDRV.Item("ORDR_NO") = ORDR_NO
        rowSOTORDRV.Item("ERROR_CODE") = ERROR_CODE
        rowSOTORDRV.Item("ERROR_STATUS") = "0"
        dst.Tables("SOTORDRV").Rows.Add(rowSOTORDRV)

    End Sub

#End Region

#Region "BuyDotCom Procedures"

    Private Sub CreateBuyComDataTable()

        With dst

            dst.Tables.Add("BUYCOM")
            With dst.Tables("BUYCOM")

                .Columns.Add("SELLER_ID", GetType(System.String))
                .Columns.Add("PARTNER_ORDR_NO", GetType(System.String))
                .Columns.Add("PARTNER_ORDR_LNO", GetType(System.Int64))
                .Columns.Add("LISTING_ID", GetType(System.Int64))
                .Columns.Add("ORDR_DATE", GetType(System.DateTime))
                .Columns.Add("PRODUCT_SKU", GetType(System.String))
                .Columns.Add("REFERENCE_ID", GetType(System.String))
                .Columns.Add("ORD_QTY", GetType(System.Int64))
                .Columns.Add("ORD_QTY_SHIP", GetType(System.Int64))
                .Columns.Add("ORD_QTY_CANCEL", GetType(System.Int64))
                .Columns.Add("TITLE", GetType(System.String))
                .Columns.Add("PRICE", GetType(System.Decimal))
                .Columns.Add("PRICE_EXT", GetType(System.Decimal))
                .Columns.Add("ORDR_FRT_AMT", GetType(System.Decimal))
                .Columns.Add("PRODUCT_OWED", GetType(System.Decimal))
                .Columns.Add("SHIPPING_OWED", GetType(System.Decimal))
                .Columns.Add("COMMISSION", GetType(System.Decimal))
                .Columns.Add("SHIPPING_FEE", GetType(System.Decimal))
                .Columns.Add("PER_ITEM_FEE", GetType(System.Decimal))
                .Columns.Add("ORDR_STAX_AMT", GetType(System.Decimal))

                .Columns.Add("BT_COMPANY_NAME", GetType(System.String))
                .Columns.Add("BT_TELEPHONE", GetType(System.String))
                .Columns.Add("BT_FIRSTNAME", GetType(System.String))
                .Columns.Add("BT_LASTNAME", GetType(System.String))
                .Columns.Add("BT_EMAIL", GetType(System.String))

                .Columns.Add("ST_FULLNAME", GetType(System.String))
                .Columns.Add("ST_COMPANY_NAME", GetType(System.String))
                .Columns.Add("ST_STREET1", GetType(System.String))
                .Columns.Add("ST_STREET2", GetType(System.String))
                .Columns.Add("ST_CITY", GetType(System.String))
                .Columns.Add("ST_STATE", GetType(System.String))
                .Columns.Add("ST_ZIP", GetType(System.String))

                .Columns.Add("DELIVERY_METHOD", GetType(System.Int16))
            End With

        End With
    End Sub

    Private Sub GetBuyComSalesOrders()

        If PARTNER_SITE_IP.Length = 0 Then Exit Sub
        If PARTNER_SITE_USER.Length = 0 Then Exit Sub
        If PARTNER_SITE_PWD.Length = 0 Then Exit Sub
        If PARTNER_SITE_OUTPUT_DIR.Length = 0 Then Exit Sub
        If PARTNER_ORDERS_DIR.Length = 0 Then Exit Sub

        Dim localFilename As String = String.Empty

        If Not My.Computer.FileSystem.DirectoryExists(PARTNER_ORDERS_DIR) Then
            Exit Sub
        End If

        If Not PARTNER_ORDERS_DIR.EndsWith("\") Then PARTNER_ORDERS_DIR &= "\"

        ' FTP file dowm from Shop.Com
        Try
            ASCMAIN1.Progress("Creating FTP Connection to Buy.com", "")

            Ftp1 = New nsoftware.IPWorks.Ftp
            Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

            ASCMAIN1.Progress("-", "RemoteHost")
            Ftp1.RemoteHost = PARTNER_SITE_IP

            ASCMAIN1.Progress("-", "User")
            Ftp1.User = PARTNER_SITE_USER

            ASCMAIN1.Progress("-", "Password")
            Ftp1.Password = PARTNER_SITE_PWD

            ASCMAIN1.Progress("-", "RemoteFile")
            Ftp1.RemoteFile = String.Empty

            ASCMAIN1.Progress("-", "Timeout")
            Ftp1.Timeout = 300

            ASCMAIN1.Progress("-", "Logon")
            Try
                Ftp1.Passive = False
                Ftp1.Logon()
            Catch ex As Exception
                Ftp1.Logoff()
                Ftp1.Passive = False
                Ftp1.Logon()
            End Try

            Ftp1.RemotePath = PARTNER_SITE_OUTPUT_DIR
            ftpFileList = New List(Of String)
            Ftp1.ListDirectory()

            For Each fileFtp As String In ftpFileList

                ASCMAIN1.Progress("Downloading: " & fileFtp, String.Empty)
                Ftp1.RemoteFile = fileFtp
                localFilename = fileFtp

                Ftp1.LocalFile = PARTNER_ORDERS_DIR & localFilename
                Ftp1.Download()
                Ftp1.DeleteFile(fileFtp)

            Next

        Catch ex As Exception
            MessageBox.Show("Error downloading order files: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Ftp1.Logoff()
            Ftp1.Dispose()
        End Try

        dst.Tables("FILES_PROCESSED").Rows.Clear()
        Dim rowSalesOrders As DataRow = Nothing

        For Each buycomSalesFile As String In My.Computer.FileSystem.GetFiles(PARTNER_ORDERS_DIR, FileIO.SearchOption.SearchTopLevelOnly, "*.txt")
            ASCMAIN1.Progress("Importing: " & buycomSalesFile, String.Empty)
            If ReadFromBuyComFile(buycomSalesFile) Then
                rowSalesOrders = dst.Tables("FILES_PROCESSED").NewRow
                rowSalesOrders.Item("FileName") = My.Computer.FileSystem.GetName(buycomSalesFile)
                rowSalesOrders.Item("FilePath") = buycomSalesFile
                dst.Tables("FILES_PROCESSED").Rows.Add(rowSalesOrders)
            End If
        Next

        CreateBuyComSalesOrders()

    End Sub

    Public Function ReadFromBuyComFile(ByVal buyComFileName As String) As Boolean

        Dim buyComColumnCount As Int16 = dst.Tables("BUYCOM").Columns.Count
        Dim currentrow(buyComColumnCount) As String

        Dim boolFileExists As Boolean = False
        Dim rowBUYCOM As DataRow = Nothing

        Try

            boolFileExists = My.Computer.FileSystem.FileExists(buyComFileName)

            If boolFileExists = False Then
                Return False
            End If

            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(buyComFileName)

                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(vbTab)

                While Not MyReader.EndOfData
                    currentrow = MyReader.ReadFields()

                    ' Ship any header records
                    If currentrow(0).ToString.Trim.ToUpper = "SellerShopperNumber".ToUpper Then
                        Continue While
                    End If

                    rowBUYCOM = dst.Tables("BUYCOM").NewRow
                    With rowBUYCOM
                        .Item("SELLER_ID") = currentrow(0)
                        .Item("PARTNER_ORDR_NO") = currentrow(1)
                        .Item("PARTNER_ORDR_LNO") = currentrow(2)
                        .Item("LISTING_ID") = currentrow(3)

                        .Item("ORDR_DATE") = currentrow(4)
                        .Item("PRODUCT_SKU") = currentrow(5)
                        .Item("REFERENCE_ID") = currentrow(6)

                        .Item("ORD_QTY") = currentrow(7)
                        .Item("ORD_QTY_SHIP") = currentrow(8)
                        .Item("ORD_QTY_CANCEL") = currentrow(9)
                        .Item("TITLE") = currentrow(10)
                        .Item("PRICE") = currentrow(11)
                        .Item("PRICE_EXT") = currentrow(12)
                        .Item("ORDR_FRT_AMT") = currentrow(13)
                        .Item("PRODUCT_OWED") = currentrow(14)
                        .Item("SHIPPING_OWED") = currentrow(15)
                        .Item("COMMISSION") = currentrow(16)
                        .Item("SHIPPING_FEE") = currentrow(17)
                        .Item("PER_ITEM_FEE") = currentrow(18)
                        .Item("ORDR_STAX_AMT") = currentrow(19)

                        .Item("BT_COMPANY_NAME") = currentrow(20)
                        .Item("BT_TELEPHONE") = currentrow(21)
                        .Item("BT_FIRSTNAME") = currentrow(22)
                        .Item("BT_LASTNAME") = currentrow(23)
                        .Item("BT_EMAIL") = currentrow(24)

                        .Item("ST_FULLNAME") = currentrow(25)
                        .Item("ST_COMPANY_NAME") = currentrow(26)
                        .Item("ST_STREET1") = currentrow(27)
                        .Item("ST_STREET2") = currentrow(28)
                        .Item("ST_CITY") = currentrow(29)
                        .Item("ST_STATE") = currentrow(30)
                        .Item("ST_ZIP") = currentrow(31)

                        .Item("DELIVERY_METHOD") = currentrow(32)
                    End With

                    dst.Tables("BUYCOM").Rows.Add(rowBUYCOM)
                End While
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Buy.Com Import Error", MessageBoxButtons.OK)
            Return False
        End Try

        Return True

    End Function

    Private Sub CreateBuyComSalesOrders()

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing

        Dim ORDR_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_BATCH_NO")
        Dim ORDR_NO As String = String.Empty
        Dim ORDR_LNO As Int16 = 0
        Dim PARTNER_ORDR_NO As String = String.Empty
        Dim telePhone As String = String.Empty
        Dim custName As String = String.Empty
        Dim maxLength As Int16 = 0

        Dim ORDR_SALES_AMT As Decimal = 0
        Dim ORDR_DISC_AMT As Decimal = 0
        Dim ORDR_DISC_PCT As Decimal = 0
        Dim ORDR_STAX_AMT As Decimal = 0
        Dim ORDR_STAX_RATE As Decimal = 0
        Dim ORDR_FRT_AMT As Decimal = 0
        Dim ORDR_TOT_AMT As Decimal = 0
        Dim SHIP_COUNTRY As String = String.Empty

        Dim ITEM_CODE As String = String.Empty

        Dim rowBUYCOM As DataRow = Nothing
        Dim sql As String = String.Empty

        ASCMAIN1.Progress("Importing Buy.Com Orders", String.Empty)

        ' Get the distinct Order Numbers
        For Each rowBUYCOMx As DataRow In ASCDATA1.SelectDistinct("BUYCOM", New String() {"PARTNER_ORDR_NO"}).Rows

            PARTNER_ORDR_NO = rowBUYCOMx.Item("PARTNER_ORDR_NO") & String.Empty
            ASCMAIN1.Progress("-", PARTNER_ORDR_NO)

            ' See if we have this Buy.Com Sales Order
            If PARTNER_ORDR_NO.Trim.Length > 0 Then
                sql = "Select * From SOTORDR1 WHERE ORDR_SOURCE_CODE = '" & PARTNER_ORDR_SOURCE_CODE & "' AND PARTNER_ORDR_NO = '" & PARTNER_ORDR_NO & "'"
                If ASCDATA1.GetDataTable(sql).Rows.Count > 0 Then
                    skippedSalesOrder.Add(PARTNER_ORDR_NO)
                    Continue For
                End If
            End If

            ' Get a single record to create the header
            rowBUYCOM = dst.Tables("BUYCOM").Select("PARTNER_ORDR_NO = '" & PARTNER_ORDR_NO & "'")(0)

            ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")

            ORDR_SALES_AMT = 0
            ORDR_DISC_AMT = 0
            ORDR_DISC_PCT = 0
            ORDR_STAX_AMT = 0
            ORDR_STAX_RATE = 0
            ORDR_FRT_AMT = 0
            ORDR_TOT_AMT = 0

            rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
            rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
            rowSOTORDR1.Item("WHSE_CODE") = "001"
            rowSOTORDR1.Item("ORDR_BATCH_NO") = ORDR_BATCH_NO
            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

            If IsDate(rowBUYCOM.Item("ORDR_DATE") & String.Empty) Then
                rowSOTORDR1.Item("ORDR_DATE") = CDate(rowBUYCOM.Item("ORDR_DATE") & String.Empty).ToString("dd-MMM-yyyy")
            Else
                rowSOTORDR1.Item("ORDR_DATE") = DateTime.Now.ToString("dd-MMM-yyyy")
            End If

            rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
            rowSOTORDR1.Item("ORDR_SOURCE_CODE") = PARTNER_ORDR_SOURCE_CODE
            'rowSOTORDR1.Item("ORDR_NO_ORIG") = ORDER_ID
            rowSOTORDR1.Item("PARTNER_ORDR_NO") = PARTNER_ORDR_NO
            'rowSOTORDR1.Item("AFFILIATE_NO") = String.Empty
            'rowSOTORDR1.Item("IP_ADDRESS") = String.Empty
            'rowSOTORDR1.Item("IP_A") = String.Empty
            'rowSOTORDR1.Item("IP_B") =  String.Empty
            'rowSOTORDR1.Item("IP_C") =  String.Empty
            'rowSOTORDR1.Item("IP_D") =  String.Empty
            'rowSOTORDR1.Item("IP_NUMBER") =  String.Empty
            'rowSOTORDR1.Item("IP_COUNTRY") = TruncateField(rowAMAZON.Item("SHIP_COUNTRY") & String.Empty, "SOTORDR1", "IP_COUNTRY")
            'rowSOTORDR1.Item("PICK_NO") =  String.Empty
            rowSOTORDR1.Item("ORDR_STATUS") = "O"
            'rowSOTORDR1.Item("INV_DATE") =  String.Empty
            'rowSOTORDR1.Item("SHIP_DATE") =  String.Empty
            'rowSOTORDR1.Item("SHIP_REF_NO") =  String.Empty
            'rowSOTORDR1.Item("USPS_ZONE") =  String.Empty
            'rowSOTORDR1.Item("ORDR_GIFTCERT_APPL") = 0
            'rowSOTORDR1.Item("ORDR_GIFT_MESSAGE") = String.Empty
            'rowSOTORDR1.Item("ORDR_NOTES") = String.Empty
            rowSOTORDR1.Item("REFERRAL") = "Buy.Com"
            'rowSOTORDR1.Item("BAD_CUST_MATCH") = String.Empty
            'rowSOTORDR1.Item("ADDRESS_TYPE") = String.Empty
            'rowSOTORDR1.Item("PYMT_METHOD") = String.Empty
            'rowSOTORDR1.Item("PYMT_AUTH_SERVICE") = String.Empty
            'rowSOTORDR1.Item("PYMT_TYPE") = String.Empty
            'rowSOTORDR1.Item("PYMT_CARD_NO") = String.Empty
            'rowSOTORDR1.Item("PYMT_EXP_DATE") = String.Empty
            'rowSOTORDR1.Item("PYMT_CARD_CVV") = String.Empty
            'rowSOTORDR1.Item("PYMT_REF_CD") = String.Empty
            'rowSOTORDR1.Item("PYMT_AUTH_CD") = String.Empty
            'rowSOTORDR1.Item("PYMT_AMT") = String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_CD") = String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_STREET") = String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_ZIP") = String.Empty
            'rowSOTORDR1.Item("PYMT_AVS_CVV2") = String.Empty
            'rowSOTORDR1.Item("PYMT_RECD") = String.Empty
            rowSOTORDR1.Item("PYMT_RECD") = "1"
            rowSOTORDR1.Item("PYMT_RECD_DATE") = rowSOTORDR1.Item("ORDR_DATE")
            'rowSOTORDR1.Item("INIT_OPER") = String.Empty
            'rowSOTORDR1.Item("INIT_DATE") = String.Empty
            'rowSOTORDR1.Item("LAST_OPER") = String.Empty
            'rowSOTORDR1.Item("LAST_DATE") = String.Empty
            rowSOTORDR1.Item("SHIP_VIA_ORIG") = TruncateField(rowBUYCOM.Item("DELIVERY_METHOD") & String.Empty, "SOTORDR1", "SHIP_VIA_ORIG")
            'rowSOTORDR1.Item("ORDR_INSTR") = rowAMAZON.Item("") & String.Empty
            rowSOTORDR1.Item("PYMT_METHOD_CODE") = "BUY"
            rowSOTORDR1.Item("PYMT_TYPE_CODE") = "BUY"

            ' SOTORDR5
            For Each addrType As String In New String() {"BT", "ST"}

                rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
                rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR5.Item("CUST_ADDR_TYPE") = addrType
                rowSOTORDR5.Item("CUST_COUNTRY") = "US"
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

                Select Case addrType

                    Case "BT"
                        rowSOTORDR5.Item("CUST_COMPANY_NAME") = TruncateField(rowBUYCOM.Item("BT_COMPANY_NAME") & String.Empty, "SOTORDR5", "CUST_COMPANY_NAME")
                        rowSOTORDR5.Item("CUST_PHONE") = TruncateField(rowBUYCOM.Item("BT_TELEPHONE") & String.Empty, "SOTORDR5", "CUST_PHONE")
                        rowSOTORDR5.Item("CUST_FIRST_NAME") = TruncateField(rowBUYCOM.Item("BT_FIRSTNAME") & String.Empty, "SOTORDR5", "CUST_FIRST_NAME")
                        rowSOTORDR5.Item("CUST_LAST_NAME") = TruncateField(rowBUYCOM.Item("BT_LASTNAME") & String.Empty, "SOTORDR5", "CUST_LAST_NAME")
                        rowSOTORDR5.Item("CUST_FULL_NAME") = TruncateField(rowBUYCOM.Item("BT_FIRSTNAME") & String.Empty & " " & rowBUYCOM.Item("BT_LASTNAME") & String.Empty, "SOTORDR5", "CUST_FULL_NAME")
                        rowSOTORDR5.Item("CUST_EMAIL") = TruncateField(rowBUYCOM.Item("BT_EMAIL") & String.Empty, "SOTORDR5", "CUST_EMAIL")

                    Case "ST"
                        rowSOTORDR5.Item("CUST_FULL_NAME") = TruncateField(rowBUYCOM.Item("ST_FULLNAME") & String.Empty, "SOTORDR5", "CUST_FULL_NAME")
                        rowSOTORDR5.Item("CUST_COMPANY_NAME") = TruncateField(rowBUYCOM.Item("ST_COMPANY_NAME") & String.Empty, "SOTORDR5", "CUST_COMPANY_NAME")
                        rowSOTORDR5.Item("CUST_ADDR1") = TruncateField(rowBUYCOM.Item("ST_STREET1") & String.Empty, "SOTORDR5", "CUST_ADDR1")
                        rowSOTORDR5.Item("CUST_ADDR2") = TruncateField(rowBUYCOM.Item("ST_STREET2") & String.Empty, "SOTORDR5", "CUST_ADDR2")
                        rowSOTORDR5.Item("CUST_CITY") = TruncateField(rowBUYCOM.Item("ST_CITY") & String.Empty, "SOTORDR5", "CUST_CITY")
                        rowSOTORDR5.Item("CUST_STATE") = TruncateField(ConvertState(rowBUYCOM.Item("ST_STATE") & String.Empty), "SOTORDR5", "CUST_STATE")
                        rowSOTORDR5.Item("CUST_ZIP_CODE") = TruncateField(rowBUYCOM.Item("ST_ZIP") & String.Empty, "SOTORDR5", "CUST_ZIP_CODE")
                End Select
            Next

            ORDR_LNO = 0
            For Each rowDetails As DataRow In dst.Tables("BUYCOM").Select("PARTNER_ORDR_NO = '" & PARTNER_ORDR_NO & "'", "PARTNER_ORDR_LNO")
                ORDR_LNO += 1

                rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR2.Item("ORDR_LNO") = ORDR_LNO
                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                ITEM_CODE = (rowDetails.Item("REFERENCE_ID") & String.Empty).ToString.Trim.ToUpper
                rowSOTORDR2.Item("ITEM_CODE") = ITEM_CODE
                rowSOTORDR2.Item("ITEM_DESC") = TruncateField(rowDetails.Item("TITLE") & String.Empty, "SOTORDR2", "ITEM_DESC")
                UpdateItemInfo(ITEM_CODE, rowSOTORDR2)

                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = Val(rowDetails.Item("PRICE") & String.Empty)
                rowSOTORDR2.Item("ORDR_QTY") = Val(rowDetails.Item("ORD_QTY") & String.Empty)


                rowSOTORDR2.Item("ORDR_EXT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE") * rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
                rowSOTORDR2.Item("ORDR_QTY_CANC") = 0
                rowSOTORDR2.Item("ORDR_QTY_ORIG") = rowSOTORDR2.Item("ORDR_QTY")
                rowSOTORDR2.Item("UNIT_WEIGHT") = 0
                rowSOTORDR2.Item("PARTNER_LN_ID") = TruncateField(rowDetails.Item("PARTNER_ORDR_LNO") & String.Empty, "SOTORDR2", "PARTNER_LN_ID")

                ORDR_SALES_AMT += Val(rowSOTORDR2.Item("ORDR_EXT_PRICE") & String.Empty)

                ORDR_DISC_AMT += 0
                ORDR_DISC_AMT = Math.Abs(ORDR_DISC_AMT) * -1

                ORDR_DISC_PCT += 0
                ORDR_STAX_AMT += Val(rowDetails.Item("ORDR_STAX_AMT") & String.Empty)
                ORDR_STAX_RATE += 0
                ORDR_FRT_AMT += Val(rowDetails.Item("ORDR_FRT_AMT") & String.Empty)
            Next

            ORDR_TOT_AMT = ORDR_SALES_AMT + ORDR_DISC_AMT + ORDR_STAX_AMT + ORDR_FRT_AMT

            rowSOTORDR1.Item("ORDR_SALES_AMT") = ORDR_SALES_AMT
            rowSOTORDR1.Item("ORDR_COGS_AMT") = 0
            rowSOTORDR1.Item("ORDR_DISC_AMT") = ORDR_DISC_AMT
            rowSOTORDR1.Item("ORDR_DISC_PCT") = ORDR_DISC_PCT
            rowSOTORDR1.Item("ORDR_STAX_AMT") = ORDR_STAX_AMT
            rowSOTORDR1.Item("ORDR_STAX_RATE") = ORDR_STAX_RATE
            rowSOTORDR1.Item("ORDR_FRT_AMT") = ORDR_FRT_AMT
            rowSOTORDR1.Item("ORDR_TOT_AMT") = ORDR_TOT_AMT
            rowSOTORDR1.Item("ORDR_TOT_WT") = 0
        Next

    End Sub
#End Region

#Region "Form Procedures"

    ''' <summary>
    ''' Truncates a fields value if the length is longer the the max length of the field
    ''' </summary>
    ''' <param name="fieldValue"></param>
    ''' <param name="TableName"></param>
    ''' <param name="FieldName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function TruncateField(ByVal fieldValue As String, ByVal TableName As String, ByVal FieldName As String) As String

        Dim rValue As String = fieldValue

        If Not dst.Tables.Contains(TableName) Then
            Return rValue
        End If

        If Not dst.Tables(TableName).Columns.Contains(FieldName) Then
            Return rValue
        End If

        Dim maxLength As Int16 = 0
        maxLength = dst.Tables(TableName).Columns(FieldName).MaxLength

        If rValue.Length > maxLength Then
            rValue = rValue.Substring(0, maxLength).Trim
        End If

        Return rValue
    End Function

    ''' <summary>
    ''' Converts State Names to State Abbreviations
    ''' </summary>
    ''' <param name="StateCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertState(ByVal StateCode As String) As String

        Select Case StateCode.Trim.ToUpper

            Case "ALABAMA" : Return "AL"
            Case "ALASKA" : Return "AK"
            Case "ARIZONA" : Return "AZ"
            Case "ARKANSAS" : Return "AR"
            Case "CALIFORNIA" : Return "CA"
            Case "COLORADO" : Return "CO"
            Case "CONNECTICUT" : Return "CT"
            Case "DELAWARE" : Return "DE"
            Case "FLORIDA" : Return "FL"
            Case "GEORGIA" : Return "GA"
            Case "HAWAII" : Return "HI"
            Case "IDAHO" : Return "ID"
            Case "ILLINOIS" : Return "IL"
            Case "INDIANA" : Return "IN"
            Case "IOWA" : Return "IA"
            Case "KANSAS" : Return "KS"
            Case "KENTUCKY" : Return "KY"
            Case "LOUISIANA" : Return "LA"
            Case "MAINE" : Return "ME"
            Case "MARYLAND" : Return "MD"
            Case "MASSACHUSETTS" : Return "MA"
            Case "MICHIGAN" : Return "MI"
            Case "MINNESOTA" : Return "MN"
            Case "MISSISSIPPI" : Return "MS"
            Case "MISSOURI" : Return "MO"
            Case "MONTANA" : Return "MT"
            Case "NEBRASKA" : Return "NE"
            Case "NEVADA" : Return "NV"
            Case "NEW HAMPSHIRE" : Return "NH"
            Case "NEW JERSEY" : Return "NJ"
            Case "NEW MEXICO" : Return "NM"
            Case "NEW YORK" : Return "NY"
            Case "NORTH CAROLINA" : Return "NC"
            Case "NORTH DAKOTA" : Return "ND"
            Case "OHIO" : Return "OH"
            Case "OKLAHOMA" : Return "OK"
            Case "OREGON" : Return "OR"
            Case "PENNSYLVANIA" : Return "PA"
            Case "PUERTO RICO" : Return "PR"
            Case "RHODE ISLAND" : Return "RI"
            Case "SOUTH CAROLINA" : Return "SC"
            Case "SOUTH DAKOTA" : Return "SD"
            Case "TENNESSEE" : Return "TN"
            Case "TEXAS" : Return "TX"
            Case "UTAH" : Return "UT"
            Case "VERMONT" : Return "VT"
            Case "VIRGINIA" : Return "VA"
            Case "WASHINGTON" : Return "WA"
            Case "WEST VIRGINIA" : Return "WV"
            Case "WEST VA" : Return "WV"
            Case "WISCONSIN" : Return "WI"
            Case "WYOMING" : Return "WY"

            Case "ALBERTA" : Return "AB"
            Case "BRITISH COLUMBIA" : Return "BC"
            Case "MANITOBA" : Return "MB"
            Case "NEW BRUNSWICK" : Return "NB"
            Case "NEWFOUNDLAND AND LABRADOR" : Return "NL"
            Case "NORTHWEST TERRITORIES" : Return "NT"
            Case "NOVA SCOTIA" : Return "NS"
            Case "NUNAVUT" : Return "NU"
            Case "ONTARIO" : Return "ON"
            Case "PRINCE EDWARD ISLAND" : Return "PE"
            Case "QUEBEC" : Return "QC"
            Case "SASKATCHEWAN" : Return "SK"
            Case "YUKON" : Return "YT"

            Case Else
                Return StateCode.Trim.ToUpper

        End Select

    End Function

    ''' <summary>
    ''' Convert Country Codes
    ''' </summary>
    ''' <param name="CountryCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ConvertCountry(ByVal CountryCode As String) As String

        Select Case CountryCode.ToUpper

            Case "PR" : Return "US"
            Case "GU" : Return "US"
            Case "VI" : Return "US"

            Case Else
                Return CountryCode

        End Select
    End Function

    Private Sub UpdateItemInfo(ByVal ITEM_CODE As String, ByRef rowSOTORDR2 As DataRow)

        rowICTITEM1 = MyBase.LookUp("ICTITEM1", ITEM_CODE)
        rowSOTORDR2.Item("ORDR_UNIT_COST") = 0

        If rowICTITEM1 IsNot Nothing Then
            rowSOTORDR2.Item("STYLE_CODE") = rowICTITEM1.Item("STYLE_CODE") & String.Empty
            rowSOTORDR2.Item("COLOR_CODE") = rowICTITEM1.Item("COLOR_CODE") & String.Empty
            rowSOTORDR2.Item("SIZE_CODE") = rowICTITEM1.Item("SIZE_CODE") & String.Empty

            rowICTSTYL1 = MyBase.LookUp("ICTSTYL1", rowICTITEM1.Item("STYLE_CODE") & String.Empty)

            If rowICTSTYL1 IsNot Nothing Then
                rowSOTORDR2.Item("ORDR_UNIT_COST") = Val(rowICTSTYL1.Item("STYLE_COST") & String.Empty)
            End If
        End If
    End Sub

    Private Sub Record_Event(ByVal ORDR_NO As String, ByVal EVENT_DESC As String)
        Dim row As DataRow = dst.Tables("SOTORDRE").NewRow
        row.Item("ORDR_NO") = ORDR_NO
        row.Item("INIT_DATE") = DateTime.Now
        row.Item("INIT_OPER") = ASCMAIN1.USER_ID
        row.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("SOTORDRE").Rows.Add(row)
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As System.Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME

            Case "ORDR_SOURCE_CODE"
                sql_where = " ORDR_SOURCE_CODE IN (SELECT PARTNER_ORDR_SOURCE_CODE FROM SOTPART1)"
        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdSOTORDR1, "SSPB", "Show Filter", "Show GroupBox", "Sales Order Entry")
        Call Load_Popup_Menu(grdSOTORDRI, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        Select Case e.Tool.Key

            Case "grdSOTORDR1", "grdSOTORDRI"
                ' Nothing 
            Case Else
                e.Cancel = True
                Exit Sub
        End Select

    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

            Case "Sales Order Entry"

                If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsFilterRow Then
                    Exit Sub
                End If

                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text

                If ORDR_NO.Length > 0 Then
                    Context_Launch("Edit", ORDR_NO, e.Tool.Key, "SOFORDR1") ', "F", "SOE")
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTORDR1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTORDR1.AfterRowActivate

        If viewSOTORDRV Is Nothing Then
            grdSOTORDRV.Text = "No selected order"
            Exit Sub
        End If

        If grdSOTORDR1.ActiveRow.IsFilterRow Then
            grdSOTORDRV.Text = "No selected order"
            viewSOTORDRV.RowFilter = "ORDR_NO = '@@@@'"
            Exit Sub
        End If

        If grdSOTORDR1.ActiveRow Is Nothing OrElse grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value & String.Empty = String.Empty Then
            grdSOTORDRV.Text = "No selected order"
            viewSOTORDRV.RowFilter = "ORDR_NO = '@@@@'"
            Exit Sub
        End If

        Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value & String.Empty
        viewSOTORDRV.RowFilter = "ORDR_NO = '" & ORDR_NO & "'"
        grdSOTORDRV.Text = "Errors for Sales Order: " & ORDR_NO

    End Sub

    Private Sub txtORDR_SOURCE_CODE_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtORDR_SOURCE_CODE.Leave
        txtORDR_SOURCE_CODE.Text = txtORDR_SOURCE_CODE.Text.Trim.ToUpper
    End Sub

    Private Sub btnGetData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetData.Click

        Try
            Dim sql As String = String.Empty

            sql = "SELECT * FROM SOTORDRI WHERE trunc(IMPORT_DATE) BETWEEN '" & CDate(dteStart.Value).ToString("dd-MMM-yyyy") & "' AND '" & CDate(dteEnd.Value).ToString("dd-MMM-yyyy") & "'"
            Fill_Records("SOTORDRI", String.Empty, True, sql)

            grdSOTORDRI.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

            grdSOTORDRI.DisplayLayout.Bands(0).SortedColumns.Clear()
            grdSOTORDRI.DisplayLayout.Bands(0).SortedColumns.Add("IMPORT_DATE", False, True)
            grdSOTORDRI.DisplayLayout.Bands(0).SortedColumns.Add("PARTNER_LNO", False)

            grdSOTORDRI.DisplayLayout.GroupByBox.Hidden = False

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        UltraExplorerBar1.Groups("Log Parameters").Visible = UltraTabControl1.SelectedTab.Key = "Sales Order Service Log"
    End Sub

#End Region
End Class
