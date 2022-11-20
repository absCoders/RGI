Imports System.Net.Http
Imports Microsoft.Office.Core
Imports Newtonsoft.Json

Public Class TACMAIN1
    Inherits ABSolution.TACMAIN1
    Public Shared CATCH_PACK As String = "000"
    Public Shared app_WHSE_CODE_PWS As String = "002"
    'Public DIVISION_CODE_of_USER_ID As String = ""
    Public Shared SREP_CODE As String = ""
    Public Shared DIVISION_CODE As String = ""
    Public Shared Receiver_ID As String = ""
    Public Shared EDI_PROCESS_IND As String = ""

    Public Shared SREP_CODE_MGR_ASST As String = ""
    Public Shared SREP_CODEs As New List(Of String)
    Public Const SSLEnabledProtocols As Int32 = 4032


#Region "NYA Canadian Warehouses"

    Private Shared NYA_CANADIAN_WAREHOUSES As New List(Of String)({"18", "21"})

    Public Shared ReadOnly Property NyaCanadaWhseList() As List(Of String)
        Get
            Return NYA_CANADIAN_WAREHOUSES
        End Get
    End Property

    Public Shared ReadOnly Property NyaCanadaWhseQueryString As String
        Get
            Return "'" & String.Join("', '", NYA_CANADIAN_WAREHOUSES.ToArray) & "'"
        End Get
    End Property

    Public Shared ReadOnly Property NyaCanadaWhseArray As String()
        Get
            Return NYA_CANADIAN_WAREHOUSES.ToArray
        End Get
    End Property

    Public Shared ReadOnly Property NyaCanadaWhseCommaSeparatedString As String
        Get
            Return String.Join(",", NYA_CANADIAN_WAREHOUSES.ToArray)
        End Get
    End Property

#End Region

#Region "Canadian Customers"

    Private Shared CanadianCustomers As New List(Of String)
    Private Shared tblCanadianCustomers As DataTable = Nothing

    Public Shared ReadOnly Property CanadaCustomerList() As List(Of String)
        Get
            fillCanadaCustomer()
            Return CanadianCustomers
        End Get
    End Property

    Public Shared ReadOnly Property CanadaCustomerQueryString As String
        Get
            fillCanadaCustomer()
            Return "'" & String.Join("', '", CanadianCustomers.ToArray) & "'"
        End Get
    End Property

    Public Shared ReadOnly Property CanadaCustomerArray As String()
        Get
            fillCanadaCustomer()
            Return CanadianCustomers.ToArray
        End Get
    End Property

    Public Shared ReadOnly Property CanadaCustomerCommaSeparatedString As String
        Get
            fillCanadaCustomer()
            Return String.Join(",", CanadianCustomers.ToArray)
        End Get
    End Property

    Private Shared Sub fillCanadaCustomer()
        If tblCanadianCustomers IsNot Nothing Then Exit Sub
        If CanadianCustomers.Count > 0 Then Exit Sub

        tblCanadianCustomers = ASCDATA1.GetDataTable("Select CUST_CODE from ARTCUST1 where CURR_CODE = 'CAD' or SEG4_CODE = '001'")
        For Each row As DataRow In tblCanadianCustomers.Select("")
            CanadianCustomers.Add(row.Item("CUST_CODE"))
        Next
    End Sub

#End Region

    Public Overrides Sub Site_Specific_Settings()

        If ASCMAIN1.CLIENT = "RGO" Then
            Exit Sub
        End If

        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            Exit Sub
        End If

        ASCMAIN1.LabelPrinterSerialPort = Nothing
        ASCMAIN1.LaserPrinterName = String.Empty

        Dim n1 As New nsoftware.IPWorks.Ftp
        Debug.Print(n1.RuntimeLicense)
        Dim n2 As New nsoftware.IPWorksEncrypt.Keymgr
        Debug.Print(n2.RuntimeLicense)
        Dim n3 As New nsoftware.IPWorksSSH.Sftp
        Debug.Print(n3.RuntimeLicense)
        Dim n4 As New nsoftware.IPWorksZip.Gzip
        Debug.Print(n4.RuntimeLicense)


        'This is causing me too many problems and should not be running at Regency in ABSolution lite.
        Dim stationID As String = System.Environment.GetEnvironmentVariable("USERNAME") & String.Empty
        Dim sql As String = "SELECT * FROM WHTLINE1 WHERE UPPER(STATION_ID) = :PARM1"
        stationID = stationID.ToUpper

        If stationID.Contains("ZENKER") Then
            stationID = "EZENKER"
        End If


        Dim rowWHTLINE1 As DataRow = Nothing

        If ASCMAIN1.CLIENT <> "ABS" Then
            rowWHTLINE1 = ASCDATA1.GetDataRow(sql, "V", New Object() {stationID})
        End If


        If (ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN") Then
            If rowWHTLINE1 Is Nothing Then
                rowWHTLINE1 = ASCDATA1.GetDataRow(sql, "V", New Object() {"DEFAULT"})
            End If
        End If

        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID <> "edz" Then
            rowWHTLINE1 = Nothing
        End If

        If rowWHTLINE1 IsNot Nothing Then

            ' ************************* Laser Printer Name *************************
            ASCMAIN1.LaserPrinterName = (rowWHTLINE1.Item("LASER_PRT_NAME") & String.Empty).ToString.Trim


            ' ************************* Label Printer *************************
            Try
                ASCMAIN1.LabelPrinterSerialPort = New System.IO.Ports.SerialPort
                Application.DoEvents()
                ASCMAIN1.LabelPrinterSerialPort.PortName = (rowWHTLINE1.Item("LABEL_PRT_COMM_PORT") & String.Empty).ToString.Trim
                Application.DoEvents()
                'Clean out the buffers
                If Not ASCMAIN1.Running_in_VS Then
                    ASCMAIN1.LabelPrinterSerialPort.Open()
                    ASCMAIN1.LabelPrinterSerialPort.DiscardInBuffer()
                    ASCMAIN1.LabelPrinterSerialPort.DiscardOutBuffer()
                End If

            Catch ex As Exception
                If ASCMAIN1.CLIENT = "RGI" Then
                    If stationID.StartsWith("WH0") OrElse stationID.StartsWith("SH0") Then
                        MessageBox.Show("Unable to connect to Label Printer: " & ex.Message, "Initialize", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
                ASCMAIN1.LabelPrinterSerialPort = Nothing
            End Try


            ' ************************* Laser Printer IP Address *************************
            ASCMAIN1.LaserPrinterIpAddress = (rowWHTLINE1.Item("LASER_PRT_IP_ADDRESS") & String.Empty).ToString.Trim
            ASCMAIN1.AltLaserPrinterIpAddress = (rowWHTLINE1.Item("LASER_PRT_ALT_IP_ADDRESS") & String.Empty).ToString.Trim

            Dim port As String = String.Empty
            If ASCMAIN1.LaserPrinterIpAddress.Contains(":") Then
                port = ASCMAIN1.LaserPrinterIpAddress.Split(":")(1)
                ASCMAIN1.LaserPrinterIpAddress = ASCMAIN1.LaserPrinterIpAddress.Split(":")(0)
            End If

            ' New code allows printer by name
            If ASCMAIN1.LaserPrinterIpAddress.Split(".").Length > 3 Then
                If Not Net.IPAddress.TryParse(ASCMAIN1.LaserPrinterIpAddress, Nothing) Then
                    ASCMAIN1.LaserPrinterIpAddress = String.Empty
                ElseIf port.Length > 0 Then
                    ASCMAIN1.LaserPrinterIpAddress &= ":" & port
                End If
            End If

            If ASCMAIN1.AltLaserPrinterIpAddress.Contains(":") Then
                port = ASCMAIN1.AltLaserPrinterIpAddress.Split(":")(1)
                ASCMAIN1.AltLaserPrinterIpAddress = ASCMAIN1.AltLaserPrinterIpAddress.Split(":")(0)
            End If

            ' New code allows printer by name
            If ASCMAIN1.AltLaserPrinterIpAddress.Split(".").Length > 3 Then
                If Not Net.IPAddress.TryParse(ASCMAIN1.AltLaserPrinterIpAddress, Nothing) Then
                    ASCMAIN1.AltLaserPrinterIpAddress = String.Empty
                ElseIf port.Length > 0 Then
                    ASCMAIN1.AltLaserPrinterIpAddress &= ":" & port
                End If
            End If

            If ASCMAIN1.LaserPrinterIpAddress.Length > 0 AndAlso ASCMAIN1.AltLaserPrinterIpAddress.Length = 0 Then
                ASCMAIN1.AltLaserPrinterIpAddress = ASCMAIN1.LaserPrinterIpAddress
            End If

            If ASCMAIN1.AltLaserPrinterIpAddress.Length > 0 AndAlso ASCMAIN1.LaserPrinterIpAddress.Length = 0 Then
                ASCMAIN1.LaserPrinterIpAddress = ASCMAIN1.AltLaserPrinterIpAddress
            End If


            ' ************************* Label Printer Name *************************
            Try
                ASCMAIN1.LabelPrinterName = rowWHTLINE1.Item("LABEL_PRT_NAME") & String.Empty
            Catch ex As Exception

            End Try


            ' ************************* Scale *************************
            Try
                ASFMAIN1.scaleComPort = rowWHTLINE1.Item("SCALE_COMM_PORT") & String.Empty
                ASFMAIN1.scaleport = New System.IO.Ports.SerialPort
                With ASFMAIN1.scaleport

                    .BaudRate = Val(rowWHTLINE1.Item("SCALE_BAUDRATE") & String.Empty)
                    If .BaudRate <= 0 Then
                        .BaudRate = 9600
                    End If

                    .DataBits = Val(rowWHTLINE1.Item("SCALE_DATABITS") & String.Empty)

                    Select Case rowWHTLINE1.Item("SCALE_STOPBITS") & String.Empty
                        Case "N"
                            .StopBits = IO.Ports.StopBits.None
                        Case "1"
                            .StopBits = IO.Ports.StopBits.One
                        Case "2"
                            .StopBits = IO.Ports.StopBits.Two
                        Case "P"
                            .StopBits = IO.Ports.StopBits.OnePointFive
                        Case Else
                            .StopBits = IO.Ports.StopBits.One
                    End Select

                    Select Case rowWHTLINE1.Item("SCALE_PARITY") & String.Empty
                        Case "N"
                            .Parity = IO.Ports.Parity.None
                        Case "E"
                            .Parity = IO.Ports.Parity.Even
                        Case "M"
                            .Parity = IO.Ports.Parity.Mark
                        Case "O"
                            .Parity = IO.Ports.Parity.Odd
                        Case "S"
                            .Parity = IO.Ports.Parity.Space
                        Case Else
                            .Parity = IO.Ports.Parity.None
                    End Select

                    .NewLine = vbCrLf

                    If .IsOpen Then
                        .Close()
                    End If

                    .PortName = ASFMAIN1.scaleComPort
                    .Open()
                End With

            Catch ex As Exception
                ASFMAIN1.scaleport = Nothing
                If ASCMAIN1.CLIENT = "RGI" Then
                    If stationID.StartsWith("WH0") OrElse stationID.StartsWith("SH0") Then
                        MessageBox.Show("Cannot connect to Scale Com Port: " & rowWHTLINE1.Item("SCALE_COMM_PORT") & ". " & ex.Message, "Initialize", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            End Try
        End If

    End Sub

    Public Overrides Sub Get_Column_Expression_Exceptions(ByVal FORM_NAME As String, ByVal DATA_SOURCE As String, ByVal COLUMN_NAME As String, ByRef sql_SELECT_col As String) ' , ByRef sql_GROUP_BY_col As String)
        Select Case FORM_NAME

            Case "GLFASUM1"
                Select Case COLUMN_NAME
                    Case "DATA_TYPE"
                        sql_SELECT_col = "'" & DATA_SOURCE & "'"
                        'sql_SELECT_col = "'" & DATA_SOURCE & "' DATA_TYPE"
                        'sql_GROUP_BY_col = "'" & DATA_SOURCE & "'"
                End Select

            Case "ICFTRNS1"
                Select Case COLUMN_NAME
                    Case "ITEM_CODE"
                        '                        xInfo(1) = "1"
                End Select

            Case "SOFWHOD1"
                Select Case COLUMN_NAME
                    Case "ITEM_CODE"
                        '                        If DATA_SOURCE <> "A" Then xInfo(1) = "G" & CStr(j)
                End Select
        End Select
    End Sub

    Public Overrides Function Get_Code_SQL_X(ByVal FORM_NAME As String, ByVal COLUMN_NAME As String, ByRef GROUP_KEY As String) As String
        Dim sql As String = ""

        ' Set up ASWGROUP w/codes & descs from all small tables

        GROUP_KEY = COLUMN_NAME

        Select Case COLUMN_NAME
            'Case "ACCT_CODE"
            '    sql = "Select ACCT_CODE, ACCT_DESC from GLTACCT1"
            Case "ACCT_TYPE" ' SHOULD BE GETTING THIS FROM THE VIEW IN THE CODE BELOW, BUT THAT CODE IS NOT SMART ENOUGH TO CONSTRUCT A SQL STMT FROM ASTVIEW1/2 FOR ASTCODE1 TYPES
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE ACCT_TYPE, T_DESC ACCT_TYPE_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'GLTACCT1' AND COLUMN_NAME = 'ACCT_TYPE'"
            Case "BRKR_TYPE"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE BRKR_TYPE, T_DESC BRKR_TYPE_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'BATBRKR1' AND COLUMN_NAME = 'BRKR_TYPE'"
            Case "CUST_BILL_TO_CUST"
                GROUP_KEY = "CUST_CODE"
                sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
            Case "CUST_BUYING_GROUP"
                GROUP_KEY = "CUST_CODE"
                sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
            Case "CUST_CITY"
                sql = "Select CUST_CITY, CUST_CITY from ARTCUST1"
                'Case "CUST_CLASS_CODE"
                '    sql = "Select CUST_CLASS_CODE, CUST_CLASS_DESC from ARTCLAS1"
                'Case "CUST_CODE"
                '    sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
            Case "CUST_CODE_SO"
                GROUP_KEY = "CUST_CODE"
                sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"

            Case "CUST_SREP_CODE"
                GROUP_KEY = "SREP_CODE"
                sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
            Case "CUST_STORE_GROUP"
                GROUP_KEY = "CUST_CODE || '-' || CUST_STORE_GROUP"
                sql = "Select CUST_STORE_GROUP, CUST_STORE_GROUP_NAME from ARTCUST8"
            Case "CUST_STORE_NO"
                GROUP_KEY = "CUST_CODE || '-' || CUST_STORE_NO"
                sql = "Select CUST_STORE_NO, DECODE (CUST_STORE_LOCATION, NULL, CUST_STORE_NAME, CUST_CODE || ':' || CUST_STORE_LOCATION) CUST_STORE_NAME from ARTCUST2"
            Case "CUST_STORE_STATE"
                GROUP_KEY = "STATE_CODE"
                sql = "Select STATE_CODE, STATE_NAME from TATSTATE"
                'Case "DIVISION_CODE"
                '    sql = "Select DIVISION_CODE, DIVISION_NAME from SOTSDIV1"
            Case "DIVISION_CODE_O"
                GROUP_KEY = "DIVISION_CODE"
                sql = "Select DIVISION_CODE, DIVISION_NAME from SOTSDIV1"
                'Case "DMA_CODE"
                '    sql = "Select DMA_CODE, DMA_DESC from SOTDMAC1"
                'Case "FUND_CODE"
                '    sql = "Select FUND_CODE, FUND_DESC from SPTFUND1"
            Case "INV_PYMT_METHOD"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE INV_PYMT_METHOD, T_DESC INV_PYMT_METHOD_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'APTINVH1' AND COLUMN_NAME = 'INV_PYMT_METHOD'"
                'Case "ITEM_CODE"
                '    sql = "Select ITEM_CODE, ITEM_DESC from ICTITEM0"
                'Case "MARKET_CODE"
                '    sql = "Select MARKET_CODE, MARKET_DESC from SOTMKTC1"
            Case "OPS_YYYYPP"
                sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2"
            Case "PO_DIV_CODE"
                sql = "Select DIVISION_CODE, DIVISION_NAME from SOTSDIV1"

            Case "PARTNER_CODE"
                sql = "SELECT ""PartnerKEY"" PARTNER_CODE, ""PartnerName"" PARTNER_NAME FROM GEN.""Partner_tb"""

            Case "PO_ORDER_TYPE"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE PO_ORDER_TYPE, T_DESC PO_ORDER_TYPE_DESC" _
                & " FROM ASTCODE1 WHERE TABLE_NAME = 'POTORDR1' AND COLUMN_NAME = 'PO_ORDER_TYPE'"
            Case "POST_CODE"
                If FORM_NAME Like "AP*" Then
                    sql = "Select POST_CODE, POST_DESC from APTPOST1"
                ElseIf FORM_NAME Like "AR*" Then
                    sql = "Select POST_CODE, POST_DESC from ARTPOST1"
                End If
            Case "PROCESSOR_CODE"
                GROUP_KEY = "USER_ID"
                sql = "Select USER_ID, USER_NAME from ASTUSER1"
                'Case "PROD_CODE"
                '    sql = "Select PROD_CODE, PROD_DESC from ICTPROD1"
                'Case "REGION_CODE"
                '    sql = "Select REGION_CODE, REGION_DESC from SOTSREG1"
                'Case "SALES_DIVISION_CODE"
                '    sql = "Select SALES_DIVISION_CODE, SALES_DIVISION_NAME from SOTSDIV1"
            Case "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"
                GROUP_KEY = "ACCT_SEG_CODE"
                sql = "Select ACCT_SEG_CODE, ACCT_SEG_DESC from GLTSEGM1 where ACCT_SEG_ID = '" & Mid(COLUMN_NAME, 4, 1) & "'"
            Case "SREP_CODE", "CUST_SREP_CODE", "SELL_CODE"
                GROUP_KEY = "SREP_CODE"
                sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
            Case "SREP_TYPE"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE SREP_TYPE, T_DESC SREP_TYPE_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'SOTSREP1' AND COLUMN_NAME = 'SREP_TYPE'"
                'Case "STATE_CODE"
                '    sql = "Select STATE_CODE, STATE_NAME from TATSTATE"
                'Case "TERM_CODE"
                '    sql = "Select TERM_CODE, TERM_DESC from TATTERM1"
                'Case "TRADE_CLASS_CODE"
                '    sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC from SOTTCLS1"
                'Case "VEND_CLASS_CODE"
                '    sql = "Select VEND_CLASS_CODE, VEND_CLASS_DESC from APTCLAS1"
                'Case "VEND_CODE"
                '    sql = "Select VEND_CODE, VEND_NAME from APTVEND1"
            Case "VEND_TYPE"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE VEND_TYPE, T_DESC VEND_TYPE_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'APTVEND1' AND COLUMN_NAME = 'VEND_TYPE'"
            Case "VEND_PYMT_METHOD"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE VEND_PYMT_METHOD, T_DESC VEND_PYMT_METHOD_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'APTVEND1' AND COLUMN_NAME = 'VEND_PYMT_METHOD'"
                'Case "WHSE_CODE"
                '    sql = "Select WHSE_CODE, WHSE_DESC from ICTWHSE1"
                'Case "ZONE_CODE"
                '    sql = "Select ZONE_CODE, ZONE_DESC from ICTZONE1"
            Case Else
                ' Stop
        End Select

        If sql = "" Then
            ASCMAIN1.sql = "SELECT * FROM ASTVIEW2 WHERE VIEW_NAME = '" & COLUMN_NAME & "' AND COLUMN_POSITION IN (1,2)"
            Dim tbl As DataTable = ASCDATA1.GetDataTable
            If tbl.Rows.Count > 2 Then
                Dim MENU_ITEM_OBJECT As String = ""
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    MENU_ITEM_OBJECT = ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT
                Else
                    MENU_ITEM_OBJECT = ASCMAIN1.MENU_ITEM_OBJECT
                End If
                'ASCMAIN1.ActiveForm IS NOTHING WHEN YOU RUN AHA.SOFRMAF1 FROM MENU AS THE 1ST FORM, AND THEN YOU BLOW UP ON NEXT LINE
                ' tbl = New DataView(tbl, "TABLE_NAME LIKE '" & Mid(ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT, 1, 2) & "%'", "COLUMN_POSITION", DataViewRowState.CurrentRows).ToTable
                tbl = New DataView(tbl, "TABLE_NAME LIKE '" & Mid(MENU_ITEM_OBJECT, 1, 2) & "%'", "COLUMN_POSITION", DataViewRowState.CurrentRows).ToTable
                'ASCMAIN1.sql &= " and TABLE_NAME LIKE '" & Mid(ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT, 1, 2) & "%'"
                'tbl = ASCDATA1.GetDataTable
            End If
            GROUP_KEY = ""
            If tbl.Rows.Count = 2 Then
                For Each row As DataRow In tbl.Select("", "COLUMN_POSITION")
                    If GROUP_KEY = "" Then GROUP_KEY = row.Item("COLUMN_NAME")
                    sql &= "," & row.Item("COLUMN_NAME")
                Next

                If GROUP_KEY = "T_CODE" Then ' NOT WORKING TOO WELL FOR ASTCODE1 - NEEDS A FANCIER SQL ASSEMBLER, USING ASTCODE1 INSTEAD OF TABLE_NAME, A WHERE CLAUSE, AND POSSIBLY COLUMN ALIASES
                    GROUP_KEY = ""
                    sql = ""
                Else
                    sql = "Select " & Mid(sql, 2) & " from " & tbl.Rows(0).Item("TABLE_NAME")
                End If

            Else
                ' NEED TO USE VIEW_NAME AND TABLE_NAME FROM THE REPORT DEFINITION FOR COLUMNS THAT MIGHT BE IN MULTIPLE MASTER TABLES, LIKE CLASS_CODE, WHICH MIGHT BE AR OR AP

            End If

        End If


        Get_Code_SQL_X = sql
    End Function

    Public Overrides Sub Write_Group_Record_X(ByVal GROUP_KEY As String, ByVal COLUMN_NAME As String, ByVal GROUP_CODEs As ArrayList, ByVal GROUP_DESCs As ArrayList)
        Select Case COLUMN_NAME
            Case "ACCT_TYPE"
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":1A", "A", "Asset")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":2L", "L", "Liability")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":3E", "E", "Equity")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":4I", "I", "Income")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":5X", "X", "Expense")

            Case "EXC_OBS_IND"
                GROUP_CODEs.Add("X") : GROUP_DESCs.Add("Excess Inventory On Hand")
                GROUP_CODEs.Add("O") : GROUP_DESCs.Add("Obsolete Inventory (No Demand)")

                'Case "MOP"
                '    GROUP_CODEs.Add("F") : GROUP_DESCs.Add("Farm Raised")
                '    GROUP_CODEs.Add("W") : GROUP_DESCs.Add("Wild Caught")
                '    GROUP_CODEs.Add("N") : GROUP_DESCs.Add("N/A")


        End Select
    End Sub

    Public Overrides Function CodeValues(ByVal TABLE_COLUMN As String) As Dictionary(Of String, String)

        Dim TABLE_NAME As String = ""
        Dim COLUMN_NAME As String = TABLE_COLUMN
        If InStr(COLUMN_NAME, ".") <> 0 Then
            TABLE_NAME = Split(TABLE_COLUMN, ".")(0)
            COLUMN_NAME = Split(TABLE_COLUMN, ".")(1)
        End If

        Dim VL As New Dictionary(Of String, String)

        Select Case COLUMN_NAME

            Case "ACCRUAL_STATUS"
                VL.Add("0", "Open")
                VL.Add("1", "Closed")
                VL.Add("2", "Netted")
                VL.Add("X", "Voided")

            Case "CCPA_REASON"
                VL.Add("A", "Auto Stmt")
                VL.Add("M", "Manual")
                VL.Add("B", "Bank")
                VL.Add("S", "Statement")
                VL.Add("O", "Order")
                VL.Add("C", "Sale Captured")

            Case "CCPA_STATUS"
                VL.Add("0", "Sales Q")
                'VL.Add("1", "In Queue")
                VL.Add("E", "Declined")
                VL.Add("A", "Approved")
                VL.Add("S", "Settled")
                VL.Add("V", "Voided")
                VL.Add("D", "Deleted")
                VL.Add("T", "Authorized")

            Case "CCPA_TYPE"
                VL.Add("A", "Auth")
                VL.Add("S", "Sale")
                VL.Add("V", "Void")
                VL.Add("C", "Credit")

            Case "COBRA_TYPE"
                VL.Add("S", "Self")
                VL.Add("P", "Plan")


            Case "COMM_PYBL_STATUS"
                VL.Add("O", "Open")
                VL.Add("P", "Paid")
                VL.Add("V", "Voided")

            Case "COMM_STATUS"
                VL.Add("A", "Active")
                VL.Add("D", "Discontinued")
                VL.Add("T", "Terminated")
                VL.Add("M", "Modified")
                VL.Add("R", "Renewed")
                VL.Add("V", "Voided")

            Case "CONT_STATUS"
                VL.Add("A", "Active")
                VL.Add("D", "Discontinued")
                VL.Add("T", "Terminated")
                VL.Add("M", "Modified")
                VL.Add("R", "Renewed")
                VL.Add("V", "Voided")

            Case "COBRA_TYPE"
                VL.Add("S", "Self")
                VL.Add("P", "Plan")
                VL.Add("3", "3rdP")

            Case "COOL_COMPLIANT"
                VL.Add("1", "COOL Compliant")
                VL.Add("0", "Not COOL")

            Case "CUST_STATUS"
                VL.Add("A", "Active")
                VL.Add("I", "Inactive")
                VL.Add("C", "Closed")

            Case "CUST_COMMENT_KEY"
                VL.Add("SOE", "Sales Order Entry")
                VL.Add("BILLING", "Billing")
                VL.Add("AR", "A/R")
                VL.Add("CREDIT", "Credit")
                VL.Add("M", "Misc")

            Case "CONTACT_TYPE"
                VL.Add("B", "Buyer")
                VL.Add("P", "A/P")
                VL.Add("W", "Whse")
                VL.Add("M", "Misc")
                VL.Add("X", "Master")

            Case "GROUP_STATUS"
                VL.Add("A", "Active")
                VL.Add("D", "Discontinued")

            Case "IN_OR_OUT"
                VL.Add("I", "In")
                VL.Add("O", "Out")
                VL.Add("A", "Adj")


            Case "ITEM_STATUS"
                VL.Add("A", "Active")
                VL.Add("I", "Inactive")

            Case "JOB_STREAM_TYPE"
                VL.Add("S", "Standard")
                VL.Add("W", "Web Reports")

            Case "MAILING_SET_DIST_BY"
                VL.Add("M", "Mail")
                VL.Add("E", "email")
                VL.Add("F", "Fax")
                VL.Add("N", "None")
                VL.Add("P", "Pickup")

            Case "MAILING_SET_TYPE"
                VL.Add("B", "Billing")
                VL.Add("F", "FTR")
                VL.Add("S", "SLR")

            Case "MOP"
                VL.Add("F", "Farm Raised")
                VL.Add("W", "Wild Caught")
                VL.Add("N", "N/A")

            Case "ORDR_LINE_STATUS"
                VL.Add("O", "Open")
                VL.Add("B", "BackOrder")
                VL.Add("P", "In Pick")
                VL.Add("F", "Completed")
                VL.Add("C", "Cancelled")

            Case "ORDR_CREDIT_STATUS"
                VL.Add("W", "Pending")
                VL.Add("A", "Approved")
                VL.Add("R", "Rejected")

            Case "ORDR_SOURCE"
                VL.Add("K", "Keyed In")
                VL.Add("E", "EDI")
                VL.Add("W", "Web")
                VL.Add("X", "XML")
                VL.Add("O", "OptiPort")
                VL.Add("V", "V-Web")
                VL.Add("I", "I-Scan")

            Case "ORDR_STATUS"
                VL.Add("O", "Open")
                VL.Add("P", "Released")
                VL.Add("F", "Completed")
                VL.Add("V", "Voided")

            Case "ORDR_UNIT_PRICE_SOURCE"
                VL.Add("L", "Lens Bank")
                VL.Add("C", "Column")
                VL.Add("P", "Price Catgy")
                VL.Add("B", "Stock Buy")
                VL.Add("A", "Annual Supply")
                VL.Add("M", "Override")
                VL.Add("S", "Promotion")
                VL.Add("T", "Pass-Thru")

            Case "PKG_TYPE"
                VL.Add("CTN", "CTN")
                VL.Add("PLT", "PLT")

            Case "PAY_OR_REV"
                VL.Add("P", "Pay")
                VL.Add("R", "Reverse")

            Case "PICK_STATUS"
                VL.Add("P", "In Pick")
                VL.Add("F", "Shipped")
                VL.Add("D", "Deleted")

                'Case "PO_ORDER_TYPE"
                '    VL.Add("R", "Rx")
                '    VL.Add("S", "Stock")

            Case "PO_STATUS"
                ' not sure of X and C
                VL.Add("C", "Closed")
                VL.Add("X", "Cancelled")
                VL.Add("O", "Open")

            Case "PO_STATUS_CODE"
                ' not sure of X and C
                VL.Add("X", "Closed")
                VL.Add("C", "Cancelled")
                VL.Add("O", "Open")

            Case "PROMO_TYPE"
                VL.Add("L", "Lens Bank")
                'VL.Add("B", "Stock Buy")
                VL.Add("A", "Annual Supply")

            Case "PROD_CATGY_STATUS"
                VL.Add("A", "Active")
                VL.Add("D", "Discontinued")
                VL.Add("I", "In-Development")

            Case "PROM_STATUS"
                VL.Add("A", "Active")
                VL.Add("D", "Discontinued")

            Case "REBATE_TYPE"
                VL.Add("U", "$/Unit")
                VL.Add("O", "$/Order")

            Case "RESPONSE_CODE"
                VL.Add("A", "Approved")
                VL.Add("E", "Error")

            Case "RTV_TYPE"
                VL.Add("DEF", "Defective")
                VL.Add("RTV", "RTV")

            Case "SUBSCRIBER_STATUS"
                VL.Add("A", "Active")
                VL.Add("C", "Cobra")
                VL.Add("L", "Leave")
                VL.Add("T", "Terminated")

            Case "TKT_IND"
                VL.Add("0", "Not Queued")
                VL.Add("1", "Queued")
                VL.Add("2", "Printed")

            Case "TRAN_SOURCE"
                VL.Add("A", "A-Rtn")
                VL.Add("B", "B-Rtn")
                VL.Add("C", "C-Rtn")
                VL.Add("E", "Entry")
                VL.Add("P", "StkPO")
                VL.Add("S", "Order")
                VL.Add("T", "BO-Xfr")
                VL.Add("X", "Rx-PO")
                VL.Add("V", "RTV")

                'Case "TRAN_TYPE"
                '    VL.Add("S", "Sale")
                '    VL.Add("R", "Return")
                '    VL.Add("P", "Rec")
                '    VL.Add("A", "Adj")
                '    VL.Add("T", "Xfr")
                '    VL.Add("V", "RTV")

        End Select

        Return VL

    End Function

    Public Shared Function Send_email_with_Attachment(
        ByVal frmASFBASE0 As ASFBASE0,
        ByVal FILENAME As String,
        ByVal ATTACHMENT As String,
        ByVal SUBJECT As String,
        Optional EMAIL_ADDRESS As String = "",
        Optional EMAIL_NAME As String = "",
        Optional EMAIL_KEY As String = "",
        Optional ENTITY_KEY As String = "",
        Optional ENTITY_NAME As String = "",
        Optional ENTITY_CAPTION As String = "")

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        ATTACHMENTs.Add(ATTACHMENT, FILENAME)

        SUBJECT = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " " & SUBJECT

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        EMAIL_ADDRESSs.Add(EMAIL_ADDRESS, EMAIL_NAME)

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, EMAIL_KEY, False, True, ENTITY_KEY, ENTITY_NAME, ENTITY_CAPTION)

        Return SEND_NO

    End Function

    Public Overrides Function Send_email(ByVal frmASFBASE0 As ASFBASE0,
                                 ByVal EMAIL_ADDRESSs As Dictionary(Of String, String),
                                 ByVal ATTACHMENTs As Dictionary(Of String, String),
                                 ByVal SUBJECT As String,
                                 ByVal EMAIL_KEY As String,
                                 Optional ByVal auto_send As Boolean = False,
                                 Optional SEND_CC_to_USER_ID As Boolean = False,
                                 Optional ENTITY_KEY As String = "",
                                 Optional ENTITY_NAME As String = "",
                                 Optional ENTITY_CAPTION As String = "",
                                 Optional EMAIL_BODY As String = "") As String

        Dim USER_ID_emailer As String = ASCMAIN1.USER_ID
        Dim rowASTUSER1_EMAIL_FROM As DataRow = frmASFBASE0.LookUp("ASTUSER1", USER_ID_emailer, True)
        Dim rowASTUSER1_EMAIL_BCC As DataRow = Nothing

        Dim USER_TELEPHONE As String = rowASTUSER1_EMAIL_FROM.Item("USER_TELEPHONE") & ""
        Dim USER_EXT As String = rowASTUSER1_EMAIL_FROM.Item("USER_EXT") & ""
        Dim USER_FAX As String = rowASTUSER1_EMAIL_FROM.Item("USER_FAX") & ""
        'Dim EMAIL_BODY As String = "Attached is the file that you have requested."

        Dim rowTATMAIL1 As DataRow = frmASFBASE0.LookUp("TATMAIL1", EMAIL_KEY)
        If rowTATMAIL1 IsNot Nothing Then
            If rowTATMAIL1.Item("EMAIL_FROM") & "" <> "" Then
                ' IF THERE IS A BAD USER ID IN TATMAIL1, THEN THE MAIL WON'T GO, AND YOU WON'T KNOW WHY - SO BETTER TO GET AN EXCEPTION IF TATMAIL1 EMAIL_FROM IS CONFIGURED WITH A BAD USER_ID
                rowASTUSER1_EMAIL_FROM = frmASFBASE0.LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_FROM")) ' frmASFBASE0.LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_FROM"), True)
            End If
            If rowTATMAIL1.Item("EMAIL_BCC") & "" <> "" Then
                rowASTUSER1_EMAIL_BCC = frmASFBASE0.LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_BCC")) '  frmASFBASE0.LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_BCC"), True)
            End If
            If EMAIL_BODY = "" Then
                If rowTATMAIL1.Item("EMAIL_BODY") & "" <> "" Then
                    EMAIL_BODY = rowTATMAIL1.Item("EMAIL_BODY")
                Else
                    EMAIL_BODY = "Attached is the file that you have requested."
                End If

            End If

        End If

        Dim USER_SIGNATURE As String =
          rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & vbCrLf _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & vbCrLf, "") _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & vbCrLf, "") _
        & IIf(USER_TELEPHONE <> "", "Tel: " & ASCMAIN1.FormatTel(USER_TELEPHONE, USER_EXT) & vbCrLf, "") _
        & IIf(USER_FAX <> "", "Fax: " & ASCMAIN1.FormatTel(USER_FAX) & vbCrLf, "") _
        & rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & vbCrLf

        Dim frmTAFSEND1 As New TAFSEND1(frmASFBASE0)
        frmTAFSEND1.EMAIL_KEY = EMAIL_KEY
        frmTAFSEND1.SEND_FROM = rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & ""
        frmTAFSEND1.SEND_FROM_NAME = rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & ""
        frmTAFSEND1.SEND_FROM_SIGNATURE = USER_SIGNATURE
        frmTAFSEND1.SEND_TOs = EMAIL_ADDRESSs
        frmTAFSEND1.SEND_TO = ""
        frmTAFSEND1.SEND_TO_NAME = ""
        frmTAFSEND1.SEND_CC = ""
        If SEND_CC_to_USER_ID Then
            frmTAFSEND1.SEND_CC = ASCMAIN1.USER_EMAIL
            frmTAFSEND1.SEND_CC_NAME = ASCMAIN1.USER_NAME
        End If
        If rowASTUSER1_EMAIL_BCC IsNot Nothing Then
            frmTAFSEND1.SEND_BCC = rowASTUSER1_EMAIL_BCC.Item("USER_EMAIL") & ""
            frmTAFSEND1.SEND_BCC_NAME = rowASTUSER1_EMAIL_BCC.Item("USER_NAME") & ""
        End If

        frmTAFSEND1.SEND_SUBJECT = SUBJECT

        frmTAFSEND1.SEND_BODY = EMAIL_BODY
        frmTAFSEND1.SEND_ENTITY_KEY = ENTITY_KEY
        frmTAFSEND1.SEND_ENTITY_NAME = ENTITY_NAME
        frmTAFSEND1.SEND_METHOD = "E"
        frmTAFSEND1.SEND_ENTITY_CAPTION = ENTITY_CAPTION
        frmTAFSEND1.SEND_ATTACHMENTs = ATTACHMENTs
        frmTAFSEND1.SEND_ATTACHMENT = ""

        If auto_send Then
            frmTAFSEND1.Send_email_automatically()
        Else
            frmTAFSEND1.ShowDialog()
        End If

        Dim SEND_STATUS As String = frmTAFSEND1.SEND_STATUS
        Dim SEND_NO As String = frmTAFSEND1.SEND_NO

        frmTAFSEND1.Dispose()
        frmTAFSEND1 = Nothing

        Return SEND_NO ' SEND_STATUS

    End Function

    Public Overrides Sub Application_Initialization()
        'PUT CODE HERE FOR WHEN YOU ARE LOGGING IN
        'Dim rowICTPARM1 As DataRow = ASCDATA1.GetDataRow("Select * from ICTPARM1 where IC_PARM_KEY = 'Z'")
        'If rowICTPARM1.Item("IC_PARM_CYW_LAST") & "" <> ASCMAIN1.CYW Then
        '    ASCMAIN1.sql = "Select * from GLTCOMP1 where COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        '    Dim rowGLTCOMP1 As DataRow = ASCDATA1.GetDataRow
        '    If rowGLTCOMP1 IsNot Nothing Then
        '        ASCMAIN1.Progress("Now Taking Weekly Inventory Snapshot")
        '        ASCDATA1.ExecuteSP("ICPLOTD9")
        '        ASCMAIN1.Progress("")
        '    End If
        'End If
    End Sub

    Public Overrides Sub Maintain_Contacts(ByVal frmASFBASE1 As ASFBASE1,
                                           ByVal CONTACT_ENTITY_TABLE As String,
                                           ByVal CONTACT_ENTITY_KEY As String,
                                           ByVal CONTACT_ENTITY_NAME As String)
        'MyBase.Maintain_Contacts()
        Using frmTAFCONT1 As New TAFCONT1(frmASFBASE1)
            With frmTAFCONT1
                .CONTACT_ENTITY_TABLE = CONTACT_ENTITY_TABLE
                .CONTACT_ENTITY_KEY = CONTACT_ENTITY_KEY
                .CONTACT_ENTITY_NAME = CONTACT_ENTITY_NAME
                .ShowDialog()
            End With
        End Using

    End Sub

    Public Overrides Function Custom_sqlwhere(
    ByVal sqlwhere As String,
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal COLUMN_NAME As String) As String
        If grd.Name = "grd" And grd.TopLevelControl.Name = "ASFCODE1" _
             And ASCMAIN1.CodeSelector IsNot Nothing _
             And ASCMAIN1.CodeSelector.VIEW_NAME = "CUST_CODE" _
             And ASCMAIN1.CodeSelector.TABLE_NAME = "ARTCUST1" _
             And COLUMN_NAME = "CUST_NAME" Then
            sqlwhere = Mid(sqlwhere, 6)
            sqlwhere = " and (" & sqlwhere & " or " & Replace(sqlwhere, "CUST_NAME", "CUST_DBA_NAME") & ")"

            If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsFilterRow Then
                grd.DisplayLayout.Bands(0).ColumnFilters("CUST_NAME").ClearFilterConditions()
                'grd.ActiveRow.Cells("CUST_NAME").Column.F.Value = DBNull.Value
            End If

        End If


        Return sqlwhere
    End Function

    Public Shared Function Get_EDI_Custs(EDI_DOC_NO As String) As List(Of String)

        Dim EDI_Custs As New List(Of String)

        ASCMAIN1.sql = "Select DISTINCT CUST_CODE from EDTTRPM1 where CUST_CODE is Not Null"
        If EDI_DOC_NO <> "" Then
            ASCMAIN1.sql &= " and EDI_DOC_NO = '" & EDI_DOC_NO & "'"
        End If
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            EDI_Custs.Add(row.Item("CUST_CODE"))
        Next
        Return EDI_Custs
    End Function

    Public Shared Function Calculate_INV_DUE_DATE(
                                                 F As ASFBASE0,
                                                 TERM_CODE As String,
                                                 rowTATTERM1 As DataRow,
                                                 INV_BASE_DATE As Object) As Date

        Dim INV_DUE_DATE As Object = Nothing

        If TERM_CODE = "" Or INV_BASE_DATE Is Nothing Then
            Return INV_BASE_DATE
            'Exit Function
        End If

        If rowTATTERM1 Is Nothing Then
            rowTATTERM1 = F.LookUp("TATTERM1", TERM_CODE, True)
        End If

        Select Case rowTATTERM1.Item("TERM_DUE_TYPE") & ""
            Case "C" ' IE, COD
                INV_DUE_DATE = INV_BASE_DATE

            Case "D"
                INV_DUE_DATE = INV_BASE_DATE.AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DUE") & ""))

            Case "S"
                If rowTATTERM1.Item("TERM_CUTOFF_DATE") & "" <> "" Then
                    Dim TERM_CUTOFF_DATE As Date = rowTATTERM1.Item("TERM_CUTOFF_DATE")
                    If Format(INV_BASE_DATE, "MMdd") > Format(TERM_CUTOFF_DATE, "MMdd") Then
                        INV_DUE_DATE = CDate(Format(TERM_CUTOFF_DATE, "MM/dd") & "/" & Format(Val(Format(INV_BASE_DATE, "yyyy")) + 1, "0000"))

                    Else
                        INV_DUE_DATE = CDate(Format(TERM_CUTOFF_DATE, "MM/dd") & "/" & Format(INV_BASE_DATE, "yyyy"))
                    End If

                End If

            Case "E"
                Dim ADD_MONTHS_BASE As Integer = 1
                Dim TERM_CUTOFF_DAY As Integer = Val(rowTATTERM1.Item("TERM_CUTOFF_DAY") & "")
                Dim BASE_DD As Integer = Val(Format(INV_BASE_DATE, "dd"))
                Dim TERM_DAYS_DUE As Integer = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & "")
                Dim TERM_ADDL_MOS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_MOS") & "")
                Dim TERM_ADDL_DAYS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_DAYS") & "")
                Dim INV_BASE_DATEx As String = Format(INV_BASE_DATE, "MM/dd/yyyy")

                Select Case rowTATTERM1.Item("TERM_EOM_TYPE") & ""
                    Case "F"
                        ASCMAIN1.sql = "Select GLTPARM2.* " _
                         & " from GLTPARM2 " _
                         & " where OPS_YYYYPP = " _
                         & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
                         & "  where GLTPARM2.PRD_END_DATE >= '" & Format(INV_BASE_DATE, "dd-MMM-yyyy") & "')"
                        Dim rowGLTPARM2 As DataRow = ASCDATA1.GetDataRow
                        Dim YYYYMM As String = ASCMAIN1.Get_YYYYMM(rowGLTPARM2.Item("OPS_YYYYPP"), 0)
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)

                    Case "C"
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)

                    Case "S"
                        'If BASE_DD <= TERM_CUTOFF_DAY _
                        'And BASE_DD <= TERM_DAYS_DUE Then
                        '    ADD_MONTHS_BASE = 0
                        'End If
                        If BASE_DD > TERM_CUTOFF_DAY Then
                            ADD_MONTHS_BASE = 2
                        End If
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)

                    Case Else
                        INV_DUE_DATE = INV_BASE_DATE

                End Select
                If TERM_ADDL_MOS > 0 Then
                    INV_DUE_DATE = CDate(INV_DUE_DATE).AddMonths(TERM_ADDL_MOS)
                End If
                If TERM_ADDL_DAYS > 0 Then
                    INV_DUE_DATE = CDate(INV_DUE_DATE).AddDays(TERM_ADDL_DAYS)
                End If
        End Select

        Return INV_DUE_DATE

    End Function

    Public Shared Function testLabel() As String

        Dim labelImage As String = String.Empty
        Dim lineNumber As Int16 = 180

        labelImage = "EPL2" & Environment.NewLine
        labelImage &= "S4" & Environment.NewLine
        labelImage &= "UN" & Environment.NewLine
        labelImage &= "WN" & Environment.NewLine
        'labelImage &= "ZB" & Environment.NewLine
        labelImage &= "ZT" & Environment.NewLine
        labelImage &= "N" & Environment.NewLine

        labelImage &= "A50,100,0,4,1,1,N," & Chr(34) & "Congratulations" & Chr(34) & Environment.NewLine
        lineNumber += 40
        labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "If you can read this information," & Chr(34) & Environment.NewLine
        lineNumber += 40
        labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "you have properly connected to" & Chr(34) & Environment.NewLine
        lineNumber += 40
        labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "your Session label Printer:" & Chr(34) & Environment.NewLine
        lineNumber += 40

        If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & ASCMAIN1.LabelPrinterSerialPort.PortName & Chr(34) & Environment.NewLine
        ElseIf ASCMAIN1.LabelPrinterName.Length > 0 Then
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & ASCMAIN1.LabelPrinterName & Chr(34) & Environment.NewLine
        End If

        lineNumber += 40

        For i As Integer = 1 To 10
            lineNumber += 40
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "Test" & Chr(34) & Environment.NewLine
            labelImage &= "A150," & lineNumber & ",0,4,1,1,N," & Chr(34) & "Line" & Chr(34) & Environment.NewLine
            labelImage &= "A300," & lineNumber & ",0,4,1,1,N," & Chr(34) & Format$(i, "000") & Chr(34) & Environment.NewLine
        Next

        labelImage &= "P1" & Environment.NewLine

        Return labelImage
    End Function

    Public Shared Function PalletCartonLabel(ByVal ShipFromName As String,
                                             ByRef rowICTWHSE1 As DataRow,
                                             ByRef tblSOTCART2 As DataTable,
                                             ByVal PoNumber As String,
                                             ByVal CartonNumber As String,
                                             ByVal Cartonsize As String,
                                             ByVal CartonCount As String,
                                             ByVal ShipmentNo As String) As String

        Dim labelImage As String = String.Empty
        Dim lineNumber As Int16 = 100

        labelImage = "EPL2" & Environment.NewLine
        labelImage &= "S4" & Environment.NewLine
        labelImage &= "UN" & Environment.NewLine
        labelImage &= "WN" & Environment.NewLine
        'labelImage &= "ZB" & Environment.NewLine
        labelImage &= "ZT" & Environment.NewLine
        labelImage &= "N" & Environment.NewLine

        labelImage &= "A50,100,0,4,1,1,N," & Chr(34) & ShipFromName & Chr(34) & Environment.NewLine
        lineNumber += 40

        For Each fieldName As String In New String() {"WHSE_ADDR1", "WHSE_ADDR2", "WHSE_ADDR3"}
            If rowICTWHSE1.Item(fieldName) & String.Empty <> String.Empty Then
                labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & rowICTWHSE1.Item(fieldName) & String.Empty & Chr(34) & Environment.NewLine
                lineNumber += 40
            End If
        Next

        Dim cityStateZip As String = rowICTWHSE1.Item("WHSE_CITY") & ", " & rowICTWHSE1.Item("WHSE_STATE") & "  " & rowICTWHSE1.Item("WHSE_ZIP_CODE")

        labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & cityStateZip & Chr(34) & Environment.NewLine
        lineNumber += 80
        labelImage &= "LO0," & lineNumber & ",800,4" & Environment.NewLine

        lineNumber += 40

        If PoNumber.Length > 0 Then
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "PO Number: " & PoNumber & Chr(34) & Environment.NewLine
            lineNumber += 40
        End If

        If CartonNumber.Length > 0 Then
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "Carton Number: " & CartonNumber & Chr(34) & Environment.NewLine
            lineNumber += 40
        End If

        If Cartonsize.Length > 0 Then
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "Carton Size: " & Cartonsize & Chr(34) & Environment.NewLine
            lineNumber += 40
        End If

        If ShipmentNo.Length > 0 Then
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "Shipment No: " & ShipmentNo & Chr(34) & Environment.NewLine
            lineNumber += 40
            labelImage &= "B10,10,0,PL,5,5,5,N," & Chr(34) & ShipmentNo & Chr(34) & Environment.NewLine
            lineNumber += 40
        End If

        lineNumber += 40
        labelImage &= "LO0," & lineNumber & ",800,4" & Environment.NewLine
        lineNumber += 40

        If CartonCount.Length > 0 Then
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "Carton Count: " & CartonCount & Chr(34) & Environment.NewLine
            lineNumber += 40
        End If

        If tblSOTCART2 IsNot Nothing AndAlso tblSOTCART2.Select("CART_NO = '" & CartonNumber & "'").Length > 0 Then
            lineNumber += 40
            labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & "Contents" & Chr(34) & Environment.NewLine
            lineNumber += 40
            For Each rowSOTCART2 As DataRow In tblSOTCART2.Select("CART_NO = '" & CartonNumber & "' AND ISNULL(QTY_PACKED, 0) > 0")
                Dim data As String = rowSOTCART2.Item("QTY_PACKED") & " - " & rowSOTCART2.Item("STYLE_CODE") & " / " & rowSOTCART2.Item("COLOR_CODE")
                data = data.Replace(Chr(34), "")
                labelImage &= "A50," & lineNumber & ",0,4,1,1,N," & Chr(34) & data & Chr(34) & Environment.NewLine
                lineNumber += 40
            Next
        End If

        labelImage &= "P1" & Environment.NewLine

        Return labelImage
    End Function

#Region "VB6"
    'Function bol(c As String, SO_PARM_UPC_VENDOR_ID As String) As String

    '    Return ("")
    'End Function

    'Sub Import_Shipment_Log(Optional FLB As Object = Nothing, Optional importFile As String = "")


    'End Sub

    Public Shared Function Validate_UPC(UPC_CODE As String, SO_PARM_UPC_VENDOR_ID As String) As String
        Stop
        'Return Null if OK or Error Msg if Bad.

        Dim VU As String = ""

        'Check that UPC is correct Length.
        If Len(UPC_CODE) <> 12 Then
            VU = VU & vbCrLf & "UPC Is Not 12 Digits."
        Else
            'Check if Check Digit is Correct.
            Dim UPC_calc As String = TAC.SOCMAIN1.UPC(Nothing, Mid(UPC_CODE, 7, 5), Mid(UPC_CODE, 1, 6), True)
            If Right(UPC_CODE, 1) <> Right(UPC_calc, 1) Then
                VU = VU & vbCrLf & "UPC Contains Incorrect Check Digit. Should Be " & Right(UPC_calc, 1)
            End If
        End If

        'Check if Correct Vendor ID was used.

        If Len(SO_PARM_UPC_VENDOR_ID) <> 6 Then
            VU = VU & vbCrLf & "VendorID In Parameter File Is Invalid."
        End If
        If Mid(UPC_CODE, 1, 6) <> SO_PARM_UPC_VENDOR_ID Then
            VU = VU & vbCrLf & "UPC Contains The Incorrect VendorID.  Should Be " & SO_PARM_UPC_VENDOR_ID
        End If

        Return VU
    End Function

    'Function Calc_GTIN(UPC_CODE As String, PACK_CODE As String)

    '    Return ""
    'End Function

    Public Shared Function Calc_Cost_DELETE_ME(FF As ASFBASE0, STYLE_CODE As String, COLOR_CODE As String, BUILD_ICWCOST1 As Boolean, ADD_QTY As Long, USING_TEMP_INV As Boolean, TINVH1 As String, TINVH2 As String, Optional SHOWMARKDOWN As Boolean = False)
        'IMPORTANT!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'An Exact copy of this function exists in VAN and VANX.
        'Any Changes you make to this function must be replicated in that function as well.
        'Failure to do so will result in your impending doom
        '
        'You will be returned a string that you must break open as follows:
        '   Dim w As String
        '   Dim A() As String
        '   ReDim A(1)
        '   w = Calc_Cost(STYLE_CODE, COLOR_CODE, False,0)
        '   A() = Split(w, "|")
        '   A(0) = The Average Cost
        '   A(1) = Total Cost Consumed
        '
        '   BUILD_ICWCOST1 if you wanna create the ICWCOST1 table
        '   ADD_QTY is additional shipment qty.  This is usefull for calculating for Orders of qty not shipped yet.
        'IMPORTANT!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        Dim a() As String
        ReDim a(1)

        If TINVH1 = "" Then
            TINVH1 = "SOTINVH1"
        End If

        If TINVH2 = "" Then
            TINVH2 = "SOTINVH2"
        End If

        'Calculate Live Period.
        ASCMAIN1.sql = " SELECT MAX(OPS_YYYYPP) FROM ICTCOST1" & vbCrLf _
         & " WHERE (TRAN_TYPE = 'B' OR TRAN_TYPE = 'M')" & vbCrLf _
         & " AND (STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "')" & vbCrLf
        Dim LIVE_DATE As String = ASCDATA1.GetDataValue
        If LIVE_DATE = "" Then
            LIVE_DATE = "200607"
        Else
            If SHOWMARKDOWN Then
                LIVE_DATE = "200607"
            End If
        End If

        If BUILD_ICWCOST1 Then
            'Build Empty MDB Table.
            ASCMAIN1.sql = " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
             & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST," & vbCrLf _
             & " 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
             & " FROM ICTCOST1" & vbCrLf _
             & " WHERE ROWNUM < 0" & vbCrLf
            FF.Create_TDA(FF.dst.Tables.Add, "ICTCOST1", "**")
        Else
            FF.dst.Tables("ICTCOST1").Rows.Clear()
        End If

        'Populate Table.
        ASCMAIN1.sql = " SELECT" & vbCrLf _
        & " STYLE_CODE, COLOR_CODE, TRAN_DATE, OPS_YYYYPP, TRAN_TYPE, TRAN_REF, TRAN_QTY," & vbCrLf _
        & " TRAN_COST, 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
        & " FROM (" & vbCrLf _
        & " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
        & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
        & " FROM ICTCOST1" & vbCrLf _
        & " WHERE ICTCOST1.STYLE_CODE =  '" & STYLE_CODE & "'" & vbCrLf _
        & " AND ICTCOST1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
        & " AND ICTCOST1.OPS_YYYYPP >= '" & LIVE_DATE & "'" & vbCrLf _
        & " UNION" & vbCrLf _
        & " SELECT POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE," & vbCrLf _
        & " ICTTRAN2.OPS_YYYYPP, 'R' TRAN_TYPE, POTSHIP2.PO_SHIPMENT_NO TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY," & vbCrLf _
        & " POTSHIP3.PO_COST_LANDED TRAN_COST" & vbCrLf _
        & " FROM POTSHIP2, POTSHIP3, POTORDR2, ICTTRAN2" & vbCrLf _
        & " WHERE POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
        & " AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
        & " AND POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
        & " AND POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
        & " AND POTORDR2.PO_ORDER_NO = ICTTRAN2.PO_ORDER_NO" & vbCrLf _
        & " AND POTORDR2.PO_ORDER_LNO = ICTTRAN2.PO_ORDER_LNO" & vbCrLf _
        & " AND POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
        & " AND POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
        & " AND ICTTRAN2.OPS_YYYYPP >= '" & LIVE_DATE & "'" & vbCrLf _
        & " UNION" & vbCrLf _
        & " SELECT D.STYLE_CODE, D.COLOR_CODE, H.INV_DATE TRAN_DATE," & vbCrLf _
        & " D.ORDR_YYYYPP_UPDATED  OPS_YYYYPP, 'S' TRAN_TYPE, '' TRAN_REF," & vbCrLf _
        & " SUM(D.ORDR_QTY_SHIP) TRAN_QTY, D.ORDR_UNIT_COST TRAN_COST" & vbCrLf _
        & " FROM " & TINVH1 & " H, " & TINVH2 & " D" & vbCrLf _
        & " WHERE H.INV_TYPE = D.INV_TYPE" & vbCrLf _
        & " AND H.INV_NO = D.INV_NO" & vbCrLf _
        & " AND D.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
        & " AND D.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
        & " AND H.INV_TYPE = 'I'" & vbCrLf
        If Not USING_TEMP_INV Then
            ASCMAIN1.sql &= " AND D.ORDR_YYYYPP_UPDATED >= '" & LIVE_DATE & "'" & vbCrLf
        End If
        ASCMAIN1.sql = "" _
        & " HAVING SUM(D.ORDR_QTY_SHIP) <> 0" & vbCrLf _
        & " GROUP BY D.STYLE_CODE, D.COLOR_CODE, H.INV_DATE," & vbCrLf _
        & " D.ORDR_YYYYPP_UPDATED, 'S', D.ORDR_UNIT_COST)" & vbCrLf _
        & " ORDER BY STYLE_CODE, COLOR_CODE, TRAN_TYPE, TRAN_DATE"
        FF.Fill_Records("ICTCOST1", "", True, ASCMAIN1.sql)

        'Calculate Consumed Qty for Receipts
        Dim W As Int64 = 0
        Dim SqlW = "(TRAN_TYPE = 'B' OR TRAN_TYPE = 'R' OR TRAN_TYPE = 'M')"
        For Each rowICTCOST1 As DataRow In FF.dst.Tables("ICTCOST1").Select(SqlW, "TRAN_DATE")
            W += rowICTCOST1.Item("TRAN_QTY")
            rowICTCOST1.Item("CUM_QTY") = W
        Next


        'Calculate Consumed Qty for Shipments
        W = 0
        SqlW = "TRAN_TYPE = 'S'"
        For Each rowICTCOST1 As DataRow In FF.dst.Tables("ICTCOST1").Select(SqlW, "TRAN_DATE")
            W += rowICTCOST1.Item("TRAN_QTY")
            rowICTCOST1.Item("CUM_QTY") = W
        Next

        'Mark Consumed Records.
        'Changed on 10/10/07 with Maurice to allow consumption of costs into future periods.
        'SQL = "UPDATE ICWCOST1 SET CONSUMED = 'Y' WHERE TRAN_TYPE = 'S' AND OPS_YYYYPP <= '" & CYP & "'"

        For Each rowICTCOST1 As DataRow In FF.dst.Tables("ICTCOST1").Select("TRAN_TYPE = 'S'")
            rowICTCOST1.Item("CONSUMED") = "Y"
        Next

        'Changed on 10/10/07 with Maurice to allow consumption of costs into future periods.
        'SQL = "SELECT SUM(TRAN_QTY) AS TOT_QTY_SHIP FROM ICWCOST1 WHERE TRAN_TYPE = 'S' AND OPS_YYYYPP <= '" & CYP & "'"

        Dim TOT_QTY_SHIP As Int64 = Val(FF.dst.Tables("ICTCOST1").Compute("SUM(TRAN_QTY)", "TRAN_TYPE = 'S'") & "")
        TOT_QTY_SHIP += Val(ADD_QTY)

        Dim TOT_SHIP_CONS As Int64 = 0

        SqlW = "(TRAN_TYPE = 'B' OR TRAN_TYPE = 'R' OR TRAN_TYPE = 'M')"
        For Each rowICTCOST1 As DataRow In FF.dst.Tables("ICTCOST1").Select(SqlW, "TRAN_TYPE, TRAN_DATE")
            TOT_SHIP_CONS = Val(rowICTCOST1.Item("CUM_QTY") & "")
            If Val(rowICTCOST1.Item("CUM_QTY") & "") < TOT_QTY_SHIP Then
                rowICTCOST1.Item("CONSUMED") = "Y"
            Else
                rowICTCOST1.Item("CONSUMED") = "Y"
                Exit For
            End If
            'If Not dynWK.EOF Then
            '    rowICTCOST1.Item("CONSUMED") = "Y"
            '    TOT_SHIP_CONS = Val(rowICTCOST1.Item("CUM_QTY"))
            'End If
        Next

        Dim TOT_R_COST As Decimal = 0
        SqlW = "(TRAN_TYPE = 'B' OR TRAN_TYPE = 'R' OR TRAN_TYPE = 'M') AND CONSUMED = 'Y'"
        For Each row As DataRow In FF.dst.Tables("ICTCOST1").Select(SqlW)
            Dim TRAN_QTY As Int64 = Val(row.Item("TRAN_QTY") & "")
            Dim TRAN_COST As Decimal = Val(row.Item("TRAN_COST") & "")
            TOT_R_COST += (TRAN_QTY * TRAN_COST)
        Next

        If Val(TOT_SHIP_CONS) = 0 Then
            a(0) = 0
        Else
            a(0) = Format(TOT_R_COST / TOT_SHIP_CONS, "0000000000.0000")
        End If
        a(1) = Format(TOT_R_COST, "0000000000.0000")
        Return Join(a, "|")
    End Function

    'Public Shared Function Calc_Cost_New_DELETE_ME(PERIOD As String, STYLE_CODE As String, COLOR_CODE As String, BUILD_ICWCOST1 As Boolean)

    '    Return 0
    'End Function
    'Public Shared Function Calc_Cost_Last_DELETE_ME(PERIOD As String, STYLE_CODE As String, COLOR_CODE As String, BUILD_ICWCOST1 As Boolean)

    '    Return 0
    'End Function

    'Public Shared Function Calc_Cost_OH_DELETE_ME(PERIOD As String, STYLE_CODE As String, COLOR_CODE As String, BUILD_ICWCOST1 As Boolean)

    '    Return 0
    'End Function

    'Public Shared Function Calc_All_Cost_Lots_DELETE_ME(PERIOD As String, STYLE_CODE As String, COLOR_CODE As String, BUILD_ICWCOST1 As Boolean)

    '    Return (0)
    'End Function

    Public Shared Function MAKE_COSTMF_DELETE_ME(FF As ASFBASE0, PERIOD As String, SQLGROUPS As String, ZERO_COST As Boolean, USING_TEMP_INV As Boolean, TINVH1 As String, TINVH2 As String, Optional ORDR_TYPE As String = "")

        ' This function will create a temp table of all styles in system 
        '  with their respective currently calculated costs
        ' for a list of groups handed to it by the SQL SQLGROUPS

        Dim TT As String = String.Empty
        Dim TABLE As String = String.Empty

        Dim W As String = String.Empty
        Dim a() As String
        ReDim a(1)

        Dim TBL1 As String = ""
        Dim TBL2 As String = ""

        If ORDR_TYPE = "R" Then
            TBL1 = "SOTRSRV1"
            TBL2 = "SOTRSRV2"
        Else
            TBL1 = "SOTORDR1"
            TBL2 = "SOTORDR2"
        End If

        '    sql = " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf
        '    sql = sql & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST," & vbCrLf
        '    sql = sql & " 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf
        '    sql = sql & " FROM ICTCOST1" & vbCrLf
        '    sql = sql & " WHERE ROWNUM < 0" & vbCrLf
        ASCMAIN1.sql = "Select ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
         & " ICTCOST1.TRAN_TYPE, '" & "".PadLeft(50, " ") & "' TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST, 0.00 QTY_USED" & vbCrLf _
         & " FROM ICTCOST1" & vbCrLf _
         & " WHERE ROWNUM < 0"
        FF.Create_TDA(FF.dst.Tables.Add, "ICTCOST1", "**", , False)

        ASCMAIN1.sql = "Select " & TBL2 & ".STYLE_CODE, " & TBL2 & ".COLOR_CODE, ICTSTYL1.STYLE_DESC, SUM(1.11) STYLE_COST" & vbCrLf _
            & "  from " & TBL1 & ", " & TBL2 & ", ICTSTYL1" & vbCrLf _
            & IIf(ORDR_TYPE = "R",
                " where " & TBL1 & ".RSRV_NO = " & TBL2 & ".RSRV_NO and " & TBL1 & ".RSRV_NO IN (" & SQLGROUPS & ")" & vbCrLf,
                " where " & TBL1 & ".ORDR_NO = " & TBL2 & ".ORDR_NO and ORDR_GROUP_NO IN (" & SQLGROUPS & ")" & vbCrLf) _
            & "  and " & TBL2 & ".STYLE_CODE = ICTSTYL1.STYLE_CODE" _
            & "  group by " & TBL2 & ".STYLE_CODE, " & TBL2 & ".COLOR_CODE, ICTSTYL1.STYLE_DESC"
        TT = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "UPDATE " & TT & " SET STYLE_COST = NULL"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "ALTER Table " & TT & " MODIFY STYLE_COST NUMBER(12,6)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "UPDATE " & TT & " SET STYLE_COST = 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Create Index I_" & TT & "_1 ON " & TT & " (STYLE_CODE, COLOR_CODE)"
        ASCDATA1.ExecuteSQL()

        If Not ZERO_COST Then
            ASCMAIN1.sql = "Select * from " & TT
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_CODE, COLOR_CODE")
                'w = Calc_Cost(dynORA.Fields("STYLE_CODE").Value, dynORA.Fields("COLOR_CODE").Value, False, 0, USING_TEMP_INV, TINVH1, TINVH2)
                'w = Calc_Cost_New(PERIOD, dynORA.Fields("STYLE_CODE").Value, dynORA.Fields("COLOR_CODE").Value, False)
                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")

                Stop ' NEXT LINE WAS REMMED BECAUSE IT WAS DELETE_ME'd
                'W = Calc_Cost_OH(PERIOD, STYLE_CODE, COLOR_CODE, False)

                a = Split(W, "|")
                row.Item("STYLE_COST") = Val(a(0))
            Next
        End If

        Return TT

    End Function

    'Function Avg_cost_DELETE_ME(STYLE_CODE As String, COLOR_CODE As String, OPS_YYYYPP As String, Optional USE_NONCOMPLETE As Boolean = False)

    '    Return 0
    'End Function
#End Region

#Region "TABMAINX"

    Public Shared Function Calc_Cost_OH_DELETE_ME(FF As ASFBASE0, PERIOD As String, STYLE_CODE As String, COLOR_CODE As String, BUILD_ICWCOST1 As Boolean)
        'IMPORTANT!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'An Exact copy of this function exists in VAN and VANX.
        'Any Changes you make to this function must be replicated in that function as well.
        'Failure to do so will result in your impending doom
        'IMPORTANT!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Dim LOT_REMAIN As Long
        Dim A() As String
        ReDim A(1)

        'Calculate Live Period.
        ASCMAIN1.sql = " SELECT MAX(OPS_YYYYPP) FROM ICTCOST1" & vbCrLf _
         & " WHERE (TRAN_TYPE = 'B' OR TRAN_TYPE = 'M')" & vbCrLf _
         & " AND (STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "')" & vbCrLf
        Dim LIVE_PERIOD As String = ASCDATA1.GetDataValue
        If LIVE_PERIOD = "" Then LIVE_PERIOD = "200607"

        If PERIOD = ASCMAIN1.CYP Then
            ASCMAIN1.sql = " SELECT NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND" & vbCrLf _
             & " FROM ICTSTAT2" & vbCrLf _
             & " WHERE STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
             & " AND COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
             & " AND WHSE_CODE = 'NJ'"
        Else
            ASCMAIN1.sql = " SELECT NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND" & vbCrLf _
             & " FROM ICTSTAT5" & vbCrLf _
             & " WHERE STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
             & " AND COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
             & " AND WHSE_CODE = 'NJ'" & vbCrLf _
             & " AND OPS_YYYYPP = '" & PERIOD & "'"
        End If
        Dim TOT_OH As Int64 = ASCDATA1.GetDataValue
        Dim OH_REMAINS As Int64 = TOT_OH

        If BUILD_ICWCOST1 Then
            ASCMAIN1.sql = "SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
             & " ICTCOST1.TRAN_TYPE, '" & "".PadLeft(50, " ") & "' TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST, 0.00 QTY_USED" & vbCrLf _
             & " FROM ICTCOST1" & vbCrLf _
             & " WHERE ROWNUM < 0"
            FF.Fill_Records("ICTCOST1", "", True, ASCMAIN1.sql)
        Else
            FF.dst.Tables("ICTCOST1").Rows.Clear()
        End If

        If OH_REMAINS <= 0 Then
            A(0) = 0
            A(1) = 0
            Return Join(A, "|")
        End If

        'Find The Lot In Memory That Represents The Highest Sales.
        ASCMAIN1.sql = " SELECT" & vbCrLf _
            & " STYLE_CODE, COLOR_CODE, TRAN_DATE, OPS_YYYYPP, TRAN_TYPE, TRAN_REF, TRAN_QTY," & vbCrLf _
            & " TRAN_COST, 'N' CONSUMED, 0.00 CUM_QTY" & vbCrLf _
            & " FROM (" & vbCrLf _
            & " SELECT ICTCOST1.STYLE_CODE, ICTCOST1.COLOR_CODE, ICTCOST1.TRAN_DATE, ICTCOST1.OPS_YYYYPP," & vbCrLf _
            & " ICTCOST1.TRAN_TYPE, ICTCOST1.TRAN_REF, ICTCOST1.TRAN_QTY, ICTCOST1.TRAN_COST" & vbCrLf _
            & " FROM ICTCOST1" & vbCrLf _
            & " WHERE ICTCOST1.STYLE_CODE =  '" & STYLE_CODE & "'" & vbCrLf _
            & " AND ICTCOST1.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & " AND ICTCOST1.OPS_YYYYPP >= '" & LIVE_PERIOD & "'" & vbCrLf _
            & " AND ICTCOST1.OPS_YYYYPP <= '" & PERIOD & "'" & vbCrLf _
            & " UNION" & vbCrLf _
            & " SELECT POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP2.PO_DATE_RECEIVED TRAN_DATE," & vbCrLf _
            & " ICTTRAN2.OPS_YYYYPP, 'R' TRAN_TYPE, POTSHIP1.PO_SHIPMENT_NO ||' - ' || POTSHIP1.PO_SHIP_VESSEL TRAN_REF, POTSHIP3.PO_QTY_REC TRAN_QTY," & vbCrLf _
            & " POTSHIP3.PO_COST_LANDED TRAN_COST" & vbCrLf _
            & " FROM POTSHIP1, POTSHIP2, POTSHIP3, POTORDR2, ICTTRAN2" & vbCrLf _
            & " WHERE POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
            & " AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & " AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
            & " AND POTSHIP3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & " AND POTSHIP3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & " AND POTORDR2.PO_ORDER_NO = ICTTRAN2.PO_ORDER_NO" & vbCrLf _
            & " AND POTORDR2.PO_ORDER_LNO = ICTTRAN2.PO_ORDER_LNO" & vbCrLf _
            & " AND POTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            & " AND POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            & " AND ICTTRAN2.OPS_YYYYPP >= '" & LIVE_PERIOD & "'" & vbCrLf _
            & " AND ICTTRAN2.OPS_YYYYPP <= '" & PERIOD & "')" & vbCrLf

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "TRAN_DATE DESC")
            Dim TRAN_QTY As Int64 = Val(row.Item("TRAN_QTY") & "")
            OH_REMAINS = OH_REMAINS - TRAN_QTY
            If OH_REMAINS <= 0 Then
                LOT_REMAIN = OH_REMAINS + TRAN_QTY
                Stop ' NEXT LINE WAS REMMED BECAUSE IT WAS DELETE_ME'd
                ' AddCostLot(FF.dst.Tables("ICTCOST1").NewRow, row, LOT_REMAIN)
                Exit For
            Else
                LOT_REMAIN = TRAN_QTY
                Stop ' NEXT LINE WAS REMMED BECAUSE IT WAS DELETE_ME'd
                'AddCostLot(FF.dst.Tables("ICTCOST1").NewRow, row, LOT_REMAIN)
            End If
        Next

        Dim ReturnCost As Decimal = 0
        Dim CostQty As Decimal = 0
        Dim CostTotal As Decimal = 0

        For Each rowICTCOST1 As DataRow In FF.dst.Tables("ICTCOST1").Select("")
            CostQty += Val(rowICTCOST1.Item("QTY_USED"))
            CostTotal += Val(rowICTCOST1.Item("COST_TOTAL"))
        Next

        If CostQty <= 0 Then
            ReturnCost = 0
        Else
            ReturnCost = CostTotal / CostQty
        End If
        A(0) = Val(ReturnCost)
        A(1) = Val(TOT_OH)
        Return Join(A, "|")
    End Function

    Public Shared Sub AddCostLot_DELETE_ME(rowICTCOST1 As DataRow, ROW As DataRow, LOT_REMAIN As Int64)
        For Each COLUMN_NAME As String In New String() _
            {"STYLE_CODE", "COLOR_CODE", "TRAN_DATE", "OPS_YYYYPP", "TRAN_TYPE", "TRAN_REF", "TRAN_QTY", "TRAN_COST"}
            rowICTCOST1.Item(COLUMN_NAME) = ROW.Item(COLUMN_NAME)
        Next
        rowICTCOST1.Item("QTY_USED") = LOT_REMAIN
        rowICTCOST1.Table.Rows.Add(rowICTCOST1)
    End Sub
#End Region

#Region "Grid Layout Saving"
    Public Overrides Sub SaveGridLayout(ByRef frm As ASFBASE0, ByRef grd As UltraWinGrid.UltraGrid)
        frm.Fill_Records("ASTGRID1", New String() {ASCMAIN1.USER_ID, frm.Name, grd.Name}, True)

        For Each rowASTGRID1 As DataRow In frm.dst.Tables("ASTGRID1").Rows
            rowASTGRID1.Delete()
        Next
        For Each grdCol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
            Dim COL_VISIBLE As String = "1"
            If grdCol.Hidden Then
                COL_VISIBLE = "0"
            End If
            Dim newASTGRID1 As DataRow = frm.dst.Tables("ASTGRID1").NewRow
            newASTGRID1.Item("USER_ID") = ASCMAIN1.USER_ID
            newASTGRID1.Item("FORM_NAME") = frm.Name
            newASTGRID1.Item("GRID_NAME") = grd.Name
            newASTGRID1.Item("COL_NAME") = grdCol.Key
            newASTGRID1.Item("COL_POS") = grdCol.Header.VisiblePosition
            newASTGRID1.Item("COL_WIDTH") = grdCol.Width
            newASTGRID1.Item("COL_VISIBLE") = COL_VISIBLE
            frm.dst.Tables("ASTGRID1").Rows.Add(newASTGRID1)
        Next
        frm.Update_Record_TDA("ASTGRID1")
        MsgBox("Save Complete", vbOKOnly, "Grid Layout")
    End Sub

    Public Overrides Sub loadGridLayout(ByRef frm As ASFBASE0, ByRef grd As UltraWinGrid.UltraGrid)
        frm.Fill_Records("ASTGRID1", New String() {ASCMAIN1.USER_ID, frm.Name, grd.Name}, True)
        If frm.dst.Tables("ASTGRID1").Rows.Count > 0 Then
            For Each grdCol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                Dim filter As String = "COL_NAME = '" & grdCol.Key & "'"
                Dim row As DataRow = frm.dst.Tables("ASTGRID1").Select(filter).FirstOrDefault
                If Not IsNothing(row) Then
                    grdCol.Header.VisiblePosition = Val(row.Item("COL_POS").ToString)
                    grdCol.Width = Val(row.Item("COL_WIDTH").ToString)
                    If row.Item("COL_VISIBLE").ToString <> "1" Then
                        grdCol.Hidden = True
                    End If
                End If
            Next
        End If
    End Sub
#End Region

    Public Overloads Shared Sub Record_Event(
    ByVal TABLE_NAME As String,
    ByVal TABLE_KEY As String,
    ByVal INIT_DATE As Date,
    ByVal INIT_OPER As String,
    ByVal EVENT_TYPE As String,
    ByVal EVENT_DESC As String,
    Optional ByVal EVENT_KEY As String = "",
    Optional FORM_NAME As String = "")

        If FORM_NAME = "" Then
            FORM_NAME = ASCMAIN1.ActiveForm.Name
        End If

        ASCDATA1.ExecuteSQL("Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME) " _
                             & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8)",
                             "VVDVVVVV",
                             New Object() {TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME})
    End Sub

    'Public Shared Function Get_CURR_EXCH_RATE( _
    '    rowGLTPARM1 As DataRow, _
    '    CURR_CODE As String, _
    '    TRAN_DATE As Date, _
    '    check_forex_if_missing_daily_rate As Boolean) As Decimal

    '    'Dim frmASFBASE0 As New ASFBASE0
    '    'frmASFBASE0.ROWs = New Dictionary(Of String, DataRow)
    '    'frmASFBASE0.ROWs.Add("GLTPARM1", rowGLTPARM1)

    '    'Return Get_CURR_EXCH_RATE(frmASFBASE0, CURR_CODE, TRAN_DATE, check_forex_if_missing_daily_rate)
    '    Return Get_CURR_EXCH_RATE(rowGLTPARM1, CURR_CODE, TRAN_DATE, check_forex_if_missing_daily_rate)
    'End Function

    Public Shared Function Get_CURR_EXCH_RATE(
        rowGLTPARM1 As DataRow,
        CURR_CODE As String,
        TRAN_DATE As Date,
        Optional check_forex_if_missing_daily_rate As Boolean = True) As Decimal

        Dim CURR_EXCH_RATE As Decimal = 0
        If TRAN_DATE.ToString = "1/1/0001 12:00:00 AM" Then
            TRAN_DATE = Now.Date
        End If

        If CURR_CODE = "" Or CURR_CODE = rowGLTPARM1.Item("GL_PARM_CURR_CODE") Then
            ' If CURR_CODE = "" Or CURR_CODE = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
            CURR_EXCH_RATE = 1
        Else
            ASCMAIN1.sql = "Select * from TATCURR3" & vbCrLf _
              & " where CURR_CODE = :PARM1" & vbCrLf _
              & "   and  CURR_DATE = (Select Max(CURR_DATE) from TATCURR3" & vbCrLf _
              & " where CURR_CODE = :PARM2 and CURR_DATE <= :PARM3)"
            Dim rowTATCURR3 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVD", New Object() {CURR_CODE, CURR_CODE, TRAN_DATE})
            If rowTATCURR3 IsNot Nothing Then
                Dim X As Integer = 30
                Dim CURR_DATE As Date = rowTATCURR3.Item("CURR_DATE")
                Dim DAYS_OLD As Integer = TRAN_DATE.Subtract(CURR_DATE).Days
                CURR_EXCH_RATE = Val(rowTATCURR3.Item("CURR_EXCH_RATE") & "")
                If DAYS_OLD > X Then
                    CURR_EXCH_RATE = 0
                End If
                'Else
                '    CURR_EXCH_RATE = 0
                '    If check_forex_if_missing_daily_rate Then
                '        TAC.TACMAIN1.Update_Forex()
                ''        CURR_EXCH_RATE = Get_CURR_EXCH_RATE(frmASFBASE0, CURR_CODE, TRAN_DATE, False)
                '        CURR_EXCH_RATE = Get_CURR_EXCH_RATE(rowGLTPARM1, CURR_CODE, TRAN_DATE, False)
                '    End If
            End If
        End If

        If CURR_EXCH_RATE = 0 Then
            If check_forex_if_missing_daily_rate Then
                TAC.TACMAIN1.Update_Forex()
                'CURR_EXCH_RATE = Get_CURR_EXCH_RATE(frmASFBASE0, CURR_CODE, TRAN_DATE, False)
                CURR_EXCH_RATE = Get_CURR_EXCH_RATE(rowGLTPARM1, CURR_CODE, TRAN_DATE, False)
            End If
        End If

        If CURR_EXCH_RATE = 0 Then
            Throw New Exception("Cannot determine Exchange Rate in TACMAIN1.Get_CURR_EXCH_RATE")
        End If

        Return CURR_EXCH_RATE
    End Function
    Public Shared Sub Update_Forex()

        ASCMAIN1.sql = "Select * from TATCURR1" _
            & " Where CURR_CODE <> 'USD'"

        For Each rowTATCURR1 As DataRow In ASCDATA1.GetDataTable.Rows

            Dim gotTodaysRate As Boolean = False
            Dim CURR_EXCH_RATE_response As String = ""
            Dim INIT_DATE As Date
            Dim INIT_OPER As String

            Dim CURR_CODE As String = rowTATCURR1.Item("CURR_CODE") & ""
            Dim sqlCurr As String = "SELECT TATCURRX.CURR_DATE_X, '" & CURR_CODE & "' CURR_CODE_X" _
            & ", NVL(TATCURR3.CURR_EXCH_RATE,0) CURR_EXCH_RATE_X" _
            & " FROM TATCURR3, (select TRUNC(SYSDATE - 60) + rownum - 1 CURR_DATE_X" _
            & " from all_objects" _
            & " where rownum <= to_date(SYSDATE,'dd-mon-yyyy')-to_date(TRUNC(SYSDATE -60),'dd-mon-yyyy')+1) TATCURRX" _
            & " WHERE " _
            & " TATCURR3.CURR_DATE (+) = TATCURRX.CURR_DATE_X"
            Dim tblTATCURRX As DataTable = ASCDATA1.GetDataTable(sqlCurr)

            If tblTATCURRX.Rows.Count <> 0 Then
                For Each rowTATCURRX As DataRow In tblTATCURRX.Select("", "CURR_DATE_X")
                    Dim CURR_EXCH_RATE As Decimal = 0
                    Dim CURR_EXCH_RATE_X As Decimal = Val(rowTATCURRX.Item("CURR_EXCH_RATE_X") & "")
                    Dim CURR_DATE_X As Date = rowTATCURRX.Item("CURR_DATE_X")
                    If CURR_EXCH_RATE_X = 0 Then
                        If CURR_DATE_X.Date = Now + ASCMAIN1.NowTSD Then
                            CURR_EXCH_RATE = Get_Current_Exchange_Rate(CURR_CODE)
                            If CURR_EXCH_RATE <> 0 Then
                                gotTodaysRate = True
                            End If
                        Else
                            CURR_EXCH_RATE = Get_Historical_Exchange_Rate(CURR_CODE, CURR_DATE_X)
                        End If

                        If CURR_EXCH_RATE <> 0 Then
                            INIT_OPER = ASCMAIN1.USER_ID
                            INIT_DATE = Now + ASCMAIN1.NowTSD
                            ASCMAIN1.sql = "Insert into TATCURR3 Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7)"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VDNDVDV", New Object() {CURR_CODE, CURR_DATE_X, CURR_EXCH_RATE, INIT_DATE, INIT_OPER, INIT_DATE, INIT_OPER})

                            Dim cd As String = Format(CURR_DATE_X.Date, "MM/dd/yyyy")
                        End If
                    Else
                        If CURR_DATE_X.Date = Now Then
                            gotTodaysRate = True
                        End If
                    End If
                Next
            End If
            'ASCMAIN1.sql = "Select GLTPARM2.*,TATCURR2.CURR_EXCH_CUR, TATCURR3.CURR_EXCH_RATE" & vbCrLf _

            Dim YP_END As String = ASCMAIN1.CYP
            Dim YP_BEG As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6)

            ASCMAIN1.sql = "Select GLTPARM2.OPS_YYYYPP,TATCURR3.CURR_CODE,TATCURR3.CURR_EXCH_RATE" & vbCrLf _
                & " from GLTPARM2,TATCURR3" & vbCrLf _
                & " where GLTPARM2.OPS_YYYYPP <= '" & YP_END & "'" & vbCrLf _
                & "   and GLTPARM2.OPS_YYYYPP >= '" & YP_BEG & "'" & vbCrLf _
                & "   and TATCURR3.CURR_DATE = GLTPARM2.PRD_END_DATE" & vbCrLf _
                & "   and TATCURR3.CURR_CODE = '" & CURR_CODE & "'" & vbCrLf _
                & "   and GLTPARM2.OPS_YYYYPP not in " & vbCrLf _
                & " (Select OPS_YYYYPP from TATCURR2 where CURR_CODE = '" & CURR_CODE & "')"
            ASCMAIN1.sql = "Insert into TATCURR2 " & vbCrLf & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Select GLTPARM2.*,TATCURR2.CURR_EXCH_CUR, TATCURR3.CURR_EXCH_RATE" & vbCrLf _
            '    & " from GLTPARM2,TATCURR2,TATCURR3" & vbCrLf _
            '    & " where GLTPARM2.OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'" & vbCrLf _
            '    & "   and GLTPARM2.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6) & "'" & vbCrLf _
            '    & "   and TATCURR2.OPS_YYYYPP (+) = GLTPARM2.OPS_YYYYPP" & vbCrLf _
            '    & "   and TATCURR2.CURR_CODE (+) = '" & CURR_CODE & "'" & vbCrLf _
            '    & "   and TATCURR3.CURR_DATE = GLTPARM2.PRD_END_DATE" & vbCrLf _
            '    & "   and TATCURR3.CURR_CODE = '" & CURR_CODE & "'" & vbCrLf

            'For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            '    Dim CURR_EXCH_CUR As Decimal = Val(row.Item("CURR_EXCH_CUR") & "")
            '    Dim CURR_EXCH_RATE As Decimal = Val(row.Item("CURR_EXCH_RATE") & "")
            '    If CURR_EXCH_RATE <> 0 And CURR_EXCH_CUR = 0 Then
            '        ASCMAIN1.sql = "Insert into TATCURR2 Values (:PARM1,:PARM2,:PARM3)"
            '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVN", New Object() {row.Item("OPS_YYYYPP"), CURR_CODE, CURR_EXCH_RATE})
            '    End If
            'Next

            If Not gotTodaysRate Then
                ' MessageBox.Show("Error getting daily exchange rate for: " & CURR_CODE, "Error", MessageBoxButtons.OK)
            End If

        Next

        ASCMAIN1.Progress("")

    End Sub
    Public Shared Function Get_Current_Exchange_Rate(forCode As String) As Decimal

        Dim CURR_EXCH_RATE As Decimal = 0
        Dim CURR_EXCH_RATE_response As String
        Dim forexSvc As String
        For Each RATE_API As String In New String() {"appspot", "ratelab", "openexchange"}
            forexSvc = RATE_API
            CURR_EXCH_RATE_response = ""
            CURR_EXCH_RATE_response = Rate_By_Service(forCode, Now + ASCMAIN1.NowTSD, RATE_API)
            ASCMAIN1.Progress("Curr: " & forCode & ", Date: " & Now + ASCMAIN1.NowTSD, "Rate: " & CURR_EXCH_RATE_response.ToString)
            If CURR_EXCH_RATE_response <> "" AndAlso (Val(CURR_EXCH_RATE_response) <> 0) Then
                Return Val(CURR_EXCH_RATE_response)
            End If
        Next

        Return CURR_EXCH_RATE

    End Function

    Public Shared Function Get_Historical_Exchange_Rate(forCode As String, forDate As Date) As Decimal

        Dim CURR_EXCH_RATE As Decimal = 0
        Dim CURR_EXCH_RATE_response As String = ""
        Dim forexSvc As String
        forexSvc = "openexchangeHistory"
        CURR_EXCH_RATE_response = Rate_By_Service(forCode, forDate, "openexchangeHistory")
        If CURR_EXCH_RATE_response <> "" AndAlso (Val(CURR_EXCH_RATE_response) <> 0) Then
            ASCMAIN1.Progress("Curr: " & forCode & ", Date: " & forDate.Date.ToString, CURR_EXCH_RATE_response.ToString)
            Return Val(CURR_EXCH_RATE_response)
        End If

        Return CURR_EXCH_RATE

    End Function



    Public Shared Function Rate_By_Service(forCode As String, rateDate As Date, rateService As String) As String
        Dim responseString As String = ""
        Dim client As New HttpClient()
        Dim API_BASE As String = ""
        Dim API_METHOD As String = ""
        Dim API_QUERY_STRING As String = ""

        Select Case rateService
            Case "appspot"
                'no api key needed
                ' Sample URL: http://rate-exchange.appspot.com/currency?from=CAD&to=USD
                ' {"to": "USD", "rate": 0.79647100000000004, "from": "CAD"}
                API_BASE = "http://rate-exchange.appspot.com/"
                API_METHOD = "currency"
                API_QUERY_STRING = "?from=" & forCode & "&to=USD"
            Case "ratelab"
                'base is USD so response needs to to be 1/rate
                'apiKey=27429B02DA56DD370A6A9091430DD0F1
                'Sample URL: http://api.exchangeratelab.com/api/single/CAD?apikey=27429B02DA56DD370A6A9091430DD0F1
                'Sample Response
                '{"rate":{"rate":1.2582,"to":"USD"},"baseCurrency":"CAD","timeStamp":1429023188,"executionTime":28,"licenseMessage":"Data Retrieved From www.ExchangeRateLab.com - Under license (Not for financial/professional use)"}
                API_BASE = "http://api.exchangeratelab.com/"
                API_METHOD = "api/single/" & forCode
                API_QUERY_STRING = "?apikey=27429B02DA56DD370A6A9091430DD0F1"

            Case "openexchange"
                'base is USD so response needs to to be 1/rate
                'app_id=44076a2ca9a243b3b61f08219fb7809f
                'Sample URL: https://openexchangerates.org/api/latest.json?app_id=44076a2ca9a243b3b61f08219fb7809f
                API_BASE = "https://openexchangerates.org/"
                API_METHOD = "api/latest.json"
                API_QUERY_STRING = "?app_id=44076a2ca9a243b3b61f08219fb7809f"
            Case "openexchangeHistory"
                'https://openexchangerates.org/api/historical/2015-04-21.json?app_id=44076a2ca9a243b3b61f08219fb7809f
                Dim rateYear As String = rateDate.Year.ToString
                Dim rateMonth As String = rateDate.Month.ToString("00")
                Dim rateDay As String = rateDate.Day.ToString("00")
                Dim rd As String = rateYear & "-" & rateMonth & "-" & rateDay
                API_BASE = "https://openexchangerates.org/"
                API_METHOD = "api/historical/" & rd & ".json"
                API_QUERY_STRING = "?app_id=44076a2ca9a243b3b61f08219fb7809f"
        End Select

        Try
            client.BaseAddress = New Uri(API_BASE)
            Dim API As String = API_METHOD & API_QUERY_STRING
            Dim response As HttpResponseMessage = client.GetAsync(API).Result
            ASCMAIN1.Progress("Fx Svc:" & API, response.StatusCode.ToString)
            If response.IsSuccessStatusCode Then
                Dim json As String = response.Content.ReadAsStringAsync().Result
                Select Case rateService
                    Case "appspot"
                        Return Newtonsoft.Json.Linq.JObject.Parse(json).SelectToken("rate").ToString
                    Case "ratelab"
                        Dim originalValue As Double = Val(Newtonsoft.Json.Linq.JObject.Parse(json).SelectToken("rate.rate").ToString & "")
                        If originalValue <> 0 Then
                            Dim reciprocalValue As Double = 1 / originalValue
                            Return reciprocalValue.ToString
                        End If
                    Case "openexchange", "openexchangeHistory"
                        Dim originalValue As Double = Val(Newtonsoft.Json.Linq.JObject.Parse(json).SelectToken("rates.CAD").ToString & "")
                        If originalValue <> 0 Then
                            Dim reciprocalValue As Double = 1 / originalValue
                            Return reciprocalValue.ToString
                        End If
                End Select
            Else
                Return responseString
            End If
        Catch ex As Exception
            Return responseString
        End Try

        Return responseString
    End Function

    Public Shared Sub Get_Unprocessed_IDOCs(frmASFBASE0 As ASFBASE0)

        Dim sftp_folder As String = "" _
            & IIf(ASCMAIN1.Running_in_VS And 1 = 1, "C:\Users\wjz\Desktop\Interparfums\", ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT")) _
            & "\IPSA\" _
            & IIf(ASCMAIN1.DBS_SERVER = "TST" Or ASCMAIN1.DBS_COMPANY = "TST", "TEST", "PROD") _
            & "\FROM_IPSA\IDOC\"

        If Not frmASFBASE0.dst.Tables.Contains("TATIDOCU") Then
            With frmASFBASE0.dst.Tables.Add("TATIDOCU")
                .Columns.Add("FILENAME")
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .Columns.Add("FILEDATE", GetType(System.DateTime))
                .Columns.Add("FILENAME_SHORT")

                .Columns.Add("INV_NUM")
                .Columns.Add("INV_DATE", GetType(System.DateTime))
                .Columns.Add("INV_AMT", GetType(System.Decimal))
                .Columns.Add("PO_ORDER_NO")
                .Columns.Add("PINV_TYPE")
                .Columns.Add("PINV_REF_INV")
            End With
        End If

        frmASFBASE0.dst.Tables("TATIDOCU").Rows.Clear()
        For Each FILENAME As String In My.Computer.FileSystem.GetFiles(sftp_folder)
            Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
            Dim rowTATIDOCU As DataRow = frmASFBASE0.dst.Tables("TATIDOCU").Rows.Add(New Object() {FILENAME, FI.Length, FI.CreationTime, FI.Name})

            Using sr As New System.IO.StreamReader(FILENAME)
                Dim T As String = sr.ReadToEnd
                Dim Ts() As String = Split(T, vbCrLf)
                Dim INV_NUM As String = ""
                Dim INV_DATE As Date = Nothing
                Dim INV_AMT As Decimal = 0
                Dim PO_ORDER_NO As String = ""
                Dim PINV_TYPE As String = ""
                Dim PINV_REF_INV As String = ""

                For i As Integer = 0 To Ts.Length - 1
                    Dim tx As String = Ts(i)
                    If tx.StartsWith("E2EDK02") Then
                        If Mid(tx, 64, 3) = "009" Then '
                            'E2EDK02                       5000000000001212899000008000000020090120069669                               20150915      
                            INV_NUM = Trim(Mid(tx, 67, 40))
                            Dim txd As String = Mid(tx, 108, 8)
                            INV_DATE = CDate(Mid(txd, 5, 2) & "/" & Mid(txd, 7, 2) & "/" & Mid(txd, 1, 4))
                        End If
                        If Mid(tx, 64, 3) = "087" Then
                            'E2EDK02                       500000000000121289900000900000002001ILLICIT MOSAIC BOARD                     20150914      
                            'E2EDK02                       500000000000107802900001300000002087132293                                          

                            PO_ORDER_NO = Trim(Mid(tx, 67, 40))
                        End If

                        If Mid(tx, 64, 3) = "017" Then
                            'E2EDK02                       5000000000001078029000012000000020170120058620                                             
                            PINV_REF_INV = Trim(Mid(tx, 67, 40))
                        End If
                    End If

                    If tx.StartsWith("E2EDK05001") Then
                        'E2EDK05001                    500000000000121289900002100000002+                                                                                      128.80                    0                                                                    USD
                        INV_AMT = Val(Trim(Mid(tx, 141, 16)))
                    End If
                    If tx.StartsWith("E2EDK01005") Then
                        'E2EDK01005                    500000000000107802900000100000001    USDEUR0.80199     Z090                                 FR39350219382       INVO0120058620                         7076.854          7076.854          KGMLR                                                     0000200656                                                                    L                                                      1.24690     
                        PINV_TYPE = Mid(tx, 143, 4)
                        PINV_TYPE = Mid(PINV_TYPE, 1, 1)
                    End If
                Next

                If PINV_TYPE = "I" Then
                    If PINV_REF_INV <> INV_NUM Then
                        PINV_TYPE = "D"
                    End If
                End If

                rowTATIDOCU.Item("INV_NUM") = INV_NUM
                rowTATIDOCU.Item("INV_DATE") = INV_DATE
                rowTATIDOCU.Item("INV_AMT") = INV_AMT
                rowTATIDOCU.Item("PO_ORDER_NO") = PO_ORDER_NO
                rowTATIDOCU.Item("PINV_TYPE") = PINV_TYPE
                rowTATIDOCU.Item("PINV_REF_INV") = PINV_REF_INV

                sr.Close()
                sr.Dispose()
            End Using


            'Dim rowE2EDK02_009 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '009'")(0)
            'Dim rowE2EDK02_001 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '001'")(0)
            'Dim rowE2EDS01 As DataRow = dst.Tables("E2EDS01").Select("SUMID = '010'")(0)

            'IDOC_DATA_KEY = Trim(rowE2EDK02_009.Item("BELNR"))
            'Dim YYYYMMDD As String = Trim(rowE2EDK02_009.Item("DATUM"))
            'IDOC_DATA_DATE = CDate(Mid(YYYYMMDD, 5, 2) & "/" & Mid(YYYYMMDD, 7, 2) & "/" & Mid(YYYYMMDD, 1, 4))
            'IDOC_DATA_AMT = Val(rowE2EDS01.Item("SUMME") & "")
            'IDOC_DATA_REF = Trim(rowE2EDK02_001.Item("BELNR"))


            'Absx1.txtFor("IDOC_DATA_KEY").Text = IDOC_DATA_KEY
            'Absx1.dteFor("IDOC_DATA_DATE").Value = IDOC_DATA_DATE
            'Absx1.numFor("IDOC_DATA_AMT").Value = IDOC_DATA_AMT
            'Absx1.txtFor("IDOC_DATA_REF").Text = IDOC_DATA_REF

        Next
    End Sub

    Public Shared Function Rename_Image_Files(frm As ASFBASE0) As Integer
        Dim FOLDER_NAME As String = ""
        Dim FILE_NAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Image File in the Folder containing the files to rename"
            openFileDialog1.Filter = "jpg files (*.jpg)|*.jpg"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILE_NAME = openFileDialog1.FileName
            End If
        End Using

        If FILE_NAME = "" Then Return 0

        Dim F As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE_NAME)
        FOLDER_NAME = F.DirectoryName

        Dim T As New DataTable
        With T.Columns
            .Add("FILE_NAME")
        End With

        Dim I As Int32 = 0
        Dim N As Int32 = 0
        Dim A As Int32 = 0
        Dim W As Int32 = 0

        ASCMAIN1.Progress("Now Renaming Image Files", "")
        frm.Cursor = Cursors.WaitCursor

        For Each FN As String In My.Computer.FileSystem.GetFiles(FOLDER_NAME)
            Dim STYLE_CODE As String = ""
            Dim COLOR_CODE As String = ""
            Dim file_ok As Boolean = False
            Dim file_ALREADY_ok As Boolean = False


            Dim rowICTSTYC1 As DataRow = Nothing

            If Not FN.ToUpper.EndsWith(".JPG") Then
                W += 1
            Else
                Dim FIL As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FN)

                Dim FNSC As String = FIL.Name

                FNSC = Mid(FNSC, 1, Len(FNSC) - 4)
                ASCMAIN1.Progress("-", FNSC & "; " & CStr(I) & "/" & CStr(N))

                If InStr(FNSC, "-") <> 0 Then
                    STYLE_CODE = Split(FNSC, "-")(0)
                    COLOR_CODE = Split(FNSC, "-")(1)
                    rowICTSTYC1 = frm.LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                End If

                If rowICTSTYC1 IsNot Nothing Then
                    A += 1
                    file_ALREADY_ok = True
                Else
                    For P As Integer = 4 To 2 Step -1
                        STYLE_CODE = Mid(FNSC, 1, Len(FNSC) - P)
                        COLOR_CODE = Mid(FNSC, Len(FNSC) - P + 1)

                        If COLOR_CODE.Contains("-") Then
                            Dim rowICTCOLR1_X As DataRow = frm.LookUp("ICTCOLR1_X", COLOR_CODE)
                            If rowICTCOLR1_X Is Nothing Then rowICTCOLR1_X = frm.LookUp("ICTCOLR1_X", Replace(COLOR_CODE, "-", "/"))
                            If rowICTCOLR1_X IsNot Nothing Then
                                COLOR_CODE = rowICTCOLR1_X.Item("COLOR_CODE_NEW")
                            End If
                        End If

                        rowICTSTYC1 = frm.LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                        If rowICTSTYC1 IsNot Nothing Then Exit For
                    Next
                    If rowICTSTYC1 IsNot Nothing Then
                        Dim NFN As String = FOLDER_NAME & "\" & STYLE_CODE & "-" & COLOR_CODE & ".jpg"
                        If Not My.Computer.FileSystem.FileExists(NFN) Then
                            Try
                                My.Computer.FileSystem.RenameFile(FN, STYLE_CODE & "-" & COLOR_CODE & ".jpg")
                                file_ok = True
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                End If
            End If


            If file_ok Then
                I += 1
            Else
                If Not file_ALREADY_ok Then
                    N += 1
                    T.Rows.Add(FN)
                End If
            End If
        Next

        frm.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

        MsgBox("Files Successfully Renamed: " & CStr(I) _
               & vbCrLf & "Files already looking good: " & CStr(A) _
               & vbCrLf & "Unidentified Files (not .jpg): " & CStr(W) _
               & vbCrLf & "Files NOT Successfully Renamed: " & CStr(N), MsgBoxStyle.OkOnly, "Results")

        If N <> 0 Then
            Using frmmsg As New ASFMSGBF
                frmmsg.Show_grd(T, frm, "Files which were NOT Successfully Renamed")
            End Using
        End If

        Return I

    End Function

    Public Shared Function FindZebraPrinter() As String
        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA450") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZP550") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZP 550") Then
                Return printerName
            End If
        Next printerName

        Return ""
    End Function

    Public Shared Function Check_Division_MixMatch(frmASFBASE0 As ASFBASE0, ByRef EMsg As String, TABLE_NAME As String, CUST_CODE As String, WHSE_CODE As String) As String
        'ALTER TABLE ARTCUST1 ADD SEG4_CODE                      VARCHAR2(6);
        'ALTER TABLE ICTWHSE1 ADD SEG4_CODE                      VARCHAR2(6);

        Dim SALES_DIVISION_CODEs As New List(Of String)
        Dim SEG4_CODEs As New List(Of String)
        For Each row As DataRow In frmASFBASE0.dst.Tables(TABLE_NAME).Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE") & ""
            If STYLE_CODE <> "" Then ' range styles do not have values in style code
                Dim rowICTSTYL1 As DataRow = frmASFBASE0.LookUp("ICTSTYL1", STYLE_CODE)
                Dim SALES_DIVISION_CODE As String = rowICTSTYL1.Item("SALES_DIVISION_CODE") & ""
                If SALES_DIVISION_CODE = "" Then
                    EMsg &= vbCr & "No Sales Divison Code set up for Style " & STYLE_CODE
                Else
                    If Not SALES_DIVISION_CODEs.Contains(SALES_DIVISION_CODE) Then
                        Dim rowSOTSDIV1 As DataRow = frmASFBASE0.LookUp("SOTSDIV1", SALES_DIVISION_CODE)
                        Dim SEG4_CODE As String = rowSOTSDIV1.Item("SEG4_CODE") & ""
                        If Not SEG4_CODEs.Contains(SEG4_CODE) Then
                            If SEG4_CODEs.Count > 0 Then
                                EMsg &= vbCr & "Cannot have Styles with Sales Divisions from multiple Companies - " & STYLE_CODE & " (" & SALES_DIVISION_CODE & ")"
                                Exit For
                            End If
                            SEG4_CODEs.Add(SEG4_CODE)
                        End If
                        SALES_DIVISION_CODEs.Add(SALES_DIVISION_CODE)
                    End If
                End If
            End If
        Next

        Dim COMPANY As String = "?"
        If SEG4_CODEs.Count = 1 Then
            COMPANY = SEG4_CODEs(0) ' = "001"

            If CUST_CODE <> "" Then
                Dim rowARTCUST1 As DataRow = frmASFBASE0.LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1.Item("SEG4_CODE") & "" <> COMPANY Then
                    EMsg &= vbCr & "Cannot use Customer " & CUST_CODE & " with Styles from Company " & IIf(COMPANY = "", "NYAG-US", COMPANY)
                End If
            End If

            If WHSE_CODE <> "" Then
                Dim rowICTWHSE1 As DataRow = frmASFBASE0.LookUp("ICTWHSE1", WHSE_CODE)
                If rowICTWHSE1.Item("SEG4_CODE") & "" <> COMPANY Then
                    EMsg &= vbCr & "Cannot use Warehouse " & WHSE_CODE & " with Styles from Company " & IIf(COMPANY = "", "NYAG-US", COMPANY)
                End If
            End If

        End If

        Return COMPANY
    End Function

    Public Shared Function getEcomInfo(ByRef frmASFBASE0 As ASFBASE0, ByVal STYLE_CODE As String, Optional ByVal COLOR_CODE As String = "") As String
        Dim RetVal As String = ""
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        Dim MSG As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("E1.ECOM_NAME")
        sql.AppendLine("FROM ECTESTY1 Y1, ECTECOM1 E1")
        sql.AppendLine("WHERE Y1.ECOM_CODE = E1.ECOM_CODE")
        sql.AppendLine(String.Format("AND Y1.STYLE_CODE = '{0}'", STYLE_CODE))
        sql.AppendLine("AND (NVL(Y1.SHIP_ECOM,'0') = '1' OR NVL(Y1.SHIP_DROP,'0') = '1')")
        sql.AppendLine("ORDER BY E1.ECOM_NAME")
        Dim tblEC As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        If tblEC.Rows.Count > 0 Then
            MSG.AppendLine("This Style Is Being Offered Through:")
            For Each rowEC As DataRow In tblEC.Rows
                Dim ECOM_NAME As String = rowEC.Item("ECOM_NAME").ToString & String.Empty
                If ECOM_NAME.Length > 0 Then
                    MSG.AppendLine("<br/>")
                    MSG.AppendLine(String.Format("&#8226; {0}", ECOM_NAME))
                End If
            Next
        End If
        If MSG.Length > 0 Then
            RetVal = MSG.ToString
        End If
        Return RetVal
    End Function

    Public Shared Sub PARSE_IMAGE(ByRef IMAGE_NAME As String, ByRef STYLE_CODE As String, ByRef COLOR_CODE As String, ByRef IMAGE_SUFFIX As String)
        Dim EXT As String = ".JPG"

        If IMAGE_NAME.Length > 4 Then
            If IMAGE_NAME.EndsWith(EXT) Then
                Dim endP As Int64 = IMAGE_NAME.ToUpper.IndexOf(EXT)
                Dim begP As Int64 = IMAGE_NAME.LastIndexOf("\") + 1
                IMAGE_NAME = IMAGE_NAME.Substring(begP, endP - begP)
                Dim FULL_STYLE As String = IMAGE_NAME
                If FULL_STYLE.Length > 1 Then
                    If FULL_STYLE.IndexOf("-") > 0 Then
                        STYLE_CODE = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-"))
                    End If
                    FULL_STYLE = FULL_STYLE.Replace(STYLE_CODE + "-", "")
                End If
                If FULL_STYLE.Length > 1 Then
                    If FULL_STYLE.IndexOf("-") > 0 Then
                        COLOR_CODE = FULL_STYLE.Substring(0, FULL_STYLE.IndexOf("-"))
                        FULL_STYLE = FULL_STYLE.Replace(COLOR_CODE + "-", "")
                    Else
                        COLOR_CODE = FULL_STYLE
                        FULL_STYLE = ""
                    End If
                End If
                If FULL_STYLE.Length > 1 Then
                    IMAGE_SUFFIX = FULL_STYLE
                End If
            End If
        End If
    End Sub
End Class

