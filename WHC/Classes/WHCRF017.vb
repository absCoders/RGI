Public Class WHCRF017
    ' Application Carton Pack no LPN

    Inherits WHCRF000

    Dim PICK_NO As String
    Dim PICK_NO_CONS As String
    Dim PICK_LNO As String
    Dim SHIP_BOL_NO As String
    Dim BILL_OF_LADING_NO As String
    Dim CART_NO As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim COLOR_DESC As String
    Dim UPC_CODE As String
    Dim CUST_CODE As String
    Dim LOCATION_CODE As String
    Dim GOTO_LOCATION As String
    Dim COLOR_CODEs As New List(Of String)
    Dim TICKET_NO As String
    Dim UNITS_MOVED As Integer
    Dim UNITS_MOVED_LST As Integer
    Dim SO_PARM_UPC_VENDOR_ID As String
    Dim Printer As String
    Dim VoidList() As String
    Dim VoidTranNo As String
    Dim CartonType As String
    Dim PICK_QTY As Integer
    Dim PACK_QTY_OPEN As Integer
    Dim INNER_PACK_QTY As Integer
    Dim CARTON_PACK_QTY As Integer
    Dim CARTONS_PER_UNIT As Integer
    Dim holdScan As String
    Dim _sql As String
    Dim PickTypeSql As String
    Dim rowWHTSCANS As DataRow
    Dim PageNo As Integer = 0
    Dim Bulk As Boolean = False

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF017"

        AppStates.Add("SCAN_PTCKT", "Scan Pick Ticket |EXIT|") 'Green
        AppStates.Add("SCAN_PRTR", "Scan Printer to Send Label")
        AppStates.Add("NEW_CARTON", "Carton Type|Inners|Units|FULL|DONE|")
        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style|MODE|Remove|Show Ctn|Cls Ctn|") 'Blue
        AppStates.Add("SCAN_COLOR", "Enter Color code|CANCEL|")
        AppStates.Add("SCAN_UNITS", "Units packed|CANCEL|")
        AppStates.Add("VERIFY", "Update|Y|N|CLEAR|EXIT|")
        AppStates.Add("LEAVE", "Units in Gun, Gun must be Empty for Pick|EXIT|")
        AppStates.Add("SCAN_VOID", "Void Line or 'ALL'|BACK|PAGE|NEW PICK|EXIT|")
        AppStates.Add("PRINT_UPC", "Print UPC Labels|LARGE|SMALL|SKIP|")

        AppState = "SCAN_PTCKT"
        'ScanLocMsg = AppStates("SCAN_LOC")
        LAST_CLR = "GREEN"

        With dst
            With .Tables.Add("WHTSCANS")
                .Columns.Add("PICK_NO")
                .Columns.Add("PICK_LNO")
                .Columns.Add("ORDR_NO")
                .Columns.Add("ORDR_LNO")
                .Columns.Add("CART_NO")
                .Columns.Add("CART_LNO")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("COLOR_DESC")
                .Columns.Add("UPC_CODE")
                .Columns.Add("RELEASE_QTY", GetType(System.Int32))
                .Columns.Add("PICK_QTY", GetType(System.Int32))
                .Columns.Add("PACK_QTY", GetType(System.Int32))
                .Columns.Add("INNER_PACK_QTY", GetType(System.Int32))
                .Columns.Add("CARTON_PACK_QTY", GetType(System.Int32))
                .Columns.Add("CARTONS_PER_UNIT", GetType(System.Int32))
                .Columns.Add("STYLE_UOM")
                .PrimaryKey = New DataColumn() { .Columns("PICK_NO"), .Columns("PICK_LNO")}
            End With

            Create_TDA(.Tables.Add, "SOTPICK5", "*")
            Create_TDA(.Tables.Add, "WHTCART1", "*")
            Create_TDA(.Tables.Add, "WHTCART2", "*")
            Create_TDA(.Tables.Add, "WHTLPRT1", "*")

            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 WHERE SOTPICK1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, True, "V", 1, "PACK_STATUS,CART_SEQ_CTR,PACK_PACKER")

        End With

        SO_PARM_UPC_VENDOR_ID = ASCDATA1.GetDataValue("select SO_PARM_UPC_VENDOR_ID from SOTPARM1 where SO_PARM_KEY = 'Z'")

        Dim styles As String = TACMAIN1.LookupLocation(Me, g.GUN_LOC)
        If styles.Length > 0 Then
            AppState = "LEAVE"
        End If

        tbl = dst.Tables("WHTSCANS")
        Fill_Records("WHTLPRT1", "", True, "Select * from WHTLPRT1")

    End Sub

    Public Overrides Function Hello() As String

        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        RESPONSE &= vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(ByVal SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        If SCANTEXT = "EXIT" Then
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "R", "EXIT")
        ElseIf SCANTEXT = "OK" Then
            If AppState = "SCAN_UPC" Then
                UPCCreateResponse()
            Else
                CreateResponse("", "B", PickMessage())
            End If
        Else
            Select Case AppState
                Case "SCAN_PTCKT"
                    Dim styles As String = TACMAIN1.LookupLocation(Me, G.GUN_LOC)
                    If styles.Length > 0 Then
                        CreateResponse("LEAVE", "R", "Gun has Units, Deposit merchandise to continue")
                        Exit Select
                    End If
                    If SCANTEXT.Length < 10 Then
                        SCANTEXT = SCANTEXT.PadLeft(10, "0")
                    End If
                    PICK_NO = SCANTEXT
                    PICK_NO_CONS = ""
                    Bulk = False

                    ASCMAIN1.Multi_Task_Cleanup()

                    Dim ORDR_NO As String = ""
                    Fill_Records("SOTPICK1", PICK_NO)
                    Dim rows() As DataRow = dst.Tables("SOTPICK1").Select("")
                    If rows.Length > 0 Then
                        For Each row As DataRow In rows
                            If Not row.Item("PICK_NO_CONS") Is Null Then
                                PICK_NO_CONS = row.Item("PICK_NO_CONS")
                            End If
                            If row.Item("PICK_STATUS") <> "P" Or (Not row.Item("PACK_STATUS").Equals(Null) AndAlso row.Item("PACK_STATUS") <> "P") Then
                                CreateResponse("", "R", "Pick Ticket not available, ask for Help")
                                Exit Select
                            Else
                                SHIP_BOL_NO = row.Item("SHIP_BOL_NO")
                                ORDR_NO = row.Item("ORDR_NO")
                            End If
                        Next
                    Else
                        CreateResponse("", "R", "Invalid Pick Ticket, Re-Scan")
                        Exit Select
                    End If
                    ' We lock scantext because of splits else we would use pick_no
                    If Not ASCMAIN1.Logical_Open("SOTPICK1", SCANTEXT) Then
                        CreateResponse("", "R", "Pick Ticket is Locked.")
                        Exit Select
                    End If

                    BILL_OF_LADING_NO = ASCDATA1.GetDataValue("select BILL_OF_LADING_NO from SOTSHIP1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                    CUST_CODE = ASCDATA1.GetDataValue("select CUST_CODE from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'")

                    ASCMAIN1.sql = "Select sotpick2.PICK_NO, sotpick2.PICK_LNO, sotordr2.ordr_no, sotordr2.ordr_lno, sotordr2.STYLE_CODE, sotordr2.COLOR_CODE, ictcolr1.COLOR_DESC, ictstyc1.UPC_CODE, " & vbCrLf _
                        & " sum(sotpick2.PICK_QTY) RELEASE_QTY, max(ictstyl1.INNER_PACK_QTY) INNER_PACK_QTY, " & vbCrLf _
                        & " max(ictstyl1.CARTON_PACK_QTY) CARTON_PACK_QTY, max(ictstyl1.CARTONS_PER_UNIT) CARTONS_PER_UNIT, max(ictstyl1.STYLE_UOM)  STYLE_UOM" & vbCrLf _
                        & " from sotpick1  " & vbCrLf _
                        & " join sotpick2 on sotpick1.PICK_NO = sotpick2.PICK_NO" & vbCrLf _
                        & " Join sotordr2 on sotordr2.ordr_no = sotpick2.ordr_no And sotordr2.ordr_lno = sotpick2.ORDR_LNO " & vbCrLf _
                        & " Join ictstyl1 on ictstyl1.style_code = sotordr2.style_code " & vbCrLf _
                        & " Join ictstyc1 on ictstyc1.style_code = sotordr2.style_code And ictstyc1.color_code = sotordr2.color_code " & vbCrLf _
                        & " Join ictcolr1 on ictcolr1.color_code = sotordr2.color_code " & vbCrLf _
                        & " where sotpick2.pick_qty > 0   And  SOTPICK1.PICK_NO = '" & PICK_NO & "'" & vbCrLf _
                        & " group by sotpick2.PICK_NO, sotpick2.PICK_LNO, sotordr2.ordr_no, sotordr2.ordr_lno, sotordr2.STYLE_CODE, sotordr2.COLOR_CODE, ictcolr1.COLOR_DESC, ictstyc1.UPC_CODE"

                    rows = ASCDATA1.GetDataTable.Select("")
                    If rows.Length = 0 Then
                        CreateResponse("", "GREEN", "Nothing left to Pack, scan new Pick Ticket")
                        Exit Select
                    End If

                    dst.Tables("WHTSCANS").Rows.Clear()
                    dst.Tables("WHTCART2").Rows.Clear()

                    Fill_Records("SOTPICK5", "", True, "Select * from SOTPICK5 where PICK_NO = '" & If(PICK_NO_CONS = "", PICK_NO, PICK_NO_CONS) & "' AND PICK_STATUS = 'P'")
                    Fill_Records("WHTCART1", "", True, "Select * from WHTCART1 where PICK_NO = '" & PICK_NO & "' ")

                    For Each rowWHTCART1 As DataRow In dst.Tables("WHTCART1").Select("")
                        Fill_Records("WHTCART2", "", False, "Select * from WHTCART2 where CART_NO = '" & rowWHTCART1.Item("CART_NO") & "' ")
                    Next

                    For Each row As DataRow In rows
                        rowWHTSCANS = dst.Tables("WHTSCANS").NewRow
                        With rowWHTSCANS
                            .Item("PICK_NO") = PICK_NO
                            .Item("PICK_LNO") = row.Item("PICK_LNO")
                            .Item("ORDR_NO") = row.Item("ORDR_NO")
                            .Item("ORDR_LNO") = row.Item("ORDR_LNO")
                            .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                            .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                            .Item("COLOR_DESC") = row.Item("COLOR_DESC")
                            .Item("UPC_CODE") = row.Item("UPC_CODE")
                            .Item("RELEASE_QTY") = row.Item("RELEASE_QTY")
                            .Item("PICK_QTY") = 0
                            .Item("PACK_QTY") = 0
                            .Item("INNER_PACK_QTY") = Val(row.Item("INNER_PACK_QTY") & "")
                            .Item("CARTON_PACK_QTY") = Val(row.Item("CARTON_PACK_QTY") & "")
                            .Item("CARTONS_PER_UNIT") = Val(row.Item("CARTONS_PER_UNIT") & "")
                            .Item("STYLE_UOM") = row.Item("STYLE_UOM")
                        End With
                        dst.Tables("WHTSCANS").Rows.Add(rowWHTSCANS)
                    Next

                    Dim rowSOTPICK5 As DataRow
                    ' this code wa rewritten to allow spreading pick_quantities and pack quantities across multiple order lines
                    'For Each rowWHTSCANS In dst.Tables("WHTSCANS").Select("")
                    '    'rowWHTSCANS.Item("PACK_QTY") = Val(rowWHTSCANS.Item("PACK_QTY") & "") + 0
                    '    rowWHTSCANS.Item("PACK_QTY") = Val(rowWHTSCANS.Item("PACK_QTY") & "") + Val(dst.Tables("WHTCART2").Compute("SUM(QTY_PACKED)", "UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'") & "")
                    '    For Each rowSOTPICK5 In dst.Tables("SOTPICK5").Select("UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'")
                    '        rowWHTSCANS.Item("PICK_QTY") = Val(rowWHTSCANS.Item("PICK_QTY") & "") + Val(rowSOTPICK5.Item("PICK_QTY") & "")
                    '    Next
                    'Next
                    For Each rowSOTPICK5 In dst.Tables("SOTPICK5").Select("", "UPC_CODE")
                        Dim pick_qty As Int32 = Val(rowSOTPICK5.Item("PICK_QTY") & "")
                        For Each rowWHTSCANS In dst.Tables("WHTSCANS").Select("UPC_CODE = '" & rowSOTPICK5.Item("UPC_CODE") & "'")
                            If Val(rowWHTSCANS.Item("PICK_QTY") & "") < Val(rowWHTSCANS.Item("RELEASE_QTY") & "") Then
                                If Val(rowWHTSCANS.Item("PICK_QTY") & "") + pick_qty > Val(rowWHTSCANS.Item("RELEASE_QTY") & "") Then
                                    pick_qty = pick_qty - (Val(rowWHTSCANS.Item("PICK_QTY") & "") - Val(rowWHTSCANS.Item("RELEASE_QTY") & ""))
                                    rowWHTSCANS.Item("PICK_QTY") = Val(rowWHTSCANS.Item("RELEASE_QTY") & "")
                                Else
                                    rowWHTSCANS.Item("PICK_QTY") = Val(rowWHTSCANS.Item("PICK_QTY") & "") + pick_qty
                                    pick_qty = 0
                                End If
                            End If
                        Next
                        'pick qty shouldn't be anything else than zero Will need to do something here, how to tell the packer there is a problem
                    Next
                    Dim rowWHTCART2 As DataRow
                    For Each rowWHTCART2 In dst.Tables("WHTCART2").Select("", "ORDR_NO,ORDR_LNO")
                        For Each rowWHTSCANS In dst.Tables("WHTSCANS").Select("ORDR_NO='" & rowWHTCART2.Item("ORDR_NO") & "' and ORDR_LNO = '" & rowWHTCART2.Item("ORDR_LNO") & "'")
                            rowWHTSCANS.Item("PACK_QTY") = Val(rowWHTSCANS.Item("PACK_QTY") & "") + Val(rowWHTCART2.Item("QTY_PACKED") & "")
                        Next
                    Next

                    ASCMAIN1.sql = "select * from WHTCART2 " & vbCrLf _
                        & " where (CART_NO, CART_LNO) = " & vbCrLf _
                        & " (select WHTCART2.CART_NO, max(CART_LNO) " & vbCrLf _
                        & " from WHTCART1, WHTCART2" & vbCrLf _
                        & " where WHTCART1.PICK_NO = '" & PICK_NO & "'" & vbCrLf _
                        & " And WHTCART1.PROCESS_STATUS = '0'" & vbCrLf _
                        & " And WHTCART2.CART_NO = WHTCART1.CART_NO" & vbCrLf _
                        & "  group by WHTCART2.CART_NO)"
                    rowWHTCART2 = ASCDATA1.GetDataTable.Select("").FirstOrDefault
                    If Not rowWHTCART2 Is Nothing Then
                        CART_NO = rowWHTCART2.Item("CART_NO")
                        Dim msg As String = "Found Open Carton" & vbCrLf
                        msg = msg & "last item loaded:" & vbCrLf _
                            & rowWHTCART2.Item("STYLE_CODE") & " -" & rowWHTCART2.Item("COLOR_CODE") & vbCrLf _
                            & "Qty: " & rowWHTCART2.Item("QTY_PACKED")
                        Fill_Records("WHTCART2", "", True, "Select * from WHTCART2 where CART_NO = '" & CART_NO & "' ")
                        Dim rowWHTCART1 As DataRow = dst.Tables("WHTCART1").Select("CART_NO = '" & CART_NO & "' ").First
                        Printer = rowWHTCART1.Item("PRINTER")
                        CartonType = "Units"
                        UPCCreateResponse("R", msg)
                    Else
                        CreateResponse("SCAN_PRTR", "YELLOW", PickMessage())
                    End If

                Case "SCAN_PRTR"
                    If SCANTEXT = "DONE" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_PTCKT", "GREEN", "Scan New Ticket")
                        Exit Select
                    Else
                        Dim row As DataRow = dst.Tables("WHTLPRT1").Select("LABEL_PRINTER_ID = '" & SCANTEXT & "'").FirstOrDefault
                        If row IsNot Nothing Then
                            Printer = SCANTEXT
                        Else
                            CreateResponse("", "R", "Unknown Printer")
                            Exit Select
                        End If
                    End If
                    CreateResponse("NEW_CARTON", "R", PickMessage())

                Case "NEW_CARTON"
                    If SCANTEXT = "DONE" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_PTCKT", "GREEN", PickMessage())
                        Exit Select
                    ElseIf SCANTEXT.ToUpper = "BULK" Then
                        CART_NO = ""
                        CartonType = SCANTEXT
                        CreateBulkCarton()
                        CreateResponse("VERIFY", "R", "Unknown Printer")
                    Else
                        'create new carton at update
                        'Prepare Scanstate message for UPC
                        CART_NO = ""
                        CartonType = SCANTEXT
                        UPCCreateResponse()
                    End If

                Case "SCAN_UPC"
                    Dim hold As String
                    'Ent Qty|Cls Ctn
                    If SCANTEXT = "Inners" Or SCANTEXT = "Units" Or SCANTEXT = "Remove" Then
                        CartonType = SCANTEXT
                        UPCCreateResponse()
                        Exit Select
                    ElseIf SCANTEXT = "Show Ctn" Then
                        Show_carton()
                        Exit Select
                    ElseIf SCANTEXT = "Cls Ctn" Then
                        Print_carton()
                        CreateResponse("NEW_CARTON", "G", PickMessage())
                        Exit Select
                    Else
                        UNITS_MOVED_LST = 0
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("UPC_CODE") Then
                            UPC_CODE = CheckResponse("UPC_CODE")
                            STYLE_CODE = CheckResponse("STYLE_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            COLOR_DESC = CheckResponse("COLOR_DESC")
                            rowWHTSCANS = tbl.Select("UPC_CODE = '" & UPC_CODE & "'").FirstOrDefault
                            If rowWHTSCANS IsNot Nothing Then
                                ProcessQty(0)
                                Exit Select
                            End If
                        Else
                            If CheckResponse.ContainsKey("STYLE_CODE") Then
                                STYLE_CODE = CheckResponse("STYLE_CODE")
                                CreateResponse("SCAN_COLOR", "G", PickMessage())
                                Exit Select
                            End If
                        End If
                    End If
                    'Error
                    hold = AppStates(AppState)
                    AppStates(AppState) = "UPC/Style not in shipment, try again|OK|"
                    CreateResponse("", "R", PickMessage())
                    AppStates(AppState) = hold

                Case "SCAN_COLOR"
                    Dim hold As String
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        'CreateResponse("SCAN_UPC", "BLUE", PickMessage)
                        UPCCreateResponse()
                        Exit Select
                    Else
                        If SCANTEXT = "0" Then
                            TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)
                            CreateResponse("", "B", "Style " & STYLE_CODE & " has been selected, All colors " & colors)
                            Exit Select
                        Else
                            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyColor(Me, SCANTEXT, STYLE_CODE, COLOR_CODEs)
                            If CheckResponse.ContainsKey("Error") Then
                                hold = AppStates("SCAN_UPC")
                                AppStates("SCAN_UPC") = "UPC/Style-Clr Not On System, Try again|OK|"
                                CreateResponse("SCAN_UPC", "R", PickMessage())
                                AppStates("SCAN_UPC") = hold
                                Exit Select
                            End If
                            UPC_CODE = CheckResponse("UPC_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            COLOR_DESC = CheckResponse("COLOR_DESC")
                            rowWHTSCANS = tbl.Select("UPC_CODE = '" & UPC_CODE & "'").First
                            If Not rowWHTSCANS Is Nothing Then
                                ProcessQty(0)
                                Exit Select
                            End If
                        End If
                    End If
                    hold = AppStates("SCAN_UPC")
                    AppStates("SCAN_UPC") = "Style Color not in shipment, try again|OK|"
                    CreateResponse("SCAN_UPC", "R", PickMessage())
                    AppStates("SCAN_UPC") = hold

                Case "PRINT_UPC"
                    CreateResponse("SCAN_UNITS", "G", PickMessage())
                    'Print large or small labels depending on response
                    'not yet implemented.


                Case "SCAN_UNITS"
                    If SCANTEXT = "CANCEL" Then
                        'CreateResponse("SCAN_UPC", "BLUE", PickMessage())
                        If CartonType = "Remove" Then
                            CartonType = "Inners"
                        End If
                        UPCCreateResponse()
                        Exit Select
                    Else
                        Dim hold As String
                        'Can we have more than 999 loose units to Move in a Pack?
                        If Val(SCANTEXT) > 9999 Or Val(SCANTEXT) < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Count " & SCANTEXT & ", Verify Units Count|OK|"
                            CreateResponse("", "R", PickMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        ProcessQty(Val(SCANTEXT))
                    End If

                Case "VERIFY"
                    If SCANTEXT.ToUpper = "Y" Then
                        ' Y is for normal Verify
                        Update_Record()
                        If CartonType = "FULL" Then
                            Print_carton()
                            CreateResponse("NEW_CARTON", "", PickMessage())
                        Else
                            'CreateResponse("SCAN_UPC", "BLUE", PickMessage())
                            UPCCreateResponse()
                        End If
                    ElseIf SCANTEXT = "YES" Then
                        ' YES is for Clear Verify
                        'CreateResponse("SCAN_UPC", "BLUE", PickMessage())
                        UPCCreateResponse()
                    ElseIf SCANTEXT.ToUpper = "N" Or SCANTEXT = "NO" Then
                        ' N is for normal Verify
                        ' NO is for Clear Verify
                        CreateResponse("SCAN_UNITS", "G", PickMessage())
                    ElseIf SCANTEXT = "CLEAR" Then
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Clear Current Pack |YES|NO|"
                        CreateResponse("", "R", PickMessage())
                        AppStates(AppState) = hold
                        Exit Select
                    Else
                        'Error
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Invalid Response|OK|"
                        CreateResponse("", "R", PickMessage())
                        AppStates(AppState) = hold
                    End If
            End Select
        End If
    End Sub

    Sub CreateBulkCarton()

        For Each rowWHTSCANS In tbl.Select("")
            Bulk = True
            rowWHTSCANS.Item("PACK_QTY") = rowWHTSCANS.Item("PICK_QTY")
            UNITS_MOVED = rowWHTSCANS.Item("PICK_QTY")
            PACK_QTY_OPEN = 0
            CARTONS_PER_UNIT = 1
            UPC_CODE = rowWHTSCANS.Item("UPC_CODE")
            STYLE_CODE = rowWHTSCANS.Item("STYLE_CODE")
            COLOR_CODE = rowWHTSCANS.Item("COLOR_CODE")
            Update_Record()
        Next

    End Sub

    Private Sub UPCCreateResponse(ByVal Optional CLR As String = "BLUE", ByVal Optional msg As String = "")
        Dim hold As String
        Dim nxtMode As String
        Select Case CartonType
            Case "Inners"
                nxtMode = "|Units"
            Case "Units"
                nxtMode = "|Inners"
            Case Else
                nxtMode = ""
        End Select
        If msg = "" Then msg = PickMessage()
        hold = AppStates("SCAN_UPC")
        If CartonType = "Remove" Then
            AppStates("SCAN_UPC") = AppStates("SCAN_UPC").Replace("|MODE|Remove", nxtMode)
        Else
            AppStates("SCAN_UPC") = AppStates("SCAN_UPC").Replace("|MODE", nxtMode)
        End If
        CreateResponse("SCAN_UPC", CLR, msg)
        AppStates("SCAN_UPC") = hold

    End Sub

    Private Sub Show_carton()
        Dim hold As String
        Dim rowWHTCART1 As DataRow
        Dim msg As String
        If CART_NO = "" Then
            msg = "No Active Carton"
        Else
            rowWHTCART1 = dst.Tables("WHTCART1").Select("CART_NO = '" & CART_NO & "'").First
            msg = "Packed in Ctn " & rowWHTCART1.Item("CART_SEQ") & vbCrLf
            For Each rowWHTCART2 As DataRow In dst.Tables("WHTCART2").Select("CART_NO = '" & CART_NO & "'", "CART_LNO desc")
                msg = msg & "Style " & rowWHTCART2.Item("STYLE_CODE") & " Clr " & rowWHTCART2.Item("COLOR_CODE") & " Qty " & rowWHTCART2.Item("QTY_PACKED") & vbCrLf
            Next
        End If

        hold = AppStates(AppState)
        AppStates(AppState) = "Continue|OK|"
        CreateResponse("", "G", msg)
        AppStates(AppState) = hold
        Exit Sub


    End Sub

    Private Sub ProcessQty(ByVal Units As Integer)
        Dim hold As String
        CARTON_PACK_QTY = rowWHTSCANS.Item("CARTON_PACK_QTY")
        CARTONS_PER_UNIT = rowWHTSCANS.Item("CARTONS_PER_UNIT")
        If (CARTONS_PER_UNIT > 1 Or CARTON_PACK_QTY = 1) And CartonType <> "FULL" Then
            hold = AppStates(AppState)
            AppStates(AppState) = "Wrong Carton Type, Pay attention|OK|"
            CreateResponse("", "R", PickMessage())
            AppStates(AppState) = hold
        End If
        If CartonType = "FULL" And Units = 0 Then
            UNITS_MOVED = CARTON_PACK_QTY
        ElseIf CartonType = "Inners" And Units = 0 Then
            UNITS_MOVED = rowWHTSCANS.Item("INNER_PACK_QTY")
        Else
            UNITS_MOVED = Units
            If CartonType = "Remove" Then
                UNITS_MOVED = -1 * UNITS_MOVED
            End If
        End If

        If UNITS_MOVED = 0 Then
            If CartonType = "Remove" Then
                hold = AppStates("SCAN_UNITS")
                AppStates("SCAN_UNITS") = "Units Removed|CANCEL|"
                CreateResponse("SCAN_UNITS", "R", PickMessage())
                AppStates("SCAN_UNITS") = hold
            Else
                ' for this new customers we need to print UPC labels before packing
                If CUST_CODE = "320214" Then
                    CreateResponse("PRINT_UPC", "G", PickMessage())
                Else
                    CreateResponse("SCAN_UNITS", "G", PickMessage())
                End If
            End If
            Exit Sub
        End If

        'PACK_QTY_OPEN = rowWHTSCANS.Item("PICK_QTY") - rowWHTSCANS.Item("PACK_QTY") - UNITS_MOVED
        PACK_QTY_OPEN = Val(dst.Tables("WHTSCANS").Compute("SUM(PICK_QTY) - SUM(PACK_QTY)", "UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'")) - UNITS_MOVED
        If CartonType = "Remove" And rowWHTSCANS.Item("PACK_QTY") - UNITS_MOVED > 0 Then
            PACK_QTY_OPEN = rowWHTSCANS.Item("PICK_QTY") - rowWHTSCANS.Item("PACK_QTY")
            hold = AppStates(AppState)
            AppStates(AppState) = "Overage, Check Count|OK|"
            CreateResponse("", "R", PickMessage())
            AppStates(AppState) = hold
            Exit Sub
        End If

        If PACK_QTY_OPEN >= 0 Then
            If CartonType = "FULL" Then
                CreateResponse("NEW_CARTON", "G", PickMessage())
                Print_carton()
            Else
                UPCCreateResponse()
            End If
            Update_Record()
            If CartonType = "Remove" Then
                CartonType = "Inners"
                UPCCreateResponse()
            ElseIf CartonType = "FULL" Then
                Print_carton()
            End If
        Else
            hold = AppStates(AppState)
            AppStates(AppState) = "Overage, moving " & UNITS_MOVED & " Units|OK|"
            CreateResponse("", "R", PickMessage())
            AppStates(AppState) = hold
        End If
    End Sub
    Function PickMessage() As String
        ''If OpenPicks = 0 Then
        ''    Return "PT " & PICK_NO & " Has no Open picks "
        ''Else
        'Dim PickQty As String = " :"

        ''I feel like I should ask for inners when displaying inners
        'Dim Open As Integer = PACK_QTY_OPEN
        'If CARTONS_PER_UNIT > 0 Then
        '    PickQty = PickQty & String.Format(" {0}ttl - {1}c per unit", Math.Truncate(Open * CARTONS_PER_UNIT), CARTONS_PER_UNIT)
        '    Open = 0
        'Else
        '    If CARTON_PACK_QTY > 0 And CARTON_PACK_QTY < Open Then
        '        PickQty = PickQty & String.Format(" {0}c", Math.Truncate(Open / CARTON_PACK_QTY))
        '        Open = Open - Math.Truncate(Open / CARTON_PACK_QTY) * CARTON_PACK_QTY
        '    End If
        '    If INNER_PACK_QTY > 0 And INNER_PACK_QTY < Open Then
        '        PickQty = PickQty & String.Format(" {0}i", Math.Truncate(Open / INNER_PACK_QTY))
        '        Open = Open - Math.Truncate(Open / INNER_PACK_QTY) * INNER_PACK_QTY
        '    End If
        'End If
        'If Open > 0 Then
        '    PickQty = PickQty & String.Format(" {0}u", Open)
        'End If
        'If Open = PACK_QTY_OPEN Then
        '    PickQty = "u"
        'End If
        Dim Cartons As Integer = dst.Tables("WHTCART1").Rows.Count
        Return " Mode: " & CartonType & vbCrLf _
                                    & "PT " & PICK_NO & " Packed " & Cartons & vbCrLf _
                                    & "Style " & STYLE_CODE & " - " & COLOR_CODE & ":" & COLOR_DESC & vbCrLf _
                                    & "Pack Qty " & UNITS_MOVED_LST & " Open Qty " & PACK_QTY_OPEN
        'End If
    End Function

    Sub Print_carton()

        ASCMAIN1.sql = "Update WHTCART1 set PROCESS_STATUS = '1' where :PARM1 in (CART_NO, CART_NO_CONS)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", CART_NO)

    End Sub

    Sub Update_Record()
        Dim HoldLoc As String = ""
        'PACK_QTY_OPEN = PACK_QTY_OPEN - UNITS_MOVED
        'rowWHTSCANS.Item("PACK_QTY") = Val(rowWHTSCANS.Item("PACK_QTY") & "") + UNITS_MOVED
        Dim CART_SEQ As Integer = 0

        BeginTrans()
        Dim rowWHTCART1 As DataRow
        If CART_NO = "" Then
            'Using SOTCART1's control number because this will ultimately become an SOTCART1 Record.
            CART_NO = TACMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & SO_PARM_UPC_VENDOR_ID)

            ASCMAIN1.sql = "Update SOTPICK1 Set PACK_STATUS = 'P', CART_SEQ_CTR = nvl(CART_SEQ_CTR,0) + 1, PACK_PACKER = '" & G.USER_ID & "'" & vbCrLf _
                & "     where PICK_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_NO)

            If BILL_OF_LADING_NO <> "" Then
                Fill_Records("WHTCART1", "", True, "SELECT C1.* FROM WHTCART1 C1, SOTPICK1 P1, SOTSHIP1 S1 WHERE C1.PICK_NO = P1.PICK_NO AND S1.SHIP_BOL_NO = P1.SHIP_BOL_NO AND S1.BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "' ")
            Else
                Fill_Records("WHTCART1", "", True, "Select * from WHTCART1 where PICK_NO = '" & PICK_NO & "' ")
            End If

            For Each rowWHTCART1 In ASCDATA1.SelectDistinct(dst.Tables("WHTCART1"), "CART_SEQ").Select("", "CART_SEQ")
                If rowWHTCART1.Item("CART_SEQ") <> CART_SEQ + 1 And rowWHTCART1.Item("CART_SEQ") <> CART_SEQ Then  ' just in case we have dup seq no
                    Exit For
                End If
                CART_SEQ = rowWHTCART1.Item("CART_SEQ")
            Next
            CART_SEQ += 1

            rowWHTCART1 = dst.Tables("WHTCART1").NewRow
            With rowWHTCART1
                .Item("CART_NO") = CART_NO
                .Item("PICK_NO") = PICK_NO
                .Item("CART_PACKER") = G.USER_ID
                .Item("CART_PACKED") = DATETIME_STAMP
                .Item("GUN_ID") = G.GUN_LOC
                .Item("PRINTER") = Printer
                .Item("PROCESS_STATUS") = "0"
                .Item("CART_SEQ") = CART_SEQ
                .Item("CARTONS_PER_UNIT") = CARTONS_PER_UNIT
            End With
            dst.Tables("WHTCART1").Rows.Add(rowWHTCART1)
        Else
            rowWHTCART1 = dst.Tables("WHTCART1").Select("CART_NO = '" & CART_NO & "'").First
        End If
        rowWHTCART1.Item("CART_TOTAL_UNITS") = Val(rowWHTCART1.Item("CART_TOTAL_UNITS") & "") + UNITS_MOVED

        If CARTONS_PER_UNIT > 1 Then
            rowWHTCART1.Item("CART_NO_CONS") = CART_NO
            For cnt As Integer = 1 To CARTONS_PER_UNIT - 1
                'Create Empty Cartons for multipack
                Dim Extra_CART_NO = TACMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & SO_PARM_UPC_VENDOR_ID)

                ASCMAIN1.sql = "Update SOTPICK1 Set PACK_STATUS = 'P', CART_SEQ_CTR = nvl(CART_SEQ_CTR,0) + 1, PACK_PACKER = '" & G.USER_ID & "'" & vbCrLf _
                    & "     where PICK_NO = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_NO)

                'Fill_Records("SOTPICK1", PICK_NO)
                'Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select("").FirstOrDefault
                CART_SEQ += 1

                rowWHTCART1 = dst.Tables("WHTCART1").NewRow
                With rowWHTCART1
                    .Item("CART_NO") = Extra_CART_NO
                    .Item("PICK_NO") = PICK_NO
                    .Item("CART_PACKER") = G.USER_ID
                    .Item("CART_PACKED") = DATETIME_STAMP
                    .Item("GUN_ID") = G.GUN_LOC
                    .Item("PRINTER") = Printer
                    .Item("PROCESS_STATUS") = "0"
                    .Item("CART_SEQ") = CART_SEQ
                    .Item("CART_TOTAL_UNITS") = 0
                    .Item("CARTONS_PER_UNIT") = CARTONS_PER_UNIT
                    .Item("CART_NO_CONS") = CART_NO
                End With
                dst.Tables("WHTCART1").Rows.Add(rowWHTCART1)
            Next
        End If


        Update_Record_TDA("WHTCART1")
        'This needs to adress items that are in multiple SO lines
        Dim PACK_QTY As Integer = UNITS_MOVED
        For Each row2 As DataRow In dst.Tables("WHTSCANS").Select("UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'", "RELEASE_QTY DESC")

            If (PACK_QTY <> 0 And Val(row2.Item("PACK_QTY") & "") < Val(row2.Item("RELEASE_QTY") & "")) Or Bulk Then

                Dim rowWHTCART2 As DataRow = dst.Tables("WHTCART2").NewRow
                rowWHTCART2.Item("CART_NO") = CART_NO
                rowWHTCART2.Item("CART_LNO") = Val(dst.Tables("WHTCART2").Compute("MAX(CART_LNO)", "CART_NO = '" & CART_NO & "'") & "") + 1
                rowWHTCART2.Item("ORDR_NO") = row2.Item("ORDR_NO")
                rowWHTCART2.Item("ORDR_LNO") = row2.Item("ORDR_LNO")
                'rowWHTCART2.Item("QTY_PACKED") = UNITS_MOVED
                rowWHTCART2.Item("UPC_CODE") = UPC_CODE
                rowWHTCART2.Item("STYLE_CODE") = STYLE_CODE
                rowWHTCART2.Item("COLOR_CODE") = COLOR_CODE
                rowWHTCART2.Item("STYLE_UOM") = row2.Item("STYLE_UOM")

                Dim applied As Integer = 0
                If (Val(row2.Item("PACK_QTY") & "") + PACK_QTY > Val(row2.Item("RELEASE_QTY") & "")) And Bulk = False Then
                    'need to test negative values in units_moved
                    Dim open As Integer = Val(row2.Item("RELEASE_QTY") & "") - Val(row2.Item("PACK_QTY") & "")
                    PACK_QTY = PACK_QTY - open
                    rowWHTCART2.Item("QTY_PACKED") = open
                    applied = open
                Else
                    rowWHTCART2.Item("QTY_PACKED") = PACK_QTY
                    applied = PACK_QTY
                    PACK_QTY = 0
                End If
                row2.Item("PACK_QTY") = Val(row2.Item("PACK_QTY") & "") + applied
                dst.Tables("WHTCART2").Rows.Add(rowWHTCART2)
                Update_Record_TDA("WHTCART2")
            End If
        Next

        PACK_QTY = 0
        For Each rowSOTPICK5 As DataRow In dst.Tables("SOTPICK5").Select("UPC_CODE = '" & UPC_CODE & "'")
            Dim open As Integer = Val(rowSOTPICK5.Item("PICK_QTY") & "") - Val(rowSOTPICK5.Item("PACK_QTY") & "")
            Dim applied As Integer = 0
            If PACK_QTY < UNITS_MOVED And open > 0 Then
                If open > UNITS_MOVED Then
                    applied = UNITS_MOVED
                Else
                    applied = open
                End If
                rowSOTPICK5.Item("PACK_QTY") = Val(rowSOTPICK5.Item("PACK_QTY") & "") + applied
                PACK_QTY = PACK_QTY + applied
            End If
        Next
        Update_Record_TDA("SOTPICK5")


        CommitTrans()

        UPC_CODE = ""
        STYLE_CODE = ""
        COLOR_CODE = ""
        UNITS_MOVED_LST = UNITS_MOVED
        UNITS_MOVED = 0

    End Sub


End Class
