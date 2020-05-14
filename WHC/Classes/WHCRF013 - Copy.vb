Public Class WHCRF013_COPY
    ' Application Receiving no LPN

    Inherits WHCRF000
    Dim Mode As String
    Dim WH_REC_NO As String
    Dim CONTAINER_NO As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim COLOR_DESC As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim UPC_cnt As Integer
    Dim ttl_ctn As Integer
    Dim ttlRcv As Integer
    Dim ORDER_QTY_SHP As Integer
    Dim PO_QTY_SHP As Integer
    Dim PO_QTY_REC As Integer
    Dim Cases As Integer
    Dim Units As Integer
    Dim RcvQty As Integer
    Dim NotOnList As Boolean
    Dim QtyKeyIn As String

    Dim FinLoc As String = ""
    Dim RecLoc As String = ""


    Dim ScanLocMsg As String
    Dim RowNum As Integer

    Dim COLOR_CODEs As New List(Of String)
    Dim colors As String = ""

    Dim PO_SHIPMENT_NO As String
    Dim PO_SHIPMENT_LNO As String

    Dim VoidList() As String
    Dim VoidTranNo As String

    Dim INNER_PACK_QTY As Integer
    Dim CARTON_PACK_QTY As Integer
    Dim CARTONS_PER_UNIT As Integer
    Dim holdScan As String
    Dim _sql As String
    Dim rowWHTSCANS As DataRow

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF013"

        AppStates.Add("MODE", "Receiving Mode|Helper|Receiver|EXIT|") 'Buttons
        'AppStates.Add("MODE", "Receiving Mode|Two Step|One Step|EXIT|") 'Buttons
        AppStates.Add("SCAN_SHPMT", "Scan Shipment Lno |EXIT|") 'Green
        'AppStates.Add("SCAN_LOC", "Scan Loc, 00:'{0}', V:Void|NEXT UPC|PREV UPC|NEXT LOC|PREV LOC|DONE|") 'Yellow
        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style|DONE|") 'Blue
        AppStates.Add("SCAN_COLOR", "Enter Color code|CANCEL|")
        AppStates.Add("SCAN_CASES", "How many cases on Cart, (0 for units)|CANCEL|")
        AppStates.Add("SCAN_UNITS", "How many units on Cart|CANCEL|")
        AppStates.Add("SCAN_LOC", "Scan Loc, 00:'{0}'|AcptLoc|CANCEL|")
        AppStates.Add("VERIFY", "Update|Y|N|CLEAR|EXIT|")
        AppStates.Add("LEAVE", "Units in Gun, Gun must be Empty for Receiving|EXIT|")
        AppStates.Add("SCAN_VOID", "Void Line or 'ALL'|BACK|NEW PICK|EXIT|")

        AppState = "MODE"
        'ScanLocMsg = AppStates("SCAN_LOC")
        LAST_CLR = "R"

        With dst

            ASCMAIN1.sql = "Select  POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO,  POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO " & vbCrLf _
                        & ", ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYC1.UPC_CODE " & vbCrLf _
                        & ", 0 PO_QTY_REC, POTSHIP3.PO_QTY_SHP " & vbCrLf _
                        & ", nvl(ICTSTYL1.CARTON_PACK_QTY, 1) CARTON_PACK_QTY, nvl(ICTSTYL1.CARTONS_PER_UNIT, 0) CARTONS_PER_UNIT " & vbCrLf _
                        & " From POTSHIP3, POTORDR2, ICTSTYL1, ICTSTYC1 " & vbCrLf _
                        & " Where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO " & vbCrLf _
                        & " And POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO " & vbCrLf _
                        & " And ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE " & vbCrLf _
                        & " And ICTSTYC1.STYLE_CODE = POTORDR2.STYLE_CODE " & vbCrLf _
                        & " And ICTSTYC1.COLOR_CODE = POTORDR2.COLOR_CODE " & vbCrLf _
                        & " And POTSHIP3.PO_SHIPMENT_NO = :PARM1" & vbCrLf _
                        & " And POTSHIP3.PO_SHIPMENT_LNO = :PARM2"

            Create_TDA(.Tables.Add, "WHTSCANS", "**", 0, False, "VV", 4)
            .Tables("WHTSCANS").Columns.Add("CARTON_REC", GetType(System.Int32), "PO_QTY_REC / CARTON_PACK_QTY")
            .Tables("WHTSCANS").Columns.Add("CARTON_OPEN", GetType(System.Int32), "(PO_QTY_SHP - PO_QTY_REC) / CARTON_PACK_QTY")
            .Tables("WHTSCANS").Columns.Add("QTY_OPEN", GetType(System.Int32), "PO_QTY_SHP - PO_QTY_REC")


            ASCMAIN1.sql = "SELECT WHTWREC1.* from WHTWREC1 where WHTWREC1.WH_REC_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTWREC1", "**", 0, True, "V", 1)
            Create_TDA(.Tables.Add, "WHTPREC2", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")

            ASCMAIN1.sql = "SELECT * from POTSHIP2 where PO_SHIPMENT_NO = :PARM1 and CONTAINER_NO = :PARM2 and PO_SHIP_STATUS = 'O'"
            Create_TDA(.Tables.Add, "POTSHIP2", "**", 0, False, "VV")

        End With

        Dim styles As String = TACMAIN1.LookupLocation(Me, g.GUN_LOC)
        If styles.Length > 0 Then
            AppState = "LEAVE"
        End If

        FinLoc = ASCDATA1.GetDataValue("select WHSE_LOC_FIN from ICTWHSE1 where whse_code = '" & g.WHSE_CODE & "'")
        RecLoc = ASCDATA1.GetDataValue("select WHSE_LOC_REC from ICTWHSE1 where whse_code = '" & g.WHSE_CODE & "'")

        tbl = dst.Tables("WHTSCANS")

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
            CreateResponse("", "B", ReceiptMessage())
        Else
            Select Case AppState
                Case "MODE"
                    Mode = SCANTEXT
                    CreateResponse("SCAN_SHPMT", "GREEN", "")

                Case "SCAN_SHPMT"
                    Dim styles As String = TACMAIN1.LookupLocation(Me, G.GUN_LOC)
                    If styles.Length > 0 Then
                        CreateResponse("LEAVE", "R", "Gun has Units, Deposit merchandise to continue")
                        Exit Select
                    End If
                    If SCANTEXT.Length < 8 Then
                        SCANTEXT = SCANTEXT.PadLeft(8, "0")
                    End If
                    If SCANTEXT.Contains("-") Then
                        Dim S() As String
                        S = SCANTEXT.Split("-")
                        PO_SHIPMENT_NO = Val(S(0))
                        PO_SHIPMENT_LNO = Val(S(1))
                    Else
                        PO_SHIPMENT_NO = SCANTEXT.Substring(0, 6)
                        PO_SHIPMENT_LNO = SCANTEXT.Substring(6, 2)
                    End If

                    tbl.Clear()
                    WH_REC_NO = ""

                    Dim row As DataRow
                    row = ASCDATA1.GetDataRow("Select POTSHIP1.* from POTSHIP1 where  POTSHIP1.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'")
                    If Not IsNothing(row) Then
                        If row.Item("WHSE_CODE") <> G.WHSE_CODE Then
                            CreateResponse("", "GREEN", "Shipment not available for Wharehouse, ask for Help")
                            Exit Select
                        End If
                    Else
                        CreateResponse("", "GREEN", "Invalid Shipment, Re-Scan")
                        Exit Select
                    End If

                    row = ASCDATA1.GetDataRow("Select POTSHIP2.* from POTSHIP2 where POTSHIP2.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO _
                                              & "' and POTSHIP2.PO_SHIPMENT_LNO = '" & PO_SHIPMENT_LNO & "'")
                    If Not IsNothing(row) Then
                        If row.Item("PO_SHIP_STATUS") <> "O" Then
                            CreateResponse("", "GREEN", "Shipment status is not Open, ask for Help")
                            Exit Select
                        Else
                            CONTAINER_NO = row("CONTAINER_NO")
                        End If
                    Else
                        CreateResponse("", "GREEN", "Invalid Shipment, Re-Scan")
                        Exit Select
                    End If
                    ttl_ctn = 0
                    ttlRcv = 0
                    If Not ASCMAIN1.Logical_Open("POTSHIP1", PO_SHIPMENT_NO) Then
                        CreateResponse("", "GREEN", "PO Shipment is locked.")
                        Exit Select
                    End If

                    Dim PO_SHIPMENT_LNOs As String = ""
                    tbl.Rows.Clear()
                    ' ASCMAIN1.sql = "SELECT * from POTSHIP2 where PO_SHIPMENT_NO = :PARM1 and CONTAINER_NO = :PARM2 and PO_SHIP_STATUS = 'O'" 'fill records
                    Fill_Records("POTSHIP2", New String() {PO_SHIPMENT_NO, CONTAINER_NO}, False)
                    For Each row0 As DataRow In dst.Tables("POTSHIP2").Select("")
                        ttl_ctn = ttl_ctn + row0.Item("PO_SHIP_CTNS")
                        Fill_Records("WHTSCANS", New String() {PO_SHIPMENT_NO, row0.Item("PO_SHIPMENT_LNO")}, False)
                        PO_SHIPMENT_LNOs += ",'" & row0.Item("PO_SHIPMENT_LNO") & "'"
                    Next

                    'For Each row0 As DataRow In dst.Tables("POTSHIP2").Select("")
                    ASCMAIN1.sql = "Select WH_REC_NO, UPC_CODE, sum(PO_QTY_REC) PO_QTY_REC from WHTPREC2 where WHTPREC2.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO _
                                    & "' and WHTPREC2.PO_SHIPMENT_LNO in (" & PO_SHIPMENT_LNOs.Substring(1) & ")" _
                                    & " group by WH_REC_NO, UPC_CODE"
                    For Each row1 As DataRow In ASCDATA1.GetDataTable.Rows
                        Dim upc_lines As Integer = tbl.Compute("Count(UPC_CODE)", "UPC_CODE = '" & row1("UPC_CODE") & "'")
                        Dim upc_cnt As Integer = 0
                        Dim po_qty_rec As Integer = row1("PO_QTY_REC")

                        For Each rowWHTSCANS In tbl.Select("UPC_CODE = '" & row1("UPC_CODE") & "'")
                            upc_cnt += 1
                            If upc_cnt = upc_lines Or rowWHTSCANS.Item("QTY_OPEN") >= po_qty_rec Then
                                rowWHTSCANS.Item("PO_QTY_REC") = rowWHTSCANS.Item("PO_QTY_REC") + po_qty_rec
                                po_qty_rec = 0
                            Else
                                rowWHTSCANS.Item("PO_QTY_REC") = rowWHTSCANS.Item("PO_QTY_REC") + rowWHTSCANS.Item("PO_QTY_SHP")
                                po_qty_rec = po_qty_rec - rowWHTSCANS.Item("PO_QTY_SHP")
                            End If
                        Next
                        ttlRcv += row1("PO_QTY_REC")
                        WH_REC_NO = row1("WH_REC_NO") & ""
                    Next
                    'Next
                    If String.IsNullOrEmpty(WH_REC_NO) Then
                        WH_REC_NO = ASCDATA1.GetDataValue("Select WH_REC_NO from WHTWREC1 where CONTAINER_NO = '" & CONTAINER_NO & "' and PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'")
                    End If

                    Fill_Records("WHTWREC1", WH_REC_NO)

                    UPC_cnt = tbl.Rows.Count
                    ORDER_QTY_SHP = Val(tbl.Compute("SUM(PO_QTY_SHP)", "") & "")

                    RowNum = 0
                    CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())

                Case "SCAN_UPC"
                    If SCANTEXT = "DONE" Then
                        CreateResponse("SCAN_SHPMT", "GREEN", "Scan New Shipment Sheet")
                        Exit Select
                    Else
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        NotOnList = False
                        PO_QTY_SHP = 0
                        PO_QTY_REC = 0

                        If CheckResponse.ContainsKey("UPC_CODE") Then
                            If SCANTEXT = CheckResponse("UPC_CODE") Or SCANTEXT.ToUpper = CheckResponse("STYLE_CODE") Then
                                UPC_CODE = CheckResponse("UPC_CODE")
                                STYLE_CODE = CheckResponse("STYLE_CODE")
                                COLOR_CODE = CheckResponse("COLOR_CODE")
                                LOCATION_CODE = GetLocation(STYLE_CODE, COLOR_CODE)
                                rowWHTSCANS = tbl.Select("UPC_CODE = '" & UPC_CODE & "'").FirstOrDefault
                                If IsNothing(rowWHTSCANS) Then
                                    NotOnList = True
                                    CreateResponse("", "RED", "UPC Not on Shipment")
                                    Exit Select
                                    'Dim row As DataRow = ASCDATA1.GetDataRow("Select CARTON_PACK_QTY, CARTONS_PER_UNIT From ICTSTYL1 where STYLE_CODE = '" & CheckResponse("STYLE_CODE") & "'")
                                    'CARTON_PACK_QTY = row("CARTON_PACK_QTY")
                                    'CARTONS_PER_UNIT = row("CARTONS_PER_UNIT")
                                Else
                                    CARTON_PACK_QTY = rowWHTSCANS("CARTON_PACK_QTY")
                                    CARTONS_PER_UNIT = rowWHTSCANS("CARTONS_PER_UNIT")
                                    PO_QTY_SHP = tbl.Compute("sum(PO_QTY_SHP)", "UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'")
                                    PO_QTY_REC = tbl.Compute("sum(PO_QTY_REC)", "UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'")
                                End If
                                'CreateResponse("SCAN_CASES", "G", ReceiptMessage())
                                If Mode = "Helper" Then
                                    Dim hold1 As String = AppStates("SCAN_UPC")
                                    AppStates("SCAN_UPC") = String.Format("Open {0} - Scan UPC|DONE|", rowWHTSCANS("CARTON_OPEN"))
                                    CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                                    AppStates("SCAN_UPC") = hold1
                                Else
                                    If Mode = "One Step" Then
                                        Dim hold1 As String = AppStates("SCAN_LOC")
                                        AppStates("SCAN_LOC") = String.Format(hold1, LOCATION_CODE)
                                        CreateResponse("SCAN_LOC", "YELLOW", ReceiptMessage())
                                        AppStates("SCAN_LOC") = hold1
                                    ElseIf Mode = "Two Step" Or Mode = "Receiver" Then
                                        CreateResponse("SCAN_CASES", "G", ReceiptMessage())
                                    End If
                                End If
                                Exit Select
                            End If
                        Else
                            If CheckResponse.ContainsKey("STYLE_CODE") Then
                                STYLE_CODE = CheckResponse("STYLE_CODE")
                                TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)
                                CreateResponse("SCAN_COLOR", "G", "Style " & STYLE_CODE & " has been selected, colors " & colors)
                                'CreateResponse("SCAN_COLOR", "G", ReceiptMessage())
                                Exit Select
                            End If
                        End If

                    End If
                    'Error
                    Dim hold As String = AppStates(AppState)
                    AppStates(AppState) = "UPC/Style Not On System, Put aside|OK|"
                    CreateResponse("", "BLUE", ReceiptMessage())
                    AppStates(AppState) = hold

                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    PO_QTY_SHP = 0
                    PO_QTY_REC = 0
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage)
                        Exit Select
                    Else
                        If SCANTEXT = "0" Then
                            TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)
                            CreateResponse("", "B", "Style " & STYLE_CODE & " has been selected, All colors " & colors)
                            Exit Select
                        Else
                            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyColor(Me, SCANTEXT, STYLE_CODE, COLOR_CODEs)
                            If CheckResponse.ContainsKey("Error") Then
                                Dim hold As String = AppStates("SCAN_UPC")
                                AppStates("SCAN_UPC") = "UPC/Style-Clr Not On System, Put aside|OK|"
                                CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                                AppStates("SCAN_UPC") = hold
                                Exit Select
                            End If
                            UPC_CODE = CheckResponse("UPC_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            LOCATION_CODE = GetLocation(STYLE_CODE, COLOR_CODE)
                            Dim rowWHTSCANS = tbl.Select("UPC_CODE = '" & UPC_CODE & "'").First
                            If IsNothing(rowWHTSCANS) Then
                                NotOnList = True
                                CreateResponse("", "RED", "UPC Not on Shipment")
                                Exit Select
                                'Dim row As DataRow = ASCDATA1.GetDataRow("Select CARTON_PACK_QTY, CARTONS_PER_UNIT From ICTSTYL1 where STYLE_CODE = '" & CheckResponse("STYLE_CODE") & "'")
                                'CARTON_PACK_QTY = row("CARTON_PACK_QTY")
                                'CARTONS_PER_UNIT = row("CARTONS_PER_UNIT")
                            Else
                                CARTON_PACK_QTY = rowWHTSCANS("CARTON_PACK_QTY")
                                CARTONS_PER_UNIT = rowWHTSCANS("CARTONS_PER_UNIT")
                                PO_QTY_SHP = tbl.Compute("sum(PO_QTY_SHP)", "UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'")
                                PO_QTY_REC = tbl.Compute("sum(PO_QTY_REC)", "UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'")
                            End If
                            'CreateResponse("SCAN_CASES", "G", ReceiptMessage())
                            If Mode = "Helper" Then
                                Dim hold1 As String = AppStates("SCAN_UPC")
                                AppStates("SCAN_UPC") = String.Format("Open {0} - Scan UPC|DONE|", rowWHTSCANS("CARTON_OPEN"))
                                CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                                AppStates("SCAN_UPC") = hold1
                            Else
                                If Mode = "One Step" Then
                                    Dim hold1 As String = AppStates("SCAN_LOC")
                                    AppStates("SCAN_LOC") = String.Format(hold1, LOCATION_CODE)
                                    CreateResponse("SCAN_LOC", "YELLOW", ReceiptMessage())
                                    AppStates("SCAN_LOC") = hold1
                                ElseIf Mode = "Two Step" Or Mode = "Receiver" Then
                                    CreateResponse("SCAN_CASES", "G", ReceiptMessage())
                                End If
                            End If
                            Exit Select
                        End If

                    End If

                Case "SCAN_LOC"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                        Exit Select
                    Else
                        If SCANTEXT = "00" Or SCANTEXT = "AcptLoc" Then
                            SCANTEXT = LOCATION_CODE
                        End If
                        Dim dResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                        If dResponse.ContainsKey("Error") Then
                            Dim hold As String = AppStates(AppState)
                            AppStates(AppState) = "Invalid Location, Choose New Location|OK|"
                            CreateResponse("", "YELLOW", dResponse("Error"))
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        LOCATION_CODE = SCANTEXT
                        CreateResponse("SCAN_CASES", "G", ReceiptMessage())
                    End If

                Case "SCAN_CASES"
                    '  Not sure how to handle CARTONS_PER_UNIT - since we cant control what's loaded on cart
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                        Exit Select
                    Else
                        Dim hold As String
                        Dim ErrMsg As String = ""
                        If Math.Abs(Val(SCANTEXT)) > 9999 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Number, Enter Cases Picked|OK|"
                            CreateResponse("", "R", ReceiptMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        QtyKeyIn = "Cases " & SCANTEXT

                        If SCANTEXT.Contains("*") Then
                            Dim S() As String
                            S = SCANTEXT.Split("*")
                            Cases = Val(S(0))
                            Units = Val(S(1))
                        Else
                            Cases = Val(SCANTEXT)
                            Units = 0
                        End If

                        If Cases + Units = 0 Then
                            CreateResponse("SCAN_UNITS", "G", ReceiptMessage())
                            Exit Select
                        End If
                        If CARTONS_PER_UNIT > 1 Then
                            RcvQty = Int(Cases / CARTONS_PER_UNIT) + Units
                        Else
                            RcvQty = Cases * CARTON_PACK_QTY + Units
                        End If

                        If PO_QTY_REC + RcvQty > PO_QTY_SHP Then
                            If CARTONS_PER_UNIT > 1 Then
                                ErrMsg = "*OVERAGE " & CARTONS_PER_UNIT * ((PO_QTY_REC + RcvQty) - PO_QTY_SHP) & "*"
                            Else
                                ErrMsg = "*OVERAGE " & ((PO_QTY_REC + RcvQty) - PO_QTY_SHP) / CARTON_PACK_QTY & "*"
                            End If
                        End If

                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = "Update " & SCANTEXT & "Cs" & ErrMsg & "|Y|N|CLEAR|Exit|"
                        CreateResponse("VERIFY", If(ErrMsg = "", "B", "R"), ReceiptMessage())
                        AppStates("VERIFY") = hold
                    End If

                Case "SCAN_UNITS"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                        Exit Select
                    Else
                        Dim hold As String
                        Dim ErrMsg As String = ""
                        'Can we have more than 999 loose units to Move in a Pick?
                        If Math.Abs(Val(SCANTEXT)) > 99999 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Count " & SCANTEXT & ", Verify Units Count|OK|"
                            CreateResponse("", "R", ReceiptMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        QtyKeyIn = "Units" & SCANTEXT
                        Cases = 0
                        Units = Val(SCANTEXT)
                        RcvQty = Units
                        If PO_QTY_REC + RcvQty > PO_QTY_SHP Then
                            ErrMsg = "*OVERAGE " & ((PO_QTY_REC + RcvQty) - PO_QTY_SHP) & "*"
                        End If
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = "Update " & SCANTEXT & "u" & ErrMsg & "|Y|N|CLEAR|Exit|"
                        CreateResponse("VERIFY", If(ErrMsg = "", "B", "R"), ReceiptMessage())
                        AppStates("VERIFY") = hold
                    End If

                Case "SCAN_VOID"
                    If SCANTEXT = "BACK" Then
                        loadLine()
                        'AppStates("SCAN_LOC") = String.Format(ScanLocMsg, GOTO_LOCATION)
                        CreateResponse("SCAN_LOC", "YELLOW", ReceiptMessage())
                        Exit Select
                    ElseIf SCANTEXT = "New PICK" Then
                        CreateResponse("SCAN_PTCKT", "GREEN", ReceiptMessage())
                        Exit Select
                    ElseIf SCANTEXT = "YES" Then
                        voidLine(holdScan)
                        holdScan = ""
                        showPicked("")
                        Exit Select
                    ElseIf SCANTEXT = "NO" Then
                        showPicked("")
                        holdScan = ""
                        Exit Select
                    Else
                        'handle voids
                        If holdScan = "" Then
                            If SCANTEXT.ToUpper = "ALL" Then
                                holdScan = SCANTEXT.ToUpper
                                showPicked(holdScan)
                                Exit Select
                            End If
                            If Val(SCANTEXT) > (VoidList.Length - 1) Or Val(SCANTEXT) < 1 Then
                                'error
                            Else
                                holdScan = SCANTEXT
                                showPicked(SCANTEXT)
                                Exit Select
                            End If
                        Else
                            showPicked(holdScan)
                            Exit Select
                        End If
                    End If
                    showPicked("")

                Case "VERIFY"
                    If SCANTEXT.ToUpper = "Y" Then
                        ' Y is for normal Verify
                        Update_Record()

                        UPC_CODE = ""
                        CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                    ElseIf SCANTEXT = "YES" Then
                        ' YES is for Clear Verify
                        UPC_CODE = ""
                        CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                    ElseIf SCANTEXT.ToUpper = "N" Or SCANTEXT = "NO" Then
                        ' N is for normal Verify
                        ' NO is for Clear Verify
                        CreateResponse("SCAN_CASES", "G", ReceiptMessage())
                    ElseIf SCANTEXT = "CLEAR" Then
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Clear Current Receipt |YES|NO|"
                        CreateResponse("", "R", ReceiptMessage())
                        AppStates(AppState) = hold
                        Exit Select
                    Else
                        'Error
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Invalid Response|OK|"
                        CreateResponse("", "R", ReceiptMessage())
                        AppStates(AppState) = hold

                    End If
            End Select
        End If
    End Sub

    Sub loadLine()

        'GOTO_LOCATION = rowWHTSCANS.Item("LOCATION_CODE")
        RowNum = rowWHTSCANS.Item("ROWNUM")
        STYLE_CODE = rowWHTSCANS.Item("STYLE_CODE")
        COLOR_CODE = rowWHTSCANS.Item("COLOR_CODE")
        COLOR_DESC = rowWHTSCANS.Item("COLOR_CODE") & ": " & rowWHTSCANS.Item("COLOR_DESC")
        'PICK_QTY_OPEN = rowWHTSCANS.Item("PICK_QTY_OPEN")
        CARTON_PACK_QTY = rowWHTSCANS.Item("CARTON_PACK_QTY")
        INNER_PACK_QTY = rowWHTSCANS.Item("INNER_PACK_QTY")
        UPC_CODE = rowWHTSCANS.Item("UPC_CODE")
        'PICK_LNO = rowWHTSCANS.Item("PICK_LNO")
        CARTONS_PER_UNIT = rowWHTSCANS.Item("CARTONS_PER_UNIT")
        'ORIGINAL_LOCATION = GOTO_LOCATION

    End Sub

    Function GetLocation(ByVal Style As String, ByVal Color As String) As String
        Dim rtn_row As DataRow

        ASCMAIN1.sql = "Select * from WHTLOCB1 " & vbCrLf _
            & "     where LOCATION_CODE > '99Z'" & vbCrLf _
            & "     and LAST_DATE > sysdate - .25" & vbCrLf _
            & "     and LOCATION_QTY > 0" & vbCrLf _
            & "     and STYLE_CODE = :PARM1" & vbCrLf _
            & "     and COLOR_CODE = :PARM2" & vbCrLf _
            & "     order by LAST_DATE DESC"
        rtn_row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {Style, Color})
        If IsNothing(rtn_row) Then
            ASCMAIN1.sql = "Select * From WHTPREC3, POTORDR2" & vbCrLf _
            & "     Where WHTPREC3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "     And WHTPREC3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO " & vbCrLf _
            & "     And PO_SHIPMENT_NO = :PARM1" & vbCrLf _
            & "     and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3 "
            rtn_row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {PO_SHIPMENT_NO, Style, Color})
        End If
        If Not IsNothing(rtn_row) Then
            LOCATION_CODE = rtn_row.Item("LOCATION_CODE") & ""
        Else
            LOCATION_CODE = "NO LOCATION"
        End If
        Return LOCATION_CODE

    End Function
    Function ReceiptMessage() As String
        Dim shp_ctn As Integer
        Dim rcv_ctn As Integer

        If PO_SHIPMENT_NO = "" Then
            Return "Scan PO Shipment Document To start"
        End If
        Dim UPC_DETAIL As String = ""
        Dim PO_QTY_SHIP As Integer
        Dim PO_QTY_REC As Integer

        If UPC_CODE <> "" Then
            PO_QTY_SHIP = tbl.Compute("sum(PO_QTY_SHP)", "UPC_CODE = '" & UPC_CODE & "'")
            PO_QTY_REC = tbl.Compute("sum(PO_QTY_REC)", "UPC_CODE = '" & UPC_CODE & "'")
            If CARTONS_PER_UNIT > 1 Then
                shp_ctn = Math.Round(PO_QTY_SHP * rowWHTSCANS.Item("CARTONS_PER_UNIT"))
                rcv_ctn = Math.Round(PO_QTY_REC * rowWHTSCANS.Item("CARTONS_PER_UNIT"))
                UPC_DETAIL = String.Format("{0}c per unit ", CARTONS_PER_UNIT & "".PadLeft(CARTONS_PER_UNIT), "*")
            Else
                shp_ctn = Math.Round(PO_QTY_SHP / rowWHTSCANS.Item("CARTON_PACK_QTY"))
                rcv_ctn = Math.Round(PO_QTY_REC / rowWHTSCANS.Item("CARTON_PACK_QTY"))
                UPC_DETAIL = String.Format("{0}u per Ctn", CARTON_PACK_QTY)
            End If
            UPC_DETAIL = vbCrLf & STYLE_CODE & "-" & COLOR_CODE & " " & rcv_ctn & "/" & shp_ctn & " - " & UPC_DETAIL & vbCrLf & LOCATION_CODE
        End If

        Dim rcv_ctn_ttl As Integer = Val(tbl.Compute("sum(CARTON_REC)", "PO_QTY_REC > 0") & "")

        Return "PO Shipment " & PO_SHIPMENT_NO & " Lno " & PO_SHIPMENT_LNO & vbCrLf _
                & "Expect " & UPC_cnt & " UPCs And " & rcv_ctn_ttl & "/" & ttl_ctn & " cartons" & vbCrLf _
                & "Units Received " & ttlRcv & " / " & ORDER_QTY_SHP _
                & UPC_DETAIL _
                & If(NotOnList, vbCrLf & UPC_CODE & " UPC Not In PO Shipment", "")

    End Function

    Sub showPicked(ByRef showline As String)
        Dim msg As String = "Picked For " & "PICK_NO"
        Dim lno As Integer = 0
        Dim lines As String = ""
        'ASCMAIN1.sql = "Select UPC_CODE, PICK_LNO, WHSE_TRAN_NO, PICK_CASES, PICK_UNITS, INIT_DATE from SOTPICK5 " & vbCrLf _
        '                & " where SOTPICK5.PICK_NO = '" & PICK_NO & "' and PICK_STATUS = 'P'"
        Dim rows() As DataRow = dst.Tables("SOTPICK5").Select("PICK_STATUS = 'P'", "INIT_DATE DESC")
        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                lno += 1
                Dim prows() As DataRow = tbl.Select("PICK_LNO = '" & ROW.Item("PICK_LNO") & "'")
                If showline = "" Or showline = "ALL" Or Val(showline) = lno Then
                    msg = msg & vbCrLf & "L" & lno & " " & prows(0).Item("LOCATION_CODE") & " " & prows(0).Item("STYLE_CODE") & " " & prows(0).Item("COLOR_CODE") & " " _
                        & ROW.Item("PICK_CASES") & "c " & ROW.Item("PICK_UNITS") & "u"
                End If
                lines &= "|" & ROW.Item("PICK_LNO")
                If lno = 5 Then Exit For
            Next
            VoidList = lines.Split("|")
            If showline <> "" Then
                Dim hold As String = AppStates(AppState)
                AppStates(AppState) = If(showline = "ALL", "Void ALL Picked lines", "Void this line") & "|YES|NO|"
                CreateResponse("", "R", msg)
                AppStates(AppState) = hold
            Else
                CreateResponse("SCAN_VOID", "R", msg)
            End If
        Else
            msg &= vbCrLf & "No Lines to Void" & vbCrLf & "Back to continue"
            VoidList = {}
            Dim hold As String = AppStates("SCAN_VOID")
            AppStates("SCAN_VOID") = "No Picks, CLICK BELOW|BACK|NEW PICK|EXIT|"
            CreateResponse("SCAN_VOID", "R", msg)
            AppStates("SCAN_VOID") = hold
        End If

    End Sub

    Sub voidLine(ByRef SCANTEXT As String)
        Dim rows() As DataRow
        Dim ss As String
        If SCANTEXT = "ALL" Then
            ss = "PICK_QTY <> PICK_QTY_OPEN"
        Else
            Dim void = VoidList(Val(SCANTEXT))
            'VoidTranNo = void(1)
            ss = "PICK_LNO = '" & void & "'"
        End If
        rows = tbl.Select(ss)
        If rows.Length = 0 Then
            'error
        Else
            For Each row As DataRow In rows
                'Figure out multipicks for SOTPICK5
                Dim rowSOTPICK5 As DataRow = dst.Tables("SOTPICK5").Select("PICK_STATUS = 'P' AND PICK_LNO = '" & row.Item("PICK_LNO") & "'").FirstOrDefault
                rowSOTPICK5.Item("PICK_STATUS") = "V"

                'CASES_MOVED = (rowSOTPICK5.Item("PICK_CASES") * -1)
                'UNITS_MOVED = (rowSOTPICK5.Item("PICK_UNITS") * -1)
                'LOCATION_CODE = rowSOTPICK5.Item("LOCATION_CODE")

                'OpenPicks = OpenPicks + 1
                rowWHTSCANS = row
                loadLine()
                Update_Record()
            Next
        End If

    End Sub

    Sub Update_Record()
        Dim rowWHTWREC1 As DataRow
        rowWHTSCANS.Item("PO_QTY_REC") = rowWHTSCANS.Item("PO_QTY_REC") + RcvQty

        Dim NEW_LOCATION As String = ""
        If Mode = "Two Step" Or Mode = "Receiver" Then
            NEW_LOCATION = RecLoc
        ElseIf Mode = "One Step" Then
            NEW_LOCATION = LOCATION_CODE
        End If

        BeginTrans()

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
        Dim WHSE_TRAN_LNO_ctr As Integer = 1

        'WHTWREC1
        If WH_REC_NO = "" Then
            WH_REC_NO = ASCMAIN1.Next_Control_No("WHTWREC1.WH_REC_NO")
            rowWHTWREC1 = dst.Tables("WHTWREC1").NewRow
            With rowWHTWREC1
                .Item("WH_REC_NO") = WH_REC_NO
                .Item("WH_DATE_RECEIVED") = DATETIME_STAMP
                .Item("WHSE_CODE") = G.WHSE_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = G.USER_ID
                .Item("LAST_OPER") = G.USER_ID
                .Item("WH_REC_STATUS") = "O"
                .Item("OPS_YYYYPP") = ""
                .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIP_VIA") = ""
                .Item("WH_DATE_DELIVERED") = DATETIME_STAMP
                .Item("WH_DATE_UNLOADED") = DATETIME_STAMP
                .Item("UNLOADED_BY_OPER") = G.USER_ID
                '.Item("TRAILER_NO") =""
                '.Item("CONTAINER_SEAL_NO") = ""
                '.Item("CONTAINER_SEAL_INTACT") = ""
                '.Item("WH_REC_EMAIL_SENT") = ""
                .Item("CONTAINER_NO") = CONTAINER_NO
            End With
            dst.Tables("WHTWREC1").Rows.Add(rowWHTWREC1)
        Else
            rowWHTWREC1 = dst.Tables("WHTWREC1").Select("WH_REC_NO = '" & WH_REC_NO & "'").First
            With rowWHTWREC1
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = G.USER_ID
            End With
        End If
        Update_Record_TDA("WHTWREC1")

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        With rowWHTMOVE1
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_TYPE") = "W"
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("WHSE_CODE") = G.WHSE_CODE
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = G.USER_ID
            .Item("STATUS") = "U"
        End With
        dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
        Update_Record_TDA("WHTMOVE1")

        Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
        With rowWHTMOVE2
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
            .Item("LOCATION_CODE_FROM") = FinLoc
            .Item("LOCATION_CODE_TO") = NEW_LOCATION
            .Item("BAR_CODE") = "0000000000"
            .Item("WHSE_TRAN_QTY") = RcvQty
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("STATUS") = "U"
            .Item("LOAD_NO_FROM") = ""
            .Item("LOAD_NO_TO") = ""
        End With
        dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
        Update_Record_TDA("WHTMOVE2")

        Dim rowWHTPREC2 As DataRow = dst.Tables("WHTPREC2").NewRow
        With rowWHTPREC2
            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("UPC_CODE") = UPC_CODE
            .Item("REC_KEYIN") = QtyKeyIn
            .Item("PO_QTY_REC") = RcvQty
            .Item("GUN_ID") = G.GUN_LOC
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("WH_REC_NO") = WH_REC_NO
        End With
        dst.Tables("WHTPREC2").Rows.Add(rowWHTPREC2)

        ttlRcv += RcvQty
        RcvQty = 0

        Update_Record_TDA("WHTPREC2")

        ASCMAIN1.sql = "Update POTSHIP2 Set WH_REC_NO = :PARM1 where PO_SHIPMENT_NO = :PARM2 and CONTAINER_NO = :PARM3"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {WH_REC_NO, PO_SHIPMENT_NO, CONTAINER_NO})

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        CommitTrans()


    End Sub


End Class
