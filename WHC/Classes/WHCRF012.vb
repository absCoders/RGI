Public Class WHCRF012
    ' Application Putaway to location for two step receiving- for warehouses with no LPN

    Inherits WHCRF000


    Dim LOCATION_FROM As String
    Dim LOCATION_TO As String
    Dim LOCATION_CODE As String
    Dim LOCATION_USE As String
    Dim TICKET_NO As String
    Dim WHSE_TRAN_TYPE As String
    Dim Styles As String = ""
    Dim FromLoc As String = ""
    Dim DmgLoc As String = ""
    Dim LNFLoc As String = ""
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim COLOR_DESC As String
    Dim UPC_CODE As String
    Dim InReceiving As Integer
    Dim Cases As Integer
    Dim Units As Integer
    Dim RcvQty As Integer
    Dim appname As String
    Dim INNER_PACK_QTY As Integer
    Dim CARTON_PACK_QTY As Integer
    Dim CARTONS_PER_UNIT As Integer

    Dim COLOR_CODEs As New List(Of String)
    Dim colors As String = ""

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF012"

        AppStates.Add("SCAN_UPC", "Scan UPC to Putaway|EXIT|") ' Red
        AppStates.Add("SCAN_COLOR", "Enter Color code|CANCEL|")
        AppStates.Add("SCAN_LOC", "Scan Putaway Location |CANCEL|") ' YELLOW
        AppStates.Add("SCAN_CASES", "How many cases, (0 for units)|CANCEL|")
        AppStates.Add("SCAN_UNITS", "How many units|CANCEL|")
        AppStates.Add("VERIFY", "Update (Y/N)|Y|N|CANCEL|")
        AppStates.Add("GET_BTNS", "OK to Continue, or New Location|NEWLOC|OK|")

        AppState = "SCAN_UPC"
        LAST_CLR = "BLUE"

        With dst

            Create_TDA(.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            Create_TDA(.Tables.Add, "SOTRMAFR", "*")

        End With

        If g.PICK_TYPE = "S" Then 'RTS
            FromLoc = ASCDATA1.GetDataValue("select WHSE_LOC_RTN from ICTWHSE1 where whse_code = '" & g.WHSE_CODE & "'")
            appname = "Returns"
        ElseIf g.PICK_TYPE = "C" Then 'RCV
            FromLoc = ASCDATA1.GetDataValue("select WHSE_LOC_REC from ICTWHSE1 where whse_code = '" & g.WHSE_CODE & "'")
            appname = "Receiving"
        ElseIf g.PICK_TYPE = "D" Then 'Damage - Breakage
            DmgLoc = ASCDATA1.GetDataValue("select WHSE_LOC_DST from ICTWHSE1 where whse_code = '" & g.WHSE_CODE & "'")
            LNFLoc = ASCDATA1.GetDataValue("select WHSE_LOC_LNF from ICTWHSE1 where whse_code = '" & g.WHSE_CODE & "'")
            appname = "Damages"
            AppStates("SCAN_UPC") = "Scan Damaged UPC|EXIT|"
        End If

        tbl = dst.Tables("WHTMOVE2") ' New DataTable

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
        Else
            Select Case AppState
                Case "SCAN_UPC"
                    SCANTEXT = Trim(SCANTEXT)
                    RcvQty = 0
                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                    LOCATION_USE = ""
                    COLOR_DESC = ""
                    If CheckResponse.ContainsKey("LOCATION_USE") Then
                        LOCATION_USE = CheckResponse("LOCATION_USE")
                        COLOR_DESC = "(III-IV)"
                    End If
                    If CheckResponse.ContainsKey("UPC_CODE") Then
                        If SCANTEXT = CheckResponse("UPC_CODE") Or SCANTEXT.ToUpper = CheckResponse("STYLE_CODE") Then
                            UPC_CODE = CheckResponse("UPC_CODE")
                            STYLE_CODE = CheckResponse("STYLE_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")

                            LOCATION_CODE = GetLocation(STYLE_CODE, COLOR_CODE)
                            InReceiving = GetInReceiving(STYLE_CODE, COLOR_CODE)

                            If InReceiving = 0 Then
                                CreateResponse("", "RED", "UPC Not found on " & appname)
                                Exit Select
                                'Dim row As DataRow = ASCDATA1.GetDataRow("Select CARTON_PACK_QTY, CARTONS_PER_UNIT From ICTSTYL1 where STYLE_CODE = '" & CheckResponse("STYLE_CODE") & "'")
                                'CARTON_PACK_QTY = row("CARTON_PACK_QTY")
                                'CARTONS_PER_UNIT = row("CARTONS_PER_UNIT")
                            Else
                                Dim row As DataRow = ASCDATA1.GetDataRow("Select nvl(CARTON_PACK_QTY,1) CARTON_PACK_QTY, nvl(CARTONS_PER_UNIT,1) CARTONS_PER_UNIT From ICTSTYL1 where STYLE_CODE = '" & CheckResponse("STYLE_CODE") & "'")
                                CARTON_PACK_QTY = row("CARTON_PACK_QTY")
                                CARTONS_PER_UNIT = row("CARTONS_PER_UNIT")
                                If G.PICK_TYPE = "D" Then
                                    CreateResponse("SCAN_CASES", "YELLOW", ReceiptMessage())
                                Else
                                    CreateResponse("SCAN_LOC", "YELLOW", ReceiptMessage())
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

                    'Error
                    Dim hold As String = AppStates(AppState)
                    AppStates(AppState) = "UPC Not On System, Put aside, Next UPC|OK|"
                    STYLE_CODE = ""
                    COLOR_CODE = ""
                    UPC_CODE = ""
                    CreateResponse("", "BLUE", ReceiptMessage())
                    AppStates(AppState) = hold

                Case "SCAN_COLOR"
                    Dim colors As String = ""

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
                                AppStates("SCAN_UPC") = "UPC Not On System, Put aside, Next UPC|OK|"
                                CreateResponse("SCAN_UPC", "BLUE", ReceiptMessage())
                                AppStates("SCAN_UPC") = hold
                                STYLE_CODE = ""
                                COLOR_CODE = ""
                                UPC_CODE = ""
                                Exit Select
                            End If
                            UPC_CODE = CheckResponse("UPC_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            LOCATION_CODE = GetLocation(STYLE_CODE, COLOR_CODE)
                            InReceiving = GetInReceiving(STYLE_CODE, COLOR_CODE)
                            If InReceiving = 0 Then
                                CreateResponse("", "RED", "UPC Not found on Receiving")
                                Exit Select
                                'Dim row As DataRow = ASCDATA1.GetDataRow("Select CARTON_PACK_QTY, CARTONS_PER_UNIT From ICTSTYL1 where STYLE_CODE = '" & CheckResponse("STYLE_CODE") & "'")
                                'CARTON_PACK_QTY = row("CARTON_PACK_QTY")
                                'CARTONS_PER_UNIT = row("CARTONS_PER_UNIT")
                            Else
                                Dim row As DataRow = ASCDATA1.GetDataRow("Select nvl(CARTON_PACK_QTY,1) CARTON_PACK_QTY, nvl(CARTONS_PER_UNIT,1) CARTONS_PER_UNIT From ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'")
                                CARTON_PACK_QTY = row("CARTON_PACK_QTY")
                                CARTONS_PER_UNIT = row("CARTONS_PER_UNIT")
                                If G.PICK_TYPE = "D" Then
                                    CreateResponse("SCAN_CASES", "YELLOW", ReceiptMessage())
                                Else
                                    CreateResponse("SCAN_LOC", "YELLOW", ReceiptMessage())
                                End If
                            End If
                            Exit Select
                        End If

                    End If
                Case "SCAN_LOC"
                    WHSE_TRAN_TYPE = "M"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Receipt Cancelled, Scan new UPC")
                        Exit Select
                    Else
                        LOCATION_TO = SCANTEXT.ToUpper
                        LOCATION_CODE = SCANTEXT.ToUpper
                        Dim dResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, LOCATION_TO, True)
                        If dResponse.ContainsKey("Error") Then
                            CreateResponse("", "YELLOW", dResponse("Error"))
                            Exit Select
                        End If
                        dst.Tables("WHTMOVE1").Rows.Clear()
                        dst.Tables("WHTMOVE2").Rows.Clear()

                        If dResponse.ContainsKey("LOCATION_USE") Then
                            If LOCATION_USE <> dResponse("LOCATION_USE") And dResponse("LOCATION_USE") = "A" Then
                                Dim msg As String = $"Location {LOCATION_CODE}{vbCrLf}{STYLE_CODE} - {COLOR_CODE}{vbCrLf}The location selected Is Not Class III-IV,{vbCrLf}This Item requires Class III-IV"
                                CreateResponse("GET_BTNS", "", msg)
                                Exit Select
                            End If
                        End If

                        'CreateResponse("VERIFY", "B", "About to Putaway merchandise" & vbCrLf & "To " & LOCATION_TO & vbCrLf & "UPC: " & UPC_CODE)
                        CreateResponse("SCAN_CASES", "B", "Enter Qty to Putaway " & vbCrLf & "Location " & LOCATION_TO & vbCrLf & " for UPC: " & UPC_CODE & vbCrLf & "Style " & STYLE_CODE & " - " _
                                       & COLOR_CODE & " " & COLOR_DESC)
                    End If

                Case "GET_BTNS"
                    If SCANTEXT = "NEWLOC" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Move Cancelled, Scan new location")
                        Exit Select
                    ElseIf SCANTEXT = "OK" Then
                        CreateResponse("SCAN_CASES", "B", "Enter Qty to Putaway " & vbCrLf & "Location " & LOCATION_TO & vbCrLf & " for UPC: " & UPC_CODE & vbCrLf & "Style " & STYLE_CODE & " - " _
                                       & COLOR_CODE & " " & COLOR_DESC)
                        Exit Select
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
                            AppStates(AppState) = "Invalid Qty, Enter Cases |OK|"
                            CreateResponse("", "R", ReceiptMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If

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
                        If InReceiving < RcvQty Then
                            hold = AppStates("VERIFY")
                            AppStates("VERIFY") = "Possible Overage, " & hold
                            CreateResponse("VERIFY", "R", ReceiptMessage())
                            AppStates("VERIFY") = hold
                        Else
                            CreateResponse("VERIFY", "B", ReceiptMessage())
                        End If
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

                        Cases = 0
                        Units = Val(SCANTEXT)
                        RcvQty = Units
                        If InReceiving < RcvQty Then
                            hold = AppStates("VERIFY")
                            AppStates("VERIFY") = "Possible Overage, " & hold
                            CreateResponse("VERIFY", "R", ReceiptMessage())
                            AppStates("VERIFY") = hold
                        Else
                            CreateResponse("VERIFY", "B", ReceiptMessage())
                        End If
                    End If

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_UPC", "BLUE", "Scan new UPC CODE")
                    ElseIf SCANTEXT = "CANCEL" Or SCANTEXT = "N" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Putaway Cancelled, Re - scan  UPC")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub
    Function ReceiptMessage() As String
        'Dim shp_ctn As Integer
        'Dim rcv_ctn As Integer


        Dim UPC_DETAIL As String = ""
        'Dim PO_QTY_SHIP As Integer
        'Dim PO_QTY_REC As Integer

        If UPC_CODE <> "" Then

            If CARTONS_PER_UNIT > 1 Then
                UPC_DETAIL = String.Format("{0}c per unit, ", CARTONS_PER_UNIT & "".PadLeft(CARTONS_PER_UNIT), "*")
            Else

                UPC_DETAIL = String.Format("{0}u per Ctn, ", CARTON_PACK_QTY)
            End If
            'UPC_DETAIL = vbCrLf & STYLE_CODE & "-" & COLOR_CODE & " - " & UPC_DETAIL & vbCrLf & LOCATION_CODE
        End If


        Dim msg As String = appname & " Putaway To " & LOCATION_CODE & vbCrLf & "UPC: " & UPC_CODE & vbCrLf & "Style " & STYLE_CODE & " - " _
                                       & COLOR_CODE & " " & COLOR_DESC & vbCrLf
        If RcvQty > 0 Then
            msg = msg & "QTY: " & Cases & "Cs, " & Units & "u " & UPC_DETAIL
            If InReceiving < RcvQty Then
                msg = msg & vbCrLf & "Possible Overage, check " & appname
            End If
        Else
            msg = msg & UPC_DETAIL
        End If

        If G.PICK_TYPE = "D" Then
            msg = appname & " Recorded To " & LOCATION_CODE & vbCrLf & "UPC: " & UPC_CODE & vbCrLf & "Style " & STYLE_CODE & " - " _
                                       & COLOR_CODE & " " & COLOR_DESC & vbCrLf & "QTY: " & Cases & "Cs, " & Units & "u " & UPC_DETAIL
        End If

        Return msg

    End Function
    Function GetLocation(ByVal Style As String, ByVal Color As String) As String
        Dim rtn_row As DataRow

        If G.PICK_TYPE = "S" Or G.PICK_TYPE = "D" Then
            ASCMAIN1.sql = "Select * from WHTLOCB1 " & vbCrLf _
                & "     where LOCATION_CODE > '99Z'" & vbCrLf _
                & "     and LOCATION_QTY > 0" & vbCrLf _
                & "     and STYLE_CODE = :PARM1" & vbCrLf _
                & "     and COLOR_CODE = :PARM2" & vbCrLf _
                & "     order by LAST_DATE DESC"
        Else
            'this was for receiving originally view last 6 hours, but now we want to see it all
            'ASCMAIN1.sql = "Select * from WHTLOCB1 " & vbCrLf _
            '    & "     where LOCATION_CODE > '99Z'" & vbCrLf _
            '    & "     and LAST_DATE > sysdate - .25" & vbCrLf _
            '    & "     and LOCATION_QTY > 0" & vbCrLf _
            '    & "     and STYLE_CODE = :PARM1" & vbCrLf _
            '    & "     and COLOR_CODE = :PARM2" & vbCrLf _
            '    & "     order by LAST_DATE DESC"

            ASCMAIN1.sql = "Select * from WHTLOCB1 " & vbCrLf _
                & "     where LOCATION_CODE > '99Z'" & vbCrLf _
                & "     and LOCATION_QTY > 0" & vbCrLf _
                & "     and STYLE_CODE = :PARM1" & vbCrLf _
                & "     and COLOR_CODE = :PARM2" & vbCrLf _
                & "     order by LAST_DATE DESC"
        End If

        rtn_row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {Style, Color})
        'If IsNothing(rtn_row) Then
        '    ASCMAIN1.sql = "Select * From WHTPREC3, POTORDR2" & vbCrLf _
        '    & "     Where WHTPREC3.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
        '    & "     And WHTPREC3.PO_ORDER_LNO = POTORDR2.PO_ORDER_LNO " & vbCrLf _
        '    & "     And PO_SHIPMENT_NO = :PARM1" & vbCrLf _
        '    & "     and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3 "
        '    rtn_row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {PO_SHIPMENT_NO, Style, Color})
        'End If
        If Not IsNothing(rtn_row) Then
            LOCATION_CODE = rtn_row.Item("LOCATION_CODE") & ""
        Else
            LOCATION_CODE = "NO LOCATION"
            If G.PICK_TYPE = "D" Then
                LOCATION_CODE = LNFLoc
            End If
        End If
        Return LOCATION_CODE

    End Function
    Function GetInReceiving(ByVal Style As String, ByVal Color As String) As Integer
        Dim rtn_row As DataRow

        If G.PICK_TYPE = "D" Then
            Return 9999
        End If

        If G.PICK_TYPE = "S" Then
            ASCMAIN1.sql = "Select NVL(sum(RA_PUTAWAY_QTY_OPEN),0) LOCATION_QTY from SOTRMAFR " & vbCrLf _
            & "     where RA_PUTAWAY_QTY_OPEN > 0" & vbCrLf _
            & "     and STYLE_CODE = :PARM1" & vbCrLf _
            & "     and COLOR_CODE = :PARM2"
        Else
            ASCMAIN1.sql = "Select nvl(LOCATION_QTY,0) LOCATION_QTY from WHTLOCB1 " & vbCrLf _
            & "     where LOCATION_CODE = '" & FromLoc & "'" & vbCrLf _
            & "     and LOCATION_QTY > 0" & vbCrLf _
            & "     and STYLE_CODE = :PARM1" & vbCrLf _
            & "     and COLOR_CODE = :PARM2"
        End If


        rtn_row = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {Style, Color})
        If Not IsNothing(rtn_row) Then
            Return Val(rtn_row("LOCATION_QTY"))
        Else
            Return 0
        End If

    End Function
    Sub Update_Record()

        BeginTrans()

        If G.PICK_TYPE = "S" Then ' RTS logic
            Dim PUTAWAY_QTY As Integer = RcvQty

            ASCMAIN1.sql = "Select * from SOTRMAFR " & vbCrLf _
                & "     where RA_PUTAWAY_QTY_OPEN > 0" & vbCrLf _
                & "     and STYLE_CODE = '" & STYLE_CODE & "' " & vbCrLf _
                & "     and COLOR_CODE = '" & COLOR_CODE & "' "

            Fill_Records("SOTRMAFR", "", True, ASCMAIN1.sql)

            For Each row As DataRow In dst.Tables("SOTRMAFR").Select("")
                If PUTAWAY_QTY > 0 Then
                    If PUTAWAY_QTY > row("RA_PUTAWAY_QTY_OPEN") Then
                        PUTAWAY_QTY = PUTAWAY_QTY - row("RA_PUTAWAY_QTY_OPEN")
                        row("RA_PUTAWAY_QTY_OPEN") = 0
                    Else
                        row("RA_PUTAWAY_QTY_OPEN") = row("RA_PUTAWAY_QTY_OPEN") - PUTAWAY_QTY
                        PUTAWAY_QTY = 0
                    End If
                    row("LAST_OPER") = G.USER_ID
                    row("LAST_DATE") = DATETIME_STAMP
                    row("RA_PUTAWAY_LOC") = LOCATION_TO
                Else
                    Exit For
                End If
            Next
            Update_Record_TDA("SOTRMAFR")
        End If

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        With rowWHTMOVE1
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_TYPE") = WHSE_TRAN_TYPE
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

        Dim WHSE_TRAN_LNO_ctr As Integer = 0
        Dim rowWHTMOVE2 As DataRow

        'Fill_Records("WHTLOCB1", "", True, "Select * from WHTLOCB1 where WHSE_CODE = '" & G.WHSE_CODE & "' AND LOCATION_CODE = '" & LOCATION_FROM & "'" & " AND LOCATION_QTY >0")

        'For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB1").Select("")

        rowWHTMOVE2 = dst.Tables("WHTMOVE2").NewRow
        With rowWHTMOVE2
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            WHSE_TRAN_LNO_ctr += 1
            .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
            If G.PICK_TYPE = "D" Then
                .Item("LOCATION_CODE_FROM") = LOCATION_CODE
                .Item("LOCATION_CODE_TO") = DmgLoc
            Else
                .Item("LOCATION_CODE_FROM") = FromLoc
                .Item("LOCATION_CODE_TO") = LOCATION_TO
            End If
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
        ' Next


        Update_Record_TDA("WHTMOVE2")

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})


        If G.PICK_TYPE = "D" Then
            ASCMAIN1.sql = "INSERT INTO ASTNOTEM " &
                       "Select 'DAMAGES' NOTE_CODE, " &
                       "NVL((SELECT max(SEND_LNO) FROM ASTNOTEM WHERE NOTE_CODE = 'DAMAGES'), 0) + 1 SEND_LNO, " &
                       "'Damaged Item Scanned: " & STYLE_CODE & "-" & COLOR_CODE & "' NOTE_MEMO " &
                       "from DUAL"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If

        CommitTrans()
        RcvQty = 0

    End Sub


End Class
