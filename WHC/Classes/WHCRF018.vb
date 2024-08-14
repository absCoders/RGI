Public Class WHCRF018
    ' Application RMA processing for non LPN Warehouse

    Inherits WHCRF000


    Dim WHSE_LOC_RTN As String
    Dim WHSE_LOC_DST As String
    Dim RA_NO As String
    Dim RA_RTN_LNO As Integer
    Dim STYLE_CODE As String
    Dim STYLE_DESC As String
    Dim COLOR_CODE As String
    Dim COLOR_DESC As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim GOTO_LOCATION As String
    Dim LOCATIONS As String
    Dim RA_RTN_STATUS As String
    Dim COLOR_CODEs As New List(Of String)
    'Dim Cases_count As Integer
    Dim CASES_BOOK As Integer
    Dim CASES_MOVED As Integer
    Dim UNITS_MOVED As Integer
    Dim TICKET_NO1 As String
    'Dim ORIGINAL_LOCATION As String
    '    Dim colors As String = ""
    Dim ScanLocMsg As String
    'Dim RowNum As Integer
    Dim VoidList() As String
    Dim VoidTranNo As String
    Dim INNER_PACK_QTY As Integer
    Dim CARTON_PACK_QTY As Integer
    Dim CARTONS_PER_UNIT As Integer
    Dim RA_RTN_QTY As Integer
    Dim holdScan As String
    Dim _sql As String
    Dim rowSOTRMAFR As DataRow
    'Dim PickSplit As String
    Dim PageNo As Integer = 0
    Dim LockWarning As Boolean

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF018"

        AppStates.Add("SCAN_RMA", "Scan RMA Num|EXIT|") 'Green - find RMA get invoice no from (SOTRMAF1)
        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style|DONE|VIEW|") 'Blue
        AppStates.Add("SCAN_COLOR", "Enter Color code|CANCEL|")
        AppStates.Add("RTS_ITEM", "Return Item to Stock|Y|N|") ' Return to stock is not broken Y=1,N=3
        AppStates.Add("NOT_FOUND", "Warning Item not in RMA Accept?|Y|N|") ' check SOTRMAF2
        AppStates.Add("SCAN_CASES", "How many cases returned, (0 for units)|CANCEL|")
        AppStates.Add("SCAN_UNITS", "How many units returned|CANCEL|")
        ' AppStates.Add("SCAN_LOC", "Scan Loc, 00:'{0}'") 'Yellow
        AppStates.Add("SCAN_LOC", "Scan Loc, 00:'{0}'|KP RTS|") 'Yellow
        AppStates.Add("VERIFY", "Update|Y|N|CLEAR|EXIT|")
        AppStates.Add("LEAVE", "Units in Gun, Gun must be Empty for Pick|EXIT|")
        AppStates.Add("CLOSE_RMA", "Close RMA|Y|N|")
        AppStates.Add("SCAN_VOID", "Void Line or 'ALL'|BACK|NXT PAGE|EXIT|")

        'Get application state ready
        AppState = "SCAN_RMA"
        LAST_CLR = "GREEN"
        ScanLocMsg = AppStates("SCAN_LOC")

        With dst
            Create_TDA(.Tables.Add, "SOTRMAFR", "*", 1)
            'Lock RMA, IF RA_QTY_USED > 0 THEN ITEM PROCESSED, DON'T ALLOW MODIFY
            'LNO GET MAX FROM RMA
            Create_TDA(.Tables.Add, "SOTRMAF1", "*", 1)
            Create_TDA(.Tables.Add, "SOTRMAF2", "*", 1)
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")

        End With

        Dim ROW As DataRow = Nothing
        ASCMAIN1.sql = "SELECT * FROM ICTWHSE1 WHERE WHSE_CODE = '" & g.WHSE_CODE & "'"
        ROW = ASCDATA1.GetDataRow(ASCMAIN1.sql)
        WHSE_LOC_RTN = ROW.Item("WHSE_LOC_RTN")
        WHSE_LOC_DST = ROW.Item("WHSE_LOC_DST")

        Dim styles As String = TACMAIN1.LookupLocation(Me, g.GUN_LOC)
        If styles.Length > 0 Then
            AppState = "LEAVE"
        End If

        tbl = dst.Tables("SOTRMAFR")



    End Sub

    Public Overrides Function Hello() As String

        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        RESPONSE &= vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(ByVal SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        Dim hold As String = String.Empty

        If SCANTEXT = "EXIT" Then
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "R", "EXIT")
        ElseIf SCANTEXT = "OK" Then
            CreateResponse("", "B", "")
        Else
            Select Case AppState
                Case "SCAN_RMA"
                    Dim styles As String = TACMAIN1.LookupLocation(Me, G.GUN_LOC)
                    If styles.Length > 0 Then
                        CreateResponse("LEAVE", "R", "Gun has Units, Deposit merchandise to continue")
                        Exit Select
                    End If
                    If SCANTEXT.Length < 6 Then
                        SCANTEXT = SCANTEXT.PadLeft(6, "0")
                    End If
                    ASCMAIN1.Multi_Task_Cleanup()

                    Fill_Records("SOTRMAF1", SCANTEXT, True)
                    Dim rows1() As DataRow = dst.Tables("SOTRMAF1").Select("")
                    If rows1.Length <> 1 Then
                        CreateResponse("", "GREEN", "Invalid RA Number, Re-Scan")

                    Else
                        RA_NO = SCANTEXT
                        If Not ASCMAIN1.Logical_Lock("SOTRMAF1", SCANTEXT) Then
                            Dim User As String = ASCMAIN1.MultiTask_Get_Users("SOTRMAF1", RA_NO, "L")
                            If (Not User.Contains(G.USER_ID)) Or LockWarning = False Then
                                EMsg = vbCr & "Return Ticket is Locked " _
                                         & User
                                CreateResponse("", "GREEN", EMsg)
                                If User.Contains(G.USER_ID) Then LockWarning = True
                                Exit Select
                            End If
                        End If
                        If rows1(0)("RA_STATUS") & "" <> "O" Then
                            CreateResponse("", "GREEN", "RA " & SCANTEXT & " Not Open, Please Verify RA")
                            Exit Select
                        End If

                    End If
                    Fill_Records("SOTRMAF2", SCANTEXT, True)

                    Fill_Records("SOTRMAFR", RA_NO, True)
                    RA_RTN_LNO = 0
                    If tbl.Rows.Count > 0 Then
                        RA_RTN_LNO = tbl.Compute("MAX(RA_RTN_LNO)", "")
                    End If

                    CreateResponse("SCAN_UPC", "BLUE", RAMessage())
                    InitItem()

                Case "SCAN_UPC"
                    If SCANTEXT = "DONE" Then
                        CreateResponse("CLOSE_RMA", "G", "Update RA Ticket for Processing")
                        Exit Select
                    ElseIf SCANTEXT = "VIEW" Then
                        ' View Scanned
                        showreturned("")
                        Exit Select
                    Else
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("UPC_CODE") Then
                            UPC_CODE = CheckResponse("UPC_CODE")
                            STYLE_CODE = CheckResponse("STYLE_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            STYLE_DESC = CheckResponse("STYLE_DESC")
                            COLOR_DESC = CheckResponse("COLOR_DESC")
                            CARTONS_PER_UNIT = CheckResponse("CARTONS_PER_UNIT")
                            CARTON_PACK_QTY = CheckResponse("CARTON_PACK_QTY")
                            INNER_PACK_QTY = CheckResponse("INNER_PACK_QTY")
                            If Not RMA_ITEM() Then
                                CreateResponse("NOT_FOUND", "G", RAMessage)
                                Exit Select
                            End If
                            CreateResponse("RTS_ITEM", "G", RAMessage)
                            Exit Select
                        Else
                            If CheckResponse.ContainsKey("STYLE_CODE") Then
                                'I'm looking for one color why give a choice? do I just ask for Yes to accept?
                                'TACMAIN1.GetColors(Me, STYLE_CODE, ORIGINAL_LOCATION, COLOR_CODEs, colors)
                                'CreateResponse("SCAN_COLOR", "G", "Style " & STYLE_CODE & " has been selected, colors " & colors)
                                CreateResponse("SCAN_COLOR", "G", RAMessage())
                                STYLE_CODE = CheckResponse("STYLE_CODE")
                                Exit Select
                            End If
                        End If
                    End If
                    'Error
                    hold = AppStates(AppState)
                    AppStates(AppState) = "Wrong UPC/Style, try again|OK|"
                    CreateResponse("", "BLUE", RAMessage())
                    AppStates(AppState) = hold

                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", RAMessage)
                        Exit Select
                    Else
                        If SCANTEXT = "0" Then
                            TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)
                            CreateResponse("", "B", "Style " & STYLE_CODE & " has been selected, All colors " & colors)
                            Exit Select
                        Else
                            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyColor(Me, SCANTEXT, STYLE_CODE, COLOR_CODEs)
                            If CheckResponse.ContainsKey("Error") Then
                                CreateResponse("", "R", CheckResponse("Error"))
                                Exit Select
                            End If
                            UPC_CODE = CheckResponse("UPC_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            STYLE_DESC = CheckResponse("STYLE_DESC")
                            COLOR_DESC = CheckResponse("COLOR_DESC")
                            CARTONS_PER_UNIT = CheckResponse("CARTONS_PER_UNIT")
                            CARTON_PACK_QTY = CheckResponse("CARTON_PACK_QTY")
                            INNER_PACK_QTY = CheckResponse("INNER_PACK_QTY")
                            If Not RMA_ITEM() Then
                                CreateResponse("NOT_FOUND", "G", RAMessage)
                                Exit Select
                            End If
                            CreateResponse("RTS_ITEM", "G", RAMessage)
                            Exit Select
                        End If
                    End If
                    'Error
                    hold = AppStates(AppState)
                    AppStates(AppState) = "Wrong UPC/Style, try again|OK|"
                    CreateResponse("", "BLUE", RAMessage())
                    AppStates(AppState) = hold

                Case "NOT_FOUND"
                    If SCANTEXT = "N" Then
                        CreateResponse("SCAN_UPC", "BLUE", STYLE_CODE & " - " & COLOR_CODE & " Will not be processed, Scan Next UPC")
                    ElseIf SCANTEXT = "Y" Then
                        CreateResponse("RTS_ITEM", "G", RAMessage)
                    End If

                Case "RTS_ITEM"
                    If SCANTEXT = "Y" Or SCANTEXT = "N" Then
                        'Found Item lets get locations
                        RA_RTN_STATUS = If(SCANTEXT = "Y", "1", "3")
                        GetLocation()
                        'CreateResponse("SCAN_CASES", "YELLOW", RAMessage())
                        CreateResponse("SCAN_UNITS", "YELLOW", RAMessage())
                    End If

                Case "SCAN_CASES"
                    '  Asking for Cases although Message may not display number of cases requested
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", STYLE_CODE & " - " & COLOR_CODE & " Will not be processed, Scan Next UPC")
                        Exit Select
                    Else
                        If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Number, Enter Cases returned|OK|"
                            CreateResponse("", "R", RAMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        If SCANTEXT.Contains("*") Then
                            Dim S() As String
                            S = SCANTEXT.Split("*")
                            CASES_MOVED = Val(S(0))
                            UNITS_MOVED = Val(S(1))
                        Else
                            CASES_MOVED = Val(SCANTEXT)
                            UNITS_MOVED = 0
                        End If

                        If CASES_MOVED + UNITS_MOVED = 0 Then
                            CreateResponse("SCAN_UNITS", "G", RAMessage())
                            Exit Select
                        End If
                        'If RA_RTN_STATUS = "1" Then
                        '    AppStates("SCAN_LOC") = String.Format(ScanLocMsg, GOTO_LOCATION)
                        '    CreateResponse("SCAN_LOC", "YELLOW", RAMessage)
                        'Else
                        '    CreateResponse("VERIFY", "B", RAMessage)
                        'End If

                        CreateResponse("VERIFY", "B", RAMessage)
                    End If

                Case "SCAN_UNITS"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", STYLE_CODE & " - " & COLOR_CODE & " Will not be processed, Scan Next UPC")
                        Exit Select
                    Else
                        'Can we have more than 999 loose units to Move in a Pick?
                        If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Count " & SCANTEXT & ", Verify Units Count|OK|"
                            CreateResponse("", "R", RAMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        UNITS_MOVED = Val(SCANTEXT)
                        'If RA_RTN_STATUS = "1" Then
                        '    AppStates("SCAN_LOC") = String.Format(ScanLocMsg, GOTO_LOCATION)
                        '    CreateResponse("SCAN_LOC", "YELLOW", RAMessage)
                        'Else
                        '    CreateResponse("VERIFY", "B", RAMessage)
                        'End If
                        CreateResponse("VERIFY", "B", RAMessage)
                    End If

                Case "SCAN_LOC"
                    ' Current location selected
                    If SCANTEXT = "KP RTS" Then
                        SCANTEXT = WHSE_LOC_RTN
                    End If
                    If SCANTEXT = "00" Then
                        SCANTEXT = GOTO_LOCATION
                    End If
                    Dim dResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                    If dResponse.ContainsKey("Error") Then
                        hold = AppStates(AppState)
                        AppStates(AppState) = "Invalid Location, Please re-Scan|OK|"
                        CreateResponse("", "YELLOW", dResponse("Error"))
                        AppStates(AppState) = hold
                        Exit Select
                    End If
                    LOCATION_CODE = SCANTEXT.ToUpper

                    hold = AppStates("VERIFY")
                    'AppStates("VERIFY") = "Update " & SCANTEXT & " Cases|Y|N|CLEAR|EXIT|"

                    AppStates("VERIFY") = "Update |Y|N|CLEAR|EXIT|"
                    CreateResponse("VERIFY", "B", RAMessage())
                    AppStates("VERIFY") = hold

                Case "SCAN_VOID"
                    If SCANTEXT = "BACK" Then
                        '
                        InitItem()
                        CreateResponse("SCAN_UPC", "BLUE", RAMessage())
                        Exit Select
                    ElseIf SCANTEXT = "NXT PAGE" Then
                        showreturned("", PageNo)
                        'CreateResponse("SCAN_PTCKT", "GREEN", PickMessage())
                        Exit Select
                    ElseIf SCANTEXT = "YES" Then
                        voidLine(holdScan)
                        holdScan = ""
                        showreturned("")
                        Exit Select
                    ElseIf SCANTEXT = "NO" Then
                        showreturned("")
                        holdScan = ""
                        Exit Select
                    Else
                        'handle voids
                        If holdScan = "" Then
                            If SCANTEXT.ToUpper = "ALL" Then
                                holdScan = SCANTEXT.ToUpper
                                showreturned(holdScan)
                                Exit Select
                            End If
                            If Val(SCANTEXT) > (VoidList.Length - 1) Or Val(SCANTEXT) < 1 Then
                                'error
                            Else
                                holdScan = SCANTEXT
                                showreturned(SCANTEXT, PageNo - 1)
                                Exit Select
                            End If
                        Else
                            showreturned(holdScan, PageNo - 1)
                            Exit Select
                        End If
                    End If
                    showreturned("")


                Case "VERIFY"
                    If SCANTEXT.ToUpper = "Y" Then
                        ' Y is for normal Verify
                        Update_Record()
                        CreateResponse("SCAN_UPC", "BLUE", RAMessage())
                    ElseIf SCANTEXT = "YES" Then
                        ' YES is for Clear Verify
                        InitItem()
                        CreateResponse("SCAN_UPC", "BLUE", RAMessage())
                    ElseIf SCANTEXT.ToUpper = "N" Or SCANTEXT = "NO" Then
                        CreateResponse("RTS_ITEM", "G", RAMessage())
                    ElseIf SCANTEXT = "CLEAR" Then
                        hold = AppStates(AppState)
                        AppStates(AppState) = "Clear Current UPC |YES|NO|"
                        CreateResponse("", "R", RAMessage())
                        AppStates(AppState) = hold
                        Exit Select
                    Else
                        'Error
                        hold = AppStates(AppState)
                        AppStates(AppState) = "Invalid Response|OK|"
                        CreateResponse("", "R", RAMessage())
                        AppStates(AppState) = hold

                    End If

                Case "CLOSE_RMA"
                    If SCANTEXT.ToUpper = "Y" Then
                        UPDATE_RA_STATUS()
                    End If
                    CreateResponse("SCAN_RMA", "GREEN", "Scan New RMA Number")
            End Select
        End If
    End Sub

    Sub GetLocation()
        LOCATIONS = ""
        GOTO_LOCATION = ""
        ASCMAIN1.sql = " select b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE " & vbCrLf _
            & " from whtlocb1 b1 " & vbCrLf _
            & "  join whtlocm1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE " & vbCrLf _
            & "  where b1.STYLE_CODE = '" & STYLE_CODE & "' and b1.COLOR_CODE = '" & COLOR_CODE & "' " & vbCrLf _
            & "  and  nvl(m1.LOCATION_USE,'A') = 'A' " & vbCrLf _
            & "  and m1.WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
            & "  order by b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"
        Dim rc As Integer = 0
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            rc = rc + 1
            If rc > 5 Then
                Exit For
            End If
            If rc = 1 Then
                GOTO_LOCATION = row.Item("LOCATION_CODE")
                LOCATIONS = GOTO_LOCATION
            Else
                LOCATIONS = LOCATIONS & "," & row.Item("LOCATION_CODE")
            End If
        Next

        If GOTO_LOCATION = "" Then
            GOTO_LOCATION = WHSE_LOC_RTN
        End If

        LOCATION_CODE = WHSE_LOC_RTN ' for now are goto location is always SET TO WHSE_LOC_RTN 
        ' LOCATION_CODE should be set from propmt - if we put back to single step

    End Sub

    Function RMA_ITEM() As Boolean
        Dim Count As Integer
        Count = dst.Tables("SOTRMAF2").Compute("Count(STYLE_CODE)", String.Format("STYLE_CODE = '{0}' and COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE))
        RA_RTN_QTY = dst.Tables("SOTRMAF2").Compute("Sum(RA_QTY)", String.Format("STYLE_CODE = '{0}' and COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE))
        If Count = 0 Then
            Return False
        End If
        Return True
    End Function

    Function RAMessage() As String

        Dim RA_MESSAGE As String = "RA " & RA_NO

        Dim RtnQty As String = ""
        Dim RA_PUTAWAY_QTY_OPEN = (CASES_MOVED * CARTON_PACK_QTY + UNITS_MOVED)


        If STYLE_CODE <> "" Then
            RA_MESSAGE = RA_MESSAGE _
            & " Style " & STYLE_CODE _
            & " - " & STYLE_DESC & vbCrLf _
            & "Color " & COLOR_DESC & vbCrLf
        End If


        If New String() {"SCAN_LOC", "VERIFY"}.Contains(AppState) Then
            If CARTONS_PER_UNIT > 0 Then
                RtnQty = RtnQty & String.Format(" {0}c per unit", CARTONS_PER_UNIT)
            Else
                If CARTON_PACK_QTY > 0 Then
                    RtnQty = RtnQty & String.Format(" {0}c", CARTON_PACK_QTY)
                End If
                If INNER_PACK_QTY > 0 Then
                    RtnQty = RtnQty & String.Format(" {0}i", INNER_PACK_QTY)
                End If
            End If
        End If

        If AppState = "VERIFY" Then
            'Show location and 
        End If

        If Val(RA_PUTAWAY_QTY_OPEN & "") <> 0 Then
            RA_MESSAGE = RA_MESSAGE _
           & " " & RA_PUTAWAY_QTY_OPEN & " Units "
        End If

        If RA_RTN_STATUS = "3" Then
            RA_MESSAGE = RA_MESSAGE _
                & "To be Destroyed " & vbCrLf
        ElseIf RA_RTN_STATUS = "1" Then
            RA_MESSAGE = RA_MESSAGE _
                & "Return to Stock " & vbCrLf
        End If

        If LOCATION_CODE <> "" Then
            RA_MESSAGE = RA_MESSAGE _
               & "To Location " & LOCATION_CODE
        End If

        'RA_MESSAGE = +" Style " & STYLE_CODE _
        '    & " - " & STYLE_DESC & vbCrLf _
        '    & "Color " & COLOR_DESC & vbCrLf _
        '    & RtnQty

        Return RA_MESSAGE
        '& "Pick Location " & GOTO_LOCATION & vbCrLf _
        '& "Pick Qty " & PICK_QTY_OPEN & PickQty & vbCrLf & "".PadRight(60)
    End Function

    Sub showreturned(ByVal showline As String, Optional ByVal PAGE As Int16 = 0)
        Dim msg As String = "returned for " & RA_NO
        Dim lno As Integer = 0
        Dim lines As String = ""
        Dim pageLno As Integer = 0

        Dim rows() As DataRow = tbl.Select("GUN_STATUS = 'P'", "RA_RTN_LNO")
        If PAGE * 5 > rows.Length Or PAGE < 0 Then PAGE = 0

        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                If pageLno / 5 >= PAGE Then
                    lno += 1
                    If showline = "" Or showline = "ALL" Or Val(showline) = lno Then
                        msg = msg & vbCrLf & "L" & lno & " " & ROW.Item("STYLE_CODE") & " " & ROW.Item("COLOR_CODE") & " " _
                           & ROW.Item("RA_PUTAWAY_QTY_OPEN") & "u " & IIf(ROW.Item("RA_RTN_STATUS") = "1", "", "D")
                    End If
                    lines &= "|" & ROW.Item("RA_RTN_LNO") & ":" & ROW.Item("RA_UPC_CODE")
                    If lno = 5 Then Exit For
                End If
                pageLno += 1
            Next
            VoidList = lines.Split("|")
            If showline <> "" Then
                Dim hold As String = AppStates(AppState)
                AppStates(AppState) = If(showline = "ALL", "Void ALL returned lines", "Void this line") & "|YES|NO|"
                CreateResponse("", "R", msg)
                AppStates(AppState) = hold
            Else
                CreateResponse("SCAN_VOID", "R", msg)
            End If
        Else
            msg &= vbCrLf & "No Lines to Void" & vbCrLf & "Back to continue"
            VoidList = {}
            Dim hold As String = AppStates("SCAN_VOID")
            AppStates("SCAN_VOID") = "No Returns, CLICK BELOW|BACK|NEW RA|EXIT|"
            CreateResponse("SCAN_VOID", "R", msg)
            AppStates("SCAN_VOID") = hold
        End If
        If showline = "" Then
            PageNo = PAGE + 1
        End If


    End Sub

    Sub voidLine(ByRef SCANTEXT As String)
        Dim rows() As DataRow
        Dim ss As String
        Dim RA_RTN_LNO As String
        Dim UPC_CODE As String
        If SCANTEXT = "ALL" Then
            ss = "GUN_STATUS = 'P'"
        Else
            Dim hold = VoidList(Val(SCANTEXT))
            Dim void = hold.split(":")
            RA_RTN_LNO = void(0)
            UPC_CODE = void(1)
            ss = "RA_RTN_LNO = '" & void(0) & "' and RA_UPC_CODE = '" & void(1) & "'"
        End If
        rows = tbl.Select(ss)
        If rows.Length = 0 Then
            'error
        Else
            For Each rowSOTRMAFR As DataRow In rows
                'Figure out multipicks for SOTPICK5
                Dim row As DataRow = tbl.Select("RA_UPC_CODE = '" & rowSOTRMAFR.Item("RA_UPC_CODE") & "'").First    '.Tables("SOTPICK5").Select("PICK_STATUS = 'P' AND PICK_LNO = '" & voidLine(0) & "' and UPC_CODE = '" & Void(1) & "'").FirstOrDefault
                rowSOTRMAFR.Item("GUN_STATUS") = "V"

                'OpenPicks = OpenPicks + 1
                'rowWHTSCANS = row
                'loadLine()
                'Update_Record()
            Next
        End If
        BeginTrans()
        Update_Record_TDA("SOTRMAFR")
        CommitTrans()

    End Sub

    Sub InitItem()
        STYLE_CODE = ""
        COLOR_CODE = ""
        UPC_CODE = ""
        CASES_MOVED = 0
        UNITS_MOVED = 0
        RA_RTN_STATUS = ""
        'extra vars
        GOTO_LOCATION = ""
        LOCATION_CODE = ""
        CARTONS_PER_UNIT = 0
        CARTON_PACK_QTY = 0
        INNER_PACK_QTY = 0
        STYLE_DESC = ""
    End Sub

    Sub Update_Record()

        Dim RA_PUTAWAY_QTY_OPEN = (CASES_MOVED * CARTON_PACK_QTY + UNITS_MOVED)

        RA_RTN_LNO = RA_RTN_LNO + 1

        BeginTrans()

        rowSOTRMAFR = dst.Tables("SOTRMAFR").NewRow
        With rowSOTRMAFR
            .Item("RA_NO") = RA_NO
            .Item("RA_RTN_LNO") = RA_RTN_LNO
            .Item("RA_RTN_QTY") = RA_RTN_QTY
            .Item("RA_UPC_CODE") = UPC_CODE
            .Item("RA_QTY_USED") = 0
            .Item("RA_PUTAWAY_QTY_OPEN") = RA_PUTAWAY_QTY_OPEN
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("RA_RTN_STATUS") = RA_RTN_STATUS
            .Item("GUN_STATUS") = "P"
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
        End With
        dst.Tables("SOTRMAFR").Rows.Add(rowSOTRMAFR)

        Update_Record_TDA("SOTRMAFR")

        'lets not accidentally do a move - lets discuss and make sure we want to do this.
        If RA_RTN_STATUS = "1" And WHSE_LOC_RTN <> LOCATION_CODE And 1 = 2 Then
            Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
            Dim WHSE_TRAN_LNO_ctr As Integer = 1

            Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
            With rowWHTMOVE1
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("WHSE_TRAN_TYPE") = "M"
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
                .Item("LOCATION_CODE_FROM") = WHSE_LOC_RTN
                .Item("LOCATION_CODE_TO") = LOCATION_CODE
                .Item("BAR_CODE") = "0000000000"
                .Item("WHSE_TRAN_QTY") = RA_RTN_QTY
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

            ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                           New Object() {WHSE_TRAN_NO, 0, 1},
                           New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})
        End If

        CommitTrans()
        InitItem()

    End Sub

    Sub UPDATE_RA_STATUS()
        BeginTrans()

        ASCMAIN1.sql = "UPDATE SOTRMAFR SET GUN_STATUS = 'F' WHERE RA_NO = '" & RA_NO & "' and GUN_STATUS = 'P'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        CommitTrans()
    End Sub


End Class
