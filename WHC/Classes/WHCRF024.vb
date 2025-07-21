Public Class WHCRF024
    ' Application Move UPC to gun - for warehouses with no LPN

    Inherits WHCRF000

    Dim BAR_CODE As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim Cases_count As Integer
    Dim TICKET_NO As String
    Dim INVALID_BAR_CODE As String
    Dim CASES_RCVD As Integer
    Dim BAR_CODE_LOCATION As String
    Dim HOLD_PROMPT As String
    Dim Mode As String


    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF024"

        AppStates.Add("SCAN_PALETTE", "Scan Palette ID|EXIT|")
        AppStates.Add("SCAN_UPC", "Scan UPC|MINUS|SHOW|DONE|") ' BLUE
        AppStates.Add("SCAN_SHOW", "Scan UPC|CANCEL|") ' BLUE
        AppStates.Add("VERIFY", "Update (Y/N)|Y|N|CANCEL|")

        AppState = "SCAN_PALETTE"
        LAST_CLR = ""

        With dst

            Create_TDA(.Tables.Add, "WHTTRAN1", "*")
            Create_TDA(.Tables.Add, "WHTTRAN2", "*")

        End With

        tbl = dst.Tables("WHTTRAN2") ' New DataTable

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
                Case "SCAN_PALETTE"
                    If SCANTEXT.Length <> 22 Then
                        CreateResponse("", "R", "Invalid Palette Barcode")
                    Else
                        Dim PICK_NO As String = SCANTEXT.Substring(1, 10)
                        Dim PALETTE_SEQ_NO As Integer = CInt(SCANTEXT.Substring(11, 3))
                        Dim PICK_NO_USL As String = SCANTEXT.Substring(14)
                        Dim SCAN_DATE As DateTime = Now

                        dst.Tables("WHTTRAN1").Rows.Clear()
                        dst.Tables("WHTTRAN2").Rows.Clear()
                        Dim rowWHTTRAN1 As DataRow = dst.Tables("WHTTRAN1").NewRow
                        rowWHTTRAN1("PALETTE_NO") = SCANTEXT
                        rowWHTTRAN1("SCAN_DATE") = SCAN_DATE
                        rowWHTTRAN1("PICK_NO") = PICK_NO
                        rowWHTTRAN1("PALETTE_SEQ_NO") = PALETTE_SEQ_NO
                        rowWHTTRAN1("PICK_NO_USL") = PICK_NO_USL
                        rowWHTTRAN1("INIT_OPER") = G.USER_ID
                        rowWHTTRAN1("INIT_DATE") = SCAN_DATE
                        dst.Tables("WHTTRAN1").Rows.Add(rowWHTTRAN1)

                    End If
                    CreateResponse("SCAN_UPC", "BLUE", "Scan UPC")

                Case "SCAN_UPC"
                    If SCANTEXT = "MINUS" Then
                        AppStates("SCAN_LOC") = "Scan Move From Location|MFG2L|Shw Loc|EXIT|"
                        Mode = SCANTEXT
                        CreateResponse("", "R", DisplayMsg("Mode was Changed"))
                        Exit Select
                    ElseIf SCANTEXT = "DONE" Then
                        AppStates("SCAN_LOC") = "Scan Move To Location|M2G|Dep WH|DepEcom|EXIT|"
                        Mode = SCANTEXT
                        CreateResponse("", "R", DisplayMsg("Mode was Changed"))
                        Exit Select
                    ElseIf SCANTEXT = "SHOW" Then
                        CreateResponse("SCAN_SHOW", "BLUE", DisplayMsg("Scan UPC to Show Locations"))
                        Exit Select
                    ElseIf SCANTEXT.Length <> 12 Then
                        CreateResponse("", "R", "Invalid Carton Barcode")
                        Exit Select
                    Else
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                            Exit Select
                        End If

                        Dim Styles As String = CheckResponse("Stylelist")

                        dst.Tables("WHTMOVE1").Rows.Clear()
                        dst.Tables("WHTMOVE2").Rows.Clear()

                        CreateResponse("SCAN_UPC", "BLUE", DisplayMsg("Styles in Location: " & Styles))
                    End If

                Case "SCAN_SHOW"
                    If SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", "Show Cancelled, Scan  location")
                        Exit Select
                    Else



                    End If
                    CreateResponse("SCAN_LOC", "YELLOW", "Show Cancelled, Scan  location")

                Case "SCAN_UPC"
                    If SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", "Move Cancelled, Scan  location")
                        Exit Select
                    Else
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                            Exit Select
                        End If

                        If CheckResponse.ContainsKey("UPC_CODE") Then
                            UPC_CODE = CheckResponse("UPC_CODE")
                            STYLE_CODE = CheckResponse("STYLE_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            CreateResponse("SCAN_CASES", "G", DisplayMsg("UPC " & UPC_CODE & " selected"))
                            Exit Select
                        End If
                        STYLE_CODE = SCANTEXT.ToUpper


                    End If


                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg("Move Cancelled, Scan  location"))
                        Exit Select
                    Else
                        If SCANTEXT = "0" Then

                            CreateResponse("", "B", DisplayMsg("All colors " & colors))
                            Exit Select
                        Else
                            'Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyColor(Me, SCANTEXT, STYLE_CODE, COLOR_CODEs)
                            'If CheckResponse.ContainsKey("Error") Then
                            '    CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                            '    Exit Select
                            'End If
                        End If
                    End If
                    CreateResponse("SCAN_CASES", "G", DisplayMsg("UPC " & UPC_CODE & " selected"))

                Case "SCAN_CASES"
                    If SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg("Move Cancelled, Scan  location"))
                        Exit Select
                    ElseIf SCANTEXT = "OK" Then
                        CreateResponse("VERIFY", "B", DisplayMsg(HOLD_PROMPT))
                        Exit Select
                    ElseIf SCANTEXT = "NEW_QTY" Then
                        CreateResponse("", "G", DisplayMsg("Re-enter Qty for " & UPC_CODE))
                        Exit Select
                    ElseIf SCANTEXT.Contains("*") Then
                        Dim S() As String
                        S = SCANTEXT.Split("*")
                        CASES_RCVD = Val(S(0))

                    Else
                        CASES_RCVD = Val(SCANTEXT)

                    End If

                    If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                        CreateResponse("", "R", DisplayMsg("Invalid Number Of cases " & SCANTEXT))
                        Exit Select
                    End If
                    'CASES_MOVED = Val(SCANTEXT)


                    'ASCMAIN1.sql = "Select CARTON_PACK_QTY from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'"
                    'Dim CARTON_PACK_QTY As String = ASCDATA1.GetDataValue
                    'UNITS_MOVED = UNITS_MOVED + Val(CARTON_PACK_QTY) * CASES_RCVD

                    'ASCMAIN1.sql = "Select LOCATION_QTY from WHTLOCB1 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & G.WHSE_CODE & "' and LOCATION_CODE = '" & LOCATION_CODE & "'"
                    'Dim LOC_QTY As String = ASCDATA1.GetDataValue
                    'If UNITS_MOVED > Val(LOC_QTY & "") And Mode = "M2G" Then
                    '    'Warning
                    '    Dim hold As String = AppStates(AppState)
                    '    AppStates(AppState) = "Warning About to create Negative O/H|OK|NEW_QTY|"
                    '    CreateResponse("", "RED", DisplayMsg("Trying to move " & UNITS_MOVED & " - ONLY " & LOC_QTY & " O/H"))
                    '    AppStates(AppState) = hold
                    '    HOLD_PROMPT = "Entered " & SCANTEXT & " Cases, Carton pack: " & CARTON_PACK_QTY & ", Units to move: " & UNITS_MOVED
                    '    Exit Select
                    'End If

                    'CreateResponse("VERIFY", "B", DisplayMsg("Entered " & SCANTEXT & " Cases, Carton pack: " & CARTON_PACK_QTY & ", Units to move: " & UNITS_MOVED))
                    'Exit Select

                Case "SCAN_UNITS"
                    If SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg("Move Cancelled, Scan  location"))
                        Exit Select
                    ElseIf SCANTEXT = "OK" Then
                        CreateResponse("VERIFY", "B", DisplayMsg(HOLD_PROMPT))
                        Exit Select
                    ElseIf SCANTEXT = "NEW_QTY" Then
                        CreateResponse("", "G", DisplayMsg("Re-enter Qty for " & UPC_CODE))
                        Exit Select
                    Else
                        'Can we have more than 99 loose units to count in a location?
                        If Val(SCANTEXT) > 99 Or Val(SCANTEXT) < 1 Then
                            CreateResponse("", "R", DisplayMsg("Invalid Number Of units " & SCANTEXT))
                            Exit Select
                        End If
                        'UNITS_MOVED = Val(SCANTEXT)

                        'ASCMAIN1.sql = "Select LOCATION_QTY from WHTLOCB1 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & G.WHSE_CODE & "' and LOCATION_CODE = '" & LOCATION_CODE & "'"
                        'Dim LOC_QTY As String = ASCDATA1.GetDataValue
                        'If UNITS_MOVED > Val(LOC_QTY & "") And Mode = "M2G" Then
                        '    'Warning
                        '    Dim hold As String = AppStates(AppState)
                        '    AppStates(AppState) = "Warning About to create Negative O/H|OK|NEW_QTY|"
                        '    CreateResponse("", "RED", DisplayMsg("Trying to move " & UNITS_MOVED & " - ONLY " & LOC_QTY & " O/H"))
                        '    AppStates(AppState) = hold
                        '    HOLD_PROMPT = "Entered " & SCANTEXT & "Scanned " & SCANTEXT & " Units to Move"
                        '    Exit Select
                        'End If

                        CreateResponse("VERIFY", "B", DisplayMsg("Scanned " & SCANTEXT & " Units to Move"))
                    End If

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg(""))
                    ElseIf SCANTEXT = "N" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg("Move Cancelled, Re - scan  location"))
                    ElseIf SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg("Move Cancelled, Re - scan  location"))
                    Else
                        CreateResponse("", "R", DisplayMsg("Invalid Response"))
                    End If
            End Select
        End If
    End Sub


    Function DisplayMsg(ByVal note As String) As String
        Dim msg As String = ""
        Dim FromLoc As String = ""

        If Mode = "M2G" Then
            msg = "Move to Gun"
            'FromLoc = LOCATION_CODE
        Else
            msg = "Move to Location"
            FromLoc = G.GUN_LOC
        End If
        'If LOCATION_CODE <> "" Then
        '    msg = msg & vbCrLf & "LOC: " & LOCATION_CODE
        '    If STYLE_CODE <> "" Then
        '        msg = msg & vbCrLf & STYLE_CODE
        '        If COLOR_CODE <> "" Then
        '            msg = msg & " " & COLOR_CODE
        '            ASCMAIN1.sql = "Select LOCATION_QTY from WHTLOCB1 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & G.WHSE_CODE & "' and LOCATION_CODE = '" & FromLoc & "'"
        '            Dim LOC_QTY As String = ASCDATA1.GetDataValue
        '            msg = msg & vbCrLf & "Found: " & FromLoc & " QTY " & ": " & LOC_QTY
        '        End If
        '    End If
        'Else
        '    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, G.GUN_LOC, False)
        '    If CheckResponse.ContainsKey("Stylelist") Then
        '        Dim Styles As String = CheckResponse("Stylelist")
        '        If Styles.Length > 2 Then
        '            msg = msg & vbCrLf & "Styles in gun " & Styles
        '        End If
        '    End If
        'End If

        If note <> "" Then
            msg = msg & vbCrLf & note
        End If
        Return msg
    End Function

    Sub ClearScanner()
        UPC_CODE = ""
    End Sub

    Sub Update_Record()

        Dim FromLoc As String = ""
        Dim ToLoc As String = ""

        'If Mode = "M2G" Then
        '    FromLoc = LOCATION_CODE
        '    ToLoc = G.GUN_LOC
        'Else 'If Mode = "MFG2L" Then
        '    FromLoc = G.GUN_LOC
        '    ToLoc = LOCATION_CODE
        'End If

        BeginTrans()

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

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

        Dim WHSE_TRAN_LNO_ctr As Integer = 0

        Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
        With rowWHTMOVE2
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            WHSE_TRAN_LNO_ctr += 1
            .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
            .Item("LOCATION_CODE_FROM") = FromLoc
            .Item("LOCATION_CODE_TO") = ToLoc
            .Item("BAR_CODE") = "0000000000"
            '.Item("WHSE_TRAN_QTY") = UNITS_MOVED
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

        ClearScanner()
        CommitTrans()
    End Sub


End Class
