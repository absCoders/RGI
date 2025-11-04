Public Class WHCRF011
    ' Application Move UPC to gun - for warehouses with no LPN

    Inherits WHCRF000

    Dim BAR_CODE As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim COLOR_CODEs As New List(Of String)
    Dim Cases_count As Integer
    Dim TICKET_NO As String
    Dim INVALID_BAR_CODE As String
    Dim CASES_BOOK As Integer
    Dim CASES_MOVED As Integer
    Dim UNITS_MOVED As Integer
    Dim TICKET_NO1 As String
    Dim BAR_CODE_LOCATION As String
    Dim HOLD_PROMPT As String
    Dim Mode As String
    Dim colors As String = ""
    Dim page As Int16 = 0
    Dim tblPage As DataTable = Nothing

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF011"

        AppStates.Add("GET_MODE", "Select Move direction|M2G|MFG2L|EXIT|")
        AppStates.Add("SCAN_LOC", "Scan Move From Location|MFG2L|EXIT|") ' YELLOW
        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style |CANCEL|") ' BLUE
        AppStates.Add("SCAN_COLOR", "Select a Color from List |CANCEL|")
        AppStates.Add("SCAN_CASES", "How many cases, (0 for units) |CANCEL|")
        AppStates.Add("SCAN_UNITS", "How many units |CANCEL|")
        AppStates.Add("SCAN_SHOW", "Scan UPC|BACK|<<|>>|") ' BLUE
        AppStates.Add("VERIFY", "Update (Y/N)|Y|N|CANCEL|")

        AppState = "GET_MODE"
        LAST_CLR = ""

        With dst
            '  With .Tables.Add("WHTSCANS").Columns
            ' .Add("BAR_CODE")
            ' .Add("NEW")
            ' .Add("SCANNED")
            ' End With
            ' .Tables("WHTSCANS").PrimaryKey = New DataColumn() {.Tables("WHTSCANS").Columns("BAR_CODE")}

            'ASCMAIN1.sql = "Select  WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE,WHTLOCB1.LOCATION_CODE, LOCATION_ROUTE_SEQ " & vbCrLf _
            '    & " from WHTLOCB1, WHTLOCM1 " & vbCrLf _
            '    & " where WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE " & vbCrLf _
            '    & " And WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE " & vbCrLf _
            '    & " and (WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_CODE) in ( " & vbCrLf _
            '    & " select distinct  WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, " & vbCrLf _
            '    & " first_value(WHTLOCB1.LOCATION_CODE) over(partition by STYLE_CODE, COLOR_CODE order by WHTLOCB1.last_date desc,WHTLOCM1.LOCATION_ROUTE_SEQ) LOCATION_CODE" & vbCrLf _
            '    & " from WHTLOCB1, WHTLOCM1 " & vbCrLf _
            '    & " where WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE " & vbCrLf _
            '    & " And WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE " & vbCrLf _
            '    & " and WHTLOCM1.LOCATION_USE = :PARM1 " & vbCrLf _
            '    & " and (WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE) in " & vbCrLf _
            '    & " ( select WHSE_CODE, STYLE_CODE, COLOR_CODE from WHTLOCB1 " & vbCrLf _
            '    & " where whse_code = :PARM2 and location_code = :PARM3 AND LOCATION_QTY <> 0))"
            ASCMAIN1.sql = "select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE,
                              max(WHTLOCB1.LOCATION_CODE) keep (dense_rank first order by WHTLOCB1.last_date desc,WHTLOCM1.LOCATION_ROUTE_SEQ) LOCATION_CODE,
                              max(WHTLOCM1.LOCATION_ROUTE_SEQ) keep (dense_rank first order by WHTLOCB1.last_date desc,WHTLOCM1.LOCATION_ROUTE_SEQ) LOCATION_ROUTE_SEQ
                              from WHTLOCB1, WHTLOCM1 
                              where WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE 
                              And WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE 
                              and WHTLOCB1.LOCATION_QTY <> 0
                              and WHTLOCM1.LOCATION_USE = :PARM1 
                              and (WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE) in 
                              ( select WHSE_CODE, STYLE_CODE, COLOR_CODE from WHTLOCB1 
                              where whse_code = :PARM2 and location_code = :PARM3 AND LOCATION_QTY <> 0)
                              group by WHTLOCB1.WHSE_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTLOCBX", "**", 0, False, "VVV", 3)

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")

        End With

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
                Case "GET_MODE"
                    If SCANTEXT = "M2G" Then
                        AppStates("SCAN_LOC") = "Scan Move From Location|MFG2L|Shw Loc|EXIT|"
                    ElseIf SCANTEXT = "MFG2L" Then
                        AppStates("SCAN_LOC") = "Scan Move To Location|M2G|Dep WH|DepEcom|EXIT|"
                    Else
                        CreateResponse("", "R", "Error Select again")
                    End If
                    Mode = SCANTEXT
                    CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg(""))

                Case "SCAN_LOC"
                    If SCANTEXT.ToUpper = "DEP WH" Or SCANTEXT.ToUpper = "DEPECOM" Then
                        CreateResponse("SCAN_LOC", "YELLOW", FindNextUPC(SCANTEXT.ToUpper))
                        Exit Select
                    ElseIf SCANTEXT = "M2G" Then
                        AppStates("SCAN_LOC") = "Scan Move From Location|MFG2L|Shw Loc|EXIT|"
                        Mode = SCANTEXT
                        CreateResponse("", "R", DisplayMsg("Mode was Changed"))
                        Exit Select
                    ElseIf SCANTEXT = "MFG2L" Then
                        AppStates("SCAN_LOC") = "Scan Move To Location|M2G|Dep WH|DepEcom|EXIT|"
                        Mode = SCANTEXT
                        CreateResponse("", "R", DisplayMsg("Mode was Changed"))
                        Exit Select
                    ElseIf SCANTEXT = "Shw Loc" Then
                        CreateResponse("SCAN_SHOW", "BLUE", DisplayMsg("Scan UPC to Show Locations"))
                        Exit Select
                    Else
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                            Exit Select
                        End If

                        LOCATION_CODE = SCANTEXT.ToUpper
                        Dim Styles As String = CheckResponse("Stylelist")

                        dst.Tables("WHTMOVE1").Rows.Clear()
                        dst.Tables("WHTMOVE2").Rows.Clear()

                        CreateResponse("SCAN_UPC", "BLUE", DisplayMsg("Styles in Location: " & Styles))
                    End If

                Case "SCAN_SHOW"
                    If SCANTEXT = "BACK" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", "Show Cancelled, Scan  location")
                        Exit Select
                    ElseIf SCANTEXT = "<<" Or SCANTEXT = ">>" Then
                        Dim output As String = TACMAIN1.PaginateDataTable(tblPage, page, 6, "LOCATION_CODE,F1, LOCATION_QTY", SCANTEXT)
                        Dim msg As String = "Style " & STYLE_CODE & " Color " & COLOR_CODE & vbCrLf & output
                        CreateResponse("", "BLUE", msg)
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
                            page = 0
                            ASCMAIN1.sql = $"select b1.LOCATION_QTY, ' #' F1, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE
                                        from WHTLOCB1 b1
                                        join WHTLOCM1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE
                                        where b1.STYLE_CODE = '{STYLE_CODE}' 
                                            and b1.COLOR_CODE = '{COLOR_CODE}' 
                                            and nvl(m1.LOCATION_USE,'A') in ('A','E') 
                                            and m1.WHSE_CODE = '{G.WHSE_CODE}'
                                        order by 
                                            case 
                                                when b1.LOCATION_QTY > 0 then 1
                                                when b1.LOCATION_QTY < 0 then 2
                                                else 3
                                            end,
                                            b1.LOCATION_QTY desc, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"
                            tblPage = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                            Dim output As String = TACMAIN1.PaginateDataTable(tblPage, page, 6, "LOCATION_CODE,F1, LOCATION_QTY", ">>")
                            Dim msg As String = "Style " & STYLE_CODE & " Color " & COLOR_CODE & vbCrLf & output
                            CreateResponse("", "BLUE", msg)
                            Exit Select
                        End If

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
                        TACMAIN1.GetColors(Me, STYLE_CODE, LOCATION_CODE, COLOR_CODEs, colors)

                    End If
                    CreateResponse("SCAN_COLOR", "G", DisplayMsg("colors " & colors))

                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_LOC", "YELLOW", DisplayMsg("Move Cancelled, Scan  location"))
                        Exit Select
                    Else
                        If SCANTEXT = "0" Then
                            TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)
                            CreateResponse("", "B", DisplayMsg("All colors " & colors))
                            Exit Select
                        Else
                            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyColor(Me, SCANTEXT, STYLE_CODE, COLOR_CODEs)
                            If CheckResponse.ContainsKey("Error") Then
                                CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                                Exit Select
                            End If
                            UPC_CODE = CheckResponse("UPC_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
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
                        CASES_MOVED = Val(S(0))
                        UNITS_MOVED = Val(S(1))
                    Else
                        CASES_MOVED = Val(SCANTEXT)
                        UNITS_MOVED = 0
                    End If

                    If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                        CreateResponse("", "R", DisplayMsg("Invalid Number Of cases " & SCANTEXT))
                        Exit Select
                    End If
                    'CASES_MOVED = Val(SCANTEXT)
                    If CASES_MOVED + UNITS_MOVED = 0 Then
                        CreateResponse("SCAN_UNITS", "G", DisplayMsg("UPC " & UPC_CODE & " selected"))
                        Exit Select
                    End If

                    ASCMAIN1.sql = "Select CARTON_PACK_QTY from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'"
                    Dim CARTON_PACK_QTY As String = ASCDATA1.GetDataValue
                    UNITS_MOVED = UNITS_MOVED + Val(CARTON_PACK_QTY) * CASES_MOVED

                    ASCMAIN1.sql = "Select LOCATION_QTY from WHTLOCB1 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & G.WHSE_CODE & "' and LOCATION_CODE = '" & LOCATION_CODE & "'"
                    Dim LOC_QTY As String = ASCDATA1.GetDataValue
                    If UNITS_MOVED > Val(LOC_QTY & "") And Mode = "M2G" Then
                        'Warning
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Warning About to create Negative O/H|OK|NEW_QTY|"
                        CreateResponse("", "RED", DisplayMsg("Trying to move " & UNITS_MOVED & " - ONLY " & LOC_QTY & " O/H"))
                        AppStates(AppState) = hold
                        HOLD_PROMPT = "Entered " & SCANTEXT & " Cases, Carton pack: " & CARTON_PACK_QTY & ", Units to move: " & UNITS_MOVED
                        Exit Select
                    End If

                    CreateResponse("VERIFY", "B", DisplayMsg("Entered " & SCANTEXT & " Cases, Carton pack: " & CARTON_PACK_QTY & ", Units to move: " & UNITS_MOVED))
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
                        UNITS_MOVED = Val(SCANTEXT)

                        ASCMAIN1.sql = "Select LOCATION_QTY from WHTLOCB1 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & G.WHSE_CODE & "' and LOCATION_CODE = '" & LOCATION_CODE & "'"
                        Dim LOC_QTY As String = ASCDATA1.GetDataValue
                        If UNITS_MOVED > Val(LOC_QTY & "") And Mode = "M2G" Then
                            'Warning
                            Dim hold As String = AppStates(AppState)
                            AppStates(AppState) = "Warning About to create Negative O/H|OK|NEW_QTY|"
                            CreateResponse("", "RED", DisplayMsg("Trying to move " & UNITS_MOVED & " - ONLY " & LOC_QTY & " O/H"))
                            AppStates(AppState) = hold
                            HOLD_PROMPT = "Entered " & SCANTEXT & "Scanned " & SCANTEXT & " Units to Move"
                            Exit Select
                        End If

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

    Function FindNextUPC(DepType As String) As String
        Dim msg As String = ""
        Dim LocType As String = "A"

        If DepType = "DEPECOM" Then
            LocType = "E"
        End If

        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE FROM WHTLOCB1 where LOCATION_QTY > 0 AND WHSE_CODE = '" & G.WHSE_CODE & "' and LOCATION_CODE = '" & G.GUN_LOC & "'"
        Dim rowStyleColor As DataRow = ASCDATA1.GetDataRow
        If Not IsNothing(rowStyleColor) Then
            Fill_Records("WHTLOCBX", {LocType, G.WHSE_CODE, G.GUN_LOC}, True)
            For Each row As DataRow In dst.Tables("WHTLOCBX").Select("", "LOCATION_ROUTE_SEQ")
                If msg = "" Then
                    msg = "Open PutAways: " & vbCrLf
                End If
                msg = msg & row.Item("STYLE_CODE") & " - " & row.Item("COLOR_CODE") & " : " & row.Item("LOCATION_CODE") & vbCrLf
                'Exit For
            Next
            If msg = "" Then
                msg = "Next PutAway: " & rowStyleColor.Item("STYLE_CODE") & " - " & rowStyleColor.Item("COLOR_CODE") & vbCrLf _
                    & " Style not in " & IIf(LocType = "A", "Normal", "Ecom") & " Inventory"
            End If
        End If
        If msg = "" Then
            msg = "Nothing to deposit, All Done!"
        End If


        Return msg
    End Function

    Function DisplayMsg(ByVal note As String) As String
        Dim msg As String = ""
        Dim FromLoc As String = ""

        If Mode = "M2G" Then
            msg = "Move to Gun"
            FromLoc = LOCATION_CODE
        Else
            msg = "Move to Location"
            FromLoc = G.GUN_LOC
        End If
        If LOCATION_CODE <> "" Then
            msg = msg & vbCrLf & "LOC: " & LOCATION_CODE
            If STYLE_CODE <> "" Then
                msg = msg & vbCrLf & STYLE_CODE
                If COLOR_CODE <> "" Then
                    msg = msg & " " & COLOR_CODE
                    ASCMAIN1.sql = "Select LOCATION_QTY from WHTLOCB1 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & G.WHSE_CODE & "' and LOCATION_CODE = '" & FromLoc & "'"
                    Dim LOC_QTY As String = ASCDATA1.GetDataValue
                    msg = msg & vbCrLf & "Found: " & FromLoc & " QTY " & ": " & LOC_QTY
                End If
            End If
        Else
            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, G.GUN_LOC, False)
            If CheckResponse.ContainsKey("Stylelist") Then
                Dim Styles As String = CheckResponse("Stylelist")
                If Styles.Length > 2 Then
                    msg = msg & vbCrLf & "Styles in gun " & Styles
                End If
            End If
        End If

        If note <> "" Then
            msg = msg & vbCrLf & note
        End If
        Return msg
    End Function

    Sub ClearScanner()

        LOCATION_CODE = ""
        UPC_CODE = ""
        STYLE_CODE = ""
        COLOR_CODE = ""
        UNITS_MOVED = 0

    End Sub

    Sub Update_Record()

        Dim FromLoc As String = ""
        Dim ToLoc As String = ""

        If Mode = "M2G" Then
            FromLoc = LOCATION_CODE
            ToLoc = G.GUN_LOC
        Else 'If Mode = "MFG2L" Then
            FromLoc = G.GUN_LOC
            ToLoc = LOCATION_CODE
        End If

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
            .Item("WHSE_TRAN_QTY") = UNITS_MOVED
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
