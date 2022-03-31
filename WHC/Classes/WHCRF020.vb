Public Class WHCRF020
    ' Application Inventory Count

    Inherits WHCRF000

    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim COLOR_CODEs As New List(Of String)
    Dim Cases_count As Integer
    Dim TICKET_LNO As Integer
    Dim INVALID_BAR_CODE As String
    Dim CASES_BOOK As Integer
    Dim CASES_PHYS As Integer
    Dim UNITS_PHYS As Integer
    Dim TICKET_NO1 As String
    Dim BAR_CODE_LOCATION As String
    Dim CARTON_PACK_QTY As String
    Dim colors As String = ""
    Dim VoidList() As String
    Dim holdScan As String = ""
    Dim TICKET_STATUS As String = ""
    Dim Styles As String = ""
    Dim RESCAN As Boolean = False
    Dim OpenTickect As Boolean = False

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF009"

        AppStates.Add("SCAN_LOC", "Scan Location for Count|EXIT|") 'Yellow
        AppStates.Add("RE_COUNT", "Location already scanned|CANCEL|RECOUNT|ADD|")
        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style|EMPTY|UPDATE|CANCEL|EXIT|") 'Blue
        AppStates.Add("SCAN_COLOR", "Select a Color from List |CANCEL|")
        AppStates.Add("SCAN_CASE", "Enter Case Count, (0 to for units) |CANCEL|")
        AppStates.Add("SCAN_UNITS", "Enter Unit Count (unitsXcases)|CANCEL|")
        AppStates.Add("QTY_OK", " OK (Y/N)|Y|N|SKIP|")
        AppStates.Add("VERIFY", "Update Location (Y/N)|Y|N|CANCEL|")
        'AppStates.Add("SCAN_VOID", "Void Line|BACK|EXIT|")

        AppState = "SCAN_LOC"
        LAST_CLR = "YELLOW"

        With dst
            '  With .Tables.Add("WHTSCANS").Columns
            ' .Add("BAR_CODE")
            ' .Add("NEW")
            ' .Add("SCANNED")
            ' End With
            ' .Tables("WHTSCANS").PrimaryKey = New DataColumn() {.Tables("WHTSCANS").Columns("BAR_CODE")}

            Create_TDA(.Tables.Add, "ICTPHYC1", "*")
            Create_TDA(.Tables.Add, "ICTPHYC2", "*")

        End With

        tbl = dst.Tables("ICTPHYC2") ' New DataTable

        ' WHSE005


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
                Case "SCAN_LOC"
                    SCANTEXT = SCANTEXT.ToUpper
                    If SCANTEXT = "00" Then
                        SCANTEXT = LOCATION_CODE
                    End If
                    If SCANTEXT = "V" Then
                        'listCounted("")
                        'CreateResponse("SCAN_VOID", "B", listCounted())
                        Exit Select
                    End If

                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "YELLOW", CheckResponse("Error"))
                        Exit Select
                    End If

                    'If Not ASCMAIN1.Logical_Lock("ICTPHYC1", G.WHSE_CODE & ":" & G.USER_ID) Then
                    '    Dim User As String = ASCMAIN1.MultiTask_Get_Users("ICTPHYC1", G.WHSE_CODE & ":" & G.USER_ID, "L")
                    '    If (Not User.Contains(G.USER_ID)) Then
                    '        EMsg = vbCr & "Pick Ticket is Locked " _
                    '                 & User
                    '        CreateResponse("", "GREEN", EMsg)
                    '        Exit Select
                    '    End If
                    'End If

                    LOCATION_CODE = SCANTEXT
                    Styles = CheckResponse("Stylelist")

                    dst.Tables("ICTPHYC2").Rows.Clear()
                    dst.Tables("ICTPHYC1").Rows.Clear()
                    TICKET_LNO = 0

                    If Chk_RESCAN() = False Then
                        Dim TICKET_NO As String = ASCMAIN1.Next_Control_No("ICTPHYC1.TICKET_NO")
                        TICKET_NO1 = TICKET_NO
                        TICKET_STATUS = ""
                        CreateResponse("SCAN_UPC", "BLUE", "Location " & SCANTEXT & " has been selected, Styles in Location: " & Styles)
                        OpenTickect = False
                    Else
                        ASCMAIN1.sql = "select * FROM ICTPHYC1 " & vbCrLf _
                            & " where nvl(TICKET_STATUS,'A') <> 'V'" & vbCrLf _
                            & " and LOCATION_CODE = '" & LOCATION_CODE & "'"
                        Fill_Records("ICTPHYC1", "", True, ASCMAIN1.sql)

                        Dim row As DataRow = dst.Tables("ICTPHYC1").Rows(0)
                        TICKET_NO1 = row.Item("TICKET_NO")
                        OpenTickect = True

                        If row.Item("TICKET_STATUS") & "" = "E" Then
                            'rescanned empty ticket
                            CreateResponse("SCAN_LOC", "RED", "Location Flagged as Empty, Void original ticket")
                            Exit Select
                        Else
                            For Each row1 As DataRow In dst.Tables("ICTPHYC1").Select("")
                                ASCMAIN1.sql = "select * FROM ICTPHYC2 " & vbCrLf _
                                & " where TICKET_NO = '" & row1.Item("TICKET_NO") & "'" & vbCrLf
                                Fill_Records("ICTPHYC2", "", False, ASCMAIN1.sql)
                            Next
                            TICKET_LNO = dst.Tables("ICTPHYC2").Compute("MAX(TICKET_LNO)", "TICKET_NO = '" & TICKET_NO1 & "'")
                        End If
                    End If

                Case "RE_COUNT"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re-scan  location")
                        Exit Select
                    ElseIf SCANTEXT = "RECOUNT" Then
                        RESCAN = True
                        'For Each row2 As DataRow In dst.Tables("ICTPHYC2").Select("STATUS IS NULL")
                        '    row2.Item("STATUS") = "V"
                        '    row2.Item("NEW_TICKET") = G.USER_ID
                        'Next
                    ElseIf SCANTEXT = "ADD" Then
                        RESCAN = False
                    End If
                    CreateResponse("SCAN_UPC", "BLUE", "Location " & SCANTEXT & " has been selected, Styles in Location: " & Styles)

                Case "SCAN_UPC"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re-scan  location")
                        Exit Select
                    ElseIf SCANTEXT = "UPDATE" Then
                        CreateResponse("VERIFY", "R", DisplayVerify())
                        Exit Select
                    ElseIf SCANTEXT = "EMPTY" Then
                        TICKET_STATUS = "E"
                        Dim hold As String = ""
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = "Update Empty Location (Y/N)|Y|N|CANCEL|"
                        CreateResponse("VERIFY", "R", String.Format("Flag {0} location as Empty", LOCATION_CODE))
                        AppStates("VERIFY") = hold
                        Exit Select
                    Else
                        SCANTEXT = Trim(SCANTEXT)
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", CheckResponse("Error"))
                            Exit Select
                        End If

                        If CheckResponse.ContainsKey("UPC_CODE") Then
                            UPC_CODE = CheckResponse("UPC_CODE")
                            STYLE_CODE = CheckResponse("STYLE_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            askCaseOrUnits()
                            Exit Select
                        End If
                        STYLE_CODE = SCANTEXT.ToUpper
                        TACMAIN1.GetColors(Me, STYLE_CODE, LOCATION_CODE, COLOR_CODEs, colors)

                    End If
                    CreateResponse("SCAN_COLOR", "G", DisplayDetails())

                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re-scan  location")
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
                        End If
                    End If
                    askCaseOrUnits()

                Case "SCAN_CASE"
                    CASES_PHYS = 0
                    UNITS_PHYS = 0
                    Dim hold As String
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re-scan location")
                        Exit Select
                    Else
                        If SCANTEXT = "0" Then
                            CreateResponse("SCAN_UNITS", "G", DisplayDetails())
                            Exit Select
                        End If
                        If SCANTEXT.Contains("*") Then
                            Dim S() As String
                            S = SCANTEXT.Split("*")
                            CASES_PHYS = Val(S(0))
                            UNITS_PHYS = Val(S(1))
                        Else
                            CASES_PHYS = Val(SCANTEXT)
                            UNITS_PHYS = 0
                        End If

                        If CASES_PHYS + UNITS_PHYS = 0 Then
                            CreateResponse("SCAN_UNITS", "G", DisplayDetails())
                            Exit Select
                        End If

                        ASCMAIN1.sql = "Select CARTON_PACK_QTY from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'"
                        CARTON_PACK_QTY = ASCDATA1.GetDataValue
                        Dim Counted As Integer = CARTON_PACK_QTY * CASES_PHYS + UNITS_PHYS

                        If Counted > 9999 Or Counted < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Number Of cases " & SCANTEXT & ", Try again|OK|"
                            CreateResponse("", "R", DisplayDetails())
                            AppStates(AppState) = hold
                            Exit Select

                        End If

                        Dim msg As String = ""
                        If CASES_PHYS > 0 Then
                            msg = "Counted: " & CASES_PHYS & "c" & If(UNITS_PHYS > 0, UNITS_PHYS & "u", "")
                        ElseIf UNITS_PHYS > 0 Then
                            msg = "Counted: " & UNITS_PHYS & "u"
                        End If
                        hold = AppStates("QTY_OK")
                        AppStates("QTY_OK") = msg & hold
                        CreateResponse("QTY_OK", "R", DisplayDetails())
                        AppStates("QTY_OK") = hold

                    End If

                Case "SCAN_UNITS"
                    CASES_PHYS = 0
                    UNITS_PHYS = 0
                    Dim hold As String
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re - scan  location")
                        Exit Select
                    Else
                        If SCANTEXT.Contains("*") Then
                            Dim S() As String
                            S = SCANTEXT.ToUpper.Split("X")
                            CASES_PHYS = 0
                            UNITS_PHYS = Val(S(0)) * Val(S(1))
                        Else
                            CASES_PHYS = 0
                            UNITS_PHYS = Val(SCANTEXT)
                        End If
                        'If Val(UNITS_PHYS) > 9999 Or Val(UNITS_PHYS) < 1 Then
                        '    'Error
                        '    hold = AppStates(AppState)
                        '    AppStates(AppState) = "Invalid Number Of units " & SCANTEXT & ", Try again|OK|"
                        '    CreateResponse("", "R", DisplayDetails())
                        '    AppStates(AppState) = hold
                        '    Exit Select

                        'End If

                        ASCMAIN1.sql = "Select CARTON_PACK_QTY from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'"
                        CARTON_PACK_QTY = ASCDATA1.GetDataValue

                        Dim msg As String = ""
                        If UNITS_PHYS > 0 Then
                            msg = "Counted: " & UNITS_PHYS & "u"
                        End If
                        hold = AppStates("QTY_OK")
                        AppStates("QTY_OK") = msg & hold
                        CreateResponse("QTY_OK", "R", DisplayDetails())
                        AppStates("QTY_OK") = hold


                        'hold = AppStates("SCAN_UPC")
                        ''don't allow empty at this point
                        'AppStates("SCAN_UPC") = "Scan UPC or Enter Style|UPDATE|CANCEL|EXIT|"
                        'CreateResponse("SCAN_UPC", "BLUE", "Location " & LOCATION_CODE & " has been selected, Styles in Location: " & Styles)
                        'AppStates("SCAN_UPC") = hold
                    End If

                Case "QTY_OK"
                    'Dim hold As String = AppStates("SCAN_LOC")
                    'AppStates("SCAN_LOC") = "Scan Location, 00:'" & LOCATION_CODE & "', V:Void|EXIT|"
                    If SCANTEXT = "Y" Then

                        If RESCAN = True Then
                            For Each row2 As DataRow In dst.Tables("ICTPHYC2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                                row2.Item("STATUS") = "V"
                                row2.Item("NEW_TICKET") = G.USER_ID
                            Next
                        End If

                        Dim rowICTPHYC2 As DataRow ' = dst.Tables("ICTPHYC2").Rows.Find(New Object() {G.WHSE_CODE, TICKET_NO1, "1"})
                        TICKET_LNO += 1
                        rowICTPHYC2 = tbl.NewRow
                        rowICTPHYC2.Item("WHSE_CODE") = G.WHSE_CODE
                        rowICTPHYC2.Item("TICKET_NO") = TICKET_NO1
                        rowICTPHYC2.Item("TICKET_LNO") = TICKET_LNO
                        rowICTPHYC2.Item("STYLE_CODE") = STYLE_CODE
                        rowICTPHYC2.Item("COLOR_CODE") = COLOR_CODE
                        rowICTPHYC2.Item("COUNT_CTNS") = CASES_PHYS
                        rowICTPHYC2.Item("CARTON_PACK_QTY") = Val(CARTON_PACK_QTY)
                        rowICTPHYC2.Item("COUNT_LOOSE") = UNITS_PHYS
                        rowICTPHYC2.Item("NEW_TICKET") = IIf(RESCAN, "YES", "")
                        tbl.Rows.Add(rowICTPHYC2)

                        Dim hold As String = AppStates("SCAN_UPC")
                        'don't allow empty at this point
                        AppStates("SCAN_UPC") = "Scan UPC or Enter Style|UPDATE|CANCEL|EXIT|"
                        CreateResponse("SCAN_UPC", "BLUE", "Location " & LOCATION_CODE & " has been selected, Styles in Location: " & Styles)
                        AppStates("SCAN_UPC") = hold
                    ElseIf SCANTEXT = "N" Then
                        CASES_PHYS = 0
                        UNITS_PHYS = 0
                        If dst.Tables("ICTPHYC2").Rows.Count = 0 Then
                            CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re - scan  location")
                        Else
                            askCaseOrUnits()
                        End If
                    ElseIf SCANTEXT = "SKIP" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Location " & SCANTEXT & " has been selected, Styles in Location: " & Styles)
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
                    'AppStates("SCAN_LOC") = hold

                Case "VERIFY"
                    'Dim hold As String = AppStates("SCAN_LOC")
                    AppStates("SCAN_LOC") = "Scan Location, 00:'" & LOCATION_CODE & "', V:Void|EXIT|"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Updated, Scan  location")
                    ElseIf SCANTEXT = "N" Then
                        CASES_PHYS = 0
                        UNITS_PHYS = 0
                        If dst.Tables("ICTPHYC2").Rows.Count = 0 Then
                            CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re - scan  location")
                        Else
                            askCaseOrUnits()
                        End If
                    ElseIf SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re - scan  location")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
                    'AppStates("SCAN_LOC") = hold
            End Select
        End If
    End Sub

    Function Chk_RESCAN() As Boolean
        ASCMAIN1.sql = "select COUNT(1) FROM ICTPHYC1 " & vbCrLf _
                    & " where nvl(TICKET_STATUS,'A') <> 'V'" & vbCrLf _
                    & " and LOCATION_CODE = '" & LOCATION_CODE & "'"
        Dim counts As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql)
        ASCMAIN1.sql = "Select COUNT(1) FROM ICTPHYC1" & vbCrLf _
            & "where VERIFIED_OPER is not null and nvl(TICKET_STATUS,'A') <> 'V'" & vbCrLf _
             & " and LOCATION_CODE = '" & LOCATION_CODE & "'"
        Dim Locked As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql)
        If Locked > 0 Then
            CreateResponse("", "R", "This Location has been Verified, Can't Recount")
            RESCAN = True
            Return RESCAN
            Exit Function
        End If
        If counts <> 0 Then
            RESCAN = True
            CreateResponse("RE_COUNT", "R", "Previous Counts Found, Please verify")
        Else
            RESCAN = False
        End If

        Return RESCAN
    End Function

    Function DisplayDetails() As String
        Dim msg As String
        msg = "LOC: " & LOCATION_CODE
        If STYLE_CODE <> "" Then
            Dim row As DataRow = ASCDATA1.GetDataRow("Select * from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'")
            msg &= vbCrLf & "Style: " & STYLE_CODE & " " & row.Item("STYLE_DESC")
            If COLOR_CODE <> "" Then
                msg &= vbCrLf & "Clr " & COLOR_CODE & ": " & ASCDATA1.GetDataValue("Select COLOR_CODE_LONG from ICTCOLR1 where COLOR_CODE = '" & COLOR_CODE & "'")
            Else
                msg &= vbCrLf & "Clrs " & colors
            End If
            If Not row.Item("CARTONS_PER_UNIT").Equals(Null) AndAlso row.Item("CARTONS_PER_UNIT") > 0 Then
                Dim csflg As String = ""
                msg &= vbCrLf & String.Format("{0}c per unit ", row.Item("CARTONS_PER_UNIT")) & csflg.PadLeft(Val(row.Item("CARTONS_PER_UNIT")), "*")
            Else
                msg &= vbCrLf & "Case Qty " & row.Item("CARTON_PACK_QTY") & "u"
                If row.Item("INNER_PACK_QTY").ToString <> "" AndAlso row.Item("INNER_PACK_QTY") > 0 Then
                    msg &= " / Inner Qty " & row.Item("INNER_PACK_QTY") & "u"
                End If
            End If
        End If

        Return msg
    End Function

    Function DisplayVerify() As String
        Dim msg As String
        msg = "LOC: " & LOCATION_CODE
        If TICKET_STATUS = "E" Then
            msg &= "Empty Location"
        Else
            'msg &= vbCrLf & "Styles Counted: " & tbl.Compute("COUNT(STYLE_CODE)", "").ToString
            msg &= vbCrLf & "Counts Entered: " & tbl.Compute("COUNT(TICKET_NO)", "").ToString
            For Each row As DataRow In tbl.Select("")
                If row.Item("STATUS") & "" <> "V" Then
                    msg &= vbCrLf & "Style/clr #" & row.Item("STYLE_CODE") & "/" & row.Item("COLOR_CODE") & " c" & row.Item("COUNT_CTNS") & " u" & row.Item("COUNT_LOOSE")
                End If
            Next

        End If

        Return msg
    End Function

    Sub askCaseOrUnits()
        Dim row As DataRow = ASCDATA1.GetDataRow("Select * from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'")
        If Not row.Item("CARTONS_PER_UNIT").Equals(Null) AndAlso row.Item("CARTONS_PER_UNIT") > 0 Then
            CreateResponse("SCAN_UNITS", "G", DisplayDetails())
        Else
            CreateResponse("SCAN_CASE", "G", DisplayDetails())
        End If
    End Sub

    Sub Update_Record()

        BeginTrans()
        Dim rowICTPHYC1 As DataRow

        If Not OpenTickect Then
            rowICTPHYC1 = dst.Tables("ICTPHYC1").NewRow
            With rowICTPHYC1
                .Item("WHSE_CODE") = G.WHSE_CODE
                .Item("TICKET_NO") = TICKET_NO1
                .Item("COUNT_BY") = G.USER_ID
                .Item("LOCATION_CODE") = LOCATION_CODE
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = G.USER_ID
                .Item("TICKET_STATUS") = TICKET_STATUS
            End With
            dst.Tables("ICTPHYC1").Rows.Add(rowICTPHYC1)
        Else
            For Each rowICTPHYC1 In dst.Tables("ICTPHYC1").Select("TICKET_NO = '" & TICKET_NO1 & "'")
                With rowICTPHYC1
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = G.USER_ID
                    .Item("TICKET_STATUS") = TICKET_STATUS
                End With
            Next
        End If

        ASCMAIN1.sql = "Update WHTLOCM1 " & vbCrLf _
            & " set LOCATION_LOCKED = '1' " & vbCrLf _
            & " where WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
            & " and LOCATION_CODE = '" & LOCATION_CODE & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "BEGIN " & vbCrLf _
            & " For i in (Select * FROM WHTLOCB1 Where WHSE_CODE = '" & G.WHSE_CODE & "' AND LOCATION_CODE = '" & LOCATION_CODE & "') " & vbCrLf _
            & " LOOP " & vbCrLf _
            & " Update WHTLOCB0 set BOOK_INVTY_ADJ = nvl(LOCATION_QTY,0) - nvl(i.LOCATION_QTY,0) " & vbCrLf _
            & " where WHSE_CODE = i.WHSE_CODE " & vbCrLf _
            & " and LOCATION_CODE = i.LOCATION_CODE " & vbCrLf _
            & " and BAR_CODE = i.BAR_CODE " & vbCrLf _
            & " and STYLE_CODE = i.STYLE_CODE " & vbCrLf _
            & " and COLOR_CODE = i.COLOR_CODE; " & vbCrLf _
            & " IF (SQL%ROWCOUNT = 0) THEN  " & vbCrLf _
            & "   INSERT INTO WHTLOCB0 (WHSE_CODE ,LOCATION_CODE ,BAR_CODE ,STYLE_CODE ,COLOR_CODE ,LOCATION_QTY ,INIT_DATE ,INIT_OPER ,LAST_DATE ,LAST_OPER ,LOCATION_QTY_WAVE, BOOK_INVTY_ADJ)  " & vbCrLf _
            & "   VALUES (i.WHSE_CODE ,i.LOCATION_CODE ,i.BAR_CODE ,i.STYLE_CODE ,i.COLOR_CODE ,0 ,i.INIT_DATE ,i.INIT_OPER ,i.LAST_DATE ,i.LAST_OPER ,i.LOCATION_QTY_WAVE, i.LOCATION_QTY ); " & vbCrLf _
            & " END IF; " & vbCrLf _
            & " END LOOP;" & vbCrLf _
            & "END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Update_Record_TDA("ICTPHYC1")
        Update_Record_TDA("ICTPHYC2")

        CommitTrans()

        STYLE_CODE = ""
        COLOR_CODE = ""
        RESCAN = False

    End Sub


End Class
