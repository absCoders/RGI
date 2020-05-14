Public Class WHCRF009
    ' Application Inventory Count

    Inherits WHCRF000

    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim COLOR_CODEs As New List(Of String)
    Dim Cases_count As Integer
    Dim TICKET_NO As String
    Dim INVALID_BAR_CODE As String
    Dim CASES_BOOK As Integer
    Dim CASES_PHYS As Integer
    Dim UNITS_PHYS As Integer
    Dim TICKET_NO1 As String
    Dim BAR_CODE_LOCATION As String
    Dim colors As String = ""
    Dim VoidList() As String
    Dim holdScan As String = ""
    Dim TICKET_STATUS As String = ""

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF009"

        AppStates.Add("SCAN_LOC", "Scan Location for Count, V:Void|EXIT|") 'Yellow
        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style|EMPTY|CANCEL|EXIT|") 'Blue
        AppStates.Add("SCAN_COLOR", "Select a Color from List |CANCEL|")
        AppStates.Add("SCAN_CASE", "Enter Case Count, (0 to for units) |CANCEL|")
        AppStates.Add("SCAN_UNITS", "Enter Unit Count |CANCEL|")
        AppStates.Add("VERIFY", "Update Count (Y/N)|Y|N|CANCEL|")
        AppStates.Add("SCAN_VOID", "Void Line|BACK|EXIT|")

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
                        listCounted("")
                        'CreateResponse("SCAN_VOID", "B", listCounted())
                        Exit Select
                    End If

                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "YELLOW", CheckResponse("Error"))
                        Exit Select
                    End If

                    LOCATION_CODE = SCANTEXT
                    Dim Styles As String = CheckResponse("Stylelist")

                    dst.Tables("ICTPHYC2").Rows.Clear()
                    dst.Tables("ICTPHYC1").Rows.Clear()

                    Dim TICKET_NO As String = ASCMAIN1.Next_Control_No("ICTPHYC1.TICKET_NO")
                    TICKET_NO1 = TICKET_NO
                    TICKET_STATUS = ""
                    CreateResponse("SCAN_UPC", "BLUE", "Location " & SCANTEXT & " has been selected, Styles in Location: " & Styles)

                Case "SCAN_UPC"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re-scan  location")
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
                        STYLE_CODE = SCANTEXT
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
                        Dim CARTON_PACK_QTY As String = ASCDATA1.GetDataValue
                        Dim Counted As Integer = CARTON_PACK_QTY * CASES_PHYS + UNITS_PHYS

                        If Counted > 9999 Or Counted < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Number Of cases " & SCANTEXT & ", Try again|OK|"
                            CreateResponse("", "R", DisplayDetails())
                            AppStates(AppState) = hold
                            Exit Select

                        End If

                        Dim rowICTPHYC2 As DataRow = dst.Tables("ICTPHYC2").Rows.Find(New Object() {G.WHSE_CODE, TICKET_NO1, "1"})
                        If rowICTPHYC2 Is Nothing Then
                            rowICTPHYC2 = tbl.NewRow
                            rowICTPHYC2.Item("WHSE_CODE") = G.WHSE_CODE
                            rowICTPHYC2.Item("TICKET_NO") = TICKET_NO1
                            rowICTPHYC2.Item("TICKET_LNO") = 1
                            rowICTPHYC2.Item("STYLE_CODE") = STYLE_CODE
                            rowICTPHYC2.Item("COLOR_CODE") = COLOR_CODE
                            rowICTPHYC2.Item("COUNT_CTNS") = CASES_PHYS
                            rowICTPHYC2.Item("CARTON_PACK_QTY") = Val(CARTON_PACK_QTY)
                            rowICTPHYC2.Item("COUNT_LOOSE") = UNITS_PHYS
                            tbl.Rows.Add(rowICTPHYC2)
                        Else
                            rowICTPHYC2.Item("STYLE_CODE") = STYLE_CODE
                            rowICTPHYC2.Item("COLOR_CODE") = COLOR_CODE
                            rowICTPHYC2.Item("COUNT_CTNS") = CASES_PHYS
                            rowICTPHYC2.Item("CARTON_PACK_QTY") = Val(CARTON_PACK_QTY)
                            rowICTPHYC2.Item("COUNT_LOOSE") = UNITS_PHYS
                        End If
                        Dim msg As String = ""
                        If CASES_PHYS > 0 Then
                            msg = "Counted: " & CASES_PHYS & "c" & If(UNITS_PHYS > 0, UNITS_PHYS & "u", "")
                        ElseIf UNITS_PHYS > 0 Then
                            msg = "Counted: " & UNITS_PHYS & "u"
                        End If
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = msg & vbCrLf & hold
                        CreateResponse("VERIFY", "R", DisplayDetails())
                        AppStates("VERIFY") = hold
                    End If

                Case "SCAN_UNITS"
                    CASES_PHYS = 0
                    UNITS_PHYS = 0
                    Dim hold As String
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "Count Cancelled, Re - scan  location")
                        Exit Select
                    Else
                        If Val(SCANTEXT) > 9999 Or Val(SCANTEXT) < 1 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Number Of units " & SCANTEXT & ", Try again|OK|"
                            CreateResponse("", "R", DisplayDetails())
                            AppStates(AppState) = hold
                            Exit Select

                        End If
                        UNITS_PHYS = Val(SCANTEXT)

                        ASCMAIN1.sql = "Select CARTON_PACK_QTY from ICTSTYL1 where STYLE_CODE = '" & STYLE_CODE & "'"
                        Dim CARTON_PACK_QTY As String = ASCDATA1.GetDataValue

                        'Do I need to try to find the record?
                        Dim rowICTPHYC2 As DataRow = dst.Tables("ICTPHYC2").Rows.Find(New Object() {G.WHSE_CODE, TICKET_NO, "1"})
                        If rowICTPHYC2 Is Nothing Then
                            rowICTPHYC2 = tbl.NewRow
                            rowICTPHYC2.Item("WHSE_CODE") = G.WHSE_CODE
                            rowICTPHYC2.Item("TICKET_NO") = TICKET_NO1
                            rowICTPHYC2.Item("TICKET_LNO") = 1
                            rowICTPHYC2.Item("STYLE_CODE") = STYLE_CODE
                            rowICTPHYC2.Item("COLOR_CODE") = COLOR_CODE
                            rowICTPHYC2.Item("COUNT_CTNS") = CASES_PHYS
                            rowICTPHYC2.Item("CARTON_PACK_QTY") = Val(CARTON_PACK_QTY)
                            rowICTPHYC2.Item("COUNT_LOOSE") = UNITS_PHYS
                            tbl.Rows.Add(rowICTPHYC2)
                        Else
                            rowICTPHYC2.Item("STYLE_CODE") = STYLE_CODE
                            rowICTPHYC2.Item("COLOR_CODE") = COLOR_CODE
                            rowICTPHYC2.Item("COUNT_CTNS") = CASES_PHYS
                            rowICTPHYC2.Item("CARTON_PACK_QTY") = Val(CARTON_PACK_QTY)
                            rowICTPHYC2.Item("COUNT_LOOSE") = UNITS_PHYS
                        End If
                        Dim msg As String = ""
                        If UNITS_PHYS > 0 Then
                            msg = "Counted: " & UNITS_PHYS & "u"
                        End If
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = msg & vbCrLf & hold
                        CreateResponse("VERIFY", "R", DisplayDetails())
                        AppStates("VERIFY") = hold
                    End If

                Case "SCAN_VOID"
                    If SCANTEXT = "BACK" Then
                        CreateResponse("SCAN_LOC", "YELLOW", "")
                        Exit Select
                    ElseIf SCANTEXT = "YES" Then
                        voidLine(holdScan)
                        holdScan = ""
                        listCounted("")
                        Exit Select
                    ElseIf SCANTEXT = "NO" Then
                        listCounted("")
                        holdScan = ""
                        Exit Select
                    Else
                        If holdScan = "" Then
                            'handle voids
                            If Val(SCANTEXT) > (VoidList.Length - 1) Or Val(SCANTEXT) < 1 Then
                                'error
                            Else
                                holdScan = SCANTEXT
                                listCounted(holdScan)
                                Exit Select
                            End If
                        Else
                            listCounted(holdScan)
                            Exit Select
                        End If
                    End If
                    listCounted("")

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

    Sub listCounted(ByRef showline As String)
        Dim msg As String = "Recent Counts"
        Dim lno As Integer = 0
        Dim lines As String = ""
        ASCMAIN1.sql = "select i1.TICKET_NO, i1.LOCATION_CODE, i1.TICKET_STATUS, i2.STYLE_CODE, i2.COLOR_CODE, " & vbCrLf _
                    & " SUM(i2.COUNT_CTNS) COUNT_CTNS, SUM(i2.COUNT_LOOSE) COUNT_LOOSE, MIN(i2.CARTON_PACK_QTY) CARTON_PACK_QTY " & vbCrLf _
                    & " From ictphyc1 i1  " & vbCrLf _
                    & " left join ictphyc2 i2 on i1.TICKET_NO = i2.TICKET_NO " & vbCrLf _
                    & " where i1.INIT_OPER = '" & G.USER_ID & "' " & vbCrLf _
                    & " and nvl(TICKET_STATUS,'A') <> 'V' " & vbCrLf _
                    & " group by i1.TICKET_NO, i1.LOCATION_CODE, i1.TICKET_STATUS, i2.STYLE_CODE, i2.COLOR_CODE "

        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("", "TICKET_NO desc")
        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                lno += 1
                If showline = "" Or Val(showline) = lno Then
                    If Not ROW.Item("TICKET_STATUS").Equals(Null) AndAlso ROW.Item("TICKET_STATUS") = "E" Then
                        msg = msg & vbCrLf & "L" & lno & " " & ROW.Item("LOCATION_CODE") & " EMPTY"
                    Else
                        msg = msg & vbCrLf & "L" & lno & " " & ROW.Item("LOCATION_CODE") & " " & ROW.Item("STYLE_CODE") & " " & ROW.Item("COLOR_CODE") & " " _
                        & If(ROW.Item("COUNT_CTNS") = 0, "", ROW.Item("COUNT_CTNS") & "c") & If(ROW.Item("COUNT_LOOSE") = 0, "", ROW.Item("COUNT_LOOSE") & "u")
                    End If
                End If
                lines &= "|" & ROW.Item("TICKET_NO")
                If lno = 5 Then Exit For
            Next
            VoidList = lines.Split("|")
            If showline <> "" Then
                Dim hold As String = AppStates("SCAN_VOID")
                AppStates("SCAN_VOID") = "Void this line|YES|NO|"
                CreateResponse("SCAN_VOID", "R", msg)
                AppStates("SCAN_VOID") = hold
            Else
                CreateResponse("SCAN_VOID", "R", msg)
            End If
        Else
            msg = "No Lines to Void" & vbCrLf & "Back to continue"
            VoidList = {}
            Dim hold As String = AppStates("SCAN_VOID")
            AppStates("SCAN_VOID") = "No Counts, CLICK BELOW|BACK|EXIT|"
            CreateResponse("SCAN_VOID", "R", msg)
            AppStates("SCAN_VOID") = hold
        End If
    End Sub
    Sub voidLine(ByRef SCANTEXT As String)

        BeginTrans()

        Dim void = VoidList(Val(SCANTEXT))

        Fill_Records("ICTPHYC1", "", True, "Select * from ICTPHYC1 where TICKET_NO = '" & void & "'")
        Fill_Records("ICTPHYC2", "", True, "Select * from ICTPHYC2 where TICKET_NO = '" & void & "'")

        Dim rowICTPHYC1 As DataRow = dst.Tables("ICTPHYC1").Select("").FirstOrDefault
        Dim rowICTPHYC2 As DataRow = dst.Tables("ICTPHYC2").Select("").FirstOrDefault

        If Not rowICTPHYC1.Item("TICKET_STATUS").Equals(Null) AndAlso rowICTPHYC1.Item("TICKET_STATUS") = "E" Then
            ' STATUS = E has no detail
        Else
            Dim row2 As DataRow = dst.Tables("ICTPHYC2").NewRow
            row2.Item("WHSE_CODE") = G.WHSE_CODE
            row2.Item("TICKET_NO") = rowICTPHYC2.Item("TICKET_NO")
            row2.Item("TICKET_LNO") = 2
            row2.Item("STYLE_CODE") = rowICTPHYC2.Item("STYLE_CODE")
            row2.Item("COLOR_CODE") = rowICTPHYC2.Item("COLOR_CODE")
            row2.Item("COUNT_CTNS") = (rowICTPHYC2.Item("COUNT_CTNS") * -1)
            row2.Item("CARTON_PACK_QTY") = rowICTPHYC2.Item("CARTON_PACK_QTY")
            row2.Item("COUNT_LOOSE") = (rowICTPHYC2.Item("COUNT_LOOSE") * -1)
            dst.Tables("ICTPHYC2").Rows.Add(row2)
        End If

        rowICTPHYC1.Item("TICKET_STATUS") = "V"
        rowICTPHYC1.Item("LAST_DATE") = DATETIME_STAMP

        Update_Record_TDA("ICTPHYC1")
        Update_Record_TDA("ICTPHYC2")

        CommitTrans()

        STYLE_CODE = ""
        COLOR_CODE = ""

    End Sub
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

        Dim rowICTPHYC1 As DataRow = dst.Tables("ICTPHYC1").NewRow
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

        Update_Record_TDA("ICTPHYC1")
        Update_Record_TDA("ICTPHYC2")

        CommitTrans()

        STYLE_CODE = ""
        COLOR_CODE = ""

    End Sub


End Class
