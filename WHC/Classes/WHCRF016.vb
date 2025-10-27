Public Class WHCRF016
    ' Application Pick Ticket no LPN

    Inherits WHCRF000

    Dim PICK_NO As String
    Dim PICK_LNO As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim COLOR_DESC As String
    Dim UPC_CODE As String
    Dim STYLE_DESC As String
    Dim LOCATION_CODE As String
    Dim GOTO_LOCATION As String
    '    Dim COLOR_CODEs As New List(Of String)
    Dim Cases_count As Integer
    Dim TICKET_NO As String
    Dim CASES_BOOK As Integer
    Dim CASES_MOVED As Integer
    Dim UNITS_MOVED As Integer
    Dim TICKET_NO1 As String
    Dim ORIGINAL_LOCATION As String
    '    Dim colors As String = ""
    Dim ScanLocMsg As String
    Dim RowNum As Integer
    Dim Picks As String
    Dim OpenPicks As Integer
    Dim VoidList() As String
    Dim VoidTranNo As String
    Dim PICK_QTY As Integer
    Dim PICK_QTY_OPEN As Integer
    Dim INNER_PACK_QTY As Integer
    Dim CARTON_PACK_QTY As Integer
    Dim CARTONS_PER_UNIT As Integer
    Dim holdScan As String
    Dim _sql As String
    Dim PickTypeSql As String
    Dim rowWHTSCANS As DataRow
    Dim PickSplit As String
    Dim PageNo As Integer = 0
    Dim LockWarning As Boolean
    Dim WHSE_CODE As String
    Dim Automated As Boolean = False
    Dim page As Int16 = 0
    Dim tblPage As DataTable = Nothing

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF016"

        AppStates.Add("SCAN_PTCKT", "Scan Pick Ticket |EXIT|") 'Green
        AppStates.Add("SCAN_LOC", "Scan Loc, 00:'{0}', V:Void|NEXT UPC|PREV UPC|<<|>>|DONE|") 'Yellow
        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style|CANCEL|") 'Blue
        AppStates.Add("SCAN_COLOR", "Enter Color code|CANCEL|")
        AppStates.Add("SCAN_CASES", "How many cases picked, (0 for units)|CANCEL|")
        AppStates.Add("SCAN_UNITS", "How many units picked|CANCEL|")
        AppStates.Add("VERIFY", "Update|Y|N|CLEAR|EXIT|")
        AppStates.Add("LEAVE", "Units in Gun, Gun must be Empty for Pick|EXIT|")
        AppStates.Add("SCAN_VOID", "Void Line or 'ALL'|BACK|NXT PAGE|NEW PICK|EXIT|")

        AppState = "SCAN_PTCKT"
        ScanLocMsg = AppStates("SCAN_LOC")
        LAST_CLR = "GREEN"

        With dst
            With .Tables.Add("WHTSCANS")
                .Columns.Add("LOCATION_ROUTE_SEQ", GetType(System.Int32))
                .Columns.Add("LOCATION_CODE")
                .Columns.Add("PICK_NO")
                .Columns.Add("PICK_LNO")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("COLOR_DESC")
                .Columns.Add("UPC_CODE")
                .Columns.Add("STYLE_DESC")
                .Columns.Add("LOCATION_QTY", GetType(System.Int32))
                .Columns.Add("PICK_QTY", GetType(System.Int32))
                .Columns.Add("PICK_QTY_OPEN", GetType(System.Int32))
                .Columns.Add("INNER_PACK_QTY", GetType(System.Int32))
                .Columns.Add("CARTON_PACK_QTY", GetType(System.Int32))
                .Columns.Add("CARTONS_PER_UNIT", GetType(System.Int32))
                .Columns.Add("STYLE_ASST_QTY", GetType(System.Int32))
                .Columns.Add("ROWNUM", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("LOCATION_ROUTE_SEQ"), .Columns("LOCATION_CODE"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With

            Create_TDA(.Tables.Add, "SOTPICK5", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")

            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where SOTPICK1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 1, True, "V", 1, "PICK_PICKER")

            'ASCMAIN1.sql = "Select SOTPICK5.* from SOTPICK5"
            'Create_TDA(.Tables.Add, "SOTPICK5_C", "**")

        End With

        WHSE_CODE = g.WHSE_CODE

        Dim styles As String = TACMAIN1.LookupLocation(Me, g.GUN_LOC)
        If styles.Length > 0 Then
            AppState = "LEAVE"
        End If

        tbl = dst.Tables("WHTSCANS")

        If g.PICK_TYPE = "C" Then ' Consalidated
            PickTypeSql = "SOTPICK1.PICK_NO_CONS = '"
        Else
            PickTypeSql = "SOTPICK1.PICK_NO = '"
        End If

        If g.PICK_TYPE = "S" Then 'split ticket
            'Split = true
        End If
        LockWarning = False


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
            CreateResponse("", "B", PickMessage())
        Else
            Select Case AppState
                Case "SCAN_PTCKT"
                    Dim styles As String = TACMAIN1.LookupLocation(Me, G.GUN_LOC)
                    Dim SHIP_BOL_NO As String = ""
                    If styles.Length > 0 Then
                        CreateResponse("LEAVE", "R", "Gun has Units, Deposit merchandise to continue")
                        Exit Select
                    End If
                    If SCANTEXT <> "" And SCANTEXT.StartsWith("A") Then
                        Automated = True
                        SCANTEXT = Format(Val(Mid(SCANTEXT, 2)), "0000000000")
                    End If
                    If SCANTEXT.Length < 10 Then
                        SCANTEXT = SCANTEXT.PadLeft(10, "0")
                    End If
                    If G.PICK_TYPE = "S" Then 'split ticket
                        PICK_NO = SCANTEXT.Substring(0, 7).PadLeft(10, "0")
                        PickSplit = SCANTEXT.Substring(7, 3)
                    Else
                        PICK_NO = SCANTEXT
                        PickSplit = ""
                    End If
                    ASCMAIN1.Multi_Task_Cleanup()
                    ASCMAIN1.sql = "Select SOTPICK1.PICK_STATUS, SOTPICK1.PACK_STATUS, SOTPICK1.SHIP_BOL_NO, SOTORDR1.ECOM_CODE from SOTPICK1, SOTORDR1 where " & PickTypeSql & PICK_NO & "' AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"
                    Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rows.Length > 0 Then
                        For Each row As DataRow In rows
                            If row.Item("PICK_STATUS") <> "P" Or row.Item("PACK_STATUS") & "" = "F" Or row.Item("ECOM_CODE") & "" <> "" Then
                                CreateResponse("", "GREEN", "Pick Ticket not available, ask for Help")
                                Exit Select
                            End If
                            SHIP_BOL_NO = row.Item("SHIP_BOL_NO")
                        Next
                    Else
                        CreateResponse("", "GREEN", "Invalid Pick Ticket, Re-Scan")
                        Exit Select
                    End If
                    ASCMAIN1.sql = "Select SOTSHIP1.WHSE_CODE from SOTSHIP1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                    If WHSE_CODE <> ASCDATA1.GetDataValue(ASCMAIN1.sql) Then
                        If WHSE_CODE = "MS" Then
                            WHSE_CODE = "NY"
                        Else
                            WHSE_CODE = "MS"
                        End If
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Invalid Warehouse, Please re-Scan|OK|"
                        CreateResponse("", "R", "Switching warehouse to " & WHSE_CODE)
                        AppStates(AppState) = hold
                        Exit Select
                    End If

                    ' We lock scantext because of splits else we would use pick_no
                    If Not ASCMAIN1.Logical_Lock("WHCRF016", SCANTEXT) Then
                        Dim User As String = ASCMAIN1.MultiTask_Get_Users("WHCRF016", SCANTEXT, "L")
                        If (Not User.Contains(G.USER_ID)) Or LockWarning = False Then
                            EMsg = vbCr & "Pick Ticket is Locked " _
                                     & User
                            CreateResponse("", "GREEN", EMsg)
                            'don't allow LockWarning bypass - Managers will Unlock Pick ticket in pack control update
                            If User.Contains(G.USER_ID) Then LockWarning = False 'LockWarning = True
                            Exit Select
                        End If
                    End If
                    If Not ASCMAIN1.Logical_Open("SOTPICK1", PICK_NO) Then
                        CreateResponse("", "GREEN", "Pick Ticket is Locked by the office.")
                        Exit Select
                    End If
                    LockWarning = False

                    _sql = "SELECT COUNT(DISTINCT STYLE_CODE || COLOR_CODE) FROM SOTPICK2 " & vbCrLf _
                        & " JOIN SOTORDR2 ON SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO " & vbCrLf _
                        & " JOIN SOTPICK1 ON SOTPICK2.PICK_NO = SOTPICK1.PICK_NO " & vbCrLf _
                        & " WHERE PICK_QTY > 0 AND " & PickTypeSql & PICK_NO & "'" & If(G.PICK_TYPE = "S", " AND SOTPICK2.PICK_SPLIT = '" & PickSplit & "'", "")
                    Picks = ASCDATA1.GetDataValue(_sql)
                    _sql = "Select count(distinct sotordr2.STYLE_CODE || sotordr2.COLOR_CODE) from SOTPICK2 " & vbCrLf _
                        & " join sotordr2 on sotordr2.ordr_no = SOTPICK2.ORDR_NO and sotordr2.ordr_lno = SOTPICK2.ORDR_lNO " & vbCrLf _
                        & " join ictstyc1 on sotordr2.STYLE_CODE = ictstyc1.STYLE_CODE and sotordr2.COLOR_CODE = ictstyc1.COLOR_CODE " & vbCrLf _
                        & " join sotpick1 on sotpick2.pick_no = sotpick1.pick_no " & vbCrLf _
                        & " left join (select PICK_NO, UPC_CODE, SUM(PICK_QTY) PICK_QTY from  SOTPICK5 " & vbCrLf _
                        & " where sotpick5.pick_no = '" & PICK_NO & "'" & If(G.PICK_TYPE = "S", " AND SOTPICK5.PICK_SPLIT = '" & PickSplit & "'", "") & " and sotpick5.pick_status = 'P' " & vbCrLf _
                        & " group by PICK_NO, UPC_CODE) sotpick5 on sotpick5.UPC_CODE = ictstyc1.UPC_CODE  " & vbCrLf _
                        & " where sotpick2.pick_qty > 0 And nvl(sotpick5.PICK_QTY,0) < sotpick2.PICK_QTY and  " & PickTypeSql & PICK_NO & "'" & If(G.PICK_TYPE = "S", " AND SOTPICK2.PICK_SPLIT = '" & PickSplit & "'", "")
                    OpenPicks = ASCDATA1.GetDataValue(_sql)

                    GOTO_LOCATION = ""
                    RowNum = 0

                    ASCMAIN1.sql = "Select sotordr2.STYLE_CODE, sotordr2.COLOR_CODE, ictcolr1.COLOR_DESC, ictstyc1.UPC_CODE, max(ictstyl1.STYLE_DESC) STYLE_DESC, " & vbCrLf _
                        & " sum(sotpick2.PICK_QTY) PICK_QTY, sum(sotpick2.PICK_QTY) PICK_QTY_OPEN, max(ictstyl1.INNER_PACK_QTY) INNER_PACK_QTY, " & vbCrLf _
                        & " max(ictstyl1.CARTON_PACK_QTY) CARTON_PACK_QTY, max(ictstyl1.CARTONS_PER_UNIT) CARTONS_PER_UNIT " & vbCrLf _
                        & " from sotpick1  " & vbCrLf _
                        & " join sotpick2 on sotpick1.PICK_NO = sotpick2.PICK_NO" & vbCrLf _
                        & " Join sotordr2 on sotordr2.ordr_no = sotpick2.ordr_no And sotordr2.ordr_lno = sotpick2.ORDR_LNO " & vbCrLf _
                        & " Join ictstyl1 on ictstyl1.style_code = sotordr2.style_code " & vbCrLf _
                        & " Join ictstyc1 on ictstyc1.style_code = sotordr2.style_code And ictstyc1.color_code = sotordr2.color_code " & vbCrLf _
                        & " Join ictcolr1 on ictcolr1.color_code = sotordr2.color_code " & vbCrLf _
                        & " where sotpick2.pick_qty > 0   And  " & PickTypeSql & PICK_NO & "'" & If(G.PICK_TYPE = "S", " AND SOTPICK2.PICK_SPLIT = '" & PickSplit & "'", "") & vbCrLf _
                        & " group by sotordr2.STYLE_CODE, sotordr2.COLOR_CODE, ictcolr1.COLOR_DESC, ictstyc1.UPC_CODE"

                    rows = ASCDATA1.GetDataTable.Select("")
                    If rows.Length = 0 Then
                        CreateResponse("", "GREEN", "Nothing left to Pick, scan new Pick Ticket")
                        Exit Select
                    End If

                    Assign_PickTicket()

                    dst.Tables("WHTSCANS").Rows.Clear()

                    Fill_Records("SOTPICK5", "", True, "Select * from SOTPICK5 where PICK_NO = '" & PICK_NO & If(G.PICK_TYPE = "S", "' AND SOTPICK5.PICK_SPLIT = '" & PickSplit & "'", "'") & " AND PICK_STATUS = 'P'")

                    For Each row As DataRow In rows
                        rowWHTSCANS = dst.Tables("WHTSCANS").NewRow
                        RowNum = 0
                        Dim rowLoc As DataRow = GetLocation(row.Item("STYLE_CODE"), row.Item("COLOR_CODE"), row.Item("PICK_QTY"))
                        With rowWHTSCANS
                            .Item("LOCATION_ROUTE_SEQ") = If(rowLoc Is Nothing, "999999", If(rowLoc.Item("LOCATION_ROUTE_SEQ").Equals(Null), "999999", rowLoc.Item("LOCATION_ROUTE_SEQ")))
                            .Item("LOCATION_CODE") = If(rowLoc Is Nothing, "999-99-A", rowLoc.Item("LOCATION_CODE"))
                            .Item("PICK_NO") = PICK_NO
                            .Item("PICK_LNO") = RowNum
                            .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                            .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                            .Item("COLOR_DESC") = row.Item("COLOR_DESC")
                            .Item("UPC_CODE") = row.Item("UPC_CODE")
                            .Item("STYLE_DESC") = row.Item("STYLE_DESC")
                            .Item("LOCATION_QTY") = If(rowLoc Is Nothing, "0", rowLoc.Item("LOCATION_QTY"))
                            .Item("PICK_QTY") = row.Item("PICK_QTY")
                            .Item("PICK_QTY_OPEN") = If(row.Item("PICK_QTY_OPEN").Equals(Null), 0, row.Item("PICK_QTY_OPEN"))
                            .Item("INNER_PACK_QTY") = If(row.Item("INNER_PACK_QTY").Equals(Null), 0, row.Item("INNER_PACK_QTY"))
                            .Item("CARTON_PACK_QTY") = If(row.Item("CARTON_PACK_QTY").Equals(Null), 0, row.Item("CARTON_PACK_QTY"))
                            .Item("CARTONS_PER_UNIT") = If(row.Item("CARTONS_PER_UNIT").Equals(Null), 0, row.Item("CARTONS_PER_UNIT"))
                            .Item("STYLE_ASST_QTY") = ASCDATA1.GetDataValue("Select nvl(STYLE_ASST_QTY,0) STYLE_ASST_QTY from ICTSTYL1 WHERE STYLE_CODE = :PARM1", "V", row.Item("STYLE_CODE"))
                            .Item("ROWNUM") = RowNum
                        End With
                        dst.Tables("WHTSCANS").Rows.Add(rowWHTSCANS)
                    Next
                    'Sort work table for picking Seq and match pick5 recs
                    RowNum = 0
                    Dim rowSOTPICK5 As DataRow
                    For Each rowWHTSCANS In dst.Tables("WHTSCANS").Select("", "LOCATION_ROUTE_SEQ, LOCATION_CODE")
                        RowNum += 1
                        rowWHTSCANS.Item("PICK_LNO") = RowNum
                        rowWHTSCANS.Item("ROWNUM") = RowNum
                        For Each rowSOTPICK5 In dst.Tables("SOTPICK5").Select("PICK_NO = '" & PICK_NO & "' and UPC_CODE = '" & rowWHTSCANS.Item("UPC_CODE") & "'")
                            rowWHTSCANS.Item("PICK_QTY_OPEN") = rowWHTSCANS.Item("PICK_QTY_OPEN") - rowSOTPICK5.Item("PICK_QTY")
                            'rowWHTSCANS.Item("PICK_LNO") = rowSOTPICK5.Item("PICK_LNO")
                        Next
                    Next

                    'Get first row - we're at last
                    Change_Style("NEXT UPC")
                    PageNo = 0

                    If IsNothing(STYLE_CODE) Then
                        CreateResponse("SCAN_LOC", "R", PickMessage())
                    Else
                        'Do this every time we change appstate to SCAN_LOC
                        AppStates("SCAN_LOC") = String.Format(ScanLocMsg, GOTO_LOCATION)
                        CreateResponse("SCAN_LOC", "YELLOW", PickMessage())
                    End If

                Case "SCAN_LOC"
                    If SCANTEXT = "DONE" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_PTCKT", "GREEN", "Scan New Ticket")
                        Exit Select
                    ElseIf SCANTEXT = "YES" Then
                        LOCATION_CODE = holdScan
                        GOTO_LOCATION = holdScan
                        holdScan = ""
                    ElseIf SCANTEXT = "NO" Then
                        CreateResponse("", "YELLOW", PickMessage())
                        holdScan = ""
                        Exit Select
                    ElseIf SCANTEXT = "<<" Or SCANTEXT = ">>" Then
                        'CreateResponse("", "YELLOW", Show_Locations())
                        If page = 0 Then
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
                        End If

                        Dim output As String = TACMAIN1.PaginateDataTable(tblPage, page, 4, "LOCATION_CODE,F1, LOCATION_QTY", SCANTEXT)
                        Dim msg As String = "Style " & STYLE_CODE & " Color " & COLOR_CODE & vbCrLf & output
                        CreateResponse("", "YELLOW", msg)
                        Exit Select

                    ElseIf SCANTEXT = "NEXT UPC" Or SCANTEXT = "PREV UPC" Then
                            Change_Style(SCANTEXT)
                            CreateResponse("", "YELLOW", PickMessage())
                            Exit Select
                        ElseIf SCANTEXT.ToUpper = "V" Then
                            showPicked("")
                            Exit Select
                        Else
                            If holdScan <> "" Then
                            Dim hold As String = AppStates(AppState)
                            AppStates(AppState) = "Continue With New Location " & holdScan & "|YES|NO|"
                            CreateResponse("", "R", "This is not the original location: " & ORIGINAL_LOCATION)
                            AppStates(AppState) = hold
                            Exit Select
                        Else
                            ' Current location selected
                            If SCANTEXT = "00" Then
                                'SCANTEXT = GOTO_LOCATION
                            End If

                            Dim dResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                            If dResponse.ContainsKey("Error") Then
                                Dim hold As String = AppStates(AppState)
                                AppStates(AppState) = "Invalid Location, Please re-Scan|OK|"
                                CreateResponse("", "YELLOW", dResponse("Error"))
                                AppStates(AppState) = hold
                                Exit Select
                            End If
                            If SCANTEXT <> ORIGINAL_LOCATION Then
                                Dim hold As String = AppStates(AppState)
                                'stop from using new location
                                'AppStates(AppState) = "Continue With New Location " & SCANTEXT & "|YES|NO|"
                                AppStates(AppState) = "Cannot pick from " & SCANTEXT & ", see supervisor|OK|"
                                CreateResponse("", "R", "This is not the original location: " & ORIGINAL_LOCATION)
                                AppStates(AppState) = hold
                                holdScan = SCANTEXT.ToUpper
                                Exit Select
                            End If

                            LOCATION_CODE = SCANTEXT
                            GOTO_LOCATION = SCANTEXT

                        End If
                    End If

                    CreateResponse("SCAN_UPC", "BLUE", PickMessage())

                Case "SCAN_UPC"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", PickMessage())
                        Exit Select
                    Else
                        'If G.USER_ID = "leox" AndAlso SCANTEXT = "*" Then '
                        '    'this is intended for leo to re-pick merchandise sent back(returned) for re-packaging but going back out again.
                        '    SCANTEXT = UPC_CODE
                        'End If
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        'No need to check, it will error by default 
                        'If CheckResponse.ContainsKey("Error") Then
                        '    CreateResponse("", "R", CheckResponse("Error"))
                        '    Exit Select
                        'End If

                        If CheckResponse.ContainsKey("UPC_CODE") Then
                            If UPC_CODE = CheckResponse("UPC_CODE") Then
                                If PICK_QTY_OPEN > CARTON_PACK_QTY Then
                                    CreateResponse("SCAN_CASES", "G", PickMessage())
                                Else
                                    CreateResponse("SCAN_UNITS", "G", PickMessage())
                                End If
                                Exit Select
                            End If
                        Else
                            If CheckResponse.ContainsKey("STYLE_CODE") AndAlso STYLE_CODE = CheckResponse("STYLE_CODE") Then
                                'I'm looking for one color why give a choice? do I just ask for Yes to accept?
                                'TACMAIN1.GetColors(Me, STYLE_CODE, ORIGINAL_LOCATION, COLOR_CODEs, colors)
                                'CreateResponse("SCAN_COLOR", "G", "Style " & STYLE_CODE & " has been selected, colors " & colors)
                                CreateResponse("SCAN_COLOR", "G", PickMessage())
                                Exit Select
                            End If
                        End If
                    End If
                    'Error
                    Dim hold As String = AppStates(AppState)
                    AppStates(AppState) = "Wrong UPC/Style, try again|OK|"
                    CreateResponse("", "BLUE", PickMessage())
                    AppStates(AppState) = hold

                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", PickMessage)
                        Exit Select
                    Else
                        'If SCANTEXT = "0" Then
                        '    TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)
                        '    CreateResponse("", "B", "Style " & STYLE_CODE & " has been selected, All colors " & colors)
                        '    Exit Select
                        'Else
                        '    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyColor(Me, SCANTEXT, STYLE_CODE, COLOR_CODEs)
                        '    If CheckResponse.ContainsKey("Error") Then
                        '        CreateResponse("", "R", CheckResponse("Error"))
                        '        Exit Select
                        '    End If
                        '    UPC_CODE = CheckResponse("UPC_CODE")
                        '    COLOR_CODE = CheckResponse("COLOR_CODE")
                        'End If
                        If SCANTEXT.ToUpper <> COLOR_CODE Then
                            'Error
                            Dim hold As String = AppStates(AppState)
                            AppStates(AppState) = "Wrong Color Code, Try again|OK|"
                            CreateResponse("", "R", PickMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                    End If
                    If PICK_QTY_OPEN > CARTON_PACK_QTY Then
                        CreateResponse("SCAN_CASES", "G", PickMessage())
                    Else
                        CreateResponse("SCAN_UNITS", "G", PickMessage())
                    End If

                Case "SCAN_CASES"
                    '  Asking for Cases although Message may not display number of cases requested
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", PickMessage())
                        Exit Select
                    Else
                        Dim hold As String
                        If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Number, Enter Cases Picked|OK|"
                            CreateResponse("", "R", PickMessage())
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
                            CreateResponse("SCAN_UNITS", "G", PickMessage())
                            Exit Select
                        End If

                        Dim picked As Integer = CARTON_PACK_QTY * CASES_MOVED + UNITS_MOVED
                        If picked > PICK_QTY_OPEN Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Too Many Cases/Units, Enter Cases Picked|OK|"
                            CreateResponse("", "R", PickMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = "Update " & SCANTEXT & " Cases|Y|N|CLEAR|EXIT|"
                        CreateResponse("VERIFY", "B", PickMessage())
                        AppStates("VERIFY") = hold
                    End If

                Case "SCAN_UNITS"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_LOC", "YELLOW", PickMessage())
                        Exit Select
                    Else
                        Dim hold As String
                        'Can we have more than 999 loose units to Move in a Pick?
                        If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Count " & SCANTEXT & ", Verify Units Count|OK|"
                            CreateResponse("", "R", PickMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        UNITS_MOVED = Val(SCANTEXT)
                        If UNITS_MOVED + (CARTON_PACK_QTY * CASES_MOVED) > PICK_QTY_OPEN Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Too Many Units " & SCANTEXT & ", Verify Units Count|OK|"
                            CreateResponse("", "R", PickMessage())
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = "Update " & SCANTEXT & " Units|Y|N|CLEAR|EXIT|"
                        CreateResponse("VERIFY", "B", PickMessage())
                        AppStates("VERIFY") = hold
                    End If

                Case "SCAN_VOID"
                    If SCANTEXT = "BACK" Then
                        loadLine()
                        AppStates("SCAN_LOC") = String.Format(ScanLocMsg, GOTO_LOCATION)
                        CreateResponse("SCAN_LOC", "YELLOW", PickMessage())
                        Exit Select
                    ElseIf SCANTEXT = "NEW PICK" Then
                        CreateResponse("SCAN_PTCKT", "GREEN", PickMessage())
                        Exit Select
                    ElseIf SCANTEXT = "NXT PAGE" Then
                        showPicked("", PageNo)
                        'CreateResponse("SCAN_PTCKT", "GREEN", PickMessage())
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
                                showPicked(SCANTEXT, PageNo - 1)
                                Exit Select
                            End If
                        Else
                            showPicked(holdScan, PageNo - 1)
                            Exit Select
                        End If
                    End If
                    showPicked("")

                Case "VERIFY"
                    If SCANTEXT.ToUpper = "Y" Then
                        ' Y is for normal Verify
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        If PICK_QTY_OPEN = 0 Then
                            OpenPicks = OpenPicks - 1
                            Change_Style("NEXT UPC")
                            CreateResponse("SCAN_LOC", "YELLOW", PickMessage())
                        Else
                            If PICK_QTY_OPEN > CARTON_PACK_QTY Then
                                CreateResponse("SCAN_CASES", "G", PickMessage())
                            Else
                                CreateResponse("SCAN_UNITS", "G", PickMessage())
                            End If
                        End If
                    ElseIf SCANTEXT = "YES" Then
                        ' YES is for Clear Verify
                        CreateResponse("SCAN_UPC", "BLUE", PickMessage())
                    ElseIf SCANTEXT.ToUpper = "N" Or SCANTEXT = "NO" Then
                        ' N is for normal Verify
                        ' NO is for Clear Verify
                        If PICK_QTY_OPEN > CARTON_PACK_QTY Then
                            CreateResponse("SCAN_CASES", "G", PickMessage())
                        Else
                            CreateResponse("SCAN_UNITS", "G", PickMessage())
                        End If
                    ElseIf SCANTEXT = "CLEAR" Then
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Clear Current Pick |YES|NO|"
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

    Function GetLocation(ByVal Style As String, ByVal Color As String, ByVal PICK_QTY As Int32) As DataRow
        Dim rtn_row As DataRow = Nothing
        Dim ab_row As DataRow = Nothing
        ' Dim tmp_LOC As String
        ASCMAIN1.sql = " select b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE, nvl(m1.LOCATION_USE,'A') LOCATION_USE " & vbCrLf _
            & " from whtlocb1 b1 " & vbCrLf _
            & "  join whtlocm1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE " & vbCrLf _
            & "  where b1.STYLE_CODE = '" & Style & "' and b1.COLOR_CODE = '" & Color & "' " & vbCrLf _
            & "  and  nvl(m1.LOCATION_USE,'A') in ('A','E') " & vbCrLf _
            & "  and m1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "  and NVL(m1.LOCATION_LOCKED,'0') <> '1'" & vbCrLf _
            & "  order by b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "LOCATION_QTY") ' "LOCATION_USE, LOCATION_QTY"
            If Style = "MTH13478" Then Stop

            rtn_row = row
            'If Val(row("LOCATION_QTY") & "") > 0 Then Exit For

            'previous logic tried to clean locs by taking picker to lowest qty > pick
            'Now we have largest qty first, so first rec always satisfies rule.
            'New change - 8/16/2025 - show lowest qty first so we can clear locations. 8/28/2025 revert change
            'tmp_LOC = row("LOCATION_CODE")
            'If row("LOCATION_QTY") > 0 And (tmp_LOC.Substring(tmp_LOC.Length - 1, 1) = "A" Or tmp_LOC.Substring(tmp_LOC.Length - 1, 1) = "B") Then
            '    ab_row = row
            'End If
            'If row("LOCATION_QTY") >= PICK_QTY Then
            '    If tmp_LOC.Substring(tmp_LOC.Length - 1, 1) = "C" AndAlso ab_row IsNot Nothing Then
            '        rtn_row = ab_row
            '    Else
            'Exit For
            '    End If
            'End If
        Next

        Return rtn_row

    End Function

    Sub Change_Location(ByVal SCANTEXT As String)
        'request from Leo to look for another location for item, previous logic would get the next or previous adjacent location
        Dim newLocation As String = ""
        Dim rtn_row As DataRow = Nothing
        Dim ab_row As DataRow = Nothing
        Dim tmp_LOC As String = ""
        ASCMAIN1.sql = " select b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE " & vbCrLf _
            & " from whtlocb1 b1 " & vbCrLf _
            & "  join whtlocm1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE " & vbCrLf _
            & "  where b1.STYLE_CODE = '" & STYLE_CODE & "' and b1.COLOR_CODE = '" & COLOR_CODE & "' " & vbCrLf _
            & "  and  nvl(m1.LOCATION_USE,'A') in ('A','E') " & vbCrLf _
            & "  and m1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "  and NVL(m1.LOCATION_LOCKED,'0') <> '1'" & vbCrLf _
            & "  order by b1.LOCATION_QTY desc, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"
        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("", "LOCATION_QTY")
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "LOCATION_QTY")
            rtn_row = row
            tmp_LOC = row("LOCATION_CODE")
            If tmp_LOC <> GOTO_LOCATION Then
                If row("LOCATION_QTY") > 0 And (tmp_LOC.Substring(tmp_LOC.Length - 1, 1) = "A" Or tmp_LOC.Substring(tmp_LOC.Length - 1, 1) = "B") Then
                    ab_row = row
                End If
                If row("LOCATION_QTY") >= PICK_QTY_OPEN Then
                    If tmp_LOC.Substring(tmp_LOC.Length - 1, 1) = "C" AndAlso ab_row IsNot Nothing Then
                        rtn_row = ab_row
                    Else
                        Exit For
                    End If
                End If
            End If
        Next
        If rtn_row IsNot Nothing Then
            GOTO_LOCATION = rtn_row("LOCATION_CODE")
            AppStates("SCAN_LOC") = String.Format(ScanLocMsg, GOTO_LOCATION)
        End If
    End Sub
    Function Show_Locations() As String
        Dim rc As Integer = 0
        ASCMAIN1.sql = " select b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE " & vbCrLf _
            & " from whtlocb1 b1 " & vbCrLf _
            & "  join whtlocm1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE " & vbCrLf _
            & "  where b1.STYLE_CODE = '" & STYLE_CODE & "' and b1.COLOR_CODE = '" & COLOR_CODE & "' " & vbCrLf _
            & "  and  nvl(m1.LOCATION_USE,'A') in ('A','E') " & vbCrLf _
            & "  and m1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "  order by b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"
        'Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("", "LOCATION_QTY")
        Dim msg As String = "Style " & STYLE_CODE & " - " & STYLE_DESC & vbCrLf _
            & "Color " & COLOR_DESC & vbCrLf _
            & "Pick Qty " & PICK_QTY_OPEN & vbCrLf
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "LOCATION_QTY desc")
            rc = rc + 1
            If rc > 4 Then Exit For
            msg = msg & row("LOCATION_CODE") & " #" & row("LOCATION_QTY") & vbCrLf
        Next
        If rc < 4 Then
            msg = msg & String.Format("FOUND IN {0} LOCS", rc) & vbCrLf
        End If
        msg = msg & "".PadRight(60)
        Return msg
    End Function
    Sub Change_Style(ByVal SCANTEXT)
        Dim SortOrder As String = "ROWNUM " & If(SCANTEXT = "NEXT UPC", "asc", "desc")
        Dim AutomatedClause = ""

        'clear paging table
        tblPage = Nothing
        page = 0

        If Automated Then
            AutomatedClause = "LOCATION_QTY > 0 AND "
        End If

        Dim rows() As DataRow = tbl.Select(AutomatedClause & "PICK_QTY_OPEN > 0 AND ROWNUM " & If(SCANTEXT = "NEXT UPC", "> ", "< ") & RowNum, SortOrder)


        If rows.Length = 0 Then
            rows = tbl.Select(AutomatedClause & "PICK_QTY_OPEN > 0 ", SortOrder)
        End If

        If rows.Length = 0 Then
            AppStates("SCAN_LOC") = "No open picks for ticket V:Void|DONE|"
            GOTO_LOCATION = "DONE"
        Else
            rowWHTSCANS = rows(0)
            loadLine()
            AppStates("SCAN_LOC") = String.Format(ScanLocMsg, GOTO_LOCATION)

        End If
    End Sub
    Sub loadLine()

        GOTO_LOCATION = rowWHTSCANS.Item("LOCATION_CODE")
        RowNum = rowWHTSCANS.Item("ROWNUM")
        STYLE_CODE = rowWHTSCANS.Item("STYLE_CODE")
        COLOR_CODE = rowWHTSCANS.Item("COLOR_CODE")
        COLOR_DESC = rowWHTSCANS.Item("COLOR_CODE") & ": " & rowWHTSCANS.Item("COLOR_DESC")
        STYLE_DESC = rowWHTSCANS.Item("STYLE_DESC")
        PICK_QTY_OPEN = rowWHTSCANS.Item("PICK_QTY_OPEN")
        CARTON_PACK_QTY = rowWHTSCANS.Item("CARTON_PACK_QTY")
        INNER_PACK_QTY = rowWHTSCANS.Item("INNER_PACK_QTY")
        UPC_CODE = rowWHTSCANS.Item("UPC_CODE")
        PICK_LNO = rowWHTSCANS.Item("PICK_LNO")
        CARTONS_PER_UNIT = rowWHTSCANS.Item("CARTONS_PER_UNIT")
        ORIGINAL_LOCATION = GOTO_LOCATION

    End Sub
    Function PickMessage() As String
        If OpenPicks = 0 Then
            Return "PT " & PICK_NO & " Has no Open picks "
        Else
            Dim PickQty As String = " :"
            'I feel like I should ask for inners when displaying inners
            Dim Open As Integer = PICK_QTY_OPEN
            If CARTONS_PER_UNIT > 0 Then
                PickQty = PickQty & String.Format(" {0}ttl - {1}c per unit", Math.Truncate(Open * CARTONS_PER_UNIT), CARTONS_PER_UNIT)
                Open = 0
            Else
                If CARTON_PACK_QTY > 0 And CARTON_PACK_QTY < Open Then
                    PickQty = PickQty & String.Format(" {0}c", Math.Truncate(Open / CARTON_PACK_QTY))
                    Open = Open - Math.Truncate(Open / CARTON_PACK_QTY) * CARTON_PACK_QTY
                End If
                If INNER_PACK_QTY > 0 And INNER_PACK_QTY < Open Then
                    PickQty = PickQty & String.Format(" {0}i", Math.Truncate(Open / INNER_PACK_QTY))
                    Open = Open - Math.Truncate(Open / INNER_PACK_QTY) * INNER_PACK_QTY
                End If
            End If
            If Open > 0 Then
                PickQty = PickQty & String.Format(" {0}u", Open)
            End If
            If Open = PICK_QTY_OPEN Then
                PickQty = "u"
            End If

            If rowWHTSCANS("STYLE_ASST_QTY") > 1 Then
                Dim Asst As Integer = Val(rowWHTSCANS("STYLE_ASST_QTY"))
                PickQty = PickQty & String.Format(", {0}Asrts, {1} in Assrt", Math.Truncate(PICK_QTY_OPEN / Asst), Asst)
            End If

            Return "Style " & STYLE_CODE _
                    & " - " & STYLE_DESC & vbCrLf _
                    & "Color " & COLOR_DESC & vbCrLf _
                    & "Pick Location " & GOTO_LOCATION & vbCrLf _
                    & "Pick Qty " & PICK_QTY_OPEN & PickQty & vbCrLf _
                    & "PT " & PICK_NO & " Open " & OpenPicks & "/" & Picks & vbCrLf & "".PadRight(60)
        End If
    End Function

    Sub showPicked(ByVal showline As String, Optional ByVal PAGE As Int16 = 0)
        Dim msg As String = "Picked for " & PICK_NO
        Dim lno As Integer = 0
        Dim lines As String = ""
        Dim pageLno As Integer = 0
        'ASCMAIN1.sql = "Select UPC_CODE, PICK_LNO, WHSE_TRAN_NO, PICK_CASES, PICK_UNITS, INIT_DATE from SOTPICK5 " & vbCrLf _
        '                & " where SOTPICK5.PICK_NO = '" & PICK_NO & "' and PICK_STATUS = 'P'"
        Dim rows() As DataRow = dst.Tables("SOTPICK5").Select("PICK_STATUS = 'P'", "INIT_DATE DESC")
        If PAGE * 5 > rows.Length Or PAGE < 0 Then PAGE = 0

        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                If pageLno / 5 >= PAGE Then
                    lno += 1
                    Dim prows() As DataRow = tbl.Select("UPC_CODE = '" & ROW.Item("UPC_CODE") & "'")
                    If showline = "" Or showline = "ALL" Or Val(showline) = lno Then
                        msg = msg & vbCrLf & "L" & lno & " " & ROW.Item("LOCATION_CODE") & " " & prows(0).Item("STYLE_CODE") & " " & prows(0).Item("COLOR_CODE") & " " _
                            & ROW.Item("PICK_CASES") & "c " & ROW.Item("PICK_UNITS") & "u"
                    End If
                    lines &= "|" & ROW.Item("PICK_LNO") & ":" & ROW.Item("UPC_CODE")
                    If lno = 5 Then Exit For
                End If
                pageLno += 1
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
        If showline = "" Then
            PageNo = PAGE + 1
        End If


    End Sub

    Sub voidLine(ByRef SCANTEXT As String)
        Dim rows() As DataRow
        Dim ss As String
        Dim PICK_LNO As String
        Dim UPC_CODE As String
        If SCANTEXT = "ALL" Then
            ss = "PICK_STATUS = 'P'"
        Else
            Dim hold = VoidList(Val(SCANTEXT))
            Dim void = hold.split(":")
            PICK_LNO = void(0)
            UPC_CODE = void(1)
            ss = "PICK_STATUS = 'P' and PICK_LNO = '" & void(0) & "' and UPC_CODE = '" & void(1) & "'"
        End If
        rows = dst.Tables("SOTPICK5").Select(ss)
        If rows.Length = 0 Then
            'error
        Else
            For Each rowSOTPICK5 As DataRow In rows
                'Figure out multipicks for SOTPICK5
                Dim row As DataRow = tbl.Select("UPC_CODE = '" & rowSOTPICK5.Item("UPC_CODE") & "'").First    '.Tables("SOTPICK5").Select("PICK_STATUS = 'P' AND PICK_LNO = '" & voidLine(0) & "' and UPC_CODE = '" & Void(1) & "'").FirstOrDefault
                rowSOTPICK5.Item("PICK_STATUS") = "V"

                CASES_MOVED = (rowSOTPICK5.Item("PICK_CASES") * -1)
                UNITS_MOVED = (rowSOTPICK5.Item("PICK_UNITS") * -1)
                LOCATION_CODE = rowSOTPICK5.Item("LOCATION_CODE")

                'need logic to void consalidated picks

                OpenPicks = OpenPicks + 1
                rowWHTSCANS = row
                loadLine()
                Update_Record()
            Next
        End If

    End Sub

    Sub Assign_PickTicket()

        Dim ROWs() As DataRow

        ASCMAIN1.sql = "Select * from SOTPICK5 where PICK_NO = '" & PICK_NO & "' and PICK_LNO =0"
        Fill_Records("SOTPICK5", "", True, ASCMAIN1.sql)

        ROWs = dst.Tables("SOTPICK5").Select("")
        If ROWs.Length = 0 Then

            Dim rowSOTPICK5 As DataRow = dst.Tables("SOTPICK5").NewRow
            With rowSOTPICK5
                .Item("PICK_NO") = PICK_NO
                .Item("PICK_LNO") = 0
                .Item("WHSE_TRAN_NO") = 0
                '.Item("WHSE_TRAN_LNO") = 0
                '.Item("UPC_CODE") = UPC_CODE
                '.Item("CARTON_PACK_QTY") = CARTON_PACK_QTY
                .Item("PICK_CASES") = 0
                .Item("PICK_UNITS") = 0
                .Item("PICK_QTY") = 0
                .Item("PICK_STATUS") = "A"
                '.Item("LOCATION_CODE") = LOCATION_CODE
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("PICK_SPLIT") = If(G.PICK_TYPE = "S", PickSplit, "")

            End With
            dst.Tables("SOTPICK5").Rows.Add(rowSOTPICK5)

            Update_Record_TDA("SOTPICK5")

            'PickTypeSql used for consalidated or non consalidated select
            ASCMAIN1.sql = "Select * from SOTPICK1 where " & PickTypeSql & PICK_NO & "'"
            Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

            For Each row As DataRow In dst.Tables("SOTPICK1").Rows
                row("PICK_PICKER") = G.USER_ID
            Next
            Update_Record_TDA("SOTPICK1")

        End If
    End Sub

    Sub Update_Record()
        Dim HoldLoc As String = ""
        Dim QtyMoved = (CASES_MOVED * CARTON_PACK_QTY + UNITS_MOVED)
        PICK_QTY_OPEN = PICK_QTY_OPEN - QtyMoved
        rowWHTSCANS.Item("PICK_QTY_OPEN") = PICK_QTY_OPEN

        HoldLoc = ASCDATA1.GetDataValue("select WHSE_LOC_SHP from ICTWHSE1 where whse_code = '" & WHSE_CODE & "'")

        BeginTrans()

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
        Dim WHSE_TRAN_LNO_ctr As Integer = 1

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        With rowWHTMOVE1
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_TYPE") = "M"
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("WHSE_CODE") = WHSE_CODE
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
            .Item("LOCATION_CODE_FROM") = LOCATION_CODE
            .Item("LOCATION_CODE_TO") = HoldLoc
            .Item("BAR_CODE") = "0000000000"
            .Item("WHSE_TRAN_QTY") = QtyMoved
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

        'Figure out what to do for multipicks 
        Dim rowSOTPICK5 As DataRow = dst.Tables("SOTPICK5").NewRow
        With rowSOTPICK5
            .Item("PICK_NO") = PICK_NO
            .Item("PICK_LNO") = PICK_LNO
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
            .Item("UPC_CODE") = UPC_CODE
            .Item("CARTON_PACK_QTY") = CARTON_PACK_QTY
            .Item("PICK_CASES") = CASES_MOVED
            .Item("PICK_UNITS") = UNITS_MOVED
            .Item("PICK_QTY") = CASES_MOVED * CARTON_PACK_QTY + UNITS_MOVED
            .Item("PICK_STATUS") = If(QtyMoved > 0, "P", "R")
            .Item("LOCATION_CODE") = LOCATION_CODE
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("PICK_SPLIT") = If(G.PICK_TYPE = "S", PickSplit, "")
        End With
        dst.Tables("SOTPICK5").Rows.Add(rowSOTPICK5)

        If G.PICK_TYPE = "C" Then
            'Add records to PICK5_C until qtymoved is depleted
            'problem if second pick for same style, need to know which picks have been filled
            'how to handle void logic
        End If


        Update_Record_TDA("SOTPICK5")


        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        CommitTrans()

        CASES_MOVED = 0
        UNITS_MOVED = 0

    End Sub
    Overrides Function Get_Anticipated_Next_Response() As String
        Select Case AppState
            Case "SCAN_PTCKT"
                Return ""
            Case "SCAN_LOC"
                Return GOTO_LOCATION
            Case "SCAN_UPC"
                Return UPC_CODE
            Case "SCAN_CASES"
                Return "*" & CStr(PICK_QTY_OPEN)
            Case "SCAN_UNITS"
                Return CStr(PICK_QTY_OPEN)
            Case "VERIFY"
                Return "Y"
            Case Else
                Return ""
        End Select
    End Function

End Class
