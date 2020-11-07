Public Class WHCRF023
    ' Unload Trailer Trip

    Inherits WHCRF000


    Dim LOCATION_CODE As String
    Dim BAR_CODE As String
    Dim Cases_count As Integer
    Dim INVALID_BAR_CODE As String
    Dim BAR_CODE_LOCATION As String
    Dim LOAD_NO As String
    Dim holdScan As String = ""
    Dim TICKET_STATUS As String = ""
    Dim QTY_ERR_FLG As Int32 = 0
    Dim rowWHTCRAN1 As DataRow
    Dim WHSE_CODE_FROM As String = ""



    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF023"

        AppStates.Add("SCAN_TRIP", "Scan Trip Id for Unload|EXIT|") ' Trailer..
        AppStates.Add("SCAN_CASE", "Scan a Case from Load|CANCEL|") ' This is the Case referenced and will define the pallet 
        AppStates.Add("CARTON_COUNT", "Count Correct?|Y|N|")
        AppStates.Add("SCAN_LOCATION", "Scan Deposit Location for Load|CANCEL|")

        AppStates.Add("VERIFY", "Update to Location (Y/N)|Y|N|CANCEL|")


        AppState = "SCAN_TRIP"
        LAST_CLR = ""
        WHSE_CODE_FROM = "NJE" ' For next move in 20 years change this value

        With dst
            Create_TDA(.Tables.Add, "WHTCRAN1", "*")
            Create_TDA(.Tables.Add, "WHTCRAN3", "*") ' SCANNED New
            Create_TDA(.Tables.Add, "WHTCRAN3X", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            Create_TDA(.Tables.Add, "ICTIXFR1", "*")
            Create_TDA(.Tables.Add, "ICTIXFR2", "*")
        End With

        tbl = dst.Tables("WHTCRAN3") ' New DataTable

    End Sub

    Public Overrides Function Hello() As String

        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        RESPONSE &= vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(ByVal SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        If SCANTEXT = "EXIT" Or G.WHSE_CODE <> "NJC" Then
            'NJC only for this app
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "R", "EXIT")
        Else
            Select Case AppState
                Case "SCAN_TRIP"
                    SCANTEXT = SCANTEXT.ToUpper

                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False, True, True, "T")
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "R", CheckResponse("Error"))
                        Exit Select
                    End If

                    ASCMAIN1.MultiTask_Release()
                    LOCATION_CODE = SCANTEXT
                    LOAD_NO = ""

                    ASCMAIN1.sql = "SELECT LOCATION_LOCKED FROM WHTLOCB1 B1,WHTLOCM1 M1, WHTBARC1 BC WHERE B1.WHSE_CODE = '" & WHSE_CODE_FROM & "' AND B1.LOCATION_CODE = '" & LOCATION_CODE & "' AND LOCATION_QTY <> 0" _
                        & " and B1.WHSE_CODE = M1.WHSE_CODE AND B1.LOCATION_CODE = M1.LOCATION_CODE and BC.BAR_CODE = B1.BAR_CODE and M1.LOCATION_USE = 'T'"
                    If ASCDATA1.GetDataValue & "" <> "1" Then
                        CreateResponse("", "R", "Trip is still open, Lock trip to unload")
                        Exit Select
                    End If

                    ASCMAIN1.sql = "SELECT DISTINCT COUNT(DISTINCT BC.LOAD_NO) FROM WHTLOCB1 B1,WHTLOCM1 M1, WHTBARC1 BC WHERE B1.WHSE_CODE = '" & WHSE_CODE_FROM & "' AND B1.LOCATION_CODE = '" & LOCATION_CODE & "' AND LOCATION_QTY <> 0" _
                        & " and B1.WHSE_CODE = M1.WHSE_CODE AND B1.LOCATION_CODE = M1.LOCATION_CODE and BC.BAR_CODE = B1.BAR_CODE and M1.LOCATION_USE = 'T'"
                    Dim LOAD_COUNT = ASCDATA1.GetDataValue
                    If LOAD_COUNT = 0 Then
                        CreateResponse("", "R", "No Loads in trailer, Verify Trip ID")
                        Exit Select
                    End If
                    CreateResponse("SCAN_CASE", "B", "Tralier " & SCANTEXT & " has been selected, Loads in Trailer: " & LOAD_COUNT)

                Case "SCAN_CASE"

                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_TRIP", "B", "Scan New Trailer ID")
                        Exit Select
                    End If
                    dst.Tables("WHTMOVE1").Rows.Clear()
                    dst.Tables("WHTMOVE2").Rows.Clear()
                    dst.Tables("ICTIXFR1").Rows.Clear()
                    dst.Tables("ICTIXFR2").Rows.Clear()

                    ASCMAIN1.sql = "Select C1.* from WHTCRAN1 C1, WHTCRAN3 C3 where C1.LOAD_NO = C3.LOAD_NO AND C3.CARTON_NO = '" & BAR_CODE & "'"
                    rowWHTCRAN1 = ASCDATA1.GetDataRow
                    If Not IsNothing(rowWHTCRAN1) Then
                        CreateResponse("", "R", "Carton not in trip, Pull carton out and Call Manager")
                        Exit Select
                    End If
                    If rowWHTCRAN1("LOAD_STATUS") <> "1" Then
                        CreateResponse("", "R", "Load Previously Unloaded, Call Manager")
                        Exit Select
                    End If

                    BAR_CODE = SCANTEXT
                    LOAD_NO = rowWHTCRAN1("LOAD_NO")
                    Dim rowWHTLOCB1 As DataRow
                    Dim LastBarcode As String = ""
                    BAR_CODE_LOCATION = ""
                    Cases_count = 0

                    If Not ASCMAIN1.Logical_Open("WHTCRAN1", LOAD_NO) Then
                        CreateResponse("", "R", "Another user is working on the same load,  Call Manager")
                        Exit Select
                    End If

                    ASCMAIN1.sql = "SELECT B1.*, nvl(M1.LOCATION_USE,'A') LOCATION_USE, C3.LOAD_NO FROM WHTLOCB1 B1,WHTLOCM1 M1, WHTCRAN3 C3 " _
                        & "WHERE B1.WHSE_CODE = '" & G.WHSE_CODE & "' AND C3.LOAD_NUMBER = '" & LOAD_NO & "' " _
                        & "AND LOCATION_QTY <> 0 " _
                        & "AND B1.WHSE_CODE = M1.WHSE_CODE " _
                        & "AND B1.LOCATION_CODE = M1.LOCATION_CODE " _
                        & "And B1.BAR_CODE = C3.BAR_CODE "
                    For Each rowWHTLOCB1 In ASCDATA1.GetDataTable.Select("")
                        'LOCATION_CODE = Trip Id
                        If LOCATION_CODE <> rowWHTLOCB1("LOCATION_CODE") Then
                            CreateResponse("", "B", "Load found in multiple locations, Pull Pallet to the side for Manager")
                            Record_error(BAR_CODE, "Load found in multiple locations, Pull Pallet to the side for Manager")
                            Exit Select
                        End If
                        If LastBarcode <> rowWHTLOCB1("BAR_CODE") Then
                            Cases_count += 1
                        End If
                        LastBarcode = rowWHTLOCB1("BAR_CODE")
                    Next

                    CreateResponse("SCAN_CARTONS", "B", "Count the cartons on the palette to verify load " & SCANTEXT & " has been selected." & vbCrLf & "Cartons on Pallet: " & Cases_count)

                Case "SCAN_CARTONS"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_CASE", "B", "Count Cancelled, Scan a new carton in Pallet")
                        Exit Select
                    End If
                    If Val(SCANTEXT) = "N" Then
                        CreateResponse("", "R", "Count Does not match, pull Pallet aside for Manager")
                        Exit Select
                    End If
                    CreateResponse("SCAN_LOCATION", "B", "Scan New Location for Pallet" & BAR_CODE & ", Cartons: " & Cases_count)

                Case "SCAN_LOCATION"
                    SCANTEXT = SCANTEXT.ToUpper

                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False, True, True, "")
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "R", CheckResponse("Error"))
                        Exit Select
                    End If
                    If CheckResponse("Stylelist") <> "" Then
                        CreateResponse("", "R", "Location " & BAR_CODE_LOCATION & " Shows Cartons, try again or call Manager")
                        Exit Select
                    End If

                    BAR_CODE_LOCATION = SCANTEXT

                    ASCMAIN1.sql = "SELECT NVL(M1.LOCATION_USE,'A') FROM WHTLOCM1 M1 WHERE M!.WHSE_CODE = '" & G.WHSE_CODE & "' AND M1.LOCATION_CODE = '" & BAR_CODE_LOCATION & "' "
                    Dim LOCATION_USE = ASCDATA1.GetDataValue
                    If LOCATION_USE > 0 Then
                        CreateResponse("", "R", "Location " & BAR_CODE_LOCATION & " Not an Inventory location, try again or call Manager")
                        Exit Select
                    End If

                    CreateResponse("VERIFY", "B", "Update Load to location " & SCANTEXT & " for " & Cases_count & " Cartons")

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_TRIP", "B", "Load Updated, Scan Trip Id for Next Load")
                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_CASE", "B", "Load Cancelled, Scan New Load")
                    ElseIf SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_CASE", "B", "Load Cancelled, Scan New Load")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
                    'AppStates("SCAN_TRIP") = hold
            End Select
        End If
    End Sub

    Sub Record_error(BARCODE, ERROR_MSG)

        ASCMAIN1.sql = "Insert into WHTTRIP2 (WHSE_CODE, TRIP_ID, BAR_CODE, ERROR_MSG, TRANS_TYPE, GUN_ID, INIT_OPER, INIT_DATE) values ('" _
            & G.WHSE_CODE & "','" & LOCATION_CODE & "','" & BARCODE & "','" & ERROR_MSG & "','LOADING','" & G.GUN_LOC & "','" & G.USER_ID & "', sysdate)"
        ASCDATA1.ExecuteSQL()

    End Sub


    Sub Update_Record()

        BeginTrans()

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
        Dim WHSE_TRAN_LNO_ctr As Integer = 0

        'Move from location on incoming Xfers
        ASCMAIN1.sql = "select WHSE_LOC_XIN from ICTWHSE1 where whse_CODE = '" & G.WHSE_CODE & "'"
        Dim XfrLoc As String = ASCDATA1.GetDataValue


        'Cleanup code for move to BAR_CODE_LOCATION (Deposit Location)
        ASCMAIN1.sql = "Update WHTCRAN1 " & vbCrLf _
            & " set LOAD_STATUS= '2', LAST_DATE = sysdate, LAST_OPER = '" & G.USER_ID & "', GUN_LOC_NJC = '" & BAR_CODE_LOCATION & "' " & vbCrLf _
            & " Where LOAD_NO = '" & LOAD_NO & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

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

        'loop through cartons and styles in load
        Dim rowWHTLOCB1 As DataRow
        ASCMAIN1.sql = "SELECT B1.*, nvl(BC.LOAD_NO,'X') LOAD_NO_FROM FROM WHTLOCB1 B1, WHTBARC1 BC, WHTCRAN3 C3 " _
                        & "WHERE B1.WHSE_CODE = '" & WHSE_CODE_FROM & "' AND C3.LOAD_NO = '" & LOAD_NO & "' " _
                        & "AND LOCATION_QTY <> 0 " _
                        & "And B1.BAR_CODE = C3.CARTON_NO " _
                        & "And BC.BAR_CODE = C3.CARTON_NO "
        For Each rowWHTLOCB1 In ASCDATA1.GetDataTable.Select("")
            WHSE_TRAN_LNO_ctr += 1
            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
            With rowWHTMOVE2
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                .Item("LOCATION_CODE_FROM") = XfrLoc ' XferIn Loc - We need to do Xfer before Move, else it's still in old whse
                .Item("LOCATION_CODE_TO") = BAR_CODE_LOCATION ' Deposit Location
                .Item("BAR_CODE") = rowWHTLOCB1("BAR_CODE")
                .Item("WHSE_TRAN_QTY") = rowWHTLOCB1("LOCATION_QTY")
                .Item("STYLE_CODE") = rowWHTLOCB1("STYLE_CODE")
                .Item("COLOR_CODE") = rowWHTLOCB1("COLOR_CODE")
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"
                .Item("LOAD_NO_FROM") = LOAD_NO
                .Item("LOAD_NO_TO") = LOAD_NO
            End With
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

            'end loop
        Next
        Update_Record_TDA("WHTMOVE2")

        'need to write to ICTIXFR1, ICTIXFR2  - execute ICTIXFRI & ICTIXFRG
        Dim Transfer_No As String = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
        Dim rowICTIXFR1 As DataRow = dst.Tables("ICTIXFR1").NewRow
        With rowICTIXFR1
            .Item("XFR_NO") = Transfer_No
            .Item("WHSE_CODE") = WHSE_CODE_FROM
            .Item("WHSE_CODE_TO") = G.WHSE_CODE
            .Item("XFR_DATE") = DATETIME_STAMP
            .Item("XFR_SOURCE") = "E"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = G.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("REGISTER_IND") = "0"
            .Item("JOURNAL_IND") = "0"
        End With
        dst.Tables("ICTIXFR1").Rows.Add(rowICTIXFR1)

        For Each rowWHTMOVE2 As DataRow In dst.Tables("WHTMOVE2").Select
            Dim rowICTIXFR2 As DataRow = dst.Tables("ICTIXFR2").NewRow
            With rowICTIXFR2
                .Item("XFR_NO") = Transfer_No
                .Item("XFR_LNO") = rowWHTMOVE2.Item("WHSE_TRAN_LNO")
                .Item("STYLE_CODE") = rowWHTMOVE2.Item("STYLE_CODE")
                .Item("COLOR_CODE") = rowWHTMOVE2.Item("COLOR_CODE")
                .Item("XFR_QTY") = rowWHTMOVE2.Item("WHSE_TRAN_QTY")
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("LOCATION_CODE") = rowWHTMOVE2.Item("LOCATION_CODE_FROM")
                .Item("BAR_CODE") = rowWHTMOVE2.Item("BAR_CODE")
                .Item("LOAD_NO") = rowWHTMOVE2.Item("LOAD_NO_FROM")

                ASCMAIN1.sql = "Select * from ICTSTYL1 " _
                & " where STYLE_CODE = '" & rowWHTMOVE2.Item("STYLE_CODE") & "" & "'"
                Dim rowICTSTYL1 As DataRow = ASCDATA1.GetDataRow
                If rowICTSTYL1 IsNot Nothing Then
                    .Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
                    .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                    .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                End If
                dst.Tables("ICTIXFR2").Rows.Add(rowICTIXFR2)
            End With
        Next
        Update_Record_TDA("ICTIXFR1")
        Update_Record_TDA("ICTIXFR2")

        ASCDATA1.ExecuteSP("ICPIXFRI", "VN", New Object() {Transfer_No, 1}, New String() {"XFR_NO_in", "S"})
        ASCDATA1.ExecuteSP("ICPIXFRG", "V", New Object() {Transfer_No}, New String() {"XFR_NO_in"})

        ASCDATA1.ExecuteSP("WHPLOCB2",
                   "VVV",
                   New String() {"B", XFR_NO_in, ASCMAIN1.SESSION_NO},
                   New String() {"WHSE_TRAN_TYPE_IN", "WHSE_TRAN_NO_IN", "SESSION_NO_IN"})

        ASCDATA1.ExecuteSP("WHPLOCB2",
                   "VVV",
                   New String() {"C", XFR_NO_in, ASCMAIN1.SESSION_NO},
                   New String() {"WHSE_TRAN_TYPE_IN", "WHSE_TRAN_NO_IN", "SESSION_NO_IN"})




        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        CommitTrans()


    End Sub


End Class

