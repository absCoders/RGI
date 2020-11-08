Public Class WHCRF022
    ' Load Trailer Trip

    Inherits WHCRF000


    Dim LOCATION_CODE As String
    Dim BAR_CODE As String
    Dim Cases_count As Integer
    Dim INVALID_BAR_CODE As String
    Dim LOAD_NO As String  ' new load 
    Dim LOAD_NO_OTHER As String ' original load no for lead carton
    Dim holdScan As String = ""
    Dim TICKET_STATUS As String = ""
    Dim QTY_ERR_FLG As Int32 = 0
    Dim WHSE_TRAN_NO As String = ""
    Dim WHSE_TRAN_LNO_ctr As Integer = 0
    Dim LoadCartonCount As Integer = 0
    Dim FIRST As Boolean = True

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF022"

        AppStates.Add("SCAN_TRIP", "Scan Trip ID|EXIT|") ' Trailer..
        AppStates.Add("SCAN_CASE", "Scan Carton BARCODE |DONE|CANCEL|") ' This is the Case referenced and will define the pallet 

        AppStates.Add("VERIFY", "Update Load on Trip (Y/N)|Y|N|CANCEL|")


        AppState = "SCAN_TRIP"
        LAST_CLR = ""

        With dst

            Create_TDA(.Tables.Add, "WHTCRAN1", "*")

            ASCMAIN1.sql = "Select Distinct WHTBARC1.BAR_CODE, WHTBARC1.LOAD_NO" & vbCrLf _
                            & " from WHTLOCB1, WHTBARC1" & vbCrLf _
                            & " where WHTBARC1.LOAD_NO = :PARM1" & vbCrLf _
                            & "   and WHTLOCB1.BAR_CODE = WHTBARC1.BAR_CODE" & vbCrLf _
                            & "   and WHTLOCB1.WHSE_CODE = :PARM2" & vbCrLf _
                            & "   and WHTLOCB1.LOCATION_QTY <> 0"
            Create_TDA(.Tables.Add, "WHTCRAN2", "**", 0, True, "VV") ' A list of cartons belonging to the Load of the 1st Carton

            ASCMAIN1.sql = "Select WHTBARC1.BAR_CODE, WHTBARC1.LOAD_NO, WHTLOCB1.WHSE_CODE, WHTLOCB1.LOCATION_CODE, WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE, WHTLOCB1.LOCATION_QTY" & vbCrLf _
                            & " from WHTLOCB1, WHTBARC1" & vbCrLf _
                            & " where WHTBARC1.LOAD_NO = :PARM1" & vbCrLf _
                            & "   and WHTLOCB1.BAR_CODE = WHTBARC1.BAR_CODE" & vbCrLf _
                            & "   and WHTLOCB1.WHSE_CODE = :PARM2" & vbCrLf _
                            & "   and WHTLOCB1.LOCATION_QTY <> 0"
            Create_TDA(.Tables.Add, "WHTCRAN2X", "**", 0, True, "VV") ' A list of cartons belonging to the Load of the 1st Carton

            Create_TDA(.Tables.Add, "WHTCRAN3", "*") ' A list of cartons scanned into the current pallet (new load)

            ASCMAIN1.sql = "Select * from WHTCRAN3"
            Create_TDA(.Tables.Add, "WHTCRAN3Z", "**", 1, False) ' To verify that a scanned carton has not already been loaded in a previous trip


            Create_TDA(.Tables.Add, "WHTCRAN3X", "*")
            Create_TDA(.Tables.Add, "WHTBARC1", "*", , False)

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            Create_TDA(.Tables.Add, "WHTBARC0", "*")

            ASCMAIN1.sql = "Select WHTLOCB1.*" & vbCrLf _
            & " from WHTLOCB1" & vbCrLf _
            & " where WHTLOCB1.WHSE_CODE = :PARM1 and WHTLOCB1.BAR_CODE = :PARM2" & vbCrLf _
            & "   and WHTLOCB1.LOCATION_QTY <> 0"
            Create_TDA(.Tables.Add, "WHTLOCB1", "**", 0, False, "VV")

            ASCMAIN1.sql = "Select WHTLOCB1.*, nvl(WHTBARC1.LOAD_NO,'X') LOAD_NO_FROM" & vbCrLf _
                & " From WHTLOCB1,  WHTBARC1, WHTCRAN3" & vbCrLf _
                & " Where WHTLOCB1.WHSE_CODE = :PARM1 And WHTCRAN3.LOAD_NO = :PARM2 " & vbCrLf _
                & " and WHTLOCB1.LOCATION_QTY <> 0" & vbCrLf _
                & " and WHTLOCB1.BAR_CODE = WHTCRAN3.BAR_CODE" & vbCrLf _
                & " and WHTBARC1.BAR_CODE = WHTCRAN3.BAR_CODE"
            Create_TDA(.Tables.Add, "WHTLOCBX", "**", 0, False, "VV")

        End With

        tbl = dst.Tables("WHTCRAN3") ' tbl reference for viewing a selected DataTable in ABSolution

    End Sub

    Public Overrides Function Hello() As String

        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        RESPONSE &= vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(ByVal SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        If SCANTEXT = "EXIT" Or G.WHSE_CODE <> "NJE" Then
            SCANTEXT = "EXIT"
            'Only NJE for this app
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "R", "EXIT")
        Else
            Select Case AppState
                Case "SCAN_TRIP"
                    SCANTEXT = SCANTEXT.ToUpper

                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, True, False, True, "T")
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "R", CheckResponse("Error"))
                        Exit Select
                    End If

                    LOCATION_CODE = SCANTEXT ' The Trip
                    WHSE_TRAN_LNO_ctr = 0
                    LoadCartonCount = 0

                    LOAD_NO_OTHER = "" ' Original Load for Lead carton - empty until 1st carton is scanned and validated

                    ASCMAIN1.MultiTask_Release()

                    For Each TABLE_NAME As String In New String() _
                        {"WHTMOVE1", "WHTMOVE2", "WHTCRAN1",
                        "WHTCRAN2", "WHTCRAN2X", "WHTCRAN3", "WHTCRAN3X"}
                        dst.Tables(TABLE_NAME).Rows.Clear()
                    Next

                    ASCMAIN1.sql = "Select DISTINCT COUNT(DISTINCT BC.LOAD_NO)" & vbCrLf _
                        & " from WHTLOCB1 B1,WHTLOCM1 M1, WHTBARC1 BC" & vbCrLf _
                        & " where B1.WHSE_CODE = :PARM1 and B1.LOCATION_CODE = :PARM2 AND LOCATION_QTY <> 0" _
                        & " and B1.WHSE_CODE = M1.WHSE_CODE and B1.LOCATION_CODE = M1.LOCATION_CODE and BC.BAR_CODE = B1.BAR_CODE and M1.LOCATION_USE = 'T'"
                    Dim LOAD_COUNT As Int32 = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New Object() {G.WHSE_CODE, LOCATION_CODE}) ' Shows the number of pallets loaded in to the  trip so far

                    CreateResponse("SCAN_CASE", "B", "Trailer " & SCANTEXT & " selected; Loads in Trailer so far: " & LOAD_COUNT)

                Case "SCAN_CASE"

                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_TRIP", "B", "Scan New Trailer Barcode")
                        Record_error(BAR_CODE, "Scan New Trailer Barcode")
                        Exit Select
                    End If
                    If SCANTEXT = "DONE" Then
                        CreateResponse("VERIFY", "B", "Verify load Update - , Cartons: " & LoadCartonCount)
                        Record_error(BAR_CODE, "Verify load Update - , Cartons: " & LoadCartonCount)
                        Exit Select
                    End If

                    BAR_CODE = SCANTEXT

                    'Error check section
                    If tbl.Rows.Find(BAR_CODE) IsNot Nothing Then
                        CreateResponse("", "B", String.Format("Duplicate Carton {0}, skipped scan", BAR_CODE))
                        Record_error(BAR_CODE, String.Format("Duplicate Carton {0}, skipped scan", BAR_CODE))
                        Exit Select
                    End If

                    If Not ASCMAIN1.Logical_Lock("WHTLOCB1", BAR_CODE) Then
                        CreateResponse("", "R", "Another user is working on the same Carton,  Call Manager")
                        Record_error(BAR_CODE, "Another user is working on the same Carton,  Call Manager")
                        Exit Select
                    End If

                    ' a carton may not have active records in any locations
                    Dim rowWHTBARC1 As DataRow = Fill_Record("WHTBARC1", BAR_CODE)
                    'Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", BAR_CODE)
                    If rowWHTBARC1 Is Nothing Then
                        CreateResponse("", "R", String.Format("Invalid Carton Scanned: {0}, Call Manager", BAR_CODE))
                        Record_error(BAR_CODE, String.Format("Invalid Carton Scanned: {0}, Call Manager", BAR_CODE))
                        Exit Select
                    End If
                    If Fill_Record("WHTCRAN3Z", BAR_CODE) IsNot Nothing Then
                        CreateResponse("", "R", String.Format("Carton Scanned: {0} was part of a previous Trip-Load, Call Manager", BAR_CODE))
                        Record_error(BAR_CODE, String.Format("Carton Scanned: {0} was part of a previous Trip-Load, Call Manager", BAR_CODE))
                        Exit Select
                    End If

                    Dim rowWHTLOCB1 As DataRow = Nothing
                    Fill_Records("WHTLOCB1", New String() {G.WHSE_CODE, BAR_CODE})
                    If dst.Tables("WHTLOCB1").Select("LOCATION_QTY_WAVE <> 0").Length > 0 Then
                        CreateResponse("", "R", "Carton Flagged for Wave, Pull carton to the side, Do not load carton")
                        Record_error(BAR_CODE, "Carton Flagged for Wave, Pull carton to the side, Do not load carton")
                        Exit Select
                    End If

                    ' a carton may be erroneously in multiple locations
                    ' a carton may be in multiple warehouses
                    ASCMAIN1.sql = "Select COUNT(DISTINCT WHTLOCB1.WHSE_CODE || WHTLOCB1.LOCATION_CODE)" & vbCrLf _
                        & " from WHTLOCB1" & vbCrLf _
                        & " where WHTLOCB1.BAR_CODE = :PARM1 AND WHTLOCB1.LOCATION_QTY <> 0"
                    Dim LOCATION_COUNT As Int32 = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {BAR_CODE}) ' Find all active locations for the barcode
                    If LOCATION_COUNT > 1 Then
                        CreateResponse("", "R", "Carton Found in multiple locations, Pull carton to the side, Do not load carton")
                        Record_error(BAR_CODE, "Carton Found in multiple locations,Pull carton to the side, Do not load carton")
                        Exit Select
                    End If

                    'End Error Check

                    If LOAD_NO_OTHER = "" Then ' this must be the 1st carton
                        LOAD_NO_OTHER = rowWHTBARC1.Item("LOAD_NO")

                        LOAD_NO = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
                        WHSE_TRAN_NO = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

                        ' GET ALL CARTONS BELONGING TO THE LOAD OF THE 1ST CARTON INTO CRAN2
                        ' GET ALL CARTON CONTENTS FOR ALL CARTONS BELONGING TO THE LOAD OF THE 1ST CARTON INTO CRAN2X

                        Fill_Records("WHTCRAN2", New String() {LOAD_NO_OTHER, G.WHSE_CODE})
                        Fill_Records("WHTCRAN2X", New String() {LOAD_NO_OTHER, G.WHSE_CODE})

                        Dim rowWHTCRAN1 As DataRow = dst.Tables("WHTCRAN1").NewRow
                        With rowWHTCRAN1
                            .Item("LOAD_NO") = LOAD_NO      ' NEW LOAD
                            .Item("BAR_CODE") = BAR_CODE   ' LEAD CARTON
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("INIT_OPER") = G.USER_ID
                            '.Item("LAST_DATE") = LOAD_NO
                            '.Item("LAST_OPER") = LOAD_NO
                            .Item("GUN_LOC_NJE") = G.GUN_LOC
                            '.Item("GUN_LOC_NJC") = LOAD_NO
                            .Item("LOAD_STATUS") = "1"
                            .Item("TRIP_NO") = LOCATION_CODE
                            .Item("LOAD_NO_ORIG") = LOAD_NO_OTHER
                        End With
                        dst.Tables("WHTCRAN1").Rows.Add(rowWHTCRAN1)

                    End If

                    Dim Styles As String = ""
                    Cases_count = 0
                    ' WRITE TO WHTCRAN3	 AND WHTCRAN3X for every scan  
                    FIRST = True
                    For Each rowWHTLOCB1 In dst.Tables("WHTLOCB1").Select("")
                        If FIRST = True Then
                            Dim rowWHTCRAN3 As DataRow = dst.Tables("WHTCRAN3").NewRow
                            With rowWHTCRAN3
                                .Item("BAR_CODE") = rowWHTLOCB1("BAR_CODE") & ""
                                .Item("LOAD_NO") = LOAD_NO
                            End With
                            tbl.Rows.Add(rowWHTCRAN3)
                            FIRST = False
                        End If
                        Dim rowWHTCRAN3X As DataRow = dst.Tables("WHTCRAN3X").NewRow
                        With rowWHTCRAN3X
                            .Item("BAR_CODE") = rowWHTLOCB1("BAR_CODE") & ""
                            .Item("LOAD_NO") = LOAD_NO
                            .Item("STYLE_CODE") = rowWHTLOCB1("STYLE_CODE") & ""
                            .Item("COLOR_CODE") = rowWHTLOCB1("COLOR_CODE") & ""
                            .Item("LOCATION_QTY") = rowWHTLOCB1("LOCATION_QTY") & ""
                        End With
                        dst.Tables("WHTCRAN3X").Rows.Add(rowWHTCRAN3X)
                    Next

                    LoadCartonCount = tbl.Compute("count(BAR_CODE)", "")
                    Dim Msg As String = "Trip ID:" & LOCATION_CODE & vbCrLf & "Last BARCODE: " & BAR_CODE & vbCrLf & "Current Load Carton Count: " & LoadCartonCount
                    CreateResponse("", "B", Msg)

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_TRIP", "B", "Load Updated, Scan Trailer ID")
                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_CASE", "B", "No Update, Scan Next Carton")
                    ElseIf SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_TRIP", "B", "Load Cancelled, Scan Trailer ID")
                        Record_error(BAR_CODE, "Load Cancelled :'Cancel', Scan Trailer ID")
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
        Dim LoadsOnPallet As String = ""
        WHSE_TRAN_LNO_ctr = 0

        Update_Record_TDA("WHTCRAN1")
        Update_Record_TDA("WHTCRAN2")
        Update_Record_TDA("WHTCRAN3")
        Update_Record_TDA("WHTCRAN2X")
        Update_Record_TDA("WHTCRAN3X")

        Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
        With rowWHTBARC0
            .Item("LOAD_NO") = LOAD_NO
            .Item("WHSE_CODE") = G.WHSE_CODE
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = G.USER_ID
            .Item("LOAD_STATUS") = "A"
            .Item("LOCATION_CODE") = LOCATION_CODE
            .Item("TRAN_TYPE") = "M"
            .Item("TRAN_NO") = WHSE_TRAN_NO
            .Item("LOAD_DATE") = DATETIME_STAMP.Date
        End With
        dst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
        Update_Record_TDA("WHTBARC0")

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
        Dim rowWHTLOCBX As DataRow
        Fill_Records("WHTLOCBX", New Object() {G.WHSE_CODE, LOAD_NO})
        For Each rowWHTLOCBX In dst.Tables("WHTLOCBX").Select("")
            WHSE_TRAN_LNO_ctr += 1
            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
            With rowWHTMOVE2
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                .Item("LOCATION_CODE_FROM") = rowWHTLOCBX("LOCATION_CODE")
                .Item("LOCATION_CODE_TO") = LOCATION_CODE ' Trip ID
                .Item("BAR_CODE") = rowWHTLOCBX("BAR_CODE")
                .Item("WHSE_TRAN_QTY") = rowWHTLOCBX("LOCATION_QTY")
                .Item("STYLE_CODE") = rowWHTLOCBX("STYLE_CODE")
                .Item("COLOR_CODE") = rowWHTLOCBX("COLOR_CODE")
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"
                .Item("LOAD_NO_FROM") = rowWHTLOCBX("LOAD_NO_FROM")
                .Item("LOAD_NO_TO") = LOAD_NO
            End With
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
        Next
        Update_Record_TDA("WHTMOVE2")

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select BAR_CODE, LOAD_NO_TO, LOAD_NO_FROM, INIT_OPER, INIT_DATE " & vbCrLf _
            & "   from WHTMOVE2" & vbCrLf _
            & "   where WHTMOVE2.WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update WHTBARC1 Set LOAD_NO = R1.LOAD_NO_TO, LAST_OPER = R1.INIT_OPER, LAST_DATE = R1.INIT_DATE " & vbCrLf _
            & "    where BAR_CODE = R1.BAR_CODE" & vbCrLf _
            & "      and LOAD_NO = R1.LOAD_NO_FROM;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        CommitTrans()


    End Sub


End Class

