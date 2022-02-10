Public Class WHCRF006
    ' Process Wave Instruction (Piece Pick)

    Inherits WHCRF000

    Dim WAVE_TYPE As String
    Dim CUST_CODE As String
    Dim ORDR_CUST_PO As String
    Dim BAR_CODE_pick_into As String = ""
    Dim BAR_CODE_pick_from As String
    Dim BAR_CODE_waved As String
    Dim WAVE_INST_NO As String
    Dim WAVE_NO As String
    Dim LOCATION_CODE As String
    Dim CASES As Integer
    Dim UNITS As Integer
    Dim PICKED As Integer
    Dim LOAD_NO As String = ""

    Dim LOCATION_QTY_on_hand As Integer

    Dim WAVE_INST_TEXT As String
    Dim SCANNED_INSTRUCTION As String
    Dim LAST_SEQ_NO As String = ""
    Dim OPEN_PICKS As String


    Dim rowWHTINST1 As DataRow

    Sub New(g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF006"

        AppStates.Add("NEXT_INST", "Enter Wave No or Press Enter for Next|EXIT|")
        AppStates.Add("SCAN_NEW", "Scan New Case ID to Pick into|ZERO|CANCEL|EXIT|")
        AppStates.Add("SCAN_LPN", "Scan Case ID to Pick from|CANCEL|EXIT|")
        AppStates.Add("QTY_PICKED", "Enter Qty Picked|CANCEL|EXIT|")
        AppStates.Add("REMOVE_LPN", "Remove Case? (Y/N)|Y|N|")
        AppStates.Add("VERIFY", "Pick Complete? (Y/N)|Y|N|CANCEL|")
        AppStates.Add("VERIFY_ZERO", "Zero Pick? (Y/N)|Y|N|CANCEL|")

        AppState = "NEXT_INST"

        With dst
            Create_TDA(.Tables.Add, "WHTINST2", "*", 1)
            Create_TDA(.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(.Tables.Add, "WHTBARC0", "*")
            Create_TDA(dst.Tables.Add, "WHTBARC1", "*")
        End With

        tbl = dst.Tables("WHTINST2") ' New DataTable

    End Sub

    Public Overrides Function Hello() As String
        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        If AppState_initial = "" Then AppState_initial = AppState
        RESPONSE &= vbCrLf & vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        If SCANTEXT = "EXIT" Then
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "R", "EXIT")
        Else

            Select Case AppState

                Case "NEXT_INST"

                    Dim WAVE_INST_NO_preferred As String = ""
                    If SCANTEXT <> "" And SCANTEXT.StartsWith("I") Then
                        WAVE_INST_NO_preferred = Format(Val(Mid(SCANTEXT, 2)), "0000000000")
                        ASCMAIN1.sql = "Select WAVE_NO from WHTINST1 where WAVE_INST_NO = '" & WAVE_INST_NO_preferred & "'"
                        SCANTEXT = ASCDATA1.GetDataValue
                        SCANNED_INSTRUCTION = WAVE_INST_NO_preferred
                        LAST_SEQ_NO = ""
                    End If

                    If SCANTEXT <> "" Then  'this means we have a new wave need to scan a new LPN
                        BAR_CODE_pick_into = ""
                    End If

                    If SCANTEXT = "" And WAVE_NO <> "" Then
                        SCANTEXT = WAVE_NO
                    End If
                    If SCANTEXT = "" Then
                        CreateResponse("", "R", "You Need to Enter a Wave No")
                        Exit Select
                    End If

                    WAVE_NO = Format(Val(SCANTEXT), "0000000000")


                    ASCMAIN1.sql = "Select count(1) from WHTINST1 " & vbCrLf _
                        & "where WHTINST1.WAVE_PICK_TYPE = '" & G.PICK_TYPE & "'" & vbCrLf _
                        & "   and WHTINST1.WAVE_INST_STATUS = '0'" & vbCrLf _
                        & "   and WHTINST1.WAVE_NO = '" & WAVE_NO & "'" & vbCrLf
                    OPEN_PICKS = ASCDATA1.GetDataValue


                    'ASCMAIN1.sql = "select wave_no from whtlocb1 l1 " _
                    '    & "join whtinst2 i2 on l1.bar_code = i2.bar_code " _
                    '    & "join whtinst1 i1 on i1.wave_inst_no = i2.wave_inst_no " _
                    '    & "where l1.LOCATION_CODE = '" & G.GUN_LOC & "' " _
                    '    & "and l1.location_qty > 0 " _
                    '    & "and rownum =1"

                    'Dim WAVE_NO_OPEN As String = ASCDATA1.GetDataValue

                    'If WAVE_NO <> WAVE_NO_OPEN Then
                    '    ' Open Wave exists, they need to deposit or continue with previous wave
                    '    CreateResponse("", "R", "Wave " & WAVE_NO_OPEN & ", is Active, cannot pick wave " & WAVE_NO & " until deposit")
                    '    Exit Select
                    'End If

                    ASCMAIN1.sql = "Select WAVE_INST_NO from ( " & vbCrLf _
                            & "select WHTINST1.WAVE_INST_NO, WHTINST1.LOCATION_CODE" & vbCrLf _
                            & "from WHTINST1,WHTWAVE1,WHTLOCM1" & vbCrLf _
                            & " where WHTINST1.WAVE_PICK_TYPE = '" & G.PICK_TYPE & "'" & vbCrLf _
                            & "   and WHTINST1.WAVE_INST_STATUS = '0'" & vbCrLf _
                            & "   and WHTWAVE1.WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                            & "   and WHTINST1.WAVE_NO = WHTWAVE1.WAVE_NO" & vbCrLf _
                            & "   and WHTINST1.WAVE_NO = '" & WAVE_NO & "'" _
                            & "   and WHTLOCM1.LOCATION_CODE = WHTINST1.LOCATION_CODE " & vbCrLf _
                            & "   and WHTLOCM1.WHSE_CODE = WHTWAVE1.WHSE_CODE " & vbCrLf _
                            & IIf(WAVE_INST_NO_preferred = "", "", "   and WHTINST1.WAVE_INST_NO = '" & WAVE_INST_NO_preferred & "'" & vbCrLf) _
                            & IIf(LAST_SEQ_NO = "", "", "   and WHTLOCM1.LOCATION_ROUTE_SEQ >= '" & LAST_SEQ_NO & "'" & vbCrLf) _
                            & "   order by WHTLOCM1.LOCATION_ROUTE_SEQ, WHTLOCM1.LOCATION_CODE) " & vbCrLf _
                            & "Where RowNum = 1"

                    WAVE_INST_NO = ASCDATA1.GetDataValue

                    If WAVE_INST_NO = "" Then
                        If Val(OPEN_PICKS) = 0 Then
                            CreateResponse("", "R", "No Picks Available for Wave (" & SCANTEXT & ")")
                            Exit Select
                        Else
                            If WAVE_INST_NO_preferred <> "" Then
                                CreateResponse("", "R", "Instruction not Available, Scan New Instruction (" & SCANTEXT & ")")
                                LAST_SEQ_NO = ""
                                Exit Select
                            Else
                                CreateResponse("", "R", "No more Picks Available, Hit Enter to check Wave (" & SCANTEXT & ")")
                                LAST_SEQ_NO = ""
                                Exit Select
                            End If

                        End If
                    End If





                    'ASCMAIN1.sql = "Select case when WHTWAVE1.WAVE_TYPE = 'W' then 'WorkOrdr' else  SOTORDR0.CUST_CODE END CUST_CODE, " _
                    '        & "case when WHTWAVE1.WAVE_TYPE = 'W' then 'WorkOrdr' else SOTORDR0.ORDR_CUST_PO end ORDR_CUST_PO, " _
                    '        & " WHTINST1.WAVE_NO, WHTINST1.LOCATION_CODE " _
                    '        & "from WHTINST1 " _
                    '        & "join WHTWAVE1 on WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO " _
                    '        & "left join whtwave3 on WHTWAVE3.WAVE_NO = WHTINST1.WAVE_NO " _
                    '        & "left join SOTSHIP1 on SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO " _
                    '        & "left join SOTORDR0 on SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                    '        & " where WHTINST1.WAVE_INST_NO = '" & WAVE_INST_NO & "'"
                    ASCMAIN1.sql = "Select SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                        & ", WHTINST1.WAVE_NO, WHTINST1.LOCATION_CODE, WHTWAVE1.WAVE_TYPE" & vbCrLf _
                        & " from WHTINST1" & vbCrLf _
                        & " join WHTWAVE1 on WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                        & " left join WHTWAVE3 on WHTWAVE3.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                        & " left join SOTSHIP1 on SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                        & " left join SOTORDR0 on SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                        & " where WHTINST1.WAVE_INST_NO = '" & WAVE_INST_NO & "'"
                    Dim rowCUSTPO As DataRow = ASCDATA1.GetDataRow
                    WAVE_TYPE = rowCUSTPO.Item("WAVE_TYPE")
                    CUST_CODE = rowCUSTPO.Item("CUST_CODE") & ""
                    ORDR_CUST_PO = rowCUSTPO.Item("ORDR_CUST_PO") & ""
                    WAVE_NO = rowCUSTPO.Item("WAVE_NO")
                    LOCATION_CODE = rowCUSTPO.Item("LOCATION_CODE")

                    If Not ASCMAIN1.Logical_Open("WHTWAVE1", WAVE_NO) Then
                        CreateResponse("", "R", "Could Not Access Wave " & WAVE_NO)
                        Exit Select
                    End If
                    'We need to lock individual pages? - not entire Wave
                    If Not ASCMAIN1.Logical_Lock("WHTINST1", WAVE_INST_NO) Then
                        If WAVE_INST_NO = SCANNED_INSTRUCTION Then
                            CreateResponse("", "R", "Could Not Access Wave Instruction " & WAVE_INST_NO)
                            Exit Select
                        Else
                            If Val(OPEN_PICKS) = 0 Then
                                CreateResponse("", "R", "Done with Case Instructions,  Scan New Instruction " & WAVE_INST_NO)
                                LAST_SEQ_NO = ""
                                Exit Select
                            Else
                                CreateResponse("", "R", "Done with Instructions, Hit Enter to check Wave " & WAVE_INST_NO)
                                LAST_SEQ_NO = ""
                                Exit Select
                            End If
                        End If
                    End If

                    rowWHTINST1 = LookUp("WHTINST1", WAVE_INST_NO)
                    If rowWHTINST1.Item("WAVE_INST_STATUS") & "" <> "0" Then
                        CreateResponse("", "R", "Wave Instruction " & WAVE_INST_NO & " no longer Open")
                        Exit Select
                    End If

                    Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {G.WHSE_CODE, LOCATION_CODE})
                    LAST_SEQ_NO = rowWHTLOCM1("LOCATION_ROUTE_SEQ") & ""




                    tbl.Rows.Clear()

                    For Each TABLE_NAME As String In New String() {"WHTINST2", "WHTLOCB1", "WHTBARC0"}
                        dst.Tables(TABLE_NAME).Rows.Clear()
                    Next

                    Fill_Records("WHTINST2", WAVE_INST_NO)

                    Dim SCQ As String = ""
                    Dim STYLE_COLORs As Integer = 0
                    BAR_CODE_waved = ""
                    UNITS = 0
                    For Each ROW As DataRow In dst.Tables("WHTINST2").Select("")
                        Dim LOCATION_QTY_WAVE As Integer = Val(ROW.Item("LOCATION_QTY_WAVE") & "")
                        SCQ &= vbCrLf & ROW.Item("STYLE_CODE") & "-" & ROW.Item("COLOR_CODE") & ":" & CStr(LOCATION_QTY_WAVE)
                        UNITS += LOCATION_QTY_WAVE
                        STYLE_COLORs += 1
                        BAR_CODE_waved = ROW.Item("BAR_CODE")
                    Next

                    WAVE_INST_TEXT = "Instruction " & WAVE_INST_NO _
                        & vbCrLf & IIf(WAVE_TYPE = "W", "Work Order", CUST_CODE & ", PO " & ORDR_CUST_PO) _
                        & vbCrLf & "Location: " & LOCATION_CODE _
                        & vbCrLf & "Case ID: " & BAR_CODE_waved _
                        & vbCrLf & CStr(STYLE_COLORs) & " " & IIf(STYLE_COLORs = 1, "Item", "Items") & ", " & CStr(UNITS) & " Units" _
                        & IIf(STYLE_COLORs = 1, Split(SCQ, ":")(0), SCQ)

                    If BAR_CODE_pick_into <> "" Then
                        Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", BAR_CODE_pick_into)
                        If rowWHTBARC1 IsNot Nothing Then
                            CreateResponse("SCAN_LPN", "G", "Case ID: " & BAR_CODE_pick_into & vbCrLf & WAVE_INST_TEXT, True)
                            Exit Select
                        Else
                            BAR_CODE_pick_into = ""
                        End If
                    End If

                    CreateResponse("SCAN_NEW", "B", WAVE_INST_TEXT, True)

                Case "SCAN_NEW"
                    If SCANTEXT = "CANCEL" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "B", "")
                    ElseIf SCANTEXT = "ZERO" Then
                        CreateResponse("VERIFY_ZERO", "B", "")
                    Else
                        Check_BAR_CODE(SCANTEXT)
                        If SCANTEXT.Length <> 8 Then
                            CreateResponse("", "R", "Invalid Valid Value for a Case ID (" & SCANTEXT & ")")
                            Exit Select
                        End If
                        Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", SCANTEXT)
                        If rowWHTBARC1 IsNot Nothing AndAlso rowWHTBARC1.Item("TRAN_TYPE") <> "V" Then
                            CreateResponse("", "R", "case id " & SCANTEXT & " exists, Scan New Case ID")
                            Exit Select
                        Else
                            If rowWHTBARC1 IsNot Nothing AndAlso Not rowWHTBARC1.Item("BC_COMMENT").ToString.Contains(WAVE_NO) Then
                                CreateResponse("", "R", "case id " & SCANTEXT & " exists for another Wave, Scan New Case ID")
                                Exit Select
                            End If
                        End If

                        BAR_CODE_pick_into = SCANTEXT
                        If Not ASCMAIN1.Logical_Lock("WHTBARC1", BAR_CODE_pick_into) Then
                            CreateResponse("", "R", "Could Not Lock Access to Case ID " & BAR_CODE_pick_into)
                            Exit Select
                        End If

                        BAR_CODE_pick_from = ""
                        '& vbCrLf & "Case ID Waved: " & BAR_CODE_waved
                        CreateResponse("SCAN_LPN", "G", "New Case ID: " & BAR_CODE_pick_into & vbCrLf & WAVE_INST_TEXT, True)
                    End If

                Case "SCAN_LPN"
                    If SCANTEXT = "CANCEL" Then
                        ASCMAIN1.MultiTask_Release()

                        CreateResponse("NEXT_INST", "B", "")
                        'If SCANTEXT = "DONE" Then
                        '    CreateResponse("VERIFY", "B", "")
                    Else
                        If SCANTEXT = "SHOW" Then
                            Dim RESPONSE As String = "NEXT "
                            Dim rows As DataRow() = tbl.Select("")
                            If rows.Length = 0 Then
                                RESPONSE = "FINISHED"
                            Else
                                For Each row As DataRow In rows
                                    RESPONSE = RESPONSE & row.Item("BAR_CODE") & ","
                                Next
                                RESPONSE = Left(RESPONSE, RESPONSE.Length - 1)
                            End If
                            CreateResponse("", "B", RESPONSE)
                        Else
                            Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", SCANTEXT)
                            If rowWHTBARC1 Is Nothing Then
                                CreateResponse("", "R", "Case ID " & SCANTEXT & " does not Exist")
                                Exit Select
                            End If

                            If Not ASCMAIN1.Logical_Lock("WHTBARC1", SCANTEXT) Then
                                CreateResponse("", "R", "Could Not Lock Access to Case ID " & SCANTEXT)
                                Exit Select
                            End If

                            ASCMAIN1.sql = "Select Sum (L.LOCATION_QTY) LOCATION_QTY from WHTLOCB1 L, WHTINST2 I" & vbCrLf _
                                & " where L.BAR_CODE = '" & SCANTEXT & "' " & vbCrLf _
                                & " and I.BAR_CODE = '" & BAR_CODE_waved & "' " & vbCrLf _
                                & " and I.STYLE_CODE = L.STYLE_CODE " & vbCrLf _
                                & " and I.COLOR_CODE = L.COLOR_CODE " & vbCrLf _
                                & " and L.LOCATION_CODE = '" & LOCATION_CODE & "' " & vbCrLf _
                                & " and L.WHSE_CODE = '" & G.WHSE_CODE & "' " & vbCrLf _
                                & " and L.LOCATION_QTY > 0"
                            LOCATION_QTY_on_hand = Val(ASCDATA1.GetDataValue & "")

                            If LOCATION_QTY_on_hand = 0 Then
                                CreateResponse("", "R", "Wrong Case " & SCANTEXT & ", Case cannot be used")
                                Exit Select
                            End If

                            BAR_CODE_pick_from = SCANTEXT
                            CreateResponse("QTY_PICKED", "G", "Pick from Case ID " & BAR_CODE_pick_from & " Scanned", True)

                        End If
                    End If

                Case "QTY_PICKED"
                    If SCANTEXT = "CANCEL" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "B", "")
                        Exit Select
                    End If

                    If Len(SCANTEXT) > 6 Or Val(SCANTEXT) < 0 Or (Val(SCANTEXT) = 0 And Trim(SCANTEXT) <> "0") Then
                        CreateResponse("", "R", "Invalid Qty " & SCANTEXT)
                        Exit Select
                    ElseIf Val(SCANTEXT) > LOCATION_QTY_on_hand Then
                        CreateResponse("", "R", "Invalid Qty " & SCANTEXT & " - there are only " & CStr(LOCATION_QTY_on_hand) & " Units in Carton")
                        Exit Select
                    End If

                    Dim rowWHTINST2 As DataRow = dst.Tables("WHTINST2").Select("")(0)
                    Dim LOCATION_QTY_WAVE As Integer = rowWHTINST2.Item("LOCATION_QTY_WAVE")

                    If Val(SCANTEXT) > LOCATION_QTY_WAVE Then
                        CreateResponse("", "R", "Invalid Qty " & SCANTEXT, ", Required: " & ToString(LOCATION_QTY_WAVE))
                        Exit Select
                    End If

                    PICKED = Val(SCANTEXT)
                    CreateResponse("VERIFY", "G", "Pick from Case ID " & BAR_CODE_pick_from & " Scanned", True)

                Case "VERIFY"
                    If SCANTEXT = "Y" Then

                        Update_Record()

                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "B", "Work on Inst " & WAVE_INST_NO & " Complete")
                        If Not ASCMAIN1.Logical_Lock("WHTINST1", SCANNED_INSTRUCTION) Then
                            CreateResponse("", "R", "Could Not Access Wave Instruction " & SCANNED_INSTRUCTION)
                            Exit Select
                        End If


                    ElseIf SCANTEXT = "N" Then
                        BAR_CODE_pick_from = ""
                        CreateResponse("SCAN_LPN", "B", "Re-Select Case ID to Pick From" & vbCrLf & WAVE_INST_TEXT)

                    ElseIf SCANTEXT = "CANCEL" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "R", "Work on Inst " & WAVE_INST_NO & " Cancelled")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If


                Case "VERIFY_ZERO"
                    If SCANTEXT = "Y" Then

                        Update_Record_Zero()

                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "B", "Inst " & WAVE_INST_NO & " Zero-Picked")

                    ElseIf SCANTEXT = "N" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "R", "Work on Inst " & WAVE_INST_NO & " Cancelled")

                    ElseIf SCANTEXT = "CANCEL" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "R", "Work on Inst " & WAVE_INST_NO & " Cancelled")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If

            End Select
        End If
    End Sub

    Overrides Function GetStatus() As String
        Dim STATUS As String = "" ' "Scanned " & tbl.Select("ISNULL(BAR_CODE_SCANNED,'') <> ''").Length & " out of " & CStr(UNITS) & " Units"
        Return STATUS
    End Function

    Sub Update_Record()

        BeginTrans()

        ' Record New Load for New Case in Gun

        Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", BAR_CODE_pick_into)
        If rowWHTBARC1 Is Nothing Then

            LOAD_NO = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
            Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
            With rowWHTBARC0
                .Item("LOAD_NO") = LOAD_NO
                .Item("WHSE_CODE") = G.WHSE_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = G.USER_ID
                .Item("LOAD_STATUS") = "A"
                .Item("LOCATION_CODE") = G.GUN_LOC
                .Item("TRAN_TYPE") = "V"
                .Item("TRAN_NO") = WAVE_INST_NO
                .Item("LOAD_DATE") = DATETIME_STAMP.Date
            End With
            dst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
            Update_Record_TDA("WHTBARC0")

            ' Record New Case ID used for Pick

            rowWHTBARC1 = dst.Tables("WHTBARC1").NewRow
            With rowWHTBARC1
                .Item("BAR_CODE") = BAR_CODE_pick_into
                .Item("TRAN_TYPE") = "V"
                .Item("TRAN_NO") = WAVE_INST_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = G.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = G.USER_ID
                .Item("STATUS_CODE") = "A"
                .Item("BC_COMMENT") = WAVE_NO & " PIECE PICK"
                .Item("LOAD_NO") = LOAD_NO
            End With
            dst.Tables("WHTBARC1").Rows.Add(rowWHTBARC1)
            Update_Record_TDA("WHTBARC1")
        Else
            LOAD_NO = rowWHTBARC1.Item("LOAD_NO")
        End If

        ' De-Commit Wave from Locator Inventory - Original Case

        For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("")
            Dim BAR_CODE As String = rowWHTINST2.Item("BAR_CODE")
            Dim STYLE_CODE As String = rowWHTINST2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTINST2.Item("COLOR_CODE")
            Dim LOCATION_QTY_WAVE As Int64 = Val(rowWHTINST2.Item("LOCATION_QTY_WAVE") & "")
            ASCMAIN1.sql = "Update WHTLOCB1" & vbCrLf _
                & " Set LOCATION_QTY_WAVE = NVL(LOCATION_QTY_WAVE,0) - " & CStr(LOCATION_QTY_WAVE) & vbCrLf _
                & " where WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                & "   and BAR_CODE = '" & BAR_CODE & "'" & vbCrLf _
                & "   and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and COLOR_CODE = '" & COLOR_CODE & "'"
            ASCDATA1.ExecuteSQL()
            rowWHTINST2.Item("LOCATION_QTY_PICK") = PICKED
            rowWHTINST2.Item("BAR_CODE") = BAR_CODE_pick_from
        Next
        Update_Record_TDA("WHTINST2", "WAVE_INST_NO = '" & WAVE_INST_NO & "'") ' delete & replace since we are changing the key

        ' Mark Instruction as Picked

        ASCMAIN1.sql = "Update WHTINST1 Set" & vbCrLf _
            & "  WAVE_INST_STATUS = '1'" & vbCrLf _
            & ", LOCATION_CODE_OTHER = '" & G.GUN_LOC & "'" & vbCrLf _
            & ", LOAD_NO_OTHER = '" & LOAD_NO & "'" & vbCrLf _
            & ", BAR_CODE_OTHER = '" & BAR_CODE_pick_into & "'" & vbCrLf _
            & ", INIT_DATE = SYSDATE, INIT_OPER = '" & G.USER_ID & "'" & vbCrLf _
            & " where WAVE_INST_NO = '" & WAVE_INST_NO & "' and WAVE_INST_STATUS = '0'"
        ASCDATA1.ExecuteSQL()

        ' Locator System Audit Record & Location Update

        ASCDATA1.ExecuteSP("WHPLOCB2", "VVV", _
                           New Object() {"V", WAVE_INST_NO, ASCMAIN1.SESSION_NO}, _
                           New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

        ' Update Wave Summary by Style/Color

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select WHTINST1.WAVE_NO, WHTINST2.WAVE_LNO, WHTINST1.WAVE_SUB, WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE" & vbCrLf _
            & "   , Sum (WHTINST2.LOCATION_QTY_PICK) WAVE_QTY_PICK" & vbCrLf _
            & "   from WHTINST2,WHTINST1" & vbCrLf _
            & "   where WHTINST1.WAVE_INST_NO = WHTINST2.WAVE_INST_NO" & vbCrLf _
            & "     and WHTINST2.WAVE_INST_NO = '" & WAVE_INST_NO & "'" & vbCrLf _
            & "   group by WHTINST1.WAVE_NO, WHTINST2.WAVE_LNO, WHTINST1.WAVE_SUB, WHTINST2.STYLE_CODE, WHTINST2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update WHTWAVE2 Set WAVE_QTY_PICK = NVL(WAVE_QTY_PICK,0) + NVL(R1.WAVE_QTY_PICK,0)" & vbCrLf _
            & "    where WAVE_NO = R1.WAVE_NO" & vbCrLf _
            & "      and WAVE_LNO = R1.WAVE_LNO" & vbCrLf _
            & "      and ((R1.WAVE_SUB = '0' and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE)" & vbCrLf _
            & "       or  (R1.WAVE_SUB = '1' and STYLE_CODE_SUB = R1.STYLE_CODE and COLOR_CODE_SUB = R1.COLOR_CODE));" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into WHTWAVE2 (WAVE_NO, STYLE_CODE, COLOR_CODE, WAVE_QTY_PICK) " & vbCrLf _
            & "     values (R1.WAVE_NO, R1.STYLE_CODE, R1.COLOR_CODE, R1.WAVE_QTY_PICK);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        CommitTrans()
    End Sub

    Sub Update_Record_Zero()

        BeginTrans()

        ' De-Commit Wave from Locator Inventory - Original Case

        For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("")
            Dim BAR_CODE As String = rowWHTINST2.Item("BAR_CODE")
            Dim STYLE_CODE As String = rowWHTINST2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTINST2.Item("COLOR_CODE")
            Dim LOCATION_QTY_WAVE As Int64 = Val(rowWHTINST2.Item("LOCATION_QTY_WAVE") & "")
            ASCMAIN1.sql = "Update WHTLOCB1" & vbCrLf _
                & " Set LOCATION_QTY_WAVE = NVL(LOCATION_QTY_WAVE,0) - " & CStr(LOCATION_QTY_WAVE) & vbCrLf _
                & " where WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                & "   and BAR_CODE = '" & BAR_CODE_waved & "'" & vbCrLf _
                & "   and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and COLOR_CODE = '" & COLOR_CODE & "'"
            ASCDATA1.ExecuteSQL()
            rowWHTINST2.Item("LOCATION_QTY_PICK") = 0
        Next
        Update_Record_TDA("WHTINST2", "WAVE_INST_NO = '" & WAVE_INST_NO & "'") ' delete & replace since we are changing the key

        ' Mark Instruction as Picked

        ASCMAIN1.sql = "Update WHTINST1 Set" & vbCrLf _
            & "  WAVE_INST_STATUS = '1'" & vbCrLf _
            & ", LOCATION_CODE_OTHER = '" & G.GUN_LOC & "'" & vbCrLf _
            & ", INIT_DATE = SYSDATE, INIT_OPER = '" & G.USER_ID & "'" & vbCrLf _
            & " where WAVE_INST_NO = '" & WAVE_INST_NO & "' and WAVE_INST_STATUS = '0'"
        ASCDATA1.ExecuteSQL()

        Lock_Location()

        CommitTrans()
    End Sub

    Function Check_BAR_CODE(BAR_CODE As String) As String
        Dim prefix As String = ""
        If BAR_CODE = "" Then Return BAR_CODE

        If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
            prefix = BAR_CODE.ToUpper.Substring(0, 1)
            BAR_CODE = BAR_CODE.Substring(1)
        End If

        If BAR_CODE.PadLeft(8, "0") <> Format(Val(BAR_CODE), "".PadLeft(8, "0")) Then
            BAR_CODE = ""
        Else
            If prefix = "" Then
                BAR_CODE = BAR_CODE.PadLeft(8, "0")
            Else
                BAR_CODE = prefix & BAR_CODE.PadLeft(7, "0")
            End If
        End If
        Return BAR_CODE
    End Function

    Sub Lock_Location()
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", G.WHSE_CODE)
        If LOCATION_CODE = rowICTWHSE1.Item("WHSE_LOC_REC") & "" Then
            ' NEVER LOCK RECEIVING
        Else
            ASCMAIN1.sql = "Update WHTLOCM1 Set LOCATION_LOCKED = '1'" _
                & " where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2"
            '& "   and NVL(LOCATION_USE,'?') = '?'"
            Dim R As Integer = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {G.WHSE_CODE, LOCATION_CODE})

            If R = 1 Then
                ASCMAIN1.Record_Event("WHTLOCM1", G.WHSE_CODE, LOCATION_CODE, Now, G.USER_ID, "LOCLCK", "CNL WVPICK", WAVE_INST_NO)
            End If
        End If
    End Sub

    Overrides Function Get_Anticipated_Next_Response() As String
        Select Case AppState
            Case "NEXT_INST"
                Return ""
            Case "SCAN_NEW"
                'Range given by Doug - when we return empty, time for another range and reset counter.
                Dim NEW_LPN As String = ASCMAIN1.Next_Control_No("WHTBARC1.AUTO_LPN")
                Dim START_LPN As Int32 = 713545
                Dim END_LPN As Int32 = 4399479
                If (START_LPN + Val(NEW_LPN)) < END_LPN Then
                    NEW_LPN = Format((START_LPN + Val(NEW_LPN)), "00000000")
                    Return NEW_LPN
                Else
                    Return ""
                End If
            Case "SCAN_LPN"
                Dim rows() As DataRow = tbl.Select("")
                If rows.Length = 0 Then
                    Return "CANCEL"
                Else
                    Return rows(0).Item("BAR_CODE")
                End If

            Case "QTY_PICKED"
                ASCMAIN1.sql = "SELECT SUM (LOCATION_QTY_WAVE) FROM WHTINST2 WHERE WAVE_INST_NO = '" & WAVE_INST_NO & "'"
                Return CStr(Val(ASCDATA1.GetDataValue))

            Case "REMOVE_LPN"
                Return ""
            Case "VERIFY"
                Return "Y"
            Case Else
                Return ""
        End Select
    End Function
End Class