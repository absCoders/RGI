Public Class WHCRF004
    ' Process Wave Instruction (Pallet Pick)

    Inherits WHCRF000

    Dim CUST_CODE As String
    Dim ORDR_CUST_PO As String
    Dim BAR_CODE As String
    Dim LOAD_NO As String
    Dim WAVE_INST_NO As String
    Dim WAVE_NO As String
    Dim LOCATION_CODE As String
    Dim CASES As Integer
    Dim WAVE_INST_TEXT As String
    Dim WAVE_INST_STATUS As String
    Dim SCANNED_INSTRUCTION As String
    Dim LAST_SEQ_NO As String = ""
    Dim OPEN_PICKS As String

    Dim rowICTWHSE1 As DataRow
    Dim rowWHTINST1 As DataRow

    Sub New(g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF004"

        AppStates.Add("NEXT_INST", "Enter Wave No or Press Enter for Next|EXIT|")
        AppStates.Add("SCAN_LPN", "Scan Case ID from Pallet to Pick|DONE|EXIT|")
        AppStates.Add("SCAN_COUNT", "How Many Cartons in Pallet|EXIT|")
        AppStates.Add("SCAN_LPN2", "Scan Next Case ID from Pallet|DONE|EXIT|")
        AppStates.Add("VERIFY", "OK to Update (Y/N)|Y|N|CANCEL|")

        AppState = "NEXT_INST"

        With dst
            With .Tables.Add("WHTSCANS").Columns
                .Add("BAR_CODE")
                .Add("LOAD_NO")
                .Add("LOCATION_CODE")
                .Add("SCANNED")
                .Add("ERROR")
            End With
            .Tables("WHTSCANS").PrimaryKey = New DataColumn() { .Tables("WHTSCANS").Columns("BAR_CODE")}

            Create_TDA(.Tables.Add, "WHTINST2", "*", 1)
            'Create_TDA(.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(.Tables.Add, "WHTBARC0", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
        End With

        tbl = dst.Tables("WHTSCANS") ' New DataTable
        rowICTWHSE1 = ASCDATA1.GetDataRow("Select * from ICTWHSE1 where WHSE_CODE = '" & g.WHSE_CODE & "'")

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

                    tbl.Rows.Clear()
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
                        CreateResponse("", "R", "No Picks Available for Wave (" & SCANTEXT & ")")
                        Exit Select
                    End If


                    ASCMAIN1.sql = "Select case when WHTWAVE1.WAVE_TYPE = 'W' then 'WorkOrdr' else  SOTORDR0.CUST_CODE END CUST_CODE, " _
                            & "case when WHTWAVE1.WAVE_TYPE = 'W' then 'WorkOrdr' else SOTORDR0.ORDR_CUST_PO end ORDR_CUST_PO, " _
                            & " WHTINST1.WAVE_NO, WHTINST1.LOCATION_CODE, WHTINST1.LOAD_NO " _
                            & "from WHTINST1 " _
                            & "join WHTWAVE1 on WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO " _
                            & "left join whtwave3 on WHTWAVE3.WAVE_NO = WHTINST1.WAVE_NO " _
                            & "left join SOTSHIP1 on SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO " _
                            & "left join SOTORDR0 on SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                            & " where WHTINST1.WAVE_INST_NO = '" & WAVE_INST_NO & "'"
                    Dim rowCUSTPO As DataRow = ASCDATA1.GetDataRow
                    CUST_CODE = rowCUSTPO.Item("CUST_CODE")
                    ORDR_CUST_PO = rowCUSTPO.Item("ORDR_CUST_PO") & ""
                    WAVE_NO = rowCUSTPO.Item("WAVE_NO")
                    LOCATION_CODE = rowCUSTPO.Item("LOCATION_CODE")
                    LOAD_NO = rowCUSTPO.Item("LOAD_NO")

                    If Not ASCMAIN1.Logical_Open("WHTWAVE1", WAVE_NO) Then
                        CreateResponse("", "R", "Could Not Access Wave " & WAVE_NO)
                        Exit Select
                    End If

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
                    ASCMAIN1.sql = "Select Distinct WHTINST2.BAR_CODE, WHTBARC1.LOAD_NO" & vbCrLf _
                        & " from WHTINST2,WHTBARC1" & vbCrLf _
                        & " where WHTINST2.WAVE_INST_NO = '" & WAVE_INST_NO & "'" & vbCrLf _
                        & "   and WHTBARC1.BAR_CODE = WHTINST2.BAR_CODE"
                    For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("", "BAR_CODE")
                        BAR_CODE = ROW.Item("BAR_CODE")
                        Dim row2 As DataRow = tbl.NewRow
                        row2.Item("BAR_CODE") = BAR_CODE
                        row2.Item("LOAD_NO") = ROW.Item("LOAD_NO")
                        ASCMAIN1.sql = "Select Distinct LOCATION_CODE from WHTLOCB1" & vbCrLf _
                            & " where WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                            & "   and BAR_CODE = '" & BAR_CODE & "'" & vbCrLf _
                            & "   and LOCATION_QTY > 0 and location_code not like '99%'"
                        Dim rows3() As DataRow = ASCDATA1.GetDataTable.Select("")
                        If rows3.Length = 0 Then
                            row2.Item("ERROR") = "Case ID " & BAR_CODE & " Not Available for Picking"
                            CreateResponse("", "R", row2.Item("ERROR"))
                            Exit Select
                        ElseIf rows3.Length > 1 Then
                            row2.Item("ERROR") = "Case ID " & BAR_CODE & " in Multiple Locations"
                            CreateResponse("", "R", row2.Item("ERROR"))
                            Exit Select
                        ElseIf rows3(0).Item("LOCATION_CODE") <> LOCATION_CODE Then
                            row2.Item("LOCATION_CODE") = rows3(0).Item("LOCATION_CODE")
                            row2.Item("ERROR") = "Case ID " & BAR_CODE & " found in Incorrect Locations"
                            CreateResponse("", "R", row2.Item("ERROR"))
                            Exit Select
                        Else
                            row2.Item("LOCATION_CODE") = rows3(0).Item("LOCATION_CODE")
                        End If

                        tbl.Rows.Add(row2)
                    Next

                    For Each TABLE_NAME As String In New String() {"WHTINST2", "WHTBARC0"}
                        dst.Tables(TABLE_NAME).Rows.Clear()
                    Next

                    CASES = tbl.Rows.Count

                    WAVE_INST_TEXT = "Instruction " & WAVE_INST_NO _
                        & vbCrLf & CUST_CODE & ", PO " & ORDR_CUST_PO _
                        & vbCrLf & "Cases on Pallet: " & CStr(CASES) _
                        & vbCrLf & "Location: " & LOCATION_CODE

                    WAVE_INST_STATUS = "C"
                    CreateResponse("SCAN_LPN", "B", WAVE_INST_TEXT, True)

                Case "SCAN_LPN"

                    If SCANTEXT = "DONE" Then
                        If WAVE_INST_STATUS = "C" Then
                            CreateResponse("VERIFY", "R", "No Case ID selected, exit with no pick")
                            Exit Select
                        End If
                        CreateResponse("VERIFY", "B", "")
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
                                CreateResponse("", "R", "Invalid Case ID " & SCANTEXT)
                                Exit Select
                            End If
                            BAR_CODE = SCANTEXT
                            If Not ASCMAIN1.Logical_Lock("WHTBARC1", BAR_CODE) Then
                                CreateResponse("", "R", "Could Not Lock Access to Case ID " & BAR_CODE)
                                Exit Select
                            End If

                            Dim row As DataRow = tbl.Rows.Find(BAR_CODE)
                            If row IsNot Nothing Then
                                ' THIS BAR_CODE IS ONE OF THE CASES THAT WAS ORIGINALLY WAVED
                                'CreateResponse("VERIFY", "G", "Case ID " & BAR_CODE & " Scanned" & vbCrLf & CStr(CASES) & " Cases on Pallet")
                                CreateResponse("SCAN_COUNT", "G", "Case ID " & BAR_CODE & " Scanned" & vbCrLf & "Verify Cases on Pallet")
                                WAVE_INST_STATUS = "1"
                                row.Item("SCANNED") = "Y"
                                Exit Select
                            Else
                                ' THIS BAR_CODE IS NOT ONE OF THE CASES THAT WAS ORIGINALLY WAVED
                                CreateResponse("", "R", "Case ID " & BAR_CODE & " is NOT part of Pallet to Pick")
                                Exit Select
                            End If
                        End If
                    End If

                Case "SCAN_COUNT"
                    If Val(SCANTEXT) = 0 Then
                        CreateResponse("", "R", "Not a valid Count, please re-enter Carton count." & vbCrLf & WAVE_INST_TEXT)
                        Exit Select
                    End If
                    If CASES = Val(SCANTEXT) Then
                        Dim rows As DataRow() = tbl.Select("")
                        For Each row As DataRow In rows
                            row.Item("SCANNED") = "Y"
                        Next
                        CreateResponse("VERIFY", "B", "")
                    Else
                        'Count doesn't match Verify Pallet - Scan each Carton
                        'SCAN_LPN2
                        CreateResponse("SCAN_LPN2", "G", "Case ID " & BAR_CODE & " Scanned" & vbCrLf & "Verify Cases on Pallet")
                        WAVE_INST_STATUS = "1"
                        Exit Select
                    End If

                Case "SCAN_LPN2"
                    If SCANTEXT = "DONE" Then
                        Dim Scanned As Integer = tbl.Compute("COUNT(BAR_CODE)", "SCANNED='Y'")
                        If Scanned = CASES Then
                            CreateResponse("VERIFY", "B", "")
                        Else
                            CreateResponse("VERIFY", "R", Scanned & " Scanned out of " & CASES)
                        End If
                        Exit Select
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
                                CreateResponse("", "R", "Invalid Case ID " & SCANTEXT)
                                Exit Select
                            End If
                            BAR_CODE = SCANTEXT
                            If Not ASCMAIN1.Logical_Lock("WHTBARC1", BAR_CODE) Then
                                CreateResponse("", "R", "Could Not Lock Access to Case ID " & BAR_CODE)
                                Exit Select
                            End If

                            Dim row As DataRow = tbl.Rows.Find(BAR_CODE)
                            If row IsNot Nothing Then
                                ' THIS BAR_CODE IS ONE OF THE CASES THAT WAS ORIGINALLY WAVED
                                CreateResponse("", "G", "Case ID " & BAR_CODE & " Scanned" & vbCrLf & "Scan all Cases on Pallet")
                                WAVE_INST_STATUS = "1"
                                row.Item("SCANNED") = "Y"
                                Exit Select
                            Else
                                ' THIS BAR_CODE IS NOT ONE OF THE CASES THAT WAS ORIGINALLY WAVED - move to LNF
                                Move_to_LNF(BAR_CODE)
                                CreateResponse("", "R", "Remove Box" & vbCrLf & "Case ID " & BAR_CODE & " is NOT part of Pallet to Pick")
                                Exit Select
                            End If
                        End If
                    End If

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
                        CreateResponse("SCAN_LPN", "B", "Resume Case Scans" & vbCrLf & WAVE_INST_TEXT, True)

                    ElseIf SCANTEXT = "CANCEL" Then
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("NEXT_INST", "B", "Work on Inst " & WAVE_INST_NO & " Cancelled")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub

    Overrides Function GetStatus() As String
        Dim STATUS As String = ""
        Return STATUS
    End Function

    Sub Move_to_LNF(BAR_CODE)
        Dim WHSE_TRAN_NO As String
        Dim WHSE_TRAN_LNO As Integer
        'Dim BAR_CODE As String
        Dim LOCATION_CODE_ORIG As String

        Dim WHSE_CODE As String = G.WHSE_CODE
        Dim LOCATION_CODE As String

        dst.Tables("WHTMOVE1").Rows.Clear()
        dst.Tables("WHTMOVE2").Rows.Clear()
        WHSE_TRAN_NO = ""
        WHSE_TRAN_LNO = 0

        BeginTrans()

        ASCMAIN1.sql = "Select Distinct LOCATION_CODE from WHTLOCB1" & vbCrLf _
                            & " where WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                            & "   and BAR_CODE = '" & BAR_CODE & "'" & vbCrLf _
                            & "   and LOCATION_QTY > 0 and location_code not like '99%'"
        Dim rows1() As DataRow = ASCDATA1.GetDataTable.Select("")
        If rows1.Length = 0 Then
            'row2.Item("ERROR") = "Case ID " & BAR_CODE & " Not Available for Move"
            Exit Sub
        ElseIf rows1.Length > 1 Then
            'row2.Item("ERROR") = "Case ID " & BAR_CODE & " in Multiple Locations"
            Exit Sub
        End If
        LOCATION_CODE_ORIG = rows1(0).Item("LOCATION_CODE") & ""
        LOCATION_CODE = rowICTWHSE1.Item("WHSE_LOC_LNF") & ""

        Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", BAR_CODE)

        ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1 " _
                                  & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                                  & " and  WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE_ORIG & "'" _
                                  & " and  WHTLOCB1.BAR_CODE = '" & BAR_CODE & "'" _
                                  & " and  WHTLOCB1.LOCATION_QTY > 0 "
        For Each rowWHTLOCB1 As DataRow In ASCDATA1.GetDataTable.Select("")

            If WHSE_TRAN_NO = "" Then
                WHSE_TRAN_NO = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
            End If

            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
            With rowWHTMOVE2

                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                WHSE_TRAN_LNO += 1
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO

                .Item("LOCATION_CODE_FROM") = LOCATION_CODE_ORIG
                .Item("LOCATION_CODE_TO") = LOCATION_CODE ' LNF
                .Item("LOAD_NO_FROM") = rowWHTBARC1.Item("LOAD_NO")
                .Item("LOAD_NO_TO") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                .Item("BAR_CODE") = BAR_CODE

                .Item("WHSE_TRAN_QTY") = rowWHTLOCB1.Item("LOCATION_QTY")
                .Item("STYLE_CODE") = rowWHTLOCB1.Item("STYLE_CODE")
                .Item("COLOR_CODE") = rowWHTLOCB1.Item("COLOR_CODE")
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"

            End With
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

        Next

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        With rowWHTMOVE1
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_TYPE") = "M"
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("STATUS") = "U"
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
        End With
        dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)

        Update_Record_TDA("WHTMOVE1")
        Update_Record_TDA("WHTMOVE2")
        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {WHSE_TRAN_NO, 0, 1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})
        CommitTrans()

    End Sub


    Sub Update_Record()

        BeginTrans()

        Dim LOAD_NO_OTHER As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
        Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
        With rowWHTBARC0
            .Item("LOAD_NO") = LOAD_NO_OTHER
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

        Fill_Records("WHTINST2", WAVE_INST_NO)
        For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("")
            Dim BAR_CODE As String = rowWHTINST2.Item("BAR_CODE")
            Dim STYLE_CODE As String = rowWHTINST2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTINST2.Item("COLOR_CODE")
            Dim LOCATION_QTY_WAVE As String = Val(rowWHTINST2.Item("LOCATION_QTY_WAVE") & "")
            Dim row As DataRow = tbl.Rows.Find(BAR_CODE)
            Dim SCANNED As String = row.Item("SCANNED") & ""

            If WAVE_INST_STATUS = "C" Or SCANNED <> "Y" Then
                rowWHTINST2.Item("LOCATION_QTY_PICK") = 0
            Else
                rowWHTINST2.Item("LOCATION_QTY_PICK") = rowWHTINST2.Item("LOCATION_QTY_WAVE")
            End If
            ASCMAIN1.sql = "Update WHTLOCB1" & vbCrLf _
                & " Set LOCATION_QTY_WAVE = NVL(LOCATION_QTY_WAVE,0) - " & CStr(LOCATION_QTY_WAVE) & vbCrLf _
                & " where WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                & "   and BAR_CODE = '" & BAR_CODE & "'" & vbCrLf _
                & "   and STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & "   and COLOR_CODE = '" & COLOR_CODE & "'"
            ASCDATA1.ExecuteSQL()
        Next
        Update_Record_TDA("WHTINST2")

        ASCMAIN1.sql = "Update WHTINST1 Set" & vbCrLf _
            & "  WAVE_INST_STATUS = '" & WAVE_INST_STATUS & "'" & vbCrLf _
            & ", LOCATION_CODE_OTHER = '" & G.GUN_LOC & "'" & vbCrLf _
            & ", LOAD_NO_OTHER = '" & LOAD_NO_OTHER & "'" & vbCrLf _
            & ", INIT_DATE = SYSDATE, INIT_OPER = '" & G.USER_ID & "'" & vbCrLf _
            & " where WAVE_INST_NO = '" & WAVE_INST_NO & "' and WAVE_INST_STATUS = '0'"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                           New Object() {"V", WAVE_INST_NO, ASCMAIN1.SESSION_NO},
                           New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

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

        For Each row As DataRow In tbl.Select("SCANNED is null")
            Move_to_LNF(row.Item("BAR_CODE") & "")
        Next

    End Sub

    Overrides Function Get_Anticipated_Next_Response() As String
        Select Case AppState
            Case "NEXT_INST"
                Return ""
            Case "SCAN_LPN"
                Dim rows() As DataRow = tbl.Select("")
                If rows.Length = 0 Then
                    Return "DONE"
                Else
                    Return rows(0).Item("BAR_CODE")
                End If
            Case "SCAN_COUNT"
                Return tbl.Select("").Count.ToString()
            Case "VERIFY"
                Return "Y"
            Case Else
                Return ""
        End Select
    End Function
End Class