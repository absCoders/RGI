Public Class WHCRF007
    ' Cycle Application Verify

    Inherits WHCRF000

    Dim WAVE_NO As String
    Dim LOCATION_CODE_DEPOSIT As String
    Dim LOAD_NO_DEPOSIT As String
    Dim BAR_CODE As String
    Dim LOCATION_CODE As String
    Dim Cases_count As Integer
    Dim CYCLE_NO As String
    Dim INVALID_BAR_CODE As String
    Dim CASES_BOOK As Integer
    Dim CASES_PHYS As Integer
    Dim CYCLE_NO1 As String
    Dim BAR_CODE_LOCATION As String



 
    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF007"

        AppStates.Add("SCAN_LOC", "Scan Location for Cycle |EXIT|")
        AppStates.Add("SCAN_LPN", "Scan Case ID to Cycle|DONE|EXIT|")
        AppStates.Add("VERIFY", "Are You Done (Y/N)|Y|N|CANCEL|")

        AppState = "SCAN_LOC"


        With dst
            Create_TDA(.Tables.Add, "WHTCYCL1", "*")
            Create_TDA(.Tables.Add, "WHTCYCL2", "*")
            Create_TDA(.Tables.Add, "WHTBARC0", "*")
            Create_TDA(.Tables.Add, "WHTBARC1", "*")
            Create_TDA(.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")


        End With

        tbl = dst.Tables("WHTCYCL2") ' New DataTable



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
                    Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {G.WHSE_CODE, SCANTEXT})
                    If rowWHTLOCM1 Is Nothing Then
                        CreateResponse("", "R", "Invalid Location " & SCANTEXT)

                        Exit Select
                    End If

                    If rowWHTLOCM1.Item("LOCATION_LOCKED") & "" = "1" Then
                        CreateResponse("", "R", "Location " & SCANTEXT & " is Locked")
                        Exit Select
                    End If

                    'If rowWHTLOCM1.Item("LOCATION_USE") & "" <> "" Then
                    '    CreateResponse("", "R", "Location " & SCANTEXT & " is not Valid for this App")
                    '    Exit Select
                    'End If

                    If Not ASCMAIN1.Logical_Lock("WHTLOCM1", SCANTEXT) Then
                        CreateResponse("", "R", "Could not lock access to Locaton " & SCANTEXT)
                        Exit Select
                    End If

                    Dim rowsWAVED() As DataRow '= ASCDATA1.GetDataTable.Select("")
                    ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                    & " where LOCATION_CODE = '" & SCANTEXT & "'" _
                    & " and WHSE_CODE = '" & G.WHSE_CODE & "'" _
                    & " and LOCATION_QTY_WAVE > 0"
                    rowsWAVED = ASCDATA1.GetDataTable.Select("")
                    If rowsWAVED.Length <> 0 Then
                        CreateResponse("", "R", "Location " & SCANTEXT & " is already Committed to a Wave")
                        Exit Select
                    End If
                
                    LOCATION_CODE = SCANTEXT

                    dst.Tables("WHTCYCL2").Rows.Clear()
                    dst.Tables("WHTCYCL1").Rows.Clear()

                    CYCLE_NO = ASCMAIN1.Next_Control_No("WHTCYCL1.CYCLE_NO")


                    ASCMAIN1.sql = "Select Distinct BAR_CODE from WHTLOCB1" & vbCrLf _
                        & " where WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.LOCATION_QTY > 0"

                    Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rows.Length > 0 Then
                        Cases_count = rows.Length
                        For Each ROW As DataRow In rows
                            Dim row2 As DataRow = tbl.NewRow
                            row2.Item("CYCLE_NO") = CYCLE_NO
                            row2.Item("BAR_CODE") = ROW.Item("BAR_CODE")
                            row2.Item("LOCATION_CODE_ORIG") = LOCATION_CODE
                            tbl.Rows.Add(row2)
                        Next
                    End If
                    CYCLE_NO1 = CYCLE_NO
                    '   LOCATION_CODE = SCANTEXT
                    CreateResponse("SCAN_LPN", "G", "Location " & SCANTEXT & " has been selected")

                Case "SCAN_LPN"
                    INVALID_BAR_CODE = ""
                    If SCANTEXT = "DONE" Then
                        CreateResponse("VERIFY", "B", "Scanned " & tbl.Select("CYCLE_SCAN = '1'").Length & " Cases")
                    Else
                        Dim rowWHTBARC1 As DataRow = LookUp("WHTBARC1", SCANTEXT)
                        If rowWHTBARC1 Is Nothing Then
                            ' CreateResponse("", "R", "Invalid Case ID " & SCANTEXT)
                            INVALID_BAR_CODE = "Y"
                            '     Exit Select
                        End If
                        BAR_CODE = SCANTEXT

                        '  If tbl.Select("SCAN_LPN = '" & BAR_CODE & "'").Length <> 0 Then
                        ' CreateResponse("", "R", "Case " & BAR_CODE & " has already been Scanned.")
                        ' Exit Select
                        'End If

                        ASCMAIN1.sql = "Select Distinct LOCATION_CODE from WHTLOCB1" & vbCrLf _
                            & " where BAR_CODE = '" & BAR_CODE & "'" _
                            & " and WHSE_CODE = '" & G.WHSE_CODE & "'" _
                            & " and LOCATION_QTY > 0"
                        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                        If rows.Length = 0 Then
                            CreateResponse("", "R", "Case ID " & BAR_CODE & " not found in Warehouse with Qty")
                            '   Exit Select
                        ElseIf rows.Length > 1 Then
                            CreateResponse("", "R", "Case ID " & BAR_CODE & " found in Multiple Locations with Qty - Call ABS")
                            Exit Select
                        Else
                            BAR_CODE_LOCATION = rows(0).Item("LOCATION_CODE")
                        End If

                        ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                          & " where BAR_CODE = '" & BAR_CODE & "'" _
                          & " and WHSE_CODE = '" & G.WHSE_CODE & "'" _
                          & " and LOCATION_QTY_WAVE > 0"
                        rows = ASCDATA1.GetDataTable.Select("")
                        If rows.Length <> 0 Then
                            CreateResponse("", "R", "Case ID " & BAR_CODE & " is already Committed to a Wave")
                            Exit Select
                        End If

                        Dim rowWHTCYCL2 As DataRow = dst.Tables("WHTCYCL2").Rows.Find(New Object() {CYCLE_NO1, BAR_CODE})
                        If rowWHTCYCL2 Is Nothing Then
                            Dim row2 As DataRow = tbl.NewRow
                            row2.Item("CYCLE_NO") = CYCLE_NO1
                            row2.Item("BAR_CODE") = BAR_CODE
                            row2.Item("LOCATION_CODE_ORIG") = BAR_CODE_LOCATION
                            row2.Item("CYCLE_SCAN") = "1"
                            row2.Item("CYCLE_NEW") = "1"
                            If INVALID_BAR_CODE = "Y" Then
                                row2.Item("BAR_CODE_INVALID") = "1"
                                row2.Item("CYCLE_NEW") = ""
                            End If
                            tbl.Rows.Add(row2)
                        Else
                            rowWHTCYCL2.Item("CYCLE_SCAN") = "1"
                        End If




                        ' tbl.Rows.Add(New Object() {BAR_CODE})

                        CreateResponse("", "B", "Case ID+ " & BAR_CODE & " Scanned")
                        Exit Select
                    End If


                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_LOC", "B", "Cycle Updated, Scan new location")
                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_LPN", "B", "Scan Ignored, Re-Scan LPN")

                    ElseIf SCANTEXT = "CANCEL" Then

                        CreateResponse("SCAN_LOC", "R", "Cycle Cancelled, Re-scan  location")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub

    Sub Update_Record()
        Dim CYCLE_RESOLUTION As String
        Dim CYCLE_STATUS As String
 
        BeginTrans()

        ' REM SQL STATMENT AGAINT CYCL2 , NEW , OR NOT SCANED, CHANGE CYCLE STATUS AND CYCLE RESOLUTION 
        CYCLE_RESOLUTION = "G"
        CYCLE_STATUS = "G"



        ASCMAIN1.sql = " CYCLE_NEW = '1' OR CYCLE_SCAN IS NULL"
        ' Dim rowWHTCYCL2() As DataRow = dst.Tables("WHTCYCL2").Select(ASCMAIN1.sql)
        If dst.Tables("WHTCYCL2").Select(ASCMAIN1.sql).Length <> 0 Then
            CYCLE_RESOLUTION = "P"
            CYCLE_STATUS = "D"
        End If


        ASCMAIN1.sql = " CYCLE_SCAN = '1'"
        Dim rowWHTCYCL2_COUNT() As DataRow = dst.Tables("WHTCYCL2").Select(ASCMAIN1.sql)
        If rowWHTCYCL2_COUNT.Length <> 0 Then
            CASES_PHYS = rowWHTCYCL2_COUNT.Length
        End If

        If CYCLE_STATUS = "D" Then
            ASCMAIN1.sql = "Update WHTLOCM1 Set LOCATION_LOCKED  = '1'" _
            & " where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2"
            ' & " where LOCATION_CODE =  '" & LOCATION_CODE & "'" _
            '     & " and WHSE_CODE = '" & G.WHSE_CODE & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {G.WHSE_CODE, LOCATION_CODE})
        End If

        Dim rowWHTCYCL1 As DataRow = dst.Tables("WHTCYCL1").NewRow
        With rowWHTCYCL1
            .Item("CYCLE_NO") = CYCLE_NO1
            .Item("CYCLE_STATUS") = CYCLE_STATUS
            .Item("CYCLE_RESOLUTION") = CYCLE_RESOLUTION
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = G.USER_ID
            .Item("LOCATION_CODE") = LOCATION_CODE
            .Item("CASES_BOOK") = Cases_count
            .Item("CASES_PHYS") = CASES_PHYS
            .Item("CYCLE_TYPE") = "V"
            .Item("WHSE_CODE") = G.WHSE_CODE

        End With
        dst.Tables("WHTCYCL1").Rows.Add(rowWHTCYCL1)

        Update_Record_TDA("WHTCYCL1")
        Update_Record_TDA("WHTCYCL2")



        CommitTrans()
    End Sub
End Class
