Public Class WHCRF008
    ' Cycle Application Cycle Count

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
        Me.MENU_ITEM_OBJECT = "WHCRF008"

        AppStates.Add("SCAN_LOC", "Scan Location for Cycle |EXIT|")
        AppStates.Add("SCAN_CASES", "Enter Number of cases in Location |CANCEL|EXIT|")
        AppStates.Add("VERIFY", "Are You Done (Y/N)|Y|N|CANCEL|")

        AppState = "SCAN_LOC"


        With dst
            '  With .Tables.Add("WHTSCANS").Columns
            ' .Add("BAR_CODE")
            ' .Add("NEW")
            ' .Add("SCANNED")
            ' End With
            ' .Tables("WHTSCANS").PrimaryKey = New DataColumn() {.Tables("WHTSCANS").Columns("BAR_CODE")}

            Create_TDA(.Tables.Add, "WHTCYCL1", "*")
            Create_TDA(.Tables.Add, "WHTCYCL2", "*")
            Create_TDA(.Tables.Add, "WHTBARC0", "*")
            Create_TDA(.Tables.Add, "WHTBARC1", "*")
            Create_TDA(.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")


        End With

        tbl = dst.Tables("WHTCYCL2") ' New DataTable


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

                    Dim rowsWAVED() As DataRow = ASCDATA1.GetDataTable.Select("")
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

                    Dim CYCLE_NO As String = ASCMAIN1.Next_Control_No("WHTCYCL1.CYCLE_NO")


                    ASCMAIN1.sql = "Select Distinct BAR_CODE from WHTLOCB1" & vbCrLf _
                        & " where WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.LOCATION_QTY > 0"

                    Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
                    Cases_count = rows.Length
                    If rows.Length > 0 Then
                        For Each ROW As DataRow In rows
                            Dim row2 As DataRow = tbl.NewRow
                            row2.Item("CYCLE_NO") = CYCLE_NO
                            row2.Item("BAR_CODE") = ROW.Item("BAR_CODE")
                            row2.Item("LOCATION_CODE_ORIG") = LOCATION_CODE
                            tbl.Rows.Add(row2)
                        Next
                    End If
                    CYCLE_NO1 = CYCLE_NO
                    '   LOCATION_CODE = SCANTEXT]
                    Dim SSSS As String = ""

                    ASCMAIN1.sql = "Select DISTINCT STYLE_CODE,COUNT(*) DDD from WHTLOCB1" & vbCrLf _
                        & " where WHTLOCB1.LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.WHSE_CODE = '" & G.WHSE_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.LOCATION_QTY > 0 GROUP BY STYLE_CODE"

                    Dim rows5() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rows5.Length > 0 Then
                        ' Cases_count = rows5.Length
                        For Each ROW As DataRow In rows5
                            SSSS = SSSS & ", " & ROW.Item("STYLE_CODE") & " " ' & ROW.Item("DDD") & " Cartons) "
                        Next
                    End If
                    If Mid(SSSS, 1, 2) = ", " Then
                        SSSS = Mid(SSSS, 3)

                    End If

                    CreateResponse("SCAN_CASES", "G", "Location " & SCANTEXT & " has been selected, Styles in Location: " & SSSS)

 
                Case "SCAN_CASES"
                    '  CASES_PHYS = 0
                    If SCANTEXT = "CANCEL" Then

                        CreateResponse("SCAN_LOC", "R", "Cycle Cancelled, Re-scan  location")
                    Else


                        If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                            CreateResponse("", "R", "Invalid Number of cases " & SCANTEXT)
                            Exit Select
                        End If
                        CASES_PHYS = SCANTEXT

                        ' tbl.Rows.Add(New Object() {BAR_CODE})

                        '  CreateResponse("", "B", "Case Count " & CASES_PHYS & " Scanned")
                        CreateResponse("VERIFY", "B", "Scanned " & CASES_PHYS & " Cases")
                        Exit Select
                    End If


                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_LOC", "B", "Scan New Location, " & LOCATION_CODE & " Updated")
                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_CASES", "B", "Scan Ignored, Re-Scan LPN")

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

          CYCLE_RESOLUTION = "G"
        CYCLE_STATUS = "G"



        'ASCMAIN1.sql = " CYCLE_NEW = '1' OR CYCLE_SCAN IS NULL"
        'Dim rowWHTCYCL2() As DataRow = dst.Tables("WHTCYCL2").Select(ASCMAIN1.sql)
        'If rowWHTCYCL2.Length <> 0 Then
        '    CYCLE_RESOLUTION = "P"
        '    CYCLE_STATUS = "D"
        'End If


        'ASCMAIN1.sql = " CYCLE_SCAN = '1'"
        'Dim rowWHTCYCL2_COUNT() As DataRow = dst.Tables("WHTCYCL2").Select(ASCMAIN1.sql)
        'If rowWHTCYCL2_COUNT.Length <> 0 Then
        '    CASES_PHYS = rowWHTCYCL2_COUNT.Length
        'End If

        If Cases_count <> CASES_PHYS Then
            CYCLE_RESOLUTION = "P"
            CYCLE_STATUS = "D"
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
            .Item("CYCLE_TYPE") = "C"
            .Item("WHSE_CODE") = G.WHSE_CODE


        End With
        dst.Tables("WHTCYCL1").Rows.Add(rowWHTCYCL1)

        Update_Record_TDA("WHTCYCL1")
        '   Update_Record_TDA("WHTCYCL2")


        CommitTrans()
    End Sub


End Class
