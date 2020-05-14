Public Class WHCRF019
    ' Application Move Entire Location - for warehouses with no LPN

    Inherits WHCRF000

    Dim BAR_CODE As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim FROM_LOCATION As String
    Dim TO_LOCATION As String
    Dim Styles As String
    Dim TICKET_NO As String
    
    Dim CASES_MOVED As Integer
    Dim UNITS_MOVED As Integer
    Dim TICKET_NO1 As String

    Dim Mode As String
    Dim colors As String = ""

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF019"

        AppStates.Add("SCAN_FROM", "Scan Move From Location|EXIT|") ' YELLOW
        AppStates.Add("SCAN_TO", "Scan Move To Location |USE GUN|CANCEL|") ' BLUE
        AppStates.Add("VERIFY", "Update (Y/N)|Y|N|CANCEL|")

        AppState = "SCAN_FROM"
        LAST_CLR = ""

        With dst
            '  With .Tables.Add("WHTSCANS").Columns
            ' .Add("BAR_CODE")
            ' .Add("NEW")
            ' .Add("SCANNED")
            ' End With
            ' .Tables("WHTSCANS").PrimaryKey = New DataColumn() {.Tables("WHTSCANS").Columns("BAR_CODE")}

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
                Case "SCAN_FROM"
                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                        Exit Select
                    End If
                    If CheckResponse("Stylelist") = "" Then
                        CreateResponse("", "R", "Emprty From Location :" & SCANTEXT.ToUpper & vbCrLf & "Verify Location")
                        Exit Select
                    End If

                    FROM_LOCATION = SCANTEXT.ToUpper
                    Styles = CheckResponse("Stylelist")

                    dst.Tables("WHTMOVE1").Rows.Clear()
                    dst.Tables("WHTMOVE2").Rows.Clear()

                    CreateResponse("SCAN_TO", "YELLOW", DisplayMsg("Styles in Location: " & Styles))


                Case "SCAN_TO"
                    If SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_FROM", "YELLOW", "Move Cancelled, Scan From location")
                        Exit Select
                    Else
                        If SCANTEXT = "USE GUN" Then
                            SCANTEXT = G.GUN_LOC
                        End If

                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                            Exit Select
                        End If

                        Dim Styles As String = CheckResponse("Stylelist")
                        'If Styles <> "" Then
                        '    CreateResponse("", "R", DisplayMsg("To Location not Empty, Select an empty location"))
                        '    Exit Select
                        'End If

                        TO_LOCATION = SCANTEXT.ToUpper
                    End If
                    CreateResponse("VERIFY", "G", DisplayMsg("Styles to move: " & Styles))

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_FROM", "YELLOW", DisplayMsg(""))
                    ElseIf SCANTEXT = "N" Or SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        CreateResponse("SCAN_FROM", "YELLOW", DisplayMsg("Move Cancelled, Re - scan  location"))
                    Else
                        CreateResponse("", "R", DisplayMsg("Invalid Response"))
                    End If
            End Select
        End If
    End Sub

    Function DisplayMsg(ByVal note As String) As String
        Dim msg As String = "Move Entire Location"
        
        If FROM_LOCATION <> "" Then
            msg = msg & vbCrLf & "FROM LOC: " & FROM_LOCATION
        End If

        If TO_LOCATION <> "" Then
            msg = msg & vbCrLf & "TO LOC: " & TO_LOCATION
        End If
        

        If note <> "" Then
            msg = msg & vbCrLf & note
        End If
        Return msg
    End Function

    Sub ClearScanner()

        LOCATION_CODE = ""
        FROM_LOCATION = ""
        TO_LOCATION = ""
        Styles = ""
        UPC_CODE = ""
        STYLE_CODE = ""
        COLOR_CODE = ""
        UNITS_MOVED = 0

    End Sub

    Sub Update_Record()

        Dim FromLoc As String = FROM_LOCATION
        Dim ToLoc As String = TO_LOCATION

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

        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, LOCATION_QTY from WHTLOCB1 " _
                        & " WHERE LOCATION_CODE = '" & FROM_LOCATION & "' and LOCATION_QTY <> 0"
        Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                STYLE_CODE = ROW("STYLE_CODE")
                COLOR_CODE = ROW("COLOR_CODE")
                UNITS_MOVED = ROW("LOCATION_QTY")

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
            Next
        End If

        Update_Record_TDA("WHTMOVE2")

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        ClearScanner()
        CommitTrans()
    End Sub


End Class
