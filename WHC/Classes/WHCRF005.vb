Public Class WHCRF005
    ' Deposit Application

    Inherits WHCRF000

    Dim BAR_CODE As String
    Dim CUST_CODE As String
    Dim ORDR_CUST_PO As String
    Dim WAVE_NO As String
    Dim LOCATION_CODE As String
    Dim LOCATION_CODE_DEPOSIT As String
    Dim LOAD_NO_DEPOSIT As String
    Dim hard_coded_wave As Boolean = False

    Sub New(g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF005"

        AppStates.Add("DEPOSIT", "Scan Location for Deposit|EXIT|")
        AppStates.Add("NODEPOSIT", "No Load to Deposit|EXIT|")
        AppStates.Add("VERIFY", "Are You Done (Y/N)|Y|N|CANCEL|")

        AppState = "DEPOSIT"

        ASCMAIN1.sql = "Select SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
        & ", WHTWAVE3.WAVE_NO, WHTINST1.LOCATION_CODE, WHTWAVE1.LOCATION_CODE_DEPOSIT, WHTWAVE1.LOAD_NO_DEPOSIT " & vbCrLf _
        & " from SOTORDR0,SOTSHIP1,WHTWAVE3,WHTINST1,WHTWAVE1" & vbCrLf _
        & " where SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
        & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
        & "   and WHTWAVE3.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
        & "   and WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
        & "   and WHTINST1.WAVE_INST_STATUS = '1'" & vbCrLf _
        & "   and ROWNUM = '1'" & vbCrLf _
        & "   AND WHTINST1.LOCATION_CODE_OTHER = '" & g.GUN_LOC & "'"

        If hard_coded_wave Then
            WAVE_NO = "0000000039"
            g.GUN_LOC = "99-G02-A"

            ASCMAIN1.sql = "Select SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ", WHTWAVE3.WAVE_NO, WHTINST1.LOCATION_CODE, WHTWAVE1.LOCATION_CODE_DEPOSIT, WHTWAVE1.LOAD_NO_DEPOSIT " & vbCrLf _
                & " from SOTORDR0,SOTSHIP1,WHTWAVE3,WHTINST1,WHTWAVE1" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = WHTWAVE3.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and WHTWAVE3.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                & "   and WHTWAVE1.WAVE_NO = WHTINST1.WAVE_NO" & vbCrLf _
                & "   and WHTINST1.WAVE_INST_STATUS = '1'" & vbCrLf _
                & "   and WHTINST1.WAVE_NO = '" & WAVE_NO & "'" _
                & "   AND WHTINST1.LOCATION_CODE_OTHER = '" & g.GUN_LOC & "'"
        End If

        Dim rowCUSTPO As DataRow = ASCDATA1.GetDataRow
        If rowCUSTPO Is Nothing Then
            CUST_CODE = ""
            ORDR_CUST_PO = ""
            WAVE_NO = ""
            LOCATION_CODE = ""
            LOCATION_CODE_DEPOSIT = ""
            LOAD_NO_DEPOSIT = ""
            AppState = "NODEPOSIT"
        Else
            CUST_CODE = rowCUSTPO.Item("CUST_CODE") & ""
            ORDR_CUST_PO = rowCUSTPO.Item("ORDR_CUST_PO") & ""
            WAVE_NO = rowCUSTPO.Item("WAVE_NO") & ""
            LOCATION_CODE = rowCUSTPO.Item("LOCATION_CODE") & ""
            LOCATION_CODE_DEPOSIT = rowCUSTPO.Item("LOCATION_CODE_DEPOSIT") & ""
            LOAD_NO_DEPOSIT = rowCUSTPO.Item("LOAD_NO_DEPOSIT") & ""

        End If

        With dst
            Create_TDA(.Tables.Add, "WHTBARC1", "*")
            Create_TDA(.Tables.Add, "WHTINST1", "*")
            Create_TDA(.Tables.Add, "WHTINST2", "*")
            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
        End With

        'tbl = dst.Tables("WHTSCANS") ' New DataTable

    End Sub

    Public Overrides Function Hello() As String

        Dim RESPONSE As String = G.THREAD_NO & ":" & G.APP_ID & ":" & G.APP_DESC & vbCrLf & Now.ToString & ":" & ASCMAIN1.USER_ID
        RESPONSE &= vbCrLf & "Deposit Wave: " & WAVE_NO & " PO: " & ORDR_CUST_PO & " for " & CUST_CODE & " to " & LOCATION_CODE_DEPOSIT
        RESPONSE &= vbCrLf & AppStates(AppState)
        Return RESPONSE
    End Function

    Public Overrides Sub GetResponseToScan(SCANTEXT As String)
        MyBase.GetResponseToScan(SCANTEXT)

        If SCANTEXT = "EXIT" Then
            ASCMAIN1.MultiTask_Release()
            CreateResponse("", "R", "EXIT")
        Else
            Select Case AppState
                Case "DEPOSIT"
                    If (SCANTEXT <> LOCATION_CODE_DEPOSIT) Then
                        CreateResponse("", "R", "Invalid Location, GOTO " & LOCATION_CODE_DEPOSIT)
                        Exit Select
                    End If
                    CreateResponse("VERIFY", "B", "Verify Deposit Wave: " & WAVE_NO & " PO: " & ORDR_CUST_PO & " for " & CUST_CODE & " to " & LOCATION_CODE_DEPOSIT)
                    Exit Select
                Case "NODEPOSIT"
                    ASCMAIN1.MultiTask_Release()
                    CreateResponse("", "R", "EXIT")

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("", "B", "EXIT")
                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("DEPOSIT", "B", "Scan Ignored, Re-Scan Shipto location")

                    ElseIf SCANTEXT = "CANCEL" Then

                        CreateResponse("DEPOSIT", "R", "Deposit Cancelled, Re-scan Shipto location")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub

    Sub Update_Record()

        ' IMPORTANT - this routine also exists in WHFWAVE1.Perform_Deposit

        BeginTrans()

        dst.Tables("WHTBARC1").Rows.Clear()

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        With rowWHTMOVE1
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_TYPE") = "D"
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

        ASCMAIN1.sql = "SELECT * FROM WHTINST1" & vbCrLf _
            & " where WAVE_INST_STATUS = '1'" _
            & " and WAVE_NO = '" & WAVE_NO & "'" _
            & " and LOCATION_CODE_OTHER = '" & G.GUN_LOC & "'"
        Fill_Records("WHTINST1", "", True, ASCMAIN1.sql)
        For Each rowWHTINST1 As DataRow In dst.Tables("WHTINST1").Select("", "WAVE_INST_NO, WAVE_NO")

            Dim WAVE_INST_NO As String = rowWHTINST1.Item("WAVE_INST_NO") & ""
            Dim LOAD_NO As String = rowWHTINST1.Item("LOAD_NO") & ""
            Dim LOCATION_CODE As String = rowWHTINST1.Item("LOCATION_CODE") & ""

            ASCMAIN1.sql = "SELECT * FROM WHTINST2 " & vbCrLf _
                    & " where WAVE_INST_NO = '" & WAVE_INST_NO & "'" _
                    & " and LOCATION_QTY_PICK > 0"
            Fill_Records("WHTINST2", "", True, ASCMAIN1.sql)
            For Each rowWHTINST2 As DataRow In dst.Tables("WHTINST2").Select("", "BAR_CODE, STYLE_CODE, COLOR_CODE")

                Dim BAR_CODE As String = rowWHTINST2.Item("BAR_CODE") & ""
                Dim rowWHTBARC1 As DataRow = Fill_Record("WHTBARC1", BAR_CODE, , False)

                Dim STYLE_CODE As String = rowWHTINST2.Item("STYLE_CODE") & ""
                Dim COLOR_CODE As String = rowWHTINST2.Item("COLOR_CODE") & ""
                Dim LOCATION_QTY_PICK As Integer = Val(rowWHTINST2.Item("LOCATION_QTY_PICK") & "")


                Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
                With rowWHTMOVE2
                    .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                    WHSE_TRAN_LNO_ctr += 1
                    .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                    .Item("LOCATION_CODE_FROM") = G.GUN_LOC
                    .Item("LOCATION_CODE_TO") = LOCATION_CODE_DEPOSIT
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("WHSE_TRAN_QTY") = LOCATION_QTY_PICK
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("INIT_OPER") = G.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("STATUS") = "U"
                    .Item("LOAD_NO_FROM") = LOAD_NO
                    .Item("LOAD_NO_TO") = LOAD_NO_DEPOSIT
                End With
                dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

                rowWHTBARC1.Item("LOAD_NO") = LOAD_NO_DEPOSIT
                'rowWHTBARC1.Item("LOCATION_CODE") = LOCATION_CODE_DEPOSIT
            Next

            rowWHTINST1.Item("WAVE_INST_STATUS") = "2"
        Next

        Update_Record_TDA("WHTBARC1")
        Update_Record_TDA("WHTINST1")
        Update_Record_TDA("WHTMOVE2")

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", _
                       New Object() {WHSE_TRAN_NO, 0, 1}, _
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        CommitTrans()
    End Sub
End Class
