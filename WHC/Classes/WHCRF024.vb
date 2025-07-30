Public Class WHCRF024
    ' Application Move UPC to gun - for warehouses with no LPN

    Inherits WHCRF000

    Dim BAR_CODE As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim Cases_count As Integer
    Dim TICKET_NO As String
    Dim INVALID_BAR_CODE As String
    Dim CASES_RCVD As Integer
    Dim BAR_CODE_LOCATION As String
    Dim HOLD_PROMPT As String
    Dim Mode As String
    Dim LockWarning As Boolean

    Dim EDI_DOC_SEQ_NO As String = ""
    Dim PALLET_NO As String = ""
    Dim upcSummaryPageIndex As Integer = 0
    Dim upcSummaryPages As Integer = 1
    Const upcSummaryLinesPerPage As Integer = 5
    Dim upcVariancePageIndex As Integer = 0
    Dim upcVarincePages As Integer = 1
    Const upcVarianceLinesPerPage As Integer = 5
    Dim updateWithVariance As Boolean = False
    Dim addedUPCs As List(Of String)
    Dim scannedUPCs As Dictionary(Of String, String)



    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF024"

        AppStates.Add("SCAN_PALLET", "Scan PALLET ID|EXIT|")
        AppStates.Add("SCAN_UPC", "Scan UPC|DONE|") ' BLUE
        AppStates.Add("SCAN_SHOW", "NEXT|PREV|SCAN|") ' BLUE
        AppStates.Add("SCAN_LOC", "Scan Deposit Location") ' yellow
        AppStates.Add("LEAVE", "Units in Gun, Gun must be Empty for Receipt|EXIT|")
        AppStates.Add("VERIFY", "Update (Y/N)|Y|N|CANCEL|")

        AppStates.Add("SHOW_VAR", "Errors|NEXT|PREV|SCAN|UPDATE|") ' BLUE

        AppState = "SCAN_PALLET"
        LAST_CLR = ""

        With dst

            With .Tables.Add("WHTTRANV")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("UPC_CODE")
                .Columns.Add("VARIANCE", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            Create_TDA(.Tables.Add, "WHTTRAN1", "*")
            Create_TDA(.Tables.Add, "WHTTRAN2", "*")
            Create_TDA(.Tables.Add, "EDT945T2", "*")

        End With

        tbl = dst.Tables("WHTTRAN2") ' New DataTable

        Dim styles As String = TACMAIN1.LookupLocation(Me, g.GUN_LOC)
        If styles.Length > 0 Then
            AppState = "LEAVE"
        End If

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
                Case "SCAN_PALLET"
                    If SCANTEXT.Length <> 20 Then
                        CreateResponse("", "R", "Invalid Pallet Barcode")
                    Else

                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyTransferPallet(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                            Exit Select
                        End If
                        PALLET_NO = SCANTEXT

                        If Not ASCMAIN1.Logical_Lock("WHCRF024", SCANTEXT) Then
                            Dim User As String = ASCMAIN1.MultiTask_Get_Users("WHCRF024", SCANTEXT, "L")
                            If (Not User.Contains(G.USER_ID)) Or LockWarning = False Then
                                EMsg = vbCr & "Pallet is Locked " _
                                     & User
                                CreateResponse("", "GREEN", EMsg)
                                If User.Contains(G.USER_ID) Then LockWarning = True
                                Exit Select
                            End If
                        End If

                        Dim PICK_NO As String = SCANTEXT.Substring(0, 10)
                        Dim PALLET_SEQ_NO As Integer = CInt(SCANTEXT.Substring(10, 3))
                        Dim PICK_NO_USL As String = SCANTEXT.Substring(13)
                        Dim SCAN_DATE As DateTime = Now

                        EDI_DOC_SEQ_NO = CheckResponse("EDI_DOC_SEQ_NO")
                        dst.Tables("WHTTRAN1").Rows.Clear()
                        dst.Tables("WHTTRAN2").Rows.Clear()
                        dst.Tables("WHTTRANV").Rows.Clear()
                        addedUPCs = New List(Of String)
                        scannedUPCs = New Dictionary(Of String, String)

                        Dim rowWHTTRAN1 As DataRow = dst.Tables("WHTTRAN1").NewRow
                        rowWHTTRAN1("PALLET_NO") = PALLET_NO
                        rowWHTTRAN1("SCAN_DATE") = SCAN_DATE
                        rowWHTTRAN1("PICK_NO") = PICK_NO
                        rowWHTTRAN1("PALLET_SEQ_NO") = PALLET_SEQ_NO
                        rowWHTTRAN1("PICK_NO_USL") = PICK_NO_USL
                        rowWHTTRAN1("INIT_OPER") = G.USER_ID
                        rowWHTTRAN1("INIT_DATE") = SCAN_DATE
                        dst.Tables("WHTTRAN1").Rows.Add(rowWHTTRAN1)

                    End If

                    CreateResponse("SCAN_UPC", "BLUE", $"Pallet: {PALLET_NO}" & vbCrLf & "Scan UPC: ")

                Case "SCAN_UPC"
                    If SCANTEXT = "MINUS" Then
                        AppStates("SCAN_UPC") = "Scan UPC|PLUS|SHOW|DONE|"
                        Mode = SCANTEXT
                        CreateResponse("", "R", DisplayMsg("Mode was Changed"))
                        Exit Select
                    ElseIf SCANTEXT = "DONE" Then
                        If dst.Tables("WHTTRAN2").Rows.Count > 0 Then

                            If Receipt_Variance(EDI_DOC_SEQ_NO, PALLET_NO) Then
                                upcVariancePageIndex = 1
                                Dim pageText As String = GetUpcVariancePage(dst.Tables("WHTTRANV"), upcVariancePageIndex)
                                Dim nextBtn As String = If(upcVarincePages = 1, "", "NEXT|")
                                AppStates("SHOW_VAR") = $"Errors|{nextBtn}SCAN|UPDATE|"
                                CreateResponse("SHOW_VAR", "R", pageText)
                            Else
                                CreateResponse("VERIFY", "", "")
                            End If

                        Else
                            ClearScanner()
                            dst.Tables("WHTTRAN1").Rows.Clear()
                            dst.Tables("WHTTRAN2").Rows.Clear()
                            dst.Tables("WHTTRANV").Rows.Clear()
                            ASCMAIN1.MultiTask_Release()
                            CreateResponse("SCAN_PALLET", "YELLOW", DisplayMsg("Transfer Receipt Cancelled, Re - scan PALLET"))
                        End If

                        Exit Select
                    ElseIf SCANTEXT = "PLUS" Then
                        AppStates("SCAN_UPC") = "Scan UPC|MINUS|SHOW|DONE|"
                        Mode = SCANTEXT
                        CreateResponse("", "R", DisplayMsg("Mode was Changed"))
                        Exit Select
                    ElseIf SCANTEXT = "SHOW" Then
                        Dim nextBtn As String = If(upcSummaryPages = 1, "", "NEXT|")
                        AppStates("SCAN_SHOW") = $"Viewing Scans|{nextBtn}SCAN|"
                        upcSummaryPageIndex = 1
                        Dim pageText As String = GetUpcSummaryPage(dst.Tables("WHTTRAN2"), upcSummaryPageIndex)
                        CreateResponse("SCAN_SHOW", "R", pageText)
                        Exit Select
                    ElseIf SCANTEXT.Length <> 12 Then
                        CreateResponse("", "R", "Invalid Carton Barcode")
                        Exit Select
                    Else
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                            Exit Select
                        Else
                            CheckResponse = TACMAIN1.VerifyTransferUPC(Me, SCANTEXT, EDI_DOC_SEQ_NO, CheckResponse)
                            If CheckResponse.ContainsKey("Error") Then
                                CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                                Exit Select
                            End If
                            If CheckResponse.ContainsKey("Added") Then
                                If Not addedUPCs.Contains(SCANTEXT) Then
                                    addedUPCs.Add(SCANTEXT)
                                End If
                            End If
                        End If

                        Dim result As Object = dst.Tables("WHTTRAN2").Compute("SUM(SCAN_QTY)", $"UPC_CODE = '{SCANTEXT}'")
                        Dim currentTotalQty As Integer = If(IsDBNull(result), 0, Convert.ToInt32(result))

                        If Mode = "MINUS" AndAlso currentTotalQty = 0 Then
                            CreateResponse("", "R", DisplayMsg("Nothing to remove"))
                            Exit Select
                        End If

                        If Not scannedUPCs.ContainsKey(SCANTEXT) Then
                            scannedUPCs.Add(SCANTEXT, $"{CheckResponse("STYLE_CODE")}-{CheckResponse("COLOR_CODE")}")
                        End If

                        Dim rowWHTTRAN2 As DataRow = dst.Tables("WHTTRAN2").NewRow
                        rowWHTTRAN2("PALLET_NO") = PALLET_NO
                        rowWHTTRAN2("UPC_CODE") = SCANTEXT
                        rowWHTTRAN2("SCAN_QTY") = If(Mode = "MINUS", -1, 1)
                        rowWHTTRAN2("UPC_CODE_ADDED") = If(addedUPCs.Contains(SCANTEXT), "1", "0")
                        rowWHTTRAN2("SCAN_DATE") = Now
                        dst.Tables("WHTTRAN2").Rows.Add(rowWHTTRAN2)
                        Dim btnToShow As String = If(Mode = "MINUS", "PLUS", "MINUS")
                        AppStates("SCAN_UPC") = $"Scan UPC|{btnToShow}|SHOW|DONE|"
                        CreateResponse("SCAN_UPC", "BLUE", DisplayMsg(""))
                        Exit Select
                    End If

                Case "SCAN_SHOW"
                    If SCANTEXT = "SCAN" Then
                        CreateResponse("SCAN_UPC", "BLUE", DisplayMsg(""))
                        Exit Select
                    ElseIf SCANTEXT = "NEXT" Then
                        upcSummaryPageIndex += 1
                        Dim pageText As String = GetUpcSummaryPage(dst.Tables("WHTTRAN2"), upcSummaryPageIndex)
                        If upcSummaryPageIndex = upcSummaryPages Then
                            AppStates("SCAN_SHOW") = "Viewing Scans|PREV|SCAN|"
                        Else
                            AppStates("SCAN_SHOW") = "Viewing Scans|NEXT|PREV|SCAN|"
                        End If
                        CreateResponse("SCAN_SHOW", "R", pageText)
                        Exit Select
                    ElseIf SCANTEXT = "PREV" Then
                        upcSummaryPageIndex -= 1
                        Dim pageText As String = GetUpcSummaryPage(dst.Tables("WHTTRAN2"), upcSummaryPageIndex)
                        If upcSummaryPageIndex = 1 Then
                            AppStates("SCAN_SHOW") = "Viewing Scans|NEXT|SCAN|"
                        Else
                            AppStates("SCAN_SHOW") = "Viewing Scans|NEXT|PREV|SCAN|"
                        End If
                        CreateResponse("SCAN_SHOW", "R", pageText)
                        Exit Select
                    End If
                Case "SCAN_LOC"

                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "R", DisplayMsg(CheckResponse("Error")))
                        Exit Select
                    Else
                        Deposit_Transfer_From_Gun(SCANTEXT)
                        ClearScanner()
                        CreateResponse("SCAN_PALLET", "YELLOW", DisplayMsg(""))
                    End If
                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_LOC", "Y", $"Deposit Pallet {PALLET_NO}")

                    ElseIf SCANTEXT = "N" Then
                        'ClearScanner()
                        Dim SCAN_QTY As Integer = CInt(dst.Tables("WHTTRAN2").Compute("SUM(SCAN_QTY)", ""))

                        CreateResponse("SCAN_UPC", "BLUE", $"Pallet: {PALLET_NO & $":{SCAN_QTY}" & vbCrLf}Scan UPC:")
                    ElseIf SCANTEXT = "CANCEL" Then
                        ClearScanner()
                        dst.Tables("WHTTRAN1").Rows.Clear()
                        dst.Tables("WHTTRAN2").Rows.Clear()
                        dst.Tables("WHTTRANV").Rows.Clear()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_PALLET", "YELLOW", DisplayMsg("Transfer Receipt Cancelled, Re - scan PALLET"))
                    Else
                        CreateResponse("", "R", DisplayMsg("Invalid Response"))
                    End If

                Case "SHOW_VAR"
                    If SCANTEXT = "SCAN" Then
                        CreateResponse("SCAN_UPC", "BLUE", DisplayMsg(""))
                        Exit Select
                    ElseIf SCANTEXT = "NEXT" Then
                        upcVariancePageIndex += 1
                        Dim pageText As String = GetUpcVariancePage(dst.Tables("WHTTRANV"), upcVariancePageIndex)
                        If upcVariancePageIndex = upcVarincePages Then
                            AppStates("SHOW_VAR") = "Errors|PREV|SCAN|UPDATE|"
                        Else
                            AppStates("SHOW_VAR") = "Errors|NEXT|PREV|SCAN|UPDATE|"
                        End If
                        CreateResponse("SHOW_VAR", "R", pageText)
                        Exit Select
                    ElseIf SCANTEXT = "PREV" Then
                        upcVariancePageIndex -= 1
                        Dim pageText As String = GetUpcVariancePage(dst.Tables("WHTTRANV"), upcVariancePageIndex)
                        If upcVariancePageIndex = 1 Then
                            AppStates("SHOW_VAR") = "Errors|NEXT|SCAN|UPDATE|"
                        Else
                            AppStates("SHOW_VAR") = "Errors|NEXT|PREV|SCAN|UPDATE|"
                        End If
                        CreateResponse("SHOW_VAR", "R", pageText)
                        Exit Select
                    ElseIf SCANTEXT = "UPDATE" Then
                        updateWithVariance = True
                        CreateResponse("VERIFY", "", "")
                    End If
            End Select
        End If
    End Sub


    Function DisplayMsg(ByVal note As String) As String
        Dim msg As String = ""
        Dim FromLoc As String = ""

        If Mode = "MINUS" Then
            msg = "Scan to remove a carton"
            'FromLoc = LOCATION_CODE
        Else
            msg = "Scan to add a carton"
        End If
        Dim summary As String = GetScannedUpcSummary(dst.Tables("WHTTRAN2"))

        If note <> "" Then
            msg = msg & vbCrLf & summary & vbCrLf & note
        Else
            msg = msg & vbCrLf & summary
        End If

        If PALLET_NO <> "" Then
            Dim SCAN_QTY As Integer = CInt(If(dst.Tables("WHTTRAN2").Rows.Count > 0, dst.Tables("WHTTRAN2").Compute("SUM(SCAN_QTY)", ""), 0))
            CreateResponse("SCAN_UPC", "BLUE", $"Pallet: {PALLET_NO}:{SCAN_QTY & vbCrLf}Scan UPC:")
            msg = $"Pallet: {PALLET_NO}:{SCAN_QTY & vbCrLf & msg}"
        End If
        Return msg
    End Function

    Sub ClearScanner()
        UPC_CODE = ""
        EDI_DOC_SEQ_NO = ""
        PALLET_NO = ""
        updateWithVariance = False
    End Sub

    Sub Update_Record()

        Dim FromLoc As String = ASCDATA1.GetDataValue($"Select WHSE_LOC_XIN from ICTWHSE1 where WHSE_CODE = '{G.WHSE_CODE}'")
        Dim ToLoc As String = G.GUN_LOC
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

        Dim WHSE_TRAN_LNO_ctr As Integer = 0
        Dim distinctUpcs = dst.Tables("WHTTRAN2").DefaultView.ToTable(True, "UPC_CODE")
        For Each row As DataRow In distinctUpcs.Rows
            Dim UPC_CODE_sum As String = row("UPC_CODE")
            Dim totalQty As Integer = Convert.ToInt32(dst.Tables("WHTTRAN2").Compute("SUM(SCAN_QTY)", $"UPC_CODE = '{UPC_CODE_sum}'"))
            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, UPC_CODE_sum)
            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
            Dim CtnQty = CInt(CheckResponse("CARTON_PACK_QTY"))
            If CtnQty = 0 Then CtnQty = 1

            With rowWHTMOVE2
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                WHSE_TRAN_LNO_ctr += 1
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                .Item("LOCATION_CODE_FROM") = FromLoc
                .Item("LOCATION_CODE_TO") = ToLoc
                .Item("BAR_CODE") = "0000000000"
                .Item("WHSE_TRAN_QTY") = totalQty * CtnQty
                .Item("STYLE_CODE") = CheckResponse("STYLE_CODE")
                .Item("COLOR_CODE") = CheckResponse("COLOR_CODE")
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"
                .Item("LOAD_NO_FROM") = ""
                .Item("LOAD_NO_TO") = ""
            End With
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
        Next

        dst.Tables("WHTTRAN1").Rows(0)("LAST_OPER") = G.USER_ID
        dst.Tables("WHTTRAN1").Rows(0)("LAST_DATE") = Now

        BeginTrans()

        Update_Record_TDA("WHTTRAN1")
        Update_Record_TDA("WHTTRAN2")
        Update_Record_TDA("WHTMOVE1")
        Update_Record_TDA("WHTMOVE2")

        Dim sql As String = "Update EDT945T3 Set EDI_STATUS = 'F' Where EDI_PALLET_NO = '" & PALLET_NO & "'"
        ASCDATA1.ExecuteSQL(sql)

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        If updateWithVariance Then
            ASCDATA1.ExecuteSQL($"Update WHTTRAN1 Set HAS_VARIANCE = '1' Where PALLET_NO = '{PALLET_NO}'")
        End If

        CommitTrans()



    End Sub

    Function GetScannedUpcSummary(dt As DataTable) As String

        Dim results As New List(Of String)()

        ' Get distinct UPC_CODE values
        Dim distinctUpcs = dt.DefaultView.ToTable(True, "UPC_CODE")

        For Each row As DataRow In distinctUpcs.Rows
            Dim UPC_CODE_sum As String = row("UPC_CODE")
            Dim totalQty As Integer = Convert.ToInt32(dt.Compute("SUM(SCAN_QTY)", $"UPC_CODE = '{UPC_CODE_sum}'"))
            results.Add($"{UPC_CODE_sum}:{totalQty}")
        Next
        Return String.Join(", ", results)
    End Function

    Function GetUpcSummaryPage(dt As DataTable, pageIndex As Integer) As String

        Dim upcList As New List(Of String)()

        ' Get distinct UPCs
        Dim distinctUpcs = dt.DefaultView.ToTable(True, "UPC_CODE")

        For Each row As DataRow In distinctUpcs.Rows

            Dim upc As String = row("UPC_CODE").ToString()
            Dim escapedUpc As String = upc.Replace("'", "''")
            Dim totalQty As Integer = Convert.ToInt32(dt.Compute("SUM(SCAN_QTY)", $"UPC_CODE = '{escapedUpc}'"))
            Dim keyToShow As String = If(scannedUPCs.ContainsKey(upc), scannedUPCs(upc), upc)
            upcList.Add($"{keyToShow}:{totalQty}")
        Next

        ' Total pages
        Dim totalPages As Integer = Math.Ceiling(upcList.Count / upcSummaryLinesPerPage)
        upcSummaryPages = totalPages

        ' Clamp page index
        If pageIndex < 1 Then pageIndex = 1
        If pageIndex > totalPages Then pageIndex = totalPages

        ' Get subset of UPCs for current page
        Dim startIdx As Integer = (pageIndex - 1) * upcSummaryLinesPerPage
        Dim pagedList = upcList.Skip(startIdx).Take(upcSummaryLinesPerPage)

        ' Build output string
        Dim output As New System.Text.StringBuilder()
        For Each entry As String In pagedList
            output.AppendLine(entry)
        Next

        output.Append($"Page {pageIndex}/{totalPages}")

        Return output.ToString()
    End Function

    Function GetUpcVariancePage(dt As DataTable, pageIndex As Integer) As String

        Dim upcList As New List(Of String)()


        For Each row As DataRow In dt.Select("")
            Dim STYLE_CODE As String = row("STYLE_CODE").ToString()
            Dim COLOR_CODE As String = row("COLOR_CODE").ToString()
            Dim UPC_CODE As String = row("UPC_CODE").ToString()
            Dim VARIANCE As Integer = Convert.ToInt32(row("VARIANCE"))
            upcList.Add($"{STYLE_CODE}-{COLOR_CODE}:{VARIANCE}")
        Next

        ' Total pages
        Dim totalPages As Integer = Math.Ceiling(upcList.Count / upcVarianceLinesPerPage)
        upcVarincePages = totalPages

        ' Clamp page index
        If pageIndex < 1 Then pageIndex = 1
        If pageIndex > totalPages Then pageIndex = totalPages

        ' Get subset of UPCs for current page
        Dim startIdx As Integer = (pageIndex - 1) * upcVarianceLinesPerPage
        Dim pagedList = upcList.Skip(startIdx).Take(upcVarianceLinesPerPage)

        ' Build output string
        Dim output As New System.Text.StringBuilder()
        For Each entry As String In pagedList
            output.AppendLine(entry)
        Next

        output.Append($"Page {pageIndex}/{totalPages}")

        Return output.ToString()
    End Function

    Sub Deposit_Transfer_From_Gun(depositLocation As String)
        Dim FromLoc As String = G.GUN_LOC
        Dim ToLoc As String = depositLocation
        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

        dst.Tables("WHTMOVE1").Clear()
        dst.Tables("WHTMOVE2").Clear()

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

        Dim WHSE_TRAN_LNO_ctr As Integer = 0

        Dim sqlWHTLOCB1 As String = $"select * from WHTLOCB1 where WHSE_CODE = '{G.WHSE_CODE}' and LOCATION_CODE = '{FromLoc}' and location_QTY > 0"
        Dim tblWHTLOCB1 As DataTable = ASCDATA1.GetDataTable(sqlWHTLOCB1)

        For Each row As DataRow In tblWHTLOCB1.Rows
            Dim totalQty As Integer = row.Item("LOCATION_QTY")

            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow

            With rowWHTMOVE2
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                WHSE_TRAN_LNO_ctr += 1
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                .Item("LOCATION_CODE_FROM") = FromLoc
                .Item("LOCATION_CODE_TO") = ToLoc
                .Item("BAR_CODE") = "0000000000"
                .Item("WHSE_TRAN_QTY") = totalQty
                .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                .Item("INIT_OPER") = G.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"
                .Item("LOAD_NO_FROM") = ""
                .Item("LOAD_NO_TO") = ""
            End With
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
        Next

        BeginTrans()

        Update_Record_TDA("WHTMOVE1")
        Update_Record_TDA("WHTMOVE2")

        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                       New Object() {WHSE_TRAN_NO, 0, 1},
                       New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})
        CommitTrans()

    End Sub

    Function Receipt_Variance(EDI_DOC_SEQ_NO As String, EDI_PALLET_NO As String) As Boolean

        dst.Tables("WHTTRANV").Rows.Clear()

        Dim sqlEDT945T2 As String = $"SELECT * FROM EDT945T2 WHERE EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}' AND EDI_PALLET_NO = '{EDI_PALLET_NO}'"
        Fill_Records("EDT945T2", "", True, sqlEDT945T2)
        Dim scannedUPCs As New List(Of String)
        Dim distinctUpcs = dst.Tables("WHTTRAN2").DefaultView.ToTable(True, "UPC_CODE")

        For Each row As DataRow In distinctUpcs.Rows
            scannedUPCs.Add(row("UPC_CODE") & "")
            Dim UPC_CODE_sum As String = row("UPC_CODE")
            Dim WHTTRAN2_Qty As Integer = Convert.ToInt32(dst.Tables("WHTTRAN2").Compute("SUM(SCAN_QTY)", $"UPC_CODE = '{UPC_CODE_sum}'"))
            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, UPC_CODE_sum)
            Dim CtnQty = CInt(CheckResponse("CARTON_PACK_QTY"))
            If CtnQty = 0 Then CtnQty = 1
            Dim WHSE_TRAN_QTY As Integer = WHTTRAN2_Qty * CtnQty
            Dim result As Object = dst.Tables("EDT945T2").Compute("SUM(EDI_SHIP_QTY)", $"UPC_CODE = '{UPC_CODE_sum}'")
            Dim EDT945T2_Qty As Integer = If(IsDBNull(result), 0, Convert.ToInt32(result))

            If WHSE_TRAN_QTY <> EDT945T2_Qty Then
                Dim rowWHTTRANV As DataRow = dst.Tables("WHTTRANV").NewRow
                rowWHTTRANV("STYLE_CODE") = CheckResponse("STYLE_CODE")
                rowWHTTRANV("COLOR_CODE") = CheckResponse("COLOR_CODE")
                rowWHTTRANV("UPC_CODE") = CheckResponse("UPC_CODE")
                rowWHTTRANV("VARIANCE") = (EDT945T2_Qty - WHSE_TRAN_QTY) * -1
                dst.Tables("WHTTRANV").Rows.Add(rowWHTTRANV)
            End If

        Next

        Dim EDT945T2_UPCs = dst.Tables("EDT945T2").DefaultView.ToTable(True, "UPC_CODE")
        For Each row As DataRow In EDT945T2_UPCs.Rows
            Dim UPC_CODE As String = row("UPC_CODE")
            If Not scannedUPCs.Contains(UPC_CODE) Then
                Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, UPC_CODE)
                Dim rowWHTTRANV As DataRow = dst.Tables("WHTTRANV").NewRow
                rowWHTTRANV("STYLE_CODE") = CheckResponse("STYLE_CODE")
                rowWHTTRANV("COLOR_CODE") = CheckResponse("COLOR_CODE")
                rowWHTTRANV("UPC_CODE") = UPC_CODE
                Dim result As Object = dst.Tables("EDT945T2").Compute("SUM(EDI_SHIP_QTY)", $"UPC_CODE = '{UPC_CODE}'")
                Dim EDT945T2_Qty As Integer = If(IsDBNull(result), 0, Convert.ToInt32(result))
                rowWHTTRANV("VARIANCE") = EDT945T2_Qty * -1
                dst.Tables("WHTTRANV").Rows.Add(rowWHTTRANV)
            End If
        Next
        Return (dst.Tables("WHTTRANV").Rows.Count > 0)
    End Function
End Class
