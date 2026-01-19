Public Class WHCRF015
    ' Application Find UPC

    Inherits WHCRF000

    Dim PICK_NO As String
    Dim PICK_LNO As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim STYLE_DESC As String
    Dim COLOR_DESC As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim COLOR_CODEs As New List(Of String)
    Dim Cases_count As Integer
    Dim TICKET_NO As String
    Dim CARTON_PACK_QTY As Integer
    Dim CASES_MOVED As Integer
    Dim UNITS_MOVED As Integer
    Dim TICKET_NO1 As String
    Dim BAR_CODE_LOCATION As String
    Dim colors As String = ""
    Dim locations As String = ""
    Dim PAGE_NO As Integer = 0

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF015"

        If "rick,victor,teo,baudilio,".Contains(g.USER_ID & ",") Then
            AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style |EXIT|NEXT|EXTRA|+OH|") ' BLUE
        Else
            AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style |EXIT|NEXT|EXTRA|") ' BLUE
        End If
        AppStates.Add("SCAN_COLOR", "Select a Color from List |CANCEL|")
        AppStates.Add("SCAN_LOCATION", "Scan Location found |CANCEL|") ' Yellow
        AppStates.Add("SCAN_CASES", "How many cases found, (0 for units)|CANCEL|")
        AppStates.Add("SCAN_UNITS", "Enter units found |CANCEL|") 'G
        AppStates.Add("VERIFY", "Update|Y|N|EXIT|")

        AppState = "SCAN_UPC"
        LAST_CLR = "BLUE"

        With dst
            Create_TDA(.Tables.Add, "WHTPICKS", "*")
        End With
        tbl = dst.Tables("WHTPICKS")

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
                Case "SCAN_UPC"
                    If SCANTEXT = "EXTRA" Then
                        locations = FINDUPC(STYLE_CODE, COLOR_CODE, True)
                        PAGE_NO = 0
                        CreateResponse("SCAN_UPC", "BLUE", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, locations))
                        Exit Select
                    ElseIf SCANTEXT = "NEXT" Then
                        locations = FINDUPC(STYLE_CODE, COLOR_CODE, False)
                        CreateResponse("SCAN_UPC", "BLUE", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, locations))
                        Exit Select
                    ElseIf SCANTEXT = "+OH" Then
                        CreateResponse("SCAN_LOCATION", "YELLOW", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, "Positive Adjustment"))
                        Exit Select
                    End If

                    If SCANTEXT.Length = 14 Then
                        SCANTEXT = SCANTEXT.Substring(0, 12)
                    End If
                    Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                    If CheckResponse.ContainsKey("Error") Then
                        CreateResponse("", "R", CheckResponse("Error"))
                        Exit Select
                    End If

                    If CheckResponse.ContainsKey("UPC_CODE") Then
                        UPC_CODE = CheckResponse("UPC_CODE")
                        STYLE_CODE = CheckResponse("STYLE_CODE")
                        COLOR_CODE = CheckResponse("COLOR_CODE")
                        STYLE_DESC = CheckResponse("STYLE_DESC")
                        COLOR_DESC = CheckResponse("COLOR_DESC")
                        PAGE_NO = 0
                        locations = FINDUPC(STYLE_CODE, COLOR_CODE, False)
                        CreateResponse("", "BLUE", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, locations))
                        Exit Select
                    End If
                    STYLE_CODE = SCANTEXT.ToUpper
                    TACMAIN1.GetColors(Me, STYLE_CODE, LOCATION_CODE, COLOR_CODEs, colors)

                    CreateResponse("SCAN_COLOR", "G", "Style " & STYLE_CODE & " has been selected, colors " & colors)

                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Cancelled")
                        Exit Select
                    Else
                        If SCANTEXT = "0" Then
                            TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)
                            CreateResponse("", "B", "Style " & STYLE_CODE & " has been selected, All colors " & colors)
                            Exit Select
                        Else
                            Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.VerifyColor(Me, SCANTEXT, STYLE_CODE, COLOR_CODEs)
                            If CheckResponse.ContainsKey("Error") Then
                                CreateResponse("", "R", CheckResponse("Error"))
                                Exit Select
                            End If
                            UPC_CODE = CheckResponse("UPC_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            STYLE_DESC = CheckResponse("STYLE_DESC")
                            COLOR_DESC = CheckResponse("COLOR_DESC")
                        End If
                    End If
                    PAGE_NO = 0
                    locations = FINDUPC(STYLE_CODE, COLOR_CODE, False)
                    CreateResponse("SCAN_UPC", "BLUE", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, locations))

                Case "SCAN_LOCATION"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Cancelled")
                        Exit Select
                    Else
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckLocation(Me, SCANTEXT, False)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", CheckResponse("Error"))
                            Exit Select
                        End If
                        BAR_CODE_LOCATION = CheckResponse("LOCATION_CODE")
                    End If
                    CreateResponse("SCAN_CASES", "BLUE", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, LOCATION_CODE))
                    dst.Tables("WHTPICKS").Rows.Clear()

                Case "SCAN_CASES"
                    '  Asking for Cases although Message may not display number of cases requested
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Cancelled")
                        Exit Select
                    Else
                        Dim hold As String
                        If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Number, Enter Cases Found|OK|"
                            CreateResponse("", "R", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, LOCATION_CODE))
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
                            CreateResponse("SCAN_UNITS", "G", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, LOCATION_CODE))
                            Exit Select
                        End If
                        ASCMAIN1.sql = "select CARTON_PACK_QTY from ICTSTYL1 Where STYLE_CODE = '" & STYLE_CODE & "'"
                        CARTON_PACK_QTY = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql) & "")
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = "Update " & SCANTEXT & " Cases|Y|N|CLEAR|EXIT|"
                        CreateResponse("VERIFY", "B", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, LOCATION_CODE))
                        AppStates("VERIFY") = hold
                    End If

                Case "SCAN_UNITS"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Cancelled")
                        Exit Select
                    Else
                        Dim hold As String
                        'Can we have more than 999 loose units to Move in a Pick?
                        If Val(SCANTEXT) > 999 Or Val(SCANTEXT) < 0 Then
                            'Error
                            hold = AppStates(AppState)
                            AppStates(AppState) = "Invalid Count " & SCANTEXT & ", Verify Units Count|OK|"
                            CreateResponse("", "R", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, LOCATION_CODE))
                            AppStates(AppState) = hold
                            Exit Select
                        End If
                        UNITS_MOVED = Val(SCANTEXT)
                        hold = AppStates("VERIFY")
                        AppStates("VERIFY") = "Update " & SCANTEXT & " Units|Y|N|CLEAR|EXIT|"
                        CreateResponse("VERIFY", "B", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, LOCATION_CODE))
                        AppStates("VERIFY") = hold
                    End If

                Case "VERIFY"
                    If SCANTEXT.ToUpper = "Y" Then
                        ' Y is for normal Verify
                        Update_Record()
                        CreateResponse("SCAN_UPC", "BLUE", "Updated")
                    ElseIf SCANTEXT.ToUpper = "N" Or SCANTEXT = "NO" Then
                        ' N is for normal Verify
                        ' NO is for Clear Verify
                        CreateResponse("SCAN_UPC", "BLUE", "Cancelled")
                    Else
                        'Error
                        Dim hold As String = AppStates(AppState)
                        AppStates(AppState) = "Invalid Response|OK|"
                        CreateResponse("", "R", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, LOCATION_CODE))
                        AppStates(AppState) = hold

                    End If
            End Select
        End If
    End Sub

    Function FINDUPC(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, EXTRA As Boolean) As String
        Dim locations As String = ""
        Dim rows() As DataRow
        Dim LOCATION_USE As String

        LOCATION_USE = ASCDATA1.GetDataValue($"Select LOCATION_USE from WHTSTLC1 WHERE STYLE_CODE = '{STYLE_CODE}'")
        If LOCATION_USE & "" <> "" Then
            locations = "Class III-IV Item" & vbCrLf
        End If
        ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1, WHTLOCM1" & vbCrLf _
                       & " where WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                       & " and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
                       & " and nvl(WHTLOCM1.LOCATION_USE,'A') in ('A','E','C')" & vbCrLf _
                       & " and WHTLOCB1.WHSE_CODE = '" & G.WHSE_CODE & "'" _
                       & " and WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" _
                       & " and WHTLOCB1.COLOR_CODE = '" & COLOR_CODE & "'" _
                       & " order by abs(sign(WHTLOCB1.LOCATION_QTY)) DESC, WHTLOCB1.LOCATION_QTY DESC, WHTLOCB1.LAST_DATE DESC"
        If EXTRA Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "nvl(WHTLOCM1.LOCATION_USE,'A') in ('A','E','C')", "WHTLOCM1.LOCATION_CODE in ('00-RCV','00-RTS','00-SHP','00-DST')")
        End If
        rows = ASCDATA1.GetDataTable.Select("")
        Dim cnt As Int32 = 0
        Dim rownum As Int32 = 0
        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                rownum += 1
                If PAGE_NO * 5 <= rownum Then
                    cnt += 1
                    If cnt > 5 Then
                        locations = locations & vbCrLf & "More ..." & vbCrLf
                        PAGE_NO += 1
                        Exit For
                    End If
                    locations = locations & If(cnt = 1, " ", vbCrLf & " ") & ROW.Item("WHSE_CODE") & ":" & ROW.Item("LOCATION_CODE") & " #" & ROW.Item("LOCATION_QTY") & " "
                End If
            Next
        End If
        If cnt < 6 Then
            PAGE_NO = 0
        End If

        Return locations

    End Function
    Sub Update_Record()
        Dim HoldLoc As String = ""
        Dim QtyFound = (CASES_MOVED * CARTON_PACK_QTY + UNITS_MOVED)

        BeginTrans()
        Dim PICK_NO As String = "OH" & ASCMAIN1.Next_Control_No("WHTPICKS.ADD_OH")
        Dim rowWHTPICKS As DataRow = dst.Tables("WHTPICKS").NewRow
        With rowWHTPICKS
            .Item("PICK_NO") = PICK_NO
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("STYLE_DESC") = STYLE_DESC
            .Item("COLOR_DESC") = COLOR_DESC
            .Item("LOCATION_CODE") = BAR_CODE_LOCATION
            .Item("SHORTAGE") = QtyFound
            .Item("STATUS") = "A"
            .Item("INIT_OPER") = G.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
        End With
        dst.Tables("WHTPICKS").Rows.Add(rowWHTPICKS)

        Update_Record_TDA("WHTPICKS")

        ASCMAIN1.sql = "INSERT INTO ASTNOTEM " &
                       "Select 'SHORTAGES' NOTE_CODE, " &
                       "NVL((SELECT max(SEND_LNO) FROM ASTNOTEM WHERE NOTE_CODE = 'SHORTAGES'), 0) + 1 SEND_LNO, " &
                       $"'Units Found {STYLE_CODE}-{COLOR_CODE}' NOTE_MEMO " &
                       "from DUAL"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        CommitTrans()

        CASES_MOVED = 0
        UNITS_MOVED = 0

    End Sub

End Class
