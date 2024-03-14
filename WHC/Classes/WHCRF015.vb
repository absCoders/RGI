Public Class WHCRF015
    ' Application Find UPC

    Inherits WHCRF000

    Dim PICK_NO As String
    Dim PICK_LNO As String
    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim LOCATION_CODE As String
    Dim COLOR_CODEs As New List(Of String)
    Dim Cases_count As Integer
    Dim TICKET_NO As String
    Dim CASES_BOOK As Integer
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

        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style |EXIT|NEXT|EXTRA|") ' BLUE
        AppStates.Add("SCAN_COLOR", "Select a Color from List |CANCEL|")

        AppState = "SCAN_UPC"
        LAST_CLR = "BLUE"

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
                        End If
                    End If
                    PAGE_NO = 0
                    locations = FINDUPC(STYLE_CODE, COLOR_CODE, False)
                    CreateResponse("SCAN_UPC", "BLUE", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, locations))

            End Select
        End If
    End Sub

    Function FINDUPC(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, EXTRA As Boolean) As String
        Dim locations As String = ""
        Dim rows() As DataRow

        ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1, WHTLOCM1" & vbCrLf _
                       & " where WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                       & " and WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" & vbCrLf _
                       & " and nvl(WHTLOCM1.LOCATION_USE,'A') in ('A','E')" & vbCrLf _
                       & " and WHTLOCB1.WHSE_CODE = '" & G.WHSE_CODE & "'" _
                       & " and WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" _
                       & " and WHTLOCB1.COLOR_CODE = '" & COLOR_CODE & "'" _
                       & " order by abs(sign(WHTLOCB1.LOCATION_QTY)) DESC, WHTLOCB1.LOCATION_QTY DESC, WHTLOCB1.LAST_DATE DESC"
        If EXTRA Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "nvl(WHTLOCM1.LOCATION_USE,'A') in ('A','E')", "WHTLOCM1.LOCATION_CODE in ('00-RCV','00-RTS','00-SHP','00-DST')")
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
                    locations = locations & If(cnt = 1, "", vbCrLf & " ") & ROW.Item("WHSE_CODE") & ":" & ROW.Item("LOCATION_CODE") & " #" & ROW.Item("LOCATION_QTY") & " "
                End If
            Next
        End If
        If cnt < 6 Then
            PAGE_NO = 0
        End If

        Return locations

    End Function

End Class
