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

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF015"

        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style |EXIT|") ' BLUE
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
                        Dim locations As String = FINDUPC(STYLE_CODE, COLOR_CODE)
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
                    Dim locations As String = FINDUPC(STYLE_CODE, COLOR_CODE)
                    CreateResponse("SCAN_UPC", "BLUE", String.Format("{0} {1} {2} {3} {4}", UPC_CODE, STYLE_CODE, COLOR_CODE, vbCrLf, locations))

            End Select
        End If
    End Sub

    Function FINDUPC(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim locations As String = ""
        Dim rows() As DataRow

        ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1" & vbCrLf _
                       & " where  WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" _
                       & " and COLOR_CODE = '" & COLOR_CODE & "'" _
                       & " order by LOCATION_QTY DESC"
        rows = ASCDATA1.GetDataTable.Select("", "LOCATION_QTY DESC")
        Dim cnt As Int32 = 0
        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                If cnt > 5 Then
                    locations = locations & vbCrLf & "More ..."
                    Exit For
                End If
                locations = locations & If(cnt = 1, "", vbCrLf) & ROW.Item("WHSE_CODE") & ":" & ROW.Item("LOCATION_CODE") & " #" & ROW.Item("LOCATION_QTY") & " "
            Next
        End If

        Return locations

    End Function

End Class
