Public Class WHCRF010
    ' Application Request UPC Labels

    Inherits WHCRF000

    Dim STYLE_CODE As String
    Dim COLOR_CODE As String
    Dim UPC_CODE As String
    Dim COLOR_CODEs As New List(Of String)
    Dim LBL_REQ_NO As String
    Dim Request_count As Integer
    Dim colors As String = ""
    Dim CARTONS_PER_UNIT As Integer
    Dim LABEL_NAME As String = ""

    Sub New(ByVal g As GunEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_TYPE = "C"
        Me.MENU_ITEM_OBJECT = "WHCRF010"

        AppStates.Add("SCAN_UPC", "Scan UPC or Enter Style |PRINT|EXIT|") ' BLUE
        AppStates.Add("SCAN_PRNTR", "Scan Printer |CANCEL|EXIT|")
        AppStates.Add("SCAN_COLOR", "Select a Color from List |CANCEL|")
        AppStates.Add("SCAN_COUNT", "Enter Number of Labels to Print |CANCEL|")
        AppStates.Add("SCAN_LBLNAME", "Print UPC Labels|LARGE|SMALL|CANCEL|")
        AppStates.Add("VERIFY", "Update Label request (Y/N)|Y|N|CANCEL|")

        AppState = "SCAN_UPC"
        LAST_CLR = "BLUE"


        With dst
            '  With .Tables.Add("WHTSCANS").Columns
            ' .Add("BAR_CODE")
            ' .Add("NEW")
            ' .Add("SCANNED")
            ' End With
            ' .Tables("WHTSCANS").PrimaryKey = New DataColumn() {.Tables("WHTSCANS").Columns("BAR_CODE")}

            'Create_TDA(.Tables.Add, "WHTUPCL1 ", "*")
            Create_TDA(.Tables.Add, "WHTUPCL1", "*")

        End With

        tbl = dst.Tables("WHTUPCL1") ' New DataTable

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
                    If SCANTEXT = "PRINT" Then
                        CreateResponse("SCAN_PRNTR", "R", "Print Requested")
                    Else
                        SCANTEXT = Trim(SCANTEXT)
                        Dim CheckResponse As Dictionary(Of String, String) = TACMAIN1.CheckUPC(Me, SCANTEXT)
                        If CheckResponse.ContainsKey("Error") Then
                            CreateResponse("", "R", CheckResponse("Error"))
                            Exit Select
                        End If

                        If CheckResponse.ContainsKey("UPC_CODE") Then
                            UPC_CODE = CheckResponse("UPC_CODE")
                            STYLE_CODE = CheckResponse("STYLE_CODE")
                            COLOR_CODE = CheckResponse("COLOR_CODE")
                            CARTONS_PER_UNIT = Val(CheckResponse("CARTONS_PER_UNIT"))
                            'CreateResponse("SCAN_COUNT", "G", "UPC " & UPC_CODE & " has been selected")
                            CreateResponse("SCAN_LBLNAME", "G", "UPC " & UPC_CODE & " has been selected")
                            Exit Select
                        End If
                        STYLE_CODE = SCANTEXT
                        TACMAIN1.GetColors(Me, STYLE_CODE, "", COLOR_CODEs, colors)

                        CreateResponse("SCAN_COLOR", "G", "Style " & STYLE_CODE & " has been selected, colors " & colors)
                    End If

                Case "SCAN_PRNTR"
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Print Cancelled, Re-scan  UPC")
                        Exit Select
                    Else
                        CreateResponse("SCAN_UPC", "BLUE", "Print sent to " & SCANTEXT)
                    End If

                Case "SCAN_COLOR"
                    Dim colors As String = ""
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Request Cancelled, Re-scan  UPC")
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
                            CARTONS_PER_UNIT = Val(CheckResponse("CARTONS_PER_UNIT"))
                        End If
                    End If

                    'CreateResponse("SCAN_COUNT", "G", "UPC " & UPC_CODE & " has been selected")
                    CreateResponse("SCAN_LBLNAME", "G", "UPC " & UPC_CODE & " has been selected")

                Case "SCAN_LBLNAME"
                    LABEL_NAME = ""
                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Request Cancelled, Re-scan  UPC")
                        Exit Select
                    Else
                        If SCANTEXT = "SMALL" Then
                            LABEL_NAME = "NEWER|smallupc.lbx|"
                        End If
                    End If
                    CreateResponse("SCAN_COUNT", "G", "UPC " & UPC_CODE & " has been selected")

                Case "SCAN_COUNT"

                    If SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "R", "Request Cancelled")
                        Exit Select
                    Else
                        If Val(SCANTEXT) > 99 Or Val(SCANTEXT) < 0 Then
                            CreateResponse("", "R", "Invalid Number Of tickets " & SCANTEXT)
                            Exit Select
                        End If
                        Request_count = Val(SCANTEXT)
                        Dim upc_suffix As Integer = 0

                        Do
                            upc_suffix += 1
                            LBL_REQ_NO = ASCMAIN1.Next_Control_No("WHTUPCL1.LBL_REQ_NO")

                            Dim row As DataRow = tbl.NewRow
                            row.Item("LBL_REQ_NO") = LBL_REQ_NO
                            If CARTONS_PER_UNIT > 1 Then
                                row.Item("UPC_CODE") = String.Format("{0}-{1}", UPC_CODE, upc_suffix)
                            Else
                                row.Item("UPC_CODE") = UPC_CODE
                            End If
                            row.Item("STYLE_CODE") = STYLE_CODE
                            row.Item("COLOR_CODE") = COLOR_CODE
                            row.Item("LBL_QTY") = Request_count
                            row.Item("LOCATION_CODE") = G.GUN_LOC
                            row.Item("INIT_OPER") = G.USER_ID
                            row.Item("INIT_DATE") = DATETIME_STAMP
                            row.Item("LAST_DATE") = DATETIME_STAMP
                            row.Item("LAST_OPER") = G.USER_ID
                            row.Item("PROCESS_IND") = "0"
                            row.Item("LABEL_NAME") = LABEL_NAME & ""
                            tbl.Rows.Add(row)

                            'upc_suffix += 1
                        Loop While CARTONS_PER_UNIT > upc_suffix

                        CreateResponse("VERIFY", "B", "Request " & SCANTEXT & " Labels")
                        'Exit Select
                    End If

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("SCAN_UPC", "BLUE", "Request Created, Scan UPC or EXIT")
                    ElseIf SCANTEXT = "N" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Request Cancelled, Re - scan  UPC")
                    ElseIf SCANTEXT = "CANCEL" Then
                        CreateResponse("SCAN_UPC", "BLUE", "Request Cancelled, Re - scan  UPC")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub

    Sub Update_Record()

        BeginTrans()

        Update_Record_TDA("WHTUPCL1")

        CommitTrans()
    End Sub


End Class
