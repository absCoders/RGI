Public Class WHCRF014
    ' Application Show Items on Location

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
        Me.MENU_ITEM_OBJECT = "WHCRF014"

        AppStates.Add("SCAN_LOC", "Scan Location |SHW GUN|EXIT|") ' YELLOW
        'AppStates.Add("VERIFY", "Update (Y/N)|Y|N|CANCEL|")

        AppState = "SCAN_LOC"
        LAST_CLR = "YELLOW"

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
                    LOCATION_CODE = SCANTEXT
                    If SCANTEXT = "SHW GUN" Then
                        LOCATION_CODE = G.GUN_LOC
                        SCANTEXT = G.GUN_LOC
                    End If

                    Dim Styles As String = TACMAIN1.LookupLocation(Me, SCANTEXT)
                    If Styles.Length = 0 Then
                        CreateResponse("", "YELLOW", "No Styles found for location " & SCANTEXT)
                        Exit Select
                    End If
                    CreateResponse("", "YELLOW", "Location " & SCANTEXT & ", Styles" & Styles)

                Case "VERIFY"
                    If SCANTEXT = "Y" Then
                        'Update_Record()
                        ASCMAIN1.MultiTask_Release()
                        CreateResponse("", "B", "Exit")
                    ElseIf SCANTEXT = "N" Then
                        'CreateResponse("SCAN_UPC", "B", "Scan Ignored, Re - Scan UPC")
                        CreateResponse("SCAN_LOC", "YELLOW", "move Cancelled, Re - scan  location")
                    ElseIf SCANTEXT = "CANCEL" Then

                        CreateResponse("SCAN_LOC", "YELLOW", "Move Cancelled, Re - scan  location")
                    Else
                        CreateResponse("", "R", "Invalid Response")
                    End If
            End Select
        End If
    End Sub

End Class
