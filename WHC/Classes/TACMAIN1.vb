Public Class TACMAIN1

    ' nSoftware License Keys
    Public nSoftwareZipkey As String = "315A4E384141315355425241533154453345383933333331580000000000000000000000000000003532323931555536000042424D454B375544463730460000"
    Public nSoftwareftpkey As String = "31504E384141315355425241533154453345383933333331580000000000000000000000000000003532323931555536000046384445474A3131583750500000"
    Public nSoftwareipportkey As String = "31504E38414131535542524153315445334538393333333158000000000000000000000000000000423252383654315A0000343854353048333650384B370000"
    Public nSoftwarepopkey As String = "31504E38414131535542524153315445334538393333333158000000000000000000000000000000394A37383437343200004E55523837363650504A42350000"
    Public nSoftwarehttpkey As String = "31504E38414131535542524153315445334538393333333158000000000000000000000000000000364E5333445A425800005555543150444D5443334A430000"
    'Public nSoftwareInship As String = "42584E354141315355425241533154453345383933333331580000000000000000000000000000004A52344B5057583900003059573859305A4A545958520000"
    Public nSoftwareInship As String = "42584E3541413153554252415331544533453839333333315800000000000000000000000000000059585246324D544600003030584254504E57374345330000"
    Public nSoftwaresftpkey As String = "31484E3841413153554252415331544533453839333333315800000000000000000000000000000039504E373857464800004A315A5A4D5A3038384259430000"

    Public Overridable Sub Site_Specific_Settings()

    End Sub

    Public Overridable Sub Get_Column_Expression_Exceptions(ByVal FORM_NAME As String, ByVal DATA_SOURCE As String, ByVal COLUMN_NAME As String, ByRef sql_SELECT_col As String) ' , ByRef sql_GROUP_BY_col As String)

    End Sub

    Public Overridable Function Get_Code_SQL_X(ByVal FORM_NAME As String, ByVal COLUMN_NAME As String, ByRef GROUP_KEY As String) As String
        Return Nothing
    End Function

    Public Overridable Sub Write_Group_Record_X(ByVal GROUP_KEY As String, ByVal COLUMN_NAME As String, ByVal GROUP_CODEs As ArrayList, ByVal GROUP_DESCs As ArrayList)

    End Sub

    Public Overridable Function CodeValues(ByVal TABLE_COLUMN As String) As Dictionary(Of String, String)
        Return Nothing
    End Function

    Public Overridable Function Send_email(ByVal frmASFBASE0 As ASCBASE0, _
                                 ByVal EMAIL_ADDRESSs As Dictionary(Of String, String), _
                                 ByVal ATTACHMENTs As Dictionary(Of String, String), _
                                 ByVal SUBJECT As String, _
                                 ByVal EMAIL_KEY As String, _
                                 Optional ByVal auto_send As Boolean = False, _
                                 Optional SEND_CC_to_USER_ID As Boolean = False, _
                                 Optional ENTITY_KEY As String = "", _
                                 Optional ENTITY_NAME As String = "", _
                                 Optional ENTITY_CAPTION As String = "") As String
        Return Nothing
    End Function

    Public Overridable Sub Application_Initialization()

    End Sub

    Public Shared Function UPC(
        ByRef clsWHCRF000 As WHCRF000,
        ByVal UPC_SEQUENCE_NO As String,
        ByRef SO_PARM_UPC_VENDOR_ID As String,
        Optional ByVal prefix_with_VENDOR_ID As Boolean = True) As String

        ' Note: Check Digit Calculation applies to the 19-digits prior to the check digit
        '       These 11 digits are made up from the 6 digit Vendor ID prepended to the 5 digit UPC Serial Number
        '       19 digits = '0000' + 6 digit SO_PARM_UPC_VENDOR_ID + 9 digit Carton Serial Number

        Dim Check_Digit_Seed As String

        If prefix_with_VENDOR_ID Then
            If SO_PARM_UPC_VENDOR_ID = "" Then
                SO_PARM_UPC_VENDOR_ID = clsWHCRF000.ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & ""
            End If

            If Len(UPC_SEQUENCE_NO) <> 5 Then
                If Len(UPC_SEQUENCE_NO) <> 9 Then
                    Stop
                End If
            End If

            Check_Digit_Seed = Mid(SO_PARM_UPC_VENDOR_ID, 1) & UPC_SEQUENCE_NO
        Else
            Check_Digit_Seed = UPC_SEQUENCE_NO
        End If

        Dim odd_digits As Integer
        Dim even_digits As Integer

        For i As Integer = 1 To Len(Check_Digit_Seed) Step 2
            odd_digits = odd_digits + Val(Mid(Check_Digit_Seed, 1, 1))
            Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            If Check_Digit_Seed <> "" Then
                even_digits = even_digits + Val(Mid(Check_Digit_Seed, 1, 1))
                Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            End If
        Next i

        Dim check_digit As Integer
        check_digit = (odd_digits * 3 + even_digits) Mod 10
        If check_digit <> 0 Then
            check_digit = 10 - check_digit
        End If

        If prefix_with_VENDOR_ID Then
            UPC = SO_PARM_UPC_VENDOR_ID & UPC_SEQUENCE_NO & Format(check_digit, "0")
        Else
            UPC = UPC_SEQUENCE_NO & Format(check_digit, "0")
        End If

    End Function

    Public Sub Record_Event(
        ByVal TABLE_NAME As String,
        ByVal TABLE_KEY As String,
        ByVal INIT_DATE As Date,
        ByVal INIT_OPER As String,
        ByVal EVENT_TYPE As String,
        ByVal EVENT_DESC As String,
        Optional ByVal EVENT_KEY As String = "",
        Optional FORM_NAME As String = "")

        'If FORM_NAME = "" Then
        '    FORM_NAME = ASCMAIN1.ActiveForm.Name
        'End If

        'ASCDATA1.ExecuteSQL("Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME) " _
        '                     & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8)", _
        '                     "VVDVVVVV", _
        '                     New Object() {TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME})

    End Sub
    Public Shared Function CheckUPC(
        clsWHCRF000 As WHCRF000,
        ByVal SCANTEXT As String) As Dictionary(Of String, String)

        Dim suffix As String = ""
        Dim RtnDict As New Dictionary(Of String, String)
        SCANTEXT = SCANTEXT.ToUpper
        If SCANTEXT.Length > 12 Then
            suffix = SCANTEXT.Substring(12)
            SCANTEXT = SCANTEXT.Substring(0, 12)
        End If

        clsWHCRF000.ASCMAIN1.sql = "Select ICTSTYC1.*, NVL(ICTCOLR1.COLOR_CODE_LONG, ICTCOLR1.COLOR_ABBR) COLOR_DESC, nvl(ICTSTYL1.CARTONS_PER_UNIT, 0) CARTONS_PER_UNIT" & vbCrLf _
                    & " ,nvl(CARTON_PACK_QTY,1) CARTON_PACK_QTY, nvl(INNER_PACK_QTY,0) INNER_PACK_QTY, STYLE_DESC" & vbCrLf _
                    & " from ICTSTYC1, ICTCOLR1, ICTSTYL1" & vbCrLf _
                    & " where  '" & SCANTEXT & "' in  (ICTSTYC1.STYLE_CODE, ICTSTYC1.UPC_CODE) and ICTSTYC1.COLOR_CODE = ICTCOLR1.COLOR_CODE and ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE"
        Dim cnt As Integer = 0
        Dim rows() As DataRow = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
        If rows.Length > 0 Then
            If rows.Length = 1 Then
                RtnDict.Add("UPC_CODE", rows(0).Item("UPC_CODE"))
                RtnDict.Add("STYLE_CODE", rows(0).Item("STYLE_CODE"))
                RtnDict.Add("STYLE_DESC", rows(0).Item("STYLE_DESC"))
                RtnDict.Add("COLOR_CODE", rows(0).Item("COLOR_CODE"))
                RtnDict.Add("COLOR_DESC", rows(0).Item("COLOR_DESC") & "")
                RtnDict.Add("CARTONS_PER_UNIT", rows(0).Item("CARTONS_PER_UNIT"))
                RtnDict.Add("CARTON_PACK_QTY", rows(0).Item("CARTON_PACK_QTY"))
                RtnDict.Add("INNER_PACK_QTY", rows(0).Item("INNER_PACK_QTY"))
                If suffix <> "" Then
                    RtnDict.Add("UPC_SUFFIX", suffix)
                End If
            Else
                If rows(0).Item("STYLE_CODE") <> SCANTEXT Then
                    RtnDict.Add("Error", "Style/UPC '" & SCANTEXT & "' not found, Try again")
                Else
                    RtnDict.Add("STYLE_CODE", rows(0).Item("STYLE_CODE"))
                End If
            End If
        Else
            RtnDict.Add("Error", "Style/UPC '" & SCANTEXT & "' not found, Try again")
        End If

        Return RtnDict
    End Function

    Public Shared Function LookupLocation(
    clsWHCRF000 As WHCRF000,
    ByVal SCANTEXT As String) As String
        Dim StyleList As String = ""
        Dim rc As Int32 = 0
        clsWHCRF000.ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                        & " where WHTLOCB1.LOCATION_CODE = '" & SCANTEXT.ToUpper & "'" & vbCrLf _
                        & " and WHTLOCB1.WHSE_CODE = '" & clsWHCRF000.G.WHSE_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.LOCATION_QTY > 0 "

        Dim rows() As DataRow = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
        If rows.Length > 0 Then
            For Each ROW As DataRow In rows
                StyleList &= vbCrLf & ROW.Item("STYLE_CODE") & ", " & ROW.Item("COLOR_CODE") & " > " & ROW.Item("LAST_OPER") & " # " & ROW.Item("LOCATION_QTY")
                rc = rc + 1
                If rc > 9 Then Exit For
            Next
        End If

        Return StyleList
    End Function


    Public Shared Function CheckLocation(
   clsWHCRF000 As WHCRF000,
   ByVal SCANTEXT As String,
   ByVal LockLocation As Boolean,
   Optional ByVal ignoreLocked As Boolean = True,
   Optional ByVal IgnoreWaveLock As Boolean = True,
   Optional ByVal LocationUse As String = "",
   Optional ByVal WHSE_CODE_IN As String = "") As Dictionary(Of String, String)

        Dim WHSE_CODE As String = WHSE_CODE_IN
        If WHSE_CODE_IN = "" Then
            WHSE_CODE = clsWHCRF000.G.WHSE_CODE
        End If

        Dim RtnDict As New Dictionary(Of String, String)
        SCANTEXT = SCANTEXT.ToUpper

        Dim rowWHTLOCM1 As DataRow = clsWHCRF000.LookUp("WHTLOCM1", New String() {WHSE_CODE, SCANTEXT})
        If rowWHTLOCM1 Is Nothing Then
            RtnDict.Add("Error", "Invalid Location " & SCANTEXT)
            Return RtnDict
            Exit Function
        End If

        If ignoreLocked = False And rowWHTLOCM1.Item("LOCATION_LOCKED") & "" = "1" Then
            RtnDict.Add("Error", "Location " & SCANTEXT & " is Locked")
            Return RtnDict
            Exit Function
        End If

        If String.IsNullOrEmpty(LocationUse) = False And rowWHTLOCM1.Item("LOCATION_USE") & "" <> LocationUse Then
            RtnDict.Add("Error", "Location " & SCANTEXT & " is not Valid for this App")
            Return RtnDict
            Exit Function
        End If

        If LockLocation Then
            If Not clsWHCRF000.ASCMAIN1.Logical_Lock("WHTLOCM1", SCANTEXT) Then
                RtnDict.Add("Error", "Could not lock access to Locaton " & SCANTEXT)
                Return RtnDict
                Exit Function
            End If
        End If

        If IgnoreWaveLock = False Then
            Dim rowsWAVED() As DataRow = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
            clsWHCRF000.ASCMAIN1.sql = "Select * from WHTLOCB1" & vbCrLf _
                        & " where LOCATION_CODE = '" & SCANTEXT & "'" _
                        & " and WHSE_CODE = '" & WHSE_CODE & "'" _
                        & " and LOCATION_QTY_WAVE > 0"
            rowsWAVED = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
            If rowsWAVED.Length <> 0 Then
                RtnDict.Add("Error", "Location " & SCANTEXT & " is already Committed to a Wave")
                Return RtnDict
                Exit Function
            End If

        End If

        Dim SSSS As String = ""
        clsWHCRF000.ASCMAIN1.sql = "Select DISTINCT STYLE_CODE,COUNT(*) DDD from WHTLOCB1" & vbCrLf _
                        & " where WHTLOCB1.LOCATION_CODE = '" & SCANTEXT & "'" & vbCrLf _
                        & " and WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                        & " and WHTLOCB1.LOCATION_QTY > 0 GROUP BY STYLE_CODE" & vbCrLf _
                        & " order by STYLE_CODE"

        Dim rows() As DataRow = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
        If rows.Length > 0 Then
            ' Cases_count = rows5.Length
            Dim cnt As Integer = 0
            For Each ROW As DataRow In rows
                SSSS = SSSS & ", " & ROW.Item("STYLE_CODE")
                cnt += 1
                If cnt > 100 Then Exit For
            Next
        End If
        If Mid(SSSS, 1, 2) = ", " Then
            SSSS = Mid(SSSS, 3)
        End If

        RtnDict.Add("LOCATION_CODE", SCANTEXT)
        RtnDict.Add("Stylelist", SSSS)

        Return RtnDict
    End Function
    Public Shared Sub GetColors(
        clsWHCRF000 As WHCRF000,
        ByVal STYLE_CODE As String,
        ByVal LOCATION_CODE As String,
        ByRef COLOR_CODEs As List(Of String),
        ByRef ColorList As String)

        Dim cnt As Integer = 0
        Dim rows() As DataRow

        STYLE_CODE = STYLE_CODE.ToUpper

        If Not String.IsNullOrEmpty(LOCATION_CODE) Then
            clsWHCRF000.ASCMAIN1.sql = "Select WHTLOCB1.* from WHTLOCB1" & vbCrLf _
                       & " where  WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" _
                       & " and LOCATION_CODE = '" & LOCATION_CODE & "'"
            rows = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
            COLOR_CODEs.Clear()
            If rows.Length > 0 Then
                For Each ROW As DataRow In rows
                    cnt = cnt + 1
                    ColorList = ColorList & String.Format(" {0}-{1}", cnt, ROW.Item("COLOR_CODE"))
                    COLOR_CODEs.Add(ROW.Item("COLOR_CODE"))
                Next
                ColorList = Mid(ColorList, 2)
            End If
        End If
        If cnt = 0 Then
            clsWHCRF000.ASCMAIN1.sql = "Select ICTSTYC1.* from ICTSTYC1" & vbCrLf _
                 & " where  ICTSTYC1.STYLE_CODE = '" & STYLE_CODE & "'"
            rows = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
            COLOR_CODEs.Clear()
            If rows.Length > 0 Then
                For Each ROW As DataRow In rows
                    cnt = cnt + 1
                    ColorList = ColorList & String.Format(" {0}-{1}", cnt, ROW.Item("COLOR_CODE"))
                    COLOR_CODEs.Add(ROW.Item("COLOR_CODE"))
                Next
                ColorList = Mid(ColorList, 2)
            End If
        End If

    End Sub

    Public Shared Function VerifyColor(
        clsWHCRF000 As WHCRF000,
        ByVal SCANTEXT As String,
        ByVal STYLE_CODE As String,
        ByVal COLOR_CODEs As List(Of String)) As Dictionary(Of String, String)

        Dim RtnDict As New Dictionary(Of String, String)
        Dim COLOR_CODE As String

        STYLE_CODE = STYLE_CODE.ToUpper

        If IsNumeric(SCANTEXT) And Val(SCANTEXT) > 0 And Val(SCANTEXT) <= COLOR_CODEs.Count Then
            COLOR_CODE = COLOR_CODEs(Val(SCANTEXT) - 1)
        Else
            COLOR_CODE = SCANTEXT.ToUpper
        End If
        '"Select ICTSTYC1.*, NVL(ICTCOLR1.COLOR_CODE_LONG, ICTCOLR1.COLOR_ABBR) COLOR_DESC from ICTSTYC1, ICTCOLR1" & vbCrLf _
        '               & " where  '" & SCANTEXT & "' in  (ICTSTYC1.STYLE_CODE, ICTSTYC1.UPC_CODE) and ICTSTYC1.COLOR_CODE = ICTCOLR1.COLOR_CODE"
        clsWHCRF000.ASCMAIN1.sql = "Select ICTSTYC1.*, NVL(ICTCOLR1.COLOR_CODE_LONG, ICTCOLR1.COLOR_ABBR) COLOR_DESC, nvl(ICTSTYL1.CARTONS_PER_UNIT, 0) CARTONS_PER_UNIT" & vbCrLf _
            & " ,nvl(CARTON_PACK_QTY,1) CARTON_PACK_QTY, nvl(INNER_PACK_QTY,0) INNER_PACK_QTY, STYLE_DESC" & vbCrLf _
            & " from ICTSTYC1, ICTCOLR1, ICTSTYL1" & vbCrLf _
            & " where  ICTSTYC1.STYLE_CODE = '" & STYLE_CODE & "' and ICTSTYC1.COLOR_CODE = '" & COLOR_CODE & "'  and ICTSTYC1.COLOR_CODE = ICTCOLR1.COLOR_CODE and ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE"
        Dim rows() As DataRow = clsWHCRF000.ASCDATA1.GetDataTable.Select("")
        If rows.Length = 1 Then
            RtnDict.Add("UPC_CODE", rows(0).Item("UPC_CODE"))
            RtnDict.Add("COLOR_CODE", rows(0).Item("COLOR_CODE"))
            RtnDict.Add("COLOR_DESC", rows(0).Item("COLOR_DESC") & "")
            RtnDict.Add("STYLE_DESC", rows(0).Item("STYLE_DESC"))
            RtnDict.Add("CARTONS_PER_UNIT", rows(0).Item("CARTONS_PER_UNIT"))
            RtnDict.Add("CARTON_PACK_QTY", rows(0).Item("CARTON_PACK_QTY"))
            RtnDict.Add("INNER_PACK_QTY", rows(0).Item("INNER_PACK_QTY"))
        Else
            Dim colors As String = ""
            For c As Integer = 1 To COLOR_CODEs.Count
                colors = colors & String.Format(", {0}: {1}", c, COLOR_CODEs(c - 1))
            Next
            RtnDict.Add("Error", "Color '" & SCANTEXT & "' not found for Style '" & STYLE_CODE & "'Try again, colors " & colors)
        End If
        Return RtnDict
    End Function

    Public Shared Function VerifyTransferPallet(
    clsWHCRF000 As WHCRF000,
    ByVal SCANTEXT As String) As Dictionary(Of String, String)

        Dim RtnDict As New Dictionary(Of String, String)
        Dim sqlWHTTRAN1 As String = $"SELECT * FROM WHTTRAN1 WHERE PALLET_NO = '{SCANTEXT}'"
        Dim tblWHTTRAN1 As DataTable = clsWHCRF000.ASCDATA1.GetDataTable(sqlWHTTRAN1)

        If tblWHTTRAN1.Rows.Count > 0 Then
            RtnDict.Add("Error", $"{SCANTEXT} has already been received.")
        Else
            Dim sqlPallet As String = $"SELECT * FROM EDT945T3 WHERE EDI_PALLET_NO = '{SCANTEXT}'"
            Dim tblEDT945T3 As DataTable = clsWHCRF000.ASCDATA1.GetDataTable(sqlPallet)
            Dim PICK_NO As String = SCANTEXT.Substring(0, 10)
            Dim PALLET_SEQ_NO As Integer = CInt(SCANTEXT.Substring(10, 3))
            Dim PICK_NO_USL As String = SCANTEXT.Substring(13)
            Dim sqlSOTPICK1 As String = $"SELECT * FROM SOTPICK1 WHERE PICK_NO = '{PICK_NO}'"
            Dim tblSOTPICK1 As DataTable = clsWHCRF000.ASCDATA1.GetDataTable(sqlSOTPICK1)

            Dim rows() As DataRow = tblEDT945T3.Select("")
            If rows.Length = 1 Then
                RtnDict.Add("EDI_DOC_SEQ_NO", rows(0)("EDI_DOC_SEQ_NO"))
                RtnDict.Add("EDI_STATUS", rows(0)("EDI_STATUS"))
                Dim rowSOTPICK1() As DataRow = tblSOTPICK1.Select("")
                If rowSOTPICK1.Length = 1 Then
                    Dim PICK_STATUS As String = rowSOTPICK1(0)("PICK_STATUS")
                    If PICK_STATUS <> "P" Then
                        RtnDict.Add("Error", $"{PICK_NO} is not in pick (PICK_STATUS <> 'P')")
                    End If
                Else
                    RtnDict.Add("Error", $"{SCANTEXT} is not a valid Pick Ticket")
                End If
            Else
                RtnDict.Add("Error", $"{SCANTEXT} is not a valid EDI Pallet No")
            End If
        End If

        Return RtnDict
    End Function

    Public Shared Function VerifyTransferUPC(
clsWHCRF000 As WHCRF000,
ByVal SCANTEXT As String,
ByVal EDI_DOC_SEQ_NO As String,
ByVal UPCDict As Dictionary(Of String, String)) As Dictionary(Of String, String)

        Dim sqlEDT945T2 As String = $"SELECT * FROM EDT945T2 WHERE EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}' AND UPC_CODE = '{SCANTEXT}'"
        Dim tblEDT945T2 As DataTable = clsWHCRF000.ASCDATA1.GetDataTable(sqlEDT945T2, "EDT945T2", "VV", New Object() {EDI_DOC_SEQ_NO, SCANTEXT})

        Dim rows() As DataRow = tblEDT945T2.Select("")
        If rows.Length > 0 Then
            Dim STYLE_CODE As String = UPCDict("STYLE_CODE")
            Dim COLOR_CODE As String = UPCDict("COLOR_CODE")
            Dim EDI_STYLE As String = rows(0)("EDI_STYLE")
            Dim EDI_COLOR As String = rows(0)("EDI_COLOR")
            If STYLE_CODE = EDI_STYLE AndAlso COLOR_CODE = EDI_COLOR Then
                UPCDict.Add("EDI_STYLE", EDI_STYLE)
                UPCDict.Add("EDI_COLOR", EDI_COLOR)
            Else
                UPCDict.Add("Error", $"Style/Color Mismatch for UPC {SCANTEXT} ")
            End If
        Else
            UPCDict.Add("Error", $"{SCANTEXT} is not a valid UPC for this Pallet")
        End If
        Return UPCDict
    End Function
End Class
