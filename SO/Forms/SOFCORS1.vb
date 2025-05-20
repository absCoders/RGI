Public Class SOFCORS1
    Public ORDR_GROUP_NOs As New List(Of String)
    Public CUST_CODE As String
    Dim rowSOTORDR1 As DataRow
    Dim ORDR_GROUP_NO As String
    Dim SheetName As New Dictionary(Of String, Integer)
    Dim CUST_NAME As String = ""

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub SOFCORS1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst

            ASCMAIN1.sql = "Select ARTCUSTD.* " _
            & " from ARTCUSTD " _
            & " where ARTCUSTD.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V", 2)
            dst.Tables("ARTCUSTD").Columns.Add("SELECTED")


            ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)

            ASCMAIN1.sql = "Select SOTORDR2.*, ICTCOLR1.COLOR_DESC, ICTSTYL1.CASE_CUBE, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_STATUS, ICTSTYL1.STYLE_ASST_QTY, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
            & ", ICTSTDQ3.DATE_1, ICTSTDQ3.QTY_1, ICTSTDQ3.DATE_2, ICTSTDQ3.QTY_2, ICTSTDQ3.DATE_3, ICTSTDQ3.QTY_3, ICTSTDQ3.DATE_4, ICTSTDQ3.QTY_4" & vbCrLf _
            & " from SOTORDR2,ICTCOLR1,ICTSTYL1,ICTSTYC1,ICTSTDQ3" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = :PARM1" & vbCrLf _
            & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
            & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
            & "   and ICTSTYC1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
            & "   and ICTSTYC1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
            & "   and ICTSTDQ3.ORDR_GROUP_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ICTSTDQ3.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
            & "   and ICTSTDQ3.COLOR_CODE (+) = SOTORDR2.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2)

            Create_TDA(.Tables.Add, "SOTORDR4", "*", 1)


        End With

        grdSOTCORS1.DataSource = dst.Tables("ARTCUSTD")
        Fill_Records("ARTCUSTD", CUST_CODE)

        With grdSOTCORS1.DisplayLayout.Bands(0)
            .Override.AllowUpdate = DefaultableBoolean.True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                End If
            Next
        End With


        grdSOTCORS1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single


    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "STYLE_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "STYLE_CODE"
        End Select
    End Sub
#End Region

    Private Sub cmdSend_Click(sender As System.Object, e As System.EventArgs) Handles cmdSend.Click

        ' 1 Create Sreadsheet
        ' 2 Send to Email include info Selected in Grid
        ' 3 ** write to events file

        EMsg = ""

        If txtINTERNAL_MESSAGE.Text = "" Then
            EMsg &= vbCr & "Internal Message for Customer Order Status File/EMail is Mandatory"
        End If
        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Send Email")
            Exit Sub
        End If
        If grdSOTCORS1.Selected.Rows.Count > 1 Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "You May not Select more then one contact in the body of the email")
            Exit Sub
        End If

        Dim MultiORD As Boolean = False
        SheetName.Clear()

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            rowSOTORDR1 = Fill_Record("SOTORDR1", ORDR_GROUP_NO)
            Fill_Records("SOTORDR2", ORDR_GROUP_NO)
            Fill_Records("SOTORDR4", ORDR_GROUP_NO)
            Create_New_Excel(MultiORD, ORDR_GROUP_NO)
            MultiORD = True
            TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_GROUP_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "CORSTA", "Email:" & txtINTERNAL_MESSAGE.Text)

            Dim NOTES_LNO As Int64 = Val(dst.Tables("SOTORDR4").Compute("MAX(ORDR_CLNO)", "") & "") + 1
            Dim rowSOTORDR4 As DataRow = dst.Tables("SOTORDR4").NewRow
            rowSOTORDR4.Item("ORDR_NO") = ORDR_GROUP_NO
            rowSOTORDR4.Item("ORDR_CLNO") = NOTES_LNO
            rowSOTORDR4.Item("ORDR_COMMENT") = "Email:" & txtINTERNAL_MESSAGE.Text
            rowSOTORDR4.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTORDR4.Item("INIT_DATE") = DATETIME_STAMP
            dst.Tables("SOTORDR4").Rows.Add(rowSOTORDR4)
            Update_Record_TDA("SOTORDR4")

        Next

        Dim filename As String = "Customer " & rowSOTORDR1.Item("CUST_CODE") & " " & CUST_NAME & ".xlsx"

        '2) only allow one checked ARTCUSTD RECORD

        Dim CONTACT_NAME As String = ""
        Dim CONTACT_EMAIL As String = ""

        For Each grow As UltraWinGrid.UltraGridRow In grdSOTCORS1.Selected.Rows
            CONTACT_NAME = "Sales Rep " & grow.Cells("CONTACT_NAME").Value & ""
            CONTACT_EMAIL = grow.Cells("CONTACT_EMAIL").Value & ""
        Next

        Dim ATTACHMENTs As New Dictionary(Of String, String)

        ATTACHMENTs.Add(filename, ASCMAIN1.Folders("Temp") & filename)

        Dim SUBJECT As String = "Customer " & rowSOTORDR1.Item("CUST_CODE") & " " & CUST_NAME
        Dim PFX As String = ""

        Dim SEND_CC_to_USER_ID As Boolean = True
        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

        If ASCMAIN1.Running_in_VS Then
            '     EMAIL_ADDRESSs.Add("dgj@absolution.com", "Darrin Joscelyn")
        End If
        EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
        SEND_CC_to_USER_ID = False
        Dim EMAIL_BODY As String = "Dear " & ASCMAIN1.USER_NAME & vbCrLf & vbCrLf _
                                                       & ASCMAIN1.USER_EMAIL & vbCrLf & vbCrLf _
                                                       & vbCrLf & vbCrLf _
                                                       & CONTACT_NAME & vbCrLf & vbCrLf _
                                                       & CONTACT_EMAIL & vbCrLf & vbCrLf

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, "CUSTORDSTA", True, SEND_CC_to_USER_ID, "", "", "    Order Number", EMAIL_BODY)
        ' DGJ ADD NEW KEY
        If SEND_NO <> "" Then
            MsgBox("email has been sent", MsgBoxStyle.OkOnly, "Verification")
        End If

        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub


    Sub Create_New_Excel(MULTI As Boolean, ORDER As String)

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet
        CUST_NAME = rowSOTORDR1.Item("CUST_NAME") & ""
        CUST_NAME = CUST_NAME.Replace("'", "")
        CUST_NAME = CUST_NAME.Replace("/", " ")
        CUST_NAME = CUST_NAME.Replace("\", " ")
        CUST_NAME = CUST_NAME.Replace(":", " ")
        CUST_NAME = CUST_NAME.Replace("-", " ")
        CUST_NAME = CUST_NAME.Replace(".", "")
        CUST_NAME = CUST_NAME.Replace("&", "")
        CUST_NAME = CUST_NAME.Replace("$", "")
        CUST_NAME = CUST_NAME.Replace("@", "")
        CUST_NAME = CUST_NAME.Replace("!", "")
        CUST_NAME = CUST_NAME.Replace("*", "")
        CUST_NAME = CUST_NAME.Replace("(", "")
        CUST_NAME = CUST_NAME.Replace(")", "")
        CUST_NAME = CUST_NAME.Replace("#", "")

        ' Dim worksheet As SpreadsheetGear.IWorksheet = _workbook.Worksheets("Samples")
        'If MULTI <> True Then
        'End If
        'Dim xls_filename As String = ASCMAIN1.Folders("Temp") & "Customer " & rowSOTORDR1.Item("CUST_CODE") & " " & CUST_NAME & ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & "_XLS") & ".xlsx"
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & "Customer " & rowSOTORDR1.Item("CUST_CODE") & " " & CUST_NAME & ".xlsx"
        If MULTI = True Then
            oWB = SpreadsheetGear.Factory.GetWorkbook(xls_filename)
            oSheet = oWB.Worksheets.Add()
        Else
            oWB = SpreadsheetGear.Factory.GetWorkbook()
            oSheet = oWB.Worksheets(0)
        End If

        Dim SHEET_NAME As String = ""

        If rowSOTORDR1.Item("ORDR_CUST_PO") & "" = "" Then
            SHEET_NAME = "ORD# " & rowSOTORDR1.Item("ORDR_NO")
        Else
            SHEET_NAME = "ORD# " & rowSOTORDR1.Item("ORDR_NO") & "-" & "PO# " & rowSOTORDR1.Item("ORDR_CUST_PO")
        End If

        SHEET_NAME = SHEET_NAME.Replace("'", "")
        SHEET_NAME = SHEET_NAME.Replace("/", " ")
        SHEET_NAME = SHEET_NAME.Replace("\", " ")
        SHEET_NAME = SHEET_NAME.Replace(":", " ")
        '   SHEET_NAME = SHEET_NAME.Replace("-", " ")
        SHEET_NAME = SHEET_NAME.Replace(".", "")
        SHEET_NAME = SHEET_NAME.Replace("&", "")
        SHEET_NAME = SHEET_NAME.Replace("$", "")
        SHEET_NAME = SHEET_NAME.Replace("@", "")
        SHEET_NAME = SHEET_NAME.Replace("!", "")
        SHEET_NAME = SHEET_NAME.Replace("*", "")
        SHEET_NAME = SHEET_NAME.Replace("(", "")
        SHEET_NAME = SHEET_NAME.Replace(")", "")
        ' SHEET_NAME = SHEET_NAME.Replace("#", "")

        If Len(SHEET_NAME) > 30 Then
            SHEET_NAME = Mid(SHEET_NAME, 1, 30)
        End If



        Dim SHEETADD As Integer = 0
        If SheetName.ContainsKey(SHEET_NAME) Then
            SHEETADD = SheetName(SHEET_NAME) + 1
            SheetName(SHEET_NAME) = SHEETADD
            SHEET_NAME = SHEET_NAME & "(" & SHEETADD & ")"
        Else
            SheetName.Add(SHEET_NAME, 1)
        End If



        oSheet.Name = SHEET_NAME

        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1:Z999")

        Dim RX As Int32 = 2
        Dim CX As Int32 = 0

        oSheet.Cells.Font.Name = "Verdana"
        oSheet.Cells.Font.Size = 10
        '    oSheet.Cells.Columns.AutoFit()

        oSheet.Cells("B2").NumberFormat = "@"
        oSheet.Cells("B2").Font.Size = 11
        oSheet.Cells("B2").Value = rowSOTORDR1.Item("CUST_CODE") & "-" & rowSOTORDR1.Item("CUST_NAME")
        oSheet.Cells("B3").NumberFormat = "@"
        oSheet.Cells("B3").Font.Size = 11
        oSheet.Cells("B3").Value = "Regency Order No " & rowSOTORDR1.Item("ORDR_NO")
        oSheet.Cells("B4").NumberFormat = "@"
        oSheet.Cells("B4").Font.Size = 11
        oSheet.Cells("B4").Value = "Customer PO " & rowSOTORDR1.Item("ORDR_CUST_PO")

        RX = 5
        CX = -1
        Dim NUMBER_COLUMNS As Integer = 15
        If chkImages.Checked Then
            CX += 1 : oSheet.Cells(RX, CX).Value = "Image"
            NUMBER_COLUMNS = 16
        End If
        CX += 1 : oSheet.Cells(RX, CX).Value = "Ln"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Style"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Description"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Color"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Ordered"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Picked"
        CX += 1 : oSheet.Cells(RX, CX).Value = "$Picked"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Shipped"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Cancel"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Open"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Avail"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Price"
        CX += 1 : oSheet.Cells(RX, CX).Value = "$Open"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Availability"
        CX += 1 : oSheet.Cells(RX, CX).Value = "Legend "
        CX += 1 : oSheet.Cells(RX, CX).Value = "Status"


        For CX = 0 To NUMBER_COLUMNS
            oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Blue
            oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.White
            oSheet.Cells(RX, CX).Font.Bold = True
        Next
        oSheet.Cells(RX + 1, 0).Activate()
        oSheet.WindowInfo.FreezePanes = True


        RX = 5
        If chkImages.Checked Then
            RX = 6
        End If
        '  CX = 0

        Dim TOT_CT As Integer = 0
        Dim TOT_ORD As Double = 0
        Dim TOT_PICK As Double = 0
        Dim TOT_PICK_SLS As Double = 0
        Dim TOT_SHP As Double = 0
        Dim TOT_CAN As Double = 0
        Dim TOT_OPN As Double = 0
        Dim TOT_SLS As Double = 0
        Dim TOT_AVAIL As Double = 0

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "ORDR_LNO")
            Dim LEGEND As String = ""
            Dim AVAILABILITY As String = ""
            Dim MULTI_Q As Integer = 0
            Dim TOT_AVAIL_QTY As Double = 0
            If Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "") = 0 And Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "") = 0 Then
            Else
                If Val(rowSOTORDR2.Item("QTY_1") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    AVAILABILITY = rowSOTORDR2.Item("QTY_1") & " Now"
                    TOT_AVAIL_QTY = TOT_AVAIL_QTY + Val(rowSOTORDR2.Item("QTY_1") & "")
                    LEGEND = "In Stock"
                End If
                If Val(rowSOTORDR2.Item("QTY_2") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    If AVAILABILITY <> "" Then AVAILABILITY = AVAILABILITY & "; "
                    AVAILABILITY = AVAILABILITY & rowSOTORDR2.Item("QTY_2") & " @ " & Format(rowSOTORDR2.Item("DATE_2"), "MM/dd/yy") '  Format(rowARTCCTR2.Item("EXPIRATION_DATE"), "MMyy")
                    '   TOT_AVAIL_QTY = TOT_AVAIL_QTY + Val(rowSOTORDR2.Item("QTY_2") & "")
                End If
                If Val(rowSOTORDR2.Item("QTY_3") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    If AVAILABILITY <> "" Then AVAILABILITY = AVAILABILITY & "; "
                    AVAILABILITY = AVAILABILITY & rowSOTORDR2.Item("QTY_3") & " @ " & Format(rowSOTORDR2.Item("DATE_3"), "MM/dd/yy")
                    '  TOT_AVAIL_QTY = TOT_AVAIL_QTY + Val(rowSOTORDR2.Item("QTY_3") & "")
                End If
                If Val(rowSOTORDR2.Item("QTY_4") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    If AVAILABILITY <> "" Then AVAILABILITY = AVAILABILITY & "; "
                    AVAILABILITY = AVAILABILITY & rowSOTORDR2.Item("QTY_4") & " @ " & Format(rowSOTORDR2.Item("DATE_4"), "MM/dd/yy")
                    ' TOT_AVAIL_QTY = TOT_AVAIL_QTY + Val(rowSOTORDR2.Item("QTY_4") & "")
                End If
                If MULTI_Q > 1 Then
                    LEGEND = "ETA-Split"
                Else
                    If LEGEND = "" Then
                        LEGEND = "ETA"
                    End If
                End If


                RX += 1
                CX = -1
                If chkImages.Checked Then
                    CX = 0
                    Dim PictureFileName As String = GetImageLocation(rowSOTORDR2.Item("STYLE_CODE"), rowSOTORDR2.Item("COLOR_CODE"))
                    If PictureFileName.Length > 0 Then
                        Add_Image_to_Worksheet(oSheet, PictureFileName, CX, RX)
                        range.Cells(RX, CX).ColumnWidth = 8
                        '              CX += 3
                        ' RX += 3
                    Else
                        oSheet.Cells(RX, CX).Value = "Missing"
                    End If
                End If
                Dim STATUS As String = ""
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE, True)
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE, True)
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_LNO")
                oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Beige
                range.Cells(RX, CX).ColumnWidth = 4
                TOT_CT = TOT_CT + 1
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_CODE")
                range.Cells(RX, CX).ColumnWidth = 11
                Dim STYLE_STATUS As String = rowSOTORDR2.Item("STYLE_STATUS") & ""
                If STYLE_STATUS = "D" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Red
                    STATUS = "Discontinued"
                End If
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_DESC")
                range.Cells(RX, CX).ColumnWidth = 54
                oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Beige
                CX += 1 : oSheet.Cells(RX, CX).Value = "'" & rowSOTORDR2.Item("COLOR_CODE")
                range.Cells(RX, CX).ColumnWidth = 7
                Dim STYLE_COLOR_STATUS As String = rowSOTORDR2.Item("STYLE_COLOR_STATUS") & ""
                If STYLE_COLOR_STATUS = "D" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Red
                    STATUS = "Discontinued"
                End If
                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                range.Cells(RX, CX).ColumnWidth = 9
                TOT_ORD = TOT_ORD + Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                range.Cells(RX, CX).ColumnWidth = 8
                TOT_PICK = TOT_PICK + Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                If Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "") <> 0 Then
                    STATUS = "In Pick"
                End If
                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "") * Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                range.Cells(RX, CX).NumberFormat = “###,###.00”
                range.Cells(RX, CX).ColumnWidth = 10
                TOT_PICK_SLS = TOT_PICK_SLS + Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "") * Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")

                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & "")
                range.Cells(RX, CX).ColumnWidth = 9
                TOT_SHP = Val(TOT_SHP) + Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & "")
                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")
                range.Cells(RX, CX).ColumnWidth = 7
                TOT_CAN = TOT_CAN + Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_QTY_OPEN")
                range.Cells(RX, CX).ColumnWidth = 9
                TOT_OPN = TOT_OPN + Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_RELEASE_AVAIL")
                range.Cells(RX, CX).ColumnWidth = 13
                range.Cells(RX, CX).Columns.Hidden = True
                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                range.Cells(RX, CX).ColumnWidth = 9
                range.Cells(RX, CX).NumberFormat = “####.00”
                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "") * Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                range.Cells(RX, CX).NumberFormat = “###,###.00”
                range.Cells(RX, CX).ColumnWidth = 12
                TOT_SLS = TOT_SLS + Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "") * Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                TOT_AVAIL = TOT_AVAIL + Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "") * TOT_AVAIL_QTY
                CX += 1 : oSheet.Cells(RX, CX).Value = AVAILABILITY
                range.Cells(RX, CX).ColumnWidth = 24
                range.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Center
                If MULTI_Q = 1 And LEGEND = "In Stock" Then
                    oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.LightGreen
                Else
                    oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Yellow
                End If

                CX += 1 : oSheet.Cells(RX, CX).Value = LEGEND
                range.Cells(RX, CX).ColumnWidth = 9.43
                CX += 1 : oSheet.Cells(RX, CX).Value = STATUS
                range.Cells(RX, CX).ColumnWidth = 12
                range.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Center
                If STATUS = "Discontinued" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Red
                End If
                If STATUS = "In Pick" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Blue
                End If
                CX += 1
                If chkImages.Checked Then
                    RX += 3
                End If
            End If

        Next
        RX += 1 : oSheet.Cells(RX, 0).Value = "Totals"
        For CX = 0 To NUMBER_COLUMNS
            oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.LightGray
        Next

        If chkImages.Checked Then
            CX = 1
        Else
            CX = 0
        End If
        RX += 1 : oSheet.Cells(RX, 0 + CX).Value = TOT_CT
        oSheet.Cells(RX, 4 + CX).Value = TOT_ORD
        oSheet.Cells(RX, 5 + CX).Value = TOT_PICK
        oSheet.Cells(RX, 6 + CX).Value = TOT_PICK_SLS
        oSheet.Cells(RX, 7 + CX).Value = TOT_SHP
        oSheet.Cells(RX, 8 + CX).Value = TOT_CAN
        oSheet.Cells(RX, 9 + CX).Value = TOT_OPN
        oSheet.Cells(RX, 12 + CX).Value = TOT_SLS
        range.Cells(RX, 12 + CX).NumberFormat = “###,###.00”
        oSheet.Cells(RX, 13 + CX).Value = "$Tot Avail " & Format(TOT_AVAIL, “###,###.00”)
        range.Cells(RX, 13 + CX).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        ' range.Cells(RX, 13 + CX).NumberFormat = “###,###.00”

        '2025 Tariff Notice
        RX += 3
        oSheet.Range($"A{RX}:N{RX}").Merge()
        'oSheet.Range($"A{RX}:N{RX}").Value = "Tariffs may come into effect on all imported items; if the US Government imposes them, a surcharge will be added to the bottom of the invoice."
        oSheet.Range($"A{RX}:N{RX}").Value = "Effective Immediately: A temporary 18% surcharge now applies to all warehouse shipments; future adjustments may occur."
        oSheet.Range($"A{RX}:N{RX}").Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
        oSheet.Range($"A{RX}:N{RX}").Borders.Weight = SpreadsheetGear.BorderWeight.Thin
        oSheet.Range($"A{RX}:N{RX}").Font.Bold = True
        oSheet.Range($"A{RX}:N{RX}").Font.Color = SpreadsheetGear.Colors.Red

        Dim SFX As String = ASCMAIN1.Next_Control_No("ExportDocuments")
        Dim XLS_FILE As String = Replace(xls_filename, "ExportDocuments", "ExportDocuments" & "_" & SFX)
        oWB.SaveAs(XLS_FILE, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        oWB.Close()
        range = Nothing
        oSheet = Nothing
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub
    Private Function GetImageLocation(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = ""
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim RO_PARM_STYLE_IMG_DIR As String = ""
        Dim FileMatch As String
        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
        Dim COLOR_CODE_LONG As String = ""
        If Not IsNothing(rowICTCOLR1) Then
            COLOR_CODE_LONG = rowICTCOLR1.Item("COLOR_CODE_LONG").ToString()
        End If

        If Not IsNothing(rowSOTPARM3) Then
            RO_PARM_STYLE_IMG_DIR = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            If RO_PARM_STYLE_IMG_DIR.Length > 0 Then
                TryPullWebImage(STYLE_CODE, COLOR_CODE)

                FileMatch = Dir(String.Format("{0}\{1}-{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                If FileMatch.Length > 0 Then
                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                Else
                    FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                    If FileMatch.Length > 0 Then
                        RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                    Else
                        FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE_LONG))
                        If FileMatch.Length > 0 Then
                            RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                        Else
                            FileMatch = Dir(String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                            If FileMatch.Length > 0 Then
                                RetVal = String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE)
                            Else
                                FileMatch = Dir(String.Format("{0}\{1}*", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                                If FileMatch.Length > 0 Then
                                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return RetVal
    End Function

    Private Sub TryPullWebImage(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String)
        Try
            Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
            Dim RO_PARM_STYLE_IMG_DIR As String = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            If RO_PARM_STYLE_IMG_DIR.Length > 0 Then
                Dim WEBURL As String = "https://www.regency-rib.com/media/product/"
                Dim FILEURL As String = WEBURL & STYLE_CODE & "-" & COLOR_CODE & ".jpg"
                If Not RO_PARM_STYLE_IMG_DIR.EndsWith("\") Then
                    RO_PARM_STYLE_IMG_DIR = RO_PARM_STYLE_IMG_DIR & "\"
                End If
                Dim TMP_FILE As String = RO_PARM_STYLE_IMG_DIR & STYLE_CODE & "-" & COLOR_CODE & ".jpg"
                If IO.Directory.Exists(RO_PARM_STYLE_IMG_DIR) Then
                    Dim web_client As New Net.WebClient
                    Dim image_stream As New IO.MemoryStream(web_client.DownloadData(FILEURL))
                    Dim img As Drawing.Image = Drawing.Image.FromStream(image_stream)
                    If IO.File.Exists(TMP_FILE) Then
                        IO.File.Delete(TMP_FILE)
                    End If
                    img.Save(TMP_FILE)
                    'FILENAME = TMP_FILE
                End If
            End If
        Catch ex As Exception
            'Just Bail
        End Try
    End Sub

    Sub Add_Image_to_Worksheet(worksheet As SpreadsheetGear.IWorksheet, ITEM_CODE As String, cx As Integer, rx As Integer)

        'Dim IMAGE_NAME As String = ITEM_CODE
        'Dim IMAGE_FOLDER As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        'If ASCMAIN1.Running_in_VS Then
        '    IMAGE_FOLDER = "C:\Share\INT\Pictures\"
        'End If
        'Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME & ".jpg"
        'If Not My.Computer.FileSystem.FileExists(imageFileStyle) Then
        '    imageFileStyle = IMAGE_FOLDER & "\" & IMAGE_NAME & ".PNG"
        'End If
        Dim imageFileStyle As String = ITEM_CODE

        Dim imageStyle As System.Drawing.Image = Nothing
        If My.Computer.FileSystem.FileExists(imageFileStyle) Then
            Dim widthStyle As Double
            Dim heightStyle As Double

            imageStyle = System.Drawing.Image.FromFile(imageFileStyle)
            Try
                widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution * 0.25
                heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution * 0.25
            Finally
                imageStyle.Dispose()
            End Try

            ' NEW 
            Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

            Dim col_adj As Decimal = 0
            If heightStyle > 45 Then
                heightStyle = 42
                widthStyle = 42
            End If
            If heightStyle > widthStyle Then
                col_adj = 0.3
            Else
                col_adj = 0.05
            End If

            Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(cx) + col_adj
            Dim topStyle As Double = windowInfoStyle.RowToPoints(rx - 1) + 0.1 ' 1.5)

            ' ImageRows = windowInfoStyle.PointsToRow(heightStyle)
            worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)

            ' END NEW 


            '' Calculate the left and top placement of the picture by converting 
            '' row and column coordinates to points.  Use fractional values to 
            '' get coordinates anywhere in between row and column boundaries.
            'Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo
            'Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(cx)
            'Dim topStyle As Double = windowInfoStyle.RowToPoints(rx)

            '' Add the picture from file.
            'worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
        End If

    End Sub

    Private Sub SplitContainer1_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub


End Class