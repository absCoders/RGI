Public Class SOFCORS1
    Public ORDR_GROUP_NOs As New List(Of String)
    Public CUST_CODE As String
    Dim rowSOTORDR1 As DataRow
    Dim ORDR_GROUP_NO As String
    Dim SheetName As New Dictionary(Of String, Integer)

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





        End With

        grdSOTCORS1.DataSource = dst.Tables("ARTCUSTD")
        ' grdSOTCORS1.s


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



        'Create_Summary(grdSOTCORS1, "SHIP_BOL_NO", "Count")



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
            Create_New_Excel(MultiORD, ORDR_GROUP_NO)
            MultiORD = True
            TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_GROUP_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "CORSTA", "Email:" & txtINTERNAL_MESSAGE.Text)
        Next
        Dim filename As String = "Customer " & rowSOTORDR1.Item("CUST_CODE") & " " & rowSOTORDR1.Item("CUST_NAME") & ".xlsx"
            '2) only allow one checked ARTCUSTD RECORD

            Dim CONTACT_NAME As String = ""
            Dim CONTACT_EMAIL As String = ""


            For Each grow As UltraWinGrid.UltraGridRow In grdSOTCORS1.Selected.Rows
            CONTACT_NAME = "Sales Rep " & grow.Cells("CONTACT_NAME").Value & ""
            CONTACT_EMAIL = grow.Cells("CONTACT_EMAIL").Value & ""

            Next

            Dim ATTACHMENTs As New Dictionary(Of String, String)

        ATTACHMENTs.Add(filename, ASCMAIN1.Folders("Temp") & filename)

        Dim SUBJECT As String = "Customer PO " & ORDR_GROUP_NO & " " & rowSOTORDR1.Item("CUST_NAME")
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

        '      TAC.TACMAIN1.Record_Event("SOTORDR1", ORDR_GROUP_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "CORSTA", txtINTERNAL_MESSAGE.Text)

        ' 3 DIFFERENT SHEETS WITH ORDER NUMBER AS THE SHEET 





        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
    End Sub


    Sub Create_New_Excel(MULTI As Boolean, ORDER As String)
        '   Dim worksheet As SpreadsheetGear.IWorksheet = Nothing

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet

        ' Dim worksheet As SpreadsheetGear.IWorksheet = _workbook.Worksheets("Samples")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & "Customer " & rowSOTORDR1.Item("CUST_CODE") & " " & rowSOTORDR1.Item("CUST_NAME") & ".xlsx"
        If MULTI = True Then
            oWB = SpreadsheetGear.Factory.GetWorkbook(xls_filename)
            oSheet = oWB.Worksheets.Add()
        Else
            oWB = SpreadsheetGear.Factory.GetWorkbook()
            oSheet = oWB.Worksheets(0)

        End If

        Dim SHEET_NAME As String = ""


        If rowSOTORDR1.Item("ORDR_CUST_PO") & "" = "" Then
            SHEET_NAME = "Regency Order " & rowSOTORDR1.Item("ORDR_NO")
        Else
            SHEET_NAME = rowSOTORDR1.Item("ORDR_CUST_PO")
        End If

        '     SHEET_NAME = "WALT"

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
        oSheet.Cells(RX, 0).Value = "Ln"
        oSheet.Cells(RX, 1).Value = "Style"
        oSheet.Cells(RX, 2).Value = "Description"
        oSheet.Cells(RX, 3).Value = "Color"
        oSheet.Cells(RX, 4).Value = "Ordered"
        oSheet.Cells(RX, 5).Value = "Picked"
        oSheet.Cells(RX, 6).Value = "Shipped"
        oSheet.Cells(RX, 7).Value = "Canceled"
        oSheet.Cells(RX, 8).Value = "Open"
        oSheet.Cells(RX, 9).Value = "Avail"
        oSheet.Cells(RX, 10).Value = "Price"
        oSheet.Cells(RX, 11).Value = "$Open"
        oSheet.Cells(RX, 12).Value = "Availability"
        oSheet.Cells(RX, 13).Value = "Legend "
        oSheet.Cells(RX, 14).Value = "Status"

        For CX = 0 To 14
            oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Blue
            oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.White
            oSheet.Cells(RX, CX).Font.Bold = True
        Next

        RX = 5
        CX = 0

        Dim TOT_CT As Integer = 0
        Dim TOT_ORD As Double = 0
        Dim TOT_PICK As Double = 0
        Dim TOT_SHP As Double = 0
        Dim TOT_CAN As Double = 0
        Dim TOT_OPN As Double = 0
        Dim TOT_SLS As Double = 0


        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "ORDR_LNO")
            Dim LEGEND As String = ""
            Dim AVAILABILITY As String = ""
            Dim MULTI_Q As Integer = 0
            If rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0 And rowSOTORDR2.Item("ORDR_QTY_PICK") = 0 Then
            Else
                If Val(rowSOTORDR2.Item("QTY_1") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    AVAILABILITY = rowSOTORDR2.Item("QTY_1") & " Now"
                    LEGEND = "In Stock"
                End If
                If Val(rowSOTORDR2.Item("QTY_2") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    If AVAILABILITY <> "" Then AVAILABILITY = AVAILABILITY & "; "
                    AVAILABILITY = AVAILABILITY & rowSOTORDR2.Item("QTY_2") & " @ " & Format(rowSOTORDR2.Item("DATE_2"), "MM/yy") '  Format(rowARTCCTR2.Item("EXPIRATION_DATE"), "MMyy")
                End If
                If Val(rowSOTORDR2.Item("QTY_3") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    If AVAILABILITY <> "" Then AVAILABILITY = AVAILABILITY & "; "
                    AVAILABILITY = AVAILABILITY & rowSOTORDR2.Item("QTY_3") & " @ " & Format(rowSOTORDR2.Item("DATE_3"), "MM/yy")
                End If
                If Val(rowSOTORDR2.Item("QTY_4") & "") <> 0 Then
                    MULTI_Q = MULTI_Q + 1
                    If AVAILABILITY <> "" Then AVAILABILITY = AVAILABILITY & "; "
                    AVAILABILITY = AVAILABILITY & rowSOTORDR2.Item("QTY_4") & " @ " & Format(rowSOTORDR2.Item("DATE_4"), "MM/yy")
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
                range.Cells(RX, CX).ColumnWidth = 14
                Dim STYLE_STATUS As String = rowSOTORDR2.Item("STYLE_STATUS") & ""
                If STYLE_STATUS = "D" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Red
                    STATUS = "Discontinued"
                End If

                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("STYLE_DESC")
                range.Cells(RX, CX).ColumnWidth = 54
                oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Beige
                CX += 1 : oSheet.Cells(RX, CX).Value = "'" & rowSOTORDR2.Item("COLOR_CODE")
                range.Cells(RX, CX).ColumnWidth = 8
                Dim STYLE_COLOR_STATUS As String = rowSOTORDR2.Item("STYLE_COLOR_STATUS") & ""
                If STYLE_COLOR_STATUS = "D" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Red
                    STATUS = "Discontinued"
                End If
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_QTY")
                range.Cells(RX, CX).ColumnWidth = 10
                TOT_ORD = TOT_ORD + Val(rowSOTORDR2.Item("ORDR_QTY"))
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_QTY_PICK")
                range.Cells(RX, CX).ColumnWidth = 10
                TOT_PICK = TOT_PICK + Val(rowSOTORDR2.Item("ORDR_QTY_PICK"))
                If Val(rowSOTORDR2.Item("ORDR_QTY_PICK")) <> 0 Then
                    STATUS = "In Pick"
                End If
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_QTY_SHIP")
                range.Cells(RX, CX).ColumnWidth = 10
                TOT_SHP = Val(TOT_SHP) + Val(rowSOTORDR2.Item("ORDR_QTY_SHIP"))
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_QTY_CANC")
                range.Cells(RX, CX).ColumnWidth = 12
                TOT_CAN = TOT_CAN + Val(rowSOTORDR2.Item("ORDR_QTY_CANC"))
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_QTY_OPEN")
                range.Cells(RX, CX).ColumnWidth = 12
                TOT_OPN = TOT_OPN + Val(rowSOTORDR2.Item("ORDR_QTY_OPEN"))
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_RELEASE_AVAIL")
                range.Cells(RX, CX).ColumnWidth = 13
                range.Cells(RX, CX).Columns.Hidden = True
                CX += 1 : oSheet.Cells(RX, CX).Value = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                range.Cells(RX, CX).ColumnWidth = 10
                range.Cells(RX, CX).NumberFormat = “####.00”
                CX += 1 : oSheet.Cells(RX, CX).Value = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE")) * Val(rowSOTORDR2.Item("ORDR_QTY_OPEN"))
                range.Cells(RX, CX).NumberFormat = “###,###.00”
                range.Cells(RX, CX).ColumnWidth = 12
                TOT_SLS = TOT_SLS + Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE")) * Val(rowSOTORDR2.Item("ORDR_QTY_OPEN"))
                CX += 1 : oSheet.Cells(RX, CX).Value = AVAILABILITY
                range.Cells(RX, CX).ColumnWidth = 24
                range.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Center
                If MULTI_Q = 1 And LEGEND = "In Stock" Then
                    oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.LightGreen
                Else
                    oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.Yellow
                End If

                CX += 1 : oSheet.Cells(RX, CX).Value = LEGEND
                range.Cells(RX, CX).ColumnWidth = 13
                CX += 1 : oSheet.Cells(RX, CX).Value = STATUS
                range.Cells(RX, CX).ColumnWidth = 13
                range.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Center
                If STATUS = "Discontinued" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Red
                End If
                If STATUS = "In Pick" Then
                    oSheet.Cells(RX, CX).Font.Color = SpreadsheetGear.Colors.Blue
                End If

            End If

        Next
        RX += 1 : oSheet.Cells(RX, 0).Value = "Totals"
        For CX = 0 To 13
            oSheet.Cells(RX, CX).Interior.Color = SpreadsheetGear.Colors.LightGray
        Next

        RX += 1 : oSheet.Cells(RX, 0).Value = TOT_CT
        oSheet.Cells(RX, 4).Value = TOT_ORD
        oSheet.Cells(RX, 5).Value = TOT_PICK
        oSheet.Cells(RX, 6).Value = TOT_SHP
        oSheet.Cells(RX, 7).Value = TOT_CAN
        oSheet.Cells(RX, 8).Value = TOT_OPN
        oSheet.Cells(RX, 11).Value = TOT_SLS
        range.Cells(RX, 11).NumberFormat = “###,###.00”

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


End Class