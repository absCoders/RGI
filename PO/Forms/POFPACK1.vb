Imports Infragistics.Win.UltraWinGrid

Public Class POFPACK1

    Dim rowPOTPACK1 As DataRow
    Dim PACK_LIST_NO As String
    Dim PACK_LIST_NO_new As String
    Dim PACK_LIST_STATUS As String

    Dim rowTATUSER1 As DataRow

    Dim sqlPOTPACKX As String
    Dim VEND_CODE As String = ""
    Dim VEND_CODE_USER As String = ""

    Dim PO_REFERENCE As String = ""
    Dim STYLE_CODE_PFX As String = ""
    Dim INITIAL_ORDER As String = ""
    Dim PO_ORDER_NO As String = ""
    Dim PO_SPEC_ORDR_NO As String = ""
    Dim Appearance_Red As New Infragistics.Win.Appearance
    Dim unFinalize As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Appearance_Red.ForeColor = Drawing.Color.Red

        If MENU_ITEM_OBJECT = "POTLTRCI" Then
            InquiryMode = True
        End If

        rowTATUSER1 = LookUp("TATUSER1", ASCMAIN1.USER_ID)
        If rowTATUSER1 IsNot Nothing AndAlso rowTATUSER1.Item("VEND_CODE") & "" <> "" Then
            VEND_CODE_USER = rowTATUSER1.Item("VEND_CODE")
        Else
            VEND_CODE_USER = ""
        End If

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
        End With

        Get_PARM("GLTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")

        With dst
            sqlPOTPACKX = "Select POTPACK1.*,APTVEND1.VEND_NAME" & vbCrLf _
                & " from POTPACK1,APTVEND1" & vbCrLf _
                & " where APTVEND1.VEND_CODE = POTPACK1.VEND_CODE"
            ASCMAIN1.sql = sqlPOTPACKX ' & "  and POTPACK1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "POTPACKX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "POTPACK1", "*")

            Create_TDA(.Tables.Add, "POTPACK2", "*", 1)

            ASCMAIN1.sql = "Select POTPACK3.*, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & " from POTPACK3, ICTSTYL1 where ICTSTYL1.STYLE_CODE = POTPACK3.STYLE_CODE and POTPACK3.PACK_LIST_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPACK3", "**", 0, True, "V")
            With .Tables("POTPACK3")
                .Columns.Add("TOTAL_UNITS", GetType(System.Int32), "CARTON_COUNT * CARTON_PACK")
                .Columns.Add("CARTON_NO_START", GetType(System.Int32))
                .Columns.Add("CARTON_NO_END", GetType(System.Int32), "CARTON_NO_START + CARTON_COUNT -1")
                .Columns.Add("TOTAL_GRS_WGT", GetType(System.Decimal), "CARTON_COUNT * CARTON_GRS_WGT")
                .Columns.Add("TOTAL_NET_WGT", GetType(System.Decimal), "CARTON_COUNT * CARTON_NET_WGT")
                ' .Columns.Add("STYLE_WEIGHT", GetType(System.Decimal), "IIF(ISNULL(CARTON_COUNT,0) = 0, 0, ISNULL(CARTON_NET_WGT,0) / ISNULL(CARTON_COUNT,0))")

            End With

            Create_Relation("POTPACK2", "POTPACK3", "PACK_LIST_NO,PACK_LIST_SHEET_NO")

            With .Tables("POTPACK2")
                .Columns.Add("COLOR_DESC")
                .Columns.Add("TOTAL_CARTONS", GetType(System.Int32), "SUM(CHILD.CARTON_COUNT)")
                .Columns.Add("TOTAL_UNITS", GetType(System.Int32), "SUM(CHILD.TOTAL_UNITS)")
                .Columns.Add("TOTAL_GRS_WGT", GetType(System.Decimal), "SUM(CHILD.TOTAL_GRS_WGT)")
                .Columns.Add("TOTAL_NET_WGT", GetType(System.Decimal), "SUM(CHILD.TOTAL_NET_WGT)")
            End With

            'With .Tables("POTPACK3")
            '    .Columns("CARTON_NO_START").Expression = "PARENT.CARTON_NO_START"
            'End With


            ASCMAIN1.sql = "Select * from POTORDR1 where PO_REFERENCE = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "POTORDR2", "*", 1, False)

            ASCMAIN1.sql = "Select PO_ORDER_NO, PO_REFERENCE, PO_SPEC_ORDR_NO, PO_DATE_SHIP_BY, PO_DATE_ETA from POTORDR1 where VEND_CODE = :PARM1 and PO_STATUS = 'O'"
            Create_TDA(.Tables.Add, "POTORDRR", "**", 0, False, "V")

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, SUM (PO_QTY_OPN) PO_QTY_OPN from POTORDR2 where PO_ORDER_NO = :PARM1 group by STYLE_CODE, COLOR_CODE"
            Create_TDA(.Tables.Add, "POTORDRD", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "WHTSCSEQ", "*", 0, False)
            Fill_Records("WHTSCSEQ")

            ASCMAIN1.sql = "Select POTLPNL1.* from POTLPNL1 where PACK_LIST_NO = :PARM1 and BARCODE_STATUS = 'A'"
            Create_TDA(.Tables.Add, "POTLPNL1", "**", 0, True, "V")

        End With

        grdPOTPACKX.DataSource = dst.Tables("POTPACKX")

        grdPOTPACK2.DataSource = dst.Tables("POTPACK2")
        grdPOTPACK3.DataSource = dst.Tables("POTPACK3")
        grdPOTLPNL1.DataSource = dst.Tables("POTLPNL1")

        grdPOTORDRR.DataSource = dst.Tables("POTORDRR")
        grdPOTORDRD.DataSource = dst.Tables("POTORDRD")

        Create_Summary(grdPOTPACKX, "PACK_LIST_NO", "Count")
        ' Create_Summary(grdPOTPACKX, New String() {"LC_AMT", "LC_PMTS", "LC_FEES", "LC_OPEN"})

        Create_Summary(grdPOTPACK2, "PACK_LIST_SHEET_NO", "Count")
        Create_Summary(grdPOTPACK2, New String() {"TOTAL_CARTONS", "TOTAL_UNITS", "TOTAL_GRS_WGT", "TOTAL_NET_WGT", "CARTON_COUNT"})

        Create_Summary(grdPOTPACK3, "PACK_LIST_SHEET_LNO", "Count")
        Create_Summary(grdPOTPACK3, New String() {"CARTON_COUNT", "TOTAL_UNITS", "TOTAL_GRS_WGT", "TOTAL_NET_WGT"})

        With grdPOTPACKX.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"PACK_LIST_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("PACK_LIST_NO").Header.Fixed = True
        End With

        With grdPOTPACK2.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                GCOL.CellActivation = Activation.NoEdit
                If New String() {"PACK_LIST_DETAILS", "CARTON_NO_START", "CARTON_COUNT", "CARTON_PACK"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.DodgerBlue '.LightGreen
                    'GCOL.CellAppearance.BackColor = System.Drawing.Color.LightGreen
                    GCOL.CellActivation = Activation.AllowEdit
                ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next

        End With

        With grdPOTPACK3.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                GCOL.CellActivation = Activation.NoEdit
                If New String() {"CARTON_COUNT", "CARTON_PACK", "CARTON_GRS_WGT", "CARTON_NET_WGT", "CARTON_DIMENSIONS"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.DodgerBlue '.LightGreen
                    ' GCOL.CellAppearance.BackColor = System.Drawing.Color.LightGreen
                    GCOL.CellActivation = Activation.AllowEdit
                ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next

        End With



        '  ASCMAIN1.Add_Value_List(grdPOTPACKX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'POTPACK1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdPOTPACKX, "PACK_LIST_STATUS", Nothing, New String() {":", "O:Open", "F:Finalized"})

        grpHeader.Visible = False

        '  Absx1.txtFor("CURR_CODE").ReadOnly = True

        Show_Filter(grdPOTPACKX, True)
        Show_Filter(grdPOTORDRR, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                unFinalize = False

                VEND_CODE = ""
                If Absx1.txtFor("VEND_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Supplier Code"
                Else
                    Dim row As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Supplier Code Entered Is Not Valid"
                    Else
                        If row.Item("VEND_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Supplier Status Is Not Active"
                        Else
                            VEND_CODE = Absx1.txtFor("VEND_CODE").Text
                        End If
                    End If
                End If

                If VEND_CODE <> VEND_CODE_USER Then
                    EMsg &= vbCr & "Invalid Vendor"
                End If
                'Dim DT As Date = Absx1.dteFor("PACK_INV_DATE").Value
                'If DT & "" = "" Then
                '    EMsg &= vbCr & "Invoice Date is Mandatory"
                'Else
                '    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                'End If

                PO_ORDER_NO = ""
                PO_SPEC_ORDR_NO = ""
                INITIAL_ORDER = "0"
                If Absx1.txtFor("PO_REFERENCE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid PO Reference"
                Else
                    PO_REFERENCE = Absx1.txtFor("PO_REFERENCE").Text
                    Fill_Records("POTORDR1", PO_REFERENCE)
                    If dst.Tables("POTORDR1").Rows.Count > 1 Then
                        EMsg &= vbCr & $"More than 1 Vandale PO is associated with PO Reference {PO_REFERENCE}"
                    ElseIf dst.Tables("POTORDR1").Rows.Count = 0 Then
                        EMsg &= vbCr & $"No record PO Reference {PO_REFERENCE}"
                    Else
                        Dim row As DataRow = dst.Tables("POTORDR1").Rows(0)
                        If row.Item("VEND_CODE") & "" <> VEND_CODE Then
                            EMsg &= vbCr & $"Invalid PO Reference {PO_REFERENCE}"
                        ElseIf row.Item("PO_STATUS") & "" <> "O" Then
                            EMsg &= vbCr & $"PO Reference {PO_REFERENCE} is not Open"
                        Else
                            PO_ORDER_NO = row.Item("PO_ORDER_NO")
                            PO_SPEC_ORDR_NO = row.Item("PO_SPEC_ORDR_NO") & ""
                            STYLE_CODE_PFX = row.Item("STYLE_CODE_PFX") & ""
                            If STYLE_CODE_PFX <> "" Then Absx1.txtFor("STYLE_CODE_PFX").Text = STYLE_CODE_PFX
                            If PO_SPEC_ORDR_NO.ToUpper.StartsWith("INITIAL") Then INITIAL_ORDER = "1"
                        End If
                    End If

                    If INITIAL_ORDER = "1" Then
                        ASCMAIN1.sql = "Select Count (*) from POTORDR2 where PO_ORDER_NO = :PARM1"
                        Dim PO_lines As Integer = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {PO_ORDER_NO}))
                        If PO_lines = 0 Then
                            EMsg &= vbCr & $"No Lines on PO {PO_REFERENCE}"
                        End If
                    Else
                        If Absx1.txtFor("STYLE_CODE_PFX").Text.Length = 0 Then
                            EMsg &= vbCr & "You must enter a Style Code Prefix"
                        Else
                            If PO_ORDER_NO <> "" Then
                                STYLE_CODE_PFX = Absx1.txtFor("STYLE_CODE_PFX").Text
                                ASCMAIN1.sql = "Select Count (*) from POTORDR2 where PO_ORDER_NO = :PARM1 and STYLE_CODE like :PARM2 || '%'"
                                Dim PO_lines As Integer = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {PO_ORDER_NO, STYLE_CODE_PFX}))
                                If PO_lines = 0 Then
                                    EMsg &= vbCr & $"No Lines on PO {PO_REFERENCE} with Style Code Prefix {STYLE_CODE_PFX}"
                                End If
                            End If
                        End If
                    End If

                End If

                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & Absx1.txtFor("VEND_CODE").Text) Then Exit Sub
                'End If

            Case "View", "Edit"

                unFinalize = False

                PACK_LIST_NO = Absx1.txtFor("PACK_LIST_NO").Text
                If PACK_LIST_NO = "" Then
                    EMsg &= vbCr & "You must specify Packing List No to View"
                Else
                    Dim row As DataRow = LookUp("POTPACK1", PACK_LIST_NO)
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & PACK_LIST_NO & " on File"
                    Else
                        If VEND_CODE_USER <> "" And row.Item("VEND_CODE") <> VEND_CODE_USER Then
                            EMsg &= vbCr & "No Record of Document " & PACK_LIST_NO & " on File"
                        End If

                        If eItemKey = "Edit" Then

                            If row.Item("PACK_LIST_STATUS") & "" = "F" Then
                                Dim VBKG_NO As String = row.Item("VBKG_NO") & ""
                                If VBKG_NO <> "" Then
                                    EMsg &= vbCr & $"Packing List {PACK_LIST_NO} has already been listed on Booking No {VBKG_NO}"
                                    EMsg &= vbCr & "- Un-Finalizing Not permitted"
                                Else
                                    If MsgBox("Already Finalized - do you want to un-Finalize?", MsgBoxStyle.YesNo,
                                          "IMPORTANT - LPNs will be regenerated") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                    unFinalize = True
                                End If
                            End If

                            If row.Item("PACK_LIST_STATUS") & "" = "F" And Not unFinalize Then
                                EMsg &= vbCr & "Document " & PACK_LIST_NO & " Is Finalized - no editing permitted"
                            End If

                            If EMsg = "" Then
                                If Not ASCMAIN1.Logical_Lock("POTPACK1", PACK_LIST_NO) Then Exit Sub
                                ' If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & row.Item("VEND_CODE")) Then Exit Sub

                            End If
                        End If
                    End If
                End If

            Case "Update"

                If Absx1.txtFor("PACK_LIST_DESC").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Packing List Description"
                End If

                Dim DT As Date = Absx1.dteFor("PACK_LIST_DATE").Value & ""
                If DT & "" = "" Then
                    EMsg &= vbCr & "Packing List Date Is Mandatory"
                Else
                    '  TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                End If

                If Absx1.txtFor("VEND_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Supplier Code"
                Else
                    Dim row As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Supplier Entered Is Not Valid"
                    Else
                        If row.Item("VEND_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Supplier Entered Is Not Active"
                        End If
                    End If
                End If

                Dim TOTAL_CARTONS As Integer = Val(dst.Tables("POTPACK2").Compute("SUM(CARTON_COUNT)", "") & "")
                Dim CARTON_COUNTer As Integer = 0
                For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "CARTON_NO_START")
                    Dim CARTON_COUNT As Integer = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")

                    Dim CARTON_NO_START As Integer = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")
                    Dim PACK_LIST_SHEET_NO As Integer = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")
                    If CARTON_NO_START <> CARTON_COUNTer + 1 Then
                        EMsg &= vbCr & $"Unexpected Starting Carton {CStr(CARTON_NO_START)} on Sheet {CStr(PACK_LIST_SHEET_NO)} - was expecting {CStr(CARTON_COUNTer + 1)}"
                        CARTON_COUNTer += CARTON_COUNT
                        Exit For
                    End If
                    Dim SQLW As String = $"PACK_LIST_NO = '{PACK_LIST_NO}' and PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}"
                    For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select(SQLW, "CARTON_NO_START")
                        CARTON_COUNTer += 1

                        Dim CARTON_NO_START3 As Integer = Val(rowPOTPACK3.Item("CARTON_NO_START") & "")
                        Dim CARTON_NO_END3 As Integer = Val(rowPOTPACK3.Item("CARTON_NO_END") & "")
                        Dim CARTON_COUNT3 As Integer = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                        Dim PACK_LIST_SHEET_LNO As Integer = Val(rowPOTPACK3.Item("PACK_LIST_SHEET_LNO") & "")
                        If CARTON_NO_START3 <> CARTON_COUNTer Then
                            EMsg &= vbCr & $"Unexpected Starting Carton {CStr(CARTON_NO_START3)} on Sheet {CStr(PACK_LIST_SHEET_NO)}, Line {PACK_LIST_SHEET_LNO} - was expecting {CStr(CARTON_COUNTer)}"
                            Exit For
                        End If
                        CARTON_COUNTer += CARTON_COUNT3 - 1
                        If CARTON_NO_END3 <> CARTON_COUNTer Then
                            EMsg &= vbCr & $"Unexpected Ending Carton {CStr(CARTON_NO_END3)} on Sheet {CStr(PACK_LIST_SHEET_NO)}, Line {PACK_LIST_SHEET_LNO} - was expecting {CStr(CARTON_COUNTer)}"
                            Exit For
                        End If
                    Next
                Next

                Dim EMsg2 As String = Generate_Carton_Nos()
                EMsg &= EMsg2

                If EMsg = "" Then
                    If chkFinalize.Checked Then
                        If MsgBox("You have chosen to Finalize this Packing List upon Update." _
                                & vbCrLf & vbCrLf & "Once you have Finalized, LPNs for Barcodes will be generated," _
                                & vbCrLf & " And you will Not be able to make further changes." _
                                & vbCrLf & vbCrLf & "Are you sure that you want to Finalize this Packing List?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Delete"


            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Print Labels"
                If dst.Tables("POTPACK3").Select("BARCODE_START IS NULL").Length > 0 Then
                    Dim RESULT As MsgBoxResult = MsgBox("Some Packing Details do not have LPNs." & vbCrLf & vbCrLf & "(Re)Generate LPNs Now?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Verification to Generate LPNs")
                    If RESULT = MsgBoxResult.Cancel Then
                        Exit Sub
                    ElseIf RESULT = MsgBoxResult.Yes Then
                        Generate_LPN_Report_File()
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Add Sheet"

                WorkbookView1.GetLock()
                Dim wsx As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.ActiveWorksheet
                'Dim ws As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.Worksheets.Add()
                Dim newSheet As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.ActiveWorksheet.CopyAfter(WorkbookView1.ActiveWorkbook.ActiveWorksheet)
                WorkbookView1.ReleaseLock()

            Case "Print Labels"
                Print_Labels()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()

                If chkFinalize.Checked Then
                    If chkFinalize.Tag & "" = "X" Then
                        chkFinalize.Tag = ""
                    Else
                        Generate_LPN_Report_File()
                        Print_Labels()
                    End If

                    Check_for_Overbooked
                End If

                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Export XLS"
                Export_XLS()

            Case "Generate Start/End"
                Generate_Carton_Nos()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode

                    If EntryMode = "V" And ScreenMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" And EntryMode <> "E" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                        .Items("Delete").Settings.Enabled = not_iScreenMode
                        .Items("Print Labels").Visible = rowPOTPACK1.Item("PACK_LIST_STATUS") & "" = "F" And EntryMode = "V"
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                        .Items("Delete").Settings.Enabled = iScreenMode
                        .Items("Print Labels").Visible = False
                    End If

                    If ScreenMode Then
                        .Items("New").Visible = False
                        .Items("View").Visible = False
                        .Items("Edit").Visible = (EntryMode = "V")
                    Else
                        .Items("New").Visible = True
                        .Items("View").Visible = True
                        .Items("Edit").Visible = True
                    End If

                    .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                    .Items("Done").Visible = ScreenMode And (EntryMode = "V")

                    If ScreenMode And EntryMode = "E" Then
                        .Items("Delete").Visible = True
                        .Items("Delete").Visible = False ' NOT UNTIL WE FIGURE OUT PROTECTIONS
                    Else
                        .Items("Delete").Visible = False
                    End If

                    If ScreenMode Then
                        .Items("Separator1").Visible = True
                        .Items("Export XLS").Visible = True
                        .Items("Generate Start/End").Visible = (EntryMode = "N" Or EntryMode = "E") And Not (INITIAL_ORDER = "1")
                    Else
                        .Items("Separator1").Visible = False
                        .Items("Export XLS").Visible = False
                        .Items("Generate Start/End").Visible = False
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                        .Items("Add Sheet").Visible = False ' True
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                        .Items("Add Sheet").Visible = False
                    End If
                End With

                ' .Groups("Totals").Visible = ScreenMode
                .Groups("Show").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        chkFinalize.Visible = Not InquiryMode And (EntryMode = "N" Or EntryMode = "E")

        splPOTPACK3.Panel2Collapsed = (Not ScreenMode Or EntryMode <> "V") OrElse rowPOTPACK1.Item("PACK_LIST_STATUS") <> "F"

        splPOTPACKX.Visible = Not ScreenMode

        If ScreenMode Then
            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            'If EntryMode = "E" Or EntryMode = "N" Then
            '    Set_Read_Only_for_ctl(Absx1.txtFor("LC_REF_NO"), False)
            '    Set_Read_Only_for_ctl(Absx1.dteFor("LC_DATE"), False)
            '    '   Set_Read_Only_for_ctl(Absx1.txtFor("CURR_CODE"), True)
            'End If

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTPACK2, grdPOTPACK3}
                For Each GCOL As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                    If GCOL.CellActivation = Activation.AllowEdit Then
                        If EntryMode = "N" Or EntryMode = "E" Then
                            ' GCOL.CellAppearance.BackColor = System.Drawing.Color.Khaki
                            GCOL.CellAppearance.BackColor = System.Drawing.Color.PowderBlue
                        Else
                            GCOL.CellAppearance.BackColor = System.Drawing.Color.Empty
                        End If
                    End If
                Next

                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name = "grdPOTPACK3" Or grd.Name = "grdPOTPACK2" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.True
                            .AllowUpdate = DefaultableBoolean.True
                        Else
                            '    '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            '    '.AllowDelete = DefaultableBoolean.True
                            '    '.AllowUpdate = DefaultableBoolean.True

                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.True
                        End If

                    End With
                Else
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                End If
            Next

            With grdPOTPACK3.DisplayLayout.Bands(0)
                If Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
                    .Columns("BARCODE_START").Hidden = True
                    .Columns("BARCODE_END").Hidden = True
                Else
                    .Columns("BARCODE_START").Hidden = False
                    .Columns("BARCODE_END").Hidden = False
                End If

                If (INITIAL_ORDER = "1") Then
                    dst.Tables("POTPACK3").Columns("TOTAL_GRS_WGT").Expression = "CARTON_GRS_WGT * PARENT(POTPACK2_POTPACK3).CARTON_COUNT"
                    dst.Tables("POTPACK3").Columns("TOTAL_NET_WGT").Expression = "CARTON_NET_WGT * PARENT(POTPACK2_POTPACK3).CARTON_COUNT"
                Else
                    dst.Tables("POTPACK3").Columns("TOTAL_GRS_WGT").Expression = "CARTON_COUNT * CARTON_GRS_WGT"
                    dst.Tables("POTPACK3").Columns("TOTAL_NET_WGT").Expression = "CARTON_COUNT * CARTON_NET_WGT"
                End If

            End With

            With grdPOTPACK3.DisplayLayout.Bands(0)
                For Each C As String In New String() {"CARTON_COUNT", "CARTON_GRS_WGT", "CARTON_NET_WGT", "CARTON_DIMENSIONS", "CARTON_NO_START", "CARTON_NO_END", "BARCODE_START", "BARCODE_END"}
                    .Columns(C).Hidden = (INITIAL_ORDER = "1")
                Next
            End With

            With grdPOTPACK2.DisplayLayout.Bands(0)
                For Each C As String In New String() {"CARTON_NO_START", "BARCODE_START", "BARCODE_END"}
                    .Columns(C).Hidden = (INITIAL_ORDER = "1")
                Next
                For Each C As String In New String() {"CARTON_PACK", "CARTON_COUNT"}
                    .Columns(C).Hidden = Not (INITIAL_ORDER = "1")
                Next

                If (INITIAL_ORDER = "1") Then
                    .Columns("TOTAL_CARTONS").Header.Caption = "Styles"
                    dst.Tables("POTPACK2").Columns("TOTAL_UNITS").Expression = "CARTON_PACK * CARTON_COUNT"
                Else
                    .Columns("TOTAL_CARTONS").Header.Caption = "Cartons"
                    dst.Tables("POTPACK2").Columns("TOTAL_UNITS").Expression = "SUM(CHILD.TOTAL_UNITS)"
                End If
            End With



            Set_Read_Only_for_ctl(Absx1.optFor("PACK_LIST_STATUS"), True)
            Set_Read_Only_for_ctl(Absx1.chkFor("INITIAL_ORDER"), True)

            lblSTYLE_CODE_PFX.Visible = Not (INITIAL_ORDER = "1")
            txtSTYLE_CODE_PFX.Visible = Not (INITIAL_ORDER = "1")

            Display_Totals()

        Else
            Clear_Record()

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTPACK1", "POTPACK2", "POTPACK3", "POTLPNL1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If VEND_CODE_USER <> "" Then
            Absx1.txtFor("VEND_CODE").Text = VEND_CODE_USER
            Absx1.txtFor("VEND_CODE").ReadOnly = True
        Else
            Absx1.txtFor("VEND_CODE").Text = ""
        End If

        chkFinalize.Checked = False
        chkFinalize.Tag = ""

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowPOTPACK1 = dst.Tables("POTPACK1").NewRow
            PACK_LIST_NO = ASCMAIN1.Next_Control_No("POTPACK1.PACK_LIST_NO")
            rowPOTPACK1.Item("PACK_LIST_NO") = PACK_LIST_NO
            rowPOTPACK1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowPOTPACK1.Item("PACK_LIST_DESC") = PO_SPEC_ORDR_NO
            rowPOTPACK1.Item("PACK_LIST_DATE") = DATETIME_STAMP.Date
            rowPOTPACK1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTPACK1.Item("INIT_DATE") = DATETIME_STAMP
            rowPOTPACK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTPACK1.Item("LAST_DATE") = DATETIME_STAMP
            rowPOTPACK1.Item("PACK_LIST_STATUS") = "O"
            rowPOTPACK1.Item("PO_REFERENCE") = HFs("PO_REFERENCE")
            If INITIAL_ORDER = "1" Then
                rowPOTPACK1.Item("STYLE_CODE_PFX") = HFs("STYLE_CODE_PFX")
                rowPOTPACK1.Item("INITIAL_ORDER") = "1"
            Else
                rowPOTPACK1.Item("STYLE_CODE_PFX") = ""
                rowPOTPACK1.Item("INITIAL_ORDER") = "0"
            End If

            rowPOTPACK1.Item("PO_ORDER_NO") = PO_ORDER_NO

            dst.Tables("POTPACK1").Rows.Add(rowPOTPACK1)

        Else
            rowPOTPACK1 = Fill_Record("POTPACK1", PACK_LIST_NO)
            VEND_CODE = rowPOTPACK1.Item("VEND_CODE")
            If VEND_CODE_USER <> "" And VEND_CODE <> VEND_CODE_USER Then
                MsgBox("Issue with Vendor Code", MsgBoxStyle.OkOnly, "Please Call ABS")
                Throw New Exception("Issue with Vendor Code")
            End If
            PO_REFERENCE = rowPOTPACK1.Item("PO_REFERENCE")
            STYLE_CODE_PFX = rowPOTPACK1.Item("STYLE_CODE_PFX")
            PO_ORDER_NO = rowPOTPACK1.Item("PO_ORDER_NO")
            INITIAL_ORDER = rowPOTPACK1.Item("INITIAL_ORDER")

            If unFinalize Then
                rowPOTPACK1.Item("PACK_LIST_STATUS") = "O"
            End If

            dst.AcceptChanges()
        End If

        PACK_LIST_STATUS = rowPOTPACK1.Item("PACK_LIST_STATUS")

        EnforceConstraints(False)

        Fill_Records("POTORDR2", PO_ORDER_NO)

        If EntryMode = "N" Then

            Fill_Records("POTORDRD", PO_ORDER_NO)

            If INITIAL_ORDER = "1" Then

                Dim rowPOTPACK2 As DataRow = dst.Tables("POTPACK2").NewRow
                rowPOTPACK2.Item("PACK_LIST_NO") = PACK_LIST_NO
                rowPOTPACK2.Item("PACK_LIST_SHEET_NO") = 1
                rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") = PO_REFERENCE & "-" & CStr(PO_SPEC_ORDR_NO)
                Dim COLOR_CODE As String = "AST"
                rowPOTPACK2.Item("COLOR_CODE") = COLOR_CODE
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                rowPOTPACK2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
                rowPOTPACK2.Item("CARTON_NO_START") = 1
                dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2)

                Dim PACK_LIST_SHEET_LNO_ctr As Integer = 0
                For Each rowPOTORDRD As DataRow In dst.Tables("POTORDRD").Select("", "STYLE_CODE, COLOR_CODE")
                    COLOR_CODE = rowPOTORDRD.Item("COLOR_CODE")
                    Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").NewRow
                    With rowPOTPACK3
                        .Item("PACK_LIST_NO") = PACK_LIST_NO
                        .Item("PACK_LIST_SHEET_NO") = 1
                        PACK_LIST_SHEET_LNO_ctr += 1
                        .Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_ctr
                        Dim STYLE_CODE As String = rowPOTORDRD.Item("STYLE_CODE")
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                        '.Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                        .Item("SIZE_CODE") = rowICTSTYL1.Item("SIZE_CODE")

                        .Item("CARTON_COUNT") = 1

                        Dim rowWHTSCSEQs() As DataRow = dst.Tables("WHTSCSEQ").Select($"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'")
                    If rowWHTSCSEQs.Length = 0 Then
                        MsgBox($"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                    ElseIf rowWHTSCSEQs.Length > 1 Then
                        MsgBox($"More than 1 Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                    Else
                        .Item("CARTON_ID") = rowWHTSCSEQs(0).Item("STYLE_SEQ")
                    End If

                    End With
            dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3)
                Next

            Else

                Dim PACK_LIST_SHEET_NO_ctr As Integer = 0
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR2"), New String() {"COLOR_CODE"}).Select("", "COLOR_CODE")
                    Dim rowPOTPACK2 As DataRow = dst.Tables("POTPACK2").NewRow
                    rowPOTPACK2.Item("PACK_LIST_NO") = PACK_LIST_NO
                    PACK_LIST_SHEET_NO_ctr += 1
                    rowPOTPACK2.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr
                    rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") = PO_REFERENCE & "-" & CStr(PACK_LIST_SHEET_NO_ctr)
                    Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                    rowPOTPACK2.Item("COLOR_CODE") = COLOR_CODE
                    Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                    rowPOTPACK2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
                    dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2)

                    Dim PACK_LIST_SHEET_LNO_ctr As Integer = 0
                    For Each rowPOTORDRD As DataRow In dst.Tables("POTORDRD").Select($"COLOR_CODE = '{COLOR_CODE}'", "STYLE_CODE")
                        Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").NewRow
                        With rowPOTPACK3
                            .Item("PACK_LIST_NO") = PACK_LIST_NO
                            .Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr
                            PACK_LIST_SHEET_LNO_ctr += 1
                            .Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_ctr
                            Dim STYLE_CODE As String = rowPOTORDRD.Item("STYLE_CODE")
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                            ' .Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                            .Item("SIZE_CODE") = rowICTSTYL1.Item("SIZE_CODE")

                            Dim rowWHTSCSEQs() As DataRow = dst.Tables("WHTSCSEQ").Select($"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'")
                            If rowWHTSCSEQs.Length = 0 Then
                                MsgBox($"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                            ElseIf rowWHTSCSEQs.Length > 1 Then
                                MsgBox($"More than 1 Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                            Else
                                .Item("CARTON_ID") = rowWHTSCSEQs(0).Item("STYLE_SEQ")
                            End If

                        End With
                        dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3)
                    Next

                Next
            End If

        Else
            Fill_Records("POTPACK2", PACK_LIST_NO)
            Fill_Records("POTPACK3", PACK_LIST_NO)

            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select("")
                Dim STYLE_CODE As String = rowPOTPACK3.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                rowPOTPACK3.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            Next
            Generate_Carton_Nos()
        End If

        If EntryMode = "V" And rowPOTPACK1.Item("PACK_LIST_STATUS") = "F" Then
            Fill_Records("POTLPNL1", PACK_LIST_NO)
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grdPOTPACK2.Rows
            grow.PerformAutoSize()
        Next

        Dim FILENAME As String = ASCMAIN1.Folders("Work") & "\" & "PACKLIST.xlsx"
        WorkbookView1.GetLock()
        WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        WorkbookView1.ReleaseLock()

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        If chkFinalize.Checked Then
            rowPOTPACK1.Item("PACK_LIST_STATUS") = "F"

            Dim BARCODE_PFX As String = "Y" ' NEED TO GET THIS FROM VENDOR MASTER
            ' AND VENDORS WITHOUT A PREFIX ARE NOT PERMITTED TO USE THIS SCREEN

            Dim tbl_BARCODE As String = "POTPACK3"
            If INITIAL_ORDER = "1" Then
                tbl_BARCODE = "POTPACK2"
            End If

            Dim generate_LPNs As Boolean = True
            Dim BARCODE_MIN As String = dst.Tables(tbl_BARCODE).Compute("MIN(BARCODE_START)", "")
            If BARCODE_MIN <> "" Then
                generate_LPNs = Generate_LPNs_Test(BARCODE_MIN)
            End If

            If generate_LPNs Then

                For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
                    Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") & ""
                    Dim CARTON_NO_START As Int32 = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")
                    Dim PACK_LIST_SHEET_NO As Int32 = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

                    For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO") ' rowPOTPACK2.GetChildRows("POTPACK2_POTPACK3")
                        Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                        If INITIAL_ORDER = "1" Then
                            CARTON_COUNT = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                        End If
                        Dim BARCODE As String = ASCMAIN1.Next_Control_No("BARCODE_" & BARCODE_PFX, CARTON_COUNT)
                        Dim BARCODE_START As String = BARCODE_PFX & BARCODE
                        Dim BARCODE_END As String = BARCODE_PFX & Format(Val(BARCODE) + CARTON_COUNT - 1, "0000000")

                        Dim rowCompare As DataRow = rowPOTPACK3
                        If INITIAL_ORDER = "1" Then rowCompare = rowPOTPACK2
                        rowCompare.Item("BARCODE_START") = BARCODE_START
                        rowCompare.Item("BARCODE_END") = BARCODE_END

                    Next
                Next

                If BARCODE_MIN <> "" Then
                    MsgBox("Note: LPNs WERE Re-Generated", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Please Note: Labels will be re-printed")
                End If

            Else

                If BARCODE_MIN <> "" Then
                    MsgBox("Note: LPNs were NOT Re-Generated", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Please Note: Labels will NOT be re-printed")
                    chkFinalize.Tag = "X"
                End If

            End If

        End If

        Dim SQLD As String = "PACK_LIST_NO = '" & PACK_LIST_NO & "'"
        INIT_LAST("POTPACK1", False, , True)

        Update_Record_TDA("POTPACK1", SQLD)
        Update_Record_TDA("POTPACK2", SQLD)
        Update_Record_TDA("POTPACK3", SQLD)

        CommitTrans("Update Complete")

    End Sub

    Function Generate_LPNs_Test(BARCODE_MIN As String) As Boolean

        Dim BARCODE_PFX As String = Mid(BARCODE_MIN, 1, 1)
        Dim BARCODE_CTR As Int32 = Val(Mid(BARCODE_MIN, 2))
        Dim regeneration_required As Boolean = False

        For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
            Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") & ""
            Dim CARTON_NO_START As Int32 = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")
            Dim PACK_LIST_SHEET_NO As Int32 = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO")
                Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                If INITIAL_ORDER = "1" Then
                    CARTON_COUNT = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                End If
                Dim BARCODE As String = Format(BARCODE_CTR, "0000000")
                Dim BARCODE_START As String = BARCODE_PFX & BARCODE
                Dim BARCODE_END As String = BARCODE_PFX & Format(Val(BARCODE) + CARTON_COUNT - 1, "0000000")

                Dim rowCompare As DataRow = rowPOTPACK3
                If INITIAL_ORDER = "1" Then rowCompare = rowPOTPACK2
                If rowCompare.Item("BARCODE_START") & "" <> BARCODE_START Or rowCompare.Item("BARCODE_END") & "" <> BARCODE_END Then
                    regeneration_required = True
                    Exit For
                End If


                BARCODE_CTR += CARTON_COUNT
            Next
            If regeneration_required Then Exit For
        Next

        Return regeneration_required
    End Function

    Sub Delete_Record()
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        ' Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"POTPACK1"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where PACK_LIST_NO = '" & PACK_LIST_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("PACK_LIST_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTPACK1"
            E.COLUMN_NAME = "PACK_LIST_NO"
            E.CODE_VALUE = Absx1.txtFor("PACK_LIST_NO").Text
            E.DESC_VALUE = Absx1.txtFor("VEND_CODE").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "POTPACK1"
        E.TABLE_KEY_CAPTION = "LC Events"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PACK_LIST_NO").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                sql_where = "VEND_TYPE = 'S'"
        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTPACKX, "SS", "Show Filter", "Show GroupBox") ', "Move to Pending", "Approve")
        Load_Popup_Menu(grdPOTORDRR, "SS", "Show Filter")
        Load_Popup_Menu(grdPOTPACK3, "B", "Add Line", "Add Lines", "Copy Value to All Lines", "Copy Pattern to Remaining Lines")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
            grd = GRDs(Mid(e.SourceControl.Name, 4))
        End If

        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If


        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name


            Case "grdPOTPACK3"

                If Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_pop.Tools("Add Line").SharedProps.Visible = Not (INITIAL_ORDER = "1")
                    tlb_pop.Tools("Add Lines").SharedProps.Visible = Not (INITIAL_ORDER = "1")

                    If Not grd.ActiveRow.DataChanged And grd.ActiveCell IsNot Nothing AndAlso New String() {"CARTON_DIMENSIONS", "CARTON_PACK"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = True
                    Else
                        tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = False
                    End If

                    If Not grd.ActiveRow.DataChanged And grd.ActiveCell IsNot Nothing AndAlso New String() {"CARTON_PACK"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = (INITIAL_ORDER = "1")
                    Else
                        tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = False
                    End If

                Else
                    tlb_pop.Tools("Add Line").SharedProps.Visible = False
                    tlb_pop.Tools("Add Lines").SharedProps.Visible = False
                    tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = False
                    tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = False
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdSPTSFOC9"
                '    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                '        tlb_btn.SharedProps.Visible = True
                '    Else
                '        tlb_btn.SharedProps.Visible = False
                '    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Add Line"
                Dim PACK_LIST_SHEET_NO As Integer = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")
                Dim PACK_LIST_SHEET_LNO As Integer = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_LNO").Value & "")
                Dim PACK_LIST_SHEET_LNO_max As Integer = dst.Tables("POTPACK3").Compute("MAX(PACK_LIST_SHEET_LNO)", $"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}")
                Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO, PACK_LIST_SHEET_LNO})
                Dim rowPOTPACK3_new As DataRow = dst.Tables("POTPACK3").NewRow
                rowPOTPACK3_new.ItemArray = rowPOTPACK3.ItemArray
                PACK_LIST_SHEET_LNO_max += 1
                rowPOTPACK3_new.Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_max
                rowPOTPACK3_new.Item("BARCODE_START") = DBNull.Value
                rowPOTPACK3_new.Item("BARCODE_END") = DBNull.Value
                dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_new)
                Sort_grdColumns(grdPOTPACK3, "STYLE_CODE,COLOR_CODE", True)

            Case "Add Lines"
                Dim PACK_LIST_SHEET_NO As Integer = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")
                Dim PACK_LIST_SHEET_LNO_max As Integer = dst.Tables("POTPACK3").Compute("MAX(PACK_LIST_SHEET_LNO)", $"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}")

                Dim STYLE_CODEs As New List(Of String)
                For Each row3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "PACK_LIST_SHEET_LNO")
                    Dim PACK_LIST_SHEET_LNO As Integer = Val(row3.Item("PACK_LIST_SHEET_LNO") & "")
                    Dim STYLE_CODE As String = row3.Item("STYLE_CODE") & ""
                    If Not STYLE_CODEs.Contains(STYLE_CODE) Then
                        Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO, PACK_LIST_SHEET_LNO})
                        Dim rowPOTPACK3_new As DataRow = dst.Tables("POTPACK3").NewRow
                        rowPOTPACK3_new.ItemArray = rowPOTPACK3.ItemArray
                        PACK_LIST_SHEET_LNO_max += 1
                        rowPOTPACK3_new.Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_max
                        dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_new)
                    End If
                Next

                Sort_grdColumns(grdPOTPACK3, "STYLE_CODE,COLOR_CODE", True)

            Case "Copy Value to All Lines"

                If grd.ActiveRow Is Nothing Or grd.ActiveCell Is Nothing Then
                Else

                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        grow.Cells(grd.ActiveCell.Column.Key).Value = grd.ActiveCell.Value
                        grow.Update()
                    Next
                End If

                'Case "Item Status Inquiry"
                '    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Text
                '    Dim rowSPTAVEH1 As DataRow = LookUp("SPTAVEH1", VEND_CODE)
                '    If rowSPTAVEH1 IsNot Nothing Then
                '        Context_Launch("View", VEND_CODE, e.Tool.Key, "ICFSTAT1")
                '    End If

            Case "Copy Pattern to Remaining Lines"

                Dim SIZE_QTYs As New Dictionary(Of String, Integer)
                For Each row As DataRow In dst.Tables("POTPACK3").Select("", "PACK_LIST_SHEET_LNO")
                    Dim SIZE_CODE As String = row.Item("SIZE_CODE")
                    Dim CARTON_PACK As Integer = Val(row.Item("CARTON_PACK") & "")
                    If SIZE_CODE <> "" Then
                        If SIZE_QTYs.ContainsKey(SIZE_CODE) Then
                            ' row.Item("CARTON_PACK") = SIZE_QTYs(SIZE_CODE)
                        Else
                            SIZE_QTYs.Add(SIZE_CODE, CARTON_PACK)
                        End If
                    End If
                Next

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    Dim SIZE_CODE As String = grow.Cells("SIZE_CODE").Value
                    grow.Cells("CARTON_PACK").Value = SIZE_QTYs(SIZE_CODE)
                    grow.Update()
                Next

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "PACK_LIST_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If

        End Select

    End Sub

    Public Overrides Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_Leave(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PO_REFERENCE"
                Absx1.txtFor("PO_REFERENCE").Text = Absx1.txtFor("PO_REFERENCE").Text.ToUpper
            Case "STYLE_CODE_PFX"
                Absx1.txtFor("STYLE_CODE_PFX").Text = Absx1.txtFor("STYLE_CODE_PFX").Text.ToUpper
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
                Fill_Records("POTORDRR", VEND_CODE)
                Sort_grdColumns(grdPOTORDRR, "PO_DATE_SHIP_BY")

                'Case "PO_REFERENCE"
                '    Absx1.txtFor("PO_REFERENCE").Text = Absx1.txtFor("PO_REFERENCE").Text.ToUpper
                'Case "STYLE_CODE_PFX"
                '    Absx1.txtFor("STYLE_CODE_PFX").Text = Absx1.txtFor("STYLE_CODE_PFX").Text.ToUpper
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "PACK_LIST_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LC_AMT"
                If ScreenMode Then Display_Totals()
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            'Case "APPR_STATUS_CODE"
            '    If Absx1.optFor("APPR_STATUS_CODE").Value = "X" Then
            '        Absx1.optFor("STATUS_CODE").Value = "C"
            '    Else

            '    End If

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "LC_DATE"
            '    If Absx1.dteFor("LC_DATE").Value & "" = "" Then
            '        Absx1.txtFor("OPS_YYYYWW").Text = ""
            '    Else
            '        Dim DATE_START As Date = Absx1.dteFor("LC_DATE").Value
            '        If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
            '            ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(DATE_START, "dd-MMM-yyyy") & "'"
            '            Dim YW As String = ASCDATA1.GetDataValue
            '            If YW <> "" Then
            '                Absx1.txtFor("OPS_YYYYWW").Text = YW
            '            End If
            '        End If
            '    End If
        End Select
    End Sub
#End Region

#Region "grdPOTLTRCP"

#End Region

    Private Sub grdSPTSFOCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTPACKX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("PACK_LIST_NO").Text = e.Row.Cells("PACK_LIST_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        EnforceConstraints(False)
        If optShow.Value = "O" Then
            ASCMAIN1.sql = sqlPOTPACKX & " and STATUS_CODE = 'O'"
            Fill_Records("POTPACKX", "", True, ASCMAIN1.sql)
            grdPOTPACKX.Text = "Open"
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlPOTPACKX
            Fill_Records("POTPACKX", "", True, ASCMAIN1.sql)
            grdPOTPACKX.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdPOTPACKX, "PACK_LIST_NO".ToLower)

    End Sub

    Private Sub optShow_ValueChanged(sender As Object, e As EventArgs) Handles optShow.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub

    Private Sub optSTATUS_CODE_ValueChanged(sender As Object, e As EventArgs)
        If ScreenMode Then
            Synch_TABLE_NAME("POTPACK1")
            Display_Totals()
        End If
    End Sub

    Sub Display_Totals()
        'Dim LC_OPEN_CALC As Decimal = 0
        'Dim LC_CANC_CALC As Decimal = 0
        'Dim LC_AMT As Decimal = Val(Absx1.numFor("LC_AMT").Value & "")
        'Dim LC_PMTS As Decimal = Val(Absx1.numFor("LC_PMTS").Value & "")
        'If optSTATUS_CODE.Value = "O" Then
        '    LC_OPEN_CALC = LC_AMT - LC_PMTS
        '    LC_CANC_CALC = 0
        'Else
        '    LC_CANC_CALC = LC_AMT - LC_PMTS
        '    LC_OPEN_CALC = 0
        'End If

        'rowPOTPACK1.Item("LC_OPEN_CALC") = LC_OPEN_CALC
        'rowPOTPACK1.Item("LC_CANC_CALC") = LC_CANC_CALC

        Display_Totals_PO()
    End Sub

    Private Sub grdPOTLTRCP_AfterRowUpdate(sender As Object, e As RowEventArgs)
        Display_Totals_PO()
    End Sub

    Sub Display_Totals_PO()

        'Dim LC_PO As Decimal =
        '    Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_OPN)", "SEL='1'") & "") +
        '    Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_SHP)", "SEL='1'") & "")
        ''Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_REC)", "SEL='1'") & "")

        'rowPOTPACK1.Item("LC_PO") = LC_PO

        'Dim LC_AMT As Decimal = Val(Absx1.numFor("LC_AMT").Value & "")

        'If LC_PO > LC_AMT Then
        '    Absx1.numFor("LC_PO").Appearance.ForeColor = Drawing.Color.Red
        'Else
        '    Absx1.numFor("LC_PO").Appearance.ForeColor = Drawing.Color.Empty
        'End If
    End Sub

    Private Sub grdPOTORDRR_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTORDRR.AfterRowActivate
        Setup_grdPOTORDRD()
    End Sub

    Sub Setup_grdPOTORDRD()
        If grdPOTORDRR.ActiveRow Is Nothing OrElse Not grdPOTORDRR.ActiveRow.IsDataRow Then
            grdPOTORDRD.Visible = False
        Else
            Dim PO_ORDER_NO As String = grdPOTORDRR.ActiveRow.Cells("PO_ORDER_NO").Value
            Fill_Records("POTORDRD", PO_ORDER_NO)
            Sort_grdColumns(grdPOTORDRD, "STYLE_CODE, COLOR_CODE")
            grdPOTORDRD.Visible = True
        End If

    End Sub

    Private Sub grdPOTORDRR_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTORDRR.DoubleClickRow
        If grdPOTORDRR.ActiveRow IsNot Nothing AndAlso grdPOTORDRR.ActiveRow.IsDataRow Then

            Absx1.txtFor("PO_REFERENCE").Text = grdPOTORDRR.ActiveRow.Cells("PO_REFERENCE").Text
            Absx1.txtFor("PACK_LIST_DESC").Text = grdPOTORDRR.ActiveRow.Cells("PO_SPEC_ORDR_NO").Text
        End If
    End Sub

    Private Sub grdPOTPACK2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTPACK2.AfterRowActivate
        Setup_grdPOTPACK3()
    End Sub

    Sub Setup_grdPOTPACK3()
        If grdPOTPACK2.ActiveRow Is Nothing OrElse Not grdPOTPACK2.ActiveRow.IsDataRow Then
            grdPOTPACK3.Visible = False
        Else
            Dim PACK_LIST_SHEET_NO As Int32 = Val(grdPOTPACK2.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")

            Dim dvw As DataView = DirectCast(grdPOTPACK3.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PACK_LIST_SHEET_NO = " & CStr(PACK_LIST_SHEET_NO)

            'Fill_Records("POTPACK3", New Object() {"", 0, 0})

            'Sort_grdColumns(grdPOTPACK3, "PACK_LIST_SHEET_LNO")
            Sort_grdColumns(grdPOTPACK3, "STYLE_CODE,COLOR_CODE", True)
            grdPOTPACK3.Visible = True

            grdPOTPACK3.Text = "Packing List Sheet Contents for Sheet " & grdPOTPACK2.ActiveRow.Cells("PACK_LIST_SHEET_NAME").Value
        End If
    End Sub

    Private Sub grdPOTPACK3_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTPACK3.InitializeRow
        Dim CARTON_GRS_WGT As Decimal = Val(e.Row.Cells("CARTON_GRS_WGT").Value & "")
        Dim CARTON_NET_WGT As Decimal = Val(e.Row.Cells.Item("CARTON_NET_WGT").Value & "")

        Dim CARTON_DIMENSIONS As String = e.Row.Cells.Item("CARTON_DIMENSIONS").Value & ""

        With e.Row.Cells("CARTON_COUNT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Count must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        'With e.Row.Cells("STYLE_WEIGHT")
        '    If Val(.Value & "") < 0 Then
        '        .ToolTipText = "Carton Pack must be > 0"
        '        .Appearance = Appearance_Red
        '    Else
        '        .ToolTipText = ""
        '        .Appearance = Nothing
        '    End If
        'End With

        With e.Row.Cells("CARTON_PACK")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Pack must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_DIMENSIONS")
            If .Value & "" = "" Then
                .ToolTipText = "Carton Dimensions are Mandatory"
                .Appearance = Appearance_Red
            Else
                Dim CARTON_VOLUME As Decimal = Get_Volume_from_Dims(CARTON_DIMENSIONS)
                If CARTON_VOLUME <= 0 Then
                    .ToolTipText = "Carton Dimensions must be expressed as: " & Replace("L' x W' x H'", "'", Chr(34))
                    .Appearance = Appearance_Red
                Else
                    .ToolTipText = ""
                    .Appearance = Nothing
                End If

            End If
        End With

        With e.Row.Cells("CARTON_GRS_WGT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Gross Weight must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_NET_WGT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Net Weight must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_GRS_WGT")
            If CARTON_GRS_WGT > 0 And CARTON_GRS_WGT < CARTON_NET_WGT Then
                .ToolTipText = "Carton Gross Weight must be > Net Weight"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

    End Sub


    Function Get_Volume_from_Dims(CARTON_DIMENSIONS As String) As Decimal
        Dim CARTON_VOLUME As Decimal = 0
        Dim D() As String = Split(Replace(CARTON_DIMENSIONS, Chr(34), "").ToUpper, "X")
        For I As Integer = 1 To D.Length
            If Val(D(I - 1)) <> 0 Then
                If CARTON_VOLUME = 0 Then CARTON_VOLUME = 1
                CARTON_VOLUME *= Val(D(I - 1))
            End If
        Next

        Return CARTON_VOLUME
    End Function


    Sub Export_XLS()

        Generate_Carton_Nos()

        Dim VBKG_NO As String = "000001"

        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        workbook = Produce_XLS(Me, VBKG_NO)

        Dim XLS_FILENAME_base As String = "Packing List " & PACK_LIST_NO & " for Booking " & VBKG_NO
        Dim XLS_FILENAME As String = XLS_FILENAME_base & ".xlsx"
        Dim retryCount As Integer = 0
        Do Until retryCount = -1 Or retryCount > 5
            If retryCount > 0 Then
                XLS_FILENAME = XLS_FILENAME_base & "_" & CStr(retryCount) & ".xlsx"
            End If
            Try
                workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                workbook.Close()
                retryCount = -1
            Catch ex As Exception
                retryCount += 1
                If retryCount > 5 Then
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Failed to Save Workbook")
                End If
            End Try
        Loop

        If retryCount = -1 Then
            Show_Document(XLS_FILENAME)
        End If
    End Sub


    Public Function Produce_XLS(frmASFBASE0 As ASFBASE0, VAN_REF As String) As SpreadsheetGear.IWorkbook

        Dim workbook As SpreadsheetGear.IWorkbook
        Dim worksheet As SpreadsheetGear.IWorksheet
        Dim worksheetBase As SpreadsheetGear.IWorksheet

        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePasteTo As SpreadsheetGear.IRange = Nothing

        Dim FILENAME As String = ASCMAIN1.Folders("Work") & "\" & "Template.xlsx"
        workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        worksheetBase = workbook.Worksheets(0)

        Dim ETD As Date = CDate("03/04/2021")
        Dim ETA As Date = CDate("05/22/2021")
        Dim INV_NO As String = "ILBD/YK/132/2021"
        Dim INV_DATE As Date = Now.Date
        Dim COUNTRY As String = "BANGLADESH"
        Dim SHIP_BY As String = "SEA"
        Dim PORT_DESC_ORIG As String = "CHITTAGONG,BANGLADESH"
        Dim PORT_DESC_DEST As String = "MAHER TERMINAL,U.S.A."

        Dim CONTAINER_NO As String = "INTEX009/2021"
        Dim EXP_NO As String = "2656 001589 2021"
        Dim ETD_CTG As String = "ETD_CTG"
        Dim BOL_NO As String = "BOL_NO"

        For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
            'worksheet = workbook.Worksheets.Add
            worksheet = worksheetBase.CopyAfter(worksheetBase)
            worksheet.Name = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME")

            worksheet.Cells(4, 16).Value = INV_NO
            worksheet.Cells(5, 16).Value = INV_DATE
            worksheet.Cells(6, 16).Value = COUNTRY
            worksheet.Cells(7, 16).Value = SHIP_BY
            worksheet.Cells(8, 16).Value = PORT_DESC_ORIG
            worksheet.Cells(9, 16).Value = PORT_DESC_DEST

            Dim CX As Integer = 0

            CX = 13
            worksheet.Cells(4, 13).Value = "'" & Format(ETD, "MM/dd/yyyy")
            worksheet.Cells(5, 13).Value = "'" & Format(ETA, "MM/dd/yyyy")

            worksheet.Cells(4, 9).Value = CONTAINER_NO
            worksheet.Cells(5, 9).Value = EXP_NO
            worksheet.Cells(6, 9).Value = STYLE_CODE_PFX
            worksheet.Cells(7, 9).Value = PO_REFERENCE
            worksheet.Cells(8, 9).Value = ETD_CTG
            worksheet.Cells(9, 9).Value = BOL_NO

            Dim RX As Integer = 0

            Dim COLOR_CODE As String = rowPOTPACK2.Item("COLOR_CODE")
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            Dim COLOR_DESC_and_CODE As String = rowICTCOLR1.Item("COLOR_DESC") & " (" & COLOR_CODE & ")"
            worksheet.Cells(15, 5).Value = COLOR_DESC_and_CODE

            Dim PACK_LIST_DETAILS As String = rowPOTPACK2.Item("PACK_LIST_DETAILS")
            worksheet.Cells(22, 0).Value = PACK_LIST_DETAILS
            'worksheet.Cells(22, 0).WrapText = False

            Dim PACK_LIST_SHEET_NO As Integer = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3") _
                .Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO") ' rowPOTPACK2.GetChildRows("POTPACK2_POTPACK3")

                If RX > 0 Then
                    worksheet.Cells(15 + RX, 0).EntireRow.Insert()
                    worksheet.Cells(15 + RX + 1, 0).EntireRow.Copy(worksheet.Cells(15 + RX, 0).EntireRow)

                End If

                Dim STYLE_CODE As String = rowPOTPACK3.Item("STYLE_CODE") & ""
                Dim SIZE_CODE As String = rowPOTPACK3.Item("SIZE_CODE") & ""
                Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                Dim CARTON_PACK As Int32 = Val(rowPOTPACK3.Item("CARTON_PACK") & "")
                Dim CARTON_NO_START As Int32 = Val(rowPOTPACK3.Item("CARTON_NO_START") & "")
                Dim CARTON_NO_END As Int32 = Val(rowPOTPACK3.Item("CARTON_NO_END") & "")

                Dim CARTON_GRS_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_GRS_WGT") & "")
                Dim CARTON_NET_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_NET_WGT") & "")

                Dim CARTON_ID As Int32 = Val(rowPOTPACK3.Item("CARTON_ID") & "")
                Dim CARTON_DIMENSIONS As String = rowPOTPACK3.Item("CARTON_DIMENSIONS") & ""
                Dim BARCODE_START As String = rowPOTPACK3.Item("BARCODE_START") & ""
                Dim BARCODE_END As String = rowPOTPACK3.Item("BARCODE_END") & ""

                worksheet.Cells(15 + RX, 0).Value = CARTON_NO_START
                '  worksheet.Cells(15 + RX, 2).Value = CARTON_NO_END

                worksheet.Cells(15 + RX, 3).Value = STYLE_CODE
                worksheet.Cells(15 + RX, 4).Value = PO_REFERENCE

                ' STYLE DESC
                ' STYLE WEIGHT

                worksheet.Cells(15 + RX, 6).Value = SIZE_CODE
                worksheet.Cells(15 + RX, 7).Value = CARTON_COUNT
                worksheet.Cells(15 + RX, 8).Value = CARTON_PACK

                worksheet.Cells(15 + RX, 13).Value = CARTON_GRS_WGT
                worksheet.Cells(15 + RX, 14).Value = CARTON_NET_WGT

                worksheet.Cells(15 + RX, 15).Value = CARTON_DIMENSIONS
                worksheet.Cells(15 + RX, 16).Value = BARCODE_START
                worksheet.Cells(15 + RX, 17).Value = BARCODE_END
                RX += 1
            Next

            worksheet.Cells(15 + RX, 0).EntireRow.Delete()

            With worksheet.Cells(15, 5, 15 + RX - 1, 5)
                .Merge()
            End With



            With worksheet.PageSetup
                .FitToPagesTall = 1
                .FitToPagesWide = 1
                .FitToPages = True
                .Orientation = SpreadsheetGear.PageOrientation.Landscape
            End With
        Next

        worksheetBase.Delete()


        Return workbook

    End Function

    Function Generate_Carton_Nos()

        Dim EMsg As String = ""

        Dim CARTONs As New List(Of Integer)
        For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("")
            Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") & ""
            Dim PACK_LIST_SHEET_NO As Int32 = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")
            Dim CARTON_NO_START As Int32 = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")

            If CARTON_NO_START <= 0 Then
                EMsg &= vbCr & "Invalid Starting Carton No on Sheet " & PACK_LIST_SHEET_NAME
            Else
                Dim CARTON_NO As Int32 = CARTON_NO_START
                For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO") ' In rowPOTPACK2.GetChildRows("POTPACK2_POTPACK3")
                    ' Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK3.GetParentRow("POTPACK2_POTPACK3").Item("PACK_LIST_SHEET_NAME")
                    Dim CARTON_GRS_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_GRS_WGT") & "")
                    Dim CARTON_NET_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_NET_WGT") & "")
                    Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                    Dim CARTON_PACK As Int32 = Val(rowPOTPACK3.Item("CARTON_PACK") & "")
                    Dim CARTON_ID As Int32 = Val(rowPOTPACK3.Item("CARTON_ID") & "")
                    Dim CARTON_DIMENSIONS As String = rowPOTPACK3.Item("CARTON_ID") & ""

                    'Dim STYLE_WEIGHT As Decimal = Val(rowPOTPACK3.Item("STYLE_WEIGHT") & "")
                    'If STYLE_WEIGHT < 0 Then
                    '    EMsg &= vbCr & "Style Weight must not be negative - see Sheet " & PACK_LIST_SHEET_NAME
                    'End If

                    rowPOTPACK3.Item("CARTON_NO_START") = CARTON_NO

                    If CARTON_COUNT <= 0 Or CARTON_PACK <= 0 Or CARTON_DIMENSIONS = "" Or CARTON_GRS_WGT <= 0 Or CARTON_NET_WGT <= 0 Or CARTON_GRS_WGT < CARTON_NET_WGT Then
                        EMsg &= vbCr & "Issue with Data on Sheet " & PACK_LIST_SHEET_NAME
                        Exit For
                    Else
                        If CARTON_COUNT > 0 Then
                            Dim overlapping As Boolean = False
                            For I As Integer = CARTON_NO To CARTON_NO + CARTON_COUNT - 1
                                If CARTONs.Contains(I) Then
                                    overlapping = True
                                Else
                                    CARTONs.Add(I)
                                End If
                            Next

                            If overlapping Then
                                EMsg &= vbCr & "Overlapping Carton Nos on Sheet " & PACK_LIST_SHEET_NAME
                                Exit For
                            Else
                                CARTON_NO += CARTON_COUNT
                            End If
                        End If
                    End If
                Next
            End If
        Next

        Return EMsg

    End Function

    Sub Generate_LPN_Report_File()

        dst.Tables("POTLPNL1").Rows.Clear()

        For Each row As DataRow In dst.Tables("POTPACK3").Select("ISNULL(BARCODE_START,'') <> ''")
            Dim CARTON_COUNT As Integer = Val(row.Item("CARTON_COUNT") & "")
            Dim BARCODE_START As String = row.Item("BARCODE_START") & ""
            Dim BARCODE_START_NO = Val(Mid(BARCODE_START, 2))
            Dim PACK_LIST_DESC As String = rowPOTPACK1.Item("PACK_LIST_DESC")
            For C As Integer = 1 To CARTON_COUNT
                Dim BARCODE_NO As Integer = BARCODE_START_NO + C - 1
                Dim BARCODE As String = Mid(BARCODE_START, 1, 1) & Format(BARCODE_NO, "0000000")

                Dim rowPOTLPNL1 As DataRow = dst.Tables("POTLPNL1").NewRow
                With rowPOTLPNL1
                    .Item("BARCODE") = BARCODE
                    .Item("PO_REFERENCE") = PO_REFERENCE
                    .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                    .Item("PACK_LIST_DESC") = PACK_LIST_DESC
                    .Item("PACK_LIST_NO") = row.Item("PACK_LIST_NO")
                    .Item("PACK_LIST_SHEET_NO") = row.Item("PACK_LIST_SHEET_NO")
                    .Item("PACK_LIST_SHEET_LNO") = row.Item("PACK_LIST_SHEET_LNO")
                    .Item("BARCODE_STATUS") = "A"
                End With
                dst.Tables("POTLPNL1").Rows.Add(rowPOTLPNL1)
            Next
        Next

        BeginTrans()

        ASCMAIN1.sql = $"Update WHTLPNL1 Set BARCODE_STATUS = 'D' where PACK_LIST_NO = '{PACK_LIST_NO}'"
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("POTLPNL1")

        CommitTrans()

    End Sub

    Sub Print_Labels()
        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("PORLPNL1")
        Print_Report_End()
    End Sub

    Private Sub grdPOTPACK3_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTPACK3.InitializeLayout

    End Sub

    Private Sub grdPOTPACK3_AfterExitEditMode(sender As Object, e As EventArgs) Handles grdPOTPACK3.AfterExitEditMode

    End Sub

    Private Sub grdPOTPACK3_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdPOTPACK3.AfterCellUpdate

        Select Case e.Cell.Column.Key
            'Case "STYLE_WEIGHT"
            '    Calculate_Net_Weight(e)
            'Case "CARTON_PACK"
            '    Calculate_Net_Weight(e)
        End Select
    End Sub

    'Sub Calculate_Net_Weight(e As CellEventArgs)

    '    If Not Me.IsLoading And ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
    '        Dim STYLE_WEIGHT As Decimal = Val(e.Cell.Row.Cells("STYLE_WEIGHT").Value & "")
    '        Dim CARTON_PACK As Integer = Val(e.Cell.Row.Cells("CARTON_PACK").Value & "")
    '        If STYLE_WEIGHT > 0 And CARTON_PACK > 0 Then
    '            e.Cell.Row.Cells("CARTON_NET_WGT").Value = STYLE_WEIGHT * CARTON_PACK
    '        End If
    '    End If
    'End Sub

    Sub Check_for_Overbooked()

        ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", SUM (PACKED_NOW) PACKED_NOW" & vbCrLf _
            & ", SUM (PACKED_OTHER) PACKED_OTHER" & vbCrLf _
            & ", SUM (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", SUM (PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & ", SUM (POTPACK3.CARTON_COUNT * POTPACK3.CARTON_PACK) PACKED_NOW, 0 PACKED_OTHER, 0 PO_QTY_OPN, 0 PO_QTY_SHP" & vbCrLf _
            & " from POTPACK3,POTPACK1" & vbCrLf _
            & $" where POTPACK1.PACK_LIST_NO = POTPACK3.PACK_LIST_NO AND POTPACK1.PO_ORDER_NO = '{PO_ORDER_NO}'" & vbCrLf _
            & $"   and POTPACK3.PACK_LIST_NO = '{PACK_LIST_NO}' and POTPACK1.PACK_LIST_STATUS = 'F'" & vbCrLf _
            & " group by POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & ", 0 PACKED_NOW, SUM (POTPACK3.CARTON_COUNT * POTPACK3.CARTON_PACK) PACKED_OTHER, 0 PO_QTY_OPN, 0 PO_QTY_SHP" & vbCrLf _
            & " from POTPACK3,POTPACK1" & vbCrLf _
            & $" where POTPACK1.PACK_LIST_NO = POTPACK3.PACK_LIST_NO AND POTPACK1.PO_ORDER_NO = '{PO_ORDER_NO}'" & vbCrLf _
            & $"   and POTPACK3.PACK_LIST_NO <> '{PACK_LIST_NO}' and POTPACK1.PACK_LIST_STATUS = 'F'" & vbCrLf _
            & " group by POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ", 0 PACKED_NOW, 0 PACKED_OTHER, SUM (POTORDR2.PO_QTY_OPN) PO_QTY_OPN, SUM (POTORDR2.PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & " from POTORDR2" & vbCrLf _
            & $" where POTORDR2.PO_ORDER_NO = '{PO_ORDER_NO}'" & vbCrLf _
            & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ") group by STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ") where PACKED_NOW + PACKED_OTHER > PO_QTY_OPN + PO_QTY_SHP"

        Dim tbl As DataTable = ASCDATA1.GetDataTable
        If tbl.Rows.Count > 0 Then
            Using f As New ASFMSGBF
                f.Show_grd(tbl, Me, "Overbooked PO - Message to Don")
            End Using
        End If
    End Sub

    Private Sub grdPOTPACK2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdPOTPACK2.KeyPress

    End Sub

    Private Sub grdPOTPACK2_CellChange(sender As Object, e As CellEventArgs) Handles grdPOTPACK2.CellChange

        If e.Cell.Column.Key = "PACK_LIST_DETAILS" Then
            Dim PACK_LIST_DETAILS As String = e.Cell.Text & ""
            If PACK_LIST_DETAILS <> "" Then

                Dim LINES As Integer = PACK_LIST_DETAILS.Count(Function(c As Char) c = vbCr)

                grdPOTPACK2.ActiveRow.Height = 17 * (LINES + 1)
            End If

        End If
    End Sub

    Private Sub grdPOTPACK2_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdPOTPACK2.AfterRowUpdate
        e.Row.PerformAutoSize()
    End Sub

    Private Sub grdPOTPACK2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdPOTPACK2.AfterCellUpdate
        If e.Cell.Column.Key = "PACK_LIST_DETAILS" Then
            e.Cell.Row.PerformAutoSize()
        End If
    End Sub

    Private Sub grdPOTPACK3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTPACK3.AfterRowActivate
        Setup_grdPOTLPNL1()
    End Sub

    Sub Setup_grdPOTLPNL1()
        If grdPOTPACK3.ActiveRow Is Nothing OrElse Not grdPOTPACK3.ActiveRow.IsDataRow Then
            grdPOTLPNL1.Visible = False
        Else
            Dim PACK_LIST_SHEET_NO As Int32 = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")
            Dim PACK_LIST_SHEET_LNO As Int32 = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_LNO").Value & "")
            Dim dvw As DataView = DirectCast(grdPOTLPNL1.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"PACK_LIST_SHEET_NO = {PACK_LIST_SHEET_NO} and PACK_LIST_SHEET_LNO = {PACK_LIST_SHEET_LNO}"
            Sort_grdColumns(grdPOTLPNL1, "BARCODE", True)
            grdPOTLPNL1.Visible = True

            grdPOTLPNL1.Text = $"LPNs for Line {PACK_LIST_SHEET_LNO}"
        End If
    End Sub
End Class