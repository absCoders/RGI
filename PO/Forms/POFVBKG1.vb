Imports Infragistics.Win.UltraWinGrid

Public Class POFVBKG1


    'TALK ABOUT ASSIGNING SEQ In ICTSTYC1
    'AUTO SIZE ROW
    ' COPY & PASTE


    Dim rowPOTVBKG1 As DataRow
    Dim VBKG_NO As String
    Dim VBKG_NO_new As String
    Dim PACK_LIST_STATUS As String

    Dim rowTATUSER1 As DataRow

    Dim sqlPOTVBKG1 As String
    Dim VEND_CODE As String = ""
    Dim VEND_CODE_USER As String = ""

    Dim VBKG_REFERENCE_NO As String = ""
    Dim VBKG_STATUS As String = ""
    Dim VBKG_SHIP_BY As String = ""
    Dim VBKG_BOL_NO As String = ""
    Dim PORT_CODE_ORIG As String = ""
    Dim PORT_CODE_DEST As String = ""
    Dim PO_SPEC_ORDR_NO As String = ""
    Dim PO_REFERENCE As String = ""
    Dim PO_ORDER_NO As String = ""
    Dim STYLE_CODE_PFX As String = ""


    Dim Appearance_Red As New Infragistics.Win.Appearance

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Appearance_Red.ForeColor = Drawing.Color.Red

        'If MENU_ITEM_OBJECT = "POTLTRCI" Then
        '    InquiryMode = True
        'End If

        rowTATUSER1 = Lookup("TATUSER1", ASCMAIN1.USER_ID)
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
            sqlPOTVBKG1 = "Select POTVBKG1.*,APTVEND1.VEND_NAME" & vbCrLf _
                & " from POTVBKG1,APTVEND1" & vbCrLf _
                & " where APTVEND1.VEND_CODE = POTVBKG1.VEND_CODE"
            ASCMAIN1.sql = sqlPOTVBKG1 ' & "  and POTPACK1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "POTVBKG1", "**", 0, False, "")

            '       Create_TDA(.Tables.Add, "POTVBKG1", "*")

            Create_TDA(.Tables.Add, "POTVBKG2", "*", 1)


            With .Tables("POTVBKG2")
                '.Columns.Add("COLOR_DESC")
                '.Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "SUM(CHILD.CARTON_COUNT)")
                '.Columns.Add("TOTAL_UNITS", GetType(System.Decimal), "SUM(CHILD.TOTAL_UNITS)")
                '.Columns.Add("TOTAL_GRS_WGT", GetType(System.Decimal), "SUM(CHILD.TOTAL_GRS_WGT)")
                '.Columns.Add("TOTAL_NET_WGT", GetType(System.Decimal), "SUM(CHILD.TOTAL_NET_WGT)")
            End With
            'ASCMAIN1.sql = "Select APTINVH1.*" & vbCrLf _
            '    & " from APTINVH1" & vbCrLf _
            '    & " where APTINVH1.VOUCHER_NO = :PARM1"
            'Create_TDA(.Tables.Add, "APTINVH1", "**", 0, False, "V")

            'ASCMAIN1.sql = "Select * from POTORDR1 where PO_REFERENCE = :PARM1"
            'Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "V")

            'Create_TDA(.Tables.Add, "POTORDR2", "*", 1, False)

            Create_TDA(.Tables.Add, "WHTSCSEQ", "*", 0, False)
            Fill_Records("WHTSCSEQ")
        End With

        'grdAPTINVH1.DataSource = dst.Tables("APTINVH1")
        grdPOTVBKG1.DataSource = dst.Tables("POTVBKG1")

        grdPOTVBKG2.DataSource = dst.Tables("POTVBKG2")

        Create_Summary(grdPOTVBKG1, "VBKG_NO", "Count")
        ' Create_Summary(grdPOTPACKX, New String() {"LC_AMT", "LC_PMTS", "LC_FEES", "LC_OPEN"})

        Create_Summary(grdPOTVBKG2, "VBKG_NO", "Count")
        '   Create_Summary(grdPOTVBKG2, New String() {"TOTAL_CARTONS", "TOTAL_UNITS", "TOTAL_GRS_WGT", "TOTAL_NET_WGT"})




        With grdPOTVBKG1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"VBKG_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    'ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                    '    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("VBKG_NO").Header.Fixed = True
        End With

        With grdPOTVBKG2.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                GCOL.CellActivation = Activation.NoEdit
                'If New String() {"PACK_LIST_DETAILS", "CARTON_NO_START"}.Contains(GCOL.Key) Then
                '    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                '    GCOL.CellAppearance.BackColor = System.Drawing.Color.LightGreen
                '    GCOL.CellActivation = Activation.AllowEdit
                'ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                '    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                'Else
                '    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                'End If
            Next

        End With


        '  ASCMAIN1.Add_Value_List(grdPOTPACKX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'POTPACK1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        'ASCMAIN1.Add_Value_List(grdPOTPACKX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'POTPACK1' and COLUMN_NAME = 'STATUS_CODE'")

        grpHeader.Visible = False

        '  Absx1.txtFor("CURR_CODE").ReadOnly = True

        Show_Filter(grdPOTVBKG1, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                VEND_CODE = ""
                If Absx1.txtFor("VEND_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Supplier Code"
                Else
                    Dim row As DataRow = Lookup("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
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

                'PO_ORDER_NO = ""
                'PO_SPEC_ORDR_NO = ""
                'If Absx1.txtFor("PO_REFERENCE").Text.Length = 0 Then
                '    EMsg &= vbCr & "You must supply a Valid PO Reference"
                'Else
                '    PO_REFERENCE = Absx1.txtFor("PO_REFERENCE").Text
                '    Fill_Records("POTORDR1", PO_REFERENCE)
                '    If dst.Tables("POTORDR1").Rows.Count > 1 Then
                '        EMsg &= vbCr & $"More than 1 Vandale PO is associated with PO Reference {PO_REFERENCE}"
                '    ElseIf dst.Tables("POTORDR1").Rows.Count = 0 Then
                '        EMsg &= vbCr & $"No record PO Reference {PO_REFERENCE}"
                '    Else
                '        Dim row As DataRow = dst.Tables("POTORDR1").Rows(0)
                '        If row.Item("VEND_CODE") & "" <> VEND_CODE Then
                '            EMsg &= vbCr & $"Invalid PO Reference {PO_REFERENCE}"
                '        ElseIf row.Item("PO_STATUS") & "" <> "O" Then
                '            EMsg &= vbCr & $"PO Reference {PO_REFERENCE} is not Open"
                '        Else
                '            PO_ORDER_NO = row.Item("PO_ORDER_NO")
                '            PO_SPEC_ORDR_NO = row.Item("PO_SPEC_ORDR_NO") & ""
                '        End If
                '    End If

                '    If Absx1.txtFor("STYLE_CODE_PFX").Text.Length = 0 Then
                '        EMsg &= vbCr & "You must enter a Style Code Prefix"
                '    Else
                '        If PO_ORDER_NO <> "" Then
                '            STYLE_CODE_PFX = Absx1.txtFor("STYLE_CODE_PFX").Text
                '            ASCMAIN1.sql = "Select Count (*) from POTORDR2 where PO_ORDER_NO = :PARM1 and STYLE_CODE like :PARM2 || '%'"
                '            Dim PO_lines As Integer = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {PO_ORDER_NO, STYLE_CODE_PFX}))
                '            If PO_lines = 0 Then
                '                EMsg &= vbCr & $"No Lines on PO {PO_REFERENCE} with Style Code Prefix {STYLE_CODE_PFX}"
                '            End If
                '        End If
                '    End If
                'End If

                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & Absx1.txtFor("VEND_CODE").Text) Then Exit Sub
                'End If

            Case "View", "Edit"
                VBKG_NO = Absx1.txtFor("VBKG_NO").Text
                If VBKG_NO = "" Then
                    EMsg &= vbCr & "You must specify an VBKG No to View"
                Else
                    Dim row As DataRow = LookUp("POTVBKG1", VBKG_NO)
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & VBKG_NO & " on File"
                    Else
                        If eItemKey = "Edit" Then
                            If EMsg = "" Then
                                If Not ASCMAIN1.Logical_Lock("POTVBKG1", VBKG_NO) Then Exit Sub
                                '   If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & row.Item("VEND_CODE")) Then Exit Sub

                            End If
                        End If
                    End If
                End If

            Case "Update"

                'If Absx1.txtFor("PACK_LIST_DESC").Text.Length = 0 Then
                '    EMsg &= vbCr & "You must supply a Packing List Description"
                'Else
                'Dim row As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                'If IsNothing(row) Then
                '    EMsg &= vbCr & "Bank Code Entered Is Not Valid"
                'Else
                '    If row.Item("BANK_STATUS").ToString <> "A" Then
                '        EMsg &= vbCr & "Bank Status Is Not Active"
                '    End If
                'End If
                'End If

                'Dim DT As Date = Absx1.dteFor("PACK_LIST_DATE").Value & ""
                'If DT & "" = "" Then
                '    EMsg &= vbCr & "Packing List Date is Mandatory"
                'Else
                '    '  TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                'End If

                If Absx1.txtFor("VEND_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Supplier Code"
                Else
                    Dim row As DataRow = Lookup("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Supplier Entered Is Not Valid"
                    Else
                        If row.Item("VEND_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Supplier Entered Is Not Active"
                        End If
                    End If
                End If

                Dim CARTONs As New List(Of Integer)
                For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select("")
                    '   Dim PACK_LIST_SHEET_NAME As String = rowPOTVBKG2.Item("PACK_LIST_SHEET_NAME") & ""
                    '   Dim CARTON_NO_START As Int32 = Val(rowPOTVBKG2.Item("CARTON_NO_START") & "")

                Next


                If EMsg = "" Then
                    'If chkFinalize.Checked Then
                    '    If MsgBox("You have chosen to Finalize this Packing List upon Update." _
                    '            & vbCrLf & vbCrLf & "Once you have Finalized, LPNs for Barcodes will be generated," _
                    '            & vbCrLf & " and you will not be able to make further changes." _
                    '            & vbCrLf & vbCrLf & "Are you sure that you want to Finalize this Packing List?",
                    '              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    '        Exit Sub
                    '    End If
                    'End If
                End If
            Case "Delete"


            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
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

                'WorkbookView1.GetLock()

                'Dim wsx As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.ActiveWorksheet

                ''Dim ws As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.Worksheets.Add()

                'Dim newSheet As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.ActiveWorksheet.CopyAfter(WorkbookView1.ActiveWorkbook.ActiveWorksheet)

                'WorkbookView1.ReleaseLock()

            Case "Print Labels"


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
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Export XLS"
                Export_XLS()
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
                        .Items("Print Labels").Visible = True
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                        .Items("Delete").Settings.Enabled = iScreenMode
                        .Items("Print Labels").Visible = False
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
                        .Items("Export XLS").Visible = True
                    Else
                        .Items("Export XLS").Visible = False
                    End If
                    .Items("Export XLS").Visible = True  ' TEMP FOR TESTING

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                        .Items("Add Sheet").Visible = True
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                        .Items("Add Sheet").Visible = False
                    End If
                End With

                .Groups("Totals").Visible = ScreenMode
                .Groups("Show").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        'chkFinalize.Visible = Not InquiryMode And (EntryMode = "N" Or EntryMode = "E")

        '   Set_Read_Only_for_ctl(Absx1.optFor("PACK_LIST_STATUS"), True)

        '  splPOTPACKX.Visible = Not ScreenMode


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

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTVBKG2}
                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name = "grdPOTVBKG2" Then
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



            Display_Totals()

        Else
            Clear_Record()
            ' grdAPTINVH1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTVBKG1", "POTVBKG2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If VEND_CODE_USER <> "" Then
            Absx1.txtFor("VEND_CODE").Text = VEND_CODE_USER
            Absx1.txtFor("VEND_CODE").ReadOnly = True
        Else
            Absx1.txtFor("VEND_CODE").Text = ""
        End If

        'chkFinalize.Checked = False

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowPOTVBKG1 = dst.Tables("POTVBKG1").NewRow
            VBKG_NO = ASCMAIN1.Next_Control_No("POTVBKG1.VBKG_NO")
            rowPOTVBKG1.Item("VBKG_NO") = VBKG_NO
            rowPOTVBKG1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowPOTVBKG1.Item("VBKG_REFERENCE_NO") = VBKG_REFERENCE_NO
            rowPOTVBKG1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTVBKG1.Item("INIT_DATE") = DATETIME_STAMP
            rowPOTVBKG1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTVBKG1.Item("VBKG_STATUS ") = "O"
            rowPOTVBKG1.Item("VBKG_REFERENCE_NO") = HFs("VBKG_REFERENCE_NO")
            rowPOTVBKG1.Item("VESSEL_NAME") = HFs("VESSEL_NAME")

            dst.Tables("POTVBKG1").Rows.Add(rowPOTVBKG1)

        Else
            rowPOTVBKG1 = Fill_Record("POTVBKG1", VBKG_NO)
            VEND_CODE = rowPOTVBKG1.Item("VEND_CODE")
            If VEND_CODE_USER <> "" And VEND_CODE <> VEND_CODE_USER Then
                MsgBox("Issue with Vendor Code", MsgBoxStyle.OkOnly, "Please Call ABS")
                Throw New Exception("Issue with Vendor Code")
            End If
            VBKG_REFERENCE_NO = rowPOTVBKG1.Item("VBKG_REFERENCE_NO")
            '  VESSEL_NAME = rowPOTVBKG1.Item("VESSEL_NAME")
            ' PO_ORDER_NO = rowPOTVBKG1.Item("PO_ORDER_NO")

            dst.AcceptChanges()
        End If

        VBKG_STATUS = rowPOTVBKG1.Item("VBKG_STATUS")

        EnforceConstraints(False)

        Fill_Records("POTVBKG2", VBKG_NO)
        ' DGJ HERE 
        If EntryMode = "N" Then
            'Dim CARTON_NO_START_ctr As Integer = 0
            ' I THINK THIS WILL WORK FOR KOHLS ONLY
            Dim PACK_LIST_SHEET_NO_ctr As Integer = 0
            'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR2"), New String() {"COLOR_CODE"}).Select("", "COLOR_CODE")
            '    Dim rowPOTVBKG2 As DataRow = dst.Tables("POTVBKG2").NewRow
            '    rowPOTVBKG2.Item("VBKG_NO") = VBKG_NO
            '    PACK_LIST_SHEET_NO_ctr += 1
            '    rowPOTVBKG2.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr
            '    rowPOTVBKG2.Item("PACK_LIST_SHEET_NAME") = PO_REFERENCE & "-" & CStr(PACK_LIST_SHEET_NO_ctr)
            '    Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            '    rowPOTVBKG2.Item("COLOR_CODE") = COLOR_CODE
            '    Dim rowICTCOLR1 As DataRow = Lookup("ICTCOLR1", COLOR_CODE)
            '    rowPOTVBKG2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
            '    dst.Tables("POTVBKG2").Rows.Add(rowPOTVBKG2)


            '    Dim PACK_LIST_SHEET_LNO_ctr As Integer = 0
            'For Each rowPOTORDRD As DataRow In dst.Tables("POTORDRD").Select($"COLOR_CODE = '{COLOR_CODE}'", "STYLE_CODE")
            '    Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").NewRow
            '    With rowPOTPACK3
            '        .Item("PACK_LIST_NO") = PACK_LIST_NO
            '        .Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr
            '        PACK_LIST_SHEET_LNO_ctr += 1
            '        .Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_ctr
            '        Dim STYLE_CODE As String = rowPOTORDRD.Item("STYLE_CODE")
            '        .Item("STYLE_CODE") = STYLE_CODE
            '        .Item("COLOR_CODE") = COLOR_CODE
            '        Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
            '        .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            '        .Item("SIZE_CODE") = rowICTSTYL1.Item("SIZE_CODE")
            '        'CARTON_NO_START_ctr += 100
            '        '.Item("CARTON_NO_START") = CARTON_NO_START_ctr

            '        Dim rowWHTSCSEQs() As DataRow = dst.Tables("WHTSCSEQ").Select($"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'")
            '        If rowWHTSCSEQs.Length = 0 Then
            '            MsgBox($"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
            '        ElseIf rowWHTSCSEQs.Length > 1 Then
            '            MsgBox($"More than 1 Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
            '        Else
            '            .Item("CARTON_ID") = rowWHTSCSEQs(0).Item("STYLE_SEQ")
            '        End If

            '    End With
            '    dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3)
            'Next

            'Next

        Else
            Fill_Records("POTVBKG2", VBKG_NO)
        End If



        'Fill_Records("APTINVH1", PACK_LIST_NO)


        'Dim FILENAME As String = ASCMAIN1.Folders("Work") & "\" & "PACKLIST.xlsx"
        'WorkbookView1.GetLock()
        'WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)

        ''workbook = WorkbookView1.ActiveWorkbook
        'worksheet = workbook.Worksheets(0)

        WorkbookView1.ReleaseLock()

        'For Each row As DataRow In dst.Tables("POTLTRCP").Select("PACK_LIST_NO = '" & PACK_LIST_NO & "'")
        '    row.Item("SEL") = "1"
        'Next

        'rowPOTPACK1.Item("LC_PMTS") = Val(dst.Tables("APTINVH1").Compute("SUM(INV_AMT)", "") & "")
        'rowPOTPACK1.Item("PYMTS") = Val(dst.Tables("APTINVH1").Compute("COUNT(VOUCHER_NO)", "") & "")
        'Synch_TABLE_NAME("POTPACK1")
        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        'If chkFinalize.Checked Then
        '    rowPOTPACK1.Item("PACK_LIST_STATUS") = "F"

        '    Dim BARCODE_PFX As String = "Y" ' NEED TO GET THIS FROM VENDOR MASTER
        '    ' AND VENDORS WITHOUT A PREFIX ARE NOT PERMITTED TO USE THIS SCREEN

        '    For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("")
        '        Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") & ""
        '        Dim CARTON_NO_START As Int32 = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")

        '        For Each rowPOTPACK3 As DataRow In rowPOTPACK2.GetChildRows("POTPACK2_POTPACK3")
        '            Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
        '            Dim BARCODE_START = ASCMAIN1.Next_Control_No("BARCODE_" & BARCODE_PFX, CARTON_COUNT)
        '            BARCODE_START = BARCODE_PFX & BARCODE_START
        '            rowPOTPACK3.Item("BARCODE_START") = BARCODE_START
        '            Dim BARCODE_END As String = Format(Val(BARCODE_START) + CARTON_COUNT - 1, "0000000")
        '            rowPOTPACK3.Item("BARCODE_END") = BARCODE_PFX & BARCODE_END
        '        Next

        '    Next
        'End If

        Dim SQLD As String = "VBKG_NO = '" & VBKG_NO & "'"
        INIT_LAST("POTVBKG1", False, , True)

        Update_Record_TDA("POTVBKG1", SQLD)
        Update_Record_TDA("POTVBKG2", SQLD)

        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        ' Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"POTVBKG1"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where VBKG_NO = '" & VBKG_NO & "'"
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
                Absx1.txtFor("VBKG_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTVBKG1"
            E.COLUMN_NAME = "VBKG_NO"
            E.CODE_VALUE = Absx1.txtFor("VBKG_NO").Text
            E.DESC_VALUE = Absx1.txtFor("VEND_CODE").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "POTVBKG1"
        E.TABLE_KEY_CAPTION = "LC Events"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("VBKG_NO").Text '  HFs("CUST_CODE")
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
        Load_Popup_Menu(grdPOTVBKG1, "SS", "Show Filter", "Show GroupBox") ', "Move to Pending", "Approve")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
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
                    tlb_pop.Tools("Add Line").SharedProps.Visible = True

                    If grd.ActiveCell IsNot Nothing AndAlso New String() {"CARTON_DIMENSIONS", "CARTON_PACK"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = True
                    Else
                        tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = True
                    End If

                Else
                    tlb_pop.Tools("Add Line").SharedProps.Visible = False
                    tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = False
                End If
        End Select

        'If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        '    e.Cancel = True
        'Else
        '    Select Case e.SourceControl.Name

        '        'Case "grdSPTSFOC9"
        '        '    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
        '        '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
        '        '        tlb_btn.SharedProps.Visible = True
        '        '    Else
        '        '        tlb_btn.SharedProps.Visible = False
        '        '    End If
        '    End Select

        'End If
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

            Case "Copy Value to All Lines"

                If grd.ActiveRow Is Nothing Or grd.ActiveCell Is Nothing Then
                Else

                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        grow.Cells(grd.ActiveCell.Column.Key).Value = grd.ActiveCell.Value
                    Next
                End If
                'Case "Item Status Inquiry"
                '    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Text
                '    Dim rowSPTAVEH1 As DataRow = LookUp("SPTAVEH1", VEND_CODE)
                '    If rowSPTAVEH1 IsNot Nothing Then
                '        Context_Launch("View", VEND_CODE, e.Tool.Key, "ICFSTAT1")
                '    End If

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
            Case "VBKG_NO"
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
            Case "VBKG_NO"
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

    Private Sub grdSPTSFOCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs)
        If e.Row.IsDataRow Then
            Absx1.txtFor("VBKG_NO").Text = e.Row.Cells("VBKG_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        EnforceConstraints(False)
        If optShow.Value = "O" Then
            ASCMAIN1.sql = sqlPOTVBKG1 & " and STATUS_CODE = 'O'"
            Fill_Records("POTVBKG1", "", True, ASCMAIN1.sql)
            grdPOTVBKG1.Text = "Open"
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlPOTVBKG1
            Fill_Records("POTVBKG1", "", True, ASCMAIN1.sql)
            grdPOTVBKG1.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdPOTVBKG1, "VBKG_NO".ToLower)
    End Sub

    Private Sub optShow_ValueChanged(sender As Object, e As EventArgs) Handles optShow.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub

    Private Sub optSTATUS_CODE_ValueChanged(sender As Object, e As EventArgs)
        If ScreenMode Then
            Synch_TABLE_NAME("POTVBKG1")
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
    Function Get_Volume_from_Dims(CARTON_DIMENSIONS As String) As Decimal
        'Dim CARTON_VOLUME As Decimal = 0
        'Dim D() As String = Split(Replace(CARTON_DIMENSIONS, Chr(34), "").ToUpper, "X")
        'For I As Integer = 1 To D.Length
        '    If Val(D(I - 1)) <> 0 Then
        '        If CARTON_VOLUME = 0 Then CARTON_VOLUME = 1
        '        CARTON_VOLUME *= Val(D(I - 1))
        '    End If
        'Next

        'Return CARTON_VOLUME
    End Function


    Sub Export_XLS()

        'Dim VBKG_NO As String = "000001"

        'Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        'workbook = Produce_XLS(Me, VBKG_NO)

        'Dim XLS_FILENAME_base As String = "Packing Lists for Booking " & VBKG_NO
        'Dim XLS_FILENAME As String = XLS_FILENAME_base & ".xlsx"
        'Dim retryCount As Integer = 0
        'Do Until retryCount = -1 Or retryCount > 5
        '    If retryCount > 0 Then
        '        XLS_FILENAME = XLS_FILENAME_base & "_" & CStr(retryCount) & ".xlsx"
        '    End If
        '    Try
        '        workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        '        workbook.Close()
        '        retryCount = -1
        '    Catch ex As Exception
        '        retryCount += 1
        '        If retryCount > 5 Then
        '            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Failed to Save Workbook")
        '        End If
        '    End Try
        'Loop

        'If retryCount = -1 Then
        '    Show_Document(XLS_FILENAME)
        'End If
    End Sub


    Public Function Produce_XLS(frmASFBASE0 As ASFBASE0, VAN_REF As String) As SpreadsheetGear.IWorkbook

        'Dim workbook As SpreadsheetGear.IWorkbook
        'Dim worksheet As SpreadsheetGear.IWorksheet
        'Dim worksheetBase As SpreadsheetGear.IWorksheet

        'Dim range As SpreadsheetGear.IRange = Nothing
        'Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        'Dim rangePasteTo As SpreadsheetGear.IRange = Nothing

        'Dim FILENAME As String = ASCMAIN1.Folders("Work") & "\" & "Template.xlsx"
        'workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        'worksheetBase = workbook.Worksheets(0)

        'Dim ETD As Date = CDate("03/04/2021")
        'Dim ETA As Date = CDate("05/22/2021")
        'Dim INV_NO As String = "ILBD/YK/132/2021"

        'For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select("", "PACK_LIST_SHEET_NO")
        '    'worksheet = workbook.Worksheets.Add
        '    worksheet = worksheetBase.CopyAfter(worksheetBase)
        '    worksheet.Name = rowPOTVBKG2.Item("PACK_LIST_SHEET_NAME")

        '    worksheet.Cells(4, 16).Value = INV_NO

        '    Dim CX As Integer = 0

        '    CX = 13
        '    worksheet.Cells(4, 13).Value = "'" & Format(ETD, "MM/dd/yyyy")
        '    worksheet.Cells(5, 13).Value = "'" & Format(ETA, "MM/dd/yyyy")

        '    worksheet.Cells(7, 9).Value = PO_REFERENCE


        '    'worksheet.Cells(3, CX + 0).Value = "PO Key"
        '    'worksheet.Cells(3, CX + 1).Value = "'" & rowpohdr.Item("POKey")

        '    Dim RX As Integer = 0

        '    Dim COLOR_CODE As String = rowPOTVBKG2.Item("COLOR_CODE")
        '    Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
        '    Dim COLOR_DESC_and_CODE As String = rowICTCOLR1.Item("COLOR_DESC") & " (" & COLOR_CODE & ")"
        '    worksheet.Cells(15, 5).Value = COLOR_DESC_and_CODE

        '    For Each rowPOTPACK3 As DataRow In rowPOTVBKG2.GetChildRows("POTVBKG2")

        '        If RX > 0 Then
        '            worksheet.Cells(15 + RX, 0).EntireRow.Insert()
        '        End If

        '        Dim STYLE_CODE As String = rowPOTPACK3.Item("STYLE_CODE") & ""
        '        Dim SIZE_CODE As String = rowPOTPACK3.Item("SIZE_CODE") & ""
        '        Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
        '        Dim CARTON_PACK As Int32 = Val(rowPOTPACK3.Item("CARTON_PACK") & "")
        '        Dim CARTON_NO_START As Int32 = Val(rowPOTPACK3.Item("CARTON_NO_START") & "")
        '        Dim CARTON_NO_END As Int32 = Val(rowPOTPACK3.Item("CARTON_NO_END") & "")

        '        Dim CARTON_GRS_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_GRS_WGT") & "")
        '        Dim CARTON_NET_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_NET_WGT") & "")

        '        Dim CARTON_ID As Int32 = Val(rowPOTPACK3.Item("CARTON_ID") & "")
        '        Dim CARTON_DIMENSIONS As String = rowPOTPACK3.Item("CARTON_DIMENSIONS") & ""
        '        Dim BARCODE_START As String = rowPOTPACK3.Item("BARCODE_START") & ""
        '        Dim BARCODE_END As String = rowPOTPACK3.Item("BARCODE_END") & ""

        '        worksheet.Cells(15 + RX, 0).Value = CARTON_NO_START
        '        worksheet.Cells(15 + RX, 2).Value = CARTON_NO_END

        '        worksheet.Cells(15 + RX, 3).Value = STYLE_CODE
        '        worksheet.Cells(15 + RX, 4).Value = PO_REFERENCE

        '        worksheet.Cells(15 + RX, 6).Value = SIZE_CODE
        '        worksheet.Cells(15 + RX, 7).Value = CARTON_COUNT
        '        worksheet.Cells(15 + RX, 8).Value = CARTON_PACK

        '        worksheet.Cells(15 + RX, 13).Value = CARTON_GRS_WGT
        '        worksheet.Cells(15 + RX, 14).Value = CARTON_NET_WGT

        '        worksheet.Cells(15 + RX, 15).Value = CARTON_DIMENSIONS
        '        worksheet.Cells(15 + RX, 16).Value = BARCODE_START
        '        worksheet.Cells(15 + RX, 17).Value = BARCODE_END
        '        RX += 1
        '    Next

        '    worksheet.Cells(15 + RX, 0).EntireRow.Delete()

        '    With worksheet.PageSetup
        '        .FitToPagesTall = 1
        '        .FitToPagesWide = 1
        '        .FitToPages = True
        '        .Orientation = SpreadsheetGear.PageOrientation.Landscape
        '    End With
        'Next

        'worksheetBase.Delete()


        'Return workbook

    End Function


End Class