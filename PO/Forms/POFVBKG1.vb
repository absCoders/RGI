Imports Infragistics.Win.UltraWinGrid

Public Class POFVBKG1



    Dim rowPOTVBKG1 As DataRow
    Dim VBKG_NO As String
    Dim VBKG_NO_new As String
    Dim PACK_LIST_STATUS As String

    Dim rowTATUSER1 As DataRow

    Dim sqlPOTVBKGX As String
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

    Dim CURR_PACK_LIST_NOS As List(Of String) = New List(Of String)

    ' Dim DEL_PACK_CODE_ALL As New List(Of String)


    Dim Appearance_Red As New Infragistics.Win.Appearance

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Appearance_Red.ForeColor = Drawing.Color.Red


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
            .Items("Delete").Visible = Not InquiryMode
        End With

        Get_PARM("GLTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")

        With dst
            sqlPOTVBKGX = "Select POTVBKG1.*,APTVEND1.VEND_NAME" & vbCrLf _
                & " from POTVBKG1,APTVEND1" & vbCrLf _
                & " where APTVEND1.VEND_CODE = POTVBKG1.VEND_CODE"
            ASCMAIN1.sql = sqlPOTVBKGX ' & "  and POTPACK1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "POTVBKGX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "POTVBKG1", "*")

            ASCMAIN1.sql = "Select POTVBKG2.*,STYLE_CODE_PFX,PO_REFERENCE,PO_ORDER_NO,PACK_LIST_STATUS,PACK_LIST_DESC,PACK_LIST_DATE,INITIAL_ORDER" _
            & " from POTVBKG2,POTPACK1 where POTPACK1.PACK_LIST_NO = POTVBKG2.PACK_LIST_NO" _
            & " AND POTVBKG2.VBKG_NO = :PARM1"
            ' Create_TDA(.Tables.Add, "POTVBKG2", "*", 1)
            Create_TDA(.Tables.Add, "POTVBKG2", "**", 0, True, "V")


            ASCMAIN1.sql = "Select * from POTPACK1 where PACK_LIST_STATUS = 'F' AND VEND_CODE = :PARM1 and VBKG_NO IS NULL"
            Create_TDA(.Tables.Add, "POTPACK1", "**", 0, False, "V")


            With .Tables("POTVBKG2")
                '.Columns.Add("PACK_LIST_DESC")
                '.Columns.Add("PACK_LIST_DATE")
                '    .Columns.Add("STYLE_CODE_PFX ")
                '    .Columns.Add("PO_REFERENCE ")
                '    .Columns.Add("PO_ORDER_NO")
                '    .Columns.Add("INITIAL_ORDER")
            End With

            Create_TDA(.Tables.Add, "WHTSCSEQ", "*", 0, False)
            Fill_Records("WHTSCSEQ")
        End With

        grdPOTVBKGX.DataSource = dst.Tables("POTVBKGX")

        grdPOTVBKG2.DataSource = dst.Tables("POTVBKG2")
        grdPOTPACK1.DataSource = dst.Tables("POTPACK1")

        Dim dvw As DataView = DirectCast(grdPOTPACK1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "VBKG_NO IS NULL"


        Create_Summary(grdPOTVBKGX, "VBKG_NO", "Count")
        ' Create_Summary(grdPOTPACKX, New String() {"LC_AMT", "LC_PMTS", "LC_FEES", "LC_OPEN"})

        Create_Summary(grdPOTVBKG2, "PACK_LIST_NO", "Count")


        With grdPOTVBKGX.DisplayLayout.Bands(0)
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

        btnShip.Visible = ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz"

        grpHeader.Visible = False

        ASCMAIN1.Add_Value_List(grdPOTVBKGX, "VBKG_STATUS", Nothing, New String() {":", "O:Open", "F:Finalized"})


        Show_Filter(grdPOTVBKGX, True)
        Refresh_Documents()
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
                Dim DT As String = Absx1.dteFor("VEND_INV_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Vendor Invoice Date is Mandatory"
                Else
                    '   TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                End If

                If Absx1.txtFor("VEND_INV_NO").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Vendor Invoice No"
                End If
                '  


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
                            Dim PO_SHIPMENT_NO As String = Absx1.txtFor("PO_SHIPMENT_NO").Text
                            If PO_SHIPMENT_NO <> "" Then
                                EMsg &= vbCr & $"Booking {VBKG_NO} has already been imported into Shipment {PO_SHIPMENT_NO} - No Edits Permitted"
                            End If
                            If EMsg = "" Then
                                If Not ASCMAIN1.Logical_Lock("POTVBKG1", VBKG_NO) Then Exit Sub
                            End If
                        End If
                    End If

                    'If EMsg = "" Then
                    '    If Not ASCMAIN1.Logical_Lock("POTVBKG1", VBKG_NO) Then Exit Sub
                    '    '   If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & row.Item("VEND_CODE")) Then Exit Sub

                    'End If
                End If


            Case "Update"


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

                If Absx1.txtFor("PORT_CODE_ORIG").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Orig Port"
                End If
                If Absx1.txtFor("PORT_CODE_DEST").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Dest Port"
                End If

                If Absx1.txtFor("VBKG_BOL_NO").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Bol No"
                End If
                If Absx1.txtFor("VBKG_REFERENCE_NO").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Reference No"
                End If
                If Absx1.txtFor("VBKG_SHIP_BY").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Ship By"
                End If
                If Absx1.txtFor("VESSEL_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Vessel"
                End If


                Dim ETADATE As String = Format(Absx1.dteFor("VBKG_ETA").Value, "yyyyMMdd")
                Dim ETDDATE As String = Format(Absx1.dteFor("VBKG_ETD").Value, "yyyyMMdd")
                If ETADATE & "" <= ETDDATE & "" Then
                    EMsg &= vbCr & "ETA Date Must be Later than the ETD Data"
                Else
                    '  TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                End If

                If ETDDATE & "" = "" Then
                    EMsg &= vbCr & "You must provide an ETD Date"
                End If



                If chkFinalize.Checked Then
                    If dst.Tables("POTVBKG2").Select("VBKG_NO = '" & VBKG_NO & "'").Length = 0 Then
                        EMsg &= vbCr & "There must be Pack Lists added when finalizing a Booking"
                    End If

                    If Absx1.txtFor("CONTAINER_NO").Text = "" Or Absx1.txtFor("CONTAINER_SEAL_NO").Text = "" Then
                        EMsg &= vbCr & "Container and Seal are mandatory when finalizing a Booking"
                    End If

                End If



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
                If MsgBox("OK to Delete Booking No " & VBKG_NO & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

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
                Refresh_Documents()

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)
                Refresh_Documents()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        'If Not tf Then
        '    Refresh_Documents()
        'End If

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
                        '      .Items("Print Labels").Visible = True
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                        .Items("Delete").Settings.Enabled = iScreenMode
                        '               .Items("Print Labels").Visible = False
                    End If

                    .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                    .Items("Done").Visible = ScreenMode And (EntryMode = "V")
                    .Items("Delete").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")

                    If ScreenMode And EntryMode = "E" AndAlso rowPOTVBKG1.Item("VBKG_STATUS") & "" <> "F" Then
                        .Items("Delete").Visible = True
                        ' .Items("Delete").Visible = False ' NOT UNTIL WE FIGURE OUT PROTECTIONS
                    Else
                        .Items("Delete").Visible = False
                    End If



                    If ScreenMode Then
                        '     .Items("Export XLS").Visible = True
                    Else
                        '    .Items("Export XLS").Visible = False
                    End If
                    ' .Items("Export XLS").Visible = True  ' TEMP FOR TESTING

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                        '      .Items("Add Sheet").Visible = True
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                        '     .Items("Add Sheet").Visible = False
                    End If
                    '   grdPOTPACK1.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                End With

                .Groups("Totals").Visible = ScreenMode
                .Groups("Show").Visible = Not ScreenMode
                grdPOTVBKGX.Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        chkFinalize.Visible = Not InquiryMode And (EntryMode = "N" Or EntryMode = "E")

        lblPO_SHIPMENT_NO.Visible = ScreenMode 'And InquiryMode
        txtPO_SHIPMENT_NO.Visible = ScreenMode 'And InquiryMode

        If ScreenMode Then


            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                    .Items("Delete").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))

            If EntryMode = "E" Or EntryMode = "N" Then
                '    Set_Read_Only_for_ctl(Absx1.txtFor("LC_REF_NO"), False)
                Set_Read_Only_for_ctl(Absx1.dteFor("VEND_INV_DATE"), False)
                Set_Read_Only_for_ctl(Absx1.txtFor("VEND_INV_NO"), False)
            End If

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTVBKG2}
                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name = "grdPOTVBKG2" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.True
                            .AllowUpdate = DefaultableBoolean.True
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.True
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

            Set_Read_Only_for_ctl(Absx1.optFor("VBKG_STATUS"), True)

            Display_Totals()

        Else
            Clear_Record()
            ' grdAPTINVH1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTVBKG1", "POTVBKG2", "POTPACK1"}
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

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowPOTVBKG1 = dst.Tables("POTVBKG1").NewRow
            VBKG_NO = ASCMAIN1.Next_Control_No("POTVBKG1.VBKG_NO")
            rowPOTVBKG1.Item("VBKG_NO") = VBKG_NO
            rowPOTVBKG1.Item("VEND_CODE") = HFs("VEND_CODE")
            '   rowPOTVBKG1.Item("VBKG_REFERENCE_NO") = HFs("VBKG_REFERENCE_NO")
            rowPOTVBKG1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTVBKG1.Item("INIT_DATE") = DATETIME_STAMP
            rowPOTVBKG1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTVBKG1.Item("VBKG_STATUS") = "O"
            rowPOTVBKG1.Item("VEND_INV_NO") = HFs("VEND_INV_NO")
            rowPOTVBKG1.Item("VEND_INV_DATE") = HFs("VEND_INV_DATE")
            '  rowPOTVBKG1.Item("VBKG_REFERENCE_NO") = HFs("VBKG_REFERENCE_NO")
            ' rowPOTVBKG1.Item("VESSEL_NAME") = HFs("VESSEL_NAME")

            dst.Tables("POTVBKG1").Rows.Add(rowPOTVBKG1)

        Else
            rowPOTVBKG1 = Fill_Record("POTVBKG1", VBKG_NO)
            Dim VEND_CODE As String = rowPOTVBKG1.Item("VEND_CODE") & ""
            If VEND_CODE_USER <> "" And VEND_CODE <> VEND_CODE_USER Then
                MsgBox("Issue with Vendor Code", MsgBoxStyle.OkOnly, "Please Call ABS")
                Throw New Exception("Issue with Vendor Code")
            End If
            VBKG_REFERENCE_NO = rowPOTVBKG1.Item("VBKG_REFERENCE_NO") & ""
            '  VESSEL_NAME = rowPOTVBKG1.Item("VESSEL_NAME")
            ' PO_ORDER_NO = rowPOTVBKG1.Item("PO_ORDER_NO")

            dst.AcceptChanges()
        End If

        VBKG_STATUS = rowPOTVBKG1.Item("VBKG_STATUS") & ""

        EnforceConstraints(False)

        Fill_Records("POTVBKG2", VBKG_NO)
        '   Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
        '     Fill_Records("POTPACK1", VEND_CODE, True)
        ' DGJ HERE 
        If EntryMode = "N" Then


        Else
            Fill_Records("POTVBKG2", VBKG_NO)
            '    Fill_Records("POTPACK1", VEND_CODE, True)

            '    DGJ
            CURR_PACK_LIST_NOS = New List(Of String)
            For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select
                Dim PACK_LIST_NO As String = rowPOTVBKG2.Item("PACK_LIST_NO") & ""
                CURR_PACK_LIST_NOS.Add(PACK_LIST_NO)
            Next
        End If

        Dim dvw As DataView = DirectCast(grdPOTPACK1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "VBKG_NO IS NULL"

        Dim vl As Infragistics.Win.ValueList

        If (Not grdPOTVBKG2.DisplayLayout.ValueLists.Exists("PACK_LIST_STATUS")) Then
            vl = grdPOTVBKG2.DisplayLayout.ValueLists.Add("PACK_LIST_STATUS")
            vl.ValueListItems.Add("O", "Open")
            vl.ValueListItems.Add("F", "Finalized")
            grdPOTVBKG2.DisplayLayout.Bands(0).Columns("PACK_LIST_STATUS").ValueList = grdPOTVBKG2.DisplayLayout.ValueLists("PACK_LIST_STATUS")
        End If

        If (Not grdPOTPACK1.DisplayLayout.ValueLists.Exists("PACK_LIST_STATUS")) Then
            vl = grdPOTPACK1.DisplayLayout.ValueLists.Add("PACK_LIST_STATUS")
            vl.ValueListItems.Add("O", "Open")
            vl.ValueListItems.Add("F", "Finalized")
            grdPOTPACK1.DisplayLayout.Bands(0).Columns("PACK_LIST_STATUS").ValueList = grdPOTPACK1.DisplayLayout.ValueLists("PACK_LIST_STATUS")
        End If

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        If chkFinalize.Checked Then
            rowPOTVBKG1.Item("VBKG_STATUS") = "F"
        End If


        For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select
            Dim PACK_LIST_NO As String = rowPOTVBKG2.Item("PACK_LIST_NO") & ""
            Dim VBKG_NO As String = rowPOTVBKG2.Item("VBKG_NO") & ""
            ASCMAIN1.sql = "Update POTPACK1 Set VBKG_NO = '" & VBKG_NO & "' where PACK_LIST_NO  = '" & PACK_LIST_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next



        For Each CURR_PACK_LIST_NO As String In CURR_PACK_LIST_NOS
            CURR_PACK_LIST_NO = CURR_PACK_LIST_NO.Trim
            Dim PACK_LIST_NO As String = CURR_PACK_LIST_NO
            '  Dim rowPOTVBKG2x As DataRow = dst.Tables("POTVBKG2").Rows.Find(New Object() {CURR_PACK_LIST_NO})

            If dst.Tables("POTVBKG2").Select("PACK_LIST_NO = '" & PACK_LIST_NO & "'").Length = 0 Then
                ASCMAIN1.sql = "Update POTPACK1 Set VBKG_NO = NULL where PACK_LIST_NO  = '" & CURR_PACK_LIST_NO & "'"
                ASCDATA1.ExecuteSQL()
            End If


        Next

        Dim SQLD As String = "VBKG_NO = '" & VBKG_NO & "'"
        INIT_LAST("POTVBKG1", False, , True)

        Update_Record_TDA("POTVBKG1", SQLD)
        Update_Record_TDA("POTVBKG2", SQLD)

        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Delete_Records()
        '  CommitTrans("Delete Complete")
        CommitTrans("Booking No " & VBKG_NO & " has been Deleted")

    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select
            Dim PACK_LIST_NO As String = rowPOTVBKG2.Item("PACK_LIST_NO") & ""
            Dim VBKG_NO As String = rowPOTVBKG2.Item("VBKG_NO") & ""
            ASCMAIN1.sql = "Update POTPACK1 Set VBKG_NO = NULL where PACK_LIST_NO  = '" & PACK_LIST_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next

        ' Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"POTVBKG1", "POTVBKG2"}
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
        Load_Popup_Menu(grdPOTVBKGX, "SS", "Show Filter", "Show GroupBox") ', "Move to Pending", "Approve")
        Load_Popup_Menu(grdPOTPACK1, "B", "Add Pack List to Booking")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Or e.SourceControl.Name = "grdPOTPACK1_EmbeddableTextBox" Then
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


            Case "grdPOTPACK1"

                If Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_pop.Tools("Add Pack List to Booking").SharedProps.Visible = True
                Else
                    tlb_pop.Tools("Add Pack List to Booking").SharedProps.Visible = False
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
            Case "Add Pack List to Booking"
                Add_Pack_List
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
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
                Fill_Records("POTPACK1", VEND_CODE, True)
                '   Sort_grdColumns(grdPOTORDRR, "PO_DATE_ETA".ToLower)

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
            ASCMAIN1.sql = sqlPOTVBKGX & " and VBKG_STATUS = 'O'"
            Fill_Records("POTVBKGX", "", True, ASCMAIN1.sql)
            grdPOTVBKGX.Text = "Open"
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlPOTVBKGX
            Fill_Records("POTVBKGX", "", True, ASCMAIN1.sql)
            grdPOTVBKGX.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdPOTVBKGX, "VBKG_NO".ToLower)
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


        Display_Totals_PO()
    End Sub

    Private Sub grdPOTLTRCP_AfterRowUpdate(sender As Object, e As RowEventArgs)
        Display_Totals_PO()
    End Sub

    Sub Display_Totals_PO()

    End Sub



    Sub Export_XLS()


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

    Private Sub grdPOTVBKGX_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTVBKGX.InitializeLayout

    End Sub

    Private Sub grdPOTVBKGX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTVBKGX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("VBKG_NO").Text = e.Row.Cells("VBKG_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdPOTVBKG2_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdPOTVBKG2.AfterRowsDeleted
        Dim DEL_PACK_CODE_ALL As New List(Of String)
        DEL_PACK_CODE_ALL = DirectCast(grdPOTVBKG2.Tag, List(Of String))
        For Each DEL_PACK_CODE_A As String In DEL_PACK_CODE_ALL
            Dim DEL_PACK_CODE As String = Split(DEL_PACK_CODE_A, vbTab)(0)
            Dim TYPE_DEL As String = Split(DEL_PACK_CODE_A, vbTab)(1)
            DEL_PACK_CODE = DEL_PACK_CODE.Trim
            If TYPE_DEL = "N" Then
                ' New to POTPACK1 dst
                ' ADD NEW POTPACK1 TO DST
                ASCMAIN1.sql = "Select * from POTPACK1" _
                & " where PACK_LIST_NO = '" & DEL_PACK_CODE & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing Then
                    Dim rowPOTPACK1 As DataRow
                    rowPOTPACK1 = dst.Tables("POTPACK1").NewRow
                    rowPOTPACK1.Item("PACK_LIST_NO") = DEL_PACK_CODE
                    rowPOTPACK1.Item("PACK_LIST_DESC") = row.Item("PACK_LIST_DESC")
                    rowPOTPACK1.Item("PACK_LIST_DATE") = row.Item("PACK_LIST_DATE")
                    rowPOTPACK1.Item("VEND_CODE") = row.Item("VEND_CODE")
                    rowPOTPACK1.Item("PACK_LIST_STATUS") = row.Item("PACK_LIST_STATUS")
                    rowPOTPACK1.Item("STYLE_CODE_PFX") = row.Item("STYLE_CODE_PFX")
                    rowPOTPACK1.Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                    rowPOTPACK1.Item("PO_REFERENCE") = row.Item("PO_REFERENCE")
                    rowPOTPACK1.Item("INIT_OPER") = row.Item("INIT_OPER")
                    rowPOTPACK1.Item("LAST_OPER") = row.Item("LAST_OPER")
                    rowPOTPACK1.Item("INIT_DATE") = row.Item("INIT_DATE")
                    rowPOTPACK1.Item("LAST_DATE") = row.Item("LAST_DATE")
                    rowPOTPACK1.Item("INITIAL_ORDER") = row.Item("INITIAL_ORDER")
                    rowPOTPACK1.Item("VBKG_NO") = DBNull.Value
                    dst.Tables("POTPACK1").Rows.Add(rowPOTPACK1)

                End If

            ElseIf TYPE_DEL = "E" Then
                Dim rowPOTPACK1 As DataRow = dst.Tables("POTPACK1").Rows.Find(New Object() {DEL_PACK_CODE})
                rowPOTPACK1.Item("VBKG_NO") = DBNull.Value
            End If

        Next

    End Sub

    Private Sub grdPOTVBKG2_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdPOTVBKG2.BeforeRowsDeleted

        Dim DEL_PACK_CODE_ALL As New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In grdPOTVBKG2.Selected.Rows
            Dim VBKG_NO As String = grow.Cells("VBKG_NO").Value
            Dim PACK_LIST_NO As String = grow.Cells("PACK_LIST_NO").Value
            If dst.Tables("POTVBKG2").Rows.Find(New String() {VBKG_NO, PACK_LIST_NO}).RowState = DataRowState.Added Then
                DEL_PACK_CODE_ALL.Add(PACK_LIST_NO & vbTab & "E")
                grdPOTVBKG2.Tag = DEL_PACK_CODE_ALL
            Else
                DEL_PACK_CODE_ALL.Add(PACK_LIST_NO & vbTab & "N")
                grdPOTVBKG2.Tag = DEL_PACK_CODE_ALL
            End If
        Next

    End Sub

    Private Sub grdPOTPACK1_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTPACK1.InitializeLayout

    End Sub

    Private Sub grdPOTPACK1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTPACK1.DoubleClickRow
        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            If Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
                Add_Pack_List()
            End If

        End If
    End Sub


    Sub Add_Pack_List()
        If grdPOTPACK1.ActiveRow.Cells("VBKG_NO").Value & "" <> "" Then
            MsgBox("Pack List already added to this Booking", MsgBoxStyle.OkOnly, "")
        Else

            If MsgBox("Are you sure that you want to Add Packing List " & grdPOTPACK1.ActiveRow.Cells("PACK_LIST_NO").Value & " to This Booking?",
                MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                Exit Sub
            End If

            Dim PACK_LIST_NO As String = grdPOTPACK1.ActiveRow.Cells("PACK_LIST_NO").Value & ""
            If Not ASCMAIN1.Logical_Lock("POTPACK1", PACK_LIST_NO) Then
                ' PROBLEM Check oracle to make sure that the VBKG_NO IS STILLBLANK
            Else
                ASCMAIN1.sql = "Select * from POTLPNL1 where PACK_LIST_NO = '" & PACK_LIST_NO & "' " _
                            & " and BARCODE_STATUS = 'A'"
                For Each rowPOTLPNL1 As DataRow In ASCDATA1.GetDataTable.Rows
                    If rowPOTLPNL1.Item("SHIP_CONF") & "" <> "S" Then
                        MsgBox("This Pack List Has Bar Codes that Have a Ship Conf Other than 'S'", MsgBoxStyle.OkOnly, "Cannot Add Pack List " & PACK_LIST_NO)
                        Exit Sub
                    End If
                Next

                ASCMAIN1.sql = "Select * from POTPACK1 where PACK_LIST_NO = '" & PACK_LIST_NO & "' " _
                    & " and VBKG_NO is Null "
                Dim tblPOTPACK1 As DataTable = ASCDATA1.GetDataTable()
                If tblPOTPACK1.Rows.Count > 0 Then
                    Dim rowPOTVBKG2_new As DataRow = dst.Tables("POTVBKG2").NewRow
                    '    rowPOTVBKG2_new.ItemArray = rowPOTVBKG2.ItemArray
                    rowPOTVBKG2_new.Item("VBKG_NO") = VBKG_NO
                    rowPOTVBKG2_new.Item("PACK_LIST_NO") = PACK_LIST_NO
                    rowPOTVBKG2_new.Item("PACK_LIST_DESC") = grdPOTPACK1.ActiveRow.Cells("PACK_LIST_DESC").Value & ""
                    rowPOTVBKG2_new.Item("PACK_LIST_DATE") = grdPOTPACK1.ActiveRow.Cells("PACK_LIST_DATE").Value & ""
                    rowPOTVBKG2_new.Item("PACK_LIST_STATUS") = grdPOTPACK1.ActiveRow.Cells("PACK_LIST_STATUS").Value & ""
                    rowPOTVBKG2_new.Item("STYLE_CODE_PFX") = grdPOTPACK1.ActiveRow.Cells("STYLE_CODE_PFX").Value & ""
                    rowPOTVBKG2_new.Item("PO_REFERENCE") = grdPOTPACK1.ActiveRow.Cells("PO_REFERENCE").Value & ""
                    rowPOTVBKG2_new.Item("PO_ORDER_NO") = grdPOTPACK1.ActiveRow.Cells("PO_ORDER_NO").Value & ""
                    rowPOTVBKG2_new.Item("INITIAL_ORDER") = grdPOTPACK1.ActiveRow.Cells("INITIAL_ORDER").Value & ""
                    dst.Tables("POTVBKG2").Rows.Add(rowPOTVBKG2_new)
                    grdPOTPACK1.ActiveRow.Cells("VBKG_NO").Value = VBKG_NO
                    grdPOTPACK1.ActiveRow.Update()
                    Sort_grdColumns(grdPOTPACK1, "PACK_LIST_NO", True)
                Else
                    MsgBox("This Pack List No is no longer available to add to Booking", MsgBoxStyle.OkOnly, "Cannot Add Pack List")
                    '       EMsg &= vbCr & " This Pack List No is no longer available to add to Booking"
                    Fill_Records("POTPACK1", VEND_CODE, True)
                End If

            End If


        End If
    End Sub

    Private Sub btnShip_Click(sender As Object, e As EventArgs) Handles btnShip.Click

        If VBKG_NO <> "" Then
            Dim PO_SHIPMENT_NO As String = ""
            PO_SHIPMENT_NO = Book2ShiP(VBKG_NO, "")
        End If

    End Sub

    Function Get_Volume_from_Dims(CARTON_DIMENSIONS As String) As Decimal ' BELONGS IN TAC - SEE POFPACK1
        Dim CARTON_VOLUME As Decimal = 0
        If CARTON_DIMENSIONS <> "" Then
            Dim D() As String = Split(Replace(CARTON_DIMENSIONS, Chr(34), "").ToUpper, "X")
            If D.Length > 0 Then
                For I As Integer = 1 To D.Length
                    If Val(D(I - 1)) <> 0 Then
                        If CARTON_VOLUME = 0 Then CARTON_VOLUME = 1
                        CARTON_VOLUME *= Val(D(I - 1))
                    End If
                Next
                If D.Length <> 3 Then CARTON_VOLUME = 0
            End If
        End If
        Return CARTON_VOLUME
    End Function

    Function Book2ShiP(VBKG_NO As String, PO_SHIPMENT_NO As String) As String

        ' THIS ROUTINE ASSUMES THAT POTVBKG1 AND POTVBKG2 EXISTS

        Dim rowPOTVBKG1 As DataRow = Fill_Record("POTVBKG1", VBKG_NO)

        If Not dst.Tables.Contains("POTSHIP1") Then
            For Each TABLE_NAME As String In New String() {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP7", "POTSHIP8", "POTPACK2", "POTPACK3", "WHTPPKM1", "WHTPPKM2"}
                Create_TDA(dst.Tables.Add, TABLE_NAME, "*", 1)
            Next
        End If

        For Each TABLE_NAME As String In New String() {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP7", "POTSHIP8", "POTPACK2", "POTPACK3", "WHTPPKM1", "WHTPPKM2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Dim rowPOTSHIP1 As DataRow = Nothing

        BeginTrans()

        If PO_SHIPMENT_NO = "" Then
            PO_SHIPMENT_NO = ASCMAIN1.Next_Control_No("PO_SHIPMENT_NO")
            rowPOTSHIP1 = dst.Tables("POTSHIP1").NewRow
            With rowPOTSHIP1
                .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIP_VESSEL") = rowPOTVBKG1.Item("VESSEL_NAME")
                .Item("PO_SHIP_ETA") = rowPOTVBKG1.Item("VBKG_ETA")
                .Item("PO_SHIP_LANDING_LEAD_DAYS") = ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETA_TO_ARR")
                .Item("PO_SHIP_REF_NO") = Val(PO_SHIPMENT_NO)
                .Item("PO_SHIP_ADV_DATE") = DATETIME_STAMP.Date
                .Item("PO_DATE_SHIPPED") = rowPOTVBKG1.Item("VBKG_ETD")
                ' .Item("PORT_CODE") = ""
                .Item("WHSE_CODE") = ROWs("POTPARM1").Item("PO_PARM_DEF_WHSE_CODE")
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("COST_IND") = "1"
                .Item("FREIGHT_ENTERED_BY") = "C"
                .Item("PO_NOTES") = "YINTAK"
                .Item("REVIEW") = "0"
                .Item("AIR_SHIP") = IIf(rowPOTVBKG1.Item("VBKG_SHIP_BY") & "" = "AIR", "1", "0")
                .Item("COST_COMPLETE") = "0"
                .Item("LP_STATUS") = "0"
                .Item("PORT_CODE_ORIG") = rowPOTVBKG1.Item("PORT_CODE_ORIG")
                .Item("PORT_CODE_DEST") = rowPOTVBKG1.Item("PORT_CODE_DEST")
                .Item("COST_FRT_METHOD") = "W"
                .Item("COST_NO_DUTY") = "0"
            End With
            dst.Tables("POTSHIP1").Rows.Add(rowPOTSHIP1)
        Else
            For Each TABLE_NAME As String In New String() {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP7", "POTSHIP8"}
                Fill_Record(TABLE_NAME, PO_SHIPMENT_NO)
            Next

            rowPOTSHIP1 = dst.Tables("POTSHIP1").Rows(0)

        End If

        Dim CONTAINER_NO As String = rowPOTVBKG1.Item("CONTAINER_NO")

        Dim PO_SHIPMENT_LNO_ctr As Integer = Val(dst.Tables("POTSHIP2").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
        Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").NewRow
        With rowPOTSHIP2
            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
            .Item("CONTAINER_NO") = CONTAINER_NO
            .Item("BOL_NO") = rowPOTVBKG1.Item("VBKG_BOL_NO")
            .Item("PO_SHIP_CTNS") = 0
            .Item("PO_SHIP_STATUS") = "O"
            '.Item("PO_SOURCE_DOC") = ""
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("CONTAINER_SIZE") = "40HC"
            .Item("COMM_INV_NO") = rowPOTVBKG1.Item("VEND_INV_NO")
            .Item("ACCRUAL_STATUS") = "0"
        End With
        dst.Tables("POTSHIP2").Rows.Add(rowPOTSHIP2)


        rowPOTVBKG1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
        rowPOTVBKG1.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr

        Fill_Records("POTVBKG2", VBKG_NO)
        Dim TOTAL_CARTONS As Integer = 0

        For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select("", "PACK_LIST_NO")
            Dim PACK_LIST_NO As String = rowPOTVBKG2.Item("PACK_LIST_NO") & ""
            Dim rowPOTPACK1 As DataRow = LookUp("POTPACK1", PACK_LIST_NO)

            Dim INITIAL_ORDER As String = rowPOTPACK1.Item("INITIAL_ORDER") & ""
            Dim CUST_CODE As String = rowPOTPACK1.Item("CUST_CODE") & ""
            ' NEED TO USE THAT PARAMETERS TABLE TO KNOW THAT WM & INITIAL = PREPACK

            Fill_Records("POTPACK2", PACK_LIST_NO)
            Fill_Records("POTPACK3", PACK_LIST_NO)
            Dim CARTON_NO_ctr As Integer = 0


            Dim PPK_CODE As String = ""
            For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
                Dim PACK_LIST_SHEET_NO As Integer = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

                Dim CARTON_COUNT2 As Integer = 0

                If CUST_CODE = "WALMART" And INITIAL_ORDER = "1" Then
                    Dim CARTON_COUNT As Integer = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                    TOTAL_CARTONS += CARTON_COUNT
                    CARTON_COUNT2 = CARTON_COUNT
                    Dim CARTON_PACK As Integer = Val(rowPOTPACK2.Item("CARTON_PACK") & "")
                    Dim CARTON_DIMENSIONS As String = rowPOTPACK2.Item("CARTON_DIMENSIONS") & ""

                    rowPOTSHIP2.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP2.Item("PO_SHIP_CTNS") & "") + CARTON_COUNT


                    PPK_CODE = ASCMAIN1.Next_Control_No("PPK_CODE") & "PPK"
                    PPK_CODE = Mid(PPK_CODE, 2)

                    Dim rowWHTPPKM1 As DataRow = dst.Tables("WHTPPKM1").NewRow
                    With rowWHTPPKM1
                        .Item("PPK_CODE") = PPK_CODE
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("PPK_DESC") = "" ' SHOULD BE SAME AS WHAT WAS LOADED ITO rowPOTSHIP7.Item("CARTON_COMMENTS")
                        .Item("LAST_DATE") = DATETIME_STAMP
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("CUSTOM_PPK") = "1"
                        .Item("PPK_QTY_TOTAL") = CARTON_PACK
                    End With
                    dst.Tables("WHTPPKM1").Rows.Add(rowWHTPPKM1)

                    Dim rowPOTSHIP7 As DataRow = dst.Tables("POTSHIP7").NewRow()
                    With rowPOTSHIP7
                        .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                        CARTON_NO_ctr += 1
                        .Item("CARTON_NO") = CARTON_NO_ctr
                        .Item("CARTONS") = CARTON_COUNT
                        .Item("CARTON_COMMENTS") = ""
                        .Item("CUSTOM_PPK") = "1"
                        .Item("PPK_CODE") = PPK_CODE
                        .Item("PO_QTY_PER_CTN") = CARTON_PACK
                        .Item("STYLE_CODE") = ""
                        .Item("COLOR_CODE") = ""
                        .Item("PPK_INNER_QTY") = 0 ' ? CARTON_PACK
                        .Item("CARTON_DIMS") = CARTON_DIMENSIONS
                        Dim CARTON_VOLUME As Decimal = Get_Volume_from_Dims(CARTON_DIMENSIONS)
                        .Item("CARTON_VOLUME") = CARTON_VOLUME
                        .Item("CARTON_WEIGHT") = rowPOTPACK2.Item("CARTON_GRS_WGT")
                    End With
                    dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
                End If

                For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, COLOR_CODE, PACK_LIST_SHEET_LNO")

                    Dim CARTON_COUNT As Integer = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                    Dim CARTON_PACK As Integer = Val(rowPOTPACK3.Item("CARTON_PACK") & "")
                    Dim CARTON_DIMENSIONS As String = rowPOTPACK3.Item("CARTON_DIMENSIONS") & ""

                    Dim PO_ORDER_NO As String = rowPOTPACK3.Item("PO_ORDER_NO") & ""
                    Dim PO_ORDER_LNO As Integer = Val(rowPOTPACK3.Item("PO_ORDER_LNO") & "")

                    If CUST_CODE = "WALMART" And INITIAL_ORDER = "1" Then
                    Else
                        rowPOTSHIP2.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP2.Item("PO_SHIP_CTNS") & "") + CARTON_COUNT
                        TOTAL_CARTONS += CARTON_COUNT
                        Dim rowPOTSHIP7 As DataRow = dst.Tables("POTSHIP7").NewRow()
                        With rowPOTSHIP7
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                            CARTON_NO_ctr += 1
                            .Item("CARTON_NO") = CARTON_NO_ctr
                            .Item("CARTONS") = CARTON_COUNT
                            .Item("CARTON_COMMENTS") = ""
                            .Item("CUSTOM_PPK") = ""
                            .Item("PPK_CODE") = ""
                            .Item("PO_QTY_PER_CTN") = CARTON_PACK
                            .Item("STYLE_CODE") = rowPOTPACK3.Item("STYLE_CODE")
                            .Item("COLOR_CODE") = rowPOTPACK3.Item("COLOR_CODE")
                            .Item("PPK_INNER_QTY") = 0
                            .Item("CARTON_DIMS") = CARTON_DIMENSIONS
                            Dim CARTON_VOLUME As Decimal = Get_Volume_from_Dims(CARTON_DIMENSIONS)
                            .Item("CARTON_VOLUME") = CARTON_VOLUME
                            .Item("CARTON_WEIGHT") = rowPOTPACK3.Item("CARTON_GRS_WGT")
                        End With
                        dst.Tables("POTSHIP7").Rows.Add(rowPOTSHIP7)
                    End If

                    ' need to repeat for 1 to carton_count, and record the lpn
                    Dim rowPOTSHIP8 As DataRow = dst.Tables("POTSHIP8").NewRow()
                    With rowPOTSHIP8
                        .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                        .Item("CARTON_NO") = CARTON_NO_ctr
                        .Item("STYLE_CODE") = rowPOTPACK3.Item("STYLE_CODE")
                        .Item("COLOR_CODE") = rowPOTPACK3.Item("COLOR_CODE")
                        .Item("QTY") = CARTON_PACK
                        .Item("DOZENS") = ""
                        '.Item("PPK_INNER_QTY") = 
                    End With
                    dst.Tables("POTSHIP8").Rows.Add(rowPOTSHIP8)

                    If PPK_CODE <> "" Then
                        Dim rowWHTPPKM2 As DataRow = dst.Tables("WHTPPKM2").NewRow
                        rowWHTPPKM2.Item("PPK_CODE") = PPK_CODE
                        rowWHTPPKM2.Item("STYLE_CODE") = rowPOTSHIP8.Item("STYLE_CODE")
                        rowWHTPPKM2.Item("COLOR_CODE") = rowPOTSHIP8.Item("COLOR_CODE")
                        rowWHTPPKM2.Item("PPK_QTY") = Val(rowPOTSHIP8.Item("QTY") & "") * IIf(rowPOTSHIP8.Item("DOZENS") & "" = "1", 12, 1)
                        dst.Tables("WHTPPKM2").Rows.Add(rowWHTPPKM2)
                    End If

                    Dim rowPOTSHIP3 As DataRow = Nothing
                    Dim rowPOTSHIP3s() As DataRow = dst.Tables("POTSHIP3").Select($"PO_ORDER_NO ='{PO_ORDER_NO}' and PO_ORDER_LNO = {CStr(PO_ORDER_LNO)}", "")

                    If rowPOTSHIP3s.Length = 0 Then
                        Dim rowPOTORDR2 As DataRow = LookUp("POTORDR2", New String() {PO_ORDER_NO, PO_ORDER_LNO})
                        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", New String() {PO_ORDER_NO})

                        Dim PO_QTY_SHP As Int32 = CARTON_COUNT * CARTON_PACK
                        Dim PO_QTY_OPN As Int32 = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")

                        If CUST_CODE = "WALMART" And INITIAL_ORDER = "1" Then
                            PO_QTY_SHP = CARTON_COUNT2 * CARTON_PACK
                        End If

                        Dim SUB_UNIT_PACK_QTY As Integer = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")

                        Dim SPQ As Integer = IIf(SUB_UNIT_PACK_QTY = 0, 12, 12 / SUB_UNIT_PACK_QTY)

                        rowPOTSHIP3 = dst.Tables("POTSHIP3").NewRow()
                        With rowPOTSHIP3
                            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            .Item("PO_QTY_SHP") = PO_QTY_SHP
                            '.Item("PO_QTY_OPN") = PO_QTY_OPN
                            .Item("PO_QTY_REC") = 0
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_DATE") = DATETIME_STAMP

                            'DUTY_RATE_CODE
                            'DUTY_RATE
                            'WEIGHT_FACTOR
                            ' PO_COST_BUFFER = 1

                            '.Item("PO_QTY_UOM") = rowPOTORDR2.Item("PO_QTY_UOM")
                            .Item("PO_COST") = Val(rowPOTORDR2.Item("PO_COST") & "")
                            .Item("PO_COST_VCOST") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "")
                            .Item("PO_COST_MATLS") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "")
                            .Item("PO_COST_VCOST_UM") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "")
                            .Item("PO_COST_MATLS_UM") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "")

                            ' IMPORTANT - note that this field is currently maintained in POTORDR2 per Dozen units, and is per unit in POTSHIP3
                            .Item("PO_COST_OTHER") = Val(rowPOTORDR2.Item("PO_COST_OTHER") & "") / SPQ

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then ' VAN PARANOIA
                                .Item("PO_COST_COMM") = Val(rowPOTORDR2.Item("PO_COST_COMM") & "")
                            Else
                                If rowPOTORDR1.Item("PO_COMM_PAYABLE_TO_BRKR") & "" = "1" Then
                                    .Item("PO_COST_COMM") = Val(rowPOTORDR1.Item("PO_COMM_PCT") & "")
                                Else
                                    .Item("PO_COST_COMM") = 0
                                End If
                            End If

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                .Item("PO_COST_BUFFER") = 5
                            End If

                            ' this is not exactly true- but we can let the calculation routines fix it later
                            .Item("PO_COST_LANDED") = Val(rowPOTORDR2.Item("PO_COST") & "")

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                If Val(rowPOTORDR2.Item("DFQUOTA") & "") = 1 Then
                                    .Item("PO_COST_QUOTA_DF") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "") / SPQ
                                    .Item("PO_COST_QUOTA_DF_DZ") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "")
                                Else
                                    .Item("PO_COST_QUOTA") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "") / SPQ
                                    .Item("PO_COST_QUOTA_DZ") = Val(rowPOTORDR2.Item("PO_COST_QUOTA") & "")
                                End If
                            End If

                            If Val(rowPOTORDR2.Item("PO_COST_VCOST_DZ") & "") = 0 Then
                                .Item("PO_COST_VCOST_DZ") = Val(rowPOTORDR2.Item("PO_COST_VCOST") & "") * SPQ
                            Else
                                .Item("PO_COST_VCOST_DZ") = Val(rowPOTORDR2.Item("PO_COST_VCOST_DZ") & "")
                            End If

                            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                If Val(rowPOTORDR2.Item("PO_COST_MATLS_DZ") & "") = 0 Then
                                    .Item("PO_COST_MATLS_DZ") = Val(rowPOTORDR2.Item("PO_COST_MATLS") & "") * SPQ
                                Else
                                    .Item("PO_COST_MATLS_DZ") = Val(rowPOTORDR2.Item("PO_COST_MATLS_DZ") & "")
                                End If
                            End If

                            .Item("PO_COST_OTHER_DZ") = Val(rowPOTORDR2.Item("PO_COST_OTHER") & "") ' see note above regarding unit of measure for PO_COST_OTHER
                            '.Item("SUB_UNIT_PACK_QTY") = Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "")
                            '.Item("CARTON_PACK_QTY") = Val(rowPOTORDR2.Item("CARTON_PACK_QTY") & "")
                            'If Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & "") = 0 Then
                            '    .Item("PO_QTY_SHP_DZ") = 0
                            '    .Item("NET_OPEN_DZ") = 0
                            'Else
                            '    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                            '        .Item("PO_QTY_SHP_DZ") = PO_QTY_SHP / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                            '    Else
                            '        .Item("PO_QTY_SHP_DZ") = PO_QTY_OPN / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                            '    End If
                            '    .Item("NET_OPEN_DZ") = PO_QTY_OPN / (12 / Val(rowPOTORDR2.Item("SUB_UNIT_PACK_QTY") & ""))
                            'End If
                            '.Item("PO_QTY_REC_DZ") = 0

                            '.Item("PO_REFERENCE") = rowPOTORDR1.Item("PO_REFERENCE")
                            '.Item("PO_DATE_SHIP_BY") = rowPOTORDR2.Item("PO_DATE_SHIP_BY")
                            .Item("FOB_CMT") = (rowPOTORDR1.Item("FOB_CMT") & "")
                            '.Item("VEND_CODE") = rowPOTORDR1.Item("VEND_CODE")

                        End With

                        dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)

                    Else
                        rowPOTSHIP3 = rowPOTSHIP3s(0)
                        rowPOTSHIP3.Item("PO_QTY_SHP") = Val(rowPOTSHIP3.Item("PO_QTY_SHP") & "") + CARTON_COUNT * CARTON_PACK
                    End If
                Next
            Next
        Next


        Dim rowPOTSHIP4 As DataRow = Nothing
        Dim rowPOTSHIP4s() As DataRow = dst.Tables("POTSHIP4").Select($"CONTAINER_NO = '{CONTAINER_NO}'")
        If rowPOTSHIP4s.Length = 0 Then
            PO_SHIPMENT_LNO_ctr = Val(dst.Tables("POTSHIP4").Compute("MAX(PO_SHIPMENT_LNO)", "") & "") + 1
            rowPOTSHIP4 = dst.Tables("POTSHIP4").NewRow
            With rowPOTSHIP4
                .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO_ctr
                .Item("CONTAINER_NO") = CONTAINER_NO
                .Item("CONTAINER_TYPE_CODE") = ""
                .Item("PO_SHIP_CTNS") = TOTAL_CARTONS
                '.Item("PO_SHIP_STATUS") = "?"
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                '.Item("TOTAL_WEIGHT") = -1
                '.Item("CBM") = -1
                '.Item("TRUCKING") = -1
                '.Item("FREIGHT_AMT") = -1
                .Item("CONTAINER_SEAL_NO") = rowPOTVBKG1.Item("CONTAINER_SEAL_NO")
                .Item("TRAILER_NO") = "?"
                .Item("CONTAINER_SEAL_INTACT") = "?"
            End With
            dst.Tables("POTSHIP4").Rows.Add(rowPOTSHIP4)
        Else
            rowPOTSHIP4 = rowPOTSHIP4s(0)
            rowPOTSHIP4.Item("PO_SHIP_CTNS") = Val(rowPOTSHIP4.Item("PO_SHIP_CTNS") & "") + TOTAL_CARTONS
        End If


        For Each TABLE_NAME As String In New String() {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTSHIP4", "POTSHIP7", "POTSHIP8"}
            Update_Record_TDA(TABLE_NAME)
        Next

        For Each TABLE_NAME As String In New String() {"WHTPPKM1", "WHTPPKM2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        Update_Record_TDA("POTVBKG1")

        ASCMAIN1.sql = "Update POTLPNL1 Set PO_SHIPMENT_NO = :PARM1, PO_SHIPMENT_LNO = :PARM2" & vbCrLf _
            & " where PACK_LIST_NO in (Select PACK_LIST_NO from POTVBKG2 where VBKG_NO = :PARM3)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VNV", New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO_ctr, VBKG_NO})

        ' PO SPLITS

        ' MOVE THIS UPDATE TO POFSHIP1 AND THEN WE WON'T NEED TO WORRY ABOUT THE FOLLOWING:
        ' NEED TO UPDATE ICTSTAT2
        ' NEED TO UPDATE POTORDR2

        ' TEST PPK WM INITIALS
        ' CREATE WHTPPKM1/2

        CommitTrans()

        Return PO_SHIPMENT_NO

    End Function
End Class