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

    Dim unFinalize As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Appearance_Red.ForeColor = Drawing.Color.Red


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

            ASCMAIN1.sql = "Select POTVBKG2.*" & vbCrLf _
                & ",POTPACK1.STYLE_CODE_PFX,POTPACK1.PO_REFERENCE,PO_ORDER_NO" & vbCrLf _
                & ",POTPACK1.PACK_LIST_STATUS,POTPACK1.PACK_LIST_DESC,POTPACK1.PACK_LIST_DATE,POTPACK1.INITIAL_ORDER" & vbCrLf _
                & ",POTPACK1.STYLE_CODE_PFX2,POTPACK1.PO_REFERENCE2,POTPACK1.PO_ORDER_NO2,POTPACK1.CUST_CODE" & vbCrLf _
                & " from POTVBKG2,POTPACK1 where POTPACK1.PACK_LIST_NO = POTVBKG2.PACK_LIST_NO" & vbCrLf _
                & " AND POTVBKG2.VBKG_NO = :PARM1"
            ' Create_TDA(.Tables.Add, "POTVBKG2", "*", 1)
            Create_TDA(.Tables.Add, "POTVBKG2", "**", 0, True, "V")

            With .Tables("POTVBKG2")
                .Columns.Add("CARTONS", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select * from POTVBKG3 where POTVBKG3.VBKG_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTVBKG3", "**", 0, True, "V")


            ASCMAIN1.sql = "Select * from POTPACK1 where PACK_LIST_STATUS = 'F' AND VEND_CODE = :PARM1 and VBKG_NO IS NULL"
            Create_TDA(.Tables.Add, "POTPACK1", "**", 0, False, "V")

            With .Tables("POTPACK1")
                .Columns.Add("CARTONS", GetType(System.Int32))
            End With
        End With

        grdPOTVBKGX.DataSource = dst.Tables("POTVBKGX")

        grdPOTVBKG2.DataSource = dst.Tables("POTVBKG2")
        grdPOTPACK1.DataSource = dst.Tables("POTPACK1")
        grdPOTVBKG3.DataSource = dst.Tables("POTVBKG3")

        Dim dvw As DataView = DirectCast(grdPOTPACK1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "VBKG_NO IS NULL"

        Create_Summary(grdPOTVBKGX, "VBKG_NO", "Count")

        Create_Summary(grdPOTVBKG2, "PACK_LIST_NO", "Count")
        Create_Summary(grdPOTVBKG2, "CARTONS")

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


        With grdPOTVBKG3.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                GCOL.CellActivation = Activation.AllowEdit
            Next
        End With
        ' grdPOTVBKG3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes

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

                If VEND_CODE <> VEND_CODE_USER And VEND_CODE_USER <> "" Then
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
                unFinalize = False
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
                            Else
                                If row.Item("VBKG_STATUS") & "" = "F" Then
                                    If MsgBox("Already Finalized - do you want to un-Finalize?", MsgBoxStyle.YesNo,
                                              "") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                    unFinalize = True
                                    chkFinalize.Checked = False
                                    rowPOTVBKG1.Item("VBKG_STATUS") = "O"
                                    VBKG_STATUS = "O"
                                End If

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
                    Dim row As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
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

                    ' check for values in POTVBKG3
                    If dst.Tables("POTVBKG3").Select("VBKG_NO = '" & VBKG_NO & "'").Length = 0 Then
                        EMsg &= vbCr & "There must be at least 1 Container added when finalizing a Booking"
                    End If


                    'If Absx1.txtFor("CONTAINER_NO").Text = "" Or Absx1.txtFor("CONTAINER_SIZE").Text = "" Or Absx1.txtFor("CONTAINER_SEAL_NO").Text = "" Then
                    '    EMsg &= vbCr & "Container, Container Size and Seal are mandatory when finalizing a Booking"
                    'End If


                End If



                If EMsg = "" Then

                    If chkFinalize.Checked Then
                        If MsgBox("You have chosen to Finalize this Booking No upon Update." _
                                & vbCrLf & vbCrLf & "Are you sure that you want to Finalize this Booking No?",
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

                    If ScreenMode And EntryMode = "E" AndAlso (rowPOTVBKG1.Item("VBKG_STATUS") & "" <> "F" Or unFinalize) Then
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

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTVBKG2, grdPOTVBKG3}
                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name = "grdPOTVBKG2" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.True
                            .AllowUpdate = DefaultableBoolean.True
                        ElseIf grd.Name = "grdPOTVBKG3" Then
                            '  .AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
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
            {"POTVBKG1", "POTVBKG2", "POTPACK1", "POTVBKG3"}
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
        optShow.Value = "O"

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

        'Fill_Records("POTVBKG2", VBKG_NO)
        ''   Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
        ''     Fill_Records("POTPACK1", VEND_CODE, True)
        '' DGJ HERE 


        If EntryMode = "N" Then


        Else
            Fill_Records("POTVBKG2", VBKG_NO)
            Fill_Records("POTVBKG3", VBKG_NO)

            For Each rowPOTVBKG2 As DataRow In dst.Tables("POTVBKG2").Select("")
                rowPOTVBKG2.Item("CARTONS") = Get_Cartons(rowPOTVBKG2.Item("PACK_LIST_NO"))
            Next
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

        If unFinalize Then
            rowPOTVBKG1.Item("VBKG_STATUS") = "O"
        End If


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
        Update_Record_TDA("POTVBKG3", SQLD)

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
            {"POTVBKG1", "POTVBKG2", "POTVBKG3"}
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
        Load_Popup_Menu(grdPOTVBKG2, "B", "Packing Lists")
        Load_Popup_Menu(grdPOTPACK1, "BB", "Add Pack List to Booking", "Packing Lists")

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
                Add_Pack_List()

            Case "Packing Lists"
                Dim PACK_LIST_NO As String = grd.ActiveRow.Cells("PACK_LIST_NO").Value
                Dim rowPOTPACK1 As DataRow = LookUp("POTPACK1", PACK_LIST_NO)
                If rowPOTPACK1 IsNot Nothing Then
                    Context_Launch("View", PACK_LIST_NO, e.Tool.Key, "POFPACK1")
                End If


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
                For Each rowPOTPACK1 As DataRow In dst.Tables("POTPACK1").Select("")
                    rowPOTPACK1.Item("CARTONS") = Get_Cartons(rowPOTPACK1.Item("PACK_LIST_NO"))
                Next
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
        ElseIf optShow.Value = "F" Then
            ASCMAIN1.sql = sqlPOTVBKGX & " and VBKG_STATUS = 'F' and PO_SHIPMENT_NO is Null"
            Fill_Records("POTVBKGX", "", True, ASCMAIN1.sql)
            grdPOTVBKGX.Text = "Finalized, Not Shipped"
        ElseIf optShow.Value = "S" Then
            ASCMAIN1.sql = sqlPOTVBKGX & " and VBKG_STATUS = 'F' and PO_SHIPMENT_NO is NOT Null"
            Fill_Records("POTVBKGX", "", True, ASCMAIN1.sql)
            grdPOTVBKGX.Text = "Finalized, Shipped"
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlPOTVBKGX
            Fill_Records("POTVBKGX", "", True, ASCMAIN1.sql)
            grdPOTVBKGX.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdPOTVBKGX, "VBKG_NO".ToLower)
        Sort_grdColumns(grdPOTPACK1, "PACK_LIST_NO".ToLower)
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
                ASCMAIN1.sql = "Select COUNT(*) from POTLPNL1 where PACK_LIST_NO = '" & PACK_LIST_NO & "' " _
                            & " and BARCODE_STATUS = 'A'" _
                            & " and SHIP_CONF <> 'S'"
                Dim UNCONF As Integer = ASCDATA1.GetDataValue

                '       Dim tblPOTLPNL1 As DataTable = ASCDATA1.GetDataTable()
                If UNCONF > 0 Then '     If tblPOTLPNL1.Rows.Count > 0 Then
                    If UNCONF = 1 Then '         If tblPOTLPNL1.Rows.Count = 1 Then
                        'MsgBox("Packing List " & PACK_LIST_NO & " has " & tblPOTLPNL1.Rows.Count & " LPN that has not been confirmed for shipment", MsgBoxStyle.OkOnly, "Cannot Add Pack List " & PACK_LIST_NO)
                        MsgBox($"Packing List {PACK_LIST_NO} has {CStr(UNCONF)} LPN that has not been confirmed for shipment", MsgBoxStyle.OkOnly, $"Cannot Add Pack List {PACK_LIST_NO}")
                    Else
                        'MsgBox("Packing List " & PACK_LIST_NO & " have " & tblPOTLPNL1.Rows.Count & " LPNs that have not been confirmed for shipment", MsgBoxStyle.OkOnly, "Cannot Add Pack List " & PACK_LIST_NO)
                        MsgBox($"Packing List {PACK_LIST_NO} have {CStr(UNCONF)} LPNs that have not been confirmed for shipment", MsgBoxStyle.OkOnly, $"Cannot Add Pack List {PACK_LIST_NO}")
                    End If
                    Exit Sub
                End If


                'For Each rowPOTLPNL1 As DataRow In ASCDATA1.GetDataTable.Rows
                '        If rowPOTLPNL1.Item("SHIP_CONF") & "" <> "S" Then
                '            MsgBox("This Pack List Has Bar Codes that Have a Ship Conf Other than 'S'", MsgBoxStyle.OkOnly, "Cannot Add Pack List " & PACK_LIST_NO)
                '            Exit Sub
                '        End If
                '    Next

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
                    rowPOTVBKG2_new.Item("CARTONS") = Get_Cartons(PACK_LIST_NO)
                    dst.Tables("POTVBKG2").Rows.Add(rowPOTVBKG2_new)
                    grdPOTPACK1.ActiveRow.Cells("VBKG_NO").Value = VBKG_NO
                    grdPOTPACK1.ActiveRow.Update()
                    '    Sort_grdColumns(grdPOTPACK1, "PACK_LIST_NO", True)
                    Sort_grdColumns(grdPOTPACK1, "PACK_LIST_NO".ToLower)
                Else
                    MsgBox("This Pack List No is no longer available to add to Booking", MsgBoxStyle.OkOnly, "Cannot Add Pack List")
                    '       EMsg &= vbCr & " This Pack List No is no longer available to add to Booking"
                    Fill_Records("POTPACK1", VEND_CODE, True)
                    For Each rowPOTPACK1 As DataRow In dst.Tables("POTPACK1").Select("")
                        rowPOTPACK1.Item("CARTONS") = Get_Cartons(rowPOTPACK1.Item("PACK_LIST_NO"))
                    Next
                End If

            End If


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

    Function Get_Cartons(PACK_LIST_NO As String) As Int32
        ' If PACK_LIST_NO = "000345" Or PACK_LIST_NO = "000316" Then Stop
        ASCMAIN1.sql = "" _
            & "Select Sum (CARTONS) from (" & vbCrLf _
            & "SELECT '2' PACK, SUM (NVL(CARTON_COUNT,0)) CARTONS from POTPACK2,POTPACK1 " & vbCrLf _
            & "where POTPACK2.PACK_LIST_NO = POTPACK1.PACK_LIST_NO and POTPACK1.INITIAL_ORDER = '1'" & vbCrLf _
            & "and POTPACK1.PACK_LIST_NO = :PARM1" & vbCrLf _
            & "union" & vbCrLf _
            & "SELECT '3' PACK, SUM (NVL(CARTON_COUNT,0)) CARTONS from POTPACK3,POTPACK1 " & vbCrLf _
            & "where POTPACK3.PACK_LIST_NO = POTPACK1.PACK_LIST_NO and not (POTPACK1.INITIAL_ORDER = '1')" & vbCrLf _
            & "and POTPACK1.PACK_LIST_NO = :PARM1" & vbCrLf _
            & ")"

        Return Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {PACK_LIST_NO}))
    End Function


    Private Sub grdPOTVBKG3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTVBKG3.AfterRowActivate
        With grdPOTVBKG3.DisplayLayout.Bands(0)
            If grdPOTVBKG3.ActiveRow.IsAddRow Then
                .Columns("CONTAINER_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CONTAINER_SEAL_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CONTAINER_SIZE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CONTAINER_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CONTAINER_SEAL_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CONTAINER_SIZE").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End With

        If grdPOTVBKG3.ActiveRow.IsAddRow Then
            grdPOTVBKG3.ActiveRow.Cells("VBKG_NO").Value = VBKG_NO
            grdPOTVBKG3.ActiveRow.Cells("LINE_NO").Value = Val(dst.Tables("POTVBKG3").Compute("Max(LINE_NO)", "") & "") + 1

        Else


        End If

        ' grdPOTVBKG3.Rows[0].Cells["abc"].Activate()
        ' grdPOTVBKG3.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode)
    End Sub

    Private Sub grdPOTVBKG3_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdPOTVBKG3.BeforeCellUpdate
        If e.Cell.Column.Key.StartsWith("CONTAINER_NO") Then
            Dim row As DataRow = dst.Tables("POTVBKG3").Select("CONTAINER_NO = '" & e.Cell.Row.Cells("CONTAINER_NO").Text & "'").FirstOrDefault
            If row Is Nothing Then
            Else
                MsgBox("Container Already Exits", MsgBoxStyle.OkOnly, "Container Entry")
                e.Cancel = True
                '  grdPOTVBKG3.Refresh()
                grdPOTVBKG3.ActiveRow.CancelUpdate()
                Exit Sub
            End If
        End If
    End Sub

    Private Sub grdPOTVBKG3_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdPOTVBKG3.BeforeRowUpdate
        If e.Row.Cells("CONTAINER_NO").Value & "" = "" Or e.Row.Cells("CONTAINER_SEAL_NO").Value & "" = "" Or e.Row.Cells("CONTAINER_SIZE").Value & "" = "" Then
            MsgBox("You Must Enter Container No, Seal No and Container Size", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            e.Cancel = True
            grdPOTVBKG3.ActiveRow.CancelUpdate()
            Exit Sub
        End If
    End Sub
End Class