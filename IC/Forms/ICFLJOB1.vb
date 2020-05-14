Public Class ICFLJOB1
    'Load_Events

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name

    Dim LABEL_JOB_NO As String
    Dim LABEL_FORMAT_CODE As String

    Dim rowICTLJOB1 As DataRow
    Dim rowARTCUST1 As DataRow      ' ARTCUST1 for the Sold-To
    Dim rowICTSTYL1 As DataRow

    Dim rowICTULBL2 As DataRow

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTLJOB1.*" & vbCrLf _
                & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & ", POTORDR1.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_ETA" & vbCrLf _
                & " from ICTLJOB1,SOTORDR0,POTORDR1" & vbCrLf _
                & " where ICTLJOB1.LABEL_JOB_STATUS = :PARM1" & vbCrLf _
                & " and SOTORDR0.ORDR_GROUP_NO (+) = ICTLJOB1.ORDR_GROUP_NO" & vbCrLf _
                & " and POTORDR1.PO_ORDER_NO (+) = ICTLJOB1.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "ICTLJOBX", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select ICTLJOB1.*, ICTULBL2.REPORT_NAME" & vbCrLf _
                & ", ICTULBL2.LABEL_MODEL_CODE, ICTULBL2.LABEL_FORMAT_DESC" & vbCrLf _
                & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & ", POTORDR1.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_ETA" & vbCrLf _
                & " from ICTLJOB1,SOTORDR0,POTORDR1,ICTULBL2,ICTULBL1" & vbCrLf _
                & " where ICTLJOB1.LABEL_JOB_NO = :PARM1" & vbCrLf _
                & " and SOTORDR0.ORDR_GROUP_NO (+) = ICTLJOB1.ORDR_GROUP_NO" & vbCrLf _
                & " and POTORDR1.PO_ORDER_NO (+) = ICTLJOB1.PO_ORDER_NO" & vbCrLf _
                & " and ICTULBL2.CUST_CODE = ICTLJOB1.CUST_CODE" & vbCrLf _
                & " and ICTULBL2.LABEL_FORMAT_CODE = ICTLJOB1.LABEL_FORMAT_CODE" & vbCrLf _
                & " and ICTULBL1.LABEL_MODEL_CODE = ICTULBL2.LABEL_MODEL_CODE"
            Create_TDA(.Tables.Add, "ICTLJOB1", "**", 0, True, "V", 1)
            '  Create_TDA(.Tables.Add, "ICTLJOB1", "*", 1)

            ASCMAIN1.sql = "Select ICTLJOB2.*, ICTSTYL1.STYLE_DESC, nvl(ICTSTYC1.STYLE_COLOR_DESC , ICTCOLR1.COLOR_DESC) COLOR_DESC, ICTSTYL1.SIZE_CODE, ICTSTYL1.SALES_DIVISION_CODE" _
            & " from ICTLJOB2,ICTSTYL1,ICTSTYC1,ICTCOLR1" _
            & " where ICTSTYL1.STYLE_CODE = ICTLJOB2.STYLE_CODE" _
            & "   and ICTSTYC1.STYLE_CODE = ICTLJOB2.STYLE_CODE" _
            & "   and ICTSTYC1.COLOR_CODE = ICTLJOB2.COLOR_CODE" _
            & "   and ICTCOLR1.COLOR_CODE = ICTLJOB2.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTLJOB2", "**", 1)

            With .Tables.Add("ICTLJOB3")
                .Columns.Add("LABEL_JOB_NO")
                .Columns.Add("LABEL_JOB_LNO", GetType(System.Int32))
                .Columns.Add("LABEL_NO", GetType(System.Int32))
                .Columns.Add("LABEL_SPACER")
                .PrimaryKey = New DataColumn() {.Columns("LABEL_JOB_NO"), .Columns("LABEL_JOB_LNO"), .Columns("LABEL_NO")}
            End With

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            'Create_TDA(.Tables.Add, "ICTULBL1", "*", 1, False)
            ASCMAIN1.sql = "Select ICTULBL2.*, ICTULBL1.LABELS_ACROSS, ICTULBL1.LABELS_DOWN" & vbCrLf _
                & " from ICTULBL1,ICTULBL2" & vbCrLf _
                & " where ICTULBL1.LABEL_MODEL_CODE = ICTULBL2.LABEL_MODEL_CODE" & vbCrLf _
                & "   and ICTULBL2.CUST_CODE = :PARM1" & vbCrLf _
                & "   and ICTULBL2.LABEL_FORMAT_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTULBL2", "**", 0, False, "VV")
        End With

        grdICTLJOBX.DataSource = dst.Tables("ICTLJOBX")
        grdICTLJOB2.DataSource = dst.Tables("ICTLJOB2")

        grdICTLJOBX.DisplayLayout.UseFixedHeaders = True
        With grdICTLJOBX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"LABEL_JOB_NO", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdICTLJOB2.DisplayLayout.UseFixedHeaders = True
        With grdICTLJOB2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"LABEL_JOB_LNO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdICTLJOB2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ORDR_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"LABEL_QTY", "LABEL_SEL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                ElseIf New String() {"LABEL_STYLE", "LABEL_PRICE", "LABEL_DGC", "LABEL_SIZE", "LABEL_UPC", "LABEL_QTY", "LABEL_COLOR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC", "SIZE_CODE", "SALES_DIVISION_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdICTLJOBX, "LABEL_JOB_NO", "Count")

        Create_Summary(grdICTLJOB2, "LABEL_JOB_LNO", "Count")
        Create_Summary(grdICTLJOB2, New String() {"ORDR_QTY", "LABEL_QTY", "LABEL_SEL"})

        Show_Filter(grdICTLJOBX, True)
        grdICTLJOBX.DisplayLayout.GroupByBox.Hidden = False

        SplitContainer1.Panel2Collapsed = True ' UNTIL WE GET USAGE STATS OR SOMETHING ELSE MEANINGFUL TO DISPLAY

        tabInfo.Tabs("Usage").Visible = False

        ASCMAIN1.Add_Value_List(grdICTLJOBX, "LABEL_JOB_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})

        Check_InquiryMode()
        MakeTransparent(chkShowSelectedOnly)
    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "ICFLJOBI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        LABEL_FORMAT_CODE = Absx1.txtFor("LABEL_FORMAT_CODE").Text
                        If LABEL_FORMAT_CODE = "" Then
                            EMsg &= vbCr & "You Must Provide a Value for Label Format Code"
                        Else
                            If LookUp("ICTULBL2", New String() {CUST_CODE, LABEL_FORMAT_CODE}) Is Nothing Then
                                EMsg &= vbCr & "Invalid value specified for Label Format Code"
                            End If
                        End If
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If optLABELS_BY.Value & "" = "" Then
                    EMsg &= vbCr & "You must choose whether Labels will be created from a Sales Order or a Purchase Order"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ICTLJOB1", CUST_CODE) Then Exit Sub
                End If

            Case "Edit", "View"

                CUST_CODE = ""
                LABEL_JOB_NO = ""

                If Absx1.txtFor("LABEL_JOB_NO").Text = "" Then
                    EMsg &= vbCr & "No Label Job No Specified"
                Else
                    LABEL_JOB_NO = Absx1.txtFor("LABEL_JOB_NO").Text
                    rowICTLJOB1 = LookUp("ICTLJOB1", LABEL_JOB_NO)
                    If rowICTLJOB1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Label Job No " & LABEL_JOB_NO
                    Else
                        CUST_CODE = rowICTLJOB1.Item("CUST_CODE")
                        If rowICTLJOB1.Item("LABEL_JOB_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowICTLJOB1.Item("LABEL_JOB_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "Label Job No " & LABEL_JOB_NO & " has been Cancelled"
                                Case "D"
                                    EMsg &= vbCr & "Label Job No " & LABEL_JOB_NO & " has been Deleted"
                                Case Else ' such as "F"
                                    EMsg &= vbCr & "Label Job No " & LABEL_JOB_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("ICTLJOB1", LABEL_JOB_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("ICTLJOB1", CUST_CODE) Then Exit Sub
                End If

            Case "Update"
                If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                    Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                Else
                    If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                     > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                    End If
                End If

                If Absx1.txtFor("ORDR_GROUP_NO").Text = "" Then
                    EMsg &= vbCr & "Order Group No is required"
                Else
                    If LookUp("SOTORDR0", Absx1.txtFor("ORDR_GROUP_NO").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Order Group No"
                    End If
                End If

                If optLABELS_BY.Value & "" = "P" Then
                    If Absx1.txtFor("PO_ORDER_NO").Text = "" Then
                        EMsg &= vbCr & "PO No is required"
                    Else
                        If LookUp("POTORDR1", Absx1.txtFor("PO_ORDER_NO").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid PO No"
                        End If
                    End If
                End If

                If optLABELS_BY.Value & "" = "S" Then
                    If Absx1.txtFor("LABEL_FORMAT_CODE").Text = "" Then
                        EMsg &= vbCr & "Label Format is required"
                    Else
                        If LookUp("ICTULBL2", New String() {Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("LABEL_FORMAT_CODE").Text}) Is Nothing Then
                            EMsg &= vbCr & "Invalid Label Format"
                        End If
                    End If
                End If

 

                If Absx1.txtFor("LABEL_FORMAT_CODE").Text = "" Then
                    EMsg &= vbCr & "Label Format is required"
                End If
                If grdICTLJOB2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Label Job"
                Else
                    If Val(dst.Tables("ICTLJOB2").Compute("COUNT(LABEL_JOB_LNO)", "LABEL_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Lines on Label Job with Qty >0"
                    End If
                    If dst.Tables("ICTLJOB2").Select("LABEL_SEL = '1'").Length = 0 Then
                        EMsg &= vbCr & "No Lines Selected for Printing"
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from ICTLJOB2 where rownum < 1" ' should be counting prints
                ASCMAIN1.sql &= " and LABEL_JOB_NO = '" & LABEL_JOB_NO & "'"

                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Label Job has been Used"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Label Job as Deleted", _
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Print"

                If grdICTLJOB2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Label Job"
                Else
                    If Val(dst.Tables("ICTLJOB2").Compute("COUNT(LABEL_JOB_LNO)", "LABEL_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Lines on Label Job with Qty >0"
                    End If
                    If dst.Tables("ICTLJOB2").Select("LABEL_SEL = '1'").Length = 0 Then
                        EMsg &= vbCr & "No Lines Selected for Printing"
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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Order()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowICTLJOB1.Item("LABEL_JOB_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If
                .Items("Update").Settings.Enabled = iScreenMode

                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Print").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Delete").Visible = (EntryMode = "E")
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
            End With

            .Groups("Special").Visible = (Not (EntryMode = "V") And ScreenMode)
        End With

        lblStatus.Visible = ScreenMode

        grdICTLJOBX.Visible = Not tf

        chkShowSelectedOnly.Checked = False


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("LABEL_FORMAT_CODE"), InquiryMode Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")))
        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))

        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTLJOB2.DisplayLayout.Bands(0).Columns
            If gcol.Key = "LABEL_PRICE" Then
                If (EntryMode = "E" Or EntryMode = "N") Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End If
        Next

        Set_Read_Only(grpPrintControls, False)



        If ScreenMode Then
            'If EntryMode = "V" Then
            '    grdICTLJOB2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            '    grdICTLJOB2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            '    grdICTLJOB2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            'Else
            grdICTLJOB2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdICTLJOB2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdICTLJOB2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            'End If

            'Set_Read_Only_for_ctl(optLABELS_BY, Not (EntryMode = "N"))

            btnRefresh.Visible = (EntryMode = "E")
            grpPO.Visible = (optLABELS_BY.Value = "P")
            grpSO.Visible = (optLABELS_BY.Value = "S")
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("LABEL_JOB_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""
        txtPO_IMPORT.Text = ""

        CUST_CODE = ""
        LABEL_JOB_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ICTLJOB1", "ICTLJOB2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Load_ICTLJOBX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            LABEL_JOB_NO = ASCMAIN1.Next_Control_No("ICTLJOB1.LABEL_JOB_NO")

            rowICTLJOB1 = dst.Tables("ICTLJOB1").NewRow
            With rowICTLJOB1
                .Item("LABEL_JOB_NO") = LABEL_JOB_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("LABEL_FORMAT_CODE") = LABEL_FORMAT_CODE
                .Item("LABEL_JOB_STATUS") = "O"
                .Item("LABEL_QTY_CALC") = "S"
                .Item("LABEL_QTY_EXTRA") = 0
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LABELS_BY") = HFs("LABELS_BY") ' "S"
            End With
            dst.Tables("ICTLJOB1").Rows.Add(rowICTLJOB1)

        Else
            rowICTLJOB1 = Fill_Record("ICTLJOB1", LABEL_JOB_NO)
            CUST_CODE = rowICTLJOB1.Item("CUST_CODE")
            LABEL_FORMAT_CODE = rowICTLJOB1.Item("LABEL_FORMAT_CODE")
        End If

        CUST_CODE = rowICTLJOB1.Item("CUST_CODE")
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)

        rowICTULBL2 = Fill_Record("ICTULBL2", New String() {CUST_CODE, LABEL_FORMAT_CODE})
        Absx1.numFor("LABELS_ACROSS").Value = Val(rowICTULBL2.Item("LABELS_ACROSS"))
        Absx1.numFor("LABELS_DOWN").Value = Val(rowICTULBL2.Item("LABELS_DOWN"))

        Fill_Records("ICTLJOB2", LABEL_JOB_NO)
        Sort_grdColumns(grdICTLJOB2, "LABEL_JOB_LNO")
        Toggle_Selected()

        lblINIT_DATE.Text = "Created " & Format(rowICTLJOB1.Item("INIT_DATE"), "MM/dd/yyyy HH:mm")

        If EntryMode = "N" Then
            lblStatus.Text = "New Job"

        Else
            Select Case rowICTLJOB1.Item("LABEL_JOB_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
            End Select
        End If

        Setup_LABELS_BY()
        Absx1.txtFor("PO_ORDER_NO").ReadOnly = True

        Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, LABEL_JOB_NO)
        For Each TABLE_NAME As String In New String() _
            {"ICTLJOB1", "ICTLJOB2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where LABEL_JOB_NO = '" & LABEL_JOB_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        If EntryMode <> "N" Then Delete_Records()

        INIT_LAST("ICTLJOB1", False, , True)
        Dim sqldelete As String = "LABEL_JOB_NO = '" & LABEL_JOB_NO & "'"
        Update_Record_TDA("ICTLJOB1", sqldelete)
        Update_Record_TDA("ICTLJOB2", sqldelete)
        Dependent_Updates(1, LABEL_JOB_NO)

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "CUST_CODE"

                sql_where = " and CUST_CODE in (Select Distinct CUST_CODE from ICTULBL2)"


            Case "LABEL_FORMAT_CODE"

                sql_where = " and ICTULBL2.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"


            Case "LABEL_JOB_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    MsgBox("You must enter a Customer Code", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""
                sql_where = " and ICTLJOB1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                If InquiryMode Then
                Else
                    sql_where &= " and ICTLJOB1.LABEL_JOB_STATUS = 'O' "
                End If

            Case "ORDR_GROUP_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    MsgBox("You must enter a Customer Code", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""
                sql_where = " and SOTORDR0.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                sql_where &= " and (SOTORDR0.ORDR_CNT_OPEN <> 0 OR SOTORDR0.ORDR_CNT_PICK <> 0)"


            Case "PO_ORDER_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    MsgBox("You must enter a Customer Code", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""
                '  sql_where = " and SOTORDR0.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                sql_where &= " and PO_STATUS = 'O'"

        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("LABEL_JOB_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ICTLJOB1"
            E.COLUMN_NAME = "LABEL_JOB_NO"
            E.CODE_VALUE = Absx1.txtFor("LABEL_JOB_NO").Text
            E.DESC_VALUE = "Label Job"
            E.ATTACHMENT_NOTES = ""
            'If rowICTLJOB1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTLJOBX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdICTLJOB2, "BBBB", "Style Status Inquiry", "Select All", "De-Select All", "Select Selected")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case grd.Name
            Case "grdICTLJOB2"
                tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N" Or EntryMode = "V")
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N" Or EntryMode = "V")
                tlb_btn = DirectCast(tlb_pop.Tools("Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N" Or EntryMode = "V")

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("LABEL_SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Select Selected"
                For Each grow As Infragistics.Win.UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        If grow.Selected Then
                            grow.Cells("LABEL_SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                            grow.Update()
                        End If
                    End If
                Next
                grd.Selected.Rows.Clear()

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_ICTLJOBX()
                End If

            Case "LABEL_FORMAT_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode _
                       And Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("LABEL_FORMAT_CODE").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "ORDR_GROUP_NO"

                If EntryMode = "N" Then
                    Fill_ORDR_GROUP_NO()
                End If

            Case "PO_ORDER_NO"

                If EntryMode = "N" Then
                    Fill_PO_ORDER_NO()
                End If

            Case "LABEL_JOB_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_ICTLJOBX()

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_ICTLJOBX()
            Case "LABEL_JOB_NO"
                Click_Command("View")

            Case "ORDR_GROUP_NO"
                If EntryMode = "N" Then
                    Fill_ORDR_GROUP_NO()
                End If

            Case "PO_ORDER_NO"
                If EntryMode = "N" Then
                    Fill_PO_ORDER_NO()
                End If

        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_ICTLJOBX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE = "" Then
            Fill_Records("ICTLJOBX", "O")
            grdICTLJOBX.Text = "Open Label Jobs"
            Sort_grdColumns(grdICTLJOBX, "LABEL_JOB_NO".ToLower)
        Else
            'ASCMAIN1.sql = "Select * from ICTLJOB1 where LABEL_JOB_STATUS = 'O'" _
            '    & " and CUST_CODE = '" & CUST_CODE & "'"
            'Fill_Records("ICTLJOBX", "", , ASCMAIN1.sql)
            'grdICTLJOBX.Text = "Open Label Jobs associated with " & CUST_CODE
            'Sort_grdColumns(grdICTLJOBX, "LABEL_JOB_NO".ToLower)
        End If
        grdICTLJOBX.Visible = True
    End Sub

    Sub Print_Record()
        ' NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM
        ' AND SHOULD BE USING THE DATALAYER OF SORUSED1

        Calc_LABEL_QTY()

        Dim LABELS_ACROSS As Integer = Val(Absx1.numFor("LABELS_ACROSS").Value & "")
        Dim LABELS_DOWN As Integer = Val(Absx1.numFor("LABELS_DOWN").Value & "")

        dst.Tables("ICTLJOB3").Rows.Clear()
        For Each rowICTLJOB2 As DataRow In dst.Tables("ICTLJOB2").Select("LABEL_SEL = '1' AND LABEL_QTY <> 0")

            Dim LABEL_SIZE As String = rowICTLJOB2.Item("LABEL_SIZE") & ""
            LABEL_SIZE = Replace(LABEL_SIZE, " ", "")
            rowICTLJOB2.Item("LABEL_SIZE") = LABEL_SIZE

            Dim LABEL_JOB_LNO As Integer = Val(rowICTLJOB2.Item("LABEL_JOB_LNO") & "")
            Dim LABEL_QTY As Integer = Val(rowICTLJOB2.Item("LABEL_QTY") & "")

            Dim LABEL_QTY_CALC As String = Absx1.optFor("LABEL_QTY_CALC").Value
            For LABEL_NO As Integer = 1 To LABEL_QTY
                Dim LABEL_SPACER As String = "0"
                If Absx1.chkFor("LABEL_SPACER").Checked Then
                    If LABEL_QTY_CALC = "O" Or LABEL_QTY_CALC = "X" Then
                        If LABEL_NO = LABEL_QTY Then LABEL_SPACER = "1"
                    End If
                    If LABEL_QTY_CALC = "R" Then
                        If LABEL_NO > LABEL_QTY - LABELS_ACROSS Then LABEL_SPACER = "1"
                    End If
                End If
                dst.Tables("ICTLJOB3").Rows.Add(New Object() {LABEL_JOB_NO, LABEL_JOB_LNO, LABEL_NO, LABEL_SPACER})
            Next
        Next

        If dst.Tables("ICTLJOB3").Rows.Count = 0 Then
            MsgBox("No Labels to Print", MsgBoxStyle.OkOnly, "Cannot Print Labels")
            Exit Sub
        End If

        Dim REPORT_NAME As String = rowICTLJOB1.Item("REPORT_NAME")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = REPORT_NAME ' unneccesary if Report Name is Like Form Name
        Generate_Report(RPT, "UPC Label", , , , , False)
        Print_Report_End()


        'Dim REPORTFILE As String = "ICRULBLA"
        'If Not REPORTS.ContainsKey(REPORTFILE) Then
        '    REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
        '    REPORTS(REPORTFILE).Prepare_dst(False, "")
        'End If

        ''To fill the report's dataset with data from Oracle, 
        '' set the parameter array to values that the Fill_Records_RPT method expects, and then call it
        ''REPORTS(REPORTFILE).Fill_Records_RPT(New String() {"USEDPER_CTL_NO = '" & USEDPER_CTL_NO & "'"})

        ''To fill the report's dataset with data from this form's dataset:
        'With REPORTS(REPORTFILE).clsASCBASE1
        '    .EnforceConstraints(False)
        '    For Each TABLE_NAME As String In New String() {"SOTPPDI1", "SOTPPDI2", "SOTPPDI3", "SOTINVH1", "SOTSVIA1"}
        '        .dst.Tables(TABLE_NAME).Rows.Clear()
        '        Dim SQL As String = ""
        '        If TABLE_NAME = "SOTINVH1" Then
        '            SQL = "LABEL_JOB_NO = '" & LABEL_JOB_NO & "'"
        '        End If

        '        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(SQL)
        '            Dim rowr As DataRow = .dst.Tables(TABLE_NAME).NewRow
        '            If TABLE_NAME = "SOTPPDI2" Or TABLE_NAME = "SOTPPDI3" Or TABLE_NAME = "SOTINVH1" Then

        '                For I As Integer = 0 To .dst.Tables(TABLE_NAME).Columns.Count - 1
        '                    Dim COLUMN_NAME As String = .dst.Tables(TABLE_NAME).Columns(I).ColumnName
        '                    rowr.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
        '                Next
        '            Else
        '                rowr.ItemArray = row.ItemArray
        '            End If
        '            .dst.Tables(TABLE_NAME).Rows.Add(rowr)
        '        Next
        '    Next
        '    .EnforceConstraints(True)
        'End With
        'With REPORTS(REPORTFILE).clsASCBASE1
        '    .Print_Report_Begin()
        '    .CR_params.Add("SUBT", "")
        '    .Generate_Report("SORUSED1", "USEDper Invoice Report", , True, , , , , False)
        '    .Print_Report_End()
        'End With

    End Sub

    Private Sub grdICTLJOBX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTLJOBX.DoubleClickRow
        Absx1.txtFor("LABEL_JOB_NO").Text = e.Row.Cells("LABEL_JOB_NO").Value
        Click_Command("View")
    End Sub


    Sub Cancel_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        Dim EMsg As String = ""
        If EntryMode = "E" Then
            Cancel_Order_1(LABEL_JOB_NO)
            EMsg = "Label Job " & LABEL_JOB_NO & " has been Cancelled"
        End If

        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Cancel_Order_1(LABEL_JOB_NO As String)
        Dependent_Updates(-1, LABEL_JOB_NO)

        ASCMAIN1.sql = "Update ICTLJOB1 Set LABEL_JOB_STATUS = :PARM1" _
            & " where LABEL_JOB_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"C", LABEL_JOB_NO})
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        Dim EMsg As String = ""

        BeginTrans()

        If EntryMode = "E" Then
            Delete_Order_1(LABEL_JOB_NO)
            EMsg = "Label Job No " & LABEL_JOB_NO & " has been marked as Deleted"
        End If

        CommitTrans(EMsg)

        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(LABEL_JOB_NO As String)
        Dependent_Updates(-1, LABEL_JOB_NO)

        ASCMAIN1.sql = "Update ICTLJOB1 Set LABEL_JOB_STATUS = :PARM1" _
            & " where LABEL_JOB_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"D", LABEL_JOB_NO})
    End Sub

    Sub Dependent_Updates(S As Integer, LABEL_JOB_NO As String)

    End Sub

    Sub Display_Totals()

    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdICTLJOB2.ActiveRow
            Select Case COLUMN_NAME

                Case "LABEL_QTY"

            End Select
        End With
    End Sub

#Region "grdICTLJOB2"

    Private Sub grdICTLJOB2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTLJOB2.AfterCellUpdate
        With grdICTLJOB2.ActiveRow
            Select Case e.Cell.Column.Key

            End Select
        End With
    End Sub

    Private Sub grdICTLJOB2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTLJOB2.AfterRowActivate


        If grdICTLJOB2.ActiveRow Is Nothing OrElse grdICTLJOB2.ActiveRow.IsAddRow OrElse Not grdICTLJOB2.ActiveRow.IsDataRow Then
            grdICTSTAT2.Visible = False
        Else
            'Dim STYLE_CODE As String = grdICTLJOB2.ActiveRow.Cells("STYLE_CODE").Value
            'Dim COLOR_CODE As String = grdICTLJOB2.ActiveRow.Cells("COLOR_CODE").Value
            ''  Fill_Records("ICTSTAT2", New String() {STYLE_CODE, COLOR_CODE})
            'grdICTSTAT2.Text = "Style Status for " & STYLE_CODE & ":" & COLOR_CODE
            'grdICTSTAT2.Visible = True
        End If
    End Sub

    Private Sub grdICTLJOB2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdICTLJOB2.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdICTLJOB2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTLJOB2.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdICTLJOB2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdICTLJOB2.BeforeCellUpdate

    End Sub

    Private Sub grdICTLJOB2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTLJOB2.BeforeExitEditMode
        If grdICTLJOB2.ActiveCell IsNot Nothing Then
            With grdICTLJOB2.ActiveCell
                Select Case .Column.Key
                    'Case "STYLE_CODE", "COLOR_CODE"
                    '    .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdICTLJOB2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTLJOB2.BeforeRowsDeleted

    End Sub

    Private Sub grdICTLJOB2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTLJOB2.BeforeRowUpdate

        If e.Row.IsAddRow Then
            e.Row.Cells("LABEL_JOB_NO").Value = LABEL_JOB_NO
            Dim LABEL_JOB_LNO As Int64 = Val(dst.Tables("ICTLJOB2").Compute("MAX(LABEL_JOB_LNO)", "") & "") + 1
            e.Row.Cells("LABEL_JOB_LNO").Value = LABEL_JOB_LNO
        End If
    End Sub

    Private Sub grdICTLJOB2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTLJOB2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

            End Select
        End With

    End Sub
#End Region

    Sub Load_Events()
        '    grdEvents.RemoveAll
        '    Call Load_Events_1("Entered", "INIT_DATE")
        '    Call Load_Events_1("Modified", "LAST_DATE")


    End Sub

    Sub Fill_ORDR_GROUP_NO()
        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", Absx1.txtFor("ORDR_GROUP_NO").Text)
        If rowSOTORDR0 IsNot Nothing Then
            Absx1.txtFor("ORDR_CUST_PO").Text = rowSOTORDR0.Item("ORDR_CUST_PO") & ""
            Absx1.dteFor("ORDR_SHIP_DATE").Value = rowSOTORDR0.Item("ORDR_SHIP_DATE")
            Absx1.dteFor("ORDR_CANCEL_DATE").Value = rowSOTORDR0.Item("ORDR_CANCEL_DATE")

            Dim ORDR_GROUP_NO As String = Absx1.txtFor("ORDR_GROUP_NO").Text

            Dim LABELS_ACROSS As Integer = Val(rowICTULBL2.Item("LABELS_ACROSS") & "")
            Dim LABELS_DOWN As Integer = Val(rowICTULBL2.Item("LABELS_DOWN") & "")

            Dim LABELS_PER_SHEET As Integer = LABELS_ACROSS * LABELS_DOWN
            If LABELS_PER_SHEET = 0 Then LABELS_PER_SHEET = 24

            If AbsCheckBox2.Checked = True Then

                ASCMAIN1.sql = "(select ICTSTYL1.STYLE_CODE, ICTSTYL1.SIZE_CODE, 0 SIZE_QTY from ICTSTYL1, ICTSTYLS" & vbCrLf _
                    & "  where ICTSTYL1.STYLE_CODE = ICTSTYLS.STYLE_CODE(+)"

                Dim Size As Integer
                For Size = 1 To 24
                    ASCMAIN1.sql = ASCMAIN1.sql & vbCrLf _
                    & "  union" & vbCrLf _
                    & "  select ICTSTYL1.STYLE_CODE, ICTSTYLS.SIZE_" & Format(Size, "00") & ", QTY_" & Format(Size, "00") & " from ICTSTYL1, ICTSTYLS" & vbCrLf _
                    & "  where ICTSTYL1.STYLE_CODE = ICTSTYLS.STYLE_CODE" & vbCrLf _
                    & "  and ICTSTYLS.SIZE_" & Format(Size, "00") & " is not null" & vbCrLf _
                    & "  and ICTSTYL1.SIZE_CODE is null"
                Next
                ASCMAIN1.sql = ASCMAIN1.sql & ")"

            Else
                ASCMAIN1.sql = "(select ICTSTYL1.STYLE_CODE, ICTSTYL1.SIZE_CODE, 1 SIZE_QTY from ICTSTYL1 )"

            End If


            ASCMAIN1.sql = "SELECT STYLE_CODE, COLOR_CODE, SIZE_CODE, " & IIf(AbsCheckBox2.Checked, " trunc(ORDR_QTY * (SIZE_QTY / PPK_CNT) + .5) ORDR_QTY", " ORDR_QTY") & vbCrLf _
                & ", LABEL_STYLE, LABEL_PRICE, LABEL_DGC, LABEL_SIZE, LABEL_UPC, LABEL_COLOR, LABEL_DESC, SIZE_QTY, PPK_CNT from " & vbCrLf _
                & "(Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.SIZE_CODE" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", EDT850T2.EDI_SKU LABEL_STYLE, EDT850T2.RETAIL_PRICE LABEL_PRICE, EDT850T2.BUYER_CATALOG_NO LABEL_DGC" & vbCrLf _
                & ", nvl(EDT850T2.EDI_SIZE_DESC, ICTSTYL1.SIZE_CODE) LABEL_SIZE, EDT850T2.EDI_UPC LABEL_UPC" & vbCrLf _
                & ", EDT850T2.EDI_COLOR_NAME LABEL_COLOR, EDT850T2.EDI_STYLE_NAME LABEL_DESC, SIZE_QTY, sum(SIZE_QTY) over (PARTITION by EDT850T2.EDI_UPC) PPK_CNT" & vbCrLf _
                & " from SOTORDR2,SOTORDR1,EDT850T2, " & ASCMAIN1.sql & " ICTSTYL1" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and (SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & IIf(txtPO_IMPORT.Text <> "", " or  ORDR_CUST_PO like '" & txtPO_IMPORT.Text & "' )", ")") & vbCrLf _
                & "   and EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYL1.SIZE_CODE" & vbCrLf _
                & ", EDT850T2.EDI_SKU, EDT850T2.RETAIL_PRICE, EDT850T2.BUYER_CATALOG_NO" & vbCrLf _
                & ", EDT850T2.EDI_SIZE_DESC, EDT850T2.EDI_UPC" & vbCrLf _
                & ", EDT850T2.EDI_COLOR_NAME, EDT850T2.EDI_STYLE_NAME, SIZE_QTY)" & vbCrLf _
                & " where SIZE_QTY <> 0 AND PPK_CNT <> 0"

            ASCMAIN1.sql = "Select '" & LABEL_JOB_NO & "' LABEL_JOB_NO, ROWNUM LABEL_JOB_LNO, NULL ORDR_NO, NULL ORDR_LNO" & vbCrLf _
                & ", X.STYLE_CODE, X.COLOR_CODE, X.SIZE_CODE, X.ORDR_QTY" & vbCrLf _
                & ", X.LABEL_STYLE, X.LABEL_PRICE, X.LABEL_DGC" & vbCrLf _
                & ", nvl(X.LABEL_SIZE, ICVLUPC1.SIZE_CODE) LABEL_SIZE, NVL(X.LABEL_UPC,NVL(ICVLUPC1.UPC_CODE,ICTSTYC2.UPC_CODE)) LABEL_UPC, X.ORDR_QTY + " & CStr(LABELS_PER_SHEET) & " * SIGN(MOD(X.ORDR_QTY," & CStr(LABELS_PER_SHEET) & ")) - MOD(X.ORDR_QTY," & CStr(LABELS_PER_SHEET) & ") LABEL_QTY" & vbCrLf _
                & ", X.LABEL_COLOR, X.LABEL_DESC" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, nvl(ICTSTYC1.STYLE_COLOR_DESC , ICTCOLR1.COLOR_DESC) COLOR_DESC" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") X,ICTSTYL1,ICTCOLR1,ICTSTYC1,ICTSTYC2,ICVLUPC1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYC2.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC2.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYC2.COLOR_CODE_UPC = X.COLOR_CODE" & vbCrLf _
                & "   and ICVLUPC1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   and ICVLUPC1.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                & IIf(rowSOTORDR0.Item("ORDR_SOURCE") = "K", " and ICVLUPC1.SIZE_CODE (+) = X.LABEL_SIZE", "   and ICVLUPC1.UPC_CODE (+) = X.LABEL_UPC")

            Fill_Records("ICTLJOB2", "", True, ASCMAIN1.sql)
        End If

    End Sub

    Sub Fill_PO_ORDER_NO()
        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", Absx1.txtFor("PO_ORDER_NO").Text)
        If rowPOTORDR1 IsNot Nothing Then
            Absx1.txtFor("PO_REFERENCE").Text = rowPOTORDR1.Item("PO_REFERENCE") & ""
            Absx1.dteFor("PO_DATE_SHIP_BY").Value = rowPOTORDR1.Item("PO_DATE_SHIP_BY")
            Absx1.dteFor("PO_DATE_ETA").Value = rowPOTORDR1.Item("PO_DATE_ETA")

            Dim PO_ORDER_NO As String = Absx1.txtFor("PO_ORDER_NO").Text

            Dim LABELS_ACROSS As Integer = Val(rowICTULBL2.Item("LABELS_ACROSS") & "")
            Dim LABELS_DOWN As Integer = Val(rowICTULBL2.Item("LABELS_DOWN") & "")

            Dim LABELS_PER_SHEET As Integer = LABELS_ACROSS * LABELS_DOWN
            If LABELS_PER_SHEET = 0 Then LABELS_PER_SHEET = 24

            ASCMAIN1.sql = "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, SUM (POTORDR2.PO_QTY_ORD) ORDR_QTY" & vbCrLf _
                & ", NULL LABEL_STYLE, NULL LABEL_PRICE, NULL LABEL_DGC" & vbCrLf _
                & ", NULL LABEL_SIZE, NULL LABEL_UPC" & vbCrLf _
                & ", NULL LABEL_COLOR, NULL LABEL_DESC" & vbCrLf _
                & " from POTORDR2,POTORDR1" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE"

            ASCMAIN1.sql = "Select '" & LABEL_JOB_NO & "' LABEL_JOB_NO, ROWNUM LABEL_JOB_LNO, NULL ORDR_NO, NULL ORDR_LNO" & vbCrLf _
                & ", X.STYLE_CODE, X.COLOR_CODE, X.SIZE_CODE, X.ORDR_QTY" & vbCrLf _
                & ", X.LABEL_STYLE, X.LABEL_PRICE, X.LABEL_DGC" & vbCrLf _
                & ", X.LABEL_SIZE, X.LABEL_UPC, X.ORDR_QTY + " & CStr(LABELS_PER_SHEET) & " * SIGN(MOD(X.ORDR_QTY," & CStr(LABELS_PER_SHEET) & ")) - MOD(X.ORDR_QTY," & CStr(LABELS_PER_SHEET) & ") LABEL_QTY" & vbCrLf _
                & ", X.LABEL_COLOR, X.LABEL_DESC" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") X,ICTSTYL1,ICTCOLR1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE"

            Fill_Records("ICTLJOB2", "", True, ASCMAIN1.sql)
        End If

    End Sub

    Private Sub optLABELS_BY_ValueChanged(sender As Object, e As EventArgs) Handles optLABELS_BY.ValueChanged
        Setup_LABELS_BY()
    End Sub

    Sub Setup_LABELS_BY()
        If optLABELS_BY.Value = "P" Then
            Absx1.txtFor("ORDR_GROUP_NO").Text = ""
            Absx1.txtFor("ORDR_CUST_PO").Text = ""
            Absx1.dteFor("ORDR_SHIP_DATE").Value = DBNull.Value
            Absx1.dteFor("ORDR_CANCEL_DATE").Value = DBNull.Value
        Else
            Absx1.txtFor("PO_ORDER_NO").Text = ""
            Absx1.txtFor("PO_REFERENCE").Text = ""
            Absx1.dteFor("PO_DATE_SHIP_BY").Value = DBNull.Value
            Absx1.dteFor("PO_DATE_ETA").Value = DBNull.Value
        End If

        Absx1.txtFor("ORDR_GROUP_NO").ReadOnly = Not (optLABELS_BY.Value = "S")
        Absx1.txtFor("PO_ORDER_NO").ReadOnly = Not (optLABELS_BY.Value = "P")
    End Sub

    Sub Toggle_Selected()
        Dim dvw As DataView = DirectCast(grdICTLJOB2.DataSource, DataTable).DefaultView
        If chkShowSelectedOnly.Checked Then
            dvw.RowFilter = "LABEL_SEL = '1'"
        Else
            dvw.RowFilter = ""
        End If
    End Sub

    Private Sub chkShowSelectedOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowSelectedOnly.CheckedChanged
        Toggle_Selected()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click, cmdPO_refresh.Click
        Dim tbl As DataTable = dst.Tables("ICTLJOB2").Copy
        If optLABELS_BY.Value = "S" Then
            Fill_ORDR_GROUP_NO()
        Else
            Fill_PO_ORDER_NO()
        End If

        Calc_LABEL_QTY()
        For Each row As DataRow In tbl.Select("LABEL_SEL = '1' or LABEL_PRICE is Not Null")
            For Each updateRow As DataRow In dst.Tables("ICTLJOB2").Select("STYLE_CODE = '" & row("STYLE_CODE") & "' and COLOR_CODE = '" & row("COLOR_CODE") & "' and (LABEL_SIZE is null or LABEL_SIZE = '" & row("LABEL_SIZE") & "')")
                updateRow("LABEL_SEL") = row("LABEL_SEL")
                updateRow("LABEL_PRICE") = row("LABEL_PRICE")
            Next
        Next

    End Sub

    Sub Calc_LABEL_QTY()

        Dim LABEL_QTY_CALC As String = optLABEL_QTY_CALC.Value
        Dim LABEL_QTY_EXTRA As Integer = Val(numExtra.Value & "")

        Dim LABELS_ACROSS As Integer = Val(Absx1.numFor("LABELS_ACROSS").Value & "")
        Dim LABELS_DOWN As Integer = Val(Absx1.numFor("LABELS_DOWN").Value & "")

        For Each rowICTLJOB2 As DataRow In dst.Tables("ICTLJOB2").Select("")
            Dim ORDR_QTY As Integer = Val(rowICTLJOB2.Item("ORDR_QTY") & "")

            Dim LABEL_QTY As Integer = 0

            If optLABEL_QTY_CALC.Value = "S" Then
                LABEL_QTY = ORDR_QTY + LABEL_QTY_EXTRA
                Dim LABEL_CALC As Integer = LABELS_ACROSS * LABELS_DOWN
                If LABEL_QTY Mod LABEL_CALC <> 0 Then
                    LABEL_QTY += LABEL_CALC - LABEL_QTY Mod LABEL_CALC
                End If

            ElseIf optLABEL_QTY_CALC.Value = "R" Then
                LABEL_QTY = ORDR_QTY + LABEL_QTY_EXTRA
                If Absx1.chkFor("LABEL_SPACER").Checked Then
                    LABEL_QTY += LABELS_ACROSS
                End If
                If LABEL_QTY Mod LABELS_ACROSS <> 0 Then
                    LABEL_QTY += LABELS_ACROSS - LABEL_QTY Mod LABELS_ACROSS
                End If

            ElseIf optLABEL_QTY_CALC.Value = "O" Then
                LABEL_QTY = ORDR_QTY + LABEL_QTY_EXTRA
                If Absx1.chkFor("LABEL_SPACER").Checked Then
                    LABEL_QTY += 1
                End If

            ElseIf optLABEL_QTY_CALC.Value = "X" Then
                LABEL_QTY = LABEL_QTY_EXTRA
                If Absx1.chkFor("LABEL_SPACER").Checked Then
                    LABEL_QTY += 1
                End If
            End If

            rowICTLJOB2.Item("LABEL_QTY") = LABEL_QTY
        Next
    End Sub

    Private Sub optLABEL_QTY_CALC_ValueChanged(sender As Object, e As EventArgs) Handles optLABEL_QTY_CALC.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Calc_LABEL_QTY()
        Absx1.chkFor("LABEL_SPACER").Visible = Not (optLABEL_QTY_CALC.Value = "S")
    End Sub
End Class