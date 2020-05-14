Public Class POFWORK1
    'Load_Events

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name

    Dim WORK_NO As String
    Dim ORDR_CUST_PO As String      ' Customer's PO No

    Dim rowPOTWORK1 As DataRow
    Dim rowARTCUST1 As DataRow      ' ARTCUST1 for the Sold-To
    Dim rowICTSTYL1 As DataRow
    Dim WORK_LNOs As New List(Of Int64) ' list of WORK_LNOs that are deleted

    Dim COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select POTWORK1.* from POTWORK1 where WORK_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTWORKX", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "POTWORK1", "*", 1)

            ASCMAIN1.sql = "Select POTWORK2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from POTWORK2,ICTSTYL1,ICTCOLR1" _
            & " where ICTSTYL1.STYLE_CODE = POTWORK2.STYLE_CODE" _
            & "   and ICTCOLR1.COLOR_CODE = POTWORK2.COLOR_CODE"
            Create_TDA(.Tables.Add, "POTWORK2", "**", 1)

            With .Tables("POTWORK2").Columns
                .Add("WORK_AMT", GetType(System.Decimal), "ISNULL(WORK_QTY,0) * ISNULL(STYLE_COST,0)")
                .Add("WORK_AMT_OPEN", GetType(System.Decimal), "ISNULL(WORK_QTY_OPEN,0) * ISNULL(STYLE_COST,0)")
                .Add("WORK_AMT_USED", GetType(System.Decimal), "ISNULL(WORK_QTY_USED,0) * ISNULL(STYLE_COST,0)")
                .Add("WORK_AMT_CANC", GetType(System.Decimal), "ISNULL(WORK_QTY_CANC,0) * ISNULL(STYLE_COST,0)")
            End With

            ASCMAIN1.sql = "Select POTWORK3.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from POTWORK3,ICTSTYL1,ICTCOLR1" _
            & " where ICTSTYL1.STYLE_CODE = POTWORK3.STYLE_CODE" _
            & "   and ICTCOLR1.COLOR_CODE = POTWORK3.COLOR_CODE"
            Create_TDA(.Tables.Add, "POTWORK3", "**", 1)

            With .Tables("POTWORK3").Columns
                .Add("PROD_AMT", GetType(System.Decimal), "ISNULL(PROD_QTY,0) * ISNULL(STYLE_COST,0)")
                .Add("PROD_AMT_OPEN", GetType(System.Decimal), "ISNULL(PROD_QTY_OPEN,0) * ISNULL(STYLE_COST,0)")
                .Add("PROD_AMT_COMP", GetType(System.Decimal), "ISNULL(PROD_QTY_COMP,0) * ISNULL(STYLE_COST,0)")
                .Add("PROD_AMT_CANC", GetType(System.Decimal), "ISNULL(PROD_QTY_CANC,0) * ISNULL(STYLE_COST,0)")
            End With

            With .Tables.Add("POTWORKT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("KEY")}
            End With

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            'ASCMAIN1.sql = "Select * from POTWORK7 where WORK_GROUP_NO = :PARM1 " & vbCrLf _
            '    & "   and POTWORK7.STYLE_CODE = :PARM2 " & vbCrLf _
            '    & "   and POTWORK7.COLOR_CODE = :PARM3" & vbCrLf _
            '    & "   and POTWORK7.ALLO_BATCH_NO is Null"
            'Create_TDA(.Tables.Add, "POTWORK7", "**", 0, True, "VVV", 1)

            ASCMAIN1.sql = "Select * from SOTWORK1 where WO_REF_TYPE = 'R' and WO_REF_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTWORK1", "**", 0, , "V", 1)
            ASCMAIN1.sql = "Select * from SOTWORK2 where WO_NO in " _
                & " (Select WO_NO from SOTWORK1 where WO_REF_TYPE = 'R' and WO_REF_NO = :PARM1)"
            Create_TDA(.Tables.Add, "SOTWORK2", "**", 0, , "V", 1)

            ASCMAIN1.sql = "Select * from ICTSTAT2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 2, False)
        End With

        grdPOTWORKX.DataSource = dst.Tables("POTWORKX")
        grdPOTWORK2.DataSource = dst.Tables("POTWORK2")
        grdPOTWORK3.DataSource = dst.Tables("POTWORK3")
        grdPOTWORKT.DataSource = dst.Tables("POTWORKT")
        grdICTSTAT2.DataSource = dst.Tables("ICTSTAT2")

        grdPOTWORKX.DisplayLayout.UseFixedHeaders = True
        With grdPOTWORKX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"WORK_NO", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With


        With grdPOTWORK2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"WORK_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdPOTWORK2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"STYLE_CODE", "COLOR_CODE", "WORK_QTY", "WORK_QTY_OPEN", "STYLE_COST", "WORK_PRIORITY", "WORK_PRIORITY_DATE", "WORK_DEMAND_DATE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"WORK_QTY", "WORK_QTY_OPEN", "WORK_QTY_ALLO", "WORK_QTY_USED", "WORK_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"WORK_AMT", "WORK_AMT_OPEN", "WORK_AMT_USED", "WORK_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With



        With grdPOTWORK3.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"WORK_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdPOTWORK3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"STYLE_CODE", "COLOR_CODE", "PROD_QTY", "PROD_QTY_OPEN", "STYLE_COST"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"PROD_QTY", "PROD_QTY_OPEN", "PROD_QTY_COMP", "PROD_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"PROD_AMT", "PROD_AMT_OPEN", "PROD_AMT_COMP", "PROD_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With

        Create_Summary(grdPOTWORKX, "WORK_NO", "Count")

        Create_Summary(grdPOTWORK2, "WORK_LNO", "Count")
        Create_Summary(grdPOTWORK2, New String() {"WORK_QTY", "WORK_QTY_OPEN", "WORK_QTY_USED", "WORK_QTY_CANC", "WORK_AMT"})
        Create_Summary(grdPOTWORK3, "WORK_LNO", "Count")
        Create_Summary(grdPOTWORK3, New String() {"PROD_QTY", "PROD_QTY_OPEN", "PROD_QTY_COMP", "PROD_QTY_CANC", "PROD_AMT"})

        With dst.Tables("POTWORKT").Rows
            .Add(New Object() {1, "WORK", 0, 0})
            .Add(New Object() {2, "Open", 0, 0})
            .Add(New Object() {3, "Allo", 0, 0})
            .Add(New Object() {4, "Used", 0, 0})
            .Add(New Object() {5, "Canc", 0, 0})
        End With
        Sort_grdColumns(grdPOTWORKT, "KEY", True)

        Show_Filter(grdPOTWORKX, True)
        grdPOTWORKX.DisplayLayout.GroupByBox.Hidden = False

        'SplitContainer1.Panel2Collapsed = True
        tabInfo.Tabs("Usage").Visible = True

        Check_InquiryMode()

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFWORKI")
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
                        ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                        If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                            EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                        End If
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("POTWORK1", CUST_CODE) Then Exit Sub
                End If

            Case "Edit", "View"

                CUST_CODE = ""
                WORK_NO = ""

                If Absx1.txtFor("WORK_NO").Text = "" Then
                    EMsg &= vbCr & "No Reservation No Specified"
                Else
                    WORK_NO = Absx1.txtFor("WORK_NO").Text
                    rowPOTWORK1 = LookUp("POTWORK1", WORK_NO)
                    If rowPOTWORK1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Reservation No " & WORK_NO
                    Else
                        CUST_CODE = rowPOTWORK1.Item("CUST_CODE")
                        If rowPOTWORK1.Item("WORK_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowPOTWORK1.Item("WORK_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "Reservation No " & WORK_NO & " has been Cancelled"
                                Case "D"
                                    EMsg &= vbCr & "Reservation No " & WORK_NO & " has been Deleted"
                                Case Else ' such as "F"
                                    EMsg &= vbCr & "Reservation No " & WORK_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("POTWORK1", WORK_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("POTWORK1", CUST_CODE) Then Exit Sub
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

                If Absx1.txtFor("SREP_CODE").Text = "" Then
                    EMsg &= vbCr & "Sales Rep is required"
                Else
                    If LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Sales Rep"
                    End If
                End If

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "Warehouse is required"
                Else
                    If LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Warehouse"
                    End If
                End If

                If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                    EMsg &= vbCr & "Customer PO is required"
                End If
                If grdPOTWORK2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Reservation"
                Else
                    If Val(dst.Tables("POTWORK2").Compute("COUNT(WORK_LNO)", "WORK_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Items on Reservation with Qty >0"
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

                ASCMAIN1.sql = "Select Count (*) from POTWORK2 where WORK_QTY_USED <> 0"
                ASCMAIN1.sql &= " and WORK_NO = '" & WORK_NO & "'"

                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Reservation has been Used"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Reservation as Deleted", _
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If


            Case "Cancel Balance"
                If EMsg = "" Then
                    If MsgBox("Do you want to Cancel (the remaining open balance on) this Reservation", _
                               vbYesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
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

            Case "Cancel Balance"
                Cancel_Order()
                Mode_Settings(False)

            Case "Work Orders"
                Using F As New TAC.SOFWORK1(Me, "R", WORK_NO, (EntryMode = "V" Or InquiryMode), _
                                            Absx1.txtFor("CUST_CODE").Text, _
                                            Absx1.txtFor("ORDR_CUST_PO").Text,
                                            Absx1.dteFor("ORDR_SHIP_DATE").Value, _
                                             Absx1.dteFor("ORDR_CANCEL_DATE").Value, _
                                            "Work Orders relating to Sales Reservation " & WORK_NO)
                    F.ShowDialog()
                End Using
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowPOTWORK1.Item("WORK_STATUS") & "" = "O" Then
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

                .Items("Cancel Balance").Settings.Enabled = iScreenMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Print").Visible = False ' ScreenMode
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Delete").Visible = (EntryMode = "E")
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)

                .Items("Cancel Balance").Visible = (EntryMode = "E")

                .Items("Work Orders").Text = "Work Orders" & IIf(dst.Tables("SOTWORK1").Rows.Count = 0, "", " (" & CStr(dst.Tables("SOTWORK1").Rows.Count) & ")")
                .Items("Work Orders").Visible = ScreenMode And Not (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")

            End With

            .Groups("Totals").Visible = ScreenMode
        End With

        lblStatus.Visible = ScreenMode

        grdPOTWORKX.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), InquiryMode Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")))
        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))

        If ScreenMode Then
            If EntryMode = "V" Then
                grdPOTWORK2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdPOTWORK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdPOTWORK2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdPOTWORK2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdPOTWORK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdPOTWORK2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

                If EntryMode <> "E" Then
                    grdPOTWORK2.DisplayLayout.Bands(0).Columns("X").Hidden = True
                Else
                    grdPOTWORK2.DisplayLayout.Bands(0).Columns("X").Hidden = False
                End If
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("WORK_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""

        CUST_CODE = ""
        WORK_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTWORK1", "POTWORK2", "SOTWORK1", "SOTWORK2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Load_POTWORKX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                WORK_NO = ASCMAIN1.Next_Control_No("WORK_NO")
            Else
                WORK_NO = ASCMAIN1.Next_Control_No("POTWORK1.WORK_NO")
            End If

            rowPOTWORK1 = dst.Tables("POTWORK1").NewRow
            With rowPOTWORK1
                .Item("WORK_NO") = WORK_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                .Item("WORK_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
                If WHSE_CODE = "" Then WHSE_CODE = ""
                .Item("WHSE_CODE") = WHSE_CODE
                '  .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
                .Item("WORK_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
            End With
            dst.Tables("POTWORK1").Rows.Add(rowPOTWORK1)

        Else
            rowPOTWORK1 = Fill_Record("POTWORK1", WORK_NO)
        End If

        CUST_CODE = rowPOTWORK1.Item("CUST_CODE")
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)

        Fill_Records("POTWORK2", WORK_NO)
        Sort_grdColumns(grdPOTWORK2, "WORK_LNO")

        'Fill_Records("SOTWORK1", WORK_NO)
        'Fill_Records("SOTWORK2", WORK_NO)

        lblINIT_DATE.Text = "Entered on " & Format(rowPOTWORK1.Item("INIT_DATE"), "MM/dd/yyyy")

        If EntryMode = "N" Then
            lblStatus.Text = "New Order"
        Else
            Select Case rowPOTWORK1.Item("WORK_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
            End Select
        End If

        With grdPOTWORK2.DisplayLayout.Bands(0)
            If (EntryMode = "E" Or EntryMode = "N") Then
                .Columns("WORK_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                If EntryMode = "E" Then
                    .Columns("WORK_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("WORK_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Else
                .Columns("WORK_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("WORK_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdPOTWORK2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            End With
            grdPOTWORK2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, False)
        Else
            With grdPOTWORK2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            grdPOTWORK2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, True)
        End If

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
        Dependent_Updates(-1, WORK_NO)
        For Each TABLE_NAME As String In New String() _
            {"POTWORK1", "POTWORK2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where WORK_NO = '" & WORK_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        If EntryMode <> "N" Then Delete_Records()

        If EntryMode = "N" Then
            rowPOTWORK1.Item("ORDR_ORIG_SHIP_DATE") = rowPOTWORK1.Item("ORDR_SHIP_DATE")
            rowPOTWORK1.Item("ORDR_ORIG_CANCEL_DATE") = rowPOTWORK1.Item("ORDR_CANCEL_DATE")
        End If

        INIT_LAST("POTWORK1", False, , True)
        Dim sqldelete As String = "WORK_NO = '" & WORK_NO & "'"
        Update_Record_TDA("POTWORK1", sqldelete)
        Update_Record_TDA("POTWORK2", sqldelete)
        Dependent_Updates(1, WORK_NO)

        'Update_Record_TDA("SOTWORK1")
        'Update_Record_TDA("SOTWORK2")

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

            Case "WORK_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and POTWORK1.WORK_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and POTWORK1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and POTWORK1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If
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

                Absx1.txtFor("WORK_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTWORK1"
            E.COLUMN_NAME = "WORK_NO"
            E.CODE_VALUE = Absx1.txtFor("WORK_NO").Text
            E.DESC_VALUE = "Reservation"
            E.ATTACHMENT_NOTES = ""
            'If rowPOTWORK1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTWORKX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdPOTWORK2, "BB", "Style Status Inquiry", "Style Multi-Color")
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
            Case "grdPOTWORK2"
                tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

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
            Case "Style Multi-Color"
                Using F As New TAC.ICFSTYCX
                    F.STYLE_CODE = ""
                    F.ShowDialog()
                    If F.STYLE_CODE <> "" Then
                        Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"))
                    End If
                End Using

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
                    Load_POTWORKX()
                End If

            Case "ORDR_CUST_PO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode _
                       And Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "WORK_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_POTWORKX()

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
                Load_POTWORKX()
            Case "WORK_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_POTWORKX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE = "" Then
            ASCMAIN1.sql = "Select * from POTWORK1 where WORK_STATUS = 'O'"
            Fill_Records("POTWORKX", "", , ASCMAIN1.sql)
            grdPOTWORKX.Text = "Open Reservations"
            Sort_grdColumns(grdPOTWORKX, "WORK_NO".ToLower)
        Else
            ASCMAIN1.sql = "Select * from POTWORK1 where WORK_STATUS = 'O'" _
                & " and CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("POTWORKX", "", , ASCMAIN1.sql)
            grdPOTWORKX.Text = "Open Reservations associated with " & CUST_CODE
            Sort_grdColumns(grdPOTWORKX, "WORK_NO".ToLower)
        End If
        grdPOTWORKX.Visible = True
    End Sub

    Sub Print_Record()
        ' NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM
        ' AND SHOULD BE USING THE DATALAYER OF SORUSED1

        'Fill_Records("SOTSVIA1", USED_CODE)

        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Dim RPT As String = "SORUSED1" ' unneccesary if Report Name is Like Form Name
        'Generate_Report(RPT, "USEDper Invoice Report", , , , , False)
        'Print_Report_End()


        Dim REPORTFILE As String = "SORUSED1"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it
        'REPORTS(REPORTFILE).Fill_Records_RPT(New String() {"USEDPER_CTL_NO = '" & USEDPER_CTL_NO & "'"})

        'To fill the report's dataset with data from this form's dataset:
        With REPORTS(REPORTFILE).clsASCBASE1
            .EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SOTPPDI1", "SOTPPDI2", "SOTPPDI3", "SOTINVH1", "SOTSVIA1"}
                .dst.Tables(TABLE_NAME).Rows.Clear()
                Dim SQL As String = ""
                If TABLE_NAME = "SOTINVH1" Then
                    SQL = "WORK_NO = '" & WORK_NO & "'"
                End If

                For Each row As DataRow In dst.Tables(TABLE_NAME).Select(SQL)
                    Dim rowr As DataRow = .dst.Tables(TABLE_NAME).NewRow
                    If TABLE_NAME = "SOTPPDI2" Or TABLE_NAME = "SOTPPDI3" Or TABLE_NAME = "SOTINVH1" Then

                        For I As Integer = 0 To .dst.Tables(TABLE_NAME).Columns.Count - 1
                            Dim COLUMN_NAME As String = .dst.Tables(TABLE_NAME).Columns(I).ColumnName
                            rowr.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                        Next
                    Else
                        rowr.ItemArray = row.ItemArray
                    End If
                    .dst.Tables(TABLE_NAME).Rows.Add(rowr)
                Next
            Next
            .EnforceConstraints(True)
        End With
        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .Generate_Report("SORUSED1", "USEDper Invoice Report", , True, , , , , False)
            .Print_Report_End()
        End With

    End Sub

    Private Sub grdPOTWORKX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTWORKX.DoubleClickRow
        Absx1.txtFor("WORK_NO").Text = e.Row.Cells("WORK_NO").Value
        Click_Command("View")
    End Sub

    Sub Update_ICTSTAT2(STYLE_CODE As String, COLOR_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVVNNNNNN", _
                           New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, _
                                         0, 0, 0, _
                                         QTY, 0, 0}, _
                           New String() {"STYLE_CODE_IN", "COLOR_CODE_IN", "WHSE_CODE_IN", _
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in", _
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Sub Cancel_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        Dim EMsg As String = ""
        If EntryMode = "E" Then
            Cancel_Order_1(WORK_NO)
            EMsg = "Reservation " & WORK_NO & " has been Cancelled"
        End If

        'ASCDATA1.ExecuteSP("SOPWORK0_G", "V", New Object() {WORK_GROUP_NO}, New String() {"WORK_GROUP_NO_IN"})
        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Cancel_Order_1(WORK_NO As String)
        Dependent_Updates(-1, WORK_NO)

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is Select * from POTWORK2 where WORK_NO = '" & WORK_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update POTWORK2" _
            & "    Set WORK_QTY_CANC = NVL(WORK_QTY_CANC,0) + NVL(R1.WORK_QTY_OPEN,0)" _
            & "      , WORK_QTY_OPEN = 0" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()
        ', WORK_STATUS = 'C'
        ASCMAIN1.sql = "Update POTWORK1 Set WORK_STATUS = :PARM1" _
            & " where WORK_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"C", WORK_NO})
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        Dim EMsg As String = ""

        BeginTrans()

        If EntryMode = "E" Then
            Delete_Order_1(WORK_NO)
            EMsg = "Reservation No " & WORK_NO & " has been marked as Deleted"
        End If

        CommitTrans(EMsg)
        'ASCDATA1.ExecuteSP("SOPWORK0_G", "V", New Object() {WORK_GROUP_NO}, New String() {"WORK_GROUP_NO_IN"})
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(WORK_NO As String)
        Dependent_Updates(-1, WORK_NO)

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select * from POTWORK2" & vbCrLf _
            & "     where WORK_NO = '" & WORK_NO & "' for Update;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update POTWORK2" & vbCrLf _
            & "    Set WORK_QTY_CANC = NVL(WORK_QTY_CANC,0) + NVL(R1.WORK_QTY_OPEN,0)" & vbCrLf _
            & "   , WORK_QTY_OPEN = 0" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update POTWORK1 Set WORK_STATUS = :PARM1" _
            & " where WORK_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"D", WORK_NO})
    End Sub

    Sub Dependent_Updates(S As Integer, WORK_NO As String)

        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from POTWORK1 where WORK_NO = '" & WORK_NO & "'"
        Dim rowPOTWORK1 As DataRow = ASCDATA1.GetDataRow
        Dim WHSE_CODE As String = rowPOTWORK1.Item("WHSE_CODE") & ""
        If WHSE_CODE = "" Then WHSE_CODE = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""

        ASCMAIN1.sql = "Select * from POTWORK2 where WORK_NO = '" & WORK_NO & "'"
        For Each rowPOTWORK2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowPOTWORK2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowPOTWORK2.Item("COLOR_CODE")

            QTY_TO_COMMIT = Val(rowPOTWORK2.Item("WORK_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowPOTWORK2.Item("STYLE_CODE")
                COLOR_CODE = rowPOTWORK2.Item("COLOR_CODE")
                Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, S * QTY_TO_COMMIT)
            End If
        Next
    End Sub

    Sub Display_Totals()

        Dim KEY As Int32 = 0
        For Each SFX As String In New String() {"", "OPEN", "USED", "CANC"}
            If SFX <> "" Then SFX = "_" & SFX
            KEY += 1
            Dim rowPOTWORKT As DataRow = dst.Tables("POTWORKT").Rows.Find(KEY)
            rowPOTWORKT.Item("QTY") = Val(dst.Tables("POTWORK2").Compute("SUM(WORK_QTY" & SFX & ")", "") & "")
            rowPOTWORKT.Item("AMT") = Val(dst.Tables("POTWORK2").Compute("SUM(WORK_AMT" & SFX & ")", "") & "")
        Next
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdPOTWORK2.ActiveRow
            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = ""
                    If Trim(.Cells("STYLE_CODE").Value & "") <> "" Then
                        STYLE_CODE = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    End If
                    Cancel = (STYLE_CODE = "")

                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Case "WORK_QTY"
                    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(.Cells("WORK_QTY").Value & "") = "" Then
                        MsgBox("Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdPOTWORK2.ActiveCell = grdPOTWORK2.ActiveRow.Cells("WORK_QTY")
                        Exit Sub
                    End If
                    If Val(.Cells("WORK_QTY").Value & "") < 0 Then
                        MsgBox("Qty May Not be Negative", vbOKOnly, "Invalid Quantity")
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim EMsg As String = ""
        If STYLE_CODE_z = "" Then Return ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            EMsg = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                EMsg = "Item Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                EMsg = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                EMsg = "Item does not have a valid Division Code" & vbCrLf
            End If
        End If

        If EMsg = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If EMsg <> "" And grdPOTWORK2.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

#Region "grdPOTWORK2"

    Private Sub grdPOTWORK2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWORK2.AfterCellUpdate
        With grdPOTWORK2.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value)
                    If STYLE_CODE <> "" Then
                        .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        If COLOR_CODEs.Count = 1 Then
                            .Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                        End If

                        Fill_Records("ICTSTAT2", New String() {STYLE_CODE, "AST"})
                        grdICTSTAT2.Text = "Style Status for " & STYLE_CODE & ":" & .Cells("COLOR_CODE").Value
                        grdICTSTAT2.Visible = True
                    End If

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & ""
                    If COLOR_CODE <> "" Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        If rowICTCOLR1 IsNot Nothing Then
                            .Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

                Case "WORK_QTY"
                    .Cells("WORK_QTY_OPEN").Value = .Cells("WORK_QTY").Value

                Case "WORK_QTY_OPEN"
                    .Cells("WORK_QTY_CANC").Value _
                        = Val(.Cells("WORK_QTY").Value & "") _
                        - Val(.Cells("WORK_QTY_USED").Value & "") _
                        - Val(.Cells("WORK_QTY_OPEN").Value & "")
                    If Val(.Cells("WORK_QTY_CANC").Value) < 0 Then
                        .Cells("WORK_QTY_CANC").Value = 0
                    End If
            End Select
        End With
    End Sub

    Private Sub grdPOTWORK2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTWORK2.AfterRowActivate


        If grdPOTWORK2.ActiveRow Is Nothing OrElse grdPOTWORK2.ActiveRow.IsAddRow OrElse Not grdPOTWORK2.ActiveRow.IsDataRow Then
            grdICTSTAT2.Visible = False
        Else
            Dim STYLE_CODE As String = grdPOTWORK2.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdPOTWORK2.ActiveRow.Cells("COLOR_CODE").Value
            Fill_Records("ICTSTAT2", New String() {STYLE_CODE, COLOR_CODE})
            grdICTSTAT2.Text = "Style Status for " & STYLE_CODE & ":" & COLOR_CODE
            grdICTSTAT2.Visible = True
        End If

        If Trim(grdPOTWORK2.ActiveRow.Cells("STYLE_CODE").Value & "") = "" And _
            (grdPOTWORK2.ActiveCell Is Nothing OrElse _
             (grdPOTWORK2.ActiveCell.Column.Key <> "STYLE_CODE")) _
        Then
            grdPOTWORK2.ActiveCell = grdPOTWORK2.ActiveRow.Cells("STYLE_CODE")
            Exit Sub
        End If

        If grdPOTWORK2.ActiveRow.IsAddRow Then
            grdPOTWORK2.DisplayLayout.Bands(0).Columns("WORK_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdPOTWORK2.DisplayLayout.Bands(0).Columns("WORK_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            If grdPOTWORK2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                grdPOTWORK2.ActiveCell = grdPOTWORK2.ActiveRow.Cells("STYLE_CODE")
            End If
        Else
            With grdPOTWORK2.DisplayLayout.Bands(0)
                Validate_Style(grdPOTWORK2.ActiveRow.Cells("STYLE_CODE").Value & "")

                If Val(grdPOTWORK2.ActiveRow.Cells("WORK_QTY_USED").Value & "") <> 0 _
                Or Val(grdPOTWORK2.ActiveRow.Cells("WORK_QTY_CANC").Value & "") <> 0 _
                Then
                    .Columns("WORK_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("WORK_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("WORK_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("WORK_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End With
        End If

    End Sub

    Private Sub grdPOTWORK2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTWORK2.AfterRowsDeleted
        Display_Totals()

        If grdPOTWORK2.Rows.Count = 0 Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = ""
        End If
    End Sub

    Private Sub grdPOTWORK2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTWORK2.AfterRowUpdate
        Display_Totals()

        If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTSTYL1.Item("SALES_DIVISION_CODE")
        End If
    End Sub

    Private Sub grdPOTWORK2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdPOTWORK2.BeforeCellUpdate

    End Sub

    Private Sub grdPOTWORK2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTWORK2.BeforeExitEditMode
        If grdPOTWORK2.ActiveCell IsNot Nothing Then
            With grdPOTWORK2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdPOTWORK2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTWORK2.BeforeRowsDeleted

        WORK_LNOs.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Val(grow.Cells("WORK_QTY_USED").Value & "") <> 0 _
            Or Val(grow.Cells("WORK_QTY_CANC").Value & "") <> 0 _
            Then
                MsgBox("Cannot Delete a Line if it has ever been " & vbCr & "Used Or Cancelled" & vbCr & "Use the Cancel Button (x)")
                e.Cancel = True
                Exit Sub
            End If

            WORK_LNOs.Add(grow.Cells("WORK_LNO").Value)
        Next
    End Sub

    Private Sub grdPOTWORK2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTWORK2.BeforeRowUpdate

        Validate_Columns("STYLE_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("COLOR_CODE", e.Cancel)
        End If
        If Not e.Cancel Then
            Validate_Columns("WORK_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("WORK_NO").Value = WORK_NO
            Dim WORK_LNO As Int64 = Val(dst.Tables("POTWORK2").Compute("MAX(WORK_LNO)", "") & "") + 1
            e.Row.Cells("WORK_LNO").Value = WORK_LNO
        End If
    End Sub

    Private Sub grdPOTWORK2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTWORK2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "X"
                    If Val(.Cells("WORK_QTY_CANC").Value) <> 0 Then
                        If MsgBox("Restore Cancelled Qty of " & .Cells("WORK_QTY_CANC").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowPOTWORKR As DataRow In dst.Tables("POTWORKR").Select("WORK_LNO = " & .Cells("WORK_LNO").Value)
                                rowPOTWORKR.Item("WORK_QTY_OPEN") = Val(rowPOTWORKR.Item("WORK_QTY_OPEN") & "") + Val(rowPOTWORKR.Item("WORK_QTY_CANC") & "")
                                rowPOTWORKR.Item("WORK_QTY_CANC") = 0
                            Next
                        End If
                        .Cells("WORK_QTY_OPEN").Value = Val(.Cells("WORK_QTY_OPEN").Value & "") + Val(.Cells("WORK_QTY_CANC").Value & "")
                        ' grdSOWWORK2_AfterColUpdate(.Cells("WORK_QTY_OPEN").position)
                        .Update()
                    Else
                        If MsgBox("Cancel Remaining Qty Open of " & .Cells("WORK_QTY_OPEN").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowPOTWORKR As DataRow In dst.Tables("POTWORKR").Select("WORK_LNO = " & .Cells("WORK_LNO").Value)
                                rowPOTWORKR.Item("WORK_QTY_OPEN") = 0
                                Dim WORK_QTY_CANC As Int64 = Val(rowPOTWORKR.Item("WORK_QTY") & "") _
                                                           - Val(rowPOTWORKR.Item("WORK_QTY_USED") & "") _
                                                           - Val(rowPOTWORKR.Item("WORK_QTY_OPEN") & "")
                                rowPOTWORKR.Item("WORK_QTY_CANC") = IIf(WORK_QTY_CANC < 0, 0, WORK_QTY_CANC)
                            Next
                        End If
                        .Cells("WORK_QTY_OPEN").Value = "0"
                        ' grdSOWWORK2_AfterColUpdate(.Cells("WORK_QTY_OPEN").position)
                        grdPOTWORK2.ActiveRow.Update()
                    End If

                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdPOTWORK2, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE IN (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE ='" & grdPOTWORK2.ActiveRow.Cells("STYLE_CODE").Value & "')"

                    grdClickCellButton(grdPOTWORK2, sql_where)
            End Select
        End With

    End Sub
#End Region

    Sub Load_Events()
        '    grdEvents.RemoveAll
        '    Call Load_Events_1("Entered", "INIT_DATE")
        '    Call Load_Events_1("Modified", "LAST_DATE")
        '    Call Load_Events_1("Released", "WORK_DATE_REL")
        '    Call Load_Events_1("ALLO Ticket", "WORK_DATE_ALLO_PRT")
        '    Call Load_Events_1("Packed", "WORK_DATE_PACKED")
        '    Call Load_Events_1("USEDped", "WORK_DATE_USEDPED")
        '    Call Load_Events_1("Invoice", "WORK_INV_DATE")
        '    Call Load_Events_1("Invoice Prt", "WORK_DATE_INV_PRT")
        '    Call Load_Events_1("Updated", "WORK_DATE_UPDATED")
        '    Call Load_Events_1("Cancelled", "WORK_DATE_CANCELLED")
    End Sub

    Sub Add_Colors(STYLE_CODE As String, tbl As DataTable)
        If tbl.Select("ISNULL(QTY,0)<>0").Length = 0 Then
            MsgBox("No Qty's Entered", MsgBoxStyle.OkOnly, "Cannot Add Colors")
            Exit Sub
        End If

        For Each rowICTCOLRM As DataRow In tbl.Select("ISNULL(QTY,0)<>0", "COLOR_CODE")
            grdPOTWORK2.DisplayLayout.Bands(0).AddNew()
            With grdPOTWORK2.ActiveRow
                .Cells("STYLE_CODE").Value = STYLE_CODE
                .Cells("COLOR_CODE").Value = rowICTCOLRM.Item("COLOR_CODE")
                .Cells("WORK_QTY").Value = rowICTCOLRM.Item("QTY")
                .Update()
            End With
        Next
        Sort_grdColumns(grdPOTWORK2, "WORK_LNO")
    End Sub
End Class