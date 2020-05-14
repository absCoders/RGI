Imports Infragistics.Win.UltraWinGrid

Public Class POFLTRC1
    Dim rowPOTLTRC1 As DataRow
    Dim LC_CTL_NO As String
    Dim LC_CTL_NO_new As String
    Dim STATUS_CODE As String

    Dim sqlPOTLTRCX As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "POTLTRCI" Then
            InquiryMode = True
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
            sqlPOTLTRCX = "Select POTLTRC1.*,APTVEND1.VEND_NAME" & vbCrLf _
                & " from POTLTRC1,APTVEND1" & vbCrLf _
                & " where APTVEND1.VEND_CODE = POTLTRC1.VEND_CODE"
            ASCMAIN1.sql = sqlPOTLTRCX ' & "  and POTLTRC1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "POTLTRCX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "POTLTRC1", "*")
            With .Tables("POTLTRC1")
                .Columns.Add("LC_CANC_CALC", GetType(System.Decimal))
                .Columns.Add("LC_OPEN_CALC", GetType(System.Decimal))
                .Columns.Add("PYMTS", GetType(System.Int32))
                .Columns.Add("LC_AMT_MAX", GetType(System.Decimal), "LC_AMT * 1.05")
                .Columns.Add("LC_AMT_MIN", GetType(System.Decimal), "LC_AMT * 0.95")
                .Columns.Add("LC_PO", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "Select POTORDR1.PO_ORDER_NO" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_ETA, POTORDR1.LC_CTL_NO" & vbCrLf _
                & ", X.PO_AMT_ORD, X.PO_AMT_OPN, X.PO_AMT_SHP, X.PO_AMT_REC" & vbCrLf _
                & " from POTORDR1, (Select PO_ORDER_NO" & vbCrLf _
                & ", Sum (PO_COST * PO_QTY_ORD) PO_AMT_ORD" & vbCrLf _
                & ", Sum (PO_COST * PO_QTY_OPN) PO_AMT_OPN" & vbCrLf _
                & ", Sum (PO_COST * PO_QTY_SHP) PO_AMT_SHP" & vbCrLf _
                & ", Sum (PO_COST * PO_QTY_REC) PO_AMT_REC" & vbCrLf _
                & " from POTORDR2 group by PO_ORDER_NO) X" & vbCrLf _
                & " where ((POTORDR1.VEND_CODE = :PARM1 And POTORDR1.LC_CTL_NO is Null And (POTORDR1.PO_STATUS = 'O' OR POTORDR1.PO_DATE_ORDERED > SYSDATE -300))" & vbCrLf _
                & "      Or POTORDR1.LC_CTL_NO = :PARM2)" & vbCrLf _
                & " and X.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTLTRCP", "**", 0, False, "VV")
            .Tables("POTLTRCP").Columns.Add("SEL")
            .Tables("POTLTRCP").Columns("SEL").DefaultValue = "0"
            '.Tables("POTLTRCP").Columns.Add("PO_AMT_ORD", GetType(System.Decimal))
            '.Tables("POTLTRCP").Columns.Add("PO_AMT_OPN", GetType(System.Decimal))
            '.Tables("POTLTRCP").Columns.Add("PO_AMT_SHP", GetType(System.Decimal))
            '.Tables("POTLTRCP").Columns.Add("PO_AMT_REC", GetType(System.Decimal))

            ASCMAIN1.sql = "Select APTINVH1.*" & vbCrLf _
                & " from APTINVH1" & vbCrLf _
                & " where APTINVH1.LC_CTL_NO = :PARM1"
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0, False, "V")

        End With

        grdPOTLTRCP.DataSource = dst.Tables("POTLTRCP")
        grdAPTINVH1.DataSource = dst.Tables("APTINVH1")
        grdPOTLTRCX.DataSource = dst.Tables("POTLTRCX")

        Create_Summary(grdPOTLTRCX, "LC_CTL_NO", "Count")
        Create_Summary(grdPOTLTRCX, New String() {"LC_AMT", "LC_PMTS", "LC_FEES", "LC_OPEN"})

        Create_Summary(grdAPTINVH1, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVH1, New String() {"INV_AMT", "LC_FEE"})

        Create_Summary(grdPOTLTRCP, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTLTRCP, New String() {"SEL", "PO_AMT_ORD", "PO_AMT_OPN", "PO_AMT_SHP", "PO_AMT_REC"})
        grdPOTLTRCP.DisplayLayout.Bands(0).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False

        With grdPOTLTRCP.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("PO_ORDER_NO").Header.Fixed = True
        End With

        With grdPOTLTRCX.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"LC_CTL_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"VEND_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("LC_CTL_NO").Header.Fixed = True
        End With


        With grdPOTLTRCP.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                GCOL.CellAppearance.BackColor = Drawing.Color.WhiteSmoke

                If New String() {"SEL"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                    GCOL.CellAppearance.BackColor = Drawing.Color.Empty

                ElseIf New String() {"VEND_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                End If
            Next
            .Columns("SEL").Header.Fixed = True
        End With



        '  ASCMAIN1.Add_Value_List(grdPOTLTRCX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'POTLTRC1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        'ASCMAIN1.Add_Value_List(grdPOTLTRCX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'POTLTRC1' and COLUMN_NAME = 'STATUS_CODE'")

        grpHeader.Visible = False

        '  Absx1.txtFor("CURR_CODE").ReadOnly = True

        Show_Filter(grdPOTLTRCX, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("VEND_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Supplier Code"
                Else
                    Dim row As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Supplier Code Entered Is Not Valid"
                    Else
                        If row.Item("VEND_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Supplier Status Is Not Active"
                        End If
                    End If
                End If

                Dim DT As Date = Absx1.dteFor("LC_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "LC Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                End If

                If Absx1.txtFor("BANK_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Bank Code"
                Else
                    Dim row As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Bank Code Entered Is Not Valid"
                    Else
                        If row.Item("BANK_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Bank Code Entered Is Not Active"
                        End If
                    End If
                End If

                'If Absx1.txtFor("LC_REF_NO").Text.Length = 0 Then
                '    EMsg &= vbCr & "You must enter an LC Reference Number"
                'End If
                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & Absx1.txtFor("VEND_CODE").Text) Then Exit Sub
                End If

            Case "View", "Edit"
                LC_CTL_NO = Absx1.txtFor("LC_CTL_NO").Text
                If LC_CTL_NO = "" Then
                    EMsg &= vbCr & "You must specify an LC No to View"
                Else
                    Dim row As DataRow = LookUp("POTLTRC1", LC_CTL_NO)
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & LC_CTL_NO & " on File"
                    Else
                        If eItemKey = "Edit" Then
                            If EMsg = "" Then
                                If Not ASCMAIN1.Logical_Lock("POTLTRC1", LC_CTL_NO) Then Exit Sub
                                If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & row.Item("VEND_CODE")) Then Exit Sub

                            End If
                        End If
                    End If
                End If

            Case "Update"

                If Absx1.txtFor("BANK_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Bank Code"
                Else
                    Dim row As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Bank Code Entered Is Not Valid"
                    Else
                        If row.Item("BANK_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Bank Status Is Not Active"
                        End If
                    End If
                End If

                Dim DT As Date = Absx1.dteFor("LC_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "LC Date is Mandatory"
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

                If Absx1.txtFor("LC_REF_NO").Text.Length = 0 Then
                    EMsg &= vbCr & "You must enter an LC Reference for this Entry"
                End If
                If Absx1.txtFor("LC_NO").Text.Length = 0 Then
                    EMsg &= vbCr & "You must enter an LC No for this Entry"
                End If

                'If grdPOTLTRCP.Rows.Count = 0 Then
                '    EMsg &= vbCr & "No Items or Collection Details Entered"
                'Else

                'End If

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
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                        .Items("Delete").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode = "E" Then
                        .Items("Delete").Visible = True
                        .Items("Delete").Visible = False ' NOT UNTIL WE FIGURE OUT PROTECTIONS
                    Else
                        .Items("Delete").Visible = False
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If
                End With

                .Groups("Totals").Visible = ScreenMode
                .Groups("Show").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode


        grdPOTLTRCX.Visible = Not ScreenMode

        If ScreenMode Then

            grdPOTLTRCP.DisplayLayout.Bands(0).Columns("SEL").Hidden = (EntryMode = "V")

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            If EntryMode = "E" Or EntryMode = "N" Then
                Set_Read_Only_for_ctl(Absx1.txtFor("LC_REF_NO"), False)
                Set_Read_Only_for_ctl(Absx1.dteFor("LC_DATE"), False)
                '   Set_Read_Only_for_ctl(Absx1.txtFor("CURR_CODE"), True)
            End If

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTLTRCP, grdAPTINVH1}
                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name = "grdPOTLTRCP" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.True
                        Else
                            '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            '.AllowDelete = DefaultableBoolean.True
                            '.AllowUpdate = DefaultableBoolean.True

                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.False
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
            grdAPTINVH1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTLTRC1", "POTLTRCP", "APTINVH1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowPOTLTRC1 = dst.Tables("POTLTRC1").NewRow
            LC_CTL_NO = ASCMAIN1.Next_Control_No("POTLTRC1.LC_CTL_NO")
            rowPOTLTRC1.Item("LC_CTL_NO") = LC_CTL_NO
            rowPOTLTRC1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowPOTLTRC1.Item("BANK_CODE") = HFs("BANK_CODE")
            rowPOTLTRC1.Item("LC_NO") = HFs("LC_NO")
            rowPOTLTRC1.Item("LC_DATE") = HFs("LC_DATE")
            rowPOTLTRC1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            rowPOTLTRC1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowPOTLTRC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTLTRC1.Item("INIT_DATE") = DATETIME_STAMP
            rowPOTLTRC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTLTRC1.Item("LAST_DATE") = DATETIME_STAMP
            rowPOTLTRC1.Item("STATUS_CODE") = "O"

            dst.Tables("POTLTRC1").Rows.Add(rowPOTLTRC1)

        Else
            rowPOTLTRC1 = Fill_Record("POTLTRC1", Absx1.txtFor("LC_CTL_NO").Text)
            dst.AcceptChanges()
        End If

        STATUS_CODE = rowPOTLTRC1.Item("STATUS_CODE")

        EnforceConstraints(False)

        Fill_Records("APTINVH1", LC_CTL_NO)

        If EntryMode = "V" Then
            Fill_Records("POTLTRCP", New String() {"", LC_CTL_NO})
        Else
            Fill_Records("POTLTRCP", New String() {Absx1.txtFor("VEND_CODE").Text, LC_CTL_NO})
        End If
        Sort_grdColumns(grdPOTLTRCP, "PO_ORDER_NO")

        For Each row As DataRow In dst.Tables("POTLTRCP").Select("LC_CTL_NO = '" & LC_CTL_NO & "'")
            row.Item("SEL") = "1"
        Next

        rowPOTLTRC1.Item("LC_PMTS") = Val(dst.Tables("APTINVH1").Compute("SUM(INV_AMT)", "") & "")
        rowPOTLTRC1.Item("PYMTS") = Val(dst.Tables("APTINVH1").Compute("COUNT(VOUCHER_NO)", "") & "")
        Synch_TABLE_NAME("POTLTRC1")
        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        'Dim OPS_YYYYPP As String = Format(rowPOTLTRC1.Item("LC_DATE"), "yyyyMM")
        'rowPOTLTRC1.Item("OPS_YYYYPP") = OPS_YYYYPP
        '  rowPOTLTRC1.Item("OPS_YYYYWW") = Set_OPS_YYYYWW()
        'rowPOTLTRC1.Item("DATE_START") = rowPOTLTRC1.Item("LC_DATE")
        'rowPOTLTRC1.Item("DATE_END") = rowPOTLTRC1.Item("LC_DATE")

        If EntryMode = "E" Then
            If STATUS_CODE <> Absx1.optFor("STATUS_CODE").Value Then
                If STATUS_CODE = "O" Then
                    ASCMAIN1.Record_Event("POTLTRC1", LC_CTL_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CLSD", "LC Closed", "")
                Else
                    ASCMAIN1.Record_Event("POTLTRC1", LC_CTL_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "REOP", "LC Re-Opened", "")
                End If
            End If
        End If

        If EntryMode = "N" Then
            ASCMAIN1.Record_Event("POTLTRC1", LC_CTL_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "ADD", "LC Created", "")
        Else
            ASCMAIN1.Record_Event("POTLTRC1", LC_CTL_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CHG", "LC Updated", "")
        End If

        If EntryMode <> "N" Then Delete_Records()

        Dim SQLD As String = "LC_CTL_NO = '" & LC_CTL_NO & "'"
        INIT_LAST("POTLTRC1", False, , True)

        Update_Record_TDA("POTLTRC1", SQLD)

        ASCMAIN1.sql = "Update POTORDR1 Set LC_CTL_NO = null where LC_CTL_NO = '" & LC_CTL_NO & "'"
        ASCDATA1.ExecuteSQL()

        For Each row As DataRow In dst.Tables("POTLTRCP").Select("SEL='1'")
            Dim PO_ORDER_NO = row.Item("PO_ORDER_NO")
            ASCMAIN1.sql = "Update POTORDR1 Set LC_CTL_NO = :PARM1 where PO_ORDER_NO = :PARM2"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {LC_CTL_NO, PO_ORDER_NO})
        Next


        Update_Record_TDA("ASTAUDT1")

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
            {"POTLTRC1"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where LC_CTL_NO = '" & LC_CTL_NO & "'"
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
                Absx1.txtFor("LC_CTL_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTLTRC1"
            E.COLUMN_NAME = "LC_CTL_NO"
            E.CODE_VALUE = Absx1.txtFor("LC_CTL_NO").Text
            E.DESC_VALUE = Absx1.txtFor("VEND_CODE").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "POTLTRC1"
        E.TABLE_KEY_CAPTION = "LC Events"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("LC_CTL_NO").Text '  HFs("CUST_CODE")
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
        Load_Popup_Menu(grdPOTLTRCX, "SS", "Show Filter", "Show GroupBox") ', "Move to Pending", "Approve")
        ' Load_Popup_Menu(grdPOTLTRC3, "SSBBBB", "Show Filter", "Show GroupBox", "Load Stores w/Attribute", "Load All Stores", "Select All", "De-Select All")
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


            Case "grdSPTSFOCX"

                'If Not InquiryMode And (optShow.Value = "P" Or optShow.Value = "G") Then
                '    tlb_pop.Tools("Move to Pending").SharedProps.Visible = (optShow.Value = "P")
                '    tlb_pop.Tools("Approve").SharedProps.Visible = True
                'Else
                '    tlb_pop.Tools("Move to Pending").SharedProps.Visible = False
                '    tlb_pop.Tools("Approve").SharedProps.Visible = False
                'End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
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
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "LC_CTL_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "LC_CTL_NO"
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

    Private Sub grdPOTLTRCP_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTLTRCP.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "ITEM_CODE"
            '    '   If cdr IsNot Nothing Then
            '    Dim ITEM_CODE As String = CStr(e.Cell.Value & "").ToUpper
            '    ' Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)

            '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            '    If rowICTITEM1 IsNot Nothing Then
            '        If e.Cell.Value & "" <> ITEM_CODE Then
            '            e.Cell.Row.Cells("ITEM_CODE").Value = ITEM_CODE
            '        End If
            '        e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
            '        e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
            '        e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
            '    End If
        End Select
    End Sub

    Private Sub grdPOTLTRCP_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTLTRCP.AfterRowActivate
        'With grdPOTLTRCP.DisplayLayout.Bands(0)
        '    If grdPOTLTRCP.ActiveRow.IsAddRow Then
        '        .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        '        grdPOTLTRCP.ActiveCell = grdPOTLTRCP.ActiveRow.Cells("COLLECTION_CODE")
        '        grdPOTLTRCP.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
        '    Else
        '        .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdPOTLTRCP_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTLTRCP.BeforeExitEditMode
        If grdPOTLTRCP.ActiveCell Is Nothing Then Exit Sub
        With grdPOTLTRCP.ActiveCell
            Select Case .Column.Key
                'Case "COLLECTION_CODE"
                '    If .Text <> "" Then
                '        If .Value IsNot Nothing Then
                '            .Value = .Text.ToUpper
                '        End If

                '    End If
                '    If .Text <> "" Then
                '        cdr = LookUp("ICTCOLL1", .Text)
                '        If cdr Is Nothing Then
                '            ASCMAIN1.Progress("Invalid Collection Code (" & .Text & ")")
                '            If .Value IsNot Nothing Then
                '                .Value = ""
                '            End If
                '            e.Cancel = True
                '        End If
                '    End If
            End Select
        End With
    End Sub

    Private Sub grdPOTLTRCP_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTLTRCP.BeforeRowUpdate
        'With grdPOTLTRCP
        '    If e.Row.Cells("COLLECTION_CODE").Text = "" Then
        '        '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
        '        ' e.Cancel = True MAGIC UNICORN IS A COMMENTED LINE
        '    Else
        '        LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Text)
        '        If cdr Is Nothing Then
        '            MsgBox("Invalid Value entered for Collection Code (" & e.Row.Cells("COLLECTION_CODE").Text & ")",
        '                   MsgBoxStyle.OkOnly, "Cannot Update Row")
        '            e.Cancel = True
        '        End If
        '    End If

        '    If e.Cancel Then
        '        e.Row.CancelUpdate()
        '    End If

        '    If Not e.Cancel Then
        '        If e.Row.Cells("LC_CTL_NO").Text = "" Then
        '            .ActiveRow.Cells("LC_CTL_NO").Value = Absx1.CtlFor("LC_CTL_NO").Text
        '            .ActiveRow.Cells("EVENT_GROUP_LNO").Value = Val(dst.Tables("POTLTRC2").Compute("Max(EVENT_GROUP_LNO)", "") & "") + 1
        '        End If
        '    End If
        'End With
    End Sub

    Private Sub grdPOTLTRCP_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTLTRCP.ClickCellButton

        'If grdPOTLTRCP.ActiveRow Is Nothing Then Exit Sub

        'Dim sql_where As String = ""
        'Select Case e.Cell.Column.Key
        '    Case "ITEM_CODE"
        '    Case "COLLECTION_CODE"
        '        sql_where = "COLLECTION_STATUS = 'A'"
        'End Select
        'grdClickCellButton(grdPOTLTRCP, sql_where, False)

    End Sub

    Private Sub grdPOTLTRCP_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdPOTLTRCP.Error
        grdPOTLTRCP.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Private Sub grdSPTSFOCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTLTRCX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("LC_CTL_NO").Text = e.Row.Cells("LC_CTL_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        EnforceConstraints(False)
        If optShow.Value = "O" Then
            ASCMAIN1.sql = sqlPOTLTRCX & " and STATUS_CODE = 'O'"
            Fill_Records("POTLTRCX", "", True, ASCMAIN1.sql)
            grdPOTLTRCX.Text = "Open"
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlPOTLTRCX
            Fill_Records("POTLTRCX", "", True, ASCMAIN1.sql)
            grdPOTLTRCX.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdPOTLTRCX, "LC_CTL_NO".ToLower)
    End Sub

    Private Sub optShow_ValueChanged(sender As Object, e As EventArgs) Handles optShow.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub

    Private Sub optSTATUS_CODE_ValueChanged(sender As Object, e As EventArgs) Handles optSTATUS_CODE.ValueChanged
        If ScreenMode Then
            Synch_TABLE_NAME("POTLTRC1")
            Display_Totals()
        End If
    End Sub

    Sub Display_Totals()
        Dim LC_OPEN_CALC As Decimal = 0
        Dim LC_CANC_CALC As Decimal = 0
        Dim LC_AMT As Decimal = Val(Absx1.numFor("LC_AMT").Value & "")
        Dim LC_PMTS As Decimal = Val(Absx1.numFor("LC_PMTS").Value & "")
        If optSTATUS_CODE.Value = "O" Then
            LC_OPEN_CALC = LC_AMT - LC_PMTS
            LC_CANC_CALC = 0
        Else
            LC_CANC_CALC = LC_AMT - LC_PMTS
            LC_OPEN_CALC = 0
        End If

        rowPOTLTRC1.Item("LC_OPEN_CALC") = LC_OPEN_CALC
        rowPOTLTRC1.Item("LC_CANC_CALC") = LC_CANC_CALC

        Display_Totals_PO()
    End Sub

    Private Sub grdPOTLTRCP_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdPOTLTRCP.AfterRowUpdate
        Display_Totals_PO()
    End Sub

    Sub Display_Totals_PO()

        Dim LC_PO As Decimal =
            Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_OPN)", "SEL='1'") & "") +
            Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_SHP)", "SEL='1'") & "")
        'Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_REC)", "SEL='1'") & "")

        rowPOTLTRC1.Item("LC_PO") = LC_PO

        Dim LC_AMT As Decimal = Val(Absx1.numFor("LC_AMT").Value & "")

        If LC_PO > LC_AMT Then
            Absx1.numFor("LC_PO").Appearance.ForeColor = Drawing.Color.Red
        Else
            Absx1.numFor("LC_PO").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub
End Class