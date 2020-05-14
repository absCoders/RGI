Public Class ICFRSTY1
    ' GRID M DISPLAY IS NOT RESET TO NO COLUMNS AFTER 1ST MULTI ENTRY
    ' NO KEY FOR ICTRSTY2 IS SCREWY 
#Region "Declarations"
    Dim CUST_CODE As String
    Dim RANGE_STYLE_CODE As String

    Dim rowICTRSTY1 As DataRow
    Dim rowICTSTYL1 As DataRow
    Dim rowARTCUST1 As DataRow 

    Dim COLOR_CODEs As New List(Of String)
    Dim COLs_max As Int64 = 60 ' number of range style qty configurations supported - need to change ICTRSTY1 if you are thinking of changing this value

    Dim STYLE_PRICE_last As Decimal

    ' ALTER TABLE ICTRSTY1 DROP COLUMN RANGE_AS_REPLACEMENT;
    ' ALTER TABLE ICTRSTY2 DROP COLUMN RANGE_STYLE_DESC;

    ' KILL ALL OF THE QTY FIELDS FROM ICTRSTY1 
    ' KILL CUST_SKU AND CUST_UPC IN ICTRSTY2 - THESE COLS RESTORED AFTER DEMO TO ALLISON - KMART USES THEM

    Dim ICTRSTY1 As String = ""
    Dim ICTRSTY2 As String = ""

    Dim EDT850TX As String = ""
    Dim sqlICTRSTYX As String = ""

    Dim QTYs() As Int64

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            sqlICTRSTYX = "Select ICTRSTY1.*, ARTCUST1.CUST_NAME" & vbCrLf _
                & " from ICTRSTY1,ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = ICTRSTY1.CUST_CODE" & vbCrLf
            ASCMAIN1.sql = sqlICTRSTYX _
                & "   and ICTRSTY1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTRSTYX", "**", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "ICTRSTY1", "*", 2)

            ASCMAIN1.sql = "Select ICTRSTY2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
             & " from ICTRSTY2, ICTSTYL1, ICTCOLR1" & vbCrLf _
             & " where ICTSTYL1.STYLE_CODE = ICTRSTY2.STYLE_CODE" & vbCrLf _
             & "   and ICTCOLR1.COLOR_CODE = ICTRSTY2.COLOR_CODE" & vbCrLf _
             & "   and ICTRSTY2.CUST_CODE = :PARM1" & vbCrLf _
             & "   and ICTRSTY2.RANGE_STYLE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTRSTY2", "**", 0, True, "VV", 0)

            Dim DT As DataTable = .Tables("ICTRSTY2").Clone
            DT.TableName = "ICTRSTYM"
            .Tables.Add(DT)
            .Tables("ICTRSTYM").Columns.Add("STYLE_AMT", GetType(System.Decimal), "STYLE_QTY * STYLE_PRICE")
            Dim C As String = ""
            For i As Integer = 1 To COLs_max
                Dim COLUMN_NAME As String = "QTY_" & Format(i, "00")
                C &= "+" & "ISNULL(QTY_" & Format(i, "00") & ",0)"
                .Tables("ICTRSTYM").Columns.Add(COLUMN_NAME, GetType(System.Int64))
            Next
            .Tables("ICTRSTYM").Columns.Add("QTY_TOTAL", GetType(System.Int64), Mid(C, 2))

            ASCMAIN1.sql = "select ICTRSTY1.CUST_CODE ,ICTRSTY1.RANGE_STYLE_CODE ,ICTRSTY1.RANGE_STYLE_DESC " & vbCrLf _
                & ", ICTRSTY1.RANGE_UPC_CODE, ICTRSTY1.RANGE_SKU, ICTRSTY1.RNG_AST_FLG, ICTRSTY2.STYLE_CODE ,ICTRSTY2.COLOR_CODE " & vbCrLf _
                & ", ICTRSTY2.SIZE_CODE ,ICTRSTY2.STYLE_PRICE ,ICTRSTY2.STYLE_QTY ,ICTRSTY2.CUST_SKU " & vbCrLf _
                & ", ICTRSTY2.CUST_UPC ,ICTRSTY2.RANGE_QTY_COL, '0' ERR_TYPE " & vbCrLf _
                & " from ICTRSTY1, ICTRSTY2 " & vbCrLf _
                & " where ICTRSTY1.CUST_CODE = ICTRSTY2.CUST_CODE " & vbCrLf _
                & "   and ICTRSTY1.RANGE_STYLE_CODE = ICTRSTY2.RANGE_STYLE_CODE" & vbCrLf _
                & "   and ROWNUM < 1"
            Create_TDA(.Tables.Add, "ICTRSTYR", "**", 0, False)
            .Tables("ICTRSTYR").Columns.Add("STYLE_AMT", GetType(System.Decimal), "STYLE_QTY * STYLE_PRICE")


            'ASCMAIN1.sql = "Select EDI_JRNL_NO, EDI_PO_NO, TO_DATE(EDI_PO_DATE,'DD-MON-YY') EDI_PO_DATE from EDTJRNL1" _
            '    & " where (EDI_PROCESS_IND is Null OR EDI_JRNL_DATE_TIME > SYSDATE -30) and DOCUMENT_ID = '850'" _
            '    & " and CUST_CODE = :PARM1"
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                ASCMAIN1.sql = "Select Distinct EDTJRNL1.EDI_JRNL_NO, EDTJRNL1.EDI_PO_NO" _
                   & ", TO_DATE(EDTJRNL1.EDI_PO_DATE,'DD-MON-YY') EDI_PO_DATE" _
                   & " from EDTJRNL1,EDT850T1,EDT850T2" _
                   & " where (EDTJRNL1.EDI_PROCESS_IND is Null OR EDTJRNL1.EDI_JRNL_DATE_TIME > SYSDATE -30)" _
                   & "   and EDTJRNL1.DOCUMENT_ID = '850' and EDTJRNL1.CUST_CODE = :PARM1" _
                   & "   and EDT850T1.EDI_JRNL_NO = EDTJRNL1.EDI_JRNL_NO" _
                   & "   and EDT850T2.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO" _
                   & "   and EDT850T2.EDI_UPC in (Select UPC_CODE from ICVLUPC1 where STYLE_CODE = :PARM2)"
            Else
                ASCMAIN1.sql = "Select Distinct EDT850T1.EDI_JRNL_NO, EDT850T1.EDI_PO_NO, EDT850T1.EDI_PO_DATE" & vbCrLf _
                    & " from EDT850T1,EDT850T2" & vbCrLf _
                    & " where (EDT850T1.EDI_PROCESS_IND is Null OR EDT850T1.INIT_DATE > SYSDATE -30)" & vbCrLf _
                    & "   and EDT850T1.CUST_CODE = :PARM1" & vbCrLf _
                    & "   and EDT850T2.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO" & vbCrLf _
                    & "   and EDT850T2.EDI_UPC in (Select UPC_CODE from ICVLUPC1 where STYLE_CODE = :PARM2)"
            End If

            Create_TDA(.Tables.Add, "EDT850TX", "**", 0, False, "VV", 0)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            'ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
            '    & " from ICTCOLR1,ICTSTYC1" _
            '    & " where ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE" _
            '    & "   and ICTSTYC1.STYLE_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "ICTCOLRM", "**", 0, False, "V", 1)
            '.Tables("ICTCOLRM").Columns.Add("QTY", GetType(System.Int32))

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)
        End With

        grdEDT850TX.DataSource = dst.Tables("EDT850TX")

        grdICTRSTYX.DataSource = dst.Tables("ICTRSTYX")
        grdICTRSTYM.DataSource = dst.Tables("ICTRSTYM")
        grdICTRSTYR.DataSource = dst.Tables("ICTRSTYR")

        grdICTRSTYX.DisplayLayout.UseFixedHeaders = True
        With grdICTRSTYX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "RANGE_STYLE_CODE", "RANGE_STYLE_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdICTRSTYM.DisplayLayout.UseFixedHeaders = True
        With grdICTRSTYM.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdICTRSTYM.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_QTY", "STYLE_PRICE", "CUST_SKU", "CUST_UPC", "SIZE_CODE"}.Contains(gcol.Key) _
                    Or gcol.Key.StartsWith("QTY_") Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If

                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"RANGE_QTY_COL", "STYLE_QTY", "STYLE_PRICE", "STYLE_AMT"}.Contains(gcol.Key) Then
                        .BackColor2 = Drawing.Color.LightBlue
                    ElseIf gcol.Key.StartsWith("QTY_") Then
                        .BackColor2 = Drawing.Color.LightGreen
                    ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                        .BackColor2 = Drawing.Color.Pink
                    ElseIf New String() {"CUST_SKU", "CUST_UPC", "SIZE_CODE"}.Contains(gcol.Key) Then
                        .BackColor2 = Drawing.Color.Orange
                    Else
                        .BackColor2 = Drawing.Color.LightGray
                    End If
                End With
            Next
        End With

        grdICTRSTYR.DisplayLayout.UseFixedHeaders = True
        With grdICTRSTYR.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "RANGE_STYLE_CODE", "RANGE_STYLE_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdICTRSTYR.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"CUST_CODE", "RANGE_STYLE_CODE", "RANGE_STYLE_DESC"}.Contains(gcol.Key) Then
                        .BackColor2 = Drawing.Color.LightBlue
                    ElseIf New String() {"RANGE_QTY_COL", "STYLE_QTY", "STYLE_PRICE", "STYLE_AMT"}.Contains(gcol.Key) Then
                        .BackColor2 = Drawing.Color.LightBlue
                    ElseIf gcol.Key.StartsWith("QTY_") Then
                        .BackColor2 = Drawing.Color.LightGreen
                    ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                        .BackColor2 = Drawing.Color.Pink
                    ElseIf New String() {"CUST_SKU", "CUST_UPC", "SIZE_CODE"}.Contains(gcol.Key) Then
                        .BackColor2 = Drawing.Color.Orange
                    Else
                        .BackColor2 = Drawing.Color.LightGray
                    End If
                End With
            Next
        End With

        Create_Summary(grdICTRSTYX, "RANGE_STYLE_CODE", "Count")

        Create_Summary(grdICTRSTYM, "STYLE_CODE", "Count")
        Create_Summary(grdICTRSTYM, New String() {"STYLE_QTY", "STYLE_AMT"})

        Create_Summary(grdICTRSTYR, "CUST_CODE", "Count")
        Show_Filter(grdICTRSTYR, True)

        Setup_QTY_COLs(0, True)

        Show_Filter(grdICTRSTYX, True)
        grdICTRSTYX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdICTRSTYX, "RNG_AST_FLG", , New String() {":", "R:Range", "A:Asst", "F:Cases", "M:Multi"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New", "Edit", "Load", "MultiRange Load"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If Not eItemKey = "MultiRange Load" Then
                    If Absx1.txtFor("RANGE_STYLE_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must First Specify a Range Style"
                    Else
                        RANGE_STYLE_CODE = Absx1.txtFor("RANGE_STYLE_CODE").Text
                        rowICTRSTY1 = LookUp("ICTRSTY1", New String() {CUST_CODE, RANGE_STYLE_CODE})
                        If eItemKey = "New" Then
                            If rowICTRSTY1 IsNot Nothing Then
                                EMsg &= vbCr & "Range Style " & RANGE_STYLE_CODE & " already exists for Customer " & CUST_CODE

                            Else
                                If Absx1.txtFor("RANGE_STYLE_DESC").Text = "" Then
                                    EMsg &= vbCr & "You Must First Specify a Range Style Description"
                                End If
                            End If
                        Else
                            If rowICTRSTY1 Is Nothing Then
                                EMsg &= vbCr & "No Record of Range Style " & RANGE_STYLE_CODE & " for Customer " & CUST_CODE
                            End If
                        End If
                    End If
                End If

                If eItemKey = "New" Or eItemKey = "Edit" Then
                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("ICTRSTY1", CUST_CODE & "||" & RANGE_STYLE_CODE) Then Exit Sub
                    End If
                End If

            Case "Update"

                If optType.Value = "M" Then
                    For i As Integer = 1 To QTYs.Length - 1
                        Dim QTY As Int64 = Val(dst.Tables("ICTRSTYM").Compute("SUM(QTY_" & Format(i, "00") & ")", "") & "")
                        If QTYs(i) <> QTY Then
                            EMsg &= vbCr & "Total Distribution (" & CStr(QTY) & ") out of Balance for Qty Column " & CStr(i) & " (Total S/B = " & CStr(QTYs(i)) & ")"
                        End If
                    Next i
                End If
                If optType.Value = "R" Then
                    If ASCDATA1.SelectDistinct(dst.Tables("ICTRSTYM"), New String() {"STYLE_PRICE"}).Rows.Count > 1 Then
                        EMsg &= vbCr & "Cannot have More Than One Price for the Style Components of a Range Style"
                    End If
                End If

                Dim sqlw As String = "ISNULL(STYLE_PRICE,0) = 0 "
                If optType.Value = "M" Then
                    sqlw &= "or ISNULL(QTY_TOTAL,0) = 0"
                Else
                    sqlw &= "or ISNULL(STYLE_QTY,0) = 0"
                End If
                If dst.Tables("ICTRSTYM").Select(sqlw).Length > 0 Then
                    EMsg &= vbCr & "Styles with no price or zero qty not allowed"
                End If

                If Absx1.txtFor("RANGE_STYLE_DESC").Text = "" And EntryMode <> "M" Then
                    EMsg &= vbCr & "Range Style Description is Required"
                End If

            Case "Cancel"
                If cmdUPC.Visible And Not cmdUPC.Enabled Then
                    If MsgBox("If you Cancel your Edits at this point," _
                              & vbCrLf _
                              & " the UPC Code which was Generated will be wasted." _
                              & vbCrLf & vbCrLf _
                              & "OK To Cancel your Edits (and lose the UPC)?", _
                              MsgBoxStyle.YesNo, "UPC was Generated") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

                If MsgBox("OK To Cancel your Edits?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

                If MsgBox("OK To Delete this Range Style Definition?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Copy and Edit"
                If grdICTRSTYX.ActiveRow IsNot Nothing AndAlso grdICTRSTYX.ActiveRow.IsDataRow Then
                    Copy_Range_Style(grdICTRSTYX.ActiveRow.Cells("CUST_CODE").Value, grdICTRSTYX.ActiveRow.Cells("RANGE_STYLE_CODE").Value)
                End If

            Case "Update"
                If EntryMode = "M" Then
                    Update_Range()
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "MultiRange Load"
                EntryMode = "M"
                grdICTRSTYR.Text = "Load Range Styles for " & CUST_CODE
                Mode_Settings(True)
                Load_MultiRange()

            Case "Done"
                Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "L" And ScreenMode) Then
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    .Items("Copy and Edit").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Copy and Edit").Settings.Enabled = not_iScreenMode
                End If
                .Items("Update").Settings.Enabled = iScreenMode

                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("Load").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode
                .Items("MultiRange Load").Settings.Enabled = not_iScreenMode

                .Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                .Items("Print").Visible = False ' ScreenMode
                .Items("Update").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Delete").Visible = (EntryMode = "E")
                .Items("Cancel").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
            End With
          
            '    .Groups("EDI Orders").Visible = ScreenMode
        End With

        'lblStatus.Visible = ScreenMode

        grdICTRSTYX.Visible = Not tf
        SplitContainer1.Panel2Collapsed = True

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        If ScreenMode Then Set_Read_Only_for_ctl(Absx1.txtFor("RANGE_STYLE_DESC"), False)

        If ScreenMode Then
            If EntryMode = "L" Then
                grdICTRSTYM.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdICTRSTYM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdICTRSTYM.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdICTRSTYM.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdICTRSTYM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdICTRSTYM.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            End If

            cmdUPC.Visible = False
            If EntryMode = "L" Then
            Else
                If Absx1.txtFor("RANGE_UPC_CODE").Text & "" = "" Then
                    cmdUPC.Visible = True
                    cmdUPC.Enabled = True
                End If
            End If

            If EntryMode = "M" Then
                SplitContainer1.Panel1Collapsed = True
                SplitContainer1.Panel2Collapsed = False
                grdICTRSTYR.Visible = True
            End If

        Else
            Clear_Record()
        End If

        ' SplitContainer1.Panel2Collapsed = True
    End Sub

    Sub Clear_Record()

        Absx1.txtFor("RANGE_STYLE_CODE").Text = ""
        Absx1.txtFor("RANGE_STYLE_DESC").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""

        CUST_CODE = ""
        RANGE_STYLE_CODE = ""
        STYLE_PRICE_last = 0

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ICTRSTY1", "ICTRSTY2", "EDT850TX", "ICTRSTYM", "ICTRSTYR"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Setup_QTY_COLs(0)

        Load_ICTRSTYX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowICTRSTY1 = dst.Tables("ICTRSTY1").NewRow
            With rowICTRSTY1
                .Item("CUST_CODE") = CUST_CODE
                .Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
                .Item("RANGE_STYLE_DESC") = Absx1.txtFor("RANGE_STYLE_DESC").Text
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("RNG_AST_FLG") = "R"
                .Item("RANGE_PRICE") = 0
            End With
            dst.Tables("ICTRSTY1").Rows.Add(rowICTRSTY1)
        Else
            rowICTRSTY1 = Fill_Record("ICTRSTY1", New String() {CUST_CODE, RANGE_STYLE_CODE})
        End If

        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)

        Fill_Records("ICTRSTY2", New String() {CUST_CODE, RANGE_STYLE_CODE})
        Fill_Records("EDT850TX", New String() {CUST_CODE, RANGE_STYLE_CODE})
        Sort_grdColumns(grdEDT850TX, "EDI_PO_DATE".ToLower)

        dst.Tables("ICTRSTYM").Rows.Clear()

        Dim RANGE_QTY_COL As Int64 = Val(dst.Tables("ICTRSTY2").Compute("MAX(RANGE_QTY_COL)", "") & "")
        Setup_QTY_COLs(RANGE_QTY_COL)
        UltraExplorerBar1.Groups("EDI Orders").Visible = (optType.Value = "M")

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ICTRSTY2"), New String() {"STYLE_CODE", "COLOR_CODE", "SIZE_CODE"}).Rows
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim SIZE_CODE As String = row.Item("SIZE_CODE")
            Dim sqlw As String = "STYLE_CODE = '" & STYLE_CODE & "'" _
                                 & " and COLOR_CODE = '" & COLOR_CODE & "'" _
                                 & " and ISNULL(SIZE_CODE,'') = '" & SIZE_CODE & "'"

            Dim rows() As DataRow = dst.Tables("ICTRSTY2").Select(sqlw)

            Dim rowICTRSTYM As DataRow = dst.Tables("ICTRSTYM").NewRow
            For I As Integer = 0 To dst.Tables("ICTRSTY2").Columns.Count - 1
                rowICTRSTYM.Item(I) = rows(0).Item(I)
            Next

            For J As Integer = 1 To RANGE_QTY_COL
                Dim QTY As Int64 = Val(dst.Tables("ICTRSTY2").Compute("SUM(STYLE_QTY)", sqlw & " and RANGE_QTY_COL = " & CStr(J)) & "")
                rowICTRSTYM.Item("QTY_" & Format(J, "00")) = QTY
            Next
            dst.Tables("ICTRSTYM").Rows.Add(rowICTRSTYM)
        Next

        Sort_grdColumns(grdICTRSTYM, "STYLE_CODE,COLOR_CODE")

        'If EntryMode = "N" Then
        '    Load_QTYs_from_EDI()
        '    If QTYs IsNot Nothing AndAlso QTYs.Length > 0 Then
        '        rowICTRSTY1.Item("RNG_AST_FLG") = "M"
        '    End If
        'End If

        With grdICTRSTYM.DisplayLayout.Bands(0)
            'If (EntryMode = "E" Or EntryMode = "N") Then
            '    .Columns("STYLE_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            '    If EntryMode = "E" Then
            '        .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
            '    Else
            '        .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            '    End If
            'Else
            '    .Columns("RSRV_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
            '    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            'End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdICTRSTYM.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            End With
            grdICTRSTYM.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, False)
        Else
            With grdICTRSTYM.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            grdICTRSTYM.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, True)
        End If

        If EntryMode = "N" Then
            lblINIT_DATE.Visible = False
        Else
            lblINIT_DATE.Visible = True
            lblINIT_DATE.Text = "Entered " & Format(rowICTRSTY1.Item("INIT_DATE"), "MM/dd/yyyy") & " by " & rowICTRSTY1.Item("INIT_OPER")
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
        ' Dependent_Updates(-1, RSRV_NO)
        For Each TABLE_NAME As String In New String() _
            {"ICTRSTY1", "ICTRSTY2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME _
            & " where CUST_CODE = '" & CUST_CODE & "'" _
            & "   and RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        If optType.Value = "M" Then
            rowICTRSTY1.Item("RANGE_PRICE") = Val(dst.Tables("ICTRSTYM").Compute("SUM(STYLE_AMT)", "RANGE_QTY_COL = 0") & "")
            'Absx1.numFor("STYLE_PRICE").Value = Val(dst.Tables("ICTRSTY2").Compute("SUM(STYLE_AMT)", "RANGE_QTY_COL = 0") & "")
        ElseIf optType.Value = "F" Or optType.Value = "R" Then
            rowICTRSTY1.Item("RANGE_PRICE") = Val(dst.Tables("ICTRSTYM").Compute("SUM(STYLE_AMT)", "") & "")

        End If

        Dim RANGE_STYLE_QTY_PER_PP As Int64 = 1
        If Not (optType.Value = "R" Or optType.Value = "M") Then
            RANGE_STYLE_QTY_PER_PP = Val(dst.Tables("ICTRSTYM").Compute("SUM(STYLE_QTY)", "") & "")
        End If
        rowICTRSTY1.Item("RANGE_STYLE_QTY_PER_PP") = RANGE_STYLE_QTY_PER_PP
        If optType.Value = "M" Then
            For i As Integer = 1 To COLs_max
                rowICTRSTY1.Item("RANGE_QTY" & IIf(i = 1, "", CStr(i - 1))) = Val(dst.Tables("ICTRSTYM").Compute("SUM(QTY_" & Format(i, "00") & ")", "") & "")
            Next
        Else
            rowICTRSTY1.Item("RANGE_QTY") = Val(dst.Tables("ICTRSTYM").Compute("SUM(STYLE_QTY)", "") & "")
        End If

        dst.Tables("ICTRSTY2").Rows.Clear()
        Dim COLs As Integer = 0
        If optType.Value = "M" Then COLs = QTYs.Length - 1
        For Each rowICTRSTYM As DataRow In dst.Tables("ICTRSTYM").Select("")
            If optType.Value = "M" Then
                For i As Integer = 1 To COLs
                    If Val(rowICTRSTYM.Item("QTY_" & Format(i, "00")) & "") <> 0 Then
                        Dim rowICTRSTY2 As DataRow = dst.Tables("ICTRSTY2").NewRow
                        'rowICTRSTY2.Item("CUST_CODE") = CUST_CODE
                        'rowICTRSTY2.Item("STYLE_CODE") = rowICTRSTYM.Item("STYLE_CODE")
                        'rowICTRSTY2.Item("COLOR_CODE") = rowICTRSTYM.Item("COLOR_CODE")
                        For j As Integer = 0 To rowICTRSTY2.Table.Columns.Count - 1
                            rowICTRSTY2.Item(j) = rowICTRSTYM.Item(j)
                        Next
                        If COLs <> 0 Then
                            rowICTRSTY2.Item("STYLE_QTY") = Val(rowICTRSTYM.Item("QTY_" & Format(i, "00")) & "")
                            rowICTRSTY2.Item("RANGE_QTY_COL") = i - 1
                        End If
                        dst.Tables("ICTRSTY2").Rows.Add(rowICTRSTY2)
                    End If
                Next
            Else
                Dim rowICTRSTY2 As DataRow = dst.Tables("ICTRSTY2").NewRow
                For j As Integer = 0 To rowICTRSTY2.Table.Columns.Count - 1
                    rowICTRSTY2.Item(j) = rowICTRSTYM.Item(j)
                Next
                dst.Tables("ICTRSTY2").Rows.Add(rowICTRSTY2)
            End If
        Next

        BeginTrans()
        INIT_LAST("ICTRSTY1", False, , True)
        Dim sqldelete As String = "CUST_CODE = '" & CUST_CODE & "' and RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'"
        Update_Record_TDA("ICTRSTY1", sqldelete)
        Update_Record_TDA("ICTRSTY2", sqldelete)

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

            Case "RANGE_STYLE_CODE"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    MsgBox("You must enter a Customer Code", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    'sql_where &= " and ICTRSTY1.RANGE_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and ICTRSTY1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
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

            Case "Load"
                If ScreenMode Then
                    Click_Command("Done")
                End If

                Absx1.txtFor("CUST_COODE").Text = Split(key, ":")(0)
                Absx1.txtFor("RANGE_STYLE_COODE").Text = Split(key, ":")(1)
                Click_Command("Load")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "ICTRSTY1"
            E.COLUMN_NAME = "CUST_CODE:RANGE_STYLE_CODE"
            E.CODE_VALUE = Absx1.txtFor("CUST_CODE").Text & ":" & Absx1.txtFor("RANGE_STYLE_CODE").Text
            E.DESC_VALUE = "Customer Range Style"
            E.ATTACHMENT_NOTES = ""
            'If rowICTRSTY1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTRSTYX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdICTRSTYM, "BSS", "Style Status Inquiry", "Style Multi-Color", "Enter Style/Colors w/out Qty", "Load from Spreadsheet")
        Load_Popup_Menu(grdICTRSTYR, "S", "Load Spreadsheet")
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
            Case "grdICTRSTYM"
                tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                'tlb_btn = DirectCast(tlb_pop.Tools("Style Status Inquiry"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                tlb_btn = DirectCast(tlb_pop.Tools("Load from Spreadsheet"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

            Case "grdICTRSTYR"
                tlb_btn = DirectCast(tlb_pop.Tools("Load Spreadsheet"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "M")

                Exit Sub
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
                    F.select_only = True
                    F.ShowDialog()
                    If F.STYLE_CODE <> "" Then
                        Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"))
                    End If
                End Using

            Case "Load Spreadsheet"
                Load_MultiRange()

            Case "Load from Spreadsheet"
                'Excel_Import(grdICTRSTYM)

                'Exit Sub
                Try
                    Dim FILENAME As String = ""
                    Using openFileDialog1 As New OpenFileDialog
                        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                        openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                        openFileDialog1.RestoreDirectory = True

                        '  Excel_Import = -1

                        If openFileDialog1.ShowDialog() = DialogResult.OK Then
                            FILENAME = openFileDialog1.FileName
                        End If
                    End Using

                    If FILENAME <> "" Then

                        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                        Dim range As SpreadsheetGear.IRange = Nothing
                        range = oSheet.UsedRange
                        Dim rmax As Integer = range.RowCount

                        Dim TYPE As String = ""
                        Dim started As Boolean = False

                        Dim INVALID_STYLES As String = ""

                        'grdICTRSTYM.SuspendLayout()

                        Dim r As Integer = 0
                        Do While r < rmax ' oSheet.Cells(r, 0).Value & "" <> ""

                            If TYPE = "" And Not started Then
                                If Trim(oSheet.Cells(r, 0).Value & "").ToUpper = "MEIJER" Then
                                    TYPE = "MEIJER"
                                End If
                                If Trim(oSheet.Cells(r, 0).Value & "").ToUpper = "SHOPKO" Then
                                    TYPE = "SHOPKO"
                                End If
                            End If


                            Dim STYLE_CODE As String = ""
                            Dim COLOR_CODE As String = ""
                            Dim SIZE_CODE As String = ""
                            Dim CUST_SKU As String = ""
                            Dim CUST_UPC As String = ""
                            Dim STYLE_PRICE As Decimal = 0
                            Dim STYLE_QTY As Int64 = 0

                            If TYPE = "" Then
                                STYLE_CODE = oSheet.Cells(r, 0).Value & ""
                                COLOR_CODE = oSheet.Cells(r, 2).Value & ""
                                SIZE_CODE = oSheet.Cells(r, 4).Value & ""
                                CUST_SKU = oSheet.Cells(r, 5).Value & ""
                                CUST_UPC = oSheet.Cells(r, 6).Value & ""
                                STYLE_PRICE = Val(oSheet.Cells(r, 7).Value & "")
                                STYLE_QTY = Val(oSheet.Cells(r, 8).Value & "")
                            ElseIf TYPE = "MEIJER" Or TYPE = "SHOPKO" Then
                                STYLE_CODE = oSheet.Cells(r, 0).Value & ""
                                COLOR_CODE = oSheet.Cells(r, 4).Value & ""
                                SIZE_CODE = oSheet.Cells(r, 3).Value & ""
                                ' CUST_SKU = oSheet.Cells(r, 5).Value & ""
                                CUST_UPC = oSheet.Cells(r, 2).Value & ""
                                STYLE_PRICE = Val(oSheet.Cells(r, 7).Value & "")
                                STYLE_QTY = Val(oSheet.Cells(r, 5).Value & "")
                            End If

                            If STYLE_QTY = 0 Then STYLE_CODE = ""

                            If LookUp("ICTSTYL1", STYLE_CODE) IsNot Nothing AndAlso LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}) IsNot Nothing Then
                                If grdICTRSTYM.ActiveRow IsNot Nothing AndAlso grdICTRSTYM.ActiveRow.IsAddRow Then
                                    grdICTRSTYM.ActiveRow.CancelUpdate()
                                End If

                                grdICTRSTYM.DisplayLayout.Bands(0).AddNew.Activate()
                                With grdICTRSTYM.ActiveRow
                                    .Cells("STYLE_CODE").Value = STYLE_CODE
                                    .Cells("COLOR_CODE").Value = COLOR_CODE
                                    .Cells("SIZE_CODE").Value = SIZE_CODE
                                    .Cells("CUST_SKU").Value = CUST_SKU
                                    .Cells("CUST_UPC").Value = CUST_UPC
                                    .Cells("STYLE_PRICE").Value = STYLE_PRICE
                                    .Cells("STYLE_QTY").Value = STYLE_QTY
                                    .Update()
                                    started = True
                                    'If grdICTRSTYM.ActiveCell.IsInEditMode Then
                                    grdICTRSTYM.ActiveRow.CancelUpdate()
                                    'End If
                                End With
                            Else
                                If STYLE_CODE <> "" And COLOR_CODE <> "" Then INVALID_STYLES &= ", " & STYLE_CODE & "-" & COLOR_CODE
                            End If


                            r += 1
                        Loop

                        ' grdICTRSTYM.ResumeLayout()

                        Sort_grdColumns(grdICTRSTYM, "STYLE_CODE,COLOR_CODE")

                        If INVALID_STYLES <> "" Then
                            MsgBox(Mid(INVALID_STYLES, 3), MsgBoxStyle.OkOnly, "The following Style-Color codes were invalid")
                        End If

                        MsgBox("Spreadsheet has been Loaded", MsgBoxStyle.OkOnly, "Success")
                    End If
                Catch ex As Exception
                    MsgBox("Error " & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load this Spreadsheet")
                End Try

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
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not ScreenMode Then
                        Load_ICTRSTYX()
                    End If
                End If

            Case "RANGE_STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not ScreenMode Then
                        Click_Command("Load")
                    End If
                End If

            Case "RANGE_STYLE_DESC"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not ScreenMode Then
                        Click_Command("New")
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_ICTRSTYX()

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
                Load_ICTRSTYX()
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_ICTRSTYX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE = "" Then
            ASCMAIN1.sql = sqlICTRSTYX
            Fill_Records("ICTRSTYX", "", , ASCMAIN1.sql)
            grdICTRSTYX.Text = "All Range Styles"
            Sort_grdColumns(grdICTRSTYX, "CUST_CODE,RANGE_STYLE_CODE")
        Else
            Fill_Records("ICTRSTYX", CUST_CODE)
            grdICTRSTYX.Text = "All Range Styles associated with " & CUST_CODE
            Sort_grdColumns(grdICTRSTYX, "RANGE_STYLE_CODE")
        End If
        grdICTRSTYX.Visible = True
    End Sub

    Sub Print_Record()
        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = ""
        Generate_Report(RPT, "Customer Range Style", , , , , False)
        Print_Report_End()
    End Sub

    Private Sub grdICTRSTYX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTRSTYX.DoubleClickRow
        Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
        Absx1.txtFor("RANGE_STYLE_CODE").Text = e.Row.Cells("RANGE_STYLE_CODE").Value
        Click_Command("Load")
    End Sub

    Sub Display_Totals()
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdICTRSTYM.ActiveRow
            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    If .Cells("STYLE_CODE").Text <> "" Then
                        Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value & "")
                        Cancel = (STYLE_CODE = "")
                    End If
                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            MsgBox("Valid Colors are: " & Join(COLOR_CODEs.ToArray, ","), MsgBoxStyle.OkOnly, "Invalid Color Code")
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Case "RSRV_QTY"
                    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Val(.Cells("RSRV_QTY").Value & "") < 0 Then
                        MsgBox("Qty May Not be Negative", MsgBoxStyle.OkOnly, "Invalid Order Quantity")
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim E As String = ""
        If STYLE_CODE_z = "" Then Return ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            E = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                E = "Item Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                E = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                E = "Item does not have a valid Division Code" & vbCrLf
            End If
        End If

        If E = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If E <> "" And grdICTRSTYM.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If E = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

#Region "grdICTRSTYM"

    Private Sub grdICTRSTYM_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTRSTYM.AfterCellUpdate
        With grdICTRSTYM.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    If STYLE_CODE <> "" Then
                        .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        If COLOR_CODEs.Count <= 1 Then
                            .Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                        End If
                    End If
                    .Cells("STYLE_PRICE").Value = STYLE_PRICE_last
                    If Absx1.optFor("RNG_AST_FLG").Value <> "M" Then .Cells("STYLE_QTY").Value = 1

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & ""
                    If COLOR_CODE <> "" Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        If rowICTCOLR1 IsNot Nothing Then
                            .Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

                Case "STYLE_PRICE"
                    STYLE_PRICE_last = Val(.Cells("STYLE_PRICE").Value & "")

            End Select
        End With
    End Sub

    Private Sub grdICTRSTY2_AfterExitEditMode(sender As Object, e As System.EventArgs) Handles grdICTRSTYM.AfterExitEditMode
        'With grdICTRSTY2
        '    Select Case .ActiveCell.Column.Key
        '        Case "STYLE_CODE"
        '            Dim STYLE_CODE As String = .ActiveCell.Text
        '            If STYLE_CODE <> "" Then
        '                .ActiveCell.Value = ASCMAIN1.Format_Field(STYLE_CODE, .ActiveCell.Column.Key)
        '            End If
        '    End Select
        'End With
    End Sub

    Private Sub grdICTRSTYM_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTRSTYM.AfterRowActivate

        If grdICTRSTYM.ActiveRow.IsAddRow Then
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Enter Style/Colors w/out Qty"), UltraWinToolbars.StateButtonTool)

            If grdICTRSTYM.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                grdICTRSTYM.ActiveCell = grdICTRSTYM.ActiveRow.Cells("STYLE_CODE")
            End If
            For I As Integer = 1 To COLs_max
                If tlb_sbt.Checked Then
                    grdICTRSTYM.DisplayLayout.Bands(0).Columns("QTY_" & Format(I, "00")).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    grdICTRSTYM.DisplayLayout.Bands(0).Columns("QTY_" & Format(I, "00")).CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        Else
            For I As Integer = 1 To COLs_max
                grdICTRSTYM.DisplayLayout.Bands(0).Columns("QTY_" & Format(I, "00")).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
            With grdICTRSTYM.DisplayLayout.Bands(0)
                Validate_Style(grdICTRSTYM.ActiveRow.Cells("STYLE_CODE").Value & "")
            End With
        End If

        If Trim(grdICTRSTYM.ActiveRow.Cells("STYLE_CODE").Value & "") = "" And _
            (grdICTRSTYM.ActiveCell Is Nothing OrElse _
             (grdICTRSTYM.ActiveCell.Column.Key <> "STYLE_CODE")) _
        Then
            grdICTRSTYM.ActiveCell = grdICTRSTYM.ActiveRow.Cells("STYLE_CODE")
            Exit Sub
        End If
    End Sub

    Private Sub grdICTRSTYM_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdICTRSTYM.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdICTRSTYM_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTRSTYM.AfterRowUpdate
        Display_Totals()
        '  Absx1.numFor("RANGE_PRICE").Value = Val(dst.Tables("ICTRSTYM").Compute("SUM(STYLE_AMT)", "") & "")
    End Sub

    Private Sub grdICTRSTYM_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdICTRSTYM.BeforeCellUpdate
        'If e.Cell.Column.Key = "STYLE_CODE" Or e.Cell.Column.Key = "COLOR_CODE" Then
        '    e.Cell.Value = e.Cell.Value.ToString.ToUpper
        'End If
    End Sub

    Private Sub grdICTRSTYM_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTRSTYM.BeforeExitEditMode
        With grdICTRSTYM
            Select Case .ActiveCell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = .ActiveCell.Text
                    If STYLE_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(STYLE_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdICTRSTYM_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTRSTYM.BeforeRowsDeleted

    End Sub

    Private Sub grdICTRSTYM_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTRSTYM.BeforeRowUpdate

        Validate_Columns("STYLE_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("COLOR_CODE", e.Cancel)
        End If

        If e.Row.Cells("CUST_UPC").Value & "" <> "" Then
            'If Validate_UPC(e.Row.Cells("CUST_UPC").Text & "") <> "" Then
            '    MsgBox(iResult, MsgBoxStyle.OkOnly, "UPC Error")
            '    e.Cancel = True
            'End If
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
            e.Row.Cells("RANGE_STYLE_CODE").Value = Absx1.txtFor("RANGE_STYLE_CODE").Text

            If e.Row.Cells("SIZE_CODE").Value & "" = "" Then
                e.Row.Cells("SIZE_CODE").Value = "AST"
            End If
        End If
    End Sub

    Private Sub grdICTRSTYM_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTRSTYM.ClickCellButton
        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdICTRSTYM, sql_where)
                Case "COLOR_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdICTRSTYM, sql_where)
            End Select
        End With
    End Sub
#End Region

    Private Sub optType_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optType.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Qty()
        UltraExplorerBar1.Groups("EDI Orders").Visible = (optType.Value = "M")
    End Sub

    Sub Set_Qty()

        For i As Integer = 1 To COLs_max
            grdICTRSTYM.DisplayLayout.Bands(0).Columns("QTY_" & Format(i, "00")).Hidden = Not (optType.Value = "M") Or (QTYs Is Nothing OrElse i > QTYs.Length - 1)
        Next i

        With grdICTRSTYM.DisplayLayout.Bands(0)
            .Columns("STYLE_QTY").Hidden = (optType.Value = "M")
            .Columns("STYLE_AMT").Hidden = (optType.Value = "M")
        End With


        'If (optType.Value = "M") And EMode <> "" Then
        '    ReDim aMsgs(10)
        '    aMsgs(0) = "1"
        '    aMessage = Chr$(0) & txtPO.Text
        '    Sql = "Select EDI_JRNL_NO,EDI_PO_NO, CUST_CODE from EDTJRNL1" _
        '        & " where EDI_PROCESS_IND is Null and DOCUMENT_ID = '850'"
        '    Sql = Sql & " and CUST_CODE = '" & CUST_CODE & "'"
        '    ASFCODE1.Show(1)
        '    If aMessage <> "" Then
        '        txtPO.Text = aCodes(1, 0)
        '        '            txtCode_LostFocus (Index)
        '        '            txtCode(Index).SetFocus
        '        jrnlno = aCodes(0, 0)
        '        Sql = "select EDI_STYLE, EDI_COLOR_CODE, EDI_UPC, EDI_SKU from edt850t1 t1, edt850t2 t2 "
        '        Sql = Sql & "where edi_jrnl_no = '" & jrnlno & "' and t1.EDI_DOC_SEQ_NO = t2.EDI_DOC_SEQ_NO "
        '        Sql = Sql & "group by EDI_STYLE, EDI_COLOR_CODE, EDI_UPC, EDI_SKU"
        '        ReDim aMsgs(10)
        '        aMsgs(0) = "1"
        '        aMessage = Chr$(0)
        '        ASFCODE1.Show(1)
        '        If aMessage <> "" Then
        '            If CUST_CODE = "JCPL" Then
        '                Sql = "Select lpad(EDI_TOTAL_QTY,5,'0') EDI_TOTAL_QTY from edt850t1 t1, edt850t2 t2 , edtjrnl1 j1"
        '                Sql = Sql & " where j1.EDI_PROCESS_IND is Null and DOCUMENT_ID = '850'"
        '                Sql = Sql & " and CUST_CODE = '" & CUST_CODE & "'"
        '                If aCodes(0, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_STYLE = '" & aCodes(0, 0) & "' "
        '                End If
        '                If aCodes(1, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_COLOR_CODE = '" & aCodes(1, 0) & "' "
        '                End If
        '                If aCodes(2, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_UPC = '" & aCodes(2, 0) & "' "
        '                End If
        '                If aCodes(3, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_SKU = '" & aCodes(3, 0) & "' "
        '                End If
        '                Sql = Sql & " and t1.EDI_JRNL_NO = j1.EDI_JRNL_NO and t1.EDI_DOC_SEQ_NO = t2.EDI_DOC_SEQ_NO"
        '                Sql = Sql & "  order by 1"
        '                dyn = OraD.CreateDynaset(Sql, 8&)
        '                SQty = ""
        '                Do While Not dyn.EOF
        '                    SQty = SQty & Format$(dyn.Fields("EDI_TOTAL_QTY").Value & "", "00000")
        '                    dyn.MoveNext()
        '                Loop
        '                dyn = Nothing
        '                '                    SQty = Sort_List(SQty, 5)
        '                c = 0
        '                T = Val(Mid(SQty, 1, 5))
        '                S = 0
        '                For i = 1 To Len(SQty) / 5
        '                    If Val(Mid(SQty, (i - 1) * 5 + 1, 5)) <> T Then
        '                        lblStrQty(c) = Str(S) & "@" & Str(T)
        '                        lblStrQty(c).tag = Str(T)
        '                        c = c + 1
        '                        S = 0
        '                        T = Val(Mid(SQty, (i - 1) * 5 + 1, 5))
        '                    End If
        '                    S = S + 1
        '                Next i
        '                lblStrQty(c) = Str(S) & "@" & Str(T)
        '                lblStrQty(c).tag = Str(T)
        '            Else
        '                'Debug.Print aMessage
        '                Sql = "SELECT T3.*  FROM EDT850T1 T1, EDT850T2 T2, EDT850T3 T3 "
        '                Sql = Sql & "WHERE EDI_JRNL_NO = '" & jrnlno & "' "
        '                If aCodes(0, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_STYLE = '" & aCodes(0, 0) & "' "
        '                End If
        '                If aCodes(1, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_COLOR_CODE = '" & aCodes(1, 0) & "' "
        '                End If
        '                If aCodes(2, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_UPC = '" & aCodes(2, 0) & "' "
        '                End If
        '                If aCodes(3, 0) <> "" Then
        '                    Sql = Sql & "AND T2.EDI_SKU = '" & aCodes(3, 0) & "' "
        '                End If
        '                Sql = Sql & "AND T1.EDI_DOC_SEQ_NO = T2.EDI_DOC_SEQ_NO AND T3.EDI_DOC_SEQ_NO = T2.EDI_DOC_SEQ_NO "
        '                Sql = Sql & "AND T3.EDI_DTL_SEQ = T2.EDI_DTL_SEQ"
        '                dyn = OraD.CreateDynaset(Sql, 8&)
        '                SQty = ""
        '                Do While Not dyn.EOF
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_01").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_02").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_03").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_04").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_05").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_06").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_07").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_08").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_09").Value & "", "00000")
        '                    SQty = SQty & Format$(dyn.Fields("EDI_QTY_10").Value & "", "00000")
        '                    dyn.MoveNext()
        '                Loop
        '                dyn = Nothing
        '                SQty = Sort_List(SQty, 5)
        '                c = 0
        '                T = Val(Mid(SQty, 1, 5))
        '                S = 0
        '                For i = 1 To Len(SQty) / 5
        '                    If Val(Mid(SQty, (i - 1) * 5 + 1, 5)) <> T Then
        '                        lblStrQty(c) = Str(S) & "@" & Str(T)
        '                        lblStrQty(c).tag = Str(T)
        '                        c = c + 1
        '                        S = 0
        '                        T = Val(Mid(SQty, (i - 1) * 5 + 1, 5))
        '                    End If
        '                    S = S + 1
        '                Next i
        '                lblStrQty(c) = Str(S) & "@" & Str(T)
        '                lblStrQty(c).tag = Str(T)
        '            End If
        '        End If
        '    End If
        'End If
        'select_qty(0)
    End Sub

    Private Sub cmdUPC_Click(sender As System.Object, e As System.EventArgs) Handles cmdUPC.Click
        Dim UPC_CODE As String = ""
        Do
            Dim UPC_CODE_CTL_NO As String = ""
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("UPC_CODE")
            Else
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("ICTUPCH1.UPC_CODE")
            End If

            UPC_CODE = TAC.SOCMAIN1.UPC(Me, UPC_CODE_CTL_NO, ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"), True)
            If LookUp("ICTUPCH1", UPC_CODE) Is Nothing Then Exit Do
        Loop

        ASCMAIN1.sql = "Insert into ICTUPCH1 (UPC_CODE,STYLE_CODE,COLOR_CODE,INIT_DATE,INIT_OPER) " & vbCrLf _
            & " values (:PARM1,:PARM2,:PARM3,SYSDATE,:PARM4)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {UPC_CODE, RANGE_STYLE_CODE, "RNG", ASCMAIN1.USER_ID})

        Absx1.txtFor("RANGE_UPC_CODE").Text = UPC_CODE
        cmdUPC.Enabled = False
    End Sub

    Sub Load_QTYs_from_EDI(EDI_JRNL_NO As String)

        Dim sql_where As String = "" _
            & "(Select EDI_DOC_SEQ_NO from EDT850T1 where EDI_JRNL_NO in (Select EDI_JRNL_NO from EDTJRNL1" & vbCrLf _
            & " where DOCUMENT_ID = '850' and CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'" & vbCrLf _
            & " and EDI_JRNL_NO = '" & EDI_JRNL_NO & "'))"

        Dim sql1 As String = "Select Distinct EDI_DOC_SEQ_NO, EDI_DTL_SEQ from EDT850T2 where EDI_DOC_SEQ_NO in " & sql_where _
            & " and EDI_UPC in (" _
            & " Select UPC_CODE from ICVLUPC1 where STYLE_CODE = '" & Absx1.txtFor("RANGE_STYLE_CODE").Text & "'" _
            & " and UPC_CODE in (" _
            & " Select EDI_UPC from EDT850T2 where EDI_DOC_SEQ_NO in " & sql_where & "))"

        If EDT850TX = "" Then
            ASCMAIN1.sql = sql1
            EDT850TX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & EDT850TX & " Add Primary Key (EDI_DOC_SEQ_NO, EDI_DTL_SEQ)")
        Else
            ASCMAIN1.sql = "Truncate Table " & EDT850TX
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into " & EDT850TX & " " & sql1
            ASCDATA1.ExecuteSQL()
        End If

        Dim sql As String = "Select QTY, COUNT (*) STORES from (" & vbCrLf
        For i As Integer = 1 To 10
            If i <> 1 Then
                sql &= " union " & vbCrLf
            End If
            sql &= Replace("Select EDI_STORE_00 STORE, TO_NUMBER(NVL(EDI_QTY_00,0)) QTY ", "_00", "_" & Format(i, "00"))
            sql &= " from EDT850T3," & EDT850TX & " EDT850TX" & vbCrLf _
                & " where EDT850T3.EDI_DOC_SEQ_NO = EDT850TX.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and EDT850T3.EDI_DTL_SEQ = EDT850TX.EDI_DTL_SEQ" & vbCrLf
        Next i
        sql &= ") group by QTY"

        Dim rows() As DataRow = ASCDATA1.GetDataTable(sql).Select("ISNULL(QTY,0) <> 0", "QTY")
        ReDim QTYs(rows.Length)
        Dim j As Int32 = 0
        For Each row As DataRow In rows
            Dim QTY As Int64 = Val(row.Item("QTY") & "")
            Dim STORES As Int64 = Val(row.Item("STORES") & "")
            j += 1
            QTYs(j) = QTY
            With grdICTRSTYM.DisplayLayout.Bands(0).Columns("QTY_" & Format(j, "00"))
                .Hidden = False
                .Header.Caption = CStr(STORES) & "@" & CStr(QTY)
            End With
        Next
    End Sub

    Sub Add_Colors(STYLE_CODE As String, tbl As DataTable)
        If tbl.Select("ISNULL(SEL,'0')='1'").Length = 0 Then
            MsgBox("No Colors Selected", MsgBoxStyle.OkOnly, "Cannot Add Colors")
            Exit Sub
        End If

        For Each rowICTCOLRM As DataRow In tbl.Select("ISNULL(SEL,'0')='1'", "COLOR_CODE")
            grdICTRSTYM.DisplayLayout.Bands(0).AddNew()
            With grdICTRSTYM.ActiveRow
                .Cells("STYLE_CODE").Value = STYLE_CODE
                .Cells("COLOR_CODE").Value = rowICTCOLRM.Item("COLOR_CODE")

                .Update()
            End With
        Next
        Sort_grdColumns(grdICTRSTYM, "STYLE_CODE,COLOR_CODE")
    End Sub

    Sub Setup_QTY_COLs(RANGE_QTY_COL As Integer, Optional initialize As Boolean = False)
        If RANGE_QTY_COL > 0 Then
            For J As Integer = 1 To RANGE_QTY_COL
                With grdICTRSTYM.DisplayLayout.Bands(0).Columns("QTY_" & Format(J, "00"))
                    .Hidden = False
                End With
            Next
        End If

        For j As Integer = RANGE_QTY_COL + 1 To COLs_max
            With grdICTRSTYM.DisplayLayout.Bands(0).Columns("QTY_" & Format(j, "00"))
                .Hidden = True
                .Header.Caption = "{" & Format(j, "00") & "}"
                .Width = 75
            End With
            If initialize Then Create_Summary(grdICTRSTYM, "QTY_" & Format(j, "00"))
        Next
    End Sub

    Private Sub grdEDT850TX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT850TX.DoubleClickRow
        If (EntryMode = "N" Or EntryMode = "E") Then
            Dim EDI_JRNL_NO As String = e.Row.Cells("EDI_JRNL_NO").Value & ""
            Load_QTYs_from_EDI(EDI_JRNL_NO)
            Set_Qty()
        End If
    End Sub

    Sub Copy_Range_Style(CUST_CODE As String, RANGE_STYLE_CODE As String)
        Dim RANGE_STYLE_CODE_NEW As String = ASCMAIN1.Get_txt_from_User _
                ("Enter New Range Style Code", "Copy Range Style", False, 12, RANGE_STYLE_CODE)

        If RANGE_STYLE_CODE_NEW = "" Then Exit Sub

        Dim row As DataRow = LookUp("ICTRSTY1", New String() {CUST_CODE, RANGE_STYLE_CODE_NEW})
        If row IsNot Nothing Then
            MsgBox("Range Style " & RANGE_STYLE_CODE_NEW & " already exists for Customer " & CUST_CODE, _
                   MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        If Not ASCMAIN1.Logical_Lock("ICTRSTY1", CUST_CODE & "||" & RANGE_STYLE_CODE_NEW) Then Exit Sub


        If ICTRSTY1 = "" Then
            ICTRSTY1 = ASCMAIN1.Temp_Table("Select * from ICTRSTY1 where ROWNUM < 1")
            ICTRSTY2 = ASCMAIN1.Temp_Table("Select * from ICTRSTY2 where ROWNUM < 1")
        End If


        ASCMAIN1.sql = "Truncate Table " & ICTRSTY1
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Truncate Table " & ICTRSTY2
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into " & ICTRSTY1 & " Select * from ICTRSTY1 where CUST_CODE = :PARM1 and RANGE_STYLE_CODE = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {CUST_CODE, RANGE_STYLE_CODE})
        ASCMAIN1.sql = "Update " & ICTRSTY1 & " Set RANGE_STYLE_CODE = :PARM1, LAST_OPER = NULL, LAST_DATE = NULL, INIT_OPER = :PARM2, INIT_DATE = SYSDATE where RANGE_STYLE_CODE = :PARM3"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {RANGE_STYLE_CODE_NEW, ASCMAIN1.USER_ID, RANGE_STYLE_CODE})
        ASCMAIN1.sql = "Insert into " & ICTRSTY2 & " Select * from ICTRSTY2 where CUST_CODE = :PARM1 and RANGE_STYLE_CODE = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {CUST_CODE, RANGE_STYLE_CODE})
        ASCMAIN1.sql = "Update " & ICTRSTY2 & " Set RANGE_STYLE_CODE = :PARM1 where RANGE_STYLE_CODE = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {RANGE_STYLE_CODE_NEW, RANGE_STYLE_CODE})
        ASCMAIN1.sql = "Insert into ICTRSTY1 Select * from " & ICTRSTY1
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into ICTRSTY2 Select * from " & ICTRSTY2
        ASCDATA1.ExecuteSQL()

        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("RANGE_STYLE_CODE").Text = RANGE_STYLE_CODE_NEW
        Click_Command("Edit")

    End Sub

    Sub Load_MultiRange()

        Try
            Dim FILENAME As String = ""
            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    FILENAME = openFileDialog1.FileName
                End If
            End Using

            If FILENAME <> "" Then

                grdICTRSTYR.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdICTRSTYR.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdICTRSTYR.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

                Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                Dim range As SpreadsheetGear.IRange = Nothing
                range = oSheet.UsedRange
                Dim rmax As Integer = range.RowCount

                Dim TYPE As String = ""
                Dim started As Boolean = False

                Dim INVALID_STYLES As String = ""
                Dim Invalid_UPCS As String = ""
                Dim Duplicate_Range As String = ""

                grdICTRSTYR.Visible = False
                ASCMAIN1.Progress("Now Loading ...")

                Dim r As Integer = 0
                Do While r < rmax ' oSheet.Cells(r, 0).Value & "" <> ""

                    'If TYPE = "" And Not started Then
                    '    If Trim(oSheet.Cells(r, 0).Value & "").ToUpper = "MEIJER" Then
                    '        TYPE = "MEIJER"
                    '    End If
                    '    If Trim(oSheet.Cells(r, 0).Value & "").ToUpper = "SHOPKO" Then
                    '        TYPE = "SHOPKO"
                    '    End If
                    'End If
                    If oSheet.Cells(r, 0).Value & "" <> CUST_CODE Then
                        r += 1
                        Continue Do
                    End If

                    Dim RANGE_CUST_CODE As String = ""
                    Dim RANGE_STYLE_CODE As String = ""
                    Dim RANGE_STYLE_DESC As String = ""
                    Dim RANGE_UPC_CODE As String = ""
                    Dim RANGE_SKU As String = ""
                    Dim RNG_AST_FLG As String = "F"

                    Dim STYLE_CODE As String = ""
                    Dim COLOR_CODE As String = ""
                    Dim SIZE_CODE As String = ""
                    Dim CUST_SKU As String = ""
                    Dim CUST_UPC As String = ""
                    Dim STYLE_PRICE As Decimal = 0
                    Dim STYLE_QTY As Int64 = 0

                    Dim LstRngUPC As String = ""
                    Dim lstRngStyle As String = ""
                    Dim lstRngUPCChkDigit As String = ""
                    Dim ERR_TYPE As String = "0"

                    RANGE_CUST_CODE = oSheet.Cells(r, 0).Value & ""
                    RANGE_STYLE_CODE = oSheet.Cells(r, 1).Value & ""
                    RANGE_STYLE_DESC = oSheet.Cells(r, 2).Value & ""
                    RANGE_UPC_CODE = oSheet.Cells(r, 3).Value & ""
                    RANGE_SKU = oSheet.Cells(r, 4).Value & ""
                    RNG_AST_FLG = "F"
                    STYLE_CODE = oSheet.Cells(r, 6).Value & ""
                    COLOR_CODE = oSheet.Cells(r, 7).Value & ""
                    SIZE_CODE = oSheet.Cells(r, 8).Value & ""
                    STYLE_PRICE = Val(oSheet.Cells(r, 9).Value & "")
                    STYLE_QTY = Val(oSheet.Cells(r, 10).Value & "")
                    CUST_SKU = oSheet.Cells(r, 11).Value & ""
                    CUST_UPC = oSheet.Cells(r, 12).Value & ""

                    'If STYLE_QTY = 0 Then STYLE_CODE = ""
                    'AndAlso LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}) IsNot Nothing 
                    If COLOR_CODE = "" Then
                        Dim Sql As String = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
                        Dim rows() As DataRow = ASCDATA1.GetDataTable(Sql).Select("")
                        If rows.Length = 1 Then
                            COLOR_CODE = rows(0).Item("COLOR_CODE") & ""
                        Else

                        End If
                    End If

                    If RANGE_UPC_CODE & RANGE_STYLE_CODE <> LstRngUPC & lstRngStyle Then
                        lstRngStyle = RANGE_STYLE_CODE
                        LstRngUPC = RANGE_UPC_CODE
                        lstRngUPCChkDigit = ""
                        'Check for existing Range
                        Dim RngSql As String = "Select * from ICTRSTY1 where CUST_CODE = '" & CUST_CODE & "' and RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "'"
                        Dim Rngrows() As DataRow = ASCDATA1.GetDataTable(RngSql).Select("")
                        If Rngrows.Length > 0 Then
                            Duplicate_Range &= ", " & CUST_CODE & "-" & RANGE_STYLE_CODE
                            ERR_TYPE = "1"
                        Else
                            If RANGE_UPC_CODE <> "" Then
                                If RANGE_UPC_CODE.Length = 11 Then
                                    lstRngUPCChkDigit = ASCMAIN1.CheckDigitUPC(RANGE_UPC_CODE)
                                End If
                                Dim Sql As String = "Select * from ICVLUPC1 where UPC_CODE = '" & RANGE_UPC_CODE & lstRngUPCChkDigit & "'"
                                Dim rows() As DataRow = ASCDATA1.GetDataTable(Sql).Select("")
                                If rows.Length = 1 Then
                                    If (RANGE_STYLE_CODE <> rows(0).Item("STYLE_CODE") & "") Then
                                        Invalid_UPCS &= ", " & RANGE_STYLE_CODE & "-" & RANGE_UPC_CODE
                                        ERR_TYPE = "1"
                                    End If
                                Else
                                    Invalid_UPCS &= ", " & RANGE_STYLE_CODE & "-" & RANGE_UPC_CODE
                                    ERR_TYPE = "1"
                                End If
                            End If
                        End If
                    End If

                    If RANGE_UPC_CODE = "" Then
                        lstRngUPCChkDigit = ""
                    End If

                    If LookUp("ICTSTYL1", STYLE_CODE) IsNot Nothing AndAlso LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}) IsNot Nothing Then
                        'If grdICTRSTYR.ActiveRow IsNot Nothing AndAlso grdICTRSTYR.ActiveRow.IsAddRow Then
                        '    grdICTRSTYR.ActiveRow.CancelUpdate()
                        'End If

                        'grdICTRSTYR.DisplayLayout.Bands(0).AddNew.Activate()

                        Dim rICTRSTYR As DataRow = dst.Tables("ICTRSTYR").NewRow
                        With rICTRSTYR
                            .Item("CUST_CODE") = RANGE_CUST_CODE
                            .Item("RANGE_STYLE_CODE") = RANGE_STYLE_CODE
                            .Item("RANGE_STYLE_DESC") = RANGE_STYLE_DESC
                            .Item("RANGE_UPC_CODE") = RANGE_UPC_CODE & lstRngUPCChkDigit
                            .Item("RANGE_SKU") = RANGE_SKU
                            .Item("RNG_AST_FLG") = RNG_AST_FLG

                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            .Item("SIZE_CODE") = SIZE_CODE
                            .Item("CUST_SKU") = CUST_SKU
                            .Item("CUST_UPC") = CUST_UPC
                            .Item("STYLE_PRICE") = STYLE_PRICE
                            .Item("STYLE_QTY") = STYLE_QTY
                            .Item("ERR_TYPE") = ERR_TYPE
                        End With
                        dst.Tables("ICTRSTYR").Rows.Add(rICTRSTYR)

                    Else
                        If STYLE_CODE <> "" And COLOR_CODE <> "" Then INVALID_STYLES &= ", " & STYLE_CODE & "-" & COLOR_CODE
                        ERR_TYPE = "1"
                    End If

                    If (r Mod 10) = 1 Then
                        ASCMAIN1.Progress("-", RANGE_STYLE_CODE)
                    End If

                    r += 1
                Loop

                grdICTRSTYR.Visible = True

                Sort_grdColumns(grdICTRSTYR, "err_type,RANGE_STYLE_CODE,STYLE_CODE,COLOR_CODE")

                If INVALID_STYLES <> "" Then
                    MsgBox(Mid(INVALID_STYLES, 3), MsgBoxStyle.OkOnly, "The following Style-Color codes were invalid")
                End If

                If Invalid_UPCS <> "" Then
                    MsgBox(Mid(Invalid_UPCS, 3), MsgBoxStyle.OkOnly, "The following Range UPC codes were invalid")
                End If
                If Duplicate_Range <> "" Then
                    MsgBox(Mid(Duplicate_Range, 3), MsgBoxStyle.OkOnly, "The following Customer Range Styles already exist")
                End If

                If Invalid_UPCS & INVALID_STYLES & Duplicate_Range <> "" Then
                    MsgBox("There were errors loading the data, please correct the spreadsheet data and try again", MsgBoxStyle.Critical, "Update not allowed")
                    'disable update
                    UltraExplorerBar1.Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                Else
                    MsgBox("Spreadsheet has been Loaded", MsgBoxStyle.OkOnly, "Success")
                End If


                grdICTRSTYR.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdICTRSTYR.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdICTRSTYR.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

                ASCMAIN1.Progress("")
            End If
        Catch ex As Exception
            MsgBox("Error " & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load this Spreadsheet")
        End Try
    End Sub

    Sub Update_Range()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        Dim last_range As String = ""
        For Each row As DataRow In dst.Tables("ICTRSTYR").Select("")
            If last_range <> row.Item("RANGE_STYLE_CODE") Then
                last_range = row.Item("RANGE_STYLE_CODE")
                Dim RANGE_PRICE As Decimal = dst.Tables("ICTRSTYR").Compute("SUM(STYLE_AMT)", "RANGE_STYLE_CODE = '" & last_range & "'")
                Dim RANGE_QTY As Int32 = dst.Tables("ICTRSTYR").Compute("SUM(STYLE_QTY)", "RANGE_STYLE_CODE = '" & last_range & "'")
                Dim rICTRSTY1 As DataRow = dst.Tables("ICTRSTY1").NewRow
                With rICTRSTY1
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("RANGE_STYLE_CODE") = last_range
                    .Item("RANGE_PRICE") = RANGE_PRICE
                    .Item("RANGE_QTY") = RANGE_QTY
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("RANGE_STYLE_DESC") = row.Item("RANGE_STYLE_DESC")
                    .Item("RANGE_UPC_CODE") = row.Item("RANGE_UPC_CODE")
                    .Item("RANGE_SKU") = row.Item("RANGE_SKU")
                    .Item("RANGE_STYLE_QTY_PER_PP") = RANGE_QTY
                    .Item("RNG_AST_FLG") = row.Item("RNG_AST_FLG")
                End With
                dst.Tables("ICTRSTY1").Rows.Add(rICTRSTY1)
                ASCMAIN1.Progress("", "Range Style: " & last_range)
            End If

            Dim rICTRSTY2 As DataRow = dst.Tables("ICTRSTY2").NewRow
            With rICTRSTY2
                .Item("CUST_CODE") = CUST_CODE
                .Item("RANGE_STYLE_CODE") = last_range
                .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                .Item("SIZE_CODE") = row.Item("SIZE_CODE")
                .Item("STYLE_PRICE") = row.Item("STYLE_PRICE")
                .Item("STYLE_QTY") = row.Item("STYLE_QTY")
                .Item("CUST_SKU") = row.Item("CUST_SKU")
                .Item("CUST_UPC") = row.Item("CUST_UPC")
            End With
            dst.Tables("ICTRSTY2").Rows.Add(rICTRSTY2)

        Next
        'MsgBox("Ranges have been created", MsgBoxStyle.OkOnly, "Success")

        BeginTrans()

        Update_Record_TDA("ICTRSTY1")
        Update_Record_TDA("ICTRSTY2")

        CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub


    Private Sub grdICTRSTYR_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTRSTYR.InitializeRow
        If e.Row.Band.Index = 0 Then
            With e.Row.Cells("ERR_TYPE")
                If .Value = "1" Then
                    e.Row.Appearance.ForeColor = Drawing.Color.Red
                End If
            End With
        End If
    End Sub
End Class