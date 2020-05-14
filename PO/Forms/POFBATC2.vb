Public Class POFBATC2

    ' allow selection of vendors to create pos for some vendors while still working on other vendors
    ' batch needs a delete remaining option

    Dim PO_BATCH_NO As String
    Dim WHSE_CODE As String
    Dim rowPOTBATC1 As DataRow
    Dim rowICTWHSE1 As DataRow

    Dim SOTORDC1 As String
    Dim SOTSLSC1 As String
    Dim POTBATC2 As String
    Dim sqlSOTORDC1 As String
    Dim sqlSOTSLSC1 As String
    Dim TC As New Dictionary(Of String, Dictionary(Of String, String))
    Dim STYLE_CLASS_CODEs As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")

        Create_Temp_Tables()

        With dst
            ASCMAIN1.sql = "Select POTBATC1.*" & vbCrLf _
            & " from POTBATC1 " & vbCrLf
            Create_TDA(.Tables.Add, "POTBATCX", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select POTBATC1.*" & vbCrLf _
            & " from POTBATC1 " & vbCrLf _
            & " where POTBATC1.PO_BATCH_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTBATC1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select :PARM1 PO_BATCH_NO, STYLE_CODE, STYLE_DESC, STYLE_STATUS, STYLE_UOM, STYLE_COST" _
                & ", CARTON_PACK_QTY, CASE_CUBE, MIN_ORDER_QTY, VEND_CODE, STYLE_CLASS_CODE" _
            & " from ICTSTYL1 where STYLE_CODE in (Select STYLE_CODE from " & POTBATC2 & ")"
            Create_TDA(.Tables.Add, "POTBATCS", "**", 0, False, "V", 2)
            .Tables("POTBATCS").Columns("PO_BATCH_NO").MaxLength = 6

            ASCMAIN1.sql = "Select POTBATC2.*" & vbCrLf _
            & " from " & POTBATC2 & " POTBATC2 "
            Create_TDA(.Tables.Add, "POTBATC2", "**", 0, True)

            Create_Relation("POTBATCS", "POTBATC2", "PO_BATCH_NO,STYLE_CODE")
            With .Tables("POTBATC2").Columns
                .Add("CUST_SOLD", GetType(System.Int32))
                .Add("QTY_SOLD", GetType(System.Int32))
                .Add("AMT_SOLD", GetType(System.Decimal))
                .Add("CASE_CUBE", GetType(System.Int32), "PARENT(POTBATCS_POTBATC2).CASE_CUBE")
                .Add("NET_POS", GetType(System.Int32), "ISNULL(QTY_ONH,0)+ISNULL(QTY_PO,0)-ISNULL(QTY_OPEN,0)")
                .Add("QTY_SHORT", GetType(System.Int32), "IIF(NET_POS<0,-1*NET_POS,NULL)")
                .Add("CASE_QTY", GetType(System.Int32), "IIF(ISNULL(PARENT(POTBATCS_POTBATC2).CARTON_PACK_QTY,0)>0,ISNULL(PARENT(POTBATCS_POTBATC2).CARTON_PACK_QTY,0),1)")
                .Add("CUBE", GetType(System.Decimal), "CASE_CUBE * PO_QTY_ROUNDED / CASE_QTY")
                .Add("PO_QTY_CALC", GetType(System.Int32), "PO_QTY + (CASE_QTY - PO_QTY%CASE_QTY)%CASE_QTY")
                .Add("NET_POS2", GetType(System.Int32), "NET_POS + PO_QTY_CALC")
            End With

            With .Tables("POTBATCS").Columns
                .Add("CUBE", GetType(System.Int32), "SUM(CHILD.CUBE)")
                .Add("QTY_SHORT", GetType(System.Int32), "SUM(CHILD.QTY_SHORT)")
                .Add("PO_QTY_ROUNDED", GetType(System.Int32), "SUM(CHILD.PO_QTY_ROUNDED)")
                .Add("TOTAL_COST", GetType(System.Int32), "STYLE_COST * PO_QTY_ROUNDED")
            End With

            ASCMAIN1.sql = "Select POTBATC3.*" & vbCrLf _
                & " from POTBATC3 where PO_BATCH_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTBATC3", "**", 0, True, "V")

            Create_Relation("POTBATC3", "POTBATCS", "VEND_CODE")
            With .Tables("POTBATC3").Columns
                .Add("CUBE", GetType(System.Int32), "SUM(CHILD.CUBE)")
                .Add("TOTAL_COST", GetType(System.Int32), "SUM(CHILD.TOTAL_COST)")
            End With

            ASCMAIN1.sql = "Select POTBATC4.*, ICTCLAS1.STYLE_CLASS_DESC" & vbCrLf _
                & " from POTBATC4,ICTCLAS1" & vbCrLf _
                & " where POTBATC4.PO_BATCH_NO = :PARM1" & vbCrLf _
                & "   and ICTCLAS1.STYLE_CLASS_CODE = POTBATC4.STYLE_CLASS_CODE"
            Create_TDA(.Tables.Add, "POTBATC4", "**", 0, True, "V")

            ASCMAIN1.sql = "Select POTBATC5.*" & vbCrLf _
                & " from POTBATC5 where PO_BATCH_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTBATC5", "**", 0, True, "V")


            ASCMAIN1.sql = "Select SOTSLSC1.*" & vbCrLf _
            & " from " & SOTSLSC1 & " SOTSLSC1 " & vbCrLf _
            & " where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add, "SOTSLSC1", "**", 0, False, "VV", 0)

            ASCMAIN1.sql = "Select SOTORDC1.*" & vbCrLf _
            & " from " & SOTORDC1 & " SOTORDC1 " & vbCrLf _
            & " where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
            Create_TDA(.Tables.Add, "SOTORDC1", "**", 0, False, "VV", 0)

            ASCMAIN1.sql = "Select ICTCLAS1.STYLE_CLASS_CODE, ICTCLAS1.STYLE_CLASS_DESC" & vbCrLf _
            & " from ICTCLAS1" & vbCrLf
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)
            .Tables("ICTCLAS1").Columns.Add("SELECTED")

        End With

        grdPOTBATCX.DataSource = dst.Tables("POTBATCX")
        grdPOTBATC2.DataSource = dst.Tables("POTBATCS")
        grdPOTBATC3.DataSource = dst.Tables("POTBATC3")
        grdSOTSLSC1.DataSource = dst.Tables("SOTSLSC1")
        grdSOTORDC1.DataSource = dst.Tables("SOTORDC1")
        grdICTCLAS1.DataSource = dst.Tables("ICTCLAS1")
        grdPOTBATC4.DataSource = dst.Tables("POTBATC4")

        Fill_Records("ICTCLAS1")
        Sort_grdColumns(grdICTCLAS1, "STYLE_CLASS_DESC", False)


        grdPOTBATC2.DisplayLayout.UseFixedHeaders = True
        With grdPOTBATC2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "STYLE_UOM", "STYLE_STATUS", "VEND_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdPOTBATC2.DisplayLayout.Bands(1)
            For Each COLUMN_NAME As String In New String() {"COLOR_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdPOTBATC2.DisplayLayout.Bands(0).Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False
        End With

        With grdPOTBATC2.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "PO_QTY" Then
                    gcol.CellAppearance.BackColor = Drawing.Color.LightYellow
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    '  gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                If New String() {"CUST_SOLD", "QTY_SOLD", "AMT_SOLD"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"PO_QTY", "PO_QTY_CALC", "PO_QTY_ROUNDED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.CellAppearance.BackColor = Drawing.Color.Orange
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"CUST_OPEN", "QTY_OPEN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.CellAppearance.BackColor = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_COLOR_STATUS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.CellAppearance.BackColor = Drawing.Color.Pink
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"QTY_SHORT", "NET_POS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"QTY_ONH", "QTY_PO", "QTY_OPEN", "CUST_OPEN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellAppearance.BackColor = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
                End If
            Next
        End With

        With grdPOTBATC3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PO_NOTES", "PO_DATE_SHIP_BY", "PO_DATE_ETA"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Create_Summary(grdPOTBATCX, "PO_BATCH_NO", "Count")

        Create_Summary(grdPOTBATC2, "STYLE_CODE", "Count")
        Create_Summary(grdPOTBATC2, New String() {"CUBE", "QTY_SHORT", "PO_QTY_ROUNDED", "TOTAL_COST"}, , "POTBATCS")

        Create_Summary(grdPOTBATC2, "COLOR_CODE", "Count", "POTBATCS_POTBATC2")
        Create_Summary(grdPOTBATC2, New String() {"PO_QTY", "PO_QTY_CALC"}, , "POTBATCS_POTBATC2")

        Create_Summary(grdPOTBATC3, "VEND_CODE", "Count")
        Create_Summary(grdPOTBATC3, New String() {"CUBE", "TOTAL_COST"})

        Create_Summary(grdSOTORDC1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDC1, New String() {"QTY", "AMT"})

        Create_Summary(grdSOTSLSC1, "CUST_CODE", "Count")
        Create_Summary(grdSOTSLSC1, New String() {"QTY", "AMT"})

        splPOTBATCA.Panel2Collapsed = True

        dteFrom.Value = Now.Date.AddMonths(-3)
        dteTo.Value = Now.Date

        grdICTCLAS1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        With grdICTCLAS1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 IsNot Nothing Then
                        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    Else
                        EMsg &= vbCr & "No Record of Whse " & Absx1.txtFor("WHSE_CODE").Text
                    End If
                End If

                STYLE_CLASS_CODEs.Clear()
                For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select("SELECTED = '1'")
                    Dim STYLE_CLASS_CODE As String = rowICTCLAS1.Item("STYLE_CLASS_CODE")
                    STYLE_CLASS_CODEs.Add(STYLE_CLASS_CODE)
                Next
                If STYLE_CLASS_CODEs.Count = 0 Then
                    EMsg &= vbCr & "No Item Classes Selected"
                Else
                    ASCMAIN1.sql = "Select Min(POTBATC1.PO_BATCH_NO) PO_BATCH_NO" _
                        & "  from POTBATC1,POTBATC4 " _
                        & " where POTBATC4.PO_BATCH_NO = POTBATC1.PO_BATCH_NO" _
                        & "   and POTBATC1.BATCH_STATUS = 'O'" _
                        & "   and POTBATC4.STYLE_CLASS_CODE in ('" & Join(STYLE_CLASS_CODEs.ToArray, "','") & "')"
                    Dim PO_BATCH_NO_in_use As String = ASCDATA1.GetDataValue
                    If PO_BATCH_NO_in_use <> "" Then
                        EMsg &= vbCr & "Batch is Already in progress with some of the Item Classes Selected (see " & PO_BATCH_NO_in_use & ")"""
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("POTBATC1", "WHSE_CODE:" & WHSE_CODE) Then Exit Sub
                    For Each STYLE_CLASS_CODE As String In STYLE_CLASS_CODEs
                        If Not ASCMAIN1.Logical_Lock("POTBATC1", "STYLE_CLASS_CODE:" & STYLE_CLASS_CODE) Then Exit Sub
                    Next
                End If

            Case "Edit", "Load"

                WHSE_CODE = ""
                PO_BATCH_NO = ""

                If Absx1.txtFor("PO_BATCH_NO").Text = "" Then
                    EMsg &= vbCr & "No  Batch No Specified"
                Else
                    PO_BATCH_NO = Absx1.txtFor("PO_BATCH_NO").Text
                    rowPOTBATC1 = LookUp("POTBATC1", PO_BATCH_NO)
                    If rowPOTBATC1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Batch No " & PO_BATCH_NO
                    Else
                        WHSE_CODE = rowPOTBATC1.Item("WHSE_CODE")
                        If rowPOTBATC1.Item("BATCH_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowPOTBATC1.Item("BATCH_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " has been Cancelled"
                                Case "D"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " has been Deleted"
                                Case Else ' such as "F"
                                    EMsg &= vbCr & "Batch No " & PO_BATCH_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" Then
                    STYLE_CLASS_CODEs.Clear()
                    ASCMAIN1.sql = "Select STYLE_CLASS_CODE from POTBATC4 where PO_BATCH_NO = '" & PO_BATCH_NO & "'"
                    For Each rowPOTBATC4 As DataRow In ASCDATA1.GetDataTable.Rows
                        Dim STYLE_CLASS_CODE As String = rowPOTBATC4.Item("STYLE_CLASS_CODE")
                        STYLE_CLASS_CODEs.Add(STYLE_CLASS_CODE)
                    Next
                End If

                If EMsg = "" And EntryMode = "E" Then
                    If Not ASCMAIN1.Logical_Lock("POTBATC1", PO_BATCH_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("POTBATC1", "WHSE_CODE:" & WHSE_CODE) Then Exit Sub
                    For Each STYLE_CLASS_CODE As String In STYLE_CLASS_CODEs
                        If Not ASCMAIN1.Logical_Lock("POTBATC1", "STYLE_CLASS_CODE:" & STYLE_CLASS_CODE) Then Exit Sub
                    Next
                End If

            Case "Cancel"
                If MsgBox("Are you sure that you want to Cancel?", MsgBoxStyle.YesNo, "Verification to Cancel Changes Made to this Batch") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"
                'If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                '    Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                '    EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                'Else
                '    If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                '     > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                '        EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                '    End If
                'End If

                If grdPOTBATC2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Styles on Batch"
                Else
                    If Val(dst.Tables("POTBATC2").Compute("COUNT(STYLE_CODE)", "PO_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Styles on Batch with PO Qty >0"
                    End If
                End If

            Case "Delete"
                If EMsg = "" Then
                    If MsgBox("Do you want to Mark this Batch as Deleted", _
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
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

            Case "Done"
                Mode_Settings(False)

            Case "Scan for New Styles"
                Scan_for_New_Styles()
        End Select

    End Sub


    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "L" And ScreenMode) Then
                        If rowPOTBATC1.Item("BATCH_STATUS") & "" = "O" Then
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                        Else
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                        End If
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Generate PO").Visible = (EntryMode = "N" Or EntryMode = "E")

                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    .Items("Print").Visible = ScreenMode
                    .Items("Update").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                    .Items("Delete").Visible = (EntryMode = "E")
                    .Items("Cancel").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                End With
                .Groups("Sales History").Visible = ScreenMode
                .Groups("Style Filters").Visible = ScreenMode
                .Groups("Item Class Selection").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        lblStatus.Visible = ScreenMode
        grdPOTBATCX.Visible = Not ScreenMode
        splPOTBATCA.Visible = ScreenMode

        If ScreenMode Then
            With grdPOTBATC2.DisplayLayout.Override
                If EntryMode = "L" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False

                    'With grdPOTBATC2.DisplayLayout.Bands(0)
                    '    If EntryMode <> "E" Then
                    '        .Columns("X").Hidden = True
                    '    Else
                    '        .Columns("X").Hidden = False
                    '    End If
                    'End With
                End If

            End With
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTBATC1", "POTBATC2", "POTBATC3", "POTBATC4", "POTBATC5", "POTBATCS", "SOTORDC1", "SOTSLSC1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select("")
            rowICTCLAS1.Item("SELECTED") = "0"
        Next
        Load_POTBATCX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        grdPOTBATC2.BeginUpdate()
        grdPOTBATC3.BeginUpdate()

        TC.Clear()
        For Each TABLE_NAME As String In New String() {"POTBATCS", "POTBATC1", "POTBATC3"}
            With dst.Tables(TABLE_NAME)
                Dim CE As New Dictionary(Of String, String)
                For c As Integer = .Columns.Count - 1 To 0 Step -1
                    If .Columns(c).Expression <> "" Then
                        CE.Add(.Columns(c).ColumnName, .Columns(c).Expression)
                        .Columns(c).Expression = ""
                    End If
                Next
                TC.Add(TABLE_NAME, CE)
            End With
        Next

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ASCDATA1.ExecuteSQL("Truncate Table " & POTBATC2)

        If EntryMode = "N" Then
            PO_BATCH_NO = ASCMAIN1.Next_Control_No("POTBATC1.PO_BATCH_NO")

            rowPOTBATC1 = dst.Tables("POTBATC1").NewRow
            With rowPOTBATC1
                .Item("PO_BATCH_NO") = PO_BATCH_NO
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("BATCH_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("POTBATC1").Rows.Add(rowPOTBATC1)

            ASCMAIN1.sql = "Select '" & PO_BATCH_NO & "' PO_BATCH_NO" _
                & ", ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_STATUS" _
                & " from ICTSTYL1,ICTSTYC1,ICTCOLR1" _
                & " where ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                & "   and ICTSTYL1.STYLE_CLASS_CODE in ('" & Join(STYLE_CLASS_CODEs.ToArray, "','") & "')" _
                & "   and ICTCOLR1.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & POTBATC2 _
                                & " (PO_BATCH_NO,STYLE_CODE,COLOR_CODE,COLOR_DESC,STYLE_COLOR_STATUS) " _
                                & ASCMAIN1.sql)

            For Each STYLE_CLASS_CODE As String In STYLE_CLASS_CODEs
                Dim rowICTCLAS1 As DataRow = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
                Dim rowPOTBATC4 As DataRow = dst.Tables("POTBATC4").NewRow
                rowPOTBATC4.Item("PO_BATCH_NO") = PO_BATCH_NO
                rowPOTBATC4.Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                rowPOTBATC4.Item("STYLE_CLASS_DESC") = rowICTCLAS1.Item("STYLE_CLASS_DESC")
                dst.Tables("POTBATC4").Rows.Add(rowPOTBATC4)

                ASCMAIN1.sql = "Select * from ICTCLAS2 where STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'"
                For Each rowICTCLAS2 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim rowPOTBATC5 As DataRow = dst.Tables("POTBATC5").NewRow
                    rowPOTBATC5.Item("PO_BATCH_NO") = PO_BATCH_NO
                    rowPOTBATC5.Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                    rowPOTBATC5.Item("CUST_COUNT") = rowICTCLAS2.Item("CUST_COUNT")
                    rowPOTBATC5.Item("PCT_INCREASE") = rowICTCLAS2.Item("PCT_INCREASE")
                    dst.Tables("POTBATC5").Rows.Add(rowPOTBATC5)
                Next
            Next
        Else
            ASCMAIN1.sql = "Select POTBATC2.*, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_STATUS" _
                & " from POTBATC2,ICTCOLR1,ICTSTYC1" _
                & " where ICTCOLR1.COLOR_CODE (+) = POTBATC2.COLOR_CODE" _
                & "   and POTBATC2.PO_BATCH_NO = '" & PO_BATCH_NO & "'" _
                & "   and ICTSTYC1.STYLE_CODE = POTBATC2.STYLE_CODE" _
                & "   and ICTSTYC1.COLOR_CODE = POTBATC2.COLOR_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & POTBATC2 & " " & ASCMAIN1.sql)

            rowPOTBATC1 = Fill_Record("POTBATC1", PO_BATCH_NO)

            rowPOTBATC1 = Fill_Record("POTBATC4", PO_BATCH_NO)
            rowPOTBATC1 = Fill_Record("POTBATC5", PO_BATCH_NO)
        End If

        If EntryMode = "N" Or EntryMode = "E" Then
            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is" _
                & "  Select * from ICTSTAT2 where (STYLE_CODE, COLOR_CODE) in" _
                & "   (Select STYLE_CODE, COLOR_CODE from " & POTBATC2 & ");" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Update " & POTBATC2 & " Set QTY_ONH = NVL(R1.WHSE_QTY_ON_HAND,0)" _
                & "    , QTY_OPEN = NVL(R1.WHSE_QTY_OPEN,0) + NVL(R1.WHSE_QTY_PICK,0), QTY_PO = NVL(R1.WHSE_QTY_ON_ORDER,0)" _
                & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is" _
                & "  Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, Count (Distinct SOTORDR1.CUST_CODE) CUST_OPEN" _
                & "   from SOTORDR2,SOTORDR1 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                & "   and SOTORDR1.ORDR_STATUS >= 'O' and SOTORDR1.ORDR_STATUS <= 'P'" _
                & "   and (SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE) in" _
                & "   (Select STYLE_CODE, COLOR_CODE from " & POTBATC2 & ")" _
                & "   group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Update " & POTBATC2 & " Set CUST_OPEN = R1.CUST_OPEN" _
                & "    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        Fill_Records("POTBATC3", PO_BATCH_NO)

        Fill_Records("POTBATCS", PO_BATCH_NO)
        For Each rowPOTBATCS As DataRow In dst.Tables("POTBATCS").Select("ISNULL(VEND_CODE,'') = ''")
            rowPOTBATCS.Item("VEND_CODE") = "."
        Next

        For Each rowVEND_CODE As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTBATCS"), New String() {"VEND_CODE"}).Rows
            Dim VEND_CODE As String = rowVEND_CODE.Item("VEND_CODE")
            Dim rowPOTBATC3 As DataRow = dst.Tables("POTBATC3").Rows.Find(New String() {PO_BATCH_NO, VEND_CODE})
            If rowPOTBATC3 Is Nothing Then
                rowPOTBATC3 = dst.Tables("POTBATC3").NewRow
                rowPOTBATC3.Item("PO_BATCH_NO") = PO_BATCH_NO
                rowPOTBATC3.Item("VEND_CODE") = VEND_CODE
                'rowPOTBATC3.Item("") = ""
                dst.Tables("POTBATC3").Rows.Add(rowPOTBATC3)
            End If
        Next
        Sort_grdColumns(grdPOTBATC3, "VEND_CODE")

        WHSE_CODE = rowPOTBATC1.Item("WHSE_CODE")
        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)


        Fill_Records("POTBATC2")
        If EntryMode = "N" Then
            For Each rowPOTBATC2 As DataRow In dst.Tables("POTBATC2").Select("QTY_SHORT > 0")
                rowPOTBATC2.Item("PO_QTY") = rowPOTBATC2.Item("QTY_SHORT")
                rowPOTBATC2.Item("PO_QTY_ROUNDED") = rowPOTBATC2.Item("PO_QTY_CALC")
            Next
        End If

        Sort_grdColumns(grdPOTBATC2, "STYLE_CODE")

        Setup_SOTSLSC1()
        Setup_SOTORDC1()

        If EntryMode = "N" Then
            lblStatus.Text = "New Batch"
        Else
            Select Case rowPOTBATC1.Item("BATCH_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Closed"
                Case "D"
                    lblStatus.Text = "Deleted"
            End Select
        End If

        With grdPOTBATC2.DisplayLayout.Bands(1)
            If (EntryMode = "E" Or EntryMode = "N") Then
                .Columns("PO_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                'If EntryMode = "E" Then
                '    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                '    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                'End If
            Else
                .Columns("PO_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                '.Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdPOTBATC2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.False
            End With
            'grdPOTBATC2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            With grdSOTRSRV2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            'grdPOTBATC2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        End If

        'Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        For Each TABLE_NAME As String In New String() {"POTBATCS", "POTBATC1", "POTBATC3"}
            With dst.Tables(TABLE_NAME)
                Dim CE As Dictionary(Of String, String) = TC(TABLE_NAME)
                For c As Integer = 0 To .Columns.Count - 1
                    If CE.ContainsKey(.Columns(c).ColumnName) Then
                        .Columns(c).Expression = CE(.Columns(c).ColumnName)
                    End If
                Next
            End With
        Next
        grdPOTBATC2.EndUpdate()
        grdPOTBATC3.EndUpdate()
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("PORWREC2")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_BATCH_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTBATCX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdPOTBATC2, "SSSBSBB", "Show Filter", "Show GroupBox", "Show Pins", "Inventory Status", "Show Details", "Expand All", "Collapse All", "Add by %")
        Load_Popup_Menu(grdPOTBATC3, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "Vendor Inquiry")
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
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Inventory Status"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Show Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splPOTBATCA.Panel2Collapsed = Not tlb_sbt.Checked

            Case "Show Calculations"

            Case "Expand All"
                grdPOTBATC2.Rows.ExpandAll(True)

            Case "Collapse All"
                grdPOTBATC2.Rows.CollapseAll(True)
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Private Sub grdPOTBATCX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTBATCX.AfterRowActivate

    End Sub

    Private Sub grdPOTBATCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTBATCX.DoubleClickRow
        'If grdPOTBATCX.ActiveRow IsNot Nothing Then
        '    Absx1.txtFor("PO_SHIPMENT_NO").Text = grdPOTBATCX.ActiveRow.Cells("PO_SHIPMENT_NO").Text
        '    Click_Command("View")
        'End If
    End Sub

    Sub Setup_SOTSLSC1()
        If grdPOTBATC2.ActiveRow Is Nothing OrElse Not grdPOTBATC2.ActiveRow.IsDataRow Then
            grdSOTSLSC1.Visible = False
        Else
            Dim STYLE_CODE As String = grdPOTBATC2.ActiveRow.Cells("STYLE_CODE").Value
            If grdPOTBATC2.ActiveRow.Band.Key = "POTBATCS_POTBATC2" Then
                Dim COLOR_CODE As String = grdPOTBATC2.ActiveRow.Cells("COLOR_CODE").Value
                Fill_Records("SOTSLSC1", New Object() {STYLE_CODE, COLOR_CODE})
                Sort_grdColumns(grdSOTSLSC1, "WHSE_CODE,CUST_CODE")
                grdSOTSLSC1.Text = "Style " & STYLE_CODE & " Color " & COLOR_CODE & "; Sales Summary"
            Else
                ASCMAIN1.sql = "Select SOTSLSC1.*" & vbCrLf _
                & " from " & SOTSLSC1 & " SOTSLSC1 " & vbCrLf _
                & " where STYLE_CODE = '" & STYLE_CODE & "'"
                Fill_Records("SOTSLSC1", "", True, ASCMAIN1.sql)
                Sort_grdColumns(grdSOTSLSC1, "WHSE_CODE,CUST_CODE")
                grdSOTSLSC1.Text = "Style " & STYLE_CODE & " All Colors; Sales Summary"
            End If
            grdSOTSLSC1.Visible = True
        End If
    End Sub

    Sub Setup_SOTORDC1()
        If grdPOTBATC2.ActiveRow Is Nothing OrElse Not grdPOTBATC2.ActiveRow.IsDataRow Then
            grdSOTORDC1.Visible = False
        Else
            Dim STYLE_CODE As String = grdPOTBATC2.ActiveRow.Cells("STYLE_CODE").Value

            If grdPOTBATC2.ActiveRow.Band.Key = "POTBATCS_POTBATC2" Then
                Dim COLOR_CODE As String = grdPOTBATC2.ActiveRow.Cells("COLOR_CODE").Value
                Fill_Records("SOTORDC1", New Object() {STYLE_CODE, COLOR_CODE})
                Sort_grdColumns(grdSOTORDC1, "WHSE_CODE,CUST_CODE")
                grdSOTORDC1.Text = "Style " & STYLE_CODE & " Color " & COLOR_CODE & "; Open Orders"
            Else
                ASCMAIN1.sql = "Select SOTORDC1.*" & vbCrLf _
                & " from " & SOTORDC1 & " SOTORDC1 " & vbCrLf _
                & " where STYLE_CODE = '" & STYLE_CODE & "'"
                Fill_Records("SOTORDC1", "", True, ASCMAIN1.sql)
                Sort_grdColumns(grdSOTORDC1, "WHSE_CODE,CUST_CODE")
                grdSOTORDC1.Text = "Style " & STYLE_CODE & " All Colors; Sales Summary"
            End If

            grdSOTORDC1.Visible = True
        End If
    End Sub

    Sub Load_POTBATCX()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        Fill_Records("POTBATCX")
        Sort_grdColumns(grdPOTBATCX, "PO_BATCH_NO".ToLower)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdPOTBATCX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTBATCX.InitializeRow
        If e.Row.Cells("BATCH_STATUS").Value & "" <> "O" Then
            e.Row.CellAppearance.BackColor = Drawing.Color.LightGray
        End If
    End Sub

    Private Function grdSOTRSRV2() As Object
        Throw New NotImplementedException
    End Function

    Private Sub grdPOTBATC2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTBATC2.AfterRowActivate
        Setup_SOTSLSC1()
        Setup_SOTORDC1()
    End Sub

    Private Sub grdPOTBATC2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTBATC2.AfterRowUpdate
        e.Row.Cells("PO_QTY_ROUNDED").Value = e.Row.Cells("PO_QTY_CALC").Value
    End Sub

    Private Sub grdPOTBATC2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTBATC2.InitializeRow
        If e.Row.Band.Key = "POTBATCS" Then
            If e.Row.Cells("STYLE_STATUS").Value & "" = "D" Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub cmdFetchSales_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetchSales.Click

    End Sub

    Private Sub optHistory_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optHistory.ValueChanged
        lblFrom.Visible = (optHistory.Value = "D")
        lblTo.Visible = (optHistory.Value = "D")
        dteFrom.Visible = (optHistory.Value = "D")
        dteTo.Visible = (optHistory.Value = "D")
        numLastXXMonths.Visible = (optHistory.Value = "M")
    End Sub

    Sub Create_Temp_Tables()
        ASCMAIN1.sql = "Select POTBATC2.*, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_STATUS" _
             & " from POTBATC2,ICTCOLR1,ICTSTYC1" _
             & " where ICTCOLR1.COLOR_CODE = POTBATC2.COLOR_CODE" _
             & "   and ICTSTYC1.STYLE_CODE = POTBATC2.STYLE_CODE" _
             & "   and ICTSTYC1.COLOR_CODE = POTBATC2.COLOR_CODE" _
             & "   and ROWNUM < 1"
        POTBATC2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTBATC2 & " Add Primary Key (PO_BATCH_NO, STYLE_CODE, COLOR_CODE)")
        ASCDATA1.ExecuteSQL("Create Unique Index I_" & POTBATC2 & "_1 on " & POTBATC2 & " (STYLE_CODE, COLOR_CODE)")

        ASCMAIN1.sql = "Select SOTINVH1.WHSE_CODE, SOTINVH1.CUST_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" _
            & ", SUM(SOTINVH2.ORDR_QTY_SHIP) QTY, SUM(SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) AMT" _
            & ", COUNT(*) CNT, MAX(SOTINVH1.INV_DATE) LAST_INV" _
            & " from SOTINVH1,SOTINVH2" _
            & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
            & "   and ROWNUM < 1" _
            & " group by SOTINVH1.WHSE_CODE, SOTINVH1.CUST_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE"
        SOTSLSC1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSLSC1 & " Add Primary Key (WHSE_CODE, CUST_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTSLSC1 & "_1 on " & SOTSLSC1 & " (STYLE_CODE, COLOR_CODE)")

        sqlSOTORDC1 = "Select SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE" _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
            & ", SOTORDR2.ORDR_QTY_SHIP QTY, SOTORDR2.ORDR_UNIT_PRICE PRICE" _
            & ", SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE AMT" _
            & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" _
            & " from SOTORDR1,SOTORDR2," & POTBATC2 & " POTBATC2" _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
            & "   and SOTORDR2.STYLE_CODE = POTBATC2.STYLE_CODE" _
            & "   and SOTORDR2.COLOR_CODE = POTBATC2.COLOR_CODE"
        SOTORDC1 = ASCMAIN1.Temp_Table(sqlSOTORDC1)
        ' ASCDATA1.ExecuteSQL("Alter Table " & SOTORDC1 & " Add Primary Key (WHSE_CODE, CUST_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTORDC1 & "_1 on " & SOTORDC1 & " (STYLE_CODE, COLOR_CODE)")
    End Sub

    Sub Scan_for_New_Styles()

    End Sub

    Private Sub grdPOTBATCX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTBATCX.InitializeLayout

    End Sub
End Class