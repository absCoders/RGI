Public Class SOFORELO
    Dim WHSE_CODE As String
    Dim CUST_CODE As String
    Dim rowICTWHSE1 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim SOTORDRG As String
    Dim sqlSOTORDRG As String
    Dim sqlSOTORDRS As String
    Dim refresh_SOTORDRG As Boolean = False
    Dim sql_CUST_CODE As String = ""

    Private ORDR_GROUP_NO_M As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            'If ASCMAIN1.DBS_COMPANY = "RGI" Then
            ASCMAIN1.sql = "Select SOTORDRG.*" & vbCrLf _
                   & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_CUST_PO,ARTCUST1.CUST_NAME" & vbCrLf _
                   & ",SOTORDR0.ORDR_CNT,SOTORDR0.ORDR_AMT" & vbCrLf _
                   & ",SOTORDR0.ORDR_CNT_OPEN,SOTORDR0.ORDR_AMT_OPEN" & vbCrLf _
                   & ",SOTORDR0.ORDR_DATE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                   & ",SOTORDR1.ORDR_MESSAGE,SOTORDR1.TERM_CODE" & vbCrLf _
                   & ",SOTORDR5.CUST_CITY,SOTORDR5.CUST_STATE,SOTORDR1.ORDR_PICK_SEQ, SOTORDRM.ORDR_GROUP_NO_M" & vbCrLf _
                   & " from SOTORDRG,SOTORDR0,ARTCUST1,SOTORDR5,SOTORDR1,SOTORDRM" & vbCrLf _
                   & " where SOTORDRG.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                   & "   and ARTCUST1.CUST_CODE (+) = SOTORDR0.CUST_CODE" & vbCrLf _
                   & "   and SOTORDR1.ORDR_NO (+) = SOTORDRG.ORDR_NO_MIN" & vbCrLf _
                   & "   and SOTORDR5.ORDR_NO (+) = SOTORDRG.ORDR_NO_MIN" & vbCrLf _
                   & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'" & vbCrLf _
                   & "   and SOTORDR0.ORDR_CNT_OPEN <> 0" & vbCrLf _
                   & "   and SOTORDR0.ORDR_GROUP_NO = SOTORDRM.ORDR_NO (+)" & vbCrLf _
                   & "   and SOTORDR0.WHSE_CODE = :PARM1" _
                   & "   and SOTORDR1.ECOM_CODE IS NULL"
            'Else
            '    ASCMAIN1.sql = "Select SOTORDRG.*" & vbCrLf _
            '       & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_CUST_PO,ARTCUST1.CUST_NAME" & vbCrLf _
            '       & ",SOTORDR0.ORDR_CNT,SOTORDR0.ORDR_AMT" & vbCrLf _
            '       & ",SOTORDR0.ORDR_CNT_OPEN,SOTORDR0.ORDR_AMT_OPEN" & vbCrLf _
            '       & ",SOTORDR0.ORDR_DATE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
            '       & ",SOTORDR1.ORDR_MESSAGE,SOTORDR1.TERM_CODE" & vbCrLf _
            '       & ",SOTORDR5.CUST_CITY,SOTORDR5.CUST_STATE,SOTORDR1.ORDR_PICK_SEQ" & vbCrLf _
            '       & " from SOTORDRG,SOTORDR0,ARTCUST1,SOTORDR5,SOTORDR1" & vbCrLf _
            '       & " where SOTORDRG.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
            '       & "   and ARTCUST1.CUST_CODE (+) = SOTORDR0.CUST_CODE" & vbCrLf _
            '       & "   and SOTORDR1.ORDR_NO (+) = SOTORDRG.ORDR_NO_MIN" & vbCrLf _
            '       & "   and SOTORDR5.ORDR_NO (+) = SOTORDRG.ORDR_NO_MIN" & vbCrLf _
            '       & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'" & vbCrLf _
            '       & "   and SOTORDR0.ORDR_CNT_OPEN <> 0" & vbCrLf _
            '       & "   and SOTORDR0.WHSE_CODE = :PARM1"
            'End If

            sqlSOTORDRG = ASCMAIN1.sql
            SOTORDRG = ASCMAIN1.Temp_Table(Replace(sqlSOTORDRG, ":PARM1", "''"))
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRG & " Add Primary Key (ORDR_GROUP_NO)")

            ASCMAIN1.sql = "Select * from " & SOTORDRG
            Create_TDA(.Tables.Add("SOTORDRG"), SOTORDRG, "**", 0, True, "", 1)
            .Tables("SOTORDRG").Columns("ORDR_REL_SHORT").DefaultValue = "0"
            .Tables("SOTORDRG").Columns.Add("REL_CXL_NOW")
            .Tables("SOTORDRG").Columns("REL_CXL_NOW").DefaultValue = "0"
            .Tables("SOTORDRG").Columns("ORDR_ALLO_EXCL").DefaultValue = "0"

            sqlSOTORDRS = "Select SOTORDRS.*" & vbCrLf _
                & ", ICTCOLR1.COLOR_DESC" & vbCrLf _
                & " from " & SOTORDRG & " SOTORDRG,SOTORDRS,ICTSTYL1,ICTCOLR1" & vbCrLf _
                & " where SOTORDRS.ORDR_GROUP_NO = SOTORDRG.ORDR_GROUP_NO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTORDRS.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDRS.COLOR_CODE"
            ASCMAIN1.sql = sqlSOTORDRS
            Create_TDA(.Tables.Add, "SOTORDRS", "**", 0, False, "V", 3)
            With .Tables("SOTORDRS").Columns
                .Add("ORDR_QTY_ALLO_NOW", GetType(System.Int32))
                .Add("ORDR_QTY_BACK_NOW", GetType(System.Int32))
                .Add("ORDR_QTY_CANC_NOW", GetType(System.Int32))
            End With

            If ASCMAIN1.DBS_COMPANY = "RGI" Then
                Create_TDA(.Tables.Add, "SOTORDRM", "*")
            End If

            Create_Relation("SOTORDRG", "SOTORDRS", "ORDR_GROUP_NO")
            For Each A As String In New String() {"CUR", "FUT", "CXL"}
                .Tables("SOTORDRG").Columns.Add("ORDR_AMT_ALLO_" & A, GetType(System.Decimal), "SUM(CHILD.ORDR_AMT_ALLO_" & A & ")")
                .Tables("SOTORDRG").Columns.Add("PCT_ALLO_" & A, GetType(System.Decimal), "IIF(ORDR_AMT=0,0,100*ORDR_AMT_ALLO_" & A & "/ORDR_AMT)")
            Next
            .Tables("SOTORDRG").Columns.Add("ORDR_RELEASE_AVAIL", GetType(System.DateTime), "MAX(CHILD.ORDR_RELEASE_AVAIL)")

            ASCMAIN1.sql = "Select X.WHSE_CODE, ICTWHSE1.WHSE_DESC, X.ORDR_CNT, X.ORDR_AMT_OPEN" _
                & " from ICTWHSE1" _
                & ",(Select WHSE_CODE, Count (*) ORDR_CNT, Sum (ORDR_AMT_OPEN) ORDR_AMT_OPEN " _
                & " from SOTORDR0 where ORDR_CNT_OPEN <> 0 group by WHSE_CODE) X" _
                & " where ICTWHSE1.WHSE_CODE = X.WHSE_CODE"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False)
            .Tables("ICTWHSEX").Columns("ORDR_CNT").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "Select T_CODE ORDR_REL_HOLD_CODE, T_DESC ORDR_REL_HOLD_DESC" _
                & " from ASTCODE1 where COLUMN_NAME = 'ORDR_REL_HOLD_CODE'"
            Create_TDA(.Tables.Add, "SOTORELH", "**", 0, False)

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1, False)
            Create_TDA(.Tables.Add, "SOTORDR4", "*", 1, False)

            Create_TDA(.Tables.Add, "SOTCANC0", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)

            .Tables.Add("STATS")
            With .Tables("STATS")
                .Columns.Add("SORT_NO", GetType(System.Int16))
                .Columns.Add("DESC", GetType(System.String))
                .Columns.Add("VALUE", GetType(System.String))
            End With


            Create_TDA(.Tables.Add, "ARTCUST1", "*", , , , , "CUST_ALLO_EXCL")

        End With

        Fill_Records("SOTORELH")
        Sort_grdColumns(grdSOTORELH, "ORDR_REL_HOLD_CODE")

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdSOTORDRG.DataSource = dst.Tables("SOTORDRG")
        grdSOTORDRS.DataSource = dst.Tables("SOTORDRS")
        grdSOTORELH.DataSource = dst.Tables("SOTORELH")
        grdSOTORDR4.DataSource = dst.Tables("SOTORDR4")
        grdStats.DataSource = dst.Tables("STATS")

        grdSOTORDRG.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        'grd.DisplayLayout.Bands(0).Columns("PROM_NON_QUAL").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdSOTORDRG.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "ORDR_REL_SHORT" Or gcol.Key = "ORDR_REL_SHORT_MIN" Or gcol.Key = "REL_CXL_NOW" Or gcol.Key = "ORDR_GROUP_NO_M" Or gcol.Key = "ORDR_ALLO_EXCL" Then
                    If gcol.Key = "ORDR_REL_SHORT" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    ElseIf gcol.Key = "REL_CXL_NOW" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    ElseIf gcol.Key = "REL_CXL_NOW" Then
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Lavender
                    ElseIf gcol.Key = "ORDR_GROUP_NO_M" Then
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.Beige
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    ElseIf gcol.Key = "ORDR_ALLO_EXCL" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    End If
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Yellow
            Next
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            If ASCMAIN1.DBS_COMPANY = "RGI" Then
                .Columns("ORDR_GROUP_NO_M").Header.Fixed = True
                .Columns("ORDR_ALLO_EXCL").Hidden = True
            Else

                For Each C As String In New String() {"ORDR_AMT", "ORDR_REL_SHORT", "ORDR_REL_SHORT_MIN", "ORDR_CNT_OPEN", "ORDR_MESSAGE", "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER", _
                                                      "TERM_CODE", "CUST_CITY", "CUST_STATE", "ORDR_PICK_SEQ", "ORDR_GROUP_NO_M", "REL_CXL_NOW", _
                                                      "ORDR_AMT_ALLO_CUR", "PCT_ALLO_CUR", "ORDR_AMT_ALLO_FUT", "PCT_ALLO_FUT", "ORDR_AMT_ALLO_CXL", "PCT_ALLO_CXL", "ORDR_RELEASE_AVAIL"}
                    .Columns(C).Hidden = True
                Next
            End If

            For Each A As String In New String() {"CUR", "FUT", "CXL"}
                .Columns("PCT_ALLO_" & A).Format = "##0"
                .Columns("PCT_ALLO_" & A).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns("ORDR_AMT_ALLO_" & A).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
        End With

        With grdSOTORDRS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key.StartsWith("ORDR_QTY") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 55
                    gcol.Format = "#,##0"
                ElseIf gcol.Key.StartsWith("ORDR_AMT") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Width = 70
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If

                If gcol.Key = "ORDR_QTY_ALLO_NOW" Or gcol.Key = "ORDR_QTY_BACK_NOW" Or gcol.Key = "ORDR_QTY_CANC_NOW" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEX, New String() {"ORDR_CNT", "ORDR_AMT_OPEN"})

        Create_Summary(grdSOTORDRG, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDRG, New String() {"ORDR_CNT_OPEN", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO_CUR", "ORDR_AMT_ALLO_FUT", "ORDR_AMT_ALLO_CXL"})

        Create_Summary(grdSOTORDRS, "STYLE_CODE", "Count")
        Create_Summary(grdSOTORDRS, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_ALLO", "ORDR_QTY_ALLO_CUR", "ORDR_QTY_ALLO_FUT", "ORDR_QTY_ALLO_CXL", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO_CUR", "ORDR_AMT_ALLO_FUT", "ORDR_AMT_ALLO_CXL"})
        Create_Summary(grdSOTORDRS, New String() {"ORDR_QTY_ALLO_NOW", "ORDR_QTY_BACK_NOW", "ORDR_QTY_CANC_NOW"})

        Bind_Controls(splComments.Panel1, "SOTORDR1")
        Bind_Controls(grpOrder, "SOTORDR1")

        Set_Read_Only(splComments.Panel1, True)
        Set_Read_Only(grpOrder, True)
        grdSOTORDR4.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSOTORDR4.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdSOTORDR4.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

        ASCMAIN1.Add_Value_List(grdSOTORDRS, "WIP_IND", Nothing, New String() {":", "P:PO", "S:Shp"})

        If ASCMAIN1.CLIENT = "VAN" Then
            Show_Filter(grdSOTORDRG, True)
            SplitContainer1.Panel2Collapsed = True
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                End If

                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock("SOTOREL1", Absx1.txtFor("WHSE_CODE").Text) Then Exit Sub
                'End If

            Case "Edit"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Customer"
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Open("SOTOREL1", Absx1.txtFor("WHSE_CODE").Text) Then Exit Sub
                    If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub

                    CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    ASCDATA1.ExecuteSQL("Delete from " & SOTORDRG & " where CUST_CODE = '" & CUST_CODE & "'")
                    ASCDATA1.ExecuteSQL("Insert into " & SOTORDRG & " " & Replace(sqlSOTORDRG, ":PARM1", "'" & WHSE_CODE & "'") & " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'")

                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE) Then Exit Sub

                    ASCMAIN1.Progress("Now Logically Locking Orders")
                    ASCMAIN1.sql = "Select * from " & SOTORDRG & " where CUST_CODE = '" & CUST_CODE & "'"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                        Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                        ASCMAIN1.Progress("-", ORDR_GROUP_NO)
                        If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                    Next
                    ASCMAIN1.Progress("")

                    If ASCMAIN1.CLIENT = "VAN" Then
                        If Not ASCMAIN1.Logical_Lock("ARTCUST1", CUST_CODE) Then Exit Sub
                    End If
                End If
 
            Case "Update"

                For Each rowSOTORDRG As DataRow In dst.Tables("SOTORDRG").Select("REL_CXL_NOW = '1'")
                    Dim ORDR_GROUP_NO As String = rowSOTORDRG.Item("ORDR_GROUP_NO")
                    For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                        Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_OPEN") & "")
                        Dim ORDR_QTY_ALLO_NOW As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_ALLO_NOW") & "")
                        Dim ORDR_QTY_BACK_NOW As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_BACK_NOW") & "")
                        Dim ORDR_QTY_CANC_NOW As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_CANC_NOW") & "")

                        If ORDR_QTY_ALLO_NOW < 0 Or ORDR_QTY_BACK_NOW < 0 Or ORDR_QTY_CANC_NOW < 0 Or ORDR_QTY_OPEN <> (ORDR_QTY_ALLO_NOW + ORDR_QTY_BACK_NOW + ORDR_QTY_CANC_NOW) Then
                            EMsg &= vbCr & "Out of Balance in Order Group " & ORDR_GROUP_NO & " Style " & rowSOTORDRS.Item("STYLE_CODE") & " Color " & rowSOTORDRS.Item("COLOR_CODE")
                        End If
                    Next
                Next

            Case "Cancel"
                If MsgBox("Are you sure that you want to Cancel?", _
                         MsgBoxStyle.YesNo, _
                         "Verification to Cancel Changes") = MsgBoxResult.No Then
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

            Case "View"
                refresh_SOTORDRG = True
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                'Mode_Settings(False)
                ASCMAIN1.MultiTask_Release()
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                'Mode_Settings(False)
                dst.Tables("SOTORDRG").RejectChanges()

                Absx1.txtFor("CUST_CODE").Text = ""
                ASCMAIN1.MultiTask_Release()
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                Mode_Settings(False)
                refresh_SOTORDRG = True
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Edit").Visible = (EntryMode = "V")
                    .Items("Update").Visible = (EntryMode = "E")
                    .Items("Cancel").Visible = (EntryMode = "E")
                    .Items("Done").Visible = (EntryMode = "V")
                    .Items("Refresh").Visible = (EntryMode = "V")
                End With


            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdICTWHSEX.Visible = Not ScreenMode
        UltraExplorerBar1.Groups("Order Release Holds").Visible = ScreenMode
        UltraExplorerBar1.Groups("Statistics").Visible = ScreenMode
        If ASCMAIN1.CLIENT = "VAN" Then
            UltraExplorerBar1.Groups("Statistics").Visible = False
        End If

        lblCUST_CODE.Visible = ScreenMode And (EntryMode = "E")
        txtCUST_CODE.Visible = ScreenMode And (EntryMode = "E")
        txtCUST_NAME.Visible = ScreenMode And (EntryMode = "E")

        If ASCMAIN1.CLIENT = "VAN" And ScreenMode Then
            lblCUST_CODE.Visible = True
            txtCUST_CODE.Visible = True
            txtCUST_NAME.Visible = True
            Set_Read_Only_for_ctl(txtCUST_CODE, (EntryMode = "E"))
        End If

        chkCUST_ALLOW_BACKORDER.Visible = ScreenMode And (EntryMode = "E")
        chkCUST_ALLO_EXCL.Visible = ScreenMode And (EntryMode = "E") And ASCMAIN1.CLIENT = "VAN"

        Dim multiple_order_groups As Boolean = (dst.Tables("SOTORDRG").Select("ORDR_NO_MIN <> ORDR_NO_MAX OR ORDR_CNT > 1").Length > 0)

        If ASCMAIN1.CLIENT = "VAN" Then
            ' grdSOTORDRG.DisplayLayout.Bands(0).Columns("ORDR_ALLO_EXCL").Hidden = Not (EntryMode = "E")
        Else
            grdSOTORDRG.DisplayLayout.Bands(0).Columns("ORDR_REL_SHORT").Hidden = Not (EntryMode = "E")
            grdSOTORDRG.DisplayLayout.Bands(0).Columns("REL_CXL_NOW").Hidden = Not (EntryMode = "E") And Not multiple_order_groups

            grdSOTORDRS.DisplayLayout.Bands(0).Columns("ORDR_QTY_ALLO_NOW").Hidden = Not (EntryMode = "E") And Not multiple_order_groups
            grdSOTORDRS.DisplayLayout.Bands(0).Columns("ORDR_QTY_BACK_NOW").Hidden = Not (EntryMode = "E") And Not multiple_order_groups
            grdSOTORDRS.DisplayLayout.Bands(0).Columns("ORDR_QTY_CANC_NOW").Hidden = Not (EntryMode = "E") And Not multiple_order_groups
        End If


        If EntryMode = "E" Then
            grdSOTORDRG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            If multiple_order_groups Then
                grdSOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Else
                grdSOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            End If
            chkCUST_ALLOW_BACKORDER.Checked = (rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & "" = "1")
        Else
            grdSOTORDRG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdSOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If

        If ScreenMode Then
            If ASCMAIN1.CLIENT = "VAN" Then
                Set_Read_Only_for_ctl(chkCUST_ALLO_EXCL, False)
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDRG", "SOTORDRS", "STATS"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            dst.Tables("SOTORDRM").Rows.Clear()
        End If
        EnforceConstraints(True)

        WHSE_CODE = ""
        txtCUST_CODE.Text = ""
        txtCUST_NAME.Text = ""
        Fill_Records("ICTWHSEX")
        Sort_grdColumns(grdICTWHSEX, "WHSE_CODE")
        ORDR_GROUP_NO_M = String.Empty

    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Set_Expressions(False)

        grdSOTORDRG.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO_M").Hidden = True

        If EntryMode = "V" Then
            WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

            If refresh_SOTORDRG Then
                Refresh_SOTORDRG_All()

                ASCMAIN1.Progress("Now Loading Data ...")
                Fill_Records("SOTORDRG")
                Fill_Records("SOTORDRS")
            End If
            grdSOTORDRG.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO_M").Hidden = Not (ASCMAIN1.DBS_COMPANY = "RGI")

            Dim dvw As DataView = DirectCast(grdSOTORDRG.DataSource, DataTable).DefaultView
            dvw.RowFilter = ""
        Else
            grdSOTORDRG.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO_M").Hidden = ASCMAIN1.DBS_COMPANY <> "RGI"
            If refresh_SOTORDRG = False Then
                If dst.Tables("SOTORDRG").Select("CUST_CODE = '" & CUST_CODE & "' AND ISNULL(ORDR_GROUP_NO_M, '*') <> '*'").Length > 0 Then
                    refresh_SOTORDRG = True
                End If
            End If
            sql_CUST_CODE = "CUST_CODE = '" & CUST_CODE & "'"
            Refresh_SOTORDRG_for_Customer()
            Dim dvw As DataView = DirectCast(grdSOTORDRG.DataSource, DataTable).DefaultView
            dvw.RowFilter = sql_CUST_CODE
        End If

        For Each rowSOTORDRG As DataRow In dst.Tables("SOTORDRG").Select("ISNULL(ORDR_REL_SHORT,'0')<>'1'")
            rowSOTORDRG.Item("ORDR_REL_SHORT") = "0"
            If rowSOTORDRG.Item("ORDR_ALLO_EXCL") & "" = "" Then rowSOTORDRG.Item("ORDR_ALLO_EXCL") = "0"
        Next

        dst.Tables("SOTORDRG").AcceptChanges()

        Set_Expressions(True)

        If EntryMode = "E" Then
            ' Statistics
            dst.Tables("STATS").Rows.Clear()
            Dim Cur As Decimal = Val(dst.Tables("SOTORDRG").Compute("SUM(ORDR_AMT_ALLO_CUR)", "CUST_CODE = '" & CUST_CODE & "'") & String.Empty)
            Dim Fut As Decimal = Val(dst.Tables("SOTORDRG").Compute("SUM(ORDR_AMT_ALLO_FUT)", "CUST_CODE = '" & CUST_CODE & "'") & String.Empty)
            Dim Cxl As Decimal = Val(dst.Tables("SOTORDRG").Compute("SUM(ORDR_AMT_ALLO_CXL)", "CUST_CODE = '" & CUST_CODE & "'") & String.Empty)
            Dim ORDR_AMT As Decimal = Val(dst.Tables("SOTORDRG").Compute("SUM(ORDR_AMT)", "CUST_CODE = '" & CUST_CODE & "'") & String.Empty)

            If ORDR_AMT = 0 Then
                Cur = 0
                Fut = 0
                Cxl = 0
            Else
                Cur = (Cur * 100) / ORDR_AMT
                Fut = (Fut * 100) / ORDR_AMT
                Cxl = (Cxl * 100) / ORDR_AMT
            End If

            dst.Tables("STATS").Rows.Add(New Object() {1, "% Cur", Cur.ToString("#,##0.00")})
            dst.Tables("STATS").Rows.Add(New Object() {2, "% Fut", Fut.ToString("#,##0.00")})
            dst.Tables("STATS").Rows.Add(New Object() {3, "% Cxl", Cxl.ToString("#,##0.00")})
        End If

        EnforceConstraints(True)

        Sort_grdColumns(grdSOTORDRG, "CUST_CODE,ORDR_GROUP_NO")

        If EntryMode = "E" Then
            If ASCMAIN1.CLIENT = "VAN" Then
                Dim rowARTCUST1 As DataRow = Fill_Record("ARTCUST1", CUST_CODE)
                chkCUST_ALLO_EXCL.Checked = (rowARTCUST1.Item("CUST_ALLO_EXCL") & "" = "1")
            End If
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        If ASCMAIN1.CLIENT = "VAN" Then

            BeginTrans()

            Update_Record_TDA("SOTORDRG")
            Dim UID As String = "'" & ASCMAIN1.USER_ID & "'"
            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is Select * from " & SOTORDRG & " where CUST_CODE = '" & CUST_CODE & "';" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update SOTORDRG Set ORDR_ALLO_EXCL = R1.ORDR_ALLO_EXCL" & vbCrLf _
                & " , LAST_DATE = SYSDATE, LAST_OPER = " & UID & vbCrLf _
                & "    where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                & "   If SQL%NOTFOUND and R1.ORDR_ALLO_EXCL = '1' Then " & vbCrLf _
                & "    Insert into SOTORDRG (ORDR_GROUP_NO,ORDR_ALLO_EXCL,INIT_DATE,INIT_OPER,LAST_DATE,LAST_OPER)" & vbCrLf _
                & "     Values (R1.ORDR_GROUP_NO,R1.ORDR_ALLO_EXCL,SYSDATE," & UID & ",SYSDATE," & UID & ");" & vbCrLf _
                & "   End If;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()



            Dim CUST_ALLO_EXCL As String = IIf(chkCUST_ALLO_EXCL.Checked, "1", "0")
            Dim rowARTCUST1 As DataRow = Fill_Record("ARTCUST1", CUST_CODE)
            If rowARTCUST1.Item("CUST_ALLO_EXCL") & "" = "" Then rowARTCUST1.Item("CUST_ALLO_EXCL") = "0"
            If rowARTCUST1.Item("CUST_ALLO_EXCL") & "" <> CUST_ALLO_EXCL Then

                'ASCMAIN1.sql = "Update ARTCUST1 Set CUST_ALLO_EXCL = :PARM1 where CUST_CODE = :PARM2"
                'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {CUST_ALLO_EXCL, CUST_CODE})

                rowARTCUST1.Item("CUST_ALLO_EXCL") = CUST_ALLO_EXCL
                Write_Audit_Trail(rowARTCUST1)
                Update_Record_TDA("ARTCUST1")
            End If

            CommitTrans("Update Complete")

        Else

            BeginTrans()

            Update_Record_TDA("SOTORDRG")

            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is Select * from " & SOTORDRG & " where CUST_CODE = '" & CUST_CODE & "';" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "  Update SOTORDRG Set ORDR_REL_SHORT = R1.ORDR_REL_SHORT, ORDR_REL_SHORT_MIN = R1.ORDR_REL_SHORT_MIN, LAST_DATE = R1.LAST_DATE, LAST_OPER = R1.LAST_OPER" & vbCrLf _
                & "   where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            Dim ORDR_NOs_to_release As New List(Of String)
            Dim ORDR_GROUP_NOs_to_release As New List(Of String)

            For Each rowSOTORDRG As DataRow In dst.Tables("SOTORDRG").Select("REL_CXL_NOW = '1'")
                Dim ORDR_GROUP_NO As String = rowSOTORDRG.Item("ORDR_GROUP_NO")
                Dim ORDR_NO_MIN As String = rowSOTORDRG.Item("ORDR_NO_MIN")
                Dim ORDR_NO_MAX As String = rowSOTORDRG.Item("ORDR_NO_MAX")
                Dim ORDR_CNT As String = rowSOTORDRG.Item("ORDR_CNT")
                Dim ORDR_NO As String = ORDR_NO_MIN
                If ORDR_NO_MAX <> ORDR_NO_MIN Or ORDR_CNT <> 1 Then
                    MsgBox("Problem with Order Group " & ORDR_GROUP_NO & " - multiple orders found", MsgBoxStyle.OkOnly, "Manual Release not permitted")
                    Rollback()
                    Exit Sub
                End If

                Fill_Records("SOTORDR2", ORDR_NO)

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                    rowSOTORDR2.Item("ORDR_QTY_PRE_ALLO") = DBNull.Value
                Next

                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                    'Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_OPEN") & "")
                    Dim ORDR_QTY_ALLO_NOW As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_ALLO_NOW") & "")
                    'Dim ORDR_QTY_BACK_NOW As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_BACK_NOW") & "")
                    Dim ORDR_QTY_CANC_NOW As Int64 = Val(rowSOTORDRS.Item("ORDR_QTY_CANC_NOW") & "")
                    If ORDR_QTY_ALLO_NOW <> 0 Then
                        If Not ORDR_NOs_to_release.Contains(ORDR_NO_MAX) Then ORDR_NOs_to_release.Add(ORDR_NO_MAX)
                        If Not ORDR_GROUP_NOs_to_release.Contains(ORDR_GROUP_NO) Then ORDR_GROUP_NOs_to_release.Add(ORDR_GROUP_NO)

                        Dim STYLE_CODE As String = rowSOTORDRS.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTORDRS.Item("COLOR_CODE")
                        Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
                            Dim ORDR_QTY_OPEN As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                            Dim ORDR_QTY_ALLO_NOW_LNO As Int32 = 0
                            If ORDR_QTY_OPEN >= ORDR_QTY_CANC_NOW Then
                                ORDR_QTY_ALLO_NOW_LNO = ORDR_QTY_ALLO_NOW
                            Else
                                ORDR_QTY_ALLO_NOW_LNO = ORDR_QTY_OPEN
                            End If
                            rowSOTORDR2.Item("ORDR_QTY_PRE_ALLO") = ORDR_QTY_ALLO_NOW_LNO
                            ORDR_QTY_ALLO_NOW -= ORDR_QTY_ALLO_NOW_LNO
                            If ORDR_QTY_ALLO_NOW = 0 Then Exit For
                        Next
                    End If

                    If ORDR_QTY_CANC_NOW <> 0 Then
                        Dim STYLE_CODE As String = rowSOTORDRS.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTORDRS.Item("COLOR_CODE")
                        Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
                            Dim ORDR_QTY_OPEN As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                            Dim ORDR_QTY_PICK As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                            Dim ORDR_QTY_SHIP As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & "")
                            Dim ORDR_QTY_CANC As Int32 = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")

                            Dim ORDR_QTY_CANC_NOW_LNO As Int32 = 0
                            If ORDR_QTY_OPEN >= ORDR_QTY_CANC_NOW Then
                                ORDR_QTY_CANC_NOW_LNO = ORDR_QTY_CANC_NOW
                            Else
                                ORDR_QTY_CANC_NOW_LNO = ORDR_QTY_OPEN
                            End If

                            ORDR_QTY_CANC += ORDR_QTY_CANC_NOW_LNO
                            rowSOTORDR2.Item("ORDR_QTY_CANC") = ORDR_QTY_CANC
                            ORDR_QTY_OPEN -= ORDR_QTY_CANC_NOW_LNO
                            rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN

                            ORDR_QTY_CANC_NOW -= ORDR_QTY_CANC_NOW_LNO

                            Dim ORDR_STATUS_LINE As String = IIf(ORDR_QTY_OPEN <> 0, "O", IIf(ORDR_QTY_PICK <> 0, "P", IIf(ORDR_QTY_SHIP <> 0, "F", "C")))
                            rowSOTORDR2.Item("ORDR_STATUS") = ORDR_STATUS_LINE

                            Dim rowSOTCANC0 As DataRow = dst.Tables("SOTCANC0").NewRow
                            rowSOTCANC0.Item("ORDR_NO") = ORDR_GROUP_NO
                            rowSOTCANC0.Item("ORDR_LNO") = rowSOTORDR2.Item("ORDR_LNO")
                            rowSOTCANC0.Item("STYLE_CODE") = STYLE_CODE
                            rowSOTCANC0.Item("COLOR_CODE") = COLOR_CODE
                            rowSOTCANC0.Item("ORDR_QTY_CANC_NOW") = ORDR_QTY_CANC_NOW_LNO
                            dst.Tables("SOTCANC0").Rows.Add(rowSOTCANC0)

                            TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", -1 * ORDR_QTY_CANC_NOW_LNO)

                            If ORDR_QTY_CANC_NOW = 0 Then Exit For
                        Next
                    End If

                Next

                Update_Record_TDA("SOTORDR2")

                Dim QTY_S As Int32 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_SHIP)", "") & "")
                Dim QTY_P As Int32 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_PICK)", "") & "")
                Dim QTY_O As Int32 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", "") & "")
                Dim ORDR_STATUS As String = IIf(QTY_O <> 0, "O", IIf(QTY_P <> 0, "P", IIf(QTY_S <> 0, "F", "C")))
                ASCDATA1.ExecuteSQL("Update SOTORDR1 Set ORDR_STATUS = :PARM1 where ORDR_NO = :PARM2", "VV", New String() {ORDR_STATUS, ORDR_NO})
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
            Next
            Update_Record_TDA("SOTCANC0")


            CommitTrans("Update Complete")




            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

            If ORDR_NOs_to_release.Count <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Releasing Orders")

                Dim REPORT_NAME As String = "SOROREL1"

                ' would like to eliminate the next 3 lines but troubles inside SOROREL1 starting with ICTSTYL1 datatable creation
                If REPORTS.ContainsKey(REPORT_NAME) Then
                    REPORTS.Remove(REPORT_NAME)
                End If

                If Not REPORTS.ContainsKey(REPORT_NAME) Then
                    REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
                    REPORTS(REPORT_NAME).Prepare_dst(False, "")
                End If

                REPORTS(REPORT_NAME).Fill_Records_RPT(New Object() {ORDR_GROUP_NOs_to_release})
                With REPORTS(REPORT_NAME).clsASCBASE1

                    ' Dim rowSOTPICK0 As DataRow = .dst.Tables("SOTPICK0").Rows(0)
                    Dim PICK_BATCH_NO As String = REPORTS(REPORT_NAME).XNO ' rowSOTPICK0.Item(PICK_BATCH_NO)

                    .Print_Report_Begin()
                    Dim SUBT As String = ""
                    SUBT = "Batch " & PICK_BATCH_NO & " (Manual Release)"
                    Dim RPT As String = REPORT_NAME
                    Dim RPT_TITLE As String = "Released Orders Report"
                    .CR_params.Add("SUBT", SUBT)
                    .Generate_Report(RPT, RPT_TITLE, SUBT, True, , , , , False)
                    ' .Generate_Report(RPT, RPT_TITLE, SUBT, "{SOTORDR1.WHSE_CODE}='" & WHSE_CODE & "'")

                    RPT = "SORORELA"
                    RPT_TITLE = "Un-Releasable Orders Report"
                    SUBT = "Manual Release"
                    .CR_params.Add("SORT_SREP", "0")
                    .CR_params.Add("RELEASE_DATE", Format$(Now.Date, "yyyyMMdd"))
                    .CR_params.Add("CHKALLOCATION_ONLY", "0")
                    .Generate_Report(RPT, RPT_TITLE, SUBT)

                    .Print_Report_End()
                End With

                For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs_to_release
                    Rebuild_SOTORDRG_for_ORDR_GROUP_NO(ORDR_GROUP_NO)
                Next

                Rebuild_SOTORDRG_for_CUST_CODE()

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")


                Dim rows() As DataRow = dst.Tables("SOTORDRG").Select("REL_CXL_NOW = '1'")
                For i As Integer = rows.Length - 1 To 0 Step -1
                    rows(i).Item("REL_CXL_NOW") = "0"
                Next
                'For Each rowSOTORDRG As DataRow In dst.Tables("SOTORDRG").Select("REL_CXL_NOW = '1'")
                '    rowSOTORDRG.Item("REL_CXL_NOW") = "0"
                'Next
            End If


        End If
    End Sub

#End Region

    Sub Rebuild_SOTORDRG_for_ORDR_GROUP_NO(ORDR_GROUP_NO As String)
        ASCMAIN1.sql = "Begin" & vbCrLf _
            & "DECLARE CURSOR C1 IS " & vbCrLf _
            & "SELECT SOTORDR1.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_UNIT_PRICE * SOTORDR2.ORDR_QTY) ORDR_AMT" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_UNIT_PRICE * SOTORDR2.ORDR_QTY_OPEN) ORDR_AMT_OPEN" & vbCrLf _
            & "FROM SOTORDR1,SOTORDR2 WHERE SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "AND SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & "GROUP BY SOTORDR1.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE SOTORDRS SET " & vbCrLf _
            & "  ORDR_QTY = R1.ORDR_QTY" & vbCrLf _
            & ", ORDR_QTY_OPEN = R1.ORDR_QTY_OPEN" & vbCrLf _
            & ", ORDR_QTY_ALLO = R1.ORDR_QTY_ALLO" & vbCrLf _
            & ", ORDR_QTY_PICK = R1.ORDR_QTY_PICK" & vbCrLf _
            & ", ORDR_QTY_SHIP = R1.ORDR_QTY_SHIP" & vbCrLf _
            & ", ORDR_QTY_CANC = R1.ORDR_QTY_CANC" & vbCrLf _
            & ", ORDR_AMT = R1.ORDR_AMT" & vbCrLf _
            & ", ORDR_AMT_OPEN = R1.ORDR_AMT_OPEN" & vbCrLf _
            & "WHERE ORDR_GROUP_NO = R1.ORDR_GROUP_NO" & vbCrLf _
            & "  AND STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
            & "  AND COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        ' THIS IS WIPING OUT THE HARD WORK OF AN ALLOCATION ONLY RELEASE

        '& ", ORDR_QTY_ALLO_CUR = 0" & vbCrLf _
        '& ", ORDR_QTY_ALLO_FUT = 0" & vbCrLf _
        '& ", ORDR_QTY_ALLO_CXL = 0" & vbCrLf _
        '& ", ORDR_AMT_ALLO_CUR = 0" & vbCrLf _
        '& ", ORDR_AMT_ALLO_FUT = 0" & vbCrLf _
        '& ", ORDR_AMT_ALLO_CXL = 0" & vbCrLf _
    End Sub

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDRG, "SSSPBBBPBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Order Inquiry", "Sales Order Entry", "Release Selected", "Group Selected", "Un-Group Selected", "Exclude Selected", "Un-Exclude Selected")
        Load_Popup_Menu(grdSOTORDRS, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTORDRG"
                    If ASCMAIN1.CLIENT = "VAN" Then
                        tlb_btn = DirectCast(tlb_pop.Tools("Group Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = False

                        tlb_btn = DirectCast(tlb_pop.Tools("Un-Group Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = False

                        tlb_btn = DirectCast(tlb_pop.Tools("Release Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = False

                        tlb_btn = DirectCast(tlb_pop.Tools("Exclude Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Enabled = EntryMode = "E"
                        tlb_btn = DirectCast(tlb_pop.Tools("Un-Exclude Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Enabled = EntryMode = "E"

                    Else
                        tlb_btn = DirectCast(tlb_pop.Tools("Group Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Enabled = EntryMode = "E"

                        'tlb_btn = DirectCast(tlb_pop.Tools("Un-Group Selected"), UltraWinToolbars.ButtonTool)
                        'tlb_btn.SharedProps.Enabled = EntryMode = "E"

                        tlb_btn = DirectCast(tlb_pop.Tools("Release Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = EntryMode = "E"

                        tlb_btn = DirectCast(tlb_pop.Tools("Exclude Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = False
                        tlb_btn = DirectCast(tlb_pop.Tools("Un-Exclude Selected"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = False
 
                    End If


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Customer Order Inquiry"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Context_Launch("Select", CUST_CODE & ":" & ORDR_GROUP_NO, e.Tool.Key, "SOFCORD1")


            Case "Sales Order Entry"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Dim ORDR_NO_MAX As String = grd.ActiveRow.Cells("ORDR_NO_MAX").Value & ""
                'ASCMAIN1.sql = "Select MAX(ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                'Dim ORDR_NO As String = ASCDATA1.GetDataValue
                Context_Launch("View", ORDR_NO_MAX, e.Tool.Key, "SOFORDR1")

            Case "Release Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("ORDR_REL_SHORT").Value = "1"
                    grow.Update()
                Next

            Case "Group Selected", "Un-Group Selected"
                refresh_SOTORDRG = True
                If grdSOTORDRG.Selected.Rows.Count = 0 Then
                    MessageBox.Show("You must select at least one Sales Order.", "Groupings", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                ElseIf grdSOTORDRG.Selected.Rows.Count <= 1 AndAlso e.Tool.Key = "Group Selected" Then
                    MessageBox.Show("You must select more than one Sales Order.", "Groupings", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim orderNos As New List(Of String)
                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTORDRG.Selected.Rows
                    orderNos.Add(grdRow.Cells("ORDR_GROUP_NO").Value)
                Next

                Dim sql As String = "Delete From SOTORDRM WHERE ORDR_NO IN ('" & Join(orderNos.ToArray, "', '") & "')"
                ASCDATA1.ExecuteSQL(sql)

                ' Clean up single stragglers
                sql = "Delete From SOTORDRM where ORDR_GROUP_NO_M in (Select ORDR_GROUP_NO_M FROM SOTORDRM GROUP BY ORDR_GROUP_NO_M having Count(*) = 1)"
                ASCDATA1.ExecuteSQL(sql)

                For Each ORDR_NO As String In orderNos
                    If dst.Tables("SOTORDRG").Select("ORDR_GROUP_NO = '" & ORDR_NO & "'").Length > 0 Then
                        dst.Tables("SOTORDRG").Select("ORDR_GROUP_NO = '" & ORDR_NO & "'")(0).Item("ORDR_GROUP_NO_M") = String.Empty
                    End If
                    ASCDATA1.ExecuteSQL("UPDATE " & SOTORDRG & " SET ORDR_GROUP_NO_M = NULL WHERE ORDR_GROUP_NO = '" & ORDR_NO & "'")
                Next

                If e.Tool.Key = "Un-Group Selected" Then
                    MessageBox.Show("Un-Grouping Successful.", "Grouping", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim ORDR_GROUP_NO_M As String = ASCMAIN1.Next_Control_No("ORDR_GROUP_NO_M")
                sql = "Insert Into SOTORDRM Select '" & ORDR_GROUP_NO_M & "', ORDR_NO, CUST_CODE FROM SOTORDR1 WHERE ORDR_NO IN ('" & Join(orderNos.ToArray, "', '") & "')"
                ASCDATA1.ExecuteSQL(sql)

                For Each ORDR_NO As String In orderNos
                    If dst.Tables("SOTORDRG").Select("ORDR_GROUP_NO = '" & ORDR_NO & "'").Length > 0 Then
                        dst.Tables("SOTORDRG").Select("ORDR_GROUP_NO = '" & ORDR_NO & "'")(0).Item("ORDR_GROUP_NO_M") = ORDR_GROUP_NO_M
                    End If
                    ASCDATA1.ExecuteSQL("UPDATE " & SOTORDRG & " SET ORDR_GROUP_NO_M = '" & ORDR_GROUP_NO_M & "' WHERE ORDR_GROUP_NO = '" & ORDR_NO & "'")
                Next

                MessageBox.Show("Grouping Successful.", "Grouping", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Case "Exclude Selected", "Un-Exclude Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("ORDR_ALLO_EXCL").Value = IIf(e.Tool.Key = "Exclude Selected", "1", "0")
                    grow.Update()
                Next
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And ScreenMode And (EntryMode = "V") Then
                    Click_Command("Edit", e)
                End If

            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("Edit")
            Case "WHSE_CODE"
                Click_Command("Load")
        End Select
    End Sub
#End Region

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("View")
    End Sub

    Private Sub grdSOTORDRG_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDRG.AfterRowActivate
        Setup_grdSOTORDRS()
    End Sub

    Sub Setup_grdSOTORDRS()
        If grdSOTORDRG.ActiveRow Is Nothing OrElse Not grdSOTORDRG.ActiveRow.IsDataRow Then
            grdSOTORDRS.Visible = False
        Else
            Dim ORDR_GROUP_NO As String = grdSOTORDRG.ActiveRow.Cells("ORDR_GROUP_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTORDRS.DataSource, DataTable).DefaultView
            dvw.RowFilter = "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            Sort_grdColumns(grdSOTORDRS, "STYLE_CODE,COLOR_CODE")
            grdSOTORDRS.Text = "Order Details for Order Group " & ORDR_GROUP_NO
            grdSOTORDRS.Visible = True

            Dim ORDR_NO As String = grdSOTORDRG.ActiveRow.Cells("ORDR_NO_MIN").Value & ""
            If ORDR_NO = "" Then
                ASCMAIN1.sql = "Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                ORDR_NO = ASCDATA1.GetDataValue
            End If
            Fill_Records("SOTORDR4", ORDR_NO)
            Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO)
            lblOrder.Text = "Order No " & ORDR_NO
            lblCustomer.Text = rowSOTORDR1.Item("CUST_CODE") & vbCrLf & rowSOTORDR1.Item("CUST_NAME")

            Select Case rowSOTORDR1.Item("ORDR_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "P"
                    lblStatus.Text = "In Pick"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
                Case "F"
                    lblStatus.Text = "Shipped"
            End Select
            lblStatus.Visible = True
        End If
    End Sub

    Private Sub grdSOTORDRG_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDRG.AfterRowUpdate

        If e.Row.Band.Key = "SOTORDRG" Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Setting Details")

            For Each grow As UltraWinGrid.UltraGridRow In e.Row.ChildBands(0).Rows
                If e.Row.Cells("REL_CXL_NOW").Value & "" = "1" And Val(grow.Cells("ORDR_QTY_OPEN").Value & "") <> 0 Then
                    grow.Cells("ORDR_QTY_ALLO_NOW").Value = grow.Cells("ORDR_QTY_ALLO_CUR").Value ' grow.Cells("ORDR_QTY_OPEN").Value
                    Dim ORDR_QTY_BACK_NOW As Int32 = Val(grow.Cells("ORDR_QTY_OPEN").Value & "") - Val(grow.Cells("ORDR_QTY_ALLO_CUR").Value & "")
                    If ORDR_QTY_BACK_NOW < 0 Then ORDR_QTY_BACK_NOW = 0
                    grow.Cells("ORDR_QTY_BACK_NOW").Value = ORDR_QTY_BACK_NOW
                Else
                    grow.Cells("ORDR_QTY_ALLO_NOW").Value = DBNull.Value
                    grow.Cells("ORDR_QTY_BACK_NOW").Value = DBNull.Value
                    grow.Cells("ORDR_QTY_CANC_NOW").Value = DBNull.Value
                End If
            Next

            grdSOTORDRS.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
            'Dim ORDR_GROUP_NO As String = e.Row.Cells("ORDR_GROUP_NO").Value
            'For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_QTY_OPEN <> 0")
            '    If e.Row.Cells("REL_CXL_NOW").Value & "" = "1" Then
            '        rowSOTORDRS.Item("ORDR_QTY_ALLO_NOW") = rowSOTORDRS.Item("ORDR_QTY_OPEN")
            '        rowSOTORDRS.Item("ORDR_QTY_BACK_NOW") = 0
            '        rowSOTORDRS.Item("ORDR_QTY_CANC_NOW") = 0
            '    Else
            '        rowSOTORDRS.Item("ORDR_QTY_ALLO_NOW") = DBNull.Value
            '        rowSOTORDRS.Item("ORDR_QTY_BACK_NOW") = DBNull.Value
            '        rowSOTORDRS.Item("ORDR_QTY_CANC_NOW") = DBNull.Value
            '    End If
            'Next
            Toggle_grdSOTORDRS(e.Row.Cells("REL_CXL_NOW").Value & "" = "1")
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
       
    End Sub

    Sub Toggle_grdSOTORDRS(allow_Update As Boolean)
        If allow_Update Then
            grdSOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Else
            grdSOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If
    End Sub

    Private Sub grdSOTORDRG_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDRG.BeforeRowUpdate
        If e.Row.Band.Key = "SOTORDRG" Then
            If e.Row.Cells("ORDR_REL_SHORT").Value & "" <> "1" Then ' THIS APPEARS REVERSED, BUT IT WORKS
                e.Row.Cells("ORDR_REL_SHORT_MIN").Value = e.Row.Cells("ORDR_AMT_ALLO_CUR").Value
            Else
                e.Row.Cells("ORDR_REL_SHORT_MIN").Value = DBNull.Value
            End If
        End If
    End Sub

    Private Sub grdSOTORDRG_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDRG.DoubleClickRow

        ORDR_GROUP_NO_M = String.Empty
        If EntryMode = "V" Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Text
            ORDR_GROUP_NO_M = e.Row.Cells("ORDR_GROUP_NO_M").Text
            Click_Command("Edit")
        End If

    End Sub

    Private Sub grdSOTORDRG_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRG.InitializeRow

        If e.Row.Band.Key = "SOTORDRG" Then
            With e.Row.Cells("ORDR_CANCEL_DATE")
                If Format(CDate(.Value & ""), "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                    .Appearance.ForeColor = Drawing.Color.Red
                    .ToolTipText = "Order is Past Cancel Date"
                Else
                End If
            End With
        End If

    End Sub

    Sub Refresh_SOTORDRG_All()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Orders Queue")

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Select ORDR_GROUP_NO from SOTORDR0 where ORDR_CNT_OPEN <> 0 minus Select ORDR_GROUP_NO from SOTORDRG"
            ASCMAIN1.sql = "Insert into SOTORDRG (ORDR_GROUP_NO) " & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & SOTORDRG)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDRG & " " & Replace(sqlSOTORDRG, ":PARM1", "'" & WHSE_CODE & "'"))
        refresh_SOTORDRG = False
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Refresh_SOTORDRG_for_Customer()
        
        Rebuild_SOTORDRG_for_CUST_CODE()

        ASCDATA1.DeleteRows(dst.Tables("SOTORDRG"), sql_CUST_CODE)
        ASCMAIN1.sql = "Select * from " & SOTORDRG & " where " & sql_CUST_CODE

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            If ORDR_GROUP_NO_M.Length > 0 Then
                ASCMAIN1.sql &= " AND ORDR_GROUP_NO_M = '" & ORDR_GROUP_NO_M & "'"
            Else
                ASCMAIN1.sql &= " AND ORDR_GROUP_NO_M IS NULL"
            End If
        End If

        Fill_Records("SOTORDRG", "", False, ASCMAIN1.sql)

        For Each rowSOTORDRG As DataRow In dst.Tables("SOTORDRG").Select(sql_CUST_CODE)
            Dim ORDR_GROUP_NO As String = rowSOTORDRG.Item("ORDR_GROUP_NO")
            Rebuild_SOTORDRG_for_ORDR_GROUP_NO(ORDR_GROUP_NO)
        Next

        ASCMAIN1.sql = sqlSOTORDRS & " and SOTORDRG." & sql_CUST_CODE
        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            If ORDR_GROUP_NO_M.Length > 0 Then
                ASCMAIN1.sql &= " AND SOTORDRG.ORDR_GROUP_NO_M = '" & ORDR_GROUP_NO_M & "'"
            Else
                ASCMAIN1.sql &= " AND SOTORDRG.ORDR_GROUP_NO_M IS NULL"
            End If
        End If
        Fill_Records("SOTORDRS", "", False, ASCMAIN1.sql)
    End Sub

    Sub Rebuild_SOTORDRG_for_CUST_CODE()

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Select ORDR_GROUP_NO from SOTORDR0 where ORDR_CNT_OPEN <> 0 minus Select ORDR_GROUP_NO from SOTORDRG"
            ASCMAIN1.sql = "Insert into SOTORDRG (ORDR_GROUP_NO) " & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()
        End If

        ASCDATA1.ExecuteSQL("Delete from " & SOTORDRG & " where CUST_CODE = '" & CUST_CODE & "'")
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDRG & " " & Replace(sqlSOTORDRG, ":PARM1", "'" & WHSE_CODE & "'") & " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'")
    End Sub

    Sub Set_Expressions(tf As Boolean)
        For Each A As String In New String() {"CUR", "FUT", "CXL"}
            dst.Tables("SOTORDRG").Columns("ORDR_AMT_ALLO_" & A).Expression = IIf(tf, "SUM(CHILD.ORDR_AMT_ALLO_" & A & ")", "")
            dst.Tables("SOTORDRG").Columns("PCT_ALLO_" & A).Expression = IIf(tf, "IIF(ORDR_AMT=0,0,100*ORDR_AMT_ALLO_" & A & "/ORDR_AMT)", "")
        Next
        dst.Tables("SOTORDRG").Columns("ORDR_RELEASE_AVAIL").Expression = IIf(tf, "MAX(CHILD.ORDR_RELEASE_AVAIL)", "")

    End Sub

    Private Sub grdSOTORDRS_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDRS.BeforeRowUpdate

        Dim ORDR_QTY_ALLO_NOW As Int64 = Val(e.Row.Cells("ORDR_QTY_ALLO_NOW").Value & "")
        Dim ORDR_QTY_BACK_NOW As Int64 = Val(e.Row.Cells("ORDR_QTY_BACK_NOW").Value & "")
        Dim ORDR_QTY_CANC_NOW As Int64 = Val(e.Row.Cells("ORDR_QTY_CANC_NOW").Value & "")
        Dim ORDR_QTY_OPEN As Int64 = Val(e.Row.Cells("ORDR_QTY_OPEN").Value & "")

        If grdSOTORDRG.ActiveRow.Cells("REL_CXL_NOW").Value & "" = "1" And ORDR_QTY_OPEN <> 0 Then

            If ORDR_QTY_ALLO_NOW > ORDR_QTY_OPEN Then
                e.Cancel = True
            Else
                If ORDR_QTY_OPEN <> (ORDR_QTY_ALLO_NOW + ORDR_QTY_BACK_NOW + ORDR_QTY_CANC_NOW) Then
                    If chkCUST_ALLOW_BACKORDER.Checked Then
                        e.Row.Cells("ORDR_QTY_BACK_NOW").Value = ORDR_QTY_OPEN - ORDR_QTY_ALLO_NOW
                        e.Row.Cells("ORDR_QTY_CANC_NOW").Value = 0
                    Else
                        e.Row.Cells("ORDR_QTY_CANC_NOW").Value = ORDR_QTY_OPEN - ORDR_QTY_ALLO_NOW
                        e.Row.Cells("ORDR_QTY_BACK_NOW").Value = 0
                    End If
                End If
            End If
        Else
            e.Row.Cells("ORDR_QTY_ALLO_NOW").Value = DBNull.Value
            e.Row.Cells("ORDR_QTY_BACK_NOW").Value = DBNull.Value
            e.Row.Cells("ORDR_QTY_CANC_NOW").Value = DBNull.Value
        End If
        
    End Sub
End Class