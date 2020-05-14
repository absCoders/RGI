Public Class WHFIADJ1
    Dim relation As Data.DataRelation
    Dim filterGrdWHTIADJ1 As String
    Dim filterGrdWHTIADJS As String
    Dim filterGrdWHTIADJW As String
    Dim filterGrdWHTIADJD As String

    Dim rowWHTTPLP1 As DataRow

#Region "ABS Standard Routines"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "SELECT WHTIADJ1.* " & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'S', WHTSTYLX.STYLE_CODE, NULL) STYLE_CODE " & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'S', WHTSTYLX.COLOR_CODE, NULL) COLOR_CODE " & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'S', WHTSTYLX.STYLE_DESC, NULL) STYLE_DESC " & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'S', WHTSTYLX.COLOR_DESC, NULL) COLOR_DESC " & vbCrLf _
                & " , ICTSTYC1.STYLE_COST_FIFO " & vbCrLf _
                & " FROM WHTIADJ1, WHTSTYLX, ICTSTYC1 " & vbCrLf _
                & " WHERE WHTIADJ1.ABS_STATUS     <> 'A'" & vbCrLf _
                & "   AND WHTSTYLX.ITEM_CODE      = WHTIADJ1.ITEM_CODE" & vbCrLf _
                & "   AND ICTSTYC1.STYLE_CODE (+) = WHTSTYLX.STYLE_CODE" & vbCrLf _
                & "   AND ICTSTYC1.COLOR_CODE (+) = WHTSTYLX.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTIADJ1", "**", 0, True, "", 3)
            .Tables("WHTIADJ1").Columns.Add("SEL")
            .Tables("WHTIADJ1").Columns.Add("DEL")
            .Tables("WHTIADJ1").Columns.Add("TOTAL_ADJ_AMT")

            .Tables("WHTIADJ1").Columns("SEL").DefaultValue = "0"
            .Tables("WHTIADJ1").Columns("DEL").DefaultValue = "0"

            'Get exploded STYLE_CODE and COLOR_CODE if ppk
            ASCMAIN1.sql = "Select " & vbCrLf _
                & "   WHTIADJ1.ITEM_CODE        ITEM_CODE" & vbCrLf _
                & " , WHTIADJ1.TRANS_SEQ        TRANS_SEQ" & vbCrLf _
                & " , WHTIADJ1.LP_CODE          LP_CODE" & vbCrLf _
                & " , WHTIADJ1.WHSE_CODE        WHSE_CODE" & vbCrLf _
                & " , WHTPPKM2.STYLE_CODE       STYLE_CODE" & vbCrLf _
                & " , WHTPPKM2.COLOR_CODE       COLOR_CODE" & vbCrLf _
                & " , ICTSTYL1.STYLE_DESC       STYLE_DESC" & vbCrLf _
                & " , ICTCOLR1.COLOR_DESC       COLOR_DESC" & vbCrLf _
                & " , ICTSTYC1.STYLE_COST_FIFO  STYLE_COST_FIFO" & vbCrLf _
                & " , TRUNC(NVL(WHTPPKM2.PPK_QTY,0) * NVL(WHTIADJ1.ADJQTY,0) / DECODE(NVL(WHTPPKM1.PPK_QTY_TOTAL,0),0,1,NVL(WHTPPKM1.PPK_QTY_TOTAL,0))*10000)/10000 ACTUAL_ADJ" & vbCrLf _
                & " From WHTIADJ1 WHTIADJ1, WHTSTYLX, WHTPPKM2, ICTSTYL1, ICTCOLR1, ICTSTYC1, WHTPPKM1 " & vbCrLf _
                & " Where WHTIADJ1.ABS_STATUS     <> 'A'" & vbCrLf _
                & "   And WHTSTYLX.ITEM_TYPE      =  'P'" & vbCrLf _
                & "   And WHTIADJ1.ITEM_CODE      =  WHTSTYLX.ITEM_CODE" & vbCrLf _
                & "   And WHTPPKM2.STYLE_CODE     =  ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   And WHTPPKM2.COLOR_CODE     =  ICTCOLR1.COLOR_CODE" & vbCrLf _
                & "   And WHTPPKM1.PPK_CODE       =  WHTPPKM2.PPK_CODE" & vbCrLf _
                & "   And WHTSTYLX.PPK_CODE       =  WHTPPKM2.PPK_CODE" & vbCrLf _
                & "   And ICTSTYC1.STYLE_CODE (+) =  WHTPPKM2.STYLE_CODE" & vbCrLf _
                & "   And ICTSTYC1.COLOR_CODE (+) =  WHTPPKM2.COLOR_CODE"

            Create_TDA(.Tables.Add, "WHTIADJ1_DTL", "**", 0, False, "", 0)
            Dim adjAmtExpression As String = "ISNULL(STYLE_COST_FIFO,0)* ACTUAL_ADJ"
            .Tables("WHTIADJ1_DTL").Columns.Add("ADJ_AMT", GetType(System.Decimal), adjAmtExpression)
            relation = Create_Relation("WHTIADJ1", "WHTIADJ1_DTL", "ITEM_CODE,TRANS_SEQ,LP_CODE,WHSE_CODE")

            Create_TDA(.Tables.Add, "ICTIADJ1", "*")
            Create_TDA(.Tables.Add, "ICTIADJ2", "*")

            ASCMAIN1.sql = "Select WHTPPKM2.STYLE_CODE, WHTPPKM2.COLOR_CODE" _
                & ", WHTIADJ1.ADJQTY, WHTIADJ1.LP_CODE, WHTIADJ1.WHSE_CODE, WHTIADJ1.TRANS_SEQ, WHTIADJ1.TRNDTE, WHTIADJ1.REACOD" _
                & ", WHTIADJ1.ADJ_REF1, WHTIADJ1.ADJ_REF2, ICTSTYC1.STYLE_COST_FIFO" _
                & " from WHTIADJ1,WHTPPKM2,ICTSTYC1 where ROWNUM < 1"
            Create_TDA(.Tables.Add, "WHTIADJS", "**", 0, False, "", 0)

        End With

        Dim valueList As String() = {":", _
                                     "INV:Inventory Adjustment", _
                                     "WOK:Work Order", _
                                     "SHP:Shipping Sub", _
                                     "RCP:Receipt Adjustment", _
                                     "DMG:Damaged Goods"}

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdWHTIADJ1, grdWHTIADJS, grdWHTIADJW, grdWHTIADJD}
            grd.DataSource = DVWs("WHTIADJ1")
            Create_Summary(grd, "ITEM_CODE", "Count")
            Create_Summary(grd, "ADJQTY", "SUM")

            ASCMAIN1.Add_Value_List(grd, "REACOD", , valueList)
        Next
        Create_Summary(grdWHTIADJ1, New String() {"TOTAL_ADJ_AMT", "SEL", "DEL"})

        '   grdWHTIADJD.DataSource = dst.Tables("WHTIADJD")

        grdWHTIADJ1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

        ' filterGrdWHTIADJ1 = "ABS_STATUS = 'N' AND (ISNULL(REACOD,'???') <> 'WOK' AND ISNULL(REACOD,'???') <> 'SHP')"
        filterGrdWHTIADJ1 = "ABS_STATUS = 'N' AND (ISNULL(REACOD,'???') <> 'WOK')"
        'filterGrdWHTIADJ1 = "ABS_STATUS = 'N' AND (REACOD <> 'WOK' OR REACOD IS NULL)"
        filterGrdWHTIADJS = "ABS_STATUS = 'S'"
        ' filterGrdWHTIADJW = "REACOD = 'WOK' OR REACOD = 'SHP'"
        filterGrdWHTIADJW = "REACOD = 'WOK'"
        filterGrdWHTIADJD = "ABS_STATUS = 'D'"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If EMsg = "" Then

                End If
                'LOGICAL LOCK EXAMPLE
                'If EMsg = "" Then
                'If Not ASCMAIN1.Logical_Lock("WHTSPCK1", grdICTWHSEX.ActiveRow.Cells("WHSE_CODE").Value) Then Exit Sub
                'End If

            Case "Update"
                If EMsg = "" Then

                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Mode_Settings(True)
            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Clear_Record()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Edit").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update").Visible = (EntryMode = "E")
                .Groups("Screen Control").Items("Edit").Visible = Not (EntryMode = "E")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabWHTIADJ1.Visible = ScreenMode

        grdWHTIADJ1.DisplayLayout.Bands(0).Columns("SEL").Hidden = Not (ScreenMode And EntryMode = "E")
        grdWHTIADJS.DisplayLayout.Bands(0).Columns("SEL").Hidden = Not (ScreenMode And EntryMode = "E")
        grdWHTIADJ1.DisplayLayout.Bands(0).Columns("DEL").Hidden = Not (ScreenMode And EntryMode = "E")
        grdWHTIADJS.DisplayLayout.Bands(0).Columns("DEL").Hidden = Not (ScreenMode And EntryMode = "E")

        grdWHTIADJD.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
        grdWHTIADJD.DisplayLayout.Bands(0).Columns("DEL").Hidden = True
        grdWHTIADJD.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdWHTIADJ1, grdWHTIADJS, grdWHTIADJW}
            With grd.DisplayLayout.Override
                .ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
                If ScreenMode And EntryMode = "E" Then
                    .AllowUpdate = DefaultableBoolean.True
                    For Each column As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                        If column.Key = "SEL" Or column.Key = "DEL" Or column.Key = "ABS_COMMENT" Then
                            column.CellActivation = UltraWinGrid.Activation.AllowEdit
                        Else
                            column.CellActivation = UltraWinGrid.Activation.NoEdit
                        End If
                    Next
                Else
                    .AllowUpdate = DefaultableBoolean.False

                End If
            End With
        Next

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTIADJ1", "WHTIADJ1_DTL", "ICTIADJ1", "ICTIADJ2", "WHTIADJS"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        ASCMAIN1.Progress("Retrieving data from 3PL ...")
        BeginTrans()
        TAC.WHCMAIN1.UpdateADSAndImport()
        CommitTrans()
        ASCMAIN1.Progress("Now Loading Data ...")

        EnforceConstraints(False)

        Fill_Records("WHTIADJ1")
        Fill_Records("WHTIADJ1_DTL")

        EnforceConstraints(True)

        Sort_grdColumns(grdWHTIADJ1, "ITEM_CODE", False, 0)
        Sort_grdColumns(grdWHTIADJ1, "STYLE_CODE", False, 1)
        Sort_grdColumns(grdWHTIADJS, "ITEM_CODE")
        Sort_grdColumns(grdWHTIADJW, "ITEM_CODE", False, 0)
        Sort_grdColumns(grdWHTIADJW, "STYLE_CODE", False, 1)

        For Each rowWHTIADJ1 As DataRow In dst.Tables("WHTIADJ1").Select("")
            Dim adjAmount As Decimal = 0
            Dim childRows As DataRow() = rowWHTIADJ1.GetChildRows(relation)
            If childRows.Count > 0 Then
                For Each row As DataRow In childRows
                    adjAmount += Val(row.Item("ADJ_AMT") & "")
                Next
            Else
                adjAmount = Val(rowWHTIADJ1.Item("STYLE_COST_FIFO") & "") * Val(rowWHTIADJ1.Item("ADJQTY") & "")
            End If

            rowWHTIADJ1.Item("TOTAL_ADJ_AMT") = adjAmount
        Next


        setGridFilters()
        SetupGrids()
        setVarianceColors()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
         
        Dim LP_XNO As String = TAC.WHCMAIN1.Get_LP_XNO(MENU_ITEM_OBJECT, dst.Tables("WHTIADJ1").Select().Length)
        For Each row As DataRow In dst.Tables("WHTIADJ1").Select("ABS_STATUS <> 'D'")
            'ADDED WHERE CLAUSE TO STOP DELETED ADJUSTMENTS FROM REVERTING BACK TO SUSPENDED STATUS
            row.Item("ABS_STATUS") = "S"
            row.Item("LP_XNO") = LP_XNO
        Next

        dst.Tables("WHTIADJS").Rows.Clear()

        Dim sqlw As String = "SEL = '1' or REACOD = 'WOK'"
        'Dim sqlw As String = "SEL = '1' or REACOD = 'WOK' OR REACOD = 'SHP'"
        For Each rowWHTIADJ1 As DataRow In dst.Tables("WHTIADJ1").Select(sqlw)
            If rowWHTIADJ1.Item("DEL") & "" = "1" Then
                rowWHTIADJ1.Item("ABS_STATUS") = "D"
            Else
                rowWHTIADJ1.Item("ABS_STATUS") = "A"

                Dim rowWHTIADJS As DataRow = dst.Tables("WHTIADJS").NewRow
                rowWHTIADJS.Item("REACOD") = rowWHTIADJ1.Item("REACOD")
                rowWHTIADJS.Item("TRNDTE") = CDate(rowWHTIADJ1.Item("TRNDTE")).Date
                rowWHTIADJS.Item("ADJ_REF1") = rowWHTIADJ1.Item("ADJ_REF1")
                rowWHTIADJS.Item("ADJ_REF2") = rowWHTIADJ1.Item("ADJ_REF2")

                If rowWHTIADJ1.GetChildRows(relation).Count > 0 Then 'GET CHILD ROWS IF PPK
                    For Each childRow As DataRow In rowWHTIADJ1.GetChildRows(relation)
                        rowWHTIADJS.Item("STYLE_CODE") = childRow.Item("STYLE_CODE")
                        rowWHTIADJS.Item("COLOR_CODE") = childRow.Item("COLOR_CODE")
                        rowWHTIADJS.Item("ADJQTY") = childRow.Item("ACTUAL_ADJ")
                        rowWHTIADJS.Item("LP_CODE") = childRow.Item("LP_CODE")
                        rowWHTIADJS.Item("WHSE_CODE") = childRow.Item("WHSE_CODE")
                        rowWHTIADJS.Item("TRANS_SEQ") = childRow.Item("TRANS_SEQ")
                        rowWHTIADJS.Item("STYLE_COST_FIFO") = childRow.Item("STYLE_COST_FIFO")
                    Next
                Else  'OR NOT IF STYLE CODE
                    rowWHTIADJS.Item("STYLE_CODE") = rowWHTIADJ1.Item("STYLE_CODE")
                    rowWHTIADJS.Item("COLOR_CODE") = rowWHTIADJ1.Item("COLOR_CODE")
                    rowWHTIADJS.Item("ADJQTY") = rowWHTIADJ1.Item("ADJQTY")
                    rowWHTIADJS.Item("LP_CODE") = rowWHTIADJ1.Item("LP_CODE")
                    rowWHTIADJS.Item("WHSE_CODE") = rowWHTIADJ1.Item("WHSE_CODE")
                    rowWHTIADJS.Item("TRANS_SEQ") = rowWHTIADJ1.Item("TRANS_SEQ")
                    rowWHTIADJS.Item("STYLE_COST_FIFO") = rowWHTIADJ1.Item("STYLE_COST_FIFO")
                End If
                dst.Tables("WHTIADJS").Rows.Add(rowWHTIADJS)
            End If
        Next

        For Each headerRow As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("WHTIADJS").Select(""), New String() {"WHSE_CODE", "TRNDTE", "REACOD", "ADJ_REF1", "ADJ_REF2"}).Rows

            Dim WHSE_CODE As String = headerRow.Item("WHSE_CODE")
            Dim REACOD As String = headerRow.Item("REACOD")  ' & String.Empty ' IF WE ALLOW NULLS HERE, WE WILL HAVE TO HANDLE NULL REACOD BELOW
            Dim TRNDTE As Date = headerRow.Item("TRNDTE")
            Dim ADJ_REF1 As String = headerRow.Item("ADJ_REF1") & ""
            Dim ADJ_REF2 As String = headerRow.Item("ADJ_REF2") & ""
            Dim TRAN_NO As String = ASCMAIN1.Next_Control_No("TRAN_NO_A")
            Dim TRAN_TYPE As String = "A"

            Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow()
            rowICTIADJ1.Item("ADJ_NO") = TRAN_NO
            rowICTIADJ1.Item("ADJ_DATE") = TRNDTE
            rowICTIADJ1.Item("WHSE_CODE") = WHSE_CODE
            rowICTIADJ1.Item("REASON_CODE") = "STK"
            rowICTIADJ1.Item("ADJ_NOTE") = REACOD
            rowICTIADJ1.Item("ADJ_SOURCE") = "A"
            rowICTIADJ1.Item("REGISTER_IND") = "0"
            rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIADJ1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTIADJ1.Item("TOTAL_COSTS") = 0
            rowICTIADJ1.Item("ADJ_REF") = Mid(ADJ_REF1 & IIf(ADJ_REF2 = "", "", IIf(ADJ_REF1 = "", "", ":") & ADJ_REF2), 1, 20)
            dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

            Dim TRAN_LNO As Integer = 1
            Dim TOTAL_COSTS As Decimal = 0
            'HANDLE NULL REACODS
            Dim filter As String = "WHSE_CODE = '" & WHSE_CODE & "'" _
                                   & " and REACOD = '" & REACOD & "'" _
                                   & " and TRNDTE = '" & TRNDTE & "'" _
                                   & " and ISNULL(ADJ_REF1,'') = '" & ADJ_REF1 & "'" _
                                   & " and ISNULL(ADJ_REF2,'') = '" & ADJ_REF2 & "'"
            For Each rowWHTIADJS As DataRow In dst.Tables("WHTIADJS").Select(filter)
                Dim rowICTIADJ2 As DataRow = dst.Tables("ICTIADJ2").NewRow
                rowICTIADJ2.Item("ADJ_NO") = TRAN_NO
                rowICTIADJ2.Item("ADJ_LNO") = TRAN_LNO
                rowICTIADJ2.Item("STYLE_CODE") = rowWHTIADJS.Item("STYLE_CODE")
                rowICTIADJ2.Item("COLOR_CODE") = rowWHTIADJS.Item("COLOR_CODE")
                rowICTIADJ2.Item("ADJ_QTY") = rowWHTIADJS.Item("ADJQTY")
                rowICTIADJ2.Item("STYLE_COST") = rowWHTIADJS.Item("STYLE_COST_FIFO")

                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowWHTIADJS.Item("STYLE_CODE"))

                rowICTIADJ2.Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                rowICTIADJ2.Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                rowICTIADJ2.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                rowICTIADJ2.Item("LOCATION_CODE") = DBNull.Value
                rowICTIADJ2.Item("BAR_CODE") = DBNull.Value
                rowICTIADJ2.Item("ADJ_REF") = rowWHTIADJS.Item("TRANS_SEQ")
                dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
                TOTAL_COSTS += Val(rowWHTIADJS.Item("ADJQTY") & "") * Val(rowWHTIADJS.Item("STYLE_COST_FIFO") & "")
                TRAN_LNO += 1
            Next
            rowICTIADJ1.Item("TOTAL_COSTS") = TOTAL_COSTS
        Next

        BeginTrans()
        ICCMAIN1.Update_Adjustment(Me)
        Update_Record_TDA("WHTIADJ1")
        CommitTrans("Update Complete")
    End Sub

    Sub SetupGrids()
        Dim allColumnsWithFormat As String() = {"ITEM_CODE", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
        Dim styleCaptions As String() = {"STYLE_CODE", "STYLE_DESC"}
        Dim colorCaptions As String() = {"COLOR_CODE", "COLOR_DESC"}

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdWHTIADJ1, grdWHTIADJS, grdWHTIADJW, grdWHTIADJD}

            Select Case grd.Name
                Case "grdWHTIADJ1"
                    If dst.Tables("WHTIADJ1").Select(filterGrdWHTIADJ1).Length = 0 Then grd.Text = "There are no open adjustments"
                    grd.DisplayLayout.Bands(1).Columns("ITEM_CODE").Header.Caption = "Item Code"
                Case "grdWHTIADJS"
                    If dst.Tables("WHTIADJ1").Select(filterGrdWHTIADJS).Length = 0 Then grd.Text = "There are no suspended adjustments"
                Case "grdWHTIADJW"""
                    If dst.Tables("WHTIADJ1").Select(filterGrdWHTIADJW).Length = 0 Then grd.Text = "There are no adjustments from work orders"
                Case "grdWHTIADJD"""
                    If dst.Tables("WHTIADJ1").Select(filterGrdWHTIADJD).Length = 0 Then grd.Text = "There are no Deleted Adjustments"
            End Select

            grd.DisplayLayout.Bands(0).Columns("TOTAL_ADJ_AMT").Format = "###,##0"

            For band As Integer = 0 To 1
                For Each column As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(band).Columns
                    If allColumnsWithFormat.Contains(column.Key) Then
                        column.Header.Appearance.BackColor = Drawing.Color.White
                        column.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        If column.Key = "ITEM_CODE" Then
                            column.Header.Appearance.BackColor2 = Drawing.Color.RoyalBlue
                        End If
                        If styleCaptions.Contains(column.Key) Then
                            column.Header.Appearance.BackColor2 = Drawing.Color.Green
                        End If
                        If colorCaptions.Contains(column.Key) Then
                            column.Header.Appearance.BackColor2 = Drawing.Color.Orange
                        End If
                    End If

                Next
            Next
        Next

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTIADJ1, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdWHTIADJS, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdWHTIADJW, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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
            If grd.Name = "grdWHTIADJ1" Or grd.Name = "grdWHTIADJS" Then
                For Each tool_key As String In New String() {"Select All", "De-Select All", "Select Selected", "De-Select Selected"}
                    tlb_btn = DirectCast(tlb_pop.Tools(tool_key), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                Next
            End If

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim filter As String = String.Empty
        Select Case grd.Name
            Case "grdWHTIADJ1"
                filter = filterGrdWHTIADJ1
            Case "grdWHTIADJS"
                filter = filterGrdWHTIADJS
            Case "grdWHTIADJD"
                filter = filterGrdWHTIADJD
        End Select

        Select Case e.Tool.Key
            Case "Select Selected"
                For Each row As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    row.Cells("SEL").Value = "1"
                    row.Update()
                Next
            Case "De-Select Selected"
                For Each row As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    row.Cells("SEL").Value = "0"
                    row.Update()
                Next
            Case "Select All"
                For Each row As UltraWinGrid.UltraGridRow In grd.Rows
                    row.Cells("SEL").Value = "1"
                    row.Update()
                Next
            Case "De-Select All"
                For Each row As UltraWinGrid.UltraGridRow In grd.Rows
                    row.Cells("SEL").Value = "0"
                    row.Update()
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
        End Select
    End Sub



#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'MyBase.txt_KeyDown(sender, e)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "WHSE_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Click_Command("Load", e)
        '        End If
        'End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "WHSE_CODE"
        '        Click_Command("Load")
        'End Select
    End Sub
#End Region

    Private Sub setVarianceColors()
        For Each grd As UltraWinGrid.UltraGrid In {grdWHTIADJ1, grdWHTIADJS, grdWHTIADJW}
            For Each row As UltraWinGrid.UltraGridRow In grd.Rows
                With row.Cells("TOTAL_ADJ_AMT")
                    If .Value < 0 Then
                        .Appearance.ForeColor = Drawing.Color.Red
                    Else
                        .Appearance.ForeColor = Drawing.Color.Black
                    End If
                End With
            Next
        Next
    End Sub

    Private Sub setGridFilters()
        If Not dst.Tables.Contains("WHTIADJ1") Then Exit Sub 'AVOID ERROR WHEN FORM LOADS 
        Dim filter As String = String.Empty

        Select Case tabWHTIADJ1.ActiveTab.Index
            Case 0
                filter = filterGrdWHTIADJ1
            Case 1
                filter = filterGrdWHTIADJS
            Case 2
                filter = filterGrdWHTIADJW
                grdWHTIADJW.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
                grdWHTIADJW.DisplayLayout.Bands(0).Columns("DEL").Hidden = True
            Case 3
                filter = filterGrdWHTIADJD
        End Select
        DVWs("WHTIADJ1").RowFilter = filter
    End Sub

    Private Sub tabWHTIADJ1_ActiveTabChanged(sender As Object, e As Infragistics.Win.UltraWinTabControl.ActiveTabChangedEventArgs) Handles tabWHTIADJ1.ActiveTabChanged
        setGridFilters()
    End Sub

    Private Sub grdWHTIADJ1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTIADJ1.AfterCellUpdate
        If e.Cell.Column.Key = "DEL" Then
            If e.Cell.Value & "" = "1" Then
                e.Cell.Row.Cells("SEL").Value = "1"
            End If
        End If
    End Sub

    Private Sub grdWHTIADJ1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTIADJ1.BeforeRowUpdate
        If e.Row.Cells("DEL").Value & "" = "1" Then
            e.Row.Cells("SEL").Value = "1"
        End If
    End Sub

    Private Sub grdWHTIADJS_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTIADJS.AfterCellUpdate
        If e.Cell.Column.Key = "DEL" Then
            If e.Cell.Value & "" = "1" Then
                e.Cell.Row.Cells("SEL").Value = "1"
            End If
        End If
    End Sub

    Private Sub grdWHTIADJS_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTIADJS.BeforeRowUpdate
        If e.Row.Cells("DEL").Value & "" = "1" Then
            e.Row.Cells("SEL").Value = "1"
        End If
    End Sub

    Private Sub tabWHTIADJ1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabWHTIADJ1.SelectedTabChanged

    End Sub
End Class