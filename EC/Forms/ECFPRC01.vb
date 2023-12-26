
Imports System.Text
'Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Imports Infragistics.Documents.Excel
Imports System.IO
Imports Infragistics.Win.UltraWinGrid

Public Class ECFPRC01
    Dim S As New System.Text.StringBuilder() With {.Length = 0}
    'Dim SQ1 As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
    Dim DEL_ECOM_CODES As New List(Of String)
    Dim rowECTPRCG1 As DataRow = Nothing
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        setPARMS()

        With dst
            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ECTPRCG1")
            S.AppendLine("WHERE PRCG_NO = :PARM1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ECTPRCG1", "**", 0, True, "V", 1)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ECTPRCG2")
            S.AppendLine("WHERE PRCG_NO = :PARM1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ECTPRCG2", "**", 2, True, "V", 2)

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ECTPRCG3")
            S.AppendLine("WHERE PRCG_NO = :PARM1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ECTPRCG3", "**", 3, True, "V", 3)

            'S.Length = 0
            'S.AppendLine("SELECT *")
            'S.AppendLine("FROM ECTPRCG2")
            'ASCMAIN1.sql = S.ToString()
            'Create_TDA(.Tables.Add, "ECTPRCG2", "**", 0, False)
            '.Tables("ECTPRCG2").Columns.Add("SEL", GetType(System.String))
        End With

        'Fill_Records("ECTPRCG2")
        'For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
        '    Dim COLS As String() = {"ECOM_PRICE_ADD", "ECOM_PRICE_MARKUP_PCT"}
        '    For Each COL As String In COLS
        '        If IsDBNull(rowECTPRCG2.Item(COL)) Then
        '            rowECTPRCG2.Item(COL) = 0
        '        End If
        '    Next
        'Next
        ''AddECOMFLDS()

        'For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
        '    rowECTPRCG2.Item("SEL") = "1"
        'Next
        'Fill_Records("ECTPRCG3")
        'Fill_Records("ECTPRC02")

        grdECTPRCG2.DataSource = dst.Tables("ECTPRCG2")
        grdECTPRCG3.DataSource = dst.Tables("ECTPRCG3")

        Create_Summary(grdECTPRCG3, "SKU", "Count", "", "###,###,##0")

        'ASCMAIN1.Add_Value_List(grdSOTQRDR1, "CALC_STATUS", , New String() {":", "I:Imported From Web", "L:Pulled To Laptop", "O:Finalized As Order", "X:Deleted", "M:Marked Complete", "T:Testing"})
        Sort_grdColumns(grdECTPRCG2, "ECOM_CODE", False)
        Sort_grdColumns(grdECTPRCG3, "SKU, ECOM_CODE", False)

        With grdECTPRCG3.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.True
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdECTPRCG3.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.True
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("MANUAL_PARTNER_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("MANUAL_PARTNER_PRICE").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

            ''.Columns("EDI_REPORT_DATE").Format = "MM/dd/yyyy hh:mm"
            .Columns("SET_QTY").Format = "###,###,##0"
            .Columns("STYLE_PRICE").Format = "###,###,##0.00"
            .Columns("CASE_QTY").Format = "###,###,##0"
            .Columns("WHSE_QTY_ON_HAND").Format = "###,###,##0"
            .Columns("NET_POS").Format = "###,###,##0"
            .Columns("IN_TRANS").Format = "###,###,##0"
            .Columns("FUTURE").Format = "###,###,##0"
            .Columns("STYLE_PRICE").Format = "###,###,##0.00"
            .Columns("SET_QTY").Format = "###,###,##0"
            .Columns("ECOM_UNIT_PRICE").Format = "###,###,##0.00"
            .Columns("SET_PRICE").Format = "###,###,##0.00"
            .Columns("STANDARD_PRICE").Format = "###,###,##0.00"
            .Columns("STANDARD_SET_PRICE").Format = "###,###,##0.00"
            .Columns("CARTON_SET_PRICE").Format = "###,###,##0.00"
            .Columns("STANDARD_PARTNER_PRICE").Format = "###,###,##0.00"
            .Columns("FINAL_PARTNER_PRICE").Format = "###,###,##0.00"

            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("ECOM_CODE").Header.Fixed = True

            .Columns("SET_QTY").Header.Appearance.BackColor2 = Drawing.Color.Green
            .Columns("CASE_QTY").Header.Appearance.BackColor2 = Drawing.Color.Green

            .Columns("STYLE_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("ECOM_UNIT_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Orange
            .Columns("SET_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Orange

            .Columns("STANDARD_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("STANDARD_SET_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("CARTON_SET_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("STANDARD_PARTNER_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Yellow
            .Columns("FINAL_PARTNER_PRICE").Header.Appearance.BackColor2 = Drawing.Color.Yellow
        End With

        With grdECTPRCG2.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.True
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            '.Columns("SEL").Hidden = True
            '.Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("ECOM_PRICE_ADD").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("ECOM_PRICE_MARKUP_PCT").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("ECOM_PRICE_ADD").Format = "###,###,##0.00"
            .Columns("ECOM_PRICE_MARKUP_PCT").Format = "###,###,##0.0000"
        End With

        'For i As Integer = 0 To grdECTPRCG3.DisplayLayout.Bands(0).Columns.Count - 1
        '    grdECTPRCG3.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        'Next i

        setHeaderToolTips()

        'Load_Record()

        isFormLoading = False
    End Sub

    Private Sub setHeaderToolTips()
        Dim COLS As New Dictionary(Of String, String)
        COLS.Add("STYLE_PRICE", "List Price From Style Masterfile.")
        COLS.Add("ECOM_UNIT_PRICE", "Current E-Com Price For This Parner From E-Commerce Styles.")
        COLS.Add("SET_PRICE", "Current Ecom Cost * Set Qty.")
        COLS.Add("STANDARD_PRICE", "List Price * PVC or Non-PVC Markup.")
        COLS.Add("STANDARD_SET_PRICE", "Base Price * Set Qty.")
        COLS.Add("CARTON_SET_PRICE", "Base Set Price + Carton Addition (If Applicable).")
        COLS.Add("STANDARD_PARTNER_PRICE", "(Carton + Labor) + Partner Price Add.")
        COLS.Add("MANUAL_PARTNER_PRICE", "Manual Adjustment To Final Price Provided By You.")
        COLS.Add("FINAL_PARTNER_PRICE", "Final Price To Be Uploaded To E-com Masterfile.")
        For Each col As UltraGridColumn In grdECTPRCG3.DisplayLayout.Bands(0).Columns
            If COLS.Keys.Contains(col.Key) Then
                col.Header.ToolTipText = $"{col.Header.Caption} = {COLS(col.Key)}"
            End If
        Next
    End Sub

    Private Sub ShowToolTip(sender As Object, e As System.EventArgs)
        tip.AutoPopDelay = 1000
        tip.InitialDelay = 0
        tip.DisplayStyle = ToolTipDisplayStyle.Default
        tip.ShowToolTip(sender)
    End Sub

    Private Sub setPARMS()
        Dim rowECTPARM1 As DataRow = LookUp("ECTPARM1", "Z")
        If Not IsNothing(rowECTPARM1) Then
            txtECOM_PRICE_MARKUP_PVC.Text = Format(Val(rowECTPARM1.Item("ECOM_PRICE_MARKUP_PVC").ToString & String.Empty), "##0.00")
            txtECOM_PRICE_MARKUP_NON_PVC.Text = Format(Val(rowECTPARM1.Item("ECOM_PRICE_MARKUP_NON_PVC").ToString & String.Empty), "##0.00")
            txtECOM_PRICE_CART_ADD.Text = Format(Val(rowECTPARM1.Item("ECOM_PRICE_CART_ADD").ToString & String.Empty), "$##0.00")
        End If
    End Sub

    'Private Sub AddECOMFLDS()
    '    For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
    '        Dim ECOM_CODE As String = rowECTPRCG2.Item("ECOM_CODE").ToString & String.Empty
    '        Select Case ECOM_CODE
    '            Case "AMAZON"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.15
    '            Case "HOMEDEPOT"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 1.55
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.13
    '            Case "HOUZZ"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.17
    '            Case "KIRKLANDS"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.24
    '            Case "OVERSTOCK"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 0.5
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.28
    '            Case "QVC"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 2
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.13
    '            Case "WAYFAIR"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.16279
    '            Case "XMASCENT"
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1.08
    '            Case Else
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTPRCG2.Item("ECOM_UP_PRICE_MULT") = 1
    '        End Select

    '    Next
    'End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "New"
                'EMsg = EMsg & vbCrLf & "Feature Not Done Yet."
                If txtPRCG_DESC.Text.Length = 0 Then
                    EMsg = EMsg & vbCrLf & "You Must Provide A Description For The Group."
                End If
            Case "Load"
                Dim S As New Text.StringBuilder With {.Length = 0}
                S.AppendLine("SELECT")
                S.AppendLine("PRCG_NO,")
                S.AppendLine("PRCG_DESC,")
                S.AppendLine("INIT_DATE")
                S.AppendLine("FROM ECTPRCG1")
                With ASCMAIN1.CodeSelector
                    .SQL = S.ToString
                    .MultipleSelections = False
                    .PreviouslySelectedCodes0 = ""
                    .Caption = "Select Group To Load"
                    .TABLE_NAME = ""
                    .VIEW_NAME = ""
                    .VIEW_DESC = ""
                    .COLUMN_NAME = ""
                    .COLUMN_PREKEYs = New Dictionary(Of String, String)
                    .Custom_sql_where = ""
                    .tblASTVIEW1 = New DataTable
                End With
                Dim F As New ASFCODE1
                F.ShowDialog()
                If ASCMAIN1.CodeSelector.Selections = 0 Then
                    EMsg = EMsg & vbCrLf & "No Group Selected."
                Else
                    txtPRCG_NO.Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("PRCG_NO") & ""
                    txtPRCG_DESC.Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("PRCG_DESC") & ""
                End If
            Case "Save"
                'EMsg = EMsg & vbCrLf & "Feature Not Done Yet."
            Case "Limit Excel"
                EMsg = EMsg & vbCrLf & "Feature Not Done Yet."
            Case "Refresh"
                EMsg = EMsg & vbCrLf & "Feature Not Done Yet."
            Case "Exit"
                Me.Close()
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                EntryMode = "N"
                Call Mode_Settings(True)
                Begin_New()
                Load_Partners()
                Load_Styles()
            Case "Load"
                EntryMode = "E"
                Call Mode_Settings(True)
                Load_Header()
                Load_Partners()
                Load_Styles()
            Case "Save"
                Call Mode_Settings(False)
                Update_Record()
                Clear_Record()
            Case "Limit Excel"
            Case "Refresh"
                'Call Mode_Settings(True)
                'Load_Partners()
                'Load_Styles()
                'Fill_Records("ECTTMPLT")
                'RefreshData()
            Case "Exit"
                Call Mode_Settings(False)
                Me.Close()
        End Select
    End Sub

    Private Sub Begin_New()
        txtPRCG_NO.Text = ASCMAIN1.Next_Control_No("ECTPRCG1.PRCG_NO")
        Dim newECTPRCG1 As DataRow = dst.Tables.Item("ECTPRCG1").NewRow
        newECTPRCG1.Item("PRCG_NO") = txtPRCG_NO.Text
        newECTPRCG1.Item("PRCG_DESC") = txtPRCG_DESC.Text
        newECTPRCG1.Item("PRCG_STATUS") = "O"
        newECTPRCG1.Item("PRICE_UPDATE") = Null
        newECTPRCG1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        newECTPRCG1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        newECTPRCG1.Item("INIT_DATE") = DATETIME_STAMP
        newECTPRCG1.Item("LAST_DATE") = DATETIME_STAMP
        dst.Tables.Item("ECTPRCG1").Rows.Add(newECTPRCG1)
        rowECTPRCG1 = newECTPRCG1
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("New").Visible = True
                .Groups("Screen Control").Items("Load").Visible = True
                .Groups("Screen Control").Items("Save").Visible = True
                .Groups("Screen Control").Items("Limit Excel").Visible = True
                .Groups("Screen Control").Items("Refresh").Visible = True
                .Groups("Screen Control").Items("Exit").Visible = True
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        dst.Tables("ECTPRCG1").Rows.Clear()
        dst.Tables("ECTPRCG2").Rows.Clear()
        dst.Tables("ECTPRCG3").Rows.Clear()
        txtPRCG_NO.Text = ""
        txtPRCG_DESC.Text = ""
    End Sub

    Sub Load_Header()
        S.Length = 0
        S.AppendLine("SELECT *")
        S.AppendLine("FROM ECTPRCG1")
        S.AppendLine($"WHERE PRCG_NO = '{txtPRCG_NO.Text}'")
        Fill_Records("ECTPRCG1",,, S.ToString)
        dst.AcceptChanges()
    End Sub
    Sub Load_Partners()
        Select Case EntryMode
            Case "N"
                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine($"'{txtPRCG_NO.Text}' AS PRCG_NO,")
                S.AppendLine("ECOM_CODE,")
                S.AppendLine("ECOM_NAME,")
                S.AppendLine("ECOM_PRICE_NOTES,")
                S.AppendLine("ECOM_PRICE_LAST,")
                S.AppendLine("ECOM_PRICE_ADD,")
                S.AppendLine("ECOM_PRICE_MARKUP_PCT")
                S.AppendLine("FROM ECTECOM1")
                Fill_Records("ECTPRCG2",,, S.ToString)
                For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
                    Dim COLS As String() = {"ECOM_PRICE_ADD", "ECOM_PRICE_MARKUP_PCT"}
                    For Each COL As String In COLS
                        If IsDBNull(rowECTPRCG2.Item(COL)) Then
                            rowECTPRCG2.Item(COL) = 0
                        End If
                    Next
                Next
            Case "E"
                S.Length = 0
                S.AppendLine("SELECT")
                S.AppendLine("PRCG_NO,")
                S.AppendLine("ECOM_CODE,")
                S.AppendLine("ECOM_NAME,")
                S.AppendLine("ECOM_PRICE_NOTES,")
                S.AppendLine("ECOM_PRICE_LAST,")
                S.AppendLine("ECOM_PRICE_ADD,")
                S.AppendLine("ECOM_PRICE_MARKUP_PCT")
                S.AppendLine("FROM ECTPRCG2")
                S.AppendLine($"WHERE PRCG_NO = '{txtPRCG_NO.Text}'")
                Fill_Records("ECTPRCG2",,, S.ToString)
        End Select
        'For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
        '    rowECTPRCG2.Item("SEL") = "1"
        'Next
        'Fill_Records("ECTPRCG3")
        'Fill_Records("ECTPRC02")
    End Sub

    Sub Load_Styles()
        'Call Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing Data", "")

        Select Case EntryMode
            Case "N"

                S.Length = 0
                S.AppendLine("SELECT * FROM (")
                S.AppendLine("    SELECT")
                S.AppendLine($"   '{txtPRCG_NO.Text}' AS PRCG_NO,")
                S.AppendLine("    E2.STYLE_CODE || '-' || E2.COLOR_CODE AS SKU,")
                S.AppendLine("    E1.ECOM_CODE,")
                S.AppendLine("    E2.STYLE_CODE,")
                S.AppendLine("    E2.COLOR_CODE,")
                S.AppendLine("    I1.STYLE_STATUS,")
                S.AppendLine("    I1.STYLE_DESC,")
                S.AppendLine("    NVL(I1.SIZE_CODE,'') AS SIZE_CODE,")
                S.AppendLine("    S3.ATTR_DESC,")
                S.AppendLine("    I1.CARTON_PACK_QTY AS CASE_QTY,")
                S.AppendLine("    I1.STYLE_UOM AS UOM,")
                S.AppendLine("    I1.STYLE_CLASS_CODE,")
                S.AppendLine("    NVL(E1.SHIP_DROP,0) AS SHIP_DROP,")
                S.AppendLine("    S2.WHSE_QTY_ON_HAND,")
                S.AppendLine("    S2.NET_POS,")
                S.AppendLine("    S2.IN_TRANS,")
                S.AppendLine("    S2.FUTURE,")
                S.AppendLine("    I1.STYLE_PRICE,")
                S.AppendLine("    NVL(E1.SET_QTY,0) AS SET_QTY,")
                S.AppendLine("    NVL(E1.ECOM_UNIT_PRICE,0) AS ECOM_UNIT_PRICE,")
                S.AppendLine("    (NVL(E1.ECOM_UNIT_PRICE,0) * NVL(E1.SET_QTY,0)) AS SET_PRICE,")
                S.AppendLine("    999.99 AS STANDARD_PRICE,")
                S.AppendLine("    999.99 AS STANDARD_SET_PRICE,")
                S.AppendLine("    999.99 AS CARTON_SET_PRICE,")
                S.AppendLine("    999.99 AS STANDARD_PARTNER_PRICE,")
                S.AppendLine("    999.99 AS MANUAL_PARTNER_PRICE,")
                S.AppendLine("    999.99 AS FINAL_PARTNER_PRICE")
                S.AppendLine("    FROM ECTESTY2 E2, ECTESTY1 E1, ICTSTYL1 I1,")
                S.AppendLine("    (")
                S.AppendLine("        SELECT")
                S.AppendLine("        S2.STYLE_CODE,")
                S.AppendLine("        S2.COLOR_CODE,")
                S.AppendLine("        SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
                S.AppendLine("        SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) - NVL(S2.WHSE_QTY_OPEN,0))) AS NET_POS,")
                S.AppendLine("        SUM(NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0)) AS IN_TRANS,")
                S.AppendLine("        SUM(NVL(S2.WHSE_QTY_ON_HAND,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0)) AS FUTURE")
                S.AppendLine("        FROM ICTSTAT2 S2")
                S.AppendLine("        WHERE S2.WHSE_CODE = 'MS'")
                S.AppendLine("        GROUP BY")
                S.AppendLine("        S2.STYLE_CODE,")
                S.AppendLine("        S2.COLOR_CODE,")
                S.AppendLine("        S2.WHSE_CODE")
                S.AppendLine("    ) S2,")
                S.AppendLine("    (")
                S.AppendLine("        SELECT")
                S.AppendLine("        S3.STYLE_CODE,")
                S.AppendLine("        MAX(A1.ATTR_DESC) AS ATTR_DESC")
                S.AppendLine("        FROM ICTSTYL3 S3, ICTATTR1 A1")
                S.AppendLine("        WHERE S3.ATTR_CODE = A1.ATTR_CODE")
                S.AppendLine("        AND NVL(A1.ATT_RANK,'0') = '1'")
                S.AppendLine("        GROUP BY S3.STYLE_CODE")
                S.AppendLine("    ) S3")
                S.AppendLine("    WHERE E1.STYLE_CODE = E2.STYLE_CODE")
                S.AppendLine("    AND E1.ECOM_CODE = E2.ECOM_CODE")
                S.AppendLine("    AND I1.STYLE_CODE = E2.STYLE_CODE")
                S.AppendLine("    AND S2.STYLE_CODE = E2.STYLE_CODE")
                S.AppendLine("    AND S2.COLOR_CODE = E2.COLOR_CODE")
                S.AppendLine("    AND I1.STYLE_CODE = S3.STYLE_CODE (+)")
                S.AppendLine(")")
                S.AppendLine("WHERE (STYLE_STATUS = 'A' OR STYLE_STATUS = 'N' OR ( STYLE_STATUS = 'D' AND FUTURE > 0))")
                Fill_Records("ECTPRCG3",,, S.ToString)
                Calc_Extra_Fields()
            Case "E"
                S.Length = 0
                S.AppendLine("SELECT *")
                S.AppendLine("FROM ECTPRCG3")
                S.AppendLine($"WHERE PRCG_NO = '{txtPRCG_NO.Text}'")
                Fill_Records("ECTPRCG3",,, S.ToString)
                dst.AcceptChanges()
        End Select

        EnforceConstraints(False)

        'Dim SEL_LIST As New List(Of String)
        'Dim RPL As String = ""
        'For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
        '    If rowECTPRCG2.Item("SEL").ToString & String.Empty = "1" Then
        '        If rowECTPRCG2.Item("ECOM_CODE").ToString & String.Empty <> "" Then
        '            SEL_LIST.Add("'" & rowECTPRCG2.Item("ECOM_CODE").ToString & String.Empty & "',")
        '        End If
        '    End If
        'Next
        'If SEL_LIST.Count > 0 Then
        '    Dim list As String = ""
        '    For Each l As String In SEL_LIST
        '        list += l
        '    Next
        '    list = list.Substring(0, list.Length - 1)
        '    RPL = String.Format("AND ECOM_CODE IN ({0})", list)
        'End If

        'Dim SQ As String = ""
        'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
        '    'Stop
        '    SQ = S.ToString.Replace("AND REPLACE_CUST_WHERE", RPL).ToString()
        '    'SQ = S.ToString.Replace("AND REPLACE_CUST_WHERE", "AND (STYLE_CODE, COLOR_CODE) IN (SELECT STYLE_CODE, COLOR_CODE FROM WHR_ECOM_231221) AND ECOM_CODE = 'WAYFAIR'").ToString()
        'Else
        '    SQ = S.ToString.Replace("AND REPLACE_CUST_WHERE", RPL).ToString()
        'End If

        'Dim SQ As String = SQ1.ToString.Replace("AND REPLACE_CUST_WHERE", "").ToString()

        'Fill_Records("ECTPRCG3",,, SQ)
        'Calc_Extra_Fields()

        EnforceConstraints(True)

        'If EntryMode = "N" Then
        'Else
        '    dst.AcceptChanges()
        'End If

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Calc_Extra_Fields()

        For Each rowECTPRCG3 As DataRow In dst.Tables("ECTPRCG3").Select()
            'rowECTPRCG3.Item("CURR_PRICE") = calcCURR_PRICE(rowECTPRCG3)
            'rowECTPRCG3.Item("ECOM_UP_PRICE") = calcECOM_UP_PRICE(rowECTPRCG3)
            'rowECTPRCG3.Item("ECOM_PRICE") = calcECOM_PRICE(rowECTPRCG3)
            Dim STYLE_CODE As String = rowECTPRCG3.Item("STYLE_CODE").ToString & String.Empty
            'Dim SIZE_CODE As String = rowECTPRCG3.Item("SIZE_CODE").ToString & String.Empty
            Dim SET_QTY As Int64 = Val(rowECTPRCG3.Item("SET_QTY").ToString & String.Empty)
            If SET_QTY = 0 Then SET_QTY = 1
            Dim STYLE_PRICE As Decimal = Val(rowECTPRCG3.Item("STYLE_PRICE").ToString & String.Empty)
            Dim MANUAL_PARTNER_PRICE As Decimal = Val(rowECTPRCG3.Item("MANUAL_PARTNER_PRICE").ToString & String.Empty)
            Dim STYLE_CLASS_CODE As String = rowECTPRCG3.Item("STYLE_CLASS_CODE").ToString & String.Empty
            Dim CARTON_PACK_QTY As Int64 = Val(rowECTPRCG3.Item("CASE_QTY").ToString & String.Empty)
            If CARTON_PACK_QTY = 0 Then CARTON_PACK_QTY = 1
            Dim ECOM_CODE As String = rowECTPRCG3.Item("ECOM_CODE").ToString & String.Empty
            Dim rowECTPRCG2 As DataRow = dst.Tables.Item("ECTPRCG2").Select($"ECOM_CODE = '{ECOM_CODE}'", "").FirstOrDefault
            If IsNothing(rowECTPRCG2) Then
                MsgBox($"No Record Found For Partner {ECOM_CODE}!!", vbCritical, "Pricing Not Complete!")
                Exit Sub
            End If

            Dim ECOM_PRICE_ADD As Decimal = 0
            Dim ECOM_PRICE_MARKUP_PCT As Decimal = 0
            If IsNumeric(rowECTPRCG2.Item("ECOM_PRICE_ADD").ToString & String.Empty) Then
                ECOM_PRICE_ADD = Val(rowECTPRCG2.Item("ECOM_PRICE_ADD").ToString & String.Empty)
            Else
                MsgBox($"Invalid Add For Partner {ECOM_CODE}!!", vbCritical, "Pricing Not Complete!")
                Exit Sub
            End If
            If IsNumeric(rowECTPRCG2.Item("ECOM_PRICE_MARKUP_PCT").ToString & String.Empty) Then
                ECOM_PRICE_MARKUP_PCT = Val(rowECTPRCG2.Item("ECOM_PRICE_MARKUP_PCT").ToString & String.Empty)
            Else
                MsgBox($"Invalid Markup For Partner {ECOM_CODE}!!", vbCritical, "Pricing Not Complete!")
                Exit Sub
            End If

            Dim STYLE_CLASS_CODE_MULT As Decimal = Val(txtECOM_PRICE_MARKUP_NON_PVC.Text)
            If STYLE_CLASS_CODE = "PVC" Then
                STYLE_CLASS_CODE_MULT = Val(txtECOM_PRICE_MARKUP_PVC.Text)
            End If

            Dim CARTON_PACK_QTY_ADDITION As Decimal = Val(txtECOM_PRICE_CART_ADD.Text)
            If CARTON_PACK_QTY = 1 Then
                CARTON_PACK_QTY_ADDITION = 1
            End If

            If MANUAL_PARTNER_PRICE = 999.99 Then
                MANUAL_PARTNER_PRICE = 0
            End If

            'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then Stop
            Dim STANDARD_PRICE As Decimal = 0
            Dim STANDARD_SET_PRICE As Decimal = 0
            Dim CARTON_SET_PRICE As Decimal = 0
            Dim STANDARD_PARTNER_PRICE As Decimal = 0
            Dim FINAL_PARTNER_PRICE As Decimal = 0

            STANDARD_PRICE = STYLE_PRICE * STYLE_CLASS_CODE_MULT
            STANDARD_SET_PRICE = STANDARD_PRICE * SET_QTY
            CARTON_SET_PRICE = STANDARD_SET_PRICE + CARTON_PACK_QTY_ADDITION
            STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + ECOM_PRICE_ADD
            FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * (1 + ECOM_PRICE_MARKUP_PCT)) + MANUAL_PARTNER_PRICE
            rowECTPRCG3.Item("STANDARD_PRICE") = STANDARD_PRICE
            rowECTPRCG3.Item("STANDARD_SET_PRICE") = STANDARD_SET_PRICE
            rowECTPRCG3.Item("CARTON_SET_PRICE") = CARTON_SET_PRICE
            rowECTPRCG3.Item("STANDARD_PARTNER_PRICE") = STANDARD_PARTNER_PRICE
            rowECTPRCG3.Item("FINAL_PARTNER_PRICE") = FINAL_PARTNER_PRICE

            rowECTPRCG3.Item("MANUAL_PARTNER_PRICE") = 0
            'Dim COLS As String() = {"STANDARD_PRICE", "STANDARD_SET_PRICE", "CARTON_SET_PRICE", "STANDARD_PARTNER_PRICE", "FINAL_PARTNER_PRICE"}
            'For Each COL As String In COLS
            '    Select Case COL
            '        Case "STANDARD_PRICE"
            '            'If STYLE_CODE = "MTF20964" Then Stop
            '            STANDARD_PRICE = STYLE_PRICE * STYLE_CLASS_CODE_MULT
            '            rowECTPRCG3.Item(COL) = STANDARD_PRICE
            '        Case "STANDARD_SET_PRICE"
            '            STANDARD_SET_PRICE = STANDARD_PRICE * SET_QTY
            '            rowECTPRCG3.Item(COL) = STANDARD_SET_PRICE
            '        Case "CARTON_SET_PRICE"
            '            CARTON_SET_PRICE = STANDARD_SET_PRICE + CARTON_PACK_QTY_ADDITION
            '            rowECTPRCG3.Item(COL) = CARTON_SET_PRICE
            '        Case "STANDARD_PARTNER_PRICE"
            '            Select Case ECOM_CODE
            '                Case "AMAZON"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0
            '                Case "HOMEDEPOT"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0.5
            '                Case "HOUZZ"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0
            '                Case "KIRKLANDS"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0
            '                Case "OVERSTOCK"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0.5
            '                Case "QVC"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 2
            '                Case "WAYFAIR"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0
            '                Case "XMASCENT"
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0
            '                Case Else
            '                    STANDARD_PARTNER_PRICE = CARTON_SET_PRICE + 0
            '            End Select
            '            rowECTPRCG3.Item(COL) = STANDARD_PARTNER_PRICE
            '        Case "FINAL_PARTNER_PRICE"
            '            Select Case ECOM_CODE
            '                Case "AMAZON"
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.15) + MANUAL_PARTNER_PRICE
            '                Case "HOMEDEPOT"
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.13) + MANUAL_PARTNER_PRICE
            '                Case "HOUZZ"
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.17) + MANUAL_PARTNER_PRICE
            '                Case "KIRKLANDS"
            '                    'If STYLE_CODE = "MTX72299" Then Stop
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.15) + MANUAL_PARTNER_PRICE
            '                Case "OVERSTOCK"
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.2303) + MANUAL_PARTNER_PRICE
            '                Case "QVC"
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.1875) + MANUAL_PARTNER_PRICE
            '                Case "WAYFAIR"
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.15) + MANUAL_PARTNER_PRICE
            '                Case "XMASCENT"
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1.18) + MANUAL_PARTNER_PRICE
            '                Case Else
            '                    FINAL_PARTNER_PRICE = (STANDARD_PARTNER_PRICE * 1) + MANUAL_PARTNER_PRICE
            '            End Select
            '            rowECTPRCG3.Item(COL) = FINAL_PARTNER_PRICE
            '        Case Else
            '    End Select
            'Next
        Next
    End Sub

    Private Function calcFINALPRICE() As Double

    End Function

    Sub Delete_Record(ByVal ORDR_NO As String)
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Dim SQLD As String = $" PRCG_NO = '{txtPRCG_NO.Text}'"
        Update_Record_TDA("ECTPRCG1", SQLD)
        Update_Record_TDA("ECTPRCG2", SQLD)
        Update_Record_TDA("ECTPRCG3", SQLD)
        Call CommitTrans("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdECTPRCG3, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdECTPRCG2, "BB", "Select All", "Select None")
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTORDR1"
            '    If Not InquiryOnly Then
            '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
            '    End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All"
                For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
                    rowECTPRCG2.Item("SEL") = "1"
                Next
            Case "Select None"
                For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select()
                    rowECTPRCG2.Item("SEL") = "0"
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Selected.Rows.Count = 0 Then
        '    MsgBox("You Must Select One And Only One Row First", vbOKOnly, "Select A Row")
        '    Exit Sub
        'End If

        Select Case e.Tool.Key

        End Select

        Update_Record()
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub RefreshData()
        'Using Load Records Here Instead
        'ASCMAIN1.Progress("Refreshing Styles", "")

        'Fill_Records("ECTPRCG3")

        'ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdECTPRCG2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdECTPRCG2.AfterRowActivate
        If Not IsNothing(grdECTPRCG2.ActiveRow) Then
            txtECOM_PRICE_NOTES.Text = grdECTPRCG2.ActiveRow.Cells.Item("ECOM_PRICE_NOTES").Text
        Else
            txtECOM_PRICE_NOTES.Text = ""
        End If
    End Sub

    Private Sub grdECTPRCG3_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdECTPRCG3.AfterCellUpdate
        If e.Cell.Column.Key = "MANUAL_PARTNER_PRICE" Then
            e.Cell.Row.Cells.Item("FINAL_PARTNER_PRICE").Value = Val(e.Cell.Row.Cells.Item("FINAL_PARTNER_PRICE").Value) + Val(e.Cell.Row.Cells.Item("MANUAL_PARTNER_PRICE").Value)
        End If
    End Sub

    Private Sub grdECTPRCG2_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdECTPRCG2.BeforeRowsDeleted
        If dst.Tables.Item("ECTPRCG2").Select("", "", DataViewRowState.Unchanged).Count = 1 Then
            MsgBox("You Can Not Remove Last Row!", vbCritical, "What Are You Doing?")
            e.Cancel = True
            Exit Sub
        End If
        DEL_ECOM_CODES.Clear()
        For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
            DEL_ECOM_CODES.Add(grdRow.Cells.Item("ECOM_CODE").Text)
        Next
    End Sub

    Private Sub grdECTPRCG2_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdECTPRCG2.AfterRowsDeleted
        For Each DEL_ECOM_CODE As String In DEL_ECOM_CODES
            For Each rowECTPRCG3 As DataRow In dst.Tables("ECTPRCG3").Select($"ECOM_CODE = '{DEL_ECOM_CODE}'")
                rowECTPRCG3.Delete()
            Next
        Next
        grdECTPRCG2.UpdateData()
    End Sub

#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class