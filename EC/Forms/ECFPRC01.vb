
Imports System.Text
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Imports Infragistics.Documents.Excel
Imports System.IO
Imports Infragistics.Win.UltraWinGrid

Public Class ECFPRC01
    Dim S As New System.Text.StringBuilder() With {.Length = 0}
    Dim SQ1 As New System.Text.StringBuilder() With {.Length = 0}
    Dim isFormLoading As Boolean = True
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        setPARMS()

        With dst

            SQ1.Length = 0
            SQ1.AppendLine("SELECT * FROM (")
            SQ1.AppendLine("    SELECT")
            SQ1.AppendLine("    E2.STYLE_CODE || '-' || E2.COLOR_CODE AS SKU,")
            SQ1.AppendLine("    E2.STYLE_CODE,")
            SQ1.AppendLine("    E2.COLOR_CODE,")
            SQ1.AppendLine("    I1.STYLE_STATUS,")
            SQ1.AppendLine("    I1.STYLE_DESC,")
            SQ1.AppendLine("    NVL(I1.SIZE_CODE,'') AS SIZE_CODE,")
            SQ1.AppendLine("    A1.ATTR_DESC,")
            SQ1.AppendLine("    I1.CARTON_PACK_QTY AS CASE_QTY,")
            SQ1.AppendLine("    I1.STYLE_UOM AS UOM,")
            SQ1.AppendLine("    I1.STYLE_CLASS_CODE,")
            SQ1.AppendLine("    NVL(E1.SHIP_DROP,0) AS SHIP_DROP,")
            SQ1.AppendLine("    E1.ECOM_CODE,")
            SQ1.AppendLine("    S2.WHSE_QTY_ON_HAND,")
            SQ1.AppendLine("    S2.NET_POS,")
            SQ1.AppendLine("    S2.IN_TRANS,")
            SQ1.AppendLine("    S2.FUTURE,")
            SQ1.AppendLine("    I1.STYLE_PRICE,")
            SQ1.AppendLine("    NVL(E1.SET_QTY,0) AS SET_QTY,")
            SQ1.AppendLine("    NVL(E1.ECOM_UNIT_PRICE,0) AS ECOM_UNIT_PRICE,")
            SQ1.AppendLine("    (NVL(E1.ECOM_UNIT_PRICE,0) * NVL(E1.SET_QTY,0)) AS SET_PRICE")
            SQ1.AppendLine("    FROM ECTESTY2 E2, ECTESTY1 E1, ICTATTR1 A1, ICTSTYL3 S3, ICTSTYL1 I1,")
            SQ1.AppendLine("    (")
            SQ1.AppendLine("        SELECT")
            SQ1.AppendLine("        S2.STYLE_CODE,")
            SQ1.AppendLine("        S2.COLOR_CODE,")
            SQ1.AppendLine("        SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
            SQ1.AppendLine("        SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) - NVL(S2.WHSE_QTY_OPEN,0))) AS NET_POS,")
            SQ1.AppendLine("        SUM(NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0)) AS IN_TRANS,")
            SQ1.AppendLine("        SUM(NVL(S2.WHSE_QTY_ON_HAND,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0)) AS FUTURE")
            SQ1.AppendLine("        FROM ICTSTAT2 S2")
            SQ1.AppendLine("        WHERE S2.WHSE_CODE = 'MS'")
            SQ1.AppendLine("        GROUP BY")
            SQ1.AppendLine("        S2.STYLE_CODE,")
            SQ1.AppendLine("        S2.COLOR_CODE,")
            SQ1.AppendLine("        S2.WHSE_CODE")
            SQ1.AppendLine("    ) S2")
            SQ1.AppendLine("    WHERE E1.STYLE_CODE = E2.STYLE_CODE")
            SQ1.AppendLine("    AND E1.ECOM_CODE = E2.ECOM_CODE")
            SQ1.AppendLine("    AND I1.STYLE_CODE = E2.STYLE_CODE")
            SQ1.AppendLine("    AND S2.STYLE_CODE = E2.STYLE_CODE")
            SQ1.AppendLine("    AND S2.COLOR_CODE = E2.COLOR_CODE")
            SQ1.AppendLine("    AND I1.STYLE_CODE = S3.STYLE_CODE")
            SQ1.AppendLine("    AND A1.ATTR_CODE = S3.ATTR_CODE")
            SQ1.AppendLine("    AND NVL(A1.ATT_RANK,'0') = '1'")
            SQ1.AppendLine(")")
            SQ1.AppendLine("WHERE (STYLE_STATUS = 'A' OR STYLE_STATUS = 'N' OR ( STYLE_STATUS = 'D' AND FUTURE > 0))")
            SQ1.AppendLine("AND REPLACE_CUST_WHERE")
            ASCMAIN1.sql = SQ1.ToString.Replace("AND REPLACE_CUST_WHERE", "").ToString()
            Create_TDA(.Tables.Add, "ECTPRC01", "**", 0, False)
            With .Tables("ECTPRC01").Columns
                .Add("STANDARD_PRICE", GetType(System.Decimal))
                .Add("STANDARD_SET_PRICE", GetType(System.Decimal))
                .Add("CARTON_SET_PRICE", GetType(System.Decimal))
                .Add("STANDARD_PARTNER_PRICE", GetType(System.Decimal))
                .Add("MANUAL_PARTNER_PRICE", GetType(System.Decimal))
                .Add("FINAL_PARTNER_PRICE", GetType(System.Decimal))
            End With
            dst.Tables.Item("ECTPRC01").Columns.Item("MANUAL_PARTNER_PRICE").ReadOnly = False

            S.Length = 0
            S.AppendLine("SELECT *")
            S.AppendLine("FROM ECTECOM1")
            ASCMAIN1.sql = S.ToString()
            Create_TDA(.Tables.Add, "ECTECOM1_FILTER", "**", 0, False)
            .Tables("ECTECOM1_FILTER").Columns.Add("SEL", GetType(System.String))
        End With

        Fill_Records("ECTECOM1_FILTER")
        For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
            Dim COLS As String() = {"ECOM_PRICE_ADD", "ECOM_PRICE_MARKUP_PCT"}
            For Each COL As String In COLS
                If IsDBNull(rowECTECOM1_FILTER.Item(COL)) Then
                    rowECTECOM1_FILTER.Item(COL) = 0
                End If
            Next
        Next
        'AddECOMFLDS()

        For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
            rowECTECOM1_FILTER.Item("SEL") = "1"
        Next
        'Fill_Records("ECTPRC01")
        'Fill_Records("ECTPRC02")

        grdECTECOM1_FILTER.DataSource = dst.Tables("ECTECOM1_FILTER")
        grdECTPRC01.DataSource = dst.Tables("ECTPRC01")

        Create_Summary(grdECTPRC01, "SKU", "Count", "", "###,###,##0")

        'ASCMAIN1.Add_Value_List(grdSOTQRDR1, "CALC_STATUS", , New String() {":", "I:Imported From Web", "L:Pulled To Laptop", "O:Finalized As Order", "X:Deleted", "M:Marked Complete", "T:Testing"})
        'Sort_grdColumns(grdECTTMPLT, "ItemID", False)

        With grdECTPRC01.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdECTPRC01.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
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

        With grdECTECOM1_FILTER.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("ECOM_PRICE_ADD").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("ECOM_PRICE_MARKUP_PCT").CellActivation = UltraWinGrid.Activation.AllowEdit
            .Columns("ECOM_PRICE_ADD").Format = "###,###,##0.00"
            .Columns("ECOM_PRICE_MARKUP_PCT").Format = "###,###,##0.0000"
        End With

        'For i As Integer = 0 To grdECTPRC01.DisplayLayout.Bands(0).Columns.Count - 1
        '    grdECTPRC01.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        'Next i

        setHeaderToolTips()

        Load_Record()

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
        For Each col As UltraGridColumn In grdECTPRC01.DisplayLayout.Bands(0).Columns
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
    '    For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
    '        Dim ECOM_CODE As String = rowECTECOM1_FILTER.Item("ECOM_CODE").ToString & String.Empty
    '        Select Case ECOM_CODE
    '            Case "AMAZON"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.15
    '            Case "HOMEDEPOT"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 1.55
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.13
    '            Case "HOUZZ"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.17
    '            Case "KIRKLANDS"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.24
    '            Case "OVERSTOCK"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 0.5
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.28
    '            Case "QVC"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 2
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.13
    '            Case "WAYFAIR"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.16279
    '            Case "XMASCENT"
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1.08
    '            Case Else
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_ADD") = 1
    '                rowECTECOM1_FILTER.Item("ECOM_UP_PRICE_MULT") = 1
    '        End Select

    '    Next
    'End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Refresh"

            Case "Exit"
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                Load_Record()
                'Fill_Records("ECTTMPLT")
                'RefreshData()
            Case "Exit"
                Call Mode_Settings(False)
                Me.Close()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Visible = True
                .Groups("Screen Control").Items("Exit").Visible = True
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
    End Sub

    Sub Clear_Record()
        'dst.Tables("SOTQRDR1").Rows.Clear()
    End Sub

    Sub Load_Record()
        'Call Save_Header_Fields(UltraGroupBox1)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Refreshing Data", "")

        EnforceConstraints(False)

        Dim SEL_LIST As New List(Of String)
        Dim RPL As String = ""
        For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
            If rowECTECOM1_FILTER.Item("SEL").ToString & String.Empty = "1" Then
                If rowECTECOM1_FILTER.Item("ECOM_CODE").ToString & String.Empty <> "" Then
                    SEL_LIST.Add("'" & rowECTECOM1_FILTER.Item("ECOM_CODE").ToString & String.Empty & "',")
                End If
            End If
        Next
        If SEL_LIST.Count > 0 Then
            Dim list As String = ""
            For Each l As String In SEL_LIST
                list += l
            Next
            list = list.Substring(0, list.Length - 1)
            RPL = String.Format("AND ECOM_CODE IN ({0})", list)
        End If

        Dim SQ As String = ""
        If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            SQ = SQ1.ToString.Replace("AND REPLACE_CUST_WHERE", "AND (STYLE_CODE, COLOR_CODE) IN (SELECT STYLE_CODE, COLOR_CODE FROM WHR_ECOM_231129)").ToString()
        Else
            SQ = SQ1.ToString.Replace("AND REPLACE_CUST_WHERE", RPL).ToString()
        End If

        Fill_Records("ECTPRC01",,, SQ)
        Calc_Extra_Fields()

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Calc_Extra_Fields()

        For Each rowECTPRC01 As DataRow In dst.Tables("ECTPRC01").Select()
            'rowECTPRC01.Item("CURR_PRICE") = calcCURR_PRICE(rowECTPRC01)
            'rowECTPRC01.Item("ECOM_UP_PRICE") = calcECOM_UP_PRICE(rowECTPRC01)
            'rowECTPRC01.Item("ECOM_PRICE") = calcECOM_PRICE(rowECTPRC01)
            Dim STYLE_CODE As String = rowECTPRC01.Item("STYLE_CODE").ToString & String.Empty
            'Dim SIZE_CODE As String = rowECTPRC01.Item("SIZE_CODE").ToString & String.Empty
            Dim SET_QTY As Int64 = Val(rowECTPRC01.Item("SET_QTY").ToString & String.Empty)
            If SET_QTY = 0 Then SET_QTY = 1
            Dim STYLE_PRICE As Decimal = Val(rowECTPRC01.Item("STYLE_PRICE").ToString & String.Empty)
            Dim MANUAL_PARTNER_PRICE As Decimal = Val(rowECTPRC01.Item("MANUAL_PARTNER_PRICE").ToString & String.Empty)
            Dim STYLE_CLASS_CODE As String = rowECTPRC01.Item("STYLE_CLASS_CODE").ToString & String.Empty
            Dim CARTON_PACK_QTY As Int64 = Val(rowECTPRC01.Item("CASE_QTY").ToString & String.Empty)
            If CARTON_PACK_QTY = 0 Then CARTON_PACK_QTY = 1
            Dim ECOM_CODE As String = rowECTPRC01.Item("ECOM_CODE").ToString & String.Empty
            Dim rowECTECOM1_FILTER As DataRow = dst.Tables.Item("ECTECOM1_FILTER").Select($"ECOM_CODE = '{ECOM_CODE}'", "").FirstOrDefault
            If IsNothing(rowECTECOM1_FILTER) Then
                MsgBox($"No Record Found For Partner {ECOM_CODE}!!", vbCritical, "Pricing Not Complete!")
                Exit Sub
            End If

            Dim ECOM_PRICE_ADD As Decimal = 0
            Dim ECOM_PRICE_MARKUP_PCT As Decimal = 0
            If IsNumeric(rowECTECOM1_FILTER.Item("ECOM_PRICE_ADD").ToString & String.Empty) Then
                ECOM_PRICE_ADD = Val(rowECTECOM1_FILTER.Item("ECOM_PRICE_ADD").ToString & String.Empty)
            Else
                MsgBox($"Invalid Add For Partner {ECOM_CODE}!!", vbCritical, "Pricing Not Complete!")
                Exit Sub
            End If
            If IsNumeric(rowECTECOM1_FILTER.Item("ECOM_PRICE_MARKUP_PCT").ToString & String.Empty) Then
                ECOM_PRICE_MARKUP_PCT = Val(rowECTECOM1_FILTER.Item("ECOM_PRICE_MARKUP_PCT").ToString & String.Empty)
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
            rowECTPRC01.Item("STANDARD_PRICE") = STANDARD_PRICE
            rowECTPRC01.Item("STANDARD_SET_PRICE") = STANDARD_SET_PRICE
            rowECTPRC01.Item("CARTON_SET_PRICE") = CARTON_SET_PRICE
            rowECTPRC01.Item("STANDARD_PARTNER_PRICE") = STANDARD_PARTNER_PRICE
            rowECTPRC01.Item("FINAL_PARTNER_PRICE") = FINAL_PARTNER_PRICE

            'Dim COLS As String() = {"STANDARD_PRICE", "STANDARD_SET_PRICE", "CARTON_SET_PRICE", "STANDARD_PARTNER_PRICE", "FINAL_PARTNER_PRICE"}
            'For Each COL As String In COLS
            '    Select Case COL
            '        Case "STANDARD_PRICE"
            '            'If STYLE_CODE = "MTF20964" Then Stop
            '            STANDARD_PRICE = STYLE_PRICE * STYLE_CLASS_CODE_MULT
            '            rowECTPRC01.Item(COL) = STANDARD_PRICE
            '        Case "STANDARD_SET_PRICE"
            '            STANDARD_SET_PRICE = STANDARD_PRICE * SET_QTY
            '            rowECTPRC01.Item(COL) = STANDARD_SET_PRICE
            '        Case "CARTON_SET_PRICE"
            '            CARTON_SET_PRICE = STANDARD_SET_PRICE + CARTON_PACK_QTY_ADDITION
            '            rowECTPRC01.Item(COL) = CARTON_SET_PRICE
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
            '            rowECTPRC01.Item(COL) = STANDARD_PARTNER_PRICE
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
            '            rowECTPRC01.Item(COL) = FINAL_PARTNER_PRICE
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
        'Update_Record_TDA("SOTQRDR1")
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
        Load_Popup_Menu(grdECTPRC01, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdECTECOM1_FILTER, "BB", "Select All", "Select None")
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
                For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                    rowECTECOM1_FILTER.Item("SEL") = "1"
                Next
            Case "Select None"
                For Each rowECTECOM1_FILTER As DataRow In dst.Tables("ECTECOM1_FILTER").Select()
                    rowECTECOM1_FILTER.Item("SEL") = "0"
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

        'Fill_Records("ECTPRC01")

        'ASCMAIN1.Progress("", "")
    End Sub

    Private Sub grdECTECOM1_FILTER_AfterRowActivate(sender As Object, e As EventArgs) Handles grdECTECOM1_FILTER.AfterRowActivate
        If Not IsNothing(grdECTECOM1_FILTER.ActiveRow) Then
            txtECOM_PRICE_NOTES.Text = grdECTECOM1_FILTER.ActiveRow.Cells.Item("ECOM_PRICE_NOTES").Text
        Else
            txtECOM_PRICE_NOTES.Text = ""
        End If
    End Sub

    Private Sub grdECTPRC01_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdECTPRC01.AfterCellUpdate
        If e.Cell.Column.Key = "MANUAL_PARTNER_PRICE" Then
            e.Cell.Row.Cells.Item("FINAL_PARTNER_PRICE").Value = Val(e.Cell.Row.Cells.Item("FINAL_PARTNER_PRICE").Value) + Val(e.Cell.Row.Cells.Item("MANUAL_PARTNER_PRICE").Value)
        End If
    End Sub

#End Region

#Region "Form Controls"

#Region "Grids"

#End Region
#End Region

End Class