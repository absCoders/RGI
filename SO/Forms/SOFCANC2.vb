Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

Public Class SOFCANC2

    Const maxStyles As Integer = 400
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim WHSE_CODE As String
    Dim rowICTWHSE1 As DataRow
    Dim SOTCANCY As String = ""
    Dim flgSOTCANCX As Boolean = True

    Dim sqlORDR_GROUP_NOs As String

    Dim iSC As New Dictionary(Of String, Integer)


    ' DETAIL HAS ORDR_STATUS TOO
    'MyBase.Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2, "ORDR_QTY_OPEN,ORDR_QTY_CANC")

    ' FINISH MAKING S=-1 TO USE ADO.NET


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Get_PARM("ICTPARM1")

        Create_Temp_Tables(True)

        With dst

            ASCMAIN1.sql = "Select SOTCANCY.*" & vbCrLf _
                & " from " & SOTCANCY & " SOTCANCY"
            MyBase.Create_TDA(.Tables.Add, "SOTCANCY", "**", 0, False, "", 4)

            '            select * from sotordr0, 
            '(select ordr_group_no, sum(sotordr9.RANGE_STYLE_QTY) RANGE_QTY_OPEN from sotordr1, sotordr9
            'where sotordr9.ORDR_NO = sotordr1.ORDR_NO
            'group by ordr_group_no) sotordr9 
            'where sotordr0.cust_code = 'MEIJER'
            'And sotordr0.ordr_date = '13-MAR-2017'
            'And sotordr0.ORDR_GROUP_NO = sotordr9.ORDR_GROUP_NO;


            ASCMAIN1.sql = "Select SOTORDR0.*, SOTORDR9.RANGE_QTY_OPEN" & vbCrLf _
                & " from SOTORDR0, " & vbCrLf _
                & " (select ordr_group_no, sum(sotordr9.RANGE_STYLE_QTY) RANGE_QTY_OPEN from sotordr1, sotordr9 " & vbCrLf _
                & " where sotordr9.ORDR_NO = sotordr1.ORDR_NO " & vbCrLf _
                & " group by ordr_group_no) SOTORDR9 " & vbCrLf _
                & " where CUST_CODE = :PARM1 and ORDR_DATE = :PARM2 and WHSE_CODE = :PARM3" & vbCrLf _
                & " And SOTORDR0.ORDR_GROUP_NO = SOTORDR9.ORDR_GROUP_NO"
            MyBase.Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "VDV", 1)
            With .Tables("SOTORDR0")
                .Columns.Add("SELECTED", GetType(System.String)) ', "IIF(ordr_qty - (ordr_qty_open + ordr_qty_canc) = 0 ,'1','0')")
                .Columns("SELECTED").DefaultValue = "1"
            End With

            ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV2,SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.CUST_CODE = :PARM1 " & vbCrLf _
                & "   and SOTRSRV2.STYLE_CODE = :PARM2 " & vbCrLf _
                & "   and SOTRSRV2.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
                & "   and SOTRSRV1.RSRV_STATUS = 'O'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_QTY_OPEN > 0" & vbCrLf
            Create_TDA(.Tables.Add, "SOTRSRVX", "**", 0, False, "VVV", 0)

            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & " where ORDR_NO = :PARM1"
            MyBase.Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, True, "V", 1, "ORDR_STATUS")

            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & " from SOTORDR2" & vbCrLf _
                & " where ORDR_NO = :PARM1"
            MyBase.Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2, "ORDR_QTY, ORDR_QTY_OPEN,ORDR_QTY_CANC")

            ASCMAIN1.sql = "Select SOTORDR9.*" & vbCrLf _
                & " from SOTORDR9" & vbCrLf _
                & " where ORDR_NO = :PARM1"
            MyBase.Create_TDA(.Tables.Add, "SOTORDR9", "**", 0, True, "V", 2, "RANGE_STYLE_QTY, RANGE_STYLE_PP_QTY")

            Create_TDA(.Tables.Add, "SOTRSRV1", "*")
            Create_TDA(.Tables.Add, "SOTRSRV2", "*")

            ASCMAIN1.sql = "Select RANGE_STYLE_LNO, RANGE_STYLE_CODE, CUST_STORE_NO, ORDR_NO, ORDR_STATUS, RANGE_STYLE_QTY_PER_PP, RANGE_STYLE_PP_QTY" & vbCrLf _
                & ", ORDR_QTY, ORDR_QTY_OPEN, ORDR_QTY_PICK, ORDR_QTY_SHIP, ORDR_QTY_CANC, RANGE_STYLE_TTL " & vbCrLf _
                & " from (Select RANGE_STYLE_LNO, RANGE_STYLE_CODE, CUST_STORE_NO, ORDR_NO, ORDR_STATUS, RANGE_STYLE_QTY_PER_PP, RANGE_STYLE_PP_QTY" & vbCrLf _
                & ", Sum (ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", Sum (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", Sum (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", Sum (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", Sum(QTY_PER_PP) RANGE_STYLE_TTL" & vbCrLf _
                & " from " & SOTCANCY & " group by RANGE_STYLE_LNO, RANGE_STYLE_CODE, CUST_STORE_NO, ORDR_NO, ORDR_STATUS, RANGE_STYLE_QTY_PER_PP, RANGE_STYLE_PP_QTY) "

            MyBase.Create_TDA(.Tables.Add, "SOTCANCS", "**", 0, False, String.Empty, 4)

            ASCMAIN1.sql = "Select RANGE_STYLE_CODE " & vbCrLf _
                & " from (Select RANGE_STYLE_CODE" & vbCrLf _
                & ", Sum (ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", Sum (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", Sum (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", Sum (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & " from " & SOTCANCY & " group by RANGE_STYLE_CODE) "

            MyBase.Create_TDA(.Tables.Add, "ICTSTYLX", "**", 0, False, String.Empty, 1)

            Create_Relation("ICTSTYLX", "SOTCANCS", "RANGE_STYLE_CODE")
            With .Tables("ICTSTYLX")
                .Columns.Add("ORDR_QTY", GetType(System.Int64), "SUM(CHILD.ORDR_QTY)")
                .Columns.Add("ORDR_QTY_OPEN", GetType(System.Int64), "SUM(CHILD.ORDR_QTY_OPEN)")
                .Columns.Add("ORDR_QTY_PICK", GetType(System.Int64), "SUM(CHILD.ORDR_QTY_PICK)")
                .Columns.Add("ORDR_QTY_SHIP", GetType(System.Int64), "SUM(CHILD.ORDR_QTY_SHIP)")
                .Columns.Add("ORDR_QTY_CANC", GetType(System.Int64), "SUM(CHILD.ORDR_QTY_CANC)")
            End With

            With .Tables("ICTSTYLX")
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "1"
            End With

            ASCMAIN1.sql = "Select CUST_STORE_NO, ORDR_NO, ORDR_STATUS from SOTORDR1" & vbCrLf _
                & " where ORDR_NO in (Select ORDR_NO from " & SOTCANCY & ")"
            Create_TDA(.Tables.Add, "SOTCANCX", "**", 0, False, String.Empty, 2)
            For iCtr As Integer = 1 To maxStyles
                .Tables("SOTCANCX").Columns.Add("Q" & Format(iCtr, "00"), GetType(System.Int64))
            Next

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_OPEN from ICTSTAT2"
            Create_TDA(.Tables.Add, "ICTSTATO", "**", 0, False, String.Empty, 3)
        End With

        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")
        grdSOTCANCX.DataSource = dst.Tables("SOTCANCX")
        grdSOTCANCS.DataSource = dst.Tables("SOTCANCS")
        grdSOTCANCY.DataSource = dst.Tables("SOTCANCY")

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, "SELECTED", "Sum")

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("SELECTED").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "ORDR_GROUP_NO", "ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key.StartsWith("ORDR_AMT") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    GCOL.Width = 80
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTORDR0, GCOL.Key)
                ElseIf GCOL.Key.StartsWith("ORDR_QTY") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    GCOL.Width = 70
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTORDR0, GCOL.Key)
                ElseIf GCOL.Key.StartsWith("ORDR_CNT") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                    GCOL.Width = 50
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTORDR0, GCOL.Key)
                ElseIf New String() {"ORDR_GROUP_NO", "ORDR_CUST_PO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightCoral
                    GCOL.Width = 110
                ElseIf New String() {"CUST_DC_NO", "ORDR_DEPT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightCoral
                    GCOL.Width = 70
                ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    GCOL.Width = 90
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdICTSTYLX, "RANGE_STYLE_CODE", "Count")
        Create_Summary(grdICTSTYLX, "SELECTED", "Sum")
        Create_Summary(grdICTSTYLX, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})

        With grdICTSTYLX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "RANGE_STYLE_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite ' WhiteSmoke
                End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key = "RANGE_STYLE_CODE" Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf GCOL.Key.StartsWith("ORDR_QTY") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    GCOL.Width = 65
                    GCOL.Format = "#,##0"
                ElseIf New String() {"WHSE_QTY_ON_HAND", "WHSE_QTY_ON_ORDER", "WHSE_QTY_TRAN", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    GCOL.Width = 65
                    GCOL.Format = "#,##0"
                ElseIf New String() {"OTS_INV", "NET_POS"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                    GCOL.Width = 65
                    GCOL.Format = "#,##0"
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

        With grdSOTCANCX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "ORDR_NO", "ORDR_STATUS"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = System.Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
            Next

        End With
        grdSOTCANCX.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.PaleGreen

        Create_Summary(grdSOTCANCS, "CUST_STORE_NO", "Count")
        With grdSOTCANCS.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "ORDR_NO", "ORDR_STATUS"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "RANGE_STYLE_PP_QTY" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key.StartsWith("RANGE_STYLE_PP_QTY") Or GCOL.Key.StartsWith("ORDR_QTY") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    GCOL.Width = 70
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTCANCS, GCOL.Key)
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdSOTCANCY, "CUST_STORE_NO", "Count")
        With grdSOTCANCY.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "ORDR_NO", "ORDR_STATUS"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                'If GCOL.Key = "ORDR_QTY_OPEN" Then
                '    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                GCOL.CellAppearance.BackColor = System.Drawing.Color.GhostWhite
                'End If
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key.StartsWith("ORDR_QTY") Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    GCOL.Width = 70
                    GCOL.Format = "#,##0"
                    Create_Summary(grdSOTCANCY, GCOL.Key)
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
        End With

        'ASCMAIN1.Add_Value_List(grdICTSTYLX, "ITEM_STATUS")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View", "Edit"

                CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

                If CUST_CODE = "" Then
                    EMsg &= vbCr & "No Customer Defined"
                Else
                    If grdSOTORDR0.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Orders in Selection Grid"
                    Else
                        Dim rows() As DataRow = dst.Tables("SOTORDR0").Select("SELECTED='1'")
                        If rows.Length = 0 Then
                            EMsg &= vbCr & "No Orders Selected"
                        Else
                            If rows(0).Item("CUST_CODE") <> CUST_CODE Then
                                EMsg &= vbCr & "Orders in Selection grid do not appear to belong to Customer Defined"
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    rowARTCUST1 = Lookup("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Code"
                    End If
                    rowICTWHSE1 = Lookup("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Warehouse Code"
                    End If
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        For Each row As DataRow In dst.Tables("SOTORDR0").Select("SELECTED = '1'")
                            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        Next
                        If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE) Then Exit Sub
                        If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("Are you sure you want to Cancel your changes?",
                            MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If dst.Tables("SOTCANCY").Select("ORDR_QTY_OPEN <> ORIG_QTY_OPEN").Length = 0 Then
                    EMsg &= vbCr & "No records have been updated"
                End If
        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View", "Edit"
                MyBase.EntryMode = Mid(eItemKey, 1, 1)
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Cancel"
                Me.Mode_Settings(False)

            Case "Print"
                Me.Print_Record()

            Case "Update"
                Me.Update_Record()
                Me.Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    If Not ScreenMode Or (EntryMode = "V") Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode

                    .Items("Cancel").Visible = (EntryMode = "N" Or EntryMode = "E")
                    .Items("Update").Visible = (EntryMode = "N" Or EntryMode = "E")

                    .Items("Print").Settings.Enabled = iScreenMode
                End With
                ' .Groups("Display Options").Visible = ScreenMode
                Setup_tabSOTCANCX()
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        splSOTCANCX.Visible = ScreenMode
        grdSOTORDR0.Visible = Not ScreenMode

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid _
                In New UltraWinGrid.UltraGrid() _
                {grdICTSTYLX, grdSOTCANCX, grdSOTCANCS, grdSOTCANCY}
                With grd.DisplayLayout.Override
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        If grd.Name = "grdSOTCANCX" Or grd.Name = "grdICTSTYLX" Or grd.Name = "grdSOTCANCY" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        End If
                        '  .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.True
                        .AllowDelete = DefaultableBoolean.False
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    End If
                End With
            Next
        End If

        With grdSOTORDR0.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With

        'With grdICTSTYLX.DisplayLayout.Override
        '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .AllowUpdate = DefaultableBoolean.True
        '    .AllowDelete = DefaultableBoolean.False
        'End With

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTCANCX", "ICTSTYLX", "SOTORDR0", "SOTCANCS", "SOTCANCY", "SOTORDR1", "SOTORDR2", "SOTRSRVX", "SOTRSRV1", "SOTRSRV2", "ICTSTATO"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        MyBase.EnforceConstraints(True)

        For i As Integer = 1 To maxStyles
            With grdSOTCANCX.DisplayLayout.Bands(0).Columns("Q" & Format(i, "00"))
                .Header.Caption = ""
                .Hidden = True
                .Tag = ""
            End With
        Next

        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
        iSC.Clear()

        If ASCMAIN1.Running_in_VS And CUST_CODE = "" Then
            Absx1.txtFor("CUST_CODE").Text = "COSTCOUS"
            Absx1.txtFor("WHSE_CODE").Text = "CAWHSE"
            dteORDR_DATE.Value = CDate("12/2/2020")
        End If

        Clear_All_Filters(grdSOTORDR0)
    End Sub

    Private Sub Load_Record()

        MyBase.EnforceConstraints(False)

        Create_Temp_Tables(False)

        optQ.Value = "ORDR_QTY_OPEN"

        Fill_Records("SOTCANCS")
        Fill_Records("SOTCANCY")
        Fill_Records("SOTCANCX")
        Sort_grdColumns(grdSOTCANCX, "CUST_STORE_NO")

        Fill_Records("ICTSTYLX")
        Sort_grdColumns(grdICTSTYLX, "RANGE_STYLE_CODE")

        With grdSOTCANCX.DisplayLayout.Bands(0)
            .Summaries.Clear()
            Create_Summary(grdSOTCANCX, "CUST_STORE_NO", "Count")
            For I As Integer = 1 To maxStyles
                .Columns("Q" & Format(I, "00")).Hidden = True
            Next

            iSC.Clear()
            Dim iCol As Integer = 0
            For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("", "RANGE_STYLE_CODE")
                Dim RANGE_STYLE_CODE As String = rowICTSTYLX.Item("RANGE_STYLE_CODE")
                'Dim COLOR_CODE As String = rowICTSTYLX.Item("COLOR_CODE")
                'Dim STYLE_DESC As String = rowICTSTYLX.Item("STYLE_DESC")
                iCol += 1
                Dim C As String = "Q" & Format(iCol, "00")
                With .Columns(C)
                    .Hidden = False
                    .Header.Caption = RANGE_STYLE_CODE
                    .Width = 80
                    .Format = "#,##0"
                    .Header.ToolTipText = RANGE_STYLE_CODE
                End With
                Create_Summary(grdSOTCANCX, C)
                iSC.Add(RANGE_STYLE_CODE, iCol)
            Next
        End With

        Load_Qtys(True)

        If grdICTSTYLX.ActiveRow Is Nothing Then
            Highlight_Style("")
        Else
            Highlight_Style(grdICTSTYLX.ActiveRow.Cells("RANGE_STYLE_CODE").Value)
        End If

        MyBase.EnforceConstraints(True)

        Sort_grdColumns(grdSOTCANCX, "CUST_STORE_NO")
        Update_Totals()

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Sub Print_Record()
        Create_Report()
    End Sub

    Function Create_Report() As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Dim REPORT_NAME As String = "SORALLO1"
        Dim RPT As String = REPORT_NAME

        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        dst.Tables("SOTALLOZ").Rows.Clear()

        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLOZ").Rows.Clear()

        Dim STYLE_CODEs As New List(Of String)
        Dim CUST_CODEs As New List(Of String)

        For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO1")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() _
                    {"ALLO_CTL_NO", "STYLE_CODE", "DATE_START", "DATE_END", "INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE", "ALLOW_OVER",
                     "ITEM_DESC", "COLLECTION_CODE", "BRAND_CODE", "ITEM_BASIC_PROMO", "ITEM_SNU_CODE", "QTY_ALLO_PLAN", "QTY_ALLO_TOTAL", "ITEM_DATE_TO_SHIP"}
                    If COLUMN_NAME = "BRAND_CODE" Then
                    Else
                        rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                    End If
                Next
                .Rows.Add(rowR)
            End With

            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")

            Fill_Records("SOTALLOZ", ALLO_CTL_NO, False)


            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            If Not STYLE_CODEs.Contains(STYLE_CODE) Then
                Fill_Records("ICTSTAT2", STYLE_CODE, False)

                Dim rowR As DataRow = REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").NewRow
                For Each DC As DataColumn In dst.Tables("ICTITEM1").Columns
                    Dim COLUMN_NAME As String = DC.ColumnName
                    If REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").Columns.Contains(COLUMN_NAME) Then
                        '      rowR.Item(COLUMN_NAME) = rowICTITEM1.Item(COLUMN_NAME)
                    End If
                Next
                REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTITEM1").Rows.Add(rowR)

                Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
                Dim imgba() As Byte = Nothing
                Dim IMAGE_FILENAME As String = FOLDER_NAME & "\" & STYLE_CODE & ".JPG"
                If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                    rowR.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                Else
                    IMAGE_FILENAME = FOLDER_NAME & "\" & STYLE_CODE & ".PNG"
                    If My.Computer.FileSystem.FileExists(IMAGE_FILENAME) Then
                        rowR.Item("ITEM_IMAGE") = ASCMAIN1.GetImageData(IMAGE_FILENAME)
                    End If
                End If
            End If
        Next

        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()

            Dim SUBT As String = "Allocations by Item/Customer (Screen Report)"
            .CR_params.Add("SUBT", SUBT) ' "")
            .CR_params.Add("PAGE_EJECT", "0")
            .CR_params.Add("EXC_ONLY", "0")
            .CR_params.Add("SUMMARY", "0")
            .Generate_Report(RPT, Me.Text, SUBT)
            .Print_Report_End()

        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return ""
    End Function

    Private Sub Update_Record()
        Dim ORDR_L As String = ""
        Dim ORDR_NO As String = ""
        Dim ORDR_STATUS As String = ""
        Dim OPEN As Int64 = 0
        Dim RANGE_STYLE_LNO As String = ""
        Dim RANGE_L As String = ""

        Dim rowSOTORDR1 As DataRow
        Dim rowSOTORDR2 As DataRow
        Dim rowSOTORDR9 As DataRow
        Dim rowSOTCANCS As DataRow
        Dim ORDR_GROUP_NOs As New List(Of String)
        Dim ORDR_GROUP_NO As String = ""

        Dim ORDR_NOs As New List(Of String)

        'dst.Tables("SOTORDR1").Rows.Clear()
        'dst.Tables("SOTORDR2").Rows.Clear()

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_GROUP_NO in (" & Mid(sqlORDR_GROUP_NOs, 2) & ")"
        Fill_Records("SOTORDR1", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO in (Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO in (" & Mid(sqlORDR_GROUP_NOs, 2) & "))"
        Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select * from SOTORDR9 where ORDR_NO in (Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO in (" & Mid(sqlORDR_GROUP_NOs, 2) & "))"
        Fill_Records("SOTORDR9", "", True, ASCMAIN1.sql)

        dst.Tables("ICTSTATO").Rows.Clear()

        'dst.Tables("SOTRSRVX").Rows.Clear()


        'Dim use_binding As Boolean = True
        'Create_BAs("SOTORDR2")


        Try
            MyBase.BeginTrans()
            ASCMAIN1.Progress("Now Retracting Commitments")
            For Each row As DataRow In dst.Tables("SOTCANCY").Select("ORDR_QTY_OPEN <> ORIG_QTY_OPEN", "ORDR_NO, CUST_STORE_NO, ORDR_LNO")
                ORDR_NO = row.Item("ORDR_NO")
                RANGE_STYLE_LNO = row("RANGE_STYLE_LNO")
                ASCMAIN1.Progress("-", ORDR_NO)

                If RANGE_STYLE_LNO <> RANGE_L Or ORDR_NO <> ORDR_L Then
                    rowSOTCANCS = dst.Tables("SOTCANCS").Select("ORDR_NO = '" & ORDR_NO & "' and RANGE_STYLE_LNO = '" & RANGE_STYLE_LNO & "'").First
                    rowSOTORDR9 = dst.Tables("SOTORDR9").Select("ORDR_NO = '" & ORDR_NO & "' and RANGE_STYLE_LNO = '" & RANGE_STYLE_LNO & "'").First
                    rowSOTORDR9("RANGE_STYLE_PP_QTY") = rowSOTCANCS("RANGE_STYLE_PP_QTY") & ""
                    rowSOTORDR9("RANGE_STYLE_QTY") = rowSOTCANCS("ORDR_QTY_OPEN") & ""
                End If

                If ORDR_NO <> ORDR_L Then
                    ORDR_NOs.Add(ORDR_NO)
                    'Fill_Records("SOTORDR1", New Object() {ORDR_NO}, False)
                    'Fill_Records("SOTORDR2", New Object() {ORDR_NO}, False)
                    'rowSOTORDR1 = dst.Tables("SOTORDR1").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO)).First
                    rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO) ' (String.Format("ORDR_NO = '{0}'", ORDR_NO)).First
                    ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO")
                    If ORDR_GROUP_NOs.IndexOf(ORDR_GROUP_NO) = -1 Then
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                    End If
                    Dependent_Updates(-1, ORDR_NO, ORDR_GROUP_NO)
                End If
                rowSOTORDR2 = dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}' and ORDR_LNO = {1}", ORDR_NO, row.Item("ORDR_LNO"))).First
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = row.Item("ORDR_QTY_OPEN")
                rowSOTORDR2.Item("ORDR_QTY_CANC") = row.Item("ORDR_QTY_CANC")
                rowSOTORDR2.Item("ORDR_QTY") = row.Item("ORDR_QTY") '******* Why wasn't this updated when it ran?

                ORDR_L = ORDR_NO
                RANGE_L = RANGE_STYLE_LNO
            Next

            ASCMAIN1.sql = "Select SOTRSRV1.* from SOTRSRV1" & vbCrLf _
                & " where SOTRSRV1.RSRV_STATUS = 'O' and SOTRSRV1.CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SOTRSRV1", "", True, ASCMAIN1.sql)
            ASCMAIN1.sql = "Select SOTRSRV2.* from SOTRSRV1,SOTRSRV2" & vbCrLf _
                & " where SOTRSRV1.RSRV_STATUS = 'O' and SOTRSRV1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO"
            Fill_Records("SOTRSRV2", "", True, ASCMAIN1.sql)

            ' Update_Record_TDA("SOTORDR2")

            ASCMAIN1.Progress("Now Updating Sales Order Commitments")
            For Each ORDR_NO In ORDR_NOs
                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

                'ORDR_NO = rowSOTORDR1.Item("ORDR_NO")
                ASCMAIN1.Progress("-", ORDR_NO)
                ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO")
                ORDR_STATUS = ""
                OPEN = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", String.Format("ORDR_NO = '{0}'", ORDR_NO) & ""))

                If OPEN <> 0 Then
                    ORDR_STATUS = "O"
                Else
                    ORDR_STATUS = "C"
                End If
                rowSOTORDR1.Item("ORDR_STATUS") = ORDR_STATUS
                Dependent_Updates(1, ORDR_NO, ORDR_GROUP_NO)

            Next

            ASCMAIN1.Progress("Now Saving Orders & Reservations")

            Update_Record_TDA("SOTORDR1")

            'Dim tbl As DataTable = dst.Tables("SOTORDR2").Clone

            Update_Record_TDA("SOTORDR2")

            Update_Record_TDA("SOTORDR9")
            'Update_BAs("SOTORDR2")

            Update_Record_TDA("SOTRSRV1")
            Update_Record_TDA("SOTRSRV2")

            ASCMAIN1.Progress("Now Updating Style/Color Commitments")

            For Each rowICTSTATO As DataRow In dst.Tables("ICTSTATO").Select("")
                Dim STYLE_CODE As String = rowICTSTATO.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowICTSTATO.Item("COLOR_CODE")
                Dim WHSE_CODE As String = rowICTSTATO.Item("WHSE_CODE")
                Dim WHSE_QTY_OPEN As String = Val(rowICTSTATO.Item("WHSE_QTY_OPEN") & "")
                If WHSE_QTY_OPEN <> 0 Then
                    TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", WHSE_QTY_OPEN)
                End If
            Next

            For Each ORDR_GROUP_NO In ORDR_GROUP_NOs
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
            Next

            ASCMAIN1.Progress("")
            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub
    Sub Dependent_Updates(S As Integer, ORDR_NO As String, ORDR_GROUP_NO As String)

        Dim QTY_TO_COMMIT As Int64

        Dim rowSOTORDR1 As DataRow = Nothing
        Dim rowSOTORDR2s() As DataRow

        If S = -1 Then
            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
            rowSOTORDR1 = ASCDATA1.GetDataRow

            ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
            rowSOTORDR2s = ASCDATA1.GetDataTable.Select("")
        Else
            rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            rowSOTORDR2s = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
        End If

        'ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        'For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
        For Each rowSOTORDR2 As DataRow In rowSOTORDR2s
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")
            Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")

            If S = -1 Then
                If rowSOTORDR2.Item("RSRV_NO") & "" <> "" Then
                    'Only restore this reservation line if it hasn't been substitutioned.  Per Gabe 07/30/02 - WR.
                    Dim row As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, rowSOTORDR2.Item("ORDR_LNO")})
                    If row IsNot Nothing Then  'Added for Angela. 1/24/05.  She was adding styles to range that had pulled from reservation already.
                        If row.Item("STYLE_CODE_SUB") & "" = "" Then
                            Update_SOTRSRVx(rowSOTORDR2, S, ORDR_GROUP_NO)
                        End If
                    End If
                End If
            Else

                Dim rowSOTRSRVX As DataRow = Nothing ' Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                '& " order by SOTRSRV1.ORDR_CANCEL_DATE"

                If S = -1 Then
                    rowSOTRSRVX = Fill_Record("SOTRSRVX", New String() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                Else
                    Dim sqlw As String = String.Format("STYLE_CODE = '{0}' and COLOR_CODE = '{1}' and RSRV_QTY_OPEN > 0", STYLE_CODE, COLOR_CODE)
                    Dim rows() As DataRow = dst.Tables("SOTRSRV2").Select(sqlw)
                    If rows.Length <> 0 Then
                        rowSOTRSRVX = rows(0)
                    End If
                End If

                Dim Ps() As Object

                If rowSOTRSRVX IsNot Nothing Then
                    rowSOTORDR2.Item("RSRV_NO") = rowSOTRSRVX.Item("RSRV_NO")
                    rowSOTORDR2.Item("RSRV_LNO") = rowSOTRSRVX.Item("RSRV_LNO")
                    Ps = {rowSOTRSRVX.Item("RSRV_NO"), rowSOTRSRVX.Item("RSRV_LNO")}
                    Update_SOTRSRVx(rowSOTORDR2, S, ORDR_GROUP_NO)
                Else
                    rowSOTORDR2.Item("RSRV_NO") = DBNull.Value
                    rowSOTORDR2.Item("RSRV_LNO") = DBNull.Value
                    Ps = {DBNull.Value, DBNull.Value}
                End If

                'Update_Record_TDA("SOTORDR2")

                'ASCMAIN1.sql = "Update SOTORDR2 Set RSRV_NO = :PARM1, RSRV_LNO = :PARM2" _
                '    & " where ORDR_NO = '" & ORDR_NO & "' and ORDR_LNO = " & CStr(ORDR_LNO)
                'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VN", Ps)
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTORDR2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTORDR2.Item("COLOR_CODE")
                If S = -1 Then
                    TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", S * QTY_TO_COMMIT)
                Else
                    Dim rowICTSTATO As DataRow = dst.Tables("ICTSTATO").Rows.Find(New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE})
                    If rowICTSTATO Is Nothing Then
                        rowICTSTATO = dst.Tables("ICTSTATO").Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, 0})
                    End If
                    rowICTSTATO.Item("WHSE_QTY_OPEN") = Val(rowICTSTATO.Item("WHSE_QTY_OPEN") & "") + QTY_TO_COMMIT
                End If
            End If
        Next

    End Sub

    Sub Update_SOTRSRVx(rowSOTORDR2 As DataRow, S As Integer, ORDR_GROUP_NO As String)
        Dim RSRV_NO As String = rowSOTORDR2.Item("RSRV_NO") & ""
        Dim RSRV_LNO As Int64 = Val(rowSOTORDR2.Item("RSRV_LNO") & "")

        Dim rowSOTRSRV1 As DataRow = Nothing
        Dim rowSOTRSRV2 As DataRow = Nothing
        If S = -1 Then
            rowSOTRSRV1 = Fill_Record("SOTRSRV1", RSRV_NO)
            rowSOTRSRV2 = Fill_Record("SOTRSRV2", New String() {RSRV_NO, RSRV_LNO})
        Else
            rowSOTRSRV1 = dst.Tables("SOTRSRV1").Rows.Find(RSRV_NO)
            rowSOTRSRV2 = dst.Tables("SOTRSRV2").Rows.Find(New Object() {RSRV_NO, RSRV_LNO})
        End If
        Dim WHSE_CODE As String = rowSOTRSRV1.Item("WHSE_CODE")

        With rowSOTRSRV2
            Dim RSRV_QTY As Int64 = .Item("RSRV_QTY")
            Dim RSRV_QTY_OPEN As Int64 = Val(.Item("RSRV_QTY_OPEN") & "")
            Dim RSRV_QTY_CANC As Int64 = Val(.Item("RSRV_QTY_CANC") & "")
            Dim RSRV_QTY_USED As Int64 = Val(.Item("RSRV_QTY_USED") & "") _
                          + S * Val(rowSOTORDR2.Item("ORDR_QTY") & "")

            '  + S * Val(rowSOTORDR2.Item("ORDR_QTY_ORIG") & "") - USING ORDR_QTY_ORIG WILL ALWAYS HAVE 0 IMPACT WHEN CHANGING THE ORDER
            Dim RSRV_QTY_OPEN_OLD As Int64 = RSRV_QTY_OPEN
            RSRV_QTY_OPEN = RSRV_QTY - RSRV_QTY_CANC - RSRV_QTY_USED
            If RSRV_QTY_OPEN <0 Then
                RSRV_QTY_OPEN= 0
            End If
            Dim RSRV_QTY_OPEN_NEW As Int64 = RSRV_QTY_OPEN
            .Item("RSRV_QTY_USED") = RSRV_QTY_USED
            .Item("RSRV_QTY_OPEN") = RSRV_QTY_OPEN

            Dim QTY_TO_COMMIT As Int64 = RSRV_QTY_OPEN_NEW - RSRV_QTY_OPEN_OLD
            If QTY_TO_COMMIT <> 0 Then
                Dim STYLE_CODE As String = .Item("STYLE_CODE")
                Dim COLOR_CODE As String = .Item("COLOR_CODE")
                If S = -1 Then
                    TAC.ICCMAIN1.Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, "WHSE_QTY_OPEN", QTY_TO_COMMIT)
                Else
                    Dim rowICTSTATO As DataRow = dst.Tables("ICTSTATO").Rows.Find(New String() {STYLE_CODE, COLOR_CODE, WHSE_CODE})
                    If rowICTSTATO Is Nothing Then
                        rowICTSTATO = dst.Tables("ICTSTATO").Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, 0})
                    End If
                    rowICTSTATO.Item("WHSE_QTY_OPEN") = Val(rowICTSTATO.Item("WHSE_QTY_OPEN") & "") + QTY_TO_COMMIT
                End If
            End If

        End With

        Dim RSRV_QTY_OPEN_total As Int64 = 0
        If S = -1 Then
            Update_Record_TDA("SOTRSRV2")

            ASCMAIN1.sql = "Select Sum (RSRV_QTY_OPEN) from SOTRSRV2 where RSRV_NO = :PARM1"
            RSRV_QTY_OPEN_total = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {RSRV_NO}))
        Else
            RSRV_QTY_OPEN_total = Val(dst.Tables("SOTRSRV2").Compute("sum(RSRV_QTY_OPEN)", "RSRV_NO = '" & RSRV_NO & "'") & "")
        End If

        If RSRV_QTY_OPEN_total = 0 Then
            rowSOTRSRV1.Item("RSRV_STATUS") = "F"
        Else
            rowSOTRSRV1.Item("RSRV_STATUS") = "O"
        End If

        If S = -1 Then
            Update_Record_TDA("SOTRSRV1")
        End If

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYLX, "B", "Style Status Inquiry")
        Load_Popup_Menu(grdSOTORDR0, "SBBBB", "Show Filter", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdSOTCANCX, "SSB", "Show Filter", "Show Pins", "Sales Order Inquiry")
        Load_Popup_Menu(grdSOTCANCS, "BBB", "Clear Qtys", "Restore Qtys", "Sales Order Inquiry")
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

            Case "grdICTSTYLX"

            Case "grdSOTORDR0"


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else

            Select Case e.SourceControl.Name
                Case "grdSOTALLOX", "grdICTITEM1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Me.Cursor = Cursors.WaitCursor
        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.Key <> "Show All Levels" Then grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next
            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                Next
            Case "Clear Qtys"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = grd.ActiveRow.Cells("COLOR_CODE").Value

                'grdSOTCANCY.Visible = False
                'grdICTSTYLX.Visible = False
                'dst.EnforceConstraints = False

                ASCMAIN1.Progress("Now Suspending Control Totals")
                Dim Rs As New Dictionary(Of String, String)
                For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                    With dst.Tables("ICTSTYLX").Columns(COLUMN_NAME)
                        Rs.Add(COLUMN_NAME, .Expression)
                        .Expression = ""
                    End With
                Next

                ASCMAIN1.Progress("Now Cancelling")

                'For Each row As DataRow In dst.Tables("SOTCANCY").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                '    Dim ORDR_QTY As Int64 = Val(row.Item("ORDR_QTY") & "")
                '    Dim ORDR_QTY_OPEN As Int64 = Val(row.Item("ORDR_QTY_OPEN") & "")
                '    ORDR_QTY_OPEN = 0
                '    Dim ORDR_QTY_PICK As Int64 = Val(row.Item("ORDR_QTY_PICK") & "")
                '    Dim ORDR_QTY_SHIP As Int64 = Val(row.Item("ORDR_QTY_SHIP") & "")
                '    Dim ORDR_QTY_CANC As Int64 = ORDR_QTY - (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP)
                '    row.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                '    row.Item("ORDR_QTY_CANC") = ORDR_QTY_CANC

                'Next

                'grdSOTCANCY.Visible = True
                'grdICTSTYLX.Visible = True
                'dst.EnforceConstraints = True

                For Each grow As UltraWinGrid.UltraGridRow In grdSOTCANCS.Rows
                    'grdSOTCANCS.ActiveCell = grow.Cells("ORDR_QTY_OPEN")
                    grow.Cells("ORDR_QTY_OPEN").Value = 0
                    grow.Update()
                Next

                ASCMAIN1.Progress("Now Restoring Control Totals")
                For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                    With dst.Tables("ICTSTYLX").Columns(COLUMN_NAME)
                        .Expression = Rs(COLUMN_NAME)
                    End With
                Next

                grdICTSTYLX.Rows.Refresh(RefreshRow.ReloadData) ' (RefreshRow.RefreshDisplay)

                ASCMAIN1.Progress("")

            Case "Restore Qtys"
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTCANCS.Rows
                    'grdSOTCANCS.ActiveCell = grow.Cells("ORDR_QTY_OPEN")
                    grow.Cells("ORDR_QTY_OPEN").Value += grow.Cells("ORDR_QTY_CANC").Value
                    grow.Update()
                Next

        End Select
        Me.Cursor = Cursors.Default

        If grd Is Nothing OrElse (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow) Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = Lookup("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If



        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode Then
                    If e.KeyCode = System.Windows.Forms.Keys.Enter Then
                        Set_SOTORDR0()
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)

            Case "CUST_CODE"
                Set_SOTORDR0()

        End Select
    End Sub
#End Region

#Region "grdSOTCANCX"
    Private Sub grdSOTCANCX_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCANCX.AfterCellUpdate
        Select Case e.Cell.Column.Key

        End Select
    End Sub

    Private Sub grdSOTCANCX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTCANCX.AfterRowActivate

    End Sub

    Private Sub grdSOTCANCX_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSOTCANCX.AfterRowUpdate
        If grdSOTCANCX.Tag & "" = "S" Then
        Else

        End If

        Update_Totals()
    End Sub

    Private Sub grdSOTCANCX_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTCANCX.BeforeExitEditMode
        If grdSOTCANCX.ActiveCell IsNot Nothing Then
            With grdSOTCANCX.ActiveCell
                Select Case .Column.Key
                    Case "CUST_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTCANCX_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCANCX.BeforeRowUpdate

        ' Validate_Columns("CUST_CODE", e.Cancel)
        'If Not e.Cancel Then
        '    Validate_Columns("ORDR_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        'End If

        If e.Cancel = True Then
            Exit Sub
        End If

        ' STYLE_CODE_last_entry = e.Row.Cells("STYLE_CODE").Value & ""

        'If e.Row.IsAddRow Then
        '    e.Row.Cells("ORDR_NO").Value = ORDR_NO

        'End If
    End Sub

    Private Sub grdSOTCANCX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCANCX.ClickCellButton

        Dim COLUMN_NAME As String = e.Cell.Column.Key

        Select Case COLUMN_NAME
            Case "CUST_CODE"

                Dim sql_where As String = ""
                grdClickCellButton(grdSOTCANCX, sql_where)

        End Select
    End Sub

#End Region

#Region "grdSOTALLOX"

    Private Sub grdSOTALLOX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs)

    End Sub

    Private Sub grdSOTALLOX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        With e.Row.Cells("SELECTED")
            If .Value & "" = "1" Then
                .Appearance.BackColor = System.Drawing.Color.Red
            Else
                .Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub

#End Region

#Region "grdSOTORDR0"
    Private Sub grdSOTORDR0_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR0.InitializeRow
        With e.Row.Cells("SELECTED")
            If .Value & "" = "1" Then
                .Appearance.BackColor = System.Drawing.Color.LightGreen
            Else
                .Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub
#End Region
    Private Sub Set_SOTCANCS()
        If Not grdICTSTYLX.ActiveRow Is Nothing Then
            Dim RANGE_STYLE_CODE As String = grdICTSTYLX.ActiveRow.Cells("RANGE_STYLE_CODE").Value
            'Dim ORDR_NO As String = grdICTSTYLX.ActiveRow.Cells("ORDR_NO").Value
            'Highlight_Style(RANGE_STYLE_CODE, ORDR_NO)

            Dim SQLW As String = "RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "' "
            If optStatus.Value <> "ALL" Then
                SQLW = SQLW & " and ORDR_STATUS = 'O'"
            End If
            Dim dvw As DataView = DirectCast(grdSOTCANCS.DataSource, DataTable).DefaultView
            dvw.RowFilter = SQLW
            Sort_grdColumns(grdSOTCANCS, "CUST_STORE_NO")

            grdSOTCANCS.Text = "Order Qtys by Store for " & RANGE_STYLE_CODE
        End If

    End Sub

    Private Sub Set_SOTCANCY()
        If Not grdSOTCANCS.ActiveRow Is Nothing Then
            Dim RANGE_STYLE_CODE As String = grdSOTCANCS.ActiveRow.Cells("RANGE_STYLE_CODE").Value
            Dim ORDR_NO As String = grdSOTCANCS.ActiveRow.Cells("ORDR_NO").Value
            'Highlight_Style(RANGE_STYLE_CODE)

            Dim SQLW As String = "RANGE_STYLE_CODE = '" & RANGE_STYLE_CODE & "' and ORDR_NO = '" & ORDR_NO & "'"
            If optStatus.Value <> "ALL" Then
                SQLW = SQLW & " and ORDR_STATUS = 'O'"
            End If
            Dim dvw As DataView = DirectCast(grdSOTCANCY.DataSource, DataTable).DefaultView
            dvw.RowFilter = SQLW
            Sort_grdColumns(grdSOTCANCY, "CUST_STORE_NO")

            grdSOTCANCY.Text = "Range Styles for " & RANGE_STYLE_CODE
        End If

    End Sub

#Region "grdICTSTYLX"

    Private Sub grdICTSTYLX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTYLX.AfterRowActivate
        Set_SOTCANCS()
    End Sub

    Private Sub grdICTSTYLX_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTSTYLX.AfterRowUpdate

        'Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
        'Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
        'Dim iCol As Integer = iSC(STYLE_CODE & "-" & COLOR_CODE)

        'Dim SELECTED As String = e.Row.Cells("SELECTED").Value & ""
        'grdSOTCANCX.DisplayLayout.Bands(0).Columns("Q" & Format(iCol, "00")).Hidden = (SELECTED <> "1")

        'grdSOTCANCX.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
    End Sub

    Private Sub grdICTSTYLX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYLX.InitializeRow

    End Sub

    Private Sub grdICTSTYLX_KeyDown(sender As Object, e As KeyEventArgs) Handles grdICTSTYLX.KeyDown

    End Sub
#End Region

#Region "grdSOTCANCS"
    Private Sub grdSOTCANCS_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCANCS.AfterRowActivate
        Set_SOTCANCY()
    End Sub

    Private Sub grdSOTCANCS_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdSOTCANCS.BeforeCellUpdate
        If e.Cell.Column.Key = "RANGE_STYLE_PP_QTY" And Not grdSOTCANCS.ActiveCell Is Nothing Then
            Dim ORDR_QTY As Int64 = Val(grdSOTCANCS.ActiveCell.Row.Cells("ORDR_QTY").Value & "")
            Dim ORDR_QTY_OPEN As Int64 = Val(grdSOTCANCS.ActiveCell.Row.Cells("ORDR_QTY_OPEN").Value & "")
            Dim RANGE_STYLE_PP_QTY As Int64 = Val(e.NewValue & "") 'Val(grdSOTCANCS.ActiveCell.Row.Cells("RANGE_STYLE_PP_QTY").Value & "")
            Dim RANGE_STYLE_QTY_PER_PP As Int64 = Val(grdSOTCANCS.ActiveCell.Row.Cells("RANGE_STYLE_QTY_PER_PP").Value & "")
            Dim ORDR_QTY_PICK As Int64 = Val(grdSOTCANCS.ActiveCell.Row.Cells("ORDR_QTY_PICK").Value & "")
            Dim ORDR_QTY_SHIP As Int64 = Val(grdSOTCANCS.ActiveCell.Row.Cells("ORDR_QTY_SHIP").Value & "")

            ORDR_QTY_OPEN = RANGE_STYLE_QTY_PER_PP * RANGE_STYLE_PP_QTY

            If ORDR_QTY < (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP) Then
                If MsgBox("Are you sure you want to increase Order Qty?",
                            MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
                    e.Cancel = True
                    Exit Sub
                End If
                'grdSOTCANCS.ActiveCell.Row.Cells("ORDR_QTY").Value = (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP)
            End If

            If ORDR_QTY_OPEN < 0 Then
                MsgBox("Invalid Qty")
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub grdSOTCANCS_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdSOTCANCS.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "RANGE_STYLE_PP_QTY"
                Dim ORDR_QTY As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY").Value)
                Dim ORDR_QTY_PICK As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_PICK").Value)
                Dim ORDR_QTY_SHIP As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_SHIP").Value)
                Dim RANGE_STYLE_PP_QTY As Int64 = Val(e.Cell.Row.Cells("RANGE_STYLE_PP_QTY").Value & "")
                Dim RANGE_STYLE_QTY_PER_PP As Int64 = Val(e.Cell.Row.Cells("RANGE_STYLE_QTY_PER_PP").Value & "")
                Dim RANGE_STYLE_TTL As Int64 = Val(e.Cell.Row.Cells("RANGE_STYLE_TTL").Value & "")
                Dim ORDR_QTY_OPEN As Int64 = RANGE_STYLE_QTY_PER_PP * RANGE_STYLE_PP_QTY
                Dim ORDR_QTY_CANC As Int64 = ORDR_QTY - (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP)
                e.Cell.Row.Cells("ORDR_QTY_OPEN").Value = ORDR_QTY_OPEN
                If ORDR_QTY_CANC > 0 Then
                    e.Cell.Row.Cells("ORDR_QTY_CANC").Value = ORDR_QTY_CANC
                Else
                    e.Cell.Row.Cells("ORDR_QTY_CANC").Value = 0
                End If
                e.Cell.Row.Cells("ORDR_QTY").Value = (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP)

                Dim ORDR_NO As String = e.Cell.Row.Cells("ORDR_NO").Value & ""
                Dim RANGE_STYLE_LNO As String = e.Cell.Row.Cells("RANGE_STYLE_LNO").Value & ""
                Dim ORDR_DTL As Int64
                Dim OPEN_DTL As Int64
                Dim PICK_DTL As Int64
                Dim SHIP_DTL As Int64

                If RANGE_STYLE_QTY_PER_PP <> 1 Then
                    MsgBox("Please Contact Rick, Need to Test for Ranges")
                End If

                For Each row As DataRow In dst.Tables("SOTCANCY").Select("ORDR_NO = '" & ORDR_NO & "' and RANGE_STYLE_LNO = '" & RANGE_STYLE_LNO & "'")
                    ORDR_DTL = Val(row("ORDR_QTY") & "")
                    If RANGE_STYLE_QTY_PER_PP = 1 Then ' 
                        'assortment
                        OPEN_DTL = RANGE_STYLE_PP_QTY / RANGE_STYLE_TTL * Val(row("QTY_PER_PP") & "")
                    Else
                        'range
                        OPEN_DTL = RANGE_STYLE_PP_QTY * Val(row("QTY_PER_PP") & "")
                    End If
                    PICK_DTL = Val(row("ORDR_QTY_PICK") & "")
                    SHIP_DTL = Val(row("ORDR_QTY_SHIP") & "")
                    row("ORDR_QTY_OPEN") = OPEN_DTL
                    If ORDR_DTL - (OPEN_DTL + PICK_DTL + SHIP_DTL) < 0 Then
                        row("ORDR_QTY") = OPEN_DTL + PICK_DTL + SHIP_DTL
                    Else
                        row("ORDR_QTY_CANC") = ORDR_DTL - (OPEN_DTL + PICK_DTL + SHIP_DTL)
                    End If

                Next

        End Select
    End Sub
#End Region


    Private Sub dteEndDate_ValueChanged(sender As Object, e As EventArgs) Handles dteORDR_DATE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If SOTCANCY = "" Then Exit Sub

        Set_SOTORDR0()

    End Sub

    Sub Set_SOTORDR0()

        If SOTCANCY = "" Then Exit Sub
        If ScreenMode Then Exit Sub

        '    Create_Temp_Tables(False)

        Fill_Records("SOTORDR0", New Object() {Absx1.txtFor("CUST_CODE").Text, dteORDR_DATE.Value, Absx1.txtFor("WHSE_CODE").Text})
        For Each row As DataRow In dst.Tables("SOTORDR0").Select("ORDR_QTY <> (ORDR_QTY_OPEN + ORDR_QTY_CANC) or (ORDR_QTY_OPEN + ORDR_QTY_PICK) <> RANGE_QTY_OPEN")
            row("SELECTED") = "0"
        Next

        grdSOTORDR0.Text = "Order Groups for " & Absx1.txtFor("CUST_CODE").Text & " with Order Date of " & dteORDR_DATE.Value
    End Sub

    Sub Update_Totals()
        'For ictr As Integer = 1 To iColumn
        '    If ALLO_CTL_NOi(ictr) <> "" Then
        '        Dim QTY_ALLO As Int64 = Val(dst.Tables("SOTALLOC").Compute("SUM(ALLO_" & Format(ictr, "00") & ")", "") & "")
        '        Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NOi(ictr))
        '        rowSOTALLO1.Item("QTY_ALLO_TOTAL") = QTY_ALLO
        '    End If
        'Next

    End Sub

    Sub Create_Temp_Tables(initialize As Boolean)

        If initialize Then
            dteORDR_DATE.Value = Now.Date.AddDays(-1)
        End If

        Dim CUST_CODE As String = ""
        Dim WHSE_CODE As String = ""
        Dim ORDR_DATE As String = Format(dteORDR_DATE.Value, "dd-MMM-yyyy")

        sqlORDR_GROUP_NOs = ",''"

        If Not initialize Then
            'CUST_CODE = Absx1.txtFor("CUST_CODE").Text
            'WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
            sqlORDR_GROUP_NOs = ""
            For Each row As DataRow In dst.Tables("SOTORDR0").Select("SELECTED = '1'")
                sqlORDR_GROUP_NOs &= ",'" & row("ORDR_GROUP_NO") & "'"
            Next
        End If

        ASCMAIN1.sql = "" _
            & "Select SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.ORDR_LNO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_STATUS, SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.CUST_UPC" & vbCrLf _
            & ", SOTORDR2.QTY_PER_PP" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_OPEN" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_SHIP" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_OPEN ORIG_QTY_OPEN" & vbCrLf _
            & ", SOTORDR9.RANGE_STYLE_QTY_PER_PP" & vbCrLf _
            & ", SOTORDR9.RANGE_STYLE_PP_QTY" & vbCrLf _
            & " from SOTORDR1,SOTORDR2,SOTORDR9" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   And SOTORDR1.ORDR_NO = SOTORDR9.ORDR_NO" & vbCrLf _
            & "   And SOTORDR2.RANGE_STYLE_LNO = SOTORDR9.RANGE_STYLE_LNO" & vbCrLf _
            & "   And SOTORDR1.ORDR_GROUP_NO in (" & vbCrLf _
            & Mid(sqlORDR_GROUP_NOs, 2) & ")"
        '& "   Select ORDR_GROUP_NO" & vbCrLf _
        '& " from SOTORDR0 where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
        '& "   and WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
        '& "   and ORDR_DATE = '" & ORDR_DATE & "')"

        If initialize Then
            SOTCANCY = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Delete from " & SOTCANCY)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCANCY & " " & ASCMAIN1.sql)
        End If

    End Sub

    Sub Highlight_Style(STYLE_CODE As String)
        If iSC.Count = 0 Then Exit Sub
        If 1 = 1 Then Exit Sub

        For Each STYLE_COLOR As String In iSC.Keys
            Dim i As Integer = iSC(STYLE_COLOR)
            If STYLE_COLOR = STYLE_CODE Then
                grdSOTCANCX.DisplayLayout.Bands(0).Columns("Q" & Format(i, "00")).Header.Appearance.BackColor2 = Drawing.Color.Gold
            Else
                grdSOTCANCX.DisplayLayout.Bands(0).Columns("Q" & Format(i, "00")).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            End If
        Next
    End Sub

    Sub Load_Qtys(Optional initial_load As Boolean = False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading " & optQ.Text)

        Dim Q As String = optQ.Value

        For Each rowSOTCANCX As DataRow In dst.Tables("SOTCANCX").Select("")
            If Not initial_load Then
                For i As Integer = 1 To iSC.Count
                    rowSOTCANCX.Item("Q" & Format(i, "00")) = DBNull.Value
                Next
            End If
            Dim CUST_STORE_NO As String = rowSOTCANCX.Item("CUST_STORE_NO")
            Dim ORDR_NO As String = rowSOTCANCX.Item("ORDR_NO")
            Dim SQLW As String = "CUST_STORE_NO = '" & CUST_STORE_NO & "' and ORDR_NO = '" & ORDR_NO & "'"
            For Each rowSOTCANCY As DataRow In dst.Tables("SOTCANCY").Select(SQLW)
                Dim RANGE_STYLE_CODE As String = rowSOTCANCY.Item("RANGE_STYLE_CODE")
                'Dim COLOR_CODE As String = rowSOTCANCY.Item("COLOR_CODE")
                Dim i As Integer = iSC(RANGE_STYLE_CODE)
                rowSOTCANCX.Item("Q" & Format(i, "00")) = rowSOTCANCY.Item(Q)
            Next
        Next
        flgSOTCANCX = False
        btnShowGrid.Visible = False
        grdSOTCANCX.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        grdSOTCANCX.Text = "Stores x Styles - " & optQ.Text
    End Sub

    Private Sub optQ_ValueChanged(sender As Object, e As EventArgs) Handles optQ.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        Load_Qtys()
    End Sub

    Private Sub tabSOTALLOC_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSOTCANCX.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabSOTCANCX()
    End Sub

    Sub Setup_tabSOTCANCX()
        UltraExplorerBar1.Groups("Display Options").Visible = ScreenMode And (tabSOTCANCX.SelectedTab.Key = "Summary")
        UltraExplorerBar1.Groups("Order Status").Visible = ScreenMode And (tabSOTCANCX.SelectedTab.Key = "Styles")
    End Sub


    'Private Sub grdSOTCANCY_AfterCellUpdate(sender As Object, e As CellEventArgs)
    '    Select Case e.Cell.Column.Key
    '        Case "ORDR_QTY_OPEN"
    '            Dim ORDR_QTY As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY").Value)
    '            Dim ORDR_QTY_OPEN As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_OPEN").Value)
    '            Dim ORDR_QTY_PICK As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_PICK").Value)
    '            Dim ORDR_QTY_SHIP As Int64 = Val(e.Cell.Row.Cells("ORDR_QTY_SHIP").Value)
    '            Dim ORDR_QTY_CANC As Int64 = ORDR_QTY - (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP)
    '            e.Cell.Row.Cells("ORDR_QTY_CANC").Value = ORDR_QTY_CANC

    '            If Not (flgSOTCANCX) Then
    '                flgSOTCANCX = True
    '                btnShowGrid.Visible = True
    '                grdSOTCANCX.Visible = False
    '            End If

    '    End Select
    'End Sub

    'Private Sub grdSOTCANCY_AfterRowActivate(sender As Object, e As EventArgs)
    '    If grdSOTCANCY.ActiveRow.Cells("ORDR_STATUS").Value = "O" Then
    '        grdSOTCANCY.ActiveRow.Cells("ORDR_QTY_OPEN").Column.CellActivation = Activation.AllowEdit
    '    Else
    '        grdSOTCANCY.ActiveRow.Cells("ORDR_QTY_OPEN").Column.CellActivation = Activation.NoEdit
    '    End If
    'End Sub

    'Private Sub grdSOTCANCY_BeforeExitEditMode(sender As Object, e As BeforeExitEditModeEventArgs)
    '    'If grdSOTCANCY.ActiveCell.Column.Key = "ORDR_QTY_OPEN" Then
    '    '    Dim ORDR_QTY As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY").Value & "")
    '    '    Dim ORDR_QTY_OPEN As Int64 = Val(grdSOTCANCY.ActiveCell.Text & "")

    '    '    Dim ORDR_QTY_PICK As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_PICK").Value & "")
    '    '    Dim ORDR_QTY_SHIP As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_SHIP").Value & "")
    '    '    'Dim ORDR_QTY_CANC As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_CANC").Value & "")

    '    '    If ORDR_QTY < (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP) Then
    '    '        e.Cancel = True
    '    '    End If
    '    'End If
    'End Sub

    'Private Sub grdSOTCANCY_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs)
    '    If e.Cell.Column.Key = "ORDR_QTY_OPEN" And Not grdSOTCANCY.ActiveCell Is Nothing Then
    '        Dim ORDR_QTY As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY").Value & "")
    '        Dim ORDR_QTY_OPEN As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_OPEN").Value & "") ' Val(grdSOTCANCY.ActiveCell.Text & "")
    '        'Dim STYLE_CODE As String = grdSOTCANCY.ActiveCell.Row.Cells("STYLE_CODE").Value
    '        'Dim COLOR_CODE As String = grdSOTCANCY.ActiveCell.Row.Cells("COLOR_CODE").Value
    '        Dim ORDR_QTY_PICK As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_PICK").Value & "")
    '        Dim ORDR_QTY_SHIP As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_SHIP").Value & "")
    '        'Dim ORDR_QTY_CANC As Int64 = Val(grdSOTCANCY.ActiveCell.Row.Cells("ORDR_QTY_CANC").Value & "")


    '        If ORDR_QTY < (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP) Or ORDR_QTY_OPEN < 0 Then
    '            MsgBox("Invalid Qty")
    '            e.Cancel = True
    '        End If
    '    End If
    'End Sub

    Private Sub btnShowGrid_Click(sender As Object, e As EventArgs) Handles btnShowGrid.Click
        Load_Qtys()
    End Sub

    Private Sub grdSOTORDR0_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdSOTORDR0.BeforeCellUpdate
        If e.Cell.Column.Key = "SELECTED" And Not grdSOTORDR0.ActiveCell Is Nothing Then
            If Val(grdSOTORDR0.ActiveCell.Row.Cells("ORDR_QTY").Value) <>
                Math.Abs(Val(grdSOTORDR0.ActiveCell.Row.Cells("ORDR_QTY_OPEN").Value) + Val(grdSOTORDR0.ActiveCell.Row.Cells("ORDR_QTY_CANC").Value)) Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub txtCUST_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtCUST_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If SOTCANCY = "" Then Exit Sub

        Set_SOTORDR0()
    End Sub

    Private Sub txtWHSE_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtWHSE_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If SOTCANCY = "" Then Exit Sub

        Set_SOTORDR0()
    End Sub

    Private Sub optStatus_ValueChanged(sender As Object, e As EventArgs) Handles optStatus.ValueChanged
        Set_SOTCANCY()
    End Sub
End Class