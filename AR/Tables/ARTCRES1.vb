Imports Infragistics.Win.UltraWinGrid

Public Class ARTCRES1
    Private SQL As New Text.StringBuilder With {.Length = 0}

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Get_PARM("ARTPARM1")
        Dim BEG_PERIOD As String = "201901"
        Dim END_PERIOD As String = "202912"

        With dst
            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("ARTCRES2.CUST_CODE,")
            SQL.AppendLine("ARTCRES2.REASON_CODE,")
            SQL.AppendLine("ARTREAS1.REASON_DESC")
            SQL.AppendLine("FROM ARTCRES2, ARTREAS1")
            SQL.AppendLine("WHERE ARTCRES2.REASON_CODE = ARTREAS1.REASON_CODE")
            SQL.AppendLine("AND ARTCRES2.CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTCRES2", "**", 0, True, "V", 2)

            SQL.Length = 0
            SQL.AppendLine("SELECT * FROM ARTCRES3 WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTCRES3", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("Select ARTREAS1.* from ARTREAS1")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTREAS1", "*", 0, False)
            Fill_Records("ARTREAS1")

            SQL.Length = 0
            SQL.AppendLine("SELECT * FROM GLTPARM2 WHERE OPS_YYYYPP >= :PARM1 and OPS_YYYYPP <= :PARM2")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "VV")
            Fill_Records("GLTPARM2", New String() {BEG_PERIOD, END_PERIOD})

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("DECODE(P5.CUST_CODE_SO, NULL, P2.CUST_CODE, P5.CUST_CODE_SO) CUST_CODE,")
            SQL.AppendLine("P1.OPS_YYYYPP,")
            SQL.AppendLine("G2.LEGEND,")
            SQL.AppendLine("P5.REASON_CODE,")
            SQL.AppendLine("R1.REASON_DESC,")
            SQL.AppendLine("SUM(NVL(P5.GL_DIST_AMT,0)) AS TOT_DED_ACT")
            SQL.AppendLine("FROM ARTPYMT1 P1, ARTPYMT2 P2, ARTPYMT5 P5, ARTREAS1 R1, GLTPARM2 G2")
            SQL.AppendLine("WHERE NVL(P5.CHARGEBACK_IND,'0') <> '1'")
            SQL.AppendLine("AND P5.PYMT_BATCH_NO = P1.PYMT_BATCH_NO")
            SQL.AppendLine("AND P5.PYMT_BATCH_NO = P2.PYMT_BATCH_NO")
            SQL.AppendLine("AND P5.PYMT_BATCH_LNO = P2.PYMT_BATCH_LNO")
            SQL.AppendLine("AND P5.REASON_CODE = R1.REASON_CODE")
            SQL.AppendLine("AND P1.OPS_YYYYPP = G2.OPS_YYYYPP")
            SQL.AppendLine("AND DECODE(P5.CUST_CODE_SO, NULL, P2.CUST_CODE, P5.CUST_CODE_SO) = :PARM1")
            SQL.AppendLine("AND P5.REASON_CODE IN (SELECT REASON_CODE FROM ARTCRES2 WHERE CUST_CODE = :PARM1)")
            SQL.AppendLine($"AND (P1.OPS_YYYYPP >= '{BEG_PERIOD}' AND P1.OPS_YYYYPP <= '{END_PERIOD}')")
            SQL.AppendLine("GROUP BY DECODE(P5.CUST_CODE_SO, NULL, P2.CUST_CODE, P5.CUST_CODE_SO), P1.OPS_YYYYPP, G2.LEGEND, P5.REASON_CODE, R1.REASON_DESC")
            SQL.AppendLine("ORDER BY P1.OPS_YYYYPP, P5.REASON_CODE")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTCRESX", "**", 0, False, "V")
            .Tables("ARTCRESX").Columns.Add("TOT_DED_EST", GetType(System.Decimal))

        End With

        grdARTCRES2.DataSource = dst.Tables("ARTCRES2")
        grdARTCRESX.DataSource = dst.Tables("ARTCRESX")

        Sort_grdColumns(grdARTCRESX, "OPS_YYYYPP, REASON_CODE", True)

        Create_Summary(grdARTCRESX, "TOT_DED_ACT")
        Create_Summary(grdARTCRESX, "TOT_DED_EST")

        'With grdARTCUST2.DisplayLayout.Bands(0)
        '    '.Columns("CUST_STORE_NO").Header.Fixed = True
        '    '.Columns("CUST_STORE_NAME").Header.Fixed = True
        'End With

        'ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        'Call InitializeControls(Me)
        'ASCMAIN1.Add_Value_List(grdARTCUST2, "CUST_ADDR_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive", "C:Closed"})

        'Set_Read_Only_for_ctl(Absx1.optFor("CUST_SHIP_COMPLETE"), True)
        'Set_Read_Only_for_ctl(Absx1.chkFor("CUST_CONS_INV"), True)
        '    Absx1.chkFor("CUST_SHIP_COMPLETE").Enabled = False
        '    Absx1.chkFor("CUST_CONS_INV").Enabled = False
        '    Absx1.chkFor("CUST_EDI_DTS_FLAG").Enabled = False
    End Sub

#Region "Overrides"

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Stop
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        'Stop
        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

                'If CreditCardQueue1.isInEditMode Then
                '    EMsg = "Update or Cancel Credit Card changes."
                '    Exit Select
                'End If

                'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

        End Select
    End Sub


    Overrides Sub Proceed_Update_Special_Pre()
        'Stop
        'grdARTCUST2.UpdateData()

        Dim sqlDelete = ""

        'For Each rowARTCRES2 As DataRow In dst.Tables("ARTCRES2").Select()
        '    rowARTCRES2.Item("REASON_DESC") = Null
        'Next
        Update_ARTCRES3()
        Update_Record_TDA("ARTCRES2")
        Update_Record_TDA("ARTCRES3")
    End Sub

    Private Sub Update_ARTCRES3()
        For Each rowARTCRESX As DataRow In dst.Tables("ARTCRESX").Select()
            Dim CUST_CODE As String = rowARTCRESX.Item("CUST_CODE").ToString & String.Empty
            Dim OPS_YYYYPP As String = rowARTCRESX.Item("OPS_YYYYPP").ToString & String.Empty
            Dim REASON_CODE As String = rowARTCRESX.Item("REASON_CODE").ToString & String.Empty
            Dim TOT_DED As Decimal = Val(rowARTCRESX.Item("TOT_DED_EST").ToString & String.Empty)
            Dim fltARTCRES3 As String = $"CUST_CODE = '{CUST_CODE}' AND OPS_YYYYPP = '{OPS_YYYYPP}' AND REASON_CODE = '{REASON_CODE}'"
            Dim rowARTCRES3 As DataRow = dst.Tables.Item("ARTCRES3").Select(fltARTCRES3).FirstOrDefault
            If IsNothing(rowARTCRES3) Then
                rowARTCRES3 = dst.Tables.Item("ARTCRES3").NewRow
                rowARTCRES3.Item("CUST_CODE") = CUST_CODE
                rowARTCRES3.Item("OPS_YYYYPP") = OPS_YYYYPP
                rowARTCRES3.Item("REASON_CODE") = REASON_CODE
                rowARTCRES3.Item("TOT_DED") = TOT_DED
                dst.Tables.Item("ARTCRES3").Rows.Add(rowARTCRES3)
            Else
                rowARTCRES3.Item("TOT_DED") = TOT_DED
            End If
        Next
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        'Stop
    End Sub

    Overrides Sub Show_Record_Special()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        If Not IsNothing(rowARTCUST1) Then
            txtCUST_NAME.Text = rowARTCUST1.Item("CUST_NAME").ToString & String.Empty
        End If

        Fill_Records("ARTCRES2", New String() {CUST_CODE})
        'Fill_ResonDesc()
        Fill_Records("ARTCRES3", New String() {CUST_CODE})

        FILL_ARTCRESX(CUST_CODE)

        'With grdARTCUSTD.DisplayLayout.Bands(0)
        '    For Each C As String In New String() {"CONTACT_PHONE", "CONTACT_FAX", "CONTACT_CELL"}
        '        .Columns(C).MaskInput = "" ' "(###) ###-####"
        '        .Columns(C).CellDisplayStyle = UltraWinGrid.CellDisplayStyle.Default ' UltraWinGrid.CellDisplayStyle.FormattedText
        '    Next
        'End With


        'If EntryMode = "New" Then
        '    rowASFBASE1.Item("CUST_CREDIT_LIMIT") = Val(ROWs("ARTPARM1").Item("AR_PARM_INITIAL_CR_LIMIT") & "")
        '    If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
        '        rowASFBASE1.Item("CUST_CRED_LIMIT_EST") = DATETIME_STAMP.Date
        '    End If
        '    rowASFBASE1.Item("CUST_CREDIT_LIMIT_NOTES") = "Initial Credit Limit"
        '    rowASFBASE1.Item("CUST_STMT_IND") = "M"
        '    rowASFBASE1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE")
        '    rowASFBASE1.Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE")
        '    rowASFBASE1.Item("CUST_STATUS") = "A"
        '    rowASFBASE1.Item("WHSE_CODE") = "MS"
        '    rowASFBASE1.Item("CUST_PRICE_TIER") = "PC"
        '    If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
        '        rowASFBASE1.Item("CUST_STATUS_DATE") = Now.Date ' DATETIME_STAMP.Date
        '    End If
        '    rowASFBASE1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

        '    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
        '        rowASFBASE1.Item("CUST_FACTOR_IND") = "1"
        '    End If
        'End If

        'EnforceConstraints(False)
        'Fill_Records("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text})

        'EnforceConstraints(True)
    End Sub

    Private Sub FILL_ARTCRESX(ByVal CUST_CODE As String)
        Fill_Records("ARTCRESX", New String() {CUST_CODE})
        For Each rowGLTPARM2 As DataRow In dst.Tables("GLTPARM2").Select("", "OPS_YYYYPP")
            Dim OPS_YYYYPP As String = rowGLTPARM2.Item("OPS_YYYYPP").ToString & String.Empty
            Dim LEGEND As String = rowGLTPARM2.Item("LEGEND").ToString & String.Empty
            For Each rowARTCRES2 As DataRow In dst.Tables("ARTCRES2").Select("", "REASON_CODE")
                Dim REASON_CODE As String = rowARTCRES2.Item("REASON_CODE").ToString & String.Empty
                Dim REASON_DESC As String = rowARTCRES2.Item("REASON_DESC").ToString & String.Empty
                Dim fltARTCRESX As String = $"CUST_CODE = '{CUST_CODE}' AND OPS_YYYYPP = '{OPS_YYYYPP}' AND REASON_CODE = '{REASON_CODE}'"
                Dim rowARTCRESX As DataRow = dst.Tables.Item("ARTCRESX").Select(fltARTCRESX).FirstOrDefault
                If IsNothing(rowARTCRESX) Then
                    Dim newARTCRESX As DataRow = dst.Tables.Item("ARTCRESX").NewRow
                    newARTCRESX.Item("CUST_CODE") = CUST_CODE
                    newARTCRESX.Item("OPS_YYYYPP") = OPS_YYYYPP
                    newARTCRESX.Item("LEGEND") = LEGEND
                    newARTCRESX.Item("REASON_CODE") = REASON_CODE
                    newARTCRESX.Item("REASON_DESC") = REASON_DESC
                    newARTCRESX.Item("TOT_DED_ACT") = 0
                    newARTCRESX.Item("TOT_DED_EST") = 0
                    dst.Tables.Item("ARTCRESX").Rows.Add(newARTCRESX)
                Else
                    If Val(rowARTCRESX.Item("TOT_DED_EST").ToString & String.Empty) = 0 Then
                        rowARTCRESX.Item("TOT_DED_EST") = 0
                    End If
                End If
            Next
        Next
        For Each rowARTCRES3 As DataRow In dst.Tables("ARTCRES3").Select()
            Dim OPS_YYYYPP As String = rowARTCRES3.Item("OPS_YYYYPP").ToString & String.Empty
            Dim REASON_CODE As String = rowARTCRES3.Item("REASON_CODE").ToString & String.Empty
            Dim fltARTCRESX As String = $"CUST_CODE = '{CUST_CODE}' AND OPS_YYYYPP = '{OPS_YYYYPP}' AND REASON_CODE = '{REASON_CODE}'"
            Dim rowARTCRESX As DataRow = dst.Tables.Item("ARTCRESX").Select(fltARTCRESX).FirstOrDefault
            If Not IsNothing(rowARTCRESX) Then
                rowARTCRESX.Item("TOT_DED_EST") = Val(rowARTCRES3.Item("TOT_DED").ToString & String.Empty)
            End If
        Next

        Sort_grdColumns(grdARTCRESX, "OPS_YYYYPP, REASON_CODE", True)
    End Sub

    Private Sub Fill_ResonDesc()
        For Each rowARTCRES2 As DataRow In dst.Tables("ARTCRES2").Select()
            If rowARTCRES2.Item("REASON_DESC").ToString & String.Empty = "" Then
                rowARTCRES2.Item("REASON_DESC") = getReasonDesc(rowARTCRES2.Item("REASON_CODE").ToString & String.Empty)
            End If
        Next
    End Sub

    Private Function getReasonDesc(ByVal REASON_CODE As String) As String
        Dim RetVal As String = ""
        Dim flt As String = $"REASON_CODE = '{REASON_CODE}'"
        Dim rowARTREAS1 As DataRow = dst.Tables.Item("ARTREAS1").Select(flt).FirstOrDefault
        If Not IsNothing(rowARTREAS1) Then
            RetVal = rowARTREAS1.Item("REASON_DESC").ToString & String.Empty
        End If
        Return RetVal
    End Function

    Overrides Sub Clear_Record_Special()
        'Stop
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ARTCRES2", "ARTCRES3", "ARTCRESX"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
        txtCUST_NAME.Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        'Stop
        'Set_Read_Only_for_ctl(Absx1.txtFor("CUST_NAME"), Not tf)
        'Set_Read_Only(grpCreditLimit, True)
        ' Set_Read_Only(grpOther, True)
        ' Set_Read_Only(grpCreditLimit, IIf(Not tf, ASCMAIN1.USER_SECURITY_CODEs.Contains("CL"), True))
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        With grdARTCRES2.DisplayLayout.Override
            If (EntryMode = "New" Or EntryMode = "Edit") Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With
        'For i As Integer = 0 To grdARTCRES2.DisplayLayout.Bands(0).Columns.Count - 1
        '    grdARTCRES2.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        'Next i

        With grdARTCRES2.DisplayLayout.Bands(0)
            'Dim editColumns As String() = New String() {"XXX"}
            'For Each COLNAME As String In editColumns
            '    .Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            '    .Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            'Next
            'For Each COL_NAME As String In New String() {"EMAIL", "GIVENNAME", "FAMILYNAME", "CLAIM_BY_OPER"}
            '    .Columns(COL_NAME).Header.Fixed = True
            'Next
        End With

        If (EntryMode = "Edit") Then
            With grdARTCRESX.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.False
            End With
            With grdARTCRESX.DisplayLayout.Bands(0)
                Dim editColumns As String() = New String() {"TOT_DED_EST"}
                For Each COLNAME As String In editColumns
                    .Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                Next
            End With
        Else
            With grdARTCRESX.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        End If

    End Sub

#End Region

    Private Sub grdARTCRES2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCRES2.ClickCellButton
        Dim sql_where As String = ""
        Call grdClickCellButton(grdARTCRES2, sql_where, True)
    End Sub

    Private Sub grdARTCRES2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCRES2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "REASON_CODE"
                grdARTCRES2.ActiveRow.Cells("REASON_DESC").Value = getReasonDesc(e.Cell.Text)
        End Select
    End Sub

    Private Sub grdARTCRES2_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdARTCRES2.AfterRowUpdate
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Update_Record_TDA("ARTCRES2")
        Update_ARTCRES3()
        FILL_ARTCRESX(CUST_CODE)
    End Sub
End Class