Imports Infragistics.Win.UltraWinGrid

Public Class ARTCRES2
    Private SQL As New Text.StringBuilder With {.Length = 0}
    Private SQL_ARTCRESX As String = ""
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Get_PARM("ARTPARM1")

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
            SQL.AppendLine("SELECT * FROM ARTCRESH WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTCRESH", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT * FROM ARTCRESD WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTCRESD", "**", 0, True, "V")
            .Tables("ARTCRESD").Columns.Add("DEDUCTION_ACT", GetType(System.Decimal))
            .Tables("ARTCRESD").Columns.Add("DEDUCTION_VAR", GetType(System.Decimal))

            SQL.Length = 0
            SQL.AppendLine("SELECT * FROM ARTCRESP WHERE CUST_CODE = :PARM1")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTCRESP", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("I1.PROGRAM,")
            SQL.AppendLine("I1.PROGRAM_SUB,")
            SQL.AppendLine("I1.CUST_STYLE_CODE,")
            SQL.AppendLine("I1.STYLE_CODE,")
            SQL.AppendLine("I1.COLOR_CODE,")
            SQL.AppendLine("SUM(NVL(S2.ORDR_QTY_SHIP,0)) AS ORDR_QTY_SHIP,")
            SQL.AppendLine("SUM(NVL(S2.ORDR_QTY_SHIP,0) * NVL(S2.ORDR_UNIT_PRICE,0)) AS ORDR_DOL_SHIP")
            SQL.AppendLine("FROM ARTCRESI I1, SOTINVH2 S2")
            SQL.AppendLine("WHERE I1.STYLE_CODE = S2.STYLE_CODE (+)")
            SQL.AppendLine("AND I1.COLOR_CODE = S2.COLOR_CODE (+)")
            SQL.AppendLine("GROUP BY")
            SQL.AppendLine("I1.CUST_CODE,")
            SQL.AppendLine("I1.PROGRAM,")
            SQL.AppendLine("I1.PROGRAM_SUB,")
            SQL.AppendLine("I1.CUST_STYLE_CODE,")
            SQL.AppendLine("I1.STYLE_CODE,")
            SQL.AppendLine("I1.COLOR_CODE")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTCRESI", "**", 0, True, "V")

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("P5.REASON_CODE,")
            SQL.AppendLine("P1.OPS_YYYYPP,")
            SQL.AppendLine("G2.LEGEND,")
            SQL.AppendLine("P5.PYMT_BATCH_NO,")
            SQL.AppendLine("P5.PYMT_BATCH_LNO,")
            SQL.AppendLine("P5.PYMT_BATCH_DLNO,")
            SQL.AppendLine("P5.CUST_REFERENCE,")
            SQL.AppendLine("P1.PYMT_BATCH_DATE,")
            SQL.AppendLine("P5.OUR_REFERENCE,")
            SQL.AppendLine("P5.GL_DIST_AMT")
            SQL.AppendLine("FROM ARTPYMT1 P1, ARTPYMT2 P2, ARTPYMT5 P5, ARTREAS1 R1, GLTPARM2 G2")
            SQL.AppendLine("WHERE NVL(P5.CHARGEBACK_IND,'0') <> '1'")
            SQL.AppendLine("AND P5.PYMT_BATCH_NO = P1.PYMT_BATCH_NO")
            SQL.AppendLine("AND P5.PYMT_BATCH_NO = P2.PYMT_BATCH_NO")
            SQL.AppendLine("AND P5.PYMT_BATCH_LNO = P2.PYMT_BATCH_LNO")
            SQL.AppendLine("AND P5.REASON_CODE = R1.REASON_CODE")
            SQL.AppendLine("AND P1.OPS_YYYYPP = G2.OPS_YYYYPP")
            SQL.AppendLine("AND DECODE(P5.CUST_CODE_SO, NULL, P2.CUST_CODE, P5.CUST_CODE_SO) = :PARM1")
            SQL.AppendLine("AND P5.PYMT_BATCH_NO = :PARM2")
            SQL.AppendLine("AND P5.PYMT_BATCH_LNO = :PARM3")
            SQL.AppendLine("AND P5.PYMT_BATCH_DLNO = :PARM4")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTPYMTD", "**", 0, False, "VVII")
            Create_TDA(.Tables.Add, "ARTPYMTD_P", "**", 0, False, "VVII")
            .Tables("ARTPYMTD_P").Columns.Add("PROGRAM", GetType(System.String))
            .Tables("ARTPYMTD_P").Columns.Add("PROGRAM_SUB", GetType(System.String))
            .Tables("ARTPYMTD_P").Columns.Add("DEDUCTION_TYPE", GetType(System.String))

            SQL.Length = 0
            SQL.AppendLine("SELECT")
            SQL.AppendLine("P1.OPS_YYYYPP,")
            SQL.AppendLine("G2.LEGEND,")
            SQL.AppendLine("P5.PYMT_BATCH_NO,")
            SQL.AppendLine("P5.PYMT_BATCH_LNO,")
            SQL.AppendLine("P5.PYMT_BATCH_DLNO,")
            SQL.AppendLine("P5.REASON_CODE,")
            SQL.AppendLine("P5.OUR_REFERENCE,")
            SQL.AppendLine("P5.CUST_REFERENCE,")
            SQL.AppendLine("P5.GL_DIST_COMMENT,")
            SQL.AppendLine("P5.GL_DIST_AMT")
            SQL.AppendLine("FROM ARTPYMT1 P1, ARTPYMT2 P2, ARTPYMT5 P5, ARTREAS1 R1, GLTPARM2 G2")
            SQL.AppendLine("WHERE NVL(P5.CHARGEBACK_IND,'0') <> '1'")
            SQL.AppendLine("AND P5.PYMT_BATCH_NO = P1.PYMT_BATCH_NO")
            SQL.AppendLine("AND P5.PYMT_BATCH_NO = P2.PYMT_BATCH_NO")
            SQL.AppendLine("AND P5.PYMT_BATCH_LNO = P2.PYMT_BATCH_LNO")
            SQL.AppendLine("AND P5.REASON_CODE = R1.REASON_CODE")
            SQL.AppendLine("AND P1.OPS_YYYYPP = G2.OPS_YYYYPP")
            SQL.AppendLine("AND (P5.PYMT_BATCH_NO, P5.PYMT_BATCH_LNO, P5.PYMT_BATCH_DLNO) NOT IN")
            SQL.AppendLine("(")
            SQL.AppendLine("   SELECT")
            SQL.AppendLine("   PYMT_BATCH_NO,")
            SQL.AppendLine("   PYMT_BATCH_LNO,")
            SQL.AppendLine("   PYMT_BATCH_DLNO")
            SQL.AppendLine("   FROM ARTCRESP")
            SQL.AppendLine("   WHERE CUST_CODE ='NULL'")
            SQL.AppendLine(")")
            SQL.AppendLine("AND DECODE(P5.CUST_CODE_SO, NULL, P2.CUST_CODE, P5.CUST_CODE_SO) = 'NULL'")
            SQL.AppendLine("AND P5.REASON_CODE IN (SELECT REASON_CODE FROM ARTCRES2 WHERE CUST_CODE = 'NULL')")
            SQL.AppendLine("ORDER BY")
            SQL.AppendLine("P1.OPS_YYYYPP,")
            SQL.AppendLine("G2.LEGEND,")
            SQL.AppendLine("P5.PYMT_BATCH_NO,")
            SQL.AppendLine("P5.PYMT_BATCH_LNO,")
            SQL.AppendLine("P5.PYMT_BATCH_DLNO,")
            SQL.AppendLine("P5.REASON_CODE,")
            SQL.AppendLine("P5.OUR_REFERENCE,")
            SQL.AppendLine("P5.CUST_REFERENCE,")
            SQL.AppendLine("P5.GL_DIST_COMMENT,")
            SQL.AppendLine("P5.GL_DIST_AMT")
            ASCMAIN1.sql = SQL.ToString
            Create_TDA(.Tables.Add, "ARTMATCH", "**", 0, False, "VVII")
        End With

        grdARTCRESH.DataSource = dst.Tables("ARTCRESH")
        grdARTCRESD.DataSource = dst.Tables("ARTCRESD")
        grdARTPYMTD.DataSource = dst.Tables("ARTPYMTD")
        grdARTCRESI.DataSource = dst.Tables("ARTCRESI")

        Sort_grdColumns(grdARTPYMTD, "REASON_CODE, OPS_YYYYPP, LEGEND, PYMT_BATCH_NO, PYMT_BATCH_DATE, CUST_REFERENCE", True)
        Sort_grdColumns(grdARTCRESH, "PROGRAM_START, PROGRAM", True)
        Sort_grdColumns(grdARTCRESD, "PROGRAM, PROGRAM_SUB", True)
        Sort_grdColumns(grdARTCRESI, "CUST_STYLE_CODE", True)

        Create_Summary(grdARTPYMTD, "GL_DIST_AMT")

        Create_Summary(grdARTCRESD, "DEDUCTION_AMT")
        Create_Summary(grdARTCRESD, "DEDUCTION_ACT")
        Create_Summary(grdARTCRESD, "DEDUCTION_VAR")

        Create_Summary(grdARTCRESI, "ORDR_QTY_SHIP",,, "###,###,###,###,##0")
        Create_Summary(grdARTCRESI, "ORDR_DOL_SHIP")

        Add_Attachment_Column(grdARTCRESH, 1, "Y", "ARTCRESH", "ATTACH_KEY")

        With grdARTCRESI.DisplayLayout.Bands(0)
            .Columns("ORDR_QTY_SHIP").Format = "###,###,###,###,##0"
            .Columns("ORDR_DOL_SHIP").Format = "###,###,###,###,##0.00"
        End With

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


#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCRESH, "SSB", "Show Filter", "Show GroupBox", "Print Program")
        Load_Popup_Menu(grdARTCRESD, "SSB", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdARTCRESI, "SSB", "Show Filter", "Show GroupBox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else

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
    End Sub
#End Region
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
        Update_Record_TDA("ARTCRESH")
        Update_Record_TDA("ARTCRESD")
        Update_Record_TDA("ARTCRESI")
        Update_Record_TDA("ARTCRESP")
    End Sub

    Private Sub Update_ARTCRES3()
        For Each rowARTCRESX As DataRow In dst.Tables("ARTCRESX").Select()
            Dim CUST_CODE As String = rowARTCRESX.Item("CUST_CODE").ToString & String.Empty
            Dim OPS_YYYYPP As String = rowARTCRESX.Item("OPS_YYYYPP").ToString & String.Empty
            Dim REASON_CODE As String = rowARTCRESX.Item("REASON_CODE").ToString & String.Empty
            Dim TOT_DED As Decimal = Val(rowARTCRESX.Item("TOT_DED_EST").ToString & String.Empty)
            Dim TOT_PCT As Decimal = Val(rowARTCRESX.Item("TOT_DED_PCT").ToString & String.Empty)
            Dim fltARTCRES3 As String = $"CUST_CODE = '{CUST_CODE}' AND OPS_YYYYPP = '{OPS_YYYYPP}' AND REASON_CODE = '{REASON_CODE}'"
            Dim rowARTCRES3 As DataRow = dst.Tables.Item("ARTCRES3").Select(fltARTCRES3).FirstOrDefault
            If IsNothing(rowARTCRES3) Then
                rowARTCRES3 = dst.Tables.Item("ARTCRES3").NewRow
                rowARTCRES3.Item("CUST_CODE") = CUST_CODE
                rowARTCRES3.Item("OPS_YYYYPP") = OPS_YYYYPP
                rowARTCRES3.Item("REASON_CODE") = REASON_CODE
                rowARTCRES3.Item("TOT_DED") = TOT_DED
                rowARTCRES3.Item("PCT_DED") = TOT_PCT
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
        Fill_Records("ARTCRESH", New String() {CUST_CODE})
        Fill_Records("ARTCRESD", New String() {CUST_CODE})
        Fill_Records("ARTCRESI", New String() {CUST_CODE})
        Fill_Records("ARTCRESP", New String() {CUST_CODE})

        FILL_ARTCRESX(CUST_CODE)

        FILL_MATCHED(CUST_CODE)

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

    Private Sub FILL_MATCHED(ByVal CUST_CODE As String)
        dst.Tables("ARTMATCH").Clear()

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine("P1.OPS_YYYYPP,")
        SQL.AppendLine("G2.LEGEND,")
        SQL.AppendLine("P5.PYMT_BATCH_NO,")
        SQL.AppendLine("P5.PYMT_BATCH_LNO,")
        SQL.AppendLine("P5.PYMT_BATCH_DLNO,")
        SQL.AppendLine("P5.REASON_CODE,")
        SQL.AppendLine("P5.OUR_REFERENCE,")
        SQL.AppendLine("P5.CUST_REFERENCE,")
        SQL.AppendLine("P5.GL_DIST_COMMENT,")
        SQL.AppendLine("P5.GL_DIST_AMT")
        SQL.AppendLine("FROM ARTPYMT1 P1, ARTPYMT2 P2, ARTPYMT5 P5, ARTREAS1 R1, GLTPARM2 G2")
        SQL.AppendLine("WHERE NVL(P5.CHARGEBACK_IND,'0') <> '1'")
        SQL.AppendLine("AND P5.PYMT_BATCH_NO = P1.PYMT_BATCH_NO")
        SQL.AppendLine("AND P5.PYMT_BATCH_NO = P2.PYMT_BATCH_NO")
        SQL.AppendLine("AND P5.PYMT_BATCH_LNO = P2.PYMT_BATCH_LNO")
        SQL.AppendLine("AND P5.REASON_CODE = R1.REASON_CODE")
        SQL.AppendLine("AND P1.OPS_YYYYPP = G2.OPS_YYYYPP")
        SQL.AppendLine("AND (P5.PYMT_BATCH_NO, P5.PYMT_BATCH_LNO, P5.PYMT_BATCH_DLNO) NOT IN")
        SQL.AppendLine("(")
        SQL.AppendLine("   SELECT")
        SQL.AppendLine("   PYMT_BATCH_NO,")
        SQL.AppendLine("   PYMT_BATCH_LNO,")
        SQL.AppendLine("   PYMT_BATCH_DLNO")
        SQL.AppendLine("   FROM ARTCRESP")
        SQL.AppendLine($"   WHERE CUST_CODE ='{CUST_CODE}'")
        SQL.AppendLine(")")
        SQL.AppendLine($"AND DECODE(P5.CUST_CODE_SO, NULL, P2.CUST_CODE, P5.CUST_CODE_SO) = '{CUST_CODE}'")
        SQL.AppendLine($"AND P5.REASON_CODE IN (SELECT REASON_CODE FROM ARTCRES2 WHERE CUST_CODE = '{CUST_CODE}')")
        SQL.AppendLine("ORDER BY")
        SQL.AppendLine("P1.OPS_YYYYPP,")
        SQL.AppendLine("G2.LEGEND,")
        SQL.AppendLine("P5.PYMT_BATCH_NO,")
        SQL.AppendLine("P5.PYMT_BATCH_LNO,")
        SQL.AppendLine("P5.PYMT_BATCH_DLNO,")
        SQL.AppendLine("P5.REASON_CODE,")
        SQL.AppendLine("P5.OUR_REFERENCE,")
        SQL.AppendLine("P5.CUST_REFERENCE,")
        SQL.AppendLine("P5.GL_DIST_COMMENT,")
        SQL.AppendLine("P5.GL_DIST_AMT")
        Fill_Records("ARTMATCH",, True, SQL.ToString)

    End Sub


    Private Sub FILL_ARTCRESX(ByVal CUST_CODE As String)
        Dim tSEL As String = SQL_ARTCRESX
        tSEL = tSEL.Replace(":PARM1", $"'{CUST_CODE}'")
        ASCMAIN1.sql = tSEL
        Dim TABLE_ARTCRESX As String = ASCMAIN1.Temp_Table
        Fill_Records("ARTCRESX", New String() {CUST_CODE}, True, $"SELECT * FROM {TABLE_ARTCRESX}")
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
                    If Val(rowARTCRESX.Item("TOT_DED_PCT").ToString & String.Empty) = 0 Then
                        rowARTCRESX.Item("TOT_DED_PCT") = 0
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
                rowARTCRESX.Item("TOT_DED_PCT") = Val(rowARTCRES3.Item("PCT_DED").ToString & String.Empty)
            End If
        Next


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
            For Each TABLE_NAME As String In New String() {"ARTCRES2", "ARTCRES3", "ARTCRESX", "ARTCRESH", "ARTCRESD", "ARTPYMTX", "ARTPYMTD", "ARTCRESI"}
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


        'For i As Integer = 0 To grdARTCRES2.DisplayLayout.Bands(0).Columns.Count - 1
        '    grdARTCRES2.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        'Next i


        With grdARTPYMTD.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
        End With

        With grdARTCRESH.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
        End With

        With grdARTCRESD.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
        End With

    End Sub

#End Region



    Private Sub CALC_VAR()
        For Each rowARTCRESD As DataRow In dst.Tables("ARTCRESD").Select()
            Dim PROGRAM As String = rowARTCRESD.Item("PROGRAM").ToString & String.Empty
            Dim PROGRAM_SUB As String = rowARTCRESD.Item("PROGRAM_SUB").ToString & String.Empty
            Dim DEDUCTION_TYPE As String = rowARTCRESD.Item("DEDUCTION_TYPE").ToString & String.Empty
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT SUM(P5.GL_DIST_AMT) AS  GL_DIST_AMT")
            SQLS.AppendLine("FROM ARTPYMT5 P5, ARTCRESP PP")
            SQLS.AppendLine("WHERE P5.PYMT_BATCH_NO = PP.PYMT_BATCH_NO")
            SQLS.AppendLine("AND P5.PYMT_BATCH_LNO = PP.PYMT_BATCH_LNO")
            SQLS.AppendLine("AND P5.PYMT_BATCH_DLNO = PP.PYMT_BATCH_DLNO")
            SQLS.AppendLine($"AND PP.PROGRAM = '{PROGRAM}'")
            SQLS.AppendLine($"AND PP.PROGRAM_SUB = '{PROGRAM_SUB}'")
            SQLS.AppendLine($"AND PP.DEDUCTION_TYPE = '{DEDUCTION_TYPE}'")
            ASCMAIN1.sql = SQLS.ToString()
            Dim GL_DIST_AMT As Int64 = Val(ASCDATA1.GetDataValue)
            rowARTCRESD.Item("DEDUCTION_ACT") = GL_DIST_AMT
            rowARTCRESD.Item("DEDUCTION_VAR") = Val(rowARTCRESD.Item("DEDUCTION_AMT").ToString & String.Empty) - Val(rowARTCRESD.Item("DEDUCTION_ACT").ToString & String.Empty)
        Next
    End Sub


    Private Sub grdARTCRESH_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCRESH.AfterRowActivate
        If grdARTCRESH.ActiveRow Is Nothing OrElse (Not grdARTCRESH.ActiveRow.IsDataRow Or grdARTCRESH.ActiveRow.IsAddRow) Then
            'grpSOTORDR3.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdARTCRESD.DataSource, DataTable).DefaultView
            Dim PROGRAM As String = grdARTCRESH.ActiveRow.Cells.Item("PROGRAM").Text & String.Empty
            dvw.RowFilter = $"PROGRAM = '{PROGRAM}'"
            grdARTCRESD.Text = $"Details For Program {PROGRAM}"
        End If
    End Sub

    Private Sub grdARTCRESD_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCRESD.AfterRowActivate
        If IsNothing(grdARTCRESD) Then
            grdARTPYMTD.Text = "Deductions Marked For Selected Program"
        Else
            Dim PROGRAM As String = grdARTCRESD.ActiveRow.Cells.Item("PROGRAM").Text & String.Empty
            Dim PROGRAM_SUB As String = grdARTCRESD.ActiveRow.Cells.Item("PROGRAM_SUB").Text & String.Empty
            Dim DEDUCTION_TYPE As String = grdARTCRESD.ActiveRow.Cells.Item("DEDUCTION_TYPE").Text & String.Empty
            grdARTPYMTD.Text = $"Deductions Marked For {PROGRAM} - {PROGRAM_SUB} - {DEDUCTION_TYPE}"
            Dim FLT As String = $"PROGRAM = '{PROGRAM}' AND PROGRAM_SUB = '{PROGRAM_SUB}' AND DEDUCTION_TYPE = '{DEDUCTION_TYPE}'"
            dst.Tables("ARTPYMTD").Clear()
            For Each rowARTCRESP As DataRow In dst.Tables("ARTCRESP").Select(FLT)
                Dim CUST_CODE As String = rowARTCRESP.Item("CUST_CODE").ToString & String.Empty
                Dim PYMT_BATCH_NO As String = rowARTCRESP.Item("PYMT_BATCH_NO").ToString & String.Empty
                Dim PYMT_BATCH_LNO As Int64 = Val(rowARTCRESP.Item("PYMT_BATCH_LNO").ToString & String.Empty)
                Dim PYMT_BATCH_DLNO As Int64 = Val(rowARTCRESP.Item("PYMT_BATCH_DLNO").ToString & String.Empty)
                Fill_Records("ARTPYMTD", New String() {CUST_CODE, PYMT_BATCH_NO, PYMT_BATCH_LNO, PYMT_BATCH_DLNO}, False)
            Next

            Dim dvw As DataView = DirectCast(grdARTCRESI.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"PROGRAM = '{PROGRAM}' AND PROGRAM_SUB = '{PROGRAM_SUB}'"
            grdARTCRESI.Text = $"Style For Program {PROGRAM} / Sub {PROGRAM_SUB}"

            CALC_VAR()
        End If
    End Sub

    Private Sub grdARTCRESD_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdARTCRESD.BeforeRowUpdate
        If e.Row.Cells("CUST_CODE").Text = "" Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text.ToString
        End If
        If e.Row.Cells("PROGRAM").Text = "" Then
            e.Row.Cells("PROGRAM").Value = grdARTCRESH.ActiveRow.Cells("PROGRAM").Text.ToString & String.Empty
        End If
    End Sub

    Private Sub grdARTCRESH_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdARTCRESH.BeforeRowUpdate
        With grdARTCRESH
            If e.Row.IsAddRow Then
                If .ActiveRow.Cells("ATTACH_KEY").Value & "" = "" Then
                    .ActiveRow.Cells("ATTACH_KEY").Value = ASCMAIN1.Next_Control_No("ARTCRESH.ATTACH_KEY")
                End If
            End If
        End With
    End Sub

    Private Sub grdARTCRESH_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdARTCRESH.InitializeLayout

    End Sub
End Class