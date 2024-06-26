Imports Infragistics.Win.UltraWinGrid

Public Class ARFPYMT2

    Dim rowGLTBANK1 As DataRow
    Dim rowGLTPARM2 As DataRow
    Dim rowARTCUST1 As DataRow

    Dim CUST_CODE As String
    Dim CUST_NAME As String
    Dim BANK_CODE As String
    Dim CURR_CODE As String
    Dim TOTAL_UNAPPLIED As Decimal

    Dim TRADE_CLASS_CODE As String
    Dim CHANNEL_CODE As String
    Dim rowSOTTCLS1 As DataRow
    Dim rowSOTCHAN1 As DataRow

    Dim RESPONSE_BATCH_NO As String

    Dim PYMT_BATCH_NO_new As String
    Dim PYMT_BATCH_LNO_new As Int32

    Dim CUST_PYMT_AMT_CURR As Decimal = 0
    Dim PYMT_NOTE As String

    Dim PYMT_BATCH_NO_application_only As String = ""
    Dim PYMT_BATCH_LNO_application_only As Int32
    Dim application_only As Boolean = False
    Dim reverse_application_option As String
    Dim applying_to_statement As Boolean = False
    Dim sqlARTPYMTX As String

    Dim ARTOPENX As String = ""

    Dim PYMT_BATCH_NO As String
    Dim PYMT_BATCH_LNO As Int32

    Dim AGING_DATES_ado(4) As String
    Dim AGING_DATES(4) As Date
    Dim AGED_TOTALS(4) As Decimal
    Dim AGE_DATE_COLUMN As String = "INV_DATE" ' "INV_DUE_DATE"
    Dim DTEs_YP() As Date
    Dim MOVE_PAYMENTS As New List(Of String)
    Dim EDI_DOC_SEQ_NO As String = ""
    Dim rowEDT820T1 As DataRow
    Dim rowARTREAS1 As DataRow
    Dim edi_820_in_process As Boolean

    Dim bcLightGreen As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightGreen}
    Dim bcBeige As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.Beige}
    Dim fcRed As New Infragistics.Win.Appearance With {.ForeColor = Drawing.Color.Red}
    Dim fcEmpty As New Infragistics.Win.Appearance With {.ForeColor = Drawing.Color.Red}
    Dim bcLightBlue As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.LightBlue}
    Dim bcYellow As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.Yellow}
    Dim bcPink As New Infragistics.Win.Appearance With {.BackColor = Drawing.Color.Pink}
    Dim btnSelected As New Infragistics.Win.Appearance
    Dim btnBlank As New Infragistics.Win.Appearance

    Dim EDI_TABLES() As String
    Public Structure strTOTALS
        Public APPL_TOTAL As Decimal
        Public DISC_TOTAL As Decimal
        Public WOFF_TOTAL As Decimal
        Public DED_TOTAL As Decimal
        Public CHB_TOTAL As Decimal
        Public OA_TOTAL As Decimal
        Public GL_TOTAL As Decimal
        Public NET_AR As Decimal
        Public UNAPPLIED As Decimal
    End Structure

    Dim TOTALS As strTOTALS

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        btnSelected.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "Selected")
        btnBlank.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "Blank Selection")

        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("SOTPARM1")

        With dst
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

            Create_TDA(.Tables.Add, "ARTCUST6", "*")

            ASCMAIN1.sql = "SELECT ARTPYMT1.PYMT_BATCH_NO, ARTPYMT1.PYMT_BATCH_DATE" _
            & ", ARTPYMT1.BANK_CODE, ARTPYMT1.PYMT_SOURCE, ARTPYMT1.PYMT_APPL_ONLY" _
            & ", ARTPYMT1.INIT_OPER, ARTPYMT1.INIT_DATE" _
            & ", SUM (DECODE(ARTPYMT2.PYMT_STATUS,'1',1,0)) S1" _
            & ", SUM (DECODE(ARTPYMT2.PYMT_STATUS,'2',1,0)) S2" _
            & ", COUNT (ARTPYMT2.PYMT_BATCH_LNO) RECORDS" _
            & ", SUM (CUST_PYMT_AMT) CUST_PYMT_AMT" _
            & " FROM ARTPYMT2,ARTPYMT1" _
            & " WHERE ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & " AND ARTPYMT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "' AND ARTPYMT1.STATUS <> '2'" _
            & " GROUP BY ARTPYMT1.PYMT_BATCH_NO, ARTPYMT1.PYMT_BATCH_DATE" _
            & ", ARTPYMT1.BANK_CODE, ARTPYMT1.PYMT_SOURCE, ARTPYMT1.PYMT_APPL_ONLY" _
            & ", ARTPYMT1.INIT_OPER, ARTPYMT1.INIT_DATE"
            Create_TDA(.Tables.Add, "ARTPYMTB", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*")

            Create_TDA(.Tables.Add("ARTPYMT2_SPLIT"), "ARTPYMT2", "*")

            '.Tables("ARTPYMT1").Columns.Add("TOTAL_CUST_PYMT_AMT", GetType(System.Decimal), "SUM(CHILD(ARTPYMT1_ARTPYMT2_BOX).CUST_PYMT_AMT_CURR)")
            '.Tables("ARTPYMT1").Columns.Add("OK_TO_APPLY", GetType(System.String), "MIN(CHILD(ARTPYMT1_ARTPYMT2_BOX).OK_TO_APPLY)")
            '.Tables("ARTPYMT1").Columns.Add("PYMT_COUNT", GetType(System.Int32), "COUNT(CHILD(ARTPYMT1_ARTPYMT2_BOX).OK_TO_APPLY)")
            '.Tables("ARTPYMT1").Columns.Add("PYMT_ERRORS", GetType(System.Int32))

            Create_TDA(.Tables.Add, "ARTPYMT3", "*", 2)
            .Tables("ARTPYMT3").Columns.Add("AGE_BUCKET", GetType(System.Int32))
            .Tables("ARTPYMT3").Columns.Add("INV_FREIGHT_CURR", GetType(System.Decimal))
            .Tables("ARTPYMT3").Columns.Add("INV_MISC_CHG_CURR", GetType(System.Decimal))
            .Tables("ARTPYMT3").Columns.Add("INV_TOTAL_AMOUNT_CURR", GetType(System.Decimal))

            ASCMAIN1.sql = "Select ARTPYMT4.*, GLTACCT1.ACCT_DESC" _
            & " from ARTPYMT4,GLTACCT1" _
            & " where GLTACCT1.ACCT_CODE = ARTPYMT4.ACCT_CODE"
            Create_TDA(.Tables.Add, "ARTPYMT4", "**", 2)

            ASCMAIN1.sql = "Select ARTPYMT5.*, ARTREAS1.REASON_DESC" _
            & " from ARTPYMT5,ARTREAS1" _
            & " where ARTREAS1.REASON_CODE (+) = ARTPYMT5.REASON_CODE"
            Create_TDA(.Tables.Add, "ARTPYMT5", "**", 2)
            Dim SQL As String = ""
            SQL &= "IIF(ISNULL(GL_DIST_AMT_CURR,0)>0 And ISNULL(CHARGEBACK_IND,'0')='1','Chargeback',"
            SQL &= "IIF(ISNULL(GL_DIST_AMT_CURR,0)<0 And ISNULL(CHARGEBACK_IND,'0')='1','On Account',"
            SQL &= "IIF(ISNULL(GL_DIST_AMT_CURR,0)>0 And ISNULL(CHARGEBACK_IND,'0')='0','DR (Expense)',"
            SQL &= "IIF(ISNULL(GL_DIST_AMT_CURR,0)<0 And ISNULL(CHARGEBACK_IND,'0')='0','CR (Income)',"
            SQL &= "''))))"
            .Tables("ARTPYMT5").Columns.Add("TRANSACTION_LEGEND", GetType(System.String), SQL)

            sqlARTPYMTX = "Select ARTPYMT2.*" & vbCrLf _
            & ", ARTPYMT1.PYMT_BATCH_DATE, ARTPYMT1.BANK_CODE" & vbCrLf _
            & ", ARTPYMT1.STATUS" & vbCrLf _
            & ", ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_SOURCE" & vbCrLf _
            & ", ARTPYMT1.REGISTER_IND, ARTPYMT1.REGISTER_XNO, ARTPYMT1.REGISTER_DATE" & vbCrLf _
            & " from ARTPYMT1,ARTPYMT2" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO"

            ASCMAIN1.sql = sqlARTPYMTX _
            & "   and ARTPYMT2.PYMT_STATUS = '1' " _
            & "   and ARTPYMT1.INIT_OPER LIKE :PARM1 "

            Create_TDA(.Tables.Add, "ARTPYMTX", "**", 0, False, "V", 2)

            .Tables.Add("ARTPYMTT")
            With .Tables("ARTPYMTT").Columns
                .Add("PYMT_TOTAL_CODE")
                .Add("PYMT_TOTAL_CAPTION")
                .Add("PYMT_TOTAL_AMT", GetType(System.Double))
            End With

            With .Tables("ARTPYMTT")
                .PrimaryKey = New DataColumn() { .Columns("PYMT_TOTAL_CODE")}
                .Rows.Add(New Object() {"1", "Amt Applied", 0})
                .Rows.Add(New Object() {"2", "Discounts", 0})
                .Rows.Add(New Object() {"3", "Write-Off", 0})
                .Rows.Add(New Object() {"4", "Deductions", 0})
                .Rows.Add(New Object() {"5", "ChargeBack", 0})
                .Rows.Add(New Object() {"6", "GL Dist", 0})
                .Rows.Add(New Object() {"7", "On Account", 0})
                .Rows.Add(New Object() {"8", "Net AR", 0})
                .Rows.Add(New Object() {"9", "UnApplied", 0})
            End With

            .Tables.Add("ARTOPENA")
            With .Tables("ARTOPENA").Columns
                .Add("AGE_NO", GetType(System.Int32))
                .Add("AGE_DESC")
                .Add("AGE_AMT", GetType(System.Double))
                .Add("AGE_AMT_NEW", GetType(System.Double))
            End With

            With .Tables("ARTOPENA")
                .PrimaryKey = New DataColumn() { .Columns("AGE_NO")}
                For I As Integer = 1 To 4
                    .Rows.Add(New Object() {I, ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_" & CStr(I)), 0, 0})
                Next
            End With

            .Tables.Add("ARTPYMTA")
            With .Tables("ARTPYMTA").Columns
                .Add("AR_TYPE_CODE")
                .Add("AR_TYPE_SEQ", GetType(System.Int32))
                .Add("AR_TYPE_CAPTION")
                .Add("ITEMS", GetType(System.Int32))
                .Add("AR_TYPE_AMT_OLD", GetType(System.Double))
                .Add("AR_TYPE_AMT", GetType(System.Double))
                .Add("AR_TYPE_AMT_NEW", GetType(System.Double))
            End With
            With .Tables("ARTPYMTA")
                .PrimaryKey = New DataColumn() { .Columns("AR_TYPE_CODE")}
                .Rows.Add(New Object() {"I", 1, "Invoice"})
                .Rows.Add(New Object() {"R", 2, "Returns"})
                .Rows.Add(New Object() {"D", 3, "DR Memo"})
                .Rows.Add(New Object() {"C", 4, "CR Memo"})
                .Rows.Add(New Object() {"B", 5, "ChgBack"})
                .Rows.Add(New Object() {"O", 6, "On Acct"})
            End With

            Create_TDA(.Tables.Add, "SOTTYPE1", "*", 0, False)
            Create_TDA(.Tables.Add, "ARTPOST1", "*", 0, False)
            Create_TDA(.Tables.Add, "ARTREAS1", "*", 0, False)
            Create_TDA(.Tables.Add, "ARTCUST1", "*")

            ASCMAIN1.sql = "Select EDT820T1.* from EDT820T1" & vbCrLf _
                & " where EDT820T1.EDI_PROCESS_IND = '0'" & vbCrLf _
                & "   and EDT820T1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            Create_TDA(.Tables.Add, "EDT820TX", "**", 0, False, "V", 2)
            '.Tables("EDT820TX").Columns.Add("SEL")
            '.Tables("EDT820TX").Columns("SEL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "EDT820T1", "*")
            Create_TDA(.Tables.Add, "EDT820T2", "*", 1, False)
            Create_TDA(.Tables.Add, "EDT820T3", "*", 1, False)
            Create_TDA(.Tables.Add, "EDT820T4", "*", 1, False)
            Create_TDA(.Tables.Add, "EDT820T5", "*", 1, False)

            Create_Relation("EDT820T1", "EDT820T2", "EDI_DOC_SEQ_NO")
            Create_Relation("EDT820T2", "EDT820T3", "EDI_DOC_SEQ_NO,EDI_ENT_NO")
            Create_Relation("EDT820T2", "EDT820T4", "EDI_DOC_SEQ_NO,EDI_ENT_NO")
            Create_Relation("EDT820T3", "EDT820T5", "EDI_DOC_SEQ_NO,EDI_ENT_NO,EDI_INV_SEQ")

            ASCMAIN1.sql = "Select EDTXREF1.*,ARTREAS1.REASON_DESC" & vbCrLf _
                 & " from EDTXREF1,ARTREAS1" & vbCrLf _
                 & " where ARTREAS1.REASON_CODE (+) = EDTXREF1.REASON_CODE" & vbCrLf _
                 & "   And SENDER_ID_QUAL = :PARM1 And SENDER_ID = :PARM2"

            If ASCMAIN1.CLIENT = "INT" Or ASCMAIN1.CLIENT = "AHA" Then
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, " SENDER_ID_QUAL", " EDI_TP_QUAL")
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, " SENDER_ID", " EDI_TP_ID")
            End If
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                ' REASON_DESC NOT NEC FOR NYA
                ASCMAIN1.sql = "Select EDTXREF1.*,ARTREAS1.REASON_DESC REASON_DESC2" & vbCrLf _
                 & " from EDTXREF1,ARTREAS1" & vbCrLf _
                 & " where ARTREAS1.REASON_CODE (+) = EDTXREF1.REASON_CODE" & vbCrLf _
                 & "   And SENDER_ID_QUAL = :PARM1 And SENDER_ID = :PARM2"
            End If

            Create_TDA(.Tables.Add, "EDTXREF1", "**", 0, True, "VV")

            Create_TDA(.Tables.Add, "EDTXREF2", "*", 0, False)

            With .Tables.Add("EDTINVC1")
                .Columns.Add("INV_NUM")
                .Columns.Add("PYMT_BATCH_NO")
                .Columns.Add("PYMT_BATCH_LNO", GetType(System.Int64))
                .Columns.Add("PYMT_BATCH_ILNO", GetType(System.Int64))
                .PrimaryKey = New DataColumn() { .Columns("INV_NUM")}
            End With

            With .Tables.Add("EDTERRS1")
                .Columns.Add("MSG_TYPE")
                .Columns.Add("MSG_KEY")
                .Columns.Add("MSG_TEXT")
                .PrimaryKey = New DataColumn() { .Columns("MSG_TYPE"), .Columns("MSG_KEY")}
            End With

            With .Tables.Add("ARTPYMTL")
                .Columns.Add("INV_NUM")
                .Columns.Add("INV_REF")
                .Columns.Add("INV_PMT", GetType(System.Decimal))
            End With

            With .Tables.Add("ARTPYMTM")
                .Columns.Add("MATCH_REF")
                .Columns.Add("MATCH_RC")
                .Columns.Add("MATCH_RG")
                .Columns.Add("MATCH_ACTION")
                .Columns.Add("MATCH_TOTAL", GetType(System.Decimal))
                .Columns.Add("MATCH_TOTAL_DR", GetType(System.Decimal))
                .Columns.Add("MATCH_TOTAL_CR", GetType(System.Decimal))
                .Columns.Add("INV_COUNT", GetType(System.Int32))
                .PrimaryKey = New DataColumn() { .Columns("MATCH_REF"), .Columns("MATCH_RC"), .Columns("MATCH_RG")}
            End With

            With .Tables.Add("ARTPYMTN")
                .Columns.Add("MATCH_REF")
                .Columns.Add("MATCH_RC")
                .Columns.Add("MATCH_RG")
                .Columns.Add("PYMT_BATCH_ILNO", GetType(System.Int32))
                .Columns.Add("CUST_CODE")
                .Columns.Add("INV_TYPE")
                .Columns.Add("INV_NUM")
                .Columns.Add("INV_CUST_PO")
                .Columns.Add("REASON_CODE")
                .Columns.Add("INV_DATE", GetType(System.DateTime))
                .Columns.Add("INV_BALANCE", GetType(System.Decimal))

                .Columns.Add("MATCH_TOTAL_DR", GetType(System.Decimal), "IIF(INV_BALANCE>0,INV_BALANCE,0)")
                .Columns.Add("MATCH_TOTAL_CR", GetType(System.Decimal), "IIF(INV_BALANCE<0,INV_BALANCE,0)")
                .Columns.Add("REASON_MATCH_GROUP")
                .PrimaryKey = New DataColumn() { .Columns("MATCH_REF"), .Columns("MATCH_RC"), .Columns("MATCH_RG"), .Columns("PYMT_BATCH_ILNO")}
            End With

            Create_Relation("ARTPYMTM", "ARTPYMTN", "MATCH_REF,MATCH_RC,MATCH_RG")

            'With dst.Tables("ARTPYMTM")
            '    .Columns("MATCH_TOTAL_DR").Expression = "SUM(CHILD.MATCH_TOTAL_DR)"
            '    .Columns("MATCH_TOTAL_CR").Expression = "SUM(CHILD.MATCH_TOTAL_CR)"
            '    .Columns("INV_COUNT").Expression = "COUNT(CHILD.INV_NUM)"
            'End With

        End With

        Fill_Records("SOTTYPE1")
        Fill_Records("ARTPOST1")
        Fill_Records("ARTREAS1")

        'Create_Lookup("GLTACCT1")
        'Create_Lookup("GLTBANK1")
        'Create_Lookup("EDTTRPM1")
        'Create_Lookup("ARTCUST1")
        'Create_Lookup("GLTPARM2")

        rowGLTPARM2 = LookUp("GLTPARM2", ASCMAIN1.CYP)

        grdARTPYMTX.DataSource = dst.Tables("ARTPYMTX")
        grdEDTXREF1.DataSource = dst.Tables("EDTXREF1")

        grdARTPYMTM.DataSource = dst.Tables("ARTPYMTM")
        With grdARTPYMTM.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
            .Bands(0).Override.AllowUpdate = DefaultableBoolean.True
            .Bands(1).Override.AllowUpdate = DefaultableBoolean.False
            For Each gcol As UltraWinGrid.UltraGridColumn In .Bands(0).Columns
                If gcol.Key = "MATCH_ACTION" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        grdARTPYMTB.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grdARTPYMTB.DataSource = dst.Tables("ARTPYMTB")

        grdARTPYMT2_SPLIT.DataSource = New DataView(dst.Tables("ARTPYMT2_SPLIT"), "ISNULL(PYMT_DELETED,'0')<> '1'", "", DataViewRowState.CurrentRows)

        grdARTPYMT3.DataSource = dst.Tables("ARTPYMT3")
        grdARTPYMT4.DataSource = dst.Tables("ARTPYMT4")
        grdARTPYMT5.DataSource = dst.Tables("ARTPYMT5")

        grdARTOPENA.DataSource = dst.Tables("ARTOPENA")
        grdARTPYMTT.DataSource = dst.Tables("ARTPYMTT")
        grdARTPYMTA.DataSource = dst.Tables("ARTPYMTA")
        dst.Tables("ARTPYMTA").DefaultView.RowFilter = "ITEMS <> 0"

        grdEDT820TX.DataSource = dst.Tables("EDT820TX")
        grdEDT820T1.DataSource = dst.Tables("EDT820T1")
        grdEDT820T1.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

        cbeYP_PYMTs.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP_PYMTs.SelectedItem = cbeYP_PYMTs.Items(0)

        grpLockBox.Location = UltraGroupBox1.Location
        grpLockBox.Dock = DockStyle.Fill

        Set_SEGS(grdARTPYMT4, "ARTPYMT4")
        Set_SEGS(grdARTPYMT5, "ARTPYMT5")

        Create_Summary(grdARTPYMTB, "PYMT_BATCH_NO", "Count")
        Create_Summary(grdARTPYMTB, "S1")
        Create_Summary(grdARTPYMTB, "S2")
        Create_Summary(grdARTPYMTB, "RECORDS")
        Create_Summary(grdARTPYMTB, "CUST_PYMT_AMT")

        Create_Summary(grdARTOPENA, "AGE_AMT")
        Create_Summary(grdARTOPENA, "AGE_AMT_NEW")

        Create_Summary(grdARTPYMTX, "PYMT_BATCH_LNO", "Count")
        Create_Summary(grdARTPYMTX, "CUST_PYMT_AMT_CURR")

        Create_Summary(grdARTPYMT3, "PYMT_BATCH_ILNO", "Count")
        Create_Summary(grdARTPYMT3, "INV_BALANCE_CURR")
        Create_Summary(grdARTPYMT3, "INV_PMT_CURR")
        Create_Summary(grdARTPYMT3, "INV_DISC_TAKEN_CURR")
        Create_Summary(grdARTPYMT3, "INV_WRITE_OFF_CURR")
        Create_Summary(grdARTPYMT3, "INV_BALANCE_NEW_CURR")

        Create_Summary(grdARTPYMT4, "PYMT_BATCH_GLNO", "Count")
        Create_Summary(grdARTPYMT4, "GL_DIST_AMT_CURR")

        Create_Summary(grdARTPYMT5, "PYMT_BATCH_DLNO", "Count")
        Create_Summary(grdARTPYMT5, "GL_DIST_AMT_CURR")

        Create_Summary(grdARTPYMTA, "AR_TYPE_AMT_OLD")
        Create_Summary(grdARTPYMTA, "AR_TYPE_AMT")
        Create_Summary(grdARTPYMTA, "AR_TYPE_AMT_NEW")

        Create_Summary(grdARTPYMT2_SPLIT, "CUST_CODE", "Count")
        Create_Summary(grdARTPYMT2_SPLIT, "CUST_PYMT_AMT_CURR")

        Create_Summary(grdEDT820TX, "EDI_DOC_SEQ_NO", "Count")
        Create_Summary(grdEDT820TX, "PYMT_AMT")

        With grdARTPYMT2_SPLIT.DisplayLayout.Bands("ARTPYMT2_SPLIT")
            .Columns("PYMT_BATCH_NO").Hidden = True
        End With
        grdARTPYMT2_SPLIT.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

        With grdARTPYMT3.DisplayLayout.Bands("ARTPYMT3")
            .Columns("PYMT_BATCH_ILNO").Header.Fixed = True
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NUM").Header.Fixed = True
            .Columns("INV_NO_CONS").Header.Fixed = True
            .Columns("PARTNER_ORDR_NO").Header.Fixed = True
            .Columns("ORDR_TYPE_CODE").Header.Fixed = True

            .Columns("PYMT_BATCH_ILNO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("INV_TYPE").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("INV_NUM").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("ORDR_TYPE_CODE").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("INV_PMT_CURR").CellAppearance.BackColor = Drawing.Color.LightSkyBlue
            .Columns("INV_NO_CONS").CellAppearance.BackColor = Drawing.Color.LemonChiffon
            .Columns("PARTNER_ORDR_NO").CellAppearance.BackColor = Drawing.Color.LemonChiffon

            Dim sqlTT As String = ""
            If ROWs("ARTPARM1").Item("AR_PARM_USE_DISC") & "" = "1" Then
                If ROWs("ARTPARM1").Item("AR_PARM_HDG_DISC") & "" <> "" Then
                    .Columns("INV_DISC_TAKEN_CURR").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_HDG_DISC") & ""
                End If
                chkTAKEDISC.Visible = True
                chkTAKEDISC.Text = "Take " & .Columns("INV_DISC_TAKEN_CURR").Header.Caption
                dst.Tables("ARTPYMTT").Rows.Find("2").Item("PYMT_TOTAL_CAPTION") = .Columns("INV_DISC_TAKEN_CURR").Header.Caption
            Else
                sqlTT &= " and PYMT_TOTAL_CODE <> '2'"
                .Columns("INV_DISC_TAKEN_CURR").Hidden = True
                chkTAKEDISC.Visible = False
            End If

            If ROWs("ARTPARM1").Item("AR_PARM_USE_WOFF") & "" = "1" Then
                If ROWs("ARTPARM1").Item("AR_PARM_HDG_WOFF") & "" <> "" Then
                    .Columns("INV_WRITE_OFF_CURR").Header.Caption = ROWs("ARTPARM1").Item("AR_PARM_HDG_WOFF") & ""
                End If
                chkTAKEWOFF.Visible = True
                chkTAKEWOFF.Text = "Take " & .Columns("INV_WRITE_OFF_CURR").Header.Caption
                dst.Tables("ARTPYMTT").Rows.Find("3").Item("PYMT_TOTAL_CAPTION") = .Columns("INV_WRITE_OFF_CURR").Header.Caption
            Else
                sqlTT &= " and PYMT_TOTAL_CODE <> '3'"
                .Columns("INV_WRITE_OFF_CURR").Hidden = True
                chkTAKEWOFF.Visible = False
            End If

            If sqlTT <> "" Then
                Dim dvw As DataView = DirectCast(grdARTPYMTT.DataSource, DataTable).DefaultView
                dvw.RowFilter = Mid(sqlTT, 6)
            End If
        End With


        With grdARTPYMT5.DisplayLayout.Bands("ARTPYMT5")
            .Columns("PYMT_BATCH_DLNO").Header.Fixed = True
            .Columns("REASON_CODE").Header.Fixed = True

            .Columns("PYMT_BATCH_DLNO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("REASON_CODE").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("GL_DIST_AMT_CURR").CellAppearance.BackColor = Drawing.Color.LightSkyBlue
        End With

        With grdARTPYMTX.DisplayLayout.Bands("ARTPYMTX")
            .Columns("PYMT_BATCH_NO").Header.Fixed = True
            .Columns("PYMT_BATCH_LNO").Header.Fixed = True

            .Columns("PYMT_BATCH_NO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("PYMT_BATCH_LNO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CUST_PYMT_AMT_CURR").CellAppearance.BackColor = Drawing.Color.LightSkyBlue
        End With

        Set_Read_Only(grpBatchInfo, True)

        Sort_grdColumns(grdARTPYMTT, "PYMT_TOTAL_CODE")
        Sort_grdColumns(grdARTPYMTA, "AR_TYPE_SEQ")

        grdARTPYMT3.DisplayLayout.Bands(0).Columns("INV_DISC_TAKEN_CURR").Hidden = (ROWs("ARTPARM1").Item("AR_PARM_USE_DISC") & "" <> "1")
        grdARTPYMT3.DisplayLayout.Bands(0).Columns("INV_WRITE_OFF_CURR").Hidden = (ROWs("ARTPARM1").Item("AR_PARM_USE_WOFF") & "" <> "1")

        Bind_Controls(Me, "ARTPYMT1")
        Bind_Controls(Me, "ARTCUST1")

        Set_Read_Only(grpCustomerInfo, True)

        splCC.Visible = False

        Calculate_Aging_Dates()


        tabMain.Tabs("Lock-Box Receipts (EDI)").Visible = (ROWs("ARTPARM1").Item("AR_PARM_ENABLE_EDI_823") & "" = "1")
        tabMain.Tabs("EDI (820)").Visible = (ROWs("ARTPARM1").Item("AR_PARM_ENABLE_EDI_820") & "" = "1")
        tabMain.Tabs("Credit Cards (Settled)").Visible = (ROWs("ARTPARM1").Item("AR_PARM_ENABLE_CC") & "" = "1")
        tabMain.Tabs("Electronic Payments (ACH)").Visible = (ROWs("ARTPARM1").Item("AR_PARM_ENABLE_ACH") & "" = "1")

        Show_AR_Item_Columns(True)

        DTEs_YP = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)

        tab.Visible = False

        cmdApplyXLS.Visible = False

        'If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
        If ASCMAIN1.CLIENT = "INT" Then
            grdARTPYMT3.DisplayLayout.Bands(0).Columns("PARTNER_ORDR_NO").Header.Caption = "Clarins Inv"
            cmdApplyXLS.Visible = True
        End If
        If ASCMAIN1.CLIENT = "AHA" Then
            grdARTPYMT3.DisplayLayout.Bands(0).Columns("PARTNER_ORDR_NO").Header.Caption = "Ptnr Order No"
            cmdApplyXLS.Visible = True
        End If


        If ASCMAIN1.CLIENT = "VAN" Or ASCMAIN1.CLIENT = "RGI" Then
            cmdApplyXLS.Visible = True
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Apply Payment", "View", "Reverse"

                If grdARTPYMTX.ActiveRow Is Nothing Then
                    EMsg &= "Select a Payment from the Grid"
                ElseIf grdARTPYMTX.Selected.Rows.Count > 1 Then
                    EMsg &= "This option applies to only 1 Payment at a time"
                ElseIf grdARTPYMTX.Selected.Rows.Count = 1 AndAlso Not grdARTPYMTX.Selected.Rows(0).Equals(grdARTPYMTX.ActiveRow) Then
                    EMsg &= "Selected (Highlighed) Row and Active Row (the one with the pointer) are not the same Row"
                Else
                    Absx1.txtFor("PYMT_BATCH_NO").Text = grdARTPYMTX.ActiveRow.Cells("PYMT_BATCH_NO").Text
                    Absx1.numFor("PYMT_BATCH_LNO").Value = Val(grdARTPYMTX.ActiveRow.Cells("PYMT_BATCH_LNO").Value)
                    Absx1.txtFor("CUST_CODE").Text = grdARTPYMTX.ActiveRow.Cells("CUST_CODE").Text
                    Absx1.txtFor("CUST_NAME").Text = grdARTPYMTX.ActiveRow.Cells("CUST_NAME").Text
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text, True)
                    Absx1.txtFor("CUST_CITY").Text = rowARTCUST1.Item("CUST_CITY") & ""
                    Absx1.txtFor("CUST_STATE").Text = rowARTCUST1.Item("CUST_STATE") & ""
                    Absx1.txtFor("CUST_ZIP_CODE").Text = rowARTCUST1.Item("CUST_ZIP_CODE") & ""

                    If Absx1.txtFor("PYMT_BATCH_NO").Text = "" Then
                        EMsg &= "You Must Select a Payment by Double-Clicking a Row"
                    Else
                        Dim row As DataRow = LookUp("ARTPYMT2", New String() {Absx1.txtFor("PYMT_BATCH_NO").Text, Val(Absx1.numFor("PYMT_BATCH_LNO").Value & "")})
                        If row Is Nothing Then
                            EMsg &= "Cannot determine Payment selected"
                        Else
                            If eItemKey = "Apply Payment" Then
                                If row.Item("PYMT_DELETED") & "" = "1" Then
                                    EMsg &= "Batch " & Absx1.txtFor("PYMT_BATCH_NO").Text & " Line " & Absx1.numFor("PYMT_BATCH_LNO").Value & " has been Deleted and is no longer available to be applied"
                                End If
                                If row.Item("PYMT_STATUS") <> "1" Then
                                    EMsg &= "Batch " & Absx1.txtFor("PYMT_BATCH_NO").Text & " Line " & Absx1.numFor("PYMT_BATCH_LNO").Value & " is no longer available to be applied"
                                End If
                            End If
                            If (eItemKey = "View" Or eItemKey = "Reverse") Then
                                If row.Item("PYMT_STATUS") <> "2" Then
                                    EMsg &= "Batch " & Absx1.txtFor("PYMT_BATCH_NO").Text & " Line " & Absx1.numFor("PYMT_BATCH_LNO").Value & " has not been applied (yet)"
                                End If
                                If eItemKey = "Reverse" Then
                                    If row.Item("PYMT_REVERSED") & "" = "1" Then
                                        EMsg &= "Batch " & Absx1.txtFor("PYMT_BATCH_NO").Text & " Line " & Absx1.numFor("PYMT_BATCH_LNO").Value & " has already been Reversed (See Batch-Line " & row.Item("PYMT_BATCH_NO_REV") & "-" & row.Item("PYMT_BATCH_LNO_REV") & ")"
                                    End If
                                    If row.Item("PYMT_REVERSED") & "" = "2" Then
                                        EMsg &= "Batch " & Absx1.txtFor("PYMT_BATCH_NO").Text & " Line " & Absx1.numFor("PYMT_BATCH_LNO").Value & " is a reversing entry for Batch-Line " & row.Item("PYMT_BATCH_NO_ORIG") & "-" & row.Item("PYMT_BATCH_LNO_ORIG")
                                    End If
                                    If row.Item("PYMT_DELETED") & "" = "1" Then
                                        EMsg &= "Batch " & Absx1.txtFor("PYMT_BATCH_NO").Text & " Line " & Absx1.numFor("PYMT_BATCH_LNO").Value & " was Deleted"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If EMsg = "" And (eItemKey = "Apply Payment" Or eItemKey = "Reverse") Then
                    If Not ASCMAIN1.Logical_Open("ARTPYMT1", Absx1.txtFor("PYMT_BATCH_NO").Text) Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Lock("ARTPYMT1", Absx1.txtFor("PYMT_BATCH_NO").Text & ":" & CStr(Absx1.numFor("PYMT_BATCH_LNO").Value & "")) Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then
                        Exit Sub
                    End If
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then
                            Exit Sub
                        End If
                    End If
                    If eItemKey = "Reverse" Then
                        If Not ASCMAIN1.Logical_Open("R", "ARRPYMT2") Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Reverse Application", "Reverse && Edit", "Return Payment"

                Dim row As DataRow = dst.Tables("ARTPYMT2").Rows(0)
                If eItemKey = "Reverse && Edit" And row.Item("EDI_DOC_SEQ_NO") & "" <> "" Then
                    EMsg &= vbCr & "Cannot Reverse and Edit an EDI payment" _
                        & vbCr & "You must Reverse, and then re-select to apply from the 820 tab"
                End If

                If EMsg = "" Then

                    If UltraLabel1.Text = "Non-AR" Then
                        If Absx1.txtFor("RETURNED_ITEM_REASON").Text = "" Then
                            Dim RETURNED_ITEM_REASON As String = ""
                            RETURNED_ITEM_REASON = InputBox("Enter the reason for Reversing Payment")
                            If RETURNED_ITEM_REASON.Length > 60 Then
                                RETURNED_ITEM_REASON = RETURNED_ITEM_REASON.Substring(0, 60)
                            End If
                            Absx1.txtFor("RETURNED_ITEM_REASON").Text = RETURNED_ITEM_REASON
                        End If
                    End If



                    If Absx1.txtFor("RETURNED_ITEM_REASON").Text = "" Then
                        EMsg &= vbCr & "Enter the reason for Reversing Payment"
                    End If
                End If

                If EMsg = "" Then
                    Dim sfx As String = ""

                    If eItemKey = "Reverse && Edit" Then
                        sfx = vbCrLf & vbCrLf & "Afterwards, the payment application details (as entered) will be presented," _
                            & vbCrLf & " for editing." _
                            & vbCrLf & vbCrLf & "You may then make changes as required," _
                            & vbCrLf & " and click Update to record the corrected Application," _
                            & vbCrLf & " or click Cancel to leave the Payment in the Unapplied Status."
                    End If
                    If eItemKey = "Return Payment" Then
                        Dim RETURNED_ITEM_FEE As Decimal = Val(Absx1.numFor("RETURNED_ITEM_FEE").Value & "")
                        sfx = vbCrLf & vbCrLf & "The payment will then be deleted."
                        If RETURNED_ITEM_FEE = 0 Then
                            sfx &= "" _
                            & vbCrLf & vbCrLf & "There will be no Bank Fee charged to the Customer's account."
                        Else
                            sfx &= "" _
                            & vbCrLf & vbCrLf & "The reversing transaction will include a Bank Fee charge of " & Format(RETURNED_ITEM_FEE, "$#.00") & "," _
                            & vbCrLf & " which will appear as a DR on the Customer's account."
                        End If
                    End If

                    If grdARTPYMTX.Tag = "Reverse All Selected" Then
                        ' do not ask this for each payment reversed
                    Else
                        If MsgBox("This option will Reverse the Application of this payment:" & vbCrLf _
                              & vbCrLf & vbTab & "Batch " & Absx1.txtFor("PYMT_BATCH_NO").Text & ", Line " & Absx1.numFor("PYMT_BATCH_LNO").Value _
                              & vbCrLf & vbTab & "Customer " & Absx1.txtFor("CUST_CODE").Text & " : " & Absx1.txtFor("CUST_NAME").Text _
                              & vbCrLf & vbTab & "Payment Reference " & Absx1.txtFor("CUST_PYMT_REF_NO").Text _
                              & vbCrLf & vbTab & "Payment Amount " & Format(Val(Absx1.numFor("CUST_PYMT_AMT_CURR").Value & ""), "###,##0.00") _
                              & vbCrLf & vbCr & "by entering a negative payment and applying in reverse." _
                              & sfx _
                              & vbCrLf & vbCrLf & "Click OK to Continue" _
                              , MsgBoxStyle.OkCancel, "Verification") <> MsgBoxResult.Ok Then
                            Exit Sub
                        End If

                    End If
                End If

            Case "Move Payment"

                If grdARTPYMTX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You must Select the Payments (in the grid) to Move to another Customer"
                End If

                Dim rowARTCUST1_MOVE_TO As DataRow = Nothing
                If Absx1.txtFor("CUST_CODE_MOVE_TO").Text = "" Then
                    EMsg &= vbCr & "You must Specify the Customer to Move the Payment To"
                Else
                    rowARTCUST1_MOVE_TO = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE_MOVE_TO").Text)
                    If rowARTCUST1_MOVE_TO Is Nothing Then
                        EMsg &= vbCr & "Invalid Customer Specified (" & Absx1.txtFor("CUST_CODE_MOVE_TO").Text & ")"
                    End If
                End If

                If EMsg = "" Then
                    If MsgBox("OK to Move " & CStr(grdARTPYMTX.Selected.Rows.Count) _
                              & " Selected Payment(s) to: " & vbCrLf _
                              & rowARTCUST1_MOVE_TO.Item("CUST_CODE") & ":" _
                              & rowARTCUST1_MOVE_TO.Item("CUST_NAME"),
                               MsgBoxStyle.YesNo, "Verification to Move Payment(s)") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

                If EMsg = "" Then
                    Dim PYMT_BATCH_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMTX.Selected.Rows
                        Dim PYMT_BATCH_NO As String = grow.Cells("PYMT_BATCH_NO").Text
                        Dim PYMT_BATCH_LNO As Int32 = Val(grow.Cells("PYMT_BATCH_LNO").Text)

                        Dim rowARTPYMT2 As DataRow = LookUp("ARTPYMT2", New String() {PYMT_BATCH_NO, PYMT_BATCH_LNO})
                        If rowARTPYMT2.Item("PYMT_STATUS") & "" <> "1" Then
                            ASCMAIN1.MultiTask_Release()
                            EMsg &= vbCr & "Cannot Move a Payment which has been Applied"
                            Exit For
                        End If
                        If rowARTPYMT2.Item("PYMT_DELETED") & "" = "1" Then
                            EMsg &= vbCr & "Cannot Move a Payment which has been Deleted"
                            Exit For
                        End If

                        If Not PYMT_BATCH_NOs.Contains(PYMT_BATCH_NO) Then
                            PYMT_BATCH_NOs.Add(PYMT_BATCH_NO)
                            If Not ASCMAIN1.Logical_Open("ARTPYMT1", PYMT_BATCH_NO) Then
                                Exit Sub
                            End If
                        End If
                        If Not ASCMAIN1.Logical_Lock("ARTPYMT1", PYMT_BATCH_NO & ":" & CStr(PYMT_BATCH_LNO)) Then
                            Exit Sub
                        End If
                    Next
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()

                Else
                    MOVE_PAYMENTS.Clear()
                    For Each GROW As UltraWinGrid.UltraGridRow In grdARTPYMTX.Selected.Rows
                        Dim PYMT_BATCH_NO As String = GROW.Cells("PYMT_BATCH_NO").Value
                        Dim PYMT_BATCH_LNO As Int32 = Val(GROW.Cells("PYMT_BATCH_LNO").Value & "")
                        MOVE_PAYMENTS.Add(PYMT_BATCH_NO & vbTab & CStr(PYMT_BATCH_LNO))
                    Next
                End If

            Case "Delete Payment"
                If grdARTPYMTX.Selected.Rows.Count = 0 Then
                    EMsg &= "You Must Select Payments to Delete"
                End If

                If EMsg = "" Then
                    Dim PYMT_BATCH_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMTX.Selected.Rows
                        Dim PYMT_BATCH_NO As String = grow.Cells("PYMT_BATCH_NO").Text
                        Dim PYMT_BATCH_LNO As Int32 = Val(grow.Cells("PYMT_BATCH_LNO").Text.Replace(",", ""))

                        Dim rowARTPYMT2 As DataRow = LookUp("ARTPYMT2", New String() {PYMT_BATCH_NO, PYMT_BATCH_LNO})
                        If rowARTPYMT2.Item("PYMT_STATUS") & "" <> "1" Then
                            ASCMAIN1.MultiTask_Release()
                            EMsg &= vbCr & "Cannot Delete a Payment which has been Applied"
                            Exit For
                        End If
                        If Not PYMT_BATCH_NOs.Contains(PYMT_BATCH_NO) Then
                            PYMT_BATCH_NOs.Add(PYMT_BATCH_NO)
                            If EMsg = "" Then
                                If Not ASCMAIN1.Logical_Open("ARTPYMT1", PYMT_BATCH_NO) Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        If EMsg = "" Then
                            If Not ASCMAIN1.Logical_Lock("ARTPYMT1", PYMT_BATCH_NO & ":" & CStr(PYMT_BATCH_LNO)) Then
                                Exit Sub
                            End If
                        End If
                    Next
                End If

                If EMsg = "" Then

                    Dim frmASFMSGBF As New ASFMSGBF
                    PYMT_NOTE = frmASFMSGBF.Get_txt_from_User("Reason for Deletion", "Enter a Reason for Deleting the Payment(s)")
                    If frmASFMSGBF.user_option = -1 Then
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    If MessageBox.Show(String.Format("OK to Delete the {0} Selected Payment(s)", grdARTPYMTX.Selected.Rows.Count), "Verificaton", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                        grdARTPYMTX.Selected.Rows.Clear()
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                End If

            Case "Apply w/out Payment"
                Validate_Code("CUST_CODE")

                If EMsg = "" Then

                    If PYMT_BATCH_NO_application_only <> "" Then
                        Dim row1 As DataRow = LookUp("ARTPYMT1", PYMT_BATCH_NO_application_only)
                        If row1 Is Nothing Then
                            PYMT_BATCH_LNO_application_only = 0
                        End If
                    End If


                    If PYMT_BATCH_LNO_application_only = 0 Then
                        If PYMT_BATCH_NO_application_only = "" Then
                            PYMT_BATCH_NO_application_only = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
                        End If

                        Dim PYMT_BATCH_DATE As Date = DATETIME_STAMP.Date
                        If Format(PYMT_BATCH_DATE, "yyyyMMdd") < Format(DTEs_YP(1), "yyyyMMdd") Then
                            PYMT_BATCH_DATE = DTEs_YP(1)
                        End If
                        If Format(PYMT_BATCH_DATE, "yyyyMMdd") > Format(DTEs_YP(DTEs_YP.Length - 1), "yyyyMMdd") Then
                            PYMT_BATCH_DATE = DTEs_YP(DTEs_YP.Length - 1)
                        End If

                        dst.Tables("ARTPYMT1").Rows.Clear()
                        Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
                        With rowARTPYMT1
                            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_application_only
                            .Item("PYMT_BATCH_DATE") = PYMT_BATCH_DATE
                            .Item("STATUS") = "1"
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_DATE") = DATETIME_STAMP
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("PYMT_APPL_ONLY") = "1"
                            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                            .Item("CURR_EXCH_RATE") = 1
                            .Item("PYMT_SOURCE") = "MAN"
                        End With
                        dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)
                    End If
                    PYMT_BATCH_LNO_application_only += 1

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    Dim CUST_NAME As String = Absx1.txtFor("CUST_NAME").Text

                    dst.Tables("ARTPYMT2").Rows.Clear()

                    Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
                    With rowARTPYMT2
                        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_application_only
                        .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_application_only
                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_NAME") = CUST_NAME
                        .Item("CUST_PYMT_REF_NO") = ""
                        .Item("CUST_PYMT_REF_DATE") = DATETIME_STAMP.Date
                        .Item("CUST_PYMT_AMT") = 0
                        .Item("PYMT_STATUS") = "1"
                        .Item("CUST_PYMT_AMT_CURR") = 0
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID

                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        Dim CURR_CODE As String = rowARTCUST1.Item("CURR_CODE") & ""
                        If CURR_CODE = "" Then CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

                        '.Item("CURR_CODE") = DBNull.Value
                        '.Item("CURR_EXCH_RATE") = DBNull.Value

                        .Item("CURR_CODE") = CURR_CODE
                        If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                            .Item("CURR_EXCH_RATE") = 1
                        Else
                            'Dim CURR_EXCH_RATE As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, Absx1.txtFor("CURR_CODE").Text, Absx1.dteFor("PYMT_BATCH_DATE").Value)
                            Dim CURR_EXCH_RATE As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me.ROWs("GLTPARM1"), Absx1.txtFor("CURR_CODE").Text, Absx1.dteFor("PYMT_BATCH_DATE").Value)

                            .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                        End If
                    End With
                    dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

                    INIT_LAST("ARTPYMT1")

                    BeginTrans()
                    Update_Record_TDA("ARTPYMT1")
                    Update_Record_TDA("ARTPYMT2")
                    CommitTrans()

                    Absx1.txtFor("PYMT_BATCH_NO").Text = PYMT_BATCH_NO_application_only
                    Absx1.numFor("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO_application_only

                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Open("ARTPYMT1", Absx1.txtFor("PYMT_BATCH_NO").Text) Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Lock("ARTPYMT1", Absx1.txtFor("PYMT_BATCH_NO").Text & ":" & CStr(Absx1.numFor("PYMT_BATCH_LNO").Value & "")) Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Lock("ARTOPEN1", Absx1.txtFor("CUST_CODE").Text) Then
                        Exit Sub
                    End If
                End If


            Case "Split Payment"
                If grdARTPYMTX.ActiveRow Is Nothing Then
                    EMsg &= "Select an Unapplied Payment from the Grid"
                ElseIf grdARTPYMTX.Selected.Rows.Count > 1 Then
                    EMsg &= "This option applies to only 1 row at a time"
                ElseIf grdARTPYMTX.Selected.Rows.Count <> 1 Then
                    EMsg &= vbCr & "You must select a single Payment to Split"
                ElseIf Not grdARTPYMTX.Selected.Rows(0).Equals(grdARTPYMTX.ActiveRow) Then
                    EMsg &= "Selected (Highlighed) Row and Active Row (the one with the pointer) are not the same Row"
                End If

                If EMsg = "" Then
                    Dim CUST_CODE As String = grdARTPYMTX.ActiveRow.Cells("CUST_CODE").Text
                    Dim PYMT_BATCH_NO As String = grdARTPYMTX.ActiveRow.Cells("PYMT_BATCH_NO").Text
                    Dim PYMT_BATCH_LNO As Int32 = Val(grdARTPYMTX.ActiveRow.Cells("PYMT_BATCH_LNO").Text)

                    Dim rowARTPYMT2 As DataRow = LookUp("ARTPYMT2", New String() {PYMT_BATCH_NO, PYMT_BATCH_LNO})
                    If rowARTPYMT2.Item("PYMT_STATUS") & "" <> "1" Then
                        ASCMAIN1.MultiTask_Release()
                        EMsg &= vbCr & "Cannot Split a Payment which has been Applied"
                    End If
                    If rowARTPYMT2.Item("PYMT_DELETED") & "" = "1" Then
                        EMsg &= vbCr & "Cannot Split a Payment which has been Deleted"
                    End If

                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("ARTPYMT1", PYMT_BATCH_NO) Then
                            Exit Sub
                        End If
                        If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then
                            Exit Sub
                        End If
                        If Not ASCMAIN1.Logical_Lock("ARTOPEN1", CUST_CODE) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Update", "Update 820"

                If application_only Then
                    Dim DTE As Date = Absx1.dteFor("PYMT_BATCH_DATE").Value
                    If Format(DTE, "yyyyMMdd") < Format(DTEs_YP(1), "yyyyMMdd") _
                        Or Format(DTE, "yyyyMMdd") > Format(DTEs_YP(DTEs_YP.Length - 1), "yyyyMMdd") Then
                        EMsg &= vbCr & "Batch Date must be between " & Format(DTEs_YP(1), "MM/dd/yyyy") & " and " & Format(DTEs_YP(DTEs_YP.Length - 1), "MM/dd/yyyy")
                    End If
                End If

                If EntryMode = "S" Then
                    Dim CUST_PYMT_AMT_CURR As Decimal = Val(grdARTPYMTX.Rows(0).Cells("CUST_PYMT_AMT_CURR").Value & "")
                    Dim CUST_PYMT_AMT_CURR_SPLIT As Decimal = Val(dst.Tables("ARTPYMT2_SPLIT").Compute("SUM(CUST_PYMT_AMT_CURR)", "") & "")

                    If CUST_PYMT_AMT_CURR <> CUST_PYMT_AMT_CURR_SPLIT Then
                        EMsg &= vbCr & "Total of Splits must equal the Original Payment"
                    End If

                    If dst.Tables("ARTPYMT2_SPLIT").Compute("Count (PYMT_BATCH_LNO)", "ISNULL(CUST_PYMT_AMT_CURR,0) = 0") <> 0 Then
                        EMsg &= vbCr & "Split Payment Amounts may not be Zero - Delete the Line instead of Zeroing the Amount"
                    End If
                End If


                If EntryMode = "E" Or EntryMode = "N" Then
                    Dim activePYMT3 As Integer = Val(dst.Tables("ARTPYMT3").Compute("COUNT (PYMT_BATCH_NO)", "INV_PMT_CURR <> 0 OR INV_DISC_TAKEN_CURR <> 0 OR INV_WRITE_OFF_CURR <> 0") & "")
                    Dim activePYMT4 As Integer = Val(dst.Tables("ARTPYMT4").Compute("COUNT (PYMT_BATCH_NO)", "GL_DIST_AMT_CURR <> 0") & "")
                    Dim activePYMT5 As Integer = Val(dst.Tables("ARTPYMT5").Compute("COUNT (PYMT_BATCH_NO)", "GL_DIST_AMT_CURR <> 0") & "")

                    If activePYMT3 + activePYMT4 + activePYMT5 = 0 Then
                        EMsg &= vbCr & "No Application Details Found"
                    End If

                    TOTAL_UNAPPLIED = Get_TOTAL_UNAPPLIED()
                    If TOTAL_UNAPPLIED <> 0 Then
                        EMsg &= vbCr & "Total Unapplied is not 0"
                    End If
                End If

                For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("INV_PMT_CURR <> 0 OR INV_DISC_TAKEN_CURR <> 0 OR INV_WRITE_OFF_CURR <> 0", "", DataViewRowState.CurrentRows)
                    With rowARTPYMT3
                        Dim rowARTOPEN1 As DataRow = LookUp("ARTOPEN1", New String() {HFs("CUST_CODE"), .Item("INV_TYPE"), .Item("INV_NUM")})
                        If rowARTOPEN1.Item("OPS_YYYYPP") & "" > ASCMAIN1.CYP Then
                            EMsg &= vbCr & "You may not apply Items which have a future posting date" & vbCr & " (See AR Item " & rowARTPYMT3.Item("INV_NUM") & ")"
                        End If
                    End With
                Next

                For Each rowARTPYMT4 As DataRow In dst.Tables("ARTPYMT4").Select("ISNULL(GL_DIST_AMT_CURR,0) = 0", "", DataViewRowState.CurrentRows)
                    EMsg &= vbCr & "Distribution Amount may not be 0 (See Line " & rowARTPYMT4.Item("PYMT_BATCH_GLNO") & ")"
                Next

                EMsg &= Validate_Accounts_and_Segments_EMsg(dst.Tables("ARTPYMT4"), False)

                'For Each row As DataRow In dst.Tables("ARTPYMT4").Select("")
                '    Dim ACCT_CODE As String = row.Item("ACCT_CODE") & ""
                '    If LookUp("GLTACCT1", ACCT_CODE) Is Nothing Then
                '        EMsg &= vbCr & "Invalid Account Code " & ACCT_CODE
                '    Else
                '        If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                '            EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is not Active"
                '        End If
                '        If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                '            EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is a Control Account - no Manual J/E permitted"
                '        End If
                '    End If
                'Next

                ' Dim tbl As DataTable = New DataView(dst.Tables("ARTPYMT5"), "ISNULL(CHARGEBACK_IND,'0')<>'1'", "", DataViewRowState.CurrentRows).ToTable
                Dim tbl As DataTable = New DataView(dst.Tables("ARTPYMT5"), "", "", DataViewRowState.CurrentRows).ToTable

                Dim BAD_REASON_CODEs As New List(Of String)
                For Each row As DataRow In dst.Tables("ARTPYMT5").Select("") ' tbl.Rows
                    Dim REASON_CODE As String = row.Item("REASON_CODE")
                    Dim rowARTREAS1 As DataRow = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
                    If rowARTREAS1 Is Nothing Then
                        If Not BAD_REASON_CODEs.Contains(REASON_CODE) Then
                            EMsg &= vbCr & "Invalid Reason Code in Deduction grid (" & REASON_CODE & ")"
                            BAD_REASON_CODEs.Add(REASON_CODE)
                        End If
                    Else
                        If row.Item("CHARGEBACK_IND") & "" <> "1" Then
                            row.Item("ACCT_CODE") = rowARTREAS1.Item("ACCT_CODE")
                        End If
                    End If
                Next

                ' tbl = New DataView(tbl, "ISNULL(CHARGEBACK_IND,'0')<>'1'", "", DataViewRowState.CurrentRows).ToTable
                tbl = New DataView(dst.Tables("ARTPYMT5"), "ISNULL(CHARGEBACK_IND,'0')<>'1'", "", DataViewRowState.CurrentRows).ToTable

                EMsg &= Validate_Accounts_and_Segments_EMsg(tbl, False)

                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    If dst.Tables("ARTPYMT5").Select("ISNULL(CHARGEBACK_IND,'0') <> '1'").Length > 0 Then
                        EMsg &= vbCr & "No Write-off permitted using Deduction Codes (as per decision to use GL)"
                    End If
                End If

            Case "Delete 820"
                If grdEDT820TX.Selected.Rows.Count = 0 Then
                    If grdEDT820TX.ActiveRow IsNot Nothing Then
                        grdEDT820TX.ActiveRow.Selected = True
                    End If
                End If

                If grdEDT820TX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must First Select a Record or Records to Delete"
                Else

                    For Each grow As UltraWinGrid.UltraGridRow In grdEDT820TX.Selected.Rows
                        Dim EDI_DOC_SEQ_NO As String = grow.Cells("EDI_DOC_SEQ_NO").Value
                        If Not ASCMAIN1.Logical_Lock("EDT820T1", EDI_DOC_SEQ_NO) Then
                            Exit Sub
                        End If
                    Next

                    If MsgBox("OK to Delete the " & CStr(grdEDT820TX.Selected.Rows.Count) & " 820 Records Selected", MsgBoxStyle.YesNo, "Verification to Delete") = MsgBoxResult.No Then
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                End If

            Case "Process 820"

                If grdEDT820TX.Selected.Rows.Count = 0 Then
                    If grdEDT820TX.ActiveRow IsNot Nothing Then
                        grdEDT820TX.ActiveRow.Selected = True
                    End If
                End If

                If grdEDT820TX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must First Select Records to Cancel"

                ElseIf grdEDT820TX.Selected.Rows.Count > 1 Then
                    EMsg &= vbCr & "You May Process Only 1 Record at a Time"

                ElseIf grdEDT820TX.ActiveRow Is Nothing OrElse grdEDT820TX.ActiveRow IsNot grdEDT820TX.Selected.Rows(0) Then
                    EMsg &= vbCr & "Not Clear on which row to Process"

                End If

                If EMsg = "" Then
                    EDI_DOC_SEQ_NO = grdEDT820TX.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value
                    If Not ASCMAIN1.Logical_Lock("EDT820T1", EDI_DOC_SEQ_NO) Then
                        Exit Sub
                    End If

                    CUST_CODE = grdEDT820TX.ActiveRow.Cells("CUST_CODE").Value & ""
                    If CUST_CODE = "" Then
                        EMsg &= vbCr & "Cannot Determine Customer to use in EDI Payment " & EDI_DOC_SEQ_NO
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Open("ARTOPEN1", "*") Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Lock("ARTOPEN1", CUST_CODE) Then
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

            Case "Apply Payment"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Reverse"
                EntryMode = "X"
                Load_Record()
                Mode_Settings(True)

            Case "Reverse Application", "Reverse && Edit", "Return Payment"

                reverse_application_option = eItemKey
                Reverse_Payment()

                Dim sfx As String = ""
                If eItemKey = "Reverse && Edit" Then
                    sfx = vbCrLf & vbCrLf & "It will now be restored for Editing." _
                        & vbCrLf & vbCrLf & "If you click Cancel, it will remain as an unapplied payment."
                Else

                End If

                If grdARTPYMTX.Tag = "Reverse All Selected" Then
                Else
                    MsgBox("This payment has been successfully reversed with a negative payment." & sfx,
                           MsgBoxStyle.OkOnly, "Verification")
                End If

                If eItemKey = "Reverse && Edit" Then
                    Restore_Payment()

                    tabMain.SelectedTab = tabMain.Tabs("Unapplied Payments")

                    PYMT_BATCH_NO = PYMT_BATCH_NO_new
                    PYMT_BATCH_LNO = PYMT_BATCH_LNO_new

                    EntryMode = "R"
                    Load_Record()
                    EntryMode = "E"
                    Mode_Settings(True)

                    HFs("PYMT_BATCH_NO") = PYMT_BATCH_NO
                    HFs("PYMT_BATCH_LNO") = PYMT_BATCH_LNO

                Else
                    Mode_Settings(False)
                End If

            Case "Update", "Update 820"
                If EntryMode = "S" Then
                    Update_Split()
                Else
                    Update_Record()
                End If

                Mode_Settings(False)

            Case "Delete Payment"
                dst.Tables("ARTPYMT2").Rows.Clear()
                For Each grow As UltraWinGrid.UltraGridRow _
                In grdARTPYMTX.Selected.Rows
                    Fill_Record("ARTPYMT2", New String() _
                    {grow.Cells("PYMT_BATCH_NO").Text,
                     grow.Cells("PYMT_BATCH_LNO").Text.Replace(",", "")}, False, False)
                Next

                For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Rows
                    rowARTPYMT2.Item("PYMT_NOTE") = PYMT_NOTE
                    rowARTPYMT2.Item("PYMT_DELETED") = "1"
                    rowARTPYMT2.Item("PYMT_STATUS") = "2"
                    rowARTPYMT2.Item("LAST_DATE") = DATETIME_STAMP
                    rowARTPYMT2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                Next

                BeginTrans()
                Update_Record_TDA("ARTPYMT2")
                CommitTrans(grdARTPYMTX.Selected.Rows.Count & " Payment(s) have been Deleted")

                Mode_Settings(False)


            Case "Move Payment"
                Mode_Settings(True)

                'dst.Tables("ARTPYMT2").Rows.Clear()
                For Each P As String In MOVE_PAYMENTS
                    Dim PYMT_BATCH_NO As String = Split(P, vbTab)(0)
                    Dim PYMT_BATCH_LNO As Int32 = Val(Split(P, vbTab)(1))
                    Fill_Record("ARTPYMT2", New String() _
                    {PYMT_BATCH_NO,
                     PYMT_BATCH_LNO}, False, False)
                Next

                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE_MOVE_TO").Text
                Dim CUST_NAME As String = LookUp("ARTCUST1", CUST_CODE).Item("CUST_NAME") & ""

                For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Rows
                    If rowARTPYMT2.Item("CUST_CODE_ORIG") & "" = "" Then
                        rowARTPYMT2.Item("CUST_CODE_ORIG") = rowARTPYMT2.Item("CUST_CODE")
                    End If
                    rowARTPYMT2.Item("CUST_CODE") = CUST_CODE
                    rowARTPYMT2.Item("CUST_NAME") = CUST_NAME
                Next

                Dim pymt_count As Int32 = MOVE_PAYMENTS.Count

                BeginTrans()
                Update_Record_TDA("ARTPYMT2")
                CommitTrans()
                Mode_Settings(False)

                MsgBox(CStr(pymt_count) & " Payment(s) have been Moved to " & CUST_CODE, MsgBoxStyle.OkOnly, "Verification")

            Case "Apply w/out Payment"
                application_only = True
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Split Payment"
                EntryMode = "S"
                Setup_Split_Payment()
                Mode_Settings(True)

            Case "Cancel", "Done", "Cancel 820"
                If eItemKey = "Done" And application_only Then
                    application_only = False ' otherwise, Clear_Record might have code that deletes rows from ARTPYMT2 after calling up an application to view it
                End If
                If eItemKey = "Cancel" And application_only Then
                    ARCMAIN1.Clean_Out_ARTPYMT2_Started_NOT_Completed(HFs("PYMT_BATCH_NO"), Val(HFs("PYMT_BATCH_LNO")), False)
                End If
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Excel"
                If EntryMode = "C" Then
                    Export_to_Excel(New UltraWinGrid.UltraGrid() {grdARTCCPA0, grdARTCCPS2, grdARTCCPA1})
                    grdARTCCPA0.Tag = "*"
                End If

            Case "Print"
                If EntryMode = "L" Then
                    Print_Receipt()
                End If

            Case "Delete 820"
                Delete_820()

            Case "Process 820"
                EntryMode = "E"
                edi_820_in_process = True

                EnforceConstraints(False)
                rowEDT820T1 = Fill_Record("EDT820T1", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T2", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T3", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T4", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T5", EDI_DOC_SEQ_NO)
                dst.Tables("EDTERRS1").Rows.Clear()
                dst.Tables("EDTINVC1").Rows.Clear()
                EnforceConstraints(True)

                Create_ARTCASH1_2()

                Absx1.txtFor("PYMT_BATCH_NO").Text = PYMT_BATCH_NO
                Absx1.numFor("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO
                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                ' Absx1.txtFor("CUST_NAME").Text = ""
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE, True)
                Absx1.txtFor("CUST_CITY").Text = rowARTCUST1.Item("CUST_CITY") & ""
                Absx1.txtFor("CUST_STATE").Text = rowARTCUST1.Item("CUST_STATE") & ""
                Absx1.txtFor("CUST_ZIP_CODE").Text = rowARTCUST1.Item("CUST_ZIP_CODE") & ""

                Load_Record()
                Delete_ARTCASH1_2()

                dst.Tables("ARTPYMT1").Rows(0).SetAdded()
                dst.Tables("ARTPYMT2").Rows(0).SetAdded()

                'If EDI_Process() Then

                'End If
                EDI_Process()
                Mode_Settings(True)

                'If ASCMAIN1.CLIENT = "VAN" Then
                '    Set_Read_Only_for_ctl(Absx1.dteFor("PYMT_BATCH_DATE"), False)
                'End If

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                If Not ScreenMode Or EntryMode = "E" Or EntryMode = "V" Or EntryMode = "X" Then
                    With .Groups("Screen Control")
                        .Items("Apply Payment").Settings.Enabled = not_iScreenMode
                        .Items("Apply Payment").Visible = Not ScreenMode
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Update").Visible = ScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                        .Items("Cancel").Visible = ScreenMode
                        .Items("Delete Payment").Settings.Enabled = not_iScreenMode
                        .Items("Delete Payment").Visible = Not ScreenMode
                        .Items("Apply w/out Payment").Settings.Enabled = not_iScreenMode
                        .Items("Apply w/out Payment").Visible = Not ScreenMode
                        .Items("Split Payment").Settings.Enabled = not_iScreenMode
                        .Items("Split Payment").Visible = Not ScreenMode
                    End With

                    .Groups("Match").Visible = False

                    With .Groups("EDI (820)")
                        .Items("Process 820").Visible = Not ScreenMode
                        .Items("Delete 820").Visible = Not ScreenMode
                        .Items("Update 820").Visible = ScreenMode And (dst.Tables("EDTERRS1").Rows.Count = 0)
                        .Items("Cancel 820").Visible = ScreenMode
                    End With

                    With .Groups("Post Application Options")
                        .Items("View").Settings.Enabled = not_iScreenMode
                        .Items("View").Visible = Not ScreenMode Or (EntryMode = "V")
                        .Items("Reverse").Settings.Enabled = not_iScreenMode
                        .Items("Reverse").Visible = Not ScreenMode Or (EntryMode = "V")
                        .Items("Reverse Application").Visible = ScreenMode And (EntryMode = "X")
                        .Items("Reverse && Edit").Visible = ScreenMode And (EntryMode = "X")
                        .Items("Return Payment").Visible = ScreenMode And (EntryMode = "X")
                        .Items("Done").Settings.Enabled = iScreenMode
                    End With

                    .Groups("Unapplied Payment Options").Visible = Not ScreenMode
                    .Groups("Payment Info").Visible = ScreenMode
                    .Groups("Control Totals").Visible = ScreenMode 'And Absx1.txtFor("CUST_CODE").Text <> ""
                End If

                If Not tf Or EntryMode = "L" Then
                    With .Groups("Lock Box Options")
                        .Items("Select Receipt").Settings.Enabled = not_iScreenMode
                        .Items("Print").Settings.Enabled = iScreenMode
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End With
                End If

                If Not tf Or EntryMode = "C" Then
                    With .Groups("Credit Card Options")
                        .Items("Inquiry Batch").Settings.Enabled = not_iScreenMode
                        .Items("Excel").Settings.Enabled = iScreenMode
                        .Items("Settle Batch").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End With
                End If
            End With
        End If

        splARTPYMT2.Panel2Collapsed = (Not ScreenMode) Or (EntryMode <> "S")
        tabARTPYMT3.Tabs("EDI (820)").Visible = (ScreenMode AndAlso (edi_820_in_process Or (EntryMode = "V" And EDI_DOC_SEQ_NO <> "")))
        tabARTPYMT3.Tabs("Match").Visible = (ScreenMode And (EntryMode = "E")) And application_only
        cmdMatchApply.Visible = False

        tabDeductions.Tabs("Code XRef").Visible = (ScreenMode And edi_820_in_process)

        tabSummary.Tabs("Reverse Payment").Visible = ScreenMode And (EntryMode = "X")
        If EntryMode = "X" Then
            tabSummary.SelectedTab = tabSummary.Tabs("Reverse Payment")
        End If


        If EntryMode = "L" Then
            'UltraGroupBox1.Visible = Not tf
            grpLockBox.Visible = True
            For Each tab As Infragistics.Win.UltraWinTabControl.UltraTab In tabMain.Tabs
                tab.Enabled = (tab.Text = "Lock-Box Receipts")
            Next
            tabLockBoxDetails.Tabs("Payments to be Applied").Visible = True
        ElseIf EntryMode = "S" Then
            For Each tab As Infragistics.Win.UltraWinTabControl.UltraTab In tabMain.Tabs
                tab.Enabled = (tab.Text = "Unapplied Payments")
            Next
            Setup_Control_Panel()
        Else
            For Each tab As Infragistics.Win.UltraWinTabControl.UltraTab In tabMain.Tabs
                tab.Enabled = True
            Next
            tabLockBoxDetails.Tabs("Payments to be Applied").Visible = False
            grpLockBox.Visible = False
            Set_Read_Only(UltraGroupBox1, ScreenMode)

            If ScreenMode And application_only Then
                'Set_Read_Only_for_ctl(Absx1.dteFor("PYMT_BATCH_DATE"), False)
            Else
                'Set_Read_Only_for_ctl(Absx1.dteFor("PYMT_BATCH_DATE"), True)
            End If

            grpBatchInfo.Visible = ScreenMode

            tabMain.Visible = Not ScreenMode
            tabARTPYMT3.Visible = ScreenMode

            Setup_Control_Panel()

            txtIGNORE.Text = ""

            If ScreenMode Then

                If ASCMAIN1.CLIENT = "AHA" Then
                    cmdApplyXLS.Visible = (HFs("CUST_CODE") = "ULTA")
                End If

                grdARTPYMT5.DisplayLayout.Bands(0).Columns("GL_DIST_COMMENT").Hidden = Not edi_820_in_process

                For Each tab As Infragistics.Win.UltraWinTabControl.UltraTab In tabARTPYMT3.Tabs
                    tab.Enabled = ((HFs.ContainsKey("CUST_CODE") AndAlso HFs("CUST_CODE") <> "") Or tab.Text = "Deductions, Chargebacks, On/Account")
                Next

                Dim inq As Boolean = (EntryMode = "V" Or EntryMode = "X")
                grpApplyOptions.Visible = Not inq

                grpApply.Visible = Not inq
                'cmdApplyAll.Visible = Not inq
                'cmdApplyToStmt.Visible = Not inq
                'cmdApplySel.Visible = Not inq
                grpInvoiceRange.Visible = Not inq

                cmdUnApplyAll.Visible = Not inq
                cmdAutoApply.Visible = Not inq

                grdARTPYMT3.DisplayLayout.Bands(0).Columns("PAY").Hidden = inq
                grdARTOPENA.DisplayLayout.Bands(0).Columns("PAY").Hidden = inq
                Set_Read_Only_for_ctl(txtPYMT_NOTE, inq)

                With grdARTPYMT5.DisplayLayout.Bands(0)
                    .Columns("INV_TYPE_CB").Hidden = (EntryMode = "E")
                    .Columns("CHARGEBACK_NO").Hidden = (EntryMode = "E")
                End With

                tabDeductions.Tabs("By Reason Code").Enabled = (HFs.ContainsKey("CUST_CODE") AndAlso HFs("CUST_CODE") <> "")
                tabDeductions.Tabs("By GL Account").Enabled = ((HFs.ContainsKey("CUST_CODE") AndAlso HFs("CUST_CODE") <> "") Or UltraLabel1.Text = "Non-AR")

                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdARTPYMT3, grdARTPYMT4, grdARTPYMT5}
                    With grd.DisplayLayout.Override
                        If EntryMode = "V" Or EntryMode = "X" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.False
                        Else
                            If grd.Equals(grdARTPYMT3) Then
                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                .AllowDelete = DefaultableBoolean.False
                                .AllowUpdate = DefaultableBoolean.True
                            Else
                                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                                .AllowDelete = DefaultableBoolean.True
                                .AllowUpdate = DefaultableBoolean.True
                            End If
                        End If
                    End With
                Next
                Show_AR_Item_Columns()
            Else
                Clear_Record()
            End If
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
        {"ARTPYMT1", "ARTPYMT2", "ARTPYMT2_SPLIT", "ARTPYMT3",
         "ARTPYMT4", "ARTPYMT5", "ARTOPEN1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        For Each TABLE_NAME As String In New String() _
        {"EDT820T1", "EDT820T2", "EDT820T3", "EDT820T4",
         "EDT820T5", "EDTERRS1", "EDTINVC1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        For Each TABLE_NAME As String In New String() {"ARTPYMTM", "ARTPYMTN"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        dst.AcceptChanges()

        Fill_Records("ARTPYMTB")

        Show_Filter(grdARTPYMT3, False)
        grdARTPYMT3.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Dim dv As DataView = DirectCast(grdARTPYMTX.DataSource, DataTable).DefaultView
        dv.RowFilter = ""

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("CUST_CODE").Visible = True
        UltraLabel1.Appearance.ForeColor = Drawing.Color.Empty
        UltraLabel1.Text = "Customer"

        splCC.Visible = False

        ' there is code already in Proceed_Prereq to clear out ARTPYMT2/1
        'If application_only Then
        '    ASCMAIN1.sql = "Delete from ARTPYMT2" _
        '    & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO_application_only & "'" _
        '    & "   and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO_application_only)
        '    ASCDATA1.ExecuteSQL()
        '    PYMT_BATCH_LNO_application_only -= 1

        '    If PYMT_BATCH_LNO_application_only = 0 Then
        '        ASCMAIN1.sql = "Delete from ARTPYMT1" _
        '        & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO_application_only & "'"
        '        ASCDATA1.ExecuteSQL()
        '    End If
        'End If

        application_only = False
        edi_820_in_process = False

        Absx1.txtFor("CUST_CODE_MOVE_TO").Text = ""
        chkTAKEDISC.Checked = False

        If tabMain.SelectedTab.Key = "Unapplied Payments" Then
            Setup_Control_Panel()
        Else
            tabMain.SelectedTab = tabMain.Tabs("Unapplied Payments")
        End If

        Absx1.numFor("RETURNED_ITEM_FEE").Value = Val(ROWs("ARTPARM1").Item("AR_PARM_RTN_ITEM_FEE") & "")
        Absx1.txtFor("RETURNED_ITEM_REASON").Text = ""
        Absx1.txtFor("RETURNED_ITEM_REASON_CODE").Text = ROWs("ARTPARM1").Item("AR_PARM_RTN_ITEM_REASON_CODE") & ""

        If ROWs("ARTPARM1").Item("AR_PARM_ENABLE_EDI_820") & "" = "1" Then Load_EDI_Grid()
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up for Payment Application")

        Save_Header_Fields(UltraGroupBox1)
        Save_Header_Fields(grpPaymentInfo, False)

        If EntryMode = "N" Then
            HFs("PYMT_BATCH_NO") = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        End If

        dst.AcceptChanges()
        ' LOAD RECORD APPEARS TO BE USED ONLY BY UNAPPLIED PAYMENT SELECTION, SO WE SHOULD CLEAR THE DATATABLES BEFORE FILLING

        Dim rowARTPYMT1 As DataRow
        If EntryMode = "R" Then
            rowARTPYMT1 = Fill_Record("ARTPYMT1", PYMT_BATCH_NO_new)
        Else
            rowARTPYMT1 = Fill_Record("ARTPYMT1", HFs("PYMT_BATCH_NO"))
            If Not application_only AndAlso rowARTPYMT1.Item("PYMT_APPL_ONLY") & "" = "1" Then
                application_only = True
            End If
        End If

        Dim rowARTPYMT2 As DataRow
        If EntryMode = "R" Then
            rowARTPYMT2 = dst.Tables("ARTPYMT2").Rows(0)
            'rowARTPYMT2.SetAdded()
        Else
            rowARTPYMT2 = Fill_Record("ARTPYMT2", New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO")})
        End If
        CUST_PYMT_AMT_CURR = Val(rowARTPYMT2.Item("CUST_PYMT_AMT_CURR") & "")

        rowARTCUST1 = LookUp("ARTCUST1", rowARTPYMT2("CUST_CODE") & "")

        If EntryMode = "N" Then
            rowARTPYMT1.Item("BANK_CODE") = HFs("BANK_CODE")
            rowARTPYMT1.Item("CURR_CODE") = HFs("CURR_CODE")
            rowARTPYMT1.Item("CURR_EXCH_RATE") = HFs("CURR_EXCH_RATE")
            rowARTPYMT1.Item("PYMT_BATCH_DATE") = HFs("PYMT_BATCH_DATE")
            rowARTPYMT1.Item("STATUS") = "0"
            rowARTPYMT1.Item("PYMT_SOURCE") = "MAN"
            rowARTPYMT1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        Else
            HFs("CUST_CODE") = rowARTPYMT2("CUST_CODE") & ""

            If HFs("CUST_CODE") = "" Then
                Absx1.txtFor("CUST_CODE").Visible = False
                UltraLabel1.Text = "Non-AR"
                UltraLabel1.Appearance.ForeColor = Drawing.Color.Red
                tabARTPYMT3.SelectedTab = tabARTPYMT3.Tabs("Deductions, Chargebacks, On/Account")
                tabDeductions.SelectedTab = tabDeductions.Tabs("By GL Account")

                TRADE_CLASS_CODE = ""
                rowSOTTCLS1 = Nothing
                CHANNEL_CODE = ""
                rowSOTCHAN1 = Nothing
            Else
                UltraLabel1.Text = ""
                tabARTPYMT3.SelectedTab = tabARTPYMT3.Tabs("Open AR Items")
                tabDeductions.SelectedTab = tabDeductions.Tabs("By Reason Code")

                TRADE_CLASS_CODE = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""
                rowSOTTCLS1 = LookUp("SOTTCLS1", TRADE_CLASS_CODE)
                If rowSOTTCLS1 Is Nothing Then
                    CHANNEL_CODE = ""
                Else
                    CHANNEL_CODE = rowSOTTCLS1.Item("CHANNEL_CODE") & ""
                End If
                rowSOTCHAN1 = LookUp("SOTCHAN1", CHANNEL_CODE)
            End If
        End If


        If EntryMode = "V" Then
            EDI_DOC_SEQ_NO = rowARTPYMT2.Item("EDI_DOC_SEQ_NO") & ""
            If EDI_DOC_SEQ_NO <> "" Then
                rowEDT820T1 = Fill_Record("EDT820T1", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T2", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T3", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T4", EDI_DOC_SEQ_NO)
                Fill_Records("EDT820T5", EDI_DOC_SEQ_NO)
            End If
        End If

        ASCMAIN1.Design_FCB("PAY", grdARTPYMT3)

        grdARTPYMT3.DisplayLayout.Bands(0).Columns("INV_NO_CONS").Hidden = True
        grdARTPYMT3.DisplayLayout.Bands(0).Columns("PARTNER_ORDR_NO").Hidden = True

        If EntryMode <> "V" And EntryMode <> "X" Then

            Dim tblARTPYMT3 As New DataTable
            If EntryMode = "R" Then
                tblARTPYMT3 = dst.Tables("ARTPYMT3").Copy
            Else
                dst.Tables("ARTPYMT4").Rows.Clear()
                dst.Tables("ARTPYMT5").Rows.Clear()
            End If
            dst.Tables("ARTPYMT3").Rows.Clear()

            Dim PYMT_BATCH_ILNO As Long = 0
            ASCMAIN1.sql = "Select ARTOPEN1.* from ARTOPEN1" _
                & " where ARTOPEN1.CUST_CODE = '" & HFs("CUST_CODE") & "'"
            Fill_Records("ARTOPEN1", "", True, ASCMAIN1.sql)

            For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Select("", "INV_DATE,INV_TYPE,INV_NUM")
                Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
                With rowARTPYMT3
                    If EntryMode = "R" Then
                        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                        .Item("PYMT_BATCH_LNO") = 3
                    Else
                        .Item("PYMT_BATCH_NO") = HFs("PYMT_BATCH_NO")
                        .Item("PYMT_BATCH_LNO") = HFs("PYMT_BATCH_LNO")
                    End If

                    PYMT_BATCH_ILNO += 1
                    .Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
                    For Each C As String In New String() _
                        {"INV_TYPE", "INV_NUM", "INV_DATE", "INV_DUE_DATE",
                         "REASON_CODE", "CUST_CODE_SO", "CUST_STORE_NO",
                         "INV_CUST_PO", "INV_BALANCE", "INV_BALANCE_CURR",
                         "INV_FREIGHT_CURR", "INV_MISC_CHG_CURR", "INV_TOTAL_AMOUNT_CURR",
                         "POST_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "CURR_CODE", "CURR_EXCH_RATE"}
                        .Item(C) = rowARTOPEN1.Item(C)
                    Next

                    .Item("INV_NO_CONS") = rowARTOPEN1.Item("INV_NO_CONS")
                    .Item("PARTNER_ORDR_NO") = rowARTOPEN1.Item("PARTNER_ORDR_NO")

                    For Each C As String In New String() _
                        {"INV_PMT", "INV_DISC_TAKEN", "INV_WRITE_OFF",
                         "INV_PMT_CURR", "INV_DISC_TAKEN_CURR", "INV_WRITE_OFF_CURR"}
                        .Item(C) = 0
                    Next
                    .Item("INV_BALANCE_NEW") = .Item("INV_BALANCE")
                    .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_CURR")

                    dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
                End With
            Next

            If EntryMode = "R" Then
                For Each rowARTPYMT3R As DataRow In tblARTPYMT3.Rows
                    Dim INV_TYPE As String = rowARTPYMT3R.Item("INV_TYPE")
                    Dim INV_NUM As String = rowARTPYMT3R.Item("INV_NUM")
                    Dim rowARTPYMT3() As DataRow = dst.Tables("ARTPYMT3") _
                    .Select("INV_TYPE = '" & INV_TYPE & "' and INV_NUM = '" & INV_NUM & "'")
                    If rowARTPYMT3.Length = 0 Then
                        Dim row As DataRow = dst.Tables("ARTPYMT3").NewRow
                        row.ItemArray = rowARTPYMT3R.ItemArray
                        PYMT_BATCH_ILNO += 1
                        'Debug.Write("then")
                        row.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                        row.Item("PYMT_BATCH_LNO") = 3
                        row.Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
                        row.Item("INV_BALANCE") = 0
                        row.Item("INV_BALANCE_NEW") = Val(row.Item("INV_PMT") & "") + Val(row.Item("INV_DISC_TAKEN") & "") + Val(row.Item("INV_WRITE_OFF") & "")
                        row.Item("INV_BALANCE_CURR") = 0
                        row.Item("INV_BALANCE_NEW_CURR") = Val(row.Item("INV_PMT_CURR") & "") + Val(row.Item("INV_DISC_TAKEN_CURR") & "") + Val(row.Item("INV_WRITE_OFF_CURR") & "")
                        dst.Tables("ARTPYMT3").Rows.Add(row)
                    Else
                        rowARTPYMT3(0).Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                        rowARTPYMT3(0).Item("PYMT_BATCH_LNO") = 3
                        rowARTPYMT3(0).Item("INV_PMT") = rowARTPYMT3R.Item("INV_PMT")
                        rowARTPYMT3(0).Item("INV_DISC_TAKEN") = rowARTPYMT3R.Item("INV_DISC_TAKEN")
                        rowARTPYMT3(0).Item("INV_WRITE_OFF") = rowARTPYMT3R.Item("INV_WRITE_OFF")
                        rowARTPYMT3(0).Item("INV_BALANCE_NEW") = Val(rowARTPYMT3(0).Item("INV_BALANCE") & "") _
                        - (Val(rowARTPYMT3R.Item("INV_PMT") & "") + Val(rowARTPYMT3R.Item("INV_DISC_TAKEN") & "") + Val(rowARTPYMT3R.Item("INV_WRITE_OFF") & ""))

                        rowARTPYMT3(0).Item("INV_PMT_CURR") = rowARTPYMT3R.Item("INV_PMT_CURR")
                        rowARTPYMT3(0).Item("INV_DISC_TAKEN_CURR") = rowARTPYMT3R.Item("INV_DISC_TAKEN_CURR")
                        rowARTPYMT3(0).Item("INV_WRITE_OFF_CURR") = rowARTPYMT3R.Item("INV_WRITE_OFF_CURR")
                        rowARTPYMT3(0).Item("INV_BALANCE_NEW_CURR") = Val(rowARTPYMT3(0).Item("INV_BALANCE_CURR") & "") _
                        - (Val(rowARTPYMT3R.Item("INV_PMT_CURR") & "") + Val(rowARTPYMT3R.Item("INV_DISC_TAKEN_CURR") & "") + Val(rowARTPYMT3R.Item("INV_WRITE_OFF_CURR") & ""))
                    End If
                Next
            End If
        Else
            Fill_Records("ARTPYMT3", New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO")})
            Fill_Records("ARTPYMT4", New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO")})
            Fill_Records("ARTPYMT5", New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO")})
            dst.AcceptChanges()
        End If

        Calculate_Application_by_Type()
        Calculate_Aging()

        grdARTPYMT3.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdARTPYMT3.DisplayLayout.Bands(0).SortedColumns.Add("PYMT_BATCH_ILNO", False)
        grdARTPYMT4.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdARTPYMT4.DisplayLayout.Bands(0).SortedColumns.Add("PYMT_BATCH_GLNO", False)
        grdARTPYMT5.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdARTPYMT5.DisplayLayout.Bands(0).SortedColumns.Add("PYMT_BATCH_DLNO", False)

        Fill_Record("ARTCUST1", HFs("CUST_CODE"))

        Set_Read_Only(grpPaymentInfo, True)

        optShowItems.Value = "A"
        If EntryMode = "N" Or EntryMode = "E" Then
            optShowItems.Value = "O"
        End If

        Display_Application_Totals()
        Setup_tabARTPYMT3()

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Record()
        BeginTrans()

        MsgBox("ABS Stop Statement - Deleting (please call ABS)")
        ' DON'T KNOW WHY WE WOULD EVER NEED THIS - DISABLING ALL CODE
        'Delete_Rows("ARTPYMT1")
        'Delete_Rows("ARTPYMT5")

        'Update_Record_TDA("ARTPYMT1")
        'Update_Record_TDA("ARTPYMT5")

        CommitTrans("Delete")
    End Sub

    Sub Update_Record()

        Dim CURR_EXCH_RATE As Double = Val(Absx1.numFor("CURR_EXCH_RATE").Value & "") ' Val(HFs("CURR_EXCH_RATE"))
        If CURR_EXCH_RATE = 0 Then
            MsgBox("Currency Exchange Rate is 0 - PLEASE CALL ABS", MsgBoxStyle.OkOnly, "STOP - When you click OK Update will be Aborted")
            Exit Sub
        End If

        BeginTrans()

        Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").Rows(0)
        Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").Rows(0)
        rowARTPYMT2.Item("PYMT_STATUS") = "2"
        'rowARTPYMT2.Item("PYMT_NOTE") = txtPYMT_NOTE.Text
        rowARTPYMT2.Item("LAST_DATE") = DATETIME_STAMP
        rowARTPYMT2.Item("LAST_OPER") = ASCMAIN1.USER_ID

        Remove_Zeros("ARTPYMT4", "ISNULL(GL_DIST_AMT_CURR,0) <> 0")
        For Each rowARTPYMT4 As DataRow In dst.Tables("ARTPYMT4").Rows
            rowARTPYMT4.Item("GL_DIST_AMT") = Val(rowARTPYMT4.Item("GL_DIST_AMT_CURR") & "") * CURR_EXCH_RATE
            rowARTPYMT4.AcceptChanges()
            rowARTPYMT4.SetAdded()
        Next
        Update_Record_TDA("ARTPYMT4")

        If rowARTPYMT2.Item("CUST_CODE") & "" <> "" Then

            dst.Tables("ARTOPEN1").AcceptChanges()

            Remove_Zeros("ARTPYMT3", "ISNULL(INV_PMT_CURR,0) <> 0 OR ISNULL(INV_DISC_TAKEN_CURR,0) <> 0 OR ISNULL(INV_WRITE_OFF_CURR,0) <> 0")
            Remove_Zeros("ARTPYMT5", "ISNULL(GL_DIST_AMT_CURR,0) <> 0")

            Dim PYMT_BATCH_ILNO As Integer = 0
            For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("", "PYMT_BATCH_ILNO")
                Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find(New String() {HFs("CUST_CODE"), rowARTPYMT3.Item("INV_TYPE"), rowARTPYMT3.Item("INV_NUM")})

                Dim CURR_EXCH_RATE_item As Decimal = Val(rowARTOPEN1.Item("CURR_EXCH_RATE") & "")
                If CURR_EXCH_RATE_item = 0 Then CURR_EXCH_RATE_item = 1

                PYMT_BATCH_ILNO += 1
                rowARTPYMT3.Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO
                rowARTPYMT3.Item("INV_BALANCE") = Val(rowARTPYMT3.Item("INV_BALANCE_CURR") & "") * CURR_EXCH_RATE_item
                rowARTPYMT3.Item("INV_PMT") = Val(rowARTPYMT3.Item("INV_PMT_CURR") & "") * CURR_EXCH_RATE_item
                rowARTPYMT3.Item("INV_DISC_TAKEN") = Val(rowARTPYMT3.Item("INV_DISC_TAKEN_CURR") & "") * CURR_EXCH_RATE_item
                rowARTPYMT3.Item("INV_WRITE_OFF") = Val(rowARTPYMT3.Item("INV_WRITE_OFF_CURR") & "") * CURR_EXCH_RATE_item
                rowARTPYMT3.Item("INV_BALANCE_NEW") = Val(rowARTPYMT3.Item("INV_BALANCE_NEW_CURR") & "") * CURR_EXCH_RATE_item

                rowARTPYMT3.Item("CURR_CODE") = rowARTOPEN1.Item("CURR_CODE")
                rowARTPYMT3.Item("CURR_EXCH_RATE") = rowARTOPEN1.Item("CURR_EXCH_RATE")

                Dim INV_BALANCE_CURR As Decimal = Val(rowARTPYMT3.Item("INV_BALANCE_CURR") & "")
                Dim INV_BALANCE_NEW_CURR As Decimal = Val(rowARTPYMT3.Item("INV_BALANCE_NEW_CURR") & "")
                Dim CURR_GAIN_LOSS As Decimal = (INV_BALANCE_CURR - INV_BALANCE_NEW_CURR) * (CURR_EXCH_RATE - CURR_EXCH_RATE_item)
                rowARTPYMT3.Item("CURR_GAIN_LOSS") = CURR_GAIN_LOSS

                rowARTPYMT3.AcceptChanges()
                rowARTPYMT3.SetAdded()

                rowARTOPEN1.Item("INV_BALANCE") = rowARTPYMT3.Item("INV_BALANCE_NEW")
                rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTPYMT3.Item("INV_BALANCE_NEW_CURR")
                rowARTOPEN1.Item("INV_PMT") = Val(rowARTOPEN1.Item("INV_PMT") & "") + Val(rowARTPYMT3.Item("INV_PMT") & "")
                rowARTOPEN1.Item("INV_PMT_CURR") = Val(rowARTOPEN1.Item("INV_PMT_CURR") & "") + Val(rowARTPYMT3.Item("INV_PMT_CURR") & "")
                rowARTOPEN1.Item("INV_DISC_TAKEN") = Val(rowARTOPEN1.Item("INV_DISC_TAKEN") & "") + Val(rowARTPYMT3.Item("INV_DISC_TAKEN") & "")
                rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") = Val(rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") & "") + Val(rowARTPYMT3.Item("INV_DISC_TAKEN_CURR") & "")
                rowARTOPEN1.Item("INV_WRITE_OFF") = Val(rowARTOPEN1.Item("INV_WRITE_OFF") & "") + Val(rowARTPYMT3.Item("INV_WRITE_OFF") & "")
                rowARTOPEN1.Item("INV_WRITE_OFF_CURR") = Val(rowARTOPEN1.Item("INV_WRITE_OFF_CURR") & "") + Val(rowARTPYMT3.Item("INV_WRITE_OFF_CURR") & "")
                ' SHOULDN'T ALLOW 1 OR BOTH OF THESE DATES TO BE NULL
                If HFs("PYMT_BATCH_DATE") <> "" Then
                    rowARTOPEN1.Item("INV_LAST_PMT_REF") = HFs("CUST_PYMT_REF_NO") ' rowARTPYMT2.Item("CUST_PYMT_REF_NO") & ""
                End If
                If HFs("CUST_PYMT_REF_DATE") <> "" Then
                    rowARTOPEN1.Item("INV_LAST_PMT_REF_DT") = HFs("CUST_PYMT_REF_DATE") ' rowARTPYMT2.Item("CUST_PYMT_REF_DATE") & ""
                End If
                rowARTOPEN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowARTOPEN1.Item("LAST_DATE") = DATETIME_STAMP
            Next
            For Each rowARTPYMT5 As DataRow In dst.Tables("ARTPYMT5").Rows
                rowARTPYMT5.Item("GL_DIST_AMT") = Val(rowARTPYMT5.Item("GL_DIST_AMT_CURR") & "") * CURR_EXCH_RATE

                If rowARTPYMT5.Item("CHARGEBACK_IND") & "" = "1" Then
                    Me.Load_Open_AR_from_CB(rowARTPYMT5, rowARTPYMT2, rowARTPYMT1.Item("PYMT_BATCH_DATE"))
                End If

                rowARTPYMT5.AcceptChanges()
                rowARTPYMT5.SetAdded()
            Next

            Update_Record_TDA("ARTPYMT3")
            Update_Record_TDA("ARTPYMT5")
            Update_Record_TDA("ARTOPEN1")

        End If

        INIT_LAST("ARTPYMT1", , , True)

        'Update_Record_TDA("ARTPYMT1")

        Update_Record_TDA("ARTPYMT2")

        ASCDATA1.ExecuteSP("ARPPYMTP", "VN" _
                           , New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO")} _
                           , New String() {"PYMT_BATCH_NO_IN", "PYMT_BATCH_LNO_IN"})

        If rowARTPYMT2.Item("CUST_CODE") & "" <> "" AndAlso
           rowARTPYMT1.Item("PYMT_APPL_ONLY") & "" <> "1" AndAlso
           Not ",CC,B2C,BOX,".Contains("," & rowARTPYMT1.Item("PYMT_SOURCE") & ",") Then

            ASCDATA1.ExecuteSP("ARPCUST6_PYMT", "VN" _
                               , New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO")} _
                               , New String() {"PYMT_BATCH_NO_IN", "PYMT_BATCH_LNO_IN"})
        End If

        If edi_820_in_process Then
            'rowEDT820T1.Item("EDI_PROCESS_IND") = "1"
            'Update_Record_TDA("EDT820T1")
            ASCMAIN1.sql = "Update EDT820T1 set EDI_PROCESS_IND = '1' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_PROCESS_IND = '0'"
            ASCDATA1.ExecuteSQL()

            Update_Record_TDA("ARTPYMT1")
        End If

        If application_only Then
            application_only = False
        End If

        Dim msg As String = "Update Complete"
        If applying_to_statement = True Then
            msg = ""
        End If

        CommitTrans(msg)
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        'Select Case COLUMN_NAME
        '    Case "PYMT_BATCH_NO"
        '        'sql_where = "STATUS = '0'"
        'End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTPYMTX, "SBB", "Show Filter", "Customer Inquiry", "Reverse All Selected")
        Load_Popup_Menu(grdARTPYMTM, "BB", "Select All", "De-Select All")

        Load_Popup_Menu(grdARTPYMT3, "SBSSSSSSSS", "Show Filter", "Write-Off Balance",
            "INV_NO_CONS|Cons Inv No", "PARTNER_ORDR_NO|Ptnr Ordr No", "CUST_CODE_SO|Sold-To", "CUST_STORE_NO|Location", "INV_DUE_DATE|Due Date",
            "POST_CODE|Post Code", "REASON_CODE|Reason Code", "INV_CUST_PO|Customer Ref No")
        Load_Popup_Menu(grdARTPYMT5, "BSSSSSSB", "Split Line",
            "ACCT_CODE|GL Acct", "GL_DIST_COMMENT|Comment", "CUST_CODE_SO|Sold-To",
            "OUR_REFERENCE|Our Reference", "CUST_REFERENCE|Customer Reference", "Import From Amazon")
        Load_Popup_Menu(grdEDT820T1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show Raw EDI")
        Load_Popup_Menu(grdEDT820TX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show Raw EDI")
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

        'If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        '    'e.Cancel = True
        'Else
        Select Case e.SourceControl.Name

            Case "grdARTPYMT3"
                For Each COLUMN_NAME In New String() _
                {"INV_NO_CONS", "PARTNER_ORDR_NO", "CUST_CODE_SO", "CUST_STORE_NO", "INV_DUE_DATE", "POST_CODE", "REASON_CODE", "INV_CUST_PO"}
                    tlb_sbt = DirectCast(tlb_pop.Tools(COLUMN_NAME), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.Tag = "X"
                    tlb_sbt.Checked = Not grdARTPYMT3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
                    tlb_sbt.Tag = ""
                    If COLUMN_NAME = "PARTNER_ORDR_NO" Then
                        If ASCMAIN1.CLIENT = "INT" Then
                            tlb_sbt.SharedProps.Caption = "Clarins Inv No"
                        End If
                        If ASCMAIN1.CLIENT = "AHA" Then
                            tlb_sbt.SharedProps.Caption = "Ptnr Ordr No"
                        End If
                    End If
                Next

                tlb_btn = DirectCast(tlb_pop.Tools("Write-Off Balance"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "RGI")



            Case "grdARTPYMT5"
                For Each COLUMN_NAME In New String() _
                {"ACCT_CODE", "GL_DIST_COMMENT", "CUST_CODE_SO", "OUR_REFERENCE", "CUST_REFERENCE"}
                    tlb_sbt = DirectCast(tlb_pop.Tools(COLUMN_NAME), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.Tag = "X"
                    tlb_sbt.Checked = Not grdARTPYMT5.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
                    tlb_sbt.Tag = ""
                Next

                tlb_btn = DirectCast(tlb_pop.Tools("Split Line"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E")

                If ASCMAIN1.CLIENT = "VAN" Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Import From Amazon"), UltraWinToolbars.ButtonTool)
                    If txtARTCUST1.Text = "AMAZONFBA" Or txtARTCUST1.Text = "AMAZONFBA03" Then
                        tlb_btn.SharedProps.Visible = True
                    Else
                        tlb_btn.SharedProps.Visible = False
                    End If
                Else
                    tlb_btn.SharedProps.Visible = False
                End If
            Case "grdARTPYMTX"
                tlb_btn = DirectCast(tlb_pop.Tools("Reverse All Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode And (tabMain.SelectedTab IsNot Nothing AndAlso tabMain.SelectedTab.Key = "Applied Payments")
                If ASCMAIN1.CLIENT = "INT" Then
                    tlb_btn.SharedProps.Visible = False
                End If
        End Select

        'End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Write-Off Balance"
                If grdARTPYMT3.Selected.Rows.Count = 0 Then
                    MsgBox("No Rows Selected")
                Else

                    If MsgBox("OK to write off Non-Zero Balances for " & CStr(grdARTPYMT3.Selected.Rows.Count) & " Invoices Selected?", MsgBoxStyle.YesNo, "") = MsgBoxResult.Yes Then
                        Dim AMT As Decimal = 0
                        For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT3.Selected.Rows
                            Dim INV_BALANCE_CURR_OLD As Decimal = Val(grow.Cells("INV_BALANCE_NEW_CURR").Value & "")
                            grow.Cells("INV_PMT_CURR").Value = Val(grow.Cells("INV_PMT_CURR").Value & "") + INV_BALANCE_CURR_OLD
                            AMT += INV_BALANCE_CURR_OLD
                            grow.Update()
                        Next

                        Dim rowARTPYMT4 As DataRow = dst.Tables("ARTPYMT4").NewRow
                        rowARTPYMT4.Item("PYMT_BATCH_NO") = HFs("PYMT_BATCH_NO")
                        rowARTPYMT4.Item("PYMT_BATCH_LNO") = HFs("PYMT_BATCH_LNO")
                        rowARTPYMT4.Item("PYMT_BATCH_GLNO") = Val(dst.Tables("ARTPYMT4").Compute("Max(PYMT_BATCH_GLNO)", "") & "") + 1
                        Dim ACCT_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_ACCT_ROUNDING")
                        Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                        rowARTPYMT4.Item("ACCT_CODE") = rowGLTACCT1.Item("ACCT_CODE")
                        rowARTPYMT4.Item("ACCT_DESC") = rowGLTACCT1.Item("ACCT_DESC")
                        rowARTPYMT4.Item("GL_DIST_REF") = "Bulk W-Off"
                        For i As Integer = 2 To 4
                            Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                            rowARTPYMT4.Item(COLUMN_NAME) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                        Next

                        rowARTPYMT4.Item("GL_DIST_AMT_CURR") = AMT
                        dst.Tables("ARTPYMT4").Rows.Add(rowARTPYMT4)
                        Display_Application_Totals()

                        grdARTPYMT3.Selected.Rows.Clear()
                        MsgBox("Write-Off for " & Format(AMT, "$#,##0.00") & " has been entered into the GL Deductions grid to GL Account " & ACCT_CODE, vbOKOnly, "Verification")
                        tabARTPYMT3.SelectedTab = tabARTPYMT3.Tabs(1)
                        ' tabDeductions.SelectedTab = tabDeductions.Tabs(1)

                    End If


                End If
        End Select

        If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
            If e.Tool.Key <> "Split Line" And e.Tool.Key <> "Write-Off Balance" Then
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    Select Case e.Tool.Key
                        Case "Show Filter"
                        Case "Import From Amazon"
                            Dim RetVal As New Text.StringBuilder With {.Length = 0}
                            Dim FILENAME As String = ""
                            Using openFileDialog1 As New OpenFileDialog
                                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                                'Dim filter As String = "xlsb files (*.xlsb)|*.xlsx|All files (*.*)|*.*"
                                Dim filter As String = "All files (*.*)|*.*"
                                openFileDialog1.Filter = filter
                                openFileDialog1.RestoreDirectory = True
                                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                                    FILENAME = openFileDialog1.FileName
                                End If
                            End Using

                            If FILENAME <> "" Then

                                Dim BRANDS As New Dictionary(Of String, String)
                                BRANDS.Add("Anne Klein".ToUpper, "26")
                                BRANDS.Add("Izod Intimates".ToUpper, "12")
                                BRANDS.Add("Jessica Simpson".ToUpper, "15")
                                BRANDS.Add("Live 2 Lounge".ToUpper, "01")
                                BRANDS.Add("Lucky Brand".ToUpper, "24")
                                BRANDS.Add("Rampage".ToUpper, "02")
                                BRANDS.Add("Steve Madden".ToUpper, "27")
                                BRANDS.Add("Vince Camuto".ToUpper, "22")
                                BRANDS.Add("Z-Not Defined".ToUpper, "01")
                                Dim proceed As Boolean = False
                                ASCMAIN1.Progress("Now Building Deductions From Excel", "")
                                Cursor = Cursors.WaitCursor

                                Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
                                Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
                                Dim XWS As New Microsoft.Office.Interop.Excel.Worksheet
                                Dim WSFound As Boolean = False
                                For Each WS As Microsoft.Office.Interop.Excel.Worksheet In XWB.Worksheets
                                    If WS.Name.ToUpper = "SKU SUMMARY" Then
                                        WSFound = True
                                        XWS = WS
                                    End If
                                Next

                                Dim errMsg As New Text.StringBuilder With {.Length = 0}
                                Dim CUR_ROW As Int64 = 3
                                If WSFound = False Then
                                    errMsg.AppendLine("Can Not Find Sheet Named original.")
                                End If
                                If (XWS.Cells(CUR_ROW, 4).text.ToString & String.Empty).Trim.ToUpper <> "TOTAL UNITS" Then
                                    errMsg.AppendLine("Can Not Find Header Rows.")
                                End If
                                If errMsg.Length = 0 Then
                                    Dim BookReturns As Boolean = False
                                    Dim iResult As MsgBoxResult
                                    Dim iTitle As String = "Book Returns"
                                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                                    iMSG.AppendLine("Do You Want To Book Returns Derived")
                                    iMSG.AppendLine("From The Supplied Spreadsheet?")
                                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                                    If iResult = MsgBoxResult.Yes Then
                                        BookReturns = True
                                    End If

                                    dst.Tables.Item("ARTPYMT5").Clear()

                                    Dim PYMT_BATCH_DLNO As Int64 = 0
                                    Dim GL_DIST_AMT_CURR As Decimal
                                    Dim PYMT_BATCH_NO As String = Absx1.txtFor("PYMT_BATCH_NO").Text
                                    Dim PYMT_BATCH_LNO = Val(Absx1.numFor("PYMT_BATCH_LNO").Value & "")
                                    For i As Int64 = CUR_ROW To 100
                                        CUR_ROW += 1
                                        If BRANDS.ContainsKey((XWS.Cells(CUR_ROW, 3).text.ToString & String.Empty).Trim.ToUpper) Then
                                            'Book Advertising (AADV)
                                            GL_DIST_AMT_CURR = getValFromStr(XWS.Cells(CUR_ROW, 18).text.ToString & String.Empty)
                                            If GL_DIST_AMT_CURR <> 0 Then
                                                GL_DIST_AMT_CURR = GL_DIST_AMT_CURR * -1
                                            End If
                                            If GL_DIST_AMT_CURR <> 0 Then
                                                Dim newARTPYMT5 As DataRow = dst.Tables.Item("ARTPYMT5").NewRow
                                                PYMT_BATCH_DLNO += 1
                                                newARTPYMT5.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                                                newARTPYMT5.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                                                newARTPYMT5.Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO
                                                newARTPYMT5.Item("REASON_CODE") = "AADV"
                                                newARTPYMT5.Item("SEG2_CODE") = "000"
                                                newARTPYMT5.Item("SEG3_CODE") = "000"
                                                newARTPYMT5.Item("SEG4_CODE") = BRANDS((XWS.Cells(CUR_ROW, 3).text.ToString & String.Empty).Trim.ToUpper)
                                                newARTPYMT5.Item("GL_DIST_AMT_CURR") = GL_DIST_AMT_CURR
                                                newARTPYMT5.Item("REASON_DESC") = "AMAZONFBA ADVERTISING"
                                                newARTPYMT5.Item("TRANSACTION_LEGEND") = "DR (Expense)"
                                                dst.Tables.Item("ARTPYMT5").Rows.Add(newARTPYMT5)
                                            End If

                                            'Book Other Fees (AOTHER)
                                            GL_DIST_AMT_CURR = getValFromStr(XWS.Cells(CUR_ROW, 20).text.ToString & String.Empty)
                                            If GL_DIST_AMT_CURR <> 0 Then
                                                GL_DIST_AMT_CURR = GL_DIST_AMT_CURR * -1
                                            End If
                                            If GL_DIST_AMT_CURR <> 0 Then
                                                Dim newARTPYMT5 As DataRow = dst.Tables.Item("ARTPYMT5").NewRow
                                                PYMT_BATCH_DLNO += 1
                                                newARTPYMT5.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                                                newARTPYMT5.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                                                newARTPYMT5.Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO
                                                newARTPYMT5.Item("REASON_CODE") = "AOTHER"
                                                newARTPYMT5.Item("SEG2_CODE") = "000"
                                                newARTPYMT5.Item("SEG3_CODE") = "000"
                                                newARTPYMT5.Item("SEG4_CODE") = BRANDS((XWS.Cells(CUR_ROW, 3).text.ToString & String.Empty).Trim.ToUpper)
                                                newARTPYMT5.Item("GL_DIST_AMT_CURR") = GL_DIST_AMT_CURR
                                                newARTPYMT5.Item("REASON_DESC") = "AMAZON OTHER FEES"
                                                newARTPYMT5.Item("TRANSACTION_LEGEND") = "DR (Expense)"
                                                dst.Tables.Item("ARTPYMT5").Rows.Add(newARTPYMT5)
                                            End If

                                            'Book Optional Returns
                                            If BookReturns Then
                                                PYMT_BATCH_DLNO += 1
                                                GL_DIST_AMT_CURR = getValFromStr(XWS.Cells(CUR_ROW, 14).text.ToString & String.Empty)
                                                If GL_DIST_AMT_CURR <> 0 Then
                                                    GL_DIST_AMT_CURR = GL_DIST_AMT_CURR * -1
                                                End If
                                                If GL_DIST_AMT_CURR <> 0 Then
                                                    Dim newARTPYMT5 As DataRow = dst.Tables.Item("ARTPYMT5").NewRow
                                                    newARTPYMT5.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                                                    newARTPYMT5.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                                                    newARTPYMT5.Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO
                                                    newARTPYMT5.Item("REASON_CODE") = "RTN"
                                                    newARTPYMT5.Item("SEG2_CODE") = "000"
                                                    newARTPYMT5.Item("SEG3_CODE") = "000"
                                                    newARTPYMT5.Item("SEG4_CODE") = BRANDS((XWS.Cells(CUR_ROW, 3).text.ToString & String.Empty).Trim.ToUpper)
                                                    newARTPYMT5.Item("GL_DIST_AMT_CURR") = GL_DIST_AMT_CURR
                                                    newARTPYMT5.Item("REASON_DESC") = "RETURN"
                                                    newARTPYMT5.Item("TRANSACTION_LEGEND") = "DR (Expense)"
                                                    dst.Tables.Item("ARTPYMT5").Rows.Add(newARTPYMT5)
                                                End If
                                            End If
                                        Else
                                            'Stop 'We done?
                                        End If
                                    Next

                                    'Book Reserves (ARESD)
                                    GL_DIST_AMT_CURR = getValFromStr(XWS.Cells(2, 21).text.ToString & String.Empty) + getValFromStr(XWS.Cells(2, 22).text.ToString & String.Empty)
                                    If GL_DIST_AMT_CURR <> 0 Then
                                        GL_DIST_AMT_CURR = GL_DIST_AMT_CURR * -1
                                    End If
                                    If GL_DIST_AMT_CURR <> 0 Then
                                        Dim newARTPYMT5 As DataRow = dst.Tables.Item("ARTPYMT5").NewRow
                                        PYMT_BATCH_DLNO += 1
                                        newARTPYMT5.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                                        newARTPYMT5.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                                        newARTPYMT5.Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO
                                        newARTPYMT5.Item("REASON_CODE") = "ARESD"
                                        newARTPYMT5.Item("SEG2_CODE") = "000"
                                        newARTPYMT5.Item("SEG3_CODE") = "000"
                                        newARTPYMT5.Item("SEG4_CODE") = "00"
                                        newARTPYMT5.Item("GL_DIST_AMT_CURR") = GL_DIST_AMT_CURR
                                        newARTPYMT5.Item("REASON_DESC") = "AMAZON RESERVE DIFFERENCE"
                                        newARTPYMT5.Item("TRANSACTION_LEGEND") = "DR (Expense)"
                                        dst.Tables.Item("ARTPYMT5").Rows.Add(newARTPYMT5)
                                    End If
                                Else
                                    MsgBox(errMsg.ToString, vbExclamation, "Error(s) Importing Excel")
                                End If
                                Display_Application_Totals()
                            End If
                        Case Else
                            grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
                    End Select
                    If e.Tool.Key = "Show Filter" Then
                    Else

                    End If
                End If
            End If
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Name = "grdARTPYMTX" Then

        '    Select Case e.Tool.Key

        '    End Select
        'End If

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each row As DataRow In dst.Tables("ARTPYMTM").Select("")
                    row.Item("MATCH_ACTION") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next

            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                If CUST_CODE <> "" Then
                    Context_Launch("Select Customer", CUST_CODE, "Customer Inquiry", "ARFCINQ1")
                End If

            Case "Show Raw EDI"

                If grd IsNot Nothing AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, , "820")
                    Dim rowEDT820T1 As DataRow = LookUp("EDT820T1", EDI_DOC_SEQ_NO)
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & CUST_CODE & " Check No " & rowEDT820T1.Item("TRACE_ID")
                        frm.ShowDialog()
                    End Using
                End If


            Case "Split Line"
                If grd.ActiveRow.IsDataRow AndAlso Not grd.ActiveRow.IsAddRow Then
                    Dim VI As Integer = Val(grd.ActiveRowScrollRegion.FirstRow.Cells("PYMT_BATCH_DLNO").Value & "")

                    Dim PYMT_BATCH_DLNO As Integer = Val(grd.ActiveRow.Cells("PYMT_BATCH_DLNO").Value & "")
                    ' Dim grow As UltraWinGrid.UltraGridRow = grd.ActiveRow
                    For Each row5 As DataRow In dst.Tables("ARTPYMT5").Select("PYMT_BATCH_DLNO>" & CStr(PYMT_BATCH_DLNO), "PYMT_BATCH_DLNO DESC")
                        row5.Item("PYMT_BATCH_DLNO") = Val(row5.Item("PYMT_BATCH_DLNO") & "") + 1
                    Next

                    Dim row As DataRow = dst.Tables("ARTPYMT5").Rows.Find(New Object() {PYMT_BATCH_NO, PYMT_BATCH_LNO, PYMT_BATCH_DLNO})
                    Dim row2 As DataRow = dst.Tables("ARTPYMT5").NewRow
                    row2.ItemArray = row.ItemArray
                    row2.Item("PYMT_BATCH_DLNO") = Val(dst.Tables("ARTPYMT5").Compute("MAX(PYMT_BATCH_DLNO)", "") & "") + 1
                    row2.Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO + 1
                    row2.Item("GL_DIST_AMT") = 0
                    row2.Item("GL_DIST_AMT_CURR") = 0
                    dst.Tables("ARTPYMT5").Rows.Add(row2)

                    'grd.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

                    Sort_grdColumns(grdARTPYMT5, "PYMT_BATCH_DLNO")
                    For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT5.Rows
                        If grow.IsDataRow AndAlso Val(grow.Cells("PYMT_BATCH_DLNO").Value & "") = VI Then
                            grdARTPYMT5.ActiveRowScrollRegion.FirstRow = grow
                        End If
                        If grow.IsDataRow AndAlso Val(grow.Cells("PYMT_BATCH_DLNO").Value & "") = PYMT_BATCH_DLNO Then
                            grd.ActiveRow = grow
                            Exit For
                        End If
                    Next
                End If

            Case "Reverse All Selected"
                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("No Rows Selected")
                ElseIf grd.Selected.Rows.Count < 2 Then
                    MsgBox("Use Reverse Command in Control Panel to Reverse a Single Payment")
                Else
                    If MsgBox("Are you sure that you want to Reverse all " & CStr(grd.Selected.Rows.Count) & " Payments Selected?", MsgBoxStyle.YesNo, "Verification") = vbYes Then

                        Dim reversal_reason As String = ASCMAIN1.Get_txt_from_User("Reason", "Enter a Reason to use for All Reversals")
                        If reversal_reason = "" Then Exit Sub

                        Dim PBLs As New List(Of String)
                        For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                            Dim PYMT_BATCH_NO As String = grow.Cells("PYMT_BATCH_NO").Value
                            Dim PYMT_BATCH_LNO As Integer = Val(grow.Cells("PYMT_BATCH_LNO").Value & "")
                            PBLs.Add(PYMT_BATCH_NO & vbTab & CStr(PYMT_BATCH_LNO))
                        Next

                        grdARTPYMTX.Tag = "Reverse All Selected"
                        For Each PBL As String In PBLs
                            Dim PYMT_BATCH_NO As String = Split(PBL, vbTab)(0)
                            Dim PYMT_BATCH_LNO As Integer = Val(Split(PBL, vbTab)(1))
                            grdARTPYMTX.Selected.Rows.Clear()

                            For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMTX.Rows
                                If grow.Cells("PYMT_BATCH_NO").Value = PYMT_BATCH_NO And Val(grow.Cells("PYMT_BATCH_LNO").Value & "") = PYMT_BATCH_LNO Then
                                    grow.Activate()
                                    grow.Selected = True
                                    Exit For
                                End If
                            Next

                            Me.Cursor = Cursors.WaitCursor
                            ASCMAIN1.Progress(PBL)

                            If grdARTPYMTX.Selected.Rows.Count <> 1 Then
                                grdARTPYMTX.Tag = "Could not continue Reverse All Selected"
                                Exit For
                            End If
                            Click_Command("Reverse")
                            If Not ScreenMode Then
                                grdARTPYMTX.Tag = "Could not continue Reverse All Selected"
                                Exit For
                            End If

                            Absx1.txtFor("RETURNED_ITEM_REASON").Text = "Reverse All Selected"

                            Click_Command("Reverse Application")
                            If ScreenMode Then
                                grdARTPYMTX.Tag = "Could not continue Reverse All Selected"
                                Exit For
                            End If
                            '  MsgBox("")
                            Me.Cursor = Cursors.Default
                            ASCMAIN1.Progress("")


                            tabMain.SelectedTab = tabMain.Tabs("Applied Payments")

                            Load_grdARTPYMTX()

                        Next

                        If grdARTPYMTX.Tag = "Reverse All Selected" Then
                            MsgBox("Selected Payments have been Reversed", MsgBoxStyle.OkOnly, "Success")
                        Else
                            MsgBox(grdARTPYMTX.Tag, MsgBoxStyle.OkOnly, "Something went Wrong - reverting to Manual")
                        End If

                        grdARTPYMTX.Tag = ""
                    End If
                End If

        End Select

    End Sub

    Private Function getValFromStr(ByVal inStr As String) As Decimal
        Dim RetVal As Decimal = 0
        Dim isNeg As Boolean = False
        If inStr.Contains("(") And inStr.Contains(")") Then
            isNeg = True
        End If
        Dim REP As String() = {",", "$", "(", ")"}
        For Each R As String In REP
            inStr = inStr.Replace(R, "")
        Next
        If IsNumeric(inStr) Then
            If isNeg And Val(inStr) > 0 Then
                RetVal = Val(inStr) * -1
            Else
                RetVal = Val(inStr)
            End If
        End If
        Return RetVal
    End Function

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("New", e)
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Click_Command("New")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)
        Select Case COLUMN_NAME
            Case "CURR_CODE"
                If Absx1.txtFor("CURR_CODE").Text <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
                    txtctl.Appearance.ForeColor = Drawing.Color.Red
                Else
                    txtctl.Appearance.ForeColor = Drawing.Color.Empty
                End If
        End Select
    End Sub
#End Region

    Private Sub grdARTPYMTX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTPYMTX.DoubleClickRow
        If Not ScreenMode And e.Row.IsDataRow Then
            If tabMain.SelectedTab.Key = "Unapplied Payments" Then
                Dim EDI_DOC_SEQ_NO As String = e.Row.Cells("EDI_DOC_SEQ_NO").Value & ""
                If EDI_DOC_SEQ_NO <> "" Then
                    MsgBox("Use the 820 tab to apply EDI Payments", MsgBoxStyle.OkOnly, "Wrong tab for EDI 820 Payment Application")
                Else
                    Click_Command("Apply Payment")
                End If
            Else
                Absx1.txtFor("PYMT_BATCH_NO").Text = e.Row.Cells("PYMT_BATCH_NO").Value
                Absx1.numFor("PYMT_BATCH_LNO").Value = e.Row.Cells("PYMT_BATCH_LNO").Value
                Click_Command("View")
            End If
        End If
    End Sub

    Private Sub chkMyOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMyOnly.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ARTPYMTX()
    End Sub

    Sub Load_ARTPYMTX(Optional ByVal sql As String = "")
        If SELECTION_NO = 0 Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Payment Batch Data")

        Dim USER_ID As String = ASCMAIN1.USER_ID
        If Not chkMyOnly.Checked Then
            USER_ID = "%"
        End If

        If sql = "" Then
            'USER_ID = "melinda"
            Fill_Records("ARTPYMTX", New String() {USER_ID})
        Else
            Fill_Records("ARTPYMTX", "", True, sql)
        End If
        Sort_grdColumns(grdARTPYMTX, "PYMT_BATCH_NO,PYMT_BATCH_LNO")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdLoadRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLoadRows.Click
        Load_grdARTPYMTX()
    End Sub

    Sub Load_grdARTPYMTX()

        Dim sql As String = sqlARTPYMTX _
        & "   and ARTPYMT2.PYMT_STATUS = '2' " _
        & "   and ARTPYMT1.OPS_YYYYPP = '" & cbeYP_PYMTs.Value & "'"

        grdARTPYMTX.Text = "Payments Applied in " & cbeYP_PYMTs.Text

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE <> "" Then
            sql &= " and ARTPYMT2.CUST_CODE = '" & CUST_CODE & "'"
            grdARTPYMTX.Text &= " for Customer " & CUST_CODE
        End If

        Load_ARTPYMTX(sql)

    End Sub
    Private Sub cmdAutoApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAutoApply.Click

        ASCMAIN1.Progress("Now Automatically Applying Payment")

        Dim UNAPPLIED_AMT As Double = Val(dst.Tables("ARTPYMTT").Rows.Find("9").Item("PYMT_TOTAL_AMT") & "")
        Dim AMT_APPLIED As Double = 0

        'For Each row As DataRow In dst.Tables("ARTPYMT3") _
        ' .Select("INV_BALANCE_NEW_CURR <> 0", "INV_DATE,INV_NUM,INV_BALANCE_NEW_CURR")
        ' USE DUE DATE NOT INV DATE PER LR SO THAT 90 DAY BUCKET GETS CLEARED BEFORE 60 DAY BUCKETS
        ' USE INV DATE NOT DUE DATE PER LR SO THAT AUTO APPLY WORKS - LR SAYS RIGHT
        For Each row As DataRow In dst.Tables("ARTPYMT3") _
         .Select("INV_BALANCE_NEW_CURR <> 0", "INV_DATE,INV_NUM,INV_BALANCE_NEW_CURR")

            AMT_APPLIED = Val(row.Item("INV_BALANCE_NEW_CURR") & "")
            If AMT_APPLIED > UNAPPLIED_AMT Then
                AMT_APPLIED = UNAPPLIED_AMT
            End If

            If AMT_APPLIED <> 0 Then
                row.Item("INV_PMT_CURR") =
                 System.Math.Round(Val(row.Item("INV_PMT_CURR") & "") + AMT_APPLIED, 2)
                row.Item("INV_BALANCE_NEW_CURR") =
                 System.Math.Round(Val(row.Item("INV_BALANCE_NEW_CURR") & "") - AMT_APPLIED, 2)
                UNAPPLIED_AMT -= AMT_APPLIED
            End If
            If UNAPPLIED_AMT = 0 Then Exit For
        Next

        Calculate_Application_by_Type()

        ASCMAIN1.Progress("")
    End Sub

#Region "grdARTPYMT3"

    Private Sub grdARTPYMT3_AfterCellCancelUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT3.AfterCellCancelUpdate

    End Sub
    Private Sub grdARTPYMT3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT3.AfterCellUpdate
        If e.Cell.Column.Key = "INV_PMT_CURR" _
        Or e.Cell.Column.Key = "INV_DISC_TAKEN_CURR" _
        Or e.Cell.Column.Key = "INV_WRITE_OFF_CURR" Then
            Dim INV_BALANCE_CURR As Decimal = Val(e.Cell.Row.Cells("INV_BALANCE_CURR").Value & "")
            Dim INV_PMT_CURR As Decimal = Val(e.Cell.Row.Cells("INV_PMT_CURR").Value & "")
            Dim INV_DISC_TAKEN_CURR As Decimal = Val(e.Cell.Row.Cells("INV_DISC_TAKEN_CURR").Value & "")
            Dim INV_WRITE_OFF_CURR As Decimal = Val(e.Cell.Row.Cells("INV_WRITE_OFF_CURR").Value & "")
            Dim INV_BALANCE_NEW_CURR As Decimal = INV_BALANCE_CURR - INV_PMT_CURR - INV_DISC_TAKEN_CURR - INV_WRITE_OFF_CURR
            e.Cell.Row.Cells("INV_BALANCE_NEW_CURR").Value = INV_BALANCE_NEW_CURR
        End If
    End Sub

    Private Sub grdARTPYMT3_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTPYMT3.AfterRowUpdate
        Calculate_Application_by_Type()
    End Sub

    Private Sub grdARTPYMT3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMT3.BeforeRowUpdate

        ' SCENARIO - YOU WILL GET 2 DIFFERENT RESULTS:
        ' ENTER A VALUE IN DISCOUNTS, AND THEN 
        ' 1) CLICK ONTO THE NEXT ROW, VS 
        ' 2) CLICK ON A CONTROL TO THE RIGHT OF THE GRID CAUSING THE GRID TO LOSE FOCUS


        ' THE DATAROW APPROACH DOES NOT WORK BECAUSE THE ROW HAS NOT YET BEEN UPDATED
        ' THE .TEXT APPROACH DOES NOT WORK BECAUSE THE .TEXT PROPERTY HAS THE UNDERSCORE MASK CHARACTERS IN IT
        ' THE .VALUE APPROACH DOES NOT WORK BECAUSE IT SHOWS THE PRE-UPDATED VALUE

        'Dim row As DataRow = dst.Tables("ARTPYMT3").Rows.Find(New Object() _
        '    {e.Row.Cells("PYMT_BATCH_NO").Value, e.Row.Cells("PYMT_BATCH_LNO").Value, e.Row.Cells("PYMT_BATCH_ILNO").Value})
        'Dim INV_BALANCE_CURR As Decimal = Val(row.Item("INV_BALANCE_CURR") & "")
        'Dim INV_PMT_CURR As Decimal = Val(row.Item("INV_PMT_CURR") & "")
        'Dim INV_DISC_TAKEN_CURR As Decimal = Val(row.Item("INV_DISC_TAKEN_CURR") & "")
        'Dim INV_WRITE_OFF_CURR As Decimal = Val(row.Item("INV_WRITE_OFF_CURR") & "")
        'Dim INV_BALANCE_NEW_CURR As Decimal = INV_BALANCE_CURR - INV_PMT_CURR - INV_DISC_TAKEN_CURR - INV_WRITE_OFF_CURR

        ' this is the section that was live
        'Dim INV_BALANCE_CURR As Decimal = Val(e.Row.Cells("INV_BALANCE_CURR").value & "")
        'Dim INV_PMT_CURR As Decimal = Val(e.Row.Cells("INV_PMT_CURR").value & "")
        'Dim INV_DISC_TAKEN_CURR As Decimal = Val(e.Row.Cells("INV_DISC_TAKEN_CURR").value & "")
        'Dim INV_WRITE_OFF_CURR As Decimal = Val(e.Row.Cells("INV_WRITE_OFF_CURR").value & "")
        'Dim INV_BALANCE_NEW_CURR As Decimal = INV_BALANCE_CURR - INV_PMT_CURR - INV_DISC_TAKEN_CURR - INV_WRITE_OFF_CURR
        'e.Row.Cells("INV_BALANCE_NEW_CURR").Value = INV_BALANCE_NEW_CURR

        'Dim INV_BALANCE_NEW_CURR_CALC As Decimal = Val(e.Row.Cells("INV_BALANCE_NEW_CURR_CALC").Value & "")
        'e.Row.Cells("INV_BALANCE_NEW_CURR").Value = INV_BALANCE_NEW_CURR_CALC
    End Sub

    Private Sub grdARTPYMT3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT3.ClickCellButton
        If EntryMode = "V" Then Exit Sub

        Click_Pay()

    End Sub

    Sub Click_Pay()

        ASCMAIN1.Set_FCB("PAY", grdARTPYMT3)

        Calculate_Totals()
        Dim INV_TYPE As String = grdARTPYMT3.ActiveRow.Cells("INV_TYPE").Value
        Dim INV_TOTAL_AMOUNT_CURR As Decimal = Val(grdARTPYMT3.ActiveRow.Cells("INV_TOTAL_AMOUNT_CURR").Value & "")
        Dim INV_BALANCE_CURR As Decimal = Val(grdARTPYMT3.ActiveRow.Cells("INV_BALANCE_CURR").Value & "")
        Dim INV_FREIGHT_CURR As Decimal = Val(grdARTPYMT3.ActiveRow.Cells("INV_FREIGHT_CURR").Value & "")
        Dim INV_MISC_CHG_CURR As Decimal = Val(grdARTPYMT3.ActiveRow.Cells("INV_MISC_CHG_CURR").Value & "")
        Dim INV_PMT_CURR As Decimal = INV_BALANCE_CURR

        If optApply.Value = "I" Then

        Else
            If INV_BALANCE_CURR > 0 And INV_BALANCE_CURR > IIf(chkTAKEDISC.Checked, TOTALS.UNAPPLIED + INV_FREIGHT_CURR, TOTALS.UNAPPLIED) Then
                INV_PMT_CURR = TOTALS.UNAPPLIED
            End If
        End If

        Dim INV_DISC_TAKEN_CURR As Decimal = 0
        Dim INV_WRITE_OFF_CURR As Decimal = 0

        If chkTAKEDISC.Checked And INV_TYPE = "I" And INV_TOTAL_AMOUNT_CURR = INV_BALANCE_CURR Then
            INV_DISC_TAKEN_CURR = System.Math.Round(0.02 * INV_TOTAL_AMOUNT_CURR, 2) ' INV_MISC_CHG_CURR
            If optApply.Value = "I" Then INV_PMT_CURR -= INV_DISC_TAKEN_CURR
        End If
        If chkTAKEWOFF.Checked And INV_TYPE = "I" And INV_TOTAL_AMOUNT_CURR = INV_BALANCE_CURR Then
            INV_WRITE_OFF_CURR = System.Math.Round(0.03 * INV_TOTAL_AMOUNT_CURR, 2) '  INV_FREIGHT_CURR
            If optApply.Value = "I" Then INV_PMT_CURR -= INV_WRITE_OFF_CURR
        End If

        ' Get an application amount when doing work in application only mode
        If INV_PMT_CURR = 0 And INV_BALANCE_CURR <> 0 _
        And Val(grdARTPYMT3.ActiveRow.Cells("INV_PMT_CURR").Value & "") = 0 _
        And application_only Then
            INV_PMT_CURR = INV_BALANCE_CURR
        End If

        Dim INV_BALANCE_NEW_CURR As Decimal = INV_BALANCE_CURR - INV_PMT_CURR - INV_DISC_TAKEN_CURR - INV_WRITE_OFF_CURR

        'If Val(grdARTPYMT3.ActiveRow.Cells("INV_BALANCE_NEW_CURR").Value & "") = 0 And _
        '  Val(grdARTPYMT3.ActiveRow.Cells("INV_PMT_CURR").Value & "") <> 0 Then
        If Val(grdARTPYMT3.ActiveRow.Cells("INV_BALANCE_NEW_CURR").Value & "") <> Val(grdARTPYMT3.ActiveRow.Cells("INV_BALANCE_CURR").Value & "") Then
            grdARTPYMT3.ActiveRow.Cells("INV_PMT_CURR").Value = 0
            grdARTPYMT3.ActiveRow.Cells("INV_DISC_TAKEN_CURR").Value = 0 ' System.Math.Round(INV_DISC_TAKEN_CURR, 2)
            grdARTPYMT3.ActiveRow.Cells("INV_WRITE_OFF_CURR").Value = 0 ' System.Math.Round(INV_WRITE_OFF_CURR, 2)
            grdARTPYMT3.ActiveRow.Cells("INV_BALANCE_NEW_CURR").Value = System.Math.Round(INV_BALANCE_CURR, 2)
        Else
            grdARTPYMT3.ActiveRow.Cells("INV_PMT_CURR").Value = System.Math.Round(INV_PMT_CURR, 2)
            grdARTPYMT3.ActiveRow.Cells("INV_DISC_TAKEN_CURR").Value = System.Math.Round(INV_DISC_TAKEN_CURR, 2)
            grdARTPYMT3.ActiveRow.Cells("INV_WRITE_OFF_CURR").Value = System.Math.Round(INV_WRITE_OFF_CURR, 2)
            grdARTPYMT3.ActiveRow.Cells("INV_BALANCE_NEW_CURR").Value = System.Math.Round(INV_BALANCE_NEW_CURR, 2)
        End If
        grdARTPYMT3.UpdateData()
    End Sub
#End Region

    Sub Calculate_Totals()

        TOTALS.APPL_TOTAL = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_PMT_CURR)", "") & "")
        TOTALS.DISC_TOTAL = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_DISC_TAKEN_CURR)", "") & "")
        TOTALS.WOFF_TOTAL = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_WRITE_OFF_CURR)", "") & "")

        TOTALS.DED_TOTAL = Val(dst.Tables("ARTPYMT5").Compute("SUM (GL_DIST_AMT_CURR)", "CHARGEBACK_IND IS NULL OR CHARGEBACK_IND = '0'") & "")
        TOTALS.CHB_TOTAL = Val(dst.Tables("ARTPYMT5").Compute("SUM (GL_DIST_AMT_CURR)", "GL_DIST_AMT_CURR > 0 AND CHARGEBACK_IND = '1'") & "")
        TOTALS.OA_TOTAL = Val(dst.Tables("ARTPYMT5").Compute("SUM (GL_DIST_AMT_CURR)", "GL_DIST_AMT_CURR < 0 AND CHARGEBACK_IND = '1'") & "")
        TOTALS.GL_TOTAL = Val(dst.Tables("ARTPYMT4").Compute("SUM (GL_DIST_AMT_CURR)", "") & "")
        TOTALS.NET_AR = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_NEW_CURR)", "") & "") - Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_CURR)", "") & "") + TOTALS.CHB_TOTAL + TOTALS.OA_TOTAL ' + TOTALS.WOFF_TOTAL
        TOTALS.UNAPPLIED = System.Math.Round(CUST_PYMT_AMT_CURR - (TOTALS.APPL_TOTAL - TOTALS.DED_TOTAL - TOTALS.CHB_TOTAL - TOTALS.OA_TOTAL - TOTALS.GL_TOTAL), 2)

    End Sub

    Sub Display_Application_Totals()
        Calculate_Totals()
        With dst.Tables("ARTPYMTT").Rows
            .Find("1").Item("PYMT_TOTAL_AMT") = TOTALS.APPL_TOTAL
            .Find("2").Item("PYMT_TOTAL_AMT") = TOTALS.DISC_TOTAL
            .Find("3").Item("PYMT_TOTAL_AMT") = TOTALS.WOFF_TOTAL
            .Find("4").Item("PYMT_TOTAL_AMT") = TOTALS.DED_TOTAL
            .Find("5").Item("PYMT_TOTAL_AMT") = TOTALS.CHB_TOTAL
            .Find("6").Item("PYMT_TOTAL_AMT") = TOTALS.GL_TOTAL
            .Find("7").Item("PYMT_TOTAL_AMT") = TOTALS.OA_TOTAL * -1
            .Find("8").Item("PYMT_TOTAL_AMT") = TOTALS.NET_AR
            .Find("9").Item("PYMT_TOTAL_AMT") = TOTALS.UNAPPLIED
            cmdLeaveOA.Visible = (TOTALS.UNAPPLIED > 0)
        End With
        grdARTPYMTT.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        If ScreenMode Then
            If TOTALS.UNAPPLIED < 0 Then
                ' MsgBox("You are now out of Funds.", MsgBoxStyle.OkOnly, "Notification")
            End If
        End If
    End Sub

    Sub Calculate_Application_by_Type()
        Display_Application_Totals()
        For Each row As DataRow In dst.Tables("ARTPYMTA").Rows
            Dim AR_TYPE_CODE As String = row.Item("AR_TYPE_CODE")
            Dim ITEMS = Val(dst.Tables("ARTPYMT3").Compute("COUNT (INV_NUM)", "INV_TYPE = '" & AR_TYPE_CODE & "'") & "")
            Dim INV_BALANCE_NEW_CURR = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_NEW_CURR)", "INV_TYPE = '" & AR_TYPE_CODE & "'") & "")
            Dim INV_BALANCE_CURR = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_CURR)", "INV_TYPE = '" & AR_TYPE_CODE & "'") & "")
            Dim AR_TYPE_AMT = INV_BALANCE_CURR - INV_BALANCE_NEW_CURR
            'Dim AR_TYPE_AMT = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_PMT_CURR)", "INV_TYPE = '" & AR_TYPE_CODE & "'") & "")
            row("ITEMS") = ITEMS
            row("AR_TYPE_AMT_OLD") = INV_BALANCE_CURR
            row("AR_TYPE_AMT") = AR_TYPE_AMT
            row("AR_TYPE_AMT_NEW") = INV_BALANCE_NEW_CURR
        Next

        For i As Integer = 1 To 4
            Dim AGE_AMT As Decimal = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_NEW_CURR)", "AGE_BUCKET = " & CStr(i)) & "")
            Dim rowARTOPENA As DataRow = dst.Tables("ARTOPENA").Rows.Find(i)
            rowARTOPENA.Item("AGE_AMT_NEW") = AGE_AMT
        Next
    End Sub

    Private Sub optShowItems_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShowItems.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub

        Select Case optShowItems.Value
            Case "A"
                dst.Tables("ARTPYMT3").DefaultView.RowFilter = ""
            Case "O"
                dst.Tables("ARTPYMT3").DefaultView.RowFilter = "INV_BALANCE_CURR <> 0"
            Case "P"
                dst.Tables("ARTPYMT3").DefaultView.RowFilter = "INV_BALANCE_NEW_CURR <> INV_BALANCE_CURR"
        End Select
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_Control_Panel()
    End Sub

    Sub Setup_Control_Panel()
        If Me.SELECTION_NO = 0 Then Exit Sub

        With UltraExplorerBar1
            .Groups("Unapplied Payment Options").Visible = Not ScreenMode And (tabMain.SelectedTab.Key = "Unapplied Payments") And (EntryMode <> "S")
            .Groups("Lock Box Options").Visible = (tabMain.SelectedTab.Text = "Lock-Box Receipts (EDI)")
            .Groups("EDI (820)").Visible = (tabMain.SelectedTab.Text = "EDI (820)")
            .Groups("Screen Control").Visible = (tabMain.SelectedTab.Key = "Unapplied Payments") And (EntryMode <> "S")
            .Groups("Credit Card Options").Visible = (tabMain.SelectedTab.Text = "Credit Cards (Settled)")
            .Groups("Split Payment Options").Visible = (EntryMode = "S")
            .Groups("Find Applied Payments").Visible = Not ScreenMode And (tabMain.SelectedTab.Key = "Applied Payments") And (EntryMode <> "S")
            .Groups("Post Application Options").Visible = (tabMain.SelectedTab.Key = "Applied Payments")
        End With

        If tabMain.SelectedTab.Key = "Unapplied Payments" Then
            If EntryMode = "S" Then
                grdARTPYMTX.Text = "Payment to Split"
            Else
                Load_ARTPYMTX()
                grdARTPYMTX.Text = "Double-Click a Payment to begin Application"
                splARTPYMT2.Parent = tabMain.SelectedTab.TabPage
            End If
        ElseIf tabMain.SelectedTab.Key = "Applied Payments" Then
            dst.Tables("ARTPYMTX").Rows.Clear()
            grdARTPYMTX.Text = "Select a Month, and (optionally) a Customer, and then click Fetch Payments"
            splARTPYMT2.Parent = tabMain.SelectedTab.TabPage
        End If
    End Sub

    Sub Print_Receipt()

        Print_Report_Begin()
        'CR_params.Add("SHOW_CVX_NAME", "0")
        Generate_Report("ARRPYMTR", "Lock-Box Receipts")
        Print_Report_End()

        tabLockBoxDetails.Tag = "Y"
    End Sub

    Sub Show_AR_Item_Columns(Optional ByVal initialize As Boolean = False)
        With grdARTPYMT3.DisplayLayout.Bands("ARTPYMT3")
            Dim P As UltraWinToolbars.PopupMenuTool = tlb.Tools("grdARTPYMT3")
            For Each tT As UltraWinToolbars.ToolBase In P.Tools
                If tT.Key <> "Write-Off Balance" Then
                    Dim T As UltraWinToolbars.StateButtonTool = DirectCast(tT, UltraWinToolbars.StateButtonTool)
                    If initialize Then
                        If New String() {"INV_DUE_DATE", "POST_CODE", "REASON_CODE", "INV_CUST_PO"}.Contains(T.Key) Then
                            T.Checked = True
                        End If
                    End If
                    If T.Key = "Show Filter" Then
                    Else
                        .Columns(T.Key).Hidden = Not T.Checked
                    End If
                End If
            Next
        End With
    End Sub

    Sub Remove_Zeros(ByVal TABLE_NAME As String, ByVal keep_rows_if As String)
        Dim dvw As DataView = New DataView(dst.Tables(TABLE_NAME))
        dvw.RowFilter = keep_rows_if
        Dim tbl As DataTable = dvw.ToTable
        dst.Tables(TABLE_NAME).Rows.Clear()
        dst.Tables(TABLE_NAME).Merge(tbl)
    End Sub

#Region "grdARTPYMT4"

    Private Sub grdARTPYMT4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT4.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdARTPYMT4, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next

                If rowSOTCHAN1 IsNot Nothing AndAlso rowSOTCHAN1.Item("SEG2_CODE") & "" <> "" Then
                    e.Cell.Row.Cells("SEG2_CODE").Value = rowSOTCHAN1.Item("SEG2_CODE")
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    'Dim ACCT_CODE As String = e.Cell.Row.Cells("ACCT_CODE").Value & ""
                    'If ACCT_CODE <> "" Then
                    '    Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                    'End If
                    ' MAYBE WE USE GLTACCT1 TO CHECK SEG3 REQD?
                    If rowSOTTCLS1 IsNot Nothing Then
                        Dim SEG3_CODE As String = rowSOTTCLS1.Item("SEG3_CODE") & ""
                        If SEG3_CODE = "" Then SEG3_CODE = rowSOTTCLS1.Item("TRADE_CLASS_CODE") & ""
                        e.Cell.Row.Cells("SEG3_CODE").Value = SEG3_CODE
                    End If
                End If

        End Select
    End Sub

    Private Sub grdARTPYMT4_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT4.AfterExitEditMode
        With grdARTPYMT4
            Select Case .ActiveCell.Column.Key
                Case "ACCT_CODE"
                    Dim ACCT_CODE As String = .ActiveCell.Text
                    If ACCT_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdARTPYMT4_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT4.AfterRowActivate
        With grdARTPYMT4
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdARTPYMT4.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdARTPYMT4_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT4.AfterRowsDeleted
        Display_Application_Totals()
    End Sub

    Private Sub grdARTPYMT4_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTPYMT4.AfterRowUpdate
        Display_Application_Totals()
    End Sub

    Private Sub grdARTPYMT4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMT4.BeforeRowUpdate
        With grdARTPYMT4
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is not Active", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                    If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is a Control Account - no Manual J/E permitted", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If

                    Dim DIST_APP_CODE As String = "AR"
                    If LookUp("GLTDSTR1", DIST_APP_CODE) IsNot Nothing AndAlso cdr.Item("DIST_APP_STATUS") & "" = "A" Then
                        If LookUp("GLTDSTR2", New String() {DIST_APP_CODE, e.Row.Cells("ACCT_CODE").Text}) Is Nothing Then
                            MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is not permitted for Posting in this Application (" & DIST_APP_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If

                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("PYMT_BATCH_NO").Text = "" Then
                    .ActiveRow.Cells("PYMT_BATCH_NO").Value = HFs("PYMT_BATCH_NO")
                    .ActiveRow.Cells("PYMT_BATCH_LNO").Value = HFs("PYMT_BATCH_LNO")
                    .ActiveRow.Cells("PYMT_BATCH_GLNO").Value = Val(dst.Tables("ARTPYMT4").Compute("Max(PYMT_BATCH_GLNO)", "") & "") + 1
                End If
                If ASCMAIN1.CLIENT = "AHA" Then
                    If rowSOTCHAN1 IsNot Nothing AndAlso rowSOTCHAN1.Item("SEG2_CODE") & "" <> "" Then
                        .ActiveRow.Cells("SEG2_CODE").Value = rowSOTCHAN1.Item("SEG2_CODE")
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub grdARTPYMT4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT4.ClickCellButton
        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim DIST_APP_CODE As String = "AR"

                If LookUp("GLTDSTR1", DIST_APP_CODE) IsNot Nothing AndAlso cdr.Item("DIST_APP_STATUS") & "" = "A" Then
                    sql_where = "ACCT_CODE in (Select ACCT_CODE from GLTDSTR2 where DIST_APP_CODE = '" & DIST_APP_CODE & "')"
                End If

            Case "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"

        End Select

        grdClickCellButton(grdARTPYMT4, sql_where, sql_where <> "")
    End Sub

    Private Sub grdARTPYMT4_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdARTPYMT4.DoubleClickCell
        If e.Cell.Column.Key = "GL_DIST_AMT_CURR" Then
            If grdARTPYMT4.ActiveCell Is Nothing Then
            Else
                'grdARTPYMT4.ActiveCell.Value = Val(grdARTPYMT4.ActiveCell.Value & "") + Get_TOTAL_UNAPPLIED()
                'grdARTPYMT4.UpdateData()

                grdARTPYMT4.ActiveCell.Value = 0
                grdARTPYMT4.ActiveRow.Update()
                grdARTPYMT4.ActiveCell.Value = -1 * Get_TOTAL_UNAPPLIED()
                grdARTPYMT4.ActiveRow.Update()

            End If
        End If

    End Sub

    Private Sub grdARTPYMT4_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTPYMT4.Error
        grdARTPYMT4.ActiveRow.CancelUpdate()
    End Sub

#End Region

#Region "grdARTPYMT5"

    Private Sub grdARTPYMT5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT5.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "REASON_CODE"
                Dim REASON_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdARTPYMT5, "ARTREAS1", "REASON_CODE", "REASON_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next

                If rowSOTCHAN1 IsNot Nothing AndAlso rowSOTCHAN1.Item("SEG2_CODE") & "" <> "" Then
                    e.Cell.Row.Cells("SEG2_CODE").Value = rowSOTCHAN1.Item("SEG2_CODE")
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    'Dim ACCT_CODE As String = e.Cell.Row.Cells("ACCT_CODE").Value & ""
                    'If ACCT_CODE <> "" Then
                    '    Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                    'End If
                    ' MAYBE WE USE GLTACCT1 TO CHECK SEG3 REQD?
                    If rowSOTTCLS1 IsNot Nothing Then
                        Dim SEG3_CODE As String = rowSOTTCLS1.Item("SEG3_CODE") & ""
                        If SEG3_CODE = "" Then SEG3_CODE = rowSOTTCLS1.Item("TRADE_CLASS_CODE") & ""
                        e.Cell.Row.Cells("SEG3_CODE").Value = SEG3_CODE
                    End If

                    Dim INV_NOs As New List(Of String)
                    For Each row As DataRow In dst.Tables("ARTPYMT3").Select("INV_TYPE = 'I'", "INV_PMT DESC, INV_BALANCE DESC")
                        INV_NOs.Add(row.Item("INV_NUM"))
                        If INV_NOs.Count >= 10 Then Exit For
                    Next
                    Dim BRAND_CODE As String = "AV"
                    If INV_NOs.Count <> 0 Then
                        ASCMAIN1.sql = "Select BRAND_CODE from (" & vbCrLf _
                            & "SELECT BRAND_CODE FROM (" & vbCrLf _
                            & "Select ICTCOLL1.BRAND_CODE, SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) AMT" & vbCrLf _
                            & " FROM SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
                            & " WHERE INV_TYPE = 'I' AND INV_NO IN ('" & Join(INV_NOs.ToArray, "','") & "')" & vbCrLf _
                            & "   AND ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                            & "   AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                            & " GROUP BY ICTCOLL1.BRAND_CODE" & vbCrLf _
                            & ") ORDER BY AMT DESC" & vbCrLf _
                            & ") WHERE ROWNUM < 2"
                        BRAND_CODE = ASCDATA1.GetDataValue
                    End If
                    If BRAND_CODE <> "" Then e.Cell.Row.Cells("SEG4_CODE").Value = BRAND_CODE


                End If
        End Select
    End Sub

    Private Sub grdARTPYMT5_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT5.AfterExitEditMode
        With grdARTPYMT5
            Select Case .ActiveCell.Column.Key
                Case "REASON_CODE"
                    Dim REASON_CODE As String = .ActiveCell.Text
                    If REASON_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(REASON_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdARTPYMT5_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT5.AfterRowActivate
        With grdARTPYMT5
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("REASON_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdARTPYMT5.ActiveRow.Cells("REASON_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("REASON_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdARTPYMT5_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT5.AfterRowsDeleted
        Display_Application_Totals()
    End Sub

    Private Sub grdARTPYMT5_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTPYMT5.AfterRowUpdate
        Display_Application_Totals()
    End Sub

    Private Sub grdARTPYMT5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMT5.BeforeRowUpdate
        With grdARTPYMT5
            If e.Row.Cells("REASON_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("ARTREAS1", e.Row.Cells("REASON_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Reason Code (" & e.Row.Cells("REASON_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("PYMT_BATCH_NO").Text = "" Then
                    .ActiveRow.Cells("PYMT_BATCH_NO").Value = HFs("PYMT_BATCH_NO")
                    .ActiveRow.Cells("PYMT_BATCH_LNO").Value = HFs("PYMT_BATCH_LNO")
                    .ActiveRow.Cells("PYMT_BATCH_DLNO").Value = Val(dst.Tables("ARTPYMT5").Compute("Max(PYMT_BATCH_DLNO)", "") & "") + 1
                End If
                If (Val(e.Row.Cells("GL_DIST_AMT_CURR").Text & "") > 0 And e.Row.Cells("CHARGEBACK_IND").Text & "" = "1") And e.Row.Cells("CUST_CODE_SO").Text = "" Then
                    .ActiveRow.Cells("CUST_CODE_SO").Value = HFs("CUST_CODE")
                End If

                If .ActiveRow.Cells("CHARGEBACK_IND").Text & "" = "1" Then
                    .ActiveRow.Cells("SEG2_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                Else
                    If rowSOTCHAN1 IsNot Nothing AndAlso rowSOTCHAN1.Item("SEG2_CODE") & "" <> "" Then
                        .ActiveRow.Cells("SEG2_CODE").Value = rowSOTCHAN1.Item("SEG2_CODE")
                    End If
                End If

            End If
        End With
    End Sub

    Private Sub grdARTPYMT5_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT5.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdARTPYMT5, sql_where, sql_where <> "")
    End Sub

    Private Sub grdARTPYMT5_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdARTPYMT5.DoubleClickCell
        If e.Cell.Column.Key = "GL_DIST_AMT_CURR" Then
            If grdARTPYMT5.ActiveCell Is Nothing Then
            Else
                'grdARTPYMT5.ActiveCell.Value = Val(grdARTPYMT5.ActiveCell.Value & "") + Get_TOTAL_UNAPPLIED()
                'grdARTPYMT5.UpdateData()
                grdARTPYMT5.ActiveCell.Value = 0
                grdARTPYMT5.ActiveRow.Update()
                grdARTPYMT5.ActiveCell.Value = -1 * Get_TOTAL_UNAPPLIED()
                grdARTPYMT5.ActiveRow.Update()
            End If
        End If
    End Sub

    Private Sub grdARTPYMT5_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTPYMT5.Error
        grdARTPYMT5.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Function Get_TOTAL_UNAPPLIED() As Double

        If Absx1.txtFor("CUST_CODE").Text = "" Then
            Dim TOTAL4 As Decimal = Val(dst.Tables("ARTPYMT4").Compute("SUM(GL_DIST_AMT_CURR)", "") & "")
            Return System.Math.Round(CUST_PYMT_AMT_CURR + TOTAL4, 2)
            'Return System.Math.Round(CUST_PYMT_AMT_CURR + TOTAL4, 2)
        Else
            Return System.Math.Round(Val(dst.Tables("ARTPYMTT").Rows.Find("9").Item("PYMT_TOTAL_AMT") & ""), 2)
        End If
        'Return System.Math.Round(Val(dst.Tables("ARTPYMTT").Rows.Find("9").Item("PYMT_TOTAL_AMT") & ""), 2)
    End Function

    Sub Load_Open_AR_from_CB(
    ByVal rowARTPYMT5 As DataRow,
    ByVal rowARTPYMT2 As DataRow,
    ByVal PYMT_BATCH_DATE As Date)

        Dim GL_DIST_AMT As Decimal = Val(rowARTPYMT5.Item("GL_DIST_AMT") & "")
        Dim GL_DIST_AMT_CURR As Decimal = Val(rowARTPYMT5.Item("GL_DIST_AMT_CURR") & "")
        Dim INV_TYPE_CB As String = rowARTPYMT5.Item("INV_TYPE_CB") & ""
        If INV_TYPE_CB = "" Then
            If GL_DIST_AMT < 0 Then
                INV_TYPE_CB = "O"
            Else
                INV_TYPE_CB = "B"
            End If
            rowARTPYMT5.Item("INV_TYPE_CB") = INV_TYPE_CB
        End If

        Dim CHARGEBACK_NO As String = ASCMAIN1.Next_Control_No("INV_NUM_" & INV_TYPE_CB)
        rowARTPYMT5.Item("CHARGEBACK_NO") = CHARGEBACK_NO

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        rowARTOPEN1.Item("CUST_CODE") = rowARTPYMT2.Item("CUST_CODE")
        rowARTOPEN1.Item("INV_TYPE") = INV_TYPE_CB
        rowARTOPEN1.Item("INV_NUM") = CHARGEBACK_NO
        rowARTOPEN1.Item("INV_DATE") = PYMT_BATCH_DATE
        rowARTOPEN1.Item("INV_DUE_DATE") = rowARTOPEN1.Item("INV_DATE")
        rowARTOPEN1.Item("INV_CUST_PO") = rowARTPYMT5.Item("CUST_REFERENCE")
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT") = GL_DIST_AMT
        rowARTOPEN1.Item("INV_BALANCE") = GL_DIST_AMT
        rowARTOPEN1.Item("REASON_CODE") = rowARTPYMT5.Item("REASON_CODE")

        rowARTOPEN1.Item("INV_LAST_PMT_REF") = rowARTPYMT2.Item("CUST_PYMT_REF_NO")
        rowARTOPEN1.Item("INV_LAST_PMT_REF_DT") = rowARTPYMT2.Item("CUST_PYMT_REF_DATE")
        rowARTOPEN1.Item("INV_NOTES") = rowARTPYMT5.Item("OUR_REFERENCE")

        Dim CC_LOOKUP As String = ""
        If rowARTPYMT5.Item("CUST_CODE_SO") & "" <> "" Then
            rowARTOPEN1.Item("CUST_CODE_SO") = rowARTPYMT5.Item("CUST_CODE_SO")
            CC_LOOKUP = rowARTPYMT5.Item("CUST_CODE_SO")
        Else
            CC_LOOKUP = rowARTPYMT2.Item("CUST_CODE")
        End If
        LookUp("ARTCUST1", CC_LOOKUP)
        rowARTOPEN1.Item("SREP_CODE") = cdr.Item("SREP_CODE") & ""
        rowARTOPEN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowARTOPEN1.Item("INIT_DATE") = DATETIME_STAMP
        rowARTOPEN1.Item("INV_MISC_CHG") = GL_DIST_AMT
        rowARTOPEN1.Item("CURR_CODE") = Absx1.txtFor("CURR_CODE").Text   ' "USD" ' CURR_CODE
        rowARTOPEN1.Item("CURR_EXCH_RATE") = Val(Absx1.numFor("CURR_EXCH_RATE").Value & "")  ' 1 ' CURR_EXCH_RATE
        rowARTOPEN1.Item("INV_MISC_CHG_CURR") = GL_DIST_AMT_CURR
        rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = GL_DIST_AMT_CURR
        rowARTOPEN1.Item("INV_BALANCE_CURR") = GL_DIST_AMT_CURR


        Dim ORDR_TYPE_CODE As String
        Dim rowSOTTYPE1 As DataRow
        If INV_TYPE_CB = "O" Then
            ORDR_TYPE_CODE = ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_OA")
        Else
            ORDR_TYPE_CODE = ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_CB")
        End If
        rowSOTTYPE1 = dst.Tables("SOTTYPE1").Rows.Find(ORDR_TYPE_CODE)

        Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE")
        Dim rowARTPOST1 As DataRow = dst.Tables("ARTPOST1").Rows.Find(POST_CODE)

        rowARTOPEN1.Item("SEG2_CODE") = rowARTPOST1.Item("SEG2_CODE")
        rowARTOPEN1.Item("SEG3_CODE") = rowARTPOST1.Item("SEG3_CODE")
        rowARTOPEN1.Item("SEG4_CODE") = rowARTPOST1.Item("SEG4_CODE")

        rowARTOPEN1.Item("POST_CODE") = POST_CODE
        rowARTOPEN1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE_0")

        rowARTOPEN1.Item("ORDR_TYPE_CODE") = ORDR_TYPE_CODE
        'rowARTOPEN1.Item("INV_REF") = rowARTPYMT5.Item("OUR_REFERENCE")
        rowARTOPEN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)

        rowARTPYMT5.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
        rowARTPYMT5.Item("SEG2_CODE") = rowARTPOST1.Item("SEG2_CODE")
        rowARTPYMT5.Item("SEG3_CODE") = rowARTPOST1.Item("SEG3_CODE")
        rowARTPYMT5.Item("SEG4_CODE") = rowARTPOST1.Item("SEG4_CODE")

    End Sub

    Private Sub cmdApplyAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApplyAll.Click
        ApplyAll(True)
    End Sub

    Private Sub cmdUnApplyAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnApplyAll.Click
        ApplyAll(False)
    End Sub

    Sub ApplyAll(ByVal apply As Boolean)

        ASCMAIN1.Progress("Now Automatically Applying All")

        Dim INV_PMT_CURR As Decimal = 0

        For Each row As DataRow In dst.Tables("ARTPYMT3").Select("", "")
            If apply Then
                INV_PMT_CURR = Val(row.Item("INV_BALANCE_CURR") & "")
            Else
                INV_PMT_CURR = 0
            End If

            row.Item("INV_PMT_CURR") = INV_PMT_CURR
            row.Item("INV_DISC_TAKEN_CURR") = 0
            row.Item("INV_WRITE_OFF_CURR") = 0
            Dim INV_BALANCE_CURR As Decimal = Val(row.Item("INV_BALANCE_CURR") & "")
            Dim INV_BALANCE_NEW_CURR As Decimal = INV_BALANCE_CURR - INV_PMT_CURR
            row.Item("INV_BALANCE_NEW_CURR") = INV_BALANCE_NEW_CURR
        Next

        Calculate_Application_by_Type()

        ASCMAIN1.Progress("")
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub grdARTPYMTT_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMTT.InitializeRow
        Dim K As String = e.Row.Cells("PYMT_TOTAL_CODE").Value & ""
        If K = "9" Then
            If Val(e.Row.Cells("PYMT_TOTAL_AMT").Value & "") = 0 Then
                e.Row.Cells("PYMT_TOTAL_AMT").Appearance.BackColor = Drawing.Color.LightGreen
            Else
                e.Row.Cells("PYMT_TOTAL_AMT").Appearance.BackColor = Drawing.Color.Yellow
            End If
            cmdLeaveOA.Visible = (Val(e.Row.Cells("PYMT_TOTAL_AMT").Value & "") > 0)
        End If
    End Sub

    Private Sub cmdLeaveOA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLeaveOA.Click
        Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
        rowARTPYMT5.Item("PYMT_BATCH_NO") = HFs("PYMT_BATCH_NO")
        rowARTPYMT5.Item("PYMT_BATCH_LNO") = HFs("PYMT_BATCH_LNO")
        rowARTPYMT5.Item("PYMT_BATCH_DLNO") = Val(dst.Tables("ARTPYMT5").Compute("Max(PYMT_BATCH_DLNO)", "") & "") + 1

        Dim rowSOTTYPE1 As DataRow = dst.Tables("SOTTYPE1").Rows.Find(ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_OA") & "")
        rowARTPYMT5.Item("REASON_CODE") = rowSOTTYPE1.Item("REASON_CODE")

        Dim rowARTREAS1 As DataRow = dst.Tables("ARTREAS1").Rows.Find(rowSOTTYPE1.Item("REASON_CODE") & "")
        rowARTPYMT5.Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
        For i As Integer = 2 To 4
            Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
            If rowARTREAS1.Item(COLUMN_NAME) & "" <> "" Then
                rowARTPYMT5.Item(COLUMN_NAME) = rowARTREAS1.Item(COLUMN_NAME)
            Else
                rowARTPYMT5.Item(COLUMN_NAME) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
            End If
        Next
        Dim AMT As Decimal = Val(dst.Tables("ARTPYMTT").Rows.Find("9").Item("PYMT_TOTAL_AMT") & "")
        rowARTPYMT5.Item("GL_DIST_AMT_CURR") = -1 * AMT
        rowARTPYMT5.Item("CHARGEBACK_IND") = "1"
        rowARTPYMT5.Item("CUST_REFERENCE") = Absx1.txtFor("CUST_PYMT_REF_NO").Text & ""
        dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)
        Display_Application_Totals()
    End Sub

    Sub Setup_Split_Payment()

        dst.Tables("ARTPYMT2_SPLIT").Rows.Clear()

        dst.Tables("ARTPYMTX").AcceptChanges()

        Dim dv As DataView = DirectCast(grdARTPYMTX.DataSource, DataTable).DefaultView
        PYMT_BATCH_NO = grdARTPYMTX.Selected.Rows(0).Cells("PYMT_BATCH_NO").Text
        PYMT_BATCH_LNO = Val(grdARTPYMTX.Selected.Rows(0).Cells("PYMT_BATCH_LNO").Text)
        dv.RowFilter = "PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "' and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)

        'PYMT_BATCH_NO_new = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        'PYMT_BATCH_LNO_new = 0

        PYMT_BATCH_NO_new = PYMT_BATCH_NO
        ASCMAIN1.sql = "Select Max(PYMT_BATCH_LNO) from ARTPYMT2 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
        PYMT_BATCH_LNO_new = Val(ASCDATA1.GetDataValue)

        Dim rowARTPYMTX As DataRow = dst.Tables("ARTPYMTX").Rows.Find(New Object() {PYMT_BATCH_NO, PYMT_BATCH_LNO})

        grdARTPYMT2_SPLIT.Text = "Splitting Batch " & PYMT_BATCH_NO & " Line " & CStr(PYMT_BATCH_LNO) & ", " & rowARTPYMTX.Item("CUST_NAME") & " Payment of " & Format(Val(rowARTPYMTX.Item("CUST_PYMT_AMT_CURR") & ""), "$#,##0.00")

        rowARTPYMTX.Item("PYMT_DELETED") = "1"
        rowARTPYMTX.Item("PYMT_STATUS") = "2"
        rowARTPYMTX.Item("LAST_DATE") = DATETIME_STAMP
        rowARTPYMTX.Item("LAST_OPER") = ASCMAIN1.USER_ID
    End Sub

#Region "grdARTPYMT2_SPLIT"
    Private Sub grdARTPYMT2_SPLIT_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT2_SPLIT.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                If e.Cell.Text = "" Then
                Else
                    grdCodeDesc(grdARTPYMT2_SPLIT, "ARTCUST1", "CUST_CODE", "CUST_NAME")
                    If grdARTPYMT2_SPLIT.ActiveRow.Cells("CUST_NAME").Text = "" Then
                        grdARTPYMT2_SPLIT.PerformAction(UltraWinGrid.UltraGridAction.PrevCell)
                    Else
                        grdARTPYMT2_SPLIT.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                    End If
                    grdARTPYMT2_SPLIT.ActiveRow.Cells("CUST_PYMT_REF_NO").Value = grdARTPYMTX.ActiveRow.Cells("CUST_PYMT_REF_NO").Value
                    grdARTPYMT2_SPLIT.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value = grdARTPYMTX.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value
                End If
        End Select
    End Sub

    Private Sub grdARTPYMT2_SPLIT_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT2_SPLIT.AfterRowActivate
        With grdARTPYMT2_SPLIT.DisplayLayout.Bands(0)
            If grdARTPYMT2_SPLIT.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdARTPYMT2_SPLIT.ActiveCell = grdARTPYMT2_SPLIT.ActiveRow.Cells("CUST_CODE")
                grdARTPYMT2_SPLIT.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdARTPYMT2_SPLIT_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMT2_SPLIT.BeforeRowUpdate
        With grdARTPYMT2_SPLIT
            If e.Row.Cells("CUST_CODE").Text = "" Then
                MsgBox("Missing Value for Customer Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Customer Code (" & e.Row.Cells("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PYMT_BATCH_NO").Text = "" Then
                    .ActiveRow.Cells("PYMT_BATCH_NO").Value = PYMT_BATCH_NO_new
                    PYMT_BATCH_LNO_new += 1
                    .ActiveRow.Cells("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO_new
                    .ActiveRow.Cells("PYMT_STATUS").Value = "1"

                    .ActiveRow.Cells("CUST_CODE_ORIG").Value = grdARTPYMTX.ActiveRow.Cells("CUST_CODE").Value
                    .ActiveRow.Cells("CURR_CODE").Value = grdARTPYMTX.ActiveRow.Cells("CURR_CODE").Value
                    .ActiveRow.Cells("CURR_EXCH_RATE").Value = grdARTPYMTX.ActiveRow.Cells("CURR_EXCH_RATE").Value
                    .ActiveRow.Cells("PYMT_BATCH_NO_ORIG").Value = PYMT_BATCH_NO
                    .ActiveRow.Cells("PYMT_BATCH_LNO_ORIG").Value = PYMT_BATCH_LNO

                    LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Value)
                    e.Row.Cells("CUST_NAME").Value = cdr.Item("CUST_NAME")
                End If
                .ActiveRow.Cells("CUST_PYMT_AMT").Value = .ActiveRow.Cells("CUST_PYMT_AMT_CURR").Text
            End If
        End With
    End Sub

    Private Sub grdARTPYMT2_SPLIT_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT2_SPLIT.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdARTPYMT2_SPLIT, sql_where, sql_where <> "")
    End Sub

    Private Sub grdARTPYMT2_SPLIT_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTPYMT2_SPLIT.Error
        grdARTPYMT2_SPLIT.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdARTPYMT2_SPLIT_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTPYMT2_SPLIT.BeforeExitEditMode
        With grdARTPYMT2_SPLIT.ActiveCell
            Select Case .Column.Key
                Case "CUST_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    End If
            End Select
        End With
    End Sub
#End Region

    Sub Update_Split()

        Dim rowARTPYMTX As DataRow = dst.Tables("ARTPYMTX").Rows.Find(New Object() {PYMT_BATCH_NO, PYMT_BATCH_LNO})
        Dim rowARTPYMT2_SPLIT As DataRow = dst.Tables("ARTPYMT2_SPLIT").NewRow
        For I As Integer = 0 To dst.Tables("ARTPYMT2_SPLIT").Columns.Count - 1
            rowARTPYMT2_SPLIT.Item(I) = rowARTPYMTX.Item(I)
        Next

        rowARTPYMT2_SPLIT.Item("PYMT_NOTE") = PYMT_NOTE
        rowARTPYMT2_SPLIT.Item("PYMT_DELETED") = "1"
        rowARTPYMT2_SPLIT.Item("PYMT_STATUS") = "2"
        rowARTPYMT2_SPLIT.Item("LAST_DATE") = DATETIME_STAMP
        rowARTPYMT2_SPLIT.Item("LAST_OPER") = ASCMAIN1.USER_ID

        dst.Tables("ARTPYMT2_SPLIT").Rows.Add(rowARTPYMT2_SPLIT)
        rowARTPYMT2_SPLIT.AcceptChanges()
        rowARTPYMT2_SPLIT.SetModified()

        BeginTrans()
        Update_Record_TDA("ARTPYMT2_SPLIT")

        ' these updates happen after applying the cash - incorrect to do this now
        'For Each row As DataRow In dst.Tables("ARTPYMT2_SPLIT").Rows
        '    ASCDATA1.ExecuteSP("ARPCUST6_PYMT", "VN" _
        '           , New Object() {row.Item("PYMT_BATCH_NO"), row.Item("PYMT_BATCH_LNO")} _
        '           , New String() {"PYMT_BATCH_NO_IN", "PYMT_BATCH_LNO_IN"})
        'Next

        CommitTrans("Split Payment Update Complete")
    End Sub

    Private Sub grdARTPYMT3_ClientSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT3.ClientSizeChanged

    End Sub

    Private Sub grdARTPYMT3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMT3.InitializeRow
        Dim INV_PMT As Decimal = Val(e.Row.Cells("INV_PMT_CURR").Value & "")
        Dim INV_DISC_TAKEN As Decimal = Val(e.Row.Cells("INV_DISC_TAKEN_CURR").Value & "")
        Dim INV_WRITE_OFF As Decimal = Val(e.Row.Cells("INV_WRITE_OFF_CURR").Value & "")
        Dim INV_BALANCE_NEW_CURR As Decimal = Val(e.Row.Cells("INV_BALANCE_NEW_CURR").Value & "")
        Dim CURR_CODE As String = e.Row.Cells("CURR_CODE").Value & ""

        ' ASCMAIN1.Initalize_FCB("PAY", e)
        If e.Row.Cells("PAY").Value & "" = "1" Then
            e.Row.Cells("PAY").ButtonAppearance = btnSelected
        Else
            e.Row.Cells("PAY").ButtonAppearance = btnBlank
        End If

        If INV_PMT <> 0 Or INV_DISC_TAKEN <> 0 Or INV_WRITE_OFF <> 0 Then
            e.Row.Cells("INV_NUM").Appearance = bcLightGreen
            If INV_BALANCE_NEW_CURR <> 0 Then
                e.Row.Cells("INV_NUM").Appearance = fcRed
            Else
                'e.Row.Cells("INV_NUM").Appearance = fcEmpty
            End If
        Else
            e.Row.Cells("INV_NUM").Appearance = bcBeige
            '  e.Row.Cells("INV_NUM").Appearance = fcEmpty
        End If

        If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
            e.Row.Cells("INV_BALANCE_CURR").Appearance = fcRed
            e.Row.Cells("INV_BALANCE_CURR").ToolTipText = "Invoice Balance is in a Foreign Currency (" & CURR_CODE & ")"
        End If

        Select Case e.Row.Cells("AGE_BUCKET").Text
            Case "1"
                e.Row.Cells("AGE_BUCKET").Appearance = bcLightGreen
            Case "2"
                e.Row.Cells("AGE_BUCKET").Appearance = bcLightBlue
            Case "3"
                e.Row.Cells("AGE_BUCKET").Appearance = bcYellow
            Case "4"
                e.Row.Cells("AGE_BUCKET").Appearance = bcPink
        End Select
    End Sub

    Private Sub grdARTOPENA_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTOPENA.ClickCellButton
        Apply_to_Statement(e.Cell.Row.Cells("AGE_NO").Value)
        Calculate_Application_by_Type()
    End Sub

    Sub Calculate_Aging_Dates()
        ' Calculate Dates used in Aging
        For i As Integer = 1 To 4
            ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 " _
            & " where OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i) & "'"
            Dim PRD_END_DATE As Date = ASCDATA1.GetDataValue
            AGING_DATES_ado(i) = "#" & Format(PRD_END_DATE, "MM/dd/yyyy") & "#"
            AGING_DATES(i) = PRD_END_DATE
        Next i
    End Sub

    Sub Calculate_Aging_for_Unapplied()

    End Sub

    Sub Calculate_Aging()

        Dim AGE_WHERE As String
        Dim T As Double = 0
        Dim AGE_AMT As Double = 0

        ReDim AGED_TOTALS(4)

        Dim rowARTOPENA As DataRow

        For i As Integer = 1 To 4
            If i = 1 Then
                AGE_WHERE = AGE_DATE_COLUMN & " > " & AGING_DATES_ado(i)
            ElseIf i = 4 Then
                AGE_WHERE = AGE_DATE_COLUMN & " <= " & AGING_DATES_ado(i - 1)
            Else
                AGE_WHERE = AGE_DATE_COLUMN & " > " & AGING_DATES_ado(i) & " and " & AGE_DATE_COLUMN & " <= " & AGING_DATES_ado(i - 1)
            End If

            For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select(AGE_WHERE)
                rowARTPYMT3.Item("AGE_BUCKET") = i
            Next

            AGE_AMT = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE)", AGE_WHERE) & "")
            T = T + AGE_AMT

            AGED_TOTALS(i) = AGE_AMT

            rowARTOPENA = dst.Tables("ARTOPENA").Rows.Find(i)
            rowARTOPENA.Item("AGE_AMT") = AGE_AMT
            rowARTOPENA.Item("AGE_AMT_NEW") = AGE_AMT
        Next
        AGED_TOTALS(0) = T

        'rowARTOPENA = dst.Tables("ARTOPENA").Rows.Find(9)
        'rowARTOPENA.Item("AGE_AMT") = T
    End Sub

    Sub Reverse_Payment()

        PYMT_BATCH_NO = Absx1.txtFor("PYMT_BATCH_NO").Text
        PYMT_BATCH_LNO = Val(Absx1.numFor("PYMT_BATCH_LNO").Value & "")

        PYMT_BATCH_NO_new = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        PYMT_BATCH_LNO_new = 1

        Dim sql As String = ""

        dst.Tables("ARTPYMT1").Rows.Clear()
        For i As Integer = 2 To 5
            sql = "Select * from ARTPYMT" & CStr(i) _
            & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
            If i > 1 Then
                sql &= " and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)
            End If
            Fill_Records("ARTPYMT" & CStr(i), "", True, sql)
        Next

        If ARTOPENX = "" Then
            ARTOPENX = ASCMAIN1.Temp_Table("Select * from ARTOPENX where ROWNUM < 1")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTOPENX)
        End If

        sql = "Select ARTOPEN1.* from ARTOPEN1 ARTOPEN1,ARTPYMT3,ARTPYMT2" & vbCrLf _
            & " where ARTOPEN1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
            & "   and ARTOPEN1.INV_TYPE = ARTPYMT3.INV_TYPE" & vbCrLf _
            & "   and ARTOPEN1.INV_NUM = ARTPYMT3.INV_NUM" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO) & vbCrLf _
            & "   and ARTPYMT3.PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'" & vbCrLf _
            & "   and ARTPYMT3.PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)

        ASCDATA1.ExecuteSQL("Insert into " & ARTOPENX & " " & Replace(sql, "from ARTOPEN1", "from ARTOPENX"))
        ' Fill_Records("ARTOPEN1", "", True, "Select * from " & ARTOPENX)
        ' Fill_Records("ARTOPEN1", "", False, sql)
        Fill_Records("ARTOPEN1", "", True, sql)


        sql = "Select ARTOPEN1.* from ARTOPEN1 ARTOPEN1,ARTPYMT5,ARTPYMT2" & vbCrLf _
            & " where ARTOPEN1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
            & "   and ARTOPEN1.INV_TYPE = ARTPYMT5.INV_TYPE_CB" & vbCrLf _
            & "   and ARTOPEN1.INV_NUM = ARTPYMT5.CHARGEBACK_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT5.PYMT_BATCH_LNO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)

        ASCDATA1.ExecuteSQL("Insert into " & ARTOPENX & " " & Replace(sql, "from ARTOPEN1", "from ARTOPENX"))
        ' Fill_Records("ARTOPEN1", "", False, "Select * from " & ARTOPENX)
        Fill_Records("ARTOPEN1", "", False, sql)

        Fill_Records("ARTOPEN1", "", False, "Select * from " & ARTOPENX)


        Dim PYMT_BATCH_DATE As Date = DATETIME_STAMP.Date
        Dim DTES() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
        Dim DTES_LAST = DTES(DTES.Length - 1)
        If Format(PYMT_BATCH_DATE, "yyyyMMdd") > Format(DTES_LAST, "yyyyMMdd") Then
            PYMT_BATCH_DATE = DTES_LAST
        End If

        Dim rowARTPYMT1_orig As DataRow = LookUp("ARTPYMT1", PYMT_BATCH_NO)
        dst.Tables("ARTPYMT1").Rows.Clear()
        Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
        With rowARTPYMT1
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
            .Item("PYMT_BATCH_DATE") = PYMT_BATCH_DATE
            .Item("BANK_CODE") = rowARTPYMT1_orig.Item("BANK_CODE")
            .Item("CURR_CODE") = rowARTPYMT1_orig.Item("CURR_CODE")
            .Item("CURR_EXCH_RATE") = rowARTPYMT1_orig.Item("CURR_EXCH_RATE")
            .Item("STATUS") = "1"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("PYMT_SOURCE") = "REV"
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
        End With
        dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

        Dim rowARTPYMT2_orig As DataRow = dst.Tables("ARTPYMT2").Rows(0)
        Dim CUST_CODE As String = rowARTPYMT2_orig.Item("CUST_CODE") & ""
        rowARTPYMT2_orig.Item("PYMT_BATCH_NO_REV") = PYMT_BATCH_NO_new
        rowARTPYMT2_orig.Item("PYMT_BATCH_LNO_REV") = PYMT_BATCH_LNO_new
        rowARTPYMT2_orig.Item("PYMT_REVERSED") = "1"
        Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
        With rowARTPYMT2
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
            .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_new
            .Item("CUST_CODE") = rowARTPYMT2_orig.Item("CUST_CODE")
            .Item("CUST_NAME") = rowARTPYMT2_orig.Item("CUST_NAME")
            .Item("CUST_PYMT_REF_NO") = rowARTPYMT2_orig.Item("CUST_PYMT_REF_NO")
            .Item("CUST_PYMT_REF_DATE") = rowARTPYMT2_orig.Item("CUST_PYMT_REF_DATE")
            .Item("CUST_PYMT_AMT") = -1 * Val(rowARTPYMT2_orig.Item("CUST_PYMT_AMT") & "")
            .Item("PYMT_STATUS") = "2"
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("CUST_PYMT_AMT_CURR") = -1 * Val(rowARTPYMT2_orig.Item("CUST_PYMT_AMT_CURR") & "")
            .Item("PYMT_BATCH_NO_ORIG") = PYMT_BATCH_NO
            .Item("PYMT_BATCH_LNO_ORIG") = PYMT_BATCH_LNO
            .Item("PYMT_REVERSED") = "2"
            .Item("PYMT_NOTE") = Absx1.txtFor("RETURNED_ITEM_REASON").Text
            .Item("CURR_CODE") = rowARTPYMT2_orig.Item("CURR_CODE")
            .Item("CURR_EXCH_RATE") = rowARTPYMT2_orig.Item("CURR_EXCH_RATE")
            dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)
        End With

        For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Rows
            With rowARTPYMT3
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_new
                .Item("INV_PMT") = -1 * Val(.Item("INV_PMT") & "")
                .Item("INV_DISC_TAKEN") = -1 * Val(.Item("INV_DISC_TAKEN") & "")
                .Item("INV_WRITE_OFF") = -1 * Val(.Item("INV_WRITE_OFF") & "")
                .Item("INV_PMT_CURR") = -1 * Val(.Item("INV_PMT_CURR") & "")
                .Item("INV_DISC_TAKEN_CURR") = -1 * Val(.Item("INV_DISC_TAKEN_CURR") & "")
                .Item("INV_WRITE_OFF_CURR") = -1 * Val(.Item("INV_WRITE_OFF_CURR") & "")

                Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find _
                (New Object() {CUST_CODE, .Item("INV_TYPE"), .Item("INV_NUM")})

                .Item("INV_BALANCE") = rowARTOPEN1.Item("INV_BALANCE")
                .Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE_CURR")

                rowARTOPEN1.Item("INV_BALANCE") = Val(rowARTOPEN1.Item("INV_BALANCE") & "") - Val(.Item("INV_PMT") & "") - Val(.Item("INV_DISC_TAKEN") & "") - Val(.Item("INV_WRITE_OFF") & "")
                rowARTOPEN1.Item("INV_BALANCE_CURR") = Val(rowARTOPEN1.Item("INV_BALANCE_CURR") & "") - Val(.Item("INV_PMT_CURR") & "") - Val(.Item("INV_DISC_TAKEN_CURR") & "") - Val(.Item("INV_WRITE_OFF_CURR") & "")

                rowARTOPEN1.Item("INV_PMT") = Val(rowARTOPEN1.Item("INV_PMT") & "") + Val(.Item("INV_PMT") & "")
                rowARTOPEN1.Item("INV_PMT_CURR") = Val(rowARTOPEN1.Item("INV_PMT_CURR") & "") + Val(.Item("INV_PMT_CURR") & "")
                rowARTOPEN1.Item("INV_DISC_TAKEN") = Val(rowARTOPEN1.Item("INV_DISC_TAKEN") & "") + Val(.Item("INV_DISC_TAKEN") & "")
                rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") = Val(rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") & "") + Val(.Item("INV_DISC_TAKEN_CURR") & "")
                rowARTOPEN1.Item("INV_WRITE_OFF") = Val(rowARTOPEN1.Item("INV_WRITE_OFF") & "") + Val(.Item("INV_WRITE_OFF") & "")
                rowARTOPEN1.Item("INV_WRITE_OFF_CURR") = Val(rowARTOPEN1.Item("INV_WRITE_OFF_CURR") & "") + Val(.Item("INV_WRITE_OFF_CURR") & "")

                rowARTOPEN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowARTOPEN1.Item("LAST_DATE") = DATETIME_STAMP

                .Item("INV_BALANCE_NEW") = rowARTOPEN1.Item("INV_BALANCE")
                .Item("INV_BALANCE_NEW_CURR") = rowARTOPEN1.Item("INV_BALANCE_CURR")

                rowARTOPEN1.Item("AMT_PAID") = DBNull.Value
                rowARTOPEN1.Item("OPS_YYYYPP_PAID") = DBNull.Value
                rowARTOPEN1.Item("OPS_YYYYPP_F") = DBNull.Value
                rowARTOPEN1.Item("DTP") = DBNull.Value
                rowARTOPEN1.Item("DATE_PAID") = DBNull.Value

                .AcceptChanges()
                .SetAdded()
            End With
        Next

        For Each rowARTPYMT4 As DataRow In dst.Tables("ARTPYMT4").Rows
            With rowARTPYMT4
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_new
                .Item("GL_DIST_AMT_CURR") = -1 * Val(.Item("GL_DIST_AMT_CURR") & "")
                .Item("GL_DIST_AMT") = -1 * Val(.Item("GL_DIST_AMT") & "")
                .AcceptChanges()
                .SetAdded()
            End With
        Next



        Dim PYMT_BATCH_ILNO_ctr As Integer = 0

        Dim CBs As New Dictionary(Of String, String)

        For Each rowARTPYMT5 As DataRow In dst.Tables("ARTPYMT5").Rows
            With rowARTPYMT5
                Dim GL_DIST_AMT_CURR As Decimal = Val(.Item("GL_DIST_AMT_CURR") & "")
                Dim GL_DIST_AMT As Decimal = Val(.Item("GL_DIST_AMT") & "")

                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_new
                .Item("GL_DIST_AMT_CURR") = -1 * GL_DIST_AMT_CURR
                .Item("GL_DIST_AMT") = -1 * GL_DIST_AMT
                If .Item("CHARGEBACK_IND") & "" = "1" Then
                    Dim CHARGEBACK_NO_old As String = rowARTPYMT5.Item("CHARGEBACK_NO")
                    Me.Load_Open_AR_from_CB(rowARTPYMT5, rowARTPYMT2_orig, DATETIME_STAMP.Date)
                    Dim CHARGEBACK_NO_new As String = rowARTPYMT5.Item("CHARGEBACK_NO")
                    CBs.Add(CHARGEBACK_NO_old, CHARGEBACK_NO_new)

                    For CB As Integer = 0 To 1
                        Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").NewRow
                        With rowARTPYMT3
                            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                            .Item("PYMT_BATCH_LNO") = 2 ' PYMT_BATCH_LNO_new

                            PYMT_BATCH_ILNO_ctr += 1
                            .Item("PYMT_BATCH_ILNO") = PYMT_BATCH_ILNO_ctr

                            Dim INV_PMT As Decimal = 0
                            Dim INV_PMT_CURR As Decimal = 0

                            .Item("INV_TYPE") = rowARTPYMT5.Item("INV_TYPE_CB")
                            If CB = 0 Then
                                .Item("INV_NUM") = CHARGEBACK_NO_old
                                INV_PMT = GL_DIST_AMT
                                INV_PMT_CURR = GL_DIST_AMT_CURR
                            Else
                                .Item("INV_NUM") = CHARGEBACK_NO_new
                                INV_PMT = -1 * GL_DIST_AMT
                                INV_PMT_CURR = -1 * GL_DIST_AMT_CURR
                            End If

                            .Item("INV_PMT") = INV_PMT
                            .Item("INV_DISC_TAKEN") = 0
                            .Item("INV_WRITE_OFF") = 0
                            .Item("INV_PMT_CURR") = INV_PMT_CURR
                            .Item("INV_DISC_TAKEN_CURR") = 0
                            .Item("INV_WRITE_OFF_CURR") = 0

                            Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").Rows.Find _
                                (New Object() {CUST_CODE, .Item("INV_TYPE"), .Item("INV_NUM")})

                            For Each C As String In New String() {"REASON_CODE", "INV_DATE", "INV_DUE_DATE", "CUST_CODE_SO", "CUST_STORE_NO", "INV_CUST_PO", "POST_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ORDR_TYPE_CODE", "CURR_CODE", "CURR_EXCH_RATE", "INV_NO_CONS", "PARTNER_ORDR_NO"}
                                .Item(C) = rowARTOPEN1.Item(C)
                            Next

                            .Item("INV_BALANCE") = rowARTOPEN1.Item("INV_BALANCE")
                            .Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE_CURR")

                            rowARTOPEN1.Item("INV_BALANCE") = Val(rowARTOPEN1.Item("INV_BALANCE") & "") - Val(.Item("INV_PMT") & "") - Val(.Item("INV_DISC_TAKEN") & "") - Val(.Item("INV_WRITE_OFF") & "")
                            rowARTOPEN1.Item("INV_BALANCE_CURR") = Val(rowARTOPEN1.Item("INV_BALANCE_CURR") & "") - Val(.Item("INV_PMT_CURR") & "") - Val(.Item("INV_DISC_TAKEN_CURR") & "") - Val(.Item("INV_WRITE_OFF_CURR") & "")

                            rowARTOPEN1.Item("INV_PMT") = Val(rowARTOPEN1.Item("INV_PMT") & "") + Val(.Item("INV_PMT") & "")
                            rowARTOPEN1.Item("INV_PMT_CURR") = Val(rowARTOPEN1.Item("INV_PMT_CURR") & "") + Val(.Item("INV_PMT_CURR") & "")
                            rowARTOPEN1.Item("INV_DISC_TAKEN") = Val(rowARTOPEN1.Item("INV_DISC_TAKEN") & "") + Val(.Item("INV_DISC_TAKEN") & "")
                            rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") = Val(rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") & "") + Val(.Item("INV_DISC_TAKEN_CURR") & "")
                            rowARTOPEN1.Item("INV_WRITE_OFF") = Val(rowARTOPEN1.Item("INV_WRITE_OFF") & "") + Val(.Item("INV_WRITE_OFF") & "")
                            rowARTOPEN1.Item("INV_WRITE_OFF_CURR") = Val(rowARTOPEN1.Item("INV_WRITE_OFF_CURR") & "") + Val(.Item("INV_WRITE_OFF_CURR") & "")

                            rowARTOPEN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            rowARTOPEN1.Item("LAST_DATE") = DATETIME_STAMP

                            .Item("INV_BALANCE_NEW") = rowARTOPEN1.Item("INV_BALANCE")
                            .Item("INV_BALANCE_NEW_CURR") = rowARTOPEN1.Item("INV_BALANCE_CURR")

                            rowARTOPEN1.Item("AMT_PAID") = DBNull.Value
                            rowARTOPEN1.Item("OPS_YYYYPP_PAID") = DBNull.Value
                            rowARTOPEN1.Item("OPS_YYYYPP_F") = DBNull.Value
                            rowARTOPEN1.Item("DTP") = DBNull.Value
                            rowARTOPEN1.Item("DATE_PAID") = DBNull.Value

                            dst.Tables("ARTPYMT3").Rows.Add(rowARTPYMT3)
                        End With
                    Next
                End If

                .AcceptChanges()
                .SetAdded()
            End With
        Next

        If CBs.Count <> 0 Then
            Dim rowARTPYMT2_new2 As DataRow = dst.Tables("ARTPYMT2").NewRow
            With rowARTPYMT2_new2
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                .Item("PYMT_BATCH_LNO") = 2 ' PYMT_BATCH_LNO_new
                .Item("CUST_CODE") = rowARTPYMT2_orig.Item("CUST_CODE")
                .Item("CUST_NAME") = rowARTPYMT2_orig.Item("CUST_NAME")
                .Item("CUST_PYMT_REF_NO") = rowARTPYMT2_orig.Item("CUST_PYMT_REF_NO")
                .Item("CUST_PYMT_REF_DATE") = rowARTPYMT2_orig.Item("CUST_PYMT_REF_DATE")
                .Item("CUST_PYMT_AMT") = 0
                .Item("PYMT_STATUS") = "2"
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("CUST_PYMT_AMT_CURR") = 0
                .Item("PYMT_REVERSED") = "2"
                .Item("PYMT_NOTE") = "Net Chargebacks"
                .Item("CURR_CODE") = rowARTPYMT2_orig.Item("CURR_CODE")
                .Item("CURR_EXCH_RATE") = rowARTPYMT2_orig.Item("CURR_EXCH_RATE")
                dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2_new2)
            End With
        End If

        If reverse_application_option = "Return Payment" Then
            Dim RETURNED_ITEM_FEE As Decimal = Val(Absx1.numFor("RETURNED_ITEM_FEE").Value & "")
            If RETURNED_ITEM_FEE <> 0 Then
                Dim rowARTPYMT5 As DataRow = Nothing
                For Each CHARGEBACK_IND As String In New String() {"0", "1"}
                    rowARTPYMT5 = dst.Tables("ARTPYMT5").NewRow
                    With rowARTPYMT5
                        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                        .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_new
                        Dim sqlw As String = "PYMT_BATCH_NO = '" & PYMT_BATCH_NO_new & "' and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO_new)
                        .Item("PYMT_BATCH_DLNO") = Val(dst.Tables("ARTPYMT5").Compute("MAX(PYMT_BATCH_DLNO)", sqlw) & "") + 1
                        .Item("REASON_CODE") = Absx1.txtFor("RETURNED_ITEM_REASON_CODE").Text
                        rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(.Item("REASON_CODE"))
                        If rowARTREAS1 IsNot Nothing Then .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
                        RETURNED_ITEM_FEE = -1 * RETURNED_ITEM_FEE
                        .Item("GL_DIST_AMT") = RETURNED_ITEM_FEE
                        .Item("GL_DIST_AMT_CURR") = RETURNED_ITEM_FEE
                        .Item("GL_DIST_COMMENT") = Absx1.txtFor("RETURNED_ITEM_REASON").Text
                        .Item("CHARGEBACK_IND") = CHARGEBACK_IND
                        .Item("CUST_REFERENCE") = "RTN PYMT FEE"
                        .Item("INV_TYPE_CB") = IIf(CHARGEBACK_IND = "1", "B", "")
                        .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                        .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                        .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                        'ACCT_CODE
                        'CUST_CODE_SO
                    End With
                    dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)
                Next
                Me.Load_Open_AR_from_CB(rowARTPYMT5, rowARTPYMT2_orig, DATETIME_STAMP.Date)
            End If
        End If

        BeginTrans()

        For i As Integer = 1 To 5
            Update_Record_TDA("ARTPYMT" & CStr(i))
        Next

        ASCMAIN1.sql = "Insert into ARTOPEN1 Select * from " & ARTOPENX
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
        & "Begin " _
        & " Declare Cursor C1 is Select * from " & ARTOPENX & ";" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Delete from ARTOPENX " _
        & "    where CUST_CODE = R1.CUST_CODE " _
        & "      and INV_TYPE = R1.INV_TYPE " _
        & "      and INV_NUM = R1.INV_NUM;" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("ARTOPEN1")
        If CUST_CODE <> "" Then
            ASCDATA1.ExecuteSP("ARPCUST6_PYMT", "VN" _
                       , New Object() {PYMT_BATCH_NO_new, PYMT_BATCH_LNO_new} _
                       , New String() {"PYMT_BATCH_NO_IN", "PYMT_BATCH_LNO_IN"})
        End If

        rowARTPYMT1 = LookUp("ARTPYMT1", PYMT_BATCH_NO)
        rowARTPYMT2 = LookUp("ARTPYMT2", New String() {PYMT_BATCH_NO, PYMT_BATCH_LNO})

        If rowARTPYMT1.Item("PYMT_SOURCE") & "" = "820" Then
            Dim EDI_DOC_SEQ_NO As String = rowARTPYMT2.Item("EDI_DOC_SEQ_NO")
            ASCMAIN1.sql = "Update EDT820T1 Set EDI_PROCESS_IND = '0' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_PROCESS_IND = '1'"
            ASCDATA1.ExecuteSQL()
        End If

        CommitTrans()
    End Sub

    Private Sub grdARTPYMTX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMTX.InitializeRow
        Dim PYMT_REVERSED As String = e.Row.Cells("PYMT_REVERSED").Value & ""
        If PYMT_REVERSED = "1" Then
            e.Row.Cells("PYMT_BATCH_NO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PYMT_REVERSED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PYMT_BATCH_NO").ToolTipText = "Payment was Reversed"
        ElseIf PYMT_REVERSED = "2" Then
            e.Row.Cells("PYMT_BATCH_NO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PYMT_REVERSED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("CUST_PYMT_AMT_CURR").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PYMT_BATCH_NO").ToolTipText = "Payment Reversed another Payment"
        End If
        Dim PYMT_DELETED As String = e.Row.Cells("PYMT_DELETED").Value & ""
        If PYMT_DELETED = "1" Then
            e.Row.Cells("PYMT_BATCH_NO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PYMT_DELETED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PYMT_BATCH_NO").ToolTipText = "Payment was Deleted"
        End If

    End Sub

    Sub Match_CBs_to_Reversed_CBs()
        PYMT_BATCH_LNO_new = 2
    End Sub

    Sub Restore_Payment()
        PYMT_BATCH_LNO_new = 3

        Dim sql As String = ""

        For i As Integer = 2 To 5
            sql = "Select * from ARTPYMT" & CStr(i) _
            & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
            If i > 1 Then
                sql &= " and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)
            End If
            Fill_Records("ARTPYMT" & CStr(i), "", True, sql)

            If i > 2 Then
                For Each row As DataRow In dst.Tables("ARTPYMT" & CStr(i)).Rows
                    row.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
                    row.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_new
                    row.AcceptChanges()
                    row.SetAdded()
                Next
            End If
        Next

        Dim rowARTPYMT2_orig As DataRow = dst.Tables("ARTPYMT2").Rows(0)
        Dim CUST_CODE As String = rowARTPYMT2_orig.Item("CUST_CODE") & ""
        Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
        With rowARTPYMT2
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO_new
            .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO_new
            .Item("CUST_CODE") = rowARTPYMT2_orig.Item("CUST_CODE")
            .Item("CUST_NAME") = rowARTPYMT2_orig.Item("CUST_NAME")
            .Item("CUST_PYMT_REF_NO") = rowARTPYMT2_orig.Item("CUST_PYMT_REF_NO")
            .Item("CUST_PYMT_REF_DATE") = rowARTPYMT2_orig.Item("CUST_PYMT_REF_DATE")
            .Item("CUST_PYMT_AMT") = Val(rowARTPYMT2_orig.Item("CUST_PYMT_AMT") & "")
            .Item("PYMT_STATUS") = "1"
            .Item("CUST_PYMT_AMT_CURR") = Val(rowARTPYMT2_orig.Item("CUST_PYMT_AMT_CURR") & "")
            .Item("PYMT_BATCH_NO_ORIG") = PYMT_BATCH_NO
            .Item("PYMT_BATCH_LNO_ORIG") = PYMT_BATCH_LNO
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("CURR_CODE") = rowARTPYMT2_orig.Item("CURR_CODE")
            .Item("CURR_EXCH_RATE") = rowARTPYMT2_orig.Item("CURR_EXCH_RATE")
            dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)
        End With
        rowARTPYMT2_orig.Delete()
        rowARTPYMT2_orig.AcceptChanges()

        Update_Record_TDA("ARTPYMT2")

    End Sub

    Private Sub grdARTPYMT2_SPLIT_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMT2_SPLIT.InitializeRow
        Dim PYMT_DELETED As String = e.Row.Cells("PYMT_DELETED").Text
        If PYMT_DELETED = "1" Then
            e.Row.Cells("CUST_PYMT_AMT").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("CUST_PYMT_AMT_CURR").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub cmdMovePayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMovePayment.Click
        Absx1.txtFor("CUST_CODE_MOVE_TO").Text = ASCMAIN1.Format_Field(Absx1.txtFor("CUST_CODE_MOVE_TO").Text, "CUST_CODE")
        Click_Command("Move Payment")
    End Sub

    Private Sub grdARTOPENA_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTOPENA.InitializeRow
        Select Case e.Row.Cells("AGE_NO").Text
            Case "1"
                e.Row.Cells("AGE_NO").Appearance.BackColor = Drawing.Color.LightGreen
            Case "2"
                e.Row.Cells("AGE_NO").Appearance.BackColor = Drawing.Color.LightBlue
            Case "3"
                e.Row.Cells("AGE_NO").Appearance.BackColor = Drawing.Color.Yellow
            Case "4"
                e.Row.Cells("AGE_NO").Appearance.BackColor = Drawing.Color.Pink
        End Select
    End Sub

    Sub Apply_to_Statement_Oldest_First(ByVal Pay_to_Bucket As Int32)
        For i As Int32 = 4 To Pay_to_Bucket Step -1
            Apply_to_Statement(i)
        Next
        Calculate_Application_by_Type()
    End Sub

    Sub Apply_to_Statement(ByVal AGE_BUCKET As Int32)

        Dim CASH_AVAIL As Decimal = Val(dst.Tables("ARTPYMTT").Rows.Find("9").Item("PYMT_TOTAL_AMT") & "")

        Dim TOTAL_BUCKET As Decimal = Val(dst.Tables("ARTPYMT3").Compute("SUM (INV_BALANCE_CURR)", "AGE_BUCKET = " & CStr(AGE_BUCKET)) & "")
        Dim ok_to_go_negative As Boolean = False

        If TOTAL_BUCKET <= CASH_AVAIL Then
            ok_to_go_negative = True
        End If

        If CASH_AVAIL <= 0 And Not ok_to_go_negative And Not application_only And Not applying_to_statement Then
            MsgBox("No Cash Available to Apply", MsgBoxStyle.OkOnly, "Applying to Aging Column " & CStr(AGE_BUCKET))
            Exit Sub
        End If

        For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3") _
        .Select("AGE_BUCKET = " & CStr(AGE_BUCKET), "INV_DUE_DATE,INV_NUM")

            Dim PYMT_AMT_REQUIRED As Decimal = Val(rowARTPYMT3.Item("INV_BALANCE_CURR") & "")
            Dim PYMT_AMT As Decimal = 0
            Debug.Print(PYMT_AMT_REQUIRED)

            If Not application_only And Not applying_to_statement And CASH_AVAIL = 0 And PYMT_AMT_REQUIRED > 0 Then
                MsgBox("Ran out of Cash before Paying Entire Aging Bucket", MsgBoxStyle.OkOnly, "Please Note")
                Exit For
            End If

            If CASH_AVAIL >= PYMT_AMT_REQUIRED Or applying_to_statement Or ok_to_go_negative Then
                PYMT_AMT = PYMT_AMT_REQUIRED
                CASH_AVAIL = CASH_AVAIL - PYMT_AMT
            Else
                PYMT_AMT = CASH_AVAIL
                CASH_AVAIL = 0
            End If

            rowARTPYMT3.Item("INV_PMT_CURR") = PYMT_AMT
            rowARTPYMT3.Item("INV_DISC_TAKEN_CURR") = 0
            rowARTPYMT3.Item("INV_WRITE_OFF_CURR") = 0
            rowARTPYMT3.Item("INV_BALANCE_NEW_CURR") = PYMT_AMT_REQUIRED - PYMT_AMT
        Next
    End Sub

#Region "grdARTPYMTB"
    Private Sub grdARTPYMTB_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMTB.AfterRowsDeleted
        Dim K As List(Of String) = grdARTPYMTB.Tag

        For Each PYMT_BATCH_NO As String In K
            ASCDATA1.ExecuteSQL("Delete from ARTPYMT2 where PYMT_BATCH_NO = :PARM1 AND NVL(CUST_PYMT_AMT,0) = 0 AND PYMT_STATUS = '1'", "V", New Object() {PYMT_BATCH_NO})
            Dim R As Int32 = Val(ASCDATA1.GetDataValue("Select Count (*) from ARTPYMT2 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"))
            If R = 0 Then
                ASCDATA1.ExecuteSQL("Delete from ARTPYMT1 where PYMT_BATCH_NO = :PARM1 AND PYMT_APPL_ONLY = '1'", "V", New Object() {PYMT_BATCH_NO})
            End If
        Next

        MsgBox("Deletion Completed", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Private Sub grdARTPYMTB_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTPYMTB.BeforeRowsDeleted
        Dim k As New List(Of String)

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("PYMT_APPL_ONLY").Value & "" <> "1" Or Val(grow.Cells("S2").Value & "") <> 0 Or Val(grow.Cells("CUST_PYMT_AMT").Value & "") <> 0 Then
                MsgBox("Deletion of an Actual Payment is Not Permitted", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
                Exit Sub
            End If

            k.Add(grow.Cells("PYMT_BATCH_NO").Value)
            grdARTPYMTB.Tag = k
        Next
    End Sub
#End Region


    Public Overrides Function Remote_Control(
ByVal command As String,
Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "View Payment Application"
                Dim batchInfo As String() = Split(key, ":")

                Absx1.txtFor("PYMT_BATCH_NO").Text = batchInfo(0)
                Absx1.numFor("PYMT_BATCH_LNO").Value = Val(batchInfo(1) & "")
                Click_Command("View")

            Case "Done"
                Click_Command("Done")
        End Select

        Return return_key
    End Function

#Region "EDI (820)"

    Sub Delete_820()
        BeginTrans()
        For Each grow As UltraWinGrid.UltraGridRow In grdEDT820TX.Selected.Rows
            Dim EDI_DOC_SEQ_NO As String = grow.Cells("EDI_DOC_SEQ_NO").Value
            ASCMAIN1.sql = "Update EDT820T1 set EDI_PROCESS_IND = 'D' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_PROCESS_IND = '0'"
            ASCDATA1.ExecuteSQL()
        Next
        CommitTrans()

        ASCMAIN1.MultiTask_Release()
        MsgBox(CStr(grdEDT820TX.Selected.Rows.Count) & " Records have been Deleted", MsgBoxStyle.OkOnly, "Verification")

        Load_EDI_Grid()

    End Sub

    Sub Create_ARTCASH1_2()

        If dst.Tables("ARTPYMT1").Select("").Length <> 0 _
            Or dst.Tables("ARTPYMT2").Select("").Length <> 0 _
            Or dst.Tables("ARTPYMT3").Select("").Length <> 0 _
            Or dst.Tables("ARTPYMT4").Select("").Length <> 0 _
            Or dst.Tables("ARTPYMT5").Select("").Length <> 0 _
            Or dst.Tables("ARTOPEN1").Select("").Length <> 0 _
            Or dst.Tables("EDTERRS1").Select("").Length <> 0 Then Stop

        PYMT_BATCH_NO = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        PYMT_BATCH_LNO = 1

        BANK_CODE = ROWs("ARTPARM1").Item("AR_PARM_BANK_CODE")
        ' SHOULD BE DETERMINED FROM EDI
        ' OR ELSE AR_PARM_BANK_CODE_EDI

        Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
        With rowARTPYMT1
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO

            .Item("BANK_CODE") = BANK_CODE
            .Item("CURR_CODE") = rowEDT820T1.Item("EDI_CURR_CODE") & ""
            If .Item("CURR_CODE") & "" = "" Then
                .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            End If
            If .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                .Item("CURR_EXCH_RATE") = 1
            Else
                Stop
                ' WHAT NOW
            End If

            .Item("PYMT_BATCH_DATE") = CDate(rowEDT820T1.Item("EDI_RECD_DATE") & "").Date
            .Item("STATUS") = "1"
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("PYMT_SOURCE") = "820"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
        End With
        dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

        CURR_CODE = rowARTPYMT1.Item("CURR_CODE")
        '   CURR_EXCH_RATE = rowARTPYMT1.Item("CURR_EXCH_RATE")

        Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
        With rowARTPYMT2
            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
            .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO

            .Item("CUST_CODE") = rowEDT820T1.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowEDT820T1.Item("CUST_CODE"))
            .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
            .Item("CUST_PYMT_REF_NO") = rowEDT820T1.Item("TRACE_ID")
            .Item("CUST_PYMT_REF_DATE") = rowEDT820T1.Item("DOC_DATE")
            .Item("CUST_PYMT_AMT") = rowEDT820T1.Item("PYMT_AMT")
            .Item("CUST_PYMT_AMT_CURR") = rowEDT820T1.Item("PYMT_AMT")
            .Item("PYMT_STATUS") = "D"
            .Item("PYMT_NOTE") = "EDI 820 Payment Receipt Record"
            .Item("CUST_PYMT_ROUTING_NO") = rowEDT820T1.Item("ORIGINATING_DFI_ID")
            .Item("CUST_PYMT_BANK_ACCT_NO") = rowEDT820T1.Item("ORIGINATING_ACCT_NUMBER")
            .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO

            .Item("CURR_CODE") = "USD"
            .Item("CURR_EXCH_RATE") = 1

            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP

        End With
        dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)

        CUST_CODE = rowARTPYMT2.Item("CUST_CODE")

        BeginTrans()
        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")
        CommitTrans()

    End Sub

    Sub Delete_ARTCASH1_2()
        ASCMAIN1.sql = "Delete from ARTPYMT2 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
        'ASCMAIN1.sql = "Delete from ARTPYMT2 where PYMT_BATCH_NO IN (Select PYMT_BATCH_NO from ARTPYMT1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from ARTPYMT1 where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
        'ASCMAIN1.sql = "Delete from ARTPYMT1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Function EDI_Process() As Boolean

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Processing 820", "")

        Dim new_method As Boolean = True

        Dim CURR_EXCH_RATE As Decimal = 1
        Dim AR_PARM_AUTO_WOFF As Decimal = 0 ' Val(ROWs("ARTPARM1").Item("AR_PARM_AUTO_WOFF") & "")

        Dim ctr820T2 As Int64
        Dim ctr820T2_nogo As Int64
        Dim INV_PMT As Decimal

        Dim T3 As Decimal = 0
        Dim T4 As Decimal = 0
        Dim T5 As Decimal = 0

        Dim ctr820T4 As Int64 = 0
        Dim ctr820T4_matched As Int64 = 0

        Dim EDI_TP_QUAL As String = rowEDT820T1.Item("EDI_TP_QUAL")
        Dim EDI_TP_ID As String = rowEDT820T1.Item("EDI_TP_ID")

        Fill_Records("EDTXREF1", New String() {EDI_TP_QUAL, EDI_TP_ID})
        Fill_Records("EDTXREF2")

        Dim rowSOTTYPE1_OA As DataRow = LookUp("SOTTYPE1", ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_OA"))
        Dim rowSOTTYPE1_CB As DataRow = LookUp("SOTTYPE1", ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_CB"))

        For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select("INV_TYPE = 'I'")
            dst.Tables("EDTINVC1").Rows.Add(New Object() {rowARTPYMT3.Item("INV_NUM"), rowARTPYMT3.Item("PYMT_BATCH_NO"), rowARTPYMT3.Item("PYMT_BATCH_LNO"), rowARTPYMT3.Item("PYMT_BATCH_ILNO")})
        Next

        Dim macys_accounts As Boolean =
            (CUST_CODE = "MACYS" Or CUST_CODE = "BLOOMIES" _
             Or CUST_CODE = "BLOOMCOM" Or CUST_CODE = "MACYSCOM" Or CUST_CODE = "MACYSBACK")

        Dim PYMT_BATCH_DLNO_ctr As Integer = 0
        Dim PYMT_BATCH_ILNO As Integer = 0

        ASCMAIN1.Progress("Now Processing Invoices Paid")

        For Each rowEDT820T2 As DataRow In dst.Tables("EDT820T2").Select("")

            Dim EDI_ENT_NO As Integer = Val(rowEDT820T2.Item("EDI_ENT_NO") & "")
            Dim rowEDTXREF2 As DataRow = dst.Tables("EDTXREF2").Rows.Find _
                                                (New Object() {EDI_TP_QUAL, EDI_TP_ID, EDI_ENT_NO})
            '? - GET THIS FROM ORDER TYPE FILE
            rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find("000")

            For Each rowEDT820T3 As DataRow In dst.Tables("EDT820T3") _
                .Select("EDI_ENT_NO = " & CStr(EDI_ENT_NO))

                Dim AMT_NET_DUE As Decimal = Val(rowEDT820T3.Item("AMT_NET_DUE") & "")
                Dim AMT_GROSS As Decimal = Val(rowEDT820T3.Item("AMT_GROSS") & "")
                Dim AMT_DISCOUNT As Decimal = Val(rowEDT820T3.Item("AMT_DISCOUNT") & "")

                If CUST_CODE = "DILLARDS" Or CUST_CODE = "NORDSTROM" Then
                    INV_PMT = AMT_NET_DUE + AMT_DISCOUNT
                ElseIf CUST_CODE = "KOHLS" Then
                    If ASCMAIN1.CLIENT = "AHA" Or ASCMAIN1.CLIENT = "INT" Or ASCMAIN1.CLIENT = "VAN" Then
                        INV_PMT = AMT_GROSS  ' CHARGEBACK THE DISCOUNT

                        Dim REASON_CODE As String = ""
                        ' rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
                        Dim ADJ_AMT As Decimal = AMT_DISCOUNT
                        T5 = T5 - ADJ_AMT

                        AMT_DISCOUNT = 0
                        Dim CUST_REFERENCE As String = rowEDT820T3.Item("EDI_INVOICE_NO") & ""
                        Dim RMR_DESC As String = rowEDT820T3.Item("RMR_DESC") & ""
                        Dim CUST_CODE_SO As String = ""
                        If rowEDTXREF2 IsNot Nothing Then
                            CUST_CODE_SO = rowEDTXREF2.Item("CUST_CODE_SO")
                        End If

                        '  Stop

                        If ASCMAIN1.CLIENT = "INT" Then
                            Dim OUR_REFERENCE As String = CUST_REFERENCE
                            REASON_CODE = "J07" : rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
                            Dim ADJ_AMT_J07 As Decimal = System.Math.Round(AMT_GROSS * 0.01, 2)
                            Record_Chargeback(PYMT_BATCH_DLNO_ctr, REASON_CODE, -1 * ADJ_AMT_J07, CUST_REFERENCE, OUR_REFERENCE, RMR_DESC, CUST_CODE_SO)
                            Dim ADJ_AMT_J15 As Decimal = ADJ_AMT - ADJ_AMT_J07
                            REASON_CODE = "J15" : rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
                            Record_Chargeback(PYMT_BATCH_DLNO_ctr, REASON_CODE, -1 * ADJ_AMT_J15, CUST_REFERENCE, OUR_REFERENCE, RMR_DESC, CUST_CODE_SO)
                        Else

                            REASON_CODE = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_WOFF")
                            rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
                            Dim rowARTPYMT5 As DataRow = Record_Chargeback(PYMT_BATCH_DLNO_ctr, REASON_CODE, -1 * ADJ_AMT, CUST_REFERENCE, "", "Invoice Deduction", CUST_CODE_SO)

                            If ASCMAIN1.CLIENT = "VAN" Then
                                rowARTPYMT5.Item("CHARGEBACK_IND") = "0"
                                rowARTPYMT5.Item("OUR_REFERENCE") = CStr(100 * Val(rowARTPYMT5.Item("GL_DIST_AMT") & "") / INV_PMT) & "%"
                            End If
                            'Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
                            'With rowARTPYMT5
                            '    .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                            '    .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                            '    PYMT_BATCH_DLNO_ctr += 1
                            '    .Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO_ctr

                            '    .Item("REASON_CODE") = REASON_CODE
                            '    If REASON_CODE <> "" And rowARTREAS1 IsNot Nothing Then
                            '        .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
                            '    End If

                            '    .Item("GL_DIST_AMT") = ADJ_AMT * CURR_EXCH_RATE
                            '    .Item("GL_DIST_COMMENT") = "Invoice Deduction"
                            '    .Item("CHARGEBACK_IND") = "1"
                            '    .Item("CHARGEBACK_NO") = DBNull.Value
                            '    .Item("CUST_REFERENCE") = CUST_REFERENCE

                            '    .Item("CUST_CODE_SO") = CUST_CODE_SO
                            '    .Item("ACCT_CODE") = DBNull.Value
                            '    .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                            '    .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                            '    .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                            '    .Item("INV_TYPE_CB") = DBNull.Value
                            '    .Item("OUR_REFERENCE") = CUST_REFERENCE
                            '    .Item("GL_DIST_AMT_CURR") = ADJ_AMT

                            '    T5 = T5 - ADJ_AMT
                            'End With

                            'dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)

                        End If
                    Else
                        INV_PMT = AMT_GROSS - AMT_DISCOUNT
                    End If

                ElseIf CUST_CODE = "ECKERD" Or CUST_CODE = "WALMART" Or CUST_CODE = "JCPENNEY" Then
                    INV_PMT = AMT_GROSS
                ElseIf macys_accounts Then ' CUST_CODE = "MACYS" Or CUST_CODE = "BLOOMIES" Then
                    If new_method Then
                        INV_PMT = AMT_GROSS ' AMT_NET_DUE - USING GROSS SINCE WE WILL NOW BE CHARGING BACK PORTION NOT PAID
                    Else
                        INV_PMT = AMT_NET_DUE
                    End If
                Else
                    INV_PMT = AMT_GROSS
                End If

                ctr820T2 = ctr820T2 + 1
                Dim EDI_INVOICE_NO As String = rowEDT820T3.Item("EDI_INVOICE_NO") & ""
                ASCMAIN1.Progress("-", EDI_INVOICE_NO)

                Dim INV_NUM As String = EDI_INVOICE_NO.PadLeft(10, "0")
                If INV_NUM.Length > 10 Then
                    INV_NUM = Mid(INV_NUM, INV_NUM.Length - 10 + 1, 10)
                End If

                ' If ASCMAIN1.Running_in_VS And INV_NUM.EndsWith("32819") Then Stop

                Dim rowARTPYMT3s() As DataRow

                Dim SQLC As String = "INV_TYPE = 'I' and INV_NO_CONS = '" & INV_NUM & "'"
                rowARTPYMT3s = dst.Tables("ARTPYMT3").Select(SQLC)

                If rowARTPYMT3s.Length = 0 Then
                    Dim row As DataRow = dst.Tables("EDTINVC1").Rows.Find(INV_NUM)

                    If row IsNot Nothing Then
                        Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").Rows.Find _
                                                  ({row.Item("PYMT_BATCH_NO"),
                                                    row.Item("PYMT_BATCH_LNO"),
                                                    row.Item("PYMT_BATCH_ILNO")})
                        rowARTPYMT3s = New DataRow() {rowARTPYMT3}
                    Else
                        SQLC = "INV_TYPE = 'I' and INV_NUM = '" & INV_NUM & "'"
                        rowARTPYMT3s = dst.Tables("ARTPYMT3").Select(SQLC)
                        If rowARTPYMT3s.Length = 0 Then
                            SQLC = "INV_TYPE = 'I' and INV_NO_CONS = '" & INV_NUM & "'"
                            rowARTPYMT3s = dst.Tables("ARTPYMT3").Select(SQLC)
                        End If
                        If rowARTPYMT3s.Length = 0 Then
                            SQLC = "INV_TYPE = 'I' and PARTNER_ORDR_NO = '" & INV_NUM & "'"
                            rowARTPYMT3s = dst.Tables("ARTPYMT3").Select(SQLC)
                        End If
                        If rowARTPYMT3s.Length = 0 Then
                            SQLC = "INV_TYPE = 'B' and INV_CUST_PO LIKE '*" & INV_NUM & "'"
                            rowARTPYMT3s = dst.Tables("ARTPYMT3").Select(SQLC)
                        End If
                        If rowARTPYMT3s.Length = 0 Then
                            If Val(INV_NUM) <> 0 Then
                                SQLC = "INV_TYPE = 'B' and INV_CUST_PO LIKE '*" & CStr(Val(INV_NUM)) & "'"
                                rowARTPYMT3s = dst.Tables("ARTPYMT3").Select(SQLC)
                            End If
                        End If
                    End If
                End If

                '  If rowARTPYMT3s.Length = 0 Or rowARTPYMT3s.Length <> 1 Then Stop

                If rowARTPYMT3s.Length = 0 Then
                    ctr820T2_nogo = ctr820T2_nogo + 1
                    'Stop ' need to review amt_gross vs inv_pmt, and cust_code_so

                    ' why are we not using Record_Chargeback?

                    PYMT_BATCH_DLNO_ctr = PYMT_BATCH_DLNO_ctr + 1
                    Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
                    With rowARTPYMT5
                        .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                        .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                        .Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO_ctr
                        .Item("REASON_CODE") = rowSOTTYPE1_OA.Item("REASON_CODE") ' ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_OA")
                        rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(.Item("REASON_CODE"))
                        If rowARTREAS1 IsNot Nothing Then .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
                        .Item("ACCT_CODE") = DBNull.Value
                        .Item("GL_DIST_AMT") = INV_PMT * CURR_EXCH_RATE * -1
                        .Item("GL_DIST_COMMENT") = "MISSING INVOICE"
                        .Item("CHARGEBACK_IND") = "1"
                        .Item("CHARGEBACK_NO") = DBNull.Value
                        .Item("CUST_REFERENCE") = EDI_INVOICE_NO

                        If rowEDTXREF2 IsNot Nothing Then
                            .Item("CUST_CODE_SO") = rowEDTXREF2.Item("CUST_CODE_SO")
                        End If
                        .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                        .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                        .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                        .Item("INV_TYPE_CB") = DBNull.Value
                        .Item("OUR_REFERENCE") = EDI_INVOICE_NO
                        .Item("GL_DIST_AMT_CURR") = INV_PMT * -1
                    End With
                    dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)

                    T3 = T3 + INV_PMT

                Else

                    For Each rowARTPYMT3 As DataRow In rowARTPYMT3s
                        PYMT_BATCH_ILNO = PYMT_BATCH_ILNO + 1

                        With rowARTPYMT3
                            Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")

                            If ASCMAIN1.Running_in_VS AndAlso .Item("INV_NUM") & "" = "0020196170" Then Stop

                            Dim INV_PMT_invoice As Decimal = INV_PMT
                            If rowARTPYMT3s.Length > 1 Then
                                INV_PMT_invoice = INV_BALANCE
                                ' we should probably check to see if the sum of the invoices matches with INV_PMT before falling into this section
                                ' if they do not agree, the then application will end up out of balance
                            End If

                            Dim INV_DISC_TAKEN As Decimal = 0
                            Dim INV_WRITE_OFF As Decimal = 0
                            If CUST_CODE = "KOHLS" Then
                                INV_WRITE_OFF = AMT_DISCOUNT
                            End If

                            .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_PMT_invoice * CURR_EXCH_RATE
                            .Item("INV_DISC_TAKEN") = INV_DISC_TAKEN
                            .Item("INV_WRITE_OFF") = INV_WRITE_OFF
                            .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - (INV_PMT_invoice + INV_WRITE_OFF) * CURR_EXCH_RATE
                            .Item("INV_PMT_CURR") = Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice
                            .Item("INV_DISC_TAKEN_CURR") = INV_DISC_TAKEN
                            .Item("INV_WRITE_OFF_CURR") = INV_WRITE_OFF
                            .Item("INV_BALANCE_NEW_CURR") = Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_WRITE_OFF)


                        End With
                    Next

                    T3 = T3 + INV_PMT
                End If

                If ASCMAIN1.Running_in_VS Then
                    If CUST_CODE = "WALMART" Then Stop
                End If


                '                If (new_method And (CUST_CODE = "MACYS" Or CUST_CODE = "BLOOMIES")) _
                If (new_method And macys_accounts) _
                Or CUST_CODE = "DILLARDS" Or CUST_CODE = "WALMART" Then
                    Dim EDI_INV_SEQ As Integer = Val(rowEDT820T3.Item("EDI_INV_SEQ") & "")

                    For Each rowEDT820T5 As DataRow In dst.Tables("EDT820T5") _
                        .Select("EDI_ENT_NO = " & CStr(EDI_ENT_NO) & " and EDI_INV_SEQ = " & CStr(EDI_INV_SEQ))

                        PYMT_BATCH_DLNO_ctr = PYMT_BATCH_DLNO_ctr + 1

                        Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
                        With rowARTPYMT5
                            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                            .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                            .Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO_ctr

                            Dim ADJ_REAS_CODE As String = rowEDT820T5.Item("ADJ_REAS_CODE")
                            Dim REASON_CODE As String = Get_REASON_CODE(EDI_TP_QUAL, EDI_TP_ID, ADJ_REAS_CODE)

                            Dim ADJ_AMT As Decimal = Val(rowEDT820T5.Item("ADJ_AMT") & "")

                            .Item("REASON_CODE") = REASON_CODE
                            If REASON_CODE <> "" And rowARTREAS1 IsNot Nothing Then .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
                            .Item("ACCT_CODE") = DBNull.Value
                            .Item("GL_DIST_AMT") = ADJ_AMT * CURR_EXCH_RATE * -1
                            .Item("GL_DIST_COMMENT") = "Invoice Not Paid for Reasons Specified"
                            .Item("CHARGEBACK_IND") = "1"
                            .Item("CHARGEBACK_NO") = DBNull.Value

                            Dim CUST_REFERENCE As String = rowEDT820T5.Item("ADJ_CUST_PO") & ""
                            If CUST_REFERENCE = "" Then CUST_REFERENCE = rowEDT820T5.Item("ADJ_REF_NO") & ""
                            'If CUST_CODE = "NORDSTROM" Then ' FOR NORDSTROM, USE ADJ_REF_NO NOT ADJ_CUST_PO - MAYBE FOR OTHERS, PC WILL ADVISE
                            '    If rowEDT820T5.Item("ADJ_REF_NO") & "" <> "" Then
                            '        CUST_REFERENCE = rowEDT820T5.Item("ADJ_REF_NO") & ""
                            '    End If
                            'End If

                            If macys_accounts And rowEDT820T5.Item("ADJ_REF_NO") & "" <> "" Then ' WITH REFERENCE TO PC EMAIL 07/19
                                CUST_REFERENCE = rowEDT820T5.Item("ADJ_REF_NO") & ""
                            End If

                            .Item("CUST_REFERENCE") = CUST_REFERENCE

                            If rowEDTXREF2 IsNot Nothing Then
                                .Item("CUST_CODE_SO") = rowEDTXREF2.Item("CUST_CODE_SO")
                            End If
                            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                            .Item("INV_TYPE_CB") = DBNull.Value
                            .Item("OUR_REFERENCE") = EDI_INVOICE_NO
                            .Item("GL_DIST_AMT_CURR") = ADJ_AMT * -1

                            T5 = T5 + ADJ_AMT
                        End With
                        dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)
                    Next
                End If
            Next

            'ASCMAIN1.Progress("Now Processing Deductions", "")

            Dim rowEDT820T4_prior As DataRow = Nothing
            Dim rowARTPYMT5_prior As DataRow = Nothing

            For Each rowEDT820T4 As DataRow In dst.Tables("EDT820T4") _
                .Select("EDI_ENT_NO = " & CStr(EDI_ENT_NO), "EDI_ADJ_SEQ")
                ctr820T4 = ctr820T4 + 1
                Dim ADJ_AMT As Decimal = Val(rowEDT820T4.Item("ADJ_AMT") & "")
                Dim ADJ_REF_NO As String = rowEDT820T4.Item("ADJ_REF_NO") & ""
                Dim macys_reference As String = ""

                If ADJ_AMT = 0 Then
                    If rowEDT820T4_prior IsNot Nothing AndAlso
                        (rowEDT820T4_prior.Item("ADJ_REAS_CODE") & "" = rowEDT820T4.Item("ADJ_REAS_CODE") & "" And
                         rowEDT820T4_prior.Item("ADJ_REF_QUAL") & "" = rowEDT820T4.Item("ADJ_REF_QUAL") & "" And
                         rowEDT820T4_prior.Item("ADJ_REF_NO") & "" = rowEDT820T4.Item("ADJ_REF_NO") & "" And
                         rowEDT820T4.Item("ADJ_DESC") & "" <> "") Then
                        Dim ADJ_DESC As String = rowEDT820T4.Item("ADJ_DESC") & ""
                        If rowARTPYMT5_prior IsNot Nothing AndAlso rowARTPYMT5_prior.Item("CUST_REFERENCE") & "" = rowEDT820T4.Item("ADJ_REF_NO") & "" Then
                            Dim GL_DIST_COMMENT As String = rowARTPYMT5_prior.Item("GL_DIST_COMMENT") & ""
                            If GL_DIST_COMMENT <> "" Then GL_DIST_COMMENT &= vbCrLf
                            GL_DIST_COMMENT &= ADJ_DESC
                            rowARTPYMT5_prior.Item("GL_DIST_COMMENT") = GL_DIST_COMMENT
                            'If CUST_CODE = "MACYS" And ADJ_DESC.StartsWith("RA   = ") Then
                            ' not sure if this is applicable any more with the macys_reference changes
                            If macys_accounts And ADJ_DESC.StartsWith("RA   = ") Then
                                rowARTPYMT5_prior.Item("OUR_REFERENCE") = Mid(ADJ_DESC, "RA   = ".Length + 1)
                            End If
                        End If
                    End If
                Else
                    ASCMAIN1.Progress("-", ADJ_REF_NO)
                    Dim ZZ As String = ""
                    Dim ZZBAL As String = ""
                    If rowEDT820T4.Item("ADJ_REF_QUAL") = "CM" Then
                        'zz = "'I', 'B', 'D'"
                        'ZZ = "'I'"
                        ZZ = "'I','B'"
                        ZZBAL = "INV_BALANCE > 0 AND "
                    Else
                        ZZ = "'C', 'O', 'R'"
                        ZZBAL = "INV_BALANCE < 0 AND "
                    End If

                    Dim T4_ADJ_DESC As String = rowEDT820T4.Item("ADJ_DESC") & ""
                    Dim T4_ADJ_REF As String = ""

                    If CUST_CODE = "DILLARDS" Or CUST_CODE = "NORDSTROM" Or CUST_CODE = "ECKERD" Or CUST_CODE = "WALMART" Or CUST_CODE = "JCPENNEY" Then
                        T4_ADJ_REF = rowEDT820T4.Item("ADJ_CUST_PO") & ""
                        If T4_ADJ_REF = "" Then
                            T4_ADJ_REF = rowEDT820T4.Item("ADJ_REF_NO") & ""
                        End If

                        If CUST_CODE = "NORDSTROM" Then ' FOR NORDSTROM, USE ADJ_REF_NO NOT ADJ_CUST_PO - MAYBE FOR OTHERS, PC WILL ADVISE
                            If rowEDT820T4.Item("ADJ_REF_NO") & "" <> "" Then
                                T4_ADJ_REF = rowEDT820T4.Item("ADJ_REF_NO") & ""
                            End If
                        End If

                    ElseIf macys_accounts Then ' CUST_CODE = "MACYS" Or CUST_CODE = "BLOOMIES" Then
                        T4_ADJ_REF = rowEDT820T4.Item("ADJ_REF_NO") & ""
                        T4_ADJ_DESC = Replace(T4_ADJ_DESC, "'", "")
                    Else
                        ' same as dillards for now
                        T4_ADJ_REF = rowEDT820T4.Item("ADJ_CUST_PO") & ""
                        If T4_ADJ_REF = "" Then
                            T4_ADJ_REF = rowEDT820T4.Item("ADJ_REF_NO") & ""
                        End If
                    End If

                    'If T4_ADJ_REF.EndsWith("388503") Then Stop
                    'If T4_ADJ_DESC.EndsWith("388503") Then Stop

                    Dim ADJ_REAS_CODE As String = rowEDT820T4.Item("ADJ_REAS_CODE")
                    Dim REASON_CODE As String = Get_REASON_CODE(EDI_TP_QUAL, EDI_TP_ID, ADJ_REAS_CODE)

                    Dim sqlc As String = ZZBAL & " INV_TYPE IN (" & ZZ & ")" _
                        & " and (INV_CUST_PO = '" & T4_ADJ_REF & "'" _
                        & " or INV_CUST_PO = '" & CStr(Val(rowEDT820T4.Item("ADJ_REF_NO"))) & "'" _
                        & " or INV_CUST_PO = '" & rowEDT820T4.Item("ADJ_REF_NO") & "')"

                    ADJ_AMT = Val(rowEDT820T4.Item("ADJ_AMT") & "")

                    'If ASCMAIN1.Running_in_VS And T4_ADJ_REF.EndsWith("23476109") Then Stop
                    'If ASCMAIN1.Running_in_VS And T4_ADJ_REF.EndsWith("6121") Then Stop
                    If ASCMAIN1.Running_in_VS And T4_ADJ_REF.EndsWith("00003791482") And ADJ_AMT = -162.7 Then Stop

                    If macys_accounts And ASCMAIN1.CLIENT = "INT" And (T4_ADJ_DESC.StartsWith("CREDIT   RTV FREIGHT") Or T4_ADJ_DESC.StartsWith("REVERSE RTV HANDLING")) Then
                        sqlc &= " AND 1<>1" ' 07/27/16 this code appears twice in this section - Petra wants (for MACYS only) for ABSolution to NOT apply repayment deduction records – she wants them placed on account (for now) – she may want me to change the code back later
                    End If

                    Dim rowARTPYMT3s() As DataRow = dst.Tables("ARTPYMT3").Select(sqlc)

                    If rowARTPYMT3s.Length = 0 Then
                        Dim SQLX As String = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
                                             & " and EDI_ENT_NO = " & rowEDT820T4.Item("EDI_ENT_NO") _
                                             & " and EDI_ADJ_SEQ = " & Val(rowEDT820T4.Item("EDI_ADJ_SEQ") & "") + 1 _
                                             & " and ADJ_AMT = 0 and ADJ_REF_NO = '" & rowEDT820T4.Item("ADJ_REF_NO") & "'"
                        Dim rowEDT820T4next() As DataRow = dst.Tables("EDT820T4").Select(SQLX)

                        If rowEDT820T4next.Length = 1 Then
                            Dim T4_ADJ_DESC_next As String = rowEDT820T4next(0).Item("ADJ_DESC") & ""
                            T4_ADJ_DESC_next = Replace(T4_ADJ_DESC_next, "'", "") ' DID NOT PROVIDE UPC'S ELECTRONICALLY
                            Dim T4_ADJ_DESC_next10 As String = ""
                            If macys_accounts And T4_ADJ_DESC_next.StartsWith("RA   = ") Then
                                T4_ADJ_DESC_next = Mid(T4_ADJ_DESC_next, "RA   = ".Length + 1)
                                If T4_ADJ_DESC_next.Length > 10 Then
                                    T4_ADJ_DESC_next10 = Mid(T4_ADJ_DESC_next, T4_ADJ_DESC_next.Length - 10 + 1, 10)
                                End If
                                macys_reference = T4_ADJ_DESC_next
                                If T4_ADJ_DESC_next10 <> "" Then macys_reference = T4_ADJ_DESC_next10

                            End If

                            sqlc = "INV_TYPE IN (" & ZZ & ")" _
                                & " and (INV_CUST_PO = '" & T4_ADJ_DESC_next & "'" & IIf(T4_ADJ_DESC_next10 <> "", " OR INV_CUST_PO = '" & T4_ADJ_DESC_next10 & "'", "") & ")"


                            If macys_accounts And ASCMAIN1.CLIENT = "INT" And (T4_ADJ_DESC.StartsWith("CREDIT   RTV FREIGHT") Or T4_ADJ_DESC.StartsWith("REVERSE RTV HANDLING")) Then
                                sqlc &= " AND 1<>1" ' 07/27/16 this code appears twice in this section - Petra wants (for MACYS only) for ABSolution to NOT apply repayment deduction records – she wants them placed on account (for now) – she may want me to change the code back later
                            End If

                            rowARTPYMT3s = dst.Tables("ARTPYMT3").Select(sqlc)

                        End If
                    End If


                    If rowARTPYMT3s.Length <> 0 Then
                        ctr820T4_matched = ctr820T4_matched + 1
                    End If

                    For Each rowARTPYMT3 As DataRow In rowARTPYMT3s
                        Dim INV_BALANCE_NEW As Decimal = Val(rowARTPYMT3.Item("INV_BALANCE_NEW") & "")
                        If ASCMAIN1.Running_in_VS AndAlso rowARTPYMT3.Item("INV_NUM") & "" = "0020196170" Then Stop
                        If ADJ_AMT <> 0 And INV_BALANCE_NEW <> 0 Then
                            Dim ADJ_TRAN_AMT As Decimal = 0
                            If INV_BALANCE_NEW > 0 Then
                                If INV_BALANCE_NEW > -1 * ADJ_AMT Then
                                    ADJ_TRAN_AMT = ADJ_AMT
                                Else
                                    ADJ_TRAN_AMT = -1 * INV_BALANCE_NEW
                                End If
                            Else
                                If -1 * INV_BALANCE_NEW > -1 * ADJ_AMT Then
                                    ADJ_TRAN_AMT = ADJ_AMT
                                Else
                                    ADJ_TRAN_AMT = INV_BALANCE_NEW
                                End If
                            End If

                            PYMT_BATCH_ILNO = PYMT_BATCH_ILNO + 1

                            With rowARTPYMT3
                                .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + ADJ_TRAN_AMT * CURR_EXCH_RATE
                                .Item("INV_DISC_TAKEN") = 0
                                .Item("INV_WRITE_OFF") = 0
                                .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - ADJ_TRAN_AMT * CURR_EXCH_RATE
                                .Item("INV_PMT_CURR") = Val(.Item("INV_PMT_CURR") & "") + ADJ_TRAN_AMT
                                .Item("INV_DISC_TAKEN_CURR") = 0
                                .Item("INV_WRITE_OFF_CURR") = 0
                                .Item("INV_BALANCE_NEW_CURR") = Val(.Item("INV_BALANCE_NEW_CURR") & "") - ADJ_TRAN_AMT
                            End With

                            ADJ_AMT = ADJ_AMT - ADJ_TRAN_AMT
                            T4 = T4 - ADJ_TRAN_AMT
                        End If
                    Next

                    rowARTPYMT5_prior = Nothing

                    If ADJ_AMT <> 0 Then
                        PYMT_BATCH_DLNO_ctr = PYMT_BATCH_DLNO_ctr + 1

                        Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
                        With rowARTPYMT5
                            .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                            .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                            .Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO_ctr

                            'ADJ_REAS_CODE = does this get set somewhere in here?
                            REASON_CODE = Get_REASON_CODE(EDI_TP_QUAL, EDI_TP_ID, ADJ_REAS_CODE)

                            .Item("REASON_CODE") = REASON_CODE
                            If REASON_CODE <> "" Then
                                .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
                            Else
                                .Item("REASON_DESC") = ADJ_REAS_CODE & ": Unmapped Code"
                            End If
                            .Item("ACCT_CODE") = DBNull.Value
                            .Item("GL_DIST_AMT") = ADJ_AMT * CURR_EXCH_RATE * -1
                            .Item("GL_DIST_COMMENT") = rowEDT820T4.Item("ADJ_DESC")
                            If System.Math.Abs(ADJ_AMT * -1 * CURR_EXCH_RATE) <= AR_PARM_AUTO_WOFF Then
                                .Item("CHARGEBACK_IND") = "0"
                            Else
                                .Item("CHARGEBACK_IND") = "1"
                            End If
                            .Item("CHARGEBACK_NO") = DBNull.Value
                            .Item("CUST_REFERENCE") = T4_ADJ_REF

                            '
                            'If ASCMAIN1.Running_in_VS And T4_ADJ_REF.EndsWith("23476109") Then Stop
                            'If ASCMAIN1.Running_in_VS And T4_ADJ_REF.EndsWith("6121") Then Stop
                            '   If ASCMAIN1.Running_in_VS And T4_ADJ_REF.EndsWith("00003791482") Then Stop

                            If rowEDTXREF2 IsNot Nothing Then
                                .Item("CUST_CODE_SO") = rowEDTXREF2.Item("CUST_CODE_SO")
                            End If

                            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                            .Item("INV_TYPE_CB") = DBNull.Value
                            .Item("OUR_REFERENCE") = DBNull.Value
                            .Item("GL_DIST_AMT_CURR") = ADJ_AMT * -1
                            T4 = T4 + ADJ_AMT * -1

                            If macys_accounts And macys_reference <> "" Then
                                .Item("CUST_REFERENCE") = macys_reference
                                .Item("OUR_REFERENCE") = T4_ADJ_REF
                            End If
                        End With
                        dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)
                        rowARTPYMT5_prior = rowARTPYMT5
                    End If
                End If

                rowEDT820T4_prior = rowEDT820T4
            Next
        Next

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default

        Dim zMSG As String = ""

        If dst.Tables("EDTERRS1").Rows.Count <> 0 Then
            Using F As New ASFMSGBF
                F.Show_grd(dst.Tables("EDTERRS1"), Me, "Un-Mapped Deduction Codes (Default 000 was used)")
            End Using
            zMSG = "You will Not be able to Update until you have Mapped the Deduction Codes and Re-Processed this Payment"
        Else
            zMSG = "Make Edits as Required and then Update"
        End If

        '   Setup_Displays()

        MsgBox("Auto-Application of EDI Remittance has Completed" _
               & vbCrLf & vbCrLf & CStr(ctr820T2) & " A/R Items Indicated on Remittance, " _
               & CStr(ctr820T2_nogo) & " Items Not Matched" _
               & vbCrLf & CStr(ctr820T4) & " Deductions Indicated in Remittance Advice, " _
               & CStr(ctr820T4_matched) & " Matched against Credits Anticipating Deduction" _
               & vbCrLf & vbCrLf & zMSG, vbOKOnly, "Verification")

        If dst.Tables("EDTXREF1").Select("", "", DataViewRowState.Added).Length <> 0 Then
            Update_Record_TDA("EDTXREF1")
        End If

        Calculate_Application_by_Type()
        Display_Application_Totals()

        Return (dst.Tables("EDTERRS1").Rows.Count = 0)

    End Function

    Function Get_REASON_CODE(EDI_SENDER_QUAL As String, EDI_SENDER_ID As String, ADJ_REAS_CODE As String) As String

        Dim REASON_CODE As String = "000"

        Dim rowEDTXREF1 As DataRow = dst.Tables("EDTXREF1").Rows.Find _
                                     (New Object() {EDI_SENDER_QUAL, EDI_SENDER_ID, ADJ_REAS_CODE})
        If rowEDTXREF1 Is Nothing Then
            If dst.Tables("EDTERRS1").Rows.Find(New String() {"ADJ_REASON_CODE", ADJ_REAS_CODE}) Is Nothing Then
                dst.Tables("EDTERRS1").Rows.Add(New String() {"ADJ_REASON_CODE", ADJ_REAS_CODE, "Missing X_Ref for " & CUST_CODE & ", " & EDI_SENDER_QUAL & ":" & EDI_SENDER_ID})
                dst.Tables("EDTXREF1").Rows.Add(EDI_SENDER_QUAL, EDI_SENDER_ID, ADJ_REAS_CODE)
            End If
        Else
            REASON_CODE = rowEDTXREF1.Item("REASON_CODE") & ""
            rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
            If rowARTREAS1 Is Nothing Then
                If dst.Tables("EDTERRS1").Rows.Find(New String() {"ADJ_REASON_CODE", ADJ_REAS_CODE}) Is Nothing Then
                    dst.Tables("EDTERRS1").Rows.Add(New String() {"ADJ_REASON_CODE", ADJ_REAS_CODE, "Missing X_Ref for " & CUST_CODE & ", " & EDI_SENDER_QUAL & ":" & EDI_SENDER_ID})
                End If
            End If
        End If

        Return REASON_CODE
    End Function

    Private Sub Load_EDI_Grid()

        If EDI_TABLES Is Nothing Then
            ReDim EDI_TABLES(7)
            For I As Integer = 1 To 7
                ASCMAIN1.sql = "Select * from EDT820T" & CStr(I) & " where ROWNUM < 1"
                EDI_TABLES(I) = ASCMAIN1.Temp_Table
            Next I
        End If

        BeginTrans()
        ASCDATA1.ExecuteSQL("Update EDT820T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_PROCESS_IND is Null")

        ASCMAIN1.sql = "Update EDT820T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
               & " where EDI_OUR_ID = TRIM(EDT820T1.EDI_OUR_ID) and EDI_TP_ID = TRIM(EDT820T1.EDI_TP_ID))" _
               & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
        ASCDATA1.ExecuteSQL()
        'ASCMAIN1.sql = "Update EDT820T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM2" _
        '       & " where EDI_TP_QUAL = EDT820T1.EDI_TP_QUAL and EDI_TP_ID = EDT820T1.EDI_TP_ID and EDI_DOC_NO = '820' and EDI_DEPT_NO = EDT820T1.EDI_DEPARTMENT)" _
        '       & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        'ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update EDT820T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
               & " where EDI_TP_QUAL = TRIM(EDT820T1.EDI_TP_QUAL) and EDI_TP_ID = TRIM(EDT820T1.EDI_TP_ID) and EDI_DOC_NO = '820')" _
               & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from EDT820T1 where CUST_CODE = 'MACYS'" & vbCrLf _
            & " And EDI_PROCESS_IND = '0' and SUBSTR(EDI_DOC_SEQ_NO,1,1) between '0' and '9'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim EDI_DOC_SEQ_NO As String = row.Item("EDI_DOC_SEQ_NO")
            Split_MACYS_by_Entity(EDI_DOC_SEQ_NO)
        Next

        ASCMAIN1.sql = "Select * from EDT820T1 where CUST_CODE = 'KMART'" & vbCrLf _
            & " And EDI_PROCESS_IND = '0' and SUBSTR(EDI_DOC_SEQ_NO,1,1) between '0' and '9'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim EDI_DOC_SEQ_NO As String = row.Item("EDI_DOC_SEQ_NO")
            Split_KMART_by_Entity(EDI_DOC_SEQ_NO)
        Next

        CommitTrans()

        Fill_Records("EDT820TX")
        Sort_grdColumns(grdEDT820TX, "CUST_CODE")

        BeginTrans()
        ASCDATA1.ExecuteSQL("Update EDT864T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_PROCESS_IND is Null")
        ASCMAIN1.sql = "Update EDT864T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
               & " where EDI_OUR_ID = TRIM(EDT864T1.EDI_OUR_ID) and EDI_TP_ID = TRIM(EDT864T1.EDI_TP_ID))" _
               & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
        ASCDATA1.ExecuteSQL()
        'ASCMAIN1.sql = "Update EDT864T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM2" _
        '       & " where EDI_TP_QUAL = EDT864T1.EDI_TP_QUAL and EDI_TP_ID = EDT864T1.EDI_TP_ID and EDI_DOC_NO = '864' and EDI_DEPT_NO = EDT864T1.EDI_DEPARTMENT)" _
        '       & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        'ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update EDT864T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
               & " where EDI_TP_QUAL = EDT864T1.EDI_TP_QUAL and EDI_TP_ID = EDT864T1.EDI_TP_ID and EDI_DOC_NO = '864')" _
               & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()
        CommitTrans()


        'dillards used to record chargeback data in the 864

        'BeginTrans()

        'ASCMAIN1.sql = "" _
        '    & "BEGIN DECLARE CURSOR C1 IS " & vbCrLf _
        '    & " SELECT * FROM EDT864T3 WHERE EDI_DOC_SEQ_NO IN" & vbCrLf _
        '    & " (SELECT EDI_DOC_SEQ_NO FROM EDT864T1 WHERE EDI_DIL_XREF IS NULL)" & vbCrLf _
        '    & " AND EDI_MSG_TEXT LIKE '%CHARGEBACK#%';" & vbCrLf _
        '    & " O EDT864T3%ROWTYPE;" & vbCrLf _
        '    & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
        '    & " SELECT EDT864T3.* INTO O FROM EDT864T3 WHERE EDI_DOC_SEQ_NO = R1.EDI_DOC_SEQ_NO" & vbCrLf _
        '    & " AND EDI_DTL_SEQ = R1.EDI_DTL_SEQ AND EDI_MSG_SEQ = R1.EDI_MSG_SEQ + 2;" & vbCrLf _
        '    & " IF NOT SQL%NOTFOUND THEN" & vbCrLf _
        '    & " If Trim(SUBSTR(o.EDI_MSG_TEXT, 67, 10)) Is Not Null Then" & vbCrLf _
        '    & " INSERT INTO EDT820TD VALUES (SUBSTR(R1.EDI_MSG_TEXT,17,10)," & vbCrLf _
        '    & " SUBSTR(O.EDI_MSG_TEXT,20,10), SUBSTR(O.EDI_MSG_TEXT,67,10));" & vbCrLf _
        '    & " END IF;" & vbCrLf _
        '    & " END IF;" & vbCrLf _
        '    & " END LOOP;" & vbCrLf _
        '    & " UPDATE EDT864T1 SET EDI_DIL_XREF = '1' WHERE EDI_DIL_XREF IS NULL;" & vbCrLf _
        '    & " END; END;"
        'ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Update EDT820T4" & vbCrLf _
        '    & " SET ADJ_CUST_PO = (SELECT DIL_VEND_XREF FROM EDT820TD" & vbCrLf _
        '    & " WHERE DIL_CHARGEBACK_NO = EDT820T4.ADJ_REF_NO)" & vbCrLf _
        '    & " WHERE EDI_JRNL_NO IN (SELECT EDI_JRNL_NO FROM EDT820T1 WHERE EDI_PROCESS_IND IS NULL)" & vbCrLf _
        '    & " AND ADJ_CUST_PO IS NULL"
        'OraD.ExecuteSQL Sql

        'CommitTrans()

    End Sub

#End Region

    Private Sub grdEDT820TX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT820TX.DoubleClickRow
        If Not ScreenMode And grdEDT820TX.ActiveRow IsNot Nothing Then Click_Command("Process 820")
    End Sub

    Sub Split_MACYS_by_Entity(EDI_DOC_SEQ_NO As String)

        'BeginTrans()

        Dim r As Integer = 0

        For I As Integer = 1 To 7
            Dim T As String = EDI_TABLES(I)

            ASCMAIN1.sql = "Delete from " & T
            ASCDATA1.ExecuteSQL()
            Dim sqlI As String = "Insert into " & T & " " _
                & "Select * from EDT820T" & CStr(I) _
                & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"

            If I = 1 Then
                Dim sqlx As String = "" _
                    & "Select EDT820T2.EDI_ENT_NO, TRIM(EDT820T2.EDI_ID) EDI_ID, TRIM(EDT820T1.EDI_TP_QUAL) EDI_TP_QUAL, TRIM(EDT820T1.EDI_TP_ID) EDI_TP_ID, EDT820T1.CUST_CODE" & vbCrLf _
                    & " from EDT820T2,EDT820T1" & vbCrLf _
                    & " where EDT820T1.EDI_DOC_SEQ_NO = EDT820T2.EDI_DOC_SEQ_NO" & vbCrLf _
                    & "   and EDT820T2.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                ASCMAIN1.sql = "Select EDT820TX.EDI_ENT_NO, NVL(EDTXREF2.CUST_CODE,EDT820TX.CUST_CODE) CUST_CODE" & vbCrLf _
                    & " from (" & sqlx & ") EDT820TX,EDTXREF2" & vbCrLf _
                    & " where EDTXREF2.EDI_TP_QUAL (+) = EDT820TX.EDI_TP_QUAL" & vbCrLf _
                    & "   and EDTXREF2.EDI_TP_ID (+) = EDT820TX.EDI_TP_ID" & vbCrLf _
                    & "   and EDTXREF2.EDI_ENT_ID (+) = EDT820TX.EDI_ID"

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.CLIENT = "VAN" Then
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDTXREF2.EDI_ENT_ID", "EDTXREF2.EDI_ENT_CODE")
                End If

                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "EDI_ENT_NO")
                    Dim EDI_ENT_NO As String = row.Item("EDI_ENT_NO")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    ASCDATA1.ExecuteSQL(sqlI)
                    ASCMAIN1.sql = "Update " & T _
                        & " Set EDI_DOC_SEQ_NO = 'M' || '" & EDI_ENT_NO & Mid(EDI_DOC_SEQ_NO, 3) & "'" & vbCrLf _
                        & ", PYMT_AMT = 0, CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    ASCDATA1.ExecuteSQL()
                Next

                ASCMAIN1.sql = "Update EDT820T" & CStr(I) & " Set EDI_PROCESS_IND = 'S' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_PROCESS_IND = '0'"
                r = ASCDATA1.ExecuteSQL()

            Else
                ASCDATA1.ExecuteSQL(sqlI)
                ASCMAIN1.sql = "Update " & T _
                    & " Set EDI_DOC_SEQ_NO = 'M' || TRIM(TO_CHAR(EDI_ENT_NO)) || '" & Mid(EDI_DOC_SEQ_NO, 3) & "'"
                ASCDATA1.ExecuteSQL()
            End If

            ASCMAIN1.sql = "Insert into EDT820T" & CStr(I) & " Select * from " & T
            ASCDATA1.ExecuteSQL()

        Next

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor c1 is " & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (AMT) AMT FROM (" & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (AMT_GROSS) AMT" & vbCrLf _
            & " FROM EDT820T3 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
            & "GROUP BY EDI_ENT_NO" & vbCrLf _
            & "UNION" & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (ADJ_AMT) AMT" & vbCrLf _
            & " FROM EDT820T4 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
            & "GROUP BY EDI_ENT_NO" & vbCrLf _
            & "UNION" & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (ADJ_AMT) AMT" & vbCrLf _
            & " FROM EDT820T5 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
            & "GROUP BY EDI_ENT_NO" & vbCrLf _
            & ") GROUP By EDI_ENT_NO;" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & "  Update EDT820T1 Set PYMT_AMT = R1.AMT" & vbCrLf _
            & "   where EDI_DOC_SEQ_NO = 'M' || R1.EDI_ENT_NO || '" & Mid(EDI_DOC_SEQ_NO, 3) & "';" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()



        'If r = 1 Then
        '    CommitTrans()
        'Else
        '    Rollback()
        'End If

    End Sub

    Sub Split_KMART_by_Entity(EDI_DOC_SEQ_NO As String)

        'BeginTrans()

        If 1 = 1 Then
            Exit Sub
        End If

        Dim r As Integer = 0

        For I As Integer = 1 To 7
            Dim T As String = EDI_TABLES(I)

            ASCMAIN1.sql = "Delete from " & T
            ASCDATA1.ExecuteSQL()
            Dim sqlI As String = "Insert into " & T & " " _
                & "Select * from EDT820T" & CStr(I) _
                & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"

            If I = 1 Then
                Dim sqlx As String = "" _
                    & "Select EDT820T2.EDI_ENT_NO, TRIM(EDT820T2.EDI_ID) EDI_ID, TRIM(EDT820T1.EDI_TP_QUAL) EDI_TP_QUAL, TRIM(EDT820T1.EDI_TP_ID) EDI_TP_ID, EDT820T1.CUST_CODE" & vbCrLf _
                    & " from EDT820T2,EDT820T1" & vbCrLf _
                    & " where EDT820T1.EDI_DOC_SEQ_NO = EDT820T2.EDI_DOC_SEQ_NO" & vbCrLf _
                    & "   and EDT820T2.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                ASCMAIN1.sql = "Select EDT820TX.EDI_ENT_NO, NVL(EDTXREF2.CUST_CODE,EDT820TX.CUST_CODE) CUST_CODE" & vbCrLf _
                    & " from (" & sqlx & ") EDT820TX,EDTXREF2" & vbCrLf _
                    & " where EDTXREF2.EDI_TP_QUAL (+) = EDT820TX.EDI_TP_QUAL" & vbCrLf _
                    & "   and EDTXREF2.EDI_TP_ID (+) = EDT820TX.EDI_TP_ID" & vbCrLf _
                    & "   and EDTXREF2.EDI_ENT_ID (+) = EDT820TX.EDI_ID"

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.CLIENT = "VAN" Then
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDTXREF2.EDI_TP_QUAL", "EDTXREF2.SENDER_ID_QUAL")
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDTXREF2.EDI_TP_ID", "EDTXREF2.SENDER_ID")
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDTXREF2.EDI_ENT_ID", "EDTXREF2.EDI_ENT_CODE")
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "EDTXREF2.CUST_CODE", "EDTXREF2.CUST_CODE_SO")

                End If

                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "EDI_ENT_NO")
                    Dim EDI_ENT_NO As String = row.Item("EDI_ENT_NO")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    ASCDATA1.ExecuteSQL(sqlI)
                    ASCMAIN1.sql = "Update " & T _
                        & " Set EDI_DOC_SEQ_NO = 'M' || '" & EDI_ENT_NO & Mid(EDI_DOC_SEQ_NO, 3) & "'" & vbCrLf _
                        & ", PYMT_AMT = 0, CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    ASCDATA1.ExecuteSQL()
                Next

                ASCMAIN1.sql = "Update EDT820T" & CStr(I) & " Set EDI_PROCESS_IND = 'S' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_PROCESS_IND = '0'"
                r = ASCDATA1.ExecuteSQL()

            Else
                ASCDATA1.ExecuteSQL(sqlI)
                ASCMAIN1.sql = "Update " & T _
                    & " Set EDI_DOC_SEQ_NO = 'M' || TRIM(TO_CHAR(EDI_ENT_NO)) || '" & Mid(EDI_DOC_SEQ_NO, 3) & "'"
                ASCDATA1.ExecuteSQL()
            End If

            ASCMAIN1.sql = "Insert into EDT820T" & CStr(I) & " Select * from " & T
            ASCDATA1.ExecuteSQL()

        Next

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor c1 is " & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (AMT) AMT FROM (" & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (AMT_GROSS) AMT" & vbCrLf _
            & " FROM EDT820T3 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
            & "GROUP BY EDI_ENT_NO" & vbCrLf _
            & "UNION" & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (ADJ_AMT) AMT" & vbCrLf _
            & " FROM EDT820T4 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
            & "GROUP BY EDI_ENT_NO" & vbCrLf _
            & "UNION" & vbCrLf _
            & "SELECT EDI_ENT_NO, SUM (ADJ_AMT) AMT" & vbCrLf _
            & " FROM EDT820T5 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
            & "GROUP BY EDI_ENT_NO" & vbCrLf _
            & ") GROUP By EDI_ENT_NO;" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & "  Update EDT820T1 Set PYMT_AMT = R1.AMT" & vbCrLf _
            & "   where EDI_DOC_SEQ_NO = 'M' || R1.EDI_ENT_NO || '" & Mid(EDI_DOC_SEQ_NO, 3) & "';" & vbCrLf _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()



        'If r = 1 Then
        '    CommitTrans()
        'Else
        '    Rollback()
        'End If

    End Sub

    Private Sub cmdApplyXLS_Click(sender As Object, e As EventArgs) Handles cmdApplyXLS.Click

        If dst.Tables("ARTPYMT5").Select("").Length <> 0 Or dst.Tables("ARTPYMT4").Select("").Length <> 0 Or dst.Tables("ARTPYMT3").Select("INV_PMT<> 0").Length <> 0 Then
            MsgBox("This function is only applicable at the start of Payment Application", MsgBoxStyle.OkOnly, "Cannot Apply XLS if application has been started")
            Exit Sub
        End If

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            Dim filter As String = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            If ASCMAIN1.CLIENT = "RGI" And HFs("CUST_CODE") = "031013" Then
                filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt"
            End If
            If ASCMAIN1.CLIENT = "RGI" And (HFs("CUST_CODE") = "110451" Or HFs("CUST_CODE") = "311823") Then
                filter = "CTX 820 Text files (*.CTX820.txt)|*.CTX820.txt| xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls3"
            End If
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True

            '  Excel_Import = -1

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        Try

            Dim Vs As New Dictionary(Of String, Integer)

            If FILENAME <> "" Then
                Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                Dim range As SpreadsheetGear.IRange = Nothing

                Dim PYMT_BATCH_DLNO As Integer = 0
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                Dim BAD_REASON_CODEs As New List(Of String)
                Dim rowARTREAS1 As DataRow = Nothing

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Applying from Spreadsheet")


                If ASCMAIN1.CLIENT = "AHA" Then

                    Dim r As Integer = 0 ' NO HEADING
                    Do While oSheet.Cells(r, 0).Value & "" <> ""
                        Dim INV_NUM As String = ""
                        Dim INV_REF As String = Trim(oSheet.Cells(r, 1).Value & "")
                        Dim INV_REF_DTvalue As Int64 = Val(Trim(oSheet.Cells(r, 2).Value & ""))
                        If INV_REF_DTvalue > 59 Then INV_REF_DTvalue -= 1 '  Excel/Lotus 2/29/1900 bug   

                        Dim INV_REF_DT As Date = New DateTime(1899, 12, 31).AddDays(INV_REF_DTvalue)
                        Dim INV_PMTUSD As String = Replace(Replace(Replace(Trim(oSheet.Cells(r, 3).Value & ""), "USD", ""), "$", ""), ",", "")

                        Dim INV_PMT As Decimal = Val(INV_PMTUSD)
                        Dim record_processed As Boolean = False


                        If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_REF)
                        r += 1

                        If INV_PMT > 0 Then ' PAID ITEM
                            Dim INV_NUM_cleaned As String = INV_REF
                            Dim Sql As String = "INV_NUM = '" & INV_NUM_cleaned.PadLeft(10, "0") & "'"
                            Sql &= " and INV_BALANCE = INV_TOTAL_AMOUNT_CURR"
                            Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)
                            If rows.Length = 1 Then
                                ' RECORD A PAYMENT TO THE INVOICE
                                Dim rowARTPYMT3 As DataRow = rows(0)
                                With rowARTPYMT3

                                    Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                                    Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                    Dim INV_PMT_invoice As Decimal = INV_PMT
                                    Dim INV_DSC_invoice As Decimal = INV_TOTAL_AMT * 0.07

                                    .Item("INV_PMT") = INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                    .Item("INV_DISC_TAKEN") = 0
                                    .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice

                                    .Item("INV_BALANCE_NEW") = 0 ' Val(.Item("INV_BALANCE_NEW") & "") - (INV_PMT_invoice + INV_DSC_invoice)  
                                    .Item("INV_PMT_CURR") = INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice

                                    .Item("INV_DISC_TAKEN_CURR") = 0
                                    .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice

                                    .Item("INV_BALANCE_NEW_CURR") = 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)

                                    Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice
                                    If System.Math.Abs(INV_BALANCE_NEW) < 0.05 Then
                                        INV_DSC_invoice += INV_BALANCE_NEW
                                        INV_BALANCE_NEW = 0
                                    End If
                                    If INV_DSC_invoice <> 0 Then
                                        Record_Chargeback(PYMT_BATCH_DLNO, "CM", -1 * INV_DSC_invoice, INV_NUM_cleaned)
                                    End If

                                    If INV_BALANCE_NEW <> 0 Then
                                        Record_Chargeback(PYMT_BATCH_DLNO, "CB", -1 * INV_BALANCE_NEW, INV_NUM_cleaned)
                                    End If

                                End With

                                record_processed = True
                            End If

                        End If

                        If INV_PMT < 0 Or Not record_processed Then ' DEDUCTION

                            Dim REF As String = INV_REF
                            Dim LNO As Integer = -1

                            Dim REASON_CODE As String = "CB"

                            If CUST_CODE = "ULTA" And REF.StartsWith("V0000") Then
                                REASON_CODE = "RM"

                                REF = Mid(REF, 6, 5)
                                If Vs.ContainsKey(Mid(INV_REF, 1, 10)) Then
                                    LNO = Vs(Mid(INV_REF, 1, 10))
                                Else
                                    Vs.Add(Mid(INV_REF, 1, 10), PYMT_BATCH_DLNO + 1)
                                End If
                            End If

                            ' RECORD A CHARGEBACK
                            If LNO <> -1 Then
                                ' ACCUMULATE
                                Dim row As DataRow = dst.Tables("ARTPYMT5").Rows.Find(New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO"), LNO})
                                With row
                                    .Item("GL_DIST_AMT_CURR") = Val(.Item("GL_DIST_AMT_CURR") & "") - 1 * INV_PMT
                                    .Item("GL_DIST_AMT") = Val(.Item("GL_DIST_AMT") & "") - 1 * INV_PMT
                                End With
                            Else
                                Dim rowARTPYMT5 As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, REASON_CODE, INV_PMT, REF)
                            End If
                        End If
                    Loop

                ElseIf ASCMAIN1.CLIENT = "RGI" Then

                    If HFs("CUST_CODE") = "031013" Then

                        Dim TXT As String = ""
                        Using SW As New System.IO.StreamReader(FILENAME)
                            TXT = SW.ReadToEnd()
                        End Using

                        Dim CHECK_NUM As String = ""

                        Dim r As Int64 = 0
                        Dim TOTAL_APPLIED As Decimal = 0
                        Dim skip As Integer = 0

                        Dim lines() As String

                        If TXT.Contains(vbCrLf) Then
                            lines = Split(TXT, vbCrLf)
                        ElseIf TXT.Contains(Chr(10)) Then
                            lines = Split(TXT, Chr(10))
                        Else
                            lines = Split(TXT, vbCrLf)
                        End If

                        For Each line As String In lines

                            Dim fields() As String = Split(line, ",")

                            If skip > 0 Then
                                skip -= 1

                            ElseIf line.StartsWith("Wayfair Voucher #:") Then
                                CHECK_NUM = Replace(line, "Wayfair Voucher #: ", "")

                            ElseIf line.StartsWith("Wayfair Remittance #:") Then
                                CHECK_NUM = Replace(line, "Wayfair Remittance #:  ", "")

                            ElseIf line.Length > 30 AndAlso line.StartsWith("Allowance for Damages/Defects/Returns") Then

                                Dim ALLOW_PCT As Decimal = Val(fields(2))
                                Dim ALLOW_AMT As Decimal = TOTAL_APPLIED * ALLOW_PCT

                                Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW", 1 * Val(fields(3)), fields(2) & " Allowance ")
                                row.Item("CHARGEBACK_IND") = "0"

                            ElseIf line.Length > 10 AndAlso line.StartsWith(",,Total:") Then

                                Record_Chargeback(PYMT_BATCH_DLNO, "ONA", -1 * fields(3), CHECK_NUM)


                            ElseIf line.Length > 10 AndAlso line.StartsWith("Credit") Then

                                Record_Chargeback(PYMT_BATCH_DLNO, "MISC", 1 * Val(fields(3)), fields(2))

                            ElseIf line.Length > 10 AndAlso (line.StartsWith("Item: ") Or line.StartsWith("Customer: ") Or line.StartsWith("Desc: ")) Then

                            ElseIf line.Length > 10 AndAlso line.StartsWith("Invoice #, PO #") Then
                                ' heading line

                            ElseIf line.Length > 20 AndAlso line.Contains(",") Then

                                Dim INV_NUM As String = fields(0)
                                'If INV_NUM = "CS453626254" Then
                                '    Stop
                                'End If
                                Dim INV_CUST_PO As String = fields(1)
                                If INV_NUM.Length < 10 Then
                                    INV_NUM = INV_NUM.PadLeft(10, "0")
                                End If
                                Dim INV_PMT As Decimal = Val(fields(3))
                                Dim record_processed As Boolean = False

                                If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_NUM)
                                r += 1

                                If INV_PMT > 0 Or INV_PMT < 0 Then ' PAID ITEM - ADDED <0 BECAUSE THERE ARE NEGATIVE INVOICES

                                    Dim Sql As String = "INV_NUM = '" & INV_NUM & "'"
                                    'Sql &= " and INV_BALANCE = INV_TOTAL_AMOUNT_CURR"
                                    Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)
                                    If rows.Length = 0 Then
                                        Sql = "INV_CUST_PO = '" & INV_CUST_PO & "'"
                                        rows = dst.Tables("ARTPYMT3").Select(Sql)
                                    End If

                                    If rows.Length = 1 Then
                                        ' RECORD A PAYMENT TO THE INVOICE
                                        Dim rowARTPYMT3 As DataRow = rows(0)
                                        With rowARTPYMT3

                                            Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                                            Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                            Dim INV_PMT_invoice As Decimal = INV_PMT
                                            Dim INV_DSC_invoice As Decimal = 0
                                            .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                            .Item("INV_DISC_TAKEN") = 0
                                            .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice

                                            .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                                            .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice

                                            .Item("INV_DISC_TAKEN_CURR") = 0
                                            .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice

                                            .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)

                                            TOTAL_APPLIED += INV_PMT

                                            Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice
                                            'If System.Math.Abs(INV_BALANCE_NEW) < 0.05 Then
                                            '    INV_DSC_invoice += INV_BALANCE_NEW
                                            '    INV_BALANCE_NEW = 0
                                            'End If
                                            'If INV_DSC_invoice <> 0 Then
                                            '    Record_Chargeback(PYMT_BATCH_DLNO, "CM", -1 * INV_DSC_invoice, INV_NUM_cleaned)
                                            'End If

                                            'If INV_BALANCE_NEW <> 0 Then
                                            '    Record_Chargeback(PYMT_BATCH_DLNO, "CB", -1 * INV_BALANCE_NEW, INV_NUM_cleaned)
                                            'End If

                                        End With

                                        record_processed = True
                                    Else
                                        Record_Chargeback(PYMT_BATCH_DLNO, "ONA", (INV_PMT), INV_NUM, , "Cannot Find AR Item")
                                    End If

                                End If
                            End If
                        Next

                        'ElseIf HFs("CUST_CODE") = "031013" Then

                        '    Dim TXT As String = ""
                        '    Using SW As New System.IO.StreamReader(FILENAME)
                        '        TXT = SW.ReadToEnd()
                        '    End Using

                        '    Dim CHECK_NUM As String = ""

                        '    Dim r As Int64 = 0
                        '    Dim TOTAL_APPLIED As Decimal = 0
                        '    Dim skip As Integer = 0

                        '    Dim lines() As String

                        '    If TXT.Contains(vbCrLf) Then
                        '        lines = Split(TXT, vbCrLf)
                        '    ElseIf TXT.Contains(Chr(10)) Then
                        '        lines = Split(TXT, Chr(10))
                        '    Else
                        '        lines = Split(TXT, vbCrLf)
                        '    End If

                        '    For Each line As String In lines

                        '        Dim fields() As String = Split(line, ",")

                        '        If skip > 0 Then
                        '            skip -= 1

                        '        ElseIf line.StartsWith("Wayfair Voucher #:") Then
                        '            CHECK_NUM = Replace(line, "Wayfair Voucher #: ", "")

                        '        ElseIf line.StartsWith("Wayfair Remittance #:") Then
                        '            CHECK_NUM = Replace(line, "Wayfair Remittance #:  ", "")

                        '        ElseIf line.Length > 30 AndAlso line.StartsWith("Allowance for Damages/Defects/Returns") Then
                        '            ' need to see this in new format
                        '            Dim ALLOW_PCT As Decimal = Val(fields(7))
                        '            Dim ALLOW_AMT As Decimal = TOTAL_APPLIED * ALLOW_PCT

                        '            Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW", 1 * Val(fields(8)), fields(7) & " Allowance ")
                        '            row.Item("CHARGEBACK_IND") = "0"

                        '        ElseIf line.Length > 10 AndAlso line.StartsWith(",,,,,,,Total (USD):") Then

                        '            ' Record_Chargeback(PYMT_BATCH_DLNO, "ONA", -1 * fields(8), CHECK_NUM)
                        '            Stop

                        '        ElseIf line.Length > 10 AndAlso line.StartsWith("Total (USD):") Then

                        '            Record_Chargeback(PYMT_BATCH_DLNO, "ONA", -1 * fields(1), CHECK_NUM)


                        '        ElseIf line.Length > 10 AndAlso line.StartsWith("Deduction") Then

                        '            Record_Chargeback(PYMT_BATCH_DLNO, "ONA", fields(8), fields(1))


                        '        ElseIf line.Length > 10 AndAlso line.StartsWith("Credit") Then

                        '            Record_Chargeback(PYMT_BATCH_DLNO, "MISC", 1 * Val(fields(8)), fields(7))

                        '        ElseIf line.Length > 10 AndAlso (line.StartsWith("Item: ") Or line.StartsWith("Customer: ") Or line.StartsWith("Desc: ")) Then

                        '        ElseIf line.Length > 10 AndAlso line.StartsWith(",,Sub-total:") Then

                        '        ElseIf line.Length > 10 AndAlso line.StartsWith("Invoice #, PO #") Then
                        '            ' heading line
                        '        ElseIf line.Length > 10 AndAlso line.StartsWith("100 Huntington Avenue") Then
                        '        ElseIf line.Length > 10 AndAlso line.StartsWith("4 Copley Place") Then
                        '        ElseIf line.Length > 10 AndAlso line.StartsWith("Boston") Then
                        '        ElseIf line.Length > 20 And line.Length < 100 AndAlso line.Contains(",") Then

                        '            Dim INV_NUM As String = fields(0)
                        '            Dim INV_CUST_PO As String = fields(1)
                        '            If INV_NUM.Length < 10 Then
                        '                INV_NUM = INV_NUM.PadLeft(10, "0")
                        '            End If
                        '            Dim INV_PMT As Decimal = Val(fields(8))
                        '            Dim record_processed As Boolean = False

                        '            If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_NUM)
                        '            r += 1

                        '            If INV_PMT > 0 Or INV_PMT < 0 Then ' PAID ITEM - ADDED <0 BECAUSE THERE ARE NEGATIVE INVOICES

                        '                Dim Sql As String = "INV_NUM = '" & INV_NUM & "'"
                        '                'Sql &= " and INV_BALANCE = INV_TOTAL_AMOUNT_CURR"
                        '                Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)
                        '                If rows.Length = 0 Then
                        '                    Sql = "INV_CUST_PO = '" & INV_CUST_PO & "'"
                        '                    rows = dst.Tables("ARTPYMT3").Select(Sql)
                        '                End If

                        '                If rows.Length = 1 Then
                        '                    ' RECORD A PAYMENT TO THE INVOICE
                        '                    Dim rowARTPYMT3 As DataRow = rows(0)
                        '                    With rowARTPYMT3

                        '                        Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                        '                        Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                        '                        Dim INV_PMT_invoice As Decimal = INV_PMT
                        '                        Dim INV_DSC_invoice As Decimal = 0
                        '                        .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                        '                        .Item("INV_DISC_TAKEN") = 0
                        '                        .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice

                        '                        .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                        '                        .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice

                        '                        .Item("INV_DISC_TAKEN_CURR") = 0
                        '                        .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice

                        '                        .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)

                        '                        TOTAL_APPLIED += INV_PMT

                        '                        Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice
                        '                        'If System.Math.Abs(INV_BALANCE_NEW) < 0.05 Then
                        '                        '    INV_DSC_invoice += INV_BALANCE_NEW
                        '                        '    INV_BALANCE_NEW = 0
                        '                        'End If
                        '                        'If INV_DSC_invoice <> 0 Then
                        '                        '    Record_Chargeback(PYMT_BATCH_DLNO, "CM", -1 * INV_DSC_invoice, INV_NUM_cleaned)
                        '                        'End If

                        '                        'If INV_BALANCE_NEW <> 0 Then
                        '                        '    Record_Chargeback(PYMT_BATCH_DLNO, "CB", -1 * INV_BALANCE_NEW, INV_NUM_cleaned)
                        '                        'End If

                        '                    End With

                        '                    record_processed = True
                        '                Else
                        '                    Record_Chargeback(PYMT_BATCH_DLNO, "ONA", (INV_PMT), INV_NUM, , "Cannot Find AR Item")
                        '                End If

                        '            End If
                        '        End If
                        '    Next


                    ElseIf HFs("CUST_CODE") = "110451" Or HFs("CUST_CODE") = "311823" Then

                        Dim CHECK_NUM As String = ""
                        Dim r As Int64 = 0
                        Dim TOTAL_APPLIED As Decimal = 0
                        Dim skip As Integer = 0
                        Dim TOTAL_DISCOUNT As Decimal = 0

                        If FILENAME.ToUpper.EndsWith("XLS") Or FILENAME.ToUpper.EndsWith("XLSX") Then
                            ' change this for 5%
                            Dim CHECK_NUM_CHK As String
                            Dim TOTAL_APPLIED_CHK As String
                            Dim invamts As New Dictionary(Of String, Decimal)
                            '      Dim r As Integer = 0 ' NO HEADING
                            Do While oSheet.Cells(r, 0).Value & "" <> "Totals"

                                CHECK_NUM_CHK = oSheet.Cells(r, 0).Value
                                If Mid(CHECK_NUM_CHK, 1, 7) = "Check #" Then
                                    CHECK_NUM = "Chk#" & Mid(CHECK_NUM_CHK, 8)
                                End If

                                TOTAL_APPLIED_CHK = oSheet.Cells(r, 0).Value
                                If Mid(TOTAL_APPLIED_CHK, 1, 8) = "Amount $" Then
                                    TOTAL_APPLIED = Mid(TOTAL_APPLIED_CHK, 9)
                                    Record_Chargeback(PYMT_BATCH_DLNO, "ONA", -1 * TOTAL_APPLIED, CHECK_NUM, CHECK_NUM)
                                End If

                                If Mid(oSheet.Cells(r, 0).Value & "", 1, 7) = "REGENCY" Then
                                    ' PROCESS FILL invamts array with invoices
                                    Dim INV_NO As String = oSheet.Cells(r, 1).Value & ""
                                    Dim INV_LINE_AMT As Decimal = oSheet.Cells(r, 2).Value & ""
                                    If IsNumeric(INV_NO) Then
                                        If Not invamts.ContainsKey(INV_NO) Then
                                            invamts.Add(INV_NO, 0)
                                        End If
                                        invamts(INV_NO) += INV_LINE_AMT
                                    End If
                                End If


                                If r > 10 And oSheet.Cells(r, 0).Value & "" = "" Then
                                    Exit Do
                                End If
                                r = r + 1
                            Loop

                            ' Process INVOICESfrom invamts array
                            For Each INV_NO As String In invamts.Keys
                                Dim INV_AMT As Decimal = invamts(INV_NO)
                                Dim TOTAL_APPLIED1 As Decimal = 0
                                '  Dim skip As Integer = 0

                                Dim INV_NUM As String = INV_NO
                                If INV_NUM.Length < 10 Then
                                    INV_NUM = INV_NUM.PadLeft(10, "0")
                                End If
                                Dim INV_PMT As Decimal = INV_AMT
                                Dim record_processed As Boolean = False

                                Dim Sql As String = "INV_NUM = '" & INV_NUM & "'"
                                Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)

                                If rows.Length = 1 Then
                                    ' RECORD A PAYMENT TO THE INVOICE
                                    Dim rowARTPYMT3 As DataRow = rows(0)
                                    With rowARTPYMT3

                                        Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                                        Dim DIFF As Decimal = INV_BALANCE - INV_PMT
                                        If HFs("CUST_CODE") = "311823" Then
                                            If DIFF <> 0 And INV_PMT > 0 And INV_BALANCE > 0 Then
                                                Dim PCT As Decimal = 100 * DIFF / INV_BALANCE
                                                If System.Math.Abs(PCT) > 4.9 And System.Math.Abs(PCT) < 5.1 Then
                                                    TOTAL_DISCOUNT += INV_BALANCE - INV_PMT
                                                    INV_PMT = INV_BALANCE
                                                End If
                                            End If

                                        Else
                                            If DIFF <> 0 And INV_PMT > 0 And INV_BALANCE > 0 Then
                                                Dim PCT As Decimal = 100 * DIFF / INV_BALANCE
                                                If System.Math.Abs(PCT) > 2.9 And System.Math.Abs(PCT) < 3.1 Then
                                                    TOTAL_DISCOUNT += INV_BALANCE - INV_PMT
                                                    INV_PMT = INV_BALANCE
                                                End If
                                            End If

                                        End If

                                        Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                        Dim INV_PMT_invoice As Decimal = INV_PMT
                                        Dim INV_DSC_invoice As Decimal = 0
                                        .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                        .Item("INV_DISC_TAKEN") = 0
                                        .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice
                                        .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                                        .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice
                                        .Item("INV_DISC_TAKEN_CURR") = 0
                                        .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice
                                        .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)
                                        TOTAL_APPLIED += INV_PMT
                                        Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice
                                    End With
                                    record_processed = True
                                Else
                                    Record_Chargeback(PYMT_BATCH_DLNO, "ONA", (INV_PMT), INV_NUM, , "Cannot Find AR Item")
                                End If
                            Next
                            If TOTAL_DISCOUNT <> 0 Then
                                'Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "MISC", (-1 * TOTAL_DISCOUNT), "", "3% Disc", "Total Discounts Taken")
                                'row.Item("CHARGEBACK_IND") = "0"

                                Dim rowARTPYMT4 As DataRow = dst.Tables("ARTPYMT4").NewRow
                                rowARTPYMT4.Item("PYMT_BATCH_NO") = HFs("PYMT_BATCH_NO")
                                rowARTPYMT4.Item("PYMT_BATCH_LNO") = HFs("PYMT_BATCH_LNO")
                                rowARTPYMT4.Item("PYMT_BATCH_GLNO") = Val(dst.Tables("ARTPYMT4").Compute("Max(PYMT_BATCH_GLNO)", "") & "") + 1
                                Dim ACCT_CODE As String = "3065"
                                Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                                rowARTPYMT4.Item("ACCT_CODE") = rowGLTACCT1.Item("ACCT_CODE")
                                rowARTPYMT4.Item("ACCT_DESC") = rowGLTACCT1.Item("ACCT_DESC")
                                If HFs("CUST_CODE") = "311823" Then
                                    rowARTPYMT4.Item("GL_DIST_REF") = "5% Disc"
                                Else
                                    rowARTPYMT4.Item("GL_DIST_REF") = "3% Disc"
                                End If
                                For i As Integer = 2 To 4
                                    Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                                    rowARTPYMT4.Item(COLUMN_NAME) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                                Next
                                rowARTPYMT4.Item("GL_DIST_AMT") = TOTAL_DISCOUNT
                                rowARTPYMT4.Item("GL_DIST_AMT_CURR") = TOTAL_DISCOUNT
                                dst.Tables("ARTPYMT4").Rows.Add(rowARTPYMT4)
                                Display_Application_Totals()

                            End If

                        Else
                            Dim TXT As String = ""
                            Using SW As New System.IO.StreamReader(FILENAME)
                                TXT = SW.ReadToEnd()
                            End Using

                            CHECK_NUM = ""
                            r = 0
                            TOTAL_APPLIED = 0
                            skip = 0
                            TOTAL_DISCOUNT = 0

                            Dim lines() As String = Split(TXT, vbCrLf)
                            For Each line As String In lines

                                Dim fields() As String = Split(line, ",")

                                If skip > 0 Then
                                    skip -= 1

                                ElseIf line.Length > 20 AndAlso line.StartsWith("DEOR") Then

                                    Dim INV_NUM As String = ""
                                    Dim INV_PMT As Decimal = -1 * Val(Mid(line, 27, 12)) / 100
                                    ' Dim INV_DISC As Decimal = Val(Mid(line, 50, 7)) / 100
                                    Dim record_processed As Boolean = False

                                    If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_NUM)
                                    r += 1

                                    If INV_PMT > 0 Or INV_PMT < 0 Then ' PAID ITEM - ADDED <0 BECAUSE THERE ARE NEGATIVE INVOICES

                                        Dim Sql As String = "INV_BALANCE = " & CStr(INV_PMT)

                                        Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)

                                        If rows.Length = 1 Then
                                            ' RECORD A PAYMENT TO THE INVOICE
                                            Dim rowARTPYMT3 As DataRow = rows(0)
                                            With rowARTPYMT3

                                                Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                                                Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                                Dim INV_PMT_invoice As Decimal = INV_PMT
                                                Dim INV_DSC_invoice As Decimal = 0
                                                .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                                .Item("INV_DISC_TAKEN") = 0
                                                .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice
                                                .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                                                .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice
                                                .Item("INV_DISC_TAKEN_CURR") = 0
                                                .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice
                                                .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)
                                                TOTAL_APPLIED += INV_PMT
                                                Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice
                                            End With

                                            record_processed = True
                                        Else
                                            Record_Chargeback(PYMT_BATCH_DLNO, "ONA", (INV_PMT), INV_NUM, , "Cannot Find OA Item")
                                        End If

                                    End If

                                ElseIf line.Length > 20 AndAlso line.StartsWith("RM IV") Then

                                    Dim INV_NUM As String = Mid(line, 6, 10)
                                    Dim INV_PMT As Decimal = Val(Mid(line, 46, 10)) / 100
                                    Dim INV_DISC As Decimal = Val(Mid(line, 56, 10)) / 100
                                    Dim record_processed As Boolean = False

                                    If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_NUM)
                                    r += 1

                                    If INV_PMT > 0 Or INV_PMT < 0 Then ' PAID ITEM - ADDED <0 BECAUSE THERE ARE NEGATIVE INVOICES

                                        Dim Sql As String = "INV_NUM = '" & INV_NUM & "'"

                                        Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)

                                        If rows.Length = 1 Then
                                            ' RECORD A PAYMENT TO THE INVOICE
                                            Dim rowARTPYMT3 As DataRow = rows(0)
                                            With rowARTPYMT3

                                                Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")

                                                Dim DIFF As Decimal = INV_BALANCE - INV_PMT
                                                If DIFF <> 0 And INV_PMT > 0 And INV_BALANCE > 0 Then
                                                    Dim PCT As Decimal = 100 * DIFF / INV_BALANCE
                                                    If System.Math.Abs(PCT) > 2.9 And System.Math.Abs(PCT) < 3.1 Then
                                                        TOTAL_DISCOUNT += INV_BALANCE - INV_PMT
                                                        INV_PMT = INV_BALANCE
                                                    End If

                                                End If

                                                Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                                Dim INV_PMT_invoice As Decimal = INV_PMT
                                                Dim INV_DSC_invoice As Decimal = 0
                                                .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                                .Item("INV_DISC_TAKEN") = 0
                                                .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice
                                                .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                                                .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice
                                                .Item("INV_DISC_TAKEN_CURR") = 0
                                                .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice
                                                .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)
                                                TOTAL_APPLIED += INV_PMT
                                                Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice
                                            End With

                                            record_processed = True
                                        Else
                                            Record_Chargeback(PYMT_BATCH_DLNO, "ONA", (INV_PMT), INV_NUM, , "Cannot Find AR Item")
                                        End If

                                    End If
                                End If
                            Next

                            If TOTAL_DISCOUNT <> 0 Then
                                'Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "MISC", (-1 * TOTAL_DISCOUNT), "", "3% Disc", "Total Discounts Taken")
                                'row.Item("CHARGEBACK_IND") = "0"

                                Dim rowARTPYMT4 As DataRow = dst.Tables("ARTPYMT4").NewRow
                                rowARTPYMT4.Item("PYMT_BATCH_NO") = HFs("PYMT_BATCH_NO")
                                rowARTPYMT4.Item("PYMT_BATCH_LNO") = HFs("PYMT_BATCH_LNO")
                                rowARTPYMT4.Item("PYMT_BATCH_GLNO") = Val(dst.Tables("ARTPYMT4").Compute("Max(PYMT_BATCH_GLNO)", "") & "") + 1
                                Dim ACCT_CODE As String = "3051"
                                Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                                rowARTPYMT4.Item("ACCT_CODE") = rowGLTACCT1.Item("ACCT_CODE")
                                rowARTPYMT4.Item("ACCT_DESC") = rowGLTACCT1.Item("ACCT_DESC")
                                rowARTPYMT4.Item("GL_DIST_REF") = "3% Disc"
                                For i As Integer = 2 To 4
                                    Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                                    rowARTPYMT4.Item(COLUMN_NAME) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                                Next
                                rowARTPYMT4.Item("GL_DIST_AMT") = TOTAL_DISCOUNT
                                rowARTPYMT4.Item("GL_DIST_AMT_CURR") = TOTAL_DISCOUNT
                                dst.Tables("ARTPYMT4").Rows.Add(rowARTPYMT4)
                                Display_Application_Totals()

                            End If

                        End If


                    ElseIf HFs("CUST_CODE") = "308423" Then

                        Dim Invoice_Date As Boolean = False
                        Dim Returns_Date As Boolean = False
                        Dim Adj_date As Boolean = False
                        Dim CHECK_NUM As String = ""
                        Dim CHECK_AMT As Decimal
                        Dim CHECK_TOT As String
                        Dim CHECK_NUM_CHK As String
                        Dim invamts As New Dictionary(Of String, Decimal)
                        Dim r As Integer = 0 ' NO HEADING

                        CHECK_NUM_CHK = oSheet.Cells(2, 0).Value

                        Dim TOT_LINE As Int64 = 0
                        For x As Int64 = 1 To 50
                            If Not IsNothing(oWB.Worksheets(0).Cells(x, 0).Value) Then
                                If (oWB.Worksheets(0).Cells(x, 0).Value).ToString.Trim.ToUpper.StartsWith("CHECK TOTAL") Then
                                    TOT_LINE = x
                                    Exit For
                                End If
                            End If
                        Next

                        If Mid(CHECK_NUM_CHK, 1, 12) = "Check Number" Then
                            CHECK_NUM = oSheet.Cells(2, 0).Value & ""
                            CHECK_NUM = "Chk#" & Mid(CHECK_NUM, 16)
                        End If
                        CHECK_TOT = oSheet.Cells(TOT_LINE, 3).Value

                        If CHECK_TOT > 0 Then 'removed Net from before check total
                            CHECK_AMT = Val(oSheet.Cells(TOT_LINE, 3).Value & "") 'from r,9
                            Record_Chargeback(PYMT_BATCH_DLNO, "ONA", -1 * CHECK_AMT, CHECK_NUM, CHECK_NUM)
                        End If

                        Do While oWB.Worksheets(1).Cells(r, 0).Value & "" <> "" '"Totals"  'EDITED
                            'CHECK_NUM_CHK = oSheet.Cells(2, 0).Value

                            'If Mid(CHECK_NUM_CHK, 1, 12) = "Check Number" Then
                            '    CHECK_NUM = oSheet.Cells(2, 0).Value & ""
                            '    CHECK_NUM = "Chk#" & Mid(CHECK_NUM, 16)
                            'End If
                            'CHECK_TOT = oSheet.Cells(21, 3).Value

                            'If CHECK_TOT > 0 Then 'removed Net from before check total
                            '    CHECK_AMT = Val(oSheet.Cells(21, 3).Value & "") 'from r,9
                            '    Record_Chargeback(PYMT_BATCH_DLNO, "ONA", -1 * CHECK_AMT, CHECK_NUM, CHECK_NUM)
                            'End If

                            r = r + 1  'same line now

                            If oWB.Worksheets(1).Cells(r, 0).Value & "" = "Sales" Then '"Invoice Date" Then  
                                Invoice_Date = True
                                Returns_Date = False
                                Adj_date = False
                            ElseIf oWB.Worksheets(1).Cells(r, 0).Value & "" = "Returns" Then '"Returns Date" Then
                                Returns_Date = True
                                Invoice_Date = False
                                Adj_date = False
                            ElseIf oWB.Worksheets(1).Cells(r, 0).Value & "" = "Adjustments" Then '"Adjustment Date" Then
                                Adj_date = True
                                Invoice_Date = False
                                Returns_Date = False
                            ElseIf oWB.Worksheets(1).Cells(r, 0).Value & "" = "Supplier Oasis Fees" Then '"Supplier Oasis Fees" Then
                                Adj_date = True
                                Invoice_Date = False
                                Returns_Date = False
                            End If

                            'r = r + 1  'same line now

                            If Invoice_Date = True Then
                                ' PROCESS FILL invamts array with invoices
                                If ASCMAIN1.Running_in_VS Then
                                    If r >= 218 Then Stop
                                End If

                                If oWB.Worksheets(1).Cells(r, 1).Value & "" <> "" Then ' "Total" Then   'changed r,0 to r,1
                                    Dim INV_NO As String = oWB.Worksheets(1).Cells(r, 6).Value & ""  'r,4
                                    Dim INV_LINE_AMT As Decimal = oWB.Worksheets(1).Cells(r, 11).Value & ""  'r,9
                                    If Not invamts.ContainsKey(INV_NO) Then
                                        invamts.Add(INV_NO, 0)
                                    End If
                                    invamts(INV_NO) += INV_LINE_AMT
                                Else
                                    Invoice_Date = False
                                End If
                            ElseIf Returns_Date = True Then
                                If oWB.Worksheets(1).Cells(r, 1).Value & "" <> "" Then '"Total" Then 'changed r,0 to r,1
                                    Dim INV_LINE_AMT As Decimal = oWB.Worksheets(1).Cells(r, 11).Value & ""  'r,9
                                    Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW1", 1 * INV_LINE_AMT, INV_LINE_AMT & " Allowance ")
                                    row.Item("CHARGEBACK_IND") = "0"
                                Else
                                    Returns_Date = False
                                End If
                            ElseIf Adj_date = True Then
                                If oWB.Worksheets(1).Cells(r, 0).Value & "" <> "" Then '"Total" Then 
                                    Dim INV_LINE_AMT As Decimal = oWB.Worksheets(1).Cells(r, 11).Value & ""   'r,9
                                    ' Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW2", 1 * Val(INV_LINE_AMT), INV_LINE_AMT)
                                    Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW2", 1 * Val(INV_LINE_AMT), INV_LINE_AMT & " Allowance ")
                                    row.Item("CHARGEBACK_IND") = "0"
                                Else
                                    Adj_date = False
                                End If
                            End If

                            If r = 6000 Then
                                Exit Do
                            End If
                        Loop

                        ' Process INVOICESfrom invamts array
                        For Each INV_NO As String In invamts.Keys
                            Dim INV_AMT As Decimal = invamts(INV_NO)
                            Dim TOTAL_APPLIED As Decimal = 0
                            Dim skip As Integer = 0
                            Dim INV_NUM As String = INV_NO
                            If INV_NUM.Length < 10 Then
                                INV_NUM = INV_NUM.PadLeft(10, "0")
                            End If

                            Dim INV_PMT As Decimal = INV_AMT
                            Dim record_processed As Boolean = False

                            If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_NUM)
                            r += 1
                            If INV_PMT > 0 Or INV_PMT < 0 Then ' PAID ITEM - ADDED <0 BECAUSE THERE ARE NEGATIVE INVOICES
                                Dim Sql As String = "INV_NUM = '" & INV_NUM & "'"
                                Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)

                                If rows.Length = 1 Then
                                    ' RECORD A PAYMENT TO THE INVOICE
                                    Dim rowARTPYMT3 As DataRow = rows(0)
                                    With rowARTPYMT3

                                        Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                                        Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                        Dim INV_PMT_invoice As Decimal = INV_PMT
                                        Dim INV_DSC_invoice As Decimal = 0
                                        .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                        .Item("INV_DISC_TAKEN") = 0
                                        .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice

                                        .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                                        .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice

                                        .Item("INV_DISC_TAKEN_CURR") = 0
                                        .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice

                                        .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)

                                        TOTAL_APPLIED += INV_PMT

                                        Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice

                                    End With
                                    'With rowARTPYMT3
                                    '    Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                                    '    Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                    '    Dim INV_PMT_invoice As Decimal = INV_PMT
                                    '    Dim INV_DSC_invoice As Decimal = 0
                                    '    .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                    '    .Item("INV_DISC_TAKEN") = 0
                                    '    .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice
                                    '    .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                                    '    .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice
                                    '    .Item("INV_DISC_TAKEN_CURR") = 0
                                    '    .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice
                                    '    .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)
                                    '    TOTAL_APPLIED += INV_PMT
                                    '    Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice
                                    'End With
                                    record_processed = True
                                Else
                                    Record_Chargeback(PYMT_BATCH_DLNO, "ONA", (INV_PMT), INV_NUM, , "Cannot Find AR Itjem")
                                End If
                            End If
                        Next
                    ElseIf HFs("CUST_CODE") = "310921" Then

                        Dim Invoice_Date As Boolean = False
                        Dim Returns_Date As Boolean = False
                        Dim Adj_date As Boolean = False
                        Dim CHECK_NUM As String = ""
                        Dim CHECK_AMT As Decimal
                        Dim CHECK_TOT As String = ""
                        Dim CHECK_NUM_CHK As String



                        Dim invamts As New Dictionary(Of String, Decimal)
                        Dim r As Integer = 0 ' NO HEADING
                        Do While oSheet.Cells(r, 0).Value & "" <> "Totals"

                            CHECK_NUM_CHK = oSheet.Cells(r, 8).Value
                            If Mid(CHECK_NUM_CHK, 1, 5) = "EDIAP" Then
                                CHECK_NUM = oSheet.Cells(r, 4).Value & ""
                                CHECK_AMT = Val(oSheet.Cells(r, 6).Value & "")
                                Record_Chargeback(PYMT_BATCH_DLNO, "ONA", CHECK_AMT, CHECK_NUM, CHECK_NUM)
                                ' CHECK_NUM = "Chk#" & Mid(CHECK_NUM, 16)
                            End If


                            'CHECK_TOT = oSheet.Cells(r, 7).Value
                            'If Mid(CHECK_TOT, 1, 15) = "Net Check Total" Then
                            '    CHECK_AMT = Val(oSheet.Cells(r, 9).Value & "")
                            '    Record_Chargeback(PYMT_BATCH_DLNO, "ONA", -1 * CHECK_AMT, CHECK_NUM, CHECK_NUM)
                            'End If


                            If oSheet.Cells(r, 0).Value & "" = "Voucher" Then
                                Invoice_Date = True
                                Returns_Date = False
                                Adj_date = False
                                'ElseIf oSheet.Cells(r, 0).Value & "" = "Returns Date" Then
                                '    Returns_Date = True
                                '    Invoice_Date = False
                                '    Adj_date = False

                                'ElseIf oSheet.Cells(r, 0).Value & "" = "Adjustment Date" Then
                                '    Adj_date = True
                                '    Invoice_Date = False
                                '    Returns_Date = False

                            End If

                            r = r + 1
                            If Invoice_Date = True Then
                                ' PROCESS FILL invamts array with invoices
                                If oSheet.Cells(r, 0).Value & "" <> "" Then
                                    Dim INV_NO As String = oSheet.Cells(r, 7).Value & ""
                                    Dim INV_LINE_AMT As Decimal = oSheet.Cells(r, 17).Value & ""

                                    If Not invamts.ContainsKey(INV_NO) Then
                                        invamts.Add(INV_NO, 0)
                                    End If
                                    invamts(INV_NO) += INV_LINE_AMT
                                Else
                                    '  Invoice_Date = False

                                End If

                                'ElseIf Returns_Date = True Then
                                '    If oSheet.Cells(r, 0).Value & "" <> "Total" Then
                                '        Dim INV_LINE_AMT As Decimal = oSheet.Cells(r, 9).Value & ""
                                '        Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW1", 1 * INV_LINE_AMT, INV_LINE_AMT & " Allowance ")
                                '        row.Item("CHARGEBACK_IND") = "0"
                                '    Else
                                '        Returns_Date = False
                                '    End If

                                'ElseIf Adj_date = True And oSheet.Cells(r, 2).Value = "Supplier Oasis Transaction Fees" Then
                                '    If oSheet.Cells(r, 0).Value & "" <> "Total" Then
                                '        Dim INV_LINE_AMT As Decimal = oSheet.Cells(r, 9).Value & ""
                                '        ' Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW2", 1 * Val(INV_LINE_AMT), INV_LINE_AMT)
                                '        Dim row As DataRow = Record_Chargeback(PYMT_BATCH_DLNO, "ALLOW2", 1 * Val(INV_LINE_AMT), INV_LINE_AMT & " Allowance ")

                                '        row.Item("CHARGEBACK_IND") = "0"
                                '    Else
                                '        Adj_date = False
                                '    End If
                            End If
                            If r = 3000 Then
                                Exit Do
                            End If
                        Loop
                        ' Process INVOICESfrom invamts array

                        For Each INV_NO As String In invamts.Keys
                            Dim INV_AMT As Decimal = invamts(INV_NO)
                            Dim TOTAL_APPLIED As Decimal = 0
                            Dim skip As Integer = 0

                            Dim INV_NUM As String = INV_NO
                            If INV_NUM.Length < 10 Then
                                INV_NUM = INV_NUM.PadLeft(10, "0")
                            End If
                            Dim INV_PMT As Decimal = INV_AMT
                            Dim record_processed As Boolean = False

                            If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_NUM)
                            r += 1

                            If INV_PMT > 0 Or INV_PMT < 0 Then ' PAID ITEM - ADDED <0 BECAUSE THERE ARE NEGATIVE INVOICES
                                Dim Sql As String = "INV_CUST_PO  = '" & INV_NUM & "'"
                                Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(Sql)

                                If rows.Length = 1 Then
                                    ' RECORD A PAYMENT TO THE INVOICE
                                    Dim rowARTPYMT3 As DataRow = rows(0)
                                    With rowARTPYMT3

                                        Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")
                                        Dim INV_TOTAL_AMT As Decimal = Val(.Item("INV_TOTAL_AMOUNT_CURR") & "")
                                        Dim INV_PMT_invoice As Decimal = INV_PMT
                                        Dim INV_DSC_invoice As Decimal = 0
                                        .Item("INV_PMT") = INV_PMT ' INV_BALANCE ' Val(.Item("INV_PMT") & "") + INV_PMT_invoice
                                        .Item("INV_DISC_TAKEN") = 0
                                        .Item("INV_WRITE_OFF") = 0 ' Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice

                                        .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT '  (INV_PMT_invoice + INV_DSC_invoice)
                                        .Item("INV_PMT_CURR") = .Item("INV_PMT")  ' INV_BALANCE ' Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice ' - INV_DSC_invoice

                                        .Item("INV_DISC_TAKEN_CURR") = 0
                                        .Item("INV_WRITE_OFF_CURR") = 0 ' Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice

                                        .Item("INV_BALANCE_NEW_CURR") = .Item("INV_BALANCE_NEW") ' 0 ' Val(.Item("INV_BALANCE_NEW_CURR") & "") - (INV_PMT_invoice + INV_DSC_invoice)

                                        TOTAL_APPLIED += INV_PMT

                                        Dim INV_BALANCE_NEW As Decimal = INV_BALANCE - INV_PMT - INV_DSC_invoice

                                    End With
                                    record_processed = True
                                Else
                                    Record_Chargeback(PYMT_BATCH_DLNO, "ONA", (INV_PMT), INV_NUM, , "Cannot Find AR Itjem")
                                End If

                            End If
                        Next

                        '      Stop

                    Else

                        '????
                    End If
                Else

                    ' INT

                    dst.Tables("ARTPYMTL").Rows.Clear()

                    Dim r As Integer = 1 ' SKIP 1ST ROW - IT IS A HEADING
                    Do While oSheet.Cells(r, 0).Value & "" <> ""
                        Dim INV_NUM As String = Trim(oSheet.Cells(r, 0).Value & "")
                        Dim INV_REF As String = Trim(oSheet.Cells(r, 1).Value & "")
                        Dim INV_PMT As Decimal = Val(oSheet.Cells(r, 2).Value & "")
                        Dim REASON_CODE As String = Trim(oSheet.Cells(r, 3).Value & "").ToUpper
                        Dim EXP As String = Trim(oSheet.Cells(r, 4).Value & "")
                        Dim INV_DSC As Decimal = Val(oSheet.Cells(r, 5).Value & "")

                        If r Mod 100 = 0 Then ASCMAIN1.Progress("-", INV_REF)

                        If REASON_CODE <> "" Then
                            rowARTREAS1 = dst.Tables("ARTREAS1").Rows.Find(REASON_CODE)
                            If rowARTREAS1 Is Nothing Then
                                If Not BAD_REASON_CODEs.Contains(REASON_CODE) Then
                                    BAD_REASON_CODEs.Add(REASON_CODE)
                                End If
                            End If
                        End If

                        dst.Tables("ARTPYMTL").Rows.Add(New Object() {INV_NUM, INV_REF, INV_PMT})
                        r += 1

                        Dim sql As String = ""

                        If INV_NUM <> "" Then
                            Dim INV_NUM_cleaned As String = ""
                            For i As Integer = 1 To INV_NUM.Length
                                Dim C As String = Mid(INV_NUM, i, 1)
                                If C >= "0" And C <= "9" Then
                                    INV_NUM_cleaned = Mid(INV_NUM, i)
                                    Exit For
                                End If
                            Next

                            Dim record_processed As Boolean = False

                            If INV_NUM_cleaned.Length <= 10 Then
                                sql = "INV_NUM = '" & INV_NUM_cleaned.PadLeft(10, "0") & "'"
                                Dim rows() As DataRow = dst.Tables("ARTPYMT3").Select(sql)
                                If rows.Length = 1 AndAlso (REASON_CODE = "" Or REASON_CODE = rows(0).Item("REASON_CODE") & "") Then
                                    ' RECORD A PAYMENT TO THE INVOICE
                                    Dim rowARTPYMT3 As DataRow = rows(0)
                                    With rowARTPYMT3

                                        Dim INV_BALANCE As Decimal = Val(.Item("INV_BALANCE") & "")

                                        Dim INV_PMT_invoice As Decimal = INV_PMT
                                        Dim INV_DSC_invoice As Decimal = INV_DSC

                                        .Item("INV_PMT") = Val(.Item("INV_PMT") & "") + INV_PMT_invoice - INV_DSC_invoice ' * CURR_EXCH_RATE
                                        .Item("INV_DISC_TAKEN") = Val(.Item("INV_DISC_TAKEN") & "") + INV_DSC_invoice ' * CURR_EXCH_RATE
                                        .Item("INV_WRITE_OFF") = 0
                                        .Item("INV_BALANCE_NEW") = Val(.Item("INV_BALANCE_NEW") & "") - INV_PMT_invoice  ' * CURR_EXCH_RATE
                                        .Item("INV_PMT_CURR") = Val(.Item("INV_PMT_CURR") & "") + INV_PMT_invoice - INV_DSC_invoice
                                        .Item("INV_DISC_TAKEN_CURR") = Val(.Item("INV_DISC_TAKEN_CURR") & "") + INV_DSC_invoice
                                        .Item("INV_WRITE_OFF_CURR") = 0
                                        .Item("INV_BALANCE_NEW_CURR") = Val(.Item("INV_BALANCE_NEW_CURR") & "") - INV_PMT_invoice

                                    End With

                                    record_processed = True
                                End If
                            End If

                            If Not record_processed Then
                                ' RECORD A CHARGEBACK
                                Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
                                With rowARTPYMT5
                                    .Item("PYMT_BATCH_NO") = HFs("PYMT_BATCH_NO")
                                    .Item("PYMT_BATCH_LNO") = HFs("PYMT_BATCH_LNO")
                                    PYMT_BATCH_DLNO += 1
                                    .Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO
                                    If REASON_CODE <> "" Then
                                        .Item("REASON_CODE") = REASON_CODE
                                        If rowARTREAS1 IsNot Nothing Then
                                            .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
                                        End If
                                    Else
                                        .Item("REASON_CODE") = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_DEFAULT") & ""
                                    End If
                                    .Item("GL_DIST_AMT") = -1 * (INV_PMT - INV_DSC)  '  INV_PMT
                                    .Item("GL_DIST_COMMENT") = DBNull.Value

                                    If EXP.ToUpper = "X" Then
                                        If rowARTREAS1 IsNot Nothing Then ' BAD REASON CODES WILL SHOW UP AT THE END BECAUSE THEY ARE LOADED INTO BAD_REASON_CODEs
                                            .Item("ACCT_CODE") = rowARTREAS1.Item("ACCT_CODE")
                                            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                                            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                                            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                                        End If

                                    Else
                                        .Item("CHARGEBACK_IND") = "1"
                                        Dim INV_TYPE_CB As String = "B"
                                        If INV_PMT > 0 Then INV_TYPE_CB = "O"

                                        Dim ORDR_TYPE_CODE As String
                                        Dim rowSOTTYPE1 As DataRow
                                        If INV_TYPE_CB = "O" Then
                                            ORDR_TYPE_CODE = ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_OA")
                                        Else
                                            ORDR_TYPE_CODE = ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_CB")
                                        End If
                                        rowSOTTYPE1 = dst.Tables("SOTTYPE1").Rows.Find(ORDR_TYPE_CODE)
                                        Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE")
                                        Dim rowARTPOST1 As DataRow = dst.Tables("ARTPOST1").Rows.Find(POST_CODE)
                                        ' why generating the CHARGEBACK_NO here if we are going to do it all over again after clicking update?
                                        ' and why isn't this routine using Record_Chargeback?
                                        .Item("CHARGEBACK_NO") = ASCMAIN1.Next_Control_No("INV_NUM_" & INV_TYPE_CB)
                                        .Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
                                        .Item("SEG2_CODE") = rowARTPOST1.Item("SEG2_CODE")
                                        .Item("SEG3_CODE") = rowARTPOST1.Item("SEG3_CODE")
                                        .Item("SEG4_CODE") = rowARTPOST1.Item("SEG4_CODE")
                                        .Item("INV_TYPE_CB") = INV_TYPE_CB
                                        .Item("OUR_REFERENCE") = INV_NUM ' .Item("CUST_REFERENCE")
                                    End If

                                    Dim REF As String = IIf(INV_REF = "", INV_NUM, INV_REF)
                                    If CUST_CODE = "ULTA" And REF.StartsWith("V0000") Then
                                        REF = Mid(REF, 6)
                                    End If

                                    .Item("CUST_REFERENCE") = Mid(Trim(REF), 1, 20)
                                    .Item("CUST_CODE_SO") = DBNull.Value
                                    .Item("OUR_REFERENCE") = .Item("CUST_REFERENCE")
                                    .Item("GL_DIST_AMT_CURR") = -1 * (INV_PMT - INV_DSC)  '  INV_PMT
                                End With

                                dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)
                            End If
                        End If
                    Loop
                End If

                Calculate_Application_by_Type()
                Display_Application_Totals()

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

                If BAD_REASON_CODEs.Count = 0 Then
                    MsgBox("Application Worksheet has been Loaded", MsgBoxStyle.OkOnly, "Success")
                Else
                    MsgBox("Application Worksheet has been Loaded, " _
                           & vbCrLf & " There was bad data in the Reason Code Column (D)" _
                           & vbCrLf & " " & Join(BAD_REASON_CODEs.ToArray, ","), MsgBoxStyle.OkOnly, "Please Review Values loaded into Spreadsheet")
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Processing File")

        End Try

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Record_Chargeback( _
        ByRef PYMT_BATCH_DLNO As Int32, _
        REASON_CODE As String, _
        INV_PMT As Decimal, _
        CUST_REFERENCE As String, _
        Optional OUR_REFERENCE As String = "", _
        Optional GL_DIST_COMMENT As String = "", _
        Optional CUST_CODE_SO As String = "") As DataRow

        Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
        With rowARTPYMT5
            .Item("PYMT_BATCH_NO") = HFs("PYMT_BATCH_NO")
            .Item("PYMT_BATCH_LNO") = HFs("PYMT_BATCH_LNO")
            PYMT_BATCH_DLNO += 1
            .Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO

            If REASON_CODE <> "" Then
                .Item("REASON_CODE") = REASON_CODE
                rowARTREAS1 = LookUp("ARTREAS1", REASON_CODE)
                If rowARTREAS1 IsNot Nothing Then
                    .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")
                End If
            Else
                .Item("REASON_CODE") = ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_DEFAULT") & ""
            End If
            .Item("GL_DIST_AMT") = -1 * INV_PMT
            .Item("GL_DIST_COMMENT") = GL_DIST_COMMENT

            .Item("CHARGEBACK_IND") = "1"
            Dim INV_TYPE_CB As String = "B"
            If INV_PMT > 0 Then INV_TYPE_CB = "O"

            ' why generating the CHARGEBACK_NO here if we are going to do it all over again after clicking update?
            ' - also - this happens in XLS processing
            ' ok - remming out the code to get the CHARGEBACK_NO here
            .Item("CHARGEBACK_NO") = DBNull.Value ' ASCMAIN1.Next_Control_No("INV_NUM_" & INV_TYPE_CB)

            Dim ORDR_TYPE_CODE As String
            Dim rowSOTTYPE1 As DataRow
            If INV_TYPE_CB = "O" Then
                ORDR_TYPE_CODE = ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_OA")
            Else
                ORDR_TYPE_CODE = ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_CB")
            End If
            rowSOTTYPE1 = dst.Tables("SOTTYPE1").Rows.Find(ORDR_TYPE_CODE)
            Dim POST_CODE As String = rowSOTTYPE1.Item("POST_CODE")
            Dim rowARTPOST1 As DataRow = dst.Tables("ARTPOST1").Rows.Find(POST_CODE)

            .Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            .Item("INV_TYPE_CB") = INV_TYPE_CB
            .Item("OUR_REFERENCE") = OUR_REFERENCE ' "" ' INV_NUM ' .Item("CUST_REFERENCE")

            If CUST_REFERENCE.Length > dst.Tables("ARTPYMT5").Columns("CUST_REFERENCE").MaxLength Then
                CUST_REFERENCE = CUST_REFERENCE.Substring(0, dst.Tables("ARTPYMT5").Columns("CUST_REFERENCE").MaxLength).Trim
            End If
            .Item("CUST_REFERENCE") = CUST_REFERENCE
            .Item("CUST_CODE_SO") = CUST_CODE_SO

            .Item("GL_DIST_AMT_CURR") = -1 * INV_PMT
        End With

        dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)

        Return rowARTPYMT5
    End Function

    Private Sub optMatchChars_ValueChanged(sender As Object, e As EventArgs) Handles optMatchChars.ValueChanged
        numMatchChars.Visible = (optMatchChars.Value = "L")
    End Sub

    Private Sub tabARTPYMT3_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabARTPYMT3.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabARTPYMT3()
    End Sub
    Sub Setup_tabARTPYMT3()
        UltraExplorerBar1.Groups("Match").Visible = (tabARTPYMT3.SelectedTab.Key = "Match")
        UltraExplorerBar1.Groups("Control Totals").Visible = Not (tabARTPYMT3.SelectedTab.Key = "Match")
    End Sub

    Private Sub cmdMatch_Click(sender As Object, e As EventArgs) Handles cmdMatch.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Scanning for Matches")

        dst.Tables("ARTPYMTN").Rows.Clear()
        dst.Tables("ARTPYMTM").Rows.Clear()

        Dim REASON_NO_MATCHs As New List(Of String)
        For Each row As DataRow In dst.Tables("ARTREAS1").Select("REASON_NO_MATCH = '1'")
            Dim REASON_CODE As String = row.Item("REASON_CODE")
            REASON_NO_MATCHs.Add(REASON_CODE)
        Next
        'ASCMAIN1.sql = "Select * from ARTREAS1"
        'For Each row As DataRow In ASCDATA1.GetDataTable.Select("REASON_NO_MATCH = '1'")
        '    Dim REASON_CODE As String = row.Item("REASON_CODE")
        '    REASON_NO_MATCHs.Add(REASON_CODE)
        'Next

        Dim SQL As String = "INV_BALANCE <> 0 and INV_CUST_PO is Not Null and INV_TYPE <> 'I'"
        If chkMatchNoPartials.Checked Then SQL &= " AND INV_TOTAL_AMT = INV_BALANCE"
        For Each rowARTPYMT3 As DataRow In dst.Tables("ARTPYMT3").Select(SQL)
            Dim MATCH_REF As String = rowARTPYMT3.Item("INV_CUST_PO") & ""
            ' If ASCMAIN1.Running_in_VS And MATCH_REF.StartsWith("7203") Then Stop
            If MATCH_REF <> "" Then
                Do While MATCH_REF.StartsWith("0")
                    MATCH_REF = Mid(MATCH_REF, 2)
                Loop
            End If
            If txtIGNORE.Text <> "" Then
                If txtIGNORE.Text = "*" Then
                    Dim MATCH_REF_orig As String = MATCH_REF
                    For i As Integer = MATCH_REF.Length To 1 Step -1
                        If Mid(MATCH_REF, MATCH_REF.Length, 1).ToUpper >= "A" And Mid(MATCH_REF, MATCH_REF.Length, 1).ToUpper <= "Z" Then
                            MATCH_REF = Mid(MATCH_REF, 1, Len(MATCH_REF) - 1)
                        Else
                            Exit For
                        End If
                    Next
                    If MATCH_REF = "" Then MATCH_REF = MATCH_REF_orig
                ElseIf MATCH_REF.EndsWith(txtIGNORE.Text) Then
                    MATCH_REF = Mid(MATCH_REF, 1, MATCH_REF.Length - txtIGNORE.Text.Length)
                End If
            End If
            If optMatchChars.Value = "L" Then
                If MATCH_REF.Length > numMatchChars.Value Then
                    MATCH_REF = Mid(MATCH_REF, MATCH_REF.Length - numMatchChars.Value + 1, numMatchChars.Value)
                End If
            End If

            Dim MATCH_RC As String = rowARTPYMT3.Item("REASON_CODE") & ""
            Dim rowARTREAS1 As DataRow = dst.Tables("ARTREAS1").Rows.Find(MATCH_RC)
            Dim MATCH_RG As String = ""
            Dim REASON_MATCH_GROUP As String = ""
            If rowARTREAS1 IsNot Nothing Then
                REASON_MATCH_GROUP = rowARTREAS1.Item("REASON_MATCH_GROUP") & ""
                MATCH_RG = REASON_MATCH_GROUP
            End If
            If MATCH_RG = "" Then MATCH_RG = "*"

            If REASON_NO_MATCHs.Contains(MATCH_RC) Then
            Else
                If Not chkMatchByReason.Checked Then
                    MATCH_RC = "*"
                End If
                If Not chkMatchByGroup.Checked Then
                    MATCH_RG = "*"
                End If

                Dim rowARTPYMTM As DataRow = dst.Tables("ARTPYMTM").Rows.Find(New String() {MATCH_REF, MATCH_RC, MATCH_RG})
                If rowARTPYMTM Is Nothing Then
                    rowARTPYMTM = dst.Tables("ARTPYMTM").Rows.Add(New String() {MATCH_REF, MATCH_RC, MATCH_RG})
                End If
                rowARTPYMTM.Item("INV_COUNT") = Val(rowARTPYMTM.Item("INV_COUNT") & "") + 1
                Dim INV_BALANCE As Decimal = Val(rowARTPYMT3.Item("INV_BALANCE") & "")
                rowARTPYMTM.Item("MATCH_TOTAL") = Val(rowARTPYMTM.Item("MATCH_TOTAL") & "") + INV_BALANCE
                If INV_BALANCE > 0 Then
                    rowARTPYMTM.Item("MATCH_TOTAL_DR") = Val(rowARTPYMTM.Item("MATCH_TOTAL_DR") & "") + INV_BALANCE
                Else
                    rowARTPYMTM.Item("MATCH_TOTAL_CR") = Val(rowARTPYMTM.Item("MATCH_TOTAL_CR") & "") + INV_BALANCE
                End If
                Dim rowARTPYMTN As DataRow = dst.Tables("ARTPYMTN").NewRow
                With rowARTPYMTN
                    .Item("MATCH_REF") = MATCH_REF
                    .Item("MATCH_RC") = MATCH_RC
                    .Item("MATCH_RG") = MATCH_RG
                    .Item("PYMT_BATCH_ILNO") = rowARTPYMT3.Item("PYMT_BATCH_ILNO")
                    .Item("INV_TYPE") = rowARTPYMT3.Item("INV_TYPE")
                    .Item("INV_NUM") = rowARTPYMT3.Item("INV_NUM")
                    .Item("CUST_CODE") = HFs("CUST_CODE")
                    .Item("INV_CUST_PO") = rowARTPYMT3.Item("INV_CUST_PO")
                    .Item("REASON_CODE") = rowARTPYMT3.Item("REASON_CODE")
                    .Item("REASON_MATCH_GROUP") = REASON_MATCH_GROUP
                    .Item("INV_DATE") = rowARTPYMT3.Item("INV_DATE")
                    .Item("INV_BALANCE") = rowARTPYMT3.Item("INV_BALANCE")
                End With
                dst.Tables("ARTPYMTN").Rows.Add(rowARTPYMTN)
            End If
        Next

        For Each rowARTPYMTM As DataRow In dst.Tables("ARTPYMTM").Select("")
            Dim MATCH_TOTAL As Decimal = Val(rowARTPYMTM.Item("MATCH_TOTAL") & "")
            Dim MATCH_TOTAL_DR As Decimal = Val(rowARTPYMTM.Item("MATCH_TOTAL_DR") & "")
            Dim MATCH_TOTAL_CR As Decimal = Val(rowARTPYMTM.Item("MATCH_TOTAL_CR") & "")
            Dim INV_COUNT As Integer = Val(rowARTPYMTM.Item("INV_COUNT") & "")
            Dim MATCH_ACTION As String = "0"
            If INV_COUNT = 0 Then MATCH_ACTION = "2"
            If chkMatchTolerance.Checked And System.Math.Abs(MATCH_TOTAL) > numMatchTolerance.Value Then MATCH_ACTION = "2"
            If MATCH_TOTAL_CR = 0 Or MATCH_TOTAL_DR = 0 Then MATCH_ACTION = "2"
            rowARTPYMTM.Item("MATCH_ACTION") = MATCH_ACTION
        Next

        ASCDATA1.DeleteRows("ARTPYMTM", "MATCH_ACTION = '2'")
        Sort_grdColumns(grdARTPYMTM, "MATCH_REF,MATCH_RC")

        If dst.Tables("ARTPYMTM").Select("").Length = 0 Then
            MsgBox("No Matches Found using Criteria Specified", MsgBoxStyle.OkOnly, "No Matches")
            cmdMatchApply.Visible = False
        Else
            cmdMatchApply.Visible = True
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdMatchApply_Click(sender As Object, e As EventArgs) Handles cmdMatchApply.Click
        If dst.Tables("ARTPYMTM").Select("MATCH_ACTION = '1'").Length = 0 Then
            MsgBox("No Matches Selected to Apply", MsgBoxStyle.OkOnly, "Cannot Apply Matches")
            Exit Sub
        End If

        If MsgBox("The selected Matches will be keyed off," _
                  & vbCrLf & " and new DR or CR items will be created for each Match where" _
                  & vbCrLf & " the Net Match is non-zero." _
                  & vbCrLf & vbCrLf & "OK to proceed with Matching process?",
                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        Dim PYMT_BATCH_DLNO_ctr As Integer = Val(dst.Tables("ARTPYMT5").Compute("MAX(PYMT_BATCH_DLNO)", "") & "")

        For Each rowARTPYMTM As DataRow In dst.Tables("ARTPYMTM").Select("MATCH_ACTION = '1'")
            Dim MATCH_TOTAL As Decimal = Val(rowARTPYMTM.Item("MATCH_TOTAL") & "")
            Dim PYMT_BATCH_ILNO_clone As Integer = -1

            For Each rowARTPYMTN As DataRow In rowARTPYMTM.GetChildRows("ARTPYMTM_ARTPYMTN")
                Dim PYMT_BATCH_ILNO As Integer = Val(rowARTPYMTN.Item("PYMT_BATCH_ILNO"))
                Dim PMT As Decimal = Val(rowARTPYMTN.Item("INV_BALANCE") & "")
                If MATCH_TOTAL <> 0 And System.Math.Sign(MATCH_TOTAL) = System.Math.Sign(PMT) And PYMT_BATCH_ILNO_clone = -1 Then
                    PYMT_BATCH_ILNO_clone = PYMT_BATCH_ILNO
                End If

                For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT3.Rows
                    If grow.IsDataRow AndAlso Val(grow.Cells("PYMT_BATCH_ILNO").Value & "") = PYMT_BATCH_ILNO Then
                        grow.Cells("INV_PMT_CURR").Value = Val(grow.Cells("INV_PMT_CURR").Value) + PMT
                        grow.Update()
                        Exit For
                    End If
                Next
                'Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").Rows.Find _
                '                             (New Object() {PYMT_BATCH_NO, PYMT_BATCH_LNO, PYMT_BATCH_ILNO})
                'rowARTPYMT3.Item("") = Val("") + Pmt()
            Next

            If MATCH_TOTAL <> 0 Then

                Dim rowARTPYMT3 As DataRow = dst.Tables("ARTPYMT3").Rows.Find _
                                             (New Object() {HFs("PYMT_BATCH_NO"), HFs("PYMT_BATCH_LNO"), PYMT_BATCH_ILNO_clone})

                PYMT_BATCH_DLNO_ctr = PYMT_BATCH_DLNO_ctr + 1

                ' why not using Record_Chargeback?

                Dim rowARTPYMT5 As DataRow = dst.Tables("ARTPYMT5").NewRow
                With rowARTPYMT5
                    .Item("PYMT_BATCH_NO") = rowARTPYMT3.Item("PYMT_BATCH_NO")
                    .Item("PYMT_BATCH_LNO") = rowARTPYMT3.Item("PYMT_BATCH_LNO")
                    .Item("PYMT_BATCH_DLNO") = PYMT_BATCH_DLNO_ctr

                    .Item("REASON_CODE") = rowARTPYMT3.Item("REASON_CODE")

                    Dim rowARTREAS1 As DataRow = dst.Tables("ARTREAS1").Rows.Find(.Item("REASON_CODE"))
                    .Item("REASON_DESC") = rowARTREAS1.Item("REASON_DESC")

                    .Item("ACCT_CODE") = DBNull.Value
                    .Item("GL_DIST_AMT") = MATCH_TOTAL ' * CURR_EXCH_RATE
                    .Item("GL_DIST_COMMENT") = "Balance after Matching"
                    .Item("CHARGEBACK_IND") = "1"
                    .Item("CHARGEBACK_NO") = DBNull.Value

                    .Item("CUST_REFERENCE") = rowARTPYMT3.Item("INV_CUST_PO")
                    .Item("CUST_CODE_SO") = rowARTPYMT3.Item("CUST_CODE_SO")
                    .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    .Item("INV_TYPE_CB") = DBNull.Value
                    '.Item("OUR_REFERENCE") = rowARTPYMT3.Item("OUR_REFERENCE")
                    .Item("GL_DIST_AMT_CURR") = MATCH_TOTAL
                End With
                dst.Tables("ARTPYMT5").Rows.Add(rowARTPYMT5)
            End If
        Next

        Calculate_Application_by_Type()
        Display_Application_Totals()

        MsgBox("Match Application Complete - You must Still Update to finalize this Application", MsgBoxStyle.OkOnly, "Verification")

        dst.Tables("ARTPYMTN").Rows.Clear()
        dst.Tables("ARTPYMTM").Rows.Clear()
        tabARTPYMT3.SelectedTab = tabARTPYMT3.Tabs("Open AR Items")
    End Sub

    Private Sub cmdApplySel_Click(sender As Object, e As EventArgs) Handles cmdApplySel.Click

        If grdARTPYMT3.Selected.Rows.Count = 0 Then
            MsgBox("No Rows Selected to Apply/Pay", MsgBoxStyle.OkOnly, "Cannot Apply")
            Exit Sub
        End If

        Dim T As Decimal = 0
        For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT3.Selected.Rows
            Dim INV_BALANCE As Decimal = 0
            T += Val(grow.Cells("INV_BALANCE_CURR").Value & "")
        Next

        If MsgBox("OK to Click Apply for the " & grdARTPYMT3.Selected.Rows.Count _
                  & " Selected Items totaling " & Format(T, "#,##0.00") & CURR_CODE & " ?",
                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
            Exit Sub
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT3.Selected.Rows
            grow.Activate()
            Click_Pay()
        Next
        grdARTPYMT3.Selected.Rows.Clear()

    End Sub

    Sub Select_Range()
        Dim INV_NO1 As String = txtINV1.Text
        Dim INV_NO2 As String = txtINV2.Text

        INV_NO1 = INV_NO1.PadLeft(10, "0")
        INV_NO2 = INV_NO2.PadLeft(10, "0")

        txtINV1.Text = INV_NO1
        txtINV2.Text = INV_NO2

        If INV_NO2 < INV_NO1 Then
            MsgBox("Starting invoice is > than the Ending invoice", MsgBoxStyle.OkOnly, "Cannot Select Range")
            Exit Sub
        End If

        grdARTPYMT3.Selected.Rows.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT3.Rows
            Dim INV_TYPE As String = grow.Cells("INV_TYPE").Value
            Dim INV_NUM As String = grow.Cells("INV_NUM").Value
            If INV_TYPE = "I" Then
                If INV_NUM >= INV_NO1 And INV_NUM <= INV_NO2 Then
                    grow.Selected = True
                End If
            End If

        Next
    End Sub

    Private Sub txtINV1_KeyDown(sender As Object, e As KeyEventArgs) Handles txtINV1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select_Range()
        End If
    End Sub

    Private Sub txtINV2_KeyDown(sender As Object, e As KeyEventArgs) Handles txtINV2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select_Range()
        End If
    End Sub

    Private Sub btnSaveEDTXREF1_Click(sender As Object, e As EventArgs) Handles btnSaveEDTXREF1.Click
        Update_Record_TDA("EDTXREF1")
        MsgBox("EDI Reason Codes X-Ref has been saved")
    End Sub

    Private Sub grdEDTXREF1_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdEDTXREF1.ClickCellButton
        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "REASON_CODE"
                sql_where = ""


        End Select

        grdClickCellButton(grdEDTXREF1, sql_where, sql_where <> "")
    End Sub

    Private Sub grdEDTXREF1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdEDTXREF1.BeforeRowUpdate
        With grdEDTXREF1
            If e.Row.Cells("REASON_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("ARTREAS1", e.Row.Cells("REASON_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Reason Code (" & e.Row.Cells("REASON_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                End If
            End If
        End With
    End Sub

    Private Sub grdEDTXREF1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDTXREF1.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "REASON_CODE"
                Dim REASON_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdEDTXREF1, "ARTREAS1", "REASON_CODE", "REASON_DESC")
        End Select
    End Sub
End Class