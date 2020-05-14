Public Class APFVINQ1
    Dim ICTIREC1 As String
    Dim ICTIREC1_TOTALS As String
    Dim POTORDR1_TOTALS As String
    Dim AGE_DATES(,)
    Dim DAYS(4) As Integer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 12, -11) ' -11 - 12)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 12, 0) '  0 - 12)

        With dst

            ASCMAIN1.sql = "Select * from APTINVH1 where APTINVH1.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0, False, "V")

            ASCMAIN1.sql = "Select APTINVH2.*, GLTACCT1.ACCT_DESC " _
             & " from APTINVH2,GLTACCT1 " _
             & " where GLTACCT1.ACCT_CODE (+) = APTINVH2.ACCT_CODE" _
             & " and APTINVH2.VOUCHER_NO = :PARM1"
            Create_TDA(.Tables.Add, "APTINVH2", "**", 0, False, "V", 2)

            '.Relations.Add("APTINVH2", _
            'New DataColumn() {.Tables("APTINVH1").Columns("VOUCHER_NO")}, _
            'New DataColumn() {.Tables("APTINVH2").Columns("VOUCHER_NO")})


            ASCMAIN1.sql = "Select * from APTCHCK1 where VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTCHCK1", "**", 0, False, "V")

            ASCMAIN1.sql = "Select APTCHCK2.*" & vbCrLf _
                & ", APTINVH1.INV_TYPE, APTINVH1.INV_REF, APTINVH1.INV_DUE_DATE" & vbCrLf _
                & " from APTCHCK2,APTINVH1" & vbCrLf _
                & " where APTCHCK2.BANK_CODE = :PARM1 " & vbCrLf _
                & " and APTCHCK2.CHECK_NUM = :PARM2 " & vbCrLf _
                & " and APTINVH1.VOUCHER_NO (+) = APTCHCK2.VOUCHER_NO"
            Create_TDA(.Tables.Add, "APTCHCK2", "**", 0, False, "VV", 3)

            '.Relations.Add("APTCHCK2", _
            'New DataColumn() {.Tables("APTCHCK1").Columns("BANK_CODE"), _
            '                  .Tables("APTCHCK1").Columns("CHECK_NUM")}, _
            'New DataColumn() {.Tables("APTCHCK2").Columns("BANK_CODE"), _
            '                  .Tables("APTCHCK2").Columns("CHECK_NUM")})

            Dim T As DataTable
            T = .Tables("APTCHCK1").Clone
            T.TableName = "APTCHCX1"
            .Tables.Add(T)

            ASCMAIN1.sql = "Select APTCHCK2.* from APTCHCK2 " _
             & " where APTCHCK2.BANK_CODE = :PARM1 " _
             & " and APTCHCK2.CHECK_NUM = :PARM2 "
            Create_TDA(.Tables.Add, "APTCHCX2", "**", 0, False, "VV", 3)

            .Relations.Add("APTCHCX2", _
            New DataColumn() {.Tables("APTCHCX1").Columns("BANK_CODE"), _
                              .Tables("APTCHCX1").Columns("CHECK_NUM")}, _
            New DataColumn() {.Tables("APTCHCX2").Columns("BANK_CODE"), _
                              .Tables("APTCHCX2").Columns("CHECK_NUM")})

            Create_TDA(.Tables.Add, "APTVEND1", "*", , False)

            ASCMAIN1.sql = "Select APTVEND5.*" _
            & ", APTVEND1.VEND_NAME, APTVEND1.VEND_CLASS_CODE, APTVEND1.VEND_TYPE" _
            & ", APTVEND1.TERM_CODE, APTVEND1.POST_CODE, APTVEND1.BANK_CODE" _
            & ", APTVEND1.PROCESSOR_CODE, APTVEND1.VEND_CODE_AP" _
            & ", APTVEND1.VEND_ON_HOLD, APTVEND1.VEND_ON_HOLD_DATE" _
            & ", APTVEND1.VEND_PYMT_METHOD, APTVEND1.VEND_PYMT_CYCLE" _
            & ", APTVEND1.VEND_STOP_PURCHASE, APTVEND1.VEND_ALWAYS_TAKE_DISC" _
            & ", X.INV_BALANCE " _
            & " from APTVEND5,APTVEND1," _
            & " (SELECT VEND_CODE, SUM (INV_BALANCE) INV_BALANCE " _
            & "  from APTINVH1 where INV_STATUS in ('O','H') group by VEND_CODE) X " _
            & " where APTVEND1.VEND_CODE = APTVEND5.VEND_CODE " _
            & " and X.VEND_CODE (+) = APTVEND5.VEND_CODE" _
            & IIf(MENU_ITEM_PP = "S", " and APTVEND1.VEND_TYPE = 'S'", "")
            Create_TDA(.Tables.Add, "APTVEND5", "**", 0, True, , 1)

            Dim sql As String = ""
            'sql = "Select ICTIREC1.*, ICTIREC2.QTY_REC, ICTIREC2.QTY_REC * ICTIREC2.PO_COST AMT_REC from ICTIREC1,ICTIREC2 where ROWNUM < 1"
            'ICTIREC1 = ASCMAIN1.Temp_Table(sql)
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTIREC1 & " Add Primary Key (RECEIPT_NO)")
            'sql = "Select ICTIREC2.RECEIPT_NO, QTY_REC, QTY_REC * PO_COST AMT_REC from ICTIREC2 where ROWNUM < 1"
            'ICTIREC1_TOTALS = ASCMAIN1.Temp_Table(sql)
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTIREC1_TOTALS & " Add Primary Key (RECEIPT_NO)")


            'ASCMAIN1.sql = "Select ICTIREC1.* from " & ICTIREC1 & " ICTIREC1" '  & " where ICTIREC1.VEND_CODE = :PAM1 and X.RECEIPT_NO = ICTIREC1.RECEIPT_NO"
            'Create_TDA(.Tables.Add, "ICTIREC1", "**", 0, False, "", 1)

            'ASCMAIN1.sql = "Select ICTIREC2.*, ICTITEM1.ITEM_DESC " _
            ' & " from ICTIREC2,ICTITEM1 " _
            ' & " where ICTIREC2.ITEM_CODE = ICTITEM1.ITEM_CODE " _
            ' & "   and ICTIREC2.RECEIPT_NO = :PARM1"
            'Create_TDA(.Tables.Add, "ICTIREC2", "**", 0, False, "V", 2)
            '.Tables("ICTIREC2").Columns.Add("INV_COST", GetType(System.Double))

            '.Relations.Add("ICTIREC2", _
            'New DataColumn() {.Tables("ICTIREC1").Columns("RECEIPT_NO")}, _
            'New DataColumn() {.Tables("ICTIREC2").Columns("RECEIPT_NO")})


            'sql = "Select POTORDR2.PO_ORDER_NO " _
            '& ", POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST PO_AMT_ORD" _
            '& ", POTORDR2.PO_QTY_REC * POTORDR2.PO_COST PO_AMT_REC" _
            '& ", POTORDR2.PO_QTY_INV * POTORDR2.PO_COST PO_AMT_INV" _
            '& ", POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST PO_AMT_OPN " _
            '& ", POTORDR2.PO_QTY_CXL * POTORDR2.PO_COST PO_AMT_CXL " _
            '& " from POTORDR2 where ROWNUM < 1"
            'POTORDR1_TOTALS = ASCMAIN1.Temp_Table(sql)
            'ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1_TOTALS & " Add Primary Key (PO_ORDER_NO)")

            'ASCMAIN1.sql = "Select POTORDR1.* " _
            '& ", POTORDR1_TOTALS.PO_AMT_ORD" _
            '& ", POTORDR1_TOTALS.PO_AMT_REC" _
            '& ", POTORDR1_TOTALS.PO_AMT_INV" _
            '& ", POTORDR1_TOTALS.PO_AMT_OPN" _
            '& ", POTORDR1_TOTALS.PO_AMT_CXL" _
            '& " from POTORDR1," & POTORDR1_TOTALS & " POTORDR1_TOTALS" _
            '& " where POTORDR1.PO_ORDER_NO = POTORDR1_TOTALS.PO_ORDER_NO (+) " _
            '& " and POTORDR1.VEND_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "V")

            'ASCMAIN1.sql = "Select POTORDR2.* " _
            ' & " from POTORDR2 " _
            ' & " where POTORDR2.PO_ORDER_NO = :PARM1 "
            'Create_TDA(.Tables.Add, "POTORDR2", "**", 0, False, "V", 2)

            '.Relations.Add("POTORDR2", _
            'New DataColumn() {.Tables("POTORDR1").Columns("PO_ORDER_NO")}, _
            'New DataColumn() {.Tables("POTORDR2").Columns("PO_ORDER_NO")})

            ASCMAIN1.sql = "Select * from APTVEND2 where APTVEND2.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTVEND2", "**", 0, False, "V")

            .Tables.Add("APTVEND1_CODES")
            With .Tables("APTVEND1_CODES")
                .Columns.Add("CODE_TYPE", GetType(System.String))
                .Columns.Add("CODE_VALUE", GetType(System.String))
                .Columns.Add("DESC_VALUE", GetType(System.String))
            End With

            .Tables.Add("APTVEND1_STATS")
            With .Tables("APTVEND1_STATS")
                .Columns.Add("STAT_TYPE", GetType(System.String))
                .Columns.Add("STAT_MTD", GetType(System.Int32))
                .Columns.Add("STAT_YTD", GetType(System.Int32))
                .Columns.Add("STAT_LYR", GetType(System.Int32))
            End With

            .Tables.Add("APTVEND1_AGING")
            With .Tables("APTVEND1_AGING")
                .Columns.Add("AGE_CATGY", GetType(System.String))
                .Columns.Add("AGE_AMT", GetType(System.Decimal))
            End With

            .Tables.Add("APTVEND1_AGEDAP")
            With .Tables("APTVEND1_AGEDAP")
                .Columns.Add("AGE_DATE", GetType(System.DateTime))
                .Columns.Add("AGE_AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("AGE_DATE")}
            End With

            '.Tables.Add("APTVEND1_AGING")
            'With .Tables("APTVEND1_AGING")
            '    .Columns.Add("AGE_CATGY", GetType(System.String))
            '    .Columns.Add("AGE_AMT", GetType(System.Decimal))
            'End With

            .Tables.Add("APTVEND1_AP_TYPE")
            With .Tables("APTVEND1_AP_TYPE")
                .Columns.Add("AP_TYPE", GetType(System.String))
                .Columns.Add("AP_AMT", GetType(System.Decimal))
            End With

            .Tables.Add("APTVEND1_FL")
            With .Tables("APTVEND1_FL")
                .Columns.Add("EVENT_DESC", GetType(System.String))
                .Columns.Add("EVENT_ITEM", GetType(System.String))
                .Columns.Add("EVENT_DATE", GetType(System.DateTime))
                .Columns.Add("EVENT_AMT", GetType(System.Double))
            End With

            .Tables.Add("APTVEND1_OPEN")
            With .Tables("APTVEND1_OPEN")
                .Columns.Add("OPEN_CATGY", GetType(System.String))
                .Columns.Add("OPEN_AMT", GetType(System.Double))
            End With


            '   ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME" & vbCrLf _
            '       , X.INV_AMT, X.INV_CNT, X.INV_DATE from APTVEND1, (" & vbCrLf _
            '      & "Select VEND_CODE, SUM (INV_AMT) INV_AMT, COUNT (*) INV_CNT, MAX(INV_DATE) INV_DATE" & vbCrLf _
            '     & " from APTINVH1 where INV_STATUS in ('O','P')" & vbCrLf _
            '    & " and OPS_YYYYPP between :PARM1 and :PARM2" & vbCrLf _
            '   & " group by VEND_CODE) X" & vbCrLf _
            '  & " where APTVEND1.VEND_CODE = X.VEND_CODE"
            ASCMAIN1.sql = "Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME" & vbCrLf _
            & ", X.INV_AMT, X.INV_CNT, X.INV_DATE,X.CHECK_AMT, X.CHECK_CNT, X.CHECK_DATE from APTVEND1, (" & vbCrLf _
            & " Select VEND_CODE, SUM (INV_AMT) INV_AMT, SUM(INV_CNT) INV_CNT, MAX(INV_DATE) INV_DATE, SUM(CHECK_AMT) CHECK_AMT, SUM(CHECK_CNT) CHECK_CNT, MAX(CHECK_DATE) CHECK_DATE" & vbCrLf _
            & " from ( Select VEND_CODE, SUM (INV_AMT) INV_AMT, COUNT (*) INV_CNT, MAX(INV_DATE) INV_DATE, 0 CHECK_AMT, 0 CHECK_CNT, NULL CHECK_DATE" & vbCrLf _
            & " from APTINVH1 where INV_STATUS in ('O','P') and OPS_YYYYPP between :PARM1 and :PARM2" & vbCrLf _
            & " group by VEND_CODE" & vbCrLf _
            & " UNION " & vbCrLf _
            & " Select VEND_CODE, 0 INV_AMT, 0 INV_CNT, NULL INV_DATE, SUM (CHECK_AMT) CHECK_AMT, COUNT (*) CHECK_CNT,  MAX(CHECK_DATE) " & vbCrLf _
            & " from APTCHCK1 where CHECK_STATUS = 'I'" & vbCrLf _
            & " and OPS_YYYYPP between :PARM1 and :PARM2" & vbCrLf _
            & " group by VEND_CODE) GROUP BY VEND_CODE) X where APTVEND1.VEND_CODE = X.VEND_CODE "
            
            Create_TDA(.Tables.Add, "APTINVHH", "**", 0, False, "VV")
            dst.Tables("APTINVHH").Columns("INV_CNT").DataType = GetType(System.Int32)
            dst.Tables("APTINVHH").Columns("CHECK_CNT").DataType = GetType(System.Int32)

        End With

        Create_Lookup("APTVEND1")
        Create_Lookup("APTVEND5")
        Create_Lookup("TATTERM1")
        Create_Lookup("GLTBANK1")
        Create_Lookup("APTPOST1")
        Create_Lookup("ASTUSER1")
        Create_Lookup("APTCLAS1")
        Create_Lookup("GLTACCT1")
        Create_Lookup("APTVEND2")
        Create_Lookup("ASTCODE1")

        grdAPTVEND2.DataSource = dst.Tables("APTVEND2")
        grdAPTINVHH.DataSource = dst.Tables("APTINVHH")
        grdAPTVEND1_CODES.DataSource = dst.Tables("APTVEND1_CODES")
        grdAPTVEND1_STATS.DataSource = dst.Tables("APTVEND1_STATS")

        grdAPTVEND1_FL.DataSource = dst.Tables("APTVEND1_FL")
        grdAPTVEND1_AGING.DataSource = dst.Tables("APTVEND1_AGING")
        grdAPTVEND1_AP_TYPE.DataSource = dst.Tables("APTVEND1_AP_TYPE")
        grdAPTVEND1_OPEN.DataSource = dst.Tables("APTVEND1_OPEN")
        grdAPTVEND1_CODES.DataSource = dst.Tables("APTVEND1_CODES")
        grdAPTVEND1_STATS.DataSource = dst.Tables("APTVEND1_STATS")

        grdAPTINVH1.DataSource = New DataView(dst.Tables("APTINVH1"), "", "", DataViewRowState.CurrentRows)
        grdAPTINVH2.DataSource = dst.Tables("APTINVH2")
        'Call Setup_APTINVH1()
        grdAPTCHCK1.DataSource = dst.Tables("APTCHCK1")
        grdAPTCHCK2.DataSource = dst.Tables("APTCHCK2")

        grdAPTCHCX1.DataSource = dst.Tables("APTCHCX1")

        Dim LYT As UltraWinGrid.UltraGridLayout = grdAPTCHCK1.DisplayLayout
        grdAPTCHCX1.DisplayLayout.Load(LYT, UltraWinGrid.PropertyCategories.All)

        'grdAPTCHCX1.DisplayLayout.Bands(0).Layout.Load(grdAPTCHCK1.DisplayLayout.Bands(0).Layout, UltraWinGrid.PropertyCategories.All)
        'grdAPTCHCX1.DisplayLayout.Bands(1).Layout.Load(grdAPTCHCK2.DisplayLayout.Bands(0).Layout, UltraWinGrid.PropertyCategories.All)



        grdAPTVEND5.DataSource = dst.Tables("APTVEND5")
        grdAPTVEND1_AGEDAP.DataSource = dst.Tables("APTVEND1_AGEDAP")

        'grdPOTORDR1.DataSource = New DataView(dst.Tables("POTORDR1"), "", "", DataViewRowState.CurrentRows)
        'grdPOTORDR2.DataSource = dst.Tables("POTORDR2")
        'grdICTIREC1.DataSource = New DataView(dst.Tables("ICTIREC1"), "", "", DataViewRowState.CurrentRows)
        'grdICTIREC2.DataSource = dst.Tables("ICTIREC2")

        With grdAPTCHCK1.DisplayLayout.Bands("APTCHCK1")
            .Columns("BANK_CODE").Header.Fixed = True
            .Columns("CHECK_NUM").Header.Fixed = True
            .Columns("CHECK_AMT").Header.Fixed = True
            .Columns("CHECK_DATE").Header.Fixed = True
        End With

        'With grdPOTORDR1.DisplayLayout.Bands("POTORDR1")
        '    .Columns("PO_ORDER_NO").Header.Fixed = True
        'End With

        With grdAPTVEND5.DisplayLayout.Bands("APTVEND5")
            '.Columns("VEND_NAME").Header.Fixed = True
            .Groups("Vendor").Header.Fixed = True
        End With

        Create_Summary(grdAPTINVHH, "VEND_CODE", "Count")
        Create_Summary(grdAPTINVHH, "INV_AMT")
        Create_Summary(grdAPTINVHH, "CHECK_AMT")

        Call Create_Summary(grdAPTCHCK2, "SEQ_NUM", "Count", "APTCHCK2")
        Call Create_Summary(grdAPTCHCK2, New String() {"INV_AMT_APPLIED", "INV_DISC_TAKEN", "LC_FEE"})

        Call Create_Summary(grdAPTCHCX1, "SEQ_NUM", "Count", "APTCHCX2")
        Call Create_Summary(grdAPTCHCX1, "INV_AMT_APPLIED", , "APTCHCX2")
        Call Create_Summary(grdAPTCHCX1, "INV_DISC_TAKEN", , "APTCHCX2")

        Call Create_Summary(grdAPTVEND5, "VEND_CODE", "Count")
        Call Create_Summary(grdAPTVEND5, "VEND_PURCHASES_MTD")
        Call Create_Summary(grdAPTVEND5, "VEND_PURCHASES_YTD")
        Call Create_Summary(grdAPTVEND5, "VEND_PURCHASES_LYR")
        Call Create_Summary(grdAPTVEND5, "VEND_PAYMENTS_MTD")
        Call Create_Summary(grdAPTVEND5, "VEND_PAYMENTS_YTD")
        Call Create_Summary(grdAPTVEND5, "VEND_PAYMENTS_LYR")
        Call Create_Summary(grdAPTVEND5, "VEND_DISC_TAKEN_MTD")
        Call Create_Summary(grdAPTVEND5, "VEND_DISC_TAKEN_YTD")
        Call Create_Summary(grdAPTVEND5, "VEND_DISC_TAKEN_LYR")
        Call Create_Summary(grdAPTVEND5, "VEND_NUM_INV_MTD")
        Call Create_Summary(grdAPTVEND5, "VEND_NUM_INV_YTD")
        Call Create_Summary(grdAPTVEND5, "VEND_NUM_INV_LYR")
        Call Create_Summary(grdAPTVEND5, "VEND_NUM_CHKS_MTD")
        Call Create_Summary(grdAPTVEND5, "VEND_NUM_CHKS_YTD")
        Call Create_Summary(grdAPTVEND5, "VEND_NUM_CHKS_LYR")
        Call Create_Summary(grdAPTVEND5, "INV_BALANCE")

        Call Show_Filter(grdAPTVEND5)

        With grdAPTINVH1.DisplayLayout.Bands("APTINVH1")
            .Columns("VOUCHER_NO").Header.Fixed = True
            .Columns("INV_NUM").Header.Fixed = True
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_DATE").Header.Fixed = True
            .Columns("INV_AMT").Header.Fixed = True
        End With

        'With grdICTIREC1.DisplayLayout.Bands("ICTIREC1")
        '    .Columns("RECEIPT_NO").Header.Fixed = True
        'End With

        Call Create_Summary(grdAPTINVH1, "VOUCHER_NO", "Count", "APTINVH1")
        Call Create_Summary(grdAPTINVH1, "INV_AMT", , "APTINVH1")
        Call Create_Summary(grdAPTINVH1, "INV_DISC_BASED_ON", , "APTINVH1")
        Call Create_Summary(grdAPTINVH1, "INV_DISC_AMT", , "APTINVH1")
        Call Create_Summary(grdAPTINVH1, "INV_PAYMENTS", , "APTINVH1")
        Call Create_Summary(grdAPTINVH1, "INV_DISC_TAKEN", , "APTINVH1")
        Call Create_Summary(grdAPTINVH1, "INV_BALANCE", , "APTINVH1")
        Call Create_Summary(grdAPTINVH1, "INV_1099_AMT", , "APTINVH1")

        Call Create_Summary(grdAPTINVH2, "VOUCHER_LNO", "Count", "APTINVH2")
        Call Create_Summary(grdAPTINVH2, "INV_LINE_AMT", , "APTINVH2")


        Call Create_Summary(grdAPTCHCK1, "CHECK_NUM", "Count", "APTCHCK1")
        Call Create_Summary(grdAPTCHCK1, "CHECK_AMT", , "APTCHCK1")

        Call Create_Summary(grdAPTCHCX1, "CHECK_NUM", "Count", "APTCHCX1")
        Call Create_Summary(grdAPTCHCX1, "CHECK_AMT", , "APTCHCX1")

        'Call Create_Summary(grdICTIREC1, "RECEIPT_NO", "Count", "ICTIREC1")
        'Call Create_Summary(grdICTIREC1, "QTY_REC", , "ICTIREC1")
        'Call Create_Summary(grdICTIREC1, "AMT_REC", , "ICTIREC1")

        'Call Create_Summary(grdICTIREC2, "RECEIPT_LNO", "Count", "ICTIREC2")
        'Call Create_Summary(grdICTIREC2, "QTY_REC", , "ICTIREC2")
        'Call Create_Summary(grdICTIREC2, "QTY_INV", , "ICTIREC2")

        'Call Create_Summary(grdPOTORDR1, "PO_ORDER_NO", "Count", "POTORDR1")

        'Call Create_Summary(grdPOTORDR1, "PO_AMT_ORD", , "POTORDR1")
        'Call Create_Summary(grdPOTORDR1, "PO_AMT_REC", , "POTORDR1")
        'Call Create_Summary(grdPOTORDR1, "PO_AMT_OPN", , "POTORDR1")
        'Call Create_Summary(grdPOTORDR1, "PO_AMT_INV", , "POTORDR1")
        'Call Create_Summary(grdPOTORDR1, "PO_AMT_CXL", , "POTORDR1")

        'Call Create_Summary(grdPOTORDR2, "PO_ORDER_LNO", "Count", "POTORDR2")
        'Call Create_Summary(grdPOTORDR2, "PO_QTY_ORD", , "POTORDR2")
        'Call Create_Summary(grdPOTORDR2, "PO_QTY_REC", , "POTORDR2")
        'Call Create_Summary(grdPOTORDR2, "PO_QTY_OPN", , "POTORDR2")
        'Call Create_Summary(grdPOTORDR2, "PO_QTY_INV", , "POTORDR2")
        'Call Create_Summary(grdPOTORDR2, "PO_QTY_CXL", , "POTORDR2")


        Call Set_Read_Only(grpAPTVEND1, True)

        Call Get_PARM("APTPARM1")
        For i As Integer = 1 To 4
            DAYS(i) = Val(ROWs("APTPARM1").Item("AP_PARM_AGE_DAYS_" & CStr(i)) & "")
        Next

        Get_PARM("GLTPARM1")
        Set_SEGS(grdAPTINVH2, "APTINVH2")

        Set_Read_Only(grpAPTVEND1, True)
        '    Set_Read_Only(grpOtherInfo, True)

        grdAPTCHCK1.DisplayLayout.Bands("APTCHCK1").SummaryFooterCaption = "Totals for All Checks"
        ' grdPOTORDR1.DisplayLayout.Bands("POTORDR1").SummaryFooterCaption = "Totals for All Purchase Orders"
        ' grdICTIREC1.DisplayLayout.Bands("ICTIREC1").SummaryFooterCaption = "Totals for All Receipts"
        grdAPTINVH1.DisplayLayout.Bands("APTINVH1").SummaryFooterCaption = "Totals for All AP Items"

        tabMain.Tabs("Purchase Orders").Visible = False
        tabMain.Tabs("Receipts").Visible = False

        optAPTVEND1_AGEDAP.Value = ROWs("APTPARM1").Item("AP_PARM_AGE_INV_OR_DUE")

        With grdAPTCHCK1.DisplayLayout.Bands(0)
            .Columns("DATE_MAILED").Hidden = True
            .Columns("DATE_CLEARED").Hidden = True
        End With
        'ASCMAIN1.Add_Value_List(grdAPTCHCK1, "", Nothing, New String() {":", ":", ":", ":"})
        ASCMAIN1.Add_Value_List(grdAPTCHCK1, "CHECK_STATUS", Nothing, New String() {":", "I:Issued", "V:Voided"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("VEND_CODE")

                If EMsg = "" Then
                    If MENU_ITEM_PP = "S" Then
                        If cdr.Item("VEND_TYPE") & "" <> "S" Then
                            EMsg &= vbCr & "Only Suppliers may be Viewed"
                        End If
                    End If
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
                EntryMode = "I"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Data Filter").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = tf
        'grdAPTVEND5.Visible = Not tf
        'grdAPTINVHH.Visible = Not tf
        tabSummary.Visible = Not tf

        Setup_tabSummary()

        If ScreenMode Then
            'UltraTabControl1.ActiveTab = UltraTabControl1.Tabs("Vendor Name && Address")
            'UltraTabControl1.ActiveTab = UltraTabControl1.Tabs("AP Items")
            tabMain.SelectedTab = tabMain.Tabs("Vendor Name && Address")
            tabMain.SelectedTab = tabMain.Tabs("AP Items")
        Else
            Clear_Record()

        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("VEND_CODE").Text = ""

        If dst.Tables("APTVEND5").Rows.Count = 0 Then
            Fill_Records("APTVEND5")
            Sort_grdColumns(grdAPTVEND5, "VEND_CODE")
        End If

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"APTCHCK1", "APTCHCK2", "APTVEND1", "APTINVH1", "APTINVH2", "APTVEND1_AGEDAP"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        ' dst.Tables("ICTIREC1").Rows.Clear()
        ' dst.Tables("ICTIREC2").Rows.Clear()
        ' dst.Tables("POTORDR1").Rows.Clear()
        ' dst.Tables("POTORDR2").Rows.Clear()

        btnAGE.Visible = True
        tabAGE.Visible = False
        optAPTVEND1_AGEDAP.Visible = False
        lblAPTVEND1_AGEDAP.Visible = False

        EnforceConstraints(True)


        grdAPTCHCK2.DisplayLayout.Bands("APTCHCK2").SummaryFooterCaption = ""
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").SummaryFooterCaption = ""
        ' grdICTIREC2.DisplayLayout.Bands("ICTIREC2").SummaryFooterCaption = ""
        ' grdPOTORDR2.DisplayLayout.Bands("POTORDR2").SummaryFooterCaption = ""

        dst.AcceptChanges()

        'UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(0)
        'UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(2)

        UltraExplorerBar1.Groups("Data Filter").Visible = False
        'UltraExplorerBar1.Groups("Open AP Aged by").Visible = False
    End Sub

    Sub Load_Record()

        Call Save_Header_Fields(UltraGroupBox1)

        dst.EnforceConstraints = False
        grdAPTCHCK1.Tag = "*"
        grdAPTINVH1.Tag = "*"
        grdPOTORDR1.Tag = "*"
        grdICTIREC1.Tag = "*"

        Call Fill_Records("APTVEND1", HFs("VEND_CODE"))
        Call Fill_Records("APTVEND2", HFs("VEND_CODE"))
        dst.EnforceConstraints = True

        Call Setup_APTVEND1_tables()
    End Sub

    Sub Delete_Record()
        Stop
    End Sub

    Sub Update_Record()
        'BeginTrans()

        'CommitTrans("Update")
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdAPTINVH1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Voucher Inquiry", "Show Attachments")
        Load_Popup_Menu(grdAPTCHCK2, "B", "Voucher Inquiry")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                'Case "grdSOTINVH1"
                '    tlb_btn = DirectCast(tlb.Tools("Change to SRep"), UltraWinToolbars.ButtonTool)
                '    tlb_btn.SharedProps.Caption = "Change to SRep " & cmbSREP_CODE.Value
                '    tlb_btn.SharedProps.Visible = (EntryMode = "E")
                '    tlb_btn = DirectCast(tlb.Tools("Change to Comm%"), UltraWinToolbars.ButtonTool)
                '    tlb_btn.SharedProps.Caption = "Change to " & Format(numSREP_COMM_PCT.Value, "#.00") & "% Comm"
                '    tlb_btn.SharedProps.Visible = (EntryMode = "E")
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            'Case ""

            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Value
                Dim rowAPTINVH1 As DataRow = LookUp("APTINVH1", VOUCHER_NO)
                If rowAPTINVH1 IsNot Nothing Then
                    Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                End If

            Case "Show Attachments"
                Dim Entity As Dropped_On_Entity = Dropped_On_Context()
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Value
                Entity.TABLE_NAME = "APTINVH1"
                Entity.CODE_VALUE = VOUCHER_NO
                Entity.COLUMN_NAME = "VOUCHER_NO"

                If Entity.TABLE_NAME <> "" Then
                    Dim F As New ASFATTA1
                    F.ENTITY = Entity
                    F.ShowDialog()
                    F.Dispose()
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                    Call Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "VEND_CODE"
                Call Click_Command("Load")
        End Select
    End Sub

#End Region

 
    Sub Setup_APTINVH1()
        If EntryMode = "" Then
            Exit Sub
        End If

        Dim sql As String = ""
        'grdAPTINVH1.DisplayLayout.Bands("APTINVH1").ColumnFilters.ClearAllFilters()
        If optAPTINVH1.CheckedItem.DataValue = "O" Then
            '    grdAPTINVH1.DisplayLayout.Bands("APTINVH1").ColumnFilters("INV_STATUS").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.NotEquals, "P")
            sql = sql & "(INV_STATUS = 'O' or INV_STATUS = 'H')"
            chkIncludeDeleted.Visible = False
        ElseIf optAPTINVH1.CheckedItem.DataValue = "R" Then
            '    grdAPTINVH1.DisplayLayout.Bands("APTINVH1").ColumnFilters("INV_STATUS").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.NotEquals, "P")
            sql = sql & "(INV_STATUS = 'R')"
            chkIncludeDeleted.Visible = False
        Else
            chkIncludeDeleted.Visible = True
            If Not chkIncludeDeleted.Checked Then
                '    grdAPTINVH1.DisplayLayout.Bands("APTINVH1").ColumnFilters("INV_STATUS").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.NotEquals, "D")
                sql = sql & "(INV_STATUS = 'O' or INV_STATUS = 'H' or INV_STATUS = 'P')"
            End If
        End If
        DirectCast(grdAPTINVH1.DataSource, DataView).RowFilter = sql
        Call Set_grd1stRow(grdAPTINVH1, grdAPTINVH2)

    End Sub

    Private Sub optAPTINVH1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAPTINVH1.ValueChanged
        Call Setup_APTINVH1()
    End Sub

    Private Sub grdAPTVEND5_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTVEND5.DoubleClickRow
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Absx1.txtFor("VEND_CODE").Focus()
        Absx1.txtFor("VEND_CODE").Text = grdAPTVEND5.ActiveRow.Cells("VEND_CODE").Text
        Me.ProcessTabKey(True)
        Call Click_Command("Load")
        Me.Cursor = Cursors.Default
    End Sub
     
    Private Sub grdAPTCHCK1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTCHCK1.AfterRowActivate
        If grdAPTCHCK1.ActiveRow.IsGroupByRow Then
            grdAPTCHCK2.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            grdAPTCHCK2.Visible = True
            Dim BANK_CODE As String = grdAPTCHCK1.ActiveRow.Cells("BANK_CODE").Text
            Dim CHECK_NUM As String = grdAPTCHCK1.ActiveRow.Cells("CHECK_NUM").Text
            Call Fill_Records("APTCHCK2", New String() {BANK_CODE, CHECK_NUM})
            grdAPTCHCK2.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals for Check " & CHECK_NUM
            grdAPTCHCK2.Visible = True
            Me.Cursor = Cursors.Default
        End If
    End Sub
     
    Private Sub grdAPTINVH1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH1.AfterRowActivate
        If grdAPTINVH1.ActiveRow.IsGroupByRow Then
            grdAPTINVH2.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            Dim VOUCHER_NO As String = grdAPTINVH1.ActiveRow.Cells("VOUCHER_NO").Text
            Call Fill_Records("APTINVH2", New String() {VOUCHER_NO})
            grdAPTINVH2.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals for Voucher " & VOUCHER_NO
            grdAPTINVH2.Visible = True
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Sub Setup_APTVEND1_tables()

        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", HFs("VEND_CODE"), True)
        Dim rowAPTVEND5 As DataRow = LookUp("APTVEND5", HFs("VEND_CODE"), True)

        Absx1.chkFor("VEND_ON_HOLD").Text = "On Hold"
        If rowAPTVEND1.Item("VEND_ON_HOLD") & "" = "1" And rowAPTVEND1.Item("VEND_ON_HOLD_DATE") & "" <> "" Then
            Dim z As String = Format(rowAPTVEND1.Item("VEND_ON_HOLD_DATE") & "", "MM/dd/yyyy")
            If z <> "" Then
                Absx1.chkFor("VEND_ON_HOLD").Text = "On Hold since " & ""
            End If
        End If

        grdAPTVEND1_CODES.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("APTVEND1_CODES")
            .Rows.Clear()
            .Rows.Add(New Object() {"Type", rowAPTVEND1.Item("VEND_TYPE"), LookUp("ASTCODE1", New String() {"APTVEND1", "VEND_TYPE", rowAPTVEND1.Item("VEND_TYPE") & ""}, True).Item("T_DESC")})
            .Rows.Add(New Object() {"Method", rowAPTVEND1.Item("VEND_PYMT_METHOD"), IIf(rowAPTVEND1.Item("VEND_PYMT_METHOD_FIXED") & "" = "1", "(Only)", "") & LookUp("ASTCODE1", New String() {"APTVEND1", "VEND_PYMT_METHOD", rowAPTVEND1.Item("VEND_PYMT_METHOD") & ""}, True).Item("T_DESC")})
            .Rows.Add(New Object() {"Terms", rowAPTVEND1.Item("TERM_CODE"), LookUp("TATTERM1", rowAPTVEND1.Item("TERM_CODE") & "", True).Item("TERM_DESC")})
            .Rows.Add(New Object() {"Bank", rowAPTVEND1.Item("BANK_CODE"), LookUp("GLTBANK1", rowAPTVEND1.Item("BANK_CODE") & "", True).Item("BANK_DESC")})
            .Rows.Add(New Object() {"Post", rowAPTVEND1.Item("POST_CODE"), LookUp("APTPOST1", rowAPTVEND1.Item("POST_CODE") & "", True).Item("POST_DESC")})
            .Rows.Add(New Object() {"GL Acct", rowAPTVEND1.Item("ACCT_CODE"), LookUp("GLTACCT1", rowAPTVEND1.Item("ACCT_CODE") & "", True).Item("ACCT_DESC")})
            .Rows.Add(New Object() {"Processor", rowAPTVEND1.Item("PROCESSOR_CODE"), LookUp("ASTUSER1", rowAPTVEND1.Item("PROCESSOR_CODE") & "", True).Item("USER_NAME")})
            .Rows.Add(New Object() {"Buyer", rowAPTVEND1.Item("VEND_BUYER_CODE"), LookUp("ASTUSER1", rowAPTVEND1.Item("VEND_BUYER_CODE") & "", True).Item("USER_NAME")})
            .Rows.Add(New Object() {"Class", rowAPTVEND1.Item("VEND_CLASS_CODE"), LookUp("APTCLAS1", rowAPTVEND1.Item("VEND_CLASS_CODE") & "", True).Item("VEND_CLASS_DESC")})
            .Rows.Add(New Object() {"Pay Vendor", rowAPTVEND1.Item("VEND_CODE_AP"), LookUp("APTVEND1", rowAPTVEND1.Item("VEND_CODE_AP") & "", True).Item("VEND_NAME")})
            .Rows.Add(New Object() {"Pymt Addr", rowAPTVEND1.Item("VEND_PYMT_ADDR"), LookUp("APTVEND2", New String() {HFs("VEND_CODE"), rowAPTVEND1.Item("VEND_PYMT_ADDR") & ""}, True).Item("VEND_ALT_NAME")})
            .Rows.Add(New Object() {"Tax ID", rowAPTVEND1.Item("VEND_TAX_ID"), "Type: " & rowAPTVEND1.Item("VEND_TAX_ID_TYPE") & "; Box: " & rowAPTVEND1.Item("VEND_1099_BOX")})
            .Rows.Add(New Object() {"Cycle", rowAPTVEND1.Item("VEND_PYMT_CYCLE"), LookUp("ASTCODE1", New String() {"APTVEND1", "VEND_PYMT_CYCLE", rowAPTVEND1.Item("VEND_PYMT_CYCLE") & ""}, True).Item("T_DESC")})
        End With

        grdAPTVEND1_STATS.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("APTVEND1_STATS")
            .Rows.Clear()
            .Rows.Add(New Object() {"Purchases", rowAPTVEND5.Item("VEND_PURCHASES_MTD"), rowAPTVEND5.Item("VEND_PURCHASES_YTD"), rowAPTVEND5.Item("VEND_PURCHASES_LYR")})
            .Rows.Add(New Object() {"Discounts", rowAPTVEND5.Item("VEND_DISC_TAKEN_MTD"), rowAPTVEND5.Item("VEND_DISC_TAKEN_YTD"), rowAPTVEND5.Item("VEND_DISC_TAKEN_LYR")})
            .Rows.Add(New Object() {"Payments", rowAPTVEND5.Item("VEND_PAYMENTS_MTD"), rowAPTVEND5.Item("VEND_PAYMENTS_YTD"), rowAPTVEND5.Item("VEND_PAYMENTS_LYR")})
            .Rows.Add(New Object() {"Invoices", rowAPTVEND5.Item("VEND_NUM_INV_MTD"), rowAPTVEND5.Item("VEND_NUM_INV_YTD"), rowAPTVEND5.Item("VEND_NUM_INV_LYR")})
            .Rows.Add(New Object() {"Checks", rowAPTVEND5.Item("VEND_NUM_CHKS_MTD"), rowAPTVEND5.Item("VEND_NUM_CHKS_YTD"), rowAPTVEND5.Item("VEND_NUM_CHKS_LYR")})
        End With


        ASCMAIN1.sql = "Select " _
        & "  Sum (INV_BALANCE) OPEN" _
        & ", Sum (Decode(INV_TYPE,'I',INV_BALANCE,0)) OPEN_I" _
        & ", Sum (Decode(INV_TYPE,'C',INV_BALANCE,0)) OPEN_C" _
        & ", Sum (Decode(INV_TYPE,'D',INV_BALANCE,0)) OPEN_D" _
        & ", Sum (Decode(INV_TYPE,'A',INV_BALANCE,0)) OPEN_A" _
        & ", Sum (Decode(INV_TYPE,'B',INV_BALANCE,0)) OPEN_B" _
        & ", Sum (Decode(INV_TYPE,'R',INV_BALANCE,0)) OPEN_R" _
        & " from APTINVH1 where VEND_CODE = '" & HFs("VEND_CODE") & "'" _
        & "  and INV_STATUS in ('O','H')"
        Dim rowOPENAP As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select " _
        & "  Sum (POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST) OPEN" _
        & " from POTORDR1,POTORDR2 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO " _
        & "  and POTORDR1.VEND_CODE = '" & HFs("VEND_CODE") & "'" _
        & "  and POTORDR1.PO_STATUS = 'O'"
        Dim rowOPENPO As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select " _
        & "  Sum (ICTIREC2.QTY_REC * ICTIREC2.PO_COST) OPEN" _
        & " from ICTIREC1,ICTIREC2 where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO " _
        & "  and ICTIREC1.VEND_CODE = '" & HFs("VEND_CODE") & "'" _
        & "  and ICTIREC1.ACCRUAL_STATUS = '0'"
        Dim rowOPENIC As DataRow = ASCDATA1.GetDataRow


        grdAPTVEND1_AP_TYPE.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("APTVEND1_AP_TYPE")
            .Rows.Clear()
            .Rows.Add(New Object() {"Invoices", Val(rowOPENAP.Item("OPEN_I") & "")})
            .Rows.Add(New Object() {"CR Memos", Val(rowOPENAP.Item("OPEN_C") & "")})
            .Rows.Add(New Object() {"DR Memos", Val(rowOPENAP.Item("OPEN_D") & "")})
            .Rows.Add(New Object() {"Advances", Val(rowOPENAP.Item("OPEN_A") & "")})
            .Rows.Add(New Object() {"Chargebacks", Val(rowOPENAP.Item("OPEN_B") & "")})
            .Rows.Add(New Object() {"Returns", Val(rowOPENAP.Item("OPEN_R") & "")})
            .Rows.Add(New Object() {"Totals", Val(rowOPENAP.Item("OPEN") & "")})
        End With

        grdAPTVEND1_FL.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("APTVEND1_FL")
            .Rows.Clear()
            .Rows.Add(New Object() {"1st Purchase", Null, rowAPTVEND5.Item("VEND_1ST_PURCH_DATE"), Null})
            .Rows.Add(New Object() {"Last Invoice", rowAPTVEND5.Item("VEND_LAST_INV_NUM"), rowAPTVEND5.Item("VEND_LAST_INV_DATE"), rowAPTVEND5.Item("VEND_LAST_INV_AMT")})
            .Rows.Add(New Object() {"Last Payment", rowAPTVEND5.Item("VEND_LAST_CHECK_NUM"), rowAPTVEND5.Item("VEND_LAST_PMT_DATE"), rowAPTVEND5.Item("VEND_LAST_PMT_AMT")})
        End With

        grdAPTVEND1_OPEN.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("APTVEND1_OPEN")
            .Rows.Clear()
            .Rows.Add(New Object() {"Net A/P", Val(rowOPENAP.Item("OPEN") & "")})
            .Rows.Add(New Object() {"Accrued Purchases", Val(rowOPENIC.Item("OPEN") & "")})
            .Rows.Add(New Object() {"Open PO's", Val(rowOPENPO.Item("OPEN") & "")})
            '  .Rows.Add(New Object() {"Accrued Purchases", Val(rowOPENIC.Item("OPEN") & "")})
            '  .Rows.Add(New Object() {"Open PO's", Val(rowOPENPO.Item("OPEN") & "")})
            Dim T As Decimal = Val(rowOPENAP.Item("OPEN") & "") + Val(rowOPENIC.Item("OPEN") & "") + Val(rowOPENPO.Item("OPEN") & "")
            .Rows.Add(New Object() {"Total Commitments", T})
        End With

    End Sub

    Function Calc_Date(ByVal F As Integer, ByVal AP_PARM_AGE_CATG As Integer) As String
        Return "'" & Format(Now.AddDays(F * DAYS(AP_PARM_AGE_CATG)), "MM/dd/yyyy") & "'"
    End Function
     
    Private Sub grdPOTORDR1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTORDR1.AfterRowActivate
        If grdPOTORDR1.ActiveRow.IsGroupByRow Then
            grdPOTORDR2.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            Dim PO_ORDER_NO As String = grdPOTORDR1.ActiveRow.Cells("PO_ORDER_NO").Text
            Call Fill_Records("POTORDR2", New String() {PO_ORDER_NO})
            grdPOTORDR2.DisplayLayout.Bands("POTORDR2").SummaryFooterCaption = "Totals for PO " & PO_ORDER_NO
            grdPOTORDR2.Visible = True
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub grdICTIREC1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIREC1.AfterRowActivate
        If grdICTIREC1.ActiveRow.IsGroupByRow Then
            grdICTIREC2.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            Dim RECEIPT_NO As String = grdICTIREC1.ActiveRow.Cells("RECEIPT_NO").Text
            Call Fill_Records("ICTIREC2", New String() {RECEIPT_NO})
            grdICTIREC2.DisplayLayout.Bands("ICTIREC2").SummaryFooterCaption = "Totals for Receipt " & RECEIPT_NO
            grdICTIREC2.Visible = True
            Me.Cursor = Cursors.Default
        End If
    End Sub
 
    Sub Setup_ICTIREC1()
        If EntryMode = "" Then
            Exit Sub
        End If

        Dim sql As String = ""
        'grdICTIREC1.DisplayLayout.Bands("ICTIREC1").ColumnFilters.ClearAllFilters()
        If optICTIREC1.CheckedItem.DataValue = "O" Then
            'grdICTIREC1.DisplayLayout.Bands("ICTIREC1").ColumnFilters("ACCRUAL_STATUS").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.NotEquals, "1")
            sql = sql & "(ACCRUAL_STATUS = '0')"
        End If
        DirectCast(grdICTIREC1.DataSource, DataView).RowFilter = sql
        Call Set_grd1stRow(grdICTIREC1, grdICTIREC2)

    End Sub

    Sub Setup_POTORDR1()
        If EntryMode = "" Then
            Exit Sub
        End If

        Dim sql As String = ""
        'grdPOTORDR1.DisplayLayout.Bands("POTORDR1").ColumnFilters.ClearAllFilters()
        If optPOTORDR1.CheckedItem.DataValue = "O" Then
            'grdPOTORDR1.DisplayLayout.Bands("POTORDR1").ColumnFilters("PO_STATUS_CODE").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.Equals, "O")
            sql = "PO_STATUS_CODE = 'O'"
        End If
        DirectCast(grdPOTORDR1.DataSource, DataView).RowFilter = sql
        Call Set_grd1stRow(grdPOTORDR1, grdPOTORDR2)

    End Sub

    Sub Setup_APTCHCK1()

    End Sub
    Private Sub optPOTORDR1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPOTORDR1.ValueChanged
        Call Setup_POTORDR1()
    End Sub

    Private Sub optICTIREC1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optICTIREC1.ValueChanged
        Call Setup_ICTIREC1()
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Call Setup_Tab(tabMain.ActiveTab.Key)
    End Sub

    Sub Load_APTCHCK1()
        Call ASCMAIN1.Progress("Now Loading Payment History")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Call Fill_Records("APTCHCK1", HFs("VEND_CODE"))
        grdAPTCHCK1.DisplayLayout.Bands("APTCHCK1").SortedColumns.Clear()
        grdAPTCHCK1.DisplayLayout.Bands("APTCHCK1").SortedColumns.Add("CHECK_DATE", True)
        Call Set_grd1stRow(grdAPTCHCK1, grdAPTCHCK2)
        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Load_APTINVH1()
        Call ASCMAIN1.Progress("Now Loading AP Items")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Call Fill_Records("APTINVH1", HFs("VEND_CODE"))
        grdAPTINVH1.DisplayLayout.Bands("APTINVH1").SortedColumns.Clear()
        grdAPTINVH1.DisplayLayout.Bands("APTINVH1").SortedColumns.Add("VOUCHER_NO", True)
        Call Set_grd1stRow(grdAPTINVH1, grdAPTINVH2)
        Me.Cursor = Cursors.Default

        Call Setup_APTINVH1()
        Call ASCMAIN1.Progress("")

        Call Age_AP_Items_by_Date()

    End Sub

    Sub Load_POTORDR1()

        Dim Sql As String = "Select POTORDR2.PO_ORDER_NO " _
        & ", SUM (POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST) PO_AMT_ORD" _
        & ", SUM (POTORDR2.PO_QTY_REC * POTORDR2.PO_COST) PO_AMT_REC" _
        & ", SUM (POTORDR2.PO_QTY_INV * POTORDR2.PO_COST) PO_AMT_INV" _
        & ", SUM (POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST) PO_AMT_OPN " _
        & ", SUM (POTORDR2.PO_QTY_CXL * POTORDR2.PO_COST) PO_AMT_CXL " _
        & " from POTORDR2,POTORDR1 " _
        & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO " _
        & " and POTORDR1.VEND_CODE = '" & HFs("VEND_CODE") & "'" _
        & " group by POTORDR2.PO_ORDER_NO"

        Call ASCMAIN1.Progress("Now Loading Purchase Orders")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR1_TOTALS)
        ASCDATA1.ExecuteSQL("Insert into " & POTORDR1_TOTALS & " " & Sql)
        Call Fill_Records("POTORDR1", HFs("VEND_CODE"))
        grdPOTORDR1.DisplayLayout.Bands("POTORDR1").SortedColumns.Clear()
        grdPOTORDR1.DisplayLayout.Bands("POTORDR1").SortedColumns.Add("PO_ORDER_NO", True)
        Call Set_grd1stRow(grdPOTORDR1, grdPOTORDR2)
        Me.Cursor = Cursors.Default

        Call Setup_POTORDR1()
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Load_ICTIREC1()
        Call ASCMAIN1.Progress("Now Loading Receipts")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTIREC1_TOTALS)
        ASCDATA1.ExecuteSQL("Insert into " & ICTIREC1_TOTALS & " Select ICTIREC2.RECEIPT_NO, SUM (QTY_REC) QTY_REC, SUM (QTY_REC * PO_COST) AMT_REC from ICTIREC1,ICTIREC2 where ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO and ICTIREC1.VEND_CODE = '" & HFs("VEND_CODE") & "' group by ICTIREC2.RECEIPT_NO")
        ASCDATA1.ExecuteSQL("Truncate Table " & ICTIREC1)
        ASCDATA1.ExecuteSQL("Insert into " & ICTIREC1 & " Select ICTIREC1.*, ICTIREC1_TOTALS.QTY_REC, ICTIREC1_TOTALS.AMT_REC from ICTIREC1," & ICTIREC1_TOTALS & " ICTIREC1_TOTALS where ICTIREC1.VEND_CODE = '" & HFs("VEND_CODE") & "' and ICTIREC1_TOTALS.RECEIPT_NO = ICTIREC1.RECEIPT_NO")
        Call Fill_Records("ICTIREC1")

        grdICTIREC1.DisplayLayout.Bands("ICTIREC1").SortedColumns.Clear()
        grdICTIREC1.DisplayLayout.Bands("ICTIREC1").SortedColumns.Add("RECEIPT_NO", True)
        Call Set_grd1stRow(grdICTIREC1, grdICTIREC2)
        Me.Cursor = Cursors.Default

        Call Setup_ICTIREC1()
        Call ASCMAIN1.Progress("")
    End Sub

    Private Sub chkIncludeDeleted_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIncludeDeleted.CheckedChanged
        Call Setup_APTINVH1()
    End Sub

    Private Sub chkShowFilters_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowFilters.CheckedChanged
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Call Show_Filter(grdAPTINVH1, chkShowFilters.Checked)
        Call Show_Filter(grdAPTCHCK1, chkShowFilters.Checked)
        Call Show_Filter(grdPOTORDR1, chkShowFilters.Checked)
        Call Show_Filter(grdICTIREC1, chkShowFilters.Checked)

        If Not chkShowFilters.Checked Then
            grdAPTINVH1.DisplayLayout.Bands("APTINVH1").ColumnFilters.ClearAllFilters()
            grdAPTCHCK1.DisplayLayout.Bands("APTCHCK1").ColumnFilters.ClearAllFilters()
            grdPOTORDR1.DisplayLayout.Bands("POTORDR1").ColumnFilters.ClearAllFilters()
            grdICTIREC1.DisplayLayout.Bands("ICTIREC1").ColumnFilters.ClearAllFilters()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdAPTVEND1_AP_TYPE_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTVEND1_AP_TYPE.InitializeRow
        If e.Row.Cells(0).Text Like "Total*" Then
            e.Row.Appearance.BackColor = Drawing.Color.Beige
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
        End If
    End Sub

    Private Sub grdAPTVEND1_OPEN_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTVEND1_OPEN.InitializeRow
        If e.Row.Cells(0).Text Like "Total*" Then
            e.Row.Appearance.BackColor = Drawing.Color.Beige
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
        End If
    End Sub

    Private Sub grdAPTVEND1_AGING_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTVEND1_AGING.InitializeRow
        If e.Row.Cells(0).Text Like "Total*" Then
            e.Row.Appearance.BackColor = Drawing.Color.Beige
            e.Row.Appearance.FontData.Bold = DefaultableBoolean.True
        End If
    End Sub

    Private Sub optAPTVEND1_AGEDAP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAPTVEND1_AGEDAP.ValueChanged
        Call Age_AP_Items_by_Date()
    End Sub

    Sub Age_AP_Items_by_Date(Optional ByVal AgingDaysOnly As Boolean = False)
        If dst.Tables.Count = 0 Or EntryMode = "" Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Calculating Aging")
        Application.DoEvents()

        Dim AGE_DATE_COLUMN As String = ""
        If optAPTVEND1_AGEDAP.Value & "" = "I" Then
            AGE_DATE_COLUMN = "INV_DATE"
        Else
            AGE_DATE_COLUMN = "INV_DUE_DATE"
        End If

        If Not AgingDaysOnly Then

            With dst.Tables("APTVEND1_AGEDAP")
                .Rows.Clear()
                For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("INV_STATUS = 'O' or INV_STATUS = 'H'", "")
                    Dim AGE_DATE As Date
                    AGE_DATE = rowAPTINVH1.Item(AGE_DATE_COLUMN)
                    Dim row As DataRow = .Rows.Find(AGE_DATE)
                    If row Is Nothing Then
                        row = .NewRow
                        row.Item("AGE_DATE") = AGE_DATE
                        row.Item("AGE_AMT") = 0
                        .Rows.Add(row)
                    End If
                    row.Item("AGE_AMT") += Val(rowAPTINVH1.Item("INV_BALANCE") & "")
                Next
            End With

            With grdAPTVEND1_AGEDAP.DisplayLayout.Bands(0).SortedColumns
                .Clear()
                .Add("AGE_DATE", False)
            End With

            If grdAPTVEND1_AGEDAP.Rows.Count <> 0 Then
                grdAPTVEND1_AGEDAP.ActiveRow = grdAPTVEND1_AGEDAP.Rows(0)
            End If

        End If

        grdAPTVEND1_AGING.DisplayLayout.Bands(0).SortedColumns.Clear()
        With dst.Tables("APTVEND1_AGING")
            .Rows.Clear()
            Dim AGE_CATGY As String
            Dim AGE_WHERE As String
            Dim T As Double = 0
            Dim AGE_AMT As Double = 0

            For i As Integer = 0 To 5
                If i = 0 Then
                    AGE_CATGY = "Current"
                    If optAgingDirection.Value = "F" Then
                        AGE_WHERE = AGE_DATE_COLUMN & " <= " & Calc_Date(1, 0)
                    Else
                        AGE_WHERE = AGE_DATE_COLUMN & " >= " & Calc_Date(-1, 0)
                    End If
                ElseIf i = 5 Then
                    AGE_CATGY = "Over " & CStr(DAYS(i - 1))
                    If optAgingDirection.Value = "F" Then
                        AGE_WHERE = AGE_DATE_COLUMN & " > " & Calc_Date(1, 4)
                    Else
                        AGE_WHERE = AGE_DATE_COLUMN & " < " & Calc_Date(-1, 4)
                    End If
                Else
                    AGE_CATGY = Format(DAYS(i - 1) + 1, "000") & "-" & Format(DAYS(i), "000")
                    If optAgingDirection.Value = "F" Then
                        AGE_WHERE = AGE_DATE_COLUMN & " >= " & Calc_Date(1, i - 1) & " and " & AGE_DATE_COLUMN & " < " & Calc_Date(1, i)
                    Else
                        AGE_WHERE = AGE_DATE_COLUMN & " < " & Calc_Date(-1, i - 1) & " and " & AGE_DATE_COLUMN & " >= " & Calc_Date(-1, i)
                    End If
                End If
                AGE_AMT = Val(dst.Tables("APTINVH1").Compute("SUM (INV_BALANCE)", "(INV_STATUS = 'O' or INV_STATUS = 'H') and " & AGE_WHERE) & "")
                T = T + AGE_AMT

                If i <> 0 Then
                    If optAgingDirection.Value = "F" Then
                        AGE_CATGY = "Due Next " & AGE_CATGY
                    Else
                        AGE_CATGY = AGE_CATGY & " Past Due"
                    End If
                End If

                .Rows.Add(New Object() {AGE_CATGY, AGE_AMT})
            Next
            .Rows.Add(New Object() {"Total", T})
        End With

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Private Sub btnAGE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAGE.Click
        Call Setup_Tab("AP Items")
    End Sub

    Sub Setup_Tab(ByVal tabKey As String)
        Dim tf As Boolean = False
        optAPTINVH1.Visible = False
        optPOTORDR1.Visible = False
        optICTIREC1.Visible = False
        chkIncludeDeleted.Visible = False

        Select Case tabKey
            Case "Vendor Name && Address"
            Case "Info"
            Case "Vendor Name && Address"
            Case "AP Items"
                If grdAPTINVH1.Tag = "*" Then
                    Load_APTINVH1()
                    grdAPTINVH1.Tag = ""
                End If
                optAPTINVH1.Visible = True
                tf = True
                btnAGE.Visible = False
                tabAGE.Visible = True
                optAPTVEND1_AGEDAP.Visible = True
                lblAPTVEND1_AGEDAP.Visible = True
                'chkIncludeDeleted.Visible = True
            Case "Payments"
                If grdAPTCHCK1.Tag = "*" Then
                    Load_APTCHCK1()
                    grdAPTCHCK1.Tag = ""
                End If
                tf = True
            Case "Purchase Orders"
                If grdPOTORDR1.Tag = "*" Then
                    Load_POTORDR1()
                    grdPOTORDR1.Tag = ""
                End If
                optPOTORDR1.Visible = True
                tf = True
            Case "Receipts"
                If grdICTIREC1.Tag = "*" Then
                    Load_ICTIREC1()
                    grdICTIREC1.Tag = ""
                End If
                optICTIREC1.Visible = True
                tf = True
        End Select
        UltraExplorerBar1.Groups("Data Filter").Visible = tf

    End Sub

    Private Sub optAgingDirection_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAgingDirection.ValueChanged
        Call Age_AP_Items_by_Date(True)
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
     ByVal ctl As Control, _
     ByVal COLUMN_NAME As String, _
     Optional ByRef sql_where As String = "", _
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "VEND_CODE"
                If MENU_ITEM_PP = "S" Then
                    sql_where = "VEND_TYPE = 'S'"
                End If
        End Select
    End Sub

    Private Sub cmdFetch_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetch.Click
        grdAPTINVHH.Text = "Purchase Summary between " & Absx1.cmbFor("RYP0").Text & " and " & Absx1.cmbFor("RYP1").Text
        Dim RYP0 As String = Absx1.cmbFor("RYP0").Value & ""
        RYP0 = Mid(RYP0, 1, 4) & Mid(RYP0, 6, 2)

        Dim RYP1 As String = Absx1.cmbFor("RYP1").Value & ""
        RYP1 = Mid(RYP1, 1, 4) & Mid(RYP1, 6, 2)

        Fill_Records("APTINVHH", New String() {RYP0, RYP1})
        Sort_grdColumns(grdAPTINVHH, "VEND_CODE")

        grdAPTINVHH.Visible = True
        grdAPTVEND5.Visible = False
    End Sub

    Private Sub grdAPTINVHH_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTINVHH.DoubleClickRow
        Me.Cursor = Cursors.WaitCursor
        Absx1.txtFor("VEND_CODE").Focus()
        Absx1.txtFor("VEND_CODE").Text = grdAPTINVHH.ActiveRow.Cells("VEND_CODE").Text
        Me.ProcessTabKey(True)
        Click_Command("Load")
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tabSummary_SelectedTabChanged_1(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSummary.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_tabSummary()
    End Sub

    Sub Setup_tabSummary()
        UltraExplorerBar1.Groups("Period Range").Visible = (tabSummary.SelectedTab.Key = "Purchase Summary") And Not ScreenMode
    End Sub

    Private Sub grdAPTINVHH_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTINVHH.InitializeLayout

    End Sub
End Class