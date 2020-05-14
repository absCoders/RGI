Imports System.Math
Imports Infragistics.Win.UltraWinGrid

Public Class APFINVH1
    Dim rowAPTVEND1 As DataRow
    Dim rowAPTINVH1 As DataRow
    Dim rowAPTVEND5 As DataRow
    Dim ICTIREC1 As String
    Dim ICTIREC1_TOTALS As String
    Dim sql_APTINVR2 As String
    Dim sql_APTVEND1 As String
    Dim batch_update As Boolean = False
    Dim BANK_LAST_CHECK_NO As String
    Dim BANK_NEXT_CHECK_NO As String
    Dim auto_next_check As Boolean = False
    Private discrepancies_only As Boolean = False
    Dim disable_calculate_totals As Boolean = False
    Dim POTSHIP3 As String
    Dim POTSHIP2 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "APFINVHI" Then
            InquiryMode = True
        End If

        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")
        Get_PARM("ICTPARM1")

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("Load").Visible = InquiryMode
            .Items("Done").Visible = InquiryMode
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Delete").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
            .Items("Print Edit").Visible = Not InquiryMode
            .Items("New Batch").Visible = Not InquiryMode
            .Items("Multi-Invoice Edit").Visible = Not InquiryMode
        End With


        With dst
            ASCMAIN1.sql = "SELECT * FROM ASTAUDT1" _
                    & " WHERE ASTAUDT1.TABLE_NAME = 'APTINVH1' and NVL(FM_MODE,'X') <> 'L'" _
                    & " AND ASTAUDT1.KEY_VALUE = :PARM1"
            Create_TDA(.Tables.Add, "ASTAUDTX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "APTVEND1", "*")
            Create_TDA(.Tables.Add, "APTVEND2", "*")
            Create_TDA(.Tables.Add, "APTVEND5", "*")
            Create_TDA(.Tables.Add, "APTINVH1", "*")

            Create_TDA(.Tables.Add, "APTCHCK1", "*")
            Create_TDA(.Tables.Add, "APTCHCK2", "*")
            Create_TDA(.Tables.Add, "GLTBANK1", "*")

            ASCMAIN1.sql = "Select APTINVH2.*, GLTACCT1.ACCT_DESC from APTINVH2,GLTACCT1 where APTINVH2.ACCT_CODE = GLTACCT1.ACCT_CODE"
            Create_TDA(.Tables.Add, "APTINVH2", "**", 1)

            Create_TDA(.Tables.Add, "APTINVH8", "*", 1)

            ASCMAIN1.sql = "Select APTINVH7.*,POTLCST1.COST_CATGY_CODE" & vbCrLf _
                & ",POTLCST1.PO_ORDER_NO,POTLCST1.PO_SHIPMENT_NO,POTLCST1.PO_SHIPMENT_LNO" & vbCrLf _
                & ",POTSHIP1.PO_SHIP_VESSEL,POTSHIP1.PO_SHIP_REF_NO" & vbCrLf _
                & ",POTSHIP2.CONTAINER_NO,POTSHIP2.BOL_NO,POTSHIP2.COMM_INV_NO" & vbCrLf _
                & " from APTINVH7,POTLCST1,POTORDR1,POTSHIP1,POTSHIP2" & vbCrLf _
                & " where POTLCST1.CTL_NO = APTINVH7.CTL_NO and APTINVH7.VOUCHER_NO = :PARM1" & vbCrLf _
                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTLCST1.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO (+) = POTLCST1.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "APTINVH7", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select POTLCST1.*" & vbCrLf _
                & ",POTSHIP1.PO_SHIP_VESSEL,POTSHIP1.PO_SHIP_REF_NO" & vbCrLf _
                & ",POTSHIP2.CONTAINER_NO,POTSHIP2.BOL_NO,POTSHIP2.COMM_INV_NO" & vbCrLf _
                & " from POTLCST1,POTORDR1,POTSHIP1,POTSHIP2" & vbCrLf _
                & " where POTSHIP1.PO_SHIPMENT_NO (+) = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTLCST1.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO (+) = POTLCST1.PO_ORDER_NO" & vbCrLf _
                & "   and (POTLCST1.VOUCHER_NO = :PARM1 or (POTLCST1.VOUCHER_NO is Null and POTLCST1.COST_ACC <> 0))"
            Create_TDA(.Tables.Add, "POTLCST1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO, APTINVH1.INV_NUM" _
            & ", APTINVH1.INV_DATE" _
            & ", NVL(APTINVH1.INV_BALANCE,0) INV_BALANCE" _
            & ", NVL(APTINVH1.INV_DISC_AMT,0) INV_DISC_AMT" _
            & " from APTINVH1 " _
            & " where VEND_CODE = :PARM1 and VOUCHER_NO <> :PARM2 " _
            & " and INV_STATUS = 'O' and INV_TYPE in ('A','I','C','D')"
            Create_TDA(.Tables.Add, "APTINVHX", "**", 0, False, "VV", 1)
            .Tables("APTINVHX").Columns.Add("INV_PAYMENTS", GetType(System.Decimal), "INV_BALANCE - INV_DISC_AMT")
            .Tables("APTINVHX").Columns.Add("SELECTED", GetType(System.String))
            .Tables("APTINVHX").Columns("SELECTED").DefaultValue = "0"

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO" _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_AMT" _
            & ", APTINVH1.INV_TYPE, APTINVH1.TERM_CODE" _
            & ", APTINVH1.INV_PYMT_METHOD, APTINVH1.INV_PYMT_CYCLE" _
            & ", APTINVH1.BANK_CODE, APTINVH1.INV_REF" _
            & ", APTINVH1.VEND_ALT_CODE" _
            & " from APTINVH1 where APTINVH1.INV_STATUS in ('H','O') and APTINVH1.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTINVHM", "**", 0, False, "V")
            .Tables("APTINVHM").Columns.Add("UPDATE_STATUS", GetType(System.String))
            .Tables("APTINVHM").Columns.Add("UPDATE_MESSAGE", GetType(System.String))

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO" _
            & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME" _
            & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_AMT" _
            & ", APTINVH1.INV_TYPE, APTINVH1.INV_STATUS, APTINVH1.TERM_CODE" _
            & ", APTINVH1.INV_PYMT_METHOD, APTINVH1.INV_PYMT_CYCLE" _
            & ", APTINVH1.BANK_CODE, APTINVH1.INV_REF" _
            & ", APTINVH1.CHECK_NUM, APTINVH1.CHECK_DATE, APTINVH1.VEND_ALT_CODE, APTINVH1.POST_CODE" _
            & " from APTINVH1,APTVEND1 where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE"
            Create_TDA(.Tables.Add, "APTINVHB", "**", 0, False)
            .Tables("APTINVHB").Columns.Add("UPDATE_STATUS", GetType(System.String))
            .Tables("APTINVHB").Columns.Add("UPDATE_MESSAGE", GetType(System.String))

            ASCMAIN1.sql = "Select APTINVH1.* from APTINVH1"
            Create_TDA(.Tables.Add, "APTINVR1", "**", 0, False)
            .Tables("APTINVR1").Columns.Add("CHECK_AMT", GetType(System.Decimal))
            .Tables("APTINVR1").Columns.Add("CHECK_AMT_OTHERS", GetType(System.Decimal))
            .Tables("APTINVR1").Columns.Add("INV_AMT_GL", GetType(System.Decimal))

            ASCMAIN1.sql = "Select APTINVH2.*, DECODE(APTINVH2.INV_LTYP,NULL,APTINVH2.INV_LINE_AMT,0) INV_LINE_AMT_GL, GLTACCT1.ACCT_DESC from APTINVH2,GLTACCT1 where APTINVH2.ACCT_CODE = GLTACCT1.ACCT_CODE"
            Create_TDA(.Tables.Add, "APTINVR2", "**", 0, False)
            .Tables("APTINVR2").Columns.Add("OPS_YYYYPP", GetType(System.String))

            .Relations.Add("APTINVR2",
            .Tables("APTINVR1").Columns("VOUCHER_NO"),
            .Tables("APTINVR2").Columns("VOUCHER_NO"))

            .Tables("APTINVR1").Columns("INV_AMT_GL").Expression = "SUM(CHILD(APTINVR2).INV_LINE_AMT_GL)"
            .Tables("APTINVR2").Columns("OPS_YYYYPP").Expression = "PARENT(APTINVR2).OPS_YYYYPP"

            ', ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO" & vbCrLf _

            ASCMAIN1.sql = "Select APTINVH5.*, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & ", ICTIREC2.QTY_REC, ICTIREC2.PO_COST, ICTIREC2.AP_COST" & vbCrLf _
                & ", ICTIREC2.STYLE_CODE, ICTIREC2.STYLE_UOM" & vbCrLf _
                & " from ICTIREC2,APTINVH5,ICTSTYL1" & vbCrLf _
                & " where APTINVH5.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                & "   and APTINVH5.RECEIPT_LNO = ICTIREC2.RECEIPT_LNO" & vbCrLf _
                & "   and APTINVH5.VOUCHER_NO = :PARM1" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = ICTIREC2.STYLE_CODE" & vbCrLf _
                & "   and APTINVH5.RECEIPT_NO IS NOT NULL"

            ASCMAIN1.sql &= " union " & vbCrLf _
                & "Select APTINVH5.*, ICTSTYL1.STYLE_DESC" & vbCrLf _
                    & ", POTSHIP3.PO_QTY_SHP QTY_REC, POTSHIP3.PO_COST, POTSHIP3.PO_COST AP_COST" & vbCrLf _
                    & ", POTORDR2.STYLE_CODE, ICTSTYL1.STYLE_UOM" & vbCrLf _
                    & " from APTINVH5,POTSHIP3,POTSHIP2,POTORDR2,ICTSTYL1" & vbCrLf _
                    & " where APTINVH5.VOUCHER_NO = :PARM1" _
                    & "   and POTSHIP2.VOUCHER_NO = APTINVH5.VOUCHER_NO" & vbCrLf _
                    & "   and POTSHIP3.PO_SHIPMENT_NO = APTINVH5.PO_SHIPMENT_NO" & vbCrLf _
                    & "   and POTSHIP3.PO_SHIPMENT_LNO = APTINVH5.PO_SHIPMENT_LNO" & vbCrLf _
                    & "   and POTSHIP3.PO_ORDER_NO = APTINVH5.PO_ORDER_NO" & vbCrLf _
                    & "   and POTSHIP3.PO_ORDER_LNO = APTINVH5.PO_ORDER_LNO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                    & "   and APTINVH5.RECEIPT_NO IS NULL"

            Create_TDA(.Tables.Add, "APTINVH5", "**", 0, True, "V", 2)
            With .Tables("APTINVH5").Columns
                .Add("AMT_REC", GetType(System.Decimal), "QTY_REC * AP_COST")
                .Add("AMT_INV", GetType(System.Decimal), "INV_QTY * INV_COST")
                .Add("QTY_VAR", GetType(System.Int32), "INV_QTY - QTY_REC")
                .Add("AMT_VAR", GetType(System.Decimal), "AMT_INV - AMT_REC")
            End With

            .Tables.Add("APTINVH5_VAR")
            With .Tables("APTINVH5_VAR")
                .Columns.Add("STYLE_CLASS_CODE")
                .Columns.Add("ACCT_CODE_PPV")
                .Columns.Add("SEG2_CODE")
                .Columns.Add("AMT_REC", GetType(System.Decimal))
                .Columns.Add("AMT_INV", GetType(System.Decimal))
                .Columns.Add("AMT_VAR", GetType(System.Decimal))
                .Columns.Add("AMT_VAR_CB", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("STYLE_CLASS_CODE")}
            End With

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'APTINVH1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, False, "V", 0)

            .Tables.Add("APTINVH5_SUM")
            With .Tables("APTINVH5_SUM")
                .Columns.Add("RECEIPT_NO", GetType(System.String))
                .Columns.Add("RECEIPT_DATE", GetType(System.DateTime))
                .Columns.Add("PO_ORDER_NO", GetType(System.String))
                .Columns.Add("PO_REFERENCE", GetType(System.String))
                .Columns.Add("PO_SHIPMENT_NO", GetType(System.String))
                .Columns.Add("PO_SHIPMENT_LNO", GetType(System.Int32))
                .Columns.Add("PO_SHIP_VESSEL", GetType(System.String))
                .Columns.Add("PO_SHIP_REF_NO", GetType(System.String))
                .Columns.Add("PO_DATE_SHIPPED", GetType(System.DateTime))
                .Columns.Add("PORT_CODE", GetType(System.String))
                .Columns.Add("WHSE_CODE", GetType(System.String))
                .Columns.Add("CONTAINER_NO", GetType(System.String))
                .Columns.Add("BOL_NO", GetType(System.String))
                .Columns.Add("COMM_INV_NO", GetType(System.String))
                '.Columns.Add("QTY_REC", GetType(System.Int32))
                '.Columns.Add("AMT_REC", GetType(System.Decimal))
                '.Columns.Add("QTY_INV", GetType(System.Int32))
                '.Columns.Add("AMT_INV", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("RECEIPT_NO")}
            End With

            .Relations.Add("APTINVH5", .Tables("APTINVH5_SUM").Columns("RECEIPT_NO"), .Tables("APTINVH5").Columns("RECEIPT_NO"))

            With .Tables("APTINVH5_SUM")
                .Columns.Add("QTY_REC", GetType(System.Int32), "SUM(Child(APTINVH5).QTY_REC)")
                .Columns.Add("AMT_REC", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_REC)")
                .Columns.Add("QTY_INV", GetType(System.Int32), "SUM(Child(APTINVH5).INV_QTY)")
                .Columns.Add("AMT_INV", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_INV)")
                '.Columns.Add("QTY_VAR", GetType(System.Int32), "QTY_INV - QTY_REC")
                .Columns.Add("QTY_VAR", GetType(System.Int32), "SUM(Child(APTINVH5).QTY_VAR)")
                '.Columns.Add("AMT_VAR", GetType(System.Decimal), "AMT_INV - AMT_REC")
                .Columns.Add("AMT_VAR", GetType(System.Decimal), "SUM(Child(APTINVH5).AMT_VAR)")
            End With

            Dim sql As String = "Select ICTIREC1.* from ICTIREC1 where ROWNUM < 1"
            ICTIREC1 = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIREC1 & " Add Primary Key (RECEIPT_NO)")
            sql = "Select ICTIREC2.RECEIPT_NO, QTY_REC, QTY_REC * AP_COST AMT_REC from ICTIREC2 where ROWNUM < 1"
            ICTIREC1_TOTALS = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & ICTIREC1_TOTALS & " Add Primary Key (RECEIPT_NO)")

            'ASCMAIN1.sql = "Select ICTIREC1.* from " & ICTIREC1 & " ICTIREC1"
            ASCMAIN1.sql = "Select ICTIREC1.*, POTSHIP2.CONTAINER_NO, POTSHIP2.BOL_NO, POTSHIP2.COMM_INV_NO" & vbCrLf _
                                    & " from " & ICTIREC1 & " ICTIREC1,POTSHIP2" & vbCrLf _
                                    & " where POTSHIP2.PO_SHIPMENT_NO = ICTIREC1.PO_SHIPMENT_NO" & vbCrLf _
                                    & "   and POTSHIP2.PO_SHIPMENT_LNO = ICTIREC1.PO_SHIPMENT_LNO"
            Create_TDA(.Tables.Add, "ICTIREC1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTIREC2.*, ICTSTYL1.STYLE_DESC " _
                & " from ICTIREC2,ICTSTYL1 " _
                & " where ICTIREC2.STYLE_CODE = ICTSTYL1.STYLE_CODE " _
                & "   and ICTIREC2.RECEIPT_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIREC2", "**", 0, False, "V", 2)

            .Relations.Add("ICTIREC2",
            New DataColumn() { .Tables("ICTIREC1").Columns("RECEIPT_NO")},
            New DataColumn() { .Tables("ICTIREC2").Columns("RECEIPT_NO")})

            sql = "Select POTSHIP3.*,POTORDR1.VEND_CODE, 0 RECEIPT_LNO from POTSHIP3,POTORDR1 where ROWNUM < 1"
            POTSHIP3 = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & POTSHIP3 & " Add Primary Key (PO_SHIPMENT_NO,PO_SHIPMENT_LNO,PO_ORDER_NO,PO_ORDER_LNO)")

            sql = "Select POTSHIP2.* from POTSHIP2 where ROWNUM < 1"
            POTSHIP2 = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & POTSHIP2 & " Add Primary Key (PO_SHIPMENT_NO,PO_SHIPMENT_LNO)")


            ASCMAIN1.sql = "Select GLTPARM2.* " _
                & " from GLTPARM2 " _
                & " where OPS_YYYYPP = " _
                & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
                & "  where GLTPARM2.PRD_END_DATE >= :PARM1)"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO, APTINVH1.INV_DATE, APTINVH1.INV_AMT, APTINVH1.INV_RECUR_GEN " _
                & " from APTINVH1 " _
                & " where VOUCHER_NO_RECUR = :PARM1"
            Create_TDA(.Tables.Add, "APTINVHR", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "POTCATG1", "*", 0)
        End With

        Fill_Records("POTCATG1")

        grdAPTINVHB.Dock = DockStyle.Fill
        grdAPTINVR1.Dock = DockStyle.Fill
        tabMain.Dock = DockStyle.Fill

        'ASCMAIN1.sql = "Update ICTIREC1 set ACCRUAL_STATUS = '1', VOUCHER_NO = :PARM2 where RECEIPT_NO = :PARM1"
        'Create_Update_Command("ICTIREC1", "VV")

        grdAPTINVH2.DataSource = dst.Tables("APTINVH2")
        Dim dvw As DataView = DirectCast(grdAPTINVH2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "INV_LTYP is Null or INV_LTYP = 'A'"

        grdICTIREC1.DataSource = dst.Tables("ICTIREC1")
        grdAPTINVH5.DataSource = dst.Tables("APTINVH5")
        grdAPTINVR1.DataSource = dst.Tables("APTINVR1")
        grdAPTINVH5_SUM.DataSource = dst.Tables("APTINVH5_SUM")
        grdASTAUDTX.DataSource = dst.Tables("ASTAUDTX")
        grdAPTINVHX.DataSource = dst.Tables("APTINVHX")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdAPTINVHR.DataSource = dst.Tables("APTINVHR")
        grdAPTINVHB.DataSource = dst.Tables("APTINVHB")
        grdAPTINVHM.DataSource = dst.Tables("APTINVHM")
        grdAPTINVH8.DataSource = dst.Tables("APTINVH8")
        grdAPTINVH7.DataSource = dst.Tables("APTINVH7")
        grdPOTLCST1.DataSource = dst.Tables("POTLCST1")

        Set_SEGS(grdAPTINVH2, "APTINVH2")
        Set_SEGS(grdAPTINVR1, "APTINVR2")

        Bind_Controls(Me, "APTVEND1")
        Bind_Controls(Me, "APTVEND2")
        Bind_Controls(Me, "APTINVH1")

        Create_Summary(grdAPTINVH2, "VOUCHER_LNO", "Count")
        Create_Summary(grdAPTINVH2, "INV_LINE_AMT")

        Create_Summary(grdAPTINVH8, "VOUCHER_ANO", "Count")
        Create_Summary(grdAPTINVH8, "VOUCHER_ADJ_AMT")

        Create_Summary(grdAPTINVH7, "VOUCHER_CLNO", "Count")
        Create_Summary(grdAPTINVH7, "TOTAL_INV")

        Create_Summary(grdPOTLCST1, "CTL_NO", "Count")
        Create_Summary(grdPOTLCST1, New String() {"COST_ACC", "COST_ACT"})

        Create_Summary(grdAPTINVR1, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVR1, "INV_AMT")

        Create_Summary(grdAPTINVHB, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHB, "INV_AMT")

        Create_Summary(grdAPTINVHM, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHM, "INV_AMT")

        Create_Summary(grdAPTINVHR, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHR, "INV_AMT")

        Create_Summary(grdICTIREC1, "RECEIPT_NO", "Count", "ICTIREC1")
        Create_Summary(grdICTIREC1, "QTY_REC", , "ICTIREC1")
        Create_Summary(grdICTIREC1, "AMT_REC", , "ICTIREC1")

        Create_Summary(grdAPTINVH5_SUM, "RECEIPT_NO", "Count")
        Create_Summary(grdAPTINVH5_SUM, New String() {"QTY_REC", "AMT_REC", "QTY_INV", "AMT_INV", "QTY_VAR", "AMT_VAR"})

        Create_Summary(grdAPTINVH5, "VOUCHER_DLNO", "Count")
        Create_Summary(grdAPTINVH5, New String() {"QTY_REC", "AMT_REC", "INV_QTY", "AMT_INV", "QTY_VAR", "AMT_VAR"})

        Create_Summary(grdAPTINVHX, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHX, New String() {"SELECTED", "INV_BALANCE", "INV_DISC_AMT", "INV_PAYMENTS"})

        With grdAPTINVH5_SUM.DisplayLayout.Bands("APTINVH5_SUM")
            .Columns("QTY_INV").CellAppearance.BackColor = Drawing.Color.Bisque
            .Columns("AMT_INV").CellAppearance.BackColor = Drawing.Color.Bisque

            For Each C As String In New String() {"RECEIPT_NO", "RECEIPT_DATE", "QTY_REC", "AMT_REC", "QTY_INV", "AMT_INV", "QTY_VAR", "AMT_VAR"}
                .Columns(C).Header.Fixed = True
            Next
        End With

        With grdAPTINVH5.DisplayLayout.Bands(0)
            .Columns("INV_QTY").CellAppearance.BackColor = Drawing.Color.Bisque
            .Columns("AMT_INV").CellAppearance.BackColor = Drawing.Color.Bisque
        End With

        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").SummaryFooterCaption = "Voucher Totals"
        grdAPTINVH5_SUM.DisplayLayout.Bands("APTINVH5_SUM").SummaryFooterCaption = "Totals for All Receipts in the Voucher"
        grdICTIREC1.DisplayLayout.Bands("ICTIREC1").SummaryFooterCaption = "Totals for All Receipts Not Vouchered"

        Set_Read_Only(grpAPTVEND1, True)
        Set_Read_Only(grpAPTVEND2, True)
        Set_Read_Only(grpOtherVendorInfo, True)

        Absx1.txtFor("VEND_ALT_CODE").ReadOnly = False Or InquiryMode

        grpShowDistribution.Visible = InquiryMode
        tabReceipts.Tabs("Open Accrued PO Receipts").Visible = Not InquiryMode

        With grdAPTINVR1.DisplayLayout.Bands("APTINVR1")
            For i As Integer = 0 To .Columns.Count - 1
                If .Columns(i).Key = .Columns(i).Header.Caption Then
                    .Columns(i).Hidden = True
                End If
            Next
        End With

        Sort_grdColumns(grdAPTINVH8, "VOUCHER_ANO")
        Sort_grdColumns(grdAPTINVH7, "VOUCHER_CLNO")
        Sort_grdColumns(grdPOTLCST1, "CTL_NO")

        grdAPTINVHB.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdAPTINVHB.DisplayLayout.Bands(0).SortedColumns.Add("VOUCHER_NO", False)

        grdAPTINVHM.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdAPTINVHM.DisplayLayout.Bands(0).SortedColumns.Add("VOUCHER_NO", False)

        ' THIS SHOULD ALL HAPPEN IN A CALL 
        With grdAPTINVR1.DisplayLayout.Bands("APTINVR1")
            .Columns("VOUCHER_NO").Header.Fixed = True
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NUM").Header.Fixed = True
        End With

        With grdAPTINVH2.DisplayLayout.Bands(0)
            .Columns("VOUCHER_LNO").Header.Fixed = True
        End With

        grdAPTINVHX.DisplayLayout.UseFixedHeaders = True
        With grdAPTINVHX.DisplayLayout.Bands("APTINVHX")
            .Columns("SELECTED").Header.Fixed = True
        End With

        With grdAPTINVHB.DisplayLayout.Bands("APTINVHB")
            .Columns("VOUCHER_NO").Header.Fixed = True
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VEND_NAME").Header.Fixed = True
        End With

        With grdAPTINVHM.DisplayLayout.Bands("APTINVHM")
            .Columns("VOUCHER_NO").Header.Fixed = True
            .Columns("INV_NUM").Header.Fixed = True
            .Columns("INV_DATE").Header.Fixed = True
            .Columns("INV_AMT").Header.Fixed = True
        End With

        If ROWs("APTPARM1").Item("AP_PARM_ALLOW_ACCRUAL") & "" <> "1" Then
            chkACCRUE_PRIOR.Visible = False
        End If

        For Each gcol As UltraWinGrid.UltraGridColumn In grdAPTINVH7.DisplayLayout.Bands(0).Columns
            If gcol.Key = "TOTAL_INV" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                gcol.CellAppearance.BackColor = Drawing.Color.Yellow
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        If InquiryMode Then
            optINV_STATUS.ValueList.ValueListItems.Add("D", "Deleted")

            With grdAPTINVH2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With

            With grdAPTINVH7.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With

            Set_Read_Only(tabMain, True)

            grdAPTINVHX.Visible = False
            cmdNextCheckNo.Visible = False
        End If

        'tabMain.Tabs("PO Receipts").Visible = False
        'tabMain.Tabs("Other Accruals").Visible = False

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            lblBOL_NO.Text = "FCR No"
            grdAPTINVH7.DisplayLayout.Bands(0).Columns("BOL_NO").Header.Caption = "FCR No"
            grdPOTLCST1.DisplayLayout.Bands(0).Columns("BOL_NO").Header.Caption = "FCR No"
        End If

        chkQuickEntry.Checked = Not InquiryMode


        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        tab0.Tabs("PO Receipts").Visible = Not InquiryMode

        If ASCMAIN1.CLIENT = "VAN" Then
            With grdICTIREC1.DisplayLayout.Bands(0)
                .Columns("PO_SHIPMENT_NO").Header.SetVisiblePosition(2, False)
                '.Columns("PO_SHIPMENT_NO").Header.SetVisiblePosition = 2
            End With
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("VEND_CODE")

                If Absx1.txtFor("INV_NUM").Text = "" Then
                    EMsg &= vbCr & "Invoice No Required"
                End If

                If Absx1.cbeFor("INV_TYPE").Text = "" Then
                    EMsg &= vbCr & "Invoice Type Required"
                End If

                If EMsg = "" Then
                    If Check_Invoice() <> "YES" Then
                        Exit Sub
                    End If
                    If Not ASCMAIN1.Logical_Lock("APTVEND1", Absx1.txtFor("VEND_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Edit", "Load"

                If Validate_Code("VOUCHER_NO") Then
                    If Not InquiryMode Then
                        If cdr.Item("INV_STATUS") & "" = "D" Then
                            EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " has been Deleted"
                        End If
                        If cdr.Item("INV_STATUS") & "" = "P" Then
                            EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " has been Paid"
                        End If
                        If cdr.Item("BATCH_NO_PYMT") & "" <> "" Then
                            EMsg &= vbCr & "Voucher " & Absx1.txtFor("VOUCHER_NO").Text & " has been Selected for Payment in Batch " & cdr.Item("BATCH_NO_PYMT")
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If Not InquiryMode Then
                        If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("VOUCHER_NO").Text) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("APTVEND1", cdr.Item("VEND_CODE") & "") Then Exit Sub
                    End If
                End If


            Case "Multi-Invoice Edit"
                Validate_Code("VEND_CODE")

                If ASCMAIN1.USER_ID <> "mattinam" And ASCMAIN1.USER_ID <> "wjzz" Then
                    EMsg &= vbCr & "This Option is NOT Ready yet (See Maria/Walter)"
                End If

            Case "Update"
                If EntryMode = "B" Then
                    ' validate things
                ElseIf EntryMode = "M" Then
                    ' validate things
                Else

                    If Absx1.txtFor("INV_NUM").Text = "" Then
                        EMsg &= vbCr & "You Must Enter an Invoice No"
                    End If
                    If Absx1.dteFor("INV_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "You Must Enter a Document Date"
                    End If
                    If Absx1.dteFor("INV_DUE_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "You Must Enter a Valid Terms Code and Document Date so that a Due Date may be calculated"
                    End If

                    Dim DIST_OOBAL As Decimal = Val(Absx1.numFor("DIST_OOBAL").Value & "")
                    If System.Math.Round(DIST_OOBAL, 2) <> 0 Then
                        EMsg &= vbCr & "Distribution is Out of Balance"
                    End If

                    Select Case Absx1.cbeFor("INV_TYPE").Value
                        Case "I", "D"
                            Dim INV_AMT As Decimal = Val(Absx1.numFor("INV_AMT").Value & "")
                            Dim INV_AMT_VEND As Decimal = Val(Absx1.numFor("INV_AMT_VEND").Value & "")
                            If Val(Absx1.numFor("INV_AMT").Value & "") < 0 _
                            Or Val(Absx1.numFor("INV_AMT_VEND").Value & "") < 0 Then
                                EMsg &= vbCr & "Amount must be Postive (i.e., an amount owed TO the Vendor) for this type of Document"
                            End If
                        Case "C", "R"
                            If Val(Absx1.numFor("INV_AMT").Value & "") > 0 _
                            Or Val(Absx1.numFor("INV_AMT_VEND").Value & "") > 0 Then
                                EMsg &= vbCr & "Amount must be Negative (i.e., an amount owed FROM the Vendor) for this type of Document"
                            End If
                        Case "A"
                            If EntryMode = "N" Then
                                If Val(Absx1.numFor("INV_AMT").Value & "") < 0 _
                                Or Val(Absx1.numFor("INV_AMT_VEND").Value & "") < 0 Then
                                    EMsg &= vbCr & "Amount must be Postive (i.e., an amount owed TO the Vendor) for this type of Document"
                                    EMsg &= vbCr & "- An offsetting document with the same (but Negative) Amount will be automatically created"
                                End If
                            End If
                    End Select

                    If Absx1.optFor("INV_REMIT_TO").Value = "N" Then
                        If Absx1.txtFor("VEND_ALT_CODE").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_NAME").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_CITY").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_STATE").Text = "" _
                        Or Absx1.txtFor("VEND_ALT_ZIP_CODE").Text = "" Then
                            EMsg &= vbCr & "Address Code, Vendor Name, City, State & Zip are Required for a New Payment Address"
                        Else
                            LookUp("APTVEND2", New String() {HFs("VEND_CODE"), Absx1.txtFor("VEND_ALT_CODE").Text})
                            If cdr IsNot Nothing Then
                                EMsg &= vbCr & "Vendor Address Code " & Absx1.txtFor("VEND_ALT_CODE").Text & " is Already on File"
                            End If
                        End If
                    Else
                        If Absx1.optFor("INV_REMIT_TO").Value = "V" _
                        Or Absx1.optFor("INV_REMIT_TO").Value = "P" Then
                            If Absx1.txtFor("VEND_ALT_CODE").Text <> "" Then
                                Absx1.txtFor("VEND_ALT_CODE").Text = ""
                            End If
                        End If
                        If Absx1.optFor("INV_REMIT_TO").Value = "A" Then
                            If Absx1.txtFor("VEND_ALT_CODE").Text = "" Then
                                EMsg &= vbCr & "You Must Specify a Valid (Alternate) Payment Address Code"
                            Else
                                LookUp("APTVEND2", New String() {HFs("VEND_CODE"), Absx1.txtFor("VEND_ALT_CODE").Text})
                                If cdr Is Nothing Then
                                    EMsg &= vbCr & "Vendor Address Code " & Absx1.txtFor("VEND_ALT_CODE").Text & " is not on File"
                                End If
                            End If
                        End If
                    End If

                    If rowAPTINVH1("INV_STATUS") = "R" Then
                        If Absx1.cbeFor("INV_TYPE").Value <> "I" Then
                            EMsg &= vbCr & "Recurring Feature applies to Invoices Only"
                        End If
                        If Val(Absx1.numFor("INV_AMT").Value & "") <= 0 Then
                            EMsg &= vbCr & "Recurring Invoice Template must have a positive, non-zero Amount"
                        Else
                            If Val(Absx1.numFor("INV_AMT_VEND").Value & "") <= 0 Then
                                EMsg &= vbCr & "Recurring Invoice Template must have a positive, non-zero Amount"
                            End If
                        End If
                        If Absx1.optFor("INV_RECUR_CYCLE").Value & "" = "" Then
                            EMsg &= vbCr & "Recurring Invoice Template must have an Recurring Cycle"
                        End If

                        If dst.Tables("APTINVH5").Rows.Count <> 0 Then
                            EMsg &= vbCr & "Recurring Invoice Templates may not have entries for Accrued Purchases"
                        End If
                        If dst.Tables("APTINVH8").Rows.Count <> 0 Then
                            EMsg &= vbCr & "Recurring Invoice Templates may not have entries for Adjustments"
                        End If
                        If dst.Tables("APTINVH7").Rows.Count <> 0 Then
                            EMsg &= vbCr & "Recurring Invoice Templates may not have entries for Other Accruals"
                        End If

                        If Absx1.txtFor("INV_RECUR_OPS_YYYYPP_BEGIN").Text = "" Then
                            EMsg &= vbCr & "Recurring Invoice Template must have a Starting Period"
                        Else
                            If Not Validate_Code("INV_RECUR_OPS_YYYYPP_BEGIN") Then
                                EMsg &= vbCr & "Invalid Starting Period"
                            End If
                        End If
                    End If


                    Dim INV_PYMT_METHOD As String = rowAPTINVH1.Item("INV_PYMT_METHOD") & ""

                    If INV_PYMT_METHOD = "LC" Then
                        If rowAPTINVH1("INV_STATUS") <> "P" Then
                            EMsg &= vbCr & "Cannot Leave an Invoice Unpaid when Choosing to pay by LC"
                        End If
                        Dim LC_CTL_NO As String = rowAPTINVH1.Item("LC_CTL_NO") & ""
                        If LC_CTL_NO = "" Then
                            EMsg &= vbCr & "Cannot Leave the LC Fields blank when Choosing to pay by LC"
                        Else
                            If LookUp("POTLTRC1", LC_CTL_NO) Is Nothing Then
                                EMsg &= vbCr & "Cannot find a record of LC Ctl No " & LC_CTL_NO
                            Else
                                If cdr.Item("VEND_CODE") & "" <> rowAPTINVH1.Item("VEND_CODE") & "" Then
                                    EMsg &= vbCr & "Cannot use an LC from a different supplier (" & cdr.Item("VEND_CODE") & ") than the Vendor defined to the invoice (" & rowAPTINVH1.Item("VEND_CODE") & ")"
                                End If
                            End If
                        End If

                        If dst.Tables("APTINVH5").Select("").Length = 0 Then
                            EMsg &= vbCr & "Cannot use LC Payment Method without involving a Purchase Shipment or Receipt"
                        End If
                    End If


                        If rowAPTINVH1("INV_STATUS") = "P" Then
                        If Absx1.txtFor("BANK_CODE").Text = "" Then
                            EMsg &= vbCr & "You Must Specify a Bank Code to Pay upon Entry"
                        End If
                        If Absx1.txtFor("INV_PYMT_METHOD").Text = "" Then
                            EMsg &= vbCr & "You Must Specify a Payment Method to Pay upon Entry"
                        End If
                        If Absx1.txtFor("CHECK_NUM").Text = "" Or Absx1.dteFor("CHECK_DATE").Value & "" = "" Then
                            If batch_update Then
                                ' generate the check number
                            End If
                            EMsg &= vbCr & "You Must Specify a Check Number and Date to Pay upon Entry"
                        End If
                        If EMsg = "" Then
                            Dim rowAPTCHCK1 As DataRow = LookUp("APTCHCK1", New String() {Absx1.txtFor("BANK_CODE").Text, Absx1.txtFor("CHECK_NUM").Text})
                            If rowAPTCHCK1 IsNot Nothing Then
                                EMsg &= vbCr & "Check No " & Absx1.txtFor("CHECK_NUM").Text & " has already been Posted"
                            End If
                        End If
                        'If Val(Absx1.numFor("CHECK_AMT").Value) = 0 And (Absx1.txtFor("BANK_CODE").Text <> "Z") Then
                        '    EMsg &= vbCr & "You Must Use Bank Code Z for Zero Checks"
                        'End If
                        'If Val(Absx1.numFor("CHECK_AMT").Value) <> 0 And (Absx1.txtFor("BANK_CODE").Text = "Z") Then
                        '    EMsg &= vbCr & "You May NOT Use Bank Code Z for Non-Zero Checks"
                        'End If
                    Else
                        If EMsg = "" Then
                            Absx1.txtFor("CHECK_NUM").Text = ""
                            Absx1.dteFor("CHECK_DATE").Value = ""
                            If Val(Absx1.numFor("INV_AMT").Value & "") = 0 Then
                                If vbNo = MsgBox("You have not clicked 'Paid'," _
                                                 & vbCr & "  which means this Invoice will be updated as 'Open'" _
                                                 & vbCr & "  and will need to be Selected for Payment" & vbCr _
                                                 & "  to remove it from the Open AP Items Report" _
                                                 & vbCr & vbCr & "Continue Anyway?", vbQuestion + vbYesNo,
                                                 "Normally, a $0 Invoice is entered as 'Paid' on a $0 Check") Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                    Validate_Code("POST_CODE")
                    Validate_Code("TERM_CODE")
                    Validate_Code("INV_PYMT_METHOD")
                    Validate_Code("BANK_CODE", , True)
                    If Absx1.txtFor("BANK_CODE").Text <> "" Then
                        Dim row As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                        If row IsNot Nothing AndAlso row.Item("BANK_STATUS") <> "A" Then
                            EMsg &= vbCr & "Bank " & Absx1.txtFor("BANK_CODE").Text & " is Not Active"
                        End If
                    End If
                    Validate_Code("INV_PYMT_CYCLE", , True)
                    Validate_Code("CURR_CODE")
                    Validate_Code("REASON_CODE", , True) ' TRUE/FALSE SHOULD BE RELATED TO MEMO

                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                        (dst.Tables("APTINVH2").Select("INV_LTYP is Null and DIST_CODE is Not Null"),
                         New String() {"DIST_CODE"}).Select("")
                        Dim DIST_CODE As String = row.Item("DIST_CODE")
                        If LookUp("GLTDIST1", DIST_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Distribution Code " & DIST_CODE
                        End If
                    Next
                    For Each row As DataRow In dst.Tables("APTINVH2").Select("")
                        Dim ACCT_CODE As String = row.Item("ACCT_CODE") & ""
                        If LookUp("GLTACCT1", ACCT_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Account Code " & ACCT_CODE
                        Else
                            If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                                EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is not Active"
                            End If
                            If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                                EMsg &= vbCr & "Acct Code " & ACCT_CODE & " is a Control Account - no Manual J/E permitted"
                            End If
                        End If
                    Next

                    If EMsg = "" Then
                        If Val(Absx1.numFor("INV_AMT_VEND").Value & "") <> Val(Absx1.numFor("INV_AMT").Value & "") Then
                            If MsgBox("Please Verify the Following Information:" & vbCr & vbCr & "Vendor Invoice Amount: " & Format(Val(Absx1.numFor("INV_AMT_VEND").Value & ""), "#,##0.00") & vbCr & "Invoice Payable: " & Format(Val(Absx1.numFor("INV_AMT").Value & ""), "#,##0.00") & vbCr & vbCr & "OK To Continue with Update?", vbQuestion + vbYesNo, "Verification: Invoice will be Booked with Adjustments") = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If

                    If EMsg = "" Then
                        If chkINV_1099_IND.Checked And Val(Absx1.numFor("INV_1099_AMT").Value & "") = 0 Then
                            If MsgBox("1099 Amount Option is checked, but there is no 1099 Amount Entered." _
                                       & vbCrLf & vbCrLf & "Proceed with Update?",
                                       MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
                        End If
                    End If




                    If EMsg = "" And rowAPTINVH1("INV_STATUS") = "P" Then
                        Dim OTHERS As Decimal = Val(dst.Tables("APTINVHX").Compute("SUM(INV_PAYMENTS)", "SELECTED = '1'") & "")
                        Dim CHECK_AMT As Decimal = Val(Absx1.numFor("CHECK_AMT").Value & "")
                        Dim INV_PAYMENTS As Decimal = Val(Absx1.numFor("INV_AMT").Value & "") - Val(Absx1.numFor("INV_DISC_AMT").Value & "") + Val(Absx1.numFor("LC_FEE").Value & "")
                        If CHECK_AMT <> OTHERS + INV_PAYMENTS Then
                            EMsg = EMsg & vbCr & "The payment amount for this Invoice (" & Format(INV_PAYMENTS, "$##,##0.00") & ")" & vbCr & " plus Selected Other AP Items (" & Format(OTHERS, "$##,##0.00") & ")" & vbCr & " does not agree with Check Amount Specified (" & Format(CHECK_AMT, "$##,##0.00") & ")"
                        End If
                        If CHECK_AMT < 0 Then
                            EMsg = EMsg & vbCr & "This Invoice plus Selected Other AP Items has a Net Negative Balance" & vbCr & " - Negative Payment Not Permitted"
                        End If

                        Dim Dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
                        Dim CHECK_DATE As String = Format(Absx1.dteFor("CHECK_DATE").Value, "yyyyMMdd")
                        If CHECK_DATE > Format(Dates(Dates.Length - 1), "yyyyMMdd") Or CHECK_DATE < Format(Dates(1), "yyyyMMdd") Then
                            If MsgBox("Please Verify the Check Date used:" & Absx1.dteFor("CHECK_DATE").Value & vbCr & "Current Period Date Range is: " & Dates(1) & " thru " & Dates(Dates.Length - 1) & vbCr & vbCr & "OK To Continue with Update?", vbExclamation + vbYesNo, "Verification: You have entered a Date out of Range") = vbNo Then
                                Exit Sub
                            End If
                        End If


                        If EMsg = "" And Not batch_update Then
                            ' PAYEENAME IS WRONG
                            If MsgBox("Please Verify the Following Information:" & vbCr & vbCr & "Check No: " & Absx1.txtFor("CHECK_NUM").Text & ", " & Absx1.dteFor("CHECK_DATE").Value & vbCr & "Bank: " & Absx1.txtFor("BANK_DESC").Text & vbCr & "Payee: " & Absx1.txtFor("VEND_NAME").Text & vbCr & "Amount: " & Format(Val(Absx1.numFor("CHECK_AMT").Value & ""), "$###,##0.00") & IIf(Val(Absx1.numFor("INV_AMT").Value & "") = 0, vbCr & vbCr & "*** This Invoice has a $0 Balance ***", "") & vbCr & vbCr & "OK To Continue with Update?", vbQuestion + vbYesNo, "Verification: You are about to Record a Payment") = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If

                    'If EntryMode = "E" Then
                    '    If Absx1.optFor("INV_STATUS").Value = "R" _
                    '    And rowAPTINVH1("INV_STATUS", DataRowVersion.Original) <> "R" Then
                    '        EMsg &= vbCr & "Cannot Change an invoice into a Recurring Invoice Template"
                    '    End If
                    '    If Absx1.optFor("INV_STATUS").Value <> "R" _
                    '    And rowAPTINVH1("INV_STATUS", DataRowVersion.Original) = "R" Then
                    '        EMsg &= vbCr & "Cannot Use a Recurring Invoice Template for an Actual Invoice Posting"
                    '    End If
                    'End If

                    If chkACCRUE_PRIOR.Checked Then
                        'If Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Value & "" = "" Then
                        If Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text & "" = "" Then
                            EMsg &= vbCr & "Missing Accrual Period"
                        End If
                    End If

                    If EMsg = "" Then
                        If Val(Absx1.numFor("INV_AMT").Value & "") = 0 Then
                            If MsgBox("Proceed with Entry?", MsgBoxStyle.YesNo, "Invoice Payable Amount is Zero") <> MsgBoxResult.Yes Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

        End Select

        If EMsg <> "" Then
            If Not batch_update Then
                MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            End If
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

                If chkQuickEntry.Checked Then
                    Absx1.txtFor("INV_REF").Text = Absx1.txtFor("QE_INV_REF").Text
                    Absx1.dteFor("INV_DATE").Value = Absx1.dteFor("QE_INV_DATE").Value
                    Absx1.numFor("INV_AMT_VEND").Value = Absx1.numFor("QE_INV_AMT").Value
                    'Absx1.numFor("INV_AMT").Value = Absx1.numFor("QE_INV_AMT").Value

                    Calculate_INV_DUE_DATE()
                    If dst.Tables("APTINVH2").Rows.Count = 0 And Val(numINV_AMT.Value & "") <> 0 Then
                        Generate_Pre_Distribution()
                        Calc_Totals()
                    End If
                    tabMain.SelectedTab = tabMain.Tabs(2)

                End If

            Case "New Batch"
                EntryMode = "B"
                Prepare_for_Batch_Entry()
                Mode_Settings(True)

            Case "Multi-Invoice Edit"
                EntryMode = "M"
                Prepare_for_Multi_Invoice_Edit()
                Mode_Settings(True)

            Case "Edit", "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                If EntryMode = "B" Then
                    Update_Batch()
                ElseIf EntryMode = "M" Then
                    Update_Multi()
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Print Edit"
                Print_Record()

            Case "Export to Excel"
                If EntryMode = "B" Then
                    Export_to_Excel(grdAPTINVHB)
                ElseIf EntryMode = "M" Then
                    Export_to_Excel(grdAPTINVHM)
                End If

            Case "Import from Excel"
                Import_Batch_from_Excel()

            Case "All Vouchers"
                Copy_Change_to("All")

            Case "Selected Vouchers"
                Copy_Change_to("Sel")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If InquiryMode Then
                        .Items("Load").Settings.Enabled = not_iScreenMode
                        .Items("Done").Settings.Enabled = iScreenMode
                    Else
                        .Items("New").Settings.Enabled = not_iScreenMode
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode

                        If EntryMode = "B" Then
                            .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                        ElseIf EntryMode = "M" Then
                            .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                        Else
                            .Items("Delete").Settings.Enabled = iScreenMode
                        End If

                        If EntryMode = "N" Then
                            .Items("Delete").Visible = False
                        Else
                            .Items("Delete").Visible = Not InquiryMode
                        End If

                        .Items("Print Edit").Settings.Enabled = not_iScreenMode
                        .Items("New Batch").Settings.Enabled = not_iScreenMode
                        .Items("Multi-Invoice Edit").Settings.Enabled = not_iScreenMode
                    End If

                    'WJZ DEMO
                    .Items("Print Edit").Visible = False
                    .Items("New Batch").Visible = False
                    .Items("Multi-Invoice Edit").Visible = False
                End With

                If InquiryMode Then
                    .Groups("Distribution Options").Visible = False
                    .Groups("Generate $0 Accrual").Visible = False
                    .Groups("Entry Options").Visible = False
                    .Groups("Copy Last Change to ...").Visible = False
                    .Groups("Batch / Excel Options").Visible = False
                Else
                    .Groups("Copy Last Change to ...").Visible = False
                    .Groups("Distribution Options").Visible = False
                    .Groups("Generate $0 Accrual").Visible = False
                    Setup_tab0()
                    '.Groups("Entry Options").Visible = Not tf

                    If EntryMode = "B" Then
                        .Groups("Batch / Excel Options").Visible = tf
                    ElseIf EntryMode = "M" Then
                        .Groups("Copy Last Change to ...").Visible = tf
                    Else
                        .Groups("Batch / Excel Options").Visible = False
                    End If
                End If
                .Groups("PO Receipts").Visible = False
            End With
        End If

        'tabMain.Tabs("PO Receipts").Visible = ScreenMode AndAlso (dst.Tables("ICTIREC1").Rows.Count <> 0)

        'WJZ DEMO
        chkQuickEntry.Checked = False
        chkQuickEntry.Visible = False
        chkRecurring.Visible = False

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpLastChange.Visible = False

        lblDIST_OOBAL.Visible = Not InquiryMode
        numDIST_OOBAL.Visible = Not InquiryMode

        If InquiryMode Then
            Absx1.txtFor("INV_NUM").ReadOnly = True
            Absx1.txtFor("VEND_ALT_CODE").ButtonsRight(0).Enabled = False
            Absx1.txtFor("VEND_ALT_CODE").ReadOnly = True

        ElseIf EntryMode = "Z" Then
            Absx1.txtFor("INV_NUM").ReadOnly = True
            Absx1.txtFor("VEND_ALT_CODE").ButtonsRight(0).Enabled = False
            Absx1.txtFor("VEND_ALT_CODE").Enabled = False
        Else
            Absx1.txtFor("INV_NUM").ReadOnly = False
            Absx1.txtFor("VEND_ALT_CODE").ButtonsRight(0).Enabled = True
            Absx1.txtFor("VEND_ALT_CODE").Enabled = True
        End If

        'grdAPTINVR1.Visible = Not ScreenMode
        tab0.Visible = Not ScreenMode

        tabMain.Visible = tf
        grdAPTINVHB.Visible = False
        grdAPTINVHM.Visible = False



        If ScreenMode And Not InquiryMode Then
            If EntryMode = "N" Then
                chkINV_1099_IND.Checked = (rowAPTVEND1.Item("VEND_TAX_ID") & "" <> "")
            Else
                chkINV_1099_IND.Checked = (Val(rowAPTINVH1.Item("INV_1099_AMT") & "") <> 0)
            End If
            Setup_OPS_YYYYPP_ACCRUE()
        End If

        If ScreenMode Then
            grdICTIREC1.Parent = tabReceipts.Tabs("Open Accrued PO Receipts").TabPage
            With grdICTIREC1.DisplayLayout.Bands(0)
                .Columns("VEND_CODE").Hidden = True
                .Columns("SEL").Hidden = InquiryMode Or Not (EntryMode = "N" Or EntryMode = "E")
            End With

            If EntryMode = "B" Then
                tabMain.Visible = False
                grdAPTINVHB.Visible = True
                UltraGroupBox1.Visible = False
            End If

            If EntryMode = "M" Then
                tabMain.Visible = False
                grdAPTINVHM.Visible = True
                lblCOLUMN_NAME.Visible = False
                lblNEW_VALUE.Visible = False
                grpLastChange.Visible = True
            End If

        Else
            grdICTIREC1.Parent = tab0.Tabs("PO Receipts").TabPage
            With grdICTIREC1.DisplayLayout.Bands(0)
                .Columns("VEND_CODE").Hidden = False
                .Columns("SEL").Hidden = True
            End With

            tab0.SelectedTab = tab0.Tabs("Invoices Entered")

            grpLC.Visible = False

            UltraGroupBox1.Visible = True
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"APTINVH1", "APTINVH2", "APTINVH5", "APTINVH5_SUM", "APTINVH7", "APTINVH8", "ICTIREC1", "ICTIREC2", "POTLCST1", "APTCHCK1", "APTCHCK2",
             "TATEVNT1", "APTINVHX", "ASTAUDTX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If Not batch_update Then
            dst.Tables("APTINVHB").Rows.Clear()
        End If

        EnforceConstraints(True)

        If HFs.ContainsKey("VEND_CODE") AndAlso HFs("VEND_CODE") <> "" Then
            Absx1.txtFor("VEND_CODE").Text = HFs("VEND_CODE")
        End If

        disable_calculate_totals = False

        'Fill_Records("APTINVR1")
        Show_Batch()

        Absx1.numFor("DIST_GL").Value = 0
        Absx1.numFor("DIST_PO").Value = 0
        Absx1.numFor("DIST_OTHER").Value = 0

        Absx1.cbeFor("INV_TYPE").Value = "I"

        Absx1.txtFor("QE_INV_REF").Text = ""
        Absx1.dteFor("QE_INV_DATE").Value = Null
        Absx1.numFor("QE_INV_AMT").Value = 0

        Absx1.numFor("CHECK_AMT").Value = 0

        Setup_QE(chkQuickEntry.Checked)

        lblRecurringTemplate.Visible = False
        lblRecurring.Visible = False
        tabHeader.SelectedTab = tabHeader.Tabs("Codes")
        tabHeader.Tabs("Pymt Info").Enabled = False

        Clear_Other_Accrual_Controls()

        ASCMAIN1.Progress("", "")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If chkQuickEntry.Checked Then
            Setup_QE(False)
        End If

        If EntryMode = "N" Then
            HFs("VOUCHER_NO") = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
        End If

        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.SuspendBinding()

        disable_calculate_totals = True

        EnforceConstraints(False)

        rowAPTINVH1 = Fill_Record("APTINVH1", New String() {HFs("VOUCHER_NO")}, EntryMode = "N")
        If EntryMode = "E" Then
            HFs("VEND_CODE") = rowAPTINVH1.Item("VEND_CODE")
            HFs("INV_TYPE") = rowAPTINVH1.Item("INV_TYPE")
            'HFs("VEND_NAME") = rowAPTINVH1.Item("VEND_NAME")
            HFs("INV_NUM") = rowAPTINVH1.Item("INV_NUM")
        End If
        rowAPTVEND1 = Fill_Record("APTVEND1", HFs("VEND_CODE"))
        X.ResumeBinding()

        Fill_Records("APTINVH2", New String() {HFs("VOUCHER_NO")})
        Fill_Records("APTINVH5", New String() {HFs("VOUCHER_NO")}) ', HFs("VOUCHER_NO")}) ' THIS LINE WAS REMMED OUT - NOT SURE WHY - SCREWS UP VOUCHER WHEN LOADED UP TO MAKE CHANGES WHEN THIS LINE IS REMMED OUT

        Fill_Records("APTINVH8", New String() {HFs("VOUCHER_NO")})
        Fill_Records("APTINVH7", New String() {HFs("VOUCHER_NO")})
        Fill_Records("POTLCST1", New String() {HFs("VOUCHER_NO")})
        auto_next_check = False
        BANK_LAST_CHECK_NO = ""
        BANK_NEXT_CHECK_NO = ""


        If EntryMode = "N" Then
            rowAPTINVH1.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            rowAPTINVH1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowAPTINVH1.Item("INV_TYPE") = HFs("INV_TYPE")
            rowAPTINVH1.Item("INV_NUM") = HFs("INV_NUM")

            rowAPTINVH1.Item("VEND_CODE_AP") = rowAPTVEND1.Item("VEND_CODE_AP")
            If rowAPTVEND1.Item("VEND_PYMT_ADDR") & "" = "" Then
                rowAPTINVH1.Item("VEND_ALT_CODE") = ""
                rowAPTINVH1.Item("INV_REMIT_TO") = "V"
            Else
                rowAPTINVH1.Item("VEND_ALT_CODE") = rowAPTVEND1.Item("VEND_PYMT_ADDR")
                rowAPTINVH1.Item("INV_REMIT_TO") = "A"
            End If

            ' UNREM THIS WHEN READY
            'If rowAPTVEND1.Item("VEND_CODE_AP") & "" <> "" Then
            '    rowAPTINVH1.Item("INV_REMIT_TO") = "P"
            'Else
            '    If rowAPTINVH1.Item("VEND_ALT_CODE") & "" <> "" Then
            '        rowAPTINVH1.Item("INV_REMIT_TO") = "A"
            '    Else
            '        rowAPTINVH1.Item("INV_REMIT_TO") = "V"
            '    End If
            'End If


            rowAPTINVH1.Item("INV_SEP_CHECK") = rowAPTVEND1.Item("VEND_SEP_CHECKS")
            rowAPTINVH1.Item("TERM_CODE") = rowAPTVEND1.Item("TERM_CODE")

            If rowAPTVEND1.Item("BANK_CODE") & "" = "" Then
                rowAPTINVH1.Item("BANK_CODE") = ROWs("APTPARM1").Item("AP_PARM_BANK_CODE")
            Else
                rowAPTINVH1.Item("BANK_CODE") = rowAPTVEND1.Item("BANK_CODE")
            End If

            If rowAPTVEND1.Item("VEND_PYMT_METHOD") & "" = "" Then
                If rowAPTINVH1.Item("BANK_CODE") & "" <> "" Then
                    Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowAPTINVH1.Item("BANK_CODE"))
                    If rowGLTBANK1 IsNot Nothing Then
                        rowAPTINVH1.Item("INV_PYMT_METHOD") = rowGLTBANK1.Item("BANK_PYMT_METHOD")
                    End If
                End If
            Else
                rowAPTINVH1.Item("INV_PYMT_METHOD") = rowAPTVEND1.Item("VEND_PYMT_METHOD")
            End If


            rowAPTINVH1.Item("INV_PYMT_CYCLE") = rowAPTVEND1.Item("VEND_PYMT_CYCLE")
            If rowAPTVEND1.Item("POST_CODE") & "" <> "" Then
                rowAPTINVH1.Item("POST_CODE") = rowAPTVEND1.Item("POST_CODE")
            Else
                rowAPTINVH1.Item("POST_CODE") = ROWs("APTPARM1").Item("AP_PARM_POST_CODE")
            End If
            rowAPTINVH1.Item("CURR_CODE") = rowAPTVEND1.Item("CURR_CODE")
            If rowAPTINVH1.Item("CURR_CODE") & "" = "" Then
                rowAPTINVH1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            End If
            If rowAPTINVH1.Item("CURR_CODE") & "" = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
                rowAPTINVH1.Item("CURR_EXCH_RATE") = 1
            End If

            Set_Recurring(chkRecurring.Checked)
            If chkRecurring.Checked Then
                rowAPTINVH1.Item("INV_STATUS") = "R"
            Else
                rowAPTINVH1.Item("INV_STATUS") = "O"
            End If

            rowAPTINVH1.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowAPTINVH1.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowAPTINVH1.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            rowAPTINVH1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowAPTINVH1.Item("REGISTER_IND") = "0"
            rowAPTINVH1.Item("INIT_DATE") = DATETIME_STAMP
            rowAPTINVH1.Item("INIT_OPER") = ASCMAIN1.USER_ID

            chkACCRUE_PRIOR.Checked = False
        Else
            Save_Header_Fields(UltraGroupBox1)
            If EntryMode <> "V" And Not InquiryMode Then Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Called up for Changes")

            If rowAPTINVH1.Item("OPS_YYYYPP_ACCRUE") & "" <> "" Then
                chkACCRUE_PRIOR.Checked = True
            End If
            Set_Recurring(rowAPTINVH1.Item("INV_STATUS") = "R")
        End If

        cmdCheck.Visible = (rowAPTINVH1.Item("INV_STATUS") = "P")

        If rowAPTINVH1.Item("INV_STATUS") = "R" Then
            chkACCRUE_PRIOR.Checked = False
        End If

        Fill_Records("TATEVNT1", HFs("VOUCHER_NO"))
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Load_OPS_YYYYPP_ACCRUE()
        Load_CURR_EXCH_RATE()

        If InquiryMode Then
            If rowAPTINVH1.Item("INV_STATUS") & "" = "P" Then
                Dim rowAPTCHCK1 As DataRow = LookUp("APTCHCK1", New String() {rowAPTINVH1.Item("BANK_CODE") & "", rowAPTINVH1.Item("CHECK_NUM") & ""})
                'Absx1.txtFor("CHECK_NUM").Text = rowAPTCHCK1.Item("CHECK_NUM") & ""
                'Absx1.dteFor("CHECK_DATE").Value = rowAPTCHCK1.Item("CHECK_DATE") & ""
                If rowAPTCHCK1 IsNot Nothing Then
                    Absx1.numFor("CHECK_AMT").Value = Val(rowAPTCHCK1.Item("CHECK_AMT") & "")
                End If
            End If
        End If

        If Absx1.cbeFor("INV_TYPE").Value = "A" Then
            tabMain.Tabs("PO Receipts").Enabled = False
            tabMain.Tabs("Other Accruals").Enabled = False
            tabMain.Tabs("GL Distribution").Enabled = False
        Else
            tabMain.Tabs("GL Distribution").Enabled = True
            If rowAPTINVH1.Item("INV_STATUS") = "R" Then
                tabMain.Tabs("PO Receipts").Enabled = False
                tabMain.Tabs("Other Accruals").Enabled = False
                Fill_Records("APTINVHR", HFs("VOUCHER_NO"))
            Else
                Load_ICTIREC1()

                For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("")
                    Dim RECEIPT_NO As String = rowAPTINVH5.Item("RECEIPT_NO") & ""

                    Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)
                    If RECEIPT_NO <> "" Then
                        rowICTIREC1 = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)
                    Else
                        Dim PO_SHIPMENT_NO As String = rowAPTINVH5.Item("PO_SHIPMENT_NO") & ""
                        Dim PO_SHIPMENT_LNO As Integer = Val(rowAPTINVH5.Item("PO_SHIPMENT_LNO") & "")
                        Dim PO_ORDER_NO As String = rowAPTINVH5.Item("PO_ORDER_NO") & ""
                        Dim PO_ORDER_LNO As Integer = Val(rowAPTINVH5.Item("PO_ORDER_LNO") & "")
                        Dim SQLW As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)  ' & " and PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
                        rowICTIREC1 = dst.Tables("ICTIREC1").Select(SQLW)(0)
                        RECEIPT_NO = rowICTIREC1.Item("RECEIPT_NO")
                        rowAPTINVH5.Item("RECEIPT_NO") = RECEIPT_NO
                    End If

                    Dim rowAPTINVH5_SUM As DataRow = dst.Tables("APTINVH5_SUM").Rows.Find(RECEIPT_NO)
                    If rowAPTINVH5_SUM Is Nothing Then
                        rowAPTINVH5_SUM = Add_APTINVH5_SUM(RECEIPT_NO, CDate(rowICTIREC1.Item("RECEIPT_DATE")), Val(rowICTIREC1.Item("QTY_REC") & ""), Val(rowICTIREC1.Item("AMT_REC") & ""))
                    End If
                Next
            End If
        End If

        lblINV_STATUS_NOTE.Visible = False
        If rowAPTINVH1.Item("VOUCHER_NO_ORIG") & "" <> "" Then
            lblINV_STATUS_NOTE.Text = "Original Voucher " & rowAPTINVH1.Item("VOUCHER_NO_ORIG")
            lblINV_STATUS_NOTE.Visible = True
        ElseIf rowAPTINVH1.Item("INV_STATUS") & "" = "D" Then
            ASCMAIN1.sql = "Select Min (VOUCHER_NO) from APTINVH1 where VOUCHER_NO_ORIG = '" & rowAPTINVH1.Item("VOUCHER_NO") & "'"
            Dim VOUCHER_NO_reversing As String = ASCDATA1.GetDataValue
            lblINV_STATUS_NOTE.Text = "Reversed by Voucher " & VOUCHER_NO_reversing
            lblINV_STATUS_NOTE.Visible = (VOUCHER_NO_reversing <> "")
        End If

        lblRecurringTemplate.Visible = (rowAPTINVH1.Item("INV_STATUS") = "R")
        lblRecurring.Text = "from Recurring Template " & rowAPTINVH1.Item("VOUCHER_NO_RECUR")
        lblRecurring.Visible = (rowAPTINVH1.Item("VOUCHER_NO_RECUR") & "" <> "")

        tabHeader.Tabs("Adjustments").Enabled = Not (rowAPTINVH1.Item("INV_STATUS") = "R")
        tabHeader.Tabs("Recurring").Enabled = (rowAPTINVH1.Item("INV_STATUS") = "R")
        tabHeader.Tabs("Pymt Info").Enabled = (rowAPTINVH1.Item("INV_STATUS") = "P")

        If EntryMode = "N" Then
            Absx1.dteFor("INV_BL_DATE").ReadOnly = InquiryMode
        Else
            '   Setup_INV_BL_DATE()
        End If

        Absx1.txtFor("INV_PYMT_METHOD").ReadOnly = (rowAPTVEND1.Item("VEND_PYMT_METHOD_FIXED") & "" = "1") Or InquiryMode

        tabMain.SelectedTab = tabMain.Tabs(0)
        grdAPTINVH2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdAPTINVH2.DisplayLayout.Bands(0).SortedColumns.Add("VOUCHER_LNO", False)

        If EntryMode = "N" Then
            ' CAN'T GET FOCUS TO GO TO A SPECIFIC CONTROL
            tabMain.SelectedTab = tabMain.Tabs(1)
            tabMain.Focus()
            Application.DoEvents()
            For i As Integer = 1 To 5
                If Absx1.CtlFor("INV_DATE").Focused Then
                    Exit For
                End If
                SendKeys.Send(Chr(9))
                Application.DoEvents()
            Next
        End If

        Fill_Records("APTINVHX", New Object() {HFs("VEND_CODE"), HFs("VOUCHER_NO")})

        Fill_Records("ASTAUDTX", HFs("VOUCHER_NO"))
        Sort_grdColumns(grdASTAUDTX, "INIT_DATE".ToLower)

        EnforceConstraints(True)

        tabHeader.ActiveTab = tabHeader.Tabs("Codes")

        disable_calculate_totals = False

        Set_Distribution_Display_Options()

        Calc_DIST_GL()
        Calc_DIST_PO()
        'Calc_DIST_Adjustments()
        Calc_DIST_Other()
    End Sub

    Sub Update_Record()

        Dim INV_PYMT_METHOD As String = Absx1.txtFor("INV_PYMT_METHOD").Text
        ' Dim INV_PYMT_METHOD2 As String = rowAPTINVH1.Item("INV_PYMT_METHOD") & ""

        If INV_PYMT_METHOD <> "LC" Then
            rowAPTINVH1.Item("LC_CTL_NO") = DBNull.Value
            rowAPTINVH1.Item("LC_FEE") = DBNull.Value
        End If

        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()
        X.SuspendBinding()

        Dim VOUCHER_NO_TO_PAY As String = HFs("VOUCHER_NO")

        Try

            BeginTrans()

            Calculate_INV_DISC_AMT()

            Delete_Rows("APTINVH2", "INV_LTYP is Null and (INV_LINE_AMT = 0 or INV_LINE_AMT is Null)")

            If rowAPTINVH1("OPS_YYYYPP_ACCRUE") & "" _
             = rowAPTINVH1("OPS_YYYYPP") & "" Then
                rowAPTINVH1("OPS_YYYYPP_ACCRUE") = ""
            End If

            Dim pay_upon_entry As Boolean = False
            rowAPTINVH1("INV_BALANCE") = rowAPTINVH1("INV_AMT")
            If rowAPTINVH1("INV_STATUS") = "P" Then
                rowAPTINVH1("INV_PAID_UPON_ENTRY") = "1"
                pay_upon_entry = True
            Else
                rowAPTINVH1("INV_PAID_UPON_ENTRY") = Null
            End If

            If rowAPTINVH1("INV_STATUS") = "R" Then
                rowAPTINVH1("REGISTER_IND") = "R"
            End If

            If rowAPTINVH1("INV_REMIT_TO") = "N" Then
                dst.Tables("APTVEND2").AcceptChanges()
                dst.Tables("APTVEND2").Rows(0).SetAdded()
                Update_Record_TDA("APTVEND2")
                rowAPTINVH1("INV_REMIT_TO") = "A"
            End If

            rowAPTVEND5 = Fill_Record("APTVEND5", HFs("VEND_CODE"), True)

            Create_APTINVH5_VAR()
            Create_APTINVH2_P()
            ''Create_APTINVH2_R()
            Create_APTINVH2_from_APTINVH7()

            For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                Dim RECEIPT_NO As String = rowAPTINVH5.Item("RECEIPT_NO") & ""
                If RECEIPT_NO.StartsWith("S") Then ' SHIPMENT BEING PAID, NOT A RECEIPT
                    rowAPTINVH5.Item("RECEIPT_NO") = DBNull.Value
                    rowAPTINVH5.Item("RECEIPT_LNO") = DBNull.Value
                End If
            Next

            If EntryMode = "N" Then
                If rowAPTINVH1("INV_PAID_UPON_ENTRY") & "" = "1" Then
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Entered as Paid")
                Else
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Entered")
                End If

                Update_Record_TDA("APTINVH1")
                Update_Record_TDA("APTINVH2")
                Update_Record_TDA("APTINVH8")

                Update_Record_TDA("POTLCST1") ' TO GET ADDED ROWS OUT THERE - DEPENDENT UPDATES WILL DO THE REST

                For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                    rowAPTINVH5.Item("VAR_QTY") = rowAPTINVH5.Item("QTY_VAR")
                    rowAPTINVH5.Item("VAR_AMT") = rowAPTINVH5.Item("AMT_VAR")

                    ' DON'T DO THIS - RELATIONS
                    'Dim RECEIPT_NO As String = rowAPTINVH5.Item("RECEIPT_NO") & ""
                    'If RECEIPT_NO.StartsWith("S") Then ' SHIPMENT BEING PAID, NOT A RECEIPT
                    '    rowAPTINVH5.Item("RECEIPT_NO") = DBNull.Value
                    '    rowAPTINVH5.Item("RECEIPT_LNO") = DBNull.Value
                    'End If

                Next
                Update_Record_TDA("APTINVH5", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'")
                'ASCMAIN1.sql = "Update APTINVH5 Set RECEIPT_NO = NULL, RECEIPT_LNO = NULL where VOUCHER_NO = '" & HFs("VOUCHER_NO") & "' and RECEIPT_NO LIKE 'S%'"
                'ASCDATA1.ExecuteSQL()

                Update_Record_TDA("APTINVH7", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'")

                dst.Tables("APTINVH5").Rows.Clear()
                dst.Tables("APTINVH8").Rows.Clear()
                dst.Tables("APTINVH7").Rows.Clear()

                Dependent_Updates(HFs("VOUCHER_NO"), False)

                If rowAPTINVH1("INV_TYPE") = "A" Then ' If Absx1.cbeFor("INV_TYPE").Value = "A" Then
                    Dim VOUCHER_NO_ADV As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")

                    ReNumber_Voucher(HFs("VOUCHER_NO"), VOUCHER_NO_ADV)
                    rowAPTINVH1("INV_PAID_UPON_ENTRY") = Null
                    rowAPTINVH1("INV_STATUS") = "O"
                    rowAPTINVH1("CHECK_NUM") = ""
                    rowAPTINVH1("CHECK_DATE") = Null
                    Negate_Voucher(VOUCHER_NO_ADV)

                    Update_Record_TDA("APTINVH1")
                    Update_Record_TDA("APTINVH2")
                End If
                'Stop
            Else
                Write_Audit_Trail(rowAPTINVH1, Nothing, "E")

                Dim something_GL_related_was_changed As Boolean = False
                If dst.Tables("APTINVH2").GetChanges IsNot Nothing _
                    Or dst.Tables("APTINVH5").GetChanges IsNot Nothing _
                    Or dst.Tables("APTINVH7").GetChanges IsNot Nothing Then
                    something_GL_related_was_changed = True
                End If
                'If dst.Tables("APTINVH2").GetChanges IsNot Nothing Then
                '    something_GL_related_was_changed = True
                'End If
                For Each COLUMN_NAME As String In New String() _
                {"INV_AMT", "INV_DATE", "OPS_YYYYPP_ACCRUE", "POST_CODE", "CURR_CODE", "CURR_EXCH_RATE", "INV_PAID_UPON_ENTRY"}
                    If rowAPTINVH1(COLUMN_NAME, DataRowVersion.Current) & "" _
                    <> rowAPTINVH1(COLUMN_NAME, DataRowVersion.Original) & "" Then
                        something_GL_related_was_changed = True
                    End If
                Next

                If rowAPTINVH1("REGISTER_IND") = "0" _
                Or rowAPTINVH1("REGISTER_IND") = "R" _
                Or Not something_GL_related_was_changed Then
                    Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Changed")
                    Dependent_Updates(HFs("VOUCHER_NO"), True)

                    rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
                    rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID

                    Dim sql_delete As String = "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'"
                    Update_Record_TDA("APTINVH1", sql_delete)
                    Update_Record_TDA("APTINVH2", sql_delete)

                    For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                        rowAPTINVH5.Item("VAR_QTY") = rowAPTINVH5.Item("QTY_VAR")
                        rowAPTINVH5.Item("VAR_AMT") = rowAPTINVH5.Item("AMT_VAR")
                    Next
                    Update_Record_TDA("APTINVH5", sql_delete)

                    Update_Record_TDA("APTINVH8", sql_delete)
                    Update_Record_TDA("APTINVH7", sql_delete)

                    Dependent_Updates(HFs("VOUCHER_NO"), False)
                Else
                    Dim VOUCHER_NO_ORIG As String = HFs("VOUCHER_NO")
                    Dim VOUCHER_NO_NEG As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
                    Dim VOUCHER_NO_NEW As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
                    VOUCHER_NO_TO_PAY = VOUCHER_NO_NEW

                    Write_Event_Log(TABLE_NAME, VOUCHER_NO_ORIG, "Reversed (" & VOUCHER_NO_NEG & ") and Replaced (" & VOUCHER_NO_NEW & ")")

                    Dependent_Updates(VOUCHER_NO_ORIG, True)

                    ReNumber_Voucher(VOUCHER_NO_ORIG, VOUCHER_NO_NEW)
                    rowAPTINVH1("REGISTER_IND") = "0"
                    rowAPTINVH1("REGISTER_XNO") = Null
                    rowAPTINVH1("OPS_YYYYPP") = ASCMAIN1.CYP
                    rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
                    rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID
                    Update_Record_TDA("APTINVH1")
                    Update_Record_TDA("APTINVH2")
                    Update_Record_TDA("APTINVH5")
                    Update_Record_TDA("APTINVH8")
                    Update_Record_TDA("APTINVH7")

                    Dependent_Updates(VOUCHER_NO_NEW, False)

                    ReLoad_Voucher(VOUCHER_NO_ORIG)
                    rowAPTINVH1 = dst.Tables("APTINVH1").Rows(0)
                    rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
                    rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID
                    rowAPTINVH1("INV_STATUS") = "D"
                    Update_Record_TDA("APTINVH1")

                    ReNumber_Voucher(VOUCHER_NO_ORIG, VOUCHER_NO_NEG)
                    Negate_Voucher(VOUCHER_NO_NEG)
                    rowAPTINVH1("REGISTER_IND") = "0"
                    rowAPTINVH1("REGISTER_XNO") = Null
                    rowAPTINVH1("OPS_YYYYPP") = ASCMAIN1.CYP

                    Update_Record_TDA("APTINVH1")
                    Update_Record_TDA("APTINVH2")

                    For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5").Select("", "", DataViewRowState.CurrentRows)
                        rowAPTINVH5.Item("VAR_QTY") = rowAPTINVH5.Item("QTY_VAR")
                        rowAPTINVH5.Item("VAR_AMT") = rowAPTINVH5.Item("AMT_VAR")
                    Next
                    Update_Record_TDA("APTINVH5")

                    Update_Record_TDA("APTINVH8")
                    Update_Record_TDA("APTINVH7")
                End If
            End If


            If pay_upon_entry Then ' If rowAPTINVH1("INV_PAID_UPON_ENTRY") & "" = "1" Then

                Update_as_Paid(VOUCHER_NO_TO_PAY)
                Update_Record_TDA("APTCHCK1")
                Update_Record_TDA("APTCHCK2")
                Update_Record_TDA("APTINVH1")
                If auto_next_check Then
                    Update_Record_TDA("GLTBANK1")
                End If
            End If
            Update_Record_TDA("APTVEND5")

            If batch_update Then
                CommitTrans()
            Else
                CommitTrans("Update Complete")
            End If

        Catch ex As Exception
            Rollback("Error Occurred - Please call ABS" & vbCrLf & ex.Message, ex)

        End Try

        X.ResumeBinding()

    End Sub

    Sub Delete_Record()

        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()
        X.SuspendBinding()

        BeginTrans()

        If EntryMode = "N" Then
            Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Entered and Deleted before Update")
        Else
            Write_Event_Log(TABLE_NAME, HFs("VOUCHER_NO"), "Deleted")
            rowAPTVEND5 = Fill_Record("APTVEND5", HFs("VEND_CODE"), True)

            Dependent_Updates(HFs("VOUCHER_NO"), True)
            Dim VOUCHER_NO_ORIG As String = HFs("VOUCHER_NO")
            ReLoad_Voucher(VOUCHER_NO_ORIG)

            rowAPTINVH1 = dst.Tables("APTINVH1").Rows(0)
            rowAPTINVH1("LAST_DATE") = DATETIME_STAMP
            rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID
            'rowAPTINVH1("INV_BALANCE") = 0 ' NOT NEC AND NOT CONSISTENT WITH OTHER D'S
            rowAPTINVH1("INV_STATUS") = "D"
            If rowAPTINVH1("REGISTER_IND") = "0" Then
                rowAPTINVH1("REGISTER_IND") = "D"
            End If
            Update_Record_TDA("APTINVH1")
            Update_Record_TDA("APTVEND5")

            If rowAPTINVH1("REGISTER_IND") = "1" Then
                Dim VOUCHER_NO_NEG As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")

                Write_Event_Log(TABLE_NAME, VOUCHER_NO_ORIG, "Reversed (" & VOUCHER_NO_NEG & ")")

                ReNumber_Voucher(VOUCHER_NO_ORIG, VOUCHER_NO_NEG)
                Negate_Voucher(VOUCHER_NO_NEG)
                rowAPTINVH1("REGISTER_IND") = "0"
                rowAPTINVH1("REGISTER_XNO") = Null
                rowAPTINVH1("OPS_YYYYPP") = ASCMAIN1.CYP

                Update_Record_TDA("APTINVH1")
                Update_Record_TDA("APTINVH2")
                Update_Record_TDA("APTINVH5")
                Update_Record_TDA("APTINVH8")
                Update_Record_TDA("APTINVH7")
            End If
        End If

        X.ResumeBinding()
        CommitTrans("Deletion Completed")


    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'")
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "Load"
                Absx1.txtFor("VOUCHER_NO").Text = key
                Click_Command("Load")
        End Select

        Return return_key
    End Function

    Public Overrides Function Data_Export_Context() As ABSolution.ASFBASE0.Data_Export_Entity

        Dim E As New Data_Export_Entity
        E.enabled = True
        ASTDATA1s.Clear()
        'ASTDATA1s.Add("APTINVHX", "Vendor Invoices")
        ASTDATA1s.Add("APTINVH5", "Invoiced Purchase Accruals")
        ASTDATA1s.Add("APTINVH8", "Invoice Adjustments")
        ASTDATA1s.Add("APTINVH7", "Invoiced Other PO Accruals")

        Return E
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "APTINVH1"
            E.COLUMN_NAME = "VOUCHER_NO"
            E.CODE_VALUE = Absx1.txtFor("VOUCHER_NO").Text
            E.DESC_VALUE = "Vendor Invoice"
            E.ATTACHMENT_NOTES = ""
            If rowAPTINVH1.Item("INV_STATUS") & "" <> "O" And rowAPTINVH1.Item("INV_STATUS") & "" <> "H" Then
                E.RESTRICTIONS = "D"
            End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    Public Overrides Function Audit_Context() As Audit_Entity

        Dim E As New Audit_Entity
        If ScreenMode Then
            E.TABLE_NAME = "APTINVH1"
            E.KEY_VALUE = Absx1.txtFor("VOUCHER_NO").Text
            E.KEY_DESC = "Vendor Invoice"
        End If
        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "APTINVH1"
        E.TABLE_KEY_CAPTION = "AP"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("VOUCHER_NO").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text & " " & Absx1.txtFor("VEND_NAME").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "A")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTIREC1, "SSB", "Show Filter", "Show GroupBox", "PO Inquiry", "Shipment Inquiry")
        Load_Popup_Menu(grdAPTINVH5_SUM, "SS", "Show Filter", "Show GroupBox")
        ' Load_Popup_Menu(grdAPTINVH5, "SSSSSBB", "Show Filter", "Show GroupBox", "Show Description", "Show Receipt Qty/Price/Amt", "Discrepancies Only", "Copy Price to All Lines", "Check All Lines", "Uncheck All Lines")
        Load_Popup_Menu(grdAPTINVH5, "BBBBB", "PO Inquiry", "Use PO Cost as AP Cost", "Copy Price to All Lines", "Check CB on All Lines", "Uncheck CB on All Lines")
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

        If tlb_pop.Tools.Exists("Show Description") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Description"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("STYLE_DESC").Hidden
        End If

        If tlb_pop.Tools.Exists("Show Receipt Qty/Price/Amt") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Receipt Qty/Price/Amt"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("QTY_REC").Hidden
        End If

        If tlb_pop.Tools.Exists("Discrepancies Only") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Discrepancies Only"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = discrepancies_only
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdAPTINVH5"
                    For Each key As String In New String() {"Use PO Cost as AP Cost", "Copy Price to All Lines", "Check CB on All Lines", "Uncheck CB on All Lines"}
                        tlb_btn = DirectCast(tlb_pop.Tools(key), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = (EntryMode = "N" Or EntryMode = "E") And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")
                    Next
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Description"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns("STYLE_DESC").Hidden = Not tlb_sbt.Checked

            Case "Show Receipt Qty/Price/Amt"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns("QTY_REC").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Columns("PO_COST").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Columns("AP_COST").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Columns("AMT_REC").Hidden = Not tlb_sbt.Checked

            Case "Discrepancies Only"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                discrepancies_only = tlb_sbt.Checked
                If grdICTIREC1.ActiveRow IsNot Nothing Then
                    Dim RECEIPT_NO As String = grdICTIREC1.ActiveRow.Cells("RECEIPT_NO").Text
                    Setup_grdAPTINVH5(RECEIPT_NO)
                End If

            Case "Use PO Cost as AP Cost"
                If MsgBox("Use PO Cost as the AP Cost for All Lines Displayed on This Receipt?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Updating Lines")
                    For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH5.Rows
                        grow.Cells("INV_COST").Value = grow.Cells("PO_COST").Value
                        grow.Cells("CB").Value = "0"
                        grow.Update()
                    Next
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                End If


            Case "Copy Price to All Lines"
                If grd.ActiveCell Is Nothing Then
                Else
                    If grd.ActiveCell.Column.Key = "INV_COST" Then
                        Dim INV_COST As Double = Val(grd.ActiveCell.Value & "")
                        If MsgBox("Copy Invoice Cost of " & Format(INV_COST, "#.00") & " to All Lines Displayed on This Receipt?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            Me.Cursor = Cursors.WaitCursor
                            ASCMAIN1.Progress("Now Updating Lines")
                            For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH5.Rows
                                grow.Cells("INV_COST").Value = INV_COST
                                grow.Update()
                            Next
                            Me.Cursor = Cursors.Default
                            ASCMAIN1.Progress("")
                        End If
                    End If
                End If

            Case "Check CB on All Lines"
                If MsgBox("Check All Lines Displayed on This Receipt?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Updating Lines")
                    For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH5.Rows
                        grow.Cells("CB").Value = "1"
                        grow.Update()
                    Next
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                End If

            Case "Uncheck CB on All Lines"
                If MsgBox("Uncheck All Lines Displayed on This Receipt?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Updating Lines")
                    For Each grow As UltraWinGrid.UltraGridRow In grdAPTINVH5.Rows
                        grow.Cells("CB").Value = "0"
                        grow.Update()
                    Next
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                If grd.ActiveRow.Band.Key = "ICTIREC2" Or grd.ActiveRow.Band.Key = "APTINVH5" Then
                    Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value
                    Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")
                End If

            Case "Style Status Inquiry"
                If grd.ActiveRow.Band.Key = "ICTIREC2" Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Shipment Inquiry"
                'If grd.Name = "grdPOTSHIPC" And grd.ActiveRow.Band.Key <> "POTSHIPC_POTSHIPC2" Then
                '    Exit Sub
                'End If
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "VEND_CODE"

                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Not chkQuickEntry.Checked Then
                        'Click_Command("New", e)
                    End If
                End If

            Case "INV_NUM"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    If Not chkQuickEntry.Checked Then
                        Click_Command("New", e)
                    End If
                End If

            Case "VOUCHER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Click_Command("Edit", e)
                End If

        End Select

    End Sub

    Overrides Sub num_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.num_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

            Case "QE_INV_AMT"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Click_Command("New", e)
                End If


        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "TERM_CODE"
                Calculate_INV_DUE_DATE()
            Case "CURR_CODE"
                Load_CURR_EXCH_RATE()
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"

            Case "INV_PYMT_METHOD"
                grpLC.Visible = (Absx1.txtFor("INV_PYMT_METHOD").Text = "LC")
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "VOUCHER_NO"
                Click_Command("Edit")
            Case "VEND_ALT_CODE"
                Load_Alternate_Payment_Address()
            Case "OPS_YYYYPP_ACCRUE"

            Case "PO_ORDER_NO"
            Case "PO_SHIPMENT_NO"
            Case "PO_SHIPMENT_LNO"

        End Select
    End Sub

#End Region

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged

        If EntryMode = "" Then
            Exit Sub
        End If
        With UltraExplorerBar1
            ' .Groups("Entry Options").Visible = False
            .Groups("Distribution Options").Visible = False
            .Groups("Generate $0 Accrual").Visible = False

            Select Case tabMain.ActiveTab.Key
                Case "Vendor Information"
                Case "GL Distribution"
                    '     .Groups("Distribution Options").Visible = True And Not InquiryMode
                Case "Header Data"
                    '.Groups("Payment Options").Visible = True and not inquirymode
                Case "Details"

                Case "Other Accruals"
                    .Groups("Generate $0 Accrual").Visible = True And Not InquiryMode
            End Select
        End With

    End Sub

    Private Sub optINV_REMIT_TO_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optINV_REMIT_TO.ValueChanged
        lblVEND_CODE_AP.Visible = (optINV_REMIT_TO.Value = "P")
        Absx1.txtFor("VEND_CODE_AP").Visible = (optINV_REMIT_TO.Value = "P")
        Absx1.txtFor("VEND_ALT_CODE").Visible = (optINV_REMIT_TO.Value = "A" Or optINV_REMIT_TO.Value = "N")
        Absx1.txtFor("VEND_ALT_NAME").ReadOnly = (optINV_REMIT_TO.Value <> "N") Or InquiryMode
        For Each COLUMN_NAME As String In New String() {"VEND_ALT_NAME", "VEND_ALT_ADDR1", "VEND_ALT_ADDR2", "VEND_ALT_ADDR3", "VEND_ALT_CITY", "VEND_ALT_STATE", "VEND_ALT_ZIP_CODE", "VEND_ALT_PHONE", "VEND_ALT_EXT", "VEND_ALT_FAX", "VEND_ALT_COUNTRY", "VEND_ALT_CONTACT", "VEND_ALT_EMAIL"}
            If COLUMN_NAME = "VEND_ALT_PHONE" Or COLUMN_NAME = "VEND_ALT_FAX" Then
                Absx1.medFor(COLUMN_NAME).ReadOnly = (optINV_REMIT_TO.Value <> "N") Or InquiryMode
            Else
                Absx1.txtFor(COLUMN_NAME).ReadOnly = (optINV_REMIT_TO.Value <> "N") Or InquiryMode
            End If
        Next

        'If optINV_REMIT_TO.Value <> "N" Then
        Set_Payment_Address()
        'End If
    End Sub

    Private Sub optINV_STATUS_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optINV_STATUS.ValueChanged
        tabHeader.Tabs("Pymt Info").Enabled = (optINV_STATUS.Value = "P")
        If optINV_STATUS.Value = "P" AndAlso (EntryMode = "E" Or EntryMode = "N") Then
            tabHeader.Tabs("Pymt Info").Selected = True
        Else
            If tabHeader.SelectedTab IsNot Nothing AndAlso tabHeader.SelectedTab.Key = "Pymt Info" Then
                tabHeader.Tabs("Codes").Selected = True
            End If
        End If
    End Sub

    Sub Load_ICTIREC1()

        ASCMAIN1.Progress("Now Loading Receipts")
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Dim sql_where As String = ""

        If EntryMode = "" Then
            If optPOReceipts.Value = "O" Then
                sql_where = " where ICTIREC1.ACCRUAL_STATUS = '0'"
            Else
                sql_where = " where ICTIREC1.OPS_YYYYPP = '" & cbeYP.Value & "'"
            End If
        Else
            sql_where = " where ICTIREC1.VEND_CODE = '" & HFs("VEND_CODE") & "'" & vbCrLf _
                & " and ICTIREC1.RECEIPT_NO in " _
                & " (Select Distinct RECEIPT_NO from APTINVH5 where VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'"
            If InquiryMode Then
                sql_where &= ")"
            Else
                sql_where &= " union " _
                & "  Select RECEIPT_NO from ICTIREC1 where ICTIREC1.ACCRUAL_STATUS = '0' AND VEND_CODE = '" & HFs("VEND_CODE") & "')"
            End If
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTIREC1)

        If Not batch_update Then

            ASCDATA1.ExecuteSQL("Insert into " & ICTIREC1 & vbCrLf _
                    & " Select ICTIREC1.*" & vbCrLf _
                    & " from ICTIREC1" & vbCrLf _
                    & sql_where)

            If ASCMAIN1.CLIENT = "VAN" Then ' RE-ENABLING BECAUSE LC PAYMENTS REQUIRE PO DETAILS SO MARIA WOUND UP NEEDING TO USE THIS FEATURE '  If ASCMAIN1.CLIENT = "VAN" And 1 <> 1 Then ' DISABLING THIS FEATURE ON 03/04/19 AS PER CONVERSATION WITH ANNA, MARIA, NADINE - SEE EMAIL SENT

                If EntryMode = "" Then
                    If optPOReceipts.Value = "O" Then
                        sql_where = " and POTSHIP2.ACCRUAL_STATUS = '0' AND POTSHIP2.PO_SHIP_STATUS = 'O'"
                    Else
                        sql_where = " and ROWNUM < 1"
                    End If
                Else
                    sql_where = " and POTORDR1.VEND_CODE = '" & HFs("VEND_CODE") & "'" & vbCrLf _
                            & " and (POTSHIP3.PO_SHIPMENT_NO,POTSHIP3.PO_SHIPMENT_LNO) in " _
                            & " (Select Distinct PO_SHIPMENT_NO,PO_SHIPMENT_LNO from APTINVH5 where RECEIPT_NO IS NULL AND VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'"
                    If InquiryMode Then
                        sql_where &= ")"
                    Else
                        sql_where &= " union " _
                                & "  Select PO_SHIPMENT_NO,PO_SHIPMENT_LNO from POTSHIP2 where POTSHIP2.ACCRUAL_STATUS = '0' AND POTSHIP2.PO_SHIP_STATUS = 'O')"
                    End If
                End If


                ASCDATA1.ExecuteSQL("Delete from " & POTSHIP3)
                ASCMAIN1.sql = "Select POTSHIP3.*,POTORDR1.VEND_CODE, ROWNUM RECEIPT_LNO" & vbCrLf _
                                & " from POTSHIP3,POTSHIP2,POTORDR1" & vbCrLf _
                                & " where POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                                & sql_where _
                                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO"
                'If EntryMode = "N" Or EntryMode = "E" Or EntryMode = "V" Then
                '    ASCMAIN1.sql &= "   and POTORDR1.VEND_CODE = '" & HFs("VEND_CODE") & "'" & vbCrLf
                'End If
                ASCDATA1.ExecuteSQL("Insert into " & POTSHIP3 & " " & ASCMAIN1.sql)

                ASCDATA1.ExecuteSQL("Delete from " & POTSHIP2)
                ASCMAIN1.sql = "Select POTSHIP2.* from POTSHIP2" & vbCrLf _
                                & " where (PO_SHIPMENT_NO,PO_SHIPMENT_LNO) " & vbCrLf _
                                & " in (Select Distinct PO_SHIPMENT_NO, PO_SHIPMENT_LNO from " & POTSHIP3 & ")"
                ASCDATA1.ExecuteSQL("Insert into " & POTSHIP2 & " " & ASCMAIN1.sql)

                ASCMAIN1.sql = "Update " & POTSHIP2 & " Set TRAN_NO = 'S' || TRIM(TO_CHAR(ROWNUM,'00000'))"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "" & vbCrLf _
                                & "Begin" & vbCrLf _
                                & " Declare Cursor C1 is Select * from " & POTSHIP2 & ";" & vbCrLf _
                                & " Begin" & vbCrLf _
                                & "  For R1 in C1 Loop" & vbCrLf _
                                & "   Update " & POTSHIP3 & " Set RECEIPT_LNO = ROWNUM " & vbCrLf _
                                & "    where PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO" & vbCrLf _
                                & "      and PO_SHIPMENT_LNO = R1.PO_SHIPMENT_LNO;" & vbCrLf _
                                & "  End Loop;" & vbCrLf _
                                & " End;" & vbCrLf _
                                & "End;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Select" & vbCrLf _
                                & "  POTSHIP2.TRAN_NO RECEIPT_NO" & vbCrLf _
                                & ", POTSHIP1.PO_DATE_SHIPPED RECEIPT_DATE" & vbCrLf _
                                & ", POTSHIP3X.VEND_CODE" & vbCrLf _
                                & ", NULL OPS_YYYYPP" & vbCrLf _
                                & ", POTSHIP2.INIT_DATE" & vbCrLf _
                                & ", POTSHIP2.INIT_OPER" & vbCrLf _
                                & ", POTSHIP2.LAST_DATE" & vbCrLf _
                                & ", POTSHIP2.LAST_OPER" & vbCrLf _
                                & ", NULL REGISTER_IND" & vbCrLf _
                                & ", NULL REGISTER_XNO" & vbCrLf _
                                & ", POTSHIP1.WHSE_CODE" & vbCrLf _
                                & ", POTSHIP2.ACCRUAL_STATUS" & vbCrLf _
                                & ", POTSHIP2.PO_SOURCE_DOC SOURCE_DOC_NO" & vbCrLf _
                                & ", NULL VOUCHER_NO" & vbCrLf _
                                & ", POTSHIP3X.QTY_REC" & vbCrLf _
                                & ", POTSHIP3X.AMT_REC" & vbCrLf _
                                & ", 0 QTY_INV" & vbCrLf _
                                & ", 0 AMT_INV" & vbCrLf _
                                & ", NULL REVERSED_BY_RECEIPT_NO" & vbCrLf _
                                & ", NULL REVERSES_RECEIPT_NO" & vbCrLf _
                                & ", POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                                & ", POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                                & " from " & POTSHIP2 & " POTSHIP2,POTSHIP1" & vbCrLf _
                                & ", (Select VEND_CODE, PO_SHIPMENT_NO, PO_SHIPMENT_LNO" & vbCrLf _
                                & ", Sum (PO_QTY_SHP) QTY_REC, Sum (PO_QTY_SHP * PO_COST) AMT_REC" & vbCrLf _
                                & " from " & POTSHIP3 & vbCrLf _
                                & " group by VEND_CODE, PO_SHIPMENT_NO, PO_SHIPMENT_LNO) POTSHIP3X" & vbCrLf _
                                & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_NO = POTSHIP3X.PO_SHIPMENT_NO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3X.PO_SHIPMENT_LNO"
                '& "   and POTSHIP2.ACCRUAL_STATUS = '0'" & vbCrLf _
                '& "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _

                ASCDATA1.ExecuteSQL("Insert into " & ICTIREC1 & " " & ASCMAIN1.sql)

            End If
        End If


        dst.Tables("ICTIREC2").Rows.Clear()
        Fill_Records("ICTIREC1")

        Sort_grdColumns(grdICTIREC1, "RECEIPT_NO".ToLower)

        Setup_ICTIREC1()

        If EntryMode <> "" Then
            Setup_Other_Accruals()
            tabMain.Tabs("Other Accruals").Enabled = True
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Setup_Other_Accruals()
        If EntryMode = "" Then
            Exit Sub
        End If
        'If grdICTIREC1.Rows.Count <> 0 Then
        '    grdICTIREC1.ActiveRow = grdICTIREC1.Rows(0)
        '    tabMain.Tabs("PO Receipts").Enabled = True
        'Else
        '    tabMain.Tabs("PO Receipts").Enabled = False
        'End If
    End Sub

    Sub Setup_ICTIREC1()
        ' Return
        If EntryMode = "" Then
            Exit Sub
        End If
        If grdICTIREC1.Rows.Count <> 0 Then
            grdICTIREC1.ActiveRow = grdICTIREC1.Rows(0)
            tabMain.Tabs("PO Receipts").Enabled = True
        Else
            tabMain.Tabs("PO Receipts").Enabled = False
        End If
    End Sub

    Private Sub grdICTIREC1_BeforeRowExpanded(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIREC1.BeforeRowExpanded
        If grdICTIREC1.ActiveRow.IsGroupByRow Then
        Else
            Me.Cursor = Cursors.WaitCursor
            Dim RECEIPT_NO As String = e.Row.Cells("RECEIPT_NO").Text '  grdICTIREC1.ActiveRow.Cells("RECEIPT_NO").Text
            Fill_Records("ICTIREC2", New String() {RECEIPT_NO})
            grdICTIREC1.DisplayLayout.Bands("ICTIREC2").SummaryFooterCaption = "Totals for Receipt " & RECEIPT_NO
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub grdICTIREC1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTIREC1.InitializeRow
        If e.Row.Band.Key = "ICTIREC1" Then
            If e.Row.Cells("ACCRUAL_STATUS").Text = "2" Then
                e.Row.Appearance.BackColor = Drawing.Color.Yellow
            Else
                If e.Row.Cells("ACCRUAL_STATUS").Text = "0" Then
                    e.Row.Appearance.BackColor = Drawing.Color.Empty
                Else
                    e.Row.Appearance.BackColor = Drawing.Color.FromArgb(0, 0, 0, 0)
                End If
            End If
        End If

    End Sub

    Private Sub grdAPTINVH5_SUM_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH5_SUM.AfterRowActivate

        Dim RECEIPT_NO As String = grdAPTINVH5_SUM.ActiveRow.Cells("RECEIPT_NO").Text
        Setup_grdAPTINVH5(RECEIPT_NO)
        Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)
        Dim i As Integer = dst.Tables("ICTIREC1").Rows.IndexOf(rowICTIREC1)
        grdICTIREC1.Rows.GetRowWithListIndex(i).Activate()

    End Sub

    Sub Setup_grdAPTINVH5(ByVal RECEIPT_NO As String)

        Dim dvw As DataView = DirectCast(grdAPTINVH5.DataSource, DataTable).DefaultView
        dvw.RowFilter = "RECEIPT_NO = '" & RECEIPT_NO & "'"
        If discrepancies_only Then
            dvw.RowFilter &= "AND ISNULL(QTY_REC,0) <> ISNULL(INV_QTY,0)"
        End If

        grdAPTINVH5.DisplayLayout.Bands("APTINVH5").SummaryFooterCaption = "Totals for Receipt " & RECEIPT_NO
        grdAPTINVH5.Visible = True

    End Sub

    Private Sub grdAPTINVH5_SUM_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH5_SUM.AfterRowsDeleted
        grdAPTINVH5.Visible = False

        Dim RECEIPT_NOs As List(Of String) = DirectCast(grdAPTINVH5_SUM.Tag, List(Of String))

        For Each RECEIPT_NO As String In RECEIPT_NOs
            Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO)
            rowICTIREC1.Item("ACCRUAL_STATUS") = "0"
        Next
        'grdICTIREC1.ActiveRow.Cells("ACCRUAL_STATUS").Value = "0"
        'grdICTIREC1.UpdateData()

        Calc_DIST_PO()
    End Sub

    Private Sub grdAPTINVH5_SUM_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdAPTINVH5_SUM.BeforeRowsDeleted

        Dim RECEIPT_NOs As New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            RECEIPT_NOs.Add(grow.Cells("RECEIPT_NO").Value)
        Next

        grdAPTINVH5_SUM.Tag = RECEIPT_NOs

    End Sub

    'Sub Display_Totals()

    '    ' this section is unnec

    '    Dim INV_AMT As Decimal = Val(Absx1.numFor("INV_AMT").Value & "")
    '    Dim DIST_GL As Decimal = Val(dst.Tables("APTINVH2").Compute("SUM(INV_LINE_AMT)", "INV_LTYP is Null or INV_LTYP = 'A'") & "")
    '    Dim DIST_PO As Decimal = 0 ' Val(dst.Tables("APTINVH5_SUM").Compute("SUM(AMT_INV)", "") & "")

    '    Dim DIST_OTHER As Decimal = Val(dst.Tables("APTINVH7").Compute("SUM(TOTAL_INV)", "") & "")
    '    Dim DIST_OOBAL As Decimal = INV_AMT - DIST_GL - DIST_PO - DIST_OTHER

    '    Absx1.numFor("DIST_GL").Value = DIST_GL
    '    Absx1.numFor("DIST_PO").Value = DIST_PO
    '    Absx1.numFor("DIST_OTHER").Value = DIST_OTHER
    '    Absx1.numFor("DIST_OOBAL").Value = DIST_OOBAL
    'End Sub

    Sub Set_Payment_Address()
        If EntryMode = "" Then
            Exit Sub
        End If
        If optINV_REMIT_TO.Value & "" = "" Then
            Exit Sub
        End If
        Dim rowAPTVEND2 As DataRow

        With dst.Tables("APTVEND2")
            .Rows.Clear()
            rowAPTVEND2 = .NewRow
            rowAPTVEND2.Item("VEND_CODE") = HFs("VEND_CODE")
            'rowAPTVEND2.Item("VEND_CODE") = rowAPTINVH1.Item("VEND_CODE")

            Select Case optINV_REMIT_TO.Value
                Case "V"
                    rowAPTVEND2.Item("VEND_ALT_CODE") = "VENDOR"
                    Absx1.txtFor("VEND_ALT_CODE").Text = ""
                    For i As Integer = 0 To .Columns.Count - 1
                        Dim COLUMN_NAME As String = .Columns(i).ColumnName
                        If COLUMN_NAME <> "VEND_ALT_CODE" Then
                            COLUMN_NAME = Replace(COLUMN_NAME, "_ALT_", "_")
                            rowAPTVEND2.Item(i) = rowAPTVEND1.Item(COLUMN_NAME)
                        End If
                    Next

                Case "P"
                    rowAPTVEND2.Item("VEND_ALT_CODE") = "VENDOR"
                    Absx1.txtFor("VEND_ALT_CODE").Text = ""
                    If Absx1.txtFor("VEND_CODE_AP").Text = "" Then
                        Absx1.txtFor("VEND_CODE_AP").Text = rowAPTVEND2.Item("VEND_CODE")
                    End If
                    LookUp("APTVEND1", Absx1.txtFor("VEND_CODE_AP").Text, True)
                    For i As Integer = 0 To .Columns.Count - 1
                        Dim COLUMN_NAME As String = .Columns(i).ColumnName
                        If COLUMN_NAME <> "VEND_ALT_CODE" Then
                            COLUMN_NAME = Replace(COLUMN_NAME, "_ALT_", "_")
                            rowAPTVEND2.Item(i) = cdr.Item(COLUMN_NAME)
                        End If
                    Next

                Case "A"
                    Absx1.txtFor("VEND_ALT_CODE").ReadOnly = InquiryMode
                    rowAPTVEND2.Item("VEND_ALT_CODE") = rowAPTINVH1.Item("VEND_ALT_CODE")

                    LookUp("APTVEND2", New String() {HFs("VEND_CODE"), rowAPTVEND2.Item("VEND_ALT_CODE") & ""}, True)
                    'LookUp("APTVEND2", New String() {rowAPTVEND2.Item("VEND_CODE") & "", rowAPTVEND2.Item("VEND_ALT_CODE") & ""}, True)
                    If cdr IsNot Nothing Then
                        rowAPTVEND2.ItemArray = cdr.ItemArray
                    End If

                Case "N"
                    Absx1.txtFor("VEND_ALT_CODE").ReadOnly = InquiryMode
                    rowAPTVEND2.Item("VEND_ALT_CODE") = "NEW"
                    Absx1.txtFor("VEND_ALT_CODE").Text = "NEW"
            End Select
            .Rows.Add(rowAPTVEND2)
        End With

    End Sub

    Private Sub grdAPTINVH2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH2.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdAPTINVH2, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
        End Select
    End Sub

    Private Sub grdAPTINVH2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH2.AfterExitEditMode
        With grdAPTINVH2
            Select Case .ActiveCell.Column.Key
                Case "ACCT_CODE"
                    Dim ACCT_CODE As String = .ActiveCell.Text
                    If ACCT_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdAPTINVH2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH2.AfterRowActivate
        With grdAPTINVH2
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdAPTINVH2.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdAPTINVH2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH2.AfterRowsDeleted
        Calc_DIST_GL()
    End Sub

    Private Sub grdAPTINVH2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH2.AfterRowUpdate
        Calc_DIST_GL()
    End Sub

    Private Sub grdAPTINVH2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVH2.BeforeRowUpdate
        With grdAPTINVH2
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
                If e.Row.Cells("VOUCHER_NO").Text = "" Then
                    .ActiveRow.Cells("VOUCHER_NO").Value = Absx1.CtlFor("VOUCHER_NO").Text
                    .ActiveRow.Cells("VOUCHER_LNO").Value = Val(dst.Tables("APTINVH2").Compute("Max(VOUCHER_LNO)", "") & "") + 1
                End If
            End If
        End With

    End Sub

    Private Sub grdAPTINVH2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH2.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdAPTINVH2, sql_where, sql_where <> "")
    End Sub

    Private Sub numDIST_GL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_GL.ValueChanged
        Calc_Totals()
    End Sub

    Private Sub numDIST_OTHER_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_OTHER.ValueChanged
        Calc_Totals()
    End Sub

    Private Sub numDIST_PO_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_PO.ValueChanged
        Calc_Totals()
    End Sub

    Sub Calc_Totals()
        If disable_calculate_totals Then Exit Sub

        numINV_AMT.Value = Val(numINV_AMT_VEND.Value & "") + Val(Absx1.numFor("INV_ADJUSTMENTS").Value & "") '  - Val(Absx1.numFor("INV_ALLOWANCES").Value & "")

        numDIST_OOBAL.Value = Val(numINV_AMT.Value & "") _
                            - Val(numDIST_GL.Value & "") _
                            - Val(numDIST_PO.Value & "") _
                            + Val(Absx1.numFor("INV_ALLOWANCES").Value & "") _
                            - Val(numDIST_OTHER.Value & "")

        '                            - Val(Absx1.numFor("INV_ADJUSTMENTS").Value & "") _

        If Abs(Val(numINV_AMT.Value & "") - Val(numINV_AMT_VEND.Value & "")) > 0.01 Then
            numINV_AMT.Appearance.ForeColor = Drawing.Color.Red
            Absx1.numFor("INV_ADJUSTMENTS").Appearance.ForeColor = Drawing.Color.Red
        Else
            numINV_AMT.Appearance.ForeColor = Drawing.Color.Empty
            Absx1.numFor("INV_ADJUSTMENTS").Appearance.ForeColor = Drawing.Color.Empty
        End If

    End Sub

    Private Sub numINV_AMT_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles numINV_AMT.Leave
        Automatic_Distribution()
    End Sub

    Private Sub numINV_AMT_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numINV_AMT.ValueChanged
        Calc_Totals()
    End Sub

    Sub Calc_DIST_GL()
        Dim DIST_GL As Decimal = Val(dst.Tables("APTINVH2").Compute("SUM(INV_LINE_AMT)", "INV_LTYP is Null or INV_LTYP = 'A'") & "")
        numDIST_GL.Value = DIST_GL
        Calc_Totals()
    End Sub

    Sub Calc_DIST_PO()
        Calc_DIST_Adjustments()
        Dim DIST_PO As Decimal = Val(dst.Tables("APTINVH5_SUM").Compute("SUM(AMT_REC)", "") & "")
        Absx1.numFor("DIST_PO").Value = System.Math.Round(DIST_PO, 2, MidpointRounding.AwayFromZero)
        Calc_Totals()
    End Sub

    Sub Calc_DIST_Adjustments()
        Dim INV_ADJUSTMENTS As Decimal = -1 * Val(dst.Tables("APTINVH5").Compute("SUM(AMT_VAR)", "ISNULL(CB,'0') = '1'") & "") + 1 * Val(dst.Tables("APTINVH8").Compute("SUM(VOUCHER_ADJ_AMT)", "") & "")
        Absx1.numFor("INV_ADJUSTMENTS").Value = INV_ADJUSTMENTS
        Dim INV_ALLOWANCES As Decimal = -1 * Val(dst.Tables("APTINVH5").Compute("SUM(AMT_VAR)", "ISNULL(CB,'0') = '0'") & "")
        Absx1.numFor("INV_ALLOWANCES").Value = INV_ALLOWANCES
        Calc_Totals()
    End Sub

    Sub Calc_DIST_Other()
        Dim DIST_OTHER As Decimal = Val(dst.Tables("APTINVH7").Compute("SUM(TOTAL_INV)", "") & "")
        Absx1.numFor("DIST_OTHER").Value = DIST_OTHER
        Calc_Totals()
    End Sub

    Private Sub chkINV_1099_IND_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkINV_1099_IND.CheckedChanged
        numINV_1099_AMT.ReadOnly = Not chkINV_1099_IND.Checked Or InquiryMode
        If Not chkINV_1099_IND.Checked Then
            numINV_1099_AMT.Value = 0
        End If
    End Sub

    Private Sub chkEXP_CATGY_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEXP_CATGY_CODE.CheckedChanged
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("EXP_CATGY_CODE").Hidden = Not chkEXP_CATGY_CODE.Checked
    End Sub

    Private Sub chkDIST_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDIST_CODE.CheckedChanged
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("DIST_CODE").Hidden = Not chkDIST_CODE.Checked
    End Sub

    Private Sub optAPTINVH2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAPTINVH2.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Distribution_Display_Options()
    End Sub

    Sub Set_Distribution_Display_Options()

        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("INV_LTYP").Hidden = Not (optAPTINVH2.Value = "Manual Entries")
        grdAPTINVH2.DisplayLayout.Bands("APTINVH2").Columns("INV_DLNO").Hidden = Not (optAPTINVH2.Value = "Manual Entries")

        Dim dvw As DataView = DirectCast(grdAPTINVH2.DataSource, DataTable).DefaultView
        If optAPTINVH2.Value = "Manual Entries" Then
            dvw.RowFilter = "INV_LTYP is Null or INV_LTYP = 'A'"
        Else
            dvw.RowFilter = ""
        End If

    End Sub
    Private Sub dteINV_DATE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dteINV_DATE.ValueChanged
        Calculate_INV_DUE_DATE()
    End Sub

    Private Sub dteINV_BL_DATE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dteINV_BL_DATE.ValueChanged
        Calculate_INV_DUE_DATE()
    End Sub

    Sub Calculate_INV_DUE_DATE()
        If EntryMode = "" Then
            Exit Sub
        End If

        Dim INV_DATE As Object = Absx1.dteFor("INV_DATE").Value
        Dim INV_BL_DATE As Object = Absx1.dteFor("INV_BL_DATE").Value
        Dim INV_BASE_DATE As Object = Nothing

        If INV_BL_DATE & "" <> "" And Not Absx1.chkFor("VEND_DUE_FROM_INV_DATE").Checked Then
            INV_BASE_DATE = INV_BL_DATE
        Else
            INV_BASE_DATE = INV_DATE
        End If
        If INV_BASE_DATE Is Nothing Then Exit Sub
        If Absx1.txtFor("TERM_CODE").Text = "" Then
            Absx1.dteFor("INV_DUE_DATE").Value = DBNull.Value
        Else
            Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, Absx1.txtFor("TERM_CODE").Text, Nothing, INV_BASE_DATE)
            Absx1.dteFor("INV_DUE_DATE").Value = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, Absx1.txtFor("TERM_CODE").Text, Nothing, INV_BASE_DATE)
        End If

        Calculate_INV_DISC_AMT(True)
    End Sub

    Sub Generate_Pre_Distribution()
        Dim VOUCHER_LNO_ctr As Integer = 0
        Dim DIST_AMT As Decimal

        If ASCMAIN1.CLIENT = "VAN" Then ' PROB OPEN UP TO ALL
            If dst.Tables("ICTIREC1").Rows.Count <> 0 Then
                Exit Sub
            End If
        End If

        ASCMAIN1.sql = "Select * from APTVEND9 where VEND_CODE = '" & HFs("VEND_CODE") & "'"
        Dim tblAPTVEND9 As DataTable = ASCDATA1.GetDataTable
        If Absx1.cbeFor("INV_TYPE").Value = "A" Then ' Absx1.CtlFor("INV_TYPE").Text = "A" Then
            tblAPTVEND9.Rows.Clear()
            Dim rowAPTVEND9 As DataRow = tblAPTVEND9.NewRow
            rowAPTVEND9.Item("VEND_CODE") = HFs("VEND_CODE")
            rowAPTVEND9.Item("ACCT_CODE") = ROWs("APTPARM1").Item("AP_PARM_ACCT_CODE_ADVANCES")
            rowAPTVEND9.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowAPTVEND9.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowAPTVEND9.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowAPTVEND9.Item("DIST_AMT") = 100
            tblAPTVEND9.Rows.Add(rowAPTVEND9)
        Else
            If tblAPTVEND9.Rows.Count = 0 Then
                If rowAPTVEND1.Item("ACCT_CODE") & "" <> "" Then
                    Dim rowAPTVEND9 As DataRow = tblAPTVEND9.NewRow
                    rowAPTVEND9.Item("VEND_CODE") = HFs("VEND_CODE")
                    rowAPTVEND9.Item("ACCT_CODE") = rowAPTVEND1.Item("ACCT_CODE")
                    rowAPTVEND9.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowAPTVEND9.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    rowAPTVEND9.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    If rowAPTVEND1.Item("VEND_PRE_DIST_TYPE") & "" = "A" Then
                        DIST_AMT = Val(numINV_AMT.Value & "")
                    Else
                        DIST_AMT = 100
                    End If
                    rowAPTVEND9.Item("DIST_AMT") = DIST_AMT
                    tblAPTVEND9.Rows.Add(rowAPTVEND9)
                End If
            End If
        End If

        For Each rowAPTVEND9 As DataRow In tblAPTVEND9.Rows
            If rowAPTVEND1.Item("VEND_PRE_DIST_TYPE") & "" = "A" And Absx1.cbeFor("INV_TYPE").Value <> "A" Then 'Absx1.CtlFor("INV_TYPE").Text <> "A" Then
                DIST_AMT = Val(rowAPTVEND9.Item("DIST_AMT") & "")
            Else
                DIST_AMT = Val(rowAPTVEND9.Item("DIST_AMT") & "") * Val(numINV_AMT.Value & "") / 100
            End If
            Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
            rowAPTINVH2.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            VOUCHER_LNO_ctr += 1
            rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
            rowAPTINVH2.Item("ACCT_CODE") = rowAPTVEND9.Item("ACCT_CODE")
            rowAPTINVH2.Item("SEG2_CODE") = rowAPTVEND9.Item("SEG2_CODE")
            rowAPTINVH2.Item("SEG3_CODE") = rowAPTVEND9.Item("SEG3_CODE")
            rowAPTINVH2.Item("SEG4_CODE") = rowAPTVEND9.Item("SEG4_CODE")
            rowAPTINVH2.Item("ACCT_DESC") = LookUp("GLTACCT1", rowAPTINVH2.Item("ACCT_CODE") & "", True).ITEM("ACCT_DESC")
            rowAPTINVH2.Item("INV_LINE_AMT") = DIST_AMT
            If Absx1.cbeFor("INV_TYPE").Value = "A" Then
                rowAPTINVH2.Item("INV_LTYP") = "A"
            End If
            dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        Next
        Calc_DIST_GL()
    End Sub

    Private Sub chkACCRUE_PRIOR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkACCRUE_PRIOR.CheckedChanged
        Setup_OPS_YYYYPP_ACCRUE()
    End Sub

    Sub Setup_OPS_YYYYPP_ACCRUE()
        'Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Visible = True
        Absx1.txtFor("OPS_YYYYPP_ACCRUE").Visible = True
        If Not chkACCRUE_PRIOR.Checked Then
            'Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Text = ""
            Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text = ""
            'Absx1.cmbFor("OPS_YYYYPP_ACCRUE").Visible = False
            Absx1.txtFor("OPS_YYYYPP_ACCRUE").Visible = False
        Else
            If Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text = "" Then
                If rowAPTINVH1.RowState = DataRowState.Modified Then
                    If rowAPTINVH1.Item("OPS_YYYYPP_ACCRUE", DataRowVersion.Original) & "" <> "" Then
                        Absx1.txtFor("OPS_YYYYPP_ACCRUE").Text = rowAPTINVH1.Item("OPS_YYYYPP_ACCRUE", DataRowVersion.Original) & ""
                    End If
                End If
            End If
        End If
    End Sub

    Sub Load_OPS_YYYYPP_ACCRUE()
        If ROWs("APTPARM1").Item("AP_PARM_ALLOW_ACCRUAL") & "" = "0" Then
            If chkACCRUE_PRIOR.Checked = False Then
                Setup_OPS_YYYYPP_ACCRUE()
            Else
                chkACCRUE_PRIOR.Checked = False
            End If
            chkACCRUE_PRIOR.Visible = False
        Else
            chkACCRUE_PRIOR.Visible = True
        End If
    End Sub

    Sub Load_CURR_EXCH_RATE()
        If Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
            Absx1.numFor("CURR_EXCH_RATE").Value = 1
            Absx1.numFor("CURR_EXCH_RATE").ReadOnly = True Or InquiryMode
        Else
            ' USE STD RATE? - IFSO - ONLY ON LOADING NEW VOUCHER
            Absx1.numFor("CURR_EXCH_RATE").ReadOnly = True Or InquiryMode ' False
        End If
    End Sub

    Sub Setup_INV_BL_DATE()
        'If dst.Tables("APTINVH5_SUM").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
        '    Absx1.dteFor("INV_BL_DATE").ReadOnly = False
        'Else
        '    Absx1.dteFor("INV_BL_DATE").ReadOnly = True
        '    Absx1.dteFor("INV_BL_DATE").Value = dst.Tables("APTINVH5_SUM").Compute("MIN(RECEIPT_DATE)", "")
        'End If
    End Sub

    Sub Negate_Voucher(ByVal VOUCHER_NO As String)
        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            For Each COLUMN_NAME As String In New String() _
            {"INV_AMT", "INV_AMT_VEND", "INV_DISC_BASED_ON", "INV_DISC_AMT", "INV_BALANCE", "INV_1099_AMT"}
                rowAPTINVH1.Item(COLUMN_NAME) = -1 * Val(rowAPTINVH1.Item(COLUMN_NAME) & "")
            Next
        Next

        For Each rowAPTINVH2 As DataRow In dst.Tables("APTINVH2") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH2.Item("INV_LINE_AMT") = -1 * Val(rowAPTINVH2.Item("INV_LINE_AMT") & "")
        Next

        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH5.Item("INV_QTY") = -1 * Val(rowAPTINVH5.Item("INV_QTY") & "")
        Next

        For Each rowAPTINVH8 As DataRow In dst.Tables("APTINVH8") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH8.Item("VOUCHER_ADJ_AMT") = -1 * Val(rowAPTINVH8.Item("VOUCHER_ADJ_AMT") & "")
        Next

        For Each rowAPTINVH7 As DataRow In dst.Tables("APTINVH7") _
            .Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
            rowAPTINVH7.Item("TOTAL_INV") = -1 * Val(rowAPTINVH7.Item("TOTAL_INV") & "")
        Next

    End Sub

    Sub ReNumber_Voucher(ByVal VOUCHER_NO_old As String, ByVal VOUCHER_NO_new As String)
        dst.EnforceConstraints = False
        For Each TABLE_NAME As String In New String() _
            {"APTINVH1", "APTINVH2", "APTINVH5", "APTINVH8", "APTINVH7"}
            dst.Tables(TABLE_NAME).AcceptChanges() ' IF YOU DON'T DO THIS, THEN THE DELETED ROWS WILL REMAIN IN THE TABLE AND BE DELETED FROM THE VOUCHER YOU ARE RENUMBERING FROM 
            ReNumber_Voucher_1(TABLE_NAME, VOUCHER_NO_old, VOUCHER_NO_new)
        Next
        dst.EnforceConstraints = True
    End Sub

    Sub ReNumber_Voucher_1(ByVal TABLE_NAME As String,
    ByVal VOUCHER_NO_old As String,
    ByVal VOUCHER_NO_new As String)
        For Each row As DataRow In dst.Tables(TABLE_NAME).Select("VOUCHER_NO = '" & VOUCHER_NO_old & "'", "")
            row.Item("VOUCHER_NO") = VOUCHER_NO_new
            If TABLE_NAME = "APTINVH1" Then
                row.Item("VOUCHER_NO_ORIG") = VOUCHER_NO_old
            End If
            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub

    Function Check_Invoice() As String
        ASCMAIN1.sql = "Select * from APTINVH1 " _
        & " where VEND_CODE = :PARM1" _
        & "   and INV_NUM = :PARM2" _
        & "   and INV_TYPE = :PARM3" _
        & "   and INV_STATUS IN ('O','H','P')"

        If Absx1.txtFor("VOUCHER_NO").Text <> "" Then
            ASCMAIN1.sql &= "   and VOUCHER_NO <> '" & Absx1.txtFor("VOUCHER_NO").Text & "'"
        End If

        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New Object() {Absx1.txtFor("VEND_CODE").Text, Absx1.txtFor("INV_NUM").Text, Absx1.cbeFor("INV_TYPE").Value})

        If row IsNot Nothing And Not batch_update Then
            If "YES" <> InputBox("Invoice Number " & Absx1.txtFor("INV_NUM").Text & " has Already been Entered" & vbCrLf & "(Voucher " & row.Item("VOUCHER_NO") & ")" & vbCrLf & vbCrLf & "Enter YES to Proceed", "Enter YES to Proceed") Then
                Check_Invoice = "NO"
                Exit Function
            Else
                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    MsgBox("Sorry, You cannot enter the same invoice twice", MsgBoxStyle.OkOnly, "Jim Says NO")
                    Check_Invoice = "NO"
                    Exit Function
                End If
            End If
        End If
        Check_Invoice = "YES"
    End Function

    Sub Load_Alternate_Payment_Address()
        LookUp("APTVEND2", New String() {HFs("VEND_CODE"), Absx1.txtFor("VEND_ALT_CODE").Text})
        If cdr IsNot Nothing Then
            Dim rowAPTVEND2 As DataRow = dst.Tables("APTVEND2").NewRow
            rowAPTVEND2.ItemArray = cdr.ItemArray
            dst.Tables("APTVEND2").Rows.Clear()
            dst.Tables("APTVEND2").Rows.Add(rowAPTVEND2)

            'For i As Integer = 0 To .Columns.Count - 1
            '    Dim COLUMN_NAME As String = .Columns(i).ColumnName
            '    If COLUMN_NAME <> "VEND_ALT_CODE" Then
            '        COLUMN_NAME = Replace(COLUMN_NAME, "_ALT_", "_")
            '        rowAPTVEND2.Item(i) = rowAPTVEND1.Item(COLUMN_NAME)
            '    End If
            'Next

        End If
    End Sub


    Private Sub numDIST_OOBAL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numDIST_OOBAL.ValueChanged
        grdAPTINVH2.DisplayLayout.Bands(0).SummaryFooterCaption = "Voucher Totals; Total Un-Distributed Amount = " & Format(numDIST_OOBAL.Value, "$#,##0.00") & "; Double-Click the Amount Cell to Auto-Balance"
    End Sub

    Private Sub grdAPTINVR1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTINVR1.DoubleClickRow
        If grdAPTINVR1.ActiveRow IsNot Nothing Then
            Absx1.txtFor("VOUCHER_NO").Text = grdAPTINVR1.ActiveRow.Cells("VOUCHER_NO").Text
            Click_Command("Edit")
        End If

    End Sub

    Private Sub grdAPTINVH2_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdAPTINVH2.DoubleClickCell
        If e.Cell.Column.Key = "INV_LINE_AMT" Then
            If grdAPTINVH2.ActiveCell Is Nothing Then
            Else
                grdAPTINVH2.ActiveCell.Value = Val(grdAPTINVH2.ActiveCell.Value & "") + Val(numDIST_OOBAL.Value & "")
                grdAPTINVH2.UpdateData()
            End If
        End If
    End Sub

    Private Sub grpMode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub chkQuickEntry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkQuickEntry.CheckedChanged
        Setup_QE(chkQuickEntry.Checked)
    End Sub

    Sub Setup_QE(ByVal tf As Boolean)
        lblQE_INV_DATE.Visible = tf
        lblQE_INV_AMT.Visible = tf
        lblQE_INV_REF.Visible = tf
        Absx1.CtlFor("QE_INV_DATE").Visible = tf
        Absx1.CtlFor("QE_INV_AMT").Visible = tf
        Absx1.CtlFor("QE_INV_REF").Visible = tf

        lblINV_TYPE.Visible = Not tf
        Absx1.CtlFor("INV_TYPE").Visible = Not tf
        lblVOUCHER_NO.Visible = Not tf
        Absx1.CtlFor("VOUCHER_NO").Visible = Not tf
    End Sub

    Sub ReLoad_Voucher(ByVal VOUCHER_NO As String)
        EnforceConstraints(False)
        Fill_Records("APTINVH1", VOUCHER_NO)
        Fill_Records("APTINVH2", VOUCHER_NO)
        'Fill_Records("APTINVH5", VOUCHER_NO)
        Fill_Records("APTINVH8", VOUCHER_NO)
        Fill_Records("APTINVH7", VOUCHER_NO)
        EnforceConstraints(True)
    End Sub

    Sub Dependent_Updates(ByVal VOUCHER_NO As String, ByVal reverse As Boolean)
        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare " & vbCrLf _
            & "  Cursor C1 is" & vbCrLf _
            & "   Select Distinct RECEIPT_NO from APTINVH5" & vbCrLf _
            & "    where VOUCHER_NO = '" & VOUCHER_NO & "'" & vbCrLf _
            & "      and RECEIPT_NO Is Not Null;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTIREC1 set " & vbCrLf _
            & IIf(reverse,
                "VOUCHER_NO = NULL, ACCRUAL_STATUS = '0'",
                "VOUCHER_NO = '" & VOUCHER_NO & "', ACCRUAL_STATUS = '1'") & vbCrLf _
            & "    where RECEIPT_NO = R1.RECEIPT_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare " & vbCrLf _
            & "  Cursor C1 is" & vbCrLf _
            & "   Select Distinct PO_SHIPMENT_NO,PO_SHIPMENT_LNO from APTINVH5" & vbCrLf _
            & "    where VOUCHER_NO = '" & VOUCHER_NO & "'" & vbCrLf _
            & "      and RECEIPT_NO Is Null;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update POTSHIP2 set " & vbCrLf _
            & IIf(reverse,
                "VOUCHER_NO = NULL, ACCRUAL_STATUS = '0'",
                "VOUCHER_NO = '" & VOUCHER_NO & "', ACCRUAL_STATUS = '1'") & vbCrLf _
            & "    where PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO" & vbCrLf _
            & "      and PO_SHIPMENT_LNO = R1.PO_SHIPMENT_LNO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare " & vbCrLf _
            & "  Cursor C1 is Select * from APTINVH7 where VOUCHER_NO = '" & VOUCHER_NO & "';" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update POTLCST1 set " & vbCrLf _
            & IIf(reverse,
                    "VOUCHER_NO = NULL, COST_ACT = 0",
                    "VOUCHER_NO = R1.VOUCHER_NO, COST_ACT = R1.TOTAL_INV") _
            & "    where CTL_NO = R1.CTL_NO;" _
            & "  End Loop; " _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        If reverse Then
            '? - PROBABLY SHOULD WRITE OUT A NEGATIVE RECORD, AND CLOSE OUT AGAINST POSITIVE RECORD IF IT IS STILL AVAILABLE
            ' DON'T LIKE THE DELETE CONCEPT ON POTLCST2
            ' BUT KEY DOES NOT SUPPORT NEGATIVE RECORDS

            ASCMAIN1.sql = "Select APTINVH7.CTL_NO, POTLCST1.PO_SHIPMENT_NO, POTLCST1.PO_ORDER_NO" & vbCrLf _
                & " from APTINVH7,POTLCST1" & vbCrLf _
                & " where APTINVH7.VOUCHER_NO = '" & VOUCHER_NO & "'" & vbCrLf _
                & "   and POTLCST1.CTL_NO = APTINVH7.CTL_NO"

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim CTL_NO As String = row.Item("CTL_NO")
                ASCMAIN1.sql = "Delete from POTLCST2 where CTL_NO = '" & CTL_NO & "'" _
                    & " and INV_NO is Null"
                ' protecting POTLCST2 records which were used in a Debit Memo
                ' - this is the ugly part of deleting these records - it's almost as if we need to write out a negative potlcst2 record now to promote a customer credit memo
                ASCDATA1.ExecuteSQL()
            Next
        Else

            ' HOW COULD THIS HAVE WORKED - POTLCST2 IS NOT EVEN WRITTEN OUT YET?
            'ASCMAIN1.sql = "Select APTINVH7.CTL_NO, POTLCST1.PO_SHIPMENT_NO, MAX(POTLCST2.PO_ORDER_NO) PO_ORDER_NO" & vbCrLf _
            '    & " from APTINVH7,POTLCST2,POTLCST1" & vbCrLf _
            '    & " where APTINVH7.VOUCHER_NO = '" & VOUCHER_NO & "'" & vbCrLf _
            '    & "   and POTLCST1.CTL_NO = APTINVH7.CTL_NO" & vbCrLf _
            '    & "   and POTLCST2.CTL_NO (+) = APTINVH7.CTL_NO" & vbCrLf _
            '    & " group by APTINVH7.CTL_NO, POTLCST1.PO_SHIPMENT_NO"

            ASCMAIN1.sql = "Select APTINVH7.CTL_NO, POTLCST1.PO_SHIPMENT_NO, POTLCST1.PO_ORDER_NO" & vbCrLf _
                & " from APTINVH7,POTLCST1" & vbCrLf _
                & " where APTINVH7.VOUCHER_NO = '" & VOUCHER_NO & "'" & vbCrLf _
                & "   and POTLCST1.CTL_NO = APTINVH7.CTL_NO"

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim CTL_NO As String = row.Item("CTL_NO")
                Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO") & ""
                Dim PO_SHIPMENT_NO As String = row.Item("PO_SHIPMENT_NO") & ""
                If PO_SHIPMENT_NO = "" Then ' PO_ORDER_NO <> "" Then

                    ASCMAIN1.sql = "Insert into POTLCST2 (CTL_NO, PO_ORDER_NO, CHARGEBACK_STATUS, OPS_YYYYPP, COST_ACT_PO)" & vbCrLf _
                        & " Select POTLCST1.CTL_NO, POTLCST1.PO_ORDER_NO" & vbCrLf _
                        & ", DECODE(POTORDR1.ORDR_NO,NULL,'0',NVL(POTCATG1.CHARGEBACK_IND,'0')) CHARGEBACK_STATUS" & vbCrLf _
                        & ", POTLCST1.OPS_YYYYPP, POTLCST1.COST_ACT" & vbCrLf _
                        & " from POTLCST1,POTORDR1,POTCATG1" & vbCrLf _
                        & " where POTCATG1.COST_CATGY_CODE = POTLCST1.COST_CATGY_CODE" & vbCrLf _
                        & "   and POTORDR1.PO_ORDER_NO = POTLCST1.PO_ORDER_NO" & vbCrLf _
                        & "   and POTLCST1.CTL_NO = '" & CTL_NO & "'"
                    ASCDATA1.ExecuteSQL()

                Else

                    Dim SQL As String = ""
                    SQL &= "" _
                        & "Select POTLCST1.CTL_NO, POTLCST1.CHARGEBACK_IND, POTLCST1.OPS_YYYYPP, POTSHIP3.PO_ORDER_NO, POTORDR1.ORDR_NO" & vbCrLf _
                        & ", POTLCST1.COST_ACT COST_ACT_PO " & vbCrLf _
                        & ", SUM(POTLCST1.COST_ACT * POTSHIP3.WEIGHT_FACTOR * POTSHIP3.PO_QTY_SHP) COST_ACT_PO_EXT " & vbCrLf _
                        & ", SUM(POTSHIP3.WEIGHT_FACTOR * POTSHIP3.PO_QTY_SHP) WEIGHT " & vbCrLf _
                        & " from POTSHIP3,POTLCST1,POTORDR1" & vbCrLf _
                        & " where POTLCST1.CTL_NO = '" & CTL_NO & "'" & vbCrLf _
                        & IIf(PO_SHIPMENT_NO <> "",
                              "" _
                              & "   and POTSHIP3.PO_SHIPMENT_NO = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
                              & "   and (POTSHIP3.PO_SHIPMENT_LNO = POTLCST1.PO_SHIPMENT_LNO OR NVL(POTLCST1.PO_SHIPMENT_LNO,0) = 0)" & vbCrLf,
                              "" _
                              & "   and POTSHIP3.PO_ORDER_NO = POTLCST1.PO_ORDER_NO" & vbCrLf) _
                        & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                        & " group by POTLCST1.CTL_NO, POTLCST1.CHARGEBACK_IND, POTLCST1.OPS_YYYYPP, POTSHIP3.PO_ORDER_NO, POTORDR1.ORDR_NO, POTLCST1.COST_ACT"

                    ASCMAIN1.sql = "Select Sum (WEIGHT) WEIGHT, Count (*) RECS from (" & SQL & ")"
                    Dim rowT As DataRow = ASCDATA1.GetDataRow
                    Dim RECS As Decimal = Val(rowT.Item("RECS") & "")
                    If RECS <> 0 Then
                        Dim WEIGHT As Decimal = Val(rowT.Item("WEIGHT") & "")
                        Dim COST_CALC As String = "COST_ACT_PO_EXT " & " / " & CStr(WEIGHT)
                        If WEIGHT = 0 Or RECS = 1 Then COST_CALC = "COST_ACT_PO " & " / " & CStr(RECS)
                        ASCMAIN1.sql = "Insert into POTLCST2 (CTL_NO, PO_ORDER_NO, CHARGEBACK_STATUS, OPS_YYYYPP, COST_ACT_PO)" & vbCrLf _
                            & " Select CTL_NO, PO_ORDER_NO, DECODE(ORDR_NO,NULL,'0',NVL(CHARGEBACK_IND,'0')) CHARGEBACK_STATUS, OPS_YYYYPP, " & COST_CALC & vbCrLf _
                            & " from (" & SQL & ")"
                        ASCDATA1.ExecuteSQL()
                    End If
                End If
            Next
        End If

        Dim INV_AMT As Decimal = Val(rowAPTINVH1("INV_AMT") & "")
        If reverse Then
            Dim row As DataRow = LookUp("APTINVH1", VOUCHER_NO)
            INV_AMT = -1 * Val(row("INV_AMT") & "")
        End If
        'rowAPTVEND5 = Fill_Record("APTVEND5", HFs("VEND_CODE"), True)

        rowAPTVEND5.Item("VEND_PURCHASES_MTD") = Val(rowAPTVEND5.Item("VEND_PURCHASES_MTD") & "") + INV_AMT
        rowAPTVEND5.Item("VEND_PURCHASES_YTD") = Val(rowAPTVEND5.Item("VEND_PURCHASES_YTD") & "") + INV_AMT
        rowAPTVEND5.Item("VEND_NUM_INV_MTD") = Val(rowAPTVEND5.Item("VEND_NUM_INV_MTD") & "") + IIf(reverse, -1, 1)
        rowAPTVEND5.Item("VEND_NUM_INV_YTD") = Val(rowAPTVEND5.Item("VEND_NUM_INV_YTD") & "") + IIf(reverse, -1, 1)
        If Not reverse Then
            rowAPTVEND5.Item("VEND_LAST_INV_DATE") = rowAPTINVH1("INV_DATE")
            rowAPTVEND5.Item("VEND_LAST_INV_AMT") = rowAPTINVH1("INV_AMT")
            rowAPTVEND5.Item("VEND_LAST_INV_NUM") = rowAPTINVH1("INV_NUM")
            If rowAPTVEND5.Item("VEND_1ST_PURCH_DATE") & "" = "" Then
                rowAPTVEND5.Item("VEND_1ST_PURCH_DATE") = rowAPTINVH1("INV_DATE")
            End If
        End If
    End Sub

    Private Sub chkRecurring_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRecurring.CheckedChanged
        Show_Batch()
        'optstatus will be R only and invisible
    End Sub

    Sub Show_Batch()

        If dst.Tables.Count = 0 Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()

        Dim sql As String = "from APTINVH1,APTVEND1 where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE"
        Dim caption As String = ""

        If chkRecurring.Checked Then
            sql &= " and APTINVH1.INV_STATUS = 'R'"
            caption = "Recurring Invoice Templates"
        Else
            sql &= " and APTINVH1.REGISTER_IND = '0'"
            caption = "Invoices Pending Invoice Register Update"
        End If

        If optFilter.Value = "R" Then
            sql &= " and APTINVH1.INIT_OPER = '" & ASCMAIN1.USER_ID & "'"
            caption &= ", entered by " & ASCMAIN1.USER_ID
        ElseIf optFilter.Value = "V" Then
            sql &= " and APTVEND1.PROCESSOR_CODE = '" & ASCMAIN1.USER_ID & "'"
            caption &= ", for processor " & ASCMAIN1.USER_ID
        End If

        EnforceConstraints(False)
        sql_APTINVR2 = "Select APTINVH2.*,GLTACCT1.ACCT_DESC from APTINVH2,GLTACCT1 where GLTACCT1.ACCT_CODE = APTINVH2.ACCT_CODE and APTINVH2.VOUCHER_NO in (Select APTINVH1.VOUCHER_NO " & sql & ")"
        sql_APTVEND1 = "Select APTVEND1.* from APTVEND1 where APTVEND1.VEND_CODE in (Select DISTINCT APTINVH1.VEND_CODE " & sql & ")"
        sql = "Select APTINVH1.* " & sql
        grdAPTINVR1.Text = caption
        Fill_Records("APTINVR1", "", True, sql)
        Fill_Records("APTINVR2", "", , sql_APTINVR2)
        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
    End Sub

    Sub Set_Recurring(ByVal Recurring As Boolean)
        'tabHeader.Tabs("Recurring").Enabled = (optINV_STATUS.Value = "R")
        grpINV_STATUS.Visible = Not Recurring
    End Sub

    Private Sub UltraTabPageControl2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles UltraTabPageControl2.Paint

    End Sub

    Sub Calculate_INV_DISC_AMT(Optional ByVal update_UI As Boolean = False)

        'Me.Validate()
        If Absx1.dteFor("INV_DATE").Value & "" = "" Then Exit Sub

        Dim TERM_CODE As String
        If update_UI Then
            TERM_CODE = Absx1.txtFor("TERM_CODE").Text
        Else
            TERM_CODE = rowAPTINVH1("TERM_CODE")
        End If
        Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE, True)

        Dim INV_DISC_BASED_ON As Double = Val(rowAPTINVH1("INV_AMT") & "")
        If update_UI Then
            INV_DISC_BASED_ON = Val(Absx1.numFor("INV_AMT").Value & "")
            'Absx1.numFor("INV_DISC_BASED_ON").Value = INV_DISC_BASED_ON
        Else
            rowAPTINVH1("INV_DISC_BASED_ON") = INV_DISC_BASED_ON
        End If

        Dim TERM_DAYS_DISC As Double = Val(rowTATTERM1("TERM_DAYS_DISC") & "")
        Dim TERM_DISC_PERC As Double = Val(rowTATTERM1("TERM_DISC_PERC") & "")

        If TERM_DISC_PERC = 0 Or HFs("INV_TYPE") <> "I" Then
            If update_UI Then
                Absx1.dteFor("INV_DISC_DUE").Value = Null
                Absx1.numFor("INV_DISC_AMT").Value = 0
            Else
                rowAPTINVH1("INV_DISC_DUE") = Null
                rowAPTINVH1("INV_DISC_AMT") = 0
            End If
        Else
            If rowTATTERM1("TERM_DISC_ELIG_DUE") & "" = "1" Then
                If update_UI Then
                    Absx1.dteFor("INV_DISC_DUE").Value = Absx1.dteFor("INV_DUE_DATE").Value
                Else
                    rowAPTINVH1("INV_DISC_DUE") = rowAPTINVH1("INV_DUE_DATE")
                End If
            Else
                If update_UI Then
                    Absx1.dteFor("INV_DISC_DUE").Value = DateValue(Absx1.dteFor("INV_DATE").Value).AddDays(TERM_DAYS_DISC)
                Else
                    rowAPTINVH1("INV_DISC_DUE") = DateValue(rowAPTINVH1("INV_DATE")).AddDays(TERM_DAYS_DISC)
                End If
            End If
            If update_UI Then
                Absx1.numFor("INV_DISC_AMT").Value = Round(INV_DISC_BASED_ON * TERM_DISC_PERC / 100, 2)
            Else
                rowAPTINVH1("INV_DISC_AMT") = Round(INV_DISC_BASED_ON * TERM_DISC_PERC / 100, 2)
            End If
        End If
    End Sub

    Private Sub numQE_INV_AMT_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles numQE_INV_AMT.Enter
        'numQE_INV_AMT.SelectAll()
    End Sub

    Private Sub numQE_INV_AMT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles numQE_INV_AMT.GotFocus
        numQE_INV_AMT.SelectAll()
    End Sub

    Sub Print_Record()
        Print_Report_Begin()

        'Fill_Records("APTINVR2", "", , sql_APTINVR2)
        Fill_Records("APTVEND1", "", , sql_APTVEND1)

        Dim SUBT As String = "Edit List"
        'CR_params.Add("REPORT_PARAMETER_NAME", "VALUE")
        'RecordSelectionFormula = "{EDTSLSVP.SLS} > 10"
        Generate_Report("APRINVR1", "Vendor Invoice Entry", SUBT)

        Print_Report_End()

        dst.Tables("APTVEND1").Rows.Clear()
    End Sub

    Private Sub UltraNumericEditor4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraNumericEditor4.ValueChanged

    End Sub

    Private Sub optFilter_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optFilter.ValueChanged
        Show_Batch()
    End Sub

    Sub Update_as_Paid(ByVal VOUCHER_NO As String)

        rowAPTINVH1 = Fill_Record("APTINVH1", VOUCHER_NO)
        Dim rowAPTINVHX As DataRow = dst.Tables("APTINVHX").NewRow
        rowAPTINVHX("VOUCHER_NO") = rowAPTINVH1("VOUCHER_NO")
        rowAPTINVHX("INV_NUM") = rowAPTINVH1("INV_NUM")
        rowAPTINVHX("INV_DATE") = rowAPTINVH1("INV_DATE")
        rowAPTINVHX("INV_BALANCE") = rowAPTINVH1("INV_BALANCE")
        rowAPTINVHX("INV_DISC_AMT") = rowAPTINVH1("INV_DISC_AMT")
        rowAPTINVHX("SELECTED") = "1"
        dst.Tables("APTINVHX").Rows.Add(rowAPTINVHX)

        Dim VEND_CODE As String = rowAPTINVH1("VEND_CODE") & ""
        Dim VEND_CODE_AP As String = rowAPTINVH1("VEND_CODE_AP") & ""
        If VEND_CODE_AP = "" Then
            VEND_CODE_AP = VEND_CODE
        End If
        Dim VEND_ALT_CODE As String = rowAPTINVH1("VEND_ALT_CODE") & ""

        Dim VEND_NAME As String
        Dim rowPayee As DataRow
        If VEND_CODE_AP <> "" And VEND_CODE_AP <> VEND_CODE Then
            rowPayee = LookUp("APTVEND1", VEND_CODE_AP)
            VEND_NAME = rowPayee.Item("VEND_NAME")
        Else
            VEND_NAME = HFs("VEND_NAME")
        End If

        Dim BANK_CODE As String = rowAPTINVH1("BANK_CODE")
        Dim PYMT_METHOD As String = rowAPTINVH1("INV_PYMT_METHOD")
        Dim CHECK_NUM As String = rowAPTINVH1("CHECK_NUM")
        Dim CHECK_DATE As Date = rowAPTINVH1("CHECK_DATE")
        Dim CHECK_AMT As Double = Absx1.numFor("CHECK_AMT").Value

        Dim SEQ_NUM As Integer
        SEQ_NUM = 0
        For Each rowAPTINVHX In dst.Tables("APTINVHX").Select("SELECTED = '1'", "")
            If rowAPTINVHX("VOUCHER_NO") = VOUCHER_NO Then
                rowAPTINVH1 = dst.Tables("APTINVH1").Rows.Find(VOUCHER_NO)
            Else
                rowAPTINVH1 = Fill_Record("APTINVH1", rowAPTINVHX("VOUCHER_NO"), , False)
            End If
            rowAPTINVH1("INV_STATUS") = "P"
            rowAPTINVH1("INV_PAYMENTS") = rowAPTINVHX("INV_BALANCE")
            rowAPTINVH1("INV_DISC_TAKEN") = rowAPTINVHX("INV_DISC_AMT")
            rowAPTINVH1("INV_LAST_PMT_DATE") = CHECK_DATE
            rowAPTINVH1("BATCH_NO_PYMT") = ""
            rowAPTINVH1("INV_BALANCE") = 0
            rowAPTINVH1("BATCH_PYMT") = 0
            rowAPTINVH1("BATCH_DISC") = 0
            rowAPTINVH1("BANK_CODE") = BANK_CODE
            rowAPTINVH1("CHECK_NUM") = CHECK_NUM
            rowAPTINVH1("CHECK_DATE") = CHECK_DATE

            Dim rowAPTCHCK2 As DataRow = dst.Tables("APTCHCK2").NewRow
            rowAPTCHCK2("BANK_CODE") = BANK_CODE
            rowAPTCHCK2("CHECK_NUM") = CHECK_NUM
            SEQ_NUM = SEQ_NUM + 1
            rowAPTCHCK2("SEQ_NUM") = SEQ_NUM
            rowAPTCHCK2("VEND_CODE") = rowAPTINVH1("VEND_CODE")
            rowAPTCHCK2("INV_NUM") = rowAPTINVH1("INV_NUM")
            rowAPTCHCK2("INV_DATE") = rowAPTINVH1("INV_DATE")
            rowAPTCHCK2("VOUCHER_NO") = rowAPTINVH1("VOUCHER_NO")
            rowAPTCHCK2("INV_AMT_APPLIED") = rowAPTINVH1("INV_AMT")
            rowAPTCHCK2("INV_DISC_TAKEN") = rowAPTINVH1("INV_DISC_AMT")
            rowAPTCHCK2("LC_FEE") = rowAPTINVH1("LC_FEE")
            dst.Tables("APTCHCK2").Rows.Add(rowAPTCHCK2)
        Next

        Dim rowAPTCHCK1 As DataRow = dst.Tables("APTCHCK1").NewRow
        rowAPTCHCK1("BANK_CODE") = BANK_CODE
        rowAPTCHCK1("CHECK_NUM") = CHECK_NUM
        rowAPTCHCK1("CHECK_DATE") = CHECK_DATE
        rowAPTCHCK1("CHECK_AMT") = CHECK_AMT
        rowAPTCHCK1("PYMT_METHOD") = PYMT_METHOD
        rowAPTCHCK1("VEND_CODE") = HFs("VEND_CODE")
        rowAPTCHCK1("VEND_CODE_AP") = VEND_CODE_AP
        rowAPTCHCK1("VEND_ALT_CODE") = VEND_ALT_CODE
        rowAPTCHCK1("OPS_YYYYPP") = ASCMAIN1.CYP
        rowAPTCHCK1("CHECK_STATUS") = "I"
        rowAPTCHCK1("VEND_NAME") = VEND_NAME
        rowAPTCHCK1("INIT_DATE") = DATETIME_STAMP
        rowAPTCHCK1("INIT_OPER") = ASCMAIN1.USER_ID
        rowAPTCHCK1("REGISTER_IND") = "0"
        dst.Tables("APTCHCK1").Rows.Add(rowAPTCHCK1)

        Dim INV_PAYMENTS As Double = Val(dst.Tables("APTCHCK2").Compute("SUM(INV_AMT_APPLIED)", "") & "")
        Dim INV_DISC_TAKEN As Double = Val(dst.Tables("APTCHCK2").Compute("SUM(INV_DISC_TAKEN)", "") & "")

        rowAPTVEND5.Item("VEND_PAYMENTS_MTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_MTD") & "") + INV_PAYMENTS
        rowAPTVEND5.Item("VEND_PAYMENTS_YTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_YTD") & "") + INV_PAYMENTS
        rowAPTVEND5.Item("VEND_DISC_TAKEN_MTD") = Val(rowAPTVEND5.Item("VEND_DISC_TAKEN_MTD") & "") + INV_DISC_TAKEN
        rowAPTVEND5.Item("VEND_DISC_TAKEN_YTD") = Val(rowAPTVEND5.Item("VEND_DISC_TAKEN_YTD") & "") + INV_DISC_TAKEN
        rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") & "") + 1
        rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") & "") + 1
        rowAPTVEND5.Item("VEND_LAST_PMT_DATE") = CHECK_DATE
        rowAPTVEND5.Item("VEND_LAST_PMT_AMT") = INV_PAYMENTS

        If auto_next_check Then
            Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", rowAPTINVH1("BANK_CODE"))
            If rowGLTBANK1("BANK_LAST_CHECK_NO") & "" = BANK_LAST_CHECK_NO Then
                rowGLTBANK1("BANK_LAST_CHECK_NO") = BANK_NEXT_CHECK_NO
            End If
        End If

        'If last_check_no <> "" And LAST_CHECK_NO_bank = datAPWINVH1.Recordset.Fields("BANK_CODE").Value Then
        '    OraD.Parameters("CODE").Value = datAPWINVH1.Recordset.Fields("BANK_CODE").Value
        '    dynGLTBANK1.Refresh()
        '    dynGLTBANK1.Edit()
        '    dynGLTBANK1.Fields("LAST_CHECK_NO").Value = last_check_no
        '    dynGLTBANK1.Update()
        'End If
    End Sub

    Private Sub grdAPTINVHX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHX.AfterCellUpdate
    End Sub

    Private Sub grdAPTINVHX_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVHX.AfterRowUpdate
        grdAPTINVHX.DisplayLayout.Bands(0).SummaryFooterCaption = "Total Amount Selected = " & Format(Val(dst.Tables("APTINVHX").Compute("Sum(INV_PAYMENTS)", "SELECTED = '1'") & ""), "$#,##0.00")
    End Sub

    Private Sub grdAPTINVHX_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHX.CellChange
        grdAPTINVHX.Update()
    End Sub

    'Private Sub chkBatchEntry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBatchEntry.CheckedChanged
    '    chkQuickEntry.Enabled = Not chkBatchEntry.Checked
    '    chkRecurring.Enabled = Not chkBatchEntry.Checked
    '    optFilter.Enabled = Not chkBatchEntry.Checked
    '    grdAPTINVR1.Visible = Not chkBatchEntry.Checked
    '    UltraGroupBox1.Visible = Not chkBatchEntry.Checked
    '    grdAPTINVHB.Visible = chkBatchEntry.Checked
    'End Sub

    Private Sub UltraTextEditor2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraTextEditor2.ValueChanged

    End Sub

    Sub Prepare_for_Multi_Invoice_Edit()

        Fill_Record("APTINVHM", Absx1.txtFor("VEND_CODE").Text)
        grdAPTINVHM.DisplayLayout.Bands(0).Columns("UPDATE_STATUS").Hidden = True
        grdAPTINVHM.DisplayLayout.Bands(0).Columns("UPDATE_MESSAGE").Hidden = True
        grpLastChange.Dock = DockStyle.None
        grpLastChange.Left = Absx1.txtFor("VEND_NAME").Left + Absx1.txtFor("VEND_NAME").Width + 2
        grpLastChange.Dock = DockStyle.Right
    End Sub

    Sub Prepare_for_Batch_Entry()

    End Sub

    Private Sub grdAPTINVHB_AfterCellUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHB.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "VEND_CODE"
                Dim VEND_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdAPTINVHB, "APTVEND1", "VEND_CODE", "VEND_NAME")
        End Select
    End Sub

    Private Sub grdAPTINVHB_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVHB.AfterExitEditMode
        With grdAPTINVHB
            Select Case .ActiveCell.Column.Key
                Case "VEND_CODE"
                    Dim VEND_CODE As String = .ActiveCell.Text
                    If VEND_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(VEND_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdAPTINVHB_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVHB.AfterRowActivate
        With grdAPTINVHB
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("VEND_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdAPTINVHB.ActiveRow.Cells("VEND_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else

            End If
        End With
    End Sub

    Private Sub grdAPTINVHB_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVHB.BeforeRowUpdate
        With grdAPTINVHB
            If e.Row.Cells("VEND_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("APTVEND1", e.Row.Cells("VEND_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Vendor Code (" & e.Row.Cells("VEND_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.Cells("VOUCHER_NO").Text = "" Then
                    .ActiveRow.Cells("VOUCHER_NO").Value = Format(Val(dst.Tables("APTINVHB").Compute("Max(VOUCHER_NO)", "") & "") + 1, "0000000000")
                End If
            End If
        End With

    End Sub

    Private Sub grdAPTINVHB_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHB.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdAPTINVHB, sql_where, sql_where <> "")
    End Sub

    Sub Update_Batch()
        ' HAVE A LINE NUMBER SERVE AS THE KEY

        batch_update = True
        Click_Command("Cancel")
        chkQuickEntry.Checked = False
        Dim V As Integer = 0
        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()

        Application.DoEvents()

        Dim STATS(2) As Integer
        For Each rowAPTINVHB As DataRow In dst.Tables("APTINVHB").Rows
            STATS(0) += 1
            V = V + 1
            rowAPTINVHB("VOUCHER_NO") = "X" & Format(V, "000000000")
            Absx1.txtFor("VEND_CODE").Text = rowAPTINVHB("VEND_CODE") & ""
            Absx1.txtFor("INV_NUM").Text = rowAPTINVHB("INV_NUM") & ""
            If rowAPTINVHB("INV_TYPE") & "" = "" Then
                Absx1.cbeFor("INV_TYPE").Text = "I"
            Else
                'Stop
                Absx1.cbeFor("INV_TYPE").Text = rowAPTINVHB("INV_TYPE")
            End If
            Application.DoEvents()
            X.EndCurrentEdit()

            Click_Command("New")
            Application.DoEvents()
            If EMsg <> "" Then
                rowAPTINVHB("UPDATE_STATUS") = "ERROR"
                rowAPTINVHB("UPDATE_MESSAGE") = EMsg
                STATS(2) += 1
            Else
                rowAPTINVHB("VEND_NAME") = Absx1.txtFor("VEND_NAME").Text
                'rowAPTINVH1("INV_DATE") = rowAPTINVHB("INV_DATE")
                Absx1.dteFor("INV_DATE").Value = rowAPTINVHB("INV_DATE")
                'Absx1.numFor("INV_AMT").Value = rowAPTINVHB("INV_AMT")
                Absx1.numFor("INV_AMT_VEND").Value = rowAPTINVHB("INV_AMT")
                Application.DoEvents()
                Automatic_Distribution()
                If rowAPTINVHB("BANK_CODE") & "" <> "" Then
                    Absx1.txtFor("BANK_CODE").Text = rowAPTINVHB("BANK_CODE") & ""
                End If
                If rowAPTINVHB("TERM_CODE") & "" <> "" Then
                    Absx1.txtFor("TERM_CODE").Text = rowAPTINVHB("TERM_CODE") & ""
                End If
                Absx1.txtFor("INV_REF").Text = rowAPTINVHB("INV_REF") & ""
                If rowAPTINVHB("CHECK_NUM") & "" <> "" Then
                    Absx1.txtFor("CHECK_NUM").Text = rowAPTINVHB("CHECK_NUM") & ""
                    Absx1.dteFor("CHECK_DATE").Value = rowAPTINVHB("CHECK_DATE")
                End If
                If rowAPTINVHB("INV_STATUS") & "" = "" Then
                    Absx1.optFor("INV_STATUS").Value = "O"
                Else
                    Absx1.optFor("INV_STATUS").Value = rowAPTINVHB("INV_STATUS") & ""
                End If
                If Absx1.optFor("INV_STATUS").Value = "P" Then
                    If Absx1.txtFor("CHECK_NUM").Text = "" Then
                        If Absx1.txtFor("BANK_CODE").Text <> "" Then
                            Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                            BANK_LAST_CHECK_NO = rowGLTBANK1("BANK_LAST_CHECK_NO") & ""
                            BANK_NEXT_CHECK_NO = CStr(Val(BANK_LAST_CHECK_NO) + 1)
                            BANK_NEXT_CHECK_NO = ASCMAIN1.Format_Field(BANK_NEXT_CHECK_NO, "CHECK_NUM")
                            Absx1.txtFor("CHECK_NUM").Text = BANK_NEXT_CHECK_NO
                            Absx1.dteFor("CHECK_DATE").Value = rowAPTINVHB("INV_DATE")
                            auto_next_check = True

                            rowAPTINVHB("CHECK_NUM") = BANK_NEXT_CHECK_NO
                            rowAPTINVHB("CHECK_DATE") = rowAPTINVHB("INV_DATE")
                        End If
                    End If
                    Absx1.numFor("CHECK_AMT").Value = Absx1.numFor("INV_AMT").Value
                End If
                If rowAPTINVHB("INV_PYMT_METHOD") & "" <> "" Then
                    Absx1.txtFor("INV_PYMT_METHOD").Text = rowAPTINVHB("INV_PYMT_METHOD") & ""
                End If
                If rowAPTINVHB("INV_PYMT_CYCLE") & "" <> "" Then
                    Absx1.txtFor("INV_PYMT_CYCLE").Text = rowAPTINVHB("INV_PYMT_CYCLE") & ""
                End If
                If rowAPTINVHB("VEND_ALT_CODE") & "" <> "" Then
                    Absx1.txtFor("VEND_ALT_CODE").Text = rowAPTINVHB("VEND_ALT_CODE") & ""
                End If
                If rowAPTINVHB("POST_CODE") & "" <> "" Then
                    Absx1.txtFor("POST_CODE").Text = rowAPTINVHB("POST_CODE") & ""
                End If

                Application.DoEvents()
                tabMain.SelectedTab = tabMain.Tabs("GL Distribution")
                Application.DoEvents()
                X.EndCurrentEdit()

                Click_Command("Update")
                Application.DoEvents()
                If EMsg <> "" Then
                    rowAPTINVHB("UPDATE_STATUS") = "ERROR"
                    rowAPTINVHB("UPDATE_MESSAGE") = EMsg
                    Click_Command("Cancel")
                    STATS(2) += 1
                Else
                    rowAPTINVHB("UPDATE_STATUS") = "UPDATED"
                    STATS(1) += 1
                    rowAPTINVHB("VOUCHER_NO") = HFs("VOUCHER_NO")
                End If
            End If
            Application.DoEvents()
        Next
        batch_update = False
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_STATUS").Hidden = False
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_MESSAGE").Hidden = False
        Export_to_Excel(grdAPTINVHB)
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_STATUS").Hidden = True
        grdAPTINVHB.DisplayLayout.Bands("APTINVHB").Columns("UPDATE_MESSAGE").Hidden = True

        MsgBox("Batch Update Complete." & vbCr & vbCr & STATS(0) & " Records Processed" & vbCr & STATS(1) & " Records Updated Successfully" & vbCr & STATS(2) & " Records Were NOT Updated" & vbCr & vbCr & "Voucher (and Check Numbers, if any) appear in the Workbook Generated" & vbCr & "Processing Errors (if any) also appear in the 'Update Message' Column of the Workbook", MsgBoxStyle.OkOnly, "Verification")
    End Sub

    Sub Import_Batch_from_Excel()
        Dim FILENAME As String = ""

        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.InitialDirectory = "c:\"
            openFileDialog1.Title = "Locate the workbook containing AP Items to Import"
            openFileDialog1.Filter = "txt files (*.xls)|*.xls|All files (*.*)|*.*"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then

            Dim GCs As New Dictionary(Of String, String)
            With grdAPTINVHB.DisplayLayout.Bands("APTINVHB")
                For j As Integer = 1 To .Columns.Count
                    GCs.Add(.Columns(j - 1).Header.Caption, .Columns(j - 1).Key)
                Next
            End With

            grdAPTINVHB.DataSource = DirectCast(grdAPTINVHB.DataSource, DataTable).Clone

            Dim xlApp As Object
            Dim xlBook As Object
            Dim xlSheet As Object

            Try
                ASCMAIN1.Progress("Now Examining XLS Workbook")
                Me.Cursor = Cursors.WaitCursor

                Dim XLS As New Infragistics.Documents.Excel.Workbook

                ' Create the Excel App Object 
                xlApp = CreateObject("Excel.Application")
                ' Create the Excel Workbook Object. 
                xlBook = xlApp.Workbooks.Open(FILENAME)
                ' XLS = DirectCast(xlBook, Infragistics.Documents.Excel.Workbook)
                xlSheet = xlBook.Sheets(1)
                Dim heading_row As Integer = 0
                Dim columns_to_import As Integer = 0
                Dim COLUMN_NAMEs() As String
                ReDim COLUMN_NAMEs(0)
                Dim found_heading_row As Boolean
                For i As Integer = 1 To 10
                    found_heading_row = False
                    If xlSheet.cells(i, 1).text <> "" Then
                        found_heading_row = True
                        ReDim COLUMN_NAMEs(0)
                        For j As Integer = 1 To GCs.Count
                            Dim CellText As String = xlSheet.cells(i, j).text
                            If j > 1 And CellText = "" Then
                                columns_to_import = j - 1
                                Exit For
                            End If
                            If Not GCs.ContainsKey(CellText) Then
                                found_heading_row = False
                                Exit For
                            Else
                                ReDim Preserve COLUMN_NAMEs(j)
                                COLUMN_NAMEs(j) = GCs(CellText)
                                columns_to_import = j
                            End If
                        Next
                        If found_heading_row Then
                            heading_row = i
                            Exit For
                        End If
                    End If
                Next
                If heading_row = 0 Then
                    MsgBox("Cannot Find Heading Row", MsgBoxStyle.OkOnly, "Problem with Workbook Selected")
                Else
                    ASCMAIN1.Progress("Now Importing Data")
                    dst.Tables("APTINVHB").Rows.Clear()

                    Dim XR As Integer = heading_row + 1
                    Dim XI As Integer = 0
                    Do While xlSheet.cells(XR, 1).text <> "" And xlSheet.cells(XR, 1).text <> "Totals"
                        ASCMAIN1.Progress("-", CStr(XR - heading_row))
                        Dim rowAPTINVHB As DataRow = dst.Tables("APTINVHB").NewRow
                        For XC As Integer = 1 To columns_to_import
                            Dim CellText As String = xlSheet.cells(XR, XC).value & ""
                            COLUMN_NAME = COLUMN_NAMEs(XC)
                            If CellText <> "" Then
                                rowAPTINVHB(COLUMN_NAME) = CellText
                            End If
                        Next
                        If Len(rowAPTINVHB("INV_TYPE") & "") > 1 Then
                            If rowAPTINVHB("INV_TYPE") = "ChargeBack" Then
                                rowAPTINVHB("INV_TYPE") = "B"
                            Else
                                rowAPTINVHB("INV_TYPE") = Mid(rowAPTINVHB("INV_TYPE") & "", 1, 1)
                            End If
                        End If
                        If rowAPTINVHB("INV_STATUS") & "" <> "" Then

                        End If
                        rowAPTINVHB("VOUCHER_NO") = "" ' make sure that we use our own
                        If rowAPTINVHB("VEND_CODE") & "" <> "" Then
                            rowAPTINVHB("VEND_CODE") = ASCMAIN1.Format_Field(rowAPTINVHB("VEND_CODE"), "VEND_CODE")
                        End If
                        If rowAPTINVHB("INV_STATUS") & "" <> "" Then
                            rowAPTINVHB("INV_STATUS") = rowAPTINVHB("INV_STATUS").ToString.ToUpper
                            If rowAPTINVHB("INV_STATUS") <> "P" And rowAPTINVHB("INV_STATUS") <> "O" And rowAPTINVHB("INV_STATUS") <> "H" Then
                                rowAPTINVHB("INV_STATUS") = "O"
                            End If
                        End If

                        If rowAPTINVHB("VOUCHER_NO") & "" = "" Then
                            rowAPTINVHB("VOUCHER_NO") = Format(XR - heading_row, "0000000000")
                        End If
                        If rowAPTINVHB("UPDATE_STATUS") & "" = "UPDATED" Then
                            XI = XI + 1
                        Else
                            dst.Tables("APTINVHB").Rows.Add(rowAPTINVHB)
                        End If
                        XR = XR + 1
                    Loop
                    MsgBox("Import Successful" & vbCr & vbCr & "Records Processed = " & CStr(XR - (heading_row + 1)) & vbCr & "Records Imported = " & CStr(XR - (heading_row + 1) - XI) & vbCr & "Records Ignored (Updated Already) = " & CStr(XI), MsgBoxStyle.OkOnly, "Verification")
                End If

                xlApp.DisplayAlerts = False
                xlApp.Quit()

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Exception Occurred")
            Finally

                xlSheet = Nothing
                xlBook = Nothing
                xlApp = Nothing

                ASCMAIN1.Progress("")
                Me.Cursor = Cursors.Default
            End Try

            grdAPTINVHB.DataSource = dst.Tables("APTINVHB")
            grdAPTINVHB.Refresh()

        End If
    End Sub

    Sub Automatic_Distribution()
        If EntryMode = "N" Then
            If Absx1.cbeFor("INV_TYPE").Value = "A" Then
                'If Absx1.CtlFor("INV_TYPE").Text = "A" Then
                dst.Tables("APTINVH2").Rows.Clear()
            End If
            If dst.Tables("APTINVH2").Rows.Count = 0 And Val(numINV_AMT.Value & "") <> 0 Then
                Generate_Pre_Distribution()
                Calc_Totals()
            End If
        End If
        'Calculate_INV_DISC_AMT()
    End Sub

    Private Sub cmdNextCheckNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNextCheckNo.Click


        Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
        If rowGLTBANK1 Is Nothing Then
            MsgBox("Invalid Bank Code", MsgBoxStyle.OkOnly, "Could Not Generate the Next Check")
            auto_next_check = False
            BANK_LAST_CHECK_NO = ""
            BANK_NEXT_CHECK_NO = ""
        Else
            BANK_LAST_CHECK_NO = rowGLTBANK1("BANK_LAST_CHECK_NO") & ""
            BANK_NEXT_CHECK_NO = CStr(Val(BANK_LAST_CHECK_NO) + 1)
            BANK_NEXT_CHECK_NO = ASCMAIN1.Format_Field(BANK_NEXT_CHECK_NO, "CHECK_NUM")
            Absx1.txtFor("CHECK_NUM").Text = BANK_NEXT_CHECK_NO
            auto_next_check = True
        End If

    End Sub

    Private Sub SplitContainer2_Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub grdAPTINVHM_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHM.AfterCellUpdate
        lblCOLUMN_NAME.Visible = True
        lblCOLUMN_NAME.Text = e.Cell.Column.Header.Caption
        lblNEW_VALUE.Visible = True
        lblNEW_VALUE.Text = e.Cell.Text
        lblCOLUMN_NAME.Tag = e.Cell.Column.Key
    End Sub

    Private Sub grdAPTINVHM_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVHM.ClickCellButton
        Dim sql_where As String = ""
        If e.Cell.Column.Key = "VEND_ALT_CODE" Then
            sql_where = "VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
        End If
        grdClickCellButton(grdAPTINVHM, sql_where, sql_where <> "")
    End Sub

    Private Sub grdAPTINVHM_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTINVHM.InitializeLayout
        For Each C As UltraWinGrid.UltraGridColumn In grdAPTINVHM.DisplayLayout.Bands(0).Columns
            If C.CellActivation = UltraWinGrid.Activation.NoEdit Then
                C.CellAppearance.BackColor = Drawing.Color.LightGray
            End If
        Next
    End Sub

    Sub Update_Multi()

        batch_update = True
        Click_Command("Cancel")
        chkQuickEntry.Checked = False
        Dim V As Integer = 0
        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("APTINVH1"))
        X.EndCurrentEdit()

        Application.DoEvents()

        Dim STATS(2) As Integer
        For Each rowAPTINVHM As DataRow In dst.Tables("APTINVHM").Select("", "", DataViewRowState.ModifiedCurrent)
            STATS(0) += 1
            V = V + 1
            Absx1.txtFor("VOUCHER_NO").Text = rowAPTINVHM("VOUCHER_NO") & ""
            Application.DoEvents()
            X.EndCurrentEdit()

            Click_Command("Edit")
            Application.DoEvents()
            If EMsg <> "" Then
                rowAPTINVHM("UPDATE_STATUS") = "ERROR"
                rowAPTINVHM("UPDATE_MESSAGE") = EMsg
                STATS(2) += 1
            Else
                For c As Integer = 0 To rowAPTINVHM.ItemArray.Length - 1
                    Dim COLUMN_NAME As String = rowAPTINVHM.Table.Columns(c).ColumnName
                    If grdAPTINVHM.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit Then
                        Absx1.txtFor(COLUMN_NAME).Text = rowAPTINVHM.Item(c) & ""
                    End If

                    'If rowAPTINVHM.Item(c, DataRowVersion.Current) & "" _
                    '<> rowAPTINVHM.Item(c, DataRowVersion.Original) & "" Then
                    '    Stop

                    '    If grdAPTINVHM.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit Then
                    '        Stop
                    '    End If

                    '    Absx1.txtFor(COLUMN_NAME).Text = rowAPTINVHM.Item(c, DataRowVersion.Current) & ""
                    'End If
                Next

                Application.DoEvents()
                tabMain.SelectedTab = tabMain.Tabs("GL Distribution")
                Application.DoEvents()
                X.EndCurrentEdit()

                Click_Command("Update")
                Application.DoEvents()

                If EMsg <> "" Then
                    rowAPTINVHM("UPDATE_STATUS") = "ERROR"
                    rowAPTINVHM("UPDATE_MESSAGE") = EMsg
                    Click_Command("Cancel")
                    STATS(2) += 1
                Else
                    rowAPTINVHM("UPDATE_STATUS") = "UPDATED"
                    STATS(1) += 1
                End If
            End If
            Application.DoEvents()
        Next
        batch_update = False

        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_STATUS").Hidden = False
        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_MESSAGE").Hidden = False
        Export_to_Excel(grdAPTINVHM)
        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_STATUS").Hidden = True
        grdAPTINVHM.DisplayLayout.Bands("APTINVHM").Columns("UPDATE_MESSAGE").Hidden = True

        MsgBox("Multiple-Invoice Edit Complete." & vbCr & vbCr & STATS(0) & " Records Processed" & vbCr & STATS(1) & " Records Updated Successfully" & vbCr & STATS(2) & " Records Were NOT Updated" & vbCr & vbCr & "Processing Errors (if any) will appear in the 'Update Message' Column of the Workbook", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Private Sub splMulti_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs)

    End Sub

    Private Sub UltraGroupBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Sub Copy_Change_to(ByVal Copy_to As String)
        If Copy_to = "All" Then
            For Each row As DataRow In dst.Tables("APTINVHM").Rows
                row.Item(lblCOLUMN_NAME.Tag) = lblNEW_VALUE.Text
            Next
        Else
            For Each C As UltraWinGrid.UltraGridRow In grdAPTINVHM.Selected.Rows
                C.Cells(lblCOLUMN_NAME.Tag).Value = lblNEW_VALUE.Text
                grdAPTINVHM.UpdateData()
            Next
        End If
    End Sub

    Private Sub txtOPS_YYYYPP_ACCRUE_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles txtOPS_YYYYPP_ACCRUE.EditorButtonClick
        txtOPS_YYYYPP_ACCRUE.ReadOnly = InquiryMode
    End Sub

    Private Sub txtOPS_YYYYPP_ACCRUE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOPS_YYYYPP_ACCRUE.ValueChanged

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
     ByVal ctl As Control,
     ByVal COLUMN_NAME As String,
     Optional ByRef sql_where As String = "",
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "OPS_YYYYPP_ACCRUE"
                If ROWs("APTPARM1").Item("AP_PARM_ALLOW_ACCRUAL") & "" = "1" Then
                    sql_where = "OPS_YYYYPP >= (SELECT GL_PARM_CURRENT_YYYYPP FROM GLTPARM1) and OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"
                Else
                    sql_where = "OPS_YYYYPP > (SELECT GL_PARM_CURRENT_YYYYPP FROM GLTPARM1) and OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"
                End If
                If ASCMAIN1.CLIENT = "VAN" Then
                    sql_where = "(" & sql_where & ") or OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1) & "'"
                End If
                txtOPS_YYYYPP_ACCRUE.ReadOnly = True Or InquiryMode
            Case "LC_CTL_NO"
                sql_where = "VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "' and STATUS_CODE = 'O'"
        End Select
    End Sub



#Region "grdAPTINVH5"

    Private Sub grdAPTINVH5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH5.AfterCellUpdate
        If e.Cell.Column.Key = "INV_COST" Or e.Cell.Column.Key = "INV_QTY" Then
            Dim AMT_VAR As Decimal = 0

            Dim AP_COST As Decimal = Val(e.Cell.Row.Cells("AP_COST").Value & "")
            Dim INV_COST As Decimal = Val(e.Cell.Row.Cells("INV_COST").Value & "")
            Dim QTY_REC As Int32 = Val(e.Cell.Row.Cells("QTY_REC").Value & "")
            Dim INV_QTY As Int32 = Val(e.Cell.Row.Cells("INV_QTY").Value & "")

            AMT_VAR = (QTY_REC * AP_COST) - (INV_QTY * INV_COST)
            If AMT_VAR < 0 Then
                e.Cell.Row.Cells("CB").Value = "1"
            ElseIf AMT_VAR >= 0 Then
                e.Cell.Row.Cells("CB").Value = "0"
            End If
        End If

    End Sub

    Private Sub grdAPTINVH5_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH5.AfterRowsDeleted

    End Sub

    Private Sub grdAPTINVH5_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH5.AfterRowUpdate
        Calc_DIST_PO()
    End Sub

    Private Sub grdAPTINVH5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVH5.BeforeRowUpdate
        Dim AMT_VAR As Decimal = 0

        Dim AP_COST As Decimal = Val(e.Row.Cells("AP_COST").Value & "")
        Dim INV_COST As Decimal = Val(e.Row.Cells("INV_COST").Value & "")
        Dim QTY_REC As Int32 = Val(e.Row.Cells("QTY_REC").Value & "")
        Dim INV_QTY As Int32 = Val(e.Row.Cells("INV_QTY").Value & "")

        AMT_VAR = (QTY_REC * AP_COST) - (INV_QTY * INV_COST)
        If AMT_VAR < 0 Then
            'e.Row.Cells("CB").Value = "1"
        ElseIf AMT_VAR >= 0 Then
            e.Row.Cells("CB").Value = "0"
        End If
    End Sub

    Private Sub grdAPTINVH5_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTINVH5.InitializeRow
        If Val(e.Row.Cells("QTY_REC").Value & "") - Val(e.Row.Cells("INV_QTY").Value & "") <> 0 Then
            e.Row.Cells("INV_QTY").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("INV_QTY").Appearance.FontData.Bold = DefaultableBoolean.True
        Else
            e.Row.Cells("INV_QTY").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("INV_QTY").Appearance.FontData.Bold = DefaultableBoolean.False
        End If

        If Val(e.Row.Cells("AP_COST").Value & "") <> Val(e.Row.Cells("INV_COST").Value & "") Then
            e.Row.Cells("INV_COST").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("INV_COST").Appearance.FontData.Bold = DefaultableBoolean.True
        Else
            e.Row.Cells("INV_COST").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("INV_COST").Appearance.FontData.Bold = DefaultableBoolean.False
        End If

        Dim AMT_VAR As Decimal = Val(e.Row.Cells("AMT_VAR").Value & "")

        If AMT_VAR < 0 Then
            e.Row.Cells("CB").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("AMT_VAR").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("QTY_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Cells("AMT_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
        ElseIf AMT_VAR > 0 Then
            e.Row.Cells("CB").Appearance.BackColor = Drawing.Color.LightGreen
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.DarkGreen
            e.Row.Cells("AMT_VAR").Appearance.ForeColor = Drawing.Color.DarkGreen
            e.Row.Cells("QTY_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
            e.Row.Cells("AMT_VAR").Appearance.FontData.Bold = DefaultableBoolean.True
        Else
            e.Row.Cells("CB").Appearance.BackColor = Drawing.Color.Empty
            e.Row.Cells("QTY_VAR").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("AMT_VAR").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("QTY_VAR").Appearance.FontData.Bold = DefaultableBoolean.False
            e.Row.Cells("AMT_VAR").Appearance.FontData.Bold = DefaultableBoolean.False
        End If
    End Sub

    Sub Create_Chargeback()

        ' this routine is no longer used - it was intended to create an independent chargeback record

        Dim VOUCHER_NO_CB As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
        Dim INV_ADJUSTMENTS As Decimal = Val(Absx1.numFor("INV_ADJUSTMENTS").Value & "")

        Dim rowAPTINVH1_INV As DataRow = dst.Tables("APTINVH1").Rows.Find(HFs("VOUCHER_NO"))

        Dim rowAPTINVH1 As DataRow = dst.Tables("APTINVH1").NewRow
        rowAPTINVH1.Item("VOUCHER_NO") = VOUCHER_NO_CB
        For Each COLUMN_NAME As String In New String() _
        {"VEND_CODE", "INV_NUM", "INV_DATE", "VEND_ALT_CODE", "INV_PYMT_CYCLE",
         "TERM_CODE", "INV_DUE_DATE", "INV_DISC_DUE", "INV_DISC_AMT", "POST_CODE",
         "REASON_CODE", "INIT_OPER", "LAST_OPER", "INIT_DATE", "LAST_DATE",
         "CURR_CODE", "CURR_EXCH_RATE", "INV_SEP_CHECK", "VEND_CODE_AP", "BANK_CODE",
         "REGISTER_IND", "REGISTER_XNO", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE",
         "INV_PYMT_METHOD", "INV_REMIT_TO", "INV_PAID_UPON_ENTRY"}
            rowAPTINVH1.Item(COLUMN_NAME) = rowAPTINVH1_INV.Item(COLUMN_NAME)
        Next
        rowAPTINVH1.Item("INV_TYPE") = "B"
        rowAPTINVH1.Item("INV_AMT") = -1 * INV_ADJUSTMENTS
        rowAPTINVH1.Item("INV_BALANCE") = -1 * INV_ADJUSTMENTS
        rowAPTINVH1.Item("INV_STATUS") = "O"
        rowAPTINVH1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowAPTINVH1.Item("MEMO_PRINT_IND") = "0"
        dst.Tables("APTINVH1").Rows.Add(rowAPTINVH1)

        Dim VOUCHER_LNO_ctr As Int32 = 0
        For Each rowAPTINVH5_VAR As DataRow In dst.Tables("APTINVH5_VAR").Select("AMT_VAR_CB <> 0")
            Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
            rowAPTINVH2.Item("VOUCHER_NO") = VOUCHER_NO_CB
            VOUCHER_LNO_ctr += 1
            rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
            rowAPTINVH2.Item("ACCT_CODE") = rowAPTINVH5_VAR.Item("ACCT_CODE_PPV")
            rowAPTINVH2.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowAPTINVH2.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowAPTINVH2.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowAPTINVH2.Item("INV_LINE_AMT") = Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & "")
            dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        Next

    End Sub

    Sub Create_APTINVH5_VAR()

        dst.Tables("APTINVH5_VAR").Rows.Clear()

        For Each rowAPTINVH5 As DataRow In dst.Tables("APTINVH5") _
        .Select("", "", DataViewRowState.CurrentRows)
            Dim STYLE_CLASS_CODE As String = rowAPTINVH5.Item("STYLE_CLASS_CODE") & ""
            Dim rowICTCLAS1 As DataRow = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
            Dim SEG2_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            ' Dim SEG2_CODE As String = rowICTCOST1.Item("SEG2_CODE") & ""
            'If SEG2_CODE = "" Then
            '    SEG2_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            'End If
            Dim rowAPTINVH5_VAR As DataRow = dst.Tables("APTINVH5_VAR").Rows.Find(STYLE_CLASS_CODE)
            If rowAPTINVH5_VAR Is Nothing Then
                rowAPTINVH5_VAR = dst.Tables("APTINVH5_VAR").NewRow
                rowAPTINVH5_VAR.Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                ' rowAPTINVH5_VAR.Item("ACCT_CODE_PPV") = rowAPTINVH5.Item("ACCT_CODE_PPV")
                Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(rowAPTINVH5.Item("RECEIPT_NO"))
                rowAPTINVH5_VAR.Item("SEG2_CODE") = SEG2_CODE ' rowICTIREC1.Item("SEG2_CODE")
                dst.Tables("APTINVH5_VAR").Rows.Add(rowAPTINVH5_VAR)
            End If
            rowAPTINVH5_VAR.Item("AMT_REC") = Val(rowAPTINVH5_VAR.Item("AMT_REC") & "") + Val(rowAPTINVH5.Item("AMT_REC") & "")
            rowAPTINVH5_VAR.Item("AMT_INV") = Val(rowAPTINVH5_VAR.Item("AMT_INV") & "") + Val(rowAPTINVH5.Item("AMT_INV") & "")
            rowAPTINVH5_VAR.Item("AMT_VAR") = Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") + Val(rowAPTINVH5.Item("AMT_VAR") & "")
            If rowAPTINVH5.Item("CB") & "" = "1" Then
                rowAPTINVH5_VAR.Item("AMT_VAR_CB") = Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & "") + Val(rowAPTINVH5.Item("AMT_VAR") & "")
            End If
        Next

    End Sub

    Sub Create_APTINVH2_P()

        Delete_Rows("APTINVH2", "INV_LTYP = 'P'")

        Dim VOUCHER_LNO_ctr As Int32 = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'") & "")

        Dim INV_LINE_AMT As Decimal = Val(dst.Tables("APTINVH5_VAR").Compute("SUM(AMT_REC)", "") & "")
        Dim rowAPTINVH2 As DataRow

        If INV_LINE_AMT <> 0 Then
            rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
            rowAPTINVH2.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            VOUCHER_LNO_ctr += 1
            rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
            rowAPTINVH2.Item("ACCT_CODE") = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_INVTY_PUR")
            rowAPTINVH2.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowAPTINVH2.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowAPTINVH2.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowAPTINVH2.Item("INV_LINE_AMT") = INV_LINE_AMT
            rowAPTINVH2.Item("INV_LTYP") = "P"
            dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        End If

        For Each rowAPTINVH5_VAR As DataRow In dst.Tables("APTINVH5_VAR").Select("ISNULL(AMT_VAR,0) - ISNULL(AMT_VAR_CB,0) <> 0")
            ' INV_LINE_AMT = -1 * (Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") - Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & ""))
            INV_LINE_AMT = (Val(rowAPTINVH5_VAR.Item("AMT_VAR") & "") - Val(rowAPTINVH5_VAR.Item("AMT_VAR_CB") & ""))
            If INV_LINE_AMT <> 0 Then
                rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                rowAPTINVH2.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
                VOUCHER_LNO_ctr += 1
                rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                ' CHANGE BELOW TO MOVE FROM H/C 5000 TO INVTY_PUR WAS MADE 07/08 WJZ IN RESPONSE TO BP SAYING HE WANTED THE AUTOMATED ENTRY TO HIT 5150 INSTEAD OF 5000
                rowAPTINVH2.Item("ACCT_CODE") = ROWs("ICTPARM1").Item("IC_PARM_ACCT_CODE_INVTY_PUR") ' "5000" ' rowAPTINVH5_VAR.Item("ACCT_CODE_PPV")
                rowAPTINVH2.Item("SEG2_CODE") = rowAPTINVH5_VAR.Item("SEG2_CODE")
                rowAPTINVH2.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                rowAPTINVH2.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                rowAPTINVH2.Item("INV_LINE_AMT") = INV_LINE_AMT
                rowAPTINVH2.Item("INV_LTYP") = "P"
                dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
            End If
        Next
    End Sub

    Sub Create_APTINVH2_R()

        Delete_Rows("APTINVH2", "INV_LTYP = 'R'")

        Dim VOUCHER_LNO_ctr As Int32 = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'") & "")

        Dim INV_LINE_AMT As Decimal = -1 * Val(dst.Tables("APTINVH7").Compute("SUM(REBATE_USED)", "") & "")
        Dim rowAPTINVH2 As DataRow

        If INV_LINE_AMT <> 0 Then
            rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
            rowAPTINVH2.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            VOUCHER_LNO_ctr += 1
            rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
            rowAPTINVH2.Item("ACCT_CODE") = ROWs("PPTPARM1").Item("PP_PARM_REBATE_ACCRUAL")
            rowAPTINVH2.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            rowAPTINVH2.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            rowAPTINVH2.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            rowAPTINVH2.Item("INV_LINE_AMT") = INV_LINE_AMT
            rowAPTINVH2.Item("INV_LTYP") = "R"
            dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
        End If

        ' NO VARIANCE LOGIC YET - WE STILL HAVE TO DEAL WITH HOW TO HANDLE VARIANCES BETWEEN WHAT WAS ACCRUED AND WHAT WAS OFFERED BY THE VENDOR
    End Sub

    Sub Create_APTINVH2_from_APTINVH7()

        Delete_Rows("APTINVH2", "INV_LTYP = 'O'")

        Dim VOUCHER_LNO_ctr As Int32 = Val(dst.Tables("APTINVH2").Compute("MAX(VOUCHER_LNO)", "VOUCHER_NO = '" & HFs("VOUCHER_NO") & "'") & "")

        For Each rowAPTINVH7 As DataRow In dst.Tables("APTINVH7").Select("", "VOUCHER_CLNO")
            Dim CTL_NO As String = rowAPTINVH7.Item("CTL_NO")
            Dim rowPOTLCST1 As DataRow = dst.Tables("POTLCST1").Rows.Find(CTL_NO)
            Dim COST_ACC As Decimal = Val(rowPOTLCST1.Item("COST_ACC") & "")
            Dim COST_ACT As Decimal = Val(rowPOTLCST1.Item("COST_ACT") & "")
            Dim COST_CATGY_CODE As String = rowAPTINVH7.Item("COST_CATGY_CODE")
            Dim rowPOTCATG1 As DataRow = dst.Tables("POTCATG1").Rows.Find(COST_CATGY_CODE)
            Dim rowAPTINVH2 As DataRow = Nothing

            If COST_ACC <> 0 Then
                rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                rowAPTINVH2.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
                VOUCHER_LNO_ctr += 1
                rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                rowAPTINVH2.Item("ACCT_CODE") = rowPOTCATG1.Item("ACCT_CODE_ACC")
                rowAPTINVH2.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                rowAPTINVH2.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                rowAPTINVH2.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                rowAPTINVH2.Item("INV_LINE_AMT") = COST_ACC
                rowAPTINVH2.Item("INV_LTYP") = "O"
                dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
            End If

            If COST_ACT - COST_ACC <> 0 Then
                rowAPTINVH2 = dst.Tables("APTINVH2").NewRow
                rowAPTINVH2.Item("VOUCHER_NO") = HFs("VOUCHER_NO")
                VOUCHER_LNO_ctr += 1
                rowAPTINVH2.Item("VOUCHER_LNO") = VOUCHER_LNO_ctr
                rowAPTINVH2.Item("ACCT_CODE") = rowPOTCATG1.Item("ACCT_CODE_EXP")
                rowAPTINVH2.Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                rowAPTINVH2.Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                rowAPTINVH2.Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                rowAPTINVH2.Item("INV_LINE_AMT") = COST_ACT - COST_ACC
                rowAPTINVH2.Item("INV_LTYP") = "O"
                dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
            End If
        Next
    End Sub
#End Region


    Private Sub grdICTIREC1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIREC1.ClickCellButton
        If InquiryMode Then Exit Sub

        If e.Cell.Row.IsDataRow Then
            ProcessDoubleClickedRow()
        End If
    End Sub

    Private Sub ProcessDoubleClickedRow()

        If InquiryMode Then Exit Sub

        Dim RECEIPT_NO As String = grdICTIREC1.ActiveRow.Cells("RECEIPT_NO").Text

        If dst.Tables("APTINVH5_SUM").Rows.Find(RECEIPT_NO) Is Nothing Then
            Me.Cursor = Cursors.WaitCursor
            Application.DoEvents()

            Dim APTINVH5_SUM As DataRow = Add_APTINVH5_SUM(RECEIPT_NO, _
                                                           CDate(grdICTIREC1.ActiveRow.Cells("RECEIPT_DATE").Text), _
                                                          Val(grdICTIREC1.ActiveRow.Cells("QTY_REC").Text), _
                                                          Val(grdICTIREC1.ActiveRow.Cells("AMT_REC").Text))
            Dim VOUCHER_DLNO_max As Integer = Val(dst.Tables("APTINVH5").Compute("MAX(VOUCHER_DLNO)", "") & "")

            Dim sql As String = ""
            If RECEIPT_NO.StartsWith("S") Then
                sql = "Select '" & HFs("VOUCHER_NO") & "' VOUCHER_NO " & vbCrLf _
                    & ", POTSHIP3.RECEIPT_LNO + " & CStr(VOUCHER_DLNO_max) & " VOUCHER_DLNO " & vbCrLf _
                    & ", POTSHIP2.TRAN_NO RECEIPT_NO, POTSHIP3.RECEIPT_LNO" & vbCrLf _
                    & ", POTSHIP3.PO_QTY_SHP INV_QTY, POTSHIP3.PO_COST INV_COST" & vbCrLf _
                    & ", ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                    & ", POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                    & ", POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & ", ICTSTYL1.STYLE_DESC" & vbCrLf _
                    & ", POTSHIP3.PO_QTY_SHP QTY_REC, POTSHIP3.PO_COST, POTSHIP3.PO_COST AP_COST" & vbCrLf _
                    & ", POTORDR2.STYLE_CODE, ICTSTYL1.STYLE_UOM" & vbCrLf _
                    & " from " & POTSHIP3 & " POTSHIP3," & POTSHIP2 & " POTSHIP2,POTORDR2,ICTSTYL1" & vbCrLf _
                    & " where POTSHIP2.TRAN_NO = '" & RECEIPT_NO & "'" _
                    & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                    & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                    & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE"
            Else
                sql = "Select '" & HFs("VOUCHER_NO") & "' VOUCHER_NO " & vbCrLf _
                    & ", ICTIREC2.RECEIPT_LNO + " & CStr(VOUCHER_DLNO_max) & " VOUCHER_DLNO " & vbCrLf _
                    & ", ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO" & vbCrLf _
                    & ", ICTIREC2.QTY_REC INV_QTY, ICTIREC2.AP_COST INV_COST" & vbCrLf _
                    & ", ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                    & ", ICTIREC2.PO_SHIPMENT_NO, ICTIREC2.PO_SHIPMENT_LNO" & vbCrLf _
                    & ", ICTIREC2.PO_ORDER_NO, ICTIREC2.PO_ORDER_LNO" & vbCrLf _
                    & ", ICTSTYL1.STYLE_DESC" & vbCrLf _
                    & ", ICTIREC2.QTY_REC, ICTIREC2.PO_COST, ICTIREC2.AP_COST" & vbCrLf _
                    & ", ICTIREC2.STYLE_CODE, ICTIREC2.STYLE_UOM" & vbCrLf _
                    & " from ICTIREC2,ICTSTYL1" & vbCrLf _
                    & " where ICTIREC2.RECEIPT_NO = '" & RECEIPT_NO & "'" _
                    & "   and ICTSTYL1.STYLE_CODE = ICTIREC2.STYLE_CODE"
            End If
            Fill_Records("APTINVH5", , False, sql)

            grdICTIREC1.ActiveRow.Cells("ACCRUAL_STATUS").Value = "2"
            grdICTIREC1.UpdateData()

            'dst.Tables("APTINVH5_SUM").Rows.Find("RECEIPT_NO = '" & RECEIPT_NO & "'")
            grdAPTINVH5_SUM.Rows(grdAPTINVH5_SUM.Rows.Count - 1).Activate()

            'dteINV_BL_DATE.Value = grdICTIREC1.ActiveRow.Cells("DATE_RECEIVED").Value
            Setup_INV_BL_DATE()
            Calculate_INV_DUE_DATE()
            Calc_DIST_PO()
            'Display_Totals()
            Me.Cursor = Cursors.Default
            grdAPTINVH5.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        Else
            MsgBox("Receipt " & RECEIPT_NO & " is already part of this Voucher", MsgBoxStyle.OkOnly, "Duplicate Selection")
        End If
    End Sub




    Private Sub numINV_AMT_VEND_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles numINV_AMT_VEND.Leave
        Automatic_Distribution()
    End Sub

    Private Sub numINV_AMT_VEND_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numINV_AMT_VEND.ValueChanged
        Calc_Totals()
    End Sub

    Private Sub grdAPTINVH8_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTINVH8.AfterRowsDeleted
        Calc_DIST_Adjustments()
    End Sub

    Private Sub grdAPTINVH8_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH8.AfterRowUpdate
        Calc_DIST_Adjustments()
    End Sub

    Private Sub grdAPTINVH8_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVH8.BeforeRowUpdate
        With grdAPTINVH8
            If Not e.Cancel Then
                If e.Row.Cells("VOUCHER_NO").Text = "" Then
                    .ActiveRow.Cells("VOUCHER_NO").Value = Absx1.CtlFor("VOUCHER_NO").Text
                    .ActiveRow.Cells("VOUCHER_ANO").Value = Val(dst.Tables("APTINVH8").Compute("Max(VOUCHER_ANO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub optAssociate_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optAssociate.ValueChanged
        txtPO_ORDER_NO.Visible = (optAssociate.Value = "3")
        txtPO_SHIPMENT_NO.Visible = (optAssociate.Value = "1")
        If optAssociate.Value = "2" Then
            'ASCMAIN1.sql = "" _
            '     & sqlPOTORDR1 & " where PO_STATUS = 'O'" _
            '     & " union " _
            '     & sqlPOTORDR1 & " where PO_ORDER_NO in (Select Distinct PO_ORDER_NO from POTSHIP3 where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "')"
            'Dim tbl As DataTable = ASCDATA1.GetDataTable

            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PO_SHIPMENT_LNO")

            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = False
                ' ASCMAIN1.CodeSelector.UseDataFromTable = tbl
                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    txtPO_SHIPMENT_LNO.Text = ASCMAIN1.CodeSelector.SelectedCode
                    txtPO_SHIPMENT_NO.Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("PO_SHIPMENT_NO")
                    txtCONTAINER_NO.Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("CONTAINER_NO")
                    txtBOL_NO.Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("BOL_NO")
                    txtCOMM_INV_NO.Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("COMM_INV_NO")
                Else
                    optAssociate.Value = "1"
                End If
            End If
        Else
            txtPO_ORDER_NO.Text = ""
            txtPO_SHIPMENT_NO.Text = ""
            txtPO_SHIPMENT_LNO.Value = 0
            txtCONTAINER_NO.Text = ""
            txtBOL_NO.Text = ""
            txtCOMM_INV_NO.Text = ""
        End If
    End Sub

    Private Sub cmdCreateAccrual_Click(sender As System.Object, e As System.EventArgs) Handles cmdCreateAccrual.Click

        Dim EMsg As String = ""
        Dim rowPOTCATG1 As DataRow = Nothing

        Dim COST_CATGY_CODE As String = Absx1.txtFor("COST_CATGY_CODE").Text
        If COST_CATGY_CODE = "" Then
            EMsg &= vbCr & "No Value Specified for Cost Category Code"
        Else
            rowPOTCATG1 = LookUp("POTCATG1", COST_CATGY_CODE)
            If rowPOTCATG1 Is Nothing Then
                EMsg &= vbCr & "Invalid Value Specified for Cost Category Code"
            End If
        End If

        Dim PO_ORDER_NO As String = txtPO_ORDER_NO.Text
        Dim PO_SHIPMENT_NO As String = txtPO_SHIPMENT_NO.Text
        Dim PO_SHIPMENT_LNO As Integer = Val(txtPO_SHIPMENT_LNO.Text)

        Dim rowPOTORDR1 As DataRow = Nothing
        Dim rowPOTSHIP1 As DataRow = Nothing
        Dim rowPOTSHIP2 As DataRow = Nothing

        If optAssociate.Value = "3" Then
            If PO_ORDER_NO = "" Then
                EMsg &= vbCr & "No Value Specified for PO No"
            Else
                rowPOTORDR1 = LookUp("POTORDR1", PO_ORDER_NO)
                If rowPOTORDR1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value Specified for PO No"
                End If
            End If
        ElseIf optAssociate.Value = "1" Then
            If PO_SHIPMENT_NO = "" Then
                EMsg &= vbCr & "No Value Specified for PO Shipment No"
            Else
                rowPOTSHIP1 = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                If rowPOTSHIP1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value Specified for PO Shipment No"
                End If
            End If
        ElseIf optAssociate.Value = "2" Then
            If PO_SHIPMENT_NO = "" Then
                EMsg &= vbCr & "No Value Specified for PO Shipment No"
            Else
                rowPOTSHIP1 = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                If rowPOTSHIP1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value Specified for PO Shipment No"
                Else
                    rowPOTSHIP2 = LookUp("POTSHIP2", New String() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
                    If rowPOTSHIP2 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Bill of Lading No"
                    End If
                End If
            End If
        End If

        If numCOST_ACT.Value = 0 Then
            EMsg &= vbCr & "Invalid Value Specified for Actual Cost"
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Create $0 Accrual")
            Exit Sub
        End If

        Dim rowPOTLCST1 As DataRow = dst.Tables("POTLCST1").NewRow
        Dim CTL_NO As String = ASCMAIN1.Next_Control_No("POTLCST1.CTL_NO")
        With rowPOTLCST1
            .Item("CTL_NO") = CTL_NO
            .Item("VEND_CODE") = Absx1.txtFor("VEND_CODE").Text
            .Item("PO_ORDER_NO") = PO_ORDER_NO
            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            .Item("COST_CATGY_CODE") = COST_CATGY_CODE
            .Item("CHARGEBACK_IND") = rowPOTCATG1.Item("CHARGEBACK_IND")
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("COST_ACT") = numCOST_ACT.Value

            If rowPOTSHIP1 IsNot Nothing Then
                .Item("PO_SHIP_VESSEL") = rowPOTSHIP1.Item("PO_SHIP_VESSEL")
                .Item("PO_SHIP_REF_NO") = rowPOTSHIP1.Item("PO_SHIP_REF_NO")
            End If
            If rowPOTSHIP2 IsNot Nothing Then
                .Item("CONTAINER_NO") = rowPOTSHIP2.Item("CONTAINER_NO")
                .Item("BOL_NO") = rowPOTSHIP2.Item("BOL_NO")
                .Item("COMM_INV_NO") = rowPOTSHIP2.Item("COMM_INV_NO")
            End If

        End With
        dst.Tables("POTLCST1").Rows.Add(rowPOTLCST1)

        Add_APTINVH7(CTL_NO)

        Dim PO_SHIPMENT_NO_SAVE As String = txtPO_SHIPMENT_NO.Text
        Clear_Other_Accrual_Controls()

        If optAssociate.Value = "1" Then
            txtPO_SHIPMENT_NO.Text = PO_SHIPMENT_NO_SAVE
        End If

    End Sub

    Sub Clear_Other_Accrual_Controls()
        optAssociate.Value = "1"
        Absx1.txtFor("COST_CATGY_CODE").Text = ""
        txtPO_ORDER_NO.Text = ""
        txtPO_SHIPMENT_NO.Text = ""
        txtPO_SHIPMENT_LNO.Text = ""
        numCOST_ACT.Value = 0
        txtCONTAINER_NO.Text = ""
        txtBOL_NO.Text = ""
        txtCOMM_INV_NO.Text = ""
    End Sub

    Sub Add_APTINVH7(CTL_NO As String)

        Dim rowPOTLCST1 As DataRow = dst.Tables("POTLCST1").Rows.Find(CTL_NO)
        Dim rowAPINVH7() As DataRow = dst.Tables("APTINVH7").Select("CTL_NO = '" & CTL_NO & "'")
        If rowAPINVH7.Length <> 0 Then
            MsgBox("Ctl No " & CTL_NO & " is already added to this Voucher", MsgBoxStyle.OkOnly, "Cannot add same records twice")
            Exit Sub
        End If

        rowPOTLCST1.Item("VOUCHER_NO") = HFs("VOUCHER_NO")

        Dim rowAPTINVH7 As DataRow = dst.Tables("APTINVH7").NewRow
        With rowAPTINVH7
            .Item("VOUCHER_NO") = HFs("VOUCHER_NO")
            Dim VOUCHER_CLNO As Integer = Val(dst.Tables("APTINVH7").Compute("MAX(VOUCHER_CLNO)", "") & "") + 1
            .Item("VOUCHER_CLNO") = VOUCHER_CLNO
            .Item("CTL_NO") = CTL_NO
            .Item("COST_CATGY_CODE") = rowPOTLCST1.Item("COST_CATGY_CODE")
            .Item("TOTAL_INV") = rowPOTLCST1.Item("COST_ACT")
            .Item("CHARGEBACK_IND") = rowPOTLCST1.Item("CHARGEBACK_IND")

            .Item("PO_ORDER_NO") = rowPOTLCST1.Item("PO_ORDER_NO")
            .Item("PO_SHIPMENT_NO") = rowPOTLCST1.Item("PO_SHIPMENT_NO")
            .Item("PO_SHIPMENT_LNO") = rowPOTLCST1.Item("PO_SHIPMENT_LNO")

            .Item("PO_SHIP_VESSEL") = rowPOTLCST1.Item("PO_SHIP_VESSEL")
            .Item("PO_SHIP_REF_NO") = rowPOTLCST1.Item("PO_SHIP_REF_NO")
            .Item("CONTAINER_NO") = rowPOTLCST1.Item("CONTAINER_NO")
            .Item("BOL_NO") = rowPOTLCST1.Item("BOL_NO")
            .Item("COMM_INV_NO") = rowPOTLCST1.Item("COMM_INV_NO")
        End With
        dst.Tables("APTINVH7").Rows.Add(rowAPTINVH7)
        Calc_DIST_Other()
    End Sub

    Private Sub grdPOTLCST1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTLCST1.DoubleClickRow
        Add_APTINVH7(e.Row.Cells("CTL_NO").Value)
    End Sub

    Private Sub grdPOTLCST1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTLCST1.InitializeRow
        If e.Row.Cells("VOUCHER_NO").Value & "" = HFs("VOUCHER_NO") Then
            e.Row.CellAppearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.CellAppearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

#Region "grdAPTINVH7"

    Private Sub grdAPTINVH7_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdAPTINVH7.AfterRowsDeleted

        Dim CTL_NOs As List(Of String) = DirectCast(grdAPTINVH7.Tag, List(Of String))
        For Each CTL_NO As String In CTL_NOs
            Dim rowPOTLCST1 As DataRow = dst.Tables("POTLCST1").Rows.Find(CTL_NO)
            rowPOTLCST1.Item("VOUCHER_NO") = ""
        Next

        Calc_DIST_Other()
    End Sub

    Private Sub grdAPTINVH7_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdAPTINVH7.AfterRowUpdate
        Dim CTL_NO As String = e.Row.Cells("CTL_NO").Value
        Dim rowPOTLCST1 As DataRow = dst.Tables("POTLCST1").Rows.Find(CTL_NO)
        rowPOTLCST1.Item("COST_ACT") = e.Row.Cells("TOTAL_INV").Value
        Calc_DIST_Other()
    End Sub

    Private Sub grdAPTINVH7_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdAPTINVH7.BeforeRowsDeleted
        Dim CTL_NOs As New List(Of String)
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim CTL_NO As String = grow.Cells("CTL_NO").Value
            CTL_NOs.Add(CTL_NO)
        Next
        grdAPTINVH7.Tag = CTL_NOs
    End Sub

#End Region

    Private Sub optPOReceipts_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPOReceipts.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        cbeYP.Visible = (optPOReceipts.Value & "" = "R")
        Load_ICTIREC1()
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
        If (tab0.SelectedTab.Key = "PO Receipts") Then
            Load_ICTIREC1()
        End If
    End Sub

    Sub Setup_tab0()
        With UltraExplorerBar1
            .Groups("Entry Options").Visible = Not ScreenMode And (tab0.SelectedTab.Key = "Invoices Entered")
            .Groups("PO Receipts").Visible = Not ScreenMode And (tab0.SelectedTab.Key = "PO Receipts")
        End With
    End Sub

    Private Sub grdICTIREC1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIREC1.DoubleClickRow
        If EntryMode <> "" Then Exit Sub

        If e.Row.IsDataRow Then
            Absx1.txtFor("VEND_CODE").Text = e.Row.Cells("VEND_CODE").Value & ""
            Absx1.txtFor("INV_NUM").Text = e.Row.Cells("COMM_INV_NO").Value & ""
            Dim RECEIPT_NO As String = e.Row.Cells("RECEIPT_NO").Value
            Click_Command("New")
            If ScreenMode Then
                tabMain.SelectedTab = tabMain.Tabs("PO Receipts")
                For Each grow As UltraWinGrid.UltraGridRow In grdICTIREC1.Rows
                    If grow.Cells("RECEIPT_NO").Value = RECEIPT_NO Then
                        grdICTIREC1.ActiveRow = grow
                        ProcessDoubleClickedRow()
                        Exit Sub
                    End If
                Next
            End If
        End If


    End Sub

    Private Sub cbeYP_ValueChanged(sender As Object, e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ICTIREC1()
    End Sub

    Sub Print_Check(COPY As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Check")

        Dim REPORTFILE As String = "APRCHKP1"
        Dim RPT As String = ""
        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowAPTINVH1.Item("BANK_CODE"))
        If rowGLTBANK1.Item("CHECK_REPORT") & "" <> "" Then
            RPT = rowGLTBANK1.Item("CHECK_REPORT")
        End If
        If RPT = "" Then RPT = REPORTFILE

        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {"", rowAPTINVH1.Item("BANK_CODE"), rowAPTINVH1.Item("CHECK_NUM")})
        Dim REPORT_NO As String = ""

        Dim make_pdf As Boolean = (COPY = "1")
        Dim FILENAME_body As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("COPY", COPY)
            'If make_pdf Then
            'REPORT_NO = .Generate_Report(RPT, "Check", , True, , , "PDF", FILENAME_body, False)
            'Show_Document(FILENAME_body)
            'Else
            REPORT_NO = .Generate_Report(RPT, "Check", , True, , , , , False)
            'End If
            '   .Print_Report_End(make_pdf, make_pdf)
            .Print_Report_End(False, False)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdCheck_Click(sender As System.Object, e As System.EventArgs) Handles cmdCheck.Click
        Print_Check("1")
    End Sub

    Function Add_APTINVH5_SUM(RECEIPT_NO As String, RECEIPT_DATE As Date, _
                              Optional QTY_REC As Int64 = 0, _
                              Optional AMT_REC As Decimal = 0) As DataRow
        Dim rowAPTINVH5_SUM As DataRow = dst.Tables("APTINVH5_SUM").NewRow
        rowAPTINVH5_SUM.Item("RECEIPT_NO") = RECEIPT_NO
        rowAPTINVH5_SUM.Item("RECEIPT_DATE") = RECEIPT_DATE
        rowAPTINVH5_SUM.Item("QTY_REC") = QTY_REC
        rowAPTINVH5_SUM.Item("AMT_REC") = AMT_REC
        rowAPTINVH5_SUM.Item("QTY_INV") = QTY_REC
        rowAPTINVH5_SUM.Item("AMT_INV") = AMT_REC

        'If RECEIPT_NO.StartsWith("S") Then ' this is a shipment, not a receipt
        'Else

        'End If
        Dim rowICTIREC1 As DataRow = dst.Tables("ICTIREC1").Rows.Find(RECEIPT_NO) ' LookUp("ICTIREC1", RECEIPT_NO)
        'If rowICTIREC1 IsNot Nothing Then

        'End If
        'Dim PO_ORDER_NO As String = rowICTIREC1.Item("PO_ORDER_NO")
        Dim PO_SHIPMENT_NO As String = rowICTIREC1.Item("PO_SHIPMENT_NO")
        Dim PO_SHIPMENT_LNO As Int64 = Val(rowICTIREC1.Item("PO_SHIPMENT_LNO") & "")
        Dim rowPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIPMENT_NO)
        Dim rowPOTSHIP2 As DataRow = LookUp("POTSHIP2", New String() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
        'Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)

        'rowAPTINVH5_SUM.Item("PO_ORDER_NO") = PO_ORDER_NO
        'rowAPTINVH5_SUM.Item("PO_REFERENCE") = rowPOTORDR1.Item("PO_REFERENCE")
        rowAPTINVH5_SUM.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
        rowAPTINVH5_SUM.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
        rowAPTINVH5_SUM.Item("PO_SHIP_VESSEL") = rowPOTSHIP1.Item("PO_SHIP_VESSEL")
        rowAPTINVH5_SUM.Item("PO_SHIP_REF_NO") = rowPOTSHIP1.Item("PO_SHIP_REF_NO")
        rowAPTINVH5_SUM.Item("PO_DATE_SHIPPED") = rowPOTSHIP1.Item("PO_DATE_SHIPPED")
        rowAPTINVH5_SUM.Item("PORT_CODE") = rowPOTSHIP1.Item("PORT_CODE")
        rowAPTINVH5_SUM.Item("WHSE_CODE") = rowPOTSHIP1.Item("WHSE_CODE")
        rowAPTINVH5_SUM.Item("CONTAINER_NO") = rowPOTSHIP2.Item("CONTAINER_NO")
        rowAPTINVH5_SUM.Item("BOL_NO") = rowPOTSHIP2.Item("BOL_NO")
        rowAPTINVH5_SUM.Item("COMM_INV_NO") = rowPOTSHIP2.Item("COMM_INV_NO")

        dst.Tables("APTINVH5_SUM").Rows.Add(rowAPTINVH5_SUM)

        Return rowAPTINVH5_SUM
    End Function

End Class