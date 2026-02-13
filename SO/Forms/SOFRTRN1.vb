Imports System.Drawing
Imports System.Math
Imports Infragistics.Win.UltraWinGrid

Public Class SOFRTRN1
    Private rowSOTRTRN1 As DataRow
    Private location_support As Boolean = False
    Private rowICTWHSE1 As DataRow
    Private whse_is_a_3PL As Boolean
    Private Whse_Rtn_no As String

    Private RECORD_INDEXs As List(Of Int32)
    Private INV_NO_RETURNED As String
    Private KEY_3PL_RECORD As String
    Private PRD_END_DATE As Date
    Private REFRESH_FLAG As Boolean = False

    Private rowARTCUST1 As DataRow
    Private rowSOTINVH1_Ret As DataRow
    Private tblSOTINVH2_Ret As DataTable
    Private priceChange As Boolean = False
    Private preloadInvoiceDetails As Boolean = False

    Private is180Customer As Boolean = False
    Private sqlSOTORDR2180 As String = String.Empty

    Private RA_NO As String = String.Empty
    Private RTRN_NO As String = String.Empty
    Private SO_PARM_WEB_INVOICES As String = String.Empty
    Private GL_PARM_CURR_CODE As String = String.Empty
    Private SO_PARM_MISC_CHG_RTN As String = String.Empty

    Private tab0SelectedTab As String = String.Empty

    Private WHSE_LOC_RFB As String = String.Empty
    Private WHSE_LOC_DST As String = String.Empty

    Private LoopReturnID As String = String.Empty

    'DROP TABLE SOTRMAFR CASCADE CONSTRAINTS ; 

    'CREATE TABLE SOTRMAFR ( 
    '  RA_NO                VARCHAR2 (6)  NOT NULL, 
    '  RA_RTN_LNO           NUMBER (3)    NOT NULL, 
    '  RA_RTN_QTY           NUMBER (7), 
    '  RA_UPC_CODE          VARCHAR2 (14), 
    '  RA_QTY_USED          NUMBER (7), 
    '  STYLE_CODE           VARCHAR2 (25), 
    '  COLOR_CODE           VARCHAR2 (6), 
    '  RA_RTN_STATUS        VARCHAR2 (1), 
    '  GUN_STATUS           VARCHAR2 (1), 
    '  RA_PUTAWAY_QTY_OPEN  NUMBER (7), 
    '  RA_PUTAWAY_LOC       VARCHAR2 (10), 
    '  INIT_OPER            VARCHAR2 (20), 
    '  INIT_DATE            DATE, 
    '  LAST_OPER            VARCHAR2 (20), 
    '  LAST_DATE            DATE, 
    '  PRIMARY KEY ( RA_NO, RA_RTN_LNO ) ) ;

    'ALTER TABLE SOTRMAF1 ADD RA_CARRIER_CODE VARCHAR2(6);
    'ALTER TABLE SOTRMAF1 ADD RA_CONTACT VARCHAR2(50);
    'ALTER TABLE SOTRMAF1 ADD RA_EMAIL VARCHAR2(100)	;
    'ALTER TABLE SOTRMAF1 ADD RA_PHONE VARCHAR2(20);
    'ALTER TABLE SOTRMAF1 ADD RA_CARTONS NUMBER(6);
    'ALTER TABLE SOTRMAF1 ADD REASON_CODE NUMBER(6);

    'ALTER TABLE SOTRTRN1 ADD RTRN_SALES_CURR NUMBER(13,2);
    'ALTER TABLE SOTRTRN1 ADD RTRN_STAX_CURR NUMBER(13,2);
    'ALTER TABLE SOTRTRN1 ADD RTRN_FREIGHT_CURR NUMBER(13,2);
    'ALTER TABLE SOTRTRN1 ADD RTRN_HANDLING_CURR NUMBER(13,2);
    'ALTER TABLE SOTRTRN1 ADD RTRN_AMOUNT_CURR NUMBER(13,2);
    'ALTER TABLE SOTRTRN1 ADD CURR_CODE VARCHAR2(3);
    'ALTER TABLE SOTRTRN1 ADD CURR_EXCH_RATE NUMBER(12,6);

    'ALTER TABLE SOTRTRN2 ADD RTRN_PRICE_CURR NUMBER(12,6);

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFRTRNI" Then
            InquiryMode = True
        End If

        ' Used in external class procedures
        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

        GL_PARM_CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & String.Empty
        SO_PARM_MISC_CHG_RTN = ROWs("SOTPARM1").Item("SO_PARM_MISC_CHG_RTN") & String.Empty

        Dim rowSOTMISC1 As DataRow = LookUp("SOTMISC1", SO_PARM_MISC_CHG_RTN)
        If rowSOTMISC1 Is Nothing Then
            SO_PARM_MISC_CHG_RTN = String.Empty
        End If

        Dim SO_PARM_WEB_INVOICES As String = (ROWs("SOTPARM1").Item("SO_PARM_WEB_INVOICES") & String.Empty).ToString.Trim
        If SO_PARM_WEB_INVOICES.Length > 0 AndAlso Not My.Computer.FileSystem.DirectoryExists(SO_PARM_WEB_INVOICES) Then
            SO_PARM_WEB_INVOICES = String.Empty
        End If
        If SO_PARM_WEB_INVOICES.Length > 0 AndAlso Not SO_PARM_WEB_INVOICES.EndsWith("\") Then
            SO_PARM_WEB_INVOICES &= "\"
        End If

        With dst
            ASCMAIN1.sql = "Select SOTRTRN1.*" _
            & " from SOTRTRN1 where SOTRTRN1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SOTRTRNX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select SOTRTRN3.*, GLTACCT1.ACCT_DESC" _
            & ", SOTRTRN1.RTRN_DATE, SOTRTRN1.WHSE_CODE, SOTRTRN1.REASON_CODE" _
            & ", SOTRTRN1.RTRN_NOTE, SOTRTRN1.INIT_OPER, SOTRTRN1.INIT_DATE" _
            & ", SOTRTRN1.RTRN_SOURCE, SOTRTRN1.OPS_YYYYPP" _
            & " from SOTRTRN1,SOTRTRN3,GLTACCT1 where SOTRTRN1.OPS_YYYYPP = :PARM1" _
            & " and GLTACCT1.ACCT_CODE = SOTRTRN3.ACCT_CODE" _
            & " and SOTRTRN3.RTRN_NO = SOTRTRN1.RTRN_NO"
            Create_TDA(.Tables.Add, "SOTRTRNG", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "SOTRTRN1", "*")
            Create_TDA(.Tables.Add("SOTRTRN1P"), "SOTRTRN1", "*")

            ASCMAIN1.sql = "Select SOTRTRN2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from SOTRTRN2,ICTSTYL1,ICTCOLR1 where ICTSTYL1.STYLE_CODE = SOTRTRN2.STYLE_CODE" _
            & " and ICTCOLR1.COLOR_CODE = SOTRTRN2.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTRTRN2", "**", 1)
            .Tables("SOTRTRN2").Columns.Add("RECORD_INDEX", GetType(System.Int32))
            .Tables("SOTRTRN2").Columns.Add("RTRN_QTY_REFUSED", GetType(System.Int32))
            .Tables("SOTRTRN2").Columns.Add("LINE_SALES", GetType(System.Decimal), "(ISNULL(RTRN_QTY, 0) - ISNULL(RTRN_QTY_REFUSED, 0)) * ISNULL(RTRN_PRICE,0)")
            .Tables("SOTRTRN2").Columns.Add("LINE_SALES_CURR", GetType(System.Decimal), "(ISNULL(RTRN_QTY, 0) - ISNULL(RTRN_QTY_REFUSED, 0))  * ISNULL(RTRN_PRICE_CURR,0)")
            .Tables("SOTRTRN2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "(ISNULL(RTRN_QTY, 0) - ISNULL(RTRN_QTY_REFUSED, 0))  * ISNULL(STYLE_COST,0)")
            .Tables("SOTRTRN2").Columns.Add("RTRN_QTY_TOTAL", GetType(System.Decimal), "ISNULL(RTRN_QTY_1,0) + ISNULL(RTRN_QTY_2,0) + ISNULL(RTRN_QTY_3,0) + ISNULL(RTRN_QTY_REFUSED,0)")
            .Tables("SOTRTRN2").Columns.Add("IMPORTED", GetType(System.String))
            .Tables("SOTRTRN2").Columns.Add("SURCHARGE_PERC", GetType(System.Decimal))
            .Tables("SOTRTRN2").Columns.Add("LINE_TARIFF", GetType(System.Decimal)) ', "Round(ISNULL(RTRN_PRICE,0) * (ISNULL(SURCHARGE_PERC,0) / 100), 2) * ISNULL(RTRN_QTY,0)")
            .Tables("SOTRTRN2").Columns.Add("MISC_CHG_CODE", GetType(System.String))
            .Tables("SOTRTRN2").Columns.Add("COUNTRY_CODE", GetType(System.String))
            .Tables("SOTRTRN2").Columns.Add("REFUND_TAX", GetType(System.Decimal))
            .Tables("SOTRTRN2").Columns.Add("WEB_RETURN_REASON", GetType(System.String))
            .Tables("SOTRTRN2").Columns.Add("LINE_ITEM_ID", GetType(System.String))
            .Tables("SOTRTRN2").Columns.Add("CUST_UPC", GetType(System.String))

            Create_TDA(.Tables.Add("SOTRTRN2P"), "SOTRTRN2", "**", 1)
            .Tables("SOTRTRN2P").Columns.Add("RECORD_INDEX", GetType(System.Int32))
            .Tables("SOTRTRN2P").Columns.Add("RTRN_QTY_REFUSED", GetType(System.Int32))
            .Tables("SOTRTRN2P").Columns.Add("LINE_SALES", GetType(System.Decimal), "(ISNULL(RTRN_QTY, 0) - ISNULL(RTRN_QTY_REFUSED, 0)) * ISNULL(RTRN_PRICE,0)")
            .Tables("SOTRTRN2P").Columns.Add("LINE_SALES_CURR", GetType(System.Decimal), "(ISNULL(RTRN_QTY, 0) - ISNULL(RTRN_QTY_REFUSED, 0)) * ISNULL(RTRN_PRICE_CURR,0)")
            .Tables("SOTRTRN2P").Columns.Add("LINE_COSTS", GetType(System.Decimal), "(ISNULL(RTRN_QTY, 0) - ISNULL(RTRN_QTY_REFUSED, 0)) * ISNULL(STYLE_COST,0)")
            .Tables("SOTRTRN2P").Columns.Add("RTRN_QTY_TOTAL", GetType(System.Decimal), "ISNULL(RTRN_QTY_1,0) + ISNULL(RTRN_QTY_2,0) + ISNULL(RTRN_QTY_3,0) + ISNULL(RTRN_QTY_REFUSED,0)")

            Create_TDA(.Tables.Add("SOTRTRN2_RPT"), "SOTRTRN2", "*")

            If Not .Tables("SOTRTRN2").Columns.Contains("RTN_LOCATION_CODE") Then
                .Tables("SOTRTRN2").Columns.Add("RTN_LOCATION_CODE", GetType(System.String))
                .Tables("SOTRTRN2P").Columns.Add("RTN_LOCATION_CODE", GetType(System.String))
                .Tables("SOTRTRN2_RPT").Columns.Add("RTN_LOCATION_CODE", GetType(System.String))
            End If

            .Tables("SOTRTRN2_RPT").Columns.Add("STYLE_DESC", GetType(System.String))
            .Tables("SOTRTRN2_RPT").Columns.Add("COLOR_DESC", GetType(System.String))

            Create_Relation("SOTRTRN1P", "SOTRTRN2P", "RTRN_NO")

            ASCMAIN1.sql = "Select SOTRTRN3.*, GLTACCT1.ACCT_DESC" _
            & " from SOTRTRN3,GLTACCT1 where GLTACCT1.ACCT_CODE = SOTRTRN3.ACCT_CODE"
            Create_TDA(.Tables.Add, "SOTRTRN3", "**", 1)

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
            & " from ICTSTAT2 where STYLE_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            .Tables.Add("SOTRTRN0")
            .Tables("SOTRTRN0").Columns.Add("KEY")
            .Tables("SOTRTRN0").Columns.Add("DESCRIPTION")

            'ASCMAIN1.sql = "Select * from ICTREAS1"
            'Create_TDA(.Tables.Add, "ICTREAS1", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            ASCMAIN1.sql = "SELECT * FROM ARTREAS1 WHERE NVL(RETURN_IND, '0') = '1'"
            Create_TDA(.Tables.Add, "ARTREAS1", "**", 0, False)

            Create_TDA(.Tables.Add, "ICTIADJ1", "*")
            Create_TDA(.Tables.Add, "ICTIADJ2", "*")
            .Tables("ICTIADJ2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(ADJ_QTY,0) * ISNULL(STYLE_COST,0)")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "SOTINVHM", "*")
            Create_TDA(.Tables.Add, "ICTWHSE1", "*")

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")

            Create_TDA(.Tables.Add, "SOTRMAFR", "*", 1)

            ASCMAIN1.sql = "Select SOTINVHM.*, SOTINVH2.STYLE_CODE from SOTINVHM, SOTINVH2
                             where SOTINVHM.INV_TYPE = 'I' and SOTINVHM.INV_NO = :PARM1 and SOTINVHM.MISC_CHARGE_TYPE = 'T'
                             and SOTINVH2.INV_TYPE = SOTINVHM.INV_TYPE and SOTINVH2.INV_NO = SOTINVHM.INV_NO 
                             and SOTINVH2.INV_LNO = SOTINVHM.INV_LNO"
            Create_TDA(.Tables.Add, "SOTINVHMR", "**", 0, False, "V", 3)

            ASCMAIN1.sql = "select RA_UPC_CODE ,STYLE_CODE ,COLOR_CODE ,
                            max(RA_RTN_QTY) RA_RTN_QTY ,
                            sum(RA_PUTAWAY_QTY_OPEN) RA_PUTAWAY_QTY_OPEN,
                            sum(RA_QTY_USED) RA_QTY_USED
                            from SOTRMAFR
                            where ra_no = :PARM1
                            and gun_status <> 'V'
                            group by RA_UPC_CODE ,STYLE_CODE ,COLOR_CODE 
                            order by Style_code"
            Create_TDA(.Tables.Add, "SOTRMAFRS", "**", 0, False, "V", 3)

            ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO" & vbCrLf _
                & ", SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & ", SOTINVH1.ORDR_NO, SOTINVH1.WHSE_CODE" & vbCrLf _
                & ", SOTINVH1.INV_SALES, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & ", SOTINVH1.ORDR_DEPT, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.CUST_BILL_TO_CUST" & vbCrLf _
                & " from SOTINVH1" & vbCrLf _
                & " where SOTINVH1.INV_TYPE = 'I' and SOTINVH1.CUST_CODE = :PARM1 and SOTINVH1.ORDR_YYYYPP_UPDATED = :PARM2"
            Create_TDA(.Tables.Add, "SOTINVHH", "**", 0, False, "VV", 2)

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO" & vbCrLf _
                & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_UNIT_PRICE_CURR, SOTINVH2.ORDR_QTY_SHIP" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & " from SOTINVH2,ICTSTYL1,ICTCOLR1,SOTINVH1" & vbCrLf _
                & " where SOTINVH2.INV_TYPE = 'I' and SOTINVH2.CUST_CODE = :PARM1 and SOTINVH2.ORDR_YYYYPP_UPDATED > :PARM2" & vbCrLf _
                & "  and SOTINVH2.STYLE_CODE = :PARM3 " & vbCrLf _
                & "  and ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE and ICTCOLR1.COLOR_CODE = SOTINVH2.COLOR_CODE" & vbCrLf _
                & "  and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "VVV", 3)

            ASCMAIN1.sql = "SELECT * FROM WHTWRTN1 WHERE WH_RTN_STATUS = 'C'"
            Create_TDA(.Tables.Add, "WHTWRTN1", "**", 0, False, , 1)


            ASCMAIN1.sql = "Select * from WHTWRTN1 "
            Create_TDA(.Tables.Add, "WHTWRTR1", "**", 0, False, , 1)


            ASCMAIN1.sql = " Select R2.WH_RTN_NO, R2.WH_RTN_LNO, R2. STYLE_CODE, IC.STYLE_DESC, COLOR_CODE, CTN_PACK_QTY, CARTONS, CTN_PACK_QTY * CARTONS" & vbCrLf _
            & " From WHTWRTN2 R2, ICTSTYL1 IC," & vbCrLf _
            & " (Select WH_RTN_NO, WH_RTN_LNO, COUNT(*) CARTONS from WHTWRTN3 Group By WH_RTN_NO, WH_RTN_LNO) R3" & vbCrLf _
            & " Where R2.WH_RTN_NO = R3.WH_RTN_NO" & vbCrLf _
            & " And R2.WH_RTN_LNO = R3.WH_RTN_LNO" _
            & " And R2.STYLE_CODE = IC.STYLE_CODE"
            Create_TDA(.Tables.Add, "WHTWRTR2", "**", 0, False, , 2)

            ASCMAIN1.sql = "SELECT SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, " _
                & " SOTORDR2.STYLE_CODE, SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE, " _
                & " SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NAME, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_CUST_PO," _
                & " SOTINVH1.INV_NO, SOTINVH1.INV_DATE " _
                & " FROM SOTORDR1, SOTORDR2, SOTINVH1 " _
                & " WHERE SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                & " AND SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO"

            sqlSOTORDR2180 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTORDR2180", ASCMAIN1.sql & " AND ROWNUM < 1", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "SOTORDR2", "*", , False)

            If ASCMAIN1.CLIENT = "NYA" Then
                Create_TDA(.Tables.Add, "EDTTRPM1", "*", 0)
                Fill_Records("EDTTRPM1", String.Empty, True, "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '180'")

                Create_TDA(.Tables.Add, "EDTSYSIH", "*")
                Create_TDA(.Tables.Add, "EDT180O1", "*")
                Create_TDA(.Tables.Add, "EDT180O2", "*")
            Else
                is180Customer = False
            End If

            ASCMAIN1.sql = "Select SOTRMAF1.*, ARTCUST1.CUST_NAME from SOTRMAF1, ARTCUST1 where RA_STATUS = 'O' AND SOTRMAF1.CUST_CODE = ARTCUST1.CUST_CODE (+)"
            Create_TDA(.Tables.Add, "SOTRMAFX", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select * from SOTRMAF2"
            Create_TDA(.Tables.Add, "SOTRMAFX2", "**", 0, False)

            ASCMAIN1.sql = "Select * from SOTRMAFR"
            Create_TDA(.Tables.Add, "SOTRMAFXR", "**", 0, False)

            Create_Relation("SOTRMAFX", "SOTRMAFX2", "RA_NO")
            Create_Relation("SOTRMAFX", "SOTRMAFXR", "RA_NO")

            .Tables("SOTRMAFX").Columns.Add("NUM_DETAILS", GetType(System.Int32), "COUNT(CHILD(SOTRMAFX_SOTRMAFX2).STYLE_CODE)")
            .Tables("SOTRMAFX").Columns.Add("NUM_SCANS", GetType(System.Int32), "COUNT(CHILD(SOTRMAFX_SOTRMAFXR).STYLE_CODE)")

            Create_TDA(.Tables.Add, "SOTRMAF1", "*", 1)

            ASCMAIN1.sql = "Select SOTRMAF2.*, ICTSTYL1.STYLE_DESC" _
                & " from SOTRMAF2, ICTSTYL1" _
                & " where ICTSTYL1.STYLE_CODE (+) = SOTRMAF2.STYLE_CODE"
            Create_TDA(.Tables.Add, "SOTRMAF2", "**", 1)
            .Tables("SOTRMAF2").Columns.Add("IMPORTED", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM SOTRTNL1 WHERE NVL(PROCESS_IND, '0') = '0'"
            Create_TDA(.Tables.Add, "SOTRTNL1", ASCMAIN1.sql, 0, True)

            ASCMAIN1.sql = "SELECT * FROM SOTRTNL2 WHERE RETURN_ID IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "SOTRTNL2", ASCMAIN1.sql, 0, False, "V")

            ASCMAIN1.sql = "SELECT * FROM SOTRTNL3 WHERE RETURN_ID IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "SOTRTNL3", ASCMAIN1.sql, 0, False, "V")

            ASCMAIN1.sql = "SELECT SOTRTNL1.RETURN_ID, SOTRTRN2.ORDR_NO, SOTRTRN2.ORDR_LNO, SOTRTNL2.RTRN_QTY,
                                (SOTRTRN2.RTRN_QTY - NVL(RTRN_QTY_1, 0) - NVL(RTRN_QTY_2, 0) - NVL(RTRN_QTY_3, 0)) QTY_REFUSED
                                FROM SOTRTRN2, SOTRTNL1, SOTRTNL2
                                WHERE SOTRTRN2.RTRN_NO = SOTRTNL1.RTRN_NO
                                AND SOTRTNL1.RETURN_ID = SOTRTNL2.RETURN_ID
                                AND SOTRTNL2.ORDR_NO = SOTRTRN2.ORDR_NO
                                AND SOTRTNL2.ORDR_LNO = SOTRTRN2.ORDR_LNO
                                AND SOTRTRN2.RTRN_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTRTNL2_REF", ASCMAIN1.sql, 0, False, "V")


            Create_TDA(.Tables.Add, "ECTECOMD", "*")
            Fill_Records("ECTECOMD", "", True, "Select * from ECTECOMD")

        End With

        Fill_Records("ARTREAS1")
        Fill_Records("ICTCLAS1")

        Show_Filter(grdSOTINVHH, True)

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -48) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        cbeInvoiceHistory.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -48) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeInvoiceHistory.SelectedItem = cbeInvoiceHistory.Items(3)

        grdSOTRTRN0.DataSource = dst.Tables("SOTRTRN0")
        grdSOTRTRN2.DataSource = dst.Tables("SOTRTRN2")
        grdSOTRTRN3.DataSource = dst.Tables("SOTRTRN3")
        grdSOTRTRNX.DataSource = dst.Tables("SOTRTRNX")
        grdSOTRTRNG.DataSource = dst.Tables("SOTRTRNG")
        grdSOTRTRN1P.DataSource = dst.Tables("SOTRTRN1P")
        grdSOTRMAFR.DataSource = dst.Tables("SOTRMAFR")
        grdSOTRMAFRS.DataSource = dst.Tables("SOTRMAFRS")

        grdSOTINVHH.DataSource = dst.Tables("SOTINVHH")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")
        grdWHTRTRN1.DataSource = dst.Tables("WHTWRTN1")

        grdSOTORDR2180.DataSource = dst.Tables("SOTORDR2180")

        Create_Relation("SOTRTNL1", "SOTRTNL2", "RETURN_ID", "RETURN_ID")
        Create_Relation("SOTRTNL1", "SOTRTNL3", "RETURN_ID", "RETURN_ID")
        grdSOTRTNL1.DataSource = dst.Tables("SOTRTNL1")

        ASCMAIN1.Add_Value_List(grdSOTRTNL1, "REFUND_BEFORE_INSPECTION", , New String() {":", "0:False", "1:True"})
        ASCMAIN1.Add_Value_List(grdSOTRTNL1, "WAS_PROCESSED", , New String() {":", "0:False", "1:True"})
        ASCMAIN1.Add_Value_List(grdSOTRTNL1, "MULTI_CURRENCY", , New String() {":", "0:False", "1:True"})

        Create_Summary(grdSOTRTRNX, "RTRN_NO", "Count")
        Create_Summary(grdSOTRTRNX, New String() {"RTRN_SALES", "RTRN_COSTS", "RTRN_STAX", "RTRN_FREIGHT", "RTRN_HANDLING", "RTRN_AMOUNT"})

        Create_Summary(grdSOTRTRNG, "RTRN_NO", "Count")
        Create_Summary(grdSOTRTRNG, "DIST_AMT")

        Create_Summary(grdSOTRTRN2, "RTRN_LNO", "Count")
        Create_Summary(grdSOTRTRN2, New String() {"RTRN_QTY", "RTRN_QTY_1", "RTRN_QTY_2", "RTRN_QTY_3", "RTRN_QTY_REFUSED", "LINE_SALES", "LINE_SALES_CURR", "LINE_COSTS", "LINE_TARIFF", "REFUND_TAX"})

        Create_Summary(grdSOTRTRN3, "RTRN_GNO", "Count")
        Create_Summary(grdSOTRTRN3, "DIST_AMT")

        Create_Summary(grdSOTRMAFR, "RA_RTN_LNO", "Count")

        With grdSOTRTRNX.DisplayLayout.Bands("SOTRTRNX")
            .Columns("RTRN_NO").Header.Fixed = True
        End With

        With grdSOTRTRNG.DisplayLayout.Bands("SOTRTRNG")
            .Columns("RTRN_NO").Header.Fixed = True
        End With

        With grdSOTRTRN1P.DisplayLayout.Bands("SOTRTRN1P")
            .Columns("RTRN_NO").Header.Fixed = True
        End With

        With grdSOTRMAFR.DisplayLayout.Bands(0)
            .Columns("RA_NO").Header.Fixed = True
            .Columns("RA_RTN_LNO").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdSOTRTRNX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ARTREAS1 order by REASON_DESC")
        ASCMAIN1.Add_Value_List(grdSOTRTRNG, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ARTREAS1 order by REASON_DESC")
        ASCMAIN1.Add_Value_List(grdWHTRTRN1, "WH_RTN_STATUS", , New String() {":", "S:SAVED", "C:COMPLETED"})
        ASCMAIN1.Add_Value_List(grdWHTRTRN1, "CUST_CODE", "Select CUST_CODE, CUST_NAME from ARTCUST1")

        ASCMAIN1.Add_Value_List(grdSOTRMAFR, "RA_RTN_STATUS", , New String() {":", "1:Return To Stock", "2:Refurbish", "3:Destroy"})

        For Each grdBand As Infragistics.Win.UltraWinGrid.UltraGridBand In grdSOTRTNL1.DisplayLayout.Bands
            grdBand.Hidden = False
            For Each gcol As UltraWinGrid.UltraGridColumn In grdBand.Columns
                gcol.Hidden = False
                gcol.Header.Caption = StrConv(gcol.Header.Caption.Replace("_", " "), VbStrConv.ProperCase)
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        Next
        grdSOTRTNL1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        Create_Summary(grdSOTRTNL1, "RETURN_ID", "Count")

        grdSOTRTRN0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Set_SEGS(grdSOTRTRN3, "SOTRTRN3")

        Bind_Controls(grpTotals, "SOTRTRN1")
        'Set_Read_Only(grpTotals, True)
        If InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") = 0 Then
            'grpTotals.Visible = False
            lblRTRN_COSTS.Visible = False
            numRTRN_COSTS.Visible = False
            With grdSOTRTRN2.DisplayLayout.Bands(0)
                .Columns("STYLE_COST").Hidden = True
                .Columns("LINE_COSTS").Hidden = True
                .Columns("STYLE_CLASS_CODE").Hidden = True
                .Columns("SALES_DIVISION_CODE").Hidden = True
            End With
        End If

        ' Setup grdSOTRTRN1P bad 1 to look like grdSOTRTRN2
        For Each grdColumn As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdSOTRTRN1P.DisplayLayout.Bands(1).Columns
            grdColumn.Hidden = True
        Next

        For Each grdColumn As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdSOTRTRN2.DisplayLayout.Bands(0).Columns
            Dim key As String = grdColumn.Key
            For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTRTRN1P.DisplayLayout.Bands(1).Columns
                If gcol.Key = key Then
                    gcol.Hidden = grdColumn.Hidden
                    gcol.Width = grdColumn.Width
                    gcol.Header.Caption = grdColumn.Header.Caption
                    gcol.Header.VisiblePosition = grdColumn.Header.VisiblePosition
                    Exit For
                End If
            Next
        Next

        For Each band As Infragistics.Win.UltraWinGrid.UltraGridBand In New Infragistics.Win.UltraWinGrid.UltraGridBand() {grdSOTRTRN2.DisplayLayout.Bands(0), grdSOTRTRN1P.DisplayLayout.Bands(1)}
            With band ' grdSOTRTRN2.DisplayLayout.Bands(0)
                .Columns("RTRN_QTY_1").Header.Caption = "Stock"
                .Columns("RTRN_QTY_2").Header.Caption = "Refurb"
                .Columns("RTRN_QTY_3").Header.Caption = "Destroy"
                .Columns("RTRN_QTY_REFUSED").Header.Caption = "Refused"

                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackColor = Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"RTRN_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC", "STYLE_CLASS_CODE", "SALES_DIVISION_CODE"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Color.LightGray
                    End If

                    If New String() {"RTRN_QTY", "RTRN_QTY_1", "RTRN_QTY_2", "RTRN_QTY_3", "RTRN_QTY_REFUSED"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Color.LightBlue
                        If gcol.Key = "RTRN_QTY" Then
                            gcol.CellAppearance.BackColor = Color.LightBlue
                        End If
                        gcol.Width = 70
                    End If

                    If New String() {"RTRN_PRICE", "STYLE_COST", "LINE_SALES", "LINE_COSTS", "LINE_TARIFF"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Color.LightGreen
                        If gcol.Key.StartsWith("LINE") Then
                            gcol.CellAppearance.BackColor = Color.LightGreen
                            gcol.Width = 90
                        Else
                            gcol.Width = 70
                        End If
                        If ASCMAIN1.CLIENT <> "RGI" And gcol.Key.Contains("LINE_TARIFF") Then
                            gcol.Hidden = True
                        End If
                    End If

                    If gcol.Key.EndsWith("_CURR") Then
                        gcol.Header.Appearance.BackColor2 = Color.LightBlue
                        If gcol.Key.StartsWith("LINE") Then
                            gcol.CellAppearance.BackColor = Color.LightBlue
                            gcol.Width = 90
                        Else
                            gcol.Width = 70
                        End If
                    End If
                Next

                For Each COLUMN_NAME As String In New String() {"RTRN_PRICE", "RTRN_PRICE_CURR"}
                    .Columns(COLUMN_NAME).Format = "####.000"
                    .Columns(COLUMN_NAME).MaskInput = "nnnn.nnn"
                Next
            End With
        Next

        grdSOTRMAFX.DataSource = dst.Tables("SOTRMAFX")
        grdSOTRMAFX.DisplayLayout.UseFixedHeaders = True
        With grdSOTRMAFX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RA_NO", "CUST_CODE", "CUST_NAME"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        ASCMAIN1.Add_Value_List(grdSOTRMAFX, "RA_RTN_STATUS", , New String() {":", "1:Stock", "3:Damage"}, 2)

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
        PRD_END_DATE = rowGLTPARM2.Item("PRD_END_DATE")

        If ASCMAIN1.CLIENT = "RGI" Then
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY_2").Hidden = True
            grdSOTRTRN1P.DisplayLayout.Bands(1).Columns("RTRN_QTY_2").Hidden = True
        End If

        Absx1.dteFor("SEARCH_START_DATE").MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        Absx1.dteFor("SEARCH_START_DATE").MinDate = DateAdd(DateInterval.Year, -3, DateTime.Now)
        Absx1.dteFor("SEARCH_START_DATE").DateTime = DateAdd(DateInterval.Month, -6, DateTime.Now)

        grpHeader.Visible = False
        Set_SEGS(grdSOTRTRNG, "SOTRTRNG")

        Dim btn As New UltraWinEditors.EditorButton
        numRTRN_FREIGHT.ButtonsLeft.Add(btn)
        btn.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "ARROW_UP_BLUE")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New", "Price Change"

                RA_NO = String.Empty
                RTRN_NO = String.Empty

                WHSE_LOC_RFB = String.Empty
                WHSE_LOC_DST = String.Empty

                Absx1.txtFor("RA_NO").Text = Absx1.txtFor("RA_NO").Text.Trim

                If Absx1.txtFor("RA_NO").TextLength > 0 Then
                    Dim rowSOTRMAF1 As DataRow = Nothing

                    If IsNumeric(Absx1.txtFor("RA_NO").Text) Then
                        Absx1.txtFor("RA_NO").Text = ASCMAIN1.Format_Field(Absx1.txtFor("RA_NO").Text, "RA_NO")
                    End If

                    rowSOTRMAF1 = LookUp("SOTRMAF1", Absx1.txtFor("RA_NO").Text)
                    If rowSOTRMAF1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Return Authorization no: " & Absx1.txtFor("RA_NO").Text
                        Exit Select
                    ElseIf rowSOTRMAF1.Item("RA_STATUS") & String.Empty <> "O" Then
                        EMsg &= vbCr & "Return Authorization no: " & Absx1.txtFor("RA_NO").Text & " is not Open"
                        Exit Select
                    ElseIf IsDate(rowSOTRMAF1.Item("RA_EXPIRE") & String.Empty) AndAlso DateTime.Now.CompareTo(rowSOTRMAF1.Item("RA_EXPIRE")) = 1 Then
                        EMsg &= vbCr & "Return Authorization no: " & Absx1.txtFor("RA_NO").Text & " is Expired"
                        Exit Select
                    Else
                        RA_NO = Absx1.txtFor("RA_NO").Text
                    End If

                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("SOTRMAF1", RA_NO) Then Exit Sub
                        If rowSOTRMAF1 Is Nothing Then
                            If Not ASCMAIN1.Logical_Lock("SOTRMAF1", "CUST_CODE_" & MyBase.Absx1.txtFor("CUST_CODE").Text) Then Exit Sub
                        Else
                            If Not ASCMAIN1.Logical_Lock("SOTRMAF1", "CUST_CODE_" & rowSOTRMAF1.Item("CUST_CODE")) Then Exit Sub
                        End If
                    Else
                        Exit Select
                    End If

                    If rowSOTRMAF1 IsNot Nothing Then
                        Absx1.txtFor("CUST_CODE").Text = rowSOTRMAF1.Item("CUST_CODE") & String.Empty
                        Absx1.txtFor("WHSE_CODE").Text = rowSOTRMAF1.Item("WHSE_CODE") & String.Empty
                    End If

                    If Not IsDate(MyBase.Absx1.dteFor("RTRN_DATE").Value) Then
                        MyBase.Absx1.dteFor("RTRN_DATE").Value = DateTime.Now
                    End If

                Else

                    rowSOTINVH1_Ret = Nothing
                    tblSOTINVH2_Ret = Nothing

                    ' Format the Inv No if provided
                    txtINV_NO_RETURNED.Text = txtINV_NO_RETURNED.Text.Trim
                    If txtINV_NO_RETURNED.TextLength > 0 Then
                        txtINV_NO_RETURNED.Text = ASCMAIN1.Format_Field(txtINV_NO_RETURNED.Text, "INV_NO")
                    End If

                    Dim INV_NO_RETURNED As String = txtINV_NO_RETURNED.Text
                    INV_NO_RETURNED = INV_NO_RETURNED.Trim

                    If ASCMAIN1.CLIENT = "RGI" AndAlso INV_NO_RETURNED.Length = 0 Then
                        EMsg &= vbCr & "You must select and Invoice to Credit."
                        Exit Select
                    End If

                    is180Customer = False
                    If ASCMAIN1.CLIENT = "NYA" AndAlso dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'").Length > 0 Then
                        If eItemKey = "New" Then
                            If INV_NO_RETURNED.Length = 0 Then
                                EMsg &= vbCr & "Invoice Returned is required for all EDI 180 customers"
                                Exit Select
                            End If
                        End If

                        is180Customer = True
                    End If

                    If INV_NO_RETURNED.Length > 0 Then
                        rowSOTINVH1_Ret = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_TYPE = :PARM1 AND INV_NO = :PARM2", "VV", {"I", INV_NO_RETURNED})
                        If rowSOTINVH1_Ret Is Nothing Then
                            EMsg &= vbCr & "Invalid Invoice Number"
                        ElseIf rowSOTINVH1_Ret.Item("INV_TYPE") <> "I" Then
                            EMsg &= vbCr & "The provided Invoice Number is not a Sales Invoice"
                            ' ElseIf rowSOTINVH1_Ret.Item("ORDR_TYPE_CODE") & String.Empty <> "REG" Then
                            ' EMsg &= vbCr & "The provided Invoice Number has an Order Type (" & rowSOTINVH1_Ret.Item("ORDR_TYPE_CODE") & ") that does not permit Returns"
                        End If

                        If EMsg.Length > 0 Then
                            Exit Select
                        End If

                        ' See if there are any other credits for this invoice
                        ASCMAIN1.sql = "select * from sotrtrn1 where INV_NO_RETURNED = '" & INV_NO_RETURNED & "' and cust_code = '" & rowSOTINVH1_Ret.Item("CUST_CODE") & "'"
                        Dim tblPrevious As DataTable = ASCDATA1.GetDataTable()
                        If tblPrevious.Rows.Count > 0 Then
                            Dim pMsg As String = "The following returns are also for this Invoice:" & vbCr

                            pMsg &= vbCr & "Return No     Date                   Return Amount"

                            For Each rowPrevious As DataRow In tblPrevious.Select("", "RTRN_NO")
                                Dim danac As String = rowPrevious.Item("RTRN_NO")
                                pMsg &= vbCr & rowPrevious.Item("RTRN_NO") & "   " & Format(rowPrevious.Item("RTRN_DATE"), "MM/dd/yyyy") & "        " & Format(Val(rowPrevious.Item("RTRN_SALES") & ""), "#,##0.00")
                            Next

                            pMsg &= vbCr & vbCr & "Do you want to continue?"

                            If MessageBox.Show(pMsg, "New Return", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                Exit Sub
                            End If

                        End If

                        ASCMAIN1.sql = "SELECT SOTINVH2.*, SOTINVH1.ORDR_NO
                                            FROM SOTINVH1, SOTINVH2
                                            WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE
                                            AND SOTINVH1.INV_NO = SOTINVH2.INV_NO
                                            AND SOTINVH1.INV_TYPE = :PARM1
                                            AND SOTINVH1.INV_NO = :PARM2"

                        tblSOTINVH2_Ret = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", {"I", INV_NO_RETURNED})
                        If tblSOTINVH2_Ret.Rows.Count = 0 Then
                            EMsg &= vbCr & "The provided Invoice does not have details"
                            Exit Select
                        End If

                        Absx1.txtFor("WHSE_CODE").Text = rowSOTINVH1_Ret.Item("WHSE_CODE") & String.Empty
                        Absx1.txtFor("CUST_CODE").Text = rowSOTINVH1_Ret.Item("CUST_CODE") & String.Empty

                    End If

                    If EMsg.Length > 0 Then
                        Exit Select
                    End If
                End If

                Validate_Code("CUST_CODE")

                If EMsg.Length = 0 Then
                    If cdr.Item("CURR_CODE") & String.Empty <> "" AndAlso cdr.Item("CURR_CODE") & String.Empty <> GL_PARM_CURR_CODE Then
                        ' EMsg &= vbCr & "Currently, you are not permitted to process returns for non USD customers. Please call ABS."
                        ' Foreign Exchnages require an Invoice No.
                        If txtINV_NO_RETURNED.Text.Trim.Length = 0 Then
                            EMsg &= vbCr & "Returns for customers with a Foreign Currency require an Invoice Number."
                            Exit Select
                        End If
                    End If
                End If

                Validate_Code("WHSE_CODE")

                If Absx1.dteFor("RTRN_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Date Specified for Entry"
                End If

                If Absx1.txtFor("WHSE_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        If ASCMAIN1.CLIENT = "VAN" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                        Else
                            If rowSOTINVH1_Ret Is Nothing Then
                                EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                            Else
                                EMsg &= vbCr & "The provided Invoice has an invalid Warehouse"
                            End If
                        End If
                    Else
                        If rowICTWHSE1.Item("LP_CODE") & "" <> "" And KEY_3PL_RECORD = "" Then
                            If ASCMAIN1.CLIENT = "NYA" Then
                                ' NO MSGBOX REQUIRED
                            Else
                                MsgBox("You are entering a Customer Credit involving a 3PL warehouse" _
                                       & vbCrLf & vbCrLf & "You must choose a reason code that does NOT impact inventory",
                                       MsgBoxStyle.OkOnly, "Verification")
                                '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Adjustments Allowed"
                            End If
                        End If

                        WHSE_LOC_RFB = rowICTWHSE1.Item("WHSE_LOC_RFB") & String.Empty
                        WHSE_LOC_DST = rowICTWHSE1.Item("WHSE_LOC_DST") & String.Empty
                    End If
                End If

                If EMsg.Length = 0 Then
                    If ASCMAIN1.CLIENT = "VAN" AndAlso LoopReturnID.Length > 0 Then
                        If Not ASCMAIN1.Logical_Lock("SOTRTNL1", LoopReturnID) Then
                            LoopReturnID = String.Empty
                            Exit Sub
                        End If

                        ' Verify it is still open
                        Dim drSOTRTNL1 As DataRow = LookUp("SOTRTNL1", LoopReturnID)
                        If drSOTRTNL1 Is Nothing Then
                            EMsg &= vbCr & $"Web Return No ({LoopReturnID}) is invalid."
                            LoopReturnID = String.Empty
                            Exit Select
                        End If

                        If drSOTRTNL1.Item("PROCESS_IND") & String.Empty = "1" Then
                            EMsg &= vbCr & $"Web Return No ({LoopReturnID}) was already processed."
                            LoopReturnID = String.Empty
                            Exit Select
                        End If
                    End If
                End If

            Case "View"
                LoopReturnID = String.Empty
                If Absx1.txtFor("RTRN_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    Dim RTRN_NO As String = ASCMAIN1.Format_Field(Absx1.txtFor("RTRN_NO").Text, "RTRN_NO")
                    Absx1.txtFor("RTRN_NO").Text = RTRN_NO

                    rowSOTRTRN1 = LookUp("SOTRTRN1", RTRN_NO)
                    If rowSOTRTRN1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Return Document " & RTRN_NO & " on File"
                    ElseIf ASCMAIN1.CLIENT = "VAN" Then
                        Dim drSOTRTRL1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTRTNL1 WHERE RTRN_NO = :PARM1", "V", {RTRN_NO})
                        If drSOTRTRL1 IsNot Nothing Then
                            LoopReturnID = drSOTRTRL1.Item("RETURN_ID") & String.Empty
                        End If
                    End If
                End If

            Case "Update"
                If Absx1.txtFor("REASON_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Reason"
                Else
                    Dim rowARTREAS1 As DataRow

                    If ASCMAIN1.CLIENT = "VAN" Then
                        rowARTREAS1 = LookUp("ARTREAS1", Absx1.txtFor("REASON_CODE").Text)
                    Else
                        rowARTREAS1 = ASCDATA1.GetDataRow("SELECT * FROM ARTREAS1 WHERE NVL(RETURN_IND, '0') = '1' AND REASON_CODE = :PARM1", "V", Absx1.txtFor("REASON_CODE").Text)
                    End If

                    If rowARTREAS1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Return Reason"
                    Else
                        If rowARTREAS1.Item("RETURN_NO_STOCK_IND") & "" = "1" Then
                            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTRTRN2"), New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
                                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                                Dim sqlw As String = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
                                If Val(dst.Tables("SOTRTRN2").Compute("SUM(RTRN_QTY_1)", sqlw) & "") <> 0 _
                                    OrElse Val(dst.Tables("SOTRTRN2").Compute("SUM(RTRN_QTY_2)", sqlw) & "") <> 0 Then
                                    EMsg &= vbCr & "Non-Zero Qty impacting inventory for Style-Color " & STYLE_CODE & "-" & COLOR_CODE
                                End If
                            Next
                        Else
                            If ASCMAIN1.CLIENT = "VAN" Then
                                If whse_is_a_3PL And KEY_3PL_RECORD = "" Then
                                    EMsg &= vbCr & "Cannot Credit a Customer in a 3PL warehouse using a Reason Code which impacts Inventory"
                                End If
                            End If
                        End If
                    End If
                End If

                If Absx1.txtFor("WHSE_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        If ASCMAIN1.CLIENT = "VAN" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                        Else
                            If rowSOTINVH1_Ret Is Nothing Then
                                EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                            Else
                                EMsg &= vbCr & "The provided Invoice has an invalid Warehouse"
                            End If
                        End If
                    Else
                        If rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        Else
                            If rowICTWHSE1.Item("LP_CODE") & "" <> "" And KEY_3PL_RECORD = "" Then
                                If ASCMAIN1.CLIENT = "NYA" Then
                                    ' NO MSGBOX REQUIRED
                                Else
                                    MsgBox("You are entering a Customer Credit involving a 3PL warehouse" _
                                           & vbCrLf & vbCrLf & "You must choose a reason code that does NOT impact inventory",
                                           MsgBoxStyle.OkOnly, "Verification")
                                    '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Adjustments Allowed"
                                End If
                            End If
                        End If
                    End If
                End If

                If Absx1.txtFor("CUST_STORE_NO").Text = "" Then
                    ' EMsg &= vbCr & "You Must Specify a Customer Store"
                Else
                    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, "MK", Absx1.txtFor("CUST_STORE_NO").Text})
                    If rowARTCUST2 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Customer Store"
                    Else
                        dst.Tables("SOTRTRN1").Rows(0).Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_NAME")
                    End If
                End If

                If grdSOTRTRN2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("", "", DataViewRowState.CurrentRows)
                        If rowSOTRTRN2.Item("STYLE_CLASS_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Class for " & rowSOTRTRN2.Item("STYLE_CODE") & ""
                        End If
                        If rowSOTRTRN2.Item("SALES_DIVISION_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Division for " & rowSOTRTRN2.Item("STYLE_CODE") & ""
                        End If

                        If ASCMAIN1.CLIENT = "RGI" Then
                            If Val(rowSOTRTRN2.Item("RTRN_QTY_3") & "") <> 0 Then
                                If rowSOTRTRN2.Item("RTV_REASON_CODE") & "" = "" Then
                                    EMsg &= vbCr & "An RTV Reason is Required when the Destroyed Qty is not 0 - see Style " & rowSOTRTRN2.Item("STYLE_CODE") & ""
                                End If
                            End If
                        End If

                        If location_support Then
                            If Val(rowSOTRTRN2.Item("RTRN_QTY_3") & "") <> 0 Then
                                If WHSE_LOC_DST.Length = 0 Then
                                    EMsg = vbCr & "The warehouse for this return does not have an assigned Destroy Location."
                                    Exit Select
                                End If
                            End If

                            If Val(rowSOTRTRN2.Item("RTRN_QTY_2") & "") <> 0 Then
                                If WHSE_LOC_RFB.Length = 0 Then
                                    EMsg = vbCr & "The warehouse for this return does not have an assigned Refurbish Location."
                                    Exit Select
                                End If
                            End If
                        End If
                    Next
                End If

                If dst.Tables("SOTRTRN2").Select("ISNULL(RTRN_QTY,0) <> ISNULL(RTRN_QTY_TOTAL,0)").Length <> 0 Then
                    EMsg &= vbCr & "Some lines are out of Balance"
                End If

                If dst.Tables("SOTRTRN2").Select("RTRN_PRICE IS NULL").Length <> 0 Then
                    If MsgBox("Some lines do not have Price, OK to  Continue?", MsgBoxStyle.YesNo,
                        "Missing prices") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

                If Val(Absx1.numFor("RTRN_AMOUNT").Value & "") < 0 Then
                    EMsg &= vbCr & "Total Amount is not a Credit to the Customer"
                End If

                If Absx1.dteFor("RTRN_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Return Date is Mandatory"
                ElseIf Format(Absx1.dteFor("RTRN_DATE").Value, "yyyyMMdd") > Format(PRD_END_DATE, "yyyyMMdd") Then
                    EMsg &= vbCr & "Return Date may not be Later than " & Format(PRD_END_DATE, "MM/dd/yyyy")
                ElseIf Format(Absx1.dteFor("RTRN_DATE").Value, "yyyyMMdd") < Format(PRD_END_DATE, "yyyyMM") & "01" Then
                    ' this change was put in place for VAN, but probably should also be enforced for RGI, NYA, etc.
                    EMsg &= vbCr & "Return Date may not be earlier than " & Format(PRD_END_DATE, "MM/01/yyyy")
                End If

                Dim INV_MISC_CHG As Decimal = Val(Absx1.numFor("RTRN_HANDLING").Value & String.Empty)
                If INV_MISC_CHG > 0 AndAlso EMsg.Length > 0 AndAlso ASCMAIN1.CLIENT <> "RGI" Then
                    Dim MISC_CHG_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_MISC_CHG_RTN") & String.Empty
                    If MISC_CHG_CODE.Length = 0 Then
                        EMsg &= vbCr & "Handling Charge requires a Misc Charge Return Code in the SO Parameters"
                    End If
                End If

                If EMsg.Length = 0 Then
                    Dim RTRN_FREIGHT As Decimal = Val(Absx1.numFor("RTRN_FREIGHT").Value & "")
                    If RTRN_FREIGHT <> 0 Then
                        If MessageBox.Show("Do you want to credit freight in the amount of " & RTRN_FREIGHT.ToString("#,##0.00") & "?", "Return Freight", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                If EMsg.Length = 0 Then
                    'ValidateOverCredited()
                    ValidateOverCreditedRMA()

                    If EMsg.Length > 0 Then
                        If MessageBox.Show("The following is a list of over credits:" & Environment.NewLine & Environment.NewLine _
                                            & EMsg & Environment.NewLine & Environment.NewLine _
                                            & "Do you want to proceed with the Update?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Error) = Windows.Forms.DialogResult.No Then
                            EMsg = String.Empty
                            Exit Sub
                        End If
                    End If

                    EMsg = String.Empty
                End If

                'If EMsg.Length = 0 Then
                '    If ASCMAIN1.CLIENT = "VAN" AndAlso LoopReturnID.Length > 0 Then
                '        Absx1.txtFor("RTRN_NOTE").Text = Absx1.txtFor("RTRN_NOTE").Text.Trim
                '        If Absx1.txtFor("RTRN_NOTE").TextLength = 0 Then
                '            If dst.Tables("SOTRTRN2").Select("ISNULL(RTRN_QTY_REFUSED,0) > 0").Length <> 0 Then
                '                EMsg &= vbCr & "When you Refuse Items on the return you must provide a Note."
                '                Exit Select
                '            End If
                '        End If
                '    End If
                'End If

                If EMsg.Length = 0 Then
                    If MessageBox.Show("Do you want to Update this Return?", "Update",
                                       MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Reverse"
                If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Back to Stock Report"
                If Val(dst.Tables("SOTRTRN2").Compute("SUM(RTRN_QTY_1)", "") & String.Empty) = 0 Then
                    EMsg &= vbCr & "The total quantity in the Stock Column is 0."
                End If

        End Select

        If EMsg <> "" Then
            If eItemKey = "New" OrElse eItemKey = "View" Then
                LoopReturnID = String.Empty
            End If
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New", "Price Change"

                priceChange = (eItemKey = "Price Change")
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

                Select Case dst.Tables("SOTRTRN1").Rows(0).Item("CURR_CODE") & String.Empty
                    Case "", GL_PARM_CURR_CODE
                        ' Nothing at this time
                    Case Else
                        MessageBox.Show("You must enter all prices in " & dst.Tables("SOTRTRN1").Rows(0).Item("CURR_CODE") & String.Empty & " currency.", "Foreign Return", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Select

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                REFRESH_FLAG = True
                Clear_Record()
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)
                REFRESH_FLAG = False

            Case "Update"
                Update_Record()

                If ASCMAIN1.CLIENT = "RGI" Then
                    For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                        MessageBox.Show("Credit Invoice Number: " & row.Item("INV_NO"), "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Next
                End If

                Mode_Settings(False)

            Case "Reverse"
                EntryMode = "R"
                Set_Up_Reversal()
                Update_Record()
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Back to Stock Report"
                Back_to_Stock_Report(True)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Price Change").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    'If ScreenMode And EntryMode <> "N" Then
                    '    .Items("Update").Settings.Enabled = not_iScreenMode
                    '    .Items("Cancel").Settings.Enabled = not_iScreenMode
                    'Else
                    '    .Items("Update").Settings.Enabled = iScreenMode
                    '    .Items("Cancel").Settings.Enabled = iScreenMode
                    'End If
                    .Items("Refresh").Settings.Enabled = iScreenMode And ASCMAIN1.CLIENT = "RGI" And InquiryMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    'If ScreenMode And EntryMode <> "V" Then
                    '    .Items("Print").Settings.Enabled = not_iScreenMode
                    '    .Items("Done").Settings.Enabled = not_iScreenMode
                    'Else
                    '    .Items("Print").Settings.Enabled = iScreenMode
                    '    .Items("Done").Settings.Enabled = iScreenMode
                    'End If
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Back to Stock Report").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Back to Stock Report").Visible = ASCMAIN1.CLIENT = "RGI"

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V") And Not InquiryMode _
                        AndAlso rowSOTRTRN1 IsNot Nothing _
                        AndAlso rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO") Is DBNull.Value _
                        AndAlso rowSOTRTRN1.Item("REVERSES_RTRN_NO") Is DBNull.Value

                    .Items("New").Visible = Not InquiryMode
                    .Items("Price Change").Visible = False
                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                    .Items("Print").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                    .Items("Refresh").Visible = iScreenMode And ASCMAIN1.CLIENT = "RGI" And InquiryMode
                    .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                    .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                End With

                .Groups("GL Distribution").Visible = ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                '   .Groups("Show if Entered in").Visible = Not ScreenMode 'And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Totals").Visible = ScreenMode
                Set_Read_Only(grpTotals, ScreenMode And (EntryMode = "V"))

                .Groups("Events").Visible = ScreenMode And (EntryMode <> "N")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        splHeader.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        tab0.Tabs("3PL").Visible = Not InquiryMode
        tab0.Tabs("Invoice History").Visible = Not InquiryMode

        tab0.Visible = Not ScreenMode
        Setup_tab0()

        If ScreenMode Then
            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            If ASCMAIN1.CLIENT = "RGI" Then lblMisc.Text = "Tariff"

            tabDetails.Tabs("Sales History").Visible = (EntryMode = "N")
            If Not REFRESH_FLAG Then tabDetails.SelectedTab = tabDetails.Tabs("Sales History")
            tabDetails.Tabs("GL Distribution").Visible = (EntryMode = "V") And ASCMAIN1.USER_SECURITY_CODEs.Contains("X5")
            tabDetails.Tabs("Scanned Returns").Visible = (EntryMode = "N") And ASCMAIN1.CLIENT = "RGI"

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            'Set_Read_Only(splGL, (EntryMode = "V"))
            Set_Read_Only_for_ctl(chkNoImpact, True)

            If EntryMode = "V" Then
                'Set_Read_Only_for_ctl(Absx1.dteFor("RTRN_DATE"), True)
                'Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), True)
            Else
                Set_Read_Only_for_ctl(Absx1.dteFor("RTRN_DATE"), False)
                Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), LoopReturnID.Length > 0)
            End If


            If EntryMode = "N" Then
                grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY_1").CellActivation = UltraWinGrid.Activation.AllowEdit

                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTRTRN2}
                    With grd.DisplayLayout.Override
                        If is180Customer Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.True
                            .AllowUpdate = DefaultableBoolean.True
                        ElseIf whse_is_a_3PL Then
                            If KEY_3PL_RECORD = "" Then
                                ' we need to allow normal entry if we are asking to use a reason code that does not impact inventory
                                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                                .AllowDelete = DefaultableBoolean.True
                                .AllowUpdate = DefaultableBoolean.True
                            Else
                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                .AllowDelete = DefaultableBoolean.False
                                .AllowUpdate = DefaultableBoolean.True
                            End If
                        ElseIf Whse_Rtn_no <> "" Then
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.True
                        Else
                            If INV_NO_RETURNED = "" Then
                                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            Else
                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            End If
                            .AllowDelete = DefaultableBoolean.True
                            .AllowUpdate = DefaultableBoolean.True
                        End If
                    End With
                Next

                With grdSOTRTRN2.DisplayLayout.Bands(0)
                    If is180Customer Then
                        .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("RTRN_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit

                    ElseIf whse_is_a_3PL And KEY_3PL_RECORD <> "" Then
                        .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                        .Columns("RTRN_QTY").CellActivation = UltraWinGrid.Activation.NoEdit

                    Else
                        .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Columns("RTRN_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                End With


            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTRTRN2, grdSOTRTRN3}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                With grdSOTRTRN2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("COLOR_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("RTRN_QTY").CellAppearance.BackColor = Color.Empty
                End With
            End If
        Else
            Clear_Record()
            Show_Filter(grdSOTRMAFX, True)
            grdSOTRMAFX.DisplayLayout.GroupByBox.Hidden = False

            If tab0SelectedTab.Length > 0 Then
                tab0.SelectedTab = tab0.Tabs(tab0SelectedTab)
            End If
        End If

        If LoopReturnID.Length > 0 Then
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY").CellActivation = Activation.NoEdit
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_PRICE").CellActivation = Activation.NoEdit
        Else
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY").CellActivation = Activation.AllowEdit
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_PRICE").CellActivation = Activation.AllowEdit
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTRTRN0", "SOTRTRN1", "SOTRTRN2", "SOTRTRN1P", "SOTRTRN2P", "SOTRTRN3", "SOTINVHH", "SOTINVHX", "ICTIADJ1", "ICTIADJ2",
             "SOTORDR2180", "SOTORDR2", "ICTWHSE1", "SOTRTRN2_RPT", "WHTMOVE1", "WHTMOVE2", "SOTRMAF1", "SOTRMAF2",
             "SOTRTNL1", "SOTRTNL2", "SOTRTNL3", "SOTRTNL2_REF"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If ASCMAIN1.CLIENT = "NYA" Then
            For Each TABLE_NAME As String In New String() _
                {"EDTSYSIH", "EDT180O1", "EDT180O2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        End If

        EnforceConstraints(True)

        If REFRESH_FLAG Then Exit Sub

        If chkGL.Checked Then
            chkGL.Checked = False
        Else
            Refresh_Documents()
        End If
        Setup_tab0_GL()
        If ASCMAIN1.CLIENT = "VAN" Then
            'Setup_3PL()
        Else
            tab0.Tabs("3PL").Visible = False
        End If

        Whse_Rtn_no = ""
        Fill_Records("WHTWRTN1")
        Sort_grdColumns(grdWHTRTRN1, "WH_RTN_NO")

        INV_NO_RETURNED = ""
        KEY_3PL_RECORD = ""

        txtTrackingNo.Clear()
        txtReturnId.Clear()
        LoopReturnID = String.Empty

        Absx1.txtFor("WHSE_CODE").Clear()
        Absx1.dteFor("RTRN_DATE").Value = Format(Now, "MM/dd/yyyy")
        'Absx1.dteFor("RTRN_DATE").Value = Now.Date
        Absx1.txtFor("RTRN_NO").Clear()

        optGL.Tag = ""
        priceChange = False
        preloadInvoiceDetails = False
        is180Customer = False

        Absx1.txtFor("STYLE_CODE").Clear()
        Absx1.txtFor("SHIP_TO_NAME").Clear()
        RA_NO = String.Empty
        RTRN_NO = String.Empty

        Load_SOTRMAFX()

        Clear_All_Filters(grdSOTRTRN2)

        ' Foreign exchanges display column RTRN_PRICE_CURR
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_PRICE_CURR").Hidden = True
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_PRICE").Hidden = False

        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("LINE_SALES_CURR").Hidden = True
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("LINE_SALES").Hidden = False
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("REFUND_TAX").Hidden = ASCMAIN1.CLIENT <> "VAN"
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("WEB_RETURN_REASON").Hidden = ASCMAIN1.CLIENT <> "VAN"
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY_REFUSED").Hidden = ASCMAIN1.CLIENT <> "VAN"
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("CUST_UPC").Hidden = ASCMAIN1.CLIENT <> "VAN"
        grdSOTINVHX.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE_CURR").Hidden = True
        grdSOTINVHX.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Hidden = False
        grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop

        ' must have ths value set to be able to enter a Handling Charge
        Absx1.numFor("RTRN_HANDLING").Enabled = SO_PARM_MISC_CHG_RTN.Length > 0

        WHSE_LOC_RFB = String.Empty
        WHSE_LOC_DST = String.Empty

        tab0.Tabs("Web Returns").Visible = ASCMAIN1.CLIENT = "VAN"

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        If REFRESH_FLAG Then
        Else
            Save_Header_Fields(UltraGroupBox1)
        End If

        ' Preserve the selected tab
        tab0SelectedTab = tab0.SelectedTab.Key

        Dim rowSOTRMAF1 As DataRow = Nothing
        Dim rowARTCUST1 As DataRow = Nothing
        RTRN_NO = Absx1.txtFor("RTRN_NO").Text

        If EntryMode = "N" Then

            ' Grab RA NO
            rowARTCUST1 = LookUp("ARTCUST1", HFs("CUST_CODE"))
            If RA_NO.Length > 0 Then
                Fill_Records("SOTRMAF1", RA_NO)
                Fill_Records("SOTRMAF2", RA_NO)
                HFs("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & String.Empty
                If dst.Tables("SOTRMAF1").Rows.Count > 0 Then
                    rowSOTRMAF1 = dst.Tables("SOTRMAF1").Rows(0)
                End If

                Fill_Records("SOTRMAFR", RA_NO)
                Dim dvw As DataView = dst.Tables("SOTRMAFR").DefaultView
                dvw.RowFilter = "GUN_STATUS <> 'V'"

                Fill_Records("SOTRMAFRS", RA_NO)
                splSCANS.Panel2Collapsed = True
            End If

            rowSOTRTRN1 = dst.Tables("SOTRTRN1").NewRow
            If ASCMAIN1.CLIENT = "VAN" Then
                rowSOTRTRN1.Item("RTRN_NO") = ASCMAIN1.Next_Control_No("TRAN_NO_C")
            Else
                If InquiryMode Then
                    rowSOTRTRN1.Item("RTRN_NO") = "0000000000"
                Else
                    rowSOTRTRN1.Item("RTRN_NO") = ASCMAIN1.Next_Control_No("SOTRTRN1.RTRN_NO")
                End If
            End If

            rowSOTRTRN1.Item("CURR_CODE") = rowARTCUST1.Item("CURR_CODE") & String.Empty
            If rowSOTRTRN1.Item("CURR_CODE") & String.Empty = String.Empty Then
                rowSOTRTRN1.Item("CURR_CODE") = GL_PARM_CURR_CODE
            End If

            If rowSOTRTRN1.Item("CURR_CODE") = GL_PARM_CURR_CODE Then
                rowSOTRTRN1.Item("CURR_EXCH_RATE") = 1
            Else
                rowSOTRTRN1.Item("CURR_EXCH_RATE") = rowSOTINVH1_Ret.Item("CURR_EXCH_RATE")
                preloadInvoiceDetails = True
            End If

            rowSOTRTRN1.Item("CUST_CODE") = HFs("CUST_CODE")
            rowSOTRTRN1.Item("CUST_NAME") = HFs("CUST_NAME")
            rowSOTRTRN1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            rowSOTRTRN1.Item("RTRN_DATE") = HFs("RTRN_DATE")
            rowSOTRTRN1.Item("RTRN_SOURCE") = "E"
            rowSOTRTRN1.Item("RTRN_STATUS") = "U"
            rowSOTRTRN1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowSOTRTRN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSOTRTRN1.Item("INIT_DATE") = DATETIME_STAMP
            rowSOTRTRN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSOTRTRN1.Item("LAST_DATE") = DATETIME_STAMP
            rowSOTRTRN1.Item("REGISTER_IND") = "0"

            If ASCMAIN1.CLIENT = "RGI" Then
                If dst.Tables("SOTRMAF1").Rows.Count > 0 Then
                    Dim CLAIM_NO As String = dst.Tables("SOTRMAF1").Rows(0).Item("CUST_CLAIM_NO") & String.Empty
                    Dim CUST_CODE As String = dst.Tables("SOTRMAF1").Rows(0).Item("CUST_CODE") & String.Empty

                    Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_TYPE = 'I' AND INV_NO = :PARM1 AND CUST_CODE = :PARM2", "VV", New Object() {CLAIM_NO, CUST_CODE})
                    If rowSOTINVH1 IsNot Nothing Then
                        rowSOTRTRN1.Item("INV_NO_RETURNED") = rowSOTINVH1.Item("INV_NO") & String.Empty
                        rowSOTRTRN1.Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO") & String.Empty
                    End If
                End If
            End If

            If rowSOTRMAF1 IsNot Nothing Then
                rowSOTRTRN1.Item("RA_NO") = rowSOTRMAF1.Item("RA_NO")
                rowSOTRTRN1.Item("CUST_CLAIM_NO") = rowSOTRMAF1.Item("CUST_CLAIM_NO")
                If rowSOTRTRN1.Item("CUST_STORE_NO") & String.Empty = String.Empty Then
                    rowSOTRTRN1.Item("CUST_STORE_NO") = rowSOTRMAF1.Item("CUST_STORE_NO")
                End If
                rowSOTRTRN1.Item("SALES_DIVISION_CODE") = rowSOTRMAF1.Item("SALES_DIVISION_CODE")
                rowSOTRTRN1.Item("SREP_CODE") = rowSOTRMAF1.Item("SREP_CODE")
                rowSOTRTRN1.Item("REASON_CODE") = rowSOTRMAF1.Item("RA_REASON_CODE")
                rowSOTRTRN1.Item("RTRN_NOTE") = rowSOTRMAF1.Item("RA_NOTES")

            ElseIf rowSOTINVH1_Ret IsNot Nothing Then
                rowSOTRTRN1.Item("SREP_CODE") = rowSOTINVH1_Ret.Item("SREP_CODE")
                rowSOTRTRN1.Item("SALES_DIVISION_CODE") = rowSOTINVH1_Ret.Item("SALES_DIVISION_CODE")
                rowSOTRTRN1.Item("CUST_STORE_NO") = rowSOTINVH1_Ret.Item("CUST_STORE_NO")
                rowSOTRTRN1.Item("RTRN_FREIGHT") = rowSOTINVH1_Ret.Item("INV_FREIGHT")
                rowSOTRTRN1.Item("RTRN_STAX") = 0 ' rowSOTINVH1_Ret.Item("INV_STAX")
                rowSOTRTRN1.Item("RTRN_SOURCE_DOC_NO") = rowSOTINVH1_Ret.Item("INV_NO")
                rowSOTRTRN1.Item("INV_NO_RETURNED") = rowSOTINVH1_Ret.Item("INV_NO")
                rowSOTRTRN1.Item("CUST_CLAIM_NO") = (rowSOTINVH1_Ret.Item("ORDR_CUST_PO") & String.Empty).ToString.PadRight(20, " ").Substring(0, 20).Trim
                rowSOTRTRN1.Item("ORDR_NO") = rowSOTINVH1_Ret.Item("ORDR_NO")
                rowSOTRTRN1.Item("CURR_EXCH_RATE") = rowSOTINVH1_Ret.Item("CURR_EXCH_RATE")

                If (ASCMAIN1.CLIENT = "RGI") Then
                    rowSOTRTRN1.Item("RTRN_NOTE") = "Orig Inv No: " & rowSOTINVH1_Ret.Item("INV_NO")
                End If

            End If

            If rowSOTRTRN1.Item("SREP_CODE") & String.Empty = String.Empty Then
                rowSOTRTRN1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
            End If

            If ASCMAIN1.CLIENT = "VAN" Then
            Else
                If dst.Tables("ARTREAS1").Rows.Count = 1 Then
                    rowSOTRTRN1.Item("REASON_CODE") = dst.Tables("ARTREAS1").Rows(0).Item("REASON_CODE")
                End If
            End If

            dst.Tables("SOTRTRN1").Rows.Add(rowSOTRTRN1)

            If is180Customer Then
                ASCMAIN1.sql = "SELECT * FROM SOTORDR2 WHERE ORDR_NO = (SELECT ORDR_NO FROM SOTINVH1 WHERE INV_NO = '" & rowSOTRTRN1.Item("INV_NO_RETURNED") & "')"
                Fill_Records("SOTORDR2", String.Empty, True, ASCMAIN1.sql)

                If dst.Tables("SOTORDR2").Rows.Count > 0 Then
                    rowSOTRTRN1.Item("ORDR_NO") = dst.Tables("SOTORDR2").Rows(0).Item("ORDR_NO")
                End If

                If EntryMode = "N" Then
                    grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_QTY_SHIP > 0", "ORDR_LNO")
                        grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                        With grdSOTRTRN2.ActiveRow
                            .Cells("STYLE_CODE").Value = rowSOTORDR2.Item("STYLE_CODE") & String.Empty
                            .Cells("COLOR_CODE").Value = rowSOTORDR2.Item("COLOR_CODE") & String.Empty
                            .Cells("RTRN_QTY").Value = rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty
                            .Cells("RTRN_QTY_1").Value = rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty
                            .Cells("RTRN_PRICE").Value = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & String.Empty)

                            .Cells("ORDR_NO").Value = rowSOTORDR2.Item("ORDR_NO") & String.Empty
                            .Cells("ORDR_LNO").Value = rowSOTORDR2.Item("ORDR_LNO") & String.Empty

                            .Update()
                        End With
                    Next
                End If

            ElseIf ASCMAIN1.CLIENT = "VAN" AndAlso LoopReturnID.Length > 0 Then

                Dim drSOTRTNL1 As DataRow = LookUp("SOTRTNL1", LoopReturnID)
                rowSOTRTRN1.Item("RTRN_SOURCE_DOC_NO") = LoopReturnID
                rowSOTRTRN1.Item("CUST_CLAIM_NO") = drSOTRTNL1.Item("PROVIDER_ORDER_NUMBER")
                rowSOTRTRN1.Item("RTRN_HANDLING") = Math.Abs(Val(drSOTRTNL1.Item("HANDLING_FEE") & String.Empty)) * -1
                For Each drSOTRTNL2 As DataRow In dst.Tables("SOTRTNL2").Select($"RETURN_ID = '{LoopReturnID}'", "ORDR_LNO")
                    Dim ORDR_NO As String = drSOTRTNL2.Item("ORDR_NO") & String.Empty
                    Dim ORDR_LNO As String = Val(drSOTRTNL2.Item("ORDR_LNO") & String.Empty)
                    Dim drSOTORDR2 As DataRow = LookUp("SOTORDR2", {ORDR_NO, ORDR_LNO})
                    If drSOTORDR2 Is Nothing Then
                        MessageBox.Show($"Cannot locate detail line for Sales Order No {ORDR_NO}, Line No {ORDR_LNO}. Line skipped", "Load Web Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Continue For
                    End If

                    grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                    With grdSOTRTRN2.ActiveRow
                        .Cells("STYLE_CODE").Value = drSOTORDR2.Item("STYLE_CODE") & String.Empty
                        .Cells("COLOR_CODE").Value = drSOTORDR2.Item("COLOR_CODE") & String.Empty
                        .Cells("RTRN_QTY").Value = 1
                        .Cells("RTRN_QTY_1").Value = 1
                        .Cells("RTRN_PRICE").Value = Val(drSOTRTNL2.Item("PRICE") & String.Empty) - Val(drSOTRTNL2.Item("DISCOUNT") & String.Empty)
                        .Cells("RTRN_PRICE_CURR").Value = Val(drSOTRTNL2.Item("PRICE") & String.Empty) - Val(drSOTRTNL2.Item("DISCOUNT") & String.Empty)

                        ' 02/02/2026 - As per Walter do not refund Tax on Exchanges
                        If (drSOTRTNL1.Item("OUTCOME") & String.Empty).ToString.ToUpper <> "EXCHANGE" Then
                            .Cells("REFUND_TAX").Value = Val(drSOTRTNL2.Item("TAX") & String.Empty)
                        Else
                            .Cells("REFUND_TAX").Value = 0
                        End If

                        .Cells("LINE_ITEM_ID").Value = drSOTRTNL2.Item("LINE_ITEM_ID") & String.Empty
                        .Cells("CUST_UPC").Value = drSOTORDR2.Item("CUST_UPC") & String.Empty

                        .Cells("ORDR_NO").Value = drSOTRTNL2.Item("ORDR_NO") & String.Empty
                        .Cells("ORDR_LNO").Value = drSOTRTNL2.Item("ORDR_LNO") & String.Empty

                        If drSOTRTNL2.Item("RETURN_REASON") & String.Empty <> String.Empty Then
                            .Cells("WEB_RETURN_REASON").Value = drSOTRTNL2.Item("RETURN_REASON") & String.Empty
                        Else
                            .Cells("WEB_RETURN_REASON").Value = drSOTRTNL2.Item("PARENT_RETURN_REASON") & String.Empty
                        End If

                        .Update()
                    End With
                Next
                grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = AllowAddNew.No

            ElseIf tblSOTINVH2_Ret IsNot Nothing AndAlso preloadInvoiceDetails Then
                ' If new Load from invoice then load the items
                grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                For Each rowSOTINVH2 As DataRow In tblSOTINVH2_Ret.Select("ORDR_QTY_SHIP > 0", IIf(priceChange, "STYLE_CODE,COLOR_CODE", "INV_LNO"))
                    grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                    With grdSOTRTRN2.ActiveRow
                        .Cells("STYLE_CODE").Value = rowSOTINVH2.Item("STYLE_CODE") & String.Empty
                        .Cells("COLOR_CODE").Value = rowSOTINVH2.Item("COLOR_CODE") & String.Empty
                        .Cells("RTRN_QTY").Value = rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty
                        .Cells("RTRN_QTY_1").Value = rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty
                        .Cells("RTRN_PRICE").Value = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE") & String.Empty)
                        .Cells("RTRN_PRICE_CURR").Value = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") & String.Empty)

                        .Cells("ORDR_NO").Value = rowSOTINVH2.Item("ORDR_NO") & String.Empty
                        '.Cells("ORDR_LNO").Value = rowSOTINVH2.Item("ORDR_LNO") & String.Empty

                        .Update()
                    End With

                    If priceChange Then
                        grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                        With grdSOTRTRN2.ActiveRow
                            .Cells("STYLE_CODE").Value = rowSOTINVH2.Item("STYLE_CODE") & String.Empty
                            .Cells("COLOR_CODE").Value = rowSOTINVH2.Item("COLOR_CODE") & String.Empty
                            .Cells("RTRN_QTY").Value = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty) * -1
                            .Cells("RTRN_QTY_1").Value = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty) * -1
                            '.Cells("RTRN_PRICE").Value = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE") & String.Empty)
                            .Update()
                        End With
                    End If
                Next
            ElseIf dst.Tables("SOTRMAF2").Rows.Count > 0 Then

                ' Add any extra scans from SOTRMAFR
                For Each row As DataRow In dst.Tables("SOTRMAFR").Select("RA_QTY_USED < RA_RTN_QTY")
                    Dim STYLE_CODE As String = row.Item("STYLE_CODE") & String.Empty
                    Dim COLOR_CODE As String = row.Item("COLOR_CODE") & String.Empty

                    If dst.Tables("SOTRMAF2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length = 0 Then
                        Dim RA_QTY_USED As Int32 = Val(dst.Tables("SOTRMAFR").Compute("SUM(RA_QTY_USED)", "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)
                        Dim RA_RTN_QTY As Int32 = Val(dst.Tables("SOTRMAFR").Compute("SUM(RA_RTN_QTY)", "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)
                        Dim RA_QTY As Int32 = RA_RTN_QTY - RA_QTY_USED

                        If RA_QTY > 0 Then
                            Dim rowSOTRMAF2 As DataRow = dst.Tables("SOTRMAF2").NewRow
                            rowSOTRMAF2.Item("RA_NO") = RA_NO
                            rowSOTRMAF2.Item("RA_LNO") = Val(dst.Tables("SOTRMAF2").Compute("MAX(RA_LNO)", "") & String.Empty) + 1
                            rowSOTRMAF2.Item("STYLE_CODE") = STYLE_CODE
                            rowSOTRMAF2.Item("COLOR_CODE") = COLOR_CODE
                            rowSOTRMAF2.Item("RA_QTY") = RA_QTY
                            rowSOTRMAF2.Item("RA_QTY_OPEN") = RA_QTY
                            rowSOTRMAF2.Item("RA_QTY_USED") = 0
                            rowSOTRMAF2.Item("RA_QTY_CANC") = 0
                            rowSOTRMAF2.Item("IMPORTED") = "1"
                            dst.Tables("SOTRMAF2").Rows.Add(rowSOTRMAF2)
                        End If

                    End If
                Next

                For Each rowSOTRMAF2 As DataRow In dst.Tables("SOTRMAF2").Select("RA_QTY_OPEN > 0", "RA_LNO")
                    grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                    grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                    With grdSOTRTRN2.ActiveRow
                        .Cells("STYLE_CODE").Value = rowSOTRMAF2.Item("STYLE_CODE") & String.Empty
                        .Cells("COLOR_CODE").Value = rowSOTRMAF2.Item("COLOR_CODE") & String.Empty
                        .Cells("RTRN_QTY").Value = rowSOTRMAF2.Item("RA_QTY_OPEN") & String.Empty
                        .Cells("IMPORTED").Value = IIf(rowSOTRMAF2.Item("IMPORTED") & String.Empty = "", "0", "1")

                        ASCMAIN1.sql = "STYLE_CODE = '" & rowSOTRMAF2.Item("STYLE_CODE") & "' AND COLOR_CODE = '" & rowSOTRMAF2.Item("COLOR_CODE") & "' AND ISNULL(RA_QTY_USED, 0) < RA_RTN_QTY AND GUN_STATUS = 'F'"
                        If dst.Tables("SOTRMAFR").Select(ASCMAIN1.sql).Length > 0 Then
                            For Each rowSOTRMAFR As DataRow In dst.Tables("SOTRMAFR").Select(ASCMAIN1.sql)
                                Dim notUsed As Int16 = Val(rowSOTRMAFR.Item("RA_RTN_QTY") & String.Empty) - Val(rowSOTRMAFR.Item("RA_QTY_USED") & String.Empty)
                                If notUsed < 0 Then notUsed = 0
                                'rowSOTRMAFR.Item("RA_QTY_USED") = rowSOTRMAFR.Item("RA_RTN_QTY")
                                Select Case rowSOTRMAFR.Item("RA_RTN_STATUS") & String.Empty
                                    Case "1" ' Stock
                                        .Cells("RTRN_QTY_1").Value = Val(.Cells("RTRN_QTY_1").Value & String.Empty) + notUsed
                                    Case "2" ' refurbish
                                        .Cells("RTRN_QTY_2").Value = Val(.Cells("RTRN_QTY_2").Value & String.Empty) + notUsed
                                    Case "3" ' Destroy
                                        .Cells("RTRN_QTY_3").Value = Val(.Cells("RTRN_QTY_3").Value & String.Empty) + notUsed
                                    Case Else
                                        .Cells("RTRN_QTY_1").Value = Val(.Cells("RTRN_QTY_1").Value & String.Empty) + notUsed
                                End Select
                            Next
                        ElseIf dst.Tables("SOTRMAFR").Rows.Count = 0 Then
                            Select Case rowSOTRMAF1.Item("RA_REASON_CODE") & String.Empty
                                Case "D" ' Damaged
                                    .Cells("RTRN_QTY_3").Value = rowSOTRMAF2.Item("RA_QTY_OPEN") & String.Empty
                                Case "X" ' Destroyed
                                    .Cells("RTRN_QTY_3").Value = rowSOTRMAF2.Item("RA_QTY_OPEN") & String.Empty
                                Case "O" ' Overstock
                                    .Cells("RTRN_QTY_1").Value = rowSOTRMAF2.Item("RA_QTY_OPEN") & String.Empty
                                Case "Z" ' Other
                                    .Cells("RTRN_QTY_1").Value = rowSOTRMAF2.Item("RA_QTY_OPEN") & String.Empty
                            End Select
                        End If

                        .Cells("RTRN_PRICE").Value = Val(rowSOTRMAF2.Item("RA_NET_PRICE") & String.Empty)

                        .Update()
                    End With
                Next
            End If
        Else
            Fill_Record("SOTRTRN1", Absx1.txtFor("RTRN_NO").Text)
            Fill_Record("SOTRTRN2", Absx1.txtFor("RTRN_NO").Text)
            dst.AcceptChanges()

            If LoopReturnID.Length > 0 Then
                Fill_Record("SOTRTNL2_REF", Absx1.txtFor("RTRN_NO").Text)
                Dim tblSOTRTNL2 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTRTNL2 WHERE RETURN_ID = :PARM1", "SOTRTNL2", "V", {LoopReturnID})
                For Each drSOTRTNL2 As DataRow In tblSOTRTNL2.Select("")
                    Dim ORDR_NO As String = drSOTRTNL2.Item("ORDR_NO") & String.Empty
                    Dim ORDR_LNO As String = Val(drSOTRTNL2.Item("ORDR_LNO") & String.Empty)

                    Dim drSOTORDR2 As DataRow = LookUp("SOTORDR2", {ORDR_NO, ORDR_LNO})

                    For Each drSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select($"ORDR_NO = '{ORDR_NO}' and ORDR_LNO = {ORDR_LNO}")
                        drSOTRTRN2.Item("REFUND_TAX") = Val(drSOTRTNL2.Item("REFUND_TAX") & String.Empty)
                        drSOTRTRN2.Item("LINE_ITEM_ID") = drSOTRTNL2.Item("LINE_ITEM_ID") & String.Empty

                        If drSOTORDR2 IsNot Nothing Then
                            drSOTRTRN2.Item("CUST_UPC") = drSOTORDR2.Item("CUST_UPC") & String.Empty
                        End If

                        If drSOTRTNL2.Item("RETURN_REASON") & String.Empty <> String.Empty Then
                            drSOTRTRN2.Item("WEB_RETURN_REASON") = drSOTRTNL2.Item("RETURN_REASON") & String.Empty
                        Else
                            drSOTRTRN2.Item("WEB_RETURN_REASON") = drSOTRTNL2.Item("PARENT_RETURN_REASON") & String.Empty
                        End If

                        ' Refused Qty is not stored in SOTRTRN2
                        Dim RTRN_QTY_LN2 As Int32 = Val(drSOTRTNL2.Item("RTRN_QTY") & "")
                        Dim RTRN_QTY As Int32 = Val(drSOTRTRN2.Item("RTRN_QTY") & "")
                        Dim RTRN_QTY_TOTAL As Int32 = Val(drSOTRTRN2.Item("RTRN_QTY_TOTAL") & "")

                        If RTRN_QTY_LN2 <> RTRN_QTY Then
                            If RTRN_QTY > RTRN_QTY_TOTAL Then
                                drSOTRTRN2.Item("RTRN_QTY_REFUSED") = RTRN_QTY - RTRN_QTY_TOTAL
                            End If
                        End If
                    Next
                Next
                tblSOTRTNL2.Rows.Clear()
            End If

            ASCMAIN1.sql = $"Select * from SOTINVHM where INV_TYPE = 'C' and INV_NO = '{rowSOTRTRN1.Item("INV_NO")}' and MISC_CHARGE_TYPE = 'T'"
            For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Rows
                Dim rowSOTRTRN2 As DataRow = dst.Tables("SOTRTRN2").Select($"RTRN_LNO = '{row("INV_LNO")}'").First
                rowSOTRTRN2("SURCHARGE_PERC") = row("SURCHARGE_PERC")
                rowSOTRTRN2("LINE_TARIFF") = Abs(row("INV_MISC_CHG"))
                rowSOTRTRN2("MISC_CHG_CODE") = row("MISC_CHG_CODE")
                rowSOTRTRN2("COUNTRY_CODE") = row("COUNTRY_CODE")
            Next

            dst.Tables("SOTRTRN0").Rows.Add(New String() {"Entered", Format(rowSOTRTRN1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
            dst.Tables("SOTRTRN0").Rows.Add(New String() {"By", rowSOTRTRN1.Item("INIT_OPER")})
            dst.Tables("SOTRTRN0").Rows.Add(New String() {"Source", rowSOTRTRN1.Item("RTRN_SOURCE")})

            If rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO") & "" <> "" Then
                Dim row As DataRow = LookUp("SOTRTRN1", rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO"))
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"Reversed", Format(row.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"By", row.Item("INIT_OPER")})
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"using", rowSOTRTRN1.Item("REVERSED_BY_RTRN_NO")})
            ElseIf rowSOTRTRN1.Item("REVERSES_RTRN_NO") & "" <> "" Then
                dst.Tables("SOTRTRN0").Rows.Add(New String() {"Reverses", rowSOTRTRN1.Item("REVERSES_RTRN_NO")})
            End If

            Fill_Records("SOTRMAFR", RA_NO)

        End If

        rowICTWHSE1 = LookUp("ICTWHSE1", rowSOTRTRN1.Item("WHSE_CODE"))
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        whse_is_a_3PL = (rowICTWHSE1.Item("LP_CODE") & "" <> "")

        If EntryMode = "N" Or EntryMode = "E" Then
            If whse_is_a_3PL Then
                If ASCMAIN1.CLIENT = "NYA" Then
                    whse_is_a_3PL = False
                Else
                    If MsgBox("Do you want to bypass 3PL Integration for this Entry?",
                        MsgBoxStyle.YesNo,
                        "Option to enter Returns without 3PL Integration") = MsgBoxResult.Yes Then
                        whse_is_a_3PL = False
                    End If
                End If
            End If
        End If
        ' whse_is_a_3PL = True
        'With grdSOTRTRN2.DisplayLayout.Bands(0)
        '    .Columns("BAR_CODE").Hidden = True ' Not location_support
        '    .Columns("LOCATION_CODE").Hidden = Not location_support
        'End With

        If EntryMode = "N" And whse_is_a_3PL Then

            ASCMAIN1.sql = "Select RCPTHDR.TRANS_SEQ, RCPTHDR.ARRDTE, RCPTHDR.PO_SHIPMENT_NO, RCPTHDR.CONTAINER_NO " _
                & ", RCPTDTL.ITEM_CODE, RCPTDTL.RCVQTY" _
                & " from ADS.RCPTHDR@ADSIIS,ADS.RCPTDTL@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
                & " and RCPTHDR.INVTYP = 'R'" _
                & "AND RCPTHDR.TRANS_SEQ = RCPTDTL.TRANS_SEQ " _
                & "AND RCPTHDR.LP_CODE = RCPTDTL.LP_CODE " _
                & "AND RCPTHDR.WHSE_CODE = RCPTDTL.WHSE_CODE " _
                & "AND RCPTHDR.WHSE_CODE = '" & rowICTWHSE1.Item("WHSE_CODE") & "'"
        End If

        If tblSOTINVH2_Ret IsNot Nothing Then
            If Not (EntryMode = "N" AndAlso tblSOTINVH2_Ret.Rows.Count > 0) Then
                Fill_Records("SOTRTRN2", Absx1.txtFor("RTRN_NO").Text)
            End If
        End If


        Sort_grdColumns(grdSOTRTRN2, "RTRN_LNO")

        Fill_Records("SOTRTRN3", Absx1.txtFor("RTRN_NO").Text)

        If INV_NO_RETURNED <> "" Then
            rowSOTRTRN1.Item("INV_NO_RETURNED") = INV_NO_RETURNED
            ASCMAIN1.sql = "Select * from SOTINVH2 where INV_TYPE = 'I' and INV_NO = '" & INV_NO_RETURNED & "'"
            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "INV_LNO")
                grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                With grdSOTRTRN2.ActiveRow
                    .Cells("STYLE_CODE").Value = row.Item("STYLE_CODE")
                    .Cells("COLOR_CODE").Value = row.Item("COLOR_CODE")
                    .Cells("RTRN_QTY").Value = row.Item("ORDR_QTY_SHIP")
                    .Cells("RTRN_QTY_1").Value = row.Item("ORDR_QTY_SHIP")
                    .Cells("RTRN_PRICE").Value = row.Item("ORDR_UNIT_PRICE")
                    .Update()
                End With
            Next
            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Sort_grdColumns(grdSOTRTRN2, "RTRN_LNO")
            Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"I", INV_NO_RETURNED})
            'Absx1.txtFor("CUST_CLAIM_NO").Text = grdSOTINVHH.ActiveRow.Cells("ORDR_CUST_PO").Value & String.Empty
            'Absx1.txtFor("CUST_STORE_NO").Text = grdSOTINVHH.ActiveRow.Cells("CUST_STORE_NO").Value & String.Empty
            'Absx1.numFor("RTRN_FREIGHT").Value = grdSOTINVHH.ActiveRow.Cells("INV_FREIGHT").Value
            'Absx1.txtFor("SREP_CODE").Text = rowSOTINVH1.Item("SREP_CODE") & ""
            rowSOTRTRN1.Item("SREP_CODE") = rowSOTINVH1.Item("SREP_CODE")
            rowSOTRTRN1.Item("SALES_DIVISION_CODE") = rowSOTINVH1.Item("SALES_DIVISION_CODE")
            Absx1.txtFor("CUST_CLAIM_NO").Text = rowSOTINVH1.Item("ORDR_CUST_PO") & ""
            Absx1.txtFor("CUST_STORE_NO").Text = rowSOTINVH1.Item("CUST_STORE_NO") & ""
            Absx1.numFor("RTRN_FREIGHT").Value = Val(rowSOTINVH1.Item("INV_FREIGHT") & "")
        End If
        If rowSOTRTRN1.Item("INV_NO_RETURNED") & "" <> "" Then
            Fill_Records("SOTINVHMR", New String() {rowSOTRTRN1.Item("INV_NO_RETURNED")})
            If preloadInvoiceDetails Then
                For Each row As DataRow In dst.Tables("SOTINVHMR").Rows
                    For Each row2 As DataRow In dst.Tables("SOTRTRN2").Select($"STYLE_CODE = '{row("STYLE_CODE")}'")
                        If IsDBNull(row2("SURCHARGE_PERC")) Then
                            row2("SURCHARGE_PERC") = row("SURCHARGE_PERC")
                            row2("LINE_TARIFF") = Round(row2("RTRN_PRICE") * (row("SURCHARGE_PERC") / 100), 2) * row2("RTRN_QTY")
                            row2("MISC_CHG_CODE") = row("MISC_CHG_CODE")
                            row2("COUNTRY_CODE") = row("COUNTRY_CODE")
                        End If
                    Next
                Next
            End If
        End If

        rowARTCUST1 = LookUp("ARTCUST1", HFs("CUST_CODE"))

        If HFs("RA_NO") & String.Empty <> String.Empty Then
            Fill_Records("SOTRTRN1P", String.Empty, True, "SELECT * FROM SOTRTRN1 WHERE RA_NO = '" & HFs("RA_NO") & "'")
            Fill_Records("SOTRTRN2P", String.Empty, True, "SELECT * FROM SOTRTRN2 WHERE RTRN_NO IN (SELECT RTRN_NO FROM SOTRTRN1 WHERE RA_NO = '" & HFs("RA_NO") & "')")
            grdSOTRTRN1P.DisplayLayout.Bands(0).PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
            grdSOTRTRN1P.DisplayLayout.Bands(1).PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        End If

        grdSOTRMAFR.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        If rowSOTRTRN1.Item("CURR_CODE") <> GL_PARM_CURR_CODE Then
            ' Foreign exchanges display column RTRN_PRICE_CURR
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_PRICE_CURR").Hidden = False
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_PRICE").Hidden = True

            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("LINE_SALES_CURR").Hidden = False
            grdSOTRTRN2.DisplayLayout.Bands(0).Columns("LINE_SALES").Hidden = True

            grdSOTINVHX.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE_CURR").Hidden = False
            grdSOTINVHX.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").Hidden = True

        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()

            If rowSOTRTRN1.Item("CURR_CODE") = GL_PARM_CURR_CODE Then
                rowSOTRTRN1.Item("RTRN_SALES_CURR") = rowSOTRTRN1.Item("RTRN_SALES")
                rowSOTRTRN1.Item("RTRN_STAX_CURR") = rowSOTRTRN1.Item("RTRN_STAX")
                rowSOTRTRN1.Item("RTRN_FREIGHT_CURR") = rowSOTRTRN1.Item("RTRN_FREIGHT")
                rowSOTRTRN1.Item("RTRN_HANDLING_CURR") = rowSOTRTRN1.Item("RTRN_HANDLING")
                rowSOTRTRN1.Item("RTRN_AMOUNT_CURR") = rowSOTRTRN1.Item("RTRN_AMOUNT")

                For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("")
                    rowSOTRTRN2.Item("RTRN_PRICE_CURR") = rowSOTRTRN2.Item("RTRN_PRICE")
                Next
            Else
                Dim CURR_EXCH_RATE As Decimal = Val(rowSOTRTRN1.Item("CURR_EXCH_RATE") & String.Empty)

                rowSOTRTRN1.Item("RTRN_SALES_CURR") = Val(rowSOTRTRN1.Item("RTRN_SALES") & String.Empty)
                rowSOTRTRN1.Item("RTRN_SALES") = Val(rowSOTRTRN1.Item("RTRN_SALES_CURR") & String.Empty) * CURR_EXCH_RATE

                rowSOTRTRN1.Item("RTRN_STAX_CURR") = Val(rowSOTRTRN1.Item("RTRN_STAX") & String.Empty)
                rowSOTRTRN1.Item("RTRN_STAX") = Val(rowSOTRTRN1.Item("RTRN_STAX_CURR") & String.Empty) * CURR_EXCH_RATE

                rowSOTRTRN1.Item("RTRN_FREIGHT_CURR") = Val(rowSOTRTRN1.Item("RTRN_FREIGHT") & String.Empty)
                rowSOTRTRN1.Item("RTRN_FREIGHT") = Val(rowSOTRTRN1.Item("RTRN_FREIGHT_CURR") & String.Empty) * CURR_EXCH_RATE

                rowSOTRTRN1.Item("RTRN_HANDLING_CURR") = Val(rowSOTRTRN1.Item("RTRN_HANDLING") & String.Empty)
                rowSOTRTRN1.Item("RTRN_HANDLING") = Val(rowSOTRTRN1.Item("RTRN_HANDLING_CURR") & String.Empty) * CURR_EXCH_RATE

                rowSOTRTRN1.Item("RTRN_AMOUNT_CURR") _
                    = Val(rowSOTRTRN1.Item("RTRN_SALES_CURR") & "") _
                    + Val(rowSOTRTRN1.Item("RTRN_STAX_CURR") & "") _
                    + Val(rowSOTRTRN1.Item("RTRN_FREIGHT_CURR") & "") _
                    + Val(rowSOTRTRN1.Item("RTRN_HANDLING_CURR") & "")

                For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("")
                    rowSOTRTRN2.Item("RTRN_PRICE") = Val(rowSOTRTRN2.Item("RTRN_PRICE_CURR") & String.Empty) * CURR_EXCH_RATE
                Next
            End If

            rowSOTRTRN1.Item("RTRN_AMOUNT") _
                = Val(rowSOTRTRN1.Item("RTRN_SALES") & "") _
                + Val(rowSOTRTRN1.Item("RTRN_STAX") & "") _
                + Val(rowSOTRTRN1.Item("RTRN_FREIGHT") & "") _
                + Val(rowSOTRTRN1.Item("RTRN_HANDLING") & "")

            For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("")
                rowSOTRTRN2.Item("OPS_YYYYPP") = rowSOTRTRN1.Item("OPS_YYYYPP")
            Next

            If rowSOTRTRN1.Item("SALES_DIVISION_CODE") & "" = "" Then
                rowSOTRTRN1.Item("SALES_DIVISION_CODE") = dst.Tables("SOTRTRN2").Compute("MIN(SALES_DIVISION_CODE)", "")
            End If

            '************** UPDTATE TO SOTINVH1 ******************
            Dim INV_NO As String = ICCMAIN1.Update_Return(Me)
            Dim INV_NO_RETURNED As String = rowSOTRTRN1.Item("INV_NO_RETURNED") & ""

            If EntryMode = "R" Then
                Dim row As DataRow = dst.Tables("SOTRTRN1").Rows(0)
                Dim REVERSED_BY_RTRN_NO As String = row.Item("RTRN_NO")
                Dim REVERSES_RTRN_NO As String = row.Item("REVERSES_RTRN_NO")
                KEY_3PL_RECORD = row.Item("KEY_3PL_RECORD") & ""
                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select * from SOTRTRN1" _
                    & " where RTRN_NO = '" & REVERSED_BY_RTRN_NO & "';" _
                    & " Begin For R1 in C1 Loop" _
                    & "  Update SOTRTRN1 Set REVERSED_BY_RTRN_NO = R1.RTRN_NO" _
                    & ", LAST_DATE = R1.INIT_DATE" _
                    & ", LAST_OPER = R1.INIT_OPER" _
                    & " where RTRN_NO = '" & REVERSES_RTRN_NO & "';" _
                    & " End Loop; End; " _
                    & "End;"
                ASCDATA1.ExecuteSQL()
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                If INV_NO_RETURNED <> "" Then
                    ASCMAIN1.sql = "" _
                        & "Begin" & vbCrLf _
                        & " Declare Cursor C1 is" & vbCrLf _
                        & "  Select * from SOTINVH2" & vbCrLf _
                        & "   where INV_TYPE = 'C' and INV_NO = '" & INV_NO & "'" & vbCrLf _
                        & "   and ORDR_PRICE_SOURCE is Null" & vbCrLf _
                        & "   for Update;" & vbCrLf _
                        & " GOT_ONE VARCHAR2(1);" & vbCrLf _
                        & " Begin" & vbCrLf _
                        & "  For R1 in C1 Loop" & vbCrLf _
                        & "   Begin" & vbCrLf _
                        & "    Declare Cursor C2 is" & vbCrLf _
                        & "     Select * from SOTINVH2 where INV_NO = '" & INV_NO_RETURNED & "'" & vbCrLf _
                        & "       and INV_LNO = R1.INV_LNO and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
                        & "    Begin" & vbCrLf _
                        & "     GOT_ONE := '0';" & vbCrLf _
                        & "     For R2 in C2 Loop" & vbCrLf _
                        & "      Update SOTINVH2 Set COMM_RATE = R2.COMM_RATE, ORDR_PRICE_SOURCE = R2.ORDR_PRICE_SOURCE" & vbCrLf _
                        & "       where Current of C1;" & vbCrLf _
                        & "      GOT_ONE := '1';" & vbCrLf _
                        & "     End Loop;" & vbCrLf _
                        & "    End;" & vbCrLf _
                        & "   End;" & vbCrLf _
                        & "   If GOT_ONE = '0' Then " & vbCrLf _
                        & "    Begin" & vbCrLf _
                        & "     Declare Cursor C2 is" & vbCrLf _
                        & "      Select * from SOTINVH2 where INV_NO = '" & INV_NO_RETURNED & "'" & vbCrLf _
                        & "        and STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
                        & "     Begin" & vbCrLf _
                        & "      For R2 in C2 Loop" & vbCrLf _
                        & "       Update SOTINVH2 Set COMM_RATE = R2.COMM_RATE, ORDR_PRICE_SOURCE = R2.ORDR_PRICE_SOURCE" & vbCrLf _
                        & "        where Current of C1;" & vbCrLf _
                        & "      End Loop;" & vbCrLf _
                        & "     End;" & vbCrLf _
                        & "    End;" & vbCrLf _
                        & "   End If;" & vbCrLf _
                        & "  End Loop;" & vbCrLf _
                        & " End;" & vbCrLf _
                        & "End;"
                    ASCDATA1.ExecuteSQL()
                End If
            End If

            Dim usedScans As Boolean = False

            If dst.Tables("SOTRMAFR").Select("RA_QTY_USED < RA_RTN_QTY").Length > 0 Then
                For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("")
                    Dim STYLE_CODE As String = rowSOTRTRN2.Item("STYLE_CODE") & String.Empty
                    Dim COLOR_CODE As String = rowSOTRTRN2.Item("COLOR_CODE") & String.Empty
                    Dim RTRN_QTY As Int32 = Val(rowSOTRTRN2.Item("RTRN_QTY") & String.Empty)
                    usedScans = True

                    For Each rowSOTRMAFR As DataRow In dst.Tables("SOTRMAFR").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' AND RA_QTY_USED < RA_RTN_QTY")
                        Dim qtyAvail As Int32 = Val(rowSOTRMAFR.Item("RA_RTN_QTY") & String.Empty) = Val(rowSOTRMAFR.Item("RA_QTY_USED") & String.Empty)

                        If RTRN_QTY <= qtyAvail Then
                            rowSOTRMAFR.Item("RA_QTY_USED") = Val(rowSOTRMAFR.Item("RA_QTY_USED") & String.Empty) + qtyAvail
                            qtyAvail = 0
                        Else
                            RTRN_QTY -= qtyAvail
                            rowSOTRMAFR.Item("RA_QTY_USED") = Val(rowSOTRMAFR.Item("RA_RTN_QTY") & String.Empty)
                        End If
                    Next
                Next
                Update_Record_TDA("SOTRMAFR")
            End If


            'If ASCMAIN1.client  = "VAN" Then
            '    If whse_is_a_3PL Then
            '        If EntryMode = "R" Then
            '            If KEY_3PL_RECORD <> "" Then
            '                ASCDATA1.ExecuteSQL("Update ADS.RCPTHDR@ADSIIS Set STATUS = '0' where TRANS_SEQ = '" & KEY_3PL_RECORD & "'")
            '            End If
            '        Else
            '            ASCDATA1.ExecuteSQL("Update ADS.RCPTHDR@ADSIIS Set STATUS = '1' where TRANS_SEQ = '" & KEY_3PL_RECORD & "'")
            '        End If
            '    End If
            'End If

            If location_support Then
                If Whse_Rtn_no <> "" Then
                    ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                    New Object() {"Z", Whse_Rtn_no, ASCMAIN1.SESSION_NO},
                    New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

                    ASCMAIN1.sql = "Update WHTWRTN1 Set WH_RTN_STATUS = 'F' Where WH_RTN_NO = '" & Whse_Rtn_no & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                Else

                    Dim LOCS As New Dictionary(Of String, String)
                    LOCS.Add("RTRN_QTY_1", "WHSE_LOC_RTN")
                    LOCS.Add("RTRN_QTY_2", "WHSE_LOC_RFB")
                    LOCS.Add("RTRN_QTY_3", "WHSE_LOC_DST")

                    For Each Q As String In LOCS.Keys
                        ASCMAIN1.sql = "Select SOTRTRN2.RTRN_NO WHSE_TRAN_NO, SOTRTRN2.RTRN_LNO WHSE_TRAN_LNO" _
                            & ", 'C' WHSE_TRAN_TYPE, SOTRTRN1.WHSE_CODE" _
                            & ", ICTWHSE1." & LOCS(Q) & " LOCATION_CODE, SOTRTRN2.STYLE_CODE, SOTRTRN2.COLOR_CODE" _
                            & ", SOTRTRN2." & Q & " WHSE_TRAN_QTY" _
                            & " from SOTRTRN1,SOTRTRN2,ICTWHSE1" _
                            & " where SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO" _
                            & "   and ICTWHSE1.WHSE_CODE = SOTRTRN1.WHSE_CODE" _
                            & "   and SOTRTRN2." & Q & " <> 0" _
                            & "   and SOTRTRN2.RTRN_NO = '" & rowSOTRTRN1.Item("RTRN_NO") & "'"
                        WHCMAIN1.Update_WHTLOCBX(Me)
                    Next
                End If

            End If

            ' Get Return Location.
            If ASCMAIN1.CLIENT = "RGI" AndAlso location_support AndAlso Not usedScans Then
                MoveInventoryToLocation()
            End If

            If ASCMAIN1.CLIENT = "NYA" Then
                CreateEDI180()
                Update_Record_TDA("EDT180O1")
                Update_Record_TDA("EDT180O2")
                Update_Record_TDA("EDTSYSIH")
            End If

            If EntryMode = "N" And RA_NO.Length > 0 Then
                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is SELECT SOTRTRN2.STYLE_CODE, SOTRTRN2.COLOR_CODE, SUM(SOTRTRN2.RTRN_QTY) RTRN_QTY" _
                    & " FROM SOTRTRN1, SOTRTRN2" _
                    & " WHERE SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO" _
                    & " AND SOTRTRN1.RA_NO = '" & RA_NO & "' " _
                    & " GROUP BY SOTRTRN2.STYLE_CODE, SOTRTRN2.COLOR_CODE;" _
                    & " Begin For R1 in C1 Loop" _
                    & "     Update SOTRMAF2 " _
                    & "     SET RA_QTY_OPEN = GREATEST(0, RA_QTY - R1.RTRN_QTY)" _
                    & "     , RA_QTY_USED = DECODE(RA_QTY, 0, R1.RTRN_QTY, LEAST(RA_QTY, R1.RTRN_QTY))" _
                    & "     where RA_NO = '" & RA_NO & "' AND STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE;" _
                    & " Update SOTRMAF1 SET LAST_OPER = '" & ASCMAIN1.USER_ID & "', LAST_DATE = SYSDATE WHERE RA_NO = '" & RA_NO & "';" _
                    & " End Loop; End; " _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                Dim RA_QTY_OPEN As Int32 = Val(ASCDATA1.GetDataValue("SELECT SUM(RA_QTY_OPEN) FROM SOTRMAF2 WHERE RA_NO = :PARM1", "V", New Object() {RA_NO}) & String.Empty)
                If RA_QTY_OPEN = 0 Then
                    ASCDATA1.ExecuteSQL("UPDATE SOTRMAF1 SET RA_STATUS = 'F' WHERE RA_NO = '" & RA_NO & "'")
                End If

                If dst.Tables("SOTRMAFR").Rows.Count > 0 Then
                    ' Grab a fresh copy
                    Fill_Records("SOTRMAFR", RA_NO)
                    For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("")
                        Dim STYLE_CODE As String = rowSOTRTRN2.Item("STYLE_CODE") & String.Empty
                        Dim COLOR_CODE As String = rowSOTRTRN2.Item("COLOR_CODE") & String.Empty
                        Dim RTRN_QTY As Int32 = Val(rowSOTRTRN2.Item("RTRN_QTY") & String.Empty)

                        ASCMAIN1.sql = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' AND ISNULL(RA_QTY_USED, 0) < RA_RTN_QTY"
                        For Each rowSOTRMAFR As DataRow In dst.Tables("SOTRMAFR").Select(ASCMAIN1.sql)
                            Dim notUsed As Int16 = Val(rowSOTRMAFR.Item("RA_RTN_QTY") & String.Empty) - Val(rowSOTRMAFR.Item("RA_QTY_USED") & String.Empty)
                            If notUsed > RTRN_QTY Then
                                rowSOTRMAFR.Item("RA_QTY_USED") = Val(rowSOTRMAFR.Item("RA_QTY_USED") & String.Empty) + RTRN_QTY
                                notUsed -= RTRN_QTY
                            ElseIf notUsed > 0 Then
                                rowSOTRMAFR.Item("RA_QTY_USED") = Val(rowSOTRMAFR.Item("RA_QTY_USED") & String.Empty) + notUsed
                                notUsed = 0
                            End If
                            If notUsed <= 0 Then Exit For
                        Next
                    Next

                    For Each rowSOTRMAFR As DataRow In dst.Tables("SOTRMAFR").Select("")
                        If Val(rowSOTRMAFR.Item("RA_RTN_QTY") & String.Empty) < Val(rowSOTRMAFR.Item("RA_QTY_USED") & String.Empty) Then
                            rowSOTRMAFR.Item("RA_QTY_USED") = rowSOTRMAFR.Item("RA_RTN_QTY")
                        End If
                    Next

                    Update_Record_TDA("SOTRMAFR")
                End If
            End If

            CreateAdjustment(Absx1.txtFor("RTRN_NO").Text)

            If ASCMAIN1.CLIENT = "VAN" AndAlso LoopReturnID.Length > 0 Then
                ASCMAIN1.sql = "UPDATE SOTRTNL1 SET PROCESS_IND = '1', RTRN_NO = :PARM1, INV_NO = :PARM2 WHERE RETURN_ID = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", {Absx1.txtFor("RTRN_NO").Text, INV_NO, LoopReturnID})

                For Each drSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("")
                    If Val(drSOTRTRN2.Item("RTRN_QTY") & String.Empty) > 0 Then
                        Dim LINE_ITEM_ID As String = drSOTRTRN2.Item("LINE_ITEM_ID") & String.Empty
                        Dim RTRN_QTY_REFUSED As Int32 = Val(drSOTRTRN2.Item("RTRN_QTY_REFUSED") & String.Empty)
                        If LINE_ITEM_ID.Length > 0 AndAlso RTRN_QTY_REFUSED = 0 Then
                            ASCMAIN1.sql = "UPDATE SOTRTNL2 SET RTRN_QTY = :PARM1 WHERE RETURN_ID = :PARM2 AND LINE_ITEM_ID = :PARM3"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NVV", {1, LoopReturnID, LINE_ITEM_ID})
                        End If
                    End If
                Next
            End If

            CommitTrans("Update Complete")

            ' Create Web Invoices
            If ASCMAIN1.CLIENT = "RGI" Then
                If ASCMAIN1.Running_in_VS Then
                    Stop
                End If
                Try
                    ASCMAIN1.Progress("Creating Web Invoice", "")
                    For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                        TAC.SOCMAIN1.CreateWebInvoice(Me, row.Item("INV_TYPE"), row.Item("INV_NO"))
                        ' Email Invoice to Sales Rep and Customer.
                        EmailInvoice(row.Item("INV_TYPE"), row.Item("INV_NO"))
                    Next
                Catch ex As Exception

                End Try
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                Back_to_Stock_Report(False)
            End If

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor

        ASCMAIN1.Progress("Now Preparing for Printing")

        Dim REPORT_NAME As String = "SORINVP1"
        Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If RPT <> "" Then RPT = REPORT_NAME

        Dim customerReportName As String = REPORT_NAME

        Select Case ASCMAIN1.CLIENT
            Case "RGI"
                customerReportName = "SORINVPR"
            Case "NYA"
                customerReportName = "SORINVPN"
        End Select

        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If

        Dim INV_NO As String = rowSOTRTRN1.Item("INV_NO")
        Dim sqlw As String = " and SOTINVH1.INV_TYPE = 'C' and SOTINVH1.INV_NO = '" & INV_NO & "'"
        REPORTS(REPORT_NAME).Fill_Records_RPT(New Object() {sqlw, True, "C"})
        Dim FILENAME As String = ""
        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("CONS_INV", "0")

            Select Case ASCMAIN1.CLIENT
                Case "RGI"
                    .CR_params.Add("EXPORT_INFO", "0")
                Case "NYA"
                    .CR_params.Add("EXPORT_INFO", "0")
            End Select

            Dim REPORT_NO As String = .Generate_Report(customerReportName, "", "", False, False, "", "PDF", INV_NO, False)
            FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
            .Print_Report_End(, True)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Show_Document(FILENAME)
        '  Return FILENAME
    End Sub

    Private Sub Back_to_Stock_Report(ByVal displayPrompt As Boolean)

        Try
            If Val(dst.Tables("SOTRTRN2").Compute("SUM(RTRN_QTY_1)", "") & String.Empty) = 0 Then
                If displayPrompt Then
                    MessageBox.Show("The total quantity in the Stock Column is 0.", "Back to Stock Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Sub
            End If

            dst.Tables("SOTRTRN2_RPT").Rows.Clear()
            For Each row As DataRow In dst.Tables("SOTRTRN2").Select("RTRN_QTY_1 > 0")
                dst.Tables("SOTRTRN2_RPT").ImportRow(row)
            Next

            If dst.Tables("SOTRTRN2_RPT").Rows.Count = 0 Then
                If displayPrompt Then
                    MessageBox.Show("The Return does not have any items placed back into the Stock column.", "Back to Stock Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Sub
            End If

            Dim WHSE_CODE As String = dst.Tables("SOTRTRN1").Rows(0).Item("WHSE_CODE") & String.Empty
            Fill_Records("ICTWHSE1", WHSE_CODE)
            If dst.Tables("ICTWHSE1").Rows.Count = 0 Then
                Exit Sub
            End If

            Dim WHSE_LOC_RTN As String = dst.Tables("ICTWHSE1").Rows(0).Item("WHSE_LOC_RTN") & String.Empty

            Print_Report_Begin()
            CR_params.Add("WHSE_LOC_RTN", WHSE_LOC_RTN)
            Generate_Report("SORRTRBR", "Back To Stock Report", , , , , False)

            Dim defaultPrinter As String = DefaultPrinterName()
            If defaultPrinter.Length > 0 Then
                Print_Report_End(True)
            Else
                Print_Report_End()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Back To Stock Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function DefaultPrinterName() As String
        Dim oPS As New System.Drawing.Printing.PrinterSettings

        Try
            DefaultPrinterName = oPS.PrinterName
        Catch ex As System.Exception
            DefaultPrinterName = ""
        Finally
            oPS = Nothing
        End Try
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "REASON_CODE"
                sql_where &= " AND RETURN_IND = '1'"

                'Case "RSRV_NO"

                '    If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                '        MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                '        Cancel = True
                '        Exit Sub
                '    End If
                '    sql_where = ""

                '    If InquiryMode Then
                '    Else
                '        sql_where &= " and SOTRSRV1.RSRV_STATUS = 'O' "
                '    End If

                '    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                '        sql_where &= " and SOTRSRV1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                '    End If
                '    If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                '        sql_where &= " and SOTRSRV1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                '    End If
        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("RTRN_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTRTRN1"
            E.COLUMN_NAME = "RTRN_NO"
            E.CODE_VALUE = Absx1.txtFor("RTRN_NO").Text
            E.DESC_VALUE = "Return"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRSRV1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    Sub Update_WHTLOCBX()
        If dst.Tables.Contains("WHTLOCB1") Then
            dst.Tables("WHTLOCB1").Rows.Clear()
            dst.Tables("WHTLOCB2").Rows.Clear()
        Else
            Create_TDA(dst.Tables.Add, "WHTLOCB1", "*")
            Create_TDA(dst.Tables.Add, "WHTLOCB2", "*")
        End If

        Dim rowSOTRTRN1 As DataRow = dst.Tables("SOTRTRN1").Rows(0)
        For Each row As DataRow In dst.Tables("SOTRTRN2").Select("")
            Dim TRAN_NO As String = row.Item("RTRN_NO")
            Dim TRAN_LNO As Integer = row.Item("RTRN_LNO")
            Dim WHSE_CODE As String = rowSOTRTRN1.Item("WHSE_CODE")
            Dim BAR_CODE As String = "0000000000" ' row.Item("BAR_CODE")
            Dim LOCATION_CODE As String = row.Item("LOCATION_CODE")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim RTRN_QTY As Int64 = Val(row.Item("RTRN_QTY") & "")

            Dim rowWHTLOCB1 As DataRow = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE})
            If rowWHTLOCB1 Is Nothing Then
                Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE}, False)
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE})
            End If

            If rowWHTLOCB1 Is Nothing Then
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").NewRow
                With rowWHTLOCB1
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOCATION_CODE") = LOCATION_CODE
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("LOCATION_QTY") = RTRN_QTY
                End With
                dst.Tables("WHTLOCB1").Rows.Add(rowWHTLOCB1)
            Else
                rowWHTLOCB1.Item("LOCATION_QTY") = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "") + RTRN_QTY
            End If

            Dim rowWHTLOCB2 As DataRow = dst.Tables("WHTLOCB2").NewRow
            With rowWHTLOCB2
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("LOCATION_CODE") = LOCATION_CODE
                .Item("BAR_CODE") = BAR_CODE
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("WHSE_TRAN_QTY") = RTRN_QTY
                .Item("WHSE_TRAN_TYPE") = "A"
                .Item("WHSE_TRAN_NO") = TRAN_NO
                .Item("WHSE_TRAN_LNO") = TRAN_LNO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LOCATION_CODE_OTHER") = ""
                .Item("SESSION_ID") = ""
            End With
            dst.Tables("WHTLOCB2").Rows.Add(rowWHTLOCB2)
        Next

        Update_Record_TDA("WHTLOCB1")
        Update_Record_TDA("WHTLOCB2")

        dst.Tables("WHTLOCB1").Rows.Clear()
        dst.Tables("WHTLOCB2").Rows.Clear()
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTRTRNX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTRTRN2, "SBBB", "Show Filter", "Style Status Inquiry", "Copy Price to All Lines", "Copy All Lines to Negate Inventory Impact")
        Load_Popup_Menu(grdSOTRTRN3, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTINVHH, "SSB", "Show Filter", "Show GroupBox", "Credit Entire Invoice")
        Load_Popup_Menu(grdSOTINVHX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdWHTRTRN1, "B", "Whse Returns Report")
        Load_Popup_Menu(grdSOTORDR2180, "SS", "Show Filter", "Show GroupBox")

        If ASCMAIN1.CLIENT = "VAN" Then
            Load_Popup_Menu(grd3PL, "BB", "Mark as Deleted", "3PL Returns Report")
        Else
            Load_Popup_Menu(grd3PL, "B", "3PL Returns Report")
        End If

        Load_Popup_Menu(grdSOTRMAFX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Cancel RMA Balance")
        Load_Popup_Menu(grdSOTRMAFR, "SSB", "Show Filter", "Show Voids", "Show Summary")
        Load_Popup_Menu(grdSOTRMAFRS, "SB", "Show Filter", "Show Details")

        Load_Popup_Menu(grdSOTRTNL1, "SSB", "Show Filter", "Show GroupBox", "Track Package")
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

        Select Case grd.Name
            Case grdSOTINVHH.Name
                tlb_btn = DirectCast(tlb_pop.Tools("Credit Entire Invoice"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not InquiryMode

            Case grdSOTRTRN2.Name
                tlb_btn = DirectCast(tlb_pop.Tools("Copy Price To All Lines"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" AndAlso LoopReturnID.Length = 0)
                tlb_btn = DirectCast(tlb_pop.Tools("Copy All Lines To Negate Inventory Impact"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" AndAlso LoopReturnID.Length = 0)

            Case grdSOTRTNL1.Name
                tlb_btn = DirectCast(tlb_pop.Tools("Track Package"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Enabled = (grdSOTRTNL1.ActiveRow.Band.Key = grdSOTRTNL1.DisplayLayout.Bands(2).Key)

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "3PL Returns Report"
                Print_Report_Begin()
                Generate_Report("WHR3PLRR", "3PL Returns Report", , , , , False)
                Print_Report_End()

            Case "Whse Returns Report"

                ASCMAIN1.sql = "Select * from WHTWRTN1 Where WH_RTN_NO = '" & grdWHTRTRN1.ActiveRow.Cells("WH_RTN_NO").Value & "'"
                Fill_Records("WHTWRTR1", , , ASCMAIN1.sql)


                ASCMAIN1.sql = " Select R2.WH_RTN_NO, R2.WH_RTN_LNO, R2. STYLE_CODE, IC.STYLE_DESC, COLOR_CODE, CTN_PACK_QTY, CARTONS, CTN_PACK_QTY * CARTONS" & vbCrLf _
                & " From WHTWRTN2 R2, ICTSTYL1 IC," & vbCrLf _
                & " (Select WH_RTN_NO, WH_RTN_LNO, COUNT(*) CARTONS from WHTWRTN3 Group By WH_RTN_NO, WH_RTN_LNO) R3" & vbCrLf _
                & " Where R2.WH_RTN_NO = R3.WH_RTN_NO" & vbCrLf _
                & " And R2.WH_RTN_LNO = R3.WH_RTN_LNO" _
                & " And R2.STYLE_CODE = IC.STYLE_CODE" _
                & " And R2.WH_RTN_NO = '" & grdWHTRTRN1.ActiveRow.Cells("WH_RTN_NO").Value & "'"
                Fill_Records("WHTWRTR2", , , ASCMAIN1.sql)

                Print_Report_Begin()
                Generate_Report("WHRWRTNR", "Whse Returns Report", , , , , False)
                Print_Report_End()

            Case "Copy All Lines to Negate Inventory Impact"

                If ASCMAIN1.CLIENT = "RGI" Then
                    MessageBox.Show("This function is disabled. Place the quantity destroyed in the 'Destroy' column. Placing a negative numbers in the 'Stock' or 'Destroy' column messes up the inventory.",
                                    "Copy Lines", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                Else
                    Dim TBL As DataTable = dst.Tables("SOTRTRN2").Clone
                    Dim RTRN_LNO As Int32 = Val(dst.Tables("SOTRTRN2").Compute("MAX (RTRN_LNO)", "") & "")
                    For Each row As DataRow In dst.Tables("SOTRTRN2").Select("", "RTRN_LNO")
                        TBL.Rows.Add(row.ItemArray)
                    Next
                    For Each row As DataRow In TBL.Select("", "RTRN_LNO")
                        RTRN_LNO += 1
                        Dim rowSOTRTRN2 As DataRow = dst.Tables("SOTRTRN2").NewRow
                        rowSOTRTRN2.ItemArray = row.ItemArray
                        rowSOTRTRN2.Item("RTRN_LNO") = RTRN_LNO
                        rowSOTRTRN2.Item("RTRN_QTY") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY") & "")
                        rowSOTRTRN2.Item("RTRN_QTY_1") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY_1") & "")
                        rowSOTRTRN2.Item("RTRN_QTY_2") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY_2") & "")
                        rowSOTRTRN2.Item("RTRN_QTY_3") = -1 * Val(rowSOTRTRN2.Item("RTRN_QTY_3") & "")
                        dst.Tables("SOTRTRN2").Rows.Add(rowSOTRTRN2)
                    Next

                    DisplayTotals()
                End If
            Case "Show Summary", "Show Details"
                If e.Tool.Key = "Show Summary" Then
                    splSCANS.Panel1Collapsed = True
                Else
                    splSCANS.Panel2Collapsed = True
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Track Package"
                Dim TRACKING_NUMBER As String = grd.ActiveRow.Cells("TRACKING_NUMBER").Value & String.Empty
                Dim CARRIER As String = (grd.ActiveRow.Cells("CARRIER").Value & String.Empty).ToString.ToUpper

                Dim drSOTCARR1 As DataRow = LookUp("SOTCARR1", CARRIER)
                If drSOTCARR1 Is Nothing Then
                    MessageBox.Show($"Unable to track Carrier {CARRIER}.", "Track Package", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim CARRIER_URL_TRACKING As String = drSOTCARR1.Item("CARRIER_URL_TRACKING") & String.Empty
                If CARRIER_URL_TRACKING.Length = 0 Then
                    MessageBox.Show($" Carrier {CARRIER} doe not have a TRacking URL.", "Track Package", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                ' update sotcarr1 set carrier_url_tracking = 'https://www.fedex.com/fedextrack/?trknbr=' where carrier_code = 'FEDEX';

                Try
                    Process.Start(CARRIER_URL_TRACKING & TRACKING_NUMBER)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Track Package", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Case "Cancel RMA Balance"
                Dim RA_NO As String = grd.ActiveRow.Cells("RA_NO").Value & String.Empty
                If MessageBox.Show("Do you want to cancel any open quantities on RMA - " & RA_NO & "?", "Cancel RMA Balance", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If


                CancelRMABalance(RA_NO)
                Load_SOTRMAFX()
                Show_Filter(grdSOTRMAFX, True)
                grdSOTRMAFX.DisplayLayout.GroupByBox.Hidden = False


            Case "Mark as Deleted"
                Dim TRANS_SEQ As String = grd.ActiveRow.Cells("TRANS_SEQ").Value

                If MsgBox("Do you Really want to mark Trans Seq " & TRANS_SEQ & " as Deleted?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    If ASCDATA1.ExecuteSQL("Update ADS.RCPTHDR@ADSIIS Set STATUS = 'D' where TRANS_SEQ = " & TRANS_SEQ & " and (STATUS = '0' or STATUS = 'V')") = 1 Then
                        grd.ActiveRow.Delete(False)
                    End If
                End If

                MsgBox("Trans Seq " & TRANS_SEQ & " has been Deleted", MsgBoxStyle.OkOnly, "Verification")

                'Case "Acknowledge w/Notes"
                '    Log_SetMode(True, True)

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Credit Entire Invoice"
                If grd.ActiveRow.IsDataRow Then
                    INV_NO_RETURNED = grd.ActiveRow.Cells("INV_NO").Value
                    Absx1.txtFor("CUST_CODE").Text = grd.ActiveRow.Cells("CUST_CODE").Value
                    Absx1.txtFor("WHSE_CODE").Text = grd.ActiveRow.Cells("WHSE_CODE").Value
                    'Absx1.txtFor("INV_NO_RETURNED").Text = grd.ActiveRow.Cells("INV_NO").Value
                    Click_Command("New")
                    If Not ScreenMode Then
                        INV_NO_RETURNED = ""
                    Else
                        DisplayTotals()
                    End If
                End If

            Case "Copy Price to All Lines"
                Dim RTRN_PRICE As Decimal = Val(grd.ActiveRow.Cells("RTRN_PRICE").Value & "")
                If MsgBox("OK to copy price " & Format(RTRN_PRICE, "$#.00") & " to all lines?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Copying Price")

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("RTRN_PRICE").Value = RTRN_PRICE
                    grow.Update()
                Next

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Show Voids"
                Dim dvw As DataView = dst.Tables("SOTRMAFR").DefaultView
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    dvw.RowFilter = ""
                Else
                    dvw.RowFilter = "GUN_STATUS <> 'V'"
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        ' Dim COL As String = Absx1.GetABSColumnName(sender)
        Select Case Absx1.GetABSColumnName(sender)

            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("WHSE_CODE").Text <> "" Then
                        ' Click_Command("New", e)
                    End If
                End If

            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("CUST_CODE").Text <> "" Then
                        ' Click_Command("New", e)
                    End If
                End If

            Case "RTRN_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If


            Case "INV_NO_RETURNED"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not ScreenMode Then
                        Dim INV_NO As String = Absx1.txtFor("INV_NO_RETURNED").Text ' DON'T KNOW WHY THIS DID NOT RETURN ANYTHING
                        INV_NO = DirectCast(sender, UltraWinEditors.UltraTextEditor).Text
                        INV_NO = ASCMAIN1.Format_Field(INV_NO, "INV_NO")
                        Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"I", INV_NO})
                        If rowSOTINVH1 Is Nothing Then
                            MsgBox("No Record of Invoice " & INV_NO, MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Else

                            '  INV_NO_RETURNED = INV_NO
                            Absx1.txtFor("CUST_CODE").Text = rowSOTINVH1.Item("CUST_CODE")
                            Absx1.txtFor("WHSE_CODE").Text = rowSOTINVH1.Item("WHSE_CODE")

                            preloadInvoiceDetails = False

                            If ASCMAIN1.CLIENT = "RGI" _
                                    AndAlso MessageBox.Show("Is this a Price Change Credit?", "Price Change", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                                preloadInvoiceDetails = True
                                Click_Command("Price Change")
                            Else
                                If MessageBox.Show("Do you want to preload the Invoice Details?", "Load Return", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                                    preloadInvoiceDetails = True
                                End If
                                Click_Command("New")
                            End If

                            'Click_Command("New")
                            If Not ScreenMode Then
                                INV_NO_RETURNED = ""
                            End If

                        End If
                    End If

                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If Not InquiryMode And Absx1.txtFor("WHSE_CODE").Text <> "" Then
                    ' Click_Command("New")
                End If
            Case "WHSE_CODE"
                If Not InquiryMode And Absx1.txtFor("CUST_CODE").Text <> "" Then
                    ' Click_Command("New")
                End If
            Case "RTRN_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                dst.Tables("SOTINVHH").Rows.Clear()
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                If CUST_CODE <> "" Then
                    Dim row As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    If row IsNot Nothing Then
                        Load_SOTINVHH()
                    End If
                End If

            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "REASON_CODE"
                Dim REASON_CODE As String = Absx1.txtFor("REASON_CODE").Text
                chkNoImpact.Visible = False
                If REASON_CODE <> "" Then

                    Dim rowARTREAS1 As DataRow = LookUp("ARTREAS1", REASON_CODE)
                    If rowARTREAS1 IsNot Nothing Then
                        chkNoImpact.Visible = True
                        chkNoImpact.Checked = (rowARTREAS1.Item("RETURN_NO_STOCK_IND") & "" = "1")
                    End If
                End If

            Case "WHSE_CODE"
                If ScreenMode Then
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    If WHSE_CODE <> "" Then
                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                        If rowICTWHSE1 IsNot Nothing Then
                            location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
                            whse_is_a_3PL = (rowICTWHSE1.Item("LP_CODE") & "" <> "")
                        End If
                    End If
                End If
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "RTRN_SALES", "RTRN_FREIGHT", "RTRN_HANDLING", "RTRN_STAX"
                Absx1.numFor("RTRN_AMOUNT").Value _
                    = Val(Absx1.numFor("RTRN_SALES").Value & "") _
                    + Val(Absx1.numFor("RTRN_FREIGHT").Value & "") _
                    + Val(Absx1.numFor("RTRN_STAX").Value & "") _
                    + Val(Absx1.numFor("RTRN_HANDLING").Value & "")
        End Select
    End Sub

#End Region

#Region "grdSOTRTRN2"

    Private Sub grdSOTRTRN2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRTRN2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

                grdCodeDesc(grdSOTRTRN2, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = cdr.Item("SALES_DIVISION_CODE")
                    Dim STYLE_CLASS_CODE As String = cdr.Item("STYLE_CLASS_CODE") & ""
                    Dim SALES_DIVISION_CODE As String = cdr.Item("SALES_DIVISION_CODE") & ""
                    Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")
                    e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = STYLE_CLASS_CODE
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = SALES_DIVISION_CODE
                    e.Cell.Row.Cells("STYLE_COST").Value = STYLE_COST
                    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")

                    'If location_support Then
                    '    e.Cell.Row.Cells("LOCATION_CODE").Value = ROWICTWHSE1.Item("WHSE_LOC_RTN")
                    '    ' USE ITEM_BIN AS A DEFAULT FOR AHA
                    'End If

                    ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
                    Dim rowICTSTYC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rowICTSTYC1s.Length = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = rowICTSTYC1s(0).Item("COLOR_CODE")
                    End If

                    If ScreenMode And Not IsLoading Then
                        Load_SOTINVHX(STYLE_CODE)
                        Dim row As DataRow = dst.Tables("SOTINVHMR").Select($"STYLE_CODE = '{STYLE_CODE}'").FirstOrDefault
                        If row IsNot Nothing Then
                            e.Cell.Row.Cells("SURCHARGE_PERC").Value = row("SURCHARGE_PERC")
                            e.Cell.Row.Cells("MISC_CHG_CODE").Value = row("MISC_CHG_CODE")
                            e.Cell.Row.Cells("COUNTRY_CODE").Value = row("COUNTRY_CODE")
                        End If
                    End If
                Else
                    grdSOTRTRN2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "COLOR_CODE"
                grdCodeDesc(grdSOTRTRN2, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
                If cdr IsNot Nothing Then
                    e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")
                End If

            Case "RTRN_QTY"

                If ASCMAIN1.CLIENT = "RGI" Then

                    Dim STYLE_CODE As String = grdSOTRTRN2.ActiveRow.Cells("STYLE_CODE").Value & ""
                    Dim COLOR_CODE As String = grdSOTRTRN2.ActiveRow.Cells("COLOR_CODE").Value & ""
                    Dim ORDR_PRICE_SOURCE As String = ""
                    Dim ORDR_UNIT_PRICE_CALC As Decimal = 0
                    If rowARTCUST1 Is Nothing OrElse rowARTCUST1.Item("CUST_CODE") & String.Empty <> MyBase.Absx1.txtFor("CUST_CODE").Text Then
                        rowARTCUST1 = LookUp("ARTCUST1", MyBase.Absx1.txtFor("CUST_CODE").Text)
                    End If
                    If Absx1.txtFor("WHSE_CODE").Text = "FE" Then
                        ORDR_PRICE_SOURCE = "FE"
                    Else
                        ORDR_PRICE_SOURCE = "FD"
                    End If

                    If dst.Tables("SOTINVHX").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length > 0 Then
                        ORDR_UNIT_PRICE_CALC = dst.Tables("SOTINVHX").Compute("MIN(ORDR_UNIT_PRICE)", "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")
                    Else
                        ORDR_UNIT_PRICE_CALC = TAC.SOCMAIN1.Price_Line(Me, HFs("CUST_CODE"), rowARTCUST1,
                                    grdSOTRTRN2.ActiveRow.Cells("STYLE_CODE").Value & "",
                                    grdSOTRTRN2.ActiveRow.Cells("COLOR_CODE").Value & "",
                                    Val(grdSOTRTRN2.ActiveRow.Cells("RTRN_QTY").Value & ""), ORDR_PRICE_SOURCE)

                    End If

                    e.Cell.Row.Cells("RTRN_PRICE").Value = ORDR_UNIT_PRICE_CALC
                    If Not IsDBNull(e.Cell.Row.Cells("SURCHARGE_PERC").Value) Then
                        e.Cell.Row.Cells("LINE_TARIFF").Value = Round(e.Cell.Row.Cells("RTRN_PRICE").Value * (e.Cell.Row.Cells("SURCHARGE_PERC").Value / 100), 2) * e.Cell.Row.Cells("RTRN_QTY").Value
                    End If
                End If

        End Select
    End Sub

    Private Sub grdSOTRTRN2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTRTRN2.AfterExitEditMode
        Select Case grdSOTRTRN2.ActiveCell.Column.Key
            'Case "ACCT_CODE"
            '    Dim ACCT_CODE As String = grdICTIXFR2.ActiveCell.Text
            '    If ACCT_CODE <> "" Then
            '        grdICTIXFR2.ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, grdGLTJRNL2.ActiveCell.Column.Key)
            '    End If
        End Select
    End Sub

    Private Sub grdSOTRTRN2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTRTRN2.AfterRowActivate
        With grdSOTRTRN2.DisplayLayout.Bands(0)
            If grdSOTRTRN2.ActiveRow.IsAddRow AndAlso (Not whse_is_a_3PL Or KEY_3PL_RECORD = "") AndAlso LoopReturnID.Length = 0 Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSOTRTRN2.ActiveCell = grdSOTRTRN2.ActiveRow.Cells("STYLE_CODE")
                grdSOTRTRN2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If EntryMode = "N" AndAlso Not grdSOTRTRN2.ActiveRow.IsAddRow AndAlso LoopReturnID.Length = 0 Then
            Load_SOTINVHX(grdSOTRTRN2.ActiveRow.Cells("STYLE_CODE").Value)
            Dim row As DataRow = dst.Tables("SOTINVHMR").Select($"STYLE_CODE = '{grdSOTRTRN2.ActiveRow.Cells("STYLE_CODE").Value}'").FirstOrDefault
            If row IsNot Nothing Then
                grdSOTRTRN2.ActiveRow.Cells("SURCHARGE_PERC").Value = row("SURCHARGE_PERC")
                grdSOTRTRN2.ActiveRow.Cells("MISC_CHG_CODE").Value = row("MISC_CHG_CODE")
                grdSOTRTRN2.ActiveRow.Cells("COUNTRY_CODE").Value = row("COUNTRY_CODE")
            End If
        End If

        If EntryMode = "V" Then
            Show_GL()
        End If
    End Sub

    Private Sub grdSOTRTRN2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTRTRN2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdSOTRTRN2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTRTRN2.AfterRowUpdate
        DisplayTotals()
    End Sub

    Private Sub grdSOTRTRN2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTRTRN2.BeforeExitEditMode
        If grdSOTRTRN2.ActiveCell Is Nothing Then Exit Sub
        With grdSOTRTRN2.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTSTYL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If

                Case "COLOR_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTCOLR1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Color Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                        If Not e.Cancel Then
                            cdr = LookUp("ICTSTYC1", New String() { .Row.Cells("STYLE_CODE").Value, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Color Code (" & .Text & ") not set up with Style (" & .Row.Cells("STYLE_CODE").Value & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If

                Case "LOCATION_CODE"
                    If location_support Then
                        If .Text <> "" Then
                            If .Value IsNot Nothing Then
                                .Value = .Text.ToUpper
                            End If

                        End If
                        If .Text <> "" Then
                            cdr = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Invalid Location Code (" & .Text & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdSOTRTRN2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTRTRN2.BeforeRowsDeleted

        'If whse_is_a_3PL Then
        '    RECORD_INDEXs = New List(Of Int32)
        '    For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '        RECORD_INDEXs.Add(grow.Cells("RECORD_INDEX").Value)
        '    Next
        'End If

    End Sub

    Private Sub grdSOTRTRN2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTRTRN2.BeforeRowUpdate
        With grdSOTRTRN2
            If e.Row.Cells("STYLE_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Row.Cells("COLOR_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTCOLR1", e.Row.Cells("COLOR_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Color Code (" & e.Row.Cells("COLOR_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
                If Not e.Cancel Then
                    LookUp("ICTSTYC1", New String() {e.Row.Cells("STYLE_CODE").Text, e.Row.Cells("COLOR_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Color Code (" & e.Row.Cells("COLOR_CODE").Text & ") not set up for Style (" & e.Row.Cells("STYLE_CODE").Text & ")",
                                   MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If



            If e.Row.Cells("RTV_REASON_CODE").Text = "" Then
                'e.Cancel = True OK TO NOT SPECIFY
            Else
                LookUp("SOTREASV", e.Row.Cells("RTV_REASON_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for RTV Reason Code (" & e.Row.Cells("RTV_REASON_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

                'If e.Row.Cells("LOCATION_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, e.Row.Cells("LOCATION_CODE").Text})
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Location Code (" & e.Row.Cells("LOCATION_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

            End If

            If Val(e.Row.Cells("RTRN_QTY").Text) = 0 Then
                MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("RTRN_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            Dim ORDR_NO As String = e.Row.Cells("ORDR_NO").Text
            Dim ORDR_LNO As String = Val(e.Row.Cells("ORDR_LNO").Text)

            If ORDR_NO.Length > 0 AndAlso ORDR_LNO > 0 Then
                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                If rowSOTORDR2 IsNot Nothing Then
                    If Val(e.Row.Cells("RTRN_QTY").Text) > Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) Then
                        MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("RTRN_QTY").Text & "), shipped only (" & Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("RTRN_NO").Text = "" Then
                    .ActiveRow.Cells("RTRN_NO").Value = Absx1.CtlFor("RTRN_NO").Text
                    .ActiveRow.Cells("RTRN_LNO").Value = Val(dst.Tables("SOTRTRN2").Compute("Max(RTRN_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdSOTRTRN2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRTRN2.ClickCellButton

        If grdSOTRTRN2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
            Case "COLOR_CODE"
                sql_where = "COLOR_CODE in (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & e.Cell.Row.Cells("STYLE_CODE").Value & "')"
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
            Case "RTV_REASON_CODE"
                sql_where = ""
        End Select
        grdClickCellButton(grdSOTRTRN2, sql_where, False)

    End Sub

    Private Sub grdSOTRTRN2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSOTRTRN2.Error
        grdSOTRTRN2.ActiveRow.CancelUpdate()
    End Sub

#End Region

#Region "Move to another Warehouse Location"

    Private Sub MoveInventoryToLocation()

        dst.Tables("WHTMOVE1").Rows.Clear()
        dst.Tables("WHTMOVE2").Rows.Clear()

        If Val(dst.Tables("SOTRTRN2").Compute("SUM(RTRN_QTY_1)", "") & String.Empty) = 0 Then
            Exit Sub
        End If

        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
        Dim WHSE_TRAN_TYPE As String = "M"

        Dim WHSE_CODE As String = dst.Tables("SOTRTRN1").Rows(0).Item("WHSE_CODE") & String.Empty
        Fill_Records("ICTWHSE1", WHSE_CODE)
        If dst.Tables("ICTWHSE1").Rows.Count = 0 Then
            Exit Sub
        End If

        Dim WHSE_LOC_RTN As String = dst.Tables("ICTWHSE1").Rows(0).Item("WHSE_LOC_RTN") & String.Empty

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        With rowWHTMOVE1
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("WHSE_TRAN_TYPE") = WHSE_TRAN_TYPE
            .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("STATUS") = "U"
        End With
        dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)

        Dim WHSE_TRAN_LNO_ctr As Integer = 0
        For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("RTRN_QTY_1 > 0", "RTRN_LNO")

            Dim LOCATION_FROM As String = rowSOTRTRN2.Item("RTN_LOCATION_CODE") & String.Empty
            If LOCATION_FROM.Length = 0 Then
                LOCATION_FROM = WHSE_LOC_RTN
                rowSOTRTRN2.Item("RTN_LOCATION_CODE") = LOCATION_FROM
            End If

            Dim UNITS_MOVED As Int16 = Val(rowSOTRTRN2.Item("RTRN_QTY_1") & String.Empty)
            Dim STYLE_CODE As String = rowSOTRTRN2.Item("STYLE_CODE") & String.Empty
            Dim COLOR_CODE As String = rowSOTRTRN2.Item("COLOR_CODE") & String.Empty

            Dim row As DataRow = GetLocation(STYLE_CODE, COLOR_CODE, WHSE_CODE)
            If row Is Nothing Then
                rowSOTRTRN2.Item("RTN_LOCATION_CODE") = WHSE_LOC_RTN
                Continue For
            End If

            Dim LOCATION_CODE As String = row.Item("LOCATION_CODE") & String.Empty
            If LOCATION_CODE = String.Empty OrElse LOCATION_CODE = WHSE_LOC_RTN Then
                Continue For
            End If

            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
            With rowWHTMOVE2
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                WHSE_TRAN_LNO_ctr += 1
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                .Item("LOCATION_CODE_FROM") = LOCATION_FROM
                .Item("LOCATION_CODE_TO") = LOCATION_CODE
                .Item("BAR_CODE") = "0000000000"
                .Item("WHSE_TRAN_QTY") = UNITS_MOVED
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"
                .Item("LOAD_NO_FROM") = ""
                .Item("LOAD_NO_TO") = ""
            End With
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
            rowSOTRTRN2.Item("RTN_LOCATION_CODE") = LOCATION_CODE
        Next

        If dst.Tables("WHTMOVE2").Rows.Count > 0 Then
            Update_Record_TDA("SOTRTRN2")
            Update_Record_TDA("WHTMOVE1")
            Update_Record_TDA("WHTMOVE2")
            ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                           New Object() {WHSE_TRAN_NO, 0, 1},
                           New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})
        Else
            dst.Tables("WHTMOVE1").Rows.Clear()
            dst.Tables("WHTMOVE2").Rows.Clear()
        End If

    End Sub

    Function GetLocation(ByVal Style As String, ByVal Color As String, ByVal WHSE_CODE As String) As DataRow

        ASCMAIN1.sql = " select b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE " & vbCrLf _
            & " from whtlocb1 b1 " & vbCrLf _
            & "  join whtlocm1 m1 on b1.LOCATION_CODE = m1.LOCATION_CODE and b1.WHSE_CODE = m1.WHSE_CODE " & vbCrLf _
            & "  where b1.STYLE_CODE = '" & Style & "' and b1.COLOR_CODE = '" & Color & "' " & vbCrLf _
            & "  and  nvl(m1.LOCATION_USE,'A') in ('A','C') and m1.LOCATION_ROUTE_SEQ is not null" & vbCrLf _
            & "  and b1.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "  order by b1.LOCATION_QTY, m1.LOCATION_ROUTE_SEQ, m1.LOCATION_CODE"

        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        If tbl.Rows.Count = 0 Then
            Return Nothing
        Else
            Return tbl.Select("", "LOCATION_QTY DESC")(0)
        End If

    End Function

#End Region

#Region "Form Controls"

    Private Sub grdSOTRTNL1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTRTNL1.DoubleClickRow

        LoopReturnID = String.Empty

        If Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Dim RETURN_ID As String = e.Row.Cells("RETURN_ID").Value & String.Empty
        LookUpReturn(RefundLookupTypes.ReturnID, RETURN_ID)
    End Sub

    Private Sub grdSOTRTRNX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRTRNX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RTRN_NO").Text = e.Row.Cells("RTRN_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdSOTRTRNG_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRTRN3.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("RTRN_NO").Text = e.Row.Cells("RTRN_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_GL()
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Private Sub cbeInvoiceHistory_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeInvoiceHistory.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Load_SOTINVHH()
    End Sub

    Private Sub chkGL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGL.CheckedChanged
        Setup_tab0_GL()
    End Sub

    Private Sub grdSOTRTRN2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRTRN2.InitializeRow
        With e.Row.Cells("RTRN_QTY")
            If Val(e.Row.Cells("RTRN_QTY_TOTAL").Value & "") <> Val(.Value & "") Then
                .Appearance.ForeColor = Color.Red
                .ToolTipText = "Total Return does not balance with Sum of Stock + Refurb + Destroy"
            Else
                .Appearance.ForeColor = Color.Empty
                .ToolTipText = ""
            End If
        End With
        If e.Row.Cells("IMPORTED").Value.ToString = "1" Then

            e.Row.Appearance.ForeColor = Color.DarkOrange
            e.Row.ToolTipText = "This line was added by a Gun scan"
        End If
    End Sub

    Private Sub grdSOTINVHH_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTINVHH.DoubleClickRow

        preloadInvoiceDetails = False
        If Not InquiryMode Then

            ' Sometimes the doublw click does not put the Invoice Number in the textbox. Freaky
            If e.Row.Cells("INV_NO").Text = String.Empty Then
                MessageBox.Show("Cannot determine the selected Invoice.", "Select", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' Sometimes the text box clears out. force it in there
            Dim INV_NO_RETURNED As String = e.Row.Cells("INV_NO").Text

            Absx1.txtFor("INV_NO_RETURNED").Text = INV_NO_RETURNED
            txtINV_NO_RETURNED.Text = INV_NO_RETURNED

            If Absx1.txtFor("INV_NO_RETURNED").Text = String.Empty Then
                MessageBox.Show("Cannot determine the selected Invoice.", "Select", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If ASCMAIN1.CLIENT = "RGI" _
                    AndAlso MessageBox.Show("Is this a Price Change Credit?", "Price Change", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                preloadInvoiceDetails = True
                Absx1.txtFor("INV_NO_RETURNED").Text = e.Row.Cells("INV_NO").Text
                txtINV_NO_RETURNED.Text = INV_NO_RETURNED
                Click_Command("Price Change")
            Else
                If MessageBox.Show("Do you want to preload the Invoice Details?", "Load Return", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                    preloadInvoiceDetails = True
                End If
                txtINV_NO_RETURNED.Text = INV_NO_RETURNED
                Click_Command("New")
            End If
        End If
    End Sub

    Private Sub grd3PL_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grd3PL.DoubleClickRow

        If InquiryMode Then Exit Sub

        KEY_3PL_RECORD = e.Row.Cells("TRANS_SEQ").Value
        Dim WHSE_CODE As String = ""
        Dim prow As UltraWinGrid.UltraGridRow = e.Row
        If e.Row.ParentRow IsNot Nothing Then prow = e.Row.ParentRow
        WHSE_CODE = prow.Cells("WHSE_CODE").Value
        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
        If Absx1.txtFor("CUST_CODE").Text = "" And prow.Cells("REF1").Value & "" <> "" Then
            Dim REF1 As String = prow.Cells("REF1").Value & ""
            If LookUp("ARTCUST1", REF1) IsNot Nothing Then
                Absx1.txtFor("CUST_CODE").Text = REF1
            End If
        End If

        ASCMAIN1.sql = "Select * from ADS.RCPTHDR@ADSIIS where TRANS_SEQ = " & KEY_3PL_RECORD
        Dim rowRCPTHDR As DataRow = ASCDATA1.GetDataRow
        If rowRCPTHDR Is Nothing OrElse rowRCPTHDR.Item("STATUS") <> "0" Then
            MsgBox("Status Changed since Grid was Prepared - Please refresh Grid", MsgBoxStyle.OkOnly, "Cannot Process this Record - Status Changed since grid was prepared")
            KEY_3PL_RECORD = "'"
            Exit Sub
        End If

        Click_Command("New")

        If Not ScreenMode Then
            KEY_3PL_RECORD = ""
        Else
            Synch_TABLE_NAME("SOTRTRN1")
            rowSOTRTRN1.Item("KEY_3PL_RECORD") = KEY_3PL_RECORD
            Absx1.txtFor("CUST_CLAIM_NO").Text = prow.Cells("PO_SHIPMENT_NO").Value & ""
            Absx1.txtFor("RTRN_NOTE").Text = prow.Cells("CONTAINER_NO").Value & ""

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Return")

            grdSOTRTRN2.Visible = False

            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
            For Each crow As UltraWinGrid.UltraGridRow In prow.ChildBands(0).Rows
                Dim ITEM_CODE As String = crow.Cells("ITEM_CODE").Value

                'If ITEM_CODE.EndsWith("PPK") Then
                '    ASCMAIN1.sql = "Select Sum (PPK_QTY) from WHTPPKM2 where PPK_CODE = '" & ITEM_CODE & "'"
                '    Dim PPK_QTY As Int64 = Val(ASCDATA1.GetDataValue)
                '    ASCMAIN1.sql = "Select * from WHTPPKM2 where PPK_CODE = '" & ITEM_CODE & "'"
                '    For Each rowWHTPPKM2 As DataRow In ASCDATA1.GetDataTable.Select("")
                '        grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                '        With grdSOTRTRN2.ActiveRow
                '            .Cells("STYLE_CODE").Value = rowWHTPPKM2.Item("STYLE_CODE")
                '            .Cells("COLOR_CODE").Value = rowWHTPPKM2.Item("COLOR_CODE")
                '            Dim QTY As Int64 = Val(rowWHTPPKM2.Item("PPK_QTY") & "") * Val(crow.Cells("RCVQTY").Value & "") / PPK_QTY
                '            .Cells("RTRN_QTY").Value = QTY
                '            .Cells("RTRN_QTY_1").Value = QTY
                '            ' .Cells("KEY_3PL_RECORD").Value = ITEM_CODE
                '            .Update()
                '        End With
                '    Next

                'Else
                grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
                With grdSOTRTRN2.ActiveRow
                    If ITEM_CODE.EndsWith("PPK") Then
                        .Cells("STYLE_CODE").Value = crow.Cells("STYLE_CODE").Value
                        .Cells("COLOR_CODE").Value = crow.Cells("COLOR_CODE").Value
                    Else
                        .Cells("STYLE_CODE").Value = Mid(ITEM_CODE, 1, Len(ITEM_CODE) - 3)
                        .Cells("COLOR_CODE").Value = Mid(ITEM_CODE, Len(ITEM_CODE) - 3 + 1)
                    End If

                    .Cells("RTRN_QTY").Value = crow.Cells("RCVQTY").Value
                    .Cells("RTRN_QTY_1").Value = crow.Cells("RCVQTY").Value
                    .Update()
                End With
                ' End If

            Next

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

            grdSOTRTRN2.Visible = True
            grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Sort_grdColumns(grdSOTRTRN2, "RTRN_LNO")
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        Setup_tab0()
    End Sub

    Private Sub grdSOTRTRNX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRTRNX.InitializeRow
        If e.Row.Cells("REVERSED_BY_RTRN_NO").Value & "" <> "" Then
            e.Row.Cells("RTRN_NO").Appearance.ForeColor = Color.Red
            e.Row.Cells("RTRN_NO").ToolTipText = "Reversed by Return No " & e.Row.Cells("REVERSED_BY_RTRN_NO").Value
        ElseIf e.Row.Cells("REVERSES_RTRN_NO").Value & "" <> "" Then
            e.Row.Cells("RTRN_NO").Appearance.ForeColor = Color.Red
            e.Row.Cells("RTRN_NO").ToolTipText = "Reverses Return No " & e.Row.Cells("REVERSES_RTRN_NO").Value
        Else
            e.Row.Cells("RTRN_NO").Appearance.ForeColor = Color.Empty
            e.Row.Cells("RTRN_NO").ToolTipText = ""
        End If
    End Sub

    Private Sub grdSOTRMAFR_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRMAFR.InitializeRow
        If e.Row.Cells("GUN_STATUS").Value & "" = "V" Then
            e.Row.Appearance.ForeColor = Color.Red
            e.Row.ToolTipText = "This line was Voided, Do not include"
        End If
    End Sub

    Private Sub grdWHTRTRN1_DoubleClick(sender As Object, e As System.EventArgs) Handles grdWHTRTRN1.DoubleClick

        Whse_Rtn_no = grdWHTRTRN1.ActiveRow.Cells("WH_RTN_NO").Value
        Absx1.txtFor("WHSE_CODE").Text = grdWHTRTRN1.ActiveRow.Cells("WHSE_CODE").Value
        Absx1.txtFor("CUST_CODE").Text = grdWHTRTRN1.ActiveRow.Cells("CUST_CODE").Value
        Absx1.txtFor("RTRN_NOTE").Text = grdWHTRTRN1.ActiveRow.Cells("WH_RTN_COMMENT").Value & ""

        ' TIME IS IN WH_RTN_DATE
        Dim RTRN_DATE As Date = CDate(grdWHTRTRN1.ActiveRow.Cells("WH_RTN_DATE").Value).Date
        Absx1.dteFor("RTRN_DATE").Value = RTRN_DATE ' grdWHTRTRN1.ActiveRow.Cells("WH_RTN_DATE").Value

        ' Absx1.txtFor("INV_NO_RETURNED").ReadOnly = False

        Click_Command("New")

        grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
        grdSOTRTRN2.Visible = False
        ASCMAIN1.sql = " Select R2.STYLE_CODE, R2.COLOR_CODE, Sum(R3.QTY_RTN) as QTY_RTN" _
        & " From WHTWRTN2 R2, WHTWRTN3 R3" _
        & " Where R2.WH_RTN_NO = R3.WH_RTN_NO" _
        & " And R2.WH_RTN_LNO = R3.WH_RTN_LNO" _
        & " And R2.WH_RTN_NO = '" & grdWHTRTRN1.ActiveRow.Cells("WH_RTN_NO").Value & "'" _
        & " Group by R2.STYLE_CODE, R2.COLOR_CODE"
        For Each rowWHTWRTN2 As DataRow In ASCDATA1.GetDataTable.Rows
            grdSOTRTRN2.DisplayLayout.Bands(0).AddNew()
            With grdSOTRTRN2.ActiveRow
                .Cells("STYLE_CODE").Value = rowWHTWRTN2.Item("STYLE_CODE")
                .Cells("COLOR_CODE").Value = rowWHTWRTN2.Item("COLOR_CODE")
                .Cells("RTRN_QTY").Value = rowWHTWRTN2.Item("QTY_RTN")
                .Cells("RTRN_QTY_1").Value = rowWHTWRTN2.Item("QTY_RTN")
                .Update()
            End With
        Next
        'grdSOTRTRN2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
        grdSOTRTRN2.DisplayLayout.Bands(0).Columns("RTRN_QTY_1").CellActivation = UltraWinGrid.Activation.NoEdit
        Sort_grdColumns(grdSOTRTRN2, "RTRN_LNO")
        grdSOTRTRN2.Visible = True
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click

        If Absx1.txtFor("CUST_CODE").TextLength = 0 Then
            MessageBox.Show("Customer Code is required for search.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
        Dim SHIP_TO_NAME As String = Absx1.txtFor("SHIP_TO_NAME").Text
        SHIP_TO_NAME = SHIP_TO_NAME.Trim.Replace("'", "")
        Absx1.txtFor("SHIP_TO_NAME").Text = SHIP_TO_NAME

        If SHIP_TO_NAME.Length > 1 AndAlso SHIP_TO_NAME.Length < 3 Then
            MessageBox.Show("Ship To Name must be at least 3 characters.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If STYLE_CODE.Length = 0 AndAlso SHIP_TO_NAME.Length = 0 Then
            MessageBox.Show("Ship To Name and/or Style Code is required.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ASCMAIN1.sql = sqlSOTORDR2180
        ASCMAIN1.sql &= " AND SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"

        If SHIP_TO_NAME.Length > 0 Then
            ASCMAIN1.sql &= " AND UPPER(SOTORDR1.CUST_STORE_NAME) LIKE '%" & SHIP_TO_NAME.ToUpper & "%' "
        End If

        If STYLE_CODE.Length > 0 Then
            ASCMAIN1.sql &= " AND SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'"
        End If

        ASCMAIN1.sql &= " AND SOTORDR1.ORDR_DATE >= '" & Absx1.dteFor("SEARCH_START_DATE").DateTime.ToString("dd-MMM-yyyy") & "'"

        Fill_Records("SOTORDR2180", String.Empty, True, ASCMAIN1.sql)

        Sort_grdColumns(grdSOTORDR2180, "ORDR_DATE")

    End Sub

    Private Sub grdSOTORDR2180_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDR2180.DoubleClickRow
        If Not InquiryMode Then
            txtINV_NO_RETURNED.Text = e.Row.Cells("INV_NO").Text & String.Empty
        End If
    End Sub

    Private Sub grdSOTRMAFX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRMAFX.DoubleClickRow
        If grdSOTRMAFX.ActiveRow Is Nothing Then Exit Sub

        ' This code needs to act like the Invoice double click grid
        Absx1.txtFor("RA_NO").Text = grdSOTRMAFX.ActiveRow.Cells("RA_NO").Value & String.Empty
        Click_Command("New")
    End Sub

    Private Sub grdSOTRMAFX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRMAFX.InitializeRow
        If e.Row.Band.Key = grdSOTRMAFX.DisplayLayout.Bands(0).Key Then
            If IsDate(e.Row.Cells("RA_EXPIRE").Text & String.Empty) AndAlso DateTime.Now.CompareTo(CDate(e.Row.Cells("RA_EXPIRE").Text)) = 1 Then
                e.Row.Appearance.BackColor = Color.PaleVioletRed
            End If
        End If
    End Sub

    Private Sub numRTRN_FREIGHT_EditorButtonClick(sender As Object, e As UltraWinEditors.EditorButtonEventArgs) Handles numRTRN_FREIGHT.EditorButtonClick

        Dim INV_NO_RETURNED As String = dst.Tables("SOTRTRN1").Rows(0).Item("INV_NO_RETURNED") & String.Empty
        Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", New String() {"I", INV_NO_RETURNED})
        If rowSOTINVH1 IsNot Nothing Then
            numRTRN_FREIGHT.Value = Val(rowSOTINVH1.Item("INV_FREIGHT") & String.Empty)
            Exit Sub
        End If
        MessageBox.Show("Cannot determine the Invoice Number.", "Freight", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub CreateAdjustment(ByVal RTRN_NO As String)

        ' Create Adjustmenst for the destroyed quantities.

        If ASCMAIN1.CLIENT <> "RGI" Then
            Exit Sub
        End If

        If dst.Tables("SOTRTRN2").Select("RTRN_NO = '" & RTRN_NO & "' and ISNULL(RTRN_QTY_3, 0) > 0").Length = 0 Then
            Exit Sub
        End If

        Dim ADJ_NO As String = String.Empty
        Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").NewRow

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
        Else
            ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        End If

        rowICTIADJ1.Item("ADJ_NO") = ADJ_NO
        rowICTIADJ1.Item("ADJ_DATE") = DateTime.Now.ToShortDateString
        rowICTIADJ1.Item("WHSE_CODE") = Absx1.txtFor("WHSE_CODE").Text
        rowICTIADJ1.Item("REASON_CODE") = "WHADJ"
        rowICTIADJ1.Item("ADJ_NOTE") = "Return Number " & RTRN_NO
        rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("INIT_DATE") = DateTime.Now
        rowICTIADJ1.Item("REGISTER_IND") = "0"
        'rowICTIADJ1.Item("REGISTER_XNO") = String.Empty
        rowICTIADJ1.Item("ADJ_SOURCE") = "E"
        rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowICTIADJ1.Item("TOTAL_COSTS") = 0
        rowICTIADJ1.Item("RTRN_NO") = RTRN_NO
        rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("LAST_DATE") = DateTime.Now
        rowICTIADJ1.Item("REVERSED_BY_ADJ_NO") = String.Empty
        rowICTIADJ1.Item("REVERSES_ADJ_NO") = String.Empty
        'rowICTIADJ1.Item("ADJ_REF") = String.Empty
        rowICTIADJ1.Item("JOURNAL_IND") = "0"
        'rowICTIADJ1.Item("JOURNAL_XNO") = String.Empty
        dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

        Dim ADJ_LNO As Int32 = 1

        For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("RTRN_NO = '" & RTRN_NO & "' and ISNULL(RTRN_QTY_3, 0) > 0")
            Dim ADJ_QTY As Int32 = Val(rowSOTRTRN2.Item("RTRN_QTY_3") & String.Empty)

            Dim rowICTIADJ2 As DataRow = dst.Tables("ICTIADJ2").NewRow
            rowICTIADJ2.Item("ADJ_NO") = ADJ_NO
            rowICTIADJ2.Item("ADJ_LNO") = ADJ_LNO
            rowICTIADJ2.Item("STYLE_CODE") = rowSOTRTRN2.Item("STYLE_CODE") & String.Empty
            rowICTIADJ2.Item("COLOR_CODE") = rowSOTRTRN2.Item("COLOR_CODE") & String.Empty
            rowICTIADJ2.Item("ADJ_QTY") = ADJ_QTY * -1
            rowICTIADJ2.Item("STYLE_COST") = Val(rowSOTRTRN2.Item("STYLE_COST") & String.Empty)
            rowICTIADJ2.Item("STYLE_CLASS_CODE") = rowSOTRTRN2.Item("STYLE_CLASS_CODE")
            rowICTIADJ2.Item("SALES_DIVISION_CODE") = rowSOTRTRN2.Item("SALES_DIVISION_CODE")
            rowICTIADJ2.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIADJ2.Item("LOCATION_CODE") = WHSE_LOC_DST
            'rowICTIADJ2.item("BAR_CODE") = String.Empty
            'rowICTIADJ2.item("ADJ_REF") = String.Empty
            dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)

            ADJ_LNO += 1
        Next

        rowICTIADJ1.Item("TOTAL_COSTS") = Val(dst.Tables("ICTIADJ2").Compute("SUM(LINE_COSTS)", "ADJ_NO = '" & ADJ_NO & "'") & "")

        ICCMAIN1.Update_Adjustment(Me)

        If location_support Then

            ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                     New Object() {"A", rowICTIADJ1.Item("ADJ_NO"), ASCMAIN1.SESSION_NO},
                     New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
        End If

    End Sub

    Sub CancelRMABalance(ByVal RA_NO As String)

        Try
            Me.Cursor = Cursors.WaitCursor
            BeginTrans()

            ASCMAIN1.sql = "" _
                & "Begin " _
                & " Declare Cursor C1 is Select * from SOTRMAF2 where RA_NO = '" & RA_NO & "' for Update;" _
                & " Begin " _
                & "  For R1 in C1 Loop" _
                & "   Update SOTRMAF2" _
                & "    Set RA_QTY_CANC = NVL(RA_QTY_CANC,0) + NVL(R1.RA_QTY_OPEN,0)" _
                & "      , RA_QTY_OPEN = 0" _
                & "    where Current of C1;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Update SOTRMAF1 Set RA_STATUS = :PARM1, LAST_OPER = :PARM2, LAST_DATE = SYSDATE where RA_NO = :PARM3"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {"F", ASCMAIN1.USER_ID, RA_NO})

            CommitTrans("Balance Open on Returns Authorization " & RA_NO & " has been Cancelled")

        Catch ex As Exception
            Rollback("Error: " & ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub Load_SOTRMAFX()

        Static wktable As String = String.Empty

        dst.Tables("SOTRMAFXR").Rows.Clear()
        dst.Tables("SOTRMAFX2").Rows.Clear()
        dst.Tables("SOTRMAFX").Rows.Clear()

        ASCMAIN1.sql = "Select SOTRMAF1.*, ARTCUST1.CUST_NAME from SOTRMAF1, ARTCUST1 where RA_STATUS = 'O' AND SOTRMAF1.CUST_CODE = ARTCUST1.CUST_CODE (+)"

        If wktable.Length = 0 Then
            wktable = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        End If

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & wktable)
        ASCDATA1.ExecuteSQL("INSERT INTO " & wktable & " Select SOTRMAF1.*, ARTCUST1.CUST_NAME from SOTRMAF1, ARTCUST1 where RA_STATUS = 'O' AND SOTRMAF1.CUST_CODE = ARTCUST1.CUST_CODE (+)")

        ASCMAIN1.sql = "SELECT * FROM " & wktable
        Fill_Records("SOTRMAFX", "", , ASCMAIN1.sql)
        grdSOTRMAFX.Text = "Open Returns Authorizations"
        Sort_grdColumns(grdSOTRMAFX, "RA_NO".ToLower)
        grdSOTRMAFX.Visible = True

        ASCMAIN1.sql = "SELECT * FROM SOTRMAF2 WHERE RA_NO IN (Select RA_NO FROM " & wktable & ")"
        Fill_Records("SOTRMAFX2", "", , ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM SOTRMAFR WHERE RA_NO IN (Select RA_NO FROM " & wktable & ") and GUN_STATUS <> 'V'"
        Fill_Records("SOTRMAFXR", "", , ASCMAIN1.sql)

    End Sub

    Sub DisplayTotals()

        Dim RTRN_SALES As Decimal = 0
        Select Case dst.Tables("SOTRTRN1").Rows(0).Item("CURR_CODE") & String.Empty
            Case "", GL_PARM_CURR_CODE
                RTRN_SALES = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_SALES)", "") & "")
            Case Else
                RTRN_SALES = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_SALES_CURR)", "") & "")
        End Select
        Absx1.numFor("RTRN_SALES").Value = RTRN_SALES

        Dim RTRN_HANDLING As Decimal = 0
        If ASCMAIN1.CLIENT = "RGI" Then
            RTRN_HANDLING = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_TARIFF)", "") & "")
            Absx1.numFor("RTRN_HANDLING").Value = RTRN_HANDLING
        End If

        Dim RTRN_COSTS As Decimal = Val(dst.Tables("SOTRTRN2").Compute("SUM(LINE_COSTS)", "") & "")
        Absx1.numFor("RTRN_COSTS").Value = RTRN_COSTS

        If ASCMAIN1.CLIENT = "VAN" Then
            Dim REFUND_TAX As Decimal = Val(dst.Tables("SOTRTRN2").Compute("SUM(REFUND_TAX)", "") & "")
            Absx1.numFor("RTRN_STAX").Value = REFUND_TAX
        End If
    End Sub

    Sub Show_GL()

        If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
            optGL.Tag = optGL.Value
            If optGL.Value = "A" Then
                grdSOTRTRN3.DataSource = dst.Tables("SOTRTRN3")
                Dim dvw As DataView = dst.Tables("SOTRTRN3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdSOTRTRN3.DataSource = dst.Tables("SOTRTRN3")
                Dim dvw As DataView = dst.Tables("SOTRTRN3").DefaultView
                Dim RTRN_LNO As Integer = 0
                If grdSOTRTRN2.ActiveRow IsNot Nothing Then
                    RTRN_LNO = Val(grdSOTRTRN2.ActiveRow.Cells("RTRN_LNO").Text)
                End If
                dvw.RowFilter = "RTRN_LNO = " & CStr(RTRN_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("SOTRTRN3").Clone
                Dim RTRN_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("SOTRTRN3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = dst.Tables("SOTRTRN3").Compute _
                    ("SUM(DIST_AMT)",
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("RTRN_NO") = Absx1.txtFor("RTRN_NO").Text
                    row.Item("RTRN_LNO") = 0
                    RTRN_GNO += 1
                    row.Item("RTRN_GNO") = RTRN_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdSOTRTRN3.DataSource = tbl
            End If
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Fill_Records("SOTRTRNX", YP)

        grdSOTRTRNX.Text = "Entered in " & cbeYP.Text
        If chkGL.Checked Then
            Fill_Records("SOTRTRNG", YP)
            grdSOTRTRN3.Text = "Entered in " & cbeYP.Text
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            EnforceConstraints(False)
            dst.Tables("SOTRTNL3").Rows.Clear()
            dst.Tables("SOTRTNL2").Rows.Clear()
            dst.Tables("SOTRTNL1").Rows.Clear()

            Fill_Records("SOTRTNL1")
            If dst.Tables("SOTRTNL1").Rows.Count > 0 Then
                Dim lstReturnIds As New List(Of String)
                For Each drSOTRTNL1 As DataRow In dst.Tables("SOTRTNL1").Select("")
                    lstReturnIds.Add(drSOTRTNL1.Item("RETURN_ID"))
                Next
                Fill_Records("SOTRTNL2", {String.Join(",", lstReturnIds.ToArray)})
                Fill_Records("SOTRTNL3", {String.Join(",", lstReturnIds.ToArray)})
            End If
            EnforceConstraints(True)
        End If

    End Sub

    Sub Setup_tab0_GL()
        If Not chkGL.Checked Then
            tab0.Tabs(0).Selected = True
        Else
            Refresh_Documents()
        End If
        tab0.Tabs("GL").Visible = chkGL.Checked

        If chkGL.Checked Then
            tab0.Tabs("GL").Selected = True
        End If
    End Sub

    Sub Set_Up_Reversal()

        Dim REVERSED_BY_RTRN_NO As String = ""
        If ASCMAIN1.CLIENT = "VAN" Then
            REVERSED_BY_RTRN_NO = ASCMAIN1.Next_Control_No("TRAN_NO_C")
        Else
            REVERSED_BY_RTRN_NO = ASCMAIN1.Next_Control_No("SOTRTRN1.RTRN_NO")
        End If

        rowSOTRTRN1 = dst.Tables("SOTRTRN1").Rows(0)
        rowSOTRTRN1.AcceptChanges()
        rowSOTRTRN1.SetAdded()

        With rowSOTRTRN1
            .Item("REVERSES_RTRN_NO") = .Item("RTRN_NO")
            .Item("RTRN_NO") = REVERSED_BY_RTRN_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("RTRN_DATE") = .Item("RTRN_DATE") ' DATETIME_STAMP.Date
            .Item("RTRN_SALES") = -1 * Val(.Item("RTRN_SALES") & "")
            .Item("RTRN_COSTS") = -1 * Val(.Item("RTRN_COSTS") & "")
            .Item("RTRN_STAX") = -1 * Val(.Item("RTRN_STAX") & "")
            .Item("RTRN_FREIGHT") = -1 * Val(.Item("RTRN_FREIGHT") & "")
            .Item("RTRN_HANDLING") = -1 * Val(.Item("RTRN_HANDLING") & "")
            .Item("RTRN_AMOUNT") = -1 * Val(.Item("RTRN_AMOUNT") & "")
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
        End With

        'Set new RTRN_NO and reverse all quantities for this return.
        For Each row As DataRow In dst.Tables("SOTRTRN2").Rows
            row.Item("RTRN_NO") = REVERSED_BY_RTRN_NO
            For Each C As String In New String() {"RTRN_QTY", "RTRN_QTY_1", "RTRN_QTY_2", "RTRN_QTY_3"}
                If Val(row.Item(C) & "") <> 0 Then
                    row.Item(C) = -1 * Val(row.Item(C) & "")
                End If
            Next
            If row.Item("OPS_YYYYPP") IsNot DBNull.Value Then
                row.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            End If

            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub

    Sub Setup_3PL()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        'ASCMAIN1.sql = "Select RCPTHDR.TRANS_SEQ, RCPTHDR.ARRDTE, RCPTHDR.PO_SHIPMENT_NO, RCPTHDR.CONTAINER_NO " _
        '    & ", RCPTDTL.ITEM_CODE, RCPTDTL.RCVQTY" _
        '    & " from ADS.RCPTHDR@ADSIIS,ADS.RCPTDTL@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
        '    & " and RCPTHDR.INVTYP = 'R'" _
        '    & "AND RCPTHDR.TRANS_SEQ = RCPTDTL.TRANS_SEQ " _
        '    & "AND RCPTHDR.LP_CODE = RCPTDTL.LP_CODE " _
        '    & "AND RCPTHDR.WHSE_CODE = RCPTDTL.WHSE_CODE "

        If Not dst.Tables.Contains("RCPTHDR") Then
            ASCMAIN1.sql = "Select RCPTHDR.*" _
                & " from ADS.RCPTHDR@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
                & " and RCPTHDR.INVTYP = 'R'"
            Create_TDA(dst.Tables.Add, "RCPTHDR", "**", 0, False, "", 0)
            'Dim RCPTHDR As DataTable = ASCDATA1.GetDataTable

            ASCMAIN1.sql = "Select RCPTDTL.*" _
                & " from ADS.RCPTHDR@ADSIIS,ADS.RCPTDTL@ADSIIS where RCPTHDR.STATUS in ('0','V')" _
                & " and RCPTHDR.INVTYP = 'R'" _
                & " and RCPTHDR.TRANS_SEQ = RCPTDTL.TRANS_SEQ " _
                & " and RCPTHDR.LP_CODE = RCPTDTL.LP_CODE " _
                & " and RCPTHDR.WHSE_CODE = RCPTDTL.WHSE_CODE "
            Create_TDA(dst.Tables.Add, "RCPTDTL", "**", 0, False, "", 0)
            dst.Tables("RCPTDTL").Columns.Add("STYLE_CODE")
            dst.Tables("RCPTDTL").Columns.Add("STYLE_DESC")
            dst.Tables("RCPTDTL").Columns.Add("COLOR_CODE")
            'Dim RCPTDTL As DataTable = ASCDATA1.GetDataTable
            dst.Relations.Add(dst.Tables("RCPTHDR").Columns("TRANS_SEQ"), dst.Tables("RCPTDTL").Columns("TRANS_SEQ"))
        End If

        EnforceConstraints(False)
        Fill_Records("RCPTHDR")
        Fill_Records("RCPTDTL")

        Dim RCPTDTL2 As DataTable = dst.Tables("RCPTDTL").Clone

        For Each rowRCPTDTL As DataRow In dst.Tables("RCPTDTL").Select("")
            Dim ITEM_CODE As String = rowRCPTDTL.Item("ITEM_CODE")
            If ITEM_CODE.EndsWith("PPK") Then
                ASCMAIN1.sql = "Select Sum (PPK_QTY) from WHTPPKM2 where PPK_CODE = '" & ITEM_CODE & "'"
                Dim PPK_QTY As Int64 = Val(ASCDATA1.GetDataValue)
                ASCMAIN1.sql = "Select * from WHTPPKM2 where PPK_CODE = '" & ITEM_CODE & "'"
                For Each rowWHTPPKM2 As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim rowRCPTDTL2 As DataRow = RCPTDTL2.NewRow
                    rowRCPTDTL2.ItemArray = rowRCPTDTL.ItemArray
                    rowRCPTDTL2.Item("STYLE_CODE") = rowWHTPPKM2.Item("STYLE_CODE")
                    rowRCPTDTL2.Item("COLOR_CODE") = rowWHTPPKM2.Item("COLOR_CODE")
                    Dim QTY As Int64 = Val(rowWHTPPKM2.Item("PPK_QTY") & "") * Val(rowRCPTDTL.Item("RCVQTY") & "") / PPK_QTY
                    rowRCPTDTL2.Item("RCVQTY") = QTY
                    RCPTDTL2.Rows.Add(rowRCPTDTL2)
                Next
                rowRCPTDTL.Delete()
            End If
        Next

        For Each row As DataRow In RCPTDTL2.Select("")
            dst.Tables("RCPTDTL").Rows.Add(row.ItemArray)
        Next


        EnforceConstraints(True)

        For Each row As DataRow In dst.Tables("RCPTDTL").Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
            If ITEM_CODE.EndsWith("PPK") Then
                Dim STYLE_CODE As String = row.Item("STYLE_CODE") & ""
                Dim rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    row.Item("STYLE_DESC") = rowICTSTYL1.ITEM("STYLE_DESC") & ""
                End If
            Else
                If Len(ITEM_CODE) > 3 Then
                    Dim STYLE_CODE As String = Mid(ITEM_CODE, 1, Len(ITEM_CODE) - 3)
                    Dim COLOR_CODE As String = Mid(ITEM_CODE, Len(ITEM_CODE) - 2, 3)
                    row.Item("STYLE_CODE") = STYLE_CODE
                    row.Item("COLOR_CODE") = COLOR_CODE
                    Dim rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
                    If rowICTSTYL1 IsNot Nothing Then
                        row.Item("STYLE_DESC") = rowICTSTYL1.ITEM("STYLE_DESC") & ""
                    End If
                End If
            End If
        Next
        grd3PL.DataSource = dst.Tables("RCPTHDR")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_SOTINVHH()
        If SELECTION_NO = 0 Then Exit Sub
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim YP As String = cbeInvoiceHistory.Value  ' ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3)
        Fill_Records("SOTINVHH", New String() {CUST_CODE, YP})
        Sort_grdColumns(grdSOTINVHH, "INV_NO".ToLower)
    End Sub

    Sub Load_SOTINVHX(STYLE_CODE As String)
        Dim YP As String = cbeInvoiceHistory.Value

        ' Use last 6 months for Regency
        If ASCMAIN1.CLIENT = "RGI" Then
            YP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6)
        End If

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Fill_Records("SOTINVHX", New String() {CUST_CODE, YP, STYLE_CODE})
        Sort_grdColumns(grdSOTINVHX, "INV_DATE".ToLower)
        grdSOTINVHX.Text = "Recent Sales of Style " & STYLE_CODE & " To " & Absx1.txtFor("CUST_CODE").Text
    End Sub

    Sub Setup_tab0()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Invoice History Since").Visible = (tab0.SelectedTab.Key = "Invoice History") And Not ScreenMode
        UltraExplorerBar1.Groups("Show if Entered in").Visible = (tab0.SelectedTab.Key = "Returns" Or tab0.SelectedTab.Key = "GL") And Not ScreenMode
    End Sub

    Private Sub ValidateOverCredited()

        If 1 = 1 Then
            Exit Sub
        End If

        If Not (ASCMAIN1.CLIENT = "RGI") Then
            Exit Sub
        End If

        Dim INV_NO_RETURNED As String = rowSOTRTRN1.Item("INV_NO_RETURNED") & String.Empty
        INV_NO_RETURNED = INV_NO_RETURNED.Trim

        If INV_NO_RETURNED.Length = 0 Then
            Exit Sub
        End If

        Dim sql As String = String.Empty
        sql = " SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_QTY_SHIP) ORDR_QTY_SHIP"
        sql &= " FROM"
        sql &= " ("
        sql &= " SELECT 'R', STYLE_CODE, COLOR_CODE, ORDR_QTY_SHIP FROM SOTINVH2 WHERE INV_NO IN"
        sql &= " ("
        sql &= " SELECT INV_NO FROM SOTRTRN1 WHERE INV_NO_RETURNED = '" & INV_NO_RETURNED & "'"
        sql &= " )"
        sql &= " UNION"
        sql &= " SELECT 'I', STYLE_CODE, COLOR_CODE, ORDR_QTY_SHIP FROM SOTINVH2 WHERE INV_TYPE = 'I' AND  INV_NO = '" & INV_NO_RETURNED & "'"
        sql &= " )"
        sql &= " GROUP BY STYLE_CODE, COLOR_CODE"

        Dim tblSOTINVH2 As DataTable = ASCDATA1.GetDataTable(sql)
        If tblSOTINVH2.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim qtyCredited As Int16 = 0
        Dim qtyAvail As Int16 = 0
        Dim STYLE_CODE As String = String.Empty
        Dim COLOR_CODE As String = String.Empty

        For Each rowSOTRTRN2 As DataRow In ASCDATA1.SelectDistinct("SOTRTRN2", New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
            STYLE_CODE = rowSOTRTRN2.Item("STYLE_CODE") & String.Empty
            COLOR_CODE = rowSOTRTRN2.Item("COLOR_CODE") & String.Empty

            qtyCredited = Val(dst.Tables("SOTRTRN2").Compute("SUM(RTRN_QTY)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)
            qtyAvail = Val(tblSOTINVH2.Compute("SUM(ORDR_QTY_SHIP)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)

            If qtyCredited > qtyAvail Then
                EMsg &= vbCr & "Style: " & STYLE_CODE & ",Color: " & COLOR_CODE & " has only " & qtyAvail & " pieces left to credit."
            End If
        Next

    End Sub

    Private Sub ValidateOverCreditedRMA()

        If dst.Tables("SOTRMAF2").Rows.Count = 0 Then
            Exit Sub
        End If

        If Not (ASCMAIN1.CLIENT = "RGI") Then
            Exit Sub
        End If

        Dim RA_NO As String = dst.Tables("SOTRMAF1").Rows(0).Item("RA_NO") & String.Empty

        Dim sql As String = String.Empty
        sql = " SELECT STYLE_CODE, COLOR_CODE, SUM(RA_QTY) QTY_AVAIL FROM"
        sql &= " ("
        sql &= "  SELECT 'R' TYPE_CODE, STYLE_CODE, COLOR_CODE, SUM(RA_QTY) RA_QTY"
        sql &= "  FROM SOTRMAF2 WHERE RA_NO = '" & RA_NO & "' GROUP BY STYLE_CODE, COLOR_CODE"
        sql &= "  UNION "
        sql &= "  SELECT 'S' TYPE_CODE, STYLE_CODE, COLOR_CODE, SUM(NVL(RTRN_QTY, 0) * -1) RA_QTY"
        sql &= "  FROM SOTRTRN1, SOTRTRN2 WHERE SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO AND SOTRTRN1.RA_NO = '" & RA_NO & "' GROUP BY STYLE_CODE, COLOR_CODE"
        sql &= "  )"
        sql &= " GROUP BY STYLE_CODE, COLOR_CODE"

        Dim tblSOTRMAFR As DataTable = ASCDATA1.GetDataTable(sql)
        If tblSOTRMAFR.Rows.Count = 0 Then
            Exit Sub
        End If

        Dim qtyCredited As Int16 = 0
        Dim qtyAvail As Int16 = 0
        Dim STYLE_CODE As String = String.Empty
        Dim COLOR_CODE As String = String.Empty

        For Each rowSOTRTRN2 As DataRow In ASCDATA1.SelectDistinct("SOTRTRN2", New String() {"STYLE_CODE", "COLOR_CODE"}).Rows
            STYLE_CODE = rowSOTRTRN2.Item("STYLE_CODE") & String.Empty
            COLOR_CODE = rowSOTRTRN2.Item("COLOR_CODE") & String.Empty

            qtyCredited = Val(dst.Tables("SOTRTRN2").Compute("SUM(RTRN_QTY)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)
            qtyAvail = Val(tblSOTRMAFR.Compute("SUM(QTY_AVAIL)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)

            If qtyCredited > qtyAvail Then
                EMsg &= vbCr & "Style: " & STYLE_CODE & ",Color: " & COLOR_CODE & " has only " & qtyAvail & " pieces left to credit."
            End If
        Next

    End Sub


    Private Sub CreateEDI180()

        If ASCMAIN1.CLIENT <> "NYA" Then
            Exit Sub
        End If

        Dim rowSOTRTRN1 As DataRow = dst.Tables("SOTRTRN1").Rows(0)
        If dst.Tables("SOTRTRN2").Rows.Count = 0 Then Exit Sub

        Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '180' AND CUST_CODE = :PARM1", "V", rowSOTRTRN1.Item("CUST_CODE"))
        If rowEDTTRPM1 Is Nothing Then
            Exit Sub
        End If

        Dim ORDR_NO As String = dst.Tables("SOTRTRN2").Rows(0).Item("ORDR_NO") & String.Empty
        Dim ORDR_LNO As Int16 = 0

        ASCMAIN1.sql = " Select" _
            & " O1.ORDR_NO, O1.ORDR_CUST_PO, T2.*" _
            & " from EDT850T1 T1, SOTORDR1 O1,  EDT850T2 T2" _
            & " WHERE T1.EDI_DOC_SEQ_NO = O1.EDI_DOC_SEQ_NO" _
            & " AND T1.EDI_DOC_SEQ_NO = T2.EDI_DOC_SEQ_NO" _
            & " AND O1.ORDR_NO = '" & ORDR_NO & "'"

        Dim tblEDT850T2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        If tblEDT850T2.Rows.Count = 0 Then
            Throw New Exception("Create EDI180 Error: EDT850T2 records not found for sales order: " & ORDR_NO)
        End If

        Dim EDI_DOC_SEQ_NO As String = tblEDT850T2.Rows(0).Item("EDI_DOC_SEQ_NO")
        Dim rowEDT850T1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
        If rowEDT850T1 Is Nothing Then
            Throw New Exception("Create EDI180 Error: EDT850T1 record not found for EDI Document: " & EDI_DOC_SEQ_NO)
        End If

        EDI_DOC_SEQ_NO = String.Empty

        Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")

        Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
        rowEDTSYSIH.Item("COMPANY_CODE") = ASCMAIN1.CLIENT
        rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
        rowEDTSYSIH.Item("EDI_APPLICATION_ID") = "AN"
        If rowEDTTRPM1.Item("EDI_STATUS") = "T" Then
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = "T"
        Else
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = "1"
        End If
        rowEDTSYSIH.Item("EDI_OUR_ID") = rowEDTTRPM1.Item("EDI_OUR_ID")
        rowEDTSYSIH.Item("EDI_TP_ID") = rowEDTTRPM1.Item("EDI_TP_ID")
        rowEDTSYSIH.Item("INIT_DATE") = DateTime.Now
        rowEDTSYSIH.Item("INIT_OPER") = ASCMAIN1.USER_ID
        dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

        Dim rowEDT180O1 As DataRow = dst.Tables("EDT180O1").NewRow
        rowEDT180O1.Item("COMPANY_CODE") = ASCMAIN1.CLIENT
        rowEDT180O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
        rowEDT180O1.Item("TRANS_PURPOSE_CODE") = "00" 'Original
        rowEDT180O1.Item("TRANS_REF_NO") = rowSOTRTRN1.Item("INV_NO") 'Our Credit Invoice Number
        rowEDT180O1.Item("TRANS_DATE") = rowSOTRTRN1.Item("RTRN_DATE")
        rowEDT180O1.Item("EDI_PO_NO") = rowEDT850T1.Item("EDI_PO_NO")
        rowEDT180O1.Item("EDI_SUPPLIER_NO") = rowEDT850T1.Item("EDI_SUPPLIER_NO")
        'rowEDT180O1.Item("TRANS_TYPE_CODE") = String.Empty
        'rowEDT180O1.Item("BATCH_NUMBER") = String.Empty
        'rowEDT180O1.Item("ENTITY_ID") = String.Empty
        'rowEDT180O1.Item("ENTITY_NAME") = String.Empty
        'rowEDT180O1.Item("ID_QUAL") = rowEDT850T1.Item("EDI_TP_QUAL") & String.Empty
        'rowEDT180O1.Item("ID_CODE") = rowEDT850T1.Item("EDI_TP_ID") & String.Empty
        rowEDT180O1.Item("RMA_NUMBER") = rowSOTRTRN1.Item("RTRN_NO") 'no RMA available
        rowEDT180O1.Item("CUST_CODE") = rowSOTRTRN1.Item("CUST_CODE")
        dst.Tables("EDT180O1").Rows.Add(rowEDT180O1)

        Dim EDI_DOC_LNO As Int32 = 0
        Dim rowICTSTYL1 As DataRow = Nothing
        For Each rowSOTRTRN2 As DataRow In dst.Tables("SOTRTRN2").Select("", "STYLE_CODE")

            ORDR_NO = rowSOTRTRN2.Item("ORDR_NO")
            ORDR_LNO = rowSOTRTRN2.Item("ORDR_LNO")
            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})

            Dim EDI_DTL_SEQ As Int16 = rowSOTORDR2.Item("EDI_DTL_SEQ")
            EDI_DOC_SEQ_NO = rowSOTORDR2.Item("EDI_DOC_SEQ_NO")

            Dim rowEDT850T2 As DataRow = tblEDT850T2.Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND EDI_DTL_SEQ = " & EDI_DTL_SEQ)(0)

            Dim rowEDT180O2 As DataRow = dst.Tables("EDT180O2").NewRow

            rowEDT180O2.Item("COMPANY_CODE") = rowEDT180O1.Item("COMPANY_CODE")
            rowEDT180O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO

            EDI_DOC_LNO += 1
            rowEDT180O2.Item("EDI_DOC_LNO") = EDI_DOC_LNO
            If rowICTSTYL1 Is Nothing OrElse rowICTSTYL1.Item("STYLE_CODE") <> rowSOTRTRN2.Item("STYLE_CODE") Then
                rowICTSTYL1 = LookUp("ICTSTYL1", rowSOTRTRN2.Item("STYLE_CODE"))
            End If

            rowEDT180O2.Item("EDI_QTY_RETURNED") = rowSOTRTRN2.Item("RTRN_QTY")
            rowEDT180O2.Item("EDI_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
            rowEDT180O2.Item("EDI_UNIT_PRICE") = rowEDT850T2.Item("EDI_PRICE")
            rowEDT180O2.Item("EDI_ITEM_UP") = rowEDT850T2.Item("EDI_UPC")
            rowEDT180O2.Item("EDI_ITEM_EN") = rowEDT850T2.Item("EDI_EAN")
            rowEDT180O2.Item("EDI_ITEM_GTIN") = rowEDT850T2.Item("EDI_GTIN")
            rowEDT180O2.Item("EDI_BUYER_ITEM") = rowEDT850T2.Item("EDI_SKU")
            rowEDT180O2.Item("EDI_SELLER_ITEM") = rowSOTORDR2.Item("STYLE_CODE")
            rowEDT180O2.Item("EDI_ITEM_DESC") = rowEDT850T2.Item("EDI_ITEM_DESC")
            rowEDT180O2.Item("EDI_BUYER_STYLE") = rowEDT850T2.Item("EDI_STYLE")
            rowEDT180O2.Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
            rowEDT180O2.Item("EDI_PO4_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
            rowEDT180O2.Item("EDI_PO4_INNER") = rowEDT850T2.Item("EDI_PO4_INNER")
            rowEDT180O2.Item("EDI_PO_LNO") = rowEDT850T2.Item("EDI_PO_LNO")
            'rowEDT180O2.Item("EDI_SIZE_CODE") = String.Empty
            rowEDT180O2.Item("EDI_SIZE_DESC") = rowEDT850T2.Item("EDI_SIZE_DESC")
            rowEDT180O2.Item("EDI_COLOR_CODE") = rowEDT850T2.Item("EDI_COLOR_CODE")
            rowEDT180O2.Item("EDI_COLOR_NAME") = rowEDT850T2.Item("EDI_COLOR_NAME")
            'rowEDT180O2.Item("EDI_DISPOSITION_CODE") = String.Empty
            'rowEDT180O2.Item("EDI_REQUEST_REASON_CODE") = String.Empty
            'rowEDT180O2.Item("EDI_RESPONSE_REASON_CODE") = String.Empty
            'rowEDT180O2.Item("EDI_RETURN_DESC") = String.Empty

            dst.Tables("EDT180O2").Rows.Add(rowEDT180O2)
        Next

    End Sub

    Public Function EmailInvoice(ByVal INV_TYPE As String, ByVal INV_NO As String) As Boolean

        Dim attachFileName As String = String.Empty

        Try

            If Not (ASCMAIN1.DBS_SERVER = "RGI" AndAlso ASCMAIN1.DBS_COMPANY = "RGI") Then
                Exit Function
            End If

            ' Only permit this to work in Production for regency.
            If Not (ASCMAIN1.DBS_SERVER = ASCMAIN1.DBS_COMPANY) Then
                Exit Function
            End If

            If dst.Tables("SOTINVH1").Select("INV_TYPE = '" & INV_TYPE & "' AND INV_NO = '" & INV_NO & "'").Length = 0 Then
                Return False
            End If

            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Select("INV_TYPE = '" & INV_TYPE & "' AND INV_NO = '" & INV_NO & "'")(0)
            Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE") & String.Empty
            Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty
            Dim SREP_CODE As String = rowSOTINVH1.Item("SREP_CODE") & String.Empty
            Dim salesRepEmail As String = String.Empty
            Dim CUST_XMIT_INV_VIA As String = String.Empty

            Dim CONS_INV As String = "0"
            If rowSOTINVH1.Item("INV_NO_CONS") & String.Empty <> String.Empty Then
                CONS_INV = "1"
            End If

            ' 10/10/2020 - As per Danny do not email House Account credits to the sales rep.
            If ASCMAIN1.CLIENT = "RGI" Then
                If SREP_CODE = "HO" Then
                    SREP_CODE = String.Empty
                End If
            End If

            ' See if the customer receives an acknowledgment
            Dim rowSOTSREP1 As DataRow = LookUp("SOTSREP1", SREP_CODE)
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

            Select Case ASCMAIN1.CLIENT
                Case "RGI"
                    If rowSOTSREP1 Is Nothing AndAlso rowARTCUST1 Is Nothing Then
                        Return False
                    End If

                Case Else
                    If rowARTCUST1 Is Nothing Then
                        Return False
                    End If
            End Select

            ' See if we have anyone to email to - Only RGI sends a copy of the invoice to the sales rep
            If ASCMAIN1.CLIENT = "RGI" Then
                If rowSOTSREP1 IsNot Nothing AndAlso rowSOTSREP1.Item("SREP_EMAIL") & String.Empty <> String.Empty Then
                    salesRepEmail = rowSOTSREP1.Item("SREP_EMAIL") & String.Empty
                End If
            Else
                salesRepEmail = String.Empty
            End If

            If rowARTCUST1 IsNot Nothing Then
                ' Mail, Email, Both
                CUST_XMIT_INV_VIA = (rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty).ToString.Trim
                If CUST_XMIT_INV_VIA.Length > 0 AndAlso "EB".Contains(CUST_XMIT_INV_VIA) Then
                    If rowARTCUST1.Item("CUST_INV_EMAIL") & String.Empty <> String.Empty Then
                        salesRepEmail &= ";" & rowARTCUST1.Item("CUST_INV_EMAIL") & String.Empty
                    End If

                    If rowARTCUST1.Item("CUST_INV_CC") & String.Empty <> String.Empty Then
                        salesRepEmail &= ";" & rowARTCUST1.Item("CUST_INV_CC") & String.Empty
                    End If
                End If
            End If

            ' remove double semi-colons
            salesRepEmail = salesRepEmail.Replace(",", ";")
            salesRepEmail = salesRepEmail.Replace(";;", ";")
            salesRepEmail = salesRepEmail.Replace(" ", "")

            ' should be at least 5 characters
            If salesRepEmail.Replace(";", "").Trim.Length < 5 Then
                Return False
            End If

            attachFileName = rowARTCUST1.Item("CUST_NAME") & " " & INV_NO

            For Each invalidChar As String In New String() {"\", "/", ":", "*", "?", "<", ">", "|", "."}
                attachFileName = attachFileName.Replace(invalidChar, "")
            Next
            attachFileName = attachFileName.Replace(" ", "_")

            Dim invNos As String = String.Empty
            For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                invNos &= ", '" & row.Item("INV_NO") & "'"
            Next
            invNos = invNos.Substring(1).Trim

            ASCMAIN1.Progress("Emailing Invoice", "")

            Dim invoiceLocation As String = String.Empty
            If SO_PARM_WEB_INVOICES.Length > 0 Then
                invoiceLocation = SO_PARM_WEB_INVOICES & CUST_CODE & "\" & INV_NO & ".pdf"
                If Not My.Computer.FileSystem.FileExists(invoiceLocation) Then
                    invoiceLocation = String.Empty
                End If
            End If

            Dim ATTACHMENTs As New Dictionary(Of String, String)

            If invoiceLocation.Length = 0 Then
                Dim RPT As String = "SORINVP1"
                If Not REPORTS.ContainsKey(RPT) Then
                    REPORTS.Add(RPT, Load_rptClass(RPT))
                    REPORTS(RPT).Prepare_dst(False, "")
                End If

                REPORTS(RPT).Fill_Records_RPT(New String() {" and SOTINVH1.INV_NO IN (" & invNos & ")"})

                Dim REPORT_NO As String = String.Empty
                With REPORTS(RPT).clsASCBASE1
                    .Print_Report_Begin()
                    .CR_params.Add("SUBT", "")
                    .CR_params.Add("CONS_INV", CONS_INV)
                    .CR_params.Add("EXPORT_INFO", "0")

                    ' Set the customers Invoice
                    Select Case ASCMAIN1.CLIENT
                        Case "RGI"
                            RPT = "SORINVPR"

                    End Select

                    REPORT_NO = .Generate_Report(RPT, "Invoice", , True, , , "PDF", attachFileName, False)
                    .Print_Report_End(True, True)
                End With
                ATTACHMENTs.Add(attachFileName & ".pdf", ASCMAIN1.Folders("Temp") & attachFileName & ".pdf")
            Else
                ATTACHMENTs.Add("Invoice: " & INV_NO, invoiceLocation)
            End If

            Dim SUBJECT As String = String.Empty
            SUBJECT = "Sales Invoice (" & INV_NO & ") for customer " & rowARTCUST1.Item("CUST_NAME")

            ' Concatentate and process all email addresses
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each emailAddress As String In (salesRepEmail).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Return True
            End If

            Dim EMAIL_KEY As String = "INV"
            Select Case ASCMAIN1.CLIENT
                Case "RGI"
                    EMAIL_KEY = "AUTOINV"
            End Select

            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    SUBJECT, EMAIL_KEY,
                    True, False, CUST_CODE, rowARTCUST1.Item("CUST_NAME"), "Customer")


            ' Mark email Only Invoices as Mailed
            Try
                If CUST_XMIT_INV_VIA = "E" Then
                    For Each rowSOTINVH1 In dst.Tables("SOTINVH1").Rows
                        INV_NO = rowSOTINVH1.Item("INV_NO")
                        ASCDATA1.ExecuteSQL("Update SOTINVH1 Set INV_PRINTED = SYSDATE where INV_NO = '" & INV_NO & "'")
                    Next
                End If
            Catch ex As Exception
                ' nothing 
            End Try

            EmailInvoice = True

        Catch ex As Exception
            EmailInvoice = False
        End Try

    End Function

    Private Sub txtTrackingNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtTrackingNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            LoopReturnID = String.Empty
            Dim TRACKING_NUMBER As String = txtTrackingNo.Text.Trim

            If TRACKING_NUMBER.Length = 0 Then
                Exit Sub
            End If

            LookUpReturn(RefundLookupTypes.TrackingNumber, TRACKING_NUMBER)
        End If
    End Sub

    Private Sub txtReturnId_KeyDown(sender As Object, e As KeyEventArgs) Handles txtReturnId.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            LoopReturnID = String.Empty
            Dim RETURN_ID As String = txtReturnId.Text.Trim

            If RETURN_ID.Length = 0 Then
                Exit Sub
            End If

            LookUpReturn(RefundLookupTypes.ReturnID, RETURN_ID)
        End If
    End Sub

    Private Enum RefundLookupTypes
        ReturnID
        TrackingNumber
    End Enum

    Private Sub LookUpReturn(lookupType As RefundLookupTypes, inValue As String)

        Try
            LoopReturnID = String.Empty
            Select Case lookupType
                Case RefundLookupTypes.ReturnID
                    ASCMAIN1.sql = "SELECT RETURN_ID FROM SOTRTNL1 WHERE RETURN_ID = :PARM1"

                Case RefundLookupTypes.TrackingNumber
                    ASCMAIN1.sql = "SELECT MAX(RETURN_ID) RETURN_ID FROM SOTRTNL3 WHERE :PARM1 LIKE '%' ||TRACKING_NUMBER"
            End Select

            Dim RETURN_ID As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", {inValue}) & String.Empty

            If RETURN_ID.Length = 0 Then
                MessageBox.Show("Cannot locate the Web Return.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim drSOTRTNL1 As DataRow = LookUp("SOTRTNL1", RETURN_ID)
            If drSOTRTNL1 Is Nothing Then
                MessageBox.Show("Cannot locate the Web Return.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If drSOTRTNL1.Item("PROCESS_IND") & String.Empty = "1" Then
                MessageBox.Show($"Web Return {RETURN_ID} was already processed.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTRTNL3 WHERE RETURN_ID = :PARM1", "", "V", {RETURN_ID})
            If tbl.Rows.Count > 1 Then
                MessageBox.Show($"This return should have {tbl.Rows.Count} cartons.", "Search Tracking No", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

            LoopReturnID = RETURN_ID
            Dim ORDR_NO As String = drSOTRTNL1.Item("ORDR_NO") & String.Empty

            Dim drECTECOMD As DataRow = dst.Tables("ECTECOMD").Rows.Find("SHOPIFY")
            If drECTECOMD Is Nothing Then
                MessageBox.Show($"Cannot locate attributes for Ecommerce Partner SHOPIFY", "Process Return", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Absx1.txtFor("CUST_CODE").Text = drECTECOMD.Item("ECOM_CUST_CODE") & String.Empty
            Absx1.txtFor("WHSE_CODE").Text = drECTECOMD.Item("ECOM_WHSE_CODE") & String.Empty

            If IsDate(drSOTRTNL1.Item("CREATED_AT") & String.Empty) Then
                Absx1.dteFor("RTRN_DATE").Value = CDate(drSOTRTNL1.Item("CREATED_AT") & String.Empty).ToShortDateString
            Else
                Absx1.dteFor("RTRN_DATE").Value = DateTime.Now.ToShortDateString
            End If
            LoopReturnID = RETURN_ID

            Click_Command("New")
            If LoopReturnID.Length > 0 Then
                Absx1.txtFor("CUST_STORE_NO").Value = drECTECOMD.Item("ECOM_CUST_ADDR_CODE") & String.Empty
                'If ORDR_NO.Length > 0 Then
                '    Absx1.txtFor("RTRN_NOTE").Value = $"Sales Order No: {ORDR_NO}"
                'End If
                Absx1.txtFor("REASON_CODE").Value = "RTN"
                DisplayTotals()
                grdSOTRTRN2.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            End If

        Catch ex As Exception
            MessageBox.Show($"Error {ex.Message}", "Locate Web Refund", MessageBoxButtons.OK, MessageBoxIcon.Error)
            LoopReturnID = String.Empty
        Finally
            txtTrackingNo.Clear()
            txtReturnId.Clear()
        End Try

    End Sub

#End Region


End Class