Public Class TATPRTN1

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            grpExtracts.Visible = True
            Set_Read_Only(grpExtracts, False)
        Else
            grpExtracts.Visible = False
        End If
        grpExtracts.Visible = True
        Set_Read_Only(grpExtracts, False)
        'btnFix.Visible = tf
        'If tf And ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" ThenS
        '    btnFix.Visible = True
        'End If
    End Sub
    Public Overrides Sub Show_Record_Special()
        '  Set_Read_Only(grpExtracts, False)
    End Sub
    Private Sub UltraButton1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub btnOpenOrders_Click(sender As Object, e As EventArgs) Handles btnOpenOrders.Click


        Dim PARTNER_CODE As String = "EFNY"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim worksheet1 As SpreadsheetGear.IWorksheet = Nothing

        Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
        Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")


        dst.Tables("APTVVVV1").Rows.Clear()
        ' Fill_Records("APTVVVV1", PARTNER_CODE
        Fill_Records("APTVVVV1", New Object() {dteStart.Value, dteEnd.Value})

        Dim DT As DataTable = dst.Tables("APTVVVV1")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("TATPRTN1.EXTRACT_NO") & "-" & PARTNER_CODE & ".xlsX" '  ASCMAIN1.Folders("Archive") & "\Frame Supplier Account Info.xlsX"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")


        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If

        Next


        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        range = oSheet.Cells("A1:AY1")
        range.Interior.Color = SpreadsheetGear.Colors.Beige
        '  range.EntireColumn.AutoFilter()
        range.EntireColumn.AutoFit()

        'oSheet.WindowInfo.ScrollColumn = 0
        'oSheet.WindowInfo.SplitColumns = 1

        ' Split after row 2 (ScrollRow Is zero based).  
        oSheet.WindowInfo.ScrollRow = 0
        oSheet.WindowInfo.SplitRows = 0

        ' Freeze the panes. 
        oSheet.WindowInfo.FreezePanes = True

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(xls_filename)
        oWB.Close()

    End Sub

    Private Sub dteDueDate_ValueChanged(sender As Object, e As EventArgs) Handles dteEnd.ValueChanged

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub TATPRTN1_Load(sender As Object, e As EventArgs) Handles Me.Load
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            With dst
                '     Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
                '     Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")

                ASCMAIN1.sql = "Select 'NYAG' COMPANY, ICTSTYL1.SALES_DIVISION_CODE, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_NO" _
                    & " , SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.TERM_CODE, SOTORDR1.ORDR_MESSAGE" _
                    & " , SOTORDR1.WHSE_CODE, SOTORDR1.SREP_CODE, SOTORDR1.ORDR_CUST_PO, NULL COL014, NULL COL015, SOTORDR1.ORDR_SHIP_INSTR" _
                    & " , SOTORDR1.ORDR_INV_COMMENT, NULL COL018, NULL COL019, NULL COL020" _
                    & " , SOTORDR1.ORDR_TYPE_CODE, ICTSEAS1.SEASON_TYPE, ICTSEAS1.SEASON_YEAR" _
                    & " , SOTORDR2.STYLE_CODE, NULL SKU200, NULL SKU300, NULL SKU400, NULL COL028" _
                    & " , NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) QTY, SOTORDR2.ORDR_UNIT_PRICE" _
                    & " , DECODE(SOTORDR1.CURR_CODE, 'CAD', SOTORDR2.ORDR_UNIT_PRICE_CURR, SOTORDR2.ORDR_UNIT_PRICE) ORDR_UNIT_PRICE_CURR" _
                    & " , NULL COL032, TO_CHAR(SOTORDR1.ORDR_DATE,'MM') MONTH, TO_CHAR(SOTORDR1.ORDR_DATE,'YYYY') YEAR" _
                    & " , NULL COL035, NULL COL036" _
                    & " , DECODE(SOTAUTH1.ORDR_CRED_CLR_AUTH,NULL,NULL,SOTAUTH1.ORDR_CRED_CLR_AUTH || '-' || SOTAUTH1.ORDR_CRED_CLR_BY) CREDIT_APPR" _
                    & " , SOTORDR1.FRT_TERMS, NULL COL039, NULL COL040" _
                    & " , NULL MISC, NULL FRT, NULL STAX, NULL COL044, NULL COL045, NULL COL046, SOTORDR1.ORDR_SOURCE" _
                    & " , NVL(NVL(SOTORDR2.CUST_SKU,SOTORDR2.CUST_STYLE_CODE),SOTORDR2.CUST_UPC) CUST_STYLE" _
                    & " , SOTORDR1.CUST_FACTOR_IND , SOTORDR1.ORDR_DEPT, SOTORDR1.ORDR_STATUS" _
                    & " From SOTORDR1, ICTSTYL1, ICTSEAS1, ICTSTYC1, SOTAUTH1, SOTORDR2" _
                    & " Where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE And ICTSEAS1.SEASON_CODE(+) = ICTSTYL1.SEASON_CODE" _
                    & " And ICTSTYC1.STYLE_CODE = SOTORDR2.STYLE_CODE" _
                    & " And ICTSTYC1.COLOR_CODE = SOTORDR2.COLOR_CODE" _
                    & " And SOTAUTH1.ORDR_GROUP_NO (+) = SOTORDR1.ORDR_GROUP_NO" _
                    & " And SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                    & " And SOTORDR1.ORDR_STATUS IN ('O','P','F')" _
                    & " And SOTORDR1.INIT_DATE BETWEEN :PARM1 and :PARM2" _
                    & " ORDER BY SOTORDR1.ORDR_NO"
                Create_TDA(.Tables.Add, "APTVVVV1", "**", 0, False, "DD")


                ASCMAIN1.sql = "Select  ICTSTYL1.SALES_DIVISION_CODE, POTORDR1.FACTORY_CODE, POTORDR1.PO_ORDER_NO" _
                   & " , POTORDR1.PO_REVISION_NOTE, ICTSTYL1.SEASON_CODE, POTORDR2.STYLE_CODE, NULL SKU200" _
                   & " , NULL SKU300, NULL SKU400, NULL SIZE_SCALE, NULL SIZECODE, POTORDR2.PO_QTY_OPN" _
                   & " , POTORDR1.WHSE_CODE, POTORDR2.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_CANCEL" _
                   & " , POTORDR1.PO_SHIP_VIA, POTORDR2.PO_COST,             POTORDR2.PO_LINE_NOTE_INT" _
                   & " , NULL MARKAS, NULL FCTCLR, POTORDR1.VEND_CODE" _
                   & " , POTORDR1.CUST_CODE, ARTCUST1.SREP_CODE, SOTORDR1.ORDR_SHIP_DATE" _
                   & " , SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_CUST_PO, POTORDR1.PORT_CODE_ORIG" _
                   & " , NULL PORT_OF_ARRIVAL, NULL FREIGHT_FORWARDER" _
                   & " , POTORDR1.FOB_CMT, NULL SUBDIV, POTORDR1.TERM_CODE TERM_PO, POTORDR1.PO_MESSAGE" _
                   & " , ' ' BUSUNIT, ' 'BUM, ' 'PRODMGR, ' 'PRODSPC" _
                   & " , ' ' SRCMGR, ' ' SRCSPC" _
                   & " , ' ' POTEMP, ' ' POCYCLE, ' ' CUBIC, ' ' GRSWGT, ' ' PIECES, ' ' CTNS" _
                   & " , ' ' RDYDTE, ' ' OPSMGR, ' ' OPSSPC, POTORDR1.PO_COMM_PCT" _
                   & " , ' ' POCOMMINCL, ' ' ROYALTYPCT, ' ' ROYALTYINCL, SOTORDR1.TERM_CODE" _
                   & " , ' ' TCR, POTORDR2.PO_DATE_ETA DELSCHED, ' ' ACTDELIV,POTORDR1.PO_DATE_ORDERED" _
                   & " From POTORDR1, POTORDR2, ICTSTYL1, SOTORDR1, ARTCUST1" _
                   & " Where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                   & " And SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO" _
                   & " And ARTCUST1.CUST_CODE (+) = SOTORDR1.CUST_CODE" _
                   & " And ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" _
                   & " And POTORDR2.PO_QTY_OPN <> 0" _
                   & " And POTORDR1.PO_DATE_ORDERED BETWEEN :PARM1 and :PARM2"
                Create_TDA(.Tables.Add, "APTVVVV2", "**", 0, False, "DD")




                ASCMAIN1.sql = "Select 'NYAG' COMPANY, ICTSTYL1.SALES_DIVISION_CODE, ICTSEAS1.SEASON_TYPE, ICTSEAS1.SEASON_YEAR" _
                    & " , ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, SUBSTR(ICTSTYL1.SIZE_SCALE,1,5) SIZE_SCALE, NULL SKU200, NULL SKU300" _
                    & " , NULL SKU400, NULL COLOR_DESC, NULL SIZE_SCALE1, NULL SIZES, NULL SIZE_INDEX, NULL UPC_INFO, NULL UPC_MFR, NULL UPC_COL, NULL UPC_CD, ICTSTYL1.STYLE_GROUP_CODE" _
                    & " , ICTSGRP1.STYLE_GROUP_DESC, ICTSTYL1.STYLE_CLASS_CODE CATGY_CODE, ICTCLAS1.STYLE_CLASS_DESC CATGY_DESC" _
                    & " , ICTSTYL1.STYLE_WEIGHT, DECODE(DGJPRICE.ORDR_UNIT_PRICE,NULL,ICTSTYL1.STYLE_PRICE,DGJPRICE.ORDR_UNIT_PRICE) AAA3" _
                    & " , ICTSTYL1.STYLE_RETAIL, NULL QRS_SELCODE, NULL GENDER, NULL SUBDIV, ICTSTYL1.ROYALTY_CODE" _
                    & " , ICTSEAS1.SEASON_TYPE DELPERIOD, DECODE(ICTSTYV1.PO_COST  ,NULL,DGJPOCOST.PO_COST,ICTSTYV1.PO_COST ) FACTORY_COST, ICTSTYL1.DUTY_RATE_CODE, NULL HTSYEAR, NULL PRODTYPE, ICTSTYC1.UPC_CODE" _
                    & " , ICTSTYL1.STYLE_DESC2 USERDEF01, NULL USERDEF02, NULL SUBCAT, NULL GARMENTTYPE, NULL DESC2, NULL MATL, 'P' PACKORHANG, ICTSTYL1.STYLE_MATL_DESC" _
                    & " , NULL SKUNULL, NULL COLLECTION, NULL CUSTDEF01, NULL CUSTDEF02, NULL CUSTDEF03, NULL CUSTDEF04, NULL CUSTDEF05, NULL CUSTDEF06" _
                    & " , NULL STYLEXREF, NULL SELLPRD, NULL SELLYEAR, ICTSTYL1.COUNTRY_CODE COO, DECODE(ICTSTYL1.VEND_CODE,NULL,DGJPOCOST.VEND_CODE,ICTSTYL1.VEND_CODE) VEND_CODE, ICTSTYL1.STYLE_STATUS, ICTSTYC1.NRF_COLOR_CODE, NULL CUSTOMSDESC" _
                    & " , ICTSTYC1.UPC_CODE_CASE GTIN, NULL FORLANG, NULL CASELEN, NULL CASEWID, NULL CASEHGT, NULL REF5, NULL UPCXREFLONG, NULL BODYTYPE, NULL DESIGNER, NULL CYCLE" _
                    & " , ICTSTYL1.CUST_CODE NYAGCUSTCODE, ICTSTYL1.CUST_STYLE_CODE  NYAGCUSTOMERSTYLECODE, ICTSTYL1.STYLE_CODE_PLM PLMSTYLE, ICTPLIN2.DESIGN_STYLE_NO" _
                    & " , ICTSTYL1.CARTON_PACK_QTY NYAGUNITSCASE, ICTSTYL1.INNER_PACK_QTY NYAGUNITSINNER, ICTSTYC1.UPC_CODE_INNER NYAGUPCINNER,DGJPOCOST.STYLE_CODE STYLE_CODE1, DGJPOCOST.VEND_CODE VEND_CODE1, DGJPOCOST.PO_COST,ICTSTYL1.INIT_DATE,ICTSTYL1.SIZE_SCALE SIZE_SCALE_FULL" _
                    & " From ICTSTYL1, ICTSEAS1, ICTSTYC1, ICTSGRP1, ICTCLAS1, ICTPLIN2, ICTSTYV1, DGJPRICE, DGJPOCOST" _
                    & " Where ICTSEAS1.SEASON_CODE(+) = ICTSTYL1.SEASON_CODE" _
                    & " And ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE And ICTSTYC1.COLOR_CODE = 'AST'" _
                    & " And ICTSGRP1.STYLE_GROUP_CODE (+) = ICTSTYL1.STYLE_GROUP_CODE" _
                    & " And ICTCLAS1.STYLE_CLASS_CODE (+) = ICTSTYL1.STYLE_CLASS_CODE" _
                    & " And ICTPLIN2.STYLE_CODE_PLM (+) = ICTSTYL1.STYLE_CODE_PLM" _
                    & " And ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" _
                    & " And ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" _
                    & " And DGJPRICE.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" _
                    & " And DGJPOCOST.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" _
                    & " And ICTSTYL1.INIT_DATE  BETWEEN :PARM1 and :PARM2" _
                    & " And ICTSTYL1.STYLE_CODE IN" _
                    & " (SELECT DISTINCT STYLE_CODE FROM ICTSTAT2 WHERE (NVL(WHSE_QTY_ON_HAND,0) <> 0 Or NVL(WHSE_QTY_ON_ORDER,0) <> 0 Or NVL(WHSE_QTY_TRAN,0) <> 0))" _
                    & " ORDER BY ICTSTYL1.STYLE_CODE"
                Create_TDA(.Tables.Add, "APTVVVV3", "**", 0, False, "DD")


                ASCMAIN1.sql = "Select 'NYAG' COMPANY, ICTSTYL1.SALES_DIVISION_CODE, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_NO" _
                    & " , SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTINVH1.TERM_CODE, SOTORDR1.ORDR_MESSAGE" _
                    & " , SOTINVH1.WHSE_CODE, SOTINVH1.SREP_CODE, SOTINVH1.ORDR_CUST_PO, NULL COL014, NULL COL015, SOTORDR1.ORDR_SHIP_INSTR" _
                    & " , SOTORDR1.ORDR_INV_COMMENT, NULL COL018, NULL COL019, NULL COL020, SOTORDR1.ORDR_TYPE_CODE, ICTSEAS1.SEASON_TYPE, ICTSEAS1.SEASON_YEAR" _
                    & " , SOTINVH2.STYLE_CODE, NULL SKU200, NULL SKU300, NULL SKU400, NULL COL028, SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
                    & " , DECODE(SOTINVH1.CURR_CODE, 'CAD', SOTINVH2.ORDR_UNIT_PRICE_CURR, SOTINVH2.ORDR_UNIT_PRICE) ORDR_UNIT_PRICE_CURR" _
                    & " , NULL COL032, SUBSTR(SOTINVH1.ORDR_YYYYPP_UPDATED,5,2) MONTH, SUBSTR(SOTINVH1.ORDR_YYYYPP_UPDATED,1,4) YEAR" _
                    & " , NULL COL035, NULL COL036" _
                    & " , DECODE(SOTAUTH1.ORDR_CRED_CLR_AUTH,NULL,NULL,SOTAUTH1.ORDR_CRED_CLR_AUTH || '-' || SOTAUTH1.ORDR_CRED_CLR_BY) CREDIT_APPR" _
                    & " , SOTORDR1.FRT_TERMS, NULL COL039, NULL COL040" _
                    & " , SOTINVH1.INV_MISC_CHG, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_STAX, NULL COL044, NULL COL045, NULL COL046, SOTORDR1.ORDR_SOURCE" _
                    & " , NVL(NVL(SOTORDR2.CUST_SKU,SOTORDR2.CUST_STYLE_CODE),SOTORDR2.CUST_UPC) CUST_STYLE" _
                    & " , SOTORDR1.CUST_FACTOR_IND , SOTORDR1.ORDR_DEPT,SOTINVH1.INV_NO,SOTINVH1.INV_date" _
                    & " From SOTINVH1, SOTORDR1, ICTSTYL1, ICTSEAS1, ICTSTYC1, SOTINVH2, SOTAUTH1, SOTORDR2" _
                    & " Where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE And ICTSEAS1.SEASON_CODE(+) = ICTSTYL1.SEASON_CODE" _
                    & " And ICTSTYC1.STYLE_CODE = SOTINVH2.STYLE_CODE" _
                    & " And ICTSTYC1.COLOR_CODE = SOTINVH2.COLOR_CODE" _
                    & " And SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" _
                    & " And SOTINVH2.INV_NO = SOTINVH1.INV_NO" _
                    & " And SOTAUTH1.ORDR_GROUP_NO (+) = SOTORDR1.ORDR_GROUP_NO" _
                    & " And SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" _
                    & " And SOTORDR2.ORDR_NO = SOTINVH1.ORDR_NO" _
                    & " And SOTORDR2.ORDR_LNO = SOTINVH2.INV_LNO" _
                    & " And SOTINVH1.INV_DATE BETWEEN :PARM1 and :PARM2" _
                    & " ORDER BY SOTINVH1.INV_NO"
                Create_TDA(.Tables.Add, "APTVVVV4", "**", 0, False, "DD")



                ASCMAIN1.sql = "Select * From POTORDR6 Where (POTORDR6.PO_ORDER_NO, POTORDR6.PO_ORDER_LNO) In (" _
                    & " Select  DISTINCT POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO" _
                    & " From POTORDR1, POTORDR2, ICTSTYL1, SOTORDR1, ARTCUST1" _
                    & " Where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                    & " And SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO" _
                    & " And ARTCUST1.CUST_CODE (+) = SOTORDR1.CUST_CODE" _
                    & " And ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" _
                    & " And POTORDR2.PO_QTY_OPN <> 0" _
                    & " And POTORDR1.PO_DATE_ORDERED BETWEEN :PARM1 and :PARM2" & ")"
                Create_TDA(.Tables.Add, "APTVVVV5", "**", 0, False, "DD")

                ASCMAIN1.sql = "Select ICTSTYL1.SALES_DIVISION_CODE, POTORDR1.FACTORY_CODE, POTORDR1.PO_ORDER_NO" _
                    & " , POTORDR1.PO_REVISION_NOTE, ICTSTYL1.SEASON_CODE, POTORDR2.STYLE_CODE, NULL SKU200" _
                    & " , NULL SKU300, NULL SKU400, NULL SIZE_SCALE, NULL SIZECODE, POTORDR2.PO_QTY_SHP In_Transit" _
                    & " , POTORDR1.WHSE_CODE, POTORDR2.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_CANCEL" _
                    & " , POTORDR1.PO_SHIP_VIA, POTORDR2.PO_COST,             POTORDR2.PO_LINE_NOTE_INT" _
                    & " , NULL MARKAS, NULL FCTCLR, POTORDR1.VEND_CODE" _
                    & " , POTORDR1.CUST_CODE, ARTCUST1.SREP_CODE, SOTORDR1.ORDR_SHIP_DATE" _
                    & " , SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_CUST_PO, POTORDR1.PORT_CODE_ORIG" _
                    & " , NULL PORT_OF_ARRIVAL, NULL FREIGHT_FORWARDER" _
                    & " , POTORDR1.FOB_CMT, NULL SUBDIV, POTORDR1.TERM_CODE TERM_PO, POTORDR1.PO_MESSAGE" _
                    & " , ' ' BUSUNIT, ' 'BUM, ' 'PRODMGR, ' 'PRODSPC" _
                    & " , ' ' SRCMGR, ' ' SRCSPC" _
                    & " , ' ' POTEMP, ' ' POCYCLE, ' ' CUBIC, ' ' GRSWGT, ' ' PIECES, ' ' CTNS" _
                    & " , ' ' RDYDTE, ' ' OPSMGR, ' ' OPSSPC, POTORDR1.PO_COMM_PCT" _
                    & " , ' ' POCOMMINCL, ' ' ROYALTYPCT, ' ' ROYALTYINCL, SOTORDR1.TERM_CODE" _
                    & " , ' ' TCR, POTORDR2.PO_DATE_ETA DELSCHED, ' ' ACTDELIV" _
                    & " From POTORDR1, POTORDR2, ICTSTYL1, SOTORDR1, ARTCUST1" _
                    & " Where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                    & " And SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO" _
                    & " And ARTCUST1.CUST_CODE (+) = SOTORDR1.CUST_CODE" _
                    & " And ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" _
                    & " And (POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO) IN (" _
                    & " Select PO_ORDER_NO,PO_ORDER_LNO FROM POTSHIP3 WHERE PO_SHIPMENT_NO In (" _
                    & " Select PO_SHIPMENT_NO from  POTSHIp2 WHERE PO_SHIP_STATUS = 'O')" _
                    & " )"
                Create_TDA(.Tables.Add, "APTVVVV6", "**", 0, False, "DD")


                ASCMAIN1.sql = "Select ICTSTYL1.SALES_DIVISION_CODE, POTORDR1.FACTORY_CODE, POTORDR1.PO_ORDER_NO" _
                    & " , POTORDR1.PO_REVISION_NOTE, ICTSTYL1.SEASON_CODE, POTORDR2.STYLE_CODE, NULL SKU200" _
                    & " , NULL SKU300, NULL SKU400, NULL SIZE_SCALE, NULL SIZECODE, POTORDR2.PO_QTY_REC Received" _
                    & " , POTORDR1.WHSE_CODE, POTORDR2.PO_DATE_SHIP_BY, POTORDR1.PO_DATE_CANCEL" _
                    & " , POTORDR1.PO_SHIP_VIA, POTORDR2.PO_COST,             POTORDR2.PO_LINE_NOTE_INT" _
                    & " , NULL MARKAS, NULL FCTCLR, POTORDR1.VEND_CODE" _
                    & " , POTORDR1.CUST_CODE, ARTCUST1.SREP_CODE, SOTORDR1.ORDR_SHIP_DATE" _
                    & " , SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_CUST_PO, POTORDR1.PORT_CODE_ORIG" _
                    & " , NULL PORT_OF_ARRIVAL, NULL FREIGHT_FORWARDER" _
                    & " , POTORDR1.FOB_CMT, NULL SUBDIV, POTORDR1.TERM_CODE TERM_PO, POTORDR1.PO_MESSAGE" _
                    & " , ' ' BUSUNIT, ' 'BUM, ' 'PRODMGR, ' 'PRODSPC" _
                    & " , ' ' SRCMGR, ' ' SRCSPC" _
                    & " , ' ' POTEMP, ' ' POCYCLE, ' ' CUBIC, ' ' GRSWGT, ' ' PIECES, ' ' CTNS" _
                    & " , ' ' RDYDTE, ' ' OPSMGR, ' ' OPSSPC, POTORDR1.PO_COMM_PCT" _
                    & " , ' ' POCOMMINCL, ' ' ROYALTYPCT, ' ' ROYALTYINCL, SOTORDR1.TERM_CODE" _
                    & " , ' ' TCR, POTORDR2.PO_DATE_ETA DELSCHED, ' ' ACTDELIV" _
                    & " From POTORDR1, POTORDR2, ICTSTYL1, SOTORDR1, ARTCUST1" _
                    & " Where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
                    & " And SOTORDR1.ORDR_NO (+) = POTORDR1.ORDR_NO" _
                    & " And ARTCUST1.CUST_CODE (+) = SOTORDR1.CUST_CODE" _
                    & " And ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" _
                    & " And (POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO) IN (" _
                    & " Select PO_ORDER_NO,PO_ORDER_LNO FROM POTSHIP3 WHERE PO_SHIPMENT_NO In (" _
                    & " Select PO_SHIPMENT_NO from  POTSHIp2 WHERE PO_SHIP_STATUS = 'C')" _
                    & " )"
                Create_TDA(.Tables.Add, "APTVVVV7", "**", 0, False, "DD")

            End With
        End If

    End Sub

    Private Sub UltraButton1_Click_1(sender As Object, e As EventArgs) Handles UltraButton1.Click
        Dim PARTNER_CODE As String = "EFNY"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim worksheet1 As SpreadsheetGear.IWorksheet = Nothing

        Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
        Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")


        dst.Tables("APTVVVV2").Rows.Clear()
        ' Fill_Records("APTVVVV1", PARTNER_CODE
        Fill_Records("APTVVVV2", New Object() {dteStart.Value, dteEnd.Value})

        Dim DT As DataTable = dst.Tables("APTVVVV2")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("TATPRTN1.EXTRACT_NO") & "-" & PARTNER_CODE & ".xlsX" '  ASCMAIN1.Folders("Archive") & "\Frame Supplier Account Info.xlsX"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")


        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If

        Next

        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        range = oSheet.Cells("A1:BF1")
        range.Interior.Color = SpreadsheetGear.Colors.Beige
        '  range.EntireColumn.AutoFilter()
        range.EntireColumn.AutoFit()

        'oSheet.WindowInfo.ScrollColumn = 0
        'oSheet.WindowInfo.SplitColumns = 1

        ' Split after row 2 (ScrollRow Is zero based).  
        oSheet.WindowInfo.ScrollRow = 0
        oSheet.WindowInfo.SplitRows = 0

        ' Freeze the panes. 
        oSheet.WindowInfo.FreezePanes = True

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(xls_filename)
        oWB.Close()

    End Sub

    Private Sub UltraButton2_Click(sender As Object, e As EventArgs) Handles UltraButton2.Click
        Dim PARTNER_CODE As String = "EFNY"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim worksheet1 As SpreadsheetGear.IWorksheet = Nothing

        Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
        Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")


        dst.Tables("APTVVVV3").Rows.Clear()
        Fill_Records("APTVVVV3", New Object() {dteStart.Value, dteEnd.Value})

        Dim DT As DataTable = dst.Tables("APTVVVV3")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("TATPRTN1.EXTRACT_NO") & "-" & PARTNER_CODE & ".xlsX" '  ASCMAIN1.Folders("Archive") & "\Frame Supplier Account Info.xlsX"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")


        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If

        Next

        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        range = oSheet.Cells("A1:CD1")
        range.Interior.Color = SpreadsheetGear.Colors.Beige
        '  range.EntireColumn.AutoFilter()
        range.EntireColumn.AutoFit()

        'oSheet.WindowInfo.ScrollColumn = 0
        'oSheet.WindowInfo.SplitColumns = 1

        ' Split after row 2 (ScrollRow Is zero based).  
        oSheet.WindowInfo.ScrollRow = 0
        oSheet.WindowInfo.SplitRows = 0

        ' Freeze the panes. 
        oSheet.WindowInfo.FreezePanes = True

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(xls_filename)
        oWB.Close()

    End Sub

    Private Sub UltraButton3_Click(sender As Object, e As EventArgs) Handles UltraButton3.Click
        Dim PARTNER_CODE As String = "EFNY"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim worksheet1 As SpreadsheetGear.IWorksheet = Nothing

        Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
        Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")


        dst.Tables("APTVVVV4").Rows.Clear()
        Fill_Records("APTVVVV4", New Object() {dteStart.Value, dteEnd.Value})

        Dim DT As DataTable = dst.Tables("APTVVVV4")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("TATPRTN1.EXTRACT_NO") & "-" & PARTNER_CODE & ".xlsX" '  ASCMAIN1.Folders("Archive") & "\Frame Supplier Account Info.xlsX"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")


        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If

        Next

        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        range = oSheet.Cells("A1:AZ1")
        range.Interior.Color = SpreadsheetGear.Colors.Beige
        '  range.EntireColumn.AutoFilter()
        range.EntireColumn.AutoFit()

        'oSheet.WindowInfo.ScrollColumn = 0
        'oSheet.WindowInfo.SplitColumns = 1

        ' Split after row 2 (ScrollRow Is zero based).  
        oSheet.WindowInfo.ScrollRow = 0
        oSheet.WindowInfo.SplitRows = 0

        ' Freeze the panes. 
        oSheet.WindowInfo.FreezePanes = True

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(xls_filename)
        oWB.Close()

    End Sub

    Private Sub UltraButton4_Click(sender As Object, e As EventArgs) Handles UltraButton4.Click
        Dim PARTNER_CODE As String = "EFNY"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim worksheet1 As SpreadsheetGear.IWorksheet = Nothing

        Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
        Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")


        dst.Tables("APTVVVV5").Rows.Clear()
        Fill_Records("APTVVVV5", New Object() {dteStart.Value, dteEnd.Value})

        Dim DT As DataTable = dst.Tables("APTVVVV5")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("TATPRTN1.EXTRACT_NO") & "-" & PARTNER_CODE & ".xlsX" '  ASCMAIN1.Folders("Archive") & "\Frame Supplier Account Info.xlsX"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")


        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If

        Next

        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        range = oSheet.Cells("A1:H1")
        range.Interior.Color = SpreadsheetGear.Colors.Beige
        '  range.EntireColumn.AutoFilter()
        range.EntireColumn.AutoFit()

        'oSheet.WindowInfo.ScrollColumn = 0
        'oSheet.WindowInfo.SplitColumns = 1

        ' Split after row 2 (ScrollRow Is zero based).  
        oSheet.WindowInfo.ScrollRow = 0
        oSheet.WindowInfo.SplitRows = 0

        ' Freeze the panes. 
        oSheet.WindowInfo.FreezePanes = True

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(xls_filename)
        oWB.Close()

    End Sub

    Private Sub UltraButton5_Click(sender As Object, e As EventArgs) Handles UltraButton5.Click
        Dim PARTNER_CODE As String = "EFNY"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim worksheet1 As SpreadsheetGear.IWorksheet = Nothing

        Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
        Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")


        dst.Tables("APTVVVV6").Rows.Clear()
        Fill_Records("APTVVVV6", New Object() {dteStart.Value, dteEnd.Value})

        Dim DT As DataTable = dst.Tables("APTVVVV6")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("TATPRTN1.EXTRACT_NO") & "-" & PARTNER_CODE & ".xlsX" '  ASCMAIN1.Folders("Archive") & "\Frame Supplier Account Info.xlsX"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")


        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If

        Next

        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        range = oSheet.Cells("A1:BD1")
        range.Interior.Color = SpreadsheetGear.Colors.Beige
        '  range.EntireColumn.AutoFilter()
        range.EntireColumn.AutoFit()

        'oSheet.WindowInfo.ScrollColumn = 0
        'oSheet.WindowInfo.SplitColumns = 1

        ' Split after row 2 (ScrollRow Is zero based).  
        oSheet.WindowInfo.ScrollRow = 0
        oSheet.WindowInfo.SplitRows = 0

        ' Freeze the panes. 
        oSheet.WindowInfo.FreezePanes = True

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(xls_filename)
        oWB.Close()

    End Sub

    Private Sub UltraButton6_Click(sender As Object, e As EventArgs) Handles UltraButton6.Click
        Dim PARTNER_CODE As String = "EFNY"

        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim worksheet1 As SpreadsheetGear.IWorksheet = Nothing

        Dim STARTDT As String = Format(dteStart.Value, "dd-MMM-yyyy")
        Dim ENDDT As String = Format(dteEnd.Value, "dd-MMM-yyyy")


        dst.Tables("APTVVVV7").Rows.Clear()
        Fill_Records("APTVVVV7", New Object() {dteStart.Value, dteEnd.Value})

        Dim DT As DataTable = dst.Tables("APTVVVV7")
        Dim xls_filename As String = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("TATPRTN1.EXTRACT_NO") & "-" & PARTNER_CODE & ".xlsX" '  ASCMAIN1.Folders("Archive") & "\Frame Supplier Account Info.xlsX"
        Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()

        Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = oSheet.Cells("A1")


        For c As Integer = 0 To DT.Columns.Count - 1
            If DT.Columns(c).DataType.ToString = "System.String" Then
                oSheet.Cells(0, c).EntireColumn.NumberFormat = "@"
            End If

        Next

        range.CopyFromDataTable(DT, SpreadsheetGear.Data.SetDataFlags.None)

        range = oSheet.Cells("A1:BD1")
        range.Interior.Color = SpreadsheetGear.Colors.Beige
        '  range.EntireColumn.AutoFilter()
        range.EntireColumn.AutoFit()

        'oSheet.WindowInfo.ScrollColumn = 0
        'oSheet.WindowInfo.SplitColumns = 1

        ' Split after row 2 (ScrollRow Is zero based).  
        oSheet.WindowInfo.ScrollRow = 0
        oSheet.WindowInfo.SplitRows = 0

        ' Freeze the panes. 
        oSheet.WindowInfo.FreezePanes = True

        oWB.SaveAs(xls_filename, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(xls_filename)
        oWB.Close()

    End Sub

    Private Sub ABSCheckBox5_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
End Class