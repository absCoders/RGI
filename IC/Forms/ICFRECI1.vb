Imports System.Drawing
Imports System.Math

Public Class ICFRECI1
    Dim RYP As String
    'Dim ICTRECI0 As String
    'Dim COLs_WITH_DESCs As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select Y.*" & vbCrLf _
                & ", ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & " from ICTSTYL1, (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", SUM (BOMQTY) BOMQTY, SUM (BOMCST) BOMCST" & vbCrLf _
                & ", SUM (RECQTY) RECQTY, SUM (RECCST) RECCST" & vbCrLf _
                & ", SUM (SHPQTY) SHPQTY, SUM (SHPCST) SHPCST" & vbCrLf _
                & ", SUM (ADJQTY) ADJQTY, SUM (ADJCST) ADJCST" & vbCrLf _
                & ", SUM (EOMQTY) EOMQTY, SUM (EOMCST) EOMCST" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select ICTSTAT5.STYLE_CODE, ICTSTAT5.COLOR_CODE" & vbCrLf _
                & ", SUM (ICTSTAT5.WHSE_QTY_ON_HAND) BOMQTY" & vbCrLf _
                & ", SUM (NVL(ICTSTAT5.WHSE_QTY_ON_HAND,0) * NVL(ICTCOSTA.STYLE_COST,0)) BOMCST" & vbCrLf _
                & ", 0 RECQTY, 0 RECCST" & vbCrLf _
                & ", 0 SHPQTY, 0 SHPCST" & vbCrLf _
                & ", 0 ADJQTY, 0 ADJCST" & vbCrLf _
                & ", 0 EOMQTY, 0 EOMCST" & vbCrLf _
                & " from ICTSTAT5,ICTCOSTA" & vbCrLf _
                & " where ICTSTAT5.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP (+) = ICTSTAT5.OPS_YYYYPP" & vbCrLf _
                & "   and ICTCOSTA.STYLE_CODE (+) = ICTSTAT5.STYLE_CODE" & vbCrLf _
                & "   and ICTCOSTA.COLOR_CODE (+) = ICTSTAT5.COLOR_CODE" & vbCrLf _
                & " group by ICTSTAT5.STYLE_CODE, ICTSTAT5.COLOR_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                & ", 0 BOMQTY, 0 BOMCST" & vbCrLf _
                & ", SUM (POTSHIP3.PO_QTY_REC) RECQTY" & vbCrLf _
                & ", SUM (POTSHIP3.PO_QTY_REC * POTSHIP3.PO_COST_LANDED) RECCST" & vbCrLf _
                & ", 0 SHPQTY, 0 SHPCST" & vbCrLf _
                & ", 0 ADJQTY, 0 ADJCST" & vbCrLf _
                & ", 0 EOMQTY, 0 EOMCST" & vbCrLf _
                & " from POTSHIP2,POTSHIP3,POTORDR2,ICTSTYL1 WHERE POTSHIP2.OPS_YYYYPP = :PARM2" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                & "" & vbCrLf _
                & " union " & vbCrLf _
                & "Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                & ", 0 BOMQTY, 0 BOMCST" & vbCrLf _
                & ", 0 RECQTY, 0 RECCST" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP) SHPQTY" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_COST) SHPCST" & vbCrLf _
                & ", 0 ADJQTY, 0 ADJCST" & vbCrLf _
                & ", 0 EOMQTY, 0 EOMCST" & vbCrLf _
                & " from SOTINVH2,ICTSTYL1 WHERE SOTINVH2.ORDR_YYYYPP_UPDATED = :PARM3" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE" & vbCrLf _
                & " group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                & "" & vbCrLf _
                & " union " & vbCrLf _
                & "Select ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE" & vbCrLf _
                & ", 0 BOMQTY, 0 BOMCST" & vbCrLf _
                & ", 0 RECQTY, 0 RECCST" & vbCrLf _
                & ", 0 SHPQTY, 0 SHPCST" & vbCrLf _
                & ", SUM (ICTSTAT1.WHSE_QTY_ADJ) ADJQTY" & vbCrLf _
                & ", 0 ADJCST" & vbCrLf _
                & ", 0 EOMQTY, 0 EOMCST" & vbCrLf _
                & " from ICTSTAT1,ICTSTYL1" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = ICTSTAT1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTAT1.OPS_YYYYPP = :PARM4" & vbCrLf _
                & " group by ICTSTAT1.STYLE_CODE, ICTSTAT1.COLOR_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select ICTSTAT5.STYLE_CODE, ICTSTAT5.COLOR_CODE" & vbCrLf _
                & ", 0 BOMQTY, 0 BOMCST" & vbCrLf _
                & ", 0 RECQTY, 0 RECCST" & vbCrLf _
                & ", 0 SHPQTY, 0 SHPCST" & vbCrLf _
                & ", 0 ADJQTY, 0 ADJCST" & vbCrLf _
                & ", SUM (ICTSTAT5.WHSE_QTY_ON_HAND) EOMQTY" & vbCrLf _
                & ", SUM (NVL(ICTSTAT5.WHSE_QTY_ON_HAND,0) * NVL(ICTCOSTA.STYLE_COST,0)) EOMCST" & vbCrLf _
                & " from ICTSTAT5,ICTCOSTA" & vbCrLf _
                & " where ICTSTAT5.OPS_YYYYPP = :PARM5" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP (+) = ICTSTAT5.OPS_YYYYPP" & vbCrLf _
                & "   and ICTCOSTA.STYLE_CODE (+) = ICTSTAT5.STYLE_CODE" & vbCrLf _
                & "   and ICTCOSTA.COLOR_CODE (+) = ICTSTAT5.COLOR_CODE" & vbCrLf _
                & " group by ICTSTAT5.STYLE_CODE, ICTSTAT5.COLOR_CODE" & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE) Y" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = Y.STYLE_CODE" & vbCrLf _
                & " and (Y.BOMQTY <> 0 or Y.RECQTY <> 0 or Y.SHPQTY <> 0 or Y.ADJQTY <> 0 or Y.EOMQTY <> 0" _
                & "   or Y.BOMCST <> 0 or Y.RECCST <> 0 or Y.SHPCST <> 0 or Y.ADJCST <> 0 or Y.EOMCST <> 0)"

            Create_TDA(.Tables.Add, "ICTRECI0", "**", 0, False, "VVVVV")
            .Tables("ICTRECI0").Columns.Add("OOBQTY", GetType(System.Int64), "ISNULL(BOMQTY,0)+ISNULL(RECQTY,0)-ISNULL(SHPQTY,0)+ISNULL(ADJQTY,0)-ISNULL(EOMQTY,0)")
            .Tables("ICTRECI0").Columns.Add("OOBCST", GetType(System.Int64), "ISNULL(BOMCST,0)+ISNULL(RECCST,0)-ISNULL(SHPCST,0)+ISNULL(ADJCST,0)-ISNULL(EOMCST,0)")

            'ASCMAIN1.sql = "Select * from ICTCLAS1"
            'Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False, 1)

            ' .Relations.Add(ASCDATA1.GetRelation(dst, "ICTCLAS1", "ICTRECI0", "CLASS_CODE"))

            '  .Tables("ICTRECI0").Columns.Add("CLASS_DESC", GetType(System.String), "PARENT(ICTCLAS1_ICTRECI0).CLASS_DESC")


            ASCMAIN1.sql = "SELECT GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & ", SUM (GLTDETL1.DETL_POSTING_AMT) GL_AMT" & vbCrLf _
            & ", SUM (DECODE(GLTDETL1.ACCT_CODE,'125600',GLTDETL1.DETL_POSTING_AMT,0)) GL_AMT_C" & vbCrLf _
            & ", SUM (DECODE(GLTDETL1.ACCT_CODE,'130000',GLTDETL1.DETL_POSTING_AMT,0)) GL_AMT_R" & vbCrLf _
            & " from GLTDETL1, GLTJRNL1" & vbCrLf _
            & " WHERE GLTDETL1.ACCT_CODE IN ('130000','125600')" & vbCrLf _
            & " AND GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO" & vbCrLf _
            & " AND GLTDETL1.OPS_YYYYPP = :PARM1" & vbCrLf _
            & " GROUP BY GLTJRNL1.JOURNAL_TYPE"
            Create_TDA(.Tables.Add, "ICTRECIG", "**", 0, False, "V", 1)
            .Tables("ICTRECIG").Columns.Add("IC_AMT", GetType(System.Decimal))
            .Tables("ICTRECIG").Columns.Add("IC_AMT_C", GetType(System.Decimal))
            .Tables("ICTRECIG").Columns.Add("IC_AMT_R", GetType(System.Decimal))
            .Tables("ICTRECIG").Columns("IC_AMT").Expression = "ISNULL(IC_AMT_C,0)+ISNULL(IC_AMT_R,0)"

            .Tables("ICTRECIG").Columns.Add("DIFF", GetType(System.Decimal), "ISNULL(IC_AMT,0)-ISNULL(GL_AMT,0)")

            ' old
            ''ASCMAIN1.sql = " Select POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, MIN(DECODE(ICTIREC1.VOUCHER_NO,NULL,POTSHIP2.VOUCHER_NO,ICTIREC1.VOUCHER_NO)),MIN (POTSHIP2.PO_DATE_RECEIVED) PO_DATE_RECEIVED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
            ''& " , SUM (POTSHIP3.PO_QTY_REC) QTY" & vbCrLf _
            ''& " , SUM (POTSHIP3.PO_QTY_REC * POTSHIP3.PO_COST_LANDED) AMT from POTSHIP2,POTSHIP3,POTSHIP1" & vbCrLf _
            ''& " WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            ''& " AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            ''& " AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            ''& " AND POTSHIP2.OPS_YYYYPP = :PARM1" & vbCrLf _
            ''& " AND ICTIREC1.RECEIPT_NO(+) = POTSHIP2.TRAN_NO" & vbCrLf _
            ''& " GROUP BY POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, " & vbCrLf _
            ''& " POTSHIP1.PO_SHIP_ETA"

            'new 

            ASCMAIN1.sql = "Select pO_SHIPMENT_NO,PO_SHIP_VESSEL,PO_DATE_RECEIVED,PO_SHIP_ETA,QTY,AMT,APTINVH1.CHECK_date FROM (" & vbCrLf _
            & " Select POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, MAX(DECODE(ICTIREC1.VOUCHER_NO,NULL,POTSHIP2.VOUCHER_NO,ICTIREC1.VOUCHER_NO)) VOUCHER_NO,MIN (POTSHIP2.PO_DATE_RECEIVED) PO_DATE_RECEIVED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
            & " , SUM (POTSHIP3.PO_QTY_REC) QTY" & vbCrLf _
            & " , SUM (POTSHIP3.PO_QTY_REC * POTSHIP3.PO_COST_LANDED) AMT from POTSHIP2,POTSHIP3,POTSHIP1,ICTIREC1" & vbCrLf _
            & " WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & " And POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & " And POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & " AND POTSHIP2.OPS_YYYYPP = :PARM1" & vbCrLf _
            & "  And ICTIREC1.RECEIPT_NO(+) = POTSHIP2.TRAN_NO" & vbCrLf _
            & "  GROUP by POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
            & ", POTSHIP1.PO_SHIP_ETA) DD, APTINVH1" & vbCrLf _
            & " WHERE APTINVH1.VOUCHER_NO(+) = DD.VOUCHER_NO"
            Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "V", 1)




            Create_TDA(.Tables.Add, "POTSHIP2", "*", 1)



            ''ASCMAIN1.sql = " Select POTSHIPH.*,E_RECEIVED, POTSHIP1.PO_SHIP_ETA" & vbCrLf _
            ''& " , SUM (POTSHIP3.PO_QTY_REC) QTY" & vbCrLf _
            ''& " , SUM (POTSHIP3.PO_QTY_REC * POTSHIP3.PO_COST_LANDED) AMT from POTSHIP2,POTSHIP3,POTSHIP1" & vbCrLf _
            ''& " WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            ''& " And POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            ''& " And POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            ''& " And POTSHIP2.OPS_YYYYPP = :PARM1" & vbCrLf _
            ''& " GROUP BY POTSHIP1.PO_SHIPMENT_NO, POTSHIP1.PO_SHIP_VESSEL, " & vbCrLf _
            ''& " POTSHIP1.PO_SHIP_ETA"
            ''Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "V", 1)


            ASCMAIN1.sql = "Select POTSHIPH.*,APTVEND1.VEND_CODE,APTVEND1.VEND_NAME,APTINVH1.INV_NUM,APTINVH1.INV_DATE,APTINVH1.CHECK_DATE FROM POTSHIPH,APTINVH1,APTVEND1 " & vbCrLf _
            & " WHERE POTSHIPH.OPS_YYYYPP  = :PARM1" & vbCrLf _
            & " And APTINVH1.VOUCHER_NO(+) = POTSHIPH.VOUCHER_NO" & vbCrLf _
            & " And APTVEND1.VEND_CODE(+) = APTINVH1.VEND_CODE"
            Create_TDA(.Tables.Add, "POTSHIPH", "**", 0, False, "V", 5)

            '    Create_TDA(.Tables.Add, "POTSHIPH", "*", 1)
            With .Tables("POTSHIPH").Columns
                ' .Add("PO_QTY_SHP_EXT", GetType(System.Int32), "PO_QTY_SHP * (PO_COST)")
                .Add("PO_QTY_SHP_EXT", GetType(System.Decimal), "PO_QTY_SHP * (PO_COST_VCOST + PO_COST_MATLS + PO_COST_OTHER)")

            End With



            ASCMAIN1.sql = "Select PO_ORDER_NO, PO_ORDER_LNO, PO_QTY_OPN from POTORDR2" _
                    & " where (PO_ORDER_NO, PO_ORDER_LNO) in " _
                    & " (Select Distinct PO_ORDER_NO, PO_ORDER_LNO from POTSHIP3 where PO_SHIPMENT_NO = :PARM1)"
                Create_TDA(.Tables.Add, "POTORDRO", "**", 0, False, "V", 2)


                ASCMAIN1.sql = "Select POTSHIP3.* " & vbCrLf _
                  & ", POTORDR1.VEND_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE " & vbCrLf _
                  & ", POTORDR2.PO_QTY_OPN, POTORDR2.PO_QTY_UOM, POTORDR2.PO_COST ORDR2_COST" & vbCrLf _
                  & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_BODY_CODE, POTORDR2.SUB_UNIT_PACK_QTY, POTORDR2.CARTON_PACK_QTY" & vbCrLf _
                  & ", POTORDR1.PO_REFERENCE, POTORDR2.PO_DATE_SHIP_BY" & vbCrLf _
                  & ", POTSHIP3.PO_QTY_REC PO_QTY_REC_OLD" & vbCrLf _
                  & " from POTSHIP3,POTORDR2,ICTSTYL1,POTORDR1 " & vbCrLf _
                  & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                  & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                  & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                  & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                  & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1"
                Create_TDA(.Tables.Add, "POTSHIP3", "**", 0, True, "V", 4)

                Create_Relation("POTORDRO", "POTSHIP3", "PO_ORDER_NO,PO_ORDER_LNO")
                .Tables("POTORDRO").Columns.Add("PO_QTY_SHP", GetType(System.Int32), "SUM(CHILD(POTORDRO_POTSHIP3).PO_QTY_SHP)")
                .Tables("POTORDRO").Columns.Add("PO_QTY_OPN_PRE", GetType(System.Int32))

                Create_Relation("POTSHIP2", "POTSHIP3", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")
                With .Tables("POTSHIP3").Columns
                    .Add("PO_QTY_VAR", GetType(System.Int32), "IIF(PARENT(POTSHIP2_POTSHIP3).PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_REC,0) - ISNULL(PO_QTY_SHP,0) = 0,NULL,ISNULL(PO_QTY_REC,0) - ISNULL(PO_QTY_SHP,0)))")
                    .Add("PO_QTY_OPN_PRE", GetType(System.Int32), "PARENT(POTORDRO_POTSHIP3).PO_QTY_OPN_PRE")
                    .Add("PO_QTY_SHP_DZ", GetType(System.Int32), "PO_QTY_SHP / (12 / SUB_UNIT_PACK_QTY)")
                    .Add("PO_QTY_REC_DZ", GetType(System.Int32), "PO_QTY_REC / (12 / SUB_UNIT_PACK_QTY)")
                    .Add("PO_SHIP_STATUS", GetType(System.String), "PARENT(POTSHIP2_POTSHIP3).PO_SHIP_STATUS")
                    .Add("PO_QTY_SR", GetType(System.Int32), "ISNULL(PO_QTY_SHP,0)")

                    .Add("PO_QTY_SR_DZ", GetType(System.Int32), "IIF(PO_SHIP_STATUS='C',ISNULL(PO_QTY_REC_DZ,0),ISNULL(PO_QTY_SHP_DZ,0))")
                    .Add("TOTAL_DUTY", GetType(System.Decimal), "PO_QTY_SR * ISNULL(PO_COST_DUTY,0)")
                    .Add("CONTAINER_NO", GetType(System.String), "PARENT(POTSHIP2_POTSHIP3).CONTAINER_NO")
                    .Add("NET_OPEN", GetType(System.Decimal), "PO_QTY_OPN")

                    .Add("NET_OPEN_DZ", GetType(System.Decimal), "NET_OPEN / (12 / SUB_UNIT_PACK_QTY)")
                    .Add("PO_AMT_REC", GetType(System.Decimal), "ISNULL(PO_COST_VCOST,0) * ISNULL(PO_QTY_REC,0)")
                    .Add("FIRST_COST_TOTAL", GetType(System.Decimal), "ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0)")
                    .Add("FIRST_COST_TOTAL_DZ", GetType(System.Decimal), "(PO_COST_VCOST_DZ + PO_COST_MATLS_DZ + PO_COST_OTHER_DZ)")
                    .Add("COMMISSION_COST", GetType(System.Decimal), "(((ISNULL(PO_COST_COMM,0)+ISNULL(PO_COST_BUFFER,0)) / 100) * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0) + ISNULL(PO_COST_QUOTA,0)))")
                    .Add("COMMISSION_COST_DZ", GetType(System.Decimal), "(((ISNULL(PO_COST_COMM,0)+ISNULL(PO_COST_BUFFER,0)) / 100) * (ISNULL(PO_COST_VCOST_DZ,0) + ISNULL(PO_COST_MATLS_DZ,0) + ISNULL(PO_COST_OTHER_DZ,0) + ISNULL(PO_COST_QUOTA_DZ,0)))")

                    .Add("EXT_WEIGHT_FACTOR", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(WEIGHT_FACTOR,0)")
                    .Add("EXT_VCOST", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_VCOST,0)")
                    .Add("EXT_MATLS", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_MATLS,0)")
                    .Add("EXT_OTHER", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_OTHER,0)")
                    .Add("EXT_FIRST", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(FIRST_COST_TOTAL,0)")
                    .Add("EXT_COMM", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_COMM,0)/100 * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0) + ISNULL(PO_COST_QUOTA,0))")
                    .Add("EXT_BUFFER", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_BUFFER,0)/100 * (ISNULL(PO_COST_VCOST,0) + ISNULL(PO_COST_MATLS,0) + ISNULL(PO_COST_OTHER,0) + ISNULL(PO_COST_QUOTA,0))")
                    .Add("EXT_QUOTA", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_QUOTA,0)")
                    .Add("EXT_QUOTA_DF", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_QUOTA_DF,0)")
                    .Add("EXT_FREIGHT", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_FREIGHT_IN,0)")
                    .Add("EXT_CUSTOMS", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_CUSTOMS,0)")
                    .Add("EXT_DUTY", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_DUTY,0)")
                    .Add("EXT_TRUCKING", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_TRUCKING,0)")
                    .Add("EXT_LANDED", GetType(System.Decimal), "ISNULL(PO_QTY_SR,0) * ISNULL(PO_COST_LANDED,0)")
                    .Add("EXT_FIRST_CALC", GetType(System.Decimal), "EXT_VCOST + EXT_MATLS + EXT_OTHER")
                    .Add("EXT_LANDED_CALC", GetType(System.Decimal), "EXT_FIRST + EXT_COMM + EXT_BUFFER + EXT_QUOTA + EXT_QUOTA_DF + EXT_FREIGHT + EXT_CUSTOMS + EXT_DUTY + EXT_TRUCKING")

                    'If Not cost_calc Then
                    '    For Each COLUMN_NAME As String In New String() _
                    '        {"EXT_WEIGHT_FACTOR", "EXT_VCOST", "EXT_MATLS", "EXT_OTHER", "EXT_FIRST", _
                    '         "EXT_COMM", "EXT_BUFFER", "EXT_QUOTA", "EXT_QUOTA_DF", "EXT_FREIGHT", _
                    '         "EXT_CUSTOMS", "EXT_DUTY", "EXT_TRUCKING", "EXT_LANDED"}
                    '        dst.Tables("POTSHIP3").Columns(COLUMN_NAME).Expression = ""
                    '    Next
                    'End If

                    .Add("LINE_EXACT", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_SHP,0) = ISNULL(PO_QTY_REC,0),1,0))")
                    .Add("LINE_OVER", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_SHP,0) < ISNULL(PO_QTY_REC,0),1,0))")
                    .Add("LINE_SHORT", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_SHP,0) > ISNULL(PO_QTY_REC,0),1,0))")
                    .Add("LINE_ZERO", GetType(System.Int64), "IIF(PO_SHIP_STATUS='O',Null,IIF(ISNULL(PO_QTY_REC,0) = 0,1,0))")
                    .Add("CBM", GetType(System.Decimal))

                End With

            End With

            ' COLs_WITH_DESCs.Add("CLASS_CODE")

            grdICTRECI0.DataSource = dst.Tables("ICTRECI0")
        grdICTRECIG.DataSource = dst.Tables("ICTRECIG")
        grdPOTSHIPX.DataSource = dst.Tables("POTSHIPX")
        grdPOTSHIP3.DataSource = dst.Tables("POTSHIP3")
        grdPOTSHIPH.DataSource = dst.Tables("POTSHIPH")
        ' Fill_Records("ICTCLAS1")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTRECIG}
            With grd.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"GL_AMT", "GL_AMT_C", "GL_AMT_R"}
                    With .Columns(COLUMN_NAME)
                        .Header.Appearance.BackColor2 = Color.Yellow
                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        .CellAppearance.BackColor = Color.LightYellow
                        .Width = 100
                        .Format = "###,##0.00"
                    End With
                    Create_Summary(grd, COLUMN_NAME)
                Next

                For Each COLUMN_NAME As String In New String() {"IC_AMT", "IC_AMT_C", "IC_AMT_R"}
                    With .Columns(COLUMN_NAME)
                        .Header.Appearance.BackColor2 = Color.Lime
                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        .CellAppearance.BackColor = Color.LightYellow
                        .Width = 100
                        .Format = "###,##0.00"
                    End With
                    Create_Summary(grd, COLUMN_NAME)
                Next

                For Each COLUMN_NAME As String In New String() {"DIFF"}
                    With .Columns(COLUMN_NAME)
                        .Header.Appearance.BackColor2 = Color.HotPink
                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        .CellAppearance.BackColor = Color.Beige
                        .Width = 100
                        .Format = "###,##0.00"
                    End With
                    Create_Summary(grd, COLUMN_NAME)
                Next

                .Columns("IC_AMT").CellAppearance.BackColor = Color.Beige
                .Columns("GL_AMT").CellAppearance.BackColor = Color.Beige
                .Columns("DIFF").CellAppearance.BackColor = Color.Beige

                .Columns("JOURNAL_TYPE").CellAppearance.BackColor = Color.LightGray

            End With
        Next

        With grdICTRECI0.DisplayLayout.Bands(0)
            .Columns("STYLE_CLASS_CODE").Hidden = False
            For Each COLUMN_NAME As String In New String() _
            {"STYLE_CODE", "COLOR_CODE", "STYLE_CLASS_CODE", "STYLE_DESC"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor = Color.White
                    .Header.Appearance.BackColor2 = Color.Yellow
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .CellAppearance.BackColor = Color.LightYellow
                End With
            Next
            Create_Summary(grdICTRECI0, "STYLE_CODE", "Count")
            For Each COLUMN_NAME As String In New String() _
                {"BOMQTY", "RECQTY", "SHPQTY", "ADJQTY", "EOMQTY", "OOBQTY"}
                With .Columns(COLUMN_NAME)
                    .Format = "###,##0"
                    .Width = 70
                    .Header.Appearance.BackColor = Color.White
                    .Header.Appearance.BackColor2 = Color.LightBlue
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                Create_Summary(grdICTRECI0, COLUMN_NAME)
            Next
            For Each COLUMN_NAME As String In New String() _
               {"BOMCST", "RECCST", "SHPCST", "ADJCST", "EOMCST", "OOBCST"}
                With .Columns(COLUMN_NAME)
                    .Format = "###,##0"
                    .Width = 80
                    .Header.Appearance.BackColor = Color.White
                    .Header.Appearance.BackColor2 = Color.LightGreen
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                Create_Summary(grdICTRECI0, COLUMN_NAME)
            Next
        End With

        Create_Summary(grdICTRECIG, "JOURNAL_TYPE", "Count")
        Create_Summary(grdPOTSHIPX, New String() {"QTY", "AMT"})

        Create_Summary(grdPOTSHIPH, New String() {"PO_QTY_SHP", "PO_QTY_SHP_EXT"})

        With grdICTRECI0.DisplayLayout.Bands("ICTRECI0")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("STYLE_CLASS_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
        End With


        grdPOTSHIP3.DisplayLayout.UseFixedHeaders = True
        With grdPOTSHIP3.DisplayLayout.Bands(0)
            .Columns("PO_SHIPMENT_LNO").CellActivation = UltraWinGrid.Activation.NoEdit

            For Each COLUMN_NAME As String In New String() _
                    {"EXT_WEIGHT_FACTOR", "CBM", "COST_CHANGED", "EXT_VCOST", "EXT_MATLS", "EXT_OTHER", "EXT_FIRST",
                     "EXT_COMM", "EXT_BUFFER", "EXT_QUOTA", "EXT_QUOTA_DF", "EXT_FREIGHT",
                     "EXT_CUSTOMS", "EXT_DUTY", "EXT_TRUCKING", "EXT_LANDED", "EXT_FIRST_CALC", "EXT_LANDED_CALC",
                     "LINE_EXACT", "LINE_OVER", "LINE_SHORT", "LINE_ZERO"}
                .Columns(COLUMN_NAME).Hidden = True
            Next

            For Each COLUMN_NAME In New String() {"PO_ORDER_NO", "PO_ORDER_LNO", "PO_REFERENCE", "PO_DATE_SHIP_BY"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.Beige


                If New String() {"PO_ORDER_NO", "PO_ORDER_LNO", "PO_REFERENCE", "PO_DATE_SHIP_BY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PO_QTY_SHP", "PO_QTY_REC", "NET_OPEN", "PO_QTY_VAR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                ElseIf New String() {"PO_QTY_SHP_DZ", "PO_QTY_REC_DZ", "NET_OPEN_DZ"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightSalmon
                ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"PO_UOM", "SUB_UNIT_PACK_QTY", "CARTON_PACK_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"PO_COST_QUOTA", "PO_COST_QUOTA_DF", "PO_COST_QUOTA_DZ", "PO_COST_QUOTA_DF_DZ",
                                     "PO_COST_BUFFER", "PO_COST_COMM", "COMMISSION_COST", "COMMISSION_COST_DZ",
                                     "PO_COST_VCOST_UM", "PO_COST_VCOST_DZ", "PO_COST_MATLS_UM", "PO_COST_MATLS_DZ",
                                     "PO_COST_OTHER", "PO_COST_OTHER_DZ", "FIRST_COST_TOTAL", "FIRST_COST_TOTAL_DZ",
                                     "PO_COST_FREIGHT_IN", "PO_COST_CUSTOMS", "PO_COST_DUTY", "PO_COST_TRUCKING", "PO_COST_LANDED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        ASCMAIN1.Add_Value_List(grdPOTSHIPH, "ACCRUAL_STATUS", Nothing, New String() {":", "O:Not Paid", "1:Invoiced(Paid)"})


        cbeOPS_YYYYPP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' and OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' order by OPS_YYYYPP DESC")
        cbeOPS_YYYYPP.ValueMember = "OPS_YYYYPP"
        cbeOPS_YYYYPP.DisplayMember = "LEGEND"
        cbeOPS_YYYYPP.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If Absx1.cbeFor("OPS_YYYYPP").Value = "" Then
                    EMsg &= vbCrLf & "You must specify a Period to View"
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

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)


            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Options").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode

        Setup_tabMain()

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("ICTRECI0").Rows.Clear()
        dst.Tables("ICTRECIG").Rows.Clear()
        EnforceConstraints(True)

        chkOOB.Checked = False
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        RYP = Absx1.cbeFor("OPS_YYYYPP").Value

        Load_ICTRECI0()
        Load_ICTRECIG()
        Load_POTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub



    Sub Update_Record()

        Call BeginTrans()
        Stop
        Call CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdICTRECI0, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
        Call Load_Popup_Menu(grdICTRECIG, "SSB", "Show Filter", "Show GroupBox")

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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key


            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "OPS_YYYYPP"
                Call Click_Command("View")
        End Select
    End Sub
#End Region

    Sub Load_POs_into_POTSHIP3()

        grdPOTSHIP3.Text = "Bill of Lading Details for Style: " & grdICTRECI0.ActiveRow.Cells("STYLE_CODE").Value & "" & " Color: " & grdICTRECI0.ActiveRow.Cells("COLOR_CODE").Value & ""

        dst.Tables("POTSHIP3").Rows.Clear()
        ASCMAIN1.sql = " select POTORDR2.*, POTSHIP2.PO_SHIPMENT_NO, POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
        & " from POTSHIP2,POTSHIP3,POTSHIP1,POTORDR2" & vbCrLf _
        & " WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
        & " AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
        & " AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
        & " AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
        & " AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
        & " AND POTORDR2.STYLE_CODE = '" & grdICTRECI0.ActiveRow.Cells("STYLE_CODE").Value & "'" & vbCrLf _
        & " AND POTORDR2.COLOR_CODE = '" & grdICTRECI0.ActiveRow.Cells("COLOR_CODE").Value & "'" & vbCrLf _
        & " AND POTSHIP2.OPS_YYYYPP = '" & RYP & "'"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim PO_SHIPMENT_NO As String = row.Item("PO_SHIPMENT_NO") & ""
            Dim PO_SHIPMENT_LNO As Integer = Val(row.Item("PO_SHIPMENT_LNO") & "")
            Dim PO_ORDER_NO As Integer = Val(row.Item("PO_ORDER_NO") & "")
            Dim PO_ORDER_LNO As Integer = Val(row.Item("PO_ORDER_LNO") & "")
            Dim SUB_UNIT_PACK_QTY As Integer = Val(row.Item("SUB_UNIT_PACK_QTY") & "")
            Dim SPQ As Integer = IIf(SUB_UNIT_PACK_QTY = 0, 12, SUB_UNIT_PACK_QTY)
            Dim PO_QTY_OPN As Int64 = Val(row.Item("PO_QTY_OPN") & "")

            Fill_Records("POTORDRO", PO_SHIPMENT_NO)
            'dst.Tables("POTSHIP3").Columns("CBM").Expression = ""
            Fill_Records("POTSHIP2", PO_SHIPMENT_NO)

            Dim rowPOTSHIP3 As DataRow = dst.Tables("POTSHIP3").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO, PO_ORDER_NO, PO_ORDER_LNO})
            If rowPOTSHIP3 Is Nothing Then
                rowPOTSHIP3 = dst.Tables("POTSHIP3").NewRow
                With rowPOTSHIP3

                    .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                    .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                    .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                    .Item("PO_QTY_SHP") = Val(row.Item("PO_QTY_SHP") & "")
                    .Item("PO_QTY_OPN") = PO_QTY_OPN
                    .Item("PO_QTY_REC") = Val(row.Item("PO_QTY_REC") & "")
                    .Item("PO_QTY_UOM") = row.Item("PO_QTY_UOM")
                    .Item("PO_COST") = Val(row.Item("PO_COST") & "")
                    .Item("PO_COST_VCOST") = Val(row.Item("PO_COST_VCOST") & "")
                    .Item("PO_COST_MATLS") = Val(row.Item("PO_COST_MATLS") & "")
                    .Item("PO_COST_OTHER") = Val(row.Item("PO_COST_OTHER") & "") / SPQ
                    .Item("PO_COST_COMM") = Val(row.Item("PO_COST_COMM") & "")
                    .Item("PO_COST_BUFFER") = 5
                    .Item("PO_COST_LANDED") = Val(row.Item("PO_COST") & "")
                    If Val(row.Item("DFQUOTA") & "") = 1 Then
                        .Item("PO_COST_QUOTA_DF") = Val(row.Item("PO_COST_QUOTA") & "") / SPQ
                        .Item("PO_COST_QUOTA_DF_DZ") = Val(row.Item("PO_COST_QUOTA") & "")
                    Else
                        .Item("PO_COST_QUOTA") = Val(row.Item("PO_COST_QUOTA") & "") / SPQ
                        .Item("PO_COST_QUOTA_DZ") = Val(row.Item("PO_COST_QUOTA") & "")
                    End If
                    If Val(row.Item("PO_COST_VCOST_DZ") & "") = 0 Then
                        .Item("PO_COST_VCOST_DZ") = Val(row.Item("PO_COST_VCOST") & "") * SPQ
                    Else
                        .Item("PO_COST_VCOST_DZ") = Val(row.Item("PO_COST_VCOST_DZ") & "")
                    End If
                    If Val(row.Item("PO_COST_MATLS_DZ") & "") = 0 Then
                        .Item("PO_COST_MATLS_DZ") = Val(row.Item("PO_COST_MATLS") & "") * SPQ
                    Else
                        .Item("PO_COST_MATLS_DZ") = Val(row.Item("PO_COST_MATLS_DZ") & "")
                    End If
                    .Item("PO_COST_OTHER_DZ") = Val(row.Item("PO_COST_OTHER") & "")
                    .Item("SUB_UNIT_PACK_QTY") = Val(row.Item("SUB_UNIT_PACK_QTY") & "")
                    .Item("CARTON_PACK_QTY") = Val(row.Item("CARTON_PACK_QTY") & "")
                    If Val(row.Item("SUB_UNIT_PACK_QTY") & "") = 0 Then
                        .Item("PO_QTY_SHP_DZ") = 0
                        .Item("NET_OPEN_DZ") = 0
                    Else
                        .Item("PO_QTY_SHP_DZ") = PO_QTY_OPN / (12 / Val(row.Item("SUB_UNIT_PACK_QTY") & ""))
                        .Item("NET_OPEN_DZ") = PO_QTY_OPN / (12 / Val(row.Item("SUB_UNIT_PACK_QTY") & ""))
                    End If
                    .Item("PO_QTY_REC_DZ") = 0

                    Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                    .Item("PO_REFERENCE") = rowPOTORDR1.Item("PO_REFERENCE")
                    .Item("PO_DATE_SHIP_BY") = row.Item("PO_DATE_SHIP_BY")
                    .Item("FOB_CMT") = (rowPOTORDR1.Item("FOB_CMT") & "")
                    .Item("VEND_CODE") = rowPOTORDR1.Item("VEND_CODE")

                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", row.Item("STYLE_CODE"))
                    .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                End With
                dst.Tables("POTSHIP3").Rows.Add(rowPOTSHIP3)
            End If
        Next
    End Sub

    Sub Load_POTSHIPX()
        Fill_Records("POTSHIPX", RYP)
        Fill_Records("POTSHIPH", RYP)

        grdPOTSHIPH.Text = "In Transit: " & RYP

    End Sub

    Sub Load_ICTRECI0()

        'ASCDATA1.ExecuteSQL("Truncate Table " & ICTRECI0)


        'ASCMAIN1.sql = "DELETE FROM " & ICTRECI0 & " WHERE NVL(CASES,0) = 0 AND NVL(UNITS,0) = 0 AND NVL(XSTDC,0) = 0"
        'ASCDATA1.ExecuteSQL()

        Fill_Records("ICTRECI0", New String() {ASCMAIN1.Period_Calc(RYP, -1), RYP, RYP, RYP, RYP})

    End Sub

    Sub Load_ICTRECIG()
        Fill_Records("ICTRECIG", RYP)
  
        Dim GLIC As New Dictionary(Of String, String)
        GLIC.Add("OPCG", "CGS")
        GLIC.Add("APIN", "API")
        GLIC.Add("ICLL", "LIQ")
        GLIC.Add("ICPR", "REC")
        GLIC.Add("ICQA", "ADJ")
        GLIC.Add("ICSR", "WHS")
        GLIC.Add("OPXA", "OPX")
        GLIC.Add("ICTR", "XFR")

        For Each GLIC_key As String In GLIC.Keys
            Dim rowICTRECIG As DataRow = dst.Tables("ICTRECIG").Rows.Find(GLIC_key)
            If rowICTRECIG Is Nothing Then
                rowICTRECIG = dst.Tables("ICTRECIG").NewRow
                rowICTRECIG.Item("JOURNAL_TYPE") = GLIC_key
                dst.Tables("ICTRECIG").Rows.Add(rowICTRECIG)
            End If
            'rowICTRECIG.Item("IC_AMT_C") = Val(dst.Tables("ICTRECI0").Compute("SUM(" & GLIC(GLIC_key) & ")", "CON_REG_IND = 'C'") & "")
            'rowICTRECIG.Item("IC_AMT_R") = Val(dst.Tables("ICTRECI0").Compute("SUM(" & GLIC(GLIC_key) & ")", "CON_REG_IND = 'R'") & "")
        Next

      
    End Sub

    Private Sub grdICTRECI0_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTRECI0.AfterRowActivate
        Load_POs_into_POTSHIP3()
    End Sub



    Private Sub grdICTRECI0_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTRECI0.InitializeRow
        Dim OOBQTY As Decimal = Val(e.Row.Cells("OOBQTY").Value & "")
        If Abs(OOBQTY) > 0.01 Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Color.Red
            e.Row.Cells("OOBQTY").Appearance.ForeColor = Color.Red
        End If
        Dim OOBCST As Decimal = Val(e.Row.Cells("OOBCST").Value & "")
        If Abs(OOBCST) > 1 Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Color.Red
            e.Row.Cells("OOBCST").Appearance.ForeColor = Color.Red

        End If
    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

    Private Sub chkOOB_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkOOB.CheckedChanged
        Dim dvw As DataView = DirectCast(grdICTRECI0.DataSource, DataTable).DefaultView
        If chkOOB.Checked Then
            dvw.RowFilter = "OOBQTY <> 0 or OOBCST <> 0"
            grdICTRECI0.Text = "Inventory Roll Forward - Out of Balance Records Only"
        Else
            dvw.RowFilter = ""
            grdICTRECI0.Text = "Inventory Roll Forward"
        End If
    End Sub

    Private Sub grdICTRECI0_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTRECI0.InitializeLayout

    End Sub
End Class