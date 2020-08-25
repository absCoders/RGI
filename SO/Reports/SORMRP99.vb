Imports System.Security
Imports System.Text

Public Class SORMRP99
    Private xRYP0_legend As String
    Private xRYP1_legend As String
    Private xRYP0 As String
    Private xRYP1 As String
    Dim SQLs As StringBuilder

    'Dim ARTCUST1 As String
    Dim BPeriod As String
    Dim EPeriod As String
    Dim SOTCSTY1 As String = ""
    Dim TEMP_LIMIT As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        chkMAKEFLAT_STATE()

        With grdSOWMRP9X.DisplayLayout.Bands(0)
            For Each COLNAME As String In New String() {"WIP_1", "WIP_2", "WIP_3", "WIP_4", "WIP_5", "TRAN",
                "ON_HAND", "ON_OPEN", "ON_PICK", "ORDR", "CANCL", "SHIPPED_1", "SHIPPED_2", "SHIPPED_3",
                "SHIPPED_4", "SHIPPED_5"}
                .Columns(COLNAME).Format = "###,##0"
            Next
        End With
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        Dim REF_NO As New String(" ", 40)

        xRYP0_legend = Absx1.cmbFor("RYP0").Value
        xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
        xRYP1_legend = Absx1.cmbFor("RYP1").Value
        xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
        BPeriod = Format(ASCMAIN1.Get_Dates(xRYP0).ElementAt(1), "dd-MMM-yyyy")
        EPeriod = Format(ASCMAIN1.Get_Dates(xRYP1).ElementAt(ASCMAIN1.Get_Dates(xRYP1).Length - 1), "dd-MMM-yyyy")
        If xRYP0 = xRYP1 Then
            SUBT = "Data Posted in " & xRYP0_legend
        Else
            SUBT = String.Format("Data Posted between {0} and {1}", xRYP0_legend, xRYP1_legend)
        End If

        ASCMAIN1.Progress("Gathering Data", "")
        Dim SOTCSTY1_SQL As String = "SELECT STYLE_CODE, COLOR_CODE, CUST_CODE, MIN (CUST_STYLE_CODE) CUST_STYLE_CODE, MIN(VENDOR_STOCK_NO) VENDOR_STOCK_NO"
        SOTCSTY1_SQL &= " FROM SOTCSTY1"
        SOTCSTY1_SQL &= " GROUP BY STYLE_CODE, COLOR_CODE, CUST_CODE"
        SOTCSTY1 = ASCMAIN1.Temp_Table(SOTCSTY1_SQL)

        Dim CUST_CODEs As String() = Split(SQLA("CUST_CODE"), ",")
        Dim sqlw As String = SQL_in("STYLE_CODE", "SJ.STYLE_CODE")
        Dim sqlw_SJ As String = sqlw + " " + SQL_in("WHSE_CODE", "SJ.WHSE_CODE")
        Dim sqlw_POTORDR1 As String = sqlw + " " + SQL_in("WHSE_CODE", "POTORDR1.WHSE_CODE")
        Dim sqlw_SOTORDR1 As String = sqlw + " " + SQL_in("WHSE_CODE", "SOTORDR1.WHSE_CODE")
        Dim sqlw_SOTINVH1 As String = SQL_in("WHSE_CODE", "SOTINVH1.WHSE_CODE")

        'Dim s As New StringBuilder With {.Length = 0}
        's.AppendLine("SELECT SOTCSTY1.STYLE_CODE, SOTCSTY1.COLOR_CODE, SOTCSTY1.CUST_CODE, MIN (SOTCSTY1.CUST_STYLE_CODE) CUST_STYLE_CODE")
        's.AppendLine("FROM SOTCSTY1, ICTSTAT2")
        's.AppendLine("WHERE SOTCSTY1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
        's.AppendLine("AND SOTCSTY1.COLOR_CODE = ICTSTAT2.COLOR_CODE")
        'If CUST_CODEs.Length > 0 Then
        '    s.AppendLine(String.Format("AND SOTCSTY1.CUST_CODE = '{0}'", CUST_CODEs(0)))
        'End If
        's.AppendLine("AND (NVL(WHSE_QTY_ON_ORDER,0) > 0 OR NVL(WHSE_QTY_TRAN,0) > 0 OR NVL(WHSE_QTY_OPEN,0) > 0 OR NVL(WHSE_QTY_ON_HAND,0) > 0)")
        's.AppendLine("GROUP BY SOTCSTY1.STYLE_CODE, SOTCSTY1.COLOR_CODE, SOTCSTY1.CUST_CODE")
        'SOTCSTY1 = ASCMAIN1.Temp_Table(s.ToString)

        With dst
            SQLs = New StringBuilder() With {.Length = 0}
            SQLs.AppendLine(String.Format("SELECT CUST_CODE, STYLE_CODE, COLOR_CODE, VENDOR_STOCK_NO, REC_TYPE, X_DATE,'{0}' REF_NO,", REF_NO))
            SQLs.AppendLine(" SUM(NVL(WIP,0)) WIP,")
            SQLs.AppendLine(" SUM(NVL(TRAN,0)) TRAN,")
            SQLs.AppendLine(" SUM(NVL(ON_HAND,0)) ON_HAND,")
            SQLs.AppendLine(" SUM(NVL(ON_OPEN,0)) ON_OPEN,")
            SQLs.AppendLine(" SUM(NVL(ON_PICK,0)) ON_PICK,")
            SQLs.AppendLine(" SUM(NVL(SHIPPED,0)) SHIPPED,")
            SQLs.AppendLine(" SUM(NVL(SHIPPED,0)) ORDR,")
            SQLs.AppendLine(" SUM(NVL(SHIPPED,0)) CANCL,")
            SQLs.AppendLine(" SUM(NVL(SHIPPED,0)) RECD,")
            SQLs.AppendLine(" CUST_STYLE_CODE")
            SQLs.AppendLine(" FROM")
            SQLs.AppendLine(" (")
            SQLs.AppendLine(" SELECT ICTSTYL1.CUST_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO , 'W' REC_TYPE,")
            SQLs.AppendLine(" DECODE(POTSHIP1.PO_SHIP_ETA, NULL,POTORDR2.PO_DATE_ETA,POTSHIP1.PO_SHIP_ETA) X_DATE,")
            SQLs.AppendLine(String.Format(" '{0}' REF_NO,", REF_NO))
            SQLs.AppendLine(" SUM(NVL(POTORDR2.PO_QTY_ORD,0) - NVL(POTORDR2.PO_QTY_SHP,0)) WIP,")
            SQLs.AppendLine(" SUM(NVL(POTORDR2.PO_QTY_SHP,0) - NVL(POTORDR2.PO_QTY_REC,0)) TRAN,")
            SQLs.AppendLine(" SUM(0) ON_HAND,")
            SQLs.AppendLine(" SUM(0) ON_OPEN,")
            SQLs.AppendLine(" SUM(0) ON_PICK,")
            SQLs.AppendLine(" SUM(0) SHIPPED,")
            SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
            SQLs.AppendLine(" FROM POTORDR2, POTSHIP3, POTSHIP1, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
            SQLs.AppendLine(" WHERE POTORDR2.PO_ORDER_NO =  POTSHIP3.PO_ORDER_NO (+)")
            SQLs.AppendLine(" AND POTORDR2.PO_ORDER_LNO =  POTSHIP3.PO_ORDER_LNO (+)")
            SQLs.AppendLine(" AND POTSHIP3.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO (+)")
            SQLs.AppendLine(" AND POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            'SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
            'SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
            SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE")
            SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE")
            SQLs.AppendLine(" AND ROWNUM < 0")
            SQLs.AppendLine(String.Format(" GROUP BY ICTSTYL1.CUST_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO , 'P', DECODE(POTSHIP1.PO_SHIP_ETA, NULL,POTORDR2.PO_DATE_ETA,POTSHIP1.PO_SHIP_ETA),'{0}', SOTCSTY1.CUST_STYLE_CODE", REF_NO))
            SQLs.AppendLine(" )")
            SQLs.AppendLine(String.Format(" GROUP BY CUST_CODE, STYLE_CODE, COLOR_CODE, VENDOR_STOCK_NO, REC_TYPE, X_DATE,'{0}', CUST_STYLE_CODE", REF_NO))
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "SOWMRP99", "**", 0, False, "", 0)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("ICTSTYL1.CUST_CODE,")
            SQLs.AppendLine("POTORDR2.STYLE_CODE,")
            SQLs.AppendLine("POTORDR2.COLOR_CODE,")
            SQLs.AppendLine("SOTCSTY1.VENDOR_STOCK_NO,")
            SQLs.AppendLine("SOTCSTY1.CUST_STYLE_CODE,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS WIP_1,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD)AS WIP_2,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS WIP_3,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS WIP_4,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS WIP_5,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS TRAN,")
            SQLs.AppendLine("POTSHIP1.PO_SHIP_ETA AS TRAN_DATE,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS ON_HAND,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS ON_OPEN,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS ON_PICK,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS ORDR,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS CANCL,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS SHIPPED_1,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS SHIPPED_2,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS SHIPPED_3,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS SHIPPED_4,")
            SQLs.AppendLine("SUM(POTORDR2.PO_QTY_ORD) AS SHIPPED_5")
            SQLs.AppendLine("FROM POTORDR2, POTSHIP3, POTSHIP1, ICTSTYL1, SOTCSTY1")
            SQLs.AppendLine("WHERE POTORDR2.PO_ORDER_NO =  POTSHIP3.PO_ORDER_NO (+)")
            SQLs.AppendLine("AND POTORDR2.PO_ORDER_LNO =  POTSHIP3.PO_ORDER_LNO (+)")
            SQLs.AppendLine("AND POTSHIP3.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO (+)")
            SQLs.AppendLine("AND POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            SQLs.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE")
            SQLs.AppendLine("AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE")
            SQLs.AppendLine("AND ROWNUM < 0")
            SQLs.AppendLine("GROUP BY")
            SQLs.AppendLine("ICTSTYL1.CUST_CODE,")
            SQLs.AppendLine("POTORDR2.STYLE_CODE,")
            SQLs.AppendLine("POTORDR2.COLOR_CODE,")
            SQLs.AppendLine("SOTCSTY1.VENDOR_STOCK_NO,")
            SQLs.AppendLine("POTSHIP1.PO_SHIP_ETA,")
            SQLs.AppendLine("SOTCSTY1.CUST_STYLE_CODE")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "SOWMRP9X", "**", 0, False, "", 0)

            ASCMAIN1.Progress("Master Files", "")

            ASCMAIN1.sql = "Select * from ARTCUST1"
            .Tables.Add(ASCDATA1.GetDataTable("", "ARWCUST1", 1))

            ASCMAIN1.sql = "Select * from ICTSTYL1"
            .Tables.Add(ASCDATA1.GetDataTable("", "ICWSTYL1", 1))

            'Dim CUST_CODEs As String() = Split(SQLA("CUST_CODE"), ",")
            'Dim sqlw As String = SQL_in("STYLE_CODE", "SJ.STYLE_CODE")
            'Dim sqlw_SJ As String = sqlw + " " + SQL_in("WHSE_CODE", "SJ.WHSE_CODE")
            'Dim sqlw_POTORDR1 As String = sqlw + " " + SQL_in("WHSE_CODE", "POTORDR1.WHSE_CODE")
            'Dim sqlw_SOTORDR1 As String = sqlw + " " + SQL_in("WHSE_CODE", "SOTORDR1.WHSE_CODE")
            'Dim sqlw_SOTINVH1 As String = SQL_in("WHSE_CODE", "SOTINVH1.WHSE_CODE")
            For Each CUST_CODE As String In CUST_CODEs
                ASCMAIN1.Progress("Building WIP", CUST_CODE)
                'SQLs.Length = 0
                'SQLs.AppendLine(" SELECT ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, 'W' REC_TYPE,")
                'SQLs.AppendLine(" DECODE(POTSHIP1.PO_SHIP_ETA, NULL,SJ.PO_DATE_ETA,POTSHIP1.PO_SHIP_ETA) X_DATE,")
                'SQLs.AppendLine(" DECODE(POTSHIP1.PO_SHIP_ETA, NULL,POTORDR1.PO_REFERENCE,POTORDR1.PO_REFERENCE || ' - ' || POTSHIP1.PO_SHIP_VESSEL) REF_NO,")
                'SQLs.AppendLine(" SUM(SJ.PO_QTY_ORD - SJ.PO_QTY_SHP) WIP,")
                'SQLs.AppendLine(" SUM(SJ.PO_QTY_SHP - SJ.PO_QTY_REC) TRAN,")
                'SQLs.AppendLine(" SUM(0) ON_HAND,")
                'SQLs.AppendLine(" SUM(0) ON_OPEN,")
                'SQLs.AppendLine(" SUM(0) ON_PICK,")
                'SQLs.AppendLine(" SUM(0) SHIPPED,")
                'SQLs.AppendLine(" SUM(0) ORDR,")
                'SQLs.AppendLine(" SUM(0) CANCL,")
                'SQLs.AppendLine(" SUM(0) RECD")
                'SQLs.AppendLine(" FROM POTORDR1, POTORDR2 SJ, POTSHIP3, POTSHIP1, ICTSTYL1")
                'SQLs.AppendLine(" WHERE SJ.PO_ORDER_NO =  POTSHIP3.PO_ORDER_NO (+)")
                'SQLs.AppendLine(" AND SJ.PO_ORDER_LNO =  POTSHIP3.PO_ORDER_LNO (+)")
                'SQLs.AppendLine(" AND POTORDR1.PO_ORDER_NO = SJ.PO_ORDER_NO")
                'SQLs.AppendLine(" AND SJ.PO_STATUS <> 'C'")
                'SQLs.AppendLine(" AND POTSHIP3.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO (+)")
                'SQLs.AppendLine(" AND SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                'If sqlw.Length > 0 Then
                '    SQLs.AppendLine(sqlw)
                'End If
                'SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                'SQLs.AppendLine(" HAVING SUM((SJ.PO_QTY_ORD - SJ.PO_QTY_SHP) - (SJ.PO_QTY_SHP - SJ.PO_QTY_REC)) <> 0")
                'SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, 'P', DECODE(POTSHIP1.PO_SHIP_ETA, NULL,SJ.PO_DATE_ETA,POTSHIP1.PO_SHIP_ETA), DECODE(POTSHIP1.PO_SHIP_ETA, NULL,POTORDR1.PO_REFERENCE,POTORDR1.PO_REFERENCE || ' - ' || POTSHIP1.PO_SHIP_VESSEL)")
                'ASCMAIN1.sql = SQLs.ToString()
                'Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)

                SQLs.Length = 0
                SQLs.AppendLine("SELECT ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'W' REC_TYPE,")
                If chkMAKEFLAT.Checked = False Then
                    SQLs.AppendLine(" SYSDATE X_DATE,")
                Else
                    SQLs.AppendLine(" POTORDR1.PO_DATE_ETA AS X_DATE,")
                End If
                SQLs.AppendLine(" NULL REF_NO,")
                SQLs.AppendLine(" SUM(NVL(SJ.PO_QTY_OPN,0)) WIP,")
                SQLs.AppendLine(" SUM(0) TRAN,")
                SQLs.AppendLine(" SUM(0) ON_HAND,")
                SQLs.AppendLine(" SUM(0) ON_OPEN,")
                SQLs.AppendLine(" SUM(0) ON_PICK,")
                SQLs.AppendLine(" SUM(0) SHIPPED,")
                SQLs.AppendLine(" SUM(0) ORDR,")
                SQLs.AppendLine(" SUM(0) CANCL,")
                SQLs.AppendLine(" SUM(0) RECD,")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                SQLs.AppendLine(" FROM POTORDR1, POTORDR2 SJ, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                SQLs.AppendLine(" WHERE(POTORDR1.PO_ORDER_NO = SJ.PO_ORDER_NO)")
                SQLs.AppendLine(" AND SJ.PO_STATUS <> 'C'")
                SQLs.AppendLine(" AND SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                If sqlw_POTORDR1.Length > 0 Then
                    SQLs.AppendLine(sqlw_POTORDR1)
                End If
                SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                SQLs.AppendLine(" HAVING SUM(SJ.PO_QTY_OPN) <> 0")
                SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'W', SOTCSTY1.CUST_STYLE_CODE")
                If chkMAKEFLAT.Checked = True Then
                    SQLs.AppendLine(" ,POTORDR1.PO_DATE_ETA")
                End If
                ASCMAIN1.sql = SQLs.ToString()
                Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)

                SQLs.Length = 0
                SQLs.AppendLine("SELECT ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'W' REC_TYPE,")
                If chkMAKEFLAT.Checked = False Then
                    SQLs.AppendLine(" SYSDATE X_DATE,")
                Else
                    SQLs.AppendLine(" POTSHIP1.PO_SHIP_ETA AS X_DATE,")
                End If
                SQLs.AppendLine(" NULL REF_NO,")
                SQLs.AppendLine(" SUM(0) WIP,")
                'SQLs.AppendLine(" SUM(NVL(SJ.PO_QTY_SHP,0) - NVL(SJ.PO_QTY_REC,0)) TRAN,")
                'SQLs.AppendLine(" DECODE(SUM(NVL(SJ.PO_QTY_OPN,0)),0,0,SUM(NVL(SJ.PO_QTY_SHP,0) - NVL(SJ.PO_QTY_REC,0))) TRAN,")
                SQLs.AppendLine(" DECODE(SUM(NVL(SJ.PO_QTY_OPN,0)),0,SUM(NVL(SJ.PO_QTY_SHP,0) - NVL(SJ.PO_QTY_REC,0)),0) TRAN,")
                SQLs.AppendLine(" SUM(0) ON_HAND,")
                SQLs.AppendLine(" SUM(0) ON_OPEN,")
                SQLs.AppendLine(" SUM(0) ON_PICK,")
                SQLs.AppendLine(" SUM(0) SHIPPED,")
                SQLs.AppendLine(" SUM(0) ORDR,")
                SQLs.AppendLine(" SUM(0) CANCL,")
                SQLs.AppendLine(" SUM(0) RECD,")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                SQLs.AppendLine(" FROM POTORDR1, POTORDR2 SJ, POTSHIP3, POTSHIP2, POTSHIP1, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                SQLs.AppendLine(" WHERE SJ.PO_ORDER_NO =  POTSHIP3.PO_ORDER_NO (+)")
                SQLs.AppendLine(" AND SJ.PO_ORDER_LNO =  POTSHIP3.PO_ORDER_LNO (+)")
                SQLs.AppendLine(" AND SJ.PO_ORDER_NO =  POTORDR1.PO_ORDER_NO")
                SQLs.AppendLine(" AND POTSHIP2.PO_SHIPMENT_NO =  POTSHIP3.PO_SHIPMENT_NO")
                SQLs.AppendLine(" AND POTSHIP2.PO_SHIPMENT_LNO =  POTSHIP3.PO_SHIPMENT_LNO")
                SQLs.AppendLine(" AND POTSHIP3.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO (+)")
                SQLs.AppendLine(" AND SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                SQLs.AppendLine(" AND POTSHIP2.PO_SHIP_STATUS <> 'C'")
                If sqlw_POTORDR1.Length > 0 Then
                    SQLs.AppendLine(sqlw_POTORDR1)
                End If
                SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                SQLs.AppendLine(" HAVING DECODE(SUM(NVL(SJ.PO_QTY_OPN,0)),0,SUM(NVL(SJ.PO_QTY_SHP,0) - NVL(SJ.PO_QTY_REC,0)),0) > 0")
                SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'W',")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                If chkMAKEFLAT.Checked = True Then
                    SQLs.AppendLine(" , POTSHIP1.PO_SHIP_ETA")
                End If
                ASCMAIN1.sql = SQLs.ToString()
                Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)

                ASCMAIN1.Progress("Building On Hand", CUST_CODE)
                'Dim ST_TABLE As String = "ICTSTAT2"
                SQLs.Length = 0
                SQLs.AppendLine(" SELECT ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'H' REC_TYPE, SYSDATE X_DATE,")
                SQLs.AppendLine(" NULL REF_NO,")
                SQLs.AppendLine(" SUM(0) WIP,")
                SQLs.AppendLine(" SUM(0) TRAN,")
                SQLs.AppendLine(" SUM(NVL(WHSE_QTY_ON_HAND,0)) ON_HAND,")
                SQLs.AppendLine(" SUM(0) ON_OPEN,")
                SQLs.AppendLine(" SUM(0) ON_PICK,")
                SQLs.AppendLine(" SUM(0) SHIPPED,")
                SQLs.AppendLine(" SUM(0) ORDR,")
                SQLs.AppendLine(" SUM(0) CANCL,")
                SQLs.AppendLine(" SUM(0) RECD,")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                SQLs.AppendLine(" FROM ICTSTAT2 SJ, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                SQLs.AppendLine(" WHERE SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                If sqlw_SJ.Length > 0 Then
                    SQLs.AppendLine(sqlw_SJ)
                End If
                SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'O', SOTCSTY1.CUST_STYLE_CODE")
                ASCMAIN1.sql = SQLs.ToString()
                Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)

                ASCMAIN1.Progress("Building Open and Pick", CUST_CODE)
                SQLs.Length = 0
                SQLs.AppendLine("SELECT ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, SJ.ORDR_STATUS REC_TYPE, SYSDATE X_DATE,")
                SQLs.AppendLine(" NULL REF_NO,")
                SQLs.AppendLine(" SUM(0) WIP,")
                SQLs.AppendLine(" SUM(0) TRAN,")
                SQLs.AppendLine(" SUM(0) ON_HAND,")
                SQLs.AppendLine(" SUM(NVL(SJ.ORDR_QTY_OPEN,0)) ON_OPEN,")
                SQLs.AppendLine(" SUM(NVL(SJ.ORDR_QTY_PICK,0)) ON_PICK,")
                SQLs.AppendLine(" SUM(0) SHIPPED,")
                SQLs.AppendLine(" SUM(0) ORDR,")
                SQLs.AppendLine(" SUM(0) CANCL,")
                SQLs.AppendLine(" SUM(0) RECD,")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                SQLs.AppendLine(" FROM SOTORDR1, SOTORDR2 SJ, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                SQLs.AppendLine(" WHERE SOTORDR1.ORDR_NO = SJ.ORDR_NO")
                SQLs.AppendLine(" AND SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                If sqlw_SOTORDR1.Length > 0 Then
                    SQLs.AppendLine(sqlw_SOTORDR1)
                End If
                SQLs.AppendLine(" AND (SJ.ORDR_STATUS = 'P' OR SJ.ORDR_STATUS = 'O')")
                SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, SJ.ORDR_STATUS, SOTCSTY1.CUST_STYLE_CODE")
                ASCMAIN1.sql = SQLs.ToString()
                Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)

                If chkIncludeReservations.Checked Then
                    SQLs.Length = 0
                    SQLs.AppendLine("SELECT ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'V' REC_TYPE, SYSDATE X_DATE,")
                    SQLs.AppendLine(" NULL REF_NO,")
                    SQLs.AppendLine(" SUM(0) WIP,")
                    SQLs.AppendLine(" SUM(0) TRAN,")
                    SQLs.AppendLine(" SUM(0) ON_HAND,")
                    SQLs.AppendLine(" SUM(NVL(SJ.RSRV_QTY_OPEN,0)) ON_OPEN,")
                    SQLs.AppendLine(" SUM(0) ON_PICK,")
                    SQLs.AppendLine(" SUM(0) SHIPPED,")
                    SQLs.AppendLine(" SUM(0) ORDR,")
                    SQLs.AppendLine(" SUM(0) CANCL,")
                    SQLs.AppendLine(" SUM(0) RECD,")
                    SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                    SQLs.AppendLine(" FROM SOTRSRV1, SOTRSRV2 SJ, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                    SQLs.AppendLine(" WHERE SOTRSRV1.RSRV_NO = SJ.RSRV_NO")
                    SQLs.AppendLine(" AND SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                    SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                    SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                    SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                    If sqlw_SOTORDR1.Length > 0 Then
                        SQLs.AppendLine(sqlw_SOTORDR1)
                    End If
                    SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, SOTCSTY1.CUST_STYLE_CODE")
                    ASCMAIN1.sql = SQLs.ToString()
                    Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)
                End If

                ASCMAIN1.Progress("Building Order and Cancel", CUST_CODE)
                SQLs.Length = 0
                SQLs.AppendLine(" SELECT ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'X' REC_TYPE, SYSDATE X_DATE,")
                SQLs.AppendLine(" NULL REF_NO,")
                SQLs.AppendLine(" SUM(0) WIP,")
                SQLs.AppendLine(" SUM(0) TRAN,")
                SQLs.AppendLine(" SUM(0) ON_HAND,")
                SQLs.AppendLine(" SUM(0) ON_OPEN,")
                SQLs.AppendLine(" SUM(0) ON_PICK,")
                SQLs.AppendLine(" SUM(0) SHIPPED,")
                SQLs.AppendLine(" SUM(NVL(SJ.ORDR_QTY,0)) ORDR,")
                SQLs.AppendLine(" SUM(NVL(SJ.ORDR_QTY_CANC,0)) CANCL,")
                SQLs.AppendLine(" SUM(0) RECD,")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                SQLs.AppendLine(" FROM SOTORDR1, SOTORDR2 SJ, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                SQLs.AppendLine(" WHERE SOTORDR1.ORDR_NO = SJ.ORDR_NO")
                SQLs.AppendLine(" AND SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                If sqlw_SOTORDR1.Length > 0 Then
                    SQLs.AppendLine(sqlw_SOTORDR1)
                End If
                SQLs.AppendLine(String.Format(" AND (SOTORDR1.ORDR_DATE_BOOKED >= '{0}' AND SOTORDR1.ORDR_DATE_BOOKED <= '{1}')", BPeriod, EPeriod))
                SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'X', SOTCSTY1.CUST_STYLE_CODE")
                ASCMAIN1.sql = SQLs.ToString()
                Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)

                '--- Add Received Records
                SQLs.Length = 0
                SQLs.AppendLine(" SELECT ")
                SQLs.AppendLine(" ICTSTYL1.CUST_CODE, ")
                SQLs.AppendLine(" SJ.STYLE_CODE, ")
                SQLs.AppendLine(" SJ.COLOR_CODE, ")
                SQLs.AppendLine(" SOTCSTY1.VENDOR_STOCK_NO,")
                SQLs.AppendLine(" 'R' REC_TYPE,")
                SQLs.AppendLine(" SYSDATE X_DATE,")
                SQLs.AppendLine(" NULL REF_NO,")
                SQLs.AppendLine(" SUM(0) WIP,")
                SQLs.AppendLine(" SUM(0) TRAN,")
                SQLs.AppendLine(" SUM(0) ON_HAND,")
                SQLs.AppendLine(" SUM(0) ON_OPEN,")
                SQLs.AppendLine(" SUM(0) ON_PICK,")
                SQLs.AppendLine(" SUM(0) SHIPPED,")
                SQLs.AppendLine(" SUM(0) ORDR,")
                SQLs.AppendLine(" SUM(0) CANCL,")
                SQLs.AppendLine(" SUM(NVL(SJ.PO_QTY_REC,0)) RECD,")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                SQLs.AppendLine(" FROM POTORDR1, POTORDR2 SJ, POTSHIP3, POTSHIP1, POTSHIP2, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                SQLs.AppendLine(" WHERE SJ.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                SQLs.AppendLine(" AND SJ.PO_ORDER_NO =  POTSHIP3.PO_ORDER_NO (+)")
                SQLs.AppendLine(" AND SJ.PO_ORDER_LNO =  POTSHIP3.PO_ORDER_LNO (+)")
                SQLs.AppendLine(" AND POTORDR1.PO_ORDER_NO = SJ.PO_ORDER_NO")
                SQLs.AppendLine(" AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO (+)")
                SQLs.AppendLine(" AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO")
                SQLs.AppendLine(" AND POTSHIP2.PO_SHIPMENT_LNO = POTSHIP3.PO_SHIPMENT_LNO")
                SQLs.AppendLine(" AND SJ.PO_STATUS <> 'C'")
                SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                SQLs.AppendLine(String.Format(" AND POTSHIP2.OPS_YYYYPP >= '{0}' AND POTSHIP2.OPS_YYYYPP <= '{0}'", xRYP0))
                SQLs.AppendLine(String.Format(" AND ICTSTYL1.CUST_CODE = '{0}'", CUST_CODE))
                If sqlw_POTORDR1.Length > 0 Then
                    SQLs.AppendLine(sqlw_POTORDR1)
                End If
                SQLs.AppendLine(" HAVING SUM(SJ.PO_QTY_REC) > 0")
                SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SJ.STYLE_CODE, SJ.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'R', ")
                SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                ASCMAIN1.sql = SQLs.ToString()
                Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)

                If Absx1.chkFor("CHKSHIPPED").Checked = True Then
                    Dim STYLE_CODE As String = ""
                    Dim COLOR_CODE As String = ""
                    For Each rowSOWMRP99 As DataRow In dst.Tables("SOWMRP99").Select("", "STYLE_CODE, COLOR_CODE")
                        If (STYLE_CODE <> rowSOWMRP99.Item("STYLE_CODE").ToString) Or (COLOR_CODE <> rowSOWMRP99.Item("COLOR_CODE").ToString) Then
                            STYLE_CODE = rowSOWMRP99.Item("STYLE_CODE").ToString
                            COLOR_CODE = rowSOWMRP99.Item("COLOR_CODE").ToString
                            SQLs.Length = 0
                            SQLs.AppendLine(" SELECT ICTSTYL1.CUST_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'S' REC_TYPE,")
                            If chkMAKEFLAT.Checked = False Then
                                SQLs.AppendLine(" SYSDATE X_DATE,")
                            Else
                                SQLs.AppendLine(" SOTINVH1.INV_DATE AS X_DATE,")
                            End If
                            SQLs.AppendLine(" NULL REF_NO,")
                            SQLs.AppendLine(" SUM(0) WIP,")
                            SQLs.AppendLine(" SUM(0) TRAN,")
                            SQLs.AppendLine(" SUM(0) ON_HAND,")
                            SQLs.AppendLine(" SUM(0) ON_OPEN,")
                            SQLs.AppendLine(" SUM(0) ON_PICK,")
                            SQLs.AppendLine(" SUM(NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SHIPPED,")
                            SQLs.AppendLine(" SUM(0) ORDR,")
                            SQLs.AppendLine(" SUM(0) CANCL,")
                            SQLs.AppendLine(" SUM(0) RECD,")
                            SQLs.AppendLine(" SOTCSTY1.CUST_STYLE_CODE")
                            SQLs.AppendLine(" FROM SOTINVH1, SOTINVH2, ICTSTYL1, " & SOTCSTY1 & " SOTCSTY1")
                            SQLs.AppendLine(" WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
                            SQLs.AppendLine(" AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE")
                            SQLs.AppendLine(" AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
                            SQLs.AppendLine(String.Format(" AND SOTINVH2.STYLE_CUST_CODE = '{0}'", CUST_CODE))
                            SQLs.AppendLine(String.Format(" AND SOTINVH2.STYLE_CODE = '{0}'", STYLE_CODE))
                            SQLs.AppendLine(String.Format(" AND SOTINVH2.COLOR_CODE = '{0}'", COLOR_CODE))
                            SQLs.AppendLine(" AND SOTINVH1.INV_TYPE = 'I'")
                            If chkMAKEFLAT.Checked = False Then
                                SQLs.AppendLine(String.Format(" AND SOTINVH1.ORDR_YYYYPP_UPDATED >= '{0}'", xRYP0))
                                SQLs.AppendLine(String.Format(" AND SOTINVH1.ORDR_YYYYPP_UPDATED <= '{0}'", xRYP1))
                            End If
                            SQLs.AppendLine(" AND ICTSTYL1.STYLE_CODE = SOTCSTY1.STYLE_CODE (+)")
                            SQLs.AppendLine(" AND ICTSTYL1.CUST_CODE = SOTCSTY1.CUST_CODE (+)")
                            If sqlw_POTORDR1.Length > 0 Then
                                SQLs.AppendLine(sqlw_SOTINVH1)
                            End If
                            SQLs.AppendLine(" GROUP BY ICTSTYL1.CUST_CODE, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTCSTY1.VENDOR_STOCK_NO, 'S', SOTCSTY1.CUST_STYLE_CODE")
                            If chkMAKEFLAT.Checked = True Then
                                SQLs.AppendLine(" , SOTINVH1.INV_DATE")
                            End If
                            ASCMAIN1.Progress(String.Format("Building Shipped for {0}", CUST_CODE), String.Format("{0}", STYLE_CODE))
                            ASCMAIN1.sql = SQLs.ToString()
                            Fill_Records("SOWMRP99", "", False, ASCMAIN1.sql)
                        End If
                    Next
                End If
            Next



        End With

        Dim Filter As String = "WIP = 0" &
        " AND TRAN = 0" &
        " AND ON_HAND = 0" &
        " AND ON_PICK = 0" &
        " AND ON_OPEN = 0" &
        " AND SHIPPED = 0" &
        " AND ORDR = 0" &
        " AND CANCL = 0" &
        " AND RECD = 0"
        For Each rowSOWMRP99 As DataRow In dst.Tables("SOWMRP99").Select(Filter)
            rowSOWMRP99.Delete()
        Next

        If chkMAKEFLAT.Checked Then
            chkMAKEFLAT.Visible = True
            txtCrop.Visible = True
            lblCrop.Visible = True
            UltraTabControl2.Tabs.Item("Flat").Visible = True
            MakeFlatData()
            grdSOWMRP9X.DataSource = dst.Tables.Item("SOWMRP9X")
        Else
            chkMAKEFLAT.Visible = False
            txtCrop.Visible = False
            lblCrop.Visible = False
            txtCrop.Text = 0
            UltraTabControl2.Tabs.Item("Flat").Visible = False
        End If
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        Dim z As String = ""
        CR_params.Add("SHOW_SHIP", IIf(Absx1.chkFor("CHKSHIPPED").Checked, "1", "0"))
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Sub Import_SOTCSTY1()

        If dst.Tables.Contains("SOTCSTY1") Then
            dst.Tables("SOTCSTY1").Rows.Clear()
        Else
            ASCMAIN1.sql = "SELECT * FROM SOTCSTY1 WHERE ROWNUM < 1"
            Create_TDA(dst.Tables.Add, "SOTCSTY1", "**", 0, True, "", 0)
        End If

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then

            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
            Dim range As SpreadsheetGear.IRange = Nothing

            Dim ACCT_CODEs As New Dictionary(Of String, Decimal)
            Dim CASHYNs As New Dictionary(Of String, Decimal)
            Dim r As Integer = 1
            Dim recordsImported As Integer = 0
            Dim dupRecords As Integer = 0
            Dim dupMsg As String = ""

            Do While oSheet.Cells(r, 0).Value & "" <> ""
                Dim CUST_STYLE_CODE As String = oSheet.Cells(r, 0).Value & ""
                Dim UPC_CODE As String = oSheet.Cells(r, 2).Value & ""

                Dim SOTCSTY1_SQL As String = "SELECT ICTSTYL1.CUST_CODE, '" & CUST_STYLE_CODE & "' CUST_STYLE_CODE" _
                        & ", ICTSTYC3.SIZE_CODE, ICTSTYC4.STYLE_CODE, ICTSTYC4.COLOR_CODE, ICTSTYC4.UPC_CODE" _
                        & ", '" & ASCMAIN1.USER_ID & "' INIT_OPER, SYSDATE INIT_DATE" _
                        & " FROM ICTSTYC4,ICTSTYC3, ICTSTYL1" _
                        & "  WHERE ICTSTYC4.UPC_CODE LIKE '%" & UPC_CODE & "%'" _
                        & " AND ICTSTYC3.STYLE_CODE = ICTSTYC4.STYLE_CODE" _
                        & " AND ICTSTYC3.COLOR_CODE = ICTSTYC4.COLOR_CODE" _
                        & " AND ICTSTYC3.SIZE_INDEX = ICTSTYC4.SIZE_INDEX" _
                        & " AND ICTSTYL1.STYLE_CODE = ICTSTYC4.STYLE_CODE"
                Dim row As DataRow = ASCDATA1.GetDataRow(SOTCSTY1_SQL)
                If row IsNot Nothing Then
                    Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
                    Dim dupCheckSql As String = "SELECT * FROM SOTCSTY1 WHERE CUST_CODE = :PARM1 AND CUST_STYLE_CODE = :PARM1"
                    Dim rowDupCheck As DataRow = ASCDATA1.GetDataRow(dupCheckSql, "VV", New String() {CUST_CODE, CUST_STYLE_CODE})
                    If rowDupCheck Is Nothing Then
                        Dim rowSOWCSTY1 As DataRow = dst.Tables("SOTCSTY1").NewRow
                        With rowSOWCSTY1
                            .Item("CUST_CODE") = CUST_CODE
                            .Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
                            .Item("SIZE_DESC") = row.Item("SIZE_CODE") & ""
                            .Item("STYLE_CODE") = row.Item("STYLE_CODE") & ""
                            .Item("COLOR_CODE") = row.Item("COLOR_CODE") & ""
                            .Item("CUST_UPC") = row.Item("UPC_CODE") & ""
                            .Item("INIT_OPER") = row.Item("INIT_OPER") & ""
                            .Item("INIT_DATE") = row.Item("INIT_DATE") & ""
                        End With
                        dst.Tables("SOTCSTY1").Rows.Add(rowSOWCSTY1)
                        recordsImported += 1
                    Else
                        'duplicate record
                        dupRecords += 1
                    End If
                End If

                r += 1
            Loop

            If dupRecords > 0 Then
                dupMsg = vbCrLf & dupRecords & " Duplicate Records Skipped."
            End If

            BeginTrans()
            Update_Record_TDA("SOTCSTY1")
            MsgBox(recordsImported & " Records Imported." & dupMsg, MsgBoxStyle.OkOnly, "Success")
            CommitTrans()

        End If



    End Sub

    Private Sub cmdImport_Click(sender As Object, e As EventArgs)
        Import_SOTCSTY1()
    End Sub

    Private Sub MakeFlatData()
        Dim CURDATE As Date = DateSerial(Now().Year, Now().Month, 1)

        Dim WP1 As Date = CURDATE.AddMonths(1)
        Dim WP1N As String = MonthName(WP1.Month)
        Dim WP2 As Date = CURDATE.AddMonths(2)
        Dim WP2N As String = MonthName(WP2.Month)
        Dim WP3 As Date = CURDATE.AddMonths(3)
        Dim WP3N As String = MonthName(WP3.Month)
        Dim WP4 As Date = CURDATE.AddMonths(4)
        Dim WP4N As String = MonthName(WP4.Month)

        Dim SP1 As Date = CURDATE.AddMonths(-1)
        Dim SP1N As String = MonthName(SP1.Month)
        Dim SP2 As Date = CURDATE.AddMonths(-2)
        Dim SP2N As String = MonthName(SP2.Month)
        Dim SP3 As Date = CURDATE.AddMonths(-3)
        Dim SP3N As String = MonthName(SP3.Month)
        Dim SP4 As Date = CURDATE.AddMonths(-4)
        Dim SP4N As String = MonthName(SP4.Month)

        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_1").Header.Caption = MonthName(Now().Month)
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_1").Header.Appearance.BackColor = Drawing.Color.LightBlue
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_2").Header.Caption = WP1N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_2").Header.Appearance.BackColor = Drawing.Color.LightBlue
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_3").Header.Caption = WP2N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_3").Header.Appearance.BackColor = Drawing.Color.LightBlue
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_4").Header.Caption = WP3N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_4").Header.Appearance.BackColor = Drawing.Color.LightBlue
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_5").Header.Caption = WP4N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("WIP_5").Header.Appearance.BackColor = Drawing.Color.LightBlue

        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_1").Header.Caption = MonthName(Now().Month)
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_1").Header.Appearance.BackColor = Drawing.Color.LightGreen
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_2").Header.Caption = SP1N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_2").Header.Appearance.BackColor = Drawing.Color.LightGreen
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_3").Header.Caption = SP2N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_3").Header.Appearance.BackColor = Drawing.Color.LightGreen
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_4").Header.Caption = SP3N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_4").Header.Appearance.BackColor = Drawing.Color.LightGreen
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_5").Header.Caption = SP4N
        grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("SHIPPED_5").Header.Appearance.BackColor = Drawing.Color.LightGreen

        If Val(txtCrop.Text) > 0 Then
            grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("VENDOR_STOCK_NO").Hidden = True
            grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("CUST_STYLE_CODE").Hidden = True
        Else
            grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("VENDOR_STOCK_NO").Hidden = False
            grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("CUST_STYLE_CODE").Hidden = False
        End If

        Dim CUST_CODE As String = ""
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        Dim VENDOR_STOCK_NO As String = ""
        Dim CUST_STYLE_CODE As String = ""
        Dim rowSOWMRP9X As DataRow = Nothing

        Dim SORT_ORDER As String = "CUST_CODE, STYLE_CODE, COLOR_CODE, VENDOR_STOCK_NO, CUST_STYLE_CODE"
        For Each rowSOWMRP99 As DataRow In dst.Tables("SOWMRP99").Select("", SORT_ORDER)
            Dim STYLE_CODE_CROP As String = (rowSOWMRP99.Item("STYLE_CODE").ToString & String.Empty).Substring(0, (rowSOWMRP99.Item("STYLE_CODE").ToString & String.Empty).Length - Val(txtCrop.Text))
            Dim VENDOR_STOCK_NO_CROP = ""
            Dim CUST_STYLE_CODE_CROP = ""
            If Val(txtCrop.Text) = 0 Then
                VENDOR_STOCK_NO_CROP = rowSOWMRP99.Item("VENDOR_STOCK_NO").ToString & String.Empty
                CUST_STYLE_CODE_CROP = rowSOWMRP99.Item("CUST_STYLE_CODE").ToString & String.Empty
            End If
            If CUST_CODE <> rowSOWMRP99.Item("CUST_CODE").ToString & String.Empty _
                Or STYLE_CODE <> STYLE_CODE_CROP _
                Or COLOR_CODE <> rowSOWMRP99.Item("COLOR_CODE").ToString & String.Empty _
                Or VENDOR_STOCK_NO <> VENDOR_STOCK_NO_CROP _
                Or CUST_STYLE_CODE <> CUST_STYLE_CODE_CROP Then
                If Not IsNothing(rowSOWMRP9X) Then
                    dst.Tables.Item("SOWMRP9X").Rows.Add(rowSOWMRP9X)
                End If
                rowSOWMRP9X = dst.Tables.Item("SOWMRP9X").NewRow
                CUST_CODE = rowSOWMRP99.Item("CUST_CODE").ToString & String.Empty
                STYLE_CODE = STYLE_CODE_CROP
                If Val(txtCrop.Text) > 0 Then
                    VENDOR_STOCK_NO = ""
                    CUST_STYLE_CODE = ""
                    grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("VENDOR_STOCK_NO").Hidden = True
                    grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("CUST_STYLE_CODE").Hidden = True
                Else
                    grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("VENDOR_STOCK_NO").Hidden = False
                    grdSOWMRP9X.DisplayLayout.Bands(0).Columns.Item("CUST_STYLE_CODE").Hidden = False
                End If
                COLOR_CODE = rowSOWMRP99.Item("COLOR_CODE").ToString & String.Empty
                VENDOR_STOCK_NO = VENDOR_STOCK_NO_CROP
                CUST_STYLE_CODE = CUST_STYLE_CODE_CROP
                rowSOWMRP9X.Item("CUST_CODE") = CUST_CODE
                rowSOWMRP9X.Item("STYLE_CODE") = STYLE_CODE_CROP
                rowSOWMRP9X.Item("COLOR_CODE") = COLOR_CODE
                rowSOWMRP9X.Item("VENDOR_STOCK_NO") = VENDOR_STOCK_NO
                rowSOWMRP9X.Item("CUST_STYLE_CODE") = CUST_STYLE_CODE
            End If
            Select Case rowSOWMRP99.Item("REC_TYPE").ToString & String.Empty
                Case "W" 'WIP / In Transit
                    If Val(rowSOWMRP99.Item("TRAN").ToString & String.Empty) > 0 Then
                        If IsDate(rowSOWMRP99.Item("X_DATE").ToString & String.Empty) Then
                            rowSOWMRP9X.Item("TRAN_DATE") = CDate(rowSOWMRP99.Item("X_DATE").ToString & String.Empty)
                        End If
                        rowSOWMRP9X.Item("TRAN") = Val(rowSOWMRP9X.Item("TRAN").ToString & String.Empty) + Val(rowSOWMRP99.Item("TRAN").ToString & String.Empty)
                    End If
                    If Val(rowSOWMRP99.Item("WIP").ToString & String.Empty) > 0 Then
                        If IsDate(rowSOWMRP99.Item("X_DATE").ToString & String.Empty) Then
                            Dim X_DATE As Date = CDate(CDate(rowSOWMRP99.Item("X_DATE").ToString & String.Empty).ToShortDateString)
                            If X_DATE >= WP4 Then
                                rowSOWMRP9X.Item("WIP_5") = Val(rowSOWMRP9X.Item("WIP_5").ToString & String.Empty) + Val(rowSOWMRP99.Item("WIP").ToString & String.Empty)
                            Else
                                If X_DATE >= WP3 Then
                                    rowSOWMRP9X.Item("WIP_4") = Val(rowSOWMRP9X.Item("WIP_4").ToString & String.Empty) + Val(rowSOWMRP99.Item("WIP").ToString & String.Empty)
                                Else
                                    If X_DATE >= WP2 Then
                                        rowSOWMRP9X.Item("WIP_3") = Val(rowSOWMRP9X.Item("WIP_3").ToString & String.Empty) + Val(rowSOWMRP99.Item("WIP").ToString & String.Empty)
                                    Else
                                        If X_DATE >= WP1 Then
                                            rowSOWMRP9X.Item("WIP_2") = Val(rowSOWMRP9X.Item("WIP_2").ToString & String.Empty) + Val(rowSOWMRP99.Item("WIP").ToString & String.Empty)
                                        Else
                                            rowSOWMRP9X.Item("WIP_1") = Val(rowSOWMRP9X.Item("WIP_1").ToString & String.Empty) + Val(rowSOWMRP99.Item("WIP").ToString & String.Empty)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Case "H" 'On Hand
                    rowSOWMRP9X.Item("ON_HAND") = Val(rowSOWMRP9X.Item("ON_HAND").ToString & String.Empty) + Val(rowSOWMRP99.Item("ON_HAND").ToString & String.Empty)
                Case "P" 'In Pick
                    rowSOWMRP9X.Item("ON_PICK") = Val(rowSOWMRP9X.Item("ON_PICK").ToString & String.Empty) + Val(rowSOWMRP99.Item("ON_PICK").ToString & String.Empty)
                Case "V" 'Open
                    rowSOWMRP9X.Item("ON_OPEN") = Val(rowSOWMRP9X.Item("ON_OPEN").ToString & String.Empty) + Val(rowSOWMRP99.Item("ON_OPEN").ToString & String.Empty)
                Case "X" 'Ordered / Cancelled
                    rowSOWMRP9X.Item("ORDR") = Val(rowSOWMRP9X.Item("ORDR").ToString & String.Empty) + Val(rowSOWMRP99.Item("ORDR").ToString & String.Empty)
                    rowSOWMRP9X.Item("CANCL") = Val(rowSOWMRP9X.Item("CANCL").ToString & String.Empty) + Val(rowSOWMRP99.Item("CANCL").ToString & String.Empty)
                Case "S" 'Shipped
                    If Val(rowSOWMRP99.Item("SHIPPED").ToString & String.Empty) > 0 Then
                        If IsDate(rowSOWMRP99.Item("X_DATE").ToString & String.Empty) Then
                            Dim X_DATE As Date = CDate(CDate(rowSOWMRP99.Item("X_DATE").ToString & String.Empty).ToShortDateString)
                            If X_DATE < SP3 Then
                                rowSOWMRP9X.Item("SHIPPED_5") = Val(rowSOWMRP9X.Item("SHIPPED_5").ToString & String.Empty) + Val(rowSOWMRP99.Item("SHIPPED").ToString & String.Empty)
                            Else
                                If X_DATE < SP2 Then
                                    rowSOWMRP9X.Item("SHIPPED_4") = Val(rowSOWMRP9X.Item("SHIPPED_4").ToString & String.Empty) + Val(rowSOWMRP99.Item("SHIPPED").ToString & String.Empty)
                                Else
                                    If X_DATE < SP1 Then
                                        rowSOWMRP9X.Item("SHIPPED_3") = Val(rowSOWMRP9X.Item("SHIPPED_3").ToString & String.Empty) + Val(rowSOWMRP99.Item("SHIPPED").ToString & String.Empty)
                                    Else
                                        If X_DATE < DateSerial(Now().Year, Now().Month, 1) Then
                                            rowSOWMRP9X.Item("SHIPPED_2") = Val(rowSOWMRP9X.Item("SHIPPED_2").ToString & String.Empty) + Val(rowSOWMRP99.Item("SHIPPED").ToString & String.Empty)
                                        Else
                                            rowSOWMRP9X.Item("SHIPPED_1") = Val(rowSOWMRP9X.Item("SHIPPED_1").ToString & String.Empty) + Val(rowSOWMRP99.Item("SHIPPED").ToString & String.Empty)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
            End Select
        Next
        If Not IsNothing(rowSOWMRP9X) Then
            dst.Tables.Item("SOWMRP9X").Rows.Add(rowSOWMRP9X)
        End If
    End Sub

    Private Sub chkMAKEFLAT_CheckedChanged(sender As Object, e As EventArgs) Handles chkMAKEFLAT.CheckedChanged
        If chkMAKEFLAT.Checked Then
            txtCrop.Visible = True
            lblCrop.Visible = True
        Else
            txtCrop.Visible = False
            lblCrop.Visible = False
            txtCrop.Text = 0
        End If
    End Sub

    Private Sub chkMAKEFLAT_STATE()
        If chkMAKEFLAT.Checked Then
            UltraTabControl2.Tabs.Item("Flat").Visible = True
        Else
            UltraTabControl2.Tabs.Item("Flat").Visible = False
        End If
    End Sub

    Private Sub txtCrop_TextChanged(sender As Object, e As EventArgs) Handles txtCrop.TextChanged
        If Not IsNumeric(txtCrop.Text) Then
            txtCrop.Text = 0
        End If
    End Sub
End Class