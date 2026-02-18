Imports Infragistics.Win.UltraWinGrid

Public Class ARFECOMS

    Dim ARTCUSTX As String
    Dim CUST_CODE As String

    Dim sqlSOTORDR1 As String = ""
    Dim sqlARTPYMTS As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            Create_ARTCUSTX(True)

            ASCMAIN1.sql = "Select
SOTORDR1.ORDR_NO,ORDR_DATE,SOTORDR1.ORDR_CUST_PO,SOTORDR1.SHIP_VIA_CODE,SOTORDR1.INIT_DATE,SOTORDR1.LAST_DATE,SOTORDR1.ORDR_SOURCE,
SOTORDR1.FRT_TERMS,SOTORDR1.ORDR_DATE_BOOKED,SOTORDR1.ORDR_YYYYPP_BOOKED,SOTORDR1.ORDR_STATUS,SOTORDR1.ORDR_HOLD,SOTORDR1.ORDR_NO_WEB,
SOTORDR1.ECOM_CODE,SOTORDR1.ORDR_BUYER_NAME,SOTORDR1.ORDR_BUYER_EMAIL,SOTORDR1.ORDR_WEB_IND,SOTORDR1.ORDR_WEB_ID,
SOTORDR2S.ORDR_QTY,SOTORDR2S.ORDR_QTY_OPEN,SOTORDR2S.ORDR_QTY_PICK,SOTORDR2S.ORDR_QTY_SHIP,SOTORDR2S.ORDR_QTY_CANC,
SOTORDR2S.ORDR_AMT,SOTORDR2S.ORDR_AMT_OPEN,SOTORDR2S.ORDR_AMT_PICK,SOTORDR2S.ORDR_AMT_SHIP,SOTORDR2S.ORDR_AMT_CANC,
SOTORDR2S.ORDR_GRS_AMT,SOTORDR2S.ORDR_GRS_AMT_OPEN,SOTORDR2S.ORDR_GRS_AMT_PICK,SOTORDR2S.ORDR_GRS_AMT_SHIP,SOTORDR2S.ORDR_GRS_AMT_CANC
from SOTORDR1,(Select SOTORDR2.ORDR_NO
, SUM (SOTORDR2.ORDR_QTY) ORDR_QTY
, SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK
, SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP, SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC
, SUM (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT
, SUM (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN, SUM (SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_PICK
, SUM (SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP, SUM (SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_CANC
, SUM (SOTORDR2.ORDR_QTY * NVL(SOTORDR2.ORDR_RETAIL_PRICE,SOTORDR2.ORDR_UNIT_PRICE)) ORDR_GRS_AMT
, SUM (SOTORDR2.ORDR_QTY_OPEN * NVL(SOTORDR2.ORDR_RETAIL_PRICE,SOTORDR2.ORDR_UNIT_PRICE)) ORDR_GRS_AMT_OPEN
, SUM (SOTORDR2.ORDR_QTY_PICK * NVL(SOTORDR2.ORDR_RETAIL_PRICE,SOTORDR2.ORDR_UNIT_PRICE)) ORDR_GRS_AMT_PICK
, SUM (SOTORDR2.ORDR_QTY_SHIP * NVL(SOTORDR2.ORDR_RETAIL_PRICE,SOTORDR2.ORDR_UNIT_PRICE)) ORDR_GRS_AMT_SHIP
, SUM (SOTORDR2.ORDR_QTY_CANC * NVL(SOTORDR2.ORDR_RETAIL_PRICE,SOTORDR2.ORDR_UNIT_PRICE)) ORDR_GRS_AMT_CANC
from SOTORDR2
group by SOTORDR2.ORDR_NO) SOTORDR2S
where SOTORDR2S.ORDR_NO = SOTORDR1.ORDR_NO"
            sqlSOTORDR1 = ASCMAIN1.sql
            ASCMAIN1.sql &= " AND ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)
            With .Tables("SOTORDR1").Columns
                .Add("ORDR_DSC_AMT", GetType(System.Decimal), "ISNULL(ORDR_GRS_AMT,0) - ISNULL(ORDR_AMT,0)")
                .Add("ORDR_DSC_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_GRS_AMT_OPEN,0) - ISNULL(ORDR_AMT_OPEN,0)")
                .Add("ORDR_DSC_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_GRS_AMT_PICK,0) - ISNULL(ORDR_AMT_PICK,0)")
                .Add("ORDR_DSC_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_GRS_AMT_SHIP,0) - ISNULL(ORDR_AMT_SHIP,0)")
                .Add("ORDR_DSC_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_GRS_AMT_CANC,0) - ISNULL(ORDR_AMT_CANC,0)")
            End With

            ASCMAIN1.sql = "Select
SOTORDR2.ORDR_NO,SOTORDR2.ORDR_LNO,SOTORDR2.STYLE_CODE,SOTORDR2.COLOR_CODE,SOTORDR2.STYLE_DESC,SOTORDR2.ORDR_EXTD_COST,SOTORDR2.ORDR_UNIT_PRICE,
SOTORDR2.ORDR_QTY,SOTORDR2.ORDR_QTY_OPEN,SOTORDR2.ORDR_QTY_PICK,SOTORDR2.ORDR_QTY_SHIP,SOTORDR2.ORDR_QTY_CANC,SOTORDR2.ORDR_STATUS,SOTORDR2.CUST_UPC,
SOTORDR2.STYLE_PRICE,SOTORDR2.STYLE_RETAIL,SOTORDR2.ORDR_RETAIL_PRICE,SOTORDR2.ORDR_FULLFILL_FEE,SOTORDR2.ORDR_SELLER_FEE,SOTORDR2.PARTNER_LN_ID,
ICTCOLR1.COLOR_DESC
from SOTORDR2,ICTCOLR1 where ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE and SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "V", 2)
            With .Tables("SOTORDR2").Columns

                .Add("ORDR_UNIT_DISC", GetType(System.Decimal), "ISNULL(ORDR_RETAIL_PRICE,0) - ISNULL(ORDR_UNIT_PRICE,0)")

                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY,0)")
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY_OPEN,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY_PICK,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY_SHIP,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE,0) * ISNULL(ORDR_QTY_CANC,0)")

                .Add("ORDR_DSC", GetType(System.Decimal), "ISNULL(ORDR_UNIT_DISC,0) * ISNULL(ORDR_QTY,0)")
                .Add("ORDR_DSC_OPEN", GetType(System.Decimal), "ISNULL(ORDR_UNIT_DISC,0) * ISNULL(ORDR_QTY_OPEN,0)")
                .Add("ORDR_DSC_PICK", GetType(System.Decimal), "ISNULL(ORDR_UNIT_DISC,0) * ISNULL(ORDR_QTY_PICK,0)")
                .Add("ORDR_DSC_SHIP", GetType(System.Decimal), "ISNULL(ORDR_UNIT_DISC,0) * ISNULL(ORDR_QTY_SHIP,0)")
                .Add("ORDR_DSC_CANC", GetType(System.Decimal), "ISNULL(ORDR_UNIT_DISC,0) * ISNULL(ORDR_QTY_CANC,0)")

                .Add("ORDR_GRS", GetType(System.Decimal), "ORDR_AMT + ORDR_DSC")
                .Add("ORDR_GRS_OPEN", GetType(System.Decimal), "ORDR_AMT_OPEN + ORDR_DSC_OPEN")
                .Add("ORDR_GRS_PICK", GetType(System.Decimal), "ORDR_AMT_PICK + ORDR_DSC_PICK")
                .Add("ORDR_GRS_SHIP", GetType(System.Decimal), "ORDR_AMT_SHIP + ORDR_DSC_SHIP")
                .Add("ORDR_GRS_CANC", GetType(System.Decimal), "ORDR_AMT_CANC + ORDR_DSC_CANC")
            End With


            ASCMAIN1.sql = "Select * from " & ARTCUSTX
            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False, "", 1)

            With .Tables.Add("ARTCUSTX_DASH")
                .Columns.Add("TYPE")
                .PrimaryKey = New DataColumn() { .Columns("TYPE")}
            End With

            For Each TYPE As String In New String() {"OPEN", "BOOK", "SHIP", "RTRN", "PYMT"}
                Dim rowARTCUSTX_DASH As DataRow = dst.Tables("ARTCUSTX_DASH").NewRow
                rowARTCUSTX_DASH.Item("TYPE") = TYPE
                dst.Tables("ARTCUSTX_DASH").Rows.Add(rowARTCUSTX_DASH)

                With .Tables.Add($"ARTCUSTX_{TYPE}")
                    .Columns.Add("TYPE")
                    .Columns.Add("YP")
                    .Columns.Add("CNT", GetType(System.Int32))

                    If TYPE = "PYMT" Then
                        .Columns.Add("PYMT", GetType(System.Decimal))
                        .Columns.Add("APPL", GetType(System.Decimal))
                        .Columns.Add("FEE", GetType(System.Decimal))
                        .Columns.Add("OTHER", GetType(System.Decimal))
                        .Columns.Add("GL", GetType(System.Decimal))
                        .Columns.Add("CBOA", GetType(System.Decimal))
                        .Columns.Add("TB", GetType(System.Decimal))
                    Else

                        .Columns.Add("QTY", GetType(System.Int32))
                        .Columns.Add("AMT", GetType(System.Decimal))
                        If TYPE = "OPEN" Then
                            .Columns.Add("MIN_DATE", GetType(System.DateTime))
                            .Columns.Add("MAX_DATE", GetType(System.DateTime))
                        ElseIf TYPE = "BOOK" Then
                            .Columns.Add("CNT_OPEN", GetType(System.Int32))
                            .Columns.Add("CNT_PICK", GetType(System.Int32))
                            .Columns.Add("CNT_SHIP", GetType(System.Int32))
                            .Columns.Add("CNT_CANC", GetType(System.Int32))
                        ElseIf TYPE = "SHIP" Then
                            .Columns.Add("GRS", GetType(System.Decimal))
                            .Columns.Add("DSC", GetType(System.Decimal))
                            .Columns.Add("UNPAID", GetType(System.Int32))
                        ElseIf TYPE = "RTRN" Then
                            .Columns.Add("GRS", GetType(System.Decimal))
                            .Columns.Add("DSC", GetType(System.Decimal))
                            .Columns.Add("UNPAID", GetType(System.Int32))
                        End If
                    End If
                    .PrimaryKey = New DataColumn() { .Columns("TYPE"), .Columns("YP")}
                End With
                Create_Relation("ARTCUSTX_DASH", $"ARTCUSTX_{TYPE}", "TYPE")
            Next

            ASCMAIN1.sql = "SELECT ':PARM1' YP, RECS, P2.PYMT, P3.APPL, P5_0.FEE, P5_0.OTHER, P4.GL, P5_1.CBOA" & vbCrLf _
                & ", NVL(PYMT,0) + NVL(FEE,0) + NVL(OTHER,0) + NVL(GL,0) + NVL(CBOA,0) - NVL(APPL,0) TB " & vbCrLf _
                & "FROM " & vbCrLf _
                & "(" & vbCrLf _
                & "SELECT 'X' KEY, COUNT (*) RECS, SUM (ARTPYMT2.CUST_PYMT_AMT) PYMT FROM ARTPYMT1,ARTPYMT2" & vbCrLf _
                & "WHERE ARTPYMT1.OPS_YYYYPP = ':PARM1'" & vbCrLf _
                & "AND ARTPYMT2.CUST_CODE = ':PARM2'" & vbCrLf _
                & "AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT2.PYMT_STATUS = '2'" & vbCrLf _
                & ") P2, (" & vbCrLf _
                & "SELECT 'X' KEY, SUM (ARTPYMT3.INV_PMT) APPL FROM ARTPYMT1,ARTPYMT2,ARTPYMT3" & vbCrLf _
                & "WHERE ARTPYMT1.OPS_YYYYPP = ':PARM1'" & vbCrLf _
                & "AND ARTPYMT2.CUST_CODE = ':PARM2'" & vbCrLf _
                & "AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT2.PYMT_STATUS = '2'" & vbCrLf _
                & "AND ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
                & ") P3, (" & vbCrLf _
                & "SELECT 'X' KEY, SUM (CASE WHEN REASON_CODE = 'SHOPF' THEN DED ELSE 0 END) FEE" & vbCrLf _
                & ", SUM (CASE WHEN REASON_CODE = 'SHOPF' THEN 0 ELSE DED END) OTHER" & vbCrLf _
                & "FROM (" & vbCrLf _
                & "SELECT ARTPYMT5.REASON_CODE, SUM (GL_DIST_AMT) DED" & vbCrLf _
                & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT5" & vbCrLf _
                & "WHERE ARTPYMT1.OPS_YYYYPP = ':PARM1'" & vbCrLf _
                & "AND ARTPYMT2.CUST_CODE = ':PARM2'" & vbCrLf _
                & "AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT2.PYMT_STATUS = '2'" & vbCrLf _
                & "AND ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
                & "AND NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" & vbCrLf _
                & "GROUP BY ARTPYMT5.REASON_CODE)" & vbCrLf _
                & ") P5_0, (" & vbCrLf _
                & "SELECT 'X' KEY, SUM (CBOA) CBOA FROM (" & vbCrLf _
                & "SELECT ARTPYMT5.REASON_CODE, SUM (GL_DIST_AMT) CBOA" & vbCrLf _
                & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT5" & vbCrLf _
                & "WHERE ARTPYMT1.OPS_YYYYPP = ':PARM1'" & vbCrLf _
                & "AND ARTPYMT2.CUST_CODE = ':PARM2'" & vbCrLf _
                & "AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT2.PYMT_STATUS = '2'" & vbCrLf _
                & "AND ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
                & "AND NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '1'" & vbCrLf _
                & "GROUP BY ARTPYMT5.REASON_CODE)" & vbCrLf _
                & ") P5_1, (" & vbCrLf _
                & "SELECT 'X' KEY, SUM (GL) GL FROM (" & vbCrLf _
                & "SELECT ARTPYMT4.ACCT_CODE, SUM (GL_DIST_AMT) GL" & vbCrLf _
                & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT4" & vbCrLf _
                & "WHERE ARTPYMT1.OPS_YYYYPP = ':PARM1'" & vbCrLf _
                & "AND ARTPYMT2.CUST_CODE = ':PARM2'" & vbCrLf _
                & "AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT2.PYMT_STATUS = '2'" & vbCrLf _
                & "AND ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                & "AND ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
                & "GROUP BY ARTPYMT4.ACCT_CODE)" & vbCrLf _
                & ") P4" & vbCrLf _
                & "WHERE P3.KEY (+) = P2.KEY AND P5_0.KEY = P2.KEY AND P5_1.KEY = P2.KEY AND P4.KEY (+) = P2.KEY"
            sqlARTPYMTS = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "ARTPYMTS", "**", 0, False, "", 1)
            With .Tables("ARTPYMTS")
                .Columns("YP").MaxLength = 6
            End With

            With .Tables.Add("ARTCUSTX_RF")
                .Columns.Add("YP")
                .Columns.Add("BEG_BAL", GetType(System.Decimal))
                .Columns.Add("GRS_SLS", GetType(System.Decimal))
                .Columns.Add("DISC", GetType(System.Decimal))
                .Columns.Add("NET_SLS", GetType(System.Decimal))
                .Columns.Add("SLS_TAX", GetType(System.Decimal))
                .Columns.Add("FRT_CHG", GetType(System.Decimal))
                .Columns.Add("MISC_CHG", GetType(System.Decimal))
                .Columns.Add("NET_AMT", GetType(System.Decimal))
                .Columns.Add("GRS_PMT", GetType(System.Decimal))
                .Columns.Add("FEE_AMT", GetType(System.Decimal))
                .Columns.Add("OTH_DED", GetType(System.Decimal))
                .Columns.Add("NET_PMT", GetType(System.Decimal))
                .Columns.Add("CGS", GetType(System.Decimal))
                .Columns.Add("GRS_PRF", GetType(System.Decimal), "ISNULL(NET_SLS,0) - ISNULL(CGS,0)")
                .Columns.Add("NET_PRF", GetType(System.Decimal), "ISNULL(GRS_PRF,0) - ISNULL(FEE_AMT,0) - ISNULL(OTH_DED,0)")
                .Columns.Add("END_BAL", GetType(System.Decimal))
                .Columns.Add("OOBAL", GetType(System.Decimal), "ISNULL(BEG_BAL,0)+ISNULL(NET_AMT,0)-ISNULL(NET_PMT,0)-ISNULL(END_BAL,0)")
                .PrimaryKey = New DataColumn() { .Columns("YP")}
            End With

            ASCMAIN1.sql = $"Select * from GLTPARM2 where OPS_YYYYPP between '{ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)}' and '{ASCMAIN1.CYP}'"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "", 1)
            Fill_Records("GLTPARM2")



            ASCMAIN1.sql = "Select
SOTINVH1.CUST_CODE, SOTINVH1.ORDR_YYYYPP_UPDATED YP
, SUM (INV_SALES) INV_SALES, SUM (INV_FREIGHT) INV_FREIGHT, SUM (INV_MISC_CHG) INV_MISC_CHG, SUM (INV_STAX) INV_STAX
, SUM (INV_TOTAL_AMOUNT) INV_TOTAL_AMOUNT, SUM (INV_COGS) INV_COGS
from SOTINVH1 where CUST_CODE = :PARM1 and ORDR_YYYYPP_UPDATED = :PARM2
group by SOTINVH1.CUST_CODE, SOTINVH1.ORDR_YYYYPP_UPDATED"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "VV", 2)

        End With

        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        grdARTCUSTX.DataSource = dst.Tables("ARTCUSTX")
        grdARTCUSTX_DASH.DataSource = dst.Tables("ARTCUSTX_DASH")
        grdARTCUSTX_RF.DataSource = dst.Tables("ARTCUSTX_RF")

        With grdSOTORDR1.DisplayLayout.Bands(0)
            .Columns("ORDR_NO").Header.Fixed = True
            .Columns("ORDR_DATE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
        End With


        With grdSOTORDR2.DisplayLayout.Bands(0)
            .Columns("ORDR_NO").Hidden = True
            .Columns("ORDR_LNO").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
        End With

        With grdARTCUSTX_RF.DisplayLayout.Bands(0)
            .Columns("YP").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"YP"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"BEG_BAL", "END_BAL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"GRS_SLS", "DISC", "NET_SLS", "SLS_TAX", "FRT_CHG", "MISC_CHG", "NET_AMT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"GRS_PMT", "FEE_AMT", "OTH_DED", "NET_PMT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"GRS_PRF", "CGS", "NET_PRF"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Violet
                End If
            Next
        End With



        Dim g As UltraWinGrid.UltraGridGroup

        grdARTCUSTX.DisplayLayout.UseFixedHeaders = True
        With grdARTCUSTX.DisplayLayout.Bands(0)

            Dim g_CUST As UltraWinGrid.UltraGridGroup = .Groups.Add("CUST")
            g_CUST.Header.Caption = "Customer Info"
            With g_CUST.Header.Appearance
                .BackColor = Drawing.Color.White
                '.BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = Drawing.Color.LightGray
            End With

            Dim g_OPEN As UltraWinGrid.UltraGridGroup = .Groups.Add("OPEN")
            g_OPEN.Header.Caption = "Open Orders"
            With g_OPEN.Header.Appearance
                .BackColor = Drawing.Color.White
                '.BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = Drawing.Color.Orange
            End With

            Dim g_BOOK As UltraWinGrid.UltraGridGroup = .Groups.Add("BOOK")
            g_BOOK.Header.Caption = "Booked, This MTD"
            With g_BOOK.Header.Appearance
                .BackColor = Drawing.Color.White
                '.BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = Drawing.Color.LightGreen
            End With

            Dim g_SHIP As UltraWinGrid.UltraGridGroup = .Groups.Add("SHIP")
            g_SHIP.Header.Caption = "Shipped, This MTD"
            With g_SHIP.Header.Appearance
                .BackColor = Drawing.Color.White
                '.BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = Drawing.Color.LightBlue
            End With

            Dim g_RTRN As UltraWinGrid.UltraGridGroup = .Groups.Add("RTRN")
            g_RTRN.Header.Caption = "Returned, This MTD"
            With g_RTRN.Header.Appearance
                .BackColor = Drawing.Color.White
                '.BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = Drawing.Color.Pink
            End With

            Dim g_AR As UltraWinGrid.UltraGridGroup = .Groups.Add("AR")
            g_AR.Header.Caption = "Open AR"
            With g_AR.Header.Appearance
                .BackColor = Drawing.Color.White
                '.BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = Drawing.Color.Gold
            End With

            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key.StartsWith("AR_") Then
                    If gcol.Key.EndsWith("_CNT") Then
                        gcol.Header.Caption = "#Invs"
                        gcol.Width = 60
                    ElseIf gcol.Key.EndsWith("_AMT") Then
                        gcol.Header.Caption = "$Amt"
                        gcol.Width = 100
                    End If
                    If gcol.Key.StartsWith("AR_OPEN_") Then
                        gcol.Header.Caption &= " Open"
                    Else
                        gcol.Header.Caption &= " >3d"
                    End If
                ElseIf gcol.Key.EndsWith("_CNT") Then
                    gcol.Header.Caption = "Count"
                    gcol.Width = 60
                ElseIf gcol.Key.EndsWith("_QTY") Then
                    gcol.Header.Caption = "Qty"
                    gcol.Width = 70
                ElseIf gcol.Key.EndsWith("_AMT") Then
                    gcol.Header.Caption = "Net Amt"
                    gcol.Width = 100
                    'ElseIf gcol.Key = "GRS" Then
                    '    gcol.Header.Caption = "Grs Amt"
                    '    gcol.Width = 100
                    'ElseIf gcol.Key = "DSC" Then
                    '    gcol.Header.Caption = "Disc"
                    '    gcol.Width = 80
                    'ElseIf gcol.Key = "UNPAID" Then
                    '    gcol.Header.Caption = "UnPd"
                    '    gcol.Width = 60
                End If

                If gcol.Key.StartsWith("OPEN_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Group = g_OPEN
                ElseIf gcol.Key.StartsWith("BOOK_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Group = g_BOOK
                ElseIf gcol.Key.StartsWith("SHIP_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Group = g_SHIP
                ElseIf gcol.Key.StartsWith("RTRN_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    gcol.Group = g_RTRN
                ElseIf gcol.Key.StartsWith("AR_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    gcol.Group = g_AR
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Group = g_CUST
                End If
            Next
        End With

        For Each B As String In New String() {"OPEN", "BOOK", "SHIP", "RTRN", "PYMT"}
            With grdARTCUSTX_DASH.DisplayLayout.Bands("ARTCUSTX_DASH_ARTCUSTX_" & B)
                .Columns("TYPE").Hidden = True
                .Columns("YP").Width = 70

                .Columns("CNT").Width = 55
                .Columns("CNT").Format = "#,##0"
                .Columns("CNT").Header.Caption = "Count"
                If B = "PYMT" Then

                Else
                    .Columns("QTY").Width = 60
                    .Columns("QTY").Format = "#,##0"
                    .Columns("QTY").Header.Caption = "Qty"
                    .Columns("AMT").Width = 90
                    .Columns("AMT").Format = "#,##0.00"
                    .Columns("AMT").Header.Caption = "Net Amt"
                End If
                If B = "SHIP" Then
                    .Columns("GRS").Width = 90
                    .Columns("GRS").Format = "#,##0.00"
                    .Columns("GRS").Header.Caption = "Grs Amt"
                    .Columns("DSC").Width = 80
                    .Columns("DSC").Format = "#,##0.00"
                    .Columns("DSC").Header.Caption = "Disc"
                    .Columns("UNPAID").Width = 60
                    .Columns("UNPAID").Format = "#,##0"
                    .Columns("UNPAID").Header.Caption = "unpd"
                End If
                If B = "RTRN" Then
                    .Columns("GRS").Width = 90
                    .Columns("GRS").Format = "#,##0.00"
                    .Columns("GRS").Header.Caption = "Grs Amt"
                    .Columns("DSC").Width = 80
                    .Columns("DSC").Format = "#,##0.00"
                    .Columns("DSC").Header.Caption = "Disc"
                    .Columns("UNPAID").Width = 60
                    .Columns("UNPAID").Format = "#,##0"
                    .Columns("UNPAID").Header.Caption = "unpd"
                End If
            End With
        Next

        Create_Summary(grdARTCUSTX, "CUST_CODE", "Count")
        Create_Summary(grdSOTORDR1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDR1, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"})
        Create_Summary(grdSOTORDR1, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})
        Create_Summary(grdSOTORDR1, New String() {"ORDR_GRS_AMT", "ORDR_DSC_AMT"})

        ASCMAIN1.Add_Value_List(grdARTCUSTX_DASH, "TYPE",, New String() {":", "OPEN:Open Orders", "BOOK:Booked Orders", "SHIP:Shipped Orders", "RTRN:Returned Orders", "PYMT:Payments"})

        ' these work, but are unnec because of the YPs_List
        'ASCMAIN1.Add_Value_List(grdARTCUSTX_DASH, "YP",, New String() {":", "000000:Open"}, 1)
        'ASCMAIN1.Add_Value_List(grdARTCUSTX_DASH, "YP",,, 2, $"Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP between '{ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)}' and '{ASCMAIN1.CYP}'")

        Dim YPs_List As New List(Of String)
        YPs_List.Add(":")
        YPs_List.Add("000000:Open")
        For Each rowGLTPARM2 As DataRow In dst.Tables("GLTPARM2").Select("")
            YPs_List.Add($"{rowGLTPARM2.Item("OPS_YYYYPP")}:{Mid(rowGLTPARM2.Item("LEGEND"), 10, 6)}")
        Next
        For b As Integer = 1 To 5
            ASCMAIN1.Add_Value_List(grdARTCUSTX_DASH, "YP",, YPs_List.ToArray, b)
        Next

        ASCMAIN1.Add_Value_List(grdARTCUSTX_RF, "YP",, YPs_List.ToArray)

        ASCMAIN1.Add_Value_List(grdSOTORDR1, "ORDR_STATUS",, New String() {":", "O:Open", "P:Pick", "F:Shipped", "C:Canceled"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Refresh"
                If ScreenMode Then
                    Dim CUST_CODE_SAVE As String = CUST_CODE
                    Click_Command("Done")
                    Absx1.txtFor("CUST_CODE").Text = CUST_CODE_SAVE
                    Click_Command("Load")
                Else
                    Clear_Record()
                End If

            Case "Load"
                EntryMode = "E"
                CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdARTCUSTX.Visible = Not tf
        splMain.Visible = tf

        If ScreenMode Then

        Else
            Clear_Record()
            splMain.Panel2Collapsed = True
        End If

    End Sub

    Sub Clear_Record()

        ' , "ARTCUSTX_DASH"
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ARTCUSTX", "ARTPYMTS" _
                , "ARTCUSTX_OPEN", "ARTCUSTX_BOOK", "ARTCUSTX_SHIP", "ARTCUSTX_RTRN", "ARTCUSTX_PYMT" _
                , "ARTCUSTX_RF", "SOTORDR1", "SOTORDR2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        CUST_CODE = ""

        Create_ARTCUSTX(False)

        EnforceConstraints(False)
        Fill_Records("ARTCUSTX")
        Sort_grdColumns(grdARTCUSTX, "CUST_CODE")
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        CUST_CODE = Absx1.txtFor("CUST_CODE").Text

        ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)
        Get_Data()
        EnforceConstraints(True)

        grdARTCUSTX_DASH.Rows.ExpandAll(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                sql_where = "POST_CODE = 'B2C'"
        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCUSTX, "SSB", "Show Filter", "Show GroupBox", "Customer Inquiry")
        Load_Popup_Menu(grdSOTORDR1, "SSB", "Show Filter", "Show GroupBox", "Sales Order Inquiry")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name
            'Case "grdARTCOLL1"
            '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool
            '    Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden

            Case Else
        End Select
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

            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "Track Shipment"
                If grd.ActiveRow.Cells("SHIP_REF").Text <> "" Then
                    'SOCMAIN1.Track_Shipment(grd.ActiveRow.Cells("SHIP_VIA_CODE").Text, grd.ActiveRow.Cells("SHIP_REF").Text)
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                If ORDR_NO <> "" Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI", "F", "SO")
                End If
        End Select


    End Sub

#End Region

    Sub Create_ARTCUSTX(initialize As Boolean)

        'ASCMAIN1.sql = "Select '000000' YP, ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME from ARTCUST1 where ARTCUST1.POST_CODE = 'B2C'"
        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME from ARTCUST1 where ARTCUST1.POST_CODE = 'B2C'"

        If initialize Or ARTCUSTX = "" Then
            ARTCUSTX = ASCMAIN1.Temp_Table
            For Each COL As String In New String() {"OPEN", "BOOK", "SHIP", "RTRN"}
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_CNT NUMBER (6,0)")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_QTY NUMBER (6,0)")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_AMT NUMBER (12,2)")

                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_MIN_DATE DATE")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_MAX_DATE DATE")

                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_CNT_OPEN NUMBER (6,0)")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_CNT_PICK NUMBER (6,0)")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_CNT_SHIP NUMBER (6,0)")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_CNT_CANC NUMBER (6,0)")

                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_GRS NUMBER (12,2)")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_DSC NUMBER (12,2)")
                ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add {COL}_UNPAID NUMBER (6,0)")

            Next
            ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add AR_OPEN_CNT NUMBER (6,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add AR_OPEN_AMT NUMBER (12,2)")
            ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add AR_OVER_CNT NUMBER (6,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add AR_OVER_AMT NUMBER (12,2)")
            ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add AR_OPEN_CNT_C NUMBER (6,0)")
            ASCDATA1.ExecuteSQL($"Alter Table {ARTCUSTX} Add AR_OPEN_AMT_C NUMBER (12,2)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCUSTX)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCUSTX & " (CUST_CODE, CUST_NAME) " & ASCMAIN1.sql)

            ASCMAIN1.sql = $"
Begin Declare Cursor C1 is 
 Select * from {ARTCUSTX} for Update;
 Begin For R1 in C1 Loop

  Begin Declare Cursor C2 is 
   Select Count(Distinct SOTORDR1.ORDR_NO) OPEN_CNT, Sum (SOTORDR2.ORDR_QTY_OPEN) OPEN_QTY, Sum (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) OPEN_AMT
    , MIN (SOTORDR1.ORDR_DATE) MIN_DATE, MAX (SOTORDR1.ORDR_DATE) MAX_DATE
    from SOTORDR1,SOTORDR2
    where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR1.CUST_CODE = R1.CUST_CODE
      and SOTORDR1.ORDR_STATUS = 'O';
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set OPEN_CNT = R2.OPEN_CNT, OPEN_QTY = R2.OPEN_QTY, OPEN_AMT = R2.OPEN_AMT
    , OPEN_MIN_DATE = R2.MIN_DATE, OPEN_MAX_DATE = R2.MAX_DATE
     where Current of C1;
   End Loop; End;
  End;
  
  {sqlSOTORDR1_BOOK_SHIP_RTRN(ASCMAIN1.CYP)}

 End Loop; End;
End;"

            ASCDATA1.ExecuteSQL()

        End If

    End Sub

    Function sqlSOTORDR1_BOOK_SHIP_RTRN(YP As String)
        Dim sql As String = $"
  Begin Declare Cursor C2 is 
   Select Count(Distinct SOTORDR1.ORDR_NO) BOOK_CNT, Sum (SOTORDR2.ORDR_QTY) BOOK_QTY, Sum (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) BOOK_AMT
    , MIN (SOTORDR1.ORDR_DATE) MIN_DATE, MAX (SOTORDR1.ORDR_DATE) MAX_DATE
    , SUM (CASE WHEN SOTORDR1.ORDR_STATUS = 'O' THEN 1 ELSE 0 END) CNT_OPEN
    , SUM (CASE WHEN SOTORDR1.ORDR_STATUS = 'P' THEN 1 ELSE 0 END) CNT_PICK
    , SUM (CASE WHEN SOTORDR1.ORDR_STATUS = 'F' THEN 1 ELSE 0 END) CNT_SHIP
    , SUM (CASE WHEN SOTORDR1.ORDR_STATUS = 'C' THEN 1 ELSE 0 END) CNT_CANC
    from SOTORDR1,SOTORDR2
    where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR1.CUST_CODE = R1.CUST_CODE
      and SOTORDR1.ORDR_YYYYPP_BOOKED = '{YP}';
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set BOOK_CNT = R2.BOOK_CNT, BOOK_QTY = R2.BOOK_QTY, BOOK_AMT = R2.BOOK_AMT
    , BOOK_MIN_DATE = R2.MIN_DATE, BOOK_MAX_DATE = R2.MAX_DATE
    , BOOK_CNT_OPEN = R2.CNT_OPEN, BOOK_CNT_PICK = R2.CNT_PICK, BOOK_CNT_SHIP = R2.CNT_SHIP, BOOK_CNT_CANC = R2.CNT_CANC
     where Current of C1;
   End Loop; End;
  End;

  Begin Declare Cursor C2 is 
   Select Count(Distinct SOTINVH1.INV_NO) SHIP_CNT, Sum (SOTINVH2.ORDR_QTY_SHIP) SHIP_QTY, Sum (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) SHIP_AMT
    , Sum (SOTINVH2.ORDR_QTY_SHIP * SOTORDR2.ORDR_RETAIL_PRICE) SHIP_GRS
    , Sum (SOTINVH2.ORDR_QTY_SHIP * (SOTORDR2.ORDR_RETAIL_PRICE - SOTORDR2.ORDR_UNIT_PRICE)) SHIP_DSC
    from SOTINVH1,SOTINVH2,SOTORDR2,SOTPICK2
    where SOTINVH1.CUST_CODE = R1.CUST_CODE
AND SOTPICK2.PICK_NO = SOTINVH1.PICK_NO AND SOTPICK2.PICK_LNO = SOTINVH2.INV_LNO
      AND SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO
      and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO
      and SOTINVH1.INV_TYPE = 'I' and SOTINVH1.ORDR_YYYYPP_UPDATED = '{YP}';
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set SHIP_CNT = R2.SHIP_CNT, SHIP_QTY = R2.SHIP_QTY, SHIP_AMT = R2.SHIP_AMT
    , SHIP_GRS = R2.SHIP_GRS, SHIP_DSC = R2.SHIP_DSC
     where Current of C1;
   End Loop; End;
  End;

  Begin Declare Cursor C2 is 
   Select Count(Distinct X.INV_NO) RTRN_CNT, Sum (X.ORDR_QTY_SHIP) RTRN_QTY, Sum (X.ORDR_QTY_SHIP * X.ORDR_UNIT_PRICE) RTRN_AMT
   , Sum (X.ORDR_QTY_SHIP * DECODE(X.ORDR_NO,NULL,SOTORDR2.ORDR_UNIT_PRICE,SOTORDR2.ORDR_RETAIL_PRICE)) RTRN_GRS
   , Sum (X.ORDR_QTY_SHIP * DECODE(X.ORDR_NO,NULL,SOTORDR2.ORDR_UNIT_PRICE,(SOTORDR2.ORDR_RETAIL_PRICE - SOTORDR2.ORDR_UNIT_PRICE))) RTRN_DSC
   from (
    Select SOTINVH1.INV_NO, SOTRTNL1.ORDR_NO, SOTINVH2.INV_LNO ORDR_LNO, SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE
        from SOTINVH1,SOTINVH2,SOTRTNL1
        where SOTINVH1.CUST_CODE =  R1.CUST_CODE
          and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO
          and SOTINVH1.INV_TYPE = 'C' and SOTINVH1.ORDR_YYYYPP_UPDATED = '{YP}'
          AND SOTRTNL1.INV_NO (+) = SOTINVH1.INV_NO
    ) X,SOTORDR2
   where SOTORDR2.ORDR_NO (+) = X.ORDR_NO
     and SOTORDR2.ORDR_LNO (+) = X.ORDR_LNO;
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set RTRN_CNT = R2.RTRN_CNT, RTRN_QTY = R2.RTRN_QTY, RTRN_AMT = R2.RTRN_AMT
    , RTRN_GRS = R2.RTRN_GRS, RTRN_DSC = R2.RTRN_DSC
     where Current of C1;
   End Loop; End;
  End;
"

        If YP = ASCMAIN1.CYP Then
            sql &= $"
  Begin Declare Cursor C2 is 
   Select Count(Distinct ARTOPEN1.INV_NUM) AR_OPEN_CNT, Sum (ARTOPEN1.INV_BALANCE) AR_OPEN_AMT
    , Sum (Case when ARTOPEN1.INV_DATE > TRUNC(SYSDATE-3) then 1 else 0 end) AR_OVER_CNT
    , Sum (Case when ARTOPEN1.INV_DATE > TRUNC(SYSDATE-3) then ARTOPEN1.INV_BALANCE else 0 end) AR_OVER_AMT
    from ARTOPEN1
    where ARTOPEN1.CUST_CODE = R1.CUST_CODE and ARTOPEN1.OPS_YYYYPP <= '{ASCMAIN1.CYP}' and ARTOPEN1.INV_TYPE = 'I';
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set AR_OPEN_CNT = R2.AR_OPEN_CNT, AR_OPEN_AMT = R2.AR_OPEN_AMT, AR_OVER_CNT = R2.AR_OVER_CNT, AR_OVER_AMT = R2.AR_OVER_AMT
     where Current of C1;
   End Loop; End;
  End;
  Begin Declare Cursor C2 is 
   Select Count(Distinct ARTOPEN1.INV_NUM) AR_OPEN_CNT, Sum (ARTOPEN1.INV_BALANCE) AR_OPEN_AMT
    from ARTOPEN1
    where ARTOPEN1.CUST_CODE = R1.CUST_CODE and ARTOPEN1.OPS_YYYYPP <= '{ASCMAIN1.CYP}' and ARTOPEN1.INV_TYPE <> 'I';
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set AR_OPEN_CNT_C = R2.AR_OPEN_CNT, AR_OPEN_AMT_C = R2.AR_OPEN_AMT
     where Current of C1;
   End Loop; End;
  End;
"
        Else
            sql &= $"
  Begin Declare Cursor C2 is 
   Select Count(Distinct ARTOPEN1.INV_NUM) AR_OPEN_CNT, Sum (ARTOPEN1.INV_BALANCE) AR_OPEN_AMT
    from (Select DETL_CVX_NO CUST_CODE, DETL_CVX_TYPE INV_TYPE, DETL_CTL_NO INV_NUM, CREC_AMT INV_BALANCE
           from GLTCREC3 where OPS_YYYYPP = '{YP}' and DETL_CTL_TYPE = 'I' and CREC_TYPE_CODE = 'AR') ARTOPEN1
    where ARTOPEN1.CUST_CODE = R1.CUST_CODE;
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set AR_OPEN_CNT = R2.AR_OPEN_CNT, AR_OPEN_AMT = R2.AR_OPEN_AMT, AR_OVER_CNT = 0, AR_OVER_AMT = 0
     where Current of C1;
   End Loop; End;
  End;

  Begin Declare Cursor C2 is 
   Select Count(Distinct ARTOPEN1.INV_NUM) AR_OPEN_CNT, Sum (ARTOPEN1.INV_BALANCE) AR_OPEN_AMT
    from (Select DETL_CVX_NO CUST_CODE, DETL_CVX_TYPE INV_TYPE, DETL_CTL_NO INV_NUM, CREC_AMT INV_BALANCE
           from GLTCREC3 where OPS_YYYYPP = '{YP}' and DETL_CTL_TYPE <> 'I' and CREC_TYPE_CODE = 'AR') ARTOPEN1
    where ARTOPEN1.CUST_CODE = R1.CUST_CODE;
   Begin For R2 in C2 Loop
    Update {ARTCUSTX} Set AR_OPEN_CNT_C = R2.AR_OPEN_CNT, AR_OPEN_AMT_C = R2.AR_OPEN_AMT
     where Current of C1;
   End Loop; End;
  End;
"
            'and ARTOPEN1.INV_BALANCE > 0
            'and ARTOPEN1.INV_TYPE = 'I';
        End If

        Return Sql
    End Function

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()

        Print_Report_End()
    End Sub

    Private Sub grdARTCUSTX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdARTCUSTX.DoubleClickRow
        If e.Row Is Nothing Or Not e.Row.IsDataRow Or e.Row.IsFilterRow Then
            Exit Sub
        End If

        CUST_CODE = e.Row.Cells("CUST_CODE").Value & ""
        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Click_Command("Load")

    End Sub

    Sub Get_Data()

        For Each TYPE As String In New String() {"OPEN", "BOOK", "SHIP", "RTRN", "PYMT"}
            dst.Tables($"ARTCUSTX_{TYPE}").Rows.Clear()
        Next
        dst.Tables("ARTPYMTS").Rows.Clear()

        Dim LAST_YP As String = ""

        For YPs As Integer = 0 To 2

            Dim YP As String = ASCMAIN1.CYP
            If YP > 0 Then
                YP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * YPs)
            End If

            If YPs > 0 Then
                ASCMAIN1.sql = $"
Begin Declare Cursor C1 is 
 Select * from {ARTCUSTX} where CUST_CODE = '{CUST_CODE}' for Update;
 Begin For R1 in C1 Loop  
  {sqlSOTORDR1_BOOK_SHIP_RTRN(YP)}
 End Loop; End;
End;"
                ASCDATA1.ExecuteSQL()

                Fill_Records("ARTCUSTX")
            End If

            Dim rowARTCUSTX As DataRow = dst.Tables("ARTCUSTX").Rows.Find(CUST_CODE)

            'Dim rowARTPYMTS As DataRow = Fill_Record("ARTPYMTS", New String() {YP, CUST_CODE})
            ASCMAIN1.sql = Replace(Replace(sqlARTPYMTS, ":PARM1", YP), ":PARM2", CUST_CODE)
            Fill_Records("ARTPYMTS",,, ASCMAIN1.sql)
            Dim rowARTPYMTS As DataRow = dst.Tables("ARTPYMTS").Rows(0)

            For Each TYPE As String In New String() {"OPEN", "BOOK", "SHIP", "RTRN", "PYMT"}

                If TYPE = "OPEN" And YPs > 0 Then
                    ' DO NOTHING FOR OPEN WHEN PULLING IN HISTORY
                Else

                    Dim rowARTCUSTX_TYPE As DataRow = dst.Tables($"ARTCUSTX_{TYPE}").NewRow
                    With rowARTCUSTX_TYPE
                        .Item("TYPE") = TYPE
                        .Item("YP") = YP

                        If TYPE = "PYMT" Then
                            .Item("CNT") = rowARTPYMTS.Item("RECS")
                            .Item("PYMT") = rowARTPYMTS.Item("PYMT")
                            .Item("APPL") = rowARTPYMTS.Item("APPL")
                            .Item("FEE") = rowARTPYMTS.Item("FEE")
                            .Item("OTHER") = rowARTPYMTS.Item("OTHER")
                            .Item("CBOA") = rowARTPYMTS.Item("CBOA")
                            .Item("TB") = rowARTPYMTS.Item("TB")
                        Else
                            .Item("CNT") = rowARTCUSTX.Item($"{TYPE}_CNT")
                            .Item("QTY") = rowARTCUSTX.Item($"{TYPE}_QTY")
                            .Item("AMT") = rowARTCUSTX.Item($"{TYPE}_AMT")
                        End If

                        If TYPE = "OPEN" Then
                            .Item("YP") = "000000"
                            .Item("MIN_DATE") = rowARTCUSTX.Item($"{TYPE}_MIN_DATE")
                            .Item("MAX_DATE") = rowARTCUSTX.Item($"{TYPE}_MAX_DATE")
                        ElseIf TYPE = "BOOK" Then
                            .Item("CNT_OPEN") = rowARTCUSTX.Item($"{TYPE}_CNT_OPEN")
                            .Item("CNT_PICK") = rowARTCUSTX.Item($"{TYPE}_CNT_PICK")
                            .Item("CNT_SHIP") = rowARTCUSTX.Item($"{TYPE}_CNT_SHIP")
                            .Item("CNT_CANC") = rowARTCUSTX.Item($"{TYPE}_CNT_CANC")
                        ElseIf TYPE = "SHIP" Then
                            .Item("GRS") = rowARTCUSTX.Item($"{TYPE}_GRS")
                            .Item("DSC") = rowARTCUSTX.Item($"{TYPE}_DSC")
                            .Item("UNPAID") = rowARTCUSTX.Item($"AR_OPEN_CNT")
                        ElseIf TYPE = "RTRN" Then
                            .Item("GRS") = rowARTCUSTX.Item($"{TYPE}_GRS")
                            .Item("DSC") = rowARTCUSTX.Item($"{TYPE}_DSC")
                            .Item("UNPAID") = rowARTCUSTX.Item($"AR_OPEN_CNT_C")
                        End If
                    End With
                    dst.Tables($"ARTCUSTX_{TYPE}").Rows.Add(rowARTCUSTX_TYPE)
                End If
            Next

            Dim rowSOTINVHX As DataRow = Fill_Record("SOTINVHX", New String() {CUST_CODE, YP}, True)
            Dim rowARTCUSTX_RF As DataRow = dst.Tables($"ARTCUSTX_RF").NewRow
            With rowARTCUSTX_RF

                .Item("YP") = YP

                .Item("BEG_BAL") = 0
                .Item("GRS_SLS") = Val(rowARTCUSTX.Item("SHIP_GRS") & "") + Val(rowARTCUSTX.Item("RTRN_GRS") & "")
                .Item("DISC") = Val(rowARTCUSTX.Item("SHIP_DSC") & "") + Val(rowARTCUSTX.Item("RTRN_DSC") & "")
                .Item("NET_SLS") = rowSOTINVHX.Item("INV_SALES")
                .Item("SLS_TAX") = rowSOTINVHX.Item("INV_STAX")
                .Item("FRT_CHG") = rowSOTINVHX.Item("INV_FREIGHT")
                .Item("MISC_CHG") = rowSOTINVHX.Item("INV_MISC_CHG")
                .Item("NET_AMT") = rowSOTINVHX.Item("INV_TOTAL_AMOUNT")
                .Item("GRS_PMT") = Val(rowARTPYMTS.Item("PYMT") & "")
                .Item("FEE_AMT") = Val(rowARTPYMTS.Item("FEE") & "")
                .Item("OTH_DED") = Val(rowARTPYMTS.Item("OTHER") & "") + Val(rowARTPYMTS.Item("GL") & "")
                .Item("NET_PMT") = Val(rowARTPYMTS.Item("APPL") & "") + Val(rowARTPYMTS.Item("CBOA") & "")
                .Item("CGS") = rowSOTINVHX.Item("INV_COGS")
                .Item("END_BAL") = Val(rowARTCUSTX.Item("AR_OPEN_AMT") & "") + Val(rowARTCUSTX.Item("AR_OPEN_AMT_C") & "")
                dst.Tables($"ARTCUSTX_RF").Rows.Add(rowARTCUSTX_RF)

                If YPs > 0 Then
                    rowARTCUSTX_RF = dst.Tables("ARTCUSTX_RF").Rows.Find(LAST_YP)
                    rowARTCUSTX_RF.Item("BEG_BAL") = rowARTCUSTX.Item("AR_OPEN_AMT")
                End If

            End With

            LAST_YP = YP
        Next
    End Sub

    Private Sub grdARTCUSTX_DASH_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUSTX_DASH.AfterRowActivate
        If grdARTCUSTX_DASH.ActiveRow IsNot Nothing AndAlso grdARTCUSTX_DASH.ActiveRow.IsDataRow AndAlso Not grdARTCUSTX_DASH.ActiveRow.IsFilterRow Then

            If grdARTCUSTX_DASH.ActiveRow.Band.Index = 0 Then
                Exit Sub
            End If

            With grdARTCUSTX_DASH.ActiveRow

                ASCMAIN1.sql = sqlSOTORDR1 & vbCrLf & $" And SOTORDR1.CUST_CODE = '{CUST_CODE}'"

                    Dim TYPE As String = .Cells("TYPE").Value & ""
                Dim YP As String = .Cells("YP").Value & ""
                Dim LEGEND As String = ""
                If YP <> "000000" Then
                    Dim rowGLTPARM2 As DataRow = dst.Tables("GLTPARM2").Rows.Find(YP)
                    LEGEND = rowGLTPARM2.Item("LEGEND")
                End If

                Select Case TYPE

                    Case "OPEN"
                        ASCMAIN1.sql &= vbCrLf & " and SOTORDR1.ORDR_STATUS = 'O'"
                        Fill_Records("SOTORDR1", ,, ASCMAIN1.sql)
                        Sort_grdColumns(grdSOTORDR1, "ORDR_NO")

                        grdSOTORDR1.Text = "Open Orders"
                        splMain.Panel2Collapsed = False

                    Case "BOOK"
                        ASCMAIN1.sql &= vbCrLf & $" and SOTORDR1.ORDR_YYYYPP_BOOKED = '{YP}'"
                        Fill_Records("SOTORDR1", ,, ASCMAIN1.sql)
                        Sort_grdColumns(grdSOTORDR1, "ORDR_NO")

                        grdSOTORDR1.Text = $"Orders Booked in {LEGEND}"
                        splMain.Panel2Collapsed = False

                    Case "SHIP"
                        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "from SOTORDR1", "from SOTORDR1,SOTINVH1")
                        ASCMAIN1.sql &= vbCrLf & $" and SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO and SOTINVH1.ORDR_YYYYPP_UPDATED = '{YP}'"
                        Fill_Records("SOTORDR1", ,, ASCMAIN1.sql)
                        Sort_grdColumns(grdSOTORDR1, "ORDR_NO")

                        grdSOTORDR1.Text = $"Orders Shipped in {LEGEND}"
                        splMain.Panel2Collapsed = False

                    Case Else
                        splMain.Panel2Collapsed = True
                End Select
            End With
        End If

    End Sub

    Private Sub grdSOTORDR1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDR1.AfterRowActivate
        If grdSOTORDR1.ActiveRow IsNot Nothing AndAlso grdSOTORDR1.ActiveRow.IsDataRow AndAlso Not grdSOTORDR1.ActiveRow.IsFilterRow Then
            Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value

            Fill_Records("SOTORDR2", ORDR_NO)
            Sort_grdColumns(grdSOTORDR2, "ORDR_NO,ORDR_LNO")

            'dst.Tables("SOTORDR2").Rows.Clear()
            'For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            '    Fill_Records("SOTORDR2", rowSOTORDR1.Item("ORDR_NO"), False)
            'Next
        End If
    End Sub

    Private Sub grdARTCUSTX_RF_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdARTCUSTX_RF.InitializeRow
        If e.Row IsNot Nothing AndAlso e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            Dim GRS_SLS As Decimal = Val(e.Row.Cells("GRS_SLS").Value & "")
            Dim DISC As Decimal = Val(e.Row.Cells("DISC").Value & "")
            Dim NET_SLS As Decimal = Val(e.Row.Cells("NET_SLS").Value & "")
            Dim SLS_TAX As Decimal = Val(e.Row.Cells("SLS_TAX").Value & "")
            Dim FRT_CHG As Decimal = Val(e.Row.Cells("FRT_CHG").Value & "")
            Dim MISC_CHG As Decimal = Val(e.Row.Cells("MISC_CHG").Value & "")
            Dim NET_AMT As Decimal = Val(e.Row.Cells("NET_AMT").Value & "")

            Dim GRS_PMT As Decimal = Val(e.Row.Cells("GRS_PMT").Value & "")
            Dim FEE_AMT As Decimal = Val(e.Row.Cells("FEE_AMT").Value & "")
            Dim OTH_DED As Decimal = Val(e.Row.Cells("OTH_DED").Value & "")
            Dim NET_PMT As Decimal = Val(e.Row.Cells("NET_PMT").Value & "")

            If GRS_SLS - DISC - NET_SLS <> 0 Then
                e.Row.Cells("NET_SLS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("NET_SLS").ToolTipText = "Net Sls <> Grs Sls - Disc; Out by " & Format(GRS_SLS - DISC - NET_SLS, "#,###.00")
            Else
                e.Row.Cells("NET_SLS").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("NET_SLS").ToolTipText = ""
            End If

            If NET_SLS + SLS_TAX + FRT_CHG + MISC_CHG - NET_AMT <> 0 Then
                e.Row.Cells("NET_AMT").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("NET_AMT").ToolTipText = "Net Amt <> Net Sls + Sls Tax + Frt + Misc; Out by " & Format(NET_SLS + SLS_TAX + FRT_CHG + MISC_CHG - NET_AMT, "#,###.00")
            Else
                e.Row.Cells("NET_AMT").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("NET_AMT").ToolTipText = ""
            End If

            If GRS_PMT + FEE_AMT + OTH_DED - NET_PMT <> 0 Then
                e.Row.Cells("NET_PMT").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("NET_PMT").ToolTipText = "Net Pmt <> Grs Pmt + Fee + Other Ded; Out by " & Format(GRS_PMT + FEE_AMT + OTH_DED - NET_PMT, "#,###.00")
            Else
                e.Row.Cells("NET_PMT").Appearance.ForeColor = Drawing.Color.Empty
                e.Row.Cells("NET_PMT").ToolTipText = ""
            End If
        End If

    End Sub
End Class