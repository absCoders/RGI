Public Class SAFSLSA1
    Dim SATCSLS1 As String
    Dim SATCSLSX As String

    Dim sqlSATCSLS1 As String
    Dim sqlSATCSLSX As String
    Dim sqlSOTINVHX As String

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW0 As String
    Dim RYW1 As String
    Dim Periods As Integer

    Dim CUST_CODE As String
    Dim SREP_CODE As String
    Dim CHECK_BOX As String


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        With dst

            sqlSATCSLSX = "Select STYLE_CODE, COLOR_CODE, CUST_CODE" & vbCrLf _
                & ", Sum (SLS_AMT) SLS_AMT, Sum (SLS_AMT_WHSE1) SLS_AMT_WHSE1, Sum (SLS_AMT_WHSE2) SLS_AMT_WHSE2, Sum (SLS_AMT_WHSEX) SLS_AMT_WHSEX" & vbCrLf _
                & ", Sum (SLS_QTY) SLS_QTY, Sum (SLS_QTY_WHSE1) SLS_QTY_WHSE1, Sum (SLS_QTY_WHSE2) SLS_QTY_WHSE2, Sum (SLS_QTY_WHSEX) SLS_QTY_WHSEX" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE" & vbCrLf _
                & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0))) SLS_AMT_WHSEX" & vbCrLf _
                & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',0,'WHSE2',0,NVL(SOTINVH2.ORDR_QTY_SHIP,0))) SLS_QTY_WHSEX" & vbCrLf _
                & " from SOTINVH2,SOTINVH1" & vbCrLf _
                & " where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "   and SOTINVH2.ORDR_QTY_SHIP <> 0" & vbCrLf _
                & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                & "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
                & " group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH1.CUST_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
                & ", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) SLS_AMT" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',0,'WHSE2',0,(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0))) SLS_AMT_WHSEX" & vbCrLf _
                & ", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0))) SLS_QTY" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE2" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',0,'WHSE2',0,(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)))) SLS_QTY_WHSEX" & vbCrLf _
                & " from SOTORDR2,SOTORDR1" & vbCrLf _
                & " where SOTORDR2.ORDR_STATUS BETWEEN 'Z' AND 'Z'" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) <> 0" & vbCrLf _
                & "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_YYYYPP_BOOKED BETWEEN 'YP1' AND 'YP2'" & vbCrLf _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE, CUST_CODE"

            ASCMAIN1.sql = sqlSATCSLSX
            SATCSLSX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATCSLSX & " Add Primary Key (STYLE_CODE, COLOR_CODE, CUST_CODE)")

            sqlSATCSLS1 = "Select X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYC1.STYLE_COLOR_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, APTVEND1.PORT_CODE, ICTSTYL1.CARTON_PACK_QTY" & vbCrLf _
                & ", ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST" & vbCrLf _
                & ", X.SLS_AMT, X.SLS_AMT_WHSE1, X.SLS_AMT_WHSE2" & vbCrLf _
                & ", X.SLS_QTY, X.SLS_QTY_WHSE1, X.SLS_QTY_WHSE2" & vbCrLf _
                & ", Z.SLS_CUSTS, Z.SLS_CUSTS_WHSE1, Z.SLS_CUSTS_WHSE2, Z.SLS_CUSTS_WHSEX" & vbCrLf _
                & ", Y.FUT_QTY, Y.FUT_QTY_WHSE1, Y.FUT_QTY_WHSE2 , ICTSTYC1.THEME_CODE" & vbCrLf _
                & " from ICTSTYL1, ICTCOLR1, ICTSTYV1, ICTSTYC1, APTVEND1" & vbCrLf _
                & ",(Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", Sum (SLS_AMT) SLS_AMT, Sum (SLS_AMT_WHSE1) SLS_AMT_WHSE1, Sum (SLS_AMT_WHSE2) SLS_AMT_WHSE2" & vbCrLf _
                & ", Sum (SLS_QTY) SLS_QTY, Sum (SLS_QTY_WHSE1) SLS_QTY_WHSE1, Sum (SLS_QTY_WHSE2) SLS_QTY_WHSE2" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS_AMT" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2" & vbCrLf _
                & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) SLS_QTY" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE1',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTINVH1.WHSE_CODE,'WHSE2',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) SLS_QTY_WHSE2" & vbCrLf _
                & " from SOTINVH2,SOTINVH1" & vbCrLf _
                & " where SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN 'YP1' AND 'YP2'" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "   and SOTINVH2.ORDR_QTY_SHIP <> 0" & vbCrLf _
                & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                & "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
                & " group by SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) SLS_AMT" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0),0)) SLS_AMT_WHSE2" & vbCrLf _
                & ", SUM ((NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0))) SLS_QTY" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE1',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE1" & vbCrLf _
                & ", SUM (DECODE(SOTORDR1.WHSE_CODE,'WHSE2',(NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)),0)) SLS_QTY_WHSE2" & vbCrLf _
                & " from SOTORDR2,SOTORDR1" & vbCrLf _
                & " where SOTORDR2.ORDR_STATUS BETWEEN 'Z' AND 'Z'" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) <> 0" & vbCrLf _
                & "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_YYYYPP_BOOKED BETWEEN 'YP1' AND 'YP2'" & vbCrLf _
                & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ") X" & vbCrLf _
                & ", (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_QTY" & vbCrLf _
                & ", SUM (DECODE(WHSE_CODE,'WHSE1',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) FUT_QTY_WHSE1" & vbCrLf _
                & ", SUM (DECODE(WHSE_CODE,'WHSE2',NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0),0)) FUT_QTY_WHSE2" & vbCrLf _
                & " from ICTSTAT2 GROUP BY STYLE_CODE, COLOR_CODE) Y" & vbCrLf _
                & ", (Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ", SUM(CASE WHEN SLS_QTY > 0 THEN 1 ELSE 0 END) SLS_CUSTS" & vbCrLf _
                & ", SUM(CASE WHEN SLS_QTY_WHSE1 > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSE1" & vbCrLf _
                & ", SUM(CASE WHEN SLS_QTY_WHSE2 > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSE2" & vbCrLf _
                & ", SUM(CASE WHEN SLS_QTY_WHSEX > 0 THEN 1 ELSE 0 END) SLS_CUSTS_WHSEX" & vbCrLf _
                & " from " & SATCSLSX & " group by STYLE_CODE, COLOR_CODE) Z" & vbCrLf _
                & " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   and Y.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                & "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   and Z.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" & vbCrLf _
                & "   and APTVEND1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" & vbCrLf

            ASCMAIN1.sql = sqlSATCSLS1 ' Replace(Replace(Replace(Replace(sqlSATCSLS1, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''")
            SATCSLS1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATCSLS1 & " Add Primary Key (STYLE_CODE, COLOR_CODE)")

            ASCMAIN1.sql = "Select * from " & SATCSLS1
            Create_TDA(.Tables.Add, "SATCSLS1", "**", 0, False)
            ' Create_TDA(.Tables.Add, "SATCSLS1", "**", 0, False, "VVVV")
            With .Tables("SATCSLS1")
                .Columns.Add("SLS_AMT_WHSEX", GetType(System.Decimal), "ISNULL(SLS_AMT,0) - ISNULL(SLS_AMT_WHSE1,0) - ISNULL(SLS_AMT_WHSE2,0)")
                .Columns.Add("SLS_QTY_WHSEX", GetType(System.Int32), "ISNULL(SLS_QTY,0) - ISNULL(SLS_QTY_WHSE1,0) - ISNULL(SLS_QTY_WHSE2,0)")
                .Columns.Add("FUT_QTY_WHSEX", GetType(System.Int32), "ISNULL(FUT_QTY,0) - ISNULL(FUT_QTY_WHSE1,0) - ISNULL(FUT_QTY_WHSE2,0)")
            End With

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO" _
                & ", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.ORDR_NO" _
                & ", SOTINVH1.SREP_CODE, SOTINVH1.WHSE_CODE, SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, SOTINVH1.CUST_STORE_NO" _
                & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION" _
                & ", SOTINVH2.STYLE_CODE, ICTSTYL1.STYLE_DESC" _
                & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
                & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" _
                & " from SOTINVH2,ICTSTYL1,ARTCUST2,SOTINVH1,ARTCUST1 " _
                & " where ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE " _
                & " and ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE " _
                & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' " _
                & " and ARTCUST2.CUST_ADDR_CODE (+) = SOTINVH1.CUST_STORE_NO " _
                & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE " _
                & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & " and SOTINVH2.ORDR_QTY_SHIP <> 0" & vbCrLf _
                & " and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
                & " and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE"
            sqlSOTINVHX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "", 0)
        End With

        grdSATCSLS1.DataSource = dst.Tables("SATCSLS1")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_AMT_SHIP")

        Create_Summary(grdSATCSLS1, "STYLE_CODE", "Count")
        Create_Summary(grdSATCSLS1, New String() {"SLS_AMT", "SLS_AMT_WHSE1", "SLS_AMT_WHSE2", "SLS_AMT_WHSEX", "SLS_QTY", "SLS_QTY_WHSE1", "SLS_QTY_WHSE2", "SLS_QTY_WHSEX", "FUT_QTY", "FUT_QTY_WHSE1", "FUT_QTY_WHSE2", "FUT_QTY_WHSEX"}, , , "#,##0")

        With grdSATCSLS1.DisplayLayout.Bands("SATCSLS1")
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("COLOR_DESC").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"SLS_AMT", "SLS_AMT_WHSE1", "SLS_AMT_WHSE2", "SLS_AMT_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"SLS_QTY", "SLS_QTY_WHSE1", "SLS_QTY_WHSE2", "SLS_QTY_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"FUT_QTY", "FUT_QTY_WHSE1", "FUT_QTY_WHSE2", "FUT_QTY_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                End If
                If New String() {"SLS_CUSTS", "SLS_CUSTS_WHSE1", "SLS_CUSTS_WHSE2", "SLS_CUSTS_WHSEX"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Olive
                    gcol.Width = 60
                    gcol.Format = "#,##0"
                End If
                If New String() {"PORT_CODE"}.Contains(gcol.Key) Then
                    If (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI") Then
                        gcol.Width = 50
                        gcol.Header.VisiblePosition = 11
                        gcol.Hidden = False
                    Else
                        gcol.Hidden = True
                    End If
                End If
                If New String() {"THEME_CODE"}.Contains(gcol.Key) Then
                    If (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI") Then
                        gcol.Hidden = False
                    Else
                        gcol.Hidden = True
                    End If
                End If
            Next
        End With

        ASCMAIN1.Add_Value_List(grdSATCSLS1, "STYLE_STATUS")
        ASCMAIN1.Add_Value_List(grdSATCSLS1, "STYLE_COLOR_STATUS")

        chkShowOriginalStyle.Visible = (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") And (1 <> 1) ' NOT CODED YET



    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then Validate_Code("CUST_CODE")
                If Absx1.txtFor("SREP_CODE").Text <> "" Then Validate_Code("SREP_CODE")


                If EMsg = "" Then
                    If Absx1.cmbFor("RYP0").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify a Starting Period"
                    End If
                    If Absx1.cmbFor("RYP1").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify an Ending Period"
                    End If

                    If EMsg = "" Then
                        RYP0 = Absx1.cmbFor("RYP0").Value
                        RYP1 = Absx1.cmbFor("RYP1").Value
                        Periods = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1
                    End If

                    If Periods > 24 Or Periods < 1 Then
                        EMsg &= vbCr & "Periods must be in chronological order and not more than 24 months apart"
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
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print Report"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Data Options").Visible = tf
                .Groups("Options").Visible = Not tf
                .Groups("Period Range").Visible = Not tf
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = tf

        spl.Panel1Collapsed = ScreenMode And Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("SREP_CODE").Text = ""

        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("CUST_NAME").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("SREP_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
        End With

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SATCSLS1", "SOTINVHX"} ', "SATCSLS1_DTL", "SATCSLS2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        chkShow0Sales.Checked = False
        chkOpenOrders.Checked = True

        ' Absx1.txtFor("CUST_CODE").Text = ""
        ' Absx1.txtFor("SREP_CODE").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Customer Sales Data")

        Save_Header_Fields(UltraGroupBox1)

        CUST_CODE = HFs("CUST_CODE")
        Create_SATCSLS1()

        Set_DataSource()

        Sort_grdColumns(grdSATCSLS1, "STYLE_CODE,COLOR_CODE")
        Setup_grdSOTINVHX()

        With grdSATCSLS1.DisplayLayout.Bands(0)
            For Each COL As String In New String() {"SLS_AMT", "SLS_QTY", "FUT_QTY", "SLS_CUSTS"}
                Dim PFX As String = IIf(COL = "SLS_AMT", "$", IIf(COL = "SLS_QTY", "#", IIf(COL = "FUT_QTY", "Fut", "#C")))
                .Columns(COL).Header.Caption = PFX & " Total"
                .Columns(COL & "_WHSE1").Header.Caption = PFX & " " & Absx1.txtFor("WHSE_CODE1").Text
                .Columns(COL & "_WHSE2").Header.Caption = PFX & " " & Absx1.txtFor("WHSE_CODE2").Text
                .Columns(COL & "_WHSEX").Header.Caption = PFX & " " & "Other"
                .Columns(COL & "_WHSE1").Hidden = (Absx1.txtFor("WHSE_CODE1").Text = "")
                .Columns(COL & "_WHSE2").Hidden = (Absx1.txtFor("WHSE_CODE2").Text = "")
                .Columns(COL & "_WHSEX").Hidden = (Absx1.txtFor("WHSE_CODE1").Text = "") And (Absx1.txtFor("WHSE_CODE2").Text = "")
            Next
        End With

        grdSATCSLS1.Text = "Sales & Inventory by Style Color, Showing Sales from " & Trim(ASCMAIN1.Get_Legend(RYP0)) & " thru " & Trim(ASCMAIN1.Get_Legend(RYP1)) _
            & IIf(Absx1.txtFor("CUST_CODE").Text <> "", ", Customer " & Absx1.txtFor("CUST_CODE").Text & ":" & Absx1.txtFor("CUST_NAME").Text, "") _
            & IIf(Absx1.txtFor("SREP_CODE").Text <> "", ", Sales Rep " & Absx1.txtFor("SREP_CODE").Text & ":" & Absx1.txtFor("SREP_NAME").Text, "")
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATCSLS1, "SSBB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Style Master File")
        Load_Popup_Menu(grdSOTINVHX, "SSBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Show Invoice")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If


            Case "Style Master File"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    'Context_Launch("View", STYLE_CODE, e.Tool.Key, "ICTSTYL1")
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    ' If ASCMAIN1.Running_in_VS Then Stop ' NOT WORKING
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                End If

            Case "Show Invoice"
                Dim FILENAME As String = ""
                If grd.ActiveRow IsNot Nothing Then
                    If Not grd.ActiveRow.Selected Then
                        grd.Selected.Rows.Clear()
                        grd.ActiveRow.Selected = True
                    End If
                End If

                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.Selected Then
                    Exit Sub
                End If

                Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value & ""

                FILENAME = TAC.SOCMAIN1.Create_Invoice(Me, INV_NO)

                Show_Document(FILENAME)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

#End Region

    Sub Create_SATCSLS1()

        ASCMAIN1.sql = "Truncate Table " & SATCSLSX
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & SATCSLSX & " " _
            & Replace(Replace(Replace(Replace(Replace(Replace(Replace(sqlSATCSLSX, _
                                    "'WHSE1'", "'" & Absx1.txtFor("WHSE_CODE1").Text & "'"), _
                                    "'WHSE2'", "'" & Absx1.txtFor("WHSE_CODE2").Text & "'"), _
                                    "BETWEEN 'Z' AND 'Z'", IIf(chkOpenOrders.Checked, "BETWEEN 'O' AND 'P'", "BETWEEN 'Z' AND 'Z'")), _
                                    "'YP1'", "'" & RYP0 & "'"), "'YP2'", "'" & RYP1 & "'"), _
                             "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTINVH1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf), _
                             "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf)
        ASCDATA1.ExecuteSQL()
        Dim danasql As String = ASCMAIN1.sql
         

        ASCMAIN1.sql = "Truncate Table " & SATCSLS1
        ASCDATA1.ExecuteSQL()

        Dim DANA As String = chkOpenOrders.CheckedValue

        ASCMAIN1.sql = "Insert into " & SATCSLS1 & " " _
            & Replace(Replace(Replace(Replace(Replace(Replace(Replace(sqlSATCSLS1, _
                                    "'WHSE1'", "'" & Absx1.txtFor("WHSE_CODE1").Text & "'"), _
                                    "'WHSE2'", "'" & Absx1.txtFor("WHSE_CODE2").Text & "'"), _
                                    "BETWEEN 'Z' AND 'Z'", IIf(chkOpenOrders.Checked, "BETWEEN 'O' AND 'P'", "BETWEEN 'Z' AND 'Z'")), _
                                    "'YP1'", "'" & RYP0 & "'"), "'YP2'", "'" & RYP1 & "'"), _
                             "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTINVH1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf), _
                             "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf)

        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {Absx1.txtFor("WHSE_CODE1").Text, Absx1.txtFor("WHSE_CODE2").Text, RYP0, RYP1})
        ASCDATA1.ExecuteSQL()

        If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("SREP_CODE").Text = "" Then

            ASCMAIN1.sql = "Insert into " & SATCSLS1 & " (STYLE_CODE,COLOR_CODE,STYLE_DESC,STYLE_STATUS,VEND_CODE,FACTORY_CODE,STYLE_UOM,STYLE_CLASS_CODE,CARTON_PACK_QTY,COLOR_DESC,PO_COST) " _
                & " Select X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.CARTON_PACK_QTY" & vbCrLf _
                & ", ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST" & vbCrLf _
                & " from ICTSTYL1,ICTCOLR1,ICTSTYV1" & vbCrLf _
                & ", (Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE from ICTSTYC1,ICTSTYL1" & vbCrLf _
                & "     where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE and ICTSTYL1.STYLE_STATUS = 'A'" & vbCrLf _
                & "   minus Select STYLE_CODE, COLOR_CODE from " & SATCSLS1 & ") X" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE"
            ASCDATA1.ExecuteSQL()

            chkShow0Sales.Visible = True
        Else
            chkShow0Sales.Visible = False

        End If

        Fill_Records("SATCSLS1")

    End Sub

    Sub Print_Report()
        Call Print_Report_Begin()

        Dim SUBT As String = ""
        Dim RecordSelectionFormula As String = ""
        Generate_Report("SARCSLS1", "", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

    Private Sub chkShowDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs)
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub

    End Sub

    Private Sub grdSATCSLS1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSATCSLS1.AfterRowActivate
        Setup_grdSOTINVHX()
    End Sub

    Sub Setup_grdSOTINVHX()
        If grdSATCSLS1.ActiveRow Is Nothing OrElse Not grdSATCSLS1.ActiveRow.IsDataRow Then
            grdSOTINVHX.Visible = False
        Else

            ASCMAIN1.Progress("Now Retrieving Sales Documents")
            Me.Cursor = Cursors.WaitCursor

            grdSOTINVHX.Visible = True

            Dim STYLE_CODE As String = grdSATCSLS1.ActiveRow.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = grdSATCSLS1.ActiveRow.Cells("COLOR_CODE").Value & ""
            ASCMAIN1.sql = sqlSOTINVHX & vbCrLf _
                & " and SOTINVH2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                & " and SOTINVH2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                & " and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
                & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTINVH1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf

            Fill_Records("SOTINVHX", "", , ASCMAIN1.sql)

            'If RYP1 = ASCMAIN1.CYP And chkOpenOrders.Checked Then
            '    ASCMAIN1.sql = "Select SOTORDR2.ORDR_STATUS INV_TYPE, SOTORDR2.ORDR_NO INV_NO" & vbCrLf _
            '        & ", SOTORDR1.ORDR_SHIP_DATE INV_DATE, SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
            '        & ", '000000' OPS_YYYYPP, SOTORDR1.ORDR_NO" & vbCrLf _
            '        & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, ARTCUST1.CUST_NAME, SOTORDR1.CUST_STORE_NO" & vbCrLf _
            '        & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION" & vbCrLf _
            '        & ", SOTORDR2.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            '        & ", NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
            '        & ", (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
            '        & " from SOTORDR2,ICTSTYL1,ARTCUST2,SOTORDR1,ARTCUST1 " & vbCrLf _
            '        & " where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE " & vbCrLf _
            '        & " and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE " & vbCrLf _
            '        & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' " & vbCrLf _
            '        & " and ARTCUST2.CUST_ADDR_CODE (+) = SOTORDR1.CUST_STORE_NO " & vbCrLf _
            '        & " and SOTORDR2.ORDR_STATUS BETWEEN 'O' AND 'P'" & vbCrLf _
            '        & " and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            '        & " and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            '        & " and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            '        & " and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            '        & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
            '        & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf

            '    Fill_Records("SOTINVHX", "", False, ASCMAIN1.sql)
            'End If


            'If RYP1 <> ASCMAIN1.CYP And chkOpenOrders.Checked Then
            '    ASCMAIN1.sql = "Select SOTORDR2.ORDR_STATUS INV_TYPE, SOTORDR2.ORDR_NO INV_NO" & vbCrLf _
            '        & ", SOTORDR1.ORDR_SHIP_DATE INV_DATE, SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
            '        & ", '000000' OPS_YYYYPP, SOTORDR1.ORDR_NO" & vbCrLf _
            '        & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, ARTCUST1.CUST_NAME, SOTORDR1.CUST_STORE_NO" & vbCrLf _
            '        & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION" & vbCrLf _
            '        & ", SOTORDR2.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            '        & ", NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
            '        & ", (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
            '        & " from SOTORDR2,ICTSTYL1,ARTCUST2,SOTORDR1,ARTCUST1 " & vbCrLf _
            '        & " where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE " & vbCrLf _
            '        & " and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE " & vbCrLf _
            '        & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' " & vbCrLf _
            '        & " and ARTCUST2.CUST_ADDR_CODE (+) = SOTORDR1.CUST_STORE_NO " & vbCrLf _
            '        & " and SOTORDR2.ORDR_STATUS BETWEEN 'O' AND 'P'" & vbCrLf _
            '        & " and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            '        & " and SOTORDR1.ORDR_YYYYPP_BOOKED < = '" & RYP1 & "'" & vbCrLf _
            '        & " and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            '        & " and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
            '        & " and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
            '        & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
            '        & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf

            '    Fill_Records("SOTINVHX", "", False, ASCMAIN1.sql)
            'End If

            If chkOpenOrders.Checked Then
                ASCMAIN1.sql = "Select SOTORDR2.ORDR_STATUS INV_TYPE, SOTORDR2.ORDR_NO INV_NO" & vbCrLf _
                    & ", SOTORDR1.ORDR_SHIP_DATE INV_DATE, SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                    & ", '000000' OPS_YYYYPP, SOTORDR1.ORDR_NO" & vbCrLf _
                    & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE, ARTCUST1.CUST_NAME, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                    & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION" & vbCrLf _
                    & ", SOTORDR2.STYLE_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
                    & ", NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                    & ", (NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0)) * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
                    & " from SOTORDR2,ICTSTYL1,ARTCUST2,SOTORDR1,ARTCUST1 " & vbCrLf _
                    & " where ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE " & vbCrLf _
                    & " and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE " & vbCrLf _
                    & " and ARTCUST2.CUST_ADDR_TYPE (+) = 'MK' " & vbCrLf _
                    & " and ARTCUST2.CUST_ADDR_CODE (+) = SOTORDR1.CUST_STORE_NO " & vbCrLf _
                    & " and SOTORDR2.ORDR_STATUS BETWEEN 'O' AND 'P'" & vbCrLf _
                    & " and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                    & " and SOTORDR1.ORDR_YYYYPP_BOOKED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
                    & " and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                    & " and SOTORDR2.STYLE_CODE = '" & STYLE_CODE & "'" & vbCrLf _
                    & " and SOTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                    & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                    & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf

                Fill_Records("SOTINVHX", "", False, ASCMAIN1.sql)
            End If


            grdSOTINVHX.Text = "Sales Documents for " & STYLE_CODE & "-" & COLOR_CODE

            ASCMAIN1.Progress("")
            Me.Cursor = Cursors.Default
        End If

    End Sub

    Private Sub chkShow0Sales_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShow0Sales.CheckedChanged
        Set_Datasource()
    End Sub

    Sub Set_DataSource()
        Dim dvw As DataView = DirectCast(grdSATCSLS1.DataSource, DataTable).DefaultView
        If chkShow0Sales.Checked Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "SLS_QTY <> 0"
        End If
    End Sub

    Private Sub grdSATCSLS1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATCSLS1.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("STYLE_STATUS").Value & "" <> e.Row.Cells("STYLE_COLOR_STATUS").Value & "" Then
                e.Row.Cells("STYLE_COLOR_STATUS").Appearance.ForeColor = Drawing.Color.Red
                e.Row.Cells("STYLE_COLOR_STATUS").ToolTipText = "Color Status is not in agreement with Style Status"
            End If
        End If
    End Sub
End Class