Imports System.Text
Imports Infragistics.Win.UltraWinGrid

Public Class SAFSLSA2
    'This was copied from SAFSLSA1 on 7/24/16 to make specific changes for Regency.
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
    Dim TTM As New UltraWinToolTip.UltraToolTipManager
    Dim isDataLoading As Boolean = True

#Region "ABS Standard Routines"
    'These Routines should be found in all Forms which Launch from the Menu.

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
                & ", Y.FUT_QTY, Y.FUT_QTY_WHSE1, Y.FUT_QTY_WHSE2 , ICTSTYC1.THEME_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
                & " from ICTSTYL1, ICTCOLR1, ICTSTYV1, ICTSTYC1, APTVEND1, ARTCUST1" & vbCrLf _
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
                & "   and APTVEND1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" & vbCrLf _
                & "   and ICTSTYL1.CUST_CODE = ARTCUST1.CUST_CODE (+)" & vbCrLf

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
                .Columns.Add("ATTR_CODE1")
                .Columns.Add("ATTR_CODE2")
                .Columns.Add("ATTR_CODE3")
                .Columns.Add("ATTR_CODE4")
                .Columns.Add("IS_ECOM")
            End With

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO" _
                & ", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.ORDR_NO" _
                & ", SOTINVH1.SREP_CODE, SOTINVH1.WHSE_CODE, SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, SOTINVH1.CUST_STORE_NO" _
                & ", NVL(ARTCUST2.CUST_ADDR_NAME,ARTCUST2.CUST_NAME) CUST_STORE_LOCATION" _
                & ", SOTINVH2.STYLE_CODE, ICTSTYL1.STYLE_DESC" _
                & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
                & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" _
                & ", SYSDATE AS ORDR_DATE_RECD" _
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
            sqlSOTINVHX = sqlSOTINVHX.Replace("SYSDATE", "NULL")

            ASCMAIN1.sql = "SELECT *" _
                & " FROM ASTGRID1" _
                & " WHERE USER_ID = :PARM1" _
                & " AND FORM_NAME = :PARM2" _
                & " AND GRID_NAME = :PARM3"
            Create_TDA(.Tables.Add, "ASTGRID1", "**", 0, True, "VVV")

            'ASCMAIN1.sql = "SELECT * FROM ICTSTYL3 where STYLE_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False, "V", 0)
            Dim s As New Text.StringBuilder With {.Length = 0}
            s.AppendLine("SELECT S3.STYLE_CODE,")
            s.AppendLine("S3.ATTR_CODE,")
            s.AppendLine("NVL(A1.ATT_RANK,9) ATT_RANK")
            s.AppendLine("FROM ICTSTYL3 S3, ICTATTR1 A1")
            s.AppendLine("WHERE S3.ATTR_CODE = A1.ATTR_CODE")
            ASCMAIN1.sql = s.ToString
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False)
        End With

        Fill_Records("ICTSTYL3")

        grdSATCSLS1.DataSource = dst.Tables("SATCSLS1")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_AMT_SHIP")

        Create_Summary(grdSATCSLS1, "STYLE_CODE", "Count")
        Create_Summary(grdSATCSLS1, New String() {"SLS_AMT", "SLS_AMT_WHSE1", "SLS_AMT_WHSE2", "SLS_AMT_WHSEX", "SLS_QTY", "SLS_QTY_WHSE1", "SLS_QTY_WHSE2", "SLS_QTY_WHSEX", "FUT_QTY", "FUT_QTY_WHSE1", "FUT_QTY_WHSE2", "FUT_QTY_WHSEX"}, , , "#,##0")

        grdSOTINVHX.DisplayLayout.Bands(0).Columns("ORDR_DATE_RECD").Format = "MM/dd/yyyy"

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
                    gcol.Width = 50
                    gcol.Hidden = False
                End If
                If New String() {"THEME_CODE"}.Contains(gcol.Key) Then
                    gcol.Hidden = False
                End If
            Next
        End With

        ASCMAIN1.Add_Value_List(grdSATCSLS1, "STYLE_STATUS")
        ASCMAIN1.Add_Value_List(grdSATCSLS1, "STYLE_COLOR_STATUS")

        chkShowOriginalStyle.Visible = (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") And (1 <> 1) ' NOT CODED YET


        isDataLoading = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then Validate_Code("CUST_CODE")
                If Absx1.txtFor("SREP_CODE").Text <> "" Then Validate_Code("SREP_CODE")
                If chkStockStyles.Checked = False And chkNonStockStyles.Checked = False Then
                    EMsg &= vbCr & "You Must Select Stock, Non-Stock or Both"
                End If

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
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            UltraExplorerBar1.Groups("Style Image").Visible = True
        Else
            UltraExplorerBar1.Groups("Style Image").Visible = False
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = tf

        spl.Panel1Collapsed = ScreenMode And Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("SREP_CODE").Text = ""

        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("CUST_NAME").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("SREP_CODE").Hidden = (Absx1.txtFor("CUST_CODE").Text <> "")
        End With

        ASCMAIN1.TACMAIN1.loadGridLayout(Me, grdSATCSLS1)

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

    Sub Print_Report()
        Call Print_Report_Begin()

        Dim SUBT As String = ""
        Dim RecordSelectionFormula As String = ""
        Generate_Report("SARCSLS1", "", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATCSLS1, "SSBBBBBBBS", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Style Master File", "Save Grid Layout", "Discontinue All Colors", "Discontinue This Color", "DNR All Colors", "DNR This Color", "Show All Attributes")
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
                    If grdSATCSLS1.Selected.Rows.Count = 1 Then
                        Dim STYLE_STATUS As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value
                        Dim STYLE_COLOR_STATUS As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Value
                        Select Case STYLE_STATUS
                            Case "D"
                                e.Tool.ToolbarsManager.Tools("Discontinue All Colors").SharedProps.Visible = False
                                e.Tool.ToolbarsManager.Tools("Discontinue This Color").SharedProps.Visible = False
                                e.Tool.ToolbarsManager.Tools("DNR This Color").SharedProps.Visible = False
                                e.Tool.ToolbarsManager.Tools("DNR All Colors").SharedProps.Visible = False
                            Case "N"
                                e.Tool.ToolbarsManager.Tools("Discontinue All Colors").SharedProps.Visible = True
                                e.Tool.ToolbarsManager.Tools("Discontinue This Color").SharedProps.Visible = True
                                e.Tool.ToolbarsManager.Tools("DNR This Color").SharedProps.Visible = False
                                e.Tool.ToolbarsManager.Tools("DNR All Colors").SharedProps.Visible = False
                            Case "A"
                                e.Tool.ToolbarsManager.Tools("Discontinue All Colors").SharedProps.Visible = True
                                e.Tool.ToolbarsManager.Tools("Discontinue This Color").SharedProps.Visible = True
                                e.Tool.ToolbarsManager.Tools("DNR This Color").SharedProps.Visible = True
                                e.Tool.ToolbarsManager.Tools("DNR All Colors").SharedProps.Visible = True
                        End Select
                    Else
                        e.Tool.ToolbarsManager.Tools("Discontinue All Colors").SharedProps.Visible = False
                        e.Tool.ToolbarsManager.Tools("DNR All Colors").SharedProps.Visible = False
                        e.Tool.ToolbarsManager.Tools("Discontinue This Color").SharedProps.Visible = False
                        e.Tool.ToolbarsManager.Tools("DNR This Color").SharedProps.Visible = False
                    End If

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

            Case "Save Grid Layout"
                ASCMAIN1.TACMAIN1.SaveGridLayout(Me, grd)
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

            Case "Discontinue All Colors", "DNR All Colors"
                DiscontinueAll(e.Tool.Key)
            Case "Discontinue This Color", "DNR This Color"
                DiscontinueColor(e.Tool.Key)
            Case "Show All Attributes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
                tlb_sbt = DirectCast(tlb.Tools("Show All Attributes"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Me.Cursor = Cursors.WaitCursor
                    For Each rowSATCSLS1 As DataRow In dst.Tables("SATCSLS1").Select("", "STYLE_CODE, COLOR_CODE")
                        Dim STYLE_CODE As String = rowSATCSLS1.Item("STYLE_CODE").ToString
                        Dim rowICTSTYL31 As DataRow = dst.Tables("ICTSTYL3").Select(String.Format("STYLE_CODE = '{0}' AND ATT_RANK = '1'", STYLE_CODE)).FirstOrDefault
                        If Not IsNothing(rowICTSTYL31) Then
                            rowSATCSLS1.Item("ATTR_CODE1") = rowICTSTYL31.Item("ATTR_CODE")
                        End If
                        Dim nextI As Integer = 2
                        For Each rowICTSTYL3 As DataRow In dst.Tables("ICTSTYL3").Select(String.Format("STYLE_CODE = '{0}' AND ATT_RANK <> '1'", STYLE_CODE), "ATTR_CODE")
                            If nextI > 4 Then Exit For
                            rowSATCSLS1.Item(String.Format("ATTR_CODE{0}", nextI)) = rowICTSTYL3.Item("ATTR_CODE")
                            nextI += 1
                        Next
                    Next
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE1").Hidden = False
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = False
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = False
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = False
                    grdSATCSLS1.UpdateData()
                Else
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE1").Hidden = True
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = True
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = True
                    grdSATCSLS1.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = True
                End If
                Me.Cursor = Cursors.Default
        End Select
    End Sub

    Private Sub DiscontinueAll(ByVal DiscType As String)
        If grdSATCSLS1.Selected.Rows.Count <> 1 Then
            MsgBox("You Must Select One Row", vbOKOnly, "Row Selection")
            Exit Sub
        End If
        Dim ORIG_STATUS As String = ""
        Dim NEW_STATUS As String = "D"
        Dim NEW_STATUS_DESC As String = "Discontinue"
        If DiscType = "DNR All Colors" Then
            NEW_STATUS = "N"
            NEW_STATUS_DESC = "DNR"
        End If
        If NEW_STATUS = "N" And grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Text = "D" Then
            MsgBox("You May Not DNR A Style That Is Already Discontinued", vbOKOnly, "No update Allowed")
            Exit Sub
        End If
        Dim STYLE_CODE As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
        Dim iResult As MsgBoxResult
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine(String.Format("This Will {0} The Following Item", NEW_STATUS_DESC))
        iMSG.AppendLine(String.Format("And All Of It's Non-{0}ed Colors:", NEW_STATUS_DESC))
        iMSG.AppendLine(STYLE_CODE)
        iMSG.AppendLine("")
        iMSG.AppendLine("Is This Really What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, NEW_STATUS_DESC & " Item")
        If iResult = MsgBoxResult.Yes Then
            If ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE) Then
                BeginTrans()
                Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                SQLS.AppendLine(String.Format("UPDATE ICTSTYL1 SET STYLE_STATUS = '{0}' WHERE STYLE_CODE = '{1}'", NEW_STATUS, STYLE_CODE))
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
                ORIG_STATUS = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value
                grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value = NEW_STATUS

                Dim sql As New Text.StringBuilder With {.Length = 0}
                sql.AppendLine("SELECT *")
                sql.AppendLine("FROM ICTSTYC1")
                sql.AppendLine("WHERE STYLE_CODE = '" & STYLE_CODE & "'")
                Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
                For Each rowICTSTYC1 As DataRow In tbl.Rows
                    Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE").ToString
                    SQLS.Length = 0
                    SQLS.AppendLine(String.Format("UPDATE ICTSTYC1 SET STYLE_COLOR_STATUS = '{0}' WHERE STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", NEW_STATUS, STYLE_CODE, COLOR_CODE))
                    ASCMAIN1.sql = SQLS.ToString
                    ASCDATA1.ExecuteSQL()
                Next

                For Each rowSATCSLS1 As DataRow In dst.Tables("SATCSLS1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                    If rowSATCSLS1.Item("STYLE_COLOR_STATUS") & "" <> NEW_STATUS Then
                        If NEW_STATUS = "N" And rowSATCSLS1.Item("STYLE_COLOR_STATUS") & "" = "D" Then
                            'Don't DNR colors that are Discontinued
                        Else
                            rowSATCSLS1.Item("STYLE_COLOR_STATUS") = NEW_STATUS
                            rowSATCSLS1.Item("STYLE_STATUS") = NEW_STATUS
                        End If
                    End If
                Next

                Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                With rowASTAUDT1
                    .Item("TABLE_NAME") = "ICTSTYL1"
                    .Item("KEY_VALUE") = STYLE_CODE
                    .Item("COLUMN_NAME") = "STYLE_STATUS"
                    .Item("USER_ID") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                    .Item("OLD_VALUE") = ORIG_STATUS
                    .Item("NEW_VALUE") = NEW_STATUS
                    .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                    .Item("SELECTION_NO") = Me.SELECTION_NO
                    .Item("XNO") = Me.XNO
                    .Item("FM_MODE") = "E"
                    .Item("NOTES") = "SAFSLSA2"
                End With
                dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                Update_Record_TDA("ASTAUDT1")
                CommitTrans()
                ASCMAIN1.MultiTask_Release(, , )
            End If
        Else
            MsgBox("Nothing Was Done.", vbOKOnly, NEW_STATUS_DESC & " Item")
        End If
    End Sub

    Private Sub DiscontinueColor(ByVal DiscType As String)
        If grdSATCSLS1.Selected.Rows.Count <> 1 Then
            MsgBox("You Must Select One Row", vbOKOnly, "Row Selection")
            Exit Sub
        End If
        Dim ORIG_STATUS As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_STATUS").Value
        Dim STYLE_CODE As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
        Dim COLOR_CODE As String = grdSATCSLS1.Selected.Rows(0).Cells.Item("COLOR_CODE").Text

        Dim discStyle As Boolean = False
        Dim NEW_STATUS As String = "D"
        Dim NEW_STATUS_DESC As String = "Discontinue"
        If DiscType = "DNR This Color" Then
            NEW_STATUS = "N"
            NEW_STATUS_DESC = "DNR"
        End If
        If NEW_STATUS = "N" And grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Value = "D" Then
            MsgBox("You May Not DNR A Style That Is Already Discontinued", vbOKOnly, "No update Allowed")
            Exit Sub
        End If

        Dim Scheck As New Text.StringBuilder With {.Length = 0}
        Select Case NEW_STATUS
            Case "D"
                Scheck.AppendLine("SELECT COUNT(*) FROM ICTSTYC1 WHERE STYLE_CODE = '" & STYLE_CODE & "' AND STYLE_COLOR_STATUS IN ('A','N')")
            Case "N"
                Scheck.AppendLine("SELECT COUNT(*) FROM ICTSTYC1 WHERE STYLE_CODE = '" & STYLE_CODE & "' AND STYLE_COLOR_STATUS IN ('A')")
        End Select

        ASCMAIN1.sql = Scheck.ToString()
        Dim countLeft As Int16 = Val(ASCDATA1.GetDataValue)
        If countLeft = 1 Then
            Dim xResult As MsgBoxResult
            Dim xMSG As New System.Text.StringBuilder With {.Length = 0}
            xMSG.AppendLine("This is the Last Color.")
            xMSG.AppendLine("Proceeding Will Change The Style As Well.")
            xMSG.AppendLine("Is That What You Want?")
            xResult = MsgBox(xMSG.ToString(), MsgBoxStyle.YesNo, "Last Color Rule")
            If xResult <> MsgBoxResult.Yes Then
                Exit Sub
            Else
                discStyle = True
            End If
        End If
        If ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE) Then
            BeginTrans()
            If discStyle Then
                Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                SQLS.AppendLine(String.Format("UPDATE ICTSTYL1 SET STYLE_STATUS = '{0}' WHERE STYLE_CODE = '{1}'", NEW_STATUS, STYLE_CODE))
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()

                For Each rowSATCSLS1 As DataRow In dst.Tables("SATCSLS1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                    If rowSATCSLS1.Item("STYLE_COLOR_STATUS") & "" <> NEW_STATUS Then
                        If NEW_STATUS = "N" And rowSATCSLS1.Item("STYLE_COLOR_STATUS") & "" = "D" Then
                            'Don't DNR colors that are Discontinued
                        Else
                            rowSATCSLS1.Item("STYLE_COLOR_STATUS") = NEW_STATUS
                            rowSATCSLS1.Item("STYLE_STATUS") = NEW_STATUS
                        End If
                    End If
                Next

                Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
                With rowASTAUDT1
                    .Item("TABLE_NAME") = "ICTSTYL1"
                    .Item("KEY_VALUE") = STYLE_CODE
                    .Item("COLUMN_NAME") = "STYLE_STATUS"
                    .Item("USER_ID") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                    .Item("OLD_VALUE") = ORIG_STATUS
                    .Item("NEW_VALUE") = NEW_STATUS
                    .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                    .Item("SELECTION_NO") = Me.SELECTION_NO
                    .Item("XNO") = Me.XNO
                    .Item("FM_MODE") = "E"
                    .Item("NOTES") = "SAFSLSA2"
                End With
                dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
                Update_Record_TDA("ASTAUDT1")
            End If

            grdSATCSLS1.Selected.Rows(0).Cells.Item("STYLE_COLOR_STATUS").Value = NEW_STATUS

            Dim SQLC As New Text.StringBuilder With {.Length = 0}
            SQLC.AppendLine(String.Format("UPDATE ICTSTYC1 SET STYLE_COLOR_STATUS = '{0}' WHERE STYLE_CODE = '{1}' AND COLOR_CODE = '{2}'", NEW_STATUS, STYLE_CODE, COLOR_CODE))
            ASCMAIN1.sql = SQLC.ToString
            ASCDATA1.ExecuteSQL()

            CommitTrans()
            ASCMAIN1.MultiTask_Release(, , )
        End If
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

#Region "Custom Methods"
    Sub Create_SATCSLS1()

        ASCMAIN1.sql = "Truncate Table " & SATCSLSX
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into " & SATCSLSX & " " _
            & Replace(Replace(Replace(Replace(Replace(Replace(Replace(sqlSATCSLSX,
                                    "'WHSE1'", "'" & Absx1.txtFor("WHSE_CODE1").Text & "'"),
                                    "'WHSE2'", "'" & Absx1.txtFor("WHSE_CODE2").Text & "'"),
                                    "BETWEEN 'Z' AND 'Z'", IIf(chkOpenOrders.Checked, "BETWEEN 'O' AND 'P'", "BETWEEN 'Z' AND 'Z'")),
                                    "'YP1'", "'" & RYP0 & "'"), "'YP2'", "'" & RYP1 & "'"),
                             "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTINVH1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf),
                             "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf)
        ASCDATA1.ExecuteSQL()
        Dim danasql As String = ASCMAIN1.sql


        ASCMAIN1.sql = "Truncate Table " & SATCSLS1
        ASCDATA1.ExecuteSQL()

        Dim DANA As String = chkOpenOrders.CheckedValue

        ASCMAIN1.sql = "Insert into " & SATCSLS1 & " " _
            & Replace(Replace(Replace(Replace(Replace(Replace(Replace(sqlSATCSLS1,
                                    "'WHSE1'", "'" & Absx1.txtFor("WHSE_CODE1").Text & "'"),
                                    "'WHSE2'", "'" & Absx1.txtFor("WHSE_CODE2").Text & "'"),
                                    "BETWEEN 'Z' AND 'Z'", IIf(chkOpenOrders.Checked, "BETWEEN 'O' AND 'P'", "BETWEEN 'Z' AND 'Z'")),
                                    "'YP1'", "'" & RYP0 & "'"), "'YP2'", "'" & RYP1 & "'"),
                             "   and SOTINVH2.CUST_CODE = SOTINVH2.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTINVH1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf),
                             "   and SOTORDR1.CUST_CODE = SOTORDR1.CUST_CODE", "" _
                 & IIf(Absx1.txtFor("CUST_CODE").Text <> "", " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'", "") & vbCrLf _
                 & IIf(Absx1.txtFor("SREP_CODE").Text <> "", " and SOTORDR1.SREP_CODE = '" & Absx1.txtFor("SREP_CODE").Text & "'", "") & vbCrLf)

        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {Absx1.txtFor("WHSE_CODE1").Text, Absx1.txtFor("WHSE_CODE2").Text, RYP0, RYP1})
        ASCDATA1.ExecuteSQL()

        If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("SREP_CODE").Text = "" Then

            Dim ACTIVE_ONLY As String = ""
            If chkActiveStylesOnly.Checked Then
                ACTIVE_ONLY = " and ICTSTYL1.STYLE_STATUS = 'A'"
            End If
            Dim STOCK_STYLES As String = ""
            'If chkStockStyles.Checked = False Or chkNonStockStyles.Checked = False Then
            '    If chkStockStyles.Checked = False Then
            '        STOCK_STYLES = " AND NVL(CUST_CODE,'NULL') = 'NULL'"
            '    End If
            '    If chkNonStockStyles.Checked = False Then
            '        STOCK_STYLES = " AND NVL(CUST_CODE,'NULL') <> 'NULL'"
            '    End If
            'End If
            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("RGI", "")
                Application.DoEvents()


                Dim S As New StringBuilder With {.Length = 0}
                Dim WHSE1 As String = Absx1.txtFor("WHSE_CODE1").Text
                Dim WHSE2 As String = Absx1.txtFor("WHSE_CODE2").Text
                Dim SWHERE As String = ""
                If WHSE1.Length > 0 Then
                    SWHERE = $"'{WHSE1}'"
                End If
                If WHSE2.Length > 0 Then
                    If SWHERE.Length = 0 Then
                        SWHERE = $"'{WHSE2}'"
                    Else
                        SWHERE = $"{SWHERE},'{WHSE2}'"
                    End If
                End If
                If SWHERE.Length > 0 Then
                    SWHERE = $"    WHERE WHSE_CODE IN ({SWHERE})"
                End If

                S.AppendLine("INSERT INTO " & SATCSLS1)
                S.AppendLine("(STYLE_CODE,COLOR_CODE,STYLE_COLOR_STATUS, STYLE_DESC,STYLE_STATUS,VEND_CODE,FACTORY_CODE,STYLE_UOM,STYLE_CLASS_CODE,CARTON_PACK_QTY,COLOR_DESC,PO_COST, CUST_NAME, FUT_QTY)")
                S.AppendLine("SELECT X.STYLE_CODE, X.COLOR_CODE, X.STYLE_COLOR_STATUS,")
                S.AppendLine("ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.CARTON_PACK_QTY,")
                S.AppendLine("ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST, ARTCUST1.CUST_NAME, Y.FUT_QTY")
                S.AppendLine("FROM ICTSTYL1,ICTCOLR1,ICTSTYV1, ARTCUST1,")
                S.AppendLine("(")
                S.AppendLine("  SELECT ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYC1.STYLE_COLOR_STATUS")
                S.AppendLine("  FROM ICTSTYC1,ICTSTYL1 where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & ACTIVE_ONLY & STOCK_STYLES)
                S.AppendLine("  MINUS SELECT STYLE_CODE, COLOR_CODE, STYLE_COLOR_STATUS from " & SATCSLS1)
                S.AppendLine(") X,")
                S.AppendLine("(")
                S.AppendLine("    SELECT STYLE_CODE, COLOR_CODE,")
                S.AppendLine("    SUM (NVL(WHSE_QTY_ON_HAND,0)+NVL(WHSE_QTY_ON_ORDER,0)+NVL(WHSE_QTY_TRAN,0)-NVL(WHSE_QTY_OPEN,0)-NVL(WHSE_QTY_PICK,0)) FUT_QTY")
                S.AppendLine("    FROM ICTSTAT2")
                S.AppendLine(SWHERE)
                S.AppendLine("    GROUP BY STYLE_CODE, COLOR_CODE")
                S.AppendLine(") Y")
                S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = X.STYLE_CODE")
                S.AppendLine("AND ICTCOLR1.COLOR_CODE = X.COLOR_CODE")
                S.AppendLine("AND ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE")
                S.AppendLine("AND ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE")
                S.AppendLine("AND ICTSTYL1.CUST_CODE = ARTCUST1.CUST_CODE (+)")
                S.AppendLine("AND X.STYLE_CODE (+) = Y.STYLE_CODE")
                S.AppendLine("AND X.COLOR_CODE (+) = Y.COLOR_CODE")
                ASCMAIN1.sql = S.ToString

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
                Application.DoEvents()
            Else
                ASCMAIN1.sql = "Insert into " & SATCSLS1 & " (STYLE_CODE,COLOR_CODE,STYLE_COLOR_STATUS, STYLE_DESC,STYLE_STATUS,VEND_CODE,FACTORY_CODE,STYLE_UOM,STYLE_CLASS_CODE,CARTON_PACK_QTY,COLOR_DESC,PO_COST, CUST_NAME) " _
                    & " Select X.STYLE_CODE, X.COLOR_CODE, X.STYLE_COLOR_STATUS" & vbCrLf _
                    & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYL1.VEND_CODE, ICTSTYL1.FACTORY_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.CARTON_PACK_QTY" & vbCrLf _
                    & ", ICTCOLR1.COLOR_DESC, CASE WHEN NVL(NEW_PO_COST_DATE,TRUNC(SYSDATE+1)) <= TRUNC(SYSDATE) THEN NEW_PO_COST ELSE PO_COST END PO_COST, ARTCUST1.CUST_NAME" & vbCrLf _
                    & " from ICTSTYL1,ICTCOLR1,ICTSTYV1, ARTCUST1" & vbCrLf _
                    & ", (Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTSTYC1.STYLE_COLOR_STATUS from ICTSTYC1,ICTSTYL1" & vbCrLf _
                    & "     where ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" & ACTIVE_ONLY & STOCK_STYLES & vbCrLf _
                    & "   minus Select STYLE_CODE, COLOR_CODE, STYLE_COLOR_STATUS from " & SATCSLS1 & ") X" & vbCrLf _
                    & " where ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                    & "   and ICTCOLR1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                    & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                    & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" _
                    & "   AND ICTSTYL1.CUST_CODE = ARTCUST1.CUST_CODE (+)"
            End If

            ASCDATA1.ExecuteSQL()

            chkShow0Sales.Visible = True
        Else
            chkShow0Sales.Visible = False

        End If

        If chkStockStyles.Checked = False Or chkNonStockStyles.Checked = False Then
            Dim sqlSNS As New StringBuilder With {.Length = 0}
            sqlSNS.AppendLine(String.Format("DELETE FROM {0}", SATCSLS1))
            sqlSNS.AppendLine(" WHERE STYLE_CODE IN ")
            If chkStockStyles.Checked = False Then
                sqlSNS.AppendLine("(SELECT STYLE_CODE FROM ICTSTYL1 WHERE NVL(CUST_CODE,'NULL') = 'NULL')")
            End If
            If chkNonStockStyles.Checked = False Then
                sqlSNS.AppendLine("(SELECT STYLE_CODE FROM ICTSTYL1 WHERE NVL(CUST_CODE,'NULL') <> 'NULL')")
            End If
            ASCMAIN1.sql = sqlSNS.ToString
            ASCDATA1.ExecuteSQL()
        End If

        Fill_Records("SATCSLS1")
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            Dim sql As New Text.StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT")
            sql.AppendLine("Y1.STYLE_CODE")
            sql.AppendLine("FROM ECTESTY1 Y1, ECTECOM1 E1")
            sql.AppendLine("WHERE Y1.ECOM_CODE = E1.ECOM_CODE")
            sql.AppendLine("AND (NVL(Y1.SHIP_ECOM,'0') = '1' OR NVL(Y1.SHIP_DROP,'0') = '1')")
            sql.AppendLine("GROUP BY Y1.STYLE_CODE")
            Dim tblECOM As DataTable = ASCDATA1.GetDataTable(sql.ToString())
            For Each rowSATCSLS1 As DataRow In dst.Tables("SATCSLS1").Select()
                Dim EFilter As String = String.Format("STYLE_CODE = '{0}'", rowSATCSLS1.Item("STYLE_CODE").ToString & String.Empty)
                If tblECOM.Select(EFilter).Count > 0 Then
                    rowSATCSLS1.Item("IS_ECOM") = "1"
                Else
                    rowSATCSLS1.Item("IS_ECOM") = "0"
                End If
            Next
        End If
    End Sub

    Sub Set_DataSource()
        Dim dvw As DataView = DirectCast(grdSATCSLS1.DataSource, DataTable).DefaultView
        If chkShow0Sales.Checked Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "SLS_QTY <> 0"
        End If
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
                    & ", SOTORDR1.ORDR_DATE_RECD" & vbCrLf _
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

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

#End Region

#Region "Form Controls"
    Private Sub chkShowDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Private Sub chkShow0Sales_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShow0Sales.CheckedChanged
        Set_DataSource()
    End Sub

    Private Sub grdSATCSLS1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSATCSLS1.AfterRowActivate
        Setup_grdSOTINVHX()
        If Not grdSATCSLS1.ActiveRow Is Nothing And grdSATCSLS1.ActiveRow.IsDataRow Then
            Dim STYLE_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("STYLE_CODE").Text & String.Empty
            Dim COLOR_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("COLOR_CODE").Text & String.Empty
            If STYLE_CODE_IMG.Length > 0 And COLOR_CODE_IMG.Length > 0 Then
                FetchImage(STYLE_CODE_IMG, COLOR_CODE_IMG)
            End If
            EcomIndicator()
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

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs)
        Setup_tabDetails()
    End Sub
#End Region

    Function FetchImage(ByVal STYLE_CODE_IMG As String, ByVal COLOR_CODE_IMG As String) As Byte()
        Dim IMAGE_NAME As String = STYLE_CODE_IMG & "-" & COLOR_CODE_IMG

        Dim imgba() As Byte = Nothing
        If IMAGE_NAME <> "" Then
            imgSTYLE.Image = Get_Style_Image(IMAGE_NAME, imgba)
            UltraExplorerBar1.Groups("Style Image").Text = "Style " & STYLE_CODE_IMG & "-" & COLOR_CODE_IMG
        Else
            imgSTYLE.Image = Nothing
            UltraExplorerBar1.Groups("Style Image").Text = "Style Image"
        End If

        Return imgba
    End Function

    Function Get_Style_Image(
        ByVal IMAGE_NAME As String,
        Optional ByRef imgba() As Byte = Nothing) As System.Drawing.Bitmap
        Dim rowICTPARM1 As DataRow = LookUp("ICTPARM1", "Z")
        Dim FOLDER_NAME As String = rowICTPARM1.Item("IC_PARM_STYLE_IMG_DIR") & ""

        'Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        Return ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
    End Function

    Private Sub imgSTYLE_DoubleClick(sender As Object, e As System.EventArgs) Handles imgSTYLE.DoubleClick
        If Not IsNothing(grdSATCSLS1.ActiveRow) And grdSATCSLS1.ActiveRow.IsDataRow Then
            Dim STYLE_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("STYLE_CODE").Text & String.Empty
            Dim STYLE_DESC_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("STYLE_DESC").Text & String.Empty
            Dim COLOR_CODE_IMG As String = grdSATCSLS1.ActiveRow.Cells.Item("COLOR_CODE").Text & String.Empty
            Using F As New ASFMSGBF
                F.Show_img(imgSTYLE.Image, Me, "Style " & STYLE_CODE_IMG & ":" & STYLE_DESC_IMG)
            End Using
        End If
    End Sub

    Private Sub EcomIndicator()
        Try
            If Not (grdSATCSLS1.ActiveRow Is Nothing OrElse Not grdSATCSLS1.ActiveRow.IsDataRow) Then
                Dim STYLE_CODE As String = grdSATCSLS1.ActiveRow.Cells("STYLE_CODE").Value
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim ECOM_MSG As String = TAC.TACMAIN1.getEcomInfo(Me, STYLE_CODE)
                    If ECOM_MSG.Length > 0 Then
                        lblEcomStyle.Visible = True
                        Dim TTI As New UltraWinToolTip.UltraToolTipInfo
                        If Not IsNothing(TTM.GetUltraToolTip(lblEcomStyle)) Then
                            TTI.ToolTipTitle = "E-Commerce Information:"
                            TTM.AutoPopDelay = 20000
                            TTI.ToolTipTextFormatted = ECOM_MSG
                            TTM.SetUltraToolTip(lblEcomStyle, TTI)
                        Else
                            TTI.ToolTipTextFormatted = ECOM_MSG
                        End If
                    Else
                        lblEcomStyle.Visible = False
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub chkStockStyles_CheckedChanged(sender As Object, e As EventArgs) Handles chkStockStyles.CheckedChanged, chkStockStyles.CheckedChanged
        If Not isDataLoading Then
            If chkStockStyles.Checked = False And chkNonStockStyles.Checked = False Then
                MsgBox("You Must Choose Stock, Non-Stock Or Both", vbOKOnly, "Stock Error")
                chkStockStyles.Checked = True
                chkNonStockStyles.Checked = True
            End If
        End If
    End Sub
End Class