Imports System.Math

Public Class ARRDEFI1

    Private xRYP0 As String = String.Empty
    Private xRYP1 As String = String.Empty
    Private ARTDEFI1 As String = String.Empty
    Shadows SUBT As String = String.Empty
    Private sqlWarehouse As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 0, 0)
        UltraTabControl1.Tabs.Item(5).Visible = False
        UltraTabControl1.Tabs.Item(5).Text = "Matrix"
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            Create_Summary(grdMATRIX, "YTD_NET_UNITS", "Sum", "", "###,##0")
            Create_Summary(grdMATRIX, "YTD_DEF_UNITS", "Sum", "", "###,##0")
            Create_Summary(grdMATRIX, "MTRX_DOL_CLAIMED", "Sum", "", "###,##0.00")
            Create_Summary(grdMATRIX, "MTRX_TWO_PCT", "Sum", "", "###,##0.00")
            Load_Popup_Menu(grdMATRIX, "SSB", "Show Filter", "Show GroupBox")
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "N"
        SUBT = String.Empty
        Dim sqlw As String = String.Empty

        ''sqlw &= SQL_in("REASON_CODE")
        ''sqlw = sqlw.Replace(" REASON_CODE", " NVL(SOTRTRN2.RTV_REASON_CODE, SOTRTRN1.REASON_CODE)")


        sqlw &= SQL_in("RTV_REASON_CODE")
        sqlw = sqlw.Replace(" RTV_REASON_CODE", " NVL(SOTRTRN2.RTV_REASON_CODE, SOTRTRN1.REASON_CODE)")

        sqlw &= SQL_in("VEND_CODE")
        sqlw = sqlw.Replace(" VEND_CODE", " ICTSTYL1.VEND_CODE")

        sqlw &= SQL_in("CUST_CODE")
        sqlw = sqlw.Replace(" CUST_CODE", " SOTRTRN1.CUST_CODE")

        sqlw &= SQL_in("STYLE_CODE")
        sqlw = sqlw.Replace(" STYLE_CODE", " SOTRTRN2.STYLE_CODE")

        sqlWarehouse = SQL_in("WHSE_CODE")
        sqlWarehouse = sqlWarehouse.Replace(" WHSE_CODE", " SOTRTRN1.WHSE_CODE")

        sqlw &= sqlWarehouse

        Prepare_dst(True, sqlw)

        Check_if_Empty("ARTDEFI1")
    End Sub

    Public Overrides Sub Print_Report()
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            FillMatrixFields()
            UltraTabControl1.Tabs.Item(5).Visible = True
            grdMATRIX.DataSource = dst.Tables("ARTDEFI1")
            'Create_Summary(grdMATRIX, "YTD_NET_UNITS", "Sum", "", "###,##0")
            'Create_Summary(grdMATRIX, "YTD_DEF_UNITS", "Sum", "", "###,##0")
            'Create_Summary(grdMATRIX, "MTRX_DOL_CLAIMED", "Sum", "", "###,##0.00")
            'Create_Summary(grdMATRIX, "MTRX_TWO_PCT", "Sum", "", "###,##0.00")
            'Load_Popup_Menu(grdMATRIX, "SSB", "Show Filter", "Show GroupBox")
        End If
        RPT = "ARRDEFI2" ' IIf(MENU_ITEM_FORM = "", MENU_ITEM_OBJECT, MENU_ITEM_FORM)
        SUBT = "For Period: " & Absx1.cmbFor("RYP0").SelectedRow.Cells("LEGEND").Value & " Thru " & Absx1.cmbFor("RYP1").SelectedRow.Cells("LEGEND").Value

        If sqlWarehouse.Length = 0 Then
            SUBT &= " - All Warehouses "
        ElseIf sqlWarehouse.Contains("=") Then
            SUBT &= " - Warehouses: " & sqlWarehouse.Split("=")(1).Replace("'", "").Trim
        ElseIf sqlWarehouse.Contains("(") Then
            SUBT &= " - Warehouses: " & sqlWarehouse.Split("(")(1).Replace("'", "").Replace(")", "").Trim
        End If
        If optDS.Value = "S" Then
            SUBT &= " Stock Qty Returns "
        Else
            SUBT &= " Destoyed Qty Returns "
        End If
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP0").Value & "" = "" Or Absx1.cmbFor("RYP1").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Report Periods"
            Else
                xRYP0 = Absx1.cmbFor("RYP0").SelectedRow.Cells("OPS_YYYYPP").Value
                xRYP1 = Absx1.cmbFor("RYP1").SelectedRow.Cells("OPS_YYYYPP").Value
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Create_TDA(dst.Tables.Add, "ARTREAS1", "SELECT REASON_CODE, REASON_DESC FROM ARTREAS1")
        Create_TDA(dst.Tables.Add, "APTVEND1", "SELECT VEND_CODE, VEND_NAME FROM APTVEND1")
        'Create_TDA(dst.Tables.Add, "ICTSTYV1", "*")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = ""

        sql = " SELECT NVL(ICTSTYL1.VEND_CODE, '?') VEND_CODE, SOTRTRN2.STYLE_CODE, SOTRTRN2.COLOR_CODE"
        sql &= ", ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_DESC, SOTRTRN2.RTV_REASON_CODE REASON_CODE"
        sql &= ", 0.00 MTD_NET_SALES, 0.00 YTD_NET_SALES"
        sql &= ", 0 MTD_NET_UNITS, 0 YTD_NET_UNITS"
        sql &= ", 0.00 MTD_DEF_CLAIMS,  0.00 YTD_DEF_CLAIMS"
        sql &= ", 0 MTD_DEF_UNITS,  0 YTD_DEF_UNITS"
        sql &= ", 0.00 MTD_PER_SALES, 0.00 YTD_PER_SALES"
        sql &= ", 0.00 MTD_PER_UNITS, 0.00 YTD_PER_UNITS"
        sql &= ", 0.00 STYLE_COST, 0.00 STYLE_COST_EXT"
        sql &= " FROM SOTRTRN1, SOTRTRN2, ICTSTYL1"
        sql &= " WHERE SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO"
        sql &= " AND SOTRTRN2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)"
        sql &= " AND ROWNUM < 1"
        ARTDEFI1 = ASCMAIN1.Temp_Table(sql)

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY MTD_NET_SALES NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY YTD_NET_SALES NUMBER(13,2)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY MTD_NET_UNITS NUMBER(8)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY YTD_NET_UNITS NUMBER(8)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY MTD_DEF_CLAIMS NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY YTD_DEF_CLAIMS NUMBER(13,2)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY MTD_DEF_UNITS NUMBER(8)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY YTD_DEF_UNITS NUMBER(8)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY MTD_PER_SALES NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY YTD_PER_SALES NUMBER(13,2)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY MTD_PER_UNITS NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY YTD_PER_UNITS NUMBER(13,2)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY STYLE_COST NUMBER(13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " MODIFY STYLE_COST_EXT NUMBER(13,2)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTDEFI1 & " ADD PRIMARY KEY (VEND_CODE, STYLE_CODE, COLOR_CODE, REASON_CODE)")

        sql = "SELECT * FROM " & ARTDEFI1
        Create_TDA(dst.Tables.Add("ARTDEFI1"), ARTDEFI1, "*")
        With dst.Tables("ARTDEFI1")
            .Columns.Add("REASON_DESC", GetType(System.String))
            .Columns.Add("MTRX_PCT_CLAIMED", GetType(System.Double))
            .Columns.Add("MTRX_DOL_CLAIMED", GetType(System.Double))
            .Columns.Add("MTRX_TWO_PCT", GetType(System.Double))
        End With
        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        EnforceConstraints(False)

        Dim sql As String = String.Empty
        'sql = "SELECT REASON_CODE, INITCAP(REASON_DESC) REASON_DESC FROM ARTREAS1 WHERE REASON_CODE NOT IN (Select RTV_REASON_CODE from SOTREASV)"
        'sql &= " Union "
        sql = "Select RTV_REASON_CODE REASON_CODE, INITCAP(RTV_REASON_DESC) REASON_DESC from SOTREASV"

        Fill_Records("ARTREAS1", String.Empty, True, sql)
        Fill_Records("APTVEND1", String.Empty, True, "SELECT VEND_CODE, INITCAP(VEND_NAME) VEND_NAME FROM APTVEND1")

        ASCMAIN1.Progress("Gather Returns", "")
        ' Get all the details for the current year that have returns


        If optDS.Value = "S" Then
            sql = " SELECT NVL(ICTSTYL1.VEND_CODE, '?') VEND_CODE, SOTRTRN2.STYLE_CODE, SOTRTRN2.COLOR_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_DESC, NVL(SOTRTRN2.RTV_REASON_CODE, SOTRTRN1.REASON_CODE) REASON_CODE"
            sql &= ", 0.00 MTD_NET_SALES, 0.00 YTD_NET_SALES"
            sql &= ", CASE SOTRTRN2.OPS_YYYYPP WHEN '" & xRYP0 & "' THEN (SOTRTRN2.RTRN_QTY_1 * NVL(SOTRTRN2.RTRN_PRICE, 0)) ELSE 0 END MTD_DEF_CLAIMS"
            'sql &= ", CASE SUBSTR(SOTRTRN2.OPS_YYYYPP, 1, 4) WHEN '" & xRYP0.Substring(0, 4) & "' THEN (SOTRTRN2.RTRN_QTY_1 * NVL(SOTRTRN2.RTRN_PRICE, 0)) ELSE 0 END YTD_DEF_CLAIMS"
            sql &= ", CASE WHEN SOTRTRN2.OPS_YYYYPP  BETWEEN '" & xRYP0 & "' AND '" & xRYP1 & "' THEN (SOTRTRN2.RTRN_QTY_1 * NVL(SOTRTRN2.RTRN_PRICE, 0)) ELSE 0 END YTD_DEF_CLAIMS"
            sql &= ", CASE SOTRTRN2.OPS_YYYYPP WHEN '" & xRYP0 & "' THEN SOTRTRN2.RTRN_QTY_1  ELSE 0 END MTD_DEF_UNITS"
            'sql &= ", CASE SUBSTR(SOTRTRN2.OPS_YYYYPP, 1, 4) WHEN '" & xRYP0.Substring(0, 4) & "' THEN SOTRTRN2.RTRN_QTY_1  ELSE 0 END YTD_DEF_UNITS"
            sql &= ", CASE  WHEN SOTRTRN2.OPS_YYYYPP  BETWEEN '" & xRYP0 & "' AND '" & xRYP1 & "' THEN SOTRTRN2.RTRN_QTY_1  ELSE 0 END YTD_DEF_UNITS"
            sql &= ", ICTSTYV1.PO_COST STYLE_COST"
            sql &= " FROM SOTRTRN1, SOTRTRN2, ICTSTYL1, ICTSTYV1"
            sql &= " WHERE SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO"
            sql &= " AND SOTRTRN2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)"
            sql &= " AND SOTRTRN2.STYLE_CODE = ICTSTYV1.STYLE_CODE (+)"
            sql &= " AND SOTRTRN2.RTRN_QTY_1 <> 0"
            'sql &= " AND SOTRTRN1.OPS_YYYYPP BETWEEN '" & xRYP0.Substring(0, 4) & "01' AND '" & xRYP0 & "'"
            sql &= " AND SOTRTRN1.OPS_YYYYPP BETWEEN '" & xRYP0 & "' AND '" & xRYP1 & "'"
        Else

            sql = " SELECT NVL(ICTSTYL1.VEND_CODE, '?') VEND_CODE, SOTRTRN2.STYLE_CODE, SOTRTRN2.COLOR_CODE, ICTSTYL1.STYLE_UOM, ICTSTYL1.STYLE_DESC, NVL(SOTRTRN2.RTV_REASON_CODE, SOTRTRN1.REASON_CODE) REASON_CODE"
            sql &= ", 0.00 MTD_NET_SALES, 0.00 YTD_NET_SALES"
            sql &= ", CASE SOTRTRN2.OPS_YYYYPP WHEN '" & xRYP0 & "' THEN (SOTRTRN2.RTRN_QTY_3 * NVL(SOTRTRN2.RTRN_PRICE, 0)) ELSE 0 END MTD_DEF_CLAIMS"
            'sql &= ", CASE SUBSTR(SOTRTRN2.OPS_YYYYPP, 1, 4) WHEN '" & xRYP0.Substring(0, 4) & "' THEN (SOTRTRN2.RTRN_QTY_3 * NVL(SOTRTRN2.RTRN_PRICE, 0)) ELSE 0 END YTD_DEF_CLAIMS"
            sql &= ", CASE WHEN SOTRTRN2.OPS_YYYYPP  BETWEEN '" & xRYP0 & "' AND '" & xRYP1 & "' THEN (SOTRTRN2.RTRN_QTY_3 * NVL(SOTRTRN2.RTRN_PRICE, 0)) ELSE 0 END YTD_DEF_CLAIMS"
            sql &= ", CASE SOTRTRN2.OPS_YYYYPP WHEN '" & xRYP0 & "' THEN SOTRTRN2.RTRN_QTY_3  ELSE 0 END MTD_DEF_UNITS"
            'sql &= ", CASE SUBSTR(SOTRTRN2.OPS_YYYYPP, 1, 4) WHEN '" & xRYP0.Substring(0, 4) & "' THEN SOTRTRN2.RTRN_QTY_3  ELSE 0 END YTD_DEF_UNITS"
            sql &= ", CASE  WHEN SOTRTRN2.OPS_YYYYPP  BETWEEN '" & xRYP0 & "' AND '" & xRYP1 & "' THEN SOTRTRN2.RTRN_QTY_3  ELSE 0 END YTD_DEF_UNITS"
            sql &= ", ICTSTYV1.PO_COST STYLE_COST"
            sql &= " FROM SOTRTRN1, SOTRTRN2, ICTSTYL1, ICTSTYV1"
            sql &= " WHERE SOTRTRN1.RTRN_NO = SOTRTRN2.RTRN_NO"
            sql &= " AND SOTRTRN2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)"
            sql &= " AND SOTRTRN2.STYLE_CODE = ICTSTYV1.STYLE_CODE (+)"
            sql &= " AND SOTRTRN2.RTRN_QTY_3 > 0"
            'sql &= " AND SOTRTRN1.OPS_YYYYPP BETWEEN '" & xRYP0.Substring(0, 4) & "01' AND '" & xRYP0 & "'"
            sql &= " AND SOTRTRN1.OPS_YYYYPP BETWEEN '" & xRYP0 & "' AND '" & xRYP1 & "'"
        End If

        sql &= parms(0)
        Dim wktable As String = ASCMAIN1.Temp_Table(sql)

        ASCDATA1.ExecuteSQL("UPDATE " & wktable & " SET REASON_CODE = '?' WHERE REASON_CODE IS NULL")

        ASCMAIN1.Progress("Sum Data", "")
        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & ARTDEFI1)
        sql = "INSERT INTO " & ARTDEFI1
        sql &= " SELECT VEND_CODE, STYLE_CODE, COLOR_CODE, STYLE_UOM, STYLE_DESC, REASON_CODE"
        sql &= ", SUM(0) MTD_NET_SALES, SUM(0) YTD_NET_SALES"
        sql &= ", SUM(0) MTD_NET_UNITS, SUM(0) YTD_NET_UNITS"
        sql &= ", SUM(MTD_DEF_CLAIMS) MTD_DEF_CLAIMS"
        sql &= ", SUM(YTD_DEF_CLAIMS) YTD_DEF_CLAIMS"
        sql &= ", SUM(MTD_DEF_UNITS) MTD_DEF_UNITS"
        sql &= ", SUM(YTD_DEF_UNITS) YTD_DEF_UNITS"
        sql &= ", SUM(0) MTD_PER_SALES"
        sql &= ", SUM(0) YTD_PER_SALES"
        sql &= ", SUM(0) MTD_PER_UNITS"
        sql &= ", SUM(0) YTD_PER_UNITS"
        sql &= ", STYLE_COST"
        sql &= ", SUM(0) STYLE_COST_EXT"
        sql &= " from " & wktable
        sql &= " group by VEND_CODE, STYLE_CODE, COLOR_CODE, STYLE_UOM, STYLE_DESC, REASON_CODE, STYLE_COST"
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.Progress("Update Monthly Sales/Units", "")
        sql = "BEGIN DECLARE CURSOR C1 IS SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_UNIT_PRICE * ORDR_QTY_SHIP) MTD_NET_SALES, SUM(ORDR_QTY_SHIP) MTD_NET_UNITS"
        sql &= " FROM SOTINVH2"
        sql &= " WHERE ORDR_YYYYPP_UPDATED = '" & xRYP0 & "'"
        sql &= " AND INV_TYPE = 'I'"
        sql &= " AND (STYLE_CODE, COLOR_CODE) IN (SELECT STYLE_CODE, COLOR_CODE FROM " & ARTDEFI1 & ") GROUP BY STYLE_CODE, COLOR_CODE;"
        sql &= " BEGIN FOR R1 IN C1 LOOP "
        sql &= "    UPDATE " & ARTDEFI1 & " SET MTD_NET_SALES = R1.MTD_NET_SALES,  MTD_NET_UNITS = R1.MTD_NET_UNITS WHERE STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE;"
        sql &= "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.Progress("Update Yearly Sales/Units", "")
        sql = "BEGIN DECLARE CURSOR C1 IS SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_UNIT_PRICE * ORDR_QTY_SHIP) YTD_NET_SALES, SUM(ORDR_QTY_SHIP) YTD_NET_UNITS"
        sql &= " FROM SOTINVH2"
        sql &= " WHERE ORDR_YYYYPP_UPDATED BETWEEN '" & xRYP0 & "' AND '" & xRYP1 & "'"
        sql &= " AND INV_TYPE = 'I'"
        sql &= " AND (STYLE_CODE, COLOR_CODE) IN (SELECT STYLE_CODE, COLOR_CODE FROM " & ARTDEFI1 & ") GROUP BY STYLE_CODE, COLOR_CODE;"
        sql &= "  BEGIN FOR R1 IN C1 LOOP "
        sql &= "    UPDATE " & ARTDEFI1 & " SET YTD_NET_SALES = R1.YTD_NET_SALES, YTD_NET_UNITS = R1.YTD_NET_UNITS WHERE STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE;"
        sql &= "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

        For Each field As String In New String() {"MTD_NET_SALES", "YTD_NET_SALES", "MTD_DEF_CLAIMS", "YTD_DEF_CLAIMS",
                                                  "MTD_NET_UNITS", "YTD_NET_UNITS", "MTD_DEF_UNITS", "YTD_DEF_UNITS", "MTD_PER_UNITS", "YTD_PER_UNITS"}
            ASCDATA1.ExecuteSQL("Update " & ARTDEFI1 & " set " & field & " = 0 where " & field & " is null")
        Next

        ' Delete any extra items. I do not think there will be any.
        sql = "Delete from " & ARTDEFI1
        sql &= " where MTD_NET_SALES = 0 and YTD_NET_SALES = 0 and MTD_DEF_CLAIMS = 0 and YTD_DEF_CLAIMS = 0"
        sql &= " and MTD_NET_UNITS = 0 and YTD_NET_UNITS = 0 and MTD_DEF_UNITS = 0 and YTD_DEF_UNITS  = 0"
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.Progress("Calculate Percentages", "")
        sql = "Update " & ARTDEFI1 & " set MTD_PER_SALES = 0, YTD_PER_SALES = 0"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update " & ARTDEFI1 & " set MTD_PER_SALES =  (MTD_DEF_CLAIMS / MTD_NET_SALES) * 100 where MTD_NET_SALES <> 0"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update " & ARTDEFI1 & " set YTD_PER_SALES =  (YTD_DEF_CLAIMS / YTD_NET_SALES) * 100 where YTD_NET_SALES <> 0"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update " & ARTDEFI1 & " set MTD_PER_UNITS =  (MTD_DEF_UNITS / MTD_NET_UNITS) * 100 where MTD_NET_UNITS <> 0"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update " & ARTDEFI1 & " set YTD_PER_UNITS =  (YTD_DEF_UNITS / YTD_NET_UNITS) * 100 where YTD_NET_UNITS <> 0"
        ASCDATA1.ExecuteSQL(sql)

        Fill_Records("ARTDEFI1", String.Empty, True, "SELECT * FROM " & ARTDEFI1)

        EnforceConstraints(True)

        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub FillMatrixFields()
        For Each rowARTDEFI1 As DataRow In dst.Tables("ARTDEFI1").Select()
            Dim REASON_DESC As String = ""
            Dim YTD_NET_UNITS As Double = Val(rowARTDEFI1.Item("YTD_NET_UNITS") & "")
            Dim YTD_DEF_UNITS As Double = Val(rowARTDEFI1.Item("YTD_DEF_UNITS") & "")
            Dim STYLE_COST As Double = Val(rowARTDEFI1.Item("STYLE_COST") & "")
            Dim REASON_CODE As String = rowARTDEFI1.Item("REASON_CODE") & ""
            Dim MTRX_PCT_CLAIMED As Double = 0
            Dim MTRX_TWO_PCT As Double = 0
            If YTD_NET_UNITS = 0 Then
                MTRX_PCT_CLAIMED = 0
            Else
                MTRX_PCT_CLAIMED = (YTD_DEF_UNITS / YTD_NET_UNITS) * 100
            End If

            If MTRX_PCT_CLAIMED >= 2 Then
                MTRX_TWO_PCT = STYLE_COST * YTD_DEF_UNITS
            Else
                MTRX_TWO_PCT = 0
            End If

            Dim rowARTREAS1 As DataRow = dst.Tables.Item("ARTREAS1").Select(String.Format("REASON_CODE = '{0}'", REASON_CODE)).FirstOrDefault()
            If Not IsNothing(rowARTREAS1) Then
                REASON_DESC = rowARTREAS1.Item("REASON_DESC").ToString() & ""
            End If
            rowARTDEFI1.Item("REASON_DESC") = REASON_DESC

            rowARTDEFI1.Item("MTRX_PCT_CLAIMED") = MTRX_PCT_CLAIMED

            rowARTDEFI1.Item("MTRX_DOL_CLAIMED") = STYLE_COST * YTD_DEF_UNITS

            rowARTDEFI1.Item("MTRX_TWO_PCT") = MTRX_TWO_PCT
        Next
    End Sub

End Class