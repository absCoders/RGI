Imports Microsoft.Office.Interop

Public Class SORCMSM1

    Dim MD() As String
    Dim S As New Text.StringBuilder With {.Length = 0}
    Dim TABLE_TEMP As String = ""
    Dim CUST_LIST As New List(Of String)
    Dim YEARS As New List(Of String)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim CYEAR As String = ASCMAIN1.CYP.Substring(0, 4)
        Dim cboYEARS_DAT As New List(Of String)
        For i As Int64 = Val(CYEAR) To Val(CYEAR) - 10 Step -1
            cboYEARS_DAT.Add($"{i.ToString}")
        Next
        cboBaseYear.DataSource = cboYEARS_DAT
        cboBaseYear.SelectedIndex = 0
        numYears.Value = 1

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables
        'ReDim MD(12)
        'For i As Integer = 1 To 12
        '    Dim Z As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP0, (i - 1)))
        '    MD(i) = Mid$(Z, 10, 6)
        'Next i


        ' Prepare filters from Run-Time Options

        MyBase.Get_SQL("*")

        'Build the temp Table.
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("ARTCUST1.CUST_CODE,")
        S.AppendLine("SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4) AS YR,")
        S.AppendLine("1 as RPT_ORDER,")
        S.AppendLine("'S' AS DATA_TYPE,")
        S.AppendLine("ARTCUST1.CUST_NAME,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202101', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_01,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202102', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_02,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202103', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_03,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202104', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_04,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202105', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_05,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202106', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_06,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202107', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_07,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202108', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_08,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202109', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_09,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202110', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_10,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202111', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_11,")
        S.AppendLine("SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'202112', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) MO_12")
        S.AppendLine("FROM SOTINVH2,ICTSTYL1,SOTINVH1, ARTCUST1")
        S.AppendLine("WHERE SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE")
        S.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE")
        S.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
        S.AppendLine("AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
        S.AppendLine("AND SOTINVH2.CUST_CODE = NULL")
        S.AppendLine("GROUP BY ARTCUST1.CUST_CODE, SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4), ARTCUST1.CUST_NAME")
        ASCMAIN1.sql = S.ToString
        TABLE_TEMP = ASCMAIN1.Temp_Table

        Dim CUST_STR As String = ""
        For Each CUST As String In CUST_LIST
            CUST_STR = CUST_STR & CUST & "','"
        Next
        CUST_STR = "'" & CUST_STR.Substring(0, CUST_STR.Length - 2)

        For Each CUST As String In CUST_LIST
            For Each YR As String In YEARS
                ASCMAIN1.Progress($"Processing {CUST} for {YR}", "")
                'Insert Sales(S) into temp table.
                S.Length = 0
                S.AppendLine($"INSERT INTO {TABLE_TEMP}")
                S.AppendLine("SELECT")
                S.AppendLine("ARTCUST1.CUST_CODE,")
                S.AppendLine("SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4) AS YR,")
                S.AppendLine("1 as RPT_ORDER,")
                S.AppendLine("'S' AS DATA_TYPE,")
                S.AppendLine("ARTCUST1.CUST_NAME,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}01', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_01,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}02', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_02,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}03', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_03,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}04', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_04,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}05', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_05,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}06', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_06,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}07', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_07,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}08', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_08,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}09', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_09,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}10', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_10,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}11', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_11,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}12', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) TY_12")
                S.AppendLine("FROM SOTINVH2,ICTSTYL1,SOTINVH1, ARTCUST1")
                S.AppendLine("WHERE SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE")
                S.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE")
                S.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
                S.AppendLine("AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
                S.AppendLine("AND SOTINVH1.INV_TYPE = 'I'")
                S.AppendLine($"AND SOTINVH1.CUST_CODE = '{CUST}'")
                S.AppendLine("AND (SOTINVH2.CUST_CODE <> 'SAMPLES' AND SOTINVH2.CUST_CODE <> 'TRANSFERS')")
                S.AppendLine($"AND SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN '{YR}01' and '{YR}12'")
                S.AppendLine("GROUP BY ARTCUST1.CUST_CODE, SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4), ARTCUST1.CUST_NAME")
                ASCMAIN1.sql = S.ToString
                ASCDATA1.ExecuteSQL()

                'Insert Costs(C) into temp table.
                S.Length = 0
                S.AppendLine($"INSERT INTO {TABLE_TEMP}")
                S.AppendLine("SELECT")
                S.AppendLine("ARTCUST1.CUST_CODE,")
                S.AppendLine("SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4) AS YR,")
                S.AppendLine("2 as RPT_ORDER,")
                S.AppendLine("'C' AS DATA_TYPE,")
                S.AppendLine("ARTCUST1.CUST_NAME,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}01', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_01,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}02', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_02,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}03', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_03,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}04', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_04,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}05', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_05,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}06', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_06,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}07', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_07,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}08', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_08,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}09', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_09,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}10', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_10,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}11', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_11,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}12', NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_12")
                S.AppendLine("FROM SOTINVH2,ICTSTYL1,SOTINVH1, ARTCUST1")
                S.AppendLine("WHERE SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE")
                S.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE")
                S.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
                S.AppendLine("AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
                S.AppendLine("AND SOTINVH1.INV_TYPE = 'I'")
                S.AppendLine($"AND SOTINVH1.CUST_CODE = '{CUST}'")
                S.AppendLine("AND (SOTINVH2.CUST_CODE <> 'SAMPLES' AND SOTINVH2.CUST_CODE <> 'TRANSFERS')")
                S.AppendLine($"AND SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN '{YR}01' and '{YR}12'")
                S.AppendLine("GROUP BY ARTCUST1.CUST_CODE, SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4), ARTCUST1.CUST_NAME")
                ASCMAIN1.sql = S.ToString
                ASCDATA1.ExecuteSQL()

                'Insert Credits(R) into temp table.
                S.Length = 0
                S.AppendLine($"INSERT INTO {TABLE_TEMP}")
                S.AppendLine("SELECT")
                S.AppendLine("ARTCUST1.CUST_CODE,")
                S.AppendLine("SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4) AS YR,")
                S.AppendLine("3 as RPT_ORDER,")
                S.AppendLine("'R' AS DATA_TYPE,")
                S.AppendLine("ARTCUST1.CUST_NAME,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}01', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_01,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}02', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_02,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}03', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_03,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}04', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_04,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}05', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_05,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}06', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_06,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}07', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_07,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}08', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_08,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}09', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_09,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}10', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_10,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}11', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_11,")
                S.AppendLine($"SUM (DECODE (SOTINVH2.ORDR_YYYYPP_UPDATED,'{YR}12', (NVL(SOTINVH2.ORDR_QTY_SHIP,0)* -1) * NVL(SOTINVH2.ORDR_UNIT_COST,0),0)) TY_12")
                S.AppendLine("FROM SOTINVH2,ICTSTYL1,SOTINVH1, ARTCUST1")
                S.AppendLine("WHERE SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE")
                S.AppendLine("AND ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE")
                S.AppendLine("AND SOTINVH1.INV_NO = SOTINVH2.INV_NO")
                S.AppendLine("AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE")
                S.AppendLine("AND SOTINVH1.INV_TYPE = 'C'")
                S.AppendLine($"AND SOTINVH1.CUST_CODE = '{CUST}'")
                S.AppendLine("AND (SOTINVH2.CUST_CODE <> 'SAMPLES' AND SOTINVH2.CUST_CODE <> 'TRANSFERS')")
                S.AppendLine($"AND SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN '{YR}01' and '{YR}12'")
                S.AppendLine("GROUP BY ARTCUST1.CUST_CODE, SUBSTR(SOTINVH2.ORDR_YYYYPP_UPDATED,0,4), ARTCUST1.CUST_NAME")
                ASCMAIN1.sql = S.ToString
                ASCDATA1.ExecuteSQL()

                'Insert Deductions(D) into temp table.
                S.Length = 0
                S.AppendLine($"INSERT INTO {TABLE_TEMP}")
                S.AppendLine("SELECT")
                S.AppendLine("ARTCUST1.CUST_CODE,")
                S.AppendLine("SUBSTR(ARTPYMT1.OPS_YYYYPP,0,4) AS YR,")
                S.AppendLine("4 as RPT_ORDER,")
                S.AppendLine("'D' AS CATEGORY,")
                S.AppendLine("ARTCUST1.CUST_NAME,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}01' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_01,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}02' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_02,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}03' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_03,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}04' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_04,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}05' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_05,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}06' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_06,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}07' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_07,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}08' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_08,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}09' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_09,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}10' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_10,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}11' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_11,")
                S.AppendLine($"Sum (Case When ARTPYMT1.OPS_YYYYPP = '{YR}12' Then NVL(ARTPYMT5.GL_DIST_AMT,0) Else 0 End) TY_12")
                S.AppendLine("from ARTPYMT5, ARTPYMT2 X, ARTPYMT1, ARTCUST1")
                S.AppendLine("where Decode(ARTPYMT5.CUST_CODE_SO, NULL, X.CUST_CODE,ARTPYMT5.CUST_CODE_SO) = ARTCUST1.CUST_CODE")
                S.AppendLine("AND NVL(ARTPYMT5.CHARGEBACK_IND,'0') <> '1'")
                S.AppendLine("and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO")
                S.AppendLine("and ARTPYMT5.PYMT_BATCH_NO = X.PYMT_BATCH_NO")
                S.AppendLine("and ARTPYMT5.PYMT_BATCH_LNO = X.PYMT_BATCH_LNO")
                S.AppendLine($"and ARTPYMT1.OPS_YYYYPP >= '{YR}01'")
                S.AppendLine($"and ARTPYMT1.OPS_YYYYPP <= '{YR}12'")
                S.AppendLine($"AND Decode(ARTPYMT5.CUST_CODE_SO, NULL, X.CUST_CODE,ARTPYMT5.CUST_CODE_SO) = '{CUST}'")
                S.AppendLine("group by ARTCUST1.CUST_CODE, SUBSTR(ARTPYMT1.OPS_YYYYPP,0,4),ARTCUST1.CUST_NAME")
                ASCMAIN1.sql = S.ToString
                ASCDATA1.ExecuteSQL()
            Next
        Next
        With dst

            ASCMAIN1.sql = $"Select * FROM {TABLE_TEMP}"
            Create_TDA(.Tables.Add, "SOTCMSM1", "**", 0, False, "")

        End With
        Fill_Records("SOTCMSM1")
        ASCMAIN1.Progress("", "")
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'Not sure what I need to put in here for this report.
        'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then Stop
        'Throw a record into ASTSRPT1 just to keep the standards from being a jackass.
        Dim newASTSRPT1 As DataRow = dst.Tables.Item("ASTSRPT1").NewRow
        newASTSRPT1.Item("G1") = "1"
        dst.Tables.Item("ASTSRPT1").Rows.Add(newASTSRPT1)
    End Sub

    Public Overrides Sub Print_Report()

        Generate_Excel()

        'If Absx1.chkFor("CHKSORTBYSEL").Checked Then
        '    RPT = "SORSLSDV"
        'End If

        'For i As Integer = 1 To 12
        '    CR_params.Add("MD" & Format$(i, "00"), MD(i))
        'Next i

        'Generate_Report(RPT, , SUBT)
    End Sub

    Private Sub Generate_Excel()
        Dim CUR_SHT As Int64 = 0
        Dim excel As Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Dim XWB As Excel.Workbook = excel.Workbooks.Add
        Dim XWS As Excel.Worksheet = XWB.Sheets(1)
        Dim XWSS As Excel.Worksheet = XWB.Sheets(1)

        Dim CUR_ROW As Int64 = 1
        Dim CUR_COL As Int64 = 1
        Dim MAX_COL As Int64 = 18
        excel.DisplayAlerts = False


        'rng.FormulaR1C1 = "Regency International"
        'rng.Merge()
        'CUR_ROW += 1
        'XWS.Cells(CUR_ROW, CUR_COL).VALUE = 134.56
        'rng.FormulaR1C1 = "Regency International"
        If chkShowSummary.Checked Then
            CUR_SHT += 1
            XWSS.Name = "SUMMARY"
            FormatColumns(XWSS)
            Dim rng As Excel.Range
            rng = XWSS.Range(GetRange(CUR_ROW, CUR_COL, CUR_ROW, MAX_COL))
            rng.FormulaR1C1 = "Summary Sheet"
            rng.MergeCells = True
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            rng.Font.Bold = True
            rng.Interior.Color = System.Drawing.Color.Yellow
            CUR_ROW += 1
        Else
            XWSS = Nothing
        End If

        Dim CUST_STR As String = ""
        For Each CUST As String In CUST_LIST
            Dim CFilter As String = $"CUST_CODE = '{CUST}'"
            If dst.Tables.Item("SOTCMSM1").Select(CFilter).Length > 0 Then
                CUR_SHT += 1
                If CUR_SHT <> 1 Then
                    XWS = XWB.Worksheets.Add()
                    CUR_ROW = 1
                    CUR_COL = 1
                End If
                XWS.Name = CUST
                FormatColumns(XWS)
                Dim rng As Excel.Range
                rng = XWS.Range(GetRange(CUR_ROW, CUR_COL, CUR_ROW, MAX_COL))
                rng.FormulaR1C1 = dst.Tables.Item("SOTCMSM1").Select(CFilter).FirstOrDefault.Item("CUST_NAME").ToString & String.Empty
                rng.MergeCells = True
                rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                rng.Font.Bold = True
                rng.Interior.Color = System.Drawing.Color.Yellow
                CUR_ROW += 1
                For Each YR As String In YEARS
                    Dim YFilter As String = $"CUST_CODE = '{CUST}' AND YR = '{YR}'"
                    Dim rows As DataRow() = dst.Tables.Item("SOTCMSM1").Select(YFilter, "RPT_ORDER")
                    If rows.Length > 0 Then
                        If Not IsNothing(XWSS) Then
                            rng = XWSS.Range(GetRange(CUR_ROW, CUR_COL, CUR_ROW, MAX_COL))
                            rng.FormulaR1C1 = $"Year: {YR}"
                            rng.MergeCells = True
                            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            rng.Font.Bold = True
                            rng.Interior.Color = System.Drawing.Color.LightGreen
                            CUR_ROW += 1
                        End If
                        rng = XWS.Range(GetRange(CUR_ROW, CUR_COL, CUR_ROW, MAX_COL))
                        rng.FormulaR1C1 = $"Year: {YR}"
                        rng.MergeCells = True
                        rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        rng.Font.Bold = True
                        rng.Interior.Color = System.Drawing.Color.LightGreen
                        CUR_ROW += 1
                        MakeYrHeader(CUR_ROW, MAX_COL, XWS, XWSS)
                        CUR_ROW += 1
                        Dim RFilter As String = $"CUST_CODE = '{CUST}' AND YR = '{YR}' AND DATA_TYPE = 'S'"
                        Dim dRow As DataRow = dst.Tables.Item("SOTCMSM1").Select(RFilter, "RPT_ORDER").FirstOrDefault
                        MakeDataRow("1", CUR_ROW, dRow, XWS, XWSS) 'Sales
                        CUR_ROW += 1
                        RFilter = $"CUST_CODE = '{CUST}' AND YR = '{YR}' AND DATA_TYPE = 'C'"
                        dRow = dst.Tables.Item("SOTCMSM1").Select(RFilter, "RPT_ORDER").FirstOrDefault
                        MakeDataRow("2", CUR_ROW, dRow, XWS, XWSS) 'Cost
                        CUR_ROW += 1
                        MakeDataRow("3", CUR_ROW, Nothing, XWS, XWSS) 'Gross Profit
                        CUR_ROW += 1
                        MakeDataRow("4", CUR_ROW, Nothing, XWS, XWSS) 'Gross Profit Pct
                        CUR_ROW += 1
                        RFilter = $"CUST_CODE = '{CUST}' AND YR = '{YR}' AND DATA_TYPE = 'R'"
                        dRow = dst.Tables.Item("SOTCMSM1").Select(RFilter, "RPT_ORDER").FirstOrDefault
                        MakeDataRow("5", CUR_ROW, dRow, XWS, XWSS) 'Returns
                        CUR_ROW += 1
                        RFilter = $"CUST_CODE = '{CUST}' AND YR = '{YR}' AND DATA_TYPE = 'D'"
                        dRow = dst.Tables.Item("SOTCMSM1").Select(RFilter, "RPT_ORDER").FirstOrDefault
                        MakeDataRow("6", CUR_ROW, dRow, XWS, XWSS) 'Deductions
                        CUR_ROW += 1
                        MakeDataRow("7", CUR_ROW, Nothing, XWS, XWSS) 'Total Net Sales
                        CUR_ROW += 1
                        MakeDataRow("8", CUR_ROW, Nothing, XWS, XWSS) 'Net Margin $
                        CUR_ROW += 1
                        MakeDataRow("9", CUR_ROW, Nothing, XWS, XWSS) 'Net Margin %
                        'For Each row As DataRow In rows
                        '    row.Item("COLOUM_NAME") = "XXXXX"
                        'Next
                    End If
                    CUR_ROW += 2
                Next
            End If

        Next
        excel.Visible = True
        'XWB.Close()
        XWB = Nothing
        excel = Nothing
    End Sub

    Private Sub FormatColumns(ByRef XWS As Excel.Worksheet)
        XWS.Range("A1").ColumnWidth = 25
        XWS.Range("A2:R1").ColumnWidth = 15
        XWS.Range("A2:R1").HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
    End Sub

    Private Sub MakeDataRow(ByVal DATA_TYPE As String, ByVal CUR_ROW As Int64, ByRef dRow As DataRow, ByRef XWS As Excel.Worksheet, ByRef XWSS As Excel.Worksheet)
        Dim hasSummary As Boolean = Not IsNothing(XWSS)
        Dim CNAME As String = ""
        Dim NFormatStr As String = "###,###,##0"
        Select Case DATA_TYPE
            Case "1"
                CNAME = "Sales"
            Case "2"
                CNAME = "COG"
            Case "3"
                CNAME = "Gross Profit"
            Case "4"
                CNAME = "GP%"
                NFormatStr = "00.0%"
            Case "5"
                CNAME = "Returns/Credits"
            Case "6"
                CNAME = "Deductions"
            Case "7"
                CNAME = "Total Net Sales"
            Case "8"
                CNAME = "Net Margin $"
            Case "9"
                CNAME = "Net Margin %"
                NFormatStr = "00.0%"
        End Select
        XWS.Range(GetRange(CUR_ROW, 2, CUR_ROW, 17)).NumberFormat = NFormatStr
        If hasSummary Then
            XWSS.Range(GetRange(CUR_ROW, 2, CUR_ROW, 17)).NumberFormat = NFormatStr
        End If

        If IsNothing(dRow) Then
            Select Case DATA_TYPE
                Case "3"
                    XWS.Cells(CUR_ROW, 1).VALUE = CNAME
                    For i As Int64 = 2 To 13
                        XWS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 2, i)}-{GetCell(CUR_ROW - 1, i)})"
                    Next
                    XWS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    If hasSummary Then
                        XWSS.Cells(CUR_ROW, 1).VALUE = CNAME
                        For i As Int64 = 2 To 13
                            XWSS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 2, i)}-{GetCell(CUR_ROW - 1, i)})"
                        Next
                        XWSS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    End If
                Case "4"
                    XWS.Cells(CUR_ROW, 1).VALUE = CNAME
                    For i As Int64 = 2 To 14
                        If IsNumeric(XWS.Cells(CUR_ROW - 3, i).value) Then
                            If Val(XWS.Cells(CUR_ROW - 3, i).value) <> 0 Then
                                XWS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 1, i)}/{GetCell(CUR_ROW - 3, i)})"
                            End If
                        End If
                    Next
                    If hasSummary Then
                        XWSS.Cells(CUR_ROW, 1).VALUE = CNAME
                        For i As Int64 = 2 To 14
                            If IsNumeric(XWSS.Cells(CUR_ROW - 3, i).value) Then
                                If Val(XWSS.Cells(CUR_ROW - 3, i).value) <> 0 Then
                                    XWSS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 1, i)}/{GetCell(CUR_ROW - 3, i)})"
                                End If
                            End If
                        Next
                    End If
                Case "7"
                    XWS.Cells(CUR_ROW, 1).VALUE = CNAME
                    For i As Int64 = 2 To 13
                        XWS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 6, i)}-{GetCell(CUR_ROW - 2, i)}-{GetCell(CUR_ROW - 1, i)})"
                    Next
                    XWS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    'Dim rng2 As Excel.Range = XWS.Range(GetRange(CUR_ROW, 1, CUR_ROW, 18))
                    If hasSummary Then
                        XWSS.Cells(CUR_ROW, 1).VALUE = CNAME
                        For i As Int64 = 2 To 13
                            XWSS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 6, i)}-{GetCell(CUR_ROW - 2, i)}-{GetCell(CUR_ROW - 1, i)})"
                        Next
                        XWSS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    End If
                Case "8"
                    XWS.Cells(CUR_ROW, 1).VALUE = CNAME
                    For i As Int64 = 2 To 13
                        XWS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 1, i)}-{GetCell(CUR_ROW - 6, i)})"
                    Next
                    XWS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    'Dim rng2 As Excel.Range = XWS.Range(GetRange(CUR_ROW, 1, CUR_ROW, 18))
                    If hasSummary Then
                        XWSS.Cells(CUR_ROW, 1).VALUE = CNAME
                        For i As Int64 = 2 To 13
                            XWSS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 1, i)}-{GetCell(CUR_ROW - 6, i)})"
                        Next
                        XWSS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    End If
                Case "9"
                    XWS.Cells(CUR_ROW, 1).VALUE = CNAME
                    For i As Int64 = 2 To 14
                        If IsNumeric(XWS.Cells(CUR_ROW - 2, i).value) Then
                            If XWS.Cells(CUR_ROW - 2, i).value <> 0 Then
                                XWS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 1, i)}/{GetCell(CUR_ROW - 2, i)})"
                            End If
                        End If
                    Next
                    'XWS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    'Dim rng2 As Excel.Range = XWS.Range(GetRange(CUR_ROW, 1, CUR_ROW, 18))
                    If hasSummary Then
                        XWSS.Cells(CUR_ROW, 1).VALUE = CNAME
                        For i As Int64 = 2 To 14
                            If IsNumeric(XWSS.Cells(CUR_ROW - 2, i).value) Then
                                If XWSS.Cells(CUR_ROW - 2, i).value <> 0 Then
                                    XWSS.Cells(CUR_ROW, i).Formula = $"=({GetCell(CUR_ROW - 1, i)}/{GetCell(CUR_ROW - 2, i)})"
                                End If
                            End If
                        Next
                    End If
                Case Else
                    XWS.Cells(CUR_ROW, 1).VALUE = CNAME
                    For i As Int64 = 2 To 13
                        XWS.Cells(CUR_ROW, i).VALUE = 0
                    Next
                    XWS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    If hasSummary Then
                        XWSS.Cells(CUR_ROW, 1).VALUE = CNAME
                        For i As Int64 = 2 To 13
                            Dim FLDNO As String = Format(i - 1, "00")
                            Dim CVAL As Int64 = 0
                            If IsNumeric(XWSS.Cells(CUR_ROW, i).VALUE) Then
                                CVAL = Val(XWSS.Cells(CUR_ROW, i).VALUE)
                            End If
                            XWSS.Cells(CUR_ROW, i).VALUE = CVAL
                        Next
                        XWSS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
                    End If
            End Select
        Else
            XWS.Cells(CUR_ROW, 1).VALUE = CNAME
            For i As Int64 = 2 To 13
                Dim FLDNO As String = Format(i - 1, "00")
                XWS.Cells(CUR_ROW, i).VALUE = getNumVal(dRow($"MO_{FLDNO}").ToString)
            Next
            XWS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
            If hasSummary Then
                XWSS.Cells(CUR_ROW, 1).VALUE = CNAME
                For i As Int64 = 2 To 13
                    Dim FLDNO As String = Format(i - 1, "00")
                    'If CNAME = "Sales" Then
                    '    If FLDNO = "01" And i = 2 Then
                    '        If getNumVal(dRow($"MO_{FLDNO}").ToString) = 7477530.4 Then Stop
                    '        If getNumVal(dRow($"MO_{FLDNO}").ToString) = 36136.8 Then Stop
                    '        If getNumVal(dRow($"MO_{FLDNO}").ToString) = 0 Then Stop
                    '    End If
                    'End If

                    Dim CVAL As Int64 = 0
                    If IsNumeric(XWSS.Cells(CUR_ROW, i).VALUE) Then
                        CVAL = Val(XWSS.Cells(CUR_ROW, i).VALUE)
                    End If
                    XWSS.Cells(CUR_ROW, i).VALUE = CVAL + getNumVal(dRow($"MO_{FLDNO}").ToString)
                Next
                XWSS.Cells(CUR_ROW, 14).Formula = $"=SUM({GetRange(CUR_ROW, 2, CUR_ROW, 13)})"
            End If
            'Dim rng As Excel.Range = XWS.Range(GetRange(CUR_ROW, 1, CUR_ROW, MAX_COL))
            'rng.Font.Bold = True
            'rng.Interior.Color = System.Drawing.Color.LightGray
        End If
    End Sub

    Private Function getNumVal(ByVal inString As String) As Int64
        Dim RetVal As Int64 = 0
        If IsNumeric(inString) Then
            RetVal = Val(inString)
        End If
        Return RetVal
        Throw New NotImplementedException()
    End Function

    Private Sub MakeYrHeader(ByVal CUR_ROW As Long, ByVal MAX_COL As Long, ByRef XWS As Excel.Worksheet, ByRef XWSS As Excel.Worksheet)
        If Not IsNothing(XWSS) Then
            XWSS.Cells(CUR_ROW, 1).VALUE = "Category"
            XWSS.Cells(CUR_ROW, 2).VALUE = "Jan"
            XWSS.Cells(CUR_ROW, 3).VALUE = "Feb"
            XWSS.Cells(CUR_ROW, 4).VALUE = "Mar"
            XWSS.Cells(CUR_ROW, 5).VALUE = "Apr"
            XWSS.Cells(CUR_ROW, 6).VALUE = "May"
            XWSS.Cells(CUR_ROW, 7).VALUE = "Jun"
            XWSS.Cells(CUR_ROW, 8).VALUE = "Jul"
            XWSS.Cells(CUR_ROW, 9).VALUE = "Aug"
            XWSS.Cells(CUR_ROW, 10).VALUE = "Sep"
            XWSS.Cells(CUR_ROW, 11).VALUE = "Oct"
            XWSS.Cells(CUR_ROW, 12).VALUE = "Nov"
            XWSS.Cells(CUR_ROW, 13).VALUE = "Dec"
            XWSS.Cells(CUR_ROW, 14).VALUE = "Total"
            XWSS.Cells(CUR_ROW, 15).VALUE = "--"
            XWSS.Cells(CUR_ROW, 16).VALUE = "--"
            XWSS.Cells(CUR_ROW, 17).VALUE = "--"
            XWSS.Cells(CUR_ROW, 18).VALUE = "--"
            Dim rngs As Excel.Range = XWSS.Range(GetRange(CUR_ROW, 1, CUR_ROW, MAX_COL))
            rngs.Font.Bold = True
            rngs.Interior.Color = System.Drawing.Color.LightGray
            rngs = XWSS.Range(GetRange(CUR_ROW, 2, CUR_ROW, MAX_COL))
            rngs.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
        End If
        XWS.Cells(CUR_ROW, 1).VALUE = "Category"
        XWS.Cells(CUR_ROW, 2).VALUE = "Jan"
        XWS.Cells(CUR_ROW, 3).VALUE = "Feb"
        XWS.Cells(CUR_ROW, 4).VALUE = "Mar"
        XWS.Cells(CUR_ROW, 5).VALUE = "Apr"
        XWS.Cells(CUR_ROW, 6).VALUE = "May"
        XWS.Cells(CUR_ROW, 7).VALUE = "Jun"
        XWS.Cells(CUR_ROW, 8).VALUE = "Jul"
        XWS.Cells(CUR_ROW, 9).VALUE = "Aug"
        XWS.Cells(CUR_ROW, 10).VALUE = "Sep"
        XWS.Cells(CUR_ROW, 11).VALUE = "Oct"
        XWS.Cells(CUR_ROW, 12).VALUE = "Nov"
        XWS.Cells(CUR_ROW, 13).VALUE = "Dec"
        XWS.Cells(CUR_ROW, 14).VALUE = "Total"
        XWS.Cells(CUR_ROW, 15).VALUE = "--"
        XWS.Cells(CUR_ROW, 16).VALUE = "--"
        XWS.Cells(CUR_ROW, 17).VALUE = "--"
        XWS.Cells(CUR_ROW, 18).VALUE = "--"
        Dim rng As Excel.Range = XWS.Range(GetRange(CUR_ROW, 1, CUR_ROW, MAX_COL))
        rng.Font.Bold = True
        rng.Interior.Color = System.Drawing.Color.LightGray
        rng = XWS.Range(GetRange(CUR_ROW, 2, CUR_ROW, MAX_COL))
        rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
    End Sub

    Private Function GetRange(ByVal Row1 As Int64, ByVal Col1 As Int64, ByVal Row2 As Int64, ByVal Col2 As Int64) As String
        Dim RetVal As String = ""
        Dim Base As Int64 = Asc("A") - 1
        RetVal = $"{Chr(Base + Col1).ToString}{Row1}:{Chr(Base + Col2).ToString}{Row2}"
        Return RetVal
    End Function

    Private Function GetCell(ByVal R As Int64, ByVal C As Int64) As String
        Dim RetVal As String = "A1"
        Dim Base As Int64 = Asc("A") - 1
        RetVal = $"{Chr(Base + C).ToString}{R}"
        Return RetVal
    End Function

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If numYears.Value < 1 Or numYears.Value > 4 Then
                EMsg = EMsg & vbCrLf & "# Years Must Be Between 1 - 4"
            Else
                YEARS.Clear()
                Dim SELYEAR As String = cboBaseYear.Text
                For i As Int64 = Val(SELYEAR) To (Val(SELYEAR) - numYears.Value) Step -1
                    YEARS.Add($"{i.ToString}")
                Next
            End If

            Dim rows() As DataRow = tblASTDSQLA.Select("COLUMN_NAME = 'CUST_CODE'")
            If rows.Length = 1 Then
                Dim CUST_CODES As String() = rows(0).Item("CODE_VALUES").ToString.Split(",")
                Array.Reverse(CUST_CODES)
                If CUST_CODES.Length > 0 Then
                    CUST_LIST.Clear()
                    For Each CC As String In CUST_CODES
                        CUST_LIST.Add(CC)
                    Next
                Else
                    EMsg = EMsg & vbCrLf & "You Must Select At Least One Customer."
                End If
            Else
                EMsg = EMsg & vbCrLf & "Problem With Customer Selection."
            End If
        End If
    End Sub

    Private Sub cboBaseYear_ValueChanged(sender As Object, e As EventArgs) Handles cboBaseYear.ValueChanged

    End Sub
End Class