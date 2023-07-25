Imports Microsoft.Office.Interop
Imports System.Drawing
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Http.Formatting
Imports Newtonsoft.Json
Imports Newtonsoft
Imports System.Text

Public Class SOFORDRO
    Dim rowSOTORDRB As DataRow
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String
    Dim ORDR_BATCH_NO As String
    Dim SREP_CODE As String = ""
    Dim Discounts As List(Of DISCOUNTS)
    Dim SelectedTabHit As Boolean = False
    Dim PINNED_CUST_PRICE_TIER_PVC As String = ""
    Dim Remote As New REMOTE(Me)
    Private mExcelProcesses() As Process
    Dim EXCELDir As String = GetExcelFolder()
    Dim IMAGES_FOLDER As String = GetImageFolder()
    Dim SCD As Integer = 7 'Start Cell for Details on order
    Dim RowCount As Integer = 0
    Dim EndMark As Integer = 0
    Dim OrderHasFEWhse As Boolean = False
    Dim InquiryOnly As Boolean = False
    Dim CustOverride As Boolean = False
    Dim FEOverride As Boolean = False
    Dim LoadFinished As Boolean = False
    Dim EditOrderLineMode As Boolean = False
    Dim EditOrderLineWarning As Boolean = False
    Dim ExcelSort As String = ""
    Dim PictureSort As String = ""
    Dim FactorySort As String = "FACTORY_CODE, STYLE_CODE, COLOR_CODE"
    Dim ReBatch As New List(Of String)
    Dim FEFDLEVEL As String = ""
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        SREP_CODE = Remote.SREP_CODE

        Get_PARM("SOTPARM1")
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}

        With dst
            Create_TDA(.Tables.Add, "SOTORDRB", "*", 1)

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)
            .Tables("SOTORDR1").Columns.Add("FKEY", GetType(System.String))
            .Tables("SOTORDR1").Columns.Add("TCUFT", GetType(System.String))
            .Tables("SOTORDR1").Columns.Add("TORDR", GetType(System.String))
            .Tables("SOTORDR1").Columns.Add("CUST_PRICE_TIER_PVC", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            With .Tables("SOTORDR2").Columns
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("CASE_CUBE", GetType(System.Double))
                .Add("TCUFT", GetType(System.Double))
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_ALLO", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            End With

            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)

            Create_TDA(.Tables.Add, "SOTORDRL", "*", 1)

            Create_TDA(.Tables.Add, "ICTSTYL1", "*", 1)

            ASCMAIN1.sql = "SELECT * FROM ARTCUST2 WHERE NVL(CUST_ADDR_STATUS,'A') = 'A' AND CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, True, "V", 3)
            'Create_TDA(.Tables.Add, "ARTCUST2", "*", 1)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            Create_TDA(.Tables.Add, "ICTSTAT2", "*", 3)

            SQLs.Length = 0
            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM")
            SQLs.AppendLine("  (")
            SQLs.AppendLine("   SELECT C1.STYLE_CODE, C1.COLOR_CODE,")
            SQLs.AppendLine("   9999 AS ORDR_QTY,")
            SQLs.AppendLine("   C2.COLOR_DESC AS COLOR_CODE_LONG,")
            SQLs.AppendLine("   C1.STYLE_COLOR_STATUS,")
            SQLs.AppendLine("   CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) < 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE")
            SQLs.AppendLine("     WHEN 'MS'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END)")
            SQLs.AppendLine("   END AS MSOH,")
            SQLs.AppendLine("   CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) <= 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("     CASE WHEN")
            SQLs.AppendLine("       SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'MS'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) < 0")
            SQLs.AppendLine("     THEN")
            SQLs.AppendLine("       0")
            SQLs.AppendLine("     ELSE")
            SQLs.AppendLine("     SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'MS'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) END")
            SQLs.AppendLine("   END AS MSFT,")
            SQLs.AppendLine(" CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) < 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE")
            SQLs.AppendLine("     WHEN 'SW'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END)")
            SQLs.AppendLine("   END AS SWOH,")
            SQLs.AppendLine("   CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) <= 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("     CASE WHEN")
            SQLs.AppendLine("       SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'SW'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) < 0")
            SQLs.AppendLine("     THEN")
            SQLs.AppendLine("       0")
            SQLs.AppendLine("     ELSE")
            SQLs.AppendLine("     SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'SW'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) END")
            SQLs.AppendLine("   END AS SWFT,")
            SQLs.AppendLine("   C1.THEME_CODE")
            SQLs.AppendLine("   FROM ICTSTYC1 C1")
            SQLs.AppendLine("   LEFT JOIN ICTSTAT2 S2")
            SQLs.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
            SQLs.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
            SQLs.AppendLine("   INNER JOIN ICTCOLR1 C2")
            SQLs.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
            SQLs.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS, C1.THEME_CODE")
            SQLs.AppendLine("  )")
            'SQLs.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') or (MSOH <> 0) or (MSFT <> 0) or (SWOH <> 0)  or (SWFT <> 0))")
            SQLs.AppendLine("  WHERE STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQLs.ToString
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False, "V", 2)
            'Fill_Records("ICTSTYC1", "", , ASCMAIN1.sql)
            .Tables("ICTSTYC1").Columns.Add("RSV", GetType(System.String))

            Dim SQLW As String = ""
            If Remote.SQLWhere.Length > 0 Then
                SQLW = " Where " & Remote.SQLWhere
                SQLW += String.Format(" AND ORDR_DATE >= '{0}'", Format(Now.AddMonths(-1), "dd-MMM-yy"))
            Else
                SQLW = String.Format(" Where ORDR_DATE >= '{0}'", Format(Now.AddMonths(-1), "dd-MMM-yy"))
            End If
            SQLW += String.Format(" OR ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1_L WHERE SREP_CODE = '{0}')", Remote.SREP_CODE)
            'Added so Sales Reps can see all orders for their customers regardless of who wrote it. - w.r. - 8/11/14
            SQLW += String.Format(" OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE = '{0}')", Remote.SREP_CODE)
            ASCMAIN1.sql = "SELECT SOTORDR1.*,  TO_CHAR(ORDR_DATE, 'YYYY') AS YEAR FROM SOTORDR1 " & SQLW
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False)
            .Tables("SOTORDRX").Columns.Add("TORDR", GetType(System.String))
            'ASCMAIN1.sql = ShowJoinedSreps(ASCMAIN1.sql)
            'Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDX1", "**", 0, False, "V", 1)
            Create_TDA(.Tables.Add, "SOTORDP1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR2 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDX2", "**", 0, False, "V", 2)
            Create_TDA(.Tables.Add, "SOTORDP2", "**", 0, False, "V", 2)
            .Tables("SOTORDX2").Columns.Add("FACTORY_CODE", GetType(System.String))
            .Tables("SOTORDP2").Columns.Add("COLOR_DESC", GetType(System.String))
            .Tables("SOTORDP2").Columns.Add("UPC_CODE", GetType(System.String))
            .Tables("SOTORDP2").Columns.Add("FACTORY_CODE", GetType(System.String))


            ASCMAIN1.sql = "SELECT * FROM SOTORDR5 where ORDR_NO = :PARM1 AND CUST_ADDR_TYPE = 'ST'"
            Create_TDA(.Tables.Add, "SOTORDX5", "**", 0, False, "V", 2)
            Create_TDA(.Tables.Add, "SOTORDP5", "**", 0, False, "V", 2)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT WHSE_CODE, STYLE_CODE, COLOR_CODE,")
            'Changed from QTY_ATS_CUM to QTY_ATS per WZ - 7/10/14 
            'After meeting with Rich, to review the screen this turned out to be a lie.  Changed back 10/23/14.
            SQLs.AppendLine("STATUS_DATE, QTY_ATS_CUM as QTY_ATS_CUM")
            SQLs.AppendLine("FROM ICTSTDQ1")
            SQLs.AppendLine("WHERE STYLE_CODE  = :PARM1")
            SQLs.AppendLine("AND COLOR_CODE  = :PARM2")
            SQLs.AppendLine("AND WHSE_CODE  = :PARM3")
            ASCMAIN1.sql = SQLs.ToString
            Create_TDA(.Tables.Add, "ICTSTDQ1", "**", 0, False, "VVV")
            'Fill_Records("ICTSTDQ1", "", , ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "HANGTAG1", "**", 1, False)
            .Tables("HANGTAG1").Columns.Add("DATEPRINTED", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("BOXQTY", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("CARTQTY", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("COLORS", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("COLORS3", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price1_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price1_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("Price2_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price2_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("Price3_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price3_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("Price4_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price4_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("COLORSDESC", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("VEND_SUPPLIER_ID", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("PORT_CODE_ORIG", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "ICTSTYLD", "**", 1, False)
            Fill_Records("ICTSTYLD", "", , ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL3 WHERE STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False, "V")

            ASCMAIN1.sql = "Select ARTCUSTD.* " _
                & " from ARTCUSTD " _
                & " where ARTCUSTD.CUST_CODE = :PARM1" _
                & " and NVL(CONTACT_NOTE,'NULL') <> 'DELETED'"
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V", 2)

            'CHANGES FOR SHIPTO QUESTION - 
            'ASCMAIN1.sql = "Select ARTCUSTQ.* " _
            '& " from ARTCUSTQ " _
            '& " where ARTCUSTQ.CUST_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "ARTCUSTQ", "**", 0, True, "V", 3)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM ARTCUSTQ")
            SQLs.AppendLine(" WHERE CUST_CODE = :PARM1")
            SQLs.AppendLine(" AND CUST_ADDR_CODE = :PARM2")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "ARTCUSTQ", "**", 0, True, "VV", 2)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM ARTCUST2")
            SQLs.AppendLine(" WHERE CUST_CODE = :PARM1")
            SQLs.AppendLine(" AND CUST_ADDR_CODE = :PARM2")
            SQLs.AppendLine("  AND CUST_ADDR_TYPE = 'MK'")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "ARTCUSX2", "**", 0, False, "VV", 3)
            .Tables("ARTCUSX2").Columns.Add("VERIFIED", GetType(System.String))
        End With

        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdSOTORDR1.DataSource = dst.Tables("SOTORDR1")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        grdICTSTYC1.DataSource = dst.Tables("ICTSTYC1")
        grdSHIP2.DataSource = dst.Tables("SOTORDR1")
        grdICTSTDQ1.DataSource = dst.Tables("ICTSTDQ1")
        FilterColors("NONE")
        FilterAlloc("NONE", "NONE", "NONE")

        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY", "ORDR_AMT", "TCUFT", "ORDR_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})

        Create_Summary(grdSOTORDRX, "TORDR", "Sum", "", "###,##0.00")

        Sort_grdColumns(grdSOTORDRX, "ORDR_DATE, ORDR_GROUP_NO, ORDR_NO".ToLower(), False)
        Sort_grdColumns(grdSOTORDR1, "ORDR_NO", False)

        grdSHIP2.DisplayLayout.Bands(0).Columns("TCUFT").Format = "###,##0.00"
        grdSHIP2.DisplayLayout.Bands(0).Columns("TORDR").Format = "###,##0.00"
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("TORDR").Format = "###,##0.00"
        'grdICTSTYC1.DisplayLayout.Bands(0).Columns("ORDR_QTY").MaskInput = "###,##0"
        'Call Create_Summary(grdSOTORDRX, "PYMT_BATCH_NO", "Count")

        'With grdARTPYMT2.DisplayLayout.Bands(0)
        '    .Columns("PYMT_BATCH_LNO").CellAppearance.BackColor = Drawing.Color.Beige
        '    .Columns("CUST_CODE").CellAppearance.BackColor = Drawing.Color.Beige
        '    .Columns("CUST_NAME").CellAppearance.BackColor = Drawing.Color.Beige
        '    .Columns("NON_AR").CellAppearance.BackColor = Drawing.Color.Beige
        'End With
        'grdARTPYMT2.DisplayLayout.Bands(0).Columns("NON_AR").Tag = "N"
        Setup_SOTORDR2()

        'Fill_Records("SOTORDR1")

        ASCMAIN1.Add_Value_List(grdSOTORDRX, "ORDR_STATUS", , New String() {":", "L:Laptop", "Q:Quote", "C:Cancelled"})
        ASCMAIN1.Add_Value_List(grdSOTORDR1, "ORDR_STATUS", , New String() {":", "L:Laptop", "Q:Quote", "C:Cancelled"})
        grdSOTORDR1.DisplayLayout.Bands(0).Columns("ORDR_STATUS").Style = UltraWinGrid.ColumnStyle.EditButton
        grdSOTORDR1.DisplayLayout.Bands(0).Columns("ORDR_STATUS").CellButtonAppearance.Image = Nothing
        'grdSOTORDR1.DisplayLayout.Bands(0).Columns("ORDR_STATUS").CellButtonAppearance
        'Dim ORDRLIST As String() = New String() {":", "L:Laptop", "Q:Quote"}
        'grdSOTORDR1.DisplayLayout.ValueLists.Add("ORDR_STATUS")

        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_QTY"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.LightCyan
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT_OPEN", "ORDR_AMT_ALLO", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "ORDR_AMT"}
                With .Columns(COLUMN_NAME)
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
        End With

        With grdICTSTYC1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdICTSTYC1.DisplayLayout.Bands(0).Columns.Count - 1
            grdICTSTYC1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        For Each COLNAME As String In New String() {"ORDR_QTY"}
            grdICTSTYC1.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        Next
        For Each COLNAME As String In New String() {"ORDR_QTY"}
            grdICTSTYC1.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Next

        tab.Visible = False
        grdSOTORDRX.Parent = tab.Parent
        LoadFinished = True

        SplitContainer1.SplitterDistance = UltraTabPageControl1.Height - btnAddShipTo.Height

    End Sub

    Private Function VERIFY_SREP_CODE() As Boolean
        Dim retval As Boolean = False
        If Remote.IsUserSuper Then
            retval = True
        Else
            If Remote.SREP_CODE <> "" Then
                retval = True
            End If
        End If
        Return retval
    End Function

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Call Validate_Code("CUST_CODE")
                If Absx1.txtFor("CUST_CODE").Text.Length > 0 Then
                    If Not CheckForSalesHold() Then
                        MsgBox("Selected Customer Is On Sales Hold.", MsgBoxStyle.Information, "Please Note")
                    End If
                End If
                'If Absx1.txtFor("ORDR_CUST_PO").Text.Length = 0 Then
                '    EMsg &= vbCr & "You Must First Specify a Customer PO"
                'End If
                'If Not VERIFY_SREP_CODE() Then
                '    EMsg &= vbCr & "Your Account Has Not Been Set-Up As A Valid Sales Rep"
                'Else
                If Not CheckCustSrep() Then
                    EMsg &= vbCr & "Selected Customer has invalid Sales Rep Assigned"
                End If
                'End If
            Case "Edit"
                'Call Validate_Code("CUST_CODE")
                If Absx1.txtFor("ORDR_BATCH_NO").Text.Length = 0 Then
                    If grdSOTORDRX.Selected.Rows.Count <= 0 Then
                        EMsg &= vbCr & "You Must First Specify or Select an Order Group"
                    Else
                        If grdSOTORDRX.Selected.Rows.Count <> 1 Then
                            EMsg &= vbCr & "You May Only Select One Order Group At A Time To Edit"
                        Else
                            Absx1.txtFor("ORDR_BATCH_NO").Text = grdSOTORDRX.Selected.Rows(0).Cells("ORDR_GROUP_NO").Text
                        End If
                    End If
                End If
                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTORDRB", Absx1.txtFor("ORDR_BATCH_NO").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                Dim HasValidOrder As Boolean = False
                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                    If rowSOTORDR1.Item("ORDR_STATUS") = "L" Or rowSOTORDR1.Item("ORDR_STATUS") = "Q" Then
                        HasValidOrder = True
                    End If
                Next
                If HasValidOrder Then
                    Dim iResult As MsgBoxResult = MsgBox("Cancelling Will Lose Any Changes You May Have Made." & vbCrLf & "Are You Sure You Want To Cancel?", MsgBoxStyle.YesNo, "Cancel Confirmation")
                    If iResult = MsgBoxResult.No Then
                        EMsg &= vbCr & "Cancel Option Aborted"
                    End If
                End If
            Case "Update"
                Dim DUP_STYLE_LIST As System.Text.StringBuilder = CHECK_STYLE_COLOR_DUPS()
                If DUP_STYLE_LIST.Length > 0 Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Duplicate Style/Colors Found"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("The Following Duplicate Style / Colors")
                    iMSG.AppendLine("Were Found On This Order.  You Should")
                    iMSG.AppendLine("Fix Them Before Transmitting The Order.")
                    iMSG.AppendLine("")
                    iMSG.AppendLine(DUP_STYLE_LIST.ToString)
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Update Anyway?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.No Then
                        EMsg &= vbCr & "Fix Duplicate Style/Colors."
                    End If
                End If
                If shipToVerify(False) = False Then
                    EMsg &= vbCr & "You Must Verify Ship Tos."
                End If
                Dim CCOK As Boolean = AllOrdersAreCC(True)
                If Not CCOK Then
                    'Dim Msg As String = "Credit Card Orders Found With No Credit Card Information Recorded."
                    'Msg = Msg & vbCrLf & "Are You Sure You Want To Proceed?"
                    'Dim iresult As MsgBoxResult = MsgBox(Msg, vbYesNo, "Credit Card Order")
                    'If iresult <> vbYes Then
                    EMsg &= vbCr & "Please Enter Credit Card Info or Change Terms"
                    'End If
                Else
                    Dim CCPA As Boolean = CCPAOnOrders()
                    If CCPA Then
                        Dim Msg As String = "Orders Found With Credit Card Information And No Credit Card Terms."
                        Msg = Msg & vbCrLf & "Would You Like To Remove Credit Card Information?"
                        Dim iresult As MsgBoxResult = MsgBox(Msg, vbYesNo, "Credit Card Order")
                        If iresult = vbYes Then
                            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                                rowSOTORDR1.Item("CCPA_NO") = ""
                            Next
                        Else
                            EMsg &= vbCr & "Please Enter Credit Card Info or Change Terms"
                        End If
                    End If
                End If


                'For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                '    Dim TERM_CODE As String = rowSOTORDR1.Item("TERM_CODE") & ""
                '    Dim CC_TRANS_ID As String = rowSOTORDR1.Item("CC_TRANS_ID") & ""
                '    Dim CCREquired As Boolean = False
                '    Select Case TERM_CODE
                '        Case Is = "AMEX", "DISC", "MC", "VISA"
                '            If CC_TRANS_ID = "" Then
                '                CCREquired = True
                '            End If
                '    End Select
                '    If CCREquired Then
                '        EMsg &= vbCr & "Credit Card Orders Found With No Credit Card Information Recorded"
                '    End If
                'Next
            Case "Print"
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must Select An Order From The Orders List To Print"
                End If
            Case "Print Preview"
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must Select An Order From The Orders List To Print"
                End If
            Case "Delete"
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must Select An Order From The Orders List To Delete"
                Else
                    If grdSOTORDRX.Selected.Rows.Count > 1 Then
                        EMsg &= vbCr & "You Can Only Select One Order At A Time To Delete"
                    Else
                        Dim ORDR_STATUS As String = grdSOTORDRX.Selected.Rows(0).Cells.Item("ORDR_STATUS").Value
                        Dim iMsg As String = ""
                        If ORDR_STATUS = "L" Or ORDR_STATUS = "Q" Then
                            iMsg = "Deleting This Order Or Quote Will PERMANENTLY"
                            iMsg = iMsg & vbCrLf & "Remove It From The System!!"
                            iMsg = iMsg & vbCrLf & "Are You Sure This Is What You Want?"
                            Dim iResponse As MsgBoxResult = MsgBox(iMsg, MsgBoxStyle.YesNo, "Pay Attention!")
                            If iResponse <> MsgBoxResult.Yes Then
                                EMsg &= vbCr & "Delete Aborted"
                            End If
                        Else
                            EMsg &= vbCr & "You May Only Delete Orders Or Quotes That Have Not Been Transmitted."
                        End If
                    End If
                End If
            Case "Hang Tag"
                PrintHangTag()
            Case "Re-Price Order"
                If grdSOTORDR1.Selected.Rows.Count <> 1 Then
                    EMsg &= vbCr & "You Must Select One And Only One Order From The Order Grid To Reprice"
                Else
                    Dim ORDR_STATUS As String = grdSOTORDR1.Selected.Rows(0).Cells.Item("ORDR_STATUS").Value
                    Dim iMsg As String = ""
                    If ORDR_STATUS <> "L" And ORDR_STATUS <> "Q" Then
                        EMsg &= vbCr & "You May Only Re-Price Orders & Quote That Have Not Been Transmitted"
                    Else
                        If grdSOTORDR1.Selected.Rows.Count <> 1 Then
                            EMsg &= vbCr & "You Must Select One And Only One Order From The Order Grid To Reprice"
                        End If
                    End If
                End If
            Case "Record Credit Card"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Credit Cards"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine("This Will Associate The Credit Card With All Orders")
                iMSG.AppendLine("In This Group. Any Orders Added To This Group")
                iMSG.AppendLine("After This Will Not Be Accociated.")
                iMSG.AppendLine("Do You Still Want To Continue?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.No Then
                    EMsg &= vbCr & "You May Add The Credit Card Info Later Too."
                End If
            Case "Update Addresses"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Update Addresses"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine("This Will Update This Sales Order With")
                iMSG.AppendLine("Any Address Changes To The Customer")
                iMSG.AppendLine("Masterfile.")
                iMSG.AppendLine("Do You Still Want To Continue?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.No Then
                    EMsg &= vbCr & "You May Update Any Time Before Transferring."
                End If
            Case "Tablet"
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must Select An Order From The Orders List To Tabletize"
                Else
                    If grdSOTORDRX.Selected.Rows.Count > 1 Then
                        EMsg &= vbCr & "You Can Only Select One Order At A Time To Tabletize"
                    Else
                        Dim ORDR_STATUS As String = grdSOTORDRX.Selected.Rows(0).Cells.Item("ORDR_STATUS").Value
                        Dim iMsg As String = ""
                        If ORDR_STATUS = "L" Then
                            Dim ORDR_GROUP_NO As String = grdSOTORDRX.Selected.Rows(0).Cells.Item("ORDR_GROUP_NO").Text
                            Dim FILTER As String = String.Format("ORDR_GROUP_NO = '{0}'", ORDR_GROUP_NO)
                            Dim OrderCnt As Int64 = 0
                            Dim AllOrdersL As Boolean = True
                            For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select(FILTER)
                                OrderCnt += 1
                                If rowSOTORDRX.Item("ORDR_STATUS").ToString <> "L" Then
                                    AllOrdersL = False
                                End If
                            Next
                            If AllOrdersL Then
                                iMsg = String.Format("You Are Tabletizing {0} Order(s) And Will No Longer Be Able to", OrderCnt)
                                iMsg = iMsg & vbCrLf & "See or Edit them On This Screen."
                                iMsg = iMsg & vbCrLf & "Are You Sure This Is What You Want?"
                                Dim iResponse As MsgBoxResult = MsgBox(iMsg, MsgBoxStyle.YesNo, "!!!Pay Attention!!!")
                                If iResponse <> MsgBoxResult.Yes Then
                                    EMsg &= vbCr & "Tabletize Aborted"
                                End If
                            Else
                                EMsg &= vbCr & "All Orders Selected Are Not In The Same State."
                            End If

                        Else
                            EMsg &= vbCr & "You May Only Tabletize Orders That Have Not Been Transmitted."
                        End If
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Private Function CHECK_STYLE_COLOR_DUPS() As System.Text.StringBuilder
        Dim RET_VAL As New System.Text.StringBuilder With {.Length = 0}
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            Dim STYLE_COLORS As New List(Of String)
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO").ToString & String.Empty
            Dim rowFilter As String = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(rowFilter)
                Dim STYLE_COLOR As String = rowSOTORDR2.Item("STYLE_CODE").ToString & String.Empty & "-" & rowSOTORDR2.Item("COLOR_CODE").ToString
                If STYLE_COLORS.Contains(STYLE_COLOR) Then
                    RET_VAL.AppendLine(String.Format("Order:{0}: {1}", ORDR_NO, STYLE_COLOR))
                Else
                    STYLE_COLORS.Add(STYLE_COLOR)
                End If
            Next
        Next
        Return RET_VAL
    End Function

    Private Function shipToVerify(ByVal showForm As Boolean) As Boolean
        'NOTE: Very Similar Code to This is also in SOTCUST1 with the same name.
        '      If you are making changes here you should consider doing it There
        '      As well.
        Dim RetVal As Boolean = True
        dst.Tables.Item("ARTCUSTQ").Clear()
        dst.Tables.Item("ARTCUSX2").Clear()
        Dim ORDR_NO As String = ""
        Dim ORDRS As New List(Of String)
        Dim CUST_ADDR_CODEs As New List(Of String)
        For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select("", ORDR_NO)
            Dim CUST_ADDR_CODE As String = rowSOTORDR5.Item("CUST_ADDR_CODE").ToString & String.Empty
            If ORDR_NO.Length = 0 Then
                ORDR_NO = rowSOTORDR5.Item("ORDR_NO").ToString & String.Empty
            End If
            If Not CUST_ADDR_CODEs.Contains(CUST_ADDR_CODE) Then
                CUST_ADDR_CODEs.Add(CUST_ADDR_CODE)
            End If
            If Not ORDRS.Contains(rowSOTORDR5.Item("ORDR_NO").ToString & String.Empty) Then
                ORDRS.Add(rowSOTORDR5.Item("ORDR_NO").ToString & String.Empty)
            End If
        Next
        For Each CC As String In CUST_ADDR_CODEs
            Fill_Records("ARTCUSTQ", New String() {CUST_CODE, CC}, False)
            Fill_Records("ARTCUSX2", New String() {CUST_CODE, CC}, False)
        Next
        For Each rowARTCUSX2 As DataRow In dst.Tables("ARTCUSX2").Select()
            Dim CUST_ADDR_CODE As String = rowARTCUSX2.Item("CUST_ADDR_CODE").ToString & String.Empty
            Dim Filter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CUST_ADDR_CODE)
            If dst.Tables("ARTCUSTQ").Select(Filter).Count = 0 Then
                Dim newARTCUSTQ As DataRow = dst.Tables("ARTCUSTQ").NewRow
                newARTCUSTQ.Item("CUST_CODE") = CUST_CODE
                newARTCUSTQ.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                newARTCUSTQ.Item("LAST_DATE") = Null
                newARTCUSTQ.Item("LAST_OPER") = Null
                newARTCUSTQ.Item("LAST_ORDR_NO") = Null
                newARTCUSTQ.Item("RESIDENTIAL_ORDR") = "0"
                newARTCUSTQ.Item("INSIDE_REQ") = "0"
                newARTCUSTQ.Item("APPOINTMENT_REQUIRED") = "0"
                newARTCUSTQ.Item("GATE_LIFT_REQ") = "0"
                newARTCUSTQ.Item("LIMITED_ACCESS") = "0"
                newARTCUSTQ.Item("IRREGULAR_HOURS") = "0"
                newARTCUSTQ.Item("APPOINTMENT_REQUIRED") = "0"
                newARTCUSTQ.Item("BROKER") = "0"
                dst.Tables("ARTCUSTQ").Rows.Add(newARTCUSTQ)
            End If
        Next

        If showForm Then
            Dim frmSOFORDRS As New SOFORDRS(Me, Absx1.txtFor("CUST_CODE").Text, ORDR_NO)
            With frmSOFORDRS
                .ShowDialog()
            End With
        End If

        Call Update_Record_TDA("ARTCUSTQ")

        For Each rowARTCUSTQ As DataRow In dst.Tables("ARTCUSTQ").Select()
            Dim LAST_ORDR_NO As String = rowARTCUSTQ.Item("LAST_ORDR_NO").ToString & String.Empty
            Dim NOWDATE As Date = CDate(Now().ToShortDateString)
            Dim LAST_DATE As Date = CDate("01/01/1900")
            If IsDate(rowARTCUSTQ.Item("LAST_DATE").ToString & String.Empty) Then
                LAST_DATE = CDate(CDate((rowARTCUSTQ.Item("LAST_DATE").ToString & String.Empty)).ToShortDateString)
            End If
            If LAST_ORDR_NO = ORDR_NO Or LAST_DATE = NOWDATE Then
                Dim CUST_CODE As String = rowARTCUSTQ.Item("CUST_CODE").ToString & String.Empty
                Dim CUST_ADDR_CODE As String = rowARTCUSTQ.Item("CUST_ADDR_CODE").ToString & String.Empty
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("DELETE FROM ARTCUSTQ_L")
                SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                SQLS.AppendLine(String.Format("AND CUST_ADDR_CODE = '{0}'", CUST_ADDR_CODE))
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
                SQLS.Length = 0
                SQLS.AppendLine("INSERT INTO ARTCUSTQ_L")
                SQLS.AppendLine("SELECT * FROM ARTCUSTQ")
                SQLS.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                SQLS.AppendLine(String.Format("AND CUST_ADDR_CODE = '{0}'", CUST_ADDR_CODE))
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
            Else

                RetVal = False
            End If
        Next

        Return RetVal
    End Function

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                EntryMode = "N"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)
            Case "Update"
                Dim HasValidOrder As Boolean = False
                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                    If rowSOTORDR1.Item("ORDR_STATUS") = "L" Or rowSOTORDR1.Item("ORDR_STATUS") = "Q" Then
                        HasValidOrder = True
                    End If
                Next
                If HasValidOrder Then
                    Call Update_Record()
                    Dim Question As String = "Please Select The Number Of Copies You Want To Print"
                    Dim No_Of_Copies As Integer = 0
                    Dim iAnswer As String = InputBox(Question, "Print Order(s)", No_Of_Copies)
                    If IsNumeric(iAnswer) Then
                        No_Of_Copies = Val(iAnswer)
                    End If
                    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                        For i As Integer = 0 To No_Of_Copies - 1
                            Call Print_Record(True, rowSOTORDR1.Item("ORDR_NO").ToString(), rowSOTORDR1.Item("CUST_CODE").ToString(), 1)
                        Next
                    Next
                Else
                    Call Mode_Settings(False)
                End If
                Call Mode_Settings(False)
            Case "Cancel", "Done"
                Call Mode_Settings(False)

            Case "Delete"
                Call Delete_Record(grdSOTORDRX.Selected.Rows(0).Cells.Item("ORDR_NO").Text)
                Call Mode_Settings(False)
                SetFormToDefault()
                RefreshSOTORDRX()
            Case "Print"
                Call Print_Record(True, grdSOTORDRX.Selected.Rows(0).Cells("ORDR_NO").Text, grdSOTORDRX.Selected.Rows(0).Cells("CUST_CODE").Text)
            Case "Print Preview"
                Call Print_Record(False, grdSOTORDRX.Selected.Rows(0).Cells("ORDR_NO").Text, grdSOTORDRX.Selected.Rows(0).Cells("CUST_CODE").Text)
            Case "New Scan (Alt-1)"
                If VerifyEditMode() Then
                    ClearStyle()
                    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                        CalculateOrderTotal(rowSOTORDR1.Item("ORDR_NO").ToString)
                        CalculateOrderCuFt(rowSOTORDR1.Item("ORDR_NO").ToString)
                    Next
                End If
            Case "Hang Tag"
                If Absx1.txtFor("STYLE_CODE").Text.Length > 0 Then
                    Dim HANGTAG As New HANGTAG(Me, Absx1.txtFor("STYLE_CODE").Text, Discounts, "")
                    If HANGTAG.ErrMsg.Length = 0 Then
                        HANGTAG.Print()
                    End If
                End If
            Case "Excel"
                Dim ORDR_NO As String
                If grdSOTORDRX.Selected.Rows.Count = 0 Then
                    MsgBox("You Must Select An Order From The Orders List To Print", MsgBoxStyle.OkOnly, "Please Select An Order")
                    Exit Sub
                Else
                    ORDR_NO = grdSOTORDRX.Selected.Rows(0).Cells.Item("ORDR_NO").Text
                End If

                ASCMAIN1.Progress("Now Generating Excel Document")
                Cursor = Cursors.WaitCursor
                ExcelProcessInit()
                Dim excelFile As String = Generate_Excel(ORDR_NO)
                ExcelProcessKill()
                Cursor = Cursors.Default
                ASCMAIN1.Progress("")
                If excelFile <> "" Then
                    Dim start_excel As New Process
                    start_excel.StartInfo.Arguments = """" + excelFile + """ /e"
                    start_excel.StartInfo.FileName = EXCELDir & excelFile & ".xls"
                    start_excel.Start()
                End If
            Case "Customer Masterfile"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                If CUST_CODE.Length > 0 Then
                    Context_Launch("Edit", CUST_CODE, "Customer Masterfile", "SOTCUST1")
                Else
                    Context_Launch("Customer Masterfile", CUST_CODE, "Customer Masterfile", "SOTCUST1")
                End If
            Case "Refresh Orders"
                RefreshSOTORDRX()
            Case "Re-Price Order"
                Dim ValChange As Double = RePriceOrder(grdSOTORDRX.Selected.Rows(0).Cells("ORDR_NO").Text)
                MsgBox("Total Value Changed = " & Format(ValChange, "$###,##0.00"), MsgBoxStyle.OkOnly, "Order Re-Priced")
                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                    CalculateOrderTotal(rowSOTORDR1.Item("ORDR_NO").ToString)
                    CalculateOrderCuFt(rowSOTORDR1.Item("ORDR_NO").ToString)
                Next
            Case "Re-Price FEFD Order"
                Dim frmASFMSGBF As New ASFMSGBF
                Dim Label As New System.Text.StringBuilder With {.Length = 0}
                Label.AppendLine("Please Type One Of The Repricing Types To Proceed:")
                Label.AppendLine("FE")
                Label.AppendLine("FEMix")
                Label.AppendLine("FD")
                Label.AppendLine("FDMix")
                Dim Caption As String = "Re-pricing Choices"
                Dim FEFDType As String = frmASFMSGBF.Get_txtblock_from_User(Label.ToString, Caption, "", False, 5)
                FEFDType = FEFDType.ToUpper
                If FEFDType = "FE" Or FEFDType = "FD" Or FEFDType = "FEMIX" Or FEFDType = "FDMIX" Then
                    Me.Cursor = Cursors.WaitCursor
                    Dim ValChange As Double = RePriceFEFDOrder(grdSOTORDRX.Selected.Rows(0).Cells("ORDR_NO").Text, FEFDType)
                    Me.Cursor = Cursors.Default
                    MsgBox("Total Value Changed = " & Format(ValChange, "$###,##0.00"), MsgBoxStyle.OkOnly, "Order Re-Priced")
                    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                        CalculateOrderTotal(rowSOTORDR1.Item("ORDR_NO").ToString)
                        CalculateOrderCuFt(rowSOTORDR1.Item("ORDR_NO").ToString)
                    Next
                Else
                    MsgBox("You Did Not Type One Of the Correct Options", vbOKOnly, "Please Play Again Soon")
                End If
            Case "Find Style by Attribute"
                Dim ms As New System.Text.StringBuilder With {.Length = 0}
                ms.AppendLine("This Feature Has Been Moved To")
                ms.AppendLine("It's Own Screen That Is Now Avalable")
                ms.AppendLine(String.Format("On The Main Menu As {0}Search By Attribute{0}", Chr(34)))
                MsgBox(ms.ToString, vbOKOnly, "This Feature Has Moved")
                'Dim STYLE_CODE_selected As String = ""
                'Using F As New TAC.ICFATTR2(Me)
                '    Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
                '    If Not IsNothing(rowSOTPARM3) Then
                '        F.IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
                '        F.rbadDir = rowSOTPARM3.Item("RO_PARM_EXCEL_DIR").ToString()
                '    Else
                '        F.IMAGES_FOLDER = "C:\"
                '        F.rbadDir = "C:\"
                '    End If
                '    F.ShowDialog()
                '    STYLE_CODE_selected = F.STYLE_CODE
                'End Using
                'If STYLE_CODE_selected <> "" Then
                '    txtSTYLE_CODE.Text = STYLE_CODE_selected
                '    Click_Command("New Scan (Alt-1)")
                'End If
            Case "Record Credit Card"
                RecordCreditCard()
            Case "Update Addresses"
                UpdateAddresses()
                MsgBox("Addresses Updated", MsgBoxStyle.OkOnly, "Updated")
            Case "Tablet"
                Dim ORDR_GROUP_NO As String = grdSOTORDRX.Selected.Rows(0).Cells.Item("ORDR_GROUP_NO").Text
                Dim FILTER As String = String.Format("ORDR_GROUP_NO = '{0}' AND ORDR_STATUS = 'L'", ORDR_GROUP_NO)
                For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select(FILTER)
                    Dim ORDR_NO As String = rowSOTORDRX.Item("ORDR_NO").ToString
                    Call TABLETIZE_ORDER(ORDR_NO)
                    MsgBox("Order(s) Tabletized.", MsgBoxStyle.OkOnly, "Updated")
                    'Call TABLETIZE_ORDER(grdSOTORDRX.Selected.Rows(0).Cells.Item("ORDR_NO").Text)
                Next
                Call Mode_Settings(False)
                SetFormToDefault()
                RefreshSOTORDRX()
            Case "Recover"
                Dim sql As New System.Text.StringBuilder With {.Length = 0}
                sql.AppendLine("SELECT *")
                sql.AppendLine("FROM RECOVER1")
                Dim tblRECOVER1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())
                For Each rowRECOVER1 As DataRow In tblRECOVER1.Rows
                    Dim ORDR_NO As String = grdSHIP2.Rows(0).Cells("ORDR_NO").Text
                    If Not ErrIfOrderOnStyle(Absx1.txtFor("STYLE_CODE").Text, ORDR_NO) Then
                        AddSOTORDR2(ORDR_NO,
                                    rowRECOVER1.Item("STYLE_CODE").ToString,
                                    rowRECOVER1.Item("COLOR_CODE").ToString,
                                    Val(rowRECOVER1.Item("ORDR_QTY").ToString & String.Empty),
                                    Val(rowRECOVER1.Item("STYLE_PRICE").ToString & String.Empty), False)
                    End If
                Next
            Case "Verify ShipTo"
                shipToVerify(True)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Dim IsSuperUser As Boolean = False
        Call Set_ScreenMode_Base(tf)
        'tabORDERS.Tabs(1).Visible = ScreenMode

        ASCMAIN1.sql = String.Format("SELECT COUNT(*) AS RECCNT FROM ASTUSER2 WHERE USER_ID = '{0}' AND SECURITY_CODE = 'X6'", ASCMAIN1.USER_ID)
        If Val(ASCDATA1.GetDataValue) > 0 Then
            IsSuperUser = True
            btnCustOverride.Visible = True
            btnFEOverride.Visible = True
        Else
            IsSuperUser = False
            btnCustOverride.Visible = False
            btnFEOverride.Visible = False
        End If

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                '.Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Excel").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Refresh Orders").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Print Preview").Settings.Enabled = not_iScreenMode

                '.Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Delete").Settings.Enabled = not_iScreenMode
                '.Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Re-Price Order").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Find Style by Attribute").Settings.Enabled = iScreenMode

                '.Groups("Screen Control").Items("Cancel").Visible = ScreenMode And Not This_Record_Inquiry_Only
                '.Groups("Screen Control").Items("Done").Visible = This_Record_Inquiry_Only

                '.Groups("Screen Control").Items("Update").Visible = ScreenMode And Not This_Record_Inquiry_Only
                '.Groups("Screen Control").Items("Record Credit Card").Visible = ScreenMode And Not This_Record_Inquiry_Only
                .Groups("Screen Control").Items("New").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Edit").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Excel").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Refresh Orders").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Print").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Print Preview").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Delete").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode

                .Groups("Screen Control").Items("Update").Visible = ScreenMode
                .Groups("Screen Control").Items("Find Style by Attribute").Visible = ScreenMode
                .Groups("Screen Control").Items("Cancel").Visible = ScreenMode
                .Groups("Screen Control").Items("Record Credit Card").Visible = tabORDERS.Tabs(1).Visible And ScreenMode
                .Groups("Screen Control").Items("Update Addresses").Visible = tabORDERS.Tabs(1).Visible And ScreenMode
                .Groups("Screen Control").Items("Re-Price Order").Visible = tabORDERS.Tabs(1).Visible And ScreenMode
                If ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or
                    ASCMAIN1.USER_ID = "danny" Or ASCMAIN1.USER_ID = "mariog" Or
                    ASCMAIN1.USER_ID = "rich" Then
                    .Groups("Screen Control").Items("Re-Price FEFD Order").Visible = tabORDERS.Tabs(1).Visible And ScreenMode
                Else
                    .Groups("Screen Control").Items("Re-Price FEFD Order").Visible = False
                End If

                .Groups("Screen Control").Items("New Scan (Alt-1)").Visible = ScreenMode
                .Groups("Screen Control").Items("Hang Tag").Visible = ScreenMode
                .Groups("Screen Control").Items("Verify ShipTo").Visible = ScreenMode

                If IsSuperUser Then
                    .Groups("Screen Control").Items("Tablet").Visible = Not ScreenMode
                Else
                    .Groups("Screen Control").Items("Tablet").Visible = False
                End If

                If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                    .Groups("Screen Control").Items("Recover").Visible = True
                Else
                    .Groups("Screen Control").Items("Recover").Visible = False
                End If

                .Groups("Allocation").Visible = False
                .Groups("Image").Visible = False
                .Groups("FEFD").Visible = Not ScreenMode
                .Groups("FEFD").Expanded = False
                .Groups("Filter").Visible = Not ScreenMode
                .Groups("Filter").Expanded = Not ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)


        'If Not This_Record_Inquiry_Only Then
        '    Absx1.dteFor("ORDR_DATE").ReadOnly = False
        'End If

        grdSOTORDRX.Visible = Not tf

        With grdSOTORDR1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdSOTORDR1.DisplayLayout.Bands(0).Columns.Count - 1
            grdSOTORDR1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        For Each COLNAME As String In New String() {"ORDR_STATUS", "ORDR_CUST_PO", "ORDR_CATEGORY", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "SHIP_VIA_CODE", "WHSE_CODE", "ORDR_SHIP_INSTR", "TERM_CODE", "ORDR_MESSAGE", "FRT_TERMS"}
            grdSOTORDR1.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        Next
        For Each COLNAME As String In New String() {"ORDR_CUST_PO", "ORDR_CATEGORY", "ORDR_SHIP_INSTR", "ORDR_MESSAGE"}
            grdSOTORDR1.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Next

        If ScreenMode Then
            'With grdARTPYMT2.DisplayLayout.Override
            '    If This_Record_Inquiry_Only Then
            '        .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '        .AllowDelete = DefaultableBoolean.False
            '        .AllowUpdate = DefaultableBoolean.False
            '    Else
            '        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            '        .AllowDelete = DefaultableBoolean.True
            '        .AllowUpdate = DefaultableBoolean.True
            '    End If
            'End With
        Else
            RefreshSOTORDRX()
            SetFormToDefault()
        End If

        Dim SavedOverRideC As String = SOCMAIN2.GetSavedOverRide("C")
        Dim SavedOverRideF As String = SOCMAIN2.GetSavedOverRide("F")
        Dim TodaysOverRideC As String = SOCMAIN2.TodaysOverRide("C")
        Dim TodaysOverRideF As String = SOCMAIN2.TodaysOverRide("F")
        If SavedOverRideC = TodaysOverRideC Then
            txtCustOverride.Text = SavedOverRideC.ToString
            chkCustOverride.Checked = True
            txtCustOverride.Enabled = False
        Else
            If SavedOverRideC <> "" Then
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("UPDATE SOTPARMR SET RO_PARM_LAST_CUSTPASS = NULL WHERE SO_PARM_KEY = 'Z'")
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()
            End If
            txtCustOverride.Text = ""
            chkCustOverride.Checked = False
            txtCustOverride.Enabled = True
        End If

        'If SavedOverRideF = TodaysOverRideF Then
        '    txtFEOverride.Text = SavedOverRideF.ToString
        '    chkFEOverride.Checked = True
        '    txtFEOverride.Enabled = False
        '    UltraExplorerBar1.Groups("FEFD").Visible = True
        'Else
        '    txtFEOverride.Text = ""
        '    chkFEOverride.Checked = False
        '    txtFEOverride.Enabled = True
        '    If IsSuperUser Then
        '        UltraExplorerBar1.Groups("FEFD").Visible = True
        '    Else
        '        UltraExplorerBar1.Groups("FEFD").Visible = False
        '    End If

        'End If
        'Unlocked now per Danny - 01/06/14
        txtFEOverride.Text = SavedOverRideF.ToString
        chkFEOverride.Checked = True
        txtFEOverride.Enabled = False
        UltraExplorerBar1.Groups("FEFD").Visible = True
    End Sub

    Sub Clear_Record()
        dst.Tables("SOTORDRB").Rows.Clear()
        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR2").Rows.Clear()
        dst.Tables("SOTORDR5").Rows.Clear()
        dst.Tables("SOTORDRL").Rows.Clear()
        dst.Tables("ARTCUSTD").Rows.Clear()

        Absx1.txtFor("ORDR_BATCH_NO").Text = ""
        Absx1.txtFor("CUST_NAME").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""
        rowARTCUST1 = Nothing
        ReBatch.Clear()
        lblEXCLUSIVE_STYLE.Visible = False
        lblDESIGNER_STYLE.Visible = False
        '.Groups("Customer Messages").Visible = False
        'This_Record_Inquiry_Only = False
        'Call Load_ARTPYMTX()
    End Sub

    Sub Load_Record()

        Call Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            ORDR_BATCH_NO = ASCMAIN1.Next_Control_No("ORDR_BATCH_NO")
            SetrowARTCUST1()
            HFs("ORDR_BATCH_NO") = ORDR_BATCH_NO
            AddSOTORDRB(ORDR_BATCH_NO)
            Absx1.txtFor("ORDR_BATCH_NO").Text = ORDR_BATCH_NO
        End If

        Dim ORDR_NOs(0) As String
        Dim ORDRcnt As Integer = 0
        Using tblSOTORDR1 As DataTable = ASCDATA1.GetDataTable(String.Format("SELECT ORDR_NO FROM SOTORDR1 WHERE (ORDR_GROUP_NO = '{0}' OR ORDR_BATCH_NO = '{0}')", Absx1.txtFor("ORDR_BATCH_NO").Text))
            For Each rowSOTORDR1 As DataRow In tblSOTORDR1.Rows
                ReDim Preserve ORDR_NOs(ORDRcnt)
                ORDR_NOs(ORDRcnt) = rowSOTORDR1.Item("ORDR_NO").ToString
                ORDRcnt += 1
            Next
        End Using

        EnforceConstraints(False)
        If EntryMode = "E" Then
            For Each rowORDR_NO As String In ORDR_NOs
                Call Fill_Records("SOTORDR1", rowORDR_NO, False)
                Call Fill_Records("SOTORDR2", rowORDR_NO, False)
                Call Fill_Records("SOTORDR5", rowORDR_NO, False)
                Call Fill_Records("SOTORDRL", rowORDR_NO, False)
            Next
            Call Fill_Records("SOTORDRB", Absx1.txtFor("ORDR_BATCH_NO").Text, False)
            If dst.Tables.Item("SOTORDRB").Rows.Count > 0 Then
                Absx1.txtFor("CUST_CODE").Text = dst.Tables.Item("SOTORDRB").Rows(0).Item("CUST_CODE")
            End If
        End If
        Call Fill_Records("ARTCUST1", Absx1.txtFor("CUST_CODE").Text, True)
        Call Fill_Records("ARTCUST2", Absx1.txtFor("CUST_CODE").Text, True)
        'CHANGES FOR SHIPTO QUESTION - Call Fill_Records("ARTCUSTQ", Absx1.txtFor("CUST_CODE").Text, True)
        EnforceConstraints(True)


        For Each rowSOTORDRL As DataRow In dst.Tables("SOTORDRL").Select()
            Dim FEFDFACTOR As Double = 0
            FEFDFACTOR = Val(rowSOTORDRL.Item("FEFDFACTOR").ToString & "")
            FEFDLEVEL = rowSOTORDRL.Item("FEFDLEVEL").ToString & ""
            numFEFDFACTOR.Value = FEFDFACTOR
        Next

        'This will set the PINNED_CUST_PRICE_TIER_PVC for all the order on this Group.
        'Spoke to Danny about this on 6/12 and he is OK with that.
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            If rowSOTORDR1.Item("ORDR_NO_WEB") & "" <> "" Then
                PINNED_CUST_PRICE_TIER_PVC = rowSOTORDR1.Item("ORDR_NO_WEB") & ""
                If PINNED_CUST_PRICE_TIER_PVC = "" Then
                    PINNED_CUST_PRICE_TIER_PVC = "PC"
                End If
                If dst.Tables("ARTCUST1").Rows.Count > 0 Then
                    Select Case PINNED_CUST_PRICE_TIER_PVC
                        Case Is = "FC", "5C"
                            dst.Tables("ARTCUST1").Rows(0).Item("CUST_PRICE_TIER_PVC") = PINNED_CUST_PRICE_TIER_PVC
                        Case Else
                            PINNED_CUST_PRICE_TIER_PVC = dst.Tables("ARTCUST1").Rows(0).Item("CUST_PRICE_TIER_PVC") & ""
                    End Select
                End If
            End If
        Next

        SetFKeys()
        ShowStyles()

        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
        ORDR_BATCH_NO = Absx1.txtFor("ORDR_BATCH_NO").Text
        SetrowARTCUST1()

        Dim OneOrderIsTransferred As Boolean = False
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            CalculateOrderTotal(rowSOTORDR1.Item("ORDR_NO").ToString)
            CalculateOrderCuFt(rowSOTORDR1.Item("ORDR_NO").ToString)
            If Not OneOrderIsTransferred And rowSOTORDR1.Item("ORDR_STATUS") = "Q" Then
                'rdoQUOTE.Checked = True
                EnableForm(True)
            Else
                'rdoQUOTE.Checked = False
                If Not OneOrderIsTransferred And rowSOTORDR1.Item("ORDR_STATUS") = "L" Then
                    EnableForm(True)
                Else
                    EnableForm(False)
                    OneOrderIsTransferred = True
                End If
            End If
        Next

        If EntryMode = "N" Then
            EnableForm(True)
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        Call BeginTrans()
        For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
            ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next
        Call CommitTrans("Order / Quote Deleted")
    End Sub

    Private Sub SyncLTables()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Rows
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
                Dim TABLE_NAME_L As String = String.Format("{0}_L", TABLE_NAME)
                ASCMAIN1.sql = String.Format("Delete from {0} where ORDR_NO = '{1}'", TABLE_NAME_L, ORDR_NO)
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = String.Format("Insert into {0} Select * from {1} where ordr_no = '{2}'", TABLE_NAME_L, TABLE_NAME, ORDR_NO)
                ASCDATA1.ExecuteSQL()
            Next
        Next
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            If ReBatch.Contains(rowSOTORDR1.Item("ORDR_NO").ToString) Then
                rowSOTORDR1.Item("ORDR_BATCH_NO") = ASCMAIN1.Next_Control_No("ORDR_BATCH_NO")
                AddSOTORDRB(rowSOTORDR1.Item("ORDR_BATCH_NO"))
            End If
            Dim ORDR_FOUND As Boolean = False
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO").ToString & ""
            For Each rowSOTORDRL As DataRow In dst.Tables("SOTORDRL").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
                rowSOTORDRL.Item("FEFDFACTOR") = Val(numFEFDFACTOR.Value & "")
                rowSOTORDRL.Item("FEFDLEVEL") = FEFDLEVEL
                ORDR_FOUND = True
            Next
            If Not ORDR_FOUND Then
                Dim newSOTORDRL As DataRow = dst.Tables.Item("SOTORDRL").NewRow
                newSOTORDRL.Item("ORDR_NO") = ORDR_NO
                newSOTORDRL.Item("FEFDFACTOR") = Val(numFEFDFACTOR.Value & "")
                newSOTORDRL.Item("FEFDLEVEL") = FEFDLEVEL
                dst.Tables("SOTORDRL").Rows.Add(newSOTORDRL)
            End If
        Next



        Dim OrderType As String = "K"

        Call CreateSOTORDRB()

        Call Update_Record_TDA("SOTORDRB")
        Call Update_Record_TDA("SOTORDR1")
        Call Update_Record_TDA("SOTORDR2")
        Call Update_Record_TDA("SOTORDR5")
        Call Update_Record_TDA("SOTORDRL")

        SyncLTables()
        'NOTE: ARTCUST1 CUST_PRICE_TIER_PVC HAS BEEN MODIFIED.   IF YOU EVER ADD IT TO UPDATE RESET IT BACK HERE.
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
                If chkCustOverride.Checked = False Then
                    If Not Remote.IsUserSuper Then
                        sql_where = Remote.SQLWhere & " OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1_L)"
                    End If
                End If
            Case "SHIP_VIA_CODE"
                sql_where &= "NVL(SHIP_VIA_STATUS,'A') = 'A'"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
        '    Stop
        '    Dim sql As New System.Text.StringBuilder With {.Length = 0}
        '    sql.AppendLine("SELECT")
        '    sql.AppendLine("S1.CUST_NAME, S1.CUST_CODE, S1.ORDR_NO")
        '    sql.AppendLine("FROM SOTORDR1 S1, SOTORDR2 S2")
        '    sql.AppendLine("WHERE S1.ORDR_NO = S2.ORDR_NO")
        '    sql.AppendLine("AND (S1.ORDR_DATE >= '01-SEP-2018' AND ORDR_DATE <= '15-NOV-2018')")
        '    sql.AppendLine("AND S1.CUST_CODE IN")
        '    sql.AppendLine("('051315','160235','010520','305448','070255','190828','160086','230695',")
        '    sql.AppendLine("'080512','130236','230090','231417','220127','060481','190283','304308',")
        '    sql.AppendLine("'190418','214478','210067','020311','135151','080225','301205','305442',")
        '    sql.AppendLine("'040305','081762','040310','193085','010488','190364','308361','180257',")
        '    sql.AppendLine("'309541','110956','070101','160170','110418','150003')")
        '    sql.AppendLine("GROUP BY")
        '    sql.AppendLine("S1.CUST_NAME, S1.CUST_CODE,")
        '    sql.AppendLine("S1.ORDR_NO")
        '    Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        '    For Each rowORDERS As DataRow In tbl.Rows
        '        Dim CUST_CODE_PRINT As String = rowORDERS.Item("CUST_CODE").ToString & String.Empty
        '        Dim CUST_NAME_PRINT As String = rowORDERS.Item("CUST_NAME").ToString & String.Empty
        '        Dim ORDR_NO_PRINT As String = rowORDERS.Item("ORDR_NO").ToString & String.Empty
        '        Print_Records_PDF("", ORDR_NO_PRINT, CUST_CODE_PRINT, CUST_NAME_PRINT)
        '    Next
        '    Stop
        '    Exit Sub
        'End If
        Print_Report_Begin()
        'frm.CR_params.Add("SUBT", "")
        'Fill SOTORDRP records
        Fill_Records("ARTCUST1", CUST_CODE, True)
        Fill_Records("SOTORDP1", ORDR_NO, True)
        Fill_Records("SOTORDP2", ORDR_NO, True)
        Fill_Records("SOTORDP5", ORDR_NO, True)
        If dst.Tables.Item("SOTORDP1").Rows.Count > 0 Then
            Dim TERM_CODE As String = dst.Tables.Item("SOTORDP1").Rows(0).Item("TERM_CODE")
            Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE)
            CR_params.Add("TERM_DESC", rowTATTERM1.Item("TERM_DESC").ToString)
        Else
            CR_params.Add("TERM_DESC", "")
        End If
        For Each rowSOTORDP2 As DataRow In dst.Tables("SOTORDP2").Select()
            Dim STYLE_CODE As String = rowSOTORDP2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDP2.Item("COLOR_CODE")
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            rowSOTORDP2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            If IsNothing(rowICTSTYC1) Then
                rowSOTORDP2.Item("UPC_CODE") = ""
            Else
                rowSOTORDP2.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE") & ""
            End If

            Dim FACTORY_CODE As String = GetVendorData(rowICTSTYL1.Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
            rowSOTORDP2.Item("FACTORY_CODE") = FACTORY_CODE
        Next
        Generate_Report("SORORDRO")
        If PrintToDefault Then
            Dim PS As New System.Drawing.Printing.PrinterSettings
            Print_Report_End(True, False, PS.PrinterName, No_Of_Copies)
        Else
            Print_Report_End()
        End If
    End Sub

    Private Sub Print_Records_PDF(ByVal FileLoc As String, ByVal ORDR_NO As String, ByVal CUST_CODE As String, CUST_NAME As String, Optional ByVal No_Of_Copies As Integer = 1)
        Print_Report_Begin()
        'frm.CR_params.Add("SUBT", "")
        'Fill SOTORDRP records
        Fill_Records("ARTCUST1", CUST_CODE, True)
        Fill_Records("SOTORDP1", ORDR_NO, True)
        Fill_Records("SOTORDP2", ORDR_NO, True)
        Fill_Records("SOTORDP5", ORDR_NO, True)
        If dst.Tables.Item("SOTORDP1").Rows.Count > 0 Then
            Dim TERM_CODE As String = dst.Tables.Item("SOTORDP1").Rows(0).Item("TERM_CODE")
            Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE)
            CR_params.Add("TERM_DESC", rowTATTERM1.Item("TERM_DESC").ToString)
        Else
            CR_params.Add("TERM_DESC", "")
        End If
        For Each rowSOTORDP2 As DataRow In dst.Tables("SOTORDP2").Select()
            Dim STYLE_CODE As String = rowSOTORDP2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTORDP2.Item("COLOR_CODE")
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            rowSOTORDP2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            If IsNothing(rowICTSTYC1) Then
                rowSOTORDP2.Item("UPC_CODE") = ""
            Else
                rowSOTORDP2.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE") & ""
            End If

            Dim FACTORY_CODE As String = GetVendorData(rowICTSTYL1.Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
            rowSOTORDP2.Item("FACTORY_CODE") = FACTORY_CODE
        Next
        Generate_Report("SORORDRO",,, , "PDF", CUST_NAME & "_" & CUST_CODE & "_" & ORDR_NO)
        Print_Report_End(True, False, "", No_Of_Copies)
        'If PrintToDefault Then
        '    Dim PS As New System.Drawing.Printing.PrinterSettings
        '    Print_Report_End(True, False, PS.PrinterName, No_Of_Copies)
        'Else
        '    Print_Report_End()
        'End If
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDR2, "BB", "Edit Line", "Delete Selected Line")
        'Load_Popup_Menu(grdSOTORDR1, "B", "Edit Ship To")
        Load_Popup_Menu(grdSOTORDRX, "SSB", "Show Filter", "Show GroupBox")
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            Case "grdSOTORDR2"
                If Not InquiryOnly Then
                    e.Tool.ToolbarsManager.Tools("Edit Line").SharedProps.Visible = True
                    e.Tool.ToolbarsManager.Tools("Delete Selected Line").SharedProps.Visible = True
                End If
            Case "grdSOTORDR1"
                If Not InquiryOnly Then
                    e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
                End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Edit Line"
                If Not InquiryOnly Then
                    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                    Dim ORDR_LNO As Integer = grd.ActiveRow.Cells("ORDR_LNO").Text
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                    EditOrderLine(ORDR_NO, ORDR_LNO, STYLE_CODE)
                    EditOrderLineMode = True
                End If
            Case "Delete Selected Line"
                If Not InquiryOnly Then
                    Dim ORDR_NO As String = grdSOTORDR2.ActiveRow.Cells("ORDR_NO").Text
                    grdSOTORDR2.ActiveRow.Delete()
                    CalculateOrderTotal(ORDR_NO)
                    'Dim SelCnt As Integer = grdSOTORDR2.Selected.Rows.Count
                    'If SelCnt < 1 Then
                    '    MsgBox("You Must Select At Least One Line To Delete")
                    'End If
                    'Dim iResult As MsgBoxResult
                    'Dim iTitle As String = "Line Deletion"
                    'Dim iMSG As New System.Text.StringBuilder
                    'iMSG.AppendLine(String.Format("Are You Sure You Want To Delete The {0} Selected Line(s)", SelCnt))
                    'iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)Setup_SOTORDR2
                    'If iResult = MsgBoxResult.Yes Then
                    '    DeleteOrdrLines()
                    'End If
                End If
            Case "Edit Ship To"
                If Not InquiryOnly Then
                    MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "Form Controls"
    Private Sub btnAddShipTo_Click(sender As System.Object, e As System.EventArgs) Handles btnAddShipTo.Click
        Dim TATCTLN3 As New TATCTLN3("SOTORDR1.ORDR_NO", Me)
        If Not IsNothing(TATCTLN3.ErrMsg) Then
            MsgBox(TATCTLN3.ErrMsg, MsgBoxStyle.OkOnly, "Problem Getting Next Order Number")
            Exit Sub
        End If
        If InquiryOnly Then
            MsgBox("You May Not Add A Ship-To When One Of The Orders Are Already Transferred.", MsgBoxStyle.Critical, "Order Already Transferred")
            Exit Sub
        End If
        If TATCTLN3.NumbersRemaining < 20 Then
            Dim msg As String = String.Format("You Only Have {0} Order Numbers Left", TATCTLN3.NumbersRemaining)
            msg = msg & vbCrLf & "You Should Fetch Some More From The Transfer Screen Soon."
            MsgBox(msg, MsgBoxStyle.Critical, "Running Low On Order Numbers")
        End If
        Using frm As New SOFORDR2(Me, rowARTCUST1, Absx1.txtFor("ORDR_CUST_PO").Text, "", True)
            With frm
                .ShowDialog()
                If .SELECT_ST.Length <> 0 Then
                    If Not IsNothing(.rowARTCUST2) And Not IsNothing(rowARTCUST1) Then
                        PINNED_CUST_PRICE_TIER_PVC = .CUST_PRICE_TIER_PVC
                        rowARTCUST1.Item("CUST_PRICE_TIER_PVC") = .CUST_PRICE_TIER_PVC
                        AddSOTORDR1(.rowARTCUST2, TATCTLN3.Next_ctl_no, frm)
                        AddSOTORDR5(.rowARTCUST2, TATCTLN3.Next_ctl_no)
                        Dim dstChanges As DataTable = dst.Tables.Item("ARTCUSTD").GetChanges
                        'CHANGES FOR SHIPTO QUESTION - Dim dstChangesQ As DataTable = dst.Tables.Item("ARTCUSTQ").GetChanges
                        'CHANGES FOR SHIPTO QUESTION - 
                        'If Not IsNothing(dstChanges) Or Not IsNothing(dstChangesQ) Then
                        If Not IsNothing(dstChanges) Then
                            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                            Update_Record_TDA("ARTCUSTD")
                            'CHANGES FOR SHIPTO QUESTION - Update_Record_TDA("ARTCUSTQ")
                            'For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUSTD", "ARTCUSTQ"}
                            For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUSTD"}
                                ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & "_L where CUST_CODE = '" & CUST_CODE & "'")
                                ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & "_H where CUST_CODE = '" & CUST_CODE & "'")
                                ASCDATA1.ExecuteSQL("Insert into " & TABLE_NAME & "_L Select * from " & TABLE_NAME & " where CUST_CODE = '" & CUST_CODE & "'")
                                ASCDATA1.ExecuteSQL("Insert into " & TABLE_NAME & "_H Select * from " & TABLE_NAME & " where CUST_CODE = '" & CUST_CODE & "'")
                            Next
                        End If
                    End If
                End If
            End With
        End Using
        ShowStyles()
    End Sub

    Private Sub btnFDMixPrice_Click(sender As System.Object, e As System.EventArgs) Handles btnFDMixPrice.Click
        SetFEPics("picFDMixPrice")
    End Sub

    Private Sub btnFDPrice_Click(sender As System.Object, e As System.EventArgs) Handles btnFDPrice.Click
        SetFEPics("picFDPrice")
    End Sub

    Private Sub btnFEMixPrice_Click(sender As System.Object, e As System.EventArgs) Handles btnFEMixPrice.Click
        SetFEPics("picFEMixPrice")
    End Sub

    Private Sub btnFEPrice_Click(sender As System.Object, e As System.EventArgs) Handles btnFEPrice.Click
        SetFEPics("picFEPrice")
    End Sub

    Private Sub optShowOrders_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShowOrders.ValueChanged
        If Not LoadFinished Then
            Exit Sub
        Else
            RefreshSOTORDRX()
        End If
    End Sub

    Private Sub tabORDERS_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabORDERS.SelectedTabChanged
        If EditOrderLineMode Then
            tabORDERS.Tabs.Item(1).Selected = True
            If Not EditOrderLineWarning Then
                Dim msg As String = "You Must Finish Editing The Line Before Moving Off This Screen"
                MsgBox(msg, MsgBoxStyle.Critical, "Currently Editing Line")
                EditOrderLineWarning = True
            End If
            Exit Sub
        End If
        If tabORDERS.Tabs.Item(1).Selected = True Then
            Dim shipRows As Integer = dst.Tables.Item("SOTORDR5").Rows.Count
            If shipRows = 0 Then
                MsgBox("You Must Have At Least One ShipTo Selected", MsgBoxStyle.Critical, "ShipTo")
                tabORDERS.Tabs.Item(0).Selected = True
            End If
        End If
        If SelectedTabHit Then
            ClearStyle()
        Else
            SelectedTabHit = True
        End If
    End Sub

    Private Sub txtCustOverride_TextChanged(sender As Object, e As System.EventArgs) Handles txtCustOverride.TextChanged
        Dim CustPassword As String = SOCMAIN2.TodaysOverRide("C")
        If txtCustOverride.Text = CustPassword.ToString Then
            chkCustOverride.Checked = True
            txtCustOverride.Enabled = False
            btnCustOverride.Enabled = False
            SOCMAIN2.SaveOverRide("C", txtCustOverride.Text)
        Else
            chkCustOverride.Checked = False
            txtCustOverride.Enabled = True
            btnCustOverride.Enabled = True
        End If
        RefreshSOTORDRX()
    End Sub

    Private Sub txtFEOverride_TextChanged(sender As Object, e As System.EventArgs) Handles txtFEOverride.TextChanged
        Dim FEPassword As String = SOCMAIN2.TodaysOverRide("F")
        If txtFEOverride.Text = FEPassword.ToString Then
            chkFEOverride.Checked = True
            txtFEOverride.Enabled = False
            btnFEOverride.Enabled = False
            SOCMAIN2.SaveOverRide("F", txtFEOverride.Text)
            UltraExplorerBar1.Groups("FEFD").Visible = True
        Else
            chkFEOverride.Checked = False
            txtFEOverride.Enabled = True
            btnFEOverride.Enabled = True
            UltraExplorerBar1.Groups("FEFD").Visible = False
        End If
    End Sub

    Private Sub txtSTYLE_CODE_MouseEnter(sender As Object, e As System.EventArgs) Handles txtSTYLE_CODE.MouseEnter
        ShowClassTip(sender, e)
    End Sub
#End Region

#Region "Grids"
#Region "grdSHIP2"

    Private Sub grdSHIP2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSHIP2.ClickCellButton
        Dim ORDR_NO As String = grdSHIP2.ActiveRow.Cells("ORDR_NO").Text
        If Not ErrIfOrderOnStyle(Absx1.txtFor("STYLE_CODE").Text, ORDR_NO) Then
            If Not HasDiscColors(Absx1.txtFor("STYLE_CODE").Text) Then
                SaveDetailsToOrder(ORDR_NO)
                EditOrderLineMode = False
                EditOrderLineWarning = False
            End If
        End If
    End Sub

    Private Sub grdSHIP2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSHIP2.InitializeRow
        Dim FKEYTIP As String = ""
        If e.Row.Cells("ORDR_CUST_PO").Text.Length > 0 Then
            FKEYTIP = "PO: " & e.Row.Cells("ORDR_CUST_PO").Text
            If e.Row.Cells("ORDR_CATEGORY").Text.Length > 0 Then
                FKEYTIP = FKEYTIP & vbCrLf & "TH: " & e.Row.Cells("ORDR_CATEGORY").Text
            End If
        Else
            If e.Row.Cells("ORDR_CATEGORY").Text.Length > 0 Then
                FKEYTIP = "TH: " & e.Row.Cells("ORDR_CATEGORY").Text
            End If
        End If

        e.Row.Cells("TCUFT").ToolTipText = "Total Cubic Feet"
        e.Row.Cells("TORDR").ToolTipText = "Order Total"
        e.Row.Cells("FKEY").ToolTipText = FKEYTIP
        e.Row.Cells("WHSE_CODE").ToolTipText = e.Row.Cells("ORDR_NO").Text
    End Sub
#End Region

#Region "grdSOTORDR1"
    Private Sub grdSOTORDR1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR1.AfterRowActivate
        Setup_SOTORDR2()
    End Sub

    Private Sub grdSOTORDR1_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDR1.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case Is = "FRT_TERMS"
                ASCMAIN1.sql = "SELECT COUNT(*) FROM ASTCODE1"
                ASCMAIN1.sql = ASCMAIN1.sql & " WHERE COLUMN_NAME = 'FRT_TERMS'"
                ASCMAIN1.sql = ASCMAIN1.sql & " AND TABLE_NAME = 'SOTORDR1'"
                ASCMAIN1.sql = String.Format("{0} AND T_CODE = '{1}'", ASCMAIN1.sql, e.NewValue & String.Empty)
                Dim RecCount As Int16 = Val(ASCDATA1.GetDataValue)
                If RecCount = 0 Then
                    MsgBox("Invalid Freight Terms", MsgBoxStyle.Critical, "Freight")
                    e.Cancel = True
                Else
                    Dim WHSE_CODE As String = e.Cell.Row.Cells.Item("WHSE_CODE").Text.ToString & String.Empty
                    Dim FRT_TERMS As String = e.NewValue
                    If WHSE_CODE = "FE" Then
                        If FRT_TERMS <> "PPA" And FRT_TERMS <> "COL" Then
                            MsgBox("Selection of FE Warehouse Requires Either PPA or COL Freight Terms!", MsgBoxStyle.OkOnly, "Invalid Terms")
                            e.Cancel = True
                        End If
                    End If
                    If WHSE_CODE = "FD" Then
                        If FRT_TERMS <> "PPD" Then
                            MsgBox("Selection of FD Warehouse Requires PPD Freight Terms!", MsgBoxStyle.OkOnly, "Invalid Terms")
                            e.Cancel = True
                        End If
                    End If
                End If
            Case Is = "WHSE_CODE"
                Dim rowRec As DataRow = LookUp("ICTWHSE1", e.Cell.Text)
                If IsNothing(rowRec) Then
                    MsgBox("Invalid Warehouse Code", MsgBoxStyle.Critical, "Warehouse")
                    e.Cancel = True
                Else
                    Dim WHSE_CODE As String = e.NewValue
                    Dim FRT_TERMS As String = e.Cell.Row.Cells.Item("FRT_TERMS").Text.ToString & String.Empty
                    If WHSE_CODE = "FE" Then
                        If FRT_TERMS <> "PPA" And FRT_TERMS <> "COL" Then
                            MsgBox("Selection of FE Warehouse Requires Either PPA or COL Freight Terms!", MsgBoxStyle.OkOnly, "Invalid Terms")
                            e.Cancel = True
                        End If
                    End If
                    If WHSE_CODE = "FD" Then
                        If FRT_TERMS <> "PPD" Then
                            MsgBox("Selection of FD Warehouse Requires PPD Freight Terms!", MsgBoxStyle.OkOnly, "Invalid Terms")
                            e.Cancel = True
                        End If
                    End If
                End If
            Case Is = "SHIP_VIA_CODE"
                Dim rowRec As DataRow = LookUp("SOTSVIA1", e.NewValue)
                If IsNothing(rowRec) Then
                    MsgBox("Invalid Ship Via", MsgBoxStyle.Critical, "Ship Via")
                    e.Cancel = True
                End If
            Case Is = "ORDR_STATUS"
                If e.NewValue <> "L" And e.NewValue <> "Q" Then
                    MsgBox(String.Format("You May Not Change an Order To {0}", e.Cell.Text), MsgBoxStyle.Critical, "Order Type")
                    e.Cancel = True
                End If
                If e.NewValue = "L" Then
                    Dim QFound As Boolean = False
                    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                        If rowSOTORDR1.Item("ORDR_STATUS") = "Q" Then
                            QFound = "True"
                        End If
                    Next
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Un-Grouping"
                    Dim iMSG As New System.Text.StringBuilder
                    iMSG.AppendLine("Changing A Quote To An Order On A Group")
                    iMSG.AppendLine("That Contains Other Quotes Will Un-Group")
                    iMSG.AppendLine("These Orders.")
                    iMSG.AppendLine("Are You OK With That?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.Yes Then
                        ReBatch.Add(grdSOTORDR1.ActiveRow.Cells.Item("ORDR_NO").Text.ToString)
                    Else
                        e.Cancel = True
                    End If
                End If
            Case "TERM_CODE"
                If Not SOCMAINL.IsValidTerms(Absx1.txtFor("CUST_CODE").Text, e.NewValue & String.Empty) Then
                    e.Cancel = True
                End If
        End Select
    End Sub

    Private Sub grdSOTORDR1_BeforeEnterEditMode(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles grdSOTORDR1.BeforeEnterEditMode
        If grdSOTORDR1.ActiveRow.Cells("ORDR_STATUS").Value <> "L" And grdSOTORDR1.ActiveRow.Cells("ORDR_STATUS").Value <> "Q" Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdSOTORDR1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR1.ClickCellButton
        Dim ORDR_STATUS As String = grdSOTORDR1.ActiveRow.Cells("ORDR_STATUS").Value
        If ORDR_STATUS = "L" Or ORDR_STATUS = "Q" Then
            Select Case e.Cell.Column.Key
                Case "WHSE_CODE", "SHIP_VIA_CODE", "FRT_TERMS"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTORDR1, sql_where)
                Case "TERM_CODE"
                    Dim sql_where As String = "TERM_STATUS = 'A'"
                    grdClickCellButton(grdSOTORDR1, sql_where)
                    Select Case grdSOTORDR1.ActiveCell.Text
                        Case Is = "AMEX", "DISC", "MC", "VISA"
                            grdSOTORDR1.ActiveCell.Value = "CRED"
                    End Select
                Case "ORDR_STATUS"
                    If ORDR_STATUS = "L" Then
                        grdSOTORDR1.ActiveRow.Cells("ORDR_STATUS").Value = "Q"
                    End If
                    If ORDR_STATUS = "Q" Then
                        grdSOTORDR1.ActiveRow.Cells("ORDR_STATUS").Value = "L"
                    End If
            End Select
        End If
    End Sub

    Private Sub grdSOTORDR1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR1.InitializeRow
        With e.Row
            If .Cells("WHSE_CODE").Value = "FE" Or .Cells("WHSE_CODE").Value = "FD" Then
                OrderHasFEWhse = True
                tabPricing.Tabs(1).Visible = True
            End If
        End With
    End Sub
#End Region

#Region "grdSOTORDRX"
    Private Sub grdSOTORDRX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDRX.DoubleClickRow
        If Not IsDBNull(e.Row.Cells("ORDR_GROUP_NO").Value) Then
            If IsDBNull(e.Row.Cells("ORDR_BATCH_NO").Value) Then
                Absx1.txtFor("ORDR_BATCH_NO").Text = e.Row.Cells("ORDR_GROUP_NO").Value
            Else
                Absx1.txtFor("ORDR_BATCH_NO").Text = e.Row.Cells("ORDR_BATCH_NO").Value
            End If

            If e.Row.Cells("CUST_CODE").Value & String.Empty <> "" Then
                Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value & String.Empty
                Click_Command("Edit")
            End If

        End If
    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow
        With e.Row
            Dim ORDR_SOURCE As String = .Cells("ORDR_SOURCE").Value & ""
            Dim ORDR_STATUS As String = .Cells("ORDR_STATUS").Value & ""
            Dim ORDR_GROUP_NO As String = .Cells("ORDR_GROUP_NO").Value & ""
            If ORDR_GROUP_NO.Length = 0 Then
                .Cells("ORDR_GROUP_NO").Value = .Cells("ORDR_BATCH_NO").Value
            End If
            If ORDR_STATUS = "L" Then
                Select Case ORDR_SOURCE
                    Case "L"
                        .Appearance.BackColor = Drawing.Color.Empty
                    Case "Q"
                        .Appearance.BackColor = Drawing.Color.BlanchedAlmond
                    Case Else
                        .Appearance.BackColor = Drawing.Color.Cyan
                End Select
            Else
                If ORDR_STATUS = "Q" Then
                    .Appearance.BackColor = Drawing.Color.BlanchedAlmond
                Else
                    .Appearance.BackColor = Drawing.Color.Cyan
                End If
            End If
            'CalculateOrderTotal(.Cells("ORDR_NO").Value)
            grdSOTORDRX.UpdateData()
        End With
    End Sub

#End Region

#Region "grdSOTORDR2"
    Private Sub grdSOTORDR2_AfterSelectChange(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdSOTORDR2.AfterSelectChange
        If Not IsNothing(grdSOTORDR2.ActiveRow) Then
            imgSTYL1.ImageLocation = GetImageLocation(grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Text, grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Text)
            UltraExplorerBar1.Groups("Image").Visible = True
        End If
    End Sub
#End Region

#Region "grdICTSTYC1"
    Private Sub grdICTSTYC1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTYC1.AfterRowActivate
        If grdICTSTYC1.ActiveRow Is Nothing Then
            UltraExplorerBar1.Groups("Allocation").Visible = False
        Else
            Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
            Dim WHSE_CODE As String = "MS"
            FilterAlloc(STYLE_CODE, COLOR_CODE, WHSE_CODE)
            UltraExplorerBar1.Groups("Allocation").Visible = True
            imgSTYL1.ImageLocation = GetImageLocation(STYLE_CODE, COLOR_CODE)
            If imgSTYL1.ImageLocation.Length > 0 Then
                UltraExplorerBar1.Groups("Image").Visible = True
            Else
                UltraExplorerBar1.Groups("Image").Visible = False
            End If
        End If
    End Sub

    Private Sub grdICTSTYC1_AfterRowRegionScroll(sender As Object, e As UltraWinGrid.RowScrollRegionEventArgs) Handles grdICTSTYC1.AfterRowRegionScroll
        If e.RowScrollRegion.VisibleRows.Count > 0 Then
            SetColorStatusImages(e.RowScrollRegion.VisibleRows(0).Row.Index, False)
        End If
    End Sub

    Private Sub grdICTSTYC1_AfterSelectChange(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdICTSTYC1.AfterSelectChange
        grdICTSTYC1.ActiveRow.Cells("ORDR_QTY").Selected = True
    End Sub

    Private Sub grdICTSTYC1_Click(sender As Object, e As System.EventArgs) Handles grdICTSTYC1.Click
        If Not IsNothing(grdICTSTYC1.ActiveRow) Then
            grdICTSTYC1.ActiveRow.Cells("ORDR_QTY").Activated = True
        End If
    End Sub

    Private Sub grdICTSTYC1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYC1.InitializeRow
        With e.Row
            Dim STYLE_COLOR_STATUS As String = .Cells("STYLE_COLOR_STATUS").Value & ""
            Select Case STYLE_COLOR_STATUS
                Case "D"
                    .Appearance.BackColor = Drawing.Color.Red
                Case "N"
                    .Appearance.BackColor = Drawing.Color.Yellow
                Case Else
                    .Appearance.BackColor = Drawing.Color.Empty
            End Select
        End With
    End Sub
#End Region
#End Region

#Region "Custom Methods"

    Private Sub AddSOTORDR1(ByVal rowARTCUST2 As DataRow, ByVal ORDR_NO As String, ByRef frm As SOFORDR2)
        SREP_CODE = setSREP_CODE()
        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow()
        rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
        rowSOTORDR1.Item("ORDR_BATCH_NO") = Absx1.txtFor("ORDR_BATCH_NO").Text
        rowSOTORDR1.Item("ORDR_DATE") = Now().Date
        rowSOTORDR1.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
        rowSOTORDR1.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
        rowSOTORDR1.Item("CUST_STORE_NO") = rowARTCUST2.Item("CUST_ADDR_CODE")
        rowSOTORDR1.Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_NAME")
        rowSOTORDR1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowSOTORDR1.Item("INIT_DATE") = DATETIME_STAMP
        rowSOTORDR1.Item("LAST_DATE") = DATETIME_STAMP
        rowSOTORDR1.Item("ORDR_DATE_RECD") = Now().Date
        rowSOTORDR1.Item("ORDR_SOURCE") = "L"
        'If frm.rdoQUOTE.Checked Then
        '    Using frmQ As New SOFORDRQ(Me, Absx1.txtFor("CUST_CODE").Text)
        '        With frmQ
        '            .ShowDialog()
        '            rowSOTORDR1.Item("ORDR_STATUS") = "Q"
        '        End With
        '    End Using
        'Else
        '    rowSOTORDR1.Item("ORDR_STATUS") = "L"
        'End If
        If frm.rdoQUOTE.Checked Then
            rowSOTORDR1.Item("ORDR_STATUS") = "Q"
        Else
            rowSOTORDR1.Item("ORDR_STATUS") = "L"
        End If
        rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") = "MK"
        rowSOTORDR1.Item("SALES_DIVISION_CODE") = "RIB"
        rowSOTORDR1.Item("ORDR_DATE_BOOKED") = Now().Date
        rowSOTORDR1.Item("ORDR_PRIORITY") = "1"
        'rowSOTORDR1.Item("SYNC_BATCH") = Null
        rowSOTORDR1.Item("CURR_CODE") = "USD"
        rowSOTORDR1.Item("CURR_EXCH_RATE") = 1
        rowSOTORDR1.Item("ORDR_CUST_PO") = frm.ORDR_CUST_PO
        rowSOTORDR1.Item("ORDR_CATEGORY") = frm.ORDR_CATEGORY
        rowSOTORDR1.Item("ORDR_SHIP_DATE") = frm.STARTDATE
        rowSOTORDR1.Item("ORDR_CANCEL_DATE") = frm.CANCELDATE
        rowSOTORDR1.Item("SHIP_VIA_CODE") = frm.SHIP_VIA_CODE
        rowSOTORDR1.Item("ORDR_SHIP_INSTR") = frm.ORDR_SHIP_INSTR
        rowSOTORDR1.Item("ORDR_MESSAGE") = frm.ORDR_MESSAGE
        rowSOTORDR1.Item("TERM_CODE") = frm.TERM_CODE
        rowSOTORDR1.Item("FRT_TERMS") = frm.FRT_TERMS
        rowSOTORDR1.Item("SREP_CODE") = SREP_CODE
        rowSOTORDR1.Item("WHSE_CODE") = frm.WHSE_CODE
        rowSOTORDR1.Item("ORDR_SHIP_COMPLETE") = rowARTCUST1.Item("CUST_SHIP_COMPLETE") & ""
        rowSOTORDR1.Item("ORDR_NO_WEB") = ""
        rowSOTORDR1.Item("ORDR_FOB") = Get_ORDR_FOB("", frm.WHSE_CODE)
        rowSOTORDR1.Item("POST_CODE") = rowARTCUST1.Item("POST_CODE") & ""
        rowSOTORDR1.Item("CUST_BILL_TO_CUST") = Absx1.txtFor("CUST_CODE").Text
        rowSOTORDR1.Item("ORDR_ORIG_SHIP_DATE") = frm.STARTDATE
        rowSOTORDR1.Item("ORDR_ORIG_CANCEL_DATE") = frm.CANCELDATE
        rowSOTORDR1.Item("ORDR_TYPE_CODE") = "REG"
        rowSOTORDR1.Item("ORDR_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
        If rowSOTORDR1.Item("ORDR_PRIORITY") & "" = "" Then
            rowSOTORDR1.Item("ORDR_PRIORITY") = "9"
        End If
        rowSOTORDR1.Item("ORDR_BUYER_NAME") = frm.ORDR_BUYER_NAME
        rowSOTORDR1.Item("ORDR_BUYER_EMAIL") = frm.ORDR_BUYER_EMAIL
        rowSOTORDR1.Item("ORDR_BUYER_CONTACT_NO") = frm.ORDR_BUYER_CONTACT_NO
        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)
        SetOrderType()
        SetFKeys()
    End Sub

    Private Function Get_ORDR_FOB(ORDR_TYPE_CODE As String, WHSE_CODE As String) As String
        Dim ORDR_FOB As String = ""

        If ORDR_TYPE_CODE = "BTB" Then
            ORDR_FOB = "Port of Origin"
        Else
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            If rowICTWHSE1 IsNot Nothing AndAlso rowICTWHSE1.Item("WHSE_CITY") & "" <> "" Then
                ORDR_FOB = rowICTWHSE1.Item("WHSE_CITY") & "," & rowICTWHSE1.Item("WHSE_STATE")
            Else
                ORDR_FOB = ""
            End If
        End If
        Return ORDR_FOB
    End Function

    Private Sub AddSOTORDR2(ByVal ORDR_NO As String, ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal ORDR_QTY As Integer, ByVal UNIT_PRICE As Double, NetPricing As Boolean)
        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow()
        Dim NextOrderRow As Integer = 0
        For Each rowSOTNEXT As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO), "ORDR_LNO")
            NextOrderRow = rowSOTNEXT.Item("ORDR_LNO")
        Next
        NextOrderRow += 1
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
        rowSOTORDR2.Item("ORDR_LNO") = NextOrderRow
        rowSOTORDR2.Item("STYLE_CODE") = STYLE_CODE
        rowSOTORDR2.Item("COLOR_CODE") = COLOR_CODE
        rowSOTORDR2.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
        rowSOTORDR2.Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
        rowSOTORDR2.Item("STYLE_UOM") = rowICTSTYL1.Item("STYLE_UOM")
        rowSOTORDR2.Item("ORDR_EXTD_COST") = 0
        rowSOTORDR2.Item("ORDR_UNIT_PRICE") = UNIT_PRICE
        rowSOTORDR2.Item("ORDR_QTY") = ORDR_QTY
        rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY
        rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
        rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
        rowSOTORDR2.Item("ORDR_QTY_CANC") = 0
        rowSOTORDR2.Item("ORDR_STATUS") = "L"
        rowSOTORDR2.Item("ORDR_QTY_ORIG") = ORDR_QTY
        rowSOTORDR2.Item("QTY_PER_PP") = 1
        rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = UNIT_PRICE
        rowSOTORDR2.Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
        rowSOTORDR2.Item("ITEM_CODE") = String.Format("{0}{1}", STYLE_CODE, COLOR_CODE)
        If IsNumeric(txtRETAIL_PRICE.Text) Then
            rowSOTORDR2.Item("STYLE_RETAIL") = Val(txtRETAIL_PRICE.Text)
        End If
        If NetPricing Then
            rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL") = "1"
        Else
            rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL") = "0"
        End If
        rowSOTORDR2.Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
        'STYLE_PRICE
        'ORDR_UNIT_PRICE_CALC
        'ORDR_PRICE_SOURCE
        'ORDR_UNIT_PRICE_STD

        'Dim ORDR_UNIT_PRICE_CALC As Double = 0
        'Dim ORDR_PRICE_SOURCE As String = ""
        'ORDR_UNIT_PRICE_CALC = TAC.SOCMAIN1.Price_Line(Me, CUST_CODE, rowARTCUST1, STYLE_CODE, COLOR_CODE, ORDR_QTY, ORDR_PRICE_SOURCE)

        'rowSOTORDR2.Item("SYNC_BATCH") = Null
        dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
    End Sub

    Private Sub AddSOTORDR5(ByVal rowARTCUST2 As DataRow, ByVal ORDR_NO As String)
        Dim INST_COLS As String() = New String() {"CUST_ADDR_CODE", "CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow()
        rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
        rowSOTORDR5.Item("CUST_ADDR_TYPE") = "ST"
        'rowSOTORDR5.Item("SYNC_BATCH") = Null
        For Each COL As String In INST_COLS
            rowSOTORDR5.Item(COL) = rowARTCUST2.Item(COL)
        Next
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
    End Sub

    Private Sub AddSOTORDRB(ByVal ORDR_BATCH_NO As String)
        rowSOTORDRB = dst.Tables("SOTORDRB").NewRow()
        rowSOTORDRB.Item("ORDR_BATCH_NO") = ORDR_BATCH_NO
        rowSOTORDRB.Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE")
        rowSOTORDRB.Item("ORDR_DATE") = DATETIME_STAMP
        rowSOTORDRB.Item("ORDR_CUST_PO") = Absx1.txtFor("ORDR_CUST_PO").Text
        dst.Tables("SOTORDRB").Rows.Add(rowSOTORDRB)
    End Sub

    Private Sub CalculateOrderCuFt(ByVal ORDR_NO As String)
        Dim TCUFT As Double = 0
        Dim TORDR As Double = 0
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTORDR2.Item("STYLE_CODE"))
            Dim CASE_CUBE As Double = 0
            Dim CARTON_PACK_QTY As Integer = 0
            Dim LINE_QTY As Integer = 0
            'Dim LINE_PRICE As Integer = 0
            If Not IsNothing(rowICTSTYL1) Then
                CASE_CUBE = Val(rowICTSTYL1.Item("CASE_CUBE"))
                CARTON_PACK_QTY = Val(rowICTSTYL1.Item("CARTON_PACK_QTY"))
                If CARTON_PACK_QTY = 0 Then
                    CARTON_PACK_QTY = 1
                End If
                LINE_QTY = Val(rowSOTORDR2.Item("ORDR_QTY"))
                'LINE_PRICE = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE"))
                TCUFT += (LINE_QTY / CARTON_PACK_QTY) * CASE_CUBE
                'TORDR += (LINE_QTY / LINE_PRICE)
                rowSOTORDR2.Item("CASE_CUBE") = CASE_CUBE
                rowSOTORDR2.Item("TCUFT") = Format((LINE_QTY / CARTON_PACK_QTY) * CASE_CUBE, "###,##0.00")
            End If
        Next
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            rowSOTORDR1.Item("TCUFT") = Format(TCUFT, "###,##0.00")
            'rowSOTORDR1.Item("TORDR") = Format(TORDR, "###,##0.0")
        Next
    End Sub

    Private Sub CalculateOrderTotal(ByVal ORDR_NO As String)
        Dim ORDR_TOTAL As Double = 0
        Dim FORDR_TOTAL As Double = 0
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            ORDR_TOTAL += (Val(rowSOTORDR2.Item("ORDR_QTY") & "") - Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")) * Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
        Next
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            rowSOTORDR1.Item("TORDR") = Format(ORDR_TOTAL, "###,##0.00")
        Next
        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            rowSOTORDRX.Item("TORDR") = Format(ORDR_TOTAL, "###,##0.00")
        Next
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select()
            FORDR_TOTAL += Val(rowSOTORDR2.Item("ORDR_QTY") & "") * Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
        Next
        txtFTotal.Text = Format(FORDR_TOTAL, "$###,###,##0.00")
    End Sub

    Private Sub CalculateOrderTotalX()
        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select()
            ASCMAIN1.sql = String.Format("select sum(nvl(ordr_unit_price,0) * (nvl(ordr_qty,0) - nvl(ordr_qty_canc,0))) from sotordr2 where ordr_no = '{0}'", rowSOTORDRX.Item("ORDR_NO"))
            rowSOTORDRX.Item("TORDR") = Format(Val(ASCDATA1.GetDataValue), "###,##0.00")
        Next
    End Sub

    Private Function CheckCustSrep() As Boolean
        Dim RetVal As Boolean = True
        ASCMAIN1.sql = String.Format("Select SREP_CODE from ARTCUST1 where CUST_CODE = '{0}'", Absx1.txtFor("CUST_CODE").Text)
        Dim SREP_CODE_CUST As String = ASCDATA1.GetDataValue
        If Not IsNothing(SREP_CODE_CUST) Then
            SREP_CODE = SREP_CODE_CUST
        Else
            RetVal = False
        End If
        'If SREP_CODE = Remote.SREP_CODE Then
        '    RetVal = True
        'End If
        'If Remote.IsUserSuper Then
        '    RetVal = True
        'End If
        Return RetVal
    End Function

    Private Function CheckForSalesHold() As Boolean
        Dim RetVal As Boolean = False
        ASCMAIN1.sql = String.Format("Select MIN(CUST_SALES_HOLD) as CUST_SALES_HOLD from ARTCUST1 where CUST_CODE = '{0}'", Absx1.txtFor("CUST_CODE").Text)
        Dim CUST_SALES_HOLD As String = ASCDATA1.GetDataValue
        If CUST_SALES_HOLD <> "1" Then
            RetVal = True
        End If
        Return RetVal
    End Function

    Private Function CheckMinQty_ORIG(ByVal STYLE_CODE As String) As Boolean
        Stop 'This Process Was Thrown out on 6/1/15 in favor of the new CheckMinQty - W.R.
        Dim RetVal As Boolean = True
        Dim BadColors As String = ""
        Fill_Records("ICTSTYL3", STYLE_CODE, True)
        Dim ExceptionFound As Boolean = False
        For Each rowICTSTYL3 As DataRow In dst.Tables("ICTSTYL3").Select()
            Select Case rowICTSTYL3.Item("ATTR_CODE")
                Case "RIBB", "WREA", "WREAPK"
                    ExceptionFound = True
            End Select
        Next
        If Not ExceptionFound Then
            ASCMAIN1.sql = String.Format("SELECT NVL(INNER_PACK_QTY,1) AS INNER_PACK_QTY FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE)
            Dim INNER_PACK_QTY As Integer = Val(ASCDATA1.GetDataValue)
            If INNER_PACK_QTY > 0 Then
                For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                    If IsNumeric(rowICTSTYC1.Item("ORDR_QTY")) Then
                        If (Val(rowICTSTYC1.Item("ORDR_QTY") > 0)) And (Val(rowICTSTYC1.Item("ORDR_QTY")) < INNER_PACK_QTY) Then
                            BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                        Else
                            If Val(rowICTSTYC1.Item("ORDR_QTY")) > 0 And Val(rowICTSTYC1.Item("ORDR_QTY")) Mod INNER_PACK_QTY <> 0 Then
                                BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                            End If
                        End If
                    End If
                Next
                If BadColors.Length > 0 Then
                    Dim iresult As MsgBoxResult = MsgBox(String.Format("The Following Colors Are Below Minimun Qty Or Not Divisible:{0}{1}Is That OK?", BadColors, vbCrLf), MsgBoxStyle.YesNo, "Minimun Qty")
                    If iresult = MsgBoxResult.No Then
                        RetVal = False
                    End If
                End If
            Else
                ASCMAIN1.sql = String.Format("SELECT NVL(CARTON_PACK_QTY,1) AS CARTON_PACK_QTY FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE)
                Dim CARTON_PACK_QTY As Integer = Val(ASCDATA1.GetDataValue)
                If CARTON_PACK_QTY > 1 Then
                    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                        If IsNumeric(rowICTSTYC1.Item("ORDR_QTY")) Then
                            If (Val(rowICTSTYC1.Item("ORDR_QTY") > 0)) And (Val(rowICTSTYC1.Item("ORDR_QTY")) < CARTON_PACK_QTY) Then
                                BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                            Else
                                If Val(rowICTSTYC1.Item("ORDR_QTY")) > 0 And Val(rowICTSTYC1.Item("ORDR_QTY")) Mod CARTON_PACK_QTY <> 0 Then
                                    BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                                End If
                            End If
                        End If
                    Next
                    If BadColors.Length > 0 Then
                        Dim msg As New System.Text.StringBuilder() With {.Length = 0}
                        msg.AppendLine("You Must Order A Full Carton On This Item Unless")
                        msg.AppendLine("Otherwise Approved By Management. Items Which Can")
                        msg.AppendLine("Be Ordered Below Master Or Below Box Quantity")
                        msg.AppendLine("Are Picks, Ribbon, Expensive Wreaths And Garlands.")
                        msg.AppendLine("Anything Breakable Should Not Be Ordered In Less")
                        msg.AppendLine("Than Master Or Box If The Box Quantity Is Zero.")
                        msg.AppendLine("")
                        msg.AppendLine("OK To Proceed?")
                        Dim iresult As MsgBoxResult = MsgBox(msg.ToString, MsgBoxStyle.YesNo, "Minimum Qty")
                        If iresult = MsgBoxResult.No Then
                            RetVal = False
                        End If
                    End If
                End If
            End If
        End If
        Return RetVal
    End Function

    Private Function CheckASSTQty(ByVal STYLE_CODE As String) As Boolean
        Dim RetVal As Boolean = True
        Dim BadColors As String = ""
        ASCMAIN1.sql = String.Format("SELECT NVL(STYLE_ASST_QTY,1) AS STYLE_ASST_QTY FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE)
        Dim STYLE_ASST_QTY As Integer = Val(ASCDATA1.GetDataValue)
        If STYLE_ASST_QTY > 0 Then
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                If IsNumeric(rowICTSTYC1.Item("ORDR_QTY")) Then
                    If Val(rowICTSTYC1.Item("ORDR_QTY")) > 0 And Val(rowICTSTYC1.Item("ORDR_QTY")) Mod STYLE_ASST_QTY <> 0 Then
                        BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                    End If
                End If
            Next
            If BadColors.Length > 0 Then
                MsgBox(String.Format("The Following Colors Have Assortments That Are Not Divisible:{0}{1}.", BadColors, vbCrLf), MsgBoxStyle.OkOnly, "Assortment Qty")
                RetVal = False
            End If
        End If
        Return RetVal
    End Function

    Private Function CheckMODQty(ByVal STYLE_CODE As String) As Boolean
        Dim RetVal As Boolean = True
        Dim ExceptionFound As Boolean = False

        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("Select Count(*) from ICTSTYL3")
        SQLS.AppendLine(String.Format(" where STYLE_CODE = '{0}' and ATTR_CODE = 'RIBB'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)
        If REC_CNT = 0 Then
            Dim BadColors As String = ""
            ASCMAIN1.sql = String.Format("SELECT NVL(INNER_PACK_QTY,1) AS INNER_PACK_QTY FROM ICTSTYL1 WHERE STYLE_CODE = '{0}'", STYLE_CODE)
            Dim INNER_PACK_QTY As Integer = Val(ASCDATA1.GetDataValue)
            Dim IsRuleRequired As Boolean = False
            If INNER_PACK_QTY > 0 Then
                For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                    If IsNumeric(rowICTSTYC1.Item("ORDR_QTY")) Then
                        If Val(rowICTSTYC1.Item("ORDR_QTY")) > 0 And Val(rowICTSTYC1.Item("ORDR_QTY")) Mod INNER_PACK_QTY <> 0 Then
                            BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                            If IsNumeric(rowICTSTYC1.Item("MSOH").ToString & String.Empty) Then
                                If Val(rowICTSTYC1.Item("MSOH").ToString & String.Empty) <= 0 Then
                                    IsRuleRequired = True
                                Else
                                    If Val(rowICTSTYC1.Item("MSOH").ToString & String.Empty) = Val(rowICTSTYC1.Item("ORDR_QTY")) Then
                                        IsRuleRequired = False
                                    Else
                                        IsRuleRequired = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                If BadColors.Length > 0 Then
                    Dim msb As MsgBoxStyle = MsgBoxStyle.YesNo
                    Dim IsOKMsg As String = "Is That OK?"
                    Dim ExtraInfo As String = ""
                    If IsRuleRequired Then
                        msb = MsgBoxStyle.OkOnly
                        IsOKMsg = vbCrLf & "However, They May Be Ordered By Taking The Full Qty."

                    End If
                    Dim iresult As MsgBoxResult = MsgBox(String.Format("The Following Colors Have Box Qty That Are Not Divisible:{0}{1}{2}", BadColors, vbCrLf, IsOKMsg), msb, "Box Qty")
                    If iresult <> MsgBoxResult.Yes Then
                        RetVal = False
                    End If
                End If
            End If
        End If

        Return RetVal
    End Function

    Private Function CheckMinQty(ByVal STYLE_CODE As String) As Boolean
        Dim RetVal As Boolean = True
        'Special override for Orders to Regency to accomodate Samples per Danny.  WR: 4/15/17
        If Absx1.txtFor("CUST_CODE").Text <> "180000" Then
            Dim BadColors As String = ""
            Dim BelowOrderColors As String = ""
            Dim AllowedOrderColors As String = ""
            Dim STYLE_SO_QTY_MIN As Integer = Val(Absx1.txtFor("STYLE_SO_QTY_MIN").Text)
            If STYLE_SO_QTY_MIN = 0 Then
                STYLE_SO_QTY_MIN = 1
            End If
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                If IsNumeric(rowICTSTYC1.Item("ORDR_QTY")) Then
                    If (Val(rowICTSTYC1.Item("ORDR_QTY") > 0)) And (Val(rowICTSTYC1.Item("ORDR_QTY")) < STYLE_SO_QTY_MIN) Then
                        Dim STYLE_COLOR_STATUS As String = rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString
                        If Val(rowICTSTYC1.Item("MSOH").ToString & "") > Val(rowICTSTYC1.Item("ORDR_QTY").ToString) Then
                            BelowOrderColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                        Else
                            If Val(rowICTSTYC1.Item("MSOH").ToString & "") = Val(rowICTSTYC1.Item("ORDR_QTY").ToString) Then
                                AllowedOrderColors = AllowedOrderColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                            Else
                                BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                            End If
                        End If
                    End If
                End If
            Next
            If BadColors.Length > 0 Then
                MsgBox(String.Format("The Following Colors Are Below Minimum Qty: {0}{1}", BadColors, vbCrLf), MsgBoxStyle.OkOnly, "Minimun Qty")
                RetVal = False
            Else
                If AllowedOrderColors.Length > 0 Then
                    Dim msgrslt As New System.Text.StringBuilder With {.Length = 0}
                    msgrslt.AppendLine("The Following Colors Are Below Minimum Qty.")
                    msgrslt.AppendLine("However, This Order Is Being Fulfilled")
                    msgrslt.AppendLine("Because It Takes The Full Qty Available:" & BelowOrderColors & vbCrLf)
                    MsgBox(msgrslt.ToString(), MsgBoxStyle.OkOnly, "Minimun Qty")
                    RetVal = True
                Else
                    If BelowOrderColors.Length > 0 Then
                        Dim msgrslt As New System.Text.StringBuilder With {.Length = 0}
                        msgrslt.AppendLine("The Following Colors Are Below Minimum Qty.")
                        msgrslt.AppendLine("However, They May Be Ordered By Taking The")
                        msgrslt.AppendLine("Full Qty Available:" & BelowOrderColors & vbCrLf)
                        MsgBox(msgrslt.ToString(), MsgBoxStyle.OkOnly, "Minimun Qty")
                        RetVal = False
                    End If
                End If
            End If
        End If
        Return RetVal
    End Function

    Private Sub CreateSOTORDRB()

    End Sub

    Private Sub EditOrderLine(ByVal ORDER_NO As String, ByVal ORDER_LNO As Integer, ByVal STYLE_CODE As String, Optional FillStyles As Boolean = True)
        Dim Index As Integer = 0
        Dim ColorCode(Index) As String
        Dim ColorQty(Index) As Integer
        Dim ORDR_UNIT_PRICE As Double = 0
        Dim STYLE_RETAIL As Double = 0
        Dim NET_PRICE As Double = 0
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}' AND STYLE_CODE = '{1}'", ORDER_NO, STYLE_CODE))
            ReDim Preserve ColorCode(Index)
            ReDim Preserve ColorQty(Index)
            STYLE_RETAIL = Val(rowSOTORDR2.Item("STYLE_RETAIL") & "")
            ColorCode(Index) = rowSOTORDR2.Item("COLOR_CODE")
            ColorQty(Index) = rowSOTORDR2.Item("ORDR_QTY")
            ORDR_UNIT_PRICE = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
            If rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL").ToString & "" = "1" Then
                NET_PRICE = ORDR_UNIT_PRICE
            End If
            Index += 1
            rowSOTORDR2.Delete()
        Next
        tabORDERS.Tabs(1).Selected = True
        Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
        If FillStyles Then
            FillStyle()
        End If
        SetSelectedFEFDPrice(ORDR_UNIT_PRICE)
        ' Populate All of the Qty for The Colors
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
            For i As Integer = 0 To Index - 1
                If rowICTSTYC1.Item("COLOR_CODE").ToString = ColorCode(i) Then
                    rowICTSTYC1.Item("ORDR_QTY") = ColorQty(i)
                End If
            Next
        Next
        If STYLE_RETAIL <> 0 Then
            txtRETAIL_PRICE.Text = Format(STYLE_RETAIL, "###,##0.00")
        Else
            txtRETAIL_PRICE.Text = Format(STYLE_RETAIL, "###,##0.00")
        End If
        If NET_PRICE <> 0 Then
            txtNET_PRICE.Text = Format(NET_PRICE, "###,##0.00")
        Else
            txtNET_PRICE.Text = Format(NET_PRICE, "###,##0.00")
        End If
    End Sub

    Private Sub EnableForm(ByVal Enabled As Boolean)
        InquiryOnly = Not Enabled
        tabORDERS.Tabs(1).Visible = Enabled
        UltraExplorerBar1.Groups("Screen Control").Items("Record Credit Card").Visible = Enabled
    End Sub

    Private Function ErrIfOrderOnStyle(ByVal STYLE_CODE As String, ByVal ORDR_NO As String) As Boolean
        Dim RetVal As Boolean = False
        Dim Msg As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO), "", DataViewRowState.Unchanged)
            If rowSOTORDR2.Item("STYLE_CODE") = STYLE_CODE Then
                'Msg = String.Format("Style {0} Is Already On An Order {1}", STYLE_CODE, ORDR_NO) _
                '    & vbCrLf & "Please Edit The Line If You Want To Change Or Remove It"
                'MsgBox(Msg, MsgBoxStyle.Critical, "Style Already On Order")
                Dim ORDR_LNO As Integer = rowSOTORDR2.Item("ORDR_LNO").ToString
                EditOrderLine(ORDR_NO, ORDR_LNO, STYLE_CODE, False)
                RetVal = True
                EditOrderLineMode = True
                Exit For
            End If
        Next
        Return RetVal
    End Function

    Private Sub FillStyle()
        If Absx1.txtFor("STYLE_CODE").Text.Length = 0 Then
            Exit Sub
        End If
        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text.ToUpper()
        If VERIFYSTYLE(STYLE_CODE) Then
            EnforceConstraints(False)
            Fill_Records("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text, True)
            If dst.Tables.Item("ICTSTYL1").Rows.Count = 1 Then
                lblEXCLUSIVE_STYLE.Visible = dst.Tables.Item("ICTSTYL1").Rows(0).Item("EXCLUSIVE_STYLE").ToString & "" = "1"
                lblDESIGNER_STYLE.Visible = (dst.Tables.Item("ICTSTYL1").Rows(0).Item("ROYALTY_CODE").ToString & "").Length > 0
            End If
            EnforceConstraints(True)
            FilterColors(Absx1.txtFor("STYLE_CODE").Text)
            If grdICTSTYC1.Rows.Count = 0 Then
                MsgBox("This Style No Longer Has Active Colors With Future Available", MsgBoxStyle.Critical, "Discontinued Colors")
                ClearStyle()
                Exit Sub
            End If
            Bind_Controls(grpSTYL1, "ICTSTYL1")
            UltraExplorerBar1.Groups("Allocation").Visible = True

            'Dim rowARTCUST1 As DataRow
            SetrowARTCUST1()
            Discounts = SOCMAIN2.Price_Discounts(Me, CUST_CODE, rowARTCUST1, Absx1.txtFor("STYLE_CODE").Text, True)
            For i As Integer = 1 To 4
                If Discounts(i - 1).DISCOUNT_QTY = 0 Then
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = False
                    Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = False
                    Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = False
                    Absx1.CtlFor(String.Format("priceDISC{0}", i)).Visible = False
                Else
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = True
                    Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = True
                    Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = True
                    Absx1.CtlFor(String.Format("priceDISC{0}", i)).Visible = True
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Text = Discounts(i - 1).DISCOUNT_DESC
                    Absx1.CtlFor(String.Format("lblDISC{0}", i)).Tag = Discounts(i - 1).DISCOUNT_PCT 'Use for hover over.
                    Absx1.txtFor(String.Format("qtyDISC{0}", i)).Text = Discounts(i - 1).DISCOUNT_QTY
                    Absx1.CtlFor(String.Format("priceDISC{0}", i)).Text = Format$(Discounts(i - 1).DISCOUNT_PRICE, "###,##0.00")
                End If
            Next
            txtFactory.Text = GetVendorData(dst.Tables("ICTSTYL1").Rows(0).Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
            txtPort.Text = GetVendorData(dst.Tables("ICTSTYL1").Rows(0).Item("VEND_CODE").ToString, "PORT_CODE")
            txtVEND_PURCH_COMMENT.Text = GetVendorData(dst.Tables("ICTSTYL1").Rows(0).Item("VEND_CODE").ToString, "VEND_PURCH_COMMENT")
            Dim STYLE_CLASS_CODE As String = dst.Tables("ICTSTYL1").Rows(0).Item("STYLE_CLASS_CODE").ToString
            SetClassTip(STYLE_CLASS_CODE)
            SetStyleColor()
            lblSTATUS.Visible = True
            Sort_grdColumns(grdICTSTYC1, "COLOR_CODE".ToUpper, True)
            grdICTSTYC1.Focus()
            grdICTSTYC1.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            grdICTSTYC1.ActiveCell = grdICTSTYC1.Rows(0).Cells.Item("ORDR_QTY")
            Dim Factor As Double = 1
            If numFEFDFACTOR.Value <> 1 And numFEFDFACTOR.Value <> 0 Then
                Factor = CDbl(numFEFDFACTOR.Value)
            End If
            Dim FEFD As New FEFDPrice(Me, Absx1.txtFor("STYLE_CODE").Text, Factor)
            If FEFD.ErrorMsg = "" Then
                btnFEPrice.Text = Format(FEFD.FEPrice, "###,##0.00")
                btnFEMixPrice.Text = Format(FEFD.FEMixPrice, "###,##0.00")
                btnFDMixPrice.Text = Format(FEFD.FDMixPrice, "###,##0.00")
                btnFDPrice.Text = Format(FEFD.FDPrice, "###,##0.00")
                If FEFDLEVEL.Length > 0 Then
                    For Each PicPrice As String In {"picFEPrice", "picFDPrice", "picFEMixPrice", "picFDMixPrice"}
                        If FEFDLEVEL = PicPrice Then
                            Absx1.CtlFor(PicPrice).Visible = True
                        Else
                            Absx1.CtlFor(PicPrice).Visible = False
                        End If
                    Next
                End If
            Else
                Stop
            End If
            SetColorStatusImages(False, 0)
            ShowPromo(txtSTYLE_CODE.Text)
        Else
            MsgBox("Style Not Found In Masterfile", MsgBoxStyle.Critical, "Invalid Style")
            ClearStyle()
            SetColorStatusImages(False, 0)
            Exit Sub
        End If
    End Sub

    Private Function GetSelectedFEFDPrice() As Double
        Dim RetVal As Double = 0
        For Each PicPrice As String In {"picFEPrice", "picFDPrice", "picFEMixPrice", "picFDMixPrice"}
            If Absx1.CtlFor(PicPrice).Visible Then
                Dim btnName As String = "btn" & PicPrice.Substring(3, PicPrice.Length - 3)
                If IsNumeric(Absx1.CtlFor(btnName).Text) Then
                    RetVal = CDbl(Absx1.CtlFor(btnName).Text)
                End If
            End If
        Next
        Return RetVal
    End Function

    Private Sub RecordCreditCard()
        Dim Cust_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CCOK As Boolean = AllOrdersAreCC()
        If Not CCOK Then
            MsgBox("Not All Orders Have Credit Card Terms", vbOKOnly, "Terms Codes")
            Exit Sub
        End If
        Dim frmSOFORDCC As New SOFORDCC(Me, Cust_CODE)
        With frmSOFORDCC
            .ShowDialog()
        End With
        Dim frmSOFORDRC As New SOFORDRC(Me, Cust_CODE)
        'With frmSOFORDRC
        '    .ShowDialog()
        '    'If .CCProcessed = True Then
        '    '    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
        '    '        rowSOTORDR1.Item("CC_TRANS_ID") = "1"
        '    '    Next
        '    'End If
        'End With
    End Sub

    Private Sub RefreshSOTORDRX()
        Dim SQLW As String = ""
        Dim SQLD As String = ""
        Cursor = Cursors.WaitCursor

        Select Case optShowOrders.Value
            Case "M"
                SQLD = String.Format(" ORDR_DATE >= '{0}'", Format(Now.AddMonths(-1), "dd-MMM-yy"))
            Case "H"
                SQLD = String.Format(" ORDR_DATE >= '{0}'", Format(Now.AddMonths(-6), "dd-MMM-yy"))
            Case "Y"
                SQLD = String.Format(" ORDR_DATE >= '{0}'", Format(Now.AddMonths(-12), "dd-MMM-yy"))
            Case "T"
                SQLD = String.Format(" ORDR_DATE >= '{0}'", Format(Now.AddMonths(-24), "dd-MMM-yy"))
            Case "A"
                Dim iResult As MsgBoxResult
                iResult = MsgBox("The ALL Option May Take A While", vbOKCancel, "Are You Sure?")
                If iResult = vbOK Then
                    SQLD = ""
                    Exit Sub
                Else
                    optShowOrders.Value = "M"
                End If
        End Select
        If chkLapQuote.Checked Then
            If SQLD.Length = 0 Then
                SQLD += " ORDR_STATUS IN ('L','Q')"
            Else
                SQLD += " AND ORDR_STATUS IN ('L','Q')"
            End If
        End If
        If chkCustOverride.Checked = True Then
            Remote.SQLWhere = ""
        End If
        If Remote.SQLWhere.Length > 0 Then
            SQLW = String.Format(" WHERE {0}", Remote.SQLWhere)
            If SQLD <> "" Then
                'All Sales Reps See all orders not transferred now per Mario - W.R 9/9/16
                'SQLW += String.Format(" AND {0} OR ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1_L WHERE SREP_CODE = '{1}')", SQLD, Remote.SREP_CODE)
                SQLW += String.Format(" AND {0} OR ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1_L)", SQLD)
            End If
        Else
            If SQLD <> "" Then
                'All Sales Reps See all orders not transferred now per Mario - W.R 9/9/16
                'SQLW += String.Format(" WHERE {0} OR ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1_L  WHERE SREP_CODE = '{1}')", SQLD, Remote.SREP_CODE)
                SQLW += String.Format(" WHERE {0} OR ORDR_NO IN (SELECT ORDR_NO FROM SOTORDR1_L)", SQLD)
            End If
        End If

        'If chkCustOverride.Checked = True Then
        '    SQLW += " OR (SOTORDR1.SREP_CODE = 'HO' AND SOTORDR1.ORDR_DATE >= '" & Format(Now.AddDays(-4), "dd-MMM-yyyy") & "')"
        'End If


        Application.DoEvents()
        ASCMAIN1.sql = String.Format("SELECT SOTORDR1.*,  TO_CHAR(ORDR_DATE, 'YYYY') AS YEAR FROM SOTORDR1 {0}", SQLW)
        ASCMAIN1.sql = ShowJoinedSreps(ASCMAIN1.sql)
        Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)
        CalculateOrderTotalX()
        Cursor = Cursors.Default
    End Sub

    Private Function RePriceOrder(ByVal ORDR_NO As String) As Double
        Dim RetVal As Double
        Dim ORDR_UNIT_PRICE As Double
        Dim ORDR_UNIT_PRICE_NEW As Double
        Dim ORDR_QTY As Integer
        Dim ORIG_VALUE As Double = 0
        Dim NEW_VALUE As Double = 0
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            If rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL") <> "1" Then
                ORDR_QTY = rowSOTORDR2.Item("ORDR_QTY")
                ORDR_UNIT_PRICE = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                ORIG_VALUE += ORDR_QTY * ORDR_UNIT_PRICE
                Discounts = SOCMAIN2.Price_Discounts(Me, Absx1.txtFor("CUST_CODE").Text, rowARTCUST1, rowSOTORDR2.Item("STYLE_CODE"), True)
                'Begin
                Dim LastGoodBreak As Integer = 0
                Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", rowSOTORDR2.Item("STYLE_CODE"))).FirstOrDefault
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTORDR2.Item("STYLE_CODE"))
                For i As Integer = 0 To 3
                    If Discounts(i).DISCOUNT_QTY > 0 Then
                        LastGoodBreak = i
                    End If
                    If ORDR_QTY >= Discounts(i).DISCOUNT_QTY Then
                        ORDR_UNIT_PRICE_NEW = Discounts(i).DISCOUNT_PRICE
                        Exit For
                    End If
                Next
                If rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString = "PVC" Then
                    If PINNED_CUST_PRICE_TIER_PVC.Length > 0 Then
                        Select Case PINNED_CUST_PRICE_TIER_PVC
                            Case "5C"
                                If ORDR_UNIT_PRICE_NEW > Val(Discounts(1).DISCOUNT_PRICE) Then
                                    If Val(Discounts(1).DISCOUNT_PRICE) < ORDR_UNIT_PRICE_NEW Then
                                        ORDR_UNIT_PRICE_NEW = Val(Discounts(1).DISCOUNT_PRICE)
                                    End If
                                End If
                            Case "FC"
                                If ORDR_UNIT_PRICE_NEW > Val(Discounts(2).DISCOUNT_PRICE) Then
                                    If Val(Discounts(2).DISCOUNT_PRICE) < ORDR_UNIT_PRICE_NEW Then
                                        ORDR_UNIT_PRICE_NEW = Val(Discounts(2).DISCOUNT_PRICE)
                                    End If
                                End If
                        End Select
                    End If
                Else
                    Select Case rowARTCUST1.Item("CUST_PRICE_TIER").ToString
                        Case "HC"
                            If Discounts(2).DISCOUNT_PRICE < ORDR_UNIT_PRICE_NEW Then
                                ORDR_UNIT_PRICE_NEW = Discounts(2).DISCOUNT_PRICE
                            End If
                        Case "FC"
                            If Discounts(1).DISCOUNT_PRICE < ORDR_UNIT_PRICE_NEW Then
                                ORDR_UNIT_PRICE_NEW = Discounts(1).DISCOUNT_PRICE
                            End If
                    End Select
                End If
                If ORDR_UNIT_PRICE_NEW = 0 Then
                    ORDR_UNIT_PRICE_NEW = Discounts(LastGoodBreak).DISCOUNT_PRICE
                End If
                'End
                'ORDR_UNIT_PRICE_NEW = 1
                rowSOTORDR2.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW
                rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_NEW
                NEW_VALUE += ORDR_QTY * ORDR_UNIT_PRICE_NEW
            End If
        Next
        RetVal = NEW_VALUE - ORIG_VALUE
        Return RetVal
    End Function

    Private Function RePriceFEFDOrder(ByVal ORDR_NO As String, ByVal FEFDType As String) As Double
        Dim RetVal As Double
        Dim ORDR_UNIT_PRICE As Double
        Dim ORDR_UNIT_PRICE_NEW As Double
        Dim ORDR_QTY As Integer
        Dim ORIG_VALUE As Double = 0
        Dim NEW_VALUE As Double = 0
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO))
            ORDR_QTY = rowSOTORDR2.Item("ORDR_QTY")
            ORDR_UNIT_PRICE = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
            ORIG_VALUE += ORDR_QTY * ORDR_UNIT_PRICE

            Dim FEFD As New FEFDPrice(Me, rowSOTORDR2.Item("STYLE_CODE").ToString & String.Empty, 1)

            If Not IsNothing(FEFD.ErrorMsg) Then
                ORDR_UNIT_PRICE_NEW = ORDR_UNIT_PRICE
            Else
                Select Case FEFDType
                    Case "FE"
                        ORDR_UNIT_PRICE_NEW = FEFD.FEPrice
                    Case "FEMIX"
                        ORDR_UNIT_PRICE_NEW = FEFD.FEMixPrice
                    Case "FD"
                        ORDR_UNIT_PRICE_NEW = FEFD.FDPrice
                    Case "FDMIX"
                        ORDR_UNIT_PRICE_NEW = FEFD.FDMixPrice
                    Case Else
                        ORDR_UNIT_PRICE_NEW = ORDR_UNIT_PRICE
                End Select
            End If

            rowSOTORDR2.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW
            rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_NEW
            rowSOTORDR2.Item("ORDR_UNIT_PRICE_MANUAL") = "1"
            NEW_VALUE += ORDR_QTY * ORDR_UNIT_PRICE_NEW

        Next
        RetVal = NEW_VALUE - ORIG_VALUE
        Return RetVal
    End Function

    Private Sub SaveDetailsToOrder(ByVal ORDR_NO As String)
        Dim STYLE_ADDED_CNT As Int64 = 0
        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        Dim STYLE_PRICE As Double
        If PINNED_CUST_PRICE_TIER_PVC = "" And PINNED_CUST_PRICE_TIER_PVC <> rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & "" Then
            PINNED_CUST_PRICE_TIER_PVC = rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & ""
            If PINNED_CUST_PRICE_TIER_PVC = "" Then
                PINNED_CUST_PRICE_TIER_PVC = "PC"
            End If
        End If
        If STYLE_CODE.Length > 0 Then
            If CheckDNQTY(STYLE_CODE) Then
                If CheckMinQty(STYLE_CODE) And CheckMODQty(STYLE_CODE) And CheckASSTQty(STYLE_CODE) Then
                    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                        If IsNumeric(rowICTSTYC1.Item("ORDR_QTY")) Then
                            If rowICTSTYC1.Item("ORDR_QTY") > 0 Then
                                Dim NetPricing As Boolean = False
                                Dim LastGoodBreak As Integer = 0
                                Dim SelectedFEFDPrice As Double = GetSelectedFEFDPrice()
                                If SelectedFEFDPrice > 0 Then
                                    If grdSHIP2.ActiveRow.Cells("WHSE_CODE").Text <> "FD" And grdSHIP2.ActiveRow.Cells("WHSE_CODE").Text <> "FE" Then
                                        ASCMAIN1.sql = String.Format("SELECT COUNT(*) AS RECCNT FROM ASTUSER2 WHERE USER_ID = '{0}' AND SECURITY_CODE = 'X6'", ASCMAIN1.USER_ID)
                                        If Val(ASCDATA1.GetDataValue) > 0 Then
                                            Dim Msg As String = "You Selected FE/FD Pricing For The Style"
                                            Msg += vbCrLf & "But The Order You Selected Is Not FE/FD"
                                            Msg += vbCrLf & "Are You Sure That Is What You Want?"
                                            Dim iResult As MsgBoxResult = MsgBox(Msg, vbYesNo, "FE/FD Pricing Issue")
                                            If iResult = MsgBoxResult.No Then
                                                Exit Sub
                                            End If
                                        Else
                                            Dim Msg As String = "You Selected FE/FD Pricing For The Style"
                                            Msg += vbCrLf & "But The Order You Selected Is Not FE/FD"
                                            'Msg += vbCrLf & "Are You Sure That Is What You Want?"
                                            Dim iResult As MsgBoxResult = MsgBox(Msg, vbCritical, "FE/FD Pricing Issue")
                                            'If iResult = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                    End If
                                    If IsNumeric(Absx1.txtFor("CARTON_PACK_QTY").Text) Then
                                        If IsNumeric(rowICTSTYC1.Item("ORDR_QTY").ToString) Then
                                            If Val(rowICTSTYC1.Item("ORDR_QTY").ToString) < Val(Absx1.txtFor("CARTON_PACK_QTY").Text) Or Val(rowICTSTYC1.Item("ORDR_QTY").ToString) Mod Val(Absx1.txtFor("CARTON_PACK_QTY").Text) Then
                                                Dim Msg As String = "You Selected FE/FD Pricing For The Style"
                                                Msg += vbCrLf & "But The Order Qty Is Less Than Carton Qty"
                                                Msg += vbCrLf & "Are You Sure That's OK?"
                                                Dim iResult As MsgBoxResult = MsgBox(Msg, vbYesNo, "FE/FD Qty Issue")
                                                If iResult = MsgBoxResult.No Then
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                    STYLE_PRICE = SelectedFEFDPrice
                                    NetPricing = True
                                Else
                                    For i As Integer = 0 To 3
                                        If Discounts(i).DISCOUNT_QTY > 0 Then
                                            LastGoodBreak = i
                                        End If
                                        If Discounts(i).DISCOUNT_QTY = 0 Then
                                            NetPricing = False
                                            STYLE_PRICE = Val(CDbl(Discounts(LastGoodBreak).DISCOUNT_PRICE))
                                        Else
                                            'If rowICTSTYC1.Item("ORDR_QTY") >= Discounts(i).DISCOUNT_QTY Then
                                            '    If Val(Format$(CDbl(Discounts(i).DISCOUNT_PRICE), "###,##0.00")) <> Val(CDbl(Absx1.CtlFor("priceDISC" & i + 1).Text)) Then
                                            '        NetPricing = True
                                            '        STYLE_PRICE = Val(CDbl(Absx1.CtlFor("priceDISC" & i + 1).Text))
                                            '    Else
                                            '        NetPricing = False
                                            '        STYLE_PRICE = Val(CDbl(Format$(Discounts(i).DISCOUNT_PRICE, "###,##0.00")))
                                            '    End If
                                            '    Exit For
                                            'End If
                                            If txtNET_PRICE.Value > 0 Then
                                                NetPricing = True
                                                STYLE_PRICE = txtNET_PRICE.Value
                                                Exit For
                                            Else
                                                NetPricing = False
                                                STYLE_PRICE = Val(CDbl(Format$(Discounts(i).DISCOUNT_PRICE, "###,##0.00")))
                                                If rowICTSTYC1.Item("ORDR_QTY") >= Discounts(i).DISCOUNT_QTY Then
                                                    STYLE_PRICE = Val(CDbl(Format$(Discounts(i).DISCOUNT_PRICE, "###,##0.00")))
                                                    Exit For
                                                End If

                                            End If
                                        End If
                                    Next
                                    If STYLE_PRICE = 0 Then
                                        STYLE_PRICE = Val(CDbl(Discounts(3).DISCOUNT_PRICE))
                                    End If
                                    If Not NetPricing Then
                                        If rowICTSTYL1.Item("STYLE_CLASS_CODE").ToString = "PVC" Then
                                            If PINNED_CUST_PRICE_TIER_PVC.Length > 0 Then
                                                Select Case PINNED_CUST_PRICE_TIER_PVC
                                                    Case "5C"
                                                        If STYLE_PRICE > Val(CDbl(Absx1.CtlFor("priceDISC2").Text)) Then
                                                            STYLE_PRICE = Val(CDbl(Absx1.CtlFor("priceDISC2").Text))
                                                        End If
                                                    Case "FC"
                                                        If STYLE_PRICE > Val(CDbl(Absx1.CtlFor("priceDISC3").Text)) Then
                                                            STYLE_PRICE = Val(CDbl(Absx1.CtlFor("priceDISC3").Text))
                                                        End If
                                                End Select
                                            End If
                                        Else
                                            Select Case rowARTCUST1.Item("CUST_PRICE_TIER").ToString
                                                Case "HC"
                                                    If STYLE_PRICE > Val(CDbl(Absx1.CtlFor("priceDISC3").Text)) Then
                                                        STYLE_PRICE = Val(CDbl(Absx1.CtlFor("priceDISC3").Text))
                                                    End If
                                                Case "FC"
                                                    If STYLE_PRICE > Val(CDbl(Absx1.CtlFor("priceDISC2").Text)) Then
                                                        STYLE_PRICE = Val(CDbl(Absx1.CtlFor("priceDISC2").Text))
                                                    End If
                                            End Select
                                        End If
                                        If STYLE_PRICE = 0 Then
                                            If Discounts(LastGoodBreak).DISCOUNT_PRICE <> Val(CDbl(Absx1.CtlFor("priceDISC" & LastGoodBreak + 1).Text)) Then
                                                NetPricing = True
                                                STYLE_PRICE = Val(CDbl(Absx1.CtlFor("priceDISC" & LastGoodBreak + 1).Text))
                                            Else
                                                NetPricing = False
                                                STYLE_PRICE = Val(CDbl(Discounts(LastGoodBreak).DISCOUNT_PRICE))
                                            End If

                                        End If
                                    End If
                                End If
                                'Added on 4/25/13 to force Discounts of Discontinued Colors.
                                'Removed per Rich on 6/9/13.
                                'Added Back per Rich on 4/27/14.
                                'Changed again per Danny on 6/18/14 to only force when not Net Pricing.
                                'Both Danny and Rich Thought this didn't happen but it clearly does. 12/26/19
                                If rowICTSTYC1.Item("STYLE_COLOR_STATUS") = "D" And NetPricing = False Then
                                    STYLE_PRICE = Val(CDbl(rowICTSTYL1.Item("STYLE_PRICE")) & "") * (100 - 70) / 100
                                End If
                                AddSOTORDR2(ORDR_NO, rowICTSTYC1.Item("STYLE_CODE").ToString, rowICTSTYC1.Item("COLOR_CODE").ToString, rowICTSTYC1.Item("ORDR_QTY").ToString, STYLE_PRICE, NetPricing)
                                STYLE_ADDED_CNT = STYLE_ADDED_CNT + 1
                            End If
                        End If
                    Next
                    If STYLE_ADDED_CNT = 0 Then
                        MsgBox("No QTY Added To Style.", vbOKOnly, "Warning")
                    End If
                    ClearStyle()
                    CalculateOrderTotal(ORDR_NO)
                    CalculateOrderCuFt(ORDR_NO)
                    grdSOTORDR1.DataBind()
                    'grdSOTORDR1.Rows(0).Selected = True
                    'grdSOTORDR1.Update()
                    'grdSOTORDR2.Update()
                    'grdSOTORDR1.Refresh()
                    'grdSOTORDR2.Refresh()
                    AutoSave()
                End If
            End If
        End If
    End Sub

    Private Function CheckDNQTY(ByVal STYLE_CODE As String) As Boolean
        Dim RetVal As Boolean = True
        Dim BadColors As String = ""
        Dim OK_QUOTE_DNR As Boolean = False
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
            If IsNumeric(rowICTSTYC1.Item("ORDR_QTY")) Then
                If (Val(rowICTSTYC1.Item("ORDR_QTY") > 0)) Then
                    Dim STYLE_COLOR_STATUS As String = rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString
                    If STYLE_COLOR_STATUS = "D" Or STYLE_COLOR_STATUS = "N" Then

                        If Not IsNothing(grdSHIP2.ActiveRow) Then
                            If grdSHIP2.ActiveRow.Cells("WHSE_CODE").Text = "FD" Or grdSHIP2.ActiveRow.Cells("WHSE_CODE").Text = "FE" Then
                                Dim BC As String = ""
                                Dim TOTOH As Integer = Val(rowICTSTYC1.Item("MSOH").ToString & "") + Val(rowICTSTYC1.Item("MSFT").ToString & "")
                                If Val(rowICTSTYC1.Item("ORDR_QTY").ToString) > TOTOH Then
                                    BC = BC & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                                End If
                                Dim msgrslt As New System.Text.StringBuilder With {.Length = 0}
                                msgrslt.AppendLine("The Following Colors Are Discontinued Or")
                                msgrslt.AppendLine("Do Not Reorder And Exceed The Qty Available.")
                                msgrslt.AppendLine($"Because This Order Is For {grdSHIP2.ActiveRow.Cells("WHSE_CODE").Text} You Can Proceed")
                                msgrslt.AppendLine("As Long As You Are Aware.")
                                msgrslt.AppendLine(vbCrLf & BC & vbCrLf & vbCrLf)
                                msgrslt.AppendLine("Proceed?")
                                Dim rslt As MsgBoxResult = MsgBox(msgrslt.ToString(), MsgBoxStyle.YesNo, "Minimun Qty")
                                If rslt = MsgBoxResult.Yes Then
                                    OK_QUOTE_DNR = True
                                End If
                                'This was changed again by Rich to Ask If It Is OK. - 12/9/21
                                'If grdSHIP2.ActiveRow.Cells("ORDR_STATUS").Text = "Q" Then 'This used to be just N but was opened to all by Rich - 7/13/21.
                                '    OK_QUOTE_DNR = True
                                'End If
                            End If
                        End If
                        If Not OK_QUOTE_DNR Then
                            Dim TOTOH As Integer = Val(rowICTSTYC1.Item("MSOH").ToString & "") + Val(rowICTSTYC1.Item("MSFT").ToString & "")
                            If Val(rowICTSTYC1.Item("ORDR_QTY").ToString) > TOTOH Then
                                BadColors = BadColors & vbCrLf & rowICTSTYC1.Item("COLOR_CODE")
                            End If
                        End If
                    End If
                End If
            End If
        Next
        If Not OK_QUOTE_DNR Then
            If BadColors.Length > 0 Then
                Dim msgrslt As New System.Text.StringBuilder With {.Length = 0}
                msgrslt.AppendLine("The Following Colors Are Discontinued Or")
                msgrslt.AppendLine("Do Not Reorder And Exceed The Qty Available.")
                msgrslt.AppendLine("Please Adjust The Order Qty." & vbCrLf & BadColors & vbCrLf)
                MsgBox(msgrslt.ToString(), MsgBoxStyle.OkOnly, "Minimun Qty")
                RetVal = False
            End If
        End If
        Return RetVal
    End Function

    Private Sub SetClassTip(ByVal STYLE_CLASS_CODE As String)
        Dim tt As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo =
                                New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo() _
                                With {.ToolTipText = STYLE_CLASS_CODE}
        tip.SetUltraToolTip(Absx1.txtFor("STYLE_CODE"), tt)
    End Sub

    Private Sub SetFEPics(ByVal PicName As String, Optional ByVal ClearAll As Boolean = False)
        If ClearAll Then
            For Each PicPrice As String In {"picFEPrice", "picFDPrice", "picFEMixPrice", "picFDMixPrice"}
                Dim btnName As String = "btn" & PicPrice.Substring(3, PicPrice.Length - 3)
                Absx1.CtlFor(PicPrice).Visible = False
                Absx1.CtlFor(btnName).Text = "0.00"
            Next
            'FEFDLEVEL = ""
        Else
            If Absx1.CtlFor(PicName).Visible Then
                Absx1.CtlFor(PicName).Visible = False
                FEFDLEVEL = ""
            Else
                Absx1.CtlFor(PicName).Visible = True
                FEFDLEVEL = Absx1.CtlFor(PicName).Name
                For Each PicPrice As String In {"picFEPrice", "picFDPrice", "picFEMixPrice", "picFDMixPrice"}
                    If PicPrice <> Absx1.CtlFor(PicName).Name Then
                        Absx1.CtlFor(PicPrice).Visible = False
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub SetFKeys()
        Dim UsingNumerics As Boolean = False
        Dim LastFKey As Integer = 1
        Dim NRecCnt As Integer = 0
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("", "ORDR_NO")
            If LastFKey > 10 Then
                UsingNumerics = True
                NRecCnt += 1
                rowSOTORDR1.Item("FKEY") = NRecCnt
            Else
                LastFKey += 1
                rowSOTORDR1.Item("FKEY") = String.Format("F{0}", LastFKey)
            End If
        Next
        Sort_grdColumns(grdSHIP2, "ORDR_NO", True)

    End Sub

    Private Sub SetFormToDefault()
        Clear_Record()
        'Fill_Records("SOTORDRX", , True)
        UltraExplorerBar1.Groups("Customer Messages").Visible = False
        UltraExplorerBar1.Groups("FEFD").Visible = True
        UltraExplorerBar1.Groups("FEFD").Expanded = False
        UltraExplorerBar1.Groups("Filter").Visible = True
        UltraExplorerBar1.Groups("Filter").Expanded = True
        OrderHasFEWhse = False
        tabPricing.Tabs(1).Visible = False
        tabORDERS.Tabs.Item(0).Selected = True
        EditOrderLineMode = False
        EditOrderLineWarning = False
    End Sub

    Sub SetOrderType()
        Dim IsOrderType As Boolean = True
        If dst.Tables("SOTORDR1").Rows.Count() > 0 Then
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select().FirstOrDefault()
            If rowSOTORDR1.Item("ORDR_STATUS").ToString = "Q" Then
                IsOrderType = False
            End If
        End If
        'rdoORDER.Checked = IsOrderType
        'rdoQUOTE.Checked = Not IsOrderType
    End Sub

    Public Function SetrowARTCUST1() As Boolean
        Dim retval As Boolean = False
        Dim CUSTMSG As String = ""
        If IsNothing(rowARTCUST1) Then
            rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
            If IsNothing(rowARTCUST1) Then
                retval = False
            Else
                retval = True

                Absx1.txtFor("CUST_NAME").Text = rowARTCUST1.Item("CUST_NAME")
                If rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "FC" Then
                    CUSTMSG = CUSTMSG & "* Full Case Non-PVC." & vbCrLf
                End If
                If rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "HC" Then
                    CUSTMSG = CUSTMSG & "* Half Case Non-PVC." & vbCrLf
                End If
                If rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "SP" Then
                    Dim PCT As String = ""
                    If Val(rowARTCUST1.Item("CUST_DISC_PCT") & "") > 0 Then
                        PCT = Val(rowARTCUST1.Item("CUST_DISC_PCT") & "").ToString & "%"
                    Else
                        If Val(rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") & "") <> 0 Then
                            Select Case Val(rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") & "")
                                Case Is = 1
                                    PCT = "5%"
                                Case Is = 2
                                    PCT = "10%"
                            End Select
                        End If
                    End If
                    CUSTMSG = CUSTMSG & "* " & PCT & " Non-PVC." & vbCrLf
                End If
                Select Case rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & ""
                    Case "5C"
                        CUSTMSG = CUSTMSG & "* 5 Case PVC." & vbCrLf
                    Case "FC"
                        CUSTMSG = CUSTMSG & "* Full Case PVC." & vbCrLf
                    Case Else
                End Select
                If CUSTMSG.Length > 0 Then
                    lblCUSTMSG.Text = CUSTMSG
                    UltraExplorerBar1.Groups("Customer Messages").Visible = True
                Else
                    lblCUSTMSG.Text = ""
                    UltraExplorerBar1.Groups("Customer Messages").Visible = False
                End If

            End If
        Else
            retval = True
        End If
        Return retval
    End Function

    Private Sub SetSelectedFEFDPrice(ByVal ORDR_UNIT_PRICE As Double)
        Dim FEFDPriceFound As Boolean = False
        For Each btnPrice As String In {"btnFEPrice", "btnFDPrice", "btnFEMixPrice", "btnFDMixPrice"}
            If IsNumeric(Absx1.CtlFor(btnPrice).Text) Then
                If Val(Absx1.CtlFor(btnPrice).Text) = ORDR_UNIT_PRICE Then
                    Dim picName As String = "pic" & btnPrice.Substring(3, btnPrice.Length - 3)
                    Absx1.CtlFor(picName).Visible = True
                    FEFDPriceFound = True
                End If
            End If
        Next
        If FEFDPriceFound Then
            tabPricing.Tabs(1).Selected = True
        Else
            tabPricing.Tabs(0).Selected = True
        End If
    End Sub

    Private Sub SetStyleColor()
        If dst.Tables("ICTSTYL1").Rows.Count > 0 Then
            Dim STYLE_STATUS As String = dst.Tables("ICTSTYL1").Rows(0).Item("STYLE_STATUS").ToString()
            Select Case STYLE_STATUS
                Case Is = "D"
                    txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Red
                Case Is = "N"
                    txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Yellow
                Case Else
                    txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Empty
            End Select
        Else
            txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Setup_SOTORDR2()
        If grdSOTORDR1.ActiveRow Is Nothing OrElse (Not grdSOTORDR1.ActiveRow.IsDataRow Or grdSOTORDR1.ActiveRow.IsAddRow) Then
            grdSOTORDR2.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdSOTORDR2.DataSource, DataTable).DefaultView
            Dim ORDR_NO As String = grdSOTORDR1.ActiveRow.Cells("ORDR_NO").Value
            dvw.RowFilter = String.Format("ORDR_NO = '{0}'", ORDR_NO)
            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select(String.Format("ORDR_NO = '{0}'", ORDR_NO)).FirstOrDefault
                            If Not IsNothing(rowSOTORDR5) Then
                                Dim Address As String = String.Format("{0}, {1} {2} {3}", rowSOTORDR5.Item("CUST_ADDR1"), rowSOTORDR5.Item("CUST_CITY"),
                                        rowSOTORDR5.Item("CUST_STATE"), rowSOTORDR5.Item("CUST_ZIP_CODE"))
                                grdSOTORDR2.Text = String.Format("Customer Style / Color Details for Order {0} At {1}", ORDR_NO, Address)
                            Else
                                grdSOTORDR2.Text = String.Format("Customer Style / Color Details for Order {0}", ORDR_NO)
                            End If

                            grdSOTORDR2.Visible = True
                        End If
                        End Sub

    Private Sub ShowClassTip(sender As Object, e As System.EventArgs)
        tip.AutoPopDelay = 3000
        tip.InitialDelay = 3000
        tip.DisplayStyle = ToolTipDisplayStyle.BalloonTip
        tip.ShowToolTip(sender)
    End Sub

    Private Sub ShowStyles()
        If Not InquiryOnly Then
            tabORDERS.Tabs(1).Visible = dst.Tables("SOTORDR1").Rows.Count > 0
            UltraExplorerBar1.Groups("Screen Control").Items("Record Credit Card").Visible = True
        End If
    End Sub

    Private Function VERIFYSTYLE(STYLE_CODE As String) As Boolean
        Dim RETVAL As Boolean = False
        Dim NEW_STYLE As String = ""
        ASCMAIN1.sql = String.Format("SELECT COUNT(*) RECCNT FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
        Dim STYLE_COUNT As Int16 = Val(ASCDATA1.GetDataValue)
        If STYLE_COUNT = 1 Then
            ASCMAIN1.sql = String.Format("SELECT STYLE_CODE FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
            NEW_STYLE = ASCDATA1.GetDataValue
            If Absx1.txtFor("STYLE_CODE").Text <> NEW_STYLE Then
                Absx1.txtFor("STYLE_CODE").Text = NEW_STYLE
            End If
            RETVAL = True
        Else
            MsgBox("Style Not Found In Masterfile", MsgBoxStyle.Critical, "Invalid Style")
            ClearStyle()
            Exit Function
        End If
        Return RETVAL
    End Function

    Private Sub AutoSave()
        Call Update_Record_TDA("SOTORDRB")
        Call Update_Record_TDA("SOTORDR1")
        Call Update_Record_TDA("SOTORDR2")
        Call Update_Record_TDA("SOTORDR5")
    End Sub

    Private Sub btnCustOverride_Click(sender As System.Object, e As System.EventArgs) Handles btnCustOverride.Click
        Dim OverRide As String = SOCMAIN2.TodaysOverRide("C")
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Today's Override"
        Dim iMSG As String = "Today's Override Code" & vbCrLf & "For Customers Is: " & OverRide
        iResult = MsgBox(iMSG, MsgBoxStyle.OkOnly, iTitle)
    End Sub

    Private Sub btnFEOverride_Click(sender As System.Object, e As System.EventArgs) Handles btnFEOverride.Click
        Dim OverRide As String = SOCMAIN2.TodaysOverRide("F")
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Today's Override"
        Dim iMSG As String = "Today's Override Code" & vbCrLf & "For FE/FD Is: " & OverRide
        iResult = MsgBox(iMSG, MsgBoxStyle.OkOnly, iTitle)
    End Sub
#End Region

#Region "These Properties and Methods May Have Aplication In SKU Inquiry As Well"
    Private Sub ShowDisc(sender As Object, e As System.EventArgs)
        Dim this As Infragistics.Win.Misc.UltraLabel = sender
        Dim tt As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo() With {.ToolTipText = this.Tag}
        tip.SetUltraToolTip(sender, tt)
        tip.AutoPopDelay = 3000
        tip.InitialDelay = 3000
        tip.DisplayStyle = ToolTipDisplayStyle.BalloonTip
        tip.ShowToolTip(sender)
    End Sub

    Private Sub FilterColors(STYLE_CODE As String)
        'Dim dvw As DataView = DirectCast(grdICTSTYC1.DataSource, DataTable).DefaultView
        'dvw.RowFilter = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        Dim SUpdate As Boolean = UpdateICTSTAT2(STYLE_CODE, "MS")
        Fill_Records("ICTSTYC1", STYLE_CODE, True)
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select()
            rowICTSTYC1.Item("ORDR_QTY") = 0
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT SUM(NVL(RSRV_QTY_OPEN,0)) RSV")
            SQLS.AppendLine("FROM SOTRSRV2")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", rowICTSTYC1.Item("STYLE_CODE").ToString))
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", rowICTSTYC1.Item("COLOR_CODE").ToString))
            ASCMAIN1.sql = SQLS.ToString()
            Dim RSV As Int16 = Val(ASCDATA1.GetDataValue)
            rowICTSTYC1.Item("RSV") = RSV
        Next
        grdICTSTYC1.UpdateData()
        grdICTSTYC1.Text = String.Format("Colors For Style{0}", STYLE_CODE)
        grdICTSTYC1.Visible = True
        If SUpdate Then
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item(1).Header.Appearance.BackColor = Drawing.Color.Green
        Else
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item(1).Header.Appearance.BackColor = Drawing.Color.Yellow
        End If
    End Sub

    'Private Sub FilterAlloc(STYLE_CODE As String, COLOR_CODE As String, Optional WHSE_CODE As String = "MS")
    '    Dim dvw As DataView = DirectCast(grdICTSTDQ1.DataSource, DataTable).DefaultView
    '    dvw.RowFilter = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND WHSE_CODE = '{2}'", STYLE_CODE, COLOR_CODE, WHSE_CODE)
    '    'grdICTSTDQ1.Text = String.Format("Colors For Style{0}", STYLE_CODE)
    '    Fill_Records("ICTSTDQ1", New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE}, True)
    '    grdICTSTDQ1.Visible = True
    '    grdICTSTDQ1.DisplayLayout.Bands(0).Columns.Item(3).Header.Appearance.BackColor = Drawing.Color.Yellow
    'End Sub

    Private Sub FilterAlloc(STYLE_CODE As String, COLOR_CODE As String, Optional WHSE_CODE As String = "MS")
        'Dim dvw As DataView = DirectCast(grdICTSTDQ1.DataSource, DataTable).DefaultView
        'dvw.RowFilter = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND WHSE_CODE = '{2}'", STYLE_CODE, COLOR_CODE, WHSE_CODE)
        'grdICTSTDQ1.Text = String.Format("Colors For Style{0}", STYLE_CODE)
        Fill_Records("ICTSTDQ1", New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE}, True)
        'Removed 6/3/21 so Sales Reps can now see full allocation - WHR
        'Dim S As New System.Text.StringBuilder() With {.Length = 0}
        'S.AppendLine("SELECT")
        'S.AppendLine("MSOH, MSFT")
        'S.AppendLine("FROM")
        'S.AppendLine("  (")
        'S.AppendLine("   SELECT")
        'S.AppendLine("   C1.STYLE_CODE,")
        'S.AppendLine("   C1.COLOR_CODE,")
        'S.AppendLine("   C1.STYLE_COLOR_STATUS,")
        'S.AppendLine("   CASE WHEN")
        'S.AppendLine("   SUM(")
        'S.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        'S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        'S.AppendLine("     ELSE 0")
        'S.AppendLine("     END) < 0")
        'S.AppendLine("   THEN")
        'S.AppendLine("     0")
        'S.AppendLine("   ELSE")
        'S.AppendLine("   SUM(")
        'S.AppendLine("     CASE S2.WHSE_CODE")
        'S.AppendLine("     WHEN 'MS'")
        'S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        'S.AppendLine("     ELSE 0")
        'S.AppendLine("     END)")
        'S.AppendLine("   END AS MSOH,")
        'S.AppendLine("   CASE WHEN")
        'S.AppendLine("   SUM(")
        'S.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        'S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'S.AppendLine("     ELSE 0")
        'S.AppendLine("     END) <= 0")
        'S.AppendLine("   THEN")
        'S.AppendLine("     0")
        'S.AppendLine("   ELSE")
        'S.AppendLine("     CASE WHEN")
        'S.AppendLine("       SUM(")
        'S.AppendLine("       CASE S2.WHSE_CODE")
        'S.AppendLine("       WHEN 'MS'")
        'S.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'S.AppendLine("       ELSE 0")
        'S.AppendLine("       END) < 0")
        'S.AppendLine("     THEN")
        'S.AppendLine("       0")
        'S.AppendLine("     ELSE")
        'S.AppendLine("     SUM(")
        'S.AppendLine("       CASE S2.WHSE_CODE")
        'S.AppendLine("       WHEN 'MS'")
        'S.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'S.AppendLine("       ELSE 0")
        'S.AppendLine("       END) END")
        'S.AppendLine("   END AS MSFT")
        'S.AppendLine("   FROM ICTSTYC1 C1")
        'S.AppendLine("   LEFT JOIN ICTSTAT2 S2")
        'S.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
        'S.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
        'S.AppendLine("   INNER JOIN ICTCOLR1 C2")
        'S.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
        'S.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
        'S.AppendLine("  )")
        'S.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') OR (MSOH <> 0) OR (MSFT <> 0))")
        'S.AppendLine("  AND STYLE_CODE = :PARM1")
        'S.AppendLine("  AND COLOR_CODE = :PARM2")
        'Dim tbl As DataTable = ASCDATA1.GetDataTable(S.ToString(), String.Empty, "VV", New Object() {STYLE_CODE, COLOR_CODE})
        'For Each rowFUTURE As DataRow In tbl.Rows
        '    If Val(rowFUTURE.Item("MSOH").ToString & "") = 0 Then
        '        For Each rowICTSTDQ1 As DataRow In dst.Tables("ICTSTDQ1").Select("", "STATUS_DATE")
        '            If dst.Tables("ICTSTDQ1").Rows.Count > 1 Then
        '                rowICTSTDQ1.Delete()
        '                dst.Tables("ICTSTDQ1").AcceptChanges()
        '            End If
        '        Next
        '    End If
        'Next
        grdICTSTDQ1.Visible = True
        grdICTSTDQ1.DisplayLayout.Bands(0).Columns.Item(3).Header.Appearance.BackColor = Drawing.Color.Yellow
    End Sub

    Private Sub ClearStyle()
        dst.Tables("ICTSTYL1").Clear()
        ClearQtyOrdered()
        Bind_Controls(grpSTYL1, "ICTSTYL1")
        FilterColors("NONE")
        UltraExplorerBar1.Groups("Image").Visible = False
        UltraExplorerBar1.Groups("Allocation").Visible = False
        For i As Integer = 1 To 4
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Text = ""
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Tag = "Select A Style"
            Absx1.txtFor(String.Format("qtyDISC{0}", i)).Text = ""
            'Absx1.txtFor(String.Format("priceDISC{0}", i)).Text = ""
            Absx1.CtlFor(String.Format("priceDISC{0}", i)).Text = 0
        Next
        txtNET_PRICE.Value = 0
        txtPort.Text = ""
        txtFactory.Text = ""
        txtVEND_PURCH_COMMENT.Text = ""
        txtRETAIL_PRICE.Text = ""
        SetClassTip("")
        SetStyleColor()
        lblSTATUS.Visible = False
        Absx1.txtFor("STYLE_CODE").Focus()
        SetFEPics("", True)
        SetColorStatusImages(0, True)
        lblEXCLUSIVE_STYLE.Visible = False
        lblDESIGNER_STYLE.Visible = False
        lblPromo.Visible = False
        lblPromo.Text = ""
        btnShowPromo.Visible = False
    End Sub

    Private Function VerifyEditMode() As Boolean
        Dim iResult As Boolean = True
        If EditOrderLineMode Then
            Dim iResultm As MsgBoxResult
            Dim iTitle As String = "Finish Edit Mode?"
            Dim iMSG As New System.Text.StringBuilder
            iMSG.AppendLine("This Will Finish The Current Line Edit.")
            iMSG.AppendLine("Is That What You Want?")
            iResultm = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResultm = MsgBoxResult.Yes Then
                EditOrderLineMode = False
                EditOrderLineWarning = False
            Else
                iResult = False
            End If
        End If
        Return iResult
    End Function

    Private Sub ClearQtyOrdered()
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("ORDR_QTY <> 0")
            rowICTSTYC1.Item("ORDR_QTY") = 0
        Next
    End Sub

    Private Function GetImageLocation(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = ""
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim RO_PARM_STYLE_IMG_DIR As String = ""
        Dim FileMatch As String
        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
        Dim COLOR_CODE_LONG As String = ""
        If Not IsNothing(rowICTCOLR1) Then
            COLOR_CODE_LONG = rowICTCOLR1.Item("COLOR_CODE_LONG").ToString()
        End If
        Dim WebVal As String = ""
        If chkLIVEDATA.Checked Then
            'WebVal = TryWebImage(String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE))
        End If
        If Not IsNothing(rowSOTPARM3) Then
            RO_PARM_STYLE_IMG_DIR = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            If RO_PARM_STYLE_IMG_DIR.Length > 0 Then
                FileMatch = Dir(String.Format("{0}\{1}-{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                If FileMatch.Length > 0 Then
                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                Else
                    FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                    If FileMatch.Length > 0 Then
                        RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                    Else
                        FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE_LONG))
                        If FileMatch.Length > 0 Then
                            RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                        Else
                            FileMatch = Dir(String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                            If FileMatch.Length > 0 Then
                                RetVal = String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE)
                            Else
                                FileMatch = Dir(String.Format("{0}\{1}*", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                                If FileMatch.Length > 0 Then
                                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Try
            If WebVal.Length > 0 Then
                Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(WebVal)
                Dim response As System.Net.WebResponse = req.GetResponse()
                Dim stream As IO.Stream = response.GetResponseStream()
                Dim img As System.Drawing.Image = System.Drawing.Image.FromStream(stream)
                stream.Close()
                If System.IO.File.Exists(RetVal) Then
                    System.IO.File.Delete(RetVal)
                    img.Save(RetVal)
                Else
                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE))
                    img.Save(RetVal)
                End If
            End If
        Catch ex As Exception
        End Try
        Return RetVal
    End Function

    Private Function TryWebImage(ImageName As String) As String
        Dim API_BASE As String = ""
        'Dim url As New System.Uri("http://50.75.200.254:8181/images/product/" & ImageName)
        Dim url As New System.Uri("http://api.regency-rib.com:8181/images/product/" & ImageName)
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
        Dim resptest As System.Net.WebResponse
        Dim ErrorsFound As Boolean = False
        Try
            resptest = req.GetResponse()
            'ImageName = "http://50.75.200.254:8181/images/product/" & ImageName
            ImageName = "http://api.regency-rib.com:8181/images/product/" & ImageName
        Catch ex As Exception
            'Try
            '    Dim url2 As New System.Uri("http://192.168.110.224:8181/images/product/" & ImageName)
            '    Dim req2 As System.Net.WebRequest = System.Net.WebRequest.Create(url2)
            '    resptest = req2.GetResponse()
            '    ImageName = "http://192.168.110.224:8181/images/product/" & ImageName
            '    resptest.Close()
            '    req2 = Nothing
            'Catch ex2 As Exception
            ErrorsFound = True
            ImageName = ""
            req = Nothing
            'End Try
        End Try
        Return ImageName
    End Function

    Private Function GetImageFolder() As String
        Dim RetVal As String = ""
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        If Not IsNothing(rowSOTPARM3) Then
            RetVal = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
        End If
        Return RetVal
    End Function

    Private Function GetExcelFolder() As String
        Dim RetVal As String = ASCMAIN1.Folders("Work")
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        If Not IsNothing(rowSOTPARM3) Then
            RetVal = rowSOTPARM3.Item("RO_PARM_EXCEL_DIR").ToString
        End If
        If RetVal.Length > 0 Then
            If RetVal.Substring(RetVal.Length - 1, 1) <> "\" Then
                RetVal = RetVal & "\"
            End If
        End If
        Return RetVal
    End Function

    Private Sub imgSTYL1_DoubleClick(sender As Object, e As System.EventArgs) Handles imgSTYL1.DoubleClick
        'Dim frmSOFIMGV1 As New SOFIMGV1(Me, imgSTYL1.ImageLocation)
        'frmSOFIMGV1.Show()
        Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
        Dim frmIMAGE As New TAC.TAFIMGV1(Me, STYLE_CODE, COLOR_CODE, "L")
        With frmIMAGE
            .ShowDialog(Me)
        End With
    End Sub

    Private Function GetVendorData(ByVal VEND_CODE As String, ByVal COLUMN As String) As String
        Dim RetVal As String = ""
        If VEND_CODE.Length > 0 And COLUMN.Length > 0 Then
            ASCMAIN1.sql = String.Format("SELECT {0} FROM APTVEND1 WHERE VEND_CODE = '{1}'", COLUMN, VEND_CODE)
            RetVal = ASCDATA1.GetDataValue
        End If
        Return RetVal
    End Function

    Private Sub FORM_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.NumPad1 Or e.KeyCode = Keys.D1) And e.Alt Then
            Call Click_Command("New Scan (Alt-1)", e)
            Exit Sub
        End If
        If e.KeyData = Keys.Enter Then
            If Absx1.txtFor("STYLE_CODE").Focused And Absx1.txtFor("STYLE_CODE").Text.Length > 0 Then
                FillStyle()
                Exit Sub
            End If
        End If
        Dim FKEY As String = ""
        Select Case e.KeyCode
            Case Keys.F2
                FKEY = "F2"
            Case Keys.F3
                FKEY = "F3"
            Case Keys.F4
                FKEY = "F4"
            Case Keys.F5
                FKEY = "F5"
            Case Keys.F6
                FKEY = "F6"
            Case Keys.F7
                FKEY = "F7"
            Case Keys.F8
                FKEY = "F8"
            Case Keys.F9
                FKEY = "F9"
            Case Keys.F10
                FKEY = "F10"
            Case Keys.F11
                FKEY = "F11"
            Case Keys.F12
                FKEY = "F12"
        End Select
        grdICTSTYC1.UpdateData()
        If FKEY.Length > 0 Then
            grdSHIP2.Selected.Rows.Clear()
            For Each rowCurrent As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSHIP2.Rows
                If rowCurrent.GetCellValue("FKEY") = FKEY Then
                    Dim ORDR_NO As String = rowCurrent.Cells("ORDR_NO").Text
                    If Not ErrIfOrderOnStyle(Absx1.txtFor("STYLE_CODE").Text, ORDR_NO) Then
                        If Not HasDiscColors(Absx1.txtFor("STYLE_CODE").Text) Then
                            rowCurrent.Activate()
                            SaveDetailsToOrder(ORDR_NO)
                            EditOrderLineMode = False
                            EditOrderLineWarning = False
                        End If
                    End If
                End If
            Next

        End If
    End Sub

    Private Sub PrintHangTag()
        If Absx1.txtFor("STYLE_CODE").Text.Length > 0 Then
            ASCMAIN1.sql = "select ro_parm_lbl_printer from sotparm3 where ro_parm_key = 'Z'"
            Dim Printer_name As String = ASCDATA1.GetDataValue
            If Printer_name.Length = 0 Then
                MsgBox("Printer Not Defined In Parameters", vbOKOnly, "Printer Definition")
            Else
                Dim HANGTAG As New HANGTAG(Me, Absx1.txtFor("STYLE_CODE").Text, Discounts, Printer_name)
                If HANGTAG.ErrMsg.Length = 0 Then
                    HANGTAG.Print()
                End If
            End If
        End If
    End Sub

#Region "Hover Over Functionality"
    Private Sub lblDISC1_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC1.MouseHover
        ShowDisc(sender, e)
    End Sub

    Private Sub lblDISC2_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC2.MouseHover
        ShowDisc(sender, e)
    End Sub

    Private Sub lblDISC3_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC3.MouseHover
        ShowDisc(sender, e)
    End Sub

    Private Sub lblDISC4_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC4.MouseHover
        ShowDisc(sender, e)
    End Sub
#End Region
#End Region

#Region "Excel Functionality"
    Private Sub ExcelProcessInit()
        Try
            'Get all currently running process Ids for Excel applications
            mExcelProcesses = Process.GetProcessesByName("Excel")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ExcelProcessKill()
        Dim oProcesses() As Process
        Dim bFound As Boolean

        Try
            'Get all currently running process Ids for Excel applications
            oProcesses = Process.GetProcessesByName("Excel")

            If oProcesses.Length > 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    bFound = False

                    For j As Integer = 0 To mExcelProcesses.Length - 1
                        If oProcesses(i).Id = mExcelProcesses(j).Id Then
                            bFound = True
                            Exit For
                        End If
                    Next

                    If Not bFound Then
                        oProcesses(i).Kill()
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Function Generate_Excel(ByVal ORDR_NO As String) As String
        Dim excelFile As String = ""
        Dim FILE_NAME As String = ""
        Dim xlPages As New Dictionary(Of Integer, Integer)
        Dim PrntQtyOrdered As Boolean
        Dim SelMin As Boolean = False
        Dim PrintRemarks As Boolean
        Dim excel As Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Dim XWB As Excel.Workbook = excel.Workbooks.Add
        Dim oSheet As Excel.Worksheet = XWB.Sheets(1)
        Dim Iresponse As MsgBoxResult
        Dim ShowCancelledLines As Boolean = True

        Fill_Records("SOTORDX1", ORDR_NO, True)
        Fill_Records("SOTORDX2", ORDR_NO, True)
        Fill_Records("SOTORDX5", ORDR_NO, True)

        For Each rowSOTORDX2 As DataRow In dst.Tables("SOTORDX2").Select()
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTORDX2.Item("STYLE_CODE"))
            rowSOTORDX2.Item("FACTORY_CODE") = GetVendorData(rowICTSTYL1.Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
        Next

        If IMAGES_FOLDER.Length = 0 Then
            MsgBox("Location For Images Not Set Up." & vbCrLf & "Please Set It Up In Parameters Before Proceeding", MsgBoxStyle.OkOnly, "Image Location")
            Return ""
            Exit Function
        End If

        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Sort Order"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("Do You Want To Sort By Style Code?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            ExcelSort = "STYLE_CODE, COLOR_CODE"
        Else
            ExcelSort = "ORDR_LNO"
        End If
        iMSG.Length = 0
        iMSG.AppendLine("Do you want the pictures to be factory sorted?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            PictureSort = FactorySort
        Else
            PictureSort = ExcelSort
        End If

        Dim CancelCount As Integer = dst.Tables.Item("SOTORDX2").Select("ORDR_QTY_CANC > 0").Count
        If CancelCount > 0 Then
            If Not MsgBox("Show Cancelled Lines On Order?", vbYesNo, "Cancelled Lines") = vbYes Then
                ShowCancelledLines = False
            End If
        End If

        Dim rowSOTORDX1 As DataRow = dst.Tables("SOTORDX1").Rows(0)
        If rowSOTORDX1.Item("ORDR_STATUS") <> "L" Then
            Iresponse = MsgBox("Do You Want The Entered Order Qty To Appear?", vbYesNo, "Confirm Qty Printing")
            If Iresponse = vbYes Then
                PrntQtyOrdered = True
                SelMin = False
            Else
                PrntQtyOrdered = False
                Iresponse = MsgBox("Do You Want To Use The Order Qty As The Min Order Qty", vbYesNo, "Min Order Qty")
                If Iresponse = vbYes Then
                    SelMin = True
                Else
                    SelMin = False
                End If
            End If
        Else
            PrntQtyOrdered = True
            SelMin = False
        End If
        'Iresponse = MsgBox("Do You Want To Print Remarks?", vbYesNo, "Remarks")
        'Always print remarks now per Danny.
        PrintRemarks = MsgBoxResult.Yes

        If XWB.Worksheets.Count < 3 Then
            For i As Integer = XWB.Worksheets.Count To 3
                XWB.Worksheets.Add()
            Next
        End If
        XWB.Worksheets(1).Name = "SKU"
        XWB.Worksheets(2).Name = "Factory"
        XWB.Worksheets(3).Name = "Pictures"

        For i As Integer = 1 To 2
            oSheet = XWB.Worksheets(i)
            Excel_Format_Heads(oSheet, PrintRemarks, SelMin)
            Excel_Fill_Sheet(oSheet, PrintRemarks, PrntQtyOrdered, SelMin, ShowCancelledLines)
            Excel_Make_Totals(oSheet)
            Excel_Auto_Resize(oSheet)
        Next
        oSheet = XWB.Worksheets(1)
        For i As Integer = 18 To 14 Step -1
            oSheet.Columns(i).Delete()
        Next i

        'Fill In Pictures Tab
        oSheet = XWB.Worksheets(3)


        Excel_Fill_Pictures(oSheet)

        '---- End of config
        'excel.PrintCommunication = False
        With oSheet.PageSetup
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .CenterHorizontally = True
        End With
        'excel.PrintCommunication = True

        'oSheet.PageSetup.FitToPagesTall = False
        'oSheet.PageSetup.FitToPagesWide = 1


        oSheet.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlPageBreakPreview

        oSheet.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlNormalView

        Dim xlsFileName_sfx As String = ""
        Dim xlsFileName As String = ""
        Dim xlsControlNo As String = ASCMAIN1.Next_Control_No("QUOTE")
        FILE_NAME = rowSOTORDX1.Item("CUST_NAME").ToString & "-" & rowSOTORDX1.Item("ORDR_CUST_PO").ToString & "-" & Format(rowSOTORDX1.Item("ORDR_DATE"), "yy-MM-dd") & "-" & rowSOTORDX1.Item("ORDR_NO").ToString
        FILE_NAME = FILE_NAME.Replace("'", "")
        FILE_NAME = FILE_NAME.Replace(" ", "")
        FILE_NAME = FILE_NAME.Replace("/", "")
        FILE_NAME = FILE_NAME.Replace(".", "")
        FILE_NAME = FILE_NAME.Replace("&", "")
        FILE_NAME = FILE_NAME.Replace("$", "")
        FILE_NAME = FILE_NAME.Replace("@", "")
        FILE_NAME = FILE_NAME.Replace("!", "")
        FILE_NAME = FILE_NAME.Replace("*", "")
        FILE_NAME = FILE_NAME.Replace("(", "")
        FILE_NAME = FILE_NAME.Replace(")", "")
        FILE_NAME = FILE_NAME.Replace("#", "")
        FILE_NAME = InputBox("File Name:", "Save File As", FILE_NAME)
        Do
            Try
                xlsFileName = FILE_NAME
                If xlsFileName_sfx.Length = 0 Then
                    excelFile = String.Format("{0}{1}.xls", EXCELDir, xlsFileName)
                Else
                    excelFile = String.Format("{0}{1}_{2}.xls", EXCELDir, xlsFileName, xlsFileName_sfx)
                End If
                XWB.SaveAs(excelFile)
                xlsFileName_sfx = ""
            Catch ex As Exception
                xlsFileName_sfx = CStr(Val(xlsFileName_sfx) + 1)
            End Try
        Loop While xlsFileName_sfx <> "" And Val(xlsFileName_sfx) < 20

        XWB.Close()
        XWB = Nothing
        excel = Nothing
        Return xlsFileName
    End Function

    Private Sub Excel_Format_Heads(ByRef oSheet As Excel.Worksheet, ByVal PrintRemarks As Boolean, ByVal SelMin As Boolean)
        Dim rowSOTORDX1 As DataRow = dst.Tables("SOTORDX1").Rows(0)
        Dim rowSOTORDX5 As DataRow = dst.Tables("SOTORDX5").Rows(0)
        'Dim rng As Excel.Range
        If PrintRemarks Then
            EndMark = 20
        Else
            EndMark = 19
        End If
        'rng = oSheet.Range(Excel_Cell(1, 1), Excel_Cell(1, 3))
        oSheet.Range(Excel_Cell(1, 1), Excel_Cell(1, 3)).Merge()
        oSheet.Range(Excel_Cell(2, 1), Excel_Cell(2, 3)).Merge()
        oSheet.Range(Excel_Cell(3, 1), Excel_Cell(3, 3)).Merge()
        oSheet.Range(Excel_Cell(4, 1), Excel_Cell(4, 3)).Merge()
        oSheet.Range(Excel_Cell(5, 1), Excel_Cell(5, 3)).Merge()

        oSheet.Range(Excel_Cell(1, 1), Excel_Cell(1, 1)).Font.Bold = True
        oSheet.Range(Excel_Cell(1, 1), Excel_Cell(1, 1)).Value = "Regency International"
        'oSheet.Range(Excel_Cell(2, 1)).Value = "11 East 26th Street"
        'oSheet.Range(Excel_Cell(3, 1), Excel_Cell(3, 1)).Value = "New York, NY 10010"
        oSheet.Range(Excel_Cell(2, 1), Excel_Cell(2, 1)).Value = "P.(212) 947-7500  F.(212) 685-2062"
        oSheet.Range(Excel_Cell(3, 1), Excel_Cell(3, 1)).Value = String.Format("For: {0}", rowSOTORDX1.Item("CUST_NAME"))
        oSheet.Range(Excel_Cell(4, 1), Excel_Cell(4, 1)).Value = String.Format("     {0}", rowSOTORDX5.Item("CUST_ADDR1"))
        oSheet.Range(Excel_Cell(5, 1), Excel_Cell(5, 1)).Value = String.Format("     {0}, {1} {2}", rowSOTORDX5.Item("CUST_CITY"), rowSOTORDX5.Item("CUST_STATE"), rowSOTORDX5.Item("CUST_ZIP_CODE"))
        With oSheet.Range(Excel_Cell(1, 1), Excel_Cell(2, 3))
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .Interior.ColorIndex = 33
            .BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic)
        End With
        With oSheet.Range(Excel_Cell(3, 1), Excel_Cell(5, 3))
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .Interior.ColorIndex = 15
            .BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic)
        End With

        'BEGIN - New Extended Header information added 6/27/07
        oSheet.Range(Excel_Cell(1, 4), Excel_Cell(1, 7)).Merge()
        oSheet.Range(Excel_Cell(1, 8), Excel_Cell(1, 12)).Merge()
        oSheet.Range(Excel_Cell(2, 4), Excel_Cell(2, 7)).Merge()
        oSheet.Range(Excel_Cell(2, 8), Excel_Cell(2, 12)).Merge()
        oSheet.Range(Excel_Cell(3, 4), Excel_Cell(3, 7)).Merge()
        oSheet.Range(Excel_Cell(3, 8), Excel_Cell(3, 12)).Merge()
        oSheet.Range(Excel_Cell(4, 4), Excel_Cell(4, 7)).Merge()
        oSheet.Range(Excel_Cell(4, 8), Excel_Cell(4, 12)).Merge()
        oSheet.Range(Excel_Cell(5, 4), Excel_Cell(5, 7)).Merge()
        oSheet.Range(Excel_Cell(5, 8), Excel_Cell(5, 12)).Merge()
        With oSheet.Range(Excel_Cell(1, 4), Excel_Cell(5, 7))
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .Interior.ColorIndex = 20
            .BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic)
        End With
        With oSheet.Range(Excel_Cell(1, 8), Excel_Cell(5, 12))
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            .Interior.ColorIndex = 20
            .BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Excel.XlColorIndex.xlColorIndexAutomatic)
        End With
        oSheet.Range(Excel_Cell(1, 4), Excel_Cell(1, 4)).Value = "Buyer: "
        oSheet.Range(Excel_Cell(2, 4), Excel_Cell(2, 4)).Value = "SP: " & rowSOTORDX1.Item("SREP_CODE")
        oSheet.Range(Excel_Cell(3, 4), Excel_Cell(3, 4)).Value = "Ship Via: " & rowSOTORDX1.Item("SHIP_VIA_CODE")
        Dim ORDR_CUST_PO As String = rowSOTORDX1.Item("ORDR_CUST_PO") & ""
        Dim ORDR_CATEGORY As String = rowSOTORDX1.Item("ORDR_CATEGORY") & ""
        If ORDR_CUST_PO.Length = 0 Then
            ORDR_CUST_PO = "N/A"
        End If
        If ORDR_CATEGORY.Length = 0 Then
            ORDR_CATEGORY = "N/A"
        End If
        oSheet.Range(Excel_Cell(4, 4), Excel_Cell(4, 4)).Value = "PO: " & ORDR_CUST_PO & " | Theme: " & ORDR_CATEGORY
        oSheet.Range(Excel_Cell(5, 4), Excel_Cell(5, 4)).Value = String.Format("Terms: {0} as of {1}", rowSOTORDX1.Item("TERM_CODE"), rowSOTORDX1.Item("ORDR_SHIP_DATE"))
        If IsDate(rowSOTORDX1.Item("ORDR_DATE")) Then
            oSheet.Range(Excel_Cell(1, 8), Excel_Cell(1, 8)).Value = "Date: " & rowSOTORDX1.Item("ORDR_DATE")
        Else
            oSheet.Range(Excel_Cell(1, 8), Excel_Cell(1, 8)).Value = "Date: "
        End If
        If IsDate(rowSOTORDX1.Item("ORDR_SHIP_DATE")) Then
            oSheet.Range(Excel_Cell(2, 8), Excel_Cell(2, 8)).Value = "Required: " & rowSOTORDX1.Item("ORDR_SHIP_DATE")
        Else
            oSheet.Range(Excel_Cell(2, 8), Excel_Cell(2, 8)).Value = "Required: "
        End If
        If IsDate(rowSOTORDX1.Item("ORDR_CANCEL_DATE")) Then
            oSheet.Range(Excel_Cell(3, 8), Excel_Cell(3, 8)).Value = "Cancel: " & rowSOTORDX1.Item("ORDR_CANCEL_DATE")
        Else
            oSheet.Range(Excel_Cell(3, 8), Excel_Cell(3, 8)).Value = "Cancel: "
        End If
        oSheet.Range(Excel_Cell(4, 8), Excel_Cell(4, 8)).Value = "Whse: " & rowSOTORDX1.Item("WHSE_CODE")
        oSheet.Range(Excel_Cell(5, 8), Excel_Cell(5, 8)).Value = "Order/Quote: " & rowSOTORDX1.Item("ORDR_NO")
        'END - New Extended Header information added 6/27/07

        oSheet.Range(Excel_Cell(SCD - 1, 1), Excel_Cell(SCD - 1, 1)).Value = "Stock#"
        oSheet.Range(Excel_Cell(SCD - 1, 2), Excel_Cell(SCD - 1, 2)).Value = "Description"
        oSheet.Range(Excel_Cell(SCD - 1, 3), Excel_Cell(SCD - 1, 3)).Value = "Color"
        oSheet.Range(Excel_Cell(SCD - 1, 4), Excel_Cell(SCD - 1, 4)).Value = "UPC"
        oSheet.Range(Excel_Cell(SCD - 1, 5), Excel_Cell(SCD - 1, 5)).Value = "Box"
        oSheet.Range(Excel_Cell(SCD - 1, 6), Excel_Cell(SCD - 1, 6)).Value = "Cart"
        oSheet.Range(Excel_Cell(SCD - 1, 7), Excel_Cell(SCD - 1, 7)).Value = "U/M"
        oSheet.Range(Excel_Cell(SCD - 1, 8), Excel_Cell(SCD - 1, 8)).Value = "CuFt"
        oSheet.Range(Excel_Cell(SCD - 1, 9), Excel_Cell(SCD - 1, 9)).Value = "List Price"
        oSheet.Range(Excel_Cell(SCD - 1, 10), Excel_Cell(SCD - 1, 10)).Value = "Net Cost"
        oSheet.Range(Excel_Cell(SCD - 1, 10), Excel_Cell(SCD - 1, 10)).Interior.ColorIndex = 3
        oSheet.Range(Excel_Cell(SCD - 1, 11), Excel_Cell(SCD - 1, 11)).Value = "U/M"
        oSheet.Range(Excel_Cell(SCD - 1, 12), Excel_Cell(SCD - 1, 12)).Value = "Retail"
        oSheet.Range(Excel_Cell(SCD - 1, 13), Excel_Cell(SCD - 1, 13)).Value = "Qty Ordered"
        oSheet.Range(Excel_Cell(SCD - 1, 13), Excel_Cell(SCD - 1, 13)).Orientation = 90
        oSheet.Range(Excel_Cell(1, 13), Excel_Cell(SCD - 1, 13)).Merge()
        oSheet.Range(Excel_Cell(1, 13), Excel_Cell(1, 13)).Font.Bold = True
        oSheet.Range(Excel_Cell(SCD - 1, 14), Excel_Cell(SCD - 1, 14)).Value = "Cube"
        oSheet.Range(Excel_Cell(SCD - 1, 15), Excel_Cell(SCD - 1, 15)).Value = "Factory Code"
        oSheet.Range(Excel_Cell(SCD - 1, 15), Excel_Cell(SCD - 1, 15)).Orientation = 90
        oSheet.Range(Excel_Cell(1, 15), Excel_Cell(SCD - 1, 15)).Merge()
        oSheet.Range(Excel_Cell(1, 15), Excel_Cell(1, 15)).Font.Bold = True
        oSheet.Range(Excel_Cell(SCD - 2, 16), Excel_Cell(SCD - 2, 18)).Merge()
        If SelMin Then
            'oSheet.Range(Excel_Cell(SCD - 1, 16), Excel_Cell(SCD - 1, 16)).Value = ""
            oSheet.Range(Excel_Cell(SCD - 1, 16), Excel_Cell(SCD - 1, 16)).Value = "MOQ"
        Else
            'oSheet.Range(Excel_Cell(SCD - 1, 16), Excel_Cell(SCD - 1, 16)).Value = ""
            oSheet.Range(Excel_Cell(SCD - 1, 16), Excel_Cell(SCD - 1, 16)).Value = "Order Qty"
        End If
        oSheet.Range(Excel_Cell(SCD - 2, 16), Excel_Cell(SCD - 2, 16)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        oSheet.Range(Excel_Cell(SCD - 2, 16), Excel_Cell(SCD - 2, 16)).Interior.ColorIndex = 33
        oSheet.Range(Excel_Cell(SCD - 1, 16), Excel_Cell(SCD - 1, 16)).Interior.ColorIndex = 8
        oSheet.Range(Excel_Cell(SCD - 1, 17), Excel_Cell(SCD - 1, 17)).Value = "Cube"
        oSheet.Range(Excel_Cell(SCD - 2, 17), Excel_Cell(SCD - 2, 17)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        oSheet.Range(Excel_Cell(SCD - 2, 17), Excel_Cell(SCD - 2, 17)).Interior.ColorIndex = 33
        oSheet.Range(Excel_Cell(SCD - 1, 17), Excel_Cell(SCD - 1, 17)).Interior.ColorIndex = 8

        oSheet.Range(Excel_Cell(SCD - 1, 18), Excel_Cell(SCD - 1, 18)).Value = "Port"
        oSheet.Range(Excel_Cell(SCD - 1, 18), Excel_Cell(SCD - 1, 18)).Interior.ColorIndex = 8
        oSheet.Range(Excel_Cell(SCD - 1, 19), Excel_Cell(SCD - 1, 19)).Value = "C.O."
        oSheet.Range(Excel_Cell(SCD - 1, 20), Excel_Cell(SCD - 1, 20)).Value = "Total"
        If PrintRemarks Then
            oSheet.Range(Excel_Cell(SCD - 1, 21), Excel_Cell(SCD - 1, 21)).Value = "Remarks"
        End If
        For i As Integer = 1 To EndMark
            oSheet.Range(Excel_Cell(SCD - 1, i), Excel_Cell(SCD - 1, i)).Font.Bold = True
            oSheet.Range(Excel_Cell(SCD - 1, i), Excel_Cell(SCD - 1, i)).Borders.Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            oSheet.Range(Excel_Cell(SCD - 1, i), Excel_Cell(SCD - 1, i)).Font.Bold = True
            oSheet.Cells(i).Columns.AutoFit()
        Next i
    End Sub

    Private Sub Excel_Fill_Sheet(ByRef oSheet As Excel.Worksheet,
                                 ByVal PrintRemarks As Boolean,
                                 ByVal PrntQtyOrdered As Boolean,
                                 ByVal SelMin As Boolean,
                                 ByVal ShowCancelledLines As Boolean)
        RowCount = 0
        Dim FCount As Integer
        Dim LastSKU As String = ""
        Dim Sort As String = ""
        Select Case oSheet.Index
            Case 1
                Sort = ExcelSort
            Case 2
                Sort = FactorySort
            Case 3
                Sort = PictureSort
        End Select
        'If oSheet.Index = 2 Then
        '    Sort = ""
        'Else
        '    Sort = ExcelSort
        'End If
        Dim filter As String = ""
        If ShowCancelledLines = False Then
            filter = "ORDR_QTY_OPEN > 0"
        End If
        For Each rowSOTORDX2 As DataRow In dst.Tables("SOTORDX2").Select(filter, Sort)
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTORDX2.Item("STYLE_CODE"))
            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE")})
            If Not IsNothing(rowICTSTYL1) Then
                oSheet.Range(Excel_Cell(SCD + RowCount, 1), Excel_Cell(SCD + RowCount, 1)).Value = rowSOTORDX2.Item("STYLE_CODE")
                oSheet.Range(Excel_Cell(SCD + RowCount, 2), Excel_Cell(SCD + RowCount, 2)).Value = rowSOTORDX2.Item("STYLE_DESC")
                oSheet.Range(Excel_Cell(SCD + RowCount, 3), Excel_Cell(SCD + RowCount, 3)).Value = rowSOTORDX2.Item("COLOR_CODE")
                oSheet.Range(Excel_Cell(SCD + RowCount, 4), Excel_Cell(SCD + RowCount, 4)).NumberFormat = "########################"
                oSheet.Range(Excel_Cell(SCD + RowCount, 4), Excel_Cell(SCD + RowCount, 4)).Value = rowICTSTYC1.Item("UPC_CODE")
                Dim UOMnum As Integer
                Select Case rowICTSTYL1.Item("STYLE_UOM")
                    Case Is = "DZ"
                        UOMnum = 12
                    Case Is = "GR"
                        UOMnum = 144
                    Case Else
                        UOMnum = 1
                End Select
                'Editing from here
                If Not IsDBNull(rowICTSTYL1.Item("INNER_PACK_QTY")) Then
                    oSheet.Range(Excel_Cell(SCD + RowCount, 5), Excel_Cell(SCD + RowCount, 5)).Value = Val(rowICTSTYL1.Item("INNER_PACK_QTY")) / Val(UOMnum)
                Else
                    oSheet.Range(Excel_Cell(SCD + RowCount, 5), Excel_Cell(SCD + RowCount, 5)).Value = 0
                End If
                oSheet.Range(Excel_Cell(SCD + RowCount, 6), Excel_Cell(SCD + RowCount, 6)).Value = Val(rowICTSTYL1.Item("CARTON_PACK_QTY")) / Val(UOMnum)
                oSheet.Range(Excel_Cell(SCD + RowCount, 7), Excel_Cell(SCD + RowCount, 7)).Value = rowICTSTYL1.Item("STYLE_UOM")
                'Always Print Cube now for every Line Per Rich.
                'If LastSKU <> rowSOTORDX2.Item("STYLE_CODE") Then
                '    oSheet.Range(Excel_Cell(SCD + RowCount, 8), Excel_Cell(SCD + RowCount, 8)).Value = rowICTSTYL1.Item("CASE_CUBE")
                'Else
                '    oSheet.Range(Excel_Cell(SCD + RowCount, 8), Excel_Cell(SCD + RowCount, 8)).Value = ""
                'End If
                oSheet.Range(Excel_Cell(SCD + RowCount, 8), Excel_Cell(SCD + RowCount, 8)).Value = rowICTSTYL1.Item("CASE_CUBE")
                'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                '    If rowSOTORDX2.Item("STYLE_CODE") = "MTX63427" Then Stop 'Use This to Check Style price
                'End If
                'Changed to use Order List Price per Danny. - W.R. 1/26/23
                'oSheet.Range(Excel_Cell(SCD + RowCount, 9), Excel_Cell(SCD + RowCount, 9)).Value = rowICTSTYL1.Item("STYLE_PRICE")
                'oSheet.Range(Excel_Cell(SCD + RowCount, 9), Excel_Cell(SCD + RowCount, 9)).Value = rowSOTORDX2.Item("STYLE_PRICE")
                'Changes to by Hybrid per Rich - W.R> 6/7/23.
                Dim useMaster As Boolean = True
                If IsNumeric(rowSOTORDX2.Item("STYLE_PRICE")) Then
                    If Val(rowSOTORDX2.Item("STYLE_PRICE").ToString & String.Empty) > 0 Then
                        useMaster = False
                    End If
                End If
                If useMaster Then
                    oSheet.Range(Excel_Cell(SCD + RowCount, 9), Excel_Cell(SCD + RowCount, 9)).Value = rowICTSTYL1.Item("STYLE_PRICE")
                Else
                    oSheet.Range(Excel_Cell(SCD + RowCount, 9), Excel_Cell(SCD + RowCount, 9)).Value = rowSOTORDX2.Item("STYLE_PRICE")
                End If

                oSheet.Range(Excel_Cell(SCD + RowCount, 9), Excel_Cell(SCD + RowCount, 9)).NumberFormat = "$###,##0.00"
                If Val(rowSOTORDX2.Item("ORDR_QTY")) <> 0 Then
                    oSheet.Range(Excel_Cell(SCD + RowCount, 10), Excel_Cell(SCD + RowCount, 10)).Value = rowSOTORDX2.Item("ORDR_UNIT_PRICE")
                Else
                    oSheet.Range(Excel_Cell(SCD + RowCount, 10), Excel_Cell(SCD + RowCount, 10)).Value = 0
                End If
                oSheet.Range(Excel_Cell(SCD + RowCount, 10), Excel_Cell(SCD + RowCount, 10)).NumberFormat = "###,##0.00"
                oSheet.Range(Excel_Cell(SCD + RowCount, 10), Excel_Cell(SCD + RowCount, 10)).Font.ColorIndex = 32
                oSheet.Range(Excel_Cell(SCD + RowCount, 10), Excel_Cell(SCD + RowCount, 10)).Font.Bold = True
                oSheet.Range(Excel_Cell(SCD + RowCount, 11), Excel_Cell(SCD + RowCount, 11)).Value = rowICTSTYL1.Item("STYLE_UOM")
                oSheet.Range(Excel_Cell(SCD + RowCount, 12), Excel_Cell(SCD + RowCount, 12)).Value = rowSOTORDX2.Item("STYLE_RETAIL")
                oSheet.Range(Excel_Cell(SCD + RowCount, 12), Excel_Cell(SCD + RowCount, 12)).NumberFormat = "###,##0.00"
                If PrntQtyOrdered Then
                    oSheet.Range(Excel_Cell(SCD + RowCount, 13), Excel_Cell(SCD + RowCount, 13)).Value = rowSOTORDX2.Item("ORDR_QTY")
                Else
                    oSheet.Range(Excel_Cell(SCD + RowCount, 13), Excel_Cell(SCD + RowCount, 13)).Value = ""
                End If
                oSheet.Range(Excel_Cell(SCD + RowCount, 13), Excel_Cell(SCD + RowCount, 13)).NumberFormat = "###,##0"
                If LastSKU <> rowSOTORDX2.Item("STYLE_CODE") Then
                    FCount = RowCount
                End If
                Dim Formula As String = "=IF(" & Excel_Cell(SCD + RowCount, 13) & "<>" & Chr(34) & Chr(34) & ",(" & Excel_Cell(SCD + RowCount, 13) & "/" & Excel_Cell(SCD + FCount, 6) & ")*" & Excel_Cell(SCD + FCount, 8) & "," & Chr(34) & Chr(34) & ")"
                oSheet.Range(Excel_Cell(SCD + RowCount, 14), Excel_Cell(SCD + RowCount, 14)).Formula = Formula
                oSheet.Range(Excel_Cell(SCD + RowCount, 14), Excel_Cell(SCD + RowCount, 14)).NumberFormat = "###,##0.00"
                oSheet.Range(Excel_Cell(SCD + RowCount, 15), Excel_Cell(SCD + RowCount, 15)).Value = rowSOTORDX2.Item("FACTORY_CODE") & "" 'GetVendorData(rowICTSTYL1.Item("VEND_CODE"), "VEND_SUPPLIER_ID")
                oSheet.Range(Excel_Cell(SCD + RowCount, 15), Excel_Cell(SCD + RowCount, 15)).Font.Color = System.Drawing.Color.Red
                oSheet.Range(Excel_Cell(SCD + RowCount, 15), Excel_Cell(SCD + RowCount, 15)).Font.Bold = vbTrue
                If SelMin Then
                    oSheet.Range(Excel_Cell(SCD + RowCount, 16), Excel_Cell(SCD + RowCount, 16)).Value = rowSOTORDX2.Item("ORDR_QTY") & ""
                Else
                    oSheet.Range(Excel_Cell(SCD + RowCount, 16), Excel_Cell(SCD + RowCount, 16)).Value = ""
                End If
                If LastSKU <> rowSOTORDX2.Item("STYLE_CODE") Then
                    FCount = RowCount
                End If
                Formula = "=IF(" & Excel_Cell(SCD + RowCount, 16) & "<>" & Chr(34) & Chr(34) & ",(" & Excel_Cell(SCD + RowCount, 16) & "/" & Excel_Cell(SCD + FCount, 6) & ")*" & Excel_Cell(SCD + FCount, 8) & "," & Chr(34) & Chr(34) & ")"
                oSheet.Range(Excel_Cell(SCD + RowCount, 17), Excel_Cell(SCD + RowCount, 17)).Formula = Formula
                Dim Port As String = ""
                oSheet.Range(Excel_Cell(SCD + RowCount, 18), Excel_Cell(SCD + RowCount, 18)).Formula = GetVendorData(rowICTSTYL1.Item("VEND_CODE") & "", "PORT_CODE")
                oSheet.Range(Excel_Cell(SCD + RowCount, 18), Excel_Cell(SCD + RowCount, 18)).NumberFormat = "###,##0.00"
                oSheet.Range(Excel_Cell(SCD + RowCount, 19), Excel_Cell(SCD + RowCount, 19)).Value = rowICTSTYL1.Item("COUNTRY_CODE") & ""
                oSheet.Range(Excel_Cell(SCD + RowCount, 20), Excel_Cell(SCD + RowCount, 20)).Formula = "=J" & Format(SCD + RowCount, "000") & "* M" & Format(SCD + RowCount, "000")
                oSheet.Range(Excel_Cell(SCD + RowCount, 20), Excel_Cell(SCD + RowCount, 20)).NumberFormat = "$##,###,##0.00"
                If PrintRemarks = True Then
                    oSheet.Range(Excel_Cell(SCD + RowCount, 21), Excel_Cell(SCD + RowCount, 21)).Value = "" 'We still Don't know where this comes from. rsItemMaster.Fields("Remark").Value
                End If
                LastSKU = rowSOTORDX2.Item("STYLE_CODE")
                RowCount = RowCount + 1
            End If
        Next
    End Sub

    Private Sub Excel_Make_Totals(ByRef oSheet As Excel.Worksheet)
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 2), Excel_Cell(SCD + RowCount + 1, 2)).Value = "TOTALS"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 13), Excel_Cell(SCD + RowCount + 1, 13)).Formula = "=SUM(M001:M" & Format(SCD + RowCount, "000") & ")"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 13), Excel_Cell(SCD + RowCount + 1, 13)).NumberFormat = "###,##0.00"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 14), Excel_Cell(SCD + RowCount + 1, 14)).Formula = "=SUM(N001:N" & Format(SCD + RowCount, "000") & ")"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 14), Excel_Cell(SCD + RowCount + 1, 14)).NumberFormat = "###,##0.00"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 16), Excel_Cell(SCD + RowCount + 1, 16)).Formula = "=SUM(P001:P" & Format(SCD + RowCount, "000") & ")"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 16), Excel_Cell(SCD + RowCount + 1, 16)).NumberFormat = "###,##0.00"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 17), Excel_Cell(SCD + RowCount + 1, 17)).Formula = "=SUM(Q001:Q" & Format(SCD + RowCount, "000") & ")"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 17), Excel_Cell(SCD + RowCount + 1, 17)).NumberFormat = "###,##0.00"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 20), Excel_Cell(SCD + RowCount + 1, 20)).Formula = "=SUM(T001:T" & Format(SCD + RowCount, "000") & ")"
        oSheet.Range(Excel_Cell(SCD + RowCount + 1, 20), Excel_Cell(SCD + RowCount + 1, 20)).NumberFormat = "$##,###,##0.00"
        For i As Integer = 1 To EndMark
            oSheet.Range(Excel_Cell(SCD + RowCount + 1, i), Excel_Cell(SCD + RowCount + 1, i)).Font.Bold = True
            oSheet.Range(Excel_Cell(SCD + RowCount + 1, i), Excel_Cell(SCD + RowCount + 1, i)).Borders.Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlDouble
            oSheet.Range(Excel_Cell(SCD + RowCount + 1, i), Excel_Cell(SCD + RowCount + 1, i)).Borders.Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            With oSheet.Range(Excel_Cell(SCD, i), Excel_Cell(SCD + RowCount, i))
                oSheet.Range(Excel_Cell(SCD + RowCount + 1, i), Excel_Cell(SCD + RowCount + 1, i)).BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
            End With
        Next i
        'Begin - tariff Notification
        'Removed 10:30AM on a dark and rainy day April 16th 2020
        'Dim LastRow As Int64 = SCD + RowCount + 2
        'With oSheet.Range(Excel_Cell(LastRow, 1), Excel_Cell(LastRow, EndMark))
        '    .Merge()
        '    .Value = "Your order/quote includes trade tariffs. If these tariffs are not in place when Regency is receiving your goods, we shall recalculate and resend your sales confirmation with the correct lower costs for each item.  Depending on the size of the item and the duty paragraph it lies in, you may expect a reduction anywhere from 7-10%."
        '    .Font.Bold = True
        '    .Font.Color = Color.Red
        '    .RowHeight = 45
        '    .WrapText = True
        '    .VerticalAlignment = Excel.XlVAlign.xlVAlignTop
        '    .BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
        'End With
        'End - tariff Notification
    End Sub

    Private Sub Excel_Auto_Resize(oSheet As Excel.Worksheet)
        oSheet.Cells.Columns.AutoFit()
    End Sub

    Private Sub InsertPictureInRange(ByVal PictureFileName As String,
                ByVal TargetCells As Microsoft.Office.Interop.Excel.Range,
                ByVal XWS As Microsoft.Office.Interop.Excel.Worksheet)
        ASCMAIN1.Progress("Picture:" & PictureFileName)
        ' inserts a picture and resizes it to fit the TargetCells range
        Dim pp As Microsoft.Office.Interop.Excel.Shape

        If TypeName(XWS) <> "Worksheet" Then Exit Sub
        If Dir(PictureFileName) = "" Then Exit Sub

        pp = XWS.Shapes.AddPicture(PictureFileName,
           Microsoft.Office.Core.MsoTriState.msoFalse,
           Microsoft.Office.Core.MsoTriState.msoCTrue, TargetCells.Left, TargetCells.Top, TargetCells.Width, TargetCells.Height)

        'pp = XWS.Shapes.AddPicture(PictureFileName, _
        '   0, _
        '   1, TargetCells.Left, TargetCells.Top, TargetCells.Width, TargetCells.Height)
        pp.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize
        pp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse
        pp = Nothing
    End Sub

    Private Function SecondPictureCaption(ByVal rowSOTORDX2 As DataRow) As String
        Dim RetVal As String = ""
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", rowSOTORDX2.Item("STYLE_CODE"))
        If Not IsNothing(rowICTSTYL1) Then
            RetVal = "Price:" & rowSOTORDX2.Item("ORDR_UNIT_PRICE") & " |Cart: " & rowICTSTYL1.Item("CARTON_PACK_QTY") & " |Box: " & rowICTSTYL1.Item("INNER_PACK_QTY") & " |Cube: " & rowICTSTYL1.Item("CASE_CUBE")
        Else
            RetVal = "Price:" & rowSOTORDX2.Item("ORDR_UNIT_PRICE") & " |Style Masterfile Problem"
        End If
        Return RetVal
    End Function

    Private Sub Excel_Fill_Pictures(oSheet As Excel.Worksheet)
        Dim IsEven As Boolean = True
        Dim PageRow As Integer = 0
        Dim PicRange As Excel.Range
        Dim CapRange As Excel.Range
        Dim CapRange2 As Excel.Range
        Dim HBreaks As New List(Of Excel.Range)
        Dim PicNOF() As String = Nothing
        RowCount = 1
        oSheet.Range("A1", "L1").ColumnWidth = 7.0
        oSheet.Range("A1", "L10000").RowHeight = 12.75
        Dim RowHeight As Integer = 18

        For Each rowSOTORDX2 As DataRow In dst.Tables("SOTORDX2").Select("", PictureSort)
            If IsEven Then
                PicRange = oSheet.Range("A" & Format(RowCount, "0000") & ":F" & Format(RowCount + RowHeight - 3, "0000"))
                CapRange = oSheet.Range("A" & Format(RowCount + RowHeight - 2, "0000") & ":F" & Format(RowCount + RowHeight - 2, "0000"))
                CapRange2 = oSheet.Range("A" & Format(RowCount + RowHeight - 1, "0000") & ":F" & Format(RowCount + RowHeight - 1, "0000"))
                IsEven = False
            Else
                PicRange = oSheet.Range("G" & Format(RowCount, "0000") & ":L" & Format(RowCount + RowHeight - 3, "0000"))
                CapRange = oSheet.Range("G" & Format(RowCount + RowHeight - 2, "0000") & ":L" & Format(RowCount + RowHeight - 2, "0000"))
                CapRange2 = oSheet.Range("G" & Format(RowCount + RowHeight - 1, "0000") & ":L" & Format(RowCount + RowHeight - 1, "0000"))
                HBreaks.Add(oSheet.Range("A" & Format(RowCount + RowHeight, "0000") & ":L" & Format(RowCount + RowHeight, "0000")))
                IsEven = True
                RowCount += RowHeight
                PageRow += 1
            End If

            Dim PictureFileName As String = GetImageLocation(rowSOTORDX2.Item("STYLE_CODE"), rowSOTORDX2.Item("COLOR_CODE"))
            If PictureFileName.Length = 0 Then
                If IsNothing(PicNOF) Then
                    ReDim PicNOF(0)
                    PicNOF(PicNOF.Length - 1) = rowSOTORDX2.Item("STYLE_CODE")
                Else
                    If Not PicNOF.Contains(rowSOTORDX2.Item("STYLE_CODE")) Then
                        ReDim PicNOF(PicNOF.Length + 1)
                        PicNOF(PicNOF.Length - 1) = rowSOTORDX2.Item("STYLE_CODE")
                    End If
                End If
            Else
                InsertPictureInRange(PictureFileName, PicRange, oSheet)
                ASCMAIN1.Progress("Now Generating Excel Document")
                CapRange.Merge()
                CapRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                CapRange.Value = rowSOTORDX2.Item("STYLE_CODE") & "-" & rowSOTORDX2.Item("COLOR_CODE") & " " & rowSOTORDX2.Item("STYLE_DESC")
                'CapRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                CapRange2.Merge()
                CapRange2.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic)
                CapRange2.Value = SecondPictureCaption(rowSOTORDX2)
                CapRange2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            End If
        Next
        oSheet.PageSetup.PrintArea = "A1:L" & RowCount
        'oSheet.PageSetup.FitToPagesTal
        oSheet.ResetAllPageBreaks()
        Dim RowsOnPage As Integer = RowHeight * 3
        For i As Integer = RowsOnPage To RowCount Step RowsOnPage
            oSheet.HPageBreaks.Add(oSheet.Range("A" & i + 1, "L" & i + 1))
        Next
        'oSheet.VPageBreaks.Add(oSheet.Range("M1:M1"))
        'For Each rng As Excel.Range In HBreaks
        '    oSheet.HPageBreaks.Add(rng)
        'Next
        'oSheet.Range("A1", "L1").ColumnWidth = 6.6
        'oSheet.Range("A1", "L10000").RowHeight = 14
    End Sub
#End Region

#Region "Remote Service Calls"
    Function UpdateICTSTAT2(ByVal STYLE_CODE As String, ByVal WHSE_CODE As String) As Boolean
        Dim iresult As Boolean = False
        If STYLE_CODE = "NONE" Or chkLIVEDATA.Checked = False Then
            Return iresult
            Exit Function
        Else
            Dim API_BASE As String = ""
            'Dim url As New System.Uri("http://50.75.200.254:8181/")
            Dim url As New System.Uri("http://api.regency-rib.com:8181/")
            Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
            Dim resptest As System.Net.WebResponse
            req.Timeout = 2000
            Try
                resptest = req.GetResponse()
                resptest.Close()
                req = Nothing
                'API_BASE = "http://50.75.200.254:8181/"
                API_BASE = "http://api.regency-rib.com:8181/"
            Catch ex As Exception
                req = Nothing
                'API_BASE = "http://192.168.110.224:8181/"
                API_BASE = "http://api.regency-rib.com:8181/"
            End Try

            Dim API_CONTROLLER As String = "api/SalesOrder/GetICTSTAT2"
            Dim API_QUERY_STRING As String = String.Format("?STYLE_CODE={0}&COLOR_CODE=&WHSE_CODE={1}&CSQL=", STYLE_CODE, WHSE_CODE)

            Dim client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(2)
            client.BaseAddress = New Uri(API_BASE)

            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            Dim response As HttpResponseMessage
            Try
                response = client.GetAsync(API_CONTROLLER & API_QUERY_STRING).Result
            Catch ex As Exception
                Return iresult
                Exit Function
            End Try

            If response.IsSuccessStatusCode Then
                Try
                    Dim ICTSTAT2_RESPONSE As New List(Of ICTSTAT2)
                    Dim apiResponseString As String = ""
                    Dim responseObject As Object = response.Content.ReadAsAsync(Of IEnumerable(Of ICTSTAT2))().Result

                    ICTSTAT2_RESPONSE = responseObject
                    'Stop
                    For Each STAT As ICTSTAT2 In ICTSTAT2_RESPONSE
                        Fill_Record("ICTSTAT2", New String() {STAT.m_STYLE_CODE, STAT.m_COLOR_CODE, STAT.m_WHSE_CODE}, False, True)
                        If dst.Tables.Item("ICTSTAT2").Rows.Count > 0 Then
                            'dst.Tables.Item("ICTSTAT2").Rows(0).Item("INIT_DATE") = STAT.m_INIT_DATE
                            'dst.Tables.Item("ICTSTAT2").Rows(0).Item("LAST_DATE") = STAT.m_LAST_DATE
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_ALLO") = STAT.m_WHSE_QTY_ALLO
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_COMM") = STAT.m_WHSE_QTY_COMM
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_ON_HAND") = STAT.m_WHSE_QTY_ON_HAND
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_ON_ORDER") = STAT.m_WHSE_QTY_ON_ORDER
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_OPEN") = STAT.m_WHSE_QTY_OPEN
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_PICK") = STAT.m_WHSE_QTY_PICK
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_PROD") = STAT.m_WHSE_QTY_PROD
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_TRAN") = STAT.m_WHSE_QTY_TRAN
                        End If
                        Call Update_Record_TDA("ICTSTAT2")
                    Next
                    iresult = True
                Catch ex As Exception
                    iresult = False
                End Try

            Else
                iresult = False
            End If
        End If
        Return iresult
    End Function


#End Region

    Private Sub UpdateAddresses()
        For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select()
            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, "MK", rowSOTORDR5.Item("CUST_ADDR_CODE")})
            If Not IsNothing(rowARTCUST2) Then
                rowSOTORDR5.Item("CUST_NAME") = rowARTCUST2.Item("CUST_NAME")
                rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST2.Item("CUST_ADDR1")
                rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST2.Item("CUST_ADDR2")
                rowSOTORDR5.Item("CUST_CITY") = rowARTCUST2.Item("CUST_CITY")
                rowSOTORDR5.Item("CUST_STATE") = rowARTCUST2.Item("CUST_STATE")
                rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST2.Item("CUST_ZIP_CODE")
            End If
        Next
    End Sub

    Private Function setSREP_CODE() As String
        'Try to get the Sales Rep Code From Customer.   If you can't Find it at least use HO so it's not null - WR - 6/27/13
        Dim RetVal As String = "HO"
        ASCMAIN1.sql = String.Format("Select SREP_CODE from ARTCUST1 where CUST_CODE = '{0}'", Absx1.txtFor("CUST_CODE").Text)
        Dim SREP_CODE_CUST As String = ASCDATA1.GetDataValue
        If Not IsNothing(SREP_CODE_CUST) Then
            RetVal = SREP_CODE_CUST
        End If
        Return RetVal
    End Function

    Private Function AllOrdersAreCC(Optional ByVal AlsoCheckCCPA As Boolean = False) As Boolean
        Dim iResult As Boolean = False
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            Select Case rowSOTORDR1.Item("TERM_CODE")
                Case "AMEX", "DISC", "MC", "VISA", "CRED"
                    If AlsoCheckCCPA Then
                        If rowSOTORDR1.Item("CCPA_NO") & "" <> "" Then
                            iResult = True
                        End If
                    Else
                        iResult = True
                    End If
                Case Else
                    iResult = True
            End Select
        Next
        Return iResult
    End Function

    Private Function CCPAOnOrders(Optional ByVal AlsoCheckCCPA As Boolean = False) As Boolean
        Dim iResult As Boolean = False
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
            Select Case rowSOTORDR1.Item("TERM_CODE")
                Case "AMEX", "DISC", "MC", "VISA", "CRED"
                    'iResult = True
                Case Else
                    If rowSOTORDR1.Item("CCPA_NO") & "" <> "" Then
                        iResult = True
                    End If
            End Select
        Next
        Return iResult
    End Function

    Private Sub imgSTYL1_Click(sender As System.Object, e As System.EventArgs) Handles imgSTYL1.Click

    End Sub

    Private Function HasDiscColors(STYLE_CODE As String) As Boolean
        Dim RetVal As Boolean = False
        If grdSHIP2.ActiveRow.Cells("WHSE_CODE").Text <> "FD" And grdSHIP2.ActiveRow.Cells("WHSE_CODE").Text <> "FE" Then
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                If Val(rowICTSTYC1.Item("ORDR_QTY").ToString & "") <> 0 Then
                    'If rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString & "" <> "A" And rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString & "" <> "N" The
                    If rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString & "" <> "A" Then
                        Dim TOTOH As Integer = Val(rowICTSTYC1.Item("MSOH").ToString & "") + Val(rowICTSTYC1.Item("MSFT").ToString & "")
                        If TOTOH = 0 Then
                            If (ASCMAIN1.USER_ID = "rich" Or ASCMAIN1.USER_ID = "tonyg" Or ASCMAIN1.USER_ID = "james") Then
                                MsgBox("You Are Adding A Discontinued Item With No Qty To An Order/Quote.", MsgBoxStyle.Information, "Discontinued Color")
                                RetVal = False
                            Else
                                MsgBox("You May Not Add A Discontinued Item With No Qty To An Order.", MsgBoxStyle.Information, "Discontinued Color")
                                RetVal = True
                            End If
                            Return RetVal
                        End If
                    End If
                End If
            Next
        End If
        Return RetVal
    End Function

    Private Sub txtNET_PRICE_Click(sender As Object, e As System.EventArgs) Handles txtNET_PRICE.Click
        txtNET_PRICE.SelectAll()
    End Sub

    Private Function ShowJoinedSreps(sql As String) As String
        Dim RetVal As String = sql
        If sql.Contains("SREP_CODE = 'DE'") Then
            RetVal = sql.Replace("SREP_CODE = 'DE'", "(SREP_CODE = 'DE' OR SREP_CODE = 'HR')")
        End If
        Return RetVal
    End Function

    Private Sub TABLETIZE_ORDER(ByVal ORDR_NO As String)
        Call BeginTrans()

        Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
        SQLS.AppendLine(String.Format("DELETE FROM SOTORDRT_L WHERE ORDR_NO = '{0}'", ORDR_NO))
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        SQLS.Length = 0
        SQLS.AppendLine("INSERT INTO SOTORDRT_L")
        SQLS.AppendLine("VALUES (")
        SQLS.AppendLine("'" & ORDR_NO & "',")
        SQLS.AppendLine("0,")
        SQLS.AppendLine("NULL,")
        SQLS.AppendLine("0,")
        SQLS.AppendLine("0,")
        SQLS.AppendLine("NULL,")
        SQLS.AppendLine("'" & ASCMAIN1.USER_ID & "',")
        SQLS.AppendLine("'" & ASCMAIN1.USER_ID & "',")
        SQLS.AppendLine("SYSDATE,")
        SQLS.AppendLine("'" & ASCMAIN1.USER_ID & "',")
        SQLS.AppendLine("SYSDATE")
        SQLS.AppendLine(")")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        For Each TableName As String In New String() {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
            SQLS.Length = 0
            SQLS.AppendLine("DELETE")
            SQLS.AppendLine("FROM " & TableName & "_L")
            SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            SQLS.Length = 0
            SQLS.AppendLine("INSERT INTO " & TableName & "_L")
            SQLS.AppendLine("SELECT * FROM " & TableName)
            SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            SQLS.Length = 0
            SQLS.AppendLine("DELETE")
            SQLS.AppendLine("FROM " & TableName)
            SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
        Next

        SQLS.Length = 0
        SQLS.AppendLine("UPDATE SOTORDR1_L")
        SQLS.AppendLine("SET ORDR_SOURCE = 'T',")
        SQLS.AppendLine("ORDR_STATUS = 'W',")
        SQLS.AppendLine("ORDR_GROUP_NO = ORDR_BATCH_NO")
        SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        SQLS.Length = 0
        SQLS.AppendLine("UPDATE SOTORDR1_L")
        SQLS.AppendLine("SET ORDR_BATCH_NO = NULL")
        SQLS.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        Call CommitTrans("")
    End Sub

    Private Sub txtCustOverride_MaskInputRejected(sender As System.Object, e As System.Windows.Forms.MaskInputRejectedEventArgs) Handles txtCustOverride.MaskInputRejected

    End Sub

    Private Sub SetColorStatusImages(ByVal StartAt As Integer, ByVal ClearAll As Boolean)
        Dim Path As String = ASCMAIN1.Folders("Images").ToString()
        Dim active As String = Path & "16\ball_green.png"
        Dim discontinued As String = Path & "16\ball_red.png"
        Dim noReorder As String = Path & "16\ball_yellow.png"
        Dim posStart As Integer = PictureBox1.Top
        Dim posStep As Integer = 21
        If ClearAll Then
            For i As Integer = 1 To 19
                Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & i.ToString())
                pBox.Visible = False
                pBox.ImageLocation = ""
            Next i
        Else

            Dim firstRow As Integer = 0
            For i As Integer = 1 To 19
                Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & i.ToString())
                pBox.Top = posStart + (posStep * (i - 1))
            Next i

            Dim curRow As Integer = 0
            For Each grow As UltraWinGrid.UltraGridRow In grdICTSTYC1.Rows
                curRow += 1
                If grow.Index >= StartAt Then
                    If (curRow - StartAt) < 20 Then
                        Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & (curRow - StartAt).ToString())
                        Select Case grow.Cells.Item("STYLE_COLOR_STATUS").Text
                            Case Is = "A"
                                pBox.Visible = True
                                pBox.ImageLocation = active
                                'pBox.Bottom
                            Case Is = "N"
                                pBox.Visible = True
                                pBox.ImageLocation = noReorder
                            Case Is = "D"
                                pBox.Visible = True
                                pBox.ImageLocation = discontinued
                            Case Else
                                pBox.Visible = False
                                pBox.ImageLocation = ""
                        End Select
                    End If

                End If
            Next
            For i As Integer = curRow + 1 To 19
                Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & i.ToString())
                pBox.Visible = False
                pBox.ImageLocation = ""
            Next i
        End If
    End Sub

    Private Sub chkLapQuote_CheckedChanged(sender As Object, e As EventArgs) Handles chkLapQuote.CheckedChanged
        If Not LoadFinished Then
            Exit Sub
        Else
            RefreshSOTORDRX()
        End If
    End Sub

    Private Sub UltraTextEditor17_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor17.ValueChanged
        If IsNumeric(UltraTextEditor17.Value) Then
            If Val(UltraTextEditor17.Value) > 0 Then
                UltraTextEditor17.Appearance.BackColor = Color.OrangeRed
            Else
                UltraTextEditor17.Appearance.BackColor = Color.Empty
            End If
        Else
            UltraTextEditor17.Appearance.BackColor = Color.Empty
        End If
    End Sub
    Private Sub ShowPromo(ByVal STYLE_CODE As String)
        Dim OnPromo As Boolean = False
        Dim PROMO_START_DATE As DateTime
        Dim PROMO_END_DATE As DateTime
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE,")
        sql.AppendLine("MAX(P2.PROMO_UNIT_PRICE) PROMO_UNIT_PRICE")
        sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2")
        sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
        sql.AppendLine("AND P2.STYLE_CODE = :PARM1")
        sql.AppendLine("GROUP BY P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE")
        Dim tblICTPROMX As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
        For Each rowICTPROMX As DataRow In tblICTPROMX.Select("", "PROMO_START_DATE")
            PROMO_START_DATE = CDate(rowICTPROMX.Item("PROMO_START_DATE").ToString & String.Empty)
            PROMO_END_DATE = CDate(rowICTPROMX.Item("PROMO_END_DATE").ToString & String.Empty)
            If PROMO_START_DATE <= Now() And PROMO_END_DATE >= Now() Then
                OnPromo = True
            End If
        Next
        If OnPromo Then
            lblPromo.Text = String.Format("Style On Promo {0} - {1}", PROMO_START_DATE.ToShortDateString, PROMO_END_DATE.ToShortDateString)
            lblPromo.Visible = True
            btnShowPromo.Visible = True
        Else
            lblPromo.Text = ""
            lblPromo.Visible = False
            btnShowPromo.Visible = False
        End If
    End Sub

    Private Sub btnShowPromo_Click(sender As Object, e As EventArgs) Handles btnShowPromo.Click
        Dim F As New ASFMSGBF
        F.grdGroupBy = True
        F.grdFilter = True
        Dim sql As New System.Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("P1.PROMO_DESC As Promotion,")
        sql.AppendLine("P1.PROMO_START_DATE As Beginning,")
        sql.AppendLine("P1.PROMO_END_DATE As Ending,")
        sql.AppendLine("P2.STYLE_CODE As Style,")
        sql.AppendLine("S1.STYLE_DESC As Description,")
        sql.AppendLine("MAX(P2.PROMO_UNIT_PRICE) As Price")
        sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2, ICTSTYL1 S1")
        sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
        sql.AppendLine("AND P2.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND (P1.PROMO_START_DATE <= SYSDATE AND P1.PROMO_END_DATE >= SYSDATE)")
        sql.AppendLine("GROUP BY")
        sql.AppendLine("P1.PROMO_DESC,")
        sql.AppendLine("P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE,")
        sql.AppendLine("P2.STYLE_CODE,")
        sql.AppendLine("S1.STYLE_DESC")
        Dim tblICTPROMX As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        If tblICTPROMX.Rows.Count > 0 Then
            F.Show_grd(tblICTPROMX, Me, "Current Active Promotions", "")
            F.Dispose()
            F = Nothing
        End If
    End Sub

    Private Sub grdSHIP2_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSHIP2.InitializeLayout

    End Sub
End Class