Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid
Imports NLog

Public Class SOFPICKS

    Dim SOTPICK1 As String
    Dim SOTORDR1 As String
    Dim SOTORDR2 As String

    Dim SOTORDQ0 As String
    Dim sqlSOTORDQ0 As String

    Dim SOTORDQ1 As String
    Dim rowICTWHSE1 As DataRow
    Dim WHSE_CODE As String

    Dim SOTORDRX As String
    Dim sqlSOTORDRX As String

    Dim SOTORDR0 As String

    Dim ICTSTATX As String
    'Dim sqlICTSTAT2 As String

    Dim sqlSOTPICK2 As String = ""
    Dim sqlSOTPICK1 As String = ""
    Dim sqlSOTPICK0 As String = ""

    Dim PICK_BATCH_NO As String

    Private vLabelPrinter As ASCPRINT

    Dim Appearance_Red As New Infragistics.Win.Appearance
    Dim Appearance_Orange As New Infragistics.Win.Appearance
    Dim Appearance_Magenta As New Infragistics.Win.Appearance
    Dim Appearance_Empty As New Infragistics.Win.Appearance

    Dim Appearance_Green As New Infragistics.Win.Appearance
    Dim Appearance_Blue As New Infragistics.Win.Appearance
    Dim Appearance_Yellow As New Infragistics.Win.Appearance

    Dim refresh_next_time_we_look_at_orders As Boolean = False

    Dim zplPrint As New TAC.TACZPLT1()
    Dim ORDR_NOs_Tried As New List(Of String)

    'Private Shared logger As Logger = LogManager.GetCurrentClassLogger()
    Dim MAX_ORDERS_TO_RELEASE As Integer = 25
    Dim SALES_DIVISION_CODE_SKIN As String = "30"
    Dim WHSE_CODE_SKIN As String = "NJC"


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        btnMaxRelease.Text = CStr(MAX_ORDERS_TO_RELEASE)

        Appearance_Red.ForeColor = Drawing.Color.Red
        Appearance_Orange.BackColor = Drawing.Color.Orange
        Appearance_Magenta.BackColor = Drawing.Color.Magenta

        Appearance_Green.BackColor = Drawing.Color.LightGreen
        Appearance_Blue.BackColor = Drawing.Color.LightBlue
        Appearance_Yellow.BackColor = Drawing.Color.Yellow

        Create_WorkTable(True)

        InquiryMode = (MENU_ITEM_OBJECT = "SOFPICKT")

        lblORDR_NO.Visible = InquiryMode
        txtORDR_NO.Visible = InquiryMode

        With dst

            ASCMAIN1.sql = $"Select SOTORDQ1.* from {SOTORDQ1} SOTORDQ1"
            Create_TDA(.Tables.Add, "SOTORDQ1", "**", 0, False,, 2)
            With .Tables("SOTORDQ1")
                .Columns("ORDR_CNT").DataType = GetType(System.Int32)
                .Columns("ORDR_QTY_OPEN").DataType = GetType(System.Int32)
                .Columns("ORDR_QTY_PICK").DataType = GetType(System.Int32)
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = $"Select SOTORDQ0.* from {SOTORDQ0} SOTORDQ0"
            Create_TDA(.Tables.Add, "SOTORDQ0", "**", 0, False,, 1)
            With .Tables("SOTORDQ0")
                .Columns("ORDR_CNT").DataType = GetType(System.Int32)
                .Columns("ORDR_QTY_OPEN").DataType = GetType(System.Int32)
                .Columns("ORDR_QTY_PICK").DataType = GetType(System.Int32)
                .Columns.Add("QTY_SHORT", GetType(System.Int32))
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                .Columns.Add("PICK_DESCRIPTION", GetType(System.String))
            End With

            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO from SOTORDR1 where ROWNUM < 1"
            SOTORDR0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_NO)")
            'ASCMAIN1.sql = $"Select * from {SOTORDR0}"
            Create_TDA(.Tables.Add("SOTORDR0"), SOTORDR0, "*", 0, True,, 1)

            ASCMAIN1.sql = "Select SOTORDR1.*, SOTORDR1.ORDR_SHIP_COMPLETE ORDR_ALLO_COMPLETE from SOTORDR1 where ROWNUM < 1"
            SOTORDR1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")

            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_OPEN PICK_QTY, SOTORDR2.ORDR_QTY_OPEN PICK_QTY_CANC" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_OPEN QTY_AVAIL" & vbCrLf _
            & " from SOTORDR2 where ROWNUM < 1"
            SOTORDR2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add Primary Key (ORDR_NO,ORDR_LNO)")

            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where ROWNUM < 1"
            SOTPICK1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK1 & " Add Primary Key (PICK_NO)")

            ASCMAIN1.sql = "Select SOTPICK0.*, SOTTRCK1.TRUCK_TYPE, X.PICK_QTY" & vbCrLf _
                & " from SOTPICK0,SOTTRCK1" & vbCrLf _
                & ",(Select PICK_BATCH_NO, SUM (PICK_QTY) PICK_QTY from SOTPICK1,SOTPICK2" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO group by PICK_BATCH_NO) X" & vbCrLf _
                & " where SOTTRCK1.TRUCK_NO (+) = SOTPICK0.TRUCK_NO" & vbCrLf _
                & "   and X.PICK_BATCH_NO = SOTPICK0.PICK_BATCH_NO"
            sqlSOTPICK0 = ASCMAIN1.sql
            ASCMAIN1.sql &= $"   And SOTPICK0.PICK_BATCH_NO In (Select Distinct PICK_BATCH_NO from {SOTPICK1})"

            Create_TDA(.Tables.Add, "SOTPICK0", "**", 0, True)

            '                & ", SUM (NVL(SOTPICK2.PICK_QTY,0) - (NVL(SOTPICK2.PICK_QTY_CONF,0) + NVL(SOTPICK2.PICK_QTY_BACK,0))) PICK_QTY_LEFT" & vbCrLf _
            ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_COMPLETE" & vbCrLf _
                & ", Case when SOTPICK2_SUM.PICK_QTY_LEFT = 0 THEN '2' ELSE CASE when SOTPICK2_SUM.PICK_QTY_LEFT = PICK_QTY THEN '0' ELSE '1' END END PICK_STATUS_CODE" & vbCrLf _
                & ", SOTPICK2_SUM.PICK_QTY, SOTPICK2_SUM.PICK_QTY_LEFT" & vbCrLf _
                & ", SOTORDR1.ORDR_BUYER_NAME" & vbCrLf _
                & " from SOTPICK1,SOTORDR1" & vbCrLf _
                & ", (Select SOTPICK2.PICK_NO" & vbCrLf _
                & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
                & ", SUM (CASE WHEN NVL(SOTPICK2.PICK_QTY,0) = 0 THEN 0 ELSE NVL(SOTPICK2.PICK_QTY,0) - (NVL(SOTPICK2.PICK_QTY_CONF,0) + NVL(SOTPICK2.PICK_QTY_BACK,0)) END) PICK_QTY_LEFT" & vbCrLf _
                & $"  from SOTPICK2, {SOTPICK1} SOTPICK1" & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO group by SOTPICK2.PICK_NO) SOTPICK2_SUM" & vbCrLf _
                & " where SOTPICK2_SUM.PICK_NO (+) = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS <> 'D'" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
            sqlSOTPICK1 = ASCMAIN1.sql
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by SOTPICK2.PICK_NO", " and ROWNUM < 1 group by SOTPICK2.PICK_NO")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "and SOTPICK1.PICK_STATUS <> 'D'", " and ROWNUM < 1 and SOTPICK1.PICK_STATUS <> 'D'")
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, True)
            With .Tables("SOTPICK1")
                .Columns.Add("RESOLUTION")
            End With

            ' STOP - IT WOULD BE CONVENIENT IF SOTORDR2 WAS A TEMP FILE, OR IF SOTPICK2 WAS A TEMP FILE TO JOIN SC
            ASCMAIN1.sql = "Select SOTPICK2.*, ICTSTYD1.LOCATION_CODE" & vbCrLf _
            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, ICTSTYC1.UPC_CODE, SOTORDR2.STYLE_CLASS_CODE, SOTORDR1.ORDR_BUYER_NAME" & vbCrLf _
            & " from " & SOTPICK1 & " SOTPICK1, SOTPICK2, SOTORDR2, SOTORDR1, ICTSTYL1, ICTSTYD1, ICTSTYC1" & vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" _
            & "   And SOTORDR1.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   And SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   And SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & $"   And ICTSTYD1.WHSE_CODE (+) = '{WHSE_CODE_SKIN}'" & vbCrLf & vbCrLf _
            & "   And ICTSTYD1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf & vbCrLf _
            & "   And ICTSTYD1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
            & "   And ICTSTYC1.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf & vbCrLf _
            & "   And ICTSTYC1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
            & "   And ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE"
            '& "   And SOTORDR2.STYLE_CODE = :PARM1" & vbCrLf & vbCrLf _
            '& "   And SOTORDR2.COLOR_CODE = :PARM2" & vbCrLf _
            sqlSOTPICK2 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 1,, "VV", 2)
            'Create_TDA(.Tables.Add, "SOTPICK2", "**", 1,, "V")
            With .Tables("SOTPICK2").Columns
                .Add("PRI", GetType(System.Int32))
                .Add("OSL", GetType(System.Int32))
                .Add("PAL", GetType(System.Int32))
                .Add("CRT", GetType(System.Int32))
                .Add("TRK", GetType(System.Int32))
                .Add("LNF", GetType(System.Int32))
                '.Add("PICK_STATUS_CODE", GetType(System.String), "IIF(ISNULL(PICK_QTY_CONF,0) = 0 And ISNULL(PICK_QTY_BACK,0) = 0, IIF(ISNULL(PICK_QTY_CANC,0) <> 0,'3','0'), IIF(ISNULL(PICK_QTY,0) - ISNULL(PICK_QTY_CONF,0) - ISNULL(PICK_QTY_BACK,0) = 0, '2', '1'))")
                .Add("PICK_STATUS_CODE", GetType(System.String), "IIF(ISNULL(PICK_QTY_BACK,0) <> 0 AND ISNULL(PICK_QTY,0) = 0, '4', IIF(ISNULL(PICK_QTY_CONF,0) = 0 And ISNULL(PICK_QTY_BACK,0) = 0, IIF(ISNULL(PICK_QTY_CANC,0) <> 0,'3','0'), IIF(ISNULL(PICK_QTY,0) - ISNULL(PICK_QTY_CONF,0) - ISNULL(PICK_QTY_BACK,0) = 0, '2', '1')))")
                .Add("RESOLUTION")
                .Add("FORCE_BO")
            End With

            With .Tables.Add("SOTPICKG")
                .Columns.Add("PICK_NO")
                .Columns.Add("PICK_LNO", GetType(System.Int32))
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("STYLE_DESC")
                .Columns.Add("UPC_CODE")
                .Columns.Add("TOTE_NO")
                .Columns.Add("SLOT_NO", GetType(System.Int32))
                .Columns.Add("PICK_GO_BACK_QTY", GetType(System.Int32))
                .Columns.Add("LOCATION_CODE")
                .PrimaryKey = New DataColumn() { .Columns("PICK_NO"), .Columns("PICK_LNO")}
            End With


            ASCMAIN1.sql = "Select SOTORDR5.*" _
            & " from (Select Distinct ORDR_NO from " & SOTPICK1 & ") SOTPICK1_ORDR_NO,SOTORDR5 " _
            & " where SOTORDR5.ORDR_NO = SOTPICK1_ORDR_NO.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, False, "", 2)

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1, False)
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            With .Tables("SOTORDR2")
                .Columns.Add("PICK_SEQ", GetType(System.String), "IIF(ISNULL(STYLE_CLASS_CODE,'??')='OP','A','Z')")
            End With

            With .Tables.Add("ICTSTATZ")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("ACTION_IF_SHORT")
                .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With

            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_PICK" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_ON_HAND QTY_AVA, ICTSTAT2.WHSE_QTY_ON_HAND QTY_SHORT, ICTSTYD1.LOCATION_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, SOTORDR2.STYLE_CLASS_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & " from SOTORDR2,ICTSTAT2,ICTSTYD1,ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTSTATO", "**", 0, False, "", 2)
            With .Tables("ICTSTATO").Columns
                .Add("ORDR_QTY_ALLO", GetType(System.Int32))
                .Add("PRI", GetType(System.Int32))
                .Add("OSL", GetType(System.Int32))
                .Add("PAL", GetType(System.Int32))
                .Add("CRT", GetType(System.Int32))
                .Add("TRK", GetType(System.Int32))
                .Add("LNF", GetType(System.Int32))
                .Add("ACTION_IF_SHORT")
                .Add("QTY_PICKED", GetType(System.Int32))
            End With
            .Tables("ICTSTATO").Columns("LOCATION_CODE").AllowDBNull = True

            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.ORDR_NO from SOTORDR2"
            Create_TDA(.Tables.Add, "ICTSTATP", "**", 0, False, "", 3)

            Create_Relation("ICTSTATO", "ICTSTATP", "STYLE_CODE,COLOR_CODE")

            Create_TDA(.Tables.Add, "SOTTOTE1", "*")

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTORDR2.STYLE_CLASS_CODE, SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.STYLE_CODE STYLE_CODE_CASE" & vbCrLf _
                & " from SOTORDR1,SOTORDR2" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRF", "**", 0, False, "V", 2)

            ASCMAIN1.sql = $"Select * from {SOTORDRX}"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "", 0)
            .Tables("SOTORDRX").Columns("QAVA").DataType = GetType(System.Int32)
            .Tables("SOTORDRX").Columns.Add("SEL")
            .Tables("SOTORDRX").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = $"Select ICTSTATX.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC from {ICTSTATX} ICTSTATX, ICTSTYL1, ICTCOLR1 where ICTSTYL1.STYLE_CODE = ICTSTATX.STYLE_CODE and ICTCOLR1.COLOR_CODE = ICTSTATX.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTSTATX", "**", 0, False, "", 0)
            .Tables("ICTSTATX").Columns.Add("QTY_AVA", GetType(System.Int32), "ISNULL(WHSE_QTY_ON_HAND,0) - ISNULL(WHSE_QTY_PICK,0) - ISNULL(LOC_QTY_NOT_AVA,0)")
            .Tables("ICTSTATX").Columns.Add("QTY_SHORT", GetType(System.Int32), "IIF(ISNULL(QTY_AVA,0) - ISNULL(ORDR_QTY_OPEN,0) >=0, NULL, ISNULL(QTY_AVA,0) - ISNULL(ORDR_QTY_OPEN,0))")

            Create_TDA(.Tables.Add, "POTORDR1", "*")
            Create_TDA(.Tables.Add, "POTORDR2", "*")

        End With

        grdSOTORDQ1.DataSource = dst.Tables("SOTORDQ1")
        grdSOTORDQ0.DataSource = dst.Tables("SOTORDQ0")

        grdSOTPICK0.DataSource = dst.Tables("SOTPICK0")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")

        grdICTSTATO.DataSource = dst.Tables("ICTSTATO")

        grdICTSTATX.DataSource = dst.Tables("ICTSTATX")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")

        Create_Summary(grdSOTORDQ1, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDQ1, New String() {"SEL", "BO", "ORDR_QTY_OPEN", "ORDR_QTY_PICK"})

        Create_Summary(grdSOTORDQ0, "WHSE_CODE", "Count")
        Create_Summary(grdSOTORDQ0, New String() {"ORDR_CNT", "ORDR_QTY_OPEN", "ORDR_QTY_PICK"})

        Create_Summary(grdSOTPICK0, "PICK_BATCH_NO", "Count")
        Create_Summary(grdSOTPICK0, New String() {"ORDERS", "PICK_QTY"})

        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
        Create_Summary(grdSOTPICK1, New String() {"PICK_QTY", "PICK_QTY_LEFT"})

        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICK2, New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_BACK", "PICK_QTY_CANC"})

        Create_Summary(grdICTSTATO, "STYLE_CODE", "Count")
        Create_Summary(grdICTSTATO, New String() {"ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_ALLO", "WHSE_QTY_ON_HAND", "WHSE_QTY_PICK", "QTY_AVA"})

        Create_Summary(grdICTSTATX, "STYLE_CODE", "Count")
        Create_Summary(grdICTSTATX, New String() {"ORDR_QTY_OPEN", "WHSE_QTY_ON_HAND", "WHSE_QTY_PICK", "LOC_QTY_NOT_AVA", "WHSE_QTY_ALLO", "ORDERS", "ORDERS_ALLO", "QTY_AVA", "QTY_SHORT"})

        Create_Summary(grdSOTORDRX, "ORDR_NO", "Count")

        Create_Summary(grdSOTORDRX, New String() {"SEL", "ORDR_QTY_OPEN", "ORDR_QTY_PICK"})

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTORDQ0, grdSOTORDQ1, grdSOTPICK0, grdSOTPICK1, grdSOTPICK2, grdICTSTATO}
            grd.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.LightGreen
        Next

        grdSOTORDRX.DisplayLayout.GroupByBox.Hidden = False
        With grdSOTORDRX.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("SEL").Hidden = True
            .Columns("ORDR_NO").Header.Fixed = True
        End With

        Show_Filter(grdICTSTATX, True)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTORDRX, grdICTSTATX}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns

                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    If c.Key = "SEL" Then
                        c.Header.Appearance.BackColor2 = System.Drawing.Color.Pink
                        c.CellActivation = Activation.AllowEdit
                    End If

                    If grd.Name = "grdICTSTATX" Then
                        If New String() {"ORDERS", "ORDERS_ALLO"}.Contains(c.Key) Then
                            c.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                        End If
                        If New String() {"WHSE_QTY_ALLO", "ORDR_QTY_OPEN", "ORDR_QTY_BACK"}.Contains(c.Key) Then
                            c.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                        End If
                        If New String() {"QTY_SHORT"}.Contains(c.Key) Then
                            c.Header.Appearance.BackColor2 = System.Drawing.Color.Pink
                        End If
                    End If


                    If New String() {"ORDR_SOURCE"}.Contains(c.Key) Then
                        c.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    End If

                    If New String() {"ONHD", "ONPO", "PICK", "OPEN", "QAVA"}.Contains(c.Key) Then
                        c.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                    End If
                Next
            End With
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTORDQ0, grdSOTORDQ1}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                .Columns("SEL").Header.Fixed = True
            End With

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns

                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    If c.Key = "SEL" Then
                        c.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        c.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                Next
            End With
        Next

        grdSOTORDQ0.DisplayLayout.Bands(0).Columns("QTY_SHORT").Hidden = True

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTPICK0, grdSOTPICK1, grdSOTPICK2}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns

                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    If c.Key = "SEL" Then
                        c.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        c.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If

                    If grd.Name = "grdSOTPICK1" Or grd.Name = "grdSOTPICK2" Then
                        If c.Key = "RESOLUTION" Then
                            c.CellActivation = UltraWinGrid.Activation.AllowEdit
                            c.Header.Appearance.BackColor2 = System.Drawing.Color.Pink
                        Else
                            c.CellActivation = UltraWinGrid.Activation.NoEdit
                        End If
                    End If

                    If grd.Name = "grdSOTPICK2" Then
                        c.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    End If

                Next
            End With
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTSTATO}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns
                    c.CellActivation = Activation.NoEdit
                    If c.Key = "ACTION_IF_SHORT" Then c.CellActivation = Activation.AllowEdit
                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next
            End With

            With grd.DisplayLayout.Bands(1)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns
                    c.CellActivation = Activation.NoEdit
                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                Next
            End With
        Next

        ASCMAIN1.Add_Value_List(grdSOTORDQ1, "DESTINATION",, New String() {":", "S:Ship-To", "C:Consumer"})
        'ASCMAIN1.Add_Value_List(grdSOTORDQ1, "ORDR_SOURCE",, New String() {":", "W:Web", "L:Lab Order", "P:API Portal", "E:EDI", "K:Keyed", "C:Converted"})
        '('K', 'C', 'E', 'L', 'W')"
        ASCMAIN1.Add_Value_List(grdSOTORDQ1, "ORDR_SOURCE",, New String() {":", "W:Web"})

        ASCMAIN1.Add_Value_List(grdSOTORDQ0, "DESTINATION",, New String() {":", "S:Ship-To", "C:Consumer"})
        'ASCMAIN1.Add_Value_List(grdSOTORDQ0, "ORDR_SOURCE",, New String() {":", "W:Web", "L:Lab Order", "P:API Portal", "E:EDI", "K:Keyed", "C:Converted"})
        ASCMAIN1.Add_Value_List(grdSOTORDQ0, "ORDR_SOURCE",, New String() {":", "W:Web"})

        ASCMAIN1.Add_Value_List(grdSOTPICK0, "PICK_BATCH_STATUS",, New String() {":", "O:Released", "P:Picking", "N:Picked", "R:ReqRes"})
        ASCMAIN1.Add_Value_List(grdSOTPICK0, "TRUCK_TYPE",, New String() {":", "P:Pre-Config", "X:Custom", "R:Regular"})
        ' NOTE HOW THESE CUSTOM LIST VALUES ARE IGNORED - WHEN THERE ARE VALUES DEFINED IN ASTCODE1 OR IN TABMAIN1

        ' ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS_CODE",, New String() {":", "0:Not Picked", "1:Partial", "2:Complete", "3:Cancelled"})
        ASCMAIN1.Add_Value_List(grdSOTPICK2, "PICK_STATUS_CODE",, New String() {":", "0:Not Picked", "1:Partial", "2:Complete", "3:Cancelled", "4:Back-Ordered"})

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "RESOLUTION",, New String() {":", "P:Re-Pick", "T:Triage Truck", "D:De-Release", "B:Back-Order"})
        ASCMAIN1.Add_Value_List(grdSOTPICK2, "RESOLUTION",, New String() {":", "P:Re-Pick", "B:Back-Order"})

        'ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS",, New String() {":", "F:Invoiced", "C:Cancelled", "D:Deleted", "P:In Pick"})
        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS")

        ASCMAIN1.Add_Value_List(grdICTSTATO, "ACTION_IF_SHORT",, New String() {":", "B:Back-Order", "C:Cancel"})

        MakeTransparent(chkIgnoreInvtyShort)
        MakeTransparent(lblSEL)

        Show_Filter(grdSOTORDQ0, True)
        Show_Filter(grdSOTORDRX, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("WHSE_CODE")

                WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)

                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & "Invalid DC Code"
                    Exit Select
                    'ElseIf rowICTWHSE1.Item("WHSE_CODE_SL") & String.Empty = String.Empty Then
                    '    EMsg &= vbCr & "DC is not assigned a Stock Lens Warehouse"
                    '    Exit Select
                End If

                WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")

                If txtORDR_NO.TextLength > 0 Then
                    Validate_Code("ORDR_NO")
                End If

            Case "Resolve"

                If grdSOTPICK0.ActiveRow Is Nothing OrElse grdSOTPICK0.ActiveRow.IsFilterRow Then
                    EMsg &= vbCr & "You Must Select a Pick Batch that Requires Resolution in order to Resolve"
                Else

                    PICK_BATCH_NO = grdSOTPICK0.ActiveRow.Cells("PICK_BATCH_NO").Value & ""
                    Dim rowSOTPICK0 As DataRow = LookUp("SOTPICK0", PICK_BATCH_NO)

                    Dim PICK_BATCH_STATUS As String = rowSOTPICK0.Item("PICK_BATCH_STATUS") & ""
                    If PICK_BATCH_STATUS <> "R" Then
                        EMsg &= vbCr & "You Must Select a Pick Batch that Requires Resolution in order to Resolve"
                    End If

                    If EMsg = "" Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Locking Sales Orders")

                        Dim TRUCK_NO As String = rowSOTPICK0.Item("TRUCK_NO") & ""
                        If Not ASCMAIN1.Logical_Lock("SOTTRCK1", TRUCK_NO) Then Exit Sub

                        If Not ASCMAIN1.Logical_Lock("SOTPICK0", PICK_BATCH_NO) Then Exit Sub
                        ASCMAIN1.sql = sqlSOTPICK1 & $" and SOTPICK1.PICK_BATCH_NO = '{PICK_BATCH_NO}'"
                        Fill_Records("SOTPICK1",,, ASCMAIN1.sql)
                        'Fill_Records("SOTPICK1")

                        For Each row As DataRow In dst.Tables("SOTPICK1").Select($"PICK_BATCH_NO = '{PICK_BATCH_NO}'")
                            Dim PICK_NO As String = row.Item("PICK_NO")
                            If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                            Dim ORDR_NO As String = row.Item("ORDR_NO")
                            ASCMAIN1.Progress("-", ORDR_NO)
                            If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub
                        Next

                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                    End If
                End If

            Case "Update"

                If dst.Tables("SOTPICK2").Select($"ISNULL(RESOLUTION,'?') = '?' and ISNULL(PICK_CNL_STATUS,'0') = '1'").Length > 0 Then
                    EMsg &= vbCr & $"Cannot Resolve a Pick Batch if any PT lines are unresolved"
                End If

                ' TEMPORARY UNTIL WE FIGURE THIS OUT
                If dst.Tables("SOTPICK2").Select("RESOLUTION = 'T' OR RESOLUTION = 'D'").Length > 0 Then
                    EMsg &= vbCr & $"Do not Triage or De-Release until ABS has tested this with you"
                End If

                ' HOW DO WE HANDLE IT WHEN A PT IS PARTIALLY PICKED, BUT MD STILL WANTS TO DE-RELEASE
                For Each row As DataRow In dst.Tables("SOTPICK1").Select("RESOLUTION = 'D'")
                    Dim PICK_NO As String = row.Item("PICK_NO")
                    If dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' and ISNULL(PICK_QTY_CONF,0) <> 0").Length > 0 Then
                        EMsg &= vbCr & $"Cannot De-Release a PT ({PICK_NO}) that has been Partially Picked"
                    End If
                Next

                Dim otherLinesMsg As String = ""
                If EMsg = "" Then
                    ' SL only
                    Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").Rows.Find(PICK_BATCH_NO)

                    For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("FORCE_BO = '1'")
                        rowSOTPICK2.Item("FORCE_BO") = DBNull.Value
                    Next
                End If

                If EMsg = "" Then
                    If MsgBox("OK to Resolve these Pick Tickets?" & otherLinesMsg, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
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
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                ASCMAIN1.Progress("Refreshing data...")
                Setup_Screen()
                ASCMAIN1.Progress("")

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                Update_Record()
                Mode_Settings(False)
                Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
                Click_Command("Load")

            Case "Cancel"
                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                Mode_Settings(False)
                Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
                Click_Command("Load")

            Case "Resolve"
                EntryMode = "R"
                Load_Record()
                Mode_Settings(True)

            Case "Label Re-Print"
                Label_RePrint()

            Case "Fetch All Orders"
                Fetch_All_Orders()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Refresh").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Resolve").Visible = Not (EntryMode = "R") And tabMain.SelectedTab.Key = "Pick Tickets" And (MENU_ITEM_OBJECT = "SOFPICKS")

                    .Items("Load").Visible = Not (EntryMode = "R")
                    .Items("Refresh").Visible = Not (EntryMode = "R")
                    .Items("Done").Visible = Not (EntryMode = "R")

                    .Items("Update").Visible = (EntryMode = "R")
                    .Items("Cancel").Visible = (EntryMode = "R")

                End With

                .Groups("Order Details").Visible = False ' tabMain.SelectedTab.Key = "Order Details"
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        chkShowHolds.Visible = ScreenMode

        tabMain.Visible = ScreenMode
        lblRESOLUTION.Visible = ScreenMode And (EntryMode = "R")

        If ScreenMode Then
            If Not InquiryMode Then Set_Read_Only_for_ctl(chkShowHolds, False)

            tabMain.Tabs("Orders").Enabled = Not (EntryMode = "R")
            tabMain.Tabs("Ship Vias").Enabled = Not (EntryMode = "R")

            grdSOTPICK1.DisplayLayout.Bands(0).Columns("RESOLUTION").Hidden = Not (EntryMode = "R")
            grdSOTPICK2.DisplayLayout.Bands(0).Columns("RESOLUTION").Hidden = Not (EntryMode = "R")
            grdSOTPICK1.DisplayLayout.Bands(0).Columns("PICK_STATUS_CODE").Hidden = (EntryMode = "R")
            grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_STATUS_CODE").Hidden = (EntryMode = "R")


            grdSOTPICK1.DisplayLayout.Bands(0).Columns("LAST_OPER").Hidden = (EntryMode = "R")
            grdSOTPICK1.DisplayLayout.Bands(0).Columns("LAST_DATE").Hidden = (EntryMode = "R")

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDQ1", "SOTORDQ0", "ICTSTATP", "ICTSTATO", "SOTPICK0", "SOTPICK1" _
            , "SOTPICK2", "SOTORDRX", "ICTSTATX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE_SKIN
        txtORDR_NO.Clear()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Setup_Screen()

        If EntryMode = "N" Then
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").Rows.Find(PICK_BATCH_NO)
        Dim PICK_BATCH_STATUS As String = "N"
        Dim WH_TRAN_NOs As New List(Of String)
        Dim PICK_NOsD As New List(Of String)
        Dim PICK_NOsT As New List(Of String)
        Dim PICK_NOsB As New List(Of String)

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_REQ_RES = '1'")
            rowSOTPICK1.Item("PICK_REQ_RES") = "0"
        Next

        dst.Tables("SOTPICKG").Rows.Clear()

        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_CNL_STATUS = '1'")
            Dim RESOLUTION As String = rowSOTPICK2.Item("RESOLUTION")
            Dim WH_TRAN_NO As String = rowSOTPICK2.Item("WH_TRAN_NO") & ""
            Dim PICK_NO As String = rowSOTPICK2.Item("PICK_NO")

            rowSOTPICK2.Item("PICK_CNL_STATUS") = "0"

            If RESOLUTION = "B" Then ' Back-Order
                If Not PICK_NOsB.Contains(PICK_NO) Then
                    PICK_NOsB.Add(PICK_NO)
                End If
                ' Mark PT for review down below - we may need a Go-Back for the case

            ElseIf RESOLUTION = "P" Then ' Re-Pick
                rowSOTPICK2.Item("PICK_QTY_CANC") = 0
                PICK_BATCH_STATUS = "P"
                If Not WH_TRAN_NOs.Contains(WH_TRAN_NO) Then
                    WH_TRAN_NOs.Add(WH_TRAN_NO)
                End If
                rowSOTPICK2.Item("WH_TRAN_NO") = DBNull.Value

                ElseIf RESOLUTION = "D" Then ' De-Release
                If Not WH_TRAN_NOs.Contains(WH_TRAN_NO) Then
                    WH_TRAN_NOs.Add(WH_TRAN_NO)
                End If
                ' WE MAY NEED TO PUT THE SLS BACK ON THE SHELF - PHYSICALLY AS WELL AS LOGICALLY
                If Not PICK_NOsD.Contains(PICK_NO) Then PICK_NOsD.Add(PICK_NO)

                    rowSOTPICK2.Item("WH_TRAN_NO") = DBNull.Value

                ElseIf RESOLUTION = "T" Then ' Triage
                    If Not PICK_NOsT.Contains(PICK_NO) Then PICK_NOsT.Add(PICK_NO)

            End If

        Next

        For Each PICK_NO As String In PICK_NOsB

            ' if there are no SLs left on this PT to pick,
            ' and there are cloths that were picked,
            ' then put all of the cases back to primary location
            ' and issue message to manually place te cloths back on the shelf

            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            Dim WHSE_CODE As String = rowSOTPICK1.Item("WHSE_CODE")

            Dim PICK_QTY_CONF_CS_total As Int32 = 0
            Dim PICK_QTY_CONF_FR_total As Int32 = 0

            Dim FORCE_BOs As New List(Of Int32)

            Dim sqlw As String = $"PICK_NO = '{PICK_NO}' and (ISNULL(PICK_QTY,0) > 0 OR ISNULL(FORCE_BO,'0') = '1')"
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw)
                Dim PICK_QTY_CONF As Int32 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
                Dim STYLE_CLASS_CODE As String = rowSOTPICK2.Item("STYLE_CLASS_CODE")
                Dim FORCE_BO As String = rowSOTPICK2.Item("FORCE_BO") & ""
                Dim PICK_LNO As Int32 = Val(rowSOTPICK2.Item("PICK_LNO") & "")

                If FORCE_BO = "1" Then
                    FORCE_BOs.Add(PICK_LNO)
                Else
                    If STYLE_CLASS_CODE = "CC" Then
                        PICK_QTY_CONF_CS_total += PICK_QTY_CONF
                    Else
                        PICK_QTY_CONF_FR_total += PICK_QTY_CONF
                    End If
                End If
            Next

            If FORCE_BOs.Count > 0 Or (PICK_QTY_CONF_CS_total > 0 And PICK_QTY_CONF_FR_total = 0) Then
                ' we have Go-Backs
                Dim sqlw2 As String = " and (ISNULL(FORCE_BO,'0') = '1' or STYLE_CLASS_CODE = 'CC') and PICK_QTY_CONF > 0"
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw & sqlw2)
                    Dim PICK_QTY_CONF As Int32 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
                    rowSOTPICK2.Item("PICK_QTY_CONF") = 0
                    rowSOTPICK2.Item("PICK_QTY_BACK") = PICK_QTY_CONF
                    rowSOTPICK2.Item("PICK_GO_BACK") = "1"

                    ' MOVE THE CS FROM THE TRUCK BACK TO THE PRIMARY LOCATION
                    ' REPORT THE GO-BACK TO THE USER

                    Dim LOCATION_CODE As String = rowSOTPICK2.Item("LOCATION_CODE")
                    Dim PICK_LNO As Int32 = Val(rowSOTPICK2.Item("PICK_LNO") & "")

                    Dim rowSOTPICKG As DataRow = dst.Tables("SOTPICKG").NewRow
                    With rowSOTPICKG
                        .Item("PICK_NO") = PICK_NO
                        .Item("PICK_LNO") = PICK_LNO
                        .Item("STYLE_CODE") = rowSOTPICK2.Item("STYLE_CODE")
                        .Item("COLOR_CODE") = rowSOTPICK2.Item("COLOR_CODE")
                        .Item("STYLE_DESC") = rowSOTPICK2.Item("STYLE_DESC")
                        .Item("UPC_CODE") = rowSOTPICK2.Item("UPC_CODE")
                        .Item("TOTE_NO") = rowSOTPICK1.Item("TOTE_NO")
                        .Item("SLOT_NO") = rowSOTPICK1.Item("SLOT_NO")
                        .Item("PICK_GO_BACK_QTY") = PICK_QTY_CONF
                        .Item("LOCATION_CODE") = LOCATION_CODE
                    End With
                    dst.Tables("SOTPICKG").Rows.Add(rowSOTPICKG)

                    ASCDATA1.ExecuteSQL("BEGIN ICPSTAT3(:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8,:PARM9,:PARM10,:PARM11); END;", "VVVVVVNVVNV",
                           {rowSOTPICK2.Item("STYLE_CODE"), rowSOTPICK2.Item("COLOR_CODE"), HFs("WHSE_CODE"), WHSE_CODE,
                           "", "TRUCK", LOCATION_CODE, PICK_QTY_CONF,
                           "K", PICK_NO, PICK_LNO, ASCMAIN1.USER_ID})
                Next
            End If
        Next


        If dst.Tables("SOTPICKG").Rows.Count <> 0 Then
            Using FRM As New ASFMSGBF
                FRM.Show_grd(dst.Tables("SOTPICKG"), Me, "Please Return these Lenses (and/or Cloths) to the Location(s) Indicated")
            End Using
        End If

        rowSOTPICK0.Item("PICK_BATCH_STATUS") = PICK_BATCH_STATUS


        ' rowSOTPICK0.Item("ORDERS") = Val(rowSOTPICK0.Item("ORDERS") & "") - PICK_NOsD.Count - PICK_NOsT.Count
        ' presently, there is no support for Triage, and we do not allow De-Release at Order Resolution
        ' if we ever did permit either, I think we would want to reduce the ORDERS by Triage only, because the De-Released PT would still refer to the batch
        ' the truth is that this field is not really used anywhere, and used to refer to the number of orders considered for release, and now it refers to the number of PTs created
        ' it might even be reduced if a PT status changes to C
        ' it is just a helpful count on the Release Screen

        For Each WH_TRAN_NO As String In WH_TRAN_NOs
            ASCDATA1.ExecuteSP("REVERSE_CNL", "VVVV",
                               New String() {WHSE_CODE, PICK_BATCH_NO, WH_TRAN_NO, ASCMAIN1.USER_ID},
                               New String() {"P_WHSE_CODE", "P_PICK_BATCH_NO", "P_WH_TRAN_NO", "P_OPER_ID"})
        Next

        If PICK_NOsD.Count > 0 Then
            De_Release_PTs(PICK_NOsD, rowSOTPICK0)
        End If


        For Each PICK_NO As String In PICK_NOsT
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            ' CREATE A NEW BATCH WITH A TRIAGE TRUCK
            rowSOTPICK1.Item("PICK_BATCH_NO") = "NEW BATCH"
        Next

        Update_Record_TDA("SOTPICK0")
        Update_Record_TDA("SOTPICK1")
        Update_Record_TDA("SOTPICK2")

        CommitTrans("Update Complete")

    End Sub

    Sub De_Release_PTs(PICK_NOsD As List(Of String), rowSOTPICK0 As DataRow)

        Dim PICK_BATCH_NO As String = rowSOTPICK0.Item("PICK_BATCH_NO")
        Dim TRUCK_TYPE As String = rowSOTPICK0.Item("TRUCK_TYPE")
        Dim r As Integer = -1

        For Each PICK_NO As String In PICK_NOsD
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
            If rowSOTPICK1.Item("PICK_STATUS") = "C" Then
                ' Pick Ticket was Cancelled as a result of BO
            Else
                rowSOTPICK1.Item("PICK_STATUS") = "D"
                rowSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTPICK1.Item("LAST_DATE") = DateTime.Now

                Dim TOTE_NO As String = rowSOTPICK1.Item("TOTE_NO")

                If TRUCK_TYPE = "X" Then
                    ASCMAIN1.sql = $"Delete from SOTTOTE1 where TOTE_NO = :PARM1 and PICK_NO = :PARM2"
                Else
                    Dim sqlT As String = ""
                    If TRUCK_TYPE = "R" Then sqlT = ", TRUCK_NO = NULL, SLOT_NO = NULL"
                    'ASCMAIN1.sql = $"Update SOTTOTE1 Set PICK_NO = NULL{sqlT} where TOTE_NO = :PARM1 and PICK_NO = :PARM2"
                    ASCMAIN1.sql = $"Update SOTTOTE1 Set LAST_OPER = '{ASCMAIN1.USER_ID}', PICK_NO = NULL {sqlT} where TOTE_NO = :PARM1 and PICK_NO = :PARM2"
                End If
                r = -1
                r = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {TOTE_NO, PICK_NO})
                If r <> 1 Then
                    Throw New Exception($"Could not Clear Pick No {PICK_NO} from Tote {TOTE_NO}")
                End If

                De_Release_Updates(PICK_NO)
            End If

        Next
    End Sub

    Sub De_Release_Updates(PICK_NO As String)

        Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
        Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")

        ASCMAIN1.sql = "Merge into ICTSTAT2 using (" & vbCrLf _
                            & "Select SOTPICK1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, Sum(SOTPICK2.PICK_QTY) PICK_QTY from SOTPICK1" & vbCrLf _
                            & " join SOTPICK2 on (SOTPICK1.PICK_NO = SOTPICK2.PICK_NO)" & vbCrLf _
                            & " join SOTORDR2 ON (SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO and SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO)" & vbCrLf _
                            & " where SOTPICK1.PICK_NO = :PARM1" & vbCrLf _
                            & " group by SOTPICK1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                            & ") X on (X.WHSE_CODE = ICTSTAT2.WHSE_CODE and X.STYLE_CODE = ICTSTAT2.STYLE_CODE and X.COLOR_CODE = ICTSTAT2.COLOR_CODE)" & vbCrLf _
                            & " when matched Then Update" & vbCrLf _
                            & "Set WHSE_QTY_OPEN = NVL(ICTSTAT2.WHSE_QTY_OPEN,0) + X.PICK_QTY," & vbCrLf _
                            & "    WHSE_QTY_PICK = NVL(ICTSTAT2.WHSE_QTY_PICK,0) - X.PICK_QTY" & vbCrLf _
                            & " when NOT matched Then Insert (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_OPEN, WHSE_QTY_PICK)" & vbCrLf _
                            & "    Values (X.STYLE_CODE, X.COLOR_CODE, X.WHSE_CODE, X.PICK_QTY, -1 * X.PICK_QTY)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_NO)

        ASCMAIN1.sql = "Merge into SOTORDR2 using (" & vbCrLf _
                            & "Select SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO, SOTPICK1.WHSE_CODE, SOTPICK2.PICK_QTY from SOTPICK1" & vbCrLf _
                            & " join SOTPICK2 on (SOTPICK1.PICK_NO = SOTPICK2.PICK_NO)" & vbCrLf _
                            & " where SOTPICK1.PICK_NO = :PARM1" & vbCrLf _
                            & ") X on (SOTORDR2.ORDR_NO = X.ORDR_NO AND SOTORDR2.ORDR_LNO = X.ORDR_LNO)" & vbCrLf _
                            & " when Matched Then Update" & vbCrLf _
                            & "Set ORDR_QTY_OPEN = NVL(ORDR_QTY_OPEN,0) + NVL(X.PICK_QTY,0), ORDR_QTY_PICK = NVL(SOTORDR2.ORDR_QTY_PICK,0) - X.PICK_QTY"
        ' THE LINE ABOVE NEEDS TO HIT ORDR_QTY_BACK WHEN WE ARE RESTORING A BACK ORDER - WHAT A PAIN
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_NO)


        ASCMAIN1.sql = $"
          Begin 
           Declare 
            Cursor C1 Is
             Select ORDR_NO
             , Sum (ORDR_QTY_OPEN) ORDR_QTY_OPEN, Sum (ORDR_QTY_PICK) ORDR_QTY_PICK, Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP
              From SOTORDR2 Where ORDR_NO = :PARM1 group by ORDR_NO;
             Begin For R1 In C1 Loop
                Update SOTORDR1 Set
                  ORDR_STATUS = 
                    CASE WHEN R1.ORDR_QTY_OPEN > 0 THEN 'O'
                         ELSE CASE WHEN R1.ORDR_QTY_PICK > 0 THEN 'P'
                                   ELSE CASE WHEN R1.ORDR_QTY_SHIP > 0 THEN 'F'
                                             ELSE 'C' END END END
                , ORDR_PICK_SEQ = NVL(ORDR_PICK_SEQ,0) - 1
                where ORDR_NO = R1.ORDR_NO;
             End Loop; End;
            End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", ORDR_NO)

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Call Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDQ1, "SBBBBBB", "Show Filter", "Select All", "De-Select All", "Select Selected", "Release Selected Orders", "Show Inventory Requirements", "Sales Order Inquiry") ' , "De-Select All Orders with Shortages"
        Load_Popup_Menu(grdICTSTATO, "SBBBB", "Show Filter", "Item/Location Inquiry", "De-Select Order", "Cancel if Short - All Items", "Back-Order if Short - All Items")
        'Load_Popup_Menu(grdSOTORDQ0, "SBBBBBB", "Show Filter", "Select All", "De-Select All", "Release Selected Orders", "Show Inventory Requirements", "Calculate Short", "Combine Groups")
        Load_Popup_Menu(grdSOTORDQ0, "SBBB", "Show Filter", "Release Selected Order Groups", "Calculate Short", "Combine Groups")
        Load_Popup_Menu(grdSOTPICK0, "SBBBBBB", "Show Filter", "Print Custom Tote Labels", "Print Resolution Report", "Print Resolution Report - Res Only", "Print Test Report", "Print Truck Pick Tag", "De-Release Pick Batch")
        Load_Popup_Menu(grdSOTPICK1, "SBBB", "Show Filter", "Print Custom Tote Label", "Sales Order Inquiry", "De-Release Pick Tickets", "De-Pick Pick Tickets")
        Load_Popup_Menu(grdSOTPICK2, "B", "Item/Location Inquiry")
        'Load_Popup_Menu(grdSOTORDRX, "SSBBBBBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Item/Location Inquiry", "Cancel Selected Orders", "Select All", "De-Select All", "Select Selected")
        Load_Popup_Menu(grdSOTORDRX, "SSBB", "Show Filter", "Show GroupBox", "Sales Order Inquiry", "Find in Orders")
    End Sub

    Private Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles tlb.BeforeToolDropdown

        Dim grd As UltraWinGrid.UltraGrid = Nothing

        If e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        grd = GRDs(Mid(e.SourceControl.Name, 4))
        If grd.ActiveRow Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If


        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

            Case grdSOTPICK1.Name

            Case "grdSOTORDRX"



            Case "grdSOTPICK0"

                tlb_btn = DirectCast(tlb_pop.Tools("Print Custom Tote Labels"), UltraWinToolbars.ButtonTool)
                Dim TRUCK_TYPE As String = grd.ActiveRow.Cells("TRUCK_TYPE").Value & ""
                tlb_btn.SharedProps.Visible = (TRUCK_TYPE = "X")

                tlb_btn = DirectCast(tlb_pop.Tools("Print Resolution Report"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "R")
                tlb_btn = DirectCast(tlb_pop.Tools("Print Resolution Report - Res Only"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "R")

                tlb_btn = DirectCast(tlb_pop.Tools("De-Release Pick Batch"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKS") And Not InquiryMode And Not chkShowHolds.Checked

            Case "grdSOTPICK1"

                tlb_btn = DirectCast(tlb_pop.Tools("De-Release Pick Tickets"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKS") And Not InquiryMode And Not chkShowHolds.Checked

                tlb_btn = DirectCast(tlb_pop.Tools("De-Pick Pick Tickets"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKS" And ASCMAIN1.USER_SECURITY_CODEs.Contains("FY")) And Not InquiryMode And Not chkShowHolds.Checked

            Case "grdSOTORDQ0"

                'tlb_btn = DirectCast(tlb_pop.Tools("Release Selected Orders"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKS") And Not InquiryMode And Not chkShowHolds.Checked

                tlb_btn = DirectCast(tlb_pop.Tools("Release Selected Order Groups"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKS") And Not InquiryMode And Not chkShowHolds.Checked

                tlb_btn = DirectCast(tlb_pop.Tools("Combine Groups"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKS") And Not InquiryMode And Not chkShowHolds.Checked

            Case "grdSOTORDQ1"

                tlb_btn = DirectCast(tlb_pop.Tools("Release Selected Orders"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKS") And Not InquiryMode And Not chkShowHolds.Checked

            Case "grdICTSTATO"

                tlb_btn = DirectCast(tlb_pop.Tools("De-Select Order"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.Band.Index = 1)

        End Select


    End Sub

    Private Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles tlb.ToolClick

        Dim grd As UltraWinGrid.UltraGrid = Nothing

        grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Select Selected"

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = "1"
                        grow.Update()
                    End If
                Next

            Case "Cancel Selected Orders"

                Dim numSelected As Int32 = dst.Tables("SOTORDRX").Select("SEL = '1'").Length
                If numSelected = 0 Then
                    MessageBox.Show("There are no Selected Orders to Cancel.", "Cancel Selected Orders", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                If MessageBox.Show($"Do you want to Cancel the {numSelected} {If(numSelected = 1, "selected order", "selected orders")}?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                Stop ' CancelSelectedOrders()


            Case "De-Pick Pick Tickets"

                If grdSOTPICK1.Selected.Rows.Count = 0 AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    grd.Selected.Rows.Clear()
                    grd.ActiveRow.Selected = True
                End If

                If grdSOTPICK1.Selected.Rows.Count <> 1 Then
                    MessageBox.Show("You may De-Pick only 1 Pick Ticket at a time", "De-Pick Pick Tickets", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim ORDR_NO As String = grdSOTPICK1.Selected.Rows(0).Cells("ORDR_NO").Value
                Dim PICK_NO As String = grdSOTPICK1.Selected.Rows(0).Cells("PICK_NO").Value

                If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub
                If Not ASCMAIN1.Logical_Lock("SOTPICK1", ORDR_NO) Then Exit Sub

                Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", PICK_NO)
                If rowSOTPICK1.Item("PICK_STATUS") & "" <> "P" Then
                    MessageBox.Show($"Pick Status of Pick Ticket {PICK_NO} is not 'In Pick' at this time - you must refresh", "De-Pick Pick Tickets", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                Dim PICK_BATCH_NO As String = rowSOTPICK1.Item("PICK_BATCH_NO") & ""
                Dim rowSOTPICK0 As DataRow = LookUp("SOTPICK0", PICK_BATCH_NO)
                If Not "NKP".Contains(rowSOTPICK0.Item("PICK_BATCH_STATUS") & "") Then
                    ' note to ABS - is this a real requirement?
                    MessageBox.Show($"Pick Status Of Pick Batch {PICK_BATCH_NO} Is Not 'Picked' - please call ABS", "De-Pick Pick Tickets", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                ASCMAIN1.sql = $"SELECT PICK_NO, SUM (PICK_QTY) PICK_QTY, SUM (PICK_QTY_CONF) PICK_QTY_CONF, SUM (STAT4) STAT4 FROM (
                    SELECT PICK_NO, SUM (PICK_QTY) PICK_QTY, SUM (PICK_QTY_CONF) PICK_QTY_CONF, 0 STAT4 FROM SOTPICK2 WHERE PICK_NO = '{PICK_NO}' GROUP BY PICK_NO
                    UNION
                    SELECT TRAN_REF PICK_NO, 0 PICK_QTY, 0 PICK_QTY_CONF, SUM (-1 * TRAN_QTY) STAT4 FROM WHTLOCB2 WHERE TRAN_REF = '{PICK_NO}' AND TRAN_TYPE = 'K' GROUP BY TRAN_REF
                    ) GROUP BY PICK_NO"

                Dim ROW As DataRow = ASCDATA1.GetDataRow
                If Val(ROW.Item("PICK_QTY") & "") <> Val(ROW.Item("PICK_QTY_CONF") & "") Or Val(ROW.Item("PICK_QTY") & "") <> Val(ROW.Item("STAT4") & "") Then

                    ' 10/15/2024 This is a second check - did not want to replace the above original query just in case we break something.
                    ASCMAIN1.sql = "SELECT PICK_NO, SUM (PICK_QTY) PICK_QTY, SUM (PICK_QTY_CONF) PICK_QTY_CONF, SUM (STAT4) STAT4 FROM (
                        SELECT PICK_NO, SUM (PICK_QTY - NVL(PICK_QTY_BACK, 0) - NVL(PICK_QTY_CANC, 0)) PICK_QTY, SUM (PICK_QTY_CONF) PICK_QTY_CONF, 0 STAT4 FROM SOTPICK2 WHERE PICK_NO = :PARM1 GROUP BY PICK_NO
                        UNION
                        SELECT TRAN_REF PICK_NO, 0 PICK_QTY, 0 PICK_QTY_CONF, SUM ( CASE WHEN LOC_CODE = 'TRUCK' THEN TRAN_QTY ELSE -1 * TRAN_QTY END) STAT4 FROM WHTLOCB2 WHERE TRAN_REF = :PARM1 AND TRAN_TYPE = 'K' GROUP BY TRAN_REF
                        ) GROUP BY PICK_NO"

                    ROW = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", {PICK_NO})

                    If Val(ROW.Item("PICK_QTY") & "") <> Val(ROW.Item("PICK_QTY_CONF") & "") Or Val(ROW.Item("PICK_QTY") & "") <> Val(ROW.Item("STAT4") & "") Then
                        MessageBox.Show($"Pick Qty ({Val(ROW.Item("PICK_QTY") & "")}, Confirm Qty {Val(ROW.Item("PICK_QTY_CONF") & "")} and Transaction Qty {Val(ROW.Item("STAT4") & "")}) are not equivalent for Pick Ticket {PICK_NO}", "De-Pick Pick Tickets", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                End If

                If MsgBox($"OK to De-Pick Pick Ticket {PICK_NO}?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                Try
                    BeginTrans()

                    ASCMAIN1.sql = $"BEGIN DECLARE CURSOR C0 IS
                                    SELECT PICK_NO, PICK_BATCH_NO FROM SOTPICK1 WHERE PICK_NO = '{PICK_NO}';
                                    BEGIN FOR R0 IN C0 LOOP
                                    BEGIN DECLARE CURSOR C1 IS
                                    SELECT SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, WHTLOCB2.WHSE_CODE, WHTLOCB2.LOCATION_CODE, WHTLOCB2.LOCATION_CODE_OTHER, WHTLOCB2.TRAN_QTY
                                    , WHTLOCB2.TRAN_TYPE, SOTPICK2.PICK_NO, SOTPICK2.PICK_LNO, WHTLOCB2.INIT_OPER
                                    FROM WHTLOCB2, SOTPICK2, SOTORDR2
                                     WHERE SOTPICK2.PICK_NO = WHTLOCB2.TRAN_REF
                                    AND SOTPICK2.PICK_LNO = WHTLOCB2.TRAN_REF_LNO
                                    AND WHTLOCB2.TRAN_REF = R0.PICK_NO
                                    AND SOTORDR2.ORDR_NO= SOTPICK2.ORDR_NO
                                    AND SOTORDR2.ORDR_LNO= SOTPICK2.ORDR_LNO;
                                    BEGIN FOR R1 IN C1 LOOP
                                        ICPSTAT3
                                        (R1.STYLE_CODE, R1.COLOR_CODE, R1.WHSE_CODE, NULL, NULL,  R1.LOC_CODE, R1.LOC_CODE_OTHER, R1.TRAN_QTY, R1.TRAN_TYPE, R1.PICK_NO, R1.PICK_LNO, R1.INIT_OPER);
                                        UPDATE SOTPICK2 SET PICK_QTY_CONF = GREATEST(0, NVL(PICK_QTY_CONF,0) + NVL(R1.TRAN_QTY,0))
                                        WHERE PICK_NO = R1.PICK_NO AND PICK_LNO = R1.PICK_LNO;
                                        END LOOP; END;
                                    END;
                                    END LOOP; END; END;"
                    ASCDATA1.ExecuteSQL()

                    Dim PICK_BATCH_STATUS_new As String = Get_PICK_BATCH_STATUS_new(PICK_BATCH_NO)

                    ASCMAIN1.sql = $"Update SOTPICK0 Set PICK_BATCH_STATUS = :PARM1, LAST_DATE = SYSDATE, LAST_OPER = :PARM2 where PICK_BATCH_NO = :PARM3"
                    Dim r As Int32 = -1
                    r = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New String() {PICK_BATCH_STATUS_new, ASCMAIN1.USER_ID, PICK_BATCH_NO})
                    If r <> 1 Then
                        Throw New Exception($"Could not Reset Pick Batch Status to {PICK_BATCH_STATUS_new} for Pick Batch No {PICK_BATCH_NO}")
                    End If

                    TAC.TACMAIN1.Record_Event("SOTPICK1", PICK_NO, Now, ASCMAIN1.USER_ID, "DEPICK", "De-Pick PT", "", Me.Name)

                    CommitTrans()

                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

                ASCMAIN1.MultiTask_Release()

                MsgBox($"Pick Ticket {PICK_NO} has been De-Picked" & vbCrLf & vbCrLf & "IMPORTANT - Return the lenses and cloths to their original Pick Location", MsgBoxStyle.OkOnly, "Success")

                refresh_next_time_we_look_at_orders = True
                '   tabMain.SelectedTab = tabMain.Tabs("Orders")
                Click_Command("Refresh")

            Case "De-Release Pick Batch", "De-Release Pick Tickets"

                If grdSOTPICK0.ActiveRow Is Nothing Then Exit Sub
                Dim PICK_BATCH_NO As String = grdSOTPICK0.ActiveRow.Cells("PICK_BATCH_NO").Value

                Dim PICK_NOsD As New List(Of String)
                If e.Tool.Key = "De-Release Pick Batch" Then
                    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select($"PICK_BATCH_NO = '{PICK_BATCH_NO}'")
                        Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                        PICK_NOsD.Add(PICK_NO)
                    Next
                Else
                    If Not grd.ActiveRow.Selected Then
                        grd.Selected.Rows.Clear()
                        grd.ActiveRow.Selected = True
                    End If
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTPICK1.Selected.Rows
                        Dim PICK_NO As String = grow.Cells("PICK_NO").Value
                        Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                        If rowSOTPICK1.Item("PICK_BATCH_NO") <> PICK_BATCH_NO Then
                            MsgBox($"Pick Batch Mis-Match on Pick Ticket {PICK_NO}" & vbCrLf & $"Batch {PICK_BATCH_NO} vs {rowSOTPICK1.Item("PICK_BATCH_NO")}",
                           MsgBoxStyle.OkOnly, "Please contact ABS")
                            Exit Sub
                        End If
                        PICK_NOsD.Add(PICK_NO)
                    Next
                End If

                If Not ASCMAIN1.Logical_Lock("SOTPICK0", PICK_BATCH_NO) Then Exit Sub

                ASCMAIN1.sql = sqlSOTPICK0 & $"   And SOTPICK0.PICK_BATCH_NO = '{PICK_BATCH_NO}'"
                Dim rowSOTPICK0 As DataRow = ASCDATA1.GetDataRow

                Dim PICK_BATCH_STATUS As String = rowSOTPICK0.Item("PICK_BATCH_STATUS")
                If Not (PICK_BATCH_STATUS = "O" Or PICK_BATCH_STATUS = "P") Then
                    MsgBox($"Status for Pick Batch {PICK_BATCH_NO} is not 'Released' or 'Picking'", MsgBoxStyle.OkOnly, "Cannot Proceed with De-Release")
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                For Each PICK_NO As String In PICK_NOsD
                    Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", PICK_NO)
                    Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                    If Not ASCMAIN1.Logical_Lock("SOTPICK1", PICK_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub

                    If rowSOTPICK1.Item("PICK_STATUS") <> "P" And rowSOTPICK1.Item("PICK_STATUS") <> "C" And rowSOTPICK1.Item("PICK_STATUS") <> "D" Then
                        MsgBox($"Pick Ticket {PICK_NO} has already been Picked", MsgBoxStyle.OkOnly, "Cannot Proceed with De-Release")
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                Next

                ASCMAIN1.sql = $"Select SOTPICK2.PICK_NO
                    , MAX (SOTPICK2.WH_TRAN_NO) WH_TRAN_NO
                    , SUM (NVL(SOTPICK2.PICK_QTY_CONF,0)) PICK_QTY_CONF
                     from SOTPICK2 
                    where SOTPICK2.PICK_NO in ('{Join(PICK_NOsD.ToArray, "','")}') group by SOTPICK2.PICK_NO"
                ASCMAIN1.sql = $"Select PICK_NO From ({ASCMAIN1.sql}) X where X.PICK_QTY_CONF <> 0 or X.WH_TRAN_NO is Not NULL"
                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
                If row IsNot Nothing Then
                    Dim PICK_NO As String = row.Item("PICK_NO")
                    MsgBox($"Pick Ticket {PICK_NO} has already been Picked", MsgBoxStyle.OkOnly, "Cannot Proceed with De-Release")
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                If MsgBox($"OK to De-Release {CStr(PICK_NOsD.Count)} Pick Tickets from Pick Batch {PICK_BATCH_NO}", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                Dim didDeReleaase As Boolean = De_Release(PICK_BATCH_NO, rowSOTPICK0, PICK_NOsD)

                ASCMAIN1.MultiTask_Release()

                If didDeReleaase Then
                    refresh_next_time_we_look_at_orders = True
                    '   tabMain.SelectedTab = tabMain.Tabs("Orders")
                End If
                Click_Command("Refresh")


            Case "Calculate Short"
                Calculate_Short()

            Case "Show Inventory Requirements"
                If grdSOTORDQ0.ActiveRow Is Nothing Then Exit Sub

                Dim GROUP_KEY As String = grdSOTORDQ0.ActiveRow.Cells("GROUP_KEY").Value
                If dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' AND SEL='1'").Length = 0 Then
                    MsgBox("No Orders Selected", MsgBoxStyle.OkOnly, "Cannot Determine Inventory Requirements")
                    Exit Sub
                End If
                Prepare_Inventory_Requirements(New String() {GROUP_KEY}, False)

            Case "Release Selected Orders"

                If grdSOTORDQ0.ActiveRow Is Nothing Or InquiryMode Then Exit Sub

                Dim GROUP_KEY As String = grdSOTORDQ0.ActiveRow.Cells("GROUP_KEY").Value

                Release_Selected_Orders(GROUP_KEY)


                'Dim ORDRs_Selected As Integer = dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' AND SEL='1'").Length
                ''Dim ORDRs_Selected As Integer = dst.Tables("SOTORDQ1").Select($"SEL='1'").Length
                'If ORDRs_Selected = 0 Then
                '    MsgBox("No Orders Selected", MsgBoxStyle.OkOnly, "Cannot Release")
                '    Exit Sub
                'ElseIf ORDRs_Selected > MAX_ORDERS_TO_RELEASE Then
                '    MsgBox($"Max number of Orders Permitted to be Released to a Truck is {MAX_ORDERS_TO_RELEASE}", MsgBoxStyle.OkOnly, "Cannot Release")
                '    Exit Sub
                'End If

                'ORDR_NOs_Tried.Clear()

                'Release_Orders(New String() {GROUP_KEY})


            Case "Release Selected Order Groups"

                dst.Tables("ICTSTATZ").Rows.Clear()

                If grdSOTORDQ0.ActiveRow Is Nothing Or InquiryMode Then Exit Sub
                If grdSOTORDQ0.Selected.Rows.Count = 0 And grdSOTORDQ0.ActiveRow.IsDataRow And Not grdSOTORDQ0.ActiveRow.IsFilterRow Then
                    grdSOTORDQ0.ActiveRow.Selected = True
                End If
                If grdSOTORDQ0.Selected.Rows.Count = 0 Then
                    Exit Sub
                End If

                Dim GROUP_KEYs As New List(Of String)

                If grdSOTORDQ0.Selected.Rows.Count = 1 Then
                    Dim GROUP_KEY As String = grdSOTORDQ0.Selected.Rows(0).Cells("GROUP_KEY").Value
                    'GROUP_KEYs.Add(GROUP_KEY)
                    'Release_Orders(GROUP_KEYs.ToArray, , True)
                    Select_Orders(MAX_ORDERS_TO_RELEASE, True)
                    Release_Selected_Orders(GROUP_KEY)

                Else
                    MsgBox("Select a Singe Order Group", MsgBoxStyle.OkOnly, "Cannot Release Multiple Order Groups - combine them first")
                End If

                'If GROUP_KEYs.Count > 0 Then
                '    Release_Orders(GROUP_KEYs.ToArray, , True)
                'End If

            Case "Print Resolution Report"
                Print_Report()

            Case "Print Resolution Report - Res Only"
                Print_Report(True)

            Case "Cancel if Short - All Items", "Back-Order if Short - All Items"

                dst.Tables("ICTSTATZ").Rows.Clear()
                Dim ACTION_IF_SHORT As String = IIf(e.Tool.Key = "Cancel if Short - All Items", "C", "B")
                For Each rowICTSTATO As DataRow In dst.Tables("ICTSTATO").Select("QTY_SHORT < 0")
                    rowICTSTATO.Item("ACTION_IF_SHORT") = ACTION_IF_SHORT

                    Dim STYLE_CODE As String = rowICTSTATO.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowICTSTATO.Item("COLOR_CODE")
                    Dim row As DataRow = dst.Tables("ICTSTATZ").Rows.Add(New String() {STYLE_CODE, COLOR_CODE})
                    row.Item("ACTION_IF_SHORT") = ACTION_IF_SHORT

                Next

            Case "Combine Groups"

                Dim GROUP_KEY_CONS As String = ""

                EnforceConstraints(False)

                Dim errors As String = ""
                Dim GROUP_KEYs As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDQ0.Selected.Rows
                    Dim GROUP_KEY As String = grow.Cells("GROUP_KEY").Value
                    GROUP_KEYs.Add(GROUP_KEY)
                Next

                For Each GROUP_KEY As String In GROUP_KEYs

                    Dim errorGROUP_KEY As Boolean = False
                    Dim row As DataRow = dst.Tables("SOTORDQ0").Rows.Find(GROUP_KEY)

                    If GROUP_KEY_CONS = "" Then
                        GROUP_KEY_CONS = ASCMAIN1.Next_Control_No("SOTORDQ0.GROUP_KEY_CONS")
                        row.Item("GROUP_KEY") = GROUP_KEY_CONS
                        row.Item("PICK_DESCRIPTION") = "Combined Group"
                        row.Item("SHIP_VIA_CODE") = ""
                    Else
                        Dim row0 As DataRow = dst.Tables("SOTORDQ0").Rows.Find(GROUP_KEY_CONS)

                        If row.Item("ORDR_TYPE_CODE") & "" <> row0.Item("ORDR_TYPE_CODE") & "" _
                        Or row.Item("DESTINATION") & "" <> row0.Item("DESTINATION") & "" _
                        Then
                            errors &= vbCrLf & row.Item("PICK_DESCRIPTION")
                            errorGROUP_KEY = True
                        Else
                            row0.Item("ORDR_CNT") = Val(row0.Item("ORDR_CNT") & "") + Val(row.Item("ORDR_CNT") & "")
                            row0.Item("ORDR_QTY_OPEN") = Val(row0.Item("ORDR_QTY_OPEN") & "") + Val(row.Item("ORDR_QTY_OPEN") & "")
                            row.Delete()
                        End If
                    End If

                    If Not errorGROUP_KEY Then
                        For Each row1 As DataRow In dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}'")
                            row1.Item("GROUP_KEY") = GROUP_KEY_CONS
                        Next
                    End If
                Next

                If errors <> "" Then
                    MsgBox("Could Not Combine the following:" & errors, MsgBoxStyle.OkOnly, "Cannot Combine Groups with different Types, Labs, or Dests")
                End If


                EnforceConstraints(True)

                grdSOTORDQ0.Selected.Rows.Clear()

                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDQ0.Rows
                    If grow.Cells("GROUP_KEY").Value = GROUP_KEY_CONS Then
                        grow.Activate()
                        Exit For
                    End If
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If


        Select Case e.Tool.Key

            Case "Print Custom Tote Labels"

                Dim TRUCK_NO As String = grd.ActiveRow.Cells("TRUCK_NO").Value & ""
                Dim PICK_BATCH_NO As String = grdSOTPICK0.ActiveRow.Cells("PICK_BATCH_NO").Value
                Print_Custom_Tote_Labels(e.Tool.Key, PICK_BATCH_NO, TRUCK_NO)
                MsgBox("Labels Printed")

            Case "Print Custom Tote Label"
                Dim TRUCK_NO As String = grdSOTPICK0.ActiveRow.Cells("TRUCK_NO").Value & ""
                Dim PICK_BATCH_NO As String = grdSOTPICK0.ActiveRow.Cells("PICK_BATCH_NO").Value
                Dim TOTE_NO As String = grd.ActiveRow.Cells("TOTE_NO").Value & ""
                Print_Custom_Tote_Labels(e.Tool.Key, PICK_BATCH_NO, TRUCK_NO, TOTE_NO)

            Case "Print Truck Pick Tag"

                Dim TRUCK_NO As String = grd.ActiveRow.Cells("TRUCK_NO").Value & ""
                Dim PICK_DESCRIPTION As String = grd.ActiveRow.Cells("PICK_DESCRIPTION").Value & ""
                Dim PICK_BATCH_NO As String = grd.ActiveRow.Cells("PICK_BATCH_NO").Value & ""
                Dim ORDER_COUNT As Integer = dst.Tables("SOTPICK1").Rows.Count
                Print_Truck_Pick_Tag(e.Tool.Key, TRUCK_NO, PICK_DESCRIPTION, ORDER_COUNT, PICK_BATCH_NO)
                MsgBox("Tag Printed")

            Case "Item/Location Inquiry"

                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim COLOR_CODE As String = grd.ActiveRow.Cells("COLOR_CODE").Text
                Dim KEYS As New Dictionary(Of String, Object)
                KEYS.Add("WHSE_CODE", WHSE_CODE)
                KEYS.Add("STYLE_CODE", STYLE_CODE)
                KEYS.Add("COLOR_CODE", COLOR_CODE)
                Context_Launch("View", KEYS, e.Tool.Key, "ICFSTAT3")

            Case "Sales Order Inquiry"

                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                'Dim KEYS As New Dictionary(Of String, Object)
                'KEYS.Add("ORDR_NO", ORDR_NO)
                'Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDSI")
                ' Context_Launch("Load", KEYS, e.Tool.Key, "SOFORDFI")
                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")

            Case "Print Test Report"
                Dim PICK_BATCH_NO As String = grd.ActiveRow.Cells("PICK_BATCH_NO").Text
                Print_Test_Report(PICK_BATCH_NO)

            Case "De-Select Order"

                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rows() As DataRow = dst.Tables("SOTORDQ1").Select($"ORDR_NO = '{ORDR_NO}'")
                If rows.Length = 1 Then
                    Dim row As DataRow = rows(0)
                    row.Item("SEL") = "0"
                End If

            Case "Find in Orders"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim row1() As DataRow = dst.Tables("SOTORDQ1").Select($"ORDR_NO = '{ORDR_NO}'")
                If row1.Length = 1 Then
                    Dim GROUP_KEY As String = row1(0).Item("GROUP_KEY")
                    For Each grow0 As UltraWinGrid.UltraGridRow In grdSOTORDQ0.Rows
                        If grow0.Cells("GROUP_KEY").Value & "" = GROUP_KEY Then
                            grow0.Activate()
                            For Each grow1 As UltraWinGrid.UltraGridRow In grdSOTORDQ1.Rows
                                If grow1.Cells("ORDR_NO").Value & "" = ORDR_NO Then
                                    grow1.Activate()
                                    Exit For
                                End If
                            Next
                            tabMain.Tabs("Orders").Selected = True
                            Exit For
                        End If
                    Next
                End If
        End Select


    End Sub

#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If Not ScreenMode Then
                    If e.KeyCode = Windows.Forms.Keys.Enter Then
                        Click_Command("Load", e)
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If Absx1.txtFor("WHSE_CODE").Text <> "" Then
                    Click_Command("Load")
                End If

        End Select
    End Sub
#End Region

    Sub Setup_Screen()

        Create_WorkTable(False)

        SplitContainer1.Panel2Collapsed = True
        grdSOTORDQ0.Rows.ColumnFilters.ClearAllFilters()

        EnforceConstraints(False)
        Fill_Records("SOTORDQ1")
        Fill_Records("SOTORDQ0")


        Fill_Records("ICTSTATX")
        Sort_grdColumns(grdICTSTATX, "STYLE_CODE,COLOR_CODE")

        Fetch_All_Orders()

        For Each grow As UltraGridRow In grdSOTORDQ0.Rows

            Dim PICK_DESCRIPTION As String = ""

            If grow.Cells("SHIP_VIA_CODE").Text <> "" Then
                Dim SHIP_VIA_CODE As String = grow.Cells("SHIP_VIA_CODE").Text
                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                Dim SHIP_VIA_DESC As String = "Via " & SHIP_VIA_CODE
                If rowSOTSVIA1 IsNot Nothing Then
                    SHIP_VIA_DESC = rowSOTSVIA1.Item("SHIP_VIA_DESC") & ""
                End If
                PICK_DESCRIPTION &= " " & SHIP_VIA_DESC
            End If

            PICK_DESCRIPTION = Trim(PICK_DESCRIPTION)
            grow.Cells("PICK_DESCRIPTION").Value = PICK_DESCRIPTION
            grow.Update()
        Next

        Sort_grdColumns(grdSOTORDQ0, "ORDR_DATE".ToLower)
        Setup_grdSOTORDQ1()

        EnforceConstraints(False)

        Fill_Records("SOTPICK0")
        Sort_grdColumns(grdSOTPICK0, "PICK_BATCH_NO".ToLower)

        Setup_grdSOTPICK1()
        Setup_grdSOTPICK2()

        EnforceConstraints(True)

        'Setup_grdSOTPICK1()
        EnforceConstraints(True)

        chkIgnoreInvtyShort.Checked = False
    End Sub

    Sub Create_WorkTable(initialize As Boolean)

        Dim inquirySingleSalesOrder As Boolean = InquiryMode AndAlso txtORDR_NO.TextLength > 0
        ' NEED TO FIND THE ALTERNATIVE FOR DETFRMSO TO GET THE ORDR_XFR_BATCH_NO
        If initialize Then
            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE, ROWNUM ORDR_CNT, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_PICK ORDR_QTY_ALLO, SOTORDR2.ORDR_QTY_PICK ORDR_QTY_ALLO_NC, SOTORDR2.ORDR_QTY_OPEN QTY_AVA" & vbCrLf _
                & ", ICTSTYLD.VOL_INDEX * 1 * 1 VOL_INDEX_TOT, SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_INV_COMMENT" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_BUYER_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                & ", SOTORDR1.ORDR_PICK_SEQ, SOTORDR1.SHIP_VIA_CODE" & vbCrLf _
                & ", 'X' DESTINATION" & vbCrLf _
                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ICTSTYC1.UPC_CODE, ICTSTYL1.STYLE_DESC, 'X' ORDR_XFR_BATCH_NO" & vbCrLf _
                & ", SOTORDR1.INIT_DATE, SOTORDR1.ORDR_HOLD, SOTORDR1.ORDR_SHIP_COMPLETE, SOTORDR1.ORDR_SHIP_COMPLETE ORDR_ALLO_COMPLETE, '0' BO" & vbCrLf _
                & " from SOTORDR1,SOTORDR2,ICTSTYL1,ICTSTYC1,ICTSTYLD" & vbCrLf _
                & " where ROWNUM < 1"
            SOTORDQ1 = ASCMAIN1.Temp_Table(sqlGK(ASCMAIN1.sql))
            ASCDATA1.ExecuteSQL($"Alter Table {SOTORDQ1} Add Primary Key (ORDR_NO)")

            ASCMAIN1.sql = "Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_PICK" & vbCrLf _
                & ", WHTLOCB1.LOCATION_QTY LOC_QTY_NOT_AVA, ICTSTAT2.WHSE_QTY_ON_HAND WHSE_QTY_ALLO" & vbCrLf _
                & ", 0 ORDERS, 0 ORDERS_ALLO, ICTSTAT2.WHSE_QTY_OPEN ORDR_QTY_OPEN, ICTSTAT2.WHSE_QTY_OPEN ORDR_QTY_BACK" & vbCrLf _
                & " from ICTSTAT2, WHTLOCB1" & vbCrLf _
                & " where ROWNUM < 1"
            ICTSTATX = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL($"Alter Table {ICTSTATX} Add Primary Key (STYLE_CODE,COLOR_CODE)")

            sqlSOTORDQ0 = "Select GROUP_KEY, WHSE_CODE, ORDR_TYPE_CODE, SHIP_VIA_CODE" & vbCrLf _
                & ", DESTINATION, ORDR_SOURCE, ORDR_XFR_BATCH_NO" & vbCrLf _
                & ", SUM (ORDR_CNT) ORDR_CNT, SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (ORDR_QTY_PICK) ORDR_QTY_PICK, SUM (ORDR_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
                & ", MIN (ORDR_DATE) ORDR_DATE" & vbCrLf _
                & $"from {SOTORDQ1}" & vbCrLf _
                & " group by GROUP_KEY, WHSE_CODE, ORDR_TYPE_CODE, SHIP_VIA_CODE" & vbCrLf _
                & ", DESTINATION, ORDR_SOURCE, ORDR_XFR_BATCH_NO"
            SOTORDQ0 = ASCMAIN1.Temp_Table(sqlSOTORDQ0)
            ASCDATA1.ExecuteSQL($"ALTER TABLE {SOTORDQ0} ADD PRIMARY KEY (GROUP_KEY)")

            sqlSOTORDRX = $"
                Select Z.*, ICTSTYC1.UPC_CODE, ICTSTYL1.STYLE_STATUS, 
                ICTSTAT2.WHSE_QTY_ON_HAND ONHD, ICTSTAT2.WHSE_QTY_ON_ORDER ONPO, ICTSTAT2.WHSE_QTY_PICK PICK            
                , NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) QAVA 
                    from ICTSTAT2, ICTSTYC1, ICTSTYL1, (
                Select SOTORDR1.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_CUST_PO, 
                SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_SOURCE,
                SOTORDR2.ORDR_LNO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_BUYER_NAME, 
                SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, ORDR_QTY_OPEN, ORDR_QTY_PICK, SOTORDR2.STYLE_CLASS_CODE 
                , SOTORDR1.INIT_DATE, 0 ORDR_QTY_ALLO, SOTORDR1.ORDR_HOLD
                , SOTORDR1.SHIP_VIA_CODE, SOTSVIA1.SHIP_VIA_DESC
                , SOTORDR1.ORDR_SHIP_COMPLETE, SOTORDR1.ORDR_SHIP_COMPLETE ORDR_ALLO_COMPLETE
                    from SOTORDR1,SOTORDR2,SOTSVIA1 
                where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO 
                  and SOTSVIA1.SHIP_VIA_CODE (+) = SOTORDR1.SHIP_VIA_CODE 
                  and SOTORDR1.WHSE_CODE = '??' AND SOTORDR1.SALES_DIVISION_CODE = '{SALES_DIVISION_CODE_SKIN}'
                  and SOTORDR1.ORDR_STATUS = 'O' 
                  AND SOTORDR1.ORDR_SOURCE = 'W'
                  and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0)
                ) Z 
                where ICTSTYL1.STYLE_CODE = Z.STYLE_CODE
                  and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE AND ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE
                  and ICTSTAT2.STYLE_CODE (+) = Z.STYLE_CODE AND ICTSTAT2.COLOR_CODE (+) = Z.COLOR_CODE AND ICTSTAT2.WHSE_CODE (+) = Z.WHSE_CODE"

            SOTORDRX = ASCMAIN1.Temp_Table(sqlSOTORDRX & " and ROWNUM < 1")
            ASCDATA1.ExecuteSQL($"ALTER TABLE {SOTORDRX} ADD PRIMARY KEY (ORDR_NO, ORDR_LNO)")

        Else

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Creating Worktables")

            Dim WHSE_CODEs As String = $"'{WHSE_CODE_SKIN}'"

            If tabMain.SelectedTab.Key = "Orders" Or tabMain.SelectedTab.Key = "Order Details" Then

                ' SOTORDR1

                ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDR1}")
                ASCMAIN1.sql = $"Select SOTORDR1.*, SOTORDR1.ORDR_SHIP_COMPLETE ORDR_ALLO_COMPLETE 
                    from SOTORDR1 
                    where WHSE_CODE = '{Absx1.txtFor("WHSE_CODE").Text}'
                    and SALES_DIVISION_CODE = '{SALES_DIVISION_CODE_SKIN}' and SOTORDR1.WHSE_CODE IN ({WHSE_CODEs}) AND SOTORDR1.ORDR_SOURCE = 'W' "

                If inquirySingleSalesOrder Then
                    ASCMAIN1.sql &= $" AND SOTORDR1.ORDR_NO = :PARM1 " ' why do we permit orders whose status might not be O?
                    ASCDATA1.ExecuteSQL($"Insert into {SOTORDR1} {ASCMAIN1.sql}", "V", New Object() {txtORDR_NO.Text})
                Else
                    ASCMAIN1.sql &= $" AND SOTORDR1.ORDR_STATUS = 'O'"
                    ASCDATA1.ExecuteSQL($"Insert into {SOTORDR1} {ASCMAIN1.sql}")
                End If

                'SOTORDRX

                ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDRX}")
                Dim SQLX As String = sqlSOTORDRX
                SQLX = Replace(SQLX, "SOTORDR1.WHSE_CODE = '??'", $"SOTORDR1.WHSE_CODE = '{Absx1.txtFor("WHSE_CODE").Text}'")
                SQLX = Replace(SQLX, "and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0", $"and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0")
                ASCDATA1.ExecuteSQL($"Insert into {SOTORDRX} " & SQLX)

                ' ICTSTATX

                ASCDATA1.ExecuteSQL($"Truncate Table {ICTSTATX}")
                ASCMAIN1.sql = "Select ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_PICK" & vbCrLf _
                    & ", SUM (WHTLOCB1.LOCATION_QTY) LOC_QTY_NOT_AVA, 0 WHSE_QTY_ALLO, X.ORDERS, 0 ORDERS_ALLO, X.ORDR_QTY_OPEN, X.ORDR_QTY_PICK" & vbCrLf _
                    & $" from ICTSTAT2, WHTLOCB1" & vbCrLf _
                    & $", (Select STYLE_CODE, COLOR_CODE, SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (ORDR_QTY_PICK) ORDR_QTY_PICK, COUNT (*) ORDERS from {SOTORDRX}" & vbCrLf _
                    & " where NVL(ORDR_QTY_OPEN,0) <> 0 group by STYLE_CODE, COLOR_CODE) X" & vbCrLf _
                    & " where ICTSTAT2.STYLE_CODE = X.STYLE_CODE and ICTSTAT2.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                    & $"   And ICTSTAT2.WHSE_CODE = '{WHSE_CODE_SKIN}'" & vbCrLf _
                    & "   and WHTLOCB1.STYLE_CODE (+) = ICTSTAT2.STYLE_CODE" & vbCrLf _
                    & "   and WHTLOCB1.COLOR_CODE (+) = ICTSTAT2.COLOR_CODE" & vbCrLf _
                    & $"   And WHTLOCB1.WHSE_CODE (+) = '{Absx1.txtFor("WHSE_CODE").Text}'" & vbCrLf _
                    & "   and (WHTLOCB1.LOCATION_CODE (+) = 'LOST')" & vbCrLf _
                    & " group by ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_PICK, X.ORDERS, X.ORDR_QTY_OPEN, X.ORDR_QTY_PICK"
                ASCDATA1.ExecuteSQL($"Insert into {ICTSTATX} {ASCMAIN1.sql}")

                ' & "   and (ICTSTAT3.LOC_CODE (+) = 'LOST' or ICTSTAT3.LOC_CODE (+) = 'PALLET')" & vbCrLf _
                ' MM/CD WANT TO SEE PALLET AS AVAILABLE TO ALLOCATE - 10/04/2023

                Calculate_Allocations()


                'SOTORDQ0,1

                ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDQ1}")
                ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDQ0}")

                Dim sqlICTSTAT2 As String = $"Select SOTORDR2.ORDR_NO" & vbCrLf _
                    & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                    & ", SUM (SOTORDRX.ORDR_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
                    & ", SUM (CASE WHEN NVL(SOTORDR2.STYLE_CLASS_CODE,'??') = 'CC' THEN SOTORDRX.ORDR_QTY_OPEN ELSE 0 END) ORDR_QTY_ALLO_NC" & vbCrLf _
                    & ", SUM (CASE WHEN NVL(ICTSTYLD.VOL_INDEX,1) = 1 THEN 5 ELSE CASE WHEN NVL(ICTSTYLD.VOL_INDEX,1) = 2 THEN 40 ELSE 110 END END) VOL_INDEX_TOT" & vbCrLf _
                    & $" from {SOTORDRX} SOTORDRX, SOTORDR2, ICTSTYLD " & vbCrLf _
                    & IIf(inquirySingleSalesOrder, $"where SOTORDR2.ORDR_NO = '{txtORDR_NO.Text}'", "where NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0 ") & vbCrLf _
                    & "  And SOTORDRX.ORDR_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
                    & "  And SOTORDRX.ORDR_LNO (+) = SOTORDR2.ORDR_LNO" & vbCrLf _
                    & "  And ICTSTYLD.STYLE_CODE (+) = SOTORDR2.STYLE_CODE" & vbCrLf _
                    & "  And ICTSTYLD.PACK_CODE (+) = 'EA'" & vbCrLf _
                    & " group by SOTORDR2.ORDR_NO"

                Dim sql1k As String = "Select SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE, 1 ORDR_CNT, X.ORDR_QTY_OPEN, X.ORDR_QTY_PICK" & vbCrLf _
                    & ", X.ORDR_QTY_ALLO, X.ORDR_QTY_ALLO_NC, 0 QTY_AVA, X.VOL_INDEX_TOT" & vbCrLf _
                    & ", SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_INV_COMMENT" & vbCrLf _
                    & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_BUYER_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                    & ", SOTORDR1.ORDR_PICK_SEQ, SOTORDR1.SHIP_VIA_CODE" & vbCrLf _
                    & ", 'C' DESTINATION" & vbCrLf _
                    & ", NULL STYLE_CODE, NULL COLOR_CODE, NULL UPC_CODE, NULL STYLE_DESC, NULL ORDR_XFR_BATCH_NO" & vbCrLf _
                    & ", SOTORDR1.INIT_DATE, SOTORDR1.ORDR_HOLD, SOTORDR1.ORDR_SHIP_COMPLETE, SOTORDR1.ORDR_ALLO_COMPLETE, '0' BO" & vbCrLf _
                    & $" from {SOTORDR1} SOTORDR1, ({sqlICTSTAT2}) X" & vbCrLf _
                    & " where SOTORDR1.ORDR_SOURCE in ('W')" & vbCrLf

                '& " where SOTORDR1.ORDR_SOURCE in ('K', 'C', 'E', 'L', 'W','P')" & vbCrLf

                If inquirySingleSalesOrder Then
                    sql1k &= $" AND SOTORDR1.ORDR_NO = '{txtORDR_NO.Text}'"
                Else
                    sql1k &= "   and SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf
                End If

                If chkShowHolds.Checked Then
                    sql1k &= "and (NVL(SOTORDR1.ORDR_HOLD,'0') = '1'" _
                & " or ((NVL(SOTORDR1.ORDR_SHIP_COMPLETE,'0') = '1' and NVL(SOTORDR1.ORDR_ALLO_COMPLETE,'0') = '0')))"
                Else
                    sql1k &= "and (NVL(SOTORDR1.ORDR_HOLD,'0') <> '1'" _
                & " and (NVL(SOTORDR1.ORDR_SHIP_COMPLETE,'0') = '0' or NVL(SOTORDR1.ORDR_ALLO_COMPLETE,'0') = '1'))"
                End If

                sql1k &= " and X.ORDR_NO = SOTORDR1.ORDR_NO"

                ASCDATA1.ExecuteSQL($"Insert into {SOTORDQ1} {sqlGK(sql1k)}")
                ASCDATA1.ExecuteSQL($"Insert into {SOTORDQ0} {sqlSOTORDQ0}")

                ASCMAIN1.sql = $"Update {SOTORDQ1} SOTORDQ1 Set BO = '1' where NVL(SOTORDQ1.ORDR_QTY_ALLO,0) <= 0 
                or not (NVL(SOTORDQ1.ORDR_QTY_ALLO,0) > SOTORDQ1.ORDR_QTY_ALLO_NC or NVL(SOTORDQ1.ORDR_QTY_OPEN,0) = NVL(SOTORDQ1.ORDR_QTY_ALLO,0))"
                ASCDATA1.ExecuteSQL()

            End If

            ASCDATA1.ExecuteSQL($"Truncate Table {SOTPICK1}")

            Dim sqlSOTPICK1 As String = "Select SOTPICK1.* from SOTPICK1, SOTPICK0" & vbCrLf

            If inquirySingleSalesOrder Then
                sqlSOTPICK1 &= " where SOTPICK0.PICK_BATCH_STATUS IN ('N','O','R','K','P','F')" & vbCrLf
                sqlSOTPICK1 &= $" AND SOTPICK1.ORDR_NO = '{txtORDR_NO.Text}'"
            Else
                sqlSOTPICK1 &= " where SOTPICK0.PICK_BATCH_STATUS IN ('N','O','R','K','P')" & vbCrLf
                sqlSOTPICK1 &= $" AND SOTPICK1.WHSE_CODE = '{WHSE_CODE}'"
            End If

            sqlSOTPICK1 &= " and SOTPICK1.PICK_BATCH_NO = SOTPICK0.PICK_BATCH_NO" & vbCrLf _
            & " and SOTPICK1.PICK_STATUS <> 'D'" & vbCrLf _
            & " and SOTPICK0.ORDR_SOURCE = 'W'" & vbCrLf

            If EntryMode = "R" Then
                sqlSOTPICK1 &= $" and SOTPICK1.PICK_BATCH_NO = '{PICK_BATCH_NO}'"
            End If
            ASCDATA1.ExecuteSQL($"Insert into {SOTPICK1} {sqlSOTPICK1}")

            grdSOTORDQ0.DisplayLayout.Bands(0).Columns("QTY_SHORT").Hidden = True

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Function sqlGK(sql As String) As String

        ' DO WE ALLOW K?

        'Return $"Select DECODE(X.ORDR_SOURCE,
        '            'K','K' || TO_CHAR(GREATEST(X.ORDR_DATE,SYSDATE-7),'YYYYMMDD') || X.WHSE_CODE || X.ORDR_TYPE_CODE || X.SHIP_VIA_CODE || X.DESTINATION,
        '            'W','W' || TO_CHAR(GREATEST(X.ORDR_DATE,SYSDATE-7),'YYYYMMDD') || X.WHSE_CODE || X.ORDR_TYPE_CODE || X.SHIP_VIA_CODE || X.DESTINATION,
        '             '?' || NVL(X.ORDR_TYPE_CODE,'')) GROUP_KEY, X.* from ({sql}) X"

        Return $"Select DECODE(X.ORDR_SOURCE,
                    'K','K' || TO_CHAR(GREATEST(X.ORDR_DATE,SYSDATE-7),'YYYYMMDD') || X.WHSE_CODE || X.ORDR_TYPE_CODE || X.SHIP_VIA_CODE || X.DESTINATION,
                    'W','W' || TO_CHAR(X.ORDR_DATE,'YYYYMMDD') || X.WHSE_CODE || X.ORDR_TYPE_CODE || X.SHIP_VIA_CODE || X.DESTINATION,
                     '?' || NVL(X.ORDR_TYPE_CODE,'')) GROUP_KEY, X.* from ({sql}) X"
    End Function

    Sub Calculate_Allocations()

        Dim QTY_TO_ALLO_calc As String = "NVL(R2.ORDR_QTY_OPEN,0)"

        ASCMAIN1.sql = $"
        Begin
          Declare 
            Cursor C1 is Select * from {SOTORDR1} SOTORDR1 order by INIT_DATE, ORDR_NO for Update;
          Begin   
            Update {SOTORDRX} Set ORDR_QTY_ALLO = 0;
            For R1 in C1 Loop
              Begin
                Declare 
                  RX {ICTSTATX}%ROWTYPE;
                  QTY_AVA NUMBER (6,0); 
                  QTY_TO_ALLO NUMBER (6,0); 
                  Cursor C2 is Select * from {SOTORDRX} SOTORDRX where ORDR_NO = R1.ORDR_NO order by INIT_DATE for Update;
                Begin         

                  For R2 in C2 Loop      
                    QTY_TO_ALLO := {QTY_TO_ALLO_calc};          
                    Select ICTSTATX.* into RX from {ICTSTATX} ICTSTATX where STYLE_CODE = R2.STYLE_CODE and COLOR_CODE = R2.COLOR_CODE;
                    QTY_AVA := NVL(RX.WHSE_QTY_ON_HAND,0) - NVL(RX.WHSE_QTY_PICK,0) - NVL (RX.LOC_QTY_NOT_AVA,0) - NVL(RX.WHSE_QTY_ALLO,0);

                    If QTY_TO_ALLO <= QTY_AVA Then
                      Update {SOTORDRX} Set ORDR_QTY_ALLO = QTY_TO_ALLO where Current of C2;
                      Update {ICTSTATX} Set WHSE_QTY_ALLO = NVL(WHSE_QTY_ALLO,0) + QTY_TO_ALLO, ORDERS_ALLO = NVL(ORDERS_ALLO,0) + 1 where STYLE_CODE = R2.STYLE_CODE and COLOR_CODE = R2.COLOR_CODE;
                    Else
                      If NVL(R1.ORDR_SHIP_COMPLETE,'0') = '1' Then
                        Update {SOTORDR1} Set ORDR_ALLO_COMPLETE = '0' where ORDR_NO = R1.ORDR_NO;                 
                      End If;
                    End If;     
                  End Loop;
                End;
              End;                      
            End Loop;
          End;
        End;"

        ASCDATA1.ExecuteSQL()

    End Sub


    Private Sub grdSOTORDQ0_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTORDQ0.AfterRowActivate
        Setup_grdSOTORDQ1()
        dst.Tables("ICTSTATZ").Rows.Clear()
    End Sub

    Sub Setup_grdSOTORDQ1()
        If grdSOTORDQ0.ActiveRow Is Nothing OrElse (grdSOTORDQ0.ActiveRow.IsFilterRow Or Not grdSOTORDQ0.ActiveRow.IsDataRow) Then
            grdSOTORDQ1.Visible = False
            numSEL.Visible = False
            lblSEL.Visible = False
            btnMaxRelease.Visible = False
        Else
            EnforceConstraints(False)
            dst.Tables("ICTSTATP").Rows.Clear()
            dst.Tables("ICTSTATO").Rows.Clear()
            EnforceConstraints(True)

            SplitContainer1.Panel2Collapsed = True

            grdSOTORDQ1.Text = $"Open Orders"

            Dim GROUP_KEY As String = grdSOTORDQ0.ActiveRow.Cells("GROUP_KEY").Value & ""

            Dim DVW As DataView = DirectCast(grdSOTORDQ1.DataSource, DataTable).DefaultView
            DVW.RowFilter = $"GROUP_KEY = '{GROUP_KEY}'"
            DVW.Sort = "ORDR_NO"
            grdSOTORDQ1.Visible = True

            numSEL.Visible = True
            lblSEL.Visible = True
            btnMaxRelease.Visible = True
        End If
    End Sub

    Private Sub numSEL_ValueChanged(sender As Object, e As EventArgs) Handles numSEL.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

    End Sub

    Private Sub numSEL_KeyDown(sender As Object, e As KeyEventArgs) Handles numSEL.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Dim NUM As Integer = Val(numSEL.Value & "")
            If NUM > 0 And NUM <= MAX_ORDERS_TO_RELEASE Then
                Select_Orders(NUM)
            End If
        End If
    End Sub

    Sub Select_Orders(NUM As Integer, Optional clear_all_previous_selections As Boolean = False)
        Dim SEL As Integer = 0
        Dim GROUP_KEY As String = grdSOTORDQ0.ActiveRow.Cells("GROUP_KEY").Value

        If clear_all_previous_selections Then
            For Each rowSOTORDQ1 As DataRow In dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' AND SEL= '1'", "ORDR_NO")
                rowSOTORDQ1.Item("SEL") = "0"
            Next
        End If

        For Each rowSOTORDQ1 As DataRow In dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' AND SEL= '0'", "ORDR_NO")
            rowSOTORDQ1.Item("SEL") = "1"
            SEL += 1
            If SEL = NUM Then
                Exit For
            End If
        Next
    End Sub

    Private Sub grdSOTORDQ1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTORDQ1.InitializeRow
        If e.Row.IsDataRow Then

            With e.Row.Cells("SEL")
                If .Value & "" = "1" Then
                    .Appearance = Appearance_Magenta
                Else
                    .Appearance = Appearance_Empty
                End If
            End With

            With e.Row.Cells("BO")
                If .Value & "" = "1" Then
                    .Appearance.BackColor = System.Drawing.Color.Red
                    '.Appearance = Appearance_Red
                Else
                    .Appearance.BackColor = System.Drawing.Color.Empty
                    '.Appearance = Appearance_Empty
                End If
            End With


            With e.Row.Cells("ORDR_QTY_ALLO")
                Dim ORDR_QTY_OPEN As Integer = Val(e.Row.Cells("ORDR_QTY_OPEN").Value & "")
                Dim ORDR_QTY_BACK As Integer = 0 ' Val(e.Row.Cells("ORDR_QTY_BACK").Value & "")
                Dim ORDR_QTY_ALLO As Integer = Val(e.Row.Cells("ORDR_QTY_ALLO").Value & "")
                If ORDR_QTY_ALLO <> ORDR_QTY_OPEN + ORDR_QTY_BACK Then
                    .Appearance.BackColor = System.Drawing.Color.Yellow
                Else
                    .Appearance.BackColor = System.Drawing.Color.Empty
                End If
            End With
        End If
    End Sub

    Private Sub grdSOTORDQ0_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTORDQ0.InitializeRow
        If e.Row.IsDataRow Then

            'If Not grdSOTORDQ0.DisplayLayout.Bands(0).Columns("QTY_SHORT").Hidden Then
            With e.Row.Cells("QTY_SHORT")

                If Val(.Value & "") <> 0 Then
                    'Dim ORDR_QTY_BACK As Int32 = Val(e.Row.Cells("ORDR_QTY_BACK").Value & "")
                    Dim ORDR_QTY_OPEN As Int32 = Val(e.Row.Cells("ORDR_QTY_OPEN").Value & "")
                    Dim QTY_SHORT As Int32 = Val(e.Row.Cells("QTY_SHORT").Value & "")
                    If ORDR_QTY_OPEN + QTY_SHORT > 0 Then
                        .Appearance.ForeColor = System.Drawing.Color.Empty
                        .Appearance.BackColor = System.Drawing.Color.Yellow
                        .ToolTipText = "Partial Release is Possible"
                    Else
                        .Appearance.ForeColor = System.Drawing.Color.Red
                        .Appearance.BackColor = System.Drawing.Color.Empty
                        .ToolTipText = "No Release is Possible"
                    End If
                Else
                    .Appearance.ForeColor = System.Drawing.Color.Empty
                    .Appearance.BackColor = System.Drawing.Color.Empty
                    .ToolTipText = ""
                End If

            End With
            'End If
        End If
    End Sub


    Private Sub grdICTSTATO_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTSTATO.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Band.Index = 0 Then
                With e.Row.Cells("QTY_SHORT")
                    If Val(.Value & "") < 0 Then
                        .Appearance = Appearance_Red
                    Else
                        .Appearance = Appearance_Empty
                    End If
                End With
            End If
        End If
    End Sub

    Private Sub grdSOTPICK0_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICK0.AfterRowActivate
        Setup_grdSOTPICK1()
    End Sub

    Sub Setup_grdSOTPICK1()

        splPicks.Panel2Collapsed = False
        splPicks.Panel1Collapsed = True

        If grdSOTPICK0.ActiveRow Is Nothing OrElse (grdSOTPICK0.ActiveRow.IsFilterRow Or Not grdSOTPICK0.ActiveRow.IsDataRow) Then
            grdSOTPICK1.Visible = False
        Else
            Dim PICK_BATCH_NO As String = grdSOTPICK0.ActiveRow.Cells("PICK_BATCH_NO").Value & ""

            grdSOTPICK1.Text = $"Pick Tickets in Pick Batch {PICK_BATCH_NO}"

            EnforceConstraints(False)
            dst.Tables("SOTPICK2").Rows.Clear()

            ASCMAIN1.sql = sqlSOTPICK2 & $" and SOTPICK1.PICK_BATCH_NO = '{PICK_BATCH_NO}'"
            'ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM1", $"'{WHSE_CODE}'")


            ' THIS IS THE ORIGINAL QTY/LOC SQL
            'Dim sqlLoc As String = "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
            '& ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = ICTSTYD1.LOCATION_CODE THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PRI" & vbCrLf _
            '& ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE LIKE 'OS%' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) OSL" & vbCrLf _
            '& ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'PALLET' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PAL" & vbCrLf _
            '& ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'CART' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) CRT" & vbCrLf _
            '& ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'TRUCK' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) TRK" & vbCrLf _
            '& ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'LOST' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) LNF" & vbCrLf _
            '& " from WHTLOCB1,ICTSTYD1" & vbCrLf _
            '& $" where WHTLOCB1.WHSE_CODE (+) = '{WHSE_CODE}'" & vbCrLf _
            '& "   and ICTSTYD1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE and ICTSTYD1.STYLE_CODE (+) = WHTLOCB1.STYLE_CODE and ICTSTYD1.COLOR_CODE (+) = WHTLOCB1.COLOR_CODE" & vbCrLf _
            '& " group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"

            ' THIS SQL ASSUMES THAT S2-%-01 AND S2-%-02 ARE PRIMARY, AND S2-%-03/04/05/06 ARE OSL
            Dim sqlLoc As String = "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE LIKE 'S2-%-01' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-02' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-01' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-02' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PRI" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE LIKE 'S2-%-03' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-04' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-05' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-06' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-03' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-04' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-05' OR WHTLOCB1.LOCATION_CODE LIKE 'S1%-06' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) OSL" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'PALLET' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PAL" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'CART' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) CRT" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'TRUCK' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) TRK" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'LOST' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) LNF" & vbCrLf _
            & " from WHTLOCB1,ICTSTYD1" & vbCrLf _
            & $" where WHTLOCB1.WHSE_CODE = '{WHSE_CODE}'" & vbCrLf _
            & "   and ICTSTYD1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE and ICTSTYD1.STYLE_CODE (+) = WHTLOCB1.STYLE_CODE and ICTSTYD1.COLOR_CODE (+) = WHTLOCB1.COLOR_CODE" & vbCrLf _
            & " group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"
            ASCMAIN1.sql = $"Select X.*, Z.PRI, Z.OSL, Z.PAL, Z.CRT, Z.TRK, Z.LNF from ({ASCMAIN1.sql}) X, ({sqlLoc}) Z where Z.STYLE_CODE (+) = X.STYLE_CODE and Z.COLOR_CODE (+) = X.COLOR_CODE"


            Fill_Records("SOTPICK2",,, ASCMAIN1.sql)


            ASCMAIN1.sql = sqlSOTPICK1 & $" and SOTPICK1.PICK_BATCH_NO = '{PICK_BATCH_NO}'"
            Fill_Records("SOTPICK1",,, ASCMAIN1.sql)
            EnforceConstraints(True)

            Sort_grdColumns(grdSOTPICK1, "PICK_NO")

            grdSOTPICK1.Visible = True

        End If
    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        If tabMain.SelectedTab.Key = "Pick Tickets" Then
            Setup_grdSOTPICK1()
        End If
        If tabMain.SelectedTab.Key = "Orders" Then
            If refresh_next_time_we_look_at_orders Then
                refresh_next_time_we_look_at_orders = False
                Click_Command("Refresh")
            End If
        End If

        UltraExplorerBar1.Groups("Screen Control").Items("Resolve").Visible = Not (EntryMode = "R") And tabMain.SelectedTab.Key = "Pick Tickets"

        UltraExplorerBar1.Groups("Order Details").Visible = False ' tabMain.SelectedTab.Key = "Order Details"

        chkShowHolds.Visible = (tabMain.SelectedTab.Key = "Orders")
    End Sub

    Function Prepare_Inventory_Requirements(GROUP_KEYs() As String, multi_task As Boolean, Optional automatic As Boolean = False, Optional release_multiple_groups As Boolean = False)

        Dim ORDR_NOs As New List(Of String)
        dst.Tables("SOTORDR0").Rows.Clear()
        ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDR0}")

        Dim ORDR_NOs_could_not_lock As New List(Of String)

        For Each GROUP_KEY As String In GROUP_KEYs
            Dim sqlw As String = $"GROUP_KEY = '{GROUP_KEY}'"
            If Not release_multiple_groups Then
                sqlw &= " AND SEL='1'"
            End If



            For Each rowSOTORDQ1 As DataRow In dst.Tables("SOTORDQ1").Select(sqlw)
                Dim ORDR_NO As String = rowSOTORDQ1.Item("ORDR_NO")
                If automatic And ORDR_NOs_Tried.Contains(ORDR_NO) Then
                Else
                    Dim add_order_to_queue As Boolean = True
                    ' IF WE CANNOT MT, THEN WE WILL RETURN FALSE IF RELEASING FROM grdSOTORDQ1, AND WE WILL AVOID ADDING THE ORDER IF RELEASING FROM grdSOTORDQ0
                    If multi_task Then
                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO,, Not release_multiple_groups) Then
                            If release_multiple_groups Then
                                add_order_to_queue = False
                                ORDR_NOs_could_not_lock.Add(ORDR_NO)
                            Else
                                Me.Cursor = Cursors.Default
                                ASCMAIN1.Progress("")
                                Return False
                            End If
                        End If
                    End If
                    If add_order_to_queue Then
                        ORDR_NOs.Add(ORDR_NO)
                        dst.Tables("SOTORDR0").Rows.Add(New String() {ORDR_NO})
                        ASCMAIN1.Progress("-", ORDR_NO)
                    End If
                End If
            Next
        Next

        If ORDR_NOs_could_not_lock.Count > 0 Then
            MsgBox("Could Not Lock some Orders" & vbCrLf & vbCrLf & Join(ORDR_NOs_could_not_lock.ToArray, ","), MsgBoxStyle.OkOnly, "Verification - FYI")
        End If

        Update_Record_TDA("SOTORDR0")

        EnforceConstraints(False)

        ASCMAIN1.sql = $"Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.ORDR_NO from SOTORDR2,SOTORDR1,{SOTORDR0} SOTORDR0" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & $"   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0)" & vbCrLf _
            & $"    and SOTORDR2.ORDR_NO = SOTORDR0.ORDR_NO group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.ORDR_NO"

        Fill_Records("ICTSTATP", , , ASCMAIN1.sql)

        ASCMAIN1.sql = $"Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDRX.ORDR_QTY_ALLO) ORDR_QTY_ALLO" & vbCrLf _
            & ", MAX (SOTORDR2.STYLE_DESC) STYLE_DESC, MAX (SOTORDR2.STYLE_CLASS_CODE) STYLE_CLASS_CODE" & vbCrLf _
            & $" from SOTORDR2,SOTORDR1,{SOTORDRX} SOTORDRX,{SOTORDR0} SOTORDR0" & vbCrLf _
            & $" where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & $"   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0)" & vbCrLf _
            & $"   and SOTORDR2.ORDR_NO = SOTORDR0.ORDR_NO" & vbCrLf _
            & "   and SOTORDRX.ORDR_NO (+) = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDRX.ORDR_LNO (+) = SOTORDR2.ORDR_LNO" & vbCrLf _
            & " group by SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR1.WHSE_CODE"

        ASCMAIN1.sql = $"Select X.STYLE_CODE, X.COLOR_CODE, X.ORDR_QTY_OPEN, X.ORDR_QTY_PICK, X.ORDR_QTY_ALLO" & vbCrLf _
            & $", ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_PICK, X.STYLE_DESC, X.STYLE_CLASS_CODE, ICTCOLR1.COLOR_DESC" & vbCrLf _
            & $" from ICTSTAT2, ICTCOLR1, ({ASCMAIN1.sql}) X" & vbCrLf _
            & " where ICTSTAT2.STYLE_CODE (+) = X.STYLE_CODE And ICTSTAT2.COLOR_CODE (+) = X.COLOR_CODE And ICTCOLR1.COLOR_CODE = X.COLOR_CODE And ICTSTAT2.WHSE_CODE (+) = X.WHSE_CODE"

        ' THIS IS THE ORIGINAL QTY/LOC SQL
        'Dim sqlLoc As String = "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
        '    & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = ICTSTYD1.LOCATION_CODE THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PRI" & vbCrLf _
        '    & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE Like 'OS%' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) OSL" & vbCrLf _
        '    & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'PALLET' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PAL" & vbCrLf _
        '    & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'CART' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) CRT" & vbCrLf _
        '    & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'TRUCK' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) TRK" & vbCrLf _
        '    & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'LOST' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) LNF" & vbCrLf _
        '    & " from WHTLOCB1,ICTSTYD1" & vbCrLf _
        '    & $" where WHTLOCB1.WHSE_CODE (+) = '{WHSE_CODE}'" & vbCrLf _
        '    & "   and ICTSTYD1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE and ICTSTYD1.STYLE_CODE (+) = WHTLOCB1.STYLE_CODE and ICTSTYD1.COLOR_CODE (+) = WHTLOCB1.COLOR_CODE" & vbCrLf _
        '    & " group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"

        ' THIS SQL ASSUMES THAT S2-%-01 AND S2-%-02 ARE PRIMARY, AND S2-%-03/04/05/06 ARE OSL
        Dim sqlLoc As String = "Select WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE LIKE 'S2-%-01' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-02' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-01' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-02' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PRI" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE LIKE 'S2-%-03' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-04' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-05' OR WHTLOCB1.LOCATION_CODE LIKE 'S2-%-06' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-03' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-04' OR WHTLOCB1.LOCATION_CODE LIKE 'S1-%-05' OR WHTLOCB1.LOCATION_CODE LIKE 'S1%-06' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) OSL" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'PALLET' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) PAL" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'CART' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) CRT" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'TRUCK' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) TRK" & vbCrLf _
            & ", SUM (CASE WHEN WHTLOCB1.LOCATION_CODE = 'LOST' THEN WHTLOCB1.LOCATION_QTY ELSE 0 END) LNF" & vbCrLf _
            & " from WHTLOCB1,ICTSTYD1" & vbCrLf _
            & $" where WHTLOCB1.WHSE_CODE (+) = '{WHSE_CODE}'" & vbCrLf _
            & "   and ICTSTYD1.WHSE_CODE (+) = WHTLOCB1.WHSE_CODE and ICTSTYD1.STYLE_CODE (+) = WHTLOCB1.STYLE_CODE and ICTSTYD1.COLOR_CODE (+) = WHTLOCB1.COLOR_CODE" & vbCrLf _
            & " group by WHTLOCB1.STYLE_CODE, WHTLOCB1.COLOR_CODE"


        ASCMAIN1.sql = $"Select Y.*, Z.PRI, Z.OSL, Z.PAL, Z.CRT, Z.TRK, Z.LNF from ({ASCMAIN1.sql}) Y, ({sqlLoc}) Z" & vbCrLf _
            & " where Z.STYLE_CODE (+) = Y.STYLE_CODE and Z.COLOR_CODE (+) = Y.COLOR_CODE"
        ASCMAIN1.sql = $"Select X.*" & vbCrLf _
            & ", NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0)" & vbCrLf _
            & " - CASE WHEN NVL(WHTLOCB1.LOCATION_QTY,0) > 0 THEN WHTLOCB1.LOCATION_QTY ELSE 0 END QTY_AVA" & vbCrLf _
            & $", LEAST(0, NVL(ORDR_QTY_ALLO,0)" & vbCrLf _
            & " - NVL(ORDR_QTY_OPEN,0)) QTY_SHORT" & vbCrLf _
            & $", ICTSTYD1.LOCATION_CODE, WHTLOCB1.LOCATION_QTY LNF" & vbCrLf _
            & $" from ({ASCMAIN1.sql}) X, ICTSTYD1, WHTLOCB1" & vbCrLf _
            & $" where WHTLOCB1.WHSE_CODE (+) = '{WHSE_CODE}' and WHTLOCB1.LOCATION_CODE (+) = 'LOST' and WHTLOCB1.STYLE_CODE (+) = X.STYLE_CODE and WHTLOCB1.COLOR_CODE (+) = X.COLOR_CODE" & vbCrLf _
            & $"   And ICTSTYD1.WHSE_CODE (+) = '{WHSE_CODE}' and ICTSTYD1.STYLE_CODE (+) = X.STYLE_CODE and ICTSTYD1.COLOR_CODE (+) = X.COLOR_CODE"

        'ASCMAIN1.sql = $"Select * from ({ASCMAIN1.sql}) X where X.ORDR_QTY_OPEN <> 0 OR X.ORDR_QTY_BACK <> 0"
        ' INVTY REQUIREMENTS SHOULD ONLY LOOK AT PRIMARY - this was not exactly how it was implemented above

        Fill_Records("ICTSTATO", , , ASCMAIN1.sql)

        EnforceConstraints(True)

        If True Then ' If release_multiple_groups Then
            For Each row As DataRow In dst.Tables("ICTSTATO").Select("QTY_SHORT <> 0 AND ISNULL(ACTION_IF_SHORT,'?') <> 'C' AND ISNULL(ACTION_IF_SHORT,'?') <> 'B'")
                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                Dim rowICTSTATZ As DataRow = dst.Tables("ICTSTATZ").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})

                Dim ACTION_IF_SHORT_default As String = "C" ' "B"
                If rowICTSTATZ Is Nothing Then
                    dst.Tables("ICTSTATZ").Rows.Add(New String() {STYLE_CODE, COLOR_CODE, ACTION_IF_SHORT_default})
                Else
                    If rowICTSTATZ.Item("ACTION_IF_SHORT") & "" = "" Then
                        rowICTSTATZ.Item("ACTION_IF_SHORT") = ACTION_IF_SHORT_default
                    End If
                End If
            Next
        End If

        For Each row As DataRow In dst.Tables("ICTSTATZ").Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim rowICTSTATO As DataRow = dst.Tables("ICTSTATO").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If rowICTSTATO IsNot Nothing Then
                rowICTSTATO.Item("ACTION_IF_SHORT") = row.Item("ACTION_IF_SHORT")
            End If
        Next

        Sort_grdColumns(grdICTSTATO, "QTY_SHORT")
        SplitContainer1.Panel2Collapsed = False

        ASCMAIN1.Progress("")

        Return True
    End Function

    Function Release_Orders(GROUP_KEYs() As String, Optional automatic As Boolean = False, Optional release_multiple_groups As Boolean = False) As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Locking Sales Orders")

        If grdICTSTATO.ActiveRow IsNot Nothing AndAlso grdICTSTATO.ActiveRow.IsDataRow AndAlso grdICTSTATO.ActiveRow.DataChanged Then
            grdICTSTATO.ActiveRow.Update()
        End If

        Dim ErrorMessage As String = ""

        If Not Prepare_Inventory_Requirements(GROUP_KEYs, True, automatic, release_multiple_groups) Then
            ErrorMessage = "Unable to lock all orders"
        ElseIf dst.Tables("ICTSTATO").Select("QTY_SHORT <> 0 AND ISNULL(ACTION_IF_SHORT,'?') <> 'C' AND ISNULL(ACTION_IF_SHORT,'?') <> 'B'").Length > 0 And Not chkIgnoreInvtyShort.Checked Then
            ErrorMessage = "Some Items are Short"
            ' DISABLING THE NEXT 2 LINES SO THAT WE CAN LIVE WITHOUT A PRIMARY
            'ElseIf dst.Tables("ICTSTATO").Select("ISNULL(LOCATION_CODE,'?') = '?' AND QTY_AVA > 0 AND ISNULL(ACTION_IF_SHORT,'?') = '?'").Length > 0 Then
            '    ErrorMessage = "Some Items do NOT have a Primary Location"
        End If

        ASCMAIN1.Progress("Now Checking Sales Orders")

        ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_ALLO, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
            & ", SOTORDR2.STYLE_CLASS_CODE, SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
            & ", SOTORDR2.ORDR_UNIT_PRICE, 'FREEGOODSITEM' ITEM_CODE_CASE" & vbCrLf _
            & $" from SOTORDR1,SOTORDR2,{SOTORDR0} SOTORDR0" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & $"   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0)" & vbCrLf _
            & $"   And SOTORDR2.ORDR_NO = SOTORDR0.ORDR_NO"
        Fill_Records("SOTORDRF", , , ASCMAIN1.sql)

        With dst.Tables("SOTORDRF")
            Dim rows() As DataRow
            rows = .Select("ISNULL(STYLE_CLASS_CODE,'?')='?'")
            If rows.Length > 0 Then
                ErrorMessage &= vbCrLf & $"Some Items on Order {rows(0).Item("ORDR_NO")} have No Product Type - Call ABS"
            End If
            rows = .Select("STYLE_CLASS_CODE<>'INTAPP'") ' IN THE FUTURE, WE MIGHT USE STYLE_CLASS_CODE TO INDICATE THAT THE STYLE IS NOT ONLY DIV 30 BUT ALSO FEASIBLE FOR WEBSITE SALES (?)
            If rows.Length > 0 Then
                ErrorMessage &= vbCrLf & $"Some Items on Order {rows(0).Item("ORDR_NO")} have an Invalid Product Type {rows(0).Item("STYLE_CLASS_CODE")} - Call ABS"
            End If
            rows = .Select("ORDR_TYPE_CODE <> 'XFR' and STYLE_CLASS_CODE <> 'CC' and ISNULL(ORDR_UNIT_PRICE,0) <= 0")
            If rows.Length > 0 Then
                ' ErrorMessage &= vbCrLf & $"Some Items on Order {rows(0).Item("ORDR_NO")} have No Price - Call ABS"
                ' removed this block to match block removal in SOE
            End If
        End With

        ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & $"SELECT X.*, Y.ORDR_QTY_OPEN, Y.ORDR_QTY_PICK FROM {SOTORDQ1} Y, (" & vbCrLf _
            & "SELECT SOTORDR2.ORDR_NO, SUM (SOTORDR2.ORDR_QTY_OPEN) OPEN, SUM (SOTORDR2.ORDR_QTY_PICK) PICK" & vbCrLf _
            & $"from SOTORDR2,{SOTORDR0} SOTORDR0 " & vbCrLf _
            & $"   where SOTORDR2.ORDR_NO = SOTORDR0.ORDR_NO" & vbCrLf _
            & $"   and (NVL(SOTORDR2.ORDR_QTY_OPEN,0) > 0)" & vbCrLf _
            & "group by SOTORDR2.ORDR_NO) X" & vbCrLf _
            & "where X.ORDR_NO = Y.ORDR_NO" & vbCrLf _
            & ") where NVL(OPEN,0) <> NVL(ORDR_QTY_OPEN,0)"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        If tbl.Rows.Count > 0 Then
            ErrorMessage &= vbCrLf & $"Some Orders have Changed - {tbl.Rows(0).Item("ORDR_NO")} - Refresh Required"
        End If

        ASCMAIN1.Progress("")

        If ErrorMessage <> "" Then
            If automatic Then
            Else
                ' at this point, ask the user if we should recursively de-select orders that are short and add orders until we reach the number of orders selected or run out of orders, whichever comes first

                MsgBox(ErrorMessage, vbOKOnly, "Cannot Release these Orders")
                ASCMAIN1.Progress("Now Clearing Locks on Orders")
                ASCMAIN1.MultiTask_Release()
                ASCMAIN1.Progress("")
            End If

            Return ErrorMessage
        End If


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


        If release_multiple_groups Then

            ' preview orders to glean out orders that will not be sent for pick (totally backordered, or just cloth)
            ' for those orders, create a ghost PT with status = C in a separate pick batch; or just don't

            For Each row As DataRow In dst.Tables("SOTORDQ1").Select("")

            Next

            Using ff As New SOFPICKT(Me)
                ff.WHSE_CODE = WHSE_CODE
                ff.SOTORDR0 = SOTORDR0
                ff.SOTORDQ1 = SOTORDQ1
                ff.SALES_DIVISION_CODE_DC = SALES_DIVISION_CODE_SKIN

                ff.ShowDialog()

                If ff.update_flag Then
                    For Each row As DataRow In ASCDATA1.SelectDistinct(ff.tbl.Select("TRUCK_NO IS NOT NULL OR BO = '1'"), New String() {"TRUCK_NO", "BO"}).Select()
                        Dim TRUCK_NO As String = row.Item("TRUCK_NO") & ""
                        Dim BO As String = row.Item("BO") & ""
                        Dim rowSOTTRCK1 As DataRow = LookUp("SOTTRCK1", TRUCK_NO, True)
                        Dim TRUCK_TYPE As String = rowSOTTRCK1.Item("TRUCK_TYPE") & ""

                        Dim PICK_DESCRIPTION As String = IIf(BO = "1", "Back Orders", "Customer Orders")
                        Dim ORDR_SOURCE As String = "W"
                        Dim tblSOTTOTET As DataTable = ff.dst.Tables("SOTTOTET")
                        Create_Pick_Tickets_for_Batch(release_multiple_groups, "", WHSE_CODE, PICK_DESCRIPTION, ORDR_SOURCE, TRUCK_NO, TRUCK_TYPE, BO, ff.tbl, tblSOTTOTET)
                    Next

                    chkIgnoreInvtyShort.Checked = False

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now releasing MT Locks")
                    ASCMAIN1.MultiTask_Release()
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")

                    ASCMAIN1.Progress("Refreshing data...")
                    Setup_Screen()
                    Me.Cursor = DefaultCursor
                    ASCMAIN1.Progress("")
                End If
            End Using


        Else

            Using ff As New SOFPICKU(Me, Absx1.txtFor("WHSE_CODE").Text)

                Dim GROUP_KEY As String = GROUP_KEYs(0)

                ff.GROUP_KEY = GROUP_KEY
                ff.rowSOTORDQ0 = dst.Tables("SOTORDQ0").Select($"GROUP_KEY = '{GROUP_KEY}'")(0)
                ff.SOTORDQ1 = SOTORDQ1
                ff.SALES_DIVISION_CODE_DC = SALES_DIVISION_CODE_SKIN

                ff.ShowDialog()

                If ff.update_flag Then

                    Dim WHSE_CODE As String = grdSOTORDQ0.ActiveRow.Cells("WHSE_CODE").Value & ""
                    Dim PICK_DESCRIPTION As String = grdSOTORDQ0.ActiveRow.Cells("PICK_DESCRIPTION").Value & ""
                    Dim ORDR_SOURCE As String = grdSOTORDQ0.ActiveRow.Cells("ORDR_SOURCE").Value & ""
                    Dim TRUCK_NO As String = ff.TRUCK_NO
                    Dim TRUCK_TYPE As String = ff.TRUCK_TYPE
                    Dim tblSOTTOTET As DataTable = ff.dst.Tables("SOTTOTET")

                    Create_Pick_Tickets_for_Batch(release_multiple_groups, GROUP_KEY, WHSE_CODE, PICK_DESCRIPTION, ORDR_SOURCE, TRUCK_NO, TRUCK_TYPE, "0", ff.tbl, tblSOTTOTET)


                    chkIgnoreInvtyShort.Checked = False

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now releasing MT Locks")
                    ASCMAIN1.MultiTask_Release()
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")

                    ASCMAIN1.Progress("Refreshing data...")
                    Setup_Screen()

                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDQ0.Rows
                        If grow.Cells("GROUP_KEY").Value = GROUP_KEY Then
                            grow.Activate()
                            Exit For
                        End If
                    Next

                    grdSOTORDQ1.ActiveRowScrollRegion.Scroll(RowScrollAction.Top)

                    Me.Cursor = DefaultCursor
                    ASCMAIN1.Progress("")
                End If

            End Using
        End If


        Return ""

    End Function

    Sub Create_Pick_Tickets_for_Batch(release_multiple_groups As Boolean, GROUP_KEY As String,
                                      WHSE_CODE As String, PICK_DESCRIPTION As String, ORDR_SOURCE As String,
                                      TRUCK_NO As String, TRUCK_TYPE As String, BO As String, tbl As DataTable, tblSOTTOTET As DataTable)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Releasing Orders and Creating Pick Tickets")

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR2", "SOTPICK0", "SOTPICK1", "SOTPICK2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(False)

        Dim PTsD As Int32 = 0

        Dim sqlT As String = ""
        If release_multiple_groups Then
            If BO = "1" Then
                sqlT = $"BO = '1'"
            Else
                sqlT = $"TRUCK_NO = '{TRUCK_NO}'"
            End If
        End If

        'Dim PICK_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTPICK0.PICK_BATCH_NO")
        Dim PICK_BATCH_NO As String = ASCMAIN1.Next_Control_No("PICK_BATCH_NO")
        Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").NewRow
        With rowSOTPICK0
            .Item("PICK_BATCH_NO") = PICK_BATCH_NO
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("TRUCK_NO") = TRUCK_NO
            .Item("PICK_BATCH_STATUS") = IIf(BO = "1", "X", "O")
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP

            .Item("WHSE_CODE") = WHSE_CODE
            .Item("ORDERS") = IIf(release_multiple_groups, tbl.Select(sqlT).Length, tbl.Rows.Count)
            .Item("PICK_DESCRIPTION") = PICK_DESCRIPTION
            .Item("ORDR_SOURCE") = ORDR_SOURCE
            .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE_SKIN
        End With
        dst.Tables("SOTPICK0").Rows.Add(rowSOTPICK0)

        For Each row As DataRow In tbl.Select(sqlT, "SLOT_NO")
            Dim ORDR_NO As String = row.Item("ORDR_NO")

            Dim rowSOTORDR1 As DataRow = Fill_Record("SOTORDR1", ORDR_NO)

            Dim rowSOTORDQ1 As DataRow = dst.Tables("SOTORDQ1").Select($"ORDR_NO = '{ORDR_NO}'")(0)

            Dim TOTE_NO As String = row.Item("TOTE_NO") & ""

            Dim PICK_NO As String = ""
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
            With rowSOTPICK1
                '.Item("PICK_NO") = PICK_NO - ASSIGNED LATER IF WE ARE KEEPING THE PT
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_PICK_SEQ") = Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") + 1
                .Item("PICK_STATUS") = IIf(BO = "1", "D", "P")
                .Item("SHIP_VIA_CODE") = rowSOTORDR1.Item("SHIP_VIA_CODE")
                .Item("WHSE_CODE") = rowSOTORDQ1.Item("WHSE_CODE")
                '.Item("CUST_CODE") = rowSOTORDQ1.Item("CUST_CODE")
                '.Item("CUST_SHIP_TO_NO") = rowSOTORDQ1.Item("CUST_SHIP_TO_NO")
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("PICK_BATCH_NO") = PICK_BATCH_NO
                .Item("TOTE_NO") = row.Item("TOTE_NO")
                .Item("SLOT_NO") = row.Item("SLOT_NO")
                '.Item("PICK_SOURCE") = rowSOTORDR1.Item("ORDR_SOURCE") ' NECESSARY?
                .Item("WHSE_CODE") = Absx1.txtFor("WHSE_CODE").Text
                .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE_SKIN
            End With

            Fill_Records("SOTORDR2", ORDR_NO)
            Dim PICK_LNO_ctr As Integer = 0
            Dim PICK_QTY_total As Int32 = 0
            Dim PICK_QTY_BACK_total As Int32 = 0
            Dim PICK_QTY_CANC_total As Int32 = 0

            'Dim FR_TO_PICK As Int32 = 0
            'Dim FR_PICK_QTY As Int32 = 0

            Dim LIVE_PROD_TO_PICK As Int32 = 0
            Dim LIVE_PROD_PICK_QTY As Int32 = 0

            ' SORT THE PT DETAIL CANDIDATES BY PICK_SEQ: A = FR OR SG, Z = CS
            Dim sqlw As String = "ISNULL(ORDR_QTY_OPEN,0) <> 0"
            sqlw = Replace(sqlw, "SOTORDR2.", "")
            sqlw = $"ISNULL(ORDR_QTY_OPEN,0) > 0 and {sqlw}"
            For Each row2 As DataRow In dst.Tables("SOTORDR2").Select(sqlw, "PICK_SEQ")
                Dim STYLE_CODE As String = row2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = row2.Item("COLOR_CODE")
                Dim ORDR_QTY_OPEN As Int32 = Val(row2.Item("ORDR_QTY_OPEN") & "")
                Dim ORDR_QTY_BACK As Int32 = 0 ' Val(row2.Item("ORDR_QTY_BACK") & "")
                Dim QTY_TO_PICK As Int32 = ORDR_QTY_OPEN + ORDR_QTY_BACK

                'Dim FR_or_CS As String = row2.Item("STYLE_CLASS_CODE") & ""
                'If FR_or_CS <> "CS" Then FR_or_CS = "FR"
                'If FR_or_CS = "FR" Then
                '    FR_TO_PICK += QTY_TO_PICK
                'End If

                Dim PICK_SEQ As String = row2.Item("PICK_SEQ") ' A = REAL PRODUCT  OP, FR, SG, ELSE Z = CS, CL
                If PICK_SEQ <> "Z" Then
                    LIVE_PROD_TO_PICK += QTY_TO_PICK
                End If

                Dim rowICTSTATO As DataRow = dst.Tables("ICTSTATO").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                Dim QTY_PICKED As Int32 = Val(rowICTSTATO.Item("QTY_PICKED") & "")
                'Dim QTY_AVA As Int32 = Val(rowICTSTATO.Item("QTY_AVA") & "")
                Dim ORDR_LNO As Int32 = Val(row2.Item("ORDR_LNO") & "")
                Dim rowSOTORDRX() As DataRow = dst.Tables("SOTORDRX").Select($"ORDR_NO = '{ORDR_NO}' and ORDR_LNO = {CStr(ORDR_LNO)}") ' .Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                Dim ORDR_QTY_ALLO As Int32 = Val(rowSOTORDRX(0).Item("ORDR_QTY_ALLO") & "")
                Dim QTY_AVA As Int32 = ORDR_QTY_ALLO
                Dim ACTION_IF_SHORT As String = rowICTSTATO.Item("ACTION_IF_SHORT") & ""

                Dim PICK_QTY_BACK As Int32 = 0
                Dim PICK_QTY_CANC As Int32 = 0

                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
                With rowSOTPICK2
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = row2.Item("ORDR_LNO")

                    If PICK_SEQ = "Z" And LIVE_PROD_TO_PICK > 0 And LIVE_PROD_PICK_QTY = 0 Then '  If FR_or_CS = "CS" And FR_TO_PICK > 0 And FR_PICK_QTY = 0 Then
                        ' WE HAD SL TO PICK, DIDNT PICK ANY, SO AUTO BACK-ORDER THE CLOTHS
                        PICK_QTY_BACK = QTY_TO_PICK
                        .Item("PICK_QTY_BACK") = QTY_TO_PICK ' NEED TO SOMEHOW NOT OOS PO THESE *******************************
                        PICK_QTY_BACK_total += QTY_TO_PICK
                    Else

                        If QTY_TO_PICK > QTY_AVA And Not chkIgnoreInvtyShort.Checked Then ' If QTY_PICKED + QTY_TO_PICK > QTY_AVA And Not chkIgnoreInvtyShort.Checked Then
                            If ACTION_IF_SHORT = "C" Then
                                PICK_QTY_CANC = QTY_TO_PICK
                                .Item("PICK_QTY_CANC") = QTY_TO_PICK
                                PICK_QTY_CANC_total += QTY_TO_PICK
                            ElseIf ACTION_IF_SHORT = "B" Then
                                PICK_QTY_BACK = QTY_TO_PICK
                                .Item("PICK_QTY_BACK") = QTY_TO_PICK
                                PICK_QTY_BACK_total += QTY_TO_PICK
                            Else
                                Throw New Exception($"Unexpected Action if Short ({ACTION_IF_SHORT}) for Style-Color {STYLE_CODE & "-" & COLOR_CODE} on Order {ORDR_NO}")
                            End If
                        Else
                            QTY_PICKED += QTY_TO_PICK
                            rowICTSTATO.Item("QTY_PICKED") = QTY_PICKED
                            .Item("PICK_QTY") = QTY_TO_PICK
                            PICK_QTY_total += QTY_TO_PICK

                            If PICK_SEQ = "A" Then
                                LIVE_PROD_PICK_QTY += QTY_TO_PICK
                            End If
                        End If
                    End If

                    If ORDR_QTY_OPEN = 0 And ORDR_QTY_BACK = PICK_QTY_BACK Then
                        ' DO NOT WRITE SOTPICK2 - NOTHING TO DO
                    Else
                        If PICK_LNO_ctr = 0 Then
                            'PICK_NO = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")
                            PICK_NO = ASCMAIN1.Next_Control_No("PICK_NO")
                            rowSOTPICK1.Item("PICK_NO") = PICK_NO
                            dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)
                        End If

                        PICK_LNO_ctr += 1
                        .Item("PICK_NO") = PICK_NO
                        .Item("PICK_LNO") = PICK_LNO_ctr
                        dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
                    End If
                End With

            Next

            Dim ORDR_QTY_OPEN_total As Int32 = Val(dst.Tables("SOTORDR2").Compute("SUM (ORDR_QTY_OPEN)", "") & "")

            'Dim ORDR_QTY_BACK_total As Int32 = 0

            If PICK_QTY_total = 0 Then ' We have a PT with nothing on it to pick, but it may have BOs or Cancels
                If TRUCK_TYPE = "X" Then ' so delete the custom truck, if one was assigned
                    Dim rowSOTTOTET As DataRow = tblSOTTOTET.Rows.Find(TOTE_NO)
                    rowSOTTOTET.Delete()
                End If

                If ORDR_QTY_OPEN_total = 0 And PICK_QTY_CANC_total = 0 Then ' if the order had nothing open (and therefore must have had BOs) and we cancelled nothing
                    If PICK_LNO_ctr <> 0 Then ' if pick lines were assoned, or if the total BO on the order <> total BO on the PT
                        Throw New Exception($"Issue with Releasing Order {ORDR_NO} - Please Call ABS")
                    End If
                    rowSOTPICK1.Item("PICK_STATUS") = "D" ' probably the entire PT BOs an order that is already BOd, so delete the PT
                    PTsD += 1
                Else
                    rowSOTPICK1.Item("PICK_STATUS") = "C" ' cancel all of the lines on this PT
                    rowSOTPICK1.Item("TOTE_NO") = DBNull.Value
                    rowSOTPICK1.Item("SLOT_NO") = DBNull.Value

                End If
            End If
        Next

        Dim PTsP As Int32 = dst.Tables("SOTPICK1").Select("PICK_STATUS = 'P'").Length
        Dim PTsC As Int32 = dst.Tables("SOTPICK1").Select("PICK_STATUS = 'C'").Length
        ' Dim PTsD As Int32 = dst.Tables("SOTPICK1").Select("PICK_STATUS = 'D'").Length - these are not even written out to SOTPICK1
        rowSOTPICK0.Item("ORDERS") = PTsP
        Try

            BeginTrans()

            Dim r As Int32 = 0

            If PTsP = 0 Then
                If PTsC = 0 Then
                    If PTsD = 0 Then
                        Throw New Exception($"Issue with Group Key {GROUP_KEY} - Please Call ABS")
                    End If
                    rowSOTPICK0.Delete()
                Else
                    rowSOTPICK0.Item("PICK_BATCH_STATUS") = "C"
                End If

            Else
                Dim sqlTT As String = ""
                If TRUCK_TYPE = "X" Then
                    sqlTT = ", TRUCK_TYPE = 'X'"
                End If

                ASCMAIN1.sql = $"Update SOTTRCK1 Set PICK_BATCH_NO = '{PICK_BATCH_NO}'" & sqlTT & " where TRUCK_NO = :PARM1 and PICK_BATCH_NO is Null"
                r = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", TRUCK_NO)
                If r <> 1 Then
                    Throw New Exception($"Could not Update Truck {TRUCK_NO} with Pick Batch {PICK_BATCH_NO}")
                End If

                If TRUCK_TYPE = "X" Then
                    dst.Tables("SOTTOTE1").Rows.Clear()
                    For Each rowSOTTOTET As DataRow In tblSOTTOTET.Select("")
                        dst.Tables("SOTTOTE1").Rows.Add(rowSOTTOTET.ItemArray)
                    Next
                    Update_Record_TDA("SOTTOTE1")
                End If

                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_STATUS = 'P'")
                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                    Dim TOTE_NO As String = rowSOTPICK1.Item("TOTE_NO")

                    ASCMAIN1.Progress("-", PICK_NO)

                    sqlTT = ""
                    If TRUCK_TYPE = "R" Then
                        sqlTT = $", TRUCK_NO = '{TRUCK_NO}'"
                    End If

                    ASCMAIN1.sql = $"Update SOTTOTE1 Set PICK_NO = '{PICK_NO}'" & sqlTT & " where TOTE_NO = :PARM1 and PICK_NO is Null"
                    r = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", TOTE_NO)
                    If r <> 1 Then
                        Throw New Exception($"Could not Update Tote {TOTE_NO} with Pick No {PICK_NO}")
                    End If
                Next
            End If

            Update_Record_TDA("SOTPICK0")
            Update_Record_TDA("SOTPICK1")
            Update_Record_TDA("SOTPICK2")

            'ASCMAIN1.sql = "Merge into ICTSTAT2 using (" & vbCrLf _
            '        & "Select SOTPICK1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            '        & ", Sum(NVL(SOTORDR2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN" & vbCrLf _
            '        & ", Sum(NVL(SOTORDR2.ORDR_QTY_BACK,0)) ORDR_QTY_BACK" & vbCrLf _
            '        & ", Sum(NVL(SOTPICK2.PICK_QTY,0)) PICK_QTY" & vbCrLf _
            '        & ", Sum(NVL(SOTPICK2.PICK_QTY_BACK,0)) PICK_QTY_BACK" & vbCrLf _
            '        & " from SOTPICK1" & vbCrLf _
            '        & " join SOTPICK2 on (SOTPICK1.PICK_NO = SOTPICK2.PICK_NO)" & vbCrLf _
            '        & " join SOTORDR2 ON (SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO and SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO)" & vbCrLf _
            '        & " where SOTPICK1.PICK_BATCH_NO = :PARM1" & vbCrLf _
            '        & " group by SOTPICK1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            '        & ") X on (X.WHSE_CODE = ICTSTAT2.WHSE_CODE and x.STYLE_CODE = ICTSTAT2.STYLE_CODE and x.COLOR_CODE = ICTSTAT2.COLOR_CODE)" & vbCrLf _
            '        & " when matched Then Update" & vbCrLf _
            '        & "Set WHSE_QTY_OPEN = NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - (X.ORDR_QTY_OPEN + X.ORDR_QTY_BACK) + X.PICK_QTY_BACK," & vbCrLf _
            '        & "    WHSE_QTY_PICK = NVL(ICTSTAT2.WHSE_QTY_PICK,0) + X.PICK_QTY" & vbCrLf _
            '        & " when NOT matched Then" & vbCrLf _
            '        & "    Insert (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_OPEN, WHSE_QTY_PICK)" & vbCrLf _
            '        & "    Values (X.STYLE_CODE, X.COLOR_CODE, X.WHSE_CODE, -1 * (X.ORDR_QTY_OPEN + X.ORDR_QTY_BACK) + X.PICK_QTY_BACK, X.PICK_QTY)"

            ASCMAIN1.sql = "Merge into ICTSTAT2 using (" & vbCrLf _
                    & "Select SOTPICK1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                    & ", Sum(NVL(SOTORDR2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN" & vbCrLf _
                    & ", Sum(NVL(SOTPICK2.PICK_QTY,0)) PICK_QTY" & vbCrLf _
                    & " from SOTPICK1" & vbCrLf _
                    & " join SOTPICK2 on (SOTPICK1.PICK_NO = SOTPICK2.PICK_NO)" & vbCrLf _
                    & " join SOTORDR2 ON (SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO and SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO)" & vbCrLf _
                    & " where SOTPICK1.PICK_BATCH_NO = :PARM1" & vbCrLf _
                    & " group by SOTPICK1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                    & ") X on (X.WHSE_CODE = ICTSTAT2.WHSE_CODE and x.STYLE_CODE = ICTSTAT2.STYLE_CODE and x.COLOR_CODE = ICTSTAT2.COLOR_CODE)" & vbCrLf _
                    & " when matched Then Update" & vbCrLf _
                    & "Set WHSE_QTY_OPEN = NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - (X.ORDR_QTY_OPEN)," & vbCrLf _
                    & "    WHSE_QTY_PICK = NVL(ICTSTAT2.WHSE_QTY_PICK,0) + X.PICK_QTY" & vbCrLf _
                    & " when NOT matched Then" & vbCrLf _
                    & "    Insert (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_OPEN, WHSE_QTY_PICK)" & vbCrLf _
                    & "    Values (X.STYLE_CODE, X.COLOR_CODE, X.WHSE_CODE, -1 * (X.ORDR_QTY_OPEN), X.PICK_QTY)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_BATCH_NO)

            'ASCMAIN1.sql = "Merge into SOTORDR2 using (" & vbCrLf _
            '        & "Select SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO, SOTPICK1.WHSE_CODE" & vbCrLf _
            '        & ", NVL(SOTPICK2.PICK_QTY,0) PICK_QTY" & vbCrLf _
            '        & ", NVL(SOTPICK2.PICK_QTY_BACK,0) PICK_QTY_BACK" & vbCrLf _
            '        & ", NVL(SOTPICK2.PICK_QTY_CANC,0) PICK_QTY_CANC from SOTPICK1" & vbCrLf _
            '        & " join SOTPICK2 on (SOTPICK1.PICK_NO = SOTPICK2.PICK_NO)" & vbCrLf _
            '        & " where SOTPICK1.PICK_BATCH_NO = :PARM1" & vbCrLf _
            '        & ") X on (SOTORDR2.ORDR_NO = X.ORDR_NO AND SOTORDR2.ORDR_LNO = X.ORDR_LNO)" & vbCrLf _
            '        & " when Matched Then Update" & vbCrLf _
            '        & "Set ORDR_QTY_OPEN = 0, ORDR_QTY_BACK = X.PICK_QTY_BACK" & vbCrLf _
            '        & ", ORDR_QTY_CANC = NVL(SOTORDR2.ORDR_QTY_CANC,0) + X.PICK_QTY_CANC" & vbCrLf _
            '        & ", ORDR_QTY_PICK = NVL(SOTORDR2.ORDR_QTY_PICK,0) + X.PICK_QTY" & vbCrLf _
            '        & ", ORDR_LINE_STATUS = CASE WHEN X.PICK_QTY_BACK <> 0 THEN 'B' ELSE CASE WHEN X.PICK_QTY <> 0 THEN 'P' ELSE CASE WHEN SOTORDR2.ORDR_QTY_SHIP <> 0 THEN 'F' ELSE 'C' END END END"
            'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_BATCH_NO)

            ASCMAIN1.sql = "Merge into SOTORDR2 using (" & vbCrLf _
                    & "Select SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO, SOTPICK1.WHSE_CODE" & vbCrLf _
                    & ", NVL(SOTPICK2.PICK_QTY,0) PICK_QTY" & vbCrLf _
                    & ", NVL(SOTPICK2.PICK_QTY_CANC,0) PICK_QTY_CANC from SOTPICK1" & vbCrLf _
                    & " join SOTPICK2 on (SOTPICK1.PICK_NO = SOTPICK2.PICK_NO)" & vbCrLf _
                    & " where SOTPICK1.PICK_BATCH_NO = :PARM1" & vbCrLf _
                    & ") X on (SOTORDR2.ORDR_NO = X.ORDR_NO AND SOTORDR2.ORDR_LNO = X.ORDR_LNO)" & vbCrLf _
                    & " when Matched Then Update" & vbCrLf _
                    & "Set ORDR_QTY_OPEN = 0" & vbCrLf _
                    & ", ORDR_QTY_CANC = NVL(SOTORDR2.ORDR_QTY_CANC,0) + X.PICK_QTY_CANC" & vbCrLf _
                    & ", ORDR_QTY_PICK = NVL(SOTORDR2.ORDR_QTY_PICK,0) + X.PICK_QTY" & vbCrLf _
                    & ", ORDR_STATUS = CASE WHEN X.PICK_QTY <> 0 THEN 'P' ELSE CASE WHEN SOTORDR2.ORDR_QTY_SHIP <> 0 THEN 'F' ELSE 'C' END END"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", PICK_BATCH_NO)

            'ASCMAIN1.sql = "" _
            '& "Begin Declare ORDR_STATUS_calc VARCHAR2(1); Cursor C1 is" & vbCrLf _
            '& "  Select SOTORDR2.ORDR_NO" & vbCrLf _
            '& ", Sum (NVL(SOTORDR2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN" & vbCrLf _
            '& ", Sum (NVL(SOTORDR2.ORDR_QTY_BACK,0)) ORDR_QTY_BACK" & vbCrLf _
            '& ", Sum (NVL(SOTORDR2.ORDR_QTY_PICK,0)) ORDR_QTY_PICK" & vbCrLf _
            '& ", Sum (NVL(SOTORDR2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
            '& " from SOTORDR2,SOTPICK1" & vbCrLf _
            '& " where SOTORDR2.ORDR_NO = SOTPICK1.ORDR_NO and SOTPICK1.PICK_BATCH_NO = :PARM1" & vbCrLf _
            '& " group by SOTORDR2.ORDR_NO;" & vbCrLf _
            '& " Begin For R1 in C1 loop" & vbCrLf _
            '& " If R1.ORDR_QTY_OPEN > 0 OR R1.ORDR_QTY_BACK > 0 Then" & vbCrLf _
            '& "  ORDR_STATUS_calc := 'O';" & vbCrLf _
            '& " Else " & vbCrLf _
            '& "  If R1.ORDR_QTY_PICK > 0 Then" & vbCrLf _
            '& "   ORDR_STATUS_calc := 'P';" & vbCrLf _
            '& "  Else " & vbCrLf _
            '& "   If R1.ORDR_QTY_SHIP > 0 Then" & vbCrLf _
            '& "    ORDR_STATUS_calc := 'F';" & vbCrLf _
            '& "   Else " & vbCrLf _
            '& "    ORDR_STATUS_calc := 'C';" & vbCrLf _
            '& "   End if;" & vbCrLf _
            '& "  End if;" & vbCrLf _
            '& " End if;" & vbCrLf _
            '& $" Update SOTORDR1 Set ORDR_STATUS = ORDR_STATUS_calc, ORDR_PICK_SEQ = NVL(ORDR_PICK_SEQ,0) + 1, LAST_OPER = :PARM2, LAST_DATE = SYSDATE" & vbCrLf _
            '& "  where ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            '& " End Loop; End;" & vbCrLf _
            '& "End;"
            'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {PICK_BATCH_NO, ASCMAIN1.USER_ID})

            ASCMAIN1.sql = "" _
            & "Begin Declare ORDR_STATUS_calc VARCHAR2(1); Cursor C1 is" & vbCrLf _
            & "  Select SOTORDR2.ORDR_NO" & vbCrLf _
            & ", Sum (NVL(SOTORDR2.ORDR_QTY_OPEN,0)) ORDR_QTY_OPEN" & vbCrLf _
            & ", Sum (NVL(SOTORDR2.ORDR_QTY_PICK,0)) ORDR_QTY_PICK" & vbCrLf _
            & ", Sum (NVL(SOTORDR2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
            & " from SOTORDR2,SOTPICK1" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK1.ORDR_NO and SOTPICK1.PICK_BATCH_NO = :PARM1" & vbCrLf _
            & " group by SOTORDR2.ORDR_NO;" & vbCrLf _
            & " Begin For R1 in C1 loop" & vbCrLf _
            & " If R1.ORDR_QTY_OPEN > 0 Then" & vbCrLf _
            & "  ORDR_STATUS_calc := 'O';" & vbCrLf _
            & " Else " & vbCrLf _
            & "  If R1.ORDR_QTY_PICK > 0 Then" & vbCrLf _
            & "   ORDR_STATUS_calc := 'P';" & vbCrLf _
            & "  Else " & vbCrLf _
            & "   If R1.ORDR_QTY_SHIP > 0 Then" & vbCrLf _
            & "    ORDR_STATUS_calc := 'F';" & vbCrLf _
            & "   Else " & vbCrLf _
            & "    ORDR_STATUS_calc := 'C';" & vbCrLf _
            & "   End if;" & vbCrLf _
            & "  End if;" & vbCrLf _
            & " End if;" & vbCrLf _
            & $" Update SOTORDR1 Set ORDR_STATUS = ORDR_STATUS_calc, ORDR_PICK_SEQ = NVL(ORDR_PICK_SEQ,0) + 1, LAST_OPER = :PARM2, LAST_DATE = SYSDATE" & vbCrLf _
            & "  where ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & " End Loop; End;" & vbCrLf _
            & "End;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {PICK_BATCH_NO, ASCMAIN1.USER_ID})


            If PTsP > 0 Then


                Try
                    Dim ORDER_COUNT As Integer = dst.Tables("SOTPICK1").Select("PICK_STATUS = 'P'").Length
                    Print_Truck_Pick_Tag("Print Truck Tag", TRUCK_NO, PICK_DESCRIPTION, ORDER_COUNT, PICK_BATCH_NO)

                    If TRUCK_TYPE = "X" Then
                        Print_Custom_Tote_Labels("Print Custom Tote Labels", PICK_BATCH_NO, TRUCK_NO)
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Problem printing Labels: Call ABS")
                End Try

            End If

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

            If release_multiple_groups Then
            Else
                If PTsC <> 0 Or PTsD <> 0 Or PTsP = 0 Then
                    Dim MSG As String = ""
                    If PTsC <> 0 Then MSG &= vbCrLf & $"{CStr(PTsC)} PTs were Cancelled; {CStr(PTsP)} PTs issued"
                    If PTsD <> 0 Then MSG &= vbCrLf & $"{CStr(PTsD)} Orders Skipped - already on B/O"
                    If PTsP = 0 Then MSG &= vbCrLf & $"{CStr(PTsP)} PTs Generated - Truck was NOT assigned"


                    ' MSG &= vbCrLf & vbCrLf & "Enter Y to Acknowledge"
                    Do While InputBox(MSG, "Please Note the following - Enter Y to Acknowledge:") <> "Y"

                    Loop

                    'MsgBox(MSG, MsgBoxStyle.OkOnly, "Please Note the following:")
                End If

            End If

            Dim MSG2 As String = $"Pick Batch {PICK_BATCH_NO} has been Created"
            If PTsP = 0 Then
                MSG2 = "Nothing to Pick"
            End If
            If release_multiple_groups Then
                MSG2 = ""
            End If
            CommitTrans(MSG2)

            If release_multiple_groups Then
            Else
                For Each row As DataRow In dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' AND SEL='1'")
                    row.Item("SEL") = "0"
                Next
            End If


        Catch ex As Exception
            Rollback("Update Rolled Back: " & ex.Message & vbCrLf & vbCrLf & "Please STOP - call ABS")
        End Try

    End Sub

    Sub Print_Custom_Tote_Labels(LABEL_ACTION As String, PICK_BATCH_NO As String, TRUCK_NO As String, Optional TOTE_NO As String = "")

        Me.Cursor = Cursors.WaitCursor
        If TOTE_NO <> "" Then
            ASCMAIN1.Progress($"Now Printing Custom Tote Label for Tote {TOTE_NO}")
        Else
            ASCMAIN1.Progress("Now Printing Custom Tote Labels")
        End If

        zplPrint.Print_Custom_Tote_Labels(LABEL_ACTION, TRUCK_NO, TOTE_NO, PICK_BATCH_NO, 0, TRUCK_NO, 0)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Print_Truck_Pick_Tag(LABEL_ACTION As String, TRUCK_NO As String, PICK_DESCRIPTION As String, ORDER_COUNT As Integer, PICK_BATCH_NO As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Truck Tag")

        zplPrint.Print_Truck_Pick_Tag(LABEL_ACTION, TRUCK_NO, PICK_DESCRIPTION, ORDER_COUNT, PICK_BATCH_NO, 0, TRUCK_NO, 0)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICK1.AfterRowActivate
        If EntryMode = "R" Then
            If grdSOTPICK1.ActiveRow.Cells("PICK_REQ_RES").Value & "" = "1" Then
                grdSOTPICK1.DisplayLayout.Bands(0).Columns("RESOLUTION").CellActivation = Activation.AllowEdit
            Else
                grdSOTPICK1.DisplayLayout.Bands(0).Columns("RESOLUTION").CellActivation = Activation.NoEdit
            End If

            Dim RESOLUTION As String = grdSOTPICK1.ActiveRow.Cells("RESOLUTION").Value & ""
            grdSOTPICK2.DisplayLayout.Bands(0).Columns("RESOLUTION").Hidden = (RESOLUTION = "D" Or RESOLUTION = "T")
        Else

        End If
        Setup_grdSOTPICK2()
    End Sub

    Sub Setup_grdSOTPICK2()
        If grdSOTPICK1.ActiveRow Is Nothing OrElse (grdSOTPICK1.ActiveRow.IsFilterRow Or Not grdSOTPICK1.ActiveRow.IsDataRow) Then
            grdSOTPICK2.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value & ""

            grdSOTPICK2.Text = $"Pick Tickets Details for Pick No {PICK_NO}"

            Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"PICK_NO = '{PICK_NO}'"
            Sort_grdColumns(grdSOTPICK2, "PICK_LNO")

            grdSOTPICK2.Visible = True

        End If
    End Sub

    Private Sub grdSOTPICK2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPICK2.InitializeRow
        If e.Row.IsDataRow Then
            Dim PICK_STATUS_CODE As String = e.Row.Cells("PICK_STATUS_CODE").Value & ""
            Dim PICK_QTY As Integer = Val(e.Row.Cells("PICK_QTY").Value & "")
            Dim PICK_QTY_CONF As Integer = Val(e.Row.Cells("PICK_QTY_CONF").Value & "")
            Dim PICK_QTY_BACK As Integer = Val(e.Row.Cells("PICK_QTY_BACK").Value & "")
            Dim PICK_CNL_STATUS As String = e.Row.Cells("PICK_CNL_STATUS").Value & ""

            If PICK_STATUS_CODE = "0" Then
                e.Row.Cells("PICK_STATUS_CODE").Appearance = Appearance_Yellow
            ElseIf PICK_STATUS_CODE = "1" Then
                e.Row.Cells("PICK_STATUS_CODE").Appearance = Appearance_Green
            Else
                ' e.Row.Cells("PICK_STATUS").Appearance = Appearance_BLUE
            End If

            If PICK_CNL_STATUS = "1" Then
                e.Row.Cells("PICK_CNL_STATUS").Appearance.BackColor = System.Drawing.Color.Red
                e.Row.Cells("PICK_CNL_STATUS").Appearance.ForeColor = System.Drawing.Color.White
            Else
                e.Row.Cells("PICK_CNL_STATUS").Appearance.BackColor = System.Drawing.Color.Empty
                e.Row.Cells("PICK_CNL_STATUS").Appearance.ForeColor = System.Drawing.Color.Empty
            End If


            If PICK_CNL_STATUS = "1" Then
                e.Row.Cells("PICK_QTY_BACK").Appearance.BackColor = System.Drawing.Color.Yellow
            Else
                e.Row.Cells("PICK_QTY_BACK").Appearance.BackColor = System.Drawing.Color.Empty
            End If

        End If
    End Sub

    Private Sub grdSOTPICK1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPICK1.InitializeRow
        If e.Row.IsDataRow Then
            Dim PICK_STATUS_CODE As String = e.Row.Cells("PICK_STATUS_CODE").Value & ""

            If PICK_STATUS_CODE = "0" Then
                e.Row.Cells("PICK_STATUS_CODE").Appearance = Appearance_Yellow
            ElseIf PICK_STATUS_CODE = "1" Then
                e.Row.Cells("PICK_STATUS_CODE").Appearance = Appearance_Green
            Else
                ' e.Row.Cells("PICK_STATUS").Appearance = Appearance_BLUE
            End If

            Dim PICK_REQ_RES As String = e.Row.Cells("PICK_REQ_RES").Value & ""

            If PICK_REQ_RES = "1" Then
                e.Row.Cells("PICK_REQ_RES").Appearance.BackColor = System.Drawing.Color.Red
            Else
                e.Row.Cells("PICK_REQ_RES").Appearance.BackColor = System.Drawing.Color.Empty
            End If

        End If
    End Sub

    Private Sub grdSOTPICK0_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPICK0.InitializeRow
        If e.Row.IsDataRow Then
            Dim PICK_BATCH_STATUS As String = e.Row.Cells("PICK_BATCH_STATUS").Value & ""

            If PICK_BATCH_STATUS = "N" Then
                e.Row.Cells("PICK_BATCH_STATUS").Appearance = Appearance_Empty
            ElseIf PICK_BATCH_STATUS = "P" Then
                e.Row.Cells("PICK_BATCH_STATUS").Appearance = Appearance_Green
            ElseIf PICK_BATCH_STATUS = "O" Then
                e.Row.Cells("PICK_BATCH_STATUS").Appearance = Appearance_Yellow
            ElseIf PICK_BATCH_STATUS = "R" Then
                e.Row.Cells("PICK_BATCH_STATUS").Appearance = Appearance_Red
            End If
        End If
    End Sub

    Private Sub grdSOTPICK0_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTPICK0.DoubleClickRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("PICK_BATCH_STATUS").Value = "R" Then
                If (MENU_ITEM_OBJECT = "SOFPICKS") Then
                    If grdSOTPICK0.ActiveCell IsNot Nothing AndAlso grdSOTPICK0.ActiveCell.Column.Key = "PICK_BATCH_STATUS" Then
                        PICK_BATCH_NO = e.Row.Cells("PICK_BATCH_NO").Value
                        If MsgBox($"Do you want to Resolve Orders in Pick Batch {PICK_BATCH_NO}", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
                        Click_Command("Resolve")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub grdSOTPICK2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICK2.AfterRowActivate
        If EntryMode = "R" Then
            If Val(grdSOTPICK2.ActiveRow.Cells("PICK_QTY_BACK").Value & "") <> 0 Or Val(grdSOTPICK2.ActiveRow.Cells("PICK_QTY_CANC").Value & "") <> 0 Then
                grdSOTPICK2.DisplayLayout.Bands(0).Columns("RESOLUTION").CellActivation = Activation.AllowEdit
            Else
                grdSOTPICK2.DisplayLayout.Bands(0).Columns("RESOLUTION").CellActivation = Activation.NoEdit
            End If
        Else

        End If
    End Sub

    Private Sub grdSOTPICK1_AfterCellListCloseUp(sender As Object, e As CellEventArgs) Handles grdSOTPICK1.AfterCellListCloseUp
        Dim RESOLUTION As String = e.Cell.Value & ""
        grdSOTPICK1.ActiveRow.Update()
        'Stop
    End Sub

    Private Sub grdSOTPICK1_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdSOTPICK1.AfterCellUpdate
        Dim RESOLUTION As String = e.Cell.Value & ""
        Dim PICK_NO As String = e.Cell.Row.Cells("PICK_NO").Value & ""
        Set_SOTPICK2_RESOLUTION(RESOLUTION, PICK_NO)
        'Stop
    End Sub

    Sub Set_SOTPICK2_RESOLUTION(RESOLUTION As String, PICK_NO As String)

        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND ISNULL(PICK_QTY_BACK,0) <> 0")
            rowSOTPICK2.Item("RESOLUTION") = RESOLUTION
        Next

        grdSOTPICK2.DisplayLayout.Bands(0).Columns("RESOLUTION").Hidden = (RESOLUTION = "D" Or RESOLUTION = "T")
    End Sub


    Sub Print_Report(Optional resolutionsOnly As Boolean = False)
        Dim SUBT As String = ""
        Dim RESOLUTION As String = ""
        If resolutionsOnly Then
            RESOLUTION = " and {SOTPICK2.PICK_CNL_STATUS} = '1'"
            SUBT = "Showing Orders Requiring Resolution, Only"
        End If
        Print_Report_Begin()
        Generate_Report("SORPICKF", "Pick Ticket Resolution Report", SUBT, "{SOTPICK1.PICK_BATCH_NO} = " & $"'{PICK_BATCH_NO}'" & RESOLUTION)
        Print_Report_End()
    End Sub

    Sub Print_Test_Report(PICK_BATCH_NO As String)
        Dim SUBT As String = ""

        Print_Report_Begin()
        Generate_Report("SORPICKB", "Pick Test Report", SUBT, "{SOTPICK1.PICK_BATCH_NO} = " & $"'{PICK_BATCH_NO}'")
        Print_Report_End()
    End Sub

    Sub Label_RePrint()
        Using ff As New TAFZPLH1(Me, "Pick Batch No", , "Truck")
            ff.ShowDialog()
        End Using
    End Sub

    Sub Calculate_Short()

        dst.Tables("ICTSTATZ").Rows.Clear()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Calculating Qty Short for each Order Group")
        splOrders.Panel2Collapsed = True
        For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDQ0.Rows
            grow.Activate()
            For Each grow1 As UltraWinGrid.UltraGridRow In grdSOTORDQ1.Rows
                grow1.Cells("SEL").Value = "1"
                grow1.Update()
            Next
            Dim GROUP_KEY As String = grow.Cells("GROUP_KEY").Value & ""
            ASCMAIN1.Progress("-", GROUP_KEY)
            Prepare_Inventory_Requirements(New String() {GROUP_KEY}, False)
            Dim QTY_SHORT As Integer = Val(dst.Tables("ICTSTATO").Compute("SUM(QTY_SHORT)", "") & "")
            grow.Cells("QTY_SHORT").Value = QTY_SHORT
            grow.Update()
        Next

        grdSOTORDQ0.DisplayLayout.Bands(0).Columns("QTY_SHORT").Hidden = False
        splOrders.Panel2Collapsed = False

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub btnMaxRelease_Click(sender As Object, e As EventArgs) Handles btnMaxRelease.Click

        Select_Orders(MAX_ORDERS_TO_RELEASE, True)

    End Sub

    Private Sub grdICTSTATO_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTSTATO.BeforeRowUpdate
        Dim QTY_SHORT As Int32 = Val(e.Row.Cells("QTY_SHORT").Value & "")
        If QTY_SHORT = 0 Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdICTSTATO_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdICTSTATO.AfterRowUpdate
        Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value
        Dim ACTION_IF_SHORT As String = e.Row.Cells("ACTION_IF_SHORT").Value & ""
        Dim row As DataRow = dst.Tables("ICTSTATZ").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
        If row Is Nothing Then
            row = dst.Tables("ICTSTATZ").Rows.Add(New String() {STYLE_CODE, COLOR_CODE})
        End If
        row.Item("ACTION_IF_SHORT") = ACTION_IF_SHORT

    End Sub

    Private Sub grdICTSTATO_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTSTATO.AfterRowActivate
        If grdICTSTATO.ActiveRow.Band.Index = 0 Then
            If grdICTSTATO.ActiveRow IsNot Nothing AndAlso grdICTSTATO.ActiveRow.IsDataRow Then
                Dim QTY_SHORT As Int32 = Val(grdICTSTATO.ActiveRow.Cells("QTY_SHORT").Value & "")
                If QTY_SHORT = 0 Then
                    grdICTSTATO.DisplayLayout.Bands(0).Columns("ACTION_IF_SHORT").CellActivation = Activation.NoEdit
                Else
                    grdICTSTATO.DisplayLayout.Bands(0).Columns("ACTION_IF_SHORT").CellActivation = Activation.AllowEdit
                End If
            End If
        End If
    End Sub

    Sub Fetch_All_Orders()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Fetching Open Sales Order Details")

        Fill_Records("SOTORDRX")

        Sort_grdColumns(grdSOTORDRX, "INIT_DATE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow

        If e.Row.IsDataRow AndAlso e.Row.Band.Index = 0 Then
            Dim ORDR_QTY_ALLO As Int32 = Val(e.Row.Cells("ORDR_QTY_ALLO").Value & "")
            Dim ORDR_QTY_OPEN As Int32 = Val(e.Row.Cells("ORDR_QTY_OPEN").Value & "")
            Dim ORDR_QTY_BACK As Int32 = 0 ' Val(e.Row.Cells("ORDR_QTY_BACK").Value & "")

            With e.Row.Cells("ORDR_QTY_ALLO")
                If ORDR_QTY_ALLO < ORDR_QTY_OPEN + ORDR_QTY_BACK Then
                    .Appearance.BackColor = System.Drawing.Color.Red
                    .Appearance.ForeColor = System.Drawing.Color.White
                Else
                    .Appearance.BackColor = System.Drawing.Color.Empty
                    .Appearance.ForeColor = System.Drawing.Color.Empty
                    '.ToolTipText = ""
                End If
            End With

        End If
    End Sub

    Function De_Release(PICK_BATCH_NO As String, rowSOTPICK0 As DataRow, PICK_NOsD As List(Of String)) As Boolean

        ASCMAIN1.sql = sqlSOTPICK1 & $" and SOTPICK1.PICK_BATCH_NO = '{PICK_BATCH_NO}'"
        Fill_Records("SOTPICK1",,, ASCMAIN1.sql)

        Try
            BeginTrans()
            De_Release_PTs(PICK_NOsD, rowSOTPICK0)

            Update_Record_TDA("SOTPICK1")

            Dim r As Integer = 0

            ASCMAIN1.sql = $"Select * from SOTPICK1 where PICK_BATCH_NO = '{PICK_BATCH_NO}' and PICK_STATUS = 'P'"
            Dim rows() As DataRow = ASCDATA1.GetDataTable.Select("")
            If rows.Length = 0 Then

                Dim TRUCK_NO As String = rowSOTPICK0.Item("TRUCK_NO")
                Dim TRUCK_TYPE As String = rowSOTPICK0.Item("TRUCK_TYPE")

                ' At this point count up F, D and Cs if any Fs the F, else if any Ds the D else if any Cs then C else X
                ASCMAIN1.sql = "SELECT PICK_BATCH_NO
                                    , SUM (CASE WHEN PICK_STATUS = 'F' THEN 1 ELSE 0 END) F
                                    , SUM (CASE WHEN PICK_STATUS = 'D' THEN 1 ELSE 0 END) D
                                    , SUM (CASE WHEN PICK_STATUS = 'C' THEN 1 ELSE 0 END) C
                                    FROM SOTPICK1
                                    WHERE PICK_BATCH_NO = :PARM1
                                    GROUP BY PICK_BATCH_NO"

                Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", {PICK_BATCH_NO})

                Dim PICK_BATCH_STATUS As String = "X"
                If rowSOTPICK1 IsNot Nothing Then
                    If Val(rowSOTPICK1.Item("F") & String.Empty) > 0 Then
                        PICK_BATCH_STATUS = "F"
                    ElseIf Val(rowSOTPICK1.Item("D") & String.Empty) > 0 Then
                        PICK_BATCH_STATUS = "D"
                    ElseIf Val(rowSOTPICK1.Item("C") & String.Empty) > 0 Then
                        PICK_BATCH_STATUS = "C"
                    End If
                End If

                ASCMAIN1.sql = $"Update SOTPICK0 Set PICK_BATCH_STATUS = :PARM1, TRUCK_NO = NULL where TRUCK_NO = :PARM2 and PICK_BATCH_NO = :PARM3"
                r = -1
                r = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New String() {PICK_BATCH_STATUS, TRUCK_NO, PICK_BATCH_NO})
                If r <> 1 Then
                    Throw New Exception($"Could not Clear Truck {TRUCK_NO} and Status for Pick Batch No {PICK_BATCH_NO}")
                End If

                Dim sqlTRUCK_TYPE As String = ""
                If TRUCK_TYPE = "X" Then sqlTRUCK_TYPE = ", TRUCK_TYPE = 'R'"
                ASCMAIN1.sql = $"Update SOTTRCK1 Set PICK_BATCH_NO = NULL{sqlTRUCK_TYPE} where TRUCK_NO = :PARM1 and PICK_BATCH_NO = :PARM2"

                r = -1
                r = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {TRUCK_NO, PICK_BATCH_NO})
                If r <> 1 Then
                    Throw New Exception($"Could not Clear Pick Batch No {PICK_BATCH_NO} from Truck {TRUCK_NO}")
                End If
            End If


            Dim PICK_BATCH_STATUS_new As String = Get_PICK_BATCH_STATUS_new(PICK_BATCH_NO)

            ASCMAIN1.sql = $"Update SOTPICK0 Set PICK_BATCH_STATUS = :PARM1, LAST_DATE = SYSDATE, LAST_OPER = :PARM2 where PICK_BATCH_NO = :PARM3"
            r = -1
            r = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New String() {PICK_BATCH_STATUS_new, ASCMAIN1.USER_ID, PICK_BATCH_NO})
            If r <> 1 Then
                Throw New Exception($"Could not Reset Pick Batch Status to {PICK_BATCH_STATUS_new} for Pick Batch No {PICK_BATCH_NO}")
            End If

            CommitTrans($"{CStr(PICK_NOsD.Count)} Pick Tickets from Pick Batch {PICK_BATCH_NO} have been De-Released")

            Return True
        Catch ex As Exception
            Rollback()
            Return False
        End Try

    End Function

    Function Get_PICK_BATCH_STATUS_new(PICK_BATCH_NO As String) As String

        ASCMAIN1.sql = $"WITH PICK_BATCH_DATA AS (SELECT S0.PICK_BATCH_STATUS,
                                S1.PICK_BATCH_NO,
                                COUNT(DISTINCT CASE WHEN S1.PICK_STATUS = 'P' THEN S1.PICK_NO END) AS TICKETS_IN_PICK,
                                MAX(CASE WHEN S1.PICK_STATUS = 'P' THEN NVL(S1.PICK_REQ_RES, '0') END) AS REQRES,
                                SUM(CASE WHEN S1.PICK_STATUS = 'P' THEN S2.PICK_QTY END)  AS PICK_QTY,
                                SUM(CASE
                                        WHEN S1.PICK_STATUS = 'P' THEN NVL(S2.PICK_QTY, 0) - NVL(S2.PICK_QTY_CONF, 0) -
                                                                       CASE WHEN NVL(S2.PICK_QTY,0) <> 0 THEN NVL(S2.PICK_QTY_BACK, 0) ELSE 0 END END)   AS PICK_QTY_NOTYET,
                                COUNT(DISTINCT CASE WHEN S1.PICK_STATUS = 'F' THEN S1.PICK_NO END)                  AS FINALIZED_TICKETS,
                                COUNT(DISTINCT CASE WHEN S1.PICK_STATUS = 'C' THEN S1.PICK_NO END)                  AS CANCELED_TICKETS,
                                COUNT(DISTINCT CASE WHEN S1.PICK_STATUS = 'D' THEN S1.PICK_NO END)                  AS DEPICKED_TICKETS
                         FROM SOTPICK0 S0
                                  JOIN
                              SOTPICK1 S1 ON S0.PICK_BATCH_NO = S1.PICK_BATCH_NO
                                  JOIN
                              SOTPICK2 S2 ON S1.PICK_NO = S2.PICK_NO
                         WHERE S1.PICK_BATCH_NO = :PARM1
                         GROUP BY S0.PICK_BATCH_STATUS,
                                  S1.PICK_BATCH_NO)
                    SELECT * FROM (
                        SELECT
                            CASE
                                WHEN PICK_BATCH_STATUS IN ('F','C','X') THEN PICK_BATCH_STATUS
                                WHEN TICKETS_IN_PICK = 0 THEN 'X'
                                ELSE CASE
                                    WHEN PICK_QTY = PICK_QTY_NOTYET THEN
                                        CASE
                                            WHEN PICK_BATCH_STATUS = 'P' THEN 'P'
                                            ELSE 'O'
                                        END
                                    ELSE CASE
                                        WHEN PICK_QTY_NOTYET <= 0 THEN
                                            CASE
                                                WHEN REQRES = '1' THEN 'R'
                                                ELSE CASE
                                                    WHEN FINALIZED_TICKETS = 0 THEN 'N'
                                                    ELSE 'K'
                                                END
                                            END
                                        ELSE 'P'
                                    END
                                END
                            END AS PICK_BATCH_STATUS_NEW,
                            PBD.*
                        FROM PICK_BATCH_DATA PBD)"

        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", PICK_BATCH_NO)

        Return row.Item("PICK_BATCH_STATUS_NEW")

    End Function

    Private Sub grdICTSTATX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTSTATX.InitializeRow
        If e.Row.IsDataRow Then

            With e.Row.Cells("QTY_SHORT")
                If Val(.Value & "") <> 0 Then
                    .Appearance.ForeColor = System.Drawing.Color.Red
                    .Appearance.BackColor = System.Drawing.Color.Empty
                Else
                    .Appearance.ForeColor = System.Drawing.Color.Empty
                    .Appearance.BackColor = System.Drawing.Color.Empty
                    '.ToolTipText = ""
                End If
            End With
            'End If
        End If
    End Sub

    Private Sub grdICTSTATX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTSTATX.AfterRowActivate
        If grdICTSTATX.ActiveRow IsNot Nothing AndAlso grdICTSTATX.ActiveRow.IsDataRow Then
            Dim STYLE_CODE As String = grdICTSTATX.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTATX.ActiveRow.Cells("COLOR_CODE").Value
            grdSOTORDRX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdSOTORDRX.DisplayLayout.Bands(0).ColumnFilters("STYLE_CODE").FilterConditions.Add(FilterComparisionOperator.Equals, STYLE_CODE)
            grdSOTORDRX.DisplayLayout.Bands(0).ColumnFilters("COLOR_CODE").FilterConditions.Add(FilterComparisionOperator.Equals, COLOR_CODE)
        End If
    End Sub

    Private Sub chkShowHolds_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowHolds.CheckedChanged
        Click_Command("Refresh")
        If chkShowHolds.Checked Then
            Me.tabMain.Tabs("Pick Tickets").Visible = False
        Else
            Me.tabMain.Tabs("Pick Tickets").Visible = True
        End If
    End Sub

    Sub Release_Selected_Orders(GROUP_KEY As String)

        Dim ORDRs_Selected As Integer = dst.Tables("SOTORDQ1").Select($"GROUP_KEY = '{GROUP_KEY}' AND SEL='1'").Length
        'Dim ORDRs_Selected As Integer = dst.Tables("SOTORDQ1").Select($"SEL='1'").Length
        If ORDRs_Selected = 0 Then
            MsgBox("No Orders Selected", MsgBoxStyle.OkOnly, "Cannot Release")
            Exit Sub
        ElseIf ORDRs_Selected > MAX_ORDERS_TO_RELEASE Then
            MsgBox($"Max number of Orders Permitted to be Released to a Truck is {MAX_ORDERS_TO_RELEASE}", MsgBoxStyle.OkOnly, "Cannot Release")
            Exit Sub
        End If

        ORDR_NOs_Tried.Clear()

        Release_Orders(New String() {GROUP_KEY})
    End Sub
End Class
