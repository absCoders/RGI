Imports Infragistics.Win.UltraWinGrid
Imports System.Text
Imports System.Drawing.Printing

Public Class WHFQACT1
    Dim sqlWHTQACTX As String = ""

    Dim ORDR_NO As String
    Dim CUST_CODE As String = "KOHLS"

    Dim rowARTCUST1 As DataRow
    Dim SOTORDR0 As String
    Dim sqlSOTORDR0 As String
    Dim sqlSOTPICK2 As String
    Dim sqlSOTORDRS As String
    Dim sqlSOTRSRVS As String
    Dim ORDR_GROUP_NO As String
    Dim sqlSOTORDRP As String
    Dim sqlSOTORDR1 As String
    Dim SOTORDR0_ALL As String
    Dim ORDR_CUST_PO As String

    Dim sqlSOTORDRT As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        'Build_TempTable()

        With dst
            sqlWHTQACTX = "Select trunc(WHTQACT1.CART_CHECKED) CART_CHECKED, count(distinct WHTQACT1.CART_NO) CARTONS " & vbCrLf _
                & ", count( distinct WHTQACT1.CART_CHECKER) CHECKERS, sum(SOTCART1.CART_TOTAL_UNITS) UNITS" & vbCrLf _
                & ", sum(ERRORS) ERRORS, Trunc((sum(ERRORS) / sum(SOTCART1.CART_TOTAL_UNITS) * 1000))/10 PERCENT" & vbCrLf _
                & ", sum(WRONG_STYLE) WRONG_STYLE, sum(SHORT) SHORT, sum(OVER) OVER" & vbCrLf _
                & " From WHTQACT1, SOTCART1, " & vbCrLf _
                & "   (select CART_NO, count(UPC_CODE) ERRORS, " & vbCrLf _
                & "    sum(case when QTY_PACKED = 0 then 1 else 0 end) WRONG_STYLE, " & vbCrLf _
                & "    sum(case when QTY_PACKED > QTY_VERIFIED then 1 else 0 end) SHORT, " & vbCrLf _
                & "    sum(case when QTY_PACKED > 0 and QTY_PACKED < QTY_VERIFIED then 1 else 0 end) OVER" & vbCrLf _
                & "    from WHTQACT2 group by CART_NO) WHTQACT2" & vbCrLf _
                & " Where WHTQACT1.CART_NO = WHTQACT2.CART_NO(+)" & vbCrLf _
                & "  and WHTQACT1.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & "  and WHTQACT1.CART_CHECKED > sysdate - 31" & vbCrLf _
                & " group by trunc(WHTQACT1.CART_CHECKED)"
            ASCMAIN1.sql = sqlWHTQACTX.Replace("CART_CHECKED > sysdate - 31", "CART_CHECKED is null")
            Create_TDA(.Tables.Add, "WHTQACTX", "**", 0, False, "", 1)

            sqlSOTORDR0 = "Select 'O' ORDR_TYPE, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
                & ", SOTORDR0.CUST_DC_NO, SOTORDR0.ORDR_DEPT, EDT850T1.EDI_MERCH_TYPE, SOTORDR0.SALES_DIVISION_CODE, SOTORDR0.ORDR_DATE" & vbCrLf _
                & ", SOTORDR0.ORDR_SHIP_DATE,SOTORDR0. ORDR_CANCEL_DATE, SOTORDR0.ORDR_ORIG_SHIP_DATE, SOTORDR0.ORDR_ORIG_CANCEL_DATE" & vbCrLf _
                & ", SOTORDR0.WHSE_CODE, SOTORDR0.SREP_CODE" & vbCrLf _
                & ", SOTORDR0.ORDR_TYPE_CODE, SOTORDR0.ORDR_SOURCE, SOTORDR0.EDI_DOC_SEQ_NO" & vbCrLf _
                & ", SOTORDR0.ORDR_AMT, SOTORDR0.ORDR_AMT_OPEN, SOTORDR0.ORDR_AMT_PICK, SOTORDR0.ORDR_AMT_SHIP, SOTORDR0.ORDR_AMT_CANC" & vbCrLf _
                & ", SOTORDR0.ORDR_QTY, SOTORDR0.ORDR_QTY_OPEN, SOTORDR0.ORDR_QTY_PICK, SOTORDR0.ORDR_QTY_SHIP, SOTORDR0.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTORDR0.ORDR_CNT, SOTORDR0.ORDR_CNT_OPEN, SOTORDR0.ORDR_CNT_PICK" & vbCrLf _
                & ", SOTORDR0.ORDR_DATE_RECD, SOTORDR0.ORDR_PRIORITY, SOTORDR0.ORDR_ARRIVAL_DATE, SOTORDR0.ORDR_LAST_ARRIVAL_DATE" & vbCrLf _
                & ", SOTORDR0.ORDR_NO_MIN, SOTORDR0.ORDR_NO_MAX, SOTORDR0.ORDR_RELEASE_AVAIL_MIN, SOTORDR0.ORDR_RELEASE_AVAIL_MAX" & vbCrLf _
                & ", SOTORDRG.ORDR_REL_SHORT, SOTORDRG.ORDR_REL_SHORT_OPER, SOTORDRG.ORDR_REL_ACTION_DATE, SOTORDRG.ORDR_REL_ACTION_OPER" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "VAN", ", EDT850T1.EDI_CONS_NO", ", '0' EDI_CONS_NO") & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "VAN", ", SOTPCKP2.PACK_NO", ", ' ' PACK_NO") & vbCrLf _
                & " from SOTORDR0,EDT850T1,SOTORDRG" & IIf(ASCMAIN1.CLIENT = "VAN", ",SOTPCKP2", "")
            ASCMAIN1.sql = sqlSOTORDR0 & " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.CUST_CODE = ''"
            If ASCMAIN1.CLIENT = "VAN" Then
                ASCMAIN1.sql &= " and SOTPCKP2.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTPCKP2.PACK_GROUP_STATUS (+) = 'A'"
            End If
            ASCMAIN1.sql = "Select X.*, SOTORDR1.TERM_CODE, SOTORDR1.LAST_DATE, SOTORDR1.LAST_OPER, SOTORDR1.ORDR_SHIP_INSTR, SOTORDR1.ORDR_MESSAGE,  ARTCUST1.CUST_NAME, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
                & " from (" & ASCMAIN1.sql & ") X, ARTCUST1,SOTORDR1" _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE and SOTORDR1.ORDR_NO = X.ORDR_NO_MIN "
            SOTORDR0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add WAVE_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add EDI_LOAD_ID VARCHAR2(20)")
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT_ALLO_CUR NUMBER(13,2)")
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT_ALLO_FUT NUMBER(13,2)")
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT_ALLO_CXL NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_TYPE, ORDR_GROUP_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add EDI_PO_TYPE VARCHAR2(2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add QA_SCANNED NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add QA_CLOSED NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add QA_STARTED NUMBER(13,2)")
            ASCMAIN1.sql = "Select * from " & SOTORDR0
            Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "", 2)

            ASCMAIN1.sql = "SELECT " & vbCrLf _
                        & "    SOTPICK1.PICK_NO," & vbCrLf _
                        & "    SOTORDR1.ORDR_CUST_PO," & vbCrLf _
                        & "    SOTORDR1.ORDR_NO," & vbCrLf _
                        & "    SOTPICK1.INV_NO," & vbCrLf _
                        & "    SOTPICK1.SHIP_BOL_NO," & vbCrLf _
                        & "    SOTORDR1.CUST_DC_NO," & vbCrLf _
                        & "    SOTORDR1.CUST_STORE_NO," & vbCrLf _
                        & "    SOTORDR1.CUST_STORE_NAME," & vbCrLf _
                        & "    count(SOTCART1.CART_NO) CARTONS" & vbCrLf _
                        & "from" & vbCrLf _
                        & "    SOTPICK1" & vbCrLf _
                        & "    inner join SOTORDR1 on SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                        & "    inner join SOTCART1 on SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                        & "where" & vbCrLf _
                        & "    SOTORDR1.ORDR_GROUP_NO = :PARM1" & vbCrLf _
                        & "group by" & vbCrLf _
                        & "    SOTORDR1.ORDR_CUST_PO," & vbCrLf _
                        & "    SOTORDR1.ORDR_NO," & vbCrLf _
                        & "    SOTPICK1.PICK_NO," & vbCrLf _
                        & "    SOTPICK1.INV_NO," & vbCrLf _
                        & "    SOTPICK1.SHIP_BOL_NO," & vbCrLf _
                        & "    SOTORDR1.CUST_DC_NO," & vbCrLf _
                        & "    SOTORDR1.CUST_STORE_NO," & vbCrLf _
                        & "    SOTORDR1.CUST_STORE_NAME"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from WHTQACT1 where PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTQACT1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select WHTQACT2.* from WHTQACT2, WHTQACT1 where WHTQACT2.CART_NO = WHTQACT1.CART_NO and WHTQACT1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTQACT2", "**", 0, False, "V", 2)

            Create_Relation("WHTQACT1", "WHTQACT2", "CART_NO")
            .Tables("WHTQACT1").Columns.Add("REL_QTY", GetType(System.Int32), "SUM(CHILD(WHTQACT1_WHTQACT2).QTY_PACKED)")
            .Tables("WHTQACT1").Columns.Add("PICKED_QTY", GetType(System.Int32), "SUM(CHILD(WHTQACT1_WHTQACT2).QTY_PICK_SCAN)")
            .Tables("WHTQACT1").Columns.Add("VERIFIED_QTY", GetType(System.Int32), "SUM(CHILD(WHTQACT1_WHTQACT2).QTY_VERIFIED)")


            'For Each A As String In New String() {"CUR", "FUT", "CXL"}
            '    .Tables("SOTORDR0").Columns.Add("PCT_ALLO_" & A, GetType(System.Decimal), "IIF(ORDR_AMT=0,0,100*ORDR_AMT_ALLO_" & A & "/ORDR_AMT)")
            'Next

            'Create_TDA(.Tables.Add, "SOTORDRG", "*")

            'Dim TBL As DataTable = .Tables("SOTORDR0").Clone
            'TBL.TableName = "SOTCORDG"
            '.Tables.Add(TBL)

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 0, False, "", 1)
            Create_TDA(.Tables.Add, "ICTWHSE1", "*", 0, False, "", 1)
            Create_TDA(.Tables.Add, "SOTSVIA1", "*", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTSHIPB", "*", 0, True)

        End With

        grdWHTQACTX.DataSource = dst.Tables("WHTQACTX")
        grdSOTORDR0.DataSource = dst.Tables("SOTORDR0")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdWHTQACT1.DataSource = dst.Tables("WHTQACT1")

        Fill_Records("ICTWHSE1")
        Fill_Records("SOTSVIA1")

        With grdWHTQACTX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                If gcol.Key = "min_qty_ecom" Or gcol.Key = "max_qty_ecom" Or gcol.Key = "pct_qty_ecom" Or gcol.Key = "not_inseason" Then
                    gcol.CellActivation = Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"CART_CHECKED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ERRORS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Red
                ElseIf New String() {"PERCENT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"FUTURE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With



        'grdSOTORDR0.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                'If gcol.Key = "PO_QTY" Then
                '    gcol.CellAppearance.BackColor = Drawing.Color.LightYellow
                '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                'Else
                '    '  gcol.CellAppearance.BackColor = Drawing.Color.Beige
                '    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                'End If

                If gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Or gcol.Key.StartsWith("ORDR_AMT") Or gcol.Key.StartsWith("CUST_CREDIT") Then
                    gcol.Hidden = True
                End If
                ', "ORDR_ARRIVAL_DATE", "ORDR_LAST_ARRIVAL_DATE"
                If New String() {"ORDR_DATE_RECD", "ORDR_PRIORITY",
                                 "ORDR_RELEASE_AVAIL_MIN", "ORDR_RELEASE_AVAIL_MAX", "ORDR_REL_SHORT", "ORDR_REL_SHORT_OPER",
                                 "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_DEPT", "EDI_MERCH_TYPE", "ORDR_CNT_OPEN", "ORDR_CNT_PICK",
                                 "ORDR_REL_ACTION_DATE", "ORDR_REL_ACTION_OPER", "TERM_CODE", "LAST_DATE", "LAST_OPER", "ORDR_SHIP_INSTR", "ORDR_MESSAGE", "EDI_PO_TYPE",
                                 "CUST_CITY", "CUST_STATE", "CUST_COUNTRY", "SALES_DIVISION_CODE", "SREP_CODE"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                End If

                If New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf gcol.Key.StartsWith("ORDR_AMT_ALLO_") Or gcol.Key.StartsWith("PCT_ALLO_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"QA_SCANNED", "QA_CLOSED", "QA_STARTED"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.MediumPurple
                ElseIf New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"ORDR_DATE", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ORIG_SHIP_DATE", "ORDR_ORIG_CANCEL_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                ElseIf New String() {"ORDR_CUST_PO", "CUST_DC_NO", "ORDR_DEPT", "WHSE_CODE", "EDI_MERCH_TYPE", "WAVE_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Yellow
                ElseIf New String() {"CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "SALES_DIVISION_CODE", "SREP_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                ElseIf New String() {"CUST_CITY", "CUST_STATE", "CUST_COUNTRY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                Else
                    gcol.Header.Appearance.BackColor = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next

            .Columns("EDI_CONS_NO").Hidden = Not (ASCMAIN1.CLIENT = "VAN")
        End With

        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")
        Create_Summary(grdSOTORDR0, New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_CNT", "ORDR_CNT_OPEN", "ORDR_CNT_PICK"}, , , "#,##0")

        Show_Filter(grdSOTORDR0, True)
        grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = False

        grdSOTORDR0.DisplayLayout.Bands(0).Columns("ORDR_TYPE_CODE").Hidden = Not (ASCMAIN1.CLIENT = "RGI")
        grdSOTORDR0.DisplayLayout.Bands(0).Columns("ORDR_SOURCE").Hidden = Not (ASCMAIN1.CLIENT = "RGI")
        grdSOTORDR0.DisplayLayout.Bands(0).Columns("WAVE_NO").Hidden = Not (ASCMAIN1.CLIENT = "VAN")
        grdSOTORDR0.DisplayLayout.Bands(0).Columns("EDI_LOAD_ID").Hidden = Not (ASCMAIN1.CLIENT = "VAN")
        grdSOTORDR0.DisplayLayout.Bands(0).Columns("TERM_CODE").Hidden = Not (ASCMAIN1.CLIENT = "RGI")

        'grdWHTQACT1
        With grdWHTQACT1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"PICK_NO", "GUN_ID", "CART_TOTAL_UNITS", "CART_UNITS_CHECKED"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                End If

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next
        End With
        With grdWHTQACT1.DisplayLayout.Bands(1)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"CART_NO", "ORDR_NO", "ORDR_LNO", "STYLE_CODE", "COLOR_CODE", "STYLE_UOM"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                End If

                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

            Next
        End With

        '-------------
        Show_Filter(grdSOTPICK1)
        Show_Filter(grdWHTQACT1)

        'spl.Panel1Collapsed = True
        'splStats.Panel2Collapsed = True
        SplitContainer1.Panel2Collapsed = True

        lblDefaultPrinter.Text = Default_Printer()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"

                If Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    EMsg &= vbCrLf & "You must specify Customer PO"
                Else
                    ASCMAIN1.sql = "select WHTQACT1.*, SOTORDR0.ORDR_GROUP_NO from SOTORDR0, SOTORDR1, SOTPICK1, WHTQACT1" & vbCrLf _
                                & " where SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                                & " and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                                & " and WHTQACT1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                                & " and SOTORDR0.ORDR_CUST_PO = :PARM1" & vbCrLf _
                                & " and rownum < 2"
                    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", Absx1.txtFor("ORDR_CUST_PO").Text)
                    If row Is Nothing Then
                        EMsg &= vbCrLf & "This PO doesn't have any scans against it, select different PO"
                    Else
                        ORDR_GROUP_NO = row("ORDR_GROUP_NO")
                    End If

                End If

                If EMsg = "" Then
                    ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                End If

            Case "Update"
                MsgBox("No records to Update", MsgBoxStyle.Information, "Info")

            Case "Print"
                EMsg = "Nothing to Print"

            Case "Cancel"


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View", "Load"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
                Clear_Record()

            Case "Done", "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Report("")


            Case "Refresh"
                Clear_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    '.Items("Update").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = not_iScreenMode
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                End With

                '.Groups("Update").Visible = ScreenMode
            End With
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        'splShipments.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Preparing Data ....")

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTQACTX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Fill_Records("WHTQACTX",, True, sqlWHTQACTX)
        Sort_grdColumns(grdWHTQACTX, "cart_checked")

        grdSOTORDR0.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Load_SOTORDR0("", CUST_CODE)
        'Setup_SOTORDR0()

        EnforceConstraints(True)

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSP("WJZOP", "VV", New String() {Me.Name, ASCMAIN1.USER_ID}, New String() {"FORM_NAME", "USER_ID"})
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data...")
        'grdSOTORDR0.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Fill_Records("SOTPICK1", ORDR_GROUP_NO, True)

        EnforceConstraints(True)

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSP("WJZOP", "VV", New String() {Me.Name, ASCMAIN1.USER_ID}, New String() {"FORM_NAME", "USER_ID"})
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Update")
        Dim sql As String

        BeginTrans()

        ''SOTPICK1.PICK_SHIP_DATE, SOTPICK1.PICK_PRIORITY, SOTPICK1.PICK_COMPLEXITY
        'Dim dtTable As DataTable = dst.Tables("WHTPICKP").GetChanges()
        'If dtTable IsNot Nothing Then
        '    For Each row As DataRow In dtTable.Rows
        '        If Not row.Item("PICK_SHIP_DATE") & "" = "" Then
        '            sql = "Update SOTPICK1 SET PICK_SHIP_DATE = :PARM1 " & vbCrLf _
        '                & ", PICK_PRIORITY = :PARM2 " & vbCrLf _
        '                & ", PICK_COMPLEXITY = :PARM3 " & vbCrLf _
        '                & "Where  PICK_NO = :PARM4 "
        '            ASCDATA1.ExecuteSQL(sql, "VVVV", New String() {String.Format("{0:dd-MMM-yy}", row.Item("PICK_SHIP_DATE")), row.Item("PICK_PRIORITY"), row.Item("PICK_COMPLEXITY"), row.Item("PICK_NO")})
        '        Else
        '            sql = "Update SOTPICK1 SET PICK_PRIORITY = :PARM1 " & vbCrLf _
        '                & ", PICK_COMPLEXITY = :PARM2 " & vbCrLf _
        '                & "Where  PICK_NO = :PARM3 "
        '            ASCDATA1.ExecuteSQL(sql, "VVV", New String() {row.Item("PICK_PRIORITY"), row.Item("PICK_COMPLEXITY"), row.Item("PICK_NO")})
        '        End If
        '    Next
        'End If
        ASCMAIN1.MultiTask_Release("", 0, 3)
        CommitTrans("Update Complete")

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
        Load_Popup_Menu(grdWHTQACTX, "S", "Show Filter")

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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTPACKX"


            End Select

        End If
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
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                    If Not Nothing Is grow.ChildBands Then
                        ' Loop throgh each of the child bands.
                        For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                            grow2.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                            grow2.Update()
                        Next
                    End If
                Next
            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                    If Not Nothing Is grow.ChildBands Then
                        ' Loop throgh each of the child bands.
                        For Each grow2 As UltraWinGrid.UltraGridRow In grow.ChildBands(0).Rows
                            grow2.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                            grow2.Update()
                        Next
                    End If
                Next
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
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("View")
        End Select
    End Sub
#End Region

    Sub Load_SOTORDR0(Optional PARM1 As String = "", Optional CUST_CODE As String = "")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Order Summary", "")

        If CUST_CODE <> "" Then ' ScreenMode Then
            ASCMAIN1.sql = sqlSOTORDR0
            Dim sqlw As String = " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf
            If ASCMAIN1.CLIENT = "VAN" Then
                sqlw &= " and SOTPCKP2.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTPCKP2.PACK_GROUP_STATUS (+) = 'A'"
                sqlw &= " AND SOTORDR0.ORDR_QTY_PICK + SOTORDR0.ORDR_QTY_SHIP <> 0" & vbCrLf
            End If

            grdSOTORDR0.Text = "Orders for " & CUST_CODE & "; Status: Shipped"

            ASCMAIN1.sql &= sqlw


        Else

            Dim SQLW As String = ""

            If ASCMAIN1.CLIENT = "VAN" Then
                SQLW &= " and SOTPCKP2.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTPCKP2.PACK_GROUP_STATUS (+) = 'A'"
            End If

            ASCMAIN1.sql = sqlSOTORDR0
            PARM1 = Replace(Replace(PARM1, ";", ""), "'", "")

            Dim sqlORDR_STATUS As String = ""

            Select Case grdSOTORDR0.Tag & ""
                Case ""
                    grdSOTORDR0.Text = "Orders which are either Open or In Pick"
                    ASCMAIN1.sql &= " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and (SOTORDR0.ORDR_CNT_OPEN <> 0 or SOTORDR0.ORDR_CNT_PICK <> 0)"
                    ASCMAIN1.sql &= SQLW
                    'ASCMAIN1.sql &= sqlReservations

                Case "SREP_CODE"
                    grdSOTORDR0.Text = "Open Orders for Sales Rep " & PARM1
                    ASCMAIN1.sql &= " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.ORDR_GROUP_NO in " _
                        & " (Select Distinct ORDR_GROUP_NO from SOTORDR1 " _
                        & " where ORDR_STATUS >= 'O' and ORDR_STATUS <= 'P'" _
                        & "   and SREP_CODE = '" & PARM1 & "')"
                    ASCMAIN1.sql &= SQLW
                    'ASCMAIN1.sql &= Replace(sqlReservations, " group by ", " and SOTRSRV1.SREP_CODE = '" & PARM1 & "'" & vbCrLf & " group by ")

                Case "ORDR_CUST_PO"
                    grdSOTORDR0.Text = "All Customer Orders using Customer PO " & PARM1
                    ASCMAIN1.sql &= " where EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR0.EDI_DOC_SEQ_NO and SOTORDRG.ORDR_GROUP_NO (+) = SOTORDR0.ORDR_GROUP_NO and SOTORDR0.ORDR_GROUP_NO in " _
                        & " (Select Distinct ORDR_GROUP_NO from SOTORDR1 where ORDR_CUST_PO = '" & PARM1 & "')"
                    ASCMAIN1.sql &= SQLW
            End Select

        End If

        Dim sqlWHTQACT1 As String = "" _
            & "select ORDR_GROUP_NO, sum(Count) QA_SCANNED" & vbCrLf _
            & ", sum(case when STATUS = 'Closed' then COUNT else 0 end) QA_CLOSED" & vbCrLf _
            & ", sum(case when STATUS <> 'Closed' then COUNT else 0 end) QA_STARTED" & vbCrLf _
            & " from ( " & vbCrLf _
            & " select SOTSHIP1.ORDR_GROUP_NO, nvl(WHTQACT1.process_status,'Closed') STATUS, count(1) COUNT " & vbCrLf _
            & " from WHTQACT1, SOTPICK1, SOTSHIP1" & vbCrLf _
            & " where WHTQACT1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & " and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & " group by SOTSHIP1.ORDR_GROUP_NO, nvl(WHTQACT1.PROCESS_STATUS,'Closed'))" & vbCrLf _
            & " group by ORDR_GROUP_NO"



        ASCMAIN1.sql = "Select X.*, SOTORDR1.TERM_CODE, SOTORDR1.LAST_DATE, SOTORDR1.LAST_OPER, SOTORDR1.ORDR_SHIP_INSTR, SOTORDR1.ORDR_MESSAGE" & vbCrLf _
            & ", ARTCUST1.CUST_NAME, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
            & ", NULL WAVE_NO, NULL EDI_LOAD_ID" & vbCrLf _
            & ", SOTORDR1.EDI_PO_TYPE, QA_SCANNED, QA_CLOSED, QA_STARTED" & vbCrLf _
            & " from (" & ASCMAIN1.sql & ") X,ARTCUST1,SOTORDR1, (" & sqlWHTQACT1 & ") WHTQACT1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO (+) = X.ORDR_NO_MIN" & vbCrLf _
            & "   and WHTQACT1.ORDR_GROUP_NO (+) = X.ORDR_GROUP_NO"
        'Fill_Records("SOTORDR0", "", , ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Delete from " & SOTORDR0)
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0 & " " & ASCMAIN1.sql)

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select ORDR_GROUP_NO, MIN (WAVE_NO) WAVE_NO, MIN (EDI_LOAD_ID) EDI_LOAD_ID" & vbCrLf _
                & "   from SOTSHIP1 where ORDR_GROUP_NO in " & vbCrLf _
                & "    (Select ORDR_GROUP_NO from " & SOTORDR0 & " where ORDR_TYPE = 'O')" & vbCrLf _
                & "   group by ORDR_GROUP_NO;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set WAVE_NO = R1.WAVE_NO, EDI_LOAD_ID = R1.EDI_LOAD_ID" & vbCrLf _
                & "    where ORDR_TYPE = 'O' and ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        Fill_Records("SOTORDR0")
        Sort_grdColumns(grdSOTORDR0, "ORDR_SHIP_DATE".ToLower)

        grdSOTORDR0.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub


    Function Default_Printer()
        Dim settings As New PrinterSettings
        For Each printer As String In PrinterSettings.InstalledPrinters
            settings.PrinterName = printer
            If settings.IsDefaultPrinter Then
                Return printer
            End If
        Next
        Return String.Empty
    End Function

    Sub Print_Report(ByVal PICK_NO As String)

        Try

        Catch ex As Exception
            Rollback(ex.Message)
            EMsg = EMsg & vbCrLf & "Unable to Print Report"
        End Try

    End Sub


    Private Sub grdSOTORDR0_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTORDR0.DoubleClickRow
        Absx1.txtFor("ORDR_CUST_PO").Text = e.Row.Cells("ORDR_CUST_PO").Text
        Click_Command("View")
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICK1.AfterRowActivate
        load_WHTQACT()
    End Sub
    Private Sub load_WHTQACT()

        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdWHTQACT1.Visible = False
            grdWHTQACT1.Text = ""
        Else
            grdWHTQACT1.Visible = True
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            EnforceConstraints(False)
            Fill_Records("WHTQACT1", PICK_NO, True)
            Fill_Records("WHTQACT2", PICK_NO, True)
            EnforceConstraints(True)
            grdWHTQACT1.Text = String.Format("Cartons PO {0}, DC {1}  Store {2}", grdSOTPICK1.ActiveRow.Cells("ORDR_CUST_PO").Value, grdSOTPICK1.ActiveRow.Cells("CUST_DC_NO").Value, grdSOTPICK1.ActiveRow.Cells("CUST_STORE_NO").Value)

            Sort_grdColumns(grdWHTQACT1, "CART_NO")
        End If

        'SOTORDR1.ORDR_CUST_PO," & vbCrLf _
        '                & "    SOTORDR1.ORDR_NO," & vbCrLf _
        '                & "    SOTPICK1.INV_NO," & vbCrLf _
        '                & "    SOTPICK1.SHIP_BOL_NO," & vbCrLf _
        '                & "    SOTORDR1.CUST_DC_NO," & vbCrLf _
        '                & "    SOTORDR1.CUST_STORE_NO," & vbCrLf _
        '                & "    SOTORDR1.CUST_STORE_NAME," & vbCrLf _

    End Sub

End Class
