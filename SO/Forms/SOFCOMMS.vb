Public Class SOFCOMMS

    Private CUST_CODE As String = String.Empty
    Private SREP_CODE As String = String.Empty
    Private OPS_YYYYPP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)

    Private rowSOTSREP1 As DataRow
    Private rowARTCUST1 As DataRow

    Private dvSrepCode As DataView
    Private dvCustCode As DataView
    Private SOTCOMMS As String = String.Empty

    Private SOTCOMH1 As String = String.Empty
    Private SOTCOMH4 As String = String.Empty
    Private SOTCOMH5 As String = String.Empty
    Private SOTCOMH6 As String = String.Empty
    Private SOTINVHS As String = String.Empty

    Private querySOTCOMH1 As String = String.Empty
    Private querySOTCOMH4 As String = String.Empty
    Private querySOTCOMH5 As String = String.Empty
    Private querySOTCOMH6 As String = String.Empty
    Private querySOTINVHS As String = String.Empty

    Private refreshWorktables As Boolean = True
    Private commissionsDone As Boolean = False


    ' ADD COMM PAID IND TO SOTINVH1
    ' Credits in a period should go against the original sales person.
    ' Does a credit point to the original invoice??
    ' After invoicing is written to update SOTINVHS
    '   we need to see how to update SOTINVHS temp table when the grid SOTCOMH1 is updated


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SOFINVHI" Then
            InquiryMode = True
        End If

        With dst

            querySOTCOMH1 = " SELECT SOTCOMH1.*, SOTSREP1.SREP_NAME" _
              & " FROM SOTCOMH1, SOTSREP1" _
              & " WHERE SOTSREP1.SREP_CODE (+) = SOTCOMH1.SREP_CODE" _
              & " AND SOTCOMH1.OPS_YYYYPP = '" & OPS_YYYYPP & "'"
            SOTCOMH1 = ASCMAIN1.Temp_Table(querySOTCOMH1 & " and rownum < 1")
            ASCDATA1.ExecuteSQL("alter table " & SOTCOMH1 & " add primary key (OPS_YYYYPP, SREP_CODE)")
            Create_TDA(.Tables.Add, SOTCOMH1, "*")

            querySOTCOMH4 = "SELECT SOTCOMH4.*, SOTSREP1.SREP_NAME, ICTSGRP1.STYLE_GROUP_DESC" _
             & " FROM SOTCOMH4, SOTSREP1, ICTSGRP1" _
             & " WHERE SOTSREP1.SREP_CODE (+) = SOTCOMH4.SREP_CODE" _
             & " AND ICTSGRP1.STYLE_GROUP_CODE (+) = SOTCOMH4.STYLE_GROUP_CODE" _
             & " AND SOTCOMH4.OPS_YYYYPP = '" & OPS_YYYYPP & "'"
            SOTCOMH4 = ASCMAIN1.Temp_Table(querySOTCOMH4 & " and rownum < 1")
            ASCDATA1.ExecuteSQL("alter table " & SOTCOMH4 & " add primary key (OPS_YYYYPP, SREP_CODE, STYLE_GROUP_CODE)")
            Create_TDA(.Tables.Add, SOTCOMH4, "*")

            querySOTCOMH5 = " SELECT SOTCOMH5.*, SOTSREP1.SREP_NAME, ARTCUST1.CUST_NAME, SOTCOMH1.SREP_COMM_RATE SREP_COMM_RATE_M" _
             & " FROM SOTCOMH5, SOTSREP1, ARTCUST1, SOTCOMH1" _
             & " WHERE SOTSREP1.SREP_CODE (+) = SOTCOMH5.SREP_CODE" _
             & " AND ARTCUST1.CUST_CODE (+) = SOTCOMH5.CUST_CODE" _
             & " AND SOTCOMH1.SREP_CODE (+) = SOTCOMH5.SREP_CODE" _
             & " AND SOTCOMH1.OPS_YYYYPP (+) = SOTCOMH5.OPS_YYYYPP" _
             & " AND SOTCOMH5.OPS_YYYYPP = '" & OPS_YYYYPP & "'"
            SOTCOMH5 = ASCMAIN1.Temp_Table(querySOTCOMH5 & " and rownum < 1")
            ASCDATA1.ExecuteSQL("alter table " & SOTCOMH5 & " add primary key (OPS_YYYYPP, SREP_CODE, CUST_CODE)")
            Create_TDA(.Tables.Add, SOTCOMH5, "*")

            querySOTCOMH6 = " SELECT SOTCOMH6.*, SOTSREP1.SREP_NAME, ARTCUST1.CUST_NAME, ICTSGRP1.STYLE_GROUP_DESC" _
             & " FROM SOTCOMH6, SOTSREP1, ARTCUST1, ICTSGRP1" _
             & " WHERE SOTSREP1.SREP_CODE (+) = SOTCOMH6.SREP_CODE" _
             & " AND ARTCUST1.CUST_CODE (+) = SOTCOMH6.CUST_CODE" _
             & " AND ICTSGRP1.STYLE_GROUP_CODE (+) = SOTCOMH6.STYLE_GROUP_CODE " _
             & " AND SOTCOMH6.OPS_YYYYPP = '" & OPS_YYYYPP & "'"
            SOTCOMH6 = ASCMAIN1.Temp_Table(querySOTCOMH6 & " and rownum < 1")
            ASCDATA1.ExecuteSQL("alter table " & SOTCOMH6 & " add primary key (OPS_YYYYPP, SREP_CODE, CUST_CODE, STYLE_GROUP_CODE)")
            Create_TDA(.Tables.Add, SOTCOMH6, "*")

            querySOTINVHS = "SELECT * FROM SOTINVHS" _
                & " WHERE (INV_TYPE, INV_NO) IN (SELECT INV_TYPE, INV_NO FROM SOTINVH1 WHERE ORDR_YYYYPP_UPDATED = '" & OPS_YYYYPP & "')"
            SOTINVHS = ASCMAIN1.Temp_Table(querySOTINVHS & " and rownum < 1")
            ASCDATA1.ExecuteSQL("alter table " & SOTINVHS & " add primary key (INV_TYPE, INV_NO, SALES_DIVISION_CODE, SREP_CODE)")
            Create_TDA(.Tables.Add, SOTINVHS, "*")

            Create_TDA(.Tables.Add, "SOTSREP1", "*", 0, False)
            Create_TDA(.Tables.Add, "ICTSGRP1", "*", 0, False)
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 0, False)

            ASCMAIN1.sql = "Select SOTINVH2.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.STYLE_GROUP_CODE" _
                & " from SOTINVH2, ICTSTYL1" _
                & " where ICTSTYL1.STYLE_CODE (+) = SOTINVH2.STYLE_CODE" _
                & " and SOTINVH2.INV_TYPE = :PARM1 and SOTINVH2.INV_NO = :PARM2"
            Create_TDA(.Tables.Add, "SOTINVH2", "**", 0, True, "VV", 3)
            .Tables("SOTINVH2").Columns.Add("EXT_NET", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")

            SOTCOMMS = TAC.SOCCOMMS.NYAGCommissionsWorktable
            ASCMAIN1.sql = "Select * from " & TAC.SOCCOMMS.NYAGCommissionsWorktable
            Create_TDA(.Tables.Add, SOTCOMMS, "*", 0, False, "")
 
            ASCMAIN1.sql = "SELECT DISTINCT INV_TYPE, INV_NO, CUST_CODE FROM " & SOTCOMMS
            Create_TDA(.Tables.Add, "SOTINVH1", ASCMAIN1.sql, 0, False, String.Empty, 0)

            .Relations.Add("SOTSREP1_SOTINVHC", dst.Tables("SOTSREP1").Columns("SREP_CODE"), dst.Tables(SOTCOMMS).Columns("SREP_CODE"))
            .Relations.Add("ARTCUST1_SOTINVHC", dst.Tables("ARTCUST1").Columns("CUST_CODE"), dst.Tables(SOTCOMMS).Columns("CUST_CODE"))
            .Relations.Add("SOTINVH1_SOTINVHC", _
                           New DataColumn() {dst.Tables("SOTINVH1").Columns("INV_TYPE"), dst.Tables("SOTINVH1").Columns("INV_NO")},
                           New DataColumn() {dst.Tables(SOTCOMMS).Columns("INV_TYPE"), dst.Tables(SOTCOMMS).Columns("INV_NO")})

            .Tables("SOTSREP1").Columns.Add("TOT_COMM", GetType(System.Decimal), "SUM(CHILD.SREP_COMM_AMT)")
            .Tables("ARTCUST1").Columns.Add("TOT_COMM", GetType(System.Decimal), "SUM(CHILD.SREP_COMM_AMT)")

            .Tables("SOTSREP1").Columns.Add("TOT_SALES", GetType(System.Decimal), "SUM(CHILD.ORDR_AMT_SHP)")
            .Tables("ARTCUST1").Columns.Add("TOT_SALES", GetType(System.Decimal), "SUM(CHILD.ORDR_AMT_SHP)")

            .Tables("SOTINVH1").Columns.Add("TOT_COMM", GetType(System.Decimal), "SUM(CHILD.SREP_COMM_AMT)")

            .Relations.Add("SOTCOMH5_SOTCOMH6", _
                           New DataColumn() {dst.Tables(SOTCOMH5).Columns("OPS_YYYYPP"), dst.Tables(SOTCOMH5).Columns("SREP_CODE"), dst.Tables(SOTCOMH5).Columns("CUST_CODE")},
                           New DataColumn() {dst.Tables(SOTCOMH6).Columns("OPS_YYYYPP"), dst.Tables(SOTCOMH6).Columns("SREP_CODE"), dst.Tables(SOTCOMH6).Columns("CUST_CODE")})

        End With

        dvSrepCode = New DataView(dst.Tables("SOTSREP1"))
        'dvSrepCode.RowFilter = "ISNULL(TOT_COMM, 0) <> 0 and ISNULL(TOT_SALES, 0) <> 0"
        grdSOTINVHR.DataSource = dvSrepCode

        dvCustCode = New DataView(dst.Tables("ARTCUST1"))
        'dvCustCode.RowFilter = "ISNULL(TOT_COMM, 0) <> 0 and ISNULL(TOT_SALES, 0) <> 0"
        grdSOTINVHC.DataSource = dvCustCode

        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
        grdSOTINVH2.DataSource = dst.Tables("SOTINVH2")
        grdSOTCOMH1.DataSource = dst.Tables(SOTCOMH1)
        grdSOTCOMH4.DataSource = dst.Tables(SOTCOMH4)
        grdSOTCOMH5.DataSource = dst.Tables(SOTCOMH5)

        Create_Summary(grdSOTINVH1, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1, "TOT_COMM", "Sum")
        Create_Summary(grdSOTINVH1, "INV_NO", "Count", "SOTINVH1_SOTINVHC")
        Create_Summary(grdSOTINVH1, "ORDR_AMT_SHP", "Sum", "SOTINVH1_SOTINVHC")
        Create_Summary(grdSOTINVH1, "SREP_COMM_AMT", "Sum", "SOTINVH1_SOTINVHC")

        Create_Summary(grdSOTINVH2, "INV_LNO", "Count")
        Create_Summary(grdSOTINVH2, "EXT_NET", "Sum")

        Create_Summary(grdSOTINVHR, "SREP_CODE", "Count")
        Create_Summary(grdSOTINVHR, "TOT_COMM", "Sum")
        Create_Summary(grdSOTINVHR, "TOT_SALES", "Sum")
        Create_Summary(grdSOTINVHR, "INV_NO", "Count", "SOTSREP1_SOTINVHC")
        Create_Summary(grdSOTINVHR, "ORDR_AMT_SHP", "Sum", "SOTSREP1_SOTINVHC")
        Create_Summary(grdSOTINVHR, "SREP_COMM_AMT", "Sum", "SOTSREP1_SOTINVHC")

        Create_Summary(grdSOTINVHC, "CUST_CODE", "Count")
        Create_Summary(grdSOTINVHC, "TOT_COMM", "Sum")
        Create_Summary(grdSOTINVHC, "TOT_SALES", "Sum")
        Create_Summary(grdSOTINVHC, "INV_NO", "Count", "ARTCUST1_SOTINVHC")
        Create_Summary(grdSOTINVHC, "ORDR_AMT_SHP", "Sum", "ARTCUST1_SOTINVHC")
        Create_Summary(grdSOTINVHC, "SREP_COMM_AMT", "Sum", "ARTCUST1_SOTINVHC")

        Fill_Records("SOTSREP1", "", True, "SELECT * FROM SOTSREP1")
        Fill_Records("ICTSGRP1", "", True, "SELECT * FROM ICTSGRP1")
        Fill_Records("ARTCUST1", "", True, "SELECT * FROM ARTCUST1")


        commissionsDone = Val(ASCDATA1.GetDataValue("select count(*) from SOTCOMMS WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "'") & String.Empty) > 0

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("SOTCOMMS", OPS_YYYYPP) Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If EntryMode = "E" Then
                    If MessageBox.Show("Do you want to recalculate commissions applying your changes before printing report?", "Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                        Proceed("Recalculate")
                    End If
                End If

            Case "Recalculate"
                If MessageBox.Show("Do you want to recalculate commissions applying your changes?", "Recalulate", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Print"
                If EntryMode = "E" Then
                    If MessageBox.Show("Do you want to recalculate commissions applying your changes before printing report?", "Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                        Proceed("Recalculate")
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Edit", "View"
                refreshWorktables = True
                If eItemKey = "View" Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Recalculate"
                refreshWorktables = False
                EntryMode = "E"
                If UpdateWorkCommissionTables() Then
                    Load_Record()
                    Mode_Settings(True)
                End If
                refreshWorktables = True
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Recalculate").Settings.Enabled = iScreenMode

                    .Items("Edit").Visible = (Not InquiryMode) Or (Not commissionsDone)
                    .Items("Recalculate").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                    '.Items("Print").Visible = (InquiryMode Or EntryMode = "V")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                End With
            End With
        End If


        If ScreenMode Then
            If InquiryMode Or (EntryMode = "V") Then
                For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In _
                    New Infragistics.Win.UltraWinGrid.UltraGrid() {grdSOTINVHR, grdSOTINVHC, grdSOTCOMH1, grdSOTCOMH4, grdSOTCOMH5}
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                Next
            Else
                For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In _
                    New Infragistics.Win.UltraWinGrid.UltraGrid() {grdSOTCOMH4, grdSOTCOMH5}
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                Next
                For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In _
                    New Infragistics.Win.UltraWinGrid.UltraGrid() {grdSOTCOMH1}
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                Next
            End If
        End If

        splTab.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()

        EnforceConstraints(False)
        ToggleDataTableExpressions(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTINVH1", "SOTINVH2", SOTCOMH1, SOTCOMH4, SOTCOMH5, SOTCOMH6, SOTCOMMS}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        ToggleDataTableExpressions(True)
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        Try
            EnforceConstraints(False)
            ToggleDataTableExpressions(False)

            ASCMAIN1.Progress("Now Loading Data ...")
            If refreshWorktables Then
                ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTCOMH1)
                ASCDATA1.ExecuteSQL("INSERT INTO " & SOTCOMH1 & " " & querySOTCOMH1)
                Sort_grdColumns(grdSOTCOMH1, "SREP_CODE")
            End If
            Fill_Records(SOTCOMH1, "", True, "SELECT * FROM " & SOTCOMH1)

            ASCMAIN1.Progress("-", "SOTCOMH4")
            If refreshWorktables Then
                ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTCOMH4)
                ASCDATA1.ExecuteSQL("INSERT INTO " & SOTCOMH4 & " " & querySOTCOMH4)
                Sort_grdColumns(grdSOTCOMH4, "SREP_CODE,STYLE_GROUP_CODE")
            End If
            Fill_Records(SOTCOMH4, "", True, "SELECT * FROM " & SOTCOMH4)

            ASCMAIN1.Progress("-", "SOTCOMH5")
            If refreshWorktables Then
                ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTCOMH5)
                ASCDATA1.ExecuteSQL("INSERT INTO " & SOTCOMH5 & " " & querySOTCOMH5)
                Sort_grdColumns(grdSOTCOMH5, "SREP_CODE,CUST_CODE")
            End If
            Fill_Records(SOTCOMH5, "", True, "SELECT * FROM " & SOTCOMH5)

            ASCMAIN1.Progress("-", "SOTCOMH6")
            If refreshWorktables Then
                ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTCOMH6)
                ASCDATA1.ExecuteSQL("INSERT INTO " & SOTCOMH6 & " " & querySOTCOMH6)
            End If
            Fill_Records(SOTCOMH6, "", True, "SELECT * FROM " & SOTCOMH6)

            ASCMAIN1.Progress("-", "SOTINVHS")
            If refreshWorktables Then
                ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTINVHS)
                ASCDATA1.ExecuteSQL("INSERT INTO " & SOTINVHS & " " & querySOTINVHS)
            End If
            Fill_Records(SOTINVHS, "", True, "SELECT * FROM " & SOTINVHS)

            ' Get Commission data in a work table from class
            ASCMAIN1.Progress("-", "SOTCOMMS")

            TAC.SOCCOMMS.GetNYAGCommissions(OPS_YYYYPP, SOTCOMH1, SOTCOMH4, SOTCOMH5, SOTCOMH6, SOTINVHS)
            Fill_Records(SOTCOMMS, String.Empty, True, "Select * from " & TAC.SOCCOMMS.NYAGCommissionsWorktable)

            ' TAC.SOCCOMMS.GetNYAGCommissions(OPS_YYYYPP, SOTCOMH1, SOTCOMH4, SOTCOMH5, SOTCOMH6, SOTINVHS)
            ' Fill_Records(SOTCOMM2, String.Empty, True, "Select * from " & TAC.SOCCOMMS.NYAGCommissionsWorktable)

            ' relac Comm % in .net since Oracle and .net do not obtain the same sum - totals off by pennies.
            ASCMAIN1.Progress("-", "Comm Calc")
            For Each row As DataRow In dst.Tables(SOTCOMMS).Select()
                Dim SREP_COMM_RATE As Decimal = Val(row("SREP_COMM_RATE") & String.Empty)
                row.Item("SREP_COMM_AMT") = Math.Round((SREP_COMM_RATE / 100) * Val(row.Item("ORDR_AMT_SHP") & String.Empty), 2)
            Next
            dst.Tables(SOTCOMMS).AcceptChanges()

            ' There may be missing customers
            ASCMAIN1.Progress("-", "ARTCUST1")
            For Each row As DataRow In ASCDATA1.SelectDistinct(SOTCOMMS, "CUST_CODE").Select
                If dst.Tables("ARTCUST1").Select("CUST_CODE = '" & row.Item("CUST_CODE") & "'").Length = 0 Then
                    Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").NewRow
                    rowARTCUST1.Item("CUST_CODE") = row.Item("CUST_CODE") & String.Empty
                    rowARTCUST1.Item("CUST_NAME") = "Unknown Customer"
                    dst.Tables("ARTCUST1").Rows.Add(rowARTCUST1)
                End If
            Next

            ASCMAIN1.Progress("-", "SOTSREP1")
            For Each row As DataRow In ASCDATA1.SelectDistinct(SOTCOMMS, "SREP_CODE").Select
                If dst.Tables("SOTSREP1").Select("SREP_CODE = '" & row.Item("SREP_CODE") & "'").Length = 0 Then
                    Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").NewRow
                    rowSOTSREP1.Item("SREP_CODE") = row.Item("SREP_CODE") & String.Empty
                    rowSOTSREP1.Item("SREP_NAME") = "Unknown sales rep"
                    dst.Tables("SOTSREP1").Rows.Add(rowSOTSREP1)
                End If
            Next

            ASCMAIN1.Progress("-", "SOTINVH1")
            Fill_Records("SOTINVH1")

            ASCMAIN1.Progress("-", "Totals")
            CalculateTotals()

            ToggleDataTableExpressions(True)
            EnforceConstraints(True)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Load Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

    Private Sub Update_Record()

        ' do we write this row if the old comm amt was 0
        Try
            BeginTrans()

            ' Upadte SOTCOMH1, SOTCOMH4, SOTCOMH5, SOTCOMH6 for the selected period
            ' Update SOTINVHS

            ASCDATA1.ExecuteSQL("DELETE FROM SOTCOMH1 WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "'")
            ASCDATA1.ExecuteSQL("INSERT INTO SOTCOMH1 SELECT OPS_YYYYPP, SREP_CODE, SREP_COMM_RATE FROM " & SOTCOMH1)

            ASCDATA1.ExecuteSQL("DELETE FROM SOTCOMH4 WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "'")
            ASCDATA1.ExecuteSQL("INSERT INTO SOTCOMH4 SELECT OPS_YYYYPP, SREP_CODE, STYLE_GROUP_CODE, SREP_COMM_RATE FROM " & SOTCOMH4)

            ASCDATA1.ExecuteSQL("DELETE FROM SOTCOMH5 WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "'")
            ASCDATA1.ExecuteSQL("INSERT INTO SOTCOMH5 SELECT OPS_YYYYPP, SREP_CODE, CUST_CODE, SREP_COMM_RATE, SREP_COMM_USE_STD FROM " & SOTCOMH5)

            ASCDATA1.ExecuteSQL("DELETE FROM SOTCOMH6 WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "'")
            ASCDATA1.ExecuteSQL("INSERT INTO SOTCOMH6 SELECT OPS_YYYYPP, SREP_CODE, CUST_CODE, STYLE_GROUP_CODE, SREP_COMM_RATE FROM " & SOTCOMH6)

            ASCDATA1.ExecuteSQL("DELETE FROM SOTINVHS WHERE (INV_TYPE, INV_NO) IN (SELECT INV_TYPE, INV_NO FROM " & SOTINVHS & ")")
            ASCDATA1.ExecuteSQL("INSERT INTO SOTINVHS SELECT * FROM " & SOTINVHS)


            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTINVH2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Inventory Status")
        Load_Popup_Menu(grdSOTINVHR, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Inventory Status")
        Load_Popup_Menu(grdSOTINVHC, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Inventory Status")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSOTINVHC"
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTCOMH1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCOMH1.AfterRowUpdate
        ' Update the % for the sales rep for the Customer/Srep relationship
        Dim SREP_COMM_RATE As Decimal = Val(e.Row.Cells("SREP_COMM_RATE").Value & String.Empty)
        Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & String.Empty
        UpdateCommPercentages(TAC.SOCCOMMS.NYACommissionCalcTypes.Customer, SREP_CODE, SREP_COMM_RATE, String.Empty, String.Empty)
    End Sub

    Private Sub grdSOTCOMH4_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCOMH4.AfterCellUpdate

        If e.Cell.Column.Key = "SREP_CODE" Then
            If e.Cell.Row.Cells("SREP_CODE").Value & String.Empty <> String.Empty Then
                Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(e.Cell.Row.Cells("SREP_CODE").Value)
                If rowSOTSREP1 IsNot Nothing Then
                    e.Cell.Row.Cells("SREP_NAME").Value = rowSOTSREP1.Item("SREP_NAME") & String.Empty
                End If
            Else
                e.Cell.Row.Cells("SREP_NAME").Value = String.Empty
            End If
        ElseIf e.Cell.Column.Key = "STYLE_GROUP_CODE" Then
            If e.Cell.Row.Cells("STYLE_GROUP_CODE").Value & String.Empty <> String.Empty Then
                Dim rowICTSGRP1 As DataRow = dst.Tables("ICTSGRP1").Rows.Find(e.Cell.Row.Cells("STYLE_GROUP_CODE").Value)
                If rowICTSGRP1 IsNot Nothing Then
                    e.Cell.Row.Cells("STYLE_GROUP_DESC").Value = rowICTSGRP1.Item("STYLE_GROUP_DESC") & String.Empty
                End If
            Else
                e.Cell.Row.Cells("STYLE_GROUP_DESC").Value = String.Empty
            End If

        End If

    End Sub

    Private Sub grdSOTCOMH4_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCOMH4.AfterRowActivate
        If grdSOTCOMH4.ActiveRow Is Nothing OrElse Not grdSOTCOMH4.ActiveRow.IsAddRow Then
            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("SREP_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("STYLE_GROUP_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("SREP_CODE").Style = UltraWinGrid.ColumnStyle.Default
            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("STYLE_GROUP_CODE").Style = UltraWinGrid.ColumnStyle.Default

        Else
            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("SREP_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("STYLE_GROUP_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit

            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("SREP_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
            grdSOTCOMH4.DisplayLayout.Bands(0).Columns("STYLE_GROUP_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
        End If
    End Sub

    Private Sub grdSOTCOMH4_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCOMH4.BeforeRowUpdate

        ' defualt to the current period
        e.Row.Cells("OPS_YYYYPP").Value = OPS_YYYYPP

        Dim SREP_CODE As String = (e.Row.Cells("SREP_CODE").Value & String.Empty).ToString.Trim.ToUpper
        Dim STYLE_GROUP_CODE As String = (e.Row.Cells("STYLE_GROUP_CODE").Value & String.Empty).ToString.Trim.ToUpper
        Dim SREP_COMM_RATE As Decimal = Val(e.Row.Cells("SREP_COMM_RATE").Value & String.Empty)

        If SREP_CODE.Length = 0 OrElse STYLE_GROUP_CODE.Length = 0 Then
            MessageBox.Show("The Sales Rep Code and the Style Group Code are required.")
            e.Cancel = True
            Exit Sub
        End If

        If SREP_COMM_RATE <= 0 Then
            MessageBox.Show("The Sales Rep commission must be greater equal 0.")
            e.Cancel = True
            Exit Sub
        End If

        If SREP_COMM_RATE >= 6 Then
            If MessageBox.Show("The Sales Rep commission is greater equal 6%. Do you want to save the record?", "Commission %", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                e.Cancel = True
                Exit Sub
            End If
        End If

    End Sub

    Private Sub grdSOTCOMH4_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCOMH4.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SREP_CODE", "STYLE_GROUP_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTCOMH4, sql_where, False)
        End Select
    End Sub

    Private Sub grdSOTCOMH4_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCOMH4.AfterRowUpdate
        ' Update the % for the sales rep for the Customer/Srep relationship
        Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & String.Empty
        Dim STYLE_GROUP_CODE As String = e.Row.Cells("STYLE_GROUP_CODE").Value & String.Empty
        Dim SREP_COMM_RATE As Decimal = Val(e.Row.Cells("SREP_COMM_RATE").Value & String.Empty)
        UpdateCommPercentages(TAC.SOCCOMMS.NYACommissionCalcTypes.StyleGroupOverride, SREP_CODE, SREP_COMM_RATE, STYLE_GROUP_CODE, String.Empty)
    End Sub

    Private Sub grdSOTCOMH5_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCOMH5.AfterCellUpdate

        If e.Cell.Column.Key = "SREP_CODE" Then
            If e.Cell.Row.Cells("SREP_CODE").Value & String.Empty <> String.Empty Then
                Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(e.Cell.Row.Cells("SREP_CODE").Value)
                If rowSOTSREP1 IsNot Nothing Then
                    e.Cell.Row.Cells("SREP_NAME").Value = rowSOTSREP1.Item("SREP_NAME") & String.Empty
                End If
            Else
                e.Cell.Row.Cells("SREP_NAME").Value = String.Empty
            End If
        ElseIf e.Cell.Column.Key = "STYLE_GROUP_CODE" Then
            If e.Cell.Row.Cells("STYLE_GROUP_CODE").Value & String.Empty <> String.Empty Then
                Dim rowICTSGRP1 As DataRow = dst.Tables("ICTSGRP1").Rows.Find(e.Cell.Row.Cells("STYLE_GROUP_CODE").Value)
                If rowICTSGRP1 IsNot Nothing Then
                    e.Cell.Row.Cells("STYLE_GROUP_DESC").Value = rowICTSGRP1.Item("STYLE_GROUP_DESC") & String.Empty
                End If
            Else
                e.Cell.Row.Cells("STYLE_GROUP_DESC").Value = String.Empty
            End If
        ElseIf e.Cell.Column.Key = "CUST_CODE" Then
            If e.Cell.Row.Cells("CUST_CODE").Value & String.Empty <> String.Empty Then
                Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(e.Cell.Row.Cells("CUST_CODE").Value)
                If rowARTCUST1 IsNot Nothing Then
                    e.Cell.Row.Cells("CUST_NAME").Value = rowARTCUST1.Item("CUST_NAME") & String.Empty
                End If
            Else
                e.Cell.Row.Cells("CUST_NAME").Value = String.Empty
            End If

        End If

    End Sub

    Private Sub grdSOTCOMH5_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCOMH5.AfterRowActivate
        If grdSOTCOMH5.ActiveRow Is Nothing OrElse Not grdSOTCOMH5.ActiveRow.IsAddRow Then
            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("SREP_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCOMH5.DisplayLayout.Bands(1).Columns("STYLE_GROUP_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("SREP_CODE").Style = UltraWinGrid.ColumnStyle.Default
            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("CUST_CODE").Style = UltraWinGrid.ColumnStyle.Default
            grdSOTCOMH5.DisplayLayout.Bands(1).Columns("STYLE_GROUP_CODE").Style = UltraWinGrid.ColumnStyle.Default

        Else
            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("SREP_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCOMH5.DisplayLayout.Bands(1).Columns("STYLE_GROUP_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit

            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("SREP_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
            grdSOTCOMH5.DisplayLayout.Bands(0).Columns("CUST_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
            grdSOTCOMH5.DisplayLayout.Bands(1).Columns("STYLE_GROUP_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
        End If
    End Sub

    Private Sub grdSOTCOMH5_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCOMH5.BeforeRowUpdate

        ' defualt to the current period
        e.Row.Cells("OPS_YYYYPP").Value = OPS_YYYYPP
        Dim SREP_COMM_RATE As Decimal = 0

        If e.Row.Band.Index = 0 Then
            Dim SREP_CODE As String = (e.Row.Cells("SREP_CODE").Value & String.Empty).ToString.Trim.ToUpper
            Dim CUST_CODE As String = (e.Row.Cells("CUST_CODE").Value & String.Empty).ToString.Trim.ToUpper
            SREP_COMM_RATE = Val(e.Row.Cells("SREP_COMM_RATE").Value & String.Empty)

            If SREP_CODE.Length = 0 OrElse CUST_CODE.Length = 0 Then
                MessageBox.Show("The Sales Rep Code and the Customer Code are required.")
                e.Cancel = True
                Exit Sub
            End If

            If SREP_COMM_RATE < 0 Then
                MessageBox.Show("The Sales Rep commission must be greater equal 0.")
                e.Cancel = True
                Exit Sub
            End If

        Else
            Dim STYLE_GROUP_CODE As String = (e.Row.Cells("STYLE_GROUP_CODE").Value & String.Empty).ToString.Trim.ToUpper
            SREP_COMM_RATE = Val(e.Row.Cells("SREP_COMM_RATE").Value & String.Empty)

            If STYLE_GROUP_CODE.Length = 0 Then
                MessageBox.Show("The Sales Rep Code and the Style Group Code is required.")
                e.Cancel = True
                Exit Sub
            End If

            If SREP_COMM_RATE <= 0 Then
                MessageBox.Show("The Sales Rep commission must be greater equal 0.")
                e.Cancel = True
                Exit Sub
            End If

        End If

        If SREP_COMM_RATE >= 6 Then
            If MessageBox.Show("The Sales Rep commission is greater equal 6%. Do you want to save the record?", "Commission %", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                e.Cancel = True
                Exit Sub
            End If
        End If

    End Sub

    Private Sub grdSOTCOMH5_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCOMH5.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SREP_CODE", "STYLE_GROUP_CODE", "CUST_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTCOMH5, sql_where, False)
        End Select
    End Sub

    Private Sub grdSOTCOMH5_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCOMH5.AfterRowUpdate

        Select Case e.Row.Band.Key

            Case "grdSOTCOMH5"
                ' Update the % for the sales rep for the Customer/Srep relationship
                Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & String.Empty
                Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & String.Empty
                Dim SREP_COMM_RATE As Decimal = Val(e.Row.Cells("SREP_COMM_RATE").Value & String.Empty)
                Dim SREP_COMM_RATE_M As Decimal = Val(e.Row.Cells("SREP_COMM_RATE_M").Value & String.Empty)
                Dim SREP_COMM_USE_STD As String = e.Row.Cells("SREP_COMM_USE_STD").Value & String.Empty

                If SREP_COMM_USE_STD = "1" Then
                    SREP_COMM_RATE = SREP_COMM_RATE_M
                End If

                UpdateCommPercentages(TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerOverride, SREP_CODE, SREP_COMM_RATE, String.Empty, CUST_CODE)

            Case "grdSOTCOMH5_grdSOTCOMH6"
                ' Update the % for the sales rep for the Customer/Srep relationship
                Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & String.Empty
                Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & String.Empty
                Dim STYLE_GROUP_CODE As String = e.Row.Cells("STYLE_GROUP_CODE").Value & String.Empty
                Dim SREP_COMM_RATE As Decimal = Val(e.Row.Cells("SREP_COMM_RATE").Value & String.Empty)
                UpdateCommPercentages(TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerAndStyleGroupOverride, SREP_CODE, SREP_COMM_RATE, STYLE_GROUP_CODE, CUST_CODE)
        End Select

    End Sub



    Private Sub tabComm_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabComm.SelectedTabChanged

        Select Case tabComm.SelectedTab.Key
            Case "S"
                SETUP_grdSOTINVH2(grdSOTINVHR)
            Case "C"
                SETUP_grdSOTINVH2(grdSOTINVHC)
            Case "R"
                splTab.Panel2Collapsed = True
        End Select

    End Sub

    Private Sub grdSOTINVHR_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTINVHR.AfterRowActivate
        SETUP_grdSOTINVH2(grdSOTINVHR)
    End Sub

    Private Sub grdSOTINVHR_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTINVHR.InitializeRow
        CommissionCalcDescription(e)
    End Sub


    Private Sub grdSOTINVHC_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTINVHC.AfterRowActivate
        SETUP_grdSOTINVH2(grdSOTINVHC)
    End Sub

    Private Sub grdSOTINVHC_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTINVHC.InitializeRow
        CommissionCalcDescription(e)
    End Sub


    Private Sub grdSOTINVH1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTINVH1.BeforeRowUpdate
        Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & String.Empty
        SREP_CODE = SREP_CODE.Trim

        If dst.Tables("SOTSREP1").Rows.Find(New Object() {SREP_CODE}) Is Nothing Then
            e.Cancel = True
            MessageBox.Show("Invalid Sales Rep.")
        End If

    End Sub

    Private Sub grdSOTINVH1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTINVH1.AfterRowUpdate
        'Dim SREP_CODE As String = e.Row.Cells("SREP_CODE").Value & String.Empty
        'Dim INV_NO As String = e.Row.Cells("INV_NO").Value & String.Empty
        'Dim INV_TYPE As String = e.Row.Cells("INV_TYPE").Value & String.Empty

        'For Each rowSOTCOMMS As DataRow In dst.Tables(SOTCOMMS).Select("INV_NO = '" & INV_NO & "' AND INV_TYPE = '" & INV_TYPE & "' AND COMM_CALC_BY = 2")
        '    rowSOTCOMMS.Item("SREP_CODE") = SREP_CODE
        'Next
    End Sub

    Private Sub grdSOTINVH1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTINVH1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SREP_CODE"
                grdClickCellButton(grdSOTINVH1, "")
        End Select
    End Sub

    Private Sub grdSOTINVH1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTINVH1.InitializeRow
        CommissionCalcDescription(e)
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub CommissionCalcDescription(e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)

        If e.Row.Band.Index <> 1 Then Exit Sub

        Dim COMM_CALC_BY As String = e.Row.Cells("COMM_CALC_BY").Value & String.Empty

        Select Case COMM_CALC_BY
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.Customer : e.Row.Cells("COMM_CALC_BY_DESC").Value = "Customer"
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerAndDivision : e.Row.Cells("COMM_CALC_BY_DESC").Value = "Customer/Division"
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerAndStyleGroupOverride : e.Row.Cells("COMM_CALC_BY_DESC").Value = "Customer/Style Override"
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerOverride : e.Row.Cells("COMM_CALC_BY_DESC").Value = "Customer Override"
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.StyleGroupOverride : e.Row.Cells("COMM_CALC_BY_DESC").Value = "Style Override"
            Case Else : e.Row.Cells("COMM_CALC_BY_DESC").Value = "Unknown"
        End Select

    End Sub

    Private Sub SETUP_grdSOTINVH2(ByRef grd As Infragistics.Win.UltraWinGrid.UltraGrid)

        If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
            splTab.Panel2Collapsed = True
        ElseIf Not grd.ActiveRow.Band.Key.Contains("SOTINVHC") Then
            splTab.Panel2Collapsed = True
        Else
            Dim INV_TYPE As String = grd.ActiveRow.Cells("INV_TYPE").Value
            Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value

            If dst.Tables("SOTINVH2").Select("INV_NO = '" & INV_NO & "' AND INV_TYPE = '" & INV_TYPE & "'").Length = 0 Then
                Fill_Records("SOTINVH2", New String() {INV_TYPE, INV_NO})
                Sort_grdColumns(grdSOTINVH2, "INV_LNO")
                grdSOTINVH2.Text = "Sales Invoice Details for Invoice " & INV_TYPE & ":" & INV_NO
            End If
            splTab.Panel2Collapsed = False
        End If

    End Sub

    Private Sub Print_Record()

        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Printing Salesperson Commission Report")

            Initialize_Report("SORCOMMN")
            REPORTS("SORCOMMN").Fill_Records_RPT(OPS_YYYYPP, SOTCOMMS)
            With REPORTS("SORCOMMN")
                .Print_Report_Begin()
                .Print_Report()
                .Print_Report_End()
            End With

        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message, "Print")
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Private Sub UpdateCommPercentages(ByVal COMM_CALC_TYPE As String, ByVal SREP_CODE As String, ByVal SREP_COMM_RATE As Decimal, _
                                      ByVal STYLE_GROUP_CODE As String, ByVal CUST_CODE As String)

        If 1 = 1 Then Exit Sub

        MyBase.EnforceConstraints(False)
        ToggleDataTableExpressions(False)

        ' Dim sqlSearch As String = "ISNULL(MANUAL_CHANGE, '0') = '0' and COMM_CALC_TYPE = '" & COMM_CALC_TYPE & "' and SREP_CODE = '" & SREP_CODE & "'"
        Dim sqlSearch As String = "COMM_CALC_BY = '" & COMM_CALC_TYPE & "' and SREP_CODE = '" & SREP_CODE & "'"

        Select Case COMM_CALC_TYPE
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.Customer
                sqlSearch &= ""
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerAndDivision
                sqlSearch &= " and CUST_CODE = '" & CUST_CODE & "' and SALES_DIVISION_CODE = '" & STYLE_GROUP_CODE & "'"
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.StyleGroupOverride
                sqlSearch &= " and STYLE_GROUP_CODE = '" & STYLE_GROUP_CODE & "'"
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerOverride
                sqlSearch &= " and CUST_CODE = '" & CUST_CODE & "'"
            Case TAC.SOCCOMMS.NYACommissionCalcTypes.CustomerAndStyleGroupOverride
                sqlSearch &= " and CUST_CODE = '" & CUST_CODE & "' and STYLE_GROUP_CODE = '" & STYLE_GROUP_CODE & "'"
        End Select

        For Each row As DataRow In dst.Tables(SOTCOMMS).Select(sqlSearch)
            row.Item("SREP_COMM_RATE") = SREP_COMM_RATE
            row.Item("SREP_COMM_AMT") = Math.Round((SREP_COMM_RATE / 100) * Val(row.Item("ORDR_AMT_SHP") & String.Empty), 2)
        Next

        CalculateTotals()
        ToggleDataTableExpressions(True)
        MyBase.EnforceConstraints(True)
    End Sub

    Private Sub CalculateTotals()

        If 1 = 1 Then Exit Sub

        For Each table As String In New String() {"SOTSREP1", "ARTCUST1", "SOTINVH1"}
            For Each row As DataRow In dst.Tables(table).Select
                row.Item("TOT_COMM") = 0
                If table <> "SOTINVH1" Then row.Item("TOT_SALES") = 0
            Next
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(SOTCOMMS, New String() {"SREP_CODE"}).Rows
            Dim SREP_CODE As String = row.Item("SREP_CODE") & String.Empty
            Dim totComm As Decimal = Val(dst.Tables(SOTCOMMS).Compute("SUM(INV_SALES)", "SREP_CODE = '" & SREP_CODE & "'") & String.Empty)
            Dim totsls As Decimal = Val(dst.Tables(SOTCOMMS).Compute("SUM(COMM_AMT)", "SREP_CODE = '" & SREP_CODE & "'") & String.Empty)

            dst.Tables("SOTSREP1").Select("SREP_CODE = '" & SREP_CODE & "'")(0).Item("TOT_COMM") = totComm
            dst.Tables("SOTSREP1").Select("SREP_CODE = '" & SREP_CODE & "'")(0).Item("TOT_SALES") = totsls
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(SOTCOMMS, New String() {"CUST_CODE"}).Rows
            Dim CUST_CODE As String = row.Item("CUST_CODE") & String.Empty
            Dim totComm As Decimal = Val(dst.Tables(SOTCOMMS).Compute("SUM(COMM_AMT)", "CUST_CODE = '" & CUST_CODE & "'") & String.Empty)
            Dim totsls As Decimal = Val(dst.Tables(SOTCOMMS).Compute("SUM(COMM_AMT)", "CUST_CODE = '" & CUST_CODE & "'") & String.Empty)

            dst.Tables("ARTCUST1").Select("CUST_CODE = '" & CUST_CODE & "'")(0).Item("TOT_COMM") = totComm
            dst.Tables("ARTCUST1").Select("CUST_CODE = '" & CUST_CODE & "'")(0).Item("TOT_SALES") = totsls
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(SOTCOMMS, New String() {"INV_NO"}).Rows
            Dim INV_NO As String = row.Item("INV_NO") & String.Empty
            Dim totComm As Decimal = Val(dst.Tables(SOTCOMMS).Compute("SUM(COMM_AMT)", "INV_NO = '" & INV_NO & "'") & String.Empty)

            dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'")(0).Item("TOT_COMM") = totComm
        Next

        dst.Tables("SOTSREP1").AcceptChanges()
        dst.Tables("ARTCUST1").AcceptChanges()
        dst.Tables("SOTINVH1").AcceptChanges()

    End Sub

    Private Sub ToggleDataTableExpressions(ByVal tf As Boolean)

        'If 1 = 1 Then Exit Sub

        With dst.Tables("SOTSREP1")
            .Columns("TOT_COMM").Expression = IIf(Not tf, "", "SUM(CHILD.SREP_COMM_AMT)")
            '.Columns("TOT_SALES").Expression = IIf(Not tf, "", "SUM(CHILD.INV_SALES)")
        End With

        With dst.Tables("ARTCUST1")
            .Columns("TOT_COMM").Expression = IIf(Not tf, "", "SUM(CHILD.SREP_COMM_AMT)")
            '.Columns("TOT_SALES").Expression = IIf(Not tf, "", "SUM(CHILD.INV_SALES)")
        End With

        With dst.Tables("SOTINVH1")
            .Columns("TOT_COMM").Expression = IIf(Not tf, "", "SUM(CHILD.SREP_COMM_AMT)")
        End With

        If tf Then
            dvSrepCode = New DataView(dst.Tables("SOTSREP1"))
            'dvSrepCode.RowFilter = "ISNULL(TOT_COMM, 0) <> 0 and ISNULL(TOT_SALES, 0) <> 0"
            grdSOTINVHR.DataSource = dvSrepCode
            grdSOTINVHR.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

            dvCustCode = New DataView(dst.Tables("ARTCUST1"))
            'dvCustCode.RowFilter = "ISNULL(TOT_COMM, 0) <> 0 and ISNULL(TOT_SALES, 0) <> 0"
            grdSOTINVHC.DataSource = dvCustCode
            grdSOTINVHC.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)

            grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
            grdSOTINVH1.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
        Else
            grdSOTINVHR.DataSource = dst.Tables("SOTSREP1")
            grdSOTINVHC.DataSource = dst.Tables("ARTCUST1")
            grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
        End If

    End Sub

    Public Overrides Function CustomSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As Double, _
        ByVal grd As UltraWinGrid.UltraGrid) As Double

        Select Case grd.Name
            Case "G" '"grdSOTINVHR", "grdSOTINVHC", "grdSOTINVH1"
                Dim totSales As Decimal = Val(dst.Tables(SOTCOMMS).Compute("SUM(INV_SALES)", "") & String.Empty)
                Dim totComm As Decimal = Val(dst.Tables(SOTCOMMS).Compute("SUM(COMM_AMT)", "") & String.Empty)

                If summarySettings.Key = "TOT_SALES" Then
                    Return totSales
                ElseIf summarySettings.Key = "TOT_COMM" Then
                    Return totComm
                End If

        End Select

        Return CustomValue
    End Function

    Private Function UpdateWorkCommissionTables() As Boolean

        Try
            BeginTrans()

            Dim sql As String = String.Empty

            For Each row As DataRow In dst.Tables(SOTCOMH1).Select("", "", DataViewRowState.ModifiedCurrent)
                sql = "UPDATE " & SOTINVHS & " SET SREP_COMM_RATE = " & Val(row.Item("SREP_COMM_RATE") & String.Empty)
                sql &= " Where SREP_CODE = '" & row.Item("SREP_CODE") & "'"
                ASCDATA1.ExecuteSQL(sql)
            Next

            For Each row As DataRow In dst.Tables(SOTCOMMS).Select("", "", DataViewRowState.ModifiedCurrent)
                sql = "UPDATE " & SOTINVHS & " SET SREP_COMM_RATE = " & Val(row.Item("SREP_COMM_RATE") & String.Empty)
                sql &= " Where INV_NO = '" & row.Item("INV_NO") & "' AND INV_TYPE = '" & row.Item("INV_TYPE") & "'"
                ASCDATA1.ExecuteSQL(sql)
            Next

            Update_Record_TDA(SOTCOMH1)
            Update_Record_TDA(SOTCOMH4)
            Update_Record_TDA(SOTCOMH5)
            Update_Record_TDA(SOTCOMH6)

            CommitTrans()
            Return True

        Catch ex As Exception
            Rollback(ex.Message)
            Return False
        End Try
    End Function

#End Region

End Class