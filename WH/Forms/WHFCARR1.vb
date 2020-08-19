Imports Infragistics.Win.UltraWinGrid

Public Class WHFCARR1

    'CREATE TABLE WHTSHPB1 (
    '    INV_CNTL_NO VARCHAR2(10),
    '    CARRIER_CODE  VARCHAR2(6),
    '    INVOICE_NUMBER VARCHAR2(20),
    '    INIT_OPER VARCHAR2 (8),
    '    INIT_DATE DATE,
    '    MIN_DATE DATE,
    '    MAX_DATE DATE,
    '    PRIMARY KEY (INV_CNTL_NO, CARRIER_CODE, INVOICE_NUMBER));

    'Create table WHTSHPB2(
    '    INV_CNTL_NO VARCHAR2(10),
    '    CARRIER_CODE  VARCHAR2(6),
    '    INVOICE_NUMBER VARCHAR2(20),
    '    INVOICE_LNO NUMBER (5),
    '    TRACKING_NO VARCHAR2(30),
    '    BILLED_CHARGE NUMBER(13,2),
    '    PRIMARY KEY (INV_CNTL_NO, CARRIER_CODE, INVOICE_NUMBER, INVOICE_LNO));

    'CREATE TABLE WHTSHPCI (
    '   SHIP_CNTL_NO VARCHAR2(10),
    '   SHIP_PACKAGE_NO NUMBER(12),
    '   INVOICE_NUMBER VARCHAR2(20), 
    '   INVOICE_LNO NUMBER(5),    
    '   BILLED_CHARGE NUMBER(13,2),
    '   PRIMARY KEY (SHIP_CNTL_NO, SHIP_PACKAGE_NO, INVOICE_NUMBER, INVOICE_LNO));

    Private PeriodStart As String = String.Empty
    Private PeriodEnd As String = String.Empty

    Private fileToImport As String = String.Empty
    Private WHTSHPB2_WK As String = String.Empty

    Private selectedInvoiceNumber As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Dim SOTINVHX As String = "(SELECT NVL(SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_BOL_NO) BILL_OF_LADING_NO, SUM(SOTINVH1.INV_FREIGHT) INV_FREIGHT
                                        FROM SOTINVH1, SOTPICK1, SOTSHIP1
                                        WHERE SOTPICK1.INV_NO = SOTINVH1.INV_NO
                                        AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO
                                        GROUP BY NVL(SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_BOL_NO)) SOTINVHX"

            Dim WHTSHPC2 As String = "(SELECT SHIP_CNTL_NO, SUM(NET_CHARGE) NET_CHARGE FROM WHTSHPC2 GROUP BY SHIP_CNTL_NO) WHTSHPC2"

            ASCMAIN1.sql = $"Select NVL(SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_BOL_NO) BILL_OF_LADING_NO, WHTSHPC1.SHIP_DATE, WHTSHPC1.OPS_YYYYPP
                    , SOTSHIP1.FRT_TERMS, SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.CUST_CODE, ARTCUST1.CUST_NAME, WHTSHPC1.MASTER_TRACKING_NO
                     , NVL(SOTINVHX.INV_FREIGHT, 0) INV_FREIGHT
                     , NVL(WHTSHPC2.NET_CHARGE, 0) NET_CHARGE
                     , SUM(NVL(WHTSHPCI.BILLED_CHARGE, 0)) BILLED_CHARGE
                     , NVL(WHTSHPC2.NET_CHARGE, 0) - SUM(NVL(WHTSHPCI.BILLED_CHARGE, 0)) RATE_VARIANCE
                     , NVL(SOTINVHX.INV_FREIGHT, 0) - SUM(NVL(WHTSHPCI.BILLED_CHARGE, 0)) INVOICE_VARIANCE
                     FROM SOTSHIP1, SOTPICK1, SOTINVH1, WHTSHPC1, {WHTSHPC2}, WHTSHPCI, ARTCUST1, {SOTINVHX} 
                     WHERE SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                     AND WHTSHPC1.CARRIER_CODE = :PARM1
                     AND (SOTINVH1.ORDR_YYYYPP_UPDATED BETWEEN :PARM2 AND :PARM3 
                        or SOTSHIP1.SHIP_BOL_NO IN (SELECT WHTSHPC1.SHIP_BOL_NO FROM WHTSHPC1, WHTSHPCI WHERE WHTSHPC1.SHIP_CNTL_NO = WHTSHPCI.SHIP_CNTL_NO AND WHTSHPCI.INVOICE_NUMBER = :PARM4))
                     AND SOTPICK1.INV_NO = SOTINVH1.INV_NO
                     AND SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE (+) 
                     AND WHTSHPC1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO
                     AND WHTSHPC1.SHIP_CNTL_NO = WHTSHPC2.SHIP_CNTL_NO
                     AND WHTSHPC1.SHIP_CNTL_NO = WHTSHPCI.SHIP_CNTL_NO (+)
                     AND NVL(SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_BOL_NO) = SOTINVHX.BILL_OF_LADING_NO (+)
                     AND NVL(WHTSHPC1.STATUS, 'P') = 'P'
                     GROUP BY NVL(SOTSHIP1.BILL_OF_LADING_NO, SOTSHIP1.SHIP_BOL_NO), WHTSHPC1.SHIP_DATE, 
                        WHTSHPC1.OPS_YYYYPP, SOTSHIP1.FRT_TERMS, SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.CUST_CODE, 
                        ARTCUST1.CUST_NAME, WHTSHPC1.MASTER_TRACKING_NO, NVL(SOTINVHX.INV_FREIGHT, 0), NVL(WHTSHPC2.NET_CHARGE, 0)
                     HAVING NVL(SOTINVHX.INV_FREIGHT, 0) > 0 OR NVL(WHTSHPC2.NET_CHARGE, 0) > 0 OR SUM(NVL(WHTSHPCI.BILLED_CHARGE, 0)) > 0"

            Create_TDA(.Tables.Add, "WHTSHPCX_INV", ASCMAIN1.sql, 0, False, "VVVV", 0)

            ASCMAIN1.sql = "SELECT BILL_OF_LADING_NO, SHIP_DATE, OPS_YYYYPP, CUST_CODE, CUST_NAME, FRT_TERMS, SUM(INV_FREIGHT) INV_FREIGHT, SUM(NET_CHARGE) NET_CHARGE, SUM(BILLED_CHARGE) BILLED_CHARGE, SUM(RATE_VARIANCE) RATE_VARIANCE, SUM(INVOICE_VARIANCE) INVOICE_VARIANCE" _
                & " FROM (" & ASCMAIN1.sql & ") " _
                & " GROUP BY BILL_OF_LADING_NO, SHIP_DATE, OPS_YYYYPP, CUST_CODE, CUST_NAME, FRT_TERMS"

            Create_TDA(.Tables.Add, "WHTSHPCX", ASCMAIN1.sql, 0, False, "VVVV", 0)

            dst.Tables("WHTSHPCX").Columns.Add("VARIANCE", GetType(System.Decimal), "ISNULL(NET_CHARGE, 0) - ISNULL(BILLED_CHARGE, 0)")
            dst.Tables("WHTSHPCX").Columns.Add("SHIP_VARIANCE", GetType(System.Decimal), "ISNULL(INV_FREIGHT, 0) - ISNULL(BILLED_CHARGE, 0)")

            Create_TDA(.Tables.Add("WHTSHPBX"), "WHTSHPB1", "*")
            Create_TDA(.Tables.Add, "WHTSHPB1", "*")
            Create_TDA(.Tables.Add, "WHTSHPB2", "*")
            Create_TDA(.Tables.Add("WHTSHPB2_X"), "WHTSHPB2", "*")

            WHTSHPB2_WK = ASCMAIN1.Temp_Table("SELECT * FROM WHTSHPB2 WHERE ROWNUM < 1")
            ASCMAIN1.sql = "ALTER TABLE " & WHTSHPB2_WK & " ADD PRIMARY KEY (CARRIER_CODE, INVOICE_NUMBER, INVOICE_LNO)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Create_TDA(.Tables.Add, WHTSHPB2_WK, "*")

            Create_Relation("WHTSHPCX", "WHTSHPCX_INV", "BILL_OF_LADING_NO", "BILL_OF_LADING_NO")
            Create_Relation("WHTSHPCX_INV", "WHTSHPB2_X", "MASTER_TRACKING_NO", "TRACKING_NO")

        End With

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -24, 0, 0)

        grdWHTSHPCX.DataSource = dst.Tables("WHTSHPCX")
        grdWHTSHPB2.DataSource = dst.Tables("WHTSHPB2")
        grdWHTSHPB1.DataSource = dst.Tables("WHTSHPBX")

        Create_Summary(grdWHTSHPCX, "BILL_OF_LADING_NO", "Count")
        Create_Summary(grdWHTSHPCX, "INV_FREIGHT", "Sum")
        Create_Summary(grdWHTSHPCX, "NET_CHARGE", "Sum")
        Create_Summary(grdWHTSHPCX, "BILLED_CHARGE", "Sum")
        Create_Summary(grdWHTSHPCX, "RATE_VARIANCE", "Sum")
        Create_Summary(grdWHTSHPCX, "INVOICE_VARIANCE", "Sum")

        Create_Summary(grdWHTSHPB1, "CARRIER_CODE", "Count")

        Create_Summary(grdWHTSHPB2, "INVOICE_NUMBER", "Count")
        Create_Summary(grdWHTSHPB2, "BILLED_CHARGE", "SUM")

        grdWHTSHPB2.Dock = DockStyle.Fill
        grdWHTSHPCX.Dock = DockStyle.Fill
        grdWHTSHPB1.Dock = DockStyle.Fill

        grdWHTSHPCX.DisplayLayout.Override.HeaderClickAction = HeaderClickAction.SortMulti
        grdWHTSHPB1.DisplayLayout.Override.HeaderClickAction = HeaderClickAction.SortMulti
        grdWHTSHPB2.DisplayLayout.Override.HeaderClickAction = HeaderClickAction.SortMulti

        With grdWHTSHPCX.DisplayLayout.Bands(0)
            .Columns("RATE_VARIANCE").CellAppearance.BackColor = Drawing.Color.LightBlue
            .Columns("INVOICE_VARIANCE").CellAppearance.BackColor = Drawing.Color.LightGreen
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Import"
                Validate_Code("CARRIER_CODE")
                If EMsg.Length > 0 Then Exit Select

                fileToImport = String.Empty

                Select Case Absx1.txtFor("CARRIER_CODE").Text
                    Case "UPS", "FEDEX"
                        ' Valid Carrier
                    Case Else
                        EMsg &= vbCr & "The selected carrier is not supported on this screen."
                        Exit Select
                End Select

                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Open Billing File"

                    Select Case Absx1.txtFor("CARRIER_CODE").Text
                        Case "UPS", "FEDEX"
                            openFileDialog1.Filter = "Excel files (*.xls)|*.xls"

                        Case Else
                            ' Should never fire
                    End Select

                    openFileDialog1.FilterIndex = 1
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        fileToImport = openFileDialog1.FileName
                    End If

                    openFileDialog1.Dispose()
                End Using

                If fileToImport.Length = 0 Then
                    EMsg &= vbCr & "No file selected."
                    Exit Select
                End If

                If Not My.Computer.FileSystem.FileExists(fileToImport) Then
                    EMsg &= vbCr & "Unable to locate file " & fileToImport
                    Exit Select
                End If

            Case "Load"
                Validate_Code("CARRIER_CODE")
                If EMsg.Length > 0 Then Exit Select

                PeriodStart = cmbPeriodStart.SelectedRow.Cells("OPS_YYYYPP").Value & String.Empty
                If PeriodStart.Length = 0 Then
                    EMsg &= vbCr & "You must select a start period."
                    Exit Select
                End If

                PeriodEnd = cmbPeriodEnd.SelectedRow.Cells("OPS_YYYYPP").Value & String.Empty
                If PeriodEnd.Length = 0 Then
                    EMsg &= vbCr & "You must select an end period."
                    Exit Select
                End If

                If Val(PeriodStart) > Val(PeriodEnd) Then
                    EMsg &= vbCr & "Start period must be less/equal the End period."
                    Exit Select
                End If

            Case "Update"

                If dst.Tables("WHTSHPB1").Rows.Count = 0 Then
                    EMsg &= "The imported file did not create a header record"
                    Exit Select
                End If

                If dst.Tables("WHTSHPB2").Rows.Count = 0 Then
                    EMsg &= "The imported file did not have details"
                    Exit Select
                End If

                ASCMAIN1.sql = "Select * from WHTSHPB1 where (CARRIER_CODE, INVOICE_NUMBER) IN (SELECT CARRIER_CODE, INVOICE_NUMBER FROM " & WHTSHPB2_WK & ")"
                Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                If tbl.Rows.Count > 0 Then
                    Dim zMsg As String = "The following Invoices were previously imported. If you continue with the Update all previous imported data will be deleted and replaced with the current data." _
                                         & Environment.NewLine & Environment.NewLine

                    For Each row As DataRow In tbl.Select("", "INVOICE_NUMBER")
                        zMsg &= row.Item("INVOICE_NUMBER") & Environment.NewLine
                    Next

                    zMsg &= Environment.NewLine & "Do you want to Continue with the Update?"

                    If MessageBox.Show(zMsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If


                If MessageBox.Show("Do you want to Update the Imported Charges?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Cancel"
                If EntryMode = "N" Then
                    If MessageBox.Show("Do you want to Cancel the Imported Charges?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Import"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Import").Settings.Enabled = not_iScreenMode
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    If Not (ScreenMode AndAlso EntryMode = "N") Then
                        .Items("Update").Settings.Enabled = DefaultableBoolean.False
                    End If
                End With

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdWHTSHPB2.Visible = ScreenMode AndAlso EntryMode = "N"
        grdWHTSHPCX.Visible = ScreenMode AndAlso EntryMode = "L"

        grdWHTSHPB1.Visible = Not ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"WHTSHPCX", "WHTSHPB1", "WHTSHPB2", WHTSHPB2_WK, "WHTSHPBX", "WHTSHPCX_INV", "WHTSHPB2_X"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Fill_Records("WHTSHPBX", String.Empty, True, "SELECT * FROM WHTSHPB1 WHERE INIT_DATE >= SYSDATE - 180")
        Sort_grdColumns(grdWHTSHPB1, "CARRIER_CODE,INIT_DATE")
        grdWHTSHPB1.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)

        EnforceConstraints(True)

        PeriodStart = String.Empty
        PeriodEnd = String.Empty
        fileToImport = String.Empty
        selectedInvoiceNumber = String.Empty

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Select Case EntryMode

            Case "L"

                If selectedInvoiceNumber.Length = 0 Then
                    ' Prevents Carrier Invoice Number to work
                    ASCMAIN1.Progress("-", "WHTSHPCX")
                    Fill_Records("WHTSHPCX", New Object() {HFs("CARRIER_CODE"), PeriodStart, PeriodEnd, "DONOTSELECT"})
                    grdWHTSHPCX.Text = "Period range: " & PeriodStart & " to " & PeriodEnd

                    ASCMAIN1.Progress("-", "WHTSHPCX_INV")
                    Fill_Records("WHTSHPCX_INV", New Object() {HFs("CARRIER_CODE"), PeriodStart, PeriodEnd, "DONOTSELECT"})
                Else
                    ' Prevents Period range to work
                    ASCMAIN1.Progress("-", "WHTSHPCX")
                    Fill_Records("WHTSHPCX", New Object() {HFs("CARRIER_CODE"), "202002", "202001", selectedInvoiceNumber})
                    grdWHTSHPCX.Text = "Invoice Number: " & selectedInvoiceNumber

                    ASCMAIN1.Progress("-", "WHTSHPCX_INV")
                    Fill_Records("WHTSHPCX_INV", New Object() {HFs("CARRIER_CODE"), "202002", "202001", selectedInvoiceNumber})
                End If
                grdWHTSHPCX.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)

                ASCMAIN1.Progress("-", "WHTSHPB2_X")
                Dim lst As New List(Of String)
                For Each row As DataRow In dst.Tables("WHTSHPCX_INV").Select("")
                    lst.Add(row.Item("MASTER_TRACKING_NO"))

                    If lst.Count >= 25 Then
                        ASCMAIN1.sql = $"SELECT * FROM WHTSHPB2 WHERE TRACKING_NO IN ('{String.Join("', '", lst.ToArray)}')"
                        Fill_Records("WHTSHPB2_X", String.Empty, False, ASCMAIN1.sql)
                        lst.Clear()
                    End If
                Next

                If lst.Count > 0 Then
                    ASCMAIN1.sql = $"SELECT * FROM WHTSHPB2 WHERE TRACKING_NO IN ('{String.Join("', '", lst.ToArray)}')"
                    Fill_Records("WHTSHPB2_X", String.Empty, False, ASCMAIN1.sql)
                    lst.Clear()
                End If

            Case "N"

                Dim tableData As New DataTable

                Try
                    Select Case HFs("CARRIER_CODE")
                        Case "UPS", "FEDEX"

                            Using cn As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileToImport & ";Extended Properties=""Excel 12.0;HDR=YES;IMEX=1""")
                                Using cmd As New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", cn)
                                    ' Select the data from Sheet1 of the workbook.
                                    cn.Open()
                                    cmd.Fill(tableData)
                                    cn.Close()
                                    cmd.Dispose()
                                End Using
                                cn.Dispose()
                            End Using

                            Dim INVOICE_LNO As Int32 = 1

                            For Each COL As DataColumn In tableData.Columns
                                tableData.Columns(COL.ColumnName).ColumnName = COL.ColumnName.Replace(" ", "_")
                            Next

                            ' Rename the FedEx columns to match the UPS Columns
                            If HFs("CARRIER_CODE") = "FEDEX" Then
                                ' Invoice Number
                                ' Express or Ground Tracking ID
                                ' Net Charge Amount
                                ' Shipment Date

                                tableData.Columns("Express_or_Ground_Tracking_ID").ColumnName = "TRACKING_NUMBER"
                                tableData.Columns("Net_Charge_Amount").ColumnName = "BILLED_CHARGE"
                                tableData.Columns.Add("PICKUP_DATE", GetType(System.DateTime))

                                Dim dateFormat As String = "yyyyMMdd"
                                Dim provider As Globalization.CultureInfo = Globalization.CultureInfo.InvariantCulture
                                For Each row As DataRow In tableData.Select("")
                                    Dim PICKUP_DATE As String = row.Item("Shipment_Date") & String.Empty
                                    If PICKUP_DATE.Length = 8 Then
                                        row.Item("PICKUP_DATE") = Date.ParseExact(PICKUP_DATE, dateFormat, provider)
                                    End If
                                Next

                            End If

                            tableData.Columns("PICKUP_DATE").DataType = GetType(System.DateTime)

                            Dim INVOICE_NUMBER As String = String.Empty
                            Dim INV_CNTL_NO As String = ASCMAIN1.Next_Control_No("WHTSHPB1.INV_CNTL_NO")

                            For Each row As DataRow In tableData.Select("", "INVOICE_NUMBER")

                                If (row.Item("TRACKING_NUMBER") & String.Empty).ToString.Length = 0 Then
                                    Continue For
                                End If

                                If INVOICE_NUMBER <> row.Item("INVOICE_NUMBER") & String.Empty Then
                                    INVOICE_NUMBER = row.Item("INVOICE_NUMBER") & String.Empty
                                    INVOICE_LNO = 1
                                End If

                                ASCMAIN1.Progress("-", $"{INVOICE_NUMBER}/{INVOICE_LNO}")

                                Dim rowWHTSHPB2 As DataRow = dst.Tables(WHTSHPB2_WK).NewRow
                                rowWHTSHPB2.Item("INV_CNTL_NO") = INV_CNTL_NO
                                rowWHTSHPB2.Item("CARRIER_CODE") = HFs("CARRIER_CODE")
                                rowWHTSHPB2.Item("INVOICE_NUMBER") = row.Item("INVOICE_NUMBER")
                                rowWHTSHPB2.Item("INVOICE_LNO") = INVOICE_LNO
                                INVOICE_LNO += 1
                                rowWHTSHPB2.Item("TRACKING_NO") = row.Item("TRACKING_NUMBER")
                                rowWHTSHPB2.Item("BILLED_CHARGE") = row.Item("BILLED_CHARGE")
                                dst.Tables(WHTSHPB2_WK).Rows.Add(rowWHTSHPB2)
                            Next

                            ASCDATA1.ExecuteSQL("DELETE FROM " & WHTSHPB2_WK)
                            ' This may be a large file
                            'Update_Record_TDA(WHTSHPB2_WK)
                            dst.Tables(WHTSHPB2_WK).AcceptChanges()
                            For Each row As DataRow In dst.Tables(WHTSHPB2_WK).Select("")
                                row.SetAdded()
                            Next

                            Create_BAs(WHTSHPB2_WK, True)
                            Update_BAs(WHTSHPB2_WK, True)

                            ASCMAIN1.Progress("-", "WHTSHPB2")
                            Fill_Records("WHTSHPB2", String.Empty, True, "SELECT * FROM " & WHTSHPB2_WK)

                            Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("WHTSHPB2"), New String() {"INVOICE_NUMBER"})

                            ASCMAIN1.Progress("-", "WHTSHPB1")
                            For Each row As DataRow In TBL.Select("")

                                Dim rowWHTSHPB1 As DataRow = dst.Tables("WHTSHPB1").NewRow
                                rowWHTSHPB1.Item("INV_CNTL_NO") = INV_CNTL_NO
                                rowWHTSHPB1.Item("CARRIER_CODE") = HFs("CARRIER_CODE")
                                rowWHTSHPB1.Item("INVOICE_NUMBER") = row.Item("INVOICE_NUMBER")
                                rowWHTSHPB1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                rowWHTSHPB1.Item("INIT_DATE") = DateTime.Now

                                Dim mDate As String = tableData.Compute("MIN(PICKUP_DATE)", "INVOICE_NUMBER = '" & row.Item("INVOICE_NUMBER") & "'")
                                If IsDate(mDate) Then
                                    rowWHTSHPB1.Item("MIN_DATE") = mDate
                                End If

                                mDate = tableData.Compute("Max(PICKUP_DATE)", "INVOICE_NUMBER = '" & row.Item("INVOICE_NUMBER") & "'")
                                If IsDate(mDate) Then
                                    rowWHTSHPB1.Item("MAX_DATE") = mDate
                                End If

                                dst.Tables("WHTSHPB1").Rows.Add(rowWHTSHPB1)
                            Next

                        Case Else

                    End Select

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    Exit Sub
                End Try

                ASCMAIN1.Progress("-", "Grids")
                grdWHTSHPB2.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
                grdWHTSHPCX.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)

        End Select

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()

            For Each tableName As String In New String() {"WHTSHPB1", "WHTSHPB2"}
                dst.Tables(tableName).AcceptChanges()
                For Each row As DataRow In dst.Tables(tableName).Select("")
                    row.SetAdded()
                Next

                Create_BAs(tableName)
                Update_BAs(tableName)
            Next

            ' Done incase the same invoice is imported more than once
            ASCMAIN1.sql = "DELETE FROM WHTSHPCI WHERE (SHIP_CNTL_NO, SHIP_PACKAGE_NO, INVOICE_NUMBER) IN " _
                    & " (SELECT WHTSHPC2.SHIP_CNTL_NO, WHTSHPC2.SHIP_PACKAGE_NO, WHTSHPB2.INVOICE_NUMBER " _
                    & " FROM WHTSHPC1, WHTSHPC2, " & WHTSHPB2_WK & " WHTSHPB2" _
                    & " WHERE WHTSHPC1.SHIP_CNTL_NO = WHTSHPC2.SHIP_CNTL_NO" _
                    & " AND WHTSHPC1.CARRIER_CODE = WHTSHPB2.CARRIER_CODE" _
                    & " AND WHTSHPC2.TRACKING_NO = WHTSHPB2.TRACKING_NO)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "INSERT INTO WHTSHPCI" _
                    & " SELECT WHTSHPC2.SHIP_CNTL_NO, WHTSHPC2.SHIP_PACKAGE_NO," _
                    & " WHTSHPB2.INVOICE_NUMBER, WHTSHPB2.INVOICE_LNO, WHTSHPB2.BILLED_CHARGE" _
                    & " FROM WHTSHPC1, WHTSHPC2, " & WHTSHPB2_WK & " WHTSHPB2" _
                    & " WHERE WHTSHPC1.SHIP_CNTL_NO = WHTSHPC2.SHIP_CNTL_NO" _
                    & " AND WHTSHPC1.CARRIER_CODE = WHTSHPB2.CARRIER_CODE" _
                    & " AND WHTSHPC2.TRACKING_NO = WHTSHPB2.TRACKING_NO"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub


#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTSHPCX, "SS", "Show Filter", "Show GroupBox")

        Load_Popup_Menu(grdWHTSHPB1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdWHTSHPB2, "SS", "Show Filter", "Show GroupBox")
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
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTPACKX"
                Case "grdSOTORDRX"

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

            Case "Sales Order Inquiry"

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As Control, COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "CARRIER_CODE"
                sql_where &= " CARRIER_TYPE = 'U' AND CARRIER_CODE IN (SELECT CARRIER_CODE FROM SOTCARR3 WHERE CARRIER_ACCOUNT_NO IS NOT NULL AND DIVISION_CODE = '" & ASCMAIN1.CLIENT & "')"
        End Select
    End Sub

    Private Sub grdWHTSHPB1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdWHTSHPB1.DoubleClickRow

        If e.Row.IsDataRow Then
            selectedInvoiceNumber = e.Row.Cells("INVOICE_NUMBER").Value & String.Empty
            Absx1.txtFor("CARRIER_CODE").Text = e.Row.Cells("CARRIER_CODE").Value & String.Empty
            Click_Command("Load")
        End If

    End Sub

#End Region

End Class