Imports System.IO
Imports nsoftware.InShip

Public Class WHFBULK1

#Region "Form variables"

    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo
    Private SO_PARM_DEF_PICK_WHSE As String
    Private tblTATSTATE As DataTable
    Private lstLabelsToProcess As New List(Of String)

    Private sqlSOTBULK267 As String = String.Empty
    Private aswSOTBULK267 As String = String.Empty
    Private tblItemQty As String = String.Empty

    Private Enum ImportTypes
        FromFile
        FromReleasedSalesOrders
        FromOpenSalesOrders
    End Enum

    Private ImportType As ImportTypes

    Private Enum CartonizationMethods
        OneCartonForAllItems
        OneCartonPerItem
        UseCartonInnerDefinitions
    End Enum

    Private CartonizationMethod As CartonizationMethods
    Private successfulUpdate As Boolean = False

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            TABLE_NAME = "SOTBULK1"

            ASCMAIN1.sql = "SELECT ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND," _
                & "  ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO, ICTSTAT2.WHSE_QTY_ON_ORDER, " _
                & "  (ICTSTAT2.WHSE_QTY_ON_HAND - ICTSTAT2.WHSE_QTY_PICK) QTY_AVAIL" _
                & "  FROM ICTSTAT2, SOTBULKI" _
                & "  where rownum < 1"
            tblItemQty = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            Get_PARM("SOTPARM1")
            SO_PARM_DEF_PICK_WHSE = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & String.Empty

            Create_TDA(.Tables.Add, "SOTBULK1", "*")
            Create_TDA(.Tables.Add, "SOTBULK2", "*", 1)
            Create_TDA(.Tables.Add, "SOTBULK3", "*", 1)
            Create_TDA(.Tables.Add, "SOTBULK4", "*", 1)
            Create_TDA(.Tables.Add, "SOTBULK5", "*", 1)
            Create_TDA(.Tables.Add, "SOTBULK6", "*", 1)
            Create_TDA(.Tables.Add, "SOTBULK7", "*", 1)
            Create_TDA(.Tables.Add, "SOTBULK8", "*", 1)
            Create_TDA(.Tables.Add, "SOTBULK9", "*", 1)
            .Tables("SOTBULK9").Columns.Add("SHIP_VIA_DESC", GetType(System.String), "SHIP_VIA_CODE")

            Create_TDA(.Tables.Add, "SOTCARRZ", "*", 2)

            sqlSOTBULK267 = "select sotbulk2.BULK_CODE, sotbulk2.cust_addr_code,"
            sqlSOTBULK267 &= " sotbulk6.PKG_L, sotbulk6.PKG_w, sotbulk6.PKG_H, SOTBULK6.PKG_WEIGHT, sotbulk7.QTY_PACKED || ' PC ' || sotbulk7.STYLE_CODE || ' - ' || sotbulk7.COLOR_CODE DESCRIPT "
            sqlSOTBULK267 &= " , sotbulk7.QTY_PACKED "
            sqlSOTBULK267 &= " from sotbulk2, sotbulk6, sotbulk7 "
            sqlSOTBULK267 &= " where sotbulk2.bulk_code = :PARM1 "
            sqlSOTBULK267 &= " and sotbulk2.bulk_code = sotbulk6.bulk_code "
            sqlSOTBULK267 &= " and sotbulk2.bulk_pattern_no = sotbulk6.bulk_pattern_no "
            sqlSOTBULK267 &= " and sotbulk6.bulk_code = sotbulk7.bulk_code "
            sqlSOTBULK267 &= " and sotbulk6.bulk_pattern_no = sotbulk7.bulk_pattern_no "
            sqlSOTBULK267 &= " and sotbulk6.cart_no = sotbulk7.cart_no "
            'sqlSOTBULK267 &= " GROUP BY"
            'sqlSOTBULK267 &= " sotbulk2.BULK_CODE, sotbulk2.cust_addr_code,"
            'sqlSOTBULK267 &= " sotbulk6.PKG_L, sotbulk6.PKG_w, sotbulk6.PKG_H, SOTBULK6.PKG_WEIGHT, sotbulk7.QTY_PACKED || ' PC ' || sotbulk7.STYLE_CODE || ' - ' || sotbulk7.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTBULK267", sqlSOTBULK267, 0, False, "V", 0)

            Create_TDA(.Tables.Add, "SOTBULKI", "*", 1)
            .Tables("SOTBULKI").Columns.Add("STYLE_DESC", GetType(System.String))
            .Tables("SOTBULKI").Columns.Add("WHSE_QTY_ON_HAND", GetType(System.Int64))
            .Tables("SOTBULKI").Columns.Add("WHSE_QTY_PICK", GetType(System.Int64))
            .Tables("SOTBULKI").Columns.Add("QTY_AVAIL", GetType(System.Int64))
            .Tables("SOTBULKI").Columns.Add("WHSE_QTY_ON_ORDER", GetType(System.Int64))

            ASCMAIN1.sql = "Select ICTSTYC1.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.SUB_UNIT_BAG_QTY, ICTSTYL1.SUB_UNIT_PACK_QTY, ICTSTYL1.STYLE_PRICE" _
                & " FROM ICTSTYL1, ICTSTYC1" _
                & " WHERE ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" _
                & " AND (ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE) IN " _
                & " (SELECT SOTBULK3.STYLE_CODE, SOTBULK3.COLOR_CODE FROM SOTBULK3 WHERE BULK_CODE = :PARM1)"
            Create_TDA(.Tables.Add, "ICTSTYC1", ASCMAIN1.sql, 0, False, "V", 2)

            Create_TDA(.Tables.Add, "ICTSTYLD", "*", 1)
            Create_TDA(.Tables.Add, "ICTSTYL1", "*")

            Create_Relation("SOTBULK4", "SOTBULK2", "BULK_CODE,BULK_PATTERN_NO", "BULK_CODE,BULK_PATTERN_NO")
            Create_Relation("SOTBULK2", "SOTBULK3", "BULK_CODE,CUST_ADDR_CODE", "BULK_CODE,CUST_ADDR_CODE")
            Create_Relation("SOTBULKI", "SOTBULK3", "STYLE_CODE,COLOR_CODE", "STYLE_CODE,COLOR_CODE")

            Create_Relation("SOTBULK2", "SOTBULK267", "BULK_CODE,CUST_ADDR_CODE", "BULK_CODE,CUST_ADDR_CODE")
            .Tables("SOTBULK2").Columns.Add("QTY_PACKED", GetType(System.Int32), "SUM(CHILD(SOTBULK2_SOTBULK267).QTY_PACKED)")

            Create_Relation("SOTBULK4", "SOTBULK6", "BULK_CODE,BULK_PATTERN_NO", "BULK_CODE,BULK_PATTERN_NO")
            Create_Relation("SOTBULK6", "SOTBULK7", "BULK_CODE,BULK_PATTERN_NO,CART_NO", "BULK_CODE,BULK_PATTERN_NO,CART_NO")
            .Tables("SOTBULK6").Columns.Add("NUM_ITEMS", GetType(System.Int16), "COUNT(CHILD.CART_LNO)")

            .Tables("SOTBULK4").Columns.Add("NUM_ACCOUNTS", GetType(System.Int16), "COUNT(CHILD(SOTBULK4_SOTBULK2).CUST_ADDR_CODE)")
            .Tables("SOTBULK4").Columns.Add("NUM_CARTONS", GetType(System.Int16), "COUNT(CHILD(SOTBULK4_SOTBULK6).CART_NO)")
            .Tables("SOTBULKI").Columns.Add("TOTAL_QTY_ORDERED", GetType(System.Int16), "SUM(CHILD.ORDR_QTY)")

            .Tables("SOTBULKI").Columns.Add("SELECTED", GetType(System.String))

            Create_TDA(.Tables.Add, "WHTPKGM1", "*")
            Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")

            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            Fill_Records("SOTSVIA1", String.Empty, True, "SELECT * FROM SOTSVIA1")

            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Fill_Records("SOTCARR1", String.Empty, True, "SELECT * FROM SOTCARR1")

            Create_TDA(.Tables.Add, "SOTCARR2", "*")
            Fill_Records("SOTCARR2", String.Empty, True, "SELECT * FROM SOTCARR2")

            Create_TDA(.Tables.Add, "SOTCARR3", "*")
            Fill_Records("SOTCARR3", String.Empty, True, "SELECT * FROM SOTCARR3")

            Create_TDA(.Tables.Add, "ICTWHSE1", "*")
            Fill_Records("ICTWHSE1", String.Empty, True, "SELECT * FROM ICTWHSE1")

            ' Used for EDI Orders. The SOTBULK* data needs to be  
            Create_TDA(.Tables.Add, "SOTORDR0", "*")
            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR5", "*")

            Create_TDA(.Tables.Add, "SOTCART1", "*")
            Create_TDA(.Tables.Add, "SOTCART2", "*")

            tblTATSTATE = ASCDATA1.GetDataTable("SELECT * FROM TATSTATE", "TATSTATE") ' WHERE region_code is not null

            .Tables.Add("UPSINTL1")
            With .Tables("UPSINTL1")
                .Columns.Add("BULK_CODE", GetType(System.String))
                .Columns.Add("CUST_ADDR_CODE", GetType(System.String))
                .Columns.Add("UNITS", GetType(System.Int32))
                .Columns.Add("UOM", GetType(System.String))
                .Columns.Add("DESC", GetType(System.String))
                .Columns.Add("ORIGIN_COUNTRY", GetType(System.String))
                .Columns.Add("UNIT_VALUE", GetType(System.Decimal))
                .Columns.Add("TOTAL_VALUE", GetType(System.Decimal))
            End With

            .Tables.Add("UPSINTL2")
            With .Tables("UPSINTL2")
                .Columns.Add("BULK_CODE", GetType(System.String))
                .Columns.Add("CUST_ADDR_CODE", GetType(System.String))

                .Columns.Add("INVOICE_LINE_TOTAL", GetType(System.Decimal))
                .Columns.Add("DISCOUNT", GetType(System.Decimal))
                .Columns.Add("INVOICE_SUBTOTAL", GetType(System.Decimal))
                .Columns.Add("FREIGHT", GetType(System.Decimal))
                .Columns.Add("INSURANCE", GetType(System.Decimal))
                .Columns.Add("OTHER", GetType(System.Decimal))
                .Columns.Add("TOTAL_INVOICE_AMOUNT", GetType(System.Decimal))
                .Columns.Add("NUM_CARTONS", GetType(System.Int32))
                .Columns.Add("CURRENCY", GetType(System.String))
                .Columns.Add("TOTAL_WEIGHT", GetType(System.Decimal))
            End With

            .Tables.Add("UPSINTL5")
            With .Tables("UPSINTL5")
                .Columns.Add("BULK_CODE", GetType(System.String))
                .Columns.Add("CUST_ADDR_CODE", GetType(System.String))
                .Columns.Add("CUST_CONTACT", GetType(System.String))
                .Columns.Add("CUST_NAME", GetType(System.String))
                .Columns.Add("CUST_ADDR1", GetType(System.String))
                .Columns.Add("CUST_ADDR2", GetType(System.String))
                .Columns.Add("CUST_CITY", GetType(System.String))
                .Columns.Add("CUST_STATE", GetType(System.String))
                .Columns.Add("CUST_ZIP_CODE", GetType(System.String))
                .Columns.Add("CUST_PHONE", GetType(System.String))
            End With
        End With

        grdSOTBULK2.DataSource = dst.Tables("SOTBULK2")
        grdSOTBULK4.DataSource = dst.Tables("SOTBULK4")
        grdSOTBULK6.DataSource = dst.Tables("SOTBULK4")
        grdSOTBULK9.DataSource = dst.Tables("SOTBULK9")
        grdSOTBULKI.DataSource = dst.Tables("SOTBULKI")
        grdSOTBULK267.DataSource = dst.Tables("SOTBULK2")

        grdSOTBULK2_SHIP.DataSource = dst.Tables("SOTBULK2")
        grdSOTBULKI_SHIP.DataSource = dst.Tables("SOTBULKI")

        ASCMAIN1.Add_Value_List(grdSOTBULK9, "CARRIER_PAYOR", Nothing, New String() {":", "S:Sender", "R:Recipient", "T:Third Party"})
        ASCMAIN1.Add_Value_List(grdSOTBULK9, "CARRIER_DI", Nothing, New String() {":", "D:Domestic", "I:Internaltional", "B:Both"})
        ASCMAIN1.Add_Value_List(grdSOTBULK9, "SHIP_VIA_DESC", "SELECT SHIP_VIA_CODE, SHIP_VIA_DESC FROM SOTSVIA1")

        ASCMAIN1.sql = "Select Distinct SOTCARR1.CARRIER_CODE, SOTCARR1.CARRIER_DESC " _
                & " from SOTCARR1, SOTCARR2" _
                & " where SOTCARR1.CARRIER_CODE = SOTCARR2.CARRIER_CODE" _
                & " AND SOTCARR1.CARRIER_TYPE = 'U'" _
                & " and SOTCARR1.CARRIER_REMOTE_HOST_IP is not null"
        ASCMAIN1.Add_Value_List(grdSOTBULK9, "CARRIER_CODE", ASCMAIN1.sql)

        Create_Summary(grdSOTBULK2, "CUST_ADDR_CODE", "Count")
        Create_Summary(grdSOTBULKI, "STYLE_CODE", "Count")
        Create_Summary(grdSOTBULK4, "BULK_PATTERN_DESC", "Count")
        Create_Summary(grdSOTBULK4, "NUM_ACCOUNTS", "Sum")
        Create_Summary(grdSOTBULK6, "BULK_PATTERN_DESC", "Count")
        Create_Summary(grdSOTBULK6, "NUM_ACCOUNTS", "Sum")

        Create_Summary(grdSOTBULK2_SHIP, "CUST_ADDR_CODE", "Count")
        Create_Summary(grdSOTBULKI_SHIP, "STYLE_CODE", "Count")

        Create_Summary(grdSOTBULK267, "CUST_ADDR_CODE", "Count")
        Create_Summary(grdSOTBULK267, "PKG_L", "Count", "SOTBULK2_SOTBULK267")


        With ultraComboPackage.DisplayLayout.Bands(0)

            ultraComboPackage.Font = grdSOTBULK6.Font
            ultraComboPackage.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Default
            ultraComboPackage.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList

            .Columns.Add("PKG_CODE")
            .Columns("PKG_CODE").Header.Caption = "Code"
            .Columns("PKG_CODE").Width = 75

            .Columns.Add("PKG_DESC")
            .Columns("PKG_DESC").Header.Caption = "Desc"
            .Columns("PKG_DESC").Width = 75

            .Columns.Add("PKG_D")
            .Columns("PKG_D").Header.Caption = "L x W x H"
            .Columns("PKG_D").Width = 200

        End With

        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1 order by PKG_CODE")
        ultraComboPackage.ValueMember = "PKG_CODE"
        ultraComboPackage.DisplayMember = "PKG_DESC"
        grdSOTBULK6.DisplayLayout.Bands("SOTBULK4_SOTBULK6").Columns("PKG_CODE").EditorComponent = ultraComboPackage

        Dim band As Int16 = 0
        For iCtr As Int16 = 0 To grdSOTBULK6.DisplayLayout.Bands.Count - 1
            If grdSOTBULK6.DisplayLayout.Bands(iCtr).Key = "SOTBULK4_SOTBULK6" Then
                band = iCtr
                Exit For
            End If
        Next

        ASCMAIN1.Add_Value_List(grdSOTBULK6, "PACKAGING_TYPE", Nothing, Nothing, band, "SELECT SOTCARR4.PACKAGE_CODE, SOTCARR4.PACKAGE_DESC" _
                & " FROM SOTSVIA1, SOTCARR4" _
                & " WHERE SOTCARR4.CARRIER_CODE = 'UPS'" _
                & " ORDER BY PACKAGE_CODE DESC")

        ImportInstructions()

        TABLE_NAME = "SOTBULK1"

        Bind_Controls(grpBulk, "SOTBULK1")
        Bind_Controls(grpGeneral, "SOTBULK1")

        dteARRIVE_BY_DATE.MaxDate = DateAdd(DateInterval.Year, 1, DateTime.Now)
        dteARRIVE_BY_DATE.MinDate = CDate("09/01/2018")

        grdSOTBULK6.AllowDrop = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "New"

                Dim rowSOTORDR0 As DataRow = Nothing
                dst.Tables("SOTORDR0").Rows.Clear()

                Dim BULK_CODE As String = MyBase.Absx1.txtFor("BULK_CODE").Text
                Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text
                Dim ORDR_CUST_PO As String = MyBase.Absx1.txtFor("ORDR_CUST_PO").Text
                Dim ORDR_GROUP_NO As String = MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text

                If BULK_CODE.Length = 0 Then
                    MyBase.Absx1.txtFor("BULK_CODE").Text = ASCMAIN1.Next_Control_No("SOTBULK1.BULK_CODE")
                End If

                Validate_Code("BULK_CODE", True)

                If CUST_CODE.Length > 0 Then
                    Validate_Code("CUST_CODE", False, True)
                End If

                If ORDR_GROUP_NO.Length > 0 Then
                    ORDR_GROUP_NO = ORDR_GROUP_NO.PadLeft(10, "0")
                    MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text = ORDR_GROUP_NO

                    rowSOTORDR0 = LookUp("SOTORDR0", ORDR_GROUP_NO)
                    If rowSOTORDR0 Is Nothing Then
                        EMsg &= vbCr & "Invalid Order Group No"
                        Exit Select
                    Else
                        If CUST_CODE.Length > 0 Then
                            If CUST_CODE <> rowSOTORDR0.Item("CUST_CODE") & String.Empty Then
                                EMsg &= vbCr & "The provided Order Group No is for customer " & CUST_CODE
                                Exit Select
                            End If
                        End If

                        If ORDR_CUST_PO.Length > 0 Then
                            If ORDR_CUST_PO <> rowSOTORDR0.Item("ORDR_CUST_PO") & String.Empty Then
                                EMsg &= vbCr & "The provided Order Group No is for customer " & CUST_CODE & ", PO " & ORDR_CUST_PO
                                Exit Select
                            End If
                        End If
                    End If
                End If

                If rowSOTORDR0 IsNot Nothing Then
                    MyBase.Absx1.txtFor("CUST_CODE").Text = rowSOTORDR0.Item("CUST_CODE") & String.Empty
                    MyBase.Absx1.txtFor("ORDR_CUST_PO").Text = rowSOTORDR0.Item("ORDR_CUST_PO") & String.Empty
                End If

            Case "Edit"
                Validate_Code("BULK_CODE")

                If EMsg.Length = 0 Then
                    Dim BULK_CODE As String = MyBase.Absx1.txtFor("BULK_CODE").Text
                    If Not ASCMAIN1.Logical_Lock("WHTBULK1", BULK_CODE) Then
                        Exit Sub
                    End If

                    ' Done so HFS has the values
                    Dim rowSOTBULK1 As DataRow = LookUp("SOTBULK1", BULK_CODE)
                    MyBase.Absx1.txtFor("CUST_CODE").Text = rowSOTBULK1.Item("CUST_CODE") & String.Empty
                    MyBase.Absx1.txtFor("ORDR_CUST_PO").Text = rowSOTBULK1.Item("ORDR_CUST_PO") & String.Empty
                    MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text = rowSOTBULK1.Item("ORDR_GROUP_NO") & String.Empty
                End If

            Case "Cancel"
                If MessageBox.Show("Do you want to Cancel your changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Update", "Save"

                MyBase.Absx1.txtFor("BULK_DESC").Text = MyBase.Absx1.txtFor("BULK_DESC").Text.Trim
                If MyBase.Absx1.txtFor("BULK_DESC").TextLength = 0 Then
                    EMsg &= vbCr & "You must provide a description."
                    Exit Select
                End If

                VerifyPatternCartonContents(String.Empty, False)
                If EMsg.Length > 0 Then
                    Exit Select
                End If

                Dim emptyCartons As Int32 = dst.Tables("SOTBULK6").Select("NUM_ITEMS = 0").Length
                If emptyCartons > 0 Then
                    Dim zmsg As String = "There are " & emptyCartons & " empty cartons. These will be deleted if you continue. Do you want to continue?"
                    If MessageBox.Show(zmsg, "Cartonize", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

                If MessageBox.Show("Do you want to " & eItemKey & " your changes?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Import Data"
                If dst.Tables("SOTBULK2").Rows.Count > 0 Then
                    If EntryMode = "E" Then
                        EMsg &= "Data has already been imported. You must create a new Bulk Distribution."
                        Exit Select
                    Else
                        If MessageBox.Show("If you import data you will lose your existing data. Do you want to continue?", "Import Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                ImportType = ImportTypes.FromFile

                Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text.Trim
                Dim ORDR_CUST_PO As String = MyBase.Absx1.txtFor("ORDR_CUST_PO").Text.Trim
                Dim ORDR_GROUP_NO As String = MyBase.Absx1.txtFor("ORDR_GROUP_NO").Text.Trim

                If CUST_CODE.Length > 0 AndAlso ORDR_CUST_PO.Length > 0 AndAlso ORDR_GROUP_NO.Length > 0 Then
                    If MessageBox.Show("Do you want to use the provided Order Group Number to import sales order data?", _
                                        "Import Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        ImportType = ImportTypes.FromReleasedSalesOrders
                    End If
                End If

                Dim rowSOTORDR0 As DataRow = Nothing

                If ImportType = ImportTypes.FromReleasedSalesOrders Then
                    Validate_Code("CUST_CODE")

                    If EMsg.Length = 0 Then
                        ASCMAIN1.sql = "SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO = '" & HFs("ORDR_GROUP_NO") & "'" '  AND ORDR_QTY_PICK > 0"
                        Fill_Records("SOTORDR0", String.Empty, True, ASCMAIN1.sql)

                        If dst.Tables("SOTORDR0").Rows.Count = 0 Then
                            EMsg &= vbCr & "Cannot locate the supplied Order Group."
                            Exit Select
                        End If

                        rowSOTORDR0 = dst.Tables("SOTORDR0").Rows(0)
                        If Val(rowSOTORDR0.Item("ORDR_QTY_PICK") & String.Empty) = 0 Then
                            If Val(rowSOTORDR0.Item("ORDR_QTY_OPEN") & String.Empty) = 0 Then
                                EMsg &= vbCr & "The supplied Order Group does not have any quantities in Pick or Open."
                                Exit Select
                            Else
                                If MessageBox.Show("The supplied Order Group number has a status of Open. Do you want to continue using the Open Quantities?", "Import", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                                    Exit Sub
                                End If
                                ImportType = ImportTypes.FromOpenSalesOrders
                            End If
                        End If

                        If EMsg.Length > 0 Then
                            Exit Select
                        End If

                    End If
                End If

                If ImportType = ImportTypes.FromFile Then
                    If MessageBox.Show("Do you want to Import an Excel Workbook?", "Import Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Transit Times"
                If MessageBox.Show("Do you want to request Transit Times for the shipments?", "Transit Time", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Delete"
                If MessageBox.Show("Are you sure you want to Delete this Bulk Distribution?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Dim pwd As String = "Annie1003"
                Dim userResponse As String = InputBox("Provide the Delete password.", "Delete")

                If pwd <> userResponse Then
                    EMsg &= vbCr & "Invalid Delete Password"
                End If

            Case "Cartonize"

                Dim f As New ASFMSGBF
                Dim selection As Int16 = f.Get_opt_from_User("Select Cartonization Type", New String() {"One Carton For All Items", "One Carton Per Item", "Use Carton/Inner Definitions", "Cancel"}, 0, "Cartonization")
                Select Case selection
                    Case 0
                        CartonizationMethod = CartonizationMethods.OneCartonForAllItems
                    Case 1
                        CartonizationMethod = CartonizationMethods.OneCartonPerItem
                    Case 2
                        CartonizationMethod = CartonizationMethods.UseCartonInnerDefinitions
                    Case 3
                        Exit Sub
                End Select

            Case "Print Pick Slip"
                If grdSOTBULK2.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "You must select stores from the Stores tab."
                    Exit Select
                End If

                If MessageBox.Show("Do you want to Print Pick Slips for the " & grdSOTBULK2.Selected.Rows.Count & " selected shipments?", "Print Pick Slip", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Generate UCC128 Cartons"
                If dst.Tables("SOTBULK1").Rows(0).Item("IMPORT_TYPE") & String.Empty <> "O" Then
                    EMsg &= vbCr & "The data needs to be imported using an Order Group Number to Generate UCC128 Cartons"
                    Exit Sub
                End If

                If dst.Tables("SOTBULK6").Rows.Count = 0 Then
                    EMsg &= vbCr & "The shipment needs to be cartonized to Generate UCC128 Cartons"
                    Exit Sub
                End If

                VerifyPatternCartonContents(String.Empty, True)

                If EMsg.Length = 0 Then
                    Dim zMsg As String = "To Generate UCC128 Cartons the following will occur: " & Environment.NewLine & Environment.NewLine _
                                         & "The current data for the Distribution will be saved." & Environment.NewLine _
                                         & "All cartons for all Pick Tickets for this Order Group will be deleted." & Environment.NewLine _
                                         & "New cartons using this Distribution will be created for all Pick Tickets for this Order Group." _
                                         & Environment.NewLine & Environment.NewLine _
                                         & "Do you want to continue?"

                    If MessageBox.Show(zMsg, "Generate UCC128 Cartons", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If

                    If MessageBox.Show("Are you sure you want to continue?", "Generate UCC128 Cartons", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Save"
                Update_Record()

            Case "Import Data"
                ImportData()
                tabBulk.ActiveTab = tabBulk.Tabs(0)
                tabBulk.SelectedTab = tabBulk.Tabs(0)

            Case "Transit Times"
                GetTransitTimes()

            Case "Delete"
                Try
                    BeginTrans()
                    For Each tableName As String In New String() {"SOTBULK1", "SOTBULK2", "SOTBULK3", "SOTBULK4", "SOTBULK5", "SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK9", "SOTBULKI"}
                        ASCMAIN1.sql = "Delete From " & tableName & " where BULK_CODE = '" & HFs("BULK_CODE") & "'"
                    Next
                    CommitTrans()
                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

                Mode_Settings(False)

            Case "Cartonize"
                Me.Cursor = Cursors.WaitCursor
                Cartonize(CartonizationMethod)
                Me.Cursor = Cursors.Default
                MessageBox.Show("Cartonization Completed. Please verify data.", "Cartonize", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Case "Print Pick Slip"
                PrintPickSlips()

            Case "Generate UCC128 Cartons"
                GenerateUCC128Cartons()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode

                    .Items("Import Data").Settings.Enabled = iScreenMode
                    .Items("Transit Times").Settings.Enabled = iScreenMode
                    .Items("Generate UCC128 Cartons").Settings.Enabled = iScreenMode
                    .Items("Cartonize").Settings.Enabled = iScreenMode
                    .Items("Print Pick Slip").Settings.Enabled = iScreenMode

                    .Items("Import Data").Visible = dst.Tables("SOTBULK2").Rows.Count = 0

                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Save").Settings.Enabled = iScreenMode

                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Delete").Visible = ASCMAIN1.USER_ID = "edz"

                End With
            End With
        End If

        txtBULK_CODE.Clear()

        If ScreenMode Then
            Set_Read_Only(grpBulk, True)
            splHolder.Panel2Collapsed = False
        Else
            Clear_Record()
            Set_Read_Only(grpBulk, False)
            splHolder.Panel2Collapsed = True
        End If

        Bind_Controls(grpBulk, "SOTBULK1")

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SOTBULK1", "SOTBULK2", "SOTBULK3", "SOTBULK4", "SOTBULK5", "SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK9", "ICTSTYC1", _
                                                       "ICTSTYLD", "SOTORDR0", "SOTORDR1", "SOTORDR5", "SOTBULKI", "SOTBULK267", "ICTSTYL1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)

        If tblItemQty.Length > 0 Then
            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & tblItemQty)
        End If

        lblOrderType.Text = String.Empty

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Gathering Data", "")
        Me.Cursor = Cursors.WaitCursor

        Save_Header_Fields(grpBulk)
        EnforceConstraints(False)

        For Each tableName As String In New String() {"SOTBULK1", "SOTBULK2", "SOTBULK3", "SOTBULK4", "SOTBULK5", "SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK9", "ICTSTYC1", "SOTBULKI", "SOTBULK267"}

            ASCMAIN1.Progress("-", tableName)

            Select Case tableName
                Case "SOTBULKI"
                    ASCMAIN1.sql = "SELECT SOTBULKI.*, ICTSTYL1.STYLE_DESC" _
                        & " FROM SOTBULKI, ICTSTYL1 " _
                        & " WHERE SOTBULKI.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)" _
                        & " AND SOTBULKI.BULK_CODE = '" & HFs("BULK_CODE") & "'"
                    Fill_Records(tableName, String.Empty, True, ASCMAIN1.sql)

                    If dst.Tables("SOTBULKI").Rows.Count > 0 Then
                        Dim Sql As String = " SELECT ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND,"
                        Sql &= " ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO, ICTSTAT2.WHSE_QTY_ON_ORDER,"
                        Sql &= " (ICTSTAT2.WHSE_QTY_ON_HAND - ICTSTAT2.WHSE_QTY_PICK) QTY_AVAIL"
                        Sql &= " FROM ICTSTAT2, SOTBULKI"
                        Sql &= " WHERE ICTSTAT2.WHSE_CODE = 'MS'"
                        Sql &= " AND ICTSTAT2.STYLE_CODE = SOTBULKI.STYLE_CODE"
                        Sql &= " AND ICTSTAT2.COLOR_CODE = SOTBULKI.COLOR_CODE"
                        Sql &= " AND SOTBULKI.BULK_CODE = '" & HFs("BULK_CODE") & "'"
                        Sql &= " GROUP BY"
                        Sql &= " ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND,"
                        Sql &= " ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO, ICTSTAT2.WHSE_QTY_ON_ORDER"

                        Dim tbl As DataTable = ASCDATA1.GetDataTable(Sql)
                        For Each row As DataRow In tbl.Select("")
                            Dim STYLE_CODE As String = row.Item("STYLE_CODE") & String.Empty
                            Dim COLOR_CODE As String = row.Item("COLOR_CODE") & String.Empty

                            Dim rowSOTBULKI As DataRow = dst.Tables("SOTBULKI").Rows.Find(New Object() {HFs("BULK_CODE"), STYLE_CODE, COLOR_CODE})
                            If rowSOTBULKI IsNot Nothing Then
                                rowSOTBULKI.Item("WHSE_QTY_ON_HAND") = row.Item("WHSE_QTY_ON_HAND")
                                rowSOTBULKI.Item("WHSE_QTY_PICK") = row.Item("WHSE_QTY_PICK")
                                rowSOTBULKI.Item("QTY_AVAIL") = row.Item("QTY_AVAIL")
                                rowSOTBULKI.Item("WHSE_QTY_ON_ORDER") = row.Item("WHSE_QTY_ON_ORDER")
                            End If
                        Next
                    End If

                Case "ICTSTYL1"
                    ASCMAIN1.sql = "SELECT *" _
                        & " FROM ICTSTYL1 " _
                        & " WHERE STYLE_CODE IN (SELECT STYLE_CODE FROM SOTBULKI WHERE BULK_CODE = '" & HFs("BULK_CODE") & "')"
                    Fill_Records(tableName, String.Empty, True, ASCMAIN1.sql)

                Case Else
                    Fill_Records(tableName, HFs("BULK_CODE"))

            End Select
        Next

        If EntryMode = "N" Then
            dst.Tables("SOTBULK1").Rows.Add(New Object() {HFs("BULK_CODE")})
            dst.Tables("SOTBULK1").Rows(0).Item("ARRIVE_BY_DATE") = CDate(DateAdd(DateInterval.Month, 1, DateTime.Now).ToShortDateString)
            dst.Tables("SOTBULK1").Rows(0).Item("WHSE_CODE") = SO_PARM_DEF_PICK_WHSE
            dst.Tables("SOTBULK1").Rows(0).Item("CUST_CODE") = HFs("CUST_CODE")
            dst.Tables("SOTBULK1").Rows(0).Item("ORDR_GROUP_NO") = HFs("ORDR_GROUP_NO")
            dst.Tables("SOTBULK1").Rows(0).Item("ORDR_CUST_PO") = HFs("ORDR_CUST_PO")
        End If


        ASCMAIN1.Progress("Extend data", "")
        ExtendData()

        EnforceConstraints(True)

        For Each TABLE_NAME As String In New String() {"SOTBULK1", "SOTBULK2", "SOTBULK3", "SOTBULK4", "SOTBULK5", "SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK9", "ICTSTYC1", _
                                               "ICTSTYLD", "SOTORDR0", "SOTORDR5", "SOTBULKI"}
            dst.Tables(TABLE_NAME).AcceptChanges()
        Next

        If dst.Tables("SOTBULK1").Rows(0).Item("IMPORT_TYPE") & String.Empty = "O" Then
            grdSOTBULK2.DisplayLayout.Bands(0).Columns("CUST_ADDR_CODE").Header.Caption = "Order No"
            grdSOTBULK2.DisplayLayout.Bands(0).Columns("CUST_REGION").Header.Caption = "Store No"

            grdSOTBULK2_SHIP.DisplayLayout.Bands(0).Columns("CUST_ADDR_CODE").Header.Caption = "Order No"
            grdSOTBULK2_SHIP.DisplayLayout.Bands(0).Columns("CUST_REGION").Header.Caption = "Store No"

            grdSOTBULK4.DisplayLayout.Bands("SOTBULK4_SOTBULK2").Columns("CUST_ADDR_CODE").Header.Caption = "Order No"
            grdSOTBULK4.DisplayLayout.Bands("SOTBULK4_SOTBULK2").Columns("CUST_REGION").Header.Caption = "Store No"

        Else
            grdSOTBULK2.DisplayLayout.Bands(0).Columns("CUST_ADDR_CODE").Header.Caption = "Store Code"
            grdSOTBULK2.DisplayLayout.Bands(0).Columns("CUST_REGION").Header.Caption = "Region"

            grdSOTBULK2_SHIP.DisplayLayout.Bands(0).Columns("CUST_ADDR_CODE").Header.Caption = "Store Code"
            grdSOTBULK2_SHIP.DisplayLayout.Bands(0).Columns("CUST_REGION").Header.Caption = "Region"

            grdSOTBULK4.DisplayLayout.Bands("SOTBULK4_SOTBULK2").Columns("CUST_ADDR_CODE").Header.Caption = "Store Code"
            grdSOTBULK4.DisplayLayout.Bands("SOTBULK4_SOTBULK2").Columns("CUST_REGION").Header.Caption = "Region"

        End If

        ASCMAIN1.Progress("Format Grids", "")
        grdSOTBULK2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
        grdSOTBULK4.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
        grdSOTBULKI.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
        grdSOTBULK6.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
        grdSOTBULK9.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
        grdSOTBULK267.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        grdSOTBULK2_SHIP.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
        grdSOTBULKI_SHIP.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        Sort_grdColumns(grdSOTBULK267, "CUST_ADDR_CODE")
        Sort_grdColumns(grdSOTBULK267, "STYLE_CODE,COLOR_CODE", , 1)

        Sort_grdColumns(grdSOTBULK4, "BULK_PATTERN_DESC")
        Sort_grdColumns(grdSOTBULK6, "BULK_PATTERN_DESC")

        Select Case dst.Tables("SOTBULK1").Rows(0).Item("IMPORT_TYPE") & String.Empty
            Case String.Empty
                lblOrderType.Text = String.Empty

            Case "F"
                lblOrderType.Text = "Data imported from a file"

            Case "O"
                If dst.Tables("SOTBULK1").Rows(0).Item("EDI_DOC_SEQ_NO") & String.Empty <> String.Empty Then
                    lblOrderType.Text = "Imported EDI Order"
                Else
                    lblOrderType.Text = "Imported Sales Order"
                End If
        End Select

        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default

    End Sub

    Sub Update_Record()

        successfulUpdate = False

        Try
            BeginTrans()

            ASCMAIN1.Progress("Updating Data", "")

            INIT_LAST("SOTBULK1")

            For Each row As DataRow In dst.Tables("SOTBULK6").Select("NUM_ITEMS = 0")
                row.Delete()
            Next

            For Each TABLE_NAME As String In New String() {"SOTBULK1", "SOTBULK2", "SOTBULK3", "SOTBULK4", "SOTBULK5", "SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK9", "SOTBULKI"}
                ASCMAIN1.Progress("-", TABLE_NAME)
                Update_Record_TDA(TABLE_NAME, "BULK_CODE = '" & HFs("BULK_CODE") & "'")
            Next

            CommitTrans("Successful Update")
            successfulUpdate = True

        Catch ex As Exception
            Rollback(ex.Message)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTBULK2, "SSSPBB", "Show Filter", "Show GroupBox", "Show Pins", "Load Zones", "Calculate Ship Date Using Transit Times")
        Load_Popup_Menu(grdSOTBULK2_SHIP, "SSSPBBPBB", "Show Filter", "Show GroupBox", "Show Pins", "Request Shipping Labels", "Get Tracking Information", "Print Shipping Labels", "Void Shipping Labels")

        Load_Popup_Menu(grdSOTBULK9, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTBULK267, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

        Load_Popup_Menu(grdSOTBULK4, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Rename Pattern")
        Load_Popup_Menu(grdSOTBULK6, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Rename Pattern")

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case grdSOTBULK4.Name, grdSOTBULK6.Name
                    tlb_btn = DirectCast(tlb_pop.Tools("Rename Pattern"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(0).Key)

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"

        End Select

        If (grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow) AndAlso e.Tool.Key <> "Update Brands" Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Rename Pattern"
                If grd.ActiveRow.Band.Key <> grd.DisplayLayout.Bands(0).Key Then
                    Exit Sub
                End If

                Dim BULK_PATTERN_DESC As String = grd.ActiveRow.Cells("BULK_PATTERN_DESC").Value & String.Empty
                Dim BULK_PATTERN_NO As String = grd.ActiveRow.Cells("BULK_PATTERN_NO").Value & String.Empty

                BULK_PATTERN_DESC = InputBox("Please Provide a new Pattern Name.", "Rename Pattern", BULK_PATTERN_DESC)
                BULK_PATTERN_DESC = BULK_PATTERN_DESC.Trim
                If BULK_PATTERN_DESC.Length = 0 Then Exit Sub
                If BULK_PATTERN_DESC.Length > dst.Tables("SOTBULK4").Columns("BULK_PATTERN_DESC").MaxLength Then
                    BULK_PATTERN_DESC = BULK_PATTERN_DESC.Substring(0, dst.Tables("SOTBULK4").Columns("BULK_PATTERN_DESC").MaxLength).Trim
                End If

                If dst.Tables("SOTBULK4").Select("BULK_PATTERN_NO <> '" & BULK_PATTERN_NO & "' and BULK_PATTERN_DESC = '" & BULK_PATTERN_DESC & "'").Length > 0 Then
                    MessageBox.Show("The provided pattern name is already in use.", "Rename Pattern", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                dst.Tables("SOTBULK4").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")(0).Item("BULK_PATTERN_DESC") = BULK_PATTERN_DESC

            Case "Calculate Ship Date Using Transit Times"

                If Not IsDate(dst.Tables("SOTBULK1").Rows(0).Item("ARRIVE_BY_DATE") & String.Empty) Then
                    MessageBox.Show("You must set the Arrive By date", "Ship Date", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim ARRIVE_BY_DATE As Date = dst.Tables("SOTBULK1").Rows(0).Item("ARRIVE_BY_DATE")

                For Each rowSOTBULK2 As DataRow In dst.Tables("SOTBULK2").Select("", "CUST_ADDR_CODE")
                    Dim TRANSIT_DAYS As Int16 = Val(rowSOTBULK2.Item("TRANSIT_DAYS") & String.Empty)
                    If TRANSIT_DAYS <= 0 Then
                        Continue For
                    End If

                    Dim numWorkingDays As Int16 = GetBusinessDays(ARRIVE_BY_DATE, TRANSIT_DAYS * -1)

                    rowSOTBULK2.Item("SHIP_DATE") = DateAdd(DateInterval.Day, numWorkingDays * -1, ARRIVE_BY_DATE)
                Next

            Case "Load Zones"
                Dim SHIP_VIA_CODE As String = Absx1.txtFor("SHIP_VIA_CODE").Text
                If SHIP_VIA_CODE.Length = 0 Then
                    MessageBox.Show("You must provide a Shipping Method.", "Load Zones", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                If WHSE_CODE.Length = 0 Then
                    MessageBox.Show("You must provide a Warehouse Code.", "Load Zones", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim row As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                If row Is Nothing Then
                    MessageBox.Show("The provided Shipping Method cannot be found.", "Load Zones", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim CARRIER_PROD_CODE As String = row.Item("CARRIER_PROD_CODE") & String.Empty
                If row Is Nothing Then
                    MessageBox.Show("The provided Shipping Method does not have an assigned carrier product code.", "Load Zones", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim CARRIER_CODE As String = row.Item("CARRIER_CODE") & String.Empty
                If row Is Nothing Then
                    MessageBox.Show("The provided Shipping Method does not have an assigned carrier.", "Load Zones", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Fill_Records("SOTCARRZ", New Object() {CARRIER_CODE, WHSE_CODE})
                If dst.Tables("SOTCARRZ").Rows.Count = 0 Then
                    MessageBox.Show("The provided Shipping Method does not zones setup for the selected Warehouse and Shipping method.", "Load Zones", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                For Each rowSOLBULK2 As DataRow In dst.Tables("SOTBULK2").Select("CUST_COUNTRY = 'US'")
                    Dim zipCode As String = rowSOLBULK2.Item("CUST_ZIP_CODE") & String.Empty
                    If zipCode.Length = 0 Then Continue For
                    If zipCode.Length > 5 Then zipCode = zipCode.Substring(0, 5)

                    rowSOLBULK2.Item("SHIP_ZONE") = String.Empty

                    For Each row In dst.Tables("SOTCARRZ").Select("DEST_ZIP_START >= '" & zipCode & "' AND DEST_ZIP_END <= '" & zipCode & "' and CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'")
                        Dim DEST_ZIP_START As String = row.Item("DEST_ZIP_START") & String.Empty
                        If DEST_ZIP_START.Length <> zipCode.Length Then Continue For
                        rowSOLBULK2.Item("SHIP_ZONE") = row.Item("CARRIER_ZONE") & String.Empty
                        Exit For
                    Next

                    If rowSOLBULK2.Item("SHIP_ZONE") & String.Empty <> String.Empty Then Continue For

                    If zipCode.Length > 0 Then zipCode = zipCode.Substring(0, 3)
                    For Each row In dst.Tables("SOTCARRZ").Select("DEST_ZIP_START <= '" & zipCode & "' AND DEST_ZIP_END >= '" & zipCode & "' and CARRIER_PROD_CODE = '" & CARRIER_PROD_CODE & "'")
                        Dim DEST_ZIP_START As String = row.Item("DEST_ZIP_START") & String.Empty
                        If DEST_ZIP_START.Length <> zipCode.Length Then Continue For
                        rowSOLBULK2.Item("SHIP_ZONE") = row.Item("CARRIER_ZONE") & String.Empty
                        Exit For
                    Next
                Next

                MessageBox.Show("Carrier Zone successful", "Load Zones", MessageBoxButtons.OK, MessageBoxIcon.Information)


            Case "Request Shipping Labels"

                'If MyBase.Absx1.txtFor("SHIP_VIA_CODE").TextLength = 0 Then
                '    MessageBox.Show("You must provide a Shipping Method.", "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                '    Exit Sub
                'End If

                If dst.Tables("SOTBULK9").Rows.Count = 0 Then
                    MessageBox.Show("You must provide a Shipping Method.", "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim domestic As Boolean = dst.Tables("SOTBULK9").Select("CARRIER_DI = 'B' OR CARRIER_DI = 'D'").Length > 0
                Dim international As Boolean = dst.Tables("SOTBULK9").Select("CARRIER_DI = 'B' OR CARRIER_DI = 'I'").Length > 0
                Dim warningMessage As String = String.Empty

                If Not domestic Then
                    warningMessage = "You do not have a Domestic shipping method defined."
                ElseIf Not international Then
                    warningMessage = "You do not have an International shipping method defined."
                End If

                If warningMessage.Length > 0 Then
                    warningMessage &= Environment.NewLine & Environment.NewLine & "Do you want to continue?"
                    If MessageBox.Show(warningMessage, "Request Shipping Labels", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

                Dim zMsg As String = String.Empty
                lstLabelsToProcess = New List(Of String)
                Dim lstAlreadyProcessed As New List(Of String)

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTBULK2_SHIP.Selected.Rows
                    lstLabelsToProcess.Add(grdRow.Cells("CUST_ADDR_CODE").Value & String.Empty)

                    If grdRow.Cells("TRACKING_NO").Value & String.Empty <> String.Empty Then
                        lstAlreadyProcessed.Add(grdRow.Cells("CUST_ADDR_CODE").Value & String.Empty)
                    End If
                Next

                If lstLabelsToProcess.Count = 0 Then
                    MessageBox.Show("You must select at least one shipment to process.", "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                EMsg = String.Empty
                VerifyPatternCartonContents(String.Empty, True)
                If EMsg.Length > 0 Then
                    MessageBox.Show(EMsg, "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If lstAlreadyProcessed.Count > 0 Then
                    zMsg = lstAlreadyProcessed.Count & " of the " & lstLabelsToProcess.Count & " selected shipments have a Tracking No. "
                    zMsg &= "Do you want to skip these shipments?" & Environment.NewLine & Environment.NewLine
                    zMsg &= "Choose 'Yes' to skip these shipments." & Environment.NewLine
                    zMsg &= "Choose 'No' to reprocess these shipments." & Environment.NewLine
                    zMsg &= "Choose 'Cancel' to abort Requesting Shipping Labels." & Environment.NewLine

                    Select Case MessageBox.Show(zMsg, "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)

                        Case Windows.Forms.DialogResult.Yes

                            For Each CUST_ADDR_CODE As String In lstAlreadyProcessed
                                lstLabelsToProcess.Remove(CUST_ADDR_CODE)
                            Next

                        Case Windows.Forms.DialogResult.No

                        Case Windows.Forms.DialogResult.Cancel
                            Exit Sub
                    End Select

                End If

                If lstLabelsToProcess.Count = 0 Then
                    Exit Sub
                End If

                zMsg = "Do you want to Request Shipping Labels for the " & lstLabelsToProcess.Count & " selected shipments?"
                If MessageBox.Show(zMsg, "Request Shipping Labels", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                ' Print by Carrier Domestic, International, Both
                Dim rowSOTBULK9_Both As DataRow = Nothing
                Dim rowSOTBULK9_Domestic As DataRow = Nothing
                Dim rowSOTBULK9_International As DataRow = Nothing

                For Each row As DataRow In dst.Tables("SOTBULK9").Select("", "CARRIER_DI, CARRIER_CODE")
                    Select Case row.Item("CARRIER_DI")
                        Case "B"
                            rowSOTBULK9_Both = row
                            rowSOTBULK9_Domestic = Nothing
                            rowSOTBULK9_International = Nothing
                            Exit For

                        Case "D"
                            rowSOTBULK9_Domestic = row
                            rowSOTBULK9_Both = Nothing

                        Case "I"
                            rowSOTBULK9_International = row
                            rowSOTBULK9_Both = Nothing
                    End Select
                Next

                If rowSOTBULK9_Both IsNot Nothing Then
                    Try
                        Me.Cursor = Cursors.WaitCursor
                        Dim ErrorMessage As String = String.Empty
                        RequestShippingLabel(ErrorMessage, rowSOTBULK9_Both)

                        If ErrorMessage.Length > 0 Then
                            MessageBox.Show(ErrorMessage, "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            MessageBox.Show("Request Shipping Labels Complete", "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        Me.Cursor = Cursors.Default
                    End Try
                End If

                If rowSOTBULK9_Domestic IsNot Nothing Then
                    Try
                        Me.Cursor = Cursors.WaitCursor
                        Dim ErrorMessage As String = String.Empty
                        RequestShippingLabel(ErrorMessage, rowSOTBULK9_Domestic)

                        If ErrorMessage.Length > 0 Then
                            MessageBox.Show(ErrorMessage, "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            MessageBox.Show("Request Shipping Labels Complete", "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        Me.Cursor = Cursors.Default
                    End Try
                End If

                If rowSOTBULK9_International IsNot Nothing Then
                    Try
                        Me.Cursor = Cursors.WaitCursor
                        Dim ErrorMessage As String = String.Empty
                        RequestShippingLabel(ErrorMessage, rowSOTBULK9_International)

                        If ErrorMessage.Length > 0 Then
                            MessageBox.Show(ErrorMessage, "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Else
                            MessageBox.Show("Request Shipping Labels Complete", "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.Message, "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Finally
                        Me.Cursor = Cursors.Default
                    End Try
                End If


            Case "Print UPS Invoice"
                PrintUPSInvoice()

            Case "Get Tracking Information"
                Dim zMsg As String = String.Empty
                lstLabelsToProcess = New List(Of String)
                Dim lstItems As New List(Of String)

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTBULK2_SHIP.Selected.Rows
                    If grdRow.Cells("TRACKING_NO").Value & String.Empty <> String.Empty Then
                        lstLabelsToProcess.Add(grdRow.Cells("CUST_ADDR_CODE").Value & String.Empty)
                    Else
                        grdRow.Selected = False
                    End If
                Next

                If lstLabelsToProcess.Count = 0 Then
                    MessageBox.Show("You must select at least one shipment with a Tracking No.", "Get Tracking Information", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If MessageBox.Show("Do you want to Get Tracking Information for the selected (" & lstLabelsToProcess.Count & ") Shipments?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                GetTrackingInformation()

                MessageBox.Show("Get Tracking Information Complete.", "Get Tracking Information", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Case "Void Shipping Labels"
                Dim zMsg As String = String.Empty
                lstLabelsToProcess = New List(Of String)
                Dim lstItems As New List(Of String)

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTBULK2_SHIP.Selected.Rows
                    If grdRow.Cells("TRACKING_NO").Value & String.Empty <> String.Empty Then
                        lstLabelsToProcess.Add(grdRow.Cells("CUST_ADDR_CODE").Value & String.Empty)
                    Else
                        grdRow.Selected = False
                    End If
                Next

                If lstLabelsToProcess.Count = 0 Then
                    MessageBox.Show("You must select at least one shipment with a Tracking No.", "Void Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If MessageBox.Show("Do you want to Void the selected (" & lstLabelsToProcess.Count & ") Shipments?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                If InputBox("Please provide the password to Void Shipping Labels?", "Void Shipping Labels") <> "Annabelle" Then
                    MessageBox.Show("Invalid Void Shipping Labels password.", "Void Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                VoidShippingLabel()

                MessageBox.Show("Void Shipping Label Complete.", "Void Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)


            Case "Print Shipping Labels"

                Dim zMsg As String = String.Empty
                lstLabelsToProcess = New List(Of String)
                Dim lstItems As New List(Of String)

                For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTBULK2_SHIP.Selected.Rows
                    If grdRow.Cells("TRACKING_NO").Value & String.Empty <> String.Empty Then
                        lstLabelsToProcess.Add(grdRow.Cells("CUST_ADDR_CODE").Value & String.Empty)
                    Else
                        grdRow.Selected = False
                    End If
                Next

                If lstLabelsToProcess.Count = 0 Then
                    MessageBox.Show("You must select at least one shipment with a Tracking No.", "Request Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                For Each rowSOTBULKI As DataRow In dst.Tables("SOTBULKI").Select("SELECTED = '1'")
                    lstItems.Add(rowSOTBULKI.Item("STYLE_CODE") & " / " & rowSOTBULKI.Item("COLOR_CODE"))
                Next

                zMsg = "Do you want to Print Shipping Labels for the " & lstLabelsToProcess.Count & " selected shipments?"
                If lstItems.Count > 0 Then
                    zMsg &= Environment.NewLine & Environment.NewLine
                    zMsg &= "For the following items:" & Environment.NewLine

                    For Each item As String In lstItems
                        zMsg &= item & Environment.NewLine
                    Next
                End If

                If MessageBox.Show(zMsg, "Print Shipping Labels", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Try
                    Me.Cursor = Cursors.WaitCursor
                    Dim ErrorMessage As String = String.Empty
                    Dim labelsPrinted As Int64 = 0
                    PrintShippingLabel(ErrorMessage, labelsPrinted, False)

                    If labelsPrinted > 0 Then
                        If MessageBox.Show(labelsPrinted & " labels will print. Verify you have enough labels in the printer." & Environment.NewLine & Environment.NewLine _
                                           & "Click 'Yes' to print the labels" & Environment.NewLine _
                                           & "Click 'No' to not print the labels.", _
                                            "Print Shipping Labels", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    labelsPrinted = 0
                    PrintShippingLabel(ErrorMessage, labelsPrinted, True)

                    If ErrorMessage.Length > 0 Then
                        MessageBox.Show(ErrorMessage, "Print Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Else
                        MessageBox.Show("Print Shipping Labels Complete - " & labelsPrinted & " labels sent to the label printer.", "Print Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Print Shipping Labels", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    Me.Cursor = Cursors.Default
                End Try

        End Select
    End Sub


#End Region

#Region "Form Procedures"

    Private Function CountryOfOrigin(ByVal CountryCode As String) As String

        Dim convertedCountryCode As String = String.Empty

        Select Case CountryCode
            Case "ABW" 'ARUBA
            Case "AFG" 'AFGHANISTAN
            Case "AGO" 'ANGOLA
            Case "AIA" 'ANGUILLA
            Case "ALB" 'ALBANIA
            Case "AND" 'ANDORRA
            Case "ANT" 'NETHERLANDS ANTILLES
            Case "ARE" 'UNITED ARAB EMIRATES
            Case "ARG" 'ARGENTINA
            Case "ARM" 'ARMENIA
            Case "ASM" 'AMERICAN SAMOA
            Case "ATA" 'ANTARCTICA
            Case "ATF" 'FRENCH SOUTHERN TERRITORIES
            Case "ATG" 'ANTIGUA AND BARBUDA
            Case "AUS" 'AUSTRALIA
            Case "AUT" 'AUSTRIA
            Case "AZE" 'AZERBAIJAN
            Case "BDI" 'BURUNDI
            Case "BEL" 'BELGIUM
            Case "BEN" 'BENIN
            Case "BFA" 'BURKINA FASO
            Case "BGD" 'BANGLADESH
            Case "BGR" 'BULGARIA
            Case "BHR" 'BAHRAIN
            Case "BHS" 'BAHAMAS
            Case "BIH" 'BOSNIA AND HERZEGOWINA
            Case "BLR" 'BELARUS
            Case "BLZ" 'BELIZE
            Case "BMU" 'BERMUDA
            Case "BOL" 'BOLIVIA
            Case "BRA" 'BRAZIL
            Case "BRB" 'BARBADOS
            Case "BRN" 'BRUNEI DARUSSALAM
            Case "BTN" 'BHUTAN
            Case "BVT" 'BOUVET ISLAND
            Case "BWA" 'BOTSWANA
            Case "CAF" 'CENTRAL AFRICAN REPUBLIC
            Case "CAN" 'CANADA
            Case "CCK" 'COCOS (KEELING) ISLANDS
            Case "CHE" 'SWITZERLAND
                convertedCountryCode = "CH"
            Case "CHL" 'CHILE
            Case "CHN" 'CHINA
                convertedCountryCode = "CN"
            Case "CIV" 'COTE D'IVOIRE
            Case "CMR" 'CAMEROON
            Case "COD" 'CONGO, THE DRC
            Case "COG" 'CONGO
            Case "COK" 'COOK ISLANDS
            Case "COL" 'COLOMBIA
            Case "COM" 'COMOROS
            Case "CPV" 'CAPE VERDE
            Case "CRI" 'COSTA RICA
            Case "CUB" 'CUBA
            Case "CXR" 'CHRISTMAS ISLAND
            Case "CYM" 'CAYMAN ISLANDS
            Case "CYP" 'CYPRUS
            Case "CZE" 'CZECH REPUBLIC
            Case "DEU" 'GERMANY
            Case "DJI" 'DJIBOUTI
            Case "DMA" 'DOMINICA
            Case "DNK" 'DENMARK
            Case "DOM" 'DOMINICAN REPUBLIC
            Case "DZA" 'ALGERIA
            Case "ECU" 'ECUADOR
            Case "EGY" 'EGYPT
            Case "ERI" 'ERITREA
            Case "ESH" 'WESTERN SAHARA
            Case "ESP" 'SPAIN
            Case "EST" 'ESTONIA
            Case "ETH" 'ETHIOPIA
            Case "FIN" 'FINLAND
            Case "FJI" 'FIJI
            Case "FLK" 'FALKLAND ISLANDS (MALVINAS)
            Case "FRA" 'FRANCE
            Case "FRO" 'FAROE ISLANDS
            Case "FSM" 'MICRONESIA, FEDERATED STATES OF
            Case "FXX" 'FRANCE, METROPOLITAN
            Case "GAB" 'GABON
            Case "GBR" 'UNITED KINGDOM
            Case "GEO" 'GEORGIA
            Case "GHA" 'GHANA
            Case "GIB" 'GIBRALTAR
            Case "GIN" 'GUINEA
            Case "GLP" 'GUADELOUPE
            Case "GMB" 'GAMBIA
            Case "GNB" 'GUINEA-BISSAU
            Case "GNQ" 'EQUATORIAL GUINEA
            Case "GRC" 'GREECE
            Case "GRD" 'GRENADA
            Case "GRL" 'GREENLAND
            Case "GTM" 'GUATEMALA
            Case "GUF" 'FRENCH GUIANA
            Case "GUM" 'GUAM
            Case "GUY" 'GUYANA
            Case "HKG" 'HONG KONG
            Case "HMD" 'HEARD AND MC DONALD ISLANDS
            Case "HND" 'HONDURAS
                convertedCountryCode = "HN"
            Case "HRV" 'CROATIA (local name: Hrvatska)
            Case "HTI" 'HAITI
            Case "HUN" 'HUNGARY
            Case "IDN" 'INDONESIA
            Case "IND" 'INDIA
                convertedCountryCode = "IN"
            Case "IOT" 'BRITISH INDIAN OCEAN TERRITORY
            Case "IRL" 'IRELAND
            Case "IRN" 'IRAN (ISLAMIC REPUBLIC OF)
            Case "IRQ" 'IRAQ
            Case "ISL" 'ICELAND
            Case "ISR" 'ISRAEL
            Case "ITA" 'ITALY
            Case "JAM" 'JAMAICA
            Case "JOR" 'JORDAN
            Case "JPN" 'JAPAN
            Case "KAZ" 'KAZAKHSTAN
            Case "KEN" 'KENYA
            Case "KGZ" 'KYRGYZSTAN
            Case "KHM" 'CAMBODIA
                convertedCountryCode = "KH"
            Case "KIR" 'KIRIBATI
            Case "KNA" 'SAINT KITTS AND NEVIS
            Case "KOR" 'KOREA, REPUBLIC OF
            Case "KWT" 'KUWAIT
            Case "LAO" 'LAOS
            Case "LBN" 'LEBANON
            Case "LBR" 'LIBERIA
            Case "LBY" 'LIBYAN ARAB JAMAHIRIYA
            Case "LCA" 'SAINT LUCIA
            Case "LIE" 'LIECHTENSTEIN
            Case "LKA" 'SRI LANKA
            Case "LSO" 'LESOTHO
            Case "LTU" 'LITHUANIA
            Case "LUX" 'LUXEMBOURG
            Case "LVA" 'LATVIA
            Case "MAC" 'MACAU
            Case "MAR" 'MOROCCO
            Case "MCO" 'MONACO
            Case "MDA" 'MOLDOVA, REPUBLIC OF
            Case "MDG" 'MADAGASCAR
            Case "MDV" 'MALDIVES
            Case "MEX" 'MEXICO
            Case "MHL" 'MARSHALL ISLANDS
            Case "MKD" 'MACEDONIA
            Case "MLI" 'MALI
            Case "MLT" 'MALTA
            Case "MMR" 'MYANMAR (Burma)
            Case "MNE" 'MONTENEGRO
            Case "MNG" 'MONGOLIA
            Case "MNP" 'NORTHERN MARIANA ISLANDS
            Case "MOZ" 'MOZAMBIQUE
            Case "MRT" 'MAURITANIA
            Case "MSR" 'MONTSERRAT
            Case "MTQ" 'MARTINIQUE
            Case "MUS" 'MAURITIUS
            Case "MWI" 'MALAWI
            Case "MYS" 'MALAYSIA
            Case "MYT" 'MAYOTTE
            Case "NAM" 'NAMIBIA
            Case "NCL" 'NEW CALEDONIA
            Case "NER" 'NIGER
            Case "NFK" 'NORFOLK ISLAND
            Case "NGA" 'NIGERIA
            Case "NIC" 'NICARAGUA
            Case "NIU" 'NIUE
            Case "NLD" 'NETHERLANDS
            Case "NOR" 'NORWAY
            Case "NPL" 'NEPAL
            Case "NRU" 'NAURU
            Case "NZL" 'NEW ZEALAND
            Case "OMN" 'OMAN
            Case "PAK" 'PAKISTAN
            Case "PAN" 'PANAMA
            Case "PCN" 'PITCAIRN
            Case "PER" 'PERU
            Case "PHL" 'PHILIPPINES
                convertedCountryCode = "PH"
            Case "PLW" 'PALAU
            Case "PNG" 'PAPUA NEW GUINEA
            Case "POL" 'POLAND
            Case "PRI" 'PUERTO RICO
            Case "PRK" 'KOREA, D.P.R.O.
            Case "PRT" 'PORTUGAL
            Case "PRY" 'PARAGUAY
            Case "PYF" 'FRENCH POLYNESIA
            Case "QAT" 'QATAR
            Case "REU" 'REUNION
                convertedCountryCode = "RE"
            Case "ROM" 'ROMANIA
            Case "RUS" 'RUSSIAN FEDERATION
            Case "RWA" 'RWANDA
            Case "SAU" 'SAUDI ARABIA
            Case "SDN" 'SUDAN
            Case "SEN" 'SENEGAL
            Case "SGP" 'SINGAPORE
            Case "SGS" 'SOUTH GEORGIA AND SOUTH S.S.
            Case "SHN" 'ST. HELENA
            Case "SJM" 'SVALBARD AND JAN MAYEN ISLANDS
            Case "SLB" 'SOLOMON ISLANDS
            Case "SLE" 'SIERRA LEONE
            Case "SLV" 'EL SALVADOR
            Case "SMR" 'SAN MARINO
            Case "SOM" 'SOMALIA
            Case "SPM" 'ST. PIERRE AND MIQUELON
            Case "SRB" 'SERBIA
            Case "SSD" 'SOUTH SUDAN
            Case "STP" 'SAO TOME AND PRINCIPE
            Case "SUR" 'SURINAME
            Case "SVK" 'SLOVAKIA (Slovak Republic)
            Case "SVN" 'SLOVENIA
            Case "SWE" 'SWEDEN
            Case "SWZ" 'SWAZILAND
            Case "SYC" 'SEYCHELLES
            Case "SYR" 'SYRIAN ARAB REPUBLIC
            Case "TCA" 'TURKS AND CAICOS ISLANDS
            Case "TCD" 'CHAD
            Case "TGO" 'TOGO
            Case "THA" 'THAILAND
                convertedCountryCode = "TH"
            Case "TJK" 'TAJIKISTAN
            Case "TKL" 'TOKELAU
            Case "TKM" 'TURKMENISTAN
            Case "TMP" 'EAST TIMOR
            Case "TON" 'TONGA
            Case "TTO" 'TRINIDAD AND TOBAGO
            Case "TUN" 'TUNISIA
            Case "TUR" 'TURKEY
            Case "TUV" 'TUVALU
            Case "TWN" 'TAIWAN, PROVINCE OF CHINA
                convertedCountryCode = "TW"
            Case "TZA" 'TANZANIA, UNITED REPUBLIC OF
            Case "UGA" 'UGANDA
            Case "UKR" 'UKRAINE
            Case "UMI" 'U.S. MINOR ISLANDS
            Case "URY" 'URUGUAY
            Case "USA" 'UNITED STATES
                convertedCountryCode = "US"
            Case "UZB" 'UZBEKISTAN
            Case "VAT" 'HOLY SEE (VATICAN CITY STATE)
            Case "VCT" 'SAINT VINCENT AND THE GRENADINES
            Case "VEN" 'VENEZUELA
            Case "VGB" 'VIRGIN ISLANDS (BRITISH)
            Case "VIR" 'VIRGIN ISLANDS (U.S.)
            Case "VNM" 'VIET NAM
            Case "VUT" 'VANUATU
            Case "WLF" 'WALLIS AND FUTUNA ISLANDS
            Case "WSM" 'SAMOA
            Case "YEM" 'YEMEN
            Case "ZAF" 'SOUTH AFRICA
            Case "ZMB" 'ZAMBIA
            Case "ZWE" 'ZIMBABWE
            Case "ZZZ"
                convertedCountryCode = "US"
        End Select

        Return convertedCountryCode

    End Function

    Private Sub PrintUPSInvoice()

        Try
            If grdSOTBULK2_SHIP.Selected.Rows.Count = 0 Then
                MessageBox.Show("You must select at least one shipment.", "Print UPS Invoice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim lstCUST_ADDR_CODE As New List(Of String)
            For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTBULK2_SHIP.Selected.Rows
                lstCUST_ADDR_CODE.Add(grdRow.Cells("CUST_ADDR_CODE").Value)
            Next

            'Dim DefaultPrinterName As String = String.Empty
            'Dim oPS As New System.Drawing.Printing.PrinterSettings
            'Try
            '    DefaultPrinterName = oPS.PrinterName
            'Catch ex As System.Exception
            '    DefaultPrinterName = ""
            'End Try

            'If DefaultPrinterName.Length = 0 Then
            '    MessageBox.Show("You do not have a default Printer.", "Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
            '    Exit Sub
            'End If

            Dim REPORT_NAME As String = "WHRSHIPE"
            ASCMAIN1.Progress("Printing UPS Invoices", "")

            For Each tableName As String In New String() {"UPSINTL1", "UPSINTL2", "UPSINTL5"}
                dst.Tables(tableName).Rows.Clear()
            Next

            Dim TOTAL_VALUE As Double = 0

            Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = '" & HFs("CUST_CODE") & "'")

            For Each CUST_ADDR_CODE As String In lstCUST_ADDR_CODE

                ASCMAIN1.Progress("-", CUST_ADDR_CODE)

                Dim rowSOTBULK2 As DataRow = dst.Tables("SOTBULK2").Select("BULK_CODE = '" & HFs("BULK_CODE") & "' AND CUST_ADDR_CODE = '" & CUST_ADDR_CODE & "'")(0)
                Dim BULK_PATTERN_NO As String = rowSOTBULK2.Item("BULK_PATTERN_NO")

                Dim tblwork As DataTable = dst.Tables("SOTBULK7").Clone
                tblwork.TableName = "SOTBULK7_WK"
                For Each row As DataRow In dst.Tables("SOTBULK7").Select("BULK_CODE = '" & HFs("BULK_CODE") & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")
                    tblwork.ImportRow(row)
                Next

                Dim tbl As DataTable = ASCDATA1.SelectDistinct(tblwork, New String() {"STYLE_CODE", "COLOR_CODE"})

                For Each row As DataRow In tbl.Select("", "")
                    Dim STYLE_CODE As String = row.Item("STYLE_CODE") & String.Empty
                    Dim COLOR_CODE As String = row.Item("COLOR_CODE") & String.Empty

                    Dim ORDR_QTY As Int32 = Val(dst.Tables("SOTBULK5").Compute("SUM(ORDR_QTY)", _
                        "BULK_CODE = '" & HFs("BULK_CODE") & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)

                    If ORDR_QTY = 0 Then Continue For

                    Dim rowSOTBULKI As DataRow = dst.Tables("SOTBULKI").Rows.Find(New Object() {HFs("BULK_CODE"), STYLE_CODE, COLOR_CODE})
                    If rowSOTBULKI Is Nothing Then
                        Continue For
                    End If

                    Dim rowUPSINTL1 As DataRow = dst.Tables("UPSINTL1").NewRow
                    rowUPSINTL1.Item("BULK_CODE") = HFs("BULK_CODE")
                    rowUPSINTL1.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL1.Item("UNITS") = ORDR_QTY
                    rowUPSINTL1.Item("UOM") = "EA"
                    rowUPSINTL1.Item("DESC") = rowSOTBULKI.Item("STYLE_DESC") & String.Empty
                    rowUPSINTL1.Item("UNIT_VALUE") = Val(rowSOTBULKI.Item("STYLE_PRICE") & String.Empty)
                    rowUPSINTL1.Item("TOTAL_VALUE") = ORDR_QTY * rowUPSINTL1.Item("UNIT_VALUE")
                    rowUPSINTL1.Item("ORIGIN_COUNTRY") = rowSOTBULKI.Item("COUNTRY_CODE") & String.Empty
                    rowUPSINTL1.Item("ORIGIN_COUNTRY") = CountryOfOrigin(rowUPSINTL1.Item("ORIGIN_COUNTRY") & String.Empty)

                    dst.Tables("UPSINTL1").Rows.Add(rowUPSINTL1)

                    TOTAL_VALUE += rowUPSINTL1.Item("TOTAL_VALUE")
                Next

                Dim rowUPSINTL2 As DataRow = dst.Tables("UPSINTL2").NewRow
                rowUPSINTL2.Item("BULK_CODE") = HFs("BULK_CODE")
                rowUPSINTL2.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                rowUPSINTL2.Item("INVOICE_LINE_TOTAL") = TOTAL_VALUE
                rowUPSINTL2.Item("DISCOUNT") = 0
                rowUPSINTL2.Item("INVOICE_SUBTOTAL") = TOTAL_VALUE
                rowUPSINTL2.Item("FREIGHT") = 0
                rowUPSINTL2.Item("INSURANCE") = 0
                rowUPSINTL2.Item("OTHER") = 0
                rowUPSINTL2.Item("TOTAL_INVOICE_AMOUNT") = TOTAL_VALUE
                rowUPSINTL2.Item("NUM_CARTONS") = dst.Tables("SOTBULK6").Select("BULK_CODE = '" & HFs("BULK_CODE") & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'").Length
                rowUPSINTL2.Item("CURRENCY") = "USD"
                rowUPSINTL2.Item("TOTAL_WEIGHT") = Val(dst.Tables("SOTBULK6").Compute("SUM(PKG_WEIGHT)", "BULK_CODE = '" & HFs("BULK_CODE") & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'") & String.Empty)
                dst.Tables("UPSINTL2").Rows.Add(rowUPSINTL2)

                Dim rowUPSINTL5 As DataRow = dst.Tables("UPSINTL5").NewRow
                If ASCMAIN1.CLIENT = "RGI" AndAlso HFs("CUST_CODE") = "021454" And rowSOTBULK2.Item("CUST_COUNTRY") = "CA" Then
                    rowUPSINTL5.Item("BULK_CODE") = HFs("BULK_CODE")
                    rowUPSINTL5.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL5.Item("CUST_CONTACT") = "Manager"
                    rowUPSINTL5.Item("CUST_NAME") = "The Gap (Canada), Inc."
                    rowUPSINTL5.Item("CUST_ADDR1") = "60 Bloor Street West"
                    rowUPSINTL5.Item("CUST_ADDR2") = "Suite 1500"
                    rowUPSINTL5.Item("CUST_CITY") = "Toronto"
                    rowUPSINTL5.Item("CUST_STATE") = "Ontario"
                    rowUPSINTL5.Item("CUST_ZIP_CODE") = "M4W 3B8"
                    rowUPSINTL5.Item("CUST_PHONE") = rowSOTBULK2.Item("CUST_AREA_CODE") & rowSOTBULK2.Item("CUST_PHONE")

                ElseIf rowARTCUST1 IsNot Nothing Then
                    rowUPSINTL5.Item("BULK_CODE") = HFs("BULK_CODE")
                    rowUPSINTL5.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL5.Item("CUST_CONTACT") = rowARTCUST1.Item("CUST_CONTACT")
                    rowUPSINTL5.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                    rowUPSINTL5.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1")
                    rowUPSINTL5.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2")
                    rowUPSINTL5.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY")
                    rowUPSINTL5.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE")
                    rowUPSINTL5.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE")
                    rowUPSINTL5.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE")

                Else
                    rowUPSINTL5.Item("BULK_CODE") = HFs("BULK_CODE")
                    rowUPSINTL5.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL5.Item("CUST_CONTACT") = rowSOTBULK2.Item("CUST_ADDR_CODE")
                    rowUPSINTL5.Item("CUST_NAME") = rowSOTBULK2.Item("CUST_NAME")
                    rowUPSINTL5.Item("CUST_ADDR1") = rowSOTBULK2.Item("CUST_ADDR1")
                    rowUPSINTL5.Item("CUST_ADDR2") = rowSOTBULK2.Item("CUST_ADDR2")
                    rowUPSINTL5.Item("CUST_CITY") = rowSOTBULK2.Item("CUST_CITY")
                    rowUPSINTL5.Item("CUST_STATE") = rowSOTBULK2.Item("CUST_STATE")
                    rowUPSINTL5.Item("CUST_ZIP_CODE") = rowSOTBULK2.Item("CUST_ZIP_CODE")
                    rowUPSINTL5.Item("CUST_PHONE") = rowSOTBULK2.Item("CUST_AREA_CODE") & rowSOTBULK2.Item("CUST_PHONE")
                End If

                dst.Tables("UPSINTL5").Rows.Add(rowUPSINTL5)
            Next

            Print_Report_Begin()
            Generate_Report(REPORT_NAME, "UPS Invoice")
            Print_Report_End()


        Catch ex As Exception
            MessageBox.Show("The following error ocurred while Printing Pick Slips: " & ex.Message)

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub GenerateUCC128Cartons()

        Try
            successfulUpdate = False
            Update_Record()

            If Not successfulUpdate Then
                Exit Sub
            End If

            If MessageBox.Show("Are you sure you want to Generate UCC128 Cartons?", "Generate UCC128 Cartons", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If

        Catch ex As Exception
            MessageBox.Show("Generate UCC128 Cartons Error: " & ex.Message)
            Exit Sub
        End Try

        Try

            dst.Tables("SOTCART1").Rows.Clear()
            dst.Tables("SOTCART2").Rows.Clear()

            Dim temp_table As String = ASCMAIN1.Temp_Table("Select SOTPICK1.SHIP_BOL_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_NO FROM SOTPICK1 WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE ORDR_GROUP_NO = '" & HFs("ORDR_GROUP_NO") & "')")
            Dim tbl As DataTable = ASCDATA1.GetDataTable("Select * From " & temp_table)

            ' Lock Everything
            If Not ASCMAIN1.Logical_Lock("SOTORDR0", HFs("ORDR_GROUP_NO"), False, True, True, 9) Then
                Exit Sub
            End If

            For Each row As DataRow In ASCDATA1.SelectDistinct(tbl, New String() {"SHIP_BOL_NO"}).Select("")
                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", row.Item("SHIP_BOL_NO"), False, True, True, 9) Then
                    Exit Sub
                End If
            Next

            For Each row As DataRow In tbl.Select("")
                If Not ASCMAIN1.Logical_Lock("SOTORDR1", row.Item("ORDR_NO"), False, True, True, 9) Then
                    Exit Sub
                End If

                If Not ASCMAIN1.Logical_Lock("SOTPICK1", row.Item("PICK_NO"), False, True, True, 9) Then
                    Exit Sub
                End If
            Next

            BeginTrans()

            ASCMAIN1.sql = "SELECT SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                & " FROM SOTPICK2, SOTORDR2" _
                & " WHERE SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO" _
                & " AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO" _
                & " AND SOTPICK2.PICK_NO IN (SELECT PICK_NO FROM " & temp_table & ")"
            Dim tblSOTPICK2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            For Each wkRow As DataRow In tbl.Select("", "ORDR_NO")
                Dim ORDR_NO As String = wkRow.Item("ORDR_NO")
                Dim PICK_NO As String = wkRow.Item("PICK_NO")
                Dim CART_SEQ As Int32 = 0

                Dim rowSOTBULK2 As DataRow = dst.Tables("SOTBULK2").Select("CUST_ADDR_CODE = '" & ORDR_NO & "'")(0)
                Dim BULK_PATTERN_NO = rowSOTBULK2.Item("BULK_PATTERN_NO")

                For Each rowSOTBULK6 As DataRow In dst.Tables("SOTBULK6").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'", "CART_NO")
                    Dim bCART_NO As String = rowSOTBULK6.Item("CART_NO")

                    If dst.Tables("SOTBULK7").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND CART_NO = '" & bCART_NO & "'").Length = 0 Then
                        Continue For
                    End If

                    Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                    Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
                    rowSOTCART1.Item("CART_NO") = CART_NO
                    rowSOTCART1.Item("CART_FREIGHT") = 0
                    'rowSOTCART1.Item("CART_PACKER") = String.Empty
                    'rowSOTCART1.Item("CART_PACKED") = String.Empty
                    'rowSOTCART1.Item("CART_SHIPPED") = String.Empty
                    rowSOTCART1.Item("PICK_NO") = PICK_NO
                    rowSOTCART1.Item("CART_TOTAL_UNITS") = Val(dst.Tables("SOTBULK7").Compute("SUM(QTY_PACKED)", "BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND CART_NO = '" & bCART_NO & "'") & String.Empty)
                    rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = rowSOTBULK6.Item("PKG_WEIGHT")
                    rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = rowSOTBULK6.Item("PKG_WEIGHT")
                    'rowSOTCART1.Item("CART_TRACKING_NO") = String.Empty

                    CART_SEQ += 1
                    rowSOTCART1.Item("CART_SEQ") = CART_SEQ

                    'rowSOTCART1.Item("CART_MEMO") = String.Empty
                    'rowSOTCART1.Item("CART_TYPE") = String.Empty
                    rowSOTCART1.Item("PACKAGING_TYPE") = rowSOTBULK6.Item("PACKAGING_TYPE") & String.Empty
                    rowSOTCART1.Item("PKG_CODE") = rowSOTBULK6.Item("PKG_CODE") & String.Empty
                    rowSOTCART1.Item("PKG_L") = rowSOTBULK6.Item("PKG_L") & String.Empty
                    rowSOTCART1.Item("PKG_W") = rowSOTBULK6.Item("PKG_W") & String.Empty
                    rowSOTCART1.Item("PKG_H") = rowSOTBULK6.Item("PKG_H") & String.Empty
                    dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

                    Dim CART_LNO As Int16 = 0

                    For Each rowSOTBULK7 As DataRow In dst.Tables("SOTBULK7").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND CART_NO = '" & bCART_NO & "'")
                        Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                        rowSOTCART2.Item("CART_NO") = CART_NO

                        CART_LNO += 1
                        rowSOTCART2.Item("CART_LNO") = CART_LNO

                        Dim rowSOTPICK2 As DataRow = tblSOTPICK2.Select("ORDR_NO = '" & ORDR_NO & "' AND STYLE_CODE = '" & rowSOTBULK7.Item("STYLE_CODE") & "' AND COLOR_CODE = '" & rowSOTBULK7.Item("COLOR_CODE") & "'")(0)

                        rowSOTCART2.Item("ORDR_NO") = ORDR_NO
                        rowSOTCART2.Item("ORDR_LNO") = rowSOTPICK2.Item("ORDR_LNO")
                        rowSOTCART2.Item("QTY_PACKED") = rowSOTBULK7.Item("QTY_PACKED")
                        rowSOTCART2.Item("UPC_CODE") = rowSOTBULK7.Item("UPC_CODE")
                        'rowSOTCART2.Item("SKU_NO") = String.Empty
                        rowSOTCART2.Item("STYLE_CODE") = rowSOTBULK7.Item("STYLE_CODE")
                        rowSOTCART2.Item("COLOR_CODE") = rowSOTBULK7.Item("COLOR_CODE")
                        'rowSOTCART2.Item("SIZE_DESC") = String.Empty
                        'rowSOTCART2.Item("STYLE_PREPACK") = String.Empty
                        'rowSOTCART2.Item("ITEM_EXP_DATE") = String.Empty
                        dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                    Next
                Next
            Next

            ASCDATA1.ExecuteSQL("DELETE FROM SOTCART2 WHERE CART_NO IN (SELECT CART_NO FROM SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & temp_table & "))")
            ASCDATA1.ExecuteSQL("DELETE FROM SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & temp_table & ")")

            Update_Record_TDA("SOTCART1")
            Update_Record_TDA("SOTCART2")

            CommitTrans("Successful Generation of UCC128 Cartons")

        Catch ex As Exception
            Rollback("Generate UCC128 Cartons Error: " & ex.Message)

        Finally
            ASCMAIN1.MultiTask_Release(, , 9)
        End Try

    End Sub

    Private Sub ImportData()

        Dim clearedTables As Boolean = False
        Dim tableData As New DataTable
        Dim fieldName As String = String.Empty
        Dim lstItems As New List(Of String)
        Const numfields As Int16 = 14
        Dim availableLoaded As Boolean = False

        Try

            If tblItemQty.Length > 0 Then
                ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & tblItemQty)
            End If

            If ImportType = ImportTypes.FromFile Then

                Dim fileToImport As String = String.Empty

                Using openFileDialog1 As New OpenFileDialog
                    'openFileDialog1.InitialDirectory = "C:\TEMP"
                    openFileDialog1.Title = "Open Bulk Distro File"
                    openFileDialog1.Filter = "Excel files (*.xls)|*.xls"
                    openFileDialog1.FilterIndex = 1
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        fileToImport = openFileDialog1.FileName
                    End If

                    openFileDialog1.Dispose()
                End Using

                If fileToImport.Length = 0 Then
                    Exit Sub
                End If

                Dim filename As String = My.Computer.FileSystem.GetName(fileToImport)
                If filename.Length > dst.Tables("SOTBULK1").Columns("BULK_DESC").MaxLength Then
                    filename = filename.Substring(0, dst.Tables("SOTBULK1").Columns("BULK_DESC").MaxLength).Trim
                End If

                If MyBase.Absx1.txtFor("BULK_DESC").TextLength = 0 Then
                    MyBase.Absx1.txtFor("BULK_DESC").Text = filename
                    dst.Tables("SOTBULK1").Rows(0).Item("BULK_DESC") = filename
                End If

                ASCMAIN1.Progress("Reading File")
                Me.Cursor = Cursors.WaitCursor

                ' "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileToImport & ";Extended Properties=""Excel 12.0;HDR=YES;IMEX=1"""
                ' "provider=Microsoft.Jet.OLEDB.4.0;" & "data source=" & fileToImport & ";Extended Properties=Excel 8.0;"
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

                For icol As Integer = 1 To numfields
                    Select Case icol
                        Case 1
                            fieldName = "CUST_ADDR_CODE"
                        Case 2
                            fieldName = "CUST_NAME"
                        Case 3
                            fieldName = "CUST_DISTRICT"
                        Case 4
                            fieldName = "CUST_REGION"
                        Case 5
                            fieldName = "CUST_BRAND"
                        Case 6
                            fieldName = "CUST_AREA_CODE"
                        Case 7
                            fieldName = "CUST_PHONE"
                        Case 8
                            fieldName = "CUST_ADDR1"
                        Case 9
                            fieldName = "CUST_ADDR2"
                        Case 10
                            fieldName = "CUST_ADDR3"
                        Case 11
                            fieldName = "CUST_CITY"
                        Case 12
                            fieldName = "CUST_STATE"
                        Case 13
                            fieldName = "CUST_ZIP_CODE"
                        Case 14
                            fieldName = "CUST_COUNTRY"

                        Case Else
                            ' should be the style code
                    End Select

                    If fieldName.Length > 0 Then
                        tableData.Columns(icol - 1).ColumnName = fieldName.Trim
                    End If
                Next


                ' Get Items to import and validate they exist
                For ictr As Int16 = numfields To tableData.Columns.Count - 1
                    fieldName = tableData.Columns(ictr).ColumnName
                    fieldName = fieldName.Trim
                    tableData.Columns(ictr).ColumnName = fieldName

                    ASCMAIN1.Progress("-", fieldName)

                    If Not fieldName.Contains("_") Then
                        MessageBox.Show("Style / Color field (" & fieldName & ") not in the proper format. Item wll be skipped.", "Import Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Continue For
                    End If

                    Dim STYLE_CODE As String = fieldName.Split("_")(0).Trim.ToUpper
                    Dim COLOR_CODE As String = fieldName.Split("_")(1).Trim.ToUpper

                    If LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE}) Is Nothing Then
                        MessageBox.Show("Style " & STYLE_CODE & " / Color " & COLOR_CODE & " field (" & fieldName & ") not in the item master. Item wll be skipped.", "Import Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Continue For
                    End If

                    lstItems.Add(fieldName)
                Next

            Else

                ASCMAIN1.Progress("Gathering Sales Orders")
                Me.Cursor = Cursors.WaitCursor

                For icol As Integer = 1 To numfields
                    Select Case icol
                        Case 1
                            fieldName = "CUST_ADDR_CODE"
                        Case 2
                            fieldName = "CUST_NAME"
                        Case 3
                            fieldName = "CUST_DISTRICT"
                        Case 4
                            fieldName = "CUST_REGION"
                        Case 5
                            fieldName = "CUST_BRAND"
                        Case 6
                            fieldName = "CUST_AREA_CODE"
                        Case 7
                            fieldName = "CUST_PHONE"
                        Case 8
                            fieldName = "CUST_ADDR1"
                        Case 9
                            fieldName = "CUST_ADDR2"
                        Case 10
                            fieldName = "CUST_ADDR3"
                        Case 11
                            fieldName = "CUST_CITY"
                        Case 12
                            fieldName = "CUST_STATE"
                        Case 13
                            fieldName = "CUST_ZIP_CODE"
                        Case 14
                            fieldName = "CUST_COUNTRY"

                        Case Else
                            ' should be the style code
                    End Select

                    tableData.Columns.Add(fieldName, GetType(System.String))
                Next

                Dim lstOrders As New List(Of String)
                For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Rows
                    lstOrders.Add(rowSOTORDR0.Item("ORDR_GROUP_NO"))
                Next

                If MyBase.Absx1.txtFor("BULK_DESC").TextLength = 0 Then
                    MyBase.Absx1.txtFor("BULK_DESC").Text = "Customer " & Absx1.txtFor("CUST_NAME").Text & ", PO: " & Absx1.txtFor("ORDR_CUST_PO").Text & ", Order Group: " & dst.Tables("SOTORDR0").Rows(0).Item("ORDR_GROUP_NO")
                    dst.Tables("SOTBULK1").Rows(0).Item("BULK_DESC") = "Customer " & Absx1.txtFor("CUST_NAME").Text & ", PO: " & Absx1.txtFor("ORDR_CUST_PO").Text & ", Order Group: " & dst.Tables("SOTORDR0").Rows(0).Item("ORDR_GROUP_NO")
                End If

                Dim temp_table As String = String.Empty

                Select Case ImportType
                    Case ImportTypes.FromReleasedSalesOrders
                        temp_table = ASCMAIN1.Temp_Table("Select SOTPICK1.SHIP_BOL_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_NO FROM SOTPICK1 WHERE SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM SOTSHIP1 WHERE ORDR_GROUP_NO = '" & HFs("ORDR_GROUP_NO") & "')")
                    Case ImportTypes.FromOpenSalesOrders
                        temp_table = ASCMAIN1.Temp_Table("Select ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO = '" & HFs("ORDR_GROUP_NO") & "' AND ORDR_STATUS IN ('O', 'P')")
                End Select

                ASCMAIN1.Progress("-", "SOTORDR5")
                ASCMAIN1.sql = "SELECT * FROM SOTORDR5 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & temp_table & ") AND CUST_ADDR_TYPE = 'ST'"
                Fill_Records("SOTORDR5", String.Empty, True, ASCMAIN1.sql)

                ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & temp_table & ")"
                Fill_Records("SOTORDR1", String.Empty, True, ASCMAIN1.sql)

                Dim sqlAvail As String = String.Empty
                Select Case ImportType
                    Case ImportTypes.FromReleasedSalesOrders
                        ASCMAIN1.sql = "SELECT SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                            & " FROM SOTPICK2, SOTORDR2" _
                            & " WHERE SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO" _
                            & " AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO" _
                            & " AND SOTPICK2.PICK_NO IN (SELECT PICK_NO FROM " & temp_table & ")"

                    Case ImportTypes.FromOpenSalesOrders
                        ASCMAIN1.sql = "SELECT SOTORDR2.*, ORDR_QTY_OPEN PICK_QTY" _
                            & " FROM SOTORDR2" _
                            & " WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & temp_table & ") AND ORDR_QTY_OPEN > 0"
                End Select

                Dim tblSOTPICK2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

                sqlAvail = " SELECT ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND,"
                sqlAvail &= " ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO, ICTSTAT2.WHSE_QTY_ON_ORDER,"
                sqlAvail &= " (ICTSTAT2.WHSE_QTY_ON_HAND - (ICTSTAT2.WHSE_QTY_PICK - SUM(NVL(SOTORDR2.ORDR_QTY_PICK, 0)))) QTY_AVAIL"
                sqlAvail &= " FROM ICTSTAT2, SOTORDR1, SOTORDR2"
                sqlAvail &= " WHERE ICTSTAT2.WHSE_CODE = SOTORDR1.WHSE_CODE"
                sqlAvail &= " AND ICTSTAT2.STYLE_CODE = SOTORDR2.STYLE_CODE"
                sqlAvail &= " AND ICTSTAT2.COLOR_CODE = SOTORDR2.COLOR_CODE"
                sqlAvail &= " AND SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
                sqlAvail &= " AND SOTORDR1.ORDR_NO in (SELECT ORDR_NO FROM " & temp_table & ")"
                sqlAvail &= " GROUP BY"
                sqlAvail &= " ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND,"
                sqlAvail &= " ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO, ICTSTAT2.WHSE_QTY_ON_ORDER"

                ASCDATA1.ExecuteSQL("INSERT INTO " & tblItemQty & " " & sqlAvail)
                availableLoaded = True

                Dim tblITEMS As DataTable = ASCDATA1.SelectDistinct(tblSOTPICK2, New String() {"STYLE_CODE", "COLOR_CODE"})
                For Each rowItems As DataRow In tblITEMS.Select("", "STYLE_CODE,COLOR_CODE")
                    fieldName = rowItems.Item("STYLE_CODE") & "_" & rowItems.Item("COLOR_CODE")
                    lstItems.Add(fieldName)
                    tableData.Columns.Add(fieldName, GetType(System.Int32))
                Next

                For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select()
                    Dim ORDR_NO As String = rowSOTORDR5.Item("ORDR_NO")
                    Dim rowTabledata As DataRow = tableData.NewRow

                    Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                    If rowSOTORDR1 IsNot Nothing Then
                        rowTabledata.Item("CUST_ADDR_CODE") = rowSOTORDR1.Item("CUST_STORE_NO")
                    Else
                        rowTabledata.Item("CUST_ADDR_CODE") = ORDR_NO
                    End If

                    rowTabledata.Item("CUST_NAME") = rowSOTORDR5.Item("CUST_NAME")
                    rowTabledata.Item("CUST_DISTRICT") = ORDR_NO
                    rowTabledata.Item("CUST_REGION") = rowSOTORDR5.Item("CUST_ADDR_CODE")
                    rowTabledata.Item("CUST_BRAND") = txtCUST_CODE.Text
                    rowTabledata.Item("CUST_AREA_CODE") = ""
                    rowTabledata.Item("CUST_PHONE") = rowSOTORDR5.Item("CUST_PHONE")
                    rowTabledata.Item("CUST_ADDR1") = rowSOTORDR5.Item("CUST_ADDR1")
                    rowTabledata.Item("CUST_ADDR2") = rowSOTORDR5.Item("CUST_ADDR2")
                    rowTabledata.Item("CUST_ADDR3") = rowSOTORDR5.Item("CUST_ADDR3")
                    rowTabledata.Item("CUST_CITY") = rowSOTORDR5.Item("CUST_CITY")
                    rowTabledata.Item("CUST_STATE") = rowSOTORDR5.Item("CUST_STATE")
                    rowTabledata.Item("CUST_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE")
                    rowTabledata.Item("CUST_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY")

                    tableData.Rows.Add(rowTabledata)

                    ' Get the ship quantities from the Pick Ticket
                    For Each rowSOTPICK2 As DataRow In tblSOTPICK2.Select("ORDR_NO = '" & ORDR_NO & "'")
                        fieldName = rowSOTPICK2.Item("STYLE_CODE") & "_" & rowSOTPICK2.Item("COLOR_CODE")
                        rowTabledata.Item(fieldName) = Val(rowTabledata.Item(fieldName) & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty)
                    Next
                Next
            End If

            EnforceConstraints(False)

            For Each TABLE_NAME As String In New String() {"SOTBULK2", "SOTBULK3", "SOTBULK4", "SOTBULK5", "SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK9", "ICTSTYC1", "SOTBULKI"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            Dim lstICTSTYC1 As New List(Of String)
            Dim lstICTSTYL1 As New List(Of String)

            For Each item As String In lstItems
                Dim STYLE_CODE As String = item.Split("_")(0).Trim.ToUpper
                Dim COLOR_CODE As String = item.Split("_")(1).Trim.ToUpper

                lstICTSTYC1.Add("('" & STYLE_CODE & "', '" & COLOR_CODE & "')")
                lstICTSTYL1.Add("'" & STYLE_CODE & "'")

                ' This prevents a SQL string to long Oracle Error
                If lstICTSTYC1.Count > 30 Then
                    ASCMAIN1.sql = "Select * from ICTSTYLD WHERE STYLE_CODE IN (" & String.Join(",", lstICTSTYL1.ToArray) & ")"
                    Fill_Records("ICTSTYLD", String.Empty, False, ASCMAIN1.sql)

                    ASCMAIN1.sql = "Select ICTSTYC1.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.SUB_UNIT_BAG_QTY, ICTSTYL1.SUB_UNIT_PACK_QTY, ICTSTYL1.STYLE_PRICE" _
                        & " FROM ICTSTYL1, ICTSTYC1" _
                        & " WHERE ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" _
                        & " AND (ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE) IN " _
                        & "(" & String.Join(",", lstICTSTYC1.ToArray) & ")"
                    Fill_Records("ICTSTYC1", String.Empty, False, ASCMAIN1.sql)

                    If Not availableLoaded Then
                        ASCMAIN1.sql = "SELECT ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND," _
                            & "  ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO, ICTSTAT2.WHSE_QTY_ON_ORDER," _
                            & "  (ICTSTAT2.WHSE_QTY_ON_HAND - ICTSTAT2.WHSE_QTY_PICK) QTY_AVAIL" _
                            & "  FROM ICTSTAT2 " _
                            & "  WHERE ICTSTAT2.WHSE_CODE = 'MS'" _
                            & "  AND (STYLE_CODE, COLOR_CODE) IN (" & String.Join(",", lstICTSTYC1.ToArray) & ")"
                        ASCMAIN1.sql = "INSERT INTO " & tblItemQty & " " & ASCMAIN1.sql
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                    End If

                    lstICTSTYC1.Clear()
                    lstICTSTYL1.Clear()
                End If
            Next

            If lstICTSTYC1.Count > 0 Then
                ASCMAIN1.sql = "Select * from ICTSTYLD WHERE STYLE_CODE IN (" & String.Join(",", lstICTSTYL1.ToArray) & ")"
                Fill_Records("ICTSTYLD", String.Empty, True, ASCMAIN1.sql)

                ASCMAIN1.sql = "Select ICTSTYC1.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.SUB_UNIT_BAG_QTY, ICTSTYL1.SUB_UNIT_PACK_QTY, ICTSTYL1.STYLE_PRICE, ICTSTYL1.COUNTRY_CODE" _
                    & " FROM ICTSTYL1, ICTSTYC1" _
                    & " WHERE ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE" _
                    & " AND (ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE) IN " _
                    & "(" & String.Join(",", lstICTSTYC1.ToArray) & ")"
                Fill_Records("ICTSTYC1", String.Empty, False, ASCMAIN1.sql)


                If Not availableLoaded Then
                    ASCMAIN1.sql = "SELECT ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE, ICTSTAT2.WHSE_QTY_ON_HAND," _
                        & "  ICTSTAT2.WHSE_QTY_PICK, ICTSTAT2.WHSE_QTY_ALLO, ICTSTAT2.WHSE_QTY_ON_ORDER," _
                        & "  (ICTSTAT2.WHSE_QTY_ON_HAND - ICTSTAT2.WHSE_QTY_PICK) QTY_AVAIL" _
                        & "  FROM ICTSTAT2 " _
                        & "  WHERE ICTSTAT2.WHSE_CODE = 'MS'" _
                        & "  AND (STYLE_CODE, COLOR_CODE) IN (" & String.Join(",", lstICTSTYC1.ToArray) & ")"
                    ASCMAIN1.sql = "INSERT INTO " & tblItemQty & " " & ASCMAIN1.sql
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                End If
            End If

            Dim tblICTSTAT2 As DataTable = ASCDATA1.GetDataTable("Select * from " & tblItemQty)
            tblICTSTAT2.PrimaryKey = New System.Data.DataColumn() {tblICTSTAT2.Columns("STYLE_CODE"), tblICTSTAT2.Columns("COLOR_CODE")}

            ' SOTBULKI
            For Each item As String In lstItems
                Dim STYLE_CODE As String = item.Split("_")(0).Trim.ToUpper
                Dim COLOR_CODE As String = item.Split("_")(1).Trim.ToUpper

                Dim rowSOTBULKI As DataRow = dst.Tables("SOTBULKI").NewRow
                rowSOTBULKI.Item("BULK_CODE") = HFs("BULK_CODE")
                rowSOTBULKI.Item("STYLE_CODE") = STYLE_CODE
                rowSOTBULKI.Item("COLOR_CODE") = COLOR_CODE

                Dim rowICTSTAT2 As DataRow = tblICTSTAT2.Rows.Find(New Object() {STYLE_CODE, COLOR_CODE})
                If rowICTSTAT2 IsNot Nothing Then
                    rowSOTBULKI.Item("WHSE_QTY_ON_HAND") = rowICTSTAT2.Item("WHSE_QTY_ON_HAND")
                    rowSOTBULKI.Item("WHSE_QTY_PICK") = rowICTSTAT2.Item("WHSE_QTY_PICK")
                    rowSOTBULKI.Item("QTY_AVAIL") = rowICTSTAT2.Item("QTY_AVAIL")
                    rowSOTBULKI.Item("WHSE_QTY_ON_ORDER") = rowICTSTAT2.Item("WHSE_QTY_ON_ORDER")
                End If

                Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New Object() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 IsNot Nothing Then
                    rowSOTBULKI.Item("CARTON_PACK_QTY") = rowICTSTYC1.Item("CARTON_PACK_QTY") & String.Empty
                    rowSOTBULKI.Item("INNER_PACK_QTY") = rowICTSTYC1.Item("INNER_PACK_QTY") & String.Empty
                    rowSOTBULKI.Item("STYLE_DESC") = rowICTSTYC1.Item("STYLE_DESC") & String.Empty
                    rowSOTBULKI.Item("STYLE_PRICE") = rowICTSTYC1.Item("STYLE_PRICE") & String.Empty
                    rowSOTBULKI.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE") & String.Empty
                    rowSOTBULKI.Item("COUNTRY_CODE") = rowICTSTYC1.Item("COUNTRY_CODE") & String.Empty
                End If

                Dim rowICTSTYLD As DataRow = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, "CTN"})
                If rowICTSTYLD IsNot Nothing Then
                    rowSOTBULKI.Item("CARTON_PKG_L") = rowICTSTYLD.Item("LENGTH")
                    rowSOTBULKI.Item("CARTON_PKG_W") = rowICTSTYLD.Item("WIDTH")
                    rowSOTBULKI.Item("CARTON_PKG_H") = rowICTSTYLD.Item("HEIGHT")
                    rowSOTBULKI.Item("CARTON_WEIGHT") = rowICTSTYLD.Item("WEIGHT")
                End If

                rowICTSTYLD = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, "INR"})
                If rowICTSTYLD IsNot Nothing Then
                    rowSOTBULKI.Item("INNER_PKG_L") = rowICTSTYLD.Item("LENGTH")
                    rowSOTBULKI.Item("INNER_PKG_W") = rowICTSTYLD.Item("WIDTH")
                    rowSOTBULKI.Item("INNER_PKG_H") = rowICTSTYLD.Item("HEIGHT")
                    rowSOTBULKI.Item("INNER_WEIGHT") = rowICTSTYLD.Item("WEIGHT")
                End If

                dst.Tables("SOTBULKI").Rows.Add(rowSOTBULKI)

            Next

            clearedTables = True

            For Each rowData As DataRow In tableData.Select("")
                Dim CUST_ADDR_CODE As String = rowData.Item("CUST_ADDR_CODE") & String.Empty
                ASCMAIN1.Progress("-", CUST_ADDR_CODE)

                If CUST_ADDR_CODE.Length = 0 Then
                    MessageBox.Show("Line is missing the Store Number. Row will be skipped.", "Import Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    rowData.Delete()
                    Continue For
                End If

                ' SOTBULK2
                Dim rowSOTBULK2 As DataRow = dst.Tables("SOTBULK2").NewRow
                For Each dColumn As DataColumn In rowSOTBULK2.Table.Columns
                    If rowData.Table.Columns.Contains(dColumn.ColumnName) Then
                        rowSOTBULK2.Item(dColumn.ColumnName) = (rowData.Item(dColumn.ColumnName)).ToString.Trim
                    End If
                Next

                rowSOTBULK2.Item("CUST_AREA_CODE") = (rowSOTBULK2.Item("CUST_AREA_CODE") & String.Empty).ToString.Replace(" ", "").Replace("(", "").Replace(")", "")
                rowSOTBULK2.Item("CUST_PHONE") = (rowSOTBULK2.Item("CUST_PHONE") & String.Empty).ToString.Replace(" ", "").Replace("(", "").Replace(")", "")

                Dim CUST_ZIP_CODE As String = rowSOTBULK2.Item("CUST_ZIP_CODE") & String.Empty
                Dim CUST_COUNTRY As String = (rowSOTBULK2.Item("CUST_COUNTRY") & String.Empty).ToString.ToUpper.Replace(".", "")
                If CUST_COUNTRY = String.Empty OrElse CUST_COUNTRY = "US" OrElse CUST_COUNTRY = "USA" Then
                    CUST_COUNTRY = "US"
                    rowSOTBULK2.Item("CUST_COUNTRY") = "US"

                    If CUST_ZIP_CODE.Length > 0 AndAlso CUST_ZIP_CODE.Length < 5 Then
                        CUST_ZIP_CODE = CUST_ZIP_CODE.PadLeft(5, "0")
                        rowSOTBULK2.Item("CUST_ZIP_CODE") = CUST_ZIP_CODE
                    End If
                End If

                Dim CUST_BRAND As String = rowSOTBULK2.Item("CUST_BRAND") & String.Empty
                If CUST_BRAND.Length = 0 Then
                    CUST_BRAND = "****"
                End If
                CUST_BRAND = CUST_BRAND.ToUpper
                rowSOTBULK2.Item("CUST_BRAND") = CUST_BRAND

                rowSOTBULK2.Item("BULK_CODE") = HFs("BULK_CODE")
                dst.Tables("SOTBULK2").Rows.Add(rowSOTBULK2)

                ' SOTBULK3
                For Each item As String In lstItems
                    Dim STYLE_CODE As String = item.Split("_")(0).Trim.ToUpper
                    Dim COLOR_CODE As String = item.Split("_")(1).Trim.ToUpper

                    ' Make sure the value is numeric
                    rowData.Item(item) = Val(rowData.Item(item) & String.Empty)

                    Dim rowSOTBULK3 As DataRow = dst.Tables("SOTBULK3").NewRow
                    rowSOTBULK3.Item("BULK_CODE") = HFs("BULK_CODE")
                    rowSOTBULK3.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowSOTBULK3.Item("STYLE_CODE") = STYLE_CODE
                    rowSOTBULK3.Item("COLOR_CODE") = COLOR_CODE
                    rowSOTBULK3.Item("UPC_CODE") = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("UPC_CODE") & String.Empty
                    rowSOTBULK3.Item("ORDR_QTY") = Val(rowData.Item(item) & String.Empty)
                    dst.Tables("SOTBULK3").Rows.Add(rowSOTBULK3)
                Next
            Next

            ' Create Patterns
            Dim Pattern As Int16 = 0
            Dim tblPattern As DataTable = ASCDATA1.SelectDistinct(tableData, lstItems.ToArray)
            For Each rowPattern As DataRow In tblPattern.Select("")
                Pattern += 1

                Dim rowSOTBULK4 As DataRow = dst.Tables("SOTBULK4").NewRow
                rowSOTBULK4.Item("BULK_CODE") = HFs("BULK_CODE")
                rowSOTBULK4.Item("BULK_PATTERN_NO") = Pattern
                rowSOTBULK4.Item("BULK_PATTERN_DESC") = "Pattern " & Pattern
                dst.Tables("SOTBULK4").Rows.Add(rowSOTBULK4)

                ASCMAIN1.Progress("-", "Pattern " & Pattern)

                For Each dColumn As DataColumn In tblPattern.Columns
                    Dim columnName As String = dColumn.ColumnName

                    Dim STYLE_CODE As String = columnName.Split("_")(0).Trim.ToUpper
                    Dim COLOR_CODE As String = columnName.Split("_")(1).Trim.ToUpper

                    Dim rowSOTBULK5 As DataRow = dst.Tables("SOTBULK5").NewRow
                    rowSOTBULK5.Item("BULK_CODE") = HFs("BULK_CODE")
                    rowSOTBULK5.Item("BULK_PATTERN_NO") = rowSOTBULK4.Item("BULK_PATTERN_NO")
                    rowSOTBULK5.Item("STYLE_CODE") = STYLE_CODE
                    rowSOTBULK5.Item("COLOR_CODE") = COLOR_CODE
                    rowSOTBULK5.Item("ORDR_QTY") = Val(rowPattern.Item(columnName) & String.Empty)
                    dst.Tables("SOTBULK5").Rows.Add(rowSOTBULK5)
                Next
            Next

            ExtendData()

            EnforceConstraints(True)

            clearedTables = False

            If ImportType = ImportTypes.FromFile Then
                dst.Tables("SOTBULK1").Rows(0).Item("IMPORT_TYPE") = "F"
            Else
                dst.Tables("SOTBULK1").Rows(0).Item("IMPORT_TYPE") = "O"
                dst.Tables("SOTBULK1").Rows(0).Item("ORDR_GROUP_NO") = dst.Tables("SOTORDR0").Rows(0).Item("ORDR_GROUP_NO")
                dst.Tables("SOTBULK1").Rows(0).Item("EDI_DOC_SEQ_NO") = dst.Tables("SOTORDR0").Rows(0).Item("EDI_DOC_SEQ_NO")
            End If

            Select Case dst.Tables("SOTBULK1").Rows(0).Item("IMPORT_TYPE") & String.Empty
                Case String.Empty
                    lblOrderType.Text = String.Empty

                Case "F"
                    lblOrderType.Text = "Data imported from a file"

                Case "O"
                    If dst.Tables("SOTBULK1").Rows(0).Item("EDI_DOC_SEQ_NO") & String.Empty <> String.Empty Then
                        lblOrderType.Text = "Imported EDI Order"
                    Else
                        lblOrderType.Text = "Imported Sales Order"
                    End If
            End Select

            ASCMAIN1.Progress("-", "Fmt Grids")
            grdSOTBULK2.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
            grdSOTBULK4.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
            grdSOTBULKI.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
            grdSOTBULK6.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
            grdSOTBULK9.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

            grdSOTBULK2_SHIP.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)
            grdSOTBULKI_SHIP.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        Catch ex As Exception
            MessageBox.Show("The following Error occurred while importing the file: " & ex.Message, "Import Data", MessageBoxButtons.OK, MessageBoxIcon.Error)

            If clearedTables Then
                For Each TABLE_NAME As String In New String() {"SOTBULK2", "SOTBULK3", "SOTBULK4", "SOTBULK5", "SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK9", "SOTBULKI", "ICTSTYC1"}
                    dst.Tables(TABLE_NAME).RejectChanges()
                Next

                EnforceConstraints(True)
            End If

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub ExtendData()

        ' Remove any extra columns from SOTBULK4
        'BULK_CODE, BULK_PATTERN_NO, BULK_PATTERN_DESC, NUM_ACCOUNTS NUM_CARTONS
        While dst.Tables("SOTBULK4").Columns.Count > 5
            For Each dcol As DataColumn In dst.Tables("SOTBULK4").Columns
                If Not "BULK_CODE,BULK_PATTERN_NO,BULK_PATTERN_DESC,NUM_ACCOUNTS,NUM_CARTONS".Contains(dcol.ColumnName) Then
                    dst.Tables("SOTBULK4").Columns.Remove(dcol.ColumnName)

                    For Each SS As Infragistics.Win.UltraWinGrid.SummarySettings In grdSOTBULK4.DisplayLayout.Bands(0).Summaries
                        If SS.Key = dcol.ColumnName Then
                            grdSOTBULK4.DisplayLayout.Bands(0).Summaries.Remove(SS)
                            Exit For
                        End If
                    Next

                    For Each SS As Infragistics.Win.UltraWinGrid.SummarySettings In grdSOTBULK6.DisplayLayout.Bands(0).Summaries
                        If SS.Key = dcol.ColumnName Then
                            grdSOTBULK6.DisplayLayout.Bands(0).Summaries.Remove(SS)
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        End While

        ASCMAIN1.Progress("-", "Bulk 1")
        Dim lstFields As New List(Of String)
        Dim tblSOTBULK3 As New DataTable
        tblSOTBULK3.Columns.Add("BULK_CODE", GetType(System.String))
        tblSOTBULK3.Columns.Add("CUST_ADDR_CODE", GetType(System.String))
        tblSOTBULK3.PrimaryKey = New System.Data.DataColumn() {tblSOTBULK3.Columns("BULK_CODE"), tblSOTBULK3.Columns("CUST_ADDR_CODE")}

        Dim tblStyleColor As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTBULK5"), New String() {"STYLE_CODE", "COLOR_CODE"})
        For Each row As DataRow In tblStyleColor.Select("", "STYLE_CODE, COLOR_CODE")
            Dim fieldName As String = row.Item("STYLE_CODE") & "_" & row.Item("COLOR_CODE")
            lstFields.Add(fieldName)
            dst.Tables("SOTBULK4").Columns.Add(fieldName, GetType(System.Int32))
            tblSOTBULK3.Columns.Add(fieldName, GetType(System.Int32))

            For Each grd As Infragistics.Win.UltraWinGrid.UltraGrid In New Infragistics.Win.UltraWinGrid.UltraGrid() {grdSOTBULK4, grdSOTBULK6}
                grd.DisplayLayout.Bands(0).Columns(fieldName).Hidden = False
                grd.DisplayLayout.Bands(0).Columns(fieldName).Header.Caption = fieldName.Replace("_", " ")
                grd.DisplayLayout.Bands(0).Columns(fieldName).CellActivation = UltraWinGrid.Activation.NoEdit
                grd.DisplayLayout.Bands(0).Columns(fieldName).CellAppearance.BackColor = Drawing.Color.LightBlue

                'If grd.Name = grdSOTBULK4.Name Then
                'Create_Summary(grdSOTBULK4, fieldName, "Custom", , "#,##0")
                'End If

                Create_Summary(grd, fieldName, "Custom", , "#,##0")
            Next
        Next

        ASCMAIN1.Progress("-", "Bulk 2")
        For Each row As DataRow In dst.Tables("SOTBULK5").Select("")
            Dim BULK_PATTERN_NO As Int16 = row.Item("BULK_PATTERN_NO")
            Dim rowSOTBULK4 As DataRow = dst.Tables("SOTBULK4").Rows.Find(New Object() {HFs("BULK_CODE"), BULK_PATTERN_NO})
            Dim fieldName As String = row.Item("STYLE_CODE") & "_" & row.Item("COLOR_CODE")
            rowSOTBULK4.Item(fieldName) = row.Item("ORDR_QTY")
        Next

        ASCMAIN1.Progress("-", "Bulk 3")
        For Each rowSOTBULK3 As DataRow In dst.Tables("SOTBULK3").Select("")
            Dim BULK_CODE As String = rowSOTBULK3.Item("BULK_CODE")
            Dim CUST_ADDR_CODE As String = rowSOTBULK3.Item("CUST_ADDR_CODE")
            Dim fieldName As String = rowSOTBULK3.Item("STYLE_CODE") & "_" & rowSOTBULK3.Item("COLOR_CODE")

            Dim rowtblSOTBULK3 As DataRow = tblSOTBULK3.Rows.Find(New Object() {BULK_CODE, CUST_ADDR_CODE})
            If rowtblSOTBULK3 Is Nothing Then
                tblSOTBULK3.Rows.Add(New Object() {BULK_CODE, CUST_ADDR_CODE})
                rowtblSOTBULK3 = tblSOTBULK3.Rows.Find(New Object() {BULK_CODE, CUST_ADDR_CODE})
            End If
            rowtblSOTBULK3.Item(fieldName) = rowSOTBULK3.Item("ORDR_QTY")
        Next

        ASCMAIN1.Progress("-", "Bulk 4")
        For Each row As DataRow In dst.Tables("SOTBULK4").Select("")
            Dim sql As String = String.Empty
            For Each item As String In lstFields
                sql &= " AND (" & item.Split("_")(0) & "_" & item.Split("_")(1) & " = " & row.Item(item) & ")"
            Next

            sql = sql.Substring(4).Trim

            For Each rowx As DataRow In tblSOTBULK3.Select(sql)
                Dim CUST_ADDR_CODE As String = rowx.Item("CUST_ADDR_CODE")
                dst.Tables("SOTBULK2").Select("CUST_ADDR_CODE = '" & CUST_ADDR_CODE & "'")(0).Item("BULK_PATTERN_NO") = row.Item("BULK_PATTERN_NO")
            Next
        Next

        Dim band As Int16 = 0
        For iCtr As Int16 = 0 To grdSOTBULK6.DisplayLayout.Bands.Count - 1
            If grdSOTBULK6.DisplayLayout.Bands(iCtr).Key = "SOTBULK6_SOTBULK7" Then
                band = iCtr
                Exit For
            End If
        Next

        Dim tblStyle As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTBULK3"), New String() {"STYLE_CODE"})
        Dim returnValue As List(Of String) = (From r In tblStyle.AsEnumerable() Select r.Field(Of String)(0)).ToList()
        ASCMAIN1.Add_Value_List(grdSOTBULK6, "STYLE_CODE", Nothing, Nothing, band, "SELECT STYLE_CODE, STYLE_CODE STYLE_DESC FROM ICTSTYL1 WHERE STYLE_CODE IN ('" & String.Join("', '", returnValue.ToArray) & "')")
        Fill_Records("ICTSTYL1", String.Empty, True, "SELECT * FROM ICTSTYL1 WHERE STYLE_CODE IN ('" & String.Join("', '", returnValue.ToArray) & "')")

        returnValue.Clear()
        Dim tblColor As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTBULK3"), New String() {"COLOR_CODE"})
        returnValue = (From r In tblColor.AsEnumerable() Select r.Field(Of String)(0)).ToList()
        ASCMAIN1.Add_Value_List(grdSOTBULK6, "COLOR_CODE", Nothing, Nothing, band, "SELECT COLOR_CODE, COLOR_CODE COLOR_DESC FROM ICTCOLR1 WHERE COLOR_CODE IN ('" & String.Join("', '", returnValue.ToArray) & "')")

        ' If no cartons then create one carton for each pattern
        ASCMAIN1.Progress("-", "Bulk 6")
        If dst.Tables("SOTBULK6").Rows.Count = 0 AndAlso dst.Tables("SOTBULK4").Rows.Count > 0 Then
            Cartonize(CartonizationMethods.OneCartonForAllItems)
        End If

        ASCMAIN1.Progress("-", "Bulk 9")

        Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTBULK2"), New String() {"CUST_BRAND"})
        Dim brandArray() As String = {"*", "All"}
        For Each row As DataRow In tbl.Select("")
            If row.Item("CUST_BRAND") & String.Empty = String.Empty Then
                row.Delete()
            End If
        Next
        tbl.AcceptChanges()

        If tbl.Select("").Length > 0 Then
            ReDim brandArray(tbl.Select("").Length)
            brandArray(0) = ":"

            Dim iLoop As Integer = 1
            For Each row As DataRow In tbl.Select("", "CUST_BRAND")
                brandArray(iLoop) = row.Item("CUST_BRAND") & ":" & row.Item("CUST_BRAND")
                iLoop += 1
            Next
        End If

        ASCMAIN1.Add_Value_List(grdSOTBULK9, "CUST_BRAND", , brandArray)

        ASCMAIN1.Progress("", "")

    End Sub

    Overrides Function CustomSummary_End( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByVal CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid) As Double

        Select Case grd.Name
            Case "grdSOTBULK4", "grdSOTBULK6"

                Select Case summarySettings.Key

                    Case "BULK_CODE", "BULK_PATTERN_NO", "BULK_PATTERN_DESC", "NUM_ACCOUNTS", "NUM_CARTONS"

                    Case Else
                        Dim totalPieces As Int32 = 0

                        For Each row As DataRow In dst.Tables("SOTBULK4").Select("")
                            totalPieces += Val(row.Item("NUM_ACCOUNTS") & String.Empty) * Val(row.Item(summarySettings.Key) & String.Empty)
                        Next

                        Return totalPieces
                End Select
        End Select
    End Function

    Private Sub ImportInstructions()
        Dim importInstr As String = String.Empty


        importInstr = "The file to be imported must be an Excel Workbook saved as Excel 97 - 2003 Workbook (*.xls)" & Environment.NewLine
        importInstr &= "The sheet name needs to be Sheet1" & Environment.NewLine
        importInstr &= "The sheet must have a header row." & Environment.NewLine
        importInstr &= "" & Environment.NewLine & Environment.NewLine


        'STORE NUMBER	STORE NAME	DISTRICT	REGION	BRAND	Area Code	Phone #	Address Line 1	Address Line 2	City	State	Zip Code	Country
        importInstr &= "The first 14 columns must appear in the following order:" & Environment.NewLine
        importInstr &= vbTab & "Store Number" & Environment.NewLine
        importInstr &= vbTab & "Store Name" & Environment.NewLine
        importInstr &= vbTab & "District" & Environment.NewLine
        importInstr &= vbTab & "Region" & Environment.NewLine
        importInstr &= vbTab & "Brand" & Environment.NewLine
        importInstr &= vbTab & "Area Code" & Environment.NewLine
        importInstr &= vbTab & "Telephone No" & Environment.NewLine
        importInstr &= vbTab & "Address Line 1" & Environment.NewLine
        importInstr &= vbTab & "Address Line 2" & Environment.NewLine
        importInstr &= vbTab & "Address Line 3" & Environment.NewLine
        importInstr &= vbTab & "City" & Environment.NewLine
        importInstr &= vbTab & "State" & Environment.NewLine
        importInstr &= vbTab & "Zip Code" & Environment.NewLine
        importInstr &= vbTab & "Country" & Environment.NewLine
        importInstr &= vbTab & "2 Character Country Code. Blank is assumed to be US" & Environment.NewLine

        importInstr &= Environment.NewLine
        importInstr &= Environment.NewLine
        importInstr &= "The remaining columns must be the Style Code and Color Code separated by an underscore (_). Example: MTB15461_SILV" & Environment.NewLine

        importInstr &= Environment.NewLine
        importInstr &= Environment.NewLine
        importInstr &= "Note: If the data is supplied by a customer, you need to clean up their data to appear in this format. REMOVE all hidden columms." & Environment.NewLine

        txtImportInstr.Text = importInstr
    End Sub

    Private Sub GetTransitTimes()
        Try

            Dim rList(1) As WHCSHIP1.RateList
            Dim shipTime As New Dictionary(Of Int16, Int16)

            Dim WHSE_CODE As String = MyBase.Absx1.txtFor("WHSE_CODE").Text
            If WHSE_CODE.Length = 0 Then
                MessageBox.Show("You must provide the Warehouse code.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
            If rowICTWHSE1 Is Nothing Then
                MessageBox.Show("Invalid Warehouse code.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
            Dim SHIP_VIA_CODE As String = String.Empty
            If dst.Tables("SOTBULK9").Rows.Count = 0 Then
                MessageBox.Show("You must create shipping methods.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim domesticShipViaCode As String = String.Empty
            Dim internationalShipViaCode As String = String.Empty


            If dst.Tables("SOTBULK9").Select("CARRIER_DI = 'B'").Length > 0 Then
                domesticShipViaCode = dst.Tables("SOTBULK9").Select("CARRIER_DI = 'B'")(0).Item("SHIP_VIA_CODE") & String.Empty
                internationalShipViaCode = domesticShipViaCode
            Else
                If dst.Tables("SOTBULK9").Select("CARRIER_DI = 'D'").Length > 0 Then
                    domesticShipViaCode = dst.Tables("SOTBULK9").Select("CARRIER_DI = 'D'")(0).Item("SHIP_VIA_CODE") & String.Empty
                End If

                If dst.Tables("SOTBULK9").Select("CARRIER_DI = 'I'").Length > 0 Then
                    internationalShipViaCode = dst.Tables("SOTBULK9").Select("CARRIER_DI = 'I'")(0).Item("SHIP_VIA_CODE") & String.Empty
                End If
            End If

            If domesticShipViaCode.Length > 0 Then
                SHIP_VIA_CODE = domesticShipViaCode
            ElseIf internationalShipViaCode.Length > 0 Then
                SHIP_VIA_CODE = internationalShipViaCode
            End If

                If SHIP_VIA_CODE.Length = 0 Then
                    MessageBox.Show("You must provide the Ship Via.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                If rowSOTSVIA1 Is Nothing Then
                    MessageBox.Show("Invalid Ship Via.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
                If CARRIER_CODE.Length = 0 Then
                    MessageBox.Show("Ship Via does not have an assigned Carrier.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'").Length = 0 Then
                    MessageBox.Show("Ship Via does not have a valid Carrier.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'").Length = 0 Then
                    MessageBox.Show("Ship Via has a Carrier that is not listed in the Carrier Master.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim rowSOTCARR1 As DataRow = dst.Tables("SOTCARR1").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
                Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)

                If CARRIER_CODE <> "UPS" Then
                    MessageBox.Show("Transit Times is supported only by UPS.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim CARRIER_PROD_CODE As String = rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty
                Dim CARRIER_PROD_CODE_INTL As String = rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty

                If CARRIER_PROD_CODE.Length = 0 Then
                    MessageBox.Show("The provided Ship Via does not have an assigned Carrier Product code.", "Transit Times", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim carrierRates As New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)
                Me.Cursor = Cursors.WaitCursor
                Dim CARRIER_MESSAGE_LENGTH As Int32 = dst.Tables("SOTBULK2").Columns("CARRIER_MESSAGE").MaxLength

                ' Credentials
                With carrierRates
                    .Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                    .UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
                    .Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                    .AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                    .UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                    .FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
                    .FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                    .LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim
                End With

                With carrierRates.Sender
                    .Company = (rowICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                    .FirstName = (rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                    .MiddleInitial = String.Empty
                    .LastName = String.Empty
                    .Address1 = (rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                    .Address2 = (rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                    .City = (rowICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                    .State = (rowICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                    .ZipCode = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                    .CountryCode = (rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                    If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                    If .CountryCode = "USA" Then .CountryCode = "US"
                    .Phone = (rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
                End With

                carrierRates.RequestedServiceType = CARRIER_PROD_CODE
                carrierRates.UPSPickupType = nsoftware.InShip.UpsratesPickupTypes.ptDailyPickup
                carrierRates.CustomerType = nsoftware.InShip.UpsratesCustomerTypes.ccRetail


                ' calculate the transit dyas starting with a Monday
                Dim shipDate As Date = DateTime.Now
                While shipDate.DayOfWeek <> DayOfWeek.Monday
                    shipDate = DateAdd(DateInterval.Day, 1, shipDate)
                End While

                carrierRates.ShipDate = shipDate.ToString("MM/dd/yyyy")

                carrierRates.ShipmentSpecialServices = 0
                carrierRates.SignatureRequired = False

                Dim listSeqNo As New List(Of Int16)

                Dim sql As String = String.Empty

                If dst.Tables("SOTBULK2").Select("ISNULL(TRANSIT_DAYS, 0) > 0").Length > 0 Then
                    Dim zmsg As String = "It appears you have already requested Transit Times. Do you want to request transit times for shipments with transit times errors?" & Environment.NewLine _
                                         & "Choose 'Yes' to request transit times for shipments with errors." & Environment.NewLine _
                                         & "Choose 'No' to request transit times for all shipments." & Environment.NewLine _
                                         & "Choose 'Cancel' to abort requesting transit times."

                    Select Case MessageBox.Show(zmsg, "Transit Times", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)

                        Case Windows.Forms.DialogResult.No
                            sql = String.Empty

                        Case Windows.Forms.DialogResult.Yes
                            sql = "ISNULL(TRANSIT_DAYS, 0) <= 0"

                        Case Windows.Forms.DialogResult.Cancel
                            Exit Sub

                    End Select
                End If


                ASCMAIN1.Progress("Requesting Transit Times", "")

                For Each rowSOTBULK2 As DataRow In dst.Tables("SOTBULK2").Select(sql, "CUST_COUNTRY DESC,CUST_STATE,CUST_ZIP_CODE")
                    Dim CUST_STATE As String = rowSOTBULK2.Item("CUST_STATE") & String.Empty
                    Dim CUST_ZIP_CODE As String = rowSOTBULK2.Item("CUST_ZIP_CODE") & String.Empty
                    Dim CUST_COUNTRY As String = rowSOTBULK2.Item("CUST_COUNTRY") & String.Empty
                    Dim TRANSIT_DAYS As Int16 = Val(rowSOTBULK2.Item("TRANSIT_DAYS") & String.Empty)

                If TRANSIT_DAYS > 0 Then
                    Continue For
                End If

                If CUST_STATE.Length = 0 OrElse CUST_ZIP_CODE.Length = 0 OrElse CUST_COUNTRY.Length = 0 Then
                    rowSOTBULK2.Item("TRANSIT_DAYS") = -1
                    Continue For
                End If

                CUST_COUNTRY = CUST_COUNTRY.ToUpper
                If CUST_COUNTRY = "USA" Then
                    CUST_COUNTRY = "US"
                End If

                If CUST_COUNTRY = "US" AndAlso CUST_ZIP_CODE.Length < 5 Then
                    CUST_ZIP_CODE = CUST_ZIP_CODE.PadLeft(5, "0")
                End If

                CUST_ZIP_CODE = CUST_ZIP_CODE.Replace(" ", "")

                Dim BULK_CODE As String = rowSOTBULK2.Item("BULK_CODE") & String.Empty
                Dim CUST_ADDR_CODE As String = rowSOTBULK2.Item("CUST_ADDR_CODE") & String.Empty
                Dim BULK_PATTERN_NO As String = rowSOTBULK2.Item("BULK_PATTERN_NO") & String.Empty

                ASCMAIN1.Progress("-", CUST_ADDR_CODE)

                carrierRates.PackageDetailList.Clear()
                Dim RatesTotalValue As Decimal = 0


                ' If no cartons then do not do anything 
                If dst.Tables("SOTBULK6").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'").Length = 0 Then
                    Continue For
                End If

                For Each rowSOTBULK6 As DataRow In dst.Tables("SOTBULK6").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")

                    Dim pkgDetail As New nsoftware.InShip.PackageDetail

                    Dim id As String = rowSOTBULK6.Item("BULK_PATTERN_NO") & "_" & rowSOTBULK6.Item("CART_NO")
                    id = id.PadLeft(8, "0")

                    pkgDetail.Id = id
                    pkgDetail.Weight = Val(rowSOTBULK6.Item("PKG_WEIGHT") & String.Empty)
                    pkgDetail.Weight *= 16

                    If pkgDetail.Weight = "0" Then
                        pkgDetail.Weight = "16.0"
                    End If

                    pkgDetail.PackagingType = CType(Val(rowSOTBULK6.Item("PACKAGING_TYPE") & String.Empty), nsoftware.InShip.UpsratesPickupTypes)
                    pkgDetail.Length = Val(rowSOTBULK6.Item("PKG_L") & String.Empty)
                    pkgDetail.Width = Val(rowSOTBULK6.Item("PKG_W") & String.Empty)
                    pkgDetail.Height = Val(rowSOTBULK6.Item("PKG_H") & String.Empty)

                    carrierRates.PackageDetailList.Add(pkgDetail)

                    Dim CART_NO As Int16 = rowSOTBULK6.Item("CART_NO") & String.Empty
                    For Each rowSOTBULK7 As DataRow In dst.Tables("SOTBULK7").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND CART_NO = " & CART_NO)
                        Dim STYLE_CODE As String = rowSOTBULK7.Item("STYLE_CODE") & String.Empty
                        Dim COLOR_CODE As String = rowSOTBULK7.Item("COLOR_CODE") & String.Empty
                        Dim QTY_PACKED As Int32 = Val(rowSOTBULK7.Item("QTY_PACKED") & String.Empty)

                        Dim STYLE_PRICE As Decimal = Val(dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("STYLE_PRICE") & String.Empty)

                        RatesTotalValue += STYLE_PRICE * QTY_PACKED
                    Next
                Next

                With carrierRates.Recipient
                    .FirstName = rowSOTBULK2.Item("CUST_NAME") & String.Empty
                    .MiddleInitial = ""
                    .LastName = ""

                    .Address1 = rowSOTBULK2.Item("CUST_ADDR1") & String.Empty
                    .Address2 = rowSOTBULK2.Item("CUST_ADDR2") & String.Empty
                    .City = rowSOTBULK2.Item("CUST_CITY") & String.Empty
                    .State = CUST_STATE
                    .ZipCode = CUST_ZIP_CODE
                    .CountryCode = CUST_COUNTRY
                    If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                    If .CountryCode = "USA" Then .CountryCode = "US"

                    .Company = rowSOTBULK2.Item("CUST_NAME") & String.Empty
                    .Phone = rowSOTBULK2.Item("CUST_AREA_CODE") & String.Empty & rowSOTBULK2.Item("CUST_PHONE") & String.Empty

                    If .Phone.Trim.Length = 0 Then
                        .Phone = carrierRates.Sender.Phone
                    End If

                    If .Phone.Trim.Length = 0 Then
                        .Phone = (rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
                    End If

                    .IsResidental = False
                    .IsPOBox = False
                End With

                If carrierRates.Recipient.CountryCode = "US" Then
                    carrierRates.RequestedServiceType = CARRIER_PROD_CODE
                ElseIf CARRIER_PROD_CODE_INTL.Length > 0 Then
                    CARRIER_PROD_CODE = CARRIER_PROD_CODE_INTL
                    carrierRates.RequestedServiceType = CARRIER_PROD_CODE
                End If

                carrierRates.RatesTotalValue = RatesTotalValue

                Me.Cursor = Cursors.WaitCursor

                rowSOTBULK2.Item("CARRIER_MESSAGE") = ""
                rList = carrierRates.GetUPSRatesList

                Me.Cursor = Cursors.WaitCursor

                If carrierRates.LastError.Length > 0 Then
                    rowSOTBULK2.Item("TRANSIT_DAYS") = -1
                    Dim CARRIER_MESSAGE As String = carrierRates.LastError
                    If CARRIER_MESSAGE.Length > CARRIER_MESSAGE_LENGTH Then
                        CARRIER_MESSAGE = CARRIER_MESSAGE.Substring(0, CARRIER_MESSAGE_LENGTH)
                    End If
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = CARRIER_MESSAGE
                    Continue For
                End If

                If rList Is Nothing Then
                    rowSOTBULK2.Item("TRANSIT_DAYS") = -2
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = "Carrier did not return a list of available services"
                    Continue For
                End If

                ' Default Value
                rowSOTBULK2.Item("TRANSIT_DAYS") = -3
                rowSOTBULK2.Item("CARRIER_MESSAGE") = "Carrier did not return transit days for the provided shipping service"

                For iLoop As Integer = 0 To rList.Count - 1
                    With rList(iLoop)
                        If Not (.ServiceType Is Nothing OrElse (.ServiceType = 0 AndAlso .ServiceTypeDescription.Length = 0)) Then
                            If .ServiceType = CARRIER_PROD_CODE Then
                                If Val(.TransitTime & String.Empty) > 0 Then
                                    rowSOTBULK2.Item("TRANSIT_DAYS") = .TransitTime
                                    rowSOTBULK2.Item("CARRIER_MESSAGE") = ""
                                ElseIf .DeliveryDate.Length > 0 Then
                                    Dim DeliveryDate As String = .DeliveryDate
                                    DeliveryDate = DeliveryDate.Split("-")(1) & "/" & DeliveryDate.Split("-")(2) & "/" & DeliveryDate.Split("-")(0)
                                    If IsDate(DeliveryDate) Then
                                        Dim numdays As Int16 = DateDiff(DateInterval.Day, CDate(shipDate.ToShortDateString), CDate(DeliveryDate))
                                        rowSOTBULK2.Item("TRANSIT_DAYS") = numdays
                                        rowSOTBULK2.Item("CARRIER_MESSAGE") = ""
                                    End If
                                End If
                                Exit For
                            End If
                        End If
                    End With
                Next
            Next

                MessageBox.Show("Getting Carrier Rates complete", "Get Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            MessageBox.Show("The following error occurred getting Carrier Rates: " & ex.Message, "Get Rates", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Function GetBusinessDays(ByVal startDate As DateTime, ByVal numdays As Int16) As Int16

        Dim weekdayDays As Int16 = 0
        Dim weekendDays As Int16 = 0

        Select Case numdays

            Case 0
                Return 0

            Case Is < 0
                Do While weekdayDays > numdays
                    If startDate.DayOfWeek = DayOfWeek.Saturday Or startDate.DayOfWeek = DayOfWeek.Sunday Then
                        weekendDays += 1
                    Else
                        weekdayDays -= 1
                    End If
                    startDate = DateAdd(DateInterval.Day, -1, startDate)
                Loop

            Case Is > 0
                Do While weekdayDays < numdays
                    If startDate.DayOfWeek = DayOfWeek.Saturday Or startDate.DayOfWeek = DayOfWeek.Sunday Then
                        weekendDays += 1
                    Else
                        weekdayDays += 1
                    End If
                    startDate = DateAdd(DateInterval.Day, 1, startDate)
                Loop
        End Select

        Return Math.Abs(weekdayDays) + weekendDays

    End Function

    Private Sub VerifyPatternCartonContents(ByVal BULK_PATTERN_NO As String, ByVal verifyNonZeroWeight As Boolean)

        Try
            Dim errors As String = String.Empty

            Dim sql As String = String.Empty
            If BULK_PATTERN_NO.Length > 0 Then
                sql = "BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'"
            End If

            For Each rowSOTBULK4 As DataRow In dst.Tables("SOTBULK4").Select(sql, "BULK_PATTERN_DESC")
                BULK_PATTERN_NO = rowSOTBULK4.Item("BULK_PATTERN_NO")
                Dim BULK_PATTERN_DESC As String = rowSOTBULK4.Item("BULK_PATTERN_DESC")

                For Each rowSOTBULK5 As DataRow In dst.Tables("SOTBULK5").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")
                    Dim STYLE_CODE As String = rowSOTBULK5.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowSOTBULK5.Item("COLOR_CODE")

                    ASCMAIN1.sql = "BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"
                    Dim QTY_PACKED As Int32 = Val(dst.Tables("SOTBULK7").Compute("SUM(QTY_PACKED)", ASCMAIN1.sql) & String.Empty)

                    If QTY_PACKED <> Val(rowSOTBULK5.Item("ORDR_QTY") & String.Empty) Then
                        EMsg &= vbCr & "Pattern (" & BULK_PATTERN_DESC & ") requies " & Val(rowSOTBULK5.Item("ORDR_QTY") & String.Empty) _
                            & " of Style: " & STYLE_CODE & ", Color: " & COLOR_CODE & "; however, the cartons contain only " & QTY_PACKED & " pieces."
                    End If
                Next

                If verifyNonZeroWeight Then
                    ASCMAIN1.sql = "BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' and ISNULL(PKG_WEIGHT, 0) <= 0"
                    If dst.Tables("SOTBULK6").Select(ASCMAIN1.sql, "").Length > 0 Then
                        EMsg &= vbCr & "Pattern (" & BULK_PATTERN_DESC & ") has cartons with Zero weights."
                    End If
                End If
            Next

        Catch ex As Exception
            EMsg &= vbCr & "Error in Verify Pattern Carton Contents: " & ex.Message
        End Try

    End Sub

    Private Sub Cartonize(ByVal CartonizeType As CartonizationMethods)

        Dim relaxConstraints As Boolean = False

        Try
            relaxConstraints = clsASCBASE1.dst.EnforceConstraints

            If relaxConstraints Then EnforceConstraints(False)

            For Each tableName As String In New String() {"SOTBULK6", "SOTBULK7", "SOTBULK8", "SOTBULK267"}
                dst.Tables(tableName).AcceptChanges()
                dst.Tables(tableName).Rows.Clear()
            Next

            Dim lstFields As New List(Of String)
            Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTBULKI"), New String() {"STYLE_CODE", "COLOR_CODE"})
            For Each row As DataRow In tbl.Select("", "STYLE_CODE, COLOR_CODE")
                Dim fieldName As String = row.Item("STYLE_CODE") & "_" & row.Item("COLOR_CODE")
                lstFields.Add(fieldName)
            Next

            Dim CART_NO As Int16 = 0

            Select Case CartonizeType

                Case CartonizationMethods.OneCartonForAllItems
                    ' One carton for all items
                    For Each rowSOTBULK4 As DataRow In dst.Tables("SOTBULK4").Select("", "BULK_PATTERN_NO")
                        Dim rowSOTBULK6 As DataRow = dst.Tables("SOTBULK6").NewRow
                        rowSOTBULK6.Item("BULK_CODE") = rowSOTBULK4.Item("BULK_CODE")

                        Dim BULK_PATTERN_NO As String = rowSOTBULK4.Item("BULK_PATTERN_NO")
                        rowSOTBULK6.Item("BULK_PATTERN_NO") = BULK_PATTERN_NO

                        CART_NO += 1
                        rowSOTBULK6.Item("CART_NO") = CART_NO.ToString.PadLeft(10, "0")
                        rowSOTBULK6.Item("PACKAGING_TYPE") = "31"
                        rowSOTBULK6.Item("PKG_CODE") = "OTHER"
                        rowSOTBULK6.Item("PKG_L") = 12
                        rowSOTBULK6.Item("PKG_W") = 12
                        rowSOTBULK6.Item("PKG_H") = 12
                        rowSOTBULK6.Item("PKG_WEIGHT") = 0
                        rowSOTBULK6.Item("REFERENCE_CODE1") = String.Empty
                        rowSOTBULK6.Item("REFERENCE_CODE1_VALUE") = String.Empty
                        rowSOTBULK6.Item("REFERENCE_CODE2") = String.Empty
                        rowSOTBULK6.Item("REFERENCE_CODE2_VALUE") = String.Empty
                        rowSOTBULK6.Item("REFERENCE_CODE3") = String.Empty
                        rowSOTBULK6.Item("REFERENCE_CODE3_VALUE") = String.Empty
                        dst.Tables("SOTBULK6").Rows.Add(rowSOTBULK6)

                        Dim CART_LNO As Int16 = 0
                        For Each fieldName As String In lstFields
                            Dim QTY_PACKED As Int32 = Val(dst.Tables("SOTBULK4").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")(0).Item(fieldName) & String.Empty)
                            If QTY_PACKED > 0 Then
                                Dim rowSOTBULK7 As DataRow = dst.Tables("SOTBULK7").NewRow
                                rowSOTBULK7.Item("BULK_CODE") = rowSOTBULK6.Item("BULK_CODE")
                                rowSOTBULK7.Item("BULK_PATTERN_NO") = rowSOTBULK6.Item("BULK_PATTERN_NO")
                                rowSOTBULK7.Item("CART_NO") = rowSOTBULK6.Item("CART_NO")
                                CART_LNO += 1
                                rowSOTBULK7.Item("CART_LNO") = CART_LNO
                                rowSOTBULK7.Item("STYLE_CODE") = fieldName.Split("_")(0)
                                rowSOTBULK7.Item("COLOR_CODE") = fieldName.Split("_")(1)
                                rowSOTBULK7.Item("UPC_CODE") = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & fieldName.Split("_")(0) & "' and COLOR_CODE = '" & fieldName.Split("_")(1) & "'")(0).Item("UPC_CODE") & String.Empty
                                rowSOTBULK7.Item("QTY_PACKED") = QTY_PACKED
                                rowSOTBULK7.Item("PACKED_CODE") = "Each"
                                dst.Tables("SOTBULK7").Rows.Add(rowSOTBULK7)
                            End If
                        Next
                    Next

                Case CartonizationMethods.OneCartonPerItem
                    ' One Style/Color combination per carton
                    For Each rowSOTBULK4 As DataRow In dst.Tables("SOTBULK4").Select("", "BULK_PATTERN_NO")
                        Dim BULK_PATTERN_NO As String = rowSOTBULK4.Item("BULK_PATTERN_NO")

                        For Each fieldName As String In lstFields
                            Dim QTY_PACKED As Int32 = Val(dst.Tables("SOTBULK4").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")(0).Item(fieldName) & String.Empty)
                            If QTY_PACKED > 0 Then

                                Dim rowSOTBULK6 As DataRow = dst.Tables("SOTBULK6").NewRow
                                rowSOTBULK6.Item("BULK_CODE") = rowSOTBULK4.Item("BULK_CODE")

                                rowSOTBULK6.Item("BULK_PATTERN_NO") = BULK_PATTERN_NO

                                CART_NO += 1
                                rowSOTBULK6.Item("CART_NO") = CART_NO.ToString.PadLeft(10, "0")
                                rowSOTBULK6.Item("PACKAGING_TYPE") = "31"
                                rowSOTBULK6.Item("PKG_CODE") = "OTHER"
                                rowSOTBULK6.Item("PKG_L") = 12
                                rowSOTBULK6.Item("PKG_W") = 12
                                rowSOTBULK6.Item("PKG_H") = 12
                                rowSOTBULK6.Item("PKG_WEIGHT") = 0
                                rowSOTBULK6.Item("REFERENCE_CODE1") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE1_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3_VALUE") = String.Empty
                                dst.Tables("SOTBULK6").Rows.Add(rowSOTBULK6)

                                Dim rowSOTBULK7 As DataRow = dst.Tables("SOTBULK7").NewRow
                                rowSOTBULK7.Item("BULK_CODE") = rowSOTBULK6.Item("BULK_CODE")
                                rowSOTBULK7.Item("BULK_PATTERN_NO") = rowSOTBULK6.Item("BULK_PATTERN_NO")
                                rowSOTBULK7.Item("CART_NO") = rowSOTBULK6.Item("CART_NO")
                                rowSOTBULK7.Item("CART_LNO") = 1
                                rowSOTBULK7.Item("STYLE_CODE") = fieldName.Split("_")(0)
                                rowSOTBULK7.Item("COLOR_CODE") = fieldName.Split("_")(1)
                                rowSOTBULK7.Item("UPC_CODE") = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & fieldName.Split("_")(0) & "' and COLOR_CODE = '" & fieldName.Split("_")(1) & "'")(0).Item("UPC_CODE") & String.Empty
                                rowSOTBULK7.Item("QTY_PACKED") = QTY_PACKED
                                rowSOTBULK7.Item("PACKED_CODE") = "Each"
                                dst.Tables("SOTBULK7").Rows.Add(rowSOTBULK7)
                            End If
                        Next
                    Next


                Case CartonizationMethods.UseCartonInnerDefinitions
                    ' Use Style/Color Carton Inner defintions
                    ' If not definititon then one carton per item
                    For Each rowSOTBULK4 As DataRow In dst.Tables("SOTBULK4").Select("", "BULK_PATTERN_NO")
                        Dim BULK_PATTERN_NO As String = rowSOTBULK4.Item("BULK_PATTERN_NO")

                        For Each rowSOTBULK5 As DataRow In dst.Tables("SOTBULK5").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' and ORDR_QTY > 0")
                            Dim ORDR_QTY As Int32 = Val(rowSOTBULK5.Item("ORDR_QTY") & String.Empty)
                            Dim STYLE_CODE As String = rowSOTBULK5.Item("STYLE_CODE") & String.Empty
                            Dim COLOR_CODE As String = rowSOTBULK5.Item("COLOR_CODE") & String.Empty

                            Dim rowSOTBULKI As DataRow = dst.Tables("SOTBULKI").Rows.Find(New Object() {HFs("BULK_CODE"), STYLE_CODE, COLOR_CODE})
                            If rowSOTBULKI Is Nothing Then
                                ' Place into one carton
                                Dim rowSOTBULK6 As DataRow = dst.Tables("SOTBULK6").NewRow
                                rowSOTBULK6.Item("BULK_CODE") = rowSOTBULK4.Item("BULK_CODE")
                                rowSOTBULK6.Item("BULK_PATTERN_NO") = BULK_PATTERN_NO

                                CART_NO += 1
                                rowSOTBULK6.Item("CART_NO") = CART_NO.ToString.PadLeft(10, "0")
                                rowSOTBULK6.Item("PACKAGING_TYPE") = "31"
                                rowSOTBULK6.Item("PKG_CODE") = "OTHER"
                                rowSOTBULK6.Item("PKG_L") = 12
                                rowSOTBULK6.Item("PKG_W") = 12
                                rowSOTBULK6.Item("PKG_H") = 12
                                rowSOTBULK6.Item("PKG_WEIGHT") = 0
                                rowSOTBULK6.Item("REFERENCE_CODE1") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE1_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3_VALUE") = String.Empty
                                dst.Tables("SOTBULK6").Rows.Add(rowSOTBULK6)

                                Dim rowSOTBULK7 As DataRow = dst.Tables("SOTBULK7").NewRow
                                rowSOTBULK7.Item("BULK_CODE") = rowSOTBULK6.Item("BULK_CODE")
                                rowSOTBULK7.Item("BULK_PATTERN_NO") = rowSOTBULK6.Item("BULK_PATTERN_NO")
                                rowSOTBULK7.Item("CART_NO") = rowSOTBULK6.Item("CART_NO")
                                rowSOTBULK7.Item("CART_LNO") = 1
                                rowSOTBULK7.Item("STYLE_CODE") = STYLE_CODE
                                rowSOTBULK7.Item("COLOR_CODE") = COLOR_CODE
                                rowSOTBULK7.Item("UPC_CODE") = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("UPC_CODE") & String.Empty
                                rowSOTBULK7.Item("QTY_PACKED") = ORDR_QTY
                                rowSOTBULK7.Item("PACKED_CODE") = "Each"
                                dst.Tables("SOTBULK7").Rows.Add(rowSOTBULK7)
                                Continue For
                            End If

                            Dim CARTON_PACK_QTY As Int32 = Val(rowSOTBULKI.Item("CARTON_PACK_QTY") & String.Empty)
                            Dim INNER_PACK_QTY As Int32 = Val(rowSOTBULKI.Item("INNER_PACK_QTY") & String.Empty)

                            ' Fill as many full cartons as possible
                            While ORDR_QTY >= CARTON_PACK_QTY And CARTON_PACK_QTY > 0
                                Dim rowSOTBULK6 As DataRow = dst.Tables("SOTBULK6").NewRow
                                rowSOTBULK6.Item("BULK_CODE") = rowSOTBULK4.Item("BULK_CODE")
                                rowSOTBULK6.Item("BULK_PATTERN_NO") = BULK_PATTERN_NO

                                CART_NO += 1
                                rowSOTBULK6.Item("CART_NO") = CART_NO.ToString.PadLeft(10, "0")
                                rowSOTBULK6.Item("PACKAGING_TYPE") = "31"
                                rowSOTBULK6.Item("PKG_CODE") = "OTHER"
                                rowSOTBULK6.Item("PKG_L") = Val(rowSOTBULKI.Item("CARTON_PKG_L") & String.Empty)
                                rowSOTBULK6.Item("PKG_W") = Val(rowSOTBULKI.Item("CARTON_PKG_W") & String.Empty)
                                rowSOTBULK6.Item("PKG_H") = Val(rowSOTBULKI.Item("CARTON_PKG_H") & String.Empty)
                                rowSOTBULK6.Item("PKG_WEIGHT") = Val(rowSOTBULKI.Item("CARTON_WEIGHT") & String.Empty)
                                rowSOTBULK6.Item("REFERENCE_CODE1") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE1_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3_VALUE") = String.Empty
                                dst.Tables("SOTBULK6").Rows.Add(rowSOTBULK6)

                                Dim rowSOTBULK7 As DataRow = dst.Tables("SOTBULK7").NewRow
                                rowSOTBULK7.Item("BULK_CODE") = rowSOTBULK6.Item("BULK_CODE")
                                rowSOTBULK7.Item("BULK_PATTERN_NO") = rowSOTBULK6.Item("BULK_PATTERN_NO")
                                rowSOTBULK7.Item("CART_NO") = rowSOTBULK6.Item("CART_NO")
                                rowSOTBULK7.Item("CART_LNO") = 1
                                rowSOTBULK7.Item("STYLE_CODE") = STYLE_CODE
                                rowSOTBULK7.Item("COLOR_CODE") = COLOR_CODE
                                rowSOTBULK7.Item("UPC_CODE") = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("UPC_CODE") & String.Empty
                                rowSOTBULK7.Item("QTY_PACKED") = CARTON_PACK_QTY
                                rowSOTBULK7.Item("PACKED_CODE") = "Carton"
                                dst.Tables("SOTBULK7").Rows.Add(rowSOTBULK7)

                                ORDR_QTY -= CARTON_PACK_QTY
                            End While

                            While ORDR_QTY >= INNER_PACK_QTY And INNER_PACK_QTY > 0
                                Dim rowSOTBULK6 As DataRow = dst.Tables("SOTBULK6").NewRow
                                rowSOTBULK6.Item("BULK_CODE") = rowSOTBULK4.Item("BULK_CODE")
                                rowSOTBULK6.Item("BULK_PATTERN_NO") = BULK_PATTERN_NO

                                CART_NO += 1
                                rowSOTBULK6.Item("CART_NO") = CART_NO.ToString.PadLeft(10, "0")
                                rowSOTBULK6.Item("PACKAGING_TYPE") = "31"
                                rowSOTBULK6.Item("PKG_CODE") = "OTHER"
                                rowSOTBULK6.Item("PKG_L") = Val(rowSOTBULKI.Item("INNER_PKG_L") & String.Empty)
                                rowSOTBULK6.Item("PKG_W") = Val(rowSOTBULKI.Item("INNER_PKG_W") & String.Empty)
                                rowSOTBULK6.Item("PKG_H") = Val(rowSOTBULKI.Item("INNER_PKG_H") & String.Empty)
                                rowSOTBULK6.Item("PKG_WEIGHT") = Val(rowSOTBULKI.Item("INNER_WEIGHT") & String.Empty)
                                rowSOTBULK6.Item("REFERENCE_CODE1") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE1_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3_VALUE") = String.Empty
                                dst.Tables("SOTBULK6").Rows.Add(rowSOTBULK6)

                                Dim rowSOTBULK7 As DataRow = dst.Tables("SOTBULK7").NewRow
                                rowSOTBULK7.Item("BULK_CODE") = rowSOTBULK6.Item("BULK_CODE")
                                rowSOTBULK7.Item("BULK_PATTERN_NO") = rowSOTBULK6.Item("BULK_PATTERN_NO")
                                rowSOTBULK7.Item("CART_NO") = rowSOTBULK6.Item("CART_NO")
                                rowSOTBULK7.Item("CART_LNO") = 1
                                rowSOTBULK7.Item("STYLE_CODE") = STYLE_CODE
                                rowSOTBULK7.Item("COLOR_CODE") = COLOR_CODE
                                rowSOTBULK7.Item("UPC_CODE") = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("UPC_CODE") & String.Empty
                                rowSOTBULK7.Item("QTY_PACKED") = INNER_PACK_QTY
                                rowSOTBULK7.Item("PACKED_CODE") = "Inner"
                                dst.Tables("SOTBULK7").Rows.Add(rowSOTBULK7)

                                ORDR_QTY -= INNER_PACK_QTY
                            End While

                            ' Place reamining quantities inone carton 
                            If ORDR_QTY > 0 Then
                                Dim rowSOTBULK6 As DataRow = dst.Tables("SOTBULK6").NewRow
                                rowSOTBULK6.Item("BULK_CODE") = rowSOTBULK4.Item("BULK_CODE")

                                rowSOTBULK6.Item("BULK_PATTERN_NO") = BULK_PATTERN_NO

                                CART_NO += 1
                                rowSOTBULK6.Item("CART_NO") = CART_NO.ToString.PadLeft(10, "0")
                                rowSOTBULK6.Item("PACKAGING_TYPE") = "31"
                                rowSOTBULK6.Item("PKG_CODE") = "OTHER"
                                rowSOTBULK6.Item("PKG_L") = 12
                                rowSOTBULK6.Item("PKG_W") = 12
                                rowSOTBULK6.Item("PKG_H") = 12
                                rowSOTBULK6.Item("PKG_WEIGHT") = 0
                                rowSOTBULK6.Item("REFERENCE_CODE1") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE1_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE2_VALUE") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3") = String.Empty
                                rowSOTBULK6.Item("REFERENCE_CODE3_VALUE") = String.Empty
                                dst.Tables("SOTBULK6").Rows.Add(rowSOTBULK6)

                                Dim rowSOTBULK7 As DataRow = dst.Tables("SOTBULK7").NewRow
                                rowSOTBULK7.Item("BULK_CODE") = rowSOTBULK6.Item("BULK_CODE")
                                rowSOTBULK7.Item("BULK_PATTERN_NO") = rowSOTBULK6.Item("BULK_PATTERN_NO")
                                rowSOTBULK7.Item("CART_NO") = rowSOTBULK6.Item("CART_NO")
                                rowSOTBULK7.Item("CART_LNO") = 1
                                rowSOTBULK7.Item("STYLE_CODE") = STYLE_CODE
                                rowSOTBULK7.Item("COLOR_CODE") = COLOR_CODE
                                rowSOTBULK7.Item("UPC_CODE") = dst.Tables("ICTSTYC1").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("UPC_CODE") & String.Empty
                                rowSOTBULK7.Item("QTY_PACKED") = ORDR_QTY
                                rowSOTBULK7.Item("PACKED_CODE") = "Each"
                                dst.Tables("SOTBULK7").Rows.Add(rowSOTBULK7)
                            End If
                        Next
                    Next
            End Select

            ' Fill SOTBULK8
            For Each rowSOTBULK6 As DataRow In dst.Tables("SOTBULK6").Select("", "BULK_PATTERN_NO,CART_NO")
                Dim BULK_PATTERN_NO As String = rowSOTBULK6.Item("BULK_PATTERN_NO")

                For Each rowSOTBULK2 As DataRow In dst.Tables("SOTBULK2").Select("BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")
                    Dim rowSOTBULK8 As DataRow = dst.Tables("SOTBULK8").NewRow
                    rowSOTBULK8.Item("BULK_CODE") = rowSOTBULK6.Item("BULK_CODE")
                    rowSOTBULK8.Item("BULK_PATTERN_NO") = BULK_PATTERN_NO
                    rowSOTBULK8.Item("CUST_ADDR_CODE") = rowSOTBULK2.Item("CUST_ADDR_CODE")
                    rowSOTBULK8.Item("CART_NO") = rowSOTBULK6.Item("CART_NO")
                    rowSOTBULK8.Item("TRACKING_NO") = String.Empty
                    rowSOTBULK8.Item("SHIP_LABEL") = String.Empty
                    dst.Tables("SOTBULK8").Rows.Add(rowSOTBULK8)
                Next
            Next

            For Each rowSOTBULK2 As DataRow In dst.Tables("SOTBULK2").Select("", "BULK_CODE, CUST_ADDR_CODE, BULK_PATTERN_NO")
                Dim BULK_CODE As String = rowSOTBULK2.Item("BULK_CODE") & String.Empty
                Dim CUST_ADDR_CODE As String = rowSOTBULK2.Item("CUST_ADDR_CODE") & String.Empty
                Dim BULK_PATTERN_NO As String = rowSOTBULK2.Item("BULK_PATTERN_NO") & String.Empty

                For Each rowSOTBULK6 As DataRow In dst.Tables("SOTBULK6").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'", "BULK_CODE, BULK_PATTERN_NO")
                    Dim CART_NOx As String = rowSOTBULK6.Item("CART_NO") & String.Empty

                    For Each rowSOTBULK7 As DataRow In dst.Tables("SOTBULK7").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND CART_NO = '" & CART_NOx & "'", "BULK_CODE, BULK_PATTERN_NO, CART_NO, CART_LNO")
                        Dim rowSOTBULK267 As DataRow = dst.Tables("SOTBULK267").NewRow
                        rowSOTBULK267.Item("BULK_CODE") = BULK_CODE
                        rowSOTBULK267.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                        rowSOTBULK267.Item("PKG_L") = rowSOTBULK6.Item("PKG_L")
                        rowSOTBULK267.Item("PKG_W") = rowSOTBULK6.Item("PKG_W")
                        rowSOTBULK267.Item("PKG_H") = rowSOTBULK6.Item("PKG_H")
                        rowSOTBULK267.Item("PKG_WEIGHT") = rowSOTBULK6.Item("PKG_WEIGHT")
                        rowSOTBULK267.Item("DESCRIPT") = rowSOTBULK7.Item("QTY_PACKED") & " PC " & rowSOTBULK7.Item("STYLE_CODE") & " - " & rowSOTBULK7.Item("COLOR_CODE")
                        rowSOTBULK267.Item("QTY_PACKED") = rowSOTBULK7.Item("QTY_PACKED")
                        dst.Tables("SOTBULK267").Rows.Add(rowSOTBULK267)
                    Next
                Next
            Next

        Catch ex As Exception
            MessageBox.Show("The following error occurred during cartonization: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            For Each tableName As String In New String() {"SOTBULK6", "SOTBULK7", "SOTBULK267"}
                dst.Tables(tableName).RejectChanges()
            Next
        Finally
            If relaxConstraints Then EnforceConstraints(True)
        End Try
    End Sub
    Private Sub PrintPickSlips()

        Try
            Dim lstCUST_ADDR_CODE As New List(Of String)
            For Each grdRow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTBULK2.Selected.Rows
                lstCUST_ADDR_CODE.Add(grdRow.Cells("CUST_ADDR_CODE").Value)
            Next

            Dim DefaultPrinterName As String = String.Empty
            Dim oPS As New System.Drawing.Printing.PrinterSettings
            Try
                DefaultPrinterName = oPS.PrinterName
            Catch ex As System.Exception
                DefaultPrinterName = ""
            End Try

            If DefaultPrinterName.Length = 0 Then
                MessageBox.Show("You do not have a default Printer.", "Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim REPORT_NAME As String = "WHRBULK1"
            ASCMAIN1.Progress("Printing Pick Slips", "")

            For Each CUST_ADDR_CODE As String In lstCUST_ADDR_CODE

                ASCMAIN1.Progress("-", CUST_ADDR_CODE)

                Dim rpt_title As String = Absx1.txtFor("BULK_DESC").Text
                If rpt_title.Length > 60 Then
                    rpt_title = rpt_title.Substring(0, 60).Trim
                End If

                Print_Report_Begin()
                Generate_Report(REPORT_NAME, rpt_title, "", "{SOTBULK2.CUST_ADDR_CODE} = '" & CUST_ADDR_CODE & "'")
                'Print_Report_End()

                Print_Report_End(True, False, DefaultPrinterName)

            Next

        Catch ex As Exception
            MessageBox.Show("The following error ocurred while Printing Pick Slips: " & ex.Message)

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub RequestShippingLabel(ByRef ErrorMessage As String, ByVal rowSOTBULK9 As DataRow)

        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowSOTCARR1 As DataRow = Nothing

        Dim SHIP_VIA_CODE As String = String.Empty
        Dim SHIP_PACKAGE_NO As String = String.Empty
        Dim pkgId As Int64 = 0
        Dim CARRIER_MESSAGE As String = String.Empty

        Dim isDomesticCredentials As Boolean = False
        Dim isInternationalCredentials As Boolean = False

        Try
            SHIP_VIA_CODE = rowSOTBULK9.Item("SHIP_VIA_CODE") & String.Empty

            rowSOTSVIA1 = LookUp("SOTSVIA1", SHIP_VIA_CODE)
            If rowSOTSVIA1 Is Nothing Then
                ErrorMessage = "Invalid or missing Ship Via for shipping label request"
                Exit Sub
            End If

            rowSOTCARR1 = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty)
            If rowSOTCARR1 Is Nothing Then
                ErrorMessage = "Ship Via is not assigned to a carrier"
                Exit Sub
            End If

            If rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty <> "U" Then
                ErrorMessage = "The carrier for Ship Via (" & SHIP_VIA_CODE & ") is not assigned to a Small Package delivery"
                Exit Sub
            End If

            ' No Labels used for UPS Freight
            If rowSOTSVIA1.Item("FREIGHT_SHIPMENT") & String.Empty = "1" Then
                ErrorMessage = "The carrier for Ship Via (" & SHIP_VIA_CODE & ") is catagorized as a freight shipment"
                Exit Sub
            End If

        Catch ex As Exception
            ErrorMessage = "The following error occurred when evaluating a shipping label request: " & ex.Message
            Exit Sub
        End Try

        Dim rowSOTBULK2 As DataRow = Nothing

        isDomesticCredentials = ",B,D,".Contains(rowSOTBULK9.Item("CARRIER_DI") & String.Empty)
        isInternationalCredentials = ",B,I,".Contains(rowSOTBULK9.Item("CARRIER_DI") & String.Empty)

        If Not (isDomesticCredentials OrElse isInternationalCredentials) Then
            Exit Sub
        End If

        Try
            If rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty <> String.Empty Then

                Dim labelFormatDesc As String = "Unknown"

                For Each vlItem As ValueListItem In optPrint_Type.Items
                    If vlItem.DataValue = rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty Then
                        labelFormatDesc = vlItem.DisplayText
                    End If
                Next

                If optPrint_Type.Value <> rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty Then
                    If MessageBox.Show("Typically you use " & labelFormatDesc & " to print " & rowSOTCARR1.Item("CARRIER_DESC") & " Labels. Do you want to change to " & optPrint_Type.Text & " labels?", "Labels", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                        optPrint_Type.Value = rowSOTCARR1.Item("LABEL_FORMAT") & String.Empty
                    End If
                End If
            End If

            Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
            Dim CARRIER_PROD_CODE As String = rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty

            ' Load and Validate Carrier/Ship Method
            Dim rowSOTCARR2 As DataRow = LookUp("SOTCARR2", New String() {CARRIER_CODE, CARRIER_PROD_CODE})
            If rowSOTCARR2 Is Nothing Then
                ErrorMessage = "Invalid or missing Carrier / Ship Method combination for shipping label request"
                Exit Sub
            End If

            ' Credentials
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            Dim PROVIDER_TYPE As String = (rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty).ToString.Trim

            If rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty = String.Empty Then
                ErrorMessage = "Invalid or missing Carrier Account Number for shipping label request"
                Exit Sub
            End If

            Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(MyBase.Absx1.txtFor("WHSE_CODE").Text)
            If rowICTWHSE1 Is Nothing Then
                ErrorMessage = "Invalid or missing Warehouse"
                Exit Sub
            End If

            Dim clsShip As New TAC.WHCSHIP1

            ' Credentials
            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

            ' Sender Information
            With clsShip.Sender
                .Company = (rowICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .Phone = (rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim

                .FirstName = (rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty

                If .FirstName = String.Empty AndAlso .LastName = String.Empty Then
                    .FirstName = "Warehouse Supervisor"
                End If

                .Address1 = (rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (rowICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (rowICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"
            End With

            ASCMAIN1.Progress("Requesting Shipment labels", "")
            Me.Cursor = Cursors.WaitCursor

            For Each CUST_ADDR_CODE As String In lstLabelsToProcess

                ASCMAIN1.Progress("-", CUST_ADDR_CODE)

                rowSOTBULK2 = dst.Tables("SOTBULK2").Rows.Find(New Object() {HFs("BULK_CODE"), CUST_ADDR_CODE})
                If rowSOTBULK2 Is Nothing Then
                    Continue For
                End If

                If Not IsDate(rowSOTBULK2.Item("SHIP_DATE") & String.Empty) Then
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = "Invalid or missing Ship Date"
                    Continue For
                End If

                If rowSOTBULK2.Item("CUST_COUNTRY") & String.Empty = String.Empty Then
                    rowSOTBULK2.Item("CUST_COUNTRY") = "US"
                End If

                If rowSOTBULK2.Item("CUST_ADDR1") & String.Empty = String.Empty AndAlso rowSOTBULK2.Item("CUST_ADDR2") & String.Empty = String.Empty Then
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = "Invalid or missing Ship To Street Address"
                    Continue For
                ElseIf Not (rowSOTBULK2.Item("CUST_COUNTRY") & String.Empty).ToString.StartsWith("US") AndAlso (rowSOTBULK2.Item("CUST_CITY") & String.Empty = String.Empty OrElse rowSOTBULK2.Item("CUST_ZIP_CODE") & String.Empty = String.Empty) Then
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = "Invalid or missing International Ship To City and/or Zip Code"
                    Continue For
                ElseIf rowSOTBULK2.Item("CUST_CITY") & String.Empty = String.Empty OrElse rowSOTBULK2.Item("CUST_STATE") & String.Empty = String.Empty OrElse rowSOTBULK2.Item("CUST_ZIP_CODE") & String.Empty = String.Empty Then
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = "Invalid or missing Ship To City, State or Zip Code"
                    Continue For
                ElseIf rowSOTBULK2.Item("CUST_COUNTRY") & String.Empty = String.Empty Then
                    Dim STATE_CODE As String = rowSOTBULK2.Item("CUST_STATE") & String.Empty = String.Empty
                    Dim rowTATSTATE As DataRow = tblTATSTATE.Rows.Find(STATE_CODE)
                    If rowTATSTATE IsNot Nothing Then
                        rowSOTBULK2.Item("CUST_COUNTRY").Text = "US"
                    Else
                        rowSOTBULK2.Item("CARRIER_MESSAGE") = "Invalid or missing Country Code"
                        Continue For
                    End If
                End If

                Dim isInternationalShipment As Boolean = False
                Dim fedexSmartPost As Int16 = 26

                ' Recipient
                With clsShip.Recipient
                    If ASCMAIN1.CLIENT = "RGI" AndAlso dst.Tables("SOTBULK1").Rows(0).Item("ORDR_GROUP_NO") & String.Empty = String.Empty Then
                        .FirstName = "STORE " & CUST_ADDR_CODE
                    Else
                        .FirstName = rowSOTBULK2.Item("CUST_NAME") & String.Empty
                    End If

                    .MiddleInitial = ""
                    .LastName = ""

                    .Address1 = rowSOTBULK2.Item("CUST_ADDR1") & String.Empty
                    .Address2 = rowSOTBULK2.Item("CUST_ADDR2") & String.Empty
                    .City = rowSOTBULK2.Item("CUST_CITY") & String.Empty
                    .State = rowSOTBULK2.Item("CUST_STATE") & String.Empty
                    .ZipCode = rowSOTBULK2.Item("CUST_ZIP_CODE") & String.Empty
                    .CountryCode = (rowSOTBULK2.Item("CUST_COUNTRY") & String.Empty).ToUpper.Replace(".", "")
                    If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                    If .CountryCode = "USA" Then .CountryCode = "US"


                    .Company = rowSOTBULK2.Item("CUST_NAME") & String.Empty
                    .Phone = rowSOTBULK2.Item("CUST_AREA_CODE") & String.Empty & rowSOTBULK2.Item("CUST_PHONE") & String.Empty
                    .Phone = .Phone.Replace("-", "").Replace("(", "").Replace(")", "")

                    If .Phone.Trim.Length = 0 Then
                        .Phone = clsShip.Sender.Phone
                    End If

                    If .Phone.Trim.Length = 0 Then
                        .Phone = "1234567890"
                    End If

                    .IsResidental = False
                    .IsPOBox = False

                End With

                ' US Puerto Rico is considered International
                isInternationalShipment = (clsShip.Recipient.CountryCode <> clsShip.Sender.CountryCode) OrElse (clsShip.Recipient.CountryCode = "US" AndAlso clsShip.Recipient.State = "PR")

                ' See if the shipping preference is for Domestic and/or International
                If Not (isDomesticCredentials AndAlso isInternationalCredentials) Then
                    If isDomesticCredentials AndAlso isInternationalShipment Then
                        Continue For
                    End If

                    If isInternationalCredentials AndAlso Not isInternationalShipment Then
                        Continue For
                    End If
                End If

                Select Case PROVIDER_TYPE

                    Case WHCSHIP1.ProviderTypeFedex
                        If Not isInternationalShipment Then
                            clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress
                        Else
                            clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpressInternational
                        End If

                    Case WHCSHIP1.ProviderTypeUPS
                        If Not isInternationalShipment Then
                            clsShip.Service = WHCSHIP1.ServiceProviders.UPS
                        Else
                            clsShip.Service = WHCSHIP1.ServiceProviders.UPSInternational
                        End If

                    Case WHCSHIP1.ProviderTypeUSPS
                        clsShip.Service = WHCSHIP1.ServiceProviders.USPS

                    Case WHCSHIP1.ProviderTypeCanada
                        clsShip.Service = WHCSHIP1.ServiceProviders.CanadaPost

                    Case Else
                        rowSOTBULK2.Item("CARRIER_MESSAGE") = "Could Not determine the Provider type"
                        Continue For
                End Select

                clsShip.PackageDetailList.Clear()
                Dim cartSequenceNos As List(Of Int16) = New List(Of Int16)

                ' Commodities for international shipments
                clsShip.TotalCustomsValue = 0
                clsShip.CommodityDetailList.Clear()
                Dim COMMODITY_LNO As Int16 = 1
                Dim itemList As List(Of String) = New List(Of String)

                Dim BULK_CODE As String = rowSOTBULK2.Item("BULK_CODE") & String.Empty
                Dim BULK_PATTERN_NO As String = rowSOTBULK2.Item("BULK_PATTERN_NO") & String.Empty

                ' If no cartons then do not do anything 
                If dst.Tables("SOTBULK6").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'").Length = 0 Then
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = "No cartons to process"
                    Continue For
                End If

                Dim RatesTotalValue As Double

                ' DO NOT FORGET REFERENCE CODES
                For Each rowSOTBULK6 As DataRow In dst.Tables("SOTBULK6").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'", "CART_NO")

                    Dim CART_NO As String = rowSOTBULK6.Item("CART_NO") & String.Empty
                    Dim reference As String = String.Empty

                    Dim pkgDetail As New nsoftware.InShip.PackageDetail

                    Dim id As String = rowSOTBULK6.Item("CART_NO")
                    pkgDetail.Id = id
                    pkgDetail.Weight = Val(rowSOTBULK6.Item("PKG_WEIGHT") & String.Empty)
                    pkgDetail.Weight *= 16

                    If pkgDetail.Weight = "0" Then
                        pkgDetail.Weight = "16.0"
                    End If

                    pkgDetail.PackagingType = CType(Val(rowSOTBULK6.Item("PACKAGING_TYPE") & String.Empty), nsoftware.InShip.UpsratesPickupTypes)
                    pkgDetail.Length = Val(rowSOTBULK6.Item("PKG_L") & String.Empty)
                    pkgDetail.Width = Val(rowSOTBULK6.Item("PKG_W") & String.Empty)
                    pkgDetail.Height = Val(rowSOTBULK6.Item("PKG_H") & String.Empty)

                    Dim items As String = String.Empty
                    Dim itemsCount As Int16 = 0
                    For Each rowSOTBULK7 As DataRow In dst.Tables("SOTBULK7").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND CART_NO = " & CART_NO, "CART_LNO")
                        Dim STYLE_CODE As String = rowSOTBULK7.Item("STYLE_CODE") & String.Empty
                        Dim COLOR_CODE As String = rowSOTBULK7.Item("COLOR_CODE") & String.Empty
                        Dim QTY_PACKED As Int32 = Val(rowSOTBULK7.Item("QTY_PACKED") & String.Empty)

                        Dim STYLE_PRICE As Decimal = Val(dst.Tables("SOTBULKI").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("STYLE_PRICE") & String.Empty)
                        RatesTotalValue += STYLE_PRICE * QTY_PACKED

                        items &= ", " & QTY_PACKED & " PC - " & STYLE_CODE & " " & COLOR_CODE
                    Next

                    If items.Length > 0 Then
                        items = items.Substring(1).Trim
                    End If

                    If ASCMAIN1.CLIENT = "RGI" And itemsCount <= 2 Then
                        reference = ";CR:" & items
                    End If

                    If reference.StartsWith(";") Then
                        reference = reference.Substring(1).Trim
                    End If

                    pkgDetail.Reference = reference
                    clsShip.PackageDetailList.Add(pkgDetail)
                Next

                If isInternationalShipment Then
                    ' Set the Customs value
                    clsShip.TotalCustomsValue = RatesTotalValue

                    For Each rowSOTBULK7 As DataRow In dst.Tables("SOTBULK7").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")
                        Dim STYLE_CODE As String = rowSOTBULK7.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTBULK7.Item("COLOR_CODE")

                        If itemList.Contains(STYLE_CODE) Then Continue For

                        itemList.Add(STYLE_CODE)

                        Dim rowSOTBULKI As DataRow = dst.Tables("SOTBULKI").Rows.Find(New Object() {HFs("BULK_CODE"), STYLE_CODE, COLOR_CODE})
                        ' Just in case a non item is permitted in the shipmen,
                        If rowSOTBULKI Is Nothing Then Continue For

                        Dim CommodityDetail As New nsoftware.InShip.CommodityDetail
                        CommodityDetail.Description = rowSOTBULKI.Item("STYLE_DESC") & String.Empty

                        Dim NumberOfPieces As Int16 = Val(dst.Tables("SOTBULK7").Compute("SUM(QTY_PACKED)", "BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)

                        CommodityDetail.NumberOfPieces = NumberOfPieces
                        CommodityDetail.Quantity = NumberOfPieces
                        CommodityDetail.QuantityUnit = "EA"
                        CommodityDetail.UnitPrice = Val(dst.Tables("SOTBULKI").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("STYLE_PRICE") & String.Empty)

                        Dim weight As Double = 0
                        Select Case rowSOTBULK7.Item("PACKED_CODE") & String.Empty

                            Case "Carton"
                                weight = Val(rowSOTBULKI.Item("CARTON_WEIGHT") & String.Empty)
                            Case "Inner", "Each"
                                weight = Val(rowSOTBULKI.Item("INNER_WEIGHT") & String.Empty)
                        End Select

                        If weight = 0 Then
                            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
                            If rowICTSTYL1 IsNot Nothing Then
                                weight = Val(rowICTSTYL1.Item("STYLE_WEIGHT") & String.Empty)
                            End If
                        End If

                        CommodityDetail.Weight = weight
                        CommodityDetail.Manufacturer = (rowSOTBULKI.Item("COUNTRY_CODE") & String.Empty).ToString.ToUpper.Trim
                        If CommodityDetail.Manufacturer.Length = 0 Then
                            CommodityDetail.Manufacturer = "US"
                        End If

                        CommodityDetail.Manufacturer = CountryOfOrigin(CommodityDetail.Manufacturer)
                        clsShip.CommodityDetailList.Add(CommodityDetail)
                    Next
                End If

                ' Shipping Method
                If isInternationalShipment Then
                    clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty)
                Else
                    clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
                End If

                If clsShip.RequestedServiceType = fedexSmartPost Then
                    clsShip.FedexSmartPost.HubId = rowSOTCARR3.Item("FEDEX_HUB_ID") & String.Empty
                End If

                ' Payor of the Shipmenet
                clsShip.Payor = nsoftware.InShip.TPayorTypes.ptSender
                If rowSOTBULK9 IsNot Nothing AndAlso rowSOTBULK9.Item("CARRIER_ACCT_NO") & String.Empty <> String.Empty Then
                    Select Case rowSOTBULK9.Item("CARRIER_PAYOR") & String.Empty
                        Case "T"
                            clsShip.Payor = nsoftware.InShip.TPayorTypes.ptThirdParty
                        Case "R"
                            clsShip.Payor = nsoftware.InShip.TPayorTypes.ptRecipient
                        Case Else
                            clsShip.Payor = nsoftware.InShip.TPayorTypes.ptRecipient
                    End Select

                    clsShip.PayorContact.AccountNumber = rowSOTBULK9.Item("CARRIER_ACCT_NO") & String.Empty
                    clsShip.PayorContact.CountryCode = rowSOTBULK9.Item("CARRIER_3PY_COUNTRY") & String.Empty
                    clsShip.PayorContact.ZipCode = rowSOTBULK9.Item("CARRIER_3PY_ZIPCODE") & String.Empty
                    If clsShip.PayorContact.CountryCode = String.Empty Then
                        clsShip.PayorContact.CountryCode = "US"
                    End If

                    clsShip.PayorContact.Company = rowSOTBULK9.Item("COMPANY_NAME") & String.Empty
                    clsShip.PayorContact.Address1 = rowSOTBULK9.Item("COMPANY_ADDR1") & String.Empty
                    clsShip.PayorContact.Address2 = rowSOTBULK9.Item("COMPANY_ADDR2") & String.Empty
                    clsShip.PayorContact.City = rowSOTBULK9.Item("COMPANY_CITY") & String.Empty
                    clsShip.PayorContact.State = rowSOTBULK9.Item("COMPANY_STATE") & String.Empty
                End If

                ' Payor of the Duties
                clsShip.DutiesPayor = clsShip.Payor
                If isInternationalShipment Then
                    clsShip.DutiesPayor = clsShip.Payor
                    clsShip.DutiesPayorContact.AccountNumber = clsShip.PayorContact.AccountNumber
                    clsShip.DutiesPayorContact.CountryCode = clsShip.PayorContact.CountryCode
                    clsShip.DutiesPayorContact.ZipCode = clsShip.PayorContact.ZipCode

                    clsShip.DutiesPayorContact.Company = clsShip.PayorContact.Company
                    clsShip.DutiesPayorContact.Address1 = clsShip.PayorContact.Address1
                    clsShip.DutiesPayorContact.Address2 = clsShip.PayorContact.Address2
                    clsShip.DutiesPayorContact.City = clsShip.PayorContact.City
                    clsShip.DutiesPayorContact.State = clsShip.PayorContact.State
                End If

                clsShip.RequestedUPSInternationalForms.ShippersExportDeclarationInfo = New WHCSHIP1.ShippersExportDeclaration
                clsShip.RequestedUPSInternationalForms.ShippersExportDeclaration = False
                clsShip.RequestedUPSInternationalForms.CommercialInvoice = False

                If isInternationalShipment Then
                    clsShip.RequestedUPSInternationalForms.ShippersExportDeclaration = True
                    With clsShip.RequestedUPSInternationalForms.ShippersExportDeclarationInfo
                        .ImportEntryNumber = String.Empty
                        .InBond = TInBondCodes.ibcNotInBond
                        .LicenseDate = String.Empty
                        .LicenseExceptionCode = TExceptionCodes.ecNLR
                        .LicenseNumber = String.Empty
                        .PointOfOrigin = "US"
                        .ShippersTaxID = String.Empty
                        .TransPortType = String.Empty
                        .ExportingCarrier = CARRIER_CODE
                        .ExportingDate = CDate(rowSOTBULK2.Item("SHIP_DATE") & String.Empty).ToString("yyyyMMdd")
                    End With

                    clsShip.RequestedUPSInternationalForms.CommercialInvoice = True
                    With clsShip.RequestedUPSInternationalForms.CommercialInvoiceInfo
                        .Comments = String.Empty

                        Select Case dst.Tables("SOTBULK1").Rows(0).Item("IMPORT_TYPE") & String.Empty
                            Case "O"
                                .CustomersInvoiceNumber = CUST_ADDR_CODE
                            Case "F"
                                .CustomersInvoiceNumber = "Store " & CUST_ADDR_CODE
                            Case Else
                                .CustomersInvoiceNumber = CUST_ADDR_CODE
                        End Select

                        .FreightCharge = 0
                        .InvoiceDate = rowSOTBULK2.Item("SHIP_DATE")
                        .Purpose = CommercialInvoicePurposes.cipSold
                        .ShipperInsurance = 0
                        .Terms = CommercialInvoiceTerms.citCpt

                    End With
                End If

                clsShip.EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron

                Select Case optPrint_Type.Value
                    Case "E"
                        clsShip.EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron
                    Case "Z"
                        clsShip.EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itZPL
                    Case "X"
                        clsShip.EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itZebra
                End Select

                Dim ShippingLabelDirectory As String = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim

                Try
                    If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "edz" Then
                        Select Case ASCMAIN1.CLIENT
                            Case "RGI"
                                ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "R:\")
                            Case "NYA"
                                ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "N:\")
                        End Select
                    End If

                    If ShippingLabelDirectory.Length > 0 Then
                        If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                            My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                        End If
                    End If

                Catch ex As Exception

                    ShippingLabelDirectory = String.Empty
                End Try

                If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
                    ShippingLabelDirectory = ShippingLabelDirectory & "\"
                End If

                clsShip.ShippingLabelDirectory = ShippingLabelDirectory
                clsShip.ShippingLabelPrefix = HFs("BULK_CODE") & "_" & rowSOTBULK2.Item("CUST_ADDR_CODE")
                clsShip.ShipDate = CDate(rowSOTBULK2.Item("SHIP_DATE") & String.Empty).ToString("yyyy-MM-dd")

                'If ASCMAIN1.CLIENT = "RGI" AndAlso CARRIER_PROD_CODE = WHCSHIP1.UPSFreightProductCode Then
                '    Dim commodity As New nsoftware.InShip.CommodityDetail
                '    With commodity
                '        .Description = "ARTIFICIAL FLOWERS"
                '        .FreightClass = "100"
                '        .FreightNMFC = String.Empty
                '        .FreightNMFCSub = String.Empty
                '        .NumberOfPieces = dst.Tables("SOTCART1").Rows.Count
                '        .Value = Val(dst.Tables("SOTPICK1").Compute("SUM(PICK_AMT_CONF)", "") & String.Empty)
                '        .Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "") & String.Empty)
                '    End With
                '    clsShip.CommodityDetailList.Add(commodity)
                '    clsShip.HandlingUnit = "LOO:" & dst.Tables("SOTCART1").Rows.Count
                'End If

                Select Case ASCMAIN1.CLIENT
                    Case "RGI"
                        clsShip.ShipmentDescription = "Artificial Flowers / Home Decorations"
                    Case Else
                        clsShip.ShipmentDescription = "Clothing / Accessories"

                End Select

                If clsShip.RequestLabel() Then

                    rowSOTBULK2.Item("INTL_FORMS") = String.Empty
                    If isInternationalShipment Then
                        If My.Computer.FileSystem.FileExists(ShippingLabelDirectory & clsShip.ShippingLabelPrefix & TAC.WHCSHIP1.UPSnternationalFormsExtension) Then
                            rowSOTBULK2.Item("INTL_FORMS") = ShippingLabelDirectory & clsShip.ShippingLabelPrefix & TAC.WHCSHIP1.UPSnternationalFormsExtension
                        End If
                    End If

                    CARRIER_MESSAGE = clsShip.LastError & String.Empty
                    rowSOTBULK2.Item("TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty
                    rowSOTBULK2.Item("CARRIER_CODE") = CARRIER_CODE
                    rowSOTBULK2.Item("CARRIER_MESSAGE") = String.Empty

                    For Each shipPackageDetail As nsoftware.InShip.PackageDetail In clsShip.PackageDetailList
                        SHIP_PACKAGE_NO = shipPackageDetail.Id

                        If dst.Tables("SOTBULK8").Select("CUST_ADDR_CODE = '" & CUST_ADDR_CODE & "' AND CART_NO = '" & SHIP_PACKAGE_NO & "'", "").Length > 0 Then
                            Dim rowSOTBULK8 As DataRow = dst.Tables("SOTBULK8").Select("CUST_ADDR_CODE = '" & CUST_ADDR_CODE & "' AND CART_NO = '" & SHIP_PACKAGE_NO & "'", "")(0)
                            rowSOTBULK8.Item("TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                            rowSOTBULK8.Item("FREIGHT_COST") = Val(clsShip.ShipmentNetCharge(SHIP_PACKAGE_NO) & String.Empty)
                            rowSOTBULK8.Item("SHIP_LABEL") = shipPackageDetail.ShippingLabel & String.Empty
                            rowSOTBULK8.Item("COD_LABEL") = shipPackageDetail.CODLabel & String.Empty
                            rowSOTBULK8.Item("RETRUN_RECEIPT") = shipPackageDetail.ReturnReceipt & String.Empty
                        End If

                        ' This is done is case there is an error - we preserve the shipping labels
                        Try
                            Update_Record_TDA("SOTBULK2")
                            Update_Record_TDA("SOTBULK8")
                        Catch ex As Exception

                        End Try
                    Next

                    Try
                        For Each fileName As String In My.Computer.FileSystem.GetFiles(ShippingLabelDirectory, FileIO.SearchOption.SearchTopLevelOnly, clsShip.ShippingLabelPrefix & "*.*")
                            If Not fileName.EndsWith(WHCSHIP1.UPSnternationalFormsExtension) Then
                                My.Computer.FileSystem.DeleteFile(fileName)
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                Else
                    CARRIER_MESSAGE = clsShip.LastError
                    If CARRIER_MESSAGE.Length > 0 Then
                        If CARRIER_MESSAGE.Length > dst.Tables("SOTBULK2").Columns("CARRIER_MESSAGE").MaxLength Then
                            CARRIER_MESSAGE = CARRIER_MESSAGE.Substring(0, dst.Tables("SOTBULK2").Columns("CARRIER_MESSAGE").MaxLength).Trim
                        End If
                        rowSOTBULK2.Item("CARRIER_MESSAGE") = CARRIER_MESSAGE
                    End If
                    'ErrorMessage &= " " & clsShip.LastError
                End If
            Next

        Catch ex As Exception
            ErrorMessage = ex.Message
            CARRIER_MESSAGE = ex.Message
            If CARRIER_MESSAGE.Length > 0 Then
                If CARRIER_MESSAGE.Length > dst.Tables("SOTBULK2").Columns("CARRIER_MESSAGE").MaxLength Then
                    CARRIER_MESSAGE = CARRIER_MESSAGE.Substring(0, dst.Tables("SOTBULK2").Columns("CARRIER_MESSAGE").MaxLength).Trim
                End If
                rowSOTBULK2.Item("CARRIER_MESSAGE") = CARRIER_MESSAGE
            End If

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

        ErrorMessage = ErrorMessage.Trim
        ASCMAIN1.Progress("", "")

    End Sub
    Private Sub PrintShippingLabel(ByRef ErrorMessage As String, ByRef labelsPrinted As Int16, ByVal PrintLabels As Boolean)

        Try

            ASCMAIN1.Progress("Print Shipping Labels", "")
            Me.Cursor = Cursors.WaitCursor

            Dim sql As String = ""
            If dst.Tables("SOTBULKI").Select("SELECTED = '1'").Length > 0 Then
                sql = "SELECTED = '1'"
            End If

            lstLabelsToProcess.Sort()
            Dim counter As Int16 = 0
            For Each rowSOTBULKI As DataRow In dst.Tables("SOTBULKI").Select(sql, "STYLE_CODE,COLOR_CODE")
                Dim STYLE_CODE As String = rowSOTBULKI.Item("STYLE_CODE") & String.Empty
                Dim COLOR_CODE As String = rowSOTBULKI.Item("COLOR_CODE") & String.Empty

                For Each CUST_ADDR_CODE As String In lstLabelsToProcess
                    ASCMAIN1.Progress("-", CUST_ADDR_CODE)

                    For Each rowSOTBULK8 As DataRow In dst.Tables("SOTBULK8").Select("CUST_ADDR_CODE = '" & CUST_ADDR_CODE & "' AND ISNULL(TRACKING_NO, '*') <> '*'", "CART_NO")
                        Dim CART_NO As String = rowSOTBULK8.Item("CART_NO")
                        If dst.Tables("SOTBULK7").Select("CART_NO = '" & CART_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length > 0 Then
                            For Each fieldName As String In New String() {"SHIP_LABEL", "COD_LABEL", "RETRUN_RECEIPT"}

                                ' This gives a count so we can let the user verify if there are enough labels in the printer
                                If Not PrintLabels Then
                                    If rowSOTBULK8.Item(fieldName) & String.Empty <> String.Empty Then
                                        labelsPrinted += 1
                                    End If
                                    Continue For
                                End If

                                rowSOTBULK8.Item("LABEL_PRINTED_DATE") = DateTime.Now
                                rowSOTBULK8.Item("LABEL_PRINTED_OPER") = ASCMAIN1.USER_ID

                                ' This is done in case of error ro user noty updating
                                Try
                                    ASCMAIN1.sql = "UPDATE SOTBULK8 SET LABEL_PRINTED_DATE = SYSDATE, LABEL_PRINTED_OPER = '" & ASCMAIN1.USER_ID & "'" _
                                        & " WHERE BULK_CODE = '" & HFs("BULK_CODE") & "'" _
                                        & " AND BULK_PATTERN_NO = '" & rowSOTBULK8.Item("BULK_PATTERN_NO") & "'" _
                                        & " AND CUST_ADDR_CODE = '" & rowSOTBULK8.Item("CUST_ADDR_CODE") & "'" _
                                        & " AND CART_NO = '" & rowSOTBULK8.Item("CART_NO") & "'"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                                Catch ex As Exception

                                End Try

                                If rowSOTBULK8.Item(fieldName) & String.Empty <> String.Empty Then
                                    Me.Cursor = Cursors.WaitCursor
                                    counter += 1
                                    If counter = 30 Then
                                        ' Each label takes more than one second to print.
                                        ' This is done so we do not overflow the printer buffer.
                                        ' Hopefully this is long enough.
                                        System.Threading.Thread.Sleep(20000)
                                        counter = 0
                                    End If
                                    PrintShipingLabel(rowSOTBULK8.Item(fieldName) & String.Empty)
                                    labelsPrinted += 1
                                End If
                            Next

                            If PrintLabels Then
                                Try
                                    Dim rowSOTBULK2 As DataRow = dst.Tables("SOTBULK2").Rows.Find(New Object() {HFs("BULK_CODE"), rowSOTBULK8.Item("CUST_ADDR_CODE")})
                                    If rowSOTBULK2 IsNot Nothing Then
                                        rowSOTBULK2.Item("LABEL_PRINTED_DATE") = DateTime.Now
                                        rowSOTBULK2.Item("LABEL_PRINTED_OPER") = ASCMAIN1.USER_ID
                                    End If

                                    ASCMAIN1.sql = "UPDATE SOTBULK2 SET LABEL_PRINTED_DATE = SYSDATE, LABEL_PRINTED_OPER = '" & ASCMAIN1.USER_ID & "'" _
                                        & " WHERE BULK_CODE = '" & HFs("BULK_CODE") & "' AND CUST_ADDR_CODE = '" & rowSOTBULK8.Item("CUST_ADDR_CODE") & "'"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                                Catch ex As Exception

                                End Try
                            End If
                        End If
                    Next
                Next

                'If PrintLabels AndAlso lstLabelsToProcess.Count > 1 Then
                '    Dim labelImage As String = String.Empty
                '    labelImage = "EPL2" & Environment.NewLine
                '    labelImage &= "S4" & Environment.NewLine
                '    labelImage &= "UN" & Environment.NewLine
                '    labelImage &= "WN" & Environment.NewLine
                '    labelImage &= "ZT" & Environment.NewLine
                '    labelImage &= "N" & Environment.NewLine
                '    labelImage &= "A50,100,0,4,1,1,N," & Chr(34) & "End of Style " & STYLE_CODE & " / Color " & COLOR_CODE & Chr(34) & Environment.NewLine
                '    labelImage &= "P1" & Environment.NewLine
                '    PrintShipingLabel(labelImage)
                'End If
            Next

        Catch ex As Exception

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default

        End Try
    End Sub

    Public Function PrintShipingLabel(ByVal LabelData As String) As Boolean

        Try
            If (ASCMAIN1.USER_ID = "edz" OrElse ASCMAIN1.USER_ID = "wjz") AndAlso ASCMAIN1.Running_in_VS Then
                ' Find Zebra printer
                Dim zebraPrinter As String = FindZebraPrinter()

                Dim vLabelPrinter As New ASCPRINT
                Return vLabelPrinter.SendStringToPrinter(zebraPrinter, LabelData)
            End If

            ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)

        Catch ex As Exception
            MessageBox.Show("Print Shipping Label Error: " & ex.Message)
        End Try

    End Function

    Private Shared Function FindZebraPrinter() As String

        If ASCMAIN1.LabelPrinterName.Length > 0 Then
            For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
                If printerName.ToUpper = ASCMAIN1.LabelPrinterName.ToUpper Then
                    Return printerName
                End If
            Next printerName
        End If

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA450") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.Contains("450") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.Contains("550") Then
                Return printerName
            End If
        Next printerName

        Return ""
    End Function

    Private Sub GetTrackingInformation()

        ASCMAIN1.Progress("Get Tracking Information", "")

        Try

            Dim rowSOTCARR1 As DataRow = LookUp("SOTCARR1", "UPS")
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & "UPS" & "'")(0)

            For Each CUST_ADDR_CODE As String In lstLabelsToProcess
                Dim rowSOTBULK2 As DataRow = dst.Tables("SOTBULK2").Rows.Find(New Object() {HFs("BULK_CODE"), CUST_ADDR_CODE})

                If rowSOTBULK2 Is Nothing Then
                    Continue For
                End If

                If rowSOTBULK2.Item("TRACKING_NO") & String.Empty = String.Empty Then
                    Continue For
                End If

                ASCMAIN1.Progress("-", CUST_ADDR_CODE)

                Select Case rowSOTBULK2.Item("CARRIER_CODE") & String.Empty

                    Case "UPS"
                        Dim TRACKING_NO As String = rowSOTBULK2.Item("TRACKING_NO") & String.Empty
                        Dim clsShip As New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)

                        ' Credentials
                        clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                        clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
                        clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                        clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                        clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                        clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
                        clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                        clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

                        Dim repsonse As String = clsShip.UPSTrack(TRACKING_NO)
                        rowSOTBULK2.Item("CARRIER_MESSAGE") = repsonse

                        ' This is done is case there is an error - we preserve the shipping label data
                        Try
                            BeginTrans()
                            Update_Record_TDA("SOTBULK2")
                            CommitTrans()
                        Catch ex As Exception
                            Rollback()
                        End Try
                End Select
            Next

        Catch ex As Exception
            MessageBox.Show("Error Getting Tracking Information" & ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Private Sub VoidShippingLabel()

        Dim rowSOTCARR1 As DataRow = LookUp("SOTCARR1", "UPS")
        Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & "UPS" & "'")(0)


        For Each CUST_ADDR_CODE As String In lstLabelsToProcess
            Dim rowSOTBULK2 As DataRow = dst.Tables("SOTBULK2").Rows.Find(New Object() {HFs("BULK_CODE"), CUST_ADDR_CODE})

            If rowSOTBULK2 Is Nothing Then
                Continue For
            End If

            If rowSOTBULK2.Item("TRACKING_NO") & String.Empty = String.Empty Then
                Continue For
            End If

            Select Case rowSOTBULK2.Item("CARRIER_CODE") & String.Empty

                Case "UPS"
                    Dim TRACKING_NO As String = rowSOTBULK2.Item("TRACKING_NO") & String.Empty
                    Dim clsShip As New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)

                    ' Credentials
                    clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                    clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
                    clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
                    clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                    clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                    clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
                    clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
                    clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

                    If clsShip.CancelShipment(TRACKING_NO, True, 0) Then
                        rowSOTBULK2.Item("TRACKING_NO") = String.Empty
                        rowSOTBULK2.Item("LABEL_PRINTED_DATE") = DBNull.Value
                        rowSOTBULK2.Item("LABEL_PRINTED_OPER") = DBNull.Value
                        rowSOTBULK2.Item("INTL_FORMS") = DBNull.Value

                        For Each rowSOTBULK8 As DataRow In dst.Tables("SOTBULK8").Select("BULK_CODE = '" & HFs("BULK_CODE") & "' AND CUST_ADDR_CODE = '" & CUST_ADDR_CODE & "'")
                            rowSOTBULK8.Item("TRACKING_NO") = String.Empty
                            rowSOTBULK8.Item("FREIGHT_COST") = 0
                            rowSOTBULK8.Item("SHIP_LABEL") = String.Empty
                            rowSOTBULK8.Item("COD_LABEL") = String.Empty
                            rowSOTBULK8.Item("RETRUN_RECEIPT") = String.Empty
                            rowSOTBULK8.Item("LABEL_PRINTED_DATE") = DBNull.Value
                            rowSOTBULK8.Item("LABEL_PRINTED_OPER") = DBNull.Value
                        Next

                        ' This is done is case there is an error - we preserve the shipping label data
                        Try
                            BeginTrans()
                            Update_Record_TDA("SOTBULK2")
                            Update_Record_TDA("SOTBULK8")
                            CommitTrans()
                        Catch ex As Exception
                            Rollback()
                        End Try

                    End If
            End Select
        Next

    End Sub

    Private Sub BulkPrintUPSInvoice()

        Try

            EnforceConstraints(False)

            Dim RPT As String = "WHRSHIPE"
            Dim Sql As String = " select * from sotbulk2 where bulk_code like 'BR%' and cust_country = 'CA' and intl_forms is null and tracking_no is not null"

            Dim tblSOTBUK2 As DataTable = ASCDATA1.GetDataTable(Sql)
            Dim tblSOTBULK7 As DataTable = Nothing
            Dim tblSOTBULK5 As DataTable = Nothing
            Dim tblSOTBULKI As DataTable = Nothing
            Dim tblSOTBULK6 As DataTable = Nothing
            Dim rowARTCUST1 As DataRow = Nothing
            Dim rowSOTBULK1 As DataRow = Nothing
            Dim TRACKING_NO As String = String.Empty

            For Each rowSOTBULK2 As DataRow In tblSOTBUK2.Select("", "BULK_CODE, CUST_ADDR_CODE")

                For Each tableName As String In New String() {"UPSINTL1", "UPSINTL2", "UPSINTL5"}
                    dst.Tables(tableName).Rows.Clear()
                Next

                Dim TOTAL_VALUE As Double = 0
                Dim CUST_ADDR_CODE As String = rowSOTBULK2.Item("CUST_ADDR_CODE")
                Dim BULK_CODE As String = rowSOTBULK2.Item("BULK_CODE")
                TRACKING_NO = rowSOTBULK2.Item("TRACKING_NO") & String.Empty

                ASCMAIN1.Progress("Printing UPS Invoices", BULK_CODE & " / " & CUST_ADDR_CODE)

                ASCMAIN1.Progress("-", CUST_ADDR_CODE)
                Dim BULK_PATTERN_NO As String = rowSOTBULK2.Item("BULK_PATTERN_NO")

                If tblSOTBULK7 Is Nothing OrElse tblSOTBULK7.Select("BULK_CODE = '" & BULK_CODE & "'").Length = 0 Then
                    tblSOTBULK7 = ASCDATA1.GetDataTable("SELECT * FROM SOTBULK7 WHERE BULK_CODE = '" & BULK_CODE & "'", "SOTBULK7")
                    tblSOTBULK5 = ASCDATA1.GetDataTable("SELECT * FROM SOTBULK5 WHERE BULK_CODE = '" & BULK_CODE & "'", "SOTBULK5")
                    tblSOTBULKI = ASCDATA1.GetDataTable("SELECT SOTBULKI.*, ICTSTYL1.STYLE_DESC FROM SOTBULKI, ICTSTYL1 WHERE SOTBULKI.BULK_CODE = '" & BULK_CODE & "' AND ICTSTYL1.STYLE_CODE = SOTBULKI.STYLE_CODE", "SOTBULKI")
                    tblSOTBULK6 = ASCDATA1.GetDataTable("SELECT * FROM SOTBULK6 WHERE BULK_CODE = '" & BULK_CODE & "'", "SOTBULK6")
                    Fill_Records("SOTBULK1", BULK_CODE)
                    Fill_Records("SOTBULK2", BULK_CODE)
                    rowSOTBULK1 = dst.Tables("SOTBULK1").Rows(0)
                    rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = '" & rowSOTBULK1.Item("CUST_CODE") & "'")
                End If

                Dim tbl As DataTable = ASCDATA1.SelectDistinct(tblSOTBULK7, New String() {"STYLE_CODE", "COLOR_CODE"})

                For Each row As DataRow In tbl.Select("", "")
                    Dim STYLE_CODE As String = row.Item("STYLE_CODE") & String.Empty
                    Dim COLOR_CODE As String = row.Item("COLOR_CODE") & String.Empty

                    Dim ORDR_QTY As Int32 = Val(tblSOTBULK5.Compute("SUM(ORDR_QTY)", _
                        "BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'") & String.Empty)

                    If ORDR_QTY = 0 Then Continue For

                    Dim rowSOTBULKI As DataRow = tblSOTBULKI.Select("BULK_CODE = '" & BULK_CODE & "' AND STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")(0)
                    If rowSOTBULKI Is Nothing Then
                        Continue For
                    End If

                    Dim rowUPSINTL1 As DataRow = dst.Tables("UPSINTL1").NewRow
                    rowUPSINTL1.Item("BULK_CODE") = BULK_CODE
                    rowUPSINTL1.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL1.Item("UNITS") = ORDR_QTY
                    rowUPSINTL1.Item("UOM") = "EA"
                    rowUPSINTL1.Item("DESC") = rowSOTBULKI.Item("STYLE_DESC") & String.Empty
                    rowUPSINTL1.Item("UNIT_VALUE") = Val(rowSOTBULKI.Item("STYLE_PRICE") & String.Empty)
                    rowUPSINTL1.Item("TOTAL_VALUE") = ORDR_QTY * rowUPSINTL1.Item("UNIT_VALUE")
                    rowUPSINTL1.Item("ORIGIN_COUNTRY") = rowSOTBULKI.Item("COUNTRY_CODE") & String.Empty
                    rowUPSINTL1.Item("ORIGIN_COUNTRY") = CountryOfOrigin(rowUPSINTL1.Item("ORIGIN_COUNTRY") & String.Empty)

                    dst.Tables("UPSINTL1").Rows.Add(rowUPSINTL1)

                    TOTAL_VALUE += rowUPSINTL1.Item("TOTAL_VALUE")
                Next

                Dim rowUPSINTL2 As DataRow = dst.Tables("UPSINTL2").NewRow
                rowUPSINTL2.Item("BULK_CODE") = BULK_CODE
                rowUPSINTL2.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                rowUPSINTL2.Item("INVOICE_LINE_TOTAL") = TOTAL_VALUE
                rowUPSINTL2.Item("DISCOUNT") = 0
                rowUPSINTL2.Item("INVOICE_SUBTOTAL") = TOTAL_VALUE
                rowUPSINTL2.Item("FREIGHT") = 0
                rowUPSINTL2.Item("INSURANCE") = 0
                rowUPSINTL2.Item("OTHER") = 0
                rowUPSINTL2.Item("TOTAL_INVOICE_AMOUNT") = TOTAL_VALUE
                rowUPSINTL2.Item("NUM_CARTONS") = tblSOTBULK6.Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'").Length
                rowUPSINTL2.Item("CURRENCY") = "USD"
                rowUPSINTL2.Item("TOTAL_WEIGHT") = Val(tblSOTBULK6.Compute("SUM(PKG_WEIGHT)", "BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'") & String.Empty)
                dst.Tables("UPSINTL2").Rows.Add(rowUPSINTL2)

                Dim rowUPSINTL5 As DataRow = dst.Tables("UPSINTL5").NewRow
                If ASCMAIN1.CLIENT = "RGI" AndAlso rowSOTBULK1.Item("CUST_CODE") = "021454" And rowSOTBULK2.Item("CUST_COUNTRY") = "CA" Then
                    rowUPSINTL5.Item("BULK_CODE") = BULK_CODE
                    rowUPSINTL5.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL5.Item("CUST_CONTACT") = "Manager"
                    rowUPSINTL5.Item("CUST_NAME") = "The Gap (Canada), Inc."
                    rowUPSINTL5.Item("CUST_ADDR1") = "60 Bloor Street West"
                    rowUPSINTL5.Item("CUST_ADDR2") = "Suite 1500"
                    rowUPSINTL5.Item("CUST_CITY") = "Toronto"
                    rowUPSINTL5.Item("CUST_STATE") = "Ontario"
                    rowUPSINTL5.Item("CUST_ZIP_CODE") = "M4W 3B8"
                    rowUPSINTL5.Item("CUST_PHONE") = rowSOTBULK2.Item("CUST_AREA_CODE") & rowSOTBULK2.Item("CUST_PHONE")

                ElseIf rowARTCUST1 IsNot Nothing Then
                    rowUPSINTL5.Item("BULK_CODE") = BULK_CODE
                    rowUPSINTL5.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL5.Item("CUST_CONTACT") = rowARTCUST1.Item("CUST_CONTACT")
                    rowUPSINTL5.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                    rowUPSINTL5.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1")
                    rowUPSINTL5.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2")
                    rowUPSINTL5.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY")
                    rowUPSINTL5.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE")
                    rowUPSINTL5.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE")
                    rowUPSINTL5.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE")

                Else
                    rowUPSINTL5.Item("BULK_CODE") = BULK_CODE
                    rowUPSINTL5.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                    rowUPSINTL5.Item("CUST_CONTACT") = rowSOTBULK2.Item("CUST_ADDR_CODE")
                    rowUPSINTL5.Item("CUST_NAME") = rowSOTBULK2.Item("CUST_NAME")
                    rowUPSINTL5.Item("CUST_ADDR1") = rowSOTBULK2.Item("CUST_ADDR1")
                    rowUPSINTL5.Item("CUST_ADDR2") = rowSOTBULK2.Item("CUST_ADDR2")
                    rowUPSINTL5.Item("CUST_CITY") = rowSOTBULK2.Item("CUST_CITY")
                    rowUPSINTL5.Item("CUST_STATE") = rowSOTBULK2.Item("CUST_STATE")
                    rowUPSINTL5.Item("CUST_ZIP_CODE") = rowSOTBULK2.Item("CUST_ZIP_CODE")
                    rowUPSINTL5.Item("CUST_PHONE") = rowSOTBULK2.Item("CUST_AREA_CODE") & rowSOTBULK2.Item("CUST_PHONE")
                End If

                dst.Tables("UPSINTL5").Rows.Add(rowUPSINTL5)

                Dim TempExportFilename As String = TRACKING_NO
                Print_Report_Begin()
                Generate_Report(RPT, "UPS Invoice", "", "", "PDF", TempExportFilename, False)
                Print_Report_End(True, True, )

                If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & TempExportFilename & ".pdf") Then
                    My.Computer.FileSystem.MoveFile(ASCMAIN1.Folders("Temp") & TempExportFilename & ".pdf", "C:\Users\ezenker\Desktop\Regency BR\UPSIntlForms\" & TempExportFilename & ".pdf", True)
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("The following error ocurred while generating UPS Intl Invoices: " & ex.Message)

        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try

    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTBULK2_SHIP_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTBULK2_SHIP.ClickCellButton
        Select Case e.Cell.Column.Key

            Case "INTL_FORMS"
                If e.Cell.Value & String.Empty = String.Empty Then
                    grdSOTBULK2_SHIP.Selected.Rows.Clear()
                    grdSOTBULK2_SHIP.Selected.Rows.Add(e.Cell.Row)
                    PrintUPSInvoice()
                    Exit Sub
                End If

                If Not My.Computer.FileSystem.FileExists(e.Cell.Value) Then
                    MessageBox.Show("Cannot locate the International Form: " & e.Cell.Value, "International Forms", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

                If MessageBox.Show("Do you want to view the International Forms for this shipment?", "International Forms", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Try
                    System.Diagnostics.Process.Start(e.Cell.Value)
                Catch ex As Exception
                    MessageBox.Show("Error accessing the International Form: " & ex.Message)
                End Try
        End Select
    End Sub

    Private Sub txtLabelPrinter_EditorButtonClick(sender As Object, e As UltraWinEditors.EditorButtonEventArgs) Handles txtLabelPrinter.EditorButtonClick
        Try
            PrintShipingLabel(TAC.TACMAIN1.testLabel)
        Catch ex As Exception
            MessageBox.Show("Error Generating Test Label: " & ex.Message)
        End Try
    End Sub

    Private Sub txtLaserPrinter_EditorButtonClick(sender As Object, e As UltraWinEditors.EditorButtonEventArgs) Handles txtLaserPrinter.EditorButtonClick
        Try
            MessageBox.Show("This does not print a test page.", "Print", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error Generating Test report: " & ex.Message)
        End Try
    End Sub

    Private Sub grdSOTBULK4_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSOTBULK4.BeforeRowUpdate
        Dim BULK_PATTERN_DESC As String = e.Row.Cells("BULK_PATTERN_DESC").Value & String.Empty
        BULK_PATTERN_DESC = BULK_PATTERN_DESC.Trim
        If BULK_PATTERN_DESC.Length = 0 Then
            MessageBox.Show("The Pattern Description is required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        Dim BULK_PATTERN_NO As String = e.Row.Cells("BULK_PATTERN_NO").Value & String.Empty
        If dst.Tables("SOTBULK4").Select("BULK_PATTERN_DESC = '" & BULK_PATTERN_DESC & "' AND BULK_PATTERN_NO <> '" & BULK_PATTERN_NO & "'").Length > 0 Then
            MessageBox.Show("The Pattern Description is already in use", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If
    End Sub

    Private Sub grdSOTBULK2_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTBULK2.InitializeLayout
        e.Layout.UseFixedHeaders = True
        e.Layout.Bands(0).Columns("CUST_ADDR_CODE").Header.Fixed = True

        ' Set the appearance of fixed headers.
        e.Layout.Override.FixedHeaderAppearance.BackColor = Drawing.Color.LightYellow
        e.Layout.Override.FixedHeaderAppearance.ForeColor = Drawing.Color.Blue

        ' Set the appearance of cells associated witht any fixed headers.
        e.Layout.Override.FixedCellAppearance.BackColor = Drawing.Color.LightYellow
        e.Layout.Override.FixedCellAppearance.ForeColor = Drawing.Color.Blue

        ' Set the color of the separator line the separates the fixed cells
        ' from non-fixed cells.
        e.Layout.Override.FixedCellSeparatorColor = Drawing.Color.Red

    End Sub

    Private Sub grdSOTBULK6_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSOTBULK6.BeforeRowUpdate

        Select Case e.Row.Band.Key

            Case "SOTBULK4_SOTBULK6"
                e.Row.Cells("BULK_CODE").Value = e.Row.ParentRow.Cells("BULK_CODE").Value
                e.Row.Cells("BULK_PATTERN_NO").Value = e.Row.ParentRow.Cells("BULK_PATTERN_NO").Value

                Dim BULK_CODE As String = e.Row.Cells("BULK_CODE").Value
                Dim BULK_PATTERN_NO As String = e.Row.Cells("BULK_PATTERN_NO").Value

                Dim CART_NO As String = e.Row.Cells("CART_NO").Value & String.Empty
                If CART_NO.Length = 0 Then
                    CART_NO = Val(dst.Tables("SOTBULK6").Compute("MAX(CART_NO)", "BULK_CODE = '" & BULK_CODE & "'") & String.Empty) + 1
                    e.Row.Cells("CART_NO").Value = CART_NO.ToString.PadLeft(10, "0")
                End If

                Dim rowWHTPKGM1 As DataRow = Nothing
                Dim PKG_CODE As String = e.Row.Cells("PKG_CODE").Value & String.Empty

                If dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "' AND PKG_CODE <> 'OTHER'").Length > 0 Then
                    rowWHTPKGM1 = dst.Tables("WHTPKGM1").Select("PKG_CODE = '" & PKG_CODE & "'")(0)
                    e.Row.Cells("PKG_L").Value = rowWHTPKGM1.Item("PKG_L")
                    e.Row.Cells("PKG_W").Value = rowWHTPKGM1.Item("PKG_W")
                    e.Row.Cells("PKG_H").Value = rowWHTPKGM1.Item("PKG_H")
                End If

                ' Sort the values by length, width, height
                Dim PKG_L As Decimal = Val(e.Row.Cells("PKG_L").Value & String.Empty)
                Dim PKG_W As Decimal = Val(e.Row.Cells("PKG_W").Value & String.Empty)
                Dim PKG_H As Decimal = Val(e.Row.Cells("PKG_H").Value & String.Empty)
                Dim PKG_WEIGHT As Decimal = Val(e.Row.Cells("PKG_WEIGHT").Value & String.Empty)

                If PKG_L <= 0 OrElse PKG_W <= 0 OrElse PKG_H < 0 Then
                    MessageBox.Show("All dimensions must be greater than 0", "Update", MessageBoxButtons.OK)
                    e.Cancel = True
                    Exit Sub
                End If

                If PKG_WEIGHT <= 0 Then
                    MessageBox.Show("Package Weight must be greater than 0", "Update", MessageBoxButtons.OK)
                    e.Cancel = True
                    Exit Sub
                End If

                Dim dimList As New List(Of Decimal)
                dimList.Add(PKG_L)
                dimList.Add(PKG_W)
                dimList.Add(PKG_H)
                dimList.Sort()
                PKG_L = dimList(2)
                PKG_W = dimList(1)
                PKG_H = dimList(0)

                e.Row.Cells("PKG_L").Value = PKG_L
                e.Row.Cells("PKG_W").Value = PKG_W
                e.Row.Cells("PKG_H").Value = PKG_H

            Case "SOTBULK6_SOTBULK7"
                e.Row.Cells("BULK_CODE").Value = e.Row.ParentRow.Cells("BULK_CODE").Value
                e.Row.Cells("BULK_PATTERN_NO").Value = e.Row.ParentRow.Cells("BULK_PATTERN_NO").Value
                e.Row.Cells("CART_NO").Value = e.Row.ParentRow.Cells("CART_NO").Value

                Dim BULK_CODE As String = e.Row.Cells("BULK_CODE").Value
                Dim BULK_PATTERN_NO As String = e.Row.Cells("BULK_PATTERN_NO").Value
                Dim CART_NO As String = e.Row.Cells("CART_NO").Value & String.Empty
                Dim CART_LNO As Int16 = Val(e.Row.Cells("CART_LNO").Value & String.Empty)

                If CART_LNO = 0 Then
                    CART_LNO = Val(dst.Tables("SOTBULK7").Compute("MAX(CART_LNO)", "BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND CART_NO = " & CART_NO) & String.Empty) + 1
                    e.Row.Cells("CART_LNO").Value = CART_LNO
                End If

                Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & String.Empty
                Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & String.Empty

                If STYLE_CODE.Length = 0 OrElse COLOR_CODE.Length = 0 Then
                    MessageBox.Show("Style and Color are required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                    Exit Sub
                End If

                Dim fieldName As String = STYLE_CODE & "_" & COLOR_CODE
                If Not dst.Tables("SOTBULK4").Columns.Contains(fieldName) Then
                    MessageBox.Show("The provided Style and Color combination is not on this order.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                    Exit Sub
                End If

                Dim QTY_PACKED As Int32 = Val(e.Row.Cells("QTY_PACKED").Value & String.Empty)
                If QTY_PACKED <= 0 Then
                    MessageBox.Show("The Quantity must be greater than 0.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                    Exit Sub
                End If

                Dim qtyOrdered As Int32 = 0
                For Each row As DataRow In dst.Tables("SOTBULK7").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "' AND STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")
                    If row.Item("CART_NO") <> CART_NO Then
                        qtyOrdered += Val(row.Item("QTY_PACKED") & String.Empty)
                    ElseIf row.Item("CART_LNO") <> CART_LNO Then
                        qtyOrdered += Val(row.Item("QTY_PACKED") & String.Empty)
                    End If
                Next

                qtyOrdered += QTY_PACKED

                Dim totalStyleColorQty As Int32 = Val(dst.Tables("SOTBULK4").Select("BULK_CODE = '" & BULK_CODE & "' AND BULK_PATTERN_NO = '" & BULK_PATTERN_NO & "'")(0).Item(fieldName) & String.Empty)

                If qtyOrdered > totalStyleColorQty Then
                    MessageBox.Show("The total of the quantities entered (" & qtyOrdered & ") for this Style / Color exceeds the total quantity ordered (" & totalStyleColorQty & ")", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                    Exit Sub
                End If

        End Select
    End Sub

    Private Sub grdSOTBULK6_SelectionDrag(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles grdSOTBULK6.SelectionDrag
        grdSOTBULK6.DoDragDrop(grdSOTBULK6.Selected.Rows, DragDropEffects.Move)
    End Sub

    Private Sub grdSOTBULK6_DragOver(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdSOTBULK6.DragOver
        e.Effect = DragDropEffects.Move
        Dim grid As UltraWinGrid.UltraGrid = TryCast(sender, UltraWinGrid.UltraGrid)
        Dim pointInGridCoords As Drawing.Point = grid.PointToClient(New Drawing.Point(e.X, e.Y))

        If pointInGridCoords.Y < 20 Then
            'Scroll up
            grdSOTBULK6.ActiveRowScrollRegion.Scroll(UltraWinGrid.RowScrollAction.LineUp)
        ElseIf pointInGridCoords.Y > grid.Height - 20 Then
            'Scroll down
            grdSOTBULK6.ActiveRowScrollRegion.Scroll(UltraWinGrid.RowScrollAction.LineDown)
        End If
    End Sub

    Private Sub grdSOTBULK6_DragDrop(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdSOTBULK6.DragDrop
        Dim dropIndex As Integer

        'Get the position on the grid where the dragged row(s) are to be dropped.
        'get the grid coordinates of the row (the drop zone)
        Dim uieOver As UIElement = grdSOTBULK6.DisplayLayout.UIElement.ElementFromPoint(grdSOTBULK6.PointToClient(New Drawing.Point(e.X, e.Y)))

        'get the row that is the drop zone/or where the dragged row is to be dropped
        Dim ugrOver As UltraWinGrid.UltraGridRow = TryCast(uieOver.GetContext(GetType(UltraWinGrid.UltraGridRow), True), UltraWinGrid.UltraGridRow)

        If ugrOver IsNot Nothing Then
            dropIndex = ugrOver.Index    'index/position of drop zone in grid

            ' We are moving SOTBULK7 records to a new SOTBULK6 
            Dim BAND_KEY = ugrOver.Band.Key
            If BAND_KEY <> "SOTBULK4_SOTBULK6" Then
                Exit Sub
            End If

            Dim FROM_BULK_PATTERN_NO As String = ugrOver.Cells("BULK_PATTERN_NO").Value
            Dim TO_CART_NO As String = ugrOver.Cells("CART_NO").Value
            Dim numRowsMoved As Int16 = 0

            'get the dragged row(s)which are to be dragged to another position in the grid
            Dim SelRows As UltraWinGrid.SelectedRowsCollection = TryCast(DirectCast(e.Data.GetData(GetType(UltraWinGrid.SelectedRowsCollection)), UltraWinGrid.SelectedRowsCollection), UltraWinGrid.SelectedRowsCollection)

            Dim selectedRows As New List(Of UltraWinGrid.UltraGridRow)
            For iLoop As Int16 = 0 To SelRows.Count - 1
                selectedRows.Add(SelRows(iLoop))
            Next

            For Each aRow As UltraWinGrid.UltraGridRow In selectedRows

                If aRow.Band.Key <> "SOTBULK6_SOTBULK7" Then
                    Continue For
                End If

                Dim TO_BULK_PATTERN_NO As String = aRow.Cells("BULK_PATTERN_NO").Value
                If FROM_BULK_PATTERN_NO <> TO_BULK_PATTERN_NO Then
                    Continue For
                End If

                Dim BULK_CODE As String = aRow.Cells("BULK_CODE").Value
                Dim BULK_PATTERN_NO As String = aRow.Cells("BULK_PATTERN_NO").Value
                Dim CART_NO As String = aRow.Cells("CART_NO").Value
                Dim CART_LNO As Int16 = aRow.Cells("CART_LNO").Value

                Dim rowSOTBULK7 As DataRow = dst.Tables("SOTBULK7").Rows.Find(New Object() {BULK_CODE, BULK_PATTERN_NO, CART_NO, CART_LNO})
                If rowSOTBULK7 IsNot Nothing Then
                    ' do this to prevent constraint
                    rowSOTBULK7.Item("CART_LNO") = -1
                    rowSOTBULK7.Item("CART_NO") = TO_CART_NO
                    rowSOTBULK7.Item("CART_LNO") = Val(dst.Tables("SOTBULK7").Compute("MAX(CART_LNO)", "CART_NO = " & TO_CART_NO) & String.Empty) + 1
                    If rowSOTBULK7.Item("CART_LNO") = 0 Then
                        rowSOTBULK7.Item("CART_LNO") = 1
                    End If
                End If

                numRowsMoved += 1

            Next

            MessageBox.Show(numRowsMoved & " items moved.", "Move", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub grdSOTBULK9_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSOTBULK9.BeforeRowUpdate

        Dim SHIP_VIA_CODE As String = e.Row.Cells("SHIP_VIA_CODE").Value & String.Empty
        Dim CARRIER_CODE As String = e.Row.Cells("CARRIER_CODE").Value & String.Empty
        Dim CARRIER_ACCT_NO As String = e.Row.Cells("CARRIER_ACCT_NO").Value & String.Empty
        Dim CARRIER_3PY_COUNTRY As String = e.Row.Cells("CARRIER_3PY_COUNTRY").Value & String.Empty
        Dim CARRIER_3PY_ZIPCODE As String = e.Row.Cells("CARRIER_3PY_ZIPCODE").Value & String.Empty
        Dim CARRIER_PAYOR As String = e.Row.Cells("CARRIER_PAYOR").Value & String.Empty
        Dim CUST_BRAND As String = e.Row.Cells("CUST_BRAND").Value & String.Empty

        e.Row.Cells("BULK_CODE").Value = Absx1.txtFor("BULK_CODE").Text

        Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
        If rowSOTSVIA1 Is Nothing Then
            MessageBox.Show("Invalid Ship Via Code.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty <> CARRIER_CODE Then
            MessageBox.Show("Invalid Ship Via Code for this carrier.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If rowSOTSVIA1.Item("SHIP_VIA_STATUS") & String.Empty <> "A" Then
            MessageBox.Show("The selected Ship Via Code is NOT active.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If CARRIER_PAYOR.Length = 0 Then
            MessageBox.Show("Payor type is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If CUST_BRAND.Length = 0 Then
            MessageBox.Show("Brand is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If CARRIER_PAYOR = "S" Then
            If CARRIER_ACCT_NO.Length > 0 Then
                MessageBox.Show("Payor Type Sender cannot contain an Account Number.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
        Else
            If CARRIER_ACCT_NO.Length = 0 Then
                MessageBox.Show("The provided Payor Type requires an Account Number.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
        End If

        ' Country and Zip code are required when providing Carrier Account No
        If CARRIER_ACCT_NO.Length > 0 Then

            If CARRIER_3PY_COUNTRY.Length = 0 Then
                MessageBox.Show("Country code for the provided Account Number is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If

            If CARRIER_3PY_ZIPCODE.Length = 0 Then
                MessageBox.Show("Zip code for the provided Account Number is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
        End If

    End Sub

    Private Sub grdSOTBULK9_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTBULK9.ClickCellButton

        Select Case e.Cell.Column.Key

            Case "SHIP_VIA_CODE"
                grdClickCellButton(grdSOTBULK9, "CARRIER_CODE = '" & grdSOTBULK9.ActiveRow.Cells("CARRIER_CODE").Value & "' AND SHIP_VIA_STATUS = 'A'", , , "SHIP_VIA_CODE")
        End Select

    End Sub

#End Region

#Region "Overrides"
    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As Control, COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME

            Case "ORDR_CUST_PO"
                Cancel = True
                If MyBase.Absx1.txtFor("CUST_CODE").TextLength = 0 Then
                    MessageBox.Show("You must provide a customer code.", "Look Up", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                ASCMAIN1.CodeSelector.SQL = "SELECT ORDR_GROUP_NO, CUST_CODE, TRUNC(ORDR_DATE) ORDR_DATE, ORDR_CUST_PO FROM SOTORDR0 WHERE CUST_CODE =  '" & MyBase.Absx1.txtFor("CUST_CODE").Text & "' AND ORDR_DATE >= SYSDATE - 365 ORDER BY ORDR_DATE"
                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                Using F As New ASFCODE1
                    F.ShowDialog()
                End Using

                If ASCMAIN1.CodeSelector.SelectedRows.Count > 0 Then
                    MyBase.Absx1.txtFor("ORDR_CUST_PO").Text = ASCMAIN1.CodeSelector.SelectedRows(0).Item("ORDR_CUST_PO") & String.Empty
                End If

        End Select
    End Sub

    Public Overrides Sub txt_EditorButtonClick_Special(txtctl As UltraWinEditors.UltraTextEditor)
        MyBase.txt_EditorButtonClick_Special(txtctl)

        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BULK_CODE"
                If txtctl.TextLength > 0 Then
                    Click_Command("Edit")
                End If

        End Select
    End Sub

#End Region

#Region "Serial and Com Connections"

    ''' <summary>
    ''' Form activate - Calls to setup devices
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    ''' <summary>
    ''' Sets the Printer Settings
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetUpPortsAndPrinters()

        '   Dim tooltip As New System.Windows.Forms.ToolTip()

        '**************************
        '**    Laser Printer
        '**************************
        Try
            txtLaserPrinter.Text = ASCMAIN1.LaserPrinterIpAddress
            '  tooltip.SetToolTip(txtLaserPrinter, ASCMAIN1.LaserPrinterIpAddress)

            If ASCMAIN1.LaserPrinterIpAddress.Length = 0 Then
                txtLaserPrinter.Appearance.BackColor = Drawing.Color.Red
            Else
                txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow

                ' Need to Remove the Port Number if it exists
                Dim LaserPrinterIpAddress As String = ASCMAIN1.LaserPrinterIpAddress

                If LaserPrinterIpAddress.Contains(":") Then
                    LaserPrinterIpAddress = LaserPrinterIpAddress.Split(":")(0)
                End If

                If Net.IPAddress.TryParse(LaserPrinterIpAddress, Nothing) Then
                    txtLaserPrinter.Appearance.BackColor = Drawing.Color.Green
                End If

            End If

        Catch ex As Exception
            txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow
            '  tooltip.SetToolTip(txtLaserPrinter, ex.Message)
        End Try


        '**************************
        '**    Label Printer Port
        '**************************        
        Try
            txtLabelPrinter.BackColor = Drawing.Color.Red

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtLabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                '    tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            Else
                Me.txtLabelPrinter.Text = "No Port"
                '      tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            End If

            txtLabelPrinter.BackColor = Drawing.Color.Yellow
            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                If Not ASCMAIN1.Running_in_VS Then ASCMAIN1.LabelPrinterSerialPort.Open()
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                txtLabelPrinter.BackColor = Drawing.Color.Green
            End If

        Catch ex As Exception
            txtLabelPrinter.BackColor = Drawing.Color.Red
            ' tooltip.SetToolTip(txtLabelPrinter, ex.Message)
        End Try

        '**************************
        '**    Scale Port
        '************************** 
        'Try
        '    '  tooltip.SetToolTip(txtScale, ASCMAIN1.ScalePort.PortName & ", " & ASCMAIN1.ScalePort.BaudRate & ", " & ASCMAIN1.ScalePort.DataBits & ", " & ASCMAIN1.ScalePort.Parity.ToString & ", " & ASCMAIN1.ScalePort.StopBits)
        '    txtScale.Text = ASCMAIN1.ScalePort.PortName
        '    txtScale.Appearance.BackColor = Drawing.Color.Green
        'Catch ex As Exception
        '    txtScale.Text = String.Empty
        '    txtScale.Appearance.BackColor = Drawing.Color.Red
        '    ' tooltip.SetToolTip(txtScale, ex.Message)
        'End Try

    End Sub

#End Region

End Class