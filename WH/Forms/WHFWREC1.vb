Imports Infragistics.Win.UltraWinGrid

Public Class WHFWREC1


    '3) UI TO ALLOW USER TO LOCK A LOCATION

    ' ENABLE LOCATOR/CTN_CTL RECEIPTS using new tab in POFSHIP1 
    '  so that Nadine can receive these - study ADS & NYA code sections

    ' IF A WREC EXISTS - THAT MEANS YOU CANNOT EDIT A SHIPMENT AT ALL
    '   - IN POFSHIP1 - BLOCK NADINE IF WREC_NO EXISTS - potship2.wh_rec_no is not null

    ' what does reversal do?

    ' right click to collapse / expand - refactor chucnk of code in load_record for collapse
    ' MAKE A DBL CLICK JUST SELECT ALL CONTAINERS - Done
    ' DEFAULT CONTAINER INTO TRAILER - Done
    ' CHECK THAT ALL BAR_CODES in DATATABLE DO NOT exist in WHTBARC1 IN ORACLE - Done
    ' update needs to validate balancing - Done
    ' check status of WH_REC_NO (for editing) and of PO_SHIPMENT_NOs (for new receipts) after multitask is done - Done
    ' EMAIL LISTS DISTINCT CONTAINER IDS - Done
    ' Executed refresh in tabMain_SelectedTabChanged
    ' Hooked up grdTATEVNT1


    Dim PO_SHIPMENT_NO As String
    Dim PO_SHIPMENT_LNO As Integer
    Dim CONTAINER_NO As String
    Dim double_clicked As Boolean = False

    Dim rowPOTSHIP1 As DataRow
    Dim rowWHTWREC1 As DataRow

    Dim WHSE_CODE As String
    Dim LOCATION_CODE As String
    Dim User_Whse_Code As String = ""

    Dim CARTON_NO As Integer
    Dim WH_REC_NO As String
    Dim PO_ORDER_NO As String
    Dim BAR_CODE_deleted As New List(Of String)
    Dim QTY_deleted As New List(Of Int32)
    Dim sqlPOTSHIPX As String
    Dim tagCR As Boolean = False
    Dim Deleted_BarCodes As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("WHTPARM1")

        If MENU_ITEM_OBJECT = "WHFWRECI" Then
            InquiryMode = True
        End If

        With dst
            User_Whse_Code = IIf(ASCMAIN1.USER_SECURITY_CODEs.Contains("NJT"), "NJT", "NJE")
            ASCMAIN1.sql = "Select * from ICTWHSE1 where WHSE_LOCATOR = '1' And WHSE_CODE = '" & User_Whse_Code & "'"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select TATEVNT1.* " _
            & " from TATEVNT1 " _
            & " where TABLE_NAME = 'WHTWREC1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V", 0)

            ASCMAIN1.sql = "Select WHTWREC1.*, POTSHIP1.PO_SHIP_VESSEL" & vbCrLf _
            & " from WHTWREC1, POTSHIP1 where POTSHIP1.PO_SHIPMENT_NO = WHTWREC1.PO_SHIPMENT_NO" & vbCrLf _
            & " and WHTWREC1.OPS_YYYYPP between :PARM1 and :PARM2"
            Create_TDA(.Tables.Add, "WHTWRECX", "**", 0, False, "VV", 0)


            sqlPOTSHIPX = "Select POTSHIP1.*" & vbCrLf _
                & ", POTSHIP2.CONTAINER_NO, POTSHIP2.PO_SHIP_STATUS" & vbCrLf _
                & ", POTSHIP2.PO_SHIPMENT_LNO, POTSHIP2.BOL_NO, POTSHIP2.PO_SHIP_CTNS" & vbCrLf _
                & " from POTSHIP1,POTSHIP2 " & vbCrLf _
                & " where POTSHIP2.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO" & vbCrLf _
                & "   and POTSHIP1.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_LOCATOR = '1' and WHSE_CTN_CTL = 'C')"

            ASCMAIN1.sql = sqlPOTSHIPX & "   and POTSHIP2.PO_SHIP_STATUS = 'O' and WH_REC_NO is Null"
            Create_TDA(.Tables.Add, "POTSHIPX", "**", 0, False, "", 0)

            With .Tables("POTSHIPX")
                .Columns.Add("CONTAINER_TYPE_CODE")
                .Columns.Add("SELECTED")
                .PrimaryKey = New DataColumn() { .Columns("PO_SHIPMENT_NO"), .Columns("PO_SHIPMENT_LNO")}
            End With



            ASCMAIN1.sql = "Select POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                & ", POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO, POTSHIP3.PO_QTY_SHP" & vbCrLf _
                & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR2.PO_QTY_UOM" & vbCrLf _
                & ", POTORDR1.PO_REFERENCE, POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
                & ", ICTSTYL1.STYLE_DESC, ICTSTYL1.SUB_UNIT_PACK_QTY" & vbCrLf _
                & ", ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                & ", ICTSTYL1.STYLE_WEIGHT , ICTSTYL1.STYLE_UOM, ICTSTYC1.STYLE_BIN " & vbCrLf _
                & ", POTORDR1.CUST_CODE, ARTCUST1.CUST_NAME  " & vbCrLf _
                & " from POTSHIP3,POTORDR2,POTORDR1,ICTSTYL1, ICTSTYC1, ARTCUST1 " & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = POTORDR2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = POTORDR2.COLOR_CODE" & vbCrLf _
                & "   and POTORDR1.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
                & "   and POTORDR1.CUST_CODE = ARTCUST1.CUST_CODE (+)" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_NO = :PARM1" & vbCrLf _
                & "   and POTSHIP3.PO_SHIPMENT_LNO = :PARM2"

            Create_TDA(.Tables.Add, "POTSHIP3", "**", 0, False, "VN", 4)
            .Tables("POTSHIP3").Columns.Add("UNITS", GetType(System.Decimal),
                "IIF(ISNULL(PO_QTY_UOM,0)=1,PO_QTY_SHP,PO_QTY_SHP * PO_QTY_UOM / IIF(ISNULL(SUB_UNIT_PACK_QTY,0)=0,1,ISNULL(SUB_UNIT_PACK_QTY,0)))")
            .Tables("POTSHIP3").Columns.Add("CARTON_NOS")


            ASCMAIN1.sql = "Select WHTBARC1.PO_SHIPMENT_NO, WHTBARC1.PO_SHIPMENT_LNO, WHTBARC1.CARTON_NO" & vbCrLf _
                & ", WHTBARC1.BAR_CODE, WHTBARC1.BAR_CODE BAR_CODE2, PO_SHIPMENT_LNO QTY" & vbCrLf _
                & ", WHTBARC1.LOAD_NO, WHTBARC1.BAR_CODE BAR_CODE_WITH" & vbCrLf _
                & " from WHTBARC1" & vbCrLf _
                & " where WHTBARC1.PO_SHIPMENT_NO = :PARM1" & vbCrLf _
                & "   and WHTBARC1.PO_SHIPMENT_LNO = :PARM2"
            Create_TDA(.Tables.Add, "POTSHIPC", "**", 0, False, "VN", 4)
            With .Tables("POTSHIPC")
                .Columns("BAR_CODE2").AllowDBNull = True
                .Columns("QTY").AllowDBNull = True
                .Columns("BAR_CODE_WITH").AllowDBNull = True
            End With

            ASCMAIN1.sql = "Select WHTWREC7.*, WHTSCSEQ.STYLE_SEQ" & vbCrLf _
                & " from WHTWREC7, WHTSCSEQ" & vbCrLf _
                & " where WHTWREC7.WH_REC_NO = :PARM1" & vbCrLf _
                & "   and WHTWREC7.PO_SHIPMENT_NO = :PARM2" & vbCrLf _
                & "   and WHTWREC7.PO_SHIPMENT_LNO = :PARM3" & vbCrLf _
                & "   and WHTWREC7.STYLE_CODE = WHTSCSEQ.STYLE_CODE(+)" & vbCrLf _
                & "   and WHTWREC7.COLOR_CODE = WHTSCSEQ.COLOR_CODE(+)"
            Create_TDA(.Tables.Add, "WHTWREC7", "**", 0, True, "VVN", 4)

            Create_Relation("WHTWREC7", "POTSHIPC", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO,CARTON_NO")
            .Tables("WHTWREC7").Columns.Add("BAR_CODES", GetType(System.Int32), "SUM(CHILD(WHTWREC7_POTSHIPC).QTY)")

            ASCMAIN1.sql = "Select WHTWREC8.*" & vbCrLf _
                & " from WHTWREC8" & vbCrLf _
                & " where WHTWREC8.WH_REC_NO = :PARM1" & vbCrLf _
                & "   and WHTWREC8.PO_SHIPMENT_NO = :PARM2" & vbCrLf _
                & "   and WHTWREC8.PO_SHIPMENT_LNO = :PARM3"
            Create_TDA(.Tables.Add, "WHTWREC8", "**", 0, True, "VVN", 6)


            Create_Relation("WHTWREC7", "WHTWREC8", "WH_REC_NO,PO_SHIPMENT_NO,PO_SHIPMENT_LNO,CARTON_NO")
            .Tables("WHTWREC8").Columns.Add("UNITS", GetType(System.Int32), "IIF(ISNULL(PARENT(WHTWREC7_WHTWREC8).CARTON_NO_CLONED_FROM,0) = 0,QTY*IIF(ISNULL(DOZENS,'0')='1',12,1),0)")

            .Tables("WHTWREC7").Columns.Add("STYLES", GetType(System.Int32), "COUNT(CHILD(WHTWREC7_WHTWREC8).STYLE_CODE)")
            .Tables("WHTWREC7").Columns.Add("UNITS", GetType(System.Int32), "SUM(CHILD(WHTWREC7_WHTWREC8).UNITS)")
            .Tables("WHTWREC7").Columns.Add("PPK_INNER_QTY_CALC", GetType(System.Int32), "SUM(CHILD(WHTWREC7_WHTWREC8).PPK_INNER_QTY)")

            .Tables("WHTWREC8").Columns.Add("CARTONS", GetType(System.Int32), "PARENT(WHTWREC7_WHTWREC8).CARTONS")
            .Tables("WHTWREC8").Columns.Add("TOTAL_UNITS", GetType(System.Int32), "ISNULL(UNITS,0) * ISNULL(CARTONS,0)")
            .Tables("WHTWREC7").Columns.Add("TOTAL_UNITS", GetType(System.Int32), "SUM(CHILD(WHTWREC7_WHTWREC8).TOTAL_UNITS)")
            .Tables("WHTWREC7").Columns.Add("STYLE_CODE_1", GetType(System.String), "MIN(CHILD(WHTWREC7_WHTWREC8).STYLE_CODE)")
            .Tables("WHTWREC7").Columns.Add("COLOR_CODE_1", GetType(System.String), "MIN(CHILD(WHTWREC7_WHTWREC8).COLOR_CODE)")
            .Tables("WHTWREC7").Columns.Add("ITEM_CODE", GetType(System.String), "IIF(ISNULL(PPK_CODE,'')='',ISNULL(STYLE_CODE_1,'') + ISNULL(COLOR_CODE_1,''),PPK_CODE)")
            .Tables("WHTWREC7").Columns.Add("CBM", GetType(System.Decimal), "ISNULL(CARTONS,0) * ISNULL(CARTON_VOLUME,0) / 1000000")
            .Tables("WHTWREC7").Columns.Add("TOTAL_WEIGHT", GetType(System.Decimal), "ISNULL(CARTONS,0) * ISNULL(CARTON_WEIGHT,0)")
            .Tables("WHTWREC8").Columns.Add("CBM", GetType(System.Decimal), "IIF(ISNULL(PARENT(WHTWREC7_WHTWREC8).TOTAL_UNITS,0) = 0, 0, ISNULL(TOTAL_UNITS,0) * ISNULL(PARENT(WHTWREC7_WHTWREC8).CBM,0) / ISNULL(PARENT(WHTWREC7_WHTWREC8).TOTAL_UNITS,0))")
            .Tables("WHTWREC8").Columns.Add("TOTAL_UNITS_REC", GetType(System.Int32), "ISNULL(QTY,0) * ISNULL(PARENT(WHTWREC7_WHTWREC8).BAR_CODES,0)")
            .Tables("WHTWREC7").Columns.Add("UNITS_REC", GetType(System.Int32), "SUM(CHILD(WHTWREC7_WHTWREC8).TOTAL_UNITS_REC)")
            Create_Relation("POTSHIPX", "WHTWREC7", "PO_SHIPMENT_NO,PO_SHIPMENT_LNO")
            .Tables("POTSHIPX").Columns.Add("CARTONS_REC", GetType(System.Int32), "ISNULL(SUM(CHILD(POTSHIPX_WHTWREC7).BAR_CODES),0)")
            .Tables("POTSHIPX").Columns.Add("CARTONS_VAR", GetType(System.Int32), "ISNULL(PO_SHIP_CTNS,0) - ISNULL(CARTONS_REC,0)")

            Create_TDA(.Tables.Add, "POTSHIP1", "*", 1, False)
            '    Create_TDA(.Tables.Add, "POTSHIP2", "*", 2, True)

            ASCMAIN1.sql = "Select * from WHTBARC1 where TRAN_TYPE = 'W' and TRAN_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTBARC1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select * from WHTBARC1 where TRAN_TYPE = 'W' and TRAN_NO = :PARM1"
            Create_TDA(.Tables.Add, "WHTBARCC", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from WHTBARC0 where LOAD_NO in (Select Distinct LOAD_NO from WHTBARC1 where TRAN_TYPE = 'W' and TRAN_NO = :PARM1)"
            Create_TDA(.Tables.Add, "WHTBARC0", "**", 0, True, "V", 1)
            .Tables("WHTBARC0").Columns.Add("LOAD_LOCKED")
            .Tables("WHTBARC0").Columns("LOAD_LOCKED").DefaultValue = "0"

            Create_TDA(.Tables.Add, "WHTWREC1", "*", 1, True)

            Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(.Tables.Add, "WHTMOVE2", "*")

            Create_Relation("WHTBARC0", "POTSHIPC", "LOAD_NO")
            .Tables("WHTBARC0").Columns.Add("QTY", GetType(System.Int32), "SUM(CHILD(WHTBARC0_POTSHIPC).QTY)")
            .Tables("POTSHIPC").Columns.Add("LOAD_STATUS", GetType(System.String), "PARENT(WHTBARC0_POTSHIPC).LOAD_STATUS")
        End With
        Fill_Records("ICTWHSE1")
        grdPOTSHIPX.DataSource = dst.Tables("POTSHIPX")
        grdPOTSHIP3.DataSource = dst.Tables("POTSHIP3")
        grdWHTWREC7.DataSource = dst.Tables("WHTWREC7")
        grdWHTWREC8.DataSource = dst.Tables("WHTWREC8")
        grdPOTSHIPC.DataSource = dst.Tables("POTSHIPC")
        grdWHTBARC0.DataSource = dst.Tables("WHTBARC0")
        grdWHTWRECX.DataSource = dst.Tables("WHTWRECX")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")

        Create_Summary(grdPOTSHIP3, "PO_SHIPMENT_LNO", "Count")
        Create_Summary(grdPOTSHIP3, New String() {"PO_QTY_SHP", "UNITS"})

        Create_Summary(grdWHTWREC7, "CARTON_NO", "Count")
        Create_Summary(grdWHTWREC7, New String() {"CARTONS", "BAR_CODES", "UNITS", "TOTAL_UNITS", "UNITS_REC", "CARTON_VOLUME", "CBM", "TOTAL_WEIGHT"})

        Create_Summary(grdWHTWREC8, "STYLE_CODE", "Count")
        Create_Summary(grdWHTWREC8, New String() {"QTY", "UNITS", "TOTAL_UNITS", "CBM"})

        Create_Summary(grdPOTSHIPC, New String() {"QTY"})
        Create_Summary(grdWHTBARC0, New String() {"QTY"})

        ASCMAIN1.Add_Value_List(grdPOTSHIPX, "WHSE_CODE", "Select WHSE_CODE, WHSE_CODE from ICTWHSE1 Where WHSE_LOCATOR = '1'")
        ASCMAIN1.Add_Value_List(grdWHTWRECX, "WH_REC_STATUS", Nothing, New String() {":", "C:Complete", "V:Reversed", "P:In Process", "R:Received"})
        ASCMAIN1.Add_Value_List(grdPOTSHIPC, "LOAD_STATUS", Nothing, New String() {":", "P:Pending", "L:Locked", "A:Moved"})


        With grdWHTBARC0.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdWHTWREC7.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        With grdPOTSHIPC.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            If InquiryMode Then
                .AllowDelete = DefaultableBoolean.False
            Else
                .AllowDelete = DefaultableBoolean.True
            End If
            .AllowUpdate = DefaultableBoolean.False
        End With



        With grdWHTWREC7.DisplayLayout.Bands(0)
            For Each COLUMN_NAME In New String() {"PO_SHIPMENT_LNO", "CARTON_NO", "PPK_CODE", "STYLES", "UNITS", "TOTAL_UNITS", "PPK_INNER_QTY_CALC"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
            .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        With grdWHTWREC7.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"CARTON_NO", "CARTONS", "BAR_CODES", "CUSTOM_PPK", "PPK_CODE", "PO_QTY_PER_CTN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.MediumAquamarine
                ElseIf New String() {"STYLE_CODE", "COLOR_CODE", "PPK_INNER_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.BurlyWood
                ElseIf New String() {"CARTON_COMMENTS"}.Contains(gcol.Key) Then
                    'gcol.Header.Appearance.BackColor2 = Drawing.Color.BurlyWood
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        grdWHTWREC8.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdWHTWREC8.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdWHTWREC8.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

        With grdWHTWREC8.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"QTY", "DOZENS", "PPK_INNER_QTY"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"CARTON_NO", "STYLE_CODE", "COLOR_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Lime
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME In New String() {"PO_SHIPMENT_LNO", "CARTON_NO", "STYLE_CODE", "COLOR_CODE", "UNITS", "TOTAL_UNITS"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        cbeReceipts.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "' order by OPS_YYYYPP DESC")
        cbeReceipts.SelectedItem = cbeReceipts.Items(0)
        cbeReceipts2.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeReceipts2.SelectedItem = cbeReceipts2.Items(0)

        Show_Filter(grdWHTWREC7, True)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"

                WH_REC_NO = Absx1.txtFor("WH_REC_NO").Text
                rowWHTWREC1 = Fill_Record("WHTWREC1", WH_REC_NO)
                If rowWHTWREC1 Is Nothing Then
                    EMsg &= vbCr & "No Record of Warehouse Receipt " & WH_REC_NO
                Else
                    WHSE_CODE = rowWHTWREC1.Item("WHSE_CODE") & ""
                    PO_SHIPMENT_NO = rowWHTWREC1.Item("PO_SHIPMENT_NO") & ""
                    CONTAINER_NO = rowWHTWREC1.Item("CONTAINER_NO") & ""

                    If eItemKey = "Edit" Then
                        If Not ASCMAIN1.Logical_Lock("WHTWREC1", WH_REC_NO) Then Exit Sub
                        If Not ASCMAIN1.Logical_Open("POTSHIP1", PO_SHIPMENT_NO) Then Exit Sub

                        If rowWHTWREC1.Item("WH_REC_STATUS") = "C" Then
                            If ASCMAIN1.CLIENT = "VAN" And ASCMAIN1.Running_in_VS Then

                                Stop ' NOT SURE IF THIS ROUTINE IS COMPLETE
                                ' - WHAT IF THE LOADS WERE LOCKED - DOES THAT MEAN THAT THE BARCODES HAVE BEEN MOVED?
                                ' - WHAT IF THERE WAS ANY WH ACTIVITY WITH THESE BARCODES

                                If MsgBox("Receipt is marked as Complete." & vbCrLf & "Do you want to re-open it so that it is not Complete?", _
                                          MsgBoxStyle.YesNo, "Management Option") = MsgBoxResult.Yes Then
                                    BeginTrans()

                                    ASCMAIN1.sql = "Update WHTBARC0 Set LOAD_STATUS = 'P'" & vbCrLf _
                                        & " where LOAD_STATUS = 'R' and TRAN_TYPE = 'W' and TRAN_NO = '" & WH_REC_NO & "'"
                                    ASCDATA1.ExecuteSQL()

                                    ASCMAIN1.sql = "Update WHTWREC1 Set WH_REC_STATUS = 'P' Where WH_REC_STATUS = 'C' and WH_REC_NO = '" & WH_REC_NO & "'"
                                    ASCDATA1.ExecuteSQL()

                                    ASCMAIN1.sql = "Update WHTWREC8 Set QTY = -1 * QTY where  WH_REC_NO = '" & WH_REC_NO & "'"
                                    ASCDATA1.ExecuteSQL()

                                    ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                                                       New Object() {"W", WH_REC_NO, ASCMAIN1.SESSION_NO},
                                                       New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

                                    ASCMAIN1.sql = "Update WHTWREC8 Set QTY = -1 * QTY where  WH_REC_NO = '" & WH_REC_NO & "'"
                                    ASCDATA1.ExecuteSQL()

                                    Record_Event("REOPEN", "Receipt Re-Opened", True)

                                    ' NOT SURE IF WE SHOULD DO THIS
                                    '  Make_Locked_Loads_Available_for_Putaway()

                                    CommitTrans()

                                    rowWHTWREC1 = Fill_Record("WHTWREC1", WH_REC_NO)
                                End If
                            End If
                        End If

                        If rowWHTWREC1.Item("WH_REC_STATUS") = "C" Then
                            EMsg &= vbCr & "Receipt is Complete. No further maintenance permitted"
                        ElseIf rowWHTWREC1.Item("WH_REC_STATUS") = "R" Then
                            EMsg &= vbCr & "Receipt is Received. No further maintenance permitted"
                        ElseIf rowWHTWREC1.Item("WH_REC_STATUS") = "P" Then
                            'this is an ok status
                        ElseIf rowWHTWREC1.Item("WH_REC_STATUS") = "V" Then
                            EMsg &= vbCr & "Receipt is Reversed. No further maintenance permitted"
                        End If

                        If WHSE_CODE <> User_Whse_Code Then
                            MsgBox("This Receipt is marked for " & WHSE_CODE & ", You are Setup to Receive in Whse " & User_Whse_Code, vbOKOnly, "Cannot Receive")
                            Exit Sub
                        End If
                    End If

                    rowPOTSHIP1 = Fill_Record("POTSHIP1", PO_SHIPMENT_NO)
                    If rowPOTSHIP1.Item("LP_STATUS") & "" <> "1" Then
                        EMsg &= vbCr & "Shipment " & PO_SHIPMENT_NO & " has not been Released for Warehouse Receipt, please refresh screen"
                    End If
                End If

                If EMsg = "" Then
                    For Each row As DataRow In dst.Tables("POTSHIPX").Select("")
                        row.Item("SELECTED") = "0"
                    Next

                    Dim PO_SHIPMENT_LNOs As New List(Of Integer)
                    PO_SHIPMENT_LNOs.Clear()
                    ASCMAIN1.sql = "Select * from POTSHIP2 where WH_REC_NO = '" & WH_REC_NO & "'"
                    For Each rowPOTSHIP2 As DataRow In ASCDATA1.GetDataTable.Select("", "PO_SHIPMENT_LNO")
                        PO_SHIPMENT_LNO = Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & "")
                        Dim rowPOTSHIPX As DataRow = dst.Tables("POTSHIPX").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
                        If rowPOTSHIPX Is Nothing Then
                            Fill_Records("POTSHIPX", "", False, sqlPOTSHIPX & " and POTSHIP2.PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & Val(rowPOTSHIP2.Item("PO_SHIPMENT_LNO") & ""))
                            rowPOTSHIPX = dst.Tables("POTSHIPX").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
                        End If
                        If rowPOTSHIPX Is Nothing Then
                            EMsg &= vbCr & "Cannot Connect Shipment " & PO_SHIPMENT_NO & ", Line " & CStr(PO_SHIPMENT_LNO)
                        Else
                            rowPOTSHIPX.Item("SELECTED") = "1"
                            PO_SHIPMENT_LNOs.Add(PO_SHIPMENT_LNO)
                            If eItemKey = "Edit" Then
                                If Not ASCMAIN1.Logical_Lock("POTSHIP2", PO_SHIPMENT_NO & ":" & CStr(PO_SHIPMENT_LNO)) Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    Next
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If
            Case "Delete"
                Dim iResponse As MsgBoxResult = MsgBox("Are You Sure You Wish to Delete?", MsgBoxStyle.YesNo, "Pay Attention!")
                If iResponse = MsgBoxResult.No Then
                    EMsg &= vbCr & "Delete Aborted"
                End If
            Case "Receive"
                If tabMain.SelectedTab.Key <> "Shipments" Then
                    Exit Sub
                End If

                WHSE_CODE = ""

                For Each row As DataRow In dst.Tables("POTSHIPX").Select("")
                    row.Item("SELECTED") = "0"
                Next

                Dim PO_SHIPMENT_LNOs As New List(Of Integer)


                If double_clicked Then
                    double_clicked = False

                    For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
                                                                                     & " And CONTAINER_NO = '" & CONTAINER_NO & "'")

                        WHSE_CODE = rowPOTSHIPX.Item("WHSE_CODE")
                        'CONTAINER_NO = rowPOTSHIPX.Item("CONTAINER_NO")
                        rowPOTSHIPX.Item("SELECTED") = "1"
                        PO_SHIPMENT_LNOs.Add(rowPOTSHIPX.Item("PO_SHIPMENT_LNO"))
                    Next

                Else
                    CONTAINER_NO = ""
                    If grdPOTSHIPX.Selected.Rows.Count = 0 Then
                        If grdPOTSHIPX.ActiveRow IsNot Nothing And grdPOTSHIPX.ActiveRow.IsDataRow Then
                            grdPOTSHIPX.ActiveRow.Selected = True
                        End If
                    End If

                    If grdPOTSHIPX.Selected.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Containers Selected"
                    Else
                        Dim PO_SHIPMENT_NO_1ST As String = ""
                        For Each grow As UltraWinGrid.UltraGridRow In grdPOTSHIPX.Selected.Rows
                            If grow.IsDataRow Then
                                PO_SHIPMENT_NO = grow.Cells("PO_SHIPMENT_NO").Value
                                If PO_SHIPMENT_NO_1ST = "" Then
                                    PO_SHIPMENT_NO_1ST = PO_SHIPMENT_NO
                                    WHSE_CODE = grow.Cells("WHSE_CODE").Value
                                Else
                                    If PO_SHIPMENT_NO_1ST <> PO_SHIPMENT_NO Then
                                        EMsg &= vbCr & "Cannot Select Containers from Different Shipments"
                                        Exit For
                                    End If
                                End If
                                PO_SHIPMENT_LNO = grow.Cells("PO_SHIPMENT_LNO").Value
                                PO_SHIPMENT_LNOs.Add(PO_SHIPMENT_LNO)
                                CONTAINER_NO = grow.Cells("CONTAINER_NO").Value

                                Dim rowPOTSHIPX As DataRow = dst.Tables("POTSHIPX").Rows.Find(New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
                                rowPOTSHIPX.Item("SELECTED") = "1"
                            End If
                        Next
                        If PO_SHIPMENT_NO_1ST = "" Then
                            EMsg &= vbCr & "No Containers Selected"
                        Else
                            PO_SHIPMENT_NO = PO_SHIPMENT_NO_1ST
                        End If

                    End If

                End If

                If WHSE_CODE = "" Then
                    EMsg &= vbCr & "Cannot Determine Warehouse from Shipments Selected"
                Else
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & "Cannot Determine Warehouse from Shipments Selected"
                    Else
                        If rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Or rowICTWHSE1.Item("WHSE_CTN_CTL") & "" <> "C" Then
                            EMsg &= vbCr & "Cannot Receive using this Function into a Warehouse that is not using ABS WMS for Carton Control"
                        End If
                    End If
                End If

                If EMsg = "" Then
                    ASCMAIN1.sql = "Select PO_SHIPMENT_LNO from POTSHIP2 where CONTAINER_NO = :PARM1 And PO_SHIP_STATUS <> 'C'"
                    For Each ROW As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New Object() {CONTAINER_NO}).Select("")
                        Dim PO_SHIPMENT_LNO As Integer = Val(ROW.Item("PO_SHIPMENT_LNO") & "")
                        If PO_SHIPMENT_LNOs.Contains(PO_SHIPMENT_LNO) Then
                            PO_SHIPMENT_LNOs.Remove(PO_SHIPMENT_LNO)
                        Else
                            EMsg &= vbCrLf & "You must select all PO Shipment Lines for a Container"
                            PO_SHIPMENT_LNOs.Clear()
                            Exit For
                        End If
                    Next
                    If PO_SHIPMENT_LNOs.Count <> 0 Then
                        EMsg &= vbCrLf & "You cannot mix Containers in a single receipt"
                    End If
                End If

                If EMsg = "" Then
                    If eItemKey = "Receive" Then
                        If Not ASCMAIN1.Logical_Open("POTSHIP1", PO_SHIPMENT_NO) Then Exit Sub

                        For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("SELECTED = '1'")
                            Dim PO_SHIPMENT_LNO As Integer = Val(rowPOTSHIPX.Item("PO_SHIPMENT_LNO") & "")
                            If Not ASCMAIN1.Logical_Lock("POTSHIP2", PO_SHIPMENT_NO & ":" & CStr(PO_SHIPMENT_LNO)) Then
                                Exit Sub
                            End If
                        Next


                        rowPOTSHIP1 = Fill_Record("POTSHIP1", PO_SHIPMENT_NO)
                        If rowPOTSHIP1.Item("LP_STATUS") & "" <> "1" Then
                            EMsg &= vbCr & "Shipment has not been Released for Warehouse Receipt"
                        End If

                        ASCMAIN1.sql = "SELECT * FROM WHTWREC1 " & vbCrLf _
                        & " WHERE PO_SHIPMENT_NO = :PARM1 AND CONTAINER_NO = :PARM2"
                        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() _
                                                                {PO_SHIPMENT_NO, CONTAINER_NO})
                        If row IsNot Nothing Then
                            EMsg &= vbCr & "Shipment has already been received, please refresh your screen"
                        End If

                    End If
                End If


                If EMsg = "" Then
                    ASCDATA1.DeleteRows("POTSHIPX", "SELECTED = '0'")
                End If

            Case "Update"

                EMsg &= Check_Deleted_LPNs()


                If dst.Tables("WHTWREC7").Select("ISNULL(CARTONS,0) <> ISNULL(BAR_CODES,0)").Length <> 0 Then
                    ' EMsg &= vbCr & "Some Carton Types do not match the count of LPNs generated"
                End If
                If Absx1.txtFor("PO_SHIP_VIA_DESC").Text = "" Then
                    EMsg &= vbCr & "Invalid Carrier (Check Receipts Header tab)"
                End If
                If Absx1.txtFor("USER_NAME").Text = "" Then
                    EMsg &= vbCr & "Invalid Unloader (Check Receipts Header tab)"
                End If
                If Absx1.txtFor("CONTAINER_SEAL_NO").Text = "" Then
                    EMsg &= vbCr & "Invalid Container Seal No (Check Receipts Header tab)"
                End If
                If Absx1.txtFor("TRAILER_NO").Text = "" Then
                    EMsg &= vbCr & "Invalid Trailer No (Check Receipts Header tab)"
                End If
                If dteWH_DATE_DELIVERED.Text = "" Then
                    EMsg &= vbCr & "Invalid Delivered Date (Check Receipts Header tab)"
                End If
                If dteWH_DATE_UNLOADED.Text = "" Then
                    EMsg &= vbCr & "Invalid Unloaded Date (Check Receipts Header tab)"
                End If

                Dim LPN_Exists_Msg As String = ""
                Dim No_Load_No_Msg As String = ""
                For Each rowWHTBARC1 As DataRow In dst.Tables("WHTBARC1").Select("")
                    LookUp("WHTBARC1", rowWHTBARC1.Item("BAR_CODE"))
                    If cdr IsNot Nothing Then

                        If (cdr.Item("PO_SHIPMENT_NO") & "" = rowWHTBARC1.Item("PO_SHIPMENT_NO") & "") And
                            (cdr.Item("PO_SHIPMENT_LNO") & "" = rowWHTBARC1.Item("PO_SHIPMENT_LNO") & "") Then
                        Else
                            LPN_Exists_Msg &= vbCr & "  LPN: " & cdr.Item("BAR_CODE") & " Received on PO " & cdr.Item("PO_ORDER_NO") & " on " & cdr.Item("PO_DATE_RECEIVED")
                        End If
                    End If
                    If rowWHTBARC1.Item("LOAD_NO") & "" = "" Then
                        No_Load_No_Msg = rowWHTBARC1.Item("BAR_CODE")
                    End If
                Next
                If LPN_Exists_Msg <> "" Then
                    EMsg &= vbCr & "LPN's already exists in Database" & LPN_Exists_Msg
                End If
                If No_Load_No_Msg <> "" Then
                    EMsg &= vbCr & "No Load assigned to Case Id " & No_Load_No_Msg
                End If

                If EMsg = "" Then
                    Dim OOB_Msg As String = ""
                    Dim Line_Ctr As Integer = 0
                    For Each rowWHTWREC7_OOB As DataRow In dst.Tables("WHTWREC7").Select("ISNULL(CARTONS,0) <> ISNULL(BAR_CODES,0)")
                        OOB_Msg &= vbCrLf & "Line Ship Qty: " & Val(rowWHTWREC7_OOB.Item("CARTONS") & "") & "  Received Qty: " & Val(rowWHTWREC7_OOB.Item("BAR_CODES") & "")
                        Line_Ctr += 1
                        If Line_Ctr = 10 Then
                            OOB_Msg &= vbCrLf & "Too many disrepencies to list all"
                            Exit For
                        End If
                    Next
                    If OOB_Msg <> "" Then
                        Dim iResponse As MsgBoxResult = MsgBox("There are discrepencies with line below, ok to proceed?" & OOB_Msg, MsgBoxStyle.YesNo, "Pay Attention!")
                        If iResponse = MsgBoxResult.No Then
                            EMsg &= vbCr & "Update Aborted"
                        Else
                            Record_Event("DISCR", "Discrepencies acknowledged", True)
                        End If
                    End If
                End If

                If ASCMAIN1.CLIENT = "VAN" Then
                    If EMsg = "" Then
                        If chkComplete.Checked Then
                            Dim WH_PARM_REC_VAR_WARNING_UNITS As Integer = Val(ROWs("WHTPARM1").Item("WH_PARM_REC_VAR_WARNING_UNITS") & "")
                            Dim WH_PARM_REC_VAR_WARNING_PWD As String = ROWs("WHTPARM1").Item("WH_PARM_REC_VAR_WARNING_PWD") & ""
                            Dim TOTAL_UNITS As Int64 = Val(dst.Tables("WHTWREC8").Compute("SUM (TOTAL_UNITS)", "") & "")
                            Dim TOTAL_UNITS_REC As Int64 = Val(dst.Tables("WHTWREC8").Compute("SUM (TOTAL_UNITS_REC)", "") & "")
                            Dim TOTAL_UNITS_VAR = TOTAL_UNITS - TOTAL_UNITS_REC

                            If System.Math.Abs(TOTAL_UNITS_VAR) > WH_PARM_REC_VAR_WARNING_UNITS Then
                                Dim frmASFMSGBF As New ASFMSGBF
                                If frmASFMSGBF.Get_txt_from_User("Enter Password", "Large Receiving Variance", True) <> WH_PARM_REC_VAR_WARNING_PWD Then
                                    EMsg &= vbCr & "Invalid Password - please see a Manager"
                                End If
                            End If
                        End If
                    End If
                End If

            Case "Reverse"
                EMsg &= vbCr & "Currently un-available"
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub




    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Receive"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Done"
                Mode_Settings(False)

            Case "Refresh"
                Load_POTSHIPX()

            Case "Reverse"

                Reverse()
                Mode_Settings(False)
            Case "Delete"
                Delete_Record()
                Mode_Settings(False)
            Case "Print"
                Print_Record()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Receive").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Reverse").Settings.Enabled = iScreenMode

                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("View").Visible = IIf(tabMain.SelectedTab.Key = "Receipts" And MENU_ITEM_OBJECT = "WHFWRECI", True, False) 'InquiryMode ' (EntryMode = "L" Or Not ScreenMode)
                    .Items("Done").Visible = IIf(tabMain.SelectedTab.Key = "Receipts" And MENU_ITEM_OBJECT = "WHFWRECI", True, False) 'InquiryMode
                    .Items("Reverse").Visible = IIf(tabMain.SelectedTab.Key = "Receipts" And MENU_ITEM_OBJECT = "WHFWREC1", True, False)

                    .Items("Refresh").Visible = Not ScreenMode

                    .Items("Receive").Visible = Not InquiryMode
                    .Items("Update").Visible = Not InquiryMode
                    .Items("Cancel").Visible = IIf(MENU_ITEM_OBJECT = "WHFWREC1", True, False) 'Not InquiryMode
                    .Items("Delete").Visible = Not InquiryMode
                    .Items("Print").Visible = ScreenMode
                End With
                .Groups("Receipts History").Visible = InquiryMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        splLPN.Panel1Collapsed = InquiryMode


        lblStatus.Visible = ScreenMode


        tabMain.Visible = Not tf
        splPOTSHIPX.Visible = tf

        grd_Update()

        Setup_Receipts_History()

        If ScreenMode Then
            If rowWHTWREC1.Item("WH_REC_STATUS") = "C" Then
                lblStatus.Text = "Complete"
            ElseIf rowWHTWREC1.Item("WH_REC_STATUS") = "R" Then
                lblStatus.Text = "Received"
            ElseIf rowWHTWREC1.Item("WH_REC_STATUS") = "P" Then
                lblStatus.Text = "In Process"
            ElseIf rowWHTWREC1.Item("WH_REC_STATUS") = "V" Then
                lblStatus.Text = "Reversed"
            End If

            UltraExplorerBar1.Groups("Screen Control").Items("Reverse").Visible = IIf(tabMain.SelectedTab.Key = "Receipts" _
                                                                                      And MENU_ITEM_OBJECT = "WHFWREC1" _
                                                                                      And rowWHTWREC1.Item("WH_REC_STATUS") & "" = "C", True, False)


            UltraExplorerBar1.Groups("Screen Control").Items("Delete").Visible = IIf(tabMain.SelectedTab.Key = "Receipts" _
                                                                                      And MENU_ITEM_OBJECT = "WHFWREC1" _
                                                                                      And rowWHTWREC1.Item("WH_REC_STATUS") & "" = "P", True, False)

            chkComplete.Checked = False
            chkComplete.Visible = (rowWHTWREC1.Item("WH_REC_STATUS") = "P")
            chkEmail.Visible = (rowWHTWREC1.Item("WH_REC_STATUS") = "P") And (rowWHTWREC1.Item("WH_REC_EMAIL_SENT") & "" <> "1")

            Show_Filter(grdPOTSHIPX, False)
            grdPOTSHIPX.Parent = splPOTSHIPX.Panel1
            grdPOTSHIPX.Selected.Rows.Clear()

            'With grdPOTSHIPX.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '    .AllowDelete = DefaultableBoolean.False
            '    .AllowUpdate = DefaultableBoolean.False
            'End With

            Set_Read_Only(grpHeader, EntryMode <> "N" And EntryMode <> "E")

            'tabContainerDetails.SelectedTab = tabContainerDetails.Tabs("Carton Type Summary")
            tabContainerDetails.SelectedTab = tabContainerDetails.Tabs("Receipts Header")
            Setup_LPN_Entry()

        Else
            Clear_Record()
            UltraExplorerBar1.Groups("LPN Entry").Visible = False
            UltraExplorerBar1.Groups("Style Lookup").Visible = False
            Show_Filter(grdPOTSHIPX, True)
            grdPOTSHIPX.Parent = tabMain.Tabs("Shipments").TabPage

            'With grdPOTSHIPX.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '    .AllowDelete = DefaultableBoolean.False
            '    .AllowUpdate = DefaultableBoolean.True
            'End With
            'grdPOTSHIPX.DisplayLayout.Bands(0).Columns("WHSE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            'With grdPOTSHIPX.DisplayLayout.Bands(0)
            '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
            '        If New String() {"WHSE_CODE"}.Contains(gcol.Key) Then
            '            gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            '        Else
            '            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            '        End If
            '    Next
            'End With


            If InquiryMode Then
                tabMain.Tabs("Shipments").Visible = Not InquiryMode
            End If
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTSHIPX", "POTSHIP3", "WHTWREC7", "WHTWREC8", "POTSHIPC", "WHTBARC0", "WHTBARC1", "WHTBARCC", "WHTWREC1", "TATEVNT1", "WHTMOVE1", "WHTMOVE2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If Absx1.txtFor("WHSE_CODE").Text = "" Then
            If dst.Tables("ICTWHSE1").Rows.Count = 1 Then
                Absx1.txtFor("WHSE_CODE").Text = dst.Tables("ICTWHSE1").Rows(0).Item("WHSE_CODE")
            End If
        End If

        'If ASCMAIN1.USER_ID = "gcv" Then
        '    ASCMAIN1.sql = "Select Distinct WHSE_TRAN_NO from whtmove2 where load_no_FROM in (Select LOAD_NO from whtbarc0 where whse_code = 'NJT')"
        '    For Each rowSOTWORK3 As DataRow In ASCDATA1.GetDataTable.Rows
        '        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
        '               New Object() {rowSOTWORK3.Item("WHSE_TRAN_NO"), 0, 1},
        '               New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})
        '    Next
        'End If




        chkRemoteLoc.Checked = False
        txtNewContainer.Text = ""

        txtBAR_CODE.Text = ""
        txtBAR_CODE2.Text = ""
        txtBAR_CODE_WITH.Text = ""
        txtLOAD_NO.Text = ""
        Deleted_BarCodes = ""
        tabMain.SelectedTab = IIf(InquiryMode = False, tabMain.Tabs("Shipments"), tabMain.Tabs("Receipts"))

        Load_POTSHIPX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)
        If EntryMode = "N" Then
            WH_REC_NO = ASCMAIN1.Next_Control_No("WHTWREC1.WH_REC_NO")

            rowWHTWREC1 = dst.Tables("WHTWREC1").NewRow
            rowWHTWREC1.Item("WH_REC_NO") = WH_REC_NO
            rowWHTWREC1.Item("WH_DATE_RECEIVED") = DATETIME_STAMP.Date
            rowWHTWREC1.Item("WH_REC_STATUS") = "P"
            rowWHTWREC1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowWHTWREC1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowWHTWREC1.Item("WHSE_CODE") = WHSE_CODE
            rowWHTWREC1.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTWREC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTWREC1.Item("TRAILER_NO") = CONTAINER_NO
            rowWHTWREC1.Item("CONTAINER_NO") = CONTAINER_NO
            dst.Tables("WHTWREC1").Rows.Add(rowWHTWREC1)
        Else
            rowWHTWREC1 = Fill_Record("WHTWREC1", WH_REC_NO)
            PO_SHIPMENT_NO = rowWHTWREC1.Item("PO_SHIPMENT_NO")
            rowPOTSHIP1 = Fill_Record("POTSHIP1", PO_SHIPMENT_NO)
            WHSE_CODE = rowWHTWREC1.Item("WHSE_CODE")
            ASCMAIN1.sql = sqlPOTSHIPX & " and POTSHIP2.WH_REC_NO = '" & WH_REC_NO & "'"
            Fill_Records("POTSHIPX", "", True, ASCMAIN1.sql)
        End If

        With grdPOTSHIPX.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("WHSE_CODE", False)
            .SortedColumns.Add("PO_SHIPMENT_NO", False)
        End With

        Sort_grdColumns(grdPOTSHIPX, "PO_SHIPMENT_LNO")

        dst.Tables("POTSHIP3").Rows.Clear()
        dst.Tables("WHTWREC7").Rows.Clear()
        dst.Tables("WHTWREC8").Rows.Clear()
        dst.Tables("POTSHIPC").Rows.Clear()
        dst.Tables("WHTBARC0").Rows.Clear()
        dst.Tables("WHTBARC1").Rows.Clear()
        dst.Tables("WHTBARCC").Rows.Clear()

        For Each row As DataRow In dst.Tables("POTSHIPX").Select("")
            PO_SHIPMENT_LNO = Val(row.Item("PO_SHIPMENT_LNO") & "")
            Fill_Records("POTSHIP3", New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO}, False)

            If EntryMode = "N" Then
                ASCMAIN1.sql = "Select POTSHIP7.*, WHTSCSEQ.STYLE_SEQ from POTSHIP7, WHTSCSEQ" & vbCrLf _
                & " Where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                & " And PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO & vbCrLf _
                & "   and POTSHIP7.STYLE_CODE = WHTSCSEQ.STYLE_CODE(+)" & vbCrLf _
                & "   and POTSHIP7.COLOR_CODE = WHTSCSEQ.COLOR_CODE(+)"
                For Each rowPOTSHIP7 As DataRow In ASCDATA1.GetDataTable.Rows
                    Add_WHTWREC7(rowPOTSHIP7, False, 0)
                Next

                ASCMAIN1.sql = "Select * from POTSHIP8" & vbCrLf _
                & " Where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" & vbCrLf _
                & " And PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO
                For Each rowPOTSHIP8 As DataRow In ASCDATA1.GetDataTable.Rows
                    Add_WHTWREC8(rowPOTSHIP8, False)
                Next
            End If
        Next

        Fill_Records("WHTBARC0", New Object() {WH_REC_NO}, False)
        Fill_Records("WHTBARC1", New Object() {WH_REC_NO}, False)
        Fill_Records("WHTBARCC", New Object() {WH_REC_NO}, False)

        If EntryMode = "N" Then
        Else
            ASCMAIN1.sql = "Select WHTWREC7.*, WHTSCSEQ.STYLE_SEQ from WHTWREC7, WHTSCSEQ Where WH_REC_NO = '" & WH_REC_NO & "'" & vbCrLf _
                & "   and WHTWREC7.STYLE_CODE = WHTSCSEQ.STYLE_CODE(+)" & vbCrLf _
                & "   and WHTWREC7.COLOR_CODE = WHTSCSEQ.COLOR_CODE(+)"
            Fill_Records("WHTWREC7", , , ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from WHTWREC8 Where WH_REC_NO = '" & WH_REC_NO & "'"
            Fill_Records("WHTWREC8", , , ASCMAIN1.sql)

            Create_POTSHIPC_VIEW("")

            'For Each rowPOTSHIPC As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTSHIPC").Select(""), "LOAD_NO").Rows
            '    If dst.Tables("WHTBARC0").Select("LOAD_NO = '" & rowPOTSHIPC.Item("LOAD_NO") & "'").Length = 0 Then
            '        ASCDATA1.DeleteRows("POTSHIPC", "LOAD_NO = '" & rowPOTSHIPC.Item("LOAD_NO") & "'")
            '    End If
            'Next

            Fill_Records("TATEVNT1", WH_REC_NO)
            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
        End If

        EnforceConstraints(True)

        Setup_POTSHIPX_Details()

        Sort_grdColumns(grdPOTSHIP3, "PO_SHIPMENT_LNO")

    End Sub

    Sub Create_POTSHIPC_VIEW(SQL As String)

        For Each row As DataRow In dst.Tables("WHTWREC7").Select(SQL)

            PO_SHIPMENT_LNO = Val(row.Item("PO_SHIPMENT_LNO") & "")
            CARTON_NO = Val(row.Item("CARTON_NO") & "")

            Dim BAR_CODE As String = ""
            Dim BAR_CODE2 As String = ""
            Dim LOAD_NO As String = ""
            Dim QTY = 0

            For Each row2 As DataRow In dst.Tables("WHTBARC1").Select("PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO), "BAR_CODE")
                If BAR_CODE <> "" Then
                    If row2.Item("BAR_CODE") = Format(Val(BAR_CODE2) + 1, "".PadLeft(8, "0")) And LOAD_NO = row2.Item("LOAD_NO") & "" Then
                        BAR_CODE2 = row2.Item("BAR_CODE")
                        QTY += 1
                    Else
                        Write_POTSHIPC(BAR_CODE, BAR_CODE2, QTY, LOAD_NO, "")
                        BAR_CODE = ""
                    End If
                End If

                If BAR_CODE = "" Then
                    BAR_CODE = row2.Item("BAR_CODE")
                    BAR_CODE2 = row2.Item("BAR_CODE")
                    LOAD_NO = row2.Item("LOAD_NO") & ""
                    QTY = 1
                End If
            Next
            If BAR_CODE <> "" Then
                Write_POTSHIPC(BAR_CODE, BAR_CODE2, QTY, LOAD_NO, "")
            End If

        Next
    End Sub

    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String, Optional update_database As Boolean = False)

        Dim rowTATEVNT1 As DataRow = dst.Tables("TATEVNT1").NewRow
        rowTATEVNT1.Item("TABLE_NAME") = "WHTWREC1"
        rowTATEVNT1.Item("TABLE_KEY") = WH_REC_NO
        rowTATEVNT1.Item("INIT_DATE") = DATETIME_STAMP
        rowTATEVNT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowTATEVNT1.Item("EVENT_TYPE") = EVENT_TYPE
        rowTATEVNT1.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("TATEVNT1").Rows.Add(rowTATEVNT1)
        If update_database Then Update_Record_TDA("TATEVNT1")

    End Sub

    Sub Add_WHTWREC7(row As DataRow, Cloned_Flag As Boolean, Cloned_From As Integer)
        ' ALMOST identical copy of this method exists in POFSHIP1 - probably should refactor both to POCMAIN1
        Dim rowWHTWREC7 As DataRow = dst.Tables("WHTWREC7").NewRow
        With rowWHTWREC7
            .Item("WH_REC_NO") = WH_REC_NO
            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            .Item("CARTON_NO") = IIf(Cloned_Flag = False, row.Item("CARTON_NO") & "",
                                                        Val(dst.Tables("WHTWREC7").Compute("MAX(CARTON_NO)",
                                                                                           "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
                                                                                            & " And PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO) & "") + 1)
            .Item("CARTONS") = IIf(Cloned_Flag = False, Val(row.Item("CARTONS") & ""), 0)
            .Item("CARTON_COMMENTS") = row.Item("CARTON_COMMENTS") & ""
            .Item("CUSTOM_PPK") = row.Item("CUSTOM_PPK") & ""
            .Item("PPK_CODE") = row.Item("PPK_CODE") & ""
            .Item("PO_QTY_PER_CTN") = Val(row.Item("PO_QTY_PER_CTN") & "")
            .Item("STYLE_CODE") = row.Item("STYLE_CODE") & ""
            .Item("COLOR_CODE") = row.Item("COLOR_CODE") & ""
            .Item("PPK_INNER_QTY") = Val(row.Item("PPK_INNER_QTY") & "")
            .Item("CARTON_DIMS") = row.Item("CARTON_DIMS") & ""
            .Item("CARTON_VOLUME") = Val(row.Item("CARTON_VOLUME") & "")
            .Item("CARTON_WEIGHT") = Val(row.Item("CARTON_WEIGHT") & "")
            .Item("CARTON_NO_CLONED_FROM") = IIf(Cloned_Flag = False, DBNull.Value, Cloned_From)
            .Item("STYLE_SEQ") = row.Item("STYLE_SEQ")
        End With
        dst.Tables("WHTWREC7").Rows.Add(rowWHTWREC7)
    End Sub

    Sub Add_WHTWREC8(row As DataRow, Cloned_Flag As Boolean)
        ' ALMOST identical copy of this method exists in POFSHIP1 - probably should refactor both to POCMAIN1
        Dim rowWHTWREC8 As DataRow = dst.Tables("WHTWREC8").NewRow
        With rowWHTWREC8
            .Item("WH_REC_NO") = WH_REC_NO
            .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            .Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            .Item("CARTON_NO") = IIf(Cloned_Flag = False, row.Item("CARTON_NO") & "",
                                            Val(dst.Tables("WHTWREC8").Compute("MAX(CARTON_NO)",
                                                                               "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
                                                                                & " And PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO) & "") + 1)
            .Item("STYLE_CODE") = row.Item("STYLE_CODE") & ""
            .Item("COLOR_CODE") = row.Item("COLOR_CODE") & ""
            .Item("QTY") = Val(row.Item("QTY") & "")
            .Item("DOZENS") = row.Item("DOZENS") & ""
            .Item("PPK_INNER_QTY") = Val(row.Item("PPK_INNER_QTY") & "")
            .Item("QTY_SHP") = IIf(Cloned_Flag = False, Val(row.Item("QTY") & ""), 0)

        End With
        dst.Tables("WHTWREC8").Rows.Add(rowWHTWREC8)
    End Sub

    Sub Delete_Record()
        BeginTrans()
        For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("")
            PO_SHIPMENT_LNO = Val(rowPOTSHIPX.Item("PO_SHIPMENT_LNO") & "")
            ASCMAIN1.sql = "Update POTSHIP2 Set WH_REC_NO = NULL where PO_SHIPMENT_NO = :PARM1 and PO_SHIPMENT_LNO = :PARM2"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VN", New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
        Next

        For Each TABLE_NAME As String In New String() _
        {"WHTBARC1", "WHTBARC0", "WHTWREC1", "WHTWREC7", "WHTWREC8"}
            ASCMAIN1.sql = "Delete from " & TABLE_NAME & " Where " & IIf(TABLE_NAME = "WHTBARC1" Or TABLE_NAME = "WHTBARC0", "TRAN_NO", "WH_REC_NO") & " = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", WH_REC_NO)
        Next
        CommitTrans("Receipt Deleted")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim warning_issued As Integer = 0

        'this is used to clear out any LPN's that were deleted on receipt but load was locked
        If Deleted_BarCodes <> "" Then
            ASCMAIN1.sql = "Insert into WHTLOCO1 " _
            & " 	Select * from WHTLOCB1 " _
            & " 	Where BAR_CODE in (" & Mid(Deleted_BarCodes, 2) & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Insert into WHTLOCO2 " _
            & " 	Select * from WHTLOCB2 " _
            & " 	Where BAR_CODE in (" & Mid(Deleted_BarCodes, 2) & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = " 	Delete from WHTLOCB1 " _
            & " 	Where BAR_CODE in (" & Mid(Deleted_BarCodes, 2) & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = " 	Delete from WHTLOCB2 " _
            & " 	Where BAR_CODE in (" & Mid(Deleted_BarCodes, 2) & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If

        'this is here to update the load no for LPN's that might have been moved off a load and have a new load no 
        'should I change  to " and LOAD_STATUS not in ('P','L')" ?
        'ASCMAIN1.sql = " Select * from WHTBARC0,  WHTBARC1" _
        '& " Where WHTBARC0.LOAD_NO = WHTBARC1.LOAD_NO" _
        '& " and WHTBARC1.TRAN_NO = '" & WH_REC_NO & "'" _
        '& " and LOAD_STATUS = 'A'"
        'For Each rowWHTBARCL As DataRow In ASCDATA1.GetDataTable.Rows
        '    For Each rowWHTBARC1 As DataRow In dst.Tables("WHTBARC1").Select("BAR_CODE = '" & rowWHTBARCL.Item("BAR_CODE") & "'")
        '        rowWHTBARC1.Item("LOAD_NO") = rowWHTBARCL.Item("LOAD_NO")
        '    Next
        'Next
        '-----------------

        For Each rowWHTBARC1 As DataRow In dst.Tables("WHTBARC1").Select("")
            rowWHTBARC1.Item("TRAN_TYPE") = "W"
            rowWHTBARC1.Item("TRAN_NO") = WH_REC_NO
            rowWHTBARC1.Item("PO_DATE_RECEIVED") = dteWH_DATE_RECEIVED.Value
            rowWHTBARC1.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTBARC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTBARC1.Item("STATUS_CODE") = "R"

            If warning_issued < 3 Then
                Dim rowWHTWREC7 As DataRow = dst.Tables("WHTWREC7").Rows.Find(New Object() {WH_REC_NO, rowWHTBARC1.Item("PO_SHIPMENT_NO"), rowWHTBARC1.Item("PO_SHIPMENT_LNO"), rowWHTBARC1.Item("CARTON_NO")})
                If rowWHTWREC7 IsNot Nothing Then
                    Dim PPK_CODE As String = rowWHTWREC7.Item("PPK_CODE") & ""
                    If rowWHTBARC1.Item("PPK_CODE") & "" <> PPK_CODE Then
                        MsgBox("Issue with Pre-pack Code " & PPK_CODE & " and Case ID " & rowWHTBARC1.Item("BAR_CODE"), MsgBoxStyle.OkOnly, "Please Alert ABS")
                        warning_issued += 1
                    End If
                End If
            End If
        Next

        For Each rowWHTWREC7 As DataRow In dst.Tables("WHTWREC7").Select("")
            rowWHTWREC7.Item("CARTONS_RECEIVED") = Val(rowWHTWREC7.Item("BAR_CODES") & "")
        Next

        Dim CONTAINER_NOs As New List(Of String)
        For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("")
            PO_SHIPMENT_LNO = Val(rowPOTSHIPX.Item("PO_SHIPMENT_LNO") & "")
            ASCMAIN1.sql = "Update POTSHIP2 Set WH_REC_NO = :PARM1 where PO_SHIPMENT_NO = :PARM2 and PO_SHIPMENT_LNO = :PARM3"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVN", New Object() {WH_REC_NO, PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
        Next



        Update_Record_TDA("WHTWREC7", "WH_REC_NO = '" & WH_REC_NO & "'")
        Update_Record_TDA("WHTWREC8", "WH_REC_NO = '" & WH_REC_NO & "'")
        Update_Record_TDA("WHTBARC1", "TRAN_TYPE = 'W' and TRAN_NO = '" & WH_REC_NO & "'")

        If chkComplete.Checked Then
            For Each rowWHTBARC0 As DataRow In dst.Tables("WHTBARC0").Select("")
                If rowWHTBARC0.Item("LOAD_STATUS") = "P" Then
                    rowWHTBARC0.Item("LOAD_STATUS") = "R"
                End If
            Next
        End If
        Update_Record_TDA("WHTBARC0", "LOAD_NO IN (SELECT DISTINCT LOAD_NO FROM WHTBARC1 WHERE TRAN_TYPE = 'W' and TRAN_NO = '" & WH_REC_NO & "')")


        Update_Record_TDA("WHTWREC1")
        CommitTrans("Update Complete")

        If chkComplete.Checked Then
            ASCMAIN1.sql = "Update WHTWREC1 Set WH_REC_STATUS = 'C' Where WH_REC_NO = '" & WH_REC_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            'rowWHTWREC1.Item("WH_REC_STATUS") = "C"

            ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                               New Object() {"W", WH_REC_NO, ASCMAIN1.SESSION_NO},
                               New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
            Record_Event("CMPLT", "Receipt Completed", True)
        Else
            Record_Event("RCVD", "Receipt Updated", True)

            Make_Locked_Loads_Available_for_Putaway()
        End If

        ASCMAIN1.sql = " Select C1.LOAD_NO, B1.WHSE_CODE, B1.LOCATION_CODE, B1.BAR_CODE, B1.STYLE_CODE, B1.COLOR_CODE, B1.LOCATION_QTY" _
        & " from WHTLOCB1 B1,WHTBARC1 C1,WHTLOCM1 M1" _
        & " where C1.BAR_CODE (+) = B1.BAR_CODE" _
        & " and M1.LOCATION_CODE (+) = B1.LOCATION_CODE and B1.LOCATION_QTY <> 0 " _
        & " and LOAD_NO is null And M1.LOCATION_CODE not in ('00-REC-A')"
        Dim TBL As DataTable = ASCDATA1.GetDataTable
        If TBL.Rows.Count <> 0 Then
            Using F As New ASFMSGBF
                F.Show_grd(TBL, Me, "Missing Load No's")
            End Using
        End If

        'ASCMAIN1.sql = " SELECT LOAD_NO, COUNT (*) LOCS, MIN (LOCATION_CODE) LOC1, MAX (LOCATION_CODE) LOC2 FROM (" _
        '& " SELECT WHTLOCB1.LOCATION_CODE, WHTBARC1.LOAD_NO, SUM (WHTLOCB1.LOCATION_QTY) UNITS, COUNT (*) CASES" _
        '& " FROM WHTLOCB1, WHTBARC1 WHERE WHTBARC1.BAR_CODE = WHTLOCB1.BAR_CODE" _
        '& " AND WHTLOCB1.LOCATION_CODE <> '00-REC-A'" _
        '& " AND WHTLOCB1.LOCATION_CODE <> '00-REC-B'" _
        '& " AND WHTLOCB1.LOCATION_CODE NOT LIKE '99%'" _
        '& " AND WHTLOCB1.LOCATION_QTY <> 0" _
        '& " AND WHTLOCB1.BAR_CODE <> '00000000' AND WHTBARC1.LOAD_NO <> '0000000000'" _
        '& " GROUP BY WHTLOCB1.LOCATION_CODE, WHTBARC1.LOAD_NO)" _
        '& " GROUP BY LOAD_NO" _
        '& " HAVING COUNT (*) > 1"
        'Dim TBL_Split_Loads As DataTable = ASCDATA1.GetDataTable
        'If TBL_Split_Loads.Rows.Count <> 0 Then
        '    Using F As New ASFMSGBF
        '        F.Show_grd(TBL_Split_Loads, Me, "Loads that Exist in Multiple Locations")
        '    End Using
        'End If


        If chkEmail.Checked Then
            Dim Send_To_Email As String = ""
            LookUp("POTSVIA1", Absx1.txtFor("PO_SHIP_VIA").Text)
            If cdr IsNot Nothing Then
                Send_To_Email = cdr.Item("PO_SHIP_VIA_EMAIL_CONTAINER") & ""
            End If

            Using frmTAFSEND1 As New TAFSEND1(Me)
                With frmTAFSEND1
                    Dim rowASTUSER1_EMAIL_FROM As DataRow = LookUp("ASTUSER1", ASCMAIN1.USER_ID, True)
                    Dim USER_SIGNATURE As String = ""

                    Dim Email_From As String = IIf(rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & "" = "",
                                 "donotreply" & "@" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_DEFAULT_EMAIL_DOMAIN"),
                                 rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & "")

                    .SEND_TO = Send_To_Email
                    .SEND_FROM = Email_From
                    .SEND_FROM_NAME = rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & ""
                    .SEND_CC = ASCMAIN1.USER_EMAIL
                    .SEND_CC_NAME = ASCMAIN1.USER_NAME
                    .SEND_FROM_SIGNATURE = USER_SIGNATURE
                    .SEND_FROM_NAME = ASCMAIN1.USER_NAME

                    .SEND_SUBJECT = "Container Pickup"
                    .SEND_BODY = ""
                    .SEND_ATTACHMENT = ""
                    .SEND_METHOD = "E"
                    .SEND_ENTITY_CAPTION = ""
                    .SEND_ENTITY_TABLE = ""
                    .SEND_ENTITY_KEY = ""
                    .SEND_ENTITY_NAME = ""
                    .ShowDialog()

                    If .SEND_STATUS = "S" Then
                        MsgBox("Email Sent", MsgBoxStyle.OkOnly, "Sent")
                        Record_Event("EML", "Email Sent", True)
                    Else
                        MsgBox("Error Occured: Could Not Send Email", MsgBoxStyle.OkOnly, "Error")
                        Record_Event("EMLX", "Email Failed", True)
                    End If
                End With
            End Using
        End If
    End Sub

    Sub Make_Locked_Loads_Available_for_Putaway()
        Dim SQL As String = "LOAD_STATUS = 'P' and LOAD_LOCKED = '1'"

        For Each rowWHTBARC0 As DataRow In dst.Tables("WHTBARC0").Select(SQL)
            Dim LOAD_NO As String = rowWHTBARC0.Item("LOAD_NO")
            ASCMAIN1.sql = "Update WHTBARC0 Set LOAD_STATUS = 'L' where LOAD_NO = '" & LOAD_NO & "'"
            ASCDATA1.ExecuteSQL()

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            Dim LOCATION_TO As String = rowICTWHSE1.Item("WHSE_LOC_PAW")
            If chkRemoteLoc.Checked = True And WHSE_CODE = "NJE" Then
                LOCATION_TO = "00-REC-C"
            End If

            Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
            Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
            With rowWHTMOVE1
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("WHSE_TRAN_TYPE") = "M"
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("STATUS") = "U"
            End With
            dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
            Update_Record_TDA("WHTMOVE1")

            Dim WHSE_TRAN_LNO_ctr As Integer = 0

            For Each rowWHTBARC1 As DataRow In dst.Tables("WHTBARC1").Select("LOAD_NO = '" & LOAD_NO & "'", "BAR_CODE")
                Dim BAR_CODE As String = rowWHTBARC1.Item("BAR_CODE")
                Dim TRAN_NO As String = rowWHTBARC1.Item("TRAN_NO")
                Dim PO_SHIPMENT_NO As String = rowWHTBARC1.Item("PO_SHIPMENT_NO")
                Dim PO_SHIPMENT_LNO As Integer = rowWHTBARC1.Item("PO_SHIPMENT_LNO")
                Dim CARTON_NO As Integer = rowWHTBARC1.Item("CARTON_NO")
                Dim sql8 As String = "WH_REC_NO = '" & TRAN_NO & "' and PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO)
                For Each rowWHTWREC8 As DataRow In dst.Tables("WHTWREC8").Select(sql8, "STYLE_CODE,COLOR_CODE")
                    Dim STYLE_CODE As String = rowWHTWREC8.Item("STYLE_CODE")
                    Dim COLOR_CODE As String = rowWHTWREC8.Item("COLOR_CODE")
                    Dim QTY As Integer = Val(rowWHTWREC8.Item("QTY") & "")

                    Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
                    With rowWHTMOVE2
                        .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                        WHSE_TRAN_LNO_ctr += 1
                        .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO_ctr
                        .Item("LOCATION_CODE_FROM") = rowICTWHSE1.Item("WHSE_LOC_REC")
                        .Item("LOCATION_CODE_TO") = LOCATION_TO
                        .Item("BAR_CODE") = BAR_CODE
                        .Item("WHSE_TRAN_QTY") = QTY
                        .Item("STYLE_CODE") = STYLE_CODE
                        .Item("COLOR_CODE") = COLOR_CODE
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = DATETIME_STAMP
                        .Item("STATUS") = "U"
                        .Item("LOAD_NO_FROM") = LOAD_NO
                    End With
                    dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
                Next

            Next
            Update_Record_TDA("WHTMOVE2")

            ASCDATA1.ExecuteSP("WHPMOVE1", "VNN",
                           New Object() {WHSE_TRAN_NO, 0, 1},
                           New String() {"WHSE_TRAN_NO_in", "WHSE_TRAN_LNO_in", "S"})

        Next
    End Sub

    Sub Reverse()
        BeginTrans()

        For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("")
            PO_SHIPMENT_LNO = Val(rowPOTSHIPX.Item("PO_SHIPMENT_LNO") & "")
            ASCMAIN1.sql = "Update POTSHIP2 Set WH_REC_NO = NULL where PO_SHIPMENT_NO = :PARM1 and PO_SHIPMENT_LNO = :PARM2"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VN", New Object() {PO_SHIPMENT_NO, PO_SHIPMENT_LNO})
        Next

        For Each TABLE_NAME As String In New String() _
        {"WHTBARC1", "WHTBARC0", "WHTWREC1", "WHTWREC7", "WHTWREC8"}
            ASCMAIN1.sql = "Delete from " & TABLE_NAME & " Where " & IIf(TABLE_NAME = "WHTBARC1" Or TABLE_NAME = "WHTBARC0", "TRAN_NO", "WH_REC_NO") & " = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", WH_REC_NO)
        Next

        CommitTrans("Reverse Complete")
        Record_Event("RVRS", "Receipt Reversed", True)
    End Sub


    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Report")

        Print_Report_Begin()
        Generate_Report("WHRWRECR", "Receipts Report ", , , , , False)
        Print_Report_End(False)



        If dst.Tables("WHTWREC7").Select("BAR_CODES - CARTONS <> 0").Length <> 0 _
            Or dst.Tables("WHTWREC7").Select("UNITS_REC - TOTAL_UNITS <> 0").Length <> 0 Then
            Print_Report_Begin()
            Generate_Report("WHRWRECR", "Receipts Report ", , "{WHTWREC7.BAR_CODES} - {WHTWREC7.CARTONS} <> 0", , , False)
            Print_Report_End(False)

        End If


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("Now Printing Report")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where = "WHSE_LOCATOR = '1' and WHSE_CTN_CTL = 'C'"
        End Select

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTWREC7, "B", "Clone Carton Line")
        Load_Popup_Menu(grdPOTSHIPX, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "PO Shipment Inquiry", "Select Entire Container")
        Load_Popup_Menu(grdTATEVNT1, "B", "Show email")
        Load_Popup_Menu(grdPOTSHIPC, "BB", "Expand Ranges", "Collapse Ranges")
        Load_Popup_Menu(grdWHTBARC0, "B", "Lock Load")

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
            Case "grdPOTSHIPX"
                If ScreenMode Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdPOTORDRS"
                '       tlb_btn = DirectCast(tlb_pop.Tools("Update Style Master (+ POs) with Changes made to Case Packs"), UltraWinToolbars.ButtonTool)
                '       tlb_btn.SharedProps.Visible = grdPOTORDRS.DisplayLayout.Bands(0).Columns("STYLE_ACTION").Hidden And Not (MENU_ITEM_OBJECT = "POFORDRI")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdWHTBARC0"
                    tlb_btn = DirectCast(tlb_pop.Tools("Lock Load"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("LOAD_STATUS").Value = "P"))

                Case "grdTATEVNT1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Show email"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("EVENT_TYPE").Value = "EML"))
                Case "grdWHTWREC7"
                    tlb_pop.Tools("Clone Carton Line").SharedProps.Caption = "Clone Carton Line " & grd.ActiveRow.Cells("CARTON_NO").Value
                    'tlb_pop.Tools("Clone Carton Line").SharedProps.Visible = IIf(grd.ActiveRow.Cells("CARTON_NO_CLONED_FROM").Value & "" = "", True, False)
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        'Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Clone Carton Line"
                If grdWHTWREC7.ActiveRow.Cells("CARTON_NO_CLONED_FROM").Value & "" & "" <> "" Then
                    MsgBox("Cannot Clone from a Cloned Line", MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Exit Sub
                End If
                For Each TABLE_NAME As String In New String() {"WHTWREC7", "WHTWREC8"}
                    For Each row As DataRow In dst.Tables(TABLE_NAME).Select("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'" _
                                                                                      & " And PO_SHIPMENT_LNO = " & PO_SHIPMENT_LNO _
                                                                                      & " And CARTON_NO = '" & CARTON_NO & "'")
                        Select Case TABLE_NAME
                            Case "WHTWREC7"
                                Add_WHTWREC7(row, True, row.Item("CARTON_NO"))
                            Case "WHTWREC8"
                                Add_WHTWREC8(row, True)
                        End Select
                    Next
                Next

            Case "Expand Ranges"

                PO_SHIPMENT_NO = grdWHTWREC7.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                PO_SHIPMENT_LNO = Val(grdWHTWREC7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                CARTON_NO = Val(grdWHTWREC7.ActiveRow.Cells("CARTON_NO").Value & "")

                Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO)
                For Each row As DataRow In dst.Tables("POTSHIPC").Select(sqlw)
                    Dim QTY As Integer = Val(row.Item("QTY") & "")
                    If QTY <> 1 Then
                        Dim BAR_CODE_start As Integer = Val(row.Item("BAR_CODE") & "")
                        Dim LOAD_NO As String = row.Item("LOAD_NO") & ""
                        Dim BAR_CODE_WITH As String = row.Item("BAR_CODE_WITH") & ""
                        For BC As Integer = 1 To QTY
                            If BC = 1 Then
                                row.Item("BAR_CODE2") = row.Item("BAR_CODE")
                                row.Item("QTY") = 1
                            Else
                                Dim BAR_CODE As String = Format(BAR_CODE_start + BC - 1, "".PadLeft(8, "0"))
                                Write_POTSHIPC(BAR_CODE, BAR_CODE, 1, LOAD_NO, BAR_CODE_WITH)
                            End If
                        Next
                    End If
                Next

                Sort_grdColumns(grdPOTSHIPC, "BAR_CODE")



            Case "Collapse Ranges"

                PO_SHIPMENT_NO = grdWHTWREC7.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                PO_SHIPMENT_LNO = Val(grdWHTWREC7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
                CARTON_NO = Val(grdWHTWREC7.ActiveRow.Cells("CARTON_NO").Value & "")

                Dim sqlw As String = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO)
                ASCDATA1.DeleteRows("POTSHIPC", sqlw)
                Create_POTSHIPC_VIEW(sqlw)

                'Case "Show All Cartons"
                '    'If CheckEditor Then
                '    For Each rowPOTSHIPC As DataRow In dst.Tables("POTSHIPC").Select("PO_SHIPMENT_LNO = " & grdWHTWREC7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value _
                '                                                                    & " And CARTON_NO = " & grdWHTWREC7.ActiveRow.Cells("CARTON_NO").Value)

                '        Dim Starting_Barcode As Integer = Val(rowPOTSHIPC.Item("BAR_CODE"))
                '        Dim Ending_Barcode As Integer = Val(rowPOTSHIPC.Item("BAR_CODE2"))
                '        Dim Load_No As String = rowPOTSHIPC.Item("LOAD_NO")
                '        CARTON_NO = grdWHTWREC7.ActiveRow.Cells("CARTON_NO").Value

                '        Dim SQL_Delete As String = "PO_SHIPMENT_NO = '" & rowPOTSHIPC.Item("PO_SHIPMENT_NO") & "'" _
                '                                                                             & " And PO_SHIPMENT_LNO = " & rowPOTSHIPC.Item("PO_SHIPMENT_LNO") _
                '                                                                             & " And CARTON_NO = " & rowPOTSHIPC.Item("CARTON_NO") _
                '                                                                             & " And BAR_CODE = '" & Format(Starting_Barcode, "".PadLeft(8, "0")) & "'"

                '        ASCDATA1.DeleteRows("POTSHIPC", SQL_Delete)
                '        For i As Integer = Starting_Barcode To Ending_Barcode
                '            Dim Formatted_BC As String = Format(i, "".PadLeft(8, "0"))

                '            ASCDATA1.DeleteRows("WHTBARC1", "BAR_CODE = '" & Formatted_BC & "'")

                '            Write_LPNs(Formatted_BC, Formatted_BC, 1, Load_No, "")
                '        Next

                '    Next
                '    'Else
                '    'ASCDATA1.DeleteRows("POTSHIPC", "")
                '    'Create_POTSHIPC_VIEW("")
                '    'End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "PO Shipment Inquiry"
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                Context_Launch("View", PO_SHIPMENT_NO, e.Tool.Key, "POFSHIPI", "F", "POE")

            Case "Select Entire Shipment"
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    Exit Sub
                End If
                grdPOTSHIPX.Selected.Rows.Clear()
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                For Each grow As UltraWinGrid.UltraGridRow In grd.ActiveRow.ParentRow.ChildBands(0).Rows
                    If Not grow.IsFilteredOut Then
                        If grow.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO Then
                            grow.Selected = True
                        End If
                    End If
                Next

            Case "Select Entire Container"
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    Exit Sub
                End If
                grdPOTSHIPX.Selected.Rows.Clear()
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value
                Dim CONTAINER_NO As String = grd.ActiveRow.Cells("CONTAINER_NO").Value
                For Each grow As UltraWinGrid.UltraGridRow In grd.ActiveRow.ParentRow.ChildBands(0).Rows
                    If Not grow.IsFilteredOut Then
                        If grow.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO And grow.Cells("CONTAINER_NO").Value = CONTAINER_NO Then
                            grow.Selected = True
                        End If
                    End If
                Next
            Case "Show email"
                If grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "PO-XMIT" _
                    Or grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "PO-XPED" Then
                    Dim FILENAME As String = grd.ActiveRow.Cells("EVENT_KEY").Value & ".EML"
                    Show_Document(ASCMAIN1.Folders("Archive") & "\email\Sent\" & FILENAME)
                End If

            Case "Lock Load"
                grd.ActiveRow.Cells("LOAD_LOCKED").Value = "1"
                grd.ActiveRow.Update()

        End Select
    End Sub


#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WH_REC_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If InquiryMode Then
                        Click_Command("View", e)
                    Else
                        Click_Command("Edit", e)
                    End If
                End If
            Case "STYLE_SEARCH"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Style_Search()
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                Load_POTSHIPX()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WH_REC_NO"
                If InquiryMode Then
                    Click_Command("View")
                Else
                    Click_Command("Edit")
                End If
        End Select
    End Sub

#End Region

    Private Sub grdPOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTSHIPX.AfterRowActivate
        Setup_POTSHIPX_Details()

        grdPOTSHIPX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdPOTSHIPX.DisplayLayout.Bands(0).Columns("WHSE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        If Not ScreenMode Then
            If grdPOTSHIPX.ActiveRow.IsGroupByRow = False Then
                If Val(grdPOTSHIPX.ActiveRow.Cells("LP_STATUS").Value & "") = 1 Then
                    grdPOTSHIPX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdPOTSHIPX.DisplayLayout.Bands(0).Columns("WHSE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    With grdPOTSHIPX.DisplayLayout.Bands(0)
                        For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                            If New String() {"WHSE_CODE"}.Contains(gcol.Key) Then
                                gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                            Else
                                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                            End If
                        Next
                    End With
                End If
            End If

        End If


    End Sub

    Private Sub grdPOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTSHIPX.DoubleClickRow

        If Not ScreenMode Then
            If grdPOTSHIPX.ActiveRow IsNot Nothing AndAlso grdPOTSHIPX.ActiveRow.IsDataRow Then
                If grdPOTSHIPX.ActiveRow.Cells("WHSE_CODE").Value <> User_Whse_Code Then
                    MsgBox("This Receipt is marked for " & grdPOTSHIPX.ActiveRow.Cells("WHSE_CODE").Value & ", You are Setup to Receive in Whse " & User_Whse_Code, vbOKOnly, "Cannot Receive")
                    Exit Sub
                End If

                grdPOTSHIPX.Selected.Rows.Clear()
                double_clicked = True
                PO_SHIPMENT_NO = grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Text
                PO_SHIPMENT_LNO = grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_LNO").Text
                CONTAINER_NO = grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Text

                If InquiryMode Then
                    Click_Command("View")
                Else
                    Click_Command("Receive")
                End If
            End If
        End If

    End Sub

    Sub Setup_POTSHIPX_Details()
        If grdPOTSHIPX.ActiveRow Is Nothing OrElse Not grdPOTSHIPX.ActiveRow.IsDataRow Then
            splPOTSHIPX.Panel2Collapsed = True
        Else
            Dim PO_SHIPMENT_NO As String = grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Value
            Dim PO_SHIPMENT_LNO As String = Val(grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
            Dim CONTAINER_NO As String = grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Value & ""

            Dim DVW As DataView
            DVW = DirectCast(grdPOTSHIP3.DataSource, DataTable).DefaultView
            DVW.RowFilter = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)
            Sort_grdColumns(grdPOTSHIP3, "PO_ORDER_LNO")

            DVW = DirectCast(grdWHTWREC7.DataSource, DataTable).DefaultView
            DVW.RowFilter = "PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO)

            grdPOTSHIP3.Text = "Shipment " & PO_SHIPMENT_NO & " Line " & CStr(PO_SHIPMENT_LNO) & "; Container '" & CONTAINER_NO & "' Contents"
            splPOTSHIPX.Panel2Collapsed = False

            grdWHTWREC7.Text = "Shipment " & PO_SHIPMENT_NO & " Line " & CStr(PO_SHIPMENT_LNO) & "; Container '" & CONTAINER_NO & "' Carton Types"


            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                For Each rowPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
                    If Val(rowPOTSHIP3.Item("CARTON_PACK_QTY")) <> 0 Then
                        Dim CARTON_NO As Int32 = rowPOTSHIP3.Item("PO_QTY_SHP") / rowPOTSHIP3.Item("CARTON_PACK_QTY")
                        If CARTON_NO > 0 Then
                            For i As Integer = 1 To CARTON_NO
                                rowPOTSHIP3.Item("CARTON_NOS") = rowPOTSHIP3.Item("CARTON_NOS") & CStr(i) + " | "
                            Next
                        End If
                    End If
                Next
            End If

        End If
    End Sub

    Sub Load_POTSHIPX()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Fill_Records("POTSHIPX")

        For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("")
            Dim PO_SHIPMENT_NO As String = rowPOTSHIPX.Item("PO_SHIPMENT_NO")
            Dim CONTAINER_NO As String = rowPOTSHIPX.Item("CONTAINER_NO") & ""
            ASCMAIN1.sql = "Select * from POTSHIP4 where PO_SHIPMENT_NO = :PARM1 and CONTAINER_NO = :PARM2"
            Dim rowPOTSHIP4 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New String() {PO_SHIPMENT_NO, CONTAINER_NO})
            If rowPOTSHIP4 IsNot Nothing Then
                rowPOTSHIPX.Item("CONTAINER_TYPE_CODE") = rowPOTSHIP4.Item("CONTAINER_TYPE_CODE")
            End If
        Next

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        'If WHSE_CODE <> "" Then
        '    Dim dvw As DataView = DirectCast(grdPOTSHIPX.DataSource, DataTable).DefaultView
        '    dvw.RowFilter = "WHSE_CODE = '" & WHSE_CODE & "'"
        'End If

        With grdPOTSHIPX.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("WHSE_CODE", False, True)
            .SortedColumns.Add("PO_SHIPMENT_NO", True)
        End With
        grdPOTSHIPX.Rows.ExpandAll(True)

        Setup_POTSHIPX_Details()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Receipts History")
        Dim OPS_YYYYPP As String = cbeReceipts.Value
        Dim OPS_YYYYPP2 As String = cbeReceipts2.Value
        Fill_Records("WHTWRECX", New String() {OPS_YYYYPP, OPS_YYYYPP2})
        Sort_grdColumns(grdWHTWRECX, "WH_REC_NO".ToLower)

        ' Setup_ICTIRECX()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdPOTSHIPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTSHIPX.InitializeRow
        If e.Row.Cells("LP_STATUS").Value & "" <> "1" Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
            e.Row.ToolTipText = "Shipment has not been Released for Warehouse Receipt"
        Else
            e.Row.Appearance.BackColor = Drawing.Color.Empty
            e.Row.ToolTipText = ""
        End If
    End Sub

    Private Sub grdWHTWREC7_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTWREC7.AfterRowActivate
        PO_SHIPMENT_LNO = Val(grdWHTWREC7.ActiveRow.Cells("PO_SHIPMENT_LNO").Value & "")
        CARTON_NO = Val(grdWHTWREC7.ActiveRow.Cells("CARTON_NO").Value & "")
        CONTAINER_NO = grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Value & ""


        If grdWHTWREC7.ActiveRow.Cells("CARTON_NO_CLONED_FROM").Value & "" & "" <> "" Then
            grdWHTWREC8.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdWHTWREC8.DisplayLayout.Bands(0).Columns("QTY").CellActivation = UltraWinGrid.Activation.AllowEdit

        Else
            grdWHTWREC8.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdWHTWREC8.DisplayLayout.Bands(0).Columns("QTY").CellActivation = UltraWinGrid.Activation.NoEdit
        End If

        Setup_grdWHTWREC8()
        Setup_grdPOTSHIPC()
    End Sub


    Sub grd_Update()
        grdWHTWREC7.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        With grdWHTWREC7.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"CARTON_COMMENTS"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        If tabMain.SelectedTab.Key = "Receipts" Or MENU_ITEM_OBJECT = "WHFWRECI" Then
            grdPOTSHIPX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            With grdPOTSHIPX.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                Next
            End With


            grdWHTWREC8.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            With grdWHTWREC8.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                Next
            End With
        Else
            grdWHTWREC8.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            With grdWHTWREC8.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If New String() {"QTY"}.Contains(gcol.Key) Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                        gcol.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                        gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                Next
            End With
        End If
    End Sub

    Sub Setup_grdPOTSHIPC()
        If grdWHTWREC7.ActiveRow Is Nothing OrElse Not grdWHTWREC7.ActiveRow.IsDataRow Then
            grdPOTSHIPC.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdPOTSHIPC.DataSource, DataTable).DefaultView
            dvw.RowFilter = ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO))
            '  grdPOTSHIPC.Text = "Carton Type " & CStr(CARTON_NO) & " by LPN"
            grdPOTSHIPC.Visible = True
        End If
    End Sub

    Sub Setup_grdWHTWREC8()
        If grdWHTWREC7.ActiveRow Is Nothing OrElse Not grdWHTWREC7.ActiveRow.IsDataRow Then
            grdWHTWREC8.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdWHTWREC8.DataSource, DataTable).DefaultView
            dvw.RowFilter = ("PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "' and PO_SHIPMENT_LNO = " & CStr(PO_SHIPMENT_LNO) & " and CARTON_NO = " & CStr(CARTON_NO))
            grdWHTWREC8.Text = "Carton Type " & CStr(CARTON_NO) & " by Style/Color"
            grdWHTWREC8.Visible = True
        End If
    End Sub

    Private Sub txtBAR_CODE_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtBAR_CODE.KeyDown
        If e.KeyValue = Keys.Enter Then
            e.Handled = True
            Validate_BAR_CODE()
        End If
    End Sub

    Private Sub txtBAR_CODE_Leave(sender As Object, e As System.EventArgs) Handles txtBAR_CODE.Leave
        Validate_BAR_CODE()
    End Sub

    Private Sub txtBAR_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtBAR_CODE.ValueChanged
        If txtBAR_CODE.Text.Length = 8 And txtBAR_CODE.Tag & "" = "" And tagCR Then
            Validate_BAR_CODE()
        End If
    End Sub

    Private Sub txtBAR_CODE2_GotFocus(sender As Object, e As System.EventArgs) Handles txtBAR_CODE2.GotFocus
        If txtBAR_CODE.Text = "" Then
            txtBAR_CODE.Focus()
        End If
    End Sub

    Private Sub txtBAR_CODE2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtBAR_CODE2.KeyDown
        If e.KeyValue = Keys.Enter Then
            If (txtBAR_CODE2.Text.StartsWith("+") Or txtBAR_CODE2.Text.Length <= 4) And txtBAR_CODE.Text <> "" Then
                Dim C As Int64 = Val(txtBAR_CODE2.Text)
                txtBAR_CODE2.Text = Format(Val(txtBAR_CODE.Text) + C - 1, "00000000")
            End If
        End If
        If e.KeyValue = Keys.Enter Then
            e.Handled = True
            Validate_BAR_CODE2()
        End If
    End Sub

    Private Sub txtBAR_CODE2_Leave(sender As Object, e As System.EventArgs) Handles txtBAR_CODE2.Leave
        Validate_BAR_CODE2()
    End Sub

    Private Sub txtBAR_CODE2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtBAR_CODE2.ValueChanged
        If txtBAR_CODE2.Text.Length = 8 And txtBAR_CODE.Tag & "" = "" And tagCR Then
            Validate_BAR_CODE2()
        End If
    End Sub

    Function Check_BAR_CODE(BAR_CODE As String) As String

        If BAR_CODE = "" Then Return BAR_CODE

        If BAR_CODE.PadLeft(8, "0") <> Format(Val(BAR_CODE), "".PadLeft(8, "0")) Then
            BAR_CODE = ""
        Else
            BAR_CODE = BAR_CODE.PadLeft(8, "0")
        End If
        Return BAR_CODE
    End Function

    Sub Validate_BAR_CODE()
        If Not ScreenMode Then Exit Sub
        Dim BAR_CODE As String = Check_BAR_CODE(txtBAR_CODE.Text)

        If BAR_CODE <> "" Then
            LookUp("WHTBARC1", BAR_CODE)
            If cdr IsNot Nothing Then
                MsgBox("LPN already exists in Database" & vbCrLf _
                    & "Received on PO " & cdr.Item("PO_ORDER_NO") _
                    & " on " & cdr.Item("PO_DATE_RECEIVED"),
                    MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                BAR_CODE = ""
            Else
                Dim rowWHTBARC1 As DataRow = dst.Tables("WHTBARC1").Rows.Find(BAR_CODE)
                If rowWHTBARC1 IsNot Nothing Then
                    Dim rowPOTSHIPX As DataRow = dst.Tables("POTSHIPX").Rows.Find(New Object() {rowWHTBARC1.Item("PO_SHIPMENT_NO"), rowWHTBARC1.Item("PO_SHIPMENT_LNO")})
                    MsgBox("LPN already exists in Current Receipt" & vbCrLf & vbCrLf _
                        & "Received on Container " & rowPOTSHIPX.Item("CONTAINER_NO") _
                        & " Shipment Line " & rowWHTBARC1.Item("PO_SHIPMENT_LNO") _
                        & " as part of Carton Type " & rowWHTBARC1.Item("CARTON_NO"),
                        MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                    BAR_CODE = ""
                End If
            End If

            txtBAR_CODE.Text = BAR_CODE
            If BAR_CODE = "" Then
                txtBAR_CODE.Focus()
            Else
                txtBAR_CODE2.Focus()
            End If
        End If
    End Sub

    Sub Validate_BAR_CODE2()
        If Not ScreenMode Then Exit Sub

        Dim BAR_CODE As String = txtBAR_CODE.Text
        Dim BAR_CODE2 As String = Check_BAR_CODE(txtBAR_CODE2.Text)

        If BAR_CODE2 <> "" Then
            If BAR_CODE.Length <> BAR_CODE2.Length Then
                MsgBox("Invalid Value for LPN",
                    MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN2")
                BAR_CODE2 = ""
            Else
                If BAR_CODE2 < BAR_CODE Then
                    MsgBox("Invalid Range for LPNs" & vbCrLf _
                        & BAR_CODE & " thru " & BAR_CODE2,
                        MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN2")
                    BAR_CODE2 = ""
                End If
            End If
        End If

        If BAR_CODE2 <> "" Then

            ASCMAIN1.sql = "Select Count(*) BAR_CODES from WHTBARC1 where BAR_CODE >= :PARM1 and BAR_CODE <= :PARM2"
            Dim BAR_CODES As Int64 = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {BAR_CODE, BAR_CODE2}))
            If BAR_CODES <> 0 Then
                MsgBox(CStr(BAR_CODES) & " LPN(s) already exist in Database" & vbCrLf _
                    & " in Range from " & BAR_CODE & " thru " & BAR_CODE2,
                    MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                BAR_CODE2 = ""
            Else
                Dim rowWHTBARC1s() As DataRow = dst.Tables("WHTBARC1").Select("BAR_CODE >= '" & BAR_CODE & "' and BAR_CODE <= '" & BAR_CODE2 & "'")
                If rowWHTBARC1s.Length <> 0 Then
                    Dim rowPOTSHIPX As DataRow = dst.Tables("POTSHIPX").Rows.Find(New Object() {rowWHTBARC1s(0).Item("PO_SHIPMENT_NO"), rowWHTBARC1s(0).Item("PO_SHIPMENT_LNO")})
                    MsgBox(CStr(rowWHTBARC1s.Length) & " LPN(s) already exists in Current Receipt" & vbCrLf & vbCrLf _
                        & "Received in Container " & rowPOTSHIPX.Item("CONTAINER_NO") _
                        & " Shipment Line " & rowWHTBARC1s(0).Item("PO_SHIPMENT_LNO") _
                        & " as part of Carton Type " & rowWHTBARC1s(0).Item("CARTON_NO"),
                        MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                    BAR_CODE2 = ""
                End If
            End If

            txtBAR_CODE2.Text = BAR_CODE2
            If BAR_CODE2 = "" Then
                txtBAR_CODE2.Focus()
            Else
                txtBAR_CODE_WITH.Focus()
            End If
        End If
    End Sub

    Sub Validate_BAR_CODE_WITH()
        If Not ScreenMode Then Exit Sub
        If txtBAR_CODE.Text = "" Or txtBAR_CODE2.Text = "" Then Exit Sub

        Dim BAR_CODE_WITH As String = Check_BAR_CODE(txtBAR_CODE_WITH.Text)
        If BAR_CODE_WITH = "".PadLeft(8, "0") Then
            Dim LOAD_NO As String = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
            txtLOAD_NO.Text = LOAD_NO
            Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").NewRow
            rowWHTBARC0.Item("LOAD_NO") = LOAD_NO
            rowWHTBARC0.Item("WHSE_CODE") = WHSE_CODE
            rowWHTBARC0.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTBARC0.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTBARC0.Item("LOAD_STATUS") = "P"
            rowWHTBARC0.Item("LOCATION_CODE") = LOCATION_CODE
            rowWHTBARC0.Item("TRAN_TYPE") = "W"
            rowWHTBARC0.Item("TRAN_NO") = WH_REC_NO
            dst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
        Else
            If BAR_CODE_WITH = "" Then
                Exit Sub
            Else

                Dim rowWHTBARC1 As DataRow = dst.Tables("WHTBARC1").Rows.Find(BAR_CODE_WITH)
                If rowWHTBARC1 Is Nothing Then
                    MsgBox("Cannot Find LPN " & BAR_CODE_WITH & " in Current Receipt", MsgBoxStyle.OkOnly, "Invalid LPN Entry")
                    BAR_CODE_WITH = ""
                Else
                    If dst.Tables("WHTBARC0").Select("LOAD_NO= '" & rowWHTBARC1.Item("LOAD_NO") & "' and LOAD_STATUS <> 'P'").Length <> 0 Then
                        MsgBox("Load No" & BAR_CODE_WITH & " is NOT Pending, Cannot Add LPN", MsgBoxStyle.OkOnly, "Invalid LPN Entry")
                        BAR_CODE_WITH = ""
                    Else
                        txtLOAD_NO.Text = rowWHTBARC1.Item("LOAD_NO")
                    End If
                End If
            End If
        End If

        txtBAR_CODE_WITH.Text = BAR_CODE_WITH

        If txtBAR_CODE_WITH.Text = "" Then
            txtBAR_CODE_WITH.Focus()
        Else
            Write_LPNs()
        End If
    End Sub

    Private Sub tabContainerDetails_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabContainerDetails.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_LPN_Entry()
    End Sub

    Sub Setup_LPN_Entry()
        UltraExplorerBar1.Groups("LPN Entry").Visible = ScreenMode And (tabContainerDetails.SelectedTab.Key = "Carton Type Summary")
        UltraExplorerBar1.Groups("Style Lookup").Visible = ScreenMode And (tabContainerDetails.SelectedTab.Key = "Carton Type Summary")
    End Sub


    Private Sub grdWHTWREC7_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTWREC7.InitializeRow
        Dim BAR_CODES As Integer = Val(e.Row.Cells("BAR_CODES").Value & "")
        Dim CARTONS As Integer = Val(e.Row.Cells("CARTONS").Value & "")
        If BAR_CODES = CARTONS And Val(e.Row.Cells("CARTON_NO_CLONED_FROM").Value & "") = 0 Then
            e.Row.Cells("BAR_CODES").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("BAR_CODES").Appearance.BackColor = Drawing.Color.Empty
        ElseIf BAR_CODES = 0 Then
            e.Row.Cells("BAR_CODES").Appearance.BackColor = Drawing.Color.HotPink
        Else

            e.Row.Cells("BAR_CODES").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("BAR_CODES").Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub txtBAR_CODE_WITH_GotFocus(sender As Object, e As System.EventArgs) Handles txtBAR_CODE_WITH.GotFocus
        If txtBAR_CODE.Text = "" Then
            txtBAR_CODE.Focus()
        ElseIf txtBAR_CODE2.Text = "" Then
            txtBAR_CODE2.Focus()
        End If
    End Sub

    Private Sub txtBAR_CODE_WITH_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtBAR_CODE_WITH.KeyDown
        If e.KeyValue = Keys.Enter Then
            e.Handled = True
            Validate_BAR_CODE_WITH()
        End If
    End Sub

    Private Sub txtBAR_CODE_WITH_Leave(sender As Object, e As System.EventArgs) Handles txtBAR_CODE_WITH.Leave
        Validate_BAR_CODE_WITH()
    End Sub

    Private Sub txtBAR_CODE_WITH_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtBAR_CODE_WITH.ValueChanged
        If txtBAR_CODE_WITH.Text.Length = 8 And txtBAR_CODE.Tag & "" = "" And tagCR Then
            Validate_BAR_CODE_WITH()
        End If
    End Sub

    Sub Write_POTSHIPC(BAR_CODE As String, BAR_CODE2 As String, QTY As Int32, LOAD_NO As String, BAR_CODE_WITH As String)
        Dim rowPOTSHIPC As DataRow = dst.Tables("POTSHIPC").NewRow
        rowPOTSHIPC.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
        rowPOTSHIPC.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
        rowPOTSHIPC.Item("CARTON_NO") = CARTON_NO
        rowPOTSHIPC.Item("BAR_CODE") = BAR_CODE
        rowPOTSHIPC.Item("BAR_CODE2") = BAR_CODE2
        rowPOTSHIPC.Item("QTY") = QTY
        rowPOTSHIPC.Item("LOAD_NO") = LOAD_NO
        rowPOTSHIPC.Item("BAR_CODE_WITH") = BAR_CODE_WITH
        dst.Tables("POTSHIPC").Rows.Add(rowPOTSHIPC)
    End Sub

    Sub Write_LPNs()

        Dim BAR_CODE As String = txtBAR_CODE.Text
        Dim BAR_CODE2 As String = txtBAR_CODE2.Text
        Dim QTY = Val(BAR_CODE2) - Val(BAR_CODE) + 1

        Dim BAR_CODE_WITH As String = txtBAR_CODE_WITH.Text
        Dim LOAD_NO As String = txtLOAD_NO.Text
        Dim CONTAINER_NO_ACTUAL As String = CONTAINER_NO

        If txtNewContainer.Text <> "" Then
            CONTAINER_NO_ACTUAL = txtNewContainer.Text
        End If

        Write_POTSHIPC(BAR_CODE, BAR_CODE2, QTY, LOAD_NO, BAR_CODE_WITH)

        Dim BAR_CODE_first As Int64 = Val(BAR_CODE)

        For i As Integer = 1 To QTY
            Dim rowWHTBARC1 As DataRow = dst.Tables("WHTBARC1").NewRow
            rowWHTBARC1.Item("BAR_CODE") = Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
            rowWHTBARC1.Item("PO_ORDER_NO") = "?" ' PO_ORDER_NO
            rowWHTBARC1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowWHTBARC1.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            rowWHTBARC1.Item("CARTON_NO") = CARTON_NO
            rowWHTBARC1.Item("LOAD_NO") = LOAD_NO
            rowWHTBARC1.Item("CONTAINER_NO_ACTUAL") = CONTAINER_NO_ACTUAL
            'rowWHTBARC1.Item("LOCATION_CODE") = LOCATION_CODE

            Dim rowWHTWREC7 As DataRow = dst.Tables("WHTWREC7").Rows.Find(New Object() {WH_REC_NO, PO_SHIPMENT_NO, PO_SHIPMENT_LNO, CARTON_NO})
            If rowWHTWREC7 IsNot Nothing Then
                Dim PPK_CODE As String = rowWHTWREC7.Item("PPK_CODE") & ""
                rowWHTBARC1.Item("PPK_CODE") = PPK_CODE
            End If

            dst.Tables("WHTBARC1").Rows.Add(rowWHTBARC1)
            Sort_grdColumns(grdWHTBARC0, "LOAD_NO".ToLower)
        Next

        txtBAR_CODE.Text = ""
        txtBAR_CODE2.Text = ""
        txtBAR_CODE_WITH.Text = ""
        txtLOAD_NO.Text = ""
        txtBAR_CODE.Focus()
    End Sub

    Private Sub grdPOTSHIPC_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTSHIPC.AfterRowsDeleted
        Dim LOAD_NOs As New List(Of String)

        txtBAR_CODE.Tag = "X"

        For b As Integer = 1 To BAR_CODE_deleted.Count
            Dim BAR_CODE_first As Int64 = Val(BAR_CODE_deleted(b - 1))
            Dim QTY As Int32 = QTY_deleted(b - 1)
            For i As Integer = 1 To QTY
                Dim BAR_CODE As String = Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
                If b = 1 Then
                    If i = 1 Then txtBAR_CODE.Text = BAR_CODE
                    If i = QTY Then txtBAR_CODE2.Text = BAR_CODE
                End If
                Dim rowWHTBARC1 As DataRow = dst.Tables("WHTBARC1").Rows.Find(BAR_CODE)
                Dim LOAD_NO As String = rowWHTBARC1.Item("LOAD_NO")
                If Not LOAD_NOs.Contains(LOAD_NO) Then
                    LOAD_NOs.Add(LOAD_NO)
                End If
                rowWHTBARC1.Delete()
            Next
        Next
        txtBAR_CODE_WITH.Text = ""

        txtBAR_CODE.Tag = "'"

        For Each LOAD_NO As String In LOAD_NOs
            Dim rowWHTBARC0 As DataRow = dst.Tables("WHTBARC0").Rows.Find(LOAD_NO)
            If Val(rowWHTBARC0.Item("QTY") & "") = 0 Then
                rowWHTBARC0.Delete()
            End If
        Next
    End Sub

    Private Sub grdPOTSHIPC_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTSHIPC.BeforeRowsDeleted
        If grdPOTSHIPC.ActiveRow.Cells("LOAD_STATUS").Value <> "P" And grdPOTSHIPC.ActiveRow.Cells("LOAD_STATUS").Value <> "L" Then
            MsgBox("Can Only Delete LPN's attached to Loads that are Locked or Pending", MsgBoxStyle.OkOnly, "Cannot Delete")
            e.Cancel = True
        Else
            BAR_CODE_deleted.Clear()
            QTY_deleted.Clear()
            For Each grow As UltraWinGrid.UltraGridRow In e.Rows
                BAR_CODE_deleted.Add(grow.Cells("BAR_CODE").Value)
                QTY_deleted.Add(grow.Cells("QTY").Value)
            Next
        End If

    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        If tabMain.SelectedTab Is Nothing Then Exit Sub

        ' IIf(tabMain.SelectedTab.Key = "Receipts" And MENU_ITEM_OBJECT = "WHFWRECI", True, False)

        UltraExplorerBar1.Groups("Screen Control").Items("Reverse").Visible = IIf(tabMain.SelectedTab.Key = "Receipts" And MENU_ITEM_OBJECT = "WHFWRECI", False, True)
        UltraExplorerBar1.Groups("Screen Control").Items("View").Visible = IIf(tabMain.SelectedTab.Key = "Receipts" And MENU_ITEM_OBJECT = "WHFWRECI", True, False)
        UltraExplorerBar1.Groups("Screen Control").Items("Update").Visible = IIf(tabMain.SelectedTab.Key = "Shipments", True, False)
        UltraExplorerBar1.Groups("Screen Control").Items("Receive").Visible = IIf(tabMain.SelectedTab.Key = "Shipments", True, False)
        UltraExplorerBar1.Groups("Screen Control").Items("Delete").Visible = IIf(MENU_ITEM_OBJECT <> "WHFWRECI", True, False)
        UltraExplorerBar1.Groups("Receipts History").Visible = InquiryMode

        Select Case tabMain.SelectedTab.Key
            Case "Receipts"

            Case "Shipments"
                Setup_Receipts_History()
        End Select
        Click_Command("Refresh")
    End Sub

    Sub Setup_Receipts_History()
        UltraExplorerBar1.Groups("Receipts History").Visible = (tabMain.SelectedTab.Key = "Receipts") And Not ScreenMode
    End Sub

    Private Sub grdWHTWRECX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTWRECX.DoubleClickRow
        If grdWHTWRECX.ActiveRow IsNot Nothing AndAlso grdWHTWRECX.ActiveRow.IsDataRow Then
            WH_REC_NO = grdWHTWRECX.ActiveRow.Cells("WH_REC_NO").Value
            Absx1.txtFor("WH_REC_NO").Text = WH_REC_NO

            If InquiryMode Then
                Click_Command("View")
            Else
                Click_Command("Edit")
            End If
        End If
    End Sub

    Private Sub grdWHTWRECX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTWRECX.InitializeRow
        If e.Row.Cells("WH_REC_STATUS").Value & "" = "V" Then
            e.Row.Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        ElseIf e.Row.Cells("WH_REC_STATUS").Value & "" = "P" Then
            e.Row.Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub


    Private Sub chkFinalize_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkComplete.CheckedChanged
        If rowWHTWREC1.Item("WH_REC_EMAIL_SENT") & "" = "1" Then
        Else
            If chkComplete.Checked Then
                chkEmail.Checked = True
                chkEmail.Enabled = False
            Else
                chkEmail.Checked = False
                chkEmail.Enabled = True
            End If
        End If

    End Sub

    Private Sub chkTagCR_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkTagCR.CheckedChanged
        tagCR = chkTagCR.Checked
    End Sub

    Private Sub cmdStyle_Search_Click(sender As System.Object, e As System.EventArgs) Handles cmdStyle_Search.Click
        Style_Search()
    End Sub

    Sub Style_Search()
        If txtStyle_Search.Text = "" Then
            MsgBox("Please Enter a Style", MsgBoxStyle.OkOnly, "Invalid Style")
            Exit Sub
        Else
            txtStyle_Search.Text = txtStyle_Search.Text.ToUpper
        End If
        Dim Found_Style As Boolean = False
        For Each rowWHTWREC7 As DataRow In dst.Tables("WHTWREC7").Select("", "PO_SHIPMENT_NO, PO_SHIPMENT_LNO,CARTON_NO")
            If Found_Style = True Then Exit For

            If txtStyle_Search.Text = rowWHTWREC7.Item("STYLE_CODE") & "" Then
                For Each gridrow As UltraWinGrid.UltraGridRow In grdPOTSHIPX.Rows
                    If gridrow.Cells("PO_SHIPMENT_NO").Value = rowWHTWREC7.Item("PO_SHIPMENT_NO") _
                    And gridrow.Cells("PO_SHIPMENT_LNO").Value = rowWHTWREC7.Item("PO_SHIPMENT_LNO") Then
                        grdPOTSHIPX.ActiveRow = gridrow
                        Exit For
                    End If
                Next

                For Each gridrow As UltraWinGrid.UltraGridRow In grdWHTWREC7.Rows
                    If gridrow.Cells("PO_SHIPMENT_NO").Value = rowWHTWREC7.Item("PO_SHIPMENT_NO") _
                    And gridrow.Cells("PO_SHIPMENT_LNO").Value = rowWHTWREC7.Item("PO_SHIPMENT_LNO") _
                    And gridrow.Cells("STYLE_CODE").Value = rowWHTWREC7.Item("STYLE_CODE") Then
                        grdWHTWREC7.ActiveRow = gridrow
                        Found_Style = True
                        Exit For
                    End If
                Next
            End If

        Next
        If Found_Style = False Then
            MsgBox("Style Not Found", MsgBoxStyle.OkOnly, "Not Found")
        End If
        txtStyle_Search.Text = ""
    End Sub

    Private Sub grdWHTBARC0_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTBARC0.InitializeRow
        e.Row.Appearance.ForeColor = Drawing.Color.Empty
        e.Row.Appearance.BackColor = Drawing.Color.Empty

        If e.Row.Cells("LOAD_STATUS").Value = "L" Then
            e.Row.Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("LOAD_LOCKED").Value = "1" Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        End If
    End Sub

    Private Sub grdPOTSHIPC_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTSHIPC.InitializeLayout

    End Sub

    Function Check_Deleted_LPNs() As String
        For Each rowWHTBARCC As DataRow In dst.Tables("WHTBARCC").Rows
            If dst.Tables("WHTBARC1").Select("BAR_CODE = '" & rowWHTBARCC.Item("BAR_CODE") & "'").Length = 0 Then
                Deleted_BarCodes &= ",'" & rowWHTBARCC.Item("BAR_CODE") & "'"
            End If

        Next

        If Deleted_BarCodes <> "" Then
            ASCMAIN1.sql = " Select C1.LOAD_NO, B1.WHSE_CODE, B1.LOCATION_CODE, B1.BAR_CODE, B1.STYLE_CODE, B1.COLOR_CODE, B1.LOCATION_QTY " _
            & " from WHTBARC1 C1, WHTLOCB1 B1" _
            & " Where B1.LOCATION_QTY > 0" _
            & " And C1.BAR_CODE = B1.BAR_CODE" _
            & " And C1.BAR_CODE in (" & Mid(Deleted_BarCodes, 2) & ")"

            Dim TBL As DataTable = ASCDATA1.GetDataTable
            If TBL.Rows.Count <> 0 Then
                Using F As New ASFMSGBF
                    F.Show_grd(TBL, Me, "Is it ok to delete these LPN's that are part of a locked load?")
                    If F.user_option = 0 Then
                        Return ""
                    Else
                        Return "Update Aborted"
                    End If
                End Using
            Else
                Return ""
            End If
        Else
            Return ""
        End If


    End Function

    Private Sub grdWHTWREC8_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTWREC8.InitializeLayout

    End Sub

    Private Sub grdWHTWREC7_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdWHTWREC7.InitializeLayout

    End Sub

    Private Sub grdPOTSHIPX_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdPOTSHIPX.AfterCellUpdate
        'If grdPOTSHIPX.ActiveRow IsNot Nothing AndAlso grdPOTSHIPX.ActiveRow.IsDataRow Then
        '    Dim iresponse As Int16 = MsgBox("Changing this Whse Code Will Move All Containers Labeled: " & grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Value & " on This Shipment, Proceed?", MsgBoxStyle.YesNo, "Proceed")
        '    If iresponse = MsgBoxResult.No Then
        '        Load_POTSHIPX()

        '    Else
        '        For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("PO_SHIPMENT_NO = '" & grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Value & "'" _
        '                                                                         & " And CONTAINER_NO = '" & grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Value & "'")
        '            rowPOTSHIPX.Item("WHSE_CODE") = grdPOTSHIPX.ActiveRow.Cells("WHSE_CODE").Value
        '        Next
        '    End If
        'End If

    End Sub

    Private Sub grdPOTSHIPX_AfterCellListCloseUp(sender As Object, e As CellEventArgs) Handles grdPOTSHIPX.AfterCellListCloseUp

    End Sub

    Private Sub grdPOTSHIPX_CellChange(sender As Object, e As CellEventArgs) Handles grdPOTSHIPX.CellChange
        If grdPOTSHIPX.ActiveRow IsNot Nothing AndAlso grdPOTSHIPX.ActiveRow.IsDataRow Then
            grdPOTSHIPX.ActiveRow.Update()
            Dim iresponse As Int16 = MsgBox("Changing this Whse Code Will Move All Containers Labeled: " & grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Value & " on This Shipment, Proceed?", MsgBoxStyle.YesNo, "Proceed")
            If iresponse = MsgBoxResult.No Then
                Load_POTSHIPX()

            Else

                For Each rowPOTSHIPX As DataRow In dst.Tables("POTSHIPX").Select("PO_SHIPMENT_NO = '" & grdPOTSHIPX.ActiveRow.Cells("PO_SHIPMENT_NO").Value & "'" _
                                                                                 & " And CONTAINER_NO = '" & grdPOTSHIPX.ActiveRow.Cells("CONTAINER_NO").Value & "'")
                    rowPOTSHIPX.Item("WHSE_CODE") = grdPOTSHIPX.ActiveRow.Cells("WHSE_CODE").Value
                Next
            End If
        End If
    End Sub

    Private Sub grdPOTSHIPX_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTSHIPX.InitializeLayout

    End Sub

    Private Sub grdWHTWRECX_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdWHTWRECX.InitializeLayout

    End Sub

    Private Sub txtNewContainer_EditorButtonClick(sender As Object, e As UltraWinEditors.EditorButtonEventArgs) Handles txtNewContainer.EditorButtonClick
        ASCMAIN1.CodeSelector.Get_SQL("CONTAINER_NO")

        ASCMAIN1.CodeSelector.SQL = "Select DISTINCT PO_SHIPMENT_NO, CONTAINER_NO from POTSHIP2" _
                        & " Where PO_SHIPMENT_NO = '" & PO_SHIPMENT_NO & "'"
        ASCMAIN1.CodeSelector.MultipleSelections = False
        Using F As New ASFCODE1
            F.ShowDialog()
        End Using
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            txtNewContainer.Text = ASCMAIN1.CodeSelector.SelectedCode
        End If
    End Sub
End Class