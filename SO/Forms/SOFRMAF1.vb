Public Class SOFRMAF1

    ' July 13, 2017
    ' INSERT INTO SOTRMAR1 SELECT 'C', 'Customer Discount' from dual;
    ' Insert Into ARTREASR SELECT 'C', 'Cust Disc', 'RD' from dual;
    ' Alter TABLE sotrmaf1 add REASON_CODE VARCHAR2(6)
    ' alter table SOTRMAF1 ADD RA_EMAIL VARCHAR2(100);
    ' alter table SOTRMAF1 ADD RA_CONTACT VARCHAR2(60);
    ' alter table SOTRMAF1 ADD RA_PHONE VARCHAR2(20);
    ' alter table SOTRMAF1 ADD RA_CARRIER_CODE VARCHAR2(6);
    ' alter table SOTRMAF1 ADD RA_CARTONS NUMBER(3);
    ' ALTER TABLE SOTRMAF1 ADD CALL_TAG_USER VARCHAR2(10);
    ' ALTER TABLE SOTRMAF1 ADD CALL_TAG_DATE DATE;

#Region "Declarations"

    Private CUST_CODE As String
    Private CUST_NAME As String             ' Sold-To Customer Name
    Private CUST_BILL_TO_CUST As String
    Private CUST_CLAIM_NO As String
    Private ORDR_NO As String = String.Empty
    Private is180Customer As Boolean = False

    Private RA_NO As String
    Private rowSOTRMAF1 As DataRow

    Private rowARTCUST1 As DataRow          ' ARTCUST1 for the Sold-To
    Private rowARTCUST1_BT As DataRow       ' ARTCUST1 for the Bill-To
    Private rowICTSTYL1 As DataRow
    Private RA_LNOs As New List(Of Int64)   ' list of RA_LNOs that are deleted

    Private COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE
    Private tblSOTRMAF2 As New DataTable

    Private importedFromExcel As Boolean = False
    Private loadedFromExcel As Boolean = False

    Private Class ItemsImported
        Public STYLE_CODE As String = String.Empty
        Public COLOR_CODE As String = String.Empty
        Public RA_QTY As Int32 = 0
        Public RA_NET_PRICE As Decimal = 0
        Public SET_QTY As Int16 = 1
    End Class

    Private lstImportedFromExcel As New List(Of ItemsImported)

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Check_InquiryMode()
        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select SOTRMAF1.* from SOTRMAF1 where RA_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTRMAFX", "**", 0, False, "V", 1)
            With .Tables("SOTRMAFX").Columns
                .Add("CUST_NAME", GetType(System.String))
                .Add("RA_QTY", GetType(System.Int64))
                .Add("RA_QTY_OPEN", GetType(System.Int64))
                .Add("RA_QTY_USED", GetType(System.Int64))
                .Add("RA_QTY_CANC", GetType(System.Int64))
                .Add("RA_AMT", GetType(System.Decimal))
                .Add("RA_AMT_OPEN", GetType(System.Decimal))
                .Add("RA_AMT_USED", GetType(System.Decimal))
                .Add("RA_AMT_CANC", GetType(System.Decimal))
                .Add("RA_RETAIL_EXT", GetType(System.Decimal))
            End With

            Create_TDA(.Tables.Add, "SOTRMAF1", "*", 1)
            .Tables("SOTRMAF1").Columns.Add("RA_AMT", GetType(System.Decimal))
            .Tables("SOTRMAF1").Columns.Add("AR_PARM_KEY")

            ASCMAIN1.sql = "Select SOTRMAF2.*, ICTSTYL1.STYLE_DESC, NVL(ICTSTYL1.STYLE_COST, 0) STYLE_COST" _
            & " from SOTRMAF2,ICTSTYL1" _
            & " where ICTSTYL1.STYLE_CODE = SOTRMAF2.STYLE_CODE"
            Create_TDA(.Tables.Add, "SOTRMAF2", "**", 1)
            With .Tables("SOTRMAF2").Columns
                .Add("RA_AMT", GetType(System.Decimal), "ISNULL(RA_QTY,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_OPEN", GetType(System.Decimal), "ISNULL(RA_QTY_OPEN,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_USED", GetType(System.Decimal), "ISNULL(RA_QTY_USED,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_CANC", GetType(System.Decimal), "ISNULL(RA_QTY_CANC,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_RETAIL_EXT", GetType(System.Decimal), "IIF(ISNULL(RA_QTY,0)=0,RA_AMT/((100 - 0) / 100),ISNULL(RA_QTY,0) * ISNULL(RA_RETAIL,0))")
                .Add("ORDR_QTY_SHIP", GetType(System.Int32))
                .Add("RA_QTY_AVAIL", GetType(System.Int32))
            End With

            Create_TDA(.Tables.Add, "SOTRMAFL", "*", 1)

            With .Tables.Add("SOTRMAFT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("KEY")}
            End With

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "ICTSTYC1", "*")
            Create_TDA(.Tables.Add, "ECTECOM1", "*")
            Fill_Records("ECTECOM1", String.Empty, True, "SELECT * FROM ECTECOM1")
            Create_TDA(.Tables.Add, "ECTESTY1", "*")

            Create_TDA(.Tables.Add, "ARTCUST1", "*", , False)
            Create_TDA(.Tables.Add, "ARTCUST2", "*", , False)
            Create_TDA(.Tables.Add, "ICTWHSE1", "*", , False)
            Create_TDA(.Tables.Add, "SOTSREP1", "*", , False)

            Create_TDA(.Tables.Add, "ARTOPEN1", "*", 0)
            Create_TDA(.Tables.Add, "SOTINVH1", "*", 0)
            Create_TDA(.Tables.Add, "SOTINVHM", "*", 0)

            ASCMAIN1.sql = "SELECT SOTORDR2.*, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_DEPT" _
                & " FROM SOTORDR1, SOTORDR2" _
                & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO " _
                & " AND SOTORDR2.ORDR_QTY_SHIP > 0 " _
                & " AND SOTORDR1.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2X", ASCMAIN1.sql, 0, False, "V", 2)

            Create_TDA(.Tables.Add, "EDTTRPM1", "*", 0)
            Fill_Records("EDTTRPM1", String.Empty, True, "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '180'")

            With .Tables.Add("SOTRMAF0")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("ADDRESS0")
                .Columns.Add("ADDRESS1")
                .Columns.Add("ADDRESS2")
                .Columns.Add("ADDRESS3")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() { .Columns("AR_PARM_KEY")}
            End With

            ASCMAIN1.sql = "SELECT SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_NO, SOTINVH1.ORDR_DEPT, SOTINVH2.INV_LNO, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" _
                 & ", SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_QTY_SHIP" _
                 & " from SOTINVH1, SOTINVH2" _
                 & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                 & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                 & "   and SOTINVH1.INV_TYPE = 'I'" _
                 & "   and SOTINVH1.CUST_CODE = :PARM1" _
                 & "   and SOTINVH1.ORDR_CUST_PO = :PARM2"
            Create_TDA(.Tables.Add, "SOTINVH2X", ASCMAIN1.sql, 0, False, "VV", 0)

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO, SOTINVH2.INV_LNO" & vbCrLf _
                & ", SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_QTY_SHIP" & vbCrLf _
                & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & " from SOTINVH2,ICTSTYL1,ICTCOLR1,SOTINVH1" & vbCrLf _
                & " where SOTINVH2.INV_TYPE = 'I' and SOTINVH2.CUST_CODE = :PARM1 and SOTINVH2.ORDR_YYYYPP_UPDATED >= :PARM2" & vbCrLf _
                & "  and SOTINVH2.STYLE_CODE = :PARM3 " & vbCrLf _
                & "  and ICTSTYL1.STYLE_CODE = SOTINVH2.STYLE_CODE and ICTCOLR1.COLOR_CODE = SOTINVH2.COLOR_CODE" & vbCrLf _
                & "  and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False, "VVV", 3)

        End With

        grdSOTRMAFX.DataSource = dst.Tables("SOTRMAFX")
        grdSOTRMAF2.DataSource = dst.Tables("SOTRMAF2")
        grdSOTRMAFT.DataSource = dst.Tables("SOTRMAFT")
        grdSOTRMAFL.DataSource = dst.Tables("SOTRMAFL")
        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        grdSOTRMAF2.DisplayLayout.Bands(0).Columns("ORDR_QTY_SHIP").DefaultCellValue = 0

        grdSOTRMAFX.DisplayLayout.UseFixedHeaders = True
        With grdSOTRMAFX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RA_NO", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdSOTRMAF2.DisplayLayout.UseFixedHeaders = True
        With grdSOTRMAF2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RA_LNO", "STYLE_CODE", "STYLE_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each COLUMN_NAME As String In New String() {"RA_NET_PRICE"}
                .Columns(COLUMN_NAME).Format = "#,###.000"
                .Columns(COLUMN_NAME).MaskInput = "nnnn.nnn"
            Next
        End With

        With grdSOTRMAFX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.Width = 80
                ElseIf New String() {"RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.Width = 90
                ElseIf New String() {"RA_RETAIL_EXT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.Width = 90
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With

        With grdSOTRMAF2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"STYLE_CODE", "COLOR_CODE", "RA_QTY", "RA_QTY_OPEN", "RA_NET_PRICE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_CODE", "COLOR_CODE", "STYLE_DESC", "STYLE_RETAIL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"RA_NET_PRICE", "RA_RETAIL_EXT", "RA_LINE_AMT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With

        Create_Summary(grdSOTRMAFX, "RA_NO", "Count")
        Create_Summary(grdSOTRMAFX, New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC",
                                          "RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC", "RA_RETAIL_EXT"})

        Create_Summary(grdSOTRMAF2, "RA_LNO", "Count")
        Create_Summary(grdSOTRMAF2, New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC",
                                                  "RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC", "RA_LINE_AMT", "RA_RETAIL_EXT"})

        With dst.Tables("SOTRMAFT").Rows
            .Add(New Object() {1, "Auth", 0, 0})
            .Add(New Object() {2, "Open", 0, 0})
            .Add(New Object() {3, "Used", 0, 0})
            .Add(New Object() {4, "Canc", 0, 0})
        End With
        Sort_grdColumns(grdSOTRMAFT, "KEY", True)

        Dim rowSOTRMAF0 As DataRow = dst.Tables("SOTRMAF0").NewRow
        With ROWs("ARTPARM1")
            rowSOTRMAF0.Item("AR_PARM_KEY") = "Z"
            rowSOTRMAF0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowSOTRMAF0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowSOTRMAF0.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            rowSOTRMAF0.Item("REMIT3") = "Tel " & .Item("AR_PARM_REMIT_PHONE") & " Fax " & .Item("AR_PARM_REMIT_FAX")
            'rowSOTRMAF0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
        End With
        rowSOTRMAF0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        dst.Tables("SOTRMAF0").Rows.Add(rowSOTRMAF0)


        Show_Filter(grdSOTRMAFX, True)
        grdSOTRMAFX.DisplayLayout.GroupByBox.Hidden = False
        SplitContainer1.Panel2Collapsed = Not ASCMAIN1.CLIENT = "RGI"

        ASCMAIN1.Add_Value_List(grdSOTRMAFX, "RA_REASON_CODE", "SELECT * FROM SOTRMAR1")

        ASCMAIN1.Add_Value_List(grdSOTRMAFX, "RA_STATUS", Nothing, New String() {":", "O:Open", "F:Completed", "D:Deleted", "C:Cancelled"})

        ASCMAIN1.Add_Value_List(grdSOTRMAFL, "CARRIER_CODE")
        ASCMAIN1.Add_Value_List(grdSOTRMAFL, "SHIP_VIA_CODE")
        ASCMAIN1.Add_Value_List(grdSOTRMAFL, "DELIVERY_METHOD", Nothing, New String() {":", "N:Carrier Mail", "U:Carrier email", "P:Printed", "E:email to Customer", "1:UPS 1 attempt", "3:UPS 3 attempts"})

        optRA_REASON_CODE.Items.Clear()
        Dim tblSOTRMAR1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTRMAR1")
        For Each row As DataRow In tblSOTRMAR1.Select("", "RA_REASON_DESC")
            Dim item As New Infragistics.Win.ValueListItem
            item.DataValue = row.Item("RA_REASON_CODE")
            item.DisplayText = row.Item("RA_REASON_DESC")
            If item.DataValue = "X" OrElse item.DataValue = "C" Then
                item.Appearance.ForeColor = Drawing.Color.Red
            Else
                item.Appearance.ForeColor = Drawing.Color.Black
            End If
            optRA_REASON_CODE.Items.Add(item)
        Next

        If ASCMAIN1.CLIENT = "RGI" Then
            lblClaimNo.Text = "Invoice No"
        End If

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFRMAFI" OrElse MENU_ITEM_OBJECT = "SOFRMAFW")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                CUST_CLAIM_NO = Absx1.txtFor("CUST_CLAIM_NO").Text
                ORDR_NO = Absx1.txtFor("ORDR_NO").Text
                is180Customer = False
                Dim rowSOTORDR1 As DataRow = Nothing

                If ASCMAIN1.CLIENT = "RGI" AndAlso CUST_CLAIM_NO.Length > 0 Then
                    Dim INV_NO As String = CUST_CLAIM_NO
                    INV_NO = INV_NO.Replace("'", "")
                    INV_NO = INV_NO.PadLeft(10, "0")

                    Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTINVH1 WHERE INV_TYPE = 'I' AND INV_NO = :PARM1", "V", INV_NO)
                    If rowSOTINVH1 IsNot Nothing Then
                        If CUST_CODE.Length > 0 Then
                            If rowSOTINVH1.Item("CUST_CODE") & String.Empty <> CUST_CODE Then
                                rowSOTINVH1 = Nothing
                            End If
                        End If
                    End If

                    If rowSOTINVH1 IsNot Nothing Then
                        Absx1.txtFor("CUST_CODE").Text = rowSOTINVH1.Item("CUST_CODE") & String.Empty
                        Absx1.txtFor("CUST_CLAIM_NO").Text = INV_NO
                        Absx1.txtFor("ORDR_NO").Text = rowSOTINVH1.Item("ORDR_NO") & String.Empty

                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        CUST_CLAIM_NO = Absx1.txtFor("CUST_CLAIM_NO").Text
                        ORDR_NO = Absx1.txtFor("ORDR_NO").Text

                    End If
                End If

                tblSOTRMAF2.Rows.Clear()

                If CUST_CLAIM_NO.Length > 0 AndAlso CUST_CODE.Length > 0 Then
                    ASCMAIN1.sql = " SELECT SOTRMAF2.STYLE_CODE, SOTRMAF2.COLOR_CODE, SUM(NVL(SOTRMAF2.RA_QTY_OPEN, 0) + NVL(SOTRMAF2.RA_QTY_USED, 0)) RA_QTY_USED "
                    ASCMAIN1.sql &= " FROM SOTRMAF1, SOTRMAF2"
                    ASCMAIN1.sql &= " WHERE SOTRMAF1.RA_NO = SOTRMAF2.RA_NO"
                    ASCMAIN1.sql &= " AND SOTRMAF1.CUST_CODE = :PARM1"
                    ASCMAIN1.sql &= " AND SOTRMAF1.CUST_CLAIM_NO = :PARM2"
                    ASCMAIN1.sql &= " GROUP BY SOTRMAF2.STYLE_CODE, SOTRMAF2.COLOR_CODE"
                    ASCMAIN1.sql &= " HAVING SUM(NVL(SOTRMAF2.RA_QTY_OPEN, 0) + NVL(SOTRMAF2.RA_QTY_USED, 0)) > 0"
                    tblSOTRMAF2 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTRMAF2", "VV", New Object() {CUST_CODE, CUST_CLAIM_NO})
                End If

                If tblSOTRMAF2.Rows.Count > 0 Then
                    If MessageBox.Show("There are other RMAs against this Customer / Claim No. Do you want these item quantities excluded from this RMA?", "New RMA", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        tblSOTRMAF2.Rows.Clear()
                    End If
                End If

                If CUST_CODE.Length = 0 Then
                    EMsg &= vbCr & "You Must Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                        CUST_CODE = String.Empty
                    End If
                End If

                If CUST_CLAIM_NO = "" Then
                    EMsg &= vbCr & "You Must Provide a Value for Customer Claim No"
                End If

                If CUST_CODE.Length > 0 AndAlso ORDR_NO.Length > 0 Then
                    rowSOTORDR1 = LookUp("SOTORDR1", ORDR_NO)
                    If rowSOTORDR1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Order Number"
                    ElseIf rowSOTORDR1.Item("CUST_CODE") <> CUST_CODE Then
                        EMsg &= vbCr & "Invalid Order Number for this Customer"
                    End If
                End If

                If EMsg = "" Then
                    If dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                        is180Customer = True

                        If ORDR_NO.Length = 0 Then
                            EMsg &= vbCr & "Order Number is required for all EDI 180 customers"
                        Else
                            Fill_Records("SOTORDR2X", ORDR_NO)
                            If dst.Tables("SOTORDR2X").Rows.Count = 0 Then
                                EMsg &= vbCr & "The provided Sales Order Number does not have any shipped items."
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTRMAF1", CUST_CODE) Then Exit Sub
                End If

            Case "Import From Excel"
                CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                CUST_CLAIM_NO = Absx1.txtFor("CUST_CLAIM_NO").Text

                If CUST_CODE.Length = 0 Then
                    EMsg &= vbCr & "You Must Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                        CUST_CODE = String.Empty
                    End If
                End If

                If CUST_CLAIM_NO = "" Then
                    EMsg &= vbCr & "You Must Provide a Value for Customer Claim No"
                End If

                If EMsg.Length > 0 Then
                    Exit Select
                End If

                Dim userInstructions As String = "The Excel Workbook must meet the following criteria: " _
                    & Environment.NewLine & Environment.NewLine & "1) Worksheet named Details" _
                    & Environment.NewLine & Environment.NewLine & "2) Three columns on the Details Worksheet labeled: Vendor SKU, Qty and Cost." _
                    & Environment.NewLine & Environment.NewLine & "3) Vendor SKU column must split the Style Code, Color Code using an underscore (_), hyphen (-) or space. Example: MTX58549-RED" _
                    & Environment.NewLine & Environment.NewLine & "4) The first row in the Excel worksheet must be a title row containing the 3 column headers identified above." _
                    & Environment.NewLine & Environment.NewLine & "The import stops when it finds a blank value in the Vendor SKU column. Therefore, make sure there are no blank rows in the data." _
                    & Environment.NewLine & Environment.NewLine & "Do you want to continue?"

                If MessageBox.Show(userInstructions, "Import From Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                If Not ASCMAIN1.Logical_Lock("SOTRMAF1", CUST_CODE) Then Exit Sub

            Case "Edit", "View"

                CUST_CODE = ""
                RA_NO = ""
                is180Customer = False

                If Absx1.txtFor("RA_NO").Text = "" Then
                    EMsg &= vbCr & "No Returns Authorization No Specified"
                Else
                    RA_NO = Absx1.txtFor("RA_NO").Text
                    rowSOTRMAF1 = LookUp("SOTRMAF1", RA_NO)
                    If rowSOTRMAF1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Returns Authorization No " & RA_NO
                    Else
                        CUST_CODE = rowSOTRMAF1.Item("CUST_CODE")

                        If rowSOTRMAF1.Item("RA_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Dim msg As String = ""
                            Select Case rowSOTRMAF1.Item("RA_STATUS")
                                Case "C"
                                    msg = "Returns Authorization No " & RA_NO & " has been Cancelled"
                                Case "D"
                                    msg = "Returns Authorization No " & RA_NO & " has been Deleted"
                                Case Else ' such as "F"
                                    msg = "Returns Authorization No " & RA_NO & " is No Longer Open"
                            End Select

                            EMsg &= vbCr & msg
                        Else
                            ' See if there are any Returns against the RMA, if so, you cannot edit
                            If eItemKey = "Edit" AndAlso ASCDATA1.GetDataTable("Select * from SOTRTRN1 where RA_NO = '" & RA_NO & "'").Rows.Count > 0 Then
                                EMsg = "Returns Authorization No " & RA_NO & " has Returns applied to it."
                            End If
                        End If

                        If EMsg.Length = 0 AndAlso eItemKey = "Edit" Then
                            If Not ASCMAIN1.Logical_Lock("SOTRMAF1", RA_NO) Then Exit Sub
                            If Not ASCMAIN1.Logical_Lock("SOTRMAF1", CUST_CODE) Then Exit Sub

                            Dim CUST_CLAIM_NO As String = Absx1.txtFor("CUST_CLAIM_NO").Text
                            If CUST_CLAIM_NO.Length > 0 AndAlso CUST_CODE.Length > 0 Then
                                ASCMAIN1.sql = " SELECT SOTRMAF2.STYLE_CODE, SOTRMAF2.COLOR_CODE, SUM(NVL(SOTRMAF2.RA_QTY_OPEN, 0) + NVL(SOTRMAF2.RA_QTY_USED, 0)) RA_QTY_USED "
                                ASCMAIN1.sql &= " FROM SOTRMAF1, SOTRMAF2"
                                ASCMAIN1.sql &= " WHERE SOTRMAF1.RA_NO = SOTRMAF2.RA_NO"
                                ASCMAIN1.sql &= " AND SOTRMAF1.CUST_CODE = :PARM1"
                                ASCMAIN1.sql &= " AND SOTRMAF1.CUST_CLAIM_NO = :PARM2"
                                ASCMAIN1.sql &= " AND SOTRMAF1.RA_NO <> '" & Absx1.txtFor("RA_NO").Text & "'"
                                ASCMAIN1.sql &= " GROUP BY SOTRMAF2.STYLE_CODE, SOTRMAF2.COLOR_CODE"
                                ASCMAIN1.sql &= " HAVING SUM(NVL(SOTRMAF2.RA_QTY_OPEN, 0) + NVL(SOTRMAF2.RA_QTY_USED, 0)) > 0"
                                tblSOTRMAF2 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTRMAF2", "VV", New Object() {CUST_CODE, CUST_CLAIM_NO})
                            End If
                        End If
                    End If
                End If

                If EMsg <> "" Then ASCMAIN1.MultiTask_Release()

                If dst.Tables("EDTTRPM1").Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                    is180Customer = True
                End If

            Case "Update"
                If Absx1.dteFor("RA_DATE").Value & "" = "" _
                    Or Absx1.dteFor("RA_EXPIRE").Value & "" = "" Then
                    EMsg &= vbCr & "RA Date and Expiration Date are Mandatory"
                Else
                    If Format(Absx1.dteFor("RA_DATE").Value, "yyyyMMdd") _
                     > Format(Absx1.dteFor("RA_EXPIRE").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Expiration Date cannot be Prior to RA Date"
                    End If
                End If

                If Absx1.txtFor("CUST_STORE_NO").Text = "" Then
                    ' NO BIGGIE
                Else
                    If LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, "MK", Absx1.txtFor("CUST_STORE_NO").Text}) Is Nothing Then
                        EMsg &= vbCr & "Invalid Store No"
                    End If
                End If

                If Absx1.txtFor("REASON_CODE").Text = "" Then
                    EMsg &= vbCr & "Rtn Reason is required"
                Else
                    If LookUp("ARTREAS1", Absx1.txtFor("REASON_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Rtn Reason"
                    End If
                End If

                If Absx1.optFor("RA_REASON_CODE").Value = "" Then
                    EMsg &= vbCr & "RA Reason is required"
                Else
                    If LookUp("SOTRMAR1", Absx1.optFor("RA_REASON_CODE").Value) Is Nothing Then
                        EMsg &= vbCr & "Invalid RA Reason"
                    Else
                        If (Absx1.optFor("RA_REASON_CODE").Value = "X" OrElse (Absx1.optFor("RA_REASON_CODE").Value = "C") AndAlso ASCMAIN1.CLIENT <> "RGI") Then

                            Dim DT As Date = Absx1.dteFor("RA_DATE").Value
                            If DT & "" = "" Then
                                EMsg &= vbCr & "Document Date is Mandatory"
                            Else
                                TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                            End If

                            If Absx1.txtFor("CUST_CLAIM_NO").Text = "" Then
                                EMsg &= vbCr & "Customer Claim No Required for Reason Code Indicated"
                            Else
                                ASCMAIN1.sql = "" _
                                    & "Select * from ARTOPEN1 where CUST_CODE = '" & CUST_BILL_TO_CUST & "'" _
                                    & " and INV_TYPE = 'C' and INV_CUST_PO = '" & Absx1.txtFor("CUST_CLAIM_NO").Text & "'" _
                                    & " union " _
                                    & "Select * from ARTOPENX where CUST_CODE = '" & CUST_BILL_TO_CUST & "'" _
                                    & " and INV_TYPE = 'C' and INV_CUST_PO = '" & Absx1.txtFor("CUST_CLAIM_NO").Text & "'"
                                If ASCDATA1.GetDataTable.Rows.Count <> 0 Then
                                    If MessageBox.Show("Customer Claim No has already been used on an existing Credit. Proceed w/Generation of Credit Anyway?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' RELAXING THIS CONSTRAINT FOR NOW - CONVERTED DATA HAD NO CLAIM
                'If Absx1.txtFor("CUST_CLAIM_NO").Text = "" Then
                '    EMsg &= vbCr & "Customer Claim No is required"
                'End If

                If grdSOTRMAF2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Returns Authorization"
                Else
                    If Val(dst.Tables("SOTRMAF2").Compute("COUNT(RA_LNO)", "RA_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Items on Returns Authorization with Qty >0"
                    End If
                End If

                If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
                    EMsg &= vbCr & "Sales Division is required"
                Else
                    Validate_Code("SALES_DIVISION_CODE")
                End If

                If (Absx1.optFor("RA_REASON_CODE").Value = "X" OrElse (Absx1.optFor("RA_REASON_CODE").Value = "C") AndAlso ASCMAIN1.CLIENT = "RGI" AndAlso EMsg.Length = 0) Then
                    If MessageBox.Show("You choose reason " & optRA_REASON_CODE.CheckedItem.DisplayText & ". This does not require the customer to return the goods." _
                                       & Environment.NewLine & "Do you want to continue?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTRMAF2 where RA_QTY_USED <> 0"
                ASCMAIN1.sql &= " and RA_NO = '" & RA_NO & "'"

                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Returns Authorization has been Used - Delete not permitted"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Returns Authorization as Deleted",
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Cancel Balance"
                If EMsg = "" Then
                    If MsgBox("Do you want to Cancel (the remaining open balance on) this Returns Authorization",
                               vbYesNo, "Confirmation") = MsgBoxResult.No Then
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

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Order()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Cancel Balance"
                Cancel_Order()
                Mode_Settings(False)

            Case "Print Credit Memo"
                Print_Credit_Memo()

            Case "Approve RMA"
                Absx1.txtFor("APPROVED_BY").Text = ASCMAIN1.USER_ID
                Absx1.dteFor("DATE_APPROVED").DateTime = DateTime.Now.ToShortDateString
                dst.Tables("SOTRMAF1").Rows(0).Item("APPROVED_BY") = ASCMAIN1.USER_ID
                dst.Tables("SOTRMAF1").Rows(0).Item("DATE_APPROVED") = DateTime.Now.ToShortDateString

            Case "Import From Excel"
                If ImportFromExcel() Then
                    Mode_Settings(True)
                End If

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                .Items("Import From Excel").Settings.Enabled = not_iScreenMode

                If InquiryMode Then
                    .Items("Import From Excel").Visible = False
                End If

                If (EntryMode = "V" And ScreenMode) Then
                    If rowSOTRMAF1.Item("RA_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If

                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Approve RMA").Settings.Enabled = iScreenMode

                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode

                .Items("Cancel Balance").Settings.Enabled = iScreenMode


                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Print").Visible = (EntryMode = "V" And ScreenMode)

                Select Case ASCMAIN1.CLIENT
                    Case "NYA", "RGI"
                        ' valid options
                    Case Else
                        .Items("Print").Visible = False
                End Select


                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)

                .Items("Approve RMA").Visible = Not InquiryMode AndAlso EntryMode <> "V" AndAlso ScreenMode _
                    AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains("CR") _
                    AndAlso Absx1.txtFor("APPROVED_BY").TextLength = 0

                .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)

                .Items("Cancel Balance").Visible = (EntryMode = "E")
                .Items("Print Credit Memo").Visible = (InquiryMode OrElse EntryMode = "V" OrElse ScreenMode) _
                            AndAlso (rowSOTRMAF1 IsNot Nothing _
                                     AndAlso rowSOTRMAF1.Item("RA_STATUS") & String.Empty = "F" _
                                     AndAlso rowSOTRMAF1.Item("RA_REASON_CODE") & String.Empty = "X" _
                                     AndAlso rowSOTRMAF1.Item("INV_TYPE") & String.Empty <> String.Empty _
                                     AndAlso rowSOTRMAF1.Item("INV_NUM") & String.Empty <> String.Empty)

            End With

            .Groups("Totals").Visible = ScreenMode
            .Groups("Status").Visible = Not ScreenMode And InquiryMode
        End With

        lblStatus.Visible = ScreenMode

        grdSOTRMAFX.Visible = Not tf

        lblCredit.Visible = ScreenMode AndAlso (rowSOTRMAF1.Item("INV_NUM") & "" <> "")
        lblINV_NUM.Visible = ScreenMode AndAlso (rowSOTRMAF1.Item("INV_NUM") & "" <> "")
        'If ScreenMode Then
        '    lblINV_NUM.Text = ""
        'End If

        Absx1.optFor("RA_REASON_CODE").Visible = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CLAIM_NO"), InquiryMode Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")))
        Dim RA_was_used As Boolean = (dst.Tables("SOTRMAF2").Select("RA_QTY_USED <> 0").Length > 0)

        Set_Read_Only_for_ctl(Absx1.optFor("RA_REASON_CODE"), InquiryMode Or RA_was_used Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")))
        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))

        grdSOTRMAF2.DisplayLayout.Bands(0).Columns("X").Hidden = Not (ScreenMode And (EntryMode = "E"))
        grdSOTRMAF2.DisplayLayout.Bands(0).Columns("RA_LINE_AMT").Hidden = True ' THIS COL IS FOR A DIFFERENT DESIGN

        If ScreenMode Then
            If EntryMode = "V" Then
                grdSOTRMAF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTRMAF2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTRMAF2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                If is180Customer Then
                    grdSOTRMAF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grdSOTRMAF2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTRMAF2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    grdSOTRMAF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    grdSOTRMAF2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTRMAF2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If

                If EntryMode <> "E" Then
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("X").Hidden = True
                Else
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("X").Hidden = False
                End If
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("RA_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("CUST_CLAIM_NO").Text = ""
        Absx1.txtFor("ORDR_NO").Text = ""

        RA_NO = String.Empty
        CUST_CODE = String.Empty
        CUST_CLAIM_NO = String.Empty
        ORDR_NO = String.Empty
        CUST_BILL_TO_CUST = String.Empty
        is180Customer = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTRMAF1", "SOTRMAF2", "ARTOPEN1", "SOTINVH1", "SOTINVH2X", "SOTORDR2X", "SOTRMAFL", "SOTINVHX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Load_SOTRMAFX()
        tblSOTRMAF2.Rows.Clear()

        importedFromExcel = False
        loadedFromExcel = False
        lstImportedFromExcel.Clear()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            RA_NO = ASCMAIN1.Next_Control_No("SOTRMAF1.RA_NO")

            rowSOTRMAF1 = dst.Tables("SOTRMAF1").NewRow
            With rowSOTRMAF1
                .Item("RA_NO") = RA_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_CLAIM_NO") = CUST_CLAIM_NO
                .Item("ORDR_NO") = ORDR_NO
                .Item("RA_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("RA_REASON_CODE") = "Z"

                Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
                If WHSE_CODE = "" Then WHSE_CODE = ""
                .Item("WHSE_CODE") = WHSE_CODE
                '  .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
                .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & ""
                .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE") & ""

                If ORDR_NO.Length > 0 Then
                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                    If rowSOTORDR1 IsNot Nothing Then
                        .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE") & String.Empty
                        .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE") & ""
                        .Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE") & ""
                    End If
                End If
            End With
            dst.Tables("SOTRMAF1").Rows.Add(rowSOTRMAF1)
        Else
            rowSOTRMAF1 = Fill_Record("SOTRMAF1", RA_NO)
        End If

        Fill_Record("SOTRMAFL", RA_NO)
        If dst.Tables("SOTRMAFL").Rows.Count > 0 Then
            grdSOTRMAFL.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        End If

        ' Load screen with items found on Invoices for the given Customer, Customer PO
        If Not is180Customer Then
            If CUST_CLAIM_NO.Length = 0 Then
                CUST_CLAIM_NO = Absx1.txtFor("CUST_CLAIM_NO").Text
            End If
            Fill_Records("SOTINVH2X", New Object() {CUST_CODE, CUST_CLAIM_NO})
            dst.Tables("SOTORDR2X").Rows.Clear()

            If dst.Tables("SOTINVH2X").Rows.Count = 0 AndAlso ASCMAIN1.CLIENT = "RGI" Then
                Dim INV_NO As String = CUST_CLAIM_NO
                INV_NO = CUST_CLAIM_NO.Replace("'", "")
                INV_NO = INV_NO.PadLeft(10, "0")

                If ORDR_NO.Length = 0 Then
                    ORDR_NO = Absx1.txtFor("ORDR_NO").Text
                End If

                ASCMAIN1.sql = "SELECT SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_NO, SOTINVH1.ORDR_DEPT, SOTINVH2.INV_LNO, SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE" _
                     & ", SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_QTY_SHIP" _
                     & " from SOTINVH1, SOTINVH2" _
                     & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                     & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                     & "   and SOTINVH1.INV_TYPE = 'I'" _
                     & "   and SOTINVH1.CUST_CODE = '" & CUST_CODE & "'" _
                     & "   and SOTINVH1.ORDR_NO = '" & ORDR_NO & "'" _
                     & "   and SOTINVH1.INV_NO = '" & INV_NO & "'"
                Fill_Records("SOTINVH2X", String.Empty, True, ASCMAIN1.sql)
                If dst.Tables("SOTINVH2X").Rows.Count > 0 Then
                    rowSOTRMAF1.Item("CUST_CLAIM_NO") = INV_NO
                End If

            End If
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH2X"), New String() {"ORDR_NO"}).Rows
                Fill_Records("SOTORDR2X", ORDR_NO, False)
            Next
        End If

        If ORDR_NO.Length > 0 AndAlso ASCMAIN1.CLIENT = "RGI" AndAlso dst.Tables("SOTORDR2X").Rows.Count = 0 Then
            Fill_Records("SOTORDR2X", ORDR_NO)
            If dst.Tables("SOTORDR2X").Rows.Count > 0 Then
                MyBase.Absx1.txtFor("CUST_STORE_NO").Text = dst.Tables("SOTORDR2X").Rows(0).Item("CUST_STORE_NO") & String.Empty
                rowSOTRMAF1.Item("CUST_STORE_NO") = dst.Tables("SOTORDR2X").Rows(0).Item("CUST_STORE_NO") & String.Empty
            End If
        End If

        If EntryMode = "N" AndAlso dst.Tables("SOTINVH2X").Rows.Count > 0 Then
            Dim minCUST_STORE_NO As String = dst.Tables("SOTINVH2X").Compute("MIN(CUST_STORE_NO)", "") & String.Empty
            Dim maxCUST_STORE_NO As String = dst.Tables("SOTINVH2X").Compute("MAX(CUST_STORE_NO)", "") & String.Empty
            If minCUST_STORE_NO.Length > 0 AndAlso minCUST_STORE_NO = maxCUST_STORE_NO Then
                MyBase.Absx1.txtFor("CUST_STORE_NO").Text = minCUST_STORE_NO
                rowSOTRMAF1.Item("CUST_STORE_NO") = minCUST_STORE_NO
            End If
        ElseIf EntryMode = "N" AndAlso is180Customer Then
            MyBase.Absx1.txtFor("CUST_STORE_NO").Text = dst.Tables("SOTORDR2X").Rows(0).Item("CUST_STORE_NO") & String.Empty
            rowSOTRMAF1.Item("CUST_STORE_NO") = dst.Tables("SOTORDR2X").Rows(0).Item("CUST_STORE_NO") & String.Empty
        End If

        CUST_CODE = rowSOTRMAF1.Item("CUST_CODE")
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = CUST_CODE
        rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

        If ASCMAIN1.CLIENT = "RGI" And EntryMode = "N" Then
            rowSOTRMAF1.Item("RA_DATE") = DateTime.Now.ToShortDateString
            rowSOTRMAF1.Item("RA_EXPIRE") = DateAdd(DateInterval.Month, 6, DateTime.Now).ToShortDateString
            rowSOTRMAF1.Item("REQUIRES_CALL_TAGS") = "1"

            If importedFromExcel Then
                rowSOTRMAF1.Item("REQUIRES_CALL_TAGS") = "0"
            End If

            Fill_Records("ARTCUST2", New String() {CUST_CODE, "MK", Absx1.txtFor("CUST_STORE_NO").Text})
            If dst.Tables("ARTCUST2").Rows.Count > 0 Then
                Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows(0)
                rowSOTRMAF1.Item("RA_EMAIL") = rowARTCUST2.Item("CUST_EMAIL") & String.Empty
                rowSOTRMAF1.Item("RA_CONTACT") = rowARTCUST2.Item("CUST_CONTACT") & String.Empty
                rowSOTRMAF1.Item("RA_PHONE") = rowARTCUST2.Item("CUST_PHONE") & String.Empty
                '.Item("RA_CARRIER_CODE") = "1"
                '.Item("RA_CARTONS") = "1"
            ElseIf rowARTCUST1 IsNot Nothing Then
                rowSOTRMAF1.Item("RA_EMAIL") = rowARTCUST1.Item("CUST_EMAIL") & String.Empty
                rowSOTRMAF1.Item("RA_CONTACT") = rowARTCUST1.Item("CUST_CONTACT") & String.Empty
                rowSOTRMAF1.Item("RA_PHONE") = rowARTCUST1.Item("CUST_PHONE") & String.Empty
            End If
        End If

        Fill_Records("SOTRMAF2", RA_NO)
        Sort_grdColumns(grdSOTRMAF2, "RA_LNO")

        lblINIT_DATE.Text = "Entered on " & Format(rowSOTRMAF1.Item("INIT_DATE"), "MM/dd/yyyy")

        If EntryMode = "N" Then
            lblStatus.Text = "New"

            Dim loadFromInvoice As Boolean = False

            If Not is180Customer Then
                If dst.Tables("SOTORDR2X").Rows.Count > 0 Then
                    If MessageBox.Show("Load items from Invoice?", "Load", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                        loadFromInvoice = True
                    End If
                End If
            End If

            If is180Customer OrElse loadFromInvoice Then
                Absx1.txtFor("ORDR_DEPT").Text = dst.Tables("SOTORDR2X").Rows(0).Item("ORDR_DEPT") & String.Empty
                Absx1.dteFor("RA_DATE").DateTime = DateTime.Now.ToShortDateString

                rowSOTRMAF1.Item("ORDR_DEPT") = dst.Tables("SOTORDR2X").Rows(0).Item("ORDR_DEPT") & String.Empty
                rowSOTRMAF1.Item("RA_DATE") = DateTime.Now.ToShortDateString

                If ASCMAIN1.CLIENT = "RGI" Then
                    Absx1.dteFor("RA_EXPIRE").DateTime = DateAdd(DateInterval.Month, 6, DateTime.Now).ToShortDateString
                    rowSOTRMAF1.Item("RA_EXPIRE") = DateAdd(DateInterval.Month, 6, DateTime.Now).ToShortDateString
                End If

                grdSOTRMAF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes

                Dim tableName As String = "SOTORDR2X"
                Dim sortBy As String = "ORDR_LNO"

                ' Regency has back orders; therefore, use only the invoice not the sales order.
                If ASCMAIN1.CLIENT = "RGI" Then
                    tableName = "SOTINVH2X"
                    sortBy = "INV_LNO"
                End If

                For Each rowSOTORDR2X As DataRow In dst.Tables(tableName).Select("", sortBy)

                    If Val(rowSOTORDR2X.Item("ORDR_QTY_SHIP") & String.Empty) = 0 Then
                        Continue For
                    End If

                    Dim STYLE_CODE As String = rowSOTORDR2X.Item("STYLE_CODE") & String.Empty
                    Dim COLOR_CODE As String = rowSOTORDR2X.Item("COLOR_CODE") & String.Empty
                    If tblSOTRMAF2.Rows.Count > 0 AndAlso tblSOTRMAF2.Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'").Length > 0 Then
                        Dim RA_QTY_USED As Int32 = Val(tblSOTRMAF2.Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("RA_QTY_USED") & String.Empty)
                        RA_QTY_USED = Val(rowSOTORDR2X.Item("ORDR_QTY_SHIP") & String.Empty) - RA_QTY_USED
                        If RA_QTY_USED <= 0 Then Continue For
                        rowSOTORDR2X.Item("ORDR_QTY_SHIP") = RA_QTY_USED
                    End If

                    grdSOTRMAF2.DisplayLayout.Bands(0).AddNew()
                    With grdSOTRMAF2.ActiveRow
                        .Cells("STYLE_CODE").Value = rowSOTORDR2X.Item("STYLE_CODE") & String.Empty
                        .Cells("COLOR_CODE").Value = rowSOTORDR2X.Item("COLOR_CODE") & String.Empty
                        .Cells("RA_QTY").Value = rowSOTORDR2X.Item("ORDR_QTY_SHIP") & String.Empty
                        .Cells("RA_NET_PRICE").Value = rowSOTORDR2X.Item("ORDR_UNIT_PRICE") & String.Empty
                        .Cells("ORDR_NO").Value = ORDR_NO
                        .Cells("ORDR_LNO").Value = rowSOTORDR2X.Item(sortBy) & String.Empty
                        .Cells("RA_QTY_AVAIL").Value = rowSOTORDR2X.Item("ORDR_QTY_SHIP") & String.Empty
                        .Cells("ORDR_QTY_SHIP").Value = rowSOTORDR2X.Item("ORDR_QTY_SHIP") & String.Empty
                        .Update()
                    End With
                Next
            ElseIf importedFromExcel Then
                For Each custReturn As ItemsImported In lstImportedFromExcel
                    grdSOTRMAF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                    grdSOTRMAF2.DisplayLayout.Bands(0).AddNew()
                    With grdSOTRMAF2.ActiveRow
                        .Cells("STYLE_CODE").Value = custReturn.STYLE_CODE
                        .Cells("COLOR_CODE").Value = custReturn.COLOR_CODE
                        .Cells("RA_QTY").Value = custReturn.RA_QTY
                        .Cells("RA_NET_PRICE").Value = custReturn.RA_NET_PRICE
                        .Cells("RA_QTY_OPEN").Value = custReturn.RA_QTY
                        .Cells("RA_QTY_USED").Value = 0
                        .Cells("RA_QTY_CANC").Value = 0
                        .Cells("RA_QTY_AVAIL").Value = custReturn.RA_QTY
                        .Update()
                    End With
                Next
            End If
        Else
            Select Case rowSOTRMAF1.Item("RA_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
                Case "F"
                    lblStatus.Text = "Completed"
                Case Else
                    lblStatus.Text = "?"
            End Select

            ' Set Max qty that can be returned
            If tblSOTRMAF2.Columns.Count > 0 Then
                For Each rowSOTINVH2X As DataRow In dst.Tables("SOTINVH2X").Select("")
                    Dim STYLE_CODE As String = rowSOTINVH2X.Item("STYLE_CODE") & String.Empty
                    Dim COLOR_CODE As String = rowSOTINVH2X.Item("COLOR_CODE") & String.Empty

                    If dst.Tables("SOTRMAF2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'").Length > 0 Then
                        Dim rowSOTRMAF2 As DataRow = dst.Tables("SOTRMAF2").Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0)
                        Dim ORDR_QTY_SHIP As Int32 = Val(rowSOTINVH2X.Item("ORDR_QTY_SHIP") & String.Empty)
                        Dim RA_QTY As Int32 = Val(rowSOTRMAF2.Item("RA_QTY") & String.Empty)

                        Dim RA_QTY_USED As Int32 = 0
                        If tblSOTRMAF2.Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'").Length > 0 Then
                            RA_QTY_USED = Val(tblSOTRMAF2.Select("STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("RA_QTY_USED") & String.Empty)
                        End If

                        rowSOTRMAF2.Item("RA_QTY_AVAIL") = ORDR_QTY_SHIP - RA_QTY_USED
                        If rowSOTRMAF2.Item("RA_QTY_AVAIL") <= 0 Then
                            rowSOTRMAF2.Item("RA_QTY_AVAIL") = RA_QTY
                        End If

                    End If
                Next
            End If
        End If

        With grdSOTRMAF2.DisplayLayout.Bands(0)
            If (EntryMode = "E" Or EntryMode = "N") Then
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                If EntryMode = "E" Then
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Else
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdSOTRMAF2.DisplayLayout.Override
                If is180Customer Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                End If
            End With
            grdSOTRMAF2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, False)
        Else
            With grdSOTRMAF2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            grdSOTRMAF2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, True)
        End If

        Display_Totals()
        EnforceConstraints(True)

        If importedFromExcel Then
            loadedFromExcel = True
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, RA_NO)
        For Each TABLE_NAME As String In New String() _
            {"SOTRMAF1", "SOTRMAF2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where RA_NO = '" & RA_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        If EntryMode <> "N" Then Delete_Records()

        rowSOTRMAF1.Item("RA_AMT") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT)", "") & "")

        If ASCMAIN1.CLIENT <> "RGI" Then
            If Absx1.optFor("RA_REASON_CODE").Value = "X" OrElse Absx1.optFor("RA_REASON_CODE").Value = "C" Then
                rowSOTRMAF1.Item("RA_STATUS") = "F"
                rowSOTRMAF1.Item("RA_DATE_CLOSED") = DATETIME_STAMP.Date
                For Each rowSOTRMAF2 As DataRow In dst.Tables("SOTRMAF2").Select("")
                    rowSOTRMAF2.Item("RA_QTY_USED") = rowSOTRMAF2.Item("RA_QTY")
                    rowSOTRMAF2.Item("RA_QTY_OPEN") = 0
                Next
                Record_AR_Item()
            End If
        End If

        INIT_LAST("SOTRMAF1", False, , True)
        Dim sqldelete As String = "RA_NO = '" & RA_NO & "'"
        Update_Record_TDA("SOTRMAF1", sqldelete)
        Update_Record_TDA("SOTRMAF2", sqldelete)
        Dependent_Updates(1, RA_NO)

        CommitTrans("Update Complete")

        If EntryMode = "N" AndAlso ASCMAIN1.CLIENT = "RGI" AndAlso dst.Tables("SOTINVH1").Rows.Count = 0 Then
            If MessageBox.Show("Do you want to Print a RMA to send to the customer?", "Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                Print_Record()
            End If
        End If

        Dim SendCallTagEmail As Boolean = False

        If EntryMode = "N" AndAlso ASCMAIN1.CLIENT = "RGI" AndAlso rowSOTRMAF1.Item("REQUIRES_CALL_TAGS") & String.Empty = "1" Then
            SendCallTagEmail = True
        End If

        If EntryMode = "E" AndAlso ASCMAIN1.CLIENT = "RGI" AndAlso rowSOTRMAF1.Item("REQUIRES_CALL_TAGS") & String.Empty = "1" Then
            Dim row As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTRMAF1 WHERE RA_NO = '" & RA_NO & "'")
            If row IsNot Nothing AndAlso Val(row.Item("REQUIRES_CALL_TAGS") & String.Empty) <> 1 Then
                SendCallTagEmail = True
            End If
        End If

        If SendCallTagEmail Then
            Try
                Dim clsASCNOTE1 As New TAC.ASCNOTE1("RTN_CTAG", dst)
                clsASCNOTE1.Note = "Call Tags required for Returns Authorization: " & RA_NO
                clsASCNOTE1.CreateComponents()
                clsASCNOTE1.EmailDocument()

            Catch ex As Exception
                MessageBox.Show("Error sending Call Tag email: " & ex.Message, "Call Tag Email", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If


        ' See if we need to issue credit card credit.
        'If dst.Tables("SOTINVH1").Rows.Count > 0 Then
        '    Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows(0)
        '    If rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty <> String.Empty AndAlso Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty) <> 0 Then
        '        ASCMAIN1.Progress("Processing CC Credit", "")
        '        Dim errorMessage As String = String.Empty
        '        If Not SOCMAIN1.IssueCredit(rowSOTINVH1.Item("INV_NO"), errorMessage) Then
        '            MessageBox.Show("Error Processing Credit Card Refund: " & errorMessage, "CC Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        End If
        '        ASCMAIN1.Progress("", "")
        '    End If
        'End If

        ' Create Web Invoices
        Try
            ASCMAIN1.Progress("Creating Web Invoice", "")
            For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                TAC.SOCMAIN1.CreateWebInvoice(Me, row.Item("INV_TYPE"), row.Item("INV_NO"))
            Next
        Catch ex As Exception

        End Try

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "RA_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("CUST_CLAIM_NO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a Claim No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTRMAF1.RA_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTRMAF1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("CUST_CLAIM_NO").Text <> "" Then
                    sql_where &= " and SOTRMAF1.CUST_CLAIM_NO = '" & Absx1.txtFor("CUST_CLAIM_NO").Text & "'"
                End If

            Case "CUST_STORE_NO"
                sql_where &= " and CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"

        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("RA_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTRMAF1"
            E.COLUMN_NAME = "RA_NO"
            E.CODE_VALUE = Absx1.txtFor("RA_NO").Text
            E.DESC_VALUE = "Returns Authorization"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRMAF1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTRMAFX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTRMAF2, "B", "Item Status Inquiry")
        Load_Popup_Menu(grdSOTRMAFL, "B", "Track Shipment")
        Load_Popup_Menu(grdSOTINVHX, "SS", "Show Filter", "Show GroupBox")
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

        Select Case grd.Name
            'Case "grdSOTRMAF2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Track Shipment"
                Dim CARRIER_CODE As String = grd.ActiveRow.Cells("CARRIER_CODE").Value & String.Empty
                Dim TRACKING_NO As String = grd.ActiveRow.Cells("TRACKING_NO").Value & String.Empty

                If TRACKING_NO.Length = 0 Then
                    MessageBox.Show("There is no Tracking Number for this Call Tag.", "Track", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim rowSOTCARR1 As DataRow = LookUp("SOTCARR1", CARRIER_CODE)
                If rowSOTCARR1 Is Nothing Then
                    MessageBox.Show("Missing or Invalid Carrier.", "Track", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim CARRIER_URL_TRACKING As String = rowSOTCARR1.Item("CARRIER_URL_TRACKING") & String.Empty
                If CARRIER_URL_TRACKING.Length = 0 Then
                    MessageBox.Show("The Carrier is not setup with a Tracing URL.", "Track", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Try
                    Process.Start(CARRIER_URL_TRACKING & TRACKING_NO)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Track", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTRMAFX()
                End If

            Case "ORDR_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode _
                       And Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("CUST_CLAIM_NO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "RA_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If

            Case "RA_REASON_CODE"
                If Absx1.optFor("RA_REASON_CODE").Value = "X" And EntryMode = "N" Then
                    MsgBox("Credit is Issued Immediately when this Reason Code is Used",
                           MsgBoxStyle.OkOnly, "Please Note")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTRMAFX()

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTRMAFX()
            Case "RA_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

#Region "Controls"

    Private Sub grdSOTRMAFX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRMAFX.DoubleClickRow

        ' Prevebt DBNUll error
        If e.Row Is Nothing OrElse e.Row.IsFilterRow Then
            Exit Sub
        End If

        Absx1.txtFor("RA_NO").Text = e.Row.Cells("RA_NO").Value & String.Empty
        Click_Command("View")
    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTRMAFX()
    End Sub

    Private Sub optRA_REASON_CODE_ValueChanged(sender As Object, e As EventArgs) Handles optRA_REASON_CODE.ValueChanged

        If ScreenMode Then
            Dim RA_REASON_CODE As String = Absx1.optFor("RA_REASON_CODE").Value & String.Empty
            Dim rowARTREASR As DataRow = LookUp("ARTREASR", RA_REASON_CODE)
            If rowARTREASR IsNot Nothing Then
                'Absx1.txtFor("REASON_CODE").Text = rowARTREASR.Item("REASON_CODE") & String.Empty
            End If
        End If
    End Sub

#End Region

#Region "grdSOTRMAF2"

    Private Sub grdSOTRMAF2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRMAF2.AfterCellUpdate
        With grdSOTRMAF2.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value)
                    If STYLE_CODE <> "" Then
                        .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        .Cells("RA_RETAIL").Value = rowICTSTYL1.Item("STYLE_RETAIL")
                        .Cells("STYLE_COST").Value = Val(rowICTSTYL1.Item("STYLE_COST") & String.Empty)

                        ' e.Cell.Row.Cells("STYLE_UOM").Value = rowICTSTYL1.Item("STYLE_UOM") & ""

                        If COLOR_CODEs.Count = 1 Then
                            e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                        End If


                        Dim RA_DATE As Date

                        ' get lowest price from invoice sales
                        If dst.Tables("SOTINVH2X").Select("STYLE_CODE = '" & STYLE_CODE & "'").Length > 0 Then
                            Dim RA_NET_PRICE As Decimal = Val(dst.Tables("SOTINVH2X").Compute("MIN(ORDR_UNIT_PRICE)", "STYLE_CODE = '" & STYLE_CODE & "'") & String.Empty)
                            .Cells("RA_NET_PRICE").Value = RA_NET_PRICE
                        ElseIf rowSOTRMAF1.Item("RA_DATE") & "" = "" Then
                            MsgBox("Please specify an RA Date before entering Items", MsgBoxStyle.OkOnly, "Cannot Determine Price")
                        Else
                            RA_DATE = rowSOTRMAF1.Item("RA_DATE")

                            Dim ORDR_UNIT_PRICE As Decimal = 0
                            'TAC.SOCMAIN1.Get_Price _
                            '                                 (Me, _
                            '                                  PRICE_LIST_CODE, _
                            '                                  PRICE_BASIS, _
                            '                                  PRICE_BASE_DPCT, _
                            '                                  STYLE_CODE, _
                            '                                  rowICTSTYL1, _
                            '                                  rowSOTRMAF1.Item("RA_DATE")) ' MAYBE SHOULD USE RA_DATE - 60

                            .Cells("RA_NET_PRICE").Value = ORDR_UNIT_PRICE
                        End If

                        If ScreenMode And Not IsLoading Then
                            Load_SOTINVHX(STYLE_CODE)
                            If e.Cell.Row.IsAddRow Then
                                If dst.Tables("SOTINVHX").Rows.Count > 0 Then
                                    If ASCMAIN1.CLIENT = "RGI" Then
                                        If Val(.Cells("RA_NET_PRICE").Value & String.Empty) = 0 Then
                                            Dim RA_NET_PRICE As Decimal = Val(dst.Tables("SOTINVHX").Compute("MIN(ORDR_UNIT_PRICE)", "STYLE_CODE = '" & STYLE_CODE & "'") & String.Empty)
                                            .Cells("RA_NET_PRICE").Value = RA_NET_PRICE
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & "" ' grdSOTORDR2.ActiveRow.Cells("COLOR_CODE").Value & ""
                    If COLOR_CODE <> "" Then
                        If COLOR_CODEs.Contains(COLOR_CODE) Then
                            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                            '  e.Cell.Row.Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

                Case "RA_QTY"
                    If ASCMAIN1.CLIENT = "RGI" Then
                        Dim RA_QTY_AVAIL As Int32 = Val(e.Cell.Row.Cells("RA_QTY_AVAIL").Text & String.Empty)

                        If RA_QTY_AVAIL > 0 AndAlso .Cells("RA_QTY").Value > RA_QTY_AVAIL Then
                            MessageBox.Show("The Quantity provided (" & .Cells("RA_QTY").Value & ") is greater than the amount available (" & RA_QTY_AVAIL & ")", "Quatity", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            .Cells("RA_QTY").Value = RA_QTY_AVAIL
                        End If
                    End If

                    .Cells("RA_QTY_OPEN").Value = .Cells("RA_QTY").Value

                Case "RA_QTY_OPEN"
                    .Cells("RA_QTY_CANC").Value _
                        = Val(.Cells("RA_QTY").Value & "") _
                        - Val(.Cells("RA_QTY_USED").Value & "") _
                        - Val(.Cells("RA_QTY_OPEN").Value & "")
                    If Val(.Cells("RA_QTY_CANC").Value) < 0 Then
                        .Cells("RA_QTY_CANC").Value = 0
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSOTRMAF2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTRMAF2.AfterRowActivate

        If Trim(grdSOTRMAF2.ActiveRow.Cells("STYLE_CODE").Value & "") = "" And
            (grdSOTRMAF2.ActiveCell Is Nothing OrElse
             (grdSOTRMAF2.ActiveCell.Column.Key <> "STYLE_CODE")) _
        Then
            grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("STYLE_CODE")
            Exit Sub
        End If

        With grdSOTRMAF2.DisplayLayout.Bands(0)
            If grdSOTRMAF2.ActiveRow.IsAddRow Then
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                If grdSOTRMAF2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                    grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("STYLE_CODE")
                End If
            Else
                Validate_Style(grdSOTRMAF2.ActiveRow.Cells("STYLE_CODE").Value & "")
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                If Val(grdSOTRMAF2.ActiveRow.Cells("RA_QTY_USED").Value & "") <> 0 _
                Or Val(grdSOTRMAF2.ActiveRow.Cells("RA_QTY_CANC").Value & "") <> 0 _
                Then
                    .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End If
        End With

        If Not grdSOTRMAF2.ActiveRow.IsAddRow Then
            Load_SOTINVHX(grdSOTRMAF2.ActiveRow.Cells("STYLE_CODE").Value)
        End If

    End Sub

    Private Sub grdSOTRMAF2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTRMAF2.AfterRowsDeleted
        Display_Totals()

        If grdSOTRMAF2.Rows.Count = 0 Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = ""
        End If
    End Sub

    Private Sub grdSOTRMAF2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTRMAF2.AfterRowUpdate
        Display_Totals()

        If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTSTYL1.Item("SALES_DIVISION_CODE")
        End If
    End Sub

    Private Sub grdSOTRMAF2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTRMAF2.BeforeCellUpdate

        If grdSOTRMAF2.ActiveCell IsNot Nothing Then
            With grdSOTRMAF2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Text & String.Empty
                        Dim COLOR_CODE As String = e.Cell.Row.Cells("COLOR_CODE").Text & String.Empty

                        If STYLE_CODE.Length = 0 OrElse COLOR_CODE.Length = 0 Then
                            Exit Select
                        End If

                        STYLE_CODE = Validate_Style(STYLE_CODE)
                        If STYLE_CODE <> "" Then
                            If .Row.IsAddRow Then
                                If dst.Tables("SOTRMAF2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length <> 0 Then
                                    MsgBox("Style (" & STYLE_CODE & "), Color (" & COLOR_CODE & ") is already part of this RA Entry", MsgBoxStyle.OkOnly, "Cannot Add This Item")
                                    e.Cancel = True
                                    Exit Sub
                                End If
                            End If
                        Else
                            e.Cancel = True
                        End If

                End Select
            End With
        End If

    End Sub

    Private Sub grdSOTRMAF2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTRMAF2.BeforeExitEditMode
        If grdSOTRMAF2.ActiveCell IsNot Nothing Then
            With grdSOTRMAF2.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                    Case "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTRMAF2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTRMAF2.BeforeRowsDeleted

        RA_LNOs.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.IsAddRow Then
                e.Cancel = True
                grow.CancelUpdate()
                Exit Sub
            End If
            If Val(grow.Cells("RA_QTY_USED").Value & "") <> 0 _
            Or Val(grow.Cells("RA_QTY_CANC").Value & "") <> 0 _
            Then
                MsgBox("Cannot Delete a Line if it has ever been " & vbCr & "Used Or Cancelled" & vbCr & "Use the Cancel Button (x)")
                e.Cancel = True
                Exit Sub
            End If

            RA_LNOs.Add(grow.Cells("RA_LNO").Value)
        Next
    End Sub

    Private Sub grdSOTRMAF2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTRMAF2.BeforeRowUpdate

        Validate_Columns("STYLE_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("RA_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {e.Row.Cells("STYLE_CODE").Value & String.Empty, e.Row.Cells("COLOR_CODE").Value & String.Empty})
        If rowICTSTYC1 Is Nothing Then
            MessageBox.Show("The provided Style / Color combination is invalid", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("RA_NO").Value = RA_NO
            Dim RA_LNO As Int64 = Val(dst.Tables("SOTRMAF2").Compute("MAX(RA_LNO)", "") & "") + 1
            e.Row.Cells("RA_LNO").Value = RA_LNO
        End If
    End Sub

    Private Sub grdSOTRMAF2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRMAF2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "X"
                    If Val(.Cells("RA_QTY_CANC").Value) <> 0 Then
                        If MsgBox("Restore Cancelled Qty of " & .Cells("RA_QTY_CANC").Value,
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        .Cells("RA_QTY_OPEN").Value = Val(.Cells("RA_QTY_OPEN").Value & "") + Val(.Cells("RA_QTY_CANC").Value & "")

                        .Update()
                    Else
                        If MsgBox("Cancel Remaining Qty Open of " & .Cells("RA_QTY_OPEN").Value,
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        .Cells("RA_QTY_OPEN").Value = "0"
                        ' grdSOWRMAF2_AfterColUpdate(.Cells("RA_QTY_OPEN").position)
                        grdSOTRMAF2.ActiveRow.Update()
                    End If

                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTRMAF2, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                    grdClickCellButton(grdSOTRMAF2, sql_where)
            End Select
        End With

    End Sub

#End Region

#Region "Procedures"

    Private Function ImportFromExcel() As Boolean

        Try
            loadedFromExcel = False
            dst.Tables("ICTSTYC1").Rows.Clear()
            dst.Tables("ECTESTY1").Rows.Clear()

            Dim fileName As String = String.Empty

            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                Dim filter As String = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                openFileDialog1.Filter = filter
                openFileDialog1.RestoreDirectory = True
                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    fileName = openFileDialog1.FileName
                End If
            End Using

            If fileName.Length = 0 Then
                Return False
            End If

            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(fileName)
            Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
            Dim workSheetFound As Boolean = False

            For Each oSheet In oWB.Worksheets
                If oSheet.Name.ToUpper = "Details".ToUpper Then
                    workSheetFound = True
                    Exit For
                End If
            Next

            If Not workSheetFound Then
                MessageBox.Show("Could not locate a worksheet named Details.", "Import From Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Importing returned items", String.Empty)
            Dim colSku As Int16 = -1
            Dim colQty As Int16 = -1
            Dim colPrice As Int16 = -1

            Dim colNumber As Int16 = 0
            Dim lineNumber As Int16 = 0

            ' Vendor SKU, Qty and Cost
            While oSheet.Cells(lineNumber, colNumber).Value & String.Empty <> String.Empty
                Select Case (oSheet.Cells(lineNumber, colNumber).Value & String.Empty).ToString.ToUpper
                    Case "Vendor SKU".ToUpper
                        colSku = colNumber
                    Case "Qty".ToUpper
                        colQty = colNumber
                    Case "Cost".ToUpper
                        colPrice = colNumber
                End Select

                If colSku >= 0 AndAlso colQty >= 0 AndAlso colPrice >= 0 Then
                    Exit While
                End If

                colNumber += 1

            End While

            If Not (colSku >= 0 AndAlso colQty >= 0 AndAlso colPrice >= 0) Then
                MessageBox.Show("The Details Worksheet must have the following three column headers: Vendor SKU, Qty and Cost", "Import From Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            Dim STYLE_CODE As String = String.Empty
            Dim COLOR_CODE As String = String.Empty
            Dim RA_QTY As Int32 = 0
            Dim RA_NET_PRICE As Decimal = 0

            Dim ECOM_CODE As String = String.Empty
            If dst.Tables("ECTECOM1").Select($"CUST_CODE = '{CUST_CODE}'").Length > 0 Then
                ECOM_CODE = dst.Tables("ECTECOM1").Select($"CUST_CODE = '{CUST_CODE}'")(0).Item("ECOM_CODE") & String.Empty
                ASCMAIN1.sql = $"SELECT * FROM ECTESTY1 WHERE SET_QTY > 1 AND ECOM_CODE = '{ECOM_CODE}'"
                Fill_Records("ECTESTY1", String.Empty, True, ASCMAIN1.sql)
            End If

            Dim lstBadStyleColor As New List(Of String)

            lineNumber = 1
            Do While oSheet.Cells(lineNumber, colSku).Value & String.Empty <> String.Empty
                Dim customerSKU As String = oSheet.Cells(lineNumber, colSku).Value & String.Empty

                ASCMAIN1.Progress("-", customerSKU)

                While customerSKU.Contains(Space(2))
                    customerSKU = customerSKU.Replace(Space(2), Space(1))
                End While

                If customerSKU.Contains("_") Then
                    STYLE_CODE = customerSKU.Split("_")(0)
                    COLOR_CODE = customerSKU.Split("_")(1)
                ElseIf customerSKU.Contains("-") Then
                    STYLE_CODE = customerSKU.Split("-")(0)
                    COLOR_CODE = customerSKU.Split("-")(1)
                ElseIf customerSKU.Contains(" ") Then
                    STYLE_CODE = customerSKU.Split(" ")(0)
                    COLOR_CODE = customerSKU.Split(" ")(1)
                Else
                    MessageBox.Show("Vendor SKU column must split the Style Code, Color Code using either an underscor, hypen or space.", "Import From Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If

                STYLE_CODE = STYLE_CODE.ToUpper
                COLOR_CODE = COLOR_CODE.ToUpper

                Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New Object() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 Is Nothing Then
                    Fill_Records("ICTSTYC1", New Object() {STYLE_CODE, COLOR_CODE})
                End If
                rowICTSTYC1 = dst.Tables("ICTSTYC1").Rows.Find(New Object() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 Is Nothing Then
                    If Not lstBadStyleColor.Contains(customerSKU) Then
                        lstBadStyleColor.Add(customerSKU)
                    End If
                End If

                RA_QTY = Val(oSheet.Cells(lineNumber, colQty).Value & String.Empty)
                RA_NET_PRICE = Val((oSheet.Cells(lineNumber, colPrice).Value & String.Empty).ToString.Replace("$", "").Replace(" ", ""))

                Dim importedItem As New ItemsImported
                With importedItem
                    .STYLE_CODE = STYLE_CODE
                    .COLOR_CODE = COLOR_CODE
                    .RA_QTY = RA_QTY
                    .RA_NET_PRICE = RA_NET_PRICE
                    .SET_QTY = 1

                    Dim rowECTESTY1 As DataRow = dst.Tables("ECTESTY1").Rows.Find(New Object() {STYLE_CODE, ECOM_CODE})
                    If rowECTESTY1 IsNot Nothing Then
                        .SET_QTY = rowECTESTY1.Item("SET_QTY")
                        .RA_QTY *= .SET_QTY
                        .RA_NET_PRICE /= .SET_QTY
                    End If
                End With

                lstImportedFromExcel.Add(importedItem)

                lineNumber += 1
            Loop

            If lstBadStyleColor.Count > 0 Then
                MessageBox.Show($"There file contains invalid Style / Color combinations: {Environment.NewLine} {String.Join(Environment.NewLine, lstBadStyleColor.ToArray)}", "Import From Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            importedFromExcel = True
            loadedFromExcel = False
            EntryMode = "N"
            Load_Record()

            Return loadedFromExcel

        Catch ex As Exception
            MessageBox.Show($"Import From Excel Error {ex.Message}", "Import From Excel", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
            dst.Tables("ICTSTYC1").Rows.Clear()
            dst.Tables("ECTESTY1").Rows.Clear()
            Me.Cursor = Cursors.Default
        End Try

    End Function

    Sub Load_SOTRMAFX()
        Dim sqlw As String = ""
        If InquiryMode Then
            If optStatus.Value <> "A" Then
                sqlw = " and SOTRMAF1.RA_STATUS = '" & optStatus.Value & "'"
            End If
        Else
            sqlw = " and SOTRMAF1.RA_STATUS = 'O'"
        End If
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE = "" Then
            grdSOTRMAFX.Text = "Open Returns Authorizations"
        Else
            sqlw &= " and SOTRMAF1.CUST_CODE = '" & CUST_CODE & "'"
            grdSOTRMAFX.Text = "Open Returns Authorizations associated with " & CUST_CODE
        End If

        ASCMAIN1.sql = "Select SOTRMAF1.*, ARTCUST1.CUST_NAME from SOTRMAF1, ARTCUST1 WHERE SOTRMAF1.CUST_CODE = ARTCUST1.CUST_CODE (+) " & sqlw
        Fill_Records("SOTRMAFX", "", , ASCMAIN1.sql)
        Sort_grdColumns(grdSOTRMAFX, "RA_NO".ToLower)

        ASCMAIN1.sql = "Select SOTRMAF2.RA_NO" _
            & ", Sum (SOTRMAF2.RA_QTY) RA_QTY" _
            & ", Sum (SOTRMAF2.RA_QTY_OPEN) RA_QTY_OPEN" _
            & ", Sum (SOTRMAF2.RA_QTY_USED) RA_QTY_USED" _
            & ", Sum (SOTRMAF2.RA_QTY_CANC) RA_QTY_CANC" _
            & ", Sum (NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT" _
            & ", Sum (NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_OPEN" _
            & ", Sum (NVL(SOTRMAF2.RA_QTY_USED,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_USED" _
            & ", Sum (NVL(SOTRMAF2.RA_QTY_CANC,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_CANC" _
            & ", Sum (NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_RETAIL,0)) RA_RETAIL_EXT" _
            & " from SOTRMAF2,SOTRMAF1 where SOTRMAF2.RA_NO = SOTRMAF1.RA_NO" & sqlw & " group by SOTRMAF2.RA_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim RA_NO As String = row.Item("RA_NO")
            Dim rowSOTRMAFX As DataRow = dst.Tables("SOTRMAFX").Rows.Find(RA_NO)
            If rowSOTRMAFX IsNot Nothing Then
                For Each COLUMN_NAME As String In New String() _
                    {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC",
                     "RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC", "RA_RETAIL_EXT"}
                    rowSOTRMAFX.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
            End If
        Next

        grdSOTRMAFX.Visible = True
    End Sub

    Sub Print_Record()

        ' To use the data layer and dst that is associated with this form

        Fill_Records("ARTCUST1", CUST_CODE)
        Fill_Records("ARTCUST2", New String() {CUST_CODE, "MK", Absx1.txtFor("CUST_STORE_NO").Text})
        Fill_Records("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
        Fill_Records("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)

        rowSOTRMAF1.Item("AR_PARM_KEY") = "Z"
        rowSOTRMAF1.Item("RA_AMT") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT)", "") & "")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = "SORRMAP1"

        Select Case ASCMAIN1.CLIENT
            Case "NYA"
                RPT = "SORRMAPN"
            Case "RGI"
                RPT = "SORRMAPR"
        End Select

        Generate_Report(RPT, "Returns Authorization", , , , , False)
        Print_Report_End()

    End Sub

    Private Sub Print_Credit_Memo()
        Try
            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Now Preparing Credit Memo for Printing")

            Dim REPORT_NAME As String = "SORINVP1"
            Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
            If RPT = "" Then RPT = REPORT_NAME

            If Not REPORTS.ContainsKey(REPORT_NAME) Then
                REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
                REPORTS(REPORT_NAME).Prepare_dst(False, "")
            End If

            Dim sql As String = " and SOTINVH1.INV_TYPE = '" & rowSOTRMAF1.Item("INV_TYPE") & "' and SOTINVH1.INV_NO = '" & rowSOTRMAF1.Item("INV_NUM") & "'"
            Dim tempFileName As String = "Memo" & DateTime.Now.ToString("yyyyMMddHHmmss")

            REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
            Dim FILENAME As String = ""
            With REPORTS(REPORT_NAME).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "")
                Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
                FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
                .Print_Report_End(, True)
            End With

            Show_Document(FILENAME)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print Credit Memo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub Cancel_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        Dim EMsg As String = ""
        If EntryMode = "E" Then
            Cancel_Order_1(RA_NO)
            EMsg = "Balance Open on Returns Authorization " & RA_NO & " has been Cancelled"
        End If

        'ASCDATA1.ExecuteSP("SOPRSRV0_G", "V", New Object() {RSRV_GROUP_NO}, New String() {"RSRV_GROUP_NO_IN"})
        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Cancel_Order_1(RA_NO As String)
        Dependent_Updates(-1, RA_NO)

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is Select * from SOTRMAF2 where RA_NO = '" & RA_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTRMAF2" _
            & "    Set RA_QTY_CANC = NVL(RA_QTY_CANC,0) + NVL(R1.RA_QTY_OPEN,0)" _
            & "      , RA_QTY_OPEN = 0" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()
        ', RA_STATUS = 'C'
        ASCMAIN1.sql = "Update SOTRMAF1 Set RA_STATUS = :PARM1" _
            & " where RA_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"F", RA_NO})
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        Dim EMsg As String = ""

        BeginTrans()

        If EntryMode = "E" Then
            Delete_Order_1(RA_NO)
            EMsg = "Returns Authorization No " & RA_NO & " has been marked as Deleted"
        End If

        CommitTrans(EMsg)
        'ASCDATA1.ExecuteSP("SOPRSRV0_G", "V", New Object() {RSRV_GROUP_NO}, New String() {"RSRV_GROUP_NO_IN"})
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(RA_NO As String)
        Dependent_Updates(-1, RA_NO)

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select * from SOTRMAF2" & vbCrLf _
            & "     where RA_NO = '" & RA_NO & "' for Update;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SOTRMAF2" & vbCrLf _
            & "    Set RA_QTY_CANC = NVL(RA_QTY_CANC,0) + NVL(R1.RA_QTY_OPEN,0)" & vbCrLf _
            & "   , RA_QTY_OPEN = 0" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTRMAF1 Set RA_STATUS = :PARM1" _
            & " where RA_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"D", RA_NO})
    End Sub

    Sub Dependent_Updates(S As Integer, RA_NO As String)

    End Sub

    Sub Display_Totals()
        Dim KEY As Int32 = 0
        For Each SFX As String In New String() {"", "OPEN", "USED", "CANC"}
            If SFX <> "" Then SFX = "_" & SFX
            KEY += 1
            Dim rowSOTRMAFT As DataRow = dst.Tables("SOTRMAFT").Rows.Find(KEY)
            rowSOTRMAFT.Item("QTY") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_QTY" & SFX & ")", "") & "")
            rowSOTRMAFT.Item("AMT") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT" & SFX & ")", "") & "")
        Next
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTRMAF2.ActiveRow
            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = ""
                    If Trim(.Cells("STYLE_CODE").Value & "") <> "" Then
                        STYLE_CODE = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    End If
                    Cancel = (STYLE_CODE = "")

                Case "RA_QTY"
                    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If

                    If Val(.Cells("RA_QTY").Value & "") = 0 Then
                        MsgBox("Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("RA_QTY")
                        Exit Sub
                    End If

                    If Val(.Cells("RA_QTY").Value & "") < 0 Then
                        MsgBox("Qty May Not be Negative", vbOKOnly, "Invalid Quantity")
                        grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("RA_QTY")
                        Cancel = True
                    End If

                    If Val(.Cells("ORDR_QTY_SHIP").Value & String.Empty) > 0 Then
                        If Val(.Cells("RA_QTY").Value & "") > Val(.Cells("ORDR_QTY_SHIP").Value & "") Then
                            MsgBox("Qty May Not be greater Qty Sold (" & Val(.Cells("ORDR_QTY_SHIP").Value & "") & ")", vbOKOnly, "Invalid Quantity")
                            grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("RA_QTY")
                            Cancel = True
                        End If
                    End If

            End Select
        End With
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim EMsg As String = ""
        If STYLE_CODE_z = "" Then Return ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            EMsg = "Style is Not on File" & vbCrLf
        Else
            'If rowICTSTYL1.Item("ITEM_STATUS") & "" <> "A" Then
            '    EMsg = "Item Status is not Active" & vbCrLf
            'End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                EMsg = "Style does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                EMsg = "Style does not have a valid Division Code" & vbCrLf
            End If
        End If

        If EMsg = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If EMsg <> "" And grdSOTRMAF2.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

    Sub Load_Events()
        '    grdEvents.RemoveAll
        '    Call Load_Events_1("Entered", "INIT_DATE")
        '    Call Load_Events_1("Modified", "LAST_DATE")
        '    Call Load_Events_1("Released", "RSRV_DATE_REL")
    End Sub

    Sub Record_AR_Item()

        If ASCMAIN1.CLIENT = "RGI" Then
            Exit Sub
        End If

        Dim INV_NUM As String = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        Dim RA_AMT As Decimal = Val(rowSOTRMAF1.Item("RA_AMT") & "")
        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        With rowARTOPEN1
            .Item("CUST_CODE") = CUST_BILL_TO_CUST
            .Item("INV_TYPE") = "C"
            .Item("INV_NUM") = INV_NUM
            .Item("INV_DATE") = rowSOTRMAF1.Item("RA_DATE")
            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE")
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE")
            .Item("INV_DUE_DATE") = rowSOTRMAF1.Item("RA_DATE")
            .Item("INV_CUST_PO") = rowSOTRMAF1.Item("CUST_CLAIM_NO")
            .Item("INV_MISC_CHG") = -1 * RA_AMT
            .Item("INV_TOTAL_AMOUNT") = -1 * RA_AMT
            .Item("INV_BALANCE") = -1 * RA_AMT
            .Item("CUST_CODE_SO") = CUST_CODE

            .Item("REASON_CODE") = rowSOTRMAF1.Item("REASON_CODE")

            .Item("ORDR_NO") = RA_NO
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("SALES_DIVISION_CODE") = rowSOTRMAF1.Item("SALES_DIVISION_CODE")
            .Item("SREP_CODE") = rowSOTRMAF1.Item("SREP_CODE")

            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            .Item("INV_SALES_CURR") = .Item("INV_SALES")
            .Item("INV_DISC_CURR") = .Item("INV_DISC")
            .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
            .Item("INV_STAX_CURR") = .Item("INV_STAX")
            .Item("INV_MISC_CHG_CURR") = .Item("INV_MISC_CHG")
            .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT")
            .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")

            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1

            .Item("INV_NOTES") = rowSOTRMAF1.Item("RA_NOTES")
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP

            For Each field As String In New String() {"INV_SALES_CURR", "INV_DISC_CURR", "INV_FREIGHT_CURR", "INV_STAX_CURR", "INV_MISC_CHG_CURR", "INV_TOTAL_AMOUNT_CURR", "INV_PMT_CURR", "INV_DISC_TAKEN_CURR", "INV_WRITE_OFF_CURR", "INV_BALANCE_CURR", "GST_TAX_CURR"}
                field = field.Trim
                .Item(field) = .Item(field.Replace("_CURR", ""))
            Next

        End With
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
        Update_Record_TDA("ARTOPEN1")

        Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
        With rowSOTINVH1
            .Item("INV_TYPE") = "C"
            .Item("INV_NO") = INV_NUM
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_STORE_NO") = rowSOTRMAF1.Item("CUST_STORE_NO")
            If .Item("CUST_STORE_NO") & "" = "" Then
                .Item("CUST_STORE_NO") = "000000"
            End If
            .Item("ORDR_CUST_PO") = rowSOTRMAF1.Item("CUST_CLAIM_NO")
            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE")
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE")
            .Item("SREP_CODE") = rowSOTRMAF1.Item("SREP_CODE")
            .Item("REASON_CODE") = String.Empty 'rowSOTRMAF1.Item("REASON_CODE")
            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            .Item("INV_MISC_CHG") = -1 * RA_AMT
            .Item("INV_TOTAL_AMOUNT") = -1 * RA_AMT
            .Item("INV_DATE") = rowSOTRMAF1.Item("RA_DATE")
            .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1
            .Item("INV_COMMENT") = rowSOTRMAF1.Item("RA_NOTES")
            .Item("ORDR_TYPE_CODE") = "DIF" ' THIS SHOULD BE PARAMETERIZED
            .Item("REGISTER_IND") = "0"
            .Item("ORDR_DEPT") = rowSOTRMAF1.Item("ORDR_DEPT")
            .Item("SREP2_CODE") = rowSOTRMAF1.Item("SREP2_CODE")
            .Item("SALES_DIVISION_CODE") = rowSOTRMAF1.Item("SALES_DIVISION_CODE")

            For Each field As String In New String() {"INV_TOTAL_AMOUNT_CURR", "INV_SALES_CURR", "INV_FREIGHT_CURR", "INV_MISC_CHG_CURR", "GST_TAX_CURR", "INV_STAX_CURR"}
                field = field.Trim
                .Item(field) = .Item(field.Replace("_CURR", ""))
            Next

            .Item("INV_TOTAL_AMT_CURR") = .Item("INV_TOTAL_AMOUNT_CURR")

        End With
        dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)
        Update_Record_TDA("SOTINVH1")

        Dim rowSOTINVHM As DataRow = dst.Tables("SOTINVHM").NewRow
        With rowSOTINVHM
            .Item("INV_TYPE") = "C"
            .Item("INV_NO") = INV_NUM
            .Item("INV_MNO") = 1
            .Item("MISC_CHG_CODE") = "DF"
            .Item("INV_MISC_CHG") = -1 * RA_AMT
        End With
        dst.Tables("SOTINVHM").Rows.Add(rowSOTINVHM)
        Update_Record_TDA("SOTINVHM")

        ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV",
                           New Object() {"C", INV_NUM},
                           New String() {"INV_TYPE_IN", "INV_NO_IN"})

        rowSOTRMAF1.Item("INV_TYPE") = "C"
        rowSOTRMAF1.Item("INV_NUM") = INV_NUM
        rowSOTRMAF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP

    End Sub

    Sub Record_AR_Item_Reversal(row As DataRow)
        Dim INV_NUM As String = ASCMAIN1.Next_Control_No("AR_CR_MEMO_NO")
        Dim INV_TOTAL_AMOUNT As Decimal = Val(row.Item("INV_TOTAL_AMOUNT") & "")
        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        rowARTOPEN1.ItemArray = row.ItemArray
        With rowARTOPEN1
            .Item("INV_NUM").Value = INV_NUM
            .Item("INV_MISC_CHG").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_TOTAL_AMOUNT").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_BALANCE").Value = -1 * INV_TOTAL_AMOUNT

            .Item("INV_MISC_CHG_CURR").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_TOTAL_AMOUNT_CURR").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_BALANCE_CURR").Value = -1 * INV_TOTAL_AMOUNT

            .Item("INIT_OPER").Value = ASCMAIN1.USER_ID
            .Item("INIT_DATE").Value = DATETIME_STAMP
        End With
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
        Update_Record_TDA("ARTOPEN1")
    End Sub

    Sub Load_SOTINVHX(STYLE_CODE As String)
        Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Fill_Records("SOTINVHX", New String() {CUST_CODE, YP, STYLE_CODE})
        Sort_grdColumns(grdSOTINVHX, "INV_DATE".ToLower)
        grdSOTINVHX.Text = "Recent Sales of Style " & STYLE_CODE & " To " & Absx1.txtFor("CUST_CODE").Text
    End Sub

#End Region

End Class