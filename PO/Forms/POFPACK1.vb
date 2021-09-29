Imports Infragistics.Win.UltraWinGrid

Public Class POFPACK1
    ' YT HARDCODED THINGS: BARCODE_PFX = "Y" LOAD RECORD

    Dim rowPOTPACK1 As DataRow
    Dim PACK_LIST_NO As String
    Dim PACK_LIST_NO_new As String
    Dim PACK_LIST_STATUS As String

    Dim rowTATUSER1 As DataRow

    Dim sqlPOTPACKX As String
    Dim VEND_CODE As String = ""
    Dim VEND_CODE_USER As String = ""

    Dim PO_REFERENCE As String = ""
    Dim STYLE_CODE_PFX As String = ""
    Dim INITIAL_ORDER As String = ""
    Dim PO_ORDER_NO As String = ""
    Dim CUST_CODE As String = ""
    Dim PO_SPEC_ORDR_NO As String = ""
    Dim Appearance_Red As New Infragistics.Win.Appearance
    Dim unFinalize As Boolean = False

    Dim rowPOTPACKC As DataRow

    Dim PO_REFERENCE2 As String = ""
    Dim STYLE_CODE_PFX2 As String = ""
    Dim PO_ORDER_NO2 As String = ""

    Dim BARCODE_PFX As String = ""


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Appearance_Red.ForeColor = Drawing.Color.Red

        If MENU_ITEM_OBJECT = "POTLTRCI" Then
            InquiryMode = True
        End If

        rowTATUSER1 = LookUp("TATUSER1", ASCMAIN1.USER_ID)
        If rowTATUSER1 IsNot Nothing AndAlso rowTATUSER1.Item("VEND_CODE") & "" <> "" Then
            VEND_CODE_USER = rowTATUSER1.Item("VEND_CODE")
        Else
            VEND_CODE_USER = ""
        End If

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            VEND_CODE_USER = "YINTAK"
        End If

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
        End With

        Get_PARM("GLTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")

        With dst
            sqlPOTPACKX = "Select POTPACK1.*,APTVEND1.VEND_NAME" & vbCrLf _
                & ", X.CARTONS, X.CARTONS_TO_SHIP, X.CARTONS_TO_REMOVE" & vbCrLf _
                & " from POTPACK1,APTVEND1" & vbCrLf _
                & ", ( Select PACK_LIST_NO, Count (*) CARTONS, SUM (CASE WHEN SHIP_CONF='S' THEN 1 ELSE 0 END) CARTONS_TO_SHIP, SUM (CASE WHEN SHIP_CONF='S' THEN 1 ELSE 0 END) CARTONS_TO_REMOVE from POTLPNL1 where BARCODE_STATUS = 'A' group by PACK_LIST_NO ) X" & vbCrLf _
                & " where APTVEND1.VEND_CODE = POTPACK1.VEND_CODE" & vbCrLf _
                & "   And X.PACK_LIST_NO (+) = POTPACK1.PACK_LIST_NO" & vbCrLf _
                & "   And POTPACK1.PACK_LIST_STATUS <> 'D'"
            ASCMAIN1.sql = sqlPOTPACKX ' & "  and POTPACK1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "POTPACKX", "**", 0, False, "")

            Create_TDA(.Tables.Add, "POTPACK1", "*")

            Create_TDA(.Tables.Add, "POTPACK2", "*", 1)

            ASCMAIN1.sql = "Select POTPACK3.*, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & " from POTPACK3, ICTSTYL1 where ICTSTYL1.STYLE_CODE = POTPACK3.STYLE_CODE and POTPACK3.PACK_LIST_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTPACK3", "**", 0, True, "V")
            With .Tables("POTPACK3")
                .Columns.Add("TOTAL_UNITS", GetType(System.Int32), "CARTON_COUNT * CARTON_PACK")
                .Columns.Add("CARTON_NO_START", GetType(System.Int32))
                .Columns.Add("CARTON_NO_END", GetType(System.Int32), "CARTON_NO_START + CARTON_COUNT -1")
                .Columns.Add("TOTAL_GRS_WGT", GetType(System.Decimal), "CARTON_COUNT * CARTON_GRS_WGT")
                .Columns.Add("TOTAL_NET_WGT", GetType(System.Decimal), "CARTON_COUNT * CARTON_NET_WGT")
                ' .Columns.Add("STYLE_WEIGHT", GetType(System.Decimal), "IIF(ISNULL(CARTON_COUNT,0) = 0, 0, ISNULL(CARTON_NET_WGT,0) / ISNULL(CARTON_COUNT,0))")

            End With

            Create_Relation("POTPACK2", "POTPACK3", "PACK_LIST_NO,PACK_LIST_SHEET_NO")

            With .Tables("POTPACK2")
                .Columns.Add("COLOR_DESC")
                .Columns.Add("TOTAL_CARTONS", GetType(System.Int32), "SUM(CHILD.CARTON_COUNT)")
                .Columns.Add("TOTAL_UNITS", GetType(System.Int32), "SUM(CHILD.TOTAL_UNITS)")
                .Columns.Add("TOTAL_GRS_WGT", GetType(System.Decimal), "SUM(CHILD.TOTAL_GRS_WGT)")
                .Columns.Add("TOTAL_NET_WGT", GetType(System.Decimal), "SUM(CHILD.TOTAL_NET_WGT)")
                .Columns.Add("CARTON_PACK_HOLD", GetType(System.Int32), "SUM(CHILD.CARTON_PACK)")
            End With

            'With .Tables("POTPACK3")
            '    .Columns("CARTON_NO_START").Expression = "PARENT.CARTON_NO_START"
            'End With

            ' PUT THIS TABLE INTO ORACLE AFTER IT PASSES MUSTER WITH WALMART & KOHLS, AND FILL_RECORDS INSTEAD OF .ROWS.ADD
            With .Tables.Add("POTPACKC")
                .Columns.Add("CUST_CODE")
                .Columns.Add("PACK_INITIAL_BY_COLOR")
                .Columns.Add("PACK_INITIAL_ODD_CARTONS")
                .Columns.Add("PACK_INITIAL_MULTI_PO")
                ' .Columns.Add("PACK_INITIAL_SIMPLE_RATIO") NO SIMPLE RATIO IF WE ALLOW ODD CARTONS
                .PrimaryKey = New DataColumn() { .Columns("CUST_CODE")}
                .Rows.Add(New String() {"WALMART", "0", "0", "1"})
                .Rows.Add(New String() {"KOHLS", "1", "1", "0"})
                .Rows.Add(New String() {"MEIJER", "1", "1", "0"})
                .Rows.Add(New String() {"COSTCO", "0", "0", "1"})

                ' CUSTOMER MASTER: PACK INITAL PO BY COLOR, ODD CARTONS, SIMPLE RATIO
                ' MEIJERS WILL LOOK Like KOHLS
                ' COSTCO WILL LOOK Like WALMART
            End With


            ASCMAIN1.sql = "Select * from POTORDR1 where PO_REFERENCE = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "POTORDR2", "*", 1, False)

            ASCMAIN1.sql = "Select PO_ORDER_NO, PO_REFERENCE, PO_SPEC_ORDR_NO, PO_DATE_SHIP_BY, PO_DATE_ETA, CUST_CODE, STYLE_CODE_PFX, CARTON_COUNT" & vbCrLf _
                & " from POTORDR1 where VEND_CODE = :PARM1 And PO_STATUS = 'O'"
            Create_TDA(.Tables.Add, "POTORDRR", "**", 0, False, "V")

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, SUM (PO_QTY_OPN) PO_QTY_OPN, MIN (PO_ORDER_LNO) PO_ORDER_LNO" & vbCrLf _
                & " from POTORDR2" & vbCrLf _
                & " where PO_ORDER_NO = :PARM1" & vbCrLf _
                & "   and PO_QTY_OPN <> 0" & vbCrLf _
                & " group by STYLE_CODE, COLOR_CODE"
            ASCMAIN1.sql = "Select X.*, POTORDR5.CARTON_PACK from POTORDR5, (" & ASCMAIN1.sql & ") X" & vbCrLf _
                & " where POTORDR5.PO_ORDER_NO (+) = :PARM1" & vbCrLf _
                & "   and POTORDR5.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   and POTORDR5.COLOR_CODE (+) = X.COLOR_CODE"
            Create_TDA(.Tables.Add, "POTORDRD", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "WHTSCSEQ", "*", 0, False)
            Fill_Records("WHTSCSEQ")

            'ASCMAIN1.sql = "Select * from ICTSTYC1 where CARTON_ID Is Not Null"
            'Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False)
            'Fill_Records("ICTSTYC1")

            ASCMAIN1.sql = "Select POTLPNL1.* from POTLPNL1 where PACK_LIST_NO = :PARM1 And BARCODE_STATUS = 'A'"
            Create_TDA(.Tables.Add, "POTLPNL1", "**", 0, True, "V")
            With .Tables("POTLPNL1")
                .Columns.Add("CONF_SHP", GetType(System.String), "IIF(SHIP_CONF='S','1','0')")
                .Columns.Add("CONF_REM", GetType(System.String), "IIF(SHIP_CONF='R','1','0')")
                .Columns.Add("CONF_UNK", GetType(System.String), "IIF(ISNULL(SHIP_CONF,'?')='?','1','0')")
            End With

            ASCMAIN1.sql = "Select WHTPKGM1.*, " & "TO_CHAR(PKG_L) || CHR(34) || 'x' || TO_CHAR(PKG_W) || CHR(34) || 'x' || TO_CHAR(PKG_H) || CHR(34)" & " CARTON_DIMENSIONS from WHTPKGM1 where BARCODE_PFX = :PARM1 And PKG_STATUS = 'A'"
            Create_TDA(.Tables.Add, "WHTPKGM1", "**", 0, True, "V")

        End With

        grdPOTPACKX.DataSource = dst.Tables("POTPACKX")

        grdPOTPACK2.DataSource = dst.Tables("POTPACK2")
        grdPOTPACK3.DataSource = dst.Tables("POTPACK3")
        grdPOTLPNLX.DataSource = dst.Tables("POTLPNL1")
        grdPOTLPNL1.DataSource = dst.Tables("POTLPNL1")

        grdPOTORDRR.DataSource = dst.Tables("POTORDRR")
        grdPOTORDRD.DataSource = dst.Tables("POTORDRD")

        grdWHTPKGM1.DataSource = dst.Tables("WHTPKGM1")

        Create_Summary(grdPOTPACKX, "PACK_LIST_NO", "Count")
        ' Create_Summary(grdPOTPACKX, New String() {"LC_AMT", "LC_PMTS", "LC_FEES", "LC_OPEN"})

        Create_Summary(grdPOTPACK2, "PACK_LIST_SHEET_NO", "Count")
        Create_Summary(grdPOTPACK2, New String() {"TOTAL_CARTONS", "TOTAL_UNITS", "TOTAL_GRS_WGT", "TOTAL_NET_WGT", "CARTON_COUNT"})

        Create_Summary(grdPOTPACK3, "PACK_LIST_SHEET_LNO", "Count")
        Create_Summary(grdPOTPACK3, New String() {"CARTON_COUNT", "TOTAL_UNITS", "TOTAL_GRS_WGT", "TOTAL_NET_WGT"})

        Create_Summary(grdPOTLPNL1, "BARCODE", "Count")
        Create_Summary(grdPOTLPNL1, New String() {"CONF_SHP", "CONF_REM", "CONF_UNK"})

        With grdPOTPACKX.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"PACK_LIST_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                ElseIf New String() {"CARTONS", "CARTONS_TO_SHIP", "CARTONS_TO_REMOVE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("PACK_LIST_NO").Header.Fixed = True
        End With

        With grdPOTPACK2.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                GCOL.CellActivation = Activation.NoEdit
                'XCARTON_DIMENSIONS
                If New String() {"PACK_LIST_DETAILS", "CARTON_NO_START", "CARTON_COUNT", "CARTON_GRS_WGT", "CARTON_NET_WGT", "PKG_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.DodgerBlue '.LightGreen
                    'GCOL.CellAppearance.BackColor = System.Drawing.Color.LightGreen
                    GCOL.CellActivation = Activation.AllowEdit
                ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                ElseIf New String() {"CARTON_PACK_HOLD"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next

        End With

        With grdPOTPACK3.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                GCOL.CellActivation = Activation.NoEdit
                'XCARTON_DIMENSIONS
                If New String() {"CARTON_COUNT", "CARTON_PACK", "CARTON_GRS_WGT", "CARTON_NET_WGT", "PKG_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.DodgerBlue '.LightGreen
                    ' GCOL.CellAppearance.BackColor = System.Drawing.Color.LightGreen
                    GCOL.CellActivation = Activation.AllowEdit
                ElseIf New String() {"TOTAL_GRS_WGT", "TOTAL_NET_WGT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next

        End With


        With grdPOTLPNL1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"PACK_LIST_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"CONF_SHP", "CONF_REM", "CONF_UNK"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                End If
            Next
            .Columns("PACK_LIST_NO").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdPOTPACK2, "PKG_CODE", "Select PKG_CODE, PKG_DESC from WHTPKGM1 where BARCODE_PFX = 'Y' order by PKG_DESC")
        ASCMAIN1.Add_Value_List(grdPOTPACK3, "PKG_CODE", "Select PKG_CODE, PKG_DESC from WHTPKGM1 where BARCODE_PFX = 'Y' order by PKG_DESC")
        ASCMAIN1.Add_Value_List(grdPOTPACKX, "PACK_LIST_STATUS", Nothing, New String() {":", "O:Open", "F:Finalized"})

        ASCMAIN1.Add_Value_List(grdPOTLPNL1, "SHIP_CONF", Nothing, New String() {":", "S:To Ship", "R:To Remove"})

        grpHeader.Visible = False

        '  Absx1.txtFor("CURR_CODE").ReadOnly = True

        Show_Filter(grdPOTPACKX, True)
        Show_Filter(grdPOTORDRR, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                unFinalize = False

                VEND_CODE = ""
                If Absx1.txtFor("VEND_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Supplier Code"
                Else
                    Dim row As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Supplier Code Entered Is Not Valid"
                    Else
                        If row.Item("VEND_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Supplier Status Is Not Active"
                        Else
                            VEND_CODE = Absx1.txtFor("VEND_CODE").Text
                        End If
                    End If
                End If

                If VEND_CODE <> VEND_CODE_USER Then
                    EMsg &= vbCr & "Invalid Vendor (not matching Vendor in User Profile)"
                End If
                'Dim DT As Date = Absx1.dteFor("PACK_INV_DATE").Value
                'If DT & "" = "" Then
                '    EMsg &= vbCr & "Invoice Date is Mandatory"
                'Else
                '    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                'End If

                PO_ORDER_NO = ""
                PO_SPEC_ORDR_NO = ""
                CUST_CODE = ""
                PO_ORDER_NO2 = ""
                PO_REFERENCE2 = ""
                INITIAL_ORDER = "0"
                STYLE_CODE_PFX = ""
                STYLE_CODE_PFX2 = ""

                If Absx1.txtFor("PO_REFERENCE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid PO Reference"
                Else
                    PO_REFERENCE = Absx1.txtFor("PO_REFERENCE").Text
                    Fill_Records("POTORDR1", PO_REFERENCE)
                    If dst.Tables("POTORDR1").Rows.Count > 1 Then
                        EMsg &= vbCr & $"More than 1 Vandale PO is associated with PO Reference {PO_REFERENCE}"
                    ElseIf dst.Tables("POTORDR1").Rows.Count = 0 Then
                        EMsg &= vbCr & $"No record PO Reference {PO_REFERENCE}"
                    Else
                        Dim row As DataRow = dst.Tables("POTORDR1").Rows(0)
                        If row.Item("VEND_CODE") & "" <> VEND_CODE Then
                            EMsg &= vbCr & $"Invalid PO Reference {PO_REFERENCE}"
                        ElseIf row.Item("PO_STATUS") & "" <> "O" Then
                            EMsg &= vbCr & $"PO Reference {PO_REFERENCE} is not Open"
                        ElseIf row.Item("CUST_CODE") & "" = "" Then
                            EMsg &= vbCr & $"PO Reference {PO_REFERENCE} is not Associated with a Customer"
                        ElseIf dst.Tables("POTPACKC").Rows.Find(row.Item("CUST_CODE") & "") Is Nothing Then
                            EMsg &= vbCr & $"No Packing information for Customer {row.Item("CUST_CODE") & ""}"
                        Else
                            PO_ORDER_NO = row.Item("PO_ORDER_NO")
                            PO_SPEC_ORDR_NO = row.Item("PO_SPEC_ORDR_NO") & ""
                            CUST_CODE = row.Item("CUST_CODE") & ""
                            STYLE_CODE_PFX = row.Item("STYLE_CODE_PFX") & ""
                            If STYLE_CODE_PFX <> "" Then Absx1.txtFor("STYLE_CODE_PFX").Text = STYLE_CODE_PFX
                            If PO_SPEC_ORDR_NO.ToUpper.StartsWith("INITIAL") Then INITIAL_ORDER = "1"
                        End If
                    End If

                    If INITIAL_ORDER = "1" Then
                        ASCMAIN1.sql = "Select Count (*) from POTORDR2 where PO_ORDER_NO = :PARM1"
                        Dim PO_lines As Integer = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {PO_ORDER_NO}))
                        If PO_lines = 0 Then
                            EMsg &= vbCr & $"No Lines on PO {PO_REFERENCE}"
                        End If

                        If Absx1.txtFor("PO_REFERENCE2").Text.Length <> 0 Then
                            If Absx1.txtFor("PO_REFERENCE2").Text.Trim = Absx1.txtFor("PO_REFERENCE").Text.Trim Then
                                EMsg &= vbCr & $"PO Reference cannot be the same as 2nd PO Reference"
                            Else
                                PO_REFERENCE2 = Absx1.txtFor("PO_REFERENCE2").Text
                                ASCMAIN1.sql = "Select * from POTORDR1 where PO_REFERENCE = :PARM1"
                                Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "V", New String() {PO_REFERENCE2})
                                If row Is Nothing Then
                                    EMsg &= vbCr & $"No record 2nd PO Reference {PO_REFERENCE2}"
                                Else
                                    If row.Item("VEND_CODE") & "" <> VEND_CODE Then
                                        EMsg &= vbCr & $"Invalid 2nd PO Reference {PO_REFERENCE2}"
                                    ElseIf row.Item("PO_STATUS") & "" <> "O" Then
                                        EMsg &= vbCr & $"2nd PO Reference {PO_REFERENCE2} is not Open"
                                    Else
                                        PO_ORDER_NO2 = row.Item("PO_ORDER_NO")
                                        STYLE_CODE_PFX2 = row.Item("STYLE_CODE_PFX") & ""
                                    End If
                                End If
                            End If
                        End If

                    Else
                        If Absx1.txtFor("STYLE_CODE_PFX").Text.Length = 0 Then
                            EMsg &= vbCr & "You must enter a Style Code Prefix"
                        Else
                            If PO_ORDER_NO <> "" Then
                                STYLE_CODE_PFX = Absx1.txtFor("STYLE_CODE_PFX").Text
                                ASCMAIN1.sql = "Select Count (*) from POTORDR2 where PO_ORDER_NO = :PARM1 and STYLE_CODE like :PARM2 || '%'"
                                Dim PO_lines As Integer = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {PO_ORDER_NO, STYLE_CODE_PFX}))
                                If PO_lines = 0 Then
                                    EMsg &= vbCr & $"No Lines on PO {PO_REFERENCE} with Style Code Prefix {STYLE_CODE_PFX}"
                                End If
                            End If
                        End If
                    End If

                End If

                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & Absx1.txtFor("VEND_CODE").Text) Then Exit Sub
                'End If

            Case "View", "Edit"

                unFinalize = False

                PACK_LIST_NO = Absx1.txtFor("PACK_LIST_NO").Text
                If PACK_LIST_NO = "" Then
                    EMsg &= vbCr & "You must specify Packing List No to View"
                Else
                    Dim row As DataRow = LookUp("POTPACK1", PACK_LIST_NO)
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & PACK_LIST_NO & " on File"
                    Else
                        If VEND_CODE_USER <> "" And row.Item("VEND_CODE") <> VEND_CODE_USER Then
                            EMsg &= vbCr & "No Record of Document " & PACK_LIST_NO & " on File"
                        End If
                        If row.Item("PACK_LIST_STATUS") & "" = "D" Then
                            EMsg &= vbCr & $"Packing List {PACK_LIST_NO} has been Deleted"
                        Else
                            If eItemKey = "Edit" Then

                                If row.Item("PACK_LIST_STATUS") & "" = "F" Then
                                    Dim VBKG_NO As String = row.Item("VBKG_NO") & ""
                                    If VBKG_NO <> "" Then
                                        EMsg &= vbCr & $"Packing List {PACK_LIST_NO} has already been listed on Booking No {VBKG_NO}"
                                        EMsg &= vbCr & "- Un-Finalizing Not permitted"
                                    Else
                                        If MsgBox("Already Finalized - do you want to un-Finalize?", MsgBoxStyle.YesNo,
                                              "IMPORTANT - LPNs will be regenerated") = MsgBoxResult.No Then
                                            Exit Sub
                                        End If
                                        unFinalize = True
                                    End If
                                End If

                                If row.Item("PACK_LIST_STATUS") & "" = "F" And Not unFinalize Then
                                    EMsg &= vbCr & "Document " & PACK_LIST_NO & " Is Finalized - no editing permitted"
                                End If

                                If EMsg = "" Then
                                    If Not ASCMAIN1.Logical_Lock("POTPACK1", PACK_LIST_NO) Then Exit Sub
                                    ' If Not ASCMAIN1.Logical_Lock("POTORDR1", "PO:" & row.Item("VEND_CODE")) Then Exit Sub

                                End If
                            End If
                        End If
                    End If
                End If

            Case "Carton Confirm"
                If grdPOTPACKX.ActiveRow Is Nothing OrElse Not grdPOTPACKX.ActiveRow.IsDataRow Then
                    EMsg &= vbCr & "You must select a row from the grid and then select this option"
                Else
                    PACK_LIST_NO = grdPOTPACKX.ActiveRow.Cells("PACK_LIST_NO").Value
                    Dim row As DataRow = LookUp("POTPACK1", PACK_LIST_NO)
                    If row.Item("PACK_LIST_STATUS") & "" <> "F" Then
                        EMsg &= vbCr & $"Packing List {PACK_LIST_NO} is not Finalized"
                    End If
                End If



                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("POTPACK1", PACK_LIST_NO) Then Exit Sub
                End If

            Case "Update"

                If EntryMode = "L" Then

                    If chkSplitRemoved.Checked Then
                        If dst.Tables("POTLPNL1").Select("CONF_REM = '1'").Length = 0 Then
                            EMsg &= vbCr & "No Cartons were Removed from this Packing List"
                        End If
                    End If

                    If EMsg = "" Then
                        If chkSplitRemoved.Checked Then
                            Dim CTNS As Integer = dst.Tables("POTLPNL1").Select("CONF_REM = '1'").Length
                            If MsgBox($"You have chosen to Remove {CTNS} Cartons from this Packing List," & vbCrLf & " and to Create a New Packing list with those Cartons." _
                                & vbCrLf & vbCrLf & "This action is non-reversible." _
                                & vbCrLf & "You may split additional cartons later, but they will be split onto another new packing list." _
                                & vbCrLf & vbCrLf & "Are you sure that you want to Split this Packing List?",
                                MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If

                Else

                    If Absx1.txtFor("PACK_LIST_DESC").Text.Length = 0 Then
                        EMsg &= vbCr & "You must supply a Packing List Description"
                    End If

                    Dim DT As Date = Absx1.dteFor("PACK_LIST_DATE").Value & ""
                    If DT & "" = "" Then
                        EMsg &= vbCr & "Packing List Date Is Mandatory"
                    Else
                        '  TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                    End If

                    If Absx1.txtFor("VEND_CODE").Text.Length = 0 Then
                        EMsg &= vbCr & "You must supply a Valid Supplier Code"
                    Else
                        Dim row As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                        If IsNothing(row) Then
                            EMsg &= vbCr & "Supplier Entered Is Not Valid"
                        Else
                            If row.Item("VEND_STATUS") & "" <> "A" Then
                                EMsg &= vbCr & "Supplier Entered Is Not Active"
                            End If
                        End If
                    End If

                    Dim EMsg2 As String = Generate_Carton_Nos()
                    EMsg &= EMsg2

                    Dim TOTAL_CARTONS As Integer = Val(dst.Tables("POTPACK2").Compute("SUM(CARTON_COUNT)", "") & "")
                    Dim CARTON_COUNTer As Integer = 0
                    For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "CARTON_NO_START")
                        Dim CARTON_COUNT As Integer = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")

                        Dim CARTON_NO_START As Integer = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")
                        Dim PACK_LIST_SHEET_NO As Integer = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")
                        If CARTON_NO_START <> CARTON_COUNTer + 1 Then
                            EMsg &= vbCr & $"Unexpected Starting Carton {CStr(CARTON_NO_START)} on Sheet {CStr(PACK_LIST_SHEET_NO)} - was expecting {CStr(CARTON_COUNTer + 1)}"
                            CARTON_COUNTer += CARTON_COUNT
                            Exit For
                        End If

                        Dim SQLW As String = $"PACK_LIST_NO = '{PACK_LIST_NO}' and PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}"

                        If INITIAL_ORDER = "1" Then
                            Dim TOTAL_UNITS As Integer = Val(rowPOTPACK2.Item("TOTAL_UNITS") & "")

                            Dim TOTAL_UNITS_CALC As Integer = CARTON_COUNT * Val(dst.Tables("POTPACK3").Compute("SUM(TOTAL_UNITS)", SQLW) & "")
                            If TOTAL_UNITS <> TOTAL_UNITS_CALC Then
                                EMsg &= vbCr & $"Total Units in PO {TOTAL_UNITS_CALC} does not agree with Total Units Packed {TOTAL_UNITS} on Sheet {PACK_LIST_SHEET_NO}"
                            End If

                            If dst.Tables("POTPACK3").Select("ISNULL(CARTON_PACK,0) = 0").Length <> 0 Then
                                EMsg &= vbCr & $"No Carton Pack specified for at least 1 Style in Details of Sheet {PACK_LIST_SHEET_NO}"
                            End If
                        Else


                            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select(SQLW, "CARTON_NO_START")
                                CARTON_COUNTer += 1

                                Dim CARTON_NO_START3 As Integer = Val(rowPOTPACK3.Item("CARTON_NO_START") & "")
                                Dim CARTON_NO_END3 As Integer = Val(rowPOTPACK3.Item("CARTON_NO_END") & "")
                                Dim CARTON_COUNT3 As Integer = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                                Dim PACK_LIST_SHEET_LNO As Integer = Val(rowPOTPACK3.Item("PACK_LIST_SHEET_LNO") & "")
                                If CARTON_NO_START3 <> CARTON_COUNTer Then
                                    EMsg &= vbCr & $"Unexpected Starting Carton {CStr(CARTON_NO_START3)} on Sheet {CStr(PACK_LIST_SHEET_NO)}, Line {PACK_LIST_SHEET_LNO} - was expecting {CStr(CARTON_COUNTer)}"
                                    Exit For
                                End If
                                CARTON_COUNTer += CARTON_COUNT3 - 1
                                If CARTON_NO_END3 <> CARTON_COUNTer Then
                                    EMsg &= vbCr & $"Unexpected Ending Carton {CStr(CARTON_NO_END3)} on Sheet {CStr(PACK_LIST_SHEET_NO)}, Line {PACK_LIST_SHEET_LNO} - was expecting {CStr(CARTON_COUNTer)}"
                                    Exit For
                                End If
                            Next

                        End If
                    Next

                    'Dim EMsg2 As String = Generate_Carton_Nos()
                    'EMsg &= EMsg2

                    If EMsg = "" Then
                        If chkFinalize.Checked Then
                            If MsgBox("You have chosen to Finalize this Packing List upon Update." _
                                    & vbCrLf & vbCrLf & "Once you have Finalized, LPNs for Barcodes will be generated," _
                                    & vbCrLf & " And you will Not be able to make further changes." _
                                    & vbCrLf & vbCrLf & "Are you sure that you want to Finalize this Packing List?",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    Else
                        If chkFinalize.Checked Then
                            ' NOT ALLOWED TO OVERRIDE EMSG IF FINALIZING
                        Else
                            If MsgBox(EMsg & vbCrLf & vbCrLf & "OK to Update Anyway?", MsgBoxStyle.OkCancel,
                                      "There are Errors in this Packing Entry") = MsgBoxResult.Cancel Then
                                Exit Sub
                            Else
                                EMsg = ""
                            End If
                        End If
                    End If
                End If

            Case "Delete"
                If MsgBox("OK to Delete Packing List?", MsgBoxStyle.YesNo,
          "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Print Labels"
                Dim tblName As String = "POTPACK3"
                If INITIAL_ORDER = "1" Then
                    tblName = "POTPACK2"
                End If
                If dst.Tables(tblName).Select("BARCODE_START IS NULL").Length > 0 Then
                    Dim RESULT As MsgBoxResult = MsgBox("Some Packing Details do not have LPNs." & vbCrLf & vbCrLf & "(Re)Generate LPNs Now?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Verification to Generate LPNs")
                    If RESULT = MsgBoxResult.Cancel Then
                        Exit Sub
                    ElseIf RESULT = MsgBoxResult.Yes Then
                        Generate_LPN_Report_File()
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

            Case "Refresh"
                Refresh_Documents()

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Add Sheet"

                WorkbookView1.GetLock()
                Dim wsx As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.ActiveWorksheet
                'Dim ws As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.Worksheets.Add()
                Dim newSheet As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.ActiveWorksheet.CopyAfter(WorkbookView1.ActiveWorkbook.ActiveWorksheet)
                WorkbookView1.ReleaseLock()

            Case "Print Labels"
                Print_Labels()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Carton Confirm"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"

                Update_Record()

                If EntryMode = "V" Then

                Else

                    If chkFinalize.Checked Then
                        If chkFinalize.Tag & "" = "X" Then
                            chkFinalize.Tag = ""
                        Else
                            Generate_LPN_Report_File()
                            Print_Labels()
                        End If

                        Check_for_Overbooked()
                    End If
                End If


                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Export XLS"
                Export_XLS()

            Case "Generate Start/End"
                Dim msg As String = Generate_Carton_Nos()
                If msg <> "" Then
                    MsgBox(msg, MsgBoxStyle.OkOnly, "Note: there are issues in the data")
                End If
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode

                    If EntryMode = "V" And ScreenMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode

                    If EntryMode = "L" Then
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                        .Items("Delete").Visible = False
                        .Items("Print Labels").Visible = False

                        lblLastScan.Text = ""

                    Else
                        If ScreenMode And EntryMode <> "N" And EntryMode <> "E" Then
                            .Items("Update").Settings.Enabled = not_iScreenMode
                            .Items("Cancel").Settings.Enabled = not_iScreenMode
                            .Items("Delete").Settings.Enabled = not_iScreenMode
                            .Items("Print Labels").Visible = rowPOTPACK1.Item("PACK_LIST_STATUS") & "" = "F" And EntryMode = "V"
                        Else
                            .Items("Update").Settings.Enabled = iScreenMode
                            .Items("Cancel").Settings.Enabled = iScreenMode
                            .Items("Delete").Settings.Enabled = iScreenMode
                            .Items("Print Labels").Visible = False
                        End If
                    End If


                    If ScreenMode Then
                        .Items("New").Visible = False
                        .Items("View").Visible = False
                        .Items("Edit").Visible = (EntryMode = "V")
                        .Items("Carton Confirm").Visible = False
                        .Items("Refresh").Visible = False
                    Else
                        .Items("New").Visible = True
                        .Items("View").Visible = True
                        .Items("Edit").Visible = True
                        .Items("Carton Confirm").Visible = True
                        .Items("Refresh").Visible = True
                    End If

                    .Items("Update").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "L")
                    .Items("Done").Visible = ScreenMode And (EntryMode = "V")

                    If ScreenMode And EntryMode = "E" Then
                        .Items("Delete").Visible = True
                    Else
                        .Items("Delete").Visible = False
                    End If

                    If ScreenMode Then
                        .Items("Separator1").Visible = True
                        .Items("Export XLS").Visible = True
                        .Items("Generate Start/End").Visible = (EntryMode = "N" Or EntryMode = "E") And Not (INITIAL_ORDER = "1")
                    Else
                        .Items("Separator1").Visible = False
                        .Items("Export XLS").Visible = False
                        .Items("Generate Start/End").Visible = False
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                        .Items("Add Sheet").Visible = False ' True
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                        .Items("Add Sheet").Visible = False
                    End If
                End With

                ' .Groups("Totals").Visible = ScreenMode
                .Groups("Show").Visible = Not ScreenMode

                .Groups("Carton Types").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")


                .Groups("LPNs").Visible = Not ((Not ScreenMode Or EntryMode <> "V") OrElse rowPOTPACK1.Item("PACK_LIST_STATUS") <> "F")
                If EntryMode = "L" Then
                    .Groups("LPNs").Visible = False
                End If

            End With
        End If

        tab1.Tabs("Packing Lists XLS").Visible = False


        lblCUST_CODE.Visible = ScreenMode
        txtCUST_CODE.Visible = ScreenMode

        lblSTYLE_CODE_PFX.Visible = ScreenMode And (INITIAL_ORDER = "1")
        txtSTYLE_CODE_PFX.Visible = ScreenMode And (INITIAL_ORDER = "1")
        If Not ScreenMode Then
            txtSTYLE_CODE_PFX2.Visible = False
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E" Or EntryMode = "V")
        SplitContainer3.Visible = ScreenMode And (EntryMode = "L")
        grpHeader.Visible = ScreenMode

        chkFinalize.Visible = Not InquiryMode And (EntryMode = "N" Or EntryMode = "E")
        chkForceLPNRegen.Visible = chkFinalize.Visible

        splPOTPACKX.Visible = Not ScreenMode

        If ScreenMode Then

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            'If EntryMode = "E" Or EntryMode = "N" Then
            '    Set_Read_Only_for_ctl(Absx1.txtFor("LC_REF_NO"), False)
            '    Set_Read_Only_for_ctl(Absx1.dteFor("LC_DATE"), False)
            '    '   Set_Read_Only_for_ctl(Absx1.txtFor("CURR_CODE"), True)
            'End If

            If EntryMode = "L" Then
                optConfirmTo.Value = "S"
                txtLPN.Focus()
            Else
                If INITIAL_ORDER = "1" Then
                    grdPOTLPNLX.DisplayLayout.CaptionVisible = DefaultableBoolean.False
                Else
                    grdPOTLPNLX.DisplayLayout.CaptionVisible = DefaultableBoolean.True
                End If


                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTPACK2, grdPOTPACK3}
                    For Each GCOL As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                        If GCOL.CellActivation = Activation.AllowEdit Then
                            If EntryMode = "N" Or EntryMode = "E" Then
                                ' GCOL.CellAppearance.BackColor = System.Drawing.Color.Khaki
                                GCOL.CellAppearance.BackColor = System.Drawing.Color.PowderBlue
                            Else
                                GCOL.CellAppearance.BackColor = System.Drawing.Color.Empty
                            End If
                        End If
                    Next

                    If EntryMode = "N" Or EntryMode = "E" Then
                        With grd.DisplayLayout.Override
                            If grd.Name = "grdPOTPACK3" Or grd.Name = "grdPOTPACK2" Then
                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                .AllowDelete = DefaultableBoolean.True
                                .AllowUpdate = DefaultableBoolean.True
                            Else
                                '    '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                                '    '.AllowDelete = DefaultableBoolean.True
                                '    '.AllowUpdate = DefaultableBoolean.True

                                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                                .AllowDelete = DefaultableBoolean.False
                                .AllowUpdate = DefaultableBoolean.True
                            End If

                        End With
                    Else
                        With grd.DisplayLayout.Override
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                            .AllowUpdate = DefaultableBoolean.False
                        End With
                    End If
                Next

                With grdPOTPACK3.DisplayLayout.Bands(0)
                    If Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
                        .Columns("BARCODE_START").Hidden = True
                        .Columns("BARCODE_END").Hidden = True
                    Else
                        ' keep starting & ending LPN columns hidden until we figure out how to handle non-contiguous ranges
                        .Columns("BARCODE_START").Hidden = True ' False
                        .Columns("BARCODE_END").Hidden = True ' False
                    End If

                    If (INITIAL_ORDER = "1") Then
                        dst.Tables("POTPACK3").Columns("TOTAL_GRS_WGT").Expression = "CARTON_GRS_WGT * PARENT(POTPACK2_POTPACK3).CARTON_COUNT"
                        dst.Tables("POTPACK3").Columns("TOTAL_NET_WGT").Expression = "CARTON_NET_WGT * PARENT(POTPACK2_POTPACK3).CARTON_COUNT"
                    Else
                        dst.Tables("POTPACK3").Columns("TOTAL_GRS_WGT").Expression = "CARTON_COUNT * CARTON_GRS_WGT"
                        dst.Tables("POTPACK3").Columns("TOTAL_NET_WGT").Expression = "CARTON_COUNT * CARTON_NET_WGT"
                    End If

                End With

                With grdPOTPACK3.DisplayLayout.Bands(0)
                    'XCARTON_DIMENSIONS
                    For Each C As String In New String() {"CARTON_COUNT", "CARTON_PACK", "CARTON_GRS_WGT", "CARTON_NET_WGT", "PKG_CODE", "CARTON_NO_START", "CARTON_NO_END", "BARCODE_START", "BARCODE_END", "TOTAL_GRS_WGT", "TOTAL_NET_WGT"}
                        .Columns(C).Hidden = (INITIAL_ORDER = "1")
                        If C = "CARTON_PACK" And rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" Then ' Not PO_REFERENCE.StartsWith("WM") Then
                            .Columns(C).Hidden = False
                        End If
                    Next

                    ' keep starting & ending LPN columns hidden until we figure out how to handle non-contiguous ranges
                    .Columns("BARCODE_START").Hidden = True
                    .Columns("BARCODE_END").Hidden = True


                    .Columns("CARTON_DIMENSIONS").Hidden = True

                    'For Each C As String In New String() {"CARTON_PACK"}
                    '    If INITIAL_ORDER = "1" Then
                    '        .Columns(C).CellActivation = Activation.NoEdit ' MAYBE WE WILL NEED TO OPEN UP FOR EDIT
                    '        .Columns(C).CellAppearance.BackColor = System.Drawing.Color.Empty
                    '        .Columns(C).Header.Appearance.BackColor = System.Drawing.Color.Empty
                    '    Else
                    '        .Columns(C).CellActivation = Activation.AllowEdit
                    '        .Columns(C).Header.Appearance.BackColor = System.Drawing.Color.DodgerBlue
                    '    End If
                    'Next
                End With

                With grdPOTPACK2.DisplayLayout.Bands(0)
                    For Each C As String In New String() {"CARTON_NO_START"}
                        .Columns(C).Hidden = (INITIAL_ORDER = "1")
                    Next
                    For Each C As String In New String() {"CARTON_PACK", "CARTON_COUNT", "BARCODE_START", "BARCODE_END"}
                        .Columns(C).Hidden = Not (INITIAL_ORDER = "1")
                    Next

                    ' keep starting & ending LPN columns hidden until we figure out how to handle non-contiguous ranges
                    .Columns("BARCODE_START").Hidden = True ' False
                    .Columns("BARCODE_END").Hidden = True ' False

                    For Each C As String In New String() {"CARTON_PACK_HOLD"}
                        .Columns(C).Hidden = True
                    Next

                    .Columns("CARTON_GRS_WGT").Hidden = Not (INITIAL_ORDER = "1")
                    .Columns("CARTON_NET_WGT").Hidden = Not (INITIAL_ORDER = "1")
                    .Columns("CARTON_DIMENSIONS").Hidden = True  ' Not (INITIAL_ORDER = "1")
                    .Columns("PKG_CODE").Hidden = Not (INITIAL_ORDER = "1")

                    If (INITIAL_ORDER = "1") Then
                        .Columns("TOTAL_CARTONS").Header.Caption = "Styles"
                        dst.Tables("POTPACK2").Columns("TOTAL_UNITS").Expression = "CARTON_PACK * CARTON_COUNT"
                        dst.Tables("POTPACK2").Columns("TOTAL_GRS_WGT").Expression = "CARTON_COUNT * CARTON_GRS_WGT"
                        dst.Tables("POTPACK2").Columns("TOTAL_NET_WGT").Expression = "CARTON_COUNT * CARTON_NET_WGT"
                    Else
                        .Columns("TOTAL_CARTONS").Header.Caption = "Cartons"
                        dst.Tables("POTPACK2").Columns("TOTAL_UNITS").Expression = "SUM(CHILD.TOTAL_UNITS)"
                        dst.Tables("POTPACK2").Columns("TOTAL_GRS_WGT").Expression = "SUM(CHILD.TOTAL_GRS_WGT)"
                        dst.Tables("POTPACK2").Columns("TOTAL_NET_WGT").Expression = "SUM(CHILD.TOTAL_NET_WGT)"
                    End If
                End With
            End If

            Set_Read_Only_for_ctl(Absx1.optFor("PACK_LIST_STATUS"), True)
            Set_Read_Only_for_ctl(Absx1.chkFor("INITIAL_ORDER"), True)

            Display_Totals()

        Else
            Clear_Record()

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTPACK1", "POTPACK2", "POTPACK3", "POTLPNL1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If VEND_CODE_USER <> "" Then
            Absx1.txtFor("VEND_CODE").Text = VEND_CODE_USER
            Absx1.txtFor("VEND_CODE").ReadOnly = True
        Else
            Absx1.txtFor("VEND_CODE").Text = ""
        End If

        'lblSTYLE_CODE_PFX.Visible = False
        'txtSTYLE_CODE_PFX.Visible = False
        'txtSTYLE_CODE_PFX2.Visible = False

        chkFinalize.Checked = False
        chkFinalize.Tag = ""
        chkForceLPNRegen.Checked = False

        chkSplitRemoved.Checked = False
        chkDblClickToEdit.Checked = False
        Check_for_MultiPO()

        Refresh_Documents()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        PO_ORDER_NO2 = ""
        PO_REFERENCE2 = ""
        STYLE_CODE_PFX2 = ""

        If EntryMode = "N" Then
            PACK_LIST_NO = ASCMAIN1.Next_Control_No("POTPACK1.PACK_LIST_NO")
            rowPOTPACK1 = dst.Tables("POTPACK1").NewRow
            With rowPOTPACK1
                .Item("PACK_LIST_NO") = PACK_LIST_NO
                .Item("VEND_CODE") = HFs("VEND_CODE")
                .Item("PACK_LIST_DESC") = PO_SPEC_ORDR_NO
                .Item("PACK_LIST_DATE") = DATETIME_STAMP.Date
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("PACK_LIST_STATUS") = "O"
                .Item("PO_REFERENCE") = PO_REFERENCE
                If INITIAL_ORDER = "1" Then
                    .Item("STYLE_CODE_PFX") = STYLE_CODE_PFX
                    .Item("INITIAL_ORDER") = "1"
                    PO_REFERENCE2 = HFs("PO_REFERENCE2") & ""
                    If PO_REFERENCE2 <> "" Then
                        ASCMAIN1.sql = "Select PO_ORDER_NO, STYLE_CODE_PFX from POTORDR1 where PO_REFERENCE = :PARM1"
                        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "V", New String() {PO_REFERENCE2})
                        PO_ORDER_NO2 = row.Item("PO_ORDER_NO") & ""
                        STYLE_CODE_PFX2 = row.Item("STYLE_CODE_PFX") & ""
                    End If
                Else
                    .Item("STYLE_CODE_PFX") = ""
                    .Item("INITIAL_ORDER") = "0"
                End If
                .Item("STYLE_CODE_PFX2") = STYLE_CODE_PFX2
                .Item("PO_REFERENCE2") = PO_REFERENCE2
                .Item("PO_ORDER_NO2") = PO_ORDER_NO2
                .Item("PO_ORDER_NO") = PO_ORDER_NO
                .Item("CUST_CODE") = CUST_CODE
            End With

            dst.Tables("POTPACK1").Rows.Add(rowPOTPACK1)

        Else
            rowPOTPACK1 = Fill_Record("POTPACK1", PACK_LIST_NO)
            VEND_CODE = rowPOTPACK1.Item("VEND_CODE")
            If VEND_CODE_USER <> "" And VEND_CODE <> VEND_CODE_USER Then
                MsgBox("Issue with Vendor Code", MsgBoxStyle.OkOnly, "Please Call ABS")
                Throw New Exception("Issue with Vendor Code")
            End If
            PO_REFERENCE = rowPOTPACK1.Item("PO_REFERENCE")
            STYLE_CODE_PFX = rowPOTPACK1.Item("STYLE_CODE_PFX") & ""
            PO_ORDER_NO = rowPOTPACK1.Item("PO_ORDER_NO")
            INITIAL_ORDER = rowPOTPACK1.Item("INITIAL_ORDER")
            CUST_CODE = rowPOTPACK1.Item("CUST_CODE") & ""

            PO_REFERENCE2 = rowPOTPACK1.Item("PO_REFERENCE2") & ""
            STYLE_CODE_PFX2 = rowPOTPACK1.Item("STYLE_CODE_PFX2") & ""
            PO_ORDER_NO2 = rowPOTPACK1.Item("PO_ORDER_NO2") & ""

            If unFinalize Then
                rowPOTPACK1.Item("PACK_LIST_STATUS") = "O"
            End If

            dst.AcceptChanges()
        End If

        rowPOTPACKC = dst.Tables("POTPACKC").Rows.Find(CUST_CODE)

        PACK_LIST_STATUS = rowPOTPACK1.Item("PACK_LIST_STATUS")


        BARCODE_PFX = "Y" ' NEED TO GET THIS FROM VENDOR MASTER
        ' AND VENDORS WITHOUT A PREFIX ARE NOT PERMITTED TO USE THIS SCREEN
        Fill_Records("WHTPKGM1", BARCODE_PFX)
        Sort_grdColumns(grdWHTPKGM1, "CARTON_DIMENSIONS")
        Reset_ValueLists()

        Check_for_MultiPO()

        EnforceConstraints(False)

        Fill_Records("POTORDR2", PO_ORDER_NO)

        If EntryMode = "N" Then

            Fill_Records("POTORDRD", PO_ORDER_NO)

            Dim EMSGS As String = ""

            If INITIAL_ORDER = "1" Then

                Dim COLOR_CODEs As New List(Of String)
                If rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" <> "1" Then ' If PO_REFERENCE.StartsWith("WM") Then
                    COLOR_CODEs.Add("AST")
                Else
                    ASCMAIN1.sql = "Select Distinct COLOR_CODE from POTORDR2 where PO_ORDER_NO = :PARM1 and PO_QTY_OPN > 0"
                    For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {PO_ORDER_NO}).Select("", "COLOR_CODE")
                        COLOR_CODEs.Add(row.Item("COLOR_CODE"))
                    Next
                End If
                Dim PACK_LIST_SHEET_NO_ctr As Integer = 0

                For Each COLOR_CODE As String In COLOR_CODEs
                    PACK_LIST_SHEET_NO_ctr += 1
                    Dim rowPOTPACK2 As DataRow = dst.Tables("POTPACK2").NewRow
                    rowPOTPACK2.Item("PACK_LIST_NO") = PACK_LIST_NO
                    rowPOTPACK2.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr

                    Dim COLOR_CODE_SFX As String = ""
                    If rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" Then ' If Not PO_REFERENCE.StartsWith("WM") Then
                        COLOR_CODE_SFX = " - Color " & COLOR_CODE
                    End If
                    rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") = PO_REFERENCE & "-" & CStr(PO_SPEC_ORDR_NO) & COLOR_CODE_SFX

                    rowPOTPACK2.Item("COLOR_CODE") = COLOR_CODE
                    Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                    rowPOTPACK2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
                    rowPOTPACK2.Item("CARTON_NO_START") = 1
                    dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2)

                    Dim CARTON_PACK As Integer = 0
                    Dim PACK_LIST_SHEET_LNO_ctr As Integer = 0
                    For PO As Integer = 1 To 2

                        If PO = 2 Then
                            If PO_ORDER_NO2 = "" Then Exit For
                            Fill_Records("POTORDRD", PO_ORDER_NO2)
                        End If

                        Dim sqlColor As String = ""
                        If rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" Then ' If Not PO_REFERENCE.StartsWith("WM") Then
                            sqlColor = $"COLOR_CODE = '{COLOR_CODE}'"
                        End If

                        For Each rowPOTORDRD As DataRow In dst.Tables("POTORDRD").Select(sqlColor, "STYLE_CODE, COLOR_CODE")
                            COLOR_CODE = rowPOTORDRD.Item("COLOR_CODE")
                            Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").NewRow
                            With rowPOTPACK3
                                .Item("PACK_LIST_NO") = PACK_LIST_NO
                                .Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr
                                PACK_LIST_SHEET_LNO_ctr += 1
                                .Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_ctr
                                Dim STYLE_CODE As String = rowPOTORDRD.Item("STYLE_CODE")
                                .Item("STYLE_CODE") = STYLE_CODE
                                .Item("COLOR_CODE") = COLOR_CODE
                                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                                .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                                '.Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                                .Item("SIZE_CODE") = rowICTSTYL1.Item("SIZE_CODE")

                                .Item("PO_QTY_OPN") = rowPOTORDRD.Item("PO_QTY_OPN") ' NOTE THAT THERE MAY BE MORE THAN 1 LINE OPEN, SO THIS IS THE SUM
                                .Item("PO_ORDER_NO") = PO_ORDER_NO
                                .Item("PO_ORDER_LNO") = rowPOTORDRD.Item("PO_ORDER_LNO") ' NOTE THAT THERE MAY BE MORE THAN 1 LINE OPEN, SO THIS IS JUST THE MIN LINE

                                CARTON_PACK += Val(rowPOTORDRD.Item("CARTON_PACK") & "")
                                .Item("CARTON_PACK") = rowPOTORDRD.Item("CARTON_PACK")

                                .Item("CARTON_COUNT") = 1

                                Dim rowWHTSCSEQs() As DataRow = dst.Tables("WHTSCSEQ").Select($"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'")
                                If rowWHTSCSEQs.Length = 0 Then
                                    EMSGS &= vbCrLf & $"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                                    ' MsgBox($"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                                ElseIf rowWHTSCSEQs.Length > 1 Then
                                    EMSGS &= vbCrLf & $"More than 1 Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                                    ' MsgBox($"More than 1 Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                                Else
                                    .Item("CARTON_ID") = rowWHTSCSEQs(0).Item("STYLE_SEQ")
                                End If

                                'Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                                'If rowICTSTYC1 Is Nothing Then
                                '    EMSGS &= vbCrLf & $"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                                'Else
                                '    Dim CARTON_ID As Integer = Val(rowICTSTYC1.Item("CARTON_ID") & "")
                                '    If CARTON_ID <= 0 Then
                                '        EMSGS &= vbCrLf & $"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                                '    Else
                                '        Dim rowICTSTYC1s() As DataRow = dst.Tables("ICTSTYC1").Select($"CARTON_ID = {CARTON_ID} and (STYLE_CODE <> '{STYLE_CODE}' or COLOR_CODE <> '{COLOR_CODE}')")
                                '        If rowICTSTYC1s.Length > 0 Then
                                '            EMSGS &= vbCrLf & $"More than 1 Style-Color defined with Carton ID {CARTON_ID} used for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                                '        End If
                                '    End If
                                '    .Item("CARTON_ID") = CARTON_ID
                                'End If

                            End With
                            dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3)
                        Next
                        rowPOTPACK2.Item("CARTON_PACK") = CARTON_PACK
                    Next
                Next


            Else

                Dim PACK_LIST_SHEET_NO_ctr As Integer = 0
                For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("POTORDR2").Select("PO_QTY_OPN > 0"), New String() {"COLOR_CODE"}).Select("", "COLOR_CODE")
                    Dim rowPOTPACK2 As DataRow = dst.Tables("POTPACK2").NewRow
                    rowPOTPACK2.Item("PACK_LIST_NO") = PACK_LIST_NO
                    PACK_LIST_SHEET_NO_ctr += 1
                    rowPOTPACK2.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr
                    rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") = PO_REFERENCE & "-" & CStr(PACK_LIST_SHEET_NO_ctr)
                    Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                    rowPOTPACK2.Item("COLOR_CODE") = COLOR_CODE
                    Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                    rowPOTPACK2.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
                    dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2)

                    Dim PACK_LIST_SHEET_LNO_ctr As Integer = 0
                    For Each rowPOTORDRD As DataRow In dst.Tables("POTORDRD").Select($"COLOR_CODE = '{COLOR_CODE}'", "STYLE_CODE")
                        Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").NewRow
                        With rowPOTPACK3
                            .Item("PACK_LIST_NO") = PACK_LIST_NO
                            .Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_ctr
                            PACK_LIST_SHEET_LNO_ctr += 1
                            .Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_ctr
                            Dim STYLE_CODE As String = rowPOTORDRD.Item("STYLE_CODE")
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                            ' .Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                            .Item("SIZE_CODE") = rowICTSTYL1.Item("SIZE_CODE")

                            .Item("PO_QTY_OPN") = rowPOTORDRD.Item("PO_QTY_OPN") ' NOTE THAT THERE MAY BE MORE THAN 1 LINE OPEN, SO THIS IS THE SUM
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = rowPOTORDRD.Item("PO_ORDER_LNO") ' NOTE THAT THERE MAY BE MORE THAN 1 LINE OPEN, SO THIS IS JUST THE MIN LINE

                            Dim rowWHTSCSEQs() As DataRow = dst.Tables("WHTSCSEQ").Select($"STYLE_CODE = '{STYLE_CODE}' and COLOR_CODE = '{COLOR_CODE}'")
                            If rowWHTSCSEQs.Length = 0 Then
                                EMSGS &= vbCrLf & $"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                                ' MsgBox($"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                            ElseIf rowWHTSCSEQs.Length > 1 Then
                                EMSGS &= vbCrLf & $"More than 1 Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                                ' MsgBox($"More than 1 Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}", MsgBoxStyle.OkOnly, "Please Report to Vandale")
                            Else
                                .Item("CARTON_ID") = rowWHTSCSEQs(0).Item("STYLE_SEQ")
                            End If

                            'Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                            'If rowICTSTYC1 Is Nothing Then
                            '    EMSGS &= vbCrLf & $"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                            'Else
                            '    Dim CARTON_ID As Integer = Val(rowICTSTYC1.Item("CARTON_ID") & "")
                            '    If CARTON_ID <= 0 Then
                            '        EMSGS &= vbCrLf & $"No Carton ID defined for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                            '    Else
                            '        Dim rowICTSTYC1s() As DataRow = dst.Tables("ICTSTYC1").Select($"CARTON_ID = {CARTON_ID} and (STYLE_CODE <> '{STYLE_CODE}' or COLOR_CODE <> '{COLOR_CODE}')")
                            '        If rowICTSTYC1s.Length > 0 Then
                            '            EMSGS &= vbCrLf & $"More than 1 Style-Color defined with Carton ID {CARTON_ID} used for Style-Color {STYLE_CODE}-{COLOR_CODE}"
                            '        End If
                            '    End If
                            '    .Item("CARTON_ID") = CARTON_ID
                            'End If

                        End With
                        dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3)
                    Next

                Next
            End If

            If EMSGS <> "" Then
                MsgBox(EMSGS, MsgBoxStyle.OkOnly, "Please Report to Vandale")
            End If

        Else
            Fill_Records("POTPACK2", PACK_LIST_NO)
            Fill_Records("POTPACK3", PACK_LIST_NO)

            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select("")
                Dim STYLE_CODE As String = rowPOTPACK3.Item("STYLE_CODE")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                rowPOTPACK3.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            Next
            Generate_Carton_Nos()
        End If


        If INITIAL_ORDER = "1" AndAlso rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" Then ' Not PO_REFERENCE.StartsWith("WM") Then
            Sort_grdColumns(grdPOTPACK2, "COLOR_CODE", True)
        End If

        If EntryMode = "V" And rowPOTPACK1.Item("PACK_LIST_STATUS") = "F" Then
            Fill_Records("POTLPNL1", PACK_LIST_NO)
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grdPOTPACK2.Rows
            grow.PerformAutoSize()
        Next

        If EntryMode = "L" Then
            Fill_Records("POTLPNL1", PACK_LIST_NO)
            Sort_grdColumns(grdPOTLPNL1, "BARCODE")
            Dim PACK_LIST_DESC As String = rowPOTPACK1.Item("PACK_LIST_DESC") & ""
            grdPOTLPNL1.Text = $"LPNs for Packing List {PACK_LIST_NO} - {PACK_LIST_DESC}, PO {PO_REFERENCE}" & IIf(PO_REFERENCE2 = "", "", $", {PO_REFERENCE2}")
        End If



        Dim FILENAME_source As String = "R:\VDI\Templates" & "\" & "PACKLIST.xlsx"
        If ASCMAIN1.Running_in_VS Then FILENAME_source = "C:\Share\VDI\Templates\PACKLIST.xlsx"

        Dim FILENAME As String = ASCMAIN1.Folders("Work") & "\" & "PACKLIST.xlsx"

        My.Computer.FileSystem.CopyFile(FILENAME_source, FILENAME, True)

        WorkbookView1.GetLock()
        WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        WorkbookView1.ReleaseLock()

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        If EntryMode = "L" Then

            If chkSplitRemoved.Checked Then
                Dim PACK_LIST_NO_split As String = ASCMAIN1.Next_Control_No("POTPACK1.PACK_LIST_NO")

                Dim rowPOTPACK1_split As DataRow = dst.Tables("POTPACK1").NewRow
                With rowPOTPACK1_split
                    .ItemArray = rowPOTPACK1.ItemArray
                    .Item("PACK_LIST_NO") = PACK_LIST_NO_split
                    .Item("PACK_LIST_DATE") = DATETIME_STAMP.Date
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    '.Item("PACK_LIST_STATUS") = "O"
                End With
                dst.Tables("POTPACK1").Rows.Add(rowPOTPACK1_split)


                For Each rowPOTLPNL1 As DataRow In dst.Tables("POTLPNL1").Select("SHIP_CONF = 'R'")

                    rowPOTLPNL1.Item("PACK_LIST_NO") = PACK_LIST_NO_split
                    rowPOTLPNL1.Item("SHIP_CONF") = DBNull.Value

                    Dim PACK_LIST_SHEET_NO As Integer = Val(rowPOTLPNL1.Item("PACK_LIST_SHEET_NO") & "")
                    Dim PACK_LIST_SHEET_LNO As Integer = Val(rowPOTLPNL1.Item("PACK_LIST_SHEET_LNO") & "")

                    Dim rowPOTPACK2 As DataRow = dst.Tables("POTPACK2").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO})
                    If PACK_LIST_SHEET_LNO = 0 Then ' CARTON IS CONNECTED TO POTPACK2

                        rowPOTPACK2.Item("CARTON_COUNT") = Val(rowPOTPACK2.Item("CARTON_COUNT") & "") - 1
                        Dim rowPOTPACK2_split As DataRow = dst.Tables("POTPACK2").Rows.Find(New Object() {PACK_LIST_NO_split, PACK_LIST_SHEET_NO})
                        If rowPOTPACK2_split Is Nothing Then

                            rowPOTPACK2_split = dst.Tables("POTPACK2").NewRow
                            With rowPOTPACK2_split
                                .ItemArray = rowPOTPACK2.ItemArray
                                .Item("PACK_LIST_NO") = PACK_LIST_NO_split
                                .Item("CARTON_COUNT") = 1
                            End With
                            dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2_split)

                            Dim SQLW As String = $"PACK_LIST_NO = '{PACK_LIST_NO}' and PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}"
                            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select(SQLW)
                                Dim rowPOTPACK3_split As DataRow = dst.Tables("POTPACK3").NewRow
                                With rowPOTPACK3_split
                                    .ItemArray = rowPOTPACK3.ItemArray
                                    .Item("PACK_LIST_NO") = PACK_LIST_NO_split
                                End With
                                dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_split)
                            Next
                        Else
                            rowPOTPACK2_split.Item("CARTON_COUNT") = Val(rowPOTPACK2_split.Item("CARTON_COUNT") & "") + 1
                        End If

                        If rowPOTPACK2.Item("CARTON_COUNT") = 0 Then
                            rowPOTPACK2.Delete()
                        End If

                    Else ' CARTON IS CONNECTED TO POTPACK3

                        Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO, PACK_LIST_SHEET_LNO})
                        rowPOTPACK3.Item("CARTON_COUNT") = Val(rowPOTPACK3.Item("CARTON_COUNT") & "") - 1
                        Dim rowPOTPACK3_split As DataRow = dst.Tables("POTPACK3").Rows.Find(New Object() {PACK_LIST_NO_split, PACK_LIST_SHEET_NO, PACK_LIST_SHEET_LNO})
                        If rowPOTPACK3_split Is Nothing Then

                            Dim rowPOTPACK2_split As DataRow = dst.Tables("POTPACK2").Rows.Find(New Object() {PACK_LIST_NO_split, PACK_LIST_SHEET_NO})
                            If rowPOTPACK2_split Is Nothing Then
                                rowPOTPACK2_split = dst.Tables("POTPACK2").NewRow
                                With rowPOTPACK2_split
                                    .ItemArray = rowPOTPACK2.ItemArray
                                    .Item("PACK_LIST_NO") = PACK_LIST_NO_split
                                End With
                                dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2_split)
                            End If

                            rowPOTPACK3_split = dst.Tables("POTPACK3").NewRow
                            With rowPOTPACK3_split
                                .ItemArray = rowPOTPACK3.ItemArray
                                .Item("PACK_LIST_NO") = PACK_LIST_NO_split
                                .Item("CARTON_COUNT") = 1
                            End With
                            dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_split)


                        Else
                            rowPOTPACK3_split.Item("CARTON_COUNT") = Val(rowPOTPACK3_split.Item("CARTON_COUNT") & "") + 1
                        End If

                        If rowPOTPACK3.Item("CARTON_COUNT") = 0 Then
                            rowPOTPACK3.Delete()
                        End If
                    End If
                Next
            End If




            Dim SQLD As String = "PACK_LIST_NO = '" & PACK_LIST_NO & "'"
            Update_Record_TDA("POTLPNL1", SQLD)
            If chkSplitRemoved.Checked Then
                Update_Record_TDA("POTPACK1", SQLD)
                Update_Record_TDA("POTPACK2", SQLD)
                Update_Record_TDA("POTPACK3", SQLD)
            End If

        Else

            If (INITIAL_ORDER = "1") Then
                For Each row As DataRow In dst.Tables("POTPACK2").Select("")
                    row.Item("CARTON_PACK") = row.Item("CARTON_PACK_HOLD")
                Next
            End If

            If chkFinalize.Checked Then
                rowPOTPACK1.Item("PACK_LIST_STATUS") = "F"

                Dim tbl_BARCODE As String = "POTPACK3"
                If INITIAL_ORDER = "1" Then
                    tbl_BARCODE = "POTPACK2"
                End If

                Dim generate_LPNs As Boolean = True
                Dim BARCODE_MIN As String = dst.Tables(tbl_BARCODE).Compute("MIN(BARCODE_START)", "") & ""
                If BARCODE_MIN <> "" Then
                    generate_LPNs = Generate_LPNs_Test(BARCODE_MIN)
                End If

                If chkForceLPNRegen.Checked Then
                    generate_LPNs = True
                End If

                If generate_LPNs Then

                    For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
                        Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") & ""
                        Dim CARTON_NO_START As Int32 = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")
                        Dim PACK_LIST_SHEET_NO As Int32 = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

                        If INITIAL_ORDER = "1" Then
                            Dim CARTON_COUNT As Int32 = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                            Dim BARCODE As String = ASCMAIN1.Next_Control_No("BARCODE_" & BARCODE_PFX, CARTON_COUNT)
                            Dim BARCODE_START As String = BARCODE_PFX & BARCODE
                            Dim BARCODE_END As String = BARCODE_PFX & Format(Val(BARCODE) + CARTON_COUNT - 1, "0000000")
                            rowPOTPACK2.Item("BARCODE_START") = BARCODE_START
                            rowPOTPACK2.Item("BARCODE_END") = BARCODE_END
                        Else
                            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO") ' rowPOTPACK2.GetChildRows("POTPACK2_POTPACK3")
                                Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                                Dim BARCODE As String = ASCMAIN1.Next_Control_No("BARCODE_" & BARCODE_PFX, CARTON_COUNT)
                                Dim BARCODE_START As String = BARCODE_PFX & BARCODE
                                Dim BARCODE_END As String = BARCODE_PFX & Format(Val(BARCODE) + CARTON_COUNT - 1, "0000000")
                                rowPOTPACK3.Item("BARCODE_START") = BARCODE_START
                                rowPOTPACK3.Item("BARCODE_END") = BARCODE_END
                            Next
                        End If
                    Next

                    If BARCODE_MIN <> "" Then
                        MsgBox("Note: LPNs WERE Re-Generated", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Please Note: Labels will be re-printed")
                    End If

                Else

                    If BARCODE_MIN <> "" Then
                        MsgBox("Note: LPNs were NOT Re-Generated", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Please Note: Labels will NOT be re-printed")
                        chkFinalize.Tag = "X"
                    End If

                End If

            End If

            Dim SQLD As String = "PACK_LIST_NO = '" & PACK_LIST_NO & "'"
            INIT_LAST("POTPACK1", False, , True)

            Update_Record_TDA("POTPACK1", SQLD)
            Update_Record_TDA("POTPACK2", SQLD)
            Update_Record_TDA("POTPACK3", SQLD)


            Update_Record_TDA("WHTPKGM1")

        End If

        CommitTrans("Update Complete")

    End Sub

    Function Generate_LPNs_Test(BARCODE_MIN As String) As Boolean

        ' Dim BARCODE_PFX As String = Mid(BARCODE_MIN, 1, 1)
        Dim BARCODE_CTR As Int32 = Val(Mid(BARCODE_MIN, 2))
        Dim regeneration_required As Boolean = False

        For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
            Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") & ""
            Dim CARTON_NO_START As Int32 = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")
            Dim PACK_LIST_SHEET_NO As Int32 = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO")
                Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                If INITIAL_ORDER = "1" Then
                    CARTON_COUNT = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                End If
                Dim BARCODE As String = Format(BARCODE_CTR, "0000000")
                Dim BARCODE_START As String = BARCODE_PFX & BARCODE
                Dim BARCODE_END As String = BARCODE_PFX & Format(Val(BARCODE) + CARTON_COUNT - 1, "0000000")

                Dim rowCompare As DataRow = rowPOTPACK3
                If INITIAL_ORDER = "1" Then rowCompare = rowPOTPACK2
                If rowCompare.Item("BARCODE_START") & "" <> BARCODE_START Or rowCompare.Item("BARCODE_END") & "" <> BARCODE_END Then
                    regeneration_required = True
                    Exit For
                End If


                BARCODE_CTR += CARTON_COUNT
            Next
            If regeneration_required Then Exit For
        Next

        Return regeneration_required
    End Function

    Sub Delete_Record()
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub

        ' Dependent_Updates(-1, ORDR_NO)

        'For Each TABLE_NAME As String In New String() _
        '    {"POTPACK1", "POTPACK2", "POTPACK3"}
        '    Delete_Records_1(TABLE_NAME)
        'Next

        ASCMAIN1.sql = "Update POTPACK1 Set PACK_LIST_STATUS = 'D'" & " where PACK_LIST_NO = '" & PACK_LIST_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update POTLPNL1 Set BARCODE_STATUS = 'D'" & " where PACK_LIST_NO = '" & PACK_LIST_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where PACK_LIST_NO = '" & PACK_LIST_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("PACK_LIST_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTPACK1"
            E.COLUMN_NAME = "PACK_LIST_NO"
            E.CODE_VALUE = Absx1.txtFor("PACK_LIST_NO").Text
            E.DESC_VALUE = Absx1.txtFor("VEND_CODE").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "POTPACK1"
        E.TABLE_KEY_CAPTION = "LC Events"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PACK_LIST_NO").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                sql_where = "VEND_TYPE = 'S'"
        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTPACKX, "SS", "Show Filter", "Show GroupBox") ', "Confirm/Remove Cartons")
        Load_Popup_Menu(grdPOTORDRR, "SS", "Show Filter")
        Load_Popup_Menu(grdPOTPACK2, "B", "Add Sheet", "Add Sheets")
        Load_Popup_Menu(grdPOTPACK3, "B", "Add Line", "Add Lines", "Copy Value to All Lines", "Copy Pattern to Remaining Lines")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then
            grd = GRDs(Mid(e.SourceControl.Name, 4))
        End If

        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If


        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

            Case "grdPOTPACK2"

                If Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_pop.Tools("Add Sheet").SharedProps.Visible = (INITIAL_ORDER = "1") And rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" ' Not PO_REFERENCE.StartsWith("WM")
                    tlb_pop.Tools("Add Sheets").SharedProps.Visible = (INITIAL_ORDER = "1") And rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" ' Not PO_REFERENCE.StartsWith("WM")

                Else
                    tlb_pop.Tools("Add Sheet").SharedProps.Visible = False
                    tlb_pop.Tools("Add Sheets").SharedProps.Visible = False
                    'tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = False
                    'tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = False
                End If

            Case "grdPOTPACK3"

                If Not InquiryMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_pop.Tools("Add Line").SharedProps.Visible = Not (INITIAL_ORDER = "1")
                    tlb_pop.Tools("Add Lines").SharedProps.Visible = Not (INITIAL_ORDER = "1")
                    'XCARTON_DIMENSIONS
                    If Not grd.ActiveRow.DataChanged And grd.ActiveCell IsNot Nothing AndAlso New String() {"PKG_CODE", "CARTON_PACK", "CARTON_GRS_WGT", "CARTON_NET_WGT"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = True
                        If grd.ActiveCell.Column.Key = "CARTON_PACK" Then
                            tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = False ' HOW CAN WE DO THIS IF WE DO NOT PERMIT MTC TO CARTON_PACK WHEN INITIAL_ORDER = 1?
                        End If
                    Else
                        tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = False
                    End If

                    If Not grd.ActiveRow.DataChanged And grd.ActiveCell IsNot Nothing AndAlso New String() {"CARTON_PACK"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = (INITIAL_ORDER = "1")
                        tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = False ' HOW CAN WE DO THIS IF WE DO NOT PERMIT MTC TO CARTON_PACK WHEN INITIAL_ORDER = 1?
                    Else
                        tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = False
                    End If

                Else
                    tlb_pop.Tools("Add Line").SharedProps.Visible = False
                    tlb_pop.Tools("Add Lines").SharedProps.Visible = False
                    tlb_pop.Tools("Copy Value to All Lines").SharedProps.Visible = False
                    tlb_pop.Tools("Copy Pattern to Remaining Lines").SharedProps.Visible = False
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdSPTSFOC9"
                '    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                '        tlb_btn.SharedProps.Visible = True
                '    Else
                '        tlb_btn.SharedProps.Visible = False
                '    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Add Sheet"
                Dim PACK_LIST_SHEET_NO As Integer = Val(grdPOTPACK2.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")
                Dim PACK_LIST_SHEET_NO_max As Integer = dst.Tables("POTPACK2").Compute("MAX(PACK_LIST_SHEET_NO)", "")
                Dim rowPOTPACK2 As DataRow = dst.Tables("POTPACK2").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO})
                Dim rowPOTPACK2_new As DataRow = dst.Tables("POTPACK2").NewRow
                rowPOTPACK2_new.ItemArray = rowPOTPACK2.ItemArray
                PACK_LIST_SHEET_NO_max += 1
                rowPOTPACK2_new.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_max
                rowPOTPACK2_new.Item("BARCODE_START") = DBNull.Value
                rowPOTPACK2_new.Item("BARCODE_END") = DBNull.Value
                dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2_new)

                For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}")
                    Dim rowPOTPACK3_new As DataRow = dst.Tables("POTPACK3").NewRow
                    rowPOTPACK3_new.ItemArray = rowPOTPACK3.ItemArray
                    rowPOTPACK3_new.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_max
                    dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_new)
                Next

                Sort_grdColumns(grdPOTPACK2, "COLOR_CODE", True)

            Case "Add Sheets"
                Dim PACK_LIST_SHEET_NO_max As Integer = dst.Tables("POTPACK2").Compute("MAX(PACK_LIST_SHEET_NO)", "")

                'Dim STYLE_CODEs As New List(Of String)
                For Each row2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
                    Dim PACK_LIST_SHEET_NO As Integer = Val(row2.Item("PACK_LIST_SHEET_NO") & "")
                    'Dim STYLE_CODE As String = row3.Item("STYLE_CODE") & ""
                    'If Not STYLE_CODEs.Contains(STYLE_CODE) Then
                    Dim rowPOTPACK2 As DataRow = dst.Tables("POTPACK2").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO})
                    Dim rowPOTPACK2_new As DataRow = dst.Tables("POTPACK2").NewRow
                    rowPOTPACK2_new.ItemArray = rowPOTPACK2.ItemArray
                    PACK_LIST_SHEET_NO_max += 1
                    rowPOTPACK2_new.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_max
                    rowPOTPACK2_new.Item("BARCODE_START") = DBNull.Value
                    rowPOTPACK2_new.Item("BARCODE_END") = DBNull.Value
                    dst.Tables("POTPACK2").Rows.Add(rowPOTPACK2_new)
                    'End If

                    For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}")
                        Dim rowPOTPACK3_new As DataRow = dst.Tables("POTPACK3").NewRow
                        rowPOTPACK3_new.ItemArray = rowPOTPACK3.ItemArray
                        rowPOTPACK3_new.Item("PACK_LIST_SHEET_NO") = PACK_LIST_SHEET_NO_max
                        dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_new)
                    Next
                Next

                Sort_grdColumns(grdPOTPACK2, "COLOR_CODE", True)

            Case "Add Line"
                Dim PACK_LIST_SHEET_NO As Integer = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")
                Dim PACK_LIST_SHEET_LNO As Integer = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_LNO").Value & "")
                Dim PACK_LIST_SHEET_LNO_max As Integer = dst.Tables("POTPACK3").Compute("MAX(PACK_LIST_SHEET_LNO)", $"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}")
                Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO, PACK_LIST_SHEET_LNO})
                Dim rowPOTPACK3_new As DataRow = dst.Tables("POTPACK3").NewRow
                rowPOTPACK3_new.ItemArray = rowPOTPACK3.ItemArray
                PACK_LIST_SHEET_LNO_max += 1
                rowPOTPACK3_new.Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_max
                rowPOTPACK3_new.Item("BARCODE_START") = DBNull.Value
                rowPOTPACK3_new.Item("BARCODE_END") = DBNull.Value
                dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_new)
                Sort_grdColumns(grdPOTPACK3, "STYLE_CODE,COLOR_CODE", True)

            Case "Add Lines"
                Dim PACK_LIST_SHEET_NO As Integer = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")
                Dim PACK_LIST_SHEET_LNO_max As Integer = dst.Tables("POTPACK3").Compute("MAX(PACK_LIST_SHEET_LNO)", $"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}")

                Dim STYLE_CODEs As New List(Of String)
                For Each row3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "PACK_LIST_SHEET_LNO")
                    Dim PACK_LIST_SHEET_LNO As Integer = Val(row3.Item("PACK_LIST_SHEET_LNO") & "")
                    Dim STYLE_CODE As String = row3.Item("STYLE_CODE") & ""
                    If Not STYLE_CODEs.Contains(STYLE_CODE) Then
                        Dim rowPOTPACK3 As DataRow = dst.Tables("POTPACK3").Rows.Find(New Object() {PACK_LIST_NO, PACK_LIST_SHEET_NO, PACK_LIST_SHEET_LNO})
                        Dim rowPOTPACK3_new As DataRow = dst.Tables("POTPACK3").NewRow
                        rowPOTPACK3_new.ItemArray = rowPOTPACK3.ItemArray
                        PACK_LIST_SHEET_LNO_max += 1
                        rowPOTPACK3_new.Item("PACK_LIST_SHEET_LNO") = PACK_LIST_SHEET_LNO_max
                        dst.Tables("POTPACK3").Rows.Add(rowPOTPACK3_new)
                    End If
                Next

                Sort_grdColumns(grdPOTPACK3, "STYLE_CODE,COLOR_CODE", True)

            Case "Copy Value to All Lines"

                If grd.ActiveRow Is Nothing Or grd.ActiveCell Is Nothing Then
                Else

                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                        grow.Cells(grd.ActiveCell.Column.Key).Value = grd.ActiveCell.Value
                        Update_CARTON_DIMENSIONS(grow)
                        grow.Update()
                    Next
                End If

                'Case "Item Status Inquiry"
                '    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Text
                '    Dim rowSPTAVEH1 As DataRow = LookUp("SPTAVEH1", VEND_CODE)
                '    If rowSPTAVEH1 IsNot Nothing Then
                '        Context_Launch("View", VEND_CODE, e.Tool.Key, "ICFSTAT1")
                '    End If

            Case "Copy Pattern to Remaining Lines"

                Dim SIZE_QTYs As New Dictionary(Of String, Integer)
                For Each row As DataRow In dst.Tables("POTPACK3").Select("", "PACK_LIST_SHEET_LNO")
                    Dim SIZE_CODE As String = row.Item("SIZE_CODE")
                    Dim CARTON_PACK As Integer = Val(row.Item("CARTON_PACK") & "")
                    If SIZE_CODE <> "" Then
                        If SIZE_QTYs.ContainsKey(SIZE_CODE) Then
                            ' row.Item("CARTON_PACK") = SIZE_QTYs(SIZE_CODE)
                        Else
                            SIZE_QTYs.Add(SIZE_CODE, CARTON_PACK)
                        End If
                    End If
                Next

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    Dim SIZE_CODE As String = grow.Cells("SIZE_CODE").Value
                    grow.Cells("CARTON_PACK").Value = SIZE_QTYs(SIZE_CODE)
                    grow.Update()
                Next

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "PACK_LIST_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If

        End Select

    End Sub

    Public Overrides Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_Leave(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "PO_REFERENCE"
                Absx1.txtFor("PO_REFERENCE").Text = Absx1.txtFor("PO_REFERENCE").Text.ToUpper

                Check_for_MultiPO()

            Case "PO_REFERENCE2"
                Absx1.txtFor("PO_REFERENCE2").Text = Absx1.txtFor("PO_REFERENCE2").Text.ToUpper

            Case "STYLE_CODE_PFX"
                Absx1.txtFor("STYLE_CODE_PFX").Text = Absx1.txtFor("STYLE_CODE_PFX").Text.ToUpper
        End Select
    End Sub

    Sub Check_for_MultiPO()

        Dim multi As Boolean = False
        Dim PO_REFERENCE As String = Absx1.txtFor("PO_REFERENCE").Text

        If PO_REFERENCE <> "" Then ' PO_REFERENCE.StartsWith("WM") Then

            ASCMAIN1.sql = "Select * from POTORDR1 where PO_REFERENCE = :PARM1"
            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "V", New String() {PO_REFERENCE})
            If row IsNot Nothing Then
                Dim PO_SPEC_ORDR_NO As String = row.Item("PO_SPEC_ORDR_NO") & ""
                If PO_SPEC_ORDR_NO.ToUpper.StartsWith("INITIAL") Then
                    Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
                    rowPOTPACKC = dst.Tables("POTPACKC").Rows.Find(CUST_CODE)
                    If rowPOTPACKC IsNot Nothing AndAlso rowPOTPACKC.Item("PACK_INITIAL_MULTI_PO") & "" = "1" Then
                        multi = True
                    End If
                End If
            End If
        End If

        lblPO2.Visible = multi
        txtPO_REFERENCE2.Visible = multi
        txtSTYLE_CODE_PFX2.Visible = multi And ScreenMode

        If multi Then
            spl.SplitterDistance = 100
        Else
            spl.SplitterDistance = 75
        End If

    End Sub
    Public Overrides Sub txt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
                Fill_Records("POTORDRR", VEND_CODE)
                Sort_grdColumns(grdPOTORDRR, "PO_DATE_SHIP_BY,PO_REFERENCE")

                'Case "PO_REFERENCE"
                '    Absx1.txtFor("PO_REFERENCE").Text = Absx1.txtFor("PO_REFERENCE").Text.ToUpper
                'Case "STYLE_CODE_PFX"
                '    Absx1.txtFor("STYLE_CODE_PFX").Text = Absx1.txtFor("STYLE_CODE_PFX").Text.ToUpper
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "PACK_LIST_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LC_AMT"
                If ScreenMode Then Display_Totals()
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            'Case "APPR_STATUS_CODE"
            '    If Absx1.optFor("APPR_STATUS_CODE").Value = "X" Then
            '        Absx1.optFor("STATUS_CODE").Value = "C"
            '    Else

            '    End If

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "LC_DATE"
            '    If Absx1.dteFor("LC_DATE").Value & "" = "" Then
            '        Absx1.txtFor("OPS_YYYYWW").Text = ""
            '    Else
            '        Dim DATE_START As Date = Absx1.dteFor("LC_DATE").Value
            '        If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
            '            ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(DATE_START, "dd-MMM-yyyy") & "'"
            '            Dim YW As String = ASCDATA1.GetDataValue
            '            If YW <> "" Then
            '                Absx1.txtFor("OPS_YYYYWW").Text = YW
            '            End If
            '        End If
            '    End If
        End Select
    End Sub
#End Region

#Region "grdPOTLTRCP"

#End Region

    Private Sub grdSPTSFOCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTPACKX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("PACK_LIST_NO").Text = e.Row.Cells("PACK_LIST_NO").Text
            Click_Command("View")
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        EnforceConstraints(False)
        If optShow.Value = "O" Then
            ASCMAIN1.sql = sqlPOTPACKX & " and PACK_LIST_STATUS = 'O'"
            Fill_Records("POTPACKX", "", True, ASCMAIN1.sql)
            grdPOTPACKX.Text = "Open"
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlPOTPACKX
            Fill_Records("POTPACKX", "", True, ASCMAIN1.sql)
            grdPOTPACKX.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdPOTPACKX, "PACK_LIST_NO".ToLower)

    End Sub

    Private Sub optShow_ValueChanged(sender As Object, e As EventArgs) Handles optShow.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Refresh_Documents()
    End Sub

    Private Sub optSTATUS_CODE_ValueChanged(sender As Object, e As EventArgs)
        If ScreenMode Then
            Synch_TABLE_NAME("POTPACK1")
            Display_Totals()
        End If
    End Sub

    Sub Display_Totals()
        'Dim LC_OPEN_CALC As Decimal = 0
        'Dim LC_CANC_CALC As Decimal = 0
        'Dim LC_AMT As Decimal = Val(Absx1.numFor("LC_AMT").Value & "")
        'Dim LC_PMTS As Decimal = Val(Absx1.numFor("LC_PMTS").Value & "")
        'If optSTATUS_CODE.Value = "O" Then
        '    LC_OPEN_CALC = LC_AMT - LC_PMTS
        '    LC_CANC_CALC = 0
        'Else
        '    LC_CANC_CALC = LC_AMT - LC_PMTS
        '    LC_OPEN_CALC = 0
        'End If

        'rowPOTPACK1.Item("LC_OPEN_CALC") = LC_OPEN_CALC
        'rowPOTPACK1.Item("LC_CANC_CALC") = LC_CANC_CALC

        Display_Totals_PO()
    End Sub

    Private Sub grdPOTLTRCP_AfterRowUpdate(sender As Object, e As RowEventArgs)
        Display_Totals_PO()
    End Sub

    Sub Display_Totals_PO()

        'Dim LC_PO As Decimal =
        '    Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_OPN)", "SEL='1'") & "") +
        '    Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_SHP)", "SEL='1'") & "")
        ''Val(dst.Tables("POTLTRCP").Compute("SUM(PO_AMT_REC)", "SEL='1'") & "")

        'rowPOTPACK1.Item("LC_PO") = LC_PO

        'Dim LC_AMT As Decimal = Val(Absx1.numFor("LC_AMT").Value & "")

        'If LC_PO > LC_AMT Then
        '    Absx1.numFor("LC_PO").Appearance.ForeColor = Drawing.Color.Red
        'Else
        '    Absx1.numFor("LC_PO").Appearance.ForeColor = Drawing.Color.Empty
        'End If
    End Sub

    Private Sub grdPOTORDRR_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTORDRR.AfterRowActivate
        Setup_grdPOTORDRD()
    End Sub

    Sub Setup_grdPOTORDRD()
        If grdPOTORDRR.ActiveRow Is Nothing OrElse Not grdPOTORDRR.ActiveRow.IsDataRow Then
            grdPOTORDRD.Visible = False
        Else
            Dim PO_ORDER_NO As String = grdPOTORDRR.ActiveRow.Cells("PO_ORDER_NO").Value
            Fill_Records("POTORDRD", PO_ORDER_NO)
            Sort_grdColumns(grdPOTORDRD, "STYLE_CODE, COLOR_CODE")
            grdPOTORDRD.Visible = True
        End If

    End Sub

    Private Sub grdPOTORDRR_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTORDRR.DoubleClickRow
        If grdPOTORDRR.ActiveRow IsNot Nothing AndAlso grdPOTORDRR.ActiveRow.IsDataRow Then

            Absx1.txtFor("PO_REFERENCE").Text = grdPOTORDRR.ActiveRow.Cells("PO_REFERENCE").Text
            Absx1.txtFor("PACK_LIST_DESC").Text = grdPOTORDRR.ActiveRow.Cells("PO_SPEC_ORDR_NO").Text

            Check_for_MultiPO()

        End If
    End Sub

    Private Sub grdPOTPACK2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTPACK2.AfterRowActivate
        Setup_grdPOTPACK3()
    End Sub

    Sub Setup_grdPOTPACK3()
        If grdPOTPACK2.ActiveRow Is Nothing OrElse Not grdPOTPACK2.ActiveRow.IsDataRow Then
            grdPOTPACK3.Visible = False
        Else
            Dim PACK_LIST_SHEET_NO As Int32 = Val(grdPOTPACK2.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")

            Dim dvw As DataView = DirectCast(grdPOTPACK3.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PACK_LIST_SHEET_NO = " & CStr(PACK_LIST_SHEET_NO)

            'Fill_Records("POTPACK3", New Object() {"", 0, 0})

            'Sort_grdColumns(grdPOTPACK3, "PACK_LIST_SHEET_LNO")
            Sort_grdColumns(grdPOTPACK3, "STYLE_CODE,COLOR_CODE", True)
            grdPOTPACK3.Visible = True

            grdPOTPACK3.Text = "Packing List Sheet Contents for Sheet " & grdPOTPACK2.ActiveRow.Cells("PACK_LIST_SHEET_NAME").Value
        End If
    End Sub

    Private Sub grdPOTPACK3_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTPACK3.InitializeRow
        Dim CARTON_GRS_WGT As Decimal = Val(e.Row.Cells("CARTON_GRS_WGT").Value & "")
        Dim CARTON_NET_WGT As Decimal = Val(e.Row.Cells.Item("CARTON_NET_WGT").Value & "")

        Dim PKG_CODE As String = e.Row.Cells.Item("PKG_CODE").Value & ""
        Dim CARTON_DIMENSIONS As String = e.Row.Cells.Item("CARTON_DIMENSIONS").Value & ""


        With e.Row.Cells("CARTON_COUNT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Count must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        'With e.Row.Cells("STYLE_WEIGHT")
        '    If Val(.Value & "") < 0 Then
        '        .ToolTipText = "Carton Pack must be > 0"
        '        .Appearance = Appearance_Red
        '    Else
        '        .ToolTipText = ""
        '        .Appearance = Nothing
        '    End If
        'End With

        With e.Row.Cells("CARTON_PACK")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Pack must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_DIMENSIONS")
            If .Value & "" = "" Then
                .ToolTipText = "Carton Dimensions are Mandatory"
                .Appearance = Appearance_Red
            Else
                Dim CARTON_VOLUME As Decimal = Get_Volume_from_Dims(CARTON_DIMENSIONS)
                If CARTON_VOLUME <= 0 Then
                    .ToolTipText = "Carton Dimensions must be expressed as: " & Replace("L' x W' x H'", "'", Chr(34))
                    .Appearance = Appearance_Red
                Else
                    .ToolTipText = ""
                    .Appearance = Nothing
                End If

            End If
        End With

        With e.Row.Cells("CARTON_GRS_WGT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Gross Weight must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_NET_WGT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Net Weight must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_GRS_WGT")
            If CARTON_GRS_WGT > 0 And CARTON_GRS_WGT < CARTON_NET_WGT Then
                .ToolTipText = "Carton Gross Weight must be > Net Weight"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

    End Sub

    Function Get_Volume_from_Dims(CARTON_DIMENSIONS As String) As Decimal
        Dim CARTON_VOLUME As Decimal = 0
        If CARTON_DIMENSIONS <> "" Then
            Dim D() As String = Split(Replace(CARTON_DIMENSIONS, Chr(34), "").ToUpper, "X")
            If D.Length > 0 Then
                For I As Integer = 1 To D.Length
                    If Val(D(I - 1)) <> 0 Then
                        If CARTON_VOLUME = 0 Then CARTON_VOLUME = 1
                        CARTON_VOLUME *= Val(D(I - 1))
                    End If
                Next
                If D.Length <> 3 Then CARTON_VOLUME = 0
            End If
        End If
        Return CARTON_VOLUME
    End Function

    Sub Export_XLS()

        Generate_Carton_Nos()

        Dim VBKG_NO As String = "000001"

        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        workbook = Produce_XLS(Me, VBKG_NO)

        Dim XLS_FILENAME_base As String = "Packing List " & PACK_LIST_NO & " for Booking " & VBKG_NO
        Dim XLS_FILENAME As String = XLS_FILENAME_base & ".xlsx"
        Dim retryCount As Integer = 0
        Do Until retryCount = -1 Or retryCount > 5
            If retryCount > 0 Then
                XLS_FILENAME = XLS_FILENAME_base & "_" & CStr(retryCount) & ".xlsx"
            End If
            Try
                workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                workbook.Close()
                retryCount = -1
            Catch ex As Exception
                retryCount += 1
                If retryCount > 5 Then
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Failed to Save Workbook")
                End If
            End Try
        Loop

        If retryCount = -1 Then
            Show_Document(XLS_FILENAME)
        End If
    End Sub


    Public Function Produce_XLS(frmASFBASE0 As ASFBASE0, VAN_REF As String) As SpreadsheetGear.IWorkbook

        Dim workbook As SpreadsheetGear.IWorkbook
        Dim worksheet As SpreadsheetGear.IWorksheet
        Dim worksheetBase As SpreadsheetGear.IWorksheet

        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePasteTo As SpreadsheetGear.IRange = Nothing

        Dim FILENAME_source As String = "R:\VDI\Templates" & "\" & "Template.xlsx"
        If ASCMAIN1.Running_in_VS Then FILENAME_source = "C:\Share\VDI\Templates\Template.xlsx"
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & "\" & "Template.xlsx"

        My.Computer.FileSystem.CopyFile(FILENAME_source, FILENAME, True)

        workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        worksheetBase = workbook.Worksheets(0)

        Dim ETD As Date = CDate("03/04/2021")
        Dim ETA As Date = CDate("05/22/2021")
        Dim INV_NO As String = "ILBD/YK/132/2021"
        Dim INV_DATE As Date = Now.Date
        Dim COUNTRY As String = "BANGLADESH"
        Dim SHIP_BY As String = "SEA"
        Dim PORT_DESC_ORIG As String = "CHITTAGONG,BANGLADESH"
        Dim PORT_DESC_DEST As String = "MAHER TERMINAL,U.S.A."

        Dim CONTAINER_NO As String = "INTEX009/2021"
        Dim EXP_NO As String = "2656 001589 2021"
        Dim ETD_CTG As String = "ETD_CTG"
        Dim BOL_NO As String = "BOL_NO"

        For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("", "PACK_LIST_SHEET_NO")
            'worksheet = workbook.Worksheets.Add
            worksheet = worksheetBase.CopyAfter(worksheetBase)
            worksheet.Name = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME")

            worksheet.Cells(4, 16).Value = INV_NO
            worksheet.Cells(5, 16).Value = INV_DATE
            worksheet.Cells(6, 16).Value = COUNTRY
            worksheet.Cells(7, 16).Value = SHIP_BY
            worksheet.Cells(8, 16).Value = PORT_DESC_ORIG
            worksheet.Cells(9, 16).Value = PORT_DESC_DEST

            Dim CX As Integer = 0

            CX = 13
            worksheet.Cells(4, 13).Value = "'" & Format(ETD, "MM/dd/yyyy")
            worksheet.Cells(5, 13).Value = "'" & Format(ETA, "MM/dd/yyyy")

            worksheet.Cells(4, 9).Value = CONTAINER_NO
            worksheet.Cells(5, 9).Value = EXP_NO
            worksheet.Cells(6, 9).Value = STYLE_CODE_PFX & IIf(STYLE_CODE_PFX2 = "", "", " & " & STYLE_CODE_PFX2)
            worksheet.Cells(7, 9).Value = PO_REFERENCE & IIf(PO_REFERENCE2 = "", "", " & " & PO_REFERENCE2)
            worksheet.Cells(8, 9).Value = ETD_CTG
            worksheet.Cells(9, 9).Value = BOL_NO

            Dim RX As Integer = 0

            Dim COLOR_CODE As String = rowPOTPACK2.Item("COLOR_CODE")
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            Dim COLOR_DESC_and_CODE As String = rowICTCOLR1.Item("COLOR_DESC") & " (" & COLOR_CODE & ")"
            worksheet.Cells(15, 5).Value = COLOR_DESC_and_CODE

            Dim PACK_LIST_DETAILS As String = rowPOTPACK2.Item("PACK_LIST_DETAILS") & ""
            worksheet.Cells(22, 0).Value = PACK_LIST_DETAILS
            'worksheet.Cells(22, 0).WrapText = False

            Dim PACK_LIST_SHEET_NO As Integer = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")

            For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3") _
                .Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO") ' rowPOTPACK2.GetChildRows("POTPACK2_POTPACK3")

                Dim PACK_LIST_SHEET_LNO As Integer = Val(rowPOTPACK3.Item("PACK_LIST_SHEET_LNO") & "")

                If RX > 0 Then
                    worksheet.Cells(15 + RX, 0).EntireRow.Insert()
                    worksheet.Cells(15 + RX + 1, 0).EntireRow.Copy(worksheet.Cells(15 + RX, 0).EntireRow)

                End If

                Dim STYLE_CODE As String = rowPOTPACK3.Item("STYLE_CODE") & ""
                Dim SIZE_CODE As String = rowPOTPACK3.Item("SIZE_CODE") & ""
                Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                Dim CARTON_PACK As Int32 = Val(rowPOTPACK3.Item("CARTON_PACK") & "")
                Dim CARTON_NO_START As Int32 = Val(rowPOTPACK3.Item("CARTON_NO_START") & "")
                Dim CARTON_NO_END As Int32 = Val(rowPOTPACK3.Item("CARTON_NO_END") & "")

                Dim CARTON_GRS_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_GRS_WGT") & "")
                Dim CARTON_NET_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_NET_WGT") & "")

                Dim CARTON_ID As Int32 = Val(rowPOTPACK3.Item("CARTON_ID") & "")
                Dim CARTON_DIMENSIONS As String = rowPOTPACK3.Item("CARTON_DIMENSIONS") & ""
                Dim BARCODE_START As String = rowPOTPACK3.Item("BARCODE_START") & ""
                Dim BARCODE_END As String = rowPOTPACK3.Item("BARCODE_END") & ""

                Dim SQLBC As String = $"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)} and PACK_LIST_SHEET_LNO = {CStr(PACK_LIST_SHEET_LNO)}"

                Dim BARCODE_MIN As String = dst.Tables("POTLPNL1").Compute("MIN(BARCODE)", SQLBC)
                Dim BARCODE_MAX As String = dst.Tables("POTLPNL1").Compute("MAX(BARCODE)", SQLBC)


                If BARCODE_START = BARCODE_MIN And BARCODE_END = BARCODE_MAX Then
                Else
                    Dim BARCODE_LAST As String = ""
                    Dim BARCODEs As String = ""
                    Dim BARCODE_append As String = ""
                    Dim BARCODE_FIRST As String = ""
                    Dim BARCODE_count As Integer = 0

                    For Each rowPOTLPNL1 As DataRow In dst.Tables("POTLPNL1").Select(SQLBC, "BARCODE")
                        Dim BARCODE As String = rowPOTLPNL1.Item("BARCODE")
                        If BARCODE_LAST = "" Then
                            BARCODE_FIRST = BARCODE
                            BARCODEs = BARCODE
                        Else
                            If Val(Mid(BARCODE, 2)) = Val(Mid(BARCODE_LAST, 2) + 1) Then
                                BARCODE_append = "-" & BARCODE
                            Else
                                If BARCODE_append <> "" Then BARCODEs &= BARCODE_append
                                BARCODE_append = ""
                                BARCODEs &= "," & BARCODE
                            End If
                        End If
                        BARCODE_LAST = BARCODE
                        BARCODE_count += 1
                    Next
                    If BARCODE_append <> "" Then BARCODEs &= BARCODE_append

                    BARCODE_START = BARCODEs
                    BARCODE_END = ""

                    If BARCODEs.Length = (BARCODE_FIRST & "-" & BARCODE_LAST).Length And BARCODE_count = CARTON_COUNT Then
                        BARCODE_START = BARCODE_FIRST
                        BARCODE_END = BARCODE_LAST

                    End If
                End If

                worksheet.Cells(15 + RX, 0).Value = CARTON_NO_START
                '  worksheet.Cells(15 + RX, 2).Value = CARTON_NO_END

                worksheet.Cells(15 + RX, 3).Value = STYLE_CODE
                worksheet.Cells(15 + RX, 4).Value = PO_REFERENCE

                ' STYLE DESC
                ' STYLE WEIGHT

                worksheet.Cells(15 + RX, 6).Value = SIZE_CODE
                worksheet.Cells(15 + RX, 7).Value = CARTON_COUNT
                worksheet.Cells(15 + RX, 8).Value = CARTON_PACK

                worksheet.Cells(15 + RX, 13).Value = CARTON_GRS_WGT
                worksheet.Cells(15 + RX, 14).Value = CARTON_NET_WGT

                worksheet.Cells(15 + RX, 15).Value = CARTON_DIMENSIONS
                worksheet.Cells(15 + RX, 16).Value = BARCODE_START
                If BARCODE_START.Contains(",") Or BARCODE_START.Contains("-") Then
                    worksheet.Cells(15 + RX, 16).HorizontalAlignment = SpreadsheetGear.HAlign.Left
                End If
                worksheet.Cells(15 + RX, 17).Value = BARCODE_END
                RX += 1
            Next

            worksheet.Cells(15 + RX, 0).EntireRow.Delete()

            With worksheet.Cells(15, 5, 15 + RX - 1, 5)
                .Merge()
            End With



            With worksheet.PageSetup
                .FitToPagesTall = 1
                .FitToPagesWide = 1
                .FitToPages = True
                .Orientation = SpreadsheetGear.PageOrientation.Landscape
            End With
        Next

        worksheetBase.Delete()


        Return workbook

    End Function

    Function Generate_Carton_Nos()

        Dim EMsg As String = ""

        Dim CARTONs As New List(Of Integer)
        For Each rowPOTPACK2 As DataRow In dst.Tables("POTPACK2").Select("")
            Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK2.Item("PACK_LIST_SHEET_NAME") & ""
            Dim PACK_LIST_SHEET_NO As Int32 = Val(rowPOTPACK2.Item("PACK_LIST_SHEET_NO") & "")
            Dim CARTON_NO_START As Int32 = Val(rowPOTPACK2.Item("CARTON_NO_START") & "")

            If CARTON_NO_START <= 0 Then
                EMsg &= vbCr & "Invalid Starting Carton No on Sheet " & PACK_LIST_SHEET_NAME
            Else
                If INITIAL_ORDER = "1" Then

                    Dim CARTON_GRS_WGT As Decimal = Val(rowPOTPACK2.Item("CARTON_GRS_WGT") & "")
                    Dim CARTON_NET_WGT As Decimal = Val(rowPOTPACK2.Item("CARTON_NET_WGT") & "")
                    Dim CARTON_COUNT As Int32 = Val(rowPOTPACK2.Item("CARTON_COUNT") & "")
                    Dim CARTON_PACK As Int32 = Val(rowPOTPACK2.Item("CARTON_PACK") & "")
                    Dim CARTON_DIMENSIONS As String = rowPOTPACK2.Item("CARTON_DIMENSIONS") & ""

                    If CARTON_COUNT <= 0 Or CARTON_PACK <= 0 Or CARTON_DIMENSIONS = "" Or Get_Volume_from_Dims(CARTON_DIMENSIONS) <= 0 Or CARTON_GRS_WGT <= 0 Or CARTON_NET_WGT <= 0 Or CARTON_GRS_WGT < CARTON_NET_WGT Then
                        EMsg &= vbCr & "Issue with Data on Sheet " & PACK_LIST_SHEET_NAME
                    Else
                        ' only required if there are multiple sheets
                    End If

                Else
                    Dim CARTON_NO As Int32 = CARTON_NO_START
                    For Each rowPOTPACK3 As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}", "STYLE_CODE, PACK_LIST_SHEET_LNO") ' In rowPOTPACK2.GetChildRows("POTPACK2_POTPACK3")
                        ' Dim PACK_LIST_SHEET_NAME As String = rowPOTPACK3.GetParentRow("POTPACK2_POTPACK3").Item("PACK_LIST_SHEET_NAME")
                        Dim CARTON_GRS_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_GRS_WGT") & "")
                        Dim CARTON_NET_WGT As Decimal = Val(rowPOTPACK3.Item("CARTON_NET_WGT") & "")
                        Dim CARTON_COUNT As Int32 = Val(rowPOTPACK3.Item("CARTON_COUNT") & "")
                        Dim CARTON_PACK As Int32 = Val(rowPOTPACK3.Item("CARTON_PACK") & "")
                        Dim CARTON_ID As Int32 = Val(rowPOTPACK3.Item("CARTON_ID") & "")
                        Dim CARTON_DIMENSIONS As String = rowPOTPACK3.Item("CARTON_DIMENSIONS") & ""

                        'Dim STYLE_WEIGHT As Decimal = Val(rowPOTPACK3.Item("STYLE_WEIGHT") & "")
                        'If STYLE_WEIGHT < 0 Then
                        '    EMsg &= vbCr & "Style Weight must not be negative - see Sheet " & PACK_LIST_SHEET_NAME
                        'End If

                        rowPOTPACK3.Item("CARTON_NO_START") = CARTON_NO

                        If CARTON_COUNT <= 0 Or CARTON_PACK <= 0 Or CARTON_DIMENSIONS = "" Or Get_Volume_from_Dims(CARTON_DIMENSIONS) <= 0 Or CARTON_GRS_WGT <= 0 Or CARTON_NET_WGT <= 0 Or CARTON_GRS_WGT < CARTON_NET_WGT Then
                            EMsg &= vbCr & "Issue with Data on Sheet " & PACK_LIST_SHEET_NAME
                            Exit For
                        Else
                            If CARTON_COUNT > 0 Then
                                Dim overlapping As Boolean = False
                                For I As Integer = CARTON_NO To CARTON_NO + CARTON_COUNT - 1
                                    If CARTONs.Contains(I) Then
                                        overlapping = True
                                    Else
                                        CARTONs.Add(I)
                                    End If
                                Next

                                If overlapping Then
                                    EMsg &= vbCr & "Overlapping Carton Nos on Sheet " & PACK_LIST_SHEET_NAME
                                    Exit For
                                Else
                                    CARTON_NO += CARTON_COUNT
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next

        Return EMsg

    End Function

    Sub Generate_LPN_Report_File()

        dst.Tables("POTLPNL1").Rows.Clear()

        Dim tblName As String = "POTPACK3"
        If INITIAL_ORDER = "1" Then tblName = "POTPACK2"

        For Each row As DataRow In dst.Tables(tblName).Select("ISNULL(BARCODE_START,'') <> ''")
            Dim CARTON_COUNT As Integer = Val(row.Item("CARTON_COUNT") & "")
            Dim BARCODE_START As String = row.Item("BARCODE_START") & ""
            Dim BARCODE_START_NO = Val(Mid(BARCODE_START, 2))
            Dim PACK_LIST_DESC As String = rowPOTPACK1.Item("PACK_LIST_DESC")
            For C As Integer = 1 To CARTON_COUNT
                Dim BARCODE_NO As Integer = BARCODE_START_NO + C - 1
                Dim BARCODE As String = Mid(BARCODE_START, 1, 1) & Format(BARCODE_NO, "0000000")

                Dim rowPOTLPNL1 As DataRow = dst.Tables("POTLPNL1").NewRow
                With rowPOTLPNL1
                    .Item("BARCODE") = BARCODE
                    .Item("PO_REFERENCE") = PO_REFERENCE
                    .Item("STYLE_CODE_PFX") = STYLE_CODE_PFX
                    If INITIAL_ORDER = "1" Then
                        .Item("PO_REFERENCE2") = PO_REFERENCE2
                        .Item("STYLE_CODE_PFX2") = STYLE_CODE_PFX2
                        If rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" Then
                            .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                        End If
                    Else
                        .Item("STYLE_CODE") = row.Item("STYLE_CODE")
                        .Item("COLOR_CODE") = row.Item("COLOR_CODE")
                        .Item("SIZE_CODE") = row.Item("SIZE_CODE")
                    End If

                    .Item("PACK_LIST_DESC") = PACK_LIST_DESC
                    .Item("PACK_LIST_NO") = row.Item("PACK_LIST_NO")
                    .Item("PACK_LIST_SHEET_NO") = row.Item("PACK_LIST_SHEET_NO")
                    If INITIAL_ORDER = "1" Then
                    Else
                        .Item("PACK_LIST_SHEET_LNO") = row.Item("PACK_LIST_SHEET_LNO")
                    End If

                    .Item("BARCODE_STATUS") = "A"
                    .Item("CARTON_PACK") = row.Item("CARTON_PACK")

                    If INITIAL_ORDER = "1" Then
                    Else
                        .Item("CARTON_ID") = row.Item("CARTON_ID")
                    End If

                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP

                    .Item("CARTON_GRS_WGT") = row.Item("CARTON_GRS_WGT")
                    .Item("CARTON_NET_WGT") = row.Item("CARTON_NET_WGT")
                    .Item("CARTON_DIMENSIONS") = row.Item("CARTON_DIMENSIONS")
                    .Item("PKG_CODE") = row.Item("PKG_CODE")
                End With
                dst.Tables("POTLPNL1").Rows.Add(rowPOTLPNL1)
            Next
        Next

        BeginTrans()

        ASCMAIN1.sql = $"Update POTLPNL1 Set BARCODE_STATUS = 'D' where PACK_LIST_NO = '{PACK_LIST_NO}'"
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("POTLPNL1")

        CommitTrans()

    End Sub

    Sub Print_Labels()
        Print_Report_Begin()
        CR_params.Add("SUBT", "")


        ' Dim printerName As String = "ZDesigner ZT411-300dpi ZPL (redirected 4)"
        If INITIAL_ORDER = "1" Then
            Generate_Report("PORLPNL2")
        Else
            Generate_Report("PORLPNL1")
        End If

        'Print_Report_End(,, printerName)
        Print_Report_End()
    End Sub

    Private Sub grdPOTPACK3_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTPACK3.InitializeLayout

    End Sub

    Private Sub grdPOTPACK3_AfterExitEditMode(sender As Object, e As EventArgs) Handles grdPOTPACK3.AfterExitEditMode
        If INITIAL_ORDER = "1" And rowPOTPACKC.Item("PACK_INITIAL_BY_COLOR") & "" = "1" Then ' Not PO_REFERENCE.StartsWith("WM") Then
            With grdPOTPACK2.ActiveRow
                Dim PACK_LIST_SHEET_NO As Integer = Val(.Cells("PACK_LIST_SHEET_NO").Value & "")
                Dim CARTON_PACK As Integer = Val(dst.Tables("POTPACK3").Compute("SUM(CARTON_PACK)", $"PACK_LIST_SHEET_NO = {PACK_LIST_SHEET_NO}"))
                .Cells("CARTON_PACK").Value = CARTON_PACK
            End With
        End If
    End Sub

    Private Sub grdPOTPACK3_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdPOTPACK3.AfterCellUpdate

        Select Case e.Cell.Column.Key
            'Case "STYLE_WEIGHT"
            '    Calculate_Net_Weight(e)
            'Case "CARTON_PACK"
            '    Calculate_Net_Weight(e)
        End Select
    End Sub

    'Sub Calculate_Net_Weight(e As CellEventArgs)

    '    If Not Me.IsLoading And ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
    '        Dim STYLE_WEIGHT As Decimal = Val(e.Cell.Row.Cells("STYLE_WEIGHT").Value & "")
    '        Dim CARTON_PACK As Integer = Val(e.Cell.Row.Cells("CARTON_PACK").Value & "")
    '        If STYLE_WEIGHT > 0 And CARTON_PACK > 0 Then
    '            e.Cell.Row.Cells("CARTON_NET_WGT").Value = STYLE_WEIGHT * CARTON_PACK
    '        End If
    '    End If
    'End Sub

    Sub Check_for_Overbooked()

        ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", SUM(PACKED_NOW) PACKED_NOW" & vbCrLf _
            & ", SUM(PACKED_OTHER) PACKED_OTHER" & vbCrLf _
            & ", SUM(PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", SUM(PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & ", SUM(POTPACK3.CARTON_COUNT * POTPACK3.CARTON_PACK) PACKED_NOW, 0 PACKED_OTHER, 0 PO_QTY_OPN, 0 PO_QTY_SHP" & vbCrLf _
            & " from POTPACK3, POTPACK1" & vbCrLf _
            & $" where POTPACK1.PACK_LIST_NO = POTPACK3.PACK_LIST_NO And POTPACK1.PO_ORDER_NO = '{PO_ORDER_NO}'" & vbCrLf _
            & $"   and POTPACK3.PACK_LIST_NO = '{PACK_LIST_NO}' and POTPACK1.PACK_LIST_STATUS = 'F'" & vbCrLf _
            & " group by POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & ", 0 PACKED_NOW, SUM (POTPACK3.CARTON_COUNT * POTPACK3.CARTON_PACK) PACKED_OTHER, 0 PO_QTY_OPN, 0 PO_QTY_SHP" & vbCrLf _
            & " from POTPACK3,POTPACK1" & vbCrLf _
            & $" where POTPACK1.PACK_LIST_NO = POTPACK3.PACK_LIST_NO AND POTPACK1.PO_ORDER_NO = '{PO_ORDER_NO}'" & vbCrLf _
            & $"   and POTPACK3.PACK_LIST_NO <> '{PACK_LIST_NO}' and POTPACK1.PACK_LIST_STATUS = 'F'" & vbCrLf _
            & " group by POTPACK3.STYLE_CODE, POTPACK3.COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ", 0 PACKED_NOW, 0 PACKED_OTHER, SUM (POTORDR2.PO_QTY_OPN) PO_QTY_OPN, SUM (POTORDR2.PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & " from POTORDR2" & vbCrLf _
            & $" where POTORDR2.PO_ORDER_NO = '{PO_ORDER_NO}'" & vbCrLf _
            & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ") group by STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ") where PACKED_NOW + PACKED_OTHER > PO_QTY_OPN + PO_QTY_SHP"

        Dim tbl As DataTable = ASCDATA1.GetDataTable
        If tbl.Rows.Count > 0 Then
            Using f As New ASFMSGBF
                f.Show_grd(tbl, Me, "Overbooked PO - Message to Don")
            End Using
        End If
    End Sub

    Private Sub grdPOTPACK2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdPOTPACK2.KeyPress

    End Sub

    Private Sub grdPOTPACK2_CellChange(sender As Object, e As CellEventArgs) Handles grdPOTPACK2.CellChange

        If e.Cell.Column.Key = "PACK_LIST_DETAILS" Then
            Dim PACK_LIST_DETAILS As String = e.Cell.Text & ""
            If PACK_LIST_DETAILS <> "" Then

                Dim LINES As Integer = PACK_LIST_DETAILS.Count(Function(c As Char) c = vbCr)

                grdPOTPACK2.ActiveRow.Height = 17 * (LINES + 1)
            End If

        End If
    End Sub

    Private Sub grdPOTPACK2_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdPOTPACK2.AfterRowUpdate
        e.Row.PerformAutoSize()
        'If (INITIAL_ORDER = "1") Then
        '    ' MAYBE THIS IS FOR WALMART ONLY?
        '    Dim CARTON_COUNT As Integer = Val(e.Row.Cells("CARTON_COUNT").Value & "")
        '    Dim PACK_LIST_SHEET_NO As Integer = Val(e.Row.Cells("PACK_LIST_SHEET_NO").Value & "")
        '    For Each row As DataRow In dst.Tables("POTPACK3").Select($"PACK_LIST_SHEET_NO = {CStr(PACK_LIST_SHEET_NO)}")
        '        Dim PO_QTY_OPN As Integer = Val(row.Item("PO_QTY_OPN") & "")
        '        row.Item("CARTON_PACK") = PO_QTY_OPN / CARTON_COUNT
        '    Next
        'End If
    End Sub

    Private Sub grdPOTPACK2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdPOTPACK2.AfterCellUpdate
        If e.Cell.Column.Key = "PACK_LIST_DETAILS" Then
            e.Cell.Row.PerformAutoSize()
        End If
    End Sub

    Private Sub grdPOTPACK3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTPACK3.AfterRowActivate
        Setup_grdPOTLPNL1()
    End Sub

    Sub Setup_grdPOTLPNL1()
        If INITIAL_ORDER = "1" Or EntryMode = "L" Then
            Exit Sub
        End If
        If grdPOTPACK3.ActiveRow Is Nothing OrElse Not grdPOTPACK3.ActiveRow.IsDataRow Then
            grdPOTLPNLX.Visible = False
        Else
            Dim PACK_LIST_SHEET_NO As Int32 = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_NO").Value & "")
            Dim PACK_LIST_SHEET_LNO As Int32 = Val(grdPOTPACK3.ActiveRow.Cells("PACK_LIST_SHEET_LNO").Value & "")
            Dim dvw As DataView = DirectCast(grdPOTLPNLX.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"PACK_LIST_SHEET_NO = {PACK_LIST_SHEET_NO} and PACK_LIST_SHEET_LNO = {PACK_LIST_SHEET_LNO}"
            Sort_grdColumns(grdPOTLPNLX, "BARCODE", True)
            grdPOTLPNLX.Visible = True

            grdPOTLPNLX.Text = $"LPNs Line {PACK_LIST_SHEET_LNO}"
        End If
    End Sub

    Private Sub grdPOTPACK2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTPACK2.InitializeRow
        If INITIAL_ORDER = "1" Then
        Else
            Exit Sub
        End If

        Dim CARTON_GRS_WGT As Decimal = Val(e.Row.Cells("CARTON_GRS_WGT").Value & "")
        Dim CARTON_NET_WGT As Decimal = Val(e.Row.Cells.Item("CARTON_NET_WGT").Value & "")

        Dim CARTON_DIMENSIONS As String = e.Row.Cells.Item("CARTON_DIMENSIONS").Value & ""

        With e.Row.Cells("CARTON_COUNT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Count must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_PACK")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Pack must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_DIMENSIONS")
            If .Value & "" = "" Then
                .ToolTipText = "Carton Dimensions are Mandatory"
                .Appearance = Appearance_Red
            Else
                Dim CARTON_VOLUME As Decimal = Get_Volume_from_Dims(CARTON_DIMENSIONS)
                If CARTON_VOLUME <= 0 Then
                    .ToolTipText = "Carton Dimensions must be expressed as: " & Replace("L' x W' x H'", "'", Chr(34))
                    .Appearance = Appearance_Red
                Else
                    .ToolTipText = ""
                    .Appearance = Nothing
                End If

            End If
        End With

        With e.Row.Cells("CARTON_GRS_WGT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Gross Weight must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_NET_WGT")
            If Val(.Value & "") <= 0 Then
                .ToolTipText = "Carton Net Weight must be > 0"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With

        With e.Row.Cells("CARTON_GRS_WGT")
            If CARTON_GRS_WGT > 0 And CARTON_GRS_WGT < CARTON_NET_WGT Then
                .ToolTipText = "Carton Gross Weight must be > Net Weight"
                .Appearance = Appearance_Red
            Else
                .ToolTipText = ""
                .Appearance = Nothing
            End If
        End With
    End Sub

    Private Sub txtLPN_KeyDown(sender As Object, e As KeyEventArgs) Handles txtLPN.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Confirm_BARCODE
        End If
    End Sub

    Sub Confirm_BARCODE()
        If txtLPN.Text <> txtLPN.Text.ToUpper Then
            txtLPN.Text = txtLPN.Text.ToUpper
        End If
        Dim BARCODE As String = txtLPN.Text
        Dim rowPOTLPNL1 As DataRow = dst.Tables("POTLPNL1").Rows.Find(BARCODE)

        If rowPOTLPNL1 Is Nothing Then
            lblLastScan.Text = "Invalid Scan: " & BARCODE
            lblLastScan.Appearance.ForeColor = Drawing.Color.Red
            lblLastScan.Visible = True
        Else
            Dim SHIP_CONF As String = rowPOTLPNL1.Item("SHIP_CONF") & ""
            rowPOTLPNL1.Item("SHIP_CONF") = optConfirmTo.Value

            lblLastScan.Text = $"{BARCODE} Confirmed To " & optConfirmTo.Text
            lblLastScan.Appearance.ForeColor = Drawing.Color.Empty
            lblLastScan.Visible = True
        End If

        txtLPN.Text = ""
        txtLPN.Focus()
    End Sub
    Private Sub grdPOTLPNL1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTLPNL1.InitializeRow

        Dim CONF As String = ""

        With e.Row.Cells("CONF_SHP")
            CONF = .Value & ""
            If CONF = "1" Then
                .Appearance.BackColor = Drawing.Color.DarkGreen
                .Appearance.ForeColor = Drawing.Color.White
            Else
                .Appearance.BackColor = Drawing.Color.Empty
                .Appearance.ForeColor = Drawing.Color.Empty
            End If
        End With

        With e.Row.Cells("CONF_REM")
            CONF = .Value & ""
            If CONF = "1" Then
                .Appearance.BackColor = Drawing.Color.Red
                .Appearance.ForeColor = Drawing.Color.White
            Else
                .Appearance.BackColor = Drawing.Color.Empty
                .Appearance.ForeColor = Drawing.Color.Empty
            End If
        End With

        With e.Row.Cells("CONF_UNK")
            CONF = .Value & ""
            If CONF = "1" Then
                .Appearance.BackColor = Drawing.Color.Yellow
                '.Appearance.ForeColor = Drawing.Color.White
            Else
                .Appearance.BackColor = Drawing.Color.Empty
                '.Appearance.ForeColor = Drawing.Color.Empty
            End If
        End With
    End Sub

    Private Sub cmdSetUnknown_Click(sender As Object, e As EventArgs) Handles cmdSetUnknown.Click

        For Each row As DataRow In dst.Tables("POTLPNL1").Select("SHIP_CONF IS NULL")
            row.Item("SHIP_CONF") = optConfirmTo.Value
        Next
    End Sub

    Private Sub optConfirmTo_ValueChanged(sender As Object, e As EventArgs) Handles optConfirmTo.ValueChanged
        cmdSetUnknown.Text = "Set All Unknown to " & optConfirmTo.Text

        If optConfirmTo.Value = "S" Then
            cmdSetUnknown.Appearance.ForeColor = Drawing.Color.DarkGreen
            'cmdSetUnknown.Appearance.BackColor = Drawing.Color.DarkGreen
        Else
            cmdSetUnknown.Appearance.ForeColor = Drawing.Color.Red
            'cmdSetUnknown.Appearance.BackColor = Drawing.Color.Red
        End If

        cmdSetAll2Unknown.Appearance.ForeColor = Drawing.Color.Yellow
        'cmdSetAll2Unknown.Appearance.BackColor = Drawing.Color.Yellow
    End Sub


    Private Sub grdPOTLPNL1_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdPOTLPNL1.DoubleClickCell
        If grdPOTLPNL1.ActiveRow IsNot Nothing AndAlso grdPOTLPNL1.ActiveRow.IsDataRow Then
            Dim COL As String = grdPOTLPNL1.ActiveCell.Column.Key
            If COL = "BARCODE" Then
                txtLPN.Text = grdPOTLPNL1.ActiveCell.Value & ""
                Confirm_BARCODE()
            End If

        End If
    End Sub

    Private Sub grdPOTLPNL1_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdPOTLPNL1.ClickCell
        If grdPOTLPNL1.ActiveRow IsNot Nothing AndAlso grdPOTLPNL1.ActiveRow.IsDataRow And grdPOTLPNL1.ActiveCell IsNot Nothing Then
            Dim COL As String = grdPOTLPNL1.ActiveCell.Column.Key
            Dim BARCODE As String = grdPOTLPNL1.ActiveRow.Cells("BARCODE").Value & ""
            Dim rowPOTLPNL1 As DataRow = dst.Tables("POTLPNL1").Rows.Find(BARCODE)
            If COL = "CONF_SHP" Then
                optConfirmTo.Value = "S"
                txtLPN.Text = BARCODE
                Confirm_BARCODE()
                grdPOTLPNL1.ActiveCell = grdPOTLPNL1.ActiveRow.Cells("BARCODE")
            End If
            If COL = "CONF_REM" Then
                optConfirmTo.Value = "R"
                txtLPN.Text = BARCODE
                Confirm_BARCODE()
                grdPOTLPNL1.ActiveCell = grdPOTLPNL1.ActiveRow.Cells("BARCODE")
            End If
            If COL = "CONF_UNK" Then
                rowPOTLPNL1.Item("SHIP_CONF") = DBNull.Value
            End If
        End If
    End Sub

    Private Sub cmdSetAll2Unknown_Click(sender As Object, e As EventArgs) Handles cmdSetAll2Unknown.Click
        For Each row As DataRow In dst.Tables("POTLPNL1").Select("SHIP_CONF IS not NULL")
            row.Item("SHIP_CONF") = DBNull.Value
        Next
    End Sub

    Private Sub cmdAddPkg_Click(sender As Object, e As EventArgs) Handles cmdAddPkg.Click
        Dim PKG_L As Decimal = Val(numL.Value & "")
        Dim PKG_W As Decimal = Val(numW.Value & "")
        Dim PKG_H As Decimal = Val(numH.Value & "")

        Dim EMsg As String = ""
        If PKG_L < 1 Or PKG_L > 99.9 Then EMsg &= vbCrLf & "Invalid Value specified for Length"
        If PKG_W < 1 Or PKG_W > 99.9 Then EMsg &= vbCrLf & "Invalid Value specified for Width"
        If PKG_H < 1 Or PKG_H > 99.9 Then EMsg &= vbCrLf & "Invalid Value specified for Height"

        Dim CARTON_DIMENSIONS As String = ""
        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Add Carton with Dimensions Specified")
            Exit Sub
        Else
            Dim QUO As String = Chr(34)
            CARTON_DIMENSIONS = $"{CStr(PKG_L)}{QUO}x{CStr(PKG_W)}{QUO}x{CStr(PKG_H)}{QUO}"
            If dst.Tables("WHTPKGM1").Select($"CARTON_DIMENSIONS = '{CARTON_DIMENSIONS}'").Length > 0 Then
                MsgBox(EMsg, MsgBoxStyle.OkOnly, $"Duplicate Carton Dimensions: {CARTON_DIMENSIONS}")
                Exit Sub
            End If
            If MsgBox("OK to add Carton with Dimensions of:" & vbCrLf & vbCrLf & CARTON_DIMENSIONS, MsgBoxStyle.OkCancel, "Verification") = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If

        Dim CTL As String = ASCMAIN1.Next_Control_No($"WHTPKGM1_{BARCODE_PFX}.PKG_CODE")
        'Dim PKG_CODE_ctr As Integer = dst.Tables("WHTPKGM1").Select("PKG_CODE like '{BARCODE_PFX}*'").Length + 1
        Dim PKG_CODE_ctr As Integer = Val(CTL)

        Dim PKG_CODE As String = BARCODE_PFX & Format(PKG_CODE_ctr, "0000000")
        Dim rowWHTPKGM1 As DataRow = dst.Tables("WHTPKGM1").NewRow
        With rowWHTPKGM1
            .Item("PKG_CODE") = PKG_CODE
            .Item("PKG_DESC") = PKG_CODE
            .Item("PKG_L") = PKG_L
            .Item("PKG_W") = PKG_W
            .Item("PKG_H") = PKG_H
            .Item("BARCODE_PFX") = BARCODE_PFX
            .Item("PKG_STATUS") = "A"
            .Item("CARTON_DIMENSIONS") = CARTON_DIMENSIONS
        End With
        dst.Tables("WHTPKGM1").Rows.Add(rowWHTPKGM1)

        Reset_ValueLists()

        numL.Value = DBNull.Value
        numW.Value = DBNull.Value
        numH.Value = DBNull.Value
    End Sub

    Sub Reset_ValueLists()

        Dim VL2 As New ValueList
        Dim VL3 As New ValueList
        For Each ROW As DataRow In dst.Tables("WHTPKGM1").Select("", "CARTON_DIMENSIONS")
            Dim PKG_CODE As String = ROW.Item("PKG_CODE")
            Dim CARTON_DIMENSIONS As String = ROW.Item("CARTON_DIMENSIONS")
            VL2.ValueListItems.Add(PKG_CODE, CARTON_DIMENSIONS)
            VL3.ValueListItems.Add(PKG_CODE, CARTON_DIMENSIONS)
        Next

        grdPOTPACK2.DisplayLayout.Bands(0).Columns("PKG_CODE").ValueList = VL2
        grdPOTPACK3.DisplayLayout.Bands(0).Columns("PKG_CODE").ValueList = VL3

    End Sub

    Private Sub grdPOTPACK3_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdPOTPACK3.BeforeRowUpdate
        If grdPOTPACK3.ActiveRow IsNot Nothing Then
            Update_CARTON_DIMENSIONS(grdPOTPACK3.ActiveRow)
        End If
    End Sub

    Private Sub grdPOTPACK2_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdPOTPACK2.BeforeRowUpdate
        If grdPOTPACK2.ActiveRow IsNot Nothing Then
            Update_CARTON_DIMENSIONS(grdPOTPACK2.ActiveRow)
        End If
    End Sub

    Sub Update_CARTON_DIMENSIONS(grow As UltraWinGrid.UltraGridRow)
        If grow IsNot Nothing Then
            Dim PKG_CODE As String = grow.Cells("PKG_CODE").Value & ""
            Dim rowWHTPKGM1 As DataRow = dst.Tables("WHTPKGM1").Rows.Find(PKG_CODE)
            If PKG_CODE <> "" AndAlso rowWHTPKGM1 IsNot Nothing Then
                Dim CARTON_DIMENSIONS As String = rowWHTPKGM1.Item("CARTON_DIMENSIONS") & ""
                grow.Cells("CARTON_DIMENSIONS").Value = CARTON_DIMENSIONS
            Else
                grow.Cells("CARTON_DIMENSIONS").Value = ""
            End If
        End If
    End Sub

    Private Sub grdWHTPKGM1_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdWHTPKGM1.InitializeLayout

    End Sub

    Private Sub grdWHTPKGM1_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdWHTPKGM1.DoubleClickCell
        If EntryMode = "N" Or EntryMode = "E" Then
            If chkDblClickToEdit.Checked Then

                If grdPOTPACK3.ActiveRow IsNot Nothing AndAlso grdPOTPACK3.ActiveCell IsNot Nothing AndAlso grdPOTPACK3.ActiveCell.Column.Key = "PKG_CODE" Then
                grdPOTPACK3.ActiveCell.Value = e.Cell.Row.Cells("PKG_CODE").Value
                grdPOTPACK3.ActiveRow.Update()
                grdPOTPACK3.PerformAction(UltraGridAction.BelowCell)
            End If
                If grdPOTPACK2.ActiveRow IsNot Nothing AndAlso grdPOTPACK2.ActiveCell IsNot Nothing AndAlso grdPOTPACK2.ActiveCell.Column.Key = "PKG_CODE" Then
                    grdPOTPACK2.ActiveCell.Value = e.Cell.Row.Cells("PKG_CODE").Value
                    grdPOTPACK2.ActiveRow.Update()
                    grdPOTPACK2.PerformAction(UltraGridAction.BelowCell)
                End If

            End If
        End If
    End Sub
End Class