Imports Infragistics.Win.UltraWinGrid

Public Class ARTCUST1

    Private sqlARTSREP1 As String
    Private tblUPSReference As DataTable
    Private tblFDXReference As DataTable
    Private TAX_ID As String = ""
    Private TAX_ID_DOC As String = ""

    ' 02/27/2019
    ' ALTER table artcust1 add  CUST_SHIP_COMPLETE_DETAIL	 VARCHAR2(1);
    ' UPDATE ARTCUST1 SET CUST_SHIP_COMPLETE_DETAIL = '1' WHERE CUST_CODE = '310921';

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select ARTCUST2.* " _
                & " from ARTCUST2 " _
                & " where ARTCUST2.CUST_CODE = :PARM1 and ARTCUST2.CUST_ADDR_TYPE = 'MK'"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, True, "V", 3)

            ASCMAIN1.sql = "Select ARTCUSTD.* " _
                & " from ARTCUSTD " _
                & " where ARTCUSTD.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V", 2)


            sqlARTSREP1 = "Select ARTSREP1.*, SOTSDIV1.SALES_DIVISION_NAME, SOTSREP1.SREP_NAME" _
            & " from SOTSDIV1,SOTSREP1,ARTSREP1" _
            & " where SOTSDIV1.SALES_DIVISION_CODE = ARTSREP1.SALES_DIVISION_CODE" _
            & "   and SOTSREP1.SREP_CODE = ARTSREP1.SREP_CODE"
            ASCMAIN1.sql = sqlARTSREP1 _
            & "  and ARTSREP1.CUST_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ARTSREP1", "**", 0, True, "V", 3)

            Create_TDA(.Tables.Add, "ARTCUSTS", "*", 1)

            Create_TDA(.Tables.Add("SOTCARRS_FEDEX"), "SOTCARRS", "*", 2)
            Create_TDA(.Tables.Add("SOTCARRS_UPS"), "SOTCARRS", "*", 2)

            Create_TDA(.Tables.Add, "TATSHIPP", "*", 2)
            Create_TDA(.Tables.Add, "WBTCUST1", "*", 1)

            Create_TDA(.Tables.Add, "ARTCUSTM", "*", 1)
            .Tables("ARTCUSTM").Columns.Add("CUST_NAME", GetType(System.String))

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                Create_TDA(.Tables.Add, "ARTCUSTQ", "*", 1)
            End If
            If .Tables("ARTCUST1").Columns.Contains("PVT_LBL_CODE") = False Then
                .Tables("ARTCUST1").Columns.Add("PVT_LBL_CODE", GetType(System.String))
                .Tables("ARTCUST1").Columns.Add("PVT_LBL_DISC_PCT", GetType(System.Int64))
            End If
        End With

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")
        grdARTCUSTD.DataSource = dst.Tables("ARTCUSTD")
        grdARTSREP1.DataSource = dst.Tables("ARTSREP1")
        grdARTCUSTM.DataSource = dst.Tables("ARTCUSTM")

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            grdSOTCARRS_FEDEX.DataSource = dst.Tables("SOTCARRS_FEDEX")
            ASCMAIN1.Add_Value_List(grdSOTCARRS_FEDEX, "CARRIER_PROD_CODE", "SELECT CARRIER_PROD_CODE, CARRIER_PROD_DESC FROM SOTCARR2 WHERE CARRIER_CODE = 'FEDEX'
                                                                            UNION
                                                                         SELECT '*' CARRIER_PROD_CODE, 'All' CARRIER_PROD_DESC from Dual")

            grdSOTCARRS_UPS.DataSource = dst.Tables("SOTCARRS_UPS")
            ASCMAIN1.Add_Value_List(grdSOTCARRS_UPS, "CARRIER_PROD_CODE", "SELECT CARRIER_PROD_CODE, CARRIER_PROD_DESC FROM SOTCARR2 WHERE CARRIER_CODE = 'UPS'
                                                                            UNION
                                                                         SELECT '*' CARRIER_PROD_CODE, 'All' CARRIER_PROD_DESC from Dual")
        Else
            grdSOTCARRS_FEDEX.DataSource = dst.Tables("ARTCUST1")

            grdSOTCARRS_FEDEX.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_FEDEX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdSOTCARRS_FEDEX.DisplayLayout.Bands(0).Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_FEDEX.DisplayLayout.Bands(0).Override.AllowDelete = DefaultableBoolean.False


            grdSOTCARRS_UPS.DataSource = dst.Tables("ARTCUST1")
            grdSOTCARRS_UPS.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_UPS.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdSOTCARRS_UPS.DisplayLayout.Bands(0).Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_UPS.DisplayLayout.Bands(0).Override.AllowDelete = DefaultableBoolean.False

        End If

        With grdARTCUST2.DisplayLayout.Bands(0)
            '.Columns("CUST_STORE_NO").Header.Fixed = True
            '.Columns("CUST_STORE_NAME").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        'Call InitializeControls(Me)
        ASCMAIN1.Add_Value_List(grdARTCUST2, "CUST_ADDR_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive", "C:Closed"})

        Set_Read_Only_for_ctl(Absx1.optFor("CUST_SHIP_COMPLETE"), True)
        Set_Read_Only_for_ctl(Absx1.chkFor("CUST_CONS_INV"), True)
        '    Absx1.chkFor("CUST_SHIP_COMPLETE").Enabled = False
        '    Absx1.chkFor("CUST_CONS_INV").Enabled = False
        '    Absx1.chkFor("CUST_EDI_DTS_FLAG").Enabled = False

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Absx1.chkFor("CUST_FIN_CHG_IND").Visible = True

            ' lblCUST_ROUTING_INST.Text = "Order Msg"
            'Absx1.txtFor("CUST_ROUTING_INST").Visible = False

            lblCUST_ROUTING_INST.Text = "Internal Msg"
            lblCUST_SPECIAL_INST.Text = "Shipping Inst"

        Else
            Absx1.chkFor("CUST_FIN_CHG_IND").Visible = False
        End If

        Dim tbl As DataTable = ASCDATA1.GetDataTable("Select * from TATCURR1")
        If tbl.Rows.Count < 2 Or (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT") Then
            lblCURR_CODE.Visible = False
            txtCURR_CODE.Visible = False
        Else
            lblCURR_CODE.Visible = True
            txtCURR_CODE.Visible = True
        End If

        'lblCURR_CODE.Visible = (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")
        'txtCURR_CODE.Visible = (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")
        'txtCURR_DESC.Visible = (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")

        lblSTAX_CODE.Visible = (ASCMAIN1.CLIENT = "NYA")
        txtSTAX_CODE.Visible = (ASCMAIN1.CLIENT = "NYA")
        txtSTAX_DESC.Visible = (ASCMAIN1.CLIENT = "NYA")

        lblPvtLblCode.Visible = (ASCMAIN1.CLIENT = "RGI")
        Absx1.txtFor("LABEL_TEMPLATE_CODE").Visible = (ASCMAIN1.CLIENT = "RGI")

        grpRGI_Pricing.Visible = (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
        grpRGI_Pricing_PVC.Visible = (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")

        lblLABEL_TEMPLATE_CODE.Visible = (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.CLIENT = "RGI")
        Absx1.txtFor("PVT_LBL_CODE").Visible = (ASCMAIN1.CLIENT = "RGI")
        Absx1.numFor("PVT_LBL_DISC_PCT").Visible = (ASCMAIN1.CLIENT = "RGI")

        tblUPSReference = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC FROM SOTCARRR  WHERE CARRIER_CODE = 'UPS'")
        tblFDXReference = ASCDATA1.GetDataTable("SELECT REF_CODE, REF_DESC FROM SOTCARRR  WHERE CARRIER_CODE = 'FEDEX'")

        tblFDXReference.Rows.Add(New Object() {"", ""})
        tblUPSReference.Rows.Add(New Object() {"", ""})

        cbeFDXREF1.DataSource = tblFDXReference
        cbeFDXREF2.DataSource = tblFDXReference
        cbeFDXREF3.DataSource = tblFDXReference

        cbeUPSREF1.DataSource = tblUPSReference
        cbeUPSREF2.DataSource = tblUPSReference

        For Each cb As Infragistics.Win.UltraWinEditors.UltraComboEditor _
        In New Infragistics.Win.UltraWinEditors.UltraComboEditor() {cbeFDXREF1, cbeFDXREF2, cbeFDXREF3, cbeUPSREF1, cbeUPSREF2}
            cb.DisplayMember = "REF_DESC"
            cb.ValueMember = "REF_CODE"
        Next

        Bind_Controls(tabShipping.Tabs(0).TabPage, "ARTCUSTS")
        Bind_Controls(tabShipping.Tabs(1).TabPage, "ARTCUSTS")

        grdTATSHIPP.DataSource = dst.Tables("TATSHIPP")
        grdTATSHIPP.Visible = (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI")
        chkCUST_ORDR_CALL_B4_SHIPPING.Visible = ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI"

        grpSegments.Visible = (ASCMAIN1.CLIENT = "NYA")
        GL_Segments(grpSegments, ROWs("GLTPARM1"))

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            SplitContainer4.Panel2.Visible = True
            SplitContainer4.Panel2Collapsed = False
        Else
            SplitContainer4.Panel2.Visible = False
            SplitContainer4.Panel2Collapsed = True
        End If

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTSREP1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(CreditCardQueue1.UserControlGrid, "B", "Use Customer Address")
        Load_Popup_Menu(grdARTCUST2, "SSB", "Show Filter", "Show GroupBox", "Add Ship-to From Master")
        Load_Popup_Menu(grdARTCUSTD, "SSB", "Show Filter", "Show GroupBox", "Make Contact From Main")
        Load_Popup_Menu(grdARTCUSTM, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        If e.SourceControl.Name = CreditCardQueue1.UserControlGrid.Name Then
            If CreditCardQueue1.UserControlGrid.ActiveRow Is Nothing Then
                e.Cancel = True
            End If
        Else
            'e.Cancel = True
        End If

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

            Case "grdARTSREP1"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdARTCUSTD"
                tlb_btn = DirectCast(tlb_pop.Tools("Make Contact From Main"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New") And ASCMAIN1.DBS_COMPANY = "RGI"

            Case "grdARTCUST2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Ship-to From Master"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New") And ASCMAIN1.DBS_COMPANY = "RGI"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            Case "Add Codes"
                If grd.Name = "grdARTSREP1" Then
                    'Add_Codes(grdARTSREP1, "SOTSDIV1", "SALES_DIVISION_CODE", "Divisions")
                    Add_Codes(grdARTSREP1, "SOTSREP1", "SREP_CODE", "Sales Reps")
                End If

            Case "Use Customer Address"
                With CreditCardQueue1.UserControlGrid
                    If .ActiveRow IsNot Nothing Then
                        If Not ScreenMode Then Exit Sub
                        .ActiveRow.Cells("CUST_CREDIT_CARD_ADDR1").Value = MyBase.Absx1.txtFor("CUST_ADDR1").Text
                        .ActiveRow.Cells("CUST_CREDIT_CARD_CITY").Value = MyBase.Absx1.txtFor("CUST_CITY").Text
                        .ActiveRow.Cells("CUST_CREDIT_CARD_STATE").Value = MyBase.Absx1.txtFor("CUST_STATE").Text
                        .ActiveRow.Cells("CUST_CREDIT_CARD_ZIP_CODE").Value = MyBase.Absx1.txtFor("CUST_ZIP_CODE").Text
                    End If
                End With
            Case "Add Ship-to From Master"
                If grd.Name = "grdARTCUST2" Then
                    MakeARTCUST2()
                End If
            Case "Make Contact From Main"
                MakeContactFromMain()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME

            Case "SREP_CODE", "SREP2_CODE"
                sql_where = "NVL(SREP_STATUS,'A') = 'A'"
        End Select
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "Edit"

            Case "Update"

                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

                Dim rowTATSTATE As DataRow = LookUp("TATSTATE", Absx1.txtFor("CUST_STATE").Text)
                If rowTATSTATE Is Nothing Then
                    If Absx1.txtFor("CUST_COUNTRY").Text = "" _
                    Or Absx1.txtFor("CUST_COUNTRY").Text = "USA" Then
                        EMsg &= "Invalid Value Specified for State"
                    End If
                End If

                If Absx1.txtFor("CUST_STATE").Text = "US" Or Absx1.txtFor("CUST_STATE").Text = "USA" Then
                    EMsg &= "Leave Country Blank for USA"
                End If

                ' DO THE FOLLOWING FOR MANDATORY CODES
                'For Each COLUMN_NAME As String In New String() _
                '    {"TERM_CODE", "SREP_CODE", "POST_CODE", "STAX_CODE", "TRADE_CLASS_CODE", "PRICE_CLASS_CODE", "PRICE_LIST_CODE", _
                '     "POST_CODE", "CUST_CLASS_CODE", "SHIP_VIA_CODE", "ROUTING_CODE", "WHSE_CODE", "CUST_BILL_TO_CUST", "CUST_CREDIT_GROUP_CUST", "VEND_CODE"}
                '    Validate_Code(COLUMN_NAME)
                'Next

                Dim rowSOTSREP1 = LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
                If rowSOTSREP1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value specified for Sales Rep Code"
                End If

                For Each ROW As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ARTCUSTD"), New String() {"CONTACT_TYPE"}).Rows
                    Dim CONTACT_TYPE As String = ROW.Item("CONTACT_TYPE") & ""
                    Dim sqlw As String = "CONTACT_TYPE = '" & CONTACT_TYPE & "'"
                    Dim c As Integer = Val(dst.Tables("ARTCUSTD").Compute("COUNT(CONTACT_NO)", sqlw & " and CONTACT_PRIMARY = '1'") & "")
                    If c > 1 Then
                        EMsg &= vbCr & "Cannot have > 1 Primary Contact of any Type (see Type " & CONTACT_TYPE & ")"
                    ElseIf c = 0 Then
                        Dim rows() As DataRow = dst.Tables("ARTCUSTD").Select(sqlw)
                        If rows.Length = 1 Then
                            rows(0).Item("CONTACT_PRIMARY") = "1"
                        Else
                            EMsg &= vbCr & "You must select a Primary Contact for each Type of Contact (see Type " & CONTACT_TYPE & ")"
                        End If
                    End If
                Next

                If ASCMAIN1.CLIENT = "NYA" Then
                    Dim CURR_CODE As String = rowASFBASE1.Item("CURR_CODE") & "" ' Absx1.txtFor("CURR_CODE").Text
                    Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", CURR_CODE)
                    If rowTATCURR1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Currency Code Specified (" & CURR_CODE & ")"
                    End If

                    If EntryMode = "New" Then
                        Dim SEG4_CODE As String = rowASFBASE1.Item("SEG4_CODE") & ""
                        If SEG4_CODE = "001" And CURR_CODE <> "CAD" Then
                            EMsg &= vbCr & "Currency Code (" & CURR_CODE & ") not consistent with Company (" & SEG4_CODE & ")"
                        ElseIf SEG4_CODE <> "001" And CURR_CODE <> "USD" Then
                            EMsg &= vbCr & "Currency Code (" & CURR_CODE & ") not consistent with Company (" & SEG4_CODE & ")"
                        End If
                    End If
                End If


                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    If optCUST_PRICE_TIER.Value = "SP" Then
                        Dim DISC_PCT_MSG As String = DISC_PCT_CHECK()
                        If DISC_PCT_MSG.Length > 0 Then
                            EMsg &= vbCr & DISC_PCT_MSG
                        End If
                    End If
                    'If CUST_DISC_PCT <> 0 And CUST_DISC_PCT <> 52 Then
                    '    EMsg &= vbCr & "This value may only be 52% or 0%"
                    'End If
                    'If CUST_DISC_PCT <> 0 And (CUST_DISC_PCT < 0 Or CUST_DISC_PCT > 60) Then
                    '    EMsg &= vbCr & "This value may only be between 0% and 60%"
                    'End If

                    If Absx1.txtFor("PVT_LBL_CODE").Text <> "" And (Val(Absx1.numFor("PVT_LBL_DISC_PCT").Text & "") < 40 Or Val(Absx1.numFor("PVT_LBL_DISC_PCT").Text & "") > 70) Then
                        EMsg &= vbCr & "Private Label Disc% must be between 40% and 70% "
                    End If

                    If EMsg.Length = 0 Then
                        Dim creditcarderror As String = CreditCardQueue1.ValidateCardData
                        If creditcarderror.Length > 0 Then
                            MessageBox.Show($"FYI: Below are Credit Card errors.{Environment.NewLine}{Environment.NewLine}{creditcarderror}", "Update ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    End If

                    If EMsg.Length = 0 Then
                        If dst.Tables.Item("ARTCUST2").Rows.Count = 0 Then
                            Dim Msg As String = ""
                            Msg = "There Are No Ship-To's Defined."
                            Msg = Msg & vbCrLf & "Would You Like Me To Create One From The Bill-To?"
                            Dim iResult As MsgBoxResult = MsgBox(Msg, MsgBoxStyle.YesNo, "Create New Ship-To?")
                            If iResult = MsgBoxResult.Yes Then
                                MakeARTCUST2()
                            Else
                                EMsg &= vbCr & "Please Create At Least One Ship-To For This Customer"
                            End If
                        End If
                    End If
                End If

                If dst.Tables("SOTCARRS_FEDEX").Rows.Count > 0 Then
                    Select Case dst.Tables("SOTCARRS_FEDEX").Select("DEFAULT_ACCOUNT = '1'").Length
                        Case 0
                            If MessageBox.Show("You do not have a default FedEx Third Party Account. If you proceed no account will be used for Third Party Billing. Do you want to Update Anyway?", "Update ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                                EMsg &= vbCr & "No FedEx Third Party Account defined as Default"
                            End If

                        Case 1

                        Case Else
                            EMsg &= vbCr & "Only 1 FedEx Third Party Account can be defined as Default"

                    End Select

                End If

                If dst.Tables("SOTCARRS_UPS").Rows.Count > 0 Then
                    Select Case dst.Tables("SOTCARRS_UPS").Select("DEFAULT_ACCOUNT = '1'").Length
                        Case 0
                            If MessageBox.Show("You do not have a default UPS Third Party Account. If you proceed no account will be used for Third Party Billing. Do you want to Update Anyway?", "Update ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                                EMsg &= vbCr & "No UPS Third Party Account defined as Default"
                            End If

                        Case 1

                        Case Else
                            EMsg &= vbCr & "Only 1 UPS Third Party Account can be defined as Default"

                    End Select

                End If

        End Select
    End Sub

    Private Function DISC_PCT_CHECK() As String
        'The Only valid Values allowed now are 52, 54, 55, 56, 57 & 59 per Danny - WR 1/23/21
        'This Function Exists in ARTCUST1 and SOTCUST1.  Make Changes To Both Or Suffer The Consequences.
        Dim RETVAL As String = ""
        Dim VALID_PCT As Decimal() = {52, 54, 55, 56, 57, 59}
        Dim VALID_PCT_STR As String = "52, 54, 55, 56, 57 & 59"

        Dim CUST_DISC_PCT As Decimal = Val(Absx1.numFor("CUST_DISC_PCT").Value & "")
        If Not VALID_PCT.Contains(CUST_DISC_PCT) Then
            RETVAL = "Disc % Can Only Be " & VALID_PCT_STR
        End If
        Return RETVAL
    End Function

    Overrides Sub Proceed_Update_Special_Pre()
        grdARTCUST2.UpdateData()
        grdARTCUSTD.UpdateData()
        grdARTCUSTM.UpdateData()

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

        Dim sqlDelete = ""
        Update_Record_TDA("ARTCUST2")
        Update_Record_TDA("ARTCUSTD")
        Update_Record_TDA("ARTSREP1")
        Update_Record_TDA("ARTCUSTS", $"CUST_CODE = '{CUST_CODE}'")

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            Update_Record_TDA("SOTCARRS_UPS", $"DELETE FROM SOTCARRS WHERE CUST_CODE = '{CUST_CODE}' AND CARRIER_CODE = 'UPS'")
            Update_Record_TDA("SOTCARRS_FEDEX", $"DELETE FROM SOTCARRS WHERE CUST_CODE = '{CUST_CODE}' AND CARRIER_CODE = 'FEDEX'")
        End If

        Update_Record_TDA("TATSHIPP", $"TABLE_NAME = 'ARTCUST1' AND KEY_VALUE = '{CUST_CODE}'")
        Update_Record_TDA("WBTCUST1")
        Update_Record_TDA("ARTCUSTM", $"CUST_CODE = '{CUST_CODE}'")

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Update_Record_TDA("ARTCUSTQ", $"CUST_CODE = '{CUST_CODE}'")
        End If
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        If EntryMode = "New" Then
            ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        Else
            ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            CreditCardQueue1.UpdateData()
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        With grdARTCUSTD.DisplayLayout.Bands(0)
            For Each C As String In New String() {"CONTACT_PHONE", "CONTACT_FAX", "CONTACT_CELL"}
                .Columns(C).MaskInput = "" ' "(###) ###-####"
                .Columns(C).CellDisplayStyle = UltraWinGrid.CellDisplayStyle.Default ' UltraWinGrid.CellDisplayStyle.FormattedText
            Next
        End With


        If EntryMode = "New" Then
            rowASFBASE1.Item("CUST_CREDIT_LIMIT") = Val(ROWs("ARTPARM1").Item("AR_PARM_INITIAL_CR_LIMIT") & "")
            If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
                rowASFBASE1.Item("CUST_CRED_LIMIT_EST") = DATETIME_STAMP.Date
            End If
            rowASFBASE1.Item("CUST_CREDIT_LIMIT_NOTES") = "Initial Credit Limit"
            rowASFBASE1.Item("CUST_STMT_IND") = "M"
            rowASFBASE1.Item("TERM_CODE") = ROWs("ARTPARM1").Item("AR_PARM_TERM_CODE")
            rowASFBASE1.Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE")
            rowASFBASE1.Item("CUST_STATUS") = "A"
            rowASFBASE1.Item("WHSE_CODE") = "MS"
            rowASFBASE1.Item("CUST_PRICE_TIER") = "PC"
            rowASFBASE1.Item("CUST_PRICE_TIER_PVC") = "PC"
            If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
                rowASFBASE1.Item("CUST_STATUS_DATE") = Now.Date ' DATETIME_STAMP.Date
            End If
            rowASFBASE1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                rowASFBASE1.Item("CUST_FACTOR_IND") = "1"
            End If
        End If

        If ASCMAIN1.CLIENT = "NYA" Then
            If EntryMode = "New" Then
                'Set_Read_Only_for_ctl(txtCURR_CODE, False)
                txtCURR_CODE.ReadOnly = False
                '  txtCURR_CODE.Enabled = True
            Else
                'Set_Read_Only_for_ctl(txtCURR_CODE, True)
                txtCURR_CODE.ReadOnly = True
            End If
        End If

        EnforceConstraints(False)
        Fill_Records("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text})
        Fill_Records("ARTCUSTD", New String() {Absx1.txtFor("CUST_CODE").Text})
        Fill_Records("ARTSREP1", New String() {Absx1.txtFor("CUST_CODE").Text})
        Fill_Records("ARTCUSTS", New String() {Absx1.txtFor("CUST_CODE").Text})

        If ASCMAIN1.DBS_COMPANY = "RGI" Then
            Fill_Records("SOTCARRS_FEDEX", New String() {Absx1.txtFor("CUST_CODE").Text, "FEDEX"})
            Fill_Records("SOTCARRS_UPS", New String() {Absx1.txtFor("CUST_CODE").Text, "UPS"})
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Fill_Records("ARTCUSTQ", New String() {Absx1.txtFor("CUST_CODE").Text})
            btnWebTaxId.Visible = ShowWebTaxIDBtn()
        Else
            btnWebTaxId.Visible = False
        End If
        Fill_Records("TATSHIPP", New String() {"ARTCUST1", Absx1.txtFor("CUST_CODE").Text})

        ASCMAIN1.sql = "SELECT ARTCUSTM.*, ARTCUST1.CUST_NAME FROM ARTCUSTM, ARTCUST1 WHERE ARTCUSTM.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'" _
            & " AND ARTCUSTM.CUST_CODE_M = ARTCUST1.CUST_CODE (+)"
        Fill_Records("ARTCUSTM", String.Empty, True, ASCMAIN1.sql)

        If dst.Tables("ARTCUSTS").Rows.Count = 0 Then
            Dim rowARTCUSTS As DataRow = dst.Tables("ARTCUSTS").NewRow
            rowARTCUSTS.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
            dst.Tables("ARTCUSTS").Rows.Add(rowARTCUSTS)
        End If

        EnforceConstraints(True)

        Setup_splUpperRight()
        Set_Pricing_Visibility()

        If ASCMAIN1.CLIENT = "RGI" Then
            CreditCardQueue1.ClearData()
            CreditCardQueue1.AllowAutoAuthForm = True
            CreditCardQueue1.AllowEdit = EntryMode = "New" OrElse EntryMode = "Edit"
            CreditCardQueue1.DisplayData(Absx1.txtFor("CUST_CODE").Text)
        End If

    End Sub

    Private Function ShowWebTaxIDBtn() As Boolean
        Dim RetVal As Boolean = False
        Dim CC As String = Absx1.txtFor("CUST_CODE").Text & String.Empty
        TAX_ID = ""
        TAX_ID_DOC = ""
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT CUST_CODE_ACTUAL, TAX_ID, TAX_ID_DOC")
        sql.AppendLine("FROM WBTCUST1")
        sql.AppendLine($"WHERE CUST_CODE_ACTUAL = '{CC}'")
        sql.AppendLine("AND NVL(TAX_ID_DOC,'NULL') <> 'NULL'")
        sql.AppendLine("AND NVL(TAX_ID,'NULL') <> 'NULL'")
        Dim tblWBTCUST1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())
        If tblWBTCUST1.Rows.Count >= 1 Then
            RetVal = True
            TAX_ID = tblWBTCUST1.Rows(0).Item("TAX_ID").ToString & String.Empty
            TAX_ID_DOC = tblWBTCUST1.Rows(0).Item("TAX_ID_DOC").ToString & String.Empty
        End If
        Return RetVal
    End Function

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ARTCUST2", "ARTCUSTD", "ARTSREP1", "ARTCUSTS", "TATSHIPP", "ARTCUSTM", "SOTCARRS_UPS", "SOTCARRS_FEDEX"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)

            If ASCMAIN1.CLIENT = "RGI" Then
                CreditCardQueue1.ClearData()
            End If

            btnWebTaxId.Visible = False

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                dst.Tables("ARTCUSTQ").Rows.Clear()
                dteLAST_DATE.Value = Null
                txtLAST_OPER.Value = Null
                chkRESIDENTIAL_ORDR.Enabled = False
                chkRESIDENTIAL_ORDR.Checked = False
                chkINSIDE_REQ.Enabled = False
                chkINSIDE_REQ.Checked = False
                chkGATE_LIFT_REQ.Enabled = False
                chkGATE_LIFT_REQ.Checked = False
                chkLIMITED_ACCESS.Enabled = False
                chkLIMITED_ACCESS.Checked = False
                txtLIMITED_ACCESS_NOTE.Enabled = False
                txtLIMITED_ACCESS_NOTE.Text = ""
                chkIRREGULAR_HOURS.Enabled = False
                chkIRREGULAR_HOURS.Checked = False
                txtIRREGULAR_HOURS_NOTE.Enabled = False
                txtIRREGULAR_HOURS_NOTE.Text = ""
                chkBROKER.Enabled = False
                chkBROKER.Checked = False
                txtBROKER_NOTE.Enabled = False
                txtBROKER_NOTE.Text = ""
                chkAPPOINTMENT_REQUIRED.Enabled = False
                chkAPPOINTMENT_REQUIRED.Text = ""
                txtAPPOINTMENT_REQUIRED_NOTE.Enabled = False
                txtAPPOINTMENT_REQUIRED_NOTE.Text = ""
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                CreditCardQueue1.AllowEdit = False
                CreditCardQueue1.CustomerCode = String.Empty
                CreditCardQueue1.CancelUpdate()
            End If

        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        'Set_Read_Only_for_ctl(Absx1.txtFor("CUST_NAME"), Not tf)
        'Set_Read_Only(grpCreditLimit, True)
        ' Set_Read_Only(grpOther, True)
        ' Set_Read_Only(grpCreditLimit, IIf(Not tf, ASCMAIN1.USER_SECURITY_CODEs.Contains("CL"), True))
        Set_Read_Only(grpCreditLimit, True)

        tabARTCUST1.Tabs("Credit Cards").Visible = ASCMAIN1.CLIENT = "RGI"

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCUSTD, grdARTCUST2, grdARTSREP1, grdARTCUSTM}
            With grd.DisplayLayout.Override
                If (EntryMode = "New" Or EntryMode = "Edit") Then ' And grd.Name <> "grdARTCUST2" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next

        If Not ScreenMode Then
            btnNewCustomer.Visible = (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")
            btnPullFromWeb.Visible = False
        Else
            btnNewCustomer.Visible = False
            btnPullFromWeb.Visible = (ASCMAIN1.DBS_COMPANY = "RGI") And EntryMode = "New"
        End If
        btnVerifyShipToInfo.Visible = (ASCMAIN1.DBS_COMPANY = "RGI") And EntryMode = "Edit"

        CreditCardQueue1.AllowEdit = False
        If ASCMAIN1.CLIENT = "RGI" Then
            If (EntryMode = "Edit" OrElse EntryMode = "New") Then
                CreditCardQueue1.AllowEdit = True
            Else
                CreditCardQueue1.AllowEdit = False
            End If
        End If
        CreditCardQueue1.SetUpScreen()

        If ASCMAIN1.DBS_COMPANY <> "RGI" Then
            grdSOTCARRS_FEDEX.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_FEDEX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdSOTCARRS_FEDEX.DisplayLayout.Bands(0).Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_FEDEX.DisplayLayout.Bands(0).Override.AllowDelete = DefaultableBoolean.False

            grdSOTCARRS_UPS.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_UPS.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdSOTCARRS_UPS.DisplayLayout.Bands(0).Override.AllowAddNew = AllowAddNew.No
            grdSOTCARRS_UPS.DisplayLayout.Bands(0).Override.AllowDelete = DefaultableBoolean.False
        End If

    End Sub

    Public Overrides Sub isDeleteAllowed()
        MyBase.isDeleteAllowed()
        If EMsg = "" Then
            isDeleteAllowed_Check_Aliased_Columns _
            (New String() {"ARTCUST1.CUST_BILL_TO_CUST"})
        End If
    End Sub

    Public Overrides Function Set_Contact_Info() As Boolean
        If ScreenMode Then
            CONTACT_ENTITY_KEY = Absx1.txtFor("CUST_CODE").Text
            CONTACT_ENTITY_NAME = rowASFBASE1.Item("CUST_NAME") & "" ' .txtFor("CUST_NAME").Text
        End If
        Return True
    End Function

#End Region

#Region "grdARTCUST2"

    Private Sub grdARTCUST2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.AfterCellUpdate
        'Select Case e.Cell.Column.Key
        '    Case "CUST_CODE"
        '        Dim row As DataRow = LookUp("ARTCUST1", e.Cell.Value)
        '        If row IsNot Nothing Then
        '            grdARTCUST2.ActiveRow.Cells("CUST_NAME").Value = row.Item("CUST_NAME")
        '        End If
        'End Select
    End Sub

    Private Sub grdARTCUST2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUST2.AfterRowActivate
        With grdARTCUST2.DisplayLayout.Bands("ARTCUST2")
            If grdARTCUST2.ActiveRow.IsAddRow Then
                .Columns("CUST_ADDR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_ADDR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                Dim CC As String = grdARTCUST2.ActiveRow.Cells("CUST_ADDR_CODE").Value & String.Empty
                If CC.Length > 0 Then
                    bindControl(CC)
                End If
            End If
        End With
    End Sub

    Private Sub grdARTCUST2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdARTCUST2.BeforeCellUpdate

        Select Case e.Cell.Column.Key
            Case "CUST_ADDR_CODE"
                'Dim STYLE_CODE As String = Validate_Style(e.NewValue & "")
                'If STYLE_CODE = "" Then
                '    e.Cancel = True
                'End If
        End Select

    End Sub

    Private Sub grdARTCUST2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTCUST2.BeforeExitEditMode
        If grdARTCUST2.ActiveCell IsNot Nothing Then
            With grdARTCUST2.ActiveCell
                Select Case .Column.Key
                    Case "CUST_ADDR_CODE"
                        '.EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        If .EditorResolved.IsValid Then
                            .EditorResolved.Value = .EditorResolved.Value.ToString.PadLeft(6, "0")
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdARTCUST2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUST2.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        Dim EMsg As String = ""

        If e.Row.Cells("CUST_STATE").Value & "" <> "" Then
            Dim rowTATSTATE As DataRow = LookUp("TATSTATE", e.Row.Cells("CUST_STATE").Value & "")
            If rowTATSTATE Is Nothing Then
                EMsg &= vbCr & "Invalid State"
                e.Cancel = True
            End If
        End If

        If e.Row.Cells("CUST_COUNTRY").Value & "" <> "" Then
            Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", e.Row.Cells("CUST_COUNTRY").Value & "")
            If rowTATCNTRY Is Nothing Then
                EMsg &= vbCr & "Invalid Country"
                e.Cancel = True
            End If
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Data Errors")
        End If
        'If row Is Nothing Then
        '    e.Cancel = True
        'End If

        If Not e.Cancel Then
            If e.Row.IsAddRow Then
                e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            End If
            If ASCMAIN1.DBS_COMPANY = "RGI" Then
                If e.Row.Cells("CUST_ADDR_CODE").Value & "" = "" Then
                    e.Row.Cells("CUST_ADDR_CODE").Value = GetNextCUST_ADDR_CODE()
                End If
            End If
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            e.Row.Cells("CUST_ADDR_TYPE").Value = "MK"
        End If
    End Sub

    Private Sub grdARTCUST2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.ClickCellButton
        'Dim sql_where As String = ""
        'Call grdClickCellButton(grdARTCUST2, sql_where, True)
    End Sub

    Private Sub grdARTCUST2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUST2.Error
        grdARTCUST2.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdARTCUST2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUST2.InitializeRow
        If e.Row.IsDataRow Then
            Dim CUST_COUNTRY As String = e.Row.Cells("CUST_COUNTRY").Value & ""
            If CUST_COUNTRY = "" Or CUST_COUNTRY = "US" Or CUST_COUNTRY = "USA" Then
                'e.Row.Cells("CUST_PHONE").EditorComponent = medCUST_PHONE_STX
                'e.Row.Cells("CUST_FAX").EditorComponent = medCUST_FAX_STX
                e.Row.Cells("CUST_PHONE").EditorComponent = Nothing
                e.Row.Cells("CUST_FAX").EditorComponent = Nothing

            Else
                e.Row.Cells("CUST_PHONE").EditorComponent = Nothing
                e.Row.Cells("CUST_FAX").EditorComponent = Nothing
            End If

            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                If Not e.Row.IsAddRow Then
                    Dim CC As String = e.Row.Cells("CUST_ADDR_CODE").Value
                    bindControl(CC)
                End If
            End If

        End If
    End Sub

#End Region

#Region "grdARTCUSTD"

    Private Sub grdARTCUSTD_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTD.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "CUST_CODE"
            '    Dim row As DataRow = LookUp("ARTCUST1", e.Cell.Value)
            '    If row IsNot Nothing Then
            '        grdARTCUSTD.ActiveRow.Cells("CUST_NAME").Value = row.Item("CUST_NAME")
            '    End If
        End Select
    End Sub

    Private Sub grdARTCUSTD_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUSTD.AfterRowActivate
        With grdARTCUSTD.DisplayLayout.Bands("ARTCUSTD")
            'If grdARTCUSTD.ActiveRow.IsAddRow Then
            '    .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            'Else
            '    .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            'End If
        End With
    End Sub

    Private Sub grdARTCUSTD_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUSTD.BeforeRowsDeleted

    End Sub

    Private Sub grdARTCUSTD_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSTD.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        'If row Is Nothing Then
        '    e.Cancel = True
        'End If
        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
            e.Row.Cells("CONTACT_NO").Value = Val(dst.Tables("ARTCUSTD").Compute("MAX(CONTACT_NO)", "") & "") + 1
        End If
    End Sub

    Private Sub grdARTCUSTD_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTD.ClickCellButton
        'Dim sql_where As String = ""
        'grdClickCellButton(grdARTCUSTD, sql_where, True)
    End Sub

    Private Sub grdARTCUSTD_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUSTD.Error
        grdARTCUSTD.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdARTCUSTD_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUSTD.InitializeRow
    End Sub

#End Region

#Region "grdARTSREP1"

    Private Sub grdARTSREP1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTSREP1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SALES_DIVISION_CODE"
                grdCodeDesc(grdARTSREP1, "SOTSDIV1", "SALES_DIVISION_CODE", "SALES_DIVISION_NAME")
            Case "SREP_CODE"
                grdCodeDesc(grdARTSREP1, "SOTSREP1", "SREP_CODE", "SREP_NAME")
        End Select
    End Sub

    Private Sub grdARTSREP1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTSREP1.BeforeRowUpdate
        Dim row As DataRow = LookUp("SOTSDIV1", e.Row.Cells("SALES_DIVISION_CODE").Text)
        If row IsNot Nothing Then row = LookUp("SOTSREP1", e.Row.Cells("SREP_CODE").Text)
        If row Is Nothing Then e.Cancel = True

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
        End If
    End Sub

    Private Sub grdARTSREP1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTSREP1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SALES_DIVISION_CODE"
                Dim sql_where As String = "" ' Get_List_of_Codes("SOTSDIV1.SALES_DIVISION_CODE not in", "ARTSREP1", "SALES_DIVISION_CODE")
                grdClickCellButton(grdARTSREP1, sql_where, True)
            Case "SREP_CODE"
                Dim sql_where As String = "" ' Get_List_of_Codes("SOTSREP1.SREP_CODE not in", "ARTSREP1", "SREP_CODE")
                grdClickCellButton(grdARTSREP1, sql_where, True)
        End Select
    End Sub

    Private Sub grdARTSREP1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTSREP1.InitializeRow
        grd_RowColor(dst.Tables("ARTSREP1"), e.Row)
    End Sub

#End Region

#Region "grdARTCUSTM"

    Private Sub grdARTCUSTM_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSTM.BeforeRowUpdate

        e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text

        Dim CUST_CODE_M As String = e.Row.Cells("CUST_CODE_M").Value & String.Empty
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE_M)
        If rowARTCUST1 Is Nothing Then
            MessageBox.Show("Invalid Customer.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        e.Row.Cells("CUST_NAME").Value = rowARTCUST1.Item("CUST_NAME") & String.Empty
    End Sub

    Private Sub grdARTCUSTM_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTM.ClickCellButton
        Dim sql_where As String = String.Empty

        If Not (EntryMode = "New" Or EntryMode = "Edit") Then
            Exit Sub
        End If

        Select Case e.Cell.Column.Key
            Case "CUST_CODE_M"

                Dim lstCustCodes As List(Of String) = (From r In dst.Tables("ARTCUSTM").AsEnumerable() Select r.Field(Of String)("CUST_CODE_M")).ToList()
                If lstCustCodes.Count > 0 Then
                    sql_where = "CUST_CODE NOT IN ('" & String.Join("', '", lstCustCodes.ToArray) & "')"
                End If

                grdClickCellButton(grdARTCUSTM, sql_where, True, "CUST_CODE_M", "CUST_CODE")

            Case Else
                Exit Sub
        End Select

    End Sub

    Private Sub grdARTCUSTM_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUSTM.Error
        grdARTCUSTM.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub


#End Region

#Region "grdSOTCARRS_FEDEX"

    Private Sub grdSOTCARRS_FEDEX_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSOTCARRS_FEDEX.BeforeRowUpdate

        If ASCMAIN1.DBS_COMPANY <> "RGI" Then
            Exit Sub
        End If

        e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
        e.Row.Cells("CARRIER_CODE").Value = "FEDEX"

        Dim CARRIER_PROD_CODE As String = e.Row.Cells("CUST_CODE").Value & String.Empty
        Dim ACCOUNT_NO As String = e.Row.Cells("ACCOUNT_NO").Value & String.Empty
        Dim ZIP_CODE As String = e.Row.Cells("ZIP_CODE").Value & String.Empty
        Dim COUNTRY_CODE As String = e.Row.Cells("COUNTRY_CODE").Value & String.Empty

        Dim lstErrors As New List(Of String)

        If CARRIER_PROD_CODE.Length = 0 Then
            lstErrors.Add("Prod Code is Required")
        End If

        If ACCOUNT_NO.Length = 0 Then
            lstErrors.Add("Account Code is Required")
        End If

        If ZIP_CODE.Length = 0 Then
            lstErrors.Add("Zip Code is Required")
        End If

        If COUNTRY_CODE.Length = 0 Then
            lstErrors.Add("Country Code is Required")
        End If

        If lstErrors.Count > 0 Then
            MessageBox.Show(String.Join(Environment.NewLine, lstErrors.ToArray), "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

    End Sub

    Private Sub grdSOTCARRS_FEDEX_AfterRowInsert(sender As Object, e As RowEventArgs) Handles grdSOTCARRS_FEDEX.AfterRowInsert

        If ASCMAIN1.DBS_COMPANY <> "RGI" Then
            Exit Sub
        End If

        e.Row.Cells("COUNTRY_CODE").Value = "USA"
    End Sub

#End Region

#Region "grdSOTCARRS_UPS"

    Private Sub grdSOTCARRS_UPS_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSOTCARRS_UPS.BeforeRowUpdate

        If ASCMAIN1.DBS_COMPANY <> "RGI" Then
            Exit Sub
        End If

        e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
        e.Row.Cells("CARRIER_CODE").Value = "UPS"

        Dim CARRIER_PROD_CODE As String = e.Row.Cells("CUST_CODE").Value & String.Empty
        Dim ACCOUNT_NO As String = e.Row.Cells("ACCOUNT_NO").Value & String.Empty
        Dim ZIP_CODE As String = e.Row.Cells("ZIP_CODE").Value & String.Empty
        Dim COUNTRY_CODE As String = e.Row.Cells("COUNTRY_CODE").Value & String.Empty

        Dim lstErrors As New List(Of String)

        If CARRIER_PROD_CODE.Length = 0 Then
            lstErrors.Add("Prod Code is Required")
        End If

        If ACCOUNT_NO.Length = 0 Then
            lstErrors.Add("Account Code is Required")
        End If

        If ZIP_CODE.Length = 0 Then
            lstErrors.Add("Zip Code is Required")
        End If

        If COUNTRY_CODE.Length = 0 Then
            lstErrors.Add("Country Code is Required")
        End If

        If lstErrors.Count > 0 Then
            MessageBox.Show(String.Join(Environment.NewLine, lstErrors.ToArray), "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

    End Sub

    Private Sub grdSOTCARRS_UPS_AfterRowInsert(sender As Object, e As RowEventArgs) Handles grdSOTCARRS_UPS.AfterRowInsert

        If ASCMAIN1.DBS_COMPANY <> "RGI" Then
            Exit Sub
        End If

        e.Row.Cells("COUNTRY_CODE").Value = "USA"
    End Sub

#End Region

    Private Sub tabARTCUST1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabARTCUST1.SelectedTabChanged
        Setup_splUpperRight()
    End Sub

    Sub Setup_splUpperRight()
        If (ASCMAIN1.CLIENT = "VAN") Then
            'If (ASCMAIN1.CLIENT = "VAN" OrElse ASCMAIN1.CLIENT = "RGI") Then
            splUpperRight.Panel1Collapsed = (tabARTCUST1.SelectedTab.Key = "Sales")
            splUpperRight.Panel2Collapsed = Not (tabARTCUST1.SelectedTab.Key = "Sales")
        Else
            splUpperRight.Panel1Collapsed = False
            splUpperRight.Panel2Collapsed = True
        End If

    End Sub

    Private Sub optCUST_PRICE_TIER_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optCUST_PRICE_TIER.ValueChanged
        'If SELECTION_NO = 0 Then Exit Sub
        Set_Pricing_Visibility()
    End Sub

    Sub Set_Pricing_Visibility()
        'Absx1.optFor("CUST_DISC_PCT_EXTRA").Visible = Not (optCUST_PRICE_TIER.Value = "FC" Or optCUST_PRICE_TIER.Value = "HC" Or optCUST_PRICE_TIER.Value = "SP")
        'lblCUST_DISC_PCT_EXTRA.Visible = Not (optCUST_PRICE_TIER.Value = "FC" Or optCUST_PRICE_TIER.Value = "HC" Or optCUST_PRICE_TIER.Value = "SP")
        'Absx1.numFor("CUST_DISC_PCT").Visible = (optCUST_PRICE_TIER.Value = "SP")
        If Not IsNothing(optCUST_PRICE_TIER.Value) Then
            Select Case optCUST_PRICE_TIER.Value
                Case "FC"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = False
                    lblCUST_DISC_PCT_EXTRA.Visible = False
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = False
                    Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case "HC"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = False
                    lblCUST_DISC_PCT_EXTRA.Visible = False
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = False
                    Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case "PC"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = True
                    lblCUST_DISC_PCT_EXTRA.Visible = True
                    'Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = False
                    Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case "SP"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = False
                    lblCUST_DISC_PCT_EXTRA.Visible = False
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = True
                    'Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case Else
                    MsgBox("Error In Price Tier Option", vbOKOnly, "Please Let ABS Know")
            End Select
        End If
    End Sub

    Private Sub btnNewCustomer_Click(sender As System.Object, e As System.EventArgs) Handles btnNewCustomer.Click
        Dim CUST_CODE As String = ASCMAIN1.Next_Control_No("ARTCUST1.CUST_CODE")
        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Click_Command("New")
    End Sub

    Private Sub txtCUST_COUNTRY_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtCUST_COUNTRY.ValueChanged
        Dim CUST_COUNTRY As String = txtCUST_COUNTRY.Text
        If CUST_COUNTRY = "" Or CUST_COUNTRY = "US" Or CUST_COUNTRY = "USA" Then
            medCUST_PHONE.InputMask = "(###) ###-####"
            medCUST_FAX.InputMask = "(###) ###-####"
        Else
            medCUST_PHONE.InputMask = ""
            medCUST_FAX.InputMask = ""
        End If
    End Sub

    Private Sub grdTATSHIPP_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdTATSHIPP.BeforeRowUpdate
        e.Row.Cells("TABLE_NAME").Value = "ARTCUST1"
        e.Row.Cells("KEY_VALUE").Value = Absx1.txtFor("CUST_CODE").Text

        Dim errorMsg As String = String.Empty

        Dim SHIPMENT_AMT As Int32 = Val(e.Row.Cells("SHIPMENT_AMT").Value & String.Empty)
        Dim SHIPMENT_PERC As Int32 = Val(e.Row.Cells("SHIPMENT_PERC").Value & String.Empty)

        If SHIPMENT_AMT <= 0 Then
            errorMsg = "The Shipment Amount must be greater than $0.00"
        End If

        If SHIPMENT_PERC < 0 OrElse SHIPMENT_PERC > 100 Then
            If errorMsg.Length > 0 Then
                errorMsg &= Environment.NewLine
            End If
            errorMsg &= "The Shipment Percentage must be between 1 and 100. Leave blank or set to 0 to be ignored."
        End If

    End Sub

    Private Sub btnPullFromWeb_Click(sender As System.Object, e As System.EventArgs) Handles btnPullFromWeb.Click
        'Stop
        Dim EMAIL As String = Absx1.txtFor("CUST_EMAIL").Text
        Dim ErrorMsg As String = ""
        Dim rowWBTCUST1 As DataRow
        If EMAIL.Length = 0 Then
            ErrorMsg &= vbCr & "You Must Provide A "
        Else
            Fill_Records("WBTCUST1", EMAIL)
            If dst.Tables.Item("WBTCUST1").Rows.Count <> 1 Then
                ErrorMsg &= vbCr & EMAIL & " Not Found In Web Customer Database"
            Else
                If dst.Tables.Item("WBTCUST1").Rows(0).Item("STATUS").ToString <> "C" And dst.Tables.Item("WBTCUST1").Rows(0).Item("STATUS").ToString <> "N" Then
                    ErrorMsg &= vbCr & EMAIL & " Not In Credit Or New Status Any Longer"
                End If
            End If
        End If
        If ErrorMsg.Length <> 0 Then
            MsgBox(ErrorMsg, MsgBoxStyle.Critical, "Can Not Import Web Customer")
            Exit Sub
        End If
        rowWBTCUST1 = dst.Tables.Item("WBTCUST1").Rows(0)
        rowWBTCUST1.Item("CUST_CODE_ACTUAL") = Absx1.txtFor("CUST_CODE").Text
        rowWBTCUST1.Item("CONTACT_NO") = 0
        rowWBTCUST1.Item("CONTACT_TYPE") = "1"
        rowWBTCUST1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowWBTCUST1.Item("LAST_DATE") = DATETIME_STAMP
        rowWBTCUST1.Item("STATUS") = "M"

        Absx1.txtFor("CUST_CONTACT").Value = String.Format("{0} {1}", rowWBTCUST1.Item("GIVENNAME").ToString.ToUpper, rowWBTCUST1.Item("FAMILYNAME").ToString.ToUpper)
        txtCUST_NAME.Value = rowWBTCUST1.Item("COMPANY").ToString.ToUpper

        If (rowWBTCUST1.Item("STREET").ToString & String.Empty).Length <= 60 Then
            Absx1.txtFor("CUST_ADDR1").Value = (rowWBTCUST1.Item("STREET").ToString & String.Empty).ToUpper
        Else
            Absx1.txtFor("CUST_ADDR1").Value = rowWBTCUST1.Item("STREET").ToString.ToUpper.Substring(0, 59)
        End If

        If (rowWBTCUST1.Item("STREET2").ToString & String.Empty).Length <= 60 Then
            Absx1.txtFor("CUST_ADDR2").Value = (rowWBTCUST1.Item("STREET2").ToString & String.Empty).ToUpper
        Else
            Absx1.txtFor("CUST_ADDR2").Value = rowWBTCUST1.Item("STREET2").ToString.ToUpper.Substring(0, 59)
        End If

        If (rowWBTCUST1.Item("STREET3").ToString & String.Empty).Length <= 60 Then
            Absx1.txtFor("CUST_ADDR3").Value = (rowWBTCUST1.Item("STREET3").ToString & String.Empty).ToUpper
        Else
            Absx1.txtFor("CUST_ADDR3").Value = rowWBTCUST1.Item("STREET3").ToString.ToUpper.Substring(0, 59)
        End If

        If (rowWBTCUST1.Item("CITY").ToString & String.Empty).Length <= 30 Then
            Absx1.txtFor("CUST_CITY").Value = (rowWBTCUST1.Item("CITY").ToString & String.Empty).ToUpper
        Else
            Absx1.txtFor("CUST_CITY").Value = rowWBTCUST1.Item("CITY").ToString.ToUpper.Substring(0, 59)
        End If

        If (rowWBTCUST1.Item("COUNTRY").ToString & String.Empty).Length = 3 Then
            Absx1.txtFor("CUST_COUNTRY").Value = (rowWBTCUST1.Item("COUNTRY").ToString & String.Empty).ToUpper
        End If

        If (rowWBTCUST1.Item("WEBSITE").ToString & String.Empty).Length <= 255 Then
            Absx1.txtFor("CUST_URL").Value = (rowWBTCUST1.Item("WEBSITE").ToString & String.Empty).ToUpper
        End If

        If rowWBTCUST1.Item("STATE").ToString.Length <= 2 Then
            Absx1.txtFor("CUST_STATE").Value = rowWBTCUST1.Item("STATE").ToString.ToUpper
        Else
            MsgBox(rowWBTCUST1.Item("STATE").ToString & " Can Not Be Added To State", MsgBoxStyle.Critical, "State")
        End If
        Absx1.txtFor("CUST_ZIP_CODE").Value = rowWBTCUST1.Item("ZIP_CODE").ToString.ToUpper
        Dim TELEPHONE As String = rowWBTCUST1.Item("TELEPHONE").ToString.Replace("-", "").Replace("(", "").Replace(")", "")
        If Not IsNumeric(TELEPHONE) Then
            MsgBox(TELEPHONE & " Can Not Be Added To Telephone", MsgBoxStyle.Critical, "Telephone")
        Else
            If TELEPHONE.Length = 11 Then
                TELEPHONE = TELEPHONE.Substring(1, 10)
            End If
            medCUST_PHONE.Value = Val(TELEPHONE)
        End If

        If dst.Tables.Item("ARTCUST2").Rows.Count = 0 Then
            Dim newARTCUST2 As DataRow = dst.Tables.Item("ARTCUST2").NewRow
            newARTCUST2.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
            newARTCUST2.Item("CUST_ADDR_TYPE") = "MK"
            newARTCUST2.Item("CUST_ADDR_CODE") = "000001"
            newARTCUST2.Item("CUST_NAME") = rowWBTCUST1.Item("COMPANY").ToString.ToUpper
            newARTCUST2.Item("CUST_ADDR1") = rowWBTCUST1.Item("SHP_ADDR_1").ToString.ToUpper
            newARTCUST2.Item("CUST_ADDR2") = rowWBTCUST1.Item("SHP_ADDR_2").ToString.ToUpper
            newARTCUST2.Item("CUST_ADDR3") = rowWBTCUST1.Item("SHP_ADDR_3").ToString.ToUpper
            newARTCUST2.Item("CUST_CITY") = rowWBTCUST1.Item("SHP_CITY").ToString.ToUpper
            newARTCUST2.Item("CUST_STATE") = rowWBTCUST1.Item("SHP_STATE").ToString.ToUpper
            newARTCUST2.Item("CUST_ZIP_CODE") = rowWBTCUST1.Item("SHP_ZIP_CODE").ToString.ToUpper
            newARTCUST2.Item("CUST_COUNTRY") = rowWBTCUST1.Item("SHP_CNTRY").ToString.ToUpper
            'newARTCUST2.Item("CUST_CONTACT") = ""
            'newARTCUST2.Item("CUST_PHONE") = ""
            'newARTCUST2.Item("CUST_EXT") = ""
            'newARTCUST2.Item("CUST_FAX") = ""
            newARTCUST2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            newARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
            newARTCUST2.Item("INIT_DATE") = DATETIME_STAMP
            newARTCUST2.Item("LAST_DATE") = DATETIME_STAMP
            'newARTCUST2.Item("CUST_ADDR_NAME") = ""
            newARTCUST2.Item("CUST_ADDR_STATUS") = "A"
            'newARTCUST2.Item("CUST_EMAIL") = ""
            'newARTCUST2.Item("GLOBAL_LOCATION_NUMBER") = ""
            'newARTCUST2.Item("FDX_ACCT_NO") = ""
            'newARTCUST2.Item("CUST_DC_NO") = ""
            'newARTCUST2.Item("UPS_ACCT_NO") = ""
            'newARTCUST2.Item("CUST_ADDR_GROUP") = ""
            'newARTCUST2.Item("STAX_CODE") = ""
            dst.Tables.Item("ARTCUST2").Rows.Add(newARTCUST2)
        End If

        If dst.Tables.Item("ARTCUSTQ").Rows.Count = 0 Then
            Dim newARTCUSTQ As DataRow = dst.Tables.Item("ARTCUSTQ").NewRow
            newARTCUSTQ.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
            newARTCUSTQ.Item("CUST_ADDR_CODE") = "000001"
            newARTCUSTQ.Item("LAST_DATE") = DATETIME_STAMP
            newARTCUSTQ.Item("LAST_OPER") = ASCMAIN1.USER_ID
            'newARTCUSTQ.Item("LAST_ORDR_NO") = ""
            newARTCUSTQ.Item("RESIDENTIAL_ORDR") = IIf(rowWBTCUST1.Item("RESIDENTIAL").ToString = "Yes", "1", "0")
            newARTCUSTQ.Item("INSIDE_REQ") = IIf(rowWBTCUST1.Item("INSIDE").ToString = "Yes", "1", "0")
            newARTCUSTQ.Item("GATE_LIFT_REQ") = IIf(rowWBTCUST1.Item("GATE_LIFT").ToString = "Yes", "1", "0")
            newARTCUSTQ.Item("LIMITED_ACCESS") = IIf(rowWBTCUST1.Item("LIMITED_ACCESS").ToString.Length = 0, "0", "1")
            newARTCUSTQ.Item("LIMITED_ACCESS_NOTE") = rowWBTCUST1.Item("LIMITED_ACCESS").ToString
            newARTCUSTQ.Item("IRREGULAR_HOURS") = IIf(rowWBTCUST1.Item("IRREGULAR_HOURS_NOTE").ToString.Length = 0, "0", "1")
            newARTCUSTQ.Item("IRREGULAR_HOURS_NOTE") = rowWBTCUST1.Item("IRREGULAR_HOURS_NOTE").ToString
            newARTCUSTQ.Item("APPOINTMENT_REQUIRED") = IIf(rowWBTCUST1.Item("APPOINTMENT_REQUIRED_NOTE").ToString.Length = 0, "0", "1")
            newARTCUSTQ.Item("APPOINTMENT_REQUIRED_NOTE") = rowWBTCUST1.Item("APPOINTMENT_REQUIRED_NOTE").ToString
            newARTCUSTQ.Item("BROKER") = IIf(rowWBTCUST1.Item("BROKER_NOTE").ToString.Length = 0, "0", "1")
            newARTCUSTQ.Item("BROKER_NOTE") = rowWBTCUST1.Item("BROKER_NOTE").ToString

            dst.Tables.Item("ARTCUSTQ").Rows.Add(newARTCUSTQ)
        End If

        btnPullFromWeb.Visible = False

        btnWebTaxId.Visible = ShowWebTaxIDBtn()
    End Sub

    Private Sub MakeARTCUST2()
        Dim NextCUST_ADDR_CODE As String = GetNextCUST_ADDR_CODE()

        Dim newARTCUST2 As DataRow = dst.Tables("ARTCUST2").NewRow()
        newARTCUST2.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
        newARTCUST2.Item("CUST_ADDR_TYPE") = "MK"
        newARTCUST2.Item("CUST_ADDR_CODE") = NextCUST_ADDR_CODE
        newARTCUST2.Item("CUST_ADDR_STATUS") = "A"
        newARTCUST2.Item("INIT_OPER") = ASCMAIN1.USER_ID
        newARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
        newARTCUST2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
        newARTCUST2.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        For Each COLNAME As String In New String() {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY"}
            newARTCUST2.Item(COLNAME) = Absx1.txtFor(COLNAME).Text
        Next
        dst.Tables("ARTCUST2").Rows.Add(newARTCUST2)
        grdARTCUST2.Refresh()
        grdARTCUST2.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
        grdARTCUST2.Update()
        If Not IsNothing(grdARTCUST2.ActiveRow) Then
            If Not grdARTCUST2.ActiveRow Is Nothing AndAlso grdARTCUST2.ActiveRow.DataChanged Then
                grdARTCUST2.ActiveRow.CancelUpdate()
            ElseIf Not grdARTCUST2.ActiveRow IsNot Nothing AndAlso grdARTCUST2.ActiveRow.DataChanged Then
                grdARTCUST2.ActiveRow.Update()
            End If
        End If
    End Sub

    Private Function GetNextCUST_ADDR_CODE() As String
        Dim RetVal As String = ""
        Dim NextCUST_ADDR_CODE As Integer = 0
        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("", "CUST_ADDR_CODE")
            If IsNumeric(rowARTCUST2.Item("CUST_ADDR_CODE")) Then
                NextCUST_ADDR_CODE = CInt(rowARTCUST2.Item("CUST_ADDR_CODE"))
            End If
        Next
        NextCUST_ADDR_CODE += 1
        RetVal = Str(NextCUST_ADDR_CODE).Trim().PadLeft(6, "0")
        Return RetVal
    End Function

    Private Sub MakeContactFromMain()
        Dim newARTCUSTD As DataRow = dst.Tables.Item("ARTCUSTD").NewRow
        Dim CONTACT_NO As Integer = Val(dst.Tables("ARTCUSTD").Compute("MAX(CONTACT_NO)", "") & "") + 1
        With newARTCUSTD
            .Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
            .Item("CONTACT_NO") = CONTACT_NO
            .Item("CONTACT_NAME") = Absx1.txtFor("CUST_CONTACT").Text
            .Item("CONTACT_EMAIL") = Absx1.txtFor("CUST_EMAIL").Text
            .Item("CONTACT_PHONE") = medCUST_PHONE.Text
            .Item("CONTACT_EXT") = Absx1.txtFor("CUST_EXT").Text
            .Item("CONTACT_FAX") = medCUST_FAX.Text
            If CONTACT_NO = 1 Then
                .Item("CONTACT_TYPE") = "B"
                .Item("CONTACT_PRIMARY") = "1"
            Else
                .Item("CONTACT_TYPE") = "M"
                .Item("CONTACT_PRIMARY") = "0"
            End If
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
        End With
        dst.Tables.Item("ARTCUSTD").Rows.Add(newARTCUSTD)
        grdARTCUSTD.Refresh()
        grdARTCUSTD.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
        grdARTCUSTD.Update()
        SendKeys.Send(Chr(27))
    End Sub

    Private Sub chkLIMITED_ACCESS_CheckedChanged(sender As Object, e As EventArgs) Handles chkLIMITED_ACCESS.CheckedChanged
        If chkLIMITED_ACCESS.Checked Then
            txtLIMITED_ACCESS_NOTE.Visible = True
        Else
            txtLIMITED_ACCESS_NOTE.Text = ""
            txtLIMITED_ACCESS_NOTE.Visible = False
        End If
    End Sub

    Private Sub chkIRREGULAR_HOURS_CheckedChanged(sender As Object, e As EventArgs) Handles chkIRREGULAR_HOURS.CheckedChanged
        If chkIRREGULAR_HOURS.Checked Then
            txtIRREGULAR_HOURS_NOTE.Visible = True
        Else
            txtIRREGULAR_HOURS_NOTE.Text = ""
            txtIRREGULAR_HOURS_NOTE.Visible = False
        End If
    End Sub

    Private Sub chkBROKER_CheckedChanged(sender As Object, e As EventArgs) Handles chkBROKER.CheckedChanged
        If chkBROKER.Checked Then
            txtBROKER_NOTE.Visible = True
        Else
            txtBROKER_NOTE.Text = ""
            txtBROKER_NOTE.Visible = False
        End If
    End Sub

    Private Sub chkAPPOINTMENT_REQUIRED_NOTE_CheckedChanged(sender As Object, e As EventArgs) Handles chkAPPOINTMENT_REQUIRED.CheckedChanged
        If chkAPPOINTMENT_REQUIRED.Checked Then
            txtAPPOINTMENT_REQUIRED_NOTE.Visible = True
        Else
            txtAPPOINTMENT_REQUIRED_NOTE.Text = ""
            txtAPPOINTMENT_REQUIRED_NOTE.Visible = False
        End If
    End Sub

    Private Sub bindControl(ByVal CC As String)
        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text & String.Empty
            Dim Filter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CC)
            Dim rowARTCUSTQ As DataRow = dst.Tables("ARTCUSTQ").Select(Filter).FirstOrDefault
            If IsNothing(rowARTCUSTQ) Then
                rowARTCUSTQ = dst.Tables("ARTCUSTQ").NewRow
                rowARTCUSTQ.Item("CUST_CODE") = CUST_CODE
                rowARTCUSTQ.Item("CUST_ADDR_CODE") = CC
                dst.Tables("ARTCUSTQ").Rows.Add(rowARTCUSTQ)
            End If
            If IsDate(rowARTCUSTQ.Item("LAST_DATE").ToString) Then
                dteLAST_DATE.Value = CDate(rowARTCUSTQ.Item("LAST_DATE").ToString)
            Else
                dteLAST_DATE.Value = Null
            End If
            txtLAST_OPER.Value = rowARTCUSTQ.Item("LAST_OPER").ToString

            chkRESIDENTIAL_ORDR.Checked = rowARTCUSTQ.Item("RESIDENTIAL_ORDR").ToString & String.Empty = "1"
            chkINSIDE_REQ.Checked = rowARTCUSTQ.Item("INSIDE_REQ").ToString & String.Empty = "1"
            chkGATE_LIFT_REQ.Checked = rowARTCUSTQ.Item("GATE_LIFT_REQ").ToString & String.Empty = "1"

            chkLIMITED_ACCESS.Checked = rowARTCUSTQ.Item("LIMITED_ACCESS").ToString & String.Empty = "1"
            txtLIMITED_ACCESS_NOTE.Text = rowARTCUSTQ.Item("LIMITED_ACCESS_NOTE").ToString & String.Empty

            chkIRREGULAR_HOURS.Checked = rowARTCUSTQ.Item("IRREGULAR_HOURS").ToString & String.Empty = "1"
            txtIRREGULAR_HOURS_NOTE.Text = rowARTCUSTQ.Item("IRREGULAR_HOURS_NOTE").ToString & String.Empty

            chkBROKER.Checked = rowARTCUSTQ.Item("BROKER").ToString & String.Empty = "1"
            txtBROKER_NOTE.Text = rowARTCUSTQ.Item("BROKER_NOTE").ToString & String.Empty

            chkAPPOINTMENT_REQUIRED.Checked = rowARTCUSTQ.Item("APPOINTMENT_REQUIRED").ToString & String.Empty = "1"
            txtAPPOINTMENT_REQUIRED_NOTE.Text = rowARTCUSTQ.Item("APPOINTMENT_REQUIRED_NOTE").ToString & String.Empty

        End If
    End Sub

    Private Sub btnVerifyShipToInfo_Click(sender As Object, e As EventArgs) Handles btnVerifyShipToInfo.Click
        If btnVerifyShipToInfo.Text = "Edit Ship-To Info" Then
            btnVerifyShipToInfo.Text = "Save Ship-To Info"
            dteLAST_DATE.Value = CDate(Now().ToShortDateString)
            txtLAST_OPER.Value = ASCMAIN1.USER_ID
            chkRESIDENTIAL_ORDR.Enabled = True
            chkINSIDE_REQ.Enabled = True
            chkGATE_LIFT_REQ.Enabled = True
            chkLIMITED_ACCESS.Enabled = True
            txtLIMITED_ACCESS_NOTE.Enabled = True
            chkIRREGULAR_HOURS.Enabled = True
            txtIRREGULAR_HOURS_NOTE.Enabled = True
            chkBROKER.Enabled = True
            txtBROKER_NOTE.Enabled = True
            chkAPPOINTMENT_REQUIRED.Enabled = True
            txtAPPOINTMENT_REQUIRED_NOTE.Enabled = True
        Else
            btnVerifyShipToInfo.Text = "Edit Ship-To Info"
            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text & String.Empty
            Dim CC As String = grdARTCUST2.ActiveRow.Cells("CUST_ADDR_CODE").Value & String.Empty
            Dim Filter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CC)
            Dim rowARTCUSTQ As DataRow = dst.Tables("ARTCUSTQ").Select(Filter).FirstOrDefault
            If Not IsNothing(rowARTCUSTQ) Then
                rowARTCUSTQ.Item("RESIDENTIAL_ORDR") = IIf(chkRESIDENTIAL_ORDR.Checked = True, "1", "0")
                rowARTCUSTQ.Item("INSIDE_REQ") = IIf(chkINSIDE_REQ.Checked = True, "1", "0")
                rowARTCUSTQ.Item("GATE_LIFT_REQ") = IIf(chkGATE_LIFT_REQ.Checked = True, "1", "0")
                rowARTCUSTQ.Item("LIMITED_ACCESS") = IIf(chkLIMITED_ACCESS.Checked = True, "1", "0")
                rowARTCUSTQ.Item("LIMITED_ACCESS_NOTE") = txtLIMITED_ACCESS_NOTE.Text & String.Empty
                rowARTCUSTQ.Item("IRREGULAR_HOURS") = IIf(chkIRREGULAR_HOURS.Checked = True, "1", "0")
                rowARTCUSTQ.Item("IRREGULAR_HOURS_NOTE") = txtIRREGULAR_HOURS_NOTE.Text & String.Empty
                rowARTCUSTQ.Item("BROKER") = IIf(chkBROKER.Checked = True, "1", "0")
                rowARTCUSTQ.Item("BROKER_NOTE") = txtBROKER_NOTE.Text & String.Empty
                rowARTCUSTQ.Item("APPOINTMENT_REQUIRED") = IIf(chkAPPOINTMENT_REQUIRED.Checked = True, "1", "0")
                rowARTCUSTQ.Item("APPOINTMENT_REQUIRED_NOTE") = txtAPPOINTMENT_REQUIRED_NOTE.Text & String.Empty

                chkRESIDENTIAL_ORDR.Enabled = False
                chkINSIDE_REQ.Enabled = False
                chkGATE_LIFT_REQ.Enabled = False
                chkLIMITED_ACCESS.Enabled = False
                txtLIMITED_ACCESS_NOTE.Enabled = False
                chkIRREGULAR_HOURS.Enabled = False
                txtIRREGULAR_HOURS_NOTE.Enabled = False
                chkBROKER.Enabled = False
                txtBROKER_NOTE.Enabled = False
                chkAPPOINTMENT_REQUIRED.Enabled = False
                txtAPPOINTMENT_REQUIRED_NOTE.Enabled = False
            End If
        End If
    End Sub

    Private Sub ViewTaxDoc()
        TAX_ID_DOC = TAX_ID_DOC.Replace(TAX_ID & "-", "")

        Dim UserName As String = "regency-rib"
        Dim Password As String = "joydHUJ3"
        Dim RemoteHost As String = "regency-rib.com" '69.39.227.201
        Dim RemotePath As String = "www/customers"
        'Dim ServerFilePath As String = "S:\RGI\Archive\Shopsite\"

        Dim FTP_FOLDER As String = "customers\" & TAX_ID & "\"
        Dim TempFolder As String = ASCMAIN1.Folders("Temp").ToString
        If Not TempFolder.EndsWith("\") Then
            TempFolder = TempFolder & "\"
        End If
        Dim LocalFile As String = TempFolder & TAX_ID_DOC
        Dim ErrMsg As New Text.StringBuilder With {.Length = 0}
        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                If System.IO.File.Exists(LocalFile) Then
                    System.IO.File.Delete(LocalFile)
                End If
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = UserName
                .Password = Password
                .RemoteHost = RemoteHost
                .RemotePath = RemotePath & "/tax_id/" & TAX_ID
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = LocalFile
                .RemoteFile = TAX_ID_DOC
                .Overwrite = False
                If Not .FileExists() Then
                    ErrMsg.AppendLine("File Not Found On Shopsite")
                    .Logoff()
                Else
                    .Download()
                    .Logoff()
                End If
            Catch ex As Exception
                ErrMsg.AppendLine(ex.Message.ToString)
                FtpShopSite.Logoff()
            End Try
        End With
        If ErrMsg.Length > 0 Then
            MsgBox(ErrMsg.ToString, vbExclamation, "Problems Fetching Document.")
        Else
            Show_Document(LocalFile)
        End If
    End Sub

    Private Sub btnWebTaxId_Click(sender As Object, e As EventArgs) Handles btnWebTaxId.Click
        ViewTaxDoc()
    End Sub

    Public Overrides Function Remote_Control(
                    ByVal command As String,
                    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                'Click_Command("Done")

            Case "View", "Edit"
                Absx1.txtFor("CUST_CODE").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

End Class