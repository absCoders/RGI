Public Class ICTSTYL1

    Dim sqlICTSTYC1 As String = ""
    Dim sqlICTSTYL3 As String = ""
    Dim sqlICTSTYL4 As String = ""
    Dim sqlICTSTYL5 As String = ""
    Dim sqlICTSTYLC As String = ""
    Dim sqlICTDUTY4 As String = ""
    Dim sqlICTSTYV1 As String = ""
    Dim sqlICTSTYC2 As String = ""
    Dim sqlICTSTYC4 As String = ""
    Dim MAX_SIZES As Int32 = 20     ' this may be set to a higher value if need be
    Dim SO_PARM_UPC_VENDOR_ID As String = ""
    Dim rowSOTPARM1 As DataRow

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            btnAutomatic.Visible = True
            btnGenerateUPCs.Visible = True
        End If

        Get_PARM("ICTPARM1")
        rowSOTPARM1 = LookUp("SOTPARM1", "Z")
        SO_PARM_UPC_VENDOR_ID = rowSOTPARM1.Item("SO_PARM_UPC_VENDOR_ID") & ""

        AUDIT.Add("ICTSTYC1", "*")
        AUDIT.Add("ICTSTYLD", "*")

        AUDIT.Add("ICTSTYL3", "*")
        'AUDIT.Add("ICTSTYL4", "E")
        'AUDIT.Add("ICTSTYL5", "E")
        'AUDIT.Add("ICTSTYLC", "E")

        AUDIT.Add("ICTSTYV1", "*")

        'AUDIT.Add("ICTSTYC2", "E")
        'AUDIT.Add("ICTSTYC3", "E")
        'AUDIT.Add("ICTSTYC4", "E")

        With dst
            sqlICTSTYC1 = "Select ICTSTYC1.*, ICTCOLR1.COLOR_DESC, ICTTHEME.THEME_DESC" _
            & " from ICTCOLR1,ICTSTYC1,ICTtheme" _
            & " where ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE" _
            & " and ICTTHEME.THEME_CODE (+) = ICTSTYC1.THEME_CODE"
            ASCMAIN1.sql = sqlICTSTYC1 _
            & "  and ICTSTYC1.STYLE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, True, "V")
            .Tables("ICTSTYC1").Columns.Add("COLORS", GetType(List(Of String)))

            sqlICTSTYL3 = "Select ICTSTYL3.* " _
            & " from ICTSTYL3" _
            & " where ICTSTYL3.STYLE_CODE = :PARM1"
            ASCMAIN1.sql = sqlICTSTYL3
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, True, "V")

            sqlICTSTYL4 = "Select ICTSTYL4.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from ICTSTYL1,ICTCOLR1,ICTSTYL4" _
            & " where ICTSTYL1.STYLE_CODE = ICTSTYL4.STYLE_CODE" _
            & "   and ICTCOLR1.COLOR_CODE = ICTSTYL4.COLOR_CODE"
            ASCMAIN1.sql = sqlICTSTYL4 _
            & "  and ICTSTYL4.STYLE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTSTYL4", "**", 0, True, "V")

            sqlICTSTYL5 = "Select ICTSTYL5.*, ICTPANTC.PANTONE_DESC, ICTPANTC.RGB" _
             & " from ICTPANTC,ICTSTYL5" _
             & " where ICTPANTC.PANTONE_CODE = ICTSTYL5.PANTONE_CODE"
            ASCMAIN1.sql = sqlICTSTYL5 _
            & "  and ICTSTYL5.STYLE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTSTYL5", "**", 0, True, "V")

            sqlICTSTYLC = "Select ICTSTYLC.*, ICTDUTY1.DUTY_RATE_DESC" _
             & " from ICTDUTY1,ICTSTYLC" _
             & " where ICTDUTY1.DUTY_RATE_CODE = ICTSTYLC.DUTY_RATE_CODE"
            ASCMAIN1.sql = sqlICTSTYLC _
            & "  and ICTSTYLC.STYLE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTSTYLC", "**", 0, True, "V")


            sqlICTDUTY4 = "Select ICTDUTY4.*, ICTDUTY1.DUTY_RATE_DESC" _
             & " from ICTDUTY1,ICTDUTY4" _
             & " where ICTDUTY1.DUTY_RATE_CODE = ICTDUTY4.DUTY_RATE_CODE"
            ASCMAIN1.sql = sqlICTDUTY4 _
            & "  and ICTDUTY4.DUTY_RATE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTDUTY4", "**", 0, True, "V")

            sqlICTSTYV1 = "Select ICTSTYV1.*, APTVEND1.VEND_NAME" _
            & " from APTVEND1,ICTSTYV1" _
            & " where APTVEND1.VEND_CODE (+) = ICTSTYV1.VEND_CODE"
            ASCMAIN1.sql = sqlICTSTYV1 _
            & "  and ICTSTYV1.STYLE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTSTYV1", "**", 0, True, "V")

            With .Tables("ICTSTYV1")
                .Columns.Add("FUTURE_COST", GetType(System.Decimal))
                .Columns.Add("FUTURE_COST_DATE", GetType(System.DateTime))
            End With

            sqlICTSTYC2 = "Select ICTSTYC2.* " _
            & " from ICTSTYC2" _
            & " where ICTSTYC2.STYLE_CODE = :PARM1"
            ASCMAIN1.sql = sqlICTSTYC2
            Create_TDA(.Tables.Add, "ICTSTYC2", "**", 0, True, "V")

            sqlICTSTYC2 = "Select ICTSTYC3.* " _
            & " from ICTSTYC3" _
            & " where ICTSTYC3.STYLE_CODE = :PARM1"
            ASCMAIN1.sql = sqlICTSTYC2
            Create_TDA(.Tables.Add, "ICTSTYC3", "**", 0, True, "V")

            sqlICTSTYC4 = "Select ICTSTYC4.* " _
            & " from ICTSTYC4" _
            & " where ICTSTYC4.STYLE_CODE = :PARM1"
            ASCMAIN1.sql = sqlICTSTYC4
            Create_TDA(.Tables.Add, "ICTSTYC4", "**", 0, True, "V")

            Create_TDA(.Tables.Add("ICTSTYL1_NEW"), "ICTSTYL1", "*")

            .Tables.Add("ICTSTYCI")
            With .Tables("ICTSTYCI")
                .Columns.Add("LINE_NO", GetType(System.Int32))
                .Columns.Add("INSTRUCTION")
                .PrimaryKey = New DataColumn() {.Columns("LINE_NO")}
                .Rows.Add(New Object() {1, "Use the same color code in the UPC grid below as the one you have selected to the left, unless the one selected is AST and you want to define UPC's for individual colors."})
                .Rows.Add(New Object() {2, "Use the 'No Size' UPC column unless you want to define UPC's for individual Sizes.  To add Colors and Sizes, right click on the grid below."})
            End With


            .Tables.Add("ICTSTYCX")
            With .Tables("ICTSTYCX")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("COLOR_CODE_UPC")
                .Columns.Add("COLOR_DESC_UPC")
                For i As Integer = 0 To MAX_SIZES
                    .Columns.Add("UPC_CODE_" & Format(i, "00"), GetType(System.String))
                    .Columns.Add("PPK_QTY_" & Format(i, "00"), GetType(System.Int64))
                Next
                .PrimaryKey = New DataColumn() {.Columns("STYLE_CODE"), _
                                                .Columns("COLOR_CODE"), _
                                                .Columns("COLOR_CODE_UPC")}
            End With
            Create_Relation("ICTSTYC1", "ICTSTYCX", "STYLE_CODE,COLOR_CODE")

            Create_Relation("ICTSTYC1", "ICTSTYL4", "STYLE_CODE,COLOR_CODE")
            Create_Relation("ICTSTYC1", "ICTSTYL5", "STYLE_CODE,COLOR_CODE")

            ASCMAIN1.sql = "Select Distinct SIZE_INDEX, SIZE_CODE FROM ICTSTYC3" _
                & " where STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYCS", "**", 0, False, "V", 1)
            .Tables("ICTSTYCS").Columns.Add("POSITION", GetType(System.Int32))

            ASCMAIN1.sql = "Select * from ICTUOMF1"
            Create_TDA(.Tables.Add, "ICTUOMF1", "**", 0, False)

            ASCMAIN1.sql = "SELECT ICTSTYLD.*, ICTSTYLM.PACK_DESC" _
            & " FROM ICTSTYLD, ICTSTYLM" _
            & " WHERE ICTSTYLD.PACK_CODE = ICTSTYLM.PACK_CODE" _
            & "  and ICTSTYLD.STYLE_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTSTYLD", "**", 0, True, "V")

            If ASCMAIN1.CLIENT = "VAN" Then
                Create_TDA(.Tables.Add, "ICTSTYLS", "*")
            End If
        End With

        Fill_Records("ICTUOMF1")

        grdICTSTYC1.DataSource = dst.Tables("ICTSTYC1")
        grdICTSTYL3.DataSource = dst.Tables("ICTSTYL3")
        grdICTSTYL4.DataSource = dst.Tables("ICTSTYL4")
        grdICTSTYL5.DataSource = dst.Tables("ICTSTYL5")
        grdICTSTYLC.DataSource = dst.Tables("ICTSTYLC")
        grdICTDUTY4.DataSource = dst.Tables("ICTDUTY4")
        grdICTSTYV1.DataSource = dst.Tables("ICTSTYV1")
        grdICTSTYCX.DataSource = dst.Tables("ICTSTYCX")
        grdICTSTYCI.DataSource = dst.Tables("ICTSTYCI")
        grdICTSTYLD.DataSource = dst.Tables("ICTSTYLD")
        Sort_grdColumns(grdICTSTYCI, "LINE_NO", True)

        If ASCMAIN1.CLIENT = "VAN" Then
            For I As Integer = 2 To 24
                Dim txt As New UltraWinEditors.UltraTextEditor
                txt.Name = "txtSIZE_" & Format(I, "00")
                txt.Parent = splColorSize.Panel2
                txt.Left = txtSIZE_01.Left + txtSIZE_01.Width * (I - 1)
                txt.Width = txtSIZE_01.Width
                txt.Top = txtSIZE_01.Top
                Absx1.SetABSColumnName(txt, "SIZE_" & Format(I, "00"))
                Absx1.SetABSTableName(txt, "ICTSTYLS")

                Absx1.dicCOLUMN_NAME.Add("ICTSTYLS." & "SIZE_" & Format(I, "00"), txt)

                txt.Visible = True

                Dim num As New UltraWinEditors.UltraNumericEditor
                num.Name = "numQTY_" & Format(I, "00")
                num.Parent = splColorSize.Panel2
                num.Left = numQTY_01.Left + numQTY_01.Width * (I - 1)
                num.Width = numQTY_01.Width
                num.Top = numQTY_01.Top
                num.AlwaysInEditMode = True
                num.PromptChar = ""
                num.MaxValue = 100
                Absx1.SetABSColumnName(num, "QTY_" & Format(I, "00"))
                Absx1.SetABSTableName(num, "ICTSTYLS")

                Absx1.dicCOLUMN_NAME.Add("ICTSTYLS." & "QTY_" & Format(I, "00"), num)

                num.Visible = True

            Next

            Bind_Controls(splColorSize.Panel2, "ICTSTYLS")

        End If

        grdICTSTYCX.DisplayLayout.UseFixedHeaders = True
        grdICTSTYCX.DisplayLayout.Override.AllowColMoving = UltraWinGrid.AllowColMoving.NotAllowed
        With grdICTSTYCX.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Hidden = True
            .Columns("COLOR_CODE").Hidden = True
            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add("CODES")
            G.Header.Fixed = True
            G.Width = 200
            G.Header.Caption = ""
            With .Columns("COLOR_CODE_UPC")
                .Width = 60
                .Header.Caption = "Color"
                .Group = G
                .Hidden = False
                '.Style = UltraWinGrid.ColumnStyle.EditButton
                '.ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            End With
            With .Columns("COLOR_DESC_UPC")
                .Width = 140
                .Header.Caption = "Description"
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .Group = G
                .Hidden = False
                .Header.Fixed = True
            End With
            For i As Integer = 0 To MAX_SIZES
                G = .Groups.Add("SIZE_" & Format(i, "00"))
                G.Header.Fixed = (i = 0)
                G.Width = 170
                G.Header.Appearance.TextHAlign = HAlign.Center
                If i = 0 Then G.Header.Caption = "No Size"
                With .Columns("UPC_CODE_" & Format(i, "00"))
                    .Group = G
                    .Hidden = False
                    .Width = 130
                    .Header.Caption = "UPC"
                    .Style = UltraWinGrid.ColumnStyle.EditButton
                    .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                    .CellButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "barcode")
                End With
                With .Columns("PPK_QTY_" & Format(i, "00"))
                    .Group = G
                    .Hidden = False
                    .Width = 40
                    .Header.Caption = "Ppk"
                End With
            Next
        End With

        With grdICTSTYC1.DisplayLayout.Bands(0)
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("COLOR_DESC").Header.Fixed = True
            .Columns("UPC_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
            .Columns("UPC_CODE").ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
            .Columns("UPC_CODE").CellButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "barcode")
            .Columns("HIDE_COLOR_3PL").Hidden = True ' EVEN FOR NYA

            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                .Columns("UPC_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("HIDE_FROM_CAT").Hidden = True
                .Columns("CATALOG_SELECTION_CODE").Hidden = True
                .Columns("NRF_SIZE_CODE").Hidden = True
                .Columns("NRF_COLOR_CODE").Hidden = True
            End If

            If ASCMAIN1.CLIENT = "NYA" Then
            Else
                If .Columns.Contains("UPC_CODE_INNER") Then
                    .Columns("UPC_CODE_INNER").Hidden = True
                    .Columns("UPC_CODE_CASE").Hidden = True
                End If
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                .Columns("THEME_CODE").Hidden = False
                .Columns("THEME_DESC").Hidden = False
                UltraTabControl1.Tabs.Item(3).Visible = True
            ElseIf ASCMAIN1.CLIENT = "NYA" Then
                .Columns("THEME_CODE").Hidden = False
                .Columns("THEME_CODE").Header.Caption = "NG Item"
                .Columns("THEME_DESC").Hidden = True
                UltraTabControl1.Tabs.Item(3).Visible = False
            Else
                .Columns("THEME_CODE").Hidden = True
                .Columns("THEME_DESC").Hidden = True
                UltraTabControl1.Tabs.Item(3).Visible = False
            End If

            .Columns("UPC_CODE").Hidden = (ASCMAIN1.CLIENT = "VAN")

        End With

        With grdICTSTYV1.DisplayLayout.Bands(0)
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VEND_NAME").Header.Fixed = True
            .Columns("PO_COST").MaskInput = "nnnn.nnnnnn"
            .Columns("PO_COST").Format = "##0.000000"
        End With

        Create_Summary(grdICTSTYL3, "ATTR_CODE", "Count")
        Create_Summary(grdICTSTYV1, "VEND_CODE", "Count")


        ASCMAIN1.Add_Value_List(grdICTSTYC1, "STYLE_COLOR_STATUS", Nothing, New String() {":", "A:Active", "D:Discontinued", "N:Do Not Re-Order"})

        Dim dt As DataTable = ASCDATA1.GetDataTable("Select STYLE_CLASS_CODE, STYLE_CLASS_DESC from ICTCLAS1 order by STYLE_CLASS_DESC")
        Dim dvw As DataView = dt.DefaultView
        dvw.Sort = "STYLE_CLASS_CODE"
        cbeSTYLE_CLASS_CODE.DataSource = dvw

        lblSIZE_SCALE.Text = "Size/Color" & vbCrLf & "Breakdown"

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
        Else
            lblCMT_NO.Visible = False
            txtCMT_NO.Visible = False
            txtCMT_NOTES.Visible = False

            tabICTSTYC2.Tabs("Size, Color, UPC Codes").Visible = False
            'splICTSTYC1.Panel2Collapsed = True

            lblSUB_BODY_CODE.Visible = False
            txtSUB_BODY_CODE.Visible = False
            txtSUB_BODY_DESC.Visible = False

            lblSIZE_CODE.Visible = False
            txtSIZE_CODE.Visible = False

            grdICTSTYCI.Visible = False
            splICTSTYCI.Panel2Collapsed = True
        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            splICTSTYC1.Panel2Collapsed = True
            Absx1.optFor("FASHION_PROMO").Visible = False
            Absx1.numFor("STYLE_SO_QTY_MIN").Visible = True
            lblSTYLE_SO_QTY_MIN.Visible = True

            splICTSTYCI.Panel2Collapsed = False
            ' splICTSTYCI.SplitterDistance = splICTSTYCI.Width * 0.6
            Size_splICTSTYCI()

            imgSTYLE.Visible = False
            btnIMAGE_NAME.Visible = False
            txtIMAGE_NAME.Visible = False

            lblCASE_CUBE_UM.Text = "CFT"

            lblSIZE_CODE.Visible = True
            txtSIZE_CODE.Visible = True

            picStyleColor.Visible = True
            picStyleColor2.Visible = True
            UltraExplorerBar1.Groups.Item("Image").Visible = True
        Else
            lblUnitsInner.Text = "Units / Inner"
            lblInner2.Visible = False
            numInner2.Visible = False
            lblUnitsInner2.Visible = False
            picStyleColor.Visible = False
            picStyleColor2.Visible = False
            UltraExplorerBar1.Groups.Item("Image").Visible = False
            Absx1.numFor("STYLE_SO_QTY_MIN").Visible = False
            lblSTYLE_SO_QTY_MIN.Visible = False
        End If

        chkSet.Visible = ASCMAIN1.CLIENT = "NYA"

        lblCUST_STYLE_CODE.Visible = (ASCMAIN1.CLIENT = "NYA")
        txtCUST_STYLE_CODE.Visible = (ASCMAIN1.CLIENT = "NYA")

        'lblSTYLE_PROMO_PRICE.Visible = (ASCMAIN1.CLIENT = "RGI")
        'numSTYLE_PROMO_PRICE.Visible = (ASCMAIN1.CLIENT = "RGI")
        'No Longer Used At RGI
        lblSTYLE_PROMO_PRICE.Visible = False
        numSTYLE_PROMO_PRICE.Visible = False

        lblFABRIC_CODE.Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")
        txtFABRIC_CODE.Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")
        txtFABRIC_DESC.Visible = (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")

        lblSTYLE_CODE_PLM.Visible = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
        txtSTYLE_CODE_PLM.Visible = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
        Absx1.txtFor("DESIGN_STYLE_NO").Visible = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")

        lblSAFETY_STOCK.Visible = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
        numSAFETY_STOCK.Visible = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")

        lblCARTONS_PER_UNIT.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")
        numCARTONS_PER_UNIT.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")
        chkEXCLUSIVE_STYLE.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")
        lblWHSE_MESSAGE.Visible = ASCMAIN1.CLIENT = "RGI"
        txtWHSE_MESSAGE.Visible = ASCMAIN1.CLIENT = "RGI"

        grdICTSTYL3.Visible = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")

        With grdICTSTYL4.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE_COMP").Hidden = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
            .Columns("COLOR_CODE_COMP").Hidden = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
            .Columns("STYLE_DESC").Hidden = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
            .Columns("COLOR_DESC").Hidden = (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA")
        End With

        Create_Summary(grdICTSTYL4, "COMP_LNO", "Count")
        Create_Summary(grdICTSTYL4, "QTY_PER_PPK")

        grdICTSTYLC.Visible = (ASCMAIN1.CLIENT = "NYA")
        grdICTDUTY4.Visible = (ASCMAIN1.CLIENT = "RGI")
        If ASCMAIN1.CLIENT = "RGI" Then
            grdICTDUTY4.Left = grdICTSTYLC.Left
        End If

        If ASCMAIN1.CLIENT = "NYA" Then
            Absx1.chkFor("REQUIRES_EXP_DATE").Visible = True
            lblCASE_CUBE_UM.Text = "CIN"
            splICTSTYC1.SplitterDistance = 100
            '  grdICTSTYC1.DisplayLayout.Bands(0).Columns("UPC_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            Absx1.chkFor("REQUIRES_EXP_DATE").Visible = True
            tabICTSTYC2.Tabs("Component Styles").Visible = False
            tabICTSTYC2.Tabs("Pantone Colors").Visible = False

            lblSTYLE_GROUP_CODE.Visible = False
            Absx1.txtFor("STYLE_GROUP_CODE").Visible = False
            Absx1.txtFor("STYLE_GROUP_DESC").Visible = False

            lblROYALTY_CODE.Visible = False
            Absx1.txtFor("ROYALTY_CODE").Visible = False
            Absx1.txtFor("ROYALTY_DESC").Visible = False

            chkSTYLE_HIDE_FROM_3PL.Visible = False
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            lblSTYLE_GROUP_CODE.Text = "Silhouette"
            lblSTYLE_GROUP_CODE.Visible = True
            Absx1.txtFor("STYLE_GROUP_CODE").Visible = True
            Absx1.txtFor("STYLE_GROUP_DESC").Visible = True
        End If


        lblLIST_CALC_CODE.Visible = (ASCMAIN1.CLIENT = "RGI")
        txtLIST_CALC_CODE.Visible = (ASCMAIN1.CLIENT = "RGI")
        txtLIST_CALC_DESC.Visible = (ASCMAIN1.CLIENT = "RGI")
        cmdCalculateList.Visible = (ASCMAIN1.CLIENT = "RGI")

        splColorSize.Panel2Collapsed = Not (ASCMAIN1.CLIENT = "VAN")

        cmdFixSizesColors.Visible = False
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            If ASCMAIN1.USER_ID = "wjz" Or ASCMAIN1.USER_ID = "marilyn" Or ASCMAIN1.USER_ID = "joann" Then
                cmdFixSizesColors.Visible = True
            End If
        End If
    End Sub

    Sub Size_splICTSTYCI()
        Dim D As Decimal = 50
        For Each gcol As UltraWinGrid.UltraGridColumn In grdICTSTYC1.DisplayLayout.Bands(0).Columns
            If Not gcol.Hidden Then D += gcol.Width
        Next
        splICTSTYCI.SplitterDistance = D
    End Sub

    Public Overrides Function RemoteProcedureCall( _
    ByVal command As String, _
    ByVal keys As Dictionary(Of String, Object)) As Object

        Dim return_key As Object = Nothing
        '  Application.DoEvents()

        Select Case command
            Case "Fill from AT"
                Fill_from_AT(keys)

        End Select

        Return return_key
    End Function
 
    Sub Fill_from_AT(keys As Dictionary(Of String, Object))
        For Each key As String In keys.Keys
            Absx1.txtFor(key).Text = keys(key)
        Next
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        MyBase.Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYC1, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes", "Generate UPCs", "Allow Edit to this UPC")
        Load_Popup_Menu(grdICTSTYL3, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdICTSTYL4, "SSSB", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdICTSTYL5, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdICTSTYLC, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdICTSTYV1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdICTSTYCX, "SSBB", "Show Filter", "Show GroupBox", "Add Size", "Add Colors")
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
        'if not new or edit - hide add codes

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdICTSTYCX"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Colors"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New") AndAlso grdICTSTYC1.ActiveRow IsNot Nothing AndAlso grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value = "AST"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Size"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
            Case "grdICTSTYC1", "grdICTSTYV1", "grdICTSTYL3", "grdICTSTYL5", "grdICTSTYLC"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
                If grd.Name = "grdICTSTYC1" Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Generate UPCs"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
                    tlb_btn = DirectCast(tlb_pop.Tools("Allow Edit to this UPC"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "Edit")
                End If
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
                If grd.Name = "grdICTSTYV1" Then
                    Add_Codes(grdICTSTYV1, "APTVEND1", "VEND_CODE", "Vendors")
                ElseIf grd.Name = "grdICTSTYC1" Then
                    Add_Codes(grdICTSTYC1, "ICTCOLR1", "COLOR_CODE", "Colors")
                    Add_Single_Colors()
                    Setup_ICTSTYC1()
                ElseIf grd.Name = "grdICTSTYL3" Then
                    Add_Codes(grdICTSTYL3, "ICTATTR1", "ATTR_CODE", "Attributes")
                ElseIf grd.Name = "grdICTSTYL5" Then
                    Add_Codes(grdICTSTYL5, "ICTPANTC", "PANTONE_CODE", "Pantone Colors")
                ElseIf grd.Name = "grdICTSTYLC" Then
                    Add_Codes(grdICTSTYLC, "TATCNTRY", "COUNTRY_CODE", "Country Codes")
                End If

            Case "Add Colors"
                Add_Colors()

            Case "Add Size"
                Add_Size()

            Case "Generate UPCs"
                Generate_UPCs()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Allow Edit to this UPC"
                With grdICTSTYC1.DisplayLayout.Bands(0).Columns("UPC_CODE")
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    grdICTSTYC1.ActiveCell = grdICTSTYC1.ActiveRow.Cells("UPC_CODE")
                    grdICTSTYC1.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                End With

        End Select



    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "DUTY_RATE_CODE"
                SHOW_DUTY_EXCEPTIONS()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        If EntryMode = "" Then
            MyBase.txt_EditorButtonClick_Special(txtctl)
        Else
            Select Case Absx1.GetABSColumnName(txtctl)
                Case "DUTY_RATE_CODE"
                    SHOW_DUTY_EXCEPTIONS()
            End Select
        End If

    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "DUTY_RATE_CODE"
                SHOW_DUTY_EXCEPTIONS()
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                If ASCMAIN1.CLIENT = "NYA" Then
                    If txtSTYLE_CODE_PLM_SOURCE.Tag & "" = "" Then
                        EMsg &= vbCr & "You Must Use the Create Style from PLM function to add a New Style"
                    End If
                End If
            Case "Edit"
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim SC As String = Absx1.txtFor("STYLE_CODE").Text
                    Dim PARTAIALSTYLE As String = PARTIALSTYLE(SC)
                    If PARTAIALSTYLE.Length > 0 Then
                        Absx1.txtFor("STYLE_CODE").Text = PARTAIALSTYLE
                    End If
                End If
            Case "View"
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim SC As String = Absx1.txtFor("STYLE_CODE").Text
                    Dim PARTAIALSTYLE As String = PARTIALSTYLE(SC)
                    If PARTAIALSTYLE.Length > 0 Then
                        Absx1.txtFor("STYLE_CODE").Text = PARTAIALSTYLE
                    End If
                End If
            Case "Update"
                If Absx1.optFor("STYLE_STATUS").Value & "" = "" Then
                    EMsg &= vbCr & "You Must Select a Value for Style Status"
                Else
                    If Absx1.optFor("STYLE_STATUS").Value & "" <> "A" Then
                        Dim STYLE_STATUS As String = Absx1.optFor("STYLE_STATUS").Value & ""
                        If dst.Tables("ICTSTYC1").Select("STYLE_COLOR_STATUS <> '" & STYLE_STATUS & "'").Length <> 0 Then
                            If MsgBox("Since Style Status is no longer Active, some Colors will have Status Changed to " & STYLE_STATUS, _
                                   MsgBoxStyle.OkCancel, "Verification") = MsgBoxResult.Cancel Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                If Absx1.txtFor("STYLE_UOM").Text = "" Then
                    EMsg &= vbCr & "You Must Select a Value for Style Unit of Measure"
                    'Else
                    '    If dst.Tables("ICTUOMF1").Rows.Find(Absx1.txtFor("STYLE_UOM").Text) Is Nothing Then
                    '        EMsg &= vbCr & "Invalid Value specified for Unit of Measure"
                    '    End If
                End If

                If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                    If Val(Absx1.numFor("STYLE_SO_QTY_MIN").Text & "") <= 0 Then
                        EMsg &= vbCr & "Minimum Sales Order Quantity Must Be 1 or Higher."
                    End If
                    If Val(Absx1.numFor("STYLE_PRICE").Text & "") <= 0 Then
                        EMsg &= vbCr & "Style Price Must Be Greater Than 0."
                    End If
                End If

                Dim SUB_UNIT_PACK_QTY As Int16 = Val(Absx1.numFor("SUB_UNIT_PACK_QTY").Value & "")
                If SUB_UNIT_PACK_QTY = 0 Then
                    EMsg &= vbCr & "Invalid Value for Sub Unit Pack Qty"
                End If

                If ASCMAIN1.CLIENT = "NYA" Then
                    If SUB_UNIT_PACK_QTY = 1 And chkSet.Checked Then EMsg &= vbCr & "Set Qty must not be 1"
                    If SUB_UNIT_PACK_QTY > 1 And Not chkSet.Checked Then EMsg &= vbCr & "Set Qty > 1 must be marked as a set"
                End If

                If Absx1.txtFor("STYLE_DESC").Text = "" Then
                    EMsg &= vbCr & "You Must Select a Value for Style Description"
                ElseIf rowASFBASE1.Item("STYLE_DESC").Contains(vbCrLf) Or rowASFBASE1.Item("STYLE_DESC").Contains(vbLf) Then
                    EMsg &= vbCr & "Style Description cannot have Multiple Lines"
                Else
                    '    Dim row1 As DataRow = dst.Tables("ICTSTYL1").Rows(0)
                    '    If row1.Item("STYLE_DESC").Contains(vbCrLf) Or row1.Item("STYLE_DESC").Contains(vbLf) Then
                    '        EMsg &= vbCr & "Style Description cannot have Multiple Lines"
                    '    End If

                    If ASCMAIN1.CLIENT = "NYA" Then
                        'Dim rx As String = "[^a-zA-Z0-9 .,_$-]" ' Allow Upper/Lower case, numbers, space, dot, comma, underscore and dash
                        'Dim rx As String = "[^a-zA-Z0-9 .,$-]" ' Allow Upper/Lower case, numbers, space, dot, comma and dash
                        'Dim rx As String = "[^a-zA-Z0-9 .$]" ' Allow Upper/Lower case, numbers, space, dot
                        Dim rx As String = "[^-'" & Chr(34) & "A-Za-z0-9% &$!#.,+=_()@\/:;]" ' Allow Upper/Lower case, numbers, other characters
                        Dim r As New System.Text.RegularExpressions.Regex(rx)
                        If r.IsMatch(Absx1.txtFor("STYLE_DESC").Text) Then
                            EMsg &= vbCr & "Style Description has Special Characters which are not allowed"
                        End If
                    End If

                End If

                'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

                'Dim rowSOTSREP1 = LookUp("SOTSREP1", Absx1.txtFor("STYLE_CODE").Text)
                'If rowSOTSREP1 Is Nothing Then
                '    EMsg &= vbCr & "Invalid Value entered for Sales Rep Code"
                'End If

                If ASCMAIN1.CLIENT = "VAN" Then
                    ' Check Size Scale
                    Dim sizes As Integer = 0
                    For isize As Integer = 1 To 24
                        Dim SIZE_CODE As String = Trim(Absx1.txtFor("SIZE_" & Format(isize, "00")).Text).ToUpper
                        Absx1.txtFor("SIZE_" & Format(isize, "00")).Text = SIZE_CODE
                        If SIZE_CODE = "" Then
                            Exit For
                        Else
                            sizes = isize
                            If LookUp("ICTSIZE1", SIZE_CODE) Is Nothing Then
                                EMsg &= vbCr & "Invalid Size Code (" & SIZE_CODE & ")"
                            End If

                            Dim SIZE_QTY As Integer = Val(Absx1.numFor("QTY_" & Format(isize, "00")).Value & "")
                            If SIZE_QTY <= 0 And Absx1.txtFor("SIZE_CODE").Text = "" Then
                                EMsg &= vbCr & "Invalid Qty (" & CStr(SIZE_QTY) & ") for Size Code (" & SIZE_CODE & ")"
                            End If

                        End If
                    Next

                    If sizes < 24 Then
                        For isize As Integer = sizes + 1 To 24
                            If Absx1.txtFor("SIZE_" & Format(isize, "00")).Text <> "" Or Val(Absx1.numFor("QTY_" & Format(isize, "00")).Value & "") <> 0 Then
                                EMsg &= vbCr & "Cannot leave gaps in Size Scale"
                            End If
                        Next
                    End If
                End If

                If ASCMAIN1.CLIENT = "RGI" Then
                    Dim LIST_CALC_CODE As String = Absx1.txtFor("LIST_CALC_CODE").Text
                    If LIST_CALC_CODE <> "" Then
                        Dim rowICTLSTC1 As DataRow = LookUp("ICTLSTC1", LIST_CALC_CODE)
                        If rowICTLSTC1 Is Nothing Then
                            EMsg &= vbCr & "Invalid List Calculation Code (" & LIST_CALC_CODE & ")"
                        Else
                            Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text
                            Dim rowICTLSTCV As DataRow = LookUp("ICTLSTCV", New String() {LIST_CALC_CODE, VEND_CODE})
                            If rowICTLSTCV Is Nothing Then
                                EMsg &= vbCr & "List Calculation Code (" & LIST_CALC_CODE & ") not set up for Vendor " & VEND_CODE
                            End If
                        End If
                    End If
                End If

                If ASCMAIN1.CLIENT = "NYA" Then
                    If dst.Tables("ICTSTYC1").Select("").Length > 1 Then
                        EMsg &= vbCr & "Multiple Colors not Supported at this time"
                    ElseIf dst.Tables("ICTSTYC1").Select("COLOR_CODE <> 'AST'").Length <> 0 Then
                        EMsg &= vbCr & "Only color AST supported at this time"
                    Else
                        Dim rowICTSTYC1() As DataRow = dst.Tables("ICTSTYC1").Select("COLOR_CODE = 'AST'")
                        If rowICTSTYC1.Length = 1 Then
                            If rowICTSTYC1(0).Item("HIDE_COLOR_3PL") & "" <> "1" Then
                                EMsg &= vbCr & "You Must Hide Color from 3PL at this time"
                            End If
                        End If
                    End If

                    For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("ISNULL(UPC_CODE,'') <> ''", "", DataViewRowState.Added + DataViewRowState.ModifiedCurrent)
                        Dim UPC_CODE As String = rowICTSTYC1.Item("UPC_CODE") & ""
                        Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE") & ""
                        Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE") & ""
                        If Len(UPC_CODE) <> 12 Then
                            EMsg &= vbCr & "UPC Code must be 12 characters long (" & UPC_CODE & ")"
                        End If
                        If Not UPC_CODE.StartsWith(SO_PARM_UPC_VENDOR_ID) Then
                            Dim msg As String = "UPC Code " & UPC_CODE & " does not begin with UPC prefix " & SO_PARM_UPC_VENDOR_ID
                            If MsgBox(msg & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "UPC Code Prefix") = MsgBoxResult.No Then
                                EMsg &= vbCr & msg
                            End If
                        End If
                        ASCMAIN1.sql = "Select * from ICTSTYC1 where UPC_CODE = '" & UPC_CODE & "' and (STYLE_CODE <> '" & STYLE_CODE & "' or COLOR_CODE <> '" & COLOR_CODE & "')"
                        Dim ROW As DataRow = ASCDATA1.GetDataRow
                        If ROW IsNot Nothing Then
                            EMsg &= vbCr & "UPC Code " & UPC_CODE & " is already associated with Style/Color " & ROW.Item("STYLE_CODE") & "/" & ROW.Item("COLOR_CODE")
                        End If
                        If EMsg = "" Then

                            Dim UPC_CODE_x = TAC.SOCMAIN1.UPC(Me, Mid(UPC_CODE, 7, 5), Mid(UPC_CODE, 1, 6), True) ' TAC.SOCMAIN1.UPC(Me, Mid(UPC_CODE, 7, 5), SO_PARM_UPC_VENDOR_ID, True)
                            If UPC_CODE <> UPC_CODE_x Then
                                EMsg &= vbCr & "Invalid Check Digit on UPC Code entered: " & UPC_CODE & ".  Should be " & UPC_CODE_x
                            End If
                        End If
                    Next

                    If EntryMode = "New" Then
                        Dim STYLE_CODE_PLM As String = Absx1.txtFor("STYLE_CODE_PLM").Text
                        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
                        If LookUp("ICTPLIN2", STYLE_CODE_PLM) Is Nothing Then
                            EMsg &= vbCr & "Invalid PLM Style Code specified"
                        Else
                            If Not STYLE_CODE.StartsWith(STYLE_CODE_PLM) Then
                                EMsg &= vbCr & "Style Code is not properly configured to PLM Style"
                            End If
                        End If
                    End If

                    If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
                        EMsg &= vbCr & "Division Code is Mandatory"
                    Else
                        If LookUp("SOTSDIV1", Absx1.txtFor("SALES_DIVISION_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Division Code"
                        End If
                    End If

                    If EntryMode = "Edit" And EMsg = "" Then
                        Dim SALES_DIVISION_CODE_OLD As String = rowASFBASE1.Item("SALES_DIVISION_CODE", DataRowVersion.Original)
                        If SALES_DIVISION_CODE_OLD <> "" Then
                            Dim SALES_DIVISION_CODE_NEW As String = Absx1.txtFor("SALES_DIVISION_CODE").Text
                            Dim rowSOTSDIV1_NEW As DataRow = LookUp("SOTSDIV1", SALES_DIVISION_CODE_NEW)

                            Dim rowSOTSDIV1_OLD As DataRow = LookUp("SOTSDIV1", SALES_DIVISION_CODE_OLD)
                            If rowSOTSDIV1_NEW.Item("SEG4_CODE") & "" <> rowSOTSDIV1_OLD.Item("SEG4_CODE") & "" Then
                                EMsg &= vbCr & "You Cannot Change Division Code to a Different Company"
                            End If
                        End If
                    End If

                    If EMsg = "" Then
                        If Absx1.txtFor("CUST_CODE").Text = "" Then
                            If MsgBox("Is there a customer code for this style?" _
                                      & vbCrLf & vbCrLf & "If so, click 'Yes' and make sure it is entered." _
                                      & vbCrLf & vbCrLf & "If not, click 'No'.", _
                                      MsgBoxStyle.YesNo, _
                                      "This style does NOT have a Customer assigned to it") = MsgBoxResult.Yes Then
                            End If
                        End If
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text

        Dim sqlDelete = "STYLE_CODE = '" & STYLE_CODE & "'"
        'INIT_LAST("ICTSTYC1", True)

        If Absx1.optFor("STYLE_STATUS").Value & "" <> "A" Then
            Dim STYLE_STATUS As String = Absx1.optFor("STYLE_STATUS").Value & ""
            For Each ROW As DataRow In dst.Tables("ICTSTYC1").Select("STYLE_COLOR_STATUS <> '" & STYLE_STATUS & "'")
                ROW.Item("STYLE_COLOR_STATUS") = STYLE_STATUS
            Next
        End If

        Update_Record_TDA("ICTSTYC1", sqlDelete)
        Update_Record_TDA("ICTSTYLD", sqlDelete)

        Update_Record_TDA("ICTSTYL3", sqlDelete)
        Update_Record_TDA("ICTSTYL4", sqlDelete)
        Update_Record_TDA("ICTSTYL5", sqlDelete)
        Update_Record_TDA("ICTSTYLC", sqlDelete)

        If ASCMAIN1.CLIENT = "VAN" Then
            Dim rowICTSTYLS As DataRow = dst.Tables("ICTSTYLS").Rows.Find(STYLE_CODE)
            If rowICTSTYLS.Item("SIZE_01") & "" = "" Then
                ASCDATA1.ExecuteSQL("Delete from ICTSTYLS where STYLE_CODE = '" & STYLE_CODE & "'")
            Else
                Update_Record_TDA("ICTSTYLS", sqlDelete)
            End If

        End If


        For Each rowICTSTYV1 As DataRow In dst.Tables("ICTSTYV1").Select("ISNULL(FUTURE_COST,0) <> 0", "")
            If Val(rowICTSTYV1.Item("NEW_PO_COST") & "") <> 0 Then
                rowICTSTYV1.Item("PO_COST") = Val(rowICTSTYV1.Item("NEW_PO_COST") & "")
                rowICTSTYV1.Item("PO_COST_DATE") = rowICTSTYV1.Item("NEW_PO_COST_DATE")
            End If
            rowICTSTYV1.Item("NEW_PO_COST") = Val(rowICTSTYV1.Item("FUTURE_COST") & "")
            rowICTSTYV1.Item("NEW_PO_COST_DATE") = rowICTSTYV1.Item("FUTURE_COST_DATE")
        Next


        Update_Record_TDA("ICTSTYV1", sqlDelete)

        dst.Tables("ICTSTYC2").Rows.Clear()
        dst.Tables("ICTSTYC3").Rows.Clear()
        dst.Tables("ICTSTYC4").Rows.Clear()


        For Each rowICTSTYCS As DataRow In dst.Tables("ICTSTYCS").Select("")
            rowICTSTYCS.Item("POSITION") = DBNull.Value
        Next
        dst.Tables("ICTSTYCS").Rows.Add(New Object() {0, DBNull.Value, 0})

        For SIZE_INDEX As Integer = 1 To MAX_SIZES
            With grdICTSTYCX.DisplayLayout.Bands(0).Groups("SIZE_" & Format(SIZE_INDEX, "00"))
                If Not .Hidden Then
                    Dim rowICTSTYCS As DataRow = dst.Tables("ICTSTYCS").Rows.Find(SIZE_INDEX)
                    rowICTSTYCS.Item("POSITION") = .Header.VisiblePosition
                End If
            End With
        Next

        For Each rowICTSTYCX As DataRow In dst.Tables("ICTSTYCX").Select("")
            'Dim STYLE_CODE As String = rowICTSTYCX.Item("STYLE_CODE")
            If STYLE_CODE <> rowICTSTYCX.Item("STYLE_CODE") Then Throw New Exception("Style Code mismatch")
            Dim COLOR_CODE As String = rowICTSTYCX.Item("COLOR_CODE")
            Dim COLOR_CODE_UPC As String = rowICTSTYCX.Item("COLOR_CODE_UPC")

            Dim SIZE_INDEX_NEW As Integer = -1
            For Each rowICTSTYCS As DataRow In dst.Tables("ICTSTYCS").Select("POSITION IS NOT NULL", "POSITION")
                Dim SIZE_INDEX As Integer = Val(rowICTSTYCS.Item("SIZE_INDEX") & "")
                SIZE_INDEX_NEW += 1

                Dim UPC_CODE As String = rowICTSTYCX.Item("UPC_CODE_" & Format(SIZE_INDEX, "00")) & ""
                Dim PPK_QTY As Int64 = Val(rowICTSTYCX.Item("PPK_QTY_" & Format(SIZE_INDEX, "00")) & "")
                If UPC_CODE <> "" Or SIZE_INDEX = 0 Then
                    Dim TABLE_NAME As String = IIf(SIZE_INDEX = 0, "ICTSTYC2", "ICTSTYC4")
                    Dim row As DataRow = dst.Tables(TABLE_NAME).NewRow
                    row.Item("STYLE_CODE") = STYLE_CODE
                    row.Item("COLOR_CODE") = COLOR_CODE
                    row.Item("COLOR_CODE_UPC") = COLOR_CODE_UPC
                    If SIZE_INDEX <> 0 Then row.Item("SIZE_INDEX") = SIZE_INDEX_NEW
                    row.Item("UPC_CODE") = UPC_CODE
                    If PPK_QTY <> 0 Then row.Item("PPK_QTY") = PPK_QTY
                    dst.Tables(TABLE_NAME).Rows.Add(row)

                    If SIZE_INDEX <> 0 Then
                        Dim rowICTSTYC3 As DataRow = dst.Tables("ICTSTYC3").Rows.Find(New Object() {STYLE_CODE, COLOR_CODE, SIZE_INDEX_NEW})
                        If rowICTSTYC3 Is Nothing Then
                            rowICTSTYC3 = dst.Tables("ICTSTYC3").NewRow
                            rowICTSTYC3.Item("STYLE_CODE") = STYLE_CODE
                            rowICTSTYC3.Item("COLOR_CODE") = COLOR_CODE
                            rowICTSTYC3.Item("SIZE_INDEX") = SIZE_INDEX_NEW
                            rowICTSTYC3.Item("SIZE_CODE") = rowICTSTYCS.Item("SIZE_CODE")
                            dst.Tables("ICTSTYC3").Rows.Add(rowICTSTYC3)
                        End If
                    End If
                End If
            Next
        Next

        Update_Record_TDA("ICTSTYC2", sqlDelete)
        Update_Record_TDA("ICTSTYC3", sqlDelete)
        Update_Record_TDA("ICTSTYC4", sqlDelete)
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        If EntryMode = "New" Then
            'ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        ElseIf EntryMode = "Edit" Then
            'ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")

            ' we should do something about ICTUPCH1 -note that if a UPC is re-assigned, that this table is out of synch with ICTSTYC1, although we do use ICTSTYC1 for UPCs, and we do have the audit trail
        End If
    End Sub

    

    Overrides Sub Show_Record_Special()

        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text

        Dim gp As Int32 = -1
        For Each g As UltraWinGrid.UltraGridGroup In grdICTSTYCX.DisplayLayout.Bands(0).Groups
            gp += 1
            g.Header.VisiblePosition = gp
        Next

        EnforceConstraints(False)
        Fill_Records("ICTSTYC1", New String() {STYLE_CODE})
        Sort_grdColumns(grdICTSTYC1, "COLOR_CODE")
        Fill_Records("ICTSTYL3", New String() {STYLE_CODE})
        Sort_grdColumns(grdICTSTYL3, "ATTR_CODE")
        Fill_Records("ICTSTYL4", New String() {STYLE_CODE})
        Fill_Records("ICTSTYL5", New String() {STYLE_CODE})
        Fill_Records("ICTSTYLC", New String() {STYLE_CODE})

        SHOW_DUTY_EXCEPTIONS()

        Fill_Records("ICTSTYV1", New String() {STYLE_CODE})
        Sort_grdColumns(grdICTSTYV1, "VEND_CODE")
        Fill_Records("ICTSTYC2", New String() {STYLE_CODE})
        Fill_Records("ICTSTYC3", New String() {STYLE_CODE})
        Fill_Records("ICTSTYC4", New String() {STYLE_CODE})
        Fill_Records("ICTSTYCS", New String() {STYLE_CODE})
        Fill_Records("ICTSTYLD", New String() {STYLE_CODE})

        dst.Tables("ICTSTYCX").Rows.Clear()
        For Each TABLE_NAME As String In New String() {"ICTSTYC2", "ICTSTYC4"}
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                Dim SIZE_INDEX As Integer = 0
                If TABLE_NAME = "ICTSTYC4" Then SIZE_INDEX = Val(row.Item("SIZE_INDEX") & "")
                Add_rowICTSTYCX(row, SIZE_INDEX)
            Next
        Next

        Add_Single_Colors()

        With grdICTSTYCX.DisplayLayout.Bands(0)
            For I As Integer = 1 To MAX_SIZES
                .Groups("SIZE_" & Format(I, "00")).Hidden = True
            Next
            For Each rowICTSTYCS As DataRow In dst.Tables("ICTSTYCS").Select("", "SIZE_INDEX")
                Dim I As Integer = rowICTSTYCS.Item("SIZE_INDEX")
                .Groups("SIZE_" & Format(I, "00")).Hidden = False
                .Groups("SIZE_" & Format(I, "00")).Header.Caption = rowICTSTYCS.Item("SIZE_CODE")
            Next
        End With

        EnforceConstraints(True)

        Setup_ICTSTYC1()

        grdICTSTYC1.Text = "Color Codes for Style " & Absx1.txtFor("STYLE_CODE").Text
        ' grdICTSTYL3.Text = "Attributes for Style " & Absx1.txtFor("STYLE_CODE").Text
        grdICTSTYV1.Text = "Vendor Parameters for Style " & Absx1.txtFor("STYLE_CODE").Text

        If (EntryMode = "New") Then
            rowASFBASE1.Item("SUB_UNIT_PACK_QTY") = 1
        End If

        'If ASCMAIN1.CLIENT = "NYA" Then
        '    Absx1.txtFor("SALES_DIVISION_CODE").Enabled = True
        '    If (EntryMode = "Edit") Then
        '        Absx1.txtFor("SALES_DIVISION_CODE").Enabled = False
        '    End If
        'End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            Absx1.txtFor("SALES_DIVISION_CODE").Enabled = True
            If (EntryMode = "Edit") Then
                ASCMAIN1.sql = "Select * from ICTIREC2 where STYLE_CODE = '" & Absx1.txtFor("STYLE_CODE").Text & "'"
                ASCMAIN1.sql = "Select * from SOTINVH2 where STYLE_CODE = '" & Absx1.txtFor("STYLE_CODE").Text & "' and ROWNUM < 2"
                Dim row As DataRow = ASCDATA1.GetDataRow
                If row IsNot Nothing Then
                    Absx1.txtFor("SALES_DIVISION_CODE").Enabled = False
                End If
            End If
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            Dim rowICTSTYLS As DataRow = Fill_Record("ICTSTYLS", STYLE_CODE)
            If rowICTSTYLS Is Nothing Then
                dst.Tables("ICTSTYLS").Rows.Add(New String() {STYLE_CODE})
            End If
        End If

    End Sub

    Sub SHOW_DUTY_EXCEPTIONS()
        If ASCMAIN1.CLIENT = "RGI" Then
            Dim DUTY_RATE_CODE As String = Absx1.txtFor("DUTY_RATE_CODE").Text
            Fill_Records("ICTDUTY4", New String() {DUTY_RATE_CODE})
            grdICTDUTY4.Text = "Duty Rate Modifiers by Country - " & DUTY_RATE_CODE
            Sort_grdColumns(grdICTDUTY4, "COUNTRY_CODE, DUTY_RATE_BEGIN")
        End If
    End Sub

    Sub Add_rowICTSTYCX(row As DataRow, SIZE_INDEX As Integer)
        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
        Dim COLOR_CODE As String = row.Item("COLOR_CODE")
        Dim COLOR_CODE_UPC As String = row.Item("COLOR_CODE_UPC")
        Dim rowICTSTYCX As DataRow = dst.Tables("ICTSTYCX").Rows.Find(New Object() {STYLE_CODE, COLOR_CODE, COLOR_CODE_UPC})
        If rowICTSTYCX Is Nothing Then
            rowICTSTYCX = dst.Tables("ICTSTYCX").NewRow
            rowICTSTYCX.Item("STYLE_CODE") = STYLE_CODE
            rowICTSTYCX.Item("COLOR_CODE") = COLOR_CODE
            rowICTSTYCX.Item("COLOR_CODE_UPC") = COLOR_CODE_UPC
            Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
            rowICTSTYCX.Item("COLOR_DESC_UPC") = rowICTCOLR1.Item("COLOR_DESC")
            dst.Tables("ICTSTYCX").Rows.Add(rowICTSTYCX)
        End If
        rowICTSTYCX.Item("UPC_CODE_" & Format(SIZE_INDEX, "00")) = row.Item("UPC_CODE")
        rowICTSTYCX.Item("PPK_QTY_" & Format(SIZE_INDEX, "00")) = row.Item("PPK_QTY")
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() { _
                "ICTSTYC1", "ICTSTYL3", "ICTSTYL4", "ICTSTYL5", "ICTSTYLC", "ICTDUTY4", _
                "ICTSTYV1", _
                "ICTSTYC2", "ICTSTYC3", "ICTSTYC4", _
                "ICTSTYCX", "ICTSTYCS", "ICTSTYLD"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next

            If ASCMAIN1.CLIENT = "VAN" Then
                dst.Tables("ICTSTYLS").Rows.Clear()
            End If
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTSTYC1.Enabled = tf
        grdICTSTYL3.Enabled = tf
        grdICTSTYL4.Enabled = tf
        grdICTSTYL5.Enabled = tf
        grdICTSTYLC.Enabled = tf
        grdICTDUTY4.Enabled = tf
        grdICTSTYV1.Enabled = tf
        grdICTSTYCX.Enabled = tf
        grdICTSTYLD.Enabled = tf

        btnIMAGE_NAME.Enabled = tf And (EntryMode = "New" Or EntryMode = "Edit")

        grpGenerate.Left = grpClone.Left
        grpGenerate.Top = grpClone.Top
        grpPLM.Left = grpClone.Left - 50
        grpPLM.Top = grpClone.Top

        grpGenerate.Visible = False
        grpClone.Visible = tf And (EntryMode = "View")
        grpPLM.Visible = Not tf And (ASCMAIN1.CLIENT = "NYA")
        Set_Read_Only(grpGenerate, False)
        Set_Read_Only(grpClone, False)
        Set_Read_Only(grpPLM, False)

        If ASCMAIN1.CLIENT = "RGI" Then
            cmdCalculateList.Visible = tf
        End If

        'If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
        '    ' MARILYN SAYS SHE WILL NEVER NEED THIS OPTION
        '    ' IT WILL COST HER A DIET COKE IF SHE EVER WANTS THIS CHECKBOX ENABLED
        '    chkCopyColors.Checked = False
        '    chkCopyColors.Enabled = False
        'End If
        If ASCMAIN1.CLIENT = "VAN" Then
            chkCopyColors.Checked = True
        End If
        If ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA" Then
            chkCopyColors.Checked = True
            chkCopyColors.Enabled = False ' ONLY 1 COLOR, AND i WANT TO MAKE SURE WE COPY IT
        End If


        'If ASCMAIN1.CLIENT = "NYA" Then
        '    Set_Read_Only_for_ctl(Absx1.txtFor("SALES_DIVISION_CODE"), Not (EntryMode = "New"))
        '    'Absx1.txtFor("SALES_DIVISION_CODE").Enabled = True
        '    'If (EntryMode = "Edit") Then
        '    '    Absx1.txtFor("SALES_DIVISION_CODE").Enabled = False
        '    'End If
        'End If

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTSTYC1, grdICTSTYL3, grdICTSTYL4, grdICTSTYL5,
                                                                                grdICTSTYLC, grdICTSTYV1, grdICTSTYCX, grdICTSTYLD}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If

            End With
        Next

        With grdICTDUTY4.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False
        End With

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            grdICTSTYC1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            ' ONLY AST SUPPORTED AT THIS TIME
            If EntryMode = "New" Then
                Dim row As DataRow = dst.Tables("ICTSTYC1").Rows.Add(New String() {Absx1.txtFor("STYLE_CODE").Text, "AST"})
                row.Item("HIDE_COLOR_3PL") = "1"
                row.Item("COLOR_DESC") = "ASSORTED"
                row.Item("STYLE_COLOR_STATUS") = "A"
            End If
        End If

        'grdICTSTYC1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdICTSTYCX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdICTSTYCX.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If tf Then
            Else
                picStyleColor.Visible = False
                picStyleColor2.Visible = False
                UltraExplorerBar1.Groups.Item("Image").Visible = False
            End If
        End If
    End Sub

#End Region

#Region "grdICTSTYC1"

    Private Sub grdICTSTYC1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYC1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "COLOR_CODE"
                grdCodeDesc(grdICTSTYC1, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
            Case "THEME_CODE"
                grdCodeDesc(grdICTSTYC1, "ICTTHEME", "THEME_CODE", "THEME_DESC")
        End Select

    End Sub

    Private Sub grdICTSTYC1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTYC1.AfterRowActivate

        With grdICTSTYC1.DisplayLayout.Bands(0).Columns("COLOR_CODE")
            If grdICTSTYC1.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        Setup_ICTSTYC1()
        'grdICTSTYC1.DisplayLayout.Bands(0).Columns("UPC_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
    End Sub

    Private Sub grdICTSTYC1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdICTSTYC1.AfterRowsDeleted
        Setup_ICTSTYC1()
    End Sub

    Private Sub grdICTSTYC1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTSTYC1.AfterRowUpdate
        Add_Single_Colors()
        ' grdICTSTYC1.DisplayLayout.Bands(0).Columns("UPC_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
    End Sub

    Private Sub grdICTSTYC1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdICTSTYC1.BeforeRowsDeleted
        Dim EMsgs As String = ""
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value
            Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If Not rowICTSTYC1.RowState = DataRowState.Added Then

                If ASCMAIN1.CLIENT = "VAN" Then
                    ASCMAIN1.sql = "Select Count (*) from SOTORDR2 where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
                    If ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE}) <> 0 Then
                        EMsgs &= vbCrLf & "Cannot Delete a Color used in a Sales Order (" & COLOR_CODE & ")"
                    End If
                    ASCMAIN1.sql = "Select Count (*) from POTORDR2 where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
                    If ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE}) <> 0 Then
                        EMsgs &= vbCrLf & "Cannot Delete a Color used in a Purchase Order (" & COLOR_CODE & ")"
                    End If
                    ASCMAIN1.sql = "Select Count (*) from ICTQUOT3 where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
                    If ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE}) <> 0 Then
                        EMsgs &= vbCrLf & "Cannot Delete a Color used in a Quote (" & COLOR_CODE & ")"
                    End If
                    ASCMAIN1.sql = "Select Count (*) from ICTSTAT1 where STYLE_CODE = :PARM1 and COLOR_CODE = :PARM2"
                    If ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {STYLE_CODE, COLOR_CODE}) <> 0 Then
                        EMsgs &= vbCrLf & "Cannot Delete a Color which ever had Status Qtys (" & COLOR_CODE & ")"
                    End If
                Else
                    MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    e.Cancel = True
                End If
            End If
        Next

        If ASCMAIN1.CLIENT = "VAN" Then
            If EMsgs <> "" Then
                MsgBox(Mid(EMsgs, 3), MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub grdICTSTYC1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYC1.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTCOLR1", e.Row.Cells("COLOR_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("STYLE_CODE").Value = Absx1.txtFor("STYLE_CODE").Text
            If e.Row.Cells("STYLE_COLOR_STATUS").Value & "" = "" Then
                e.Row.Cells("STYLE_COLOR_STATUS").Value = "A"
            End If
        End If

    End Sub

    Private Sub grdICTSTYC1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYC1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "COLOR_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTCOLR1.COLOR_CODE not in", "ICTSTYC1", "COLOR_CODE")
                grdClickCellButton(grdICTSTYC1, sql_where, True)

            Case "UPC_CODE"
                If EntryMode = "New" Or EntryMode = "Edit" And e.Cell.Value & "" = "" Then
                    If Not grdICTSTYC1.ActiveRow.IsAddRow Then
                        Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
                        Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
                        Dim UPC_CODE As String = Get_UPC_Code(STYLE_CODE, COLOR_CODE)
                        e.Cell.Value = UPC_CODE
                    End If
                End If

            Case "STYLE_BIN"
                Dim sql_where As String = "WHSE_CODE = (SELECT SO_PARM_DEF_PICK_WHSE FROM SOTPARM1 WHERE SO_PARM_KEY = 'Z')"
                grdClickCellButton(grdICTSTYC1, sql_where, True)

            Case "CATALOG_SELECTION_CODE", "NRF_SIZE_CODE", "NRF_COLOR_CODE"
                Dim sql_where As String = ""
                grdClickCellButton(grdICTSTYC1, sql_where, True)

            Case "THEME_CODE"
                'Dim sql_where As String = Get_List_of_Codes("ICTCOLR1.COLOR_CODE not in", "ICTSTYC1", "COLOR_CODE")
                grdClickCellButton(grdICTSTYC1)
        End Select
    End Sub

    Private Sub grdICTSTYC1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYC1.InitializeRow
        If e.Row.Cells("STYLE_CODE").Text <> "" And e.Row.Cells("STYLE_CODE").Text <> Absx1.txtFor("STYLE_CODE").Text Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ICTSTYC1"), e.Row)
    End Sub

#End Region

#Region "grdICTSTYL3"

    Private Sub grdICTSTYL3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYL3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ATTR_CODE"
                grdCodeDesc(grdICTSTYL3, "ICTATTR1", "ATTR_CODE", "ATTR_DESC")
        End Select
    End Sub

    Private Sub grdICTSTYL3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYL3.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTATTR1", e.Row.Cells("ATTR_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("STYLE_CODE").Value = Absx1.txtFor("STYLE_CODE").Text
        End If

    End Sub

    Private Sub grdICTSTYL3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYL3.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ATTR_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTATTR1.ATTR_CODE not in", "ICTSTYL3", "ATTR_CODE")
                grdClickCellButton(grdICTSTYL3, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTSTYL3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYL3.InitializeRow
        If e.Row.Cells("STYLE_CODE").Text <> "" And e.Row.Cells("STYLE_CODE").Text <> Absx1.txtFor("STYLE_CODE").Text Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ICTSTYL3"), e.Row)
    End Sub

#End Region

#Region "grdICTSTYL4"

    Private Sub grdICTSTYL4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYL4.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_COMP"
                grdCodeDesc(grdICTSTYL4, "ICTSTYL1", "STYLE_CODE_COMP", "STYLE_DESC")
            Case "COLOR_CODE_COMP"
                grdCodeDesc(grdICTSTYL4, "ICTCOLR1", "COLOR_CODE_COMP", "COLOR_DESC")
        End Select
    End Sub

    Private Sub grdICTSTYL4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYL4.BeforeRowUpdate

        If e.Row.Cells("COMP_NON_STOCK").Value & "" = "" Then
            e.Row.Cells("COMP_NON_STOCK").Value = "1"
        End If

        If e.Row.Cells("COMP_NON_STOCK").Value = "1" Then
            e.Row.Cells("STYLE_CODE_COMP").Value = DBNull.Value
            e.Row.Cells("COLOR_CODE_COMP").Value = DBNull.Value
        Else
            Dim row As DataRow = LookUp("ICTSTYC1", New String() {e.Row.Cells("STYLE_CODE_COMP").Text, e.Row.Cells("COLOR_CODE_COMP").Text})

            If row Is Nothing Then
                MsgBox("Invalid Value Specified for Style and Color", MsgBoxStyle.OkOnly, "Cannot Add Prepack Component")
                e.Cancel = True
                Exit Sub
            End If
        End If

        Dim QTY_PER_PPK As Integer = Val(e.Row.Cells("QTY_PER_PPK").Value & "")
        If QTY_PER_PPK <= 0 Then
            MsgBox("Qty Per Prepack must be greater than 0", MsgBoxStyle.OkOnly, "Cannot Add Prepack Component")
            e.Cancel = True
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
            e.Row.Cells("STYLE_CODE").Value = STYLE_CODE
            e.Row.Cells("COLOR_CODE").Value = COLOR_CODE
            Dim sqlw As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            Dim COMP_LNO As Integer = Val(dst.Tables("ICTSTYL4").Compute("MAX(COMP_LNO)", sqlw) & "")
            e.Row.Cells("COMP_LNO").Value = COMP_LNO + 1
        End If

    End Sub

    Private Sub grdICTSTYL4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYL4.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_COMP"
                grdClickCellButton(grdICTSTYL4, "", True)
            Case "COLOR_CODE_COMP"
                Dim STYLE_CODE_COMP As String = e.Cell.Row.Cells("STYLE_CODE_COMP").Value & ""
                Dim sql_where As String = "COLOR_CODE in (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE_COMP & "')"
                grdClickCellButton(grdICTSTYL4, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTSTYL4_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYL4.InitializeRow
        If e.Row.Cells("STYLE_CODE").Text <> "" And e.Row.Cells("STYLE_CODE").Text <> Absx1.txtFor("STYLE_CODE").Text Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ICTSTYL4"), e.Row)
    End Sub

#End Region

#Region "grdICTSTYL5"

    Private Sub grdICTSTYL5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYL5.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "PANTONE_CODE"
                grdCodeDesc(grdICTSTYL5, "ICTPANTC", "PANTONE_CODE", "PANTONE_DESC")
                grdCodeDesc(grdICTSTYL5, "ICTPANTC", "PANTONE_CODE", "RGB")
        End Select
    End Sub

    Private Sub grdICTSTYL5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYL5.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTPANTC", e.Row.Cells("PANTONE_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
            e.Row.Cells("STYLE_CODE").Value = STYLE_CODE
            e.Row.Cells("COLOR_CODE").Value = COLOR_CODE
            e.Row.Cells("COLOR_SEQ_NO").Value = Val(dst.Tables("ICTSTYL5").Compute("MAX(COLOR_SEQ_NO)", "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'") & "") + 1
        End If

    End Sub

    Private Sub grdICTSTYL5_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYL5.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "PANTONE_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTPANTC.PANTONE_CODE not in", "ICTSTYL5", "PANTONE_CODE")
                grdClickCellButton(grdICTSTYL5, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTSTYL5_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYL5.InitializeRow
        If e.Row.Cells("STYLE_CODE").Text <> "" And e.Row.Cells("STYLE_CODE").Text <> Absx1.txtFor("STYLE_CODE").Text Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ICTSTYL5"), e.Row)

        Dim RGB As String = e.Row.Cells("RGB").Text
        If RGB = "" Then
            e.Row.Cells("SWATCH").Appearance.BackColor = System.Drawing.Color.Empty
        Else
            Dim ARGB() As String = Split(RGB & ",,,", ",")
            Dim R As Integer = Val(ARGB(0)) : If R < 0 Or R > 255 Then R = 0
            Dim G As Integer = Val(ARGB(1)) : If G < 0 Or G > 255 Then G = 0
            Dim B As Integer = Val(ARGB(2)) : If B < 0 Or B > 255 Then B = 0
            e.Row.Cells("SWATCH").Appearance.BackColor = System.Drawing.Color.FromArgb(255, R, G, B)
        End If

    End Sub

#End Region

#Region "grdICTSTYLC"

    Private Sub grdICTSTYLC_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYLC.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "DUTY_RATE_CODE"
                grdCodeDesc(grdICTSTYLC, "ICTDUTY1", "DUTY_RATE_CODE", "DUTY_RATE_DESC")

        End Select
    End Sub

    Private Sub grdICTSTYLC_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYLC.BeforeRowUpdate

        Dim row As DataRow = LookUp("TATCNTRY", e.Row.Cells("COUNTRY_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        Dim rowICTDUTY1 As DataRow = LookUp("ICTDUTY1", e.Row.Cells("DUTY_RATE_CODE").Text)

        If rowICTDUTY1 Is Nothing Then
            e.Cancel = True
        End If

        If Not e.Cancel Then
            Dim DUTY_RATE_CODE As String = e.Row.Cells("DUTY_RATE_CODE").Text
            Dim COUNTRY_CODE As String = e.Row.Cells("COUNTRY_CODE").Text

            If DUTY_RATE_CODE.Length <> 16 Then
                e.Cancel = True
            ElseIf Mid(DUTY_RATE_CODE, 13, 1) <> "-" Then
                e.Cancel = True
            End If
            If Mid(DUTY_RATE_CODE, 14, 3) <> COUNTRY_CODE Then
                e.Cancel = True
            End If
        End If


        If e.Row.IsAddRow Then
            Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
            e.Row.Cells("STYLE_CODE").Value = STYLE_CODE
        End If

    End Sub

    Private Sub grdICTSTYLC_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYLC.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "COUNTRY_CODE"
                Dim sql_where As String = Get_List_of_Codes("TATCNTRY.COUNTRY_CODE not in", "ICTSTYLC", "COUNTRY_CODE")
                grdClickCellButton(grdICTSTYLC, sql_where, True)
            Case "DUTY_RATE_CODE"
                Dim COUNTRY_CODE As String = e.Cell.Row.Cells("COUNTRY_CODE").Value & ""
                Dim sql_where As String = "ICTDUTY1.DUTY_RATE_CODE LIKE '%-" & "" & "%'"
                grdClickCellButton(grdICTSTYLC, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTSTYLC_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYLC.InitializeRow

    End Sub

#End Region

#Region "grdICTSTYV1"

    Private Sub grdICTSTYV1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYV1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "VEND_CODE"
                grdCodeDesc(grdICTSTYV1, "APTVEND1", "VEND_CODE", "VEND_NAME")
        End Select
    End Sub

    Private Sub grdICTSTYV1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYV1.BeforeRowUpdate

        Dim row As DataRow = LookUp("APTVEND1", e.Row.Cells("VEND_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("STYLE_CODE").Value = Absx1.txtFor("STYLE_CODE").Text
        End If

    End Sub

    Private Sub grdICTSTYV1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYV1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "VEND_CODE"
                Dim sql_where As String = Get_List_of_Codes("APTVEND1.VEND_CODE not in", "ICTSTYV1", "VEND_CODE")
                grdClickCellButton(grdICTSTYV1, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTSTYV1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYV1.InitializeRow
        If e.Row.Cells("STYLE_CODE").Text <> "" And e.Row.Cells("STYLE_CODE").Text <> Absx1.txtFor("STYLE_CODE").Text Then
            e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ICTSTYV1"), e.Row)
        With e.Row.Cells("VEND_REMARK")
            .ToolTipText = .Value & ""

        End With
    End Sub

#End Region

    Sub Setup_ICTSTYC1()
        grdICTSTYC1.DisplayLayout.Bands(0).Columns("UPC_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        If grdICTSTYC1.ActiveRow Is Nothing OrElse grdICTSTYC1.ActiveRow.IsAddRow OrElse Not grdICTSTYC1.ActiveRow.IsDataRow Then
            grdICTSTYCX.Visible = False
            grdICTSTYL4.Visible = False
            grdICTSTYL5.Visible = False

        Else
            Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
            Dim dvw As DataView = Nothing

            dvw = DirectCast(grdICTSTYCX.DataSource, DataTable).DefaultView
            dvw.RowFilter = "COLOR_CODE = '" & COLOR_CODE & "'"
            Sort_grdColumns(grdICTSTYCX, "COLOR_CODE_UPC")
            grdICTSTYCX.Text = "Size, Color, UPC Codes for Color " & COLOR_CODE
            grdICTSTYCX.Visible = True

            dvw = DirectCast(grdICTSTYL4.DataSource, DataTable).DefaultView
            dvw.RowFilter = "COLOR_CODE = '" & COLOR_CODE & "'"
            Sort_grdColumns(grdICTSTYL4, "STYLE_CODE_COMP, COLOR_CODE_COMP")
            grdICTSTYL4.Text = "Component Styles for Color " & COLOR_CODE
            grdICTSTYL4.Visible = True

            dvw = DirectCast(grdICTSTYL5.DataSource, DataTable).DefaultView
            dvw.RowFilter = "COLOR_CODE = '" & COLOR_CODE & "'"
            Sort_grdColumns(grdICTSTYL5, "PANTONE_CODE")
            grdICTSTYL5.Text = "Pantone Colors for Color " & COLOR_CODE
            grdICTSTYL5.Visible = True

        End If

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            If grdICTSTYC1.ActiveRow Is Nothing OrElse grdICTSTYC1.ActiveRow.IsAddRow OrElse Not grdICTSTYC1.ActiveRow.IsDataRow Then
                picStyleColor.Image = Nothing
                picStyleColor.Visible = False
                picStyleColor2.Image = Nothing
                picStyleColor2.Visible = False
                UltraExplorerBar1.Groups.Item("Image").Visible = False
            Else
                Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
                Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
                Dim IMAGE_NAME As String = STYLE_CODE & "-" & COLOR_CODE

                Dim imgba() As Byte = Nothing
                If IMAGE_NAME <> "" Then
                    Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
                    picStyleColor.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
                    picStyleColor.Visible = True
                    picStyleColor2.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
                    picStyleColor2.Visible = True
                    UltraExplorerBar1.Groups.Item("Image").Visible = True
                Else
                    picStyleColor.Image = Nothing
                    picStyleColor.Visible = False
                    picStyleColor2.Image = Nothing
                    picStyleColor2.Visible = False
                    UltraExplorerBar1.Groups.Item("Image").Visible = False
                End If
            End If
        End If
    End Sub

    Sub Add_Colors()

        Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value

        Dim sqlw As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
        Dim sql_where As String = Get_List_of_Codes("ICTCOLR1.COLOR_CODE not in", "ICTSTYCX", "COLOR_CODE_UPC", sqlw)
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("COLOR_CODE", , sql_where)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                For Each CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    Dim rowICTSTYCX As DataRow = dst.Tables("ICTSTYCX").NewRow
                    rowICTSTYCX.Item("STYLE_CODE") = STYLE_CODE
                    rowICTSTYCX.Item("COLOR_CODE") = COLOR_CODE
                    rowICTSTYCX.Item("COLOR_CODE_UPC") = CODE
                    dst.Tables("ICTSTYCX").Rows.Add(rowICTSTYCX)
                Next

                Sort_grdColumns(grdICTSTYCX, "COLOR_CODE_UPC")
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If
    End Sub

    Sub Add_Size()

        If dst.Tables("ICTSTYCS").Rows.Count >= MAX_SIZES Then
            MsgBox("Maximum Number of Sizes " & CStr(MAX_SIZES) & " has been Reached", _
                   MsgBoxStyle.OkOnly, "Cannot Add Any More Sizes")
            Exit Sub
        End If

        Dim sql_where As String = Get_List_of_Codes("ICTSIZE1.SIZE_CODE not in", "ICTSTYCS", "SIZE_CODE")
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("SIZE_CODE", , sql_where)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading")

                For Each CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    For SIZE_INDEX As Integer = 1 To MAX_SIZES
                        Dim rowICTSTYCS As DataRow = dst.Tables("ICTSTYCS").Rows.Find(SIZE_INDEX)
                        If rowICTSTYCS Is Nothing Then
                            rowICTSTYCS = dst.Tables("ICTSTYCS").NewRow
                            rowICTSTYCS.Item("SIZE_INDEX") = SIZE_INDEX
                            rowICTSTYCS.Item("SIZE_CODE") = CODE
                            dst.Tables("ICTSTYCS").Rows.Add(rowICTSTYCS)
                            With grdICTSTYCX.DisplayLayout.Bands(0).Groups("SIZE_" & Format(SIZE_INDEX, "00"))
                                .Hidden = False
                                .Header.Caption = CODE
                            End With
                            Exit For
                        End If
                    Next
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If
    End Sub

    Private Sub cmdGenerate_Click(sender As System.Object, e As System.EventArgs) Handles cmdGenerate.Click
        Dim STYLE_CLASS_CODE As String = cbeSTYLE_CLASS_CODE.Value & ""
        If STYLE_CLASS_CODE = "" Then
            MsgBox("You must first select a Class Code before Auto-Generating the next Style Code", _
                   MsgBoxStyle.OkOnly, "Cannot Auto-Generate Style")
            Exit Sub
        End If

        If Not ASCMAIN1.Logical_Lock("ICTCLAS1", STYLE_CLASS_CODE, False, True, True, 2) Then Exit Sub

        Dim EMsg As String = ""

        Dim rowICTCLAS1 As DataRow = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
        If rowICTCLAS1 Is Nothing Then
            EMsg = "Could not find Class Record for " & STYLE_CLASS_CODE
        Else
            Dim STYLE_CLASS_STYLE_MASK As String = rowICTCLAS1.Item("STYLE_CLASS_STYLE_MASK") & ""
            If STYLE_CLASS_STYLE_MASK = "" Then
                STYLE_CLASS_STYLE_MASK = "##########"
            ElseIf Not STYLE_CLASS_STYLE_MASK.Contains("#") Then
                STYLE_CLASS_STYLE_MASK &= "##########"
            End If
            If STYLE_CLASS_STYLE_MASK.Length > Absx1.txtFor("STYLE_CODE").MaxLength Then
                EMsg = "Total Mask Size Exceeds Maximum Length of Style Code"
            Else
                Dim STYLE_SEQ_START As Integer = InStr(STYLE_CLASS_STYLE_MASK, "#")
                Dim S As String = STYLE_CLASS_STYLE_MASK.Substring(STYLE_SEQ_START - 1)
                Dim STYLE_SEQ_LENGTH As Integer = 1
                Do While S.Length > STYLE_SEQ_LENGTH And Mid(S, STYLE_SEQ_LENGTH, 1) = "#"
                    STYLE_SEQ_LENGTH += 1
                Loop
                ASCDATA1.ExecuteSQL("Update ICTCLAS1 Set STYLE_CLASS_STYLE_SEQ = NVL(STYLE_CLASS_STYLE_SEQ,0) + 1 where STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'")
                Dim STYLE_CLASS_STYLE_SEQ As Int64 = Val(rowICTCLAS1.Item("STYLE_CLASS_STYLE_SEQ") & "")
                rowICTCLAS1 = LookUp("ICTCLAS1", STYLE_CLASS_CODE)
                If STYLE_CLASS_STYLE_SEQ + 1 <> Val(rowICTCLAS1.Item("STYLE_CLASS_STYLE_SEQ") & "") Then
                    EMsg = "Problem with Sequence Control"
                Else
                    STYLE_CLASS_STYLE_SEQ = Val(rowICTCLAS1.Item("STYLE_CLASS_STYLE_SEQ") & "")
                    Dim SF As String = "".PadLeft(STYLE_SEQ_LENGTH, "#")
                    Dim SS As String = Format(STYLE_CLASS_STYLE_SEQ, Replace(SF, "#", "0"))
                    Dim STYLE_CODE As String = Replace(STYLE_CLASS_STYLE_MASK, SF, SS, , 1)
                    Dim row As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    If row IsNot Nothing Then
                        EMsg = "Auto-Generated Next Style Code (" & STYLE_CODE & ") already exists"
                    Else
                        If Not ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE, False, True, True, 2) Then Exit Sub
                        Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
                        Click_Command("New")
                        Absx1.txtFor("STYLE_CLASS_CODE").Text = STYLE_CLASS_CODE
                        Absx1.optFor("STYLE_STATUS").Value = "A"
                    End If
                End If
            End If
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Auto-Generate Style")
        End If

        ASCMAIN1.MultiTask_Release(, , 2)

    End Sub

    Private Sub cmdClone_Click(sender As System.Object, e As System.EventArgs) Handles cmdClone.Click
        Dim EMsg As String = ""

        Dim STYLE_CODE As String = Trim(txtCLONE_STYLE_CODE.Text)
        Dim STYLE_CODE_ORIG As String = Absx1.txtFor("STYLE_CODE").Text
        ' MAKE SURE WE HAVE NO ILLEGAL CHARACTERS IN NEW STYLE CODE
        STYLE_CODE = ASCMAIN1.Format_Field(STYLE_CODE, "STYLE_CODE")
        txtCLONE_STYLE_CODE.Text = STYLE_CODE
        If STYLE_CODE = "" Then Exit Sub

        If Not ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE, False, True, True, 2) Then Exit Sub

        Dim row As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        If row IsNot Nothing Then
            EMsg = "Style Code " & STYLE_CODE & " already exists"
        End If

        If txtCLONE_CUST_CODE.Text <> "" Then
            row = LookUp("ARTCUST1", txtCLONE_CUST_CODE.Text)
            If row Is Nothing Then
                EMsg = "Invalid Customer Code Specified (" & txtCLONE_CUST_CODE.Text & ")"
            End If
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Clone Style")
        Else


            If MsgBox("Are you sure you want to clone Style " & STYLE_CODE_ORIG & vbCrLf & " to Style " & STYLE_CODE, _
                      MsgBoxStyle.YesNo, "The Following Action is Permanent") = MsgBoxResult.No Then
                Exit Sub
            End If

            Dim rowICTSTYL1 As DataRow = Fill_Record("ICTSTYL1_NEW", Absx1.txtFor("STYLE_CODE").Text)
            BeginTrans()
            For Each TABLE_NAME As String In New String() {"ICTSTYL1_NEW", "ICTSTYC1", "ICTSTYV1", "ICTSTYL3", "ICTSTYL4", "ICTSTYL5", "ICTSTYLC", "ICTSTYLS"}
                If TABLE_NAME = "ICTSTYC1" And Not chkCopyColors.Checked Then
                    ' DO NOT COPY COLORS
                ElseIf TABLE_NAME = "ICTSTYLS" And ASCMAIN1.CLIENT <> "VAN" Then
                    ' ONLY FOR VAN
                Else
                    For Each rowNew As DataRow In dst.Tables(TABLE_NAME).Select()
                        rowNew.Item("STYLE_CODE") = STYLE_CODE
                        If TABLE_NAME = "ICTSTYL1_NEW" Then
                            If txtCLONE_CUST_CODE.Text <> "" Then
                                rowNew.Item("CUST_CODE") = txtCLONE_CUST_CODE.Text
                                rowNew.Item("STYLE_CODE_ORIG") = STYLE_CODE_ORIG
                            End If
                            rowNew.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            rowNew.Item("INIT_DATE") = DATETIME_STAMP
                            rowNew.Item("LAST_OPER") = DBNull.Value
                            rowNew.Item("LAST_DATE") = DBNull.Value

                            If ASCMAIN1.CLIENT = "VAN" Then
                                rowNew.Item("VEND_CODE") = DBNull.Value
                                rowNew.Item("FACTORY_CODE") = DBNull.Value
                                rowNew.Item("COUNTRY_CODE") = DBNull.Value
                            End If
                        End If
                        If TABLE_NAME = "ICTSTYC1" Then
                            rowNew.Item("UPC_CODE") = DBNull.Value
                        End If
                        rowNew.AcceptChanges()
                        rowNew.SetAdded()
                    Next
                    If TABLE_NAME = "ICTSTYC1" Then
                        If ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "VAN" Then
                        Else
                            Generate_UPCs()
                        End If
                    End If
                    Update_Record_TDA(TABLE_NAME)
                End If
            Next

            CommitTrans()
        End If

        ASCMAIN1.MultiTask_Release(, , 2)

        If EMsg = "" Then
            Click_Command("Done")
            txtCLONE_CUST_CODE.Text = ""
            txtCLONE_STYLE_CODE.Text = ""
            Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
            Click_Command("Edit")
        End If

    End Sub

    Private Sub txtCLONE_CUST_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtCLONE_CUST_CODE.ValueChanged
        Dim row As DataRow = LookUp("ARTCUST1", txtCLONE_CUST_CODE.Text)
        If row IsNot Nothing Then
            Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text & row.Item("CUST_STYLE_SFX")
            txtCLONE_STYLE_CODE.Text = STYLE_CODE
        End If
    End Sub

    Function Get_UPC_Code(STYLE_CODE As String, COLOR_CODE As String) As String

        Dim UPC_CODE As String = ""
        Do
            Dim UPC_CODE_CTL_NO As String = ""
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("UPC_CODE")
            Else
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("ICTUPCH1.UPC_CODE")
            End If

            UPC_CODE = TAC.SOCMAIN1.UPC(Me, UPC_CODE_CTL_NO, SO_PARM_UPC_VENDOR_ID, True)
            If LookUp("ICTUPCH1", UPC_CODE) Is Nothing Then Exit Do
        Loop

        ASCMAIN1.sql = "Insert into ICTUPCH1 (UPC_CODE,STYLE_CODE,COLOR_CODE,INIT_DATE,INIT_OPER) " & vbCrLf _
            & " values (:PARM1,:PARM2,:PARM3,SYSDATE,:PARM4)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {UPC_CODE, STYLE_CODE, COLOR_CODE, ASCMAIN1.USER_ID})

        Return UPC_CODE
    End Function

    Sub Generate_UPCs()
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("UPC_CODE IS NULL")
            rowICTSTYC1.Item("UPC_CODE") = Get_UPC_Code(rowICTSTYC1.Item("STYLE_CODE"), rowICTSTYC1.Item("COLOR_CODE"))
        Next
    End Sub

    Private Sub grdICTSTYCX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTYCX.AfterRowActivate
        With grdICTSTYCX.DisplayLayout.Bands(0).Columns("COLOR_CODE_UPC")
            If grdICTSTYCX.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdICTSTYCX_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYCX.ClickCellButton
        Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value

        If e.Cell.Column.Key = "COLOR_CODE_UPC" Then
            Dim sqlw As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
            Dim sql_where As String = Get_List_of_Codes("ICTCOLR1.COLOR_CODE not in", "ICTSTYCX", "COLOR_CODE_UPC", sqlw)
            grdClickCellButton(grdICTSTYC1, sql_where, True)

        ElseIf e.Cell.Column.Key.StartsWith("UPC_CODE_") Then
            If EntryMode = "New" Or EntryMode = "Edit" And e.Cell.Value & "" = "" Then
                e.Cell.Value = Get_UPC_Code(STYLE_CODE, COLOR_CODE)
            End If
        End If
    End Sub

    Sub Add_Single_Colors()
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("COLOR_CODE <> 'AST'", "")
            Dim STYLE_CODE As String = rowICTSTYC1.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTSTYC1.Item("COLOR_CODE")
            Dim rowICTSTYCX As DataRow = dst.Tables("ICTSTYCX").Rows.Find(New String() {STYLE_CODE, COLOR_CODE, COLOR_CODE})
            If rowICTSTYCX Is Nothing Then
                rowICTSTYCX = dst.Tables("ICTSTYCX").NewRow
                rowICTSTYCX.Item("STYLE_CODE") = STYLE_CODE
                rowICTSTYCX.Item("COLOR_CODE") = COLOR_CODE
                rowICTSTYCX.Item("COLOR_CODE_UPC") = COLOR_CODE
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                rowICTSTYCX.Item("COLOR_DESC_UPC") = rowICTCOLR1.Item("COLOR_DESC")
                dst.Tables("ICTSTYCX").Rows.Add(rowICTSTYCX)
            End If
        Next
    End Sub

    Private Sub btnIMAGE_NAME_Click(sender As System.Object, e As System.EventArgs) Handles btnIMAGE_NAME.Click
        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Image File to Associate with Style " & Absx1.txtFor("STYLE_CODE").Text
            openFileDialog1.Filter = "jpg files (*.jpg)|*.jpg|png files (*.png)|*.png"
            openFileDialog1.RestoreDirectory = True
            openFileDialog1.InitialDirectory = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            If Not FILENAME.StartsWith(FOLDER_NAME) Then
                MsgBox("Images must be located in " & FOLDER_NAME, MsgBoxStyle.OkOnly, "Invalid Path/Folder Name")
            Else
                Dim IMAGE_NAME As String = Mid(FILENAME, Len(FOLDER_NAME) + 1)
                If IMAGE_NAME.StartsWith("\") Then IMAGE_NAME = Mid(IMAGE_NAME, 2)
                Absx1.txtFor("IMAGE_NAME").Text = IMAGE_NAME
            End If
        End If
    End Sub

    Private Sub txtIMAGE_NAME_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtIMAGE_NAME.ValueChanged
        If txtIMAGE_NAME.Text = "" Then
            imgSTYLE.Image = Nothing
        Else
            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            Dim IMAGE_NAME As String = txtIMAGE_NAME.Text
            imgSTYLE.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , ) ' imgba)
        End If
    End Sub

    Private Sub numCASE_CUBE_DoubleClick(sender As Object, e As System.EventArgs) Handles numCASE_CUBE.DoubleClick

        If EntryMode <> "Edit" Then
            Exit Sub
        End If

        Dim frm As New TAC.TAFCUBE1
        Select Case frm.ShowDialog
            Case Windows.Forms.DialogResult.OK
                numCASE_CUBE.Value = frm.calculatedCube
            Case Windows.Forms.DialogResult.Cancel
        End Select
        frm.Close()
        frm = Nothing
    End Sub

    Private Sub cmdCreatePLM_Click(sender As System.Object, e As System.EventArgs) Handles cmdCreatePLM.Click
        Create_Style_from_PLM()
    End Sub

    Sub Create_Style_from_PLM(Optional show_messages = True)

        txtSTYLE_CODE_PLM_SOURCE.Text = txtSTYLE_CODE_PLM_SOURCE.Text.ToUpper
        txtSTYLE_CODE_NEW.Text = txtSTYLE_CODE_NEW.Text.ToUpper

        Dim STYLE_CODE_PLM As String = txtSTYLE_CODE_PLM_SOURCE.Text

        If STYLE_CODE_PLM = "" Then
            MsgBox("You must first select a PLM Style Code to Create a Style from the PLM Definition", _
                   MsgBoxStyle.OkOnly, "Cannot Create Style")
            Exit Sub
        End If

        If Not ASCMAIN1.Logical_Lock("ICTPLIN2", STYLE_CODE_PLM, False, True, True, 2) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("ICTSTYL1", STYLE_CODE_PLM, False, True, True, 2) Then Exit Sub

        Dim EMsg As String = ""

        Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
        If rowICTPLIN2 Is Nothing Then
            EMsg = "Could not find PLM Style Record for " & STYLE_CODE_PLM
        Else
            'Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_PLM)
            'If rowICTSTYL1 IsNot Nothing Then
            '    EMsg = "Record already exists for Style " & STYLE_CODE_PLM
            'End If
        End If

        Dim STYLE_CODE_NEW As String = txtSTYLE_CODE_NEW.Text
        If STYLE_CODE_NEW.StartsWith(STYLE_CODE_PLM) Then
            Dim row As DataRow = LookUp("ICTSTYL1", STYLE_CODE_NEW)
            If row IsNot Nothing Then
                EMsg = "Style Code " & STYLE_CODE_NEW & " already exists"
            End If
        Else
            EMsg = "Style Code to assign must begin with the PLM Style Code " & STYLE_CODE_PLM
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Create Style")
        Else
            Dim STYLE_CODE As String = STYLE_CODE_NEW ' STYLE_CODE_PLM
            Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE

            txtSTYLE_CODE_PLM_SOURCE.Tag = STYLE_CODE
            Click_Command("New")
            txtSTYLE_CODE_PLM_SOURCE.Tag = ""

            Absx1.txtFor("STYLE_GROUP_CODE").Text = rowICTPLIN2.Item("STYLE_GROUP_CODE") & ""
            Absx1.txtFor("STYLE_CLASS_CODE").Text = rowICTPLIN2.Item("STYLE_CLASS_CODE") & ""
            Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTPLIN2.Item("SALES_DIVISION_CODE") & ""
            Absx1.txtFor("SEASON_CODE").Text = rowICTPLIN2.Item("SEASON_CODE") & ""
            Absx1.txtFor("ROYALTY_CODE").Text = rowICTPLIN2.Item("ROYALTY_CODE") & ""
            Absx1.txtFor("STYLE_DESC").Text = Mid(rowICTPLIN2.Item("STYLE_DESC") & "", 1, Absx1.txtFor("STYLE_DESC").MaxLength)
            Absx1.optFor("STYLE_STATUS").Value = "A"
            Absx1.txtFor("STYLE_UOM").Text = "EA"
            Absx1.numFor("SUB_UNIT_PACK_QTY").Value = 1
            Absx1.txtFor("STYLE_CODE_PLM").Text = STYLE_CODE_PLM
            Absx1.txtFor("IMAGE_NAME").Text = rowICTPLIN2.Item("SALES_DIVISION_CODE") & "\" & STYLE_CODE_PLM & ".JPG"
            'Absx1.numFor("STYLE_PRICE").Value = Val(rowICTPLIN2.Item("STYLE_PRICE") & "")
            'Absx1.numFor("STYLE_COST").Value = Val(rowICTPLIN2.Item("STYLE_COST") & "")
            ASCMAIN1.sql = "Select Max(DUTY_RATE_CODE) from ICTPLIN3 where STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "'"
            Dim DUTY_RATE_CODE As String = ASCDATA1.GetDataValue
            Absx1.txtFor("DUTY_RATE_CODE").Text = DUTY_RATE_CODE
            Absx1.txtFor("COUNTRY_CODE").Text = rowICTPLIN2.Item("COUNTRY_CODE") & ""
            Absx1.txtFor("VEND_CODE").Text = rowICTPLIN2.Item("VEND_CODE") & ""
            Absx1.txtFor("PURCH_NOTES").Text = rowICTPLIN2.Item("PURCH_NOTES") & ""

            Absx1.txtFor("CUST_CODE").Text = rowICTPLIN2.Item("CUST_CODE") & ""
            Absx1.txtFor("CUST_STYLE_CODE").Text = rowICTPLIN2.Item("CUST_STYLE_CODE") & ""

            Dim rowICTPLIN3 As DataRow = LookUp("ICTPLIN3", New String() {STYLE_CODE_PLM, 1})
            If rowICTPLIN3 IsNot Nothing Then
                Absx1.txtFor("STYLE_MATL_DESC").Text = rowICTPLIN3.Item("STYLE_CONTENT") & ""
            End If
            Absx1.txtFor("SIZE_SCALE").Text = rowICTPLIN2.Item("STYLE_NOTES") & ""

            If rowICTPLIN2.Item("VEND_CODE") & "" <> "" Then
                'If rowICTPLIN2.Item("VEND_ITEM_CODE") & "" <> "" And rowICTPLIN2.Item("VEND_CODE") & "" <> "" Then
                Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", rowICTPLIN2.Item("VEND_CODE") & "")
                If rowAPTVEND1 IsNot Nothing Then
                    Dim rowICTSTYV1 As DataRow = dst.Tables("ICTSTYV1").NewRow
                    rowICTSTYV1.Item("STYLE_CODE") = STYLE_CODE
                    rowICTSTYV1.Item("VEND_CODE") = rowICTPLIN2.Item("VEND_CODE") & ""
                    rowICTSTYV1.Item("VEND_NAME") = rowAPTVEND1.Item("VEND_NAME") & ""
                    rowICTSTYV1.Item("VEND_ITEM_CODE") = rowICTPLIN2.Item("VEND_ITEM_CODE") & ""
                    dst.Tables("ICTSTYV1").Rows.Add(rowICTSTYV1)
                End If
            End If

        End If

        ASCMAIN1.MultiTask_Release(, , 2)
    End Sub
    Public Overrides Function OK_to_do_View_Lookup(ByVal txtctl As UltraWinEditors.UltraTextEditor) As Boolean

        If txtctl.Name = "txtSTYLE_CODE_PLM_SOURCE" Then Return True

        If EntryMode = "" Then
            If htbkey_COLUMN_NAMEs.ContainsKey(Absx1.GetABSColumnName(txtctl)) Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function

    Private Sub numSUB_UNIT_PACK_QTY_ValueChanged(sender As System.Object, e As System.EventArgs) Handles numSUB_UNIT_PACK_QTY.ValueChanged
        'If (EntryMode = "Edit" Or EntryMode = "New") Then
        '    Dim SUB_PACK As Int64 = Val(numSUB_UNIT_PACK_QTY.Value & "")
        '    If SUB_PACK = 0 Then
        '        numSUB_UNIT_PACK_QTY.Value = 1
        '    Else
        '    End If
        'End If
        Dim SUB_PACK As Int64 = Val(numSUB_UNIT_PACK_QTY.Value & "")
        If SUB_PACK <= 1 Then
            chkSet.Checked = False

        ElseIf SUB_PACK > 1 Then
            chkSet.Checked = True
        End If

    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If Me.SELECTION_NO <> 0 Then
            Size_splICTSTYCI()
        End If
    End Sub

    Private Function PARTIALSTYLE(STYLE_CODE As String) As String
        If STYLE_CODE = "" Then Return ""
        Dim RETVAL As String = ""
        Dim NEW_STYLE As String = ""
        ASCMAIN1.sql = String.Format("SELECT COUNT(*) RECCNT FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
        Dim STYLE_COUNT As Int16 = Val(ASCDATA1.GetDataValue)
        If STYLE_COUNT = 1 Then
            ASCMAIN1.sql = String.Format("SELECT STYLE_CODE FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
            NEW_STYLE = ASCDATA1.GetDataValue
            RETVAL = NEW_STYLE
        End If
        Return RETVAL
    End Function

    Private Sub ICFSTYL1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            If (e.KeyCode = Keys.NumPad1 Or e.KeyCode = Keys.D1) And e.Alt Then
                Call Click_Command("Done", e)
            End If
        End If
    End Sub

    Private Sub grdICTSTYLD_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdICTSTYLD.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "PACK_CODE"
                grdCodeDesc(grdICTSTYLD, "ICTSTYLM", "PACK_CODE", "PACK_DESC")
        End Select
    End Sub

    Private Sub grdICTSTYLD_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdICTSTYLD.BeforeRowUpdate
        Dim row As DataRow = LookUp("ICTSTYLM", e.Row.Cells("PACK_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("STYLE_CODE").Value = Absx1.txtFor("STYLE_CODE").Text
        End If
    End Sub ' 
    Private Sub grdICTSTYLD_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdICTSTYLD.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "PACK_CODE"
                grdClickCellButton(grdICTSTYLD)
        End Select
    End Sub

    Private Sub imgSTYLE_DoubleClick(sender As Object, e As EventArgs) Handles imgSTYLE.DoubleClick
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
        Dim IMAGE_NAME As String = txtIMAGE_NAME.Text
        If Not FOLDER_NAME.EndsWith("\") Then
            FOLDER_NAME = FOLDER_NAME & "\"
        End If
        Dim frm As New ICFSIMG1(FOLDER_NAME & IMAGE_NAME)
        frm.ShowDialog()
    End Sub

    Private Sub chkSet_CheckedChanged(sender As Object, e As EventArgs) Handles chkSet.CheckedChanged
        If chkSet.Checked Then
            chkSet.Appearance.ForeColor = Color.Red
        Else
            chkSet.Appearance.ForeColor = Color.Empty
        End If
    End Sub

    Private Sub txtSTYLE_CODE_PLM_SOURCE_ValueChanged(sender As Object, e As EventArgs) Handles txtSTYLE_CODE_PLM_SOURCE.ValueChanged
        If txtSTYLE_CODE_PLM_SOURCE.Text = "" Then
            txtSTYLE_CODE_NEW.Text = ""
        Else
            If Not txtSTYLE_CODE_NEW.Text.StartsWith(txtSTYLE_CODE_PLM_SOURCE.Text) Then
                txtSTYLE_CODE_NEW.Text = txtSTYLE_CODE_PLM_SOURCE.Text
            End If
        End If
    End Sub

    Private Sub btnAutomatic_Click(sender As Object, e As EventArgs) Handles btnAutomatic.Click
        ASCMAIN1.sql = "Select * from ICTPLIN2 where SALES_DIVISION_CODE = 'JFA'" '  and INIT_DATE > TRUNC(SYSDATE)"
        '    ASCMAIN1.sql = "select * from ictplin2 where style_code_plm in (Select STYLE_CODE_PLM from ICTPLIN2 where SALES_DIVISION_CODE = 'JFA' MINUS SELECT STYLE_CODE_PLM FROM ICTSTYL1)"

        ' ASCMAIN1.sql &= " AND STYLE_CODE_PLM = '983-1000'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_CODE_PLM")
            Dim STYLE_CODE_PLM As String = row.Item(0)
            Dim STYLE_CARE_INST As String = row.Item("STYLE_CARE_INST")
            txtSTYLE_CODE_PLM_SOURCE.Text = STYLE_CODE_PLM

            Dim rowICTPLIN2 As DataRow = LookUp("ICTPLIN2", STYLE_CODE_PLM)
            Dim STYLE_NOTES As String = rowICTPLIN2.Item("STYLE_NOTES") & ""
            Dim N() = Split(STYLE_NOTES, vbCrLf)
            For NN As Integer = 0 To N.Length - 1
                N(NN) = Trim(N(NN))
            Next
            STYLE_NOTES = Join(N, vbCrLf)
            ASCDATA1.ExecuteSQL("Update ICTPLIN2 Set STYLE_NOTES = :PARM1 where STYLE_CODE_PLM = :PARM2", "VV", New String() {STYLE_NOTES, STYLE_CODE_PLM})

            Dim P As Integer = 0
            ASCMAIN1.sql = "SELECT DISTINCT PACK FROM JF_LINES WHERE STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "' ORDER BY PACK"
            Dim rows() As DataRow = ASCDATA1.GetDataTable.Select()
            For Each rowP As DataRow In rows
                ' JF + TAB NAME + CLASS CODE + SEQUENCE
                Dim PSFX As String = ""
                P += 1
                If rows.Length > 1 Then
                    PSFX = "P" & CStr(P)
                End If
                txtSTYLE_CODE_NEW.Text = STYLE_CODE_PLM & "JF" & STYLE_CARE_INST & PSFX

                ASCMAIN1.sql = "Select * from JF_LINES where STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "' AND PACK = '" & rowP.Item(0) & "'"
                Dim rowJF As DataRow = ASCDATA1.GetDataRow

                Create_Style_from_PLM(False)

                Absx1.numFor("INNER_PACK_QTY").Value = Val(rowP.Item(0))
                Absx1.numFor("CARTON_PACK_QTY").Value = Val(rowJF.Item("PACK")) * Val(rowJF.Item("PPQ"))

                Dim rowICTPLIN3 As DataRow = LookUp("ICTPLIN3", New String() {STYLE_CODE_PLM, 1})
                Dim DUTY_RATE_CODE As String = rowICTPLIN3.Item("DUTY_RATE_CODE") & ""
                'If DUTY_RATE_CODE <> "" Then
                '    DUTY_RATE_CODE = Mid(DUTY_RATE_CODE, 1, 4) & "." & Mid(DUTY_RATE_CODE, 5, 2) & "." & Mid(DUTY_RATE_CODE, 7, 4)
                'End If
                Absx1.txtFor("DUTY_RATE_CODE").Text = DUTY_RATE_CODE

                Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", rowICTPLIN2.Item("VEND_CODE") & "")
                If rowAPTVEND1 IsNot Nothing Then
                    Dim rowICTSTYV1 As DataRow = dst.Tables("ICTSTYV1").NewRow
                    rowICTSTYV1.Item("STYLE_CODE") = txtSTYLE_CODE_NEW.Text
                    rowICTSTYV1.Item("VEND_CODE") = rowICTPLIN2.Item("VEND_CODE") & ""
                    rowICTSTYV1.Item("VEND_NAME") = rowAPTVEND1.Item("VEND_NAME") & ""
                    rowICTSTYV1.Item("VEND_ITEM_CODE") = rowICTPLIN2.Item("VEND_ITEM_CODE") & ""
                    rowICTSTYV1.Item("PO_COST") = rowICTPLIN3.Item("PO_COST") & ""
                    rowICTSTYV1.Item("PO_COST_DATE") = Now.Date
                    dst.Tables("ICTSTYV1").Rows.Add(rowICTSTYV1)
                End If


                Click_Command("Update")

                If ScreenMode Then
                    Stop
                Else

                    ASCMAIN1.sql = "Insert into SOTPRIC2 (PRICE_LIST_CODE,STYLE_CODE,STYLE_PRICE,INIT_OPER,INIT_DATE)" & vbCrLf _
                        & " Values ('LOBLAW','" & txtSTYLE_CODE_NEW.Text & "', " & CStr(Val(rowJF.Item("JF_COST"))) & ", 'wjz', SYSDATE)"
                    ASCDATA1.ExecuteSQL()
                End If
            Next
        Next
    End Sub

    Private Sub btnGenerateUPCs_Click(sender As Object, e As EventArgs) Handles btnGenerateUPCs.Click

        If ASCMAIN1.CLIENT = "VAN" Then Exit Sub

        BeginTrans()

        ASCMAIN1.Progress("Now Generating UPCs")
        Dim R As Integer = 0

        ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE in (Select STYLE_CODE from ICTSTYL1 where CUST_CODE = 'LOBLAW')"

        For Each row As DataRow In ASCDATA1.GetDataTable().Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim UPC_CODE As String = row.Item("UPC_CODE") & ""
            ASCMAIN1.Progress("-", STYLE_CODE)
            If UPC_CODE <> "" Then
                Stop
            End If
            R += 1
            UPC_CODE = Get_UPC_Code(STYLE_CODE, COLOR_CODE)
            ASCMAIN1.sql = "Update ICTSTYC1 Set UPC_CODE = :PARM1 where STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3"
            Dim I As Integer = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {UPC_CODE, STYLE_CODE, COLOR_CODE})
            If I <> 1 Then
                Stop
            End If
        Next

        CommitTrans("Generated " & CStr(R) & " UPCs")
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdFixSizesColors_Click(sender As Object, e As EventArgs) Handles cmdFixSizesColors.Click
        Dim SCs As String = ""
        Dim SZs As String = ""
        Dim SQs As String = ""
        For I As Integer = 1 To 12
            Dim SIZE_CODE As String = Absx1.txtFor("SIZE_" & Format(I, "00")).Text
            If SIZE_CODE = "" Then
                Exit For
            End If
            SZs &= "-" & SIZE_CODE
            Dim QTY As Integer = Val(Absx1.numFor("QTY_" & Format(I, "00")).Value & "")
            SQs &= "/" & CStr(QTY)
        Next

        For Each ROW As DataRow In dst.Tables("ICTSTYC1").Select("", "COLOR_CODE")
            Dim COLOR_CODE = ROW.Item("COLOR_CODE")
            Dim STYLE_COLOR_DESC = ROW.Item("STYLE_COLOR_DESC") & ""
            SCs &= vbCrLf & COLOR_CODE & " " & STYLE_COLOR_DESC
        Next

        If SZs <> "" Then
            Dim SIZE_CODE As String = Absx1.txtFor("SIZE_CODE").Text
            If SIZE_CODE <> "" Then
                SCs = SIZE_CODE & " (" & Mid(SZs, 2) & " = " & Mid(SQs, 2) & ")" & SCs
            Else
                SCs = Mid(SZs, 2) & " = " & Mid(SQs, 2) & SCs
            End If

        Else
            SCs = Mid(SCs, 3)
        End If

        Absx1.txtFor("SIZE_SCALE").Text = SCs

    End Sub

    Private Sub cmdCalculateList_Click(sender As Object, e As EventArgs) Handles cmdCalculateList.Click
        If ASCMAIN1.CLIENT = "RGI" Then
            Synch_TABLE_NAME("ASTBASE1")
            Synch_TABLE_NAME("ICTSTYL1")
            Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
            Dim SILENT As Boolean = False
            Dim STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Style_Price(Me, SILENT, STYLE_CODE, rowASFBASE1)
            numSTYLE_PRICE_CALC.Value = STYLE_PRICE
        End If

    End Sub
End Class