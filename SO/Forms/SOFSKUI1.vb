Imports System.Net.Http
Imports System.Net.Http.Headers
'Imports Microsoft.Office.Interop.Excel

Public Class SOFSKUI1
    Dim Discounts As List(Of DISCOUNTS)
    Dim THEMEUSERS As String() = {"danny", "mariog", "regencyny", "tonyg", "wayne"}

#Region "ABS Standard Routines"

    ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Get_PARM("SOTPARM1")
        Dim SQLs As New Text.StringBuilder() With {.Length = 0}

        With dst
            Create_TDA(.Tables.Add, "ICTSTYL1", "*", 1)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT * FROM")
            SQLs.AppendLine("  (")
            SQLs.AppendLine("   SELECT C1.STYLE_CODE, C1.COLOR_CODE,")
            SQLs.AppendLine("   9999 AS ORDR_QTY,")
            SQLs.AppendLine("   C2.COLOR_DESC AS COLOR_CODE_LONG,")
            SQLs.AppendLine("   C1.STYLE_COLOR_STATUS,")
            SQLs.AppendLine("   CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) < 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE")
            SQLs.AppendLine("     WHEN 'MS'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END)")
            SQLs.AppendLine("   END AS MSOH,")
            SQLs.AppendLine("   CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) <= 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("     CASE WHEN")
            SQLs.AppendLine("       SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'MS'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) < 0")
            SQLs.AppendLine("     THEN")
            SQLs.AppendLine("       0")
            SQLs.AppendLine("     ELSE")
            SQLs.AppendLine("     SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'MS'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) END")
            SQLs.AppendLine("   END AS MSFT,")
            SQLs.AppendLine(" CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) < 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE")
            SQLs.AppendLine("     WHEN 'SW'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END)")
            SQLs.AppendLine("   END AS SWOH,")
            SQLs.AppendLine("   CASE WHEN")
            SQLs.AppendLine("   SUM(")
            SQLs.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            SQLs.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("     ELSE 0")
            SQLs.AppendLine("     END) <= 0")
            SQLs.AppendLine("   THEN")
            SQLs.AppendLine("     0")
            SQLs.AppendLine("   ELSE")
            SQLs.AppendLine("     CASE WHEN")
            SQLs.AppendLine("       SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'SW'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) < 0")
            SQLs.AppendLine("     THEN")
            SQLs.AppendLine("       0")
            SQLs.AppendLine("     ELSE")
            SQLs.AppendLine("     SUM(")
            SQLs.AppendLine("       CASE S2.WHSE_CODE")
            SQLs.AppendLine("       WHEN 'SW'")
            SQLs.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("       ELSE 0")
            SQLs.AppendLine("       END) END")
            SQLs.AppendLine("   END AS SWFT,")
            SQLs.AppendLine("   C1.THEME_CODE")
            SQLs.AppendLine("   FROM ICTSTYC1 C1")
            SQLs.AppendLine("   LEFT JOIN ICTSTAT2 S2")
            SQLs.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
            SQLs.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
            SQLs.AppendLine("   INNER JOIN ICTCOLR1 C2")
            SQLs.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
            SQLs.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS,  C1.THEME_CODE")
            SQLs.AppendLine("  )")
            SQLs.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') or (MSOH <> 0) or (MSFT <> 0) or (SWOH <> 0)  or (SWFT <> 0))")
            SQLs.AppendLine("  AND STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQLs.ToString
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False, "V", 2)
            'Fill_Records("ICTSTYC1", "", , ASCMAIN1.sql)
            .Tables("ICTSTYC1").Columns.Add("RSV", GetType(System.String))

            SQLs.Length = 0
            SQLs.AppendLine("SELECT WHSE_CODE, STYLE_CODE, COLOR_CODE,")
            'Changed from QTY_ATS_CUM to QTY_ATS per WZ - 7/10/14
            'This appears to be wrong per discussion with Rich - 10/23/14
            SQLs.AppendLine("STATUS_DATE, QTY_ATS_CUM as QTY_ATS_CUM")
            SQLs.AppendLine("FROM ICTSTDQ1")
            SQLs.AppendLine("WHERE STYLE_CODE  = :PARM1")
            SQLs.AppendLine("AND COLOR_CODE  = :PARM2")
            SQLs.AppendLine("AND WHSE_CODE  = :PARM3")
            ASCMAIN1.sql = SQLs.ToString
            Create_TDA(.Tables.Add, "ICTSTDQ1", "**", 0, False, "VVV")
            'Fill_Records("ICTSTDQ1", "", , ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "HANGTAG1", "**", 1, False)
            .Tables("HANGTAG1").Columns.Add("DATEPRINTED", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("BOXQTY", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("CARTQTY", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("COLORS", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("COLORS3", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price1_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price1_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("Price2_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price2_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("Price3_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price3_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("Price4_LBL", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("Price4_AMT", GetType(System.Double))
            .Tables("HANGTAG1").Columns.Add("COLORSDESC", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("VEND_SUPPLIER_ID", GetType(System.String))
            .Tables("HANGTAG1").Columns.Add("PORT_CODE_ORIG", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
            Create_TDA(.Tables.Add, "ICTSTYLD", "**", 1, False)
            Fill_Records("ICTSTYLD", "", , ASCMAIN1.sql)

            'If ASCMAIN1.Running_in_VS Then
            '    ASCMAIN1.sql = "SELECT * FROM ICTSTYC1_PICS"
            '    Create_TDA(.Tables.Add, "ICTSTYC1_PICS", "**", 0, True)
            'End If
            Create_TDA(.Tables.Add, "ICTSTAT2", "*", 3)

            SQLs.Length = 0
            SQLs.AppendLine("SELECT ICTSTYL3.*, ICTATTR1.ATTR_DESC")
            SQLs.AppendLine("FROM ICTSTYL3, ICTATTR1")
            SQLs.AppendLine("WHERE ICTSTYL3.ATTR_CODE = ICTATTR1.ATTR_CODE")
            SQLs.AppendLine("AND ICTSTYL3.STYLE_CODE = :PARM1")
            ASCMAIN1.sql = SQLs.ToString
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False, "V")

        End With

        grdICTSTYC1.DataSource = dst.Tables("ICTSTYC1")
        grdICTSTDQ1.DataSource = dst.Tables("ICTSTDQ1")
        grdICTSTYL3.DataSource = dst.Tables("ICTSTYL3")

        Sort_grdColumns(grdICTSTYC1, "COLOR_CODE".ToUpper, True)

        FilterColors("NONE")
        FilterAlloc("NONE", "NONE", "NONE")

        tab.Visible = False
        ClearStyle()

        If ASCMAIN1.Running_in_VS Then
            btnItemImg.Visible = True
        Else
            btnItemImg.Visible = False
        End If
        Absx1.txtFor("STYLE_CODE").Focus()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Done"
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If
        Call Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Cancel", "Done"
                Call Mode_Settings(False)
            Case "New Scan (Alt-1)"
                ClearStyle()
            Case "Hang Tag"
                PrintHangTag()
                Absx1.txtFor("STYLE_CODE").Focus()
            Case "Find Style by Attribute"
                Dim ms As New Text.StringBuilder With {.Length = 0}
                ms.AppendLine("This Feature Has Been Moved To")
                ms.AppendLine("It's Own Screen That Is Now Avalable")
                ms.AppendLine(String.Format("On The Main Menu As {0}Search By Attribute{0}", Chr(34)))
                MsgBox(ms.ToString, vbOKOnly, "This Feature Has Moved")
                'Dim STYLE_CODE_selected As String = ""
                'Using F As New TAC.ICFATTR2(Me)
                '    Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
                '    If Not IsNothing(rowSOTPARM3) Then
                '        If rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString.EndsWith("\") Then
                '            F.IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
                '        Else
                '            F.IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString & "\"
                '        End If
                '        F.rbadDir = rowSOTPARM3.Item("RO_PARM_EXCEL_DIR").ToString()
                '    Else
                '        F.IMAGES_FOLDER = "C: \"
                '        F.rbadDir = "C:\"
                '    End If
                '    F.ShowDialog()
                '    STYLE_CODE_selected = F.STYLE_CODE
                'End Using
                'If STYLE_CODE_selected <> "" Then
                '    txtSTYLE_CODE.Text = STYLE_CODE_selected
                '    Click_Command("New Scan (Alt-1)")
                'End If
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Allocation").Visible = False
                .Groups("Image").Visible = False
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            'With grdARTPYMT2.DisplayLayout.Override
            '    If This_Record_Inquiry_Only Then
            '        .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '        .AllowDelete = DefaultableBoolean.False
            '        .AllowUpdate = DefaultableBoolean.False
            '    Else
            '        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            '        .AllowDelete = DefaultableBoolean.True
            '        .AllowUpdate = DefaultableBoolean.True
            '    End If
            'End With
            Clear_Record()
        Else
            Clear_Record()
        End If
        Absx1.txtFor("STYLE_CODE").Focus()
    End Sub

    Sub Clear_Record()
        txtUPC_CODE.Text = ""
        lblUPC_CODE.Text = "UPC Code"
        Absx1.txtFor("STYLE_CODE").Focus()
        lblEXCLUSIVE_STYLE.Visible = False
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)
        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
        ClearStyle()
    End Sub

    Sub Delete_Record()
        'Call BeginTrans()
        'Call Delete_Rows("SOTORDR1")
        'Call CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        'Call BeginTrans()
        'Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record()
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                If Absx1.txtFor("STYLE_CODE").Text.Length = 0 Then
                    'ClearStyle()
                Else
                    If VERIFYSTYLE(Absx1.txtFor("STYLE_CODE").Text) Then
                        EnforceConstraints(False)
                        Fill_Records("ICTSTYL1", Absx1.txtFor("STYLE_CODE").Text, True)
                        If dst.Tables.Item("ICTSTYL1").Rows.Count = 1 Then
                            lblEXCLUSIVE_STYLE.Visible = dst.Tables.Item("ICTSTYL1").Rows(0).Item("EXCLUSIVE_STYLE").ToString & "" = "1"
                        End If
                        Fill_Records("ICTSTYL3", Absx1.txtFor("STYLE_CODE").Text, True)
                        EnforceConstraints(True)
                        If dst.Tables("ICTSTYL1").Rows.Count() = 0 Then
                            MsgBox("Style Not Found In Masterfile", MsgBoxStyle.Critical, "Invalid Style")
                            ClearStyle()
                            Exit Sub
                        End If
                        FilterColors(Absx1.txtFor("STYLE_CODE").Text)
                        If grdICTSTYC1.Rows.Count = 0 Then
                            MsgBox("Selected Style Does Not Have Colors Set-up", MsgBoxStyle.Critical, "Invalid Style")
                            ClearStyle()
                            Exit Sub
                        End If
                        Bind_Controls(grpSTYL1, "ICTSTYL1")
                        Bind_Controls(grpICTSTYL3, "ICTSTYL1")
                        UltraExplorerBar1.Groups("Allocation").Visible = True
                        Dim rowARTCUST1 As DataRow = Nothing
                        Discounts = SOCMAIN2.Price_Discounts(Me, "", rowARTCUST1, Absx1.txtFor("STYLE_CODE").Text, False)
                        For i As Integer = 1 To 4
                            If Discounts(i - 1).DISCOUNT_QTY = 0 Then
                                Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = False
                                Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = False
                                Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = False
                                Absx1.txtFor(String.Format("priceDISC{0}", i)).Visible = False
                            Else
                                Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = True
                                Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = True
                                Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = True
                                Absx1.txtFor(String.Format("priceDISC{0}", i)).Visible = True
                                Absx1.CtlFor(String.Format("lblDISC{0}", i)).Text = Discounts(i - 1).DISCOUNT_DESC
                                Absx1.CtlFor(String.Format("lblDISC{0}", i)).Tag = Discounts(i - 1).DISCOUNT_PCT 'Use for hover over.
                                Absx1.txtFor(String.Format("qtyDISC{0}", i)).Text = Discounts(i - 1).DISCOUNT_QTY
                                Absx1.txtFor(String.Format("priceDISC{0}", i)).Text = Format$(Discounts(i - 1).DISCOUNT_PRICE, "###,##0.00")
                            End If
                        Next
                        txtFactory.Text = GetVendorData(dst.Tables("ICTSTYL1").Rows(0).Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
                        txtPort.Text = GetVendorData(dst.Tables("ICTSTYL1").Rows(0).Item("VEND_CODE").ToString, "PORT_CODE")
                        Dim STYLE_CLASS_CODE As String = dst.Tables("ICTSTYL1").Rows(0).Item("STYLE_CLASS_CODE").ToString
                        SetClassTip(STYLE_CLASS_CODE)
                        SetStyleColor()
                        lblSTATUS.Visible = True
                        Sort_grdColumns(grdICTSTYC1, "COLOR_CODE".ToUpper, True)
                        ShowPromo(Absx1.txtFor("STYLE_CODE").Text)
                        Dim FEFD As New FEFDPrice(Me, Absx1.txtFor("STYLE_CODE").Text)
                        If FEFD.ErrorMsg = "" Then
                            btnFEPrice.Text = Format(FEFD.FEPrice, "###,##0.00")
                            btnFEMixPrice.Text = Format(FEFD.FEMixPrice, "###,##0.00")
                            btnFDMixPrice.Text = Format(FEFD.FDMixPrice, "###,##0.00")
                            btnFDPrice.Text = Format(FEFD.FDPrice, "###,##0.00")
                            Absx1.txtFor("STYLE_CODE").Focus()
                            txtSTYLE_CODE.Select(0, txtSTYLE_CODE.Text.Length)
                        Else
                            MsgBox(FEFD.ErrorMsg, MsgBoxStyle.Critical, "Error With Style")
                            ClearStyle()
                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "These Properties and Methods May Have Aplication In Order Entry As Well"
    Private Sub ShowDisc(sender As Object, e As System.EventArgs)
        Dim this As Infragistics.Win.Misc.UltraLabel = sender
        Dim tt As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo = New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo() With {.ToolTipText = this.Tag}
        tip.SetUltraToolTip(sender, tt)
        tip.AutoPopDelay = 3000
        tip.InitialDelay = 3000
        tip.DisplayStyle = ToolTipDisplayStyle.BalloonTip
        tip.ShowToolTip(sender)
    End Sub

    Private Sub FilterColors(STYLE_CODE As String)
        'Dim dvw As DataView = DirectCast(grdICTSTYC1.DataSource, DataTable).DefaultView
        'dvw.RowFilter = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        Dim SUpdate As Boolean = UpdateICTSTAT2(STYLE_CODE, "MS")
        Fill_Records("ICTSTYC1", STYLE_CODE, True)
        For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select()
            rowICTSTYC1.Item("ORDR_QTY") = 0
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT SUM(NVL(RSRV_QTY_OPEN,0)) RSV")
            SQLS.AppendLine("FROM SOTRSRV2")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", rowICTSTYC1.Item("STYLE_CODE").ToString))
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", rowICTSTYC1.Item("COLOR_CODE").ToString))
            ASCMAIN1.sql = SQLS.ToString()
            Dim RSV As Int16 = Val(ASCDATA1.GetDataValue)
            rowICTSTYC1.Item("RSV") = RSV
        Next
        grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item(1).Header.Appearance.BackColor = Drawing.Color.Yellow
        grdICTSTYC1.UpdateData()
        grdICTSTYC1.Text = String.Format("Colors For Style{0}", STYLE_CODE)
        grdICTSTYC1.Visible = True
        If SUpdate Then
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item(1).Header.Appearance.BackColor = Drawing.Color.Green
        Else
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item(1).Header.Appearance.BackColor = Drawing.Color.Yellow
        End If
        SetColorStatusImages(0, False)
    End Sub

    Private Sub ClearStyle()
        dst.Tables("ICTSTYL1").Clear()
        dst.Tables("ICTSTYL3").Clear()
        Bind_Controls(grpSTYL1, "ICTSTYL1")
        Bind_Controls(grpICTSTYL3, "ICTSTYL1")
        FilterColors("NONE")
        UltraExplorerBar1.Groups("Image").Visible = False
        UltraExplorerBar1.Groups("Allocation").Visible = False
        For i As Integer = 1 To 4
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Text = ""
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Tag = "Select A Style"
            Absx1.txtFor(String.Format("qtyDISC{0}", i)).Text = ""
            Absx1.txtFor(String.Format("priceDISC{0}", i)).Text = ""
            Absx1.CtlFor(String.Format("lblDISC{0}", i)).Visible = True
            Absx1.CtlFor(String.Format("lblDISC{0}QP", i)).Visible = True
            Absx1.txtFor(String.Format("qtyDISC{0}", i)).Visible = True
            Absx1.txtFor(String.Format("priceDISC{0}", i)).Visible = True
        Next
        txtPort.Text = ""
        txtFactory.Text = ""
        lblSTATUS.Visible = False
        SetClassTip("")
        SetStyleColor()
        Absx1.txtFor("STYLE_CODE").Focus()
        SetColorStatusImages(0, True)
        lblEXCLUSIVE_STYLE.Visible = False
        lblPromo.Visible = False
        lblPromo.Text = ""
        btnShowPromo.Visible = False
    End Sub

    Private Function GetImageLocation(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = ""
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim RO_PARM_STYLE_IMG_DIR As String = ""
        Dim FileMatch As String
        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
        Dim COLOR_CODE_LONG As String = ""
        If Not IsNothing(rowICTCOLR1) Then
            COLOR_CODE_LONG = rowICTCOLR1.Item("COLOR_CODE_LONG").ToString()
        End If
        Dim WebVal As String = ""
        If chkLIVEDATA.Checked Then
            'TryWebImage(String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE))
        End If
        If Not IsNothing(rowSOTPARM3) Then
            RO_PARM_STYLE_IMG_DIR = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            If RO_PARM_STYLE_IMG_DIR.Length > 0 Then
                FileMatch = Dir(String.Format("{0}\{1}-{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                If FileMatch.Length > 0 Then
                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                Else
                    FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                    If FileMatch.Length > 0 Then
                        RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                    Else
                        FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE_LONG))
                        If FileMatch.Length > 0 Then
                            RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                        Else
                            FileMatch = Dir(String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                            If FileMatch.Length > 0 Then
                                RetVal = String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE)
                            Else
                                FileMatch = Dir(String.Format("{0}\{1}*", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                                If FileMatch.Length > 0 Then
                                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, FileMatch)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Try
            If WebVal.Length > 0 Then
                Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(WebVal)
                Dim response As System.Net.WebResponse = req.GetResponse()
                Dim stream As IO.Stream = response.GetResponseStream()
                Dim img As System.Drawing.Image = System.Drawing.Image.FromStream(stream)
                stream.Close()
                If System.IO.File.Exists(RetVal) Then
                    System.IO.File.Delete(RetVal)
                    img.Save(RetVal)
                Else
                    RetVal = String.Format("{0}\{1}", RO_PARM_STYLE_IMG_DIR, String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE))
                    img.Save(RetVal)
                End If
            End If
        Catch ex As Exception
        End Try
        Return RetVal
    End Function

    Private Sub imgSTYL1_DoubleClick(sender As Object, e As System.EventArgs) Handles imgSTYL1.DoubleClick
        'Dim frmSOFIMGV1 As New SOFIMGV1(Me, imgSTYL1.ImageLocation)
        'frmSOFIMGV1.Show()
        Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
        Dim frmIMAGE As New TAC.TAFIMGV1(Me, STYLE_CODE, COLOR_CODE, "L")
        With frmIMAGE
            .ShowDialog(Me)
        End With
    End Sub

    Private Function GetVendorData(ByVal VEND_CODE As String, ByVal COLUMN As String) As String
        Dim RetVal As String = ""
        If VEND_CODE.Length > 0 And COLUMN.Length > 0 Then
            ASCMAIN1.sql = String.Format("SELECT {0} FROM APTVEND1 WHERE VEND_CODE = '{1}'", COLUMN, VEND_CODE)
            RetVal = ASCDATA1.GetDataValue
        End If
        Return RetVal
    End Function

    Private Sub FORM_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        'Stop
        If (e.KeyCode = Keys.NumPad1 Or e.KeyCode = Keys.D1) And e.Alt Then
            Call Click_Command("New Scan (Alt-1)", e)
        End If
        If (e.KeyCode = Keys.NumPad2 Or e.KeyCode = Keys.D2) And e.Alt Then
            Call Click_Command("Hang Tag", e)
        End If
        If e.KeyData = Keys.Enter Then
            If Absx1.txtFor("STYLE_CODE").Focused Then
                grdICTSTYC1.Focus()
            End If
        End If
    End Sub

    Private Sub PrintHangTag()
        If Absx1.txtFor("STYLE_CODE").Text.Length > 0 Then
            ASCMAIN1.sql = "select ro_parm_lbl_printer from sotparm3 where ro_parm_key = 'Z'"
            Dim Printer_name As String = ASCDATA1.GetDataValue
            If Printer_name.Length = 0 Then
                MsgBox("Printer Not Defined In Parameters", vbOKOnly, "Printer Definition")
            Else
                Dim HANGTAG As New HANGTAG(Me, Absx1.txtFor("STYLE_CODE").Text, Discounts, Printer_name)
                If HANGTAG.ErrMsg.Length = 0 Then
                    HANGTAG.Print()
                End If
            End If
        End If
    End Sub

#Region "Hover Over Functionality"
    Private Sub lblDISC1_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC1.MouseHover
        ShowDisc(sender, e)
    End Sub

    Private Sub lblDISC2_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC2.MouseHover
        ShowDisc(sender, e)
    End Sub

    Private Sub lblDISC3_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC3.MouseHover
        ShowDisc(sender, e)
    End Sub

    Private Sub lblDISC4_MouseHover(sender As Object, e As System.EventArgs) Handles lblDISC4.MouseHover
        ShowDisc(sender, e)
    End Sub
#End Region
#End Region

    Private Sub grdICTSTYC1_AfterRowRegionScroll(sender As Object, e As UltraWinGrid.RowScrollRegionEventArgs) Handles grdICTSTYC1.AfterRowRegionScroll
        SetColorStatusImages(e.RowScrollRegion.VisibleRows(0).Row.Index, False)
    End Sub

    Private Sub grdICTSTYC1_BeforeEnterEditMode(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles grdICTSTYC1.BeforeEnterEditMode
        e.Cancel = True
    End Sub

    Private Sub grdICTSTYC1_Click(sender As Object, e As System.EventArgs) Handles grdICTSTYC1.Click
        If Not IsNothing(grdICTSTYC1.ActiveRow) Then
            grdICTSTYC1.ActiveRow.Cells("ORDR_QTY").Activated = True
        End If
    End Sub

    Private Sub grdICTSTYC1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTYC1.AfterRowActivate
        If grdICTSTYC1.ActiveRow Is Nothing Then
            UltraExplorerBar1.Groups("Allocation").Visible = False
        Else
            Dim STYLE_CODE As String = grdICTSTYC1.ActiveRow.Cells("STYLE_CODE").Value
            Dim COLOR_CODE As String = grdICTSTYC1.ActiveRow.Cells("COLOR_CODE").Value
            txtUPC_CODE.Text = GetUPCCODE(STYLE_CODE, COLOR_CODE)
            If txtUPC_CODE.Text.Length > 0 Then
                lblUPC_CODE.Text = "UPC Code For " & COLOR_CODE
            Else
                lblUPC_CODE.Text = "UPC Code"
            End If
            Dim WHSE_CODE As String = "MS"
            FilterAlloc(STYLE_CODE, COLOR_CODE, WHSE_CODE)
            UltraExplorerBar1.Groups("Allocation").Visible = True
            imgSTYL1.ImageLocation = GetImageLocation(STYLE_CODE, COLOR_CODE)
            If imgSTYL1.ImageLocation.Length > 0 Then
                UltraExplorerBar1.Groups("Image").Visible = True
            Else
                UltraExplorerBar1.Groups("Image").Visible = False
            End If
        End If
    End Sub

    Private Sub FilterAlloc(STYLE_CODE As String, COLOR_CODE As String, Optional WHSE_CODE As String = "MS")
        'Dim dvw As DataView = DirectCast(grdICTSTDQ1.DataSource, DataTable).DefaultView
        'dvw.RowFilter = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}' AND WHSE_CODE = '{2}'", STYLE_CODE, COLOR_CODE, WHSE_CODE)
        'grdICTSTDQ1.Text = String.Format("Colors For Style{0}", STYLE_CODE)
        Fill_Records("ICTSTDQ1", New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE}, True)
        'Removed 6/3/21 so Sales Reps can now see full allocation - WHR
        'Dim S As New System.Text.StringBuilder() With {.Length = 0}
        'S.AppendLine("SELECT")
        'S.AppendLine("MSOH, MSFT")
        'S.AppendLine("FROM")
        'S.AppendLine("  (")
        'S.AppendLine("   SELECT")
        'S.AppendLine("   C1.STYLE_CODE,")
        'S.AppendLine("   C1.COLOR_CODE,")
        'S.AppendLine("   C1.STYLE_COLOR_STATUS,")
        'S.AppendLine("   CASE WHEN")
        'S.AppendLine("   SUM(")
        'S.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        'S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        'S.AppendLine("     ELSE 0")
        'S.AppendLine("     END) < 0")
        'S.AppendLine("   THEN")
        'S.AppendLine("     0")
        'S.AppendLine("   ELSE")
        'S.AppendLine("   SUM(")
        'S.AppendLine("     CASE S2.WHSE_CODE")
        'S.AppendLine("     WHEN 'MS'")
        'S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        'S.AppendLine("     ELSE 0")
        'S.AppendLine("     END)")
        'S.AppendLine("   END AS MSOH,")
        'S.AppendLine("   CASE WHEN")
        'S.AppendLine("   SUM(")
        'S.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        'S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'S.AppendLine("     ELSE 0")
        'S.AppendLine("     END) <= 0")
        'S.AppendLine("   THEN")
        'S.AppendLine("     0")
        'S.AppendLine("   ELSE")
        'S.AppendLine("     CASE WHEN")
        'S.AppendLine("       SUM(")
        'S.AppendLine("       CASE S2.WHSE_CODE")
        'S.AppendLine("       WHEN 'MS'")
        'S.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'S.AppendLine("       ELSE 0")
        'S.AppendLine("       END) < 0")
        'S.AppendLine("     THEN")
        'S.AppendLine("       0")
        'S.AppendLine("     ELSE")
        'S.AppendLine("     SUM(")
        'S.AppendLine("       CASE S2.WHSE_CODE")
        'S.AppendLine("       WHEN 'MS'")
        'S.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        'S.AppendLine("       ELSE 0")
        'S.AppendLine("       END) END")
        'S.AppendLine("   END AS MSFT")
        'S.AppendLine("   FROM ICTSTYC1 C1")
        'S.AppendLine("   LEFT JOIN ICTSTAT2 S2")
        'S.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
        'S.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
        'S.AppendLine("   INNER JOIN ICTCOLR1 C2")
        'S.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
        'S.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
        'S.AppendLine("  )")
        'S.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') OR (MSOH <> 0) OR (MSFT <> 0))")
        'S.AppendLine("  AND STYLE_CODE = :PARM1")
        'S.AppendLine("  AND COLOR_CODE = :PARM2")
        'Dim tbl As DataTable = ASCDATA1.GetDataTable(S.ToString(), String.Empty, "VV", New Object() {STYLE_CODE, COLOR_CODE})
        'For Each rowFUTURE As DataRow In tbl.Rows
        '    If Val(rowFUTURE.Item("MSOH").ToString & "") = 0 Then
        '        For Each rowICTSTDQ1 As DataRow In dst.Tables("ICTSTDQ1").Select("", "STATUS_DATE")
        '            If dst.Tables("ICTSTDQ1").Rows.Count > 1 Then
        '                rowICTSTDQ1.Delete()
        '                dst.Tables("ICTSTDQ1").AcceptChanges()
        '            End If
        '        Next
        '    End If
        'Next
        grdICTSTDQ1.Visible = True
        grdICTSTDQ1.DisplayLayout.Bands(0).Columns.Item(3).Header.Appearance.BackColor = Drawing.Color.Yellow
    End Sub

    Private Function GetUPCCODE(STYLE_CODE As String, COLOR_CODE As String) As String
        Dim RetVal As String = ""
        Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
        If Not IsNothing(rowICTSTYC1) Then
            If Not IsDBNull(rowICTSTYC1.Item("UPC_CODE")) Then
                RetVal = rowICTSTYC1.Item("UPC_CODE")
            End If
        End If
        Return RetVal
    End Function

    Private Function VERIFYSTYLE(STYLE_CODE As String) As Boolean
        Dim RETVAL As Boolean = False
        Dim NEW_STYLE As String = ""
        ASCMAIN1.sql = String.Format("SELECT COUNT(*) RECCNT FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
        Dim STYLE_COUNT As Int16 = Val(ASCDATA1.GetDataValue)
        If STYLE_COUNT = 1 Then
            ASCMAIN1.sql = String.Format("SELECT STYLE_CODE FROM ictstyl1 WHERE STYLE_CODE LIKE '%{0}'", STYLE_CODE)
            NEW_STYLE = ASCDATA1.GetDataValue
            If Absx1.txtFor("STYLE_CODE").Text <> NEW_STYLE Then
                Absx1.txtFor("STYLE_CODE").Text = NEW_STYLE
            End If
            RETVAL = True
        Else
            MsgBox("Style Not Found In Masterfile", MsgBoxStyle.Critical, "Invalid Style")
            ClearStyle()
            Exit Function
        End If
        Return RETVAL
    End Function

    Private Sub ShowClassTip(sender As Object, e As System.EventArgs)
        tip.AutoPopDelay = 3000
        tip.InitialDelay = 3000
        tip.DisplayStyle = ToolTipDisplayStyle.BalloonTip
        tip.ShowToolTip(sender)
    End Sub

    Private Sub txtSTYLE_CODE_MouseEnter(sender As Object, e As System.EventArgs) Handles txtSTYLE_CODE.MouseEnter
        ShowClassTip(sender, e)
    End Sub

    Private Sub SetClassTip(ByVal STYLE_CLASS_CODE As String)
        Dim tt As Infragistics.Win.UltraWinToolTip.UltraToolTipInfo =
                                New Infragistics.Win.UltraWinToolTip.UltraToolTipInfo() _
                                With {.ToolTipText = STYLE_CLASS_CODE}
        tip.SetUltraToolTip(txtSTYLE_CODE, tt)
    End Sub

    Private Sub SetStyleColor()
        If dst.Tables("ICTSTYL1").Rows.Count > 0 Then
            Dim STYLE_STATUS As String = dst.Tables("ICTSTYL1").Rows(0).Item("STYLE_STATUS").ToString()
            Select Case STYLE_STATUS
                Case Is = "D"
                    txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Red
                Case Is = "N"
                    txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Yellow
                Case Else
                    txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Empty
            End Select
        Else
            txtSTYLE_CODE.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdICTSTYC1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYC1.InitializeRow
        With e.Row
            Dim STYLE_COLOR_STATUS As String = .Cells("STYLE_COLOR_STATUS").Value & ""
            Select Case STYLE_COLOR_STATUS
                Case "D"
                    .Appearance.BackColor = Drawing.Color.Red
                Case "N"
                    .Appearance.BackColor = Drawing.Color.Yellow
                Case Else
                    .Appearance.BackColor = Drawing.Color.Empty
            End Select
        End With
    End Sub

    Private Sub btnItemImg_Click(sender As System.Object, e As System.EventArgs) Handles btnItemImg.Click
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("Truncate table ICTSTYC1_PICS")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        SQLS.Length = 0
        SQLS.AppendLine("INSERT INTO ICTSTYC1_PICS SELECT STYLE_CODE, COLOR_CODE, NULL AS STYLE_COLOR_IMAGE_NAME FROM ICTSTYC1")
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()

        Fill_Records("ICTSTYC1_PICS")
        SQLS.Length = 0
        SQLS.AppendLine("Select Count(*) from ICTSTYC1")
        ASCMAIN1.sql = SQLS.ToString()
        Dim TotalRecs As Int16 = Val(ASCDATA1.GetDataValue)
        Dim CurrentRec As Int16 = 0
        For Each rowICTSTYC1_PICS As DataRow In dst.Tables("ICTSTYC1_PICS").Select()
            CurrentRec += 1
            Dim STYLE_CODE As String = rowICTSTYC1_PICS.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowICTSTYC1_PICS.Item("COLOR_CODE")
            rowICTSTYC1_PICS.Item("STYLE_COLOR_IMAGE_NAME") = Replace(GetImageLocation(STYLE_CODE, COLOR_CODE), "E:\MASTER ITEM PHOTO FOLDER\", "")
            ASCMAIN1.Progress(Format(((CurrentRec / TotalRecs)), "###,##0.0%") & " Complete")
        Next
        Call BeginTrans()
        Call Update_Record_TDA("ICTSTYC1_PICS")
        Call CommitTrans("Update Complete")
    End Sub

    Private Function TryWebImage(ImageName As String) As String
        Dim API_BASE As String = ""
        'Dim url As New System.Uri("http://50.75.200.254:8181/images/product/" & ImageName)
        Dim url As New System.Uri("http://api.regency-rib.com:8181/images/product/" & ImageName)
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
        req.Timeout = 2000
        Dim resptest As System.Net.WebResponse
        Dim ErrorsFound As Boolean = False
        Try
            resptest = req.GetResponse()
            'ImageName = "http://50.75.200.254:8181/images/product/" & ImageName
            ImageName = "http://api.regency-rib.com:8181/images/product/" & ImageName
        Catch ex As Exception
            'Try
            '    Dim url2 As New System.Uri("http://192.168.110.224:8181/images/product/" & ImageName)
            '    'Dim url2 As New System.Uri("http://api.regency-rib.com/images/product/" & ImageName)
            '    Dim req2 As System.Net.WebRequest = System.Net.WebRequest.Create(url2)
            '    req2.Timeout = 2000
            '    resptest = req2.GetResponse()
            '    ImageName = "http://192.168.110.224:8181/images/product/" & ImageName
            '    'ImageName = "http://api.regency-rib.com/images/product/" & ImageName
            '    resptest.Close()
            '    req2 = Nothing
            'Catch ex2 As Exception
            ErrorsFound = True
            ImageName = ""
            req = Nothing
            'End Try
        End Try
        Return ImageName
    End Function

    Function UpdateICTSTAT2(ByVal STYLE_CODE As String, ByVal WHSE_CODE As String) As Boolean
        Dim iresult As Boolean = False
        If STYLE_CODE = "NONE" Or chkLIVEDATA.Checked = False Then
            Return iresult
            Exit Function
        Else
            Dim API_BASE As String = ""
            'Dim url As New System.Uri("http://50.75.200.254:8181/")
            Dim url As New System.Uri("http://api.regency-rib.com:8181/")
            Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
            Dim resptest As System.Net.WebResponse
            req.Timeout = 2000
            Try
                resptest = req.GetResponse()
                resptest.Close()
                req = Nothing
                'API_BASE = "http://50.75.200.254:8181/"
                API_BASE = "http://api.regency-rib.com:8181/"
            Catch ex As Exception
                req = Nothing
                'API_BASE = "http://192.168.110.224:8181/"
                API_BASE = "http://api.regency-rib.com:8181/"
            End Try

            Dim API_CONTROLLER As String = "api/SalesOrder/GetICTSTAT2"
            Dim API_QUERY_STRING As String = String.Format("?STYLE_CODE={0}&COLOR_CODE=&WHSE_CODE={1}&CSQL=", STYLE_CODE, WHSE_CODE)

            Dim client As New HttpClient()
            client.Timeout = TimeSpan.FromSeconds(2)
            client.BaseAddress = New Uri(API_BASE)

            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            Dim response As HttpResponseMessage
            Try
                response = client.GetAsync(API_CONTROLLER & API_QUERY_STRING).Result
            Catch ex As Exception
                Return iresult
                Exit Function
            End Try

            If response.IsSuccessStatusCode Then
                Try
                    Dim ICTSTAT2_RESPONSE As New List(Of ICTSTAT2)
                    Dim apiResponseString As String = ""
                    Dim responseObject As Object = response.Content.ReadAsAsync(Of IEnumerable(Of ICTSTAT2))().Result

                    ICTSTAT2_RESPONSE = responseObject
                    'Stop
                    For Each STAT As ICTSTAT2 In ICTSTAT2_RESPONSE
                        Fill_Record("ICTSTAT2", New String() {STAT.m_STYLE_CODE, STAT.m_COLOR_CODE, STAT.m_WHSE_CODE}, False, True)
                        If dst.Tables.Item("ICTSTAT2").Rows.Count > 0 Then
                            'dst.Tables.Item("ICTSTAT2").Rows(0).Item("INIT_DATE") = STAT.m_INIT_DATE
                            'dst.Tables.Item("ICTSTAT2").Rows(0).Item("LAST_DATE") = STAT.m_LAST_DATE
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_ALLO") = STAT.m_WHSE_QTY_ALLO
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_COMM") = STAT.m_WHSE_QTY_COMM
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_ON_HAND") = STAT.m_WHSE_QTY_ON_HAND
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_ON_ORDER") = STAT.m_WHSE_QTY_ON_ORDER
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_OPEN") = STAT.m_WHSE_QTY_OPEN
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_PICK") = STAT.m_WHSE_QTY_PICK
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_PROD") = STAT.m_WHSE_QTY_PROD
                            dst.Tables.Item("ICTSTAT2").Rows(0).Item("WHSE_QTY_TRAN") = STAT.m_WHSE_QTY_TRAN
                        End If
                        Call Update_Record_TDA("ICTSTAT2")
                    Next
                    iresult = True
                Catch ex As Exception
                    iresult = False
                End Try

            Else
                iresult = False
            End If
        End If
        Return iresult
    End Function

    Private Sub chkLongColors_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkLongColors.CheckedChanged
        SetColorCodeView()
    End Sub

    Private Sub SetColorCodeView()
        grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE_LONG").Width = grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Width
        grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE_LONG").Header.Appearance.BackColor = grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Header.Appearance.BackColor
        If chkLongColors.Checked Then
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Hidden = True
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE_LONG").Hidden = False
        Else
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Hidden = False
            grdICTSTYC1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE_LONG").Hidden = True
        End If
    End Sub

    Private Sub grdICTSTYC1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSTYC1.ClickCellButton
        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case Is = "THEME_CODE"
                    If THEMEUSERS.Contains(ASCMAIN1.USER_ID) Then
                        grdClickCellButton(grdICTSTYC1, , , , "THEME_CODE")
                        Dim STYLE_CODE As String = e.Cell.Row.Cells.Item("STYLE_CODE").Text
                        Dim COLOR_CODE As String = e.Cell.Row.Cells.Item("COLOR_CODE").Text
                        Dim THEME_CODE As String = e.Cell.Row.Cells.Item("THEME_CODE").Text
                        Dim SQLS As New System.Text.StringBuilder() With {.Length = 0}
                        SQLS.AppendLine("DELETE FROM ICTTHEMS")
                        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                        SQLS.Length = 0
                        SQLS.AppendLine("INSERT INTO ICTTHEMS")
                        SQLS.AppendLine(String.Format("VALUES ('{0}','{1}','{2}')", STYLE_CODE, COLOR_CODE, THEME_CODE))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                        SQLS.Length = 0
                        SQLS.AppendLine("UPDATE ICTSTYC1")
                        SQLS.AppendLine(String.Format("SET THEME_CODE = '{0}'", THEME_CODE))
                        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                        ASCMAIN1.sql = SQLS.ToString
                        ASCDATA1.ExecuteSQL()
                    End If
            End Select
        End With
    End Sub

    Private Sub SetColorStatusImages(ByVal StartAt As Integer, ByVal ClearAll As Boolean)
        Dim Path As String = ASCMAIN1.Folders("Images").ToString()
        Dim active As String = Path & "16\ball_green.png"
        Dim discontinued As String = Path & "16\ball_red.png"
        Dim noReorder As String = Path & "16\ball_yellow.png"
        Dim posStart As Integer = PictureBox1.Top
        Dim posStep As Integer = 21
        If ClearAll Then
            For i As Integer = 1 To 19
                Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & i.ToString())
                pBox.Visible = False
                pBox.ImageLocation = ""
            Next i
        Else

            Dim firstRow As Integer = 0
            For i As Integer = 1 To 19
                Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & i.ToString())
                pBox.Top = posStart + (posStep * (i - 1))
            Next i

            Dim curRow As Integer = 0
            For Each grow As UltraWinGrid.UltraGridRow In grdICTSTYC1.Rows
                curRow += 1
                If grow.Index >= StartAt Then
                    If (curRow - StartAt) < 20 Then
                        Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & (curRow - StartAt).ToString())
                        Select Case grow.Cells.Item("STYLE_COLOR_STATUS").Text
                            Case Is = "A"
                                pBox.Visible = True
                                pBox.ImageLocation = active
                                'pBox.Bottom
                            Case Is = "N"
                                pBox.Visible = True
                                pBox.ImageLocation = noReorder
                            Case Is = "D"
                                pBox.Visible = True
                                pBox.ImageLocation = discontinued
                            Case Else
                                pBox.Visible = False
                                pBox.ImageLocation = ""
                        End Select
                    End If

                End If
            Next
            For i As Integer = curRow + 1 To 19
                Dim pBox As PictureBox = Absx1.CtlFor("PictureBox" & i.ToString())
                pBox.Visible = False
                pBox.ImageLocation = ""
            Next i
        End If
    End Sub

    Private Sub numCARTONS_PER_UNIT_ValueChanged(sender As Object, e As EventArgs) Handles numCARTONS_PER_UNIT.ValueChanged
        If IsNumeric(numCARTONS_PER_UNIT.Value) Then
            If Val(numCARTONS_PER_UNIT.Value) > 0 Then
                numCARTONS_PER_UNIT.Appearance.BackColor = Drawing.Color.OrangeRed
            Else
                numCARTONS_PER_UNIT.Appearance.BackColor = Drawing.Color.Empty
            End If
        Else
            numCARTONS_PER_UNIT.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

#Region "Promo System"
    Private Sub btnShowPromo_Click(sender As Object, e As EventArgs) Handles btnShowPromo.Click
        Dim F As New ASFMSGBF
        F.grdGroupBy = True
        F.grdFilter = True
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("P1.PROMO_DESC As Promotion,")
        sql.AppendLine("P1.PROMO_START_DATE As Beginning,")
        sql.AppendLine("P1.PROMO_END_DATE As Ending,")
        sql.AppendLine("P2.STYLE_CODE As Style,")
        sql.AppendLine("S1.STYLE_DESC As Description,")
        sql.AppendLine("MAX(P2.PROMO_UNIT_PRICE) As Price")
        sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2, ICTSTYL1 S1")
        sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
        sql.AppendLine("AND P2.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND (P1.PROMO_START_DATE <= SYSDATE AND P1.PROMO_END_DATE >= SYSDATE)")
        sql.AppendLine("GROUP BY")
        sql.AppendLine("P1.PROMO_DESC,")
        sql.AppendLine("P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE,")
        sql.AppendLine("P2.STYLE_CODE,")
        sql.AppendLine("S1.STYLE_DESC")
        Dim tblICTPROMX As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        If tblICTPROMX.Rows.Count > 0 Then
            F.Show_grd(tblICTPROMX, Me, "Current Active Promotions", "")
            F.Dispose()
            F = Nothing
        End If
    End Sub

    Private Sub ShowPromo(ByVal STYLE_CODE As String)
        Dim OnPromo As Boolean = False
        Dim PROMO_START_DATE As DateTime
        Dim PROMO_END_DATE As DateTime
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE,")
        sql.AppendLine("MAX(P2.PROMO_UNIT_PRICE) PROMO_UNIT_PRICE")
        sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2")
        sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
        sql.AppendLine("AND P2.STYLE_CODE = :PARM1")
        sql.AppendLine("GROUP BY P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE")
        Dim tblICTPROMX As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
        For Each rowICTPROMX As DataRow In tblICTPROMX.Select("", "PROMO_START_DATE")
            PROMO_START_DATE = CDate(rowICTPROMX.Item("PROMO_START_DATE").ToString & String.Empty)
            PROMO_END_DATE = CDate(rowICTPROMX.Item("PROMO_END_DATE").ToString & String.Empty)
            If PROMO_START_DATE <= Now() And PROMO_END_DATE >= Now() Then
                OnPromo = True
            End If
        Next
        If OnPromo Then
            lblPromo.Text = String.Format("Style On Promo {0} - {1}", PROMO_START_DATE.ToShortDateString, PROMO_END_DATE.ToShortDateString)
            lblPromo.Visible = True
            btnShowPromo.Visible = True
        Else
            lblPromo.Text = ""
            lblPromo.Visible = False
            btnShowPromo.Visible = False
        End If
    End Sub

#End Region

End Class