Imports System.Drawing
Imports System.Math
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Http.Formatting
Imports Newtonsoft.Json.Linq
Imports Infragistics.Win.UltraWinGrid

Public Class ICFXLSWR
    Dim rowICTXLSW1 As DataRow
    Dim APTVENDX As String
    Dim ICTCLASX As String
    Dim XLS_NO As String = ""
    Dim XLS_IMP_NO As String = ""
    Dim uploadMsgs As String = ""
    Dim apiEndpoint As String = "http://api2.regency-rib.com:8085/api/"
    Dim importFailed As Boolean = False
    Dim responseImported As Boolean = False
    Dim listPriceMaintenanceMode As Boolean = False
    Dim vendorDimensionsUpdateMode As Boolean = False
    Dim calcCodeMaintenanceMode As Boolean = False
    Dim importErrors As Boolean = False



#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFTHEMI" Then
            InquiryMode = True
        ElseIf MENU_ITEM_OBJECT = "ICFXLSWM" Then
            listPriceMaintenanceMode = True
        ElseIf MENU_ITEM_OBJECT = "ICFXLSWC" Then
            calcCodeMaintenanceMode = True
        ElseIf MENU_ITEM_OBJECT = "ICFXLSWD" Then
            vendorDimensionsUpdateMode = True
            Me.AllowDrop = True
        Else
            Me.AllowDrop = True
        End If

        With dst
            ASCMAIN1.sql = "Select VEND_CODE, '0' SEL from APTVEND1 where VEND_SUPPLIER_ID is Not Null and VEND_STATUS = 'A'"
            APTVENDX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & APTVENDX & " Add Primary Key (VEND_CODE)")

            ASCMAIN1.sql = "Select * from " & APTVENDX
            Create_TDA(.Tables.Add("APTVENDX"), APTVENDX, "**", 0)
            dst.Tables("APTVENDX").Columns("SEL").DefaultValue = "0"
            Fill_Records("APTVENDX")

            ASCMAIN1.sql = "Select STYLE_CLASS_CODE, '0' SEL from ICTCLAS1"
            ICTCLASX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTCLASX & " Add Primary Key (STYLE_CLASS_CODE)")

            ASCMAIN1.sql = "Select * from " & ICTCLASX
            Create_TDA(.Tables.Add("ICTCLASX"), ICTCLASX, "**", 0)
            dst.Tables("ICTCLASX").Columns("SEL").DefaultValue = "0"
            Fill_Records("ICTCLASX")

            Dim SLS_YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -13)

            Dim sqlUnits As String = "SELECT " _
                & "  SOTINVH2.STYLE_CODE, SUM(SOTINVH2.ORDR_QTY_SHIP) QTY_SHP" _
                & " FROM" _
                & "  SOTINVH1, SOTINVH2" _
                & " WHERE" _
                & "  SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & "  AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & "  AND SOTINVH1.INV_TYPE = 'I'" _
                & "  AND SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & SLS_YP & "'" _
                & " GROUP BY STYLE_CODE"

            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.VEND_CODE, ICTSTYL1.STYLE_CLASS_CODE, X.PO_DATE_ORDERED" & vbCrLf _
                & ", CASE WHEN ICTSTYV1.NEW_PO_COST <> 0 THEN ICTSTYV1.NEW_PO_COST ELSE ICTSTYV1.PO_COST END PO_COST" & vbCrLf _
                & ", CASE WHEN ICTSTYV1.NEW_PO_COST <> 0 THEN ICTSTYV1.NEW_PO_COST_DATE ELSE ICTSTYV1.PO_COST_DATE END PO_COST_DATE" & vbCrLf _
                & ", ICTSTYL1.LIST_CALC_CODE, ICTLSTC1.LIST_CALC_DESC, ICTSTYL1.STYLE_PRICE, SOTINVH2.QTY_SHP" & vbCrLf _
                & ", ICTSTYV1.VEND_REMARK" & vbCrLf _
                & " from ICTSTYL1,ICTLSTC1,ICTSTYV1," & APTVENDX & " APTVENDX, " & ICTCLASX & " ICTCLASX, (" & sqlUnits & ") SOTINVH2, (SELECT POTORDR2.STYLE_CODE, MAX(POTORDR1.PO_DATE_ORDERED) PO_DATE_ORDERED" & vbCrLf _
                & " FROM POTORDR1, POTORDR2" & vbCrLf _
                & " WHERE POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & " GROUP BY POTORDR2.STYLE_CODE) X" & vbCrLf _
                & " where ICTSTYL1.STYLE_STATUS = 'A'" & vbCrLf _
                & "   and ICTSTYL1.LIST_CALC_CODE = ICTLSTC1.LIST_CALC_CODE " & vbCrLf _
                & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and SOTINVH2.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYL1.INIT_DATE < :PARM1" & vbCrLf _
                & "   and APTVENDX.VEND_CODE = ICTSTYL1.VEND_CODE and NVL(APTVENDX.SEL,'0') = '1'" & vbCrLf _
                & "   and ICTCLASX.STYLE_CLASS_CODE = ICTSTYL1.STYLE_CLASS_CODE and NVL(ICTCLASX.SEL,'0') = '1'" & vbCrLf
            Create_TDA(.Tables.Add, "ICTSTYLX", "**", 0, False, "D")
            With dst.Tables("ICTSTYLX").Columns
                .Add("DISCONTINUE", GetType(System.String))
                .Add("EXCLUDE", GetType(System.String))
                .Add("LIST_CALC_CODE_NEW", GetType(System.String))
                .Add("NEW_PO_COST", GetType(System.Decimal))
                .Add("PCTCHG", GetType(System.Decimal), "IIF(ISNULL(PO_COST,0)=0,0,100 * (ISNULL(NEW_PO_COST,0) - ISNULL(PO_COST,0)) / ISNULL(PO_COST,0))")
                .Add("NEW_STYLE_PRICE", GetType(System.Decimal))
                .Add("PCTCHG_SP", GetType(System.Decimal), "IIF(ISNULL(STYLE_PRICE,0)=0,0,100 * (ISNULL(NEW_STYLE_PRICE,0) - ISNULL(STYLE_PRICE,0)) / ISNULL(STYLE_PRICE,0))")
                .Add("SLS_DIFF", GetType(System.Decimal), "IIF(ISNULL(QTY_SHP,0)=0,0, (NEW_STYLE_PRICE * QTY_SHP) - (STYLE_PRICE * QTY_SHP))")
            End With
            dst.Tables("ICTSTYLX").Columns("DISCONTINUE").DefaultValue = "0"
            dst.Tables("ICTSTYLX").Columns("EXCLUDE").DefaultValue = "0"

            Create_TDA(.Tables.Add, "ICTSTYL1", "*", 1, , , , "STYLE_STATUS, STYLE_PRICE,LIST_CALC_CODE")
            AUDIT.Add("ICTSTYL1", "*")

            Create_TDA(.Tables.Add, "ICTSTYV1", "*", 2, False)

            Dim XLS_STATUS As String = IIf(listPriceMaintenanceMode, "'L','R'", "'G','R'")
            If vendorDimensionsUpdateMode Then
                XLS_STATUS = "'z'"
            End If
            If calcCodeMaintenanceMode Then
                XLS_STATUS = "'C'"
            End If

            ASCMAIN1.sql = "Select * from ICTXLSW1 where XLS_STATUS IN (" & XLS_STATUS & ")"
            Create_TDA(.Tables.Add, "ICTXLSW1", "**", 0)

            Create_TDA(.Tables.Add, "ICTXLSW3", "*", 1)

            Dim DT As DataTable

            DT = .Tables("ICTXLSW3").Clone
            DT.TableName = "ICTXLSW3_V"
            .Tables.Add(DT)

            Create_Relation("ICTXLSW3", "ICTXLSW3_V", "STYLE_CODE")

            '               & "   and ICTSTYL1.STYLE_STATUS = 'A'" & vbCrLf _ removed by request from Simon
            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_MATL_DESC" & vbCrLf _
               & ", ICTSTYLD_CTN.LENGTH CTN_LENGTH, ICTSTYLD_CTN.WIDTH CTN_WIDTH, ICTSTYLD_CTN.HEIGHT CTN_HEIGHT, ICTSTYLD_CTN.WEIGHT CTN_WEIGHT" & vbCrLf _
               & ", ICTSTYLD_INR.LENGTH INR_LENGTH, ICTSTYLD_INR.WIDTH INR_WIDTH, ICTSTYLD_INR.HEIGHT INR_HEIGHT, ICTSTYLD_INR.WEIGHT INR_WEIGHT" & vbCrLf _
               & ", ICTSTYLD_ITM.LENGTH ITM_LENGTH, ICTSTYLD_ITM.WIDTH ITM_WIDTH, ICTSTYLD_ITM.HEIGHT ITM_HEIGHT, ICTSTYLD_ITM.WEIGHT ITM_WEIGHT" & vbCrLf _
               & " from ICTSTYL1,ICTSTYLD ICTSTYLD_CTN,ICTSTYLD ICTSTYLD_INR,ICTSTYLD ICTSTYLD_ITM" & vbCrLf _
               & " where ICTSTYLD_CTN.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
               & "   and ICTSTYLD_CTN.PACK_CODE (+) = 'CTN'" & vbCrLf _
               & "   and ICTSTYLD_INR.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
               & "   and ICTSTYLD_INR.PACK_CODE (+) = 'INR'" & vbCrLf _
               & "   and ICTSTYLD_ITM.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
               & "   and ICTSTYLD_ITM.PACK_CODE (+) = 'ITM'" & vbCrLf _
               & "   and ICTSTYL1.STYLE_CODE in (Select STYLE_CODE from ICTXLSW3,ICTXLSW1" & vbCrLf _
               & " where ICTXLSW3.XLS_IMP_NO = ICTXLSW1.XLS_IMP_NO and ICTXLSW1.XLS_NO = :PARM1)"
            Create_TDA(.Tables.Add, "ICTXLSWD", "**", 0, False, "V")

            DT = .Tables("ICTXLSWD").Clone
            DT.TableName = "ICTXLSWD_V"
            .Tables.Add(DT)

            Create_Relation("ICTXLSWD", "ICTXLSWD_V", "STYLE_CODE")

        End With

        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")
        grdICTXLSW1.DataSource = dst.Tables("ICTXLSW1")
        grdAPTVENDX.DataSource = dst.Tables("APTVENDX")
        grdICTCLASX.DataSource = dst.Tables("ICTCLASX")

        grdICTXLSW3.DataSource = dst.Tables("ICTXLSW3")
        grdICTXLSWD.DataSource = dst.Tables("ICTXLSWD")

        With grdICTSTYLX.DisplayLayout.Bands(0)
            .Columns("DISCONTINUE").Hidden = listPriceMaintenanceMode Or calcCodeMaintenanceMode
            .Columns("NEW_PO_COST").Hidden = listPriceMaintenanceMode Or calcCodeMaintenanceMode
            .Columns("PCTCHG").Hidden = listPriceMaintenanceMode Or calcCodeMaintenanceMode
            .Columns("NEW_STYLE_PRICE").Hidden = Not listPriceMaintenanceMode
            .Columns("EXCLUDE").Hidden = Not listPriceMaintenanceMode
            .Columns("QTY_SHP").Hidden = Not listPriceMaintenanceMode
            .Columns("SLS_DIFF").Hidden = Not listPriceMaintenanceMode
            .Columns("LIST_CALC_CODE").Hidden = Not calcCodeMaintenanceMode
            .Columns("LIST_CALC_CODE_NEW").Hidden = Not calcCodeMaintenanceMode
            .Columns("LIST_CALC_DESC").Hidden = Not calcCodeMaintenanceMode
        End With

        For i As Integer = 0 To 1
            With grdICTXLSW3.DisplayLayout.Bands(i)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Hidden = True
                    If New String() {"STYLE_CODE", "STYLE_DESC", "VEND_ITEM_CODE", "VEND_REMARK", "STYLE_SO_QTY_MIN", "INNER_PACK_QTY", "CARTON_PACK_QTY", "CASE_CUBE", "PO_COST", "STYLE_PO_QTY_MIN"}.Contains(gcol.Key) Then
                        gcol.Hidden = False
                        If i = 1 Then
                            gcol.Header.VisiblePosition = grdICTXLSW3.DisplayLayout.Bands(0).Columns(gcol.Key).Header.VisiblePosition
                        End If
                    End If
                Next
            End With
        Next

        Create_Summary(grdICTSTYLX, "STYLE_CODE", "Count")
        Create_Summary(grdICTXLSW1, "XLS_NO", "Count")

        If calcCodeMaintenanceMode Then
            ASCMAIN1.Add_Value_List(grdICTSTYLX, "LIST_CALC_CODE_NEW", "SELECT LIST_CALC_CODE T_CODE, LIST_CALC_DESC T_DESC from ICTLSTC1 where LIST_CALC_STATUS = 'A'")
            ASCMAIN1.Add_Value_List(cbeLIST_CALC_CODE, "LIST_CALC_CODE",,, "SELECT LIST_CALC_CODE T_CODE, LIST_CALC_DESC T_DESC from ICTLSTC1 where LIST_CALC_STATUS = 'A'")
            With grdICTSTYLX.DisplayLayout
                .Override.AllowUpdate = DefaultableBoolean.True
                With .Bands(0)
                    .Columns("LIST_CALC_DESC").Header.Caption = "LCC Desc"
                    .Columns("LIST_CALC_DESC").Header.VisiblePosition = .Columns("LIST_CALC_CODE").Header.VisiblePosition + 1
                    .Columns("LIST_CALC_CODE").Hidden = True
                    For Each c As UltraWinGrid.UltraGridColumn In .Columns
                        If c.Key = "LIST_CALC_CODE_NEW" Then
                            .Columns(c.Key).CellActivation = Activation.AllowEdit
                            .Columns(c.Key).Header.Caption = "LCC New"
                        Else
                            .Columns(c.Key).CellActivation = Activation.NoEdit
                        End If
                    Next
                End With
            End With
        Else
            grdICTSTYLX.DisplayLayout.Bands(0).Columns("LIST_CALC_CODE_NEW").Hidden = True
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Generate"

                If dst.Tables("APTVENDX").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select 1 or more Vendors"
                End If
                If dst.Tables("ICTCLASX").Select("SEL = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select 1 or more Classes"
                End If


            Case "Import Vendor Reply"
                If Absx1.txtFor("XLS_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Vendor ReQuote Request XLS No to Edit"
                Else
                    rowICTXLSW1 = LookUp("ICTXLSW1", Absx1.txtFor("XLS_NO").Text)
                    If rowICTXLSW1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Vendor ReQuote Request XLS No " & Absx1.txtFor("XLS_NO").Text
                    Else
                        If Not (listPriceMaintenanceMode Or calcCodeMaintenanceMode) Then
                            If rowICTXLSW1.Item("XLS_STATUS") & "" <> "G" And rowICTXLSW1.Item("XLS_STATUS") & "" <> "R" Then
                                EMsg &= vbCr & "XLS No " & Absx1.txtFor("XLS_NO").Text & " is Not eligible to Import Vendor Reply"
                            End If
                        End If

                    End If
                End If

            Case "Update"

                If grdICTSTYLX.Rows.Count = 0 And Not vendorDimensionsUpdateMode Then
                    EMsg &= vbCr & "No Styles to Update"

                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
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

            Case "Generate"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Import Vendor Reply"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                If Not importErrors Then
                    Mode_Settings(False)
                End If

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Refresh"
                Refresh_Documents()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Generate").Settings.Enabled = not_iScreenMode
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                    .Items("Import Vendor Reply").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Cancel").Visible = True
                    .Items("Done").Visible = False
                End With
                .Groups("Generate Style List").Visible = Not ScreenMode
                .Groups("Re-Quote Options").Visible = ScreenMode And (EntryMode = "N" And Not listPriceMaintenanceMode)
                .Groups("LCC Update").Visible = ScreenMode And calcCodeMaintenanceMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        spl.Panel1Collapsed = Not ScreenMode

        If EntryMode = "N" Then
            Set_Read_Only(UltraGroupBox1, False)
            grpHeader2.Visible = False
            grdICTSTYLX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
        Else
            Set_Read_Only(UltraGroupBox1, False)
            grpHeader2.Visible = True
            grdICTSTYLX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        End If

        'abStyles.Tabs("Styles").Visible = (EntryMode = "E" And Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode))
        tabStyles.Tabs("Styles").Visible = (EntryMode = "E" And Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode Or calcCodeMaintenanceMode))
        tabStyles.Tabs("ReQuote").Visible = ((EntryMode = "N" And Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode)) Or calcCodeMaintenanceMode)
        tabStyles.Tabs("Dimensions").Visible = (EntryMode = "E" And Not (listPriceMaintenanceMode Or calcCodeMaintenanceMode))
        dteCostEffectiveDate.Visible = (EntryMode = "E" And Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode Or calcCodeMaintenanceMode))

        If listPriceMaintenanceMode Or vendorDimensionsUpdateMode Or calcCodeMaintenanceMode Then
            dteCostEffectiveDate.Visible = False
            dteREPLY_BY_DATE.Visible = False
            UltraLabel5.Visible = False
            UltraLabel6.Visible = False
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("Generate").Visible = False
                    .Items("Import Vendor Reply").Visible = False
                End With
                '.Groups("Generate Style List").Visible = False
                .Groups("Re-Quote Options").Visible = False
                .Groups("Generate Style List").Visible = False

            End With
        End If

        If calcCodeMaintenanceMode Then
            With UltraExplorerBar1
                .Groups("Generate Style List").Visible = Not ScreenMode
            End With
        End If

        With grdICTSTYLX.DisplayLayout.Bands(0)
            .Columns("DISCONTINUE").Hidden = calcCodeMaintenanceMode Or listPriceMaintenanceMode
            .Columns("NEW_PO_COST").Hidden = (calcCodeMaintenanceMode Or listPriceMaintenanceMode Or EntryMode = "N")
            .Columns("PCTCHG").Hidden = (calcCodeMaintenanceMode Or listPriceMaintenanceMode Or EntryMode = "N")
            .Columns("NEW_STYLE_PRICE").Hidden = (EntryMode = "N" Or calcCodeMaintenanceMode)
            .Columns("PCTCHG_SP").Hidden = (EntryMode = "N" Or calcCodeMaintenanceMode)
            .Columns("EXCLUDE").Hidden = calcCodeMaintenanceMode Or (Not listPriceMaintenanceMode)
            .Columns("QTY_SHP").Hidden = calcCodeMaintenanceMode Or (Not listPriceMaintenanceMode)
            .Columns("SLS_DIFF").Hidden = calcCodeMaintenanceMode Or (Not listPriceMaintenanceMode)
            .Columns("LIST_CALC_CODE").Hidden = Not calcCodeMaintenanceMode
            .Columns("LIST_CALC_CODE_NEW").Hidden = Not calcCodeMaintenanceMode
        End With

        grdICTXLSW1.Visible = Not ScreenMode
        tabStyles.Visible = ScreenMode
        If tf And (EntryMode = "E") Then
            'dteREPLY_BY_DATE.Visible = False
        End If
        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTSTYLX", "ICTXLSW3", "ICTXLSW3_V", "ICTXLSWD", "ICTXLSWD_V", "ICTSTYL1", "ICTSTYV1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        For Each row As DataRow In dst.Tables("APTVENDX").Select("")
            row.Item("SEL") = "0"
        Next
        Sort_grdColumns(grdAPTVENDX, "VEND_CODE")

        For Each row As DataRow In dst.Tables("ICTCLASX").Select("")
            row.Item("SEL") = "0"
        Next
        Sort_grdColumns(grdICTCLASX, "STYLE_CLASS_CODE")

        Absx1.txtFor("VEND_CODE").Text = ""
        Absx1.txtFor("XLS_NO").Text = ""
        XLS_NO = ""
        XLS_IMP_NO = ""
        uploadMsgs = ""

        dteREPLY_BY_DATE.Value = Now.Date.AddDays(45)
        dteBefore.Value = Now.Date.AddDays(-90)
        chkGenerateEmail.Checked = True
        importFailed = False
        responseImported = False
        importErrors = False
        SplitContainer2.Panel2Collapsed = True
        Refresh_Documents()
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        Dim SLS_YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -16)

        If EntryMode = "N" Then
            Update_Record_TDA("APTVENDX")
            Update_Record_TDA("ICTCLASX")


            Fill_Records("ICTSTYLX", New Object() {dteBefore.Value})

        Else ' Import Vendor Reply

            XLS_NO = Absx1.txtFor("XLS_NO").Text

            Dim sqlUnits As String = "SELECT " _
                & "  SOTINVH2.STYLE_CODE, SUM(SOTINVH2.ORDR_QTY_SHIP) QTY_SHP" _
                & " FROM" _
                & "  SOTINVH1, SOTINVH2" _
                & " WHERE" _
                & "  SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & "  AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & "  AND SOTINVH1.INV_TYPE = 'I'" _
                & "  AND SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & SLS_YP & "'" _
                & " GROUP BY STYLE_CODE"

            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.VEND_CODE, ICTSTYL1.STYLE_CLASS_CODE, X.PO_DATE_ORDERED" & vbCrLf _
                & ", CASE WHEN ICTSTYV1.NEW_PO_COST <> 0 THEN ICTSTYV1.NEW_PO_COST ELSE ICTSTYV1.PO_COST END PO_COST" & vbCrLf _
                & ", CASE WHEN ICTSTYV1.NEW_PO_COST <> 0 THEN ICTSTYV1.NEW_PO_COST_DATE ELSE ICTSTYV1.PO_COST_DATE END PO_COST_DATE" & vbCrLf _
                & ", ICTSTYL1.LIST_CALC_CODE, ICTLSTC1.LIST_CALC_DESC, ICTSTYL1.STYLE_PRICE, SOTINVH2.QTY_SHP" & vbCrLf _
                & ", ICTSTYV1.VEND_REMARK" & vbCrLf _
                & " from ICTSTYL1,ICTLSTC1, ICTSTYV1," & APTVENDX & " APTVENDX, " & ICTCLASX & " ICTCLASX, ICTXLSW1, ICTXLSW3, (" & sqlUnits & ") SOTINVH2, (SELECT POTORDR2.STYLE_CODE, MAX(POTORDR1.PO_DATE_ORDERED) PO_DATE_ORDERED" & vbCrLf _
                & " FROM POTORDR1, POTORDR2" & vbCrLf _
                & " WHERE POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & " GROUP BY POTORDR2.STYLE_CODE) X" & vbCrLf _
                & " where ICTSTYL1.STYLE_STATUS = 'A'" & vbCrLf _
                & "   and ICTSTYL1.LIST_CALC_CODE = ICTLSTC1.LIST_CALC_CODE" & vbCrLf _
                & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE" & vbCrLf _
                & "   and ICTXLSW1.XLS_IMP_NO = ICTXLSW3.XLS_IMP_NO" & vbCrLf _
                & "   and ICTXLSW3.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and APTVENDX.VEND_CODE = ICTSTYL1.VEND_CODE " & vbCrLf _
                & "   and X.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and SOTINVH2.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and ICTCLASX.STYLE_CLASS_CODE = ICTSTYL1.STYLE_CLASS_CODE" & vbCrLf _
                & "   and ICTXLSW1.XLS_NO = '" & XLS_NO & "'"

            Fill_Records("ICTSTYLX", , , ASCMAIN1.sql)
            Fill_Records("ICTXLSW3", New Object() {XLS_IMP_NO})
            Fill_Records("ICTXLSWD", New Object() {XLS_NO})

            If listPriceMaintenanceMode Then
                For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select()
                    Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE")
                    Dim VEND_CODE As String = rowICTSTYLX.Item("VEND_CODE")
                    ASCMAIN1.sql = "Select * from ICTSTYV1 where STYLE_CODE = :PARM1 and VEND_CODE = :PARM2"
                    Dim rowICTSTYV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {STYLE_CODE, VEND_CODE})

                    Dim LIST_CALC_CODE As String = rowICTSTYLX.Item("LIST_CALC_CODE") & ""
                    If LIST_CALC_CODE <> "" Then
                        Dim SILENT As Boolean = True
                        ASCMAIN1.Progress("Calculating New List Price for: " & STYLE_CODE)
                        Dim NEW_STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Style_Price(Me, SILENT, STYLE_CODE, , rowICTSTYV1)
                        rowICTSTYLX.Item("NEW_STYLE_PRICE") = NEW_STYLE_PRICE
                    End If

                Next
                ASCMAIN1.Progress("", "")
            Else
                With grdICTSTYLX.DisplayLayout.Bands(0)
                    .Columns("NEW_STYLE_PRICE").Hidden = (rowICTXLSW1.Item("XLS_STATUS") & "" = "R")
                    .Columns("PCTCHG_SP").Hidden = (rowICTXLSW1.Item("XLS_STATUS") & "" = "R")
                    .Columns("SLS_DIFF").Hidden = (rowICTXLSW1.Item("XLS_STATUS") & "" = "R")
                End With
                If calcCodeMaintenanceMode Then
                    With grdICTSTYLX.DisplayLayout
                        .Override.AllowUpdate = DefaultableBoolean.True
                        With .Bands(0)
                            .Columns("NEW_STYLE_PRICE").Hidden = True
                            .Columns("PCTCHG_SP").Hidden = True
                            .Columns("PCTCHG").Hidden = True
                            .Columns("NEW_PO_COST").Hidden = True
                            .Columns("SLS_DIFF").Hidden = True
                            .Columns("LIST_CALC_CODE").Hidden = False
                            .Columns("LIST_CALC_CODE_NEW").Hidden = False
                            For Each c As UltraWinGrid.UltraGridColumn In .Columns
                                If c.Key = "LIST_CALC_CODE_NEW" Then
                                    .Columns(c.Key).CellActivation = Activation.AllowEdit
                                    .Columns(c.Key).Header.Caption = "LCC New"
                                Else
                                    .Columns(c.Key).CellActivation = Activation.NoEdit
                                End If
                            Next
                        End With
                    End With
                End If
            End If
        End If

        With grdICTSTYLX.DisplayLayout.Bands(0).SortedColumns
            .Clear()
            .Add("VEND_CODE", False, True)
            .Add("STYLE_CODE", False)
        End With
        If vendorDimensionsUpdateMode Then
            grdICTXLSWD.Rows.CollapseAll(True)
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        If vendorDimensionsUpdateMode Then
            BeginTrans()
            uploadMsgs = ""
            Upload_Dimensions()
            Manual_Style_Dimensions_Update()
            CommitTrans("")
            MsgBox(uploadMsgs, vbOKOnly, "Update Complete.")
            Exit Sub
        End If

        BeginTrans()

        Dim vendorEmails As New Dictionary(Of String, String)
        'If calcCodeMaintenanceMode Then
        '    EntryMode = "E"
        'End If
        If EntryMode = "N" Then

            Dim STYLE_CLASS_CODEs As String = ""
            For Each row As DataRow In dst.Tables("ICTCLASX").Select("SEL='1'", "STYLE_CLASS_CODE")
                Dim STYLE_CLASS_CODE As String = row.Item("STYLE_CLASS_CODE")
                STYLE_CLASS_CODEs &= "," & STYLE_CLASS_CODE
            Next

            If Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode) Then
                For Each row As DataRow In dst.Tables("APTVENDX").Select("SEL='1'", "VEND_CODE")
                    Dim VEND_CODE As String = row.Item("VEND_CODE")
                    Dim lccSql As String = ""
                    If calcCodeMaintenanceMode Then
                        'lccSql = " AND ISNULL(LIST_CALC_CODE_NEW,'') <> '' "
                    End If
                    Dim rowICTSTYLXs() As DataRow = dst.Tables("ICTSTYLX").Select("VEND_CODE = '" & VEND_CODE & "' AND ISNULL(DISCONTINUE,'0') = '0'" & lccSql, "STYLE_CODE")
                    If rowICTSTYLXs.Length > 0 Then
                        Dim rowICTXLSW1 As DataRow = dst.Tables("ICTXLSW1").NewRow
                        Dim XLS_NO As String = ASCMAIN1.Next_Control_No("ICTXLSW1.XLS_NO")
                        Dim XLS_IMP_NO As String = ASCMAIN1.Next_Control_No("ICTXLSW2.XLS_IMP_NO")

                        vendorEmails.Add(VEND_CODE, XLS_NO)

                        ASCMAIN1.sql = "Select Max (XLS_SEQ_NO) from ICTXLSW1 where VEND_CODE = :PARM1 and XLS_TYPE = 'R'"
                        Dim XLS_SEQ_NO As Integer = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {VEND_CODE})) + 1

                        With rowICTXLSW1
                            .Item("XLS_NO") = XLS_NO
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("VEND_CODE") = VEND_CODE
                            .Item("XLS_STATUS") = IIf(calcCodeMaintenanceMode, "C", "G") ' G = GENERATED, R = REPLIED, D = DELETED
                            .Item("XLS_DESC") = Mid(STYLE_CLASS_CODEs, 2)
                            .Item("REPLY_BY_DATE") = dteREPLY_BY_DATE.Value
                            .Item("XLS_TYPE") = "R"
                            .Item("XLS_IMP_NO") = XLS_IMP_NO
                            .Item("XLS_SEQ_NO") = XLS_SEQ_NO
                        End With
                        dst.Tables("ICTXLSW1").Rows.Add(rowICTXLSW1)

                        Dim XLS_IMP_LNO As Integer = 0
                        For Each rowICTSTYLX As DataRow In rowICTSTYLXs

                            Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE")
                            Dim rowICTSTY1 As DataRow = Fill_Record("ICTSTYL1", STYLE_CODE, , False)

                            Dim rowICTXLSW3 As DataRow = dst.Tables("ICTXLSW3").NewRow
                            With rowICTXLSW3
                                .Item("XLS_IMP_NO") = XLS_IMP_NO
                                XLS_IMP_LNO += 1
                                .Item("XLS_IMP_LNO") = XLS_IMP_LNO
                                .Item("STYLE_CODE") = STYLE_CODE
                                For Each col As DataColumn In rowICTSTY1.Table.Columns
                                    If .Table.Columns.Contains(col.ColumnName) Then
                                        .Item(col.ColumnName) = rowICTSTY1.Item(col.ColumnName)
                                    End If
                                Next

                            End With
                            dst.Tables("ICTXLSW3").Rows.Add(rowICTXLSW3)

                            If calcCodeMaintenanceMode Then
                                Dim SILENT As Boolean = True
                                Dim record As Integer = 0

                                Dim LIST_CALC_CODE As String = rowICTSTYLX.Item("LIST_CALC_CODE") & ""
                                Dim LIST_CALC_CODE_NEW As String = rowICTSTYLX.Item("LIST_CALC_CODE_NEW") & ""
                                If LIST_CALC_CODE <> LIST_CALC_CODE_NEW And LIST_CALC_CODE_NEW <> "" Then
                                    rowICTSTY1.Item("LIST_CALC_CODE") = LIST_CALC_CODE_NEW
                                    ASCMAIN1.Progress("Get ICTSTYV1 for " & STYLE_CODE & ":" & VEND_CODE, record.ToString)
                                    ASCMAIN1.sql = "Select * from ICTSTYV1 where STYLE_CODE = :PARM1 and VEND_CODE = :PARM2"
                                    Dim rowICTSTYV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {STYLE_CODE, VEND_CODE})
                                    ASCMAIN1.Progress("Calculate new price for " & STYLE_CODE & ":" & VEND_CODE, record.ToString)
                                    Dim NEW_STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Style_Price(Me, SILENT, STYLE_CODE, rowICTSTY1, rowICTSTYV1)
                                    rowICTSTY1.Item("STYLE_PRICE") = NEW_STYLE_PRICE
                                End If
                            End If
                        Next
                    End If
                Next

                For Each row As DataRow In dst.Tables("ICTSTYLX").Select("DISCONTINUE = '1'")
                    Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                    Dim rowICTSTY1 As DataRow = Fill_Record("ICTSTYL1", STYLE_CODE, , False)
                    rowICTSTY1.Item("STYLE_STATUS") = "D"
                Next
            Else
                'LIST PRICE MAINTENANCE MODE
                Dim rowICTSTYLXs() As DataRow = dst.Tables("ICTSTYLX").Select("ISNULL(EXCLUDE,'0') = '0'", "STYLE_CODE")
                If rowICTSTYLXs.Length > 0 Then

                    Dim rowICTXLSW1 As DataRow = dst.Tables("ICTXLSW1").NewRow
                    Dim XLS_NO As String = ASCMAIN1.Next_Control_No("ICTXLSW1.XLS_NO")
                    Dim XLS_IMP_NO As String = ASCMAIN1.Next_Control_No("ICTXLSW2.XLS_IMP_NO")

                    With rowICTXLSW1
                        .Item("XLS_NO") = XLS_NO
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("XLS_STATUS") = "L" ' G = GENERATED, R = REPLIED, D = DELETED
                        .Item("XLS_DESC") = Mid(STYLE_CLASS_CODEs, 2)
                        .Item("REPLY_BY_DATE") = dteREPLY_BY_DATE.Value
                        .Item("XLS_TYPE") = "R"
                        .Item("XLS_IMP_NO") = XLS_IMP_NO
                    End With
                    dst.Tables("ICTXLSW1").Rows.Add(rowICTXLSW1)

                    Dim XLS_IMP_LNO As Integer = 0
                    For Each rowICTSTYLX As DataRow In rowICTSTYLXs
                        Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE")
                        Dim rowICTSTY1 As DataRow = Fill_Record("ICTSTYL1", STYLE_CODE, , False)
                        Dim rowICTXLSW3 As DataRow = dst.Tables("ICTXLSW3").NewRow
                        With rowICTXLSW3
                            .Item("XLS_IMP_NO") = XLS_IMP_NO
                            XLS_IMP_LNO += 1
                            .Item("XLS_IMP_LNO") = XLS_IMP_LNO
                            .Item("STYLE_CODE") = STYLE_CODE
                            For Each col As DataColumn In rowICTSTY1.Table.Columns
                                If .Table.Columns.Contains(col.ColumnName) Then
                                    .Item(col.ColumnName) = rowICTSTY1.Item(col.ColumnName)
                                End If
                            Next

                        End With
                        dst.Tables("ICTXLSW3").Rows.Add(rowICTXLSW3)
                    Next

                End If

            End If

            Update_Record_TDA("ICTSTYL1")
            Update_Record_TDA("ICTXLSW1")
            Update_Record_TDA("ICTXLSW3")
        Else

            If Not responseImported And Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode Or calcCodeMaintenanceMode) Then
                ASCMAIN1.sql = "Update ICTXLSW1 Set REPLY_BY_DATE = :PARM1 where XLS_NO = :PARM2"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DV", New Object() {dteREPLY_BY_DATE.Value, XLS_NO})
            Else
                Dim XLS_STATUS As String = "R"
                If calcCodeMaintenanceMode Then
                    XLS_STATUS = "C"
                End If
                ASCMAIN1.sql = "Update ICTXLSW1 Set XLS_STATUS = :PARM1 where XLS_NO = :PARM2"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {XLS_STATUS, XLS_NO})

                If Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode Or calcCodeMaintenanceMode) Then
                    For Each row As DataRow In dst.Tables("ICTSTYLX").Select("DISCONTINUE = '1'")
                        Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                        Dim rowICTSTY1 As DataRow = Fill_Record("ICTSTYL1", STYLE_CODE, , False)
                        rowICTSTY1.Item("STYLE_STATUS") = "D"
                    Next
                    ' STASHES VENDOR'S REPLY XLS
                Else
                    Dim rowICTSTYLXs() As DataRow = dst.Tables("ICTSTYLX").Select("ISNULL(EXCLUDE,'0') = '0'", "STYLE_CODE")
                    If rowICTSTYLXs.Length > 0 Then
                        For Each rowICTSTYLX As DataRow In rowICTSTYLXs
                            Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE")
                            Dim rowICTSTYL1 As DataRow = Fill_Record("ICTSTYL1", STYLE_CODE, , False)
                            If calcCodeMaintenanceMode Then
                                Dim SILENT As Boolean = True
                                Dim record As Integer = 0
                                Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text & ""

                                Dim LIST_CALC_CODE As String = rowICTSTYL1.Item("LIST_CALC_CODE") & ""
                                Dim LIST_CALC_CODE_NEW As String = rowICTSTYLX.Item("LIST_CALC_CODE_NEW") & ""
                                If LIST_CALC_CODE <> LIST_CALC_CODE_NEW And LIST_CALC_CODE_NEW <> "" Then
                                    rowICTSTYL1.Item("LIST_CALC_CODE") = LIST_CALC_CODE_NEW
                                    ASCMAIN1.Progress("Get ICTSTYV1 for " & STYLE_CODE & ":" & VEND_CODE, record.ToString)
                                    ASCMAIN1.sql = "Select * from ICTSTYV1 where STYLE_CODE = :PARM1 and VEND_CODE = :PARM2"
                                    Dim rowICTSTYV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {STYLE_CODE, VEND_CODE})
                                    ASCMAIN1.Progress("Calculate new price for " & STYLE_CODE & ":" & VEND_CODE, record.ToString)
                                    Dim NEW_STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Style_Price(Me, SILENT, STYLE_CODE, rowICTSTYL1, rowICTSTYV1)
                                    rowICTSTYL1.Item("STYLE_PRICE") = NEW_STYLE_PRICE
                                End If
                            Else
                                rowICTSTYL1.Item("STYLE_PRICE") = Val(rowICTSTYLX.Item("NEW_STYLE_PRICE") & "")
                            End If
                        Next
                    End If
                End If

                If dst.Tables("ICTSTYL1").Rows.Count > 0 Then
                    Update_Record_TDA("ICTSTYL1")
                End If
            End If


        End If

        If Not vendorDimensionsUpdateMode Then
            CommitTrans("Update Complete")
            ASCMAIN1.Progress("", "")
        End If

        If EntryMode = "N" Then
            If Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode Or calcCodeMaintenanceMode) Then
                If chkGenerateEmail.Checked Then
                    Dim e As Integer = 0
                    For Each kvp As KeyValuePair(Of String, String) In vendorEmails
                        Dim VEND_CODE As String = kvp.Key
                        Dim XLS_NO As String = kvp.Value
                        Generate_Vendor_Email(XLS_NO, VEND_CODE)
                        ASCMAIN1.Progress("Waiting for API to process reguest...", "")
                        System.Threading.Thread.Sleep(5000)
                        e += 1
                    Next
                    If e > 0 Then
                        ASCMAIN1.Progress("", "")
                        MsgBox("Generated " & e & " emails." & vbCrLf & "Please review the emails in your drafts folder and send.", vbOKOnly, "Draft(s) Generated")
                    End If
                End If
            End If
        Else
            If Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode Or calcCodeMaintenanceMode) Then
                If responseImported Then
                    uploadMsgs = ""
                    Upload_Styles()
                    Upload_Dimensions()
                    Archive_Vendor_Spreadsheet()
                    MsgBox(uploadMsgs, vbOKOnly, "Uploads Complete.")
                    With UltraExplorerBar1
                        With .Groups("Screen Control")
                            .Items("Cancel").Visible = False
                            .Items("Done").Visible = True
                        End With
                        .Groups("Generate Style List").Visible = Not ScreenMode
                        .Groups("Re-Quote Options").Visible = ScreenMode And (EntryMode = "N" And Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode))
                    End With
                End If

            End If

        End If

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTXLSW1, "SSBBB", "Show Filter", "Show GroupBox", "Generate Email(s)", "Delete Request")
        Load_Popup_Menu(grdICTSTYLX, "SSBBBB", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Style Master File", "Discontinue", "De-Discontinue", "Exclude from Update", "Remove Exclusions")
        Load_Popup_Menu(grdICTXLSW3, "SSBB", "Show Filter", "Show GroupBox", "Collapse All", "Expand All")
        Load_Popup_Menu(grdICTXLSWD, "SSBB", "Show Filter", "Show GroupBox", "Collapse All", "Expand All")
        Load_Popup_Menu(grdAPTVENDX, "BB", "Select All", "De-Select All")
        Load_Popup_Menu(grdICTCLASX, "BB", "Select All", "De-Select All")
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

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdICTSTYLX"
                tlb_btn = DirectCast(tlb_pop.Tools("Discontinue"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode)
                tlb_btn = DirectCast(tlb_pop.Tools("De-Discontinue"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not (listPriceMaintenanceMode Or vendorDimensionsUpdateMode)
                tlb_btn = DirectCast(tlb_pop.Tools("Exclude from Update"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (listPriceMaintenanceMode Or vendorDimensionsUpdateMode)
                tlb_btn = DirectCast(tlb_pop.Tools("Remove Exclusions"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (listPriceMaintenanceMode Or vendorDimensionsUpdateMode)
        End Select
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdICTSTYC1"

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

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Style Master File"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim keys As New Dictionary(Of String, Object)
                keys.Add("STYLE_CODE", grd.ActiveRow.Cells("STYLE_CODE").Text)
                Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "De-Select All", "0", "1")
                    grow.Update()
                Next
            Case "Generate Email(s)"
                Dim emailCount As Integer = 0
                If grd.Selected.Rows.Count = 1 Then
                    Dim XLS_NO As String = grd.ActiveRow.Cells("XLS_NO").Text
                    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Text
                    Generate_Vendor_Email(XLS_NO, VEND_CODE)
                    emailCount += 1
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        Dim XLS_NO As String = grow.Cells("XLS_NO").Text
                        Dim VEND_CODE As String = grow.Cells("VEND_CODE").Text
                        Generate_Vendor_Email(XLS_NO, VEND_CODE)
                        ASCMAIN1.Progress("Waiting for API to process reguest...", "")
                        System.Threading.Thread.Sleep(5000)
                        emailCount += 1
                    Next
                End If
                ASCMAIN1.Progress("", "")
                MsgBox("Generated " & emailCount & " email(s). Please verify the emails in Outlook.", vbOKOnly, "Generation Complete")
                grd.Selected.Rows.Clear()
                ASCMAIN1.Progress("", "")
            Case "Delete Request"
                Dim d As Integer = 0
                Dim dMsg As String = "Are you sure you want to Delete this Re-Quote Request?"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    d += 1
                Next
                If d > 1 Then
                    dMsg = "Are you sure you want to delete these " & d.ToString & " Re-Quote Requests?"
                End If
                If MessageBox.Show(dMsg, "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                BeginTrans()
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    Dim XLS_NO As String = grow.Cells("XLS_NO").Text
                    ASCMAIN1.sql = "Update ICTXLSW1 Set XLS_STATUS = :PARM1 where XLS_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"D", XLS_NO})
                Next
                CommitTrans("Request" & IIf(d > 1, "s", "") & " Deleted")
                Refresh_Documents()
            Case "Expand All"
                grd.Rows.ExpandAll(True)
            Case "Collapse All"
                grd.Rows.CollapseAll(True)
            Case "Discontinue", "De-Discontinue"
                Dim DISCONTINUE As String = IIf(e.Tool.Key = "Discontinue", "1", "0")
                If grd.Selected.Rows.Count <= 1 Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                    For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("STYLE_CODE = '" & STYLE_CODE & "'")
                        rowICTSTYLX.Item("DISCONTINUE") = DISCONTINUE
                    Next
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        grow.Cells("DISCONTINUE").Value = DISCONTINUE
                        grow.Update()
                    Next
                    grd.Selected.Rows.Clear()
                End If
            Case "Change Reply-By Date"


            Case "Exclude from Update", "Remove Exclusions"
                Dim EXCLUDE As String = IIf(e.Tool.Key = "Exclude from Update", "1", "0")
                If grd.Selected.Rows.Count <= 1 Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                    For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("STYLE_CODE = '" & STYLE_CODE & "'")
                        rowICTSTYLX.Item("EXCLUDE") = EXCLUDE
                    Next
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        grow.Cells("EXCLUDE").Value = EXCLUDE
                        grow.Update()
                    Next
                    grd.Selected.Rows.Clear()
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "XLS_NO"
                'If e.KeyCode = Windows.Forms.Keys.Enter Then
                '    Click_Command("Import Vendor Reply", e)
                'End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "XLS_NO"
            '    Click_Command("Import Vendor Reply")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "VEND_CODE"
        End Select
    End Sub

#End Region

    Sub Generate_Vendor_Email(XLS_NO As String, VEND_CODE As String)
        Dim iResult As String = ""
        Dim API_BASE As String = ""
        Dim apiMethod As String = "GetVendorInfoSheet"
        Dim API_CONTROLLER As String = "RGI/IC/" & apiMethod

        Dim client As New HttpClient()
        client.BaseAddress = New Uri(ASCMAIN1.Get_API_Endpoint(False))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))
        client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", ASCMAIN1.Get_User_JWT())

        Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()
        Dim XLS_REQ As New xlsRequest
        XLS_REQ.XLS_NO = XLS_NO
        XLS_REQ.REQUEST_TYPE = "R"
        Dim content As HttpContent = New ObjectContent(Of xlsRequest)(XLS_REQ, frmtr)

        Dim resp As HttpResponseMessage = Nothing
        Dim resp_err As String = ""
        Try
            ASCMAIN1.Progress("Now Generating Requote Spreadsheet...", XLS_NO)
            resp = client.PostAsync(API_CONTROLLER, content).Result
            If resp.StatusCode = Net.HttpStatusCode.OK Then
                Dim apiResponseString As String = ""
                Dim responseObject As Object = Nothing
                responseObject = resp.Content.ReadAsAsync(Of Object)().Result
                Dim xlsFilePath As String = responseObject("filePath").ToString
                Dim xlsFileName As String = responseObject("fileName").ToString
                Dim infoSheetPath_dev As String = xlsFilePath & "\" & xlsFileName
                Dim infoSheetPath As String = "\\IIS2019\spreadsheets\FactoryInfoSheet\" & xlsFileName
                If ASCMAIN1.Running_in_VS Then
                    infoSheetPath = infoSheetPath_dev
                End If
                ASCMAIN1.Progress("Now Generating Email...", VEND_CODE)
                rowICTXLSW1 = LookUp("ICTXLSW1", XLS_NO)
                Generate_Vendor_Email_Draft(infoSheetPath, VEND_CODE & " Requote Request: " & xlsFileName)
            Else
                MsgBox("Error: " & resp.StatusCode.ToString, vbOKOnly, "API Error")
            End If
        Catch ex As Exception
            resp_err = ex.InnerException.InnerException.Message
            MsgBox(resp_err, vbOKOnly, "Error Generating Spreadsheet")
        End Try

    End Sub

#Region "Drag/Drop Events"
    Private Sub Form1_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim dFiles() As String = e.Data.GetData(DataFormats.FileDrop)
        If vendorDimensionsUpdateMode Then
            EntryMode = "E"
            Mode_Settings(True)
        End If
        If EntryMode = "E" Then
            For Each df As Object In dFiles
                importFailed = False
                Import_Vendor_Reply(df.ToString)
                If Not importFailed Then
                    responseImported = True
                    If vendorDimensionsUpdateMode Then
                        grdICTXLSWD.Rows.CollapseAll(True)
                    End If
                    MsgBox(df.ToString & " Imported.", vbOKOnly, "Import Complete")
                Else
                    ASCMAIN1.Progress("", "")
                End If
            Next
        End If

    End Sub

    Private Sub Form1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
#End Region

    Private Sub grdICTXLSW1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTXLSW1.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("XLS_NO").Text = e.Row.Cells("XLS_NO").Text
            Absx1.txtFor("VEND_CODE").Text = e.Row.Cells("VEND_CODE").Text
            dteCostEffectiveDate.Value = Date.Now
            XLS_IMP_NO = e.Row.Cells("XLS_IMP_NO").Text
            Click_Command("Import Vendor Reply")
        End If
    End Sub

    Sub Refresh_Documents()
        Fill_Records("ICTXLSW1")
        Sort_grdColumns(grdICTXLSW1, "XLS_NO".ToLower)
    End Sub

    Sub Generate_Vendor_Email_Draft(FILENAME As String, MAIL_SUBJECT As String)

        Dim MAIL_BODY As String = ""

        MAIL_BODY = "Dear Vendor," & vbCrLf & vbCrLf

        MAIL_BODY &= "Enclosed please find the re-run item list for Style Data (cost, packing, etc.) and the Dimension & Weight sheet." & vbCrLf & vbCrLf

        MAIL_BODY &= "Please check Factory item#, packing, cost and MOQ of each item on the Style Data tab and the material breakdown, dimensions, and weight on the Dimensions & Weights tab." & vbCrLf & vbCrLf

        MAIL_BODY &= "Please review previous Material Component Breakdown percentage to make sure they are correct after your mass production. " & vbCrLf
        MAIL_BODY &= "Please review Style Dimensions & Weights to make sure they are correct after your mass production." & vbCrLf
        MAIL_BODY &= "Please do not rename the workbook." & vbCrLf & vbCrLf

        MAIL_BODY &= "We request to send this list back on/ before the deadline of " & rowICTXLSW1.Item("REPLY_BY_DATE")


        Create_Outlook_mailitem("", "", MAIL_SUBJECT, MAIL_BODY, New String() {FILENAME})

    End Sub

    Sub Import_Vendor_Reply(vendorWB As String)

        If vendorWB <> "" Then
            Dim eMsg As String = ""
            Dim VEND_CODE As String = ""
            ASCMAIN1.Progress("Now Importing Vendor Reply")

            Try
                Dim oWB As SpreadsheetGear.IWorkbook
                oWB = SpreadsheetGear.Factory.GetWorkbook(vendorWB)
                Dim ws As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)   'style data

                dst.Tables("ICTXLSW3_V").Rows.Clear()

                If Not vendorDimensionsUpdateMode Then
                    For r As Int64 = 2 To ws.UsedRange.RowCount - 1
                        Dim XLS_IMP_LNO As Integer = r - 1
                        Dim STYLE_CODE As String = ws.Cells(r, 0).Text

                        If STYLE_CODE = "" Then Exit For

                        Dim rowICTXLSW3_orig As DataRow = dst.Tables("ICTXLSW3").Rows.Find(New Object() {XLS_IMP_NO, XLS_IMP_LNO})
                        If rowICTXLSW3_orig Is Nothing Then
                            MsgBox("No Record of style: " & STYLE_CODE & " on line " & XLS_IMP_LNO.ToString, vbOKOnly, "Cannot Proceed")
                            importFailed = True
                            Exit Sub
                        End If

                        Dim STYLE_CODE_orig As String = rowICTXLSW3_orig.Item("STYLE_CODE") & ""
                        If STYLE_CODE <> STYLE_CODE_orig Then
                            MsgBox("Invalid Vendor Worksheet - Style Code Mismatch", vbOKOnly, "Cannot Proceed")
                            importFailed = True
                            Exit Sub
                        End If
                        VEND_CODE = Absx1.txtFor("VEND_CODE").Text
                        ASCMAIN1.sql = "Select * from ICTSTYV1 where STYLE_CODE = :PARM1 and VEND_CODE = :PARM2"
                        Dim rowICTSTYV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {STYLE_CODE, VEND_CODE})

                        Dim roqICTSTYLXs() As DataRow = dst.Tables("ICTSTYLX").Select("STYLE_CODE = '" & STYLE_CODE & "'", "")
                        If roqICTSTYLXs.Length > 0 Then
                            Dim NEW_PO_COST As String = Val(ws.Cells(r, 12).Text & "")
                            roqICTSTYLXs(0).Item("NEW_PO_COST") = NEW_PO_COST
                            roqICTSTYLXs(0).Item("VEND_REMARK") = ws.Cells(r, 5).Text
                            rowICTSTYV1.Item("NEW_PO_COST") = NEW_PO_COST
                            rowICTSTYV1.Item("NEW_PO_COST_DATE") = Now.Date
                            Dim LIST_CALC_CODE As String = roqICTSTYLXs(0).Item("LIST_CALC_CODE") & ""
                            If LIST_CALC_CODE <> "" Then
                                Dim SILENT As Boolean = True
                                ASCMAIN1.Progress("Calculating New List Price for: " & STYLE_CODE)
                                Dim NEW_STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Style_Price(Me, SILENT, STYLE_CODE, , rowICTSTYV1)
                                roqICTSTYLXs(0).Item("NEW_STYLE_PRICE") = NEW_STYLE_PRICE
                            End If
                        End If

                        Dim rowICTXLSW3_V As DataRow = dst.Tables("ICTXLSW3_V").NewRow
                        For Each col As DataColumn In rowICTXLSW3_orig.Table.Columns
                            Dim colName As String = col.ColumnName
                            rowICTXLSW3_V.Item(colName) = rowICTXLSW3_orig.Item(colName)
                        Next
                        'STYLE_SO_QTY_MIN

                        rowICTXLSW3_V.Item("VEND_ITEM_CODE") = ws.Cells(r, 3).Text
                        rowICTXLSW3_V.Item("VEND_REMARK") = ws.Cells(r, 5).Text
                        rowICTXLSW3_V.Item("STYLE_SO_QTY_MIN") = ws.Cells(r, 7).Text
                        rowICTXLSW3_V.Item("INNER_PACK_QTY") = ws.Cells(r, 8).Text
                        rowICTXLSW3_V.Item("CARTON_PACK_QTY") = ws.Cells(r, 9).Text
                        rowICTXLSW3_V.Item("CASE_CUBE") = ws.Cells(r, 10).Text
                        rowICTXLSW3_V.Item("PO_COST") = ws.Cells(r, 12).Text
                        rowICTXLSW3_V.Item("STYLE_PO_QTY_MIN") = ws.Cells(r, 13).Text

                        dst.Tables("ICTXLSW3_V").Rows.Add(rowICTXLSW3_V)

                    Next
                Else
                    VEND_CODE = Split(oWB.Name, "_")(0)
                    Absx1.txtFor("VEND_CODE").Text = VEND_CODE
                End If

                ASCMAIN1.Progress("Importing Style Dimension Data")
                dst.Tables("ICTXLSWD_V").Rows.Clear()
                ws = oWB.Worksheets(1) 'dimensions & weight
                For r As Int64 = 2 To ws.UsedRange.RowCount - 1
                    Dim STYLE_CODE As String = ws.Cells(r, 0).Text
                    If STYLE_CODE = "" Then Exit For
                    Dim rowICTXLSWD_V As DataRow = dst.Tables("ICTXLSWD_V").NewRow

                    If Not vendorDimensionsUpdateMode Then
                        Dim rowICTXLSWD_orig As DataRow = Nothing
                        Dim rowICTXLSWD_check() As DataRow = dst.Tables("ICTXLSWD").Select("STYLE_CODE = '" & STYLE_CODE & "'", "STYLE_CODE")
                        If rowICTXLSWD_check.Length > 0 Then
                            rowICTXLSWD_orig = rowICTXLSWD_check(0)
                        Else
                            MsgBox("No Record of style: " & STYLE_CODE & " on generated worksheet.", vbOKOnly, "Cannot Proceed")
                            importFailed = True
                            Exit Sub
                        End If

                        Dim STYLE_CODE_orig As String = rowICTXLSWD_orig.Item("STYLE_CODE") & ""
                        If STYLE_CODE <> STYLE_CODE_orig Then
                            MsgBox("Invalid Vendor Worksheet - Style Code Mismatch", vbOKOnly, "Cannot Proceed")
                            importFailed = True
                            Exit Sub
                        End If

                        For Each col As DataColumn In rowICTXLSWD_orig.Table.Columns
                            Dim colName As String = col.ColumnName
                            rowICTXLSWD_V.Item(colName) = rowICTXLSWD_orig.Item(colName)
                        Next
                    Else
                        ASCMAIN1.sql = "Select * from ICTSTYV1 where STYLE_CODE = :PARM1 and VEND_CODE = :PARM2"
                        Dim rowICTSTYV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {STYLE_CODE, VEND_CODE})

                        Dim STYLE_DESC As String = ws.Cells(r, 1).Text
                        rowICTXLSWD_V.Item("STYLE_CODE") = STYLE_CODE
                        rowICTXLSWD_V.Item("STYLE_DESC") = STYLE_DESC
                    End If


                    rowICTXLSWD_V.Item("STYLE_MATL_DESC") = ws.Cells(r, 2).Text

                    rowICTXLSWD_V.Item("CTN_LENGTH") = Val(ws.Cells(r, 3).Text & "")
                    rowICTXLSWD_V.Item("CTN_WIDTH") = Val(ws.Cells(r, 4).Text & "")
                    rowICTXLSWD_V.Item("CTN_HEIGHT") = Val(ws.Cells(r, 5).Text & "")
                    rowICTXLSWD_V.Item("CTN_WEIGHT") = Val(ws.Cells(r, 6).Text & "")

                    rowICTXLSWD_V.Item("INR_LENGTH") = Val(ws.Cells(r, 7).Text & "")
                    rowICTXLSWD_V.Item("INR_WIDTH") = Val(ws.Cells(r, 8).Text & "")
                    rowICTXLSWD_V.Item("INR_HEIGHT") = Val(ws.Cells(r, 9).Text & "")
                    rowICTXLSWD_V.Item("INR_WEIGHT") = Val(ws.Cells(r, 10).Text & "")

                    rowICTXLSWD_V.Item("ITM_LENGTH") = Val(ws.Cells(r, 11).Text & "")
                    rowICTXLSWD_V.Item("ITM_WIDTH") = Val(ws.Cells(r, 12).Text & "")
                    rowICTXLSWD_V.Item("ITM_HEIGHT") = Val(ws.Cells(r, 13).Text & "")
                    rowICTXLSWD_V.Item("ITM_WEIGHT") = Val(ws.Cells(r, 14).Text & "")

                    If vendorDimensionsUpdateMode Then
                        Dim rowICTXLSWD_orig As DataRow = dst.Tables("ICTXLSWD").NewRow
                        For Each col As DataColumn In rowICTXLSWD_V.Table.Columns
                            Dim colName As String = col.ColumnName
                            rowICTXLSWD_orig.Item(colName) = rowICTXLSWD_V.Item(colName)
                        Next
                        dst.Tables("ICTXLSWD").Rows.Add(rowICTXLSWD_orig)
                    End If
                    dst.Tables("ICTXLSWD_V").Rows.Add(rowICTXLSWD_V)

                Next
            Catch ex As Exception
                MsgBox(ex.Message, vbOKOnly, "Error Occurred")
            End Try
            grdICTXLSW3.Rows.ExpandAll(True)
            grdICTXLSWD.Rows.ExpandAll(True)
            grdICTSTYLX.Rows.ExpandAll(True)


        End If

        ASCMAIN1.Progress("", "")

    End Sub

    Sub Upload_Styles()
        Dim iResult As String = ""
        Dim API_BASE As String = ""
        Dim apiMethod As String = "UploadStyles"
        Dim API_CONTROLLER As String = "RGI/IC/" & apiMethod

        Dim client As New HttpClient()
        client.BaseAddress = New Uri(ASCMAIN1.Get_API_Endpoint(ASCMAIN1.Running_in_VS))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))
        client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", ASCMAIN1.Get_User_JWT())
        client.Timeout = New TimeSpan(0, 5, 0)

        Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()
        Dim US_REQ As New uploadStylesRequest
        US_REQ.XLS_NO = XLS_NO
        US_REQ.XLS_ENFORCE_REQD = "Yes"
        US_REQ.REQUOTE_REQ = "Yes"
        US_REQ.PO_COST_DATE = dteCostEffectiveDate.Value
        US_REQ.ICTXLSW3s = ASCDATA1.DataTableToJSON(dst.Tables("ICTXLSW3_V"))
        Dim content As HttpContent = New ObjectContent(Of uploadStylesRequest)(US_REQ, frmtr)

        Dim resp As HttpResponseMessage = Nothing
        Dim resp_err As String = ""

        Try
            ASCMAIN1.Progress("Now Uploading Style Data", XLS_NO)
            resp = client.PostAsync(API_CONTROLLER, content).Result

            Dim apiResponseString As String = Newtonsoft.Json.JsonConvert.SerializeObject(resp)

            If resp.StatusCode = Net.HttpStatusCode.OK Then
                Dim responseObject As Object = Nothing
                responseObject = resp.Content.ReadAsAsync(Of Object)().Result
                If responseObject("SUCCESS").ToString = "False" Then
                    Dim eMsgs As String = ""
                    For i As Integer = 0 To responseObject("ICTXLSW3_ERRs").COUNT - 1
                        eMsgs &= responseObject("ICTXLSW3_ERRs")(i).ToString & vbCrLf
                    Next
                    MsgBox(eMsgs, vbOKOnly, "Upload Error" & IIf(responseObject("ICTXLSW3_ERRs").COUNT > 1, "s", ""))
                    uploadMsgs &= "Upload Style Data Failed" & vbCrLf
                Else
                    'MsgBox("Success", vbOKOnly, "Upload Complete")
                    uploadMsgs &= "Upload Style Data Complete." & vbCrLf
                End If
            Else
                MsgBox("Error: " & resp.StatusCode.ToString, vbOKOnly, "API Error")
            End If
        Catch ex As Exception
            If ex.InnerException.InnerException IsNot Nothing Then
                resp_err = ex.InnerException.InnerException.Message
            Else
                resp_err = ex.InnerException.ToString
            End If

            MsgBox(resp_err, vbOKOnly, "Error Generating Spreadsheet")
        End Try
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Upload_Dimensions()
        Dim iResult As String = ""
        Dim API_BASE As String = ""
        Dim apiMethod As String = IIf(vendorDimensionsUpdateMode, "UploadStyleDimensionsManual", "UploadStyleDimensions")
        Dim API_CONTROLLER As String = "RGI/IC/" & apiMethod

        Dim client As New HttpClient()
        client.BaseAddress = New Uri(ASCMAIN1.Get_API_Endpoint(ASCMAIN1.Running_in_VS))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/x-www-form-urlencoded"))
        client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", ASCMAIN1.Get_User_JWT())
        client.Timeout = New TimeSpan(0, 5, 0)

        Dim frmtr As MediaTypeFormatter = New JsonMediaTypeFormatter()
        Dim USD_REQ As New uploadStyleDimensionsRequest
        USD_REQ.XLS_NO = XLS_NO
        USD_REQ.REQUOTE_REQ = "Yes"
        USD_REQ.VEND_CODE = Absx1.txtFor("VEND_CODE").Text & ""
        USD_REQ.ICTXLSW4s = ASCDATA1.DataTableToJSON(dst.Tables("ICTXLSWD_V"))

        Dim content As HttpContent = New ObjectContent(Of uploadStyleDimensionsRequest)(USD_REQ, frmtr)
        Dim resp As HttpResponseMessage = Nothing
        Dim resp_err As String = ""

        Try
            ASCMAIN1.Progress("Now Uploading Style Dimensions", XLS_NO)
            resp = client.PostAsync(API_CONTROLLER, content).Result

            Dim apiResponseString As String = Newtonsoft.Json.JsonConvert.SerializeObject(resp)

            If resp.StatusCode = Net.HttpStatusCode.OK Then
                Dim responseObject As Object = Nothing
                responseObject = resp.Content.ReadAsAsync(Of Object)().Result
                If responseObject("SUCCESS").ToString = "False" Then
                    Dim eMsgs As String = ""
                    importErrors = True
                    uploadMsgs &= "Upload Style Dimensions Failed." & vbCrLf
                    Dim objICTXLSW4 As Object = responseObject("ICTXLSW4_ERRs")
                    Dim ds As DataSet = Newtonsoft.Json.JsonConvert.DeserializeObject(Of DataSet)(objICTXLSW4.value)
                    Dim tblICTXLSW4 As DataTable = ds.Tables("ICTXLSW4")
                    Dim dvw As DataView = tblICTXLSW4.DefaultView
                    dvw.RowFilter = "ISNULL(IMPORT_ERROR,'')<>''"
                    grdICTXLSW4_ERRS.DataSource = tblICTXLSW4
                    MsgBox(grdICTXLSW4_ERRS.Rows.Count & " Import Errors", vbOKOnly, "Upload Error" & IIf(grdICTXLSW4_ERRS.Rows.Count > 1, "s", ""))
                    SplitContainer2.Panel2Collapsed = False
                Else
                    'MsgBox("Success", vbOKOnly, "Upload Complete")
                    uploadMsgs &= "Upload Style Dimensions Complete." & vbCrLf
                    SplitContainer2.Panel2Collapsed = True
                End If
            Else
                MsgBox("Error: " & resp.StatusCode.ToString, vbOKOnly, "API Error")
            End If
        Catch ex As Exception
            If ex.InnerException.InnerException IsNot Nothing Then
                resp_err = ex.InnerException.InnerException.Message
            Else
                resp_err = ex.InnerException.ToString
            End If

            MsgBox(resp_err, vbOKOnly, "Error Generating Spreadsheet")
        End Try
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Archive_Vendor_Spreadsheet()
        uploadMsgs &= "Vendor Response Archived." & vbCrLf
    End Sub

    Sub Manual_Style_Dimensions_Update()
        Dim SILENT As Boolean = True
        Dim record As Integer = 0
        Dim VEND_CODE As String = Absx1.txtFor("VEND_CODE").Text & ""
        ASCMAIN1.Progress("Update style price")
        For Each rowICTXLSW3_V As DataRow In dst.Tables("ICTXLSWD_V").Select
            Dim STYLE_CODE As String = rowICTXLSW3_V.Item("STYLE_CODE")
            ASCMAIN1.Progress("Get ICTSTYV1 for " & STYLE_CODE & ":" & VEND_CODE, record.ToString)
            ASCMAIN1.sql = "Select * from ICTSTYV1 where STYLE_CODE = :PARM1 and VEND_CODE = :PARM2"
            Dim rowICTSTYV1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {STYLE_CODE, VEND_CODE})
            ASCMAIN1.Progress("Calculate new price for " & STYLE_CODE & ":" & VEND_CODE, record.ToString)
            Dim NEW_STYLE_PRICE As Decimal = TAC.ICCMAIN1.Calculate_Style_Price(Me, SILENT, STYLE_CODE, , rowICTSTYV1)
            ASCMAIN1.sql = "Update ICTSTYL1 set STYLE_PRICE = " & NEW_STYLE_PRICE & " WHERE STYLE_CODE = '" & STYLE_CODE & "'"
            ASCDATA1.ExecuteSQL()
            record += 1
        Next
        uploadMsgs = record & " Styles updated."
        ASCMAIN1.Progress("", "")

    End Sub

    Class xlsRequest
        Public XLS_NO As String
        Public REQUEST_TYPE As String
    End Class

    Class uploadStylesRequest
        Public XLS_NO As String
        Public PO_COST_DATE As Date
        Public XLS_ENFORCE_REQD As String
        Public REQUOTE_REQ As String
        Public ICTXLSW3s As String
    End Class
    Class uploadStyleDimensionsRequest
        Public XLS_NO As String
        Public REQUOTE_REQ As String
        Public ICTXLSW4s As String
        Public VEND_CODE As String
    End Class
    Private Sub grdICTSTYLX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTSTYLX.InitializeRow
        Dim PCTCHG As Decimal = Val(e.Row.Cells("PCTCHG").Value & "")
        Dim NEW_PO_COST As Decimal = Val(e.Row.Cells("NEW_PO_COST").Value & "")
        Dim PO_COST As Decimal = Val(e.Row.Cells("PO_COST").Value & "")
        If PO_COST <> NEW_PO_COST Then
            If PCTCHG > 0.01 Then
                e.Row.Cells("PCTCHG").Appearance.ForeColor = Color.White
                e.Row.Cells("PCTCHG").Appearance.BackColor = Color.Red
            ElseIf PCTCHG < 0.01 Then
                e.Row.Cells("PCTCHG").Appearance.BackColor = Color.LightGreen
            Else
                e.Row.Cells("PCTCHG").Appearance.ForeColor = Color.Empty
                e.Row.Cells("PCTCHG").Appearance.BackColor = Color.Empty
            End If
        Else
            e.Row.Cells("PCTCHG").Appearance.ForeColor = Color.Empty
            e.Row.Cells("PCTCHG").Appearance.BackColor = Color.Empty
        End If

        Dim STYLE_PRICE As Decimal = Val(e.Row.Cells("STYLE_PRICE").Value & "")
        Dim NEW_STYLE_PRICE As Decimal = Val(e.Row.Cells("NEW_STYLE_PRICE").Value & "")
        Dim LIST_CALC_CODE As String = e.Row.Cells("LIST_CALC_CODE").Value & ""
        If STYLE_PRICE <> NEW_STYLE_PRICE AndAlso LIST_CALC_CODE <> "" Then
            e.Row.Cells("NEW_STYLE_PRICE").Appearance.ForeColor = Color.White
            e.Row.Cells("NEW_STYLE_PRICE").Appearance.BackColor = Color.Red
        Else
            e.Row.Cells("NEW_STYLE_PRICE").Appearance.ForeColor = Color.Empty
            e.Row.Cells("NEW_STYLE_PRICE").Appearance.BackColor = Color.Empty
        End If

    End Sub

    Private Sub grdICTXLSW3_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTXLSW3.InitializeRow
        If e.Row.Band.Key = "ICTXLSW3_ICTXLSW3_V" Then
            For Each COLUMN_NAME As String In New String() {"VEND_ITEM_CODE", "VEND_REMARK", "STYLE_SO_QTY_MIN", "INNER_PACK_QTY", "CARTON_PACK_QTY", "CASE_CUBE", "PO_COST", "STYLE_PO_QTY_MIN"}
                Dim nVal As String = e.Row.Cells(COLUMN_NAME).Value.ToString
                Dim oVal As String = e.Row.ParentRow.Cells(COLUMN_NAME).Value.ToString
                If nVal <> oVal AndAlso oVal <> "" Then
                    e.Row.ParentRow.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.ParentRow.Cells("STYLE_DESC").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells(COLUMN_NAME).Appearance.ForeColor = Drawing.Color.Red
                    e.Row.ParentRow.Cells(COLUMN_NAME).Appearance.ForeColor = Drawing.Color.Red
                End If
            Next
        End If

    End Sub

    Private Sub grdICTXLSWD_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTXLSWD.InitializeRow
        If e.Row.Band.Key = "ICTXLSWD_ICTXLSWD_V" Then
            For Each PFX As String In New String() {"CTN", "INR", "ITM"}
                For Each COL As String In New String() {"LENGTH", "WIDTH", "HEIGHT", "WEIGHT"}
                    Dim COLUMN_NAME As String = PFX & "_" & COL
                    Dim nVal As String = e.Row.Cells(COLUMN_NAME).Value.ToString
                    Dim oVal As String = e.Row.ParentRow.Cells(COLUMN_NAME).Value.ToString
                    If nVal <> oVal AndAlso oVal <> "" Then
                        e.Row.ParentRow.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Red
                        e.Row.ParentRow.Cells("STYLE_DESC").Appearance.ForeColor = Drawing.Color.Red
                        e.Row.Cells(COLUMN_NAME).Appearance.ForeColor = Drawing.Color.Red
                        e.Row.ParentRow.Cells(COLUMN_NAME).Appearance.ForeColor = Drawing.Color.Red
                    End If
                Next
            Next

        End If
    End Sub


    Private Sub UltraLabel5_Click(sender As Object, e As EventArgs) Handles UltraLabel5.Click

    End Sub
    Private Sub dteREPLY_BY_DATE_ValueChanged(sender As Object, e As EventArgs) Handles dteREPLY_BY_DATE.ValueChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If grdICTSTYLX.Selected.Rows.Count = 0 Then
            MsgBox("Please select a row(s) to update", vbOKOnly, "Cannot Proceed")
            Exit Sub
        End If

        For Each grow As UltraWinGrid.UltraGridRow In grdICTSTYLX.Selected.Rows
            grow.Cells("LIST_CALC_CODE_NEW").Value = cbeLIST_CALC_CODE.Value
            grow.Update()
        Next

    End Sub
End Class