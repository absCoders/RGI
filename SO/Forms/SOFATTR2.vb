Imports System.Drawing
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop

Public Class SOFATTR2
    Dim InquiryOnly As Boolean = False
    Dim S As New System.Text.StringBuilder() With {.Length = 0}

    'Dim STYLE_CLASS_CODE As String
    Dim SCCs As New List(Of String)
    Dim SCC_IN As String = ""
    Dim WHSE_CODE As String
    Public STYLE_CODE As String
    Dim sqlcols As String = "ICTSTYL1.STYLE_CODE,ICTSTYL1.STYLE_STATUS,ICTSTYL1.STYLE_DESC,ICTSTYL1.INNER_PACK_QTY,ICTSTYL1.CARTON_PACK_QTY,ICTSTYL1.STYLE_UOM,ICTSTYL1.STYLE_PRICE,ICTSTYL1.STYLE_CLASS_CODE,ICTSTYL1.SIZE_CODE, ICTSTYL1.VEND_CODE, ICTSTYL1.CASE_CUBE, ICTSTYL1.EXCLUSIVE_STYLE, 999 AS IMPORT_SORT, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_MATL_DESC, ICTSTYL1.COUNTRY_CODE, ICTSTYL1.PURCH_NOTES"
    Dim ATTR_CODE_1s As String
    Dim SIZE_CODEs As String
    Dim ATTR_CODE_2s As String
    Dim LAST_ATTR As String = ""
    Dim QUERY_NO As String

    Public rbadDir As String = ""
    Public IMAGES_FOLDER As String = ""
    Dim attachmentList As New Dictionary(Of String, String)
    Dim MAIL_SUBJECT As String = "Pricing and Availability Spreadsheet"
    Private mExcelProcesses() As Process
    Public progressSplash As ASFPROGS
    Dim progressSplashMsg1 As String = ""
    Dim progressSplashMsg2 As String = ""
    Dim progressSplashMsg3 As String = ""
    Dim myExcelHasBalls As Boolean = True
    Dim hotKeyPartOne As String = ""
    Dim xls_format As String = ".xls"
    Dim ORDR_NOs As New List(Of String)
    Dim isLaptop As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.KeyPreview = True 'used to grab hotkey combinations to circumvent this forms modality

        If ASCMAIN1.DBS_COMPANY = "RGO" Or ASCMAIN1.DBS_SERVER = "RGO" Then
            isLaptop = True
        Else
            isLaptop = False
        End If

        With dst
            ASCMAIN1.sql = "Select * from ICTATTRQ where INIT_OPER = :PARM1"
            Create_TDA(.Tables.Add, "ICTATTRQ", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False, "", 1)
            .Tables("ICTCLAS1").Columns.Add("SEL").DefaultValue = 0

            ASCMAIN1.sql = "Select * from ICTTHEME"
            Create_TDA(.Tables.Add, "ICTTHEME", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from ICTDISC1"
            Create_TDA(.Tables.Add, "ICTDISC1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from ICTSTDQ2"
            Create_TDA(.Tables.Add, "ICTSTDQ2", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select * from ICTWHSE1"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from ICTDUTY1"
            Create_TDA(.Tables.Add, "ICTDUTY1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from SOTBRAN1"
            Create_TDA(.Tables.Add, "SOTBRAN1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ATTR_CODE, ATTR_DESC from ICTATTR1"
            Create_TDA(.Tables.Add, "ICTATTR1_1", "**", 0, False, "", 1)
            .Tables("ICTATTR1_1").Columns.Add("SEL").DefaultValue = 0

            ASCMAIN1.sql = "Select ATTR_CODE, ATTR_DESC from ICTATTR1"
            Create_TDA(.Tables.Add, "ICTATTR1_2", "**", 0, False, "", 1)
            .Tables("ICTATTR1_2").Columns.Add("SEL").DefaultValue = 0

            ASCMAIN1.sql = "Select SIZE_CODE, SIZE_DESC from ICTSIZE1"
            Create_TDA(.Tables.Add, "ICTSIZE1", "**", 0, False, "", 1)
            .Tables("ICTSIZE1").Columns.Add("SEL").DefaultValue = 0

            ASCMAIN1.sql = "Select COLOR_CODE, COLOR_DESC from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select " & sqlcols & ", ICTCOLR1.COLOR_CODE,ICTATTR1.ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS,ICTCOLR1.COLOR_GROUP_CODE" _
                & " from ICTSTYL1,ICTCOLR1,ICTATTR1,ICTSTYC1"
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 0)
            With .Tables("ICTSTYL1")
                .Columns.Add("SEL").DefaultValue = 0
                .Columns.Add("ONH", GetType(System.Int32))
                .Columns.Add("ONPO", GetType(System.Int32))
                .Columns.Add("TRAN", GetType(System.Int32))
                .Columns.Add("OPEN", GetType(System.Int32))
                .Columns.Add("PICK", GetType(System.Int32))
                .Columns.Add("OTS", GetType(System.Int32), "ISNULL(ONH,0)-ISNULL(OPEN,0)-ISNULL(PICK,0)")
                .Columns.Add("FUT", GetType(System.Int32), "ISNULL(OTS,0)+ISNULL(ONPO,0)+ISNULL(TRAN,0)")
                .Columns.Add("FUT1_DATE", GetType(System.DateTime))
                .Columns.Add("FUT1_AVAIL", GetType(System.Int32))
                .Columns.Add("FUT2_DATE", GetType(System.DateTime))
                .Columns.Add("FUT2_AVAIL", GetType(System.Int32))
                .Columns.Add("FUT3_DATE", GetType(System.DateTime))
                .Columns.Add("FUT3_AVAIL", GetType(System.Int32))
                .Columns.Add("PBH1", GetType(System.String))
                .Columns.Add("PBV1", GetType(System.String))
                .Columns.Add("PBH2", GetType(System.String))
                .Columns.Add("PBV2", GetType(System.String))
                .Columns.Add("PBH3", GetType(System.String))
                .Columns.Add("PBV3", GetType(System.String))
                .Columns.Add("PBH4", GetType(System.String))
                .Columns.Add("PBV4", GetType(System.String))
                .Columns.Add("IMAGE_LOC_API", GetType(System.String))
                .Columns.Add("IMAGE_LOC_LOCAL", GetType(System.String))
                .Columns("COLOR_CODE").AllowDBNull = True
                .Columns.Add("FACTORY", GetType(System.String))
                .Columns.Add("PORT", GetType(System.String))
                .Columns.Add("FEPRICE", GetType(System.String))
                .Columns.Add("FDPRICE", GetType(System.String))
                .Columns.Add("FEMPRICE", GetType(System.String))
                .Columns.Add("FDMPRICE", GetType(System.String))
                .Columns.Add("THEME_CODE", GetType(System.String))
                .Columns.Add("SEASON_CODE", GetType(System.String))
                .Columns.Add("THEME_DESC", GetType(System.String))
                .Columns.Add("DISC_DATE", GetType(System.DateTime))
                .Columns.Add("LENGTH_ITM", GetType(System.Double))
                .Columns.Add("WIDTH_ITM", GetType(System.Double))
                .Columns.Add("HEIGHT_ITM", GetType(System.Double))
                .Columns.Add("WEIGHT_ITM", GetType(System.Double))
                .Columns.Add("LENGTH_CTN", GetType(System.Double))
                .Columns.Add("WIDTH_CTN", GetType(System.Double))
                .Columns.Add("HEIGHT_CTN", GetType(System.Double))
                .Columns.Add("WEIGHT_CTN", GetType(System.Double))
                .Columns.Add("LENGTH_INR", GetType(System.Double))
                .Columns.Add("WIDTH_INR", GetType(System.Double))
                .Columns.Add("HEIGHT_INR", GetType(System.Double))
                .Columns.Add("WEIGHT_INR", GetType(System.Double))
                .Columns.Add("LIGHT_TYPE", GetType(System.String))
                .Columns.Add("LIGHT_COLOR", GetType(System.String))
                .Columns.Add("DUTY_RATE", GetType(System.String))
                '--- PVC Data ---
                .Columns.Add("PVC_COLORS", GetType(System.String))
                .Columns.Add("PVC_HEIGHT", GetType(System.String))
                .Columns.Add("PVC_DIAMETER", GetType(System.String))
                .Columns.Add("PVC_LIGHT_TYPE_DESC", GetType(System.String))
                .Columns.Add("PVC_TIP_COUNT", GetType(System.String))
                .Columns.Add("PVC_LIGHT_COUNT", GetType(System.String))
                For i As Integer = 2 To 6
                    .Columns.Add(String.Format("ATTR_CODE{0}", i), GetType(System.String))
                Next
                .Columns.Add("ATTR_CODE_ALL", GetType(System.String))
                '--- PO Data ---
                .Columns.Add("TARIFF_PCT", GetType(System.Double))
                .Columns.Add("PO_COST", GetType(System.Double))
                .Columns.Add("STYLE_PO_QTY_MIN", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL1 where STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "RIBBON1", "**", 0, False, "V", 0)
            .Tables("RIBBON1").Columns.Add("DATEPRINTED", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("BOXQTY", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("CARTQTY", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("COLORS", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("COLORS3", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("Price1_LBL", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("Price1_AMT", GetType(System.Double))
            .Tables("RIBBON1").Columns.Add("Price2_LBL", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("Price2_AMT", GetType(System.Double))
            .Tables("RIBBON1").Columns.Add("Price3_LBL", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("Price3_AMT", GetType(System.Double))
            .Tables("RIBBON1").Columns.Add("Price4_LBL", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("Price4_AMT", GetType(System.Double))
            .Tables("RIBBON1").Columns.Add("COLORSDESC", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("VEND_SUPPLIER_ID", GetType(System.String))
            .Tables("RIBBON1").Columns.Add("PORT_CODE_ORIG", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM ICTSTYL3 where STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False, "V", 0)


            If Not isLaptop Then
                S.Length = 0
                S.AppendLine("SELECT *")
                S.AppendLine("FROM ECTPRCG1")
                S.AppendLine("WHERE PRCG_NO = :PARM1")
                ASCMAIN1.sql = S.ToString()
                Create_TDA(.Tables.Add, "ECTPRCG1", "**", 0, True, "V", 1)

                S.Length = 0
                S.AppendLine("SELECT *")
                S.AppendLine("FROM ECTPRCG2")
                S.AppendLine("WHERE PRCG_NO = :PARM1")
                ASCMAIN1.sql = S.ToString()
                Create_TDA(.Tables.Add, "ECTPRCG2", "**", 2, True, "V", 2)

                S.Length = 0
                S.AppendLine("SELECT *")
                S.AppendLine("FROM ECTPRCG3")
                S.AppendLine("WHERE PRCG_NO = :PARM1")
                ASCMAIN1.sql = S.ToString()
                Create_TDA(.Tables.Add, "ECTPRCG3", "**", 3, True, "V", 3)

            End If

        End With

        grdICTATTRQ.DataSource = dst.Tables("ICTATTRQ")
        grdICTCLAS1.DataSource = dst.Tables("ICTCLAS1")

        grdICTSTYL1.DataSource = dst.Tables("ICTSTYL1")
        Create_Summary(grdICTSTYL1, "STYLE_CODE", "Count")

        grdICTSIZE1.DataSource = dst.Tables("ICTSIZE1")
        Create_Summary(grdICTSIZE1, "SIZE_CODE", "Count")
        Create_Summary(grdICTSIZE1, "SEL")

        grdICTATTR1_1.DataSource = dst.Tables("ICTATTR1_1")
        Create_Summary(grdICTATTR1_1, "ATTR_CODE", "Count")
        Create_Summary(grdICTATTR1_1, "SEL")

        grdICTATTR1_2.DataSource = dst.Tables("ICTATTR1_2")
        Create_Summary(grdICTATTR1_2, "ATTR_CODE", "Count")
        Create_Summary(grdICTATTR1_2, "SEL")

        Fill_Records("ICTCLAS1")
        Fill_Records("ICTTHEME")

        'AddAllClass()
        Sort_grdColumns(grdICTCLAS1, "STYLE_CLASS_CODE")
        Fill_Records("ICTDISC1")
        Fill_Records("ICTCOLR1")
        Fill_Records("ICTWHSE1")
        Fill_Records("ICTDUTY1")
        Fill_Records("SOTBRAN1")

        grdICTSTYL1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        With grdICTSTYL1.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Debug.Print(gcol.Key)
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ONH", "ONPO", "TRAN", "OPEN", "PICK"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    gcol.Hidden = True
                    gcol.Width = 60
                    gcol.Format = "#,##0"
                ElseIf New String() {"OTS", "FUT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                    gcol.Width = 60
                    gcol.Format = "#,##0"
                    If gcol.Key = "OTS" Then
                        gcol.CellAppearance.BackColor = Color.LightGreen
                    End If
                ElseIf New String() {"FUT1_DATE", "FUT2_DATE", "FUT3_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.HotPink
                    gcol.Width = 60
                    gcol.Format = "MM/dd"
                ElseIf New String() {"FUT1_AVAIL", "FUT2_AVAIL", "FUT3_AVAIL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.HotPink
                    gcol.Width = 60
                    gcol.Format = "#,##0"
                ElseIf New String() {"PBH1", "PBH2", "PBH3", "PBH4"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LawnGreen
                    gcol.Width = 60
                    gcol.Hidden = True
                ElseIf New String() {"PBV1", "PBV2", "PBV3", "PBV4", "FEPRICE", "FDPRICE", "FEMPRICE", "FDMPRICE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LawnGreen
                    gcol.Width = 60
                    gcol.Format = "####,##0.00"
                    gcol.Hidden = True
                ElseIf New String() {"LENGTH_IT", "WIDTH_IT", "HEIGHT_IT", "WEIGHT_IT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.MediumVioletRed
                    gcol.Width = 100
                    gcol.Format = "####,##0.00"
                    gcol.Hidden = True
                ElseIf New String() {"LENGTH_CTN", "WIDTH_CTN", "HEIGHT_CTN", "WEIGHT_CTN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.MediumBlue
                    gcol.Width = 100
                    gcol.Format = "####,##0.00"
                    gcol.Hidden = True
                ElseIf New String() {"LENGTH_INR", "WIDTH_INR", "HEIGHT_INR", "WEIGHT_INR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.MediumPurple
                    gcol.Width = 100
                    gcol.Format = "####,##0.00"
                    gcol.Hidden = True
                ElseIf New String() {"IMAGE_LOC_API"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Wheat
                    gcol.Width = 100
                    gcol.Hidden = True
                ElseIf New String() {"IMAGE_LOC_LOCAL"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Wheat
                    gcol.Width = 100
                    gcol.Hidden = True
                ElseIf New String() {"LONG_COLOR"}.Contains(gcol.Key) Then
                    'gcol.Header.Appearance.BackColor2 = Color.Wheat
                    gcol.Header.Caption = "Color"
                    gcol.Hidden = True
                ElseIf New String() {"FACTORY", "PORT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Olive
                    gcol.Hidden = True
                ElseIf New String() {"THEME_CODE"}.Contains(gcol.Key) Then
                    gcol.Hidden = True
                ElseIf New String() {"SEASON_CODE"}.Contains(gcol.Key) Then
                    gcol.Hidden = False
                ElseIf New String() {"THEME_DESC"}.Contains(gcol.Key) Then
                    gcol.Hidden = False
                    gcol.Width = 100
                ElseIf New String() {"TARIFF_PCT", "PO_COST", "STYLE_PO_QTY_MIN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.Coral
                    gcol.Hidden = True
                End If
            Next
        End With

        ASCMAIN1.Add_Value_List(grdICTSTYL1, "STYLE_STATUS", Nothing, New String() {":", "A:Active", "D:Discontinued", "N:Do Not Re-Order"})
        ASCMAIN1.Add_Value_List(grdICTSTYL1, "STYLE_COLOR_STATUS", Nothing, New String() {":", "A:Active", "D:Discontinued", "N:Do Not Re-Order"})

        'splICTSTYL1.Visible = False
        TabStyles.Visible = False
        grdICTATTRQ.Visible = True
        Show_Filter(grdICTSTYL1, True)

        'Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
        Absx1.txtFor("WHSE_CODE").Text = "MS"

        ReParent_Tabs(splChoices.Panel2.Controls("tab1"))

        'grdICTATTR1_1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Initialize_tabAttributes()

        optFEFD.Value = "FE"
        Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        If Not IsNothing(rowSOTPARM3) Then
            IMAGES_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString
            If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                rbadDir = ASCMAIN1.Folders("Temp")
            Else
                rbadDir = rowSOTPARM3.Item("RO_PARM_EXCEL_DIR").ToString()
            End If
        Else
            IMAGES_FOLDER = "C:\"
            rbadDir = "C:\"
        End If

        If Not isLaptop Then
            Dim ECOM_LIST As New List(Of String)
            ECOM_LIST.Add("")
            Dim sql As New StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT ECOM_CODE FROM ECTECOM1 ORDER BY ECOM_CODE")
            Dim tblE As DataTable = ASCDATA1.GetDataTable(sql.ToString())
            For Each rowE As DataRow In tblE.Rows
                ECOM_LIST.Add(rowE.Item("ECOM_CODE").ToString & String.Empty)
            Next
            cboECOMPRICING.DataSource = ECOM_LIST
            cboECOMPRICING.SelectedIndex = 0
        End If

        tab.Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Find"
                'STYLE_CLASS_CODE = Absx1.txtFor("STYLE_CLASS_CODE").Text
                'Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)
                'If rowICTCLAS1 Is Nothing Then
                '    EMsg &= vbCr & "You Must Select (at minimum) a valid Class Code"
                'End If
                If SCCs.Count = 0 Then
                    EMsg &= vbCr & "You Must Select (at minimum) a valid Class Code"
                End If

                WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & "You Must Select a valid Warehouse Code"
                End If

                Dim selectedAttributes() = dst.Tables("ICTATTR1_1").Select("ISNULL(SEL,'0') = '1'")
                If selectedAttributes.Length = 0 Then
                    EMsg &= vbCr & "You Must Select (at minimum) one Attribute"
                End If
            Case "Clear"

            Case "Done"

            Case "Email"
                If dst.Tables("ICTSTYL1").Select("SEL='1'").Length = 0 Then
                    EMsg &= vbCr & "No Styles Selected"
                End If
            Case "Excel"
                If dst.Tables("ICTSTYL1").Select("SEL='1'").Length = 0 Then
                    EMsg = EMsg & vbCrLf & "No Styles Selected"
                End If

                If Not chkNonPVC.Checked Then
                    If IsNothing(optPRICE_TIER.Value) Then
                        EMsg = EMsg & vbCrLf & "You Must Select A Price Tier If Not Using Standard"
                    Else
                        If optPRICE_TIER.Value = "SP" Then
                            If IsNothing(optDISC_PCT_EXTRA.Value) And IsDBNull(numDISC_PCT.Value) Then
                                EMsg = EMsg & vbCrLf & "You Must Select A Percentage When Selecting Disc %"
                            End If
                        End If
                    End If
                End If

                If Not chkPVC.Checked Then
                    If IsNothing(optPRICE_TIER_PVC.Value) Then
                        EMsg = EMsg & vbCrLf & "You Must Select A Case Level for PVC If Not Using Standard"
                    End If
                End If

                If chkPBFE.Checked Then
                    If Not IsNumeric(numFEFDFACTOR.Value) Then
                        numFEFDFACTOR.Value = 0
                    Else
                        If numFEFDFACTOR.Value < -14 Or numFEFDFACTOR.Value > 0 Then
                            EMsg = EMsg & vbCrLf & "Discount factor must be between -14 and 0"
                        End If
                    End If
                End If
            Case "Attribute Excel"
                Dim SelOnlyWhere As String = "SEL = '1'"
                Dim SelRows As Int64 = dst.Tables.Item("ICTSTYL1").Select(SelOnlyWhere, "").Count
                If SelRows = 0 Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "No Rows Selected!"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine("You Must Select At Least One")
                    iMSG.AppendLine("Row From The Grid To Create")
                    iMSG.AppendLine("An Excel File.")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Would You Like Me To Select")
                    iMSG.AppendLine("All Of Them And Proceed?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.Yes Then
                        For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select()
                            rowICTSTYL1.Item("SEL") = "1"
                            SelRows += 1
                        Next
                    Else
                        EMsg = EMsg & vbCrLf & "No Rows Selected."
                    End If
                End If
                If SelRows > 250 Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Thats A Lot!"
                    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                    iMSG.AppendLine($"You Have Selected {SelRows} Rows!")
                    iMSG.AppendLine("")
                    iMSG.AppendLine("Are You Sure You Want To Procced?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult <> MsgBoxResult.Yes Then
                        EMsg = EMsg & vbCrLf & "Too Many Rows."
                    End If
                End If
            Case "Import"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Pick Style/Colors from Spreadsheet?"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Prompt You For An Excel File")
                iMSG.AppendLine("That Will Be Used To Populate The Styles.")
                iMSG.AppendLine("It Requires That At least One Column Be")
                iMSG.AppendLine("Titled 'Style Code' and One Column Be")
                iMSG.AppendLine("Titled 'Color Code'.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Are You Ready?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg = EMsg & vbCrLf & "OK.  Maybe Some Other Time."
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Find"
                Call Mode_Settings(True)
                Find_Styles()
                showSelectors(False)
            Case "Clear"
                Call Mode_Settings(False)
                txtOrder.Text = ""
                'Absx1.txtFor("STYLE_CLASS_CODE").Text = ""
                SCC_CLEAR(True)
                Initialize_tabAttributes()
                chkSTYLE_STATUS_D.Checked = True
                chkSTYLE_STATUS_A.Checked = True
                chkSTYLE_STATUS_N.Checked = True
                TabStyles.Visible = False
                grdICTATTRQ.Visible = True
                ORDR_NOs.Clear()
                showSelectors(True)
            Case "Done"
                Call Mode_Settings(False)
                Me.Close()
            Case "Email"
                attachmentList.Clear()
                Me.Cursor = Cursors.WaitCursor
                Dim steps As Integer = IIf(chkAttachZip.Checked, 2, 1)

                DisplaySplash("Composing Email", "Generating Excel Workbook", "Stage 1 of " & steps.ToString)
                ExcelProcessInit()
                Dim excelFile As String = Generate_Excel()
                ExcelProcessKill()

                If excelFile <> "" Then
                    Dim excelFileAttachment As String = rbadDir & excelFile & xls_format
                    attachmentList.Add("Spreadsheet", excelFileAttachment)
                End If

                If chkAttachZip.Checked Then
                    progressSplash.UpdateProgress("", "Building Zip File", "Stage 2 of 2", 10)
                    Dim zipFile As String = Generate_Zip_file()
                    If zipFile <> "" Then
                        Dim zipFileAttachment As String = rbadDir & zipFile
                        If My.Computer.FileSystem.FileExists(zipFileAttachment) Then
                            attachmentList.Add("Images", zipFileAttachment)
                        End If
                    End If
                End If

                CloseSplash()

                Me.Cursor = Cursors.Default

                Dim laptopUser As Boolean = True

                If laptopUser Then
                    Dim MAIL_BODY As String = ""
                    Dim attachments() As String
                    ReDim attachments(1)
                    Dim attachmentNo As Integer = 0
                    Dim attachmentItem As KeyValuePair(Of String, String)

                    For Each attachmentItem In attachmentList
                        attachments(attachmentNo) = attachmentItem.Value
                        attachmentNo += 1
                    Next
                    Create_Outlook_mailitem("", "", MAIL_SUBJECT, "", attachments)
                    MsgBox("Email Saved as Draft", vbInformation + vbOKOnly, "Done")
                Else
                    Send_Email(attachmentList)
                End If
            Case "Excel"
                DisplaySplash("Now Generating Excel Workbook", "", "")
                Me.Cursor = Cursors.WaitCursor
                ExcelProcessInit()
                Dim excelFile As String = Generate_Excel()
                ExcelProcessKill()
                Me.Cursor = Cursors.Default
                CloseSplash()

                If excelFile <> "" Then
                    Dim start_excel As New Process
                    start_excel.StartInfo.Arguments = """" + excelFile + """ /e"
                    start_excel.StartInfo.FileName = rbadDir & excelFile & xls_format
                    start_excel.Start()
                End If
            Case "Attribute Excel"
                DisplaySplash("Now Generating Attribute Excel", "", "")
                Me.Cursor = Cursors.WaitCursor
                ExcelProcessInit()
                Dim excelFile As String = Generate_Excel_WPICS()

                Show_Document(excelFile)

                CloseSplash()
                Me.Cursor = Cursors.Default
            Case "Zip"
                DisplaySplash("Now Building Zip File", "", "")
                Generate_Zip_file()
                CloseSplash()
            Case "Import"
                Dim Success As Boolean = Import_From_Excel()
                Call Mode_Settings(Success)
                showSelectors(Not Success)
        End Select
    End Sub

    Private Function Generate_Excel_WPICS() As String
        Dim RetVal As String = ""
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim xls_name As String = ""
        Dim xls_file_name As String = ""
        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim RowStyleColors As New Dictionary(Of Int64, String())
        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                xls_file_name = xls_path & "\" & xls_name & ".XLSx"
                If Not My.Computer.FileSystem.FileExists(xls_file_name) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook()
        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i

        oSheet = oWB.Worksheets(0)
        oSheet.Name = "Search By Attribute"
        ASCMAIN1.Progress("-", oSheet.Name)
        Dim SORT_ORDER As String = ""
        For Each SRT_COL As UltraWinGrid.UltraGridColumn In grdICTSTYL1.DisplayLayout.Bands(0).SortedColumns
            SORT_ORDER = SORT_ORDER & "," & SRT_COL.Key
            If SRT_COL.SortIndicator = UltraWinGrid.SortIndicator.Descending Then
                SORT_ORDER = SORT_ORDER & " DESC"
            End If
        Next
        If SORT_ORDER.StartsWith(",") Then
            SORT_ORDER = SORT_ORDER.Substring(1, SORT_ORDER.Length - 1)
        End If

        Dim SelOnlyWhere As String = "SEL = '1'"
        Load_DataTable_into_SGXLS(1, 3, dst.Tables.Item("ICTSTYL1"), oSheet, grdICTSTYL1, Nothing, SORT_ORDER, SelOnlyWhere)
        'Lock and color header row.
        'Dim dbRows As Int64 = grdICTSTYL1.Rows.Count
        Dim dbRows As Int64 = dst.Tables.Item("ICTSTYL1").Select(SelOnlyWhere, "").Count
        Dim dbCols As Int64 = grdICTSTYL1.DisplayLayout.Bands(0).Columns.Count
        oSheet.Cells(0, 0).Value = "Image"
        oSheet.Cells(0, 0).EntireColumn.ColumnWidth = 25
        oSheet.Cells(0, 1).Value = "SKU"
        oSheet.Cells(0, 1).EntireColumn.ColumnWidth = 25
        Dim STYLE_COL As Int64 = 0
        Dim COLOR_COL As Int64 = 0
        Dim DISC_COL As Int64 = 0
        Dim LIST_COL As Int64 = 0
        For col As Int64 = 0 To dbCols + 2
            If oSheet.Cells(0, col).Interior.Color = SpreadsheetGear.Colors.Transparent Then
                oSheet.Cells(0, col).Interior.Color = SpreadsheetGear.Colors.LightBlue
            End If
            If oSheet.Cells(0, col).Text.Trim.ToUpper = "STYLE CODE" Then
                STYLE_COL = col
            End If
            If oSheet.Cells(0, col).Text.Trim.ToUpper = "COLOR" Then
                COLOR_COL = col
            End If
            If oSheet.Cells(0, col).Text.Trim.ToUpper = "PURCH NOTES" Then
                DISC_COL = col + 1
                oSheet.Cells(0, DISC_COL).Value = "Price"
            End If
            If oSheet.Cells(0, col).Text.Trim.ToUpper = "LIST PRICE" Then
                If chkDiscSheets.Checked And numDiscSheets.Value > 0 Then
                    LIST_COL = col
                End If
            End If
        Next
        Dim r As Int64 = 0
        For rw As Int64 = 1 To dbRows
            r += 1
            oSheet.Cells(rw, 1).RowHeight = 100
            If STYLE_COL > 0 And COLOR_COL > 0 Then
                Dim STYLE_CODE As String = oSheet.Cells(rw, STYLE_COL).Text
                Dim COLOR_CODE As String = oSheet.Cells(rw, COLOR_COL).Text
                RowStyleColors.Add(rw, {STYLE_CODE, COLOR_CODE})
                oSheet.Cells(rw, 1).Value = $"{STYLE_CODE}-{COLOR_CODE}"
                If DISC_COL > 0 And LIST_COL > 0 Then
                    Dim LIST_PRICE As Decimal = Val(oSheet.Cells(rw, LIST_COL).Value)
                    oSheet.Cells(rw, DISC_COL).Value = LIST_PRICE * (1 - (numDiscSheets.Value / 100))
                    oSheet.Cells(rw, DISC_COL).NumberFormat = "###,###,###.00"
                End If
            End If
        Next

        Dim msgs As Dictionary(Of String, String) = TAC.TACMAIN1.getSalesDocMsgs("")

        ''CC Notice
        'r += 3
        'oSheet.Range($"A{r}:N{r}").Merge()
        'oSheet.Range($"A{r}:N{r}").Value = "We accept MasterCard, Visa, and Discover. Credit cards are charged approximately one week prior to shipment for the product and estimated shipping charges. Any difference at the time of shipment will be charged or credited to the same card. Each shipment will be charged separately."
        'oSheet.Range($"A{r}:N{r}").Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
        'oSheet.Range($"A{r}:N{r}").Borders.Weight = SpreadsheetGear.BorderWeight.Thin
        'oSheet.Range($"A{r}:N{r}").Font.Bold = True
        'oSheet.Range($"A{r}:N{r}").Font.Color = SpreadsheetGear.Colors.Red
        'oSheet.Range($"A{r}:N{r}").RowHeight = oSheet.Range($"A{r}:N{r}").RowHeight * 2
        'oSheet.Range($"A{r}:N{r}").WrapText = True
        ''oSheet.Range($"A{r}:N{r}").Font.Size = 8

        If msgs("T").Length > 0 Then
            '2025 Tariff Notice
            r += 3
            oSheet.Range($"A{r}:N{r}").Merge()
            oSheet.Range($"A{r}:N{r}").Value = msgs("T")
            oSheet.Range($"A{r}:N{r}").Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
            oSheet.Range($"A{r}:N{r}").Borders.Weight = SpreadsheetGear.BorderWeight.Thin
            oSheet.Range($"A{r}:N{r}").Font.Bold = True
            oSheet.Range($"A{r}:N{r}").Font.Color = SpreadsheetGear.Colors.Red
            oSheet.Range($"A{r}:N{r}").RowHeight = oSheet.Range($"A{r}:N{r}").RowHeight * 2
            oSheet.Range($"A{r}:N{r}").WrapText = True
            'oSheet.Range($"A{r}:N{r}").Font.Size = 8
        End If

        oSheet.Range(0, 0).Select()
        oSheet.WindowInfo.FreezePanes = True
        oSheet.Range("A1:A1").Select()

        oWB.SaveAs(xls_file_name, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        oWB = Nothing
        AddImagesToExcel(xls_file_name, RowStyleColors)
        RetVal = xls_file_name
        Return RetVal
    End Function

    Private Sub AddImagesToExcel(ByVal xls_file_name As String, ByVal rowStyleColors As Dictionary(Of Long, String()))
        Dim excel As Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Dim XWB As Excel.Workbook = excel.Workbooks.Open(xls_file_name)
        Dim XWS As Excel.Worksheet = XWB.Sheets(1)
        Dim rng As Excel.Range
        For Each rw As KeyValuePair(Of Long, String()) In rowStyleColors
            rng = XWS.Range($"A{rw.Key + 1}:A{rw.Key + 1}")
            Dim STYLE_CODE As String = rw.Value(0)
            Dim COLOR_CODE As String = rw.Value(1)
            Dim IMAGE_NAME As String = STYLE_CODE & "-" & COLOR_CODE & ".JPG"

            If IMAGE_NAME <> "" Then
                Dim FILENAME As String = IMAGES_FOLDER & "\" & IMAGE_NAME
                If chkWebImages.Checked Then
                    getWebImage(FILENAME, STYLE_CODE, COLOR_CODE)
                End If
                'rng = XWS.Range("A" & CStr(r) & ":" & "B" & CStr(r + 4))
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    InsertPictureInRange(FILENAME, rng, XWS, STYLE_CODE, COLOR_CODE)
                Else
                    rng.MergeCells = True
                    rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    rng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    rng.FormulaR1C1 = "No Image Available"
                    rng.Font.Bold = True
                End If
            End If
        Next
        XWB.Save()
        XWB.Close()
        XWB = Nothing
        excel = Nothing
    End Sub

    Private Sub SCC_CLEAR(ByRef clearSelection As Boolean)
        SCCs.Clear()
        SCC_IN = ""
        If clearSelection Then
            For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select()
                rowICTCLAS1.Item("SEL") = "0"
            Next
        End If
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Find").Visible = True
                .Groups("Screen Control").Items("Clear").Visible = ScreenMode
                .Groups("Screen Control").Items("Done").Visible = True
                .Groups("Screen Control").Items("Email").Visible = ScreenMode
                .Groups("Screen Control").Items("Excel").Visible = ScreenMode
                .Groups("Screen Control").Items("Attribute Excel").Visible = ScreenMode
                .Groups("Screen Control").Items("Import").Visible = True
                .Groups("Screen Control").Items("Zip").Visible = ScreenMode
                If isLaptop Then
                    .Groups("Ecom Pricing").Visible = False
                Else
                    .Groups("Ecom Pricing").Visible = tf
                End If
                SetOptionsVisible(ScreenMode)
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

    End Sub

    Private Sub SetOptionsVisible(screenMode As Boolean)
        chkAttachZip.Visible = screenMode
        chkGroupBySize.Visible = screenMode
        chkLongColors.Visible = screenMode
    End Sub

    Sub Clear_Record()
        'dst.Tables("XXXXXXXX").Rows.Clear()
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        'Call Fill_Records("ARTCUST1", Absx1.txtFor("CUST_CODE").Text, True)
        'Call Fill_Records("ARTCUST2", Absx1.txtFor("CUST_CODE").Text, True)

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()
        'For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        'Next
        'Call CommitTrans("Order / Quote Deleted")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'TODO: Remove this stop before going live.
        Stop
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        'Print_Report_Begin()
        'Generate_Report("SORORDRO")
        'Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYL1, "SSSSBBBBSSSSBBSBSSSSSB", "Show Filter", "Show GroupBox", "Show Pins", "Show All Status Qtys", "Select All", "De-Select All", "Select Selected", "De-Select Selected", "Calc Price Breaks", "Show Factory/Port", "Show Disc Date", "Show Extended Pack", "Print Ribbon Sheet", "Print Ribbon Combined", "Show All Attributes", "Style Masterfile", "Show E-Commerce", "Show Lighting", "Show PVC", "Show Purch Notes", "Show PO Info", "Dump Data To Excel")
        Load_Popup_Menu(grdICTATTR1_1, "BBBB", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdICTATTR1_2, "BBBB", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdICTSIZE1, "BBBB", "Select All", "De-Select All", "Select Selected", "De-Select Selected")
        Load_Popup_Menu(grdICTCLAS1, "BBBB", "Select All", "De-Select All", "Select Selected", "De-Select Selected")

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

        'tlb_sbt = DirectCast(tlb.Tools("Show PO Info"), UltraWinToolbars.StateButtonTool)
        If e.SourceControl.Name = "grdICTSTYL1" Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show PO Info"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = ASCMAIN1.CLIENT = "RGI"
        End If

        If grd.Name = "grd" Then
            Exit Sub
        End If

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        '  Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Select Selected", "De-Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                    grow.Update()
                Next

            Case "Show All Status Qtys"
                tlb_sbt = DirectCast(tlb.Tools("Show All Status Qtys"), UltraWinToolbars.StateButtonTool)
                For Each COLUMN_NAME As String In New String() {"ONH", "ONPO", "TRAN", "OPEN", "PICK"}
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                Next
            Case "Calc Price Breaks"
                Dim Discounts As List(Of DISCOUNTS)
                Dim LastStyle As String = ""
                Dim rowARTCUST1 As DataRow = Nothing
                Me.Cursor = Cursors.WaitCursor
                tlb_sbt = DirectCast(tlb.Tools("Calc Price Breaks"), UltraWinToolbars.StateButtonTool)
                For Each COLUMN_NAME As String In New String() {"PBH1", "PBV1", "PBH2", "PBV2", "PBH3", "PBV3", "PBH4", "PBV4", "FEPRICE", "FDPRICE", "FEMPRICE", "FDMPRICE", "IMAGE_LOC_LOCAL", "IMAGE_LOC_API"}
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                Next

                Dim bAreEventsInFilteredMode As Boolean = False
                For Each column As UltraWinGrid.UltraGridColumn In grdICTSTYL1.DisplayLayout.Bands(0).Columns
                    If grdICTSTYL1.DisplayLayout.Bands(0).ColumnFilters(column).FilterConditions.Count > 0 Then
                        bAreEventsInFilteredMode = True
                    End If
                Next
                If bAreEventsInFilteredMode Then
                    MsgBox("Only Filtered Rows Showing Will Calculate Price Breaks", vbOKOnly, "Filters In Effect")
                End If
                For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                    Dim row As UltraWinGrid.UltraGridRow = grdICTSTYL1.Rows.GetRowWithListIndex(dst.Tables("ICTSTYL1").Rows.IndexOf(rowICTSTYL1))
                    If Not row.IsFilteredOut Then
                        If LastStyle <> rowICTSTYL1.Item("STYLE_CODE").ToString Then
                            LastStyle = rowICTSTYL1.Item("STYLE_CODE").ToString
                        End If
                        Discounts = SOCMAIN2.Price_Discounts(Me, "", rowARTCUST1, LastStyle, False)
                        If Discounts(3).DISCOUNT_QTY = 0 Then
                            rowICTSTYL1.Item("PBH1") = Null
                            rowICTSTYL1.Item("PBV1") = Null
                        Else
                            rowICTSTYL1.Item("PBH1") = Discounts(3).DISCOUNT_QTY
                            rowICTSTYL1.Item("PBV1") = Format(Discounts(3).DISCOUNT_PRICE, "###,##0.00")
                        End If
                        If Discounts(2).DISCOUNT_QTY = 0 Then
                            rowICTSTYL1.Item("PBH2") = Null
                            rowICTSTYL1.Item("PBV2") = Null
                        Else
                            rowICTSTYL1.Item("PBH2") = Discounts(2).DISCOUNT_QTY
                            rowICTSTYL1.Item("PBV2") = Format(Discounts(2).DISCOUNT_PRICE, "###,##0.00")
                        End If

                        If Discounts(1).DISCOUNT_QTY = 0 Then
                            rowICTSTYL1.Item("PBH3") = Null
                            rowICTSTYL1.Item("PBV3") = Null
                        Else
                            rowICTSTYL1.Item("PBH3") = Discounts(1).DISCOUNT_QTY
                            rowICTSTYL1.Item("PBV3") = Format(Discounts(1).DISCOUNT_PRICE, "###,##0.00")
                        End If

                        If Discounts(0).DISCOUNT_QTY = 0 Then
                            rowICTSTYL1.Item("PBH4") = Null
                            rowICTSTYL1.Item("PBV4") = Null
                        Else
                            rowICTSTYL1.Item("PBH4") = Discounts(0).DISCOUNT_QTY
                            rowICTSTYL1.Item("PBV4") = Format(Discounts(0).DISCOUNT_PRICE, "###,##0.00")
                        End If

                        Dim FEFD As New FEFDPrice(Me, rowICTSTYL1.Item("STYLE_CODE").ToString)
                        If FEFD.ErrorMsg = "" Then
                            rowICTSTYL1.Item("FEPRICE") = Format(FEFD.FEPrice, "###,##0.00")
                            rowICTSTYL1.Item("FDPRICE") = Format(FEFD.FDPrice, "###,##0.00")
                            rowICTSTYL1.Item("FEMPRICE") = Format(FEFD.FEMixPrice, "###,##0.00")
                            rowICTSTYL1.Item("FDMPRICE") = Format(FEFD.FDMixPrice, "###,##0.00")
                        Else
                            rowICTSTYL1.Item("FEPRICE") = Format(0, "###,##0.00")
                            rowICTSTYL1.Item("FDPRICE") = Format(0, "###,##0.00")
                            rowICTSTYL1.Item("FEMPRICE") = Format(0, "###,##0.00")
                            rowICTSTYL1.Item("FDMPRICE") = Format(0, "###,##0.00")
                        End If

                        Dim IMAGE_NAME As String = GetImageLocation(rowICTSTYL1.Item("STYLE_CODE").ToString, rowICTSTYL1.Item("COLOR_CODE").ToString)
                        If IMAGE_NAME.Length > 0 Then
                            If Not IMAGES_FOLDER.EndsWith("\") Then
                                IMAGES_FOLDER = IMAGES_FOLDER & "\"
                            End If
                            rowICTSTYL1.Item("IMAGE_LOC_LOCAL") = IMAGES_FOLDER & IMAGE_NAME
                            rowICTSTYL1.Item("IMAGE_LOC_API") = "http://api.regency-rib.com:8181/images/product/" & IMAGE_NAME
                        End If
                    End If
                Next
                grdICTSTYL1.UpdateData()
                Me.Cursor = Cursors.Default
            Case "Show Factory/Port"
                Dim LastStyle As String = ""
                Dim LastFactory As String = ""
                Dim LastPort As String = ""
                Me.Cursor = Cursors.WaitCursor
                tlb_sbt = DirectCast(tlb.Tools("Calc Price Breaks"), UltraWinToolbars.StateButtonTool)
                For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                    If LastStyle <> rowICTSTYL1.Item("STYLE_CODE").ToString Then
                        LastStyle = rowICTSTYL1.Item("STYLE_CODE").ToString
                        LastFactory = GetVendorData(rowICTSTYL1.Item("VEND_CODE").ToString, "VEND_SUPPLIER_ID")
                        LastPort = GetVendorData(rowICTSTYL1.Item("VEND_CODE").ToString, "PORT_CODE")
                    End If
                    rowICTSTYL1.Item("FACTORY") = LastFactory
                    rowICTSTYL1.Item("PORT") = LastPort
                Next
                grdICTSTYL1.DisplayLayout.Bands(0).Columns("FACTORY").Hidden = False
                grdICTSTYL1.DisplayLayout.Bands(0).Columns("PORT").Hidden = False
                grdICTSTYL1.UpdateData()
                Me.Cursor = Cursors.Default
            Case "Show All Attributes"
                tlb_sbt = DirectCast(tlb.Tools("Show All Attributes"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Dim LastStyle As String = ""
                    Me.Cursor = Cursors.WaitCursor
                    Dim ATTR_CODE_ALL As String = ""
                    For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                        'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                        '    If rowICTSTYL1.Item("STYLE_CODE").ToString = "MTX55902" Then Stop
                        'End If

                        If LastStyle <> rowICTSTYL1.Item("STYLE_CODE").ToString Then
                            LastStyle = rowICTSTYL1.Item("STYLE_CODE").ToString
                            ATTR_CODE_ALL = ""
                            Fill_Records("ICTSTYL3", LastStyle)
                        End If

                        Dim nextI As Integer = 1
                        For Each rowICTSTYL3 As DataRow In dst.Tables("ICTSTYL3").Select()
                            If rowICTSTYL1.Item("ATTR_CODE") <> rowICTSTYL3.Item("ATTR_CODE") Then
                                ATTR_CODE_ALL = ATTR_CODE_ALL + rowICTSTYL3.Item("ATTR_CODE") + " - "
                                If nextI <= 4 Then
                                    nextI += 1
                                    rowICTSTYL1.Item(String.Format("ATTR_CODE{0}", nextI)) = rowICTSTYL3.Item("ATTR_CODE")
                                End If
                            End If
                        Next
                        If ATTR_CODE_ALL.Length >= 3 Then
                            ATTR_CODE_ALL = ATTR_CODE_ALL.Substring(0, ATTR_CODE_ALL.Length - 3)
                        End If
                        rowICTSTYL1.Item("ATTR_CODE_ALL") = ATTR_CODE_ALL
                    Next
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE_ALL").Hidden = False
                    grdICTSTYL1.UpdateData()
                Else
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE_ALL").Hidden = True
                End If
                Me.Cursor = Cursors.Default
            Case "Show Disc Date"
                Me.Cursor = Cursors.WaitCursor
                For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                    Dim row As UltraWinGrid.UltraGridRow = grdICTSTYL1.Rows.GetRowWithListIndex(dst.Tables("ICTSTYL1").Rows.IndexOf(rowICTSTYL1))
                    If Not row.IsFilteredOut Then
                        Dim STYLE_COLOR As String = rowICTSTYL1.Item("STYLE_CODE").ToString & ":" & rowICTSTYL1.Item("COLOR_CODE").ToString
                        Dim SQLS As New StringBuilder With {.Length = 0}
                        SQLS.AppendLine("SELECT MAX(INIT_DATE) AS INIT_DATE")
                        SQLS.AppendLine("FROM ASTAUDT1")
                        SQLS.AppendLine("WHERE TABLE_NAME = 'ICTSTYC1'")
                        SQLS.AppendLine("AND COLUMN_NAME = 'STYLE_COLOR_STATUS'")
                        SQLS.AppendLine("AND NEW_VALUE = 'D'")
                        SQLS.AppendLine("AND KEY_VALUE = '" & STYLE_COLOR & "'")
                        ASCMAIN1.sql = SQLS.ToString()
                        Dim INIT_DATE As String = ASCDATA1.GetDataValue
                        If IsDate(INIT_DATE) Then
                            rowICTSTYL1.Item("DISC_DATE") = INIT_DATE
                        End If
                    End If
                Next
                grdICTSTYL1.DisplayLayout.Bands(0).Columns("DISC_DATE").Hidden = False
                grdICTSTYL1.DisplayLayout.Bands(0).Columns("DISC_DATE").Format = "MM/dd/yy"
                Me.Cursor = Cursors.Default
            Case "Show Extended Pack"
                Dim CartTypes As String() = {"ITM", "CTN", "INR"}
                tlb_sbt = DirectCast(tlb.Tools("Show Extended Pack"), UltraWinToolbars.StateButtonTool)
                For Each COLUMN_NAME As String In New String() {"LENGTH", "WIDTH", "HEIGHT", "WEIGHT"}
                    For Each CartType As String In CartTypes
                        grdICTSTYL1.DisplayLayout.Bands(0).Columns(COLUMN_NAME & "_" & CartType).Hidden = Not tlb_sbt.Checked
                    Next
                Next
                If tlb_sbt.Checked Then
                    Me.Cursor = Cursors.WaitCursor
                    For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                        Dim row As UltraWinGrid.UltraGridRow = grdICTSTYL1.Rows.GetRowWithListIndex(dst.Tables("ICTSTYL1").Rows.IndexOf(rowICTSTYL1))
                        If Not row.IsFilteredOut Then
                            Dim STYLE_COLOR As String = rowICTSTYL1.Item("STYLE_CODE").ToString
                            For Each CartType As String In CartTypes
                                Dim rowICTSTYLD As DataRow = LookUp("ICTSTYLD", New String() {STYLE_COLOR, CartType})
                                If Not IsNothing(rowICTSTYLD) Then
                                    rowICTSTYL1.Item("LENGTH_" & CartType) = Val(rowICTSTYLD.Item("LENGTH").ToString & "")
                                    rowICTSTYL1.Item("WIDTH_" & CartType) = Val(rowICTSTYLD.Item("WIDTH").ToString & "")
                                    rowICTSTYL1.Item("HEIGHT_" & CartType) = Val(rowICTSTYLD.Item("HEIGHT").ToString & "")
                                    rowICTSTYL1.Item("WEIGHT_" & CartType) = Val(rowICTSTYLD.Item("WEIGHT").ToString & "")
                                Else
                                    rowICTSTYL1.Item("LENGTH_" & CartType) = Null
                                    rowICTSTYL1.Item("WIDTH_" & CartType) = Null
                                    rowICTSTYL1.Item("HEIGHT_" & CartType) = Null
                                    rowICTSTYL1.Item("WEIGHT_" & CartType) = Null
                                End If
                            Next
                        End If
                    Next
                End If
            Case "Print Ribbon Sheet"
                Dim SEL_CNT As Integer = dst.Tables.Item("ICTSTYL1").Select("SEL = '1'").Count
                If SEL_CNT = 0 Then
                    MsgBox("You Must Select A Row To Print", vbOKOnly, "Row Selection")
                Else
                    For Each rowSEL As DataRow In dst.Tables.Item("ICTSTYL1").Select("SEL = '1'")
                        Dim STYLE_CODE As String = rowSEL.Item("STYLE_CODE").ToString & String.Empty
                        dst.Tables.Item("RIBBON1").Clear()
                        Dim labelCount As Integer = 14
                        For i As Integer = 1 To labelCount
                            Fill_Records("RIBBON1", STYLE_CODE, False)

                        Next
                        FillExtraFields(labelCount, rowSEL)
                        Print_Report_Begin()

                        Generate_Report("RIBBON1")
                        Print_Report_End()
                    Next
                    'Dim STYLE_CODE As String = grd.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
                    'dst.Tables.Item("RIBBON1").Clear()
                    'Dim labelCount As Integer = 14
                    'For i As Integer = 1 To labelCount
                    '    Fill_Records("RIBBON1", STYLE_CODE, False)

                    'Next
                    'FillExtraFields(labelCount)
                    'Print_Report_Begin()

                    'Generate_Report("RIBBON1")
                    'Print_Report_End()
                End If
            Case "Print Ribbon Combined"
                Dim SEL_CNT As Integer = dst.Tables.Item("ICTSTYL1").Select("SEL = '1'").Count
                If SEL_CNT = 0 Then
                    MsgBox("You Must Select A Row To Print", vbOKOnly, "Row Selection")
                Else
                    dst.Tables.Item("RIBBON1").Clear()
                    For Each rowSEL As DataRow In dst.Tables.Item("ICTSTYL1").Select("SEL = '1'")
                        Dim STYLE_CODE As String = rowSEL.Item("STYLE_CODE").ToString & String.Empty

                        For i As Integer = 1 To 14
                            Fill_Records("RIBBON1", STYLE_CODE, False)
                        Next
                    Next
                    FillExtraFieldsAll()
                    Print_Report_Begin()
                    Generate_Report("RIBBON1")
                    Print_Report_End()
                End If
            Case "Style Masterfile"
                If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    Context_Launch("View", keys, e.Tool.Key, "ICTSTYL1")
                Else
                    MsgBox("This Feature Is Only For Big ABS")
                End If
            Case "Show E-Commerce"
                tlb_sbt = DirectCast(tlb.Tools("Show E-Commerce"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Dim LastStyle As String = ""
                    Me.Cursor = Cursors.WaitCursor
                    Dim SET_MAX As Int64 = 0
                    Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                    Dim ATTR_CODE_ALL As String = ""
                    For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                        If LastStyle <> rowICTSTYL1.Item("STYLE_CODE").ToString Then
                            LastStyle = rowICTSTYL1.Item("STYLE_CODE").ToString
                            ATTR_CODE_ALL = ""
                            Fill_Records("ICTSTYL3", LastStyle)
                        End If
                        Dim nextI As Integer = 1

                        For Each rowICTSTYL3 As DataRow In dst.Tables("ICTSTYL3").Select()
                            If rowICTSTYL1.Item("ATTR_CODE") <> rowICTSTYL3.Item("ATTR_CODE") Then
                                ATTR_CODE_ALL = ATTR_CODE_ALL + rowICTSTYL3.Item("ATTR_CODE") + " - "
                                If nextI <= 4 Then
                                    nextI += 1
                                    rowICTSTYL1.Item(String.Format("ATTR_CODE{0}", nextI)) = rowICTSTYL3.Item("ATTR_CODE")
                                End If
                            End If
                        Next
                        If ATTR_CODE_ALL.Length >= 3 Then
                            ATTR_CODE_ALL = ATTR_CODE_ALL.Substring(0, ATTR_CODE_ALL.Length - 3)
                        End If
                        rowICTSTYL1.Item("ATTR_CODE_ALL") = ATTR_CODE_ALL
                        SQLS.Length = 0
                        SQLS.AppendLine("SELECT COUNT(*)")
                        SQLS.AppendLine("FROM ICTSTYST")
                        SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", LastStyle))
                        ASCMAIN1.sql = SQLS.ToString()
                        Dim SCNT As Int16 = Val(ASCDATA1.GetDataValue)
                        If SCNT > SET_MAX Then
                            SET_MAX = SCNT
                        End If
                    Next
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE_ALL").Hidden = False

                    For i As Int64 = 1 To SET_MAX
                        Dim SETS As New List(Of String)
                        SETS.Add("Set_Depth_" & Format(i, "0#"))
                        SETS.Add("Set_Width_" & Format(i, "0#"))
                        SETS.Add("Set_Height_" & Format(i, "0#"))
                        SETS.Add("SET_Weight_" & Format(i, "0#"))
                        SETS.Add("SET_Length_" & Format(i, "0#"))
                        For Each SCOL As String In SETS
                            If Not grdICTSTYL1.DisplayLayout.Bands(0).Columns.Exists(SCOL) Then
                                grdICTSTYL1.DisplayLayout.Bands(0).Columns.Add(SCOL)
                                grdICTSTYL1.DisplayLayout.Bands(0).Columns(SCOL).Header.Caption = SCOL.Replace("_", " ")
                                grdICTSTYL1.DisplayLayout.Bands(0).Columns(SCOL).CellAppearance.TextHAlign = HAlign.Right
                            End If
                            grdICTSTYL1.DisplayLayout.Bands(0).Columns(SCOL).Hidden = False
                        Next
                    Next

                    For Each grow As UltraWinGrid.UltraGridRow In grdICTSTYL1.Rows
                        For i As Int64 = 1 To SET_MAX
                            Dim STYLE_CODE As String = grow.Cells.Item("STYLE_CODE").Text
                            Dim rowICTSTYST As DataRow = LookUp("ICTSTYST", New String() {STYLE_CODE, i})
                            Dim SET_DEPTH As String = "Set_Depth_" & Format(i, "0#")
                            Dim SET_WIDTH As String = "Set_Width_" & Format(i, "0#")
                            Dim SET_HEIGHT As String = "Set_Height_" & Format(i, "0#")
                            Dim SET_WEIGHT As String = "SET_Weight_" & Format(i, "0#")
                            Dim SET_LENGTH As String = "SET_Length_" & Format(i, "0#")
                            If Not IsNothing(rowICTSTYST) Then
                                grow.Cells(SET_DEPTH).Value = Format(Val(rowICTSTYST.Item("DEPTH").ToString), "###,##0.00")
                                grow.Cells(SET_WIDTH).Value = Format(Val(rowICTSTYST.Item("WIDTH").ToString), "###,##0.00")
                                grow.Cells(SET_HEIGHT).Value = Format(Val(rowICTSTYST.Item("HEIGHT").ToString), "###,##0.00")
                                grow.Cells(SET_WEIGHT).Value = Format(Val(rowICTSTYST.Item("WEIGHT").ToString), "###,##0.00")
                                grow.Cells(SET_LENGTH).Value = Null
                            Else
                                grow.Cells(SET_DEPTH).Value = Null
                                grow.Cells(SET_WIDTH).Value = Null
                                grow.Cells(SET_HEIGHT).Value = Null
                                grow.Cells(SET_WEIGHT).Value = Null
                                grow.Cells(SET_LENGTH).Value = Null
                            End If
                        Next
                    Next

                    grdICTSTYL1.UpdateData()
                Else
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("ATTR_CODE_ALL").Hidden = True
                    For Each grdCol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                        If grdCol.Key.ToString.StartsWith("SET_") Then
                            grdCol.Hidden = True
                        End If
                    Next
                End If
                Me.Cursor = Cursors.Default
            Case "Show Lighting"
                tlb_sbt = DirectCast(tlb.Tools("Show Lighting"), UltraWinToolbars.StateButtonTool)
                Me.Cursor = Cursors.WaitCursor
                If tlb_sbt.Checked Then
                    For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                        Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
                        Dim COLOR_CODE As String = rowICTSTYL1.Item("COLOR_CODE").ToString & String.Empty
                        Dim LIGHT_TYPE As String = ""
                        Dim LIGHT_COLOR As String = ""
                        If STYLE_CODE.Length > 1 Then
                            Select Case STYLE_CODE.Substring(STYLE_CODE.Length - 1, 1)
                                Case "A"
                                    LIGHT_TYPE = "ALWAYS LIT"
                                Case "B"
                                    LIGHT_TYPE = "LIT"
                                Case "E"
                                    LIGHT_TYPE = "EXTRA LIT"
                                Case "G"
                                    LIGHT_TYPE = "CSA LIT"
                                Case "I"
                                    LIGHT_TYPE = "CSA DUAL"
                                Case "J"
                                    LIGHT_TYPE = "CSA LED"
                                Case "L"
                                    LIGHT_TYPE = "LED"
                                Case "M"
                                    LIGHT_TYPE = "LED 5MM"
                                Case "R"
                                    LIGHT_TYPE = "MICRO LED"
                                Case "U"
                                    LIGHT_TYPE = "LED 3MM"
                                Case "V"
                                    LIGHT_TYPE = "CSA LED 3MM"
                                Case "Z"
                                    LIGHT_TYPE = "LED ADJ"
                            End Select
                        End If
                        If STYLE_CODE.Length > 2 Then
                            Select Case STYLE_CODE.Substring(STYLE_CODE.Length - 2, 2)
                                Case "MG"
                                    LIGHT_TYPE = "CSA LED 5MM"
                            End Select
                        End If
                        If COLOR_CODE.Length > 2 Then
                            Select Case COLOR_CODE.Substring(COLOR_CODE.Length - 2, 2)
                                Case "MU", "MB"
                                    LIGHT_COLOR = "MULTI"
                                Case "CB"
                                    LIGHT_COLOR = "CLEAR"
                                Case "DL"
                                    LIGHT_COLOR = "DUAL"
                            End Select
                        End If
                        rowICTSTYL1.Item("LIGHT_TYPE") = LIGHT_TYPE
                        rowICTSTYL1.Item("LIGHT_COLOR") = LIGHT_COLOR
                    Next
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("LIGHT_TYPE").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("LIGHT_COLOR").Hidden = False
                Else
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("LIGHT_TYPE").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("LIGHT_COLOR").Hidden = True
                End If
                Me.Cursor = Cursors.Default
            Case "Show PVC"
                tlb_sbt = DirectCast(tlb.Tools("Show PVC"), UltraWinToolbars.StateButtonTool)
                Me.Cursor = Cursors.WaitCursor
                If tlb_sbt.Checked Then
                    Dim STYLE_LIST As String = BuildStyleInList()
                    Dim sql As New System.Text.StringBuilder With {.Length = 0}
                    sql.AppendLine("SELECT")
                    sql.AppendLine("S1.STYLE_CODE,")
                    sql.AppendLine("PV.HEIGHT AS PVC_HEIGHT,")
                    sql.AppendLine("PV.DIAMETER AS PVC_DIAMETER,")
                    sql.AppendLine("NULL AS PVC_COLORS,")
                    sql.AppendLine("LT.LIGHT_TYPE_DESC AS PVC_LIGHT_TYPE_DESC,")
                    sql.AppendLine("PV.TIP_COUNT AS PVC_TIP_COUNT,")
                    sql.AppendLine("PV.LIGHT_COUNT AS PVC_LIGHT_COUNT")
                    sql.AppendLine("FROM ICTSTYL1 S1, ICTPVC01 PV, ICTPVCLT LT, ICTPVCCG CG, ICTPVCCL CL, ICTPVCTS TS, ICTPVCST ST, ICTPVCLC LC")
                    sql.AppendLine("WHERE S1.STYLE_CLASS_CODE = 'PVC'")
                    sql.AppendLine("AND S1.STYLE_CODE = PV.STYLE_CODE (+)")
                    sql.AppendLine("AND PV.LIGHT_TYPE_CODE = LT.LIGHT_TYPE_CODE (+)")
                    sql.AppendLine("AND PV.COLLECTION_GROUP_CODE = CG.COLLECTION_GROUP_CODE(+)")
                    sql.AppendLine("And PV.COLLECTION_CODE = CL.COLLECTION_CODE (+)")
                    sql.AppendLine("And PV.TREE_SHAPE_CODE = TS.TREE_SHAPE_CODE (+)")
                    sql.AppendLine("And PV.SETUP_CODE = ST.SETUP_CODE (+)")
                    sql.AppendLine("And PV.LIGHT_COLOR_CODE = LC.LIGHT_COLOR_CODE (+)")
                    'sql.AppendLine($"AND S1.STYLE_CODE IN ({STYLE_LIST})")
                    Dim tblPVC As DataTable = ASCDATA1.GetDataTable(sql.ToString())

                    sql.Length = 0
                    sql.AppendLine("SELECT")
                    sql.AppendLine("* FROM ICTSTYC1")
                    'sql.AppendLine($" WHERE STYLE_CODE IN ({STYLE_LIST})")
                    Dim tblICTSTYC1 As DataTable = ASCDATA1.GetDataTable(sql.ToString())

                    For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                        Dim THIS_STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
                        Dim flt As String = $"STYLE_CODE = '{THIS_STYLE_CODE}'"
                        Dim rowPCV As DataRow = tblPVC.Select(flt).FirstOrDefault
                        If Not IsNothing(rowPCV) Then
                            rowICTSTYL1.Item("PVC_HEIGHT") = rowPCV.Item("PVC_HEIGHT").ToString & String.Empty
                            rowICTSTYL1.Item("PVC_DIAMETER") = rowPCV.Item("PVC_DIAMETER").ToString & String.Empty
                            rowICTSTYL1.Item("PVC_LIGHT_TYPE_DESC") = rowPCV.Item("PVC_LIGHT_TYPE_DESC").ToString & String.Empty
                            rowICTSTYL1.Item("PVC_TIP_COUNT") = rowPCV.Item("PVC_TIP_COUNT").ToString & String.Empty
                            rowICTSTYL1.Item("PVC_LIGHT_COUNT") = rowPCV.Item("PVC_LIGHT_COUNT").ToString & String.Empty

                            Dim COLORS As String = ""
                            For Each rowICTSTYC1 As DataRow In tblICTSTYC1.Select(flt, "COLOR_CODE")
                                If rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty <> "" Then
                                    COLORS = COLORS & rowICTSTYC1.Item("COLOR_CODE").ToString & String.Empty & ","
                                End If
                            Next
                            If COLORS.Length = 0 Then
                                COLORS = "No Colors Found"
                            Else
                                COLORS = COLORS.Substring(0, COLORS.Length - 1)
                            End If
                            rowICTSTYL1.Item("PVC_COLORS") = COLORS
                        End If
                    Next
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_HEIGHT").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_DIAMETER").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_LIGHT_TYPE_DESC").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_TIP_COUNT").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_LIGHT_COUNT").Hidden = False
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_COLORS").Hidden = False
                Else
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_HEIGHT").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_DIAMETER").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_LIGHT_TYPE_DESC").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_TIP_COUNT").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_LIGHT_COUNT").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PVC_COLORS").Hidden = True
                End If
            Case "Show Purch Notes"
                tlb_sbt = DirectCast(tlb.Tools("Show Purch Notes"), UltraWinToolbars.StateButtonTool)
                Me.Cursor = Cursors.WaitCursor
                If tlb_sbt.Checked Then
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PURCH_NOTES").Hidden = False
                Else
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PURCH_NOTES").Hidden = True
                End If
                Me.Cursor = Cursors.Default
            Case "Show PO Info"
                tlb_sbt = DirectCast(tlb.Tools("Show PO Info"), UltraWinToolbars.StateButtonTool)
                Me.Cursor = Cursors.WaitCursor
                If tlb_sbt.Checked Then
                    Dim password As String
                    password = InputBox("Enter password:", "Password Required")
                    If password = "Holiday!" Then
                        ASCMAIN1.Progress("Fetching PO Data", "")
                        For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
                            Dim THIS_STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
                            Dim rowICTSTYLX As DataRow = LookUp("ICTSTYL1", THIS_STYLE_CODE)
                            Dim COUNTRY_CODE As String = rowICTSTYL1.Item("COUNTRY_CODE").ToString & String.Empty

                            Dim SQL As New System.Text.StringBuilder With {.Length = 0}
                            SQL.AppendLine("SELECT NVL(MAX(TARIFF_PCT),0) AS TARIFF_PCT")
                            SQL.AppendLine("FROM ICTTARF2")
                            SQL.AppendLine("WHERE TARIFF_START = ")
                            SQL.AppendLine("(")
                            SQL.AppendLine("    SELECT MAX(TARIFF_START)")
                            SQL.AppendLine("    FROM ICTTARF2 ")
                            SQL.AppendLine("    WHERE NVL(TARIFF_START,'01-JAN-1900')<> '01-JAN-1900'")
                            SQL.AppendLine($"    AND COUNTRY_CODE = '{COUNTRY_CODE}'")
                            SQL.AppendLine(")")
                            ASCMAIN1.sql = SQL.ToString()
                            rowICTSTYL1.Item("TARIFF_PCT") = Val(ASCDATA1.GetDataValue)

                            SQL.Length = 0
                            SQL.AppendLine("SELECT PO_COST ")
                            SQL.AppendLine("FROM ")
                            SQL.AppendLine("(")
                            SQL.AppendLine("    SELECT")
                            SQL.AppendLine("    STYLE_CODE,")
                            SQL.AppendLine("    NEW_PO_COST_DATE AS PO_COST_DATE,")
                            SQL.AppendLine("    NEW_PO_COST AS PO_COST")
                            SQL.AppendLine("    FROM ICTSTYV1")
                            SQL.AppendLine("    WHERE NVL(NEW_PO_COST_DATE,'01-JAN-1900') <> '01-JAN-1900'")
                            SQL.AppendLine("    AND NVL(NEW_PO_COST_DATE,'01-JAN-1900') <= SYSDATE")
                            SQL.AppendLine("    UNION")
                            SQL.AppendLine("    SELECT ")
                            SQL.AppendLine("    STYLE_CODE,")
                            SQL.AppendLine("    PO_COST_DATE AS PO_COST_DATE,")
                            SQL.AppendLine("    PO_COST")
                            SQL.AppendLine("    FROM ICTSTYV1")
                            SQL.AppendLine("    WHERE NVL(PO_COST_DATE,'01-JAN-1900') <> '01-JAN-1900'")
                            SQL.AppendLine("    AND NVL(PO_COST_DATE,'01-JAN-1900') <= SYSDATE")
                            SQL.AppendLine(")")
                            SQL.AppendLine($"WHERE STYLE_CODE = '{THIS_STYLE_CODE}'")
                            SQL.AppendLine("AND PO_COST_DATE = ")
                            SQL.AppendLine("(")
                            SQL.AppendLine("    SELECT MAX(PO_COST_DATE) FROM")
                            SQL.AppendLine("    (")
                            SQL.AppendLine("        SELECT")
                            SQL.AppendLine("        STYLE_CODE,")
                            SQL.AppendLine("        NEW_PO_COST_DATE AS PO_COST_DATE,")
                            SQL.AppendLine("        NEW_PO_COST AS PO_COST")
                            SQL.AppendLine("        FROM ICTSTYV1")
                            SQL.AppendLine("        WHERE NVL(NEW_PO_COST_DATE,'01-JAN-1900') <> '01-JAN-1900'")
                            SQL.AppendLine("        AND NVL(NEW_PO_COST_DATE,'01-JAN-1900') <= SYSDATE")
                            SQL.AppendLine("        UNION")
                            SQL.AppendLine("        SELECT ")
                            SQL.AppendLine("        STYLE_CODE,")
                            SQL.AppendLine("        PO_COST_DATE AS PO_COST_DATE,")
                            SQL.AppendLine("        PO_COST")
                            SQL.AppendLine("        FROM ICTSTYV1")
                            SQL.AppendLine("        WHERE NVL(PO_COST_DATE,'01-JAN-1900') <> '01-JAN-1900'")
                            SQL.AppendLine("        AND NVL(PO_COST_DATE,'01-JAN-1900') <= SYSDATE")
                            SQL.AppendLine("    )")
                            SQL.AppendLine($"    WHERE STYLE_CODE = '{THIS_STYLE_CODE}'")
                            SQL.AppendLine(")")
                            ASCMAIN1.sql = SQL.ToString()
                            rowICTSTYL1.Item("PO_COST") = Val(ASCDATA1.GetDataValue)

                            rowICTSTYL1.Item("STYLE_PO_QTY_MIN") = Val(rowICTSTYLX.Item("STYLE_PO_QTY_MIN").ToString & String.Empty)
                        Next
                        grdICTSTYL1.DisplayLayout.Bands(0).Columns("TARIFF_PCT").Hidden = False
                        grdICTSTYL1.DisplayLayout.Bands(0).Columns("PO_COST").Hidden = False
                        grdICTSTYL1.DisplayLayout.Bands(0).Columns("STYLE_PO_QTY_MIN").Hidden = False
                        ASCMAIN1.Progress("", "")
                    Else
                        grdICTSTYL1.DisplayLayout.Bands(0).Columns("TARIFF_PCT").Hidden = True
                        grdICTSTYL1.DisplayLayout.Bands(0).Columns("PO_COST").Hidden = True
                        grdICTSTYL1.DisplayLayout.Bands(0).Columns("STYLE_PO_QTY_MIN").Hidden = True
                        MsgBox("Access Denied!", MsgBoxStyle.Critical, "Error")
                    End If
                Else
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("TARIFF_PCT").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("PO_COST").Hidden = True
                    grdICTSTYL1.DisplayLayout.Bands(0).Columns("STYLE_PO_QTY_MIN").Hidden = True
                End If
                Me.Cursor = Cursors.Default
            Case "Dump Data To Excel"
                Dim oWB As SpreadsheetGear.IWorkbook
                Dim oSheet As SpreadsheetGear.IWorksheet = Nothing

                Dim xls_file_name As String = "Attributes.xlsx"
                Dim fDialog As New FolderBrowserDialog
                fDialog.Description = "Please Select The Folder To Save File"
                fDialog.ShowDialog()
                xls_file_name = $"{fDialog.SelectedPath}\{xls_file_name}"
                If System.IO.File.Exists(xls_file_name) Then
                    Dim iResult As MsgBoxResult = MsgBox("Delete it?", vbYesNo, $"{xls_file_name} Exists!")
                    If iResult <> vbYes Then
                        Exit Sub
                    Else
                        System.IO.File.Delete(xls_file_name)
                    End If
                End If

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Creating File...", "")
                Application.DoEvents()

                oWB = SpreadsheetGear.Factory.GetWorkbook()
                For i As Integer = oWB.Worksheets.Count To 2 Step -1
                    oWB.Worksheets(i).Delete()
                Next i
                oSheet = oWB.Worksheets.Add()
                oSheet.Name = "Search Results"

                If oWB.Worksheets.Count = 2 Then
                    oWB.Worksheets(0).Delete()
                End If

                ASCMAIN1.Progress("-", oSheet.Name)

                Load_DataTable_into_SGXLS(1, 1, dst.Tables.Item("ICTSTYL1"), oSheet, Nothing, Nothing, "", "")

                'Set heading to Grid Heading
                Dim colCnt As Int64 = dst.Tables.Item("ICTSTYL1").Columns.Count
                Dim rowCnt As Int64 = dst.Tables.Item("ICTSTYL1").Rows.Count
                For i As Int64 = 0 To colCnt
                    oSheet.Range(0, i).Select()
                    Dim colTitle As String = oSheet.Range(0, i).Text
                    If grdICTSTYL1.DisplayLayout.Bands(0).Columns.IndexOf(colTitle) <> -1 Then
                        Dim grdTitle As String = grdICTSTYL1.DisplayLayout.Bands(0).Columns(colTitle).Header.Caption
                        oSheet.Range(0, i).Value = grdTitle
                    End If
                Next

                'Set Columns to Auto
                oSheet.Range($"A1:ZZ1").EntireColumn.AutoFit()

                'Set Rows to 15
                oSheet.Range($"A1:A{rowCnt + 10}").RowHeight = 15

                oSheet.Range("A1:A1").Select()
                oSheet.WindowInfo.FreezePanes = True
                oSheet.Range("A1:A1").Select()

                oWB.Worksheets(0).Select()

                oWB.SaveAs(xls_file_name, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                'Show_Document(xls_file_name)
                oWB = Nothing

                Cursor = Cursors.Default
                ASCMAIN1.Progress("", "")
                MsgBox("File Created", vbOKOnly, "Done")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

    Private Function BuildStyleInList() As String
        Dim RETVAL As String = ""
        Dim STYLE_LIST As New List(Of String)
        For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString.Trim & String.Empty
            If Not STYLE_LIST.Contains(STYLE_CODE) Then
                STYLE_LIST.Add(STYLE_CODE)
            End If
        Next
        For Each STYLE As String In STYLE_LIST
            RETVAL &= $"','{STYLE}"
        Next
        If RETVAL.Length > 3 Then
            RETVAL = RETVAL.Substring(2, RETVAL.Length - 2) & "'"
        Else
            RETVAL = "'XXX'"
        End If
        Return RETVAL
    End Function
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "XXXXXX"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
        End Select

        If hotKeyPartOne <> "" Then
            Hot_Key_Part_Two(hotKeyPartOne, e)
        Else
            Hot_Key_Part_One(e)
        End If
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                'FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
            Select Case Absx1.GetABSColumnName(sender)
                Case "STYLE_CLASS_CODE"
                    'Dim STYLE_CLASS_CODE As String = Absx1.txtFor("STYLE_CLASS_CODE").Text
                    'Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)
                    Initialize_tabAttributes()
            End Select
        End If
    End Sub
#End Region

#Region "Custom Methods"

    Private Sub CloseSplash()
        If progressSplash Is Nothing OrElse progressSplash.IsHandleCreated = False Then
            Return
        End If
        progressSplash.Invoke(New EventHandler(AddressOf progressSplash.EndForm))
        progressSplash.Dispose()
        progressSplash = Nothing
    End Sub

    Sub DisplaySplash(st1 As String, st2 As String, st3 As String)

        progressSplashMsg1 = st1
        progressSplashMsg2 = st2
        progressSplashMsg3 = st3

        Dim splashThread As Threading.Thread = New Threading.Thread(New Threading.ThreadStart(AddressOf StartSplash))

        splashThread.Start()

    End Sub

    Private Sub ExcelProcessInit()
        Try
            'Get all currently running process Ids for Excel applications
            mExcelProcesses = Process.GetProcessesByName("Excel")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ExcelProcessKill()
        Dim oProcesses() As Process
        Dim bFound As Boolean

        Try
            'Get all currently running process Ids for Excel applications
            oProcesses = Process.GetProcessesByName("Excel")

            If oProcesses.Length > 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    bFound = False

                    For j As Integer = 0 To mExcelProcesses.Length - 1
                        If oProcesses(i).Id = mExcelProcesses(j).Id Then
                            bFound = True
                            Exit For
                        End If
                    Next

                    If Not bFound Then
                        oProcesses(i).Kill()
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FillExtraFields(ByVal labelCount As Integer, ByRef rowSEL As DataRow)
        'Note: Stolen form SOCMAIN2 - 
        For i As Integer = 0 To labelCount - 1
            Dim rowRIBBON1 As DataRow = dst.Tables("RIBBON1").Rows(i)
            Dim STYLE_CODE As String = rowRIBBON1.Item("STYLE_CODE")
            Dim CARTON_PACK_QTY As Integer = Val(rowRIBBON1.Item("CARTON_PACK_QTY"))
            Dim INNER_PACK_QTY As Integer = Val(rowRIBBON1.Item("INNER_PACK_QTY"))
            Dim SUB_UNIT_PACK_QTY As Integer = Val(rowRIBBON1.Item("SUB_UNIT_PACK_QTY") & "")
            Dim STYLE_UOM As String = rowRIBBON1.Item("STYLE_UOM")
            '(rowRIBBON1.Item("STYLE_CODE"))
            rowRIBBON1.Item("DATEPRINTED") = String.Format("{0}/{1}/{2}", Now.Date.Month, Now.Day, Now.Year.ToString.Substring(2, 2))
            rowRIBBON1.Item("BOXQTY") = String.Format("BOX:{0}", (INNER_PACK_QTY * SUB_UNIT_PACK_QTY))
            rowRIBBON1.Item("CARTQTY") = String.Format("CART:{0}", CARTON_PACK_QTY)
            rowRIBBON1.Item("COLORS") = "Add Colors Here"
            rowRIBBON1.Item("COLORS3") = ""
            rowRIBBON1.Item("Price1_LBL") = rowSEL.Item("PBH1").ToString & String.Empty
            If rowSEL.Item("PBV1").ToString & String.Empty <> String.Empty Then
                rowRIBBON1.Item("Price1_AMT") = rowSEL.Item("PBV1").ToString & String.Empty
            End If
            rowRIBBON1.Item("Price2_LBL") = rowSEL.Item("PBH2").ToString & String.Empty
            If rowSEL.Item("PBV2").ToString & String.Empty <> String.Empty Then
                rowRIBBON1.Item("Price2_AMT") = rowSEL.Item("PBV2").ToString & String.Empty
            End If
            rowRIBBON1.Item("Price3_LBL") = rowSEL.Item("PBH3").ToString & String.Empty
            If rowSEL.Item("PBV3").ToString & String.Empty <> String.Empty Then
                rowRIBBON1.Item("Price3_AMT") = rowSEL.Item("PBV3").ToString & String.Empty
            End If
            rowRIBBON1.Item("Price4_LBL") = rowSEL.Item("PBH4").ToString & String.Empty
            rowRIBBON1.Item("Price4_AMT") = Val(rowSEL.Item("PBV4").ToString & String.Empty)

            Dim COLORS As String = ""
            Dim sql As New System.Text.StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT ICTCOLR1.COLOR_CODE")
            sql.AppendLine("FROM ICTSTYC1, ICTCOLR1")
            sql.AppendLine("WHERE ICTSTYC1.COLOR_CODE = ICTCOLR1.COLOR_CODE")
            sql.AppendLine("AND STYLE_CODE = :PARM1")
            Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
            For Each rowICTCOLOR1 As DataRow In tbl.Rows
                COLORS = COLORS & rowICTCOLOR1.Item(0).ToString & ", "
            Next
            If COLORS.Length >= 2 Then
                COLORS = COLORS.Substring(0, COLORS.Length - 2)
            End If
            rowRIBBON1.Item("COLORSDESC") = COLORS
            rowRIBBON1.Item("VEND_SUPPLIER_ID") = ""
            rowRIBBON1.Item("PORT_CODE_ORIG") = ""
        Next
    End Sub

    Private Sub FillExtraFieldsAll()
        For Each rowRIBBON1 As DataRow In dst.Tables("RIBBON1").Select()
            Dim STYLE_CODE As String = rowRIBBON1.Item("STYLE_CODE")
            Dim rowSEL As DataRow = dst.Tables.Item("ICTSTYL1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE), "").FirstOrDefault
            If Not IsNothing(rowSEL) Then
                Dim CARTON_PACK_QTY As Integer = Val(rowRIBBON1.Item("CARTON_PACK_QTY"))
                Dim INNER_PACK_QTY As Integer = Val(rowRIBBON1.Item("INNER_PACK_QTY"))
                Dim SUB_UNIT_PACK_QTY As Integer = Val(rowRIBBON1.Item("SUB_UNIT_PACK_QTY") & "")
                Dim STYLE_UOM As String = rowRIBBON1.Item("STYLE_UOM")
                rowRIBBON1.Item("DATEPRINTED") = String.Format("{0}/{1}/{2}", Now.Date.Month, Now.Day, Now.Year.ToString.Substring(2, 2))
                rowRIBBON1.Item("BOXQTY") = String.Format("BOX:{0}", (INNER_PACK_QTY * SUB_UNIT_PACK_QTY))
                rowRIBBON1.Item("CARTQTY") = String.Format("CART:{0}", CARTON_PACK_QTY)
                rowRIBBON1.Item("COLORS") = "Add Colors Here"
                rowRIBBON1.Item("COLORS3") = ""
                rowRIBBON1.Item("Price1_LBL") = rowSEL.Item("PBH1").ToString & String.Empty
                If rowSEL.Item("PBV1").ToString & String.Empty <> String.Empty Then
                    rowRIBBON1.Item("Price1_AMT") = rowSEL.Item("PBV1").ToString & String.Empty
                End If
                rowRIBBON1.Item("Price2_LBL") = rowSEL.Item("PBH2").ToString & String.Empty
                If rowSEL.Item("PBV2").ToString & String.Empty <> String.Empty Then
                    rowRIBBON1.Item("Price2_AMT") = rowSEL.Item("PBV2").ToString & String.Empty
                End If
                rowRIBBON1.Item("Price3_LBL") = rowSEL.Item("PBH3").ToString & String.Empty
                If rowSEL.Item("PBV3").ToString & String.Empty <> String.Empty Then
                    rowRIBBON1.Item("Price3_AMT") = rowSEL.Item("PBV3").ToString & String.Empty
                End If
                rowRIBBON1.Item("Price4_LBL") = rowSEL.Item("PBH4").ToString & String.Empty
                rowRIBBON1.Item("Price4_AMT") = Val(rowSEL.Item("PBV4").ToString & String.Empty)

                Dim COLORS As String = ""
                Dim sql As New System.Text.StringBuilder With {.Length = 0}
                sql.AppendLine("SELECT ICTCOLR1.COLOR_CODE")
                sql.AppendLine("FROM ICTSTYC1, ICTCOLR1")
                sql.AppendLine("WHERE ICTSTYC1.COLOR_CODE = ICTCOLR1.COLOR_CODE")
                sql.AppendLine("AND STYLE_CODE = :PARM1")
                Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
                For Each rowICTCOLOR1 As DataRow In tbl.Rows
                    COLORS = COLORS & rowICTCOLOR1.Item(0).ToString & ", "
                Next
                If COLORS.Length >= 2 Then
                    COLORS = COLORS.Substring(0, COLORS.Length - 2)
                End If
                rowRIBBON1.Item("COLORSDESC") = COLORS
                rowRIBBON1.Item("VEND_SUPPLIER_ID") = ""
                rowRIBBON1.Item("PORT_CODE_ORIG") = ""
            Else
                ASCMAIN1.Progress("Problem With Style Masterfile", STYLE_CODE)
            End If
        Next
    End Sub

    Sub Find_Styles(Optional ByVal FetchFromOrder As Boolean = False,
                    Optional AllDiscontinued As Boolean = False,
                    Optional ByVal Load_from_Excel As Boolean = False,
                    Optional Style_List As List(Of String) = Nothing,
                    Optional Color_List As List(Of String) = Nothing,
                    Optional AllActive As Boolean = False,
                    Optional PAGE_CODE As String = "")
        Dim QD As String = ""
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        'STYLE_CLASS_CODE = Absx1.txtFor("STYLE_CLASS_CODE").Text

        Dim SQLAttribute As New StringBuilder With {.Length = 0}
        SQLAttribute.AppendLine("SELECT")
        SQLAttribute.AppendLine("S3.STYLE_CODE,")
        SQLAttribute.AppendLine("MIN(S3.ATTR_CODE) AS ATTR_CODE")
        SQLAttribute.AppendLine("FROM ICVSTYL3 S3")
        SQLAttribute.AppendLine("WHERE (S3.STYLE_CODE, NVL(S3.ATT_RANK,99))")
        SQLAttribute.AppendLine("IN")
        SQLAttribute.AppendLine("(")
        SQLAttribute.AppendLine("  SELECT")
        SQLAttribute.AppendLine("  SI.STYLE_CODE,")
        SQLAttribute.AppendLine("  MIN(NVL(SI.ATT_RANK,99)) AS ATT_RANK")
        SQLAttribute.AppendLine("  FROM ICVSTYL3 SI")
        If ATTR_CODE_1s <> "" Then
            SQLAttribute.AppendLine(" where ATTR_CODE In (" & Mid(ATTR_CODE_1s, 2) & ")")
        End If
        SQLAttribute.AppendLine("  GROUP BY SI.STYLE_CODE")
        SQLAttribute.AppendLine(")")
        SQLAttribute.AppendLine("GROUP BY S3.STYLE_CODE")
        ASCMAIN1.sql = SQLAttribute.ToString

        Dim EXCEL_LIST As String = ASCMAIN1.Temp_Table

        If Load_from_Excel Then

            ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1" & vbCrLf
            Dim sqlw As String = ""
            S.Length = 0
            S.AppendLine("SELECT STYLE_CODE, COLOR_CODE, 999 AS IMPORT_SORT")
            S.AppendLine("FROM ICTSTYC1")
            S.AppendLine("WHERE STYLE_CODE = NULL")
            ASCMAIN1.sql = S.ToString
            EXCEL_LIST = ASCMAIN1.Temp_Table
            For i As Integer = 0 To Style_List.Count - 1
                Dim S2 As New System.Text.StringBuilder With {.Length = 0}
                S2.AppendLine(String.Format("INSERT INTO {0}", EXCEL_LIST))
                S2.AppendLine(String.Format("VALUES ('{0}','{1}', {2})", Style_List(i), Color_List(i), i))
                ASCMAIN1.sql = S2.ToString
                ASCDATA1.ExecuteSQL()
            Next

            'ASCMAIN1.sql = "SELECT X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, NVL(ICTSTYC1.THEME_CODE,'') AS THEME_CODE from ICTSTYC1, (" & vbCrLf _
            '    & "Select " & sqlcols & " from ICTSTYL1" & vbCrLf _
            '    & ") X," & vbCrLf _
            '    & "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
            '    & "(SELECT S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE" & vbCrLf _
            '    & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
            '    & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
            '    & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
            '    & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
            '    & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
            '    & " from ICTSTAT2 S2  , ICTCOLR1 C1  WHERE S2.COLOR_CODE = C1.COLOR_CODE GROUP BY S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, C1.COLOR_GROUP_CODE) Z" & vbCrLf _
            '    & " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
            '    & "   and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
            '    & "   and ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
            '    & "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
            '    & "and (x.style_code, z.color_code) in (SELECT STYLE_CODE, COLOR_CODE from " & EXCEL_LIST & ")"
            ASCMAIN1.sql = "SELECT X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, NVL(ICTSTYC1.THEME_CODE,'') AS THEME_CODE from ICTSTYC1, (" & vbCrLf _
                & "Select " & sqlcols & " from ICTSTYL1" & vbCrLf _
                & ") X," & vbCrLf _
                & "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
                & "(SELECT SC.STYLE_CODE, SC.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
                & " from ICTSTYC1 SC, ICTSTAT2 S2, ICTCOLR1 C1" & vbCrLf _
                & " WHERE SC.STYLE_CODE = S2.STYLE_CODE (+)" & vbCrLf _
                & " AND SC.COLOR_CODE = S2.COLOR_CODE (+)" & vbCrLf _
                & " AND SC.COLOR_CODE = C1.COLOR_CODE" & vbCrLf _
                & " GROUP BY SC.STYLE_CODE, SC.COLOR_CODE, C1.COLOR_CODE_LONG, C1.COLOR_GROUP_CODE" & vbCrLf _
                & " ) Z" & vbCrLf _
                & " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
                & "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "and (x.style_code, z.color_code) in (SELECT STYLE_CODE, COLOR_CODE from " & EXCEL_LIST & ")"
        Else
            If FetchFromOrder Then
                Dim SQLIN As String = ""
                For Each ORDR As String In ORDR_NOs
                    SQLIN = SQLIN & "'" & ORDR & "',"
                Next
                SQLIN = SQLIN.Substring(0, SQLIN.Length - 1)

                'ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1" & vbCrLf
                'Dim sqlw As String = ""
                'ASCMAIN1.sql = "SELECT X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, NVL(ICTSTYC1.THEME_CODE,'') AS THEME_CODE from ICTSTYC1, (" & vbCrLf _
                '& ASCMAIN1.sql _
                '& ") X," & vbCrLf _
                '& "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
                '& "(SELECT S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE" & vbCrLf _
                '& ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
                '& ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
                '& ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
                '& ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
                '& ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
                '& " from ICTSTAT2 S2  , ICTCOLR1 C1  WHERE S2.COLOR_CODE = C1.COLOR_CODE GROUP BY S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, C1.COLOR_GROUP_CODE) Z" & vbCrLf _
                '& " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                '& "   and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
                '& "   and ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
                '& "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                '& "and (x.style_code, z.color_code) in (SELECT STYLE_CODE, COLOR_CODE from SOTORDR2 where ordr_no in (" & SQLIN & "))"


                ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1" & vbCrLf
                Dim sqlw As String = ""
                ASCMAIN1.sql = "SELECT X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, NVL(ICTSTYC1.THEME_CODE,'') AS THEME_CODE from ICTSTYC1, (" & vbCrLf _
                & ASCMAIN1.sql _
                & ") X," & vbCrLf _
                & "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
                & "(SELECT SC.STYLE_CODE,SC.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'MS',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'MS',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'MS',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'MS',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
                & ", SUM (DECODE(S2.WHSE_CODE,'MS',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
                & " From ICTSTYC1 SC, ICTCOLR1 C1, ICTSTAT2 S2" & vbCrLf _
                & " WHERE SC.COLOR_CODE = C1.COLOR_CODE" & vbCrLf _
                & " And sc.STYLE_CODE = s2.STYLE_CODE(+)" & vbCrLf _
                & " And SC.COLOR_CODE = S2.COLOR_CODE (+)" & vbCrLf _
                & " Group By sc.STYLE_CODE, sc.COLOR_CODE, c1.COLOR_CODE_LONG, c1.COLOR_GROUP_CODE) Z" & vbCrLf _
                & " Where y.STYLE_CODE(+) = x.STYLE_CODE" & vbCrLf _
                & "   And ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
                & "   And ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
                & "   And Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "And (x.style_code, z.color_code) In (Select STYLE_CODE, COLOR_CODE From SOTORDR2 Where ordr_no In (" & SQLIN & "))"
            Else
                If AllDiscontinued Or AllActive Then
                    If AllDiscontinued Then
                        ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1" & vbCrLf
                        Dim sqlw As String = ""
                        ASCMAIN1.sql = "Select X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, ICTSTYC1.THEME_CODE from ICTSTYC1, (" & vbCrLf _
                        & ASCMAIN1.sql _
                        & ") X," & vbCrLf _
                        & "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
                        & "(SELECT S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
                        & " from ICTSTAT2 S2  , ICTCOLR1 C1  WHERE S2.COLOR_CODE = C1.COLOR_CODE GROUP BY S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, C1.COLOR_GROUP_CODE) Z" & vbCrLf _
                        & " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
                        & "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "and x.style_status in ('D','N')"
                        '& "AND NVL(Z.COLOR_CODE,'null') <> 'null'" & vbCrLf _
                        '& "AND NVL(Y.ATTR_CODE,'null') <> 'null'"
                    Else
                        ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1" & vbCrLf
                        Dim sqlw As String = ""
                        ASCMAIN1.sql = "Select X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, ICTSTYC1.THEME_CODE from ICTSTYC1, (" & vbCrLf _
                        & ASCMAIN1.sql _
                        & ") X," & vbCrLf _
                        & "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
                        & "(SELECT S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
                        & " from ICTSTAT2 S2  , ICTCOLR1 C1  WHERE S2.COLOR_CODE = C1.COLOR_CODE GROUP BY S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, C1.COLOR_GROUP_CODE) Z" & vbCrLf _
                        & " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
                        & "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "and x.style_status in ('A')"
                        '& "AND NVL(Z.COLOR_CODE,'null') <> 'null'" & vbCrLf _
                        '& "AND NVL(Y.ATTR_CODE,'null') <> 'null'"
                    End If
                Else
                    If PAGE_CODE.Length > 0 Then
                        ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1" & vbCrLf
                        Dim sqlw As String = ""
                        ASCMAIN1.sql = "Select X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, ICTSTYC1.THEME_CODE from ICTSTYC1, (" & vbCrLf _
                        & ASCMAIN1.sql _
                        & ") X," & vbCrLf _
                        & "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
                        & "(SELECT S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
                        & " from ICTSTAT2 S2  , ICTCOLR1 C1  WHERE S2.COLOR_CODE = C1.COLOR_CODE GROUP BY S2.STYLE_CODE, S2.COLOR_CODE, C1.COLOR_CODE_LONG, C1.COLOR_GROUP_CODE) Z" & vbCrLf _
                        & " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
                        & "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "and (X.STYLE_CODE, Z.COLOR_CODE) IN" & vbCrLf _
                        & $"(SELECT STYLE_CODE, COLOR_CODE FROM WBTCATED WHERE PAGE_CODE = '{PAGE_CODE}')" & vbCrLf
                        '& "AND NVL(Z.COLOR_CODE,'null') <> 'null'" & vbCrLf _
                        '& "AND NVL(Y.ATTR_CODE,'null') <> 'null'"
                    Else
                        'If SCCs(0) = "ALL" Then
                        '    ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1 where STYLE_CLASS_CODE <> 'ALL'" & vbCrLf
                        'Else
                        ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1 where STYLE_CLASS_CODE IN (" & SCC_IN & ")" & vbCrLf
                        'End If

                        If Not chkSTYLE_STATUS_A.Checked Or Not chkSTYLE_STATUS_N.Checked Or Not chkSTYLE_STATUS_D.Checked Then
                            Dim STATUS_CODEs As String = ""
                            If chkSTYLE_STATUS_A.Checked Then STATUS_CODEs &= ",'A'"
                            If chkSTYLE_STATUS_N.Checked Then STATUS_CODEs &= ",'N'"
                            If chkSTYLE_STATUS_D.Checked Then STATUS_CODEs &= ",'D'"
                            If STATUS_CODEs = "" Then
                                MsgBox("Cannot Find Style if No Status is Selected", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                            Else
                                ASCMAIN1.sql &= " and STYLE_STATUS in (" & Mid(STATUS_CODEs, 2) & ")" & vbCrLf
                            End If
                        Else
                        End If

                        If optSN.Value = "S" Then
                            ASCMAIN1.sql &= " and NVL(CUST_CODE,'NULL') = 'NULL'"
                            QD &= " ;" & "Stock Only"
                        ElseIf optSN.Value = "N" Then
                            ASCMAIN1.sql &= " and NVL(CUST_CODE,'NULL') <> 'NULL'"
                            QD &= " ;" & "Non-Stock Only"
                        End If

                        Dim sqlw As String = ""
                        For I As Integer = 1 To 3
                            If I = 1 Then
                                ATTR_CODE_1s = Get_CODE_VALUEs(grdICTATTR1_1)
                                If ATTR_CODE_1s <> "" Then
                                    QD &= " ;" & "Any of " & Replace(Mid(ATTR_CODE_1s, 2), "'", "")
                                    ASCMAIN1.sql &= " and STYLE_CODE in (Select Distinct STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(ATTR_CODE_1s, 2) & "))" & vbCrLf
                                End If
                            End If
                            If I = 2 Then
                                SIZE_CODEs = Get_CODE_VALUEs(grdICTSIZE1)
                                If SIZE_CODEs <> "" Then
                                    QD &= " ;" & "Sizes: " & Replace(Mid(SIZE_CODEs, 2), "'", "")
                                    ASCMAIN1.sql &= " and NVL(SIZE_CODE,'?') in (" & Mid(SIZE_CODEs, 2) & ")" & vbCrLf
                                End If
                            End If
                            If I = 3 Then
                                ATTR_CODE_2s = Get_CODE_VALUEs(grdICTATTR1_2)
                                If ATTR_CODE_2s <> "" Then
                                    QD &= " ;" & "And also any of " & Replace(Mid(ATTR_CODE_2s, 2), "'", "")
                                    ASCMAIN1.sql &= " and STYLE_CODE in (Select Distinct STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(ATTR_CODE_2s, 2) & "))" & vbCrLf
                                End If
                            End If
                        Next

                        '            & ", SUM (DECODE(WHSE_CODE,'SW',WHSE_QTY_ON_HAND,0)) OHSW, SUM (DECODE(WHSE_CODE,'SW',WHSE_QTY_ON_ORDER,0)) POSW, SUM (DECODE(WHSE_CODE,'SW',WHSE_QTY_TRAN,0)) PSSW" & vbCrLf _

                        ASCMAIN1.sql = "SELECT X.*, NVL(Y.ATTR_CODE,'NONE') AS ATTR_CODE,ICTSTYC1.UPC_CODE,ICTSTYC1.STYLE_COLOR_STATUS, Z.COLOR_CODE, Z.ONH, Z.ONPO, Z.OPEN, Z.TRAN, Z.PICK, Z.COLOR_CODE_LONG AS LONG_COLOR, NVL(Z.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, Z.THEME_CODE from ICTSTYC1, (" & vbCrLf _
                        & ASCMAIN1.sql _
                        & ") X," & vbCrLf _
                        & "(" & SQLAttribute.ToString & ") Y," & vbCrLf _
                        & "(SELECT CL.STYLE_CODE, CL.COLOR_CODE, C1.COLOR_CODE_LONG, NVL(C1.COLOR_GROUP_CODE,'') AS COLOR_GROUP_CODE, NVL(CL.THEME_CODE,'') AS THEME_CODE" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_HAND,0)) ONH" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_ON_ORDER,0)) ONPO" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_TRAN,0)) TRAN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_OPEN,0)) OPEN" & vbCrLf _
                        & ", SUM (DECODE(S2.WHSE_CODE,'" & WHSE_CODE & "',S2.WHSE_QTY_PICK,0)) PICK" & vbCrLf _
                        & " from ICTSTAT2 S2, ICTCOLR1 C1, ICTSTYC1 CL  " & vbCrLf _
                        & " WHERE CL.COLOR_CODE = C1.COLOR_CODE " & vbCrLf _
                        & " AND S2.STYLE_CODE (+) = CL.STYLE_CODE" & vbCrLf _
                        & " AND S2.COLOR_CODE (+) = CL.COLOR_CODE" & vbCrLf _
                        & " GROUP BY CL.STYLE_CODE, CL.COLOR_CODE, C1.COLOR_CODE_LONG, C1.COLOR_GROUP_CODE, CL.THEME_CODE) Z" & vbCrLf _
                        & " where Y.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.STYLE_CODE (+) = Z.STYLE_CODE" & vbCrLf _
                        & "   and ICTSTYC1.COLOR_CODE (+) = Z.COLOR_CODE" & vbCrLf _
                        & "   and Z.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf

                        If Not chkSTYLE_COLOR_STATUS_A.Checked Or Not chkSTYLE_COLOR_STATUS_N.Checked Or Not chkSTYLE_COLOR_STATUS_D.Checked Then
                            Dim STATUS_COLOR_CODEs As String = ""
                            If chkSTYLE_COLOR_STATUS_A.Checked Then STATUS_COLOR_CODEs &= ",'A'"
                            If chkSTYLE_COLOR_STATUS_N.Checked Then STATUS_COLOR_CODEs &= ",'N'"
                            If chkSTYLE_COLOR_STATUS_D.Checked Then STATUS_COLOR_CODEs &= ",'D'"
                            If STATUS_COLOR_CODEs = "" Then
                                MsgBox("Cannot Find Style if No Color Status is Selected", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                            Else
                                ASCMAIN1.sql &= " and ICTSTYC1.STYLE_COLOR_STATUS in (" & Mid(STATUS_COLOR_CODEs, 2) & ")" & vbCrLf
                            End If
                        Else
                        End If

                        If optAvail.Value = "C" Then
                            ASCMAIN1.sql &= " and NVL(Z.ONH,0) - NVL(Z.OPEN,0) - NVL(Z.PICK,0) > 0"
                            ' QD &= " ;" & "Current Available to Sell Only"
                        ElseIf optAvail.Value = "A" Then
                            ASCMAIN1.sql &= " and NVL(Z.ONH,0) - NVL(Z.OPEN,0) - NVL(Z.PICK,0) + NVL(Z.ONPO,0) + NVL(Z.TRAN,0) > 0"
                            'QD &= " ;" & "Current Available to Sell Only" - THIS IS WRONG
                        End If
                        QD &= " ;" & optAvail.Text
                    End If
                End If
            End If
        End If

        grdICTSTYL1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        Dim TT As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = String.Format("ALTER TABLE {0} ADD ATTR_CODE2 VARCHAR2(6)", TT)
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = String.Format("ALTER TABLE {0} ADD ATTR_CODE3 VARCHAR2(6)", TT)
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = String.Format("ALTER TABLE {0} ADD ATTR_CODE4 VARCHAR2(6)", TT)
        ASCDATA1.ExecuteSQL()

        Fill_Records("ICTSTYL1", "", True, "Select * from " & TT)

        ASCMAIN1.sql = "Select ICTSTDQ2.* from ICTSTDQ2, " & TT & " ICTSTYL1 " _
        & " where ICTSTDQ2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
        & " and ICTSTDQ2.COLOR_CODE = ICTSTYL1.COLOR_CODE" _
        & " and ICTSTDQ2.WHSE_CODE = '" & WHSE_CODE & "'"
        Fill_Records("ICTSTDQ2", "", True, ASCMAIN1.sql)

        Dim tblEXCEL_LIST As DataTable = Nothing
        If Load_from_Excel Then
            If EXCEL_LIST.Length <> 0 Then
                Dim sql As New System.Text.StringBuilder With {.Length = 0}
                sql.AppendLine("SELECT *")
                sql.AppendLine("FROM " & EXCEL_LIST)
                tblEXCEL_LIST = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
            End If
        End If
        For Each row As DataRow In dst.Tables("ICTSTYL1").Select()
            row.Item("THEME_DESC") = GET_THEME_INFO(row.Item("THEME_CODE").ToString & String.Empty, "THEME_DESC")
            row.Item("SEASON_CODE") = GET_THEME_INFO(row.Item("THEME_CODE").ToString & String.Empty, "SEASON_CODE")

            Dim sc As String = row.Item("STYLE_CODE") & ""
            Dim cc As String = row.Item("COLOR_CODE") & ""
            Dim drc As String = row.Item("DUTY_RATE_CODE") & ""

            row.Item("DUTY_RATE") = GET_DUTY_RATE(drc)

            Dim rowICTSTDQ2 As DataRow = dst.Tables("ICTSTDQ2").Rows.Find(New Object() {WHSE_CODE, sc, cc})
            'If sc = "MTX57398" And cc = "GRSV" Then Stop
            If rowICTSTDQ2 IsNot Nothing Then
                For I As Integer = 1 To 3
                    If Val(rowICTSTDQ2.Item("QTY_" & I.ToString) & String.Empty) <> 0 Then
                        row.Item("FUT" & I.ToString & "_DATE") = rowICTSTDQ2.Item("DATE_" & I.ToString)
                        If I = 1 Then
                            If FilterAlloc(sc, cc, WHSE_CODE) = True And Val(rowICTSTDQ2.Item("QTY_" & I.ToString) & String.Empty) = 0 Then
                                row.Item("FUT" & I.ToString & "_AVAIL") = 0
                            Else
                                row.Item("FUT" & I.ToString & "_AVAIL") = rowICTSTDQ2.Item("QTY_" & I.ToString)
                            End If
                        Else
                            row.Item("FUT" & I.ToString & "_AVAIL") = rowICTSTDQ2.Item("QTY_" & I.ToString)
                        End If
                    End If
                Next
            End If
            If Load_from_Excel Then
                Dim rowEXCELSORT As DataRow = tblEXCEL_LIST.Select(String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", sc, cc)).FirstOrDefault
                If Not IsNothing(rowEXCELSORT) Then
                    row.Item("IMPORT_SORT") = rowEXCELSORT.Item("IMPORT_SORT").ToString & String.Empty
                End If
            End If
        Next

        'With grdICTSTYL1.DisplayLayout.Bands(0)
        '    For Each COLUMN_NAME As String In New String() {"OHMS", "POMS", "PSMS", "OHSW", "POSW", "PSSW"}
        '        .Columns(COLUMN_NAME).Format = "#,##0"
        '        .Columns(COLUMN_NAME).Width = 80
        '    Next
        'End With

        SetColorCodeView()

        ASCMAIN1.grdInitializeLayout(grdICTSTYL1)
        Sort_grdColumns(grdICTSTYL1, "STYLE_CODE")
        ATTR_CODE_1s = ""
        'splICTSTYL1.Visible = True
        TabStyles.Visible = True
        grdICTATTRQ.Visible = False
        txtDescription.Text = Mid(QD, 3)

        With UltraExplorerBar1
            If (grdICTSTYL1.Rows.Count > 0) Then
                .Groups("Screen Control").Items("Email").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Excel").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Attribute Excel").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Zip").Settings.Enabled = DefaultableBoolean.True
            Else
                .Groups("Screen Control").Items("Email").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Excel").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Attribute Excel").Settings.Enabled = DefaultableBoolean.False
                .Groups("Screen Control").Items("Zip").Settings.Enabled = DefaultableBoolean.False
            End If

        End With
        chkAttachZip.Enabled = (grdICTSTYL1.Rows.Count > 0)
        If Load_from_Excel Then
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("IMPORT_SORT").Hidden = False
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("IMPORT_SORT").Format = "###,###"
        Else
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("IMPORT_SORT").Hidden = True
        End If
    End Sub

    Private Sub BUILD_SCC()
        SCC_CLEAR(False)
        'Dim rowALL As DataRow = dst.Tables.Item("ICTCLAS1").Select("STYLE_CLASS_CODE = 'ALL'").FirstOrDefault
        'If rowALL.Item("SEL") = "1" Then
        '    SCCs.Add("ALL")
        'Else
        For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select()
            If rowICTCLAS1.Item("SEL") = "1" Then
                SCCs.Add(rowICTCLAS1.Item("STYLE_CLASS_CODE").ToString & String.Empty)
            End If
        Next
        'End If

        If SCCs.Count = 0 Then
            SCC_IN = "'NONE'"
        End If
        If SCCs.Count = 1 Then
            'If SCCs(0) = "ALL" Then
            '    For Each rowICTCLAS1 As DataRow In dst.Tables("ICTCLAS1").Select("STYLE_CLASS_CODE <> 'ALL'")
            '        SCC_IN = SCC_IN & "','" & rowICTCLAS1.Item("STYLE_CLASS_CODE").ToString & String.Empty
            '    Next
            'Else
            SCC_IN = $"'{SCCs(0)}'"
            'End If
        End If
        If SCCs.Count > 1 Then
            For Each SCC As String In SCCs
                SCC_IN = SCC_IN & "','" & SCC & String.Empty
            Next
            SCC_IN = SCC_IN.Substring(2, SCC_IN.Length - 2) & "'"
        End If
    End Sub

    Private Function GET_DUTY_RATE(ByVal DUTY_RATE_CODE As String) As Double
        Dim RetVal As Double = 0
        Dim filter As String = String.Format("DUTY_RATE_CODE = '{0}'", DUTY_RATE_CODE)
        Dim rowICTDUTY1 As DataRow = dst.Tables("ICTDUTY1").Select(filter).FirstOrDefault
        If Not IsNothing(rowICTDUTY1) Then
            RetVal = Val(rowICTDUTY1.Item("DUTY_RATE").ToString & String.Empty) * 0.01
        End If
        Return RetVal
    End Function

    Private Function GET_THEME_INFO(ByVal THEME_CODE As String, ByVal COL_NAME As String) As String
        Dim RetVal As String = ""
        Dim filter As String = String.Format("THEME_CODE = '{0}'", THEME_CODE)
        Dim rowICTTHEME As DataRow = dst.Tables("ICTTHEME").Select(filter).FirstOrDefault
        If Not IsNothing(rowICTTHEME) Then
            RetVal = rowICTTHEME.Item(COL_NAME).ToString & String.Empty
        End If
        Return RetVal
    End Function

    Sub Hot_Key_Part_Two(hkpo As String, e As System.Windows.Forms.KeyEventArgs)

        Dim kc As New KeysConverter
        Dim kInt As Integer = e.KeyCode
        Dim cKey As String = kc.ConvertToString(kInt)

        If (e.Alt AndAlso (cKey <> "")) Then
            Dim hkot As String = hotKeyPartOne & cKey
            Select Case hkot
                Case "DS"
                    'user has typed Alt + D + S, show the DataSet Utility
                    Dim frmASFDSET1 As New ASFDSET1(ASCMAIN1.ActiveForm)
                    frmASFDSET1.Show()
                Case "ZZ" 'set up other hot key combinations like so
                Case "AZ"
            End Select
        End If

        hotKeyPartOne = ""

    End Sub

    Sub Hot_Key_Part_One(e As System.Windows.Forms.KeyEventArgs)
        Dim kc As New KeysConverter
        Dim kInt As Integer = e.KeyCode
        Dim cKey As String = kc.ConvertToString(kInt)

        If (e.Alt AndAlso (cKey <> "" And kInt <> 18)) Then
            hotKeyPartOne = cKey
        End If
    End Sub

    Private Function Import_From_Excel() As Boolean
        Dim RetVal As Boolean = False
        OpenFileDialog1.DefaultExt = "xlsx"
        OpenFileDialog1.ShowDialog()
        Dim BadFileName As Boolean = True
        Dim ThisFileExt As String = ""
        If OpenFileDialog1.FileNames.Length = 1 Then
            Me.Cursor = Cursors.WaitCursor
            Dim excel As Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim XWB As Excel.Workbook = excel.Workbooks.Add
            Dim XWS As Excel.Worksheet = XWB.Sheets(1)
            Dim FullFileName As String = OpenFileDialog1.FileNames(0)
            Try
                XWB = excel.Workbooks.Open(FullFileName)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Opening File")
                XWB.Close()
                XWB = Nothing
                excel = Nothing
                Me.Cursor = Cursors.Default
                Return RetVal
                'Exit Sub
            End Try

            XWS = XWB.Worksheets(1)
            Dim StyleColPos As Integer = 0
            Dim ColorColPos As Integer = 0
            Dim StyleColFound As Boolean = False
            Dim ColorColFound As Boolean = False
            For i As Integer = 1 To 30
                Dim CELL_TXT As String = ""
                CELL_TXT = XWS.Cells(1, i).text.ToString & ""
                If CELL_TXT.ToUpper = "STYLE CODE" Then
                    StyleColPos = i
                    StyleColFound = True
                End If
                If CELL_TXT.ToUpper = "COLOR CODE" Then
                    ColorColPos = i
                    ColorColFound = True
                End If
            Next
            If StyleColFound And ColorColFound Then
                Dim NullTolerance As Integer = 10
                Dim NullsFound As Integer = 0
                Dim StylesToLoad As New List(Of String)
                Dim ColorsToLoad As New List(Of String)
                Dim SCCombo As New List(Of String)
                Dim currCount As Integer = 1
                Do While (NullsFound < NullTolerance)
                    currCount += 1
                    If currCount > 30000 Then
                        Exit Do 'Just to stop runaway processes
                    End If
                    Dim STYLE_CODE As String = XWS.Cells(currCount, StyleColPos).text.ToString & ""
                    Dim COLOR_CODE As String = XWS.Cells(currCount, ColorColPos).text.ToString & ""
                    If STYLE_CODE.Length = 0 Or COLOR_CODE.Length = 0 Then
                        NullsFound += 1
                    Else
                        If SCCombo.IndexOf(STYLE_CODE & "-" & COLOR_CODE) = -1 Then
                            SCCombo.Add(STYLE_CODE & "-" & COLOR_CODE)
                            StylesToLoad.Add(STYLE_CODE)
                            ColorsToLoad.Add(COLOR_CODE)
                        End If
                    End If
                Loop
                XWB.Close()
                XWB = Nothing
                excel = Nothing

                If SCCombo.Count > 0 Then
                    Find_Styles(False, False, True, StylesToLoad, ColorsToLoad)
                    RetVal = True
                End If
            Else
                MsgBox("Could Not Find Style/Color Columns", vbOKOnly, "Excel")
                Me.Cursor = Cursors.Default
                Return RetVal
            End If
        Else
            MsgBox("Bad File Selected", vbOKOnly, "Excel")
            Me.Cursor = Cursors.Default
            Return RetVal
        End If
        Me.Cursor = Cursors.Default
        Return RetVal
    End Function

    Private Sub showSelectors(ByVal showSel As Boolean)
        splChoices.Panel1Collapsed = Not showSel
        'lblClass.Visible = showSel
        'txtSTYLE_CLASS_CODE.Visible = showSel
        btnAllDiscontinued.Visible = showSel
        btnAllActive.Visible = showSel
        btnSelectClass.Visible = showSel
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            btnCatelog.Visible = showSel
        Else
            btnCatelog.Visible = False
        End If
        If showSel = True Then
            SCC_CLEAR(False)
        End If
    End Sub

    Private Sub StartSplash()
        progressSplash = New ASFPROGS(progressSplashMsg1, progressSplashMsg2, progressSplashMsg3, Me.Bounds, True)
        Application.Run(progressSplash)
    End Sub

#End Region

#Region "Form Control Events"

    Private Sub tabAttributes_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabAttributes.SelectedTabChanged
        If tabAttributes.SelectedTab.Key = "1" Then
            txtATTR_CODE_1.Focus()
            txtATTR_CODE_1.Text = ""

        ElseIf tabAttributes.SelectedTab.Key = "Size" Then
            txtSIZE_CODE.Focus()
            txtSIZE_CODE.Text = ""

            If LAST_ATTR = "1" Then
                ATTR_CODE_1s = Get_CODE_VALUEs(grdICTATTR1_1)
                SIZE_CODEs = Get_CODE_VALUEs(grdICTSIZE1)
                'If SCCs(0) = "ALL" Then
                '    'ASCMAIN1.sql = "Select Distinct SIZE_CODE from ICTSTYL1" _
                '    '    & " where STYLE_CLASS_CODE <> '" & STYLE_CLASS_CODE & "' and SIZE_CODE is Not Null" _
                '    '    & IIf(ATTR_CODE_1s = "", "", "   and STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(ATTR_CODE_1s, 2) & "))")
                '    ASCMAIN1.sql = "Select Distinct SIZE_CODE from ICTSTYL1" _
                '        & " where STYLE_CLASS_CODE NOT IN (" & SCC_IN & ") and SIZE_CODE is Not Null" _
                '        & IIf(ATTR_CODE_1s = "", "", "   and STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(ATTR_CODE_1s, 2) & "))")
                'Else
                '    'ASCMAIN1.sql = "Select Distinct SIZE_CODE from ICTSTYL1" _
                '    '& " where STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "' and SIZE_CODE is Not Null" _
                '    '& IIf(ATTR_CODE_1s = "", "", "   and STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(ATTR_CODE_1s, 2) & "))")
                ASCMAIN1.sql = "Select Distinct SIZE_CODE from ICTSTYL1" _
                    & " where STYLE_CLASS_CODE IN (" & SCC_IN & ") and SIZE_CODE is Not Null" _
                    & IIf(ATTR_CODE_1s = "", "", "   and STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(ATTR_CODE_1s, 2) & "))")
                'End If

                Fill_Records("ICTSIZE1", "", True, ASCMAIN1.sql)
                'dst.Tables("ICTSIZE1").Rows.Add(New String() {"", ""})
                Sort_grdColumns(grdICTSIZE1, "SIZE_CODE")
                If SIZE_CODEs <> "" Then
                    For Each row As DataRow In dst.Tables("ICTSIZE1").Select("SIZE_CODE in (" & Mid(SIZE_CODEs, 2) & ")")
                        row.Item("SEL") = "1"
                    Next
                End If
            End If

        ElseIf tabAttributes.SelectedTab.Key = "2" Then
            Dim AT1s = Get_CODE_VALUEs(grdICTATTR1_1)
            txtATTR_CODE_2.Focus()
            txtATTR_CODE_2.Text = ""

            If LAST_ATTR = "Size" Or LAST_ATTR = "1" Then
                SIZE_CODEs = Get_CODE_VALUEs(grdICTSIZE1)
                ATTR_CODE_2s = Get_CODE_VALUEs(grdICTATTR1_2)
                'If SCCs(0) = "ALL" Then
                '    'ASCMAIN1.sql = "Select Distinct ICTSTYL3.ATTR_CODE from ICTSTYL3,ICTSTYL1" _
                '    '& " where ICTSTYL3.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                '    '& "   and ICTSTYL1.STYLE_CLASS_CODE <> '" & STYLE_CLASS_CODE & "'" _
                '    '& IIf(SIZE_CODEs = "", "", "   and NVL(ICTSTYL1.SIZE_CODE,'?') in (" & Mid(SIZE_CODEs, 2) & ")") _
                '    '& IIf(AT1s = "", "", "   and ICTSTYL3.STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(AT1s, 2) & "))")
                '    ASCMAIN1.sql = "Select Distinct ICTSTYL3.ATTR_CODE from ICTSTYL3,ICTSTYL1" _
                '    & " where ICTSTYL3.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                '    & "   and ICTSTYL1.STYLE_CLASS_CODE IN (" & SCC_IN & ")" _
                '    & IIf(SIZE_CODEs = "", "", "   and NVL(ICTSTYL1.SIZE_CODE,'?') in (" & Mid(SIZE_CODEs, 2) & ")") _
                '    & IIf(AT1s = "", "", "   and ICTSTYL3.STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(AT1s, 2) & "))")

                'Else
                '    'ASCMAIN1.sql = "Select Distinct ICTSTYL3.ATTR_CODE from ICTSTYL3,ICTSTYL1" _
                '    '& " where ICTSTYL3.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                '    '& "   and ICTSTYL1.STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'" _
                '    '& IIf(SIZE_CODEs = "", "", "   and NVL(ICTSTYL1.SIZE_CODE,'?') in (" & Mid(SIZE_CODEs, 2) & ")") _
                '    '& IIf(AT1s = "", "", "   and ICTSTYL3.STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(AT1s, 2) & "))")
                ASCMAIN1.sql = "Select Distinct ICTSTYL3.ATTR_CODE from ICTSTYL3,ICTSTYL1" _
                & " where ICTSTYL3.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                & "   and ICTSTYL1.STYLE_CLASS_CODE IN (" & SCC_IN & ")" _
                & IIf(SIZE_CODEs = "", "", "   and NVL(ICTSTYL1.SIZE_CODE,'?') in (" & Mid(SIZE_CODEs, 2) & ")") _
                & IIf(AT1s = "", "", "   and ICTSTYL3.STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(AT1s, 2) & "))")
                'End If

                ASCMAIN1.sql = "Select ATTR_CODE, ATTR_DESC from ICTATTR1 where ATTR_CODE in (" & ASCMAIN1.sql & ") AND NVL(ATT_RANK,9) <> 1"
                Fill_Records("ICTATTR1_2", "", True, ASCMAIN1.sql)
                For Each row As DataRow In dst.Tables("ICTATTR1_2").Select("ISNULL(ATTR_DESC,'')=''")
                    row.Item("ATTR_DESC") = row.Item("ATTR_CODE")
                Next
                Sort_grdColumns(grdICTATTR1_2, "ATTR_CODE")
                If ATTR_CODE_2s <> "" Then
                    For Each row As DataRow In dst.Tables("ICTATTR1_2").Select("ATTR_CODE in (" & Mid(ATTR_CODE_2s, 2) & ")")
                        row.Item("SEL") = "1"
                    Next
                End If
            End If

        End If
    End Sub

    Private Sub txtATTR_CODE_1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtATTR_CODE_1.ValueChanged
        Set_Pointer(grdICTATTR1_1, txtATTR_CODE_1)
    End Sub

    Private Sub txtSIZE_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtSIZE_CODE.ValueChanged
        Set_Pointer(grdICTSIZE1, txtSIZE_CODE)
    End Sub

    Private Sub txtATTR_CODE_2_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtATTR_CODE_2.ValueChanged
        Set_Pointer(grdICTATTR1_2, txtATTR_CODE_2)
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If txtDescription.Text = "" Then
            MsgBox("You need to specify a Description", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Dim rowICTATTRQ As DataRow = dst.Tables("ICTATTRQ").NewRow
        If txtQueryDesc.Text.Length > rowICTATTRQ.Table.Columns.Item("QUERY_DESC").MaxLength Then
            MsgBox("Description Field Too Large.", MsgBoxStyle.OkOnly, "Cannot Save")
            Exit Sub
        End If
        If Get_CODE_VALUEs(grdICTATTR1_1).ToString.Length > rowICTATTRQ.Table.Columns.Item("ATTR_CODE_1S").MaxLength Then
            MsgBox("Attribute List 1 Too Large.", MsgBoxStyle.OkOnly, "Cannot Save")
            Exit Sub
        End If
        If Get_CODE_VALUEs(grdICTATTR1_2).ToString.Length > rowICTATTRQ.Table.Columns.Item("ATTR_CODE_2S").MaxLength Then
            MsgBox("Attribute List 2 Too Large.", MsgBoxStyle.OkOnly, "Cannot Save")
            Exit Sub
        End If
        If Get_CODE_VALUEs(grdICTSIZE1).ToString.Length > rowICTATTRQ.Table.Columns.Item("SIZE_CODES").MaxLength Then
            MsgBox("Size Code List Too Large.", MsgBoxStyle.OkOnly, "Cannot Save")
            Exit Sub
        End If

        QUERY_NO = ASCMAIN1.Next_Control_No("ICTATTRQ.QUERY_NO")
        With rowICTATTRQ
            .Item("QUERY_NO") = QUERY_NO
            .Item("QUERY_DESC") = txtQueryDesc.Text
            '.Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
            .Item("STYLE_CLASS_CODE") = SCC_IN
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("INCLUDE_STATUS") = IIf(chkSTYLE_STATUS_A.Checked, "A", "") & IIf(chkSTYLE_STATUS_N.Checked, "N", "") & IIf(chkSTYLE_STATUS_D.Checked, "D", "")
            .Item("ATTR_CODE_1S") = Get_CODE_VALUEs(grdICTATTR1_1) ' ATTR_CODE_1s
            .Item("SIZE_CODES") = Get_CODE_VALUEs(grdICTSIZE1)  ' SIZE_CODEs
            .Item("ATTR_CODE_2S") = Get_CODE_VALUEs(grdICTATTR1_2)  ' ATTR_CODE_2s
            .Item("CUR_AVAIL") = IIf(optAvail.Value = "C", "1", "0") ' IIf(chkCUR_AVA.Checked, "1", "0")
            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
        End With
        dst.Tables("ICTATTRQ").Rows.Add(rowICTATTRQ)
        Update_Record_TDA("ICTATTRQ")
        MsgBox("Query Definition Saved", MsgBoxStyle.OkOnly, "Verification")
    End Sub

#End Region

#Region "Grid Events"
    Private Sub grdICTSTYL1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs)
        If ASCMAIN1.USER_ID = "rdw" Then
        Else
            If e.Row.IsDataRow And Not e.Row.IsAddRow Then
                STYLE_CODE = e.Row.Cells("STYLE_CODE").Value
                Me.Close()
            End If
        End If

    End Sub

    Private Sub grdICTATTRQ_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTATTRQ.DoubleClickRow
        Dim QUERY_NO As String = e.Row.Cells("QUERY_NO").Value
        Dim rowICTATTRQ As DataRow = dst.Tables("ICTATTRQ").Rows.Find(QUERY_NO)

        Absx1.txtFor("STYLE_CLASS_CODE").Text = rowICTATTRQ.Item("STYLE_CLASS_CODE") & ""

        '   Initialize_tabAttributes(False)
        rowICTATTRQ = dst.Tables("ICTATTRQ").Rows.Find(QUERY_NO)
        Absx1.txtFor("WHSE_CODE").Text = rowICTATTRQ.Item("WHSE_CODE") & ""
        Dim INCLUDE_STATUS As String = rowICTATTRQ.Item("INCLUDE_STATUS") & ""

        chkSTYLE_STATUS_A.Checked = INCLUDE_STATUS.Contains("A")
        chkSTYLE_STATUS_N.Checked = INCLUDE_STATUS.Contains("N")
        chkSTYLE_STATUS_D.Checked = INCLUDE_STATUS.Contains("D")
        'chkCUR_AVA.Checked = (rowICTATTRQ.Item("CUR_AVAIL") & "" = "1")
        If rowICTATTRQ.Item("CUR_AVAIL") & "" = "1" Then
            optAvail.Value = "C"
        Else
            optAvail.Value = "N"
        End If
        Dim ATTR_CODE_1s As String = rowICTATTRQ.Item("ATTR_CODE_1S") & ""
        Dim SIZE_CODEs As String = rowICTATTRQ.Item("SIZE_CODES") & ""
        Dim ATTR_CODE_2s As String = rowICTATTRQ.Item("ATTR_CODE_2S") & ""

        If ATTR_CODE_1s <> "" Then
            For Each row As DataRow In dst.Tables("ICTATTR1_1").Select()
                If ATTR_CODE_1s.Contains(row.Item(0)) Then
                    row.Item("SEL") = "1"
                End If
            Next
        End If
        LAST_ATTR = "1"

        tabAttributes.SelectedTab = tabAttributes.Tabs("Size")
        If SIZE_CODEs <> "" Then
            For Each row As DataRow In dst.Tables("ICTSIZE1").Select()
                If row.Item(0) & "" = "" Then
                    If SIZE_CODEs.Contains("''") Then
                        row.Item("SEL") = "1"
                    End If
                Else
                    If SIZE_CODEs.Contains(row.Item(0)) Then
                        row.Item("SEL") = "1"
                    End If
                End If
            Next
        End If
        LAST_ATTR = "Size"

        tabAttributes.SelectedTab = tabAttributes.Tabs("2")
        If ATTR_CODE_2s <> "" Then
            For Each row As DataRow In dst.Tables("ICTATTR1_2").Select()
                If ATTR_CODE_2s.Contains(row.Item(0)) Then
                    row.Item("SEL") = "1"
                End If
            Next
        End If

        tabAttributes.SelectedTab = tabAttributes.Tabs("1")
        txtQueryDesc.Text = rowICTATTRQ.Item("QUERY_DESC") & ""

        'STYLE_CLASS_CODE = Absx1.txtFor("STYLE_CLASS_CODE").Text
        BUILD_SCC()
        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

        Find_Styles()
    End Sub

    Private Sub grdICTCLAS1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTCLAS1.DoubleClickRow
        'Absx1.txtFor("STYLE_CLASS_CODE").Text = e.Row.Cells("STYLE_CLASS_CODE").Value
    End Sub

    Private Sub grdICTATTR1_1_CellChange(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTATTR1_1.CellChange
        If grdICTATTR1_1.ActiveRow IsNot Nothing AndAlso grdICTATTR1_1.ActiveRow.DataChanged Then grdICTATTR1_1.ActiveRow.Update()
    End Sub

    Private Sub grdICTATTR1_1_ClickCell(sender As Object, e As Infragistics.Win.UltraWinGrid.ClickCellEventArgs) Handles grdICTATTR1_1.ClickCell
        If grdICTATTR1_1.ActiveRow IsNot Nothing AndAlso grdICTATTR1_1.ActiveRow.DataChanged Then grdICTATTR1_1.ActiveRow.Update()
    End Sub

    Private Sub grdICTATTR1_1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdICTATTR1_1.MouseClick
        If grdICTATTR1_1.ActiveRow IsNot Nothing AndAlso grdICTATTR1_1.ActiveRow.DataChanged Then grdICTATTR1_1.ActiveRow.Update()
    End Sub

    Private Sub grdICTATTR1_1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdICTATTR1_1.MouseUp
        If grdICTATTR1_1.ActiveRow IsNot Nothing AndAlso grdICTATTR1_1.ActiveRow.DataChanged Then grdICTATTR1_1.ActiveRow.Update()
    End Sub

    Private Sub grdICTATTR1_1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTATTR1_1.AfterCellUpdate
        LAST_ATTR = "1"
        If grdICTATTR1_1.ActiveRow IsNot Nothing AndAlso grdICTATTR1_1.ActiveRow.DataChanged Then grdICTATTR1_1.ActiveRow.Update()
    End Sub

    Private Sub grdICTSIZE1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSIZE1.AfterCellUpdate
        LAST_ATTR = "Size"
    End Sub

    Private Sub grdICTATTR1_2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTATTR1_2.AfterCellUpdate
        LAST_ATTR = "2"
    End Sub

#End Region

    Sub Initialize_tabAttributes()
        'STYLE_CLASS_CODE = Absx1.txtFor("STYLE_CLASS_CODE").Text
        If SCCs.Count = 0 Then
            tabAttributes.Visible = False
            grdICTCLAS1.Visible = True
            LAST_ATTR = ""
        Else
            tabAttributes.Visible = True
            grdICTCLAS1.Visible = False
            'If SCCs(0) = "ALL" Then
            '    'ASCMAIN1.sql = "Select Distinct ATTR_CODE from ICTSTYL3 " _
            '    '    & " where STYLE_CODE in (Select STYLE_CODE from ICTSTYL1 where STYLE_CLASS_CODE <> '" & STYLE_CLASS_CODE & "')"
            '    ASCMAIN1.sql = "Select Distinct ATTR_CODE from ICTSTYL3 "
            'Else
            ASCMAIN1.sql = "Select Distinct ATTR_CODE from ICTSTYL3 " _
                & " where STYLE_CODE in (Select STYLE_CODE from ICTSTYL1 where STYLE_CLASS_CODE IN (" & SCC_IN & "))"
            'End If

            ASCMAIN1.sql = "Select ATTR_CODE, ATTR_DESC from ICTATTR1 where ATTR_CODE in (" & ASCMAIN1.sql & ") AND NVL(ATT_RANK,9) = 1"
            Fill_Records("ICTATTR1_1", "", True, ASCMAIN1.sql)
            For Each row As DataRow In dst.Tables("ICTATTR1_1").Select("ISNULL(ATTR_DESC,'')=''")
                row.Item("ATTR_DESC") = row.Item("ATTR_CODE")
            Next
            Sort_grdColumns(grdICTATTR1_1, "ATTR_CODE")

            tabAttributes.SelectedTab = tabAttributes.Tabs("1")
        End If

        For Each TABLE_NAME As String In New String() {"ICTATTR1_1", "ICTATTR1_2", "ICTSIZE1", "ICTATTR1_1"}
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("SEL = '1'")
                row.Item("SEL") = "0"
            Next
        Next

        Fill_Records("ICTATTRQ", ASCMAIN1.USER_ID)
        txtQueryDesc.Text = ""

        QUERY_NO = ""
        ATTR_CODE_1s = ""
        SIZE_CODEs = ""
        ATTR_CODE_2s = ""


        'UltraExplorerBar1.Groups("Screen Control").Items("Find").Visible = tabAttributes.Visible
        'cmdFind.Enabled = tabAttributes.Visible
        UltraExplorerBar1.Groups("Screen Control").Items("Clear").Visible = tabAttributes.Visible
        txtATTR_CODE_1.Focus()


        With UltraExplorerBar1
            .Groups("Screen Control").Items("Email").Settings.Enabled = False
            .Groups("Screen Control").Items("Excel").Settings.Enabled = False
            .Groups("Screen Control").Items("Attribute Excel").Settings.Enabled = False
            .Groups("Screen Control").Items("Zip").Settings.Enabled = False
        End With

        chkAttachZip.Enabled = False
    End Sub

    Sub Size_Header()
        Dim grpsHeight As Integer = 77
        grpInclude.Height = grpsHeight
        'grpWhse.Height = grpsHeight
        'grpBtns.Height = grpsHeight - 10

        'grpClass.Width = grdICTATTR1_1.Width

        grpInclude.Width = (chkSTYLE_STATUS_D.Width + optAvail.Width + 10) '  (chkSTYLE_STATUS_D.Width + chkCUR_AVA.Width + 10)
        'grpWhse.Width = 115
        'grpBtns.Width = Me.Width - (grpClass.Width + grpInclude.Width + grpWhse.Width + 20)
        'grpBtns.Top = grpWhse.Top + 10
    End Sub

    Function Get_CODE_VALUEs(grd As UltraWinGrid.UltraGrid)
        Dim CODE_VALUEs As String = ""
        Dim tbl As DataTable = DirectCast(grd.DataSource, DataTable)
        For Each row As DataRow In tbl.Select("SEL = '1'")
            Dim CODE_VALUE As String = row.Item(0)
            If CODE_VALUE = "" Then CODE_VALUE = "?"
            CODE_VALUEs &= ",'" & CODE_VALUE & "'"
        Next
        Return CODE_VALUEs
    End Function

    Function Generate_Excel() As String
        Dim pbInt As Integer = 0
        Dim excelFile As String = ""
        Dim FILE_NAME As String = ""
        Dim UseOrderPricing As Boolean = txtOrder.Text.Length > 0


        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

        Dim BRANCH_CODE As String = "NY"
        Dim xlPages As New Dictionary(Of Integer, Integer)

        If dst.Tables("ICTSTYL1").Select("SEL='1'").Length = 0 Then
            MsgBox("No Styles Selected", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
            Return excelFile
        End If

        Dim excel As Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Dim XWB As Excel.Workbook = excel.Workbooks.Add
        Dim XWS As Excel.Worksheet = XWB.Sheets(1)
        Dim rng As Excel.Range

        myExcelHasBalls = Does_This_Version_Of_Excel_Have_Balls(excel.Version)

        'insert logo
        Dim LOGO_FILENAME As String = IMAGES_FOLDER & "rgiLogo.jpg"
        If My.Computer.FileSystem.FileExists(LOGO_FILENAME) Then
            rng = XWS.Range("A" & CStr(1) & ":" & "B" & CStr(6))
            InsertPictureInRange(LOGO_FILENAME, rng, XWS, "", "")
        End If

        rng = XWS.Range("E1:E1")
        rng.FormulaR1C1 = "Regency International"
        rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
        With rng.Font
            .Name = "Georgia"
            .Size = 18
            .Color = Color.FromArgb(79, 129, 189)
            .Bold = False
        End With

        rng = XWS.Range("E3:E3")
        rng.FormulaR1C1 = "Pricing and Availability"
        rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        rng = XWS.Range("E4:E4")
        rng.FormulaR1C1 = IIf(txtQueryDesc.Text & "" = "", "Custom Query", txtQueryDesc.Text)
        rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        rng = XWS.Range("E5:E5")
        rng.FormulaR1C1 = "Generated: " & Now.ToShortDateString()
        rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

        Dim rowSOTBRAN1 As DataRow = dst.Tables("SOTBRAN1").Rows.Find(BRANCH_CODE)

        If rowSOTBRAN1 IsNot Nothing Then
            Add_Branch_Info(XWS, rowSOTBRAN1)
        End If

        Create_Worksheet_Headers(XWS)

        'freeze header region
        XWS.Range("A10", "A10").Select()
        XWS.Application.ActiveWindow.FreezePanes = True

        Dim C As Integer = 0
        Dim R As Integer = 11
        Dim rowsToProcess As Integer = dst.Tables("ICTSTYL1").Select("SEL = '1'").Length + 1
        Dim currentRow As Integer = 1
        Dim currentPage As Integer = 1
        xlPages.Clear()

        Dim sby As String = "STYLE_CODE,COLOR_CODE"
        If chkGroupBySize.Checked Then
            sby = "SIZE_CODE," & sby
        Else
            sby = SetExcelSortBy()
        End If
        Dim SIZE_CODE_last As String = ""

        Dim totLines As New List(Of Integer)

        For Each row As DataRow In dst.Tables("ICTSTYL1").Select("SEL = '1'", sby)

            pbInt = CInt((currentRow / rowsToProcess) * 100)
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE") & ""

            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})

            Dim SIZE_CODE As String = row.Item("SIZE_CODE") & ""

            If chkGroupBySize.Checked Then
                If SIZE_CODE <> SIZE_CODE_last Then
                    SIZE_CODE_last = SIZE_CODE
                    XWS.Cells(R, 4).VALUE = "Size"
                    XWS.Cells(R, 5).VALUE = SIZE_CODE
                    R += 1
                End If
            End If

            progressSplash.UpdateProgress("Formatting Spreadsheet", "Style: " & STYLE_CODE & ":" & COLOR_CODE, "", pbInt)

            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            XWS.Cells(R, 4).VALUE = "Style"
            XWS.Cells(R, 5).VALUE = STYLE_CODE
            XWS.Cells(R + 1, 4).VALUE = "Description"
            XWS.Cells(R + 1, 5).VALUE = rowICTSTYL1.Item("STYLE_DESC") & ""
            XWS.Cells(R + 2, 4).VALUE = "Color"
            Dim rowICTCOLR1 As DataRow = dst.Tables("ICTCOLR1").Rows.Find(COLOR_CODE)
            If rowICTCOLR1 IsNot Nothing Then
                XWS.Cells(R + 2, 5).VALUE = rowICTCOLR1.Item("COLOR_DESC")
            Else
                XWS.Cells(R + 2, 5).VALUE = COLOR_CODE
            End If

            XWS.Cells(R + 3, 4).VALUE = "Size"
            Dim rowICTSIZE1 As DataRow = dst.Tables("ICTSIZE1").Rows.Find(SIZE_CODE)
            If rowICTSIZE1 IsNot Nothing Then
                XWS.Cells(R + 3, 5).VALUE = rowICTSIZE1.Item("SIZE_DESC")
            Else
                XWS.Cells(R + 3, 5).VALUE = SIZE_CODE
            End If
            XWS.Cells(R + 3, 5).VALUE = rowICTSTYL1.Item("SIZE_CODE") & ""
            rng = XWS.Range("E" & CStr(R) & ":" & "E" & CStr(R + 4))
            rng.Font.Bold = True

            XWS.Cells(R, 7).VALUE = "Inner"
            XWS.Cells(R, 8).VALUE = rowICTSTYL1.Item("INNER_PACK_QTY") & ""
            XWS.Cells(R + 1, 7).VALUE = "Case"
            XWS.Cells(R + 1, 8).VALUE = rowICTSTYL1.Item("CARTON_PACK_QTY") & ""
            XWS.Cells(R + 2, 7).VALUE = "UM"
            XWS.Cells(R + 2, 8).VALUE = rowICTSTYL1.Item("STYLE_UOM") & ""
            XWS.Cells(R + 3, 7).VALUE = "Class"
            XWS.Cells(R + 3, 8).VALUE = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
            XWS.Cells(R + 4, 7).VALUE = "Cube"
            XWS.Cells(R + 4, 8).VALUE = rowICTSTYL1.Item("CASE_CUBE") & ""
            rng = XWS.Range("H" & CStr(R) & ":" & "H" & CStr(R + 4))
            rng.Font.Bold = True
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight

            Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(rowICTSTYL1.Item("STYLE_CLASS_CODE") & "")
            If rowICTCLAS1 IsNot Nothing Then
                Dim DISC_CODE As String = rowICTCLAS1.Item("DISC_CODE") & ""
                Dim rowICTDISC1 As DataRow = dst.Tables("ICTDISC1").Rows.Find(DISC_CODE)
                If rowICTDISC1 IsNot Nothing Then
                    Dim STYLE_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PRICE") & "")
                    For I As Integer = 1 To 4
                        Dim Pos As Integer = 0
                        Select Case I
                            Case Is = 1
                                Pos = 4
                            Case Is = 2
                                Pos = 3
                            Case Is = 3
                                Pos = 2
                            Case Is = 4
                                Pos = 1
                        End Select
                        Dim PRICE As Decimal = STYLE_PRICE * (100 - Val(rowICTDISC1.Item("DISC" & CStr(I) & "_PCT"))) / 100

                        If DISC_CODE = "NONPVC" Then
                            If optPRICE_TIER.Value = "SP" Then
                                Dim CUST_DISC_PCT_EXTRA_PCT As Decimal = 0
                                If Val(numDISC_PCT.Value & "") > 0 Then
                                    PRICE = (100 - Val(numDISC_PCT.Value & "")) * Val(rowICTSTYL1.Item("STYLE_PRICE") & "") / 100
                                Else
                                    Select Case optDISC_PCT_EXTRA.Value
                                        Case Is = "1"
                                            CUST_DISC_PCT_EXTRA_PCT = 5
                                        Case Is = "2"
                                            CUST_DISC_PCT_EXTRA_PCT = 10
                                        Case Else
                                            CUST_DISC_PCT_EXTRA_PCT = Val(numDISC_PCT.Value & "")
                                    End Select
                                    PRICE = (100 - CUST_DISC_PCT_EXTRA_PCT) * PRICE / 100
                                End If

                                If PRICE < 0 Then
                                    PRICE = 0
                                End If
                            End If
                        End If

                        Dim STYLE_STATUS As String = rowICTSTYL1.Item("STYLE_STATUS") & ""
                        Dim STYLE_COLOR_STATUS As String = rowICTSTYC1.Item("STYLE_COLOR_STATUS") & ""
                        Dim IsDiscontinued As Boolean = False
                        If STYLE_STATUS = "D" Or (COLOR_CODE <> "" AndAlso STYLE_COLOR_STATUS = "D") Then
                            PRICE = 0.3 * STYLE_PRICE
                            IsDiscontinued = True
                        End If
                        Dim pricingDesc As String = String.Format("{0}({1})", rowICTDISC1.Item("DISC" & CStr(I) & "_DESC"), Val(rowICTSTYL1.Item("CARTON_PACK_QTY") * rowICTDISC1.Item("DISC" & CStr(I) & "_CASES")))
                        If chkPBFE.Checked Then
                            Dim Factor As Integer = numFEFDFACTOR.Value
                            If Factor >= 0 Then
                                Factor = 1
                            End If
                            If Factor < -14 Then
                                Factor = 1
                            End If
                            Dim FEFD As New FEFDPrice(Me, rowICTSTYL1.Item("STYLE_CODE").ToString, Factor)
                            If FEFD.ErrorMsg = "" Then
                                Select Case Pos
                                    Case Is = 4
                                        If optFEFD.Value = "FDM" Then
                                            PRICE = Format(FEFD.FDMixPrice, "###,##0.00")
                                            pricingDesc = "FED Mix Price"
                                        Else
                                            PRICE = Format(0, "###,##0.00")
                                            pricingDesc = ""
                                        End If
                                    Case Is = 3
                                        If optFEFD.Value = "FEM" Then
                                            PRICE = Format(FEFD.FEMixPrice, "###,##0.00")
                                            pricingDesc = "FE Mix Price"
                                            Pos = 4
                                        Else
                                            PRICE = Format(0, "###,##0.00")
                                            pricingDesc = ""
                                        End If
                                    Case Is = 2
                                        If optFEFD.Value = "FD" Then
                                            PRICE = Format(FEFD.FDPrice, "###,##0.00")
                                            pricingDesc = "FD Price"
                                            Pos = 4
                                        Else
                                            PRICE = Format(0, "###,##0.00")
                                            pricingDesc = ""
                                        End If
                                    Case Is = 1
                                        If optFEFD.Value = "FE" Then
                                            PRICE = Format(FEFD.FEPrice, "###,##0.00")
                                            pricingDesc = "FE Price"
                                            Pos = 4
                                        Else
                                            PRICE = Format(0, "###,##0.00")
                                            pricingDesc = ""
                                        End If
                                End Select
                            End If
                        End If

                        If UseOrderPricing Then
                            Dim ORDR_PRICE_QTY As Double() = GetOrderPrice(STYLE_CODE, COLOR_CODE)
                            Dim ORDRTOTAL As Double = ORDR_PRICE_QTY(0) * ORDR_PRICE_QTY(1)
                            Select Case Pos
                                Case Is = 4
                                    XWS.Cells(R + (Pos), 10).VALUE = ORDRTOTAL
                                    XWS.Cells(R + (Pos), 11).VALUE = "Order Total"
                                Case Is = 3
                                    XWS.Cells(R + (Pos), 10).VALUE = ORDR_PRICE_QTY(1)
                                    XWS.Cells(R + (Pos), 11).VALUE = "Order Qty"
                                Case Is = 2
                                    XWS.Cells(R + (Pos), 10).VALUE = ORDR_PRICE_QTY(0)
                                    XWS.Cells(R + (Pos), 11).VALUE = "Order Price"
                                Case Is = 1
                                    XWS.Cells(R + (Pos), 10).VALUE = ""
                                    XWS.Cells(R + (Pos), 11).VALUE = ""
                            End Select
                        Else
                            If ShowPriceLevel(DISC_CODE, I, IsDiscontinued) Then
                                If chkPBFE.Checked Then
                                    If PRICE = 0 Then
                                        XWS.Cells(R + (Pos - 1), 10).VALUE = ""
                                        XWS.Cells(R + (Pos - 1), 11).VALUE = ""
                                    Else
                                        XWS.Cells(R + (Pos - 1), 10).VALUE = PRICE
                                        XWS.Cells(R + (Pos - 1), 11).VALUE = pricingDesc
                                    End If
                                Else
                                    If chkDiscSheets.Checked Then
                                        XWS.Cells(R + (Pos - 1), 10).VALUE = STYLE_PRICE * (1 - (numDiscSheets.Value / 100))
                                        XWS.Cells(R + (Pos - 1), 11).VALUE = "Sale Price"

                                    Else
                                        XWS.Cells(R + (Pos - 1), 10).VALUE = PRICE
                                        XWS.Cells(R + (Pos - 1), 11).VALUE = pricingDesc
                                    End If

                                End If
                            End If
                        End If
                    Next
                    If UseOrderPricing Then
                        XWS.Cells(R, 10).VALUE = CDec(rowICTSTYL1.Item("STYLE_PRICE") & "")
                        XWS.Cells(R, 11).VALUE = "List Price"
                    Else
                        If chkDiscSheets.Checked Then
                            XWS.Cells(R + 3, 10).VALUE = String.Format("=SUM(J{0}*N{0})", R)
                            XWS.Cells(R + 3, 11).VALUE = "Total Available Cost"
                            XWS.Cells(R + 4, 10).VALUE = String.Format("=SUM(H{0}*N{1}/H{2})", R + 4, R, R + 1)
                            XWS.Cells(R + 4, 11).VALUE = "Total Available Cube"
                            totLines.Add(R + 3)
                        Else
                            XWS.Cells(R + 4, 10).VALUE = CDec(Val(rowICTSTYL1.Item("STYLE_PRICE") & ""))
                            XWS.Cells(R + 4, 11).VALUE = "List Price"
                        End If
                    End If

                End If
            End If

            rng = XWS.Range("J" & CStr(R) & ":" & "J" & CStr(R + 4))
            rng.Font.Bold = True
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
            rng.Style = "Currency"
            If chkDiscSheets.Checked Then
                rng = XWS.Range("J" & CStr(R + 4) & ":" & "J" & CStr(R + 4))
                rng.Style = "Normal"
                rng.NumberFormat = "###,###,##0.00"
                rng.Font.Bold = True
            End If

            If UseOrderPricing Then
                rng = XWS.Range("J" & CStr(R + 3) & ":" & "J" & CStr(R + 3))
                rng.NumberFormat = "###,###,##0"
            End If

            ASCMAIN1.sql = "SELECT * FROM ICTSTDQ2 WHERE WHSE_CODE = :PARM1" & vbCrLf _
            & " AND STYLE_CODE = :PARM2 AND COLOR_CODE = :PARM3"
            Dim rowICTSTDQ2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            If rowICTSTDQ2 IsNot Nothing Then
                For I As Integer = 1 To 4
                    Dim DATE_A As String = ""
                    Dim QTY_A As String = ""
                    Dim DC As String = "DATE_" & I.ToString
                    Dim QC As String = "QTY_" & I.ToString
                    If Val(rowICTSTDQ2.Item("QTY_" & I.ToString) & String.Empty) <> 0 Then
                        If rowICTSTDQ2.Item(DC) & "" <> "" Then
                            Dim dateAvailable As Date = rowICTSTDQ2.Item(DC)
                            DATE_A = dateAvailable.ToString("d MMM")
                            If I = 1 Then
                                If FilterAlloc(STYLE_CODE, COLOR_CODE, WHSE_CODE) = True And Val(rowICTSTDQ2.Item(QC) & "") = 0 Then
                                    QTY_A = 0
                                Else
                                    QTY_A = Val(rowICTSTDQ2.Item(QC) & "").ToString
                                End If
                            Else
                                QTY_A = Val(rowICTSTDQ2.Item(QC) & "").ToString
                            End If
                        End If
                        XWS.Cells(R + (I - 1), 13).VALUE = DATE_A
                        XWS.Cells(R + (I - 1), 14).VALUE = QTY_A
                    Else
                        XWS.Cells(R + (I - 1), 13).VALUE = ""
                        XWS.Cells(R + (I - 1), 14).VALUE = ""
                    End If
                Next
            Else
                For I As Integer = 1 To 4
                    XWS.Cells(R + (I - 1), 13).VALUE = ""
                    XWS.Cells(R + (I - 1), 14).VALUE = ""
                Next
            End If

            rng = XWS.Range("M" & CStr(R) & ":" & "M" & CStr(R + 4))
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            rng = XWS.Range("N" & CStr(R) & ":" & "N" & CStr(R + 4))
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
            rng.Font.Bold = True

            Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""

            If rowICTSTYC1 IsNot Nothing AndAlso rowICTSTYC1.Item("STYLE_COLOR_IMAGE_NAME") & "" <> "" Then
                IMAGE_NAME = rowICTSTYC1.Item("STYLE_COLOR_IMAGE_NAME") & ""
            End If

            'If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            IMAGE_NAME = STYLE_CODE & "-" & COLOR_CODE & ".JPG"
            'End If

            If IMAGE_NAME = "" Then
                '   IMAGE_NAME = GetImageLocation(STYLE_CODE, COLOR_CODE)
                'IMAGE_NAME = STYLE_CODE & COLOR_CODE & ".JPG"
            End If

            If IMAGE_NAME <> "" Then
                Dim FILENAME As String = IMAGES_FOLDER & "\" & IMAGE_NAME
                If chkWebImages.Checked Then
                    getWebImage(FILENAME, STYLE_CODE, COLOR_CODE)
                End If
                rng = XWS.Range("A" & CStr(R) & ":" & "B" & CStr(R + 4))
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    InsertPictureInRange(FILENAME, rng, XWS, STYLE_CODE, COLOR_CODE)
                Else
                    rng.MergeCells = True
                    rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    rng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                    rng.FormulaR1C1 = "No Image Available"
                    rng.Font.Bold = True
                End If
            End If

            For Each cellSet As String In New String() {"D:E", "G:H", "J:K", "M:N"}
                Dim xlCells() As String = Split(cellSet, ":")
                rng = XWS.Range(xlCells(0) & CStr(R) & ":" & xlCells(1) & CStr(R + 4))
                If myExcelHasBalls Then
                    rng.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                End If
                rng = XWS.Range(xlCells(0) & CStr(R) & ":" & xlCells(0) & CStr(R + 4))
                If myExcelHasBalls Then
                    With rng.Interior
                        .Color = Color.FromArgb(242, 242, 242)
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                End If
            Next

            R += 6
            currentRow += 1

            If ((currentPage - 1) * 13) + 12 = currentRow - 1 Then
                'at current column widths trying to show more than 12 records on the first page
                'or 13 on subsequent pages scales the record down and creats white space pn the right margin
                xlPages.Add(currentPage, (R - 1))
                currentPage += 1
            End If

        Next

        If chkDiscSheets.Checked Then
            '"=SUM(J11 + J16)"
            Dim T1 As String = "=SUM("
            Dim T2 As String = "=SUM("
            Dim tcnt As Integer = 0
            For Each N1 As Integer In totLines
                tcnt += 1
                If tcnt = 1 Then
                    T1 = T1 + "J" + N1.ToString
                    T2 = T2 + "J" + (N1 + 1).ToString
                Else
                    T1 = T1 + " + J" + N1.ToString
                    T2 = T2 + " + J" + (N1 + 1).ToString
                End If
            Next
            T1 = T1 + ")"
            T2 = T2 + ")"

            rng = XWS.Range("J6:J6")
            rng.Value = "Total Price"
            rng.Font.Bold = True

            rng = XWS.Range("J7:J7")
            rng.Value = "Total Cube"
            rng.Font.Bold = True

            rng = XWS.Range("K6:K6")
            rng.Value = T1
            rng.Style = "Currency"
            rng.Font.Bold = True
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight

            rng = XWS.Range("K7:K7")
            rng.Value = T2
            rng.Font.Bold = True
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
        End If

        Dim msgs As Dictionary(Of String, String) = TAC.TACMAIN1.getSalesDocMsgs("")

        'Not doing this for attribute printing per Andy -4/9/26 W.R.
        ''CC Notice
        'R += 1
        'rng = XWS.Range($"A{R}:N{R}")
        'rng.Merge()
        'rng.Value = "We accept MasterCard, Visa, and Discover. Credit cards are charged approximately one week prior to shipment for the product and estimated shipping charges. Any difference at the time of shipment will be charged or credited to the same card. Each shipment will be charged separately."
        'rng.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
        'rng.Font.Bold = True
        'rng.Font.Color = Color.Red
        'rng.RowHeight = rng.RowHeight * 2
        'rng.WrapText = True

        '2025 Tariff Notice
        If msgs("T").Length > 0 Then
            R += 1
            rng = XWS.Range($"A{R}:N{R}")
            rng.Merge()
            rng.Value = msgs("T")
            rng.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
            rng.Font.Bold = True
            rng.Font.Color = Color.Red
            rng.RowHeight = rng.RowHeight * 2
            rng.WrapText = True
        End If

        rng = XWS.Range("D:E")
        rng.EntireColumn.AutoFit()
        rng = XWS.Range(XC(11))
        rng.EntireColumn.ColumnWidth = 20
        rng = XWS.Range("J:J")
        rng.EntireColumn.AutoFit()
        XWS.PageSetup.PrintTitleRows = "$9:$9"
        XWS.PageSetup.Zoom = False
        XWS.PageSetup.FitToPagesWide = 1
        XWS.PageSetup.FitToPagesTall = False
        If myExcelHasBalls Then
            XWS.PageSetup.ScaleWithDocHeaderFooter = True
            XWS.PageSetup.AlignMarginsHeaderFooter = True
        End If

        XWS.PageSetup.LeftFooter = "Page &P of &N"
        XWS.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlPageBreakPreview

        For i As Integer = 1 To XWS.HPageBreaks.Count
            If i > XWS.HPageBreaks.Count Then Exit For
            If xlPages.ContainsKey(i) Then
                rng = XWS.Range("A" & xlPages(i).ToString)
                Try
                    XWS.HPageBreaks(i).Location = rng
                Catch ex As Exception
                    If ex.Message <> "Exception from HRESULT: 0x800A03EC" Then
                        MsgBox(ex.Message, vbOKOnly, "An Error Occurred")
                    End If
                End Try
            End If
        Next

        'not needed right now
        'rng = XWS.Range("O:O")
        'For i = 1 To XWS.VPageBreaks.Count
        '    XWS.VPageBreaks(i).Location = rng
        'Next

        XWS.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlNormalView

        Dim wsName As String = IIf(txtQueryDesc.Text & "" = "", "Custom Query", txtQueryDesc.Text)
        XWS.Name = wsName

        For Each ews As Excel.Worksheet In XWB.Sheets
            If ews.Name <> XWS.Name Then
                ews.Delete()
            End If
        Next

        Dim xlsFileName_sfx As String = ""
        Dim xlsFileName As String = ""
        Dim xlsControlNo As String = ASCMAIN1.Next_Control_No("ICFATTR2.XLS_NO")
        FILE_NAME = "Best"
        Do
            Try
                xlsFileName = FILE_NAME & "_" & xlsControlNo
                progressSplash.UpdateProgress("Saving Spreadsheet", xlsFileName & xls_format, "", pbInt)
                excelFile = rbadDir & xlsFileName & xls_format
                XWB.SaveAs(excelFile)

                'XWB.SaveXls(excelFile)
                'XWB.ClosePreservedXlsx()
                'XWB = Nothing

                progressSplash.UpdateProgress("", "", "Done", 100)
                xlsFileName_sfx = ""



            Catch ex As Exception
                xlsFileName_sfx = CStr(Val(xlsFileName_sfx) + 1)
            End Try
        Loop While xlsFileName_sfx <> "" And Val(xlsFileName_sfx) < 10

        XWB.Close()
        XWB = Nothing
        excel = Nothing
        Return xlsFileName
    End Function

    Private Sub getWebImage(ByRef FILENAME As String,
                            ByVal STYLE_CODE As String,
                            ByVal COLOR_CODE As String)
        Dim WEBURL As String = "https://www.regency-rib.com/media/product/"
        Dim FILEURL As String = WEBURL & STYLE_CODE & "-" & COLOR_CODE & ".jpg"
        Dim TMP_FOLDER As String = ASCMAIN1.Folders("Temp")
        Dim TMP_FILE As String = TMP_FOLDER & STYLE_CODE & "-" & COLOR_CODE & ".jpg"
        If Not TMP_FOLDER.EndsWith("\") Then
            TMP_FOLDER = TMP_FOLDER & "\"
        End If
        If IO.Directory.Exists(TMP_FOLDER) Then
            Try
                Dim web_client As New Net.WebClient
                Dim image_stream As New MemoryStream(web_client.DownloadData(FILEURL))
                Dim img As Image = Image.FromStream(image_stream)
                If IO.File.Exists(TMP_FILE) Then
                    IO.File.Delete(TMP_FILE)
                End If
                img.Save(TMP_FILE)
                FILENAME = TMP_FILE
            Catch ex As Exception

            End Try
        End If
    End Sub

    Sub Add_Branch_Info(xlws As Excel.Worksheet, dr As DataRow)

        Dim rng As Excel.Range
        Dim emptyRows As Integer = 0
        Dim eal As Integer = 0
        For I As Integer = 1 To 8
            Dim addrLine As String = ""
            rng = xlws.Range("G" & (I - eal).ToString & ":K" & (I - eal).ToString)
            rng.MergeCells = True
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
            Select Case I
                Case 1
                    If dr.Item("BRANCH_NAME") & "" <> "" Then
                        addrLine = dr.Item("BRANCH_NAME") & ""
                        rng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom
                        rng.Font.Bold = True
                    End If
                Case 3, 4, 5
                    addrLine = dr.Item(I - 1) & ""
                Case 6
                    addrLine = dr.Item(I - 1) & " " & dr.Item(I) & ", " & dr.Item(I + 1)
                Case 7
                    Dim phoneFax As String = ""
                    If dr.Item("BRANCH_PHONE") & "" <> "" Then
                        Dim strFormatedNumber As String = CLng(dr.Item("BRANCH_PHONE")).ToString("(###) ###-####")
                        phoneFax = "Phone: " & strFormatedNumber & "   "
                    End If
                    If dr.Item("BRANCH_FAX") & "" <> "" Then
                        Dim strFormatedNumber As String = CLng(dr.Item("BRANCH_FAX")).ToString("(###) ###-####")
                        phoneFax &= "Fax: " & strFormatedNumber & " "
                    End If
                    addrLine = phoneFax
                Case 8
                    If dr.Item("BRANCH_EMAIL") & "" <> "" Then
                        addrLine = "Email: " & dr.Item("BRANCH_EMAIL") & ""
                    End If
            End Select
            If addrLine = "" Then
                eal += 1
            End If
            rng.FormulaR1C1 = addrLine
        Next

    End Sub

    Sub Create_Worksheet_Headers(xlws As Excel.Worksheet)

        Dim rng As Excel.Range

        Format_Worksheet_Header("Image", xlws, xlws.Range("A9:B9"))
        xlws.Cells(9, 1).VALUE = "Image"
        rng = xlws.Range(XC(3))
        rng.EntireColumn.ColumnWidth = 0.5

        Format_Worksheet_Header("Description", xlws, xlws.Range("D9:E9"))
        xlws.Cells(9, 4).VALUE = "Description"
        rng = xlws.Range(XC(6))
        rng.EntireColumn.ColumnWidth = 0.5

        Format_Worksheet_Header("Packing", xlws, xlws.Range("G9:H9"))
        xlws.Cells(9, 7).VALUE = "Packing"
        rng = xlws.Range(XC(9))
        rng.EntireColumn.ColumnWidth = 0.5

        Format_Worksheet_Header("Pricing", xlws, xlws.Range("J9:K9"))
        xlws.Cells(9, 10).VALUE = "Pricing"
        rng = xlws.Range(XC(12))
        rng.EntireColumn.ColumnWidth = 0.5

        Format_Worksheet_Header("Availability", xlws, xlws.Range("M9:N9"))
        xlws.Cells(9, 13).VALUE = "Availability"
        rng = xlws.Range(XC(15))
        rng.EntireColumn.ColumnWidth = 0.5

    End Sub

    Sub Format_Worksheet_Header(headerText As String, xlws As Excel.Worksheet, headerRange As Excel.Range)
        With headerRange  ' XWS.Range(XC(i - 1, (S - 1) * 3 + 0), XC(i, (S - 1) * 3 + 2))
            If myExcelHasBalls Then
                .Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
                Dim grd2 As Microsoft.Office.Interop.Excel.LinearGradient
                grd2 = .Interior.Gradient
                Dim cs As Microsoft.Office.Interop.Excel.ColorStop
                cs = grd2.ColorStops.Add(0)
                cs.Color = Color.FromArgb(255, 255, 255)
                cs = grd2.ColorStops.Add(1)
                cs.Color = Color.FromArgb(79, 129, 189)
                cs.TintAndShade = 0
                grd2.Degree = 90
            Else
                .Interior.Color = Color.FromArgb(79, 129, 189)
            End If
            .Merge()
            .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            With .Font
                .Name = "Calibri"
                .Size = 11
                .Bold = True
            End With

        End With

    End Sub

    Function Generate_Zip_file() As String

        Dim fileDateTime As String = ""
        Dim pbInt As Integer = 0

        If dst.Tables("ICTSTYL1").Select("SEL='1'").Length = 0 Then
            MsgBox("No Styles Selected", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
            Return fileDateTime
        End If

        'need to set filename
        fileDateTime = DateTime.Now.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HHmmss")

        If My.Computer.FileSystem.FileExists(rbadDir & fileDateTime & ".zip") Then
            My.Computer.FileSystem.DeleteFile(rbadDir & fileDateTime & ".zip")
        End If

        Try
            Dim Zip1 As New nsoftware.IPWorksZip.Zip
            Zip1.RuntimeLicense = nSoftwareKeys("nSoftwareZipkey")
            Zip1.ArchiveFile = rbadDir & fileDateTime & ".zip"
            Dim rtp As Integer = dst.Tables("ICTSTYL1").Select("SEL = '1'").Length
            Dim zipBuffer As Integer = 1
            If rtp > 10 Then
                zipBuffer = CInt(rtp * 0.1)
            End If
            Dim rowsToProcess As Integer = rtp + zipBuffer

            Dim currentRow As Integer = 1
            For Each row As DataRow In dst.Tables("ICTSTYL1").Select("SEL = '1'")

                pbInt = CInt((currentRow / rowsToProcess) * 100)
                Dim STYLE_CODE As String = row.Item("STYLE_CODE") & ""
                Dim COLOR_CODE As String = row.Item("COLOR_CODE") & ""
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)

                Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""

                Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 IsNot Nothing AndAlso rowICTSTYC1.Item("STYLE_COLOR_IMAGE_NAME") & "" <> "" Then
                    IMAGE_NAME = rowICTSTYC1.Item("STYLE_COLOR_IMAGE_NAME") & ""
                End If

                If IMAGE_NAME = "" Then
                    ' IMAGE_NAME = GetImageLocation(STYLE_CODE, COLOR_CODE)
                    'IMAGE_NAME = STYLE_CODE & COLOR_CODE & ".JPG"
                End If
                IMAGE_NAME = STYLE_CODE & "-" & COLOR_CODE & ".jpg"

                If IMAGE_NAME <> "" Then
                    Dim FILENAME As String = IMAGES_FOLDER & "\" & IMAGE_NAME
                    If My.Computer.FileSystem.FileExists(FILENAME) Then
                        progressSplash.UpdateProgress("Adding Image to Zip", IMAGE_NAME, "", pbInt)
                        Zip1.IncludeFiles(FILENAME)
                    End If
                End If
                currentRow += 1
            Next
            progressSplash.UpdateProgress("Compressing && Saving Zip", fileDateTime & ".zip", "This may take a moment.", pbInt)
            Zip1.Compress()
            Zip1.Dispose()
            progressSplash.UpdateProgress("Saving Zip", fileDateTime & ".zip", "Done", 100)
        Catch ex As Exception
            MsgBox("Error Creating Zip FileG")
        End Try

        Return fileDateTime & ".zip"

    End Function

    Sub InsertPictureInRange(ByVal PictureFileName As String,
                    ByVal TargetCells As Microsoft.Office.Interop.Excel.Range,
                    ByVal XWS As Microsoft.Office.Interop.Excel.Worksheet,
                    STYLE_CODE As String, COLOR_CODE As String)

        ' inserts a picture and resizes it to fit the TargetCells range
        Dim pp As Microsoft.Office.Interop.Excel.Shape

        If TypeName(XWS) <> "Worksheet" Then Exit Sub
        If Dir(PictureFileName) = "" Then Exit Sub

        pp = XWS.Shapes.AddPicture(PictureFileName,
           Microsoft.Office.Core.MsoTriState.msoFalse,
           Microsoft.Office.Core.MsoTriState.msoCTrue,
           TargetCells.Left,
           TargetCells.Top,
           TargetCells.Width,
           TargetCells.Height)
        'Selection.ShapeRange.Item(1)'
        If STYLE_CODE <> "" Then
            'XWS.Hyperlinks.Add(Anchor:=pp, Address:=
            '    "http://api.regency-rib.com:8181/images/product/" & STYLE_CODE & "-" & COLOR_CODE & ".jpg", ScreenTip:="Click to view image on our Web-Site")
            XWS.Hyperlinks.Add(Anchor:=pp, Address:="https://www.regency-rib.com/media/product/" & STYLE_CODE & "-" & COLOR_CODE & ".jpg", ScreenTip:="Click to view image on our Web-Site")
        End If
        ', TextToDisplay:="text"

        'pp = XWS.Shapes.AddPicture(PictureFileName, _
        '   0, _
        '   1, TargetCells.Left, TargetCells.Top, TargetCells.Width, TargetCells.Height)
        pp.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize
        pp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse
        pp = Nothing
    End Sub

    Function XC(
ByVal C As Int16,
Optional ByVal R As Int16 = 0,
Optional ByVal absolute As Boolean = False) As String

        Dim COL As String = ""
        If C >= 1 Then
            Dim B As Int16 = (C - 1) Mod 26 + 1
            Dim A As Int16 = (C - B) / 26
            COL = Chr(Asc("A") + B - 1)
            If A > 0 Then
                COL = Chr(Asc("A") + A - 1) & COL
            End If
            If absolute Then
                COL = "$" & COL
            End If

            If R = 0 Then
                COL = COL & ":" & COL
            ElseIf R > 0 Then
                COL = COL & IIf(absolute, "$", "") & CStr(R)
            End If
        End If

        Return COL
    End Function

    Sub Set_Pointer(grd As UltraWinGrid.UltraGrid, txt As UltraWinEditors.UltraTextEditor)
        If grd.Rows.Count <> 0 Then
            If txt.Text = "" Then
                grd.ActiveRow = grd.Rows(0)
            Else
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.Cells(0).Value & "" >= txt.Text.ToUpper Then
                        grd.ActiveRow = grow
                        Exit For
                    End If
                Next
            End If
        End If
    End Sub

    Sub Send_Email(ByVal attachments As Dictionary(Of String, String))
        Using frmTAFSEND1 As New TAFSEND1(Me)
            With frmTAFSEND1
                .EMAIL_KEY = "STY"

                Dim rowTATMAIL1 As DataRow = LookUp("TATMAIL1", .EMAIL_KEY)
                Dim rowASTUSER1_EMAIL_FROM As DataRow
                Dim Email_From As String = ""
                If rowTATMAIL1 Is Nothing Then
                    rowASTUSER1_EMAIL_FROM = LookUp("ASTUSER1", ASCMAIN1.USER_ID, True)
                Else
                    If rowTATMAIL1.Item("EMAIL_FROM_USER") & "" = "1" Then
                        rowASTUSER1_EMAIL_FROM = LookUp("ASTUSER1", ASCMAIN1.USER_ID, True)
                    Else
                        rowASTUSER1_EMAIL_FROM = LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_FROM") & "", True)
                    End If
                End If
                'Dim USER_SIGNATURE As String = rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL_SIGNATURE") & ""
                Dim USER_SIGNATURE As String = ""
                Email_From = IIf(rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & "" = "",
                                 "donotreply" & "@" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_DEFAULT_EMAIL_DOMAIN"),
                                 rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & "")

                .SEND_TO = ""
                .SEND_FROM = Email_From
                .SEND_FROM_NAME = rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & ""
                .SEND_CC = ASCMAIN1.USER_EMAIL
                .SEND_CC_NAME = ASCMAIN1.USER_NAME
                .SEND_FROM_SIGNATURE = USER_SIGNATURE

                'If Absx1.txtFor("SREP_CODE").Text <> TAC.TACMAIN1.SREP_CODE Then
                '    ASCMAIN1.sql = "Select * from ASTUSER1 " _
                '    & " where USER_ID = (Select Min (USER_ID) " _
                '    & " from TATUSER1 where SREP_CODE = :PARM1)"
                '    Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("SREP_CODE").Text})
                '    If row IsNot Nothing AndAlso row.Item("USER_EMAIL") & "" <> "" Then
                '        .SEND_CC &= IIf(.SEND_CC = "", "", ";") & row.Item("USER_EMAIL")
                '        .SEND_CC_NAME &= ";" & row.Item("USER_NAME")
                '    End If
                'End If

                .SEND_SUBJECT = MAIL_SUBJECT
                .SEND_BODY = ""

                .SEND_ATTACHMENTs = attachments
                .SEND_METHOD = "E"
                .SEND_ENTITY_CAPTION = "Style List"
                .SEND_ENTITY_TABLE = "ICTSTYL1"
                .SEND_ENTITY_KEY = SCCs(0)
                .SEND_ENTITY_NAME = ""
                .ShowDialog()
                '.Send_email()

                If .SEND_STATUS <> "C" Then
                    'BeginTrans()
                    ''not sure if the event will be recorded
                    'TAC.TACMAIN1.Record_Event("SOTINVH1", INV_NO, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, _
                    '               "E", "Email List To" & .SEND_TO & " - " & .SEND_TO_NAME, _
                    '               .SEND_NO)
                    'CommitTrans("")
                End If

                'Fill_Records("TATCONT1", New String() {"ARTCUST1", Absx1.txtFor("CUST_CODE").Text})

            End With
        End Using

    End Sub

    Private Function GetImageLocation(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As String
        Dim RetVal As String = ""
        'Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
        Dim RO_PARM_STYLE_IMG_DIR As String = IMAGES_FOLDER
        Dim FileMatch As String
        If RO_PARM_STYLE_IMG_DIR.Length > 0 Then
            FileMatch = Dir(String.Format("{0}\{1}-{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
            If FileMatch.Length > 0 Then
                RetVal = String.Format("{0}", FileMatch)
            Else
                FileMatch = Dir(String.Format("{0}\{1}{2}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE, COLOR_CODE))
                If FileMatch.Length > 0 Then
                    RetVal = String.Format("{0}", FileMatch)
                Else
                    FileMatch = Dir(String.Format("{0}\{1}.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                    If FileMatch.Length > 0 Then
                        RetVal = String.Format("{0}.jpg", STYLE_CODE)
                    Else
                        FileMatch = Dir(String.Format("{0}\{1}*", RO_PARM_STYLE_IMG_DIR, STYLE_CODE))
                        If chkDefaultColor.Checked Then
                            If FileMatch.Length > 0 Then
                                RetVal = String.Format("{0}", FileMatch)
                            End If
                        Else
                            If FileMatch.Length > 0 Then
                                RetVal = String.Format("{0}\{1}-XXXX.jpg", RO_PARM_STYLE_IMG_DIR, STYLE_CODE)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return RetVal
    End Function

    Function Does_This_Version_Of_Excel_Have_Balls(versionName As String) As Boolean
        Dim DTVOEHB As Boolean = True
        Dim length As Integer = versionName.IndexOf(".")
        versionName = versionName.Substring(0, length)

        Dim versionNumber As Integer = Integer.Parse(versionName, Globalization.CultureInfo.GetCultureInfo("en-US"))

        If versionNumber < 12 Then
            'No, this version of Excel has no balls - remove all superfluous formatting
            DTVOEHB = False
        Else
            'This version of Excel has balls, but we can simulate a pseudoneuter by unremarking the following line
            'DTVOEHB = False
        End If

        Return DTVOEHB
    End Function

    Private Sub cmdOrder_Click(sender As System.Object, e As System.EventArgs) Handles cmdOrder.Click
        Dim ORDR_NO As String = txtOrder.Text
        If ORDR_NO.Length = 0 Then
            MsgBox("Please Enter An Order Number In The Box Above.", MsgBoxStyle.Critical, "Order Number Required")
        Else
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("Select Count(*) from SOTORDR1 where ORDR_NO = '{0}'", ORDR_NO))
            ASCMAIN1.sql = SQLS.ToString()
            Dim OCNT As Int16 = Val(ASCDATA1.GetDataValue)
            If OCNT = 0 Then
                MsgBox("Order " & ORDR_NO & " Is Not Valid", MsgBoxStyle.Critical, "Bad Order Number")
            Else
                If ORDR_NOs.Count > 0 Then
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Clear Results"
                    Dim iMSG As New System.Text.StringBuilder
                    iMSG.AppendLine("Do You Want To Clear The Results On The Grid Below?")
                    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    If iResult = MsgBoxResult.Yes Then
                        ORDR_NOs.Clear()
                    End If
                End If
                ORDR_NOs.Add(ORDR_NO)
                Call Mode_Settings(True)
                Find_Styles(True)
                showSelectors(False)
                'cmdOrder.Enabled = False
                'txtOrder.ReadOnly = True
                For Each grow As UltraWinGrid.UltraGridRow In grdICTSTYL1.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("SEL").Value = "1"
                    grow.Update()
                Next
            End If
        End If
    End Sub

    Private Function SetExcelSortBy() As String
        Dim Retval As String = "STYLE_CODE, COLOR_CODE"
        Dim SortCount As Integer = grdICTSTYL1.DisplayLayout.Bands(0).SortedColumns.Count
        If SortCount > 0 Then
            Retval = ""
            For i As Integer = 1 To SortCount
                Dim SortCol As String = grdICTSTYL1.DisplayLayout.Bands(0).SortedColumns(i - 1).Key & " "
                Select Case grdICTSTYL1.DisplayLayout.Bands(0).SortedColumns(i - 1).SortIndicator
                    Case Is = UltraWinGrid.SortIndicator.Ascending
                        SortCol += "ASC, "
                    Case Is = UltraWinGrid.SortIndicator.Descending
                        SortCol += "DESC, "
                    Case Else
                        SortCol = ""
                End Select
                Retval += SortCol
            Next
            If Retval.Length > 2 Then
                Retval = Retval.Substring(0, Retval.Length - 2)
            End If
        End If
        Return Retval
    End Function

    Private Sub chkLongColors_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkLongColors.CheckedChanged
        SetColorCodeView()
    End Sub

    Private Sub SetColorCodeView()
        grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("LONG_COLOR").Header.Caption = "Color Long"
        grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("LONG_COLOR").Width = grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Width
        'grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("LONG_COLOR").Lay = grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Width
        If chkLongColors.Checked Then
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("LONG_COLOR").Header.VisiblePosition = grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Header.VisiblePosition + 1
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Hidden = False
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("LONG_COLOR").Hidden = False
        Else
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Header.VisiblePosition = 4
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("COLOR_CODE").Hidden = False
            grdICTSTYL1.DisplayLayout.Bands(0).Columns.Item("LONG_COLOR").Hidden = True
        End If
    End Sub

    Private Function GetVendorData(ByVal VEND_CODE As String, ByVal COLUMN As String) As String
        Dim RetVal As String = ""
        If VEND_CODE.Length > 0 And COLUMN.Length > 0 Then
            ASCMAIN1.sql = String.Format("SELECT {0} FROM APTVEND1 WHERE VEND_CODE = '{1}'", COLUMN, VEND_CODE)
            RetVal = ASCDATA1.GetDataValue
        End If
        Return RetVal
    End Function

    'Public Sub New()
    '    ' This call is required by the designer.
    '    InitializeComponent()

    '    ' Add any initialization after the InitializeComponent() call.
    'End Sub

    Private Sub btnAllDiscontinued_Click(sender As System.Object, e As System.EventArgs) Handles btnAllDiscontinued.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Load All Discontinued"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("This Will Clear Any Existing Searches And")
        iMSG.AppendLine("Load The Grid With All Discontinued Items.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Is That What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Fetching Styles", "")
            Find_Styles(False, True)
            showSelectors(False)
            Call Mode_Settings(True)
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
        End If

    End Sub

    Private Function FilterAlloc(STYLE_CODE As String, COLOR_CODE As String, Optional WHSE_CODE As String = "MS") As Boolean
        Dim RetVal As Boolean = False
        Dim S As New System.Text.StringBuilder() With {.Length = 0}
        S.AppendLine("SELECT")
        S.AppendLine("MSOH, MSFT")
        S.AppendLine("FROM")
        S.AppendLine("  (")
        S.AppendLine("   SELECT")
        S.AppendLine("   C1.STYLE_CODE,")
        S.AppendLine("   C1.COLOR_CODE,")
        S.AppendLine("   C1.STYLE_COLOR_STATUS,")
        S.AppendLine("   CASE WHEN")
        S.AppendLine("   SUM(")
        S.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        S.AppendLine("     ELSE 0")
        S.AppendLine("     END) < 0")
        S.AppendLine("   THEN")
        S.AppendLine("     0")
        S.AppendLine("   ELSE")
        S.AppendLine("   SUM(")
        S.AppendLine("     CASE S2.WHSE_CODE")
        S.AppendLine("     WHEN 'MS'")
        S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
        S.AppendLine("     ELSE 0")
        S.AppendLine("     END)")
        S.AppendLine("   END AS MSOH,")
        S.AppendLine("   CASE WHEN")
        S.AppendLine("   SUM(")
        S.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
        S.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        S.AppendLine("     ELSE 0")
        S.AppendLine("     END) <= 0")
        S.AppendLine("   THEN")
        S.AppendLine("     0")
        S.AppendLine("   ELSE")
        S.AppendLine("     CASE WHEN")
        S.AppendLine("       SUM(")
        S.AppendLine("       CASE S2.WHSE_CODE")
        S.AppendLine("       WHEN 'MS'")
        S.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        S.AppendLine("       ELSE 0")
        S.AppendLine("       END) < 0")
        S.AppendLine("     THEN")
        S.AppendLine("       0")
        S.AppendLine("     ELSE")
        S.AppendLine("     SUM(")
        S.AppendLine("       CASE S2.WHSE_CODE")
        S.AppendLine("       WHEN 'MS'")
        S.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
        S.AppendLine("       ELSE 0")
        S.AppendLine("       END) END")
        S.AppendLine("   END AS MSFT")
        S.AppendLine("   FROM ICTSTYC1 C1")
        S.AppendLine("   LEFT JOIN ICTSTAT2 S2")
        S.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
        S.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
        S.AppendLine("   INNER JOIN ICTCOLR1 C2")
        S.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
        S.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
        S.AppendLine("  )")
        S.AppendLine("  WHERE (STYLE_COLOR_STATUS NOT IN ('D','N') OR (MSOH <> 0) OR (MSFT <> 0))")
        S.AppendLine("  AND STYLE_CODE = :PARM1")
        S.AppendLine("  AND COLOR_CODE = :PARM2")
        Dim tbl As DataTable = ASCDATA1.GetDataTable(S.ToString(), String.Empty, "VV", New Object() {STYLE_CODE, COLOR_CODE})
        For Each rowFUTURE As DataRow In tbl.Rows
            If Val(rowFUTURE.Item("MSOH").ToString & "") = 0 Then
                RetVal = True
            End If
        Next
        Return RetVal
    End Function

    Private Sub chkPBSTANDARD_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkPBSTANDARD.CheckedChanged
        Dim IsChecked = chkPBSTANDARD.Checked
        chkPBFE.Checked = Not IsChecked
        grpPBALL.Visible = IsChecked
        grpNONPVC.Visible = IsChecked
        grpPVC.Visible = IsChecked
        chkPB1.Checked = IsChecked
        chkPB2.Checked = IsChecked
        chkPB3.Checked = IsChecked
        chkPB4.Checked = IsChecked
        chkPBFE.Checked = Not IsChecked
        If IsChecked Then
            chkNonPVC.Checked = IsChecked
            chkPVC.Checked = IsChecked
        End If
    End Sub

    Private Sub chkPBFE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkPBFE.CheckedChanged
        chkPBSTANDARD.Checked = Not chkPBFE.Checked
        panFEExtra.Visible = chkPBFE.Checked
        panFEFD.Visible = chkPBFE.Checked
    End Sub

    Private Sub chkNonPVC_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkNonPVC.CheckedChanged
        Dim isChecked As Boolean = chkNonPVC.Checked
        panNONALWAYS.Visible = Not isChecked
        panEXTRA.Visible = Not isChecked
        If isChecked Then
            optPRICE_TIER.Value = Null
            optDISC_PCT_EXTRA.Value = Null
            numDISC_PCT.Value = Null
        End If
        If Not chkNonPVC.Checked Or Not chkPVC.Checked Then
            grpPBALL.Visible = False
            chkPB1.Checked = True
            chkPB2.Checked = True
            chkPB3.Checked = True
            chkPB4.Checked = True
        Else
            grpPBALL.Visible = True
        End If
    End Sub

    Private Sub chkPVC_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkPVC.CheckedChanged
        Dim isChecked As Boolean = chkPVC.Checked
        panPVCALWAYS.Visible = Not isChecked
        If isChecked Then
            optPRICE_TIER_PVC.Value = Null
        End If
        If Not chkNonPVC.Checked Or Not chkPVC.Checked Then
            grpPBALL.Visible = False
            chkPB1.Checked = True
            chkPB2.Checked = True
            chkPB3.Checked = True
            chkPB4.Checked = True
        Else
            grpPBALL.Visible = True
        End If
    End Sub

    Private Sub optPRICE_TIER_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPRICE_TIER.ValueChanged
        If optPRICE_TIER.Value = "SP" Then
            panEXTRA.Visible = True
            optDISC_PCT_EXTRA.Value = Null
            numDISC_PCT.Value = Null
        Else
            panEXTRA.Visible = False
            optDISC_PCT_EXTRA.Value = Null
            numDISC_PCT.Value = Null
        End If
    End Sub

    Private Sub optDISC_PCT_EXTRA_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optDISC_PCT_EXTRA.ValueChanged
        If optDISC_PCT_EXTRA.Value <> "" Then
            numDISC_PCT.Value = Null
        End If
    End Sub

    Private Sub numDISC_PCT_ValueChanged(sender As System.Object, e As System.EventArgs) Handles numDISC_PCT.ValueChanged
        If Not IsNothing(numDISC_PCT.Value) Then
            optDISC_PCT_EXTRA.Value = Null
        End If
    End Sub

    'Private Function Price_Line(STYLE_CODE As String, COLOR_CODE As String) As Decimal

    '    If Not frm.dst.Tables.Contains("ICTCLAS1") Then
    '        ASCMAIN1.sql = "Select * from ICTCLAS1"
    '        frm.Create_TDA(frm.dst.Tables.Add, "ICTCLAS1", "**", 0, False)
    '        frm.Fill_Records("ICTCLAS1")
    '    End If
    '    If Not frm.dst.Tables.Contains("ICTDISC1") Then
    '        ASCMAIN1.sql = "Select * from ICTDISC1"
    '        frm.Create_TDA(frm.dst.Tables.Add, "ICTDISC1", "**", 0, False)
    '        frm.Fill_Records("ICTDISC1")
    '    End If

    '    Dim rowICTSTYL1 As DataRow = frm.LookUp("ICTSTYL1", New String() {STYLE_CODE})
    '    Dim STYLE_STATUS As String = rowICTSTYL1.Item("STYLE_STATUS") & ""
    '    Dim STYLE_CLASS_CODE As String = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
    '    Dim STYLE_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PRICE") & "")
    '    Dim STYLE_PROMO_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PROMO_PRICE") & "")
    '    Dim CARTON_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
    '    Dim rowICTSTYC1 As DataRow = frm.LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})

    '    Dim STYLE_COLOR_STATUS As String = ""
    '    If rowICTSTYC1 IsNot Nothing Then
    '        STYLE_COLOR_STATUS = rowICTSTYC1.Item("STYLE_COLOR_STATUS") & ""
    '    End If

    '    Dim ORDR_UNIT_PRICE_CALC As Decimal = STYLE_PRICE

    '    ORDR_PRICE_SOURCE = "Q" ' Qty Break using Price Discount Schedule

    '    If STYLE_STATUS = "D" Or STYLE_COLOR_STATUS = "D" Then
    '        ORDR_UNIT_PRICE_CALC = 0.3 * STYLE_PRICE
    '        ORDR_PRICE_SOURCE = "D" ' Discontinued

    '    ElseIf STYLE_PROMO_PRICE <> 0 Then ' And ORDR_QTY >= CARTON_PACK_QTY Then
    '        ORDR_UNIT_PRICE_CALC = STYLE_PROMO_PRICE
    '        ORDR_PRICE_SOURCE = "P" ' Promo

    '    Else
    '        Dim CUST_DISC_PCT_EXTRA As String = rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") & ""
    '        Dim CUST_DISC_CASES As Decimal = 0
    '        Dim rowICTCLAS1 As DataRow = frm.dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)

    '        If rowICTCLAS1 IsNot Nothing Then
    '            Dim DISC_CODE As String = rowICTCLAS1.Item("DISC_CODE") & ""

    '            If DISC_CODE = "NONPVC" And rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "SP" Then
    '                Dim CUST_DISC_PCT As Decimal = Val(rowARTCUST1.Item("CUST_DISC_PCT") & "")
    '                ORDR_UNIT_PRICE_CALC = STYLE_PRICE * (100 - CUST_DISC_PCT) / 100
    '                ORDR_PRICE_SOURCE = "S" ' Special Price

    '            Else
    '                If DISC_CODE = "PVC" Then
    '                    If rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & "" = "2C" Then CUST_DISC_CASES = 2
    '                    If rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & "" = "FC" Then CUST_DISC_CASES = 1
    '                Else
    '                    If rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "FC" Then CUST_DISC_CASES = 1
    '                    If rowARTCUST1.Item("CUST_PRICE_TIER") & "" = "HC" Then CUST_DISC_CASES = 0.5
    '                End If

    '                Dim rowICTDISC1 As DataRow = frm.dst.Tables("ICTDISC1").Rows.Find(DISC_CODE)
    '                If rowICTDISC1 IsNot Nothing Then
    '                    For I As Integer = 1 To 4
    '                        Dim CASES As Decimal = Val(rowICTDISC1.Item("DISC" & CStr(I) & "_CASES") & "")
    '                        If CUST_DISC_CASES <> 0 And CUST_DISC_CASES < CASES And CUST_DISC_CASES * CARTON_PACK_QTY > ORDR_QTY Then
    '                        Else
    '                            If ORDR_QTY >= CASES * CARTON_PACK_QTY Or CASES = 0 Or CUST_DISC_CASES = CASES Then
    '                                Dim PCT As Decimal = Val(rowICTDISC1.Item("DISC" & CStr(I) & "_PCT") & "")
    '                                ORDR_UNIT_PRICE_CALC = STYLE_PRICE * (100 - PCT) / 100
    '                                If ORDR_UNIT_PRICE_CALC < 0 Then ORDR_UNIT_PRICE_CALC = 0
    '                                ' ORDR_PRICE_SOURCE = rowICTDISC1.Item("ABBR" & CStr(I)) & "" ' Qty Break
    '                                ' THIS IS THE PRICE TIER
    '                                ORDR_PRICE_SOURCE = "Q" & CStr(I)
    '                                Exit For
    '                            End If
    '                        End If
    '                    Next

    '                    If DISC_CODE = "NONPVC" Then
    '                        Dim CUST_DISC_PCT_EXTRA_PCT As Decimal = 0
    '                        If CUST_DISC_PCT_EXTRA = "1" Then CUST_DISC_PCT_EXTRA_PCT = 5
    '                        If CUST_DISC_PCT_EXTRA = "2" Then CUST_DISC_PCT_EXTRA_PCT = 10
    '                        If CUST_DISC_PCT_EXTRA_PCT <> 0 And CUST_DISC_CASES = 0 Then

    '                            ORDR_UNIT_PRICE_CALC = (100 - CUST_DISC_PCT_EXTRA_PCT) * ORDR_UNIT_PRICE_CALC / 100
    '                            ORDR_PRICE_SOURCE &= Format(CUST_DISC_PCT_EXTRA_PCT, "00")
    '                            If CUST_DISC_PCT_EXTRA_PCT = 0 Then
    '                                ORDR_PRICE_SOURCE &= "XX" ' SHOULD NEVER HAPPEN SINCE THIS BLOCK IS ONLY IF CUST_DISC_PCT_EXTRA_PCT <> 0
    '                            Else
    '                                ORDR_PRICE_SOURCE &= "X" & CUST_DISC_PCT_EXTRA
    '                            End If
    '                            ' THIS IS THE VOL DISC
    '                        End If
    '                    End If

    '                End If
    '            End If

    '        End If
    '    End If

    '    ORDR_UNIT_PRICE_CALC = System.Math.Round(ORDR_UNIT_PRICE_CALC + 0.001, 2)
    '    Return ORDR_UNIT_PRICE_CALC
    'End Function

    Private Function ShowPriceLevel(ByVal DISC_CODE As String, ByVal I As Integer, ByVal IsDiscontinued As Boolean) As Boolean
        Dim RetVal As Boolean = True
        Dim IsSpecialPricing As Boolean = False
        If IsDiscontinued Then
            If I = 4 Then
                RetVal = True
            Else
                RetVal = False
            End If
            IsSpecialPricing = True
        Else
            If chkPBFE.Checked Then
                'If I = 4 Then
                '    RetVal = True
                'Else
                '    RetVal = False
                'End If
                IsSpecialPricing = True
            Else
                Select Case DISC_CODE
                    Case Is = "PVC"
                        If Not chkPVC.Checked Then
                            IsSpecialPricing = True
                            Select Case optPRICE_TIER_PVC.Value
                                Case Is = "5C"
                                    If I <= 2 Then
                                        RetVal = True
                                    Else
                                        RetVal = False
                                    End If
                                Case Is = "FC"
                                    If I <= 3 Then
                                        RetVal = True
                                    Else
                                        RetVal = False
                                    End If
                            End Select
                        End If
                    Case Is = "NONPVC"
                        If Not chkNonPVC.Checked Then
                            IsSpecialPricing = True
                            Select Case optPRICE_TIER.Value
                                Case Is = "FC"
                                    If I <= 2 Then
                                        RetVal = True
                                    Else
                                        RetVal = False
                                    End If
                                Case Is = "HC"
                                    If I <= 3 Then
                                        RetVal = True
                                    Else
                                        RetVal = False
                                    End If
                                Case Is = "SP"
                                    If Val(numDISC_PCT.Value & "") > 0 Then
                                        If I = 4 Then
                                            RetVal = True
                                        Else
                                            RetVal = False
                                        End If
                                    Else
                                        Select Case optDISC_PCT_EXTRA.Value
                                            Case Is = "2" '10%
                                            Case Is = "1" '5%
                                            Case Else
                                        End Select
                                    End If
                            End Select
                        Else
                            RetVal = True
                        End If
                    Case Else
                        RetVal = True
                End Select
            End If
        End If
        If Not IsSpecialPricing Then
            Select Case I
                Case Is = 1
                    If chkPB4.Checked Then
                        RetVal = True
                    Else
                        RetVal = False
                    End If
                Case Is = 2
                    If chkPB3.Checked Then
                        RetVal = True
                    Else
                        RetVal = False
                    End If
                Case Is = 3
                    If chkPB2.Checked Then
                        RetVal = True
                    Else
                        RetVal = False
                    End If
                Case Is = 4
                    If chkPB1.Checked Then
                        RetVal = True
                    Else
                        RetVal = False
                    End If
            End Select
        End If
        Return RetVal
    End Function

    Private Function GetOrderPrice(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As Double()
        Dim RetVal As Double()
        ReDim RetVal(1)
        Dim SQL As New System.Text.StringBuilder
        Dim ORDR_NO As String = ""
        'If STYLE_CODE = "MTX53273" Then Stop
        For Each order As String In ORDR_NOs
            SQL.Length = 0
            SQL.AppendLine("SELECT COUNT(*) AS RECCNT")
            SQL.AppendLine("FROM SOTORDR2")
            SQL.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", order))
            SQL.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
            SQL.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQL.ToString()
            Dim RECCNT As Int16 = Val(ASCDATA1.GetDataValue)
            If RECCNT > 0 Then
                ORDR_NO = order
            End If
        Next
        SQL.Length = 0
        SQL.AppendLine("SELECT ORDR_UNIT_PRICE, ORDR_QTY")
        SQL.AppendLine("FROM SOTORDR2")
        SQL.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
        SQL.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
        SQL.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        Dim tbl As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), String.Empty)
        RetVal(0) = 0
        RetVal(1) = 0
        For Each rowSOTORDR2 As DataRow In tbl.Rows
            RetVal(0) = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE").ToString & "")
            RetVal(1) = Val(rowSOTORDR2.Item("ORDR_QTY").ToString & "")
        Next
        Return RetVal
    End Function

    Private Sub AddAllClass()
        'Dim newICTCLAS1 As DataRow = dst.Tables.Item("ICTCLAS1").NewRow
        'newICTCLAS1.Item("STYLE_CLASS_CODE") = "ALL"
        'newICTCLAS1.Item("STYLE_CLASS_DESC") = "All"
        'newICTCLAS1.Item("DISC_CODE") = "NONPVC"
        'dst.Tables.Item("ICTCLAS1").Rows.Add(newICTCLAS1)
    End Sub

    Private Sub btnSelectClass_Click(sender As Object, e As EventArgs) Handles btnSelectClass.Click
        BUILD_SCC()
        Initialize_tabAttributes()
        btnAllDiscontinued.Visible = False
        btnAllActive.Visible = False
        btnSelectClass.Visible = False
    End Sub

    Private Sub btnECOMPRICING_Click(sender As Object, e As EventArgs) Handles btnECOMPRICING.Click
        Dim eMsg As New StringBuilder With {.Length = 0}
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Create ECom Pricing?"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}

        If cboECOMPRICING.Text.Length = 0 Then
            eMsg.AppendLine("ECommerce Partner Not Selected.")
        End If

        If txtPRCG_DESC.Text.Length = 0 Then
            eMsg.AppendLine("Pricing Description Not Selected.")
        End If

        Dim SelOnlyWhere As String = "SEL = '1'"
        Dim SelRows As Int64 = dst.Tables.Item("ICTSTYL1").Select(SelOnlyWhere, "").Count
        If SelRows = 0 Then
            eMsg.AppendLine("You Must Select As Least Style.")
        End If

        If eMsg.Length = 0 Then
            iTitle = "Create ECom Pricing?"
            iMSG.Length = 0
            iMSG.AppendLine("This Will Create A New ECommerce")
            iMSG.AppendLine("Pricing Group For All Of The Selected")
            iMSG.AppendLine("Styles.")
            iMSG.AppendLine("")
            iMSG.AppendLine("Ready?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                createEcomPricing()
            End If
        Else
            iTitle = "Please Fix And Try Again."
            iMSG.Length = 0
            iMSG.AppendLine("This Will Create A New ECommerce")
            iMSG.AppendLine("Pricing Group For All Of The Selected")
            iMSG.AppendLine("Styles.")
            iMSG.AppendLine("")
            iMSG.AppendLine("Ready?")
            iResult = MsgBox(eMsg.ToString, MsgBoxStyle.Critical, iTitle)
        End If
    End Sub

    Private Sub createEcomPricing()
        dst.Tables("ECTPRCG1").Clear()

        Dim PRCG_NO As String = ASCMAIN1.Next_Control_No("ECTPRCG1.PRCG_NO")
        Dim ECOM_CODE As String = cboECOMPRICING.Text

        Dim newECTPRCG1 As DataRow = dst.Tables.Item("ECTPRCG1").NewRow
        newECTPRCG1.Item("PRCG_NO") = PRCG_NO
        newECTPRCG1.Item("PRCG_DESC") = txtPRCG_DESC.Text
        newECTPRCG1.Item("PRCG_STATUS") = "W"
        newECTPRCG1.Item("PRICE_UPDATE") = Null
        newECTPRCG1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        newECTPRCG1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        newECTPRCG1.Item("INIT_DATE") = DATETIME_STAMP
        newECTPRCG1.Item("LAST_DATE") = DATETIME_STAMP
        dst.Tables.Item("ECTPRCG1").Rows.Add(newECTPRCG1)

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine($"'{PRCG_NO}' AS PRCG_NO,")
        S.AppendLine("ECOM_CODE,")
        S.AppendLine("ECOM_NAME,")
        S.AppendLine("ECOM_PRICE_NOTES,")
        S.AppendLine("ECOM_PRICE_LAST,")
        S.AppendLine("ECOM_PRICE_ADD,")
        S.AppendLine("ECOM_PRICE_MARKUP_PCT")
        S.AppendLine("FROM ECTECOM1")
        S.AppendLine($"WHERE ECOM_CODE = '{ECOM_CODE}'")
        Fill_Records("ECTPRCG2",, True, S.ToString)
        For Each rowECTPRCG2 As DataRow In dst.Tables("ECTPRCG2").Select($"ECOM_CODE = '{ECOM_CODE}'")
            Dim COLS As String() = {"ECOM_PRICE_ADD", "ECOM_PRICE_MARKUP_PCT"}
            For Each COL As String In COLS
                If IsDBNull(rowECTPRCG2.Item(COL)) Then
                    rowECTPRCG2.Item(COL) = 0
                End If
            Next
        Next

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine($"'{PRCG_NO}' AS PRCG_NO,")
        S.AppendLine("I1.STYLE_CODE || '-' || C1.COLOR_CODE AS SKU,")
        S.AppendLine($"'{ECOM_CODE}' AS ECOM_CODE,")
        S.AppendLine("I1.STYLE_CODE,")
        S.AppendLine("C1.COLOR_CODE,")
        S.AppendLine("I1.STYLE_STATUS,")
        S.AppendLine("I1.STYLE_DESC,")
        S.AppendLine("NVL(I1.SIZE_CODE,'') AS SIZE_CODE,")
        S.AppendLine("S3.ATTR_DESC,")
        S.AppendLine("I1.CARTON_PACK_QTY AS CASE_QTY,")
        S.AppendLine("I1.STYLE_UOM AS UOM,")
        S.AppendLine("I1.STYLE_CLASS_CODE,")
        S.AppendLine("0 AS SHIP_DROP,")
        S.AppendLine("S2.WHSE_QTY_ON_HAND,")
        S.AppendLine("S2.NET_POS,")
        S.AppendLine("S2.IN_TRANS,")
        S.AppendLine("S2.FUTURE,")
        S.AppendLine("I1.STYLE_PRICE,")
        S.AppendLine("999 AS SET_QTY,")
        S.AppendLine("999.99 AS ECOM_UNIT_PRICE,")
        S.AppendLine("999.99 AS SET_PRICE,")
        S.AppendLine("999.99 AS STANDARD_PRICE,")
        S.AppendLine("999.99 AS STANDARD_SET_PRICE,")
        S.AppendLine("999.99 AS CARTON_SET_PRICE,")
        S.AppendLine("999.99 AS STANDARD_PARTNER_PRICE,")
        S.AppendLine("999.99 AS MANUAL_PARTNER_PRICE,")
        S.AppendLine("999.99 AS FINAL_PARTNER_PRICE")
        S.AppendLine("FROM ICTSTYL1 I1, ICTSTYC1 C1,")
        S.AppendLine("(")
        S.AppendLine("    SELECT")
        S.AppendLine("    S2.STYLE_CODE,")
        S.AppendLine("    S2.COLOR_CODE,")
        S.AppendLine("    SUM(NVL(S2.WHSE_QTY_ON_HAND,0)) AS WHSE_QTY_ON_HAND,")
        S.AppendLine("    SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) - NVL(S2.WHSE_QTY_OPEN,0))) AS NET_POS,")
        S.AppendLine("    SUM(NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0)) AS IN_TRANS,")
        S.AppendLine("    SUM(NVL(S2.WHSE_QTY_ON_HAND,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0)) AS FUTURE")
        S.AppendLine("    FROM ICTSTAT2 S2")
        S.AppendLine("    WHERE S2.WHSE_CODE = 'MS'")
        S.AppendLine("    GROUP BY")
        S.AppendLine("    S2.STYLE_CODE,")
        S.AppendLine("    S2.COLOR_CODE,")
        S.AppendLine("    S2.WHSE_CODE")
        S.AppendLine(") S2,")
        S.AppendLine("(")
        S.AppendLine("    SELECT")
        S.AppendLine("    S3.STYLE_CODE,")
        S.AppendLine("    MAX(A1.ATTR_DESC) AS ATTR_DESC")
        S.AppendLine("    FROM ICTSTYL3 S3, ICTATTR1 A1")
        S.AppendLine("    WHERE S3.ATTR_CODE = A1.ATTR_CODE")
        S.AppendLine("    AND NVL(A1.ATT_RANK,'0') = '1'")
        S.AppendLine("    GROUP BY S3.STYLE_CODE")
        S.AppendLine(") S3")
        S.AppendLine("WHERE I1.STYLE_CODE = C1.STYLE_CODE")
        S.AppendLine("AND I1.STYLE_CODE = S2.STYLE_CODE")
        S.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE")
        S.AppendLine("AND I1.STYLE_CODE = S3.STYLE_CODE")
        Dim SelOnlyWhere As String = "SEL = '1'"
        dst.Tables("ECTPRCG3").Clear()
        For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select(SelOnlyWhere)
            Dim SQLECTPRCG3 As String = S.ToString
            Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowICTSTYL1.Item("COLOR_CODE").ToString & String.Empty
            SQLECTPRCG3 = SQLECTPRCG3 & $" AND I1.STYLE_CODE = '{STYLE_CODE}' AND C1.COLOR_CODE = '{COLOR_CODE}'"
            Fill_Records("ECTPRCG3",, False, SQLECTPRCG3.ToString)
        Next
        Call BeginTrans()
        Dim SQLD As String = $" PRCG_NO = '{PRCG_NO}'"
        Update_Record_TDA("ECTPRCG1", SQLD)
        Update_Record_TDA("ECTPRCG2", SQLD)
        Update_Record_TDA("ECTPRCG3", SQLD)
        Call CommitTrans("")
        MsgBox($"Your New Pricing Group {PRCG_NO} Has Been Created.", vbOKOnly, "ECom Pricing!")
    End Sub

    Private Sub btnAllActive_Click(sender As Object, e As EventArgs) Handles btnAllActive.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Load All Active"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("This Will Clear Any Existing Searches And")
        iMSG.AppendLine("Load The Grid With All Active Items.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Is That What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Fetching Styles", "")
            Find_Styles(False, False,,,, True)
            showSelectors(False)
            Call Mode_Settings(True)
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
        End If
    End Sub

    Private Sub btnCatelog_Click(sender As Object, e As EventArgs) Handles btnCatelog.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Load Catalog Page"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("This Will Clear Any Existing Searches And")
        iMSG.AppendLine("Load The Grid With Selected Catalog Page.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Is That What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Dim S As New System.Text.StringBuilder With {.Length = 0}
            S.AppendLine("SELECT")
            S.AppendLine("PAGE_CODE,")
            S.AppendLine("PAGE_NAME")
            S.AppendLine("FROM WBTCATEH")
            With ASCMAIN1.CodeSelector
                .SQL = S.ToString
                .MultipleSelections = False
                .PreviouslySelectedCodes0 = ""
                .Caption = "Select Page"
                .TABLE_NAME = ""
                .VIEW_NAME = ""
                .VIEW_DESC = ""
                .COLUMN_NAME = ""
                .COLUMN_PREKEYs = New Dictionary(Of String, String)
                '.Custom_sql_where = ""
                .tblASTVIEW1 = New DataTable
            End With
            Dim F As New ASFCODE1
            F.ShowDialog()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Dim PAGE_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("PAGE_CODE") & ""
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Fetching Styles", "")
                Find_Styles(False, False,,,,, PAGE_CODE)
                showSelectors(False)
                Call Mode_Settings(True)
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("", "")
            End If
        End If
    End Sub
End Class