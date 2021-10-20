Imports Infragistics.Win.UltraWinGrid

Public Class TAFLOCM1

    'Private tblItremsToMove As DataTable = Nothing
    Public WHSE_CODE As String = String.Empty
    Private WHSE_CTN_CTL As String = String.Empty
    Private LOCATION_CODE_TO As String = String.Empty
    Public WHSE_TRAN_NO As String = String.Empty
    Private WHSE_TRAN_LNO As Int32 = 0

    Public LOCATION_CODE_FROM As String = String.Empty
    Public ADJ_NO As String = String.Empty

    Public confirm_only As Boolean = False
    Public movement_type As String
    Public rowICTWHSE1 As DataRow

    Private Const InvalidLocation As String = "L"
    Private Const InvalidItem As String = "I"
    Private Const InvalidStyleColor As String = "S"

    Public BAR_CODE_CMB As String
    Public rowWHTMOVE2_Copy As DataRow

    Public REASON_CODE As String

    Private frmDst As DataSet
    Dim tagCR As Boolean = False

    Dim COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE
    Dim rowICTSTYL1 As DataRow
    Dim LOAD_NO As String
    Dim Add_CaseIDs_Clicked As Boolean = False

    Public disableUpdate As Boolean = False


#Region "Form Events"

    Public Sub New()
        frmASFBASE1 = New ABSolution.ASFBASE1
        InitializeComponent()

        dst = frmASFBASE1.clsASCBASE1.dst

        With dst
            frmASFBASE1.Create_TDA(.Tables.Add, "ICTIXFR1", "*")
            frmASFBASE1.Create_TDA(.Tables.Add, "ICTIXFR2", "*")

            frmASFBASE1.Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            frmASFBASE1.Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            frmASFBASE1.Create_TDA(.Tables.Add, "WHTMOVE3", "*")

            frmASFBASE1.Create_TDA(.Tables.Add, "WHTBARC1", "*")
            frmASFBASE1.Create_TDA(.Tables.Add, "WHTBARC0", "*")

            frmASFBASE1.Create_TDA(.Tables.Add, "ICTIADJ1", "*")
            frmASFBASE1.Create_TDA(.Tables.Add, "ICTIADJ2", "*")

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            frmASFBASE1.Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            frmDst = dst

            With .Tables.Add("ICTIADJS")
                .Columns.Add("STYLE_CODE", GetType(System.String))
                .Columns.Add("COLOR_CODE", GetType(System.String))
                .Columns.Add("ADJ_QTY", GetType(System.Int64))
                .PrimaryKey = New DataColumn() { .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With
        End With

        With dst.Tables("WHTMOVE3")
            .Columns.Add("STYLE_DESC", GetType(System.String))
        End With

        With dst.Tables("WHTMOVE2")
            .Columns.Add("STYLE_DESC", GetType(System.String))
            .Columns.Add("ERROR_CODES", GetType(System.String))
            .Columns.Add("WHSE_TRAN_QTY_ORIG", GetType(System.Int32))
        End With
    End Sub

    Private Sub TAFLOCM1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        ASCMAIN1.grdInitializeLayout(grdWHTMOVE2)
        grdWHTMOVE2.DataSource = frmDst.Tables("WHTMOVE2")
        ASCMAIN1.grdInitializeLayout(grdWHTMOVE3)
        grdWHTMOVE3.DataSource = frmDst.Tables("WHTMOVE3")
        ASCMAIN1.grdInitializeLayout(grdWHTBARC1)
        grdWHTBARC1.DataSource = frmDst.Tables("WHTBARC1")
        grpTransfer_To.Visible = False

        Create_Summary(grdWHTMOVE2, "WHSE_TRAN_LNO", "Count")
        Create_Summary(grdWHTMOVE2, "WHSE_TRAN_QTY")

        Create_Summary(grdWHTMOVE3, "STYLE_CODE", "Count")
        Create_Summary(grdWHTMOVE3, "CASE_QTY")
        Create_Summary(grdWHTBARC1, "BAR_CODE", "Count")

        WHSE_CTN_CTL = rowICTWHSE1.Item("WHSE_CTN_CTL") & ""

        If confirm_only Then
            grpTo.Visible = False
            With grdWHTMOVE2.DisplayLayout.Bands(0)
                .Columns("LOCATION_CODE_TO").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("WHSE_TRAN_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
            End With
        End If

        If WHSE_CTN_CTL = "C" Then

            grdWHTMOVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grdWHTMOVE2.DisplayLayout.Bands(0)
                .Columns("LOCATION_CODE_TO").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("WHSE_TRAN_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("BAR_CODE").Hidden = False
            End With
        End If

        If movement_type = "CTN" Then
            Me.Text = "Re-Cartonize"

            grpREASON_CODE.Visible = True

            grdWHTMOVE2.Text = "Inventory to Consume"
            btnMove.Text = "Re-Cartonize"

            grdWHTMOVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                rowWHTMOVE2.Item("LOCATION_CODE_TO") = rowICTWHSE1.Item("WHSE_LOC_CTN")
            Next

            Get_LOAD_NO()

        ElseIf movement_type = "CFG" Then
            Me.Text = "Re-Configure"

            grpREASON_CODE.Visible = True

            Me.grpTo.Visible = False
            splLPN.Panel2Collapsed = True

            grdWHTMOVE2.Text = "Inventory to Consume"
            btnMove.Text = "Re-Configure"

            grdWHTBARC1.Visible = False
            cmdAddStyle.Visible = True
            cmdSizeScale.Visible = True

            grdWHTMOVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("LOCATION_CODE_TO").Hidden = True

            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                rowWHTMOVE2.Item("LOCATION_CODE_TO") = rowICTWHSE1.Item("WHSE_LOC_CTN")
            Next

            Get_LOAD_NO()

        ElseIf movement_type = "CCS" Then
            Me.Text = "Create Cases"
            grpREASON_CODE.Visible = True
            btnMove.Text = "Create Cases"
            Get_LOAD_NO()

        ElseIf movement_type = "LNF" Then
            Me.Text = "Move to LNF"

        ElseIf movement_type = "ADJ" Then
            Me.Text = "Location Adjustment"

            grpTo.Visible = False
            grpREASON_CODE.Visible = True
            If REASON_CODE <> "" Then
                Absx1.txtFor("REASON_CODE").Text = REASON_CODE
            End If
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("LOCATION_CODE_TO").Hidden = True
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("WHSE_TRAN_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit

            grdWHTMOVE2.Text = "Inventory to Adjust (Negative Qty will Reduce On Hand)"
            btnMove.Text = "Update"

            ' grdWHTMOVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                rowWHTMOVE2.Item("LOCATION_CODE_TO") = DBNull.Value
                rowWHTMOVE2.Item("WHSE_TRAN_QTY") = -1 * Val(rowWHTMOVE2.Item("WHSE_TRAN_QTY"))
                rowWHTMOVE2.Item("WHSE_TRAN_QTY_ORIG") = -1 * Val(rowWHTMOVE2.Item("WHSE_TRAN_QTY_ORIG"))
            Next

        ElseIf movement_type = "BTS" Then
            Me.Text = "Back-to-Stock"
            grpREASON_CODE.Visible = False
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("BAR_CODE_OTHER").Hidden = False
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("LOCATION_CODE_TO").Hidden = True
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("WHSE_TRAN_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            btnMove.Text = "Create BTS Cases"

            Get_LOAD_NO()
            ScreenMode = True

        ElseIf movement_type = "CHG" Then
            Me.Text = "Change Style"

            grpREASON_CODE.Visible = True
            splWHTMOVEX.Panel1Collapsed = True
            grpSTYLE_CODE.Visible = True

            btnMove.Text = "Change Style"

            grdWHTMOVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            For Each control As Control In New Control() {grdWHTBARC1, lblBAR_CODE2, txtBAR_CODE2, btnAddCaseIDs}
                control.Visible = False
            Next

            Get_LOAD_NO()


        ElseIf movement_type = "CMB" Then
            Me.Text = "Combine Cases into " & BAR_CODE_CMB

            grpREASON_CODE.Visible = False
            ' splWHTMOVEX.Panel1Collapsed = True

            btnMove.Text = "Combine Cases"

            grdWHTMOVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("LOCATION_CODE_TO").Hidden = True
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("BAR_CODE_OTHER").Hidden = False
            grdWHTMOVE2.DisplayLayout.Bands(0).Columns("BAR_CODE_OTHER").CellActivation = Activation.NoEdit
            'For Each control As Control In New Control() {grdWHTBARC1, lblBAR_CODE2, txtBAR_CODE2, btnAddCaseIDs}
            '    control.Visible = False
            'Next
        ElseIf movement_type = "TRN" Then
            Me.Text = "Transfer Cases"
            btnMove.Text = "Transfer"
            txtWhse_Code_From.Text = WHSE_CODE
            grpTransfer_To.Visible = True
            grpREASON_CODE.Visible = False
            grpTo.Visible = False
            grdWHTMOVE2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            With grdWHTMOVE2.DisplayLayout.Bands(0)
                .Columns("LOCATION_CODE_TO").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("WHSE_TRAN_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("BAR_CODE").Hidden = False
            End With
        End If

        If disableUpdate Then
            btnMove.Visible = False
        End If
    End Sub

    Sub Get_LOAD_NO()

        LOAD_NO = ASCMAIN1.Next_Control_No("WHTBARC0.LOAD_NO")
        Dim rowWHTBARC0 As DataRow = frmDst.Tables("WHTBARC0").NewRow
        With rowWHTBARC0
            .Item("LOAD_NO") = LOAD_NO
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LOAD_STATUS") = "A"
            .Item("LOCATION_CODE") = LOCATION_CODE_TO
            .Item("TRAN_TYPE") = "M"
            .Item("TRAN_NO") = WHSE_TRAN_NO
            .Item("LOAD_COMMENT") = movement_type
            .Item("LOAD_DATE") = DATETIME_STAMP.Date
        End With
        frmDst.Tables("WHTBARC0").Rows.Add(rowWHTBARC0)
    End Sub

    Private Sub TAFLOCM1_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown

        '' Set the warehouse code
        'If frmDst.Tables("WHTMOVE1").Rows.Count > 0 Then
        '    WHSE_CODE = frmDst.Tables("WHTMOVE1").Rows(0).Item("WHSE_CODE") & String.Empty
        'End If

        Sort_grdColumns(grdWHTMOVE2, "WHSE_TRAN_LNO")
        splLPN.Panel1Collapsed = False
        numUnits.Visible = False
        splWHTMOVEX.Panel2Collapsed = True
        lblLPN_Help.Visible = False
        If movement_type = "CTN" Or movement_type = "CFG" Then
            splWHTMOVEX.Panel2Collapsed = False
        ElseIf movement_type = "CCS" Then
            splWHTMOVEX.Panel2Collapsed = False
            splWHTMOVEX.Panel1Collapsed = True
        ElseIf movement_type = "BTS" Then
            If frmDst.Tables("WHTMOVE2").Rows.Count = 1 Then
                lblLPN_Help.Visible = True
                splWHTMOVEX.Panel2Collapsed = False
                splLPN.Panel1Collapsed = True
                numUnits.Visible = True
            End If
        End If
    End Sub

#End Region

#Region "Form Procedures"

    Public Sub ClearItemsToMove()
        frmDst.Tables("WHTMOVE1").Rows.Clear()
        frmDst.Tables("WHTMOVE2").Rows.Clear()
        ' WHSE_CODE = String.Empty
        WHSE_TRAN_NO = String.Empty
        WHSE_TRAN_LNO = 0
    End Sub

    Public Sub AddItemToMove(ByVal WHSE_CODE As String, _
                               ByVal LOCATION_CODE As String, _
                               ByVal STYLE_CODE As String, _
                               ByVal COLOR_CODE As String, _
                               ByVal BAR_CODE As String, _
                               ByVal LOAD_NO As String, _
                               ByVal QTY As Int32, _
                               Optional LOCATION_CODE_TO As String = "", _
                               Optional BAR_CODE_OTHER As String = "", _
                               Optional LOAD_NO_TO As String = "")

        Try
            WHSE_CODE = WHSE_CODE.Trim
            LOCATION_CODE = LOCATION_CODE.Trim
            STYLE_CODE = STYLE_CODE.Trim
            COLOR_CODE = COLOR_CODE.Trim

            If WHSE_TRAN_NO.Length = 0 Then
                WHSE_TRAN_NO = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

                Dim rowWHTMOVE1 As DataRow = frmDst.Tables("WHTMOVE1").NewRow
                rowWHTMOVE1.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                txtWHSE_TRAN_NO.Text = WHSE_TRAN_NO
                rowWHTMOVE1.Item("WHSE_TRAN_TYPE") = "M"
                rowWHTMOVE1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowWHTMOVE1.Item("WHSE_CODE") = WHSE_CODE
                rowWHTMOVE1.Item("STATUS") = "U"
                rowWHTMOVE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowWHTMOVE1.Item("INIT_DATE") = DATETIME_STAMP
                rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowWHTMOVE1.Item("LAST_DATE") = DATETIME_STAMP
                frmDst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
            End If

            Dim rowWHTMOVE2 As DataRow = frmDst.Tables("WHTMOVE2").NewRow
            With rowWHTMOVE2
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                WHSE_TRAN_LNO += 1
                .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                .Item("LOCATION_CODE_FROM") = LOCATION_CODE
                .Item("LOCATION_CODE_TO") = LOCATION_CODE_TO
                If rowICTWHSE1.Item("WHSE_CTN_CTL") & "" = "C" Then
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("LOAD_NO_FROM") = LOAD_NO
                    .Item("LOAD_NO_TO") = LOAD_NO_TO
                    If BAR_CODE_OTHER <> "" Then
                        .Item("BAR_CODE_OTHER") = BAR_CODE_OTHER
                    End If
                Else
                    .Item("BAR_CODE") = rowICTWHSE1.Item("WHSE_DEF_BAR_CODE")
                    .Item("LOAD_NO_FROM") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                    .Item("LOAD_NO_TO") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                End If

                .Item("WHSE_TRAN_QTY") = QTY
                .Item("WHSE_TRAN_QTY_ORIG") = QTY

                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE

                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("STATUS") = "U"

                .Item("ERROR_CODES") = String.Empty
            End With
            frmDst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

            LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE})
            If cdr Is Nothing Then
                rowWHTMOVE2.Item("ERROR_CODES") &= InvalidLocation
            End If

            LookUp("ICTSTYL1", STYLE_CODE)
            rowWHTMOVE2.Item("STYLE_DESC") = cdr.Item("STYLE_DESC") & String.Empty

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Add Item To Move", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As System.Windows.Forms.Control, COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & WHSE_CODE & "'"
            Case "WHSE_CODE"
                sql_where = "Nvl(WHSE_LOCATOR,0)  = '1' and WHSE_CODE <> '" & txtWhse_Code_From.Text & "'"
        End Select
    End Sub

    Private Sub OverMove()

        frmDst.Tables("WHTMOVE2").AcceptChanges()

        For Each row As DataRow In ASCDATA1.SelectDistinct _
                (frmDst.Tables("WHTMOVE2"), New String() {"STYLE_CODE", "COLOR_CODE", "LOCATION_CODE_FROM"}).Select()
            Dim STYLE_CODE As String = row.Item("STYLE_CODE") & String.Empty
            Dim COLOR_CODE As String = row.Item("COLOR_CODE") & String.Empty
            Dim LOCATION_CODE_FROM As String = row.Item("LOCATION_CODE_FROM") & String.Empty

            Dim sqlWhere As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and LOCATION_CODE_FROM = '" & LOCATION_CODE_FROM & "' AND ERROR_CODES = ''"

            If frmDst.Tables("WHTMOVE2").Select(sqlWhere).Length = 0 Then
                Continue For
            End If

            Dim WHSE_TRAN_QTY_ORIG As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY_ORIG)", sqlWhere) & String.Empty)
            Dim WHSE_TRAN_QTY As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY)", sqlWhere) & String.Empty)

            Dim foreColor As System.Drawing.Color = Color.Black
            If WHSE_TRAN_QTY > WHSE_TRAN_QTY_ORIG Then
                foreColor = Color.Red
            End If

            For Each gridRow As UltraWinGrid.UltraGridRow In grdWHTMOVE2.Rows
                If gridRow.Cells("STYLE_CODE").Value = STYLE_CODE AndAlso _
                    gridRow.Cells("COLOR_CODE").Value = COLOR_CODE AndAlso _
                    gridRow.Cells("LOCATION_CODE_FROM").Value = LOCATION_CODE_FROM Then
                    gridRow.Cells("WHSE_TRAN_QTY").Appearance.ForeColor = foreColor

                    If foreColor = Color.Black Then
                        gridRow.Cells("WHSE_TRAN_QTY").ToolTipText = ""
                    Else
                        gridRow.Cells("WHSE_TRAN_QTY").ToolTipText = "Original Qty was " & WHSE_TRAN_QTY_ORIG & ", over by " & WHSE_TRAN_QTY - WHSE_TRAN_QTY_ORIG
                    End If
                End If
            Next
        Next
    End Sub

#End Region

#Region "Buttons"

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        WHSE_TRAN_NO = ""
        Me.Close()
    End Sub

    Sub Update_Transfer()
        Dim Transfer_No As String = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
        Dim rowICTIXFR1 As DataRow = frmASFBASE1.clsASCBASE1.dst.Tables("ICTIXFR1").NewRow
        With rowICTIXFR1
            .Item("XFR_NO") = Transfer_No
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("WHSE_CODE_TO") = txtWhse_Code_To.Text
            .Item("XFR_DATE") = DATETIME_STAMP
            .Item("XFR_SOURCE") = "E"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("REGISTER_IND") = "0"
            .Item("JOURNAL_IND") = "0"
        End With
        frmASFBASE1.clsASCBASE1.dst.Tables("ICTIXFR1").Rows.Add(rowICTIXFR1)

        For Each rowWHTMOVE2 As DataRow In frmASFBASE1.clsASCBASE1.dst.Tables("WHTMOVE2").Select
            Dim rowICTIXFR2 As DataRow = frmASFBASE1.clsASCBASE1.dst.Tables("ICTIXFR2").NewRow
            With rowICTIXFR2
                .Item("XFR_NO") = Transfer_No
                .Item("XFR_LNO") = rowWHTMOVE2.Item("WHSE_TRAN_LNO")
                .Item("STYLE_CODE") = rowWHTMOVE2.Item("STYLE_CODE")
                .Item("COLOR_CODE") = rowWHTMOVE2.Item("COLOR_CODE")
                .Item("XFR_QTY") = rowWHTMOVE2.Item("WHSE_TRAN_QTY")
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("LOCATION_CODE") = rowWHTMOVE2.Item("LOCATION_CODE_FROM")
                .Item("BAR_CODE") = rowWHTMOVE2.Item("BAR_CODE")
                .Item("LOAD_NO") = rowWHTMOVE2.Item("LOAD_NO_FROM")

                ASCMAIN1.sql = "Select * from ICTSTYL1 " _
                & " where STYLE_CODE = '" & rowWHTMOVE2.Item("STYLE_CODE") & "" & "'"
                Dim rowICTSTYL1 As DataRow = ASCDATA1.GetDataRow
                If rowICTSTYL1 IsNot Nothing Then
                    .Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
                    .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
                    .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                End If
                frmASFBASE1.clsASCBASE1.dst.Tables("ICTIXFR2").Rows.Add(rowICTIXFR2)
            End With
        Next

    End Sub

    Private Sub btnMove_Click(sender As System.Object, e As System.EventArgs) Handles btnMove.Click
        If movement_type = "CCS" Then

        Else
            If frmDst.Tables("WHTMOVE2").Select("").Length = 0 Then
                MsgBox("Nothing to Move", MsgBoxStyle.OkOnly, "No Rows")
                Exit Sub
            End If
        End If

        If movement_type = "CTN" Or movement_type = "CFG" Or movement_type = "CCS" Then
            If frmDst.Tables("WHTMOVE3").Select("").Length = 0 Then
                MsgBox("No Case Pack Defined", MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                Exit Sub
            Else
                If movement_type = "CFG" Then
                Else

                    If frmDst.Tables("WHTMOVE3").Select("").Length > 1 Then
                        MsgBox("Pre-Pack Cases not Supported", MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                        Exit Sub
                    End If
                End If
            End If

            If movement_type = "CFG" Then
            Else

                If frmDst.Tables("WHTBARC1").Select("").Length = 0 Then
                    MsgBox("No New Case IDs Defined", MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                    Exit Sub
                End If
            End If

            For Each row As DataRow In frmDst.Tables("WHTMOVE3").Select("")
                Dim CASE_QTY As Int64 = Val(row.Item("CASE_QTY") & "")
                If CASE_QTY <= 0 Then
                    MsgBox("Invalid Case Qty for Style " & row.Item("STYLE_CODE"), MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                    Exit Sub
                End If
            Next

            For Each rowWHTBARC1 As DataRow In frmDst.Tables("WHTBARC1").Select("")
                Dim BAR_CODE As String = rowWHTBARC1.Item("BAR_CODE")
                If frmASFBASE1.LookUp("WHTBARC1", BAR_CODE) IsNot Nothing Then
                    MsgBox("Invalid Case ID " & BAR_CODE & " - already on File", MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                    Exit Sub
                Else
                    If BAR_CODE <> BAR_CODE.Substring(0, 1) & Format(Val(BAR_CODE.Substring(1)), "".PadLeft(7, "0")) Then
                        MsgBox("Invalid Case ID Format " & BAR_CODE, MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                        Exit Sub
                    End If
                End If
            Next

            If movement_type = "CFG" Then
            Else


                If Absx1.txtFor("LOCATION_CODE").Text = "" Then
                    MsgBox("No Location Defined for New Case IDs", MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                    Exit Sub
                Else
                    Dim rowWHTLOCM1 As DataRow = frmASFBASE1.LookUp("WHTLOCM1", New String() {WHSE_CODE, Absx1.txtFor("LOCATION_CODE").Text})
                    If rowWHTLOCM1 Is Nothing Then
                        MsgBox("Invalid Location Defined for New Case IDs", MsgBoxStyle.OkOnly, "Cannot " & Me.Text)
                        Exit Sub
                    Else
                        If rowWHTLOCM1.Item("LOCATION_USE") & "" = "S" Or rowWHTLOCM1.Item("LOCATION_USE") & "" = "C" Then
                            MsgBox("Cannot use Location " & Absx1.txtFor("LOCATION_CODE").Text & " for New Case IDs", MsgBoxStyle.OkOnly, "Cannot Re-Cartonize")
                            Exit Sub

                        Else

                            If rowWHTLOCM1.Item("LOCATION_SINGLE_LOAD") & "" = "1" Then
                                Dim LOCATION_CODE As String = Absx1.txtFor("LOCATION_CODE").Text

                                ASCMAIN1.sql = "Select Count (*) CASES" & vbCrLf _
                                & " from WHTLOCB1" & vbCrLf _
                                & " where WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
                                & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                                & "   and LOCATION_QTY > 0"
                                Dim CASES_in_location As Int64 = Val(ASCDATA1.GetDataValue & "")
                                If CASES_in_location <> 0 Then
                                    MsgBox("Location " & LOCATION_CODE & " is a Single Load Location and it is not Empty", MsgBoxStyle.OkOnly, "Cannot Move")
                                    Exit Sub
                                End If
                            End If

                        End If
                    End If
                End If
            End If

            If movement_type = "CTN" Then

                Dim UNITS_OUT As Int64 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY)", "") & "")

                Dim CASE_QTY2 As Int64 = Val(frmDst.Tables("WHTMOVE3").Compute("SUM(CASE_QTY)", "") & "")
                Dim CASES As Int64 = frmDst.Tables("WHTBARC1").Rows.Count
                Dim UNITS_IN As Int64 = CASE_QTY2 * CASES
                If UNITS_IN <> UNITS_OUT Then
                    If MsgBox("Units Out = " & CStr(UNITS_OUT) & ", Units In = " & CStr(UNITS_IN) & vbCrLf & vbCrLf _
                                & "Update Anyway?", MsgBoxStyle.YesNo, "Units In Do NOT Equal Units Out") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            End If

            If movement_type = "CFG" Then
                ' all records should have identical number of units out
                Dim UNITS_OUT As Int64 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY)", "") & "")
                Dim CASE_QTY2 As Int64 = Val(frmDst.Tables("WHTMOVE2").Rows.Count)
                Dim UNITS2 As Int64 = Val(frmDst.Tables("WHTMOVE3").Compute("SUM(CASE_QTY)", "") & "")

                Dim UNITS_IN As Int64 = CASE_QTY2 * UNITS2
                If UNITS_IN <> UNITS_OUT Then
                    If MsgBox("Units Out = " & CStr(UNITS_OUT) & ", Units In = " & CStr(UNITS_IN) & vbCrLf & vbCrLf _
                                & "Update Anyway?", MsgBoxStyle.YesNo, "Units In Do NOT Equal Units Out") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If
        End If

        If movement_type = "BTS" Then

            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                Dim BAR_CODE As String = rowWHTMOVE2.Item("BAR_CODE_OTHER") & ""
                If frmASFBASE1.LookUp("WHTBARC1", BAR_CODE) IsNot Nothing Then
                    MsgBox("Invalid Case ID " & BAR_CODE & " - already on File", MsgBoxStyle.OkOnly, "Cannot BTS")
                    Exit Sub
                Else
                    If BAR_CODE <> BAR_CODE.Substring(0, 1) & Format(Val(BAR_CODE.Substring(1)), "".PadLeft(7, "0")) Then
                        MsgBox("Invalid Case ID Format " & BAR_CODE, MsgBoxStyle.OkOnly, "Cannot BTS")
                        Exit Sub
                    End If
                End If

                Dim LOCATION_CODE As String = rowWHTMOVE2.Item("LOCATION_CODE_TO") & ""
                BAR_CODE = rowWHTMOVE2.Item("BAR_CODE")
                If LOCATION_CODE = "" Then
                    MsgBox("No Location Defined for Case ID " & BAR_CODE, MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Exit Sub
                Else
                    Dim rowWHTLOCM1 As DataRow = frmASFBASE1.LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE})
                    If rowWHTLOCM1 Is Nothing Then
                        MsgBox("Invalid Location Defined for Case ID " & BAR_CODE, MsgBoxStyle.OkOnly, "Cannot Proceed")
                        Exit Sub
                    Else
                        If rowWHTLOCM1.Item("LOCATION_USE") & "" = "S" Or rowWHTLOCM1.Item("LOCATION_USE") & "" = "C" Then
                            MsgBox("Cannot use Location " & Absx1.txtFor("LOCATION_CODE").Text & " for Case ID " & BAR_CODE, MsgBoxStyle.OkOnly, "Cannot Proceed")
                            Exit Sub
                        End If
                    End If
                End If

            Next

        End If

        If movement_type = "ADJ" Then
            If frmDst.Tables("WHTMOVE2").Select("ISNULL(WHSE_TRAN_QTY,0) = 0").Length <> 0 Then
                MsgBox("Cannot Adjust with 0 Qty", MsgBoxStyle.OkOnly, "Cannot Adjust")
                Exit Sub
            End If

            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                Dim WHSE_TRAN_QTY As Int64 = Val(rowWHTMOVE2.Item("WHSE_TRAN_QTY"))
                Dim WHSE_TRAN_QTY_ORIG As Int64 = Val(rowWHTMOVE2.Item("WHSE_TRAN_QTY_ORIG"))
                If WHSE_TRAN_QTY_ORIG < 0 Then
                    If WHSE_TRAN_QTY < WHSE_TRAN_QTY_ORIG Then
                        MsgBox("Cannot Adjust Case " & rowWHTMOVE2.Item("BAR_CODE") & " Qty to a Negative Value", MsgBoxStyle.OkOnly, "Cannot Adjust")
                        Exit Sub
                    End If
                Else
                    If WHSE_TRAN_QTY > WHSE_TRAN_QTY_ORIG Then
                        MsgBox("Cannot Adjust Case " & rowWHTMOVE2.Item("BAR_CODE") & " Qty to a Postive Value (presently, it is a negative value)", MsgBoxStyle.OkOnly, "Cannot Adjust")
                        Exit Sub
                    End If
                End If
            Next
        End If


        If movement_type = "CHG" Or movement_type = "ADJ" Or movement_type = "CTN" Or movement_type = "CFG" Or movement_type = "CCS" Then
            If Absx1.txtFor("REASON_CODE").Text = "" Then
                MsgBox("No Reason Defined for Adjustment", MsgBoxStyle.OkOnly, "Cannot Change")
                Exit Sub
            Else
                Dim rowICTREAS1 As DataRow = frmASFBASE1.LookUp("ICTREAS1", New String() {Absx1.txtFor("REASON_CODE").Text})
                If rowICTREAS1 Is Nothing Then
                    MsgBox("Invalid Reason Defined for Adjustment", MsgBoxStyle.OkOnly, "Cannot Update")
                    Exit Sub
                End If
            End If
        End If

        If movement_type = "CHG" Then
            If Absx1.txtFor("STYLE_CODE").Text = "" Then
                MsgBox("No Replacement Style Code Defined", MsgBoxStyle.OkOnly, "Cannot Change")
                Exit Sub
            Else
                Dim rowICTSTYL1 As DataRow = frmASFBASE1.LookUp("ICTSTYL1", New String() {Absx1.txtFor("STYLE_CODE").Text})
                If rowICTSTYL1 Is Nothing Then
                    MsgBox("Invalid Style Code Defined for Change", MsgBoxStyle.OkOnly, "Cannot Change")
                    Exit Sub
                Else
                    If frmDst.Tables("WHTMOVE2").Select("STYLE_CODE ='" & Absx1.txtFor("STYLE_CODE").Text & "'").Length <> 0 Then
                        MsgBox("Cannot Change Style to Same Value", MsgBoxStyle.OkOnly, "Cannot Change")
                        Exit Sub
                    End If
                End If
            End If


            If Absx1.txtFor("LOCATION_CODE").Text = "" Then
                MsgBox("No Location Defined", MsgBoxStyle.OkOnly, "Cannot Proceed")
                Exit Sub
            Else
                Dim rowWHTLOCM1 As DataRow = frmASFBASE1.LookUp("WHTLOCM1", New String() {WHSE_CODE, Absx1.txtFor("LOCATION_CODE").Text})
                If rowWHTLOCM1 Is Nothing Then
                    MsgBox("Invalid Location Defined for New Case ID", MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Exit Sub
                Else
                    If rowWHTLOCM1.Item("LOCATION_USE") & "" = "S" Or rowWHTLOCM1.Item("LOCATION_USE") & "" = "C" Then
                        MsgBox("Cannot use Location " & Absx1.txtFor("LOCATION_CODE").Text & " for New Case ID", MsgBoxStyle.OkOnly, "Cannot Proceed")
                        Exit Sub
                    End If
                End If
            End If

            If frmDst.Tables("WHTMOVE2").Select("ISNULL(WHSE_TRAN_QTY,0) <= 0").Length <> 0 Then
                MsgBox("Cannot Change Style with 0 or Negative Qty", MsgBoxStyle.OkOnly, "Cannot Change")
                Exit Sub
            End If
        End If

        If WHSE_CTN_CTL = "C" And movement_type <> "ADJ" And movement_type <> "CMB" Then
            Dim LOCATION_CODEs As New List(Of String)
            Dim WHSE_CODE_TO As String = WHSE_CODE
            If movement_type = "TRN" Then
                WHSE_CODE_TO = txtWhse_Code_To.Text
            End If
            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                Dim LOCATION_CODE As String = rowWHTMOVE2.Item("LOCATION_CODE_TO")
                If LOCATION_CODE = "" Then
                    MsgBox("Location " & LOCATION_CODE & " has Blank Move to Location", MsgBoxStyle.OkOnly, "Cannot Move")
                    Exit Sub
                Else
                    If Not LOCATION_CODEs.Contains(LOCATION_CODE) Then
                        LOCATION_CODEs.Add(LOCATION_CODE)
                        Dim rowWHTLOCM1 As DataRow = frmASFBASE1.LookUp("WHTLOCM1", New String() {WHSE_CODE_TO, LOCATION_CODE})
                        If rowWHTLOCM1.Item("LOCATION_SINGLE_LOAD") & "" = "1" Then
                            ASCMAIN1.sql = "Select Count (*) CASES" & vbCrLf _
                            & " from WHTLOCB1" & vbCrLf _
                            & " where WHSE_CODE = '" & WHSE_CODE_TO & "'" & vbCrLf _
                            & "   and LOCATION_CODE = '" & LOCATION_CODE & "'" & vbCrLf _
                            & "   and LOCATION_QTY > 0"
                            Dim CASES As Int64 = Val(ASCDATA1.GetDataValue & "")
                            If CASES <> 0 Then
                                MsgBox("Location " & LOCATION_CODE & " is a Single Load Location and it is not Empty", MsgBoxStyle.OkOnly, "Cannot Move")
                                Exit Sub
                            End If
                        End If
                    End If
                End If

            Next
        End If

        EMsg = String.Empty
        DATETIME_STAMP = DateTime.Now + ASCMAIN1.NowTSD

        Dim warnings As String = String.Empty
        Dim processed As New List(Of String)

        If movement_type = "CCS" Then
        Else
            frmDst.Tables("WHTMOVE2").AcceptChanges()
            Dim rowWHTMOVE1 As DataRow = frmDst.Tables("WHTMOVE1").Rows(0)
            rowWHTMOVE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTMOVE1.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowWHTMOVE1.Item("LAST_DATE") = DATETIME_STAMP
        End If


        If movement_type = "ADJ" Or movement_type = "TRN" Then
            ' THESE CHECKS ARE NOT NEC
        Else
            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("", "STYLE_CODE, COLOR_CODE, LOCATION_CODE_FROM, LOCATION_CODE_TO")
                Dim STYLE_CODE As String = rowWHTMOVE2.Item("STYLE_CODE") & String.Empty
                Dim COLOR_CODE As String = rowWHTMOVE2.Item("COLOR_CODE") & String.Empty
                Dim LOCATION_CODE_FROM As String = rowWHTMOVE2.Item("LOCATION_CODE_FROM") & String.Empty
                Dim LOCATION_CODE_TO As String = (rowWHTMOVE2.Item("LOCATION_CODE_TO") & String.Empty).ToString.Trim

                rowWHTMOVE2.Item("LOCATION_CODE_TO") = LOCATION_CODE_TO

                Dim ERROR_CODES As String = (rowWHTMOVE2.Item("ERROR_CODES") & String.Empty).ToString.Trim
                If ERROR_CODES.Length > 0 Then
                    warnings &= vbCr & "Style-Color " & STYLE_CODE & "-" & COLOR_CODE & " will be skipped since it is invalid."
                    Continue For
                End If

                If LOCATION_CODE_TO.Length = 0 Then
                    warnings &= vbCr & "Style-Color " & STYLE_CODE & "-" & COLOR_CODE & " will be skipped since the 'Location To' is empty."
                    Continue For
                End If

                Dim sqlWhere As String = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and LOCATION_CODE_FROM = '" & LOCATION_CODE_FROM & "'  and ERROR_CODES = ''"

                If processed.Contains(STYLE_CODE & "_" & COLOR_CODE & "_" & LOCATION_CODE_FROM & "_" & LOCATION_CODE_TO) Then
                    Continue For
                Else
                    processed.Add(STYLE_CODE & "_" & COLOR_CODE & "_" & LOCATION_CODE_FROM & "_" & LOCATION_CODE_TO)
                End If

                LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE_TO})
                If cdr Is Nothing Then
                    MessageBox.Show("Invalid Move-To Location (" & LOCATION_CODE_TO & ") for Item: " & STYLE_CODE, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim WHSE_TRAN_QTY_ORIG As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY_ORIG)", sqlWhere) & String.Empty)
                Dim WHSE_TRAN_QTY As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY)", sqlWhere) & String.Empty)

                If WHSE_TRAN_QTY > WHSE_TRAN_QTY_ORIG Then
                    warnings &= vbCr & "Style-Color " & STYLE_CODE & "-" & COLOR_CODE & " Original Qty: " & WHSE_TRAN_QTY_ORIG & ", Move Quantity: " & WHSE_TRAN_QTY
                End If

                If movement_type = "CMB" Then
                Else
                    If rowWHTMOVE2.Item("LOCATION_CODE_TO") & String.Empty = rowWHTMOVE2.Item("LOCATION_CODE_FROM") & String.Empty Then
                        warnings &= vbCr & "Style-Color " & STYLE_CODE & "-" & COLOR_CODE & " will be skipped since the To-Location is the same as the From-Location"
                    End If
                End If

                rowWHTMOVE2.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowWHTMOVE2.Item("INIT_DATE") = DATETIME_STAMP
                rowWHTMOVE2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowWHTMOVE2.Item("LAST_DATE") = DATETIME_STAMP
            Next
        End If


        If warnings.Length > 0 Then
            Dim msg As String = "Please review the following warnings before continuing."
            msg &= vbCrLf
            msg &= warnings
            msg &= vbCrLf & vbCrLf
            msg &= "Do you want to continue?"

            If MessageBox.Show(msg, "Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        Else
            If MessageBox.Show("OK to Proceed with this Entry?", "Verification", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If


        If movement_type = "ADJ" Or movement_type = "TRN" Then ' Or movement_type = "CMB" Then
            ' NOT NEC
        Else
            ' Remove lines with errors
            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("ERROR_CODES <> ''")
                rowWHTMOVE2.Delete()
            Next
            If movement_type = "CMB" Then
            Else
                For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("LOCATION_CODE_TO = LOCATION_CODE_FROM")
                    rowWHTMOVE2.Delete()
                Next
            End If

            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("LOCATION_CODE_TO = ''")
                rowWHTMOVE2.Delete()
            Next
        End If

        If WHSE_CTN_CTL = "C" Then
            If movement_type = "CTN" Or movement_type = "CFG" Or movement_type = "CCS" Or movement_type = "ADJ" Or movement_type = "CMB" Or movement_type = "TRN" Then
                ' Stop ' NO NEED
            Else
                Get_LOAD_NO()
                For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                    rowWHTMOVE2.Item("LOAD_NO_TO") = LOAD_NO
                Next
            End If
        End If

        frmDst.Tables("WHTMOVE2").AcceptChanges()
        If movement_type = "CCS" Then

        Else
            If frmDst.Tables("WHTMOVE2").Rows.Count = 0 Then
                MessageBox.Show("There are no Valid Location Movement to process.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        Dim STYLE_CODE_new As String = Absx1.txtFor("STYLE_CODE").Text
        If STYLE_CODE_new <> "" Then
            If LookUp("ICTSTYL1", New String() {STYLE_CODE_new}) Is Nothing Then
                MessageBox.Show("Cannot Find Style record for " & STYLE_CODE_new, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                Dim COLOR_CODE As String = rowWHTMOVE2.Item("COLOR_CODE")
                If LookUp("ICTSTYC1", New String() {STYLE_CODE_new, COLOR_CODE}) Is Nothing Then
                    MessageBox.Show("Cannot Find Style/Color record for " & STYLE_CODE_new & "/" & COLOR_CODE, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next
        End If


        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
            rowWHTMOVE2.SetAdded()
        Next

        If movement_type = "CTN" Then
            Create_Adjustment_CTN()
        ElseIf movement_type = "CFG" Then
            Create_Adjustment_CFG()
        ElseIf movement_type = "CCS" Then
            Create_Adjustment_CCS()
        ElseIf movement_type = "ADJ" Then
            Create_Adjustment()
        ElseIf movement_type = "BTS" Then
            CreateNewWHTBARC1_BTS()
        ElseIf movement_type = "CHG" Then
            CreateAdjustmentCHG()
            'ElseIf movement_type = "CMB" Then
            '    CreateAdjustmentCMB()
        ElseIf movement_type = "TRN" Then
            Update_Transfer()
        End If

        With frmASFBASE1

            Try
                .BeginTrans()

                If movement_type = "TRN" Then
                    .clsASCBASE1.Update_Record_TDA("ICTIXFR1")
                    .clsASCBASE1.Update_Record_TDA("ICTIXFR2")

                    Dim rowICTIXFR1 As DataRow = .clsASCBASE1.dst.Tables("ICTIXFR1").Rows(0)
                    Dim XFR_NO_in As String = rowICTIXFR1.Item("XFR_NO")

                    ASCDATA1.ExecuteSP("ICPIXFRI", "VN", New Object() {XFR_NO_in, 1}, New String() {"XFR_NO_in", "S"})
                    ASCDATA1.ExecuteSP("ICPIXFRG", "V", New Object() {XFR_NO_in}, New String() {"XFR_NO_in"})

                    ASCDATA1.ExecuteSP("WHPLOCB2",
                   "VVV",
                   New String() {"B", XFR_NO_in, ASCMAIN1.SESSION_NO},
                   New String() {"WHSE_TRAN_TYPE_IN", "WHSE_TRAN_NO_IN", "SESSION_NO_IN"})

                    ASCDATA1.ExecuteSP("WHPLOCB2",
                   "VVV",
                   New String() {"C", XFR_NO_in, ASCMAIN1.SESSION_NO},
                   New String() {"WHSE_TRAN_TYPE_IN", "WHSE_TRAN_NO_IN", "SESSION_NO_IN"})

                ElseIf movement_type = "ADJ" Or movement_type = "CHG" Then

                    If movement_type = "CHG" Then
                        .clsASCBASE1.Update_Record_TDA("WHTBARC0")
                        .clsASCBASE1.Update_Record_TDA("WHTBARC1") ' NOT NEC FOR CHG OR CMB

                        .clsASCBASE1.Update_Record_TDA("WHTMOVE1")
                        .clsASCBASE1.Update_Record_TDA("WHTMOVE2")
                        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {WHSE_TRAN_NO, 0, 1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})
                    End If

                    .clsASCBASE1.Update_Record_TDA("ICTIADJ1")
                    .clsASCBASE1.Update_Record_TDA("ICTIADJ2")
                    ICCMAIN1.Shuttle_ADJ_to_ICTTRAN1_SQL(ADJ_NO)

                    ASCDATA1.ExecuteSP("ICPIADJI", "VN", New Object() {ADJ_NO, 1}, New String() {"ADJ_NO_in", "S"})
                    ASCDATA1.ExecuteSP("ICPIADJG", "V", New Object() {ADJ_NO}, New String() {"ADJ_NO_in"})
                    'ICCMAIN1.Update_Adjustment(Me)
                    ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                            New Object() {"A", ADJ_NO, ASCMAIN1.SESSION_NO},
                            New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

                Else

                    If movement_type = "CFG" Then
                    Else
                        .clsASCBASE1.Update_Record_TDA("WHTMOVE1")
                        .clsASCBASE1.Update_Record_TDA("WHTMOVE2")
                    End If


                    If movement_type = "BTS" Or movement_type = "CTN" Or movement_type = "CCS" Then
                        .clsASCBASE1.Update_Record_TDA("WHTBARC1")
                        .clsASCBASE1.Update_Record_TDA("WHTBARC0")
                    End If

                    If movement_type = "CTN" Or movement_type = "CFG" Or movement_type = "CCS" Then
                        If movement_type = "CFG" Then
                        Else
                            .clsASCBASE1.Update_Record_TDA("WHTMOVE3")
                        End If

                        .clsASCBASE1.Update_Record_TDA("ICTIADJ1")
                        .clsASCBASE1.Update_Record_TDA("ICTIADJ2")
                        ICCMAIN1.Shuttle_ADJ_to_ICTTRAN1_SQL(ADJ_NO)

                        ASCDATA1.ExecuteSP("ICPIADJI", "VN", New Object() {ADJ_NO, 1}, New String() {"ADJ_NO_in", "S"})
                        ASCDATA1.ExecuteSP("ICPIADJG", "V", New Object() {ADJ_NO}, New String() {"ADJ_NO_in"})

                        'ICCMAIN1.Update_Adjustment(Me)
                        ASCDATA1.ExecuteSP("WHPLOCB2", "VVV", _
                                New Object() {"A", ADJ_NO, ASCMAIN1.SESSION_NO}, _
                                New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

                    Else
                        If WHSE_CTN_CTL = "C" And movement_type <> "CMB" Then
                            .clsASCBASE1.Update_Record_TDA("WHTBARC0")
                            For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
                                Dim BAR_CODE As String = rowWHTMOVE2.Item("BAR_CODE")
                                Dim LOAD_NO_TO As String = rowWHTMOVE2.Item("LOAD_NO_TO")
                                ASCMAIN1.sql = "Update WHTBARC1 Set LOAD_NO = '" & LOAD_NO_TO & "' where BAR_CODE = '" & BAR_CODE & "'"
                                ASCDATA1.ExecuteSQL()
                            Next
                        End If
                    End If

                    If movement_type = "CFG" Then
                    Else
                        ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {WHSE_TRAN_NO, 0, 1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})
                    End If
                End If

                .CommitTrans("Move successful")
            Catch ex As Exception
                .Rollback(ex.Message)
            End Try
        End With
        Me.Close()
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTMOVE2, "B", "Split")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdARTPYMT3"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Split"
                If grdWHTMOVE2.ActiveRow Is Nothing Then Exit Sub
                Dim rowWHTMOVE2 As DataRow = frmDst.Tables("WHTMOVE2").NewRow

                rowWHTMOVE2.Item("WHSE_TRAN_NO") = grdWHTMOVE2.ActiveRow.Cells("WHSE_TRAN_NO").Value
                WHSE_TRAN_LNO += 1
                rowWHTMOVE2.Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                rowWHTMOVE2.Item("LOCATION_CODE_FROM") = grdWHTMOVE2.ActiveRow.Cells("LOCATION_CODE_FROM").Value
                rowWHTMOVE2.Item("LOCATION_CODE_TO") = grdWHTMOVE2.ActiveRow.Cells("LOCATION_CODE_TO").Value
                rowWHTMOVE2.Item("BAR_CODE") = grdWHTMOVE2.ActiveRow.Cells("BAR_CODE").Value
                rowWHTMOVE2.Item("WHSE_TRAN_QTY") = 0 'grdWHTMOVE2.ActiveRow.Cells("WHSE_TRAN_NO").Value
                rowWHTMOVE2.Item("STYLE_CODE") = grdWHTMOVE2.ActiveRow.Cells("STYLE_CODE").Value
                rowWHTMOVE2.Item("COLOR_CODE") = grdWHTMOVE2.ActiveRow.Cells("COLOR_CODE").Value
                rowWHTMOVE2.Item("INIT_OPER") = grdWHTMOVE2.ActiveRow.Cells("INIT_OPER").Value
                rowWHTMOVE2.Item("INIT_DATE") = grdWHTMOVE2.ActiveRow.Cells("INIT_DATE").Value
                rowWHTMOVE2.Item("LAST_OPER") = grdWHTMOVE2.ActiveRow.Cells("LAST_OPER").Value
                rowWHTMOVE2.Item("LAST_DATE") = grdWHTMOVE2.ActiveRow.Cells("LAST_DATE").Value
                rowWHTMOVE2.Item("STATUS") = grdWHTMOVE2.ActiveRow.Cells("STATUS").Value
                rowWHTMOVE2.Item("STYLE_DESC") = grdWHTMOVE2.ActiveRow.Cells("STYLE_DESC").Value
                rowWHTMOVE2.Item("ERROR_CODES") = grdWHTMOVE2.ActiveRow.Cells("ERROR_CODES").Value
                rowWHTMOVE2.Item("WHSE_TRAN_QTY_ORIG") = 0 'grdWHTMOVE2.ActiveRow.Cells("WHSE_TRAN_NO").Value

                frmDst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
        '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
        '    grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
        'End If

        If grd.Name = "grdARTCCPA1" Then

            Select Case e.Tool.Key
                Case "Location Inquiry"

            End Select
        End If
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdWHTPULL2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWHTMOVE2.AfterRowsDeleted
        OverMove()
    End Sub

    Private Sub grdWHTPULL2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTMOVE2.AfterRowUpdate
        OverMove()
    End Sub

    Private Sub grdWHTPULL2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTMOVE2.BeforeRowUpdate
        Dim LOCATION_CODE_TO As String = e.Row.Cells("LOCATION_CODE_TO").Value & String.Empty

        If movement_type = "BTS" Then
            e.Row.Cells("BAR_CODE_OTHER").Value = Check_BAR_CODE(e.Row.Cells("BAR_CODE_OTHER").Value)
        End If

            LOCATION_CODE_TO = LOCATION_CODE_TO.Trim

        e.Row.Cells("LOCATION_CODE_TO").Value = LOCATION_CODE_TO

        If LOCATION_CODE_TO.Length > 0 Then
            LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE_TO})
            If cdr Is Nothing Then
                MessageBox.Show("Invalid 'Location To' for the Warehouse.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub grdWHTPULL2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTMOVE2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "LOCATION_CODE_TO"
                grdClickCellButton(grdWHTMOVE2, "WHSE_CODE = '" & WHSE_CODE & "'", True, "", "LOCATION_CODE")
        End Select
    End Sub

    Private Sub grdWHTPULL2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTMOVE2.InitializeRow
        ' When the data is loaded error messages are placed in the ERROR_MSG column.
        ' These become the tooltip for the column with the error
        Dim ERROR_CODES As String = (e.Row.Cells("ERROR_CODES").Value & String.Empty).ToString.Trim

        If ERROR_CODES.Contains(InvalidLocation) Then
            e.Row.Cells("LOCATION_CODE_TO").Appearance.ForeColor = Color.Red
            e.Row.Cells("LOCATION_CODE_TO").ToolTipText = "Invalid Location"
        ElseIf e.Row.Cells("LOCATION_CODE_FROM").Value.ToString.Trim = e.Row.Cells("LOCATION_CODE_TO").Value.ToString.Trim Then
            e.Row.Cells("LOCATION_CODE_TO").Appearance.ForeColor = Color.Red
            e.Row.Cells("LOCATION_CODE_TO").ToolTipText = "From and To locations are the same"
        Else
            e.Row.Cells("LOCATION_CODE_TO").Appearance.ForeColor = Color.Black
            e.Row.Cells("LOCATION_CODE_TO").ToolTipText = String.Empty
        End If

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LOCATION_CODE"
                ValidateSelectedNewLocation()

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "LOCATION_CODE"
                ValidateSelectedNewLocation()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "LOCATION_CODE"
                ValidateSelectedNewLocation()
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "WHSE_CODE"
                Dim rowICTWHSE1 As DataRow = clsASCBASE1.LookUp("ICTWHSE1", txtWhse_Code_To.Text)

                If rowICTWHSE1 IsNot Nothing Then
                    For Each row As DataRow In frmDst.Tables("WHTMOVE2").Select()
                        row.Item("LOCATION_CODE_TO") = rowICTWHSE1.Item("WHSE_LOC_XIN") & ""
                    Next
                End If
        End Select
    End Sub


    Private Sub ValidateSelectedNewLocation()
        LookUp("WHTLOCM1", New String() {WHSE_CODE, MyBase.Absx1.txtFor("LOCATION_CODE").Text})

        If cdr IsNot Nothing Then
            If movement_type = "CTN" Or movement_type = "CFG" Or movement_type = "CCS" Then
            Else
                Dim sqlw As String = "ISNULL(LOCATION_CODE_TO, '*') = '*' OR LOCATION_CODE_TO = ''"
                If movement_type = "CHG" Then ' I SUSPECT THAT WE NEED TO DO THIS MORE OFTEN
                    sqlw = ""
                End If
                For Each row As DataRow In frmDst.Tables("WHTMOVE2").Select(sqlw)
                    row.Item("LOCATION_CODE_TO") = MyBase.Absx1.txtFor("LOCATION_CODE").Text
                Next

                If LOCATION_CODE_TO.Length > 0 AndAlso LOCATION_CODE_TO <> MyBase.Absx1.txtFor("LOCATION_CODE").Text Then
                    ' See if the user wants to change the location
                    Dim numdiff As Int32 = frmDst.Tables("WHTMOVE2").Select("ISNULL(LOCATION_CODE_TO, '*') = '" & LOCATION_CODE_TO & "'").Length
                    If numdiff > 0 Then
                        Dim msg As String = "There are " & numdiff & " Styles stamped with the previously selected Location (" & LOCATION_CODE_TO & ")."
                        msg &= vbCrLf & vbCrLf
                        msg &= " Do you want to update these with the new location?"
                        If MessageBox.Show(msg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                            For Each row As DataRow In frmDst.Tables("WHTMOVE2").Select("ISNULL(LOCATION_CODE_TO, '*') = '" & LOCATION_CODE_TO & "'")
                                row.Item("LOCATION_CODE_TO") = MyBase.Absx1.txtFor("LOCATION_CODE").Text
                            Next
                        End If
                    End If
                End If
            End If
            LOCATION_CODE_TO = MyBase.Absx1.txtFor("LOCATION_CODE").Text
        End If
    End Sub

#End Region

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
                txtBAR_CODE2.Text = txtBAR_CODE.Text.Substring(0, 1) & Format(Val(txtBAR_CODE.Text.Substring(1)) + C - 1, "0000000")
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

        Dim prefix As String = ""
        If BAR_CODE = "" Then Return BAR_CODE

        If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
            prefix = BAR_CODE.ToUpper.Substring(0, 1)
            BAR_CODE = BAR_CODE.Substring(1)
        End If

        If BAR_CODE.PadLeft(8, "0") <> Format(Val(BAR_CODE), "".PadLeft(8, "0")) Then
            BAR_CODE = ""
        Else
            If prefix = "" Then
                BAR_CODE = BAR_CODE.PadLeft(8, "0")
            Else
                BAR_CODE = prefix & BAR_CODE.PadLeft(7, "0")
            End If
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
                    & " on " & cdr.Item("PO_DATE_RECEIVED"), _
                    MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                BAR_CODE = ""
            Else
                Dim rowWHTBARC1 As DataRow = frmDst.Tables("WHTBARC1").Rows.Find(BAR_CODE)
                If rowWHTBARC1 IsNot Nothing Then
                    MsgBox("LPN already exists in Current Transaction", _
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
                MsgBox("Invalid Value for LPN", _
                    MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN2")
                BAR_CODE2 = ""
            Else
                If BAR_CODE2 < BAR_CODE Then
                    MsgBox("Invalid Range for LPNs" & vbCrLf _
                        & BAR_CODE & " thru " & BAR_CODE2, _
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
                    & " in Range from " & BAR_CODE & " thru " & BAR_CODE2, _
                    MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                BAR_CODE2 = ""
            Else
                If movement_type = "BTS" Then
                    Dim rowWHTBARC1s() As DataRow = frmDst.Tables("WHTMOVE2").Select("BAR_CODE_OTHER >= '" & BAR_CODE & "' and BAR_CODE_OTHER <= '" & BAR_CODE2 & "'")
                    If rowWHTBARC1s.Length <> 0 Then
                        MsgBox(CStr(rowWHTBARC1s.Length) & " LPN(s) already exists in Current Transaction", _
                            MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                        BAR_CODE2 = ""
                    End If
                Else
                    Dim rowWHTBARC1s() As DataRow = frmDst.Tables("WHTBARC1").Select("BAR_CODE >= '" & BAR_CODE & "' and BAR_CODE <= '" & BAR_CODE2 & "'")
                    If rowWHTBARC1s.Length <> 0 Then
                        MsgBox(CStr(rowWHTBARC1s.Length) & " LPN(s) already exists in Current Transaction", _
                            MsgBoxStyle.OkOnly, "Invalid Value Specified for LPN")
                        BAR_CODE2 = ""
                    End If
                End If

            End If

            txtBAR_CODE2.Text = BAR_CODE2
            If BAR_CODE2 = "" Then
                txtBAR_CODE2.Focus()
            Else
                If movement_type <> "BTS" Then
                    Write_LPNs()
                    txtBAR_CODE.Focus()
                End If
            End If
        End If
    End Sub

    Sub Write_LPNs()

        If txtBAR_CODE.Text = "" Or txtBAR_CODE2.Text = "" Then Exit Sub
        Dim BAR_CODE As String = txtBAR_CODE.Text
        Dim BAR_CODE2 As String = txtBAR_CODE2.Text
        Dim QTY As Int64
        Dim BAR_CODEX As String

        If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
            QTY = Val(BAR_CODE2.Substring(1)) - Val(BAR_CODE.Substring(1)) + 1
        Else
            QTY = Val(BAR_CODE2) - Val(BAR_CODE) + 1
        End If

        Dim BAR_CODE_first As Int64

        If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
            BAR_CODE_first = Val(BAR_CODE.Substring(1))
        Else
            BAR_CODE_first = Val(BAR_CODE)
        End If

        For i As Integer = 1 To QTY
            Dim rowWHTBARC1 As DataRow = frmDst.Tables("WHTBARC1").NewRow
            If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
                BAR_CODEX = BAR_CODE.ToUpper.Substring(0, 1) & Format(BAR_CODE_first + i - 1, "".PadLeft(7, "0"))
            Else
                BAR_CODEX = Format(BAR_CODE_first + i - 1, "".PadLeft(8, "0"))
            End If
            If frmDst.Tables("WHTBARC1").Rows.Find(BAR_CODEX) IsNot Nothing Then
                MsgBox("Case ID " & BAR_CODEX & " is already in the list", MsgBoxStyle.OkOnly, "Cannot Add Case ID")
                Exit Sub
            End If

            rowWHTBARC1.Item("BAR_CODE") = BAR_CODEX

            rowWHTBARC1.Item("PO_ORDER_NO") = "?" ' PO_ORDER_NO
            'rowWHTBARC1.Item("PO_SHIPMENT_NO") = ""
            'rowWHTBARC1.Item("PO_SHIPMENT_LNO") = PO_SHIPMENT_LNO
            'rowWHTBARC1.Item("CARTON_NO") = CARTON_NO
            rowWHTBARC1.Item("LOAD_NO") = LOAD_NO
            '   rowWHTBARC1.Item("LOCATION_CODE") = rowICTWHSE1.Item("WHSE_LOC_CTN")
            frmDst.Tables("WHTBARC1").Rows.Add(rowWHTBARC1)
        Next

        txtBAR_CODE.Text = ""
        txtBAR_CODE2.Text = ""

        txtBAR_CODE.Focus()

        Display_Totals()
    End Sub

    Private Sub grdWHTMOVE3_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTMOVE3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim STYLE_CODE As String = Validate_Style(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("STYLE_CODE").Value)
                If STYLE_CODE <> "" Then
                    e.Cell.Row.Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                End If
        End Select
    End Sub

    Private Sub grdWHTMOVE3_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWHTMOVE3.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdWHTMOVE3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTMOVE3.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdWHTMOVE3_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWHTMOVE3.BeforeExitEditMode
        If grdWHTMOVE3.ActiveCell IsNot Nothing Then
            With grdWHTMOVE3.ActiveCell
                Select Case .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        'If .EditorResolved.Value & "" <> "" AndAlso .EditorResolved.Value <> CStr(.EditorResolved.Value & "").ToUpper Then
                        If .EditorResolved.IsValid AndAlso .EditorResolved.Value & "" <> "" Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value, .Column.Key)
                        End If
                        'End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdWHTMOVE3_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdWHTMOVE3.BeforeRowsDeleted
        e.DisplayPromptMsg = False
    End Sub

    Private Sub grdWHTMOVE3_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTMOVE3.BeforeRowUpdate
        Validate_Columns("STYLE_CODE", e.Cancel)
        If e.Cancel Then MsgBox("Invalid Style")
        If Not e.Cancel Then
            Validate_Columns("COLOR_CODE", e.Cancel)
            If e.Cancel Then MsgBox("Invalid Color")
        Else

        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("WHSE_TRAN_NO").Value = WHSE_TRAN_NO
        End If

    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)

        With grdWHTMOVE3.ActiveRow

            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    Cancel = (STYLE_CODE = "")

                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If

            End Select
        End With
    End Sub

    Function Select_Style(ByRef COLOR_CODE As String) As String

        Dim STYLE_CODE As String = ""

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            STYLE_CODE = ASCMAIN1.CodeSelector.SelectedCode
            STYLE_CODE = Validate_Style(STYLE_CODE)
        End If

        If COLOR_CODE <> "" Then
            If STYLE_CODE <> "" Then
                Dim rowICTSTYC1 As DataRow = frmASFBASE1.LookUp("ICTSTYC1", New String() {STYLE_CODE, COLOR_CODE})
                If rowICTSTYC1 Is Nothing Then
                    COLOR_CODE = ""
                    'MsgBox("Color Code '" & COLOR_CODE & "' is not Associated with Style " & STYLE_CODE)
                    'STYLE_CODE = ""
                End If
            End If
        End If

        If COLOR_CODE = "" Then
            If COLOR_CODEs.Count = 1 Then
                COLOR_CODE = COLOR_CODEs(0)
            Else
                Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"

                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("COLOR_CODE", , sql_where)
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""

                    'ASCMAIN1.CodeSelector.SQL = "Select * from (" & ASCMAIN1.CodeSelector.SQL & ")" _
                    '    & " where COLOR_CODE in " _
                    '    & " (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "')"

                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    COLOR_CODE = ASCMAIN1.CodeSelector.SelectedCode
                    If COLOR_CODE = "" Then STYLE_CODE = ""
                End If
            End If
        End If
        Return STYLE_CODE
    End Function

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim E As String = ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = frmASFBASE1.LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            E = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                ' E = "Style Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                E = "Style does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                E = "Style does not have a valid Division Code" & vbCrLf
            End If
        End If

        If E = "" Then
            COLOR_CODEs.Clear()

            frmASFBASE1.Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In frmDst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If E <> "" Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If E = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

    Private Sub grdWHTMOVE3_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTMOVE3.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdWHTMOVE3, sql_where)
                    Dim COLOR_CODE As String = .Cells("COLOR_CODE").Value & ""
                    Dim STYLE_CODE As String = Select_Style(COLOR_CODE)

                    If STYLE_CODE <> "" Then

                        If Validate_Style(STYLE_CODE) <> "" Then
                            .Cells("STYLE_CODE").Value = STYLE_CODE
                            .Cells("COLOR_CODE").Value = COLOR_CODE
                            .Update()
                        End If

                    End If

                Case "COLOR_CODE"

                    Dim STYLE_CODE As String = grdWHTMOVE3.ActiveRow.Cells("STYLE_CODE").Value
                    If Validate_Style(STYLE_CODE) <> "" Then
                        ' TO REFRESH COLOR_CODES
                    End If
                    Dim sql_where As String = "COLOR_CODE in ('" & Join(COLOR_CODEs.ToArray, "','") & "')"
                    grdClickCellButton(grdWHTMOVE3, sql_where, , , "COLOR_CODE")

            End Select
        End With
    End Sub

    Private Sub btnAddCaseIDs_Click(sender As System.Object, e As System.EventArgs) Handles btnAddCaseIDs.Click
        If movement_type = "BTS" Then

            If Add_CaseIDs_Clicked = True Then
                MsgBox("You have already clicked Add Case ID's, click Cancel to reset and start over", MsgBoxStyle.OkOnly, "Cannot Proceed")
                Exit Sub
            End If


            If txtBAR_CODE.Text = "" Or txtBAR_CODE2.Text = "" Then Exit Sub
            Dim BAR_CODE As String = txtBAR_CODE.Text
            Dim BAR_CODE2 As String = txtBAR_CODE2.Text
            Dim QTY As Int64
            Dim BAR_CODE_first As Int64
            Dim BAR_CODEX As String
            Add_CaseIDs_Clicked = True

            If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
                QTY = Val(BAR_CODE2.Substring(1)) - Val(BAR_CODE.Substring(1)) ' + 1 not used due to Zero base loop
                BAR_CODE_first = Val(BAR_CODE.Substring(1))
            Else
                QTY = Val(BAR_CODE2) - Val(BAR_CODE) ' + 1 not used due to Zero base loop
                BAR_CODE_first = Val(BAR_CODE)
            End If

            Dim row As DataRow = frmDst.Tables("WHTMOVE2").Rows(0)
            For i = 0 To QTY
                If BAR_CODE.ToUpper.Substring(0, 1) >= "A" Then
                    BAR_CODEX = BAR_CODE.ToUpper.Substring(0, 1) & Format(BAR_CODE_first + i, "".PadLeft(7, "0"))
                Else
                    BAR_CODEX = Format(BAR_CODE_first + i, "".PadLeft(8, "0"))
                End If
                If i = 0 Then
                    row.Item("BAR_CODE_OTHER") = BAR_CODEX
                    row.Item("WHSE_TRAN_QTY") = numUnits.Value
                Else
                    Dim rowWHTMOVE2 As DataRow = frmDst.Tables("WHTMOVE2").NewRow
                    rowWHTMOVE2.Item("WHSE_TRAN_NO") = row.Item("WHSE_TRAN_NO")
                    WHSE_TRAN_LNO += 1
                    rowWHTMOVE2.Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                    rowWHTMOVE2.Item("LOCATION_CODE_FROM") = row.Item("LOCATION_CODE_FROM")
                    rowWHTMOVE2.Item("LOCATION_CODE_TO") = row.Item("LOCATION_CODE_TO")
                    rowWHTMOVE2.Item("BAR_CODE") = row.Item("BAR_CODE")
                    rowWHTMOVE2.Item("WHSE_TRAN_QTY") = numUnits.Value
                    rowWHTMOVE2.Item("STYLE_CODE") = row.Item("STYLE_CODE")
                    rowWHTMOVE2.Item("COLOR_CODE") = row.Item("COLOR_CODE")
                    rowWHTMOVE2.Item("INIT_OPER") = row.Item("INIT_OPER")
                    rowWHTMOVE2.Item("INIT_DATE") = row.Item("INIT_DATE")
                    rowWHTMOVE2.Item("LAST_OPER") = row.Item("LAST_OPER")
                    rowWHTMOVE2.Item("LAST_DATE") = row.Item("LAST_DATE")
                    rowWHTMOVE2.Item("STATUS") = row.Item("STATUS")
                    rowWHTMOVE2.Item("STYLE_DESC") = row.Item("STYLE_DESC")
                    rowWHTMOVE2.Item("ERROR_CODES") = row.Item("ERROR_CODES")
                    rowWHTMOVE2.Item("WHSE_TRAN_QTY_ORIG") = 0 'grdWHTMOVE2.ActiveRow.Cells("WHSE_TRAN_NO").Value
                    rowWHTMOVE2.Item("BAR_CODE_OTHER") = BAR_CODEX
                    frmDst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
                End If
            Next

            txtBAR_CODE.Text = ""
            txtBAR_CODE2.Text = ""
            numUnits.Value = 0
            txtBAR_CODE.Focus()
        Else
            Write_LPNs()
        End If

    End Sub

    Sub Create_Adjustment_CTN()

        For Each rowWHTMOVE2_from As DataRow In frmDst.Tables("WHTMOVE2").Select("")
            Dim STYLE_CODE As String = rowWHTMOVE2_from.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTMOVE2_from.Item("COLOR_CODE")
            Dim ADJ_QTY As Int64 = Val(rowWHTMOVE2_from.Item("WHSE_TRAN_QTY") & "")
            Dim row As DataRow = frmDst.Tables("ICTIADJS").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If row Is Nothing Then
                row = frmDst.Tables("ICTIADJS").Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, 0})
            End If
            row.Item("ADJ_QTY") = Val(row.Item("ADJ_QTY") & "") - ADJ_QTY
        Next

        For Each rowWHTBARC1 As DataRow In frmDst.Tables("WHTBARC1").Select("")
            For Each rowWHTMOVE3 As DataRow In frmDst.Tables("WHTMOVE3").Select("")

                Dim STYLE_CODE As String = rowWHTMOVE3.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowWHTMOVE3.Item("COLOR_CODE")
                Dim ADJ_QTY As Int64 = Val(rowWHTMOVE3.Item("CASE_QTY") & "")

                Dim rowWHTMOVE2 As DataRow = frmDst.Tables("WHTMOVE2").NewRow
                With rowWHTMOVE2
                    .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                    WHSE_TRAN_LNO += 1
                    .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                    .Item("LOCATION_CODE_FROM") = rowICTWHSE1.Item("WHSE_LOC_CTN")
                    .Item("LOCATION_CODE_TO") = LOCATION_CODE_TO
                    .Item("BAR_CODE") = rowWHTBARC1.Item("BAR_CODE")
                    .Item("WHSE_TRAN_QTY") = rowWHTMOVE3.Item("CASE_QTY")
                    .Item("WHSE_TRAN_QTY_ORIG") = rowWHTMOVE3.Item("CASE_QTY")
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("STATUS") = "U"
                    .Item("LOAD_NO_FROM") = rowICTWHSE1.Item("WHSE_DEF_LOAD_NO")
                    .Item("LOAD_NO_TO") = LOAD_NO
                End With
                frmDst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

                Dim row As DataRow = frmDst.Tables("ICTIADJS").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                If row Is Nothing Then
                    row = frmDst.Tables("ICTIADJS").Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, 0})
                End If
                row.Item("ADJ_QTY") = Val(row.Item("ADJ_QTY") & "") + ADJ_QTY
            Next
        Next

        If frmDst.Tables("ICTIADJS").Select("ADJ_QTY <> 0").Length <> 0 Then

            Dim rowICTIADJ1 As DataRow = Get_ADJ_NO("Re-Cartonization")

            Dim ADJ_LNO As Int64 = 0
            Dim TOTAL_COSTS As Decimal = 0

            For Each rowICTIADJS As DataRow In frmDst.Tables("ICTIADJS").Select("ADJ_QTY <> 0")
                Dim STYLE_CODE As String = rowICTIADJS.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowICTIADJS.Item("COLOR_CODE")
                Dim ADJ_QTY As Int64 = Val(rowICTIADJS.Item("ADJ_QTY") & "")

                Record_ICTIADJ2(ADJ_LNO, _
                                rowICTWHSE1.Item("WHSE_LOC_CTN"), _
                                rowICTWHSE1.Item("WHSE_DEF_BAR_CODE"), _
                                movement_type, _
                                STYLE_CODE, COLOR_CODE, ADJ_QTY)
            Next
        End If
    End Sub

    Sub Create_Adjustment_CFG()

        Dim rowICTIADJ1 As DataRow = Get_ADJ_NO("Re-Configuration")
        Dim ADJ_LNO As Int64 = 0
        Dim TOTAL_COSTS As Decimal = 0

        For Each rowWHTMOVE2_from As DataRow In frmDst.Tables("WHTMOVE2").Select("")
            Dim STYLE_CODE As String = rowWHTMOVE2_from.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTMOVE2_from.Item("COLOR_CODE")
            Dim ADJ_QTY As Int64 = Val(rowWHTMOVE2_from.Item("WHSE_TRAN_QTY") & "")

            Dim LOCATION_CODE As String = rowWHTMOVE2_from.Item("LOCATION_CODE_FROM")
            Dim BAR_CODE As String = rowWHTMOVE2_from.Item("BAR_CODE")
            Dim LOAD_NO As String = rowWHTMOVE2_from.Item("LOAD_NO_FROM")

            Record_ICTIADJ2(ADJ_LNO, _
                            LOCATION_CODE, _
                            BAR_CODE, _
                            movement_type, _
                            STYLE_CODE, COLOR_CODE, -1 * ADJ_QTY)

            For Each rowWHTMOVE3 As DataRow In frmDst.Tables("WHTMOVE3").Select("")

                STYLE_CODE = rowWHTMOVE3.Item("STYLE_CODE")
                COLOR_CODE = rowWHTMOVE3.Item("COLOR_CODE")
                ADJ_QTY = Val(rowWHTMOVE3.Item("CASE_QTY") & "")

                'Dim rowWHTMOVE2 As DataRow = frmDst.Tables("WHTMOVE2").NewRow
                'With rowWHTMOVE2
                '    .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                '    WHSE_TRAN_LNO += 1
                '    .Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                '    .Item("LOCATION_CODE_FROM") = LOCATION_CODE
                '    .Item("LOCATION_CODE_TO") = LOCATION_CODE ' LOCATION_CODE_TO
                '    .Item("BAR_CODE") = BAR_CODE
                '    .Item("WHSE_TRAN_QTY") = ADJ_QTY ' rowWHTMOVE3.Item("CASE_QTY")
                '    .Item("WHSE_TRAN_QTY_ORIG") = ADJ_QTY ' rowWHTMOVE3.Item("CASE_QTY")
                '    .Item("STYLE_CODE") = STYLE_CODE
                '    .Item("COLOR_CODE") = COLOR_CODE
                '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                '    .Item("INIT_DATE") = DATETIME_STAMP
                '    .Item("STATUS") = "U"
                '    .Item("LOAD_NO_FROM") = LOAD_NO
                '    .Item("LOAD_NO_TO") = LOAD_NO
                'End With
                'frmDst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)


                Record_ICTIADJ2(ADJ_LNO, _
                                LOCATION_CODE, _
                                BAR_CODE, _
                                movement_type, _
                                STYLE_CODE, COLOR_CODE, ADJ_QTY)

            Next
        Next

    End Sub

    Sub Create_Adjustment_CCS()

        For Each rowWHTBARC1 As DataRow In frmDst.Tables("WHTBARC1").Select("")
            Dim BAR_CODE As String = rowWHTBARC1.Item("BAR_CODE")

            For Each rowWHTMOVE3 As DataRow In frmDst.Tables("WHTMOVE3").Select("")

                Dim STYLE_CODE As String = rowWHTMOVE3.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowWHTMOVE3.Item("COLOR_CODE")
                Dim ADJ_QTY As Int64 = Val(rowWHTMOVE3.Item("CASE_QTY") & "")

                AddItemToMove(WHSE_CODE, _
                              rowICTWHSE1.Item("WHSE_LOC_ADJ"), _
                              STYLE_CODE, _
                              COLOR_CODE, _
                               rowICTWHSE1.Item("WHSE_DEF_BAR_CODE"), _
                              rowICTWHSE1.Item("WHSE_DEF_LOAD_NO"), _
                              ADJ_QTY, _
                              txtLOCATION_CODE.Text, _
                              BAR_CODE, _
                              LOAD_NO)

                rowWHTMOVE3.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO

                Dim row As DataRow = frmDst.Tables("ICTIADJS").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                If row Is Nothing Then
                    row = frmDst.Tables("ICTIADJS").Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, 0})
                End If
                row.Item("ADJ_QTY") = Val(row.Item("ADJ_QTY") & "") + ADJ_QTY
            Next
        Next

        If frmDst.Tables("ICTIADJS").Select("ADJ_QTY <> 0").Length <> 0 Then

            Dim rowICTIADJ1 As DataRow = Get_ADJ_NO("Cases Created")

            Dim ADJ_LNO As Int64 = 0
            Dim TOTAL_COSTS As Decimal = 0

            For Each rowICTIADJS As DataRow In frmDst.Tables("ICTIADJS").Select("ADJ_QTY <> 0")
                Dim STYLE_CODE As String = rowICTIADJS.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowICTIADJS.Item("COLOR_CODE")
                Dim ADJ_QTY As Int64 = Val(rowICTIADJS.Item("ADJ_QTY") & "")

                Record_ICTIADJ2(ADJ_LNO, _
                                rowICTWHSE1.Item("WHSE_LOC_ADJ"), _
                                rowICTWHSE1.Item("WHSE_DEF_BAR_CODE"), _
                                "CCS", _
                                STYLE_CODE, COLOR_CODE, ADJ_QTY)
            Next
        End If
    End Sub

    Sub Create_Adjustment()

        Dim rowICTIADJ1 As DataRow = Get_ADJ_NO()

        Dim ADJ_LNO As Int64 = 0
        Dim TOTAL_COSTS As Decimal = 0

        Dim tableName = If(movement_type <> "BTS", "WHTMOVE2", "WHTMOVE3")
        Dim qtyColumn = If(movement_type <> "BTS", "WHSE_TRAN_QTY", "CASE_QTY")


        For Each rowWHTMOVE2 As DataRow In frmDst.Tables(tableName).Select("")
            Dim STYLE_CODE As String = rowWHTMOVE2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTMOVE2.Item("COLOR_CODE")
            Dim ADJ_QTY As Int64 = Val(rowWHTMOVE2.Item(qtyColumn) & "")

            Record_ICTIADJ2(ADJ_LNO, _
                            rowWHTMOVE2.Item("LOCATION_CODE_FROM"), _
                            rowWHTMOVE2.Item("BAR_CODE"), _
                            rowWHTMOVE2.Item("LOAD_NO_FROM"), _
                            STYLE_CODE, COLOR_CODE, ADJ_QTY)

        Next

    End Sub

    Function Get_ADJ_NO(Optional ADJ_NOTE = "Location Adjustment") As DataRow

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
        Else
            ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        End If

        Dim rowICTIADJ1 As DataRow = frmDst.Tables("ICTIADJ1").NewRow
        With rowICTIADJ1
            .Item("ADJ_NO") = ADJ_NO
            .Item("ADJ_DATE") = DATETIME_STAMP.Date
            .Item("WHSE_CODE") = WHSE_CODE
            .Item("REASON_CODE") = Absx1.txtFor("REASON_CODE").Text
            .Item("ADJ_NOTE") = ADJ_NOTE
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("REGISTER_IND") = "0"
            .Item("ADJ_SOURCE") = "L"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("TOTAL_COSTS") = 0
            .Item("ADJ_REF") = WHSE_TRAN_NO
        End With
        frmDst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

        Return rowICTIADJ1
    End Function

    Sub Record_ICTIADJ2(ByRef ADJ_LNO As Integer, _
                             LOCATION_CODE As String, _
                             BAR_CODE As String, _
                             ADJ_REF As String, _
                             STYLE_CODE As String, _
                             COLOR_CODE As String, _
                             ADJ_QTY As Int64)

        Dim rowICTIADJ2 As DataRow = frmDst.Tables("ICTIADJ2").NewRow
        With rowICTIADJ2
            .Item("ADJ_NO") = ADJ_NO
            ADJ_LNO += 1
            .Item("ADJ_LNO") = ADJ_LNO
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("ADJ_QTY") = ADJ_QTY
            Dim rowICTSTYL1 As DataRow = frmASFBASE1.LookUp("ICTSTYL1", STYLE_CODE)
            .Item("STYLE_COST") = rowICTSTYL1.Item("STYLE_COST")
            .Item("STYLE_CLASS_CODE") = rowICTSTYL1.Item("STYLE_CLASS_CODE")
            .Item("SALES_DIVISION_CODE") = rowICTSTYL1.Item("SALES_DIVISION_CODE")
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("LOCATION_CODE") = LOCATION_CODE
            .Item("BAR_CODE") = BAR_CODE
            .Item("ADJ_REF") = ADJ_REF
        End With
        frmDst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
    End Sub

    Sub CreateAdjustmentCHG()
        Dim rowICTIADJ1 As DataRow = Get_ADJ_NO()

        Dim STYLE_CODE_new As String = Absx1.txtFor("STYLE_CODE").Text

        Dim ADJ_LNO As Int64 = 0
        Dim TOTAL_COSTS As Decimal = 0

        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
            Dim STYLE_CODE As String = rowWHTMOVE2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTMOVE2.Item("COLOR_CODE")
            Dim ADJ_QTY As Int64 = -1 * Val(rowWHTMOVE2.Item("WHSE_TRAN_QTY") & "")
            Dim LOCATION_CODE As String = rowWHTMOVE2.Item("LOCATION_CODE_FROM")
            Dim ADJ_REF As String = "Chg to " & STYLE_CODE_new
            For i As Integer = 1 To 2
                Record_ICTIADJ2(ADJ_LNO, _
                                rowWHTMOVE2.Item("LOCATION_CODE_TO"), _
                                rowWHTMOVE2.Item("BAR_CODE"), _
                                ADJ_REF, STYLE_CODE, COLOR_CODE, ADJ_QTY)
                If i = 1 Then
                    ADJ_REF = "Chg from " & STYLE_CODE
                    ADJ_QTY = -1 * ADJ_QTY
                    STYLE_CODE = STYLE_CODE_new
                End If
            Next
        Next

    End Sub

    Sub CreateAdjustmentCMB()
        'Dim rowICTIADJ1 As DataRow = Get_ADJ_NO()

        'Dim STYLE_CODE_new As String = Absx1.txtFor("STYLE_CODE").Text

        'Dim ADJ_LNO As Int64 = 0
        'Dim TOTAL_COSTS As Decimal = 0

        'For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
        '    Dim STYLE_CODE As String = rowWHTMOVE2.Item("STYLE_CODE")
        '    Dim COLOR_CODE As String = rowWHTMOVE2.Item("COLOR_CODE")
        '    Dim ADJ_QTY As Int64 = Val(rowWHTMOVE2.Item("WHSE_TRAN_QTY") & "")
        '    Dim LOCATION_CODE As String = rowWHTMOVE2.Item("LOCATION_CODE_FROM")
        '    Dim ADJ_REF As String = "Cmb to " & BAR_CODE_CMB
        '    Record_ICTIADJ2(ADJ_LNO, _
        '    rowWHTMOVE2.Item("LOCATION_CODE_TO"), _
        '    rowWHTMOVE2.Item("BAR_CODE"), _
        '    ADJ_REF, STYLE_CODE, COLOR_CODE, ADJ_QTY)
        'Next

    End Sub

    Private Sub CreateNewWHTBARC1_BTS()

        ' Dim rowICTIADJ1 As DataRow = Get_ADJ_NO()

        'Dim ADJ_LNO As Int64 = 0
        'Dim TOTAL_COSTS As Decimal = 0

        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("") ' S/B ONLY 1 ROW
            Dim STYLE_CODE As String = rowWHTMOVE2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowWHTMOVE2.Item("COLOR_CODE")
            'Dim ADJ_QTY As Int64 = -1 * Val(rowWHTMOVE2.Item("WHSE_TRAN_QTY") & "")
            Dim BAR_CODE As String = rowWHTMOVE2.Item("BAR_CODE")
            'Dim LOCATION_CODE As String = rowWHTMOVE2.Item("LOCATION_CODE_FROM")
            'Record_ICTIADJ2(ADJ_LNO, LOCATION_CODE, BAR_CODE, LOAD_NO, STYLE_CODE, COLOR_CODE, ADJ_QTY)

            Dim row As DataRow = frmASFBASE1.LookUp("WHTBARC1", BAR_CODE)
            Dim rowWHTBARC1 As DataRow = frmDst.Tables("WHTBARC1").NewRow
            rowWHTBARC1.ItemArray = row.ItemArray
            BAR_CODE = rowWHTMOVE2.Item("BAR_CODE_OTHER")
            rowWHTBARC1.Item("BAR_CODE") = BAR_CODE
            rowWHTBARC1.Item("LOAD_NO") = LOAD_NO
            rowWHTBARC1.Item("TRAN_NO") = WHSE_TRAN_NO
            frmDst.Tables("WHTBARC1").Rows.Add(rowWHTBARC1)
            'LOCATION_CODE = txtLOCATION_CODE.Text
            'ADJ_QTY = -1 * ADJ_QTY
        Next
    End Sub

    Private Sub grdWHTBARC1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWHTBARC1.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdWHTBARC1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTBARC1.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdWHTBARC1_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdWHTBARC1.BeforeRowsDeleted
        e.DisplayPromptMsg = False
    End Sub

    Sub Display_Totals()
        Dim CASE_QTY As Int64 = Val(frmDst.Tables("WHTMOVE3").Compute("SUM(CASE_QTY)", "") & "")
        Dim CASES As Int64 = frmDst.Tables("WHTBARC1").Rows.Count
        lblUnits.Text = CASE_QTY * CASES
    End Sub

    Private Sub txtLOCATION_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtLOCATION_CODE.ValueChanged

    End Sub

    Private Sub cmdAddStyle_Click(sender As Object, e As EventArgs) Handles cmdAddStyle.Click
        Dim grow As UltraGridRow = grdWHTMOVE2.Rows(0)
        Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value

        Dim rowWHTMOVE3 As DataRow

        If frmDst.Tables("WHTMOVE3").Rows.Find(New String() {WHSE_TRAN_NO, STYLE_CODE, COLOR_CODE}) IsNot Nothing Then
            MsgBox("Record already exists for Style-Color " & STYLE_CODE & "-" & COLOR_CODE, MsgBoxStyle.OkOnly, "Cannot Add Style")
            Exit Sub
        End If

        rowWHTMOVE3 = frmDst.Tables("WHTMOVE3").NewRow
        With rowWHTMOVE3
            .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            .Item("STYLE_CODE") = STYLE_CODE
            rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
            .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
            .Item("COLOR_CODE") = COLOR_CODE
        End With
        frmDst.Tables("WHTMOVE3").Rows.Add(rowWHTMOVE3)

        grdWHTMOVE3.Rows.Refresh(RefreshRow.FireInitializeRow)
        Sort_grdColumns(grdWHTMOVE2, "STYLE_CODE,COLOR_CODE")

    End Sub

    Private Sub cmdSizeScale_Click(sender As Object, e As EventArgs) Handles cmdSizeScale.Click

        frmDst.Tables("WHTMOVE3").Rows.Clear()


        Dim grow As UltraGridRow = grdWHTMOVE2.Rows(0)
        Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
        Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value
        Dim WHSE_TRAN_QTY As Integer = Val(grow.Cells("WHSE_TRAN_QTY").Value & "")
        Dim rowICTSTYC1 As DataRow

        Dim rowICTSTYLS As DataRow = LookUp("ICTSTYLS", STYLE_CODE)
        If rowICTSTYLS Is Nothing Then
            MsgBox("No Size Scale on File", MsgBoxStyle.OkOnly, "Cannot find Size Scale for Style " & STYLE_CODE)
            Exit Sub
        End If

        Dim T As Integer = 0

        For I As Integer = 1 To 12
            Dim SIZE As String = rowICTSTYLS.Item("SIZE_" & Format(I, "00")) & ""
            Dim QTY As Integer = Val(rowICTSTYLS.Item("QTY_" & Format(I, "00")) & "")

            If SIZE = "" Then Exit For
            T += QTY

            Dim STYLE_CODE_SIZE As String = STYLE_CODE & SIZE
            rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_SIZE)
            If rowICTSTYL1 Is Nothing Then
                MsgBox("Sized-Style " & STYLE_CODE_SIZE & " is Not On File", MsgBoxStyle.OkOnly, "Cannot Complete Size Scale Re-Configure")
                Exit Sub
            End If
            rowICTSTYC1 = LookUp("ICTSTYC1", New String() {STYLE_CODE_SIZE, COLOR_CODE})
            If rowICTSTYC1 Is Nothing Then
                MsgBox("Color-Code " & COLOR_CODE & " for Sized-Style " & STYLE_CODE_SIZE & " is Not On File", MsgBoxStyle.OkOnly, "Cannot Complete Size Scale Re-Configure")
                Exit Sub
            End If

            Dim rowWHTMOVE3 As DataRow = frmDst.Tables("WHTMOVE3").NewRow
            With rowWHTMOVE3
                .Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                .Item("STYLE_CODE") = STYLE_CODE_SIZE
                .Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("CASE_QTY") = QTY
            End With
            frmDst.Tables("WHTMOVE3").Rows.Add(rowWHTMOVE3)
        Next

        If T = 0 Or WHSE_TRAN_QTY = 0 Then
            MsgBox("Total Qty in Carton is not congruous with Size Scale Total Qty", MsgBoxStyle.OkOnly, "Warning")
            Exit Sub
        End If

        If WHSE_TRAN_QTY Mod T <> 0 Then
            MsgBox("Total Qty in Carton is not evenly divisible by Total of Sizes", MsgBoxStyle.OkOnly, "Warning")
        Else
            Dim F As Integer = WHSE_TRAN_QTY / T
            For Each row As DataRow In frmDst.Tables("WHTMOVE3").Select("")
                row.Item("CASE_QTY") = Val(row.Item("CASE_QTY") & "") * F
            Next
        End If


    End Sub

    Private Sub grdWHTMOVE3_Error(sender As Object, e As ErrorEventArgs) Handles grdWHTMOVE3.Error
        ' issue where grid error messages appear behind the modal form
        MsgBox(e.ErrorText, MsgBoxStyle.OkOnly, "Error")
        e.Cancel = True
    End Sub

End Class