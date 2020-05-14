Imports Infragistics.Win.UltraWinGrid

Public Class SOFORDR2
    Public mode As String = "" ' N = New, "" = Update Next Step
    Public SELECT_ST As String
    Public CUST_CODE As String
    Public CUST_ADDR_TYPE As String
    Public SHIP_VIA_CODE As String
    Public STARTDATE As DateTime
    Public CANCELDATE As DateTime
    Public WHSE_CODE As String
    Public ORDR_CUST_PO As String
    Public ORDR_CATEGORY As String
    Public TERM_CODE As String
    Public CUST_TERM_CODE As String
    Public FRT_TERMS As String
    Public ORDR_SHIP_INSTR As String
    Public ORDR_MESSAGE As String
    Public rowARTCUST1 As DataRow
    Public rowARTCUST2 As DataRow
    Public CUST_PRICE_TIER_PVC As String = "PC"
    Public ORDR_BUYER_NAME As String = ""
    Public ORDR_BUYER_EMAIL As String = ""
    Public ORDR_BUYER_CONTACT_NO As Integer
    Public HAS_CONTACT_CHANGES As Boolean = False
    Private BuyerInfoLive As Boolean = True
    Private FF As ASFBASE1

    Public Sub New(ByVal frmASFBASE1 As ASFBASE1,
                   ByRef inARTCUST1 As DataRow,
                   ByVal in_CUST_PO As String,
                   ByVal in_ORDR_CATEGORY As String,
                   Optional ByVal in_BuyerInfoLive As Boolean = True)
        FF = frmASFBASE1
        InitializeComponent()
        BuyerInfoLive = in_BuyerInfoLive
        rowARTCUST1 = inARTCUST1
        ORDR_CUST_PO = in_CUST_PO
        ORDR_CATEGORY = in_ORDR_CATEGORY
        txtORDR_CUST_PO.Text = ORDR_CUST_PO
        txtORDR_CATEGORY.Text = ORDR_CATEGORY
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        grdARTCUST2.DataSource = FF.dst.Tables("ARTCUST2")
        If FF.dst.Tables("ARTCUST2").Rows.Count = 1 Then
            grdARTCUST2.Rows.Item(0).Selected = True
            SELECT_ST = grdARTCUST2.Rows.Item(0).Cells("CUST_ADDR_CODE").Text
            CUST_CODE = grdARTCUST2.Rows.Item(0).Cells("CUST_CODE").Text
            CUST_ADDR_TYPE = grdARTCUST2.Rows.Item(0).Cells("CUST_ADDR_TYPE").Text
        End If
        txtSHIP_VIA_CODE.Text = rowARTCUST1.Item("SHIP_VIA_CODE").ToString
        txtWHSE_CODE.Text = rowARTCUST1.Item("WHSE_CODE").ToString
        CUST_TERM_CODE = rowARTCUST1.Item("TERM_CODE").ToString
        txtTERM_CODE.Text = rowARTCUST1.Item("TERM_CODE").ToString
        txtFRT_TERMS.Text = rowARTCUST1.Item("FRT_TERMS").ToString
        txtORDR_SHIP_INSTR.Text = rowARTCUST1.Item("CUST_SPECIAL_INST").ToString
        txtORDR_MESSAGE.Text = rowARTCUST1.Item("CUST_ROUTING_INST").ToString
        txtORDR_CUST_PO.Text = ORDR_CUST_PO
        txtORDR_CATEGORY.Text = ORDR_CATEGORY
        SetTierFromCust()
        If BuyerInfoLive Then
            grpBuyerOnfo.Visible = True
        Else
            grpBuyerOnfo.Visible = False
        End If
    End Sub

    Private Sub grdARTCUST2_AfterSelectChange(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdARTCUST2.AfterSelectChange
        If IsNothing(grdARTCUST2.ActiveRow) Then
            SELECT_ST = ""
            CUST_CODE = ""
        Else
            SELECT_ST = grdARTCUST2.ActiveRow.Cells("CUST_ADDR_CODE").Text
            CUST_CODE = grdARTCUST2.ActiveRow.Cells("CUST_CODE").Text
            CUST_ADDR_TYPE = grdARTCUST2.ActiveRow.Cells("CUST_ADDR_TYPE").Text
        End If
        'CHANGES FOR SHIPTO QUESTION - ARTCUSTQ_SET(CUST_CODE, CUST_ADDR_TYPE, SELECT_ST)
    End Sub

    Private Sub ARTCUSTQ_SET(ByVal CUST_CODE As String, ByVal CUST_ADDR_TYPE As String, ByVal CUST_ADDR_CODE As String)
        'CHANGES FOR SHIPTO QUESTION - 
        Exit Sub
        'Dim cFilter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_TYPE = '{1}' AND CUST_ADDR_CODE = '{2}'", CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE)
        'Dim rowARTCUSTQ As DataRow = FF.dst.Tables.Item("ARTCUSTQ").Select(cFilter).FirstOrDefault
        'If IsNothing(rowARTCUSTQ) Then
        '    Dim newARTCUSTQ As DataRow = FF.dst.Tables.Item("ARTCUSTQ").NewRow
        '    newARTCUSTQ.Item("CUST_CODE") = CUST_CODE
        '    newARTCUSTQ.Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
        '    newARTCUSTQ.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
        '    newARTCUSTQ.Item("LAST_DATE") = Now()
        '    newARTCUSTQ.Item("LAST_OPER") = ASCMAIN1.USER_ID
        '    newARTCUSTQ.Item("RESIDENTIAL_ORDR") = "0"
        '    newARTCUSTQ.Item("INSIDE_REQ") = "0"
        '    newARTCUSTQ.Item("DEL_APPOINTMENT_REQ") = "0"
        '    newARTCUSTQ.Item("GATE_LIFT_REQ") = "0"
        '    newARTCUSTQ.Item("PREFERRED_SHIP_VIA") = ""
        '    FF.dst.Tables.Item("ARTCUSTQ").Rows.Add(newARTCUSTQ)
        '    rowARTCUSTQ = newARTCUSTQ
        'End If
        'dteLAST_DATE.Value = rowARTCUSTQ.Item("LAST_DATE").ToString & String.Empty
        'txtLAST_OPER.Text = rowARTCUSTQ.Item("LAST_OPER").ToString & String.Empty
        'chkRESIDENTIAL_ORDR.Checked = Val(rowARTCUSTQ.Item("RESIDENTIAL_ORDR").ToString & String.Empty) = 1
        'chkINSIDE_REQ.Checked = Val(rowARTCUSTQ.Item("INSIDE_REQ").ToString & String.Empty) = 1
        'chkDEL_APPOINTMENT_REQ.Checked = Val(rowARTCUSTQ.Item("DEL_APPOINTMENT_REQ").ToString & String.Empty) = 1
        'chkGATE_LIFT_REQ.Checked = Val(rowARTCUSTQ.Item("GATE_LIFT_REQ").ToString & String.Empty) = 1
        'txtPREFERRED_SHIP_VIA.Text = rowARTCUSTQ.Item("PREFERRED_SHIP_VIA").ToString & String.Empty
    End Sub

    Private Sub cmdDone_Click(sender As System.Object, e As System.EventArgs) Handles cmdDone.Click
        Dim Msg As String = ""
        If grdARTCUST2.Selected.Rows.Count() < 1 Then
            Msg += vbCrLf & "You Didn't Select a Ship-To."
        End If
        If datSTARTDATE.DateTime.ToShortDateString = datCANCELDATE.DateTime.ToShortDateString Then
            Dim iresult As MsgBoxResult = MsgBox("Start Date Is The Same As Cancel Date.", MsgBoxStyle.YesNo, "Is That OK With You?")
            If iresult = MsgBoxResult.No Then
                Msg += vbCrLf & "Please Select A Valid Cancel Date."
            End If
        End If
        If txtSHIP_VIA_CODE.Text.Length = 0 Then
            Msg += vbCrLf & "You Must Select A Valid Ship-Via."
        Else
            If ValidateCode("SHIP_VIA_CODE") <> True Then
                Msg += vbCrLf & "Selected Ship-Via is Invalid."
            End If
        End If
        If txtWHSE_CODE.Text.Length = 0 Then
            Msg += vbCrLf & "You Must Select A Valid Warehouse."
        Else
            If ValidateCode("WHSE_CODE") <> True Then
                Msg += vbCrLf & "Selected Warehouse is Invalid."
            End If
        End If
        If txtTERM_CODE.Text.Length = 0 Then
            Msg += vbCrLf & "You Must Select A Valid Terms Code."
        Else
            If ValidateCode("TERM_CODE") <> True Then
                Msg += vbCrLf & "Selected Terms Code is Invalid."
            End If
        End If
        If txtFRT_TERMS.Text.Length = 0 Then
            Msg += vbCrLf & "You Must Select A Valid Freight Code."
        Else
            If ValidateCode("FRT_TERMS") <> True Then
                Msg += vbCrLf & "Selected Freight Code is Invalid."
            End If
        End If
        'No Longer Manditory at least for the show, DM - 12/19/17
        'If BuyerInfoLive Then
        '    If txtORDR_BUYER_NAME.Text.Length = 0 Then
        '        Msg += vbCrLf & "You Must Select A Buyer."
        '    End If
        '    If txtORDR_BUYER_EMAIL.Text.Length = 0 Then
        '        Msg += vbCrLf & "Buyer E-mail Can Not Be Blank."
        '    End If
        'End If
        'Now Manditory Again
        If BuyerInfoLive Then
            If txtORDR_BUYER_NAME.Text.Length = 0 Then
                Msg += vbCrLf & "You Must Select A Buyer."
            End If
            'If txtORDR_BUYER_EMAIL.Text.Length = 0 Then
            '    Msg += vbCrLf & "Buyer E-mail Can Not Be Blank."
            'End If
        End If

        If Msg.Length > 0 Then
            MsgBox(Msg, MsgBoxStyle.Critical, "Ship-To Selection Issues")
            Exit Sub
        Else
            Dim CCOK As Boolean = False
            'For Each rowTABLE_NAME As DataRow In dst.Tables("TABLE_NAME").Select()
            '    rowTABLE_NAME.Item("COLOUM_NAME") = "XXXXX"
            'Next
            rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_ADDR_TYPE, SELECT_ST})
            SHIP_VIA_CODE = txtSHIP_VIA_CODE.Text
            STARTDATE = datSTARTDATE.Value
            CANCELDATE = datCANCELDATE.Value
            WHSE_CODE = txtWHSE_CODE.Text
            ORDR_SHIP_INSTR = txtORDR_SHIP_INSTR.Text
            ORDR_MESSAGE = txtORDR_MESSAGE.Text
            ORDR_CUST_PO = txtORDR_CUST_PO.Text
            ORDR_CATEGORY = txtORDR_CATEGORY.Text
            TERM_CODE = txtTERM_CODE.Text
            FRT_TERMS = txtFRT_TERMS.Text
            Me.Close()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        SELECT_ST = ""
        txtSHIP_VIA_CODE.Text = ""
        datSTARTDATE.Value = ""
        datCANCELDATE.Value = ""
        txtWHSE_CODE.Text = ""
        txtORDR_CUST_PO.Text = ""
        txtORDR_CATEGORY.Text = ""
        txtTERM_CODE.Text = ""
        txtFRT_TERMS.Text = ""
        txtORDR_SHIP_INSTR.Text = ""
        txtORDR_MESSAGE.Text = ""
        ORDR_CUST_PO = txtORDR_CUST_PO.Text
        ORDR_CATEGORY = txtORDR_CATEGORY.Text
        TERM_CODE = txtTERM_CODE.Text
        FRT_TERMS = txtFRT_TERMS.Text
        Me.Close()
    End Sub

    Private Sub datCANCELDATE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles datCANCELDATE.ValueChanged
        FixDates()
    End Sub

    Private Sub FixDates()
        If datCANCELDATE.Value < datSTARTDATE.Value Then
            datCANCELDATE.Value = datSTARTDATE.Value
        End If
    End Sub

    Private Sub ValidateStart()
        Dim TooFarDate As DateTime = Now.AddMonths(9)
        If datSTARTDATE.Value > TooFarDate Then
            Dim iResult As MsgBoxResult = MsgBox("Selected Date Is Greater Than 9 Months" & vbCrLf & "Is That OK With You?", MsgBoxStyle.YesNo, "Date Check")
            If iResult = MsgBoxResult.No Then
                datSTARTDATE.Value = Now()
                Exit Sub
            End If
        End If
        FixDates()
    End Sub

    Private Sub datSTARTDATE_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs)
        ValidateStart()
    End Sub

    Private Function ValidateCode(ByVal ValCode As String) As Boolean
        Dim RetVal As Boolean = False
        Select Case ValCode
            Case "SHIP_VIA_CODE"
                ASCMAIN1.sql = String.Format("Select Count (*) from SOTSVIA1 where SHIP_VIA_CODE = '{0}' AND NVL(SHIP_VIA_STATUS,'A') = 'A'", txtSHIP_VIA_CODE.Text)
            Case "WHSE_CODE"
                ASCMAIN1.sql = String.Format("Select Count (*) from ICTWHSE1 where WHSE_CODE = '{0}'", txtWHSE_CODE.Text)
            Case "TERM_CODE"
                ASCMAIN1.sql = String.Format("Select Count (*) from TATTERM1 where TERM_CODE = '{0}'", txtTERM_CODE.Text)
                Select Case txtTERM_CODE.Text
                    Case Is = "DISC", "MC", "VISA"
                        txtTERM_CODE.Text = "CRED"
                End Select
            Case "FRT_TERMS"
                ASCMAIN1.sql = "SELECT COUNT(*) FROM ASTCODE1"
                ASCMAIN1.sql = ASCMAIN1.sql & " WHERE COLUMN_NAME = 'FRT_TERMS'"
                ASCMAIN1.sql = ASCMAIN1.sql & " AND TABLE_NAME = 'SOTORDR1'"
                ASCMAIN1.sql = String.Format("{0} AND T_CODE = '{1}'", ASCMAIN1.sql, txtFRT_TERMS.Text)
            Case Else
                Return RetVal
        End Select
        Dim RecCount As Int16 = Val(ASCDATA1.GetDataValue)
        If RecCount > 0 Then
            RetVal = True
        End If
        Return RetVal
    End Function

    Private Sub chkPinDisc_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkPinDisc.CheckedChanged
        If chkPinDisc.Checked Then
            lblPinDisc.Visible = True
            optPinDisc1.Visible = True
            optPinDisc2.Visible = True
        Else
            lblPinDisc.Visible = False
            optPinDisc1.Visible = False
            optPinDisc2.Visible = False
            optPinDisc1.Checked = False
            optPinDisc2.Checked = False
        End If
    End Sub

#Region "rdoQUOTESandOrders"
    Private Sub rdoQUOTE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoQUOTE.CheckedChanged
        If rdoQUOTE.Checked Then
            rdoORDER.Checked = False
        Else
            rdoORDER.Checked = True
        End If
    End Sub

    Private Sub rdoORDER_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rdoORDER.CheckedChanged
        If rdoORDER.Checked Then
            rdoQUOTE.Checked = False
        Else
            rdoQUOTE.Checked = True
        End If
    End Sub

#End Region

    Private Sub txtTERM_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtTERM_CODE.ValueChanged
        If Not IsNothing(CUST_CODE) Then
            If Not SOCMAINL.IsValidTerms(CUST_CODE, txtTERM_CODE.Text) Then
                txtTERM_CODE.Text = CUST_TERM_CODE
            End If
        End If


        'This is moved to SOCMAINL.IsValidTerms now for re-use. WR: 04/15/17
        'Select Case txtTERM_CODE.Text
        '    Case Is = "DISC", "MC", "VISA"
        '        txtTERM_CODE.Text = "CRED"
        'End Select
        'Select Case CUST_TERM_CODE
        '    Case Is = "DISC", "MC", "VISA"
        '        CUST_TERM_CODE = "CRED"
        'End Select
        'If txtTERM_CODE.Text <> CUST_TERM_CODE Then
        '    Select Case CUST_TERM_CODE
        '        Case Is = "N30", "N30D", "N30ROG", "N45D", "N60", "N90", "N90D"
        '            Select Case txtTERM_CODE.Text
        '                Case Is = "N30", "N30D", "N30ROG", "N45D", "N60", "N90", "N90D", "COD", "CBD", "CRED", "XMAS", "FALL"
        '                    'OK
        '                Case Else
        '                    MsgBox("Invalid Terms Code For This Customer", MsgBoxStyle.OkOnly, "Invalid Terms")
        '                    txtTERM_CODE.Text = CUST_TERM_CODE
        '            End Select
        '        Case Is = "CRED", "COD", "CBD"
        '            Select Case txtTERM_CODE.Text
        '                Case Is = "CRED", "COD", "CBD"
        '                    'OK
        '                Case Else
        '                    MsgBox(String.Format("Customers With {0} Terms Code Must Select CRED, COD or CBD Terms", CUST_TERM_CODE), MsgBoxStyle.OkOnly, "Invalid Terms")
        '                    txtTERM_CODE.Text = CUST_TERM_CODE
        '            End Select
        '        Case Is = "AMEX"
        '            MsgBox("AMEX Is No Longer Supported As A Terms Code" & CUST_TERM_CODE, MsgBoxStyle.OkOnly, "Invalid Terms")
        '            txtTERM_CODE.Text = CUST_TERM_CODE
        '        Case Else
        '            MsgBox("Customer Terms Code Is " & CUST_TERM_CODE, MsgBoxStyle.OkOnly, "Invalid Terms")
        '            txtTERM_CODE.Text = CUST_TERM_CODE
        '    End Select
        'End If

        'Rich says don't change dates based on terms anymore - 3/4/13
        'Dim DueDate As Date = Now()
        'Dim BaseDate As Date = Now()
        '' If txtTERM_CODE.Text.Length > 0 Then
        'Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", txtTERM_CODE.Text)
        'If Not IsNothing(rowTATTERM1) Then
        '    DueDate = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, txtTERM_CODE.Text, rowTATTERM1, BaseDate)
        'End If
        'If DueDate <> "#12:00:00 AM#" Then
        '    datSTARTDATE.Value = DueDate.Date.AddMonths(-1)
        '    datCANCELDATE.Value = DueDate.Date.AddDays(-1)
        'End If
        'MsgBox(DueDate)
    End Sub

    Private Sub optPinDisc1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optPinDisc1.CheckedChanged
        SetCUST_PRICE_TIER_PVC()
    End Sub

    Private Sub SetCUST_PRICE_TIER_PVC()
        CUST_PRICE_TIER_PVC = "PC"
        If optPinDisc1.Checked = True Then
            CUST_PRICE_TIER_PVC = "FC"
        Else
            If optPinDisc2.Checked = True Then
                CUST_PRICE_TIER_PVC = "5C"
            End If
        End If
    End Sub

    Private Sub optPinDisc2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optPinDisc2.CheckedChanged
        SetCUST_PRICE_TIER_PVC()
    End Sub

    Private Sub SetTierFromCust()
        Dim CUST_TIER As String = rowARTCUST1.Item("CUST_PRICE_TIER_PVC").ToString & ""
        If CUST_TIER = "" Then
            CUST_TIER = "PC"
        End If
        Select Case CUST_TIER
            Case Is = "5C"
                chkPinDisc.Checked = True
                lblPinDisc.Visible = True
                optPinDisc1.Visible = True
                optPinDisc2.Visible = True
                optPinDisc2.Checked = True
            Case Is = "FC"
                chkPinDisc.Checked = True
                lblPinDisc.Visible = True
                optPinDisc1.Visible = True
                optPinDisc2.Visible = True
                optPinDisc1.Checked = True
            Case Else
                chkPinDisc.Checked = False
                lblPinDisc.Visible = False
                optPinDisc1.Visible = False
                optPinDisc2.Visible = False
        End Select


    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "TERM_CODE"
                sql_where = "TERM_STATUS = 'A'"
            Case "SHIP_VIA_CODE"
                sql_where &= "NVL(SHIP_VIA_STATUS,'A') = 'A'"
        End Select
    End Sub

    Private Sub txtWHSE_CODE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtWHSE_CODE.ValueChanged
        Select Case txtWHSE_CODE.Text
            Case Is = "FD"
                txtFRT_TERMS.Text = "PPD"
            Case Is = "FE"
                txtFRT_TERMS.Text = ""
        End Select
    End Sub

    Private Sub txtFRT_TERMS_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtFRT_TERMS.ValueChanged
        If txtFRT_TERMS.Text <> "" Then
            If txtWHSE_CODE.Text = "FE" Then
                If txtFRT_TERMS.Text <> "PPA" And txtFRT_TERMS.Text <> "COL" Then
                    MsgBox("Selection of FE Warehouse Requires Either PPA or COL Freight Terms!", MsgBoxStyle.OkOnly, "Invalid Terms")
                    txtFRT_TERMS.Text = ""
                End If
            End If
            If txtWHSE_CODE.Text = "FD" Then
                If txtFRT_TERMS.Text <> "PPD" Then
                    MsgBox("Selection of FD Warehouse Requires PPD Freight Terms!", MsgBoxStyle.OkOnly, "Invalid Terms")
                    txtFRT_TERMS.Text = "PPD"
                End If
            End If
        End If
    End Sub

    Private Sub btnBuyer_Click(sender As Object, e As EventArgs) Handles btnBuyer.Click
        Dim _CUST_CODE As String = ""
        If FF.dst.Tables("ARTCUST2").Rows.Count > 0 Then
            _CUST_CODE = grdARTCUST2.Rows.Item(0).Cells("CUST_CODE").Text
        End If
        If _CUST_CODE.Length > 0 Then
            Dim frmSOFORDRB As New SOFORDRB(FF, _CUST_CODE)
            With frmSOFORDRB
                .ShowDialog()
                ORDR_BUYER_NAME = .ORDR_BUYER_NAME
                txtORDR_BUYER_NAME.Text = ORDR_BUYER_NAME

                ORDR_BUYER_EMAIL = .ORDR_BUYER_EMAIL
                txtORDR_BUYER_EMAIL.Text = ORDR_BUYER_EMAIL

                ORDR_BUYER_CONTACT_NO = Val(.ORDR_BUYER_CONTACT_NO)
                HAS_CONTACT_CHANGES = .HAS_CONTACT_CHANGES
            End With
        End If
    End Sub

    Private Sub grdARTCUST2_BeforeSelectChange(sender As Object, e As BeforeSelectChangeEventArgs) Handles grdARTCUST2.BeforeSelectChange
        'CHANGES FOR SHIPTO QUESTION - ARTCUSTQ_SAVE()
    End Sub

    Private Sub ARTCUSTQ_SAVE()
        Exit Sub
        'CHANGES FOR SHIPTO QUESTION - 
        'If grdARTCUST2.Selected.Rows.Count = 1 Then
        '    Dim CUST_CODE As String = grdARTCUST2.Selected.Rows(0).Cells("CUST_CODE").Text & String.Empty
        '    Dim CUST_ADDR_TYPE As String = grdARTCUST2.Selected.Rows(0).Cells("CUST_ADDR_TYPE").Text & String.Empty
        '    Dim CUST_ADDR_CODE As String = grdARTCUST2.Selected.Rows(0).Cells("CUST_ADDR_CODE").Text & String.Empty
        '    Dim cFilter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_TYPE = '{1}' AND CUST_ADDR_CODE = '{2}'", CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE)
        '    Dim rowARTCUSTQ As DataRow = FF.dst.Tables.Item("ARTCUSTQ").Select(cFilter).FirstOrDefault
        '    If IsNothing(rowARTCUSTQ) Then
        '        Dim newARTCUSTQ As DataRow = FF.dst.Tables.Item("ARTCUSTQ").NewRow
        '        newARTCUSTQ.Item("CUST_CODE") = CUST_CODE
        '        newARTCUSTQ.Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
        '        newARTCUSTQ.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
        '        newARTCUSTQ.Item("LAST_DATE") = Now()
        '        newARTCUSTQ.Item("LAST_OPER") = ASCMAIN1.USER_ID
        '        If chkRESIDENTIAL_ORDR.Checked Then
        '            newARTCUSTQ.Item("RESIDENTIAL_ORDR") = "1"
        '        Else
        '            newARTCUSTQ.Item("RESIDENTIAL_ORDR") = "0"
        '        End If
        '        If chkINSIDE_REQ.Checked Then
        '            newARTCUSTQ.Item("INSIDE_REQ") = "1"
        '        Else
        '            newARTCUSTQ.Item("INSIDE_REQ") = "0"
        '        End If
        '        If chkDEL_APPOINTMENT_REQ.Checked Then
        '            newARTCUSTQ.Item("DEL_APPOINTMENT_REQ") = "1"
        '        Else
        '            newARTCUSTQ.Item("DEL_APPOINTMENT_REQ") = "0"
        '        End If
        '        If chkGATE_LIFT_REQ.Checked Then
        '            newARTCUSTQ.Item("GATE_LIFT_REQ") = "1"
        '        Else
        '            newARTCUSTQ.Item("GATE_LIFT_REQ") = "0"
        '        End If
        '        newARTCUSTQ.Item("PREFERRED_SHIP_VIA") = txtPREFERRED_SHIP_VIA.Text
        '        FF.dst.Tables.Item("ARTCUSTQ").Rows.Add(newARTCUSTQ)
        '        rowARTCUSTQ = newARTCUSTQ
        '    Else
        '        rowARTCUSTQ.Item("LAST_DATE") = Now()
        '        rowARTCUSTQ.Item("LAST_OPER") = ASCMAIN1.USER_ID
        '        If chkRESIDENTIAL_ORDR.Checked Then
        '            rowARTCUSTQ.Item("RESIDENTIAL_ORDR") = "1"
        '        Else
        '            rowARTCUSTQ.Item("RESIDENTIAL_ORDR") = "0"
        '        End If
        '        If chkINSIDE_REQ.Checked Then
        '            rowARTCUSTQ.Item("INSIDE_REQ") = "1"
        '        Else
        '            rowARTCUSTQ.Item("INSIDE_REQ") = "0"
        '        End If
        '        If chkDEL_APPOINTMENT_REQ.Checked Then
        '            rowARTCUSTQ.Item("DEL_APPOINTMENT_REQ") = "1"
        '        Else
        '            rowARTCUSTQ.Item("DEL_APPOINTMENT_REQ") = "0"
        '        End If
        '        If chkGATE_LIFT_REQ.Checked Then
        '            rowARTCUSTQ.Item("GATE_LIFT_REQ") = "1"
        '        Else
        '            rowARTCUSTQ.Item("GATE_LIFT_REQ") = "0"
        '        End If
        '        rowARTCUSTQ.Item("PREFERRED_SHIP_VIA") = txtPREFERRED_SHIP_VIA.Text
        '    End If
        'End If
    End Sub

End Class