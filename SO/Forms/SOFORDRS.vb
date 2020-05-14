
Imports System.ComponentModel
Imports Infragistics.Win.UltraWinGrid

Public Class SOFORDRS
    Private FF As ASFBASE1
    Private CUST_CODE As String
    Private ORDR_NO As String
    Private SQL As New Text.StringBuilder With {.Length = 0}

#Region "Standard Methods"
    Public Sub New(ByVal frmASFBASE1 As ASFBASE1, ByVal _CUST_CODE As String, ByVal _ORDR_NO As String)
        FF = frmASFBASE1
        CUST_CODE = _CUST_CODE
        ORDR_NO = _ORDR_NO
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If CUST_CODE.Length = 0 Then
        '    Me.Close()
        'End If
        'ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        'FF.Fill_Records("ARTCUSTD", CUST_CODE)
        'grdARTCUSTD.DataSource = FF.dst.Tables("ARTCUSTD")

        grdARTCUSX2.DataSource = FF.dst.Tables("ARTCUSX2")
        With grdARTCUSX2.DisplayLayout.Override
            .AllowAddNew = AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False
        End With
    End Sub

#End Region

#Region "Form Controls"

    Private Sub btnDone_Click(sender As Object, e As EventArgs) Handles btnDone.Click
        If Not grdARTCUSX2.ActiveRow Is Nothing Then
            Dim CC As String = grdARTCUSX2.ActiveRow.Cells("CUST_ADDR_CODE").Value
            updateControls(CC)
        End If

        Dim EMsg As String = Error_Checks()
        If EMsg.Length > 0 Then
            MsgBox(EMsg.ToString, vbOKOnly, "Problems")
        Else
            Me.Close()
        End If
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

    Private Sub chkAPPOINTMENT_REQUIRED_NOTE_CheckedChanged(sender As Object, e As EventArgs) Handles chkAPPOINTMENT_REQUIRED_NOTE.CheckedChanged
        If chkAPPOINTMENT_REQUIRED_NOTE.Checked Then
            txtAPPOINTMENT_REQUIRED_NOTE.Visible = True
        Else
            txtAPPOINTMENT_REQUIRED_NOTE.Text = ""
            txtAPPOINTMENT_REQUIRED_NOTE.Visible = False
        End If
    End Sub
#End Region

#Region "grdARTCUSX2"
    Private Sub grdARTCUSX2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdARTCUSX2.InitializeRow
        Dim CC As String = e.Row.Cells("CUST_ADDR_CODE").Value
        setVerifiedColor(e.Row, CC, False)
        bindControl(CC)
        grdARTCUSX2.UpdateData()

    End Sub

    Private Sub grdARTCUSX2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUSX2.AfterRowActivate
        If Not grdARTCUSX2.ActiveRow Is Nothing Then
            Dim CC As String = grdARTCUSX2.ActiveRow.Cells("CUST_ADDR_CODE").Value
            setVerifiedColor(grdARTCUSX2.ActiveRow, CC, True)
            bindControl(CC)
            grdARTCUSX2.UpdateData()
        End If
    End Sub

    Private Sub setVerifiedColor(ByRef rw As UltraGridRow, ByVal CC As String, ByVal SetVerified As Boolean)
        Dim Filter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CC)
        Dim LAST_ORDR_NO As String = ""
        If FF.dst.Tables("ARTCUSTQ").Select(Filter).Count = 1 Then
            Dim rowARTCUSTQ As DataRow = FF.dst.Tables("ARTCUSTQ").Select(Filter).FirstOrDefault
            If SetVerified Then
                rowARTCUSTQ.Item("LAST_ORDR_NO") = ORDR_NO
                LAST_ORDR_NO = ORDR_NO
            Else
                LAST_ORDR_NO = rowARTCUSTQ.Item("LAST_ORDR_NO").ToString & String.Empty
            End If
        End If
        If LAST_ORDR_NO.Length = 0 Then
            rw.Cells("VERIFIED").Value = "UnVerified"
            rw.Cells("VERIFIED").Appearance.BackColor = Drawing.Color.Red
            rw.Cells("VERIFIED").Appearance.ForeColor = Drawing.Color.White
        Else
            rw.Cells("VERIFIED").Value = "Verified"
            rw.Cells("VERIFIED").Appearance.BackColor = Drawing.Color.Transparent
            rw.Cells("VERIFIED").Appearance.ForeColor = Drawing.Color.Black
        End If
    End Sub

    Private Sub grdARTCUSX2_BeforeRowDeactivate(sender As Object, e As CancelEventArgs) Handles grdARTCUSX2.BeforeRowDeactivate
        Dim CC As String = grdARTCUSX2.ActiveRow.Cells("CUST_ADDR_CODE").Value
        updateControls(CC)
        Dim EMsg As String = Error_Checks()
        If EMsg.Length > 0 Then
            MsgBox(EMsg.ToString, vbOKOnly, "Problems")
            e.Cancel = True
        End If
    End Sub

#End Region

#Region "Custom Methods"
    Private Sub bindControl(ByVal CC As String)
        Dim Filter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CC)
        If FF.dst.Tables("ARTCUSTQ").Select(Filter).Count = 1 Then
            Dim rowARTCUSTQ As DataRow = FF.dst.Tables("ARTCUSTQ").Select(Filter).FirstOrDefault
            If IsDate(rowARTCUSTQ.Item("LAST_DATE").ToString) Then
                dteLAST_DATE.Value = CDate(rowARTCUSTQ.Item("LAST_DATE").ToString)
            Else
                dteLAST_DATE.Value = Null
            End If
            txtLAST_OPER.Value = rowARTCUSTQ.Item("LAST_OPER").ToString
            chkRESIDENTIAL_ORDR.Checked = rowARTCUSTQ.Item("RESIDENTIAL_ORDR").ToString & String.Empty = "1"
            chkINSIDE_REQ.Checked = rowARTCUSTQ.Item("INSIDE_REQ").ToString & String.Empty = "1"
            chkAPPOINTMENT_REQUIRED_NOTE.Checked = rowARTCUSTQ.Item("APPOINTMENT_REQUIRED").ToString & String.Empty = "1"
            chkGATE_LIFT_REQ.Checked = rowARTCUSTQ.Item("GATE_LIFT_REQ").ToString & String.Empty = "1"
            chkLIMITED_ACCESS.Checked = rowARTCUSTQ.Item("LIMITED_ACCESS").ToString & String.Empty = "1"
            chkIRREGULAR_HOURS.Checked = rowARTCUSTQ.Item("IRREGULAR_HOURS").ToString & String.Empty = "1"
            chkBROKER.Checked = rowARTCUSTQ.Item("BROKER").ToString & String.Empty = "1"
            chkRESIDENTIAL_ORDR.Checked = rowARTCUSTQ.Item("RESIDENTIAL_ORDR").ToString & String.Empty = "1"
            chkRESIDENTIAL_ORDR.Checked = rowARTCUSTQ.Item("RESIDENTIAL_ORDR").ToString & String.Empty = "1"
            chkRESIDENTIAL_ORDR.Checked = rowARTCUSTQ.Item("RESIDENTIAL_ORDR").ToString & String.Empty = "1"
            txtLIMITED_ACCESS_NOTE.Text = rowARTCUSTQ.Item("LIMITED_ACCESS_NOTE").ToString & String.Empty
            txtIRREGULAR_HOURS_NOTE.Text = rowARTCUSTQ.Item("IRREGULAR_HOURS_NOTE").ToString & String.Empty
            txtAPPOINTMENT_REQUIRED_NOTE.Text = rowARTCUSTQ.Item("APPOINTMENT_REQUIRED_NOTE").ToString & String.Empty
            txtBROKER_NOTE.Text = rowARTCUSTQ.Item("BROKER_NOTE").ToString & String.Empty
        End If
    End Sub

    Private Function Error_Checks() As String
        Dim retVal As New Text.StringBuilder With {.Length = 0}
        If chkLIMITED_ACCESS.Checked And txtLIMITED_ACCESS_NOTE.Text.Length = 0 Then
            retVal.AppendLine("Limited Hours Checked With No Note.")
        End If

        If chkIRREGULAR_HOURS.Checked And txtIRREGULAR_HOURS_NOTE.Text.Length = 0 Then
            retVal.AppendLine("Irregular Hours Checked With No Note.")
        End If

        If chkBROKER.Checked And txtBROKER_NOTE.Text.Length = 0 Then
            retVal.AppendLine("Broker Checked With No Note.")
        End If

        If chkAPPOINTMENT_REQUIRED_NOTE.Checked And txtAPPOINTMENT_REQUIRED_NOTE.Text.Length = 0 Then
            retVal.AppendLine("Appt Required Checked With No Note.")
        End If
        Return retVal.ToString
    End Function

    Private Sub updateControls(CC)
        Dim Filter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CC)
        If FF.dst.Tables("ARTCUSTQ").Select(Filter).Count = 1 Then
            Dim rowARTCUSTQ As DataRow = FF.dst.Tables("ARTCUSTQ").Select(Filter).FirstOrDefault
            rowARTCUSTQ.Item("LAST_OPER") = ASCMAIN1.USER_ID 'txtLAST_OPER.Value
            rowARTCUSTQ.Item("LAST_ORDR_NO") = ORDR_NO
            rowARTCUSTQ.Item("LAST_DATE") = CDate(Now().ToShortDateString)
            rowARTCUSTQ.Item("RESIDENTIAL_ORDR") = If(chkRESIDENTIAL_ORDR.Checked, "1", "0")
            rowARTCUSTQ.Item("INSIDE_REQ") = If(chkINSIDE_REQ.Checked, "1", "0")
            rowARTCUSTQ.Item("GATE_LIFT_REQ") = If(chkGATE_LIFT_REQ.Checked, "1", "0")
            rowARTCUSTQ.Item("LIMITED_ACCESS") = If(chkLIMITED_ACCESS.Checked, "1", "0")
            rowARTCUSTQ.Item("IRREGULAR_HOURS") = If(chkIRREGULAR_HOURS.Checked, "1", "0")
            rowARTCUSTQ.Item("APPOINTMENT_REQUIRED") = If(chkAPPOINTMENT_REQUIRED_NOTE.Checked, "1", "0")
            rowARTCUSTQ.Item("BROKER") = If(chkBROKER.Checked, "1", "0")
            rowARTCUSTQ.Item("LIMITED_ACCESS_NOTE") = txtLIMITED_ACCESS_NOTE.Text
            rowARTCUSTQ.Item("IRREGULAR_HOURS_NOTE") = txtIRREGULAR_HOURS_NOTE.Text
            rowARTCUSTQ.Item("APPOINTMENT_REQUIRED_NOTE") = txtAPPOINTMENT_REQUIRED_NOTE.Text
            rowARTCUSTQ.Item("BROKER_NOTE") = txtBROKER_NOTE.Text
        End If
    End Sub
#End Region
End Class