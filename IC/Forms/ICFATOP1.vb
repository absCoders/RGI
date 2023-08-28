Public Class ICFATOP1

    Public calculatedCube As Double = 0
    Public STYLE_CODE As String
    Public COLOR_CODE As String
    Public ORDR_TYPE As String
    Public ORDR_NO As String
    Public ORDR_SHIP_DATE As Date
    Public ORDR_CANCEL_DATE As Date
    Public rowICTATOP1 As DataRow
    Dim ready As Boolean = False

    Private Sub ICFATOP1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")

        AUDIT.Add("ICTATOP1", "*")

        With dst
            Create_TDA(.Tables.Add, "ICTATOP1", "*")
        End With

        Set_Read_Only(Me.UltraGroupBox1, True)

        rowICTATOP1 = Fill_Record("ICTATOP1", New String() {STYLE_CODE, COLOR_CODE, ORDR_TYPE, ORDR_NO})
        If rowICTATOP1 Is Nothing Then
            rowICTATOP1 = dst.Tables("ICTATOP1").NewRow
            With rowICTATOP1
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("ORDR_TYPE") = ORDR_TYPE
                .Item("ORDR_NO") = ORDR_NO

                .Item("ORDR_SHIP_DATE_ORIG") = ORDR_SHIP_DATE
                .Item("ORDR_CANCEL_DATE_ORIG") = ORDR_CANCEL_DATE

                .Item("STYLE_SHIP_WINDOW_DAYS") = ROWs("SOTPARM1").Item("SO_PARM_SHIP_WINDOW_DAYS")

                .Item("STYLE_AT_ONCE_UNTIL") = DBNull.Value
                .Item("STYLE_AT_ONCE_ACTIVE") = "1"

                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                .Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
            End With


            dst.Tables("ICTATOP1").Rows.Add(rowICTATOP1)
        End If

        Absx1.dteFor("ORDR_SHIP_DATE").Value = ORDR_SHIP_DATE

        If ORDR_TYPE = "O" Then
            UltraGroupBox1.Text = "Order No"
        Else
            UltraGroupBox1.Text = "Reservation No"
        End If

        ready = True
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        Dim EMsg As String = ""
        If Absx1.dteFor("STYLE_AT_ONCE_UNTIL").Value & "" = "" Then
            EMsg &= vbCr & "You must provide a value for Date Until"
        End If
        If Val(Absx1.numFor("STYLE_SHIP_WINDOW_DAYS").Value & "") <= 0 Then
            EMsg &= vbCr & "Arrival+ Days must be a positive number"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update Record")
            Exit Sub
        End If

        If rowICTATOP1.RowState = DataRowState.Added Then
        Else

            rowICTATOP1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            rowICTATOP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        End If

        Update_Record_TDA("ICTATOP1")
        DialogResult = Windows.Forms.DialogResult.OK
        Hide()
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
        Hide()
    End Sub

    Private Sub numSTYLE_SHIP_WINDOW_DAYS_ValueChanged(sender As Object, e As EventArgs) Handles numSTYLE_SHIP_WINDOW_DAYS.ValueChanged
        If ready AndAlso dteORDR_SHIP_DATE_PLUS.Tag = "X" Then Exit Sub
        numSTYLE_SHIP_WINDOW_DAYS.Tag = "X"
        Absx1.dteFor("ORDR_SHIP_DATE_PLUS").Value = CDate(Absx1.dteFor("ORDR_SHIP_DATE").Value).AddDays(Val(numSTYLE_SHIP_WINDOW_DAYS.Value & ""))
        numSTYLE_SHIP_WINDOW_DAYS.Tag = ""
    End Sub

    Private Sub dteORDR_SHIP_DATE_PLUS_ValueChanged(sender As Object, e As EventArgs) Handles dteORDR_SHIP_DATE_PLUS.ValueChanged
        If ready AndAlso numSTYLE_SHIP_WINDOW_DAYS.Tag = "X" Then Exit Sub
        dteORDR_SHIP_DATE_PLUS.Tag = "X"
        Dim STYLE_SHIP_WINDOW_DAYS As Integer = CDate(Absx1.dteFor("ORDR_SHIP_DATE_PLUS").Value).Subtract(CDate(Absx1.dteFor("ORDR_SHIP_DATE").Value)).TotalDays
        If STYLE_SHIP_WINDOW_DAYS >= 0 Then
            numSTYLE_SHIP_WINDOW_DAYS.Value = STYLE_SHIP_WINDOW_DAYS
        Else
            Absx1.dteFor("ORDR_SHIP_DATE_PLUS").Value = CDate(Absx1.dteFor("ORDR_SHIP_DATE").Value)
        End If
        dteORDR_SHIP_DATE_PLUS.Tag = ""
    End Sub
End Class