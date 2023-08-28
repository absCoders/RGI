Public Class ICFATOP2

    Public calculatedCube As Double = 0
    Public STYLE_CODE As String
    Public COLOR_CODE As String
    Public PS_CODE As String
    Public PS_NO As String
    Public PS_ETA As Date
    Public rowICTATOP2 As DataRow

    Private Sub ICFATOP2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")

        AUDIT.Add("ICTATOP2", "*")

        With dst
            Create_TDA(.Tables.Add, "ICTATOP2", "*")
        End With

        Set_Read_Only(Me.UltraGroupBox1, True)

        rowICTATOP2 = Fill_Record("ICTATOP2", New String() {STYLE_CODE, COLOR_CODE, PS_CODE, PS_NO})
        If rowICTATOP2 Is Nothing Then
            rowICTATOP2 = dst.Tables("ICTATOP2").NewRow
            With rowICTATOP2
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("PS_CODE") = PS_CODE
                .Item("PS_NO") = PS_NO
                .Item("PS_ETA") = PS_ETA

                .Item("STYLE_ARRIVAL_BUFFER_DAYS") = ROWs("SOTPARM1").Item("SO_PARM_ARRIVAL_BUFFER_DAYS")

                .Item("STYLE_AT_ONCE_UNTIL") = DBNull.Value
                .Item("STYLE_AT_ONCE_ACTIVE") = "1"

                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                .Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
            End With

            dst.Tables("ICTATOP2").Rows.Add(rowICTATOP2)
        End If
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        Dim EMsg As String = ""
        If Absx1.dteFor("STYLE_AT_ONCE_UNTIL").Value & "" = "" Then
            EMsg &= vbCr & "You must provide a value for Date Until"
        End If
        If Val(Absx1.numFor("STYLE_ARRIVAL_BUFFER_DAYS").Value & "") <= 0 Then
            EMsg &= vbCr & "Arrival+ Days must be a positive number"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update Record")
            Exit Sub
        End If

        If rowICTATOP2.RowState = DataRowState.Added Then
        Else

            rowICTATOP2.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            rowICTATOP2.Item("LAST_OPER") = ASCMAIN1.USER_ID
        End If

        Update_Record_TDA("ICTATOP2")
        DialogResult = Windows.Forms.DialogResult.OK
        Hide()
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
        Hide()
    End Sub
End Class