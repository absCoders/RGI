Public Class SOTPARM1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "TATSHIPP", "*", 2)
        End With

        grdTATSHIPP.DataSource = dst.Tables("TATSHIPP")
        grdTATSHIPP.Visible = (ASCMAIN1.CLIENT = "RGI")

        numSO_PARM_GROUP_DAYS.Visible = (ASCMAIN1.CLIENT = "RGI")
        numSO_PARM_REL_CUBE.Visible = (ASCMAIN1.CLIENT = "RGI")
        lblDays.Visible = (ASCMAIN1.CLIENT = "RGI")
        lblCube.Visible = (ASCMAIN1.CLIENT = "RGI")

        grpReleaseAtOnce.Visible = (ASCMAIN1.CLIENT = "RGI")
        grpReleaseAtOnce.Visible = (ASCMAIN1.CLIENT = "RGI") And (ASCMAIN1.USER_ID = "rich" Or ASCMAIN1.USER_ID = "danny" Or ASCMAIN1.USER_ID = "wjz")
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Update_Record_TDA("TATSHIPP", "TABLE_NAME = 'SOTPARM1' AND KEY_VALUE = '" & Absx1.txtFor("SO_PARM_KEY").Text & "'")
    End Sub

    Overrides Sub Show_Record_Special()
        Fill_Records("TATSHIPP", New String() {"SOTPARM1", Absx1.txtFor("SO_PARM_KEY").Text})
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"TATSHIPP"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)

        End If
    End Sub

    Private Sub grdTATSHIPP_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdTATSHIPP.BeforeRowUpdate
        e.Row.Cells("TABLE_NAME").Value = "SOTPARM1"
        e.Row.Cells("KEY_VALUE").Value = Absx1.txtFor("SO_PARM_KEY").Text

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

End Class