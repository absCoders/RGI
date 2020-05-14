Public Class TATSTATE

    ' ALTER TABLE TATSTATE ADD STATE_REL_PCT NUMBER (3)

    Private Sub TATSTATE_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        With dst
            Call Create_TDA(.Tables.Add, "TATSHIPP", "*", 2)
            grdTATSHIPP.DataSource = dst.Tables("TATSHIPP")
        End With

        numSTATE_REL_PCT.Visible = (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI")
        lblSTATE_REL_PCT.Visible = (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI")
        grdTATSHIPP.Visible = (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI")

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()
        Update_Record_TDA("TATSHIPP", "TABLE_NAME = 'TATSTATE' AND KEY_VALUE = '" & Absx1.txtFor("STATE_CODE").Text & "'")
    End Sub

    Overrides Sub Show_Record_Special()
        Fill_Records("TATSHIPP", New String() {"TATSTATE", Absx1.txtFor("STATE_CODE").Text})
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

#End Region

    Private Sub grdTATSHIPP_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdTATSHIPP.BeforeRowUpdate
        e.Row.Cells("TABLE_NAME").Value = "TATSTATE"
        e.Row.Cells("KEY_VALUE").Value = Absx1.txtFor("STATE_CODE").Text

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