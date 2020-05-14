Public Class ASFTEST1
    Dim dst As DataSet

    Private Sub UltraButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton1.Click

        ASCMAIN1.sql = "Select * from ASTSECM1"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        dst.Tables.Add(tbl)

        Dim RPT_FILENAME As String = ASCMAIN1.Folders("Reports") & "ASRTEST1.RPT"
        If ASCMAIN1.Running_in_VS Then
            Dim XSD_FILENAME As String = ASCMAIN1.Folders("Temp") & "ASFTEST1.XSD"
            If Not My.Computer.FileSystem.FileExists(XSD_FILENAME) Then
                dst.WriteXml(XSD_FILENAME, XmlWriteMode.WriteSchema)
            End If
        End If

        ASCMAIN1.CR_RPT.Load(RPT_FILENAME)

        ASCMAIN1.CR_RPT.SetDataSource(dst)

        For Each sr As CrystalDecisions.CrystalReports.Engine.ReportDocument In ASCMAIN1.CR_RPT.Subreports
            Try
                sr.SetDataSource(dst)
            Catch ex As Exception
                If ASCMAIN1.Running_in_VS Then
                    Stop
                End If
            End Try
        Next

        Dim REPORT_NO As String = ASCMAIN1.Next_Control_No("ASTSPRF1.REPORT_NO")
        Dim FILENAME As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO & ".RPT"
        Dim DestOpt As New CrystalDecisions.Shared.DiskFileDestinationOptions
        DestOpt.DiskFileName = ASCMAIN1.Folders("Temp") & filename

        With ASCMAIN1.CR_RPT.ExportOptions
            .DestinationOptions = DestOpt
            .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport

            'Select Case ExportFormat
            '    Case "RPT"
            '        .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.CrystalReport
            '    Case "PDF"
            '        .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
            'End Select
        End With













        Dim crv As New CrystalDecisions.Windows.Forms.CrystalReportViewer
        ' Add a Crystal Report Viewer to the Tab Page Control & Configure it
        SplitContainer1.Panel2.Controls.Add(crv)
        crv.ActiveViewIndex = -1
        crv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        crv.Dock = System.Windows.Forms.DockStyle.Fill
        crv.BackColor = System.Drawing.Color.FromArgb(222, 223, 206)

        Dim REPORT_FILENAME As String = ASCMAIN1.Folders("Temp") & FILENAME

        Dim RPT As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        Try
            RPT.Load(REPORT_FILENAME)
            crv.ReportSource = RPT

        Catch ex As Exception
            MsgBox("Problem Report: " & ASCMAIN1.CR_RPT.FileName & vbCr & vbCr & ex.Message, MsgBoxStyle.OkOnly, "Cannot Load Report " & REPORT_NO)

        End Try

    End Sub

    Private Sub UltraButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraButton2.Click

        Dim F As ASFBASE1 = ASCMAIN1.ActiveForm
        F.remotely_controlled = True
        If Not F.ScreenMode Then
            F.Click_Command("Cancel")
        End If

        ASCMAIN1.sql = "SELECT * FROM MM_TAX_ADJ"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            F.Absx1.txtFor("CUST_CODE").Text = row.Item("CUST")
            F.Absx1.txtFor("CUST_SHIP_TO_NO").Text = row.Item("SHIP_TO") & ""
            Dim TAX_ADJ As Decimal = Val(row.Item("TAX_ADJ"))
            F.Absx1.optFor("INV_TYPE").CheckedIndex = IIf(TAX_ADJ < 0, 0, 1)
            F.Click_Command("New")
            If F.ScreenMode Then
                F.Absx1.numFor("INV_STAX").Value = TAX_ADJ
                F.Absx1.txtFor("STAX_CODE").Text = row.Item("STAX_CD")
                F.Absx1.txtFor("INV_NOTES").Text = "Adj tax doc#" & row.Item("INV_NO")
                F.Absx1.txtFor("ORDR_CUST_PO").Text = row.Item("INV_NO")
                'Dim X As MsgBoxResult = MsgBox("UPDATE", MsgBoxStyle.YesNoCancel, "VERIFICATION")
                'If X = MsgBoxResult.Cancel Then
                '    Exit Sub
                'ElseIf X = MsgBoxResult.Yes Then
                '    F.Click_Command("Update")
                'End If
                F.Click_Command("Update")
            Else
                MsgBox("X", MsgBoxStyle.OkOnly, "X")
            End If
        Next
    End Sub
End Class