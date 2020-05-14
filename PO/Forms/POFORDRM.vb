Public Class POFORDRM
    Public STYLE_CODE As String
    Public COLOR_CODE As String
    Public rowPOTORDR2 As DataRow

    Public CONSUMPTION As Decimal = 0
    Public TOTAL_COST As Decimal = 0
    Public FABRIC_COST As Decimal = 0

    'Public Form_Caption As String = ""
    Public ok2update As Boolean = False

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        With dst

            With .Tables.Add("POTORDRC")
                .Columns.Add("CONSUMPTION", GetType(System.Decimal))
                .Columns.Add("COST_PER_YD", GetType(System.Decimal))
                .Columns.Add("TOTAL_COST", GetType(System.Decimal), "CONSUMPTION * COST_PER_YD")
            End With
        End With

        grdPOTORDRC.DataSource = dst.Tables("POTORDRC")
    
        Create_Summary(grdPOTORDRC, New String() {"CONSUMPTION", "TOTAL_COST"})
        Create_Summary(grdPOTORDRC, "COST_PER_YD", "Custom")

        STYLE_CODE = rowPOTORDR2.Item("STYLE_CODE") & ""
        COLOR_CODE = rowPOTORDR2.Item("COLOR_CODE") & ""

        Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
        Absx1.txtFor("COLOR_CODE").Text = COLOR_CODE

        With grdPOTORDRC.DisplayLayout.Bands(0)
            .Columns("TOTAL_COST").CellActivation = UltraWinGrid.Activation.NoEdit
        End With
    End Sub

    Sub Use_Calculated_Cost()
        Dim CONSUMPTION As Decimal = Val(dst.Tables("POTORDRC").Compute("SUM(CONSUMPTION)", "") & "")
        Dim TOTAL_COST As Decimal = Val(dst.Tables("POTORDRC").Compute("SUM(TOTAL_COST)", "") & "")
        Dim PO_ORDER_LNO As Integer = Val(rowPOTORDR2.Item("PO_ORDER_LNO") & "")
        If TOTAL_COST <> 0 Then
            If MsgBox("Do You Want to Replace Line " & CStr(PO_ORDER_LNO) & " with the Calculated Values?", _
                    MsgBoxStyle.YesNo, "Use Calculations?") = MsgBoxResult.Yes Then
                Dim COST_PER_YD As Decimal = TOTAL_COST / CONSUMPTION
                Absx1.numFor("POTORDR2_LINE.YARDS_CONSUMED").Value = CONSUMPTION
                Absx1.numFor("POTORDR2_LINE.FABRIC_COST").Value = COST_PER_YD
                Absx1.numFor("POTORDR2_LINE.PO_COST_MATLS_DZ").Value = TOTAL_COST
                ' ReCalculate_PO_Cost()
            End If
        End If
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "STYLE_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Prepare_ICTCOLRM()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "STYLE_CODE"
            '    Prepare_ICTCOLRM()
        End Select
    End Sub
#End Region

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        CONSUMPTION = Val(dst.Tables("POTORDRC").Compute("SUM(CONSUMPTION)", "") & "")
        TOTAL_COST = Val(dst.Tables("POTORDRC").Compute("SUM(TOTAL_COST)", "") & "")
        FABRIC_COST = 0
        If CONSUMPTION <> 0 Then FABRIC_COST = TOTAL_COST / CONSUMPTION
        ok2update = True
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
    End Sub


    Overrides Sub CustomSummary_DataRows( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal row As UltraWinGrid.UltraGridRow, _
    ByRef CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid)

        Select Case grd.Name
            ' Case "grdPOTORDRC"
            'If summarySettings.Key = "CONSUMPTION" Then
            '    CustomValue += Val(row.Cells("CONSUMPTION").Value & "")
            '    'Dim GROSS As Object = row.GetCellValue(summarySettings.SourceColumn.Band.Columns("GROSS"))
            'End If
            'If summarySettings.Key = "TOTAL_COST" Then
            '    CustomValue += Val(row.Cells("TOTAL_COST").Value & "")
            'End If
        End Select
    End Sub

    Overrides Function CustomSummary_End( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByVal CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid) As Double

        Select Case grd.Name
            Case "grdPOTORDRC"
                If summarySettings.Key = "COST_PER_YD" Then
                    ' THESE EXPRESSIONS DO NOT GIVE ACCURATE PRESENT VALUES - ITS LIKE THEY ARE AN UPDATE BEHIND
                    Dim CONSUMPTION As Decimal = Val(rows.SummaryValues("CONSUMPTION").Value & "")
                    Dim TOTAL_COST As Decimal = Val(rows.SummaryValues("TOTAL_COST").Value & "")

                    CONSUMPTION = 0
                    TOTAL_COST = 0
                    For Each row As UltraWinGrid.UltraGridRow In rows
                        CONSUMPTION += Val(row.Cells("CONSUMPTION").Value & "")
                        TOTAL_COST += Val(row.Cells("TOTAL_COST").Value & "")
                    Next

                    Dim COST_PER_YD As Decimal = 0
                    If CONSUMPTION <> 0 Then
                        COST_PER_YD = TOTAL_COST / CONSUMPTION
                    End If
                    Return COST_PER_YD
                End If
        End Select
    End Function
End Class