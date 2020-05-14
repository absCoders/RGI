Public Class WHFUSPPI

    Private sqlUSSPI As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Get_PARM("POTPARM1")

        With dst

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            .Tables("SOTINVH1").Columns.Add("CUST_NAME", GetType(System.String))

            sqlUSSPI = "SELECT ICTSTYL1.DUTY_RATE_CODE HTS_CODE, ICTDUTY1.DUTY_RATE_DESC HTS_DESC, NVL(SOTSHIP1.SHIP_TOTAL_WGT, 0) SHP_TOTAL_WGT, " _
                & " SUM(SOTINVH2.ORDR_QTY_SHIP) UNITS, SUM(NVL(ICTSTYL1.STYLE_WEIGHT, 0)) WEIGHT, SUM(ORDR_UNIT_PRICE * ORDR_QTY_SHIP) VALUE_USD" _
                & " FROM SOTINVH2, ICTSTYL1, ICTDUTY1, SOTPICK1, SOTSHIP1" _
                & " WHERE SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                & " AND ICTSTYL1.DUTY_RATE_CODE = ICTDUTY1.DUTY_RATE_CODE (+)" _
                & " AND SOTINVH2.INV_NO IN ('XX')  AND SOTINVH2.ORDR_QTY_SHIP > 0" _
                & " AND SOTINVH2.INV_NO = SOTPICK1.INV_NO (+)" _
                & " AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" _
                & " GROUP BY ICTSTYL1.DUTY_RATE_CODE, ICTDUTY1.DUTY_RATE_DESC, NVL(SOTSHIP1.SHIP_TOTAL_WGT, 0)"
            Create_TDA(.Tables.Add, "WHTUSPPI", sqlUSSPI, 0, False, "", 0)

            .Tables("WHTUSPPI").Columns.Add("WEIGHT_KG", GetType(System.Decimal), "WEIGHT * 0.453592")

        End With

        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")
        grdWHTUSPPI.DataSource = dst.Tables("WHTUSPPI")

        Create_Summary(grdSOTINVH1, "INV_NO", "Count")
        Create_Summary(grdWHTUSPPI, "HTS_CODE", "Count")
        Create_Summary(grdWHTUSPPI, "UNITS", "Sum")
        Create_Summary(grdWHTUSPPI, "WEIGHT", "Sum")
        Create_Summary(grdWHTUSPPI, "VALUE_USD", "Sum")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"


            Case "Cancel"
                If MessageBox.Show("Do you want to clear the contents on the screen?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Load"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                End With
            End With
        End If

        txtINV_NO.Clear()

        If ScreenMode Then
            grdWHTUSPPI.Visible = True
            grpINV_NO.Enabled = False
        Else
            Clear_Record()
            grdWHTUSPPI.Visible = False
            grpINV_NO.Enabled = True

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SOTINVH1", "WHTUSPPI"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Dim lstInvNos As New List(Of String)
        For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
            lstInvNos.Add(row.Item("INV_NO"))
        Next

        ASCMAIN1.sql = sqlUSSPI
        ASCMAIN1.sql = ASCMAIN1.sql.Replace("'XX'", "'" & String.Join("', '", lstInvNos.ToArray) & "'")
        Fill_Records("WHTUSPPI", String.Empty, True, ASCMAIN1.sql)

        grdWHTUSPPI.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        EnforceConstraints(True)

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTUSPPI, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name


            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case ""

        End Select
    End Sub


#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

#Region "From Controls"

    Private Sub txtINV_NO_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtINV_NO.KeyPress

        If Not e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Exit Sub
        End If

        Try
            e.Handled = True
            Dim INV_NO As String = txtINV_NO.Text

            INV_NO = INV_NO.Trim
            If INV_NO.Length = 0 Then
                Exit Sub
            End If

            INV_NO = ASCMAIN1.Format_Field(INV_NO, "INV_NO")

            If dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'").Length = 1 Then
                MessageBox.Show("The provided Invoice Number (" & INV_NO & ") is already loaded.")
                Exit Sub
            End If

            ASCMAIN1.sql = "Select SOTINVH1.*, ARTCUST1.CUST_NAME FROM SOTINVH1, ARTCUST1 WHERE SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE AND SOTINVH1.INV_TYPE = 'I' AND SOTINVH1.INV_NO = '" & INV_NO & "'"
            Fill_Records("SOTINVH1", String.Empty, False, ASCMAIN1.sql)

            If dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'").Length = 0 Then
                MessageBox.Show("The provided Invoice Number (" & INV_NO & ") could not be found.")
            End If


        Catch ex As Exception
            MessageBox.Show("Error acessing Invoice: " & ex.Message, "Get Invoice", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Finally
            txtINV_NO.Clear()
            txtINV_NO.Focus()
        End Try

    End Sub

#End Region

End Class