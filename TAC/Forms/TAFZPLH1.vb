Imports Infragistics.Win.UltraWinGrid

Public Class TAFZPLH1

    Dim Appearance_Magenta As New Infragistics.Win.Appearance
    Dim Appearance_Yellow As New Infragistics.Win.Appearance
    Dim Appearance_Empty As New Infragistics.Win.Appearance

    Dim hdgKEY1 As String = ""
    Dim HDGLNO1 As String = ""
    Dim hdgKEY2 As String = ""
    Dim HDGLNO2 As String = ""

    Public Sub New(ByVal FF As ASFBASE1,
                   Optional hdgKEY1 As String = "", Optional HDGLNO1 As String = "",
                   Optional hdgKEY2 As String = "", Optional HDGLNO2 As String = "")

        Me.hdgKEY1 = hdgKEY1
        Me.HDGLNO1 = HDGLNO1
        Me.hdgKEY2 = hdgKEY2
        Me.HDGLNO2 = HDGLNO2

        frmASFBASE1 = FF
        InitializeComponent()
    End Sub


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Appearance_Magenta.BackColor = Drawing.Color.Magenta
        Appearance_Yellow.BackColor = Drawing.Color.Yellow

        With dst
            ASCMAIN1.sql = $"Select TATZPLH1.* from  TATZPLH1"
            Create_TDA(.Tables.Add, "TATZPLH1", "**", 0, False)
            With .Tables("TATZPLH1")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

            Fill_Records("TATZPLH1")

        End With

        grdTATZPLH1.DataSource = dst.Tables("TATZPLH1")

        With grdTATZPLH1.DisplayLayout.Bands(0)
            .Columns("KEY1").Header.Caption = hdgKEY1
            .Columns("LNO1").Header.Caption = HDGLNO1
            .Columns("KEY2").Header.Caption = hdgKEY2
            .Columns("LNO2").Header.Caption = HDGLNO2
        End With


        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdTATZPLH1}
            grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

            With grd.DisplayLayout.Bands(0)
                For Each c As UltraWinGrid.UltraGridColumn In .Columns

                    'If New String() {"KEY1", "LNO1", "KEY2", "LNO2"}.Contains(c.Key) Then
                    '    If c.Header.Caption = "" Then c.Hidden = True
                    'End If

                    c.Header.Appearance.BackColor = System.Drawing.Color.White
                    c.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    c.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                    If New String() {"KEY1", "LNO1", "KEY2", "LNO2"}.Contains(c.Key) Then
                        If c.Header.Caption = "" Then c.Hidden = True
                        c.Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                    End If
                Next
            End With
        Next

        ASCMAIN1.sql = "" _
            & "Select LABEL_MINI_IP_ADDRESS PRINTER_IP, STATION_ID || ' (mini)' DESCRIPTION" & vbCrLf _
            & " from WHTLINE1 WHERE LABEL_MINI_IP_ADDRESS Is Not NULL" & vbCrLf _
            & " UNION" & vbCrLf _
            & "Select LABEL_IP_ADDRESS PRINTER_IP, STATION_ID || ' (4x6)'" & vbCrLf _
            & " from WHTLINE1 WHERE LABEL_IP_ADDRESS IS NOT NULL"
        cmbPrinters.DataSource = ASCDATA1.GetDataTable

        Create_Summary(grdTATZPLH1, "ZPL_CTL_NO", "Count")

    End Sub

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdTATZPLH1, "SBB", "Show Filter", "Select All", "De-Select All")
    End Sub

    Private Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles tlb.BeforeToolDropdown

        Dim grd As UltraWinGrid.UltraGrid = Nothing

        If e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        grd = GRDs(Mid(e.SourceControl.Name, 4))
        If grd.ActiveRow Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

            Case "grdSOTORDQ0"

                'tlb_btn = DirectCast(tlb_pop.Tools("Release Selected Orders"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (MENU_ITEM_OBJECT = "SOFPICKF")

        End Select

    End Sub

    Private Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles tlb.ToolClick

        Dim grd As UltraWinGrid.UltraGrid = Nothing

        grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select

    End Sub

#End Region

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrint.Click

        Me.Close()
    End Sub

    Private Sub grdTATZPLH1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdTATZPLH1.DoubleClickRow
        Dim PRINTER_IP As String = e.Row.Cells("PRINTER_IP").Value
        Dim ZPL_BODY As String = e.Row.Cells("ZPL_BODY").Value
        Dim zplPrint As New TAC.TACZPLT1()
        zplPrint.SendLabelToPrinter(PRINTER_IP, ZPL_BODY)

        'Using zplPrint As New TACZPLT1()
        '    zplPrint.SendLabelToPrinter(PRINTER_IP, ZPL_BODY)
        'End Using
    End Sub

    Private Sub optPrinter_ValueChanged(sender As Object, e As EventArgs) Handles optPrinter.ValueChanged
        If optPrinter.Value = "SEL" Then
            cmbPrinters.Enabled = True
        Else
            cmbPrinters.Enabled = False
        End If
    End Sub
End Class