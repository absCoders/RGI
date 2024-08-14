Public Class ASFJSONG
    Public user_option As Integer = 0
    Public grow As UltraWinGrid.UltraGridRow

    Public grdFilter As Boolean = False
    Public grdGroupBy As Boolean = False

    Private f_Calling_Form As ASFBASE0

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        'If optmsg.Visible Then
        '    user_option = optmsg.CheckedIndex
        '    If user_option = -1 Then
        '        Exit Sub
        '    End If
        'End If
        'If grdmsg.Visible Then
        '    If grdmsg.ActiveRow IsNot Nothing Then
        '        grow = grdmsg.ActiveRow
        '    End If
        'End If
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        user_option = -1
        'txtmsg.Text = ""
        Me.Close()
    End Sub


    Public Sub Show_Formatted_txt(
    ByVal Form_Caption As String,
    ByVal Formatted_Text As String,
    ByVal ff As Form)

        'f_Calling_Form = ff

        Me.Text = Form_Caption
        Me.Height = 0.6 * ff.Height
        Me.Width = 0.6 * ff.Width

        'fmttxt.Value = Formated_Text
        'fmttxt.ReadOnly = True
        'fmttxt.Appearance.BackColor = Color.White ' Color.Empty ' Color.White
        'fmttxt.Visible = True

        'WebBrowser1.DocumentText = Formatted_Text
        'WebBrowser1.Visible = True

        cmdOK.Visible = False
        cmdCancel.Text = "Done"
        'cmdPrint.Visible = True

        'fmttxt.Appearance.BackColorDisabled = Color.Empty ' Color.White
        'fmttxt.Appearance.ForeColorDisabled = Color.Empty ' Color.DodgerBlue
        'fmttxt.
        lblmsg.Visible = False


        Me.ShowDialog()
    End Sub

    Private Sub ASFJSONG_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ASCMAIN1.Center(Me)
        'grdmsg.DisplayLayout.Override.RowSelectors = DefaultableBoolean.Default
        'grdmsg.DisplayLayout.Override.RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
        'ASCMAIN1.grdInitializeLayout(grdmsg)

        'txtmsg.SpellChecker = ASFMAIN1.UltraSpellChecker1
        'ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(txtmsg, "txtMenu")


        ' fmttxt.Visible = False
        'grdmsg.Parent = SplitContainer1.Panel1
        'grdmsg.Dock = DockStyle.Fill
        'SplitContainer1.FixedPanel = FixedPanel.Panel2
        'If img1.Image Is Nothing Then
        '    img1.Dock = DockStyle.None
        '    img1.Visible = False
        'Else
        '    img1.Dock = DockStyle.Fill
        '    img1.Visible = True
        '    cmdPrint.Visible = True
        '    Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        'End If
    End Sub

    Private Sub cmdPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim PrintDialog1 As New System.Windows.Forms.PrintDialog

        Dim result As DialogResult = PrintDialog1.ShowDialog()

        If result = Windows.Forms.DialogResult.OK Then
            'If img1 Is Nothing Then
            '    WebBrowser1.Print()
            'Else
            Try
                    SplitContainer1.Panel2Collapsed = True

                'img = CaptureForm1()
                'pd = New System.Drawing.Printing.PrintDocument

                'pd.Print()

                Dim ppDialog As PrintPreviewDialog = New PrintPreviewDialog()
                    ppDialog.ClientSize = New Size(400, 500)
                'ppDialog.Document = pd
                'pd.Print()
                ''  ppDialog.ShowDialog()

                'pd = Nothing
                'img = Nothing

                SplitContainer1.Panel2Collapsed = False
                Catch ex As Exception
                    MsgBox("Please verify that the Default Printer is available", MsgBoxStyle.OkOnly, "Cannot Select Default Printer")
                End Try
            'End If

        End If
    End Sub

    <System.Runtime.InteropServices.DllImport("gdi32.DLL", EntryPoint:="BitBlt",
    SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode,
    ExactSpelling:=True,
    CallingConvention:=System.Runtime.InteropServices.CallingConvention.StdCall)>
    Private Shared Function BitBlt(ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As System.Int32) As Boolean

        ' Leave function empty - DLLImport attribute forwards calls to MoveFile to
        ' MoveFileW in KERNEL32.DLL.
    End Function


End Class