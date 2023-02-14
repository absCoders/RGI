<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class InputForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Barcode1 = New Barcode.Barcode
        Me.SuspendLayout()
        '
        'Barcode1
        '
        Me.Barcode1.DecoderParameters.CODABAR = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.CODE128 = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.CODE39 = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.CODE39Params.Code32Prefix = False
        Me.Barcode1.DecoderParameters.CODE39Params.Concatenation = False
        Me.Barcode1.DecoderParameters.CODE39Params.ConvertToCode32 = False
        Me.Barcode1.DecoderParameters.CODE39Params.FullAscii = False
        Me.Barcode1.DecoderParameters.CODE39Params.Redundancy = False
        Me.Barcode1.DecoderParameters.CODE39Params.ReportCheckDigit = False
        Me.Barcode1.DecoderParameters.CODE39Params.VerifyCheckDigit = False
        Me.Barcode1.DecoderParameters.D2OF5 = Barcode.DisabledEnabled.Disabled
        Me.Barcode1.DecoderParameters.EAN13 = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.EAN8 = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.I2OF5 = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.KOREAN_3OF5 = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.MSI = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.UPCA = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DecoderParameters.UPCE0 = Barcode.DisabledEnabled.Enabled
        Me.Barcode1.DeviceName = Nothing
        Me.Barcode1.EnableScanner = False
        Me.Barcode1.ScanParameters.BeepFrequency = 2670
        Me.Barcode1.ScanParameters.BeepTime = 200
        Me.Barcode1.ScanParameters.CodeIdType = Barcode.CodeIdTypes.None
        Me.Barcode1.ScanParameters.LedTime = 3000
        Me.Barcode1.ScanParameters.ScanType = Barcode.ScanTypes.Foreground
        Me.Barcode1.ScanParameters.WaveFile = ""
        '
        'InputForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(638, 455)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(50, 50)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "InputForm"
        Me.Text = "Input"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Barcode1 As Barcode.Barcode
End Class
