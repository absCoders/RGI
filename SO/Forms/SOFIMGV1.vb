
Public Class SOFIMGV1
    Public mode As String = "" ' N = New, "" = Update Next Step
    Public ImageLocation As String
    Private FF As ASFBASE1

    Public Sub New(ByVal FF As ASFBASE1, ByVal Image As String)
        ImageLocation = Image
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        imgSTYLE.ImageLocation = ImageLocation
    End Sub

    Private Sub cmdDone_Click(sender As System.Object, e As System.EventArgs) Handles cmdDone.Click
        Me.Close()
    End Sub

End Class