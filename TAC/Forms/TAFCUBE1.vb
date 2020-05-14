Public Class TAFCUBE1

    Public calculatedCube As Double = 0

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        calculatedCube = 0
    End Sub

    Private Sub btnCalculate_Click(sender As System.Object, e As System.EventArgs) Handles btnCalculate.Click

        Try

            Dim length As Double = MyBase.Absx1.numFor("PKG_L").Value
            Dim width As Double = MyBase.Absx1.numFor("PKG_W").Value
            Dim height As Double = MyBase.Absx1.numFor("PKG_H").Value

            ' Convert Centimeters to Inches
            If optEntryType.Value = "M" Then
                length *= 0.39
                width *= 0.39
                height *= 0.39
            End If

            'If you've already measured the inches, mutiply the three numbers together to get cubic inches. 
            'Then divide by 1728, the number of cubic inches in a cubic foot
            Dim cubic As Double = (length * width * height) / 1728
            MyBase.Absx1.numFor("PKG_C").Value = cubic
            calculatedCube = cubic

            ' If they use Cubic Meters
            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                ' calculatedCube *= 0.0283168466
            End If

        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message)
            calculatedCube = 0
        End Try

    End Sub

    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click
        DialogResult = Windows.Forms.DialogResult.OK
        Hide()
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        DialogResult = Windows.Forms.DialogResult.Cancel
        Hide()
    End Sub
End Class