Imports System.Runtime.InteropServices
Imports PocketPCScan.MeasureStringSample


Public Class AutoSizeLabel
    Inherits Label

    Public Overrides Property Text() As String
        Get
            Return MyBase.Text
        End Get
        Set(ByVal value As String)
            MyBase.Text = value
            ReCalculateSize()
        End Set
    End Property

    Public Overrides Property Font() As System.Drawing.Font
        Get
            Return MyBase.Font
        End Get
        Set(ByVal value As System.Drawing.Font)
            MyBase.Font = value
            ReCalculateSize()
        End Set
    End Property

    Private Sub ReCalculateSize()
        Using control As New Control()
            Using g As Graphics = control.CreateGraphics()
                Dim size As Size = CFMeasureString.MeasureString(g, MyBase.Text, MyBase.ClientRectangle, False)
                MyBase.Height = CInt(size.Height) + 1
            End Using
        End Using
    End Sub

End Class


