Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

Public Class Form1
    Public SourceImagePath As String
    Public DestinationPath As String
    ' Convert to PNG while we're at it? JPEG is a lossy image format
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim original As Image = Image.FromFile(SourceImagePath)
        Dim resized As Image = ResizeImage(original, New Size(1024, 768))

        SaveImageWithQuality(resized, DestinationPath, 75L)
    End Sub
End Class

Public Module ImageUtils
    Public Function ResizeImage(ByVal image As Image,
  ByVal size As Size, Optional ByVal preserveAspectRatio As Boolean = True) As Image

        Dim newWidth As Integer
        Dim newHeight As Integer
        If preserveAspectRatio Then
            Dim originalWidth As Integer = image.Width
            Dim originalHeight As Integer = image.Height
            Dim percentWidth As Single = CSng(size.Width) / CSng(originalWidth)
            Dim percentHeight As Single = CSng(size.Height) / CSng(originalHeight)
            Dim percent As Single = If(percentHeight < percentWidth, percentHeight, percentWidth)
            newWidth = CInt(originalWidth * percent)
            newHeight = CInt(originalHeight * percent)
        Else
            newWidth = size.Width
            newHeight = size.Height
        End If

        Dim newImage As Image = New Bitmap(newWidth, newHeight)

        Using graphicsHandle As Graphics = Graphics.FromImage(newImage)
            graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic
            graphicsHandle.DrawImage(image, 0, 0, newWidth, newHeight)
        End Using

        Return newImage

    End Function
    ' Compression
    Public Sub SaveImageWithQuality(ByVal bmp1 As Image, ByVal destinationPath As String, ByVal quality As Long)
        'Or you can use build-in method
        'Dim jgpEncoder As ImageCodecInfo = GetEncoderInfo("image/jpeg");
        Dim jgpEncoder As ImageCodecInfo = GetEncoder(ImageFormat.Jpeg)

        ' Create an Encoder object based on the GUID
        ' for the Quality parameter category.
        Dim myEncoder As System.Drawing.Imaging.Encoder = System.Drawing.Imaging.Encoder.Quality

        ' Create an EncoderParameters object.
        ' An EncoderParameters object has an array of EncoderParameter
        ' objects. In this case, there is only one
        ' EncoderParameter object in the array.
        Dim myEncoderParameters As New EncoderParameters(1)

        ' Save with 100% quality
        Dim myEncoderParameter As New EncoderParameter(myEncoder, quality)
        myEncoderParameters.Param(0) = myEncoderParameter
        bmp1.Save(destinationPath, jgpEncoder, myEncoderParameters)

    End Sub

    Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo

        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()

        Dim codec As ImageCodecInfo
        For Each codec In codecs
            If codec.FormatID = format.Guid Then
                Return codec
            End If
        Next codec
        Return Nothing

    End Function

End Module
