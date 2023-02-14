Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Drawing

Namespace MeasureStringSample
    Public Class CFMeasureString
        Private Structure Rect
            Public Left As Integer, Top As Integer, Right As Integer, Bottom As Integer
            Public Sub New(ByVal r As Rectangle)
                Me.Left = r.Left
                Me.Top = r.Top
                Me.Bottom = r.Bottom
                Me.Right = r.Right
            End Sub
        End Structure

        <DllImport("coredll.dll")> _
        Private Shared Function DrawText(ByVal hdc As IntPtr, ByVal lpStr As String, ByVal nCount As Integer, ByRef lpRect As Rect, ByVal wFormat As Integer) As Integer
        End Function
        Private Const DT_CALCRECT As Integer = &H400
        Private Const DT_WORDBREAK As Integer = &H10
        Private Const DT_EDITCONTROL As Integer = &H2000

        Public Shared Function MeasureString(ByVal gr As Graphics, ByVal text As String, ByVal rect As Rectangle, ByVal textboxControl As Boolean) As Size
            Dim bounds As New Rect(rect)
            Dim hdc As IntPtr = gr.GetHdc()
            Dim flags As Integer = DT_CALCRECT Or DT_WORDBREAK
            If textboxControl Then
                flags = flags Or DT_EDITCONTROL
            End If
            DrawText(hdc, text, text.Length, bounds, flags)
            gr.ReleaseHdc(hdc)
            Return New Size(bounds.Right - bounds.Left, bounds.Bottom - bounds.Top + (If(textboxControl, 6, 0)))
        End Function
    End Class
End Namespace