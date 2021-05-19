Public Class WHTPKGM1

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Update"

                ' Sort the values by length, width, height
                Dim PKG_L As Decimal = Absx1.numFor("PKG_L").Value
                Dim PKG_W As Decimal = Absx1.numFor("PKG_W").Value
                Dim PKG_H As Decimal = Absx1.numFor("PKG_H").Value

                If PKG_L <= 0 OrElse PKG_W <= 0 OrElse PKG_H < 0 Then
                    EMsg &= vbCr & "All dimensions must be greater than 0"
                    Exit Sub
                End If

                Dim dimList As New List(Of Decimal)
                dimList.Add(PKG_L)
                dimList.Add(PKG_W)
                dimList.Add(PKG_H)
                dimList.Sort()
                PKG_L = dimList(2)
                PKG_W = dimList(1)
                PKG_H = dimList(0)

                Dim rowWHTPKGM1 As DataRow = dst.Tables("WHTPKGM1").Rows.Find(Absx1.txtFor("PKG_CODE").Text)

                If rowWHTPKGM1 IsNot Nothing Then
                    rowWHTPKGM1.Item("PKG_L") = PKG_L
                    rowWHTPKGM1.Item("PKG_W") = PKG_W
                    rowWHTPKGM1.Item("PKG_H") = PKG_H
                End If

        End Select
    End Sub

    Private Sub UltraLabel14_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class