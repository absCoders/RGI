Public Class WHFPKGM1

    Public PKG_BOX_UPC As String = String.Empty
    Private PKG_CODE As String = String.Empty
    Public Updated As Boolean = False

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "WHTPKGM1", "*")
        End With

        PKG_CODE = ASCMAIN1.Next_Control_No("WHTPKGM1.PKG_CODE")
        PKG_CODE = "P" & StrReverse(StrReverse(PKG_CODE).Substring(0, 5))
        Fill_Records("WHTPKGM1", PKG_CODE)

        ' Prevent unique key if someone manually enters a box
        While dst.Tables("WHTPKGM1").Rows.Count <> 0
            PKG_CODE = ASCMAIN1.Next_Control_No("WHTPKGM1.PKG_CODE")
            PKG_CODE = "P" & StrReverse(StrReverse(PKG_CODE).Substring(0, 5))
            Fill_Records("WHTPKGM1", PKG_CODE)
        End While

        dst.Tables("WHTPKGM1").Rows.Add(New Object() {PKG_CODE})
        dst.Tables("WHTPKGM1").Rows(0).Item("PKG_BOX_UPC") = PKG_BOX_UPC

        Updated = False
    End Sub


    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Updated = False
        Me.Close()
    End Sub


    Private Sub cmdCCSubmit_Click(sender As Object, e As EventArgs) Handles cmdCCSubmit.Click

        EMsg = String.Empty

        Dim PKG_CODE As String = Absx1.txtFor("PKG_CODE").Text.Trim
        Dim PKG_DESC As String = Absx1.txtFor("PKG_DESC").Text.Trim
        Dim PKG_BOX_UPC As String = Absx1.txtFor("PKG_BOX_UPC").Text.Trim

        Dim PKG_L As String = Val(Absx1.numFor("PKG_L").Value & String.Empty)
        Dim PKG_W As String = Val(Absx1.numFor("PKG_W").Value & String.Empty)
        Dim PKG_H As String = Val(Absx1.numFor("PKG_H").Value & String.Empty)
        Dim PKG_WT As String = Val(Absx1.numFor("PKG_WT").Value & String.Empty)
        Dim PKG_CUBE As String = Val(Absx1.numFor("PKG_CUBE").Value & String.Empty)
        Dim PKG_COST As String = Val(Absx1.numFor("PKG_COST").Value & String.Empty)
        Dim PKG_CHARGE As String = Val(Absx1.numFor("PKG_CHARGE").Value & String.Empty)

        If PKG_CODE.Length = 0 Then
            EMsg &= vbCr & "Package Code is required."
        End If

        If PKG_BOX_UPC.Length = 0 Then
            EMsg &= vbCr & "Package barcode is required."
        End If

        If PKG_L <= 0 Then
            EMsg &= vbCr & "Package length must be greater than 0."
        End If

        If PKG_W <= 0 Then
            EMsg &= vbCr & "Package width must be greater than 0."
        End If

        If PKG_H <= 0 Then
            EMsg &= vbCr & "Package heigt must be greater than 0."
        End If

        If PKG_WT < 0 Then
            EMsg &= vbCr & "Package weight must be greater/equal 0."
        End If

        If PKG_CUBE < 0 Then
            EMsg &= vbCr & "Package cube must be greater/equal 0."
        End If

        If PKG_COST < 0 Then
            EMsg &= vbCr & "Package cost must be greater/equal 0."
        End If

        If PKG_CHARGE < 0 Then
            EMsg &= vbCr & "Package charge must be greater/equal 0."
        End If

        ' Sort the values by length, width, height
        If PKG_L <= 0 OrElse PKG_W <= 0 OrElse PKG_H < 0 Then
            EMsg &= vbCr & "All package dimensions must be greater than 0"
        End If

        If EMsg.Length > 0 Then
            MessageBox.Show(EMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Absx1.numFor("PKG_CHARGE").Value = PKG_CHARGE


        Dim dimList As New List(Of Decimal)
        dimList.Add(PKG_L)
        dimList.Add(PKG_W)
        dimList.Add(PKG_H)
        dimList.Sort()

        PKG_L = dimList(2)
        PKG_W = dimList(1)
        PKG_H = dimList(0)

        Absx1.numFor("PKG_L").Value = PKG_L
        Absx1.numFor("PKG_W").Value = PKG_W
        Absx1.numFor("PKG_H").Value = PKG_H

        If PKG_DESC.Length = 0 Then
            PKG_DESC = "Box " & PKG_L & " x " & PKG_W & " x " & PKG_H
            Absx1.txtFor("PKG_DESC").Text = PKG_DESC
        End If

        Try
            Dim rowWHTPKGM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM WHTPKGM1 WHERE PKG_BOX_UPC = '" & PKG_BOX_UPC & "'")
            If rowWHTPKGM1 IsNot Nothing Then
                Updated = True
                Me.Close()
                Exit Sub
            End If
        Catch ex As Exception
            Updated = True
            Me.Close()
            Exit Sub
        End Try

        Try
            BeginTrans()
            Update_Record_TDA("WHTPKGM1")
            CommitTrans("Package updated successfully")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

        Updated = True
        Me.Close()

    End Sub
End Class