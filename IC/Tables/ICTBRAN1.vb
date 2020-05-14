Public Class ICTBRAN1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"

            Case "Edit"
            Case "Update"

                If Absx1.optFor("BRAND_STATUS").Value & "" = "" Then
                    EMsg &= vbCr & "Status is Mandatory"
                End If
                If Absx1.txtFor("BRAND_NAME").Text & "" = "" Then
                    EMsg &= vbCr & "Brand Name is Mandatory"
                End If
 
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        'Dim sqlDelete = "CUST_CODE = '" & CUST_CODE & "'"
        'Update_Record_TDA("SPTDCOM2", sqlDelete)
    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        
    End Sub

    Overrides Sub Show_Record_Special()

        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text

        Dim FOLDER_NAME As String = ASCMAIN1.Folders("Images") & "\COLUMN_NAME\BRAND_CODE\"
        Dim IMAGE_NAME As String = BRAND_CODE & ".png"
        If My.Computer.FileSystem.FileExists(FOLDER_NAME & IMAGE_NAME) Then
            img.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , ) ' imgba)
            img.Visible = True
        End If

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
 
        If Not tf Then
            img.Visible = False
        End If

    End Sub


    Private Sub btnPickColor_Click(sender As Object, e As EventArgs)
        'Dim cDialog As New ColorDialog()
        'cDialog.Color = lblBRAND_COLOR.Appearance.ForeColor  ' initial selection is current color.
        'If (cDialog.ShowDialog() = DialogResult.OK) Then
        '    numBRAND_COLOR.Value = cDialog.Color.ToArgb
        'End If
    End Sub

    Private Sub numBRAND_COLOR_ValueChanged(sender As Object, e As EventArgs)
        'Dim rgb As Int64 = Val(numBRAND_COLOR.Value & "")
        'Dim c As System.Drawing.Color = System.Drawing.Color.FromArgb(rgb)
        '' lblBRAND_COLOR.Appearance.ForeColor = c
        'numBRAND_COLOR.Appearance.ForeColor = c
        'numBRAND_COLOR.Appearance.ForeColorDisabled = c
    End Sub
#End Region
End Class