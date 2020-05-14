Public Class SOTSDIV1

    Private Sub SOTSDIV1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        grpVAN.Visible = (ASCMAIN1.CLIENT = "VAN")
        grpSegments.Visible = (ASCMAIN1.CLIENT = "NYA")
        Get_PARM("GLTPARM1")
        GL_Segments(grpSegments, ROWs("GLTPARM1"))

        With dst
            Create_TDA(.Tables.Add, "GLTSEGM1", "*")
        End With
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        If SELECTION_NO = 0 Then Exit Sub
        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                If ASCMAIN1.CLIENT = "NYA" Then
                    Dim SEG3_CODE As String = Absx1.txtFor("SEG3_CODE").Text
                    If SEG3_CODE <> "" And SEG3_CODE <> Absx1.txtFor("SALES_DIVISION_CODE").Text Then
                        If LookUp("SOTSDIV1", SEG3_CODE) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Segment 3"
                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        If ASCMAIN1.CLIENT = "NYA" Then
            Dim SEG3_CODE As String = Absx1.txtFor("SEG3_CODE").Text
            If SEG3_CODE = "" Or SEG3_CODE = Absx1.txtFor("SALES_DIVISION_CODE").Text Then
                SEG3_CODE = Absx1.txtFor("SALES_DIVISION_CODE").Text

                Dim rowGLTSEGM1 As DataRow = Fill_Record("GLTSEGM1", New String() {"3", SEG3_CODE})
                If rowGLTSEGM1 Is Nothing Then
                    rowGLTSEGM1 = dst.Tables("GLTSEGM1").NewRow
                    rowGLTSEGM1.Item("ACCT_SEG_ID") = "3"
                    rowGLTSEGM1.Item("ACCT_SEG_CODE") = SEG3_CODE
                    rowGLTSEGM1.Item("ACCT_SEG_STATUS") = "A"
                    rowGLTSEGM1.Item("ACCT_SEG_NO_GL") = "0"
                    rowGLTSEGM1.Item("ACCT_SEG_CLASS") = "" ' Absx1.txtFor("SALES_DIVISION_CODE").Text
                    dst.Tables("GLTSEGM1").Rows.Add(rowGLTSEGM1)
                End If

                rowGLTSEGM1.Item("ACCT_SEG_DESC") = Absx1.txtFor("SALES_DIVISION_NAME").Text
                Update_Record_TDA("GLTSEGM1")
            End If
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

        If ASCMAIN1.CLIENT = "NYA" Then
            Set_Read_Only_for_ctl(Absx1.txtFor("SEG4_CODE"), Not (EntryMode = "New"))
        End If

    End Sub

#End Region

End Class