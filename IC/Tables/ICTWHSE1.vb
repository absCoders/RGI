Public Class ICTWHSE1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblWHSE_EMAIL_VAS.Visible = (ASCMAIN1.CLIENT = "NYA")
        txtWHSE_EMAIL_VAS.Visible = (ASCMAIN1.CLIENT = "NYA")

        grpSegments.Visible = (ASCMAIN1.CLIENT = "NYA")
        Get_PARM("GLTPARM1")
        GL_Segments(grpSegments, ROWs("GLTPARM1"))

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"

                If Absx1.txtFor("LP_CODE").Text <> "" And Absx1.chkFor("WHSE_LOCATOR").Checked Then
                    EMsg &= vbCr & "A 3PL Warehouse Cannot be Set Up with Locator Support"
                End If

                If Absx1.chkFor("WHSE_LOCATOR").Checked Then
                    If Absx1.optFor("WHSE_CTN_CTL").Value & "" = "" Then
                        EMsg &= vbCr & "A Carton Control method must be chosen if the warehouse has Location Control"
                    End If
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'If rowASFBASE1.Item("WHSE_LOCATOR") & "" = "1" Then
        '    Dim WHSE_CODE As String = rowASFBASE1.Item("WHSE_CODE")
        '    ASCMAIN1.sql = "Insert into WHTLOCM1 (WHSE_CODE, LOCATION_CODE, LOCATION_DESC)" & vbCrLf _
        '        & "Select '" & WHSE_CODE & "' WHSE_CODE, LOCATION_CODE, LOCATION_DESC from WHTLOCM0" & vbCrLf _
        '        & " where LOCATION_CODE in " & vbCrLf _
        '        & "(Select LOCATION_CODE from WHTLOCM0 minus " & vbCrLf _
        '        & " Select LOCATION_CODE from WHTLOCM1 where WHSE_CODE = '" & WHSE_CODE & "')"
        '    ASCDATA1.ExecuteSQL()
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        'Dim WHSE_LOCATOR As String = rowASFBASE1.Item("WHSE_LOCATOR") & ""
        grpLOCATIONs.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked) Or (EntryMode = "New")
        grpReturns.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked) Or (EntryMode = "New")
        grpVirtual.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked) Or (EntryMode = "New")
        lblWHSE_CTN_CTL.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked) Or (EntryMode = "New")
        optWHSE_CTN_CTL.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked) Or (EntryMode = "New")
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        Set_Read_Only_for_ctl(Absx1.chkFor("WHSE_LOCATOR"), Not (EntryMode = "New"))
        Set_Read_Only_for_ctl(Absx1.optFor("WHSE_CTN_CTL"), Not (EntryMode = "New"))

        If ASCMAIN1.CLIENT = "NYA" Then
            Set_Read_Only_for_ctl(Absx1.txtFor("SEG4_CODE"), Not (EntryMode = "New"))
        End If

    End Sub
#End Region

End Class