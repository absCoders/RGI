Public Class TATDASH1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

         With dst
            Create_TDA(.Tables.Add, "TATDASH2", "*", 1)
        End With

        grdTATDASH2.DataSource = dst.Tables("TATDASH2")

        'ASCMAIN1.Add_Value_List(grdTATDASH2, "CUST_COMMENT_KEY")

        'Call InitializeControls(Me)
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
        Update_Record_TDA("TATDASH2")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

        EnforceConstraints(False)
        Call Fill_Records("TATDASH2", New String() {Absx1.txtFor("DASH_CODE").Text})
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("TATDASH2").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdTATDASH2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        If Not tf Then
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(0)
        End If
        'Set_Read_Only(UltraTabControl1, Not tf)
        UltraTabControl1.Enabled = tf
    End Sub

#End Region

End Class