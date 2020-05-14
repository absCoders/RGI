Public Class WHFBACK1

    Dim connection_is_down As Boolean = False
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Get_PARM("POTPARM1")

        With dst
            
        End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Backup"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Backup"
                Backup_Tables()


        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("View").Settings.Enabled = not_iScreenMode
                End With
            End With
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        'For Each TABLE_NAME As String In New String() _
        '        {""}
        'Next
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub


    Sub Backup_Tables()
        Dim BACK_TBL() As String
        Dim W As Integer

        Dim BACK_BATCH_NO As String = ASCMAIN1.Next_Control_No("BACK_BATCH_NO")
        Dim BACK_OPER As String = ASCMAIN1.USER_ID
        Dim BACK_DATE As Date = DATETIME_STAMP
        ReDim BACK_TBL(13)

        BACK_TBL(1) = "ICTPHYB1"
        BACK_TBL(2) = "ICTPHYC1"
        BACK_TBL(3) = "ICTPHYC2"
        BACK_TBL(4) = "ICTSTAT1"
        BACK_TBL(5) = "ICTSTAT2"
        BACK_TBL(6) = "ICTSTAT5"
        BACK_TBL(7) = "WHTLOCP1"



        For W = 1 To 7
            ASCMAIN1.Progress("Now Backing Up Table " & BACK_TBL(W), "")
            ProgressBar1.Value = (W / 7) * 100
            Application.DoEvents()

            ASCMAIN1.sql = " INSERT INTO BA" & Mid(BACK_TBL(W), 3, 6) _
            & " SELECT '" & BACK_BATCH_NO.PadLeft(10, "0") & "' BACK_BATCH_NO, " _
            & " '" & BACK_OPER & "' BACK_OPER, " _
            & " SYSDATE BACK_DATE, " _
            & " " & BACK_TBL(W) & ".* " _
            & " FROM " & BACK_TBL(W)
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        Next W

        txtBACK_BATCH_NO.Text = BACK_BATCH_NO
        txtBACK_OPER.Text = BACK_OPER
        txtBACK_DATE.text = Format(BACK_DATE, "MM/dd/yyyy, HH:mm:ss")

        GroupBox2.Visible = True
        MsgBox("Back-Up is Complete.  Please Note Your Back-up Batch Data.", MsgBoxStyle.OkOnly, "Complete")
    End Sub
  
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name


            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "PO Inquiry"
               
        End Select
    End Sub


#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

End Class