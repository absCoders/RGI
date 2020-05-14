Public Class ICFUPCX1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
 
        With dst
            ASCMAIN1.sql = "SELECT ICTSTYC1.UPC_CODE, ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE " _
                & ", ICTSTYL1.STYLE_DESC" _
                & " from ICTSTYC1, ICTSTYL1" _
                & " where ICTSTYC1.UPC_CODE IS NOT NULL" _
                & "   and ICTSTYL1.STYLE_CODE = ICTSTYC1.STYLE_CODE"
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select UPC_CODE from ICTSTYC1"
            Create_TDA(.Tables.Add, "ICTUPCX1", "**", 0, False, "", 0)
         End With

        grdICTSTYC1.DataSource = dst.Tables("ICTSTYC1")
        grdICTUPCX1.DataSource = dst.Tables("ICTUPCX1")

        Create_Summary(grdICTSTYC1, "UPC_CODE", "Count")

        Create_Summary(grdICTUPCX1, "UPC_CODE", "Count")

        spl.Panel1Collapsed = True
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                'Dim rowEDT846T1 As DataRow = LookUp("EDT846T1", EDI_DOC_SEQ_NO)
                'If rowEDT846T1 Is Nothing Then
                '    Exit Sub
                'End If


                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("SEASON_CODE").Text) Then
                '        Exit Sub
                '    End If
                'End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Visible = False
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
 
        grdICTUPCX1.Visible = True

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"ICTSTYC1", "ICTUPCX1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        Load_ICTSTYC1()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            'rowICTPLIN1.Item("STMT_TYPE") = HFs("STMT_TYPE")
            'rowICTPLIN1.Item("STMT_DESC") = HFs("STMT_DESC")
        Else

        End If
 
        If EntryMode = "N" Then
        Else

        End If

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "SEASON_CODE"
            '    If Absx1.optFor("STMT_TYPE").CheckedIndex <> -1 Then
            '        sql_where = "STMT_TYPE = '" & Absx1.optFor("STMT_TYPE").Value & "'"
            '    End If
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(tvw, "BBB", "Insert Above", "Insert Below", "Insert Within")
        Load_Popup_Menu(grdICTSTYC1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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

                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Style Multi-Color"
            '    Using F As New TAC.ICFSTYCX
            '        F.STYLE_CODE = ""
            '        F.Price_Caption = "Cost" & IIf(ssdDZGRD.Value = 1, "", "/Dz")
            '        F.ShowDialog()
            '        If F.STYLE_CODE <> "" Then
            '            Add_Colors(F.STYLE_CODE, F.dst.Tables("ICTCOLRM"), F.PRICE)
            '        End If
            '    End Using

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        'Select Case Absx1.GetABSColumnName(sender)
        '    Case "LP_CODE"
        '        If e.KeyCode = Windows.Forms.Keys.Enter Then
        '            Me.UltraGroupBox1.Select() ' to force txt_Leave event to fire, for formatting
        '            Load_ICTSTYC1()
        '        End If
        'End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        'Select Case Absx1.GetABSColumnName(txtctl)
        '    Case "LP_CODE"
        '        Load_ICTSTYC1()
        'End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        'With Absx1.txtFor(COLUMN_NAME)
        '    Select Case COLUMN_NAME

        '        Case "LP_CODE"
        '            Load_ICTSTYC1()

        '    End Select

        'End With
    End Sub

#End Region

    Sub Load_ICTSTYC1()
 
        Fill_Records("ICTSTYC1")
        Sort_grdColumns(grdICTSTYC1, "UPC_CODE".ToLower)
    End Sub

    Private Sub cmdGenerateUPCs_Click(sender As System.Object, e As System.EventArgs) Handles cmdGenerateUPCs.Click
        If numUPCs.Value < 1 Or numUPCs.Value > 50 Then
            MsgBox("Number of UPCs must be between 1 and 50", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        For i As Integer = 1 To numUPCs.Value
            dst.Tables("ICTUPCX1").Rows.Add(New String() {Get_UPC_Code()})
        Next

    End Sub

    Function Get_UPC_Code() As String
        Dim UPC_CODE As String = ""
        Do
            Dim UPC_CODE_CTL_NO As String = ""
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("UPC_CODE")
            Else
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("ICTUPCH1.UPC_CODE")
            End If

            UPC_CODE = TAC.SOCMAIN1.UPC(Me, UPC_CODE_CTL_NO, ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"), True)
            If LookUp("ICTUPCH1", UPC_CODE) Is Nothing Then Exit Do
        Loop

        ASCMAIN1.sql = "Insert into ICTUPCH1 (UPC_CODE,STYLE_CODE,COLOR_CODE,INIT_DATE,INIT_OPER) " & vbCrLf _
            & " values (:PARM1,:PARM2,:PARM3,SYSDATE,:PARM4)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {UPC_CODE, "", "", ASCMAIN1.USER_ID})


        Return UPC_CODE
    End Function
End Class