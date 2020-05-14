Public Class POFEXPD1

    Dim POTORDR1 As String
    Dim VEND_CODEs As New List(Of String)
    Dim PO_ORDER_NOs As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")
        Create_Temporary_Tables()

        With dst
            ASCMAIN1.sql = "SELECT POTORDR1.*, CUST_NAME from POTORDR1, ARTCUST1 where ARTCUST1.CUST_CODE (+) = POTORDR1.CUST_CODE"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "", 1)
            With .Tables("POTORDRX")
                .Columns.Add("PO_QTY_ORD", GetType(System.Int64))
                .Columns.Add("PO_QTY_SHP", GetType(System.Int64))
                .Columns.Add("PO_QTY_REC", GetType(System.Int64))
                .Columns.Add("PO_QTY_OPN", GetType(System.Int64))
                .Columns.Add("PO_AMT_ORD", GetType(System.Decimal))
                .Columns.Add("PO_AMT_SHP", GetType(System.Decimal))
                .Columns.Add("PO_AMT_REC", GetType(System.Decimal))
                .Columns.Add("PO_AMT_OPN", GetType(System.Decimal))
                .Columns.Add("SEL", GetType(System.String))
                .Columns("SEL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "Select POTORDR1.* from " & POTORDR1 & " POTORDR1"
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "", 1)
            With .Tables("POTORDR1")
                .Columns.Add("SEL", GetType(System.String))
                .Columns("SEL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "Select POTORDR2.* from POTORDR2" & vbCrLf _
                & " where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select X.VEND_CODE, APTVEND1.VEND_NAME, X.PO_COUNT, X.PO_COUNT_TO_XPED, X.PO_COUNT_TO_XMIT, X.PO_QTY_OPN, X.PO_AMT_OPN" & vbCrLf _
                & " from APTVEND1, (Select VEND_CODE, COUNT(*) PO_COUNT" & vbCrLf _
                & ", SUM(CASE WHEN PO_XMIT_IND = '1' THEN 1 ELSE 0 END) PO_COUNT_TO_XPED" & vbCrLf _
                & ", SUM(CASE WHEN PO_XMIT_IND = '0' THEN 1 ELSE 0 END) PO_COUNT_TO_XMIT" & vbCrLf _
                & ", SUM (PO_QTY_OPN) PO_QTY_OPN, SUM (PO_AMT_OPN) PO_AMT_OPN from " & POTORDR1 & " group by VEND_CODE) X" & vbCrLf _
                & " where APTVEND1.VEND_CODE (+) = X.VEND_CODE"
            Create_TDA(.Tables.Add, "POTORDRV", "**", 0, False, "", 1)
            .Tables("POTORDRV").Columns("PO_COUNT").DataType = GetType(System.Int64)
            .Tables("POTORDRV").Columns("PO_COUNT_TO_XPED").DataType = GetType(System.Int64)
            .Tables("POTORDRV").Columns("PO_COUNT_TO_XMIT").DataType = GetType(System.Int64)
            .Tables("POTORDRV").Columns("PO_QTY_OPN").DataType = GetType(System.Int64)
        End With

        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdPOTORDR1.DataSource = dst.Tables("POTORDR1")
        grdPOTORDR2.DataSource = dst.Tables("POTORDR2")
        grdPOTORDRV.DataSource = dst.Tables("POTORDRV")

        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDRX, New String() {"SEL", "PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN", _
                                                  "PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"}, , , "#,##0")

        Create_Summary(grdPOTORDR1, "PO_ORDER_NO", "Count")
        Create_Summary(grdPOTORDR1, New String() {"PO_QTY_ORD", "PO_QTY_OPN", "PO_AMT_ORD", "PO_AMT_OPN"}, , , "#,##0")

        Create_Summary(grdPOTORDR2, "PO_ORDER_LNO", "Count")
        Create_Summary(grdPOTORDR2, New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_OPN"}, , , "#,##0")

        Create_Summary(grdPOTORDRV, "VEND_CODE", "Count")
        Create_Summary(grdPOTORDRV, New String() {"PO_COUNT", "PO_COUNT_TO_XPED", "PO_COUNT_TO_XMIT", "PO_QTY_OPN", "PO_AMT_OPN"}, , , "#,##0")

        grdPOTORDRX.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        grdPOTORDRX.DisplayLayout.UseFixedHeaders = True
        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("PO_ORDER_NO").Header.Fixed = True
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VEND_NAME").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
            Next
        End With
        grdPOTORDRX.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grdPOTORDRX, True)


        With grdPOTORDRV.DisplayLayout.Bands(0)
            .Columns("VEND_CODE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"PO_COUNT", "PO_COUNT_TO_XPED", "PO_COUNT_TO_XMIT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
            Next
        End With

        With grdPOTORDR1.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("PO_ORDER_NO").Header.Fixed = True
            .Columns("PO_DATE_ORDERED").Header.Fixed = True
            .Columns("PO_DATE_SHIP_BY").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
            Next
        End With

        With grdPOTORDR2.DisplayLayout.Bands(0)
            .Columns("PO_ORDER_LNO").Header.Fixed = True
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If New String() {"PO_QTY_ORD", "PO_QTY_SHP", "PO_QTY_REC", "PO_QTY_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PO_AMT_ORD", "PO_AMT_SHP", "PO_AMT_REC", "PO_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White
            Next
        End With

        dteShipDate.Value = Now.Date

        ASCMAIN1.Add_Value_List(grdPOTORDRX, "PO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "LABEL_RESP_CODE")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Print POs"
                If dst.Tables("POTORDRX").Select("SEL='1'").Length = 0 Then
                    EMsg &= vbCr & "No POs Selected"
                End If

            Case "Transmit POs"
                If dst.Tables("POTORDRX").Select("SEL='1'").Length = 0 Then
                    EMsg &= vbCr & "No POs Selected"
                End If

                If EMsg = "" Then
                    PO_ORDER_NOs.Clear()
                    VEND_CODEs.Clear()
                    For Each row As DataRow In dst.Tables("POTORDRX").Select("SEL='1'")

                        Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                        PO_ORDER_NOs.Add(PO_ORDER_NO)
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                        If ROWs("POTPARM1").Item("PO_PARM_APPR_REQD") & "" = "1" Then
                            If rowPOTORDR1.Item("PO_APPR_PENDING") & "" <> "1" Or rowPOTORDR1.Item("PO_APPR_BY") & "" = "" Then
                                EMsg &= vbCr & "PO " & PO_ORDER_NO & " not Approved"
                            End If
                        End If

                        Dim VEND_CODE As String = row.Item("VEND_CODE")
                        If Not VEND_CODEs.Contains(VEND_CODE) Then
                            VEND_CODEs.Add(VEND_CODE)
                            Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
                            If rowAPTVEND1.Item("VEND_STOP_PURCHASE") & "" = "1" Then
                                EMsg &= vbCr & "Vendor is On Hold for Purchasing"
                            End If
                        End If
                    Next
                    If VEND_CODEs.Count > 1 Then
                        EMsg &= vbCr & "Must Select only 1 Supplier at a time when Transitting"
                    End If
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Expedite POs"
                Dim VEND_CODE As String = grdPOTORDRV.ActiveRow.Cells("VEND_CODE").Value & ""

                Dim sqlw As String = "VEND_CODE = '" & VEND_CODE & "' and SEL='1'"
                If dst.Tables("POTORDR1").Select(sqlw & " and PO_XMIT_IND = '1'").Length = 0 Then
                    EMsg &= vbCr & "No POs selected to Expedite"
                End If
                If dst.Tables("POTORDR1").Select(sqlw & " and  ISNULL(PO_XMIT_IND,'0') <> '1'").Length <> 0 Then
                    EMsg &= vbCr & "Cannot Expedite a PO that has not been Transmitted"
                End If

                If EMsg = "" Then
                    PO_ORDER_NOs.Clear()

                    For Each row As DataRow In dst.Tables("POTORDR1").Select(sqlw)
                        Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                        PO_ORDER_NOs.Add(PO_ORDER_NO)
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                            Exit Sub
                        Else
                            Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                            If rowPOTORDR1.Item("PO_XMIT_IND") & "" <> "1" Or rowPOTORDR1.Item("PO_STATUS") & "" <> "O" Then
                                EMsg &= "Status has changed for PO " & PO_ORDER_NO & " - please refresh this list"
                            End If
                        End If
                    Next
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                Refresh_POs()

            Case "Print"
                Print_Record()

            Case "Print POs"
                PO_ORDER_NOs.Clear()
                For Each row As DataRow In dst.Tables("POTORDRX").Select("SEL='1'")
                    PO_ORDER_NOs.Add(row.Item("PO_ORDER_NO"))
                Next
                Print_POs(PO_ORDER_NOs)

            Case "Transmit POs"
                Transmit_POs()


            Case "Expedite POs"
                Expedite_POs()


        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    '.Items("View").Settings.Enabled = not_iScreenMode
                    '.Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode

                    '.Items("View").Visible = (EntryMode = "L" Or Not ScreenMode)
                    '.Items("Done").Visible = (EntryMode = "L" And ScreenMode)
                    '.Items("Print").Visible = ScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Setup_tabMain()
        tabMain.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTORDR1", "POTORDR2", "POTORDRV", "POTORDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Refresh_POs()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        ' Sort_grdColumns(grdPOTSHIP3, "PO_SHIPMENT_LNO")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Stop
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Preparing " & Me.Text)

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("PORWREC2")
        Print_Report_End()

        '    Me.Cursor = Cursors.Default
        '    ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_SHIPMENT_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTORDRX, "SSSBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "PO Entry", "Select All", "De-Select All", "Select Selected", "Select All VENDOR")
        Load_Popup_Menu(grdPOTORDR1, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry", "PO Entry")
        Load_Popup_Menu(grdPOTORDR2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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
            Case "grdPOTORDRX"
                tlb_btn = DirectCast(tlb_pop.Tools("Select All VENDOR"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Visible = True
                    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Value
                    tlb_btn.Tag = VEND_CODE
                    tlb_btn.SharedProps.Caption = "Select All " & VEND_CODE
                End If
 
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All"
                For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("")
                    rowPOTORDRX.Item("SEL") = "1"
                Next
                'For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                '    If grow.IsDataRow Then
                '        grow.Cells("SEL").Value = "1"
                '        grow.Update()
                '    End If
                'Next

            Case "De-Select All"
                For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("")
                    rowPOTORDRX.Item("SEL") = "0"
                Next
                'For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                '    If grow.IsDataRow Then
                '        grow.Cells("SEL").Value = "0"
                '        grow.Update()
                '    End If
                'Next

            Case "Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells("SEL").Value = "1"
                    grow.Update()
                Next

            Case "Select All VENDOR"
                tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Dim VEND_CODE As String = tlb_btn.Tag
                For Each row As DataRow In dst.Tables("POTORDRX").Select("VEND_CODE = '" & VEND_CODE & "'")
                    row.Item("SEL") = "1"
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "PO Entry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("Edit", PO_ORDER_NO, e.Tool.Key, "POFORDR1", "F", "POE")

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
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

    Sub Load_POTORDR1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR1)
        ASCMAIN1.sql = "Select POTORDR1.*, ARTCUST1.CUST_NAME, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL from POTORDR1,ARTCUST1 " _
            & " where POTORDR1.PO_ORDER_NO in " & vbCrLf _
            & "(Select Distinct PO_ORDER_NO from POTORDR2 where PO_STATUS = 'O' " & vbCrLf _
            & " and PO_DATE_SHIP_BY <= '" & Format(dteShipDate.Value, "dd-MMM-yyyy") & "')" & vbCrLf _
            & " and ARTCUST1.CUST_CODE (+) = POTORDR1.CUST_CODE"
        ASCMAIN1.sql = "Insert into " & POTORDR1 & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & vbCrLf _
            & "  Select POTORDR2.PO_ORDER_NO" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_ORD) PO_QTY_ORD" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_REC) PO_QTY_REC" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_ORD * POTORDR2.PO_COST) PO_AMT_ORD" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_SHP * POTORDR2.PO_COST) PO_AMT_SHP" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_REC * POTORDR2.PO_COST) PO_AMT_REC" & vbCrLf _
            & " , SUM (POTORDR2.PO_QTY_OPN * POTORDR2.PO_COST) PO_AMT_OPN" & vbCrLf _
            & " , MIN (POTORDR2.PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MIN" & vbCrLf _
            & " , MAX (POTORDR2.PO_DATE_SHIP_BY) PO_DATE_SHIP_BY_MAX" & vbCrLf _
            & "   from POTORDR2," & POTORDR1 & " POTORDR1" & vbCrLf _
            & "  where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
            & "  group by POTORDR2.PO_ORDER_NO;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & POTORDR1 & " Set" & vbCrLf _
            & "     PO_QTY_ORD = R1.PO_QTY_ORD" & vbCrLf _
            & "   , PO_QTY_SHP = R1.PO_QTY_SHP" & vbCrLf _
            & "   , PO_QTY_REC = R1.PO_QTY_REC" & vbCrLf _
            & "   , PO_QTY_OPN = R1.PO_QTY_OPN" & vbCrLf _
            & "   , PO_AMT_ORD = R1.PO_AMT_ORD" & vbCrLf _
            & "   , PO_AMT_SHP = R1.PO_AMT_SHP" & vbCrLf _
            & "   , PO_AMT_REC = R1.PO_AMT_REC" & vbCrLf _
            & "   , PO_AMT_OPN = R1.PO_AMT_OPN" & vbCrLf _
            & "   , PO_DATE_SHIP_BY_MIN = R1.PO_DATE_SHIP_BY_MIN" & vbCrLf _
            & "   , PO_DATE_SHIP_BY_MAX = R1.PO_DATE_SHIP_BY_MAX" & vbCrLf _
            & "   where PO_ORDER_NO = R1.PO_ORDER_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End; " & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Fill_Records("POTORDR1")
        Fill_Records("POTORDRV")
        Sort_grdColumns(grdPOTORDRV, "VEND_CODE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Create_Temporary_Tables()

        ASCMAIN1.sql = "Select POTORDR1.*, ARTCUST1.CUST_NAME from POTORDR1,ARTCUST1 where ROWNUM < 1"
        POTORDR1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add Primary Key (PO_ORDER_NO)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_QTY_ORD NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_QTY_SHP NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_QTY_REC NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_QTY_OPN NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_AMT_ORD NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_AMT_SHP NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_AMT_REC NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_AMT_OPN NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_DATE_SHIP_BY_MIN DATE")
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add PO_DATE_SHIP_BY_MAX DATE")


        'ASCMAIN1.sql = "Select * from POTORDR2 where ROWNUM < 1"
        'POTORDR2 = ASCMAIN1.Temp_Table
        'ASCDATA1.ExecuteSQL("Alter Table " & POTORDR2 & " Add Primary Key (PO_ORDER_NO,PO_ORDER_LNO)")
    End Sub

    Sub Load_POTORDRX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim sqlw2 As String = IIf(ROWs("POTPARM1").Item("PO_PARM_APPR_REQD") & "" = "1", " and NVL(PO_APPR_PENDING,'0') = '1' and PO_APPR_BY is Not Null", "")

        Dim sqlw As String = ""
        If optStatus.Value = "N" Then
            sqlw = " and POTORDR1.PO_STATUS = 'O' and POTORDR1.PO_XMIT_IND = '0'" & sqlw2
            grdPOTORDRX.Text = "Open POs NOT Transmitted"
        ElseIf Not InquiryMode Or optStatus.Value = "O" Then
            sqlw = " and POTORDR1.PO_STATUS = 'O'" & sqlw2
            grdPOTORDRX.Text = "Open POs"
        End If

        Dim sqlDTL As String = "Select PO_ORDER_NO" & vbCrLf _
            & ", SUM (PO_QTY_ORD) PO_QTY_ORD" & vbCrLf _
            & ", SUM (PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & ", SUM (PO_QTY_REC) PO_QTY_REC" & vbCrLf _
            & ", SUM (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", SUM (PO_QTY_ORD * PO_COST) PO_AMT_ORD" & vbCrLf _
            & ", SUM (PO_QTY_SHP * PO_COST) PO_AMT_SHP" & vbCrLf _
            & ", SUM (PO_QTY_REC * PO_COST) PO_AMT_REC" & vbCrLf _
            & ", SUM (PO_QTY_OPN * PO_COST) PO_AMT_OPN" & vbCrLf _
            & " from  POTORDR2 where PO_ORDER_NO in " & vbCrLf _
            & " (Select PO_ORDER_NO from POTORDR1 " & ASCMAIN1.SQL_Add_WHERE(sqlw) & ")" _
            & " group by PO_ORDER_NO"

        ASCMAIN1.sql = "Select POTORDR1.*, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", X.PO_QTY_ORD, X.PO_QTY_SHP, X.PO_QTY_REC, X.PO_QTY_OPN " & vbCrLf _
            & ", X.PO_AMT_ORD, X.PO_AMT_SHP, X.PO_AMT_REC, X.PO_AMT_OPN " & vbCrLf _
            & " from (" & sqlDTL & ") X,POTORDR1,ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE (+) = POTORDR1.CUST_CODE" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_NO = X.PO_ORDER_NO" & vbCrLf _
            & sqlw
        Fill_Records("POTORDRX", "", True, ASCMAIN1.sql)

        grdPOTORDRX.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdPOTORDRX.DisplayLayout.Bands(0).SortedColumns.Add("WHSE_CODE", False, True)
        grdPOTORDRX.DisplayLayout.Bands(0).SortedColumns.Add("VEND_CODE", False)
        grdPOTORDRX.DisplayLayout.Bands(0).SortedColumns.Add("PO_DATE_SHIP_BY", False)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_POTORDRX()
        Setup_tabMain()
    End Sub

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
        Refresh_POs()
    End Sub

    Sub Setup_tabMain()
        UltraExplorerBar1.Groups("Status Filter").Visible = (tabMain.SelectedTab.Key = "Transmit POs")
        UltraExplorerBar1.Groups("Ship Date Cut-off").Visible = (tabMain.SelectedTab.Key = "Expedite POs")
        With UltraExplorerBar1.Groups("Screen Control")
            .Items("Print POs").Visible = (tabMain.SelectedTab.Key = "Transmit POs")
            .Items("Transmit POs").Visible = (tabMain.SelectedTab.Key = "Transmit POs")
            If optStatus.Value = "N" Then
                '.Items("Transmit POs").Settings.Enabled = DefaultableBoolean.True
                .Items("Transmit POs").Text = "Transmit POs"
            Else
                '.Items("Transmit POs").Settings.Enabled = DefaultableBoolean.False
                .Items("Transmit POs").Text = "Re-Transmit POs"
            End If

            .Items("Expedite POs").Visible = (tabMain.SelectedTab.Key = "Expedite POs")
        End With
    End Sub

    Private Sub grdPOTORDRV_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDRV.AfterRowActivate
        Setup_POTORDRV()
    End Sub

    Sub Setup_POTORDRV()
        If grdPOTORDRV.ActiveRow Is Nothing OrElse Not grdPOTORDRV.ActiveRow.IsDataRow Then
            splPOTORDR1.Visible = False
        Else
            Dim VEND_CODE As String = grdPOTORDRV.ActiveRow.Cells("VEND_CODE").Value
            Dim dvw As DataView = DirectCast(grdPOTORDR1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "VEND_CODE = '" & VEND_CODE & "'"
            Sort_grdColumns(grdPOTORDR1, "PO_ORDER_NO")
            grdPOTORDR1.Text = "POs for Vendor " & VEND_CODE
            splPOTORDR1.Visible = True
            SETUP_POTORDR1()
        End If
    End Sub

    Private Sub grdPOTORDR1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDR1.AfterRowActivate
        SETUP_POTORDR1()
    End Sub

    Sub Setup_POTORDR1()
        If grdPOTORDR1.ActiveRow Is Nothing OrElse Not grdPOTORDR1.ActiveRow.IsDataRow Then
            grdPOTORDR2.Visible = False
        Else
            grdPOTORDR2.Visible = True
            Dim PO_ORDER_NO As String = grdPOTORDR1.ActiveRow.Cells("PO_ORDER_NO").Value
            Fill_Records("POTORDR2", New Object() {PO_ORDER_NO})
            Sort_grdColumns(grdPOTORDR2, "PO_ORDER_LNO")
            grdPOTORDR2.Text = "PO " & PO_ORDER_NO & "Details"
        End If
    End Sub

    Sub Refresh_POs()
        If tabMain.SelectedTab.Key = "Transmit POs" Then
            Load_POTORDRX()
        Else
            Load_POTORDR1()
        End If
    End Sub

    Function Print_POs(PO_ORDER_NOs As List(Of String), Optional make_pdf As Boolean = False, Optional FILENAME_body As String = "") As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing POs")

        Dim REPORTFILE As String = "POROPRT1"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and PO_ORDER_NO in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"})

        'To fill the report's dataset with data from this form's dataset:
        'With REPORTS(REPORTFILE).clsASCBASE1
        '    .EnforceConstraints(False)
        '    For Each TABLE_NAME As String In New String() {"SOTPPDI1", "SOTPPDI2", "SOTPPDI3", "SOTINVH1", "SOTSVIA1"}
        '        .dst.Tables(TABLE_NAME).Rows.Clear()
        '        Dim SQL As String = ""
        '        If TABLE_NAME = "SOTINVH1" Then
        '            SQL = "ORDR_NO = '" & ORDR_NO & "'"
        '        End If

        '        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(Sql)
        '            Dim rowr As DataRow = .dst.Tables(TABLE_NAME).NewRow
        '            If TABLE_NAME = "SOTPPDI2" Or TABLE_NAME = "SOTPPDI3" Or TABLE_NAME = "SOTINVH1" Then

        '                For I As Integer = 0 To .dst.Tables(TABLE_NAME).Columns.Count - 1
        '                    Dim COLUMN_NAME As String = .dst.Tables(TABLE_NAME).Columns(I).ColumnName
        '                    rowr.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
        '                Next
        '            Else
        '                rowr.ItemArray = row.ItemArray
        '            End If
        '            .dst.Tables(TABLE_NAME).Rows.Add(rowr)
        '        Next
        '    Next
        '    .EnforceConstraints(True)
        'End With

        Dim REPORT_NO As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("FORM_TYPE", "P")

            Dim RPT As String = REPORTFILE
            Dim PO_PARM_PO_RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
            If PO_PARM_PO_RPT <> "" Then RPT = PO_PARM_PO_RPT

            If make_pdf Then
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , "PDF", FILENAME_body, False)
            Else
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , , , False)
            End If
            .Print_Report_End(make_pdf, make_pdf)
            If ASCMAIN1.CLIENT = "NYA" Then

            End If
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return REPORT_NO
    End Function

    Sub Transmit_POs()
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim VEND_CODE As String = VEND_CODEs(0)
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

        Dim PO_ORDER_NOs_1 As New List(Of String)
        For Each PO_ORDER_NO As String In PO_ORDER_NOs
            PO_ORDER_NOs_1.Clear()
            PO_ORDER_NOs_1.Add(PO_ORDER_NO)
            Dim REPORT_NO As String = Print_POs(PO_ORDER_NOs_1, True, PO_ORDER_NO)
            ATTACHMENTs.Add(PO_ORDER_NO & ".pdf", ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".pdf")
        Next

        Dim SUBJECT As String = ""
        Dim PFX As String = ""
        If ASCMAIN1.CLIENT = "RGI" Then PFX = "Regency  "
        If PO_ORDER_NOs.Count = 1 Then
            SUBJECT = PFX & "PO " & PO_ORDER_NOs(0)
        Else
            SUBJECT = PFX & "POs " & Join(PO_ORDER_NOs.ToArray, ",")
        End If

        If optStatus.Value = "O" Then
            SUBJECT &= " - Re-Transmit"
        End If

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")
        EMAIL_ADDRESSs.Add(rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "", rowAPTVEND1.Item("VEND_PURCH_CONTACT") & "")

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                SUBJECT, "PO", False, True, VEND_CODE, rowAPTVEND1.Item("VEND_NAME"), "Supplier")

        If SEND_NO <> "" Then
            If optStatus.Value = "N" Then

                For Each PO_ORDER_NO As String In PO_ORDER_NOs
                    Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                    Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
                    ASCMAIN1.sql = "Update POTORDR1 " & vbCrLf _
                        & " Set PO_PRINTED_IND = '1', PO_DATE_PRINTED = SYSDATE" & vbCrLf _
                        & ", PO_XMIT_IND = '1', PO_XMIT_BY = '" & ASCMAIN1.USER_ID & "', PO_XMIT_DATE = SYSDATE, PO_XMIT_XNO = '" & XNO & "'" & vbCrLf _
                        & " where (PO_ORDER_NO) in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"
                    ASCDATA1.ExecuteSQL()
                    Dim rowPOTORDRX As DataRow = dst.Tables("POTORDRX").Rows.Find(PO_ORDER_NO)
                    rowPOTORDRX.Delete()

                    My.Computer.FileSystem.CopyFile( _
                        ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".PDF", _
                        ASCMAIN1.Folders("Archive") & "PO\" & PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV) & ".PDF", True)
                Next

                ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                    & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'PO-XMIT','PO Transmitted', '" & SEND_NO & "'" _
                    & " from POTORDR1 " & vbCrLf _
                    & " where (PO_ORDER_NO) in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Insert into POTORDRZ (PO_ORDER_NO,PO_HDR_CTR_REV,PO_ORDER_LNO" & vbCrLf _
                    & ",STYLE_CODE,COLOR_CODE,PO_QTY_ORD,PO_COST,PO_DATE_SHIP_BY,PO_STATUS,CARTON_PACK_QTY)" & vbCrLf _
                    & " Select POTORDR2.PO_ORDER_NO, NVL(POTORDR1.PO_HDR_CTR_REV,0), POTORDR2.PO_ORDER_LNO" & vbCrLf _
                    & ", POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR2.PO_QTY_ORD" & vbCrLf _
                    & ", POTORDR2.PO_COST, POTORDR2.PO_DATE_SHIP_BY, POTORDR2.PO_STATUS, POTORDR2.CARTON_PACK_QTY" & vbCrLf _
                    & " from POTORDR1,POTORDR2" & vbCrLf _
                    & " where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Insert into POTORDRH (PO_ORDER_NO,PO_HDR_CTR_REV,PO_REVISION_NOTE,INIT_OPER,INIT_DATE,LAST_OPER,LAST_DATE)" & vbCrLf _
                    & " Select PO_ORDER_NO, NVL(PO_HDR_CTR_REV,0), DECODE(NVL(PO_HDR_CTR_REV,0),0,'Original',PO_REVISION_NOTE), LAST_OPER, LAST_DATE, LAST_OPER, LAST_DATE" & vbCrLf _
                    & " from POTORDR1" & vbCrLf _
                    & " where POTORDR1.PO_ORDER_NO in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"
                ASCDATA1.ExecuteSQL()


                If rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "" = "" Then
                    Dim rowTATSEND1 As DataRow = LookUp("TATSEND1", SEND_NO)
                    Dim SEND_TO As String = rowTATSEND1.Item("SEND_TO") & ""
                    If SEND_TO <> "" And Not SEND_TO.Contains(";") Then
                        ASCMAIN1.sql = "Update APTVEND1 Set VEND_PURCH_EMAIL = :PARM1 where VEND_CODE = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {rowTATSEND1.Item("SEND_TO"), VEND_CODE})
                    End If
                End If
            Else
                ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                    & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'PO-XMIT','PO Re-Transmitted', '" & SEND_NO & "'" _
                    & " from POTORDR1 " & vbCrLf _
                    & " where (PO_ORDER_NO) in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"
                ASCDATA1.ExecuteSQL()
            End If
        End If
    End Sub

    Private Sub grdPOTORDR1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDR1.InitializeRow
        If e.Row.Cells("PO_XMIT_IND").Value & "" <> "1" Then
            e.Row.Cells("PO_ORDER_NO").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("PO_ORDER_NO").ToolTipText = "PO must have been Transmitted in order to be Expedited"
        Else
            e.Row.Cells("PO_ORDER_NO").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Expedite_POs()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating Expediting Reports")

        Dim EXPEDITE_NO As String = ASCMAIN1.Next_Control_No("POTORDR1.EXPEDITE_NO")
        TAC.POCMAIN1.Generate_Expedite_POs_XLS(Me, EXPEDITE_NO, PO_ORDER_NOs)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim VEND_CODE As String = grdPOTORDRV.ActiveRow.Cells("VEND_CODE").Value
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

        ATTACHMENTs.Add("Open_PO_Report_" & EXPEDITE_NO & ".xls", ASCMAIN1.Folders("Temp") & "Open_PO_Report_" & EXPEDITE_NO & ".xls")

        'Dim PO_ORDER_NOs_1 As New List(Of String)
        'For Each PO_ORDER_NO As String In PO_ORDER_NOs
        '    PO_ORDER_NOs_1.Clear()
        '    PO_ORDER_NOs_1.Add(PO_ORDER_NO)
        '    Dim REPORT_NO As String = Print_POs(PO_ORDER_NOs_1, True, PO_ORDER_NO)
        '    ATTACHMENTs.Add(PO_ORDER_NO & ".pdf", ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".pdf")
        'Next

        Dim SUBJECT As String = ""
        If PO_ORDER_NOs.Count = 1 Then
            SUBJECT = "Regency PO " & PO_ORDER_NOs(0)
        Else
            SUBJECT = "Regency POs " & Join(PO_ORDER_NOs.ToArray, ",")
        End If

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")
        EMAIL_ADDRESSs.Add(rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "", rowAPTVEND1.Item("VEND_PURCH_CONTACT") & "")

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                SUBJECT, "POEXP", False, True, VEND_CODE, rowAPTVEND1.Item("VEND_NAME"), "Supplier")

        If SEND_NO <> "" Then
            'For Each PO_ORDER_NO As String In PO_ORDER_NOs
            '    Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
            '    Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
            '    ASCMAIN1.sql = "Update POTORDR1 " & vbCrLf _
            '        & " Set PO_PRINTED_IND = '1', PO_DATE_PRINTED = SYSDATE" & vbCrLf _
            '        & ", PO_XMIT_IND = '1', PO_XMIT_BY = '" & ASCMAIN1.USER_ID & "', PO_XMIT_DATE = SYSDATE, PO_XMIT_XNO = '" & XNO & "'" & vbCrLf _
            '        & " where (PO_ORDER_NO) in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"
            '    ASCDATA1.ExecuteSQL()
            '    Dim rowPOTORDRX As DataRow = dst.Tables("POTORDRX").Rows.Find(PO_ORDER_NO)
            '    rowPOTORDRX.Delete()

            '    My.Computer.FileSystem.CopyFile( _
            '        ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".PDF", _
            '        ASCMAIN1.Folders("Archive") & "PO\" & PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV) & ".PDF")
            'Next

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'PO-XPED','PO Expedited', '" & SEND_NO & "'" _
                & " from POTORDR1 " & vbCrLf _
                & " where (PO_ORDER_NO) in ('" & Join(PO_ORDER_NOs.ToArray, "','") & "')"
            ASCDATA1.ExecuteSQL()

            If rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "" = "" Then
                Dim rowTATSEND1 As DataRow = LookUp("TATSEND1", SEND_NO)
                ASCMAIN1.sql = "Update APTVEND1 Set VEND_PURCH_EMAIL = :PARM1 where VEND_CODE = :PARM2"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {rowTATSEND1.Item("SEND_TO"), VEND_CODE})
            End If
        End If
    End Sub

    Private Sub grdPOTORDRX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdPOTORDRX.InitializeLayout

    End Sub
End Class