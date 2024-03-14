Imports Infragistics.Win.UltraWinGrid

Public Class POFVNLB1
    Dim POTVNLB1 As String 'TABLE_NAME
    Dim sqlPOTVNLB1 As String
    Dim subUPCSupport As Boolean = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "SELECT T1.PO_ORDER_NO, t1.whse_code,T2.STYLE_CODE, T4.STYLE_DESC, T2.COLOR_CODE, T3.UPC_CODE, T4.COUNTRY_CODE,T2.PO_QTY_OPN" & vbCrLf _
                & ", T4.LABEL_TYPE_CODE, T6.CUST_NAME, T5.CUST_SKU, T5.CUST_UPC, T5.CUST_STYLE_CODE, T5.CUST_COLOR_CODE, T5.CUST_SIZE_CODE, T5.STYLE_RETAIL, T1.VEND_CODE, T4.STYLE_UOM " & vbCrLf _
                & " FROM POTORDR1 T1, POTORDR2 T2, ICTSTYC1 T3, ICTSTYL1 T4, SOTORDR2 T5, SOTORDR1 T6 " & vbCrLf _
                & " WHERE T1.PO_STATUS = 'O'" & vbCrLf _
                & " AND NVL(T2.PO_QTY_OPN,0) <> 0 " & vbCrLf _
                & " AND T1.PO_ORDER_NO = T2.PO_ORDER_NO " & vbCrLf _
                & " AND T2.STYLE_CODE = T4.STYLE_CODE " & vbCrLf _
                & " AND T2.STYLE_CODE = T3.STYLE_CODE " & vbCrLf _
                & " AND T2.COLOR_CODE = T3.COLOR_CODE " & vbCrLf _
                & " AND T2.ORDR_NO = T6.ORDR_NO (+) " & vbCrLf _
                & " AND T2.ORDR_NO = T5.ORDR_NO (+) " & vbCrLf _
                & " AND T2.ORDR_LNO = T5.ORDR_LNO (+) " & vbCrLf _
                & " ORDER BY T1.PO_ORDER_NO, T2.STYLE_CODE, T2.COLOR_CODE " & vbCrLf


            sqlPOTVNLB1 = ASCMAIN1.sql

            POTVNLB1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & POTVNLB1
            Create_TDA(.Tables.Add("POTVNLB1"), POTVNLB1, "**", 0, True)
            ' Create_TDA(.Tables.Add, "POTBATC1", "**", 0, True, "V", 1)

            If subUPCSupport Then
                ASCMAIN1.sql = $"Select ICTXLSPS.* 
                    from ICTXLSPS, {POTVNLB1} POTVNLB1
                    where ICTXLSPS.STYLE_CODE = POTVNLB1.STYLE_CODE 
                    AND ICTXLSPS.COLOR_CODE = POTVNLB1.COLOR_CODE"
                Create_TDA(.Tables.Add, "ICTXLSPS", "**", 0, False)
            End If

        End With

        grdPOTVNLB1.DataSource = dst.Tables("POTVNLB1")

        ' Create_Summary(grdPOTVNLB1, "STYLE_CODE", "Count")

        With grdPOTVNLB1.DisplayLayout.Bands(0)

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink

            Next
        End With


        spl.Panel1Collapsed = True


        If subUPCSupport Then
            grdICTXLSPS.DataSource = dst.Tables("ICTXLSPS")
            Create_Summary(grdICTXLSPS, "SET_LNO", "Count")
            Sort_grdColumns(grdICTXLSPS, "SET_LNO", True)
            splSets.Panel2Collapsed = True
        End If



    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)


            Case "Done"
                Mode_Settings(False)

            Case "Update"
                'Update_Record()
                'Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If subUPCSupport Then
            With grdICTXLSPS.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        End If

        If ScreenMode Then
        Else
            'Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POFVNLB1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If subUPCSupport Then
            dst.Tables("ICTXLSPS").Rows.Clear()
        End If

        EnforceConstraints(True)

        ' Absx1.txtFor("CUST_CODE").Text = ""
        ' Absx1.txtFor("SREP_CODE").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "E" Then
        Else
            ASCMAIN1.sql = "TRUNCATE TABLE " & POTVNLB1
            ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "INSERT INTO " & POTVNLB1 & sqlPOTVNLB1
            ' ASCMAIN1.sql = "INSERT INTO " & POTVNLB1 & " SELECT X.*,'','','','','' FROM (" & sqlPOTVNLB1 & ") X "
            'ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO " & POTVNLB1 & " SELECT X.* FROM (" & sqlPOTVNLB1 & ") X "
            ' ASCMAIN1.sql = "INSERT INTO " & POTCONF1 & " SELECT X.*,'','','','','' FROM (" & sqlPOTCONF1 & ") X "
            ASCDATA1.ExecuteSQL()

            Fill_Records("POTVNLB1")
            Fill_Records("ICTXLSPS")
        End If


        Sort_grdColumns(grdPOTVNLB1, "STYLE_CODE,COLOR_CODE")

        Set_background_colors()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        EntryMode = ""

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub




#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTVNLB1, "SSBS", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Update Column")
        If subUPCSupport Then
            Load_Popup_Menu(grdICTXLSPS, "B", "Style Master File")
        End If
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        Select Case e.SourceControl.Name
            Case "grdPOFVNLB1"
                tlb_sbt = DirectCast(tlb_pop.Tools("Update Column"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = False
                If grd.ActiveCell IsNot Nothing Then
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    Dim row As DataRow = dst.Tables("TATCOLS1").Rows.Find(COLUMN_NAME)
                    If row IsNot Nothing Then
                        tlb_sbt.SharedProps.Visible = True
                        tlb_sbt.SharedProps.Caption = "Update " & row.Item("COLUMN_CAPTION")
                        tlb_sbt.Tag = ""
                        tlb_sbt.Checked = (row.Item("SEL") = "1")
                        tlb_sbt.Tag = COLUMN_NAME
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdPOFVNLB1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Update Column"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Dim COLUMN_NAME As String = tlb_sbt.Tag & ""
                If COLUMN_NAME <> "" Then
                    Dim row As DataRow = dst.Tables("TATCOLS1").Rows.Find(COLUMN_NAME)
                    If row IsNot Nothing Then
                        If tlb_sbt.Checked Then
                            row.Item("SEL") = "1"
                        Else
                            row.Item("SEL") = "0"
                        End If
                        Set_background_colors()
                    End If
                End If

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

    Sub Set_background_colors()


    End Sub

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub



#End Region

    Private Sub UltraLabel7_Click(sender As System.Object, e As System.EventArgs) Handles UltraLabel7.Click

    End Sub

    Private Sub grdPOTVNLB1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdPOTVNLB1.AfterRowActivate
        Setup_grdPOTORDR2_ActiveRow()
    End Sub
    Sub Setup_grdPOTORDR2_ActiveRow()

        If (grdPOTVNLB1.ActiveRow Is Nothing) Then
            splSets.Panel2Collapsed = True
        Else
            If subUPCSupport Then
                Setup_Sub_UPC_Grid()
            End If
        End If

    End Sub

    Sub Setup_Sub_UPC_Grid()

        Dim STYLE_CODE As String = grdPOTVNLB1.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = grdPOTVNLB1.ActiveRow.Cells("COLOR_CODE").Value & ""

        Dim dvw As DataView = DirectCast(grdICTXLSPS.DataSource, DataTable).DefaultView
        dvw.RowFilter = $"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'"

        splSets.Panel2Collapsed = Not Line_Has_Sub_UPCs(STYLE_CODE, COLOR_CODE)

        If Not splSets.Panel2Collapsed Then
            Dim PO_ORDER_NO As Int32 = Val(grdPOTVNLB1.ActiveRow.Cells("PO_ORDER_NO").Value & "")
            Sort_grdColumns(grdICTXLSPS, "SET_LNO")
            grdICTXLSPS.Text = $"Sub UPCS for Style/Color {STYLE_CODE}/{COLOR_CODE} on PO {PO_ORDER_NO}"
        End If

    End Sub

    Function Line_Has_Sub_UPCs(STYLE_CODE As String, COLOR_CODE As String) As Boolean
        Return dst.Tables("ICTXLSPS").Select($"STYLE_CODE = '{STYLE_CODE}' AND COLOR_CODE = '{COLOR_CODE}'").Count > 0
    End Function

    Private Sub grdPOTVNLB1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTVNLB1.InitializeRow
        If subUPCSupport Then
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & ""
            If Line_Has_Sub_UPCs(STYLE_CODE, COLOR_CODE) Then
                e.Row.Cells("STYLE_CODE").Appearance.ForeColor = Drawing.Color.Blue
            End If
        End If
    End Sub
End Class