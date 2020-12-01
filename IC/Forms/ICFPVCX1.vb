Public Class ICFPVCX1
    Private sql As New System.Text.StringBuilder With {.Length = 0}
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            sql.Length = 0
            sql.AppendLine("SELECT")
            sql.AppendLine("S1.STYLE_CODE,")
            sql.AppendLine("S1.STYLE_DESC,")
            sql.AppendLine("S1.CARTON_PACK_QTY,")
            sql.AppendLine("S1.STYLE_PRICE,")
            sql.AppendLine("S1.CASE_CUBE,")
            sql.AppendLine("PV.LIGHT_TYPE_CODE,")
            sql.AppendLine("LT.LIGHT_TYPE_DESC,")
            sql.AppendLine("PV.LIGHT_COLOR_CODE,")
            sql.AppendLine("LC.LIGHT_COLOR_DESC,")
            sql.AppendLine("PV.COLLECTION_GROUP_CODE,")
            sql.AppendLine("CG.COLLECTION_GROUP_DESC,")
            sql.AppendLine("PV.COLLECTION_CODE,")
            sql.AppendLine("CL.COLLECTION_DESC,")
            sql.AppendLine("PV.TREE_SHAPE_CODE,")
            sql.AppendLine("TS.TREE_SHAPE_DESC,")
            sql.AppendLine("PV.SETUP_CODE,")
            sql.AppendLine("ST.SETUP_DESC,")
            sql.AppendLine("PV.HEIGHT,")
            sql.AppendLine("PV.DIAMETER,")
            sql.AppendLine("PV.PVC_LENGTH,")
            sql.AppendLine("PV.LIGHT_COUNT,")
            sql.AppendLine("PV.G40_COUNT,")
            sql.AppendLine("PV.C7_COUNT,")
            sql.AppendLine("PV.CANDLE_COUNT,")
            sql.AppendLine("PV.TIP_COUNT,")
            sql.AppendLine("PV.CONE_COUNT")
            sql.AppendLine("FROM ICTSTYL1 S1, ICTPVC01 PV, ICTPVCLT LT, ICTPVCCG CG, ICTPVCCL CL, ICTPVCTS TS, ICTPVCST ST, ICTPVCLC LC")
            sql.AppendLine("WHERE S1.STYLE_CODE = PV.STYLE_CODE")
            sql.AppendLine("AND PV.LIGHT_TYPE_CODE = LT.LIGHT_TYPE_CODE (+)")
            sql.AppendLine("AND PV.COLLECTION_GROUP_CODE = CG.COLLECTION_GROUP_CODE(+)")
            sql.AppendLine("And PV.COLLECTION_CODE = CL.COLLECTION_CODE (+)")
            sql.AppendLine("And PV.TREE_SHAPE_CODE = TS.TREE_SHAPE_CODE (+)")
            sql.AppendLine("And PV.SETUP_CODE = ST.SETUP_CODE (+)")
            sql.AppendLine("And PV.LIGHT_COLOR_CODE = LC.LIGHT_COLOR_CODE (+)")
            ASCMAIN1.sql = sql.ToString
            Create_TDA(.Tables.Add, "ICTPVCX1", "**", 0, False, "", 0)
            With .Tables("ICTPVCX1")
                .Columns.Add("COLORS", GetType(System.String))
                .Columns.Add("BAR_CODE", GetType(System.String))
                .Columns.Add("PB1", GetType(System.Double))
                .Columns.Add("PB2", GetType(System.Double))
                .Columns.Add("PB3", GetType(System.Double))
                .Columns.Add("PB4", GetType(System.Double))
                .Columns.Add("AVAIL", GetType(System.Double))
            End With

            sql.Length = 0
            sql.AppendLine("SELECT * FROM")
            sql.AppendLine("(")
            sql.AppendLine("SELECT")
            sql.AppendLine("S1.STYLE_CODE,")
            sql.AppendLine("C1.COLOR_CODE,")
            sql.AppendLine("S1.STYLE_STATUS,")
            sql.AppendLine("C1.STYLE_COLOR_STATUS,")
            sql.AppendLine("S1.STYLE_DESC,")
            sql.AppendLine("SUM((NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))) AS AVAIL")
            sql.AppendLine("FROM ICTSTYL1 S1, ICTSTYC1 C1, ICTSTAT2 S2")
            sql.AppendLine("WHERE S1.STYLE_CODE = C1.STYLE_CODE")
            sql.AppendLine("AND C1.STYLE_CODE = S2.STYLE_CODE (+)")
            sql.AppendLine("AND C1.COLOR_CODE = S2.COLOR_CODE (+)")
            sql.AppendLine("AND S2.WHSE_CODE = 'MS'")
            sql.AppendLine("GROUP BY")
            sql.AppendLine("S1.STYLE_CODE,")
            sql.AppendLine("C1.COLOR_CODE,")
            sql.AppendLine("S1.STYLE_STATUS,")
            sql.AppendLine("C1.STYLE_COLOR_STATUS,")
            sql.AppendLine("S1.STYLE_DESC")
            sql.AppendLine(")")
            sql.AppendLine("WHERE (STYLE_COLOR_STATUS = 'A' OR AVAIL > 0)")
            ASCMAIN1.sql = sql.ToString
            Create_TDA(.Tables.Add, "ICTSTATX", "**", 0, False, "", 0)
        End With

        grdICTPVCX1.DataSource = dst.Tables("ICTPVCX1")

        Create_Summary(grdICTPVCX1, "STYLE_CODE", "Count")

        Sort_grdColumns(grdICTPVCX1, "STYLE_CODE", False)

        With grdICTPVCX1.DisplayLayout.Bands(0)
            For Each COL_NAME As String In New String() {"STYLE_CODE", "COLORS"}
                .Columns(COL_NAME).Header.Fixed = True
            Next
        End With

        spl.Panel1Collapsed = True
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Refresh"
                Me.Cursor = Cursors.WaitCursor
                'EntryMode = "E"
                Load_Record()
                Mode_Settings(True)
                Me.Cursor = Cursors.Default
            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True
                '.Groups("Screen Control").Visible = False
            End With
        End If

        'Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdICTPVCX1.Visible = True

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"ICTPVCX1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        'If EntryMode = "N" Then
        '    'rowICTPLIN1.Item("STMT_TYPE") = HFs("STMT_TYPE")
        '    'rowICTPLIN1.Item("STMT_DESC") = HFs("STMT_DESC")
        'Else

        'End If

        Fill_Records("ICTPVCX1")
        Fill_Records("ICTSTATX")
        Fill_Extra_Fields()

        If EntryMode = "N" Then
        Else

        End If

    End Sub

    Private Sub Fill_Extra_Fields()
        For Each rowICTPVCX1 As DataRow In dst.Tables("ICTPVCX1").Select()
            Dim STYLE_CODE As String = rowICTPVCX1.Item("STYLE_CODE").ToString & String.Empty
            Dim COLORS As String = ""
            Dim AVAIL As Double = 0
            Dim xFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
            For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select(xFilter, "COLOR_CODE")
                If rowICTSTATX.Item("COLOR_CODE").ToString & String.Empty <> "" Then
                    COLORS = COLORS & rowICTSTATX.Item("COLOR_CODE").ToString & String.Empty & ","
                End If

                If IsNumeric(rowICTSTATX.Item("AVAIL").ToString & String.Empty) Then
                    AVAIL = AVAIL + Val(rowICTSTATX.Item("AVAIL").ToString & String.Empty)
                End If
            Next

            If COLORS.Length = 0 Then
                COLORS = "No Colors Found"
            Else
                COLORS = COLORS.Substring(0, COLORS.Length - 1)
            End If

            rowICTPVCX1.Item("COLORS") = COLORS
            rowICTPVCX1.Item("AVAIL") = AVAIL
            rowICTPVCX1.Item("BAR_CODE") = String.Format("*{0}*", STYLE_CODE)
            'PB1
            Dim rowARTCUST1 As DataRow = Nothing
            Dim Discounts As List(Of DISCOUNTS) = SOCMAIN2.Price_Discounts(Me, "", rowARTCUST1, STYLE_CODE, False)
            If Discounts(3).DISCOUNT_QTY = 0 Then
                rowICTPVCX1.Item("PB1") = Null
            Else
                rowICTPVCX1.Item("PB1") = Format(Discounts(3).DISCOUNT_PRICE, "###,##0.00")
            End If
            If Discounts(2).DISCOUNT_QTY = 0 Then
                rowICTPVCX1.Item("PB2") = Null
            Else
                rowICTPVCX1.Item("PB2") = Format(Discounts(2).DISCOUNT_PRICE, "###,##0.00")
            End If

            If Discounts(1).DISCOUNT_QTY = 0 Then
                rowICTPVCX1.Item("PB3") = Null
            Else
                rowICTPVCX1.Item("PB3") = Format(Discounts(1).DISCOUNT_PRICE, "###,##0.00")
            End If

            If Discounts(0).DISCOUNT_QTY = 0 Then
                rowICTPVCX1.Item("PB4") = Null
            Else
                rowICTPVCX1.Item("PB4") = Format(Discounts(0).DISCOUNT_PRICE, "###,##0.00")
            End If

        Next
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(
    ByVal ctl As Control,
    ByVal COLUMN_NAME As String,
    Optional ByRef sql_where As String = "",
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
        Load_Popup_Menu(grdICTPVCX1, "SS", "Show Filter", "Show GroupBox")
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
                Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
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

End Class