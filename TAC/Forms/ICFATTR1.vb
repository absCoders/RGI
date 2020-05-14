Public Class ICFATTR1
    Public STYLE_CODE As String
    Dim sqlcols As String = "ICTSTYL1.STYLE_CODE,ICTSTYL1.STYLE_STATUS,ICTSTYL1.STYLE_DESC,ICTSTYL1.INNER_PACK_QTY,ICTSTYL1.CARTON_PACK_QTY,ICTSTYL1.STYLE_UOM,ICTSTYL1.STYLE_PRICE,ICTSTYL1.STYLE_CLASS_CODE"
    Dim ATTR_CODEs As String

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from ICTATTR1"
            Create_TDA(.Tables.Add, "ICTATTR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select " & sqlcols & ", ICTCOLR1.COLOR_CODE,ICTATTR1.ATTR_CODE from ICTSTYL1,ICTCOLR1,ICTATTR1"
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 0)
        End With

        grdICTSTYL1.DataSource = dst.Tables("ICTSTYL1")
        Create_Summary(grdICTSTYL1, "STYLE_CODE", "Count")

        Fill_Records("ICTCLAS1")
        Fill_Records("ICTATTR1")
        Show_ATTR_CODEs(0)

        grdICTSTYL1.Visible = False
        Show_Filter(grdICTSTYL1, True)
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "STYLE_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Prepare_ICTCOLRM()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "STYLE_CODE"
            '    Prepare_ICTCOLRM()
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CLASS_CODE"
                Dim STYLE_CLASS_CODE As String = Absx1.txtFor("STYLE_CLASS_CODE").Text
                Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)
                If rowICTCLAS1 IsNot Nothing Then
                    If Show_ATTR_CODEs(1) = 0 Then
                        Find_Styles()
                    End If
                Else
                    Show_ATTR_CODEs(0)
                End If
        End Select
    End Sub

    Public Overrides Sub cbe_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.cbe_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ATTR_CODE_1", "ATTR_CODE_2", "ATTR_CODE_3"
                If Absx1.cbeFor(COLUMN_NAME).Tag & "" = "X" Then
                    ' DO NOTHING - SETTING VALUE FROM WITHIN Show_ATTR_CODEs
                Else
                    Dim I As Integer = Val(Mid(COLUMN_NAME, Len(COLUMN_NAME), 1))
                    Dim ATTR_CODE As String = Absx1.cbeFor(COLUMN_NAME).Value
                    Dim rowICTATTR1 As DataRow = dst.Tables("ICTATTR1").Rows.Find(ATTR_CODE)
                    If rowICTATTR1 IsNot Nothing Then
                        Dim Choices As Integer = Show_ATTR_CODEs(I + 1)
                        If Choices = 0 Or I = 3 Then
                            Find_Styles()
                        ElseIf Choices = 1 Then
                            Absx1.cbeFor("ATTR_CODE_" & Format(I + 1, "0")).Value = Absx1.cbeFor("ATTR_CODE_" & Format(I + 1, "0")).Items(0)
                        End If
                    Else
                        Show_ATTR_CODEs(I)
                    End If
                End If
        End Select
    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYL1, "SSSBBB", "Show Filter", "Show GroupBox", "Show Pins")
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

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.Name = "grd" Then
            Exit Sub
        End If

        Select Case e.SourceControl.Name
        
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
 
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
           
        End Select
    End Sub

#End Region

    Function Show_ATTR_CODEs(Attribute_Index As Integer) As Integer

        Dim Choices As Integer = -1

        Dim tbl As DataTable = Nothing

        If Attribute_Index >= 1 And Attribute_Index <= 3 Then
            Absx1.cbeFor("ATTR_CODE_" & Format(Attribute_Index, "0")).Tag = "X"
            Absx1.cbeFor("ATTR_CODE_" & Format(Attribute_Index, "0")).Value = ""
            Absx1.cbeFor("ATTR_CODE_" & Format(Attribute_Index, "0")).Tag = ""
            ASCMAIN1.sql = "Select Distinct ATTR_CODE from ICTSTYL3 " _
                & " where STYLE_CODE in (Select STYLE_CODE from ICTSTYL1 " _
                & " where STYLE_CLASS_CODE = '" & Absx1.txtFor("STYLE_CLASS_CODE").Text & "'" _
                & IIf(chkActiveOnly.Checked, " and STYLE_STATUS = 'A'", "") _
                & ")"

            If optAndOr.Value = "AND" Then
                If Attribute_Index >= 2 Then
                    ASCMAIN1.sql &= " and STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE = '" & Absx1.cbeFor("ATTR_CODE_1").Value & "')"
                    ASCMAIN1.sql &= " and ATTR_CODE <> '" & Absx1.cbeFor("ATTR_CODE_1").Value & "'"
                    If Attribute_Index >= 3 Then
                        ASCMAIN1.sql &= " and STYLE_CODE in (Select STYLE_CODE from ICTSTYL3 where ATTR_CODE = '" & Absx1.cbeFor("ATTR_CODE_2").Value & "')"
                        ASCMAIN1.sql &= " and ATTR_CODE <> '" & Absx1.cbeFor("ATTR_CODE_2").Value & "'"
                    End If
                End If
            End If

            ASCMAIN1.sql &= " order by ATTR_CODE"
            tbl = ASCDATA1.GetDataTable
            tbl.Columns.Add("SELECTED")
            Choices = tbl.Rows.Count
            If Attribute_Index = 1 Then lblATTR_CODE_1.Text = "Attribute 1 (" & CStr(tbl.Rows.Count) & " choices)"
            If Attribute_Index = 2 Then lblATTR_CODE_2.Text = "Attribute 2 (" & CStr(tbl.Rows.Count) & " choices)"
            If Attribute_Index = 3 Then lblATTR_CODE_3.Text = "Attribute 3 (" & CStr(tbl.Rows.Count) & " choices)"
            Absx1.cbeFor("ATTR_CODE_" & Format(Attribute_Index, "0")).ValueMember = "ATTR_CODE"
            Absx1.cbeFor("ATTR_CODE_" & Format(Attribute_Index, "0")).DisplayMember = "ATTR_CODE"
            Absx1.cbeFor("ATTR_CODE_" & Format(Attribute_Index, "0")).DataSource = tbl ' .Select("", "ATTR_CODE")
        End If

        lblATTR_CODE_1.Visible = (Attribute_Index >= 1) And (Attribute_Index > 1 Or Choices <> 0)
        lblATTR_CODE_2.Visible = (Attribute_Index >= 2) And (Attribute_Index > 2 Or Choices <> 0)
        lblATTR_CODE_3.Visible = (Attribute_Index >= 3) And (Attribute_Index > 3 Or Choices <> 0)

        cbeATTR_CODE_1.Visible = (Attribute_Index >= 1) And (Attribute_Index > 1 Or Choices <> 0)
        cbeATTR_CODE_2.Visible = (Attribute_Index >= 2) And (Attribute_Index > 2 Or Choices <> 0)
        cbeATTR_CODE_3.Visible = (Attribute_Index >= 3) And (Attribute_Index > 3 Or Choices <> 0)

        If Choices <> 0 And Choices <> 1 Then
            grdICTSTYL1.Visible = False
        End If

        If Choices > 1 Then
            tbl.TableName = "ICTATTR1"
            grdICTATTR1.DataSource = tbl
            If grdICTATTR1.DisplayLayout.Bands(0).Summaries.Count = 0 Then
                Create_Summary(grdICTATTR1, "ATTR_CODE", "Count")
                Create_Summary(grdICTATTR1, "SELECTED")
            End If
            grdICTATTR1.Text = "Attributes (" & CStr(Choices) & ")"
            splChoices.Panel1Collapsed = False
        Else
            splChoices.Panel1Collapsed = True
        End If
        Return Choices
    End Function

    Private Sub cmdDone_Click(sender As System.Object, e As System.EventArgs) Handles cmdDone.Click
        STYLE_CODE = ""
        Me.Close()
    End Sub

    Private Sub cmdFind_Click(sender As System.Object, e As System.EventArgs) Handles cmdFind.Click

        Dim STYLE_CLASS_CODE = Absx1.txtFor("STYLE_CLASS_CODE").Text
        Dim rowICTCLAS1 As DataRow = dst.Tables("ICTCLAS1").Rows.Find(STYLE_CLASS_CODE)
        If rowICTCLAS1 Is Nothing Then
            MsgBox("You Must Select (at minimum) a valid Class Code", MsgBoxStyle.OkOnly, "Cannot Find")
            Exit Sub
        End If

        Find_Styles()
    End Sub

    Sub Find_Styles()
        Dim STYLE_CLASS_CODE = Absx1.txtFor("STYLE_CLASS_CODE").Text
        ASCMAIN1.sql = "Select " & sqlcols & " from ICTSTYL1 where STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'"
        If chkActiveOnly.Checked Then
            ASCMAIN1.sql &= " and STYLE_STATUS = 'A'"
        End If
        Dim sqlw As String = ""
        For I As Integer = 1 To 3
            If Absx1.cbeFor("ATTR_CODE_" & Format(I, "0")).Visible Then
                Dim ATTR_CODE As String = Absx1.cbeFor("ATTR_CODE_" & Format(I, "0")).Value
                If ATTR_CODE <> "" Then
                    Dim rowICTATTR1 As DataRow = dst.Tables("ICTATTR1").Rows.Find(ATTR_CODE)
                    If rowICTATTR1 Is Nothing Then
                        MsgBox("Invalid Attribute Specified (" & ATTR_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Find")
                        Exit Sub
                    Else
                        If optAndOr.Value = "AND" Then
                            ASCMAIN1.sql &= " and STYLE_CODE in (Select Distinct STYLE_CODE from ICTSTYL3 where ATTR_CODE = '" & ATTR_CODE & "')"
                        Else
                            sqlw &= " or ATTR_CODE = '" & ATTR_CODE & "'"
                        End If
                    End If
                Else

                    ATTR_CODEs = ""
                    If grdICTATTR1.Visible Then
                        Dim tbl As DataTable = DirectCast(grdICTATTR1.DataSource, DataTable)
                        For Each rowICTATTR1 As DataRow In tbl.Select("SELECTED = '1'")
                            Dim ATTR_CODE_selected As String = rowICTATTR1.Item("ATTR_CODE")
                            ATTR_CODEs &= ",'" & ATTR_CODE_selected & "'"
                        Next
                        If ATTR_CODEs <> "" Then
                            ASCMAIN1.sql &= " and STYLE_CODE in (Select Distinct STYLE_CODE from ICTSTYL3 where ATTR_CODE in (" & Mid(ATTR_CODEs, 2) & "))"
                        End If
                    End If
                End If
            End If
        Next

        If sqlw <> "" And optAndOr.Value = "OR" Then
            ASCMAIN1.sql &= " and ICTSTYL1.STYLE_CODE in (Select Distinct STYLE_CODE from ICTSTYL3 where (" & Mid(sqlw, 4) & "))"
        End If

        ASCMAIN1.sql = "SELECT X.*, Y.ATTR_CODE, Z.COLOR_CODE, Z.OHMS, Z.POMS, Z.PSMS, Z.OHSW, Z.POSW, Z.PSSW  FROM (" _
            & ASCMAIN1.sql _
            & ") X," _
            & "(SELECT ICTSTYL3.STYLE_CODE, MIN (ATTR_CODE) ATTR_CODE from ICTSTYL3 " & IIf(ATTR_CODEs = "", "", " where ATTR_CODE in (" & Mid(ATTR_CODEs, 2) & ")") & " group by ICTSTYL3.STYLE_CODE) Y," _
            & "(SELECT STYLE_CODE, COLOR_CODE" _
            & ", SUM (DECODE(WHSE_CODE,'MS',WHSE_QTY_ON_HAND,0)) OHMS, SUM (DECODE(WHSE_CODE,'MS',WHSE_QTY_ON_ORDER,0)) POMS, SUM (DECODE(WHSE_CODE,'MS',WHSE_QTY_TRAN,0)) PSMS" _
            & ", SUM (DECODE(WHSE_CODE,'SW',WHSE_QTY_ON_HAND,0)) OHSW, SUM (DECODE(WHSE_CODE,'SW',WHSE_QTY_ON_ORDER,0)) POSW, SUM (DECODE(WHSE_CODE,'SW',WHSE_QTY_TRAN,0)) PSSW" _
            & " from ICTSTAT2 GROUP BY STYLE_CODE, COLOR_CODE) Z" _
            & " where Y.STYLE_CODE (+) = X.STYLE_CODE" _
            & " and Z.STYLE_CODE (+) = X.STYLE_CODE"

        grdICTSTYL1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        Fill_Records("ICTSTYL1", "", True, ASCMAIN1.sql)

        'With dst.Tables("ICTSTYL1")

        'End With
        With grdICTSTYL1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"OHMS", "POMS", "PSMS", "OHSW", "POSW", "PSSW"}
                .Columns(COLUMN_NAME).Format = "#,##0"
                .Columns(COLUMN_NAME).Width = 80
            Next
        End With
        ASCMAIN1.grdInitializeLayout(grdICTSTYL1)
        Sort_grdColumns(grdICTSTYL1, "STYLE_CODE")
        ATTR_CODEs = ""
        grdICTSTYL1.Visible = True
    End Sub

    Private Sub grdICTSTYL1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSTYL1.DoubleClickRow
        If e.Row.IsDataRow And Not e.Row.IsAddRow Then
            STYLE_CODE = e.Row.Cells("STYLE_CODE").Value
            Me.Close()
        End If
    End Sub

    Private Sub chkActiveOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkActiveOnly.CheckedChanged
        If lblATTR_CODE_1.Visible Then
            If Show_ATTR_CODEs(1) = 0 Then
                Find_Styles()
            Else
                Absx1.cbeFor("ATTR_CODE_1").Value = ""
            End If
        End If
    End Sub

    Private Sub optAndOr_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optAndOr.ValueChanged
        If lblATTR_CODE_1.Visible Then
            If Show_ATTR_CODEs(1) = 0 Then
                Find_Styles()
            Else
                Absx1.cbeFor("ATTR_CODE_1").Value = ""
            End If
        End If
    End Sub
End Class