Imports Infragistics.Win.UltraWinGrid

Public Class ICFCACT1

    Private wkTable As String = String.Empty
    Private selectedRowCount As Integer = 0
    Private selectedStylesList As New HashSet(Of String)
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim fromDate As String = New DateTime(DateTime.Today.Year - 1, 1, 1).ToString("dd-MMM-yyyy").ToUpper()
        With dst
            ASCMAIN1.sql = $"SELECT DISTINCT POTORDR1.PO_ORDER_NO PO_ORDER_NO,
                            POTORDR1.PO_DATE_ORDERED PO_DATE_ORDERED,
                            POTORDR1.PO_REFERENCE PO_REFERENCE,
                            POTORDR1.TERM_CODE TERM_CODE,
                            POTORDR1.PO_STATUS PO_STATUS,
                            ICTIREC1.RECEIPT_NO RECEIPT_NO,
                            ICTIREC1.RECEIPT_DATE RECEIPT_DATE,
                            ICTIREC1.WHSE_CODE WHSE_CODE_5,
                            ICTIREC1.QTY_REC QTY_REC,
                            ICTIREC1.AMT_REC AMT_REC,
                            ICTIREC1.QTY_INV QTY_INV,
                            ICTIREC1.AMT_INV AMT_INV FROM POTORDR1, ICTIREC1, ICTIREC2
                            WHERE POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                            and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                            AND POTORDR1.PO_DATE_ORDERED >= :PARM1"
            Create_TDA(.Tables.Add, "ICTCACT1", "**", 0, False, "D")

            ASCMAIN1.sql = $"select *
                            FROM ICTIREC2
                            where ICTIREC2.PO_ORDER_NO = :PARM1
                            and ICTIREC2.RECEIPT_NO = :PARM2"
            Create_TDA(.Tables.Add, "ICTCACT2", "**", 0, False, "VV")

            With .Tables("ICTCACT2")
                .Columns.Add("SEL") ' Specify the data type as Integer for the 0 or 1 values
                .Columns("SEL").DefaultValue = "0"
                .Columns("SEL").Caption = "Sel"
                .Columns("SEL").SetOrdinal(0)
            End With

            ASCMAIN1.sql = $"select * from (
                                select POTORDR1.PO_REFERENCE, 'Received' ACTIVITY, ICTIREC1.OPS_YYYYPP Period, ICTIREC2.STYLE_CODE,  sum(nvl(ICTIREC2.QTY_REC,0)) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                union
                                Select ' ', 'Shipped',  SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CODE, SUM(nvl(SOTINVH2.ORDR_QTY_SHIP,0) * -1)
                                from SOTINVH1, SOTINVH2, SOTORDR1
                                where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE
                                and SOTINVH1.INV_NO = SOTINVH2.INV_NO
                                and SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO
                                and SOTINVH2.STYLE_CODE in('')
                                group by SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CODE
                                union
                                select ' ','Adjusted',to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), ICTIADJ2.STYLE_CODE, sum(ICTIADJ2.ADJ_QTY)
                                from ICTIADJ1, ICTIADJ2
                                Where ICTIADJ1.ADJ_NO =  ICTIADJ2.ADJ_NO
                                AND ICTIADJ1.REVERSED_BY_ADJ_NO IS NULL
                                AND ICTIADJ1.REVERSES_ADJ_NO IS NULL
                                AND ICTIADJ2.STYLE_CODE in ('')
                                group by to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), ICTIADJ2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in(''))"
            Create_TDA(.Tables.Add, "ICTCACTX", "**", 0, False)

            ASCMAIN1.sql = $"select * from (
                                select POTORDR1.PO_REFERENCE, 't Balance' ACTIVITY, ICTIREC1.OPS_YYYYPP Period, ICTIREC2.STYLE_CODE,  sum(nvl(ICTIREC2.QTY_REC,0)) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in(''))"
            Create_TDA(.Tables.Add, "ICTCACTB", "**", 0, False)
        End With

        grdICFCACT1.DataSource = dst.Tables("ICTCACT1")
        grdICFCACT2.DataSource = dst.Tables("ICTCACT2")
        grdICFCACTX.DataSource = dst.Tables("ICTCACTX")


        For Each gcol As UltraWinGrid.UltraGridColumn In grdICFCACT2.DisplayLayout.Bands(0).Columns
            gcol.CellActivation = If(gcol.Key = "SEL", UltraWinGrid.Activation.AllowEdit, UltraWinGrid.Activation.NoEdit)
        Next
        grdICFCACT2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        Create_Summary(grdICFCACT1, "PO_ORDER_NO", "Count")
        Create_Summary(grdICFCACT2, "STYLE_CODE", "Count")


        dteSearchS.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dteSearchS.MinDate = New DateTime(DateTime.Today.Year - 5, 1, 1).ToString("dd-MMM-yyyy").ToUpper()
        dteSearchS.DateTime = New DateTime(DateTime.Today.Year - 1, 1, 1).ToString("dd-MMM-yyyy").ToUpper()

        Absx1.txtFor("CUST_CODE").Value = "WALMART"



    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")
                If EMsg.Length = 0 Then

                End If

            'Case "Generate"
            '    If dst.Tables("EDTSYSIH").Rows.Count = 0 OrElse dst.Tables("EDT846O1").Rows.Count = 0 OrElse dst.Tables("EDT846O2").Rows.Count = 0 Then
            '        EMsg &= vbCr & "There is not EDI 846 to Update."
            '    Else
            '        If MessageBox.Show($"Do you want to Update the EDI 846 for Ecom Code: {Absx1.txtFor("ECOM_CODE").Text}?") = DialogResult.No Then
            '            Exit Sub
            '        End If
            '    End If

            Case "Cancel"
                selectedStylesList.Clear()
                SelectedStyles.Text = String.Empty

            Case "Refresh"
                If MessageBox.Show($"Do you want to Refresh?") = DialogResult.No Then
                    Exit Sub
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Generate"
                Generate()
            '    Mode_Settings(False)

            Case "Refresh"
                'Refresh_Entries()

        End Select

    End Sub
    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Generate").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = not_iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"ICTCACT1", "ICTCACT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        ' Refresh_Entries()

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating Detail History")

        Try
            EnforceConstraints(False)

            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
            Dim Emsg As String = String.Empty


            Fill_Records("ICTCACT1", New Object() {dteSearchS.DateTime}, True)


            EnforceConstraints(True)

        Catch ex As Exception
            MessageBox.Show($"Error Generating Detail Data {ex.Message }", "Generate Detail Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Clear_Record()
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try
    End Sub

    Private Sub Generate()

        ASCMAIN1.sql = $"select * from (
                                select POTORDR1.PO_REFERENCE, 'Received' ACTIVITY, ICTIREC1.OPS_YYYYPP Period, ICTIREC2.STYLE_CODE,  sum(nvl(ICTIREC2.QTY_REC,0)) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('{String.Join("','", selectedStylesList)}')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                union
                                Select ' ', 'Shipped',  SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CODE, SUM(nvl(SOTINVH2.ORDR_QTY_SHIP,0) * -1)
                                from SOTINVH1, SOTINVH2, SOTORDR1
                                where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE
                                and SOTINVH1.INV_NO = SOTINVH2.INV_NO
                                and SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO
                                and SOTINVH2.STYLE_CODE in('{String.Join("','", selectedStylesList)}')
                                group by SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CODE
                                union
                                select ' ','Adjusted',to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), ICTIADJ2.STYLE_CODE, sum(ICTIADJ2.ADJ_QTY)
                                from ICTIADJ1, ICTIADJ2
                                Where ICTIADJ1.ADJ_NO =  ICTIADJ2.ADJ_NO
                                AND ICTIADJ1.REVERSED_BY_ADJ_NO IS NULL
                                AND ICTIADJ1.REVERSES_ADJ_NO IS NULL
                                AND ICTIADJ2.STYLE_CODE in ('{String.Join("','", selectedStylesList)}')
                                group by to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), ICTIADJ2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in('{String.Join("','", selectedStylesList)}'))"

        Fill_Records("ICTCACTX", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = $"select * from (
                                select POTORDR1.PO_REFERENCE, 'T Bal' ACTIVITY, ICTIREC1.OPS_YYYYPP Period, ICTIREC2.STYLE_CODE,  sum(nvl(ICTIREC2.QTY_REC,0)) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('{String.Join("','", selectedStylesList)}')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in('{String.Join("','", selectedStylesList)}'))"

        Fill_Records("ICTCACTB", "", True, ASCMAIN1.sql)

        'add a row after every shipment record
        Dim SkipCols As String = "PERIOD,ACTIVITY,PO_REFERENCE"

        For Each row As DataRow In dst.Tables("ICTCACTX").Select("ACTIVITY = 'Shipped' or ACTIVITY = 'Adjusted'", "PERIOD,ACTIVITY,PO_REFERENCE")
            Dim rowdist As DataRow = dst.Tables("ICTCACTX").NewRow()
            rowdist.ItemArray = row.ItemArray
            For Each balrow As DataRow In dst.Tables("ICTCACTB").Select("")
                Dim updaterow As Boolean = False
                For Each col As DataColumn In dst.Tables("ICTCACTB").Columns
                    If Not SkipCols.Contains(col.ColumnName) Then
                        If Val(balrow(col.ColumnName) & "") > 0 And Val(balrow(col.ColumnName) & "") > 0 Then
                            updaterow = True
                            balrow("PERIOD") = row("PERIOD")
                            If Val(balrow(col.ColumnName) & "") > (Val(rowdist(col.ColumnName) & "") * -1) Then
                                balrow(col.ColumnName) = Val(balrow(col.ColumnName) & "") + Val(rowdist(col.ColumnName) & "")
                                rowdist(col.ColumnName) = 0
                            Else
                                rowdist(col.ColumnName) = Val(rowdist(col.ColumnName) & "") + Val(balrow(col.ColumnName) & "")
                                balrow(col.ColumnName) = 0
                            End If
                        End If
                    End If
                Next
                If updaterow Then
                    dst.Tables("ICTCACTX").ImportRow(balrow)
                End If
            Next
        Next




    End Sub
    Private Sub Update_Record()

        'Try
        '    BeginTrans()

        '    Update_Record_TDA("EDTSYSIH")
        '    Update_Record_TDA("EDT846O1")
        '    Update_Record_TDA("EDT846O2")

        '    CommitTrans("EDI 846 Updated")

        'Catch ex As Exception
        '    Rollback($"Error Updating EDI 846: {ex.Message }")
        'End Try

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

    Private Sub Refresh_Entries()

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

        For Each TABLE_NAME As String In New String() {"ICTCACT1", "ICTCACT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        'Try
        '    ASCDATA1.ExecuteSQL($"DELETE FROM {wkTable}")
        '    ASCMAIN1.sql = $"SELECT EDI_OUTBOUND_DOC_NO 
        '                        FROM EDTSYSIH 
        '                        where TRUNC(INIT_DATE) BETWEEN '{dteSearchS.DateTime.ToString("dd-MMM-yyyy")}' AND '{dteSearchE.DateTime.ToString("dd-MMM-yyyy")}' 
        '                        and EDI_APPLICATION_ID = 'IB'"

        '    If CUST_CODE.Length > 0 Then
        '        Dim rowECTECOM1 As DataRow = dst.Tables("ECTECOM1").Rows.Find(CUST_CODE)
        '        Dim EDI_OUR_ID As String = String.Empty

        '        Dim EDI_TP_ID As String = String.Empty
        '        Dim EDI_TP_QUAL As String = String.Empty

        '        If rowECTECOM1 IsNot Nothing Then
        '            EDI_TP_ID = rowECTECOM1.Item("EDI_TP_ID") & String.Empty
        '            EDI_TP_QUAL = rowECTECOM1.Item("EDI_TP_QUAL") & String.Empty

        '            Dim Sql As String = $"EDI_TP_QUAL = '{EDI_TP_QUAL}' and EDI_TP_ID = '{EDI_TP_ID}'"

        '            If dst.Tables("EDTTRPM1").Select(Sql).Length > 0 Then
        '                Dim rowEDTTRPM1 As DataRow = dst.Tables("EDTTRPM1").Select(Sql)(0)
        '                EDI_OUR_ID = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
        '            End If

        '        End If

        '        If EDI_OUR_ID.Length > 0 AndAlso EDI_TP_ID.Length > 0 Then
        '            ASCMAIN1.sql &= $" and EDI_OUR_ID = '{EDI_OUR_ID}' and EDI_TP_ID = '{EDI_TP_ID}' "
        '        End If
        '    End If

        '    ASCDATA1.ExecuteSQL($"INSERT INTO {wkTable}  {ASCMAIN1.sql}")

        '    Fill_Records("EDT846O1", String.Empty, True, $"SELECT EDTSYSIH.EDI_TP_ID, EDT846O1.* FROM EDT846O1, EDTSYSIH WHERE EDT846O1.EDI_OUTBOUND_DOC_NO IN (SELECT EDI_OUTBOUND_DOC_NO FROM {wkTable}) AND EDT846O1.EDI_OUTBOUND_DOC_NO = EDTSYSIH.EDI_OUTBOUND_DOC_NO (+)")

        'Catch ex As Exception
        '    MessageBox.Show($"Error getting EDI 846s: {ex.Message }", "Load EDI 846", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    For Each TABLE_NAME As String In New String() {"EDTSYSIH", "EDT846O1", "EDT846O2"}
        '        dst.Tables(TABLE_NAME).Rows.Clear()
        '    Next
        'End Try

    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICFCACT1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdICFCACT2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("EDI_STYLE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

    Private Sub grdICFCACT1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICFCACT1.AfterRowActivate

        If grdICFCACT1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        Dim PO_ORDER_NO As String = grdICFCACT1.ActiveRow.Cells("PO_ORDER_NO").Value & String.Empty
        Dim RECEIPT_NO As String = grdICFCACT1.ActiveRow.Cells("RECEIPT_NO").Value & String.Empty

        ' LOAD THEM ONLY IF THEY AERE NEEDED. TAKES TO LONG
        Me.Cursor = Cursors.WaitCursor
        Fill_Records("ICTCACT2", New Object() {PO_ORDER_NO, RECEIPT_NO}, True, String.Empty)
        ' Now ensure the previously selected styles are checked in the grid
        CheckSelectedStylesInGrid()
        Me.Cursor = Cursors.Default

        'Dim dView As New DataView(dst.Tables("ICTCACT2"))
        'dView.RowFilter = $"PO_ORDER_NO = '{PO_ORDER_NO}'"
        'grdICFCACT2.DataSource = dView

    End Sub

    Private Sub grdICFCACT2_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdICFCACT2.ClickCell
        If e.Cell.Column.Key = "SEL" AndAlso e.Cell.IsDataCell Then
            e.Cell.Value = Not CBool(e.Cell.Value)
            Dim styleCode As String = e.Cell.Row.Cells("STYLE_CODE").Value.ToString()

            If CBool(e.Cell.Value) Then
                ' Add the style code to the HashSet if the box is checked
                selectedStylesList.Add(styleCode)
            Else
                ' Remove the style code from the HashSet if the box is unchecked
                selectedStylesList.Remove(styleCode)
            End If
            ' Update the SelectedStyles TextBox
            UpdateSelectedStylesTextBox()
        End If
    End Sub
    Private Sub grdICFCACT2_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdICFCACT2.AfterCellUpdate
        ' Check if the updated cell is in the "SEL" column and is a data cell
        If e.Cell.Column.Key = "SEL" AndAlso e.Cell.IsDataCell Then
            ' Manually update the row selection state based on the "SEL" cell value
            e.Cell.Row.Selected = CBool(e.Cell.Value)
        End If
    End Sub
    Private Sub UpdateSelectedStylesTextBox()
        Dim stylesText As String = String.Join(", ", selectedStylesList)
        Dim heading As String = "Selected Styles: " & vbCrLf
        SelectedStyles.Text = heading & stylesText
    End Sub
    Private Sub CheckSelectedStylesInGrid()
        If grdICFCACT2 IsNot Nothing AndAlso grdICFCACT2.Rows.Count > 0 Then
            grdICFCACT2.SuspendLayout()
            For Each row As DataRow In dst.Tables("ICTCACT2").Select()
                Dim styleCode As String = row("STYLE_CODE") & ""

                ' Check if the current row's style code is in the selectedStylesList
                If selectedStylesList.Contains(styleCode) Then
                    ' Check the "SEL" cell without raising the ClickCell event to avoid duplicate handling
                    row("SEL") = "1"
                End If
            Next
            grdICFCACT2.ResumeLayout(True)
        End If
    End Sub

    Private Sub grdICFCACT2_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdICFCACT2.BeforeCellUpdate

    End Sub

    'Private Sub grdICFCACT2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICFCACT2.InitializeRow
    '    Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & ""
    '    If STYLE_CODE <> "" AndAlso selectedStylesList.Contains("," + STYLE_CODE) Then
    '        e.Row.Cells("SEL").Value = "1"
    '    End If
    'End Sub
End Class