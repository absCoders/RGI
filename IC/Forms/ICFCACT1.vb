Imports Infragistics.Win.UltraWinGrid

Public Class ICFCACT1

    Private wkTable As String = String.Empty
    Private selectedRowCount As Integer = 0
    Private selectedStylesList As New HashSet(Of String)
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fromDate As String = New DateTime(DateTime.Today.Year - 1, 1, 1).ToString("dd-MMM-yyyy").ToUpper()
        With dst

            ASCMAIN1.sql = $"select * from (
                                select ICTIREC1.OPS_YYYYPP PERIOD, 'Received' ACTIVITY, POTORDR1.PO_REFERENCE, ICTIREC2.STYLE_CODE,  sum(nvl(ICTIREC2.QTY_REC,0)) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                union
                                Select SOTINVH2.ORDR_YYYYPP_UPDATED, 'Shipped', ' ', SOTINVH2.STYLE_CODE, SUM(nvl(SOTINVH2.ORDR_QTY_SHIP,0) * -1)
                                from SOTINVH1, SOTINVH2, SOTORDR1
                                where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE
                                and SOTINVH1.INV_NO = SOTINVH2.INV_NO
                                and SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO
                                and SOTINVH2.STYLE_CODE in('')
                                group by SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CODE
                                union
                                select to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), 'Adjusted', ' ', ICTIADJ2.STYLE_CODE, sum(ICTIADJ2.ADJ_QTY)
                                from ICTIADJ1, ICTIADJ2
                                Where ICTIADJ1.ADJ_NO =  ICTIADJ2.ADJ_NO
                                AND ICTIADJ1.REVERSED_BY_ADJ_NO IS NULL
                                AND ICTIADJ1.REVERSES_ADJ_NO IS NULL
                                AND ICTIADJ2.STYLE_CODE in ('')
                                group by to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), ICTIADJ2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in(''))"
            Create_TDA(.Tables.Add, "ICTCACTX", "**", 0, False)

            ASCMAIN1.sql = $"select * from (
                                select ICTIREC1.OPS_YYYYPP PERIOD, 't_Bal' ACTIVITY, POTORDR1.PO_REFERENCE, ICTIREC2.STYLE_CODE, sum(nvl(ICTIREC2.QTY_REC,0)) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in(''))"
            Create_TDA(.Tables.Add, "ICTCACTB", "**", 0, False)
        End With
        grdICTCACTX.DataSource = dst.Tables("ICTCACTX")
        Absx1.txtFor("CUST_CODE").Value = "WALMART"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")
                If EMsg.Length = 0 Then

                End If

            Case "Generate"

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
                Clear_Record()
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
        cmdMulti.Visible = ScreenMode
        grdICTCACTX.Visible = False

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"ICTCACTX", "ICTCACTB"}
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
                                select ICTIREC1.OPS_YYYYPP PERIOD, 'Received' ACTIVITY, POTORDR1.PO_REFERENCE, ICTIREC2.STYLE_CODE,  sum(nvl(ICTIREC2.QTY_REC,0)) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('{String.Join("','", selectedStylesList)}')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                union
                                Select SOTINVH2.ORDR_YYYYPP_UPDATED, 'Shipped', ' ',  SOTINVH2.STYLE_CODE, SUM(nvl(SOTINVH2.ORDR_QTY_SHIP,0) * -1)
                                from SOTINVH1, SOTINVH2, SOTORDR1
                                where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE
                                and SOTINVH1.INV_NO = SOTINVH2.INV_NO
                                and SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO
                                and SOTINVH2.STYLE_CODE in('{String.Join("','", selectedStylesList)}')
                                group by SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CODE
                                union
                                select to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), 'Adjusted', ' ', ICTIADJ2.STYLE_CODE, sum(ICTIADJ2.ADJ_QTY)
                                from ICTIADJ1, ICTIADJ2
                                Where ICTIADJ1.ADJ_NO =  ICTIADJ2.ADJ_NO
                                AND ICTIADJ1.REVERSED_BY_ADJ_NO IS NULL
                                AND ICTIADJ1.REVERSES_ADJ_NO IS NULL
                                AND ICTIADJ2.STYLE_CODE in ('{String.Join("','", selectedStylesList)}')
                                group by to_CHAR(ICTIADJ1.ADJ_DATE,'YYYYmm'), ICTIADJ2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in('{String.Join("','", selectedStylesList)}'))"

        grdICTCACTX.DisplayLayout.NewColumnLoadStyle = NewColumnLoadStyle.Show
        grdICTCACTX.DisplayLayout.NewBandLoadStyle = NewBandLoadStyle.Show
        Fill_Records("ICTCACTX", "", True, ASCMAIN1.sql)
        grdICTCACTX.DataSource = dst.Tables("ICTCACTX")
        AdjustTableSchema("ICTCACTX", selectedStylesList)
        ASCMAIN1.grdInitializeLayout(grdICTCACTX)
        For Each col As UltraGridColumn In grdICTCACTX.DisplayLayout.Bands(0).Columns
            If Not {"PERIOD", "ACTIVITY", "PO_REFERENCE"}.Contains(col.Key) Then
                col.Header.Caption = col.Key.Trim("'"c)
            End If
        Next

        ASCMAIN1.sql = $"select * from (
                                select ICTIREC1.OPS_YYYYPP PERIOD, 't_Bal' ACTIVITY, POTORDR1.PO_REFERENCE, ICTIREC2.STYLE_CODE,  sum(nvl(ICTIREC2.QTY_REC,0) * 0) PO_QTY_REC
                                from POTORDR1, ICTIREC1, ICTIREC2
                                where  POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO
                                and ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO
                                and ICTIREC2.STYLE_CODE in('{String.Join("','", selectedStylesList)}')
                                group by POTORDR1.PO_REFERENCE, ICTIREC1.OPS_YYYYPP, ICTIREC2.STYLE_CODE
                                ) pivot (sum (PO_QTY_REC) for STYLE_CODE in('{String.Join("','", selectedStylesList)}'))"

        Fill_Records("ICTCACTB", "", True, ASCMAIN1.sql)

        'add a row after every shipment record
        Dim SkipCols As String = "PERIOD,ACTIVITY,PO_REFERENCE"

        For Each row As DataRow In dst.Tables("ICTCACTX").Select("ACTIVITY = 'Shipped' or ACTIVITY = 'Adjusted' or ACTIVITY = 'Received'", "PERIOD,ACTIVITY,PO_REFERENCE")
            Dim rowdist As DataRow = dst.Tables("ICTCACTX").NewRow()
            Dim forcerow As Boolean = False
            Dim balrowis0 As Boolean = True
            rowdist.ItemArray = row.ItemArray
forcerow_here:
            If rowdist("PERIOD") = "201901" And rowdist("ACTIVITY") = "Received" Then
                balrowis0 = False
            End If
            For Each balrow As DataRow In dst.Tables("ICTCACTB").Select($"PERIOD is null or PERIOD <= '{rowdist("PERIOD")}'")
                Dim updaterow As Boolean = False
                If rowdist("ACTIVITY") = "Received" And rowdist("PO_REFERENCE") <> balrow("PO_REFERENCE") Then
                    Continue For
                End If
                For Each col As DataColumn In dst.Tables("ICTCACTB").Columns
                    If Not SkipCols.Contains(col.ColumnName) Then
                        If (Val(balrow(col.ColumnName) & "") > 0 Or Val(rowdist(col.ColumnName) & "") <> 0) Then

                            updaterow = True
                            balrow("PERIOD") = row("PERIOD")
                            If Val(balrow(col.ColumnName) & "") > (Val(rowdist(col.ColumnName) & "") * -1) Or forcerow Then
                                balrow(col.ColumnName) = Val(balrow(col.ColumnName) & "") + Val(rowdist(col.ColumnName) & "")
                                rowdist(col.ColumnName) = 0
                                balrowis0 = False
                            Else
                                rowdist(col.ColumnName) = Val(rowdist(col.ColumnName) & "") + Val(balrow(col.ColumnName) & "")
                                balrow(col.ColumnName) = 0
                            End If
                        End If
                    End If
                Next
                If updaterow Then
                    Dim updated As Boolean = False
                    For Each balupdate As DataRow In dst.Tables("ICTCACTX").Select($"PO_REFERENCE = '{balrow("PO_REFERENCE")}' and ACTIVITY = '{balrow("ACTIVITY")}' and PERIOD = '{balrow("PERIOD")}'")
                        updated = True
                        For Each col As DataColumn In dst.Tables("ICTCACTB").Columns
                            If Not SkipCols.Contains(col.ColumnName) Then
                                If balrow(col.ColumnName) & "" = "" Then
                                    'Stop
                                Else
                                    balupdate(col.ColumnName) = Val(balrow(col.ColumnName) & "")
                                End If
                            End If
                        Next
                    Next
                    If Not updated Then
                        dst.Tables("ICTCACTX").ImportRow(balrow)
                    End If
                End If
            Next
            For Each col As DataColumn In dst.Tables("ICTCACTB").Columns
                If Not SkipCols.Contains(col.ColumnName) Then
                    If (Val(rowdist(col.ColumnName) & "") <> 0) Then
                        forcerow = True
                        GoTo forcerow_here
                        Stop
                    End If
                End If
            Next
        Next

        Sort_grdColumns(grdICTCACTX, "PERIOD, ACTIVITY, PO_REFERENCE")
        grdICTCACTX.Visible = True

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
    Private Sub UpdateSelectedStylesTextBox()
        Dim stylesText As String = String.Join(", ", selectedStylesList)
        Dim heading As String = "Selected Styles: " & vbCrLf
        SelectedStyles.Text = heading & stylesText
    End Sub
    Private Sub cmdMulti_Click(sender As System.Object, e As System.EventArgs) Handles cmdMulti.Click
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STYLE_CODE")
        ASCMAIN1.CodeSelector.SQL &= " where CUST_CODE ='" & Absx1.txtFor("CUST_CODE").Text & "'"
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                selectedStylesList.Clear()
                For Each STYLE_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    selectedStylesList.Add(STYLE_CODE)
                Next
            End If
            UpdateSelectedStylesTextBox()
        End If
    End Sub
    Private Sub AdjustTableSchema(ByRef table As String, ByVal selectedStyles As HashSet(Of String))
        Dim baseColumns As New List(Of String) From {"PO_REFERENCE", "ACTIVITY", "PERIOD"} ' Add all your base column names here

        Dim adjustedSelectedStyles As New HashSet(Of String)(selectedStyles.Select(Function(s) s.Replace("'", "")))
        For Each baseColumn As String In baseColumns
            adjustedSelectedStyles.Add(baseColumn)
        Next

        With grdICTCACTX.DisplayLayout.Bands(0)
            For Each grdCol As UltraGridColumn In .Columns
                Dim cleanKey As String = grdCol.Key.Trim("'")
                If adjustedSelectedStyles.Contains(cleanKey) OrElse baseColumns.Contains(cleanKey) Then
                    grdCol.Hidden = False
                Else
                    grdCol.Hidden = True
                End If
            Next
        End With
    End Sub

End Class