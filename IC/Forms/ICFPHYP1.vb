Imports Infragistics.Win.UltraWinGrid

Public Class ICFPHYP1
    Dim RYP As String
    Dim ICTPHYJW As String = ""
    Dim ICTPHYJ1 As String = ""
    Dim ICTPHYJ2 As String = ""


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ' CREATE TEMP TABLES ======

            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE from SOTORDR2 where ROWNUM < 1"
            ICTPHYJW = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTPHYJW & " add Primary Key (STYLE_CODE, COLOR_CODE)")

            ASCMAIN1.sql = "Select WHTLOCB1.* From WHTLOCB1 where ROWNUM < 1"
            ICTPHYJ1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("CREATE INDEX I_" & ICTPHYJ1 & "_1 ON " & ICTPHYJ1 & " (STYLE_CODE,COLOR_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTPHYJ1 & " ADD YELLOW VARCHAR2(1)")

            ASCMAIN1.sql = "Select LOCATION_CODE, SUM (LOCATION_QTY) QTY" & vbCrLf _
               & ", SUM (Case When YELLOW = '1' THEN LOCATION_QTY ELSE 0 END) YELLOW" & vbCrLf _
               & ", SUM (CASE WHEN YELLOW = '1' THEN 0 ELSE LOCATION_QTY END) NOTYELLOW" & vbCrLf _
               & " From " & ICTPHYJ1 & " Group By LOCATION_CODE" & vbCrLf
            ICTPHYJ2 = ASCMAIN1.Temp_Table

            ' CREATE TEMP TABLES ======

            ASCMAIN1.sql = "Select 'NO YELLOWS' LOCATION_STATUS, COUNT (*) PALLETS, SUM (YELLOW) YELLOW, SUM (NOTYELLOW) NOTYELLOW" & vbCrLf _
                & " , MIN (LOCATION_CODE) MINLOC, MAX (LOCATION_CODE) MAXLOC" & vbCrLf _
                & " From " & ICTPHYJ2 & " Where YELLOW = 0" & vbCrLf _
                & " UNION " & vbCrLf _
                & " Select 'ALL YELLOW' LOCATION_STATUS, COUNT (*) PALLETS, SUM (YELLOW) YELLOW, SUM (NOTYELLOW) NOTYELLOW" & vbCrLf _
                & " , MIN (LOCATION_CODE) MINLOC, MAX (LOCATION_CODE) MAXLOC" & vbCrLf _
                & " From " & ICTPHYJ2 & " Where NOTYELLOW = 0  And YELLOW <> 0" & vbCrLf _
                & " UNION" & vbCrLf _
                & " Select 'SOME YELLOW' LOCATION_STATUS, COUNT (*) PALLETS, SUM (YELLOW) YELLOW, SUM (NOTYELLOW) NOTYELLOW" & vbCrLf _
                & " , MIN (LOCATION_CODE) MINLOC, MAX (LOCATION_CODE) MAXLOC" & vbCrLf _
                & " From " & ICTPHYJ2 & "" & vbCrLf _
                & " Where NOTYELLOW <> 0 And YELLOW <> 0"
            Create_TDA(.Tables.Add, "ICTPHYJ2", "**", 0, False, "")

            ASCMAIN1.sql = "Select LOCATION_CODE,SUM(LOCATION_QTY) YELLOW FROM " & ICTPHYJ1 & " Where YELLOW = '1' AND LOCATION_CODe In " & vbCrLf _
            & " (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW = 0 And YELLOW <> 0) group by LOCATION_CODE"
            Create_TDA(.Tables.Add, "ICTPHYJYS", "**", 0, False, "")

            ASCMAIN1.sql = "Select * From " & ICTPHYJ1 & " Where LOCATION_CODE IN " & vbCrLf _
            & "(Select LOCATION_CODE FROM " & ICTPHYJ2 & "  WHERE NOTYELLOW = 0 And YELLOW <> 0)"
            Create_TDA(.Tables.Add, "ICTPHYJY", "**", 0, False, "")

            'ASCMAIN1.sql = "Select  LOCATION_CODE,SUM(YELLOW) LOCATION_QTY_YELLOW,SUM(NOTYELLOW) LOCATION_QTY_NOTYELLOW FROM (" & vbCrLf _
            '    & " Select LOCATION_CODE,SUM(LOCATION_QTY) YELLOW,0 NOTYELLOW FROM " & ICTPHYJ1 & " Where YELLOW = '1' AND LOCATION_CODE" & vbCrLf _
            '    & " In (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW <> 0 And YELLOW <> 0)" & vbCrLf _
            '    & " GROUP by LOCATION_CODE" & vbCrLf _
            '    & " UNION" & vbCrLf _
            '    & " Select LOCATION_CODE,0 YELLOW,SUM(LOCATION_QTY) NOTYELLOW FROM " & ICTPHYJ1 & " Where YELLOW Is NULL And LOCATION_CODE" & vbCrLf _
            '    & " In (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW <> 0 And YELLOW <> 0)" & vbCrLf _
            '    & " GROUP by LOCATION_CODE)" & vbCrLf _
            '    & " GROUP by LOCATION_CODE"
            'Create_TDA(.Tables.Add, "ICTPHYJMS", "**", 0, False, "")

            ASCMAIN1.sql = "Select LOCATION_CODE,SUM(YELLOW) LOCATION_QTY_YELLOW, SUM(NOTYELLOW) LOCATION_QTY_NOTYELLOW,  sum(STYLES) STYLES ,  sum(yellow_styles) STYLES_Y,  sum(not_yellow_s) STYLES_N FROM (" & vbCrLf _
             & "Select LOCATION_CODE,SUM(LOCATION_QTY) YELLOW,0 NOTYELLOW, 0 STYLES, 0 yellow_styles, 0 not_yellow_s FROM " & ICTPHYJ1 & " Where YELLOW = '1' AND LOCATION_CODE" & vbCrLf _
             & " In (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW <> 0 And YELLOW <> 0)" & vbCrLf _
             & " GROUP by LOCATION_CODE" & vbCrLf _
             & " UNION" & vbCrLf _
             & " Select LOCATION_CODE,0 YELLOW,SUM(LOCATION_QTY) NOTYELLOW, 0 STYLES, 0 yellow_styles, 0 not_yellow_s FROM " & ICTPHYJ1 & " Where YELLOW Is NULL And LOCATION_CODE" & vbCrLf _
             & " In (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW <> 0 And YELLOW <> 0)" & vbCrLf _
             & " GROUP by LOCATION_CODE" & vbCrLf _
             & " union" & vbCrLf _
             & " Select LOCATION_CODE,0 YELLOW,0 NOTYELLOW, count(distinct style_code || color_code) STYLES, 0 yellow_styles, 0 not_yellow_s FROM " & ICTPHYJ1 & " Where  LOCATION_CODE" & vbCrLf _
             & " In (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW <> 0 And YELLOW <> 0)" & vbCrLf _
             & " GROUP by LOCATION_CODE" & vbCrLf _
             & " union" & vbCrLf _
             & " Select LOCATION_CODE,0 YELLOW,0 NOTYELLOW, 0 STYLES, 0 yellow_styles, count(distinct style_code || color_code) not_yellow_s FROM " & ICTPHYJ1 & " Where YELLOW Is NULL And LOCATION_CODE" & vbCrLf _
             & " In (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW <> 0 And YELLOW <> 0)" & vbCrLf _
             & " GROUP by LOCATION_CODE" & vbCrLf _
             & " union" & vbCrLf _
             & " Select LOCATION_CODE,0 YELLOW,0 NOTYELLOW, 0 STYLES, count(distinct style_code || color_code)  yellow_styles, 0 not_yellow_s FROM " & ICTPHYJ1 & " Where YELLOW ='1' And LOCATION_CODE" & vbCrLf _
             & " In (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW <> 0 And YELLOW <> 0)" & vbCrLf _
             & " GROUP by LOCATION_CODE)" & vbCrLf _
             & " GROUP by LOCATION_CODE"
            Create_TDA(.Tables.Add, "ICTPHYJMS", "**", 0, False, "")


            'ASCMAIN1.sql = "Select * From " & ICTPHYJ1 & " Where LOCATION_CODE = :PARM1" & vbCrLf _
            ' & " And LOCATION_CODE In (Select LOCATION_CODE FROM " & ICTPHYJ2 & "  WHERE NOTYELLOW <> 0 And YELLOW <> 0)"
            'Create_TDA(.Tables.Add, "ICTPHYJM", "**", 0, False, "V", 0)


            ASCMAIN1.sql = "Select A.*,ICTSTAT2.WHSE_QTY_OPEN,ICTSTAT2.WHSE_QTY_PICK FROM (" & vbCrLf _
                & " Select * From  " & ICTPHYJ1 & " Where LOCATION_CODE = :PARM1" & vbCrLf _
                & " And LOCATION_CODE In (Select LOCATION_CODE From  " & ICTPHYJ2 & " Where NOTYELLOW <> 0 And YELLOW <> 0)) A,ICTSTAT2" & vbCrLf _
                & " WHERE ICTSTAT2.STYLE_CODE = a.STYLE_CODE" & vbCrLf _
                & " And ICTSTAT2.COLOR_CODE = A.COLOR_CODE" & vbCrLf _
                & " And ICTSTAT2.WHSE_CODE = 'NJC'"
            Create_TDA(.Tables.Add, "ICTPHYJM", "**", 0, False, "V", 0)




            ASCMAIN1.sql = "Select STYLE_CODE,COLOR_CODE,LOCATION_CODE,SUM(LOCATION_QTY) LOCATION_QTY FROM " & ICTPHYJ1 & " Where YELLOW = '1' AND LOCATION_CODe In " & vbCrLf _
            & " (Select LOCATION_CODE FROM " & ICTPHYJ2 & " WHERE NOTYELLOW = 0 And YELLOW <> 0) group by STYLE_CODE, COLOR_CODE, LOCATION_CODE"
            Create_TDA(.Tables.Add, "ICTPHYJP", "**", 0, False, "")

            ASCMAIN1.sql = "Select A.STYLE_CODE ,A.COLOR_CODE,A.LOCATION_QTY,ICTSTAT2.WHSE_QTY_ON_HAND,ICTSTAT2.WHSE_QTY_OPEN,ICTSTAT2.WHSE_QTY_PICK FROM (" & vbCrLf _
                & "Select STYLE_CODE,COLOR_CODE,COUNT(*),SUM(LOCATION_QTY) LOCATION_QTY From " & ICTPHYJ1 & " Where LOCATION_CODE In (" & vbCrLf _
                & " Select LOCATION_CODE From " & ICTPHYJ2 & " Where NOTYELLOW = 0 And YELLOW <> 0)" & vbCrLf _
                & " GROUP by STYLE_CODE, COLOR_CODE) A,ICTSTAT2" & vbCrLf _
                & " WHERE ICTSTAT2.STYLE_CODE = A.STYLE_CODE" & vbCrLf _
                & " And ICTSTAT2.COLOR_CODE = A.COLOR_CODE" & vbCrLf _
                & " And ICTSTAT2.WHSE_CODE = 'NJC'"
            Create_TDA(.Tables.Add, "ICTPHYJS", "**", 0, False, "")


            '''.Relations.Add("ICTPHYJS",
            '''.Tables("ICTPHYJMS").Columns("LOCATION_CODE"),
            '''.Tables("ICTPHYJS").Columns("LOCATION_CODE"))

            '''.Tables("ICTPHYJMS").Columns("STYLE_COUNT)").Expression = "COUNT(CHILD(ICTPHYJS).STYLE_CODE)"
            ''''   .Tables("ICTPHYJS").Columns("OPS_YYYYPP").Expression = "PARENT(APTINVR2).OPS_YYYYPP"



            '''   Dim STYLES_COUNT As Decimal = Val(dst.Tables("ICTPHYJM").Compute("COUNT(STYLE_CODE)", "YELLOW = '1'") & "")






            ''With .Tables("POTSHIPH").Columns
            ''    ' .Add("PO_QTY_SHP_EXT", GetType(System.Int32), "PO_QTY_SHP * (PO_COST)")
            ''    .Add("PO_QTY_SHP_EXT", GetType(System.Decimal), "PO_QTY_SHP * (PO_COST_VCOST + PO_COST_MATLS + PO_COST_OTHER)")

            ''End With

        End With

        With grdICTPHYJY.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                .Columns("LOCATION_QTY").CellAppearance.BackColor = Color.LightBlue
            Next
        End With

        With grdICTPHYJM.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                .Columns("LOCATION_QTY").CellAppearance.BackColor = Color.LightBlue
            Next
        End With

        With grdICTPHYJS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                .Columns("LOCATION_QTY").CellAppearance.BackColor = Color.LightBlue
            Next
        End With

        With grdICTPHYJYS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                .Columns("LOCATION_QTY").CellAppearance.BackColor = Color.LightBlue
            Next
        End With

        grdICTPHYJYS.DataSource = dst.Tables("ICTPHYJYS")
        grdICTPHYJY.DataSource = dst.Tables("ICTPHYJY")
        grdICTPHYJMS.DataSource = dst.Tables("ICTPHYJMS")
        grdICTPHYJ2.DataSource = dst.Tables("ICTPHYJ2")
        grdICTPHYJS.DataSource = dst.Tables("ICTPHYJS")
        grdICTPHYJP.DataSource = dst.Tables("ICTPHYJP")
        grdICTPHYJM.DataSource = dst.Tables("ICTPHYJM")

        Create_Summary(grdICTPHYJYS, "LOCATION_CODE", "Count")
        Create_Summary(grdICTPHYJYS, New String() {"YELLOW"})

        Create_Summary(grdICTPHYJY, "LOCATION_CODE", "Count")
        Create_Summary(grdICTPHYJY, New String() {"LOCATION_QTY"})

        Create_Summary(grdICTPHYJMS, "LOCATION_CODE", "Count")
        Create_Summary(grdICTPHYJMS, New String() {"LOCATION_QTY_YELLOW", "LOCATION_QTY_NOTYELLOW"})
        Create_Summary(grdICTPHYJM, "LOCATION_CODE", "Count")
        Create_Summary(grdICTPHYJM, New String() {"LOCATION_QTY"})


        Create_Summary(grdICTPHYJS, New String() {"LOCATION_QTY", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"})

        Create_Summary(grdICTPHYJP, New String() {"LOCATION_QTY"})


        ' ASCMAIN1.Add_Value_List(grdICTPHYJM, "ACCRUAL_STATUS", Nothing, New String() {":", "O:Not Paid", "1:Invoiced(Paid)"})


        'cbeOPS_YYYYPP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' and OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' order by OPS_YYYYPP DESC")
        'cbeOPS_YYYYPP.ValueMember = "OPS_YYYYPP"
        'cbeOPS_YYYYPP.DisplayMember = "LEGEND"
        'cbeOPS_YYYYPP.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Calculate Pallets"
                'If Absx1.cbeFor("OPS_YYYYPP").Value = "" Then
                '    EMsg &= vbCrLf & "You must specify a Period to View"
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

            Case "Calculate Pallets"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Upload XLS"
                Upload_XLS()


            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Calculate Pallets").Settings.Enabled = False

                .Groups("Screen Control").Items("Upload XLS").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Options").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabMain.Visible = ScreenMode

        Setup_tabMain()

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        '  dst.Tables("ICTRECI0").Rows.Clear()
        dst.Tables("ICTPHYJ2").Rows.Clear()
        dst.Tables("ICTPHYJY").Rows.Clear()
        dst.Tables("ICTPHYJM").Rows.Clear()
        dst.Tables("ICTPHYJS").Rows.Clear()
        dst.Tables("ICTPHYJP").Rows.Clear()
        dst.Tables("ICTPHYJMS").Rows.Clear()
        dst.Tables("ICTPHYJYS").Rows.Clear()
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        RYP = Absx1.cbeFor("OPS_YYYYPP").Value


        Load_POTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub



    Sub Update_Record()

        Call BeginTrans()
        Stop
        Call CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        ' Call Load_Popup_Menu(grdICTRECIG, "SSB", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTPHYJM, "SSS", "Show Filter", "Show GroupBox", "Style Status Inquiry")
        Load_Popup_Menu(grdICTPHYJY, "SSS", "Show Filter", "Show GroupBox", "Style Status Inquiry")
        Load_Popup_Menu(grdICTPHYJP, "SSS", "Show Filter", "Show GroupBox", "Style Status Inquiry")
        Load_Popup_Menu(grdICTPHYJS, "S", "Style Status Inquiry")


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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

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

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "OPS_YYYYPP"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Calculate Pallets", e)
            '    End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "OPS_YYYYPP"
            '    Call Click_Command("Calculate Pallets")
        End Select
    End Sub
#End Region


    Sub Load_POTSHIPX()
        'ICTPHYP2
        Fill_Records("ICTPHYJ2")
        Fill_Records("ICTPHYJY")
        Fill_Records("ICTPHYJS")
        Fill_Records("ICTPHYJP")
        Fill_Records("ICTPHYJMS")
        Fill_Records("ICTPHYJYS")

        '  grdICTPHYJM.Text = "In Transit: " & RYP

    End Sub


    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

    Sub Upload_XLS()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            Dim filter As String = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True
            '  Excel_Import = -1

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using


        'Try

        Dim Vs As New Dictionary(Of String, Integer)

        If FILENAME <> "" Then
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Uploading Data ...")



            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
            Dim range As SpreadsheetGear.IRange = Nothing
            Dim r As Integer = 1
            Dim CNT As Integer = 0
            Dim SQL As String = ""

            SQL = "TRUNCATE TABLE " & ICTPHYJW
            ASCDATA1.ExecuteSQL(SQL)

            ' Rip through Excel

            Do While oSheet.Cells(r, 0).Value & "" <> ""

                Dim STYLE_CODE As String = Trim(oSheet.Cells(r, 0).Value & "")
                Dim COLOR_CODE As String = Trim(oSheet.Cells(r, 1).Value & "")
                Dim STYLE_COLOR_INSERT As String = "('" & STYLE_CODE & "','" & COLOR_CODE & "')"
                If STYLE_CODE <> "" And COLOR_CODE <> "" Then
                    On Error Resume Next
                    SQL = "INSERT INTO " & ICTPHYJW & " VALUES" & STYLE_COLOR_INSERT
                    ASCDATA1.ExecuteSQL(SQL)
                    On Error GoTo 0
                End If

                r = r + 1

                'Dim rowICTPHYJW As DataRow = Nothing
                'rowICTPHYJW = dst.Tables("ICTPHYJW").NewRow
                'With rowICTPHYJW
                '    .Item("STYLE_CODE") = STYLE_CODE
                '    .Item("COLOR_CODE") = COLOR_CODE

                ' End With
                ' dst.Tables("ICTPHYJW").Rows.Add(rowICTPHYJW)
                CNT = CNT + 1
            Loop
        End If

        Dim SQL1 As String = "SELECT WHTLOCB1.*,'' FROM WHTLOCB1,WHTLOCM1 WHERE WHTLOCM1.LOCATION_CODE =  WHTLOCB1.LOCATION_CODE " _
             & " AND WHTLOCB1.LOCATION_QTY <> 0 and WHTLOCB1.WHSE_CODE = 'NJC' AND LOCATION_USE = 'A'"
        ' EXCLUDE WHAT LOCATIONS
        ASCDATA1.ExecuteSQL("Delete from " & ICTPHYJ1)
        ASCDATA1.ExecuteSQL("Insert into " & ICTPHYJ1 & " " & SQL1)

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 Is SELECT * FROM " & ICTPHYJW & ";" _
                    & " BEGIN FOR R1 IN C1 LOOP" _
                    & " Update " & ICTPHYJ1 & " SET YELLOW = '1' WHERE STYLE_CODE= R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE;" _
                    & "     END LOOP;" _
                    & "   END;" _
                    & " END;"
        ASCDATA1.ExecuteSQL()


        SQL1 = "Select LOCATION_CODE, SUM (LOCATION_QTY) QTY, SUM (Case When YELLOW = '1' THEN LOCATION_QTY ELSE 0 END) YELLOW" & vbCrLf _
            & ", SUM(CASE WHEN YELLOW = '1' THEN 0 ELSE LOCATION_QTY END) NOTYELLOW  From  " & ICTPHYJ1 & " Group By LOCATION_CODE"
        ASCDATA1.ExecuteSQL("Delete from " & ICTPHYJ2)
        ASCDATA1.ExecuteSQL("Insert into " & ICTPHYJ2 & " " & SQL1)

        Me.Cursor = Cursors.Default

        Call Click_Command("Calculate Pallets")

    End Sub

    Private Sub grdICTPHYJS_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdICTPHYJS.ClickCell
        grdICTPHYJP.Text = "Pallets for Style" & " " & grdICTPHYJS.ActiveRow.Cells("STYLE_CODE").Value & " Color " & grdICTPHYJS.ActiveRow.Cells("COLOR_CODE").Value

        Dim STYLE_CODE As String = grdICTPHYJS.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = grdICTPHYJS.ActiveRow.Cells("COLOR_CODE").Value & ""

        Dim dvw As DataView = DirectCast(grdICTPHYJP.DataSource, DataTable).DefaultView
        dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

    End Sub



    Private Sub grdICTPHYJMS_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdICTPHYJMS.ClickCell
        Dim LOCATION_CODE As String = grdICTPHYJMS.ActiveRow.Cells("LOCATION_CODE").Value & ""

        Fill_Records("ICTPHYJM", LOCATION_CODE)


        grdICTPHYJM.Text = "Cartons for Pallet/Location Code" & " " & grdICTPHYJMS.ActiveRow.Cells("LOCATION_CODE").Value


        Dim dvw As DataView = DirectCast(grdICTPHYJM.DataSource, DataTable).DefaultView
        dvw.RowFilter = "LOCATION_CODE = '" & LOCATION_CODE & "'"

    End Sub

    Private Sub grdICTPHYJYS_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdICTPHYJYS.ClickCell
        Dim LOCATION_CODE As String = grdICTPHYJYS.ActiveRow.Cells("LOCATION_CODE").Value & ""

        '       Fill_Records("ICTPHYJM", LOCATION_CODE)


        grdICTPHYJY.Text = "Cartons for Pallet/Location Code" & " " & grdICTPHYJYS.ActiveRow.Cells("LOCATION_CODE").Value


        Dim dvw As DataView = DirectCast(grdICTPHYJY.DataSource, DataTable).DefaultView
        dvw.RowFilter = "LOCATION_CODE = '" & LOCATION_CODE & "'"

    End Sub


    Private Sub grdICTPHYJM_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTPHYJM.InitializeRow
        Dim YELLOW As String = e.Row.Cells("YELLOW").Value & ""

        If YELLOW = "1" Then
            e.Row.Cells("YELLOW").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("LOCATION_CODE").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("BAR_CODE").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("STYLE_CODE").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("COLOR_CODE").Appearance.BackColor = Drawing.Color.Yellow
        End If

    End Sub
End Class