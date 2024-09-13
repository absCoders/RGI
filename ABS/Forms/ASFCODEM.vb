Imports ABSolution.Exceptions

Public Class ASFCODEM

    Protected Friend tbl As New DataTable
    Private ROW_NO As Long
    Private ROW_NO_last As Long
    Private ROW_ctr As Long
    Private FORM_NAME As String
    Private tblMRU As DataTable
    Private filter As Collection
    Protected Friend htbkey_COLUMN_NAMEs As Hashtable
    Private key_COLUMN_VALUEs() As String
    Private key_COLUMN_VALUEs_tab As String
    Private row_original As DataRow
    Private rowASTTABD1 As DataRow
    Private OK_TO_LOAD_RECORD As Boolean

    Private last_row As DataRow
    Private copy_from_row As DataRow

    Dim tblGridAttributes As New DataTable

    ' DELETES (AND THE AUDIT TRAIL FOR DELETES)
    ' not supportnig non-text ctls to hold keys

    Public CONTACT_ENTITY_KEY As String
    Public CONTACT_ENTITY_NAME As String

    Private _formLoadError As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me._formLoadError = Nothing

        If Me.DesignMode Then
            Return
        End If

        FORM_NAME = Me.Name
        TABLE_NAME = FORM_NAME

        Me.Text = TABLE_NAME

        Try
            Initialize_Form()

        Catch ex As Exception
            Me._formLoadError = ex.Message
            Exit Sub
        End Try

        'UltraExplorerBar1.Groups("Screen Mode").Items("Single Record View").Active = True
        'UltraExplorerBar1.Groups("Screen Mode").Items("Single Record View").Checked = True

        Clear_Record(True)

        If Absx1.dicCOLUMN_NAME.ContainsKey("SEG2_CODE") Then
            Dim ctl As Control = Absx1.CtlFor("SEG2_CODE", True)
            If ctl IsNot Nothing Then
                Dim row As DataRow = ASCDATA1.GetDataRow("Select * from GLTPARM1")
                Me.GL_Segments(Me, row)
            End If
        End If

        Create_Summary(grdASTAUDT1, "KEY_VALUE", "Count")

        If Set_Contact_Info() Then
            UltraExplorerBar1.Groups("Screen Control").Items("Contacts").Visible = True
        End If

        tblMRU = New DataTable("Record")
        tblMRU.Columns.Add("KEY", Type.GetType("System.String"))
    End Sub

    Private Sub Form_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        With UltraExplorerBar1.Groups("Screen Mode")
            If Not .Items("Single Record View").Checked And _
            Not .Items("Multiple Record View").Checked And _
            Not .Items("Audit Trail").Checked Then
                .Items("Single Record View").Checked = True
            End If
        End With


        If Not ScreenMode Then
            Me.ActiveControl = Absx1.CtlFor(tbl.Columns(0).ColumnName)
            Absx1.CtlFor(tbl.Columns(0).ColumnName).Focus()
        End If
    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If Me.DesignMode Then
            Return
        End If

        UltraExplorerBar1.Groups("Screen Mode").Items("Single Record View").Active = True
        UltraExplorerBar1.Groups("Screen Mode").Items("Single Record View").Checked = True


        If Me._formLoadError IsNot Nothing Then
            MessageBox.Show(Me._formLoadError, "Error Initializing Form", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
            Exit Sub
        End If

        lblOperation.Text = ""

        Set_Viable_Navigation_Options()
    End Sub

    Sub Initialize_Form()

        With tblGridAttributes
            .Columns.Add("GRID_NAME")
            .Columns.Add("ALLOW_ADDNEW", GetType(System.Int16))
            .Columns.Add("ALLOW_DELETE", GetType(System.Int16))
            .Columns.Add("ALLOW_UPDATE", GetType(System.Int16))
            .PrimaryKey = New DataColumn() {.Columns("GRID_NAME")}
        End With

        Get_Grids()

        'ASCMAIN1.sql = "Select * from ASTTABD1 where TABLE_NAME = '" & TABLE_NAME & "'"
        ASCMAIN1.sql = String.Format("Select * from ASTTABD1 where TABLE_NAME = '{0}'", TABLE_NAME)
        rowASTTABD1 = ASCDATA1.GetDataRow(, True)

        tblASFBASE1.TableName = "ASTBASE1" ' TABLE_NAME
        ASCMAIN1.sql = "Select * from " & TABLE_NAME
        tbl.TableName = TABLE_NAME ' "ASTBASE1"
        Create_TDA(tbl, TABLE_NAME, "*")
        'tda = ASCDATA1.GetDataAdapter(tbl, TABLE_NAME, ASCMAIN1.sql, True, , False)

        tblASFBASE1 = tbl.Clone
        tblASFBASE1.TableName = TABLE_NAME
        dst.Tables.Add(tblASFBASE1)

        Create_TDA(dst.Tables.Add, "ASTDFLT1", "*", 1)
        Fill_Records("ASTDFLT1", TABLE_NAME)

        Dim key_COLUMN_NAMES As String = ""

        ReDim key_COLUMN_VALUEs(tbl.PrimaryKey.Length)
        htbkey_COLUMN_NAMEs = New Hashtable
        Dim txtctl As UltraWinEditors.UltraTextEditor
        If tbl.PrimaryKey.Length > 0 Then
            For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                COLUMN_NAME = tbl.PrimaryKey(i).ColumnName
                txtctl = Absx1.txtFor(COLUMN_NAME)
                txtctl.Appearance.BackColor = Color.LightBlue '  .Aquamarine
                key_COLUMN_NAMES &= "," & COLUMN_NAME
                htbkey_COLUMN_NAMEs.Add(COLUMN_NAME, txtctl)
            Next
        Else
            Throw New Exceptions.KeyColumnsRequiredException
            'Throw New Exception("Table Requires at least 1 Key Column")
        End If

        If rowASTTABD1.Item("DO_NOT_PREFETCH_ROWS") & "" <> "1" Then
            Dim sqlSelect = TDAs(TABLE_NAME).SelectCommand.CommandText
            TDAs(TABLE_NAME).SelectCommand.CommandText = "Select * from " & TABLE_NAME & " ORDER BY " & Mid(key_COLUMN_NAMES, 2)
            TDAs(TABLE_NAME).Fill(tbl)
            TDAs(TABLE_NAME).SelectCommand.CommandText = sqlSelect
        End If

        Set_Table_Statistics()
        Setup_grd_Defaults()

        InitializeControls(Me)

        filter = New Collection
    End Sub

    Sub InitializeControls(ByVal c As Control)

        For Each ctl As Control In c.Controls
            If TypeOf ctl Is UltraWinGrid.UltraGrid Then
                Dim grdctl As UltraWinGrid.UltraGrid = DirectCast(ctl, UltraWinGrid.UltraGrid)
                AddHandler grdctl.Error, AddressOf grd_Error
                AddHandler grdctl.BeforeRowUpdate, AddressOf grd_BeforeRowUpdate
            End If

            If ctl.Controls.Count > 0 Then
                InitializeControls(ctl)
            End If
        Next
    End Sub

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdASTAUDT1, "SSSB", "Show Filter", "Show GroupBox", "Card View", "Column Chooser")
        Load_Popup_Menu(grdASTBASE1, "SSSBBBBB", "Show Filter", "Show GroupBox", "Card View", "Column Chooser", "Create Template", "Update/Append Records with Template", "Replace All Records with Template") ', "Refresh"
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

        If tlb_pop.Tools.Exists("Card View") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Card View"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = grd.DisplayLayout.Bands(0).CardView
        End If

        If tlb_pop.Tools.Exists("Update/Append Records with Template") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Update/Append Records with Template"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = (EntryMode = "Edit")
        End If
        If tlb_pop.Tools.Exists("Replace All Records with Template") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Replace All Records with Template"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Visible = (EntryMode = "Edit")
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
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
            Case "Card View"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).CardView = tlb_sbt.Checked

            Case "Column Chooser"
                grd.ShowColumnChooser()


            Case "Refresh"
                Fill_Records("ASTBASE1")

            Case "Create Template"
                Dim dvw As DataView = DirectCast(grdASTBASE1.DataSource, DataTable).DefaultView
                dvw.RowFilter = "1<>1"
                Export_to_Excel(grdASTBASE1, True, False, TABLE_NAME, "")
                dvw.RowFilter = ""

            Case "Update/Append Records with Template", "Replace All Records with Template"
                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.InitialDirectory = ASCMAIN1.Folders("Work")
                    openFileDialog1.Title = "Locate the workbook containing the data to Import"
                    openFileDialog1.Filter = "txt files (*.xls)|*.xls|All files (*.*)|*.*"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True
                    If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using

                Dim cancel_MRM As Boolean = False

                If FILENAME <> "" Then
                    Dim xlApp As Object
                    Dim xlBook As Object
                    Dim xlSheet As Object

                    Try
                        ASCMAIN1.Progress("Now Examining XLS Workbook")
                        Me.Cursor = Cursors.WaitCursor

                        Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FILENAME & "; Extended Properties=""Excel 8.0; HDR=Yes; IMEX=1"""
                        Dim objConn As New System.Data.OleDb.OleDbConnection(sConnectionString)
                        objConn.Open()

                        Dim TABLE_NAME_xls As String = "" ' TABLE_NAME & "$"
                        Dim dto As DataTable = objConn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
                        For Each rto As DataRow In dto.Select("")
                            If CStr(rto.Item("TABLE_NAME") & "").StartsWith(TABLE_NAME) Then
                                TABLE_NAME_xls = CStr(rto.Item("TABLE_NAME") & "")
                            End If
                        Next
                        If TABLE_NAME_xls = "" Then
                            If dto.Rows.Count = 1 Then
                                TABLE_NAME_xls = dto.Rows(0).Item("TABLE_NAME")
                            End If
                        End If

                        'MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [sheet1$]", MyConnection)

                        Dim KEYs As New List(Of String())

                        Dim MRM_TABLENAME As String = "ASTMRMX1"
                        Using objAdapter As New System.Data.OleDb.OleDbDataAdapter()
                            Dim objCmdSelect As New System.Data.OleDb.OleDbCommand("SELECT * FROM [" & TABLE_NAME_xls & "]", objConn)
                            objAdapter.SelectCommand = objCmdSelect

                            If dst.Tables.Contains(MRM_TABLENAME) Then
                                dst.Tables.Remove(MRM_TABLENAME)
                            End If
                            dst.Tables.Add(MRM_TABLENAME)

                            objAdapter.Fill(dst.Tables(MRM_TABLENAME))

                            Dim TBL As DataTable = DirectCast(Me.grdASTBASE1.DataSource, DataTable)
                            Dim empty_row As Boolean = False

                            TBL.Rows.Clear()
                            If e.Tool.Key = "Replace All Records with Template" Then
                                TDAs(TABLE_NAME).Fill(TBL)
                            End If

                            For Each rowIMPORT As DataRow In dst.Tables(MRM_TABLENAME).Rows
                                empty_row = False
                                Dim KEY() As String
                                ReDim KEY(htbkey_COLUMN_NAMEs.Count - 1)

                                Dim K As Integer = 0

                                For Each c As DictionaryEntry In htbkey_COLUMN_NAMEs
                                    If rowIMPORT.Item(K) & "" = "" Then
                                        empty_row = True
                                        Exit For
                                    End If
                                    Try
                                        KEY(K) = rowIMPORT.Item(c.Key.ToString)
                                        K += 1
                                        TDAs(TABLE_NAME).SelectCommand.Parameters("S_" & c.Key.ToString).Value = ""
                                        TDAs(TABLE_NAME).SelectCommand.Parameters("S_" & c.Key.ToString).Value = rowIMPORT.Item(c.Key.ToString)
                                    Catch ex As Exception
                                        Stop
                                    End Try

                                Next

                                If empty_row Then
                                    Exit For
                                End If

                                If KEYs.Contains(KEY) Then
                                    Throw New Exception("Duplicate Key: " & Join(KEY, ":"))
                                End If

                                Dim row As DataRow = Nothing
                                If e.Tool.Key = "Replace All Records with Template" Then
                                    row = TBL.Rows.Find(KEY)
                                    If row Is Nothing Then
                                        row = TBL.NewRow
                                    End If
                                Else
                                    If TDAs(TABLE_NAME).Fill(TBL) = 0 Then
                                        row = TBL.NewRow
                                    Else
                                        row = TBL.Rows.Find(KEY)
                                    End If
                                End If

                                For Each c As DataColumn In dst.Tables(MRM_TABLENAME).Columns
                                    Try
                                        row.Item(Replace(c.ColumnName, " ", "_")) = rowIMPORT.Item(c.ColumnName)
                                    Catch ex As Exception

                                    End Try
                                Next
                                If row.RowState = DataRowState.Detached Then
                                    TBL.Rows.Add(row)
                                End If
                            Next
                        End Using

                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Proceed")
                        cancel_MRM = True
                    Finally

                        xlSheet = Nothing
                        xlBook = Nothing
                        xlApp = Nothing

                        ASCMAIN1.Progress("")
                        Me.Cursor = Cursors.Default

                    End Try

                End If

                If cancel_MRM Then
                    Click_Command("Cancel")
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Customer Inquiry"
            '    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
            '    Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

        End Select
    End Sub

#End Region

    Sub Set_Table_Filter()
        If filter.Count = 0 Then
            tbl.Select("")
        Else
            Dim sql As String = ""
            For i As Integer = 0 To filter.Count - 1
                Dim ctl As Control
                ctl = filter.Item(i + 1)
                Dim z As String = ctl.Name

                If ctl.Name Like "txt*" Then
                    Stop
                    Dim COLUMN_NAME As String = Mid$(z, 4)


                    sql = sql & " and " & COLUMN_NAME & " like '" & ctl.Text & "*'"
                    'sql = sql & " and " & COLUMN_NAME & " = " & ctl.Text & "*"

                End If

                Dim dc As New DataColumn
            Next
            tbl.Select(ASCMAIN1.SQL_Add_WHERE(sql))

        End If
    End Sub

    Sub Set_Filter()
        filter.Clear()

        Dim z As String
        Dim COLUMN_NAME As String

        Dim tdr() As DataRow = tbl.Select("CUST_STATE = 'NY'")
        Dim tdv As New DataView(tbl)
        Stop
        tdv.RowFilter = "CUST_STATE = 'NY'"
        Stop

        For Each ctl As Control In Me.Panel1.Controls
            z = ctl.Name
            COLUMN_NAME = Mid$(z, 4)
            Select Case Mid$(z, 1, 3)
                Case "txt"
                    Dim txt As UltraWinEditors.UltraTextEditor = ctl
                    If txt.Text <> "" Then
                        filter.Add(txt, COLUMN_NAME)
                    End If
                Case "chk"
                    Dim chk As Infragistics.Win.UltraWinEditors.UltraCheckEditor = ctl
                    filter.Add(chk, COLUMN_NAME)
            End Select
        Next
    End Sub

    Sub Set_Table_Statistics()
        ROW_ctr = tbl.Rows.Count
        ROW_NO_last = -1
    End Sub

    Sub Setup_grd_Defaults()

        ' Audit Trail

        With grdASTAUDT1.DisplayLayout.Override
            .HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
            .RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
            .FilterUIType = UltraWinGrid.FilterUIType.FilterRow
            .FilterClearButtonLocation = UltraWinGrid.FilterClearButtonLocation.Row
        End With


        ' Misc Settings

        With grdASTBASE1.DisplayLayout.Override
            .HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
            .RowSelectorHeaderStyle = UltraWinGrid.RowSelectorHeaderStyle.SeparateElement
            .TemplateAddRowPrompt = "Click here to add a new record..."
        End With
        grdASTBASE1.DisplayLayout.ColumnChooserEnabled = DefaultableBoolean.True

    End Sub

    Function Show_Record(ByVal Get_Data_from_Database As Boolean) As Boolean
        If ROW_ctr = 0 And Not Get_Data_from_Database Then
            Exit Function
        End If

        If EntryMode = "New" Then
            UltraExplorerBar1.Groups("Record Tracking").Items("Record Tracking").Text = "Record {New}"
        Else
            UltraExplorerBar1.Groups("Record Tracking").Items("Record Tracking").Text = "Record " & Format$(ROW_NO + 1, "###,##0") & " of " & Format$(ROW_ctr, "###,##0")
        End If
        ' WHAT ABOUT DEFAULTS

        If EntryMode = "Defaults" Then
            Fill_Records("ASTDFLT1", TABLE_NAME)
            rowASFBASE1 = dst.Tables(TABLE_NAME).NewRow
            For Each rowASTDFLT1 As DataRow In dst.Tables("ASTDFLT1").Rows
                COLUMN_NAME = rowASTDFLT1.Item("COLUMN_NAME")
                If rowASFBASE1.Table.Columns.Contains(COLUMN_NAME) Then
                    If Absx1.dicCOLUMN_NAME.ContainsKey(COLUMN_NAME) _
                    Or Absx1.dicCOLUMN_NAME.ContainsKey(TABLE_NAME & "." & COLUMN_NAME) Then
                        'Absx1.CtlFor(COLUMN_NAME).Text = rowASTDFLT1.Item("DEFAULT_VALUE")
                        If TypeOf (Absx1.CtlFor(COLUMN_NAME)) Is UltraWinEditors.UltraCheckEditor Then
                            Absx1.chkFor(COLUMN_NAME).Checked = (rowASTDFLT1.Item("DEFAULT_VALUE") & "" = "1")
                        ElseIf TypeOf (Absx1.CtlFor(COLUMN_NAME)) Is UltraWinEditors.UltraOptionSet Then
                            Absx1.optFor(COLUMN_NAME).Value = rowASTDFLT1.Item("DEFAULT_VALUE") & ""
                        ElseIf TypeOf (Absx1.CtlFor(COLUMN_NAME)) Is UltraWinEditors.UltraNumericEditor Then
                            Absx1.numFor(COLUMN_NAME).Value = rowASTDFLT1.Item("DEFAULT_VALUE") & ""
                        Else
                            Absx1.CtlFor(COLUMN_NAME).Text = rowASTDFLT1.Item("DEFAULT_VALUE") & ""
                        End If
                    End If
                End If

                '    rowASFBASE1.Item(COLUMN_NAME) = rowASTDFLT1.Item("DEFAULT_VALUE")
            Next
            Exit Function
        Else

        End If


        If Get_Data_from_Database Then
            For Each dc As DataColumn In tbl.PrimaryKey
                COLUMN_NAME = dc.ColumnName
                TDAs(TABLE_NAME).SelectCommand.Parameters("S_" & dc.ColumnName).Value = Absx1.CtlFor(COLUMN_NAME).Text
                If EntryMode = "New" Then
                    rowASFBASE1 = tblASFBASE1.NewRow
                    rowASFBASE1.Item(COLUMN_NAME) = Absx1.CtlFor(COLUMN_NAME).Text
                End If
            Next
            tblASFBASE1.Rows.Clear()
            TDAs(TABLE_NAME).Fill(tblASFBASE1)
            If EntryMode <> "New" Then
                If tblASFBASE1.Rows.Count = 0 Then
                    ' THERE IS NO ROW TO SHOW (YET)
                    ' BUT WE NEED TO RESTORE THE VALUES IN THE BOUND CONTROLS
                    rowASFBASE1 = tblASFBASE1.NewRow
                    For Each dc As DataColumn In tbl.PrimaryKey
                        COLUMN_NAME = dc.ColumnName
                        Absx1.CtlFor(COLUMN_NAME).Text = TDAs(TABLE_NAME).SelectCommand.Parameters("S_" & COLUMN_NAME).Value & ""
                        rowASFBASE1.Item(COLUMN_NAME) = TDAs(TABLE_NAME).SelectCommand.Parameters("S_" & COLUMN_NAME).Value & ""
                    Next

                    Exit Function
                    'rowASFBASE1 = tblASFBASE1.NewRow
                Else
                    rowASFBASE1 = tblASFBASE1.Rows(0)
                    OK_TO_LOAD_RECORD = True
                End If
                'rowASFBASE1 = tblASFBASE1.Rows(0)
            Else
                ' ESTABLISH THE NEW ROW AND REPACK THE KEY FIELDS
                rowASFBASE1 = tblASFBASE1.NewRow
                For Each dc As DataColumn In tbl.PrimaryKey
                    COLUMN_NAME = dc.ColumnName
                    Absx1.CtlFor(COLUMN_NAME).Text = TDAs(TABLE_NAME).SelectCommand.Parameters("S_" & COLUMN_NAME).Value & ""
                    rowASFBASE1.Item(COLUMN_NAME) = TDAs(TABLE_NAME).SelectCommand.Parameters("S_" & COLUMN_NAME).Value & ""
                Next
                tblASFBASE1.Rows.Add(rowASFBASE1)

                If optBeginNewRecordsWith.Value = "D" Then
                    For Each rowASTDFLT1 As DataRow In dst.Tables("ASTDFLT1").Rows
                        COLUMN_NAME = rowASTDFLT1.Item("COLUMN_NAME")
                        If rowASFBASE1.Table.Columns.Contains(COLUMN_NAME) Then
                            If Absx1.dicCOLUMN_NAME.ContainsKey(TABLE_NAME & "." & COLUMN_NAME) Then
                                rowASFBASE1.Item(COLUMN_NAME) = rowASTDFLT1.Item("DEFAULT_VALUE")
                                'If TypeOf (Absx1.CtlFor(COLUMN_NAME)) Is UltraWinEditors.UltraCheckEditor Then
                                '    Absx1.chkFor(COLUMN_NAME).Checked = (rowASTDFLT1.Item("DEFAULT_VALUE") & "" = "1")
                                'ElseIf TypeOf (Absx1.CtlFor(COLUMN_NAME)) Is UltraWinEditors.UltraOptionSet Then
                                '    Absx1.optFor(COLUMN_NAME).Value = rowASTDFLT1.Item("DEFAULT_VALUE") & ""
                                'ElseIf TypeOf (Absx1.CtlFor(COLUMN_NAME)) Is UltraWinEditors.UltraNumericEditor Then
                                '    Absx1.numFor(COLUMN_NAME).Value = rowASTDFLT1.Item("DEFAULT_VALUE") & ""
                                'Else
                                '    Absx1.CtlFor(COLUMN_NAME).Text = rowASTDFLT1.Item("DEFAULT_VALUE") & ""
                                'End If
                            End If
                        End If
                    Next
                ElseIf optBeginNewRecordsWith.Value = "U" Then
                    Dim pKeyColumns As New List(Of String)
                    For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                        pKeyColumns.Add(tbl.Columns(i).ColumnName)
                    Next
                    For Each ctl As Control In htbCOLUMN_NAME.Values
                        COLUMN_NAME = Absx1.GetABSColumnName(ctl)
                        If tblASFBASE1.Columns.Contains(COLUMN_NAME) Then
                            If Not pKeyColumns.Contains(COLUMN_NAME) Then
                                rowASFBASE1.Item(COLUMN_NAME) = last_row.Item(COLUMN_NAME)
                            End If
                        End If
                    Next
                ElseIf optBeginNewRecordsWith.Value = "P" Then
                    Dim pKeyColumns As New List(Of String)
                    For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                        pKeyColumns.Add(tbl.Columns(i).ColumnName)
                    Next
                    For Each ctl As Control In htbCOLUMN_NAME.Values
                        COLUMN_NAME = Absx1.GetABSColumnName(ctl)
                        If tblASFBASE1.Columns.Contains(COLUMN_NAME) Then
                            If Not pKeyColumns.Contains(COLUMN_NAME) Then
                                rowASFBASE1.Item(COLUMN_NAME) = copy_from_row.Item(COLUMN_NAME)
                            End If
                        End If
                    Next
                End If
            End If
        Else
            Dim row As DataRow = tbl.Rows(ROW_NO)
            Dim rownew As DataRow = tblASFBASE1.NewRow
            For i As Integer = 0 To tbl.Columns.Count - 1
                rownew.Item(i) = row.Item(i)
            Next
            tblASFBASE1.Rows.Add(rownew)
            rowASFBASE1 = tblASFBASE1.Rows(0)

        End If

        row_original = tbl.NewRow
        row_original.ItemArray = rowASFBASE1.ItemArray

        For Each ctl As Control In htbCOLUMN_NAME.Values
            COLUMN_NAME = Absx1.GetABSColumnName(ctl)
            If Get_Data_from_Database And htbkey_COLUMN_NAMEs.Contains(COLUMN_NAME) Then
            Else
                Try
                    If TypeOf ctl Is UltraWinEditors.UltraTextEditor Then
                        ''ctl.Text = rowASFBASE1.Item(COLUMN_NAME).ToString
                        If Absx1.GetABSHasButton(ctl) Then
                            Populate_Controls_with_Parents(COLUMN_NAME, ctl)
                        End If
                    End If

                Catch ex As Exception
                    Stop
                End Try
            End If

        Next

        Show_Record_Special()

        If EntryMode = "New" Then
        Else

            ' NEEDS TO BE CHANGED FOR MULTI-KEYED TABLES
            Dim RECORD_KEY As String = rowASFBASE1.Item(0).ToString
            Dim r As DataRow
            r = tblMRU.NewRow
            r.Item(0) = RECORD_KEY
            tblMRU.Rows.Add(tblMRU.NewRow)
        End If

        Set_Viable_Navigation_Options()
    End Function

    Sub Load_Data_into_Row_from_Controls()
        For Each ctl As Control In htbCOLUMN_NAME.Values
            Try
                COLUMN_NAME = Absx1.GetABSColumnName(ctl)
                'If COLUMN_NAME = "FRT_CONT_DATE_END" Then Stop
                If TypeOf ctl Is UltraWinEditors.UltraTextEditor Then
                    rowASFBASE1.Item(COLUMN_NAME) = ctl.Text
                End If
                If TypeOf ctl Is UltraWinMaskedEdit.UltraMaskedEdit Then
                    Dim msk As UltraWinMaskedEdit.UltraMaskedEdit

                    msk = ctl
                    msk.DataMode = UltraWinMaskedEdit.MaskMode.Raw
                    rowASFBASE1.Item(COLUMN_NAME) = msk.Text
                End If
                If TypeOf ctl Is UltraWinEditors.UltraNumericEditor Then
                    Dim num As UltraWinEditors.UltraNumericEditor
                    num = ctl
                    rowASFBASE1.Item(COLUMN_NAME) = num.Value
                End If
                If TypeOf ctl Is UltraWinEditors.UltraDateTimeEditor Then
                    Dim dte As UltraWinEditors.UltraDateTimeEditor
                    dte = ctl
                    If dte.Value Is Nothing Then
                        rowASFBASE1.Item(COLUMN_NAME) = DBNull.Value
                    Else

                        rowASFBASE1.Item(COLUMN_NAME) = dte.Value
                    End If
                End If
                If TypeOf ctl Is UltraWinEditors.UltraOptionSet Then
                    Dim opt As UltraWinEditors.UltraOptionSet
                    opt = ctl
                    rowASFBASE1.Item(COLUMN_NAME) = opt.Value
                End If
                If TypeOf ctl Is UltraWinEditors.UltraCheckEditor Then
                    Dim chk As Infragistics.Win.UltraWinEditors.UltraCheckEditor
                    chk = ctl
                    If chk.Checked Then
                        rowASFBASE1.Item(COLUMN_NAME) = "1"
                    Else
                        rowASFBASE1.Item(COLUMN_NAME) = "0"
                    End If
                End If
            Catch ex As Exception
                Dim PARENT_COLUMN_NAME As String = Absx1.GetABSParentColumnName(ctl)
                If PARENT_COLUMN_NAME = "" Then
                    Stop ' YOU PROBABLY HAVE A CONTROL ON THE FORM WITH AN INVALID ABSColumnName - see ctl.Name
                End If
            End Try
        Next
    End Sub

    Sub Set_Viable_Navigation_Options()
        With UltraExplorerBar1.Groups("Navigation")
            .Items("First Record").Settings.Enabled = IIf(ROW_NO = 0 Or ROW_ctr = 0, DefaultableBoolean.False, DefaultableBoolean.True)
            .Items("Previous Record").Settings.Enabled = IIf(ROW_NO = 0 Or ROW_ctr = 0, DefaultableBoolean.False, DefaultableBoolean.True)
            .Items("Next Record").Settings.Enabled = IIf(ROW_NO = ROW_ctr - 1 Or ROW_ctr = 0, DefaultableBoolean.False, DefaultableBoolean.True)
            .Items("Last Record").Settings.Enabled = IIf(ROW_NO = ROW_ctr - 1 Or ROW_ctr = 0, DefaultableBoolean.False, DefaultableBoolean.True)
        End With
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            Select Case .Groups("Screen Mode").CheckedItem.Key
                Case "Single Record View"

                    If Not ScreenMode Then
                        lblOperation.Text = ""
                    Else
                        lblOperation.Text = EntryMode
                    End If
                    spl.Panel1Collapsed = False

                    ' this is essential done within SetReadOnly below
                    'Dim txtctl As UltraWinEditors.UltraTextEditor
                    'For Each txtctl In htbkey_COLUMN_NAMEs.Values
                    '    txtctl.ReadOnly = ScreenMode
                    'Next

                    With .Groups("Screen Control")
                        .Items("New").Settings.Enabled = not_iScreenMode
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                        .Items("View").Settings.Enabled = not_iScreenMode
                        .Items("Contacts").Settings.Enabled = iScreenMode

                        .Items("New").Visible = True
                        .Items("View").Visible = True
                        .Items("Save").Visible = True
                        .Items("Delete").Visible = True
                        .Items("Done").Visible = True
                        .Items("Set Copy-From").Visible = True
                        .Items("Defaults").Visible = True

                        If EntryMode = "View" Then
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                            .Items("Save").Settings.Enabled = DefaultableBoolean.False
                            .Items("Update").Settings.Enabled = DefaultableBoolean.False
                            .Items("Cancel").Settings.Enabled = DefaultableBoolean.False
                            .Items("Delete").Settings.Enabled = DefaultableBoolean.False
                            .Items("Done").Settings.Enabled = DefaultableBoolean.True
                        Else
                            .Items("Save").Settings.Enabled = iScreenMode
                            .Items("Update").Settings.Enabled = iScreenMode
                            .Items("Cancel").Settings.Enabled = iScreenMode
                            .Items("Delete").Settings.Enabled = iScreenMode
                            .Items("Done").Settings.Enabled = DefaultableBoolean.False
                        End If

                        If ScreenMode And EntryMode <> "Defaults" Then
                            .Items("Set Copy-From").Settings.Enabled = DefaultableBoolean.True
                        Else
                            .Items("Set Copy-From").Settings.Enabled = DefaultableBoolean.False
                        End If
                        .Items("Defaults").Settings.Enabled = not_iScreenMode

                    End With

                    With .Groups("Navigation")
                        .Items("First Record").Settings.Enabled = not_iScreenMode
                        .Items("Previous Record").Settings.Enabled = not_iScreenMode
                        .Items("Next Record").Settings.Enabled = not_iScreenMode
                        .Items("Last Record").Settings.Enabled = not_iScreenMode
                        '  If filter Is Nothing Then filter = New Collection 
                        ' this error happenened with pam accidentally added a style to the style master with a lowercase character
                        ' i would prefer that things blow up than to let the error go unnoticed, so i am remming out the if filter isnothing bandaid
                        .Items("Clear Filter").Settings.Enabled = IIf(filter.Count = 0, Infragistics.Win.DefaultableBoolean.False, Infragistics.Win.DefaultableBoolean.True)
                        .Items("Show Filter").Settings.Enabled = IIf(filter.Count = 0, Infragistics.Win.DefaultableBoolean.False, Infragistics.Win.DefaultableBoolean.True)

                        .Visible = True
                        .Visible = False ' UNTIL WE GET THIS WORKING
                    End With


                    .Groups("Record Tracking").Visible = False ' True
                    .Groups("Screen Control").Visible = True
                    .Groups("Begin New Records With").Visible = Not ScreenMode

                    With .Groups("Screen Mode")
                        .Items("Single Record View").Settings.Enabled = not_iScreenMode
                        .Items("Multiple Record View").Settings.Enabled = not_iScreenMode
                        .Items("Audit Trail").Settings.Enabled = not_iScreenMode
                    End With

                    'If Not ScreenMode Then
                    '    Me.ActiveControl = Absx1.CtlFor(tbl.Columns(0).ColumnName)
                    '    Absx1.CtlFor(tbl.Columns(0).ColumnName).Focus()
                    'End If
                    .Groups("Default Mode").Visible = Not ScreenMode

                    If Not ScreenMode Or EntryMode = "View" Then
                        Set_Readonly(True)
                    ElseIf ScreenMode Then
                        Set_Readonly(False)
                    End If

                    If grdASTBASE1.Tag & "" = "2" Then
                        UltraExplorerBar1.Groups("Screen Mode").Items("Multiple Record View").Checked = True
                    End If

                Case "Multiple Record View"
                    spl.Panel1Collapsed = True
                    .Groups("Navigation").Visible = False
                    .Groups("Record Tracking").Visible = False

                    With .Groups("Screen Control")
                        .Visible = True ' False
                        .Items("New").Visible = False
                        .Items("View").Visible = False
                        .Items("Save").Visible = False
                        .Items("Delete").Visible = False
                        .Items("Done").Visible = False
                        .Items("Set Copy-From").Visible = False
                        .Items("Defaults").Visible = False

                        If EntryMode = "Edit" Then
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                            .Items("Update").Settings.Enabled = DefaultableBoolean.True
                            .Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                            grdASTBASE1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                        Else
                            .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                            .Items("Update").Settings.Enabled = DefaultableBoolean.False
                            .Items("Cancel").Settings.Enabled = DefaultableBoolean.False
                            grdASTBASE1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                        End If
                    End With


                    .Groups("Begin New Records With").Visible = False

                    .Groups("Screen Control").Items("Defaults").Settings.Enabled = Infragistics.Win.DefaultableBoolean.False

                    .Groups("Default Mode").Visible = False

                Case "Audit Trail"
                    spl.Panel1Collapsed = True
                    .Groups("Navigation").Visible = False
                    .Groups("Record Tracking").Visible = False
                    .Groups("Screen Control").Visible = False
                    .Groups("Begin New Records With").Visible = False

                    .Groups("Screen Control").Items("Defaults").Settings.Enabled = Infragistics.Win.DefaultableBoolean.False

                    .Groups("Default Mode").Visible = False

            End Select
        End With

        Set_ScreenMode_Special(tf)

    End Sub

    Sub Set_Readonly(ByVal tf As Boolean)
        Set_Read_Only(Panel1, tf)
        If Not ScreenMode Or Not tf Then
            ' FOR TABLES WITH 2 PART KEYS, THIS I-LOOP WAS TRYING TO ITERATE THRU THE 2ND KEY AND THEN THE 1ST KEY, BUT SOMEWHERE IN THIS LOOP, THE ORDER OF THE KEYS IS FLIPPED, AND IT WIND UP DOING THE 1ST KEY TWICW
            'For I As Integer = htbkey_COLUMN_NAMEs.Count To 1 Step -1
            '    COLUMN_NAME = htbkey_COLUMN_NAMEs.Keys(I - 1)
            '    Set_Read_Only_for_ctl(htbkey_COLUMN_NAMEs(COLUMN_NAME), Not tf)
            'Next
            ' and this next attempt fails because it says that the collection was modified, although the collection really wasn't modified
            'For Each COLUMN_NAME As String In htbkey_COLUMN_NAMEs.Keys
            '    Set_Read_Only_for_ctl(htbkey_COLUMN_NAMEs(COLUMN_NAME), Not tf)
            'Next

            Dim COLUMN_NAMEs As New List(Of String)
            For Each COLUMN_NAME As String In htbkey_COLUMN_NAMEs.Keys
                COLUMN_NAMEs.Add(COLUMN_NAME)
            Next
            For Each COLUMN_NAME As String In COLUMN_NAMEs
                Set_Read_Only_for_ctl(htbkey_COLUMN_NAMEs(COLUMN_NAME), Not tf)
            Next
        End If

        If ScreenMode And EntryMode <> "View" Then
            Set_Grids(True)
        Else
            Set_Grids(False)
        End If
    End Sub

    Overridable Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Private Sub UltraExplorerBar1_ItemCheckStateChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemCheckStateChanged

        If UltraExplorerBar1.CheckedItem Is Nothing Then Exit Sub

        Select Case UltraExplorerBar1.CheckedItem.Key
            Case "Single Record View"
                Panel1.Visible = True
                grdASTBASE1.Visible = False
                grdASTAUDT1.Visible = False
                Mode_Settings(False)

            Case "Multiple Record View"
                'grdASTBASE1.Dock = DockStyle.Fill

                Dim keep_position As Boolean = False
                If grdASTBASE1.Tag & "" <> "" Then
                    keep_position = True
                    Dim t As Integer = Val(grdASTBASE1.Tag & "")
                    t = t - 1
                    If t = 0 Then
                        grdASTBASE1.Tag = ""
                    Else
                        grdASTBASE1.Tag = CStr(t)
                    End If
                End If

                If keep_position Then
                    grdASTBASE1.ActiveRow.Selected = True
                Else
                    grdASTBASE1.DataSource = tbl
                    Sort_grdColumns(grdASTBASE1, tblASFBASE1.Columns(0).ColumnName)

                    If grdASTBASE1.DisplayLayout.Bands(0).Summaries.Count = 0 Then
                        Create_Summary(grdASTBASE1, tblASFBASE1.Columns(0).ColumnName, "Count")
                    End If

                    For Each grcol As UltraWinGrid.UltraGridColumn In _
                      grdASTBASE1.DisplayLayout.Bands(0).Columns
                        Dim COLUMN_NAME As String = grcol.Key
                        If Absx1.dicCOLUMN_NAME.ContainsKey(TABLE_NAME & "." & COLUMN_NAME) Then
                        Else
                            grcol.Hidden = True
                        End If
                    Next
                End If

                grdASTBASE1.Visible = True
                grdASTAUDT1.Visible = False
                Panel1.Visible = False
                Mode_Settings(False)

            Case "Audit Trail"
                grdASTBASE1.Visible = False
                Dim sql As String = "Select * from ASTAUDT1 where TABLE_NAME = '" & TABLE_NAME & "'" '  order by COLUMN_NAME, INIT_DATE"
                grdASTAUDT1.DataSource = ASCDATA1.GetDataTable(sql)
                Sort_grdColumns(grdASTAUDT1, "INIT_DATE".ToLower)
                grdASTAUDT1.Visible = True
                Panel1.Visible = False
                Mode_Settings(False)
        End Select
    End Sub

    Sub Check_Key_Fields()

        ReDim key_COLUMN_VALUEs(tbl.PrimaryKey.Length)
        key_COLUMN_VALUEs_tab = ""
        For i As Integer = 0 To tbl.PrimaryKey.Length - 1
            COLUMN_NAME = tbl.Columns(i).ColumnName
            CODE_VALUE = Absx1.txtFor(COLUMN_NAME).Text
            If CODE_VALUE = "" Then
                EMsg &= vbCr & "No Value Specified for " & ASCMAIN1.Make_Caption(COLUMN_NAME)
            ElseIf CODE_VALUE.Contains("'") Then
                EMsg &= vbCr & "Invalid Characters in Value Specified for " & ASCMAIN1.Make_Caption(COLUMN_NAME)
            ElseIf CODE_VALUE.Contains(" ") Then
                EMsg &= vbCr & "Cannot embed spaces in Value Specified for " & ASCMAIN1.Make_Caption(COLUMN_NAME)
            Else
                If i <> tbl.PrimaryKey.Length - 1 Then
                    'Stop
                    '   2 verify it's value if it is not the last keyed field
                End If
                'Select * from ASTVIEW1 where (VIEW_NAME=:S_VIEW_NAME) and (TABLE_NAME=:S_TABLE_NAME)
                TDAs(TABLE_NAME).SelectCommand.Parameters(i).Value = CODE_VALUE
                key_COLUMN_VALUEs(i) = CODE_VALUE
                key_COLUMN_VALUEs_tab &= vbTab & CODE_VALUE
            End If
        Next
        key_COLUMN_VALUEs_tab = Mid$(key_COLUMN_VALUEs_tab, 2)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Proceed_PreReq_Special(eItemKey)
                Check_Key_Fields()

                If EMsg = "" Then
                    If optBeginNewRecordsWith.Value = "U" And last_row Is Nothing Then
                        EMsg &= "There is no Last Record Updated (yet)"
                    ElseIf optBeginNewRecordsWith.Value = "P" And copy_from_row Is Nothing Then
                        EMsg &= "No Record has been Selected (yet) as the Copy-From Record"
                    End If
                End If

                If EMsg = "" Then
                    Dim tblr As New DataTable
                    TDAs(TABLE_NAME).Fill(tblr)
                    If tblr.Rows.Count <> 0 Then
                        EMsg &= vbCr & "Record already exists for {" & Replace$(key_COLUMN_VALUEs_tab, vbTab, ":") & "}"
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, key_COLUMN_VALUEs_tab) Then
                        Exit Sub
                    End If
                End If

            Case "Edit", "View"
                Proceed_PreReq_Special(eItemKey)

                If UltraExplorerBar1.Groups("Screen Mode").Items("Multiple Record View").Checked Then
                Else
                    Check_Key_Fields()

                    If EMsg = "" Then
                        Dim tblr As New DataTable
                        TDAs(TABLE_NAME).Fill(tblr)
                        If tblr.Rows.Count = 0 Then
                            EMsg &= vbCr & "No Record of {" & Replace$(key_COLUMN_VALUEs_tab, vbTab, ":") & "}"
                        End If
                    End If

                    If EMsg = "" Then
                        If eItemKey = "Edit" Then
                            If Not ASCMAIN1.Logical_Lock(TABLE_NAME, key_COLUMN_VALUEs_tab) Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If


            Case "Save"

            Case "Update"
                If EntryMode = "Defaults" Then
                Else
                    Proceed_PreReq_Special(eItemKey)
                    If ASCMAIN1.CLIENT = "RGI" And Me.Text = "EDTXREF4" Then
                        'this screen uses multi key validation which code below doesn't support.
                    Else
                        Validate_Lookups(Me)
                    End If
                End If


            Case "Delete"
                isDeleteAllowed()
                If EMsg = "" Then
                    If MsgBox("Are you sure that you want to Delete this Record?", _
                              MsgBoxStyle.YesNo, _
                              "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Contacts"
                Set_Contact_Info()
                If CONTACT_ENTITY_KEY = "" Then
                    EMsg &= vbCr & "No Contact Key established"
                End If

            Case Else
                If UltraExplorerBar1.Groups("Special Functions").Items.Exists(eItemKey) Then
                    Proceed_PreReq_Special(eItemKey)
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

            Case "Defaults"
                EntryMode = "Defaults"
                Show_Record(True)
                Mode_Settings(True)

            Case "First Record"
                ROW_NO = 0
                Show_Record(False)

            Case "Previous Record"
                If ROW_NO = -1 Then
                    ROW_NO = -1
                Else
                    ROW_NO = ROW_NO - 1
                End If
                If ROW_NO < 0 Then
                    ROW_NO = 0
                End If
                Show_Record(False)

            Case "Next Record"
                If ROW_NO = -1 Then
                    ROW_NO = 0
                Else
                    ROW_NO = ROW_NO + 1
                End If
                If ROW_NO > ROW_ctr - 1 Then
                    ROW_NO = ROW_ctr - 1
                End If
                Show_Record(False)

            Case "Last Record"
                ROW_NO = ROW_ctr - 1
                Show_Record(False)

            Case "New"
                EntryMode = "New"
                Show_Record(True)
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "Edit"
                If UltraExplorerBar1.Groups("Screen Mode").Items("Multiple Record View").Checked Then
                Else
                    Show_Record(True)
                End If
                Mode_Settings(True)

            Case "View"
                EntryMode = "View"
                Show_Record(True)
                Mode_Settings(True)

            Case "Save"
                Dim save_key_COLUMN_VALUEs() As String = key_COLUMN_VALUEs

                Click_Command("Update")

                If EntryMode = "" Then
                    For Each COLUMN_NAME As String In htbkey_COLUMN_NAMEs.Keys
                        Absx1.CtlFor(COLUMN_NAME).Text = last_row.Item(COLUMN_NAME)
                    Next
                    Click_Command("View")
                End If

            Case "Update"

                If EntryMode = "Defaults" Then

                    For Each ctl As Control In htbCOLUMN_NAME.Values
                        COLUMN_NAME = Absx1.GetABSColumnName(ctl)
                        Dim rowASTDFLT1 As DataRow = dst.Tables("ASTDFLT1").Rows.Find(New String() {TABLE_NAME, COLUMN_NAME})

                        Dim DEFAULT_VALUE As Object = Nothing
                        If TypeOf (ctl) Is UltraWinEditors.UltraCheckEditor Then
                            If Absx1.chkFor(COLUMN_NAME).Checked Then DEFAULT_VALUE = "1"
                        ElseIf TypeOf (ctl) Is UltraWinEditors.UltraOptionSet Then
                            If Absx1.optFor(COLUMN_NAME).Value & "" <> "" Then DEFAULT_VALUE = Absx1.optFor(COLUMN_NAME).Value
                        ElseIf TypeOf (ctl) Is UltraWinEditors.UltraNumericEditor Then
                            If Val(Absx1.numFor(COLUMN_NAME).Value & "") <> 0 Then DEFAULT_VALUE = Val(Absx1.numFor(COLUMN_NAME).Value & "")
                        ElseIf TypeOf (ctl) Is UltraWinEditors.UltraDateTimeEditor Then
                            If Absx1.dteFor(COLUMN_NAME).Value & "" <> "" And Absx1.dteFor(COLUMN_NAME).Value & "" <> "1/1/1753" Then DEFAULT_VALUE = Absx1.dteFor(COLUMN_NAME).Value & ""
                        ElseIf TypeOf (ctl) Is UltraWinMaskedEdit.UltraMaskedEdit Then
                            If Absx1.medFor(COLUMN_NAME).Value & "" <> "" Then DEFAULT_VALUE = Absx1.medFor(COLUMN_NAME).Value & ""
                        Else
                            If Absx1.CtlFor(COLUMN_NAME).Text <> "" Then DEFAULT_VALUE = Absx1.CtlFor(COLUMN_NAME).Text
                        End If

                        If rowASFBASE1 IsNot Nothing AndAlso Not rowASFBASE1.Table.Columns.Contains(COLUMN_NAME) _
                        Or DEFAULT_VALUE Is Nothing OrElse DEFAULT_VALUE & "" = "" Then
                            If rowASTDFLT1 IsNot Nothing Then
                                rowASTDFLT1.Delete()
                            End If
                        Else
                            If rowASTDFLT1 Is Nothing Then
                                rowASTDFLT1 = dst.Tables("ASTDFLT1").NewRow
                                rowASTDFLT1.Item("TABLE_NAME") = TABLE_NAME
                                rowASTDFLT1.Item("COLUMN_NAME") = COLUMN_NAME
                                dst.Tables("ASTDFLT1").Rows.Add(rowASTDFLT1)
                            End If
                            rowASTDFLT1.Item("DEFAULT_VALUE") = DEFAULT_VALUE
                        End If
                    Next
                    Update_Record_TDA("ASTDFLT1")

                Else
                    If UltraExplorerBar1.Groups("Screen Mode").Items("Multiple Record View").Checked Then
                    Else
                        last_row = rowASFBASE1.Table.NewRow
                        last_row.ItemArray = rowASFBASE1.ItemArray
                    End If

                    Try

                        BeginTrans()

                        If UltraExplorerBar1.Groups("Screen Mode").Items("Multiple Record View").Checked Then
                            TDAs(TABLE_NAME).Update(tbl)
                        Else
                            Dim KEY_VALUE As String = ""
                            For i As Integer = 0 To tblASFBASE1.PrimaryKey.Length - 1
                                KEY_VALUE &= ":" & rowASFBASE1.Item(i)
                            Next
                            KEY_VALUE = Mid$(KEY_VALUE, 2)

                            Write_Audit_Trail(rowASFBASE1, row_original)

                            If EntryMode = "New" Then
                                If rowASFBASE1.Table.Columns.Contains("INIT_OPER") Then
                                    Try
                                        rowASFBASE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        rowASFBASE1.Item("INIT_DATE") = DATETIME_STAMP
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If

                            If rowASFBASE1.Table.Columns.Contains("LAST_OPER") Then
                                Try
                                    rowASFBASE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                    rowASFBASE1.Item("LAST_DATE") = DATETIME_STAMP
                                Catch ex As Exception

                                End Try
                            End If

                            Proceed_Update_Special_Pre()
                            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                                Enlist_Transaction()
                            End If
                            TDAs(TABLE_NAME).Update(tblASFBASE1)
                            Proceed_Update_Special_Post()

                            tbl.Merge(tblASFBASE1, False)
                        End If

                        CommitTrans()

                        If UltraExplorerBar1.Groups("Screen Mode").Items("Multiple Record View").Checked Then
                            MsgBox("Update Complete")
                        End If

                    Catch ex As Exception
                        Rollback()
                        MsgBox("Update did not occur" & vbCrLf & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Failure")
                    End Try

                End If

                Clear_Record(False)
                Mode_Settings(False)

            Case "Cancel", "Done"
                Clear_Record(False)
                Mode_Settings(False)

                COLUMN_NAME = tbl.Columns(0).ColumnName
                Absx1.txtFor(COLUMN_NAME).Focus()

            Case "Delete"
                Try
                    BeginTrans()

                    rowASFBASE1.Delete()
                    TDAs(TABLE_NAME).Update(tblASFBASE1)
                    Write_Audit_Trail(rowASFBASE1, row_original, "D")

                    CommitTrans("Record has been Deleted")

                Catch ex As Exception
                    Rollback()
                    MsgBox("Delete did not occur" _
                           & vbCrLf & vbCrLf & ex.Message, _
                           MsgBoxStyle.OkOnly, "Failure")
                End Try

                Mode_Settings(False)

            Case "Set Copy-From"

                copy_from_row = rowASFBASE1.Table.NewRow
                copy_from_row.ItemArray = rowASFBASE1.ItemArray
                MsgBox("Copy-From Record has been Set", MsgBoxStyle.OkOnly, "Verification")

            Case "Set Filter"
                Set_Filter()
                Mode_Settings(False)
                Set_Table_Filter()

            Case "Show Filter"
                Mode_Settings(False)

            Case "Clear Filter"
                filter.Clear()
                Set_Table_Filter()
                Set_Table_Statistics()
                Clear_Record(False)
                'Mode_Settings(False)

            Case "Contacts"
                ASCMAIN1.TACMAIN1.Maintain_Contacts(Me,
                       TABLE_NAME,
                       CONTACT_ENTITY_KEY,
                       CONTACT_ENTITY_NAME)
            Case Else
                If UltraExplorerBar1.Groups("Special Functions").Items.Exists(eItemKey) Then
                    DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                    Proceed_Special(eItemKey)
                End If

        End Select

    End Sub

    Overridable Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        'Stop
    End Sub
    Overridable Sub Proceed_Special(ByVal eItemKey As String)
        'Stop
    End Sub
    Overridable Sub Proceed_Update_Special_Pre()

    End Sub

    Overridable Sub Proceed_Update_Special_Post()

    End Sub

    Overridable Sub Show_Record_Special()

    End Sub

    Overridable Sub Clear_Record_Special()

    End Sub

    Private Sub UltraGridExcelExporter1_EndExport(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.EndExportEventArgs)

    End Sub

    Private Sub UltraGridExcelExporter1_InitializeColumn(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ExcelExport.InitializeColumnEventArgs)
        If e.Column.DataType.Name = "DateTime" Then
            'e.ExcelFormatStr = "[$-409]mmmm d, yyyy;@"
            e.ExcelFormatStr = "MM/DD/YYYY"
        End If
    End Sub

    Sub Clear_Record(ByVal initializing As Boolean)
        IsDone = True ' SETTING THIS FLAG A LITTLE EARLIER SINCE FM CALLS CLEAR RECORD BEFORE CALLING MODES(FALSE)
        UltraExplorerBar1.Groups("Record Tracking").Items("Record Tracking").Text = Format$(ROW_ctr, "###,##0") & " Records"

        If initializing Then
            ROW_NO = -1
        End If

        tblASFBASE1.Rows.Clear()
        Clear_Controls(Me.Panel1, initializing)

        Clear_Record_Special()

    End Sub

    Sub Clear_Controls(ByVal pc As Control, ByVal initializing As Boolean)
        Dim COLUMN_NAME As String
        For Each ctl As Control In pc.Controls
            If ctl.Controls.Count > 1 Then
                Clear_Controls(ctl, initializing)
            End If

            COLUMN_NAME = Absx1.GetABSColumnName(ctl)

            If TypeOf ctl Is UltraWinEditors.UltraTextEditor Then
                ctl.Text = ""
                If initializing Then
                    Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(ctl, UltraWinEditors.UltraTextEditor)
                    'AddHandler txtctl.ValueChanged, AddressOf UltraTextEditor1_ValueChanged
                    'AddHandler txtctl.DoubleClick, AddressOf UltraTextEditor1_DoubleClick
                    'AddHandler txtctl.Enter, AddressOf UltraTextEditor1_Enter
                    'AddHandler txtctl.Leave, AddressOf UltraTextEditor1_Leave
                End If
            End If
            If TypeOf ctl Is UltraWinGrid.UltraCombo Then
                ctl.Text = ""
            End If
            If TypeOf ctl Is UltraWinEditors.UltraComboEditor Then
                ctl.Text = ""
            End If
            If TypeOf ctl Is UltraWinMaskedEdit.UltraMaskedEdit Then
                ctl.Text = ""
                Dim txtctl As UltraWinMaskedEdit.UltraMaskedEdit = DirectCast(ctl, UltraWinMaskedEdit.UltraMaskedEdit)
            End If
            If TypeOf ctl Is UltraWinEditors.UltraCheckEditor Then
                Dim chk As UltraWinEditors.UltraCheckEditor
                chk = ctl
                chk.Checked = False
            End If
            If TypeOf ctl Is UltraWinEditors.UltraNumericEditor Then
                Dim num As UltraWinEditors.UltraNumericEditor
                num = ctl
                If num.MinValue > 0 Then
                    num.Value = num.MinValue
                Else
                num.Value = 0
            End If

            End If
            If TypeOf ctl Is UltraWinEditors.UltraDateTimeEditor Then
                Dim dte As UltraWinEditors.UltraDateTimeEditor
                dte = ctl
                dte.Value = DBNull.Value ' ""
            End If
            If TypeOf ctl Is UltraWinEditors.UltraOptionSet Then
                Dim opt As Infragistics.Win.UltraWinEditors.UltraOptionSet
                opt = ctl
                opt.CheckedIndex = -1
            End If
        Next

    End Sub

    Public Overrides Sub Leaving_txt_Special_Before(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        ' why did this procedure not need to know the COLUMN_NAME
        If Not ScreenMode And htbkey_COLUMN_NAMEs.Contains(COLUMN_NAME) Then
            'Show_Record(True)

        End If
        Leaving_txt_Special(COLUMN_NAME, ctl)
    End Sub

    Overridable Sub Leaving_txt_Special(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)

    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        MyBase.txt_KeyDown(sender, e)

        If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
            COLUMN_NAME = Absx1.GetABSColumnName(sender)

            If htbkey_COLUMN_NAMEs.Contains(COLUMN_NAME) Then
                Click_Command(optDefaultMode.Text)
            End If
        End If
    End Sub

    Public Overrides Sub num_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If EntryMode = "" Then
            Dim numctl As UltraWinEditors.UltraNumericEditor = DirectCast(sender, UltraWinEditors.UltraNumericEditor)
            If htbkey_COLUMN_NAMEs.ContainsKey(Absx1.GetABSColumnName(numctl)) Then
            Else
                e.SuppressKeyPress = True
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub cmb_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If EntryMode = "" Then
            Dim cmbctl As UltraWinGrid.UltraCombo = DirectCast(sender, UltraWinGrid.UltraCombo)
            If htbkey_COLUMN_NAMEs.ContainsKey(Absx1.GetABSColumnName(cmbctl)) Then
            Else
                e.SuppressKeyPress = True
                Exit Sub
            End If
        End If
    End Sub

    Public Overloads Sub chk_BeforeCheckStateChanged(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If EntryMode = "" Then
            Dim chkctl As UltraWinEditors.UltraCheckEditor = DirectCast(sender, UltraWinEditors.UltraCheckEditor)
            If htbkey_COLUMN_NAMEs.ContainsKey(Absx1.GetABSColumnName(chkctl)) Then
            Else
                e.Cancel = True
                Exit Sub
            End If
        End If

    End Sub

    Public Overrides Function OK_to_do_View_Lookup(ByVal txtctl As UltraWinEditors.UltraTextEditor) As Boolean
        If EntryMode = "" Then
            If htbkey_COLUMN_NAMEs.ContainsKey(Absx1.GetABSColumnName(txtctl)) Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
    End Function

    Sub GL_Segments(ByVal parent_ctl As Control, ByVal row As DataRow)
        For Each ctl As Control In parent_ctl.Controls
            If ctl.HasChildren Then
                Me.GL_Segments(ctl, row)
            End If
            If TypeOf ctl Is Misc.UltraLabel Then
                Dim lbl As Misc.UltraLabel = DirectCast(ctl, Misc.UltraLabel)
                If lbl.Text = "Segment 2" _
                Or lbl.Text = "Segment 3" _
                Or lbl.Text = "Segment 4" Then
                    Dim SEG As String = Mid(lbl.Text, Len(lbl.Text), 1)
                    lbl.Text = row.Item("GL_PARM_SEG" & SEG & "_DESC") & ""
                    If lbl.Text = "" Then
                        Absx1.txtFor("SEG" & SEG & "_CODE").Visible = False
                        For Each c2 As Control In Absx1.CtlsFor("SEG" & SEG & "_CODE")
                            c2.Visible = False
                        Next
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub grd_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs)
        If e.Row.IsAddRow Then
            For i As Integer = 0 To tbl.PrimaryKey.Length - 1
                Dim COLUMN_NAME As String = tbl.PrimaryKey(i).ColumnName
                If e.Row.Cells.Count > i AndAlso e.Row.Cells(i).Column.Key = COLUMN_NAME Then
                    e.Row.Cells(COLUMN_NAME).Value = Absx1.txtFor(COLUMN_NAME).Text
                End If
            Next
        End If
    End Sub

    Private Sub grd_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs)
        Dim EMSG As String = e.ErrorText
        Dim i As Integer = InStr(EMSG, ".  Value ")
        If i <> 0 And InStr(EMSG, "constrained to be unique.  Value") <> 0 Then
            EMSG = Mid(EMSG, i + 3)
            e.ErrorText = EMSG
        End If
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        grd.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdASTBASE1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdASTBASE1.DoubleClickRow
        If EntryMode = "Edit" Then
            Exit Sub
        End If

        'Click_Command("Single Record View")
        UltraExplorerBar1.Groups("Screen Mode").Items("Single Record View").Checked = True

        Dim keys As New Dictionary(Of String, Object)
        keys.Add(grdASTBASE1.ActiveRow.Cells(0).Column.Key, grdASTBASE1.ActiveRow.Cells(0).Text)
        Remote_Control("Edit", keys)
        grdASTBASE1.Tag = "2"
        '  Click_Command("Edit")
    End Sub

    Private Sub grdASTBASE1_InitializeRow(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTBASE1.InitializeRow
        Dim R As Integer = e.Row.ListIndex
        Select Case tbl.Rows(R).RowState
            Case DataRowState.Added
                e.Row.Appearance.ForeColor = Color.Blue
                e.Row.Appearance.BackColor = Color.Empty
            Case DataRowState.Modified
                e.Row.Appearance.ForeColor = Color.Empty
                e.Row.Appearance.BackColor = Color.Yellow
            Case Else
                e.Row.Appearance.ForeColor = Color.Empty
                e.Row.Appearance.BackColor = Color.Empty
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)
        If Not ScreenMode And htbkey_COLUMN_NAMEs.Contains(COLUMN_NAME) Then
            If optDefaultMode.Text <> "New" Then
                For Each COLUMN_NAME In htbkey_COLUMN_NAMEs.Keys
                    If Absx1.CtlFor(COLUMN_NAME).Text = "" Then
                        Exit Sub
                    End If
                Next

                Click_Command(optDefaultMode.Text)
            End If
        End If
    End Sub

    Sub Get_Grids()

        For Each grdName As String In GRDs.Keys
            If New String() {"ASFBASEX", "ASFBASE1", "ASTAUDT1"}.Contains(grdName) Then
            Else
                Dim row As DataRow = tblGridAttributes.NewRow
                row.Item("GRID_NAME") = grdName
                row.Item("ALLOW_ADDNEW") = GRDs(grdName).DisplayLayout.Override.AllowAddNew
                row.Item("ALLOW_DELETE") = GRDs(grdName).DisplayLayout.Override.AllowDelete
                row.Item("ALLOW_UPDATE") = GRDs(grdName).DisplayLayout.Override.AllowUpdate
                tblGridAttributes.Rows.Add(row)
            End If
        Next

    End Sub

    Sub Set_Grids(ByVal tf As Boolean)
        For Each row As DataRow In tblGridAttributes.Rows
            Dim GRID_NAME As String = row.Item("GRID_NAME")
            If GRDs.ContainsKey(GRID_NAME) Then
                Dim G As UltraWinGrid.UltraGrid = DirectCast(GRDs(GRID_NAME), UltraWinGrid.UltraGrid)
                If tf Then
                    G.DisplayLayout.Override.AllowAddNew = Val(row.Item("ALLOW_ADDNEW"))
                    G.DisplayLayout.Override.AllowDelete = Val(row.Item("ALLOW_DELETE"))
                    G.DisplayLayout.Override.AllowUpdate = Val(row.Item("ALLOW_UPDATE"))
                Else
                    G.DisplayLayout.Override.AllowAddNew = DefaultableBoolean.False
                    G.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    G.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                End If
            End If
        Next
    End Sub

    Public Overrides Function Audit_Context() As Audit_Entity

        Dim E As New Audit_Entity

        E.TABLE_NAME = TABLE_NAME
        E.TABLE_DESC = MENU_ITEM_DESC

        If ScreenMode Then
            E.KEY_VALUE = ""
            For Each COLUMN_NAME As String In htbkey_COLUMN_NAMEs.Keys
                If E.KEY_VALUE <> "" Then E.KEY_VALUE &= ":"
                E.KEY_VALUE &= Absx1.CtlFor(COLUMN_NAME).Text
            Next
            E.KEY_DESC = ""
        End If

        Return E
    End Function

    ''' <summary>
    ''' Validates the Lookup textboxes if they contain a value
    ''' </summary>
    ''' <param name="c">Parent Control to Search. Child controls with controls are searched as well</param>
    ''' <remarks></remarks>
    Private Sub Validate_Lookups(ByVal c As Control)

        Dim COLUMN_NAME As String = ""
        For Each ctl As Control In c.Controls
            If TypeOf ctl Is Infragistics.Win.UltraWinEditors.UltraTextEditor Then
                If MyBase.Absx1.GetABSHasButton(ctl) = True AndAlso MyBase.Absx1.GetABSColumnName(ctl).ToString.Trim.Length > 0 Then
                    COLUMN_NAME = MyBase.Absx1.GetABSColumnName(ctl).ToString.Trim

                    If MyBase.Absx1.txtFor(COLUMN_NAME).Text.Trim.Length > 0 AndAlso Not htbkey_COLUMN_NAMEs.Contains(COLUMN_NAME) Then
                        'Me.Validate_Code(COLUMN_NAME, Precedent_Key_Values)
                        Dim PRECEDENT_KEYS As String = Absx1.GetABSPrecedentKeys(ctl)
                        If PRECEDENT_KEYS <> "" Then
                            Me.Validate_Code(COLUMN_NAME, Add_PK(PRECEDENT_KEYS))
                        Else
                            Me.Validate_Code(COLUMN_NAME)
                        End If
                    End If
                End If
            End If

            If ctl.Controls.Count > 0 Then
                Validate_Lookups(ctl)
            End If
        Next

    End Sub

    Sub isDeleteAllowed_Check_Aliased_Columns(ByVal TABLE_COLUMN As String())
        Dim KEY_VALUE As String = Absx1.txtFor(htbkey_COLUMN_NAMEs.Keys(0)).Text

        For Each TC As String In TABLE_COLUMN
            Dim TABLE_NAME As String = Split(TC, ".")(0)
            Dim COLUMN_NAME As String = Split(TC, ".")(1)
            ASCMAIN1.sql = "Select Count (*) from " & TABLE_NAME _
            & " where " & COLUMN_NAME & " = '" & KEY_VALUE & "'"
            Dim C As Int64 = ASCDATA1.GetDataValue
            If C <> 0 Then
                EMsg &= vbCr & "Delete not allowed - Code (" & KEY_VALUE _
                & ") already in use (" & TABLE_NAME & ") - Column " _
                & COLUMN_NAME
                Exit For
            End If
        Next
    End Sub


    ''' <summary>
    ''' Checks to see if it is OK to delete the record.
    ''' 
    ''' Sets a value in EMsg
    '''  (ie vbCR + {the reason why deletion is not permitted})
    '''  to prevent the Delete from occurring.
    ''' 
    ''' When Overriding this function,
    '''  be sure to include a call to MyBase.isDeleteAllowed
    '''  if you want to extend the functionality of this routine,
    '''  and omit the call to MyBase
    '''  if you want to replace the functionality.
    ''' </summary>
    ''' <remarks></remarks>
    Overridable Sub isDeleteAllowed()

        ASCMAIN1.Progress("Checking if delete is allowed...")
        Dim sqlFindTables As String = ""
        For Each keyCol As String In htbkey_COLUMN_NAMEs.Keys
            sqlFindTables &= String.Format(" INTERSECT" _
                        & " SELECT DISTINCT TABLE_NAME FROM USER_TAB_COLS WHERE" _
                        & " COLUMN_NAME='{0}'", keyCol)
        Next
        sqlFindTables = sqlFindTables.Substring(11)
        sqlFindTables = String.Format("SELECT TABLE_NAME FROM" _
                        & "({0})" _
                        & " WHERE LENGTH(TABLE_NAME) = 8 AND SUBSTR(TABLE_NAME,3,1) = 'T'" _
                        & " AND TABLE_NAME NOT LIKE 'ASW%' AND" _
                        & " TABLE_NAME <> '{1}'", sqlFindTables, TABLE_NAME)

        Dim dtTables As DataTable = ASCDATA1.GetDataTable(sqlFindTables)

        Dim keyValues As String = ""
        Dim sqlWhereClause As String = ""
        For i As Integer = 0 To tbl.PrimaryKey.Length - 1
            Dim keyColumn As String = tbl.Columns(i).ColumnName
            Dim keyValue As String = Absx1.txtFor(keyColumn).Text
            keyValues &= "," & keyValue
            sqlWhereClause &= String.Format(" AND {0}='{1}'", keyColumn, keyValue)
        Next
        sqlWhereClause = sqlWhereClause.Substring(4)

        For Each tableRow As DataRow In dtTables.Rows
            Dim tableToCheck As String = tableRow.Item("TABLE_NAME")
            Dim sqlCheckUse As String = String.Format("SELECT COUNT(*) FROM {0} WHERE ", tableToCheck)
            sqlCheckUse &= sqlWhereClause
            Dim useCount As Integer = Val(ASCDATA1.GetDataValue(sqlCheckUse))
            If useCount > 0 Then
                EMsg &= vbCr & "Delete not allowed - Code (" & Mid(keyValues, 2) & ") already in use (" & tableToCheck & ")"
                Exit For
            End If
        Next

        ASCMAIN1.Progress("")
    End Sub

    Public Overridable Function Set_Contact_Info() As Boolean
        'CONTACT_ENTITY_KEY = ""
        'CONTACT_ENTITY_NAME = ""
        'Return True
    End Function


    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            If htbkey_COLUMN_NAMEs.Count = 1 AndAlso TypeOf (htbkey_COLUMN_NAMEs(htbkey_COLUMN_NAMEs.Keys(0))) Is Infragistics.Win.UltraWinEditors.UltraTextEditor Then
                E.TABLE_NAME = Me.TABLE_NAME
                E.COLUMN_NAME = htbkey_COLUMN_NAMEs.Keys(0)
                E.CODE_VALUE = Absx1.txtFor(E.COLUMN_NAME).Text
                'MAYBE WE NEED TO CONFIGURE THE DESC FIELD FOR A TABLE
                Dim COLUMN_NAME_DESC As String = Replace(E.COLUMN_NAME, "_CODE", "_DESC")
                If Not tbl.Columns.Contains(COLUMN_NAME_DESC) Then
                    COLUMN_NAME_DESC = Replace(E.COLUMN_NAME, "_CODE", "_NAME")
                    If Not tbl.Columns.Contains(COLUMN_NAME_DESC) Then
                        COLUMN_NAME_DESC = ""
                    End If
                Else

                End If
                If COLUMN_NAME_DESC <> "" Then
                    E.DESC_VALUE = Absx1.txtFor(Me.TABLE_NAME & "." & COLUMN_NAME_DESC).Text
                End If

                E.ATTACHMENT_NOTES = ""
                E.RESTRICTIONS = "D"
                'E.READ_ONLY = True
            End If
        End If

        Return E
    End Function

    Private Sub grdASTBASE1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTBASE1.InitializeLayout

    End Sub
End Class